from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

from openpyxl import Workbook
from datetime import datetime
import os
import json
import time
import sys
import subprocess
import re


LINKS_FILE = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "meet_links.txt"
)

CACHE_FILE = os.path.join(
    os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
    "cache_performances.json"
)

with open(LINKS_FILE, encoding="utf-8") as f:
    URLS = [line.strip() for line in f if line.strip()]

all_results = []
scraped_urls = set()


def perf_float(p):
    if not p:
        return 999.0
    clean = re.sub(r'[^\d.]', '', p.replace("(", "").split()[0])
    # Handle empty or malformed (e.g. "..", ".") values
    try:
        return float(clean) if clean else 999.0
    except ValueError:
        return 999.0


GREEK_TO_LATIN = {
    ord('Α'): 'A', ord('α'): 'a',
    ord('Β'): 'V', ord('β'): 'v',
    ord('Γ'): 'G', ord('γ'): 'g',
    ord('Δ'): 'D', ord('δ'): 'd',
    ord('Ε'): 'E', ord('ε'): 'e',
    ord('Ζ'): 'Z', ord('ζ'): 'z',
    ord('Η'): 'I', ord('η'): 'i',
    ord('Θ'): 'TH', ord('θ'): 'th',
    ord('Ι'): 'I', ord('ι'): 'i',
    ord('Κ'): 'K', ord('κ'): 'k',
    ord('Λ'): 'L', ord('λ'): 'l',
    ord('Μ'): 'M', ord('μ'): 'm',
    ord('Ν'): 'N', ord('ν'): 'n',
    ord('Ξ'): 'X', ord('ξ'): 'x',
    ord('Ο'): 'O', ord('ο'): 'o',
    ord('Π'): 'P', ord('π'): 'p',
    ord('Ρ'): 'R', ord('ρ'): 'r',
    ord('Σ'): 'S', ord('σ'): 's', ord('ς'): 's',
    ord('Τ'): 'T', ord('τ'): 't',
    ord('Υ'): 'Y', ord('υ'): 'y',
    ord('Φ'): 'F', ord('φ'): 'f',
    ord('Χ'): 'CH', ord('χ'): 'ch',
    ord('Ψ'): 'PS', ord('ψ'): 'ps',
    ord('Ω'): 'O', ord('ω'): 'o',
    ord('Ά'): 'A', ord('ά'): 'a',
    ord('Έ'): 'E', ord('έ'): 'e',
    ord('Ή'): 'I', ord('ή'): 'i',
    ord('Ί'): 'I', ord('ί'): 'i', ord('ΐ'): 'i',
    ord('Ό'): 'O', ord('ό'): 'o',
    ord('Ύ'): 'Y', ord('ύ'): 'y', ord('ΰ'): 'y',
    ord('Ώ'): 'O', ord('ώ'): 'o',
}

def normalize_name(name):
    """Transliterate Greek to Latin so ΓΕΩΡΓΙΑ ΣΤΑΘΟΠΟΥΛΟΥ and GEORGIA STATHOPOULOU match."""
    return name.translate(GREEK_TO_LATIN).upper().strip()

def nk_latin(name):
    """Extract surname as Latin string for cross-script matching."""
    n = normalize_name(name)
    n = n.replace("OY", "OU")
    parts = n.split()
    return parts[-1] if parts else ""

def norm_full(name):
    """Full name normalized to Latin, for cross-script identity matching."""
    n = normalize_name(name)
    n = n.replace("OY", "OU").replace("GG", "NG")
    return n


# =========================
# LOAD EXISTING CACHE
# =========================
if os.path.exists(CACHE_FILE):
    print("Loading existing cache...")
    try:
        with open(CACHE_FILE, 'r', encoding='utf-8') as f:
            data = json.load(f)
        if isinstance(data, dict) and "performances" in data:
            all_results = data["performances"]
            scraped_urls = set(data.get("scraped_urls", []))
        elif isinstance(data, list):
            all_results = data
        # Filter out entries with unparseable performances
        before = len(all_results)
        all_results = [r for r in all_results if perf_float(r["performance"]) < 900]
        if len(all_results) < before:
            print(f"  Removed {before - len(all_results)} corrupted entry(ies)")
        print(f"Loaded {len(all_results)} performances from cache\n")
    except (json.JSONDecodeError, KeyError):
        print("[WARN] Cache file is corrupted. Starting fresh...")
        all_results = []

# =========================
# DETERMINE NEW URLS
# =========================
urls_to_scrape = [url for url in URLS if url not in scraped_urls]
print(f"Total links in file: {len(URLS)} | Already cached: {len(URLS) - len(urls_to_scrape)} | New: {len(urls_to_scrape)}")

if urls_to_scrape:
    print(f"Scraping {len(urls_to_scrape)} new link(s)...\n")

    # =========================
    # CHROME OPTIONS
    # =========================
    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    chrome_options.add_argument("--disable-dev-shm-features")
    chrome_options.add_argument("--no-sandbox")

    # =========================
    # SCRAPE LOOP (PARALLEL)
    # =========================
    from concurrent.futures import ThreadPoolExecutor, as_completed

    NUM_WORKERS = 3

    PERFORMANCE_RE = re.compile(r"\d+\.\d+")

    def find_performance_cell(row):
        cells = row.find_elements(By.TAG_NAME, "td")
        for idx, cell in enumerate(cells):
            text = cell.text.split("\n")[0].strip()
            if "DNS" in text or "DNF" in text:
                continue
            if PERFORMANCE_RE.search(text):
                perf = text.replace("SB", "").replace("PB", "").strip()
                return perf, idx
        return None, None

    def start_driver():
        return webdriver.Chrome(
            service=Service(ChromeDriverManager().install()),
            options=chrome_options
        )

    def worker(urls_batch, worker_id):
        driver = start_driver()
        local_results = []
        succeeded = []
        try:
            for url_index, url in enumerate(urls_batch):
                print(f"\n[W{worker_id}] ({url_index + 1}/{len(urls_batch)}) {url}")
                try:
                    driver.get(url)

                    try:
                        WebDriverWait(driver, 15).until(
                            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr"))
                        )
                    except TimeoutException:
                        print(f"[W{worker_id}]   [WARN] Timeout: Could not load schedule from {url}")
                        succeeded.append(url)
                        continue

                    time.sleep(1)

                    # =========================
                    # SCRAPE MEET INFO
                    # =========================
                    body_text = driver.find_element(By.TAG_NAME, "body").text
                    lines = body_text.split("\n")

                    meet_name = ""
                    meet_date = ""
                    meet_location = ""

                    for i, line in enumerate(lines):
                        line = line.strip()
                        if line == "ΟΝΟΜΑ" and i + 1 < len(lines):
                            raw = lines[i + 1].strip()
                            if " · " in raw:
                                meet_name = raw.split(" · ", 1)[1].strip()
                            else:
                                meet_name = raw
                        if line == "ΗΜΕΡΟΜΗΝΙΑ & ΩΡΑ" and i + 1 < len(lines):
                            raw = lines[i + 1].strip()
                            meet_date = raw[:10]
                        if line == "ΠΟΛΗ & ΧΩΡΑ" and i + 1 < len(lines):
                            raw = lines[i + 1].strip()
                            meet_location = re.sub(r'^\d+\s*', '', raw).strip()

                    meet_name = re.sub(r'^Roster Athletics\s*·\s*', '', meet_name).strip()

                    if not meet_name:
                        meet_name = driver.title.strip()

                    print(f"[W{worker_id}]   Meet: {meet_name} | Date: {meet_date} | Location: {meet_location}")

                    # =========================
                    # FIND WOMEN'S 100m FINALS
                    # =========================
                    rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
                    print(f"[W{worker_id}]   Found {len(rows)} rows in schedule")

                    matching_indices = []
                    for i, row in enumerate(rows):
                        text = row.text.strip()
                        if (
                            "100μ" in text
                            and "Τελικός" in text
                            and "Εμπόδια" not in text
                            and "4Χ100" not in text.upper()
                            and "4X100" not in text.upper()
                            and ("Γυναίκες" in text or "Κορίτσια" in text)
                        ):
                            matching_indices.append(i)
                            print(f"[W{worker_id}]   [OK] Matched row {i}: {repr(text)}")

                    print(f"[W{worker_id}]   Found {len(matching_indices)} matching women's 100m finals rows")

                    results_urls = []

                    for i in matching_indices:
                        try:
                            WebDriverWait(driver, 15).until(
                                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr"))
                            )
                            time.sleep(0.5)
                            rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")

                            if i >= len(rows):
                                print(f"[W{worker_id}]   [WARN] Row {i} no longer exists (only {len(rows)} rows)")
                                continue

                            row = rows[i]
                            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row)
                            time.sleep(0.5)

                            anchors = row.find_elements(By.TAG_NAME, "a")
                            if anchors:
                                href = anchors[0].get_attribute("href")
                                if href:
                                    results_urls.append(href)
                                    print(f"[W{worker_id}]   [OK] Found results URL: {href}")
                                    continue

                            td = row.find_elements(By.TAG_NAME, "td")
                            target = td[1] if len(td) > 1 else row

                            actions = ActionChains(driver)
                            actions.move_to_element(target).click().perform()

                            try:
                                WebDriverWait(driver, 10).until(
                                    lambda d: "meId" in d.current_url
                                )
                            except TimeoutException:
                                print(f"[W{worker_id}]   [WARN] URL did not change after clicking row {i}")
                                driver.back()
                                WebDriverWait(driver, 15).until(
                                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr"))
                                )
                                continue

                            current = driver.current_url
                            results_urls.append(current)
                            print(f"[W{worker_id}]   [OK] Found results URL: {current}")
                            driver.back()
                            WebDriverWait(driver, 15).until(
                                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr"))
                            )

                        except Exception as click_error:
                            print(f"[W{worker_id}]   [WARN] Could not click row {i}: {click_error}")
                            continue

                    if not results_urls:
                        print(f"[W{worker_id}]   [WARN] No women's 100m finals found in {url}")
                        succeeded.append(url)
                        continue

                    # =========================
                    # SCRAPE EACH RESULTS PAGE
                    # =========================
                    for results_url in results_urls:
                        print(f"[W{worker_id}]   Scraping results: {results_url}")

                        try:
                            driver.get(results_url)

                            try:
                                WebDriverWait(driver, 15).until(
                                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr"))
                                )
                            except TimeoutException:
                                print(f"[W{worker_id}]   [WARN] Timeout loading results: {results_url}")
                                continue

                            time.sleep(1)

                            body_text = driver.find_element(By.TAG_NAME, "body").text
                            lines = body_text.split("\n")

                            heat_name = ""
                            for line in lines:
                                line = line.strip()
                                if "100μ" in line and "Τελικός" in line:
                                    parts = line.split("Τελικός")
                                    if len(parts) > 1:
                                        after = parts[1].strip()
                                        letter = after.split()[0] if after else ""
                                        heat_name = f"Τελικός {letter}" if letter else "Τελικός"
                                    break

                            wind = ""
                            for line in lines:
                                line = line.strip()
                                if line.startswith("Άνεμος:"):
                                    wind = line.replace("Άνεμος:", "").strip()
                                    break

                            result_rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
                            print(f"[W{worker_id}]   [OK] Βρέθηκαν {len(result_rows)} γραμμές")

                            for row in result_rows:
                                try:
                                    cells = row.find_elements(By.TAG_NAME, "td")

                                    if len(cells) < 5:
                                        continue

                                    lane = cells[1].text.strip()
                                    athlete_info = cells[2].text.split("\n")
                                    name = athlete_info[0].strip()

                                    if len(athlete_info) < 2:
                                        continue

                                    try:
                                        birth_year = int(athlete_info[1].strip())
                                    except ValueError:
                                        continue

                                    performance, perf_idx = find_performance_cell(row)

                                    if not performance:
                                        continue

                                    all_cells = row.find_elements(By.TAG_NAME, "td")
                                    club = all_cells[perf_idx - 1].text.strip() if perf_idx and perf_idx > 0 else ""
                                    if club and name and club == name.split()[-1]:
                                        club = ""

                                    if 2009 <= birth_year <= 2012:
                                        local_results.append({
                                            "name": name,
                                            "birth_year": birth_year,
                                            "club": club,
                                            "performance": performance,
                                            "wind": wind,
                                            "competition": meet_name,
                                            "date": meet_date,
                                            "location": meet_location,
                                            "heat": heat_name,
                                            "lane": lane
                                        })
                                except Exception as row_error:
                                    print(f"[W{worker_id}]     [WARN] Error parsing row: {row_error}")
                                    continue

                        except Exception as results_error:
                            print(f"[W{worker_id}]   [ERR] Error scraping results page: {results_error}")
                            continue

                except Exception as e:
                    print(f"[W{worker_id}]   [ERR] Error scraping {url}: {e}")
                    if "invalid session id" in str(e).lower() or "no such window" in str(e).lower():
                        try:
                            driver.quit()
                        except Exception:
                            pass
                        driver = start_driver()
                    continue

                succeeded.append(url)
        finally:
            try:
                driver.quit()
            except Exception:
                pass
        return local_results, succeeded

    # Distribute URLs evenly across workers
    batches = [[] for _ in range(NUM_WORKERS)]
    for i, url in enumerate(urls_to_scrape):
        batches[i % NUM_WORKERS].append(url)
    batches = [b for b in batches if b]

    all_new_results = []
    all_succeeded = []
    with ThreadPoolExecutor(max_workers=len(batches)) as executor:
        futures = {executor.submit(worker, batch, i): i for i, batch in enumerate(batches)}
        for future in as_completed(futures):
            try:
                results, succeeded = future.result()
                all_new_results.extend(results)
                all_succeeded.extend(succeeded)
            except Exception as e:
                print(f"[ERR] Worker failed: {e}")

    all_results.extend(all_new_results)
    scraped_urls.update(all_succeeded)

    # Remove duplicates
    unique = {}
    for r in all_results:
        key = (
            r["name"],
            r["birth_year"],
            r["competition"],
            r["performance"]
        )
        unique[key] = r

    all_results = list(unique.values())

    # =========================
    # SAVE CACHE
    # =========================
    print(f"\nSaving {len(all_results)} performances to cache...")
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump({
            "performances": all_results,
            "scraped_urls": list(scraped_urls)
        }, f, ensure_ascii=False, indent=2)
    print("[OK] Cache saved successfully")
else:
    print("All links already scraped. Using cached data.")


# =========================
# CLEAN COMPETITION NAMES
# =========================
for r in all_results:
    r["competition"] = re.sub(r'^Roster Athletics\s*·\s*', '', r["competition"]).strip()

# =========================
# NORMALIZE NAMES TO GREEK
# =========================
# Build a mapping of (normalized_name, year) -> preferred Greek name
preferred_names = {}
for r in all_results:
    has_greek = any('\u0370' <= c <= '\u03FF' for c in r['name'])
    key = (norm_full(r["name"]), r["birth_year"])
    if has_greek and key not in preferred_names:
        preferred_names[key] = r['name']

# Replace Latin names with their Greek equivalent where available
for r in all_results:
    has_greek = any('\u0370' <= c <= '\u03FF' for c in r['name'])
    if not has_greek:
        key = (norm_full(r["name"]), r["birth_year"])
        if key in preferred_names:
            r['name'] = preferred_names[key]

# =========================
# SEASON BEST
# =========================
season_best = {}

for r in all_results:
    try:
        perf = perf_float(r["performance"])
    except ValueError:
        continue

    key = (norm_full(r["name"]), r["birth_year"])

    if key not in season_best:
        season_best[key] = r
    else:
        if perf < perf_float(season_best[key]["performance"]):
            season_best[key] = r

ranking = sorted(
    season_best.values(),
    key=lambda x: perf_float(x["performance"])
)

# =========================
# OUTPUT FILE
# =========================
base_dir = os.path.dirname(os.path.abspath(__file__))
output_dir = os.path.join(base_dir, "output")
os.makedirs(output_dir, exist_ok=True)

timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")

filename = os.path.join(
    output_dir,
    f"2026_U18_100m_GR_{timestamp}.xlsx"
)

# =========================
# EXCEL
# =========================
wb = Workbook()

ws1 = wb.active
ws1.title = "All_Performances"

ws1.append([
    "Rank", "Name", "Birth Year", "Club", "Performance",
    "Wind", "Competition", "Date", "Location", "Heat", "Lane", "Notes"
])

sorted_all = sorted(
    all_results,
    key=lambda x: perf_float(x["performance"])
)

for i, r in enumerate(sorted_all, 1):
    ws1.append([
        i,
        r["name"],
        r["birth_year"],
        r["club"].upper(),
        r["performance"],
        r["wind"],
        r["competition"].upper(),
        r["date"],
        r["location"],
        r["heat"],
        r["lane"],
        ""
    ])

ws2 = wb.create_sheet("Season_Best")

ws2.append([
    "Rank", "Name", "Birth Year", "Club", "Best Performance",
    "Wind", "Competition", "Date", "Location", "Heat", "Lane", "Notes"
])

for i, r in enumerate(ranking, 1):
    ws2.append([
        i,
        r["name"],
        r["birth_year"],
        r["club"].upper(),
        r["performance"],
        r["wind"],
        r["competition"].upper(),
        r["date"],
        r["location"],
        r["heat"],
        r["lane"],
        ""
    ])

wb.save(filename)

print("\n" + "="*50)
print("DONE")
print("="*50)
print(f"Excel file: {filename}")
print(f"Total performances: {len(all_results)}")
print(f"Season best athletes: {len(ranking)}")

if sys.platform == "win32":
    os.startfile(filename)
elif sys.platform == "darwin":
    subprocess.run(["open", filename])
else:
    subprocess.run(["xdg-open", filename])