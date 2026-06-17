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

# =========================
# CHECK FOR CACHE
# =========================
USE_CACHE = False
if os.path.exists(CACHE_FILE):
    response = input(f"\nCache file found ({os.path.getsize(CACHE_FILE)} bytes). Use it? (y/n): ").lower()
    USE_CACHE = response == 'y'

if USE_CACHE:
    print("Loading from cache...")
    try:
        with open(CACHE_FILE, 'r', encoding='utf-8') as f:
            all_results = json.load(f)
        print(f"Loaded {len(all_results)} performances from cache\n")
    except json.JSONDecodeError:
        print("⚠️ Cache file is corrupted. Starting fresh...")
        all_results = []
        USE_CACHE = False
else:
    # =========================
    # CHROME OPTIONS
    # =========================
    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    chrome_options.add_argument("--disable-dev-shm-features")
    chrome_options.add_argument("--no-sandbox")

    # =========================
    # SCRAPE LOOP
    # =========================
    for url_index, url in enumerate(URLS):
        print(f"\n[{url_index + 1}/{len(URLS)}] Επεξεργασία: {url}")

        driver = None
        try:
            driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=chrome_options
            )

            driver.get(url)

            try:
                WebDriverWait(driver, 15).until(
                    EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr"))
                )
            except TimeoutException:
                print(f"⚠️ Timeout: Could not load schedule from {url}")
                driver.quit()
                continue

            time.sleep(2)

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

            # Remove "Roster Athletics · " prefix if present
            meet_name = re.sub(r'^Roster Athletics\s*·\s*', '', meet_name).strip()

            if not meet_name:
                meet_name = driver.title.strip()

            print(f"  Meet: {meet_name} | Date: {meet_date} | Location: {meet_location}")

            # =========================
            # FIND WOMEN'S 100m FINALS
            # =========================
            rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
            print(f"  Found {len(rows)} rows in schedule")

            matching_indices = []
            for i, row in enumerate(rows):
                text = row.text.strip()
                if (
                    "100μ" in text
                    and "Τελικός" in text
                    and "Εμπόδια" not in text
                    and ("Γυναίκες" in text or "Κορίτσια" in text)
                ):
                    matching_indices.append(i)
                    print(f"  ✓ Matched row {i}: {repr(text)}")

            print(f"  Found {len(matching_indices)} matching women's 100m finals rows")

            results_urls = []

            for i in matching_indices:
                try:
                    WebDriverWait(driver, 15).until(
                        EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr"))
                    )
                    time.sleep(1)
                    rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")

                    if i >= len(rows):
                        print(f"  ⚠️ Row {i} no longer exists (only {len(rows)} rows)")
                        continue

                    row = rows[i]
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", row)
                    time.sleep(1)

                    td = row.find_elements(By.TAG_NAME, "td")
                    target = td[1] if len(td) > 1 else row

                    actions = ActionChains(driver)
                    actions.move_to_element(target).click().perform()

                    try:
                        WebDriverWait(driver, 10).until(
                            lambda d: "meId" in d.current_url
                        )
                    except TimeoutException:
                        print(f"  ⚠️ URL did not change after clicking row {i}")
                        driver.back()
                        time.sleep(3)
                        continue

                    current = driver.current_url
                    results_urls.append(current)
                    print(f"  ✓ Found results URL: {current}")
                    driver.back()
                    time.sleep(3)

                except Exception as click_error:
                    print(f"  ⚠️ Could not click row {i}: {click_error}")
                    continue

            if not results_urls:
                print(f"  ⚠️ No women's 100m finals found in {url}")
                continue

            # =========================
            # SCRAPE EACH RESULTS PAGE
            # =========================
            for results_url in results_urls:
                print(f"  Scraping results: {results_url}")

                try:
                    driver.get(results_url)

                    try:
                        WebDriverWait(driver, 15).until(
                            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr"))
                        )
                    except TimeoutException:
                        print(f"  ⚠️ Timeout loading results: {results_url}")
                        continue

                    time.sleep(2)

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
                    print(f"  ✓ Βρέθηκαν {len(result_rows)} γραμμές")

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

                            club = cells[3].text.strip()
                            performance_text = cells[4].text.split("\n")[0].strip()

                            if performance_text == "" or "DNS" in performance_text or "DNF" in performance_text:
                                continue

                            performance = performance_text.replace("SB", "").replace("PB", "").strip()

                            if 2009 <= birth_year <= 2012:
                                all_results.append({
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
                            print(f"    ⚠️ Error parsing row: {row_error}")
                            continue

                except Exception as results_error:
                    print(f"  ❌ Error scraping results page: {results_error}")
                    continue

        except Exception as e:
            print(f"❌ Error scraping {url}: {e}")
            continue

        finally:
            if driver:
                try:
                    driver.quit()
                except Exception:
                    pass

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

    print(f"\nSaving {len(all_results)} performances to cache...")
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    print("✓ Cache saved successfully")


# =========================
# CLEAN COMPETITION NAMES
# =========================
for r in all_results:
    r["competition"] = re.sub(r'^Roster Athletics\s*·\s*', '', r["competition"]).strip()

# =========================
# SEASON BEST
# =========================
season_best = {}

for r in all_results:
    try:
        perf = float(r["performance"])
    except ValueError:
        continue

    key = (r["name"], r["birth_year"])

    if key not in season_best:
        season_best[key] = r
    else:
        if perf < float(season_best[key]["performance"]):
            season_best[key] = r

ranking = sorted(
    season_best.values(),
    key=lambda x: float(x["performance"])
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
    "Wind", "Competition", "Date", "Location", "Heat", "Lane"
])

sorted_all = sorted(
    all_results,
    key=lambda x: float(x["performance"].replace("(", "").split()[0]) if x["performance"] else 999
)

for i, r in enumerate(sorted_all, 1):
    ws1.append([
        i,
        r["name"],
        r["birth_year"],
        r["club"],
        r["performance"],
        r["wind"],
        r["competition"],
        r["date"],
        r["location"],
        r["heat"],
        r["lane"]
    ])

ws2 = wb.create_sheet("Season_Best")

ws2.append([
    "Rank", "Name", "Birth Year", "Club", "Best Performance",
    "Wind", "Competition", "Date", "Location", "Heat", "Lane"
])

for i, r in enumerate(ranking, 1):
    ws2.append([
        i,
        r["name"],
        r["birth_year"],
        r["club"],
        r["performance"],
        r["wind"],
        r["competition"],
        r["date"],
        r["location"],
        r["heat"],
        r["lane"]
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