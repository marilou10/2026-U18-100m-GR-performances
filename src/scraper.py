# Copyright (c) 2026 Μαρία Ελένη Αντωνοπούλου
# Licensed under the MIT License. See LICENSE file.
#
# 2026 U18 100m Greek performances — scraper & orchestrator.
# Data sourced from Roster Athletics (https://www.rosterathletics.com).
# I do not own or claim ownership of any athlete data, results, or meet information.

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.action_chains import ActionChains
from webdriver_manager.chrome import ChromeDriverManager

import json
import time
import sys
import re
import unicodedata
import os

from config import (
    LINKS_FILE, CACHE_FILE, NOTES_FILE, NUM_WORKERS,
    OVERRIDE_EXCLUDE, NON_GREEK_CLUBS, MANUAL_CLUB,
)
from normalizer import (
    perf_float, norm_full, nk_latin, has_greek_chars,
    latin_to_greek, clean_competition_location, translate_location,
    clean_performance, is_wind_legal, fmt_wind, fmt_comp, fmt_loc,
    uppercase_all, clear_placeholder_clubs, clear_lane_placeholders,
    _club_key,
)
from exporter import (
    create_excel, export_csv, export_pdf, open_file,
)


# =========================
# LOAD EXISTING CACHE
# =========================
all_results = []
scraped_urls = set()

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
        before = len(all_results)
        all_results = [r for r in all_results if 0 < perf_float(r["performance"]) < 25.0]
        if len(all_results) < before:
            print(f"  Removed {before - len(all_results)} corrupted/non-100m entry(ies)")
        print(f"Loaded {len(all_results)} performances from cache\n")
    except (json.JSONDecodeError, KeyError):
        print("[WARN] Cache file is corrupted. Starting fresh...")
        all_results = []

# =========================
# DETERMINE NEW URLS
# =========================
fixed_links = []
with open(LINKS_FILE, 'r', encoding="utf-8") as f:
    raw_lines = f.readlines()

for line in raw_lines:
    stripped = line.strip()
    if not stripped:
        fixed_links.append(line)
        continue
    original = stripped
    stripped = stripped.replace("about?id=", "schedule?id=")
    fixed_links.append(stripped + "\n")

URLS = [l.strip() for l in fixed_links if l.strip()]

if any("about?id=" in l for l in raw_lines):
    with open(LINKS_FILE, 'w', encoding="utf-8") as f:
        f.writelines(fixed_links)
    print("[OK] Auto-converted about?id= to schedule?id= in meet_links.txt")

urls_to_scrape = [url for url in URLS if url not in scraped_urls]
print(f"Total links in file: {len(URLS)} | Already cached: {len(URLS) - len(urls_to_scrape)} | New: {len(urls_to_scrape)}")

if urls_to_scrape:
    print(f"Scraping {len(urls_to_scrape)} new link(s)...\n")

    chrome_options = Options()
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    chrome_options.add_argument("--disable-dev-shm-features")
    chrome_options.add_argument("--no-sandbox")

    from concurrent.futures import ThreadPoolExecutor, as_completed

    PERFORMANCE_RE = re.compile(r"\d+\.\d+")

    def find_performance_cell(row):
        cells = row.find_elements(By.TAG_NAME, "td")
        for idx, cell in enumerate(cells):
            text = cell.text.split("\n")[0].strip()
            if "DNS" in text or "DNF" in text:
                continue
            m = PERFORMANCE_RE.search(text)
            if m:
                perf = m.group(0)
                rest = text[m.end():]
                for suffix in ("w", "h"):
                    if suffix in rest:
                        perf += suffix
                        break
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
                            has_100m = False
                            for line in lines:
                                line = line.strip()
                                if "100μ" in line and "Τελικός" in line:
                                    has_100m = True
                                    parts = line.split("Τελικός")
                                    if len(parts) > 1:
                                        after = parts[1].strip()
                                        letter = after.split()[0] if after else ""
                                        heat_name = f"Τελικός {letter}" if letter else "Τελικός"
                                    break

                            if not has_100m:
                                print(f"[W{worker_id}]   [SKIP] No 100m event on this results page")
                                continue

                            wind = ""
                            for line in lines:
                                line = line.strip()
                                if line.startswith("Άνεμος:"):
                                    wind = line.replace("Άνεμος:", "").strip()
                                    break

                            try:
                                tables = driver.find_elements(By.TAG_NAME, "table")
                                table_100m = None
                                for t in tables:
                                    before = t.find_elements(By.XPATH, "./preceding::*[contains(., '100μ') and contains(., 'Τελικός')][1]")
                                    if before:
                                        table_100m = t
                                        break
                                if table_100m:
                                    result_rows = table_100m.find_elements(By.CSS_SELECTOR, "tbody tr")
                                else:
                                    result_rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")
                            except Exception:
                                print(f"[W{worker_id}]   [WARN] Could not isolate 100m table, using all rows")
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
                                    if perf_float(performance) > 25.0:
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

    print(f"Scraped {len(all_new_results)} new performances\n")

# =========================
# DEDUP (always runs)
# =========================
unique = {}
for r in all_results:
    key = (
        r["name"],
        r["birth_year"],
        r["competition"],
        r["performance"]
    )
    unique[key] = r
if len(all_results) != len(unique):
    print(f"  Dedup: {len(all_results)} -> {len(unique)} entries")
all_results = list(unique.values())

# =========================
# CLEAN COMPETITION NAMES & LOCATIONS
# =========================
for r in all_results:
    clean_competition_location(r)
    translate_location(r)

# Backfill clubs across entries for the same athlete
club_lookup = {}
for r in all_results:
    c = r.get("club", "").strip()
    if c and c not in {"GREECE", "Greece"}:
        k1, k2 = _club_key(r["name"])
        club_lookup.setdefault(k1, set()).add(c)
        club_lookup.setdefault(k2, set()).add(c)
for r in all_results:
    c = r.get("club", "").strip()
    if not c or c in {"GREECE", "Greece"}:
        k1, k2 = _club_key(r["name"])
        for k in (k1, k2):
            if k in club_lookup:
                r["club"] = list(club_lookup[k])[0]
                break
    if r.get("club", "").strip() in {"GREECE", "Greece"}:
        r["club"] = ""

# Clean malformed performances
for r in all_results:
    r["performance"] = clean_performance(r["performance"])

# Override exclude
all_results = [r for r in all_results if r["name"] not in OVERRIDE_EXCLUDE]

# Non-Greek club filter
before = len(all_results)
all_results = [r for r in all_results if r.get("club", "").strip() not in NON_GREEK_CLUBS]
if len(all_results) < before:
    print(f"[OK] Removed {before - len(all_results)} entries by non-Greek club athletes")

# =========================
# SAVE CACHE
# =========================
print(f"\nSaving {len(all_results)} performances to cache...")
with open(CACHE_FILE, 'w', encoding='utf-8') as f:
    json.dump({
        "performances": all_results,
        "scraped_urls": sorted(scraped_urls)
    }, f, ensure_ascii=False, indent=2)
print("[OK] Cache saved successfully")

# =========================
# NORMALIZE NAMES TO GREEK
# =========================
preferred_names = {}
for r in all_results:
    has_greek = has_greek_chars(r['name'])
    key = (norm_full(r["name"]), r["birth_year"])
    if has_greek and key not in preferred_names:
        preferred_names[key] = r['name']

for r in all_results:
    has_greek = has_greek_chars(r['name'])
    if not has_greek:
        key = (norm_full(r["name"]), r["birth_year"])
        if key in preferred_names:
            r['name'] = preferred_names[key]

surname_greek = {}
for r in all_results:
    has_greek = has_greek_chars(r['name'])
    if has_greek:
        key = (nk_latin(r['name']), r['birth_year'])
        if key not in surname_greek:
            surname_greek[key] = set()
        surname_greek[key].add(r['name'])

for r in all_results:
    has_greek = has_greek_chars(r['name'])
    if not has_greek:
        key = (nk_latin(r['name']), r['birth_year'])
        if key in surname_greek and len(surname_greek[key]) == 1:
            r['name'] = list(surname_greek[key])[0]

for r in all_results:
    has_greek = has_greek_chars(r['name'])
    if not has_greek:
        greek_attempt = latin_to_greek(r['name'])
        if not any('A' <= c <= 'Z' for c in greek_attempt.upper()):
            from config import GREEK_STRONG_DIGRAPHS
            upper_name = r['name'].upper()
            surname = upper_name.split()[-1] if ' ' in upper_name else upper_name
            has_strong = False
            for d in GREEK_STRONG_DIGRAPHS:
                if d not in upper_name:
                    continue
                if d == "EV" and surname.endswith("EVA") and "EV" in surname[-3:]:
                    continue
                has_strong = True
                break
            if has_strong:
                r['name'] = greek_attempt

# =========================
# FILTER NON-GREEK ATHLETES
# =========================
before = len(all_results)
all_results = [r for r in all_results if has_greek_chars(r['name'])]
removed = before - len(all_results)
if removed:
    print(f"[OK] Removed {removed} entries by non-Greek athletes")

# Manual club fixups
for r in all_results:
    if r.get("club", "").strip():
        continue
    sur = nk_latin(r["name"])
    if sur in MANUAL_CLUB:
        r["club"] = MANUAL_CLUB[sur]

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
    if key not in season_best or perf < perf_float(season_best[key]["performance"]):
        season_best[key] = r

ranking = sorted(
    season_best.values(),
    key=lambda x: perf_float(x["performance"])
)

# =========================
# WIND-LEGAL BEST
# =========================
wind_legal_best = {}
for r in all_results:
    if not is_wind_legal(r):
        continue
    try:
        perf = perf_float(r["performance"])
    except ValueError:
        continue
    key = (norm_full(r["name"]), r["birth_year"])
    if key not in wind_legal_best or perf < perf_float(wind_legal_best[key]["performance"]):
        wind_legal_best[key] = r

wind_legal_ranking = sorted(
    wind_legal_best.values(),
    key=lambda x: perf_float(x["performance"])
)

# =========================
# WIND-AIDED ONLY
# =========================
wind_aided_only = {}
for r in all_results:
    if is_wind_legal(r):
        continue
    try:
        perf = perf_float(r["performance"])
    except ValueError:
        continue
    key = (norm_full(r["name"]), r["birth_year"])
    if key in wind_legal_best:
        continue
    if key not in wind_aided_only or perf < perf_float(wind_aided_only[key]["performance"]):
        wind_aided_only[key] = r

wind_aided_ranking = sorted(
    wind_aided_only.values(),
    key=lambda x: perf_float(x["performance"])
)

# =========================
# UPPERCASE & CLEAN UP
# =========================
uppercase_all(all_results)
clear_placeholder_clubs(all_results)
clear_lane_placeholders(all_results)

# =========================
# LOAD NOTES
# =========================
_notes_data = {}
if os.path.exists(NOTES_FILE):
    try:
        with open(NOTES_FILE, 'r', encoding='utf-8') as f:
            _notes_data = json.load(f)
    except (json.JSONDecodeError, IOError):
        _notes_data = {}

def fmt_note(r, nf=None, nd=None):
    k = f"{norm_full(r['name'])}|{r.get('birth_year','')}|{r.get('date','')}"
    return _notes_data.get(k, "")

# =========================
# EXPORT
# =========================
filename = create_excel(all_results, ranking, wind_legal_ranking, wind_aided_ranking, fmt_note, fmt_wind, fmt_comp, fmt_loc)
export_csv(all_results, filename, fmt_wind, fmt_comp, fmt_loc, fmt_note)
export_pdf(all_results, wind_legal_ranking, wind_aided_ranking, filename, lambda r, k: str(r.get(k, "")), fmt_wind, fmt_comp, fmt_loc, fmt_note)

print(f"\nExcel file: {filename}")
print(f"Total performances: {len(all_results)}")
print(f"Season best athletes: {len(ranking)}")
print(f"Wind-legal best athletes: {len(wind_legal_ranking)}")

# Stats
clubs = {}
ages = {}
for r in all_results:
    c = r.get("club", "").strip()
    if c:
        clubs[c] = clubs.get(c, 0) + 1
    by = r.get("birth_year", "")
    if by:
        ages[str(by)] = ages.get(str(by), 0) + 1
top_clubs = sorted(clubs.items(), key=lambda x: -x[1])[:10]
print(f"\nTop clubs: {', '.join(f'{c} ({n})' for c, n in top_clubs)}")
age_dist = sorted(ages.items())
print(f"Age dist: {', '.join(f'{y}: {n}' for y, n in age_dist)}")

open_file(filename)
pdf_file = filename.replace(".xlsx", ".pdf")
if os.path.exists(pdf_file):
    open_file(pdf_file)
