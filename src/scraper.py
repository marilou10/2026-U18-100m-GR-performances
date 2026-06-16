from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from webdriver_manager.chrome import ChromeDriverManager

from openpyxl import Workbook
from datetime import datetime
import os
import json
import time
import sys
import subprocess
import re
from urllib.parse import urlparse, parse_qs


# =========================
# MEETS (heats / finals)
# =========================

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
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")

    # =========================
    # HELPER FUNCTION: Extract 100m finals meIds
    # =========================
    def extract_100m_finals_meids(driver, race_url):
        """
        Given a race schedule page, find all 100m finals and return their meIds
        """
        meids = []
        try:
            # Wait for the schedule to load
            WebDriverWait(driver, 15).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "a[href*='meId']"))
            )
            
            time.sleep(2)
            
            # Get all event links
            event_links = driver.find_elements(By.CSS_SELECTOR, "a[href*='meId']")
            
            for link in event_links:
                link_text = link.text.strip()
                link_href = link.get_attribute("href")
                
                # Check if it's a 100m final
                if "100μ" in link_text and "Τελικός" in link_text:
                    # Extract meId from URL
                    match = re.search(r'meId=(\d+)', link_href)
                    if match:
                        meid = match.group(1)
                        meids.append({
                            "meId": meid,
                            "event_name": link_text,
                            "race_url": race_url
                        })
                        print(f"  ✓ Found: {link_text} (meId: {meid})")
            
            if not meids:
                print(f"  ⚠️ No 100m finals found on this page")
        
        except Exception as e:
            print(f"  ⚠️ Error extracting meIds: {e}")
        
        return meids

    # =========================
    # SCRAPE LOOP
    # =========================
    for url_index, race_url in enumerate(URLS):
        print(f"\n[{url_index + 1}/{len(URLS)}] Processing race: {race_url}")
        
        driver = None
        try:
            # Create new driver for this race
            driver = webdriver.Chrome(
                service=Service(ChromeDriverManager().install()),
                options=chrome_options
            )
            
            # Load the race schedule page to find 100m finals
            driver.get(race_url)
            
            # Extract all 100m finals meIds
            finals = extract_100m_finals_meids(driver, race_url)
            
            if not finals:
                print(f"  Skipping race (no 100m finals found)")
                driver.quit()
                continue
            
            # Now scrape each 100m final
            for final_index, final in enumerate(finals):
                meid = final["meId"]
                event_name = final["event_name"]
                
                # Build the final results page URL
                # Extract the race id from the race_url
                race_id_match = re.search(r'id=(\d+)', race_url)
                if not race_id_match:
                    print(f"    Could not extract race ID from {race_url}")
                    continue
                
                race_id = race_id_match.group(1)
                final_url = f"https://meets.rosterathletics.com/public/competitions/details/results?id={race_id}&meId={meid}"
                
                print(f"\n    [{final_index + 1}/{len(finals)}] Scraping: {event_name}")
                print(f"    URL: {final_url}")
                
                try:
                    # Load the final results page
                    driver.get(final_url)
                    
                    # Wait for table to load (max 15 seconds)
                    try:
                        WebDriverWait(driver, 15).until(
                            EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr"))
                        )
                    except TimeoutException:
                        print(f"    ⚠️ Timeout: Could not load table data from {final_url}")
                        continue
                    
                    # Extra wait for content to render
                    time.sleep(2)
                    
                    heat_name = event_name
                    body_text = driver.find_element(By.TAG_NAME, "body").text
                    lines = body_text.split("\n")

                    date = ""
                    location = ""
                    wind = ""

                    # Extract date, location, and wind
                    for i, line in enumerate(lines):
                        line = line.strip()

                        if line == "ΗΜΕΡΟΜΗΝΙΑ & ΩΡΑ" and i + 1 < len(lines):
                            date = lines[i + 1].strip()

                        if line == "ΠΟΛΗ & ΧΩΡΑ" and i + 1 < len(lines):
                            location = lines[i + 1].strip()

                        if line.startswith("Άνεμος:"):
                            wind = line.replace("Άνεμος:", "").strip()

                    rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")

                    print(f"    ✓ Found {len(rows)} athletes")

                    for row in rows:
                        try:
                            cells = row.find_elements(By.TAG_NAME, "td")

                            if len(cells) < 6:
                                continue

                            lane = cells[1].text.strip()

                            athlete_info = cells[2].text.split("\n")

                            name = athlete_info[0].strip()

                            try:
                                birth_year = int(athlete_info[1].strip())
                            except ValueError:
                                continue

                            club = cells[4].text.strip()

                            performance_text = cells[5].text.split("\n")[0].strip()

                            if performance_text == "" or "DNS" in performance_text or "DNF" in performance_text:
                                continue

                            performance = performance_text.replace("SB", "").strip()

                            if 2009 <= birth_year <= 2012:
                                all_results.append({
                                    "name": name,
                                    "birth_year": birth_year,
                                    "club": club,
                                    "performance": performance,
                                    "wind": wind,
                                    "competition": race_url,
                                    "date": date,
                                    "location": location,
                                    "heat": heat_name,
                                    "lane": lane
                                })
                        except Exception as row_error:
                            print(f"      ⚠️ Error parsing row: {row_error}")
                            continue

                except Exception as final_error:
                    print(f"    ❌ Error scraping final: {final_error}")
                    continue

        except Exception as e:
            print(f"❌ Error processing race {race_url}: {e}")
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

    # Save to cache
    print(f"\nSaving {len(all_results)} performances to cache...")
    with open(CACHE_FILE, 'w', encoding='utf-8') as f:
        json.dump(all_results, f, ensure_ascii=False, indent=2)
    print("✓ Cache saved successfully")


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
    "Name", "Birth Year", "Club", "Performance",
    "Wind", "Competition", "Date", "Location",
    "Heat", "Lane"
])

for r in all_results:
    ws1.append([
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
    "Rank", "Name", "Birth Year", "Club",
    "Best Performance", "Heat", "Lane"
])

for i, r in enumerate(ranking, 1):
    ws2.append([
        i,
        r["name"],
        r["birth_year"],
        r["club"],
        r["performance"],
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

# Open file with cross-platform support
if sys.platform == "win32":
    os.startfile(filename)
elif sys.platform == "darwin":
    subprocess.run(["open", filename])
else:  # Linux
    subprocess.run(["xdg-open", filename])
