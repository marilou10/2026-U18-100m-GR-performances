from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

from openpyxl import Workbook
from datetime import datetime
import os
import json
import time


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
    with open(CACHE_FILE, 'r', encoding='utf-8') as f:
        all_results = json.load(f)
    print(f"Loaded {len(all_results)} performances from cache\n")
else:
    # =========================
    # SCRAPE LOOP
    # =========================
    driver = webdriver.Chrome(
        service=Service(ChromeDriverManager().install())
    )

    for url in URLS:
        print("\nΕπεξεργασία:", url)
        
        try:
            driver.get(url)
            
            # Wait for table to load (max 10 seconds)
            WebDriverWait(driver, 10).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, "tbody tr"))
            )
            
            heat_name = ""
            body_text = driver.find_element(By.TAG_NAME, "body").text
            lines = body_text.split("\n")

            for i, line in enumerate(lines):
                line = line.strip()

                if "100μ" in line and "Τελικός" in line:
                    heat_name = line
                    break

            date = ""
            location = ""
            wind = ""

            for i, line in enumerate(lines):
                line = line.strip()

                if line == "ΗΜΕΡΟΜΗΝΙΑ & ΩΡΑ" and i + 1 < len(lines):
                    date = lines[i + 1].strip()

                if line == "ΠΟΛΗ & ΧΩΡΑ" and i + 1 < len(lines):
                    location = lines[i + 1].strip()

                if line.startswith("Άνεμος:"):
                    wind = line.replace("Άνεμος:", "").strip()

            rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")

            print("Βρέθηκαν", len(rows), "γραμμές")

            for row in rows:
                cells = row.find_elements(By.TAG_NAME, "td")

                if len(cells) < 6:
                    continue

                lane = cells[1].text.strip()

                athlete_info = cells[2].text.split("\n")

                name = athlete_info[0].strip()

                try:
                    birth_year = int(athlete_info[1].strip())
                except:
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
                        "competition": url,
                        "date": date,
                        "location": location,
                        "heat": heat_name,
                        "lane": lane
                    })

        except Exception as e:
            print(f"Error scraping {url}: {e}")
            continue

    driver.quit()

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
    print("Cache saved successfully")


# =========================
# SEASON BEST
# =========================
season_best = {}

for r in all_results:
    try:
        perf = float(r["performance"])
    except:
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

print("\nDONE")
print("Excel file:", filename)
print("Total performances:", len(all_results))
print("Season best athletes:", len(ranking))

os.startfile(filename)
