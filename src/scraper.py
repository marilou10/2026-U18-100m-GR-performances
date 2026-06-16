from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager

from openpyxl import Workbook
from datetime import datetime
import os
import time


# =========================
# MEETS (heats / finals)
# =========================
URLS = [
    {
        "url": "https://meets.rosterathletics.com/public/competitions/details/results?id=29233&meId=432494",
        "heat": "Final A"
    },
    {
        "url": "https://meets.rosterathletics.com/public/competitions/details/results?id=29233&meId=432495",
        "heat": "Final B"
    },
]


driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install())
)

all_results = []


# =========================
# SCRAPE LOOP
# =========================
for item in URLS:
    url = item["url"]
    heat_name = item["heat"]

    print("\nΕπεξεργασία:", url)
    print("Heat:", heat_name)

    driver.get(url)
    time.sleep(5)
    
    competition = driver.title.replace("Roster Athletics · ", "").strip()

    body_text = driver.find_element(By.TAG_NAME, "body").text
    lines = body_text.split("\n")

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

        lane = cells[1].text.strip()   # ⭐ FIX: lane

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
                "competition": competition,
                "date": date,
                "location": location,
                "heat": heat_name,
                "lane": lane   # ⭐ NEW
            })


driver.quit()


# =========================
# SEASON BEST
# =========================
season_best = {}

for r in all_results:
    try:
        perf = float(r["performance"])
    except:
        continue

    name = r["name"]

    if name not in season_best:
        season_best[name] = r
    else:
        if perf < float(season_best[name]["performance"]):
            season_best[name] = r

ranking = sorted(season_best.values(), key=lambda x: float(x["performance"]))


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