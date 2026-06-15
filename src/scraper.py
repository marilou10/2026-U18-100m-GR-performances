from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time

URLS = [
    "https://meets.rosterathletics.com/public/competitions/details/results?id=29233&meId=432494",
    "https://meets.rosterathletics.com/public/competitions/details/results?id=29233&meId=432495",
]

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install())
)

results = []

for url in URLS:
    print("\n" + "=" * 60)
    print("Επεξεργασία:", url)

    driver.get(url)

    # Περιμένουμε να φορτώσει η σελίδα
    time.sleep(5)

    # ========================
    # Στοιχεία αγώνα
    # ========================

    competition = driver.title.replace("Roster Athletics · ", "").strip()

    page_text = driver.find_element(By.TAG_NAME, "body").text

    date = ""
    location = ""
    wind = ""

    lines = page_text.split("\n")

    for i, line in enumerate(lines):
        line = line.strip()

        if line == "ΗΜΕΡΟΜΗΝΙΑ & ΩΡΑ" and i + 1 < len(lines):
            date = lines[i + 1].strip()

        elif line == "ΠΟΛΗ & ΧΩΡΑ" and i + 1 < len(lines):
            location = lines[i + 1].strip()

        elif line.startswith("Άνεμος:"):
            wind = line.replace("Άνεμος:", "").strip()

    print(f"Αγώνας: {competition}")
    print(f"Ημερομηνία: {date}")
    print(f"Τοποθεσία: {location}")
    print(f"Άνεμος: {wind}")

    # ========================
    # Αποτελέσματα
    # ========================

    rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")

    print(f"Βρέθηκαν {len(rows)} γραμμές")

    for row in rows:
        cells = row.find_elements(By.TAG_NAME, "td")

        # Παραλείπουμε άκυρες γραμμές
        if len(cells) < 6:
            continue

        athlete_info = cells[2].text.split("\n")

        name = athlete_info[0].strip()

        try:
            birth_year = int(athlete_info[1].strip())
        except (IndexError, ValueError):
            continue

        club = cells[4].text.strip()

        performance = cells[5].text.split("\n")[0].strip()

        # Κ18 (2009–2010) + Κ16 με δικαίωμα συμμετοχής (2011–2012)
        if 2009 <= birth_year <= 2012:
            results.append({
                "name": name,
                "birth_year": birth_year,
                "club": club,
                "performance": performance,
                "wind": wind,
                "competition": competition,
                "date": date,
                "location": location,
            })

driver.quit()

# ========================
# Συνολικά αποτελέσματα
# ========================

print("\n" + "=" * 60)
print("ΣΥΝΟΛΙΚΑ Κ18 ΑΠΟΤΕΛΕΣΜΑΤΑ")
print("=" * 60)

for athlete in results:
    print(athlete)

print("\nΣυνολικές Κ18 εγγραφές:", len(results))