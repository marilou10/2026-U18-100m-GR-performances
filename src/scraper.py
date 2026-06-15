from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time

URL = (
    "https://meets.rosterathletics.com/public/"
    "competitions/details/results?id=29233&meId=432494"
)

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install())
)

driver.get(URL)

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

print("\n" + "=" * 60)
print("ΣΤΟΙΧΕΙΑ ΑΓΩΝΑ")
print("=" * 60)
print("Αγώνας:", competition)
print("Ημερομηνία:", date)
print("Τοποθεσία:", location)
print("Άνεμος:", wind)

# ========================
# Αποτελέσματα Κ18
# ========================

rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")

print(f"\nΒρέθηκαν {len(rows)} γραμμές")

results = []

for row in rows:
    cells = row.find_elements(By.TAG_NAME, "td")

    # Παραλείπουμε γραμμές που δεν έχουν τα αναμενόμενα δεδομένα
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

print("\n" + "=" * 60)
print("Κ18 ΑΠΟΤΕΛΕΣΜΑΤΑ")
print("=" * 60)

for athlete in results:
    print(athlete)