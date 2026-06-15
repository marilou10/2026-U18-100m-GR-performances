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

# Παίρνουμε όλες τις γραμμές του πίνακα αποτελεσμάτων
rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")

print(f"Βρέθηκαν {len(rows)} γραμμές")

results = []

for row in rows:
    cells = row.find_elements(By.TAG_NAME, "td")

    # Προστασία σε περίπτωση που κάποια γραμμή δεν είναι κανονικό αποτέλεσμα
    if len(cells) < 6:
        continue

    # Στήλη 3: Όνομα + έτος γέννησης + χώρα
    athlete_info = cells[2].text.split("\n")

    name = athlete_info[0].strip()

    try:
        birth_year = int(athlete_info[1].strip())
    except (IndexError, ValueError):
        continue

    # Στήλη 5: Σύλλογος
    club = cells[4].text.strip()

    # Στήλη 6: Επίδοση
    performance = cells[5].text.split("\n")[0].strip()

    # Φιλτράρουμε μόνο Κ18 (2009–2012)
    if 2009 <= birth_year <= 2012:
        results.append({
            "name": name,
            "birth_year": birth_year,
            "club": club,
            "performance": performance,
        })

driver.quit()

print("\nΚ18 αποτελέσματα:\n")

for athlete in results:
    print(athlete)