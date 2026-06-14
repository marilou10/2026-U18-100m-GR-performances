
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

# Περιμένουμε λίγα δευτερόλεπτα να φορτώσει η σελίδα
time.sleep(5)

athletes = driver.find_elements(By.CSS_SELECTOR, "a.athlete-name")

print(f"Βρέθηκαν {len(athletes)} αθλήτριες:\n")

for i, athlete in enumerate(athletes, start=1):
    print(f"{i}. {athlete.text}")

input("\nPress Enter to close...")

driver.quit()