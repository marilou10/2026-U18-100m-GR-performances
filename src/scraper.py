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

time.sleep(5)

rows = driver.find_elements(
    By.CSS_SELECTOR,
    "tbody tr"
)

print(f"Βρέθηκαν {len(rows)} γραμμές\n")

for i, row in enumerate(rows, start=1):
    print("=" * 60)
    print(f"ΓΡΑΜΜΗ {i}")

    cells = row.find_elements(By.TAG_NAME, "td")

    for j, cell in enumerate(cells, start=1):
        print(f"Στήλη {j}: '{cell.text}'")

input("\nPress Enter to close...")
driver.quit()