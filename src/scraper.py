from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

URL = (
    "https://meets.rosterathletics.com/public/"
    "competitions/details/results?id=29233&meId=432494"
)

driver = webdriver.Chrome(
    service=Service(ChromeDriverManager().install())
)

print("Opening browser...")

driver.get(URL)

print("Page title:")
print(driver.title)

input("\nPress Enter to close the browser...")

driver.quit()