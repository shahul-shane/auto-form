import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Load Excel data
df = pd.read_excel("form_data.xlsx")

# Setup ChromeDriver
CHROMEDRIVER_PATH = "./chromedriver.exe"
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service)

form_url = "https://docs.google.com/forms/d/e/1FAIpQLSc33QW9-7z5jkLqCtC6WLn-W2c7z4LAvWnOqmERVJSO0fa5Iw/viewform"

for index, row in df.iterrows():
    print(f"\n🚀 Submitting row {index + 1}")
    driver.get(form_url)

    # Wait until the first textarea loads
    WebDriverWait(driver, 15).until(
        EC.presence_of_element_located((By.XPATH, "(//textarea)[1]"))
    )

    # Fill the textareas in order
    driver.find_element(By.XPATH, "(//textarea)[1]").send_keys(str(row["SME"]))
    driver.find_element(By.XPATH, "(//textarea)[2]").send_keys(str(row["Batch Name"]))
    driver.find_element(By.XPATH, "(//textarea)[3]").send_keys(str(row["Course Event"]))
    driver.find_element(By.XPATH, "(//textarea)[4]").send_keys(str(row["Comments"]))

    time.sleep(1)

    # Submit the form
    submit_button = driver.find_element(By.XPATH, '//span[text()="Submit"]')
    submit_button.click()

    print(f"✅ Row {index + 1} submitted!")
    time.sleep(2)

driver.quit()
