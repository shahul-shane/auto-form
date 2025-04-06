import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import traceback

# Load Excel data
file_path = "form_data.xlsx"
df = pd.read_excel(file_path)

# Add a 'Status' column if it doesn't exist
if "Status" not in df.columns:
    df["Status"] = ""

# Setup ChromeDriver
CHROMEDRIVER_PATH = "./chromedriver.exe"
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service)

form_url = "https://docs.google.com/forms/d/e/1FAIpQLSc33QW9-7z5jkLqCtC6WLn-W2c7z4LAvWnOqmERVJSO0fa5Iw/viewform"

for index, row in df.iterrows():
    # Skip rows already marked as submitted
    

    print(f"\n🚀 Submitting row {index + 1}")
    driver.get(form_url)
    
    try:
        

        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "(//textarea)[1]"))
        )

        # Fill the form
        driver.find_element(By.XPATH, "(//textarea)[1]").send_keys(str(row["SME"]))
        driver.find_element(By.XPATH, "(//textarea)[2]").send_keys(str(row["Batch Name"]))
        driver.find_element(By.XPATH, "(//textarea)[3]").send_keys(str(row["Course Event"]))
        driver.find_element(By.XPATH, "(//textarea)[4]").send_keys(str(row["Comments"]))
        time.sleep(1)

        # Submit
        driver.find_element(By.XPATH, '//span[text()="Submit"]').click()
        print(f"✅ Row {index + 1} submitted!")

        # Mark as submitted
        df.at[index, "Status"] = "Submitted"

    except Exception as e:
        print(f"❌ Error on row {index + 1}: {type(e).__name__} - {e}")
        traceback.print_exc()
        df.at[index, "Status"] = "Failed"

    time.sleep(2)

driver.quit()

# Save updated Excel with status
try:
    df.to_excel(file_path, index=False)
    print("\n📄 Excel file updated with submission status.")
except PermissionError:
    print("\n⚠️ Cannot write to Excel file. Please make sure it's closed and not open in Excel.")
