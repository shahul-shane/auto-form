import time
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
import tkinter as tk
from tkinter import simpledialog


# Load Excel data
file_path = "form_data.xlsx"
df = pd.read_excel(file_path)

def prompt_attendees_gui(row_number):
    root = tk.Tk()
    root.withdraw()  # Hide the main window
    value = simpledialog.askstring(
        title="Manual Input Required",
        prompt=f"Enter Total Attendees (online + offline) for row {row_number}:"
    )
    root.destroy()
    return value

# Setup ChromeDriver
CHROMEDRIVER_PATH = "./chromedriver.exe"
service = Service(CHROMEDRIVER_PATH)
driver = webdriver.Chrome(service=service)

form_url = "https://docs.google.com/forms/d/e/1FAIpQLSc33QW9-7z5jkLqCtC6WLn-W2c7z4LAvWnOqmERVJSO0fa5Iw/viewform"

# Create timestamp for column header
timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

for index, row in df.iterrows():
    print(f"\n🚀 Submitting row {index + 1}")
    driver.get(form_url)

    try:
        WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.XPATH, "//form"))
        )
                # === Fill Email ===
        email_value = str(row.get("Email", "")).strip()

        if not email_value or email_value.lower() == "nan":
            email_value = simpledialog.askstring(
                title="Manual Email Required",
                prompt=f"Enter Email for row {index + 1}:"
            )

        # Adjust this if the email field is NOT type='email'
        try:
            email_input = driver.find_element(By.XPATH, "//input[@type='email']")
        except:
            email_input = driver.find_element(By.XPATH, "(//input[@type='text'])[1]")

        email_input.send_keys(email_value)
        print(f"✅ Filled Email with: {email_value}")


        # === Fill Date ===
        session_date = datetime.now().strftime("%m/%d/%Y")
        driver.find_element(By.XPATH, "//input[@type='date']").send_keys(session_date)

        # === Textarea Inputs ===
        driver.find_element(By.XPATH, "(//textarea)[1]").send_keys(str(row["SME"]))
        driver.find_element(By.XPATH, "(//textarea)[2]").send_keys(str(row["Batch Name"]))
        driver.find_element(By.XPATH, "(//textarea)[3]").send_keys(str(row["Course Event"]))

        # === Total Attendees Input (before Comments) ===
        attendees_value = str(row.get("Total attendees (online + offline)", "")).strip()

        if not attendees_value or attendees_value.lower() == "nan":
            attendees_value = prompt_attendees_gui(index + 1)

        attendees_input = driver.find_element(By.XPATH, "(//textarea)[3]/following::input[@type='text'][1]")
        attendees_input.send_keys(attendees_value)
        print(f"✅ Filled Total attendees with: {attendees_value}")
        time.sleep(1)

        # === Ratings Less Than 4 ===
        ratings_raw = row.get("How many ratings less than 4 for today's session? (in any category)", "")
        ratings_value = str(ratings_raw).strip()

        if not ratings_value or ratings_value.lower() == "nan":
            ratings_value = simpledialog.askstring(
                title="Manual Input Required",
                prompt=f"Enter number of ratings < 4 for row {index + 1}:"
            )

        ratings_input = driver.find_element(By.XPATH, "(//textarea)[3]/following::input[@type='text'][2]")
        ratings_input.send_keys(ratings_value)
        print(f"✅ Filled Ratings < 4 with: {ratings_value}")
        time.sleep(1)

        # === Comments ===
        driver.find_element(By.XPATH, "(//textarea)[4]").send_keys(str(row["Comments"]))

        # === Radio Buttons ===
        radio_questions = [
            "Camera On While Delivering",
            "Class Started on Time",
            "Zoom Poll Taken / Feedback Poll Taken",
            "Resolution of Non Tech query",
            "Resolution of Tech query",
            "Refer and earn slide shown",
            "Participant Engagement",
            "Technical glitch (if any)",
            "Was there any disruption during the session?"
        ]

        radiogroups = driver.find_elements(By.XPATH, "//div[@role='radiogroup']")

        for i, question in enumerate(radio_questions):
            try:
                answer_raw = str(row[question]).strip().lower()
                answer = "Yes" if answer_raw == "yes" else "No" if answer_raw == "no" else None

                if not answer:
                    raise ValueError(f"Invalid answer: '{row[question]}' for '{question}'")

                group = radiogroups[i]
                option = group.find_element(By.XPATH, f".//div[@data-value='{answer}']")

                driver.execute_script("arguments[0].scrollIntoView(true);", option)
                ActionChains(driver).move_to_element(option).pause(0.2).click(option).perform()

                print(f"✅ Selected '{answer}' for: {question}")
                time.sleep(0.1)

            except Exception as e:
                print(f"❌ Error selecting for '{question}': {e}")

        # === Submit ===
        driver.find_element(By.XPATH, '//span[text()="Submit"]').click()
        print(f"✅ Row {index + 1} submitted!")

        df.at[index, timestamp] = "Submitted"

    except Exception as e:
        print(f"❌ Error on row {index + 1}: {e}")
        df.at[index, timestamp] = f"Error: {e}"

    time.sleep(2)

driver.quit()

# Save Excel with updated status
df.to_excel(file_path, index=False)
print(f"\n📄 Excel updated with new column: {timestamp}")
 # type: ignore