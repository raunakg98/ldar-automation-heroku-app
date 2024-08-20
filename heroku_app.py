import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import json
from dotenv import load_dotenv

# Load environment variables from .env
load_dotenv()

# Access the credentials from environment variable
creds_json = os.getenv("GOOGLE_SHEETS_CREDENTIALS")
creds = json.loads(creds_json)

# Google Sheets API setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_dict(creds, scope)
client = gspread.authorize(creds)

# Open the Google Sheet
sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1vtpN5oMHuS0uicTZGg3oosPZRFBFsWy6cYJR8Ymk9pY/edit?usp=sharing')


# ---------------------------- for heroku use ----------------------------
# Initialize the WebDriver
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--headless")
chrome_options.add_argument("--disable-gpu")
chrome_options.add_argument("--no-sandbox")
chrome_options.add_argument("--disable-dev-shm-usage")
chrome_options.binary_location = os.getenv("GOOGLE_CHROME_BIN", "/app/.apt/usr/bin/google-chrome")

# Initialize the WebDriver without `executable_path`
driver = webdriver.Chrome(service=webdriver.chrome.service.Service(os.getenv("CHROMEDRIVER_PATH", "/app/.chromedriver/bin/chromedriver")), options=chrome_options)

# ---------------------------------------------------------------------

# ---------------------------- for local use ----------------------------
# # Path to the uploaded chromedriver
# chromedriver_path = r'C:\Users\rauna\Downloads\chrome-win64\chrome-win64\chrome.exe'

# # Initialize the WebDriver
# chrome_options = webdriver.ChromeOptions()
# chrome_options.binary_location = chromedriver_path
# driver = webdriver.Chrome(options=chrome_options)
# ---------------------------------------------------------------------

def find_animal_and_click(animal_name):
    try:
        results = driver.find_elements(By.CSS_SELECTOR, '.views-field-title a')
        for result in results:
            if result.text.lower() == animal_name.lower():
                result.click()
                return True
        return False
    except Exception as e:
        print(f"Error in find_animal_and_click: {e}")
        return False

def process_worksheet(sheet, worksheet_index, start_row):
    worksheet = sheet.get_worksheet(worksheet_index)
    animal_names = worksheet.col_values(1)[start_row-1:150]
    updates = []

    for idx, animal_name in enumerate(animal_names, start=start_row):
        if not animal_name or "Animal" in animal_name:
            continue

        try:
            search_box = WebDriverWait(driver, 12).until(
                EC.presence_of_element_located((By.ID, 'edit-title'))
            )
            search_box.clear()
            search_box.send_keys(animal_name)

            apply_button = WebDriverWait(driver, 12).until(
                EC.element_to_be_clickable((By.ID, 'edit-submit-adoptable-animals'))
            )
            apply_button.click()

            time.sleep(5)

            if driver.find_elements(By.CLASS_NAME, 'view-empty'):
                updates.append((idx, 'Animal not found'))
                continue

            if not driver.find_elements(By.CLASS_NAME, 'view-content'):
                updates.append((idx, 'Animal not found'))
                continue

            if not find_animal_and_click(animal_name):
                updates.append((idx, 'Animal not found'))
                continue

            driver.switch_to.window(driver.window_handles[1])

            img_element = WebDriverWait(driver, 32).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '.rgPetDetailsLargePhoto img'))
            )
            img_url = img_element.get_attribute('src')
            formatted_url = f'=image("{img_url}", 4, 100, 100)'

            updates.append((idx, formatted_url))

            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        except Exception as e:
            updates.append((idx, 'Error occurred'))

    for idx, update in updates:
        worksheet.update_acell(f'B{idx}', update)

try:
    driver.get('https://www.luckydoganimalrescue.org/adopt')

    worksheets_to_process = [5]
    start_row = 2

    for worksheet_index in worksheets_to_process:
        process_worksheet(sheet, worksheet_index, start_row)

finally:
    driver.quit()