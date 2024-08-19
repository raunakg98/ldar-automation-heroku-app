import gspread
from oauth2client.service_account import ServiceAccountCredentials
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

# Google Sheets API setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name('sunny-strategy-428116-g6-77eaa542d7d3.json', scope)
client = gspread.authorize(creds)

# Open the Google Sheet
sheet = client.open_by_url('https://docs.google.com/spreadsheets/d/1F-Ga7tOJI5_gvJdq5jtG6J-g5sSRm54g-9b6w_G62LY/edit?usp=sharing')

# Path to the uploaded chromedriver
chromedriver_path = r'C:\Users\rauna\Downloads\chrome-win64\chrome-win64\chrome.exe'

# Initialize the WebDriver
chrome_options = webdriver.ChromeOptions()
chrome_options.binary_location = chromedriver_path
driver = webdriver.Chrome(options=chrome_options)

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
    # Read all relevant cells in a single API call starting from the specified row
    animal_names = worksheet.col_values(1)[start_row-1:150]  # Adjust range as needed
    updates = []

    for idx, animal_name in enumerate(animal_names, start=start_row):
        if not animal_name or "Animal" in animal_name:
            # If the cell in column A is empty or contains the term "Animal", skip this row
            continue

        try:
            # Find the search box and enter the animal name
            search_box = WebDriverWait(driver, 12).until(
                EC.presence_of_element_located((By.ID, 'edit-title'))
            )
            search_box.clear()
            search_box.send_keys(animal_name)

            # Click the "Apply" button
            apply_button = WebDriverWait(driver, 12).until(
                EC.element_to_be_clickable((By.ID, 'edit-submit-adoptable-animals'))
            )
            apply_button.click()

            # Wait for 5 seconds
            time.sleep(5)

            # Check if the "view-empty" element is present
            if driver.find_elements(By.CLASS_NAME, 'view-empty'):
                updates.append((idx, 'Animal not found'))
                print(f"Animal not found for row {idx}")
                continue  # Move to the next animal

            # Check if the "view-content" element is present
            if not driver.find_elements(By.CLASS_NAME, 'view-content'):
                updates.append((idx, 'Animal not found'))
                print(f"Animal not found for row {idx}")
                continue  # Move to the next animal

            # Find the correct animal link and click it
            if not find_animal_and_click(animal_name):
                updates.append((idx, 'Animal not found'))
                print(f"Exact match not found for row {idx}")
                continue  # Move to the next animal

            # Switch to the new tab
            driver.switch_to.window(driver.window_handles[1])

            # Get the main picture URL
            img_element = WebDriverWait(driver, 32).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, '.rgPetDetailsLargePhoto img'))
            )
            img_url = img_element.get_attribute('src')
            formatted_url = f'=image("{img_url}", 4, 100, 100)'

            # Collect the update
            updates.append((idx, formatted_url))
            print(f"Updated row {idx} with image URL")

            # Close the current tab and switch back to the main tab
            driver.close()
            driver.switch_to.window(driver.window_handles[0])

        except Exception as e:
            updates.append((idx, 'Error occurred'))
            print(f"Error for row {idx}: {str(e)}")

    # Batch update the Google Sheet
    for idx, update in updates:
        worksheet.update_acell(f'B{idx}', update)

try:
    # Open the Lucky Dog Animal Rescue adopt page once
    driver.get('https://www.luckydoganimalrescue.org/adopt')

    # List of worksheet indices to process
    worksheets_to_process = [5]
    start_row = 2 # Change this to the desired starting row

    # Process each worksheet
    for worksheet_index in worksheets_to_process:
        process_worksheet(sheet, worksheet_index, start_row)

finally:
    # Close the WebDriver
    driver.quit()
