import os
import pandas as pd
import time
import sys
import random
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Set console to UTF-8 encoding
sys.stdout.reconfigure(encoding='utf-8')

# Function to initialize WebDriver with options
def initialize_driver():
    chrome_options = Options()
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--incognito")
    chrome_options.add_experimental_option("debuggerAddress", "localhost:8989")
    return webdriver.Chrome(options=chrome_options)

# Function to close the WebDriver
def close_driver(driver):
    driver.quit()

# Initialize WebDriver
driver = initialize_driver()

# Load the Excel file
input_file = r'D:/Selenium/Group5-2.xlsx'
try:
    df = pd.read_excel(input_file)
except Exception as e:
    print(f"Error loading Excel file: {e}", file=sys.stderr)
    close_driver(driver)
    exit()

# Checkpoint file to track progress
checkpoint_file = 'checkpoint.txt'

# Load the last processed index from the checkpoint file
if os.path.exists(checkpoint_file):
    with open(checkpoint_file, 'r') as file:
        last_processed_index = int(file.read().strip())
else:
    last_processed_index = -1  # Start from the beginning if no checkpoint exists

# Create a list to collect results
results_list = []

# URL to be accessed
base_url = 'https://dogsearch.moag.gov.il/#/pages/pets'

# Function to select Hebrew language only once
def select_language_once():
    try:
        driver.get(base_url)
        WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "div.p-dropdown-trigger.ng-tns-c58-3"))
        ).click()
        WebDriverWait(driver, 2).until(
            EC.element_to_be_clickable((By.XPATH, "//li[@aria-label='עברית']"))
        ).click()
        time.sleep(random.uniform(1, 2))
    except Exception as e:
        print(f"Error interacting with language dropdown: {e}", file=sys.stderr)
        close_driver(driver)
        exit()

select_language_once()

# Update dictionary to include the new attribute
dog_info_ids = {
    'שם בעל החיים': 'for-name',
    'מין': 'for-gender',
    'גזע': 'for-breed',
    'צבע': 'for-color',
    'תאריך לידה': 'for-birthDate',
    'בעלים': 'for-owner',
    'כתובת': 'for-address',
    'ישוב': 'for-city',
    'טלפון': 'for-phone1',
    'עיקור/סירוס': 'for-neutering',
    'מספר חיסוני כלבת': 'for-rabies-vaccine',
    'חיסון קודם לכלבת': 'for-rabies-p-vaccine-date',
    'חיסון אחרון לכלבת': 'for-rabies-vaccine-date',
    'רופא מחסן': 'for-vet',
    'בדיקת נוגדני כלבת': 'for-viewReport',
    'מספר רישיון': 'for-license',
    'הנפקת רישיון': 'for-license-date-start',
    'שם הרשות': 'for-domain',
    'תאריך עדכון אחרון': 'for-license-latest-update',
    'סטטוס': 'for-status',
    'כרטיס כלב ומספר שבב': 'head_resulte'  # New attribute added
}

def scrape_page(search_name):
    global driver
    print(f"Scraping results for: {search_name}")

    data_found = False

    try:
        WebDriverWait(driver, 2).until(
            EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'app-locate-results > div'))
        )
        result_divs = driver.find_elements(By.CSS_SELECTOR, 'app-locate-results > div')

        if not result_divs:
            print("No results found for this name, skipping...")
            return False  
        else:
            data_found = True
            for result_div in result_divs:
                data = {field: None for field in dog_info_ids.keys()}
                data['Search Name'] = search_name

                for field_name, css_id in dog_info_ids.items():
                    try:
                        field_element = result_div.find_element(By.CSS_SELECTOR, f'div#{css_id}')
                        data[field_name] = field_element.text.strip() if field_element else 'null'
                    except Exception as e:
                        data[field_name] = 'null'

                print(f"Appending result: {data}")
                results_list.append(data)
    except Exception as e:
        print(f"Error scraping page: {e}", file=sys.stderr)

    return data_found

def go_to_next_page():
    global driver
    try:
        next_button = driver.find_element(By.CSS_SELECTOR, 'button.p-paginator-next.p-paginator-element.p-link.p-ripple')
        if next_button.is_enabled() and 'p-disabled' not in next_button.get_attribute('class'):
            next_button.click()
            WebDriverWait(driver, 2).until(
                EC.presence_of_all_elements_located((By.CSS_SELECTOR, 'app-locate-results > div'))
            )
            time.sleep(random.uniform(1, 2))
            print("Navigating to the next page...")
            return True
        else:
            print("Next page button is disabled or not available.")
    except Exception as e:
        print(f"Error navigating to the next page: {e}", file=sys.stderr)
    return False

def scrape_all_pages(search_name):
    global driver
    while True:
        data_found = scrape_page(search_name)
        if not go_to_next_page():
            if not data_found:
                print("No data found, proceeding to next search.")
                time.sleep(random.uniform(1, 2))
            break

    if data_found:
        print("Data found, waiting before next search.")
        time.sleep(random.uniform(1, 2))

def handle_captcha():
    global driver
    try:
        captcha_button = WebDriverWait(driver, 5).until(
            EC.presence_of_element_located((By.ID, 'amzn-captcha-verify-button'))
        )
        if captcha_button:
            print("CAPTCHA detected, attempting to resolve...")
            captcha_button.click()
            time.sleep(random.uniform(15, 25))
    except Exception as e:
        print(f"Error handling CAPTCHA: {e}", file=sys.stderr)

search_counter = 0


# Load existing results if file exists
output_file = r'D:/Selenium/Group5-2Result.xlsx'
if os.path.exists(output_file):
    try:
        existing_results_df = pd.read_excel(output_file)
        existing_results_list = existing_results_df.to_dict(orient='records')
    except Exception as e:
        print(f"Error loading existing results: {e}", file=sys.stderr)
        existing_results_list = []
else:
    existing_results_list = []

# Start iterating from the last processed index + 1
for index, row in df.iterrows():
    if index <= last_processed_index:
        continue  # Skip already processed rows

    search_name = row.get('Full Name (English + Hebrew)', '')
    if not search_name:
        continue

    full_url = f'{base_url}?search={search_name}'
    driver.get(full_url)

    handle_captcha()

    try:
        WebDriverWait(driver, 2).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, 'input[formcontrolname="locPet"]'))
        )

        search_box = driver.find_element(By.CSS_SELECTOR, 'input[formcontrolname="locPet"]')
        search_box.clear()
        search_box.send_keys(search_name)
        
        search_button = driver.find_element(By.ID, 'locPetButton')
        search_button.click()
        
        time.sleep(random.uniform(1, 2))

        scrape_all_pages(search_name)

    except Exception as e:
        print(f"Exception occurred: {e}", file=sys.stderr)

    search_counter += 1

    # Save the current progress in the checkpoint file
    with open(checkpoint_file, 'w') as file:
        file.write(str(index))

   

    if search_counter % 100 == 0:
        print("Taking a 10-minute break after 100 searches.")
        time.sleep(60)

    if search_counter % 1 == 0:
        driver.execute_script("location.reload();")
        print("Page reloaded after every 1 search.")

    # Append new results to existing results
    try:
        results_df = pd.DataFrame(results_list)
        if not results_df.empty:
            if os.path.exists(output_file):
                existing_results_df = pd.read_excel(output_file)
                combined_df = pd.concat([existing_results_df, results_df], ignore_index=True)
            else:
                combined_df = results_df
            
            combined_df.to_excel(output_file, index=False)
            results_list = []  # Clear the results_list after saving
    except Exception as e:
        print(f"Error saving results to Excel: {e}", file=sys.stderr)

close_driver(driver)
