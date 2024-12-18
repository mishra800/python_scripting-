import os
import time
from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PIL import Image
import pytesseract
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

# Path to the ChromeDriver (use webdriver_manager to automatically download and use the latest version)
service = Service(ChromeDriverManager().install())  # This will automatically manage the chromedriver version

# Create the WebDriver instance
driver = webdriver.Chrome(service=service)

# Specify the path to Tesseract executable
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'

# Save directory for HTML files
SAVE_DIR = 'transactions'
os.makedirs(SAVE_DIR, exist_ok=True)

# WebDriver setup
driver.maximize_window()
driver.get('https://freesearchigrservice.maharashtra.gov.in/')

# Wait object for dynamic elements
wait = WebDriverWait(driver, 20)  # Increased wait time to 20 seconds

# Step 1: Select "Rest of Maharashtra"
try:
    wait.until(EC.presence_of_element_located((By.ID, 'rdbRestOfMaharashtra')))
    wait.until(EC.element_to_be_clickable((By.ID, 'rdbRestOfMaharashtra'))).click()
    print("Clicked on 'Rest of Maharashtra' option")
except Exception as e:
    print(f"Error while clicking on the element: {e}")

# Step 2: Fill form details
Select(driver.find_element(By.ID, 'ddlYear')).select_by_visible_text('2023')
Select(driver.find_element(By.ID, 'ddlDistrict')).select_by_visible_text('Pune')
Select(driver.find_element(By.ID, 'ddlTahsil')).select_by_visible_text('Haveli')
Select(driver.find_element(By.ID, 'ddlVillage')).select_by_visible_text('Wakad')

# Captcha solving function
def solve_captcha():
    captcha_img = driver.find_element(By.ID, 'imgCaptcha')
    captcha_img.screenshot('captcha.png')
    captcha_text = pytesseract.image_to_string(Image.open('captcha.png')).strip()
    print(f"Solved Captcha: {captcha_text}")
    return captcha_text

# Step 3: Iterate over property numbers
for property_number in range(1, 11):  # Adjust range as required
    try:
        # Input property number
        driver.find_element(By.ID, 'txtPropertyNumber').clear()
        driver.find_element(By.ID, 'txtPropertyNumber').send_keys(str(property_number))

        # Solve captcha and fill it
        captcha_text = solve_captcha()
        driver.find_element(By.ID, 'txtCaptcha').clear()
        driver.find_element(By.ID, 'txtCaptcha').send_keys(captcha_text)

        # Submit the form
        driver.find_element(By.ID, 'btnSearch').click()
        time.sleep(5)  # Wait for results

        # Check for "No Data Found"
        if "No Data Found" in driver.page_source:
            print(f"No data found for property number {property_number}")
            continue

        # Step 4: Scrape transactions on current page
        while True:
            rows = driver.find_elements(By.CSS_SELECTOR, '#gridResults tr')  # Adjust selector
            for row in rows[1:]:  # Skip header row
                try:
                    columns = row.find_elements(By.TAG_NAME, 'td')
                    doc_num = columns[0].text
                    sro_code = columns[1].text
                    year = '2023'

                    # Click the document link
                    doc_link = columns[-1].find_element(By.TAG_NAME, 'a')
                    doc_link.click()

                    # Switch to new tab and save HTML
                    driver.switch_to.window(driver.window_handles[-1])
                    html_content = driver.page_source
                    filename = f"{doc_num}_{sro_code}_{year}.html"
                    with open(os.path.join(SAVE_DIR, filename), 'w', encoding='utf-8') as f:
                        f.write(html_content)
                    print(f"Saved: {filename}")

                    # Close the tab and return to the main page
                    driver.close()
                    driver.switch_to.window(driver.window_handles[0])
                except Exception as e:
                    print(f"Error processing row: {e}")

            # Handle pagination
            try:
                next_button = driver.find_element(By.ID, 'btnNext')
                if next_button.is_enabled():
                    next_button.click()
                    time.sleep(5)
                else:
                    break
            except Exception:
                break
    except Exception as e:
        print(f"Error with property number {property_number}: {e}")

# Close the WebDriver
driver.quit()
