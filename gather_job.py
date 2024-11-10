from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl

# Initialize WebDriver
driver = webdriver.Chrome()

# Log in to Naukri
driver.get("https://www.naukri.com/nlogin/login")
email = "EMAIL"
password = "PASSWORD"

try:
    email_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "usernameField"))
    )
    password_input = driver.find_element(By.ID, "passwordField")
    
    email_input.send_keys(email)
    password_input.send_keys(password)
    password_input.send_keys(Keys.RETURN)
    
    print("Logged in successfully.")
except Exception as e:
    print("Login failed:", e)
    driver.quit()

time.sleep(3)
# Navigate to the first page of job listings
driver.get("https://www.naukri.com/python-developer-jobs?k=python%20developer&nignbevent_src=jobsearchDeskGNB&experience=0&cityTypeGid=17&ugTypeGid=12&ugTypeGid=9502")
time.sleep(5)

# Initialize Excel workbook
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Job Listings"
ws.append(["Job Title", "Link", "Experience Required"])  # Removed "Status"

# Function to detect if the next button is disabled
def is_next_button_disabled():
    try:
        next_button = driver.find_element(By.XPATH, "//a[contains(@class, 'styles_btn-secondary__2AsIP') and contains(., 'Next')]")
        return "disabled" in next_button.get_attribute("class")
    except:
        return False

# Loop through pages by pressing "Next"
while True:
    # Close any overlay/pop-up if it appears
    try:
        popup = WebDriverWait(driver, 3).until(
            EC.element_to_be_clickable((By.CLASS_NAME, "styles_ppContainer__eeZyG"))
        )
        popup_close_button = popup.find_element(By.TAG_NAME, "button")  # Adjust this if there's a close button within the popup
        popup_close_button.click()
        print("Closed popup/overlay.")
        time.sleep(2)
    except Exception:
        # If no popup is found, continue normally
        pass

    # Find all job cards on the current page
    job_cards = driver.find_elements(By.CLASS_NAME, "srp-jobtuple-wrapper")
    
    if not job_cards:
        print("No more job cards found, stopping pagination.")
        break

    # Scrape job details from each job card
    for job in job_cards:
        try:
            # Extract title, link, and experience
            title_element = job.find_element(By.CLASS_NAME, "title")
            title = title_element.text
            link = title_element.get_attribute("href")
            
            experience_element = job.find_element(By.CLASS_NAME, "expwdth")
            experience = experience_element.text
            
            # Append data to Excel (only first 3 columns)
            ws.append([title, link, experience])  # Removed the status column
            print(f"Scraped: {title}")
        except Exception as e:
            print("Error scraping job:", e)
            continue

    # Save data periodically
    wb.save("Naukri_Job_Listings.xlsx")

    # Check if "Next" button is disabled (last page)
    if is_next_button_disabled():
        print("Reached the last page.")
        break
    
    # Try to find and click the "Next" button
    try:
        next_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.XPATH, "//a[contains(@class, 'styles_btn-secondary__2AsIP') and contains(., 'Next')]"))
        )
        driver.execute_script("arguments[0].click();", next_button)  # Use JavaScript click to avoid interception
        time.sleep(5)  # Wait for the next page to load
    except Exception as e:
        print("Next button not found or last page reached:", e)
        break

# Final save
wb.save("Naukri_Job_Listings.xlsx")
print("Job data saved to 'Naukri_Job_Listings.xlsx'.")

# Close the driver
driver.quit()