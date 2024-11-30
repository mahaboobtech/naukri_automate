import openpyxl
import os
import logging
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys

# Define logging
# Define logging to append to the log file
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    filename="application_log.log",  # Log file name
    filemode="a"  # Append mode
)

# Define the file to store last applied job index
last_applied_file = "last_applied_job.txt"

def log_status_and_index(index, status, title):
    """Log the current application status and index."""
    logging.info(f"Index: {index}, Job Title: '{title}', Status: {status}")

# Function to read the last applied job index
def get_last_applied_index():
    if os.path.exists(last_applied_file):
        with open(last_applied_file, "r") as file:
            return int(file.read().strip())
    return 0

# Function to save the last applied job index
def save_last_applied_index(index):
    with open(last_applied_file, "w") as file:
        file.write(str(index))

# Load the prepared job links from the Excel file
input_wb = openpyxl.load_workbook("Naukri_Job_Listings.xlsx")
input_ws = input_wb.active

# Initialize WebDriver
driver = webdriver.Chrome()

# Login credentials
email = "EMAIL"
password = "PASSWORD"

# File paths
final_file = "final_ouptut_x.xlsx"
applied_file = "Successfully_Applied_Jobs.xlsx"
failed_file = "Failed_Jobs.xlsx"

# Load or create final_output.xlsx
if os.path.exists(final_file):
    final_wb = openpyxl.load_workbook(final_file)
    final_ws = final_wb.active
else:
    final_wb = openpyxl.Workbook()
    final_ws = final_wb.active
    final_ws.append(["Job Title", "Link", "Experience", "Status", "External Link or Notice"])

# Login to Naukri
try:
    driver.get("https://www.naukri.com/nlogin/login")
    time.sleep(5)
    email_input = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "usernameField"))
    )
    password_input = driver.find_element(By.ID, "passwordField")
    email_input.send_keys(email)
    password_input.send_keys(password)
    password_input.send_keys(Keys.RETURN)
    logging.info("Logged in successfully.")
    time.sleep(5)
except Exception as e:
    logging.error(f"Login failed: {e}")
    driver.quit()

# Get the last applied job index
last_applied_index = get_last_applied_index()

try:
    # Loop through each job link starting from the last applied job
    for index, row in enumerate(input_ws.iter_rows(min_row=2, values_only=True), start=1):
        if index <= last_applied_index:
            continue

        title, link, experience = row

        # Open the job link
        try:
            driver.get(link)
            time.sleep(3)
        except Exception as e:
            logging.error(f"Failed to open job link '{title}': {e}")
            final_ws.append([title, link, experience, "Failed", "Link not accessible"])
            continue

        try:
            # Check for already applied status
            already_applied = driver.find_elements(By.ID, "already-applied")
            if already_applied:
                status = "Already Applied"
                logging.info(f"Job '{title}' is already applied.")
                final_ws.append([title, link, experience, status, ""])
                save_last_applied_index(index)
                log_status_and_index(index, status, title)
                continue

            # Apply for the job
            apply_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "apply-button"))
            )
            apply_button.click()
            time.sleep(5)  # Delay after pressing the "Apply" button to wait for any banners or popups

            # Check for error banner
            try:
                error_banner = WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CLASS_NAME, "styles_user-msg__YLRsE"))
                )
                if error_banner:
                    status = "Server Error"
                    logging.info(f"Server error detected for job '{title}'.")
                    final_ws.append([title, link, experience, status, ""])
                    save_last_applied_index(index)
                    log_status_and_index(index, status, title)
                    continue
            except Exception:
                # If no error banner, proceed to check for chatbot popup
                try:
                    chatbot_popup = WebDriverWait(driver, 5).until(
                        EC.presence_of_element_located((By.XPATH, "//div[contains(@class, 'chatbot_Drawer')]"))
                    )
                    status = "Chatbot Notice"
                    logging.info(f"Chatbot detected for job '{title}', marking as {status}.")
                    final_ws.append([title, link, experience, status, ""])
                    save_last_applied_index(index)
                    log_status_and_index(index, status, title)
                    continue  # Skip to the next job if chatbot is detected
                except Exception:
                    # If neither the banner nor chatbot popup is detected, assume successful application
                    status = "Applied"
                    logging.info(f"Successfully applied for job: {title}")
                    final_ws.append([title, link, experience, status, ""])
                    save_last_applied_index(index)
                    log_status_and_index(index, status, title)

        except Exception as e:
            status = "Failed"
            logging.warning(f"Failed to apply for job '{title}': {e}")
            log_status_and_index(index, status, title)
            try:
                # Check for "Apply on company site" button
                company_site_button = driver.find_element(By.ID, "company-site-button")
                company_site_button.click()
                time.sleep(2)

                # Switch to the new tab
                driver.switch_to.window(driver.window_handles[1])
                external_url = driver.current_url
                status = "Redirected to Company Site"
                logging.info(f"Redirected to company site for job '{title}': {external_url}")
                final_ws.append([title, link, experience, status, external_url])
                save_last_applied_index(index)
                log_status_and_index(index, status, title)

                # Close the new tab and switch back to the main tab
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
            except Exception as sub_e:
                status = "Failed"
                logging.error(f"Failed to retrieve company site link for '{title}': {sub_e}")
                final_ws.append([title, link, experience, status, "No link found"])
                save_last_applied_index(index)
                log_status_and_index(index, status, title)

        except Exception as e:
            logging.warning(f"Failed to apply for job '{title}': {e}")
            try:
                # Check for "Apply on company site" button
                company_site_button = driver.find_element(By.ID, "company-site-button")
                company_site_button.click()
                time.sleep(2)

                # Switch to the new tab
                driver.switch_to.window(driver.window_handles[1])
                external_url = driver.current_url
                logging.info(f"Redirected to company site for job '{title}': {external_url}")

                # Log external site link
                final_ws.append([title, link, experience, "Redirected to Company Site", external_url])

                # Close the new tab and switch back to the main tab
                driver.close()
                driver.switch_to.window(driver.window_handles[0])
            except Exception as sub_e:
                logging.error(f"Failed to retrieve company site link for '{title}': {sub_e}")
                final_ws.append([title, link, experience, "Failed", "No link found"])

finally:
    # Save the final Excel file
    final_wb.save(final_file)
    logging.info(f"Saved data to {final_file}")

    # Quit the WebDriver
    driver.quit()
