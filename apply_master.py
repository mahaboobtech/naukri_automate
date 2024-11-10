from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl
import logging

# Logging setup to save in a separate file called "application_log.log"
logging.basicConfig(
    filename="application_log.log",  # Log file name
    level=logging.INFO,  # Log level
    format="%(asctime)s - %(levelname)s - %(message)s",  # Log format
    filemode="w"  # Overwrite log file on each run
)

# Load the prepared job links from the Excel file
input_wb = openpyxl.load_workbook("Naukri_Job_Listings.xlsx")
input_ws = input_wb.active

# Create workbooks for successful and failed applications
applied_wb = openpyxl.Workbook()
failed_wb = openpyxl.Workbook()
applied_ws = applied_wb.active
failed_ws = failed_wb.active
applied_ws.append(["Job Title", "Link", "Experience"])
failed_ws.append(["Job Title", "Link", "Experience", "Company Site Link or Chatbot Notice"])

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
    
    logging.info("Logged in successfully.")
    time.sleep(5)  # Wait for the page to load after login
except Exception as e:
    logging.error(f"Login failed: {e}")
    driver.quit()

try:
    # Loop through each job link and apply
    for row in input_ws.iter_rows(min_row=2, values_only=True):  # Assuming the first row is headers
        title, link, experience = row

        # Open the job link
        try:
            driver.get(link)
            time.sleep(3)  # Wait for the job page to load
        except Exception as get_link_error:
            logging.error(f"Failed to open link for job '{title}': {get_link_error}")
            failed_ws.append([title, link, experience, "Link not accessible"])
            continue  # Skip to the next job link

        try:
            # Check if the "Server Error" message is present within 2 seconds after applying
            apply_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "apply-button"))
            )
            apply_button.click()
            logging.info(f"Attempted to apply for job: {title}")

            # Wait for 2 seconds to check for error message
            time.sleep(2)

            # Check if the error message exists after applying
            error_message_section = driver.find_elements(By.CSS_SELECTOR, "section.styles_user-msg__YLRsE span")
            if error_message_section:
                error_text = error_message_section[0].text
                if "There was an error while processing your request, please try again later" in error_text:
                    logging.error(f"Server error for job '{title}', marking as 'Server Error'.")
                    failed_ws.append([title, link, experience, "Server Error"])
                    continue  # Skip to the next job link

            # Check if the job has already been applied
            already_applied = driver.find_elements(By.ID, "already-applied")
            if already_applied:
                # If the "Already Applied" element is found, mark the status as "Already Applied"
                logging.info(f"Job '{title}' is already applied. Marking as 'Already Applied'.")
                applied_ws.append([title, link, experience])  # Still mark it as applied
                continue  # Skip to the next job link

            # Try to find and click the "Apply" button again if no error
            apply_button = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.ID, "apply-button"))
            )
            apply_button.click()

            # Wait for the application to process
            time.sleep(3)

        except Exception as e:
            logging.warning(f"Failed to apply for job '{title}': {e}")
            
            # Check if the "Apply on company site" button exists
            try:
                company_site_button = driver.find_element(By.ID, "company-site-button")
                company_site_button.click()
                time.sleep(2)  # Wait for the new tab to open
                
                # Switch to the new tab
                driver.switch_to.window(driver.window_handles[1])
                external_url = driver.current_url  # Get the external URL
                logging.info(f"Redirected to company site for job '{title}': {external_url}")
                
                # Log to failed Excel with external site link
                failed_ws.append([title, link, experience, external_url])
                
                # Close the new tab and switch back to the main tab
                driver.close()
                driver.switch_to.window(driver.window_handles[0])

            except Exception as sub_e:
                logging.error(f"Failed to retrieve company site link for '{title}': {sub_e}")
                # Log to failed Excel without external link if button is not found
                failed_ws.append([title, link, experience, "No link found"])
            continue

finally:
    # Save Excel files
    applied_wb.save("Successfully_Applied_Jobs.xlsx")
    failed_wb.save("Failed_Jobs.xlsx")
    logging.info("Application process completed for all jobs and files saved.")

    # Merge data into a final_output.xlsx file
    final_wb = openpyxl.Workbook()
    final_ws = final_wb.active
    final_ws.append(["Job Title", "Link", "Experience", "Status", "External Link or Notice"])

    # Add original job listings with "Status" and "External Link" columns
    for row in input_ws.iter_rows(min_row=2, values_only=True):
        title, link, experience = row
        applied_status = "Failed"
        external_link_or_notice = ""
        
        # Check if job is in applied list
        for applied_row in applied_ws.iter_rows(min_row=2, values_only=True):
            if title == applied_row[0] and link == applied_row[1]:
                applied_status = "Applied"
                break
        
        # If not applied, check if in failed list and get external link or notice if available
        if applied_status == "Failed":
            for failed_row in failed_ws.iter_rows(min_row=2, values_only=True):
                if title == failed_row[0] and link == failed_row[1]:
                    external_link_or_notice = failed_row[3]
                    break
        
        # If the job was already applied, mark it as "Already Applied"
        if applied_status == "Failed" and external_link_or_notice == "":
            for applied_row in applied_ws.iter_rows(min_row=2, values_only=True):
                if title == applied_row[0] and link == applied_row[1]:
                    applied_status = "Already Applied"
                    break
        
        final_ws.append([title, link, experience, applied_status, external_link_or_notice])

    final_wb.save("final_output.xlsx")
    logging.info("Merged data saved to final_output.xlsx")

    # Quit the driver
    driver.quit()