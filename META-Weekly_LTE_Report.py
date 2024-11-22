import os 
import subprocess
import time
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import fnmatch


# Path to the download directory (update as needed)
download_dir = "C:/Users/10015340/Downloads"

# Function to configure and create a new Chrome WebDriver instance
def create_driver():
    options = Options()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")  # Connect to the remote debugging port
    options.add_argument("--headless")  # Optionally run Chrome headlessly
    return webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

# Function to check if the file was downloaded successfully
def is_file_downloaded(report_name):
    # Check if the downloaded file exists in the download directory
    for filename in os.listdir(download_dir):
        if fnmatch.fnmatch(filename, f"*{report_name}*.csv"):  # File name contains the report_name and ends with .csv
            print(f"Found downloaded file: {filename}")
            return True
    return False

def open_new_tab_and_select_report(report_name, driver):
    """Helper function to open a new tab and perform report download steps."""
    driver.execute_script("window.open('');")  # Open a new tab
    driver.switch_to.window(driver.window_handles[-1])  # Switch to the newly opened tab
    driver.get("https://partners.metaenterprise.com/actionable_insights/global/dashboards/pinned?locale=en_US&partner_id=45985835156&dashboard_id=367012886275356&dashboard_partner_id=45985835156")  # Replace with actual URL
    time.sleep(3)  # Let the page load

    select_and_download_report(driver, report_name)

def select_and_download_report(driver, report_name, retries=3):
    """Function to perform the steps for selecting a report and downloading it."""
    wait = WebDriverWait(driver, 100)  # Increased wait time for elements to load
    
    attempt = 0
    while attempt < retries:
        try:
            # Step 1: Click on "View Report Templates"
            print(f"[Attempt {attempt + 1}] Searching for 'View Report Templates' button for {report_name}...")
            reports = wait.until(EC.element_to_be_clickable((By.XPATH, '//*[@id="XMPPInsights"]/div/div/div[1]/div[1]/div/div/div/div[2]/div[1]/div/nav[1]/div/div[2]/div[1]/div')))
            reports.click()
            print(f'[{report_name}] View Report Templates clicked.')

            # Step 2: Scroll and search for the report template dynamically
            print(f"[Attempt {attempt + 1}] Scrolling to find the report template: {report_name}...")
            grid_element = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, '.ReactVirtualized__Grid._1zmk')))
            found = False

            # Scroll within the grid element until the report name is found or max retries reached
            while not found:
                # Search for the report template in the current view
                try:
                    report_element = grid_element.find_element(By.XPATH, f".//div//span[contains(text(), '{report_name}')]")
                    report_element.click()  # Click the found report template
                    print(f'[{report_name}] Report template selected.')
                    found = True  # Stop scrolling if the element is found
                except:
                    # Scroll the grid down if the element is not found in the current view
                    driver.execute_script("arguments[0].scrollBy(0, 100);", grid_element)  # Scroll down 100px
                    time.sleep(1)  # Wait for the content to load
                    print(f"[Attempt {attempt + 1}] Scrolling down the grid...")

            # Step 3: Wait for the table to be visible
            print(f"[Attempt {attempt + 1}] Waiting for KPI table to be visible...")
            WebDriverWait(driver, 100).until(
                lambda driver: int(driver.find_element(By.XPATH, "//table[@aria-label='Untitled report. Table.']").get_attribute("aria-rowcount")) > 0
            )
            print(f'[{report_name}] KPI table is shown.')

            # Step 4: Click on the 'Actions' dropdown
            print(f"[Attempt {attempt + 1}] Searching for 'Actions' dropdown...")
            actions_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[@role='button']//span//div[contains(text(), 'Actions')]")))
            actions_button.click()
            print(f'[{report_name}] Preparing to export.')

            # Step 5: Click on the 'Export' option
            print(f"[Attempt {attempt + 1}] Searching for 'Export' option...")
            export_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//div[contains(text(), 'Export')]")))
            export_button.click()
            print(f'[{report_name}] Export initiated.')

            # Step 6: Click 'Download'
            print(f"[Attempt {attempt + 1}] Searching for 'Download' button...")
            download_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//*[@id='facebook']/body/div[6]/div[1]/div[1]/div/div/div/div/div[3]/div/div[2]/div")))
            download_button.click()
            print(f'[{report_name}] Download started.')

            # Check if the report is downloaded, retry if not
            attempts = 0
            while not is_file_downloaded(report_name) and attempts < 3:
                print(f"[{report_name}] Download not found. Retrying... (Attempt {attempts + 1})")
                time.sleep(5)  # Wait for a few seconds before retrying
                attempts += 1

            if attempts >= 3:
                print(f"[{report_name}] Failed to download report after 3 attempts.")
                return  # Return early if the download fails after retrying
            else:
                print(f"[{report_name}] Downloaded successfully.")
                return  # Successfully downloaded, exit function

        except Exception as e:
            print(f"Error while processing {report_name} (Attempt {attempt + 1}): {str(e)}")
            attempt += 1
            if attempt < retries:
                print(f"Retrying {report_name}... (Attempt {attempt + 1})")
            else:
                print(f"[{report_name}] Failed after {retries} attempts.")
                return  # If retry limit reached, exit the function


def kill_chrome_processes():
    """Ensure all Chrome processes are killed."""
    subprocess.run(["taskkill", "/f", "/im", "chrome.exe"], stdout=subprocess.PIPE, stderr=subprocess.PIPE)


# List of report names
report_names = [
    "meta_1week_region_android_lte_mobile_metrics_part_1", "meta_1week_region_android_lte_mobile_metrics_part_2", "meta_1week_region_android_lte_mobile_metrics_part_3", 
    "meta_1week_town_android_lte_mobile_metrics_part_1", "meta_1week_town_android_lte_mobile_metrics_part_2", "meta_1week_town_android_lte_mobile_metrics_part_3",
    "meta_1week_province_android_lte_mobile_metrics_part_1", "meta_1week_province_android_lte_mobile_metrics_part_2", "meta_1week_province_android_lte_mobile_metrics_part_3",
    "meta_1week_barangay_android_lte_mobile_metrics_part_1", "meta_1week_barangay_android_lte_mobile_metrics_part_2", "meta_1week_barangay_android_lte_mobile_metrics_part_3"
]

# Step 1: Open CMD and send the command to launch Chrome with remote debugging mode
subprocess.run(["start", "chrome", "--remote-debugging-port=9222"], shell=True)

# Wait a few seconds for Chrome to open and initialize remote debugging
time.sleep(3)

# Step 2: Create a single WebDriver instance
driver = create_driver()

try:
    # Process each report one at a time, in sequence
    for report_name in report_names:
        open_new_tab_and_select_report(report_name, driver)

finally:
    # Step 3: Closing the driver after all reports are processed
    driver.quit()
    print("Driver closed after all reports are processed.")
    # Explicitly kill any chrome processes to ensure the browser is closed
    kill_chrome_processes()