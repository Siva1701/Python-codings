import os
import time
import logging
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# ✅ **Configure Logging**
log_file = "automation_log.txt"
logging.basicConfig(
    filename=log_file,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S"
)

# ✅ **Set up Selenium WebDriver**
options = webdriver.ChromeOptions()
download_folder = r"C:\Users\sivagmm\Downloads\T2 Automation\PC_Timeout\Downloads"
prefs = {"download.default_directory": download_folder}
options.add_experimental_option("prefs", prefs)
driver = webdriver.Chrome(options=options)

def extract_task_ids(csv_file):
    """Extract task IDs and return them as a comma-separated string."""
    try:
        data = pd.read_csv(csv_file)
        task_ids = data['task_id'].dropna().astype(str).tolist()
        task_ids_str = ",".join(task_ids)
        logging.info(f"Extracted Task IDs: {task_ids_str}")
        return task_ids_str
    except Exception as e:
        logging.error(f"Error reading the CSV file: {e}")
        return ""

def reflow_task_ids(task_ids_str):
    """Enter all task IDs at once in a comma-separated format."""
    try:
        driver.get("https://canvas.verizon.com/uuigui/uui/espresso/static/uui3_infra/dist/index.html?_loc=/dashboard")
        wait = WebDriverWait(driver, 10)

        # ✅ **Navigate to Mass Task Processing**
        pc_workflow = wait.until(EC.presence_of_element_located((By.XPATH, "//div[@title='PC Workflow Controller Tools']")))
        pc_workflow.click()
        logging.info("PC Workflow clicked")

        time.sleep(10)
        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[@id='2']")))
        driver.switch_to.frame("frame1")
        logging.info("Switched to Work Management frame")

        work_management = wait.until(EC.presence_of_element_located((By.XPATH, "//*[@id='upiivoipmenu_work']")))
        work_management.click()

        mass_task_processing = wait.until(EC.presence_of_element_located((By.XPATH, "//div[1]/div/table/tbody/tr/td/div[1]/a[2]")))
        mass_task_processing.click()

        driver.switch_to.default_content()
        wait.until(EC.frame_to_be_available_and_switch_to_it((By.XPATH, "//iframe[@id='iframe0_1']")))

        # ✅ **Enter all Task IDs**
        input_box = wait.until(EC.presence_of_element_located((By.XPATH, "//textarea[@name='taskId']")))
        input_box.clear()
        input_box.send_keys(task_ids_str)  
        logging.info(f"Entered Task IDs: {task_ids_str}")

        # ✅ **Click 'Reflow' Button**
        reflow_button = wait.until(EC.presence_of_element_located((By.XPATH, "//button[text()='Reflow']")))
        reflow_button.click()

        # ✅ **Log Reflow Status**
        status = wait.until(EC.presence_of_element_located((By.XPATH, "//span[@id='statusField']")))
        logging.info(f"Reflow Status: {status.text}")

    except Exception as e:
        logging.error(f"Error during reflow process: {e}")

# ✅ **Main Execution Flow**
try:
    downloaded_file = "task_data.csv"  # Example CSV file
    task_ids_str = extract_task_ids(downloaded_file)

    if task_ids_str:
        reflow_task_ids(task_ids_str)

    logging.info("Automation completed successfully.")
finally:
    driver.quit()
    logging.info("Browser closed.")
