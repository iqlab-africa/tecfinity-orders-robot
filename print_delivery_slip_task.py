from robocorp.windows import desktop
from robocorp.tasks import task
from RPA.JSON import JSON
from RPA.Excel.Files import Files
from RPA.Desktop import Desktop
import time
import logging
import os


# Initialize the libraries
json_lib = JSON()
excel_lib = Files()
desktop_lib = Desktop()

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Paths to your files
CREDENTIALS_JSON_FILE_PATH = "devdata/creds/mainframe_credentials.json"
MAINFRAME_CLIENT_PATH = r"C:\\Users\\27810\\OneDrive\\Documentos\\Dynamic Connect\\Session\\TECFINITY.dcs"
ORDERS_INPUT_FILE_PATH = "devdata/input/testdata.xlsx"
SCREENSHOT_DIR = "output/screenshots"
WORK_ITEMS_FILE_PATH = "output/workitems.json"

# Ensure directories exist
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

def maximize_window():
    try:
        # Send Alt + Space to open the window's system menu
        desktop().send_keys('{Alt} {Space}')
        time.sleep(1)  # Wait for the system menu to open
        
        # Send 'X' to select the maximize option
        desktop().send_keys('X')
        time.sleep(1)  # Wait for the window to maximize
        print("Window maximized successfully.")
    except Exception as e:
        print(f"Failed to maximize window: {e}")


def load_credentials():
    """Load credentials from the JSON file."""
    try:
        logger.info("Loading credentials from JSON file...")
        credentials_list = json_lib.load_json_from_file(CREDENTIALS_JSON_FILE_PATH)
        
        if not credentials_list:
            logger.info("Credentials list is empty.")
            return None
        
        credentials_payload = credentials_list[0]['payload'] if isinstance(credentials_list, list) else credentials_list['payload']
        logger.info("Credentials loaded successfully.")
        return credentials_payload["username"], credentials_payload["password"]
    except Exception as e:
        logger.error(f"Failed to load credentials: {e}")
        return None
    

def rollback_to_main_screen():
    """Rollback to the main screen by sending F1 key 4 times."""
    try:
        logger.info("Rolling back to the main screen by sending F1 key 4 times...")
        send_keys_multiple_times('{F1}', 4)
        logger.info("Rollback to main screen completed.")
    except Exception as e:
        logger.error(f"An error occurred during rollback: {e}")

def rollback_from_sub_screen():
    """Rollback from a sub-screen to the main screen."""
    try:
        logger.info("Rolling back to the main screen by exiting sub screen...")
        desktop().send_keys('{F1}')
        desktop().send_keys('{RIGHT}')
        desktop().send_keys('{Enter}')
        time.sleep(2)  # Adjust the sleep time if necessary
        logger.info("Rollback to main screen completed.")
    except Exception as e:
        logger.error(f"An error occurred during rollback: {e}")

def press_enter(times=1):
    """Press the Enter key a specified number of times."""
    try:
        for _ in range(times):
            desktop().send_keys('{Enter}')
            time.sleep(3)  # Adjust the sleep time if necessary
    except Exception as e:
        logger.error(f"Failed to press Enter: {e}")

def enter_value(param, enter_after=True):
    """Enter a value and optionally press Enter."""
    try:
        desktop().send_keys(f"{param}")
        time.sleep(1) 
        if enter_after:
            press_enter(1)
    except Exception as e:
        logger.error(f"Failed to enter value: {e}")

def close_mainframe_client():
    """Close the mainframe client."""
    try:
        logger.info("Attempting to close the mainframe client...")
        desktop().send_keys('{Alt}{F4}')
        desktop().send_keys('{Tab}')
        press_enter(1)
        logger.info("Sent Alt + F4 to the mainframe window.")
    except Exception as e:
        logger.error(f"An error occurred while trying to close the mainframe client: {e}")

def load_customer_data():
    """Load customer data from the Excel file."""
    try:
        logger.info("Loading customer data from Excel file...")
        excel_lib.open_workbook(ORDERS_INPUT_FILE_PATH)
        rows = excel_lib.read_worksheet_as_table(header=True)
        excel_lib.close_workbook()
        logger.info("Customer data loaded successfully.")
        return rows
    except Exception as e:
        logger.error(f"Failed to load customer data: {e}")
        return []

def send_keys_multiple_times(key, times):
    """Send a specified key multiple times."""
    try:
        for _ in range(times):
            desktop().send_keys(key)
            time.sleep(3)  # Adjust the sleep time if necessary
    except Exception as e:
        logger.error(f"Failed to send keys multiple times: {e}")



def print_delivery_slip(pnumber, no_of_labels, total_weight, packer, checker):
    """Print delivery slip using parcel number, number of labels, total weight, packer, and checker."""
    try:
        logger.info(f"Processing customer p_number: {pnumber}")
        desktop().send_keys("+{Enter}") 
        enter_value(6) 
        enter_value(pnumber)
        press_enter(1)
        enter_value(no_of_labels)
        enter_value(total_weight)
        enter_value(packer)
        enter_value(checker)
        enter_value(checker)
        press_enter(2)
    except Exception as e:
        logger.error(f"An error occurred while processing customer p_number {pnumber}: {e}")
        rollback_to_main_screen()
        close_mainframe_client()
        raise


def print_option():
    """Process each customer and enter order details."""
    try:
        working_items_path = "output/workitems.json"
        working_items = json_lib.load_json_from_file(working_items_path)

        for item in working_items:
            pnumber = item['pnumber']
            no_of_labels = item['no_of_labels']
            total_weight = item['total_weight']
            packer = item['packer']
            checker = item['checker']
            
            if pnumber:
                print_delivery_slip(pnumber, no_of_labels, total_weight, packer, checker)
            else:
                logger.error("Failed to extract pnumber from ocr output")
    except Exception as e:
        logger.error(f"An error occurred while processing working items: {e}")
        raise
