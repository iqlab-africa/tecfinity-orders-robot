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

def start_mainframe_client():
    """Start the mainframe client."""
    try:
        logger.info("Opening mainframe client...")
        desktop().windows_run(MAINFRAME_CLIENT_PATH)
        logger.info("Waiting for the mainframe client to load...")
        time.sleep(4) 
        maximize_window()
        time.sleep(6)  # Adjust this time based on your application load time
    except Exception as e:
        logger.error(f"Failed to start mainframe client: {e}")

def login(username, password):
    """Perform the login with the provided credentials."""
    try:
        logger.info("Sending login credentials...")
        press_enter(1)  # Send the Enter key to start the login process
        logger.info(f"Entering username: {username}")
        enter_value(username)
        logger.info("Entering password.")
        enter_value(password)
        logger.info("Login process completed.")
        logger.info("Sending login credentials...TO SUBSCREEN")
        press_enter(1)  # Send the Enter key to start the login process
        logger.info(f"Entering username: {username}")
        enter_value(username)
        logger.info("Entering password.")
        enter_value(password)
        logger.info("Login process completed.")
    except Exception as e:
        logger.error(f"Failed to login: {e}")

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
        time.sleep(3)  # Adjust the sleep time if necessary
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

def press_arrow_down(times=1):
    """Press the arrow down key a specified number of times."""
    try:
        for _ in range(times):
            desktop().send_keys('{DOWN}')
            time.sleep(2)  # Adjust the sleep time if necessary
    except Exception as e:
        logger.error(f"Failed to press arrow down: {e}")

def press_arrow_right(times=1):
    """Press the arrow right key a specified number of times."""
    try:
        for _ in range(times):
            desktop().send_keys('{RIGHT}')
            time.sleep(2)  # Adjust the sleep time if necessary
    except Exception as e:
        logger.error(f"Failed to press arrow right: {e}")

def precheck_picking_slip(pnumber, stock_no, quantity_value):
    """4. Precheck picking slip using parcel number, stock number, and quantity value."""
    try:
        logger.info(f"Processing customer p_number: {pnumber}")
        desktop().send_keys("+{Enter}")  # Send Shift + Enter
        enter_value(4)
        enter_value(pnumber)
        press_enter(1)
        enter_value(stock_no)
        enter_value(quantity_value)
        press_enter(1)
        time.sleep(4)
    except Exception as e:
        logger.error(f"An error occurred while processing customer p_number {pnumber}: {e}")
        rollback_to_main_screen()
        close_mainframe_client()
        raise

def process_customers():
    """Process each customer and enter order details."""
    try:
        # Load working items from JSON
        working_items_path = "output/workitems.json"
        working_items = json_lib.load_json_from_file(working_items_path)

        for item in working_items:
            pnumber = item['pnumber']
            stock_no = item['stock_no']
            quantity_value = item['quantity_value']

            if pnumber:
               precheck_picking_slip(pnumber, stock_no, quantity_value)
            else:
                logger.error("Failed to extract pnumber from ocr output")

            time.sleep(4)
    except Exception as e:
        logger.error(f"An error occurred while processing working items: {e}")
        time.sleep(4)
        raise

@task
def main():
    """Main function to run the automation task."""
    credentials = load_credentials()
    if credentials:
        username, password = credentials
        start_mainframe_client()
        login(username, password)
        customer_data = load_customer_data()
        if customer_data:
            process_customers(customer_data)
        close_mainframe_client()
    else:
        print("No credentials available. Terminating the process.")

if __name__ == "__main__":
    main()