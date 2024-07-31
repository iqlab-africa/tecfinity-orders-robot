from robocorp.windows import desktop
from robocorp.tasks import task
from RPA.JSON import JSON
from RPA.Excel.Files import Files
from RPA.Desktop import Desktop
import pyautogui
import pyperclip
import time
import re
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

def clear_clipboard():
    """Clear the clipboard."""
    try:
        logger.info("Clearing the clipboard.")
        pyperclip.copy('')
    except Exception as e:
        logger.error(f"An error occurred while clearing the clipboard: {e}")

def highlight_and_copy(start_x, start_y, end_x, end_y):
    """Highlight text on the screen and copy it to the clipboard."""
    try:
        # Clear the clipboard
        clear_clipboard()

        logger.info(f"Moving to start position ({start_x}, {start_y}).")
        pyautogui.moveTo(start_x, start_y)
        
        logger.info("Clicking and holding to start selection.")
        pyautogui.mouseDown(button='left')
        
        logger.info(f"Moving to end position ({end_x}, {end_y}) to select text.")
        pyautogui.moveTo(end_x, end_y, duration=0.5)
        
        logger.info("Releasing mouse button to end selection.")
        pyautogui.mouseUp(button='left')
        
        logger.info("Copying selection to clipboard.")
        pyautogui.hotkey('ctrl', 'c')
        
        logger.info("Waiting for the clipboard to update.")
        time.sleep(0.5)
        logger.info("Retrieving text from clipboard.")
        text = pyperclip.paste()

        logger.info("Clipboard content saved to clipboard_content.txt")

        clipboard_text = desktop_lib.get_clipboard_value()
        logger.info(f"Captured text from clipboard: {clipboard_text}")

        return clipboard_text
    except Exception as e:
        logger.error(f"An error occurred while capturing screen text: {e}")
        return None
    

def save_ocr_output(customer_number, clipboard_text):
    """Save OCR text output."""
    try:
        screentext_path = os.path.join(SCREENSHOT_DIR, f"Cust_{customer_number}_output_text.txt")
        
        logger.info(f"Saving clipboard text to {screentext_path}.")
        with open(screentext_path, 'w', encoding='utf-8') as f:
            f.write(clipboard_text)
        
        logger.info(f"OCR output saved for customer number {customer_number} at {screentext_path}")
    
    except Exception as e:
        logger.error(f"An error occurred while saving OCR output for customer number {customer_number}: {e}")

def extract_pnumber_from_text(screen_text):
    """Get output text and use regex to extract pnumber."""
    try:
        # Regex find the P number by the pretext
        match = re.search(r"CONF. NO \w+", screen_text)
        if match:
            extracted_p_number = match.group()
            # Clean the matched text
            extracted_p_number = extracted_p_number.replace("CONF. NO ", "")
            logger.info(f"Extracted parcel number: {extracted_p_number}")
            return extracted_p_number
        else:
            logger.warning("Parcel number not found in the text.")
            return None
    
    except Exception as e:
        logger.error(f"An error occurred while extracting parcel number: {e}")
        return None
    
def capture_new_order(customer_number, orderdesc, stock_no, quantity_value, comment):
    """Capture a new order in the system."""
    try:
        logger.info(f"Processing customer number: {customer_number}")

        desktop().send_keys("+{Enter}")
        logger.info("Sent Shift + Enter.")
        time.sleep(3)
        press_enter(2)
        time.sleep(3)
        enter_value(customer_number)
        time.sleep(4)
        press_enter(1)
        send_keys_multiple_times("{Esc}", 1)
        press_enter(2)
        enter_value(orderdesc)
        time.sleep(5)
        press_enter(4)
        enter_value(stock_no)
        press_enter(3)
        enter_value(quantity_value)
        press_enter(1)
        enter_value("C1")
        enter_value(comment)
        press_enter(4)
        logger.info("New order captured successfully.")
    
    except Exception as e:
        logger.error(f"An error occurred while capturing new order for customer number {customer_number}: {e}")

            
    
def process_customers(customer_data):
    work_items = []
    for row in customer_data:
        try:
            customer_number = row['Account No']
            stock_no = row['Stock No']
            quantity_value = row['Quantity']
            allocated_user = row['Allocated User']
            no_of_labels = row['No of Labels']
            total_weight = row['Total Weight']
            orderdesc = row['Order Description']
            comment = row['Comment']
            packer = row['Packer']
            checker = row['Checker']

            capture_new_order(customer_number, orderdesc, stock_no, quantity_value, comment)

            start_pos = (390, 218)
            end_pos = (1168, 542)
            clipboard_text = highlight_and_copy(start_pos[0], start_pos[1], end_pos[0], end_pos[1])
            save_ocr_output(customer_number, clipboard_text)
            pnumber = extract_pnumber_from_text(clipboard_text)
            had_error = 'no' if pnumber else 'yes'

            work_item = {
                "pnumber": pnumber,
                "allocated_user": allocated_user,
                "stock_no": stock_no,
                "quantity_value": quantity_value,
                "no_of_labels": no_of_labels,
                "total_weight": total_weight,
                "packer": packer,
                "checker": checker,
                "had_error": had_error
            }
            work_items.append(work_item)

            time.sleep(4)
        except Exception as e:
            logger.error(f"An error occurred while processing customer number {customer_number}: {e}")
            time.sleep(4)
            raise

    json_lib.save_to_file(WORK_ITEMS_FILE_PATH, work_items)


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