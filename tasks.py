from robocorp.windows import desktop, find_window
from robocorp.tasks import task
from RPA.JSON import JSON
from RPA.Excel.Files import Files
from RPA.Desktop import Desktop
from RPA import OCR
import time
import re

# Initialize the libraries
json_lib = JSON()
excel_lib = Files()
desktop_lib = Desktop()
ocr_lib = OCR()

# Paths to your files
CREDENTIALS_JSON_FILE_PATH = "devdata/creds/mainframe_credentials.json"
MAINFRAME_CLIENT_PATH = r"C:\\Users\\27810\\OneDrive\\Documentos\\Dynamic Connect\\Session\\TECFINITY.dcs"
ORDERS_INPUT_FILE_PATH = "devdata/input/OWN FLEET TESTING CUSTOMER LIST 2024-07.xlsx"

def load_credentials():
    """Load credentials from the JSON file."""
    try:
        print("Loading credentials from JSON file...")
        credentials_list = json_lib.load_json_from_file(CREDENTIALS_JSON_FILE_PATH)
        
        if not credentials_list:
            print("Credentials list is empty.")
            return None
        
        credentials_payload = credentials_list[0]['payload'] if isinstance(credentials_list, list) else credentials_list['payload']
        print("Credentials loaded successfully.")
        return credentials_payload["username"], credentials_payload["password"]
    except Exception as e:
        print(f"Failed to load credentials: {e}")
        return None

def start_mainframe_client():
    """Start the mainframe client."""
    try:
        print("Opening mainframe client...")
        desktop().windows_run(MAINFRAME_CLIENT_PATH)
        print("Waiting for the mainframe client to load...")
        time.sleep(10)  # Adjust this time based on your application load time
    except Exception as e:
        print(f"Failed to start mainframe client: {e}")

def login(username, password):
    """Perform the login with the provided credentials."""
    try:
        print("Sending login credentials...")
        press_enter(1)  # Send the Enter key to start the login process
        
        print(f"Entering username: {username}")
        enter_value(username)
        
        print("Entering password.")
        enter_value(password)
        print("Login process completed.")
    except Exception as e:
        print(f"Failed to login: {e}")

def rollback_to_main_screen():
    """Rollback to the main screen by sending F1 key 4 times."""
    try:
        print("Rolling back to the main screen by sending F1 key 4 times...")
        send_keys_multiple_times('{F1}', 4)
        print("Rollback to main screen completed.")
    except Exception as e:
        print(f"An error occurred during rollback: {e}")

def rollback_from_sub_screen():
    """Rollback from a sub-screen to the main screen."""
    try:
        print("Rolling back to the main screen by exiting sub screen times...")
        desktop().send_keys('{F1}')
        desktop().send_keys('{RIGHT}')
        desktop().send_keys('{Enter}')
        time.sleep(1)  # Adjust the sleep time if necessary
        print("Rollback to main screen completed.")
    except Exception as e:
        print(f"An error occurred during rollback: {e}")

def press_enter(times=1):
    """Press the Enter key a specified number of times."""
    try:
        for _ in range(times):
            desktop().send_keys('{Enter}')
            time.sleep(3)  # Adjust the sleep time if necessary
    except Exception as e:
        print(f"Failed to press Enter: {e}")

def enter_value(param, enter_after=True):
    """Enter a value and optionally press Enter."""
    try:
        desktop().send_keys(f"{param}")
        time.sleep(1)  # Adjust the sleep time if necessary
        if enter_after:
            press_enter(1)
    except Exception as e:
        print(f"Failed to enter value: {e}")

def close_mainframe_client():
    """Close the mainframe client."""
    try:
        print("Attempting to close the mainframe client...")
        desktop().send_keys('{Alt}{F4}')
        desktop().send_keys('{Tab}')
        press_enter()
        print("Sent Alt + F4 to the mainframe window.")
    except Exception as e:
        print(f"An error occurred while trying to close the mainframe client: {e}")

def load_customer_data():
    """Load customer data from the Excel file."""
    try:
        print("Loading customer data from Excel file...")
        excel_lib.open_workbook(ORDERS_INPUT_FILE_PATH)
        rows = excel_lib.read_worksheet_as_table(header=True)
        excel_lib.close_workbook()
        print("Customer data loaded successfully.")
        return rows
    except Exception as e:
        print(f"Failed to load customer data: {e}")
        return []

def send_keys_multiple_times(key, times):
    """Send a specified key multiple times."""
    try:
        for _ in range(times):
            desktop().send_keys(key)
            time.sleep(3)  # Adjust the sleep time if necessary
    except Exception as e:
        print(f"Failed to send keys multiple times: {e}")

def process_customers(customer_data):
    """Process each customer and enter order details."""
    for row in customer_data:
        try:
            customer_number = row['Account No']
            stock_no = row['Stock No']
            quantity_value = row['Quantity']
            allocated_user = row['Allocated User']
            no_of_labels = row['No of Labels']
            total_weight = row['Total Weight']
            orderdesc = row['Order Description']
            comment =row['Comment']
            packer = row['Packer']
            checker = row['Checker']
            
            print(f"Processing customer number: {customer_number}")
            desktop().send_keys("+{Enter}")  # Send Shift + Enter
            press_enter(2)
            enter_value(customer_number)
            time.sleep(4)  # Adjust the sleep time if necessary
            press_enter(2)
            send_keys_multiple_times("{Esc}", 2)
            # Enter order tag
            enter_value(orderdesc)
            time.sleep(5)
            press_enter(5)
            # Enter order info
            enter_value(stock_no)
            press_enter(2)
            # Enter quantity value
            enter_value(quantity_value)
            press_enter(4)
            # End enter order info
            enter_value("C1")
            enter_value(comment)
            press_enter(3)
            rollback_to_main_screen()
            # Process the rest using the extracted pnumber
            pnumber = extract_pnumber_from_screenshot(capture_screenshot())
            if pnumber:
                release_onhold_order(pnumber)
                allocate_picking_slip(pnumber, allocated_user)
                precheck_picking_slip(pnumber, stock_no, quantity_value)
                scan_picking_slip(pnumber, stock_no, quantity_value)
                print_delivery_slip(pnumber, no_of_labels, total_weight, packer, checker)
            close_mainframe_client()
        except Exception as e:
            print(f"An error occurred while processing customer number {customer_number}: {e}")
            rollback_to_main_screen()
            close_mainframe_client()
            raise

def press_arrow_down(times=1):
    """Press the arrow down key a specified number of times."""
    try:
        for _ in range(times):
            desktop().send_keys('{DOWN}')
            time.sleep(1)  # Adjust the sleep time if necessary
    except Exception as e:
        print(f"Failed to press arrow down: {e}")

def press_arrow_right(times=1):
    """Press the arrow right key a specified number of times."""
    try:
        for _ in range(times):
            desktop().send_keys('{RIGHT}')
            time.sleep(1)  # Adjust the sleep time if necessary
    except Exception as e:
        print(f"Failed to press arrow right: {e}")

def release_onhold_order(pnumber):
    """Release on hold order using parcel number."""
    try:
        print(f"Processing customer p_number: {pnumber}")
        desktop().send_keys("+{Enter}")  # Send Shift + Enter
        enter_value(2)
        enter_value(pnumber)
        press_enter(1)
        enter_value("N")
        press_enter(1)
        enter_value("RL")
        press_enter(1)
        press_arrow_down(1)
        press_enter(1)
        enter_value(1)
        press_enter(1)
        press_arrow_right(1)   
        press_enter(2)
        time.sleep(4)
    except Exception as e:
        print(f"An error occurred while processing customer p_number {pnumber}: {e}")
        rollback_to_main_screen()
        close_mainframe_client()
        raise

def allocate_picking_slip(pnumber, allocated_user):
    """Allocate picking slip using parcel number and allocated user."""
    try:
        print(f"Processing customer p_number: {pnumber}")
        desktop().send_keys("+{Enter}")  # Send Shift + Enter
        enter_value(3)
        enter_value(allocated_user)
        press_enter(1)
        enter_value(pnumber)
        press_enter(1)
        time.sleep(3)
    except Exception as e:
        print(f"An error occurred while processing customer p_number {pnumber}: {e}")
        rollback_to_main_screen()
        close_mainframe_client()
        raise

def precheck_picking_slip(pnumber, stock_no, quantity_value):
    """Precheck picking slip using parcel number, stock number, and quantity value."""
    try:
        print(f"Processing customer p_number: {pnumber}")
        desktop().send_keys("+{Enter}")  # Send Shift + Enter
        enter_value(4)
        enter_value(pnumber)
        press_enter(2)
        enter_value(stock_no)
        press_enter(2)
        enter_value(quantity_value)
        press_enter(2)
        enter_value("C1")
        press_enter(2)
        time.sleep(4)
    except Exception as e:
        print(f"An error occurred while processing customer p_number {pnumber}: {e}")
        rollback_to_main_screen()
        close_mainframe_client()
        raise

def scan_picking_slip(pnumber, stock_no, quantity_value):
    """Scan picking slip using parcel number, stock number, and quantity value."""
    try:
        print(f"Processing customer p_number: {pnumber}")
        desktop().send_keys("+{Enter}")  # Send Shift + Enter
        enter_value(5)
        enter_value(pnumber)
        press_enter(2)
        enter_value(stock_no)
        press_enter(2)
        enter_value(quantity_value)
        press_enter(4)
        time.sleep(4)
    except Exception as e:
        print(f"An error occurred while processing customer p_number {pnumber}: {e}")
        rollback_to_main_screen()
        close_mainframe_client()
        raise

def print_delivery_slip(pnumber, no_of_labels, total_weight, packer, checker):
    """Print delivery slip using parcel number, number of labels, total weight, packer, and checker."""
    try:
        print(f"Processing customer p_number: {pnumber}")
        desktop().send_keys("+{Enter}")  # Send Shift + Enter
        enter_value(6)
        enter_value(pnumber)
        press_enter(1)
        enter_value(no_of_labels)
        press_enter(1)
        enter_value(total_weight)
        press_enter(1)
        enter_value(packer)
        press_enter(1)
        enter_value(checker)
        press_enter(1)
        press_arrow_right(1)
        press_enter(1)
        time.sleep(4)
    except Exception as e:
        print(f"An error occurred while processing customer p_number {pnumber}: {e}")
        rollback_to_main_screen()
        close_mainframe_client()
        raise
def capture_screenshot():
    """Capture a screenshot."""
    screenshot_path = "screenshot.png"
    try:
        desktop.screenshot(path=screenshot_path)
        return screenshot_path
    except Exception as e:
        print(f"Failed to capture screenshot: {e}")
        return None

def extract_pnumber_from_screenshot(screenshot_path):
    """Extract the parcel number from a screenshot using OCR."""
    try:
        if screenshot_path:
            print("Extracting parcel number from screenshot using OCR...")
            ocr_text = ocr_lib.image_to_string(screenshot_path)
            print(f"OCR extracted text: {ocr_text}")
            pnumber = extract_pnumber_from_text(ocr_text)
            print(f"Extracted parcel number: {pnumber}")
            return pnumber
        else:
            print("No screenshot path provided.")
            return None
    except Exception as e:
        print(f"Failed to extract parcel number from screenshot: {e}")
        return None

def extract_pnumber_from_text(ocr_text):
    """Extract the parcel number from the OCR text."""
    try:
        pattern = r'\b\d{6}\b'  # Adjust the pattern based on the actual format
        match = re.search(pattern, ocr_text)
        if match:
            return match.group(0)
        else:
            return None
    except Exception as e:
        print(f"Failed to extract parcel number from text: {e}")
        return None

@task
def process_orders():
    credentials = load_credentials()
    if credentials:
        username, password = credentials
        start_mainframe_client()
        login(username, password)
        
        customer_data = load_customer_data()
        if customer_data:
            process_customers(customer_data)
            
        close_mainframe_client()

if __name__ == "__main__":
    process_orders()