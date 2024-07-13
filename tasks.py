from robocorp.windows import desktop, find_window
from robocorp.tasks import task
from RPA.JSON import JSON
from RPA.Excel.Files import Files
import time

# Initialize the JSON and Excel libraries
json_lib = JSON()
excel_lib = Files()

# Paths to your files
CREDENTIALS_JSON_FILE_PATH = "devdata/creds/mainframe_credentials.json"
MAINFRAME_CLIENT_PATH = r"C:\\Users\\27810\\OneDrive\\Documentos\\Dynamic Connect\\Session\\TECFINITY.dcs"
ORDERS_INPUT_FILE_PATH = "devdata/input/OWN FLEET TESTING CUSTOMER LIST 2024-07.xlsx"
def load_credentials():
    """Load credentials from the JSON file."""
    print("Loading credentials from JSON file...")
    credentials_list = json_lib.load_json_from_file(CREDENTIALS_JSON_FILE_PATH)
    
    if not credentials_list:
        print("Credentials list is empty.")
        return None
    
    credentials_payload = credentials_list[0]['payload'] if isinstance(credentials_list, list) else credentials_list['payload']
    print("Credentials loaded successfully.")
    return credentials_payload["username"], credentials_payload["password"]

def start_mainframe_client():
    """Start the mainframe client."""
    print("Opening mainframe client...")
    desktop().windows_run(MAINFRAME_CLIENT_PATH)
    print("Waiting for the mainframe client to load...")
    time.sleep(10)  # Adjust this time based on your application load time

def login(username, password):
    """Perform the login with the provided credentials."""
    print("Sending login credentials...")
    desktop().send_keys("{Enter}")  # Send the Enter key to start the login process
    time.sleep(3)  # Wait a moment before sending the next inputs
    
    print(f"Entering username: {username}")
    desktop().send_keys(f"{username}{{Enter}}")
    time.sleep(1)  # Wait a moment before sending the password
    
    print("Entering password.")
    desktop().send_keys(f"{password}{{Enter}}")
    time.sleep(1)  # Wait a moment for the login process to complete
    print("Login process completed.")

def rollback_to_main_screen():
    """Rollback to the main screen by sending F1 key 4 times."""
    try:
        print("Rolling back to the main screen by sending F1 key 4 times...")
        for _ in range(4):
            desktop().send_keys('{F1}')
            time.sleep(1)  # Adjust the sleep time if necessary
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

def close_mainframe_client():
    """Close the mainframe client."""
    try:
        print("Attempting to close the mainframe client...")
        desktop().send_keys('{Alt}{F4}')
        desktop().send_keys('{Tab}')
        desktop().send_keys('{Enter}')
        print("Sent Alt + F4 to the mainframe window.")
    except Exception as e:
        print(f"An error occurred while trying to close the mainframe client: {e}")

def load_customer_numbers():
    """Load customer numbers from the Excel file."""
    print("Loading customer numbers from Excel file...")
    excel_lib.open_workbook(ORDERS_INPUT_FILE_PATH)
    rows = excel_lib.read_worksheet_as_table(header=True)
    excel_lib.close_workbook()
    print("Customer numbers loaded successfully.")
    return [row['Account No'] for row in rows if 'Account No' in row]

def process_customers(customer_numbers):
    """Process each customer by performing the required key strokes."""
    for customer_number in customer_numbers:
        try:
            print(f"Processing customer number: {customer_number}")
            desktop().send_keys("+{Enter}")  # Send Shift + , then Enter
            desktop().send_keys("{Enter}") 
            desktop().send_keys("{Enter}")
            desktop().send_keys(f"{customer_number}")  # Type or paste the customer number
            time.sleep(10)  # Adjust the sleep time if necessary
            desktop().send_keys("{Enter}")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            desktop().send_keys("{Esc}")
            time.sleep(5)
            desktop().send_keys("{Esc}")
            time.sleep(5)
            # enter order tag
            desktop().send_keys("TEST")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            # ENTER ORDER INFO
            desktop().send_keys("BRMD")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            # END ENTER ORDER INFO
            desktop().send_keys("C1")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            desktop().send_keys("TEST ORDER")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            desktop().send_keys("{Enter}")
            time.sleep(5)
            rollback_to_main_screen()
            close_mainframe_client()

        except Exception as e:
            print(f"An error occurred while processing customer number {customer_number}: {e}")
            rollback_to_main_screen()
            close_mainframe_client()
            raise

@task
def open_mainframe_client():
    """Main task to open and interact with the mainframe client."""
    try:
        username, password = load_credentials()
        if not username or not password:
            return
        
        start_mainframe_client()
        login(username, password)
        
        # Add your additional steps here
        print("Waiting for the inner frame client to load...")
        time.sleep(3)  # Adjust this time based on your application load time
        
        login(username, password)
        
        print("Sending login credentials again...")
        desktop().send_keys("{Enter}")  # Send the Enter key to start the login process
        time.sleep(10)  # Wait a moment before sending the next inputs
        
        customer_numbers = load_customer_numbers()
        process_customers(customer_numbers)
        rollback_from_sub_screen()
        close_mainframe_client()

    except Exception as e:
        print(f"An error occurred: {e}")
        rollback_to_main_screen()
        close_mainframe_client()

# Example usage
open_mainframe_client()
