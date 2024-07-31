from robocorp.windows import desktop
from robocorp.tasks import task
from RPA.JSON import JSON
from RPA.Excel.Files import Files
from RPA.Desktop import Desktop
import logging
import os
import time

# Initialize the libraries
json_lib = JSON()
desktop_lib = Desktop()

# Set up logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Paths to your files
CREDENTIALS_JSON_FILE_PATH = "devdata/creds/mainframe_credentials.json"
MAINFRAME_CLIENT_PATH = r"C:\\Users\\27810\\OneDrive\\Documentos\\Dynamic Connect\\Session\\TECFINITY.dcs"
SCREENSHOT_DIR = "output/screenshots"

# Ensure directories exist
os.makedirs(SCREENSHOT_DIR, exist_ok=True)

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

@task
def main():
    """Main function to run the login task."""
    credentials = load_credentials()
    if credentials:
        username, password = credentials
        start_mainframe_client()
        login(username, password)
        close_mainframe_client()
    else:
        print("No credentials available. Terminating the process.")

if __name__ == "__main__":
    main()
