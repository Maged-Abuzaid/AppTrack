import os
import json
import shutil
from appdirs import user_data_dir
from src.utils.utils import resource_path

# Application-specific information
app_name = "AppTrack"

# Base path to store files in a user-specific local directory (AppData or equivalent)
base_path = user_data_dir(app_name)

# Paths for configuration, assets, and data directories
CONFIG_DIR = os.path.join(base_path, 'config')
ASSETS_DIR = os.path.join(base_path, 'assets')
DATA_DIR = os.path.join(base_path, 'Data')

# Paths for application files stored exclusively in AppData
CONFIG_JSON_PATH = os.path.join(CONFIG_DIR, 'app_config.json')
ICON_PATH = os.path.join(ASSETS_DIR, 'app_icon.png')
PERSONAL_INFO_FILE = os.path.join(DATA_DIR, 'personal_info.json')
DATA_FILE_PATH = os.path.join(DATA_DIR, 'Applications.xlsx')
SERVICE_ACCOUNT_FILE = os.path.join(CONFIG_DIR, 'service_account.json')

# Default configurations for settings
# config/settings_manager.py

default_config = {
    "theme": "Light",
    "clipboard_enabled": True,
    "clipboard_side": "right",
    "SERVICE_ACCOUNT_FILE": "/path/to/service_account.json",
    "SPREADSHEET_ID": "your_spreadsheet_id_here",
    "ENABLE_GOOGLE_SYNC": False,
    # ... other default settings
}


# Ensure required directories in AppData exist
os.makedirs(ASSETS_DIR, exist_ok=True)
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(CONFIG_DIR, exist_ok=True)

# Copy assets directory if it doesn't exist in AppData
project_assets_path = resource_path('assets')
if not os.listdir(ASSETS_DIR):
    if os.path.exists(project_assets_path):
        try:
            shutil.copytree(project_assets_path, ASSETS_DIR, dirs_exist_ok=True)
            print(f"Copied assets folder to {ASSETS_DIR}")
        except Exception as e:
            print(f"Error copying assets folder: {e}")
    else:
        print("Error: assets folder is missing from the project directory.")

# Load configurations from app_config.json if available, or initialize it in AppData
if os.path.exists(CONFIG_JSON_PATH):
    with open(CONFIG_JSON_PATH, "r") as config_file:
        try:
            user_config = json.load(config_file)
        except json.JSONDecodeError:
            user_config = default_config
            print("Warning: app_config.json is malformed. Using default configurations.")
else:
    # If app_config.json doesn't exist, create it with default settings in AppData
    user_config = default_config
    with open(CONFIG_JSON_PATH, "w") as config_file:
        json.dump(default_config, config_file, indent=4)
        print("Initialized app_config.json with default configurations in AppData.")

# Merge any missing default keys with loaded configurations to ensure all are present
for key, value in default_config.items():
    user_config.setdefault(key, value)

# Export configuration variables from loaded user_config
ENABLE_GOOGLE_SYNC = user_config["ENABLE_GOOGLE_SYNC"]
SPREADSHEET_ID = user_config["SPREADSHEET_ID"]
theme = user_config["theme"]

# Example range for Google Sheets
RANGE_NAME = "Sheet1!A1:E"  # Adjust this to the actual range as needed

def save_theme(new_theme):
    """Updates the theme in app_config.json located in AppData."""
    global theme  # Update the global theme variable to match the new setting
    theme = new_theme

    # Load the current configuration from CONFIG_JSON_PATH to ensure no other settings are overwritten
    try:
        with open(CONFIG_JSON_PATH, "r") as config_file:
            config_data = json.load(config_file)
    except (FileNotFoundError, json.JSONDecodeError):
        # If file does not exist or is corrupted, start with default configuration
        config_data = default_config

    # Update the theme in the loaded configuration data
    config_data["theme"] = theme

    # Save the updated configuration back to app_config.json
    try:
        with open(CONFIG_JSON_PATH, "w") as config_file:
            json.dump(config_data, config_file, indent=4)
        print(f"[DEBUG] Theme set to '{theme}' and saved in app_config.json in AppData.")
    except Exception as e:
        print(f"[ERROR] Failed to save theme to app_config.json: {e}")

# Debugging: Print paths to verify correct file locations are set to AppData
print(f"App Local Storage Path: {base_path}")
print(f"Config JSON Path: {CONFIG_JSON_PATH}")
print(f"Data File Path: {DATA_FILE_PATH}")
print(f"Service Account File Path: {SERVICE_ACCOUNT_FILE}")
print(f"Icon Path: {ICON_PATH}")