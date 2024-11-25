# src/utils/file_io.py

import pandas as pd
from config.settings_manager import DATA_FILE_PATH

def read_applications_from_excel(file_path=DATA_FILE_PATH):
    try:
        return pd.read_excel(file_path)
    except FileNotFoundError:
        return pd.DataFrame(columns=["Company", "Position", "Application Portal URL", "Date Applied", "Status"])

def save_applications_to_excel(df, file_path=DATA_FILE_PATH):
    df.to_excel(file_path, index=False)