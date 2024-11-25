# src/utils/google_sheets.py

import logging
import pandas as pd
from googleapiclient.discovery import build
from google.oauth2 import service_account
from config.settings_manager import (
    SERVICE_ACCOUNT_FILE,
    SPREADSHEET_ID,
    RANGE_NAME
)

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

def get_service():
    """
    Authenticates with Google Sheets API using a service account file and returns the service object.
    """
    try:
        creds = service_account.Credentials.from_service_account_file(
            SERVICE_ACCOUNT_FILE, scopes=SCOPES
        )
        service = build('sheets', 'v4', credentials=creds, cache_discovery=False)
        return service
    except Exception as e:
        logging.error(f"Error obtaining Google Sheets service: {e}")
        raise

def read_from_google_sheets():
    """
    Reads data from the specified Google Sheets document and returns it as a pandas DataFrame.
    """
    try:
        service = get_service()
        sheet = service.spreadsheets()
        result = sheet.values().get(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME
        ).execute()
        values = result.get('values', [])
        if values:
            df = pd.DataFrame(values[1:], columns=values[0])
            return df
        else:
            print("No data found in Google Sheets.")
            return pd.DataFrame()
    except Exception as e:
        logging.error(f"Error reading from Google Sheets: {e}")
        return pd.DataFrame()

def write_to_google_sheets(df):
    """
    Writes the provided pandas DataFrame to the specified Google Sheets document.
    """
    try:
        service = get_service()
        # Replace NaN values with empty strings
        df = df.fillna('')
        # Convert DataFrame to a list of lists
        values = [df.columns.values.tolist()] + df.values.tolist()
        body = {'values': values}

        # Clear the existing data in the sheet before writing new data
        clear_values_request_body = {}
        service.spreadsheets().values().clear(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME,
            body=clear_values_request_body
        ).execute()

        # Write the new data to the sheet
        service.spreadsheets().values().update(
            spreadsheetId=SPREADSHEET_ID,
            range=RANGE_NAME,
            valueInputOption='RAW',
            body=body
        ).execute()
        print("Data written to Google Sheets successfully.")
    except Exception as e:
        logging.error(f"Error writing to Google Sheets: {e}")
        raise

def delete_row_in_google_sheets(row_index):
    """
    Deletes a row in the Google Sheets document at the specified index.
    Note: row_index is 1-based (1 corresponds to the first row).
    """
    try:
        service = get_service()
        requests = [{
            'deleteDimension': {
                'range': {
                    'sheetId': 0,  # Default sheet ID; change if necessary
                    'dimension': 'ROWS',
                    'startIndex': row_index - 1,  # Zero-based index
                    'endIndex': row_index        # Exclusive end index
                }
            }
        }]
        body = {'requests': requests}
        service.spreadsheets().batchUpdate(
            spreadsheetId=SPREADSHEET_ID,
            body=body
        ).execute()
        print(f"Row {row_index} deleted from Google Sheets successfully.")
    except Exception as e:
        logging.error(f"Error deleting row in Google Sheets: {e}")
        raise