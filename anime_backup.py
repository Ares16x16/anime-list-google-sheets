import os
from datetime import datetime
import google.auth
from google.oauth2 import service_account
from googleapiclient.discovery import build

# Authenticate using a service account
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SERVICE_ACCOUNT_FILE = ''
creds = None
creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)

SPREADSHEET_ID = ''

# Connect to the Google Sheets API
service = build('sheets', 'v4', credentials=creds)

def update_anime(sheet_name, anime_name, num_episodes):
    """_summary_

    Args:
        sheet_name (_type_): Sheet name of anime to be update
        anime_name (_type_): Name of anime to be updated
        num_episodes (_type_): Desire number of episodes

    Returns:
        _type_: Return True if the update was successful, else False
    """
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=sheet_name).execute()
    values = result.get('values', [])
    row_index = None
    for i, row in enumerate(values):
        try:
            if row[2] == anime_name:
                row_index = i + 1
                break
        except:
            continue
    if row_index is None:
        return False

    range_name = f'{sheet_name}!D{row_index}'
    value_input_option = 'RAW'
    value_range_body = {
        'values': [[num_episodes]]
    }
    sheet.values().update(
        spreadsheetId=SPREADSHEET_ID, range=range_name,
        valueInputOption=value_input_option, body=value_range_body).execute()
    return True
    


def search_anime(sheet_name, anime_name):
    """_summary_

    Args:
        sheet_name (_type_):  Sheet name of anime to be update
        anime_name (_type_): Name of anime to be searched

    Returns:
        _type_: Return None if the anime is not found
                else Return number of episodes
    """
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=sheet_name).execute()
    values = result.get('values', [])
    row_index = None
    for i, row in enumerate(values):
        try:
            if row[2] == anime_name:
                row_index = i + 1
                break
        except:
            continue
    if row_index is None:
        return None

    range_name = f'{sheet_name}!D{row_index}'
    result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=range_name).execute()
    num_episodes = result['values'][0][0]
    return num_episodes

def list_anime(sheet_name):
    """_summary_

    Args:
        sheet_name (_type_): The sheet to be listed

    Returns:
        _type_: Return list of list of the anime data
    """
    sheet_metadata = service.spreadsheets().get(spreadsheetId=SPREADSHEET_ID).execute().get('sheets', '')
    sheet_id = None
    for sheet in sheet_metadata:
        if sheet['properties']['title'] == sheet_name:
            sheet_id = sheet['properties']['sheetId']
            break
    
    if sheet_id == None:
        return []

    range_name = f"{sheet_name}!A1:F1000"

    sheet = service.spreadsheets().values().get(spreadsheetId=SPREADSHEET_ID, range=range_name).execute()

    values = sheet.get('values', [])
    filtered_values = [row[:6] for row in values if len(row) >= 3 and row[2] != '']

    return filtered_values
            

def find_anime_in_sheets(anime_name):
    """_summary_

    Args:
        anime_name (_type_): Name of the anime

    Returns:
        _type_: List of matches: (sheets name, episodes)
    """
    current_year = datetime.now().year
    current_month = datetime.now().month
    seasons = ["Winter", "Spring", "Summer", "Fall"]
    sheet_names = ["All Time"]
    for year in range(2018, current_year + 1):
        for i, season in enumerate(seasons):
            if year == current_year and i*3+1 > current_month:
                break
            sheet_names.append(f"{year} {season}")
    
    matches = []
    for sheet_name in sheet_names:
        result = search_anime(sheet_name, anime_name)
        if result is not None:
            matches.append((sheet_name, result))

    return matches