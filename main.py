import gspread
from oauth2client.service_account import ServiceAccountCredentials


# Path to  service account JSON file
SERVICE_ACCOUNT_FILE = 'iu-patient-tracking-040d4a8eeb27.json'  # Update with your actual filename


# Google Sheets URL and worksheet name
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1fpXJYv8QRXx2B5qbDoTVec5MoXxE6WjEtzqe5xKQroI/edit?usp=sharing'
WORKSHEET_NAME = 'Form Responses 1'


# Set up credentials and client
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
client = gspread.authorize(creds)


# Open the sheet
sheet = client.open_by_url(SHEET_URL)
worksheet = sheet.worksheet(WORKSHEET_NAME)


# Get all values
data = worksheet.get_all_values()




# Print columns and their values
if data:
    headers = data[0]
    columns = {header: [] for header in headers}
    for row in data[1:]:
        for i, value in enumerate(row):
            if i < len(headers):
                columns[headers[i]].append(value)
    for header in headers:
        print(f"Column: {header}")
        for value in columns[header]:
            print(f"  {value}")
        print()


    # Conflict detection and update
    location_col = "Patient location/destination (Physical location/Room)"
    team_col = "Your team (not the team you're handing off to)"
    status_col = "Status"
    flag_reason_col = "Flag Reason"


    # Build index for location and team
    location_index = {}
    team_index = {}

    for idx, row in enumerate(data[1:], start=2):  # start=2 for correct row number in sheet
        location = row[headers.index(location_col)] if location_col in headers else None
        team = row[headers.index(team_col)] if team_col in headers else None


        status_col_num = headers.index(status_col)+1
        flag_reason_col_num = headers.index(flag_reason_col)+1


        # Check for location conflict
        if location and location in location_index:
            print(f"Updating row {idx}: Status -> CONFLICT, Flag Reason -> Location Already Occupied")
            worksheet.update_cell(idx, status_col_num, "CONFLICT")
            worksheet.update_cell(idx, flag_reason_col_num, "Location Already Occupied")
        else:
            location_index[location] = idx


        # Check for team conflict
        if team and team in team_index:
            print(f"Updating row {idx}: Status -> CONFLICT, Flag Reason -> Duplicate Team Assignment")
            worksheet.update_cell(idx, status_col_num, "CONFLICT")
            worksheet.update_cell(idx, flag_reason_col_num, "Duplicate Team Assignment")
        else:
            team_index[team] = idx


else:
    print("No data found in the sheet.")
