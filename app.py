from __future__ import print_function
import pickle
import os.path
from googleapiclient import discovery
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

from datetimerange import DateTimeRange
import re

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms'
SAMPLE_RANGE_NAME = 'Form Responses 1!A2:AJ'

INPUT_SPREADSHEET_ID = '1OGJHwOWpzCu7L6uG2uech6tEqBl6jCRYiDW6c9JDx2k'
OUTPUT_SPREADSHEET_ID = '1_tL1iArkqfsQ0GRcc286AvcmCZnT_lmVfNmCVEWpdro'
OUTPUT_SPREADSHEET_ID_COPY = '1Zz6otTDbeqJWWatzNf4ErWvc3lVuDUn4b1EPKff6xpk'

OUTPUT_RANGE = "Fixed!A2:I"


class LaborPosition:
    def __init__(self, _name, start_hour, finish_hour, start_minute="00", finish_minute="00"):
        # start_hour and finish_hour must be 24hr time, integer
        self.name = _name
        # example
        # time_range = DateTimeRange("T10:00:00-0600", "T10:10:00-0600")
        # +0900 is for timezone
        self.time_range = DateTimeRange("T"+start_hour+start_minute+":00-0600", "T"+finish_hour+finish_minute+":00-0600")

        # Days that this labor position does labor
        self.days = [[],[],[],[],[],[],[]] # People available for that day; Monday is first list, Sunday is last list

    def __str__(self):
        return str(self.name) + " " + str(self.time_range) + " " + str(self.days)


class LaborSchedule:
    def __init__(self):
        self.positions = []

def main():
    """Shows basic usage of the Sheets API.
    Prints values from a sample spreadsheet.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)

    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    result = sheet.values().get(spreadsheetId=INPUT_SPREADSHEET_ID, range=SAMPLE_RANGE_NAME).execute()
    values = result.get('values', [])
    # formats = sheet.get(spreadsheetId=INPUT_SPREADSHEET_ID, includeGridData=True).execute()

    output_result = sheet.values().get(spreadsheetId=OUTPUT_SPREADSHEET_ID, range=OUTPUT_RANGE).execute()
    output_values = output_result.get('values', [])


    for line in output_values:
        print(line)

    # For each line, create a labor position and store in schedule

    PATTERN_FOR_HOURS = "((\d+)?(:\d\d)?([ap]m)?)-((\d+)?(:\d\d)?([ap]m)?)"
    laborSchedule = LaborSchedule()

    for position in output_values:
        # _pos = LaborPosition()
        _name = position[0]
        match_hours = re.findall(PATTERN_FOR_HOURS, _name)
        match_name = re.match("[^(]*", _name).group(0).strip()
        # print(match_hours)
        # group 2 is start hour and group 6 is end hour 
        # group 4 is the start hours am or pm and group 8 is the end hour am or pm
        # group 3 is start minutes an group 7 is the end minutes

        start_hour = match_hours[0][1] if match_hours[0][3] == "am" else str(int(match_hours[0][1]) + 12)
        end_hour = match_hours[0][5] if match_hours[0][7] == "am" else str(int(match_hours[0][5]) + 12)

        start_minutes = ":00" if not match_hours[0][2] else match_hours[0][2]
        end_minutes = ":00" if not match_hours[0][6] else match_hours[0][6]

        # special case for 12pm as start hour; don't add 12
        if match_hours[0][1] == "12" and match_hours[0][3] == "pm":
            start_hour = match_hours[0][1]

        # special case for 12am as start hour; add 12
        if match_hours[0][1] == "12" and match_hours[0][3] == "am":
            start_hour = str(int(match_hours[0][1]) + 12)


        # special case for 12pm as end hour; don't add 12
        if match_hours[0][5] == "12" and match_hours[0][7] == "pm":
            end_hour = match_hours[0][5]

        # special case for 12am as end hour; add 11 hours and 59 minutes
        if match_hours[0][5] == "12" and match_hours[0][7] == "am":
            end_hour = str(int(match_hours[0][5]) + 11)
            end_minutes = ":59"


        

        # print(match_name, match_hours[0][0], match_hours[0][4], "****", start_hour +  start_minutes, end_hour + end_minutes)

        # If there is a position with multiple times, create the positions 
        # separately
        if len(match_hours) > 1: # this position has hours that vary per day
            start_hour_2 = match_hours[1][1] if match_hours[1][3] == "am" else str(int(match_hours[1][1]) + 12)
            end_hour_2 = match_hours[1][5] if match_hours[1][7] == "am" else str(int(match_hours[1][5]) + 12)
            start_minutes_2 = ":00" if not match_hours[1][2] else match_hours[1][2]
            end_minutes_2 = ":00" if not match_hours[1][6] else match_hours[1][6]
            # print(match_name, match_hours[0][0], match_hours[0][4], match_hours[1][0], match_hours[1][4])
            if "Sundays" in _name:
                position_1 = LaborPosition(match_name, start_hour, end_hour, start_minutes, end_minutes)
                position_2 = LaborPosition(match_name + " Sundays", start_hour_2, end_hour_2, start_minutes_2, end_minutes_2)
                laborSchedule.positions.append(position_1)
                laborSchedule.positions.append(position_2)
                print("Creating two positions", match_name)

        else:
            # _start_minutes = ":00" if not match_hours[0][2] else match_hours[0][2]
            # _end_minutes = ":00" if not match_hours[0][6] else match_hours[0][6]
            # print(match_name, match_hours[0][0], match_hours[0][4])
            position = LaborPosition(match_name, start_hour, end_hour, start_minutes, end_minutes)
            print("Creating a position", match_name)
            laborSchedule.positions.append(position)

    for l in laborSchedule.positions:
        print(l)
    print("\n\n\n\n")

    # Start populating the positions
    

    if not values:
        print('No data found.')
    else:
        print('Full Name, :')
        DAYS = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        COL_OFFSET = 4 # Offset from the start of the timeslots
        for person in values:
            print("")
            # Print columns A and E, which correspond to indices 0 and 4.
            # print('%s %s, ' % (person[2], person[3]))
            # print("~ ", len(person), str(person))
            full_name = person[2] + " " + person[3]
            print("Person:", full_name)
            for time_slot in range(COL_OFFSET, len(person)): # first time_slot is 8:00am; Minus, timestamp, email, first name, last name 
                i = time_slot - 4
                # print("T"+str(8 + i//2) + ":" + str(3*(i%2)) + "0"+":00-0600", "T"+str(8 + i//2) + ":" + str(3*(i%2)) + "5"+":00-0600")
                start_time = DateTimeRange("T"+str(8 + i//2) + ":" + str(3*(i%2)) + "0"+":00-0600", "T"+str(8 + i//2) + ":" + str(3*(time_slot%2)) + "5"+":00-0600")
                pos_time = l.time_range
                # print("Position name:", l.name, "Position time", pos_time, "Start time:", start_time)
                print()
                # print("Timeslot", time_slot - COL_OFFSET, "len", len(person))
                print()
                days_available = person[time_slot - COL_OFFSET].split(", ")
                print("Days", full_name, "is able to work:", days_available)
                for position in laborSchedule.positions:
                    for day in days_available: # days that they are able to work for this timeslot
                        if start_time in pos_time and day: # they are able to work during this time
                            print("Start_time", start_time, "pos_time", pos_time, "day", day)
                            position.days[DAYS.index(day)].append(full_name)
                            print("Adding", full_name, "to", l.name, "on", day)

    for l in laborSchedule.positions:
        pass
        # print()
        # print(l)
        # print()

            

if __name__ == '__main__':
    main()
    colors = []

    # for row_num, row in enumerate(["sheets"]["data"]["rowData"]):
    #     row_of_colors = []
    #     for col_num, cell in enumerate(row["values"]):
    #         cell_color = cell["userEnteredFormat"]["backgroundColor"]

    #         # Is the color of the cell a custom color, other than white?
    #         # If so, then this is a blacked out cell
    #         # if cell_color["red"] < 1.0 or cell_color["green"] < 1.0 or cell_color["blue"] < 1.0:

    #     colors.append(row_of_colors)
