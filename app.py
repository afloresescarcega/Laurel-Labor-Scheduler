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

OUTPUT_OUTPUT_RANGE = "Fixed!C2:I"


class Person:
    def __init__(self, _name):
        self.name = _name
        self.days_available = [[], [], [], [], [], [], []] # first index has available time ranges for Monday; last index for Sunday

    def __str__(self):
        return self.name + ":\n" + "Mondays: " + str(self.days_available[0]) + "\nTuesdays: "  + str(self.days_available[1]) + "\nWednesdays: " + str(self.days_available[2]) + "\nThursdays: " + str(self.days_available[3]) + "\nFridays: " + str(self.days_available[4]) + "\nSaturdays: " + str(self.days_available[5]) + "\nSundays: " + str(self.days_available[6]) + "\n"
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
        self.people = []

def publish(laborSchedule=None):
    from pprint import pprint

    from googleapiclient import discovery

    # TODO: Change placeholder below to generate authentication credentials. See
    # https://developers.google.com/sheets/quickstart/python#step_3_set_up_the_sample
    #
    # Authorize using one of the following scopes:
    #     'https://www.googleapis.com/auth/drive'
    #     'https://www.googleapis.com/auth/drive.file'
    #     'https://www.googleapis.com/auth/spreadsheets'
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

    # The ID of the spreadsheet to update.
    spreadsheet_id = '1Zz6otTDbeqJWWatzNf4ErWvc3lVuDUn4b1EPKff6xpk'  # TODO: Update placeholder value.

    # The A1 notation of the values to update.
    range_ = OUTPUT_OUTPUT_RANGE  # TODO: Update placeholder value.

    # How the input data should be interpreted.
    value_input_option = 'RAW'  # TODO: Update placeholder value.

    # rows are labor positions
    # columns are days that people are available for the time that the position occurs

    values = []

    # for dealing with DCU edge case
    monday_thru_sat = []
    need_to_add_dcu_mon_sat = True
    finished_dcu_mon_sat = False

    for position in laborSchedule.positions:
        # Sunday DCU is separate in laborSchedule but needs to combined with normal DCU
        if True:
            values.append([", ".join([name for name in names]) for names in position.days])
        # else:
        #     if "Sundays" not in position.name:
                
        #         for index, day in enumerate(position.days):
        #             if index != 6:
        #                 # values.append([position.name + ", ".join([name for name in names]) for names in position.days])
        #                 monday_thru_sat.append([position.name + ", ".join([name for name in position.days[index]])])
        #             else:
        #                 break
        #     else: # Sundays
        #                                         #      [position.name + ", ".join([name for names]) for names in position.days[index]]
        #         monday_thru_sunday = monday_thru_sat + [ position.name + ", ".join([name for name in position.days[6]])]
        #         print("UUUUUU")
        #         print("UUUUUU")
        #         print("UUUUUU")
        #         print("UUUUUU")
        #         print("UUUUUU")
        #         h = [l[0] for l in monday_thru_sunday]
        #         print(h) 
        #         return
        #         values.append(h) 
    # print()
    # print()
    # print()
    # print(values)
    # return 
    value_range_body = {
        'values': values
    }

    request = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_, valueInputOption=value_input_option, body=value_range_body)
    response = request.execute()

    # TODO: Change code below to process the `response` dict:
    pprint(response)

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

    output_result = sheet.values().get(spreadsheetId=OUTPUT_SPREADSHEET_ID_COPY, range=OUTPUT_RANGE).execute()
    output_values = output_result.get('values', [])


    # for line in output_values:
    #     print(line)

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
            # if "Sundays" in _name:
            #     position_1 = LaborPosition(match_name, start_hour, end_hour, start_minutes, end_minutes)
            #     position_2 = LaborPosition(match_name + " Sundays", start_hour_2, end_hour_2, start_minutes_2, end_minutes_2)
            #     laborSchedule.positions.append(position_1)
            #     laborSchedule.positions.append(position_2)
            #     print("Creating two positions", match_name)

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

        # Make the time ranges for every person 
        for person_row in values:
            # create a person object
            full_name = person_row[2] + " " + person_row[3]
            person = Person(full_name)
            for time_slot in range(COL_OFFSET, len(person_row)):
                i = time_slot - COL_OFFSET
                if not person_row[time_slot]: # Skip because not days with this time chunk
                    # print(full_name, "skipped time slot", time_slot)
                    continue
                days_available_in_chunk = person_row[time_slot].split(", ")
                # print(full_name, "has these days", days_available_in_chunk)
                if days_available_in_chunk: # has a day available for this time chunk

                    if i < 31:
                        person_available_time_chunk = DateTimeRange("T"+str(8 + i//2) + ":" + str(3*(i%2)) + "0"+":00-0600", "T"+str(8 + (i + 1)//2) + ":" + str(3*((i + 1)%2)) + "0"+":00-0600")
                        # print(full_name, "has this person_available_time_chunk", person_available_time_chunk)
                    elif i == 31:
                        person_available_time_chunk = DateTimeRange("T"+str(8 + i//2) + ":" + str(3*(i%2)) + "0"+":00-0600", "T"+str(8 + (i)//2) + ":59"+":00-0600")
                        # print(full_name, "has this person_available_time_chunk", person_available_time_chunk)
                    for day in days_available_in_chunk:
                        if not day: # Skip because not days with this time chunk
                            # print(full_name, "skipped day", day)
                            continue
                        day_index = DAYS.index(day) # Monday = 0 ... Sunday = 6
                        assert(day_index <= 6)
                        if not person.days_available[day_index]: # If initial time chunk for this day, don't coalesce
                            # print(full_name, "has no previous time for", day_index, ". This is the before:", person.days_available[day_index])
                            person.days_available[day_index].append(person_available_time_chunk)
                            # print(full_name, "now has time for", day_index,":", "This is the after:", person.days_available[day_index])
                        if len(person.days_available[day_index]) >= 1: # There is a previous time chunk for this day
                            # print(full_name, "has  previous time for", day_index, ". This is the before:", person.days_available[day_index])
                            last_time_chunk = person.days_available[day_index][-1]
                            # print(full_name, "'s previous time for", day_index, ":",last_time_chunk)
                            if last_time_chunk.is_intersection(person_available_time_chunk): # times do intersect
                                if last_time_chunk.intersection(person_available_time_chunk).get_timedelta_second() == 0: # coalesce
                                    # print("Coalescing last time chunk: ", last_time_chunk, "and", person_available_time_chunk)
                                    s = last_time_chunk.encompass(person_available_time_chunk)
                                    person.days_available[day_index].pop() # remove the previous time because we are coalescing them
                                    person.days_available[day_index].append(s)
                            else:
                                # print(full_name, "'s times do not intersect:", last_time_chunk.is_intersection(person_available_time_chunk))
                                person.days_available[day_index].append(person_available_time_chunk)

                                




                # # Comes up with the 30 minute time chunk
                # person_available_time_chunk = DateTimeRange("T"+str(8 + i//2) + ":" + str(3*(i%2)) + "0"+":00-0600", "T"+str(8 + i//2) + ":" + str(3*(time_slot%2)) + "5"+":00-0600")
                # if i == 0: # first doesn't have a previous time chunk
                #     person.
            laborSchedule.people.append(person)
        for p in laborSchedule.people:
            print(p)
            for labor_position in laborSchedule.positions:
                for day_index in range(7): # 0 = Monday
                    for person_available_time_chunk in p.days_available[day_index]:
                        if labor_position.time_range in person_available_time_chunk:
                            labor_position.days[day_index].append(p.name)
                            break # don't need to continue looking for a time

        print("\n\n\n\n")
        for l in laborSchedule.positions:
            print("\n")
            print(l)
        print("\n\n\n\n")

        publish(laborSchedule)


        """
        Do not use the bottom
        """
        # for person in values:
        #     print("")
        #     # Print columns A and E, which correspond to indices 0 and 4.
        #     # print('%s %s, ' % (person[2], person[3]))
        #     # print("~ ", len(person), str(person))
        #     full_name = person[2] + " " + person[3]
        #     print("Person:", full_name)
        #     for time_slot in range(COL_OFFSET, len(person)): # first time_slot is 8:00am; Minus, timestamp, email, first name, last name 
        #         i = time_slot - 4
        #         # print("T"+str(8 + i//2) + ":" + str(3*(i%2)) + "0"+":00-0600", "T"+str(8 + i//2) + ":" + str(3*(i%2)) + "5"+":00-0600")
        #         start_time = DateTimeRange("T"+str(8 + i//2) + ":" + str(3*(i%2)) + "0"+":00-0600", "T"+str(8 + i//2) + ":" + str(3*(time_slot%2)) + "5"+":00-0600")
        #         pos_time = l.time_range
        #         # print("Position name:", l.name, "Position time", pos_time, "Start time:", start_time)
        #         print()
        #         # print("Timeslot", time_slot - COL_OFFSET, "len", len(person))
        #         print()
        #         days_available = person[time_slot - COL_OFFSET].split(", ")
        #         print("Days", full_name, "is able to work:", days_available)
        #         for position in laborSchedule.positions:
        #             for day in days_available: # days that they are able to work for this timeslot
        #                 if start_time in pos_time and day: # they are able to work during this time
        #                     print("Start_time", start_time, "pos_time", pos_time, "day", day)
        #                     position.days[DAYS.index(day)].append(full_name)
        #                     print("Adding", full_name, "to", l.name, "on", day)

    # for l in laborSchedule.positions:
    #     pass
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
