#############################################################
#           PROJECT: TimeManager                            #
#           TEAM MEMBER: Lingchao Mao                       #
#                        YuQuan Cui                         #
#                        Ziwei Liu                          #
#                        Wenting Zheng                      #
#                                                           #
#         Some code is copied from this site                #
# https://developers.google.com/calendar/quickstart/python  #
#############################################################

#############################################################
#       This part includes functions used in the main()     #
#############################################################

from __future__ import print_function
from datetime import datetime, timedelta
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import pandas as pd
import numpy as np
from sklearn import linear_model, preprocessing
# from datetime import datetime, timedelta
# import operator
#########################################
# The function is to print the upcoming #
# ten events in the calendar.           #
#########################################
def getEvents(service):
    # Call the Calendar API
    now = datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    print('Getting the upcoming 10 events')
    events_result = service.events().list(calendarId='primary', timeMin=now,
                                        maxResults=10, singleEvents=True,
                                        orderBy='startTime').execute()
    events = events_result.get('items', [])

    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])

##############################
# The function find the free #
# time in the calendar       #
##############################
def getBusy(service):
    today = datetime.now()
    today = today.replace(hour=0, minute = 0, second = 0)
    the_datetime = today - timedelta(days=today.weekday()-14)
    the_datetime2 = the_datetime + timedelta(days=7)

    body = {
      "timeMin": the_datetime.isoformat() + 'Z',
      "timeMax": the_datetime2.isoformat() + 'Z',
      "timeZone": 'America/New_York',
      "items": [{"id": 'wzheng8@ncsu.edu'}]
    }
    events = service.freebusy().query(body=body).execute()
    return events['calendars']['wzheng8@ncsu.edu']['busy']

#############################
# The function is to add an #
# event to the calendar.    #
#############################
def addEvent(service, summaries, start, duration):
    summary = ""
    for sum in summaries:
        summary = summary + sum[2] + ", "
    end = start + timedelta(hours = duration)
    event = {
      'summary': summary,
      'description': 'TimeManager Team Meeting',
      'start': {
        'dateTime': start.strftime("%Y-%m-%dT%H:%M:%S"),
        'timeZone': 'America/New_York'
      },
      'end': {
        'dateTime': end.strftime("%Y-%m-%dT%H:%M:%S"),
        'timeZone': 'America/New_York'
      },
      'attendees': [
        {'email': 'wzheng8@ncsu.edu', 'displayName': 'Wenting Zheng', 'organizer': True, 'self': True, 'responseStatus': 'accepted'}
      ],
      'reminders': {
        'useDefault': False
      }
    }
    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created')

#################################
#     This function is to       #
#     get the free blocks       #
#################################
def get_free_blocks(events, schedule):
    for event in events:
        sta = datetime.strptime(event['start'][:-6], "%Y-%m-%dT%H:%M:%S")
        end = datetime.strptime(event['end'][:-6], "%Y-%m-%dT%H:%M:%S")
        day = sta.isoweekday()
        for val in schedule[day]:
            if val[0].time() < sta.time() and val[1].time() > end.time():
                schedule[day].append([end, val[1]])
                val[1] = val[1].replace(hour=sta.hour, minute=sta.minute)
#             overlap case 1: c < A < d < B
            if val[0].time() >= sta.time() and val[0].time() < end.time():
                val[0] = val[0].replace(hour=end.hour, minute=end.minute)
#             overlap case 2: A < c < B < d
            if val[1].time() > sta.time() and val[1].time() <= end.time():
                val[1] = val[1].replace(hour=sta.hour, minute=sta.minute)
    return schedule

#################################
#     This function is to get   #
#     duration time             #
#################################
def getduration(times):
    diff = times[1] - times[0]
    return diff.total_seconds()/3600

#################################
#     This function is to get   #
#     possible event options    #
#################################
def get_Event_Options(slot, tasks):
    result = {
        5:[],
        4:[],
        3:[],
        2:[],
        1:[]
    }
    for k, task in tasks.items():
        ratio = task[1]/getduration(slot)
        if slot[0] < task[3] and ratio <= 1:
            result[task[2]].append([k, ratio, task[4]])
#     sort based on ratio
    for k, val in result.items():
        if val:
            result[k] = sorted(result[k], key=lambda x: -x[1])
    return result

#################################
#     This function is to       #
#     select events             #
#################################
def choose(tasks):
    result = []
    current_ratio = 0
    for k, tasks in tasks.items():
        if tasks:
            for task in tasks:
                if current_ratio + task[1] <= 1:
                    result.append(task)
                    current_ratio = current_ratio + task[1]
    return result

#################################
#     This function is to       #
#     print time and tasks      #
#################################
def printTimeTask(start, duration, tasks):
    end = start + timedelta(hours = duration)
    print(start.strftime("%m/%d/%y %H:%M-"), end.strftime("%H:%M tasks: "), end="")
    for task in tasks:
        print("( id =",task[0], ")", task[2], end=", ")
    print()

#############################################################
#   This part is to import the historical data(xlsm) and    #
#   build a model to predict the duration of task.          #
#############################################################
#############
# load data #
#############
test = pd.read_excel(r'Tasks_test.xlsx')
testDF = pd.DataFrame(data=test)
test_column_list = testDF.columns.values.tolist()
test_x_array = np.array(testDF[ [test_column_list[3], test_column_list[4], test_column_list[5], test_column_list[6] ] ] )

task = pd.read_excel(r'Tasks_Training.xlsx')
taskDF = pd.DataFrame(data=task)
column_list = taskDF.columns.values.tolist()
##################################
# initialize duration dictionary #
##################################
duration = {}
test_id_array = testDF['ID']
length = len(test_id_array)

for i in range(length):
    duration[i] = []

test_category_array = testDF['Category']
category_num = []

# convert the category string into numbers
for category in test_category_array:
    if category == "School":
        category_num.append(2)
    elif category == "Personal":
        category_num.append(1)
    elif category == "Health":
        category_num.append(4)
    elif category == "Work":
        category_num.append(3)

for i in range(length):
    duration[i] = [ category_num[i] ]

list = []
my_array = np.array([])
for key in duration:
    key_value = duration[key] #a list
    value = key_value[0]
    #print(value)
    list.append(value)
my_array = np.array(list)

testDF[ test_column_list[2] ] =  my_array

test_x_array = np.array(testDF[ [test_column_list[2], test_column_list[3], test_column_list[4], test_column_list[5], test_column_list[6] ] ] )
test_x_array_normalized = preprocessing.normalize(test_x_array)
test_x_array_normalized_df = pd.DataFrame(test_x_array_normalized)
###############
# build model #
###############
x_array = np.array(taskDF[ [ column_list[10], column_list[3], column_list[4], column_list[5], column_list[6] ] ] )
x_array_normalized = preprocessing.normalize(x_array)


x_array_normalized_df = pd.DataFrame(x_array_normalized)
y_df = taskDF[ [column_list[7]] ]

### Return a dict for each category ###
multi_regr = linear_model.LinearRegression()
multi_regr.fit(x_array_normalized_df, y_df)
predict_y = multi_regr.predict(test_x_array_normalized)
predict_y_list = predict_y.tolist()

for key in duration:
    value_list = duration[key]
    value_list.append(predict_y_list[key][0])
print(duration)
#############################################################
#   This part is to connect to Google calendar and get the  #
#   schedule. The connection enables read/write.            #
#############################################################

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']
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
        creds = flow.run_local_server()
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)
# const {google} = require('googleapis')
# const calendar = google.calendar("v3")
service = build('calendar', 'v3', credentials=creds)
# getEvents(service)

######### Event object #########
# 2019-03-08T17:00:00-05:00 
# {'kind': 'calendar#event', 
# 'etag': '"3100936860372000"', 
# 'id': '500f86m17jsr2sts8074bfqpbo_20190308T220000Z', 
# 'status': 'confirmed', 
# 'htmlLink': 'https://www.google.com/calendar/event?eid=NTAwZjg2bTE3anNyMnN0czgwNzRiZnFwYm9fMjAxOTAzMDhUMjIwMDAwWiB3emhlbmc4QG5jc3UuZWR1', 
# 'created': '2019-02-18T05:39:17.000Z', 
# 'updated': '2019-02-18T05:40:30.186Z', 
# 'summary': 'Workout', 
# 'creator': {
#     'email': 'wzheng8@ncsu.edu', 
#     'displayName': 'Wenting Zheng', 
#     'self': True
# }, 
# 'organizer': {
#     'email': 'wzheng8@ncsu.edu', 
#     'displayName': 'Wenting Zheng', 
#     'self': True
# }, 
# 'start': {
#     'dateTime': '2019-03-08T17:00:00-05:00', 
#     'timeZone': 'America/New_York'
# }, 
# 'end': {
#     'dateTime': '2019-03-08T18:00:00-05:00', 
#     'timeZone': 'America/New_York'
# }, 
# 'recurringEventId': '500f86m17jsr2sts8074bfqpbo_R20190301T220000', 
# 'originalStartTime': {
#     'dateTime': '2019-03-08T17:00:00-05:00', 
#     'timeZone': 'America/New_York'
# }, 
# 'visibility': 'private', 
# 'iCalUID': '500f86m17jsr2sts8074bfqpbo_R20190301T220000@google.com', 
# 'sequence': 0, 
# 'attendees': [
#     {'email': 'lmao3@ncsu.edu', 'responseStatus': 'accepted'}, 
#     {'email': 'wzheng8@ncsu.edu', 'displayName': 'Wenting Zheng', 'organizer': True, 'self': True, 'responseStatus': 'accepted'}
# ], 
# 'extendedProperties': {'private': {'everyoneDeclinedDismissed': '-1'}}, 
# 'reminders': {'useDefault': True}}
######### Event object #########

#############################################################
#   This part is to get the free time blocks based on the   #
#   schedule in Google calendar.                            #
#############################################################
# Read in schedule in an array of dictionary
#  get next week's start date and end date
events = getBusy(service)
today = datetime.now().date()
print(today)
start = today - timedelta(days=today.weekday()-14)
start_time = datetime.strptime(input("Enter the time you prefer to start the work(eg. 13:30): "), "%H:%M").replace(year=start.year, month = start.month, day = start.day)
end_time = datetime.strptime(input("Enter the time you prefer to end the work(eg. 13:30): "), "%H:%M").replace(year=start.year, month = start.month, day = start.day)
free_time = {
    1:[[start_time, end_time]],
    2:[[start_time + timedelta(days=1), end_time + timedelta(days=1)]],
    3:[[start_time + timedelta(days=2), end_time + timedelta(days=2)]],
    4:[[start_time + timedelta(days=3), end_time + timedelta(days=3)]],
    5:[[start_time + timedelta(days=4), end_time + timedelta(days=4)]],
}
free_time = get_free_blocks(events, free_time)
tasks = np.array(testDF[ [test_column_list[4], test_column_list[7], test_column_list[1] ] ] )
i = 0
########################
# initialize task dict #
########################
# id => [category#, duration, priority, due time, name]
for k,v in duration.items():
    tasks[i][1] = tasks[i][1].to_pydatetime()
    duration[k].extend(tasks[i])
    i = i + 1
############################
# assign task to free time #
############################
i = 0
for time_slots in free_time.values():
    for slot in time_slots: #for each free time block
        possible_tasks = get_Event_Options(slot, duration)
        selected_tasks = choose(possible_tasks)
        if selected_tasks:
            sum_T = 0
            for t in selected_tasks:
                sum_T = sum_T + duration[t[0]][1]
                del duration[t[0]]
            printTimeTask(slot[0], sum_T,selected_tasks)
#             addEvent(service, selected_tasks, slot[0], sum_T)
        i = i + 1
####################
# fail to schedule #
####################
if duration:
    for k,v in duration.items():
        print("Task: ", v[4], "cannot fit in preferred working block. You have to use your leisure time!")
# addEvent(service)