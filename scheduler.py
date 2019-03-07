#############################################################
#           PROJECT: TimeManager                            #
#           TEAM MEMBER: Lingchao Mao                       #
#                        YuQuan Cui                         #
#                        Ziwei Liu                          #
#                        Wenting Zheng                      #
#############################################################

# import numpy as np
# from datetime import datetime, timedelta
# import operator
# from sklearn import linear_model, preprocessing
# import pandas as pd
#############################################################
#         Some code is copied from this site                #
# https://developers.google.com/calendar/quickstart/python  #
#############################################################
from __future__ import print_function
import datetime
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']


#########################################
# The function is to print the upcoming #
# ten events in the calendar.           #
#########################################
def getEvents(service):
    # Call the Calendar API
    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
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

#############################
# The function is to add an #
# event to the calendar.    #
#############################
def addEvent(service):
    event = {
      'summary': 'TimeManager Try',
      'description': 'TimeManager Team Meeting',
      'start': {
        'dateTime': '2019-03-09T17:00:00-05:00',
        'timeZone': 'America/New_York'
      },
      'end': {
        'dateTime': '2019-03-09T18:00:00-05:00',
        'timeZone': 'America/New_York'
      },
      'attendees': [
        {'email': 'lmao3@ncsu.edu'},
        {'email': 'wzheng8@ncsu.edu', 'displayName': 'Wenting Zheng', 'organizer': True, 'self': True, 'responseStatus': 'accepted'}
      ],
      'reminders': {
        'useDefault': False
      }
    }
    event = service.events().insert(calendarId='primary', body=event).execute()
    print('Event created')
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
getEvents(service)
addEvent(service)