from __future__ import print_function
import os, time, schedule
from datetime import datetime, timedelta
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

import xlsxwriter
from xlsxwriter import Workbook


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']

total = 0

def main():
    """Shows basic usage of the Google Calendar API.
    Prints the start and name of the next 10 events on the user's calendar.
    """
    creds = None
    page_token = None
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

    service = build('calendar', 'v3', credentials=creds)

    # Work calendarId
    workCalendar = "c8qj3ofbinh3fg2boqe3vbuohg@group.calendar.google.com"

    # Call the Calendar API
    weekAgo = datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
    weekAgo_object = datetime.strptime(weekAgo, "%Y-%m-%dT%H:%M:%S.%fZ")
    weekAgo_object = weekAgo_object - timedelta(days=14);
    weekAgo = datetime.strftime(weekAgo_object, "%Y-%m-%dT%H:%M:%S.%fZ" )
    now = datetime.utcnow().isoformat() + 'Z'
    print('Getting the upcoming 10 events')
    events_result = service.events().list(calendarId = workCalendar, timeMin=weekAgo,
                                        timeMax=now, singleEvents=True,
                                        orderBy='startTime').execute()

    events = events_result.get('items', []);

    wb = Workbook("invoice " + datetime.strftime(weekAgo_object,"%b %d %Y") + " - " + datetime.strftime(datetime.now(),"%b %d %Y") +  ".xlsx")
    sheet1 = wb.add_worksheet()
    
    sheet1.set_column(1, 1, 15)
    
    sheet1.write(0,0, "Client")
    sheet1.write(0,1, "Start Date")
    sheet1.write(0,2, "End Date")
    sheet1.write(0,3, "No. Hours")
    sheet1.write(0,4, "Rate")
    sheet1.write(0,5, "Class Type")
    sheet1.write(0,6, "Total")
    
    row = 1
    total = 0
    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])
        if "COLES" in event['summary']:
            continue
        else:
            writeToInvoice(event, wb, sheet1, row)
        row = row + 1

    calculateTotalRevenue(sheet1, row)
    wb.close()

# Extracts the client, date and class type from the event
def extractClientFromEvent(event):
    client = []
    eventSummary = event['summary']
    # Extract client name
    i = 0
    while eventSummary[i] != ':':
        client.append(eventSummary[i])
        i = i + 1
    # Convert from list object to string
    return ''.join(client)

def extractClassTypeFromEvent(event):
    classType = []
    eventSummary = event['summary']

    # Extract client name
    i = 0
    while eventSummary[i] != ':':
        i = i + 1
    
    i = i + 1
    while eventSummary[i] != '(':
        classType.append(eventSummary[i])
        i = i + 1
    # Convert from list object to string
    return ''.join(classType)

def writeToInvoice(event, wb, sheet1, row):
    global total
    text_format = wb.add_format({'text_wrap' : True})
    start = event['start'].get('dateTime', event['start'].get('date'))
    end = event['end'].get('dateTime', event['start'].get('date'))
    if "1.5hour" in event['summary']:
        hours = 1.5
    else:
        hours = 1
    rate = 27
    sheet1.write(row,0, extractClientFromEvent(event), text_format)
    sheet1.write(row,1, start, text_format)
    sheet1.write(row,2, end, text_format)
    sheet1.write(row,3, hours)
    sheet1.write(row,4, 27)
    sheet1.write(row,5, extractClassTypeFromEvent(event), text_format)
    sheet1.write(row,6, hours * rate)
    total += hours * rate

def calculateTotalRevenue(sheet1, row):
    sheet1.write(row, 6, total)


main()

'''
schedule.every().tuesday.do(main)
while True:
    
    start = time.time()
    schedule.run_pending()
    time.sleep(5)
    end = time.time()
    # If five seconds have elapsed, not tuesday, exit
    if end - start > 0:
        break
        
'''    