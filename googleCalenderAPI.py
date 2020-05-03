 # some of this code was adapted from
# https://towardsdatascience.com/accessing-google-calendar-events-data-using-python-e915599d3ae2
import exchangeSearch
import datetime
import pickle
import os
from apiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

month_dictionary = {
    "January" : "01",
    "February" : "02",
    "March" : "03",
    "April" : "04",
    "May" : "05",
    "June" : "06",
    "July" : "07",
    "August" : "08",
    "September" : "09",
    "October" : "10",
    "November" : "11",
    "December" : "12"
}

# get time
now = datetime.datetime.now()

#
## Create events via google script API
def create_event(importantEmail, service):
    # extract important email info
    subject_line = importantEmail.subject
    email_body = importantEmail.email_body
    deadline = importantEmail.deadline.group().split()
    # check if month is this year or next year
    if int(month_dictionary[deadline[0]]) > now.month:
        currentYear = now.year
    else:
        currentYear = now.year + 1
    deadline_formatted = str(currentYear) + "-" + month_dictionary[deadline[0]] + "-" + deadline[1] + "T09:00:00+03:00"
    event = {
    'summary': subject_line,
    #'location': location,  # optional addition
    'description': email_body,
    'start': {
        'dateTime': deadline_formatted,
        'timeZone': 'Asia/Qatar',
    },
    'end': {
        'dateTime': deadline_formatted,
        'timeZone': 'Asia/Qatar',
    },
    }
    event = service.events().insert(calendarId='primary', body=event).execute()


#
## Main program
def main():
    # specify scope for project (what permissions im getting from user)
    scopes = ['https://www.googleapis.com/auth/calendar']
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
                "client_secret.json", scopes)
            creds = flow.run_local_server()
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)


    service = build("calendar", "v3", credentials=creds)

    result = service.calendarList().list().execute()
    calendar_id = result['items'][0]['id']
    result = service.events().list(calendarId=calendar_id).execute()
    # note: getting event summaries isn't working so I store created events
    # in a local file
    important_emails = exchangeSearch.collect_important_emails()
    # initialize boolean to check if event already registered
    eventCreated = False
    # open written events file
    with open("events_created.txt", "a+") as eventFile:
        eventFile.seek(0)
        eventTitles = eventFile.read().split('\n')
        # go through each important email, check if written,if not add to file
        for email in important_emails:
            if eventTitles:     # if not empty check events
                for event in eventTitles:
                    if event == important_emails[email].subject:
                        eventCreated = True
            if not eventCreated:
                create_event(important_emails[email], service)
                eventFile.write(important_emails[email].subject + "\n")
            eventCreated = False    # reset boolean
        eventFile.close()   # close when done


# Run program
main()
