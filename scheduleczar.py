####################################################################################################################################
# There are two parts of this script. One that uses the google calendar API to check for the presence of the string '[T]' in the title
# of a microscope slot on the calendar and a second that sends an email if certain timing conditions are met based on the current time 
# and the appointment start time.
#
# This script is meant to run as an hourly cron job or something similar (see czar.sh)
# 
# Setting up the google calendar API requires a google cloud account and the app must be registered. I pay a few cents per month for this 
# see here for some semi-useful info: https://developers.google.com/calendar/api/quickstart/python
####################################################################################################################################


from __future__ import print_function

import datetime
import os.path
import numpy as np

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

import smtplib
from string import Template
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/calendar', 'https://www.googleapis.com/auth/calendar.events']


def main():


    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time. i.e. it will open a browser and you can log in
    if os.path.exists('/home/jason/Software/calendarHandler/token.json'):
        creds = Credentials.from_authorized_user_file('/home/jason/Software/calendarHandler/token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                '/home/jason/Software/calendarHandler/credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('/home/jason/Software/calendarHandler/token.json', 'w') as token:
            token.write(creds.to_json())

    try:
        service = build('calendar', 'v3', credentials=creds)

        # Call the Calendar API
        now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
        dt_now = datetime.datetime.strptime(now[:-1] + '-00:00', "%Y-%m-%dT%H:%M:%S.%f%z")

        print(now + ' Getting the upcoming 20 events')
        events_result = service.events().list(calendarId='hnestvmas1von3v9hufn1jkugo@group.calendar.google.com', timeMin=now,
                                              maxResults=20, singleEvents=True,
                                              orderBy='startTime').execute()

        events = events_result.get('items', [])

        if not events:
            print('No upcoming events found.')
            return

        # Prints the start and name of the next 10 events
        for event in events:
            start = event['start'].get('dateTime', event['start'].get('date'))

            dt_now = datetime.datetime.strptime(now[:-1] + '-00:00', "%Y-%m-%dT%H:%M:%S.%f%z")
            dt_start = datetime.datetime.strptime(start, "%Y-%m-%dT%H:%M:%S%z")

            hours_to_start = np.round((dt_start - dt_now).total_seconds()/3600., 2)
            #


            if ('[T]' in event['summary']) & (hours_to_start < 72):
                print('DELETING')
                print(event['id'], start, hours_to_start, '[T]' in event['summary'], event['summary'], event['creator']['email'])

                multipart_message = MIMEMultipart()
                SEND_FROM = 'weiner.microscope.czar@gmail.com'
                EMAIL_PWD = '' #ask Jason for the password and put it here
                msg = f"Hello,\n\nIn accordance with the lab's policy concerning opt-in usage of the microscope, I am deleting a slot marked as tentative from the calendar. The deleted appointment, '{event['summary']}', was set to start at {start}.\n\nThank you and I apologize for any inconvenience,\nThe Weiner Lab Microscope Czar"
                multipart_message['From'] = SEND_FROM
                multipart_message['To'] = event['creator']['email']
                multipart_message['Subject'] = "Tentative Microscope Slot Removed"

                # set up the SMTP server
                smtplib_server = smtplib.SMTP(host='smtp.gmail.com', port=587)
                smtplib_server.starttls()
                smtplib_server.login(SEND_FROM, EMAIL_PWD)

                # add in the message body
                multipart_message.attach(MIMEText(msg, 'plain'))
                # send the message via the server set up earlier.
                smtplib_server.send_message(multipart_message)
                del multipart_message
                # Terminate the SMTP session and close the connection
                smtplib_server.quit()

                service.events().delete(calendarId='hnestvmas1von3v9hufn1jkugo@group.calendar.google.com', eventId=event['id']).execute()

            elif ('[T]' in event['summary']) & (hours_to_start > 95) & (hours_to_start < 97):
                print('WARNING')
                print(event['id'], start, hours_to_start, '[T]' in event['summary'], event['summary'], event['creator']['email'])

                multipart_message = MIMEMultipart()
                SEND_FROM = 'weiner.microscope.czar@gmail.com'
                EMAIL_PWD = '' #ask Jason for the password and put it here
                msg = f"Hello,\n\nIn accordance with the lab's policy concerning opt-in usage of the microscope, I am considering deleting a slot marked as tentative from the calendar. The soon-to-be deleted appointment, '{event['summary']}', is set to start at {start}. If you wish to avoid this, please remove the [T] marker from your scheduled appointment. \n\nThank you and I apologize for any inconvenience,\nThe Weiner Lab Microscope Czar"
                multipart_message['From'] = SEND_FROM
                multipart_message['To'] = event['creator']['email']
                multipart_message['Subject'] = "Warning: Tentative Microscope Slot Removal"

                # set up the SMTP server
                smtplib_server = smtplib.SMTP(host='smtp.gmail.com', port=587)
                smtplib_server.starttls()
                smtplib_server.login(SEND_FROM, EMAIL_PWD)

                # add in the message body
                multipart_message.attach(MIMEText(msg, 'plain'))
                # send the message via the server set up earlier.
                smtplib_server.send_message(multipart_message)
                del multipart_message
                # Terminate the SMTP session and close the connection
                smtplib_server.quit()

    except HttpError as error:
        print('An error occurred: %s' % error)


if __name__ == '__main__':
    main()
