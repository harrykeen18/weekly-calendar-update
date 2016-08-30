from __future__ import print_function
import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

import datetime

# email stuff
import base64
from email.mime.audio import MIMEAudio
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import mimetypes
import os

from apiclient import errors

from time import strftime, strptime
from datetime import timedelta

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/calendar-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/calendar.readonly https://www.googleapis.com/auth/gmail.compose'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Google Calendar API Python Quickstart'


def get_credentials():
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'calendar-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def create_message(sender, to, subject, message_text):
    """Create a message for an email.

    Args:
    sender: Email address of the sender.
    to: Email address of the receiver.
    subject: The subject of the email message.
    message_text: The text of the email message.

    Returns:
    An object containing a base64url encoded email object.
    """
    message = MIMEText(message_text)
    message['to'] = to
    message['from'] = sender
    message['subject'] = subject
    return {'raw': base64.urlsafe_b64encode(message.as_string())}

def send_message(service, user_id, message):
    """Send an email message.

    Args:
    service: Authorized Gmail API service instance.
    user_id: User's email address. The special value "me"
    can be used to indicate the authenticated user.
    message: Message to be sent.

    Returns:
    Sent Message.
    """
    try:
        message = (service.users().messages().send(userId=user_id, body=message).execute())
        print('Message Id: %s' % message['id'])
        return message
    except errors.HttpError, error:
        print('An error occurred: %s' % error)

def main():
    """Shows basic usage of the Google Calendar API.

    Creates a Google Calendar API service object and outputs a list of the next
    10 events on the user's calendar.
    """

    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    service = discovery.build('calendar', 'v3', http=http)

    now = datetime.datetime.utcnow().isoformat() + 'Z' # 'Z' indicates UTC time
    three_weeks_from_now = datetime.datetime.utcnow() + timedelta(days=21)
    three_weeks_from_now = three_weeks_from_now.isoformat() + 'Z'

    print('Getting all the events in the next three weeks')
    eventsResult = service.events().list(
        calendarId='fabhub.io_5mietl2ubcfslni3ltlr61n28c@group.calendar.google.com', timeMin=now, timeMax=three_weeks_from_now, singleEvents=True,
        orderBy='startTime').execute()
    events = eventsResult.get('items', [])

    raw_event_array = []
    event_array = []

    if not events:
        print('No upcoming events found.')
    for event in events:
        # start = event['start'].get('dateTime', event['start'].get('date'))
        start = event['start'].get('date')
        end = event['end'].get('date')
        raw_event_array.append([event['start'], event['end'], event['summary']])
        # event_array.append(event['summary'])

    email_service = discovery.build('gmail', 'v1', http=http)

    for event in raw_event_array:

        start_end = []

        for n in [0, 1]:

            test_date = event[n]

            if "dateTime" in test_date.keys():
                date_time = str(test_date['dateTime'])
                temp_date = date_time[:10]
                date_format = strptime(temp_date, "%Y-%m-%d")

            else:
                date_time = str(test_date['date'])
                date_format = strptime(date_time, "%Y-%m-%d")

            start_end.append(strftime("%a, %d %b %Y", date_format))

        event_array.append([start_end[0], start_end[1], event[2]])

    # print(event_array)

    search_words = ["holiday", "Holiday", "HOLIDAY", "off", "OFF", "Off", "Away", "AWAY", "away", "out", "OUT", "Out", "leave", "Leave", "LEAVE"]
    email_string = "Good morning Opendesk, Basil here, your friendly holiday manager, in the next three weeks these lucky people are on holiday;\n\n"
    for event in event_array:
        for word in search_words:
            try:
                if word in str(event[2]):

                    email_string = email_string + str(event[2]) + " from " + event[0] + " back on " + event[1]
                    email_string = email_string + "\n\n"
            except:
                print("error, ascii error probably")

    # create_message(sender, to, subject, message_text):
    email_message = create_message("harry@opendesk.cc", "internal@opendesk.cc", "Who's on holiday?", email_string)

    send_message(email_service, "harry@opendesk.cc", email_message)

if __name__ == '__main__':
    main()