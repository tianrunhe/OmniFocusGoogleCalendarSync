import datetime
import pickle
import os.path
import iso8601
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/calendar']


def main(arguments):
    import rfc3339

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

    service = build('calendar', 'v3', credentials=creds)

    event = {
        'summary': arguments.title,
        'start': {
            'dateTime': rfc3339.rfc3339(arguments.starttime)
        },
        'end': {
            'dateTime': rfc3339.rfc3339(arguments.endtime)
        }
    }
    event = service.events().insert(calendarId=arguments.calendarId, body=event).execute()
    print('Event created: %s' % (event.get('htmlLink')))

    now = datetime.datetime.utcnow().isoformat() + 'Z'  # 'Z' indicates UTC time
    events_result = service.events().list(calendarId=arguments.calendarId,
                                          timeMin=now,
                                          maxResults=10, singleEvents=True,
                                          orderBy='startTime').execute()
    events = events_result.get('items', [])
    if not events:
        print('No upcoming events found.')
    for event in events:
        start = event['start'].get('dateTime', event['start'].get('date'))
        print(start, event['summary'])


def valid_date(s):
    try:
        return iso8601.parse_date(s)
    except ValueError:
        msg = "Not a valid date: '{0}'.".format(s)
        raise argparse.ArgumentTypeError(msg)


if __name__ == '__main__':
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("-i", "--calendarId", help="Calendar Id", required=False, type=str, default="primary")
    parser.add_argument("-t", "--title", help="Title of the event", required=True, type=str)
    parser.add_argument("-s", "--starttime",
                        help="Start datetime of the event - format e.g., 2020-04-23T09:00:00-07:00",
                        required=True,
                        type=valid_date)
    parser.add_argument("-e", "--endtime",
                        help="End datetime of the event - format e.g., 2020-04-23T09:00:00-07:00",
                        required=True,
                        type=valid_date)
    args = parser.parse_args()

    main(args)
