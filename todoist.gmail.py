import todoist
import logging
import logging.handlers
from datetime import datetime

import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools

LOG_FILENAME = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'todoist.gmail.log')
SCOPES = ['https://www.googleapis.com/auth/userinfo.email', 'https://www.googleapis.com/auth/gmail.modify']
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Gmail API Quickstart'
USER_ID = 'me'

# Set up a specific logger with our desired output level
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)

# Format.
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# Add the log message handler to the logger
handler = logging.handlers.RotatingFileHandler(LOG_FILENAME, maxBytes=1024*1024, backupCount=5)
handler.setFormatter(formatter)
logger.addHandler(handler)

handler = logging.StreamHandler()
handler.setFormatter(formatter)
logger.addHandler(handler)

def get_credentials(): 
    """Straight out of Google's quickstart example:
        https://developers.google.com/gmail/api/quickstart/python
    
    Gets valid user credentials from storage.

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
                                   'gmail-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatability with Python 2.6
            credentials = tools.run(flow, store)
    return credentials

def main():

    config = {}
    execfile("todoist.conf", config) 

    # Initialize Google credentials.
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())

    # Initialize Todist API.
    api = todoist.TodoistAPI(config["TODOIST_API_TOKEN"])
    api.sync(resource_types=['all'])
    service = discovery.build('gmail', 'v1', http=http)

    # Get all gmail messages that are starred:
    response = service.users().messages().list(userId=USER_ID, q='label:starred').execute()

    messages = []
    if 'messages' in response:
        messages.extend(response['messages'])

    while 'nextPageToken' in response:
      page_token = response['nextPageToken']
      response = service.users().messages().list(userId=USER_ID,
                                                 labelIds=label_ids,
                                                 pageToken=page_token).execute()
      messages.extend(response['messages'])

    if len(messages) == 0:
        logger.info("No starred message(s).")
    else:
        logger.info("{0} starred messages found.".format(len(messages)))
    
    for m in messages:
        message = service.users().messages().get(userId=USER_ID, id=m["id"]).execute()
        
        subject = "No Subject"

        # Get email subject.
        for header in message['payload']['headers']:
            if header["name"] == "Subject":
                subject = header["value"]

        # Create a todoist item.
        item = api.items.add(u'https://mail.google.com/mail/u/0/#inbox/{0} (Review Email: {1})'.format(m["id"], subject), config["TODOIST_PROJECT_ID_GMAIL"], date_string="today")
        logger.info("Processing {0}: {1}.".format(m["id"], subject.encode("utf-8")))
        r = api.commit()

        # Skip this item on error.
        if "error_code" in r:
            logger.info ("Task for email {0} was not created. Error: {1}:{2}".format(m["id"], r["error_code"], r["error_string"]))
            continue

        # A note that the task was auto-generated.
        note = api.notes.add(item["id"], u'Automatically Generated Task by Todoist Task Creator Script')
        r = api.commit()

        # Mark message as read and unstar them.
        thread = service.users().threads().modify(userId=USER_ID, id=m["threadId"], body={'removeLabelIds': ['UNREAD', 'STARRED'], 'addLabelIds': []}).execute()

if __name__ == '__main__':
    main()