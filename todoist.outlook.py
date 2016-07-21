import os
import sys
import runpy
import base64
import win32com.client
import todoist
import logging
import logging.handlers
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import pywintypes

MY_LOCATION = os.path.dirname(os.path.realpath(__file__))
LOG_FILENAME = os.path.join(MY_LOCATION, 'todoist.outlook.log')

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

def pyWinDate2datetime(now_pytime):
    """Converts pyWinDate to Python datetime. More info:
    http://timgolden.me.uk/python/win32_how_do_i/use-a-pytime-value.html"""

    now_datetime = datetime (
      year=now_pytime.year,
      month=now_pytime.month,
      day=now_pytime.day,
      hour=now_pytime.hour,
      minute=now_pytime.minute,
      second=now_pytime.second
    )
    
    return now_datetime

def main():

    # Get config.
    config = runpy.run_path(os.path.join(MY_LOCATION, "todoist.conf"))

    # Init Todoist API.
    api = todoist.TodoistAPI(config["TODOIST_API_TOKEN"])
    api.sync()

    # For some background on Outlook Interop model, see:
    # https://msdn.microsoft.com/en-us/library/office/ff861868%28v=office.15%29.aspx
    # https://msdn.microsoft.com/en-us/library/microsoft.office.interop.outlook.mailitem_properties(v=office.14).aspx
    olFolderTodo = 28
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    todo_folder = ns.GetDefaultFolder(olFolderTodo)
    todo_items = todo_folder.Items

    message = todo_items.GetFirst()
    total = 0

    # Iterate through messages.
    while message:
        try:        
            start_date = pyWinDate2datetime(message.TaskStartDate)
            end_date = pyWinDate2datetime(message.TaskCompletedDate)

            # Tasks that are not completed in Outlook receive a completion date
            # in the future something like 4501-01-01.
            difference_in_years = relativedelta(end_date, start_date).years

            # If the completion date is recent, then it's actually a completed
            # task and we don't need to generate a task item for it.
            if (difference_in_years < 10):
                message = todo_items.GetNext()
                continue

        except AttributeError as ex:
            message = todo_items.GetNext()
            continue  
     
        # How Todoist generates an ID for an Outlook email:
        # https://todoist.com/Support/show/30790/

        # How to get PR_INTERNET_MESSAGE_ID for an Outlook item:
        # http://www.slipstick.com/developer/read-mapi-properties-exposed-outlooks-object-model/           
        PR_INTERNET_MESSAGE_ID = message.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001E")
        todoist_message_id = base64.b64encode(bytes("id={0};mid={1}".format(message.EntryID, PR_INTERNET_MESSAGE_ID), "utf-8")).decode("utf-8")
        
        # Add a task.
        item = api.items.add(u'[[outlook=id3={0}, Review Email: {1}]]'.format(todoist_message_id, message.Subject), config["TODOIST_PROJECT_ID_OUTLOOK"], date_string="today")
        logger.info("Processing '{0}'.".format(message.Subject.encode("utf-8")))
        r = api.commit()

        # If error occures, move on.
        if "error_code" in r:
            logger.info ("Task for email '{0}' was not created. Error: {1}:{2}".format(message.Subject, r["error_code"], r["error_string"]))
            message = todo_items.GetNext()
            continue

        # Add a note.
        note = api.notes.add(item["id"], u'Automatically Generated Task by Todoist Task Creator Script')
        r = api.commit()

        # Mark message as read.
        message.TaskCompletedDate = message.TaskStartDate
        message.Save()

        # Advanced to the next message.
        message = todo_items.GetNext()        
        total += 1

    if (total == 0):
        logger.info("No flagged emails found.")
    else:
        logger.info("{0} flagged emails processed.".format(total))

if __name__=="__main__":
    main()
