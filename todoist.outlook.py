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
    tasks = todo_items.Restrict("[Complete] = FALSE")

    if (len(tasks) == 0):
        logger.info("No flagged emails found.")
    else:
        logger.info("{0} flagged emails will be processed.".format(len(tasks)))

    for task in tasks:
        try:
            creation_date = pyWinDate2datetime(task.CreationTime)

            # Do not create reminders for emails that came in today.
            if not creation_date.date() < datetime.today().date():
                print("Skipped, since it was created today.")
                continue

        except AttributeError as ex:
            logger.exception("Exception occured while processing the message.", ex)
            message = todo_items.GetNext()
            continue

        # How Todoist generates an ID for an Outlook email:
        # https://todoist.com/Support/show/30790/

        # How to get PR_INTERNET_MESSAGE_ID for an Outlook item:
        # http://www.slipstick.com/developer/read-mapi-properties-exposed-outlooks-object-model/
        PR_INTERNET_MESSAGE_ID = task.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001E")
        todoist_message_id = base64.b64encode(bytes("id={0};mid={1}".format(task.EntryID, PR_INTERNET_MESSAGE_ID), "utf-8")).decode("utf-8")

        # Add a task.
        item = api.items.add(u'[[outlook=id3={0}, Review Email: {1}]]'.format(todoist_message_id, task.Subject), config["TODOIST_PROJECT_ID_OUTLOOK"], date_string="")
        logger.info("Processing '{0}'.".format(task.Subject.encode("utf-8")))
        r = api.commit()

        # If error occures, move on.
        if "error_code" in r:
            logger.info ("Task for email '{0}' was not created. Error: {1}:{2}".format(task.Subject, r["error_code"], r["error_string"]))
            continue

        # Add a note.
        note = api.notes.add(item["id"], u'Automatically Generated Task by Todoist Task Creator Script')
        r = api.commit()

        # Mark message as read.
        task.TaskCompletedDate = task.CreationTime
        task.Save()

if __name__=="__main__":
    main()
