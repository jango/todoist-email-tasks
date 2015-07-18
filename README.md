# todoist-email-tasks
Two Python scripts that will generate a task for each starred (Gmail) or flagged (Outlook) email in [Todoist].

## Motivation
I am a big fan of Todoist and use it as my sole task management system. When I review personal or work email often there are messages that require follow up that I can not afford to do right away. I used to flag those emails and eventally add them to Todoist one by one, but every message takes time. I did not want to use forwarding feature of Todoist either, because I do not want the content of my emails to be included in the task's note.

These two scripts are now setup to run every few minutes to create a task for any email that I have flagged and clear the email flag. This way I can continue using Todoist as my primary task management system and avoid doing "double entry". I hope that someone else may find it useful, too.

## Setup - Both Scripts
As a general suggestion, before scheduling either of the scripts, make sure that you are able to run them through the command line.

Start by renaming `todoist.sample.conf` to `todoist.conf` and specify the Todoist API key (can be found under your account settings) and Project IDs (can be found in the URL when you open any of your project folders in the browser window).

After that, install requirements in `requirements.txt` via `pip`.

## Setup - Gmail
In order to access Gmail from python, you need to [generate] `client_secret.json` for your Gmail account and drop it into the project folder (follow the list of steps under `Step 1: Enable the Gmail API` only).

I run the Gmail script on Linux, so the cron setup looks like this for me:
```cron
# Runs Todoist Task Creator.
* * * * * python /opt/scripts/todoist-email-tasks/todoist.gmail.py
```

## Setup - Outlook
Besides dependencies, you may need to install [Outlook Interop Aseembly References].

You can create a Windows Scheduler task to run a batch file that will run the Outlook script:
```sh
python C:\todoist-email-tasks\todoist.outlook.py
```
You can then use this cool [trick] to make the command Window invisible when the script runs.

License
----

MIT

[Todoist]:https://todoist.com
[trick]:http://superuser.com/questions/62525/run-a-batch-file-in-a-completely-hidden-way
[Outlook Interop Aseembly References]:https://www.microsoft.com/en-ca/download/details.aspx?id=3508
[generate]:https://developers.google.com/gmail/api/quickstart/python
