from __future__ import print_function
import smtplib, ssl
import datetime
import os.path
import sys
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
# this is MIT Runs spreadsheet
SAMPLE_SPREADSHEET_ID = '129RJkESc7-t2A2n1IHKkl_b2M_8oGHJML_0baoG_ooA'
SAMPLE_RANGE_NAME = 'Sheet1'

#this is MIT Routes spreadsheet
SAMPLE_SPREADSHEET_ID_2 = '1lXqu7hjugLT6beK-riPsy3DWl2HVdJ1knXqMO21tdu8'
SAMPLE_RANGE_NAME_2 = 'Sheet1'


port = 465  # For SSL
smtp_server = "smtp.gmail.com"
sender_email = "mitclubrunning@gmail.com"  # Enter your address
receiver_email = "mit-running-club@mit.edu"  # Enter receiver address
#receiver_email = "kmorey@mit.edu"
#password 

"""Shows basic usage of the Sheets API.
Prints values from a sample spreadsheet.
"""
creds = None
# The file token.json stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
if os.path.exists('token.json'):
    creds = Credentials.from_authorized_user_file('token.json', SCOPES)
# If there are no (valid) credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(
            'credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.json', 'w') as token:
        token.write(creds.to_json())

try:
    service = build('sheets', 'v4', credentials=creds)

    # Call the Sheets API
    sheet = service.spreadsheets()
    #gets values of the MIT Runs spreadsheet
    result = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID,
                                range=SAMPLE_RANGE_NAME).execute()
    #gets values of the MIT Routes spreadsheet
    result2 = sheet.values().get(spreadsheetId=SAMPLE_SPREADSHEET_ID_2,
                                range=SAMPLE_RANGE_NAME_2).execute()

    #allows lookup of various routes from the Runs spreadsheet to the routes spreadsheet
    dictionary = result2.get('values', [])

    #data on all the runs as stored in the MIT Runs spreadsheet
    values = result.get('values', [])

    #if there is no data found in the MIT Runs spreadsheet
    if not values:
        sys.exit("No data found.")

    
    #the local time in Boston, which is different then server UTC
    toddt = datetime.datetime.now() - datetime.timedelta(hours=4)
    #gets the day in Boston
    tod = toddt.date()
    oneday = datetime.timedelta(days =1)
    send_emails = []

    for row in values[1:]:
        # Print columns A and E, which correspond to indices 0 and 4.
        #print('%s, %s' % (row[0], row[2]))
        date_object = datetime.datetime.strptime(row[0], "%m/%d/%Y")
        #only send emails if its the day before a declared run
        if(date_object.date()-tod == oneday):
            send_emails.append(row)
    
    #sends a new email for each run in the spreadsheet
    for line in send_emails:
        #gets data from the MIT Run spreadsheet data
        date, time, day_week, isemailsent, hosting, route = line
        #only sends email if instructed to
        if (isemailsent !='Yes'):
            continue

        
        message = MIMEMultipart("alternative")
        if hosting == 'MITRC':
            for exroute in dictionary:
                if route == exroute[0]:
                    break
            
            options = exroute[2:]
            noarticle = exroute[0]
            article = exroute[1]

            text_string = ""
            html_string = ""
            num_options = len(options)//3
            
            for k in range(num_options):
                distance, img_link, strava_link = options[k*3:(k+1)*3]
                if k == num_options-1:
                    text_string += f" and {distance} options, as shown below."
                    html_string += f' and <a href="{strava_link}">{distance}</a> options, as shown below.'
                else:
                    text_string += f"{distance}, "
                    html_string += f'<a href="{strava_link}">{distance}</a>, '
            
            text_string+= '\n\n Hope to see you there tomorrow!\n\nHappy Running,\nRunning Club Exec'
            html_string+='<br><br>Hope to see you there tomorrow!<br><br>Happy Running,<br>Running Club Exec</p>'


            for k in range(num_options):
                distance, img_link, strava_link = options[k*3:(k+1)*3]
                html_string += f'<br><p><img src="{img_link}" alt="Picture" width="380" height="479"></p>'
            
            html_string+='</body></html>'
            message["Subject"] = f"{noarticle} Run Tomorrow at {time}"
    
            text = f"""\
            Hello,

            We will be running tomorrow at {time} tomorrow, meeting in front of the Z-center on the Kresge side. Tomorrow's route will be going to {article}!

            There are """+text_string
            html = f"""\
            <html>
              <body>
                <p>Hello,<br><br>
                   We will be running tomorrow at {time} tomorrow, meeting in front of the Z-center on the Kresge side. Tomorrow's route will be going to {article}!<br><br>
                   There are """+html_string

        if hosting == 'Tracksmith':
            message["Subject"] = f"{hosting} Run Tomorrow at {time}"
            text = """\
                    Hello,
                    
                    We will be doing a Tracksmith workout tomorrow, meeting at 6:15 pm in front of the Z-center and at 6:30 at the Tracksmith Shop on 285 Newbury Street. For those of you who don't know, Tracksmith is a running shop that organizes group workouts on Wednesdays, for anyone in the Boston community to join. To be able to enjoy the workout, its generally reccomended that you be able to run 4-5 miles without stopping. Here is the Strava page for Tracksmith, which has the details of upcoming Tracksmith events including their workouts. Hope to see you there tomorrow!
                    
                    Happy Running,,
                    Running Club Exec""" 
            html = """\
            <html>
              <body>
                <p>Hello,<br><br>
                    We will be doing a Tracksmith workout tomorrow, meeting at 6:15 pm in front of the Z-center and at 6:30 at the Tracksmith Shop on 285 Newbury Street. For those of you who don't know, Tracksmith is a running shop that organizes group workouts on Wednesdays, for anyone in the Boston community to join. <b>To be able to enjoy the workout, its generally reccomended that you be able to run 4-5 miles without stopping</b>. Here is the <a href = "https://www.strava.com/clubs/176890">Strava page</a> for Tracksmith, which has the details of upcoming Tracksmith events including their workouts. Hope to see you there tomorrow!<br><br>
                    Happy Running,<br>
                    Running Club Exec
        </p>
        </body>
        </html>
        """
        message["From"] = sender_email
        message["To"] = receiver_email

        # Create the plain-text and HTML version of your message
    
        part1 = MIMEText(text, "plain")
        part2 = MIMEText(html, "html")

        # Add HTML/plain-text parts to MIMEMultipart message
        # The email client will try to render the last part first
        message.attach(part1)
        message.attach(part2)


        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
            server.login(sender_email, password)
            server.sendmail(sender_email, receiver_email, message.as_string())

except HttpError as err:
    print(err)
