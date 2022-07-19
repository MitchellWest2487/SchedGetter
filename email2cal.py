import base64
from datetime import datetime
import os.path
import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import json

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
MSG_SEARCH = 'from:timberlineliquorstore@gmail.com has:attachment filename:xlsx'
TOKEN = 'token.json'
MSG_ID = 'msgId.txt'
SCHED_DATA = 'data.json'

def create_google_service():
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.

    creds = None
    if os.path.exists(TOKEN):
        creds = Credentials.from_authorized_user_file(TOKEN, SCOPES)
        
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(TOKEN, 'w') as token:
            token.write(creds.to_json())


    return build('gmail', 'v1', credentials=creds)

#checks if the msg_id is the same as previous run and returns true if it is and false if it has changed
def is_old_sched(new_id):
    if os.path.exists(MSG_ID):
        with open(MSG_ID, 'r') as file:
            old_id = file.read()
            return True if old_id == new_id else False
        
#gets the excel schedule from the email and returns it as a panda df   
def get_sched_from_email(email_service, msg_id):
    msg = email_service.users().messages().get(userId='me', id=msg_id).execute()
    attachment_id = msg['payload']['parts'][1]['body']['attachmentId']
    attachment = email_service.users().messages().attachments().get(userId='me',messageId = msg_id, id=attachment_id).execute()
    raw_data = base64.urlsafe_b64decode(attachment['data'].encode('UTF-8'))
    return pd.read_excel(raw_data, header=1)

#cleans the schedule df and returns it as a dictionary
def clean_sched(sched_df):
    #keys for the df to exclude the unnamed col
    keys = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
    sched_dic = {}
    dates_one_week = {}
    print(sched_df.keys())
    
    for _, row in sched_df.iterrows():
        
        #retrieves the scheduled dates to be worked that week and adds them as keys in the schedule dictionary 
        if isinstance(row['Monday'], datetime):
            for key in keys:
                sched_dic[str(row[key].date())] = {}
                dates_one_week[key] = str(row[key].date())
        
        #skips rows that are empty and the start of the next weeks schedule if there are multiple in one file
        if pd.isnull(row['Name']) or row['Name'] == 'Name': continue
        
        #maps the name and hours to be worked to the date
        name = row['Name']
        uname_count = 2
        for key in keys:
            
            #adds people to only the days they work
            if not pd.isnull(row[key]):
                time_in = str(row[key]) 
                time_out = str(row[f'Unnamed: {uname_count}'])
                date = dates_one_week[key]
                sched_dic[date][name] = {'in': time_in, 'out': time_out}
            uname_count = uname_count + 2
            
    return sched_dic

#returns a dictionary of the most recent work schedule
def get_sched():
    email_service = create_google_service()
    msgs = email_service.users().messages().list(userId='me', q = MSG_SEARCH).execute()
    msg_id = msgs['messages'][0]['id']

    if is_old_sched(msg_id) and os.path.exists(SCHED_DATA):
        with open(SCHED_DATA, 'r') as file:
            return json.load(file)
    
    schedule = get_sched_from_email(email_service, msg_id)
    clean_schedule = clean_sched(schedule)
    
    #writes the new msg_id to the file
    with open(MSG_ID, 'w') as file:
        file.write(msg_id)
    
    #converts the new schedule to json and writes it to the file
    with open(SCHED_DATA, 'w') as file:
        json.dump(clean_schedule,file, indent=2)

    return clean_schedule


schedule = get_sched()
print(schedule.keys())