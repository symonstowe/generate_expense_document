from __future__ import print_function

import os.path
import base64
import time
import datetime
import numpy as np
from pathlib import Path
from sys import platform
import subprocess

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

class ExpenseDocSetup ():
    def __init__(self):
        self.company = 'Tidal'
        self.company_full = 'Tidal Medical Inc.'
        self.doc_ID = 'TM-Ex-001'
        self.start_date = datetime.date(2022, 6, 12)
        self.end_date = datetime.date(2022,6,17)
        self.user = 'Symon Stowe'
        self.user_email = 'symonstowe@gmail.com'
        self.user_company = None
        self.base_path = "./output/" + str(self.end_date.year)+ "/" + str(self.start_date) + "--" + str(self.end_date)
        self.im_path = self.base_path + "/imgs/"
        Path(self.im_path).mkdir(parents=True, exist_ok=True) # Create directory if not already created

class GmailConnect ():
    def __init__(self):
        creds = None
        # If modifying these scopes, delete the file token.json.
        SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']
        # Look for saved credentials
        if os.path.exists('token.json'):
            creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        # Login online and save credentials
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
        self.creds = creds

def save_data_from_part(part,msg,service,doc_setup):
    if 'data' in part['body']:
        data = part['body']['data']
    else:
        att_id = part['body']['attachmentId']
        att = service.users().messages().attachments().get(userId="me", messageId=msg['id'],id=att_id).execute()
        data = att['data']
    file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
    path = part['filename']
    with open(doc_setup.im_path + path, 'wb') as f:
        f.write(file_data)
    path = doc_setup.im_path + path
    return  path

def sort_by_date(dates,captions,imgs):
    # Sort everything that we are interested in printing to the tex document by date
    idx = np.argsort(dates)
    dates = [dates[i] for i in idx]
    captions = [captions[i] for i in idx]
    imgs = [imgs[i] for i in idx]
    return dates, captions, imgs 

def latex_output(doc_setup,dates,captions,imgs):
    i=0
    summary_table = ""
    # TODO - parse the email to get the amounts and sum them from the email
    dates, captions, imgs = sort_by_date(dates, captions, imgs)
    days = [time.strftime('%Y-%m-%d', time.localtime(int(date)/1000)) for date in dates]
    for i, caption in enumerate(captions):
        summary_table += days[i] + " & " + caption + " & \$\,XX.xx \\\ "

    # TODO add a sum to the table!
    # If dollar sign in email do.... else print the Xs (gross)
    # TODO - add href to the corresponding receipt....
    
    with open('tex_template.tex', 'r') as file:
        filedata = file.read()

    # Replace the target strings
    filedata = filedata.replace('@EXPENSE_SUMMARY', summary_table)
    filedata = filedata.replace('@COMPANY_ADDRESS', doc_setup.company)
    filedata = filedata.replace('@COMPANY', doc_setup.company_full)
    filedata = filedata.replace('@USER_NAME', doc_setup.user)
    filedata = filedata.replace('@USER_EMAIL', doc_setup.user_email)
    filedata = filedata.replace('@DATE_RANGE', str(doc_setup.start_date) + " - " + str(doc_setup.end_date))
    filedata = filedata.replace('@FILE_ID', doc_setup.doc_ID)
    filedata = filedata.replace('@DATE', str(datetime.date.today()))
    filedata = filedata.replace('@IMG_LOC', doc_setup.im_path[2:])
    filedata = filedata.replace('@LOGO_LOC', 'personal_imgs/')

    receipt_section = ""
    img_c = 0
    for i, date in enumerate(dates): # For each email arrival date put all the images in 
        day = time.strftime('%Y-%m-%d', time.localtime(int(date)/1000))
        longform_day = time.strftime('%B %d %Y', time.localtime(int(date)/1000))
        receipt_section += "\section*{" + longform_day + "} \n"
        atchd_imgs = imgs[i]
        for img in atchd_imgs:
            img_c += 1
            # Pull in the image template
            with open('img_template.tex', 'r') as file:
                img_tex = file.read()
            img_tex = img_tex.replace('@IMG_NAME', img.split('/')[-1])
            img_tex = img_tex.replace('@EXPENSE_INFO', captions[i])
            receipt_section += img_tex + "\n"
            if img_c%2 == 0:
                receipt_section += "\\newpage \n"
                
    filedata = filedata.replace('@RECEIPTS', receipt_section)

    fname = doc_setup.base_path + "/" + str(doc_setup.start_date) + "--" + str(doc_setup.end_date) + ".tex"
    doc_setup.fname = fname.split('/')[-1]
    if os.path.isfile(fname):
        overwrite = input('File for this date range already exists... Overwrite? \ny = yes, \nn = no \nr = recompile pdf but keep tex\n')
        if overwrite.lower() == 'y':
            with open(fname, 'w') as file:
                file.write(filedata)
        elif overwrite.lower() == 'r':
            mk_pdf_from_tex(doc_setup)
        else:
            raise ValueError('File exists - not overwriting.')
    else:
        with open(fname, 'w') as file:
            file.write(filedata)
    return doc_setup
    
def mk_pdf_from_tex(doc_setup):
    if platform == "linux" or platform == "linux2":
        # linux
        try: # Run twice for the page numbering...
            bashCommand = "lualatex " + doc_setup.fname
            process = subprocess.Popen(bashCommand.split(), stdout=subprocess.PIPE,cwd=doc_setup.base_path)
            output, error = process.communicate()
            process = subprocess.Popen(bashCommand.split(), stdout=subprocess.PIPE,cwd=doc_setup.base_path)
            output, error = process.communicate()
        except ValueError:
            print("Bash command did not work.")

    elif platform == "darwin":
        # OS X
        raise ValueError("No Tex compilation set up for Mac - you must compile manualy.")

    elif platform == "win32":
        # Windows...
        raise ValueError("No Tex compilation set up for Windows - you must compile manualy.")


if __name__ == '__main__':
    g_api = GmailConnect()
    doc_setup = ExpenseDocSetup()
    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=g_api.creds)
        #results = service.users().labels().list(userId='me').execute()
        results = service.users().messages().list(userId='me').execute()
        message_ids = results.get('messages',[])
        msg_data = [service.users().messages().get(userId="me", id=msg_id['id'], format="full", metadataHeaders=None).execute() for msg_id in message_ids] 
        c = 0
        t_rec = [0] * len(msg_data)
        sbjct = [0] * len(msg_data)
        content = [0] * len(msg_data)
        attch_pths = [0] * len(msg_data)
        for msg in msg_data:
            t_rec[c] = msg['internalDate'] 
            headers = msg['payload']['headers']
            sbjct[c]= [i['value'] for i in headers if i["name"]=="Subject"]
            if doc_setup.company in sbjct[c][0]: # Begin expense entry
                content[c] = msg['snippet']
                # Stupid nested parts...
                img_pths = []
                for part in msg['payload']['parts']:
                    if part['filename']:
                        path = save_data_from_part(part,msg,service,doc_setup)
                        img_pths.append(path)

                    elif 'parts' in part.keys():
                        for pt in part['parts']:
                            if pt['filename']:
                                path = save_data_from_part(pt,msg,service,doc_setup)
                                img_pths.append(path)
            attch_pths[c] = img_pths
            c+=1
        latex_output(doc_setup,t_rec,content,attch_pths)
        mk_pdf_from_tex(doc_setup)

    except HttpError as error:
        # TODO(developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')
