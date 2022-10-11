#!/usr/bin/env python
# coding: utf-8

# In[1]:


import pickle
import os
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from google.auth.transport.requests import Request


def Create_Service(client_secret_file, api_name, api_version, *scopes):
    print(client_secret_file, api_name, api_version, scopes, sep='-')
    CLIENT_SECRET_FILE = client_secret_file
    API_SERVICE_NAME = api_name
    API_VERSION = api_version
    SCOPES = [scope for scope in scopes[0]]
    print(SCOPES)

    cred = None

    pickle_file = f'token_{API_SERVICE_NAME}_{API_VERSION}.pickle'
    # print(pickle_file)

    if os.path.exists(pickle_file):
        with open(pickle_file, 'rb') as token:
            cred = pickle.load(token)

    if not cred or not cred.valid:
        if cred and cred.expired and cred.refresh_token:
            cred.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET_FILE, SCOPES)
            cred = flow.run_local_server()

        with open(pickle_file, 'wb') as token:
            pickle.dump(cred, token)

    try:
        service = build(API_SERVICE_NAME, API_VERSION, credentials=cred)
        print(API_SERVICE_NAME, 'service created successfully')
        return service
    except Exception as e:
        print('Unable to connect.')
        print(e)
        return None

def convert_to_RFC_datetime(year=1900, month=1, day=1, hour=0, minute=0):
    dt = datetime.datetime(year, month, day, hour, minute, 0).isoformat() + 'Z'
    return dt


# In[2]:



import base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import mimetypes


# In[3]:


CLIENT_SECRET_FILE = 'client_secret_file.json'
API_NAME = 'gmail'
API_VERSION = 'v1'
SCOPES = ['https://mail.google.com/']


# In[4]:


service = Create_Service(CLIENT_SECRET_FILE, API_NAME, API_VERSION, SCOPES)


# In[5]:


def sendEmail(msg_string, recepient, subject, template='html'):
    emailMsg = msg_string
    mimeMessage = MIMEMultipart()
    mimeMessage['to'] = recepient
    mimeMessage['subject'] = subject
    mimeMessage.attach(MIMEText(emailMsg, template))
    raw_string = base64.urlsafe_b64encode(mimeMessage.as_bytes()).decode()
    message = service.users().messages().send(userId='me', body={'raw': raw_string}).execute()
    if message['labelIds']== 'SENT':
        print('Email Successfully Sent')


# In[ ]:





# In[29]:


# HTML message, would use mako templating in real scenario
msg_html = """
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
  <title>html title</title>
  <style type="text/css" media="screen">
.attribution {{ color: #aaaaaa; font-size: 8pt }}
.greeting {{ font-size: 54pt; font-styweight: bold}}

td {
  {text-align: right;}
}

tr:nth-child(even) {
  background-color: #D6EEEE;
}
</style></head>
<body>
<span class="greeting">Hello, {}!</span>
<p>As our valued customer, we would like to invite you to our annual sale!</p>
<p class="attribution">
<a href="https://www.youtube.com/watch?v=JRCJ6RtE3xU&ab_channel=CoreySchafer">
Image by FreeVector.com
</a></p>
<table>
<tr>
    <th>Has Negative Inventory</th>
    <th>Location</th>
    <th>Count</th>
  </tr>
  <tr>
    <td>True</td>
    <td>LA STOCK</td>
    <td>20</td>
  </tr>
</table>
</body></html>
"""


# In[30]:


sendEmail(msg_html, 'paulgathondudev@gmail.com', 'testing html')


# In[ ]:


from credsTest.GoogleSheetsAPI import *


# In[4]:


import pandas as pd

import numpy as np
Url='https://docs.google.com/spreadsheets/d/1EPo-yeHzap_BKWz0VvdOpFsN1Atw4jL2ZA7aXhl1aps/edit#gid=273865484'
gsheetId=getSheetId(Url)
df=pd.DataFrame({'header':[1 , 2, 3, 4, 5]})
writeDataToSheetDf('Sheet7',gsheetId,df)

