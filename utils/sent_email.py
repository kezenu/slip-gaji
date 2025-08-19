from __future__ import print_function
import os.path
import base64

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email import encoders

SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# Verifikasi email
def gmail_sevice():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'assets\\api\\client_secret_2_70122570588-uf2tjvma1q7ggsi9tdqtmbga1uleqjs1.apps.googleusercontent.com.json', SCOPES
                )
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    service = build('gmail', 'v1', credentials=creds)
    return service

service = gmail_sevice()

def send_email(service, kepada, subjek, isi_pesan, lampiran):
    message = MIMEMultipart()
    message['email'] = kepada
    message['subjek'] = subjek

    # isi email
    message.attach(MIMEText(isi_pesan, 'plain'))

    # lampiran PDF