import os
import base64
from email.mime.text import MIMEText
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/gmail.send']

def get_gmail_service():
    """จัดการสิทธิ์และสร้างการเชื่อมต่อกับ Gmail API"""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
        
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
            
    return build('gmail', 'v1', credentials=creds)

def send_ncr_alert(to_email, subject, body_text):
    """สั่งส่งอีเมลผ่าน HTTPS (Port 443) บายพาส Firewall บริษัท"""
    try:
        service = get_gmail_service()
        message = MIMEText(body_text, 'plain', 'utf-8')
        message['to'] = to_email
        message['subject'] = subject
        
        raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')
        service.users().messages().send(userId='me', body={'raw': raw_message}).execute()
        return True, "ส่งอีเมลแจ้งเตือนเรียบร้อยแล้ว!"
    except Exception as e:
        return False, f"เกิดข้อผิดพลาด: {str(e)}"