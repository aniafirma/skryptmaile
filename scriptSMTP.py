from ldap3 import Server, Connection, ALL
from datetime import date, datetime,timedelta
from dotenv import load_dotenv
import os
import winfiletime
import sys
from smtplib import SMTP_SSL as SMTP
from email.mime.text import MIMEText

load_dotenv()
max_days=90

server = Server(os.getenv('AD_SERVER'))
username = os.getenv('AD_USERNAME')
password = os.getenv('AD_PASSWORD')
BASE_DN = 'nie wiem'

SMTPserver = os.getenv('SMTP_SERVER')
sender = os.getenv('SMTP_SENDER')
SMTPusername = os.getenv('SMTP_USERNAME')
SMTPpassword = os.getenv('SMTP_PASSWORD')
text_subtype = 'plain'

def send_notification(recipient,message):
    try:
        msg = MIMEText(message, text_subtype)
        msg['Subject'] = 'Hasło wkrótce wygaśnie'
        msg['From'] = sender
        conn = SMTP(SMTPserver)
        conn.set_debuglevel(False)
        conn.login(SMTPusername, SMTPpassword)
        try:
            conn.sendmail(sender, recipient, msg.as_string())
        finally:
            conn.quit()
    except Exception as e:
        sys.exit(f"Sending email failed - {e}")

try:
    conn = Connection(server, user=username, password=password)
    conn.bind()
    conn.search(search_base=BASE_DN, search_filter='(objectclass=user)', attributes=['mail','pwdLastSet'])

    for entry in conn.entries:
        date_changed=entry['pwdLastSet'].value
        when_set=winfiletime.to_datetime(date_changed)
        mail=entry['mail'].value
        current_date = datetime.now()

        age=current_date-when_set #zwraca timedelta

        if age>timedelta(days=max_days-1): #powiadamiamy dzień przed
            send_notification(mail,'Za 1 dzień haslo wygasa. Proszę je pilnie zmienić albo konto zostanie zablokowane!')

        elif age>timedelta(days=max_days-7): #powiadamiamy tydzien przed
            send_notification(mail,'Za 7 dni hasło wygasa. Proszę je zmienić')

except Exception as e:
    print(f"Connection to AD server failed :( - {e}")
