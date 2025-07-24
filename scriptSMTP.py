from ldap3 import Server, Connection, ALL
from datetime import date, datetime,timedelta
import winfiletime
import sys
from smtplib import SMTP_SSL as SMTP
from email.mime.text import MIMEText

#https://dnmtechs.com/python-3-ldap-authentication-with-active-directory/
#https://ldap3.readthedocs.io/en/latest/
#https://learn.microsoft.com/en-us/windows/win32/ad/user-object-attributes
#https://github.com/jleclanche/winfiletime
#https://stackoverflow.com/questions/6332577/send-outlook-email-via-python
#https://stackoverflow.com/questions/24192252/python-sending-outlook-email-from-different-address-using-pywin32
#https://stackoverflow.com/questions/64505/sending-mail-from-python-using-smtp

max_days=90

server = Server('ip AD??')
username = 'nie wiem'
password = 'tez nie wiem'
BASE_DN = 'nie wiem'

def send_notification(recipient,message):
    SMTPserver = 'nasz serwer'
    sender = 'mail'

    USERNAME = "USER_NAME_FOR_INTERNET_SERVICE_PROVIDER"
    PASSWORD = "PASSWORD_INTERNET_SERVICE_PROVIDER"
    text_subtype = 'plain'
    try:
        msg = MIMEText(message, text_subtype)
        msg['Subject'] = 'Hasło wkrótce wygaśnie'
        msg['From'] = sender
        conn = SMTP(SMTPserver)
        conn.set_debuglevel(False)
        conn.login(USERNAME, PASSWORD)
        try:
            conn.sendmail(sender, recipient, msg.as_string())
        finally:
            conn.quit()
    except Exception as e:
        sys.exit(f"mail failed {e}")

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

        if age>timedelta(days=max_days-1): #powiadamiamy tydzien przed
            send_notification(mail,'Za 1 dzień haslo wygasa. Proszę je pilnie zmienić albo konto zostanie zablokowane!')

        elif age>timedelta(days=max_days-7): #powiadamiamy tydzien przed
            send_notification(mail,'Za 7 dni hasło wygasa. Proszę je zmienić')

except Exception as e:
    print(f"Connection failed :( - {e}")
