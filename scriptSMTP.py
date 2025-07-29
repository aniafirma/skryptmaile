from email.mime.multipart import MIMEMultipart
from ldap3 import Server, Connection, ALL
from datetime import date, datetime,timedelta,timezone
import winfiletime
import sys
import smtplib
from email.mime.text import MIMEText
from cryptography.fernet import Fernet
import os
from dotenv import load_dotenv

load_dotenv()

max_days=90

server = Server(os.getenv('AD_SERVER'))
username = os.getenv('AD_USERNAME')

#AD password
pathAD=os.getenv('AD_PATH')
with open(pathAD, "rb") as key_file:
    key = key_file.read()

f = Fernet(key)
encrypted_password = os.getenv("AD_PASSWORD").encode()
password = f.decrypt(encrypted_password).decode()

BASE_DN = os.getenv('BASE_DN')

SMTPport=(int)(os.getenv('SMTP_PORT'))
TLS = os.getenv("TLS", "False").lower() == "true"
SMTPserver = os.getenv('SMTP_SERVER')
sender = os.getenv('SMTP_SENDER')
SMTPusername = os.getenv('SMTP_USERNAME')

#SMTP password
pathSMTP=os.getenv('SMTP_PATH')
with open(pathSMTP, "rb") as key_file:
    key = key_file.read()

f = Fernet(key)
encrypted_password = os.getenv("SMTP_PASSWORD").encode()
SMTPpassword = f.decrypt(encrypted_password).decode()

BASE_DN = os.getenv('BASE_DN')

def send_notification(recipient,message):
    try:
        part1=MIMEText(message,"plain")

        html_content = f"""\
               <html>
               <head>
                 <style>
                   body {{
                     font-family: Arial, sans-serif;
                     background-color: #f4f4f4;
                     margin: 0;
                     padding: 0;
                   }}
                   .container {{
                     background-color: #ffffff;
                     max-width: 600px;
                     margin: 30px auto;
                     padding: 20px;
                     border-radius: 8px;
                     box-shadow: 0 2px 8px rgba(0, 0, 0, 0.1);
                   }}
                   h2 {{
                     color: #2c3e50;
                     text-align: center;
                   }}
                   p {{
                     color: #333333;
                     line-height: 1.6;
                   }}
                   .footer {{
                     margin-top: 30px;
                     font-size: 12px;
                     color: #999999;
                     text-align: center;
                   }}
                 </style>
               </head>
               <body>
                 <div class="container">
                   <h2> Twoje hasło wkrótce wygaśnie</h2>
                   <p>{message.replace('\n', '<br>')}</p>
                   <div class="footer">
                     <p>To jest automatyczna wiadomość. Prosimy na nią nie odpowiadać. W przypadku dalszych pytań prosimy o kontakt z działem IT</p>
                   </div>
                 </div>
               </body>
               </html>
               """

        part2=MIMEText(html_content,"html")

        msg = MIMEMultipart("alternative")

        msg['Subject'] = 'Hasło wkrótce wygaśnie'
        msg['From'] = sender
        msg.attach(part1)
        msg.attach(part2)
        print(f"Connecting to SMTP server: {SMTPserver}:{SMTPport}, TLS: {TLS}")

        if TLS:
            conn = smtplib.SMTP(SMTPserver, SMTPport)
            conn.ehlo()
            conn.starttls()
            conn.ehlo()
        else:
            conn = smtplib.SMTP_SSL(SMTPserver, SMTPport)

        conn.set_debuglevel(False)
        conn.login(SMTPusername, SMTPpassword)
        try:
            conn.sendmail(sender, recipient, msg.as_string())
            print('email sent')
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
        mail = entry['mail'].value

        if not mail:
            print(f"User has no email : {entry['distinguishedName']}")
            continue

        current_date = datetime.now(timezone.utc)
        when_set = when_set.astimezone(timezone.utc)

        days_since_change = (current_date - when_set).days
        days_until_expiry = max_days - days_since_change

        if days_until_expiry == 1: #powiadamiamy dzień przed
            send_notification(mail,"Twoje hasło do systemu wygaśnie za 1 dzień.\n"
            "Prosimy o jego niezwłoczną zmianę, aby uniknąć zablokowania konta.\n\n"
            "Instrukcja zmiany hasła:\n"
            "- Połącz się z firmową siecią przez VPN\n"
            "- Naciśnij Ctrl + Alt + Delete\n"
            "- Wybierz opcję 'Zmień hasło'\n"
            "- Wpisz stare hasło oraz nowe hasło dwukrotnie\n\n"
            "Wymagania dla nowego hasła:\n" 
            "- Minimum 8 znaków\n"
            "- Duże i małe litery\n"
            "- Cyfry\n"
            "- Znaki specjalne\n"
            "- Hasło nie może być wcześniej używane")

        elif days_until_expiry == 7: #powiadamiamy tydzien przed
            send_notification(mail,"Twoje hasło do systemu wygaśnie za 1 dzień.\n"
            "Prosimy o jego niezwłoczną zmianę, aby uniknąć zablokowania konta.\n\n"
            "Instrukcja zmiany hasła:\n"
            "- Połącz się z firmową siecią przez VPN\n"
            "- Naciśnij Ctrl + Alt + Delete\n"
            "- Wybierz opcję 'Zmień hasło'\n"
            "- Wpisz stare hasło oraz nowe hasło dwukrotnie\n\n"
            "Wymagania dla nowego hasła:\n" 
            "- Minimum 8 znaków\n"
            "- Duże i małe litery\n"
            "- Cyfry\n"
            "- Znaki specjalne\n"
            "- Hasło nie może być wcześniej używane")

except Exception as e:
    print(f"Connection to AD server failed :( - {e}")

