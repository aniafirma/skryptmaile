import ldap3
from ldap3 import Server, Connection, ALL
from datetime import date, datetime,timedelta
import winfiletime
import win32com.client as win32

#https://dnmtechs.com/python-3-ldap-authentication-with-active-directory/
#https://ldap3.readthedocs.io/en/latest/
#https://learn.microsoft.com/en-us/windows/win32/ad/user-object-attributes
#https://github.com/jleclanche/winfiletime
#https://stackoverflow.com/questions/6332577/send-outlook-email-via-python
#https://stackoverflow.com/questions/24192252/python-sending-outlook-email-from-different-address-using-pywin32

max_days=90

server = Server('ip AD??')
username = 'nie wiem'
password = 'tez nie wiem'
BASE_DN = 'DC=example,DC=com'

def send_notification(recipient): #wysyla sie z konta osoby ktora jest zalogowana na tym komputerze
    outlook = win32.Dispatch('outlook.application')
    mail=outlook.CreateItem(0)
    mail.To=recipient
    mail.Subject='Hasło zaraz wygaśnie'
    mail.Body='Dzień dobry proszę zmienić hasło bo zaraz wygaśnie'

    From = None
    for myEmailAddress in outlook.Session.Accounts:
        if "anna.smuga" in str(myEmailAddress):
            From = myEmailAddress
            break

    if From != None:
        # This line basically calls the "mail.SendUsingAccount = xyz@email.com" outlook VBA command
        mail._oleobj_.Invoke(*(64209, 0, 8, 0, From))

        mail.Send()

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

        if age>timedelta(days=max_days-7): #powiadamiamy tydzien przed
            send_notification(mail)

except Exception as e:
    print(f"Connection failed:(+ {e}")
