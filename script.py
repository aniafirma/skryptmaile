import ldap3
from ldap3 import Server, Connection, ALL
from datetime import date, datetime

#https://dnmtechs.com/python-3-ldap-authentication-with-active-directory/
#https://ldap3.readthedocs.io/en/latest/
#https://learn.microsoft.com/en-us/windows/win32/ad/user-object-attributes

max_days=90

server = Server('ip AD??')
username = 'nie wiem'
password = 'tez nie wiem'
BASE_DN = 'DC=example,DC=com'

def send_notification():
    print('notifying user')

try:
    conn = Connection(server, user=username, password=password)
    conn.bind()
    conn.search(search_base=BASE_DN, search_filter='(objectclass=user)', attributes=['mail','pwdLastSet'])

    for entry in conn.entries:
        when_set=entry['pwdLastSet'].value
        mail=entry['mail'].value
        current_date = datetime.now()
        age=current_date-when_set #windows file time jak konwertowac

        if age>max_days-7: #powiadamiamy tydzien przed
            send_notification()

except Exception as e:
    print(f"Connection failed:(+ {e}")
