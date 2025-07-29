from cryptography.fernet import Fernet

key = Fernet.generate_key()
with open("secret.keySMTP", "wb") as key_file:
    key_file.write(key)

f = Fernet(key)

password = b"your SMTP password"
encrypted = f.encrypt(password)
print(encrypted.decode()) #encrypted password, put in .env file

key = Fernet.generate_key()

with open("secret.keyAD", "wb") as key_file:
    key_file.write(key)

f = Fernet(key)

password = b"your AD password"
encrypted = f.encrypt(password)
print(encrypted.decode()) #encrypted password, put in .env file

#put path to files with keys in env file
