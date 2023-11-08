import pyodbc
from django.contrib.auth.hashers import make_password,check_password
import os
os.environ.setdefault("DJANGO_SETTINGS_MODULE", "config.settings")
import django
django.setup()
import bcrypt


def connect():
    server = 'aca-mysqlserver1.database.windows.net'
    database = 'Database1'
    username = 'aca-user1'
    password = 'Pokemon!123'
    port = '1433'

    driver = '{SQL Server}'
    
    cnxn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    return cnxn


def createUser():
    username = 'cumberland'
    password = 'cumb3rland'
    role = 'cumberland'
    
    hashed_password = make_password(password)
    
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute("INSERT INTO [dbo].[User] (Username, Password, Role) VALUES (?, ?, ?)", (username, hashed_password, role))
    cnxn.commit()
    print(username)
if __name__ == "__main__":
    createUser()