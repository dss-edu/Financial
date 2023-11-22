import pyodbc
from dotenv import load_dotenv
import os
import base64
from cryptography.fernet import Fernet

load_dotenv()

def connect():
    key = base64.b64decode(os.getenv('FERNET_KEY'))
    fernet = Fernet(key)
    server = fernet.decrypt(os.getenv('SERVER')).decode()
    database = fernet.decrypt(os.getenv('DATABASE')).decode()
    username = fernet.decrypt(os.getenv('DB_USERNAME')).decode()
    password = fernet.decrypt(os.getenv('DB_PASSWORD')).decode()
    port = fernet.decrypt(os.getenv('DB_PORT')).decode()
    

    # driver = '{/usr/lib/libmsodbcsql-17.so}'
    #driver = '{ODBC Driver 17 for SQL Server}'
    driver = '{SQL Server}'

    cnxn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    return cnxn
