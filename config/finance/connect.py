
import pyodbc

def connect():
    server = 'aca-mysqlserver1.database.windows.net'
    database = 'Database1'
    username = 'aca-user1'
    password = 'Pokemon!123'
    port = '1433'
    

    driver = '{/usr/lib/libmsodbcsql-17.so}'
    #driver = '{ODBC Driver 17 for SQL Server}'
    #driver = '{SQL Server}'

    cnxn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    return cnxn