from finance.connect import connect
from finance.backend import update_fy
from config.settings import SCHOOLS

if __name__ == "__main__":
    for school in SCHOOLS.keys():
        update_fy(school, '2024')

    # --- UPDATE ASCENDER CLIENTS ---
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute("update [dbo].[AscenderDownloader] set status = '1'")
    cnxn.commit()
    cnxn.close()
