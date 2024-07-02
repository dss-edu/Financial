from finance.connect import connect
from finance.backend import update_fy
from config.settings import SCHOOLS
from config.settings import schoolMonths

if __name__ == "__main__":
     # --- UPDATE ASCENDER CLIENTS ---
    cnxn = connect()
    cursor = cnxn.cursor()
    #cursor.execute("update [dbo].[AscenderData_ACA] set Date = '2023-12-27' where Date = '2024-12-27' and FY = '2023-2024'")
    #cursor.execute("update [dbo].[AscenderData_Cumberland] set Date = '2023-12-20' where Date = '2024-12-20' and FY = '2023-2024'")
    #cursor.execute("update [dbo].[AscenderData_SA] set Date = '2024-01-16' where Date = '2024-06-16' and FY = '2023-2024'")
    #cnxn.commit()
        
    for school in SCHOOLS.keys():
        if school in schoolMonths["septemberSchool"]:
            update_fy(school, '2024')

    # --- UPDATE ASCENDER CLIENTS ---
    cursor.execute("update [dbo].[AscenderDownloader] set status = '1'")
    cnxn.commit()
    cnxn.close()
