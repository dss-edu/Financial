from .connect import connect
# from connect import connect
from time import strftime
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import json
import os
import re
from collections import defaultdict
from config import settings
import shutil
# from django.core.files.base import ContentFile
# from django.core.files.storage import default_storage
# from django.core.files.storage import FileSystemStorage

# Get the current date
current_date = datetime.now()
# Extract the month number from the current date
month_number = current_date.month
curr_year = current_date.year


JSON_DIR = settings.MEDIA_ROOT
SCHOOLS = settings.SCHOOLS
db = settings.db
schoolCategory = settings.schoolCategory
schoolMonths = settings.schoolMonths
school_fye = settings.school_fye


def update_db():
    for school, name in SCHOOLS.items():
        writeCodes(school, db[school]['db'], year)
        profit_loss(school) 
        balance_sheet(school)
        cashflow(school)
        excel(school)
        charter_first(school)
        updateGraphDB(school, False)
        profit_loss_chart(school)
        profit_loss_date(school)
        
def update_school(school):
    anchor_year=""
    profit_loss(school,anchor_year)  # 
    balance_sheet(school,anchor_year) #
    cashflow(school,anchor_year) #
    charter_first(school) #
    excel(school,anchor_year) #
    updateGraphDB(school, False)
    profit_loss_chart(school) #
    profit_loss_date(school)  # 

def update_fy(school,year):    
    writeCodes(school, db[school]['db'], year)
    updateDescription(db[school]['db'], school)
    profit_loss(school,year) 
    balance_sheet(school,year)
    cashflow(school,year)
    charter_first(school)
    updateGraphDB(school, True)
    profit_loss_chart(school)
    profit_loss_date(school)
    excel(school,year)
    if school in schoolCategory["ascender"]:
        balance_sheet_asc(school,year)        
    school_status(school)
    
def profit_loss(school,year):
    print("profit_loss")
    print(school)
    print(year)
    present_date = datetime.today().date()   
    present_year = present_date.year
    today_date = datetime.now()
    today_month = today_date.month
    next_month = present_date + timedelta(days=30)
    last_update = today_date.strftime('%Y-%m-%d')


    #LAST UPDATE


    if year:
        year = int(year)
        if year == present_year:
            
            print("year",year)

            if school in schoolMonths["septemberSchool"]:
                if today_month <= 8:
                    
                    start_year = year - 1
                    present_year = present_year - 1
                    FY_year_current = year - 1
                else: 
                    start_year = year 
                    FY_year_current = year
            else:
                if today_month <= 6:
                    start_year = year - 1
                    present_year = present_year - 1
                    FY_year_current = year - 1
                else: 
                    start_year = year 
                    FY_year_current = year
        else:
            start_year = year 
            FY_year_current = year







        
        
        
        if school in schoolMonths["julySchool"]:
            current_date = datetime(start_year, 7, 1).date()
            
        else:
            current_date = datetime(start_year, 9, 1).date() 
        current_year = current_date.year
     
    else:
        start_year = 2021
        current_date = datetime.today().date()   
        current_year = current_date.year
        FY_year_current = current_year
        
    while start_year <= FY_year_current:
        print(start_year)
        print(FY_year_current)
        # if today_month == 1:
        #     start_year = start_year - 1 
        FY_year_1 = start_year
        FY_year_2 = start_year + 1 
        july_date_start  = datetime(FY_year_1, 7, 1).date()
        
        july_date_end  = datetime(FY_year_2, 6, 30).date()
        september_date_start  = datetime(FY_year_1, 9, 1).date()
        september_date_end  = datetime(FY_year_2, 8, 31).date()

        start_year = FY_year_2

        
      

          
        cnxn = connect()
        cursor = cnxn.cursor()
        cursor.execute(f"SELECT  * FROM [dbo].{db[school]['object']};")
        rows = cursor.fetchall()
        

        data = []
        for row in rows:
            if row[5] == school:

                row_dict = {
                    "fund": row[0],
                    "obj": row[1],
                    "description": row[2],
                    "category": row[3],
                    "value": row[4], #NOT BEING USED. DATA IS COMING FROM GL
                    "school":row[5],
                }
                data.append(row_dict)

        cursor.execute(f"SELECT  * FROM [dbo].{db[school]['function']};")

        rows = cursor.fetchall()

        data2 = []
        for row in rows:
            if row[5] == school:        
                row_dict = {
                    "func_func": row[0],
                    "obj": row[1],
                    "desc": row[2],
                    "category": row[3],
                    "budget":row[4], #NOT BEING USED. DATA IS COMING FROM GL
                    "school": row[5],

                }
                data2.append(row_dict)

        if school in schoolCategory["ascender"]:
            # cursor.execute(
            #     f"SELECT * FROM [dbo].{db[school]['db']}  as AA where AA.Number != 'BEGBAL' and AA.Type != 'EN'  AND (UPPER(AA.WorkDescr) NOT LIKE '%BEG BAL%' AND UPPER(AA.WorkDescr) NOT LIKE '%BEGBAL%') AND UPPER(AA.WorkDescr) NOT LIKE '%BEGINNING BAL%'"
            # )
            cursor.execute(
                f"SELECT * FROM [dbo].{db[school]['db']}  as AA where AA.Number != 'BEGBAL' and AA.Type != 'EN'  "
            )
        else:
            cursor.execute(f"SELECT * FROM [dbo].{db[school]['db']} where source != 'RE';")

        rows = cursor.fetchall()

        data3 = []

       
        if school in schoolMonths["julySchool"]:
            current_month = july_date_start
        else:
            current_month = september_date_start

        
        
        
        last_month = ""
        last_month_name = ""
        last_month_number = ""
        formatted_last_month = ""


        if school in schoolCategory["ascender"]:
            for row in rows:
                expend = float(row[17])
                date = row[11]
                if isinstance(row[11], datetime):
                    date = row[11].strftime("%Y-%m-%d")
                acct_per_month_string = datetime.strptime(date, "%Y-%m-%d")
                acct_per_month = acct_per_month_string.strftime("%m")


                db_date = row[22].split('-')[0]

                if isinstance(row[11],datetime):
                    date_checker = row[11].date()
                else:
                    date_checker = datetime.strptime(row[11], "%Y-%m-%d").date()
                   
    

                #convert data
                db_date = int(db_date)
                curr_fy = int(FY_year_1)
   

                if db_date == curr_fy:
                    if next_month > date_checker: #checks whether the date in data3 will be greater than next month. 
                        if date_checker > current_month:
                            current_month = date_checker.replace(day=1)
                            
                        
                        
                        
                        row_dict = {
                            "fund": row[0],
                            "func": row[1],
                            "obj": row[2],
                            "sobj": row[3],
                            "org": row[4],
                            "fscl_yr": row[5],
                            "pgm": row[6],
                            "edSpan": row[7],
                            "projDtl": row[8],
                            "AcctDescr": row[9],
                            "Number": row[10],
                            "Date": date,
                            "AcctPer": row[12],
                            "Est": row[13],
                            "Real": row[14],
                            "Appr": row[15],
                            "Encum": row[16],
                            "Expend": expend,
                            "Bal": row[18],
                            "WorkDescr": row[19],
                            "Type": row[20],
                            "Contr": row[21],
                        }
                        data3.append(row_dict)
                
        

        else:        
            for row in rows:
                amount = float(row[19])
                date = row[9]
                
                if isinstance(row[9], datetime):
                    date = row[9].strftime("%Y-%m-%d")
                acct_per_month_string = datetime.strptime(date, "%Y-%m-%d")
                acct_per_month = acct_per_month_string.strftime("%m")

                if isinstance(row[9], (datetime, datetime.date)):
                    date_checker = row[9].date()
                else:
                    date_checker = datetime.strptime(row[9], "%Y-%m-%d").date()

                if school in schoolMonths["julySchool"]:
                
                    if date_checker >= july_date_start and date_checker <= july_date_end:
                        if date_checker > current_month:
                            current_month = date_checker.replace(day=1)

                        
                        row_dict = {
                            "fund": row[0],
                            "func": row[2],
                            "obj": row[3],
                            "sobj": row[4],
                            "org": row[5],
                            "fscl_yr": row[6],
                            "Date": date,
                            "AcctPer":row[10],
                            "Amount": amount,
                            "Budget":row[20],
                            "BegBal":row[21],
                            
                        }

                        data3.append(row_dict)

                else:
                    if date_checker >= september_date_start and date_checker <= september_date_end:
                        if date_checker >= current_month:
                            current_month = date_checker.replace(day=1)

                       
                        row_dict = {
                            "fund": row[0],
                            "func": row[2],
                            "obj": row[3],
                            "sobj": row[4],
                            "org": row[5],
                            "fscl_yr": row[6],
                            "Date": date,
                            "AcctPer":row[10],
                            "Amount": amount,
                            "Budget":row[20],
                            "BegBal":row[21],
                            
                        }

                   
                        data3.append(row_dict)
  
        if FY_year_1 == present_year:
            print("current_month")
    
        else:
            if school in schoolMonths["julySchool"]:
                current_month = july_date_end
            else:
                current_month = september_date_end

        # last_month = (current_month.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)                      
        last_month = current_month.replace(day=1) - relativedelta(days=1)                      
        last_month_name = last_month.strftime("%B")
        last_month_number = last_month.month
        formatted_last_month = last_month.strftime('%B %d, %Y')
        db_last_month = last_month.strftime("%Y-%m-%d")
        
  
   
        if present_year == FY_year_1:
            first_day_of_next_month = current_month.replace(day=1, month=current_month.month%12 + 1)
            last_day_of_current_month = first_day_of_next_month - timedelta(days=1)

            if current_month <= last_day_of_current_month:
                current_month = current_month.replace(day=1) - timedelta(days=1)
                last_month = (current_month.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)                      
                last_month_name = last_month.strftime("%B")
                last_month_number = last_month.month
                formatted_last_month = last_month.strftime('%B %d, %Y')  
                db_last_month = last_month.strftime("%Y-%m-%d")
   



    


        cursor.execute(f"SELECT * FROM [dbo].{db[school]['adjustment']} ")
        rows = cursor.fetchall()

        adjustment = []

        if school in schoolCategory["ascender"]:
            for row in rows:
                expend = float(row[17])
                row_dict = {
                    "fund": row[0],
                    "func": row[1],
                    "obj": row[2],
                    "sobj": row[3],
                    "org": row[4],
                    "fscl_yr": row[5],
                    "pgm": row[6],
                    "edSpan": row[7],
                    "projDtl": row[8],
                    "AcctDescr": row[9],
                    "Number": row[10],
                    "Date": row[11],
                    "AcctPer": row[12],
                    "Est": row[13],
                    "Real": row[14],
                    "Appr": row[15],
                    "Encum": row[16],
                    "Expend": expend,
                    "Bal": row[18],
                    "WorkDescr": row[19],
                    "Type": row[20],
                    "School": row[21],
                }
                adjustment.append(row_dict)

        cursor.execute(f"SELECT * FROM [dbo].{db[school]['code']};")
        rows = cursor.fetchall()

        data_expensebyobject = []

        for row in rows:
            budgetformat = "{:,.0f}".format(float(row[2])) if row[2] else ""
            row_dict = {
                "obj": row[0],
                "Description": row[1],
                "budget": budgetformat,
            }

            data_expensebyobject.append(row_dict)

        cursor.execute(f"SELECT * FROM [dbo].{db[school]['activities']};")
        rows = cursor.fetchall()

        data_activities = []

        for row in rows:
            if row[3] == school:
                row_dict = {
                    "obj": row[0],
                    "Description": row[1],
                    "Category": row[2],
                    "school": row[3],
                }

                data_activities.append(row_dict)

        def format_value_dollars(value):
            value = round(value,2)
            if value > 0:
                return "${:,.0f}".format(round(value))
            elif value < 0:
                return "$({:,.0f})".format(abs(round(value)))
            else:
                return ""
        def format_value(value):

            if value > 0:
                return "{:,.0f}".format(round(value))
            elif value < 0:
                return "({:,.0f})".format(abs(round(value)))
            else:
                return ""

        def format_value_dollars_negative(value):
            value = round(value,2)
            if value > 0:
                return "$({:,.0f})".format(abs(round(value)))
                
            elif value < 0:
                
                return "${:,.0f}".format(abs(round(value)))
            else:
                return ""

        def format_value_negative(value):
            if value > 0:
                return "({:,.0f})".format(abs(round(value)))
                
            elif value < 0:
                
                return "{:,.0f}".format(abs(round(value)))
            else:
                return ""

        
        
        # current_month = current_date.replace(day=1)
        # last_month = current_month - relativedelta(days=1)
        # last_month_name = last_month.strftime("%B")
        # formatted_last_month = last_month.strftime('%B %d, %Y')
        # last_month_number = last_month.month

        month_exception = abs(last_month_number) + 1 
        if month_exception == 13:
            month_exception = 1
            
        month_exception_str = str(month_exception).zfill(2)


        if school in schoolMonths["julySchool"]:
                ytd_budget_test = last_month_number - 6
                if month_exception == 7:
                    month_exception = ""
                    month_exception_str = ""             
        else:
            if last_month_number >= 9:

                ytd_budget_test = last_month_number - 8
            else:
                ytd_budget_test = last_month_number + 4

            if month_exception == 9:
                    month_exception = "" 
                    month_exception_str = ""
        ytd_budget = abs(ytd_budget_test) / 12


        if ytd_budget_test == 1 or ytd_budget_test == 12:
            formatted_ytd_budget = f"{ytd_budget * 100:.0f}"
        
        else:
            formatted_ytd_budget = (
            f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
            )
            if formatted_ytd_budget.startswith("0."):
                formatted_ytd_budget = formatted_ytd_budget[2:]

        expend_key = "Expend"
        est_key = "Est"
        expense_key = "Expend"
        real_key = "Real"
        appr_key = "Appr"
        encum_key = "Encum"
        if school in schoolCategory["skyward"]:
            expense_key = "Amount"
            expend_key = "Amount"
            est_key = "Budget"
            real_key = "Amount"
            appr_key = "Budget"
            encum_key = "Amount"

        
        acct_per_values = [
            "01",
            "02",
            "03",
            "04",
            "05",
            "06",
            "07",
            "08",
            "09",
            "10",
            "11",
            "12",
        ]

        # for item in data:
        #     fund = item["fund"]
        #     obj = item["obj"]
            
        #     for i, acct_per in enumerate(acct_per_values, start=1):
        #         total_real = sum(
        #             entry[real_key]
        #             for entry in data3
        #             if entry["fund"] == fund
        #             and entry["obj"] == obj
        #             and entry["AcctPer"] == acct_per
        #         )
        #         total_adjustment = sum(
        #                 entry[real_key]
        #                 for entry in adjustment
        #                 if entry["fund"] == fund
        #                 and entry["AcctPer"] == acct_per
        #                 and entry["obj"] == obj
        #                 and entry["School"] == school
        #             )
        #         item[f"total_check{i}"] = total_real + total_adjustment


    # july_date  = datetime(current_year, 7, 1).date()
    # september_date  = datetime(current_year, 9, 1).date()
    # FY_year_1 = last_year
    # FY_year_2 = current_year
    # for item in data3:
    #     date_str = item["Date"]
    #     if date_str:
    #         if school == 'manara' or school == 'leadership':
               
    #             date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
    #             if date_obj >= july_date:

    #                 FY_year_1 = current_year
    #                 FY_year_2 = next_year
    #         else:
    #             date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
    #             if date_obj >= september_date:
    #                 september_date = datetime(next_year, 9, 1).date()
    #                 FY_year_1 = current_year
    #                 FY_year_2 = next_year

        #checks if the last month column is empty. if empty. last month will be set to  last two months.
        # if all(item[f"total_check{last_month_number}"] == 0 for item in data):
        #     last_2months = current_month - relativedelta(months=1)
        #     last_2months = last_2months - relativedelta(days=1)
        #     last_month_number = last_2months.month
        #     last_month_name = last_2months.strftime("%B")
        #     formatted_last_month = last_2months.strftime('%B %d, %Y')
        #     last_month_number = last_2months.month
            
        #     if school == 'manara' or school == 'leadership':
        #             ytd_budget_test = last_month_number - 6             
        #     else:
        #         if last_month_number >= 9:

        #             ytd_budget_test = last_month_number - 8
        #         else:
        #             ytd_budget_test = last_month_number + 4
            
        #     ytd_budget = abs(ytd_budget_test) / 12

        #     if ytd_budget_test == 1 or ytd_budget_test == 12:
        #         formatted_ytd_budget = f"{ytd_budget * 100:.0f}"
                
        #     else:

        #         formatted_ytd_budget = (
        #         f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
        #         )


        #         if formatted_ytd_budget.startswith("0."):
        #             formatted_ytd_budget = formatted_ytd_budget[2:]
    
        # CALCULATIONS START REVENUES 
        total_lr =  {acct_per: 0 for acct_per in acct_per_values}
        total_spr =  {acct_per: 0 for acct_per in acct_per_values}
        total_fpr =  {acct_per: 0 for acct_per in acct_per_values}
        total_revenue = {acct_per: 0 for acct_per in acct_per_values}
        ytd_total_revenue = 0
        ytd_total_lr  = 0
        ytd_total_spr = 0
        ytd_total_fpr = 0
        variances_revenue = 0

        totals = {
            "total_ammended": 0,
            "total_ammended_lr": 0,
            "total_ammended_spr": 0,
            "total_ammended_fpr": 0,
        }
                
                
        for item in data:
            fund = item["fund"]
            obj = item["obj"]
            category = item["category"]
            ytd_total = 0

            #PUT IT BACK WHEN YOU WANT TO GET THE GL FOR AMMENDED BUDGET FOR REVENUES
            if school in schoolCategory["skyward"]:
                
                total_budget = sum(
                    entry[est_key]
                    for entry in data3
                    if entry["fund"] == fund
                    and entry["obj"] == obj
                    and entry["Date"] <= db_last_month
                  
            
                                
                )
                total_adjustment_budget = sum(
                    entry[est_key]
                    for entry in adjustment
                    if entry["fund"] == fund
                    and entry["obj"] == obj
                    and entry["School"] == school
                    and entry[est_key] is not None 
                    and not isinstance(entry[est_key], str) 
                    and entry["Date"] <= db_last_month
                                
                )
                item["total_budget"] = total_adjustment_budget + total_budget
            else:
                total_budget = sum(
                    entry[est_key]
                    for entry in data3
                    if entry["fund"] == fund
                    and entry["obj"] == obj
                    and entry["Type"] == "GJ" 
             
                          
                )
                total_adjustment_budget = sum(
                    entry[est_key]
                    for entry in adjustment
                    if entry["fund"] == fund
                    and entry["obj"] == obj
                    and entry["School"] == school 
                    and entry[est_key] is not None 
                    and not isinstance(entry[est_key], str)              
                )
                item["total_budget"] = total_adjustment_budget + total_budget
                

            totals["total_ammended"] += item["total_budget"]
            item[f"ytd_budget"] = item["total_budget"] * ytd_budget
                    
            if category == 'Local Revenue':
                totals["total_ammended_lr"] += item["total_budget"]
            elif category == 'State Program Revenue':
                totals["total_ammended_spr"] += item["total_budget"]
            elif category == 'Federal Program Revenue':
                totals["total_ammended_fpr"] += item["total_budget"]
                
            for i, acct_per in enumerate(acct_per_values, start=1):
                if school in schoolMonths['julySchool']:
                    total_real = sum(
                        entry[real_key]
                        for entry in data3
                        if entry["fund"] == fund
                        and entry["obj"] == obj
                        and entry["AcctPer"] == acct_per
                        
                    )
                else:
                    total_real = sum(
                        entry[real_key]
                        for entry in data3
                        if entry["fund"] == fund
                        and entry["obj"] == obj
                        and entry["AcctPer"] == acct_per
                    
                    )
                total_adjustment = sum(
                        entry[real_key]
                        for entry in adjustment
                        if entry["fund"] == fund
                        and entry["AcctPer"] == acct_per
                        and entry["obj"] == obj
                        and entry["School"] == school
                        and entry[real_key] is not None 
                        and not isinstance(entry[real_key], str) 
                    )
                item[f"total_real{i}"] = total_real + total_adjustment 
            
                total_revenue[acct_per] += (item[f"total_real{i}"])

                if category == 'Local Revenue':
                    total_lr[acct_per] += (item[f"total_real{i}"])
                    if i != month_exception:
                        ytd_total_lr += (item[f"total_real{i}"])
                    
                if category == 'State Program Revenue':
                    total_spr[acct_per] += (item[f"total_real{i}"])
                    if i != month_exception:
                        ytd_total_spr += (item[f"total_real{i}"])
                    
                if category == 'Federal Program Revenue':
                    total_fpr[acct_per] += (item[f"total_real{i}"])
                    if i != month_exception:
                        ytd_total_fpr += (item[f"total_real{i}"])

            for month_number in range(1, 13):
                if month_number != month_exception:
                    ytd_total += (item[f"total_real{month_number}"])
           

            item["ytd_total"] = ytd_total
            item["variances"] = item["ytd_total"] +item[f"ytd_budget"]
            item[f"ytd_budget"] = format_value(item[f"ytd_budget"])
        
        ytd_total_revenue = abs(sum(value for key, value in total_revenue.items() if key != month_exception_str))
        #ytd_total_revenue = abs(sum(total_revenue.values())) abs(sum(value for key, value in total_revenue.items() if key != month_exception_str))
        ytd_ammended_total = totals["total_ammended"] * ytd_budget
        ytd_ammended_total_lr = totals["total_ammended_lr"] * ytd_budget
        ytd_ammended_total_spr = totals["total_ammended_spr"] * ytd_budget
        ytd_ammended_total_fpr = totals["total_ammended_fpr"] * ytd_budget

        variances_revenue = (ytd_total_revenue - ytd_ammended_total)
        variances_revenue_lr = (ytd_total_lr + ytd_ammended_total_lr)
        variances_revenue_spr = (ytd_total_spr + ytd_ammended_total_spr)
        variances_revenue_fpr = (ytd_total_fpr + ytd_ammended_total_fpr)

        var_ytd = "{:d}%".format(round(abs(ytd_total_revenue / totals["total_ammended"]*100))) if totals["total_ammended"] != 0 else ""
        var_ytd_lr = "{:d}%".format(round(abs(ytd_total_lr / totals["total_ammended_lr"]*100))) if totals["total_ammended_lr"] != 0 else ""
        var_ytd_spr = "{:d}%".format(round(abs(ytd_total_spr / totals["total_ammended_spr"]*100))) if totals["total_ammended_spr"] != 0 else ""
        var_ytd_fpr = "{:d}%".format(round(abs(ytd_total_fpr / totals["total_ammended_fpr"]*100))) if totals["total_ammended_fpr"] != 0 else ""
        #REVENUES CALCULATIONS END
        
        # CALCULATION START FIRST TOTAL AND DEPRECIATION AND AMORTIZATION (SBD) 
        first_total = 0
        first_ytd_total = 0
        first_total_months =  {acct_per: 0 for acct_per in acct_per_values}
        ytd_ammended_total_first=0
        variances_first_total = 0
        var_ytd_first_total = 0

        dna_total = 0
        dna_ytd_total = 0
        dna_total_months =  {acct_per: 0 for acct_per in acct_per_values}
        ytd_ammended_dna=0
        variances_dna = 0
        var_ytd_dna = 0
    
        for item in data2:
            if item["category"] != "Depreciation and Amortization":
                func = item["func_func"]
                obj = item["obj"]
                ytd_total = 0

                if school in schoolCategory["skyward"]:
                    total_func_func = sum(
                            entry[appr_key]
                            for entry in data3
                            if entry["func"] == func  
                            and entry["obj"] != '6449'
                            and entry["Date"] <= db_last_month
                            
                        )
                else:
                    total_func_func = sum(
                            entry[appr_key]
                            for entry in data3
                            if entry["func"] == func  
                            and entry["obj"] != '6449'
                            and entry["Type"] == 'GJ'
                            and entry["Date"] <= db_last_month
                      
                        
                        )
                total_adjustment_func = sum(
                        entry[appr_key]
                        for entry in adjustment
                        if entry["func"] == func  
                        and entry["obj"] != '6449' 
                        and entry["School"] == school
                        and entry[appr_key] is not None 
                        and not isinstance(entry[appr_key], str)  
                    )
                
                if school in schoolCategory["skyward"]:
                    item['total_budget'] = total_func_func + total_adjustment_func
                else:
                    item['total_budget'] = -(total_func_func + total_adjustment_func)
                
                for i, acct_per in enumerate(acct_per_values, start=1):
                    total_func = sum(
                        entry[expend_key]
                        for entry in data3
                        if entry["func"] == func 
                        and entry["AcctPer"] == acct_per 
                        and entry["obj"] != '6449'
                    )
                    total_adjustment = sum(
                        entry[expend_key]
                        for entry in adjustment
                        if entry["func"] == func 
                        and entry["AcctPer"] == acct_per 
                        and entry["obj"] != '6449' 
                        and entry["School"] == school
                    )
                    item[f"total_func{i}"] = total_func + total_adjustment
                    first_total_months[acct_per] += item[f"total_func{i}"]



                for month_number in range(1, 13):
                    if month_number != month_exception:
                        ytd_total += (item[f"total_func{month_number}"])
            
                item["ytd_total"] = ytd_total
                first_total += item['total_budget']
                first_ytd_total += item["ytd_total"]
                item[f"ytd_budget"] = item['total_budget'] * ytd_budget

                item["variances"] =  item[f"ytd_budget"] -item["ytd_total"]
                variances_first_total += item["variances"]
    
            
                item["var_ytd"] =  "{:d}%".format(round(abs(item["ytd_total"] /item['total_budget'] *100))) if item['total_budget'] != 0 else ""
            
        ytd_ammended_total_first = first_total * ytd_budget
        var_ytd_first_total = "{:d}%".format(round(abs( first_ytd_total/first_total*100))) if first_total != 0 else ""


        for item in data2:
            if item["category"] == "Depreciation and Amortization":
                func = item["func_func"]
                obj = item["obj"]
                ytd_total = 0            
            
                if school in schoolCategory["skyward"]:
                    total_func_func = sum(
                            entry[appr_key]
                            for entry in data3
                            if entry["func"] == func  
                            and entry["obj"] == '6449'
                            and entry["Date"] <= db_last_month 
                          
                        )
                else:
                    total_func_func = sum(
                        entry[appr_key]
                        for entry in data3
                        if entry["func"] == func  
                        and entry["obj"] == '6449'
                        and entry["Type"] == 'GJ'
                        and entry["Date"] <= db_last_month 
                    )
                total_adjustment_func = sum(
                        entry[appr_key]
                        for entry in adjustment
                        if entry["func"] == func  
                        and entry["obj"] == '6449' 
                        and entry["School"] == school
                        and entry[appr_key] is not None 
                        and not isinstance(entry[appr_key], str)
                    )

                if school in schoolCategory["skyward"]:
                    item['total_budget'] = total_func_func + total_adjustment_func
                else:
                    item['total_budget'] = -(total_func_func + total_adjustment_func)

                
                for i, acct_per in enumerate(acct_per_values, start=1):
                    total_func = sum(
                        entry[expend_key]
                        for entry in data3
                        if entry["func"] == func
                        and entry["AcctPer"] == acct_per
                        and entry["obj"] == obj
                    )
                    total_adjustment = sum(
                        entry[expend_key]
                        for entry in adjustment
                        if entry["func"] == func
                        and entry["AcctPer"] == acct_per
                        and entry["obj"] == obj
                        and entry["School"] == school
                        and entry[expend_key] is not None 
                        and not isinstance(entry[expend_key], str)
                    )
                
                    item[f"total_func2_{i}"] = total_func + total_adjustment
                    dna_total_months[acct_per] += item[f"total_func2_{i}"]
                
                for month_number in range(1, 13):
                    if month_number != month_exception:
                        ytd_total += (item[f"total_func2_{month_number}"])
            
                item["ytd_total"] = ytd_total
                dna_total += item['total_budget']
                dna_ytd_total += item["ytd_total"]
                item[f"ytd_budget"] = item['total_budget'] * ytd_budget
                item["variances"] =  item[f"ytd_budget"] -item["ytd_total"]
                variances_dna+= item["variances"]
                item["var_ytd"] =  "{:d}%".format(round(abs( item["ytd_total"]/item['total_budget'] *100))) if item['total_budget'] != 0 else ""
                ytd_ammended_dna = dna_total * ytd_budget
                var_ytd_dna = "{:d}%".format(round(abs(dna_ytd_total / ytd_ammended_dna*100))) if ytd_ammended_dna != 0 else ""
        #CALCULATION END FIRST TOTAL AND DNA
        
        #CALCULATION START SURPLUS BEFORE DEFICIT
        total_SBD =  {acct_per: 0 for acct_per in acct_per_values}
        ammended_budget_SBD = 0
        ytd_ammended_SBD = 0 
        ytd_SBD = 0 
        variances_SBD = 0 
        var_SBD = 0

        total_SBD = {
            acct_per: abs(total_revenue[acct_per]) - first_total_months[acct_per]
            for acct_per in acct_per_values
        }

        ammended_budget_SBD = abs(totals["total_ammended"]) - abs(first_total) 

        ytd_ammended_SBD =  abs(ytd_ammended_total) - abs(ytd_ammended_total_first)

        ytd_SBD = ytd_total_revenue - first_ytd_total
        variances_SBD =  ytd_SBD - ytd_ammended_SBD
        var_SBD = "{:d}%".format(round(abs( ytd_SBD/ ammended_budget_SBD*100))) if ammended_budget_SBD != 0 else ""

        #CALCULATION END SURPLUS BEFORE DEFICIT

        #CALCULATION START NET SURPLUS
        total_netsurplus_months =  {acct_per: 0 for acct_per in acct_per_values}
        ammended_budget_netsurplus = 0
        ytd_ammended_netsurplus = 0 
        ytd_netsurplus = 0
        variances_netsurplus = 0
        var_netsurplus = 0

        total_netsurplus_months = {
            acct_per: total_SBD[acct_per] - dna_total_months[acct_per]
            for acct_per in acct_per_values
        }
        ammended_budget_netsurplus = ammended_budget_SBD - dna_total
        print("ammended_budget_netsurplus",ammended_budget_netsurplus)
        ytd_ammended_netsurplus = ytd_ammended_SBD - ytd_ammended_dna
        ytd_netsurplus =  ytd_SBD - dna_ytd_total 
        bs_ytd_netsurplus = ytd_netsurplus
        variances_netsurplus = ytd_netsurplus - ytd_ammended_netsurplus
        var_netsurplus = "{:d}%".format(round(abs(ytd_netsurplus / ammended_budget_netsurplus*100))) if ammended_budget_netsurplus != 0 else "0%"



        # FOR YTD EXPEND PAGE
        obj_ranges = ["61", "62", "63", "64", "65", "66"]
        full_obj_ranges = ["6100","6200","6300","6400","6500","6600"]

        expend_fund = {}
        for item in data3:
            fund_value = item["fund"]
            if fund_value not in expend_fund and fund_value != '000':
                expend_fund[fund_value] = {}
                for i in range(1, len(acct_per_values) + 1):
                    for obj_range in full_obj_ranges:
                        expend_fund[fund_value][f"total_expend_{obj_range}_{i}"] = 0
                        expend_fund[fund_value][f"total_expend_{obj_range}_ytd"] = 0
                        expend_fund[fund_value][f"total_PB_{obj_range}"] = 0
                    expend_fund[fund_value][f"total_PB"] = 0    
                    expend_fund[fund_value][f"total_ytd"] = 0
                    expend_fund[fund_value][f"total_{i}"] = 0
                    expend_fund[fund_value][f"total_budget"] = 0
                    for obj_range in obj_ranges:
                        expend_fund[fund_value][f"total_budget_{obj_range}"] = 0





        for fund_value in expend_fund.keys(): #total fund within each month
            if school in schoolCategory["skyward"]:
                for obj_range in obj_ranges:
                    total_budget = sum(
                        entry[appr_key]
                        for entry in data3
                        if entry["fund"] == fund_value
                        and entry["obj"].startswith(obj_range)
                        )
                
                    expend_fund[fund_value][f"total_budget_{obj_range}"] = total_budget
                    expend_fund[fund_value]["total_budget"] += total_budget
            else:
                for obj_range in obj_ranges:
                    total_budget = sum(
                        entry[appr_key]
                        for entry in data3
                        if entry["fund"] == fund_value
                        and entry["Type"] == 'GJ'
                        and entry["obj"].startswith(obj_range)
                        )
                
                    expend_fund[fund_value][f"total_budget_{obj_range}"] = total_budget
                    expend_fund[fund_value]["total_budget"] += total_budget

            for i, acct_per in enumerate(acct_per_values, start=1):
                total_expend_6100 = sum(
                    entry[expend_key]
                    for entry in data3
                    if entry["fund"] == fund_value
                    and entry["AcctPer"] == acct_per
                    and entry["obj"].startswith("61")
                )
                
                total_expend_6200 = sum(
                    entry[expend_key]
                    for entry in data3
                    if entry["fund"] == fund_value
                    and entry["AcctPer"] == acct_per
                    and entry["obj"].startswith("62")
                )
                total_expend_6300 = sum(
                    entry[expend_key]
                    for entry in data3
                    if entry["fund"] == fund_value
                    and entry["AcctPer"] == acct_per
                    and entry["obj"].startswith("63")
                )
                total_expend_6400 = sum(
                    entry[expend_key]
                    for entry in data3
                    if entry["fund"] == fund_value
                    and entry["AcctPer"] == acct_per
                    and entry["obj"].startswith("64")
                )
                total_expend_6500 = sum(
                    entry[expend_key]
                    for entry in data3
                    if entry["fund"] == fund_value
                    and entry["AcctPer"] == acct_per
                    and entry["obj"].startswith("65")
                )
                total_expend_6600 = sum(
                    entry[expend_key]
                    for entry in data3
                    if entry["fund"] == fund_value
                    and entry["AcctPer"] == acct_per
                    and entry["obj"].startswith("66")
                )
                expend_fund[fund_value][f"total_expend_6100_{i}"] += total_expend_6100
                expend_fund[fund_value][f"total_expend_6200_{i}"] += total_expend_6200
                expend_fund[fund_value][f"total_expend_6300_{i}"] += total_expend_6300
                expend_fund[fund_value][f"total_expend_6400_{i}"] += total_expend_6400
                expend_fund[fund_value][f"total_expend_6500_{i}"] += total_expend_6500
                expend_fund[fund_value][f"total_expend_6600_{i}"] += total_expend_6600
                expend_fund[fund_value][f"total_{i}"] += total_expend_6100 + total_expend_6200 + total_expend_6300 + total_expend_6400 + total_expend_6500 + total_expend_6600

                if i != month_exception:
                    expend_fund[fund_value][f"total_expend_6100_ytd"] += total_expend_6100
                    expend_fund[fund_value][f"total_expend_6200_ytd"] += total_expend_6200
                    expend_fund[fund_value][f"total_expend_6300_ytd"] += total_expend_6300
                    expend_fund[fund_value][f"total_expend_6400_ytd"] += total_expend_6400
                    expend_fund[fund_value][f"total_expend_6500_ytd"] += total_expend_6500
                    expend_fund[fund_value][f"total_expend_6600_ytd"] += total_expend_6600
                    expend_fund[fund_value][f"total_ytd"] += total_expend_6100 + total_expend_6200 + total_expend_6300 + total_expend_6400 + total_expend_6500 +total_expend_6600

            for obj_range in obj_ranges:
                if school in schoolCategory["skyward"]:
                    expend_fund[fund_value][f"total_PB_{obj_range}00"] =   expend_fund[fund_value][f"total_budget_{obj_range}"] - expend_fund[fund_value][f"total_expend_{obj_range}00_ytd"]
                else:
                    expend_fund[fund_value][f"total_PB_{obj_range}00"] =   expend_fund[fund_value][f"total_budget_{obj_range}"] + expend_fund[fund_value][f"total_expend_{obj_range}00_ytd"]
                expend_fund[fund_value][f"total_PB"] += expend_fund[fund_value][f"total_PB_{obj_range}00"]



        for fund_value in expend_fund:
            for i, acct_per in enumerate(acct_per_values, start=1):
                expend_fund[fund_value][f"total_expend_6100_{i}"] = format_value(expend_fund[fund_value][f"total_expend_6100_{i}"])
                expend_fund[fund_value][f"total_expend_6200_{i}"]   = format_value(expend_fund[fund_value][f"total_expend_6200_{i}"])
                expend_fund[fund_value][f"total_expend_6300_{i}"]   = format_value(expend_fund[fund_value][f"total_expend_6300_{i}"]) 
                expend_fund[fund_value][f"total_expend_6400_{i}"]  = format_value(expend_fund[fund_value][f"total_expend_6400_{i}"])
                expend_fund[fund_value][f"total_expend_6500_{i}"]  = format_value(expend_fund[fund_value][f"total_expend_6500_{i}"])
                expend_fund[fund_value][f"total_expend_6600_{i}"]  = format_value(expend_fund[fund_value][f"total_expend_6600_{i}"])
                expend_fund[fund_value][f"total_{i}"] = format_value(expend_fund[fund_value][f"total_{i}"])
            
            expend_fund[fund_value][f"total_expend_6100_ytd"] = format_value(expend_fund[fund_value][f"total_expend_6100_ytd"])
            expend_fund[fund_value][f"total_expend_6200_ytd"] = format_value(expend_fund[fund_value][f"total_expend_6200_ytd"])
            expend_fund[fund_value][f"total_expend_6300_ytd"] = format_value(expend_fund[fund_value][f"total_expend_6300_ytd"])
            expend_fund[fund_value][f"total_expend_6400_ytd"] = format_value(expend_fund[fund_value][f"total_expend_6400_ytd"])
            expend_fund[fund_value][f"total_expend_6500_ytd"] = format_value(expend_fund[fund_value][f"total_expend_6500_ytd"])
            expend_fund[fund_value][f"total_expend_6600_ytd"] = format_value(expend_fund[fund_value][f"total_expend_6600_ytd"])
            expend_fund[fund_value][f"total_ytd"] = format_value(expend_fund[fund_value][f"total_ytd"]) 

            if school in schoolCategory["skyward"]:
                for obj_range in obj_ranges:
                    expend_fund[fund_value][f"total_PB_{obj_range}00"] = format_value(expend_fund[fund_value][f"total_PB_{obj_range}00"])
                    expend_fund[fund_value][f"total_budget_{obj_range}"] = format_value(expend_fund[fund_value][f"total_budget_{obj_range}"])
                expend_fund[fund_value]["total_budget"] = format_value(expend_fund[fund_value]["total_budget"])
                expend_fund[fund_value][f"total_PB"] = format_value(expend_fund[fund_value][f"total_PB"])
            else:
                for obj_range in obj_ranges:
                    expend_fund[fund_value][f"total_PB_{obj_range}00"] = format_value_negative(expend_fund[fund_value][f"total_PB_{obj_range}00"])
                    expend_fund[fund_value][f"total_budget_{obj_range}"] = format_value_negative(expend_fund[fund_value][f"total_budget_{obj_range}"])
                expend_fund[fund_value]["total_budget"] = format_value_negative(expend_fund[fund_value]["total_budget"])
                expend_fund[fund_value][f"total_PB"] = format_value_negative(expend_fund[fund_value][f"total_PB"])


        # END OF YTD EXPEND PAGE
             
        #CALCULATION EXPENSE BY OBJECT(EOC) AND TOTAL EXPENSE

        total_EOC_pc =  {acct_per: 0 for acct_per in acct_per_values} # PAYROLL COSTS
        total_EOC_pcs =  {acct_per: 0 for acct_per in acct_per_values}#Professional and Cont Svcs
        total_EOC_sm =  {acct_per: 0 for acct_per in acct_per_values}#Supplies and Materials
        total_EOC_ooe =  {acct_per: 0 for acct_per in acct_per_values}#Other Operating Expenses
        total_EOC_te =  {acct_per: 0 for acct_per in acct_per_values}#Total Expense
        total_EOC_oe =  {acct_per: 0 for acct_per in acct_per_values}#Other expenses 6449
        total_EOC_cpa =  {acct_per: 0 for acct_per in acct_per_values}#FOR FIXED/CAPITAL ASSETS

        ytd_EOC_pc   = 0
        ytd_EOC_pcs  = 0
        ytd_EOC_sm   = 0
        ytd_EOC_ooe  = 0
        ytd_EOC_te   = 0
        ytd_EOC_oe = 0
        ytd_EOC_cpa = 0 

        #FOR TOTAL EXPENSE
        total_expense = 0 
        total_expense_ytd_budget = 0
        total_expense_months =  {acct_per: 0 for acct_per in acct_per_values}
        total_expense_ytd = 0

        
        total_budget_pc  = 0
        total_budget_pcs = 0
        total_budget_sm = 0
        total_budget_ooe = 0
        total_budget_oe = 0
        total_budget_te = 0
        total_budget_cpa = 0

        ytd_budget_pc = 0
        ytd_budget_pcs = 0
        ytd_budget_sm = 0
        ytd_budget_ooe = 0 
        ytd_budget_oe = 0 
        ytd_budget_te = 0
        ytd_budget_cpa = 0
        
        for item in data_activities:
            obj = item["obj"]
            category = item["Category"]
            ytd_total = 0
            total_budget_data_activities = 0
            
            item["total_budget"] = 0

            if school in schoolCategory["skyward"]:
                total_budget_data_activities = sum(
                    entry[appr_key]
                    for entry in data3
                    if entry["obj"] == obj
           
                    )
                item["total_budget"] = total_budget_data_activities
            else:
                total_budget_data_activities = sum(
                entry[appr_key]
                for entry in data3
                if entry["obj"] == obj
                and entry["Type"] == 'GJ'
          
        
                )
                item["total_budget"] = -(total_budget_data_activities)
            
            item["ytd_budget"] =  item["total_budget"] * ytd_budget
            total_expense += item["total_budget"]  
            total_expense_ytd_budget += item[f"ytd_budget"]
            if category == "Payroll and Benefits":
                total_budget_pc += item["total_budget"]                

            if category == "Professional and Contract Services":          
                total_budget_pcs += item["total_budget"] 

            if category == "Materials and Supplies":       
                total_budget_sm += item["total_budget"]     
                
            if category == "Other Operating Costs":
                total_budget_ooe += item["total_budget"]  

            if category == "Depreciation":  
                total_budget_oe += item["total_budget"]     
                
            if category == "Debt Services": 
                total_budget_te += item["total_budget"]     

            if category == "FIXED/CAPITAL ASSETS": 
                total_budget_cpa += item["total_budget"]      

            for i, acct_per in enumerate(acct_per_values, start=1):
                total_activities = sum(
                    entry[expense_key]
                    for entry in data3
                    if entry["obj"] == obj and entry["AcctPer"] == acct_per
                )
                total_adjustment = sum(
                    entry[expense_key]
                    for entry in adjustment
                    if entry["obj"] == obj 
                    and entry["AcctPer"] == acct_per 
                    and entry["School"] == school
                    and entry[expense_key] is not None 
                    and not isinstance(entry[expense_key], str)
                )

                item[f"total_activities{i}"] = total_activities + total_adjustment
                
                if category == "Payroll and Benefits":
                    total_EOC_pc[acct_per] += item[f"total_activities{i}"]
                
                elif category == "Professional and Contract Services":
                    total_EOC_pcs[acct_per] += item[f"total_activities{i}"]

                elif category == "Materials and Supplies":
                    total_EOC_sm[acct_per] += item[f"total_activities{i}"]
                
                elif category == "Other Operating Costs":
                    total_EOC_ooe[acct_per] += item[f"total_activities{i}"]

                elif category == "Depreciation":
                    total_EOC_oe[acct_per] += item[f"total_activities{i}"]
                
                elif category == "Debt Services":
                    total_EOC_te[acct_per] += item[f"total_activities{i}"]

                elif category == "FIXED/CAPITAL ASSETS":
                    total_EOC_cpa[acct_per] += item[f"total_activities{i}"]

                total_expense_months[acct_per] += item[f"total_activities{i}"]  

            for month_number in range(1, 13):
                if month_number != month_exception:
                    ytd_total += (item[f"total_activities{month_number}"])
            item["ytd_total"] = ytd_total

        total_expense += dna_total
        total_expense_ytd_budget += ytd_ammended_dna
        for acct_per, dna_value in dna_total_months.items():
    
            if acct_per in total_expense_months:
            
                total_expense_months[acct_per] += dna_value

        # ytd_EOC_pc  = sum(total_EOC_pc.values())
        # ytd_EOC_pcs = sum(total_EOC_pcs.values())
        # ytd_EOC_sm  = sum(total_EOC_sm.values())
        # ytd_EOC_ooe = sum(total_EOC_ooe.values())
        # ytd_EOC_te  = sum(total_EOC_te.values())
        # ytd_EOC_oe  = sum(total_EOC_oe.values())
        # ytd_EOC_cpa  = sum(total_EOC_cpa.values())

        ytd_EOC_pc  = (sum(value for key, value in total_EOC_pc.items() if key != month_exception_str))
        ytd_EOC_pcs =  (sum(value for key, value in total_EOC_pcs.items() if key != month_exception_str))
        ytd_EOC_sm  =  (sum(value for key, value in total_EOC_sm.items() if key != month_exception_str))
        ytd_EOC_ooe =  (sum(value for key, value in total_EOC_ooe.items() if key != month_exception_str))
        ytd_EOC_te  =  (sum(value for key, value in total_EOC_te.items() if key != month_exception_str))
        ytd_EOC_oe  =  (sum(value for key, value in total_EOC_oe.items() if key != month_exception_str))
        ytd_EOC_cpa  =  (sum(value for key, value in total_EOC_cpa.items() if key != month_exception_str))
        
        ytd_budget_pc = total_budget_pc * ytd_budget
        ytd_budget_pcs = total_budget_pcs * ytd_budget
        ytd_budget_sm = total_budget_sm * ytd_budget
        ytd_budget_ooe = total_budget_ooe  * ytd_budget
        ytd_budget_oe = total_budget_oe * ytd_budget
        ytd_budget_te = total_budget_te * ytd_budget
        ytd_budget_cpa = total_budget_cpa * ytd_budget

        #temporarily for 6500
        budget_for_6500 = 0
        ytd_budget_for_6500 = 0 

        for item in data_expensebyobject:
            obj = item["obj"]
        

            if obj == "6100":
                category = "Payroll and Benefits"
                item["variances"] = ytd_budget_pc - ytd_EOC_pc
                item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_pc / total_budget_pc*100))) if total_budget_pc != 0 else ""
            elif obj == "6200":
                category = "Professional and Contract Services"
                item["variances"] = ytd_budget_pcs - ytd_EOC_pcs
                item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_pcs / total_budget_pcs*100))) if total_budget_pcs != 0 else ""
            elif obj == "6300":
                category = "Materials and Supplies"
                item["variances"] = ytd_budget_sm - ytd_EOC_sm
                item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_sm / total_budget_sm*100))) if total_budget_sm != 0 else ""
            elif obj == "6400":
                category = "Other Operating Costs"
                item["variances"] = ytd_budget_ooe - ytd_EOC_ooe
                item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_ooe / total_budget_ooe*100))) if total_budget_ooe != 0 else ""
            elif obj == "6449":
                category = "Depreciation"
                item["variances"] = ytd_budget_oe - ytd_EOC_oe
                item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_oe / total_budget_oe*100))) if total_budget_oe != 0 else ""
            elif obj == "6500":
                category = "Debt Services"
                item["variances"] = ytd_budget_te - ytd_EOC_te
                item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_te / total_budget_te*100))) if total_budget_te != 0 else ""
            else:
                
                category = "FIXED/CAPITAL ASSETS"
                item["variances"] = ytd_budget_cpa - ytd_EOC_cpa
                item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_cpa / total_budget_cpa*100))) if total_budget_cpa != 0 else ""

            for i, acct_per in enumerate(acct_per_values, start=1):
                item[f"total_expense{i}"] = sum(
                    entry[f"total_activities{i}"]
                    for entry in data_activities
                    if entry["Category"] == category
                )

            
        #CONTINUATION COMPUTATION TOTAL EXPENSE
        total_expense_ytd = sum([ytd_EOC_te, ytd_EOC_ooe, ytd_EOC_sm, ytd_EOC_pcs, ytd_EOC_pc,dna_ytd_total,ytd_EOC_cpa])
        variances_total_expense = total_expense_ytd_budget - total_expense_ytd
        var_total_expense = "{:d}%".format(round(abs(total_expense_ytd / total_expense*100))) if total_expense != 0 else ""
            

        #CALCULATIONS START NET INCOME
        net_income_budget = 0
        ytd_budget_net_income = 0 
        total_net_income_months =  {acct_per: 0 for acct_per in acct_per_values}
        ytd_net_income = 0
        variances_net_income = 0
        var_net_income = 0 


        budget_net_income = totals["total_ammended"] - total_expense
        print("budget_net_income",budget_net_income)
        print(totals["total_ammended"])
        print(total_expense)
    
        ytd_budget_net_income = ytd_ammended_total - total_expense_ytd_budget
        ytd_net_income = ytd_total_revenue - total_expense_ytd
        variances_net_income = ytd_net_income - ytd_budget_net_income
        var_net_income = "{:d}%".format(round(abs(ytd_net_income / budget_net_income * 100))) if budget_net_income != 0 else "0%"
    
        
        total_net_income_months = {
            acct_per: abs(total_revenue[acct_per]) - total_expense_months[acct_per]
            for acct_per in acct_per_values
        }  

        #FORMAT FOR REVENUE
        formatted_ammended = format_value_dollars(totals["total_ammended"]) if totals["total_ammended"] != 0 else ""
        formatted_ammended_lr = format_value_dollars(totals["total_ammended_lr"]) if totals["total_ammended_lr"] != 0 else ""
        formatted_ammended_spr = format_value(totals["total_ammended_spr"]) if totals["total_ammended_spr"] != 0 else ""
        formatted_ammended_fpr = format_value(totals["total_ammended_fpr"]) if totals["total_ammended_fpr"] != 0 else ""
            

        formatted_total_lr = {acct_per: format_value_dollars_negative(value) for acct_per, value in total_lr.items() if value != 0}
        formatted_total_spr = {acct_per: format_value_negative(value) for acct_per, value in total_spr.items() if value != 0}
        formatted_total_fpr = {acct_per: format_value_negative(value) for acct_per, value in total_fpr.items() if value != 0}
        formatted_total_revenue = {acct_per: format_value_dollars_negative(value) for acct_per, value in total_revenue.items() if value != 0}
        
        ytd_ammended_total = format_value_dollars(ytd_ammended_total )
        ytd_ammended_total_lr =format_value_dollars(ytd_ammended_total_lr ) 
        ytd_ammended_total_spr=format_value(ytd_ammended_total_spr) 
        ytd_ammended_total_fpr=format_value(ytd_ammended_total_fpr) 
        
        ytd_total_revenue = format_value_dollars(ytd_total_revenue)
        ytd_total_lr  = format_value_dollars_negative(ytd_total_lr)
        ytd_total_spr = format_value_negative(ytd_total_spr)
        ytd_total_fpr = format_value_negative(ytd_total_fpr)

        variances_revenue = format_value_dollars(variances_revenue)
        variances_revenue_lr=format_value_dollars_negative(variances_revenue_lr)
        variances_revenue_spr=format_value_negative(variances_revenue_spr)
        variances_revenue_fpr=format_value_negative(variances_revenue_fpr)


        for row in data:
            ytd_total = float(row["ytd_total"])
        
            variances =float(row["variances"])
            total_budget = float(row["total_budget"])
    
            if total_budget is None or total_budget == 0:
                row["total_budget"] = ""
            else:
                row["total_budget"] = format_value(total_budget) 

            if ytd_total is None or ytd_total == 0:
                row["ytd_total"] = ""
            else:
                row["ytd_total"] = format_value_negative(ytd_total)

            if variances is None or variances == 0:
                row["variances"] = ""
            else:
                row["variances"] = format_value(variances)

        # FOR EXPENSE BY OBJECT DEPRECIATION ONLY        
        dna_total_6449 = format_value(dna_total)
        ytd_ammended_dna_6449 = format_value(ytd_ammended_dna)
        dna_ytd_total_6449 = format_value(dna_ytd_total)
        variances_dna_6449 = format_value(variances_dna)   
        dna_total_months_6449 = {acct_per: format_value(value) for acct_per, value in dna_total_months.items() if value != 0}    


        #FORMAT FIRST TOTAL AND DEPRECIATION AND AMORTIZATION(DNA)
        dna_total = format_value_dollars(dna_total)
        first_total = format_value_dollars(first_total)

        ytd_ammended_dna = format_value_dollars(ytd_ammended_dna)
        ytd_ammended_total_first = format_value_dollars(ytd_ammended_total_first)
        
        dna_ytd_total = format_value_dollars(dna_ytd_total)
        first_ytd_total = format_value_dollars(first_ytd_total)

        variances_first_total = format_value_dollars(variances_first_total)
        variances_dna = format_value_dollars(variances_dna)

        first_total_months = {acct_per: format_value_dollars(value) for acct_per, value in first_total_months.items() if value != 0}
        dna_total_months = {acct_per: format_value_dollars(value) for acct_per, value in dna_total_months.items() if value != 0}

        for row in data2:
            ytd_budget =float(row[f"ytd_budget"])
            ytd_total = float(row["ytd_total"])
            variances = float(row["variances"])
            budget = row["total_budget"]
            
            

            if ytd_total is None or ytd_total == 0:
                row[f"ytd_total"] = ""
            else:
                row[f"ytd_total"] = format_value(ytd_total) 
            if var_ytd is None or var_ytd == 0:
                row[f"variances"] = ""
            else:
                row[f"variances"] = format_value(variances)
            
                
            if budget is None or budget == 0:
                row[f"total_budget"] = ""
            else:
                row[f"total_budget"] = format_value(budget)
            if ytd_budget is None or ytd_budget == 0:
                row[f"ytd_budget"] = ""
            else:
                row[f"ytd_budget"] = format_value(ytd_budget)
    



        #FORMAT SURPLUS BEFORE DEFICIT   
        ammended_budget_SBD = format_value_dollars(ammended_budget_SBD)
        ytd_ammended_SBD = format_value_dollars(ytd_ammended_SBD)
        ytd_SBD = format_value_dollars(ytd_SBD)
        variances_SBD = format_value_dollars(variances_SBD)
        
        total_SBD = {acct_per: format_value_dollars(value) for acct_per, value in total_SBD.items() if value != 0}
        
        
        #FORMAT NET SURPLUS 
        ammended_budget_netsurplus = format_value_dollars(ammended_budget_netsurplus)

  
        ytd_ammended_netsurplus = format_value_dollars(ytd_ammended_netsurplus)

        if ammended_budget_netsurplus == "":
            ammended_budget_netsurplus = 0
            ytd_ammended_netsurplus = 0

        ytd_netsurplus = format_value_dollars(ytd_netsurplus)
        variances_netsurplus = format_value_dollars(variances_netsurplus)
        
        total_netsurplus_months = {acct_per: format_value_dollars(value) for acct_per, value in total_netsurplus_months.items() if value != 0}
        
        #FORMAT EXPENSE BY OBJECT CODES
        for row in data_activities:
        
            ytd_total = (row["ytd_total"])
            total_expense_budget = row["total_budget"]
            ytd_budget = row["ytd_budget"]
        
        
            if ytd_total is None or ytd_total == 0:
                row[f"ytd_total"] = ""
            else:
                row[f"ytd_total"] = format_value(ytd_total)

            if total_expense_budget is None or total_expense_budget == 0:
                row[f"total_budget"] = ""
            else:
                row[f"total_budget"] = format_value(total_expense_budget)

            if ytd_budget is None or ytd_budget == 0:
                row[f"ytd_budget"] = ""
            else:
                row[f"ytd_budget"] = format_value(ytd_budget)
        
        for row in data_expensebyobject:
            variances = float(row["variances"])
            if variances is None or variances == 0:
                row[f"variances"] = ""
            else:
                row[f"variances"] = format_value(variances)

        total_EOC_pc = {acct_per: format_value(value) for acct_per, value in total_EOC_pc.items() if value != 0}
        total_EOC_pcs = {acct_per: format_value(value) for acct_per, value in total_EOC_pcs.items() if value != 0} 
        total_EOC_sm = {acct_per: format_value(value) for acct_per, value in total_EOC_sm.items() if value != 0} 
        total_EOC_ooe = {acct_per: format_value(value) for acct_per, value in total_EOC_ooe.items() if value != 0} 
        total_EOC_te = {acct_per: format_value(value) for acct_per, value in total_EOC_te.items() if value != 0}
        total_EOC_oe = {acct_per: format_value(value) for acct_per, value in total_EOC_oe.items() if value != 0}
        total_EOC_cpa = {acct_per: format_value(value) for acct_per, value in total_EOC_cpa.items() if value != 0}  

        ytd_EOC_pc  = format_value(ytd_EOC_pc)
        ytd_EOC_pcs = format_value(ytd_EOC_pcs)
        ytd_EOC_sm  = format_value(ytd_EOC_sm)
        ytd_EOC_ooe = format_value(ytd_EOC_ooe)
        ytd_EOC_te  = format_value(ytd_EOC_te)
        ytd_EOC_oe  = format_value(ytd_EOC_oe)
        ytd_EOC_cpa  = format_value(ytd_EOC_cpa)

        total_budget_pc =  format_value(total_budget_pc)
        total_budget_pcs = format_value(total_budget_pcs)
        total_budget_sm =  format_value(total_budget_sm)
        total_budget_ooe = format_value(total_budget_ooe)
        total_budget_oe =  format_value(total_budget_oe)   
        total_budget_te =  format_value(total_budget_te)
        total_budget_cpa =  format_value(total_budget_cpa)

        ytd_budget_pc = format_value(ytd_budget_pc)
        ytd_budget_pcs =format_value(ytd_budget_pcs)
        ytd_budget_sm = format_value(ytd_budget_sm)
        ytd_budget_ooe =format_value(ytd_budget_ooe)
        ytd_budget_oe = format_value(ytd_budget_oe)
        ytd_budget_te = format_value(ytd_budget_te)
        ytd_budget_cpa = format_value(ytd_budget_cpa)

        #EXPENSE OBJECT FOR FIX
        budget_for_6500 = format_value(budget_for_6500)
        ytd_budget_for_6500 = format_value(ytd_budget_for_6500)

        #FORMAT TOTAL EXPENSE
        total_expense = format_value_dollars(total_expense)
        total_expense_ytd_budget = format_value_dollars(total_expense_ytd_budget)
        total_expense_months = {acct_per: format_value_dollars(value) for acct_per, value in total_expense_months.items() if value != 0} 
        total_expense_ytd = format_value_dollars(total_expense_ytd)
        variances_total_expense =format_value_dollars(variances_total_expense)
            
        
        #FORMAT NET INCOME
        budget_net_income = format_value_dollars(budget_net_income)

        ytd_budget_net_income = format_value_dollars(ytd_budget_net_income)
        if budget_net_income == "":
            budget_net_income = 0
            ytd_budget_net_income = 0
        total_net_income_months = {acct_per: format_value_dollars(value) for acct_per, value in total_net_income_months.items() if value != 0}
        ytd_net_income = format_value_dollars(ytd_net_income)
        variances_net_income = format_value_dollars(variances_net_income)     

        keys_to_check = [
            "total_real1",
            "total_real2",
            "total_real3",
            "total_real4",
            "total_real5",
            "total_real6",
            "total_real7",
            "total_real8",
            "total_real9",
            "total_real10",
            "total_real11",
            "total_real12",
        ]

        keys_to_check_expense = [
            "total_activities1",
            "total_activities2",
            "total_activities3",
            "total_activities4",
            "total_activities5",
            "total_activities6",
            "total_activities7",
            "total_activities8",
            "total_activities9",
            "total_activities10",
            "total_activities11",
            "total_activities12",
        ]
        keys_to_check_expense2 = [
            "total_expense1",
            "total_expense2",
            "total_expense3",
            "total_expense4",
            "total_expense5",
            "total_expense6",
            "total_expense7",
            "total_expense8",
            "total_expense9",
            "total_expense10",
            "total_expense11",
            "total_expense12",
        ]

        
        for row in data_activities:
            for key in keys_to_check_expense:
                value = float(row[key])
                if value == 0:
                    row[key] = ""
                elif value < 0:
                    row[key] = "({:,.0f})".format(abs(float(row[key])))
                elif value != "":
                    row[key] = "{:,.0f}".format(float(row[key]))

        for row in data_expensebyobject:
            for key in keys_to_check_expense2:
                value = float(row[key])
                if value == 0:
                    row[key] = ""
                elif value < 0:
                    row[key] = "({:,.0f})".format(abs(float(row[key])))
                elif value != "":
                    row[key] = "{:,.0f}".format(float(row[key]))

            # for row in data_activitybs:
            # for key in keys_to_check:
            #     value = int(row[key])
            #     if value == 0:
            #         row[key] = ""
            #     elif value < 0:
            #         row[key] = "({:,.0f})".format(abs(float(row[key])))
            #     elif value != "":
            #         row[key] = "{:,.0f}".format(float(row[key]))

        for row in data:
            for key in keys_to_check:
                value = int(row[key])
                if value == 0:
                    row[key] = ""
                elif value < 0:
                    row[key] = "{:,.0f}".format(abs(float(row[key])))
                elif value != "":
                    row[key] = "({:,.0f})".format(float(row[key]))

        # for row in data:
        #     for key in keys_to_check:
        #         if row[key] < 0:
        #             row[key] = -row[key]
        #         else if row[key] > 0:
        #             row[key] = row[key]

        # for row in data:
        #     for key in keys_to_check:
        #         if row[key] != "":
        #             row[key] = "{:,.0f}".format(row[key])


        keys_to_check_func = [
            "total_func1",
            "total_func2",
            "total_func3",
            "total_func4",
            "total_func5",
            "total_func6",
            "total_func7",
            "total_func8",
            "total_func9",
            "total_func10",
            "total_func11",
            "total_func12",
        ]

        keys_to_check_func_2 = [
            "total_func2_1",
            "total_func2_2",
            "total_func2_3",
            "total_func2_4",
            "total_func2_5",
            "total_func2_6",
            "total_func2_7",
            "total_func2_8",
            "total_func2_9",
            "total_func2_10",
            "total_func2_11",
            "total_func2_12",
        ]

        for row in data2:
            for key in keys_to_check_func:
                if key in row and row[key] is not None:
                    row[key] = format_value(row[key])
                else:
                    row[key] = ""


        for row in data2:
            for key in keys_to_check_func_2:
                if key in row and row[key] is not None :
                    row[key] = format_value(row[key])
                else:
                    row[key] = ""
        # for row in data2:
        #     for key in keys_to_check_func_2:
        #         if row[key] != "":
        #             row[key] = "{:,.0f}".format(row[key])

        # if not school == "village-tech":
        #     lr_funds = list(set(row["fund"] for row in data3 if "fund" in row))
        #     lr_funds_sorted = sorted(lr_funds)
        #     lr_obj = list(set(row["obj"] for row in data3 if "obj" in row))
        #     lr_obj_sorted = sorted(lr_obj)
        #
        #     func_choice = list(set(row["func"] for row in data3 if "func" in row))
        #     func_choice_sorted = sorted(func_choice)

        # current_date = datetime.today().date()
        # current_year = current_date.year
        # last_year = current_date - timedelta(days=365)
        # current_month = current_date.replace(day=1)
        # last_month = current_month - relativedelta(days=1)
        # last_month_number = last_month.month
        # ytd_budget_test = last_month_number + 4
        # ytd_budget = ytd_budget_test / 12
        
        sorted_data2 = sorted(data2, key=lambda x: x['func_func'])
        sorted_data = sorted(data, key=lambda x: x['obj'])
        data_activities = sorted(data_activities, key=lambda x: x['obj'])
    

        context = {
            "data": sorted_data,
            "data2": sorted_data2,
            "data3": data3,
            "data_expensebyobject": data_expensebyobject,
            "data_activities": data_activities,
            "last_update": last_update,
            "expend_fund": expend_fund,
            "months":
                    {
                "last_month": formatted_last_month,
                "last_month_number": last_month_number,
                "last_month_name": last_month_name,
                "format_ytd_budget": formatted_ytd_budget,
                "ytd_budget": ytd_budget,
                "FY_year_1":FY_year_1,
                "FY_year_2":FY_year_2,
                "db_last_month": db_last_month,
                "month_exception": month_exception,
                "month_exception_str": month_exception_str,
                

                },
            "totals":{
                #FOR REVENUES
                "total_lr": formatted_total_lr,
                "total_spr": formatted_total_spr,
                "total_fpr": formatted_total_fpr,
                "total_revenue": formatted_total_revenue,
                "total_ammended": formatted_ammended,
                "total_ammended_lr": formatted_ammended_lr,
                "total_ammended_spr": formatted_ammended_spr,
                "total_ammended_fpr": formatted_ammended_fpr,
                "ytd_ammended_total":ytd_ammended_total,
                "ytd_ammended_total_lr":ytd_ammended_total_lr,
                "ytd_ammended_total_spr":ytd_ammended_total_spr,
                "ytd_ammended_total_fpr":ytd_ammended_total_fpr,
                "ytd_total_revenue": ytd_total_revenue,
                "ytd_total_lr": ytd_total_lr,
                "ytd_total_spr": ytd_total_spr,
                "ytd_total_fpr": ytd_total_fpr,
                "variances_revenue":variances_revenue,
                "variances_revenue_lr":variances_revenue_lr,
                "variances_revenue_spr":variances_revenue_spr,
                "variances_revenue_fpr":variances_revenue_fpr,
                "var_ytd":var_ytd,
                "var_ytd_lr":var_ytd_lr,
                "var_ytd_spr":var_ytd_spr,
                "var_ytd_fpr":var_ytd_fpr,

                #FIRST TOTAL
                "first_total":first_total,
                "first_total_months":first_total_months,
                "first_ytd_total":first_ytd_total,
                "ytd_ammended_total_first": ytd_ammended_total_first,
                "variances_first_total":variances_first_total,
                "var_ytd_first_total": var_ytd_first_total,

                # DEPRECIATION AND AMORTIZATION
                "dna_total":dna_total,
                "dna_total_months":dna_total_months,
                "dna_ytd_total":dna_ytd_total,
                "ytd_ammended_dna": ytd_ammended_dna,
                "variances_dna":variances_dna,
                "var_ytd_dna":var_ytd_dna,

                #SURPLUS BEFORE DEFICIT(SBD)
                "total_SBD": total_SBD,
                "ammended_budget_SBD": ammended_budget_SBD,
                "ytd_ammended_SBD": ytd_ammended_SBD,
                "ytd_SBD":ytd_SBD,
                "variances_SBD": variances_SBD,
                "var_SBD":var_SBD,

                #NET SURPLUS    
                "total_netsurplus_months": total_netsurplus_months,
                "ammended_budget_netsurplus": ammended_budget_netsurplus,
                "ytd_ammended_netsurplus" : ytd_ammended_netsurplus,
                "ytd_netsurplus": ytd_netsurplus,
                "variances_netsurplus": variances_netsurplus,
                "var_netsurplus":var_netsurplus,

                #EXPENSE BY OBJECT 
                "total_EOC_pc":total_EOC_pc,
                "total_EOC_pcs":total_EOC_pcs,
                "total_EOC_sm":total_EOC_sm,
                "total_EOC_ooe":total_EOC_ooe,
                "total_EOC_te":total_EOC_te,
                "total_EOC_oe":total_EOC_oe,
                "total_EOC_cpa":total_EOC_cpa,
                "ytd_EOC_pc":ytd_EOC_pc,
                "ytd_EOC_pcs":ytd_EOC_pcs,
                "ytd_EOC_sm":ytd_EOC_sm,
                "ytd_EOC_ooe":ytd_EOC_ooe,
                "ytd_EOC_te":ytd_EOC_te,
                "ytd_EOC_oe":ytd_EOC_oe,
                "ytd_EOC_cpa":ytd_EOC_cpa,
                "total_budget_pc":total_budget_pc,
                "total_budget_pcs":total_budget_pcs,
                "total_budget_sm":total_budget_sm,
                "total_budget_ooe":total_budget_ooe,
                "total_budget_oe":total_budget_oe,
                "total_budget_te":total_budget_te,
                "total_budget_cpa":total_budget_cpa,
                "ytd_budget_pc":ytd_budget_pc,
                "ytd_budget_pcs":ytd_budget_pcs,
                "ytd_budget_sm":ytd_budget_sm,
                "ytd_budget_ooe":ytd_budget_ooe,
                "ytd_budget_oe":ytd_budget_oe,
                "ytd_budget_te":ytd_budget_te,
                "ytd_budget_cpa":ytd_budget_cpa,
                #EXPENSE BY OBJECT FOR DEPRECIATION AND AMORTIZATION
                "dna_total_6449":dna_total_6449,
                "ytd_ammended_dna_6449":ytd_ammended_dna_6449,
                "dna_ytd_total_6449":dna_ytd_total_6449,
                "variances_dna_6449":variances_dna_6449,
                "dna_total_months_6449":dna_total_months_6449,

                #FIX SOON
                "budget_for_6500":budget_for_6500,
                "ytd_budget_for_6500": ytd_budget_for_6500,
                
                #TOTAL EXPENSE 
                "total_expense": total_expense,
                "total_expense_ytd_budget": total_expense_ytd_budget,
                "total_expense_months":total_expense_months,
                "total_expense_ytd":total_expense_ytd,
                "variances_total_expense":variances_total_expense,
                "var_total_expense":var_total_expense,

                #NET INCOME
                "budget_net_income": budget_net_income,
                "ytd_budget_net_income":ytd_budget_net_income,
                "total_net_income_months":total_net_income_months,
                "variances_net_income": variances_net_income,
                "ytd_net_income": ytd_net_income,
                "var_net_income":var_net_income,

                #FOR BS
                "bs_ytd_netsurplus":bs_ytd_netsurplus,
            }
        }

        # if not school == "village-tech":
        #     context["total_lr"] = formatted_total_lr,
            # context["total_netsurplus"] = formatted_total_netsurplus
            # context["total_SBD"] = total_SBD
            # context["ytd_netsurplus"] = formated_ytdnetsurplus


        # if not school == "village-tech":
        #     context["lr_funds"] = lr_funds_sorted
        #     context["lr_obj"] = lr_obj_sorted
        #     context["func_choice"] = func_choice_sorted

        # dict_keys = ["data", "data2", "data3", "data_expensebyobject", "data_activities"]
        


        if FY_year_1 == present_year:
            relative_path = os.path.join("profit-loss", school)
        else:
            relative_path = os.path.join(str(FY_year_1), "profit-loss", school)


        json_path = os.path.join(JSON_DIR, relative_path)

        shutil.rmtree(json_path, ignore_errors=True)
        os.makedirs(json_path, exist_ok=True)

        for key, val in context.items():
            file_path = os.path.join(json_path, f"{key}.json")
            
  
            with open(file_path, "w") as file:
                json.dump(val, file)
                
            print(file_path)

def profit_loss_date(school):

 

    present_date = datetime.today().date()   
    present_year = present_date.year
 
    current_date = datetime.today().date()   
    current_year = current_date.year
    FY_year_current = current_year
     
    FY_year_1 = current_year
    FY_year_2 = current_year + 1 
    july_date_start  = datetime(FY_year_1, 7, 1).date()
    
    july_date_end  = datetime(FY_year_2, 6, 30).date()
    september_date_start  = datetime(FY_year_1, 9, 1).date()
    september_date_end  = datetime(FY_year_2, 8, 31).date()

    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['object']};")
    rows = cursor.fetchall()
    
    data = []
    for row in rows:
        if row[5] == school:
            row_dict = {
                "fund": row[0],
                "obj": row[1],
                "description": row[2],
                "category": row[3],
                "value": row[4], #NOT BEING USED. DATA IS COMING FROM GL
                "school":row[5],
            }
            data.append(row_dict)
    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['function']};")
    rows = cursor.fetchall()
    data2 = []
    for row in rows:
        if row[5] == school:        
            row_dict = {
                "func_func": row[0],
                "obj": row[1],
                "desc": row[2],
                "category": row[3],
                "budget":row[4], #NOT BEING USED. DATA IS COMING FROM GL
                "school": row[5],
            }
            data2.append(row_dict)
    #
    if school in schoolCategory["ascender"]:
        cursor.execute(
            f"SELECT * FROM [dbo].{db[school]['db']}  as AA where AA.Number != 'BEGBAL'  and AA.Type != 'EN';"
        )
    else:
        cursor.execute(f"SELECT * FROM [dbo].{db[school]['db']} where source != 'RE';")
    rows = cursor.fetchall()
    data3 = []
    if school in schoolMonths["julySchool"]:
        current_month = july_date_start
    else:
        current_month = september_date_start
    
    last_month = ""
    last_month_name = ""
    last_month_number = ""
    formatted_last_month = ""
    if school in schoolCategory["ascender"]:
        for row in rows:
            expend = float(row[17])
            date = row[11]
            if isinstance(row[11], datetime):
                date = row[11].strftime("%Y-%m-%d")
            acct_per_month_string = datetime.strptime(date, "%Y-%m-%d")
            acct_per_month = acct_per_month_string.strftime("%m")
            db_date = row[22].split('-')[0]
            if isinstance(row[11],datetime):
                date_checker = row[11].date()
            else:
                date_checker = datetime.strptime(row[11], "%Y-%m-%d").date()
               

            #convert data
            db_date = int(db_date)
            curr_fy = int(FY_year_1)

                
            if db_date == curr_fy:
                if date_checker > current_month:
                    current_month = date_checker.replace(day=1)
                    
                
                
                
                row_dict = {
                    "fund": row[0],
                    "func": row[1],
                    "obj": row[2],
                    "sobj": row[3],
                    "org": row[4],
                    "fscl_yr": row[5],
                    "pgm": row[6],
                    "edSpan": row[7],
                    "projDtl": row[8],
                    "AcctDescr": row[9],
                    "Number": row[10],
                    "Date": date,
                    "AcctPer": acct_per_month,
                    "Est": row[13],
                    "Real": row[14],
                    "Appr": row[15],
                    "Encum": row[16],
                    "Expend": expend,
                    "Bal": row[18],
                    "WorkDescr": row[19],
                    "Type": row[20],
                    "Contr": row[21],
                }
                data3.append(row_dict)
            
        

    else:        
        for row in rows:
            amount = float(row[19])
            date = row[9]
            
            if isinstance(row[9], datetime):
                date = row[9].strftime("%Y-%m-%d")
            acct_per_month_string = datetime.strptime(date, "%Y-%m-%d")
            acct_per_month = acct_per_month_string.strftime("%m")
            if isinstance(row[9], (datetime, datetime.date)):
                date_checker = row[9].date()
            else:
                date_checker = datetime.strptime(row[9], "%Y-%m-%d").date()
            if school in schoolMonths["julySchool"]:
            
                if date_checker >= july_date_start and date_checker <= july_date_end:
                    if date_checker > current_month:
                        current_month = date_checker.replace(day=1)
                    row_dict = {
                        "fund": row[0],
                        "func": row[2],
                        "obj": row[3],
                        "sobj": row[4],
                        "org": row[5],
                        "fscl_yr": row[6],
                        "Date": date,
                        "AcctPer":acct_per_month,
                        "Amount": amount,
                        "Budget":row[20],
                    }
                    data3.append(row_dict)
            else:
                if date_checker >= september_date_start and date_checker <= september_date_end:
                    if date_checker >= current_month:
                        current_month = date_checker.replace(day=1)
                    row_dict = {
                        "fund": row[0],
                        "func": row[2],
                        "obj": row[3],
                        "sobj": row[4],
                        "org": row[5],
                        "fscl_yr": row[6],
                        "Date": date,
                        "AcctPer": row[10],
                        "Amount": amount,
                        "Budget":row[20],
                    }
                    data3.append(row_dict)

    if FY_year_1 == present_year:
        print("current_month")

    else:
        if school in schoolMonths["julySchool"]:
            current_month = july_date_end
        else:
            current_month = september_date_end
    
    last_month = (current_month.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)                      
    last_month_name = last_month.strftime("%B")
    last_month_number = last_month.month
    formatted_last_month = last_month.strftime('%B %d, %Y')
    db_last_month = last_month.strftime("%Y-%m-%d")
    


    if present_year == FY_year_1:
        first_day_of_next_month = current_month.replace(day=1, month=current_month.month%12 + 1)
        last_day_of_current_month = first_day_of_next_month - timedelta(days=1)

        if current_month <= last_day_of_current_month:
            current_month = current_month.replace(day=1) - timedelta(days=1)
            last_month = (current_month.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)                      
            last_month_name = last_month.strftime("%B")
            last_month_number = last_month.month
            formatted_last_month = last_month.strftime('%B %d, %Y')  
            db_last_month = last_month.strftime("%Y-%m-%d")
   
    cursor.execute(f"SELECT * FROM [dbo].{db[school]['adjustment']} ")
    rows = cursor.fetchall()
    adjustment = []

    if school in schoolCategory["ascender"]:
        for row in rows:
            expend = float(row[17])
            row_dict = {
                "fund": row[0],
                "func": row[1],
                "obj": row[2],
                "sobj": row[3],
                "org": row[4],
                "fscl_yr": row[5],
                "pgm": row[6],
                "edSpan": row[7],
                "projDtl": row[8],
                "AcctDescr": row[9],
                "Number": row[10],
                "Date": row[11],
                "AcctPer": row[12],
                "Est": row[13],
                "Real": row[14],
                "Appr": row[15],
                "Encum": row[16],
                "Expend": expend,
                "Bal": row[18],
                "WorkDescr": row[19],
                "Type": row[20],
                "School": row[21],
            }
            adjustment.append(row_dict)
    cursor.execute(f"SELECT * FROM [dbo].{db[school]['code']};")
    rows = cursor.fetchall()
    data_expensebyobject = []
    for row in rows:
        budgetformat = "{:,.0f}".format(float(row[2])) if row[2] else ""
        row_dict = {
            "obj": row[0],
            "Description": row[1],
            "budget": budgetformat,
        }
        data_expensebyobject.append(row_dict)
    cursor.execute(f"SELECT * FROM [dbo].{db[school]['activities']};")
    rows = cursor.fetchall()
    data_activities = []
    for row in rows:
        if row[3] == school:
            row_dict = {
                "obj": row[0],
                "Description": row[1],
                "Category": row[2],
                "school": row[3],
            }
            data_activities.append(row_dict)
    def format_value_dollars(value):
        if value > 0:
            return "${:,.0f}".format(round(value))
        elif value < 0:
            return "$({:,.0f})".format(abs(round(value)))
        else:
            return ""
    def format_value(value):
        if value > 0:
            return "{:,.0f}".format(round(value))
        elif value < 0:
            return "({:,.0f})".format(abs(round(value)))
        else:
            return ""
    def format_value_dollars_negative(value):
        if value > 0:
            return "$({:,.0f})".format(abs(round(value)))
            
        elif value < 0:
            
            return "${:,.0f}".format(abs(round(value)))
        else:
            return ""
    def format_value_negative(value):
        if value > 0:
            return "({:,.0f})".format(abs(round(value)))
            
        elif value < 0:
            
            return "{:,.0f}".format(abs(round(value)))
        else:
            return ""

        
        
    # current_month = current_date.replace(day=1)
    # last_month = current_month - relativedelta(days=1)
    # last_month_name = last_month.strftime("%B")
    # formatted_last_month = last_month.strftime('%B %d, %Y')
    # last_month_number = last_month.month
    month_exception = abs(last_month_number) + 1 
    if month_exception == 13:
        month_exception = 1
        
    month_exception_str = str(month_exception).zfill(2)
    if school in schoolMonths["julySchool"]:
            ytd_budget_test = last_month_number - 6
            if month_exception == 7:
                month_exception = ""
                month_exception_str = ""             
    else:
        if last_month_number >= 9:
            ytd_budget_test = last_month_number - 8
        else:
            ytd_budget_test = last_month_number + 4
        if month_exception == 9:
                month_exception = "" 
                month_exception_str = ""
    ytd_budget = abs(ytd_budget_test) / 12
    if ytd_budget_test == 1 or ytd_budget_test == 12:
        formatted_ytd_budget = f"{ytd_budget * 100:.0f}"
    
    else:
        formatted_ytd_budget = (
        f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
        )
        if formatted_ytd_budget.startswith("0."):
            formatted_ytd_budget = formatted_ytd_budget[2:]
    expend_key = "Expend"
    est_key = "Est"
    expense_key = "Expend"
    real_key = "Real"
    appr_key = "Appr"
    encum_key = "Encum"
    if school in schoolCategory["skyward"]:
        expense_key = "Amount"
        expend_key = "Amount"
        est_key = "Budget"
        real_key = "Amount"
        appr_key = "Budget"
        encum_key = "Amount"
    
    acct_per_values = [
        "01",
        "02",
        "03",
        "04",
        "05",
        "06",
        "07",
        "08",
        "09",
        "10",
        "11",
        "12",
    ]
    

    # CALCULATIONS START REVENUES 
    total_lr =  {acct_per: 0 for acct_per in acct_per_values}
    total_spr =  {acct_per: 0 for acct_per in acct_per_values}
    total_fpr =  {acct_per: 0 for acct_per in acct_per_values}
    total_revenue = {acct_per: 0 for acct_per in acct_per_values}
    ytd_total_revenue = 0
    ytd_total_lr  = 0
    ytd_total_spr = 0
    ytd_total_fpr = 0
    variances_revenue = 0
    totals = {
        "total_ammended": 0,
        "total_ammended_lr": 0,
        "total_ammended_spr": 0,
        "total_ammended_fpr": 0,
    }
                
                
    for item in data:
        fund = item["fund"]
        obj = item["obj"]
        category = item["category"]
        ytd_total = 0
        #PUT IT BACK WHEN YOU WANT TO GET THE GL FOR AMMENDED BUDGET FOR REVENUES
        if school in schoolCategory["skyward"]:
            
            total_budget = sum(
                entry[est_key]
                for entry in data3
                if entry["fund"] == fund
                and entry["obj"] == obj
                and entry["Date"] <= db_last_month
              
        
                            
            )
            total_adjustment_budget = sum(
                entry[est_key]
                for entry in adjustment
                if entry["fund"] == fund
                and entry["obj"] == obj
                and entry["School"] == school
                and entry[est_key] is not None 
                and not isinstance(entry[est_key], str) 
                and entry["Date"] <= db_last_month
                            
            )
            item["total_budget"] = total_adjustment_budget + total_budget
        else:
            total_budget = sum(
                entry[est_key]
                for entry in data3
                if entry["fund"] == fund
                and entry["obj"] == obj
                and entry["Type"] == "GJ" 
         
                      
            )
            total_adjustment_budget = sum(
                entry[est_key]
                for entry in adjustment
                if entry["fund"] == fund
                and entry["obj"] == obj
                and entry["School"] == school 
                and entry[est_key] is not None 
                and not isinstance(entry[est_key], str)              
            )
            item["total_budget"] = total_adjustment_budget + total_budget
            
        totals["total_ammended"] += item["total_budget"]
        item[f"ytd_budget"] = item["total_budget"] * ytd_budget
                
        if category == 'Local Revenue':
            totals["total_ammended_lr"] += item["total_budget"]
        elif category == 'State Program Revenue':
            totals["total_ammended_spr"] += item["total_budget"]
        elif category == 'Federal Program Revenue':
            totals["total_ammended_fpr"] += item["total_budget"]
            
        for i, acct_per in enumerate(acct_per_values, start=1):
            if school in schoolMonths['julySchool']:
                total_real = sum(
                    entry[real_key]
                    for entry in data3
                    if entry["fund"] == fund
                    and entry["obj"] == obj
                    and entry["AcctPer"] == acct_per
                    
                )
            else:
                total_real = sum(
                    entry[real_key]
                    for entry in data3
                    if entry["fund"] == fund
                    and entry["obj"] == obj
                    and entry["AcctPer"] == acct_per
                
                )
            total_adjustment = sum(
                    entry[real_key]
                    for entry in adjustment
                    if entry["fund"] == fund
                    and entry["AcctPer"] == acct_per
                    and entry["obj"] == obj
                    and entry["School"] == school
                    and entry[real_key] is not None 
                    and not isinstance(entry[real_key], str) 
                )
            item[f"total_real{i}"] = total_real + total_adjustment 
        
            total_revenue[acct_per] += (item[f"total_real{i}"])
            if category == 'Local Revenue':
                total_lr[acct_per] += (item[f"total_real{i}"])
                if i != month_exception:
                    ytd_total_lr += (item[f"total_real{i}"])
                
            if category == 'State Program Revenue':
                total_spr[acct_per] += (item[f"total_real{i}"])
                if i != month_exception:
                    ytd_total_spr += (item[f"total_real{i}"])
                
            if category == 'Federal Program Revenue':
                total_fpr[acct_per] += (item[f"total_real{i}"])
                if i != month_exception:
                    ytd_total_fpr += (item[f"total_real{i}"])
        for month_number in range(1, 13):
            if month_number != month_exception:
                ytd_total += (item[f"total_real{month_number}"])
       
        item["ytd_total"] = ytd_total
        item["variances"] = item["ytd_total"] +item[f"ytd_budget"]
        item[f"ytd_budget"] = format_value(item[f"ytd_budget"])
        
    ytd_total_revenue = abs(sum(value for key, value in total_revenue.items() if key != month_exception_str))
    #ytd_total_revenue = abs(sum(total_revenue.values())) abs(sum(value for key, value in total_revenue.items() if key != month_exception_str))
    ytd_ammended_total = totals["total_ammended"] * ytd_budget
    ytd_ammended_total_lr = totals["total_ammended_lr"] * ytd_budget
    ytd_ammended_total_spr = totals["total_ammended_spr"] * ytd_budget
    ytd_ammended_total_fpr = totals["total_ammended_fpr"] * ytd_budget
    variances_revenue = (ytd_total_revenue - ytd_ammended_total)
    variances_revenue_lr = (ytd_total_lr + ytd_ammended_total_lr)
    variances_revenue_spr = (ytd_total_spr + ytd_ammended_total_spr)
    variances_revenue_fpr = (ytd_total_fpr + ytd_ammended_total_fpr)
    var_ytd = "{:d}%".format(round(abs(ytd_total_revenue / totals["total_ammended"]*100))) if totals["total_ammended"] != 0 else ""
    var_ytd_lr = "{:d}%".format(round(abs(ytd_total_lr / totals["total_ammended_lr"]*100))) if totals["total_ammended_lr"] != 0 else ""
    var_ytd_spr = "{:d}%".format(round(abs(ytd_total_spr / totals["total_ammended_spr"]*100))) if totals["total_ammended_spr"] != 0 else ""
    var_ytd_fpr = "{:d}%".format(round(abs(ytd_total_fpr / totals["total_ammended_fpr"]*100))) if totals["total_ammended_fpr"] != 0 else ""
    #REVENUES CALCULATIONS END
    
    # CALCULATION START FIRST TOTAL AND DEPRECIATION AND AMORTIZATION (SBD) 
    first_total = 0
    first_ytd_total = 0
    first_total_months =  {acct_per: 0 for acct_per in acct_per_values}
    ytd_ammended_total_first=0
    variances_first_total = 0
    var_ytd_first_total = 0
    dna_total = 0
    dna_ytd_total = 0
    dna_total_months =  {acct_per: 0 for acct_per in acct_per_values}
    ytd_ammended_dna=0
    variances_dna = 0
    var_ytd_dna = 0

    for item in data2:
        if item["category"] != "Depreciation and Amortization":
            func = item["func_func"]
            obj = item["obj"]
            ytd_total = 0
            if school in schoolCategory["skyward"]:
                total_func_func = sum(
                        entry[appr_key]
                        for entry in data3
                        if entry["func"] == func  
                        and entry["obj"] != '6449'
                        and entry["Date"] <= db_last_month
                        
                    )
            else:
                total_func_func = sum(
                        entry[appr_key]
                        for entry in data3
                        if entry["func"] == func  
                        and entry["obj"] != '6449'
                        and entry["Type"] == 'GJ'
                        and entry["Date"] <= db_last_month
                  
                    
                    )
            total_adjustment_func = sum(
                    entry[appr_key]
                    for entry in adjustment
                    if entry["func"] == func  
                    and entry["obj"] != '6449' 
                    and entry["School"] == school
                    and entry[appr_key] is not None 
                    and not isinstance(entry[appr_key], str)  
                )
            
            if school in schoolCategory["skyward"]:
                item['total_budget'] = total_func_func + total_adjustment_func
            else:
                item['total_budget'] = -(total_func_func + total_adjustment_func)
            
            for i, acct_per in enumerate(acct_per_values, start=1):
                total_func = sum(
                    entry[expend_key]
                    for entry in data3
                    if entry["func"] == func 
                    and entry["AcctPer"] == acct_per 
                    and entry["obj"] != '6449'
                )
                total_adjustment = sum(
                    entry[expend_key]
                    for entry in adjustment
                    if entry["func"] == func 
                    and entry["AcctPer"] == acct_per 
                    and entry["obj"] != '6449' 
                    and entry["School"] == school
                )
                item[f"total_func{i}"] = total_func + total_adjustment
                first_total_months[acct_per] += item[f"total_func{i}"]
            for month_number in range(1, 13):
                if month_number != month_exception:
                    ytd_total += (item[f"total_func{month_number}"])
        
            item["ytd_total"] = ytd_total
            first_total += item['total_budget']
            first_ytd_total += item["ytd_total"]
            item[f"ytd_budget"] = item['total_budget'] * ytd_budget
            item["variances"] =  item[f"ytd_budget"] -item["ytd_total"]
            variances_first_total += item["variances"]

        
            item["var_ytd"] =  "{:d}%".format(round(abs(item["ytd_total"] /item['total_budget'] *100))) if item['total_budget'] != 0 else ""
        
    ytd_ammended_total_first = first_total * ytd_budget
    var_ytd_first_total = "{:d}%".format(round(abs( first_ytd_total/first_total*100))) if first_total != 0 else ""


    for item in data2:
        if item["category"] == "Depreciation and Amortization":
            func = item["func_func"]
            obj = item["obj"]
            ytd_total = 0            
        
            if school in schoolCategory["skyward"]:
                total_func_func = sum(
                        entry[appr_key]
                        for entry in data3
                        if entry["func"] == func  
                        and entry["obj"] == '6449'
                        and entry["Date"] <= db_last_month 
                      
                    )
            else:
                total_func_func = sum(
                    entry[appr_key]
                    for entry in data3
                    if entry["func"] == func  
                    and entry["obj"] == '6449'
                    and entry["Type"] == 'GJ'
                    and entry["Date"] <= db_last_month 
                )
            total_adjustment_func = sum(
                    entry[appr_key]
                    for entry in adjustment
                    if entry["func"] == func  
                    and entry["obj"] == '6449' 
                    and entry["School"] == school
                    and entry[appr_key] is not None 
                    and not isinstance(entry[appr_key], str)
                )
            item['total_budget'] = total_func_func + total_adjustment_func
            
            for i, acct_per in enumerate(acct_per_values, start=1):
                total_func = sum(
                    entry[expend_key]
                    for entry in data3
                    if entry["func"] == func
                    and entry["AcctPer"] == acct_per
                    and entry["obj"] == obj
                )
                total_adjustment = sum(
                    entry[expend_key]
                    for entry in adjustment
                    if entry["func"] == func
                    and entry["AcctPer"] == acct_per
                    and entry["obj"] == obj
                    and entry["School"] == school
                    and entry[expend_key] is not None 
                    and not isinstance(entry[expend_key], str)
                )
            
                item[f"total_func2_{i}"] = total_func + total_adjustment
                dna_total_months[acct_per] += item[f"total_func2_{i}"]
            
            for month_number in range(1, 13):
                if month_number != month_exception:
                    ytd_total += (item[f"total_func2_{month_number}"])
        
            item["ytd_total"] = ytd_total
            dna_total += item['total_budget']
            dna_ytd_total += item["ytd_total"]
            item[f"ytd_budget"] = item['total_budget'] * ytd_budget
            item["variances"] =  item[f"ytd_budget"] -item["ytd_total"]
            variances_dna+= item["variances"]
            item["var_ytd"] =  "{:d}%".format(round(abs( item["ytd_total"]/item['total_budget'] *100))) if item['total_budget'] != 0 else ""
            ytd_ammended_dna = dna_total * ytd_budget
            var_ytd_dna = "{:d}%".format(round(abs(dna_ytd_total / ytd_ammended_dna*100))) if ytd_ammended_dna != 0 else ""
    #CALCULATION END FIRST TOTAL AND DNA
    
    #CALCULATION START SURPLUS BEFORE DEFICIT
    total_SBD =  {acct_per: 0 for acct_per in acct_per_values}
    ammended_budget_SBD = 0
    ytd_ammended_SBD = 0 
    ytd_SBD = 0 
    variances_SBD = 0 
    var_SBD = 0
    total_SBD = {
        acct_per: abs(total_revenue[acct_per]) - first_total_months[acct_per]
        for acct_per in acct_per_values
    }
    ammended_budget_SBD = abs(totals["total_ammended"]) - abs(first_total) 
    ytd_ammended_SBD =  abs(ytd_ammended_total) - abs(ytd_ammended_total_first)
    ytd_SBD = ytd_total_revenue - first_ytd_total
    variances_SBD =  ytd_SBD - ytd_ammended_SBD
    var_SBD = "{:d}%".format(round(abs( ytd_SBD/ ammended_budget_SBD*100))) if ammended_budget_SBD != 0 else ""
    #CALCULATION END SURPLUS BEFORE DEFICIT
    #CALCULATION START NET SURPLUS
    total_netsurplus_months =  {acct_per: 0 for acct_per in acct_per_values}
    ammended_budget_netsurplus = 0
    ytd_ammended_netsurplus = 0 
    ytd_netsurplus = 0
    variances_netsurplus = 0
    var_netsurplus = 0
    total_netsurplus_months = {
        acct_per: total_SBD[acct_per] - dna_total_months[acct_per]
        for acct_per in acct_per_values
    }
    ammended_budget_netsurplus = ammended_budget_SBD - dna_total
    ytd_ammended_netsurplus = ytd_ammended_SBD - ytd_ammended_dna
    ytd_netsurplus =  ytd_SBD - dna_ytd_total 
    bs_ytd_netsurplus = ytd_netsurplus
    variances_netsurplus = ytd_netsurplus - ytd_ammended_netsurplus
    var_netsurplus = "{:d}%".format(round(abs(ytd_netsurplus / ammended_budget_netsurplus*100))) if ammended_budget_netsurplus != 0 else ""
    #CALCULATION EXPENSE BY OBJECT(EOC) AND TOTAL EXPENSE
    total_EOC_pc =  {acct_per: 0 for acct_per in acct_per_values} # PAYROLL COSTS
    total_EOC_pcs =  {acct_per: 0 for acct_per in acct_per_values}#Professional and Cont Svcs
    total_EOC_sm =  {acct_per: 0 for acct_per in acct_per_values}#Supplies and Materials
    total_EOC_ooe =  {acct_per: 0 for acct_per in acct_per_values}#Other Operating Expenses
    total_EOC_te =  {acct_per: 0 for acct_per in acct_per_values}#Total Expense
    total_EOC_oe =  {acct_per: 0 for acct_per in acct_per_values}#Other expenses 6449
    total_EOC_cpa =  {acct_per: 0 for acct_per in acct_per_values}#FOR FIXED/CAPITAL ASSETS
    ytd_EOC_pc   = 0
    ytd_EOC_pcs  = 0
    ytd_EOC_sm   = 0
    ytd_EOC_ooe  = 0
    ytd_EOC_te   = 0
    ytd_EOC_oe = 0
    ytd_EOC_cpa = 0 
    #FOR TOTAL EXPENSE
    total_expense = 0 
    total_expense_ytd_budget = 0
    total_expense_months =  {acct_per: 0 for acct_per in acct_per_values}
    total_expense_ytd = 0
    
    total_budget_pc  = 0
    total_budget_pcs = 0
    total_budget_sm = 0
    total_budget_ooe = 0
    total_budget_oe = 0
    total_budget_te = 0
    total_budget_cpa = 0
    ytd_budget_pc = 0
    ytd_budget_pcs = 0
    ytd_budget_sm = 0
    ytd_budget_ooe = 0 
    ytd_budget_oe = 0 
    ytd_budget_te = 0
    ytd_budget_cpa = 0
        
    for item in data_activities:
        obj = item["obj"]
        category = item["Category"]
        ytd_total = 0
        total_budget_data_activities = 0
        
        item["total_budget"] = 0
        if school in schoolCategory["skyward"]:
            total_budget_data_activities = sum(
                entry[appr_key]
                for entry in data3
                if entry["obj"] == obj
       
                )
            item["total_budget"] = total_budget_data_activities
        else:
            total_budget_data_activities = sum(
            entry[appr_key]
            for entry in data3
            if entry["obj"] == obj
            and entry["Type"] == 'GJ'
      
    
            )
            item["total_budget"] = -(total_budget_data_activities)
        
        item["ytd_budget"] =  item["total_budget"] * ytd_budget
        total_expense += item["total_budget"]  
        total_expense_ytd_budget += item[f"ytd_budget"]
        if category == "Payroll and Benefits":
            total_budget_pc += item["total_budget"]                
        if category == "Professional and Contract Services":          
            total_budget_pcs += item["total_budget"] 
        if category == "Materials and Supplies":       
            total_budget_sm += item["total_budget"]     
            
        if category == "Other Operating Costs":
            total_budget_ooe += item["total_budget"]  
        if category == "Depreciation":  
            total_budget_oe += item["total_budget"]     
            
        if category == "Debt Services": 
            total_budget_te += item["total_budget"]     
        if category == "FIXED/CAPITAL ASSETS": 
            total_budget_cpa += item["total_budget"]      
        for i, acct_per in enumerate(acct_per_values, start=1):
            total_activities = sum(
                entry[expense_key]
                for entry in data3
                if entry["obj"] == obj and entry["AcctPer"] == acct_per
            )
            total_adjustment = sum(
                entry[expense_key]
                for entry in adjustment
                if entry["obj"] == obj 
                and entry["AcctPer"] == acct_per 
                and entry["School"] == school
                and entry[expense_key] is not None 
                and not isinstance(entry[expense_key], str)
            )
            item[f"total_activities{i}"] = total_activities + total_adjustment
            
            if category == "Payroll and Benefits":
                total_EOC_pc[acct_per] += item[f"total_activities{i}"]
            
            elif category == "Professional and Contract Services":
                total_EOC_pcs[acct_per] += item[f"total_activities{i}"]
            elif category == "Materials and Supplies":
                total_EOC_sm[acct_per] += item[f"total_activities{i}"]
            
            elif category == "Other Operating Costs":
                total_EOC_ooe[acct_per] += item[f"total_activities{i}"]
            elif category == "Depreciation":
                total_EOC_oe[acct_per] += item[f"total_activities{i}"]
            
            elif category == "Debt Services":
                total_EOC_te[acct_per] += item[f"total_activities{i}"]
            elif category == "FIXED/CAPITAL ASSETS":
                total_EOC_cpa[acct_per] += item[f"total_activities{i}"]
            total_expense_months[acct_per] += item[f"total_activities{i}"]  
        for month_number in range(1, 13):
            if month_number != month_exception:
                ytd_total += (item[f"total_activities{month_number}"])
        item["ytd_total"] = ytd_total
    total_expense += dna_total
    total_expense_ytd_budget += ytd_ammended_dna
    for acct_per, dna_value in dna_total_months.items():

        if acct_per in total_expense_months:
        
            total_expense_months[acct_per] += dna_value
    # ytd_EOC_pc  = sum(total_EOC_pc.values())
    # ytd_EOC_pcs = sum(total_EOC_pcs.values())
    # ytd_EOC_sm  = sum(total_EOC_sm.values())
    # ytd_EOC_ooe = sum(total_EOC_ooe.values())
    # ytd_EOC_te  = sum(total_EOC_te.values())
    # ytd_EOC_oe  = sum(total_EOC_oe.values())
    # ytd_EOC_cpa  = sum(total_EOC_cpa.values())
    ytd_EOC_pc  = abs(sum(value for key, value in total_EOC_pc.items() if key != month_exception_str))
    ytd_EOC_pcs =  abs(sum(value for key, value in total_EOC_pcs.items() if key != month_exception_str))
    ytd_EOC_sm  =  abs(sum(value for key, value in total_EOC_sm.items() if key != month_exception_str))
    ytd_EOC_ooe =  abs(sum(value for key, value in total_EOC_ooe.items() if key != month_exception_str))
    ytd_EOC_te  =  abs(sum(value for key, value in total_EOC_te.items() if key != month_exception_str))
    ytd_EOC_oe  =  abs(sum(value for key, value in total_EOC_oe.items() if key != month_exception_str))
    ytd_EOC_cpa  =  abs(sum(value for key, value in total_EOC_cpa.items() if key != month_exception_str))
    
    ytd_budget_pc = total_budget_pc * ytd_budget
    ytd_budget_pcs = total_budget_pcs * ytd_budget
    ytd_budget_sm = total_budget_sm * ytd_budget
    ytd_budget_ooe = total_budget_ooe  * ytd_budget
    ytd_budget_oe = total_budget_oe * ytd_budget
    ytd_budget_te = total_budget_te * ytd_budget
    ytd_budget_cpa = total_budget_cpa * ytd_budget
    #temporarily for 6500
    budget_for_6500 = 0
    ytd_budget_for_6500 = 0 
    for item in data_expensebyobject:
        obj = item["obj"]
    
        if obj == "6100":
            category = "Payroll and Benefits"
            item["variances"] = ytd_budget_pc - ytd_EOC_pc
            item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_pc / total_budget_pc*100))) if total_budget_pc != 0 else ""
        elif obj == "6200":
            category = "Professional and Contract Services"
            item["variances"] = ytd_budget_pcs - ytd_EOC_pcs
            item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_pcs / total_budget_pcs*100))) if total_budget_pcs != 0 else ""
        elif obj == "6300":
            category = "Materials and Supplies"
            item["variances"] = ytd_budget_sm - ytd_EOC_sm
            item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_sm / total_budget_sm*100))) if total_budget_sm != 0 else ""
        elif obj == "6400":
            category = "Other Operating Costs"
            item["variances"] = ytd_budget_ooe - ytd_EOC_ooe
            item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_ooe / total_budget_ooe*100))) if total_budget_ooe != 0 else ""
        elif obj == "6449":
            category = "Depreciation"
            item["variances"] = ytd_budget_oe - ytd_EOC_oe
            item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_oe / total_budget_oe*100))) if total_budget_oe != 0 else ""
        elif obj == "6500":
            category = "Debt Services"
            item["variances"] = ytd_budget_te - ytd_EOC_te
            item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_te / total_budget_te*100))) if total_budget_te != 0 else ""
        else:
            category = "FIXED/CAPITAL ASSETS"
            item["variances"] = ytd_budget_cpa - ytd_EOC_cpa
            item["var_EOC"] = "{:d}%".format(round(abs(ytd_EOC_cpa / total_budget_cpa*100))) if total_budget_cpa != 0 else ""
        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f"total_expense{i}"] = sum(
                entry[f"total_activities{i}"]
                for entry in data_activities
                if entry["Category"] == category
            )
        
    #CONTINUATION COMPUTATION TOTAL EXPENSE
    total_expense_ytd = sum([ytd_EOC_te, ytd_EOC_ooe, ytd_EOC_sm, ytd_EOC_pcs, ytd_EOC_pc,dna_ytd_total,ytd_EOC_cpa])
    variances_total_expense = total_expense_ytd_budget - total_expense_ytd
    var_total_expense = "{:d}%".format(round(abs(total_expense_ytd / total_expense*100))) if total_expense != 0 else ""
        
    #CALCULATIONS START NET INCOME
    net_income_budget = 0
    ytd_budget_net_income = 0 
    total_net_income_months =  {acct_per: 0 for acct_per in acct_per_values}
    ytd_net_income = 0
    variances_net_income = 0
    var_net_income = 0 
    budget_net_income = totals["total_ammended"] - total_expense

    ytd_budget_net_income = ytd_ammended_total - total_expense_ytd_budget
    ytd_net_income = ytd_total_revenue - total_expense_ytd
    variances_net_income = ytd_net_income - ytd_budget_net_income
    var_net_income = "{:d}%".format(round(abs(ytd_net_income / budget_net_income * 100))) if budget_net_income != 0 else ""

    
    total_net_income_months = {
        acct_per: abs(total_revenue[acct_per]) - total_expense_months[acct_per]
        for acct_per in acct_per_values
    }  
    #FORMAT FOR REVENUE
    formatted_ammended = format_value_dollars(totals["total_ammended"]) if totals["total_ammended"] != 0 else ""
    formatted_ammended_lr = format_value_dollars(totals["total_ammended_lr"]) if totals["total_ammended_lr"] != 0 else ""
    formatted_ammended_spr = format_value(totals["total_ammended_spr"]) if totals["total_ammended_spr"] != 0 else ""
    formatted_ammended_fpr = format_value(totals["total_ammended_fpr"]) if totals["total_ammended_fpr"] != 0 else ""
        
    formatted_total_lr = {acct_per: format_value_dollars_negative(value) for acct_per, value in total_lr.items() if value != 0}
    formatted_total_spr = {acct_per: format_value_negative(value) for acct_per, value in total_spr.items() if value != 0}
    formatted_total_fpr = {acct_per: format_value_negative(value) for acct_per, value in total_fpr.items() if value != 0}
    formatted_total_revenue = {acct_per: format_value_dollars_negative(value) for acct_per, value in total_revenue.items() if value != 0}
    
    ytd_ammended_total = format_value_dollars(ytd_ammended_total )
    ytd_ammended_total_lr =format_value_dollars(ytd_ammended_total_lr ) 
    ytd_ammended_total_spr=format_value(ytd_ammended_total_spr) 
    ytd_ammended_total_fpr=format_value(ytd_ammended_total_fpr) 
    
    ytd_total_revenue = format_value_dollars(ytd_total_revenue)
    ytd_total_lr  = format_value_dollars_negative(ytd_total_lr)
    ytd_total_spr = format_value_negative(ytd_total_spr)
    ytd_total_fpr = format_value_negative(ytd_total_fpr)
    variances_revenue = format_value_dollars(variances_revenue)
    variances_revenue_lr=format_value_dollars_negative(variances_revenue_lr)
    variances_revenue_spr=format_value_negative(variances_revenue_spr)
    variances_revenue_fpr=format_value_negative(variances_revenue_fpr)


    for row in data:
        ytd_total = float(row["ytd_total"])
    
        variances =float(row["variances"])
        total_budget = float(row["total_budget"])

        if total_budget is None or total_budget == 0:
            row["total_budget"] = ""
        else:
            row["total_budget"] = format_value(total_budget) 
        if ytd_total is None or ytd_total == 0:
            row["ytd_total"] = ""
        else:
            row["ytd_total"] = format_value_negative(ytd_total)
        if variances is None or variances == 0:
            row["variances"] = ""
        else:
            row["variances"] = format_value(variances)
    # FOR EXPENSE BY OBJECT DEPRECIATION ONLY        
    dna_total_6449 = format_value(dna_total)
    ytd_ammended_dna_6449 = format_value(ytd_ammended_dna)
    dna_ytd_total_6449 = format_value(dna_ytd_total)
    variances_dna_6449 = format_value(variances_dna)   
    dna_total_months_6449 = {acct_per: format_value(value) for acct_per, value in dna_total_months.items() if value != 0}    
    #FORMAT FIRST TOTAL AND DEPRECIATION AND AMORTIZATION(DNA)
    dna_total = format_value_dollars(dna_total)
    first_total = format_value_dollars(first_total)
    ytd_ammended_dna = format_value_dollars(ytd_ammended_dna)
    ytd_ammended_total_first = format_value_dollars(ytd_ammended_total_first)
    
    dna_ytd_total = format_value_dollars(dna_ytd_total)
    first_ytd_total = format_value_dollars(first_ytd_total)
    variances_first_total = format_value_dollars(variances_first_total)
    variances_dna = format_value_dollars(variances_dna)
    first_total_months = {acct_per: format_value_dollars(value) for acct_per, value in first_total_months.items() if value != 0}
    dna_total_months = {acct_per: format_value_dollars(value) for acct_per, value in dna_total_months.items() if value != 0}
    for row in data2:
        ytd_budget =float(row[f"ytd_budget"])
        ytd_total = float(row["ytd_total"])
        variances = float(row["variances"])
        budget = row["total_budget"]
        
        
        if ytd_total is None or ytd_total == 0:
            row[f"ytd_total"] = ""
        else:
            row[f"ytd_total"] = format_value(ytd_total) 
        if var_ytd is None or var_ytd == 0:
            row[f"variances"] = ""
        else:
            row[f"variances"] = format_value(variances)
        
            
        if budget is None or budget == 0:
            row[f"total_budget"] = ""
        else:
            row[f"total_budget"] = format_value(budget)
        if ytd_budget is None or ytd_budget == 0:
            row[f"ytd_budget"] = ""
        else:
            row[f"ytd_budget"] = format_value(ytd_budget)

    #FORMAT SURPLUS BEFORE DEFICIT   
    ammended_budget_SBD = format_value_dollars(ammended_budget_SBD)
    ytd_ammended_SBD = format_value_dollars(ytd_ammended_SBD)
    ytd_SBD = format_value_dollars(ytd_SBD)
    variances_SBD = format_value_dollars(variances_SBD)
    
    total_SBD = {acct_per: format_value_dollars(value) for acct_per, value in total_SBD.items() if value != 0}
    
    #FORMAT NET SURPLUS 
    ammended_budget_netsurplus = format_value_dollars(ammended_budget_netsurplus)
    ytd_ammended_netsurplus = format_value_dollars(ytd_ammended_netsurplus)
    ytd_netsurplus = format_value_dollars(ytd_netsurplus)
    variances_netsurplus = format_value_dollars(variances_netsurplus)
    
    total_netsurplus_months = {acct_per: format_value_dollars(value) for acct_per, value in total_netsurplus_months.items() if value != 0}
    
    #FORMAT EXPENSE BY OBJECT CODES
    for row in data_activities:
    
        ytd_total = (row["ytd_total"])
        total_expense_budget = row["total_budget"]
        ytd_budget = row["ytd_budget"]
    
    
        if ytd_total is None or ytd_total == 0:
            row[f"ytd_total"] = ""
        else:
            row[f"ytd_total"] = format_value(ytd_total)
        if total_expense_budget is None or total_expense_budget == 0:
            row[f"total_budget"] = ""
        else:
            row[f"total_budget"] = format_value(total_expense_budget)
        if ytd_budget is None or ytd_budget == 0:
            row[f"ytd_budget"] = ""
        else:
            row[f"ytd_budget"] = format_value(ytd_budget)
    
    for row in data_expensebyobject:
        variances = float(row["variances"])
        if variances is None or variances == 0:
            row[f"variances"] = ""
        else:
            row[f"variances"] = format_value(variances)
    total_EOC_pc = {acct_per: format_value(value) for acct_per, value in total_EOC_pc.items() if value != 0}
    total_EOC_pcs = {acct_per: format_value(value) for acct_per, value in total_EOC_pcs.items() if value != 0} 
    total_EOC_sm = {acct_per: format_value(value) for acct_per, value in total_EOC_sm.items() if value != 0} 
    total_EOC_ooe = {acct_per: format_value(value) for acct_per, value in total_EOC_ooe.items() if value != 0} 
    total_EOC_te = {acct_per: format_value(value) for acct_per, value in total_EOC_te.items() if value != 0}
    total_EOC_oe = {acct_per: format_value(value) for acct_per, value in total_EOC_oe.items() if value != 0}
    total_EOC_cpa = {acct_per: format_value(value) for acct_per, value in total_EOC_cpa.items() if value != 0}  
    ytd_EOC_pc  = format_value(ytd_EOC_pc)
    ytd_EOC_pcs = format_value(ytd_EOC_pcs)
    ytd_EOC_sm  = format_value(ytd_EOC_sm)
    ytd_EOC_ooe = format_value(ytd_EOC_ooe)
    ytd_EOC_te  = format_value(ytd_EOC_te)
    ytd_EOC_oe  = format_value(ytd_EOC_oe)
    ytd_EOC_cpa  = format_value(ytd_EOC_cpa)
    total_budget_pc =  format_value(total_budget_pc)
    total_budget_pcs = format_value(total_budget_pcs)
    total_budget_sm =  format_value(total_budget_sm)
    total_budget_ooe = format_value(total_budget_ooe)
    total_budget_oe =  format_value(total_budget_oe)   
    total_budget_te =  format_value(total_budget_te)
    total_budget_cpa =  format_value(total_budget_cpa)
    ytd_budget_pc = format_value(ytd_budget_pc)
    ytd_budget_pcs =format_value(ytd_budget_pcs)
    ytd_budget_sm = format_value(ytd_budget_sm)
    ytd_budget_ooe =format_value(ytd_budget_ooe)
    ytd_budget_oe = format_value(ytd_budget_oe)
    ytd_budget_te = format_value(ytd_budget_te)
    ytd_budget_cpa = format_value(ytd_budget_cpa)
    #EXPENSE OBJECT FOR FIX
    budget_for_6500 = format_value(budget_for_6500)
    ytd_budget_for_6500 = format_value(ytd_budget_for_6500)
    #FORMAT TOTAL EXPENSE
    total_expense = format_value_dollars(total_expense)
    total_expense_ytd_budget = format_value_dollars(total_expense_ytd_budget)
    total_expense_months = {acct_per: format_value_dollars(value) for acct_per, value in total_expense_months.items() if value != 0} 
    total_expense_ytd = format_value_dollars(total_expense_ytd)
    variances_total_expense =format_value_dollars(variances_total_expense)
        
    #FORMAT NET INCOME
    budget_net_income = format_value_dollars(budget_net_income)
    ytd_budget_net_income = format_value_dollars(ytd_budget_net_income)
    total_net_income_months = {acct_per: format_value_dollars(value) for acct_per, value in total_net_income_months.items() if value != 0}
    ytd_net_income = format_value_dollars(ytd_net_income)
    variances_net_income = format_value_dollars(variances_net_income)     
    keys_to_check = [
        "total_real1",
        "total_real2",
        "total_real3",
        "total_real4",
        "total_real5",
        "total_real6",
        "total_real7",
        "total_real8",
        "total_real9",
        "total_real10",
        "total_real11",
        "total_real12",
    ]
    keys_to_check_expense = [
        "total_activities1",
        "total_activities2",
        "total_activities3",
        "total_activities4",
        "total_activities5",
        "total_activities6",
        "total_activities7",
        "total_activities8",
        "total_activities9",
        "total_activities10",
        "total_activities11",
        "total_activities12",
    ]
    keys_to_check_expense2 = [
        "total_expense1",
        "total_expense2",
        "total_expense3",
        "total_expense4",
        "total_expense5",
        "total_expense6",
        "total_expense7",
        "total_expense8",
        "total_expense9",
        "total_expense10",
        "total_expense11",
        "total_expense12",
    ]
    
    for row in data_activities:
        for key in keys_to_check_expense:
            value = float(row[key])
            if value == 0:
                row[key] = ""
            elif value < 0:
                row[key] = "({:,.0f})".format(abs(float(row[key])))
            elif value != "":
                row[key] = "{:,.0f}".format(float(row[key]))
    for row in data_expensebyobject:
        for key in keys_to_check_expense2:
            value = float(row[key])
            if value == 0:
                row[key] = ""
            elif value < 0:
                row[key] = "({:,.0f})".format(abs(float(row[key])))
            elif value != "":
                row[key] = "{:,.0f}".format(float(row[key]))
        # for row in data_activitybs:
        # for key in keys_to_check:
        #     value = int(row[key])
        #     if value == 0:
        #         row[key] = ""
        #     elif value < 0:
        #         row[key] = "({:,.0f})".format(abs(float(row[key])))
        #     elif value != "":
        #         row[key] = "{:,.0f}".format(float(row[key]))
    for row in data:
        for key in keys_to_check:
            value = int(row[key])
            if value == 0:
                row[key] = ""
            elif value < 0:
                row[key] = "{:,.0f}".format(abs(float(row[key])))
            elif value != "":
                row[key] = "({:,.0f})".format(float(row[key]))
    # for row in data:
    #     for key in keys_to_check:
    #         if row[key] < 0:
    #             row[key] = -row[key]
    #         else if row[key] > 0:
    #             row[key] = row[key]
    # for row in data:
    #     for key in keys_to_check:
    #         if row[key] != "":
    #             row[key] = "{:,.0f}".format(row[key])
    keys_to_check_func = [
        "total_func1",
        "total_func2",
        "total_func3",
        "total_func4",
        "total_func5",
        "total_func6",
        "total_func7",
        "total_func8",
        "total_func9",
        "total_func10",
        "total_func11",
        "total_func12",
    ]
    keys_to_check_func_2 = [
        "total_func2_1",
        "total_func2_2",
        "total_func2_3",
        "total_func2_4",
        "total_func2_5",
        "total_func2_6",
        "total_func2_7",
        "total_func2_8",
        "total_func2_9",
        "total_func2_10",
        "total_func2_11",
        "total_func2_12",
    ]
    for row in data2:
        for key in keys_to_check_func:
            if key in row and row[key] is not None:
                row[key] = format_value(row[key])
            else:
                row[key] = ""
    for row in data2:
        for key in keys_to_check_func_2:
            if key in row and row[key] is not None :
                row[key] = format_value(row[key])
            else:
                row[key] = ""
    # if not school == "village-tech":
    #     lr_funds = list(set(row["fund"] for row in data3 if "fund" in row))
    #     lr_funds_sorted = sorted(lr_funds)
    #     lr_obj = list(set(row["obj"] for row in data3 if "obj" in row))
    #     lr_obj_sorted = sorted(lr_obj)
    #
    #     func_choice = list(set(row["func"] for row in data3 if "func" in row))
    #     func_choice_sorted = sorted(func_choice)
    # current_date = datetime.today().date()
    # current_year = current_date.year
    # last_year = current_date - timedelta(days=365)
    # current_month = current_date.replace(day=1)
    # last_month = current_month - relativedelta(days=1)
    # last_month_number = last_month.month
    # ytd_budget_test = last_month_number + 4
    # ytd_budget = ytd_budget_test / 12
    
    sorted_data2 = sorted(data2, key=lambda x: x['func_func'])
    sorted_data = sorted(data, key=lambda x: x['obj'])
    data_activities = sorted(data_activities, key=lambda x: x['obj'])
    

    context = {
        "data": sorted_data,
        "data2": sorted_data2,
        "data3": data3,
        "data_expensebyobject": data_expensebyobject,
        "data_activities": data_activities,
        "months":
                {
            "last_month": formatted_last_month,
            "last_month_number": last_month_number,
            "last_month_name": last_month_name,
            "format_ytd_budget": formatted_ytd_budget,
            "ytd_budget": ytd_budget,
            "FY_year_1":FY_year_1,
            "FY_year_2":FY_year_2,
            "db_last_month": db_last_month,
            "month_exception": month_exception,
            "month_exception_str": month_exception_str,
            },
        "totals":{
            #FOR REVENUES
            "total_lr": formatted_total_lr,
            "total_spr": formatted_total_spr,
            "total_fpr": formatted_total_fpr,
            "total_revenue": formatted_total_revenue,
            "total_ammended": formatted_ammended,
            "total_ammended_lr": formatted_ammended_lr,
            "total_ammended_spr": formatted_ammended_spr,
            "total_ammended_fpr": formatted_ammended_fpr,
            "ytd_ammended_total":ytd_ammended_total,
            "ytd_ammended_total_lr":ytd_ammended_total_lr,
            "ytd_ammended_total_spr":ytd_ammended_total_spr,
            "ytd_ammended_total_fpr":ytd_ammended_total_fpr,
            "ytd_total_revenue": ytd_total_revenue,
            "ytd_total_lr": ytd_total_lr,
            "ytd_total_spr": ytd_total_spr,
            "ytd_total_fpr": ytd_total_fpr,
            "variances_revenue":variances_revenue,
            "variances_revenue_lr":variances_revenue_lr,
            "variances_revenue_spr":variances_revenue_spr,
            "variances_revenue_fpr":variances_revenue_fpr,
            "var_ytd":var_ytd,
            "var_ytd_lr":var_ytd_lr,
            "var_ytd_spr":var_ytd_spr,
            "var_ytd_fpr":var_ytd_fpr,
            #FIRST TOTAL
            "first_total":first_total,
            "first_total_months":first_total_months,
            "first_ytd_total":first_ytd_total,
            "ytd_ammended_total_first": ytd_ammended_total_first,
            "variances_first_total":variances_first_total,
            "var_ytd_first_total": var_ytd_first_total,
            # DEPRECIATION AND AMORTIZATION
            "dna_total":dna_total,
            "dna_total_months":dna_total_months,
            "dna_ytd_total":dna_ytd_total,
            "ytd_ammended_dna": ytd_ammended_dna,
            "variances_dna":variances_dna,
            "var_ytd_dna":var_ytd_dna,
            #SURPLUS BEFORE DEFICIT(SBD)
            "total_SBD": total_SBD,
            "ammended_budget_SBD": ammended_budget_SBD,
            "ytd_ammended_SBD": ytd_ammended_SBD,
            "ytd_SBD":ytd_SBD,
            "variances_SBD": variances_SBD,
            "var_SBD":var_SBD,
            #NET SURPLUS    
            "total_netsurplus_months": total_netsurplus_months,
            "ammended_budget_netsurplus": ammended_budget_netsurplus,
            "ytd_ammended_netsurplus" : ytd_ammended_netsurplus,
            "ytd_netsurplus": ytd_netsurplus,
            "variances_netsurplus": variances_netsurplus,
            "var_netsurplus":var_netsurplus,
            #EXPENSE BY OBJECT 
            "total_EOC_pc":total_EOC_pc,
            "total_EOC_pcs":total_EOC_pcs,
            "total_EOC_sm":total_EOC_sm,
            "total_EOC_ooe":total_EOC_ooe,
            "total_EOC_te":total_EOC_te,
            "total_EOC_oe":total_EOC_oe,
            "total_EOC_cpa":total_EOC_cpa,
            "ytd_EOC_pc":ytd_EOC_pc,
            "ytd_EOC_pcs":ytd_EOC_pcs,
            "ytd_EOC_sm":ytd_EOC_sm,
            "ytd_EOC_ooe":ytd_EOC_ooe,
            "ytd_EOC_te":ytd_EOC_te,
            "ytd_EOC_oe":ytd_EOC_oe,
            "ytd_EOC_cpa":ytd_EOC_cpa,
            "total_budget_pc":total_budget_pc,
            "total_budget_pcs":total_budget_pcs,
            "total_budget_sm":total_budget_sm,
            "total_budget_ooe":total_budget_ooe,
            "total_budget_oe":total_budget_oe,
            "total_budget_te":total_budget_te,
            "total_budget_cpa":total_budget_cpa,
            "ytd_budget_pc":ytd_budget_pc,
            "ytd_budget_pcs":ytd_budget_pcs,
            "ytd_budget_sm":ytd_budget_sm,
            "ytd_budget_ooe":ytd_budget_ooe,
            "ytd_budget_oe":ytd_budget_oe,
            "ytd_budget_te":ytd_budget_te,
            "ytd_budget_cpa":ytd_budget_cpa,
            #EXPENSE BY OBJECT FOR DEPRECIATION AND AMORTIZATION
            "dna_total_6449":dna_total_6449,
            "ytd_ammended_dna_6449":ytd_ammended_dna_6449,
            "dna_ytd_total_6449":dna_ytd_total_6449,
            "variances_dna_6449":variances_dna_6449,
            "dna_total_months_6449":dna_total_months_6449,
            #FIX SOON
            "budget_for_6500":budget_for_6500,
            "ytd_budget_for_6500": ytd_budget_for_6500,
            
            #TOTAL EXPENSE 
            "total_expense": total_expense,
            "total_expense_ytd_budget": total_expense_ytd_budget,
            "total_expense_months":total_expense_months,
            "total_expense_ytd":total_expense_ytd,
            "variances_total_expense":variances_total_expense,
            "var_total_expense":var_total_expense,
            #NET INCOME
            "budget_net_income": budget_net_income,
            "ytd_budget_net_income":ytd_budget_net_income,
            "total_net_income_months":total_net_income_months,
            "variances_net_income": variances_net_income,
            "ytd_net_income": ytd_net_income,
            "var_net_income":var_net_income,
            #FOR BS
            "bs_ytd_netsurplus":bs_ytd_netsurplus,
        }
    }
    # if not school == "village-tech":
    #     context["total_lr"] = formatted_total_lr,
        # context["total_netsurplus"] = formatted_total_netsurplus
        # context["total_SBD"] = total_SBD
        # context["ytd_netsurplus"] = formated_ytdnetsurplus
    # if not school == "village-tech":
    #     context["lr_funds"] = lr_funds_sorted
    #     context["lr_obj"] = lr_obj_sorted
    #     context["func_choice"] = func_choice_sorted
    # dict_keys = ["data", "data2", "data3", "data_expensebyobject", "data_activities"]
    
    
    relative_path = os.path.join("profit-loss-date", school)
    
    json_path = os.path.join(JSON_DIR, relative_path)
    
    os.makedirs(json_path, exist_ok=True)
    for key, val in context.items():
        file_path = os.path.join(json_path, f"{key}.json")
        

        with open(file_path, "w") as file:
            json.dump(val, file)
        print(file_path)

def balance_sheet(school,year):
    print("balance")
    present_date = datetime.today().date()   
    present_year = present_date.year
    
    today_date = datetime.now()
    
    today_month = today_date.month

    if year:
        year = int(year)
        if year == present_year:
            
            print("year",year)

            if school in schoolMonths["septemberSchool"]:
                if today_month <= 8:
                    
                    start_year = year - 1
                    present_year = present_year - 1
                    FY_year_current = year - 1
                else: 
                    start_year = year 
                    FY_year_current = year
            else:
                if today_month <= 6:
                    start_year = year - 1
                    present_year = present_year - 1
                    FY_year_current = year - 1
                else: 
                    start_year = year 
                    FY_year_current = year
        else:
            start_year = year 
            FY_year_current = year
        
        if school in schoolMonths["julySchool"]:
            current_date = datetime(start_year, 7, 1).date()
            
        else:
            current_date = datetime(start_year, 9, 1).date() 
        current_year = current_date.year
    else:
        start_year = 2021
        current_date = datetime.today().date()   
        current_year = current_date.year
        FY_year_current = current_year

    while start_year <= FY_year_current:
        print(start_year)
        FY_year_1 = start_year
        FY_year_2 = start_year + 1 
        start_year = FY_year_2

        cnxn = connect()
        cursor = cnxn.cursor()
        # cursor.execute(f"SELECT  * FROM [dbo].{db[school]['bs']} AS T1 LEFT JOIN [dbo].{db[school]['bs_fye']} AS T2 ON T1.BS_id = T2.BS_id ;  ")
        # rows = cursor.fetchall()
        # bs_checker = []

        # for row in rows:
        #     if row[8] == school:
        #         if FY_year_1 == row[9]:
        #             row_dict = {
        #                 "Activity": row[0],
        #                 "Description": row[1],
        #                 "Category": row[2],
        #                 "Subcategory": row[3],
        #                 "FYE": row[4],
        #                 "BS_id": row[5],
        #                 "school": row[8],
                
        #             }
        #             bs_checker.append(row_dict)
                    
            
        cursor.execute(f"SELECT  * FROM [dbo].{db[school]['bs']} ;  ")
        rows = cursor.fetchall()

        unique_bs_id = []
        for row in rows:
            if row[5] not in unique_bs_id:
                unique_bs_id.append(row[5])
        
        largest_unique_id = max(unique_bs_id)
     
        for i in range(1, int(largest_unique_id)):
            cursor.execute(f"SELECT  * FROM [dbo].{db[school]['bs_fye']} where  BS_id = {i} and school = '{school}' and year = '{FY_year_1}';  ")
            row = cursor.fetchone()
            if row is None:
                fye = '0'
                query = "INSERT INTO [dbo].[BS_FYE] (BS_id, FYE, school,year) VALUES (?, ?, ?,?)"
                cursor.execute(query, (i, fye, school,FY_year_1)) 
                print(f"Data Inserted to BS_FYE DB  with BS_id = {i} and year = {FY_year_1}")
                cnxn.commit()





        # if bs_checker == []:
        #     for i in range(1, int(largest_unique_id)):
        #         fye = '0'
        #         query = "INSERT INTO [dbo].[BS_FYE] (BS_id, FYE, school,year) VALUES (?, ?, ?,?)"
        #         cursor.execute(query, (i, fye, school,FY_year_1)) 
        #         cnxn.commit()
                
        
        cursor.execute(f"SELECT  * FROM [dbo].{db[school]['bs']} AS T1 LEFT JOIN [dbo].{db[school]['bs_fye']} AS T2 ON T1.BS_id = T2.BS_id ;  ")
        rows = cursor.fetchall()

        data_balancesheet = []

        for row in rows:
            if row[8] == school:
                fye = float(row[7]) if row[7] else 0
                if fye == 0:
                    fyeformat = ""
                else:
                    if row[0] == 'Cash' or row[0] == 'AP':
                        fyeformat = (
                            "${:,.0f}".format(abs(fye)) if fye >= 0 else "$({:,.0f})".format(abs(fye))
                        )
                    else:
                        fyeformat = (
                            "{:,.0f}".format(abs(fye)) if fye >= 0 else "({:,.0f})".format(abs(fye))
                        )

                if FY_year_1 == row[9]:
                    print(row[0])
                    row_dict = {
                        "Activity": row[0],
                        "Description": row[1],
                        "Category": row[2],
                        "Subcategory": row[3],
                        "FYE": fyeformat, #should now be total fye coming from GL(data3)
                        "BS_id": row[5], #wont be used
                        "school": row[8],
                

                    }

                    data_balancesheet.append(row_dict)

        # cursor.execute(f"SELECT  * FROM [dbo].{db[school]['object']};")
        # rows = cursor.fetchall()
        #
        # data = []
        if FY_year_1 == present_year:
            relative_path = os.path.join("profit-loss", school)
        else:
            relative_path = os.path.join(str(FY_year_1), "profit-loss", school)

        # json_path = JSON_DIR.path(relative_path)
        json_path = os.path.join(JSON_DIR, relative_path)
        with open(os.path.join(json_path, "data.json"), "r") as f:
            data = json.load(f)
        # for row in rows:
        #     if row[4] is None:
        #         row[4] = ""
        #     valueformat = "{:,.0f}".format(float(row[4])) if row[4] else ""
        #     row_dict = {
        #         "fund": row[0],
        #         "obj": row[1],
        #         "description": row[2],
        #         "category": row[3],
        #         "value": valueformat,
        #     }
        #     data.append(row_dict)

        # cursor.execute(f"SELECT  * FROM [dbo].{db[school]['function']};")
        # rows = cursor.fetchall()
        #
        # data2 = []
        with open(os.path.join(json_path, "data2.json"), "r") as f:
            data2 = json.load(f)
        # for row in rows:
        #     budgetformat = "{:,.0f}".format(float(row[3])) if row[3] else ""
        #     row_dict = {
        #         "func_func": row[0],
        #         "desc": row[1],
        #         "category": row[2],
        #         "obj": row[4],
        #         "budget": budgetformat,
        #     }
        #     data2.append(row_dict)

        cursor.execute(f"SELECT * FROM [dbo].{db[school]['bs_activity']}")
        rows = cursor.fetchall()

        data_activitybs = []

        for row in rows:
            if row[3] == school:
                row_dict = {
                    "Activity": row[0],
                    "obj": row[1],
                    "Description2": row[2],
                    "school": row[3],
                }
        
                data_activitybs.append(row_dict)

        with open(os.path.join(json_path, "data3.json"), "r") as f:
            data3 = json.load(f)
        # if not school == "village-tech":
        #     cursor.execute(
        #         f"SELECT * FROM [dbo].{db[school]['db']}  as AA where AA.Number != 'BEGBAL'"
        #     )
        # else:
        #     cursor.execute(
        #         f"SELECT * FROM [dbo].{db[school]['db']}"
        #     )
        # rows = cursor.fetchall()
        #
        # data3 = []
        #
        # if not school == "village-tech":
        #     for row in rows:
        #         row_dict = {
        #             "fund": row[0],
        #             "func": row[1],
        #             "obj": row[2],
        #             "sobj": row[3],
        #             "org": row[4],
        #             "fscl_yr": row[5],
        #             "pgm": row[6],
        #             "edSpan": row[7],
        #             "projDtl": row[8],
        #             "AcctDescr": row[9],
        #             "Number": row[10],
        #             "Date": row[11],
        #             "AcctPer": row[12],
        #             "Est": row[13],
        #             "Real": row[14],
        #             "Appr": row[15],
        #             "Encum": row[16],
        #             "Expend": row[17],
        #             "Bal": row[18],
        #             "WorkDescr": row[19],
        #             "Type": row[20],
        #             "Contr": row[21],
        #         }
        #
        #         data3.append(row_dict)
        # else:
        #     for row in rows:
        #         amount = float(row[19])
        #         row_dict = {
        #             "fund": row[0],
        #             "func": row[2],
        #             "obj": row[3],
        #             "sobj": row[4],
        #             "org": row[5],
        #             "fscl_yr": row[6],
        #             "Date": row[9],
        #             "AcctPer": row[10],
        #             "Amount": amount,
        #         }
        #         data3.append(row_dict)

        with open(os.path.join(json_path, "totals.json"), "r") as f:
            totals = json.load(f)

        with open(os.path.join(json_path, "months.json"), "r") as f:
            months = json.load(f)

        # FY_year_1 = months["FY_year_1"] 
        # FY_year_2 = months["FY_year_2"]
        last_month = months["last_month"]
        last_month_number = months["last_month_number"]
        last_month_name = months["last_month_name"]
   
        db_last_month = months["db_last_month"]
        month_exception = months["month_exception"]
        month_exception_str = months["month_exception_str"]
        

     
        cursor.execute(f"SELECT * FROM [dbo].{db[school]['adjustment']} ")
        rows = cursor.fetchall()

        adjustment = []

        if school in schoolCategory["ascender"]:
            for row in rows:
                expend = float(row[17])
                row_dict = {
                    "fund": row[0],
                    "func": row[1],
                    "obj": row[2],
                    "sobj": row[3],
                    "org": row[4],
                    "fscl_yr": row[5],
                    "pgm": row[6],
                    "edSpan": row[7],
                    "projDtl": row[8],
                    "AcctDescr": row[9],
                    "Number": row[10],
                    "Date": row[11],
                    "AcctPer": row[12],
                    "Est": row[13],
                    "Real": row[14],
                    "Appr": row[15],
                    "Encum": row[16],
                    "Expend": expend,
                    "Bal": row[18],
                    "WorkDescr": row[19],
                    "Type": row[20],
                    "School": row[21],
                }
                adjustment.append(row_dict)

        acct_per_values = [
            "01",
            "02",
            "03",
            "04",
            "05",
            "06",
            "07",
            "08",
            "09",
            "10",
            "11",
            "12",
        ]

        real_key = "Real"        
        bal_key = "Bal"
        expend_key = "Expend"
        begbal_key = "BegBal"
        if school in schoolCategory["skyward"]:
            bal_key = "Amount"
            real_key = "Amount"
            expend_key = "Amount"
            begbal_key = "BegBal"





        unique_act = []
        for item in data_balancesheet:
            Activity = item["Activity"]
            

            if item['Subcategory'] == 'Long Term Debt' or  item['Subcategory'] == 'Current Liabilities' or item['Category'] == 'Net Assets':
                if Activity not in unique_act:
                    unique_act.append(Activity)

        # if school == 'goldenrule':
        #     numberstack = set()

        #     for item in data3:
        #         if item["obj"] == '3600':
                    
        #             numberstack.add(item["Number"])
            
        #     numberstack = list(numberstack)
            



                    

        for item in data_activitybs:
            Activity = item["Activity"]
            obj = item["obj"]
            item["fytd"] = 0
            
            for i, acct_per in enumerate(acct_per_values, start=1):
                if school in schoolCategory["ascender"]:
                    total_data3 = sum(
                        entry[bal_key]
                        for entry in data3
                        if entry["obj"] == obj 
                        and entry["AcctPer"] == acct_per
                        and entry["fund"] != '000'
                        and "BEG BAL" not in entry["WorkDescr"]
                        and "BEGBAL" not in entry["WorkDescr"]
                        and "BEGINNING BAL" not in entry["WorkDescr"]
                    )
                else:
                    total_data3 = sum(
                        entry[bal_key]
                        for entry in data3
                        if entry["obj"] == obj 
                        and entry["AcctPer"] == acct_per
                        and entry["fund"] != '000'

                    )
                total_adjustment = sum(
                    entry[bal_key]
                    for entry in adjustment
                    if entry["obj"] == obj
                    and entry["AcctPer"] == acct_per 
                    and entry["School"] == school
                    and entry[bal_key] is not None 
                    and not isinstance(entry[bal_key], str)
                )
             
                item[f"total_bal{i}"] = total_data3 + total_adjustment
                if i != month_exception:
                    item["fytd"] += item[f"total_bal{i}"]

            if school in schoolCategory["skyward"]:
                activity_fye = sum(
                        entry[begbal_key]
                        for entry in data3
                        if entry["obj"] == obj 
                        and entry[begbal_key] is not None                   
                    )

                
                if Activity in unique_act:
                    item["activity_fye"] = -(activity_fye)
                else:
                    item["activity_fye"] = activity_fye
            
            if school in school_fye:
                activity_fye = sum(
                        entry["Bal"]
                        for entry in data3
                        if entry["obj"] == obj
                        and entry["fund"] == '000'
                        and entry["Bal"] is not None                   
                    )
                
                
                if Activity in unique_act:
                    item["activity_fye"] = -(activity_fye)
                else:
                    item["activity_fye"] = activity_fye

            # if school == 'goldenrule':

            #     activity_fye = sum(
            #         entry[bal_key]
            #         for entry in data3
            #         if entry["obj"] == obj
            #         and entry["Type"] == "GJ"
            #         and entry["Number"] in numberstack
            #     )
            #     if Activity in unique_act:
            #         item["activity_fye"] = -(activity_fye)
            #     else:
            #         item["activity_fye"] = activity_fye

        # if school == 'goldenrule':
        #     for item in data_balancesheet:
        #         Activity = item["Activity"]
        #         item["FYE"] = sum(
        #             entry["activity_fye"]
        #             for entry in data_activitybs
        #             if entry["Activity"] == Activity
        #         )
        #         print(Activity,item["FYE"])


                
            
            
        activity_sum_dict = {}


        for item in data_activitybs:
            Activity = item["Activity"]
            
            for i in range(1, 13):
                total_sum_i = sum(
                    float(entry[f"total_bal{i}"])
                    if entry[f"total_bal{i}"] and entry["Activity"] == Activity
                    else 0
                    for entry in data_activitybs
                )
                activity_sum_dict[(Activity, i)] = total_sum_i
            


      
        if school in schoolCategory["skyward"] or school in school_fye:
            
            for item in data_activitybs:
                Activity = item["Activity"]
                
                
                if school in schoolMonths['septemberSchool']:
                    if Activity in unique_act:

                        item["total_bal9"] -= item["activity_fye"] 
                    else:
                        item["total_bal9"] += item["activity_fye"]

                    item["total_bal10"] += item["total_bal9"]
                    item["total_bal11"] += item["total_bal10"]
                    item["total_bal12"] += item["total_bal11"]
                    item["total_bal1"] += item["total_bal12"]
                    item["total_bal2"] += item["total_bal1"]
                    item["total_bal3"] += item["total_bal2"]
                    item["total_bal4"] += item["total_bal3"]
                    item["total_bal5"] += item["total_bal4"]
                    item["total_bal6"] += item["total_bal5"]
                    item["total_bal7"] += item["total_bal6"]
                    item["total_bal8"] += item["total_bal7"]
                    item["last_month_bal"] = item[f'total_bal{last_month_number}']

                else:
                    if Activity in unique_act:
                        item["total_bal7"] -= item["activity_fye"] 
                    else:
                        item["total_bal7"] += item["activity_fye"] 

                    item["total_bal7"] += item["activity_fye"] 
                    item["total_bal8"] += item["total_bal7"]
                    item["total_bal9"] +=  item["total_bal8"]
                    item["total_bal10"] += item["total_bal9"]
                    item["total_bal11"] += item["total_bal10"]
                    item["total_bal12"] += item["total_bal11"]
                    item["total_bal1"] += item["total_bal12"]
                    item["total_bal2"] += item["total_bal1"]
                    item["total_bal3"] += item["total_bal2"]
                    item["total_bal4"] += item["total_bal3"]
                    item["total_bal5"] += item["total_bal4"]
                    item["total_bal6"] += item["total_bal5"]
                    item["last_month_bal"] = item[f'total_bal{last_month_number}']


                    



        
                


        for row in data_balancesheet:
            activity = row["Activity"]
            
            
            for i in range(1, 13):
                key = (activity, i)
                row[f"total_sum{i}"] = (activity_sum_dict.get(key, 0))

            if school in schoolCategory["skyward"] or school in school_fye:
                total_fye = sum(
                        entry["activity_fye"]
                        for entry in data_activitybs
                        if entry["Activity"] == activity 
                        and entry["activity_fye"] is not None                   
                    )
                
                row["total_fye"] =  total_fye



        # TOTAL REVENUE
        total_revenue = {acct_per: 0 for acct_per in acct_per_values}
        for item in data:
            fund = item["fund"]
            obj = item["obj"]

            for i, acct_per in enumerate(acct_per_values, start=1):
                total_real = sum(
                    entry[real_key]
                    for entry in data3
                    if entry["fund"] == fund
                    and entry["obj"] == obj
                    and entry["AcctPer"] == acct_per
                 
                )
                total_adjustment = sum(
                        entry[real_key]
                        for entry in adjustment
                        if entry["fund"] == fund
                        and entry["AcctPer"] == acct_per
                        and entry["obj"] == obj
                        and entry["School"] == school
                        and entry[real_key] is not None 
                        and not isinstance(entry[real_key], str)
                    )
                item[f"total_real{i}"] = total_real + total_adjustment
                # if i == last_month_number and (item[f"total_real{i}"] == 0):

                #         last_2months = current_month - relativedelta(months=1)
                #         last_2months = last_2months - relativedelta(days=1)
                #         last_month_name = last_2months.strftime("%B")
                #         formatted_last_month = last_2months.strftime('%B %d, %Y')
                
                total_revenue[acct_per] += (item[f"total_real{i}"])

        # if all(item[f"total_real{last_month_number}"] == 0 for item in data):
        #     last_2months = current_month - relativedelta(months=1)
        #     last_2months = last_2months - relativedelta(days=1)
        #     last_month_number = last_2months.month
        #     last_month_name = last_2months.strftime("%B")
        #     formatted_last_month = last_2months.strftime('%B %d, %Y')                    


        # total surplus / first total
        total_surplus = {acct_per: 0 for acct_per in acct_per_values}

        for item in data2:
            if item["category"] != "Depreciation and Amortization":
                func = item["func_func"]

                for i, acct_per in enumerate(acct_per_values, start=1):
                    total_func = sum(
                        entry[expend_key]
                        for entry in data3
                        if entry["func"] == func and entry["AcctPer"] == acct_per and entry["obj"] != '6449'
                    )
                    total_adjustment = sum(
                        entry[expend_key]
                        for entry in adjustment
                        if entry["func"] == func 
                        and entry["AcctPer"] == acct_per 
                        and entry["obj"] != '6449' 
                        and entry["School"] == school
                        and entry[expend_key] is not None 
                        and not isinstance(entry[expend_key], str)
                    )
                    item[f"total_func{i}"] = total_func + total_adjustment
                    total_surplus[acct_per] += item[f"total_func{i}"]

        # difference_func_values = {i: 0 for i in range(1, 13)}
        # monthly_totals_func = {i: 0 for i in range(1, 13)}
        # monthly_totals_func2 = {i: 0 for i in range(1, 13)}

        # ---- Depreciation and ammortization total
        total_DnA = {acct_per: 0 for acct_per in acct_per_values}

        for item in data2:
            if item["category"] == "Depreciation and Amortization":
                func = item["func_func"]
                obj = item["obj"]

                for i, acct_per in enumerate(acct_per_values, start=1):
                    total_func = sum(
                        entry[expend_key]
                        for entry in data3
                        if entry["func"] == func
                        and entry["AcctPer"] == acct_per
                        and entry["obj"] == obj
                    )
                    total_adjustment = sum(
                        entry[expend_key]
                        for entry in adjustment
                        if entry["func"] == func
                        and entry["AcctPer"] == acct_per
                        and entry["obj"] == obj
                        and entry["School"] == school
                        and entry[expend_key] is not None 
                        and not isinstance(entry[expend_key], str)
                        

                    )
                    item[f"total_func2_{i}"] = total_func + total_adjustment
                    total_DnA[acct_per] += item[f"total_func2_{i}"]

        total_SBD = {
            acct_per: abs(total_revenue[acct_per]) - total_surplus[acct_per]
            for acct_per in acct_per_values
        }
        total_netsurplus = {
            acct_per: total_SBD[acct_per] - total_DnA[acct_per] #dna_total_months in pl.. SBD same as pl
            for acct_per in acct_per_values
        }
    
        ytd_DnA = sum(total_DnA.values())
        ytd_netsurplus = sum(total_netsurplus.values())

        # for month, total in monthly_totals_func2.items():
        #     print(f'MonthFUNC2 {month}: {total}')

        # for key, value in difference_func_values.items():
        #     print(f'{key}: {value}')

        def format_with_parentheses(value):
            if value >= 1:
                return "${:,.0f}".format(round(value))
            elif value <= -1:
                return "$({:,.0f})".format(abs(round(value)))
            else:
                return ""

        def format_with_parentheses2(value):
            if value == 0:
                return ""
            formatted_value = "{:,.0f}".format(abs(round(value)))
            return "({})".format(formatted_value) if value > 0 else formatted_value

        def format_value_dollars(value):
            if value > 0:
                return "${:,.0f}".format(round(value))
            elif value < 0:
                return "$({:,.0f})".format(abs(round(value)))
            else:
                return ""
        def format_value(value):
            if value > 0:
                return "{:,.0f}".format(round(value))
            elif value < 0:
                return "({:,.0f})".format(abs(round(value)))
            else:
                return ""

        def format_negative(value):
            if value > 0:
                return "({:,.0f})".format(round(value))
            elif value < 0:
                return "{:,.0f}".format(abs(round(value)))
            else:
                return ""

        for row in data_balancesheet:
            if row["school"] == school:
                
                if school in schoolCategory["skyward"] or school in school_fye:
                    FYE_value = row["total_fye"]
                    
                # elif school =="goldenrule":
                #     FYE_value = float(row["FYE"])
                else:
                    FYE_value = (float(row["FYE"].replace("$","").replace(",", "").replace("(", "-").replace(")", ""))
                        if row["FYE"]
                        else 0
                    )
                total_sum9_value = float(row["total_sum9"])
                total_sum10_value = float(row["total_sum10"])
                total_sum11_value = float(row["total_sum11"])
                total_sum12_value = float(row["total_sum12"])
                total_sum1_value = float(row["total_sum1"])
                total_sum2_value = float(row["total_sum2"])
                total_sum3_value = float(row["total_sum3"])
                total_sum4_value = float(row["total_sum4"])
                total_sum5_value = float(row["total_sum5"])
                total_sum6_value = float(row["total_sum6"])
                total_sum7_value = float(row["total_sum7"])
                total_sum8_value = float(row["total_sum8"])

                total_sums = [
                                float(row[f"total_sum{i}"]) for i in range(1, 13)
                            ]
                if school in schoolMonths['septemberSchool']:
                
                    # Calculate the differences and store them in the row dictionary
                    row["difference_9"] = (FYE_value + total_sum9_value)
                    row["difference_10"] =(row["difference_9"] + total_sum10_value)
                    row["difference_11"] =(row["difference_10"] + total_sum11_value)
                    row["difference_12"] =(row["difference_11"]  + total_sum12_value )
                    row["difference_1"] = (row["difference_12"] + total_sum1_value )
                    row["difference_2"] = (row["difference_1"] + total_sum2_value )
                    row["difference_3"] = (row["difference_2"] + total_sum3_value )
                    row["difference_4"] = (row["difference_3"] + total_sum4_value )
                    row["difference_5"] = (row["difference_4"] + total_sum5_value )
                    row["difference_6"] = (row["difference_5"] + total_sum6_value )
                    row["difference_7"] = (row["difference_6"] + total_sum7_value )
                    row["difference_8"] = (row["difference_7"] + total_sum8_value )
    
                
                    row["last_month_difference"] = row[f"difference_{last_month_number}"] 
              

                    if month_exception != "":
                        
                        row["fytd"] = sum(
                            total_sum
                            for i, total_sum in enumerate(total_sums, start=1)
                            if i != month_exception
                        )
                    else:
                        row["fytd"] =sum(total_sums)

                    row["debt_9"]  = (FYE_value - total_sum9_value)
                    row["debt_10"] = (row["debt_9"] - total_sum10_value)
                    row["debt_11"] = (row["debt_10"] - total_sum11_value)
                    row["debt_12"] = (row["debt_11"] - total_sum12_value)
                    row["debt_1"] = (row["debt_12"] - total_sum1_value)
                    row["debt_2"] = (row["debt_1"] - total_sum2_value)
                    row["debt_3"] = (row["debt_2"] - total_sum3_value)
                    row["debt_4"] = (row["debt_3"]- total_sum4_value)
                    row["debt_5"] = (row["debt_4"]  - total_sum5_value )
                    row["debt_6"] = (row["debt_5"]- total_sum6_value)
                    row["debt_7"] = (row["debt_6"] - total_sum7_value)
                    row["debt_8"] = (row["debt_7"] - total_sum8_value)
                    row["last_month_debt"] = row[f"debt_{last_month_number}"] 



                    if month_exception != "":
                        
                        row["debt_fytd"] = -sum(
                            total_sum
                            for i, total_sum in enumerate(total_sums, start=1)
                            if i != month_exception
                        )
                    else:
                        row["debt_fytd"] =-sum(total_sums)
                    

                    row["net_assets9"] = (FYE_value + total_netsurplus["09"])
                    row["net_assets10"] = (row["net_assets9"] + total_netsurplus["10"])
                    row["net_assets11"] = (row["net_assets10"]+ total_netsurplus["11"])
                    row["net_assets12"] = (row["net_assets11"]+ total_netsurplus["12"])
                    row["net_assets1"] = (row["net_assets12"] + total_netsurplus["01"])
                    row["net_assets2"] = (row["net_assets1"] + total_netsurplus["02"])
                    row["net_assets3"] = (row["net_assets2"]+ total_netsurplus["03"])
                    row["net_assets4"] = (row["net_assets3"] + total_netsurplus["04"])
                    row["net_assets5"] = (row["net_assets4"] + total_netsurplus["05"])
                    row["net_assets6"] = (row["net_assets5"]  + total_netsurplus["06"])
                    row["net_assets7"] = (row["net_assets6"] + total_netsurplus["07"])
                    row["net_assets8"] = (row["net_assets7"] + total_netsurplus["08"])
                    row["last_month_net_assets"] = row[f"net_assets{last_month_number}"]
                    
                else:
                                    # Calculate the differences and store them in the row dictionary
                    row["difference_7"] = (FYE_value + total_sum7_value )
            
                    row["difference_8"] = (row["difference_7"] + total_sum8_value )
                    row["difference_9"] = (row["difference_8"]  + total_sum9_value)
                    row["difference_10"] =(row["difference_9"] + total_sum10_value)
                    row["difference_11"] =(row["difference_10"] + total_sum11_value)
                    row["difference_12"] =(row["difference_11"]  + total_sum12_value )
                    row["difference_1"] = (row["difference_12"] + total_sum1_value )
                    row["difference_2"] = (row["difference_1"] + total_sum2_value )
                    row["difference_3"] = (row["difference_2"] + total_sum3_value )
                    row["difference_4"] = (row["difference_3"] + total_sum4_value )
                    row["difference_5"] = (row["difference_4"] + total_sum5_value )
                    row["difference_6"] = (row["difference_5"] + total_sum6_value )
                    
                    row["last_month_difference"] = row[f"difference_{last_month_number}"] 
                    

                    if month_exception != "":
                        
                        row["fytd"] = sum(
                            total_sum
                            for i, total_sum in enumerate(total_sums, start=1)
                            if i != month_exception
                        )
                    else:
                        row["fytd"] =sum(total_sums)

                    row["debt_7"] = (FYE_value - total_sum7_value)
                    row["debt_8"] = (row["debt_7"] - total_sum8_value)
                    row["debt_9"]  = (row["debt_8"] - total_sum9_value)
                    row["debt_10"] = (row["debt_9"] - total_sum10_value)
                    row["debt_11"] = (row["debt_10"] - total_sum11_value)
                    row["debt_12"] = (row["debt_11"] - total_sum12_value)
                    row["debt_1"] = (row["debt_12"] - total_sum1_value)
                    row["debt_2"] = (row["debt_1"] - total_sum2_value)
                    row["debt_3"] = (row["debt_2"] - total_sum3_value)
                    row["debt_4"] = (row["debt_3"]- total_sum4_value)
                    row["debt_5"] = (row["debt_4"]  - total_sum5_value )
                    row["debt_6"] = (row["debt_5"]- total_sum6_value)
                    row["last_month_debt"] = row[f"debt_{last_month_number}"] 
    
                    
                    if month_exception != "":
                        
                        row["debt_fytd"] = -sum(
                            total_sum
                            for i, total_sum in enumerate(total_sums, start=1)
                            if i != month_exception
                        )
                    else:
                        row["debt_fytd"] =-sum(total_sums)

                    row["net_assets7"] = (FYE_value + total_netsurplus["07"])
                    row["net_assets8"] = (row["net_assets7"] + total_netsurplus["08"])
                    row["net_assets9"] = (row["net_assets8"]  + total_netsurplus["09"])
                    row["net_assets10"] = (row["net_assets9"] + total_netsurplus["10"])
                    row["net_assets11"] = (row["net_assets10"]+ total_netsurplus["11"])
                    row["net_assets12"] = (row["net_assets11"]+ total_netsurplus["12"])
                    row["net_assets1"] = (row["net_assets12"] + total_netsurplus["01"])
                    row["net_assets2"] = (row["net_assets1"] + total_netsurplus["02"])
                    row["net_assets3"] = (row["net_assets2"]+ total_netsurplus["03"])
                    row["net_assets4"] = (row["net_assets3"] + total_netsurplus["04"])
                    row["net_assets5"] = (row["net_assets4"] + total_netsurplus["05"])
                    row["net_assets6"] = (row["net_assets5"]  + total_netsurplus["06"])
                    row["last_month_net_assets"] = row[f"net_assets{last_month_number}"]

        total_current_assets = {acct_per: 0 for acct_per in acct_per_values}
        total_current_assets_fye = 0
        total_current_assets_fytd = 0 
        last_month_current_assets = 0 

        total_capital_assets = {acct_per: 0 for acct_per in acct_per_values}
        total_capital_assets_fye = 0
        total_capital_assets_fytd = 0 
        last_month_total_capital_assets = 0

        total_current_liabilities = {acct_per: 0 for acct_per in acct_per_values}
        total_current_liabilities_fye = 0
        total_current_liabilities_fytd = 0
        last_month_total_current_liabilities = 0

        total_liabilities = {acct_per: 0 for acct_per in acct_per_values}
        total_liabilities_fye = 0
        total_liabilities_fytd = 0
        last_month_total_liabilities = 0

        total_assets = {acct_per: 0 for acct_per in acct_per_values}
        total_assets_fye = 0
        total_assets_fye_fytd = 0
        last_month_total_assets = 0 
        

        total_LNA = {acct_per: 0 for acct_per in acct_per_values} # LIABILITES AND NET ASSETS 
        total_LNA_fye = 0
        total_LNA_fytd = 0
        total_net_assets_fytd = 0
        last_month_total_LNA = 0

        total_net_assets_fytd = totals["bs_ytd_netsurplus"]    #assign the value coming from profitloss totals
        
    

        for row in data_balancesheet:
            if row["school"] == school:
                subcategory =  row["Subcategory"]

                # if school == "goldenrule":
                #     fye = float(row["FYE"])
                # else:
                fye =  float(row["FYE"].replace("$","").replace(",", "").replace("(", "-").replace(")", "")) if row["FYE"] else 0
                if school in schoolCategory["skyward"] or school in school_fye:
                    fye = row["total_fye"]

                if subcategory == 'Current Assets':
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_current_assets[acct_per] += row[f"difference_{i}"]
                        if i == last_month_number:
                            last_month_current_assets += row[f"difference_{i}"]                      
                    total_current_assets_fytd += row["fytd"]

                    total_current_assets_fye +=  fye
                if subcategory == 'Capital Assets, Net':
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_capital_assets[acct_per] += row[f"difference_{i}"]
                        if i == last_month_number:
                            last_month_total_capital_assets += row[f"difference_{i}"]
                    
                            

                    total_capital_assets_fytd += row["fytd"]
                    total_capital_assets_fye +=  fye

                    
                if subcategory == 'Current Liabilities':
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_current_liabilities[acct_per] += row[f"debt_{i}"]
                        if i == last_month_number:
                            last_month_total_current_liabilities += row[f"debt_{i}"]
                    total_current_liabilities_fytd += row["debt_fytd"]
                    total_current_liabilities_fye +=  fye

        total_liabilities_fytd_2 = 0
        for row in data_balancesheet:
            if row["school"] == school:
                subcategory =  row["Subcategory"]
                
                # if school == "goldenrule":
                #     fye = float(row["FYE"])
                # else:
                fye =  float(row["FYE"].replace("$","").replace(",", "").replace("(", "-").replace(")", "")) if row["FYE"] else 0
                if school in schoolCategory["skyward"] or school in school_fye:
                    fye = row["total_fye"]
                if subcategory == 'Long Term Debt':
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_liabilities[acct_per] += row[f"debt_{i}"] + total_current_liabilities[acct_per]
                        if i == last_month_number:
                            last_month_total_liabilities += row[f"debt_{i}"] + total_current_liabilities[acct_per]
                    total_liabilities_fytd_2 += row["debt_fytd"]
                    total_liabilities_fye +=   total_current_liabilities_fye + fye
        total_liabilities_fytd = total_liabilities_fytd_2 + total_current_liabilities_fytd

        for row in data_balancesheet:
            if row["school"] == school:
                
                # if school == "goldenrule":
                #     fye = float(row["FYE"])
                # else:
                fye =  float(row["FYE"].replace("$","").replace(",", "").replace("(", "-").replace(")", "")) if row["FYE"] else 0
                if school in schoolCategory["skyward"] or school in school_fye:
                    fye = row["total_fye"]
                if  row["Category"] == "Net Assets":
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_LNA[acct_per] += row[f"net_assets{i}"] + total_liabilities[acct_per]
                        if i == last_month_number:
                            last_month_total_LNA += row[f"net_assets{i}"] + total_liabilities[acct_per]

                    total_LNA_fye += total_liabilities_fye + fye
        
        total_assets = {
            acct_per: total_current_assets[acct_per] + total_capital_assets[acct_per]
            for acct_per in acct_per_values

        }

        last_month_number_str = f"{last_month_number:02}"  
        last_month_total_assets  = total_assets[last_month_number_str]
        
        

        total_assets_fye = total_current_assets_fye + total_capital_assets_fye

        total_assets_fye_fytd = total_current_assets_fytd + total_capital_assets_fytd
        total_LNA_fytd = total_net_assets_fytd + total_liabilities_fytd
        total_net_assets_fytd = format_value(total_net_assets_fytd)
        total_current_assets_fye = format_value(total_current_assets_fye)
        total_capital_assets_fye = format_value(total_capital_assets_fye)
        total_current_liabilities_fye = format_value(total_current_liabilities_fye)
        total_liabilities_fye = format_value(total_liabilities_fye)
        total_assets_fye = format_value_dollars(total_assets_fye)
        total_LNA_fye = format_value_dollars(total_LNA_fye)

        total_current_assets_fytd = format_value(total_current_assets_fytd)
        total_capital_assets_fytd = format_value(total_capital_assets_fytd)
        total_current_liabilities_fytd = format_value(total_current_liabilities_fytd)
        total_liabilities_fytd = format_value(total_liabilities_fytd)
        total_assets_fye_fytd = format_value_dollars(total_assets_fye_fytd)
        total_LNA_fytd = format_value_dollars(total_LNA_fytd)

        total_current_assets = {acct_per: format_value(value) for acct_per, value in total_current_assets.items() if value != 0}
        total_capital_assets = {acct_per: format_value(value) for acct_per, value in total_capital_assets.items() if value != 0}
        total_current_liabilities = {acct_per: format_value(value) for acct_per, value in total_current_liabilities.items() if value != 0}
        total_liabilities = {acct_per: format_value(value) for acct_per, value in total_liabilities.items() if value != 0}
        total_assets = {acct_per: format_value_dollars(value) for acct_per, value in total_assets.items() if value != 0}
        total_LNA = {acct_per: format_value_dollars(value) for acct_per, value in total_LNA.items() if value != 0}

        last_month_current_assets = format_value(last_month_current_assets)
        last_month_total_capital_assets = format_value(last_month_total_capital_assets)
        last_month_total_assets = format_value_dollars(last_month_total_assets)
        last_month_total_current_liabilities = format_value(last_month_total_current_liabilities)
        last_month_total_liabilities = format_value(last_month_total_liabilities)
        last_month_total_LNA = format_value_dollars(last_month_total_LNA)
        
    
        for row in data_balancesheet:
            if row["school"] == school:

                if school in schoolCategory["skyward"] or school in school_fye:
                    if row["Activity"] == 'Cash' or row["Activity"] == 'AP':

                        row["total_fye"] = format_value_dollars(row["total_fye"]) 
                    else:
                        row["total_fye"] = format_value(row["total_fye"]) 

                # if school == 'goldenrule':
                #     if row["Activity"] == 'Cash' or row["Activity"] == 'AP':

                #         row["FYE"] = format_value_dollars(row["FYE"])
                #     else:
                #         row["FYE"] = format_value(row["FYE"])

                if row["Activity"] == 'Cash':
                    
                    row["difference_9"] = format_value_dollars(row["difference_9"]) 
                    row["difference_10"]= format_value_dollars(row["difference_10"])
                    row["difference_11"]= format_value_dollars(row["difference_11"])
                    row["difference_12"]= format_value_dollars(row["difference_12"])
                    row["difference_1"] = format_value_dollars(row["difference_1"] )
                    row["difference_2"] = format_value_dollars(row["difference_2"] )
                    row["difference_3"] = format_value_dollars(row["difference_3"] )
                    row["difference_4"] = format_value_dollars(row["difference_4"] )
                    row["difference_5"] = format_value_dollars(row["difference_5"] )
                    row["difference_6"] = format_value_dollars(row["difference_6"] )
                    row["difference_7"] = format_value_dollars(row["difference_7"] )
                    row["difference_8"] = format_value_dollars(row["difference_8"] )
                    
                    row["last_month_difference"] = format_value_dollars(row["last_month_difference"] )
                    row["fytd"] = format_value_dollars(row["fytd"])
                else:
                    
                    row["difference_9"] = format_value(row["difference_9"]) 
                    row["difference_10"]= format_value(row["difference_10"])
                    row["difference_11"]= format_value(row["difference_11"])
                    row["difference_12"]= format_value(row["difference_12"])
                    row["difference_1"] = format_value(row["difference_1"] )
                    row["difference_2"] = format_value(row["difference_2"] )
                    row["difference_3"] = format_value(row["difference_3"] )
                    row["difference_4"] = format_value(row["difference_4"] )
                    row["difference_5"] = format_value(row["difference_5"] )
                    row["difference_6"] = format_value(row["difference_6"] )
                    row["difference_7"] = format_value(row["difference_7"] )
                    row["difference_8"] = format_value(row["difference_8"] )
                    row["last_month_difference"] = format_value(row["last_month_difference"] )
                    row["fytd"] = format_value(row["fytd"])
                
                if row['Activity'] == 'AP':
                    
                    row["debt_9"] =  format_value_dollars(row["debt_9"] )
                    row["debt_10"]=  format_value_dollars(row["debt_10"])
                    row["debt_11"]=  format_value_dollars(row["debt_11"])
                    row["debt_12"]=  format_value_dollars(row["debt_12"])
                    row["debt_1"] =  format_value_dollars(row["debt_1"] )
                    row["debt_2"] =  format_value_dollars(row["debt_2"] )
                    row["debt_3"] =  format_value_dollars(row["debt_3"] )
                    row["debt_4"] =  format_value_dollars(row["debt_4"] )
                    row["debt_5"] =  format_value_dollars(row["debt_5"] )
                    row["debt_6"] =  format_value_dollars(row["debt_6"] )
                    row["debt_7"] =  format_value_dollars(row["debt_7"] )
                    row["debt_8"] =  format_value_dollars(row["debt_8"] )
                    row["debt_fytd"]=format_value_dollars(row["debt_fytd"])
                    row["last_month_debt"] = format_value_dollars(row["last_month_debt"] )

                else:
                    
                    row["debt_9"] =  format_value(row["debt_9"] )
                    row["debt_10"]=  format_value(row["debt_10"])
                    row["debt_11"]=  format_value(row["debt_11"])
                    row["debt_12"]=  format_value(row["debt_12"])
                    row["debt_1"] =  format_value(row["debt_1"] )
                    row["debt_2"] =  format_value(row["debt_2"] )
                    row["debt_3"] =  format_value(row["debt_3"] )
                    row["debt_4"] =  format_value(row["debt_4"] )
                    row["debt_5"] =  format_value(row["debt_5"] )
                    row["debt_6"] =  format_value(row["debt_6"] )
                    row["debt_7"] =  format_value(row["debt_7"] )
                    row["debt_8"] =  format_value(row["debt_8"] )
                    row["last_month_debt"] = format_value(row["last_month_debt"] )
                    row["debt_fytd"]=format_value(row["debt_fytd"])
        

                row["net_assets9"]  = format_value(row["net_assets9"])
                row["net_assets10"] = format_value(row["net_assets10"])
                row["net_assets11"] = format_value(row["net_assets11"])
                row["net_assets12"] = format_value(row["net_assets12"])
                row["net_assets1"]  = format_value(row["net_assets1"])
                row["net_assets2"]  = format_value(row["net_assets2"])
                row["net_assets3"]  = format_value(row["net_assets3"])
                row["net_assets4"]  = format_value(row["net_assets4"])
                row["net_assets5"]  = format_value(row["net_assets5"])
                row["net_assets6"]  = format_value(row["net_assets6"])
                row["net_assets7"]  = format_value(row["net_assets7"])
                row["net_assets8"]  = format_value(row["net_assets8"])
                row["last_month_net_assets"] = format_value(row["last_month_net_assets"])

        keys_to_check = [
            "total_bal1",
            "total_bal2",
            "total_bal3",
            "total_bal4",
            "total_bal5",
            "total_bal6",
            "total_bal7",
            "total_bal8",
            "total_bal9",
            "total_bal10",
            "total_bal11",
            "total_bal12",
            "fytd"
        ]

        threshold = 0.50
        if school in schoolCategory["skyward"] or school in school_fye:
            for row in data_activitybs:
                Activity = row["Activity"]

                if Activity in unique_act:
                    row["last_month_bal"] = format_negative(row["last_month_bal"])
                    for key in keys_to_check:
                        value = float(row[key])
                        
                        if value == 0 or value == 0.00 or value == 0.0  :
                            row[key] = ""
                        elif value >=0:
                            
                            row[key] = "({:,.0f})".format(abs(float(row[key])))
                        elif value < 0:                            
                            row[key] = "{:,.0f}".format(abs(float(row[key])))
                        elif value != "":
                            row[key] = "{:,.0f}".format(float(row[key]))
                else:
                    
                    row["last_month_bal"] = format_value(row["last_month_bal"])
                    for key in keys_to_check:
                        value = float(row[key])
                        if value == 0:
                            row[key] = ""
                        elif value < 0:
                            
                            row[key] = "({:,.0f})".format(abs(float(row[key])))
                        elif value != "":
                            row[key] = "{:,.0f}".format(float(row[key]))

        else:
            for row in data_activitybs:

                for key in keys_to_check:
                    value = float(row[key])
                    if value == 0:
                        row[key] = ""
                    elif value > 0:
                        
                        row[key] = "{:,.0f}".format(abs(float(row[key])))
                    elif value != "":
                        row[key] = "({:,.0f})".format(float(row[key]))


        if school in schoolCategory["skyward"] or school in school_fye:
            for row in data_activitybs:

                row["activity_fye"] = format_value(row["activity_fye"])

        # for row in data_balancesheet:
        #     subcategory = row["Subcategory"]
        #     fye = float(row["FYE"])

        #     row["total_fye"][subcategory] += fye
        #     row["total_fye"] = total_fye[subcategory]

        # for row in data_balancesheet:
        #     row['diffunc9']

        # keys_to_check_func = ['total_func1', 'total_func2', 'total_func3', 'total_func4', 'total_func5','total_func6','total_func7','total_func8','total_func9','total_func10','total_func11','total_func12']
        # keys_to_check_func_2 = ['total_func2_1', 'total_func2_2', 'total_func2_3', 'total_func2_4', 'total_func2_5','total_func2_6','total_func2_7','total_func2_8','total_func2_9','total_func2_10','total_func2_11','total_func2_12']

        # for row in data2:
        #     for key in keys_to_check_func:
        #         if row[key] > 0:
        #             row[key] = row[key]
        #         else:
        #             row[key] = ''
        # for row in data2:
        #     for key in keys_to_check_func:
        #         if row[key] != "":
        #             row[key] = "{:,.0f}".format(row[key])

        # for row in data2:
        #     for key in keys_to_check_func_2:
        #         if row[key] > 0:
        #             row[key] = row[key]
        #         else:
        #             row[key] = ''
        # for row in data2:
        #     for key in keys_to_check_func_2:
        #         if row[key] != "":
        #             row[key] = "{:,.0f}".format(row[key])

        formatted_total_netsurplus = {
            acct_per: "${:,}".format(abs(int(value)))
            if value > 0
            else "(${:,})".format(abs(int(value)))
            if value < 0
            else ""
            for acct_per, value in total_netsurplus.items()
            if value != 0
        }
        formatted_total_DnA = {
            acct_per: "{:,}".format(abs(int(value)))
            if value >= 0
            else "({:,})".format(abs(int(value)))
            if value < 0
            else ""
            for acct_per, value in total_DnA.items()
            if value != 0
        }

        formated_ytdnetsurplus = format_with_parentheses(ytd_netsurplus)

        bs_activity_list = list(
            set(row["Activity"] for row in data_balancesheet if "Activity" in row)
        )
        bs_activity_list_sorted = sorted(bs_activity_list)
        gl_obj = list(set(row["obj"] for row in data3 if "obj" in row))
        gl_obj_sorted = sorted(gl_obj)

        # func_choice = list(set(row['func'] for row in data3 if 'func' in row))
        # func_choice_sorted = sorted(func_choice)

        data_activitybs = sorted(data_activitybs, key=lambda x: x['obj'])
        #difference_key = "difference_" + str(last_month_number)
        context = {
            "data_balancesheet": data_balancesheet,
            "data_activitybs": data_activitybs,
            # "data3": data3,
            "bs_activity_list": bs_activity_list_sorted,
            "gl_obj": gl_obj_sorted,
            # "button_rendered": button_rendered,
        
            "last_month": last_month,
            "last_month_number": last_month_number,
            "last_month_name": last_month_name,
            "FY_year_1":FY_year_1,
            "FY_year_2":FY_year_2,
            "totals_bs":{
                "total_current_assets":total_current_assets,
                "total_current_assets_fye":total_current_assets_fye,
                "last_month_current_assets":last_month_current_assets,
                "total_capital_assets":total_capital_assets,
                "total_capital_assets_fye":total_capital_assets_fye,
                "last_month_total_capital_assets":last_month_total_capital_assets,
                "total_current_liabilities":total_current_liabilities,
                "total_current_liabilities_fye":total_current_liabilities_fye,
                "last_month_total_current_liabilities":last_month_total_current_liabilities,
                "total_liabilities":total_liabilities,
                "total_liabilities_fye":total_liabilities_fye,
                "last_month_total_liabilities":last_month_total_liabilities,
                "total_assets": total_assets,
                "total_assets_fye":total_assets_fye,
                "last_month_total_assets":last_month_total_assets,
                "total_LNA_fye":total_LNA_fye,
                "total_LNA":total_LNA,
                "total_current_assets_fytd":total_current_assets_fytd,
                "total_capital_assets_fytd":total_capital_assets_fytd,
                "total_current_liabilities_fytd":total_current_liabilities_fytd,
                "total_liabilities_fytd":total_liabilities_fytd,
                "total_assets_fye_fytd":total_assets_fye_fytd,
                "total_net_assets_fytd":total_net_assets_fytd,
                "total_LNA_fytd":total_LNA_fytd,
                "last_month_total_LNA":last_month_total_LNA
            }

            
            #"difference_key": difference_key,
            # "format_ytd_budget": formatted_ytd_budget,
            # "ytd_budget": ytd_budget,
        }

        if school in schoolCategory["ascender"]:
            context["total_DnA"] = (formatted_total_DnA,)
            context["total_netsurplus"] = formatted_total_netsurplus
            context["total_SBD"] = total_SBD
            context["ytd_netsurplus"] = formated_ytdnetsurplus
    # return context
    # dict_keys = ["data", "data2", "data3", "data_expensebyobject", "data_activities"]

        if FY_year_1 == present_year:
            relative_path = os.path.join("balance-sheet", school)
        else:
            relative_path = os.path.join(str(FY_year_1), "balance-sheet", school)

        # json_path = JSON_DIR.path(relative_path)  
        json_path = os.path.join(JSON_DIR,relative_path)

        shutil.rmtree(json_path, ignore_errors=True)
        if not os.path.exists(json_path):
            os.makedirs(json_path)

        for key, val in context.items():
            file = os.path.join(json_path, f"{key}.json")
            with open(file, "w") as f:
                json.dump(val, f)

def cashflow(school,year):

    def format_value_dollars(value):
            if value > 0:
                return "${:,.0f}".format(round(value))
            elif value < 0:
                return "$({:,.0f})".format(abs(round(value)))
            else:
                return ""
    def format_value(value):
        if value > 0:
            return "{:,.0f}".format(round(value))
        elif value < 0:
            return "({:,.0f})".format(abs(round(value)))
        else:
            return ""
    def format_value_dollars_negative(value):
        if value > 0:
            return "$({:,.0f})".format(abs(round(value)))
            
        elif value < 0:
            
            return "${:,.0f}".format(abs(round(value)))
        else:
            return ""
    def format_value_negative(value):
        if value > 0:
            return "({:,.0f})".format(abs(round(value)))
            
        elif value < 0:
            
            return "{:,.0f}".format(abs(round(value)))
        else:
            return ""

    
    present_date = datetime.today().date()   
    present_year = present_date.year
    today_date = datetime.now()
    
    today_month = today_date.month

    if year:
        year = int(year)
        if year == present_year:
            
            print("year",year)

            if school in schoolMonths["septemberSchool"]:
                if today_month <= 8:
                    
                    start_year = year - 1
                    present_year = present_year - 1
                    FY_year_current = year - 1
                else: 
                    start_year = year 
                    FY_year_current = year
            else:
                if today_month <= 6:
                    start_year = year - 1
                    present_year = present_year - 1
                    FY_year_current = year - 1
                else: 
                    start_year = year 
                    FY_year_current = year
        else:
            start_year = year 
            FY_year_current = year
     
        if school in schoolMonths["julySchool"]:
            current_date = datetime(start_year, 7, 1).date()
            
        else:
            current_date = datetime(start_year, 9, 1).date() 
        current_year = current_date.year
    else:
        start_year = 2021
        current_date = datetime.today().date()   
        current_year = current_date.year
        FY_year_current = current_year


    while start_year <= FY_year_current:
        FY_year_1 = start_year
        FY_year_2 = start_year + 1
        start_year = FY_year_2
        cnxn = connect()
        cursor = cnxn.cursor()


        if FY_year_1 == present_year:
            relative_path = os.path.join("profit-loss", school)
        else:
            relative_path = os.path.join(str(FY_year_1), "profit-loss", school)
        # json_path = JSON_DIR.path(relative_path)
        json_path = os.path.join(JSON_DIR, relative_path)

        with open(os.path.join(json_path, "data.json"), "r") as f:
            data = json.load(f)

        with open(os.path.join(json_path, "data2.json"), "r") as f:
            data2 = json.load(f)

        with open(os.path.join(json_path, "data3.json"), "r") as f:
            data3 = json.load(f)

        with open(os.path.join(json_path, "data_expensebyobject.json"), "r") as f:
            data_expensebyobject = json.load(f)

        with open(os.path.join(json_path, "data_activities.json"), "r") as f:
            data_activities = json.load(f)

        with open(os.path.join(json_path, "totals.json"), "r") as f:
            totals = json.load(f)


        with open(os.path.join(json_path, "months.json"), "r") as f:
            months = json.load(f)

        cursor.execute(f"SELECT * FROM [dbo].{db[school]['cashflow']};")
        rows = cursor.fetchall()

        data_cashflow = []

        for row in rows:
            row_dict = {
                "Category": row[0],
                "Activity": row[1],
                "Description": row[2],
                "obj": str(row[3]),
            }
          

            data_cashflow.append(row_dict)


        if FY_year_1 == present_year:
            relative_path = os.path.join( "balance-sheet", school)
        else:
            relative_path = os.path.join(str(FY_year_1), "balance-sheet", school)
        
        # json_path = JSON_DIR.path(relative_path)
        json_path = os.path.join(JSON_DIR, relative_path)
        with open(os.path.join(json_path, "data_activitybs.json"), "r") as f:
            data_activitybs = json.load(f)

        with open(os.path.join(json_path, "data_balancesheet.json"), "r") as f:
            data_balancesheet = json.load(f)




        acct_per_values = [
            "01",
            "02",
            "03",
            "04",
            "05",
            "06",
            "07",
            "08",
            "09",
            "10",
            "11",
            "12",
        ]

        last_month = months["last_month"]
        last_month_number = months["last_month_number"]
        last_month_name = months["last_month_name"]
   
        db_last_month = months["db_last_month"]
        month_exception = months["month_exception"]
        month_exception_str = months["month_exception_str"]

        activity_key = "Bal"
        if school in schoolCategory["skyward"]:
            activity_key = "Amount"

        # ---------- FOR EXPENSE TOTAL -------


        chars_to_remove = r'[(),]'
        for i, acct_per in enumerate(acct_per_values, start=1):
            key = f"total_bal{i}"
            for entry in data_activitybs:
                sign = 1
                if '(' in entry[key]:
                    sign = -1
                try:
                    entry[key] = float(re.sub(chars_to_remove, '', entry[key])) * sign
                except (TypeError, ValueError):
                    if entry[key] == '':
                        entry[key] = 0
                    else:
                        print(entry[key])


        total_activity = {acct_per: 0 for acct_per in acct_per_values}

        total_investing = {acct_per: 0 for acct_per in acct_per_values}
        total_operating = {acct_per: 0 for acct_per in acct_per_values}
        total_operating_ytd = 0 
        total_investing_ytd = 0

        cfchecker= {acct_per: 0 for acct_per in acct_per_values}


        dna_months = totals["dna_total_months"]
        total_netsurplus = totals["total_netsurplus_months"]


        school_fye = ['aca','advantage','cumberland','pro-vision','manara','stmary','sa']


        for item in data_cashflow:
            activity = item["Activity"]

            item["fytd_1"] = 0

            for i, acct_per in enumerate(acct_per_values, start=1):
                key = f"total_bal{i}"
                item[f"total_operating{i}"] = sum(
                    entry[key]
                    for entry in data_activitybs
                    if entry["Activity"] == activity
                )

                total_activity[acct_per] += item[f"total_operating{i}"]
                total_operating[acct_per] += item[f"total_operating{i}"]

               


                
                if i != month_exception:
                    item["fytd_1"] += item[f"total_operating{i}"]
        
        
   

               

        for item in data_cashflow:
            obj = item["obj"]
            item["fytd_2"] = 0

            for i, acct_per in enumerate(acct_per_values, start=1):
                item[f"total_investing{i}"] = sum(
                    entry[activity_key]
                    for entry in data3
                    if entry["obj"] == obj and entry["AcctPer"] == acct_per
                )

                total_activity[acct_per] += item[f"total_investing{i}"]
                total_investing[acct_per] += item[f"total_investing{i}"]
                
                
              
                if i != month_exception:
                    item["fytd_2"] += item[f"total_investing{i}"] 
        
        

        def stringParser(value):
            print(value)
            if value == "" or value == 0:
                return 0

           
            if "(" in value:
                
                formatted = "".join(value.strip().replace("$", "").replace("(", "-").replace(")", "").split(","))
                
                if "." in formatted:
                  
                    return float(formatted)
             
                return int(formatted) 
            
            formatted = "".join(value.strip().replace("$", "").split(","))

            if "." in formatted:
                return float(formatted)
            return int(formatted)
            
        



        # PART OF TOTAL OPERATING

        #add for CFS 
        total_activity = {
            acct_per: total_activity[acct_per] + stringParser(dna_months.get(acct_per, 0))
            for acct_per in acct_per_values 
        }


        total_activity = {
            acct_per: total_activity[acct_per] - stringParser(total_netsurplus.get(acct_per, 0))
            for acct_per in acct_per_values 
        }

        #add for total operating total
        total_operating= {
            acct_per: total_operating[acct_per] + stringParser(dna_months.get(acct_per, 0))
            for acct_per in acct_per_values 
        }

        total_operating = {
            acct_per: total_operating[acct_per] - stringParser(total_netsurplus.get(acct_per, 0))
            for acct_per in acct_per_values 
        }
        #END OF TOTAL OPERATING

        total_activity_ytd = sum(total_activity.values())
        total_operating_ytd = sum(total_operating.values())
        total_investing_ytd = sum(total_investing.values())



       

        

        for row in data_balancesheet:
            if row["school"] == school and row["Category"] == "Assets" and row["Activity"] == "Cash":
                if school == 'goldenrule':
                    begbal = stringParser(row["FYE"])
                else:
                    begbal = stringParser(row["FYE"])
                    
                if school in schoolCategory["skyward"] or school in school_fye:
                    begbal = stringParser(row["total_fye"])
                
         
                cfchecker["09"] = begbal- stringParser(row["difference_9"]) + total_activity["09"]
                cfchecker["10"] = stringParser(row["difference_9"]) -  stringParser(row["difference_10"]) + total_activity["10"]
                cfchecker["11"] = stringParser(row["difference_10"]) - stringParser(row["difference_11"]) + total_activity["11"]
                cfchecker["12"] = stringParser(row["difference_11"]) - stringParser(row["difference_12"]) + total_activity["12"]
                cfchecker["01"] = stringParser(row["difference_12"]) - stringParser(row["difference_1"]) + total_activity["01"]
                cfchecker["02"] = stringParser(row["difference_1"]) -  stringParser(row["difference_2"]) + total_activity["02"]
                cfchecker["03"] = stringParser(row["difference_2"]) -  stringParser(row["difference_3"]) + total_activity["03"]
                cfchecker["04"] = stringParser(row["difference_3"]) -  stringParser(row["difference_4"]) + total_activity["04"]
                cfchecker["05"] = stringParser(row["difference_4"]) -  stringParser(row["difference_5"]) + total_activity["05"]
                cfchecker["06"] = stringParser(row["difference_5"]) -  stringParser(row["difference_6"]) + total_activity["06"]
                cfchecker["07"] = stringParser(row["difference_6"]) -  stringParser(row["difference_7"]) + total_activity["07"]
                cfchecker["08"] = stringParser(row["difference_7"]) -  stringParser(row["difference_8"]) + total_activity["08"]

        


               
        cfchecker = {acct_per: format_value_negative(value) for acct_per, value in cfchecker.items() if value != 0}
        total_investing = {acct_per: format_value_negative(value) for acct_per, value in total_investing.items() if value != 0}
        total_operating = {acct_per: format_value_negative(value) for acct_per, value in total_operating.items() if value != 0}
        total_activity = {acct_per: format_value_negative(value) for acct_per, value in total_activity.items() if value != 0}

        total_operating_ytd = format_value(total_operating_ytd)
        total_investing_ytd = format_value(total_investing_ytd)
        total_activity_ytd = format_value(total_activity_ytd)

    

    
        
        context = {
            "cf_totals":{
                "cfchecker":cfchecker,
                "total_investing":total_investing,
                "total_operating":total_operating,
                "total_activity":total_activity,
                "total_operating_ytd":total_operating_ytd,
                "total_investing_ytd":total_investing_ytd,
                "total_activity_ytd":total_activity_ytd,




            }
        }

        
                

        data_key = "Expend"
        if school in schoolCategory["skyward"]:
            data_key = "Amount"

        keys_to_check_cashflow = [
            "total_operating1",
            "total_operating2",
            "total_operating3",
            "total_operating4",
            "total_operating5",
            "total_operating6",
            "total_operating7",
            "total_operating8",
            "total_operating9",
            "total_operating10",
            "total_operating11",
            "total_operating12",
            
        ]
        keys_to_check_cashflow2 = [
            "total_investing1",
            "total_investing2",
            "total_investing3",
            "total_investing4",
            "total_investing5",
            "total_investing6",
            "total_investing7",
            "total_investing8",
            "total_investing9",
            "total_investing10",
            "total_investing11",
            "total_investing12",
            
        ]
        for row in data_cashflow:
            for key in keys_to_check_cashflow:
                value = float(row[key])
                if value == 0:
                    row[key] = ""
                elif value > 0:
                    row[key] = "({:,.0f})".format(abs(float(row[key])))
                else:
                    row[key] = "{:,.0f}".format(abs(float(row[key])))
            
        for row in data_cashflow:
            for key in keys_to_check_cashflow2:
                value = float(row[key])
                if value == 0:
                    row[key] = ""
                elif value > 0:
                    row[key] = "({:,.0f})".format(abs(float(row[key])))
                else:
                    row[key] = "{:,.0f}".format(abs(float(row[key])))

        for row in data_cashflow:
            f1 = row["fytd_1"]
            f2 = row["fytd_2"]

            if f1 is None or f1 == 0:
                row["fytd_1"] = ""
            else:
                row["fytd_1"] = format_value(f1) 

            if f2 is None or f2 == 0:
                row["fytd_2"] = ""
            else:
                row["fytd_2"] = format_value(f2) 


        if FY_year_1 == present_year:
            relative_path = os.path.join('cashflow', school)
        else:
            relative_path = os.path.join(str(FY_year_1), 'cashflow', school)
        

        # cashflow_path = JSON_DIR.path(relative_path)
        cashflow_path = os.path.join(JSON_DIR, relative_path)
        
        shutil.rmtree(cashflow_path, ignore_errors=True)
        if not os.path.exists(cashflow_path):
            os.makedirs(cashflow_path)


        cashflow_file = os.path.join(cashflow_path, "data_cashflow.json")
        
        with open(cashflow_file, "w") as f:
            json.dump(data_cashflow, f)

        for key, val in context.items():
            cashflow_file = os.path.join(cashflow_path, f"{key}.json")  
            with open(cashflow_file, "w") as f:
                json.dump(val, f)

            

    cursor.close()
    cnxn.close()

def excel(school,year):
    print("excel")
    present_date = datetime.today().date()   
    present_year = present_date.year
    today_date = datetime.now()
    print("today_date",today_date)
    print("present_date",present_date)
    
    today_month = today_date.month
 
    if year:
        year = int(year)
        if today_month == 1:
            start_year = year - 1
            present_year = present_year - 1
            FY_year_current = year - 1
        else: 
            start_year = year 
            FY_year_current = year
   
        if school in schoolMonths["julySchool"]:
            current_date = datetime(start_year, 7, 1).date()
            
        else:
            current_date = datetime(start_year, 9, 1).date() 
        current_year = current_date.year
    else:
        start_year = 2021
        current_date = datetime.today().date()   
        current_year = current_date.year
        FY_year_current = current_year
    while start_year <= FY_year_current:
     
        FY_year_1 = start_year
        FY_year_2 = start_year + 1 
        july_date_start  = datetime(FY_year_1, 7, 1).date()
        
        july_date_end  = datetime(FY_year_2, 6, 30).date()
        september_date_start  = datetime(FY_year_1, 9, 1).date()
        september_date_end  = datetime(FY_year_2, 8, 31).date()
        start_year = FY_year_2

        cnxn = connect()
        cursor = cnxn.cursor()
        cursor.execute(f"SELECT  * FROM [dbo].{db[school]['object']};")
        rows = cursor.fetchall()

        
        data = []
        for row in rows:
            if row[5] == school:

                row_dict = {
                    "fund": row[0],
                    "obj": row[1],
                    "description": row[2],
                    "category": row[3],
                    "value": row[4], #NOT BEING USED. DATA IS COMING FROM GL
                    "school":row[5],
                }
                data.append(row_dict)

        cursor.execute(f"SELECT  * FROM [dbo].{db[school]['function']};")
        rows = cursor.fetchall()

        data2 = []
        for row in rows:
            if row[5] == school:        
                row_dict = {
                    "func_func": row[0],
                    "obj": row[1],
                    "desc": row[2],
                    "category": row[3],
                    "budget":row[4], #NOT BEING USED. DATA IS COMING FROM GL
                    "school": row[5],

                }
                data2.append(row_dict)

        #
        if school in schoolCategory["ascender"]:
            # cursor.execute(
            #      f"SELECT * FROM [dbo].{db[school]['db']}  as AA where AA.Number != 'BEGBAL' and AA.Type != 'EN'  AND (UPPER(AA.WorkDescr) NOT LIKE '%BEG BAL%' AND UPPER(AA.WorkDescr) NOT LIKE '%BEGBAL%')"
            # )
                        cursor.execute(
                f"SELECT * FROM [dbo].{db[school]['db']}  as AA where AA.Number != 'BEGBAL' and AA.Type != 'EN'  "
            )
        else:
            cursor.execute(f"SELECT * FROM [dbo].{db[school]['db']} where source != 'RE';")

        rows = cursor.fetchall()

        data3 = []

        if school in schoolMonths["julySchool"]:
            current_month = july_date_start
        else:
            current_month = september_date_start
        
        last_month = ""
        last_month_name = ""
        last_month_number = ""
        formatted_last_month = ""


        if school in schoolCategory["ascender"]:
            for row in rows:
                expend = float(row[17])
                date = row[11]
                if isinstance(row[11], datetime):
                    date = row[11].strftime("%Y-%m-%d")
                acct_per_month_string = datetime.strptime(date, "%Y-%m-%d")
                acct_per_month = acct_per_month_string.strftime("%m")
        

                if isinstance(row[11],datetime):
                    date_checker = row[11].date()
                else:
                    date_checker = datetime.strptime(row[11], "%Y-%m-%d").date()
                   
                db_date = row[22].split('-')[0]

                db_date = int(db_date)
                curr_fy = int(FY_year_1)
   
                    
                if db_date == curr_fy:
                    if date_checker > current_month:
                        current_month = date_checker.replace(day=1)
                        
                    
                    
                    
                    row_dict = {
                        "fund": row[0],
                        "func": row[1],
                        "obj": row[2],
                        "sobj": row[3],
                        "org": row[4],
                        "fscl_yr": row[5],
                        "pgm": row[6],
                        "edSpan": row[7],
                        "projDtl": row[8],
                        "AcctDescr": row[9],
                        "Number": row[10],
                        "Date": date,
                        "AcctPer": row[12],
                        "Est": row[13],
                        "Real": row[14],
                        "Appr": row[15],
                        "Encum": row[16],
                        "Expend": expend,
                        "Bal": row[18],
                        "WorkDescr": row[19],
                        "Type": row[20],
                        "Contr": row[21],
                    }
                    data3.append(row_dict)
                

        

        else:        
            for row in rows:
                amount = float(row[19])
                date = row[9]
                
                if isinstance(row[9], datetime):
                    date = row[9].strftime("%Y-%m-%d")
                acct_per_month_string = datetime.strptime(date, "%Y-%m-%d")
                acct_per_month = acct_per_month_string.strftime("%m")

                if isinstance(row[9], (datetime, datetime.date)):
                    date_checker = row[9].date()
                else:
                    date_checker = datetime.strptime(row[9], "%Y-%m-%d").date()

                if school in schoolMonths["julySchool"]:
                
                    if date_checker >= july_date_start and date_checker <= july_date_end:
                        if date_checker > current_month:
                            current_month = date_checker.replace(day=1)

                        row_dict = {
                            "fund": row[0],
                            "func": row[2],
                            "obj": row[3],
                            "sobj": row[4],
                            "org": row[5],
                            "fscl_yr": row[6],
                            "Date": date,
                            "AcctPer":row[10],
                            "Amount": amount,
                            "Budget":row[20],
                        }

                        data3.append(row_dict)

                else:
                    if date_checker >= september_date_start and date_checker <= september_date_end:
                        if date_checker >= current_month:
                            current_month = date_checker.replace(day=1)

                        row_dict = {
                            "fund": row[0],
                            "func": row[2],
                            "obj": row[3],
                            "sobj": row[4],
                            "org": row[5],
                            "fscl_yr": row[6],
                            "Date": date,
                            "AcctPer": row[10],
                            "Amount": amount,
                            "Budget":row[20],
                        }

                        data3.append(row_dict)


        if FY_year_1 == present_year:
            print("current_month")
    
        else:
            if school in schoolMonths["julySchool"]:
                current_month = july_date_end
            else:
                current_month = september_date_end

        
        last_month = (current_month.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)                      
        last_month_name = last_month.strftime("%B")
        last_month_number = last_month.month
        formatted_last_month = last_month.strftime('%B %d, %Y')
        db_last_month = last_month.strftime("%Y-%m-%d")
        
  

        if present_year == FY_year_1:
            first_day_of_next_month = current_month.replace(day=1, month=current_month.month%12 + 1)
            last_day_of_current_month = first_day_of_next_month - timedelta(days=1)

            if current_month <= last_day_of_current_month:
                current_month = current_month.replace(day=1) - timedelta(days=1)
                last_month = (current_month.replace(day=1) + timedelta(days=32)).replace(day=1) - timedelta(days=1)                      
                last_month_name = last_month.strftime("%B")
                last_month_number = last_month.month
                formatted_last_month = last_month.strftime('%B %d, %Y')  
                db_last_month = last_month.strftime("%Y-%m-%d")

        cursor.execute(f"SELECT * FROM [dbo].{db[school]['adjustment']} ")
        rows = cursor.fetchall()

        adjustment = []

        if school in schoolCategory["ascender"]:
            for row in rows:
                expend = float(row[17])
                row_dict = {
                    "fund": row[0],
                    "func": row[1],
                    "obj": row[2],
                    "sobj": row[3],
                    "org": row[4],
                    "fscl_yr": row[5],
                    "pgm": row[6],
                    "edSpan": row[7],
                    "projDtl": row[8],
                    "AcctDescr": row[9],
                    "Number": row[10],
                    "Date": row[11],
                    "AcctPer": row[12],
                    "Est": row[13],
                    "Real": row[14],
                    "Appr": row[15],
                    "Encum": row[16],
                    "Expend": expend,
                    "Bal": row[18],
                    "WorkDescr": row[19],
                    "Type": row[20],
                    "School": row[21],
                }
                adjustment.append(row_dict)

        cursor.execute(f"SELECT * FROM [dbo].{db[school]['code']};")
        rows = cursor.fetchall()

        data_expensebyobject = []

        for row in rows:
            
            row_dict = {
                "obj": row[0],
                "Description": row[1],
                "budget": row[2],
            }

            data_expensebyobject.append(row_dict)

        cursor.execute(f"SELECT * FROM [dbo].{db[school]['activities']};")
        rows = cursor.fetchall()

        data_activities = []

        for row in rows:
            if row[3] == school:
                row_dict = {
                    "obj": row[0],
                    "Description": row[1],
                    "Category": row[2],
                    "school": row[3],
                }

                data_activities.append(row_dict)



        #END OF PL DATA

        #CHARTER FIRST

        #BS START

        cursor.execute(f"SELECT  * FROM [dbo].{db[school]['bs']} AS T1 LEFT JOIN [dbo].{db[school]['bs_fye']} AS T2 ON T1.BS_id = T2.BS_id ;  ")
        rows = cursor.fetchall()

        data_balancesheet = []

        for row in rows:
            if row[8] == school:
                fye = float(row[7]) if row[7] else 0
            
                if FY_year_1 == row[9]:
                    row_dict = {
                        "Activity": row[0],
                        "Description": row[1],
                        "Category": row[2],
                        "Subcategory": row[3],
                        "FYE": fye,
                        "BS_id": row[5],
                        "school": row[8],

                    }

                    data_balancesheet.append(row_dict)

        cursor.execute(f"SELECT * FROM [dbo].{db[school]['bs_activity']}")
        rows = cursor.fetchall()

        data_activitybs = []

        for row in rows:
            if row[3] == school:
                row_dict = {
                    "Activity": row[0],
                    "obj": row[1],
                    "Description2": row[2],
                    "school": row[3]
                }
                
                data_activitybs.append(row_dict)

        
        cursor.execute(f"SELECT * FROM [dbo].{db[school]['cashflow']};")
        rows = cursor.fetchall()

        data_cashflow = []

        for row in rows:
            row_dict = {
                "Category": row[0],
                "Activity": row[1],
                "Description": row[2],
                "obj": str(row[3]),
            }

            data_cashflow.append(row_dict)

    
        
        # current_date = datetime.today().date()
        
        # current_year = current_date.year
        # next_year = current_date.year + 1
        # last_year = current_date.year - 1
        # current_month = current_date.replace(day=1)
        # last_month = current_month - relativedelta(days=1)
        # last_month_name = last_month.strftime("%B")
        # formatted_last_month = last_month.strftime('%B %d, %Y')
        # last_month_number = last_month.month
        # if school == 'manara' or school == 'leadership':
        #         ytd_budget_test = last_month_number - 6             
        # else:
        #     if last_month_number >= 9:

        #         ytd_budget_test = last_month_number - 8
        #     else:
        #         ytd_budget_test = last_month_number + 4
        # ytd_budget = abs(ytd_budget_test) / 12
        # if ytd_budget_test == 1 or ytd_budget_test == 12:
        #     formatted_ytd_budget = f"{ytd_budget * 100:.0f}"
        
        # else:
        #     formatted_ytd_budget = (
        #     f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
        #     )
        #     if formatted_ytd_budget.startswith("0."):
        #         formatted_ytd_budget = formatted_ytd_budget[2:]

        if school in schoolMonths["julySchool"]:
                ytd_budget_test = last_month_number - 6             
        else:
            if last_month_number >= 9:

                ytd_budget_test = last_month_number - 8
            else:
                ytd_budget_test = last_month_number + 4
        ytd_budget = abs(ytd_budget_test) / 12


        if ytd_budget_test == 1 or ytd_budget_test == 12:
            formatted_ytd_budget = f"{ytd_budget * 100:.0f}"
        
        else:
            formatted_ytd_budget = (
            f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
            )
            if formatted_ytd_budget.startswith("0."):
                formatted_ytd_budget = formatted_ytd_budget[2:]

        expend_key = "Expend"
        est_key = "Est"
        expense_key = "Expend"
        real_key = "Real"
        appr_key = "Appr"
        encum_key = "Encum"
        bal_key = "Bal"
        if school in schoolCategory["skyward"]:
            expense_key = "Amount"
            expend_key = "Amount"
            est_key = "Budget"
            real_key = "Amount"
            appr_key = "Budget"
            encum_key = "Amount"
            bal_key = "Amount"

        
        acct_per_values = [
            "01",
            "02",
            "03",
            "04",
            "05",
            "06",
            "07",
            "08",
            "09",
            "10",
            "11",
            "12",
        ]


        # for item in data:
        #     fund = item["fund"]
        #     obj = item["obj"]
            
        #     for i, acct_per in enumerate(acct_per_values, start=1):
        #         total_real = sum(
        #             entry[real_key]
        #             for entry in data3
        #             if entry["fund"] == fund
        #             and entry["obj"] == obj
        #             and entry["AcctPer"] == acct_per
        #         )
        #         total_adjustment = sum(
        #                 entry[real_key]
        #                 for entry in adjustment
        #                 if entry["fund"] == fund
        #                 and entry["AcctPer"] == acct_per
        #                 and entry["obj"] == obj
        #                 and entry["School"] == school
        #             )
        #         item[f"total_check{i}"] = total_real + total_adjustment


        # july_date  = datetime(current_year, 7, 1).date()
        # september_date  = datetime(current_year, 9, 1).date()
        # FY_year_1 = last_year
        # FY_year_2 = current_year
        # for item in data3:
        #     date_str = item["Date"]
        #     if date_str:
        #         if school == 'manara' or school == 'leadership':
                
        #             date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
        #             if date_obj > july_date: # if date is higher than july 1 this year
        #             FY_year_1 = current_year
        #             FY_year_2 = next_year
        #         else:
        #             date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
        #             if date_obj > september_date: # if date is higher than july 1 this year
        #             FY_year_1 = current_year
        #             FY_year_2 = next_year
                    

        # #checks if the last month column is empty. if empty. last month will be set to  last two months.
        # if all(item[f"total_check{last_month_number}"] == 0 for item in data):
        #     last_2months = current_month - relativedelta(months=1)
        #     last_2months = last_2months - relativedelta(days=1)
        #     last_month_number = last_2months.month
        #     last_month_name = last_2months.strftime("%B")
        #     formatted_last_month = last_2months.strftime('%B %d, %Y')
        #     last_month_number = last_2months.month
        #     if school == 'manara' or school == 'leadership':
        #             ytd_budget_test = last_month_number - 6             
        #     else:
        #         if last_month_number >= 9:

        #             ytd_budget_test = last_month_number - 8
        #         else:
        #             ytd_budget_test = last_month_number + 4
        #     ytd_budget = abs(ytd_budget_test) / 12
            
        # if ytd_budget_test == 1 or ytd_budget_test == 12:
        #     formatted_ytd_budget = f"{ytd_budget * 100:.0f}"
        
        # else:
        #     formatted_ytd_budget = (
        #     f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
        #     )
        #     if formatted_ytd_budget.startswith("0."):
        #         formatted_ytd_budget = formatted_ytd_budget[2:]
            
        # CALCULATIONS START REVENUES 
        total_lr =  {acct_per: 0 for acct_per in acct_per_values}
        total_spr =  {acct_per: 0 for acct_per in acct_per_values}
        total_fpr =  {acct_per: 0 for acct_per in acct_per_values}
        total_revenue = {acct_per: 0 for acct_per in acct_per_values}
        ytd_total_revenue = 0
        ytd_total_lr  = 0
        ytd_total_spr = 0
        ytd_total_fpr = 0
        variances_revenue = 0

        totals = {
            "total_ammended": 0,
            "total_ammended_lr": 0,
            "total_ammended_spr": 0,
            "total_ammended_fpr": 0,
        }
                
                
        for item in data:
            fund = item["fund"]
            obj = item["obj"]
            category = item["category"]
            ytd_total = 0


            #PUT IT BACK WHEN YOU WANT TO GET THE GL FOR AMMENDED BUDGET FOR REVENUES
            if school in schoolCategory["skyward"]:
                
                total_budget = sum(
                    entry[est_key]
                    for entry in data3
                    if entry["fund"] == fund
                    and entry["obj"] == obj
                                
                )
                total_adjustment_budget = sum(
                    entry[est_key]
                    for entry in adjustment
                    if entry["fund"] == fund
                    and entry["obj"] == obj
                    and entry["School"] == school
                    and entry[est_key] is not None 
                    and not isinstance(entry[est_key], str) 
                                
                )
                item["total_budget"] = total_adjustment_budget + total_budget
            else:
                total_budget = sum(
                    entry[est_key]
                    for entry in data3
                    if entry["fund"] == fund
                    and entry["obj"] == obj
                    and entry["Type"] == "GJ"                
                )
                total_adjustment_budget = sum(
                    entry[est_key]
                    for entry in adjustment
                    if entry["fund"] == fund
                    and entry["obj"] == obj
                    and entry["School"] == school 
                    and entry[est_key] is not None 
                    and not isinstance(entry[est_key], str)              
                )
                item["total_budget"] = total_adjustment_budget + total_budget

            totals["total_ammended"] += item["total_budget"]
            item[f"ytd_budget"] = item["total_budget"] * ytd_budget
            

                
            if category == 'Local Revenue':
                totals["total_ammended_lr"] += item["total_budget"]
            elif category == 'State Program Revenue':
                totals["total_ammended_spr"] += item["total_budget"]
            elif category == 'Federal Program Revenue':
                totals["total_ammended_fpr"] += item["total_budget"]
            

            for i, acct_per in enumerate(acct_per_values, start=1):
                total_real = sum(
                    entry[real_key]
                    for entry in data3
                    if entry["fund"] == fund
                    and entry["obj"] == obj
                    and entry["AcctPer"] == acct_per
                )
                total_adjustment = sum(
                        entry[real_key]
                        for entry in adjustment
                        if entry["fund"] == fund
                        and entry["AcctPer"] == acct_per
                        and entry["obj"] == obj
                        and entry["School"] == school
                        and entry[real_key] is not None 
                        and not isinstance(entry[real_key], str) 
                    )
                item[f"total_real{i}"] = total_real + total_adjustment          
                total_revenue[acct_per] += (item[f"total_real{i}"])

                

                if category == 'Local Revenue':
                    total_lr[acct_per] += (item[f"total_real{i}"])
                    ytd_total_lr += (item[f"total_real{i}"])
                    
                if category == 'State Program Revenue':
                    total_spr[acct_per] += (item[f"total_real{i}"])
                    ytd_total_spr += (item[f"total_real{i}"])
                    
                if category == 'Federal Program Revenue':
                    total_fpr[acct_per] += (item[f"total_real{i}"])
                    ytd_total_fpr += (item[f"total_real{i}"])

            for month_number in range(1, 13):
                ytd_total += (item[f"total_real{month_number}"])
            
            item["ytd_total"] = ytd_total

            item["variances"] = item["ytd_total"] +item[f"ytd_budget"]

            item[f"ytd_budget"] = (item[f"ytd_budget"])
        
        
        ytd_total_revenue = abs(sum(total_revenue.values()))  
    
        ytd_ammended_total = totals["total_ammended"] * ytd_budget
        ytd_ammended_total_lr = totals["total_ammended_lr"] * ytd_budget
        ytd_ammended_total_spr = totals["total_ammended_spr"] * ytd_budget
        ytd_ammended_total_fpr = totals["total_ammended_fpr"] * ytd_budget

        variances_revenue = (ytd_total_revenue - ytd_ammended_total)
        variances_revenue_lr = (ytd_total_lr + ytd_ammended_total_lr)
        variances_revenue_spr = (ytd_total_spr + ytd_ammended_total_spr)
        variances_revenue_fpr = (ytd_total_fpr + ytd_ammended_total_fpr)

        var_ytd = (abs(int(ytd_total_revenue / totals["total_ammended"]*100))) if totals["total_ammended"] != 0 else ""
        var_ytd_lr = (abs(int(ytd_total_lr / totals["total_ammended_lr"]*100))) if totals["total_ammended_lr"] != 0 else ""
        var_ytd_spr = (abs(int(ytd_total_spr / totals["total_ammended_spr"]*100))) if totals["total_ammended_spr"] != 0 else ""
        var_ytd_fpr = (abs(int(ytd_total_fpr / totals["total_ammended_fpr"]*100))) if totals["total_ammended_fpr"] != 0 else ""
        #REVENUES CALCULATIONS END
        
        
        # CALCULATION START FIRST TOTAL AND DEPRECIATION AND AMORTIZATION (SBD) 
        first_total = 0
        first_ytd_total = 0
        first_total_months =  {acct_per: 0 for acct_per in acct_per_values}
        ytd_ammended_total_first=0
        variances_first_total = 0
        var_ytd_first_total = 0

        dna_total = 0
        dna_ytd_total = 0
        dna_total_months =  {acct_per: 0 for acct_per in acct_per_values}
        ytd_ammended_dna=0
        variances_dna = 0
        var_ytd_dna = 0

        for item in data2:
            if item["category"] != "Depreciation and Amortization":
                func = item["func_func"]
                obj = item["obj"]
                
                ytd_total = 0
                
                if school in schoolCategory["skyward"]:
                    total_func_func = sum(
                            entry[appr_key]
                            for entry in data3
                            if entry["func"] == func  
                            and entry["obj"] != '6449'
                        )
                else:
                    total_func_func = sum(
                            entry[appr_key]
                            for entry in data3
                            if entry["func"] == func  
                            and entry["obj"] != '6449'
                            and entry["Type"] == 'GJ' 
                        
                        )
                total_adjustment_func = sum(
                        entry[appr_key]
                        for entry in adjustment
                        if entry["func"] == func  
                        and entry["obj"] != '6449' 
                        and entry["School"] == school
                        and entry[appr_key] is not None 
                        and not isinstance(entry[appr_key], str)  
                    )
                if school in schoolCategory["skyward"]:
                    item['total_budget'] = total_func_func + total_adjustment_func
                else:
                    item['total_budget'] = -(total_func_func + total_adjustment_func)
            
                for i, acct_per in enumerate(acct_per_values, start=1):
                    total_func = sum(
                        entry[expend_key]
                        for entry in data3
                        if entry["func"] == func and entry["AcctPer"] == acct_per and entry["obj"] != '6449'
                    )
                    total_adjustment = sum(
                        entry[expend_key]
                        for entry in adjustment
                        if entry["func"] == func and entry["AcctPer"] == acct_per and entry["obj"] != '6449' and entry["School"] == school
                    )
                    item[f"total_func{i}"] = total_func + total_adjustment
                    first_total_months[acct_per] += item[f"total_func{i}"]

                for month_number in range(1, 13):
                    ytd_total += (item[f"total_func{month_number}"])
            
                item["ytd_total"] = ytd_total
                first_total += item['total_budget']
                first_ytd_total += item["ytd_total"]
                item[f"ytd_budget"] = item['total_budget'] * ytd_budget

                item["variances"] =  item[f"ytd_budget"] -item["ytd_total"]
                variances_first_total += item["variances"]
                item["var_ytd"] =  (abs(int(item["ytd_total"] /  item['total_budget']*100))) if item['total_budget'] != 0 else ""
        
        ytd_ammended_total_first = first_total * ytd_budget
        var_ytd_first_total = (abs(int(first_ytd_total / ytd_ammended_total_first*100))) if ytd_ammended_total_first != 0 else ""


        for item in data2:
            if item["category"] == "Depreciation and Amortization":
                func = item["func_func"]
                obj = item["obj"]
                ytd_total = 0
                
                if school in schoolCategory["skyward"]:
                    total_func_func = sum(
                            entry[appr_key]
                            for entry in data3
                            if entry["func"] == func  
                            and entry["obj"] == '6449'
                        
                        )
                else:
                    total_func_func = sum(
                        entry[appr_key]
                        for entry in data3
                        if entry["func"] == func  
                        and entry["obj"] == '6449'
                        and entry["Type"] == 'GJ'
                    )
                total_adjustment_func = sum(
                        entry[appr_key]
                        for entry in adjustment
                        if entry["func"] == func  
                        and entry["obj"] == '6449' 
                        and entry["School"] == school
                        and entry[appr_key] is not None 
                        and not isinstance(entry[appr_key], str)
                    )
                item['total_budget'] = -(total_func_func + total_adjustment_func)
                
                for i, acct_per in enumerate(acct_per_values, start=1):
                    total_func = sum(
                        entry[expend_key]
                        for entry in data3
                        if entry["func"] == func
                        and entry["AcctPer"] == acct_per
                        and entry["obj"] == obj
                    )
                    total_adjustment = sum(
                        entry[expend_key]
                        for entry in adjustment
                        if entry["func"] == func
                        and entry["AcctPer"] == acct_per
                        and entry["obj"] == obj
                        and entry["School"] == school
                        and entry[expend_key] is not None 
                        and not isinstance(entry[expend_key], str)
                    )
                
                    item[f"total_func2_{i}"] = total_func + total_adjustment
                    dna_total_months[acct_per] += item[f"total_func2_{i}"]
                
                

                for month_number in range(1, 13):
                    ytd_total += (item[f"total_func2_{month_number}"])
            
                item["ytd_total"] = ytd_total
                dna_total += item['total_budget']
                dna_ytd_total += item["ytd_total"]
                item[f"ytd_budget"] = item['total_budget'] * ytd_budget
                item["variances"] =  item[f"ytd_budget"] -item["ytd_total"]
                variances_dna+= item["variances"]
                item["var_ytd"] = (abs(int( item["ytd_total"]/item['total_budget'] *100))) if item['total_budget']  != 0 else ""
                ytd_ammended_dna = dna_total * ytd_budget
                var_ytd_dna = (abs(int(dna_ytd_total / ytd_ammended_dna*100))) if ytd_ammended_dna != 0 else ""
        #CALCULATION END FIRST TOTAL AND DNA
        
        #CALCULATION START SURPLUS BEFORE DEFICIT
        total_SBD =  {acct_per: 0 for acct_per in acct_per_values}
        ammended_budget_SBD = 0
        ytd_ammended_SBD = 0 
        ytd_SBD = 0 
        variances_SBD = 0 
        var_SBD = 0

        total_SBD = {
            acct_per: abs(total_revenue[acct_per]) - first_total_months[acct_per]
            for acct_per in acct_per_values
        }

        ammended_budget_SBD = abs(totals["total_ammended"]) - abs(first_total) 

        ytd_ammended_SBD =  abs(ytd_ammended_total) - abs(ytd_ammended_total_first)

        ytd_SBD = ytd_total_revenue - first_ytd_total
        variances_SBD =  ytd_SBD - ytd_ammended_SBD
        var_SBD = (abs(int(  ytd_SBD/ ammended_budget_SBD*100))) if ammended_budget_SBD != 0 else ""
        #CALCULATION END SURPLUS BEFORE DEFICIT


        #CALCULATION START NET SURPLUS
        total_netsurplus_months =  {acct_per: 0 for acct_per in acct_per_values}
        ammended_budget_netsurplus = 0
        ytd_ammended_netsurplus = 0 
        ytd_netsurplus = 0
        variances_netsurplus = 0
        var_netsurplus = 0

        total_netsurplus_months = {
            acct_per: total_SBD[acct_per] - dna_total_months[acct_per]
            for acct_per in acct_per_values
        }
        ammended_budget_netsurplus = ammended_budget_SBD - dna_total
        ytd_ammended_netsurplus = ytd_ammended_SBD - ytd_ammended_dna
        ytd_netsurplus =  ytd_SBD - dna_ytd_total 
        variances_netsurplus = ytd_netsurplus - ytd_ammended_netsurplus
        var_netsurplus = (abs(int(ytd_netsurplus / ammended_budget_netsurplus*100))) if ammended_budget_netsurplus != 0 else ""

        #CALCULATION EXPENSE BY OBJECT(EOC) AND TOTAL EXPENSE

        total_EOC_pc =  {acct_per: 0 for acct_per in acct_per_values} # PAYROLL COSTS
        total_EOC_pcs =  {acct_per: 0 for acct_per in acct_per_values}#Professional and Cont Svcs
        total_EOC_sm =  {acct_per: 0 for acct_per in acct_per_values}#Supplies and Materials
        total_EOC_ooe =  {acct_per: 0 for acct_per in acct_per_values}#Other Operating Expenses
        total_EOC_te =  {acct_per: 0 for acct_per in acct_per_values}#Total Expense
        total_EOC_oe =  {acct_per: 0 for acct_per in acct_per_values}#Other expenses 6449
        ytd_EOC_pc   = 0
        ytd_EOC_pcs  = 0
        ytd_EOC_sm   = 0
        ytd_EOC_ooe  = 0
        ytd_EOC_te   = 0
        ytd_EOC_oe = 0

        #FOR TOTAL EXPENSE
        total_expense = 0 
        total_expense_ytd_budget = 0
        total_expense_months =  {acct_per: 0 for acct_per in acct_per_values}
        total_expense_ytd = 0

        total_budget_pc  = 0
        total_budget_pcs = 0
        total_budget_sm = 0
        total_budget_ooe = 0
        total_budget_oe = 0
        total_budget_te = 0

        ytd_budget_pc = 0
        ytd_budget_pcs = 0
        ytd_budget_sm = 0
        ytd_budget_ooe = 0 
        ytd_budget_oe = 0 
        ytd_budget_te = 0

        for item in data_activities:
            obj = item["obj"]
            category = item["Category"]
            ytd_total = 0
            
            item["total_budget"] = 0

            if school in schoolCategory["skyward"]:
                total_budget_data_activities = sum(
                    entry[appr_key]
                    for entry in data3
                    if entry["obj"] == obj

                    )
            else:
                total_budget_data_activities = sum(
                entry[appr_key]
                for entry in data3
                if entry["obj"] == obj
                and entry["Type"] == 'GJ'
        
                )

            item["total_budget"] = -(total_budget_data_activities)
            
            item["ytd_budget"] =  item["total_budget"] * ytd_budget
            total_expense += item["total_budget"]  
            total_expense_ytd_budget += item[f"ytd_budget"]
            if category == "Payroll and Benefits":
                total_budget_pc += item["total_budget"]                

            if category == "Professional and Contract Services":          
                total_budget_pcs += item["total_budget"] 

            if category == "Materials and Supplies":       
                total_budget_sm += item["total_budget"]     
                
            if category == "Other Operating Costs":
                total_budget_ooe += item["total_budget"]  

            if category == "Depreciation":  
                total_budget_oe += item["total_budget"]     
                
            if category == "Debt Services": 
                total_budget_te += item["total_budget"]         

            for i, acct_per in enumerate(acct_per_values, start=1):
                total_activities = sum(
                    entry[expense_key]
                    for entry in data3
                    if entry["obj"] == obj and entry["AcctPer"] == acct_per
                )
                total_adjustment = sum(
                    entry[expense_key]
                    for entry in adjustment
                    if entry["obj"] == obj 
                    and entry["AcctPer"] == acct_per 
                    and entry["School"] == school
                    and entry[expense_key] is not None 
                    and not isinstance(entry[expense_key], str)
                )

                item[f"total_activities{i}"] = total_activities + total_adjustment


                
                if category == "Payroll and Benefits":
                    total_EOC_pc[acct_per] += item[f"total_activities{i}"]
                

                if category == "Professional and Contract Services":
                    total_EOC_pcs[acct_per] += item[f"total_activities{i}"]

                if category == "Materials and Supplies":
                    total_EOC_sm[acct_per] += item[f"total_activities{i}"]
                    

                if category == "Other Operating Costs":
                    total_EOC_ooe[acct_per] += item[f"total_activities{i}"]

                if category == "Depreciation":
                    total_EOC_oe[acct_per] += item[f"total_activities{i}"]
                    

                if category == "Debt Services":
                    total_EOC_te[acct_per] += item[f"total_activities{i}"]

                total_expense_months[acct_per] += item[f"total_activities{i}"]  

            for month_number in range(1, 13):
                ytd_total += (item[f"total_activities{month_number}"])
            item["ytd_total"] = ytd_total


        total_expense += dna_total
        total_expense_ytd_budget += ytd_ammended_dna
        for acct_per, dna_value in dna_total_months.items():
    
            if acct_per in total_expense_months:
            
                total_expense_months[acct_per] += dna_value

        ytd_EOC_pc  = sum(total_EOC_pc.values())
        ytd_EOC_pcs = sum(total_EOC_pcs.values())
        ytd_EOC_sm  = sum(total_EOC_sm.values())
        ytd_EOC_ooe = sum(total_EOC_ooe.values())
        ytd_EOC_te  = sum(total_EOC_te.values())
        ytd_EOC_oe  = sum(total_EOC_oe.values())

        
        ytd_budget_pc = total_budget_pc * ytd_budget
        ytd_budget_pcs = total_budget_pcs * ytd_budget
        ytd_budget_sm = total_budget_sm * ytd_budget
        ytd_budget_ooe = total_budget_ooe  * ytd_budget
        ytd_budget_oe = total_budget_oe * ytd_budget
        ytd_budget_te = total_budget_te * ytd_budget

            
        #temporarily for 6500
        budget_for_6500 = 0
        ytd_budget_for_6500 = 0 

        for item in data_expensebyobject:
            obj = item["obj"]
        
            if obj == "6100":
                category = "Payroll and Benefits"
                item["variances"] = ytd_budget_pc - ytd_EOC_pc
            
            elif obj == "6200":
                category = "Professional and Contract Services"
                item["variances"] = ytd_budget_pcs - ytd_EOC_pcs
            
            elif obj == "6300":
                category = "Materials and Supplies"
                item["variances"] = ytd_budget_sm - ytd_EOC_sm
                
            elif obj == "6400":
                category = "Other Operating Costs"
                item["variances"] = ytd_budget_ooe - ytd_EOC_ooe
                
            elif obj == "6449":
                category = "Depreciation"
                item["variances"] = ytd_budget_oe - ytd_EOC_oe
                
            else:
                category = "Debt Services"
                item["variances"] = ytd_budget_te - ytd_EOC_te
                

            for i, acct_per in enumerate(acct_per_values, start=1):
                item[f"total_expense{i}"] = sum(
                    entry[f"total_activities{i}"]
                    for entry in data_activities
                    if entry["Category"] == category
                )

            
        #CONTINUATION COMPUTATION TOTAL EXPENSE
        total_expense_ytd = sum([ytd_EOC_te, ytd_EOC_ooe, ytd_EOC_sm, ytd_EOC_pcs, ytd_EOC_pc,dna_ytd_total])
        variances_total_expense = total_expense_ytd_budget - total_expense_ytd
        var_total_expense = (abs(int(total_expense_ytd / total_expense*100))) if total_expense != 0 else ""
            

        #CALCULATIONS START NET INCOME
        net_income_budget = 0
        ytd_budget_net_income = 0 
        total_net_income_months =  {acct_per: 0 for acct_per in acct_per_values}
        ytd_net_income = 0
        variances_net_income = 0
        var_net_income = 0 


        budget_net_income = totals["total_ammended"] - total_expense
    
        ytd_budget_net_income = ytd_ammended_total - total_expense_ytd_budget
        ytd_net_income = ytd_total_revenue - total_expense_ytd
        variances_net_income = ytd_net_income - ytd_budget_net_income
        var_net_income = (abs(int(ytd_net_income / net_income_budget*100))) if net_income_budget != 0 else ""


        
        total_net_income_months = {
            acct_per: abs(total_revenue[acct_per]) - total_expense_months[acct_per]
            for acct_per in acct_per_values
        }  


        #BS CALCULATION START
        for item in data_activitybs:
            obj = item["obj"]
            item["fytd"] = 0

            for i, acct_per in enumerate(acct_per_values, start=1):
                total_data3 = sum(
                    entry[bal_key]
                    for entry in data3
                    if entry["obj"] == obj 
                    and entry["AcctPer"] == acct_per
                )
                total_adjustment = sum(
                    entry[bal_key]
                    for entry in adjustment
                    if entry["obj"] == obj
                    and entry["AcctPer"] == acct_per 
                    and entry["School"] == school
                    and entry[bal_key] is not None 
                    and not isinstance(entry[bal_key], str)
                )

                item[f"total_bal{i}"] = total_data3 + total_adjustment
                item["fytd"] += item[f"total_bal{i}"]
        
        activity_sum_dict = {}
        for item in data_activitybs:
            Activity = item["Activity"]
            
            for i in range(1, 13):
                total_sum_i = sum(
                    float(entry[f"total_bal{i}"])
                    if entry[f"total_bal{i}"] and entry["Activity"] == Activity
                    else 0
                    for entry in data_activitybs
                )
                activity_sum_dict[(Activity, i)] = total_sum_i

        for row in data_balancesheet:
            activity = row["Activity"]
            print(activity)
            for i in range(1, 13):
                key = (activity, i)
                row[f"total_sum{i}"] = (activity_sum_dict.get(key, 0))
                


        for row in data_balancesheet:
            if row["school"] == school:
                FYE_value = (float(row["FYE"])
                    if row["FYE"]
                    else 0
                )
                total_sum9_value = float(row["total_sum9"])
                total_sum10_value = float(row["total_sum10"])
                total_sum11_value = float(row["total_sum11"])
                total_sum12_value = float(row["total_sum12"])
                total_sum1_value = float(row["total_sum1"])
                total_sum2_value = float(row["total_sum2"])
                total_sum3_value = float(row["total_sum3"])
                total_sum4_value = float(row["total_sum4"])
                total_sum5_value = float(row["total_sum5"])
                total_sum6_value = float(row["total_sum6"])
                total_sum7_value = float(row["total_sum7"])
                total_sum8_value = float(row["total_sum8"])

                if school in schoolMonths['septemberSchool']:
                    # Calculate the differences and store them in the row dictionary
                    row["difference_9"] = (FYE_value + total_sum9_value)
                    row["difference_10"] =(row["difference_9"] + total_sum10_value)
                    row["difference_11"] =(row["difference_10"] + total_sum11_value)
                    row["difference_12"] =(row["difference_11"]  + total_sum12_value )
                    row["difference_1"] = (row["difference_12"] + total_sum1_value )
                    row["difference_2"] = (row["difference_1"] + total_sum2_value )
                    row["difference_3"] = (row["difference_2"] + total_sum3_value )
                    row["difference_4"] = (row["difference_3"] + total_sum4_value )
                    row["difference_5"] = (row["difference_4"] + total_sum5_value )
                    row["difference_6"] = (row["difference_5"] + total_sum6_value )
                    row["difference_7"] = (row["difference_6"] + total_sum7_value )
                    row["difference_8"] = (row["difference_7"] + total_sum8_value )
                    row["last_month_difference"] = row[f"difference_{last_month_number}"] 
        

                    row["fytd"] = ( total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value )

                    row["debt_9"]  = (FYE_value - total_sum9_value)
                    row["debt_10"] = (row["debt_9"] - total_sum10_value)
                    row["debt_11"] = (row["debt_10"] - total_sum11_value)
                    row["debt_12"] = (row["debt_11"] - total_sum12_value)
                    row["debt_1"] = (row["debt_12"] - total_sum1_value)
                    row["debt_2"] = (row["debt_1"] - total_sum2_value)
                    row["debt_3"] = (row["debt_2"] - total_sum3_value)
                    row["debt_4"] = (row["debt_3"]- total_sum4_value)
                    row["debt_5"] = (row["debt_4"]  - total_sum5_value )
                    row["debt_6"] = (row["debt_5"]- total_sum6_value)
                    row["debt_7"] = (row["debt_6"] - total_sum7_value)
                    row["debt_8"] = (row["debt_7"] - total_sum8_value)
                    row["last_month_debt"] = row[f"debt_{last_month_number}"]
                    row["debt_fytd"] = -( total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value)

            
                    row["net_assets9"] = (FYE_value + total_netsurplus_months["09"])
                    row["net_assets10"] = (row["net_assets9"] + total_netsurplus_months["10"])
                    row["net_assets11"] = (row["net_assets10"]+ total_netsurplus_months["11"])
                    row["net_assets12"] = (row["net_assets11"]+ total_netsurplus_months["12"])
                    row["net_assets1"] = (row["net_assets12"] + total_netsurplus_months["01"])
                    row["net_assets2"] = (row["net_assets1"] + total_netsurplus_months["02"])
                    row["net_assets3"] = (row["net_assets2"]+ total_netsurplus_months["03"])
                    row["net_assets4"] = (row["net_assets3"] + total_netsurplus_months["04"])
                    row["net_assets5"] = (row["net_assets4"] + total_netsurplus_months["05"])
                    row["net_assets6"] = (row["net_assets5"]  + total_netsurplus_months["06"])
                    row["net_assets7"] = (row["net_assets6"] + total_netsurplus_months["07"])
                    row["net_assets8"] = (row["net_assets7"] + total_netsurplus_months["08"])
                    row["last_month_net_assets"] = row[f"net_assets{last_month_number}"]
                else:
                                    # Calculate the differences and store them in the row dictionary
                    row["difference_7"] = (FYE_value + total_sum7_value )
                    row["difference_8"] = (row["difference_7"] + total_sum8_value )
                    row["difference_9"] = (row["difference_8"]  + total_sum9_value)
                    row["difference_10"] =(row["difference_9"] + total_sum10_value)
                    row["difference_11"] =(row["difference_10"] + total_sum11_value)
                    row["difference_12"] =(row["difference_11"]  + total_sum12_value )
                    row["difference_1"] = (row["difference_12"] + total_sum1_value )
                    row["difference_2"] = (row["difference_1"] + total_sum2_value )
                    row["difference_3"] = (row["difference_2"] + total_sum3_value )
                    row["difference_4"] = (row["difference_3"] + total_sum4_value )
                    row["difference_5"] = (row["difference_4"] + total_sum5_value )
                    row["difference_6"] = (row["difference_5"] + total_sum6_value )
                    row["last_month_difference"] = row[f"difference_{last_month_number}"] 
                    

                    row["fytd"] = ( total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value )


                    row["debt_7"] = (FYE_value - total_sum7_value)
                    row["debt_8"] = (row["debt_7"] - total_sum8_value)
                    row["debt_9"]  = (row["debt_8"] - total_sum9_value)
                    row["debt_10"] = (row["debt_9"] - total_sum10_value)
                    row["debt_11"] = (row["debt_10"] - total_sum11_value)
                    row["debt_12"] = (row["debt_11"] - total_sum12_value)
                    row["debt_1"] = (row["debt_12"] - total_sum1_value)
                    row["debt_2"] = (row["debt_1"] - total_sum2_value)
                    row["debt_3"] = (row["debt_2"] - total_sum3_value)
                    row["debt_4"] = (row["debt_3"]- total_sum4_value)
                    row["debt_5"] = (row["debt_4"]  - total_sum5_value )
                    row["debt_6"] = (row["debt_5"]- total_sum6_value)
                    row["last_month_debt"] = row[f"debt_{last_month_number}"] 
    
                    row["debt_fytd"] = -( total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value)


                    row["net_assets7"] = (FYE_value + total_netsurplus_months["07"])
                    row["net_assets8"] = (row["net_assets7"] + total_netsurplus_months["08"])
                    row["net_assets9"] = (row["net_assets8"]  + total_netsurplus_months["09"])
                    row["net_assets10"] = (row["net_assets9"] + total_netsurplus_months["10"])
                    row["net_assets11"] = (row["net_assets10"]+ total_netsurplus_months["11"])
                    row["net_assets12"] = (row["net_assets11"]+ total_netsurplus_months["12"])
                    row["net_assets1"] = (row["net_assets12"] + total_netsurplus_months["01"])
                    row["net_assets2"] = (row["net_assets1"] + total_netsurplus_months["02"])
                    row["net_assets3"] = (row["net_assets2"]+ total_netsurplus_months["03"])
                    row["net_assets4"] = (row["net_assets3"] + total_netsurplus_months["04"])
                    row["net_assets5"] = (row["net_assets4"] + total_netsurplus_months["05"])
                    row["net_assets6"] = (row["net_assets5"]  + total_netsurplus_months["06"])
                    row["last_month_net_assets"] = row[f"net_assets{last_month_number}"]

        total_current_assets = {acct_per: 0 for acct_per in acct_per_values}
        total_current_assets_fye = 0
        total_current_assets_fytd = 0 
        

        total_capital_assets = {acct_per: 0 for acct_per in acct_per_values}
        total_capital_assets_fye = 0
        total_capital_assets_fytd = 0 

        total_current_liabilities = {acct_per: 0 for acct_per in acct_per_values}
        total_current_liabilities_fye = 0
        total_current_liabilities_fytd = 0

        total_liabilities = {acct_per: 0 for acct_per in acct_per_values}
        total_liabilities_fye = 0
        total_liabilities_fytd = 0

        total_assets = {acct_per: 0 for acct_per in acct_per_values}
        total_assets_fye = 0
        total_assets_fye_fytd = 0
        

        total_LNA = {acct_per: 0 for acct_per in acct_per_values} # LIABILITES AND NET ASSETS 
        total_LNA_fye = 0
        total_LNA_fytd = 0


        
        total_net_assets_fytd = 0
        total_net_assets_fytd = ytd_netsurplus    
        


        for row in data_balancesheet:
            if row["school"] == school:
                subcategory =  row["Subcategory"]
                fye =  float(row["FYE"]) if row["FYE"] else 0

                if subcategory == 'Current Assets':
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_current_assets[acct_per] += row[f"difference_{i}"]
                    total_current_assets_fytd += row["fytd"]

                    total_current_assets_fye +=  fye
                if subcategory == 'Capital Assets, Net':
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_capital_assets[acct_per] += row[f"difference_{i}"]
                    total_capital_assets_fytd += row["fytd"]
                    total_capital_assets_fye +=  fye
                if subcategory == 'Current Liabilities':
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_current_liabilities[acct_per] += row[f"debt_{i}"]
                    total_current_liabilities_fytd += row["debt_fytd"]
                    total_current_liabilities_fye +=  fye

        
        total_liabilities_fytd_2 = 0
        for row in data_balancesheet:
            if row["school"] == school:
                subcategory =  row["Subcategory"]
                fye =  float(row["FYE"]) if row["FYE"] else 0
                if subcategory == 'Long Term Debt':
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_liabilities[acct_per] += row[f"debt_{i}"] + total_current_liabilities[acct_per]
                    total_liabilities_fytd_2 += row["debt_fytd"]
                    total_liabilities_fye +=  + total_current_liabilities_fye + fye
        total_liabilities_fytd = total_liabilities_fytd_2 + total_current_liabilities_fytd

        for row in data_balancesheet:
            if row["school"] == school:
                fye =  float(row["FYE"]) if row["FYE"] else 0
                if  row["Category"] == "Net Assets":
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_LNA[acct_per] += row[f"net_assets{i}"] + total_liabilities[acct_per]

                    total_LNA_fye += total_liabilities_fye + fye

    
        total_assets = {
            acct_per: total_current_assets[acct_per] + total_capital_assets[acct_per]
            for acct_per in acct_per_values

        }
        total_assets_fye = total_current_assets_fye + total_capital_assets_fye
        total_assets_fye_fytd = total_current_assets_fytd + total_capital_assets_fytd

        net = float(total_net_assets_fytd) if total_net_assets_fytd else 0
        total_LNA_fytd = net + total_liabilities_fytd


        # FOR CASHFLOW

        for item in data_cashflow:
            activity = item["Activity"]
            item["fytd_1"] = 0
            obj = item["obj"]
            item["fytd_2"] = 0

            for i, acct_per in enumerate(acct_per_values, start=1):
                total_bal_key = f"total_bal{i}"
                item[f"total_operating{i}"] = sum(
                    entry[total_bal_key]
                    for entry in data_activitybs
                    if entry["Activity"] == activity
                )
                item["fytd_1"] += item[f"total_operating{i}"]
            for i, acct_per in enumerate(acct_per_values, start=1):
                item[f"total_investing{i}"] = sum(
                    entry[bal_key]
                    for entry in data3
                    if entry["obj"] == obj and entry["AcctPer"] == acct_per
                )
                item["fytd_2"] += item[f"total_investing{i}"]  
        
        sorted_data2 = sorted(data2, key=lambda x: x['func_func'])
        sorted_data = sorted(data, key=lambda x: x['obj'])


        cursor.execute(f"SELECT * FROM [dbo].[AscenderData_CharterFirst]  \
                WHERE school = '{school}' \
                AND year = '{present_year}' \
                AND month = {last_month_number};") 
        row = cursor.fetchone()

        data_charterfirst = []
    

        # for row in rows:
        #     print(row[1])

        #     if row[0] == school and row[2]== (last_month_number) and row[1] == FY_year_1:
        #         print("enter")
        if row is not None:

            row_dict = {
                "school": row[0],
                "year": row[1],
                "month": row[2],
                "net_income_ytd":row[3],
                "indicators": row[4],
                "net_assets": row[5],
                "days_coh": row[6],
                "current_assets": row[7],
                "net_earnings": row[8],
                "budget_vs_revenue": row[9],
                "total_assets": row[10],
                "debt_service": row[11],
                "debt_capitalization": row[12],
                "ratio_administrative": row[13],
                "ratio_student_teacher": row[14],
                "estimated_actual_ada": row[15],
                "reporting_peims": row[16],
                "annual_audit": row[17],
                "post_financial_info": row[18],
                "approved_geo_boundaries": row[19],
                "estimated_first_rating": row[20],
            }
            data_charterfirst.append(row_dict)


        context = {
            "data": sorted_data,
            "data2": sorted_data2,
            "data3": data3,
            "data_expensebyobject": data_expensebyobject,
            "data_activities": data_activities,
            "data_charterfirst":data_charterfirst,
            "data_balancesheet":data_balancesheet,
            "data_activitybs":data_activitybs,
            "data_cashflow":data_cashflow,


            "months":
                    {
                "last_month": formatted_last_month,
                "last_month_number": last_month_number,
                "last_month_name": last_month_name,
                "format_ytd_budget": formatted_ytd_budget,
                "ytd_budget": ytd_budget,
                "FY_year_1":FY_year_1,
                "FY_year_2":FY_year_2,

                },
            "totals":{
                #FOR REVENUES
                "total_lr": total_lr,
                "total_spr": total_spr,
                "total_fpr": total_fpr,
                "total_revenue": total_revenue,
                "total_ammended": totals["total_ammended"],
                "total_ammended_lr": totals["total_ammended_lr"],
                "total_ammended_spr": totals["total_ammended_spr"],
                "total_ammended_fpr": totals["total_ammended_fpr"],
                "ytd_ammended_total":ytd_ammended_total,
                "ytd_ammended_total_lr":ytd_ammended_total_lr,
                "ytd_ammended_total_spr":ytd_ammended_total_spr,
                "ytd_ammended_total_fpr":ytd_ammended_total_fpr,
                "ytd_total_revenue": ytd_total_revenue,
                "ytd_total_lr": ytd_total_lr,
                "ytd_total_spr": ytd_total_spr,
                "ytd_total_fpr": ytd_total_fpr,
                "variances_revenue":variances_revenue,
                "variances_revenue_lr":variances_revenue_lr,
                "variances_revenue_spr":variances_revenue_spr,
                "variances_revenue_fpr":variances_revenue_fpr,
                "var_ytd":var_ytd,
                "var_ytd_lr":var_ytd_lr,
                "var_ytd_spr":var_ytd_spr,
                "var_ytd_fpr":var_ytd_fpr,

                #FIRST TOTAL
                "first_total":first_total,
                "first_total_months":first_total_months,
                "first_ytd_total":first_ytd_total,
                "ytd_ammended_total_first": ytd_ammended_total_first,
                "variances_first_total":variances_first_total,
                "var_ytd_first_total": var_ytd_first_total,

                # DEPRECIATION AND AMORTIZATION
                "dna_total":dna_total,
                "dna_total_months":dna_total_months,
                "dna_ytd_total":dna_ytd_total,
                "ytd_ammended_dna": ytd_ammended_dna,
                "variances_dna":variances_dna,
                "var_ytd_dna":var_ytd_dna,

                #SURPLUS BEFORE DEFICIT(SBD)
                "total_SBD": total_SBD,
                "ammended_budget_SBD": ammended_budget_SBD,
                "ytd_ammended_SBD": ytd_ammended_SBD,
                "ytd_SBD":ytd_SBD,
                "variances_SBD": variances_SBD,
                "var_SBD":var_SBD,

                #NET SURPLUS    
                "total_netsurplus_months": total_netsurplus_months,
                "ammended_budget_netsurplus": ammended_budget_netsurplus,
                "ytd_ammended_netsurplus" : ytd_ammended_netsurplus,
                "ytd_netsurplus": ytd_netsurplus,
                "variances_netsurplus": variances_netsurplus,
                "var_netsurplus":var_netsurplus,

                #EXPENSE BY OBJECT 
                "total_EOC_pc":total_EOC_pc,
                "total_EOC_pcs":total_EOC_pcs,
                "total_EOC_sm":total_EOC_sm,
                "total_EOC_ooe":total_EOC_ooe,
                "total_EOC_te":total_EOC_te,
                "total_EOC_oe":total_EOC_oe,
                "ytd_EOC_pc":ytd_EOC_pc,
                "ytd_EOC_pcs":ytd_EOC_pcs,
                "ytd_EOC_sm":ytd_EOC_sm,
                "ytd_EOC_ooe":ytd_EOC_ooe,
                "ytd_EOC_te":ytd_EOC_te,
                "ytd_EOC_oe":ytd_EOC_oe,
                "total_budget_pc":total_budget_pc,
                "total_budget_pcs":total_budget_pcs,
                "total_budget_sm":total_budget_sm,
                "total_budget_ooe":total_budget_ooe,
                "total_budget_oe":total_budget_oe,
                "total_budget_te":total_budget_te,
                "ytd_budget_pc":ytd_budget_pc,
                "ytd_budget_pcs":ytd_budget_pcs,
                "ytd_budget_sm":ytd_budget_sm,
                "ytd_budget_ooe":ytd_budget_ooe,
                "ytd_budget_oe":ytd_budget_oe,
                "ytd_budget_te":ytd_budget_te,
                #FIX SOON
                "budget_for_6500":budget_for_6500,
                "ytd_budget_for_6500": ytd_budget_for_6500,
                
                #TOTAL EXPENSE 
                "total_expense": total_expense,
                "total_expense_ytd_budget": total_expense_ytd_budget,
                "total_expense_months":total_expense_months,
                "total_expense_ytd":total_expense_ytd,
                "variances_total_expense":variances_total_expense,
                "var_total_expense":var_total_expense,

                #NET INCOME
                "budget_net_income": budget_net_income,
                "ytd_budget_net_income":ytd_budget_net_income,
                "total_net_income_months":total_net_income_months,
                "variances_net_income": variances_net_income,
                "ytd_net_income": ytd_net_income,
                "var_net_income":var_net_income,            
            },
            "total_bs":{
                "total_current_assets":total_current_assets,
                "total_current_assets_fye":total_current_assets_fye,
                "total_capital_assets":total_capital_assets,
                "total_capital_assets_fye":total_capital_assets_fye,
                "total_current_liabilities":total_current_liabilities,
                "total_current_liabilities_fye":total_current_liabilities_fye,
                "total_liabilities":total_liabilities,
                "total_liabilities_fye":total_liabilities_fye,
                "total_assets": total_assets,
                "total_assets_fye":total_assets_fye,
                "total_LNA_fye":total_LNA_fye,
                "total_LNA":total_LNA,
                "total_current_assets_fytd":total_current_assets_fytd,
                "total_capital_assets_fytd":total_capital_assets_fytd,
                "total_current_liabilities_fytd":total_current_liabilities_fytd,
                "total_liabilities_fytd":total_liabilities_fytd,
                "total_assets_fye_fytd":total_assets_fye_fytd,
                "total_net_assets_fytd":total_net_assets_fytd,
                "total_LNA_fytd":total_LNA_fytd,
            }
        }
    
        if FY_year_1 == present_year: 
            relative_path = os.path.join( "excel", school)
        else:
            relative_path = os.path.join(str(FY_year_1), "excel", school)
        # json_path = JSON_DIR.path(relative_path)
        json_path = os.path.join(JSON_DIR, relative_path)

        shutil.rmtree(json_path, ignore_errors=True)
        if not os.path.exists(json_path):
            os.makedirs(json_path)

        for key, val in context.items():
            file = os.path.join(json_path, f"{key}.json")
            with open(file, "w") as f:
                json.dump(val, f)

    
def charter_first(school):
    context = {}
    # first check if previous month is in database already
    cnxn = connect()
    cursor = cnxn.cursor()

    current_date = datetime.now()
    month_number = current_date.month
    curr_year = current_date.year
    # never use global declaration as it changes the variable and  when the page does not reload it will result to an error when the function is reused

  
    # if january query to dec last year else last month of current year 
    if month_number == 1:
        month_number = 12
        curr_year = curr_year - 1
        prev_query = f"SELECT * from [dbo].[AscenderData_CharterFirst] WHERE month={month_number} AND year={curr_year} AND school='{school}';"
        print("first_query")
    else:
        month_number -= 1
        prev_query = f"SELECT * from [dbo].[AscenderData_CharterFirst] WHERE month={month_number} AND year={curr_year} AND school='{school}';"
        print("second_query")

   
    cursor.execute(prev_query)
    rows = cursor.fetchone()

    shouldUpdate = False
    if rows is not None:
        shouldUpdate = True


    first_columns = [
        "school", "year", "month", "net_income_ytd", "indicators", "net_assets", 
        "days_coh", "current_assets", "net_earnings", "budget_vs_revenue", "total_assets", 
        "debt_service", "debt_capitalization", "ratio_administrative", "ratio_student_teacher", 
        "estimated_actual_ada", "reporting_peims", "annual_audit", "post_financial_info", 
        "approved_geo_boundaries", "estimated_first_rating" 
    ]
    # if None then create one
    insert_query = f"INSERT INTO [dbo].[AscenderData_CharterFirst] \
    ({', '.join(first_columns)}) \
    VALUES ({', '.join(['?'] * 21)})"

    context["school"] = school
    context["year"] = int(curr_year)

    # should stay since the month is already declared as 12. otherwise other months should be last month

    context["month"] = int(month_number)
    
    
        
        


    totals_file_path = os.path.join("profit-loss", school, "totals.json")
    # totals_file = JSON_DIR.path(totals_file_path)
    totals_file = os.path.join(JSON_DIR, totals_file_path)
    with open(totals_file, "r") as f:
        pl_totals = json.load(f)

        
    data2_file_path = os.path.join("profit-loss", school, "data2.json")
    # totals_file = JSON_DIR.path(totals_file_path)
    data2_file = os.path.join(JSON_DIR, data2_file_path)
    with open(data2_file, "r") as f:
        data2 = json.load(f)


    data3_file_path = os.path.join("profit-loss", school, "data3.json")
    # totals_file = JSON_DIR.path(totals_file_path)
    data3_file = os.path.join(JSON_DIR, data3_file_path)
    with open(data3_file, "r") as f:
        data3 = json.load(f)

    

    # get cash equivalents
    bs_file_path = os.path.join( "balance-sheet", school, "data_balancesheet.json")
    # bs_file = JSON_DIR.path(bs_file_path)
    bs_file = os.path.join(JSON_DIR, bs_file_path)
    with open(bs_file, "r") as f:
        balance_sheet = json.load(f)

    cash_equivalents = ""
    for item in balance_sheet:
        if item["Activity"].strip().lower() == "cash" and item["Description"].strip().lower() == "cash and cash equivalents" and item["school"].strip().lower() == school:
            cash_equivalents = dollar_parser(item["last_month_difference"])
            break
    
    print("cash_equivalents",cash_equivalents)

    pl_activities_file_path = os.path.join( "profit-loss", school, "data_activities.json")
    # pl_activities_file = JSON_DIR.path(pl_activities_file_path)
    pl_activities_file = os.path.join(JSON_DIR, pl_activities_file_path)
    with open(pl_activities_file, "r") as f:
        pl_activities = json.load(f)



    # get fy start
    pl_months_file_path = os.path.join( "profit-loss", school, "months.json")
    # pl_months_file = JSON_DIR.path(pl_months_file_path)
    pl_months_file = os.path.join(JSON_DIR, pl_months_file_path)
    with open(pl_months_file, "r") as f:
        pl_months = json.load(f)



    date_string = pl_months["last_month"]
    FY_year_1 = pl_months["FY_year_1"]

    pl_lmn = pl_months["last_month_number"]

    key_month = "0" + str(pl_lmn) if pl_lmn < 10 else str(pl_lmn)


    
    ytd_expense = dollar_parser(pl_totals["first_ytd_total"])
    debt_services = dollar_parser(pl_totals["ytd_EOC_te"])


    date_format = "%B %d, %Y"
    fy_curr = datetime.strptime(date_string, date_format)
    print("fycurr",fy_curr)

    fy_start = ""
    if school in schoolMonths['julySchool']:

        fy_start = datetime(FY_year_1, 7, 1)    
    else: 
        fy_start = datetime(FY_year_1, 9, 1)

    fy_diff = fy_curr - fy_start
    print("FY_DIFF",fy_diff)
    days = fy_diff.days

    expense_per_day = ytd_expense / days

    if expense_per_day != 0:
        context["days_coh"] = cash_equivalents / expense_per_day
    else:
        # Handle the case where expense_per_day is zero
        context["days_coh"] = 0

 
    # current assets
    bs_totals_file_path = os.path.join("balance-sheet", school, "totals_bs.json")
    # bs_totals_file = JSON_DIR.path(bs_totals_file_path)
    bs_totals_file = os.path.join(JSON_DIR, bs_totals_file_path)
    with open(bs_totals_file, "r") as f:
        bs_totals = json.load(f)
 
    try:
        total_current_assets = dollar_parser(bs_totals["total_current_assets"][key_month])
        total_current_liabilities = dollar_parser(bs_totals["total_current_liabilities"][key_month])
    except KeyError:

        total_current_assets = 0
        total_current_liabilities = 0
    

    
    if total_current_liabilities != 0:
        current_ratio = total_current_assets / total_current_liabilities
    else:
        current_ratio = 0
  


    # net earnings 
 


    


    #debt_service = (deficitsurplus + 71 debt service)/ (-increase or decrease in bonds(cashflow) + debt service)
    debtservice = 0
    for row in data2:        
        if row["category"] != 'Depreciation and Amortization' and row["func_func"] == '71':
            debtservice = row[f"total_func{pl_lmn}"]
            
            break

    
    try:
        deficitsurplus = dollar_parser(pl_totals["total_SBD"][key_month])
    except KeyError:
        deficitsurplus = 0

    print(debtservice)

    try:
        if debtservice:
            debtservice = dollar_parser(debtservice)
        else:
            debtservice = 0
    except KeyError:
        debtservice = 0



    if debtservice != 0:
        debt_service = (deficitsurplus+debtservice)/debtservice
        debt_service = round(debt_service, 2)
    else:
        debt_service = 0


   


    ltd = 0
    equity = 0
    for item in balance_sheet:
        if item["Activity"].strip().lower() == "ltd" and item["Subcategory"].strip().lower() == "long term debt" and item["Category"].strip().lower() == "debt" and item["school"].strip().lower() == school:
            ltd = item[f"debt_{pl_lmn}"]
        if item["Activity"].strip().lower() == "equity" and item["Category"].strip().lower() == "net assets" and item["school"].strip().lower() == school:
            equity = item[f"net_assets{pl_lmn}"]

    

    if equity:
        equity = dollar_parser(equity)
    else:
        equity = 0

    if ltd:
        ltd = dollar_parser(ltd)
    else:
        ltd = 0

    # ltd/(ltd+equity) x100
    if ltd != 0:
        debt_capitalization = ltd/(ltd+equity)*100
        debt_capitalization_ratio_rounded = round(debt_capitalization, 2)
    else:
        debt_capitalization_ratio_rounded = 0




    # num 11 criteria
    try:
        total_assets = dollar_parser(bs_totals["total_assets"][key_month])    
    except KeyError:
        total_assets = 0
    
    if total_assets != 0:
        lt_ratio = ltd / total_assets
    else:
        lt_ratio = 0

    expend_key = "Expend"
    est_key = "Est"
    expense_key = "Expend"
    real_key = "Real"
    appr_key = "Appr"
    encum_key = "Encum"
    if school in schoolCategory["skyward"]:
        expense_key = "Amount"
        expend_key = "Amount"
        est_key = "Budget"
        real_key = "Amount"
        appr_key = "Budget"
        encum_key = "Amount"


    def calculate_first_func(func_value):
        return sum(
            entry[expend_key]
            for entry in data3
            if entry["fund"] in ['420', '199']
            and 6100 <= int(entry["obj"]) <= 6499
            and entry['func'] == func_value
            and entry["AcctPer"] == key_month
        )

    def calculate_second_func(func_value):
        return sum(
            entry[expend_key]
            for entry in data3
            if entry["fund"] in ['420', '199','266','281','282','283']
            and 6100 <= int(entry["obj"]) <= 6499
            and entry['func'] == func_value
            and entry["AcctPer"] == key_month
        )

    first_21 = calculate_first_func('21')
    first_41 = calculate_first_func('41')
    first_11 = calculate_first_func('11')
    first_12 = calculate_first_func('12')
    first_13 = calculate_first_func('13')
    first_31 = calculate_first_func('31')

    second_21 = calculate_second_func('21')
    second_41 = calculate_second_func('41')
    second_11 = calculate_second_func('11')
    second_12 = calculate_second_func('12')
    second_13 = calculate_second_func('13')
    second_31 = calculate_second_func('31')

    first_denominator = first_11 + first_12 + first_13 + first_31
    if first_denominator != 0:
        first_AR =  (first_21 + first_41) / (first_denominator) * 100    
    else: 
        first_AR = 0  
    
    second_denominator = second_11 + second_12 + second_13 + second_31
    
    if second_denominator != 0:
        
        second_AR = (second_21 + second_41) / (second_denominator) * 100
    else:
        second_AR = 0

    first_AR = round(first_AR, 2)
    second_AR = round(second_AR, 2)

    administrative_ratio = str(first_AR) + '% /' + str(second_AR) + '%'
    



    # debt_capitalization
    # skipped because there is no loan payable

    # administrative ratio
    # skipped because instructions are unclear
    context["net_income_ytd"] = dollar_parser(pl_totals["ytd_netsurplus"])
    context["current_assets"] = round(current_ratio, 2)
    context["net_earnings"] = dollar_parser(pl_totals["ytd_SBD"])
    context["total_assets"] = round(lt_ratio, 2)
    context["indicators"] = "Pass"
    context["net_assets"] = "projected"
    context["budget_vs_revenue"] = "projected"
    context["debt_service"] = debt_service
    context["debt_capitalization"] = debt_capitalization_ratio_rounded
    context["ratio_administrative"] = administrative_ratio
    context["ratio_student_teacher"] = "Not measured by DSS"
    context["estimated_actual_ada"] = "projected"
    context["reporting_peims"] = "projected"
    context["annual_audit"] = "projected"
    context["post_financial_info"] = "projected"
    context["approved_geo_boundaries"] = "Not measured by DSS"
    context["estimated_first_rating"] = 88

    # print(list((context[key] for key in first_columns)))
    if shouldUpdate:
        update_query = f"UPDATE [dbo].[AscenderData_CharterFirst] \
        SET {', '.join([f'{col} = ?' for col in first_columns])} \
        WHERE month={month_number} AND year={curr_year} AND school='{school}';"

        cursor.execute(update_query, tuple((context[key] for key in first_columns)))
        cnxn.commit()
        print("updated")

    else:
        cursor.execute(insert_query, tuple((context[key] for key in first_columns)))
        print("inserted")
        cnxn.commit()

    cursor.close()
    cnxn.close()

def dollar_parser(dollar):
    if dollar.strip() == "":
        return 0
    if "(" in dollar:
        formatted = "".join(dollar.strip().replace("$", "").replace("(", "").replace(")", "").split(","))
        if "." in formatted:
            return float(formatted) * -1
        return int(formatted) * -1
    
    formatted = "".join(dollar.strip().replace("$", "").split(","))

    if "." in formatted:
        return float(formatted) * -1
    return int(formatted)



def profit_loss_chart(school):
        
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['pl_chart']};")
    rows = cursor.fetchall()

    data_chart = []
    for row in rows:
        if row[2] == school:
           
            row_dict = {
                "date": row[0],
                "data":row[1],
                "school":row[2]

            }
            data_chart.append(row_dict)


    context = {
        "data_chart": data_chart,

    }
    pl_json_path = os.path.join( "profit-loss-chart", school)
    # json_path = JSON_DIR.path(pl_json_path)
    json_path = os.path.join(JSON_DIR, pl_json_path)
    shutil.rmtree(json_path, ignore_errors=True)
    if not os.path.exists(json_path):
        os.makedirs(json_path)

    for key, val in context.items():
        file = os.path.join(json_path, f"{key}.json")
        with open(file, "w") as f:
            json.dump(val, f)
            
def updateGraphDB(school, fye):
    
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['pl_chart']};")
    rows = cursor.fetchall()

    dataDate = {}
    for row in rows:        
        dataDate[row[0]] = 1 
    yearPath = ['', '2022', '2021']
    if fye == True:   
        yearPath = ['']
        

    for yPath in yearPath:
        with open(os.path.join(settings.BASE_DIR, 'finance', 'json', str(yPath), 'profit-loss', str(school), 'data.json'), 'r') as f:
            data = json.load(f)

        with open(os.path.join(settings.BASE_DIR, 'finance', 'json', str(yPath), 'profit-loss', str(school), 'data2.json'), 'r') as f:
            dataExpense = json.load(f)

        with open(os.path.join(settings.BASE_DIR, 'finance', 'json', str(yPath), 'profit-loss', str(school), 'data_activities.json'), 'r') as f:
            dataExpensebyObject = json.load(f)
        
        for x in range(1,13):
            insertqueryStatement = 'INSERT INTO dbo.PLData (date,data,school) VALUES (?,?,?)'
            localRevenue = '{ "localRevenue": { '
            stateprogramRevenue = '"stateprogramRevenue": { '
            federalprogramRevenue = '"federalprogramRevenue": { '
            expense = '"expense": { '
            payrollandBenefits = '"payrollandBenefits": { '
            professionalandcontractServices = '"professionalandcontractServices": { '
            materialsandSupplies = '"materialsandSupplies": { '
            otheroperatingCosts = '"otheroperatingCosts": { '
            depreciation = '"depreciation": { '
            debtServices = '"debtServices": { '

            localRevenueTotal = 0
            stateprogramRevenueTotal = 0
            federalprogramRevenueTotal = 0
            expenseTotal = 0
            payrollandBenefitsTotal = 0
            professionalandcontractServicesTotal = 0
            materialsandSuppliesTotal = 0
            otheroperatingCostsTotal = 0
            depreciationTotal = 0
            debtServicesTotal = 0        
            for i in data:        
                if i['total_real' + str(x)] != '':
                    amount = i['total_real' + str(x)].replace(',','')
                    if amount.find('(') != -1:
                        amount = '-' + amount.replace('(','').replace(')','')

                    if i['category'] == 'Local Revenue':
                        localRevenue = localRevenue + '"' + str(i['obj']) + '": ' +  amount + ','
                        localRevenueTotal = localRevenueTotal + float(amount)
                    elif i['category'] == 'State Program Revenue':
                        stateprogramRevenue = stateprogramRevenue + '"' + str(i['obj']) + '": ' +  amount + ','
                        stateprogramRevenueTotal = stateprogramRevenueTotal + float(amount)
                    elif i['category'] == 'Federal Program Revenue':
                        federalprogramRevenue = federalprogramRevenue + '"' + str(i['obj']) + '": '  +  amount + ','
                        federalprogramRevenueTotal = federalprogramRevenueTotal + float(amount)
            for i in dataExpensebyObject:
                if i['total_activities' + str(x)] != '':
                    amount = i['total_activities' + str(x)].replace(',','')
                    if amount.find('(') != -1:
                        amount = '-' + amount.replace('(','').replace(')','')
                    objCodes = int(i['obj'])
                    if objCodes >= 6100 and objCodes < 6200:
                        payrollandBenefits = payrollandBenefits + '"' + str(i['obj']) + '": ' +  amount + ','
                        payrollandBenefitsTotal = payrollandBenefitsTotal + float(amount)
                    elif objCodes >= 6200 and objCodes < 6300:
                        professionalandcontractServices = professionalandcontractServices + '"' + str(i['obj']) + '": ' +  amount + ','
                        professionalandcontractServicesTotal = professionalandcontractServicesTotal + float(amount)
                    elif objCodes >= 6300 and objCodes < 6400:
                        materialsandSupplies = materialsandSupplies + '"' + str(i['obj']) + '": ' +  amount + ','
                        materialsandSuppliesTotal = materialsandSuppliesTotal + float(amount)
                    elif objCodes >= 6400 and objCodes < 6500:
                        otheroperatingCosts = otheroperatingCosts + '"' + str(i['obj']) + '": ' +  amount + ','
                        otheroperatingCostsTotal = otheroperatingCostsTotal + float(amount)
                    elif objCodes == 6449 :
                        depreciation = depreciation + '"' + str(i['obj']) + '": ' +  amount + ','
                        depreciationTotal = depreciationTotal + float(amount)
                    elif objCodes >= 6500 and objCodes < 6600:                                    
                        debtServices = debtServices + '"' + str(i['obj']) + '": ' +  amount + ','
                        debtServicesTotal = debtServicesTotal + float(amount)
            for i in dataExpense:                
                if i['total_func' + str(x)] != '':
                    amount = i['total_func' + str(x)].replace(',','')
                    if amount.find('(') != -1:
                        amount = '-' + amount.replace('(','').replace(')','')
                    expense = expense + '"' + str(i['func_func']) + ' - ' + str(i['desc']) + '": ' +  amount + ','
                    expenseTotal = expenseTotal + float(amount)

            year = yPath
            if year == '':
                date_object = datetime.today()
                year = str(date_object.strftime("%Y"))
                
            dateStore = year + "-{:02}".format(x)
            startMonth = 9
            if school in schoolMonths['julySchool']:
                startMonth = 7
            if startMonth <= x:
                yearstore = int(year) - 1
                dateStore = str(yearstore) + "-{:02}".format(x)
            
            if localRevenueTotal == 0 and stateprogramRevenueTotal == 0 and federalprogramRevenueTotal == 0 and expenseTotal == 0:
                pass
            else:
                jsonString = localRevenue + '"total": ' + str(int(localRevenueTotal)) + '},' \
                            + stateprogramRevenue + '"total": ' + str(int(stateprogramRevenueTotal)) + '},' \
                            + federalprogramRevenue + '"total": ' + str(int(federalprogramRevenueTotal)) + '},' \
                            + expense + '"total": ' + str(int(expenseTotal)) + '},' \
                            + payrollandBenefits + '"total": ' + str(int(payrollandBenefitsTotal)) + '},' \
                            + professionalandcontractServices + '"total": ' + str(int(professionalandcontractServicesTotal)) + '},' \
                            + materialsandSupplies + '"total": ' + str(int(materialsandSuppliesTotal)) + '},' \
                            + otheroperatingCosts + '"total": ' + str(int(otheroperatingCostsTotal)) + '},' \
                            + depreciation + '"total": ' + str(int(depreciationTotal)) + '},' \
                            + debtServices + '"total": ' + str(int(debtServicesTotal)) + '}}'
                if dateStore in dataDate:
                    queryStatement = "delete from dbo.PLData where date = '" + dateStore + "' and school = '" + school + "'"
                    cursor.execute(queryStatement)
                    cnxn.commit()
                                        
                cursor.execute(insertqueryStatement, dateStore, jsonString, school)
                cnxn.commit()

def writeCodes(school, table, year):    
    
    cnxn = connect()
    cursor = cnxn.cursor()

    my_dict = {}
    my_func = {}
    my_activities = {}
    dataOBJ = {}

    cursor.execute(f"SELECT  * FROM dbo.PL_Definition_obj where school = '" + school + "'")
    rows = cursor.fetchall()
    for row in rows:
        my_dict[row[0]+ '-' + row[1]] = row[0] + ';' + row[1] + ';' + row[2] + ';' + row[3] + ';0;' + row[5]
    
    cursor.execute(f"SELECT  * FROM dbo.PL_Definition_func where school = '" + school + "'")
    rows = cursor.fetchall()
    for row in rows:
        my_func[row[0] + row[3]] = row[0] + ';' + row[1] + ';' + row[2] + ';' + row[3] + ';0;' + row[5] 

    cursor.execute(f"SELECT  * FROM dbo.PL_Activities where school = '" + school + "'")
    rows = cursor.fetchall()
    for row in rows:
        my_activities[row[0]] = row[1]

    cursor.execute(f"SELECT  * FROM dbo.ActivityBS where school = '" + school + "'")
    rows = cursor.fetchall()

    dataBS = {}
    for row in rows:
        dataBS[row[1]] = row[0] + ';' + row[2]

    cursor.execute(f"SELECT  * FROM dbo.List_Activities;")
    rows = cursor.fetchall()
   
    dataFUNC = {}

    for row in rows:
        if row[1] != '':
            dataFUNC[row[1]] = row[3]
        elif row[0] != '' and row[2] != '':
            dataOBJ[row[0] + '-' + row[2]] = row[3]
        elif row[2] != '':
            if not row[2] in dataBS:
                dataBS[row[2]] = row[3]

    cursor.execute("delete from dbo.PL_Definition_func where school = '" + school + "'")
    cnxn.commit()
    
    plOBJqueryStatement = 'INSERT INTO dbo.PL_Definition_obj (fund,obj,Description,Category,budget,school) VALUES (?,?,?,?,?,?)'
    plFUNCqueryStatement = 'INSERT INTO dbo.PL_Definition_func (func,obj,Description,Category,budget,school) VALUES (?,?,?,?,?,?)'
    plACTqueryStatement = 'INSERT INTO dbo.PL_Activities (obj,Description,Category,school) VALUES (?,?,?,?)'
    plBSqueryStatement = 'INSERT INTO dbo.ActivityBS (Activity,obj,Description,school) VALUES (?,?,?,?)'

    selectQuery = f"SELECT * FROM " + table + " where Date >= '2023-07-01' and Type != 'EN'"
    if school in schoolCategory['skyward']:
        selectQuery = f"SELECT fund,func,obj,sobj,org,fscl_yr,PI,LOC,PostingDate,TransactionDescr,Month,Source,Subsource,Batch,Vendor,InvoiceDate,CheckNumber,CheckDate,Amount,budgetOrigin FROM " + table + " where PostingDate >= '2023-09-01' and source != 'RE'"
    cursor.execute(selectQuery)
    rows = cursor.fetchall()

    storeREV = {}
    storeACT = {}
    storeBS = {}
    
    for line in rows:        
        try:
            objCodes = int(line[2])                    
            if objCodes >= 5700 and objCodes < 6000:
                key = line[0] + '-' + line[2]
                if key not in my_dict:
                    if key not in storeREV:
                        value = ''
                        description = line[9][0:49]
                        if description == '':
                            if key in dataOBJ:
                                description = dataOBJ[key]                  
                        if objCodes >= 5700 and objCodes < 5800:
                            value = line[0] + ';' + line[2] + ';' + description + ';Local Revenue;0;' + school
                        elif objCodes >= 5800 and objCodes < 5900:
                            value = line[0] + ';' + line[2] + ';' + description + ';State Program Revenue;0;' + school
                        elif objCodes >= 5900 and objCodes < 6000:
                            value = line[0] + ';' + line[2] + ';' + description + ';Federal Program Revenue;0;' + school
                        if value != '':
                            storeREV[key] = value                
            elif objCodes >= 6100 and objCodes != 6449:
                key = line[2]
                if key not in my_activities:
                    if key not in storeACT:
                        value = ''
                        description = line[9][0:49]                    
                        if objCodes >= 6100 and objCodes < 6200:
                            value = line[2] + ';' + description + ';Payroll and Benefits;' + school
                        elif objCodes >= 6200 and objCodes < 6300:
                            value = line[2] + ';' + description + ';Professional and Contract Services;' + school
                        elif objCodes >= 6300 and objCodes < 6400:
                            value = line[2] + ';' + description + ';Materials and Supplies;' + school
                        elif objCodes >= 6400 and objCodes < 6500:
                            value = line[2] + ';' + description + ';Other Operating Costs;' + school
                        elif objCodes >= 6500 and objCodes < 6600:
                            value = line[2] + ';' + description + ';Debt Services;' + school
                        if value != '':
                            storeACT[key] = value
            elif objCodes < 4500:
                if school in schoolCategory['ascender']:                                                   
                    if line[14] == '0.00' and line[17] == '0.00' and line[18] == '0.00' :
                        pass
                    else:
                        if line[2] not in dataBS:
                            if line[2] not in storeBS:
                                activityType = assignedType(objCodes)         
                                storeBS[line[2]] = activityType[1] + ';' + line[2] + ';' + activityType[0]
                        
                else:
                    if line[18] != '0.00':
                        if line[2] not in dataBS:
                            if line[2] not in storeBS:
                                activityType = assignedType(objCodes)         
                                storeBS[line[2]] = activityType[1] + ';' + line[2] + ';' + activityType[0]

            func = line[1].replace('=', '').replace('"', '')
            if func != '00':
                ctgry = ''
                if line[2] == '6449':
                    ctgry = 'Depreciation and Amortization'

                key = func + ctgry
                if key not in my_func:
                    description = line[9]
                    if key in dataFUNC:
                        description = str(dataFUNC[key])
                    description = description[0:49]
                    value = func + ';' + line[2] + ';' + description + ';' + ctgry + ';0;' + school
                    my_func[key] = value

        except Exception as error:
            pass

    for plCodes in storeREV:        
        splitCodes = str(storeREV[plCodes]).split(';')        
        cursor.execute(plOBJqueryStatement,splitCodes[0], splitCodes[1], splitCodes[2], splitCodes[3], int(splitCodes[4]), splitCodes[5])
    for plCodes in my_func:
        splitCodes = str(my_func[plCodes]).split(';')
        cursor.execute(plFUNCqueryStatement,splitCodes[0], splitCodes[1], splitCodes[2], splitCodes[3], int(splitCodes[4]), splitCodes[5])

    for plCodes in storeACT:
        splitCodes = str(storeACT[plCodes]).split(';')
        cursor.execute(plACTqueryStatement,splitCodes[0], splitCodes[1], splitCodes[2], splitCodes[3])

    for plCodes in storeBS:
        splitCodes = str(storeBS[plCodes]).split(';')
        cursor.execute(plBSqueryStatement,splitCodes[0], splitCodes[1], splitCodes[2], school)

    cnxn.commit()

def assignedType(objectCode):
    if objectCode < 1200:
        return ['Cash and Cash Equivalent', 'Cash' ]
    elif objectCode == 1210:
        return ['Property Taxes To Be Passed Through School Districts', 'DFS+F']
    elif objectCode == 1220:
        return ['Contributions Receivable', 'DFS+F']
    elif objectCode == 1220:
        return ['Allowance for Uncollected Receivables (credit)', 'DFS+F']
    elif objectCode < 1242:
        return ['Receivables', 'DFS+F']
    elif objectCode < 1251:
        return ['Due from other Governments', 'DFS+F']
    elif objectCode < 1294:
        return ['General Fund', 'OTHR']
    elif objectCode < 1311:
        return ['Inventories- Supplies and Materials', 'Inventory']
    elif objectCode < 1411:
        return ['Deferred Expenses', 'OTHR']
    elif objectCode < 1421:
        return ['Pre-paid Workers Comp', 'Acc-Exp']
    elif objectCode < 1490:
        return ['Capitalized Bond Costs', 'LTD']
    elif objectCode < 1500:
        return ['Other Current Assets', 'OTHR']
    elif objectCode < 1520:
        return ['Land Purchase and Improvements', 'FA-L']
    elif objectCode < 1530:
        return ['Buildings and Improvements', 'FA-BFE']
    elif objectCode < 1538:
        return ['Vehicles', 'FA-L']
    elif objectCode == 1538:
        return ['Equipment- Technology', 'FA-BFE']
    elif objectCode < 1550:
        return ['Furniture and Equipment', 'FA-BFE']
    elif objectCode < 1581:
        return ['Capital Leases- Equipment', 'FA-BFE']
    elif objectCode < 1590:
        return ['Accumulated Depreciation', 'FA-AD']
    elif objectCode < 1699:
        return ['Infrastructure Assets', 'OCA']
    elif objectCode < 1799:
        return ['Deferred Outflows of Resources', 'DR']
    elif objectCode < 1828:
        return ['Restricted Cash', 'Restr']
    elif objectCode < 2110:
        return ['Long-term Investments', 'OTHR']
    elif objectCode < 2120:
        return ['Accounts Payable', 'AP']
    elif objectCode < 2124:
        return ['Bonds and Loans Payable', 'OtherLiab']
    elif objectCode < 2131:
        return ['Capital Leases Payable- Current Year', 'Debt-C']
    elif objectCode < 2151:
        return ['Loan Interest Payable', 'ACC-Int']
    elif objectCode < 2162:
        return ['Payroll Deductions and Withholdings', 'Acc-Exp']
    elif objectCode < 2180:
        return ['Special Revenue Fund', 'OTHR']
    elif objectCode < 2190:
        return ['Due to Other Governments', 'AP']
    elif objectCode < 2220:
        return ['ACCRUED PAYROLL EXPENSES', 'Acc-Exp']
    elif objectCode < 2300:
        return ['Accrued Expenditures or Expenses', 'Acc-Exp']
    elif objectCode < 2400:
        return ['Deferred Revenue', 'Debt-D']
    elif objectCode < 2500:
        return ['Payable from Restricted Assets', 'ACC-Int']
    elif objectCode < 3000:
        return ['Bonds and Loans Payable- Long Term', 'LTD']
    elif objectCode >= 3000:
        return ['Net Assets', 'Equity']
    return ''
    
def updateDescription(table, school):

    cnxn = connect()
    cursor = cnxn.cursor()

    my_obj = {}
    my_fund = {}

    cursor.execute("SELECT distinct obj, AcctDescr FROM " + table + " where AcctDescr != ''")
    rows = cursor.fetchall()

    for row in rows:
        my_obj[row[0]] = row[1]

    cursor.execute("SELECT distinct fund, obj, AcctDescr FROM " + table + " where AcctDescr != ''")
    rows = cursor.fetchall()

    for row in rows:
        my_fund[row[0] + '-' + row[1]] = row[2]
                
    cursor.execute(f"SELECT  * FROM [dbo].[ActivityBS] where school = '" + school + "'")
    rows = cursor.fetchall()

    for row in rows:
        try:
            queryStatement = "update [dbo].[ActivityBS] set Description = '" + my_obj[row[1]] + "' where obj = '" + row[1] + "' and school = '" + school + "'"
            cursor.execute(queryStatement)
        except:
            pass
        
    cursor.execute(f"SELECT  * FROM [dbo].[PL_Activities] where school = '" + school + "'")
    rows = cursor.fetchall()

    for row in rows:
        try:
            queryStatement = "update [dbo].[PL_Activities] set Description = '" + my_obj[row[0]] + "' where obj = '" + row[0] + "' and school = '" + school + "'"
            cursor.execute(queryStatement)
        except:
            pass

    cursor.execute(f"SELECT  * FROM [dbo].[PL_Definition_obj] where school = '" + school + "'")
    rows = cursor.fetchall()

    for row in rows:
        try:
            queryStatement = "update [dbo].[PL_Definition_obj] set Description = '" + my_fund[row[0] + '-' + row[1]] + "' where fund = '" + row[0] + "' and obj = '" + row[1] + "' and school = '" + school + "'"
            cursor.execute(queryStatement)
        except:
            pass

    cnxn.commit()

def balance_sheet_asc(school,year):
    print("balance_sheet_asc")
    present_date = datetime.today().date()   
    present_year = present_date.year
    
    today_date = datetime.now()
    
    today_month = today_date.month

    if year:
        year = int(year)
        if year == present_year:
            
            print("year",year)

            if school in schoolMonths["septemberSchool"]:
                if today_month <= 8:
                    
                    start_year = year - 1
                    present_year = present_year - 1
                    FY_year_current = year - 1
                else: 
                    start_year = year 
                    FY_year_current = year
            else:
                if today_month <= 6:
                    start_year = year - 1
                    present_year = present_year - 1
                    FY_year_current = year - 1
                else: 
                    start_year = year 
                    FY_year_current = year
        else:
            start_year = year 
            FY_year_current = year
        
        if school in schoolMonths["julySchool"]:
            current_date = datetime(start_year, 7, 1).date()
            
        else:
            current_date = datetime(start_year, 9, 1).date() 
        current_year = current_date.year
    else:
        start_year = 2021
        current_date = datetime.today().date()   
        current_year = current_date.year
        FY_year_current = current_year

    while start_year <= FY_year_current:
        print(start_year)
        FY_year_1 = start_year
        FY_year_2 = start_year + 1 
        start_year = FY_year_2

        cnxn = connect()
        cursor = cnxn.cursor()
        cursor.execute(f"SELECT  * FROM [dbo].{db[school]['bs']} AS T1 LEFT JOIN [dbo].{db[school]['bs_fye']} AS T2 ON T1.BS_id = T2.BS_id ;  ")
        rows = cursor.fetchall()
        bs_checker = []

        for row in rows:
            if row[8] == school:
                if FY_year_1 == row[9]:
                    row_dict = {
                        "Activity": row[0],
                        "Description": row[1],
                        "Category": row[2],
                        "Subcategory": row[3],
                        "FYE": row[4],
                        "BS_id": row[5],
                        "school": row[8],
                
                    }
                    bs_checker.append(row_dict)
            
        
        if bs_checker == []:
            for i in range(1, 30):
                fye = '0'
                query = "INSERT INTO [dbo].[BS_FYE] (BS_id, FYE, school,year) VALUES (?, ?, ?,?)"
                cursor.execute(query, (i, fye, school,FY_year_1)) 
                cnxn.commit()
                
        
        cursor.execute(f"SELECT  * FROM [dbo].{db[school]['bs']} AS T1 LEFT JOIN [dbo].{db[school]['bs_fye']} AS T2 ON T1.BS_id = T2.BS_id ;  ")
        rows = cursor.fetchall()

        data_balancesheet = []

        for row in rows:
            if row[8] == school:
                fye = float(row[7]) if row[7] else 0
                if fye == 0:
                    fyeformat = ""
                else:
                    if row[0] == 'Cash' or row[0] == 'AP':
                        fyeformat = (
                            "${:,.0f}".format(abs(fye)) if fye >= 0 else "$({:,.0f})".format(abs(fye))
                        )
                    else:
                        fyeformat = (
                            "{:,.0f}".format(abs(fye)) if fye >= 0 else "({:,.0f})".format(abs(fye))
                        )
                if FY_year_1 == row[9]:
                    row_dict = {
                        "Activity": row[0],
                        "Description": row[1],
                        "Category": row[2],
                        "Subcategory": row[3],
                        "FYE": fyeformat, #should now be total fye coming from GL(data3)
                        "BS_id": row[5], #wont be used
                        "school": row[8],
                

                    }

                    data_balancesheet.append(row_dict)

        if FY_year_1 == present_year:
            relative_path = os.path.join("profit-loss", school)
        else:
            relative_path = os.path.join(str(FY_year_1), "profit-loss", school)

        json_path = os.path.join(JSON_DIR, relative_path)
        with open(os.path.join(json_path, "data.json"), "r") as f:
            data = json.load(f)

        with open(os.path.join(json_path, "data2.json"), "r") as f:
            data2 = json.load(f)


        cursor.execute(f"SELECT * FROM [dbo].{db[school]['bs_activity']}")
        rows = cursor.fetchall()

        data_activitybs = []

        for row in rows:
            if row[3] == school:
                row_dict = {
                    "Activity": row[0],
                    "obj": row[1],
                    "Description2": row[2],
                    "school": row[3],
                }
        
                data_activitybs.append(row_dict)

        with open(os.path.join(json_path, "data3.json"), "r") as f:
            data3 = json.load(f)
       

        with open(os.path.join(json_path, "totals.json"), "r") as f:
            totals = json.load(f)

        with open(os.path.join(json_path, "months.json"), "r") as f:
            months = json.load(f)

        last_month = months["last_month"]
        last_month_number = months["last_month_number"]
        last_month_name = months["last_month_name"]
   
        db_last_month = months["db_last_month"]
        month_exception = months["month_exception"]
        month_exception_str = months["month_exception_str"]
        

     
        cursor.execute(f"SELECT * FROM [dbo].{db[school]['adjustment']} ")
        rows = cursor.fetchall()

        adjustment = []

        if school in schoolCategory["ascender"]:
            for row in rows:
                expend = float(row[17])
                row_dict = {
                    "fund": row[0],
                    "func": row[1],
                    "obj": row[2],
                    "sobj": row[3],
                    "org": row[4],
                    "fscl_yr": row[5],
                    "pgm": row[6],
                    "edSpan": row[7],
                    "projDtl": row[8],
                    "AcctDescr": row[9],
                    "Number": row[10],
                    "Date": row[11],
                    "AcctPer": row[12],
                    "Est": row[13],
                    "Real": row[14],
                    "Appr": row[15],
                    "Encum": row[16],
                    "Expend": expend,
                    "Bal": row[18],
                    "WorkDescr": row[19],
                    "Type": row[20],
                    "School": row[21],
                }
                adjustment.append(row_dict)

        acct_per_values = [
            "01",
            "02",
            "03",
            "04",
            "05",
            "06",
            "07",
            "08",
            "09",
            "10",
            "11",
            "12",
        ]

        real_key = "Real"        
        bal_key = "Bal"
        expend_key = "Expend"
        begbal_key = "BegBal"
        if school in schoolCategory["skyward"]:
            bal_key = "Amount"
            real_key = "Amount"
            expend_key = "Amount"
            begbal_key = "BegBal"




        school_fye = ['aca','advantage','cumberland','pro-vision','manara','stmary','sa']

        unique_act = []
        for item in data_balancesheet:
            Activity = item["Activity"]

            if item['Subcategory'] == 'Long Term Debt' or  item['Subcategory'] == 'Current Liabilities' or item['Category'] == 'Net Assets':
                if Activity not in unique_act:
                    unique_act.append(Activity)

        if school in schoolCategory["ascender"]:
            numberstack = set()

            for item in data3:
                if item["obj"] == '3600':
                    
                    numberstack.add(item["Number"])
                    print("Number",item["Number"])
            
            numberstack = list(numberstack)
            



                    

        for item in data_activitybs:
            Activity = item["Activity"]
            obj = item["obj"]
            item["fytd"] = 0
            
            for i, acct_per in enumerate(acct_per_values, start=1):
                if school in schoolCategory["ascender"]:
                    total_data3 = sum(
                        entry[bal_key]
                        for entry in data3
                        if entry["obj"] == obj 
                        and entry["AcctPer"] == acct_per
                        and entry["fund"] != '000'
                        and "BEG BAL" not in entry["WorkDescr"]
                        and "BEGBAL" not in entry["WorkDescr"]
                        and "BEGINNING BAL" not in entry["WorkDescr"]
                    )
                else:
                    total_data3 = sum(
                        entry[bal_key]
                        for entry in data3
                        if entry["obj"] == obj 
                        and entry["AcctPer"] == acct_per
                        and entry["fund"] != '000'

                    )
                total_adjustment = sum(
                    entry[bal_key]
                    for entry in adjustment
                    if entry["obj"] == obj
                    and entry["AcctPer"] == acct_per 
                    and entry["School"] == school
                    and entry[bal_key] is not None 
                    and not isinstance(entry[bal_key], str)
                )
             
                item[f"total_bal{i}"] = total_data3 + total_adjustment
                if i != month_exception:
                    item["fytd"] += item[f"total_bal{i}"]

            


            if school in schoolCategory["ascender"]:

                activity_fye = sum(
                    entry[bal_key]
                    for entry in data3
                    if entry["obj"] == obj
                    and entry["Type"] == "GJ"
                    and entry["Number"] in numberstack
                )
                print(activity_fye)
                if Activity in unique_act:
                    item["activity_fye"] = -(activity_fye)
                else:
                    item["activity_fye"] = activity_fye
             

        if school in schoolCategory["ascender"]:
            for item in data_balancesheet:
                Activity = item["Activity"]
                item["FYE"] = sum(
                    entry["activity_fye"]
                    for entry in data_activitybs
                    if entry["Activity"] == Activity
                )
                print(Activity,item["FYE"])


                
            
            
        activity_sum_dict = {}


        for item in data_activitybs:
            Activity = item["Activity"]
            
            for i in range(1, 13):
                total_sum_i = sum(
                    float(entry[f"total_bal{i}"])
                    if entry[f"total_bal{i}"] and entry["Activity"] == Activity
                    else 0
                    for entry in data_activitybs
                )
                activity_sum_dict[(Activity, i)] = total_sum_i
            
        for row in data_balancesheet:
            activity = row["Activity"]
            
            for i in range(1, 13):
                key = (activity, i)
                row[f"total_sum{i}"] = (activity_sum_dict.get(key, 0))


      
      



        # TOTAL REVENUE
        total_revenue = {acct_per: 0 for acct_per in acct_per_values}
        for item in data:
            fund = item["fund"]
            obj = item["obj"]

            for i, acct_per in enumerate(acct_per_values, start=1):
                total_real = sum(
                    entry[real_key]
                    for entry in data3
                    if entry["fund"] == fund
                    and entry["obj"] == obj
                    and entry["AcctPer"] == acct_per
                 
                )
                total_adjustment = sum(
                        entry[real_key]
                        for entry in adjustment
                        if entry["fund"] == fund
                        and entry["AcctPer"] == acct_per
                        and entry["obj"] == obj
                        and entry["School"] == school
                        and entry[real_key] is not None 
                        and not isinstance(entry[real_key], str)
                    )
                item[f"total_real{i}"] = total_real + total_adjustment

                
                total_revenue[acct_per] += (item[f"total_real{i}"])

                


        # total surplus / first total
        total_surplus = {acct_per: 0 for acct_per in acct_per_values}

        for item in data2:
            if item["category"] != "Depreciation and Amortization":
                func = item["func_func"]

                for i, acct_per in enumerate(acct_per_values, start=1):
                    total_func = sum(
                        entry[expend_key]
                        for entry in data3
                        if entry["func"] == func and entry["AcctPer"] == acct_per and entry["obj"] != '6449'
                    )
                    total_adjustment = sum(
                        entry[expend_key]
                        for entry in adjustment
                        if entry["func"] == func 
                        and entry["AcctPer"] == acct_per 
                        and entry["obj"] != '6449' 
                        and entry["School"] == school
                        and entry[expend_key] is not None 
                        and not isinstance(entry[expend_key], str)
                    )
                    item[f"total_func{i}"] = total_func + total_adjustment
                    total_surplus[acct_per] += item[f"total_func{i}"]



        # ---- Depreciation and ammortization total
        total_DnA = {acct_per: 0 for acct_per in acct_per_values}

        for item in data2:
            if item["category"] == "Depreciation and Amortization":
                func = item["func_func"]
                obj = item["obj"]

                for i, acct_per in enumerate(acct_per_values, start=1):
                    total_func = sum(
                        entry[expend_key]
                        for entry in data3
                        if entry["func"] == func
                        and entry["AcctPer"] == acct_per
                        and entry["obj"] == obj
                    )
                    total_adjustment = sum(
                        entry[expend_key]
                        for entry in adjustment
                        if entry["func"] == func
                        and entry["AcctPer"] == acct_per
                        and entry["obj"] == obj
                        and entry["School"] == school
                        and entry[expend_key] is not None 
                        and not isinstance(entry[expend_key], str)
                        

                    )
                    item[f"total_func2_{i}"] = total_func + total_adjustment
                    total_DnA[acct_per] += item[f"total_func2_{i}"]

        total_SBD = {
            acct_per: abs(total_revenue[acct_per]) - total_surplus[acct_per]
            for acct_per in acct_per_values
        }
        total_netsurplus = {
            acct_per: total_SBD[acct_per] - total_DnA[acct_per] #dna_total_months in pl.. SBD same as pl
            for acct_per in acct_per_values
        }
    
        ytd_DnA = sum(total_DnA.values())
        ytd_netsurplus = sum(total_netsurplus.values())

        # for month, total in monthly_totals_func2.items():
        #     print(f'MonthFUNC2 {month}: {total}')

        # for key, value in difference_func_values.items():
        #     print(f'{key}: {value}')

        def format_with_parentheses(value):
            if value >= 1:
                return "${:,.0f}".format(round(value))
            elif value <= -1:
                return "$({:,.0f})".format(abs(round(value)))
            else:
                return ""

        def format_with_parentheses2(value):
            if value == 0:
                return ""
            formatted_value = "{:,.0f}".format(abs(round(value)))
            return "({})".format(formatted_value) if value > 0 else formatted_value

        def format_value_dollars(value):
            if value > 0:
                return "${:,.0f}".format(round(value))
            elif value < 0:
                return "$({:,.0f})".format(abs(round(value)))
            else:
                return ""
        def format_value(value):
            if value > 0:
                return "{:,.0f}".format(round(value))
            elif value < 0:
                return "({:,.0f})".format(abs(round(value)))
            else:
                return ""

        def format_negative(value):
            if value > 0:
                return "({:,.0f})".format(round(value))
            elif value < 0:
                return "{:,.0f}".format(abs(round(value)))
            else:
                return ""

        for row in data_balancesheet:
            if row["school"] == school:    
                if school in schoolCategory["ascender"]:
                    FYE_value = float(row["FYE"])
                else:
                    FYE_value = (float(row["FYE"].replace("$","").replace(",", "").replace("(", "-").replace(")", ""))
                        if row["FYE"]
                        else 0
                    )
                total_sum9_value = float(row["total_sum9"])
                total_sum10_value = float(row["total_sum10"])
                total_sum11_value = float(row["total_sum11"])
                total_sum12_value = float(row["total_sum12"])
                total_sum1_value = float(row["total_sum1"])
                total_sum2_value = float(row["total_sum2"])
                total_sum3_value = float(row["total_sum3"])
                total_sum4_value = float(row["total_sum4"])
                total_sum5_value = float(row["total_sum5"])
                total_sum6_value = float(row["total_sum6"])
                total_sum7_value = float(row["total_sum7"])
                total_sum8_value = float(row["total_sum8"])

                total_sums = [
                                float(row[f"total_sum{i}"]) for i in range(1, 13)
                            ]
                if school in schoolMonths['septemberSchool']:
                
                    # Calculate the differences and store them in the row dictionary
                    row["difference_9"] = (FYE_value + total_sum9_value)
                    row["difference_10"] =(row["difference_9"] + total_sum10_value)
                    row["difference_11"] =(row["difference_10"] + total_sum11_value)
                    row["difference_12"] =(row["difference_11"]  + total_sum12_value )
                    row["difference_1"] = (row["difference_12"] + total_sum1_value )
                    row["difference_2"] = (row["difference_1"] + total_sum2_value )
                    row["difference_3"] = (row["difference_2"] + total_sum3_value )
                    row["difference_4"] = (row["difference_3"] + total_sum4_value )
                    row["difference_5"] = (row["difference_4"] + total_sum5_value )
                    row["difference_6"] = (row["difference_5"] + total_sum6_value )
                    row["difference_7"] = (row["difference_6"] + total_sum7_value )
                    row["difference_8"] = (row["difference_7"] + total_sum8_value )
    
                
                    row["last_month_difference"] = row[f"difference_{last_month_number}"] 
              

                    if month_exception != "":
                        
                        row["fytd"] = sum(
                            total_sum
                            for i, total_sum in enumerate(total_sums, start=1)
                            if i != month_exception
                        )
                    else:
                        row["fytd"] =sum(total_sums)

                    row["debt_9"]  = (FYE_value - total_sum9_value)
                    row["debt_10"] = (row["debt_9"] - total_sum10_value)
                    row["debt_11"] = (row["debt_10"] - total_sum11_value)
                    row["debt_12"] = (row["debt_11"] - total_sum12_value)
                    row["debt_1"] = (row["debt_12"] - total_sum1_value)
                    row["debt_2"] = (row["debt_1"] - total_sum2_value)
                    row["debt_3"] = (row["debt_2"] - total_sum3_value)
                    row["debt_4"] = (row["debt_3"]- total_sum4_value)
                    row["debt_5"] = (row["debt_4"]  - total_sum5_value)
                    row["debt_6"] = (row["debt_5"]- total_sum6_value)
                    row["debt_7"] = (row["debt_6"] - total_sum7_value)
                    row["debt_8"] = (row["debt_7"] - total_sum8_value)
                    row["last_month_debt"] = row[f"debt_{last_month_number}"] 



                    if month_exception != "":
                        
                        row["debt_fytd"] = -sum(
                            total_sum
                            for i, total_sum in enumerate(total_sums, start=1)
                            if i != month_exception
                        )
                    else:
                        row["debt_fytd"] =-sum(total_sums)
                    

                    row["net_assets9"] = (FYE_value + total_netsurplus["09"])
                    row["net_assets10"] = (row["net_assets9"] + total_netsurplus["10"])
                    row["net_assets11"] = (row["net_assets10"]+ total_netsurplus["11"])
                    row["net_assets12"] = (row["net_assets11"]+ total_netsurplus["12"])
                    row["net_assets1"] = (row["net_assets12"] + total_netsurplus["01"])
                    row["net_assets2"] = (row["net_assets1"] + total_netsurplus["02"])
                    row["net_assets3"] = (row["net_assets2"]+ total_netsurplus["03"])
                    row["net_assets4"] = (row["net_assets3"] + total_netsurplus["04"])
                    row["net_assets5"] = (row["net_assets4"] + total_netsurplus["05"])
                    row["net_assets6"] = (row["net_assets5"]  + total_netsurplus["06"])
                    row["net_assets7"] = (row["net_assets6"] + total_netsurplus["07"])
                    row["net_assets8"] = (row["net_assets7"] + total_netsurplus["08"])
                    row["last_month_net_assets"] = row[f"net_assets{last_month_number}"]
                    
                else:
                                    # Calculate the differences and store them in the row dictionary
                    row["difference_7"] = (FYE_value + total_sum7_value )
            
                    row["difference_8"] = (row["difference_7"] + total_sum8_value )
                    row["difference_9"] = (row["difference_8"]  + total_sum9_value)
                    row["difference_10"] =(row["difference_9"] + total_sum10_value)
                    row["difference_11"] =(row["difference_10"] + total_sum11_value)
                    row["difference_12"] =(row["difference_11"]  + total_sum12_value )
                    row["difference_1"] = (row["difference_12"] + total_sum1_value )
                    row["difference_2"] = (row["difference_1"] + total_sum2_value )
                    row["difference_3"] = (row["difference_2"] + total_sum3_value )
                    row["difference_4"] = (row["difference_3"] + total_sum4_value )
                    row["difference_5"] = (row["difference_4"] + total_sum5_value )
                    row["difference_6"] = (row["difference_5"] + total_sum6_value )
                    
                    row["last_month_difference"] = row[f"difference_{last_month_number}"] 
                    

                    if month_exception != "":
                        
                        row["fytd"] = sum(
                            total_sum
                            for i, total_sum in enumerate(total_sums, start=1)
                            if i != month_exception
                        )
                    else:
                        row["fytd"] =sum(total_sums)

                    row["debt_7"] = (FYE_value - total_sum7_value)
                    row["debt_8"] = (row["debt_7"] - total_sum8_value)
                    row["debt_9"]  = (row["debt_8"] - total_sum9_value)
                    row["debt_10"] = (row["debt_9"] - total_sum10_value)
                    row["debt_11"] = (row["debt_10"] - total_sum11_value)
                    row["debt_12"] = (row["debt_11"] - total_sum12_value)
                    row["debt_1"] = (row["debt_12"] - total_sum1_value)
                    row["debt_2"] = (row["debt_1"] - total_sum2_value)
                    row["debt_3"] = (row["debt_2"] - total_sum3_value)
                    row["debt_4"] = (row["debt_3"]- total_sum4_value)
                    row["debt_5"] = (row["debt_4"]  - total_sum5_value )
                    row["debt_6"] = (row["debt_5"]- total_sum6_value)
                    row["last_month_debt"] = row[f"debt_{last_month_number}"] 
    
                    
                    if month_exception != "":
                        
                        row["debt_fytd"] = -sum(
                            total_sum
                            for i, total_sum in enumerate(total_sums, start=1)
                            if i != month_exception
                        )
                    else:
                        row["debt_fytd"] =-sum(total_sums)

                    row["net_assets7"] = (FYE_value + total_netsurplus["07"])
                    row["net_assets8"] = (row["net_assets7"] + total_netsurplus["08"])
                    row["net_assets9"] = (row["net_assets8"]  + total_netsurplus["09"])
                    row["net_assets10"] = (row["net_assets9"] + total_netsurplus["10"])
                    row["net_assets11"] = (row["net_assets10"]+ total_netsurplus["11"])
                    row["net_assets12"] = (row["net_assets11"]+ total_netsurplus["12"])
                    row["net_assets1"] = (row["net_assets12"] + total_netsurplus["01"])
                    row["net_assets2"] = (row["net_assets1"] + total_netsurplus["02"])
                    row["net_assets3"] = (row["net_assets2"]+ total_netsurplus["03"])
                    row["net_assets4"] = (row["net_assets3"] + total_netsurplus["04"])
                    row["net_assets5"] = (row["net_assets4"] + total_netsurplus["05"])
                    row["net_assets6"] = (row["net_assets5"]  + total_netsurplus["06"])
                    row["last_month_net_assets"] = row[f"net_assets{last_month_number}"]

        total_current_assets = {acct_per: 0 for acct_per in acct_per_values}
        total_current_assets_fye = 0
        total_current_assets_fytd = 0 
        last_month_current_assets = 0 

        total_capital_assets = {acct_per: 0 for acct_per in acct_per_values}
        total_capital_assets_fye = 0
        total_capital_assets_fytd = 0 
        last_month_total_capital_assets = 0

        total_current_liabilities = {acct_per: 0 for acct_per in acct_per_values}
        total_current_liabilities_fye = 0
        total_current_liabilities_fytd = 0
        last_month_total_current_liabilities = 0

        total_liabilities = {acct_per: 0 for acct_per in acct_per_values}
        total_liabilities_fye = 0
        total_liabilities_fytd = 0
        last_month_total_liabilities = 0

        total_assets = {acct_per: 0 for acct_per in acct_per_values}
        total_assets_fye = 0
        total_assets_fye_fytd = 0
        last_month_total_assets = 0 
        

        total_LNA = {acct_per: 0 for acct_per in acct_per_values} # LIABILITES AND NET ASSETS 
        total_LNA_fye = 0
        total_LNA_fytd = 0
        total_net_assets_fytd = 0
        last_month_total_LNA = 0

        total_net_assets_fytd = totals["bs_ytd_netsurplus"]    #assign the value coming from profitloss totals
        
    

        for row in data_balancesheet:
            if row["school"] == school:
                subcategory =  row["Subcategory"]

                if school in schoolCategory["ascender"]:
                    fye = float(row["FYE"])
                else:
                    fye =  float(row["FYE"].replace("$","").replace(",", "").replace("(", "-").replace(")", "")) if row["FYE"] else 0


                if subcategory == 'Current Assets':
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_current_assets[acct_per] += row[f"difference_{i}"]
                        if i == last_month_number:
                            last_month_current_assets += row[f"difference_{i}"]                      
                    total_current_assets_fytd += row["fytd"]

                    total_current_assets_fye +=  fye
                if subcategory == 'Capital Assets, Net':
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_capital_assets[acct_per] += row[f"difference_{i}"]
                        if i == last_month_number:
                            last_month_total_capital_assets += row[f"difference_{i}"]
                    
                            

                    total_capital_assets_fytd += row["fytd"]
                    total_capital_assets_fye +=  fye

                    
                if subcategory == 'Current Liabilities':
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_current_liabilities[acct_per] += row[f"debt_{i}"]
                        if i == last_month_number:
                            last_month_total_current_liabilities += row[f"debt_{i}"]
                    total_current_liabilities_fytd += row["debt_fytd"]
                    total_current_liabilities_fye +=  fye

        total_liabilities_fytd_2 = 0
        for row in data_balancesheet:
            if row["school"] == school:
                subcategory =  row["Subcategory"]
                
                if school in schoolCategory["ascender"]:
                    fye = float(row["FYE"])
                else:
                    fye =  float(row["FYE"].replace("$","").replace(",", "").replace("(", "-").replace(")", "")) if row["FYE"] else 0
               
                if subcategory == 'Long Term Debt':
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_liabilities[acct_per] += row[f"debt_{i}"] + total_current_liabilities[acct_per]
                        if i == last_month_number:
                            last_month_total_liabilities += row[f"debt_{i}"] + total_current_liabilities[acct_per]
                    total_liabilities_fytd_2 += row["debt_fytd"]
                    total_liabilities_fye +=   total_current_liabilities_fye + fye
        total_liabilities_fytd = total_liabilities_fytd_2 + total_current_liabilities_fytd

        for row in data_balancesheet:
            if row["school"] == school:
                
                if school in schoolCategory["ascender"]:
                    fye = float(row["FYE"])
                else:
                    fye =  float(row["FYE"].replace("$","").replace(",", "").replace("(", "-").replace(")", "")) if row["FYE"] else 0
                
                if  row["Category"] == "Net Assets":
                    for i, acct_per in enumerate(acct_per_values,start = 1):
                        total_LNA[acct_per] += row[f"net_assets{i}"] + total_liabilities[acct_per]
                        if i == last_month_number:
                            last_month_total_LNA += row[f"net_assets{i}"] + total_liabilities[acct_per]

                    total_LNA_fye += total_liabilities_fye + fye
        
        total_assets = {
            acct_per: total_current_assets[acct_per] + total_capital_assets[acct_per]
            for acct_per in acct_per_values

        }

        last_month_number_str = f"{last_month_number:02}"  
        last_month_total_assets  = total_assets[last_month_number_str]
        
        

        total_assets_fye = total_current_assets_fye + total_capital_assets_fye

        total_assets_fye_fytd = total_current_assets_fytd + total_capital_assets_fytd
        total_LNA_fytd = total_net_assets_fytd + total_liabilities_fytd
        total_net_assets_fytd = format_value(total_net_assets_fytd)
        total_current_assets_fye = format_value(total_current_assets_fye)
        total_capital_assets_fye = format_value(total_capital_assets_fye)
        total_current_liabilities_fye = format_value(total_current_liabilities_fye)
        total_liabilities_fye = format_value(total_liabilities_fye)
        total_assets_fye = format_value_dollars(total_assets_fye)
        total_LNA_fye = format_value_dollars(total_LNA_fye)

        total_current_assets_fytd = format_value(total_current_assets_fytd)
        total_capital_assets_fytd = format_value(total_capital_assets_fytd)
        total_current_liabilities_fytd = format_value(total_current_liabilities_fytd)
        total_liabilities_fytd = format_value(total_liabilities_fytd)
        total_assets_fye_fytd = format_value_dollars(total_assets_fye_fytd)
        total_LNA_fytd = format_value_dollars(total_LNA_fytd)

        total_current_assets = {acct_per: format_value(value) for acct_per, value in total_current_assets.items() if value != 0}
        total_capital_assets = {acct_per: format_value(value) for acct_per, value in total_capital_assets.items() if value != 0}
        total_current_liabilities = {acct_per: format_value(value) for acct_per, value in total_current_liabilities.items() if value != 0}
        total_liabilities = {acct_per: format_value(value) for acct_per, value in total_liabilities.items() if value != 0}
        total_assets = {acct_per: format_value_dollars(value) for acct_per, value in total_assets.items() if value != 0}
        total_LNA = {acct_per: format_value_dollars(value) for acct_per, value in total_LNA.items() if value != 0}

        last_month_current_assets = format_value(last_month_current_assets)
        last_month_total_capital_assets = format_value(last_month_total_capital_assets)
        last_month_total_assets = format_value_dollars(last_month_total_assets)
        last_month_total_current_liabilities = format_value(last_month_total_current_liabilities)
        last_month_total_liabilities = format_value(last_month_total_liabilities)
        last_month_total_LNA = format_value_dollars(last_month_total_LNA)
        
    
        for row in data_balancesheet:
            if row["school"] == school:
                if school in schoolCategory["ascender"]:
                    if row["Activity"] == 'Cash' or row["Activity"] == 'AP':

                        row["FYE"] = format_value_dollars(row["FYE"])
                    else:
                        row["FYE"] = format_value(row["FYE"])

                if row["Activity"] == 'Cash':
                    row["difference_9"] = format_value_dollars(row["difference_9"]) 
                    row["difference_10"]= format_value_dollars(row["difference_10"])
                    row["difference_11"]= format_value_dollars(row["difference_11"])
                    row["difference_12"]= format_value_dollars(row["difference_12"])
                    row["difference_1"] = format_value_dollars(row["difference_1"] )
                    row["difference_2"] = format_value_dollars(row["difference_2"] )
                    row["difference_3"] = format_value_dollars(row["difference_3"] )
                    row["difference_4"] = format_value_dollars(row["difference_4"] )
                    row["difference_5"] = format_value_dollars(row["difference_5"] )
                    row["difference_6"] = format_value_dollars(row["difference_6"] )
                    row["difference_7"] = format_value_dollars(row["difference_7"] )
                    row["difference_8"] = format_value_dollars(row["difference_8"] )
                    
                    row["last_month_difference"] = format_value_dollars(row["last_month_difference"] )
                    row["fytd"] = format_value_dollars(row["fytd"])
                else:
                    
                    row["difference_9"] = format_value(row["difference_9"]) 
                    row["difference_10"]= format_value(row["difference_10"])
                    row["difference_11"]= format_value(row["difference_11"])
                    row["difference_12"]= format_value(row["difference_12"])
                    row["difference_1"] = format_value(row["difference_1"] )
                    row["difference_2"] = format_value(row["difference_2"] )
                    row["difference_3"] = format_value(row["difference_3"] )
                    row["difference_4"] = format_value(row["difference_4"] )
                    row["difference_5"] = format_value(row["difference_5"] )
                    row["difference_6"] = format_value(row["difference_6"] )
                    row["difference_7"] = format_value(row["difference_7"] )
                    row["difference_8"] = format_value(row["difference_8"] )
                    row["last_month_difference"] = format_value(row["last_month_difference"] )
                    row["fytd"] = format_value(row["fytd"])
                
                if row['Activity'] == 'AP':
                    
                    row["debt_9"] =  format_value_dollars(row["debt_9"] )
                    row["debt_10"]=  format_value_dollars(row["debt_10"])
                    row["debt_11"]=  format_value_dollars(row["debt_11"])
                    row["debt_12"]=  format_value_dollars(row["debt_12"])
                    row["debt_1"] =  format_value_dollars(row["debt_1"] )
                    row["debt_2"] =  format_value_dollars(row["debt_2"] )
                    row["debt_3"] =  format_value_dollars(row["debt_3"] )
                    row["debt_4"] =  format_value_dollars(row["debt_4"] )
                    row["debt_5"] =  format_value_dollars(row["debt_5"] )
                    row["debt_6"] =  format_value_dollars(row["debt_6"] )
                    row["debt_7"] =  format_value_dollars(row["debt_7"] )
                    row["debt_8"] =  format_value_dollars(row["debt_8"] )
                    row["debt_fytd"]=format_value_dollars(row["debt_fytd"])
                    row["last_month_debt"] = format_value_dollars(row["last_month_debt"] )

                else:
                    
                    row["debt_9"] =  format_value(row["debt_9"] )
                    row["debt_10"]=  format_value(row["debt_10"])
                    row["debt_11"]=  format_value(row["debt_11"])
                    row["debt_12"]=  format_value(row["debt_12"])
                    row["debt_1"] =  format_value(row["debt_1"] )
                    row["debt_2"] =  format_value(row["debt_2"] )
                    row["debt_3"] =  format_value(row["debt_3"] )
                    row["debt_4"] =  format_value(row["debt_4"] )
                    row["debt_5"] =  format_value(row["debt_5"] )
                    row["debt_6"] =  format_value(row["debt_6"] )
                    row["debt_7"] =  format_value(row["debt_7"] )
                    row["debt_8"] =  format_value(row["debt_8"] )
                    row["last_month_debt"] = format_value(row["last_month_debt"] )
                    row["debt_fytd"]=format_value(row["debt_fytd"])
        

                row["net_assets9"]  = format_value(row["net_assets9"])
                row["net_assets10"] = format_value(row["net_assets10"])
                row["net_assets11"] = format_value(row["net_assets11"])
                row["net_assets12"] = format_value(row["net_assets12"])
                row["net_assets1"]  = format_value(row["net_assets1"])
                row["net_assets2"]  = format_value(row["net_assets2"])
                row["net_assets3"]  = format_value(row["net_assets3"])
                row["net_assets4"]  = format_value(row["net_assets4"])
                row["net_assets5"]  = format_value(row["net_assets5"])
                row["net_assets6"]  = format_value(row["net_assets6"])
                row["net_assets7"]  = format_value(row["net_assets7"])
                row["net_assets8"]  = format_value(row["net_assets8"])
                row["last_month_net_assets"] = format_value(row["last_month_net_assets"])

        keys_to_check = [
            "total_bal1",
            "total_bal2",
            "total_bal3",
            "total_bal4",
            "total_bal5",
            "total_bal6",
            "total_bal7",
            "total_bal8",
            "total_bal9",
            "total_bal10",
            "total_bal11",
            "total_bal12",
            "fytd"
        ]

        threshold = 0.50
        
        for row in data_activitybs:
            for key in keys_to_check:
                value = float(row[key])
                if value == 0:
                    row[key] = ""
                elif value > 0:
                    
                    row[key] = "{:,.0f}".format(abs(float(row[key])))
                elif value != "":
                    row[key] = "({:,.0f})".format(float(row[key]))


        if school in schoolCategory["ascender"]:
            for row in data_activitybs:
                if row['Activity'] == "AP" or row["Activity"] == 'Cash':
                    row["activity_fye"] = format_value_dollars(row["activity_fye"])
                else:
                    row["activity_fye"] = format_value(row["activity_fye"])

        # for row in data_balancesheet:
        #     subcategory = row["Subcategory"]
        #     fye = float(row["FYE"])

        #     row["total_fye"][subcategory] += fye
        #     row["total_fye"] = total_fye[subcategory]

        # for row in data_balancesheet:
        #     row['diffunc9']

        # keys_to_check_func = ['total_func1', 'total_func2', 'total_func3', 'total_func4', 'total_func5','total_func6','total_func7','total_func8','total_func9','total_func10','total_func11','total_func12']
        # keys_to_check_func_2 = ['total_func2_1', 'total_func2_2', 'total_func2_3', 'total_func2_4', 'total_func2_5','total_func2_6','total_func2_7','total_func2_8','total_func2_9','total_func2_10','total_func2_11','total_func2_12']

        # for row in data2:
        #     for key in keys_to_check_func:
        #         if row[key] > 0:
        #             row[key] = row[key]
        #         else:
        #             row[key] = ''
        # for row in data2:
        #     for key in keys_to_check_func:
        #         if row[key] != "":
        #             row[key] = "{:,.0f}".format(row[key])

        # for row in data2:
        #     for key in keys_to_check_func_2:
        #         if row[key] > 0:
        #             row[key] = row[key]
        #         else:
        #             row[key] = ''
        # for row in data2:
        #     for key in keys_to_check_func_2:
        #         if row[key] != "":
        #             row[key] = "{:,.0f}".format(row[key])

        formatted_total_netsurplus = {
            acct_per: "${:,}".format(abs(int(value)))
            if value > 0
            else "(${:,})".format(abs(int(value)))
            if value < 0
            else ""
            for acct_per, value in total_netsurplus.items()
            if value != 0
        }
        formatted_total_DnA = {
            acct_per: "{:,}".format(abs(int(value)))
            if value >= 0
            else "({:,})".format(abs(int(value)))
            if value < 0
            else ""
            for acct_per, value in total_DnA.items()
            if value != 0
        }

        formated_ytdnetsurplus = format_with_parentheses(ytd_netsurplus)

        bs_activity_list = list(
            set(row["Activity"] for row in data_balancesheet if "Activity" in row)
        )
        bs_activity_list_sorted = sorted(bs_activity_list)
        gl_obj = list(set(row["obj"] for row in data3 if "obj" in row))
        gl_obj_sorted = sorted(gl_obj)


        context = {
            "data_balancesheet": data_balancesheet,
            "data_activitybs": data_activitybs,
            # "data3": data3,
            "bs_activity_list": bs_activity_list_sorted,
            "gl_obj": gl_obj_sorted,
            # "button_rendered": button_rendered,
        
            "last_month": last_month,
            "last_month_number": last_month_number,
            "last_month_name": last_month_name,
            "FY_year_1":FY_year_1,
            "FY_year_2":FY_year_2,
            "totals_bs":{
                "total_current_assets":total_current_assets,
                "total_current_assets_fye":total_current_assets_fye,
                "last_month_current_assets":last_month_current_assets,
                "total_capital_assets":total_capital_assets,
                "total_capital_assets_fye":total_capital_assets_fye,
                "last_month_total_capital_assets":last_month_total_capital_assets,
                "total_current_liabilities":total_current_liabilities,
                "total_current_liabilities_fye":total_current_liabilities_fye,
                "last_month_total_current_liabilities":last_month_total_current_liabilities,
                "total_liabilities":total_liabilities,
                "total_liabilities_fye":total_liabilities_fye,
                "last_month_total_liabilities":last_month_total_liabilities,
                "total_assets": total_assets,
                "total_assets_fye":total_assets_fye,
                "last_month_total_assets":last_month_total_assets,
                "total_LNA_fye":total_LNA_fye,
                "total_LNA":total_LNA,
                "total_current_assets_fytd":total_current_assets_fytd,
                "total_capital_assets_fytd":total_capital_assets_fytd,
                "total_current_liabilities_fytd":total_current_liabilities_fytd,
                "total_liabilities_fytd":total_liabilities_fytd,
                "total_assets_fye_fytd":total_assets_fye_fytd,
                "total_net_assets_fytd":total_net_assets_fytd,
                "total_LNA_fytd":total_LNA_fytd,
                "last_month_total_LNA":last_month_total_LNA
            }

            
            #"difference_key": difference_key,
            # "format_ytd_budget": formatted_ytd_budget,
            # "ytd_budget": ytd_budget,
        }

        if school in schoolCategory["ascender"]:
            context["total_DnA"] = (formatted_total_DnA,)
            context["total_netsurplus"] = formatted_total_netsurplus
            context["total_SBD"] = total_SBD
            context["ytd_netsurplus"] = formated_ytdnetsurplus
    # return context
    # dict_keys = ["data", "data2", "data3", "data_expensebyobject", "data_activities"]

        if FY_year_1 == present_year:
            relative_path = os.path.join("balance-sheet-asc", school)
        else:
            relative_path = os.path.join(str(FY_year_1), "balance-sheet-asc", school)

        # json_path = JSON_DIR.path(relative_path)  
        json_path = os.path.join(JSON_DIR,relative_path)

        shutil.rmtree(json_path, ignore_errors=True)
        if not os.path.exists(json_path):
            os.makedirs(json_path)

        for key, val in context.items():
            file = os.path.join(json_path, f"{key}.json")
            with open(file, "w") as f:
                json.dump(val, f)


def school_status(request):
    month_names = {
    1: "January",
    2: "February",
    3: "March",
    4: "April",
    5: "May",
    6: "June",
    7: "July",
    8: "August",
    9: "September",
    10: "October",
    11: "November",
    12: "December"
    }

    

    sept = ["09","10","11","12","01","02","03","04","05","06","07","08"]
    july = ["07","08","09","10","11","12","01","02","03","04","05","06"]
    school_data = []
    BS_status=""
    for key,value in SCHOOLS.items():
        print(key)
        
   
        BS_status = ""
        pl_path = os.path.join("profit-loss", key)
        js_path = os.path.join(JSON_DIR, pl_path)
        
        with open(os.path.join(js_path, "months.json"), "r") as f:
            months = json.load(f)

        with open(os.path.join(js_path, "data.json"), "r") as f:
            data = json.load(f)

        with open(os.path.join(js_path, "data2.json"), "r") as f:
            data2 = json.load(f)

        with open(os.path.join(js_path, "totals.json"), "r") as f:
            totals = json.load(f)

        

        

        if os.path.exists(os.path.join(js_path, "last_update.json")):
            with open(os.path.join(js_path, "last_update.json"), "r") as f:
                last_update = json.load(f)
        else:
            last_update = ""

        with open(os.path.join(js_path, "data_expensebyobject.json"), "r") as f:
            data_expensebyobject = json.load(f)

        relative_path = os.path.join("balance-sheet", key)
        json_path = os.path.join(JSON_DIR, relative_path)
        with open(os.path.join(json_path, "totals_bs.json"), "r") as f:
            totals_bs = json.load(f)

        cashflow_path = os.path.join("cashflow", key)
        cashflow_file = os.path.join(JSON_DIR, cashflow_path)
        if os.path.exists(os.path.join(cashflow_file, "cf_totals.json")):
            with open(os.path.join(cashflow_file, "cf_totals.json"), "r") as f:
                cf_totals = json.load(f)
        else:
            cf_totals = ""


        PLbudget_status = "No Budget"
        PLbudget_counter = 0
        for row in data:
            budget = row.get("total_budget")
   
            if budget != "":
                PLbudget_counter += 1

        if PLbudget_counter > 1:
            PLbudget_status ="True"
                
               

        

     
        if "month_exception" in months:
            last_month_number = months["last_month_number"]
           
            month_exception = months["month_exception"]
            month_exception_str = months["month_exception_str"]
            last_month_number_str = str(last_month_number).zfill(2)
            
            total_LNA = totals_bs.get("total_LNA", {})
            total_assets = totals_bs.get("total_assets", {})
            month_exception_str = months["month_exception_str"]

            ytd_netsurplus = totals["ytd_netsurplus"]
            variances_netsurplus = totals["variances_netsurplus"]
            ytd_netincome = totals["ytd_net_income"]
            variances_netincome = totals["variances_net_income"]

            pl_balanced = "Not Balanced"
            if ytd_netsurplus == ytd_netincome and variances_netsurplus == variances_netincome:
                pl_balanced = "True"
               
      
            month_name = month_names.get(last_month_number, "Current")

            PLrevenue_status = f"No Revenue for {month_name}"
            PLrevenue_counter = 0
            for row in data:
                revenue = row.get(f'total_real{last_month_number}')
                if revenue != "":
                    PLrevenue_counter += 1

            if PLrevenue_counter > 1:
                PLrevenue_status ="True"

            PLexpense_status =f"No Expense for {month_name}"
            PLexpense_counter = 0
            for row in data2:
                expense = row.get(f'total_func{last_month_number}')
                if expense != "":
                    PLexpense_counter += 1

            if PLexpense_counter>1:
                PLexpense_status = "True"

            PLtotalexpense_status = f"No Expense by object for {month_name}"
            PLtotalexpense_counter = 0        
            for row in data_expensebyobject:
                total_expense = row.get(f'total_expense{last_month_number}')
                if total_expense != "":
                    PLtotalexpense_counter += 1
            
            if PLtotalexpense_counter > 1:
                PLtotalexpense_status = "True"

            
      

            if 'cfchecker' in cf_totals and isinstance(cf_totals['cfchecker'], dict):
                
                checker_values = [
                    cf_totals['cfchecker'].get(month, 0) for month in ["09", "10", "11", "12", "01", "02", "03", "04", "05", "06", "07", "08"]
                ]
            
                all_zero = all(value == 0 for value in checker_values)

               
                CF_status = "BALANCED" if all_zero else "NOT BALANCED"
                print(key, CF_status)
            else:
                CF_status = "NO DATA"



            if key in schoolMonths["septemberSchool"]:
                for month in sept:
                    if month == month_exception_str:
         

                        break
                    
                    else:
                        total_LNA_value = int(total_LNA.get(month, "0").replace(",", "").replace("$", "").replace("(", "-").replace(")", ""))
                        total_assets_value = int(total_assets.get(month, "0").replace(",", "").replace("$", "").replace("(", "-").replace(")", ""))
                        BS_status = "BALANCED"
                    
                        if total_LNA_value != total_assets_value:
                            BS_status = "NOT BALANCED"
            else:
                for month in july:
                    if month == month_exception_str:
                        
                        break
                    
                    else:
                        total_LNA_value = int(total_LNA.get(month, "0").replace(",", "").replace("$", "").replace("(", "-").replace(")", ""))
                        total_assets_value = int(total_assets.get(month, "0").replace(",", "").replace("$", "").replace("(", "-").replace(")", ""))
                        BS_status = "BALANCED"
                    
                        if total_LNA_value != total_assets_value:
                            BS_status = "NOT BALANCED"


            db_string = (db[key]['db'])
            db_string = db_string.strip('[]')
            cnxn = connect()
            cursor = cnxn.cursor()
            query = "SELECT * FROM [dbo].[AscenderDownloader] WHERE db = ?"
            cursor.execute(query,db_string)
            print(db_string)
            row = cursor.fetchone()
            update_status = ""
            if row:
                update_status = row[5]
                print("db",row[4])

            ascender = 'True'
            if key in schoolCategory["skyward"]:
                ascender = 'False'

        
        if key in schoolCategory["ascender"]:
            row_data = {
                "school_key":key,
                "school_name": value,
                "school_category": "ascender",
                "BS_status": BS_status,
                "PLbudget_status":PLbudget_status,
                "PLrevenue_status":PLrevenue_status,
                "PLexpense_status":PLexpense_status,
                "PLtotalexpense_status":PLtotalexpense_status,
                "pl_balanced":pl_balanced,
                "last_update": last_update,
                "CF_status": CF_status,
                "update_status":update_status,
                "ascender":ascender,
            }
            school_data.append(row_data)
        else:
            row_data = {
                "school_key":key,
                "school_name": value,
                "school_category": "skyward",
                "BS_status": BS_status,
                "PLbudget_status":PLbudget_status,
                "PLrevenue_status":PLrevenue_status,
                "PLexpense_status":PLexpense_status,
                "PLtotalexpense_status":PLtotalexpense_status,
                "pl_balanced":pl_balanced,
                "last_update": last_update,
                "update_status":update_status,
                "ascender":ascender,
            }
            school_data.append(row_data)

    def custom_sort(entry):

        if entry["school_key"] == "advantage":
            return (0, entry["school_key"])
        else:
            return (1, entry["school_key"])

    sorted_school_data = sorted(school_data, key=custom_sort)




    context = {
      
        'school_data': sorted_school_data
    }

    cursor.close()
    cnxn.close()    


    json_path = os.path.join(JSON_DIR,"school-status")
    shutil.rmtree(json_path, ignore_errors=True)
    if not os.path.exists(json_path):
        os.makedirs(json_path)
    for key, val in context.items():
        file = os.path.join(json_path, f"{key}.json")
        with open(file, "w") as f:
            json.dump(val, f)

if __name__ == "__main__":
    update_db()
    # charter_first("advantage")
