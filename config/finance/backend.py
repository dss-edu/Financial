from .connect import connect
from time import strftime
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import json
import os
import re
import pprint
from collections import defaultdict

pp = pprint.PrettyPrinter()
# Get the current date
current_date = datetime.now()
# Extract the month number from the current date
month_number = current_date.month
curr_year = current_date.year

JSON_DIR = os.path.join(os.getcwd(), "finance", "json")

SCHOOLS = {
    "advantage": "ADVANTAGE ACADEMY",
    "cumberland": "CUMBERLAND ACADEMY",
    "village-tech": "VILLAGE TECH",
    "prepschool": "LEADERSHIP PREP SCHOOL",
    "manara": "MANARA ACADEMY",
}

db = {
    "advantage": {
        "object": "[PL_Definition_obj]",
        "function": "[PL_Definition_func]",
        "db": "[AscenderData_Advantage_new]",
        "code": "[PL_ExpensesbyObjectCode]",
        "activities": "[PL_Activities]",
        "bs": "[AscenderData_Advantage_Balancesheet]",
        "bs_activity": "[ActivityBS]",
        "cashflow": "[AscenderData_Advantage_Cashflow]",
        "adjustment": "[Adjustment]",
        "bs_fye":"[Balancesheet_FYE]",
    },
    "cumberland": {
        "object": "[PL_Definition_obj]",
        "function": "[PL_Definition_func]",
        "db": "[AscenderData_Cumberland_new]",
        "code": "[PL_ExpensesbyObjectCode]",
        "activities": "[PL_Activities]",
        "bs": "[AscenderData_Advantage_Balancesheet]",
        "bs_activity": "[ActivityBS]",
        "cashflow": "[AscenderData_Advantage_Cashflow]",
        "adjustment": "[Adjustment]",
        "bs_fye":"[Balancesheet_FYE]",
    },
    "village-tech": {
        "object": "[PL_Definition_obj]",
        "function": "[PL_Definition_func]",
        "db": "[Skyward_VillageTech]",
        "code": "[PL_ExpensesbyObjectCode]",
        "activities": "[PL_Activities]",
        "bs": "[AscenderData_Advantage_Balancesheet]",
        "bs_activity": "[ActivityBS]",
        "cashflow": "[AscenderData_Advantage_Cashflow]",
        "adjustment": "[Adjustment]",
        "bs_fye":"[Balancesheet_FYE]",
    },
    "prepschool": {
        "object": "[PL_Definition_obj]",
        "function": "[PL_Definition_func]",
        "db": "[AscenderData_Leadership]",
        "code": "[PL_ExpensesbyObjectCode]",
        "activities": "[PL_Activities]",
        "bs": "[AscenderData_Advantage_Balancesheet]",
        "bs_activity": "[ActivityBS]",
        "cashflow": "[AscenderData_Advantage_Cashflow]",
        "adjustment": "[Adjustment]",
        "bs_fye":"[Balancesheet_FYE]",
    },
    "manara": {
        "object": "[PL_Definition_obj]",
        "function": "[PL_Definition_func]",
        "db": "[AscenderData_Manara]",
        "code": "[PL_ExpensesbyObjectCode]",
        "activities": "[PL_Activities]",
        "bs": "[AscenderData_Advantage_Balancesheet]",
        "bs_activity": "[ActivityBS]",
        "cashflow": "[AscenderData_Advantage_Cashflow]",
        "adjustment": "[Adjustment]",
        "bs_fye":"[Balancesheet_FYE]",
    },
}


def update_db():
    # profit_loss("advantage")
    # balance_sheet("cumberland")
    # cashflow("advantage")
    # excel("advantage")
    # balance_sheet("manara")
    # profit_loss("manara")
    for school, name in SCHOOLS.items():
        profit_loss(school) 
        balance_sheet(school)
        cashflow(school)
        excel(school)
        
def update_school(school):
    profit_loss(school) 
    balance_sheet(school)
    cashflow(school)
    excel(school)




def profit_loss(school):


    
    
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
    if not school == "village-tech":
        cursor.execute(
            f"SELECT * FROM [dbo].{db[school]['db']}  as AA where AA.Number != 'BEGBAL';"
        )
    else:
        cursor.execute(f"SELECT * FROM [dbo].{db[school]['db']};")

    rows = cursor.fetchall()

    data3 = []

    if not school == "village-tech":
        for row in rows:
            expend = float(row[17])
            date = row[11]
            if isinstance(row[11], datetime):
                date = row[11].strftime("%Y-%m-%d")

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

    cursor.execute(f"SELECT * FROM [dbo].{db[school]['adjustment']} ")
    rows = cursor.fetchall()

    adjustment = []

    if school != "village-tech":
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

    
    current_date = datetime.today().date()
    
    current_year = current_date.year
    next_year = current_date.year + 1
    last_year = current_date.year - 1
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_name = last_month.strftime("%B")
    formatted_last_month = last_month.strftime('%B %d, %Y')
    last_month_number = last_month.month
    if school == 'manara' or school == 'prepschool':
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
    if school == "village-tech":
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
                )
            item[f"total_check{i}"] = total_real + total_adjustment


    july_date  = datetime(current_year, 7, 1).date()
    september_date  = datetime(current_year, 9, 1).date()
    FY_year_1 = last_year
    FY_year_2 = current_year
    for item in data3:
        date_str = item["Date"]
        if date_str:
            if school == 'manara' or school == 'prepschool':
               
                date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
                if date_obj > july_date: # if date is higher than july 1 this year
                  FY_year_1 = current_year
                  FY_year_2 = next_year
            else:
                date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
                
                if date_obj > september_date: # if date is higher than july 1 this year
                  
                  FY_year_1 = current_year
                  FY_year_2 = next_year
                
           
            


    #checks if the last month column is empty. if empty. last month will be set to  last two months.
    if all(item[f"total_check{last_month_number}"] == 0 for item in data):
        last_2months = current_month - relativedelta(months=1)
        last_2months = last_2months - relativedelta(days=1)
        last_month_number = last_2months.month
        last_month_name = last_2months.strftime("%B")
        formatted_last_month = last_2months.strftime('%B %d, %Y')
        last_month_number = last_2months.month
        if school == 'manara' or school == 'prepschool':
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
        if school == "village-tech":
               
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

        item[f"ytd_budget"] = format_value(item[f"ytd_budget"])
    


    
    ytd_total_revenue = abs(sum(total_revenue.values()))  
   
    ytd_ammended_total = totals["total_ammended"] * ytd_budget
    ytd_ammended_total_lr = totals["total_ammended_lr"] * ytd_budget
    ytd_ammended_total_spr = totals["total_ammended_spr"] * ytd_budget
    ytd_ammended_total_fpr = totals["total_ammended_fpr"] * ytd_budget

    variances_revenue = (ytd_total_revenue - ytd_ammended_total)
    variances_revenue_lr = (ytd_total_lr + ytd_ammended_total_lr)
    variances_revenue_spr = (ytd_total_spr + ytd_ammended_total_spr)
    variances_revenue_fpr = (ytd_total_fpr + ytd_ammended_total_fpr)

    var_ytd = "{:d}%".format(abs(int(ytd_total_revenue / totals["total_ammended"]*100))) if totals["total_ammended"] != 0 else ""
    var_ytd_lr = "{:d}%".format(abs(int(ytd_total_lr / totals["total_ammended_lr"]*100))) if totals["total_ammended_lr"] != 0 else ""
    var_ytd_spr = "{:d}%".format(abs(int(ytd_total_spr / totals["total_ammended_spr"]*100))) if totals["total_ammended_spr"] != 0 else ""
    var_ytd_fpr = "{:d}%".format(abs(int(ytd_total_fpr / totals["total_ammended_fpr"]*100))) if totals["total_ammended_fpr"] != 0 else ""
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

    print(school)
    for item in data2:
        if item["category"] != "Depreciation and Amortization":
            func = item["func_func"]
            obj = item["obj"]
            
            ytd_total = 0


            if school == 'village-tech':
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
            
            if school == 'village-tech':
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
            item["var_ytd"] =  "{:d}%".format(abs(int(item['total_budget'] / item["ytd_total"]*100))) if item["ytd_total"] != 0 else ""
    
    ytd_ammended_total_first = first_total * ytd_budget
    var_ytd_first_total = "{:d}%".format(abs(int(first_ytd_total / ytd_ammended_total_first*100))) if ytd_ammended_total_first != 0 else ""


    for item in data2:
        if item["category"] == "Depreciation and Amortization":
            func = item["func_func"]
            obj = item["obj"]
            ytd_total = 0
            
           
            if school == 'village-tech':
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
                ytd_total += (item[f"total_func2_{month_number}"])
        
            item["ytd_total"] = ytd_total
            dna_total += item['total_budget']
            dna_ytd_total += item["ytd_total"]
            item[f"ytd_budget"] = item['total_budget'] * ytd_budget
            item["variances"] =  item[f"ytd_budget"] -item["ytd_total"]
            variances_dna+= item["variances"]
            item["var_ytd"] =  "{:d}%".format(abs(int(item['total_budget'] / item["ytd_total"]*100))) if item["ytd_total"] != 0 else ""
            ytd_ammended_dna = dna_total * ytd_budget
            var_ytd_dna = "{:d}%".format(abs(int(dna_ytd_total / ytd_ammended_dna*100))) if ytd_ammended_dna != 0 else ""
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
    var_SBD = "{:d}%".format(abs(int(  ytd_SBD/ ammended_budget_SBD*100))) if ammended_budget_SBD != 0 else ""
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
    var_netsurplus = "{:d}%".format(abs(int(ytd_netsurplus / ammended_budget_netsurplus*100))) if ammended_budget_netsurplus != 0 else ""

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

        if school == 'village-tech':
            item["total_budget"] = sum(
                entry[appr_key]
                for entry in data3
                if entry["obj"] == obj

                )
        else:
            item["total_budget"] = sum(
            entry[appr_key]
            for entry in data3
            if entry["obj"] == obj
            and entry["Type"] == 'GJ'
       
            )
        
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
            item["var_EOC"] = "{:d}%".format(abs(int(ytd_EOC_pc / total_budget_pc*100))) if total_budget_pc != 0 else ""
        elif obj == "6200":
            category = "Professional and Contract Services"
            item["variances"] = ytd_budget_pcs - ytd_EOC_pcs
            item["var_EOC"] = "{:d}%".format(abs(int(ytd_EOC_pcs / total_budget_pcs*100))) if total_budget_pcs != 0 else ""
        elif obj == "6300":
            category = "Materials and Supplies"
            item["variances"] = ytd_budget_sm - ytd_EOC_sm
            item["var_EOC"] = "{:d}%".format(abs(int(ytd_EOC_sm / total_budget_sm*100))) if total_budget_sm != 0 else ""
        elif obj == "6400":
            category = "Other Operating Costs"
            item["variances"] = ytd_budget_ooe - ytd_EOC_ooe
            item["var_EOC"] = "{:d}%".format(abs(int(ytd_EOC_ooe / total_budget_ooe*100))) if total_budget_ooe != 0 else ""
        elif obj == "6449":
            category = "Depreciation"
            item["variances"] = ytd_budget_oe - ytd_EOC_oe
            item["var_EOC"] = "{:d}%".format(abs(int(ytd_EOC_oe / total_budget_oe*100))) if total_budget_oe != 0 else ""
        else:
            category = "Debt Services"
            item["variances"] = ytd_budget_te - ytd_EOC_te
            item["var_EOC"] = "{:d}%".format(abs(int(ytd_EOC_te / total_budget_te*100))) if total_budget_te != 0 else ""

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f"total_expense{i}"] = sum(
                entry[f"total_activities{i}"]
                for entry in data_activities
                if entry["Category"] == category
            )

        
    #CONTINUATION COMPUTATION TOTAL EXPENSE
    total_expense_ytd = sum([ytd_EOC_te, ytd_EOC_ooe, ytd_EOC_sm, ytd_EOC_pcs, ytd_EOC_pc,dna_ytd_total])
    variances_total_expense = total_expense_ytd_budget - total_expense_ytd
    var_total_expense = "{:d}%".format(abs(int(total_expense_ytd / total_expense*100))) if total_expense != 0 else ""
        

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
    variances_net_income = variances_revenue - variances_total_expense
    var_net_income = "{:d}%".format(abs(int(ytd_net_income / net_income_budget*100))) if net_income_budget != 0 else ""


    
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
            row["variances"] = format_value_negative(variances)

    # FOR EXPENSE BY OBJECT DEPRECIATION ONLY        
    dna_total_6449 = format_value_negative(dna_total)
    ytd_ammended_dna_6449 = format_value_negative(ytd_ammended_dna)
    dna_ytd_total_6449 = format_value(dna_ytd_total)
    variances_dna_6449 = format_value(variances_dna)   
    dna_total_months_6449 = {acct_per: format_value(value) for acct_per, value in dna_total_months.items() if value != 0}    


    #FORMAT FIRST TOTAL AND DEPRECIATION AND AMORTIZATION(DNA)
    dna_total = format_value_dollars_negative(dna_total)
    first_total = format_value_dollars_negative(first_total)

    ytd_ammended_dna = format_value_dollars_negative(ytd_ammended_dna)
    ytd_ammended_total_first = format_value_dollars_negative(ytd_ammended_total_first)
    
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
            row[f"total_budget"] = format_value_negative(total_expense_budget)

        if ytd_budget is None or ytd_budget == 0:
            row[f"ytd_budget"] = ""
        else:
            row[f"ytd_budget"] = format_value_negative(ytd_budget)
    
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

    ytd_EOC_pc  = format_value_negative(ytd_EOC_pc)
    ytd_EOC_pcs = format_value_negative(ytd_EOC_pcs)
    ytd_EOC_sm  = format_value_negative(ytd_EOC_sm)
    ytd_EOC_ooe = format_value_negative(ytd_EOC_ooe)
    ytd_EOC_te  = format_value_negative(ytd_EOC_te)
    ytd_EOC_oe  = format_value_negative(ytd_EOC_oe)

    total_budget_pc = format_value_negative(total_budget_pc)
    total_budget_pcs = format_value_negative(total_budget_pcs)
    total_budget_sm = format_value_negative(total_budget_sm)
    total_budget_ooe = format_value_negative(total_budget_ooe)
    total_budget_oe = format_value_negative(total_budget_oe)   
    total_budget_te = format_value_negative(total_budget_te)

    ytd_budget_pc = format_value_negative(ytd_budget_pc)
    ytd_budget_pcs =format_value_negative(ytd_budget_pcs)
    ytd_budget_sm = format_value_negative(ytd_budget_sm)
    ytd_budget_ooe = format_value_negative(ytd_budget_ooe)
    ytd_budget_oe = format_value_negative(ytd_budget_oe)
    ytd_budget_te = format_value_negative(ytd_budget_te)




    #EXPENSE OBJECT FOR FIX
    budget_for_6500 = format_value(budget_for_6500)
    ytd_budget_for_6500 = format_value(ytd_budget_for_6500)

    #FORMAT TOTAL EXPENSE
    total_expense = format_value_dollars_negative(total_expense)
    total_expense_ytd_budget = format_value_dollars_negative(total_expense_ytd_budget)
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
            if key in row and row[key] is not None and row[key] > 0:
                row[key] = row[key]
            else:
                row[key] = ""
    for row in data2:
        for key in keys_to_check_func:
            if row[key] != "":
                row[key] = "{:,.0f}".format(row[key])

    for row in data2:
        for key in keys_to_check_func_2:
            if key in row and row[key] is not None and row[key] > 0:
                row[key] = row[key]
            else:
                row[key] = ""
    for row in data2:
        for key in keys_to_check_func_2:
            if row[key] != "":
                row[key] = "{:,.0f}".format(row[key])



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

    json_path = os.path.join(JSON_DIR, "profit-loss", school)
    if not os.path.exists(json_path):
        os.makedirs(json_path)

    for key, val in context.items():
        file = os.path.join(json_path, f"{key}.json")
        with open(file, "w") as f:
            json.dump(val, f)


def balance_sheet(school):
    current_date = datetime.today().date()
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_name = last_month.strftime("%B")
    formatted_last_month = last_month.strftime('%B %d, %Y')
    last_month_number = last_month.month


    cnxn = connect()
    cursor = cnxn.cursor()

    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['bs']} AS T1 LEFT JOIN [dbo].{db[school]['bs_fye']} AS T2 ON T1.BS_id = T2.BS_id ;  ")
    rows = cursor.fetchall()

    data_balancesheet = []

    for row in rows:
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

        row_dict = {
            "Activity": row[0],
            "Description": row[1],
            "Category": row[2],
            "Subcategory": row[3],
            "FYE": fyeformat,
            "BS_id": row[5],
            "school": row[8],

        }

        data_balancesheet.append(row_dict)

    # cursor.execute(f"SELECT  * FROM [dbo].{db[school]['object']};")
    # rows = cursor.fetchall()
    #
    # data = []
    json_path = os.path.join(JSON_DIR, "profit-loss", school)
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

    cursor.execute(f"SELECT * FROM [dbo].{db[school]['adjustment']} ")
    rows = cursor.fetchall()

    adjustment = []

    if school != "village-tech":
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
    if school == "village-tech":
        bal_key = "Amount"
        real_key = "Amount"
        expend_key = "Amount"

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
            # if i == last_month_number and (item[f"total_real{i}"] == 0):

            #         last_2months = current_month - relativedelta(months=1)
            #         last_2months = last_2months - relativedelta(days=1)
            #         last_month_name = last_2months.strftime("%B")
            #         formatted_last_month = last_2months.strftime('%B %d, %Y')
            
            total_revenue[acct_per] += (item[f"total_real{i}"])

    if all(item[f"total_real{last_month_number}"] == 0 for item in data):
        last_2months = current_month - relativedelta(months=1)
        last_2months = last_2months - relativedelta(days=1)
        last_month_number = last_2months.month
        last_month_name = last_2months.strftime("%B")
        formatted_last_month = last_2months.strftime('%B %d, %Y')                    


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
        if value > 0:
            return "${:,.0f}".format(round(value))
        elif value < 0:
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

    for row in data_balancesheet:
        if row["school"] == school:
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

            if school != 'manara' or school != 'prepschool':
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
 
                total_sum_value = totals.get(f"total_sum{last_month_number + 1}_value", 0)
                row["last_month_difference"] = row[f"difference_{last_month_number}"] + total_sum_value

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
                row["debt_fytd"] = -( total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value)

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
                
                total_sum_value = totals.get(f"total_sum{last_month_number}_value", 0)
                row["last_month_difference"] = row[f"difference_{last_month_number-1}"] + total_sum_value
                

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
  
                row["debt_fytd"] = -( total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value)


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
    
    total_net_assets_fytd = totals["bs_ytd_netsurplus"]    #assign the value coming from profitloss totals
    
    FY_year_1 = months["FY_year_1"] 
    FY_year_2 = months["FY_year_2"]

    for row in data_balancesheet:
        if row["school"] == school:
            subcategory =  row["Subcategory"]
            fye =  float(row["FYE"].replace("$","").replace(",", "").replace("(", "-").replace(")", "")) if row["FYE"] else 0

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
            fye =  float(row["FYE"].replace("$","").replace(",", "").replace("(", "-").replace(")", "")) if row["FYE"] else 0
            if subcategory == 'Long Term Debt':
                for i, acct_per in enumerate(acct_per_values,start = 1):
                    total_liabilities[acct_per] += row[f"debt_{i}"] + total_current_liabilities[acct_per]
                total_liabilities_fytd_2 += row["debt_fytd"]
                total_liabilities_fye +=  + total_current_liabilities_fye + fye
    total_liabilities_fytd = total_liabilities_fytd_2 + total_current_liabilities_fytd

    for row in data_balancesheet:
        if row["school"] == school:
            fye =  float(row["FYE"].replace("$","").replace(",", "").replace("(", "-").replace(")", "")) if row["FYE"] else 0
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
    print(school,total_assets_fye_fytd)


    
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
    total_assets_fye_fytd = format_value(total_assets_fye_fytd)
    total_LNA_fytd = format_value(total_LNA_fytd)

    total_current_assets = {acct_per: format_value(value) for acct_per, value in total_current_assets.items() if value != 0}
    total_capital_assets = {acct_per: format_value(value) for acct_per, value in total_capital_assets.items() if value != 0}
    total_current_liabilities = {acct_per: format_value(value) for acct_per, value in total_current_liabilities.items() if value != 0}
    total_liabilities = {acct_per: format_value(value) for acct_per, value in total_liabilities.items() if value != 0}
    total_assets = {acct_per: format_value_dollars(value) for acct_per, value in total_assets.items() if value != 0}
    total_LNA = {acct_per: format_value_dollars(value) for acct_per, value in total_LNA.items() if value != 0}

   
    for row in data_balancesheet:
        if row["school"] == school:
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

    for row in data_activitybs:
        for key in keys_to_check:
            value = float(row[key])
            if value == 0:
                row[key] = ""
            elif value < 0:
                row[key] = "({:,.0f})".format(abs(float(row[key])))
            elif value != "":
                row[key] = "{:,.0f}".format(float(row[key]))


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

    #difference_key = "difference_" + str(last_month_number)
    context = {
        "data_balancesheet": data_balancesheet,
        "data_activitybs": data_activitybs,
        # "data3": data3,
        "bs_activity_list": bs_activity_list_sorted,
        "gl_obj": gl_obj_sorted,
        # "button_rendered": button_rendered,
       
        "last_month": formatted_last_month,
        "last_month_number": last_month_number,
        "last_month_name": last_month_name,
        "FY_year_1":FY_year_1,
        "FY_year_2":FY_year_2,
        "totals_bs":{
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

        
        #"difference_key": difference_key,
        # "format_ytd_budget": formatted_ytd_budget,
        # "ytd_budget": ytd_budget,
    }

    if not school == "village-tech":
        context["total_DnA"] = (formatted_total_DnA,)
        context["total_netsurplus"] = formatted_total_netsurplus
        context["total_SBD"] = total_SBD
        context["ytd_netsurplus"] = formated_ytdnetsurplus
    # return context

    # dict_keys = ["data", "data2", "data3", "data_expensebyobject", "data_activities"]

    json_path = os.path.join(JSON_DIR, "balance-sheet", school)
    if not os.path.exists(json_path):
        os.makedirs(json_path)

    for key, val in context.items():
        file = os.path.join(json_path, f"{key}.json")
        with open(file, "w") as f:
            json.dump(val, f)

def cashflow(school):
    cnxn = connect()
    cursor = cnxn.cursor()

    json_path = os.path.join(JSON_DIR, "profit-loss", school)
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

    json_path = os.path.join(JSON_DIR, "balance-sheet", school)
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

    activity_key = "Bal"
    if school == "village-tech":
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
            item["fytd_2"] += item[f"total_investing{i}"] 
            

    data_key = "Expend"
    if school == "village-tech":
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

    # --- total revenue
    # total_revenue = {acct_per: 0 for acct_per in acct_per_values}
    #
    # data_key = "Real"
    # if school == "village-tech":
    #     data_key = "Amount"
    #
    # for item in data:
    #     fund = item["fund"]
    #     obj = item["obj"]
    #
    #     for i, acct_per in enumerate(acct_per_values, start=1):
    #         item[f"total_real{i}"] = sum(
    #             entry[data_key]
    #             for entry in data3
    #             if entry["fund"] == fund
    #             and entry["obj"] == obj
    #             and entry["AcctPer"] == acct_per
    #         )
    #
    #         total_revenue[acct_per] += abs(item[f"total_real{i}"])

    # keys_to_check = [
    #     "total_real1",
    #     "total_real2",
    #     "total_real3",
    #     "total_real4",
    #     "total_real5",
    #     "total_real6",
    #     "total_real7",
    #     "total_real8",
    #     "total_real9",
    #     "total_real10",
    #     "total_real11",
    #     "total_real12",
    # ]
    #
    # acct_per_values2 = [
    #     "01",
    #     "02",
    #     "03",
    #     "04",
    #     "05",
    #     "06",
    #     "07",
    #     "08",
    #     "09",
    #     "10",
    #     "11",
    #     "12",
    # ]

    # total_surplus = {acct_per: 0 for acct_per in acct_per_values2}

    # data_key = "Expend"
    # if school == "village-tech":
    #     data_key = "Amount"
    # for item in data2:
    #     if item["category"] != "Depreciation and Amortization":
    #         func = item["func_func"]
    #
    #         for i, acct_per in enumerate(acct_per_values2, start=1):
    #             item[f"total_func{i}"] = sum(
    #                 entry[data_key]
    #                 for entry in data3
    #                 if entry["func"] == func and entry["AcctPer"] == acct_per
    #             )
    #             total_surplus[acct_per] += item[f"total_func{i}"]
    #
    # ---- Depreciation and ammortization total
    # total_DnA = {acct_per: 0 for acct_per in acct_per_values2}
    #
    # data_key = "Expend"
    # if school == "village-tech":
    #     data_key = "Amount"
    # for item in data2:
    #     func = item["func_func"]
    #     obj = item["obj"]
    #
    #     for i, acct_per in enumerate(acct_per_values2, start=1):
    #         item[f"total_func2_{i}"] = sum(
    #             entry[data_key]
    #             for entry in data3
    #             if entry["func"] == func
    #             and entry["AcctPer"] == acct_per
    #             and entry["obj"] == obj
    #         )
    #         total_DnA[acct_per] += item[f"total_func2_{i}"]
    #
    # total_SBD = {
    #     acct_per: total_revenue[acct_per] - total_surplus[acct_per]
    #     for acct_per in acct_per_values
    # }
    # total_netsurplus = {
    #     acct_per: total_SBD[acct_per] - total_DnA[acct_per]
    #     for acct_per in acct_per_values
    # }
    # formatted_total_netsurplus = {
    #     acct_per: "${:,}".format(abs(int(value)))
    #     if value > 0
    #     else "(${:,})".format(abs(int(value)))
    #     if value < 0
    #     else ""
    #     for acct_per, value in total_netsurplus.items()
    #     if value != 0
    # }
    # formatted_total_DnA = {
    #     acct_per: "{:,}".format(abs(int(value)))
    #     if value >= 0
    #     else "({:,})".format(abs(int(value)))
    #     if value < 0
    #     else ""
    #     for acct_per, value in total_DnA.items()
    #     if value != 0
    # }
    #
    # if formatted_ytd_budget.startswith("0."):
    #     formatted_ytd_budget = formatted_ytd_budget[2:]

    cashflow_path = os.path.join(JSON_DIR, 'cashflow', school)
    if not os.path.exists(cashflow_path):
        os.makedirs(cashflow_path)

    cashflow_file = os.path.join(cashflow_path, "data_cashflow.json")
    with open(cashflow_file, "w") as f:
        json.dump(data_cashflow, f)

    cursor.close()
    cnxn.close()

def excel(school):
        
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
    if not school == "village-tech":
        cursor.execute(
            f"SELECT * FROM [dbo].{db[school]['db']}  as AA where AA.Number != 'BEGBAL';"
        )
    else:
        cursor.execute(f"SELECT * FROM [dbo].{db[school]['db']};")

    rows = cursor.fetchall()
    data3 = []

    if not school == "village-tech":
        for row in rows:
            expend = float(row[17])
            date = row[11]
            if isinstance(row[11], datetime):
                date = row[11].strftime("%Y-%m-%d")

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

    cursor.execute(f"SELECT * FROM [dbo].{db[school]['adjustment']} ")
    rows = cursor.fetchall()

    adjustment = []

    if school != "village-tech":
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
    cursor.execute("SELECT * FROM [dbo].[AscenderData_CharterFirst]") 
    rows = cursor.fetchall()

    data_charterfirst = []

    for row in rows:
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
    #BS START

    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['bs']} AS T1 LEFT JOIN [dbo].{db[school]['bs_fye']} AS T2 ON T1.BS_id = T2.BS_id ;  ")
    rows = cursor.fetchall()

    data_balancesheet = []

    for row in rows:
        fye = float(row[7]) if row[7] else 0
       

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
        row_dict = {
            "Activity": row[0],
            "obj": row[1],
            "Description2": row[2],
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

   
    
    current_date = datetime.today().date()
    
    current_year = current_date.year
    next_year = current_date.year + 1
    last_year = current_date.year - 1
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_name = last_month.strftime("%B")
    formatted_last_month = last_month.strftime('%B %d, %Y')
    last_month_number = last_month.month
    if school == 'manara' or school == 'prepschool':
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
    if school == "village-tech":
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
                )
            item[f"total_check{i}"] = total_real + total_adjustment


    july_date  = datetime(current_year, 7, 1).date()
    september_date  = datetime(current_year, 9, 1).date()
    FY_year_1 = last_year
    FY_year_2 = current_year
    for item in data3:
        date_str = item["Date"]
        if date_str:
            if school == 'manara' or school == 'prepschool':
               
                date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
                if date_obj > july_date: # if date is higher than july 1 this year
                  FY_year_1 = current_year
                  FY_year_2 = next_year
            else:
                date_obj = datetime.strptime(date_str, "%Y-%m-%d").date()
                if date_obj > september_date: # if date is higher than july 1 this year
                  FY_year_1 = current_year
                  FY_year_2 = next_year
                
           
            


    #checks if the last month column is empty. if empty. last month will be set to  last two months.
    if all(item[f"total_check{last_month_number}"] == 0 for item in data):
        last_2months = current_month - relativedelta(months=1)
        last_2months = last_2months - relativedelta(days=1)
        last_month_number = last_2months.month
        last_month_name = last_2months.strftime("%B")
        formatted_last_month = last_2months.strftime('%B %d, %Y')
        last_month_number = last_2months.month
        if school == 'manara' or school == 'prepschool':
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
        if school == "village-tech":
               
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


            
            if school == 'village-tech':
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
            if school == 'village-tech':
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
            item["var_ytd"] =  (abs(int(item['total_budget'] / item["ytd_total"]*100))) if item["ytd_total"] != 0 else ""
    
    ytd_ammended_total_first = first_total * ytd_budget
    var_ytd_first_total = (abs(int(first_ytd_total / ytd_ammended_total_first*100))) if ytd_ammended_total_first != 0 else ""


    for item in data2:
        if item["category"] == "Depreciation and Amortization":
            func = item["func_func"]
            obj = item["obj"]
            ytd_total = 0
     
           
            
            if school == 'village-tech':
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
                ytd_total += (item[f"total_func2_{month_number}"])
        
            item["ytd_total"] = ytd_total
            dna_total += item['total_budget']
            dna_ytd_total += item["ytd_total"]
            item[f"ytd_budget"] = item['total_budget'] * ytd_budget
            item["variances"] =  item[f"ytd_budget"] -item["ytd_total"]
            variances_dna+= item["variances"]
            item["var_ytd"] = (abs(int(item['total_budget'] / item["ytd_total"]*100))) if item["ytd_total"] != 0 else ""
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

        if school == 'village-tech':
            item["total_budget"] = sum(
                entry[appr_key]
                for entry in data3
                if entry["obj"] == obj

                )
        else:
            item["total_budget"] = sum(
            entry[appr_key]
            for entry in data3
            if entry["obj"] == obj
            and entry["Type"] == 'GJ'
       
            )
        
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
    variances_net_income = variances_revenue - variances_total_expense
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

            if school != 'manara' or school != 'prepschool':
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
                total_sum_value = totals.get(f"total_sum{last_month_number + 1}_value", 0)
                row["last_month_difference"] = row[f"difference_{last_month_number}"] + total_sum_value
     

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
                row["debt_fytd"] = ( total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value)

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
                total_sum_value = totals.get(f"total_sum{last_month_number}_value", 0)
                row["last_month_difference"] = row[f"difference_{last_month_number-1}"] + total_sum_value
                

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
  
                row["debt_fytd"] = ( total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value)


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
    json_path = os.path.join(JSON_DIR, "excel", school)
    if not os.path.exists(json_path):
        os.makedirs(json_path)

    for key, val in context.items():
        file = os.path.join(json_path, f"{key}.json")
        with open(file, "w") as f:
            json.dump(val, f)


if __name__ == "__main__":
    update_db()
