from .connect import connect
from time import strftime
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from django.views.decorators.cache import cache_control
import json
import os

# Get the current date
current_date = datetime.now()
# Extract the month number from the current date
month_number = current_date.month
curr_year = current_date.year

SCHOOLS = {
    "advantage": "ADVANTAGE ACADEMY",
    "cumberland": "CUMBERLAND ACADEMY",
    "village-tech": "VILLAGE TECH",
    "prepschool": "LEADERSHIP PREP SCHOOL",
    "manara": "MANARA ACADEMY",
}

db = {
    "advantage": {
        "object": "[AscenderData_Advantage_Definition_obj]",
        "function": "[AscenderData_Advantage_Definition_func]",
        "db": "[AscenderData_Advantage]",
        "code": "[AscenderData_Advantage_PL_ExpensesbyObjectCode]",
        "activities": "[AscenderData_Advantage_PL_Activities]",
        "bs": "[AscenderData_Advantage_Balancesheet]",
        "bs_activity": "[AscenderData_Advantage_ActivityBS]",
        "cashflow": "[AscenderData_Advantage_Cashflow]",
        "adjustment": "[Adjustment]",
    },
    "cumberland": {
        "object": "[AscenderData_Cumberland_Definition_obj]",
        "function": "[AscenderData_Cumberland_Definition_func]",
        # "function": "[AscenderData_Advantage_Definition_func]",
        "db": "[AscenderData_Cumberland]",
        "code": "[AscenderData_Cumberland_PL_ExpensesbyObjectCode]",
        "activities": "[AscenderData_Cumberland_PL_Activities]",
        "bs": "[AscenderData_Cumberland_Balancesheet]",
        "bs_activity": "[AscenderData_Cumberland_ActivityBS]",
        "cashflow": "[AscenderData_Advantage_Cashflow]",
        "adjustment": "[Adjustment]",
    },
    "village-tech": {
        "object": "[AscenderData_Advantage_Definition_obj]",
        "function": "[AscenderData_Advantage_Definition_func]",
        "db": "[Skyward_VillageTech]",
        "code": "[AscenderData_Advantage_PL_ExpensesbyObjectCode]",
        "activities": "[AscenderData_Advantage_PL_Activities]",
        "bs": "[AscenderData_Advantage_Balancesheet]",
        "bs_activity": "[AscenderData_Advantage_ActivityBS]",
        "cashflow": "[AscenderData_Advantage_Cashflow]",
        "adjustment": "[Adjustment]",
    },
    "prepschool": {
        "object": "[AscenderData_Advantage_Definition_obj]",
        "function": "[AscenderData_Advantage_Definition_func]",
        "db": "[AscenderData_Leadership]",
        "code": "[AscenderData_Advantage_PL_ExpensesbyObjectCode]",
        "activities": "[AscenderData_Advantage_PL_Activities]",
        "bs": "[AscenderData_Advantage_Balancesheet]",
        "bs_activity": "[AscenderData_Advantage_ActivityBS]",
        "cashflow": "[AscenderData_Advantage_Cashflow]",
        "adjustment": "[Adjustment]",
    },
    "manara": {
        "object": "[AscenderData_Advantage_Definition_obj]",
        "function": "[AscenderData_Advantage_Definition_func]",
        "db": "[AscenderData_Manara]",
        "code": "[AscenderData_Advantage_PL_ExpensesbyObjectCode]",
        "activities": "[AscenderData_Advantage_PL_Activities]",
        "bs": "[AscenderData_Advantage_Balancesheet]",
        "bs_activity": "[AscenderData_Advantage_ActivityBS]",
        "cashflow": "[AscenderData_Advantage_Cashflow]",
        "adjustment": "[Adjustment]",
    },
}


def charter_first(school):
    # need to validate and sanitize school to avoid SQLi
    cnxn = connect()
    cursor = cnxn.cursor()
    query = f"SELECT * FROM [dbo].[AscenderData_CharterFirst] \
                WHERE school = '{school}' \
                AND month = {month_number - 1};"
    cursor.execute(query)
    row = cursor.fetchone()

    context = {
        "school": school,
        "school_name": SCHOOLS[school],
        "year": row[1],
        "month": row[2],
        "net_income_ytd": row[3],
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
    return context


def profit_loss(school):
    current_date = datetime.today().date()
    # current_year = current_date.year
    # last_year = current_date - timedelta(days=365)
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 4
    ytd_budget = ytd_budget_test / 12
    formatted_ytd_budget = (
        f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
    )

    context = {
        "school": school,
        "school_name": SCHOOLS[school],
        "last_month": last_month,
        "last_month_number": last_month_number,
        "format_ytd_budget": formatted_ytd_budget,
        "ytd_budget": ytd_budget,
    }

    BASE_DIR = os.getcwd()
    JSON_DIR = os.path.join(BASE_DIR, "finance", "json", "profit-loss", school)
    files = os.listdir(JSON_DIR)

    for file in files:
        with open(os.path.join(JSON_DIR, file), "r") as f:
            basename = os.path.splitext(file)[0]
            context[basename] = json.load(f)

    if not school == "village-tech":
        lr_funds = list(set(row["fund"] for row in context["data3"] if "fund" in row))
        lr_funds_sorted = sorted(lr_funds)
        lr_obj = list(set(row["obj"] for row in context["data3"] if "obj" in row))
        lr_obj_sorted = sorted(lr_obj)

        func_choice = list(
            set(row["func"] for row in context["data3"] if "func" in row)
        )
        func_choice_sorted = sorted(func_choice)

        context["lr_funds"] = lr_funds_sorted
        context["lr_obj"] = lr_obj_sorted
        context["func_choice"] = func_choice_sorted
        
    return context


def balance_sheet(school):
    

    current_date = datetime.today().date()
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_name = last_month.strftime("%B")

    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 4
    ytd_budget = ytd_budget_test / 12
    formatted_ytd_budget = (
        f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
    )

    if formatted_ytd_budget.startswith("0."):
        formatted_ytd_budget = formatted_ytd_budget[2:]
    context = {
            "school": school,
            "school_name": SCHOOLS[school],
            
            "last_month": last_month,
            "last_month_number": last_month_number,
            "last_month_name": last_month_name,
            "format_ytd_budget": formatted_ytd_budget,
            "ytd_budget": ytd_budget,
            }

    BASE_DIR = os.getcwd()
    JSON_DIR = os.path.join(BASE_DIR, "finance", "json", "balance-sheet", school)
    files = os.listdir(JSON_DIR)

    for file in files:
        with open(os.path.join(JSON_DIR, file), "r") as f:
            basename = os.path.splitext(file)[0]
            context[basename] = json.load(f)

    return context



def cashflow(school):
    current_date = datetime.today().date()
    current_year = current_date.year
    last_year = current_date - timedelta(days=365)
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 4
    ytd_budget = ytd_budget_test / 12
    formatted_ytd_budget = f"{ytd_budget:.2f}"

    if formatted_ytd_budget.startswith("0."):
        formatted_ytd_budget = formatted_ytd_budget[2:]

    context = {
        "school": school,
        "school_name": SCHOOLS[school],
        # "data_cashflow": data_cashflow,
        # "data_activitybs": data_activitybs,
        # "data_balancesheet": data_balancesheet,
        "last_month": last_month,
        "last_month_number": last_month_number,
        "format_ytd_budget": formatted_ytd_budget,
        "ytd_budget": ytd_budget,
    }

    # all  of profit loss
    JSON_DIR = os.path.join(os.getcwd(), "finance", "json")
    PL_DIR = os.path.join(JSON_DIR, "profit-loss", school)
    files = os.listdir(PL_DIR)

    for file in files:
        with open(os.path.join(PL_DIR, file), "r") as f:
            basename = os.path.splitext(file)[0]
            context[basename] = json.load(f)

    if not school == "village-tech":
        lr_funds = list(set(row["fund"] for row in context["data3"] if "fund" in row))
        lr_funds_sorted = sorted(lr_funds)
        lr_obj = list(set(row["obj"] for row in context["data3"] if "obj" in row))
        lr_obj_sorted = sorted(lr_obj)

        func_choice = list(
            set(row["func"] for row in context["data3"] if "func" in row)
        )
        func_choice_sorted = sorted(func_choice)

        context["lr_funds"] = lr_funds_sorted
        context["lr_obj"] = lr_obj_sorted
        context["func_choice"] = func_choice_sorted

    BS_DIR = os.path.join(JSON_DIR, "balance-sheet", school)
    files = os.listdir(BS_DIR)

    for file in files:
        with open(os.path.join(BS_DIR, file), "r") as f:
            basename = os.path.splitext(file)[0]
            context[basename] = json.load(f)

    CF_DIR = os.path.join(JSON_DIR, "cashflow", school)
    files = os.listdir(CF_DIR)

    for file in files:
        with open(os.path.join(CF_DIR, file), "r") as f:
            basename = os.path.splitext(file)[0]
            context[basename] = json.load(f)

    return context


def general_ledger(school):
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute(f"SELECT  TOP(500)* FROM [dbo].{db[school]['db']} ORDER BY Date DESC")
    rows = cursor.fetchall()

    data3 = []
    for row in rows:
        date_str = row[11]
        # if date_str is not None:
        #         date_without_time = date_str.strftime('%b. %d, %Y')
        # else:
        #         date_without_time = None
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
            "Date": date_str,
            "AcctPer": row[12],
            "Est": row[13],
            "Real": row[14],
            "Appr": row[15],
            "Encum": row[16],
            "Expend": row[17],
            "Bal": row[18],
            "WorkDescr": row[19],
            "Type": row[20],
            "Contr": row[21],
        }

        data3.append(row_dict)

    context = {
        "data3": data3,
    }
    context["school"] = school
    context["school_name"] = SCHOOLS[school]
    return context
