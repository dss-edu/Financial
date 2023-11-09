from .connect import connect
from time import strftime
from datetime import datetime, timedelta, date
from dateutil.relativedelta import relativedelta
from django.views.decorators.cache import cache_control
import json
import os
import re
import math
from config import settings

# Get the current date
current_date = datetime.now()
# Extract the month number from the current date
month_number = current_date.month
curr_year = current_date.year

SCHOOLS = settings.SCHOOLS
db = settings.db

def dashboard(school):
    # current_date = datetime.today().date()
    # # current_year = current_date.year
    # # last_year = current_date - timedelta(days=365)
    # current_month = current_date.replace(day=1)
    # last_month = current_month - relativedelta(days=1)

    # need to validate and sanitize school to avoid SQLi
    cnxn = connect()
    cursor = cnxn.cursor()
    for i in range(month_number - 1, 0, -1):
        query = f"SELECT * FROM [dbo].[AscenderData_CharterFirst] \
                    WHERE school = '{school}' \
                    AND month = {i};"
        cursor.execute(query)
        row = cursor.fetchone()
        if row is not None:
            last_month = date(curr_year, i + 1, 1)
            last_month = last_month - relativedelta(days=1)
            break

    context = {
        "school": school,
        "school_name": SCHOOLS[school],
        "date": last_month,
        "net_income_ytd": row[3],  ###
        "days_coh": row[6],  ###
        "net_earnings": row[8],  ###
        "debt_service": row[11],  ###
        "ratio_administrative": row[13],  ###
        ## "ratio_student_teacher": row[14],
    }

    cursor.close()
    cnxn.close()
    return context

def float_to_ratio(float_number):
    if float_number < 0:
        sign = '-'
        float_number = abs(float_number)
    else:
        sign = ''

    # Multiply the float by a power of 10 to make it an integer
    numerator = int(float_number * 100)
    denominator = 100

    # Find the greatest common divisor (GCD) and divide both numerator and denominator
    gcd = math.gcd(numerator, denominator)
    numerator //= gcd
    denominator //= gcd

    return f"{sign}{numerator}:{denominator}"

def percent_to_ratio(percentage):
    # Use regular expressions to extract the numeric part of the percentage string
    match = re.match(r"(\d+)%", percentage)

    if match:
        # Extract the numeric part and convert it to an integer
        percent = int(match.group(1))

        # Calculate the ratio
        numerator = percent
        denominator = 100
        # Find the greatest common divisor (GCD) and divide both numerator and denominator
        gcd = math.gcd(numerator, denominator)
        numerator //= gcd
        denominator //= gcd

        return f"{numerator}:{denominator}"
    else:
        return "Invalid input"


def charter_first(school):
    # need to validate and sanitize school to avoid SQLi
    cnxn = connect()
    cursor = cnxn.cursor()
    for i in range(month_number - 1, 0, -1):
        query = f"SELECT * FROM [dbo].[AscenderData_CharterFirst] \
                    WHERE school = '{school}' \
                    AND month = {i};"
        cursor.execute(query)
        row = cursor.fetchone()
        if row is not None:
            break

    context = {
        "school": school,
        "school_name": SCHOOLS[school],
        "year": row[1],
        "month": row[2],
        "net_income_ytd": row[3],  ###
        "indicators": row[4],
        "net_assets": row[5],
        "days_coh": row[6],  ###
        "current_assets": row[7],
        "net_earnings": row[8],  ###
        "budget_vs_revenue": row[9],
        "total_assets": row[10],
        "debt_service": row[11],  ###
        "debt_capitalization": row[12],
        "ratio_administrative": row[13],  ###
        "ratio_student_teacher": row[14],
        "estimated_actual_ada": row[15],
        "reporting_peims": row[16],
        "annual_audit": row[17],
        "post_financial_info": row[18],
        "approved_geo_boundaries": row[19],
        "estimated_first_rating": row[20],
    }

    net_ytd = context["net_income_ytd"]
    net_earnings = context["net_earnings"]

    if net_ytd < 0:
        context["net_income_ytd"] = f"$({net_ytd * -1:.0f})"
    else:
        context["net_income_ytd"] = f"${net_ytd:.0f}"

    if net_earnings < 0:
        context["net_earnings"] = f"$({net_earnings * -1:.0f})"
    else:
        context["net_earnings"] = f"${net_earnings:.0f}"

    context["debt_capitalization"] = f"{context['debt_capitalization']:.0f}%"

    # turn int into month name
    month = context["month"]
    year = context["year"]
    next_month = datetime(year, month + 1, 1)
    this_month = next_month - relativedelta(days=1)
    context["date"] = this_month

    # for FY
    fiscal_year = year
    if school in ["advantage", "cumberland", "village-tech"]:
        if month < 9:
            fiscal_year = year - 1

    if school in ["manara", "leadership"]:
        if month < 7:
            fiscal_year = year - 1

    context["fiscal_year"] = fiscal_year
    context["next_fiscal_year"] = fiscal_year + 1
    return context


def profit_loss(school, anchor_year):
    current_date = datetime.today().date()
    # current_year = current_date.year
    # last_year = current_date - timedelta(days=365)
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    formatted_last_month = last_month.strftime("%B %d, %Y")
    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 4
    ytd_budget = ytd_budget_test / 12
    formatted_ytd_budget = (
        f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
    )
    context = {
        "school": school,
        "school_name": SCHOOLS[school],
        # "last_month": formatted_last_month,
        # "last_month_number": last_month_number,
        # "format_ytd_budget": formatted_ytd_budget,
        # "ytd_budget": ytd_budget,
    }

    # JSON_DIR = os.path.join(BASE_DIR, "finance", "json", "profit-loss", school)
    JSON_DIR = os.path.join(settings.BASE_DIR, "finance", "json", "profit-loss", school)
    if anchor_year:  # anchor_year is by default = ""
        JSON_DIR = os.path.join(
            settings.BASE_DIR, "finance","json", str(anchor_year), "profit-loss", school
        )
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
        context["anchor_year"] = anchor_year

    return context


def profit_loss_chart(school, anchor_year):
    cnxn = connect()
    cursor = cnxn.cursor()
    for i in range(month_number - 1, 0, -1):
        query = f"SELECT * FROM [dbo].[AscenderData_CharterFirst] \
                    WHERE school = '{school}' \
                    AND month = {i};"
        cursor.execute(query)
        row = cursor.fetchone()
        if row is not None:
            break

    context = {
        "school": school,
        "school_name": SCHOOLS[school],
        "year": row[1],
        "month": row[2],
        "net_income_ytd": row[3],  ###
        "indicators": row[4],
        "net_assets": row[5],
        "days_coh": row[6],  ###
        "current_assets": row[7],
        "net_earnings": row[8],  ###
        "budget_vs_revenue": row[9],
        "total_assets": row[10],
        "debt_service": row[11],  ###
        "debt_capitalization": row[12],
        "ratio_administrative": row[13],  ###
        "ratio_student_teacher": row[14],
        "estimated_actual_ada": row[15],
        "reporting_peims": row[16],
        "annual_audit": row[17],
        "post_financial_info": row[18],
        "approved_geo_boundaries": row[19],
        "estimated_first_rating": row[20],
    }

    # JSON_DIR = os.path.join(BASE_DIR, "finance", "json", "profit-loss", school)
    JSON_DIR = os.path.join(settings.BASE_DIR, "finance", "json", "profit-loss-chart", school)

    if anchor_year:  # anchor_year is by default = ""
        JSON_DIR = os.path.join(
            settings.BASE_DIR, "finance", str(anchor_year), "profit-loss", school
        )
    files = os.listdir(JSON_DIR)

    for file in files:
        with open(os.path.join(JSON_DIR, file), "r") as f:
            basename = os.path.splitext(file)[0]
            context[basename] = json.load(f)

    return context


def balance_sheet(school, anchor_year):
    current_date = datetime.today().date()
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_name = last_month.strftime("%B")
    formatted_last_month = last_month.strftime("%B %d, %Y")
    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 4
    ytd_budget = ytd_budget_test / 12
    formatted_ytd_budget = (
        f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
    )
    context = {
        "school": school,
        "school_name": SCHOOLS[school],
        # "last_month": formatted_last_month,
        # "last_month_number": last_month_number,
        # "last_month_name": last_month_name,
        # "format_ytd_budget": formatted_ytd_budget,
        # "ytd_budget": ytd_budget,
    }

    if formatted_ytd_budget.startswith("0."):
        formatted_ytd_budget = formatted_ytd_budget[2:]

    JSON_DIR = os.path.join(settings.BASE_DIR, "finance", "json", "balance-sheet", school)
    if anchor_year:
        JSON_DIR = os.path.join(
            settings.BASE_DIR, "finance", "json", str(anchor_year), "balance-sheet", school
        )
    files = os.listdir(JSON_DIR)

    for file in files:
        with open(os.path.join(JSON_DIR, file), "r") as f:
            basename = os.path.splitext(file)[0]
            context[basename] = json.load(f)

    context["anchor_year"] = anchor_year

    cnxn = connect()
    cursor = cnxn.cursor()

    query = f"SELECT * FROM [dbo].{db[school]['bs_activity']} WHERE Activity IS NULL or Activity = '';"

    rows = cursor.execute(query)
    missing_act_list = []
    for row in rows:
        if row[3] == school:
            data = {"Activity": row[0], "obj": row[1], "Description": row[2]}
            missing_act_list.append(data)

    missing_act_list = sorted(missing_act_list, key=lambda x: x["obj"])
    context["missing_activities"] = missing_act_list

    query_notmissing = f"SELECT * FROM [dbo].{db[school]['bs_activity']} WHERE Activity IS NOT NULL or Activity != '';"

    rows = cursor.execute(query_notmissing)
    not_missing = []
    for row in rows:
        if row[3] == school:
            data = {"Activity": row[0], "obj": row[1], "Description": row[2]}
            not_missing.append(data)

    not_missing = sorted(not_missing, key=lambda x: x["obj"])
    context["not_missing"] = not_missing

    query = f"SELECT DISTINCT Activity FROM [dbo].{db[school]['bs_activity']}"
    opts = cursor.execute(query)

    options = []
    for opt in opts:
        if opt[0] not in [None, ""]:
            options.append(opt[0])

    context["activity_options"] = options

    return context


def cashflow(school, anchor_year):
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
        "anchor_year": anchor_year,
    }

    # all  of profit loss
    JSON_DIR = os.path.join(settings.BASE_DIR, "finance", "json")
    if anchor_year:
        JSON_DIR = os.path.join(settings.BASE_DIR, "finance", str(anchor_year))
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
    query = f"SELECT  TOP(500)* FROM [dbo].{db[school]['db']} ORDER BY Date DESC"
    if school == "village-tech":
        query = (
            f"SELECT  TOP(500)* FROM [dbo].{db[school]['db']} ORDER BY PostingDate DESC"
        )
    cursor.execute(query)
    rows = cursor.fetchall()

    data3 = []
    if school != "village-tech":
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
                # "Contr": row[21],
            }

            data3.append(row_dict)
    else:
        for row in rows:
            date_str = row[11]
            # if date_str is not None:
            #         date_without_time = date_str.strftime('%b. %d, %Y')
            # else:
            #         date_without_time = None
            row_dict = {
                "fund": row[0],
                "T": row[1],
                "func": row[2],
                "obj": row[3],
                "sobj": row[4],
                "org": row[5],
                "fscl_yr": row[6],
                "PI": row[7],
                "LOC": row[8],
                "PostingDate": row[9],
                "Month": row[10],
                "Source": date_str,
                "SubSource": row[12],
                "Batch": row[13],
                "Vendor": row[14],
                "TransactionDescr": row[15],
                "InvoiceDate": row[16],
                "CheckNumber": row[17],
                "CheckDate": row[18],
                "Amount": row[19],
                "Budget": row[20],
                # "Contr": row[21],
            }

            data3.append(row_dict)

    context = {
        "data3": data3,
    }
    context["school"] = school
    context["school_name"] = SCHOOLS[school]
    return context


def manual_adjustments(school):
    cnxn = connect()
    cursor = cnxn.cursor()
    query = "SELECT * FROM [dbo].[Adjustment] WHERE School = ? ;"
    # AND month = {month_number - 1};"
    cursor.execute(query, school)
    # cursor.execute(query)
    rows = cursor.fetchall()

    data3 = []
    for row in rows:
        date_str = row[11]
        # if date_str is not None:
        #     date_without_time = date_str.strftime("%b. %d, %Y")
        # else:
        #     date_without_time = None
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
            # "Contr": row[21],
        }

        data3.append(row_dict)

    query = f"SELECT * FROM [dbo].{db[school]['db']};"
    cursor.execute(query)
    rows = cursor.fetchall()
    options = {
        "fund": [],
        "func": [],
        "obj": [],
        "sobj": [],
        "org": [],
        "fscl_yr": [],
        "pgm": [],
        "edSpan": [],
        "projDtl": [],
        "AcctDescr": [],
        "Number": [],
        # "Date": date_str,""
        "AcctPer": [
            "00",
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
        ],
        # "Est": row[13],
        # "Real": row[14],
        # "Appr": row[15],
        # "Encum": row[16],
        # "Expend": row[17],
        # "Bal": row[18],
        "WorkDescr": [],
        "Type": [],
        # "Contr": row[21],
    }

    for row in rows:
        if row[0] not in options["fund"]:
            options["fund"].append(row[0])
        if row[1] not in options["func"]:
            options["func"].append(row[1])
        if row[2] not in options["obj"]:
            options["obj"].append(row[2])
        if row[3] not in options["sobj"]:
            options["sobj"].append(row[3])
        if row[4] not in options["org"]:
            options["org"].append(row[4])
        if row[5] not in options["fscl_yr"]:
            options["fscl_yr"].append(row[5])
        if row[6] not in options["pgm"]:
            options["pgm"].append(row[6])
        if row[7] not in options["edSpan"]:
            options["edSpan"].append(row[7])
        if row[8] not in options["projDtl"]:
            options["projDtl"].append(row[8])
        if row[9] not in options["AcctDescr"]:
            options["AcctDescr"].append(row[9])
        if row[10] not in options["Number"]:
            options["Number"].append(row[10])
        if row[19] not in options["WorkDescr"]:
            options["WorkDescr"].append(row[19])
        if row[20] not in options["Type"]:
            options["Type"].append(row[20])

    cursor.close()
    cnxn.close()

    context = {
        "school": school,
        "school_name": SCHOOLS[school],
        "data3": data3,
        "options": options,
    }
    return context


def activity_edits(school, body):
    cnxn = connect()
    cursor = cnxn.cursor()
    # delete the empties first

    ins_query = f"UPDATE [dbo].{db[school]['bs_activity']} SET Activity = ? , Description = ? where obj = ? and school = ?"
    for item in body:
        item["school"] = school
        values = (item["activity"], item["description"], item["obj"], item["school"])
        print(values)
        cursor.execute(ins_query, values)
        cnxn.commit()

    cursor.close()
    cnxn.close()
    return True
