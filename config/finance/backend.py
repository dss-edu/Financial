from connect import connect
from time import strftime
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import json
import pprint
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


def update_db():
    for school, name in SCHOOLS.items():
        profit_loss(school)


def profit_loss(school):
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['object']};")
    rows = cursor.fetchall()

    data = []
    for row in rows:
        if row[4] is None:
            row[4] = ""
        valueformat = "{:,.0f}".format(float(row[4])) if row[4] else ""
        row_dict = {
            "fund": row[0],
            "obj": row[1],
            "description": row[2],
            "category": row[3],
            "value": valueformat,
        }
        data.append(row_dict)

    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['function']};")

    rows = cursor.fetchall()

    data2 = []
    for row in rows:
        budgetformat = "{:,.0f}".format(float(row[3])) if row[3] else ""
        row_dict = {
            "func_func": row[0],
            "desc": row[1],
            "category": row[2],
            "obj": row[4],
            "budget": budgetformat,
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
            }

            data3.append(row_dict)

    cursor.execute(f"SELECT * FROM [dbo].{db[school]['adjustment']} ")
    rows = cursor.fetchall()

    adjustment = []

    if not school == "village-tech":
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
        row_dict = {
            "obj": row[0],
            "Description": row[1],
            "Category": row[2],
        }

        data_activities.append(row_dict)

    # ---------- FOR EXPENSE TOTAL -------
    expense_key = "Expend"
    if school == "village-tech":
        expense_key = "Amount"

    acct_per_values_expense = [
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
    for item in data_activities:
        obj = item["obj"]

        for i, acct_per in enumerate(acct_per_values_expense, start=1):
            item[f"total_activities{i}"] = sum(
                entry[expense_key]
                for entry in data3
                if entry["obj"] == obj and entry["AcctPer"] == acct_per
            )
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

    for item in data_expensebyobject:
        obj = item["obj"]
        if obj == "6100":
            category = "Payroll Costs"
        elif obj == "6200":
            category = "Professional and Cont Svcs"
        elif obj == "6300":
            category = "Supplies and Materials"
        elif obj == "6400":
            category = "Other Operating Expenses"
        else:
            category = "Total Expense"

        for i, acct_per in enumerate(acct_per_values_expense, start=1):
            item[f"total_expense{i}"] = sum(
                entry[f"total_activities{i}"]
                for entry in data_activities
                if entry["Category"] == category
            )

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

    # ---- for data ------
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

    data_key = "Real"
    if school == "village-tech":
        data_key = "Amount"

    for item in data:
        fund = item["fund"]
        obj = item["obj"]

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f"total_real{i}"] = sum(
                entry[data_key]
                for entry in data3
                if entry["fund"] == fund
                and entry["obj"] == obj
                and entry["AcctPer"] == acct_per
            )

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

    for row in data:
        for key in keys_to_check:
            if row[key] < 0:
                row[key] = -row[key]
            else:
                row[key] = ""

    for row in data:
        for key in keys_to_check:
            if row[key] != "":
                row[key] = "{:,.0f}".format(row[key])

    acct_per_values2 = [
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

    data_key = "Expend"
    if school == "village-tech":
        data_key = "Amount"
    for item in data2:
        if item["category"] != "Depreciation and Amortization":
            func = item["func_func"]
            for i, acct_per in enumerate(acct_per_values2, start=1):
                total_func = sum(
                    entry[data_key]
                    for entry in data3
                    if entry["func"] == func and entry["AcctPer"] == acct_per
                )
                total_adjustment = sum(
                    entry[data_key]
                    for entry in adjustment
                    if entry["func"] == func and entry["AcctPer"] == acct_per
                )
                item[f"total_func{i}"] = total_func + total_adjustment

    for item in data2:
        if item["category"] == "Depreciation and Amortization":
            func = item["func_func"]
            obj = item["obj"]
            for i, acct_per in enumerate(acct_per_values2, start=1):
                total_func = sum(
                    entry[data_key]
                    for entry in data3
                    if entry["func"] == func
                    and entry["AcctPer"] == acct_per
                    and entry["obj"] == obj
                )
                total_adjustment = sum(
                    entry[data_key]
                    for entry in adjustment
                    if entry["func"] == func
                    and entry["AcctPer"] == acct_per
                    and entry["obj"] == obj
                )
                item[f"total_func2_{i}"] = total_func + total_adjustment

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
    # formatted_ytd_budget = (
    #     f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
    # )
    #
    # if formatted_ytd_budget.startswith("0."):
    #     formatted_ytd_budget = formatted_ytd_budget[2:]

    context = {
        "data": data,
        "data2": data2,
        "data3": data3,
        "data_expensebyobject": data_expensebyobject,
        "data_activities": data_activities,
        # "last_month": last_month,
        # "last_month_number": last_month_number,
        # "format_ytd_budget": formatted_ytd_budget,
        # "ytd_budget": ytd_budget,
    }

    # if not school == "village-tech":
    #     context["lr_funds"] = lr_funds_sorted
    #     context["lr_obj"] = lr_obj_sorted
    #     context["func_choice"] = func_choice_sorted

    dict_keys = ["data", "data2", "data3", "data_expensebyobject", "data_activities"]
    JSON_DIR = os.path.join(os.getcwd(), "finance", "json", "profit-loss", school)

    if not os.path.exists(JSON_DIR):
        os.makedirs(JSON_DIR)

    for key in dict_keys:
        if key in context.keys():
            json_path = os.path.join(JSON_DIR, f"{key}.json")
            with open(json_path, "w") as f:
                json.dump(context[key], f)


if __name__ == "__main__":
    update_db()
