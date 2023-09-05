from .connect import connect
from time import strftime
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta

SCHOOLS = {
    "advantage": "ADVANTAGE ACADEMY",
    "cumberland": "CUMBERLAND ACADEMY",
    "village-tech": "VILLAGE TECH",
    "prepschool": "LEADERSHIP PREP SCHOOL",
    "manara": "MANARA ACADEMY"
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
        "cashflow" : "[AscenderData_Advantage_Cashflow]",
    },
    "cumberland": {
        "object": "[AscenderData_Cumberland_Definition_obj]", 
        #"function": "[AscenderData_Cumberland_Definition_func]",
        "function": "[AscenderData_Advantage_Definition_func]",
        "db": "[AscenderData_Cumberland]",
        "code": "[AscenderData_Cumberland_PL_ExpensesbyObjectCode]",
        "activities": "[AscenderData_Cumberland_PL_Activities]",
        "bs": "[AscenderData_Cumberland_Balancesheet]",
        "bs_activity": "[AscenderData_Cumberland_ActivityBS]",
        "cashflow" : "[AscenderData_Advantage_Cashflow]",
    },
    "village-tech": {
        "object": "[AscenderData_Advantage_Definition_obj]", 
        "function": "[AscenderData_Advantage_Definition_func]",
        "db": "[Skyward_VillageTech]",
        "code": "[AscenderData_Advantage_PL_ExpensesbyObjectCode]",
        "activities": "[AscenderData_Advantage_PL_Activities]",
        "bs": "[AscenderData_Advantage_Balancesheet]",
        "bs_activity": "[AscenderData_Advantage_ActivityBS]",
        "cashflow" : "[AscenderData_Advantage_Cashflow]",
    },
    "prepschool": {
        "object": "[AscenderData_Advantage_Definition_obj]", 
        "function": "[AscenderData_Advantage_Definition_func]",
        "db": "[AscenderData_Leadership]",
        "code": "[AscenderData_Advantage_PL_ExpensesbyObjectCode]",
        "activities": "[AscenderData_Advantage_PL_Activities]",
        "bs": "[AscenderData_Advantage_Balancesheet]",
        "bs_activity": "[AscenderData_Advantage_ActivityBS]",
        "cashflow" : "[AscenderData_Advantage_Cashflow]",
    },
    "manara": {
        "object": "[AscenderData_Advantage_Definition_obj]", 
        "function": "[AscenderData_Advantage_Definition_func]",
        "db": "[AscenderData_Manara]",
        "code": "[AscenderData_Advantage_PL_ExpensesbyObjectCode]",
        "activities": "[AscenderData_Advantage_PL_Activities]",
        "bs": "[AscenderData_Advantage_Balancesheet]",
        "bs_activity": "[AscenderData_Advantage_ActivityBS]",
        "cashflow" : "[AscenderData_Advantage_Cashflow]",
        },
}


def profit_loss(school):
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['object']};") 
    rows = cursor.fetchall()
    
    data = []
    for row in rows:
        if row[4] is None:
            row[4] = ''
        valueformat = "{:,.0f}".format(float(row[4])) if row[4] else ""
        row_dict = {
            'fund': row[0],
            'obj': row[1],
            'description': row[2],
            'category': row[3],
            'value': valueformat  
        }
        data.append(row_dict)

    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['function']};") 

    rows = cursor.fetchall()


    data2=[]
    for row in rows:
        budgetformat = "{:,.0f}".format(float(row[3])) if row[3] else ""
        row_dict = {
            'func_func': row[0],
            'desc': row[1],
            'category': row[2],
            'obj': row[4],
            'budget': budgetformat,
            
        }
        data2.append(row_dict)



    #
    cursor.execute(f"SELECT * FROM [dbo].{db[school]['db']};") 
    rows = cursor.fetchall()
    
    data3=[]
    
    
    if not school == "village-tech":
        for row in rows:
            expend = float(row[17])

            row_dict = {
                'fund':row[0],
                'func':row[1],
                'obj':row[2],
                'sobj':row[3],
                'org':row[4],
                'fscl_yr':row[5],
                'pgm':row[6],
                'edSpan':row[7],
                'projDtl':row[8],
                'AcctDescr':row[9],
                'Number':row[10],
                'Date':row[11],
                'AcctPer':row[12],
                'Est':row[13],
                'Real':row[14],
                'Appr':row[15],
                'Encum':row[16],
                'Expend':expend,
                'Bal':row[18],
                'WorkDescr':row[19],
                'Type':row[20],
                'Contr':row[21]
                }

            data3.append(row_dict)
            
    else:
        for row in rows:
            amount = float(row[19])
            row_dict = {
                'fund':row[0],
                'func':row[2],
                'obj':row[3],
                'sobj':row[4],
                'org':row[5],
                'fscl_yr':row[6],
     
                'Date':row[9],
                'AcctPer':row[10],
                'Amount':amount,
            }

            data3.append(row_dict)


    cursor.execute(f"SELECT * FROM [dbo].{db[school]['code']};") 
    rows = cursor.fetchall()
    
    data_expensebyobject=[]
    
    
    for row in rows:
        
        budgetformat = "{:,.0f}".format(float(row[2])) if row[2] else ""
        row_dict = {
            'obj':row[0],
            'Description':row[1],
            'budget':budgetformat,
            
            }
        
        data_expensebyobject.append(row_dict)

    cursor.execute(f"SELECT * FROM [dbo].{db[school]['activities']};") 
    rows = cursor.fetchall()
    
    data_activities=[]
    
    
    for row in rows:
        
      
        row_dict = {
            'obj':row[0],
            'Description':row[1],
            'Category':row[2],
            
            }
        
        data_activities.append(row_dict)
    

    #---------- FOR EXPENSE TOTAL -------
    expense_key = 'Expend'
    if school  == "village-tech":
        expense_key = 'Amount'

    acct_per_values_expense = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
    for item in data_activities:
        
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values_expense, start=1):
            item[f'total_activities{i}'] = sum(
                entry[expense_key] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
            )
    keys_to_check_expense = ['total_activities1', 'total_activities2', 'total_activities3', 'total_activities4', 'total_activities5','total_activities6','total_activities7','total_activities8','total_activities9','total_activities10','total_activities11','total_activities12']
    keys_to_check_expense2 = ['total_expense1', 'total_expense2', 'total_expense3', 'total_expense4', 'total_expense5','total_expense6','total_expense7','total_expense8','total_expense9','total_expense10','total_expense11','total_expense12']



    for item in data_expensebyobject:
        obj = item['obj']
        if obj == '6100':
            category = 'Payroll Costs'
        elif obj == '6200':
            category = 'Professional and Cont Svcs'
        elif obj == '6300':
            category = 'Supplies and Materials'
        elif obj == '6400':
            category = 'Other Operating Expenses'
        else:
            category = 'Total Expense'

        for i, acct_per in enumerate(acct_per_values_expense, start=1):
            item[f'total_expense{i}'] = sum(
                entry[f'total_activities{i}'] for entry in data_activities if entry['Category'] == category 
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
        
    #---- for data ------
    acct_per_values = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']

    data_key = 'Real'
    if school == "village-tech":
        data_key = 'Amount'

    for item in data:
        fund = item['fund']
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_real{i}'] = sum(
                entry[data_key] for entry in data3 if entry['fund'] == fund and entry['obj'] == obj and entry['AcctPer'] == acct_per
            )

    keys_to_check = ['total_real1', 'total_real2', 'total_real3', 'total_real4', 'total_real5','total_real6','total_real7','total_real8','total_real9','total_real10','total_real11','total_real12']
 
    for row in data:
        for key in keys_to_check:
            if row[key] < 0:
                row[key] = -row[key]
            else:
                row[key] = ''

    for row in data:
        for key in keys_to_check:
            if row[key] != "":
                row[key] = "{:,.0f}".format(row[key])
                
    acct_per_values2 = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']

    data_key = 'Expend'
    if school == "village-tech":
        data_key = 'Amount'
    for item in data2:
        func = item['func_func']
        for i, acct_per in enumerate(acct_per_values2, start=1):
            item[f'total_func{i}'] = sum(
                entry[data_key] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per
            )
    for item in data2:
        func = item['func_func']
        for i, acct_per in enumerate(acct_per_values2, start=1):
            item[f'total_func2_{i}'] = sum(
                entry[data_key] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per and entry['obj'] == '6449'
            )  

    keys_to_check_func = ['total_func1', 'total_func2', 'total_func3', 'total_func4', 'total_func5','total_func6','total_func7','total_func8','total_func9','total_func10','total_func11','total_func12']
    keys_to_check_func_2 = ['total_func2_1', 'total_func2_2', 'total_func2_3', 'total_func2_4', 'total_func2_5','total_func2_6','total_func2_7','total_func2_8','total_func2_9','total_func2_10','total_func2_11','total_func2_12']

    for row in data2:
        for key in keys_to_check_func:
            if row[key] > 0:
                row[key] = row[key]
            else:
                row[key] = ''
    for row in data2:
        for key in keys_to_check_func:
            if row[key] != "":
                row[key] = "{:,.0f}".format(row[key])

    for row in data2:
        for key in keys_to_check_func_2:
            if row[key] > 0:
                row[key] = row[key]
            else:
                row[key] = ''
    for row in data2:
        for key in keys_to_check_func_2:
            if row[key] != "":
                row[key] = "{:,.0f}".format(row[key])

    if not school == "village-tech":
        lr_funds = list(set(row['fund'] for row in data3 if 'fund' in row))
        lr_funds_sorted = sorted(lr_funds)
        lr_obj = list(set(row['obj'] for row in data3 if 'obj' in row))
        lr_obj_sorted = sorted(lr_obj)

        func_choice = list(set(row['func'] for row in data3 if 'func' in row))
        func_choice_sorted = sorted(func_choice)
    
            
    current_date = datetime.today().date()
    current_year = current_date.year
    last_year = current_date - timedelta(days=365)
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 4
    ytd_budget = ytd_budget_test / 12
    formatted_ytd_budget = f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places

    if formatted_ytd_budget.startswith("0."):
        formatted_ytd_budget = formatted_ytd_budget[2:]
   
    if not school == "village-tech":
        context = {
             'data': data, 
             'data2':data2 , 
             'data3': data3 ,
              'lr_funds':lr_funds_sorted, 
              'lr_obj':lr_obj_sorted, 
              'func_choice':func_choice_sorted ,
              'data_expensebyobject': data_expensebyobject,
              'data_activities': data_activities,
              'last_month':last_month,
              'last_month_number':last_month_number,
              'format_ytd_budget': formatted_ytd_budget,
              'ytd_budget':ytd_budget,
              }
    else:
        context = {
             'data': data, 
             'data2':data2 , 
             'data3': data3 ,
              'data_expensebyobject': data_expensebyobject,
              'data_activities': data_activities,
              'last_month':last_month,
              'last_month_number':last_month_number,
              'format_ytd_budget': formatted_ytd_budget,
              'ytd_budget':ytd_budget,
              }

    context["school"] = school
    context["school_name"] = SCHOOLS[school]

    return context

def balance_sheet(school):
    cnxn = connect()
    cursor = cnxn.cursor()
    
    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['bs']}") 
    rows = cursor.fetchall()
    
    data_balancesheet=[]
    
    for row in rows:
        fye = int(row[4]) if row[4] else 0
        if fye == 0:
            fyeformat = ""
        else:
            fyeformat = "{:,.0f}".format(abs(fye)) if fye >= 0 else "({:,.0f})".format(abs(fye))
        row_dict = {
            'Activity':row[0],
            'Description':row[1],
            'Category':row[2],
            'Subcategory':row[3],
            'FYE':fyeformat,
            
            
            }
        
        data_balancesheet.append(row_dict)


    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['object']};") 
    rows = cursor.fetchall()

    
    data = []
    for row in rows:
        if row[4] is None:
            row[4] = ''
        valueformat = "{:,.0f}".format(float(row[4])) if row[4] else ""
        row_dict = {
            'fund': row[0],
            'obj': row[1],
            'description': row[2],
            'category': row[3],
            'value': valueformat  
        }
        data.append(row_dict)

    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['function']};") 
    rows = cursor.fetchall()


    data2=[]
    for row in rows:
        budgetformat = "{:,.0f}".format(float(row[3])) if row[3] else ""
        row_dict = {
            'func_func': row[0],
            'desc': row[1],
            'category': row[2],
            'obj': row[4],
            'budget': budgetformat,
            
        }
        data2.append(row_dict)

    cursor.execute(f"SELECT * FROM [dbo].{db[school]['bs_activity']}") 
    rows = cursor.fetchall()
    
    data_activitybs=[]
    
    
    for row in rows:

        row_dict = {
            'Activity':row[0],
            'obj':row[1],
            'Description2':row[2],
            
            
            }
        
        data_activitybs.append(row_dict)

    cursor.execute(f"SELECT * FROM [dbo].{db[school]['db']}") 
    rows = cursor.fetchall()
    
    data3=[]
    
    
    if not school == "village-tech":
        for row in rows:

            row_dict = {
                'fund':row[0],
                'func':row[1],
                'obj':row[2],
                'sobj':row[3],
                'org':row[4],
                'fscl_yr':row[5],
                'pgm':row[6],
                'edSpan':row[7],
                'projDtl':row[8],
                'AcctDescr':row[9],
                'Number':row[10],
                'Date':row[11],
                'AcctPer':row[12],
                'Est':row[13],
                'Real':row[14],
                'Appr':row[15],
                'Encum':row[16],
                'Expend':row[17],
                'Bal':row[18],
                'WorkDescr':row[19],
                'Type':row[20],
                'Contr':row[21]
                }
            
            data3.append(row_dict)
    else:
        for row in rows:
            amount = float(row[19])
            row_dict = {
                'fund':row[0],
                'func':row[2],
                'obj':row[3],
                'sobj':row[4],
                'org':row[5],
                'fscl_yr':row[6],
     
                'Date':row[9],
                'AcctPer':row[10],
                'Amount':amount,
                }
            data3.append(row_dict)

    acct_per_values = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']

    activity_key = 'Bal'
    if school == "village-tech":
        activity_key = 'Amount'

    for item in data_activitybs:
        
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_bal{i}'] = sum(
                entry[activity_key] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
            )

    keys_to_check = ['total_bal1', 'total_bal2', 'total_bal3', 'total_bal4', 'total_bal5','total_bal6','total_bal7','total_bal8','total_bal9','total_bal10','total_bal11','total_bal12']
    
    
                
    for row in data_activitybs:
        for key in keys_to_check:
            value = int(row[key])
            if value == 0:
                row[key] = ""
            elif value < 0:
                row[key] = "({:,.0f})".format(abs(float(row[key]))) 
            elif value != "":
                row[key] = "{:,.0f}".format(float(row[key]))
                


    activity_sum_dict = {} 
    for item in data_activitybs:
        Activity = item['Activity']
        for i in range(1, 13):
            total_sum_i = sum(
                int(entry[f'total_bal{i}'].replace(',', '').replace('(', '-').replace(')', '')) if entry[f'total_bal{i}'] and entry['Activity'] == Activity else 0
                for entry in data_activitybs
            )
            activity_sum_dict[(Activity, i)] = total_sum_i
    

 
    for row in data_balancesheet:
        
        activity = row['Activity']
        for i in range(1, 13):
            key = (activity, i)
            row[f'total_sum{i}'] = "{:,.0f}".format(activity_sum_dict.get(key, 0))

    # TOTAL REVENUE
    total_revenue = {acct_per: 0 for acct_per in acct_per_values}
    data_key2 = 'Real'
    if school == "village-tech":
        data_key2 = 'Amount'
        
    for item in data:
        fund = item['fund']
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_real{i}'] = sum(
                entry[data_key2] for entry in data3 if entry['fund'] == fund and entry['obj'] == obj and entry['AcctPer'] == acct_per
            )

            total_revenue[acct_per] += abs(item[f'total_real{i}'])
    
    #total surplus
    total_surplus = {acct_per: 0 for acct_per in acct_per_values}
    data_key = 'Expend'
    if school == "village-tech":
        data_key = 'Amount'
    for item in data2:
        if item['category'] != 'Depreciation and Amortization':
            func = item['func_func']

            for i, acct_per in enumerate(acct_per_values, start=1):
                item[f'total_func{i}'] = sum(
                    entry[data_key] for entry in data3 if entry['func'] == func and entry['AcctPer'] == acct_per
                )
                total_surplus[acct_per] += item[f'total_func{i}']

    # difference_func_values = {i: 0 for i in range(1, 13)}
    # monthly_totals_func = {i: 0 for i in range(1, 13)}
    # monthly_totals_func2 = {i: 0 for i in range(1, 13)}

    #---- Depreciation and ammortization total
    total_DnA = {acct_per: 0 for acct_per in acct_per_values}
    
    for item in data2:
        if item['category'] == 'Depreciation and Amortization':
            func = item['func_func']
            obj =  item['obj']

            for i, acct_per in enumerate(acct_per_values, start=1):
                item[f'total_func2_{i}'] = sum(
                    entry[data_key] for entry in data3 if entry['func'] == func and entry['AcctPer'] == acct_per and entry['obj'] == obj
                )
                total_DnA[acct_per] += item[f'total_func2_{i}']

    total_SBD = {acct_per: total_revenue[acct_per] - total_surplus[acct_per]  for acct_per in acct_per_values}
    total_netsurplus = {acct_per: total_SBD[acct_per] - total_DnA[acct_per]  for acct_per in acct_per_values}
    
    


    # for month, total in monthly_totals_func2.items():
    #     print(f'MonthFUNC2 {month}: {total}')
           
    
    # for key, value in difference_func_values.items():
    #     print(f'{key}: {value}')
    
    def format_with_parentheses(value):
        if value == 0:
            return ""
        formatted_value = "{:,.0f}".format(abs(value))
        return "({})".format(formatted_value) if value < 0 else formatted_value
    
    for row in data_balancesheet:
    
        FYE_value = float(row['FYE'].replace(',', '').replace('(', '-').replace(')', '')) if row['FYE'] else 0
        total_sum9_value = float(row['total_sum9'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum9'] else 0
        total_sum10_value = float(row['total_sum10'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum10'] else 0
        total_sum11_value = float(row['total_sum11'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum11'] else 0
        total_sum12_value = float(row['total_sum12'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum12'] else 0
        total_sum1_value = float(row['total_sum1'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum1'] else 0
        total_sum2_value = float(row['total_sum2'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum2'] else 0
        total_sum3_value = float(row['total_sum3'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum3'] else 0
        total_sum4_value = float(row['total_sum4'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum4'] else 0
        total_sum5_value = float(row['total_sum5'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum5'] else 0
        total_sum6_value = float(row['total_sum6'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum6'] else 0
        total_sum7_value = float(row['total_sum7'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum7'] else 0
        total_sum8_value = float(row['total_sum8'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum8'] else 0

        # Calculate the differences and store them in the row dictionary
        row['difference_9'] = format_with_parentheses(FYE_value + total_sum9_value)
        row['difference_10'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value)
        row['difference_11'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value)
        row['difference_12'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value)
        row['difference_1'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value)
        row['difference_2'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value)
        row['difference_3'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value)
        row['difference_4'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value)
        row['difference_5'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value)
        row['difference_6'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value)
        row['difference_7'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value)
        row['difference_8'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value)

        row['debt_9'] = format_with_parentheses(FYE_value  - total_sum9_value)
        row['debt_10'] = format_with_parentheses(FYE_value - total_sum9_value - total_sum10_value)
        row['debt_11'] = format_with_parentheses(FYE_value - total_sum9_value - total_sum10_value - total_sum11_value)
        row['debt_12'] = format_with_parentheses(FYE_value - total_sum9_value - total_sum10_value - total_sum11_value - total_sum12_value)
        row['debt_1'] = format_with_parentheses(FYE_value  - total_sum9_value - total_sum10_value - total_sum11_value - total_sum12_value - total_sum1_value)
        row['debt_2'] = format_with_parentheses(FYE_value  - total_sum9_value - total_sum10_value - total_sum11_value - total_sum12_value - total_sum1_value - total_sum2_value)
        row['debt_3'] = format_with_parentheses(FYE_value  - total_sum9_value - total_sum10_value - total_sum11_value - total_sum12_value - total_sum1_value - total_sum2_value - total_sum3_value)
        row['debt_4'] = format_with_parentheses(FYE_value  - total_sum9_value - total_sum10_value - total_sum11_value - total_sum12_value - total_sum1_value - total_sum2_value - total_sum3_value + total_sum4_value)
        row['debt_5'] = format_with_parentheses(FYE_value  - total_sum9_value - total_sum10_value - total_sum11_value - total_sum12_value - total_sum1_value - total_sum2_value - total_sum3_value + total_sum4_value - total_sum5_value)
        row['debt_6'] = format_with_parentheses(FYE_value  - total_sum9_value - total_sum10_value - total_sum11_value - total_sum12_value - total_sum1_value - total_sum2_value - total_sum3_value + total_sum4_value - total_sum5_value - total_sum6_value)
        row['debt_7'] = format_with_parentheses(FYE_value  - total_sum9_value - total_sum10_value - total_sum11_value - total_sum12_value - total_sum1_value - total_sum2_value - total_sum3_value + total_sum4_value - total_sum5_value - total_sum6_value - total_sum7_value)
        row['debt_8'] = format_with_parentheses(FYE_value  - total_sum9_value - total_sum10_value - total_sum11_value - total_sum12_value - total_sum1_value - total_sum2_value - total_sum3_value + total_sum4_value - total_sum5_value - total_sum6_value - total_sum7_value - total_sum8_value)
        
        row['fytd'] = format_with_parentheses(total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value)
        row['net_assets9'] = format_with_parentheses(FYE_value + total_netsurplus['09'])
        row['net_assets10'] =  format_with_parentheses(FYE_value + total_netsurplus['09'] +total_netsurplus['10'])
        row['net_assets11'] =  format_with_parentheses(FYE_value + total_netsurplus['09'] +total_netsurplus['10'] +total_netsurplus['11'])
        row['net_assets12'] =  format_with_parentheses(FYE_value + total_netsurplus['09'] +total_netsurplus['10'] +total_netsurplus['11'] +total_netsurplus['12'])
      
        row['net_assets1'] =  format_with_parentheses(FYE_value + total_netsurplus['09'] +total_netsurplus['10'] +total_netsurplus['11'] +total_netsurplus['12'] +total_netsurplus['01'])
        row['net_assets2'] =  format_with_parentheses(FYE_value + total_netsurplus['09']+total_netsurplus['10'] +total_netsurplus['11'] +total_netsurplus['12'] +total_netsurplus['01'] +total_netsurplus['02'])
        row['net_assets3'] =  format_with_parentheses(FYE_value + total_netsurplus['09']+total_netsurplus['10'] +total_netsurplus['11'] +total_netsurplus['12'] +total_netsurplus['01'] +total_netsurplus['02']+total_netsurplus['03'])
        row['net_assets4'] =  format_with_parentheses(FYE_value + total_netsurplus['09'] +total_netsurplus['10'] +total_netsurplus['11'] +total_netsurplus['12'] +total_netsurplus['01'] +total_netsurplus['02']+total_netsurplus['03']+total_netsurplus['04'])
        row['net_assets5'] =  format_with_parentheses(FYE_value + total_netsurplus['09'] +total_netsurplus['10'] +total_netsurplus['11'] +total_netsurplus['12'] +total_netsurplus['01'] +total_netsurplus['02']+total_netsurplus['03']+total_netsurplus['04']+total_netsurplus['05'])
        row['net_assets6'] =  format_with_parentheses(FYE_value + total_netsurplus['09'] +total_netsurplus['10'] +total_netsurplus['11'] +total_netsurplus['12'] +total_netsurplus['01'] +total_netsurplus['02']+total_netsurplus['03']+total_netsurplus['04']+total_netsurplus['05']+total_netsurplus['06'])
        row['net_assets7'] =  format_with_parentheses(FYE_value + total_netsurplus['09'] +total_netsurplus['10'] +total_netsurplus['11'] +total_netsurplus['12'] +total_netsurplus['01'] +total_netsurplus['02']+total_netsurplus['03']+total_netsurplus['04']+total_netsurplus['05']+total_netsurplus['06']+total_netsurplus['07'])
        row['net_assets8'] =  format_with_parentheses(FYE_value + total_netsurplus['09'] +total_netsurplus['10'] +total_netsurplus['11'] +total_netsurplus['12'] +total_netsurplus['01'] +total_netsurplus['02']+total_netsurplus['03']+total_netsurplus['04']+total_netsurplus['05']+total_netsurplus['06']+total_netsurplus['07']+total_netsurplus['08'])
       

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
        acct_per: "${:,}".format(abs(int(value))) if value > 0 else "(${:,})".format(abs(int(value))) if value < 0 else ""
        for acct_per, value in total_netsurplus.items() if value != 0
    }
    formatted_total_DnA = {
        acct_per: "{:,}".format(abs(int(value))) if value >= 0 else "({:,})".format(abs(int(value))) if value < 0 else ""
        for acct_per, value in total_DnA.items() if value!=0
    }

    
    
    bs_activity_list = list(set(row['Activity'] for row in data_balancesheet if 'Activity' in row))
    bs_activity_list_sorted = sorted(bs_activity_list)
    gl_obj = list(set(row['obj'] for row in data3 if 'obj' in row))
    gl_obj_sorted = sorted(gl_obj)

    # func_choice = list(set(row['func'] for row in data3 if 'func' in row))
    # func_choice_sorted = sorted(func_choice)        

    button_rendered = 0

    current_date = datetime.today().date()
    current_year = current_date.year
    last_year = current_date - timedelta(days=365)
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_name = last_month.strftime('%B')

    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 4
    ytd_budget = ytd_budget_test / 12
    formatted_ytd_budget = f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places

    if formatted_ytd_budget.startswith("0."):
        formatted_ytd_budget = formatted_ytd_budget[2:]
    
    context = { 
               'school': school,
               'school_name': SCHOOLS[school],
               'data_balancesheet': data_balancesheet ,
               'data_activitybs': data_activitybs,
               'data3': data3,
               'bs_activity_list': bs_activity_list_sorted,
               'gl_obj':gl_obj_sorted,
               'button_rendered': button_rendered,
               'last_month':last_month,
               'last_month_number':last_month_number,
               'last_month_name':last_month_name,
               'format_ytd_budget': formatted_ytd_budget,
               'ytd_budget':ytd_budget,
               'total_DnA': formatted_total_DnA,
               'total_netsurplus':formatted_total_netsurplus,
               'total_SBD':total_SBD,
               }

    return context



def cashflow(school):
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['object']};") 
    rows = cursor.fetchall()

    
    data = []
    for row in rows:
        if row[4] is None:
            row[4] = ''
        valueformat = "{:,.0f}".format(float(row[4])) if row[4] else ""
        row_dict = {
            'fund': row[0],
            'obj': row[1],
            'description': row[2],
            'category': row[3],
            'value': valueformat  
        }
        data.append(row_dict)

    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['function']};") 
    rows = cursor.fetchall()


    data2=[]
    for row in rows:
        budgetformat = "{:,.0f}".format(float(row[3])) if row[3] else ""
        row_dict = {
            'func_func': row[0],
            'desc': row[1],
            'category': row[2],
            'obj': row[4],
            'budget': budgetformat,
            
        }
        data2.append(row_dict)



    #
    cursor.execute(f"SELECT * FROM [dbo].{db[school]['db']};")  
    rows = cursor.fetchall()
    
    data3=[]
    
    
    if not school == "village-tech":
        for row in rows:
            expend = float(row[17])

            row_dict = {
                'fund':row[0],
                'func':row[1],
                'obj':row[2],
                'sobj':row[3],
                'org':row[4],
                'fscl_yr':row[5],
                'pgm':row[6],
                'edSpan':row[7],
                'projDtl':row[8],
                'AcctDescr':row[9],
                'Number':row[10],
                'Date':row[11],
                'AcctPer':row[12],
                'Est':row[13],
                'Real':row[14],
                'Appr':row[15],
                'Encum':row[16],
                'Expend':expend,
                'Bal':row[18],
                'WorkDescr':row[19],
                'Type':row[20],
                'Contr':row[21]
                }

            data3.append(row_dict)
            
    else:
        for row in rows:
            amount = float(row[19])
            row_dict = {
                'fund':row[0],
                'func':row[2],
                'obj':row[3],
                'sobj':row[4],
                'org':row[5],
                'fscl_yr':row[6],
     
                'Date':row[9],
                'AcctPer':row[10],
                'Amount':amount,
            }

            data3.append(row_dict)


    cursor.execute(f"SELECT * FROM [dbo].{db[school]['code']};") 
    rows = cursor.fetchall()
    
    data_expensebyobject=[]
    
    
    for row in rows:
        
        budgetformat = "{:,.0f}".format(float(row[2])) if row[2] else ""
        row_dict = {
            'obj':row[0],
            'Description':row[1],
            'budget':budgetformat,
            
            }
        
        data_expensebyobject.append(row_dict)

    cursor.execute(f"SELECT * FROM [dbo].{db[school]['activities']};") 
    rows = cursor.fetchall()
    
    data_activities=[]
    
    
    for row in rows:
        
      
        row_dict = {
            'obj':row[0],
            'Description':row[1],
            'Category':row[2],
            
            }
        
        data_activities.append(row_dict)

    cursor.execute(f"SELECT * FROM [dbo].{db[school]['cashflow']};")  
    rows = cursor.fetchall()
    
    data_cashflow=[]
    
    
    for row in rows:
        
      
        row_dict = {
            'Category':row[0],
            'Activity':row[1],
            'Description':row[2],
            'obj':str(row[3]),
            
            }
        
        data_cashflow.append(row_dict)

    cursor.execute(f"SELECT * FROM [dbo].{db[school]['bs_activity']};")   
    rows = cursor.fetchall()
    
    data_activitybs=[]
    
    
    for row in rows:

        row_dict = {
            'Activity':row[0],
            'obj':row[1],
            'Description2':row[2],
            
            
            }
        
        data_activitybs.append(row_dict)

    cursor.execute(f"SELECT  * FROM [dbo].{db[school]['bs']}") 
    rows = cursor.fetchall()
    
    data_balancesheet=[]
    
    
    for row in rows:
        fye = int(row[4]) if row[4] else 0
        if fye == 0:
            fyeformat = ""
        else:
            fyeformat = "{:,.0f}".format(abs(fye)) if fye >= 0 else "({:,.0f})".format(abs(fye))
        row_dict = {
            'Activity':row[0],
            'Description':row[1],
            'Category':row[2],
            'Subcategory':row[3],
            'FYE':fyeformat,
            }
        
        data_balancesheet.append(row_dict)
    
    acct_per_values = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']

    activity_key = 'Bal'
    if school == "village-tech":
        activity_key = 'Amount'

    for item in data_activitybs:
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            key = f'total_bal{i}' 
            item[key] = sum(
                entry[activity_key] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
            )  


    #---------- FOR EXPENSE TOTAL -------
    for item in data_cashflow:
        activity = item['Activity']

        for i, acct_per in enumerate(acct_per_values, start=1):
            key = f'total_bal{i}' 
            item[f'total_operating{i}'] = sum(
                entry[key] for entry in data_activitybs if entry['Activity'] == activity
            )

    for item in data_cashflow:
        
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_investing{i}'] = sum(
                entry['Bal'] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
            )
            

    
    for item in data_activities:
        
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_activities{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
            )

    activity_sum_dict = {} 
    for item in data_activitybs:
        Activity = item['Activity']
        for i in range(1, 13):
            total_sum_i = sum(
                int(entry[f'total_bal{i}']) if entry[f'total_bal{i}'] and entry['Activity'] == Activity else 0
                for entry in data_activitybs
            )
            activity_sum_dict[(Activity, i)] = total_sum_i
    

 
    for row in data_balancesheet:
        
        activity = row['Activity']
        for i in range(1, 13):
            key = (activity, i)
            row[f'total_sum{i}'] = "{:,.0f}".format(activity_sum_dict.get(key, 0))
            

    def format_with_parentheses(value):
        if value == 0:
            return ""
        formatted_value = "{:,.0f}".format(abs(value))
        return "({})".format(formatted_value) if value < 0 else formatted_value
    
    for row in data_balancesheet:
    
        FYE_value = int(row['FYE'].replace(',', '').replace('(', '-').replace(')', '')) if row['FYE'] else 0
        total_sum9_value = int(row['total_sum9'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum9'] else 0
        total_sum10_value = int(row['total_sum10'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum10'] else 0
        total_sum11_value = int(row['total_sum11'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum11'] else 0
        total_sum12_value = int(row['total_sum12'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum12'] else 0
        total_sum1_value = int(row['total_sum1'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum1'] else 0
        total_sum2_value = int(row['total_sum2'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum2'] else 0
        total_sum3_value = int(row['total_sum3'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum3'] else 0
        total_sum4_value = int(row['total_sum4'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum4'] else 0
        total_sum5_value = int(row['total_sum5'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum5'] else 0
        total_sum6_value = int(row['total_sum6'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum6'] else 0
        total_sum7_value = int(row['total_sum7'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum7'] else 0
        total_sum8_value = int(row['total_sum8'].replace(',', '').replace('(', '-').replace(')', '')) if row['total_sum8'] else 0

        # Calculate the differences and store them in the row dictionary
        row['difference_9'] = format_with_parentheses(FYE_value + total_sum9_value)
        row['difference_10'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value)
        row['difference_11'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value)
        row['difference_12'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value)
        row['difference_1'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value)
        row['difference_2'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value)
        row['difference_3'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value)
        row['difference_4'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value)
        row['difference_5'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value)
        row['difference_6'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value)
        row['difference_7'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value)
        row['difference_8'] = format_with_parentheses(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value)

        row['fytd'] = format_with_parentheses(total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value)
    
    keys_to_check_expense = ['total_activities1', 'total_activities2', 'total_activities3', 'total_activities4', 'total_activities5','total_activities6','total_activities7','total_activities8','total_activities9','total_activities10','total_activities11','total_activities12']
    keys_to_check_expense2 = ['total_expense1', 'total_expense2', 'total_expense3', 'total_expense4', 'total_expense5','total_expense6','total_expense7','total_expense8','total_expense9','total_expense10','total_expense11','total_expense12']

    keys_to_check_cashflow = ['total_operating1','total_operating2','total_operating3','total_operating4','total_operating5','total_operating6','total_operating7','total_operating8','total_operating9','total_operating10','total_operating11','total_operating12']
    keys_to_check_cashflow2 = ['total_investing1','total_investing2','total_investing3','total_investing4','total_investing5','total_investing6','total_investing7','total_investing8','total_investing9','total_investing10','total_investing11','total_investing12']
    for item in data_expensebyobject:
        obj = item['obj']
        if obj == '6100':
            category = 'Payroll Costs'
        elif obj == '6200':
            category = 'Professional and Cont Svcs'
        elif obj == '6300':
            category = 'Supplies and Materials'
        elif obj == '6400':
            category = 'Other Operating Expenses'
        else:
            category = 'Total Expense'

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_expense{i}'] = sum(
                entry[f'total_activities{i}'] for entry in data_activities if entry['Category'] == category 
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

    for row in data_cashflow:
      for key in keys_to_check_cashflow:
          value = float(row[key])
          if value == 0:
              row[key] = ""
          elif value < 0:
              row[key] = "({:,.0f})".format(abs(float(row[key]))) 
          elif value != "":
              row[key] = "{:,.0f}".format(float(row[key])) 


    for row in data_cashflow:
      for key in keys_to_check_cashflow2:
          value = float(row[key])
          if value == 0:
              row[key] = ""
          elif value < 0:
              row[key] = "({:,.0f})".format(abs(float(row[key]))) 
          elif value != "":
              row[key] = "{:,.0f}".format(float(row[key]))      
        
    
   
    

    


    #---- for data ------
   
    
    
    
    #--- total revenue
    total_revenue = {acct_per: 0 for acct_per in acct_per_values}

    for item in data:
        fund = item['fund']
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_real{i}'] = sum(
                entry['Real'] for entry in data3 if entry['fund'] == fund and entry['obj'] == obj and entry['AcctPer'] == acct_per
            )

            total_revenue[acct_per] += abs(item[f'total_real{i}'])
    # total_revenue9 = 0 
    # total_revenue10 = 0
    # total_revenue11 = 0
    # total_revenue12 = 0
    # total_revenue1 = 0
    # total_revenue2 = 0
    # total_revenue3 = 0
    # total_revenue4 = 0
    # total_revenue5 = 0
    # total_revenue6 = 0
    # total_revenue7 = 0
    # total_revenue8 = 0
   
    # for item in data:
    #     fund = item['fund']
    #     obj = item['obj']

    #     for i, acct_per in enumerate(acct_per_values, start=1):
    #         item[f'total_real{i}'] = sum(
    #             entry['Real'] for entry in data3 if entry['fund'] == fund and entry['obj'] == obj and entry['AcctPer'] == acct_per
    #         )
            
    
    #         if acct_per == '09':
    #             total_revenue9 += item[f'total_real{i}']
    #         if acct_per == '10':
    #             total_revenue10 += item[f'total_real{i}']
    #         if acct_per == '11':
    #             total_revenue11 += item[f'total_real{i}']
    #         if acct_per == '12':
    #             total_revenue12 += item[f'total_real{i}']
    #         if acct_per == '01':
    #             total_revenue1 += item[f'total_real{i}']
    #         if acct_per == '02':
    #             total_revenue2 += item[f'total_real{i}']
    #         if acct_per == '03':
    #             total_revenue3 += item[f'total_real{i}']
    #         if acct_per == '04':
    #             total_revenue4 += item[f'total_real{i}']
    #         if acct_per == '05':
    #             total_revenue5 += item[f'total_real{i}']
    #         if acct_per == '06':
    #             total_revenue6 += item[f'total_real{i}']
    #         if acct_per == '07':
    #             total_revenue7 += item[f'total_real{i}']
    #         if acct_per == '08':
    #             total_revenue8 += item[f'total_real{i}']
                
    keys_to_check = ['total_real1', 'total_real2', 'total_real3', 'total_real4', 'total_real5','total_real6','total_real7','total_real8','total_real9','total_real10','total_real11','total_real12']
 
    for row in data:
        for key in keys_to_check:
            if row[key] < 0:
                row[key] = -row[key]
            else:
                row[key] = ''

    for row in data:
        for key in keys_to_check:
            if row[key] != "":
                row[key] = "{:,.0f}".format(row[key])
                
    



    acct_per_values2 = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
    
    
    total_surplus = {acct_per: 0 for acct_per in acct_per_values2}

    for item in data2:
        if item['category'] != 'Depreciation and Amortization':
            func = item['func_func']

            for i, acct_per in enumerate(acct_per_values2, start=1):
                item[f'total_func{i}'] = sum(
                    entry['Expend'] for entry in data3 if entry['func'] == func and entry['AcctPer'] == acct_per
                )
                total_surplus[acct_per] += item[f'total_func{i}']
    # total_surplus9 = 0 
    # total_surplus10 = 0
    # total_surplus11 = 0
    # total_surplus12 = 0
    # total_surplus1 = 0
    # total_surplus2 = 0
    # total_surplus3 = 0
    # total_surplus4 = 0
    # total_surplus5 = 0
    # total_surplus6 = 0
    # total_surplus7 = 0
    # total_surplus8 = 0
    # for item in data2:
    #     func = item['func_func']
        

    #     for i, acct_per in enumerate(acct_per_values2, start=1):
    #         item[f'total_func{i}'] = sum(
    #             entry['Expend'] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per
    #         )
    #         if acct_per == '09':
    #             total_surplus9 += item[f'total_func{i}']
    #         if acct_per == '10':
    #             total_surplus10 += item[f'total_func{i}']
    #         if acct_per == '11':
    #             total_surplus11 += item[f'total_func{i}']
    #         if acct_per == '12':
    #             total_surplus12 += item[f'total_func{i}']
    #         if acct_per == '01':
    #             total_surplus1 += item[f'total_func{i}']
    #         if acct_per == '02':
    #             total_surplus2 += item[f'total_func{i}']
    #         if acct_per == '03':
    #             total_surplus3 += item[f'total_func{i}']
    #         if acct_per == '04':
    #             total_surplus4 += item[f'total_func{i}']
    #         if acct_per == '05':
    #             total_surplus5 += item[f'total_func{i}']
    #         if acct_per == '06':
    #             total_surplus6 += item[f'total_func{i}']
    #         if acct_per == '07':
    #             total_surplus7 += item[f'total_func{i}']
    #         if acct_per == '08':
    #             total_surplus8 += item[f'total_func{i}']

    #---- Depreciation and ammortization total
    total_DnA = {acct_per: 0 for acct_per in acct_per_values2}
    
    for item in data2:
        func = item['func_func']
        obj =  item['obj']
    
        for i, acct_per in enumerate(acct_per_values2, start=1):
            item[f'total_func2_{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['func'] == func and entry['AcctPer'] == acct_per and entry['obj'] == obj
            )
            total_DnA[acct_per] += item[f'total_func2_{i}']
        
    
    total_SBD = {acct_per: total_revenue[acct_per] - total_surplus[acct_per]  for acct_per in acct_per_values}
    total_netsurplus = {acct_per: total_SBD[acct_per] - total_DnA[acct_per]  for acct_per in acct_per_values}
    formatted_total_netsurplus = {
        acct_per: "${:,}".format(abs(int(value))) if value > 0 else "(${:,})".format(abs(int(value))) if value < 0 else ""
        for acct_per, value in total_netsurplus.items() if value != 0
    }
    formatted_total_DnA = {
        acct_per: "{:,}".format(abs(int(value))) if value >= 0 else "({:,})".format(abs(int(value))) if value < 0 else ""
        for acct_per, value in total_DnA.items() if value!=0
    }
    
    # total_surplusBD9 = 0 
    # total_surplusBD10 = 0
    # total_surplusBD11 = 0
    # total_surplusBD12 = 0
    # total_surplusBD1 = 0
    # total_surplusBD2 = 0
    # total_surplusBD3 = 0
    # total_surplusBD4 = 0
    # total_surplusBD5 = 0
    # total_surplusBD6 = 0
    # total_surplusBD7 = 0
    # total_surplusBD8 = 0
    # for item in data2:
    #     func = item['func_func']
        

    #     for i, acct_per in enumerate(acct_per_values2, start=1):
    #         item[f'total_func2_{i}'] = sum(
    #             entry['Expend'] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per and entry['obj'] == '6449'
    #         )
            # if acct_per == '09':
            #     total_surplusBD9 += item[f'total_func2_{i}']
            # if acct_per == '10':
            #     total_surplusBD10 += item[f'total_func2_{i}']
            # if acct_per == '11':
            #     total_surplusBD11 += item[f'total_func2_{i}']
            # if acct_per == '12':
            #     total_surplusBD12 += item[f'total_func2_{i}']
            # if acct_per == '01':
            #     total_surplusBD1 += item[f'total_func2_{i}']
            # if acct_per == '02':
            #     total_surplusBD2 += item[f'total_func2_{i}']
            # if acct_per == '03':
            #     total_surplusBD3 += item[f'total_func2_{i}']
            # if acct_per == '04':
            #     total_surplusBD4 += item[f'total_func2_{i}']
            # if acct_per == '05':
            #     total_surplusBD5 += item[f'total_func2_{i}']
            # if acct_per == '06':
            #     total_surplusBD6 += item[f'total_func2_{i}']
            # if acct_per == '07':
            #     total_surplusBD7 += item[f'total_func{i}']
            # if acct_per == '08':
            #     total_surplusBD8 += item[f'total_func{i}']


    
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

                
          



    lr_funds = list(set(row['fund'] for row in data3 if 'fund' in row))
    lr_funds_sorted = sorted(lr_funds)
    lr_obj = list(set(row['obj'] for row in data3 if 'obj' in row))
    lr_obj_sorted = sorted(lr_obj)

    func_choice = list(set(row['func'] for row in data3 if 'func' in row))
    func_choice_sorted = sorted(func_choice)
    
            
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
         'school' : school,
         'school_name': SCHOOLS[school],
         'data': data, 
         'data2':data2 , 
         'data3': data3 ,
         'data_cashflow': data_cashflow,
         'data_activitybs': data_activitybs,
         'data_balancesheet':data_balancesheet,
         'lr_funds':lr_funds_sorted, 
         'lr_obj':lr_obj_sorted, 
         'func_choice':func_choice_sorted ,
         'data_expensebyobject': data_expensebyobject,
         'data_activities': data_activities,
         'last_month':last_month,
         'last_month_number':last_month_number,
         'format_ytd_budget': formatted_ytd_budget,
         'ytd_budget':ytd_budget,
         'total_DnA': formatted_total_DnA,
         'total_netsurplus':formatted_total_netsurplus,
         'total_SBD':total_SBD,
          }
    return context

def general_ledger(school):
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute(f"SELECT  TOP(300)* FROM [dbo].{db[school]['db']}")
    rows = cursor.fetchall()
    
    data3 = []
    for row in rows:
        date_str = row[11]
        # if date_str is not None:
        #         date_without_time = date_str.strftime('%b. %d, %Y')
        # else:
        #         date_without_time = None 
        row_dict = {
            'fund':row[0],
            'func':row[1],
            'obj':row[2],
            'sobj':row[3],
            'org':row[4],
            'fscl_yr':row[5],
            'pgm':row[6],
            'edSpan':row[7],
            'projDtl':row[8],
            'AcctDescr':row[9],
            'Number':row[10],
            'Date':date_str,
            'AcctPer':row[12],
            'Est':row[13],
            'Real':row[14],
            'Appr':row[15],
            'Encum':row[16],
            'Expend':row[17],
            'Bal':row[18],
            'WorkDescr':row[19],
            'Type':row[20],
            'Contr':row[21]
            }
        
        data3.append(row_dict)

    context = {
        'data3': data3, 
         }
    context["school"] = school
    context["school_name"] = SCHOOLS[school]
    return context
