from django.shortcuts import render,redirect,get_object_or_404
from .forms import UploadForm
import pandas as pd
import io
from .models import User
from django.contrib import auth
from django.contrib.auth import login, authenticate
from django.contrib.auth import logout
from django.urls import reverse_lazy
import sys
from time import strftime
import os
import pyodbc
import sqlalchemy as sa
from sqlalchemy import create_engine
from urllib.parse import quote_plus
import datetime
import locale
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
import json
from datetime import datetime,timedelta
from dateutil.relativedelta import relativedelta


def connect():
    server = 'aca-mysqlserver1.database.windows.net'
    database = 'Database1'
    username = 'aca-user1'
    password = 'Pokemon!123'
    port = '1433'
    
    driver = '{/usr/lib/libmsodbcsql-17.so}'

    cnxn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    return cnxn

def loginView(request):
    if request.method == 'POST':
        username = request.POST['username']
        password = request.POST['password']
        user = authenticate(username=username, password=password)
        if user is not None and user.is_active:
            if user.is_admin or user.is_superuser:
                auth.login(request, user)
                return redirect('pl_advantage')
            elif user is not None and user.is_employee:
                auth.login(request, user)
                return redirect('pl_advantage')
            
            else:
                return redirect('login_form')
        else:
            
            return redirect('login_form')

def login_form(request):
    return render(request,'login.html')



def logoutView(request):
	logout(request)
	return redirect('login_form')

def pl_advantage(request):
    
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Advantage_Definition_obj];") 
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

    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Advantage_Definition_func];") 
    rows = cursor.fetchall()


    data2=[]
    for row in rows:
        budgetformat = "{:,.0f}".format(float(row[3])) if row[3] else ""
        row_dict = {
            'func_func': row[0],
            'desc': row[1],
            'budget': budgetformat,
            
        }
        data2.append(row_dict)



    #
    cursor.execute("SELECT * FROM [dbo].[AscenderData_Advantage];") 
    rows = cursor.fetchall()
    
    data3=[]
    
    
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


    cursor.execute("SELECT * FROM [dbo].[AscenderData_Advantage_PL_ExpensesbyObjectCode];") 
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

    cursor.execute("SELECT * FROM [dbo].[AscenderData_Advantage_PL_Activities];") 
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
    acct_per_values_expense = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
    for item in data_activities:
        
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values_expense, start=1):
            item[f'total_activities{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
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

    for item in data:
        fund = item['fund']
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_real{i}'] = sum(
                entry['Real'] for entry in data3 if entry['fund'] == fund and entry['obj'] == obj and entry['AcctPer'] == acct_per
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

    for item in data2:
        func = item['func_func']
        

        for i, acct_per in enumerate(acct_per_values2, start=1):
            item[f'total_func{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per
            )

    for item in data2:
        func = item['func_func']
        

        for i, acct_per in enumerate(acct_per_values2, start=1):
            item[f'total_func2_{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per and entry['obj'] == '6449'
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
    return render(request,'dashboard/advantage/pl_advantage.html', context)

def first_advantage(request):
    current_date = datetime.today().date()
    current_year = current_date.year
    last_year = current_date - timedelta(days=365)
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 4 
    ytd_budget = ytd_budget_test / 12
    formatted_ytd_budget = f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
    print(last_month)
    if formatted_ytd_budget.startswith("0."):
        formatted_ytd_budget = formatted_ytd_budget[2:]
    context = {
        'last_month':last_month,
        'last_month_number':last_month_number,
        'format_ytd_budget': formatted_ytd_budget,
        'ytd_budget':ytd_budget,
    }
    return render(request,'dashboard/advantage/first_advantage.html', context)

def first_cumberland(request):
    current_date = datetime.today().date()
    current_year = current_date.year
    last_year = current_date - timedelta(days=365)
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 4 
    ytd_budget = ytd_budget_test / 12
    formatted_ytd_budget = f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
    print(last_month)
    if formatted_ytd_budget.startswith("0."):
        formatted_ytd_budget = formatted_ytd_budget[2:]
    context = {
        'last_month':last_month,
        'last_month_number':last_month_number,
        'format_ytd_budget': formatted_ytd_budget,
        'ytd_budget':ytd_budget,
    }
    return render(request,'dashboard/cumberland/first_cumberland.html', context)

def pl_cumberland(request):
    
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Cumberland_Definition_obj];") 
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
            'value':  valueformat
        }
        data.append(row_dict)

    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Cumberland_Definition_func];") 
    rows = cursor.fetchall()


    data2=[]
    for row in rows:
        budgetformat = "{:,.0f}".format(float(row[3])) if row[3] else ""
        row_dict = {
            'func_func': row[0],
            'desc': row[1],
            'budget': budgetformat,
            
        }
        data2.append(row_dict)


    #
    cursor.execute("SELECT * FROM [dbo].[AscenderData_Cumberland];") 
    rows = cursor.fetchall()
    
    data3=[]
    
    
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
    
    




    cursor.execute("SELECT * FROM [dbo].[AscenderData_Cumberland_PL_ExpensesbyObjectCode];") 
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

    cursor.execute("SELECT * FROM [dbo].[AscenderData_Cumberland_PL_Activities];") 
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
    acct_per_values_expense = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
    for item in data_activities:
        
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values_expense, start=1):
            

            item[f'total_activities{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
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

    for item in data:
        fund = item['fund']
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_real{i}'] = sum(
                entry['Real'] for entry in data3 if entry['fund'] == fund and entry['obj'] == obj and entry['AcctPer'] == acct_per
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

    for item in data2:
        func = item['func_func']
        

        for i, acct_per in enumerate(acct_per_values2, start=1):
            item[f'total_func{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per
            )

    for item in data2:
        func = item['func_func']
        

        for i, acct_per in enumerate(acct_per_values2, start=1):
            item[f'total_func2_{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per and entry['obj'] == '6449'
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
    print(ytd_budget) 
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
    return render(request,'dashboard/cumberland/pl_cumberland.html', context)




    

def insert_row(request):
    if request.method == 'POST':
        print(request.POST)
        try:
            #<<<------------------ LR INSERT FUNCTION ---------------------->>>
            funds = request.POST.getlist('fund[]')
            objs = request.POST.getlist('obj[]')
            descriptions = request.POST.getlist('Description[]')
            budgets = request.POST.getlist('budget[]')

            data_list_LR = []
            for fund, obj, description, budget in zip(funds, objs, descriptions, budgets):
                if fund.strip() and obj.strip() and budget.strip():
                    data_list_LR.append({
                        'fund': fund,
                        'obj': obj,
                        'description': description,
                        'budget': budget,
                        'category': 'Local Revenue',
                    })

          
            cnxn = connect()
            cursor = cnxn.cursor()

          
            for data in data_list_LR:
                fund = data['fund']
                obj = data['obj']
                description = data['description']
                budget = data['budget']
                category = data['category']

                query = "INSERT INTO [dbo].[AscenderData_Advantage_Definition_obj] (budget, fund, obj, Description, Category) VALUES (?, ?, ?, ?, ?)"
                cursor.execute(query, (budget, fund, obj, description, category))
                cnxn.commit()

            



            #<--------------UPDATE FOR LOCAL REVENUE,SPR,FPR----------
            updatefunds = request.POST.getlist('updatefund[]')  
            updatevalues = request.POST.getlist('updatevalue[]')
            updateobjs = request.POST.getlist('updateobj[]')
            
            

            updatedata_list = []

            for updatefund,updatevalue,updateobj in zip(updatefunds, updatevalues,updateobjs):
                if updatefund.strip() and updatevalue.strip() :
                    updatedata_list.append({
                        'updatefund': updatefund,
                        'updateobj':updateobj,
                        
                        'updatevalue': updatevalue,
                        
                    })
            for data in updatedata_list:
                updatefund= data['updatefund']
                updateobj=data['updateobj']
                
                updatevalue = data['updatevalue']

                try:
                    query = "UPDATE [dbo].[AscenderData_Advantage_Definition_obj] SET budget = ? WHERE fund = ? and obj = ? "
                    cursor.execute(query, (updatevalue, updatefund,updateobj))
                    cnxn.commit()
                    print(f"Rows affected for fund={updatefund}: {cursor.rowcount}")
                except Exception as e:
                    print(f"Error updating fund={updatefund}: {str(e)}")


            #<---------------------------------- INSERT FOR SPR ---->>>
            fundsSPR = request.POST.getlist('fundSPR[]')
            objsSPR = request.POST.getlist('objSPR[]')
            descriptionsSPR = request.POST.getlist('DescriptionSPR[]')
            budgetsSPR = request.POST.getlist('budgetSPR[]')

            data_list_SPR = []
            for fund, obj, description, budget in zip(fundsSPR, objsSPR, descriptionsSPR, budgetsSPR):
                if fund.strip() and obj.strip() and budget.strip():
                    data_list_SPR.append({
                        'fund': fund,
                        'obj': obj,
                        'description': description,
                        'budget': budget,
                        'category': 'State Program Revenue',
                    })
          
            for data in data_list_SPR:
                fund = data['fund']
                obj = data['obj']
                description = data['description']
                budget = data['budget']
                category = data['category']

                query = "INSERT INTO [dbo].[AscenderData_Advantage_Definition_obj] (budget, fund, obj, Description, Category) VALUES (?, ?, ?, ?, ?)"
                cursor.execute(query, (budget, fund, obj, description, category))
                cnxn.commit()

            # #<--------------UPDATE FOR STATE PROGRAM REVENUE ----------
            # updatefundsSPR = request.POST.getlist('updatefundSPR[]')  
            # updatevaluesSPR = request.POST.getlist('updatevalueSPR[]')
            # updateobjsSPR = request.POST.getlist('updateobjSPR[]')
            


            # updatedata_list_SPR = []

            # for updatefund,updatevalue,updateobj in zip(updatefundsSPR, updatevaluesSPR,updateobjsSPR):
            #     if updatefund.strip() and updatevalue.strip() and updatevalue.strip() != " ":
            #         updatedata_list_SPR.append({
            #             'updatefund': updatefund,
            #             'updateobj':updateobj,
                        
            #             'updatevalue': updatevalue,
                        
            #         })
            # for data in updatedata_list_SPR:
            #     updatefund= data['updatefund']
            #     updateobj=data['updateobj']
                
            #     updatevalue = data['updatevalue']

            #     try:
            #         query = "UPDATE [dbo].[AscenderData_Definition_obj] SET budget = ? WHERE fund = ? and obj = ? "
            #         cursor.execute(query, (updatevalue, updatefund,updateobj))
            #         cnxn.commit()
            #         print(f"Rows affected for fund={updatefund}: {cursor.rowcount}")
            #     except Exception as e:
            #         print(f"Error updating fund={updatefund}: {str(e)}")    
            
            #<---------------------------------- INSERT FOR FPR ---->>>
            fundsFPR = request.POST.getlist('fundFPR[]')
            objsFPR = request.POST.getlist('objFPR[]')
            descriptionsFPR = request.POST.getlist('DescriptionFPR[]')
            budgetsFPR = request.POST.getlist('budgetFPR[]')

            data_list_FPR = []
            for fund, obj, description, budget in zip(fundsFPR, objsFPR, descriptionsFPR, budgetsFPR):
                if fund.strip() and obj.strip() and budget.strip():
                    data_list_FPR.append({
                        'fund': fund,
                        'obj': obj,
                        'description': description,
                        'budget': budget,
                        'category': 'Federal Program Revenue',
                    })
          
            for data in data_list_FPR:
                fund = data['fund']
                obj = data['obj']
                description = data['description']
                budget = data['budget']
                category = data['category']

                query = "INSERT INTO [dbo].[AscenderData_Advantage_Definition_obj] (budget, fund, obj, Description, Category) VALUES (?, ?, ?, ?, ?)"
                cursor.execute(query, (budget, fund, obj, description, category))
                cnxn.commit()
            
            #<---------------------------------- INSERT FOR Func ---->>>
            newfuncs = request.POST.getlist('newfunc[]')
            newDescriptionfuncs = request.POST.getlist('newDescriptionfunc[]')
            newBudgetfuncs = request.POST.getlist('newBudgetfunc[]')
            

            data_list_func = []
            for func, description, budget in zip(newfuncs, newDescriptionfuncs,newBudgetfuncs):
                if func.strip() and budget.strip():
                    data_list_func.append({
                        'func': func,
                        'description': description,
                        'budget': budget,
                        
                    })
          
            for data in data_list_func:
                func = data['func']
                description = data['description']
                budget = data['budget']
                

                query = "INSERT INTO [dbo].[AscenderData_Advantage_Definition_func] (budget, func, Description) VALUES (?, ?, ?)"
                cursor.execute(query, (budget, func, description))
                cnxn.commit()
            
            updatefunds = request.POST.getlist('updatefund[]')  
            updatevalues = request.POST.getlist('updatevalue[]')
            updateobjs = request.POST.getlist('updateobj[]')
            

            #<------------------ update for func func ---------------
            updatefuncfuncs = request.POST.getlist('updatefuncfunc[]')  
            updatefuncbudgets = request.POST.getlist('updatefuncbudget[]')
            


            updatedata_list_func = []

            for updatefunc,updatebudget in zip(updatefuncfuncs, updatefuncbudgets):
                if updatefunc.strip() and updatebudget.strip() and updatebudget.strip() != " ":
                    updatedata_list_func.append({
                        'updatefunc': updatefunc,
                        'updatebudget':updatebudget,
                        
                        
                    })
            for data in updatedata_list_func:
                updatefunc= data['updatefunc']
                updatebudget=data['updatebudget']
                
                

                try:
                    query = "UPDATE [dbo].[AscenderData_Advantage_Definition_func] SET budget = ? WHERE func = ? "
                    cursor.execute(query, (updatebudget, updatefunc))
                    cnxn.commit()
                    print(f"Rows affected for fund={updatefunc}: {cursor.rowcount}")
                except Exception as e:
                    print(f"Error updating fund={updatefunc}: {str(e)}")

                


            cursor.close()
            cnxn.close()

            return redirect('pl_advantage')

        except Exception as e:
            return JsonResponse({'status': 'error', 'message': str(e)})

    return JsonResponse({'status': 'error', 'message': 'Invalid request method'})


def delete(request, fund, obj):
    try:
        
        cnxn = connect()
        cursor = cnxn.cursor()

        
        query = "DELETE FROM [dbo].[AscenderData_Definition_obj] WHERE fund = ? AND obj = ?"
        cursor.execute(query, (fund, obj))
        cnxn.commit()

        
        cursor.close()
        cnxn.close()

        return redirect('pl_advantage')

    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})

def delete_func(request, func):
    try:
        
        cnxn = connect()
        cursor = cnxn.cursor()

        
        query = "DELETE FROM [dbo].[AscenderData_Definition_func] WHERE func = ?"
        cursor.execute(query, (func))
        cnxn.commit()

        
        cursor.close()
        cnxn.close()

        return redirect('pl_advantage')

    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})


def delete_bs(request, description, subcategory):
    try:
        print(request)
        
        cnxn = connect()
        cursor = cnxn.cursor()

        
        query = "DELETE FROM [dbo].[AscenderData_Advantage_Balancesheet] WHERE Description = ? and Subcategory = ?"
        cursor.execute(query, (description,subcategory))
        cnxn.commit()

        
        cursor.close()
        cnxn.close()

        return redirect('bs_advantage')

    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})

def delete_bsa(request, obj, Activity):
    try:
     
        cnxn = connect()
        cursor = cnxn.cursor()

        
        query = "DELETE FROM [dbo].[AscenderData_Advantage_ActivityBS] WHERE obj = ? and Activity = ?"
        cursor.execute(query, (obj,Activity))
        cnxn.commit()

        
        cursor.close()
        cnxn.close()

        return redirect('bs_advantage')

    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})



def pl_advantagechart(request):
    return render(request,'dashboard/advantage/pl_advantagechart.html')

def pl_cumberlandchart(request):
    return render(request,'dashboard/cumberland/pl_cumberlandchart.html')

def viewgl(request,fund,obj,yr):
    print(request)
    try:
        
        cnxn = connect()
        cursor = cnxn.cursor()
        query = "SELECT * FROM [dbo].[AscenderData_Advantage] WHERE fund = ? and obj = ? and AcctPer = ?"
        cursor.execute(query, (fund,obj,yr))
        
        rows = cursor.fetchall()
    
        gl_data=[]
    
    
        for row in rows:
            date_str=row[11]

            # date_without_time = date_str.strftime('%b. %d, %Y')


            real = float(row[14]) if row[14] else 0
            if real == 0:
                realformat = ""
            else:
                realformat = "{:,.0f}".format(abs(real)) if real >= 0 else "({:,.0f})".format(abs(real))

            
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
                'Real':realformat,
                'Appr':row[15],
                'Encum':row[16],
                'Expend':row[17],
                'Bal':row[18],
                'WorkDescr':row[19],
                'Type':row[20],
                'Contr':row[21]
            }

            gl_data.append(row_dict)
        
        total_bal = sum(float(row['Real'].replace(',', '').replace('(', '-').replace(')', '')) for row in gl_data)
        total_bal = "{:,.0f}".format(abs(total_bal)) if total_bal >= 0 else "({:,.0f})".format(abs(total_bal))

        context = { 
            'gl_data':gl_data,
            'total_bal':total_bal
        }

        
        cursor.close()
        cnxn.close()

        return JsonResponse({'status': 'success', 'data': context})

    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})

def viewgl_cumberland(request,fund,obj,yr):
    print(request)
    try:
        
        cnxn = connect()
        cursor = cnxn.cursor()


        
        
        query = "SELECT * FROM [dbo].[AscenderData_Cumberland] WHERE fund = ? and obj = ? and AcctPer = ?"
        cursor.execute(query, (fund,obj,yr))
        
        rows = cursor.fetchall()
    
        gl_data=[]
    
    
        for row in rows:
            date_str=row[11]

          

            # real = float(row[14]) if row[14] else 0
            # if real == 0:
            #     realformat = ""
            # else:
            #     realformat = "{:,.0f}".format(abs(real)) if real >= 0 else "({:,.0f})".format(abs(real))

            
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

            gl_data.append(row_dict)
        
        total_bal = sum(float(row['Real']) for row in gl_data)
        total_bal = "{:,.0f}".format(abs(total_bal)) if total_bal >= 0 else "({:,.0f})".format(abs(total_bal))
        
        
        context = { 
            'gl_data':gl_data,
            'total_bal':total_bal
        }

        
        cursor.close()
        cnxn.close()

        return JsonResponse({'status': 'success', 'data': context})

    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})

def gl_advantage(request):
    
    cnxn = connect()
    cursor = cnxn.cursor()
    
    cursor.execute("SELECT  TOP(300)* FROM [dbo].[AscenderData_Advantage]") 
    rows = cursor.fetchall()
    
    data3=[]
    
    
    for row in rows:
        date_str=row[11]
        
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
        
        'data3': data3 , 
         }
    return render(request,'dashboard/advantage/gl_advantage.html', context)


def gl_cumberland(request):
    
    cnxn = connect()
    cursor = cnxn.cursor()
    
    cursor.execute("SELECT  TOP(300)* FROM [dbo].[AscenderData_Cumberland]") 
    rows = cursor.fetchall()
    
    data3=[]
    
    
    for row in rows:
        date_str=row[11]
        
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
        
        'data3': data3 , 
         }
    return render(request,'dashboard/cumberland/gl_cumberland.html', context)

def bs_advantage(request):
    
    cnxn = connect()
    cursor = cnxn.cursor()
    
    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Advantage_Balancesheet]") 
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

    cursor.execute("SELECT * FROM [dbo].[AscenderData_Advantage_ActivityBS]") 
    rows = cursor.fetchall()
    
    data_activitybs=[]
    
    
    for row in rows:

        row_dict = {
            'Activity':row[0],
            'obj':row[1],
            'Description2':row[2],
            
            
            }
        
        data_activitybs.append(row_dict)

    cursor.execute("SELECT * FROM [dbo].[AscenderData_Advantage]") 
    rows = cursor.fetchall()
    
    data3=[]
    
    
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

    acct_per_values = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']

    for item in data_activitybs:
        
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_bal{i}'] = sum(
                entry['Bal'] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
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
    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 4 
    ytd_budget = ytd_budget_test / 12
    formatted_ytd_budget = f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places

    if formatted_ytd_budget.startswith("0."):
        formatted_ytd_budget = formatted_ytd_budget[2:]
    
    context = { 
        
        'data_balancesheet': data_balancesheet ,
        'data_activitybs': data_activitybs,
        'data3': data3,
        'bs_activity_list': bs_activity_list_sorted,
        'gl_obj':gl_obj_sorted,
        'button_rendered': button_rendered,
        'last_month':last_month,
        'last_month_number':last_month_number,
        'format_ytd_budget': formatted_ytd_budget,
        'ytd_budget':ytd_budget,
        
         }

    return render(request,'dashboard/advantage/bs_advantage.html', context)


def bs_cumberland(request):
    
    cnxn = connect()
    cursor = cnxn.cursor()
    
    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Cumberland_Balancesheet]") 
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

    cursor.execute("SELECT * FROM [dbo].[AscenderData_Cumberland_ActivityBS]") 
    rows = cursor.fetchall()
    
    data_activitybs=[]
    
    
    for row in rows:

        row_dict = {
            'Activity':row[0],
            'obj':row[1],
            'Description2':row[2],
            
            
            }
        
        data_activitybs.append(row_dict)

    cursor.execute("SELECT * FROM [dbo].[AscenderData_Cumberland]") 
    rows = cursor.fetchall()
    
    data3=[]
    
    
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

    acct_per_values = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']

    for item in data_activitybs:
        
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_bal{i}'] = sum(
                entry['Bal'] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
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
    



    
    
    bs_activity_list = list(set(row['Activity'] for row in data_balancesheet if 'Activity' in row))
    bs_activity_list_sorted = sorted(bs_activity_list)
    gl_obj = list(set(row['obj'] for row in data3 if 'obj' in row))
    gl_obj_sorted = sorted(gl_obj)

    # func_choice = list(set(row['func'] for row in data3 if 'func' in row))
    # func_choice_sorted = sorted(func_choice)        

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
    
    context = { 
        
        'data_balancesheet': data_balancesheet ,
        'data_activitybs': data_activitybs,
        'data3': data3,
        'bs_activity_list': bs_activity_list_sorted,
        'gl_obj':gl_obj_sorted,
        'last_month':last_month,
        'last_month_number':last_month_number,
        'format_ytd_budget': formatted_ytd_budget,
        'ytd_budget':ytd_budget,
        
        
         }

    return render(request,'dashboard/cumberland/bs_cumberland.html', context)


def insert_bs_advantage(request):
    if request.method == 'POST':
        print(request.POST)
        try:
            #<<<------------------ CURRENT ASSETS INSERT FUNCTION ---------------------->>>
            Descriptions = request.POST.getlist('Description[]')
            fyes = request.POST.getlist('fye[]')
            Activitys = request.POST.getlist('Activity[]')
           

            data_list_current_assets = []
            for Description, fye , Activity in zip(Descriptions, fyes,Activitys):
                if Description.strip() and fye.strip() and Activity.strip():
                    data_list_current_assets.append({
                        'Description': Description,
                        'FYE': fye,
                        'Subcategory': 'Current Assets',
                        'Category': 'Assets',
                        'Activity': Activity,
                    })

          
            cnxn = connect()
            cursor = cnxn.cursor()


            for data in data_list_current_assets:
                Description = data['Description']
                FYE = data['FYE']
                Subcategory = data['Subcategory']
                Category = data['Category']
                Activity = data['Activity']

                query = "INSERT INTO [dbo].[AscenderData_Advantage_Balancesheet] (Description, FYE, Subcategory, Category, Activity) VALUES (?, ?, ?, ?, ?)"
                cursor.execute(query, (Description, FYE, Subcategory, Category, Activity))
                cnxn.commit()

            
            #<<<------------------ Capital  ASSETS INSERT FUNCTION ---------------------->>>
            Descriptions_Capital_Assets = request.POST.getlist('Description_Capital_Assets[]')
            fyes_Capital_Assets = request.POST.getlist('fye_Capital_Assets[]')
            Activitys_Capital_Assets = request.POST.getlist('Activity_Capital_Assets[]')
           

            data_list_capital_assets = []
            for Description, fye , Activity in zip(Descriptions_Capital_Assets, fyes_Capital_Assets,Activitys_Capital_Assets):
                if Description.strip() and fye.strip() and Activity.strip():
                    data_list_capital_assets.append({
                        'Description': Description,
                        'FYE': fye,
                        'Subcategory': 'Capital Assets, Net',
                        'Category': 'Assets',
                        'Activity': Activity,
                    })

          
     

          
            for data in data_list_capital_assets:
                Description = data['Description']
                FYE = data['FYE']
                Subcategory = data['Subcategory']
                Category = data['Category']
                Activity = data['Activity']

                query = "INSERT INTO [dbo].[AscenderData_Advantage_Balancesheet] (Description, FYE, Subcategory, Category, Activity) VALUES (?, ?, ?, ?, ?)"
                cursor.execute(query, (Description, FYE, Subcategory, Category, Activity))
                cnxn.commit()

             
            #<<<------------------ CURRENT LIABILITIES INSERT FUNCTION ---------------------->>>
            Descriptions_CLs = request.POST.getlist('Description_CL[]')
            fyes_CLs = request.POST.getlist('fye_CL[]')
            Activitys_CLs = request.POST.getlist('Activity_CL[]')
           

            data_list_CL = []
            for Description, fye , Activity in zip(Descriptions_CLs, fyes_CLs, Activitys_CLs):
                if Description.strip() and fye.strip() and Activity.strip():
                    data_list_CL.append({
                        'Description': Description,
                        'FYE': fye,
                        'Subcategory': 'Current Liabilities',
                        'Category': 'Liabilities and Net Assets',
                        'Activity': Activity,
                    })

            for data in data_list_CL:
                Description = data['Description']
                FYE = data['FYE']
                Subcategory = data['Subcategory']
                Category = data['Category']
                Activity = data['Activity']

                query = "INSERT INTO [dbo].[AscenderData_Advantage_Balancesheet] (Description, FYE, Subcategory, Category, Activity) VALUES (?, ?, ?, ?, ?)"
                cursor.execute(query, (Description, FYE, Subcategory, Category, Activity))
                cnxn.commit()


            #<<<------------------ LONG TERM DEBT INSERT FUNCTION ---------------------->>>
            Descriptions_LTDs = request.POST.getlist('Description_LTD[]')
            fyes_LTDs = request.POST.getlist('fye_LTD[]')
            Activitys_LTDs = request.POST.getlist('Activity_LTD[]')
           

            data_list_LTD = []
            for Description, fye , Activity in zip(Descriptions_LTDs, fyes_LTDs, Activitys_LTDs):
                if Description.strip() and fye.strip() and Activity.strip():
                    data_list_LTD.append({
                        'Description': Description,
                        'FYE': fye,
                        'Subcategory': 'Long Term Debt',
                        'Category': 'Debt',
                        'Activity': Activity,
                    })

            for data in data_list_LTD:
                Description = data['Description']
                FYE = data['FYE']
                Subcategory = data['Subcategory']
                Category = data['Category']
                Activity = data['Activity']

                query = "INSERT INTO [dbo].[AscenderData_Advantage_Balancesheet] (Description, FYE, Subcategory, Category, Activity) VALUES (?, ?, ?, ?, ?)"
                cursor.execute(query, (Description, FYE, Subcategory, Category, Activity))
                cnxn.commit()

            
            #<<<------------------ BALANCE SHEET ACTIVITY INSERT FUNCTION ---------------------->>>
            Description_BSAs = request.POST.getlist('Description_BSA[]')
            obj_BSAs = request.POST.getlist('obj_BSA[]')
            Activity_BSAs = request.POST.getlist('Activity_BSA[]')
           

            data_list_BSA = []
            for Description, obj , Activity in zip(Description_BSAs, obj_BSAs, Activity_BSAs):
                if Description.strip() and obj.strip() and Activity.strip():
                    data_list_BSA.append({
                        'Description': Description,
                        'obj': obj,
                        'Activity': Activity,
                    })

            for data in data_list_BSA:
                Description = data['Description']
                obj = data['obj']
                Activity = data['Activity']

                query = "INSERT INTO [dbo].[AscenderData_Advantage_ActivityBS] (Description, obj, Activity) VALUES (?, ?, ?)"
                cursor.execute(query, (Description, obj, Activity))
                cnxn.commit()
            

            
            #<--------------UPDATE FOR LOCAL REVENUE,SPR,FPR----------
            updatesubs = request.POST.getlist('updatesub[]')  
            updatedescs = request.POST.getlist('updatedesc[]')
            updatefyes = request.POST.getlist('updatefye[]')
            
            

            updatedata_bs = []

            for updatesub,updatedesc,updatefye in zip(updatesubs, updatedescs,updatefyes):
                if updatesub.strip() and updatedesc.strip()  and updatefye.strip():
                    updatedata_bs.append({
                        'updatesub': updatesub,
                        'updatedesc':updatedesc,
                        
                        'updatefye': updatefye,
                        
                    })
            for data in updatedata_bs:
                updatesub= data['updatesub']
                updatedesc=data['updatedesc']
                
                updatefye = data['updatefye']

                try:
                    query = "UPDATE [dbo].[AscenderData_Advantage_Balancesheet] SET FYE = ? WHERE Description = ? and Subcategory = ? "
                    cursor.execute(query, (updatefye, updatedesc,updatesub))
                    cnxn.commit()
                    print(f"Rows affected for fund={updatefye}: {cursor.rowcount}")
                except Exception as e:
                    print(f"Error updating fund={updatefye}: {str(e)}")
            
            


            return redirect('bs_advantage')

        except Exception as e:
            return JsonResponse({'status': 'error', 'message': str(e)})

    return JsonResponse({'status': 'error', 'message': 'Invalid request method'})



def viewgl_activitybs(request,obj,yr):
    print(request)
    try:
        
        cnxn = connect()
        cursor = cnxn.cursor()

        
        query = "SELECT * FROM [dbo].[AscenderData_Advantage] WHERE obj = ? and AcctPer = ?"
        cursor.execute(query, (obj,yr))
        
        rows = cursor.fetchall()
    
        glbs_data=[]
    
    
        for row in rows:
            date_str=row[11]
        
           
            bal = float(row[18]) if row[18] else 0
            if bal == 0:
                balformat = ""
            else:
                balformat = "{:,.0f}".format(abs(bal)) if bal >= 0 else "({:,.0f})".format(abs(bal))


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
                'Bal':balformat,
                'WorkDescr':row[19],
                'Type':row[20],
                'Contr':row[21]
            }

            glbs_data.append(row_dict)

        total_expend = 0 
        for row in glbs_data:
            expend_str = row['Bal'].replace(',','').replace('(','-').replace(')','')
            try:
                expend_value = float(expend_str)
                total_expend += expend_value
                
            except ValueError:
                pass
            
        
        

        # total_bal = sum(row['Bal'] for row in glbs_data)
        total_bal = "{:,}".format(total_expend)
        
        context = { 
            'glbs_data':glbs_data,
            'total_bal':total_bal,

            }

        
        cursor.close()
        cnxn.close()

        return JsonResponse({'status': 'success', 'data': context})

    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})

def viewgl_activitybs_cumberland(request,obj,yr):
    print(request)
    try:
        
        cnxn = connect()
        cursor = cnxn.cursor()

        
        query = "SELECT * FROM [dbo].[AscenderData_Cumberland] WHERE obj = ? and AcctPer = ?"
        cursor.execute(query, (obj,yr))
        
        rows = cursor.fetchall()
    
        glbs_data=[]
    
    
        for row in rows:
            date_str=row[11]
        
            

            bal = float(row[18]) if row[18] else 0
            if bal == 0:
                balformat = ""
            else:
                balformat = "{:,.0f}".format(abs(bal)) if bal >= 0 else "({:,.0f})".format(abs(bal))


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
                'Bal':balformat,
                'WorkDescr':row[19],
                'Type':row[20],
                'Contr':row[21]
            }

            glbs_data.append(row_dict)

        total_expend = 0 
        for row in glbs_data:
            expend_str = row['Bal'].replace(',','').replace('(','-').replace(')','')
            try:
                expend_value = float(expend_str)
                total_expend += expend_value
                
            except ValueError:
                pass
            
        
        

        # total_bal = sum(row['Bal'] for row in glbs_data)
        total_bal = "{:,}".format(total_expend)
        
        context = { 
            'glbs_data':glbs_data,
            'total_bal':total_bal,

            }

        
        cursor.close()
        cnxn.close()

        return JsonResponse({'status': 'success', 'data': context})

    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})

def viewglfunc(request,func,yr):
    print(request)
    try:
        
        cnxn = connect()
        cursor = cnxn.cursor()

        
        query = "SELECT * FROM [dbo].[AscenderData_Advantage] WHERE func = ? and AcctPer = ?"
        cursor.execute(query, (func,yr))
        
        rows = cursor.fetchall()
    
        glfunc_data=[]
    
    
        for row in rows:
            date_str=row[11]
        
            
            expend = float(row[17]) if row[17] else 0
            if expend == 0:
                expendformat = ""
            else:
                expendformat = "{:,.0f}".format(abs(expend)) if expend >= 0 else "({:,.0f})".format(abs(expend))

            
            
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
                'Expend':expendformat,
                'Bal':row[18],
                'WorkDescr':row[19],
                'Type':row[20],
                'Contr':row[21]
            }

            glfunc_data.append(row_dict)



        total_expend = 0 
        for row in glfunc_data:
            expend_str = row['Expend'].replace(',','').replace('(','-').replace(')','')
            try:
                expend_value = float(expend_str)
                total_expend += expend_value
                
            except ValueError:
                pass
            
        
        
        # total_bal = sum(float(row['Expend'].replace(',','')) for row in glfunc_data)
        total_bal = "{:,}".format(total_expend)
        
       
        context = { 
            'glfunc_data':glfunc_data,
            'total_bal':total_bal
            }

        
        cursor.close()
        cnxn.close()

        return JsonResponse({'status': 'success', 'data': context})

    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})

def viewglfunc_cumberland(request,func,yr):
    print(request)
    try:
        
        cnxn = connect()
        cursor = cnxn.cursor()

        
        query = "SELECT * FROM [dbo].[AscenderData_Cumberland] WHERE func = ? and AcctPer = ?"
        cursor.execute(query, (func,yr))
        
        rows = cursor.fetchall()
    
        glfunc_data=[]
    
    
        for row in rows:
            date_str=row[11]
        
            
            expend = float(row[17]) if row[17] else 0
            if expend == 0:
                expendformat = ""
            else:
                expendformat = "{:,.0f}".format(abs(expend)) if expend >= 0 else "({:,.0f})".format(abs(expend))

            
            
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
                'Expend':expendformat,
                'Bal':row[18],
                'WorkDescr':row[19],
                'Type':row[20],
                'Contr':row[21]
            }

            glfunc_data.append(row_dict)



        total_expend = 0 
        for row in glfunc_data:
            expend_str = row['Expend'].replace(',','').replace('(','-').replace(')','')
            try:
                expend_value = float(expend_str)
                total_expend += expend_value
                
            except ValueError:
                pass
            
        
        
        # total_bal = sum(float(row['Expend'].replace(',','')) for row in glfunc_data)
        total_bal = "{:,}".format(total_expend)
        
       
        context = { 
            'glfunc_data':glfunc_data,
            'total_bal':total_bal
            }

        
        cursor.close()
        cnxn.close()

        return JsonResponse({'status': 'success', 'data': context})

    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})

def viewglexpense(request,obj,yr):
    print(request)
    try:
        
        cnxn = connect()
        cursor = cnxn.cursor()
        

        cursor.execute("SELECT * FROM [dbo].[AscenderData_Advantage_PL_ExpensesbyObjectCode];") 
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

        cursor.execute("SELECT * FROM [dbo].[AscenderData_Advantage_PL_Activities];") 
        rows = cursor.fetchall()

        data_activities=[]


        for row in rows:

        
            row_dict = {
                'obj':row[0],
                'Description':row[1],
                'Category':row[2],

                }

            data_activities.append(row_dict)
        


            
        
        
        query = "SELECT * FROM [dbo].[AscenderData_Advantage] WHERE obj = ? and AcctPer = ? "
        cursor.execute(query, (obj,yr))
        
        rows = cursor.fetchall()
    
        glfunc_data=[]
    
    
        for row in rows:
            date_str=row[11]
        
            
            expend = float(row[17]) if row[17] else 0
            if expend == 0:
                expendformat = ""
            else:
                expendformat = "{:,.0f}".format(abs(expend)) if expend >= 0 else "({:,.0f})".format(abs(expend))

            
            
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
                'Expend':expendformat,
                'Bal':row[18],
                'WorkDescr':row[19],
                'Type':row[20],
                'Contr':row[21]
            }

            glfunc_data.append(row_dict)



        total_expend = 0 
        for row in glfunc_data:
            expend_str = row['Expend'].replace(',','').replace('(','-').replace(')','')
            try:
                expend_value = float(expend_str)
                total_expend += expend_value
                
            except ValueError:
                pass
            
        
        
        # total_bal = sum(float(row['Expend'].replace(',','')) for row in glfunc_data)
        total_bal = "{:,}".format(total_expend)
        
       
        context = { 
            'glfunc_data':glfunc_data,
            'total_bal':total_bal
            }

        
        cursor.close()
        cnxn.close()

        return JsonResponse({'status': 'success', 'data': context})

    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})

def viewglexpense_cumberland(request,obj,yr):
    print(request)
    try:
        
        cnxn = connect()
        cursor = cnxn.cursor()
        

        cursor.execute("SELECT * FROM [dbo].[AscenderData_Cumberland_PL_ExpensesbyObjectCode];") 
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

        cursor.execute("SELECT * FROM [dbo].[AscenderData_Cumberland_PL_Activities];") 
        rows = cursor.fetchall()

        data_activities=[]


        for row in rows:

        
            row_dict = {
                'obj':row[0],
                'Description':row[1],
                'Category':row[2],

                }

            data_activities.append(row_dict)
        


            
        
        
        query = "SELECT * FROM [dbo].[AscenderData_Cumberland] WHERE obj = ? and AcctPer = ? "
        cursor.execute(query, (obj,yr))
        
        rows = cursor.fetchall()
    
        glfunc_data=[]
    
    
        for row in rows:
            date_str=row[11]
        
           
            expend = float(row[17]) if row[17] else 0
            if expend == 0:
                expendformat = ""
            else:
                expendformat = "{:,.0f}".format(abs(expend)) if expend >= 0 else "({:,.0f})".format(abs(expend))

            
            
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
                'Expend':expendformat,
                'Bal':row[18],
                'WorkDescr':row[19],
                'Type':row[20],
                'Contr':row[21]
            }

            glfunc_data.append(row_dict)



        total_expend = 0 
        for row in glfunc_data:
            expend_str = row['Expend'].replace(',','').replace('(','-').replace(')','')
            try:
                expend_value = float(expend_str)
                total_expend += expend_value
                
            except ValueError:
                pass
            
        
        
        total_bal = sum(float(row['Expend'].replace(',','')) for row in glfunc_data)
        total_bal = "{:,}".format(total_expend)
        
       
        context = { 
            'glfunc_data':glfunc_data,
            'total_bal':total_bal
            }

        
        cursor.close()
        cnxn.close()

        return JsonResponse({'status': 'success', 'data': context})

    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})

        
def cashflow_advantage(request):
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Advantage_Definition_obj];") 
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

    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Advantage_Definition_func];") 
    rows = cursor.fetchall()


    data2=[]
    for row in rows:
        budgetformat = "{:,.0f}".format(float(row[3])) if row[3] else ""
        row_dict = {
            'func_func': row[0],
            'desc': row[1],
            'budget': budgetformat,
            
        }
        data2.append(row_dict)



    #
    cursor.execute("SELECT * FROM [dbo].[AscenderData_Advantage];") 
    rows = cursor.fetchall()
    
    data3=[]
    
    
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


    cursor.execute("SELECT * FROM [dbo].[AscenderData_Advantage_PL_ExpensesbyObjectCode];") 
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

    cursor.execute("SELECT * FROM [dbo].[AscenderData_Advantage_PL_Activities];") 
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
    acct_per_values_expense = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
    for item in data_activities:
        
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values_expense, start=1):
            item[f'total_activities{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
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
    
        for i, acct_per in enumerate(acct_per_values2, start=1):
            item[f'total_func2_{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['func'] == func and entry['AcctPer'] == acct_per and entry['obj'] == '6449'
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
    print(formatted_total_netsurplus)
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
          'total_DnA': formatted_total_DnA,
          'total_netsurplus':formatted_total_netsurplus,
          'total_SBD':total_SBD,
          
        
          

          }
    return render(request,'dashboard/advantage/cashflow_advantage.html',context)

def cashflow_cumberland(request):
    return render(request,'dashboard/cumberland/cashflow_cumberland.html')


