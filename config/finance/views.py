from django.shortcuts import render,redirect,get_object_or_404
from .forms import UploadForm, ReportsForm
import pandas as pd
import io
from .models import User, Item
from django.contrib import auth
from django.contrib.auth import login, authenticate
from django.contrib.auth import logout
from django.urls import reverse_lazy
from django.utils.safestring import mark_safe
import sys
from time import strftime
import os
import pyodbc
import sqlalchemy as sa
from sqlalchemy import create_engine
from urllib.parse import quote_plus
import datetime
import locale
from django.http import JsonResponse,HttpResponse
from django.views.decorators.csrf import csrf_exempt
import json
from datetime import datetime,timedelta
from dateutil.relativedelta import relativedelta
import shutil
import openpyxl
from django.conf import settings
from openpyxl.utils import get_column_letter
from bs4 import BeautifulSoup
from openpyxl.styles import Font,NamedStyle, Border, Side, Alignment
from .connect import connect





def loginView(request):
    if request.method == 'POST':
        username = request.POST['username']
        password = request.POST['password']
        user = authenticate(username=username, password=password)
        if user is not None and user.is_active:
            if user.is_admin or user.is_superuser:
                auth.login(request, user)
                return redirect('dashboard/advantage')
            elif user is not None and user.is_employee:
                auth.login(request, user)
                return redirect('dashboard/advantage')
            
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
    ytd_budget_test = last_month_number + 3 
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

def pl_villagetech(request):
    
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
    cursor.execute("SELECT * FROM [dbo].[Skyward_VillageTech];") 
    rows = cursor.fetchall()
    
    data3=[]
    
    
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
    

    # #---------- FOR EXPENSE TOTAL -------
    acct_per_values_expense = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
    for item in data_activities:
        
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values_expense, start=1):
            item[f'total_activities{i}'] = sum(
                entry['Amount'] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
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
                entry['Amount'] for entry in data3 if entry['fund'] == fund and entry['obj'] == obj and entry['AcctPer'] == acct_per
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
                entry['Amount'] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per
            )

    for item in data2:
        func = item['func_func']
        

        for i, acct_per in enumerate(acct_per_values2, start=1):
            item[f'total_func2_{i}'] = sum(
                entry['Amount'] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per and entry['obj'] == '6449'
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
                
                



    # lr_funds = list(set(row['fund'] for row in data3 if 'fund' in row))
    # lr_funds_sorted = sorted(lr_funds)
    # lr_obj = list(set(row['obj'] for row in data3 if 'obj' in row))
    # lr_obj_sorted = sorted(lr_obj)

    # func_choice = list(set(row['func'] for row in data3 if 'func' in row))
    # func_choice_sorted = sorted(func_choice)
    
            
    current_date = datetime.today().date()
    current_year = current_date.year
    last_year = current_date - timedelta(days=365)
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 3 
    ytd_budget = ytd_budget_test / 12
    formatted_ytd_budget = f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places

    if formatted_ytd_budget.startswith("0."):
        formatted_ytd_budget = formatted_ytd_budget[2:]
   
    context = {
         'data': data, 
         'data2':data2 , 
         'data3': data3 ,
        #   'lr_funds':lr_funds_sorted, 
        #   'lr_obj':lr_obj_sorted, 
        #   'func_choice':func_choice_sorted ,
          'data_expensebyobject': data_expensebyobject,
          'data_activities': data_activities,
          'last_month':last_month,
          'last_month_number':last_month_number,
          'format_ytd_budget': formatted_ytd_budget,
          'ytd_budget':ytd_budget,
          

          }
    return render(request,'dashboard/villagetech/pl_villagetech.html', context)

def first_advantage(request):
    current_date = datetime.today().date()
    current_year = current_date.year
    last_year = current_date - timedelta(days=365)
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 3 
    ytd_budget = ytd_budget_test / 12
    formatted_ytd_budget = f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
    
    if formatted_ytd_budget.startswith("0."):
        formatted_ytd_budget = formatted_ytd_budget[2:]
    context = {
        'last_month':last_month,
        'last_month_number':last_month_number,
        'format_ytd_budget': formatted_ytd_budget,
        'ytd_budget':ytd_budget,
    }
    return render(request,'dashboard/advantage/first_advantage.html', context)

def dashboard_advantage(request):
    data = {"accomplishments":"", "activities":"", "agendas": ""}
    delimiter = "\n"

    cnxn = connect()
    cursor = cnxn.cursor()
    school_name = 'advantage'

    if request.method == 'POST':
        form = ReportsForm(request.POST)
        activities = request.POST.get('activities')
        accomplishments = request.POST.get('accomplishments')
        agendas = request.POST.get('agendas')

         
        update_query = "UPDATE [dbo].[Report] SET accomplishments = ?, activities = ?, agendas = ? WHERE school = ?"
        cursor.execute(update_query, (accomplishments, activities, agendas, school_name))
        # update_query = "UPDATE [dbo].[Report] SET accomplishments = ?, activities = ? WHERE school = ?"
        # cursor.execute(update_query, (accomplishments, activities, school_name))
        cnxn.commit()

        data["accomplishments"] = mark_safe(accomplishments)
        data["activities"] = mark_safe(activities)
        data["agendas"] = mark_safe(activities)
    else:
        # check if it exists
        # query for the school
        query = "SELECT * FROM [dbo].[Report] WHERE school = ?"
        cursor.execute(query, school_name)
        row = cursor.fetchone()

        if row is None:
            # Insert query if it does noes exists

            insert_query = "INSERT INTO [dbo].[Report] (school, accomplishments, activities, agendas) VALUES (?, ?, ?, ?)"
            accomplishments = "No accomplishments for this school yet. Click edit and add bullet points. It is important that the inserted accomplishments are in bullet points.\n"
            activities = "No activities for this school yet. Click edit and add bullet points. It is important that the inserted activities are in bullet points.\n"
            agendas = "No agenda for this school yet. Click edit and add bullet points. It is important that the inserted agenda are in bullet points.\n"

            # Execute the INSERT query
            cursor.execute(insert_query, (school_name, accomplishments, activities, agendas))
            # cursor.execute(insert_query, (school_name, accomplishments, activities))

            # Commit the transaction
            cnxn.commit()
            print("Row inserted successfully.")
        else:
            # row = (school, accomplishments, activities)
            if row[1]:
                data["accomplishments"] = mark_safe(row[1])
            if row[2]:
                data["activities"] = mark_safe(row[2])
            try:
                if row[3]:
                    data["agendas"] = mark_safe(row[3])
            except:
                pass

    # form = CKEditorForm(initial={'form_field_name': initial_content})
    form = ReportsForm(initial={'accomplishments': data["accomplishments"], 'activities': data["activities"], "agendas": data["agendas"]})

    cursor.close()
    cnxn.close()
    return render(request,'dashboard/advantage/dashboard_advantage.html', {'form': form, "data":data})


def dashboard_cumberland(request):
    data = {"accomplishments":"", "activities":""}
    delimiter = "\n"

    cnxn = connect()
    cursor = cnxn.cursor()
    school_name = 'cumberland'

    if request.method == 'POST':
        form = ReportsForm(request.POST)
        activities_text = request.POST.get('activities')
        accomplishments_text = request.POST.get('accomplishments')

        # Parse the HTML content
        soup_activities = BeautifulSoup(activities_text, 'html.parser')
        soup_accomplishments = BeautifulSoup(accomplishments_text, 'html.parser')

        # Find all list items within the <ul> tag
        activities_items = soup_activities.find_all('li')
        accomplishments_items = soup_accomplishments.find_all('li')

        # Extract the text content from each list item
        activities_list = [item.get_text()+delimiter for item in activities_items]
        accomplishments_list = [item.get_text()+delimiter for item in accomplishments_items]

        activities = "".join(activities_list)
        accomplishments = "".join(accomplishments_list)

        update_query = "UPDATE [dbo].[Report] SET accomplishments = ?, activities = ? WHERE school = ?"
        cursor.execute(update_query, (accomplishments, activities, school_name))
        cnxn.commit()

        data["accomplishments"] = accomplishments.split(delimiter)[:-1]
        data["activities"] = activities.split(delimiter)[:-1]
    else:
        # check if it exists
        # query for the school
        query = "SELECT * FROM [dbo].[Report] WHERE school = ?"
        cursor.execute(query, school_name)
        row = cursor.fetchone()

        if row is None:
            # Insert query if it does noes exists
            insert_query = "INSERT INTO [dbo].[Report] (school, accomplishments, activities) VALUES (?, ?, ?)"
            accomplishments = "No accomplishments for this school yet. Click edit and add bullet points. It is important that the inserted accomplishments are in bullet points.\n"
            activities = "No activities for this school yet. Click edit and add bullet points. It is important that the inserted activities are in bullet points.\n"

            # Execute the INSERT query
            cursor.execute(insert_query, (school_name, accomplishments, activities))

            # Commit the transaction
            cnxn.commit()
            print("Row inserted successfully.")
        else:
            # row = (school, accomplishments, activities)
            if row[1]:
                data["accomplishments"] = row[1].split(delimiter)[:-1]
            if row[2]:
                data["activities"] = row[2].split(delimiter)[:-1]

    activities = "</li>".join(["<li>" + w for w in data["activities"]])
    accomplishments = "</li>" .join(["<li>" + w for w in data["accomplishments"]])
    # form = CKEditorForm(initial={'form_field_name': initial_content})
    form = ReportsForm(initial={'accomplishments': accomplishments, 'activities': activities})

    cursor.close()
    cnxn.close()
    return render(request,'dashboard/cumberland/dashboard_cumberland.html', {'form': form, "data":data})

def first_cumberland(request):
    current_date = datetime.today().date()
    current_year = current_date.year
    last_year = current_date - timedelta(days=365)
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 3 
    ytd_budget = ytd_budget_test / 12
    formatted_ytd_budget = f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
    
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
    ytd_budget_test = last_month_number + 3 
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

def first_advantagechart(request):
    return render(request,'dashboard/advantage/first_advantagechart.html')

def pl_advantagechart(request):
    return render(request,'dashboard/advantage/pl_advantagechart.html')

def pl_cumberlandchart(request):
    return render(request,'dashboard/cumberland/pl_cumberlandchart.html')

def viewgl(request,fund,obj,yr):
    
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
    ytd_budget_test = last_month_number + 3 
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


def bs_villagetech(request):
    
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

    cursor.execute("SELECT * FROM [dbo].[Skyward_VillageTech]") 
    rows = cursor.fetchall()
    
    data3=[]
    
    
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

    for item in data_activitybs:
        
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_bal{i}'] = sum(
                entry['Amount'] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
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
    ytd_budget_test = last_month_number + 3
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

    return render(request,'dashboard/villagetech/bs_villagetech.html', context)


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
    ytd_budget_test = last_month_number + 3
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
    ytd_budget_test = last_month_number + 3
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

def generate_excel(request):
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Advantage_Definition_obj];") 
    rows = cursor.fetchall()

    
    data = []
    for row in rows:
        if row[4] is None:
            row[4] = ''
       
        
        row_dict = {
            'fund': row[0],
            'obj': row[1],
            'description': row[2],
            'category': row[3],
            'value': row[4]
        }
        data.append(row_dict)

    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Advantage_Definition_func];") 
    rows = cursor.fetchall()


    data2=[]
    for row in rows:
       
        row_dict = {
            'func_func': row[0],
            'desc': row[1],
            'budget': row[3],
            
        }
        data2.append(row_dict)


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
        
        
        row_dict = {
            'obj':row[0],
            'Description':row[1],
            'budget':row[2],
            
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
    #----------------------------------------END OF PL DATA


    #----------------- BS DATA
    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Advantage_Balancesheet]") 
    rows = cursor.fetchall()
    
    data_balancesheet=[]
    
    
    for row in rows:
       
        row_dict = {
            'Activity':row[0],
            'Description':row[1],
            'Category':row[2],
            'Subcategory':row[3],
            'FYE':row[4],
            
            
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

    

    acct_per_values = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']
    
    #---------- ADDITIONAL PL DATAS
    for item in data_activities:
        
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_activities{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
            )
    for item in data:
        fund = item['fund']
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_real{i}'] = sum(
                entry['Real'] for entry in data3 if entry['fund'] == fund and entry['obj'] == obj and entry['AcctPer'] == acct_per
            )


    for item in data2:
        func = item['func_func']
        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_func{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per
            )

    for item in data2:
        func = item['func_func']
        

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_func2_{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per and entry['obj'] == '6449'
            )  
    # END OF ADDITIONAL PL DATAS


    # BS ADDITIONAL DATAS
    for item in data_activitybs:
        
        obj = item['obj']

        for i, acct_per in enumerate(acct_per_values, start=1):
            item[f'total_bal{i}'] = sum(
                entry['Bal'] for entry in data3 if entry['obj'] == obj and entry['AcctPer'] == acct_per
            )

    keys_to_check = ['total_bal1', 'total_bal2', 'total_bal3', 'total_bal4', 'total_bal5','total_bal6','total_bal7','total_bal8','total_bal9','total_bal10','total_bal11','total_bal12']
  
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
            row[f'total_sum{i}'] = activity_sum_dict.get(key, 0)
    

    
    for row in data_balancesheet:
    
        FYE_value = int(row['FYE']) if row['FYE'] else 0
        total_sum9_value = int(row['total_sum9']) if row['total_sum9'] else 0
        total_sum10_value = int(row['total_sum10']) if row['total_sum10'] else 0
        total_sum11_value = int(row['total_sum11']) if row['total_sum11'] else 0
        total_sum12_value = int(row['total_sum12']) if row['total_sum12'] else 0
        total_sum1_value = int(row['total_sum1']) if row['total_sum1'] else 0
        total_sum2_value = int(row['total_sum2']) if row['total_sum2'] else 0
        total_sum3_value = int(row['total_sum3']) if row['total_sum3'] else 0
        total_sum4_value = int(row['total_sum4']) if row['total_sum4'] else 0
        total_sum5_value = int(row['total_sum5']) if row['total_sum5'] else 0
        total_sum6_value = int(row['total_sum6']) if row['total_sum6'] else 0
        total_sum7_value = int(row['total_sum7']) if row['total_sum7'] else 0
        total_sum8_value = int(row['total_sum8']) if row['total_sum8'] else 0
    
        # Calculate the differences and store them in the row dictionary
        row['difference_9'] = float(FYE_value + total_sum9_value)
        row['difference_10'] = float(FYE_value + total_sum9_value + total_sum10_value)
        row['difference_11'] = float(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value)
        row['difference_12'] = float(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value)
        row['difference_1'] = float(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value)
        row['difference_2'] = float(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value)
        row['difference_3'] = float(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value)
        row['difference_4'] = float(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value)
        row['difference_5'] = float(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value)
        row['difference_6'] = float(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value)
        row['difference_7'] = float(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value)
        row['difference_8'] = float(FYE_value + total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value)
    
        row['fytd'] = float(total_sum9_value + total_sum10_value + total_sum11_value + total_sum12_value + total_sum1_value + total_sum2_value + total_sum3_value + total_sum4_value + total_sum5_value + total_sum6_value + total_sum7_value + total_sum8_value)
    
    template_path = os.path.join(settings.BASE_DIR, 'finance', 'static', 'template.xlsx')


    generated_excel_path = os.path.join(settings.BASE_DIR, 'finance', 'static', 'GeneratedExcel.xlsx')



    shutil.copyfile(template_path, generated_excel_path)

 
    workbook = openpyxl.load_workbook(generated_excel_path)
    sheet_names = workbook.sheetnames
    fontbold = Font(bold=True)


    first = sheet_names[0]
    pl = sheet_names[1]
    bs = sheet_names[2]
    cashflow= sheet_names[3]
    

    first_sheet = workbook[first]
    pl_sheet = workbook[pl]
    bs_sheet = workbook[bs]
    cashflow_sheet = workbook[cashflow]

    #------ FIRST DESIGN
    first_sheet.column_dimensions['A'].width = 46
    first_sheet.column_dimensions['B'].width = 20
    first_sheet.column_dimensions['C'].width = 9
    first_sheet.column_dimensions['D'].width = 12
    first_sheet.column_dimensions['E'].width = 38
    for row in range(3,26):
        first_sheet.row_dimensions[row].height = 30
    first_sheet.row_dimensions[1].height = 21
    first_sheet.row_dimensions[1].height = 21
  
    

    #------- PL DESIGN
    for col in range(7, 19):
        col_letter = get_column_letter(col)
        pl_sheet.column_dimensions[col_letter].outline_level = 1
        pl_sheet.column_dimensions[col_letter].hidden = True

    pl_sheet.row_dimensions[1].height = 64
    for row in range(2,181):
        pl_sheet.row_dimensions[row].height = 19
    pl_sheet.row_dimensions[17].height = 26 #local revenue
    pl_sheet.row_dimensions[20].height = 26 #spr
    pl_sheet.row_dimensions[33].height = 26 #fpr
    pl_sheet.row_dimensions[34].height = 26 #total
    pl_sheet.column_dimensions['A'].width = 8
    pl_sheet.column_dimensions['B'].width = 28
    pl_sheet.column_dimensions['C'].width = 10
    pl_sheet.column_dimensions['D'].width = 14
    pl_sheet.column_dimensions['E'].width = 14
    pl_sheet.column_dimensions['F'].hidden = True
    pl_sheet.column_dimensions['G'].width = 12
    pl_sheet.column_dimensions['H'].width = 12
    pl_sheet.column_dimensions['I'].width = 12
    pl_sheet.column_dimensions['J'].width = 12
    pl_sheet.column_dimensions['K'].width = 12
    pl_sheet.column_dimensions['L'].width = 12
    pl_sheet.column_dimensions['M'].width = 12
    pl_sheet.column_dimensions['N'].width = 12
    pl_sheet.column_dimensions['O'].width = 12
    pl_sheet.column_dimensions['P'].width = 12
    pl_sheet.column_dimensions['Q'].width = 12
    pl_sheet.column_dimensions['R'].width = 12
    pl_sheet.column_dimensions['S'].width = 3
    pl_sheet.column_dimensions['T'].width = 17
    pl_sheet.column_dimensions['U'].width = 17
    pl_sheet.column_dimensions['V'].width = 12

   

    total_vars = ['value', 'total_real9', 'total_real10', 'total_real11', 'total_real12', 
              'total_real1', 'total_real2', 'total_real3', 'total_real4', 'total_real5', 
              'total_real6', 'total_real7', 'total_real8']
    totals = {var: 0 for var in total_vars}

    # PL START OF DESIGN
    start_row = 5
    lr_row_start = start_row
    for row_data in data:
        if row_data['category'] == 'Local Revenue': 
            pl_sheet[f'A{start_row}'] = row_data['fund']
            pl_sheet[f'B{start_row}'] = f'{row_data["obj"]} - {row_data["description"]}'
            pl_sheet[f'D{start_row}'] = row_data['value']
            pl_sheet[f'E{start_row}'] = row_data['value'] * .922222
            
            pl_sheet[f'G{start_row}'] = row_data['total_real9']
            pl_sheet[f'H{start_row}'] = row_data['total_real10']
            pl_sheet[f'I{start_row}'] = row_data['total_real11']
            pl_sheet[f'J{start_row}'] = row_data['total_real12']
            pl_sheet[f'K{start_row}'] = row_data['total_real1']
            pl_sheet[f'L{start_row}'] = row_data['total_real2']
            pl_sheet[f'M{start_row}'] = row_data['total_real3']
            pl_sheet[f'N{start_row}'] = row_data['total_real4']
            pl_sheet[f'O{start_row}'] = row_data['total_real5']
            pl_sheet[f'P{start_row}'] = row_data['total_real6']
            pl_sheet[f'Q{start_row}'] = row_data['total_real7']
            pl_sheet[f'R{start_row}'] = row_data['total_real8']
            pl_sheet[f'T{start_row}'].value = f'=SUM(G{start_row}:R{start_row})' 
            lr_row_end = start_row
            
            for var in total_vars:
                totals[var] += row_data.get(var, 0)
            
            start_row += 1
            
    lr_end = start_row
    #local revenue total
    for col in range(2, 19):  
        cell = pl_sheet.cell(row=start_row, column=col)
        cell.font = fontbold
    pl_sheet[f'B{start_row}'] = 'Local Revenue'
    pl_sheet[f'D{start_row}'] = totals['value']
    pl_sheet[f'E{start_row}'].value = f'=SUM(E{lr_row_start}:E{lr_row_end})'  
    pl_sheet[f'G{start_row}'] = totals['total_real9']
    pl_sheet[f'H{start_row}'] = totals['total_real10']
    pl_sheet[f'I{start_row}'] = totals['total_real11']
    pl_sheet[f'J{start_row}'] = totals['total_real12']
    pl_sheet[f'K{start_row}'] = totals['total_real1']
    pl_sheet[f'L{start_row}'] = totals['total_real2']
    pl_sheet[f'M{start_row}'] = totals['total_real3']
    pl_sheet[f'N{start_row}'] = totals['total_real4']
    pl_sheet[f'O{start_row}'] = totals['total_real5']
    pl_sheet[f'P{start_row}'] = totals['total_real6']
    pl_sheet[f'Q{start_row}'] = totals['total_real7']
    pl_sheet[f'R{start_row}'] = totals['total_real8']
    pl_sheet[f'T{start_row}'].value = f'=SUM(T{lr_row_start}:T{lr_row_end})'  
    pl_sheet[f'U{start_row}'].value = f'=+T{start_row}-E{start_row}'
    pl_sheet[f'V{start_row}'].value = f'=+T{start_row}/D{start_row}'
    

    start_row += 1  

    
    

    # for row in total_vars:
    #     globals()[row] = 0
    totals = {var: 0 for var in total_vars} # reset the totals
    spr_row_start = start_row
    for row_data in data:
        if row_data['category'] == 'State Program Revenue':
        
            pl_sheet[f'A{start_row}'] = row_data['fund']
            pl_sheet[f'B{start_row}'] = f'{row_data["obj"]} - {row_data["description"]}'
            pl_sheet[f'D{start_row}'] = row_data['value']
            pl_sheet[f'E{start_row}'] = row_data['value'] * .922222
            pl_sheet[f'G{start_row}'] = row_data['total_real9']
            pl_sheet[f'H{start_row}'] = row_data['total_real10']
            pl_sheet[f'I{start_row}'] = row_data['total_real11']        
            pl_sheet[f'J{start_row}'] = row_data['total_real12']
            pl_sheet[f'K{start_row}'] = row_data['total_real1']
            pl_sheet[f'L{start_row}'] = row_data['total_real2']
            pl_sheet[f'M{start_row}'] = row_data['total_real3']
            pl_sheet[f'N{start_row}'] = row_data['total_real4']
            pl_sheet[f'O{start_row}'] = row_data['total_real5']
            pl_sheet[f'P{start_row}'] = row_data['total_real6']
            pl_sheet[f'Q{start_row}'] = row_data['total_real7']
            pl_sheet[f'R{start_row}'] = row_data['total_real8']
            pl_sheet[f'T{start_row}'].value = f'=SUM(G{start_row}:R{start_row})' 
             
            pl_sheet[f'U{start_row}'].value = f'=+T{start_row}-E{start_row}'
            spr_row_end = start_row

            for var in total_vars:
                totals[var] += row_data.get(var, 0)
            
            start_row += 1
            
            
    spr_end = start_row
    # STATE PROGRAM TOTAL      
    for col in range(2, 19):  
        cell = pl_sheet.cell(row=start_row, column=col)
        cell.font = fontbold
    pl_sheet[f'B{start_row}'] = 'State Program Revenue'
    pl_sheet[f'D{start_row}'] = totals['value']
    pl_sheet[f'E{start_row}'].value = f'=SUM(E{spr_row_start}:E{spr_row_end})' 
    pl_sheet[f'G{start_row}'] = totals['total_real9']
    pl_sheet[f'H{start_row}'] = totals['total_real10']
    pl_sheet[f'I{start_row}'] = totals['total_real11']
    pl_sheet[f'J{start_row}'] = totals['total_real12']
    pl_sheet[f'K{start_row}'] = totals['total_real1']
    pl_sheet[f'L{start_row}'] = totals['total_real2']
    pl_sheet[f'M{start_row}'] = totals['total_real3']
    pl_sheet[f'N{start_row}'] = totals['total_real4']
    pl_sheet[f'O{start_row}'] = totals['total_real5']
    pl_sheet[f'P{start_row}'] = totals['total_real6']
    pl_sheet[f'Q{start_row}'] = totals['total_real7']
    pl_sheet[f'R{start_row}'] = totals['total_real8']
    pl_sheet[f'T{start_row}'].value = f'=SUM(T{spr_row_start}:T{spr_row_end})'  
    pl_sheet[f'U{start_row}'].value = f'=+T{start_row}-E{start_row}'  
    pl_sheet[f'V{start_row}'].value = f'=+T{start_row}/D{start_row}' 
    start_row += 1

    totals = {var: 0 for var in total_vars} # reset the totals
    fpr_row_start = start_row
    for row_data in data:
        if row_data['category'] == 'Federal Program Revenue':
        
            pl_sheet[f'A{start_row}'] = row_data['fund']
            pl_sheet[f'B{start_row}'] = f'{row_data["obj"]} - {row_data["description"]}'
            pl_sheet[f'D{start_row}'] = row_data['value']
            pl_sheet[f'E{start_row}'] = row_data['value'] * .922222
            pl_sheet[f'G{start_row}'] = row_data['total_real9']
            pl_sheet[f'H{start_row}'] = row_data['total_real10']
            pl_sheet[f'I{start_row}'] = row_data['total_real11']
            pl_sheet[f'J{start_row}'] = row_data['total_real12']
            pl_sheet[f'K{start_row}'] = row_data['total_real1']
            pl_sheet[f'L{start_row}'] = row_data['total_real2']
            pl_sheet[f'M{start_row}'] = row_data['total_real3']
            pl_sheet[f'N{start_row}'] = row_data['total_real4']
            pl_sheet[f'O{start_row}'] = row_data['total_real5']
            pl_sheet[f'P{start_row}'] = row_data['total_real6']
            pl_sheet[f'Q{start_row}'] = row_data['total_real7']
            pl_sheet[f'R{start_row}'] = row_data['total_real8']
            pl_sheet[f'T{start_row}'].value = f'=SUM(G{start_row}:R{start_row})'
            pl_sheet[f'U{start_row}'].value = f'=+T{start_row}-E{start_row}' 
            fpr_row_end = start_row
            for var in total_vars:
                totals[var] += row_data.get(var, 0)
            
            start_row += 1
            
    fpr_end = start_row
        # FEDERAL PROGRAM REVENUE TOTAL
    for col in range(2, 19):  
        cell = pl_sheet.cell(row=start_row, column=col)
        cell.font = fontbold
    pl_sheet[f'B{start_row}'] = 'Federal Program Revenue'
    pl_sheet[f'D{start_row}'] = totals['value']
    pl_sheet[f'E{start_row}'].value = f'=SUM(E{fpr_row_start}:E{fpr_row_end})' 
    pl_sheet[f'G{start_row}'] = totals['total_real9']
    pl_sheet[f'H{start_row}'] = totals['total_real10']
    pl_sheet[f'I{start_row}'] = totals['total_real11']
    pl_sheet[f'J{start_row}'] = totals['total_real12']
    pl_sheet[f'K{start_row}'] = totals['total_real1']
    pl_sheet[f'L{start_row}'] = totals['total_real2']
    pl_sheet[f'M{start_row}'] = totals['total_real3']
    pl_sheet[f'N{start_row}'] = totals['total_real4']
    pl_sheet[f'O{start_row}'] = totals['total_real5']
    pl_sheet[f'P{start_row}'] = totals['total_real6']
    pl_sheet[f'Q{start_row}'] = totals['total_real7']
    pl_sheet[f'R{start_row}'] = totals['total_real8']
    pl_sheet[f'T{start_row}'].value = f'=SUM(T{fpr_row_start}:T{fpr_row_end})'  
    pl_sheet[f'U{start_row}'].value = f'=+T{start_row}-E{start_row}'
    pl_sheet[f'V{start_row}'].value = f'=+T{start_row}/D{start_row}'   
    start_row += 1

    total_revenue_row = start_row
    for col in range(2, 19):  
        cell = pl_sheet.cell(row=start_row, column=col)
        cell.font = fontbold
    pl_sheet[f'B{start_row}'] = 'Total Revenue'
    pl_sheet[f'D{start_row}'].value = f'=SUM(D{spr_end},D{fpr_end},D{lr_end})'
    pl_sheet[f'E{start_row}'].value = f'=SUM(E{spr_end},E{fpr_end},E{lr_end})'
    pl_sheet[f'G{start_row}'].value = f'=SUM(G{spr_end},G{fpr_end},G{lr_end})'
    pl_sheet[f'H{start_row}'].value = f'=SUM(H{spr_end},H{fpr_end},H{lr_end})'
    pl_sheet[f'I{start_row}'].value = f'=SUM(I{spr_end},I{fpr_end},I{lr_end})'
    pl_sheet[f'J{start_row}'].value = f'=SUM(J{spr_end},J{fpr_end},J{lr_end})'
    pl_sheet[f'K{start_row}'].value = f'=SUM(K{spr_end},K{fpr_end},K{lr_end})'
    pl_sheet[f'L{start_row}'].value = f'=SUM(L{spr_end},L{fpr_end},L{lr_end})'
    pl_sheet[f'M{start_row}'].value = f'=SUM(M{spr_end},M{fpr_end},M{lr_end})'
    pl_sheet[f'N{start_row}'].value = f'=SUM(N{spr_end},N{fpr_end},N{lr_end})'
    pl_sheet[f'O{start_row}'].value = f'=SUM(O{spr_end},O{fpr_end},O{lr_end})'
    pl_sheet[f'P{start_row}'].value = f'=SUM(P{spr_end},P{fpr_end},P{lr_end})'
    pl_sheet[f'Q{start_row}'].value = f'=SUM(Q{spr_end},Q{fpr_end},Q{lr_end})'
    pl_sheet[f'R{start_row}'].value = f'=SUM(R{spr_end},R{fpr_end},R{lr_end})'
    pl_sheet[f'T{start_row}'].value = f'=SUM(T{spr_end},T{fpr_end},T{lr_end})'   
    pl_sheet[f'U{start_row}'].value = f'=+T{start_row}-E{start_row}'
    pl_sheet[f'V{start_row}'].value = f'=+T{start_row}/D{start_row}'     

    start_row += 4   
    first_total_start = start_row
    for row_data in data2: #1st TOTAL
        pl_sheet[f'B{start_row}'] = f'{row_data["func_func"]} - {row_data["desc"]}'
        pl_sheet[f'D{start_row}'] = row_data['budget']
        pl_sheet[f'E{start_row}'] = row_data['budget'] * .922222
        pl_sheet[f'G{start_row}'] = row_data['total_func9']
        pl_sheet[f'H{start_row}'] = row_data['total_func10']
        pl_sheet[f'I{start_row}'] = row_data['total_func11']
        pl_sheet[f'J{start_row}'] = row_data['total_func12']
        pl_sheet[f'K{start_row}'] = row_data['total_func1']
        pl_sheet[f'L{start_row}'] = row_data['total_func2']
        pl_sheet[f'M{start_row}'] = row_data['total_func3']
        pl_sheet[f'N{start_row}'] = row_data['total_func4']
        pl_sheet[f'O{start_row}'] = row_data['total_func5']
        pl_sheet[f'P{start_row}'] = row_data['total_func6']
        pl_sheet[f'Q{start_row}'] = row_data['total_func7']
        pl_sheet[f'R{start_row}'] = row_data['total_func8']
        pl_sheet[f'T{start_row}'].value = f'=SUBTOTAL(109,G{start_row}:R{start_row})'
        pl_sheet[f'U{start_row}'].value = f'=E{start_row}-T{start_row}' 
        pl_sheet[f'v{start_row}'].value = f'=IFERROR(T{start_row}/D{start_row},"    ")'
        first_total_end = start_row
        start_row += 1

 
    first_total_row = start_row
    for col in range(2, 19):  
        cell = pl_sheet.cell(row=start_row, column=col)
        cell.font = fontbold
    pl_sheet[f'B{start_row}'] = 'Total'
    pl_sheet[f'D{start_row}'].value = f'=SUM(D{first_total_start}:D{first_total_end})'
    pl_sheet[f'E{start_row}'].value = f'=SUM(E{first_total_start}:E{first_total_end})'
    pl_sheet[f'G{start_row}'].value = f'=SUM(G{first_total_start}:G{first_total_end})'
    pl_sheet[f'H{start_row}'].value = f'=SUM(H{first_total_start}:H{first_total_end})'
    pl_sheet[f'I{start_row}'].value = f'=SUM(I{first_total_start}:I{first_total_end})'
    pl_sheet[f'J{start_row}'].value = f'=SUM(J{first_total_start}:J{first_total_end})'
    pl_sheet[f'K{start_row}'].value = f'=SUM(K{first_total_start}:K{first_total_end})'
    pl_sheet[f'L{start_row}'].value = f'=SUM(L{first_total_start}:L{first_total_end})'
    pl_sheet[f'M{start_row}'].value = f'=SUM(M{first_total_start}:M{first_total_end})'
    pl_sheet[f'N{start_row}'].value = f'=SUM(N{first_total_start}:N{first_total_end})'
    pl_sheet[f'O{start_row}'].value = f'=SUM(O{first_total_start}:O{first_total_end})'
    pl_sheet[f'P{start_row}'].value = f'=SUM(P{first_total_start}:P{first_total_end})'
    pl_sheet[f'Q{start_row}'].value = f'=SUM(Q{first_total_start}:Q{first_total_end})'
    pl_sheet[f'R{start_row}'].value = f'=SUM(R{first_total_start}:R{first_total_end})'
    pl_sheet[f'T{start_row}'].value = f'=SUM(T{first_total_start}:T{first_total_end})'  
    pl_sheet[f'U{start_row}'].value = f'=SUM(U{first_total_start}:U{first_total_end})'
    pl_sheet[f'V{start_row}'].value = f'=+T{start_row}/D{start_row}' 
    pl_sheet[f'v{start_row}'].value = f'=IFERROR(+T{start_row}/E{start_row},"    ")'
    
    start_row += 2 #surplus (deficits) before depreciation
    surplus_row = start_row
    pl_sheet[f'B{start_row}'] = 'Surplus (Deficits) before Depreciation'
    pl_sheet[f'D{start_row}'].value = f'=(D{total_revenue_row}-D{first_total_row})'
    pl_sheet[f'E{start_row}'].value = f'=(E{total_revenue_row}-E{first_total_row})'
    pl_sheet[f'G{start_row}'].value = f'=(G{total_revenue_row}-G{first_total_row})'
    pl_sheet[f'H{start_row}'].value = f'=(H{total_revenue_row}-H{first_total_row})'
    pl_sheet[f'I{start_row}'].value = f'=(I{total_revenue_row}-I{first_total_row})'
    pl_sheet[f'J{start_row}'].value = f'=(J{total_revenue_row}-J{first_total_row})'
    pl_sheet[f'K{start_row}'].value = f'=(K{total_revenue_row}-K{first_total_row})'
    pl_sheet[f'L{start_row}'].value = f'=(L{total_revenue_row}-L{first_total_row})'
    pl_sheet[f'M{start_row}'].value = f'=(M{total_revenue_row}-M{first_total_row})'
    pl_sheet[f'N{start_row}'].value = f'=(N{total_revenue_row}-N{first_total_row})'
    pl_sheet[f'O{start_row}'].value = f'=(O{total_revenue_row}-O{first_total_row})'
    pl_sheet[f'P{start_row}'].value = f'=(P{total_revenue_row}-P{first_total_row})'
    pl_sheet[f'Q{start_row}'].value = f'=(Q{total_revenue_row}-Q{first_total_row})'
    pl_sheet[f'R{start_row}'].value = f'=(R{total_revenue_row}-R{first_total_row})'
    pl_sheet[f'T{start_row}'].value = f'=(T{total_revenue_row}-T{first_total_row})'  
    pl_sheet[f'U{start_row}'].value = f'=+T{start_row}-E{start_row}'
    pl_sheet[f'V{start_row}'].value = f'=IFERROR(T{start_row}/D{start_row},"    ")' 

    start_row += 2  

    dna_row_start = start_row
    for row_data in data2: #Depreciation and amortization
        if row_data["func_func"] == '51':
            pl_sheet[f'B{start_row}'] = f'{row_data["func_func"]} - {row_data["desc"]}'
            pl_sheet[f'C{start_row}'] = '6449'
            
            pl_sheet[f'G{start_row}'] = row_data['total_func2_9']
            pl_sheet[f'H{start_row}'] = row_data['total_func2_10']
            pl_sheet[f'I{start_row}'] = row_data['total_func2_11']
            pl_sheet[f'J{start_row}'] = row_data['total_func2_12']
            pl_sheet[f'K{start_row}'] = row_data['total_func2_1']
            pl_sheet[f'L{start_row}'] = row_data['total_func2_2']
            pl_sheet[f'M{start_row}'] = row_data['total_func2_3']
            pl_sheet[f'N{start_row}'] = row_data['total_func2_4']
            pl_sheet[f'O{start_row}'] = row_data['total_func2_5']
            pl_sheet[f'P{start_row}'] = row_data['total_func2_6']
            pl_sheet[f'Q{start_row}'] = row_data['total_func2_7']
            pl_sheet[f'R{start_row}'] = row_data['total_func2_8']
            pl_sheet[f'T{start_row}'].value = f'=SUBTOTAL(109,G{start_row}:R{start_row})'
            pl_sheet[f'U{start_row}'].value = f'=E{start_row}-T{start_row}'
            pl_sheet[f'V{start_row}'].value = f'=IFERROR(T{start_row}/D{start_row},"    ")'  
            dna_row_end = start_row
            start_row += 1

    start_row += 1
    dna_row = start_row
    pl_sheet[f'B{start_row}'] = 'Depreciation and Amortization'
    pl_sheet[f'D{start_row}'].value = f'=SUM(D{dna_row_start}:D{dna_row_end})'
    pl_sheet[f'E{start_row}'].value = f'=SUM(E{dna_row_start}:E{dna_row_end})'
    pl_sheet[f'G{start_row}'].value = f'=SUM(G{dna_row_start}:G{dna_row_end})'
    pl_sheet[f'H{start_row}'].value = f'=SUM(H{dna_row_start}:H{dna_row_end})'
    pl_sheet[f'I{start_row}'].value = f'=SUM(I{dna_row_start}:I{dna_row_end})'
    pl_sheet[f'J{start_row}'].value = f'=SUM(J{dna_row_start}:J{dna_row_end})'
    pl_sheet[f'K{start_row}'].value = f'=SUM(K{dna_row_start}:K{dna_row_end})'
    pl_sheet[f'L{start_row}'].value = f'=SUM(L{dna_row_start}:L{dna_row_end})'
    pl_sheet[f'M{start_row}'].value = f'=SUM(M{dna_row_start}:M{dna_row_end})'
    pl_sheet[f'N{start_row}'].value = f'=SUM(N{dna_row_start}:N{dna_row_end})'
    pl_sheet[f'O{start_row}'].value = f'=SUM(O{dna_row_start}:O{dna_row_end})'
    pl_sheet[f'P{start_row}'].value = f'=SUM(P{dna_row_start}:P{dna_row_end})'
    pl_sheet[f'Q{start_row}'].value = f'=SUM(Q{dna_row_start}:Q{dna_row_end})'
    pl_sheet[f'R{start_row}'].value = f'=SUM(R{dna_row_start}:R{dna_row_end})'
    pl_sheet[f'T{start_row}'].value = f'=SUBTOTAL(109,G{start_row}:R{start_row})'  
    pl_sheet[f'U{start_row}'].value = f'=E{start_row}-T{start_row}' 
    pl_sheet[f'V{start_row}'].value = f'=IFERROR(T{start_row}/E{start_row},"    ")' 

    start_row += 2
    netsurplus_row = start_row
    for col in range(2, 19):  
        cell = pl_sheet.cell(row=start_row, column=col)
        cell.font = fontbold
    pl_sheet[f'B{start_row}'] = 'Net Surplus(Deficit)'
    pl_sheet[f'D{start_row}'].value = f'=(D{surplus_row}-D{dna_row})'
    pl_sheet[f'E{start_row}'].value = f'=(E{surplus_row}-E{dna_row})'
    pl_sheet[f'G{start_row}'].value = f'=(G{surplus_row}-G{dna_row})'
    pl_sheet[f'H{start_row}'].value = f'=(H{surplus_row}-H{dna_row})'
    pl_sheet[f'I{start_row}'].value = f'=(I{surplus_row}-I{dna_row})'
    pl_sheet[f'J{start_row}'].value = f'=(J{surplus_row}-J{dna_row})'
    pl_sheet[f'K{start_row}'].value = f'=(K{surplus_row}-K{dna_row})'
    pl_sheet[f'L{start_row}'].value = f'=(L{surplus_row}-L{dna_row})'
    pl_sheet[f'M{start_row}'].value = f'=(M{surplus_row}-M{dna_row})'
    pl_sheet[f'N{start_row}'].value = f'=(N{surplus_row}-N{dna_row})'
    pl_sheet[f'O{start_row}'].value = f'=(O{surplus_row}-O{dna_row})'
    pl_sheet[f'P{start_row}'].value = f'=(P{surplus_row}-P{dna_row})'
    pl_sheet[f'Q{start_row}'].value = f'=(Q{surplus_row}-Q{dna_row})'
    pl_sheet[f'R{start_row}'].value = f'=(R{surplus_row}-R{dna_row})'
    pl_sheet[f'T{start_row}'].value = f'=(T{surplus_row}-T{dna_row})' 
    pl_sheet[f'U{start_row}'].value = f'=+T{start_row}-E{start_row}'
    pl_sheet[f'V{start_row}'].value = f'=IFERROR(T{start_row}/D{start_row},"    ")'   
    



    start_row += 2
    pl_sheet[f'B{start_row}'] = 'Expense By Object Codes'
    pl_sheet[f'B{start_row}'].font = fontbold

    start_row += 1
    payroll_row_start = start_row
    for row_data in data_activities: 
        if row_data['Category'] == 'Payroll Costs':
            pl_sheet[f'B{start_row}'] = f'{row_data["obj"]} - {row_data["Description"]}'
            
            pl_sheet[f'G{start_row}'] = row_data['total_activities9']
            pl_sheet[f'H{start_row}'] = row_data['total_activities10']
            pl_sheet[f'I{start_row}'] = row_data['total_activities11']
            pl_sheet[f'J{start_row}'] = row_data['total_activities12']
            pl_sheet[f'K{start_row}'] = row_data['total_activities1']
            pl_sheet[f'L{start_row}'] = row_data['total_activities2']
            pl_sheet[f'M{start_row}'] = row_data['total_activities3']
            pl_sheet[f'N{start_row}'] = row_data['total_activities4']
            pl_sheet[f'O{start_row}'] = row_data['total_activities5']
            pl_sheet[f'P{start_row}'] = row_data['total_activities6']
            pl_sheet[f'Q{start_row}'] = row_data['total_activities7']
            pl_sheet[f'R{start_row}'] = row_data['total_activities8']
            pl_sheet[f'T{start_row}'].value = f'=SUBTOTAL(109,G{start_row}:R{start_row})'
            payroll_row_end = start_row
            start_row += 1
            
    payroll_row = start_row
    for row_data in data_expensebyobject: 
        if row_data['obj'] == '6100':
            pl_sheet[f'B{start_row}'] = f'{row_data["obj"]} - {row_data["Description"]}'
            pl_sheet[f'D{start_row}'] = row_data['budget']
            pl_sheet[f'E{start_row}'] = row_data['budget'] * .922222
            pl_sheet[f'G{start_row}'].value = f'=SUM(G{payroll_row_start}:G{payroll_row_end})'
            pl_sheet[f'H{start_row}'].value = f'=SUM(H{payroll_row_start}:H{payroll_row_end})'
            pl_sheet[f'I{start_row}'].value = f'=SUM(I{payroll_row_start}:I{payroll_row_end})'
            pl_sheet[f'J{start_row}'].value = f'=SUM(J{payroll_row_start}:J{payroll_row_end})'
            pl_sheet[f'K{start_row}'].value = f'=SUM(K{payroll_row_start}:K{payroll_row_end})'
            pl_sheet[f'L{start_row}'].value = f'=SUM(L{payroll_row_start}:L{payroll_row_end})'
            pl_sheet[f'M{start_row}'].value = f'=SUM(M{payroll_row_start}:M{payroll_row_end})'
            pl_sheet[f'N{start_row}'].value = f'=SUM(N{payroll_row_start}:N{payroll_row_end})'
            pl_sheet[f'O{start_row}'].value = f'=SUM(O{payroll_row_start}:O{payroll_row_end})'
            pl_sheet[f'P{start_row}'].value = f'=SUM(P{payroll_row_start}:P{payroll_row_end})'
            pl_sheet[f'Q{start_row}'].value = f'=SUM(Q{payroll_row_start}:Q{payroll_row_end})'
            pl_sheet[f'R{start_row}'].value = f'=SUM(R{payroll_row_start}:R{payroll_row_end})'
            pl_sheet[f'T{start_row}'].value = f'=SUBTOTAL(109,G{start_row}:R{start_row})'  
            pl_sheet[f'U{start_row}'].value = f'=E{start_row}-T{start_row}' 
            pl_sheet[f'V{start_row}'].value = f'=IFERROR(T{start_row}/D{start_row},"    ")' 
            start_row += 1

    pcs_row_start = start_row
    for row_data in data_activities: 
        if row_data['Category'] == 'Professional and Cont Svcs':
            pl_sheet[f'B{start_row}'] = f'{row_data["obj"]} - {row_data["Description"]}'
            pl_sheet[f'G{start_row}'] = row_data['total_activities9']
            pl_sheet[f'H{start_row}'] = row_data['total_activities10']
            pl_sheet[f'I{start_row}'] = row_data['total_activities11']
            pl_sheet[f'J{start_row}'] = row_data['total_activities12']
            pl_sheet[f'K{start_row}'] = row_data['total_activities1']
            pl_sheet[f'L{start_row}'] = row_data['total_activities2']
            pl_sheet[f'M{start_row}'] = row_data['total_activities3']
            pl_sheet[f'N{start_row}'] = row_data['total_activities4']
            pl_sheet[f'O{start_row}'] = row_data['total_activities5']
            pl_sheet[f'P{start_row}'] = row_data['total_activities6']
            pl_sheet[f'Q{start_row}'] = row_data['total_activities7']
            pl_sheet[f'R{start_row}'] = row_data['total_activities8']
            pl_sheet[f'T{start_row}'].value = f'=SUBTOTAL(109,G{start_row}:R{start_row})'
            pcs_row_end = start_row
            start_row += 1

    pcs_row = start_row
    for row_data in data_expensebyobject: 
        if row_data['obj'] == '6200':
            pl_sheet[f'B{start_row}'] = f'{row_data["obj"]} - {row_data["Description"]}'
            pl_sheet[f'D{start_row}'] = row_data['budget']
            pl_sheet[f'E{start_row}'] = row_data['budget'] * .922222
            pl_sheet[f'G{start_row}'].value = f'=SUM(G{pcs_row_start}:G{pcs_row_end})'
            pl_sheet[f'H{start_row}'].value = f'=SUM(H{pcs_row_start}:H{pcs_row_end})'
            pl_sheet[f'I{start_row}'].value = f'=SUM(I{pcs_row_start}:I{pcs_row_end})'
            pl_sheet[f'J{start_row}'].value = f'=SUM(J{pcs_row_start}:J{pcs_row_end})'
            pl_sheet[f'K{start_row}'].value = f'=SUM(K{pcs_row_start}:K{pcs_row_end})'
            pl_sheet[f'L{start_row}'].value = f'=SUM(L{pcs_row_start}:L{pcs_row_end})'
            pl_sheet[f'M{start_row}'].value = f'=SUM(M{pcs_row_start}:M{pcs_row_end})'
            pl_sheet[f'N{start_row}'].value = f'=SUM(N{pcs_row_start}:N{pcs_row_end})'
            pl_sheet[f'O{start_row}'].value = f'=SUM(O{pcs_row_start}:O{pcs_row_end})'
            pl_sheet[f'P{start_row}'].value = f'=SUM(P{pcs_row_start}:P{pcs_row_end})'
            pl_sheet[f'Q{start_row}'].value = f'=SUM(Q{pcs_row_start}:Q{pcs_row_end})'
            pl_sheet[f'R{start_row}'].value = f'=SUM(R{pcs_row_start}:R{pcs_row_end})'
            pl_sheet[f'T{start_row}'].value = f'=SUBTOTAL(109,G{start_row}:R{start_row})' 
           
            pl_sheet[f'U{start_row}'].value = f'=E{start_row}-T{start_row}' 
            pl_sheet[f'V{start_row}'].value = f'=IFERROR(T{start_row}/D{start_row},"    ")' 
            start_row += 1

    sm_row_start = start_row
    for row_data in data_activities: 
        if row_data['Category'] == 'Supplies and Materials':
            pl_sheet[f'B{start_row}'] = f'{row_data["obj"]} - {row_data["Description"]}'
            pl_sheet[f'G{start_row}'] = row_data['total_activities9']
            pl_sheet[f'H{start_row}'] = row_data['total_activities10']
            pl_sheet[f'I{start_row}'] = row_data['total_activities11']
            pl_sheet[f'J{start_row}'] = row_data['total_activities12']
            pl_sheet[f'K{start_row}'] = row_data['total_activities1']
            pl_sheet[f'L{start_row}'] = row_data['total_activities2']
            pl_sheet[f'M{start_row}'] = row_data['total_activities3']
            pl_sheet[f'N{start_row}'] = row_data['total_activities4']
            pl_sheet[f'O{start_row}'] = row_data['total_activities5']
            pl_sheet[f'P{start_row}'] = row_data['total_activities6']
            pl_sheet[f'Q{start_row}'] = row_data['total_activities7']
            pl_sheet[f'R{start_row}'] = row_data['total_activities8']
            pl_sheet[f'T{start_row}'].value = f'=SUBTOTAL(109,G{start_row}:R{start_row})'
            sm_row_end = start_row
            start_row += 1

    sm_row = start_row
    for row_data in data_expensebyobject: 
        if row_data['obj'] == '6300':
            pl_sheet[f'B{start_row}'] = f'{row_data["obj"]} - {row_data["Description"]}'
            pl_sheet[f'D{start_row}'] = row_data['budget']
            pl_sheet[f'E{start_row}'] = row_data['budget'] * .922222
            pl_sheet[f'G{start_row}'].value = f'=SUM(G{sm_row_start}:G{sm_row_end})'
            pl_sheet[f'H{start_row}'].value = f'=SUM(H{sm_row_start}:H{sm_row_end})'
            pl_sheet[f'I{start_row}'].value = f'=SUM(I{sm_row_start}:I{sm_row_end})'
            pl_sheet[f'J{start_row}'].value = f'=SUM(J{sm_row_start}:J{sm_row_end})'
            pl_sheet[f'K{start_row}'].value = f'=SUM(K{sm_row_start}:K{sm_row_end})'
            pl_sheet[f'L{start_row}'].value = f'=SUM(L{sm_row_start}:L{sm_row_end})'
            pl_sheet[f'M{start_row}'].value = f'=SUM(M{sm_row_start}:M{sm_row_end})'
            pl_sheet[f'N{start_row}'].value = f'=SUM(N{sm_row_start}:N{sm_row_end})'
            pl_sheet[f'O{start_row}'].value = f'=SUM(O{sm_row_start}:O{sm_row_end})'
            pl_sheet[f'P{start_row}'].value = f'=SUM(P{sm_row_start}:P{sm_row_end})'
            pl_sheet[f'Q{start_row}'].value = f'=SUM(Q{sm_row_start}:Q{sm_row_end})'
            pl_sheet[f'R{start_row}'].value = f'=SUM(R{sm_row_start}:R{sm_row_end})'
            pl_sheet[f'T{start_row}'].value = f'=SUBTOTAL(109,G{start_row}:R{start_row})'  
            
            pl_sheet[f'U{start_row}'].value = f'=E{start_row}-T{start_row}' 
            pl_sheet[f'V{start_row}'].value = f'=IFERROR(T{start_row}/D{start_row},"    ")' 
            start_row += 1
        
    ooe_row_start = start_row
    for row_data in data_activities: 
        if row_data['Category'] == 'Other Operating Expenses':
            pl_sheet[f'B{start_row}'] = f'{row_data["obj"]} - {row_data["Description"]}'
            pl_sheet[f'G{start_row}'] = row_data['total_activities9']
            pl_sheet[f'H{start_row}'] = row_data['total_activities10']
            pl_sheet[f'I{start_row}'] = row_data['total_activities11']
            pl_sheet[f'J{start_row}'] = row_data['total_activities12']
            pl_sheet[f'K{start_row}'] = row_data['total_activities1']
            pl_sheet[f'L{start_row}'] = row_data['total_activities2']
            pl_sheet[f'M{start_row}'] = row_data['total_activities3']
            pl_sheet[f'N{start_row}'] = row_data['total_activities4']
            pl_sheet[f'O{start_row}'] = row_data['total_activities5']
            pl_sheet[f'P{start_row}'] = row_data['total_activities6']
            pl_sheet[f'Q{start_row}'] = row_data['total_activities7']
            pl_sheet[f'R{start_row}'] = row_data['total_activities8']
            ooe_row_end = start_row
            start_row += 1

    ooe_row = start_row
    for row_data in data_expensebyobject: 
        if row_data['obj'] == '6400':
            pl_sheet[f'B{start_row}'] = f'{row_data["obj"]} - {row_data["Description"]}'
            pl_sheet[f'D{start_row}'] = row_data['budget']
            pl_sheet[f'E{start_row}'] = row_data['budget'] * .922222
            pl_sheet[f'G{start_row}'].value = f'=SUM(G{ooe_row_start}:G{ooe_row_end})'
            pl_sheet[f'H{start_row}'].value = f'=SUM(H{ooe_row_start}:H{ooe_row_end})'
            pl_sheet[f'I{start_row}'].value = f'=SUM(I{ooe_row_start}:I{ooe_row_end})'
            pl_sheet[f'J{start_row}'].value = f'=SUM(J{ooe_row_start}:J{ooe_row_end})'
            pl_sheet[f'K{start_row}'].value = f'=SUM(K{ooe_row_start}:K{ooe_row_end})'
            pl_sheet[f'L{start_row}'].value = f'=SUM(L{ooe_row_start}:L{ooe_row_end})'
            pl_sheet[f'M{start_row}'].value = f'=SUM(M{ooe_row_start}:M{ooe_row_end})'
            pl_sheet[f'N{start_row}'].value = f'=SUM(N{ooe_row_start}:N{ooe_row_end})'
            pl_sheet[f'O{start_row}'].value = f'=SUM(O{ooe_row_start}:O{ooe_row_end})'
            pl_sheet[f'P{start_row}'].value = f'=SUM(P{ooe_row_start}:P{ooe_row_end})'
            pl_sheet[f'Q{start_row}'].value = f'=SUM(Q{ooe_row_start}:Q{ooe_row_end})'
            pl_sheet[f'R{start_row}'].value = f'=SUM(R{ooe_row_start}:R{ooe_row_end})'
            pl_sheet[f'T{start_row}'].value = f'=SUBTOTAL(109,G{start_row}:R{start_row})'  
           
            pl_sheet[f'U{start_row}'].value = f'=E{start_row}-T{start_row}' 
            pl_sheet[f'V{start_row}'].value = f'=IFERROR(T{start_row}/D{start_row},"    ")' 
            
            start_row += 1


    total_expense_row_start = start_row
    for row_data in data_activities: 
        if row_data['Category'] == 'Total Expense':
            pl_sheet[f'B{start_row}'] = f'{row_data["obj"]} - {row_data["Description"]}'
            pl_sheet[f'G{start_row}'] = row_data['total_activities9']
            pl_sheet[f'H{start_row}'] = row_data['total_activities10']
            pl_sheet[f'I{start_row}'] = row_data['total_activities11']
            pl_sheet[f'J{start_row}'] = row_data['total_activities12']
            pl_sheet[f'K{start_row}'] = row_data['total_activities1']
            pl_sheet[f'L{start_row}'] = row_data['total_activities2']
            pl_sheet[f'M{start_row}'] = row_data['total_activities3']
            pl_sheet[f'N{start_row}'] = row_data['total_activities4']
            pl_sheet[f'O{start_row}'] = row_data['total_activities5']
            pl_sheet[f'P{start_row}'] = row_data['total_activities6']
            pl_sheet[f'Q{start_row}'] = row_data['total_activities7']
            pl_sheet[f'R{start_row}'] = row_data['total_activities8']
            pl_sheet[f'T{start_row}'].value = f'=SUBTOTAL(109,G{start_row}:R{start_row})'
            
            
            total_expense_row_end = start_row
            start_row += 1

    total_expense_row = start_row
    for row_data in data_expensebyobject: 
        if row_data['obj'] == '6500':
            pl_sheet[f'B{start_row}'] = f'{row_data["obj"]} - {row_data["Description"]}'
            pl_sheet[f'D{start_row}'] = row_data['budget']
            pl_sheet[f'E{start_row}'] = row_data['budget'] * .922222
            pl_sheet[f'G{start_row}'].value = f'=SUM(G{total_expense_row_start}:G{total_expense_row_end})'
            pl_sheet[f'H{start_row}'].value = f'=SUM(H{total_expense_row_start}:H{total_expense_row_end})'
            pl_sheet[f'I{start_row}'].value = f'=SUM(I{total_expense_row_start}:I{total_expense_row_end})'
            pl_sheet[f'J{start_row}'].value = f'=SUM(J{total_expense_row_start}:J{total_expense_row_end})'
            pl_sheet[f'K{start_row}'].value = f'=SUM(K{total_expense_row_start}:K{total_expense_row_end})'
            pl_sheet[f'L{start_row}'].value = f'=SUM(L{total_expense_row_start}:L{total_expense_row_end})'
            pl_sheet[f'M{start_row}'].value = f'=SUM(M{total_expense_row_start}:M{total_expense_row_end})'
            pl_sheet[f'N{start_row}'].value = f'=SUM(N{total_expense_row_start}:N{total_expense_row_end})'
            pl_sheet[f'O{start_row}'].value = f'=SUM(O{total_expense_row_start}:O{total_expense_row_end})'
            pl_sheet[f'P{start_row}'].value = f'=SUM(P{total_expense_row_start}:P{total_expense_row_end})'
            pl_sheet[f'Q{start_row}'].value = f'=SUM(Q{total_expense_row_start}:Q{total_expense_row_end})'
            pl_sheet[f'R{start_row}'].value = f'=SUM(R{total_expense_row_start}:R{total_expense_row_end})'
            pl_sheet[f'T{start_row}'].value = f'=SUBTOTAL(109,G{start_row}:R{start_row})'  
            
            pl_sheet[f'U{start_row}'].value = f'=E{start_row}-T{start_row}' 
            pl_sheet[f'V{start_row}'].value = f'=IFERROR(T{start_row}/D{start_row},"    ")' 
            start_row += 1

    start_row += 1

    total_expense_total = start_row
    pl_sheet[f'B{start_row}'] = 'Total Expense'
    pl_sheet[f'D{start_row}'].value = f'=SUM(D{payroll_row},D{pcs_row},D{sm_row},D{ooe_row},D{total_expense_row})'
    pl_sheet[f'E{start_row}'].value = f'=SUM(E{payroll_row},E{pcs_row},E{sm_row},E{ooe_row},E{total_expense_row})'
    pl_sheet[f'G{start_row}'].value = f'=SUM(G{payroll_row},G{pcs_row},G{sm_row},G{ooe_row},G{total_expense_row})'
    pl_sheet[f'H{start_row}'].value = f'=SUM(H{payroll_row},H{pcs_row},H{sm_row},H{ooe_row},H{total_expense_row})'
    pl_sheet[f'I{start_row}'].value = f'=SUM(I{payroll_row},I{pcs_row},I{sm_row},I{ooe_row},I{total_expense_row})'
    pl_sheet[f'J{start_row}'].value = f'=SUM(J{payroll_row},J{pcs_row},J{sm_row},J{ooe_row},J{total_expense_row})'
    pl_sheet[f'K{start_row}'].value = f'=SUM(K{payroll_row},K{pcs_row},K{sm_row},K{ooe_row},K{total_expense_row})'
    pl_sheet[f'L{start_row}'].value = f'=SUM(L{payroll_row},L{pcs_row},L{sm_row},L{ooe_row},L{total_expense_row})'
    pl_sheet[f'M{start_row}'].value = f'=SUM(M{payroll_row},M{pcs_row},M{sm_row},M{ooe_row},M{total_expense_row})'
    pl_sheet[f'N{start_row}'].value = f'=SUM(N{payroll_row},N{pcs_row},N{sm_row},N{ooe_row},N{total_expense_row})'
    pl_sheet[f'O{start_row}'].value = f'=SUM(O{payroll_row},O{pcs_row},O{sm_row},O{ooe_row},O{total_expense_row})'
    pl_sheet[f'P{start_row}'].value = f'=SUM(P{payroll_row},P{pcs_row},P{sm_row},P{ooe_row},P{total_expense_row})'
    pl_sheet[f'Q{start_row}'].value = f'=SUM(Q{payroll_row},Q{pcs_row},Q{sm_row},Q{ooe_row},Q{total_expense_row})'
    pl_sheet[f'R{start_row}'].value = f'=SUM(R{payroll_row},R{pcs_row},R{sm_row},R{ooe_row},R{total_expense_row})'
     
    pl_sheet[f'T{start_row}'].value = f'=SUM(T{payroll_row},T{pcs_row},T{sm_row},T{ooe_row},T{total_expense_row})'
    pl_sheet[f'U{start_row}'].value = f'=SUM(U{payroll_row},U{pcs_row},U{sm_row},U{ooe_row},U{total_expense_row})'
    pl_sheet[f'V{start_row}'].value = f'=IFERROR(T{start_row}/D{start_row},"    ")' 
    

    start_row += 1
    pl_sheet[f'B{start_row}'] = 'Net Income'
    pl_sheet[f'D{start_row}'].value = f'=(D{total_revenue_row}-D{total_expense_total})'
    pl_sheet[f'E{start_row}'].value = f'=(E{total_revenue_row}-E{total_expense_total})'
    pl_sheet[f'G{start_row}'].value = f'=(G{total_revenue_row}-G{total_expense_total})'
    pl_sheet[f'H{start_row}'].value = f'=(H{total_revenue_row}-H{total_expense_total})'
    pl_sheet[f'I{start_row}'].value = f'=(I{total_revenue_row}-I{total_expense_total})'
    pl_sheet[f'J{start_row}'].value = f'=(J{total_revenue_row}-J{total_expense_total})'
    pl_sheet[f'K{start_row}'].value = f'=(K{total_revenue_row}-K{total_expense_total})'
    pl_sheet[f'L{start_row}'].value = f'=(L{total_revenue_row}-L{total_expense_total})'
    pl_sheet[f'M{start_row}'].value = f'=(M{total_revenue_row}-M{total_expense_total})'
    pl_sheet[f'N{start_row}'].value = f'=(N{total_revenue_row}-N{total_expense_total})'
    pl_sheet[f'O{start_row}'].value = f'=(O{total_revenue_row}-O{total_expense_total})'
    pl_sheet[f'P{start_row}'].value = f'=(P{total_revenue_row}-P{total_expense_total})'
    pl_sheet[f'Q{start_row}'].value = f'=(Q{total_revenue_row}-Q{total_expense_total})'
    pl_sheet[f'R{start_row}'].value = f'=(R{total_revenue_row}-R{total_expense_total})'
    
    pl_sheet[f'T{start_row}'].value = f'=(T{total_revenue_row}-T{total_expense_total})'
    pl_sheet[f'U{start_row}'].value = f'=(U{total_revenue_row}-U{total_expense_total})'
    pl_sheet[f'V{start_row}'].value = f'=IFERROR(T{start_row}/D{start_row},"    ")' 
    


    start_row += 4 #Total expense and Net income
 
    # BS DESIGN
    for col in range(7, 19):
        col_letter = get_column_letter(col)
        bs_sheet.column_dimensions[col_letter].outline_level = 1
        bs_sheet.column_dimensions[col_letter].hidden = True
    for row in range(2,181):
        bs_sheet.row_dimensions[row].height = 19
    # bs_sheet.row_dimensions[17].height = 26 #local revenue
    # bs_sheet.row_dimensions[20].height = 26 #spr
    # bs_sheet.row_dimensions[33].height = 26 #fpr

    

    bs_sheet.column_dimensions['A'].width = 8
    bs_sheet.column_dimensions['B'].width = 32
    bs_sheet.column_dimensions['C'].hidden = True
    bs_sheet.column_dimensions['D'].width = 14
    bs_sheet.column_dimensions['E'].width = 28
    bs_sheet.column_dimensions['F'].width = 15
    bs_sheet.column_dimensions['G'].width = 15
    bs_sheet.column_dimensions['H'].width = 13
    bs_sheet.column_dimensions['I'].width = 13
    bs_sheet.column_dimensions['J'].width = 13
    bs_sheet.column_dimensions['K'].width = 13
    bs_sheet.column_dimensions['L'].width = 13
    bs_sheet.column_dimensions['M'].width = 13
    bs_sheet.column_dimensions['N'].width = 13
    bs_sheet.column_dimensions['O'].width = 13
    bs_sheet.column_dimensions['P'].width = 13
    bs_sheet.column_dimensions['Q'].width = 13
    bs_sheet.column_dimensions['R'].width = 13
    bs_sheet.column_dimensions['S'].width = 3
    bs_sheet.column_dimensions['T'].width = 17
    bs_sheet.column_dimensions['U'].width = 17
    bs_sheet.column_dimensions['V'].width = 12

    indent_style = NamedStyle(name="indent_style", alignment=Alignment(indent=2))
    indent_style2 = NamedStyle(name="indent_style2", alignment=Alignment(indent=4))


    #--- BS INSERT
    start_row_bs = 6
    bs_sheet[f'D{start_row_bs}'] = 'Current Assets'
    
    for row in data_activitybs:
        if row['Activity'] == 'Cash':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] =  row['Description2']
            bs_sheet[f'G{start_row_bs}'] =  row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] =  row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] =  row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] =  row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] =  row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] =  row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] =  row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] =  row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] =  row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] =  row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] =  row['total_bal7']
            #bs_sheet[f'R{start_row_bs}'] =  row['total_bal8']



    for row in data_balancesheet:
        if row['Activity'] == 'Cash':
            if row['Category'] == 'Assets':
                if row['Subcategory'] == 'Current Assets':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    cash_row_bs = start_row_bs

    

    for row in data_activitybs:
        if row['Activity'] == 'Restr':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] = row['Description2']
            bs_sheet[f'G{start_row_bs}'] = row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] = row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] = row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] = row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] = row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] = row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] = row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] = row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] = row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] = row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] = row['total_bal7']
            #bs_sheet[f'R{start_row_bs}'] = row['total_bal8']


    for row in data_balancesheet:
        if row['Activity'] == 'Restr':
            if row['Category'] == 'Assets':
                if row['Subcategory'] == 'Current Assets':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
           
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    restr_row_bs = start_row_bs

    for row in data_activitybs: 
        if row['Activity'] == 'DFS+F':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] = row['Description2']
            bs_sheet[f'G{start_row_bs}'] = row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] = row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] = row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] = row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] = row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] = row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] = row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] = row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] = row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] = row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] = row['total_bal7']
            #bs_sheet[f'R{start_row_bs}'] = row['total_bal8']


    for row in data_balancesheet:
        if row['Activity'] == 'DFS+F':
            if row['Category'] == 'Assets':
                if row['Subcategory'] == 'Current Assets':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    dfs_row_bs = start_row_bs
    
    for row in data_activitybs: 
        if row['Activity'] == 'OTHR':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] = row['Description2']
            bs_sheet[f'G{start_row_bs}'] = row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] = row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] = row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] = row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] = row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] = row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] = row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] = row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] = row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] = row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] = row['total_bal7']
            #bs_sheet[f'R{start_row_bs}'] = row['total_bal8']
            

    for row in data_balancesheet:
        if row['Activity'] == 'OTHR':
            if row['Category'] == 'Assets':
                if row['Subcategory'] == 'Current Assets':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    othr_row_bs = start_row_bs

    for row in data_activitybs: 
        if row['Activity'] == 'Inventory':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] = row['Description2']
            bs_sheet[f'G{start_row_bs}'] = row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] = row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] = row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] = row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] = row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] = row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] = row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] = row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] = row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] = row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] = row['total_bal7']
            #bs_sheet[f'R{start_row_bs}'] = row['total_bal8']
          

    for row in data_balancesheet:
        if row['Activity'] == 'Inventory':
            if row['Category'] == 'Assets':
                if row['Subcategory'] == 'Current Assets':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    inventory_row_bs = start_row_bs



    for row in data_activitybs: 
        if row['Activity'] == 'PPD':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] = row['Description2']
            bs_sheet[f'G{start_row_bs}'] = row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] = row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] = row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] = row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] = row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] = row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] = row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] = row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] = row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] = row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] = row['total_bal7']
            #bs_sheet[f'R{start_row_bs}'] = row['total_bal8']


    for row in data_balancesheet:
        if row['Activity'] == 'PPD':
            if row['Category'] == 'Assets':
                if row['Subcategory'] == 'Current Assets':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    ppd_row_bs = start_row_bs

    start_row_bs += 1
    total_current_assets_row_bs = start_row_bs
    bs_sheet[f'D{start_row_bs}'].style = indent_style2
    bs_sheet[f'D{start_row_bs}'].font = fontbold
    bs_sheet[f'D{start_row_bs}'] = 'Total Current Assets'
    bs_sheet[f'F{start_row_bs}'].value = f'=SUM(F{cash_row_bs},F{restr_row_bs},F{dfs_row_bs},F{othr_row_bs},F{inventory_row_bs},F{ppd_row_bs})'
    bs_sheet[f'G{start_row_bs}'].value = f'=SUM(G{cash_row_bs},G{restr_row_bs},G{dfs_row_bs},G{othr_row_bs},G{inventory_row_bs},G{ppd_row_bs})'
    bs_sheet[f'H{start_row_bs}'].value = f'=SUM(H{cash_row_bs},H{restr_row_bs},H{dfs_row_bs},H{othr_row_bs},H{inventory_row_bs},H{ppd_row_bs})'
    bs_sheet[f'I{start_row_bs}'].value = f'=SUM(I{cash_row_bs},I{restr_row_bs},I{dfs_row_bs},I{othr_row_bs},I{inventory_row_bs},I{ppd_row_bs})'
    bs_sheet[f'J{start_row_bs}'].value = f'=SUM(J{cash_row_bs},J{restr_row_bs},J{dfs_row_bs},J{othr_row_bs},J{inventory_row_bs},J{ppd_row_bs})'
    bs_sheet[f'K{start_row_bs}'].value = f'=SUM(K{cash_row_bs},K{restr_row_bs},K{dfs_row_bs},K{othr_row_bs},K{inventory_row_bs},K{ppd_row_bs})'
    bs_sheet[f'L{start_row_bs}'].value = f'=SUM(L{cash_row_bs},L{restr_row_bs},L{dfs_row_bs},L{othr_row_bs},L{inventory_row_bs},L{ppd_row_bs})'
    bs_sheet[f'M{start_row_bs}'].value = f'=SUM(M{cash_row_bs},M{restr_row_bs},M{dfs_row_bs},M{othr_row_bs},M{inventory_row_bs},M{ppd_row_bs})'
    bs_sheet[f'N{start_row_bs}'].value = f'=SUM(N{cash_row_bs},N{restr_row_bs},N{dfs_row_bs},N{othr_row_bs},N{inventory_row_bs},N{ppd_row_bs})'
    bs_sheet[f'O{start_row_bs}'].value = f'=SUM(O{cash_row_bs},O{restr_row_bs},O{dfs_row_bs},O{othr_row_bs},O{inventory_row_bs},O{ppd_row_bs})'
    bs_sheet[f'P{start_row_bs}'].value = f'=SUM(P{cash_row_bs},P{restr_row_bs},P{dfs_row_bs},P{othr_row_bs},P{inventory_row_bs},P{ppd_row_bs})'
    bs_sheet[f'Q{start_row_bs}'].value = f'=SUM(Q{cash_row_bs},Q{restr_row_bs},Q{dfs_row_bs},Q{othr_row_bs},Q{inventory_row_bs},Q{ppd_row_bs})'
    #bs_sheet[f'R{start_row_bs}'].value = f'=SUM(R{cash_row_bs},R{restr_row_bs},R{dfs_row_bs},R{othr_row_bs},R{inventory_row_bs},R{ppd_row_bs})'
    bs_sheet[f'T{start_row_bs}'].value = f'=SUM(T{cash_row_bs},T{restr_row_bs},T{dfs_row_bs},T{othr_row_bs},T{inventory_row_bs},T{ppd_row_bs})'  
    bs_sheet[f'U{start_row_bs}'].value = f'=SUM(U{cash_row_bs},U{restr_row_bs},U{dfs_row_bs},U{othr_row_bs},U{inventory_row_bs},U{ppd_row_bs})'
    
    start_row_bs += 1
    bs_sheet[f'D{start_row_bs}'] = 'Capital Assets , Net'
    bs_sheet.row_dimensions[start_row_bs].height = 37 

    for row in data_activitybs: 
        if row['Activity'] == 'FA-L':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] = row['Description2']
            bs_sheet[f'G{start_row_bs}'] = row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] = row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] = row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] = row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] = row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] = row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] = row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] = row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] = row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] = row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] = row['total_bal7']
            #bs_sheet[f'R{start_row_bs}'] = row['total_bal8']


    for row in data_balancesheet:
        if row['Activity'] == 'FA-L':
            if row['Category'] == 'Assets':
                if row['Subcategory'] == 'Capital Assets, Net':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    fal_row_bs = start_row_bs

    for row in data_activitybs: 
        if row['Activity'] == 'FA-BFE':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] = row['Description2']
            bs_sheet[f'G{start_row_bs}'] = row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] = row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] = row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] = row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] = row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] = row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] = row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] = row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] = row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] = row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] = row['total_bal7']
            #bs_sheet[f'R{start_row_bs}'] = row['total_bal8']


    for row in data_balancesheet:
        if row['Activity'] == 'FA-BFE':
            if row['Category'] == 'Assets':
                if row['Subcategory'] == 'Capital Assets, Net':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    fabfe_row_bs = start_row_bs

    for row in data_activitybs: 
        if row['Activity'] == 'FA-AD':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] = row['Description2']
            bs_sheet[f'G{start_row_bs}'] = row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] = row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] = row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] = row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] = row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] = row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] = row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] = row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] = row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] = row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] = row['total_bal7']
            #bs_sheet[f'R{start_row_bs}'] = row['total_bal8']
   

    for row in data_balancesheet:
        if row['Activity'] == 'FA-AD':
            if row['Category'] == 'Assets':
                if row['Subcategory'] == 'Capital Assets, Net':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    faad_row_bs = start_row_bs


    start_row_bs += 1
    total_capital_assets_row_bs = start_row_bs
    bs_sheet[f'D{start_row_bs}'].style = indent_style2
    bs_sheet[f'D{start_row_bs}'].font = fontbold
    bs_sheet[f'D{start_row_bs}'] = 'Total Capital Assets'
    
    bs_sheet[f'F{start_row_bs}'].value = f'=SUM(F{fal_row_bs},F{fabfe_row_bs},F{faad_row_bs})'
    bs_sheet[f'G{start_row_bs}'].value = f'=SUM(G{fal_row_bs},G{fabfe_row_bs},G{faad_row_bs})'
    bs_sheet[f'H{start_row_bs}'].value = f'=SUM(H{fal_row_bs},H{fabfe_row_bs},H{faad_row_bs})'
    bs_sheet[f'I{start_row_bs}'].value = f'=SUM(I{fal_row_bs},I{fabfe_row_bs},I{faad_row_bs})'
    bs_sheet[f'J{start_row_bs}'].value = f'=SUM(J{fal_row_bs},J{fabfe_row_bs},J{faad_row_bs})'
    bs_sheet[f'K{start_row_bs}'].value = f'=SUM(K{fal_row_bs},K{fabfe_row_bs},K{faad_row_bs})'
    bs_sheet[f'L{start_row_bs}'].value = f'=SUM(L{fal_row_bs},L{fabfe_row_bs},L{faad_row_bs})'
    bs_sheet[f'M{start_row_bs}'].value = f'=SUM(M{fal_row_bs},M{fabfe_row_bs},M{faad_row_bs})'
    bs_sheet[f'N{start_row_bs}'].value = f'=SUM(N{fal_row_bs},N{fabfe_row_bs},N{faad_row_bs})'
    bs_sheet[f'O{start_row_bs}'].value = f'=SUM(O{fal_row_bs},O{fabfe_row_bs},O{faad_row_bs})'
    bs_sheet[f'P{start_row_bs}'].value = f'=SUM(P{fal_row_bs},P{fabfe_row_bs},P{faad_row_bs})'
    bs_sheet[f'Q{start_row_bs}'].value = f'=SUM(Q{fal_row_bs},Q{fabfe_row_bs},Q{faad_row_bs})'
    #bs_sheet[f'R{start_row_bs}'].value = f'=SUM(R{fal_row_bs},R{fabfe_row_bs},R{faad_row_bs})'
    bs_sheet[f'T{start_row_bs}'].value = f'=SUM(T{fal_row_bs},T{fabfe_row_bs},T{faad_row_bs})'  
    bs_sheet[f'U{start_row_bs}'].value = f'=SUM(U{fal_row_bs},U{fabfe_row_bs},U{faad_row_bs})'
    
    start_row_bs += 1
    total_assets_row_bs = start_row_bs
    bs_sheet.row_dimensions[start_row_bs].height = 37         
    bs_sheet[f'D{start_row_bs}'].font = fontbold
    bs_sheet[f'D{start_row_bs}'] = 'Total  Assets'
    bs_sheet[f'F{start_row_bs}'].value = f'=SUM(F{total_capital_assets_row_bs},F{total_current_assets_row_bs})'
    bs_sheet[f'G{start_row_bs}'].value = f'=SUM(G{total_capital_assets_row_bs},G{total_current_assets_row_bs})'
    bs_sheet[f'H{start_row_bs}'].value = f'=SUM(H{total_capital_assets_row_bs},H{total_current_assets_row_bs})'
    bs_sheet[f'I{start_row_bs}'].value = f'=SUM(I{total_capital_assets_row_bs},I{total_current_assets_row_bs})'
    bs_sheet[f'J{start_row_bs}'].value = f'=SUM(J{total_capital_assets_row_bs},J{total_current_assets_row_bs})'
    bs_sheet[f'K{start_row_bs}'].value = f'=SUM(K{total_capital_assets_row_bs},K{total_current_assets_row_bs})'
    bs_sheet[f'L{start_row_bs}'].value = f'=SUM(L{total_capital_assets_row_bs},L{total_current_assets_row_bs})'
    bs_sheet[f'M{start_row_bs}'].value = f'=SUM(M{total_capital_assets_row_bs},M{total_current_assets_row_bs})'
    bs_sheet[f'N{start_row_bs}'].value = f'=SUM(N{total_capital_assets_row_bs},N{total_current_assets_row_bs})'
    bs_sheet[f'O{start_row_bs}'].value = f'=SUM(O{total_capital_assets_row_bs},O{total_current_assets_row_bs})'
    bs_sheet[f'P{start_row_bs}'].value = f'=SUM(P{total_capital_assets_row_bs},P{total_current_assets_row_bs})'
    bs_sheet[f'Q{start_row_bs}'].value = f'=SUM(Q{total_capital_assets_row_bs},Q{total_current_assets_row_bs})'
    #bs_sheet[f'R{start_row_bs}'].value = f'=SUM(R{total_capital_assets_row_bs},R{total_current_assets_row_bs})'
    bs_sheet[f'T{start_row_bs}'].value = f'=SUM(T{total_capital_assets_row_bs},T{total_current_assets_row_bs})'  
    bs_sheet[f'U{start_row_bs}'].value = f'=SUM(U{total_capital_assets_row_bs},U{total_current_assets_row_bs})'

    start_row_bs += 1
    bs_sheet.row_dimensions[start_row_bs].height = 37 
    
    bs_sheet[f'D{start_row_bs}'].font = fontbold
    bs_sheet[f'D{start_row_bs}'] = 'Liabilities and Net Assets'
    
    start_row_bs += 1
    bs_sheet.row_dimensions[start_row_bs].height = 37 
    bs_sheet[f'D{start_row_bs}'] = 'Current Liabilities'

    for row in data_activitybs: 
        if row['Activity'] == 'AP':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] = row['Description2']
            bs_sheet[f'G{start_row_bs}'] = row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] = row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] = row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] = row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] = row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] = row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] = row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] = row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] = row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] = row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] = row['total_bal7']
            #bs_sheet[f'R{start_row_bs}'] = row['total_bal8']
       

    for row in data_balancesheet:
        if row['Activity'] == 'AP':
            if row['Category'] == 'Liabilities and Net Assets':
                if row['Subcategory'] == 'Current Liabilities':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    ap_row_bs = start_row_bs

    for row in data_activitybs: 
        if row['Activity'] == 'Acc-Exp':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] = row['Description2']
            bs_sheet[f'G{start_row_bs}'] = row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] = row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] = row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] = row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] = row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] = row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] = row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] = row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] = row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] = row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] = row['total_bal7']
           #bs_sheet[f'R{start_row_bs}'] = row['total_bal8']
   

    for row in data_balancesheet:
        if row['Activity'] == 'Acc-Exp':
            if row['Category'] == 'Liabilities and Net Assets':
                if row['Subcategory'] == 'Current Liabilities':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    accexp_row_bs = start_row_bs

    for row in data_activitybs: 
        if row['Activity'] == 'OtherLab':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] = row['Description2']
            bs_sheet[f'G{start_row_bs}'] = row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] = row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] = row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] = row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] = row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] = row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] = row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] = row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] = row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] = row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] = row['total_bal7']
            #bs_sheet[f'R{start_row_bs}'] = row['total_bal8']
          

    for row in data_balancesheet:
        if row['Activity'] == 'OtherLab':
            if row['Category'] == 'Liabilities and Net Assets':
                if row['Subcategory'] == 'Current Liabilities':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    otherlab_row_bs = start_row_bs

    for row in data_activitybs: 
        if row['Activity'] == 'Debt-C':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] = row['Description2']
            bs_sheet[f'G{start_row_bs}'] = row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] = row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] = row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] = row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] = row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] = row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] = row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] = row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] = row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] = row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] = row['total_bal7']
            #bs_sheet[f'R{start_row_bs}'] = row['total_bal8']
            bs_sheet[f'T{start_row_bs}'] = 'ACTIVITYBS YTD'

    for row in data_balancesheet:
        if row['Activity'] == 'Debt-C':
            if row['Category'] == 'Liabilities and Net Assets':
                if row['Subcategory'] == 'Current Liabilities':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    debtc_row_bs = start_row_bs

    start_row_bs += 1
    total_current_liabilites_row_bs = start_row_bs
    
    bs_sheet[f'D{start_row_bs}'].style = indent_style2
    bs_sheet[f'D{start_row_bs}'].font = fontbold
  
    bs_sheet[f'D{start_row_bs}'] ='Total Current Liabilities'
    bs_sheet[f'F{start_row_bs}'].value = f'=SUM(F{ap_row_bs},F{accexp_row_bs},F{otherlab_row_bs},F{debtc_row_bs})'
    bs_sheet[f'G{start_row_bs}'].value = f'=SUM(G{ap_row_bs},G{accexp_row_bs},G{otherlab_row_bs},G{debtc_row_bs})'
    bs_sheet[f'H{start_row_bs}'].value = f'=SUM(H{ap_row_bs},H{accexp_row_bs},H{otherlab_row_bs},H{debtc_row_bs})'
    bs_sheet[f'I{start_row_bs}'].value = f'=SUM(I{ap_row_bs},I{accexp_row_bs},I{otherlab_row_bs},I{debtc_row_bs})'
    bs_sheet[f'J{start_row_bs}'].value = f'=SUM(J{ap_row_bs},J{accexp_row_bs},J{otherlab_row_bs},J{debtc_row_bs})'
    bs_sheet[f'K{start_row_bs}'].value = f'=SUM(K{ap_row_bs},K{accexp_row_bs},K{otherlab_row_bs},K{debtc_row_bs})'
    bs_sheet[f'L{start_row_bs}'].value = f'=SUM(L{ap_row_bs},L{accexp_row_bs},L{otherlab_row_bs},L{debtc_row_bs})'
    bs_sheet[f'M{start_row_bs}'].value = f'=SUM(M{ap_row_bs},M{accexp_row_bs},M{otherlab_row_bs},M{debtc_row_bs})'
    bs_sheet[f'N{start_row_bs}'].value = f'=SUM(N{ap_row_bs},N{accexp_row_bs},N{otherlab_row_bs},N{debtc_row_bs})'
    bs_sheet[f'O{start_row_bs}'].value = f'=SUM(O{ap_row_bs},O{accexp_row_bs},O{otherlab_row_bs},O{debtc_row_bs})'
    bs_sheet[f'P{start_row_bs}'].value = f'=SUM(P{ap_row_bs},P{accexp_row_bs},P{otherlab_row_bs},P{debtc_row_bs})'
    bs_sheet[f'Q{start_row_bs}'].value = f'=SUM(Q{ap_row_bs},Q{accexp_row_bs},Q{otherlab_row_bs},Q{debtc_row_bs})'
    #bs_sheet[f'R{start_row_bs}'].value = f'=SUM(R{ap_row_bs},R{accexp_row_bs},R{otherlab_row_bs},R{debtc_row_bs})'
    bs_sheet[f'T{start_row_bs}'].value = f'=SUM(T{ap_row_bs},T{accexp_row_bs},T{otherlab_row_bs},T{debtc_row_bs})'  
    bs_sheet[f'U{start_row_bs}'].value = f'=SUM(U{ap_row_bs},U{accexp_row_bs},U{otherlab_row_bs},U{debtc_row_bs})'

    start_row_bs += 1
    bs_sheet.row_dimensions[start_row_bs].height = 37 
    bs_sheet[f'D{start_row_bs}'] ='Long Term Debt'

    for row in data_activitybs: 
        if row['Activity'] == 'LTD':
            start_row_bs += 1
            bs_sheet[f'D{start_row_bs}'].style = indent_style
            bs_sheet[f'D{start_row_bs}'] = row['Description2']
            bs_sheet[f'G{start_row_bs}'] = row['total_bal9']
            bs_sheet[f'H{start_row_bs}'] = row['total_bal10']
            bs_sheet[f'I{start_row_bs}'] = row['total_bal11']
            bs_sheet[f'J{start_row_bs}'] = row['total_bal12']
            bs_sheet[f'K{start_row_bs}'] = row['total_bal1']
            bs_sheet[f'L{start_row_bs}'] = row['total_bal2']
            bs_sheet[f'M{start_row_bs}'] = row['total_bal3']
            bs_sheet[f'N{start_row_bs}'] = row['total_bal4']
            bs_sheet[f'O{start_row_bs}'] = row['total_bal5']
            bs_sheet[f'P{start_row_bs}'] = row['total_bal6']
            bs_sheet[f'Q{start_row_bs}'] = row['total_bal7']
            #bs_sheet[f'R{start_row_bs}'] = row['total_bal8']
  

    for row in data_balancesheet:
        if row['Activity'] == 'LTD':
            if row['Category'] == 'Debt':
                if row['Subcategory'] == 'Long Term Debt':
                    start_row_bs += 1
                    bs_sheet[f'D{start_row_bs}'].style = indent_style
                    bs_sheet[f'D{start_row_bs}'] = row['Description']
                    bs_sheet[f'F{start_row_bs}'] = row['FYE']
                    bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                    bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                    bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                    bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                    bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                    bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                    bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                    bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                    bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                    bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                    bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                    #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                    bs_sheet[f'T{start_row_bs}'] = row['fytd']
                    bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                    ltd_row_bs = start_row_bs

    start_row_bs += 1
    total_liabilites_row_bs = start_row_bs
    bs_sheet.row_dimensions[start_row_bs].height = 37 
    bs_sheet[f'D{start_row_bs}'].font = fontbold
    bs_sheet[f'D{start_row_bs}'] = 'Total Liabilities'
    bs_sheet[f'F{start_row_bs}'].value = f'=SUM(F{total_current_liabilites_row_bs},F{ltd_row_bs})'
    bs_sheet[f'G{start_row_bs}'].value = f'=SUM(G{total_current_liabilites_row_bs},G{ltd_row_bs})'
    bs_sheet[f'H{start_row_bs}'].value = f'=SUM(H{total_current_liabilites_row_bs},H{ltd_row_bs})'
    bs_sheet[f'I{start_row_bs}'].value = f'=SUM(I{total_current_liabilites_row_bs},I{ltd_row_bs})'
    bs_sheet[f'J{start_row_bs}'].value = f'=SUM(J{total_current_liabilites_row_bs},J{ltd_row_bs})'
    bs_sheet[f'K{start_row_bs}'].value = f'=SUM(K{total_current_liabilites_row_bs},K{ltd_row_bs})'
    bs_sheet[f'L{start_row_bs}'].value = f'=SUM(L{total_current_liabilites_row_bs},L{ltd_row_bs})'
    bs_sheet[f'M{start_row_bs}'].value = f'=SUM(M{total_current_liabilites_row_bs},M{ltd_row_bs})'
    bs_sheet[f'N{start_row_bs}'].value = f'=SUM(N{total_current_liabilites_row_bs},N{ltd_row_bs})'
    bs_sheet[f'O{start_row_bs}'].value = f'=SUM(O{total_current_liabilites_row_bs},O{ltd_row_bs})'
    bs_sheet[f'P{start_row_bs}'].value = f'=SUM(P{total_current_liabilites_row_bs},P{ltd_row_bs})'
    bs_sheet[f'Q{start_row_bs}'].value = f'=SUM(Q{total_current_liabilites_row_bs},Q{ltd_row_bs})'
    #bs_sheet[f'R{start_row_bs}'].value = f'=SUM(R{total_current_liabilites_row_bs},R{ltd_row_bs})'
    bs_sheet[f'T{start_row_bs}'].value = f'=SUM(T{total_current_liabilites_row_bs},T{ltd_row_bs})'  
    bs_sheet[f'U{start_row_bs}'].value = f'=SUM(U{total_current_liabilites_row_bs},U{ltd_row_bs})'


    for row in data_balancesheet:
        if row['Activity'] == 'Equity':
            if row['Category'] == 'Net Assets':
                
                start_row_bs += 1
                bs_sheet.row_dimensions[start_row_bs].height = 37 
                bs_sheet[f'D{start_row_bs}'].font = fontbold
                bs_sheet[f'D{start_row_bs}'] = 'Net Assets'
                
                bs_sheet[f'F{start_row_bs}'] = row['FYE']
                bs_sheet[f'G{start_row_bs}'] = row['difference_9']
                bs_sheet[f'H{start_row_bs}'] = row['difference_10']
                bs_sheet[f'I{start_row_bs}'] = row['difference_11']
                bs_sheet[f'J{start_row_bs}'] = row['difference_12']
                bs_sheet[f'K{start_row_bs}'] = row['difference_1']
                bs_sheet[f'L{start_row_bs}'] = row['difference_2']
                bs_sheet[f'M{start_row_bs}'] = row['difference_3']
                bs_sheet[f'N{start_row_bs}'] = row['difference_4']
                bs_sheet[f'O{start_row_bs}'] = row['difference_5']
                bs_sheet[f'P{start_row_bs}'] = row['difference_6']
                bs_sheet[f'Q{start_row_bs}'] = row['difference_7']
                #bs_sheet[f'R{start_row_bs}'] = row['difference_8']
                bs_sheet[f'T{start_row_bs}'] = row['fytd']
                bs_sheet[f'U{start_row_bs}'] = row['difference_6']
                net_assets_row_bs = start_row_bs

    start_row_bs += 1
    bs_sheet.row_dimensions[start_row_bs].height = 37 
    bs_sheet[f'D{start_row_bs}'].style = indent_style2
    bs_sheet[f'D{start_row_bs}'].font = fontbold
   
    bs_sheet[f'D{start_row_bs}'] = 'Total Liabilities and Net Assets'
    bs_sheet[f'F{start_row_bs}'].value = f'=SUM(F{total_liabilites_row_bs},F{net_assets_row_bs})'
    bs_sheet[f'G{start_row_bs}'].value = f'=SUM(G{total_liabilites_row_bs},G{net_assets_row_bs})'
    bs_sheet[f'H{start_row_bs}'].value = f'=SUM(H{total_liabilites_row_bs},H{net_assets_row_bs})'
    bs_sheet[f'I{start_row_bs}'].value = f'=SUM(I{total_liabilites_row_bs},I{net_assets_row_bs})'
    bs_sheet[f'J{start_row_bs}'].value = f'=SUM(J{total_liabilites_row_bs},J{net_assets_row_bs})'
    bs_sheet[f'K{start_row_bs}'].value = f'=SUM(K{total_liabilites_row_bs},K{net_assets_row_bs})'
    bs_sheet[f'L{start_row_bs}'].value = f'=SUM(L{total_liabilites_row_bs},L{net_assets_row_bs})'
    bs_sheet[f'M{start_row_bs}'].value = f'=SUM(M{total_liabilites_row_bs},M{net_assets_row_bs})'
    bs_sheet[f'N{start_row_bs}'].value = f'=SUM(N{total_liabilites_row_bs},N{net_assets_row_bs})'
    bs_sheet[f'O{start_row_bs}'].value = f'=SUM(O{total_liabilites_row_bs},O{net_assets_row_bs})'
    bs_sheet[f'P{start_row_bs}'].value = f'=SUM(P{total_liabilites_row_bs},P{net_assets_row_bs})'
    bs_sheet[f'Q{start_row_bs}'].value = f'=SUM(Q{total_liabilites_row_bs},Q{net_assets_row_bs})'
    #bs_sheet[f'R{start_row_bs}'].value = f'=SUM(R{total_liabilites_row_bs},R{net_assets_row_bs})'
    bs_sheet[f'T{start_row_bs}'].value = f'=SUM(T{total_liabilites_row_bs},T{net_assets_row_bs})'  
    bs_sheet[f'U{start_row_bs}'].value = f'=SUM(U{total_liabilites_row_bs},U{net_assets_row_bs})'


    


  


 

    workbook.save(generated_excel_path)

    # Serve the generated Excel file for download
    with open(generated_excel_path, 'rb') as excel_file:
        response = HttpResponse(excel_file.read(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        response['Content-Disposition'] = f'attachment; filename={os.path.basename(generated_excel_path)}'

    # Remove the generated Excel file (optional)
    os.remove(generated_excel_path)

    return response
