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

def connect():
    server = 'aca-mysqlserver1.database.windows.net'
    database = 'Database1'
    username = 'aca-user1'
    password = 'Pokemon!123'
    port = '1433'
    
    driver = '{/usr/lib/libmsodbcsql-17.so}'

    cnxn = pyodbc.connect(f'DRIVER={driver};SERVER={server};DATABASE={database};UID={username};PWD={password}')
    return cnxn

def upload_view(request):
    if request.method == 'POST':
        form = UploadForm(request.POST, request.FILES)
        if form.is_valid():
            xlsb_file = request.FILES['xlsb_file']
            # Process the XLSB file
            with io.BytesIO(xlsb_file.read()) as f:
                excel_data = pd.ExcelFile(f, engine='pyxlsb')
                sheet_names = excel_data.sheet_names

                sheet_data = {}  # Dictionary to store sheet data
                for sheet_name in sheet_names:
                    df = excel_data.parse(sheet_name)
                    rows = df.values.tolist()
                    sheet_data[sheet_name] = rows

                # Create context data to pass to the template
                context = {
                    'sheet_names': sheet_names,
                    'sheet_data': sheet_data,
                }
                return render(request, 'your_template.html', context)
    else:
        form = UploadForm()
    return render(request, 'your_upload_form.html', {'form': form})



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
    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Definition_obj];") 
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
            'value': int(row[4]) if row[4] else 0  # Convert to int if not None, else set as 0
        }
        data.append(row_dict)

    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Definition_func];") 
    rows = cursor.fetchall()


    data2=[]
    for row in rows:
        row_dict = {
            'func_func': row[0],
            'desc': row[1],
            'budget': int(row[2]),
            
        }
        data2.append(row_dict)


    #
    cursor.execute("SELECT * FROM [dbo].[AscenderData_Advantage];") 
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



    acct_per_values2 = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']

    for item in data2:
        func = item['func_func']
        

        for i, acct_per in enumerate(acct_per_values2, start=1):
            item[f'total_func{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per
            ) 

    keys_to_check2 = ['total_func1', 'total_func2', 'total_func3', 'total_func4', 'total_func5','total_func6','total_func7','total_func8','total_func9','total_func10','total_func11','total_func12']

    for row in data2:
        for key in keys_to_check2:
            if row[key] > 0:
                row[key] = row[key]
            else:
                row[key] = ''            

      

    context = { 'data': data, 'data2':data2 , 'data3': data3}
    return render(request,'dashboard/pl_advantage.html', context)

def pl_cumberland(request):
    
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Definition_obj];") 
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
            'value': int(row[4]) if row[4] else 0  # Convert to int if not None, else set as 0
        }
        data.append(row_dict)

    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Definition_func];") 
    rows = cursor.fetchall()


    data2=[]
    for row in rows:
        row_dict = {
            'func_func': row[0],
            'desc': row[1],
            'budget': int(row[2]),
            
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



    acct_per_values2 = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10', '11', '12']

    for item in data2:
        func = item['func_func']
        

        for i, acct_per in enumerate(acct_per_values2, start=1):
            item[f'total_func{i}'] = sum(
                entry['Expend'] for entry in data3 if entry['func'] == func  and entry['AcctPer'] == acct_per
            ) 

    keys_to_check2 = ['total_func1', 'total_func2', 'total_func3', 'total_func4', 'total_func5','total_func6','total_func7','total_func8','total_func9','total_func10','total_func11','total_func12']

    for row in data2:
        for key in keys_to_check2:
            if row[key] > 0:
                row[key] = row[key]
            else:
                row[key] = ''            

      

    context = { 'data': data, 'data2':data2 , 'data3': data3}
    return render(request,'dashboard/pl_cumberland.html', context)



def insert_row(request):
    if request.method == 'POST':
        try:
            budget = request.POST.get('budget')
            fund = request.POST.get('fund')
            obj = request.POST.get('obj')
            description = request.POST.get('Description')
            category = 'Local' 
            

            if fund is None or fund == '':
                return JsonResponse({'status': 'error', 'message': 'fund value is missing'})
            if obj is None or obj =='':
                return JsonResponse({'status': 'error', 'message': 'obj value is missing'})
            if description is None  or description=='':
                return JsonResponse({'status': 'error', 'message': 'description value is missing'})
            if category is None   or category=='':
                return JsonResponse({'status': 'error', 'message': 'category value is missing'})
            if budget is None  or budget=='':
                return JsonResponse({'status': 'error', 'message': 'Budget value is missing'})


            


            cnxn = connect()
            cursor = cnxn.cursor()
            query = "INSERT INTO [dbo].[AscenderData_Definition_obj] (budget, fund, obj, Description, Category) VALUES (?, ?, ?, ?, ?)"
            cursor.execute(query, (budget, fund, obj, description, category))
            cnxn.commit()

            return JsonResponse({'message': 'Data inserted successfully'})

        except Exception as e:
            return JsonResponse({'status': 'error', 'message': str(e)})

    return JsonResponse({'status': 'error', 'message': 'Invalid request method'})

    

