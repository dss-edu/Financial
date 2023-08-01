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
        valueformat = "{:,.0f}".format(int(row[4])) if row[4] else ""
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
        budgetformat = "{:,.0f}".format(int(row[3])) if row[3] else ""
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

    keys_to_check2 = ['total_func1', 'total_func2', 'total_func3', 'total_func4', 'total_func5','total_func6','total_func7','total_func8','total_func9','total_func10','total_func11','total_func12']

    for row in data2:
        for key in keys_to_check2:
            if row[key] > 0:
                row[key] = row[key]
            else:
                row[key] = ''
    for row in data2:
        for key in keys_to_check2:
            if row[key] != "":
                row[key] = "{:,.0f}".format(row[key])

                

    lr_funds = list(set(row['fund'] for row in data3 if 'fund' in row))
    lr_funds_sorted = sorted(lr_funds)
    lr_obj = list(set(row['obj'] for row in data3 if 'obj' in row))
    lr_obj_sorted = sorted(lr_obj)

    func_choice = list(set(row['func'] for row in data3 if 'func' in row))
    func_choice_sorted = sorted(func_choice)
    
            

    context = { 'data': data, 'data2':data2 , 'data3': data3 , 'lr_funds':lr_funds_sorted, 'lr_obj':lr_obj_sorted, 'func_choice':func_choice_sorted }
    return render(request,'dashboard/pl_advantage.html', context)

def pl_cumberland(request):
    
    cnxn = connect()
    cursor = cnxn.cursor()
    cursor.execute("SELECT  * FROM [dbo].[AscenderData_Cumberland_Definition_obj];") 
    rows = cursor.fetchall()

    
    data = []
    for row in rows:
        if row[4] is None:
            row[4] = ''
        valueformat = "{:,.0f}".format(int(row[4])) if row[4] else ""
            
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
        budgetformat = "{:,.0f}".format(int(row[3])) if row[3] else ""
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

    keys_to_check2 = ['total_func1', 'total_func2', 'total_func3', 'total_func4', 'total_func5','total_func6','total_func7','total_func8','total_func9','total_func10','total_func11','total_func12']

    for row in data2:
        for key in keys_to_check2:
            if row[key] > 0:
                row[key] = row[key]
            else:
                row[key] = ''
    for row in data2:
        for key in keys_to_check2:
            if row[key] != "":
                row[key] = "{:,.0f}".format(row[key])



      
    lr_funds = list(set(row['fund'] for row in data3 if 'fund' in row))
    lr_funds_sorted = sorted(lr_funds)
    lr_obj = list(set(row['obj'] for row in data3 if 'obj' in row))
    lr_obj_sorted = sorted(lr_obj)

    func_choice = list(set(row['func'] for row in data3 if 'func' in row))
    func_choice_sorted = sorted(func_choice)
    
            

    context = { 'data': data, 'data2':data2 , 'data3': data3 , 'lr_funds':lr_funds_sorted, 'lr_obj':lr_obj_sorted, 'func_choice':func_choice_sorted }
    return render(request,'dashboard/pl_cumberland.html', context)




    

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


def pl_advantagechart(request):
    return render(request,'dashboard/pl_advantagechart.html')

def pl_cumberlandchart(request):
    return render(request,'dashboard/pl_cumberlandchart.html')

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

            gl_data.append(row_dict)
        
        context = { 'gl_data':gl_data}

        
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
    
    

            

    context = { 
        
        'data3': data3 , 
         }
    return render(request,'dashboard/gl_advantage.html', context)

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


    


    
            

    
    
            

    context = { 
        
        'data_balancesheet': data_balancesheet ,
        'data_activitybs': data_activitybs,
        'data3': data3,
         }

    return render(request,'dashboard/bs_advantage.html', context)

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

            glbs_data.append(row_dict)
        
        context = { 'glbs_data':glbs_data}

        
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

            glfunc_data.append(row_dict)
        
        context = { 'glfunc_data':glfunc_data}

        
        cursor.close()
        cnxn.close()

        return JsonResponse({'status': 'success', 'data': context})

    except Exception as e:
        return JsonResponse({'status': 'error', 'message': str(e)})

