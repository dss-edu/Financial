from django.shortcuts import render, redirect, get_object_or_404
from .forms import ReportsForm
from django.utils.safestring import mark_safe
from django.http import JsonResponse, HttpResponse
import datetime
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
from .connect import connect
from . import modules
from django.views.decorators.csrf import requires_csrf_token, csrf_exempt
from django.contrib.auth.decorators import login_required
import calendar
import json
from .decorators import permission_required,custom_login_required
from config import settings
from django.core.files.storage import default_storage
from django.core.files.storage import FileSystemStorage
import os
from azure.storage.blob import BlobServiceClient, BlobClient, ContainerClient
from io import BytesIO


SCHOOLS = settings.SCHOOLS
db = settings.db
schoolCategory = settings.schoolCategory
schoolMonths = settings.schoolMonths
school_fye = settings.school_fye
media_root = settings.MEDIA_ROOT
JSON_DIR = FileSystemStorage(location=media_root)

present_date = datetime.today().date()   
present_year = present_date.year
present_year = int(present_year)
present_month = present_date.month
curr_fy = int(present_year)
if present_month == 1:
    last_month_number = 12
    curr_fy = curr_fy - 1  
else:
    last_month_number = present_month - 1


def getStatusCode(school):
    db_string = (db[school]['db'])
    db_string = db_string.strip('[]')
    cnxn = connect()
    cursor = cnxn.cursor()
    query = "SELECT * FROM [dbo].[AscenderDownloader] WHERE db = ?"
    cursor.execute(query,db_string)
    print(db_string)
    row = cursor.fetchone()
    status = ""
    if row:
        status = row[3]


    return status


def dashboard_notes(request, school , anchor_year="" , anchor_month=""):

    cnxn = connect()
    cursor = cnxn.cursor()
    if request.method == "POST":
        notes = request.POST.getlist("notesList[]")
        data = json.dumps({i: note for i, note in enumerate(notes)})

    if anchor_month:
        # select_query = "SELECT * FROM [dbo].[Reports] WHERE school = ? and year = ? and month = ?"
        # cursor.execute(select_query,(school,anchor_year, anchor_month))
        # row = cursor.fetchone()
        # if row is None:
        #     insert_query = "INSERT INTO [dbo].[Reports] (school, accomplishments, activities, agendas,notes,year, month) VALUES (?, ?, ?, ?, ?, ?, ?)"
        #     cursor.execute(insert_query, ( school,"","","","",anchor_year ,anchor_month ))
            
        update_query = "UPDATE [dbo].[Reports] SET notes = ? WHERE school = ? and year = ? and month = ?"
        cursor.execute(update_query, (data, school,anchor_year ,anchor_month ))
   
    else:
        # select_query = "SELECT * FROM [dbo].[Reports] WHERE school = ? and year = ? and month = ?"
        # cursor.execute(select_query,(school,curr_fy, last_month_number))
        # row = cursor.fetchone()
        # if row is None:
        #     insert_query = "INSERT INTO [dbo].[Reports] (school, accomplishments, activities, agendas,notes,year, month) VALUES (?, ?, ?, ?, ?, ?, ?)"
        #     cursor.execute(insert_query, ( school,"","","","",curr_fy ,last_month_number ))

        update_query = "UPDATE [dbo].[Reports] SET notes = ? WHERE school = ? and year = ? and month = ?"
        cursor.execute(update_query, (data, school,curr_fy,last_month_number))
      
    # update_query = "UPDATE [dbo].[Report] SET notes = ?"
    # cursor.execute(update_query, data)

    cnxn.commit()
    cursor.close()
    cnxn.close()
    # return JsonResponse({1: "hello world"}, safe=False)
    return HttpResponse(status=200)


@custom_login_required
@permission_required
def dashboard(request, school, anchor_year="",anchor_month=""):
    data = {"accomplishments": "", "activities": "", "agendas": ""}
   
    
    cnxn = connect()
    cursor = cnxn.cursor()
    school_name = school

    if request.method == "POST":
        form = ReportsForm(request.POST)
        activities = request.POST.get("activities")
        accomplishments = request.POST.get("accomplishments")
        agendas = request.POST.get("agendas")

        update_query = "UPDATE [dbo].[Reports] SET accomplishments = ?, activities = ?, agendas = ? WHERE school = ? and year = ? and month = ?"
        if anchor_month:
            cursor.execute(
                update_query, (accomplishments, activities, agendas, school_name,anchor_year,anchor_month)
            )
        else:
            cursor.execute(
                update_query, (accomplishments, activities, agendas, school_name,curr_fy,last_month_number)
            )

        cnxn.commit()

        data["accomplishments"] = mark_safe(accomplishments)
        data["activities"] = mark_safe(activities)
        data["agendas"] = mark_safe(agendas)
    else:
        # check if it exists
        # query for the school
        query = "SELECT * FROM [dbo].[Reports] WHERE school = ? and year =? and month=?"
        if anchor_month:            
            cursor.execute(query, school_name,anchor_year,anchor_month)
        else:
            cursor.execute(query, school_name,curr_fy,last_month_number)
        row = cursor.fetchone()
       
        if row is None:
            # Insert query if it does noes exists
            insert_query = "INSERT INTO [dbo].[Reports] (school, accomplishments, activities, agendas, year, month) VALUES (?, ?, ?, ?,?,?)"
            # insert_query = "INSERT INTO [dbo].[Report] (school, accomplishments, activities) VALUES (?, ?, ?)"
            accomplishments = "No accomplishments for this school yet. Click edit and add bullet points. It is important that the inserted accomplishments are in bullet points.\n"
            activities = "No activities for this school yet. Click edit and add bullet points. It is important that the inserted activities are in bullet points.\n"
            agendas = "No agenda for this school yet. Click edit and add bullet points. It is important that the inserted agenda are in bullet points.\n"

            # Execute the INSERT query
            if anchor_month:
                cursor.execute(
                    insert_query, (school_name, accomplishments, activities, agendas,anchor_year,anchor_month)
                )
            else:
                cursor.execute(
                    insert_query, (school_name, accomplishments, activities, agendas,curr_fy,last_month_number)
                )

            # Commit the transaction
            cnxn.commit()
       
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
            if row[4]:
                data["notes"] = json.loads(row[4])

    # form = CKEditorForm(initial={'form_field_name': initial_content})
    form = ReportsForm(
        initial={
            "accomplishments": data["accomplishments"],
            "activities": data["activities"],
            "agendas": data["agendas"],
        }
    )

    if anchor_month:
        context = modules.dashboard(school,anchor_year, anchor_month)
     
        
    else:

        context = modules.dashboard(school,anchor_year, anchor_month)



    net_ytd = context["net_income_ytd"]
    net_earnings = context["net_earnings"]
    
    if net_ytd:
        if net_ytd < 0:
            context["net_income_ytd"] = f"$({net_ytd * -1:.0f})"
        else:
            context["net_income_ytd"] = f"${net_ytd:.0f}"

    if net_earnings:
        if net_earnings < 0:
            context["net_earnings"] = f"$({net_earnings * -1:.0f})"
        else:
            context["net_earnings"] = f"${net_earnings:.0f}"

    # turn int into month name
    # context["month"] = calendar.month_name[context["month"]]
    print(context["anchor_year"])
    cursor.close()
    cnxn.close()
    
    context["form"] = form
    context["data"] = data
    context["anchor_year"] = anchor_year
    context["anch_month"] = anchor_month
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username
    context["ascender"] = 'True'
    if school in schoolCategory["skyward"]:
        context["ascender"] = 'False'
    context["iconStatusCode"] = getStatusCode(school)
   
    # title="code : {{ iconStatusCode }}"
    return render(request, "temps/dashboard.html", context)


@custom_login_required
@permission_required
def charter_first(request, school, anchor_year="",anchor_month=""):
    context = modules.charter_first(school,anchor_year,anchor_month)


    context["anchor_year"] = anchor_year


    
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username
    context["ascender"] = 'True'
    if school in schoolCategory["skyward"]:
        context["ascender"] = 'False'
    context["iconStatusCode"] = getStatusCode(school)
    return render(request, "temps/charter-first.html", context)


@custom_login_required
@permission_required
def charter_first_charts(request, school):
    # if school = "advantage":
    context = {"school": school, "school_name": SCHOOLS[school]}
    role = request.session.get('user_role')
    context["role"] = role
    context["ascender"] = 'True'
    if school in schoolCategory["skyward"]:
        context["ascender"] = 'False'
    return render(request, "temps/charter-first-charts.html", context)




@custom_login_required
@permission_required
def profit_loss(request, school, anchor_year=""):
    context = modules.profit_loss(school, anchor_year)
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username

    context["September"] = 'True'
    if school in schoolMonths["julySchool"]:
        context["September"] = 'False'
    context["present_year"] = present_year
    
    context["ascender"] = 'True'
    if school in schoolCategory["skyward"]:
        context["ascender"] = 'False'


    context["iconStatusCode"] = getStatusCode(school)
    return render(request, "temps/profit-loss.html", context)


@custom_login_required
@permission_required
def ytd_expend(request, school, anchor_year=""):
    context = modules.profit_loss(school, anchor_year)
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username

    context["September"] = 'True'
    if school in schoolMonths["julySchool"]:
        context["September"] = 'False'
    context["present_year"] = present_year
    
    context["ascender"] = 'True'
    if school in schoolCategory["skyward"]:
        context["ascender"] = 'False'


    context["iconStatusCode"] = getStatusCode(school)
    return render(request, "temps/ytd-expend.html", context)


@custom_login_required
@permission_required
def profit_loss_date(request, school, anchor_year=""):
    context = modules.profit_loss_date(school, anchor_year)
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username

    context["September"] = 'True'
    if school in schoolMonths["julySchool"]:
        context["September"] = 'False'

    context["present_year"] = present_year
    context["ascender"] = 'True'
    if school in schoolCategory["skyward"]:
        context["ascender"] = 'False'
    return render(request, "temps/profit-loss.html", context)

@custom_login_required
@permission_required
def profit_loss_charts(request, school, anchor_year=""):
    context = modules.profit_loss_chart(school, anchor_year)
    # context = {"school": school, "school_name": SCHOOLS[school]}
    net_ytd = context["net_income_ytd"]
    anchor_month = ""
    first_context = modules.charter_first(school,anchor_year,anchor_month)
    # first_context["debt_capitalization"] = modules.percent_to_ratio(
    #     first_context["debt_capitalization"]
    # )
    # first_context["total_assets"] = modules.float_to_ratio(
    #     first_context["total_assets"]
    # )
    # first_context["debt_service"] = modules.float_to_ratio(
    #     first_context["debt_service"]
    # )
    # first_context["current_assets"] = modules.float_to_ratio(
    #     first_context["current_assets"]
    # )

    if net_ytd < 0:
        context["net_income_ytd"] = f"${net_ytd * -1:,.0f}"
    else:
        context["net_income_ytd"] = f"${net_ytd:,.0f}"

    context["first_context"] = first_context

    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username
    context["ascender"] = 'True'
    if school in schoolCategory["skyward"]:
        context["ascender"] = 'False'
    return render(request, "temps/profit-loss-charts.html", context)


@custom_login_required
@permission_required
def balance_sheet(request, school, anchor_year=""):
    
    context = modules.balance_sheet(school, anchor_year)
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username
    context["September"] = 'True'
    if school in schoolMonths["julySchool"]:
        context["September"] = 'False'
    context["present_year"] = present_year
    context["ascender"] = 'True'
    if school in schoolCategory["skyward"]:
        context["ascender"] = 'False'
    context["school_bs"] = "False"
    if school in school_fye:
        context["school_bs"] = "True"
    context["school_bs_asc"] = ""

    
    context["iconStatusCode"] = getStatusCode(school)
    return render(request, "temps/balance-sheet.html", context)

@custom_login_required
@permission_required
def balance_sheet_asc(request, school, anchor_year=""):
    
    context = modules.balance_sheet_asc(school, anchor_year)
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username
    context["September"] = 'True'
    if school in schoolMonths["julySchool"]:
        context["September"] = 'False'
    context["present_year"] = present_year
    context["ascender"] = 'True'
    if school in schoolCategory["skyward"]:
        context["ascender"] = 'False'

    school_fye = []
    context["school_bs"] = ""
    
    context["school_bs_asc"] = "False"
    if school in schoolCategory["ascender"]:
        context["school_bs_asc"] = "True"
    context["iconStatusCode"] = getStatusCode(school)
    return render(request, "temps/balance-sheet.html", context)


@custom_login_required
@permission_required
def balance_sheet_charts(request, school, anchor_year=""):
    context = modules.profit_loss_chart(school, anchor_year)
    # context = {"school": school, "school_name": SCHOOLS[school]}
    net_ytd = context["net_income_ytd"]
    anchor_month=""
    first_context = modules.charter_first(school,anchor_year,anchor_month)
    # first_context["debt_capitalization"] = modules.percent_to_ratio(
    #     first_context["debt_capitalization"]
    # )
    # first_context["total_assets"] = modules.float_to_ratio(
    #     first_context["total_assets"]
    # )
    # first_context["debt_service"] = modules.float_to_ratio(
    #     first_context["debt_service"]
    # )
    # first_context["current_assets"] = modules.float_to_ratio(
    #     first_context["current_assets"]
    # )

    if net_ytd < 0:
        context["net_income_ytd"] = f"${net_ytd * -1:,.0f}"
    else:
        context["net_income_ytd"] = f"${net_ytd:,.0f}"

    context["first_context"] = first_context
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username
    context["ascender"] = 'True'
    if school in schoolCategory["skyward"]:
        context["ascender"] = 'False'
    return render(request, "temps/profit-loss-charts.html", context)

@custom_login_required
@permission_required
def cashflow(request, school, anchor_year=""):
    context = modules.cashflow(school, anchor_year)
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username
    context["September"] = 'True'
    if school in schoolMonths["julySchool"]:
        context["September"] = 'False'
    context["present_year"] = present_year
    context["ascender"] = 'True'
    if school in schoolCategory["skyward"]:
        context["ascender"] = 'False'

    
    context["school_bs"] = "False"
    if school in school_fye:
        context["school_bs"] = "True"
    context["iconStatusCode"] = getStatusCode(school)    
    return render(request, "temps/cashflow.html", context)


@custom_login_required
@permission_required
def cashflow_charts(request, school, anchor_year=""):
    context = modules.profit_loss_chart(school, anchor_year)
    # context = {"school": school, "school_name": SCHOOLS[school]}
    net_ytd = context["net_income_ytd"]

    first_context = modules.charter_first(school)
    # first_context["debt_capitalization"] = modules.percent_to_ratio(
    #     first_context["debt_capitalization"]
    # )
    # first_context["total_assets"] = modules.float_to_ratio(
    #     first_context["total_assets"]
    # )
    # first_context["debt_service"] = modules.float_to_ratio(
    #     first_context["debt_service"]
    # )
    # first_context["current_assets"] = modules.float_to_ratio(
    #     first_context["current_assets"]
    # )

    if net_ytd < 0:
        context["net_income_ytd"] = f"${net_ytd * -1:,.0f}"
    else:
        context["net_income_ytd"] = f"${net_ytd:,.0f}"

    context["first_context"] = first_context
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username
    context["ascender"] = 'True'
    if school in schoolCategory["skyward"]:
        context["ascender"] = 'False'
    return render(request, "temps/profit-loss-charts.html", context)

@custom_login_required
@permission_required
def general_ledger(request, school):
    context = modules.general_ledger(school)
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username
    if school in schoolCategory["skyward"]:
        return render(request, "temps/gl-vtech.html", context)
    return render(request, "temps/general-ledger.html", context)

@custom_login_required
@permission_required
def general_ledger_range(request, school, date_start="", date_end=""):
    try:
        data = modules.general_ledger(school, date_start, date_end)

        return JsonResponse(data["data3"], safe=False)


    except Exception as e:
        print(e)
        return JsonResponse({'message': str(e)})
    



@custom_login_required
@permission_required
def manual_adjustments(request, school):
    context = modules.manual_adjustments(school)
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username
    
    return render(request, "temps/manual-adjustments.html", context)


@custom_login_required
@permission_required
def add_adjustments(request):
    if request.method == "POST":
        school = request.POST.get("school")

        # save to database
        values = []
        floats = ["est", "acct-real", "appr", "expend", "bal"]
        for key in request.POST.keys():
            if key == "csrfmiddlewaretoken":
                continue
            if key in floats:
                values.append(
                    float(request.POST.get(key)) if request.POST.get(key) else 0
                )
            else:
                values.append(request.POST.get(key))
        cnxn = connect()
        cursor = cnxn.cursor()

        insert_query = f"INSERT INTO [dbo].[Adjustment] (fund, func, obj, org, fscl_yr, pgm, edSpan, projDtl, AcctDescr, Number, Date, AcctPer, Real, Expend, Bal, WorkDescr, Type, School) VALUES ({','.join(['?' for i in range(len(request.POST.keys())-1)])})"
        cursor.execute(insert_query, tuple(values))
        cnxn.commit()

        cursor.close()
        cnxn.close()
    # return redirect(f"/manual-adjustments/{school}")
    return HttpResponse(status=200)


@custom_login_required
@permission_required
def delete_adjustments(request):
    if request.method == "POST":
        rows = json.loads(request.POST.get("adjustments"))
        school = request.POST.get("school")
        floats = ["est", "acct-real", "appr", "expend", "bal"]

        for row in rows:
            for key, val in row.items():
                if key in floats:
                    row[key] = float(val)

        cnxn = connect()
        cursor = cnxn.cursor()

        # columns = "fund, func, obj, org, fscl_yr, pgm, edSpan, projDtl, AcctDescr, Number, Date, AcctPer, Real, Expend, Bal, WorkDescr, Type, School"
        columns = [
            "fund",
            "func",
            "obj",
            "org",
            "fscl_yr",
            "pgm",
            "edSpan",
            "projDtl",
            "AcctDescr",
            "Number",
            "Date",
            "AcctPer",
            "Real",
            "Expend",
            "Bal",
            "WorkDescr",
            "Type",
            "School",
        ]

        delete_query = f"DELETE FROM [dbo].[Adjustment] WHERE { ' AND '.join([col + ' = ?' for col in columns]) };"

        for row in rows:
            if row["Date"] == "None":
                temp_delete_query = delete_query.replace("Date = ?", "Date IS NULL")
                del row["Date"]

                cursor.execute(temp_delete_query, tuple(row.values()))
                cnxn.commit()
                continue

            input_string = row["Date"]
            input_format = "%m-%d-%Y"
            datetime_obj = datetime.strptime(input_string, input_format)
            row["Date"] = datetime_obj

            cursor.execute(delete_query, tuple(row.values()))
            cnxn.commit()

        cursor.close()
        cnxn.close()
    return HttpResponse(status=200)


@custom_login_required
@permission_required
def update_adjustments(request):
    pass


@custom_login_required
@permission_required
def activity_edits(request, school):
    if request.method == "POST":
        body = json.loads(request.body)

        status = modules.activity_edits(school, body)

    if status:
        return HttpResponse(status=200)
    else:
        return HttpResponse(status=400)


@custom_login_required
@permission_required
def access_charts(request):
    # context = {
    #     'school': school,
    # }
    # close connection
    role = request.session.get('user_role')

    username = request.session.get('username')
    

    context = {
        "role":role,
        "username": username,
    }
    return render(request, "temps/access-charts.html",context)

@custom_login_required
@permission_required
def access_date_count(request):
    if request.method == 'GET':
        cnxn = connect()
        cursor = cnxn.cursor()
        query = """
        SELECT CAST(access_date AS DATE) AS date_only, school, COUNT(*) FROM [dbo].[Access_Logs]
        WHERE username != 'admin'
        GROUP BY CAST(access_date AS DATE), school
        ORDER BY date_only;
        """
        try: 
            date_group = list(cursor.execute(query))
        except Exception as e:
            print(e)
        finally:
            cursor.close()
            cnxn.close()
        data = {}
        for i, (date,school, count) in enumerate(date_group):
            if data.get(school, False):
                data[school].append({
                    "date": date,
                    "count": count
                })
            else:
                data[school] = [{
                    "date":date,
                    "count": count
                }]

    return JsonResponse(data, safe=False)


@custom_login_required
@permission_required
def all_schools(request, school):
    context = {}
    JSON_DIR = os.path.join(settings.BASE_DIR, "finance","json","school-status")
    files = os.listdir(JSON_DIR)
    for file in files:
        with open(os.path.join(JSON_DIR, file), "r") as f:
            basename = os.path.splitext(file)[0]
            context[basename] = json.load(f)

    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username

    return render(request, "temps/schools.html", context)

@custom_login_required
@permission_required
def home(request):

    school_data = []
    for key,name in SCHOOLS.items():
        row_data = {
            "school_key":key,
            "school_name":name
        }
        school_data.append(row_data)


    def custom_sort(entry):

        if entry["school_key"] == "advantage":
            return (0, entry["school_key"])
        else:
            return (1, entry["school_key"])
    sorted_school_data = sorted(school_data, key=custom_sort)
    context = {
        'schools': sorted_school_data
    }
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username

    return render(request, "temps/home.html", context)



@custom_login_required
@permission_required
def data_processing(request,school):

    context = {}
    username = request.session.get('username')
   
    storage_connection_string = "DefaultEndpointsProtocol=https;AccountName=blogdataprocessing;AccountKey=qI+FDF3NPvo6SkYUpr00yw4MiQB0lofHF+lmnZ+8S686FPXBUTMJZUo31N3KoNefctV/rfR0dTFe+AStBaDTpA==;EndpointSuffix=core.windows.net"
    
    # Create a BlobServiceClient
    blob_service_client = BlobServiceClient.from_connection_string(storage_connection_string)

    container_id = school
    container_client = blob_service_client.get_container_client(container_id)

    # Get a list of blobs in the container
    blobs = container_client.list_blobs()
    
    context['blobs'] = blobs

    cnxn = connect()
    cursor = cnxn.cursor()

    side_query = "SELECT * FROM [dbo].[InvoiceSubmission] WHERE [user] = ? "
    cursor.execute(side_query,(username))
    # query = "SELECT * FROM [dbo].[InvoiceSubmission] WHERE blobPath LIKE ? "
    # blob_url_pattern = f"blob-{school}/%"
    # cursor.execute(query,(blob_url_pattern))
    rows = cursor.fetchall()
    file_data = []
    for row in rows:
        
        doc = {
            "PO_Number":row[0],
            "blobPath":row[1],
            "client":row[2],
            "user":row[3],
            "status":row[4],
            "logs":row[5],
            "fileName":row[6],
            "Date":row[7]


        }

        file_data.append(doc)


    side_query = "SELECT * FROM [dbo].[InvoiceSubmission] WHERE [user] = ? "
    cursor.execute(side_query,(username))
    rows = cursor.fetchall()

    side_data = []

    for row in rows:

        doc = {
            "PO_Number":row[0],
            "blobPath":row[1],
            "client":row[2],
            "user":row[3],
            "status":row[4],
            "logs":row[5],
            "fileName":row[6],
            "Date":row[7]

        }
        print(row[0])
        side_data.append(doc)









    cursor.close()
    cnxn.close()

    context["side_data"] = side_data
    context["file_data"] = file_data
    context['school'] = school
    role = request.session.get('user_role')
    context["role"] = role
    username = request.session.get('username')
    context["username"] = username

    return render(request, "temps/data-processing.html", context)