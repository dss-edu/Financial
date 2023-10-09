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
from .decorators import permission_required

SCHOOLS = {
    "advantage": "ADVANTAGE ACADEMY",
    "cumberland": "CUMBERLAND ACADEMY",
    "village-tech": "VILLAGE TECH",
    "leadership": "Leadership Prep School",
    "manara": "MANARA ACADEMY",
}


@login_required
@permission_required
def dashboard_notes(request, school):
    cnxn = connect()
    cursor = cnxn.cursor()
    if request.method == "POST":
        notes = request.POST.getlist("notesList[]")
        data = json.dumps({i: note for i, note in enumerate(notes)})

    update_query = "UPDATE [dbo].[Report] SET notes = ? WHERE school = ?"
    cursor.execute(update_query, (data, school))
    # update_query = "UPDATE [dbo].[Report] SET notes = ?"
    # cursor.execute(update_query, data)

    cnxn.commit()
    cursor.close()
    cnxn.close()
    # return JsonResponse({1: "hello world"}, safe=False)
    return HttpResponse(status=200)


@login_required
@permission_required
def dashboard(request, school, anchor_year=""):
    data = {"accomplishments": "", "activities": "", "agendas": ""}

    cnxn = connect()
    cursor = cnxn.cursor()
    school_name = school

    if request.method == "POST":
        form = ReportsForm(request.POST)
        activities = request.POST.get("activities")
        accomplishments = request.POST.get("accomplishments")
        agendas = request.POST.get("agendas")

        update_query = "UPDATE [dbo].[Report] SET accomplishments = ?, activities = ?, agendas = ? WHERE school = ?"
        cursor.execute(
            update_query, (accomplishments, activities, agendas, school_name)
        )

        cnxn.commit()

        data["accomplishments"] = mark_safe(accomplishments)
        data["activities"] = mark_safe(activities)
        data["agendas"] = mark_safe(agendas)
    else:
        # check if it exists
        # query for the school
        query = "SELECT * FROM [dbo].[Report] WHERE school = ?"
        cursor.execute(query, school_name)
        row = cursor.fetchone()

        if row is None:
            # Insert query if it does noes exists
            insert_query = "INSERT INTO [dbo].[Report] (school, accomplishments, activities, agendas) VALUES (?, ?, ?, ?)"
            # insert_query = "INSERT INTO [dbo].[Report] (school, accomplishments, activities) VALUES (?, ?, ?)"
            accomplishments = "No accomplishments for this school yet. Click edit and add bullet points. It is important that the inserted accomplishments are in bullet points.\n"
            activities = "No activities for this school yet. Click edit and add bullet points. It is important that the inserted activities are in bullet points.\n"
            agendas = "No agenda for this school yet. Click edit and add bullet points. It is important that the inserted agenda are in bullet points.\n"

            # Execute the INSERT query
            cursor.execute(
                insert_query, (school_name, accomplishments, activities, agendas)
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

    context = modules.dashboard(school)

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

    # turn int into month name
    # context["month"] = calendar.month_name[context["month"]]

    cursor.close()
    cnxn.close()

    context["form"] = form
    context["data"] = data
    context["anchor_year"] = anchor_year
    return render(request, "temps/dashboard.html", context)


@login_required
@permission_required
def charter_first(request, school, anchor_year=""):
    context = modules.charter_first(school)
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

    context["anchor_year"] = anchor_year

    # current_date = datetime.today().date()
    # current_year = current_date.year
    # last_year = current_date - timedelta(days=365)
    # current_month = current_date.replace(day=1)
    # last_month = current_month - relativedelta(days=1)
    # last_month_number = last_month.month
    # ytd_budget_test = last_month_number + 3
    # ytd_budget = ytd_budget_test / 12
    # formatted_ytd_budget = (
    #     f"{ytd_budget:.2f}"  # Formats the float to have 2 decimal places
    # )
    #
    # if formatted_ytd_budget.startswith("0."):
    #     formatted_ytd_budget = formatted_ytd_budget[2:]
    # context = {
    #     "school": school,
    #     "school_name": SCHOOLS[school],
    #     "last_month": last_month,
    #     "last_month_number": last_month_number,
    #     "format_ytd_budget": formatted_ytd_budget,
    #     "ytd_budget": ytd_budget,
    # }
    return render(request, "temps/charter-first.html", context)


@login_required
@permission_required
def charter_first_charts(request, school):
    # if school = "advantage":
    context = {"school": school, "school_name": SCHOOLS[school]}
    return render(request, "temps/charter-first-charts.html", context)


@login_required
@permission_required
def profit_loss(request, school, anchor_year=""):
    context = modules.profit_loss(school, anchor_year)
    return render(request, "temps/profit-loss.html", context)


@login_required
@permission_required
def profit_loss_charts(request, school):
    # if school = "advantage":
    context = {"school": school, "school_name": SCHOOLS[school]}
    return render(request, "temps/profit-loss-charts.html", context)


@login_required
@permission_required
def balance_sheet(request, school, anchor_year=""):
    context = modules.balance_sheet(school, anchor_year)
    return render(request, "temps/balance-sheet.html", context)


@login_required
@permission_required
def balance_sheet_charts(request, school):
    context = {"school": school, "school_name": SCHOOLS[school]}
    return render(request, "temps/profit-loss-charts.html", context)


@login_required
@permission_required
def cashflow(request, school, anchor_year=""):
    context = modules.cashflow(school, anchor_year)

    return render(request, "temps/cashflow.html", context)


@login_required
@permission_required
def cashflow_charts(request, school):
    context = {"school": school, "school_name": SCHOOLS[school]}
    return render(request, "temps/profit-loss-charts.html", context)


@login_required
@permission_required
def general_ledger(request, school):
    context = modules.general_ledger(school)
    if school == "village-tech":
        return render(request, "temps/gl-vtech.html", context)
    return render(request, "temps/general-ledger.html", context)


@login_required
@permission_required
def manual_adjustments(request, school):
    context = modules.manual_adjustments(school)
    return render(request, "temps/manual-adjustments.html", context)


@login_required
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


@login_required
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


@login_required
@permission_required
def update_adjustments(request):
    pass


@login_required
@permission_required
def activity_edits(request, school):
    if request.method == "POST":
        body = json.loads(request.body)

        status = modules.activity_edits(school, body)

    if status:
        return HttpResponse(status=200)
    else:
        return HttpResponse(status=400)
