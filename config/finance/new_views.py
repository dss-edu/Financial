from django.shortcuts import render, redirect, get_object_or_404
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
from django.http import JsonResponse, HttpResponse
from django.views.decorators.csrf import csrf_exempt
import json
from datetime import datetime, timedelta
from dateutil.relativedelta import relativedelta
import shutil
import openpyxl
from django.conf import settings
from openpyxl.utils import get_column_letter
from bs4 import BeautifulSoup
from openpyxl.styles import Font, NamedStyle, Border, Side, Alignment
from .connect import connect
from . import modules

SCHOOLS = {
    "advantage": "ADVANTAGE ACADEMY",
    "cumberland": "CUMBERLAND ACADEMY",
    "village-tech": "VILLAGE TECH",
    "prepschool": "Leadership Prep School",
    "manara": "MANARA ACADEMY",
}


def dashboard(request, school):
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

        # update_query = "UPDATE [dbo].[Report] SET accomplishments = ?, activities = ? WHERE school = ?"
        # cursor.execute(update_query, (accomplishments, activities, school_name))
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
    form = ReportsForm(
        initial={
            "accomplishments": data["accomplishments"],
            "activities": data["activities"],
            "agendas": data["agendas"],
        }
    )

    cursor.close()
    cnxn.close()
    context = {
        "school": school,
        "school_name": SCHOOLS[school],
        "form": form,
        "data": data,
    }
    return render(request, "temps/dashboard.html", context)


def charter_first(request, school):
    current_date = datetime.today().date()
    current_year = current_date.year
    last_year = current_date - timedelta(days=365)
    current_month = current_date.replace(day=1)
    last_month = current_month - relativedelta(days=1)
    last_month_number = last_month.month
    ytd_budget_test = last_month_number + 3
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
        "format_ytd_budget": formatted_ytd_budget,
        "ytd_budget": ytd_budget,
    }
    return render(request, "temps/charter-first.html", context)


def charter_first_charts(request, school):
    # if school = "advantage":
    context = {"school": school, "school_name": SCHOOLS[school]}
    return render(request, "temps/charter-first-charts.html", context)


def profit_loss(request, school):
    context = modules.profit_loss(school)
    return render(request, "temps/profit-loss.html", context)


def profit_loss_charts(request, school):
    # if school = "advantage":
    context = {"school": school, "school_name": SCHOOLS[school]}
    return render(request, "temps/profit-loss-charts.html", context)


def balance_sheet(request, school):
    context = modules.balance_sheet(school)
    return render(request, "temps/balance-sheet.html", context)


def balance_sheet_charts(request, school):
    context = {"school": school, "school_name": SCHOOLS[school]}
    return render(request, "temps/profit-loss-charts.html", context)


def cashflow(request, school):
    context = modules.cashflow(school)

    return render(request, "temps/cashflow.html", context)


def cashflow_charts(request, school):
    context = {"school": school, "school_name": SCHOOLS[school]}
    return render(request, "temps/profit-loss-charts.html", context)


def general_ledger(request, school):
    if school == "advantage":
        context = modules.general_ledger()
        context["school"] = school
        context["school_name"] = SCHOOLS[school]
        return render(request, "temps/general-ledger.html", context)
