from django.urls import path
from . import views
from . import new_views


urlpatterns = [
    path("", views.loginView, name="login"),
    path("login/", views.loginView, name="login"),
    path("logout/", views.logoutView, name="logout"),
    path("home/schools", new_views.home, name="home"),
    path("change_password/<str:school>/", views.change_password, name="change_password"),
    path("users/", views.users, name="users"),
    path("add_user/", views.add_user, name="add_user"),
    path("view_user/<str:username>", views.view_user, name="view_user"),
    path("edit_user", views.edit_user, name="edit_user"),
    path("delete_user", views.delete_user, name="delete_user"),

    # path('pl_advantage/',views.pl_advantage,name = 'pl_advantage'),
    # path('pl_advantagechart/', views.pl_advantagechart, name='pl_advantagechart'),
    # path('first_advantage/', views.first_advantage, name='first_advantage'),
    # path('first_advantagechart/', views.first_advantagechart, name='first_advantagechart'),
    # path('gl_advantage/', views.gl_advantage, name='gl_advantage'),
    # path('bs_advantage/', views.bs_advantage, name='bs_advantage'),
    # path('cashflow_advantage/', views.cashflow_advantage, name='cashflow_advantage'),
    # path('dashboard_advantage/', views.dashboard_advantage, name='dashboard_advantage'),
    path("updatedb/", views.updatedb, name="updatedb"),
    path("updateschool/<str:school>", views.updateschool, name="updateschool"),
    path("updatefy/<str:school>", views.updatefy, name="updatefy"),
    path("updatefy/<str:school>/<year>", views.updatefy, name="updatefy"),
    path("insert_row/", views.insert_row, name="insert-row"),
    path("update_row/<str:school>/<int:year>", views.update_row, name="update_row"),
    path("delete/<str:fund>/<str:obj>/", views.delete, name="delete"),
    path("insert_bs_advantage/", views.insert_bs_advantage, name="insert_bs_advantage"),
    path(
        "delete_bs/<str:description>/<str:subcategory>/",
        views.delete_bs,
        name="delete_bs",
    ),
    path("delete_bsa/<str:obj>/<str:Activity>/", views.delete_bsa, name="delete_bsa"),

    path(
        "generate_excel/<str:school>",
        views.generate_excel,
        name="generate_excel",
    ),
    path(
        "generate_excel/<str:school>/<int:anchor_year>",
        views.generate_excel,
        name="generate_excel",
    ),
    path(
        "general-ledger/export/excel/<str:school>/<str:start>/<str:end>/",
        views.general_ledger_excel,
        name="general_ledger_excel",
    ),
    path(
        "general-ledger/export/excel/<str:school>/",
        views.general_ledger_excel,
        name="general_ledger_excel",
    ),
    path("delete_func/<str:func>/", views.delete_func, name="delete_func"),
    path(
        "viewgltotalrevenueytd/<str:school>/<str:year>/<str:url>/<str:category>/", views.viewgltotalrevenueytd, name="viewgltotalrevenueytd"
    ),
    path(
        "viewglrevenueytd/<str:fund>/<str:obj>/<str:school>/<str:year>/<str:url>/", views.viewglrevenueytd, name="viewglrevenueytd"
    ),
    path(
        "viewgl/<str:fund>/<str:obj>/<str:yr>/<str:school>/<str:year>/<str:url>/", views.viewgl, name="viewgl"
    ),
    path(
        "viewgl-all/<str:school>/<str:year>/<str:url>/", views.viewgl_all, name="viewgl-all"
    ),
    path(
        "viewgl-all/<str:school>/<str:year>/<str:url>/<str:yr>", views.viewgl_all, name="viewgl-all"
    ),
    path(
        "viewglfunc/<str:func>/<str:yr>/<str:school>/<str:year>/<str:url>/",
        views.viewglfunc,
        name="viewglfunc",
    ),
    path(
        "viewglfunc-all/<str:school>/<str:year>/<str:url>/",
        views.viewglfunc_all,
        name="viewglfunc-all",
    ),
    path(
        "viewglfunc-all/<str:school>/<str:year>/<str:url>/<str:yr>/",
        views.viewglfunc_all,
        name="viewglfunc-all",
    ),
    path(
        "viewgldna/<str:func>/<str:yr>/<str:school>/<str:year>/<str:url>/",
        views.viewgldna,
        name="viewgldna",
    ),
    path(
        "viewgl_activitybs/<str:yr>/<str:school>/<str:year>/<str:url>/",
        views.viewgl_activitybs,
        name="viewgl_activitybs",
    ),
    path(
        "viewglexpense/<str:obj>/<str:yr>/<str:school>/<str:year>/<str:url>/",
        views.viewglexpense,
        name="viewglexpense",
    ),
    path(
        "viewglexpense-all/<str:school>/<str:year>/<str:url>/",
        views.viewglexpense_all,
        name="viewglexpense-all",
    ),
    path(
        "viewglexpense-all/<str:school>/<str:year>/<str:url>/<str:yr>/",
        views.viewglexpense_all,
        name="viewglexpense-all",
    ),
    path(
        "download-csv/<str:school>/",
        views.download_csv,
        name="download_csv",
    ),
    # path("generate_excel/", new_views.generate_excel, name="generate_excel"),
    # path(
    #     "viewgl_cumberland/<str:fund>/<str:obj>/<str:yr>/",
    #     views.viewgl_cumberland,
    #     name="viewgl_cumberland",
    # ),
    # path(
    #     "viewglfunc_cumberland/<str:func>/<str:yr>/",
    #     views.viewglfunc_cumberland,
    #     name="viewglfunc_cumberland",
    # ),
    # path(
    #     "viewgl_activitybs_cumberland/<str:obj>/<str:yr>/",
    #     views.viewgl_activitybs_cumberland,
    #     name="viewgl_activitybs_cumberland",
    # ),
    # path(
    #     "viewglexpense_cumberland/<str:obj>/<str:yr>/",
    #     views.viewglexpense_cumberland,
    #     name="viewglexpense_cumberland",
    # ),
    # refactored urls
    path("dashboard/<str:school>", new_views.dashboard),
    path("dashboard/notes/<str:school>", new_views.dashboard_notes),
    path("dashboard/notes/<str:school>/<int:anchor_year>/<int:anchor_month>", new_views.dashboard_notes),
    path("dashboard/<str:school>/<int:anchor_year>", new_views.dashboard),
    path("dashboard/<str:school>/<int:anchor_year>/<int:anchor_month>", new_views.dashboard),
    path("charter-first/<str:school>", new_views.charter_first),
    path("charter-first/<str:school>/<int:anchor_year>", new_views.charter_first),
    path("charter-first-monthly/<str:school>/<int:anchor_year>/<int:anchor_month>", new_views.charter_first),
    path("charter-first-charts/<str:school>", new_views.charter_first_charts),
    path("profit-loss/<str:school>", new_views.profit_loss),
    path("profit-loss/<str:school>/<int:anchor_year>", new_views.profit_loss),
    path("profit-loss-monthly/<str:school>/<str:monthly>", new_views.profit_loss_monthly),
    path("ytd-expend/<str:school>", new_views.ytd_expend),
    path("ytd-expend/<str:school>/<int:anchor_year>", new_views.ytd_expend),
    path("profit-loss-date/<str:school>", new_views.profit_loss_date),
    path("profit-loss-date/<str:school>/<int:anchor_year>", new_views.profit_loss_date),
    path("profit-loss-charts/<str:school>", new_views.profit_loss_charts),
    path("profit-loss-charts/<str:school>/<int:anchor_year>", new_views.profit_loss_charts),
    path("bs/activity-edits/<str:school>", new_views.activity_edits),
    path("balance-sheet/<str:school>", new_views.balance_sheet),
    path("balance-sheet/<str:school>/<int:anchor_year>", new_views.balance_sheet),
    path("balance-sheet-monthly/<str:school>/<str:monthly>", new_views.balance_sheet_monthly),
    path("balance-sheet-asc/<str:school>", new_views.balance_sheet_asc),
    path("balance-sheet-asc/<str:school>/<int:anchor_year>", new_views.balance_sheet_asc),
    path("balance-sheet-charts/<str:school>", new_views.balance_sheet_charts),
    path("cashflow-statement/<str:school>", new_views.cashflow),
    path("cashflow-statement/<str:school>/<int:anchor_year>", new_views.cashflow),
    path("cashflow-statement-monthly/<str:school>/<str:monthly>", new_views.cashflow_monthly),
    path("cashflow-statement-charts/<str:school>", new_views.cashflow_charts),
    path("general-ledger/<str:school>", new_views.general_ledger),
    path("general-ledger-range/<str:school>/<str:date_start>/<str:date_end>", new_views.general_ledger_range),
    path("general-ledger-range/<str:school>", new_views.general_ledger_range),
    path("manual-adjustments/<str:school>", new_views.manual_adjustments),
    path("add-adjustments/", new_views.add_adjustments, name="add_adjustments"),
    path(
        "update-adjustments/", new_views.update_adjustments, name="update_adjustments"
    ),
    path(
        "delete-adjustments/", new_views.delete_adjustments, name="delete_adjustments"
    ),
    path(
        "home/analytics/", new_views.access_charts, name="access_charts"
    ),
    path(
        "access-date-count/", new_views.access_date_count, name="access_date_count"
    ),
    path(
        "school-status/<str:school>", new_views.all_schools, name="all_schools"
    ),
    path(

        "mockup/", views.mockup, name="mockup"
    ),
    path(
        "data-processing/<str:school>", new_views.data_processing, name="data_processing"
    ),
    path(
        "upload-data/<str:school>", views.upload_data, name="upload_data"
    ),
    path('download_file/<path:name>/<str:school>/', views.download_file, name='download_file'),

]
