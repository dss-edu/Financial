from django.urls import path
from . import views
from . import new_views


urlpatterns = [
    path("", views.loginView, name="login"),
    path("login/", views.loginView, name="login"),
    path("logout/", views.logoutView, name="logout"),
    path("change_password/<str:school>/", views.change_password, name="change_password"),
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
    path("delete_func/<str:func>/", views.delete_func, name="delete_func"),
    path(
        "viewgl/<str:fund>/<str:obj>/<str:yr>/<str:school>", views.viewgl, name="viewgl"
    ),
    path(
        "viewglfunc/<str:func>/<str:yr>/<str:school>/<str:year>",
        views.viewglfunc,
        name="viewglfunc",
    ),
    path(
        "viewgldna/<str:func>/<str:yr>/<str:school>/<str:year>",
        views.viewgldna,
        name="viewgldna",
    ),
    path(
        "viewgl_activitybs/<str:obj>/<str:yr>/",
        views.viewgl_activitybs,
        name="viewgl_activitybs",
    ),
    path(
        "viewglexpense/<str:obj>/<str:yr>/<str:school>",
        views.viewglexpense,
        name="viewglexpense",
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
    path("dashboard/<str:school>/<int:anchor_year>", new_views.dashboard),
    path("charter-first/<str:school>", new_views.charter_first),
    path("charter-first/<str:school>/<int:anchor_year>", new_views.charter_first),
    path("charter-first-charts/<str:school>", new_views.charter_first_charts),
    path("profit-loss/<str:school>", new_views.profit_loss),
    path("profit-loss/<str:school>/<int:anchor_year>", new_views.profit_loss),
    path("profit-loss-charts/<str:school>", new_views.profit_loss_charts),
    path("profit-loss-charts/<str:school>/<int:anchor_year>", new_views.profit_loss_charts),
    path("bs/activity-edits/<str:school>", new_views.activity_edits),
    path("balance-sheet/<str:school>", new_views.balance_sheet),
    path("balance-sheet/<str:school>/<int:anchor_year>", new_views.balance_sheet),
    path("balance-sheet-charts/<str:school>", new_views.balance_sheet_charts),
    path("cashflow-statement/<str:school>", new_views.cashflow),
    path("cashflow-statement/<str:school>/<int:anchor_year>", new_views.cashflow),
    path("cashflow-statement-charts/<str:school>", new_views.cashflow_charts),
    path("general-ledger/<str:school>", new_views.general_ledger),
    path("manual-adjustments/<str:school>", new_views.manual_adjustments),
    path("add-adjustments/", new_views.add_adjustments, name="add_adjustments"),
    path(
        "update-adjustments/", new_views.update_adjustments, name="update_adjustments"
    ),
    path(
        "delete-adjustments/", new_views.delete_adjustments, name="delete_adjustments"
    ),
]
