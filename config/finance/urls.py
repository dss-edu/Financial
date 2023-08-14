from django.urls import path
from . import views

urlpatterns = [
    
    path('login/',views.loginView,name = 'login'),
    path('login_form/',views.login_form,name = 'login_form'),
    path('logout/', views.logoutView, name='logout'),
    path('pl_advantage/',views.pl_advantage,name = 'pl_advantage'),
    path('pl_advantagechart/', views.pl_advantagechart, name='pl_advantagechart'),
    path('gl_advantage/', views.gl_advantage, name='gl_advantage'),
    path('bs_advantage/', views.bs_advantage, name='bs_advantage'),
    path('cashflow_advantage/', views.cashflow_advantage, name='cashflow_advantage'),

    
    path('insert_row/', views.insert_row, name='insert-row'),
    path('delete/<str:fund>/<str:obj>/', views.delete, name='delete'),
    path('insert_bs_advantage/', views.insert_bs_advantage, name='insert_bs_advantage'),
    path('delete_bs/<str:description>/<str:subcategory>/', views.delete_bs, name='delete_bs'),
    path('delete_bsa/<str:obj>/<str:Activity>/', views.delete_bsa, name='delete_bsa'),
  
    path('delete_func/<str:func>/', views.delete_func, name='delete_func'),
    path('viewgl/<str:fund>/<str:obj>/<str:yr>/', views.viewgl, name='viewgl'),
    path('viewglfunc/<str:func>/<str:yr>/', views.viewglfunc, name='viewglfunc'),
    path('viewgl_activitybs/<str:obj>/<str:yr>/', views.viewgl_activitybs, name='viewgl_activitybs'),
    path('viewglexpense/<str:obj>/<str:yr>/', views.viewglexpense, name='viewglexpense'),

    

    
    path('pl_cumberlandchart/', views.pl_cumberlandchart, name='pl_cumberlandchart'),
    path('pl_cumberland/',views.pl_cumberland,name = 'pl_cumberland'),

]
