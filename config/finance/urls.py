from django.urls import path
from . import views

urlpatterns = [
    path('login/',views.loginView,name = 'login'),
    path('login_form/',views.login_form,name = 'login_form'),
    path('logout/', views.logoutView, name='logout'),
    path('pl_advantage/',views.pl_advantage,name = 'pl_advantage'),
    path('pl_cumberland/',views.pl_cumberland,name = 'pl_cumberland'),
    path('insert_row/', views.insert_row, name='insert-row'),
    path('delete/<str:fund>/<str:obj>/', views.delete, name='delete'),
    
]
