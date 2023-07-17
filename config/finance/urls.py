from django.urls import path
from . import views

urlpatterns = [
    path('upload/', views.upload_view, name='upload'),
    path('login/',views.loginView,name = 'login'),
    path('login_form/',views.login_form,name = 'login_form'),
    path('logout/', views.logoutView, name='logout'),
    path('dashboard/',views.dashboard,name = 'dashboard'),
]