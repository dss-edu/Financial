from django.shortcuts import redirect
from django.http import HttpResponse
from functools import wraps

def permission_required(view_func):
    def wrapper(request, *args, **kwargs):
        school = kwargs.get("school")
        user_role = request.session.get('user_role')

        if user_role == "admin" :
                return view_func(request, *args, **kwargs)
           
        if user_role == school :
                return view_func(request, *args, **kwargs)
        return HttpResponse(status=403)
    return wrapper
    
def custom_login_required(view_func):
    @wraps(view_func)
    def _wrapped_view(request, *args, **kwargs):
        if not request.session.get('user_role'):
            return redirect('login') 
        return view_func(request, *args, **kwargs)
    return _wrapped_view

