from django.shortcuts import redirect
from django.http import HttpResponse


def permission_required(view_func):
    def wrapper(request, *args, **kwargs):
        school = kwargs.get("school")
        groups = None
        if request.user.groups.exists():
            groups = request.user.groups.all()

        for group in groups:
            if group.name == "admin":
                return view_func(request, *args, **kwargs)
            if group.name == school:
                return view_func(request, *args, **kwargs)
        return HttpResponse(status=403)

    return wrapper
