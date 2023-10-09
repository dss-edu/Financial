# custom_filters.py
from django import template
from django.contrib.auth.models import Group

register = template.Library()


@register.filter(name="get_dict_value")
def get_dict_value(dictionary, key):
    try:
        return dictionary[key]
    except (KeyError, TypeError):
        return None


@register.filter(name="has_group")
def has_group(user, group_name):
    group = Group.objects.get(name=group_name)
    return True if group in user.groups.all() else False
