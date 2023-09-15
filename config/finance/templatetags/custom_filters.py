# custom_filters.py
from django import template

register = template.Library()

@register.filter(name='get_dict_value')
def get_dict_value(dictionary, key):
    try:
        return dictionary[key]
    except (KeyError, TypeError):
        return None
