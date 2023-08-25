from django.contrib.auth.models import AbstractUser
from django.db import models
from ckeditor.fields import RichTextField


class User(AbstractUser):
    is_employee = models.BooleanField(default=False)
    is_admin = models.BooleanField(default=False)

class Item(models.Model):
    accomplishments = models.JSONField()
    activities = models.JSONField()
