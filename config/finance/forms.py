from django import forms
from ckeditor.widgets import CKEditorWidget

class UploadForm(forms.Form):
    xlsb_file = forms.FileField(label='Upload XLSB File')

class ReportsForm(forms.Form):
    activities = forms.CharField(widget=CKEditorWidget(config_name='financial_config', 
                                                       attrs={"id":"activities-form"}))
    accomplishments = forms.CharField(widget=CKEditorWidget(config_name='financial_config', 
                                                            attrs={"id":"accomplishments-form"}))
    
