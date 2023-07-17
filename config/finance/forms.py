from django import forms

class UploadForm(forms.Form):
    xlsb_file = forms.FileField(label='Upload XLSB File')
    