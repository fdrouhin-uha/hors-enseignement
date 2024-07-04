from django import forms

class UploadFilesForm(forms.Form):
    file1 = forms.FileField()
    file2 = forms.FileField()