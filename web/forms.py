from django import forms
from django.forms import fields
from .models import Photo

class UploadForm(forms.ModelForm):
    class Meta:
        model = Photo
        fields = "__all__"