# ocr_app/forms.py
from django import forms
from .models import ReceiptImage

class ReceiptImageForm(forms.ModelForm):
    class Meta:
        model = ReceiptImage
        fields = ['image']
