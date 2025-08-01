from django import forms
from .models import ReceiptImage

class ReceiptImageForm(forms.ModelForm):
    cash = forms.DecimalField(
        max_digits=10, 
        decimal_places=2, 
        required=True, 
        label='Cash',
        widget=forms.NumberInput(attrs={'type': 'number', 'step': '0.01', 'inputmode': 'decimal'})
    )
    card = forms.DecimalField(
        max_digits=10, 
        decimal_places=2, 
        required=True, 
        label='Card',
        widget=forms.NumberInput(attrs={'type': 'number', 'step': '0.01', 'inputmode': 'decimal'})
    )
    cupon = forms.DecimalField(
        max_digits=10, 
        decimal_places=2, 
        required=True, 
        label='Cupon',
        widget=forms.NumberInput(attrs={'type': 'number', 'step': '0.01', 'inputmode': 'decimal'})
    )
    profature = forms.DecimalField(
        max_digits=10, 
        decimal_places=2, 
        required=True, 
        label='Profature',
        widget=forms.NumberInput(attrs={'type': 'number', 'step': '0.01', 'inputmode': 'decimal'})
    )
    ambasada = forms.DecimalField(
        max_digits=10, 
        decimal_places=2, 
        required=True, 
        label='Ambasada',
        widget=forms.NumberInput(attrs={'type': 'number', 'step': '0.01', 'inputmode': 'decimal'})
    )
    
    class Meta:
        model = ReceiptImage
        fields = ['image', 'cash', 'card', 'cupon', 'profature', 'ambasada']
