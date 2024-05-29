from django import forms


class PCForm(forms.Form):
    postcode = forms.CharField(label="Postcode", max_length=8)
