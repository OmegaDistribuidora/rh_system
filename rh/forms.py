from django import forms
from .models import Distrato, Admissao


class DistratoForm(forms.ModelForm):

    class Meta:
        model = Distrato
        fields = "__all__"
        widgets = {
            "banco": forms.TextInput(attrs={
                "placeholder": "Ex: Caixa Econômica",
                "maxlength": "50"
            }),
            "agencia": forms.TextInput(attrs={
                "placeholder": "Somente números",
                "maxlength": "10"
            }),
            "operacao": forms.TextInput(attrs={
                "placeholder": "Ex: 001",
                "maxlength": "5"
            }),
            "conta_corrente": forms.TextInput(attrs={
                "placeholder": "Conta sem dígito",
                "maxlength": "15"
            }),
            "titular": forms.TextInput(attrs={
                "placeholder": "Nome completo do titular",
                "maxlength": "100"
            }),
            "telefone": forms.TextInput(attrs={
                "placeholder": "(99) 99999-9999",
                "maxlength": "15"
            }),
        }

    def clean_agencia(self):
        agencia = self.cleaned_data.get("agencia")
        if agencia and not agencia.isdigit():
            raise forms.ValidationError("A agência deve conter apenas números.")
        return agencia

    def clean_conta_corrente(self):
        conta = self.cleaned_data.get("conta_corrente")
        if conta and not conta.isdigit():
            raise forms.ValidationError("A conta corrente deve conter apenas números.")
        return conta


class AdmissaoForm(forms.ModelForm):

    class Meta:
        model = Admissao
        fields = "__all__"
        widgets = {
            "cpf": forms.TextInput(attrs={
                "placeholder": "000.000.000-00",
                "maxlength": "14"
            }),
            "fone": forms.TextInput(attrs={
                "placeholder": "(99) 99999-9999",
                "maxlength": "15"
            }),
            "email": forms.EmailInput(attrs={
                "placeholder": "email@exemplo.com"
            }),
            "cep": forms.TextInput(attrs={
                "placeholder": "00000-000",
                "maxlength": "9"
            }),
        }

    def clean_cpf(self):
        cpf = self.cleaned_data.get("cpf")
        if cpf:
            numeros = [d for d in cpf if d.isdigit()]
            if len(numeros) != 11:
                raise forms.ValidationError("O CPF deve ter 11 dígitos numéricos.")
        return cpf

    def clean_fone(self):
        fone = self.cleaned_data.get("fone")
        if fone and not any(ch.isdigit() for ch in fone):
            raise forms.ValidationError("O telefone deve conter números.")
        return fone
