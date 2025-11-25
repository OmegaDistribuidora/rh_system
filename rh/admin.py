from django import forms
from django.contrib import admin
from django.urls import path, reverse
from django.core.exceptions import PermissionDenied, ValidationError
from django.utils.html import format_html
from django.db.models import Count
from django.http import HttpResponse
from django.utils.translation import gettext_lazy as _

from .forms import DistratoForm
from .models import Desligamento, Admissao, Distrato, Hierarquia, RelatorioRH, Caju

# from .services.notifications import notificar_admissao  # envio desativado

from .services.excel import (
    exportar_desligamento_excel,
    exportar_admissao_excel,
    exportar_distrato_excel,
)
from .services.permission import users_visiveis_para
from .services.relatorio import gerar_relatorio_pdf


#-------------------------------------------------------------------------------------------------------------------
# Formulários customizados

class DesligamentoForm(forms.ModelForm):
    class Meta:
        model = Desligamento
        fields = "__all__"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        obrigatorios = ["motivo"]
        for campo in obrigatorios:
            self.fields[campo].required = True


class AdmissaoForm(forms.ModelForm):
    class Meta:
        model = Admissao
        fields = "__all__"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        obrigatorios = [
            "nascimento", "naturalidade", "uf",
            "endereco", "bairro", "cidade", "estado", "cep",
            "fone", "email", "rg", "orgao_exp", "emissao", "cpf",
            "banco", "agencia", "conta", "operacao",
            "data_admissao", "cargo", "supervisor_responsavel"
        ]
        for campo in obrigatorios:
            self.fields[campo].required = True

    def clean(self):
        cleaned_data = super().clean()
        cpf = cleaned_data.get("cpf")

        if cpf and Admissao.objects.filter(cpf=cpf).exclude(pk=self.instance.pk).exists():
            raise ValidationError(f"Já existe uma admissão registrada com o CPF {cpf}.")
        return cleaned_data


class CajuForm(forms.ModelForm):
    class Meta:
        model = Caju
        fields = "__all__"

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        obrigatorios = ["cpf", "nome_completo", "email", "telefone"]
        for campo in obrigatorios:
            self.fields[campo].required = True

    def clean(self):
        cleaned_data = super().clean()
        cpf = cleaned_data.get("cpf")
        email = cleaned_data.get("email")

        if cpf and Caju.objects.filter(cpf=cpf).exclude(pk=self.instance.pk).exists():
            raise ValidationError(f"Já existe um registro de Caju com o CPF {cpf}.")
        if email and Caju.objects.filter(email=email).exclude(pk=self.instance.pk).exists():
            raise ValidationError(f"Já existe um registro de Caju com o e-mail {email}.")
        return cleaned_data

#-------------------------------------------------------------------------------------------------------------------
#admin: Desligamento

@admin.register(Desligamento)
class DesligamentoAdmin(admin.ModelAdmin):
    form = DesligamentoForm
    list_display = (
        "nome", "codigo", "supervisor", "demissao",
        "area_atuacao", "criado_por", "status", "qtd_desligamentos_colaborador"
    )
    search_fields = ("nome", "codigo", "area_atuacao")
    list_filter = ("status", "area_atuacao", "demissao", "criado_por")
    list_editable = ("status",)

    fieldsets = (
        ('📌 Dados do Colaborador', {
            'fields': ('codigo', 'nome', 'contato', 'admissao', 'demissao', 'area_atuacao')
        }),
        ('📄 Motivo do desligamento', {
            'fields': ('motivo',)
        }),
        ('📦 Itens a devolver', {
            'fields': (
                'fardamento', 'chip_voz', 'chip_dados', 'tablet',
                'carregador_tablet', 'fone_tablet', 'catalogo',
                'bloco_pedido', 'carta_pedido_demissao', 'relatorio_inadimplencia'
            )
        }),
        ('🔎 Perguntas extras', {
            'fields': ('substituto', 'telemarketing', 'nova_contratacao')
        }),
        ('📊 Status', {
            'fields': ('status',)
        }),
    )

    def get_changelist_formset(self, request, **kwargs):
        formset = super().get_changelist_formset(request, **kwargs)
        if not (request.user.is_superuser or request.user.groups.filter(name="RH").exists()):
            formset.form.base_fields["status"].disabled = True
        return formset

    def get_readonly_fields(self, request, obj=None):
        readonly = list(super().get_readonly_fields(request, obj))
        if not (request.user.is_superuser or request.user.groups.filter(name="RH").exists()):
            readonly.append("status")
        return readonly

    def save_model(self, request, obj, form, change):
        is_new = obj.pk is None
        if is_new:
            obj.criado_por = request.user
            if Desligamento.objects.filter(codigo=obj.codigo, demissao=obj.demissao).exists():
                raise ValidationError("Já existe um desligamento registrado para este colaborador nessa data.")
        super().save_model(request, obj, form, change)

        # Envio desativado
        # if is_new:
        #     notificar_desligamento(obj, request.user)

    def get_queryset(self, request):
        qs = super().get_queryset(request)
        users_visiveis = users_visiveis_para(request.user)
        return qs.filter(criado_por__in=users_visiveis).annotate(
            total_desligamentos=Count("criado_por")
        )

    def qtd_desligamentos_colaborador(self, obj):
        return obj.total_desligamentos
    qtd_desligamentos_colaborador.short_description = "Qtd desligamentos"

    def has_export_permission(self, request):
        return request.user.is_superuser or request.user.groups.filter(name="RH").exists()

    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path(
                '<int:desligamento_id>/exportar_excel/',
                self.admin_site.admin_view(self.exportar_excel),
                name="rh_desligamento_exportar_excel_individual",
            ),
        ]
        return custom_urls + urls

    def change_view(self, request, object_id, form_url='', extra_context=None):
        extra_context = extra_context or {}
        export_url = reverse("admin:rh_desligamento_exportar_excel_individual", args=[object_id])
        extra_context['extra_buttons'] = format_html(
            f'<a class="button" style="margin-left:10px;" href="{export_url}">📤 Exportar Excel</a>'
        )
        return super().change_view(request, object_id, form_url, extra_context=extra_context)

    def exportar_excel(self, request, desligamento_id):
        if not self.has_export_permission(request):
            raise PermissionDenied("Você não tem permissão para exportar este registro.")
        desligamento = Desligamento.objects.get(id=desligamento_id)
        return exportar_desligamento_excel(desligamento)

    def has_view_permission(self, request, obj=None):
        return True

    def has_module_permission(self, request):
        return self.has_view_permission(request)

#-------------------------------------------------------------------------------------------------------------------
#admin: Admissao

@admin.register(Admissao)
class AdmissaoAdmin(admin.ModelAdmin):
    form = AdmissaoForm
    list_display = ("nome", "codigo", "supervisor", "data_admissao", "cargo", "criado_por", "status")
    search_fields = ("nome", "codigo", "cpf", "cargo", "supervisor_responsavel")
    list_filter = ("status", "cargo", "data_admissao", "criado_por")
    list_editable = ("status",)

    fieldsets = (
        ("📌 Dados Pessoais", {
            "fields": ("codigo", "nome", "nascimento", "naturalidade", "uf", "mae", "pai")
        }),
        ("🏠 Endereço", {
            "fields": ("endereco", "bairro", "cidade", "estado", "cep")
        }),
        ("📞 Contato", {
            "fields": ("fone", "email")
        }),
        ("📄 Documentos", {
            "fields": ("rg", "orgao_exp", "emissao", "cpf")
        }),
        ("🏦 Dados Bancários", {
            "fields": ("banco", "agencia", "conta", "operacao")
        }),
        ("💼 Condições de Admissão", {
            "fields": ("data_admissao", "cargo", "substituicao", "supervisor_responsavel", "coordenador")
        }),
        ("🔐 Conta Gov", {
            "fields": ("conta_gov", "senha_gov")
        }),
        ("📝 Observações", {
            "fields": ("observacoes",)
        }),
        ("📊 Status", {
            "fields": ("status",)
        }),
    )

    def get_changelist_formset(self, request, **kwargs):
        formset = super().get_changelist_formset(request, **kwargs)
        if not (request.user.is_superuser or request.user.groups.filter(name="RH").exists()):
            formset.form.base_fields["status"].disabled = True
        return formset

    def get_readonly_fields(self, request, obj=None):
        readonly = list(super().get_readonly_fields(request, obj))
        if not (request.user.is_superuser or request.user.groups.filter(name="RH").exists()):
            readonly.append("status")
        return readonly

    def save_model(self, request, obj, form, change):
        is_new = obj.pk is None
        if is_new:
            obj.criado_por = request.user
            if Admissao.objects.filter(codigo=obj.codigo, data_admissao=obj.data_admissao).exists():
                raise ValidationError("Já existe uma admissão registrada para este código nessa data.")
        super().save_model(request, obj, form, change)

        # Envio desativado
        # if is_new:
        #     notificar_admissao(obj, request.user)

    def get_queryset(self, request):
        qs = super().get_queryset(request)
        users_visiveis = users_visiveis_para(request.user)
        return qs.filter(criado_por__in=users_visiveis)

    def has_export_permission(self, request):
        return request.user.is_superuser or request.user.groups.filter(name="RH").exists()

    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path(
                '<int:admissao_id>/exportar_excel/',
                self.admin_site.admin_view(self.exportar_excel),
                name="rh_admissao_exportar_excel_individual",
            ),
        ]
        return custom_urls + urls

    def change_view(self, request, object_id, form_url='', extra_context=None):
        extra_context = extra_context or {}
        export_url = reverse("admin:rh_admissao_exportar_excel_individual", args=[object_id])
        extra_context['extra_buttons'] = format_html(
            f'<a class="button" style="margin-left:10px;" href="{export_url}">📤 Exportar Excel</a>'
        )
        return super().change_view(request, object_id, form_url, extra_context=extra_context)

    def exportar_excel(self, request, admissao_id):
        if not self.has_export_permission(request):
            raise PermissionDenied("Você não tem permissão para exportar este registro.")
        admissao = Admissao.objects.get(id=admissao_id)
        return exportar_admissao_excel(admissao)

    def has_view_permission(self, request, obj=None):
        return True

    def has_module_permission(self, request):
        return self.has_view_permission(request)

#-------------------------------------------------------------------------------------------------------------------
#admin: Distrato

@admin.register(Distrato)
class DistratoAdmin(admin.ModelAdmin):
    form = DistratoForm
    list_display = ("nome", "cpf", "data_admissao", "data_demissao",
                    "total_geral", "total_ultimos_3_meses", "criado_por", "status")
    search_fields = ("nome", "cpf", "rg")
    list_filter = ("status", "data_demissao", "criado_por")
    list_editable = ("status",)

    fieldsets = (
        ("📌 Dados do Representante", {
            "fields": ("nome", "cpf", "rg")
        }),
        ("📅 Datas", {
            "fields": ("data_admissao", "data_demissao")
        }),
        ("💰 Totais (somente campos amarelos)", {
            "fields": ("total_geral", "total_ultimos_3_meses")
        }),
        ("🏦 Dados Bancários", {
            "fields": ("banco", "agencia", "operacao", "conta_corrente", "titular", "telefone")
        }),
        ("📊 Status", {
            "fields": ("status",)
        }),
    )

    def get_changelist_formset(self, request, **kwargs):
        formset = super().get_changelist_formset(request, **kwargs)
        if not (request.user.is_superuser or request.user.groups.filter(name="RH").exists()):
            formset.form.base_fields["status"].disabled = True
        return formset

    def get_readonly_fields(self, request, obj=None):
        readonly = list(super().get_readonly_fields(request, obj))
        if not (request.user.is_superuser or request.user.groups.filter(name="RH").exists()):
            readonly.append("status")
        return readonly

    def save_model(self, request, obj, form, change):
        if not obj.pk:
            obj.criado_por = request.user
            if Distrato.objects.filter(cpf=obj.cpf, data_demissao=obj.data_demissao).exists():
                raise ValidationError("Já existe um distrato registrado para este CPF nessa data.")
        super().save_model(request, obj, form, change)

    def get_queryset(self, request):
        qs = super().get_queryset(request)
        users_visiveis = users_visiveis_para(request.user)
        return qs.filter(criado_por__in=users_visiveis)

    def has_export_permission(self, request):
        return request.user.is_superuser or request.user.groups.filter(name="RH").exists()

    def get_urls(self):
        urls = super().get_urls()
        custom_urls = [
            path(
                '<int:distrato_id>/exportar_excel/',
                self.admin_site.admin_view(self.exportar_excel),
                name="rh_distrato_exportar_excel_individual",
            ),
        ]
        return custom_urls + urls

    def change_view(self, request, object_id, form_url='', extra_context=None):
        extra_context = extra_context or {}
        export_url = reverse("admin:rh_distrato_exportar_excel_individual", args=[object_id])
        extra_context['extra_buttons'] = format_html(
            f'<a class="button" style="margin-left:10px;" href="{export_url}">📤 Exportar Distrato</a>'
        )
        return super().change_view(request, object_id, form_url, extra_context=extra_context)

    def exportar_excel(self, request, distrato_id):
        if not self.has_export_permission(request):
            raise PermissionDenied("Você não tem permissão para exportar este registro.")
        distrato = Distrato.objects.get(id=distrato_id)
        return exportar_distrato_excel(distrato)

    def has_view_permission(self, request, obj=None):
        return request.user.is_superuser or request.user.groups.filter(name="RH").exists()

    def has_module_permission(self, request):
        return self.has_view_permission(request)

#-------------------------------------------------------------------------------------------------------------------
#admin: Hierarquia

@admin.register(Hierarquia)
class HierarquiaAdmin(admin.ModelAdmin):
    list_display = ("coordenador", "supervisor")
    search_fields = ("coordenador__username", "supervisor__username")

    def has_view_permission(self, request, obj=None):
        return request.user.is_superuser or request.user.groups.filter(name="RH").exists()

    def has_change_permission(self, request, obj=None):
        return request.user.is_superuser or request.user.groups.filter(name="RH").exists()

    def has_add_permission(self, request):
        return request.user.is_superuser or request.user.groups.filter(name="RH").exists()

    def has_delete_permission(self, request, obj=None):
        return request.user.is_superuser or request.user.groups.filter(name="RH").exists()

    def has_module_permission(self, request):
        return self.has_view_permission(request)

#-------------------------------------------------------------------------------------------------------------------
#admin: Relatorio

@admin.register(RelatorioRH)
class RelatorioRHAdmin(admin.ModelAdmin):
    def has_add_permission(self, request):
        return False

    def has_change_permission(self, request, obj=None):
        return False

    def has_delete_permission(self, request, obj=None):
        return False

    def changelist_view(self, request, extra_context=None):
        return gerar_relatorio_pdf()

#-------------------------------------------------------------------------------------------------------------------
#admin: Caju

@admin.register(Caju)
class CajuAdmin(admin.ModelAdmin):
    form = CajuForm
    list_display = ("nome_completo", "cpf", "email", "telefone", "data_cadastro")
    search_fields = ("nome_completo", "cpf", "email", "telefone")
    list_filter = ("data_cadastro",)

    fieldsets = (
        ("📌 Dados do Caju", {
            "fields": ("cpf", "nome_completo", "email", "telefone")
        }),
        ("📊 Registro", {
            "fields": ("data_cadastro",)
        }),
    )

    readonly_fields = ("data_cadastro",)

    def get_queryset(self, request):
        qs = super().get_queryset(request)
        return qs
