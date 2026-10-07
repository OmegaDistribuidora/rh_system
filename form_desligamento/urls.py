from django.contrib import admin
from django.urls import path
from django.shortcuts import redirect

from rh.views import sso_entry, sso_exchange

urlpatterns = [
    path("", lambda request: redirect("/admin/login/")),

    path("sso/", sso_entry, name="sso_entry"),
    path("sso/exchange/", sso_exchange, name="sso_exchange"),

    path("admin/", admin.site.urls),
]
