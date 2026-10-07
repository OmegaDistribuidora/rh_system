from django.contrib import admin
from django.urls import path
from django.shortcuts import redirect

from rh.views import admin_login_bridge, sso_entry, sso_exchange

urlpatterns = [
    path("", lambda request: redirect("/admin/login/")),

    path("sso/", sso_entry, name="sso_entry"),
    path("sso/exchange/", sso_exchange, name="sso_exchange"),

    path("admin/login/", admin_login_bridge, name="admin_login_bridge"),
    path("admin/", admin.site.urls),
]
