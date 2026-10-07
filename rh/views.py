import logging

from django.conf import settings
from django.shortcuts import redirect, render
from django.urls import reverse
from django.utils.http import url_has_allowed_host_and_scheme
from django.views.decorators.cache import never_cache
from django.views.decorators.csrf import csrf_exempt
from django.views.decorators.debug import sensitive_post_parameters
from django.views.decorators.http import require_GET, require_POST

from .services.sso import SsoAuthenticationError, autenticar_via_sso


logger = logging.getLogger(__name__)


def _render_sso(request, context, status=200):
    response = render(request, "sso/entry.html", context, status=status)
    response["Cache-Control"] = "no-store, max-age=0"
    response["Pragma"] = "no-cache"
    response["Referrer-Policy"] = "no-referrer"
    response["Content-Security-Policy"] = (
        "default-src 'none'; style-src 'unsafe-inline'; script-src 'unsafe-inline'; "
        "form-action 'self'; base-uri 'none'; frame-ancestors 'none'"
    )
    return response


def _safe_next_url(request):
    candidate = request.POST.get("next") or request.GET.get("next") or reverse("admin:index")
    if url_has_allowed_host_and_scheme(
        candidate,
        allowed_hosts={request.get_host()},
        require_https=request.is_secure(),
    ):
        return candidate
    return reverse("admin:index")


@never_cache
@require_GET
def sso_entry(request):
    status = 200 if settings.ECOSYSTEM_SSO_ENABLED else 404
    return _render_sso(
        request,
        {
            "sso_enabled": settings.ECOSYSTEM_SSO_ENABLED,
            "next_url": _safe_next_url(request),
        },
        status=status,
    )


@sensitive_post_parameters("token")
@csrf_exempt
@never_cache
@require_POST
def sso_exchange(request):
    next_url = _safe_next_url(request)
    if not settings.ECOSYSTEM_SSO_ENABLED:
        return _render_sso(
            request,
            {"sso_enabled": False, "next_url": next_url},
            status=404,
        )
    try:
        autenticar_via_sso(request, request.POST.get("token", ""))
    except SsoAuthenticationError as exc:
        logger.warning("Login SSO recusado: %s", exc)
        return _render_sso(
            request,
            {
                "sso_enabled": settings.ECOSYSTEM_SSO_ENABLED,
                "exchange_attempted": True,
                "error": "Não foi possível validar seu acesso. Abra o sistema novamente pelo Ecossistema Omega.",
                "next_url": next_url,
            },
            status=401,
        )

    return redirect(next_url)
