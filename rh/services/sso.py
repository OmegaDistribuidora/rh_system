import hashlib
import logging
from datetime import datetime, timezone as datetime_timezone

import jwt
from django.conf import settings
from django.contrib.auth import login
from django.contrib.auth.models import User
from django.db import IntegrityError, transaction
from django.utils import timezone
from django.views.decorators.debug import sensitive_variables

from rh.models import SsoTokenConsumido


logger = logging.getLogger(__name__)


class SsoAuthenticationError(Exception):
    pass


@sensitive_variables("token")
def _decode_token(token):
    if not settings.ECOSYSTEM_SSO_ENABLED:
        raise SsoAuthenticationError("SSO desabilitado")
    if not token or len(token) > 8192:
        raise SsoAuthenticationError("Token ausente ou grande demais")

    try:
        payload = jwt.decode(
            token,
            settings.ECOSYSTEM_SSO_SHARED_SECRET,
            algorithms=["HS256"],
            issuer=settings.ECOSYSTEM_SSO_ISSUER,
            audience=settings.ECOSYSTEM_SSO_AUDIENCE,
            leeway=settings.ECOSYSTEM_SSO_CLOCK_SKEW_SECONDS,
            options={"require": ["exp", "iat", "jti", "sub", "targetLogin"]},
        )
    except jwt.PyJWTError as exc:
        raise SsoAuthenticationError("Token inválido ou expirado") from exc

    issued_at = payload.get("iat")
    expires_at = payload.get("exp")
    if not isinstance(issued_at, (int, float)) or not isinstance(expires_at, (int, float)):
        raise SsoAuthenticationError("Datas do token inválidas")
    if expires_at <= issued_at:
        raise SsoAuthenticationError("Período do token inválido")
    if expires_at - issued_at > settings.ECOSYSTEM_SSO_MAX_TOKEN_LIFETIME_SECONDS:
        raise SsoAuthenticationError("Tempo de vida do token excedido")

    jti = str(payload.get("jti") or "").strip()
    target_login = str(payload.get("targetLogin") or "").strip()
    ecosystem_username = str(payload.get("ecosystemUsername") or "").strip()
    if not jti or not target_login or len(target_login) > 150:
        raise SsoAuthenticationError("Token incompleto")

    return payload, jti, target_login, ecosystem_username, expires_at


@sensitive_variables("token")
def autenticar_via_sso(request, token):
    payload, jti, target_login, ecosystem_username, expires_at = _decode_token(token)

    try:
        user = User.objects.get(username=target_login)
    except User.DoesNotExist as exc:
        raise SsoAuthenticationError("Usuário de destino não encontrado") from exc

    if not user.is_active or not user.is_staff:
        raise SsoAuthenticationError("Usuário de destino sem acesso ao sistema")

    if user.is_superuser:
        ecosystem_is_admin = payload.get("ecosystemIsAdmin") is True
        ecosystem_login = ecosystem_username.lower()
        same_login = ecosystem_login == user.username.lower()
        allowlisted = ecosystem_login in settings.ECOSYSTEM_SSO_ADMIN_USERS
        if not ecosystem_is_admin or not (same_login or allowlisted):
            raise SsoAuthenticationError("SSO sem autorização administrativa")

    jti_hash = hashlib.sha256(jti.encode("utf-8")).hexdigest()
    expiration = datetime.fromtimestamp(expires_at, tz=datetime_timezone.utc)

    try:
        with transaction.atomic():
            SsoTokenConsumido.objects.filter(expira_em__lt=timezone.now()).delete()
            SsoTokenConsumido.objects.create(
                jti_hash=jti_hash,
                expira_em=expiration,
                usuario=user,
                usuario_ecossistema=ecosystem_username[:150],
                login_destino=target_login,
            )
    except IntegrityError as exc:
        raise SsoAuthenticationError("Token SSO já utilizado") from exc

    login(request, user, backend="django.contrib.auth.backends.ModelBackend")
    logger.info(
        "Login SSO concluído: usuario=%s usuario_ecossistema=%s",
        user.username,
        ecosystem_username or "-",
    )
    return user
