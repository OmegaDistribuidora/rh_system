import time
import uuid

import jwt
from django.contrib.auth.models import User
from django.contrib.sessions.models import Session
from django.test import Client, TestCase, override_settings
from django.urls import reverse


SSO_SECRET = "segredo-de-teste-com-mais-de-trinta-e-dois-caracteres"


@override_settings(
    ECOSYSTEM_SSO_ENABLED=True,
    ECOSYSTEM_SSO_ISSUER="ecosistema-omega",
    ECOSYSTEM_SSO_AUDIENCE="rh_system",
    ECOSYSTEM_SSO_SHARED_SECRET=SSO_SECRET,
    ECOSYSTEM_SSO_ADMIN_USERS=set(),
    ECOSYSTEM_SSO_MAX_TOKEN_LIFETIME_SECONDS=300,
    ECOSYSTEM_SSO_CLOCK_SKEW_SECONDS=5,
)
class SsoLoginTests(TestCase):
    def setUp(self):
        self.user = User.objects.create_user(
            username="supervisor.rh",
            password="senha-local-de-contingencia",
            is_staff=True,
        )

    def token(self, **overrides):
        now = int(time.time())
        payload = {
            "iss": "ecosistema-omega",
            "aud": "rh_system",
            "sub": "1",
            "iat": now,
            "exp": now + 45,
            "jti": str(uuid.uuid4()),
            "ecosystemUsername": "supervisor.rh",
            "ecosystemIsAdmin": False,
            "targetLogin": self.user.username,
        }
        payload.update(overrides)
        return jwt.encode(payload, SSO_SECRET, algorithm="HS256")

    def test_pagina_de_entrada_contem_post_protegido_por_csrf(self):
        response = self.client.get(reverse("sso_entry"))

        self.assertEqual(response.status_code, 200)
        self.assertContains(response, "csrfmiddlewaretoken")
        self.assertContains(response, "window.location.hash")

    def test_login_admin_encaminha_token_recebido_no_fragmento(self):
        response = self.client.get(reverse("admin:login"))

        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, "admin/sso_login.html")
        self.assertContains(response, "window.location.replace")

    def test_login_admin_processa_novo_token_mesmo_com_sessao_ativa(self):
        self.client.force_login(self.user)

        response = self.client.get(reverse("admin:login"))

        self.assertEqual(response.status_code, 200)
        self.assertTemplateUsed(response, "sso/entry.html")
        self.assertContains(response, "window.location.hash")

    def test_troca_aceita_origin_nulo_pois_o_jwt_e_a_credencial(self):
        client = Client(enforce_csrf_checks=True)

        response = client.post(
            reverse("sso_exchange"),
            {"token": self.token()},
            HTTP_ORIGIN="null",
        )

        self.assertEqual(response.status_code, 302)
        self.assertEqual(int(client.session["_auth_user_id"]), self.user.pk)
        self.assertEqual(Session.objects.filter(session_key__startswith="sso").count(), 1)

    def test_fluxo_com_csrf_valido_abre_sessao(self):
        client = Client(enforce_csrf_checks=True)
        entry_response = client.get(reverse("sso_entry"))
        csrf_token = client.cookies["csrftoken"].value

        response = client.post(
            reverse("sso_exchange"),
            {"token": self.token()},
            HTTP_X_CSRFTOKEN=csrf_token,
        )

        self.assertEqual(entry_response.status_code, 200)
        self.assertEqual(response.status_code, 302)
        self.assertEqual(int(client.session["_auth_user_id"]), self.user.pk)

    def test_token_valido_abre_sessao_django(self):
        response = self.client.post(
            reverse("sso_exchange"),
            {"token": self.token(), "next": reverse("admin:index")},
        )

        self.assertRedirects(response, reverse("admin:index"), fetch_redirect_response=False)
        self.assertEqual(int(self.client.session["_auth_user_id"]), self.user.pk)
        self.assertEqual(Session.objects.filter(session_key__startswith="sso").count(), 1)

    def test_token_nao_pode_ser_reutilizado(self):
        token = self.token()
        primeira = self.client.post(reverse("sso_exchange"), {"token": token})
        self.client.logout()
        segunda = self.client.post(reverse("sso_exchange"), {"token": token})

        self.assertEqual(primeira.status_code, 302)
        self.assertEqual(segunda.status_code, 401)
        self.assertEqual(Session.objects.filter(session_key__startswith="sso").count(), 1)

    def test_novo_sso_substitui_usuario_da_sessao_anterior(self):
        segundo_usuario = User.objects.create_user(
            username="arleilson",
            password="outra-senha-local",
            is_staff=True,
        )
        self.client.force_login(self.user)

        response = self.client.post(
            reverse("sso_exchange"),
            {
                "token": self.token(
                    targetLogin=segundo_usuario.username,
                    ecosystemUsername=segundo_usuario.username,
                )
            },
        )

        self.assertEqual(response.status_code, 302)
        self.assertEqual(
            int(self.client.session["_auth_user_id"]),
            segundo_usuario.pk,
        )

    def test_rejeita_audience_incorreta(self):
        response = self.client.post(
            reverse("sso_exchange"),
            {"token": self.token(aud="outro-sistema")},
        )

        self.assertEqual(response.status_code, 401)
        self.assertNotIn("_auth_user_id", self.client.session)

    def test_rejeita_redirecionamento_para_dominio_externo(self):
        response = self.client.post(
            reverse("sso_exchange"),
            {"token": self.token(), "next": "https://exemplo-invalido.test/roubo"},
        )

        self.assertRedirects(response, reverse("admin:index"), fetch_redirect_response=False)

    def test_superusuario_exige_administrador_do_ecossistema(self):
        self.user.is_superuser = True
        self.user.save(update_fields=["is_superuser"])

        response = self.client.post(reverse("sso_exchange"), {"token": self.token()})

        self.assertEqual(response.status_code, 401)
        self.assertNotIn("_auth_user_id", self.client.session)

    def test_superusuario_equivalente_e_administrador_pode_entrar(self):
        self.user.is_superuser = True
        self.user.save(update_fields=["is_superuser"])

        response = self.client.post(
            reverse("sso_exchange"),
            {"token": self.token(ecosystemIsAdmin=True)},
        )

        self.assertEqual(response.status_code, 302)
        self.assertEqual(int(self.client.session["_auth_user_id"]), self.user.pk)

    def test_administrador_do_ecossistema_nao_promove_usuario_comum(self):
        response = self.client.post(
            reverse("sso_exchange"),
            {"token": self.token(ecosystemIsAdmin=True)},
        )

        self.user.refresh_from_db()
        self.assertEqual(response.status_code, 302)
        self.assertFalse(self.user.is_superuser)

    def test_superusuario_rejeita_administrador_com_login_diferente(self):
        self.user.is_superuser = True
        self.user.save(update_fields=["is_superuser"])

        response = self.client.post(
            reverse("sso_exchange"),
            {
                "token": self.token(
                    ecosystemIsAdmin=True,
                    ecosystemUsername="outro-administrador",
                )
            },
        )

        self.assertEqual(response.status_code, 401)
        self.assertNotIn("_auth_user_id", self.client.session)

    def test_usuario_sem_acesso_ao_admin_e_rejeitado(self):
        self.user.is_staff = False
        self.user.save(update_fields=["is_staff"])

        response = self.client.post(reverse("sso_exchange"), {"token": self.token()})

        self.assertEqual(response.status_code, 401)
        self.assertNotIn("_auth_user_id", self.client.session)


class SsoDesabilitadoTests(TestCase):
    @override_settings(ECOSYSTEM_SSO_ENABLED=False)
    def test_endpoint_de_entrada_fica_indisponivel(self):
        response = self.client.get(reverse("sso_entry"))

        self.assertEqual(response.status_code, 404)

    @override_settings(ECOSYSTEM_SSO_ENABLED=False)
    def test_endpoint_de_troca_fica_indisponivel(self):
        response = self.client.post(reverse("sso_exchange"), {"token": "qualquer"})

        self.assertEqual(response.status_code, 404)
