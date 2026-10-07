from datetime import date

from django.db import IntegrityError, transaction
from django.test import TestCase

from .admin import AdmissaoForm
from .models import Admissao


class AdmissaoCodigoTests(TestCase):
    def test_codigo_nao_e_obrigatorio_no_formulario(self):
        form = AdmissaoForm()

        self.assertFalse(form.fields["codigo"].required)

    def test_permite_multiplas_admissoes_sem_codigo_na_mesma_data(self):
        data_admissao = date(2026, 10, 7)

        primeira = Admissao.objects.create(
            nome="Primeiro vendedor",
            codigo="",
            data_admissao=data_admissao,
        )
        segunda = Admissao.objects.create(
            nome="Segundo vendedor",
            codigo="",
            data_admissao=data_admissao,
        )

        self.assertNotEqual(primeira.pk, segunda.pk)

    def test_rejeita_codigo_repetido_na_mesma_data(self):
        data_admissao = date(2026, 10, 7)
        Admissao.objects.create(
            nome="Primeiro vendedor",
            codigo="123",
            data_admissao=data_admissao,
        )

        with self.assertRaises(IntegrityError), transaction.atomic():
            Admissao.objects.create(
                nome="Segundo vendedor",
                codigo="123",
                data_admissao=data_admissao,
            )
