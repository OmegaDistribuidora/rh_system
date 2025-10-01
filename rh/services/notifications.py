import logging
from django.core.mail import send_mail
from django.conf import settings

logger = logging.getLogger(__name__)


def notificar_admissao(obj, usuario):
    """
    Envia e-mail notificando nova admissão pelo SendGrid.
    """
    assunto = "📥 Nova admissão registrada"
    corpo = (
        f"Uma nova admissão foi registrada:\n\n"
        f"Nome: {obj.nome}\n"
        f"Código RCA: {getattr(obj, 'codigo', '-')}\n"
        f"Data de Admissão: "
        f"{obj.data_admissao.strftime('%d/%m/%Y') if getattr(obj, 'data_admissao', None) else '-'}\n"
        f"Cargo: {getattr(obj, 'cargo', '-')}\n"
        f"Supervisor Responsável: {getattr(obj, 'supervisor_responsavel', '-')}\n"
        f"Registrado por: {usuario.get_username()}"
    )

    try:
        enviados = send_mail(
            subject=assunto,
            message=corpo,
            from_email=settings.DEFAULT_FROM_EMAIL,
            recipient_list=settings.EMAIL_DESTINATARIOS,
            fail_silently=False,
        )
        if enviados > 0:
            logger.info("E-mail de admissão enviado com sucesso para %s", settings.EMAIL_DESTINATARIOS)
        else:
            logger.warning("Nenhum e-mail de admissão foi enviado.")
    except Exception as e:
        logger.error("Erro ao enviar e-mail de admissão: %s", e, exc_info=True)
        raise


def notificar_desligamento(obj, usuario):
    """
    Envia e-mail notificando novo desligamento pelo SendGrid.
    """
    assunto = "📤 Novo desligamento registrado"
    corpo = (
        f"Uma nova saída foi registrada:\n\n"
        f"Nome: {obj.nome}\n"
        f"Código RCA: {getattr(obj, 'codigo', '-')}\n"
        f"Data de Desligamento: "
        f"{obj.data_desligamento.strftime('%d/%m/%Y') if getattr(obj, 'data_desligamento', None) else '-'}\n"
        f"Cargo: {getattr(obj, 'cargo', '-')}\n"
        f"Supervisor Responsável: {getattr(obj, 'supervisor_responsavel', '-')}\n"
        f"Registrado por: {usuario.get_username()}"
    )

    try:
        enviados = send_mail(
            subject=assunto,
            message=corpo,
            from_email=settings.DEFAULT_FROM_EMAIL,
            recipient_list=settings.EMAIL_DESTINATARIOS,
            fail_silently=False,
        )
        if enviados > 0:
            logger.info("E-mail de desligamento enviado com sucesso para %s", settings.EMAIL_DESTINATARIOS)
        else:
            logger.warning("Nenhum e-mail de desligamento foi enviado.")
    except Exception as e:
        logger.error("Erro ao enviar e-mail de desligamento: %s", e, exc_info=True)
        raise
