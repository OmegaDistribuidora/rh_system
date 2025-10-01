from django.contrib.auth.models import User
from rh.models import Hierarquia

GRUPO_RH = "RH"
GRUPO_COORDENADOR = "COORDENADORES"
GRUPO_SUPERVISOR = "COLABORADORES"


def users_visiveis_para(user):

    if user.is_superuser or user.groups.filter(name=GRUPO_RH).exists():
        return User.objects.all()

    elif user.groups.filter(name=GRUPO_COORDENADOR).exists():
        supervisores = Hierarquia.objects.filter(coordenador=user).values_list("supervisor_id", flat=True)
        return User.objects.filter(id__in=list(supervisores) + [user.id])

    elif user.groups.filter(name=GRUPO_SUPERVISOR).exists():
        return User.objects.filter(id=user.id)

    return User.objects.filter(id=user.id)
