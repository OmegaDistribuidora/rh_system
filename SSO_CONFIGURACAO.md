# SSO do Ecossistema Omega

O RH recebe um JWT de uso único emitido pelo Ecossistema Omega e, depois de
validá-lo, abre uma sessão Django normal. Usuários, grupos e permissões continuam
sendo administrados no RH.

O endpoint de troca não usa a validação CSRF baseada em `Origin`, porque alguns
navegadores enviam `Origin: null` no redirecionamento delegado. Ele aceita somente
JWT assinado para este sistema, com expiração curta e identificador de uso único.
O hash desse identificador é registrado na tabela nativa `django_session`,
impedindo reutilização inclusive entre réplicas. As demais rotas continuam
protegidas pelo middleware CSRF do Django.

## 1. Primeiro deploy do RH

Faça o primeiro deploy com `ECOSYSTEM_SSO_ENABLED=False`. O arquivo
`railway.json` executa automaticamente antes de cada versão:

```text
python manage.py migrate
```

O login local permanece funcionando durante toda a implantação.

## 2. Variáveis do RH

Configure no serviço do RH:

```text
ECOSYSTEM_SSO_ENABLED=True
ECOSYSTEM_SSO_ISSUER=ecosistema-omega
ECOSYSTEM_SSO_AUDIENCE=rh_system
ECOSYSTEM_SSO_SHARED_SECRET=SEGREDO_ALEATORIO_COM_32_OU_MAIS_CARACTERES
```

`ECOSYSTEM_SSO_ADMIN_USERS` é opcional. Ele recebe logins do Ecossistema,
separados por vírgula, autorizados a abrir contas superusuárias do RH. Um
superusuário também pode entrar quando possui o mesmo login nos dois sistemas e
é administrador no Ecossistema.

## 3. Configuração no Ecossistema

No cadastro do sistema Forms Omega/RH, use:

```text
URL=https://forms-omega.up.railway.app/sso/
SSO habilitado=sim
Chave SSO=rh_system
```

O endereço antigo `/admin/login/` também reconhece o token e encaminha para o
fluxo SSO, inclusive quando já existe outra sessão ativa. Usar `/sso/`
diretamente deixa a configuração mais explícita.

No serviço do Ecossistema, configure:

```text
SSO_SECRET_RH_SYSTEM=O_MESMO_SEGREDO_CONFIGURADO_NO_RH
SSO_AUDIENCE_RH_SYSTEM=rh_system
SSO_TTL_RH_SYSTEM=45
```

Para cada pessoa liberada no módulo, o login externo deve ser exatamente o
`username` de um usuário Django ativo e com acesso de staff ao RH.

Use um segredo exclusivo para este sistema. Não reutilize `SECRET_KEY`, senha de
banco, segredo de sessão ou segredos SSO de outros módulos.
