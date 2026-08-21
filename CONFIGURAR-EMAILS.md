# Configuração do envio de emails

A aplicação envia emails através do Microsoft Graph em nome do utilizador autenticado.

## Permissão necessária no Microsoft Entra

Na aplicação registada com o ID `81a4b1c0-13eb-4c3e-bb82-283fa7d52334`:

1. Abrir **Microsoft Entra admin center**.
2. Entrar em **Aplicações > Registos de aplicações**.
3. Abrir a aplicação da Gestão de Despesas.
4. Entrar em **Permissões de API > Adicionar uma permissão**.
5. Escolher **Microsoft Graph > Permissões delegadas**.
6. Adicionar `Mail.Send`.
7. Conceder consentimento de administrador, caso a política da organização o exija.

Depois da configuração, os utilizadores devem terminar sessão na aplicação e entrar novamente para autorizar a nova permissão.

## Emails enviados

- Ao submeter uma nota: email para `Aprovador1Email` e, quando definido, `Aprovador2Email`.
- Ao aprovar ou recusar: email para `CriadoPorEmail`.
- Nas recusas, o email inclui a justificação.

Os emails são enviados pela caixa de correio do utilizador que executa a ação e ficam nos respetivos Itens Enviados.

## Power Automate anterior

Depois de validar este sistema, desativar o fluxo antigo que envia estes mesmos emails para evitar mensagens duplicadas.
