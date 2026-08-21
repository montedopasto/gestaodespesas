# Configuração do envio de emails

A aplicação envia emails através do Microsoft Graph em nome do utilizador autenticado.

## Permissão necessária no Microsoft Entra

Na aplicação registada com o ID `81a4b1c0-13eb-4c3e-bb82-283fa7d52334`:

1. Abrir **Microsoft Entra admin center**.
2. Entrar em **Aplicações > Registos de aplicações**.
3. Abrir a aplicação da Gestão de Despesas.
4. Entrar em **Permissões de API > Adicionar uma permissão**.
5. Escolher **Microsoft Graph > Permissões delegadas**.
6. Adicionar `Mail.Send` e `Mail.Send.Shared`.
7. Conceder consentimento de administrador, caso a política da organização o exija.

Depois da configuração, os utilizadores devem terminar sessão na aplicação e entrar novamente para autorizar a nova permissão.

## Emails enviados

- Ao submeter uma nota: email para `Aprovador1Email` e, quando definido, `Aprovador2Email`.
- Ao aprovar ou recusar: email para `CriadoPorEmail`.
- Nas recusas, o email inclui a justificação.

Os emails são enviados pela caixa partilhada `gestaodespesas@montedopasto.pt`, apresentada como **App Gestão de Despesas**.

Cada utilizador que possa submeter, aprovar ou recusar notas tem de possuir a permissão Exchange **Enviar como** sobre essa caixa partilhada. Esta permissão é independente das permissões Microsoft Graph.

## Power Automate anterior

Depois de validar este sistema, desativar o fluxo antigo que envia estes mesmos emails para evitar mensagens duplicadas.
