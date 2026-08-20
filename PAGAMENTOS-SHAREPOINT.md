# Configuração do módulo Pagamentos

O módulo não altera o campo `Estado` nem o fluxo de aprovação existente. Apenas apresenta itens da lista `NotasDespesa` cujo `Estado` seja `Aprovado`.

## Colunas necessárias em `NotasDespesa`

| Nome interno | Tipo | Configuração |
| --- | --- | --- |
| `EstadoPagamento` | Escolha | Valores `Por pagar` e `Pago`; predefinição `Por pagar` |
| `Pago` | Sim/Não | Predefinição `Não` |
| `PagoPorNome` | Uma linha de texto | Sem valor predefinido |
| `PagoPorEmail` | Uma linha de texto | Sem valor predefinido |
| `DataPagamento` | Data e hora | Incluir hora |
| `NotasPagamento` | Várias linhas de texto | Texto simples |

Os itens antigos que ainda não tenham `EstadoPagamento` são apresentados como `Por pagar`.

## Power Automate

O botão de confirmação atualiza os seis campos numa única operação. O fluxo de notificação pode usar o acionador SharePoint **Quando um item é criado ou modificado** na lista `NotasDespesa` e avançar apenas quando:

- `EstadoPagamento` é igual a `Pago`;
- `Pago` é igual a `Sim`.

Para evitar emails duplicados em futuras edições do mesmo item, o fluxo deve manter a condição de mudança de valor (por exemplo, através de **Obter alterações para um item ou ficheiro**) e enviar o email apenas quando `EstadoPagamento` tiver acabado de mudar para `Pago`.

Campos disponíveis para o email: `CriadoPorNome`, `CriadoPorEmail`, `TipoDocumento`, `TotalRecebido`, `DataPagamento`, `PagoPorNome` e `NotasPagamento`.
