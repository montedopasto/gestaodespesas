# Configurar a correção de valores

Na lista SharePoint **NotasDespesa**, criar estas colunas com exatamente estes nomes:

| Nome | Tipo | Valor predefinido |
|---|---|---|
| `EmCorrecao` | Sim/Não | Não |
| `MotivoCorrecao` | Várias linhas de texto | vazio |
| `DevolvidoPorNome` | Uma linha de texto | vazio |
| `DevolvidoPorEmail` | Uma linha de texto | vazio |
| `DataDevolucao` | Data e hora | vazio |

Não é necessário alterar as escolhas das colunas `Estado` ou `EstadoPagamento`.

## Funcionamento

1. Um Admin ou GestorFaturas abre uma nota aprovada e ainda não paga.
2. Escolhe **Devolver para correção** e indica o motivo.
3. A pessoa que submeteu a nota recebe um email.
4. No Dashboard, essa pessoa pode alterar apenas os valores ou KMs apresentados.
5. Ao guardar, a nota mantém o estado **Aprovado** e regressa diretamente a **Pagamentos**.
6. Quem devolveu a nota recebe um email a informar que os valores foram corrigidos.

Os anexos, datas, rubricas, descrições, matrícula, aprovadores e restantes dados não podem ser alterados neste fluxo.
