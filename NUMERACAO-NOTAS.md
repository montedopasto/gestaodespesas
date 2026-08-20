# Configuração da numeração das notas

Antes de publicar esta funcionalidade, criar na lista SharePoint `NotasDespesa` uma coluna com estas definições:

- Nome: `NumeroNota`
- Tipo: **Uma linha de texto**
- Exigir que esta coluna contenha informações exclusivas: **Sim**
- Valor predefinido: deixar vazio

As novas notas recebem automaticamente um número no formato `ND-ANO-SEQUÊNCIA`, por exemplo `ND-2026-000001`. A sequência reinicia no início de cada ano.

Os anexos das notas de tipo **Outras despesas** são guardados em `AnexosDespesas` com nomes como:

- `ND-2026-000001-01-1755700000000-fatura.pdf`
- `ND-2026-000001-02-1755700000001-recibo.jpg`

O valor numérico adicional evita que dois uploads simultâneos substituam acidentalmente o ficheiro um do outro.

As notas KM recebem número, mas não têm anexos.
