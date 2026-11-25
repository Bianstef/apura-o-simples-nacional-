# Apuração Simples Nacional

Skill para o Claude que calcula o DAS do Simples Nacional pelo método da planilha do escritório
(regras até 2026 e de 2027–2028, com CBS/IBS e a nova janela da RBT12), com ICMS/PR (isenção e redução),
vencimento do DAS e simulador de CBS/IBS para a nota fiscal.

## Arquivos

| Arquivo | O que é |
|---|---|
| `SKILL.md` | Instruções da skill (rotinas, método de cálculo, tabela CNAE → anexo, scripts) |
| `apuracao_simples.py` | Calcula o DAS e gera a planilha Excel da apuração |
| `ler_extrato_pgdas.py` | Lê o extrato do PGDAS-D (PDF) e devolve o faturamento mês a mês |

## Como usar

**Como skill no Claude:** baixe a pasta, compacte em .zip e envie em Configurações → Capabilities → Skills.

**Scripts direto no Python:**

```bash
pip install openpyxl
python3 apuracao_simples.py cliente.json "Apuracao Simples - CLIENTE - MM.AAAA.xlsx"
python3 ler_extrato_pgdas.py extrato.pdf     # precisa do pdftotext (poppler-utils)
```

O formato do `cliente.json` está no Passo 4 do `SKILL.md`.

## Atenção

- A planilha de histórico citada no `SKILL.md` é privada de quem criou a skill. Quem for usar precisa
  criar a sua e trocar o link/id.
- O cadastro de clientes vem da skill `analise-notas-de-entrada`, que não está neste repositório.
- Tabelas cadastradas até 2028. Confira sempre com o PGDAS-D/e-CAC. Não substitui a análise do contador.
