---
name: "apuracao-simples-nacional"
description: "Calcula o DAS do Simples Nacional (regras de 2027 com CBS/IBS e nova janela da RBT12) pelo método da planilha do escritório, guarda o faturamento mensal de cada cliente na planilha de histórico e entrega relatório e planilha Excel com a aba do simulador de CBS/IBS para a nota."
---

# Apuração Simples Nacional

Responda sempre em português do Brasil, de forma direta. A usuária é a Bianca (escritório contábil no PR; clientes do Simples Nacional no PR, SC e SP).

## Visão geral – duas rotinas

**A) Carga do histórico (extratos do PGDAS-D).** A Bianca manda o extrato do PGDAS-D (PDF da Receita) de cada cliente. Você lê o faturamento mês a mês (e a folha, quando houver) e grava na **planilha de histórico**. Ver "Rotina A".

**B) Apuração mensal.** Depois da carga, ela manda **só o faturamento do mês** de cada cliente (separado em comércio normal / c/ ST, indústria, serviço s/ e c/ retenção) e, para clientes com Fator R, a **folha do mês**. Você monta a RBT12 com o histórico guardado, calcula o DAS e **grava o faturamento (e a folha) do mês no histórico**, para os meses seguintes. Ver "Rotina B".

Ela não precisa mandar a RBT12. Se mandar, use a dela e avise se for diferente da que o histórico dá.

## Planilha de histórico (a "memória" do faturamento)

- **Histórico de Faturamento – Simples Nacional**: https://claude.ai/artifact/TqyKfvkp5m2Lzp8JuvhgM1 (id `d9640407-23fc-4fc1-8022-9e940faf94d4`). É uma planilha do tipo Sheets; a Bianca vê e pode corrigir à mão.
- Aba **Faturamento**, uma linha por cliente e mês, a partir da linha 2:

  | A CNPJ | B Cliente | C Competência | D Receita bruta (R$) | E Folha c/ encargos (R$) | F Origem | G Lançado em |
  |---|---|---|---|---|---|---|
  | 00.000.000/0001-00 | EMPRESA EXEMPLO LTDA | 10/2026 | 50000.00 | | Apuração | 2026-11-18 |

  - CNPJ completo e formatado (do cadastro). Competência sempre como texto `'MM/AAAA` (com o apóstrofo na escrita).
  - Receita bruta = total do mês (mercado interno + externo; todas as atividades, inclusive ST e retenção).
  - Folha c/ encargos = **FS12 do escritório: Tributável INSS + FGTS** (ver "Rotina C"). Gravada para todos os clientes que têm folha no SCI; só é usada no cálculo dos clientes com Fator R.
  - Origem: `Extrato PGDAS-D` (carga), `Apuração` (rotina B) ou `Informado` (a Bianca digitou ou ditou).
- **Como acessar** (conector Claude Docs): antes da primeira chamada da conversa, `guide(items=["topic.sheets"])`. Ache o nó: `read(ref={"object":"project","id":"d9640407-23fc-4fc1-8022-9e940faf94d4"})` → `files[0].content.id` (engine `workbook`). Depois `read(... payload={"kind":"readSheets"})` para a `revision` e o `filledRange`, e `readCells` da aba `Faturamento` (de A1 até a última linha; leia em partes de até 2.000 linhas).
- **Gravar**: `update` com `{"kind":"applyEdits","base":<revision>,"edits":[{"kind":"writeCells","sheet":"Faturamento","topLeft":"A<primeira linha vazia>","rows":[[...7 colunas...], ...]}]}`. Só acrescente linhas no fim; nunca reescreva linhas que já existem sem a Bianca pedir.
- **Duplicidade**: a chave é CNPJ + competência. Se o mês já existe com o mesmo valor, não grave de novo. Se existe com valor diferente, **não sobrescreva**: mostre os dois valores e pergunte qual vale (pode ser retificação).
- Se o conector Claude Docs não estiver disponível, diga isso, faça o cálculo com o que a Bianca mandar e liste as linhas que ficaram para gravar depois.

## Rotina A – Carga dos extratos do PGDAS-D

1. Para cada PDF, rode `python3 ler_extrato_pgdas.py extrato1.pdf extrato2.pdf ...` (script no fim da skill). Ele devolve CNPJ básico, nome, PA, receita do PA, os meses das "Receitas Brutas Anteriores" (interno + externo), a folha, e se a soma dos meses da RBT12 **confere** com a RBT12 do extrato.
2. Se `confere` for false ou faltar campo, abra o PDF (leia as páginas) e corrija os valores à mão. Nunca grave mês sem conferir.
3. Ache o CNPJ completo no cadastro (Passo 1) pelo CNPJ básico (8 primeiros dígitos) e pelo nome.
4. Leia o histórico do cliente na planilha e grave os meses que faltam (Origem `Extrato PGDAS-D`). Um extrato traz os meses anteriores + o PA.
5. Responda com uma tabela: cliente, meses gravados (de–até), meses que já existiam, divergências, e a RBT12 que o próximo PA vai usar.

**Carga grande (muitas linhas):** em vez de `writeCells`, grave um CSV (sem cabeçalho, competência **sem apóstrofo**, ex.: `10/2026`, números no padrão US), suba como asset da planilha (`upload_asset`), importe com `importFile` (`sheetName: "Carga"`), copie para a aba Faturamento com `copyCells` e apague a aba Carga. Com apóstrofo no CSV, a competência entra com o `'` literal.

## Rotina C – Carga da folha (planilha FOLHA_PGDAS_FS12 do RH)

A Bianca manda a planilha **FOLHA_PGDAS_FS12.xlsx** (fonte: Espelhos da Folha do SCI Visual Practice). Use a aba **Lançamentos**: colunas Cód., Empresa, **CNPJ**, **Competência** (data, dia 1), Tributável INSS, FGTS mensal, FGTS rescisório, Tributável INSS 13º, FGTS 13º/adiant., **Total PGDAS**, Soma 12 meses anteriores, Observação.

- **Folha do mês = Total PGDAS** = Tributável INSS (empregados + pró-labore + autônomos) + FGTS mensal + FGTS rescisório + Tributável INSS do 13º + FGTS do 13º/adiantamento. Confira que a soma das colunas E a I bate com a coluna J.
- Critérios do escritório: bolsa de estagiário **não** entra; multa/indenização compensatória do FGTS **não** entra; 13º: FGTS do adiantamento em 11, Tributável INSS do 13º e FGTS da 2ª parcela em 12.
- Linhas de meses futuros vêm vazias (Total 0): **ignore** competências depois do último mês fechado.
- "Sem espelho no SCI nesta competência" = folha 0 (empresa ainda sem folha ou encerrada).
- Gravação: preencha a **coluna E** das linhas que já existem (chave CNPJ + competência). Se a coluna E já tiver valor diferente (ex.: veio do extrato do PGDAS-D), **não sobrescreva**: mostre os dois e pergunte. Empresa ou mês sem linha de receita no histórico: não crie linha só com folha (a receita vazia viraria zero na RBT12); liste e peça o extrato.
- Na apuração mensal (Rotina B), se a Bianca não mandar a folha do mês de um cliente com Fator R, procure-a na planilha FS12 mais recente que ela mandar.

## Rotina B – Apuração mensal

1. Ache o cliente no cadastro (Passo 1) e defina o anexo (Passo 2). Para cliente com Fator R, peça também a **folha do mês** (vai em `folha_mes`).
2. Leia o histórico dele na planilha. A janela da RBT12 e da folha depende do ano da competência:
   - **Até 12/2026**: os 12 meses anteriores ao PA (ex.: PA 10/2026 → 10/2025 a 09/2026).
   - **A partir de 01/2027**: os 12 meses **antes do mês anterior** ao PA; o mês imediatamente anterior não entra (ex.: PA 01/2027 → 12/2025 a 11/2026; PA 05/2027 → 04/2026 a 03/2027).
   O script calcula a janela sozinho: mande todo o histórico do cliente em `historico` e ele escolhe os meses e avisa os que faltam.
3. Faltando mês na janela: pergunte o valor (ou o extrato). Se a Bianca mandar seguir, o script conta como zero e avisa.
4. Rode o script (Passo 4), entregue planilha e relatório (Passo 5). A planilha já sai com a aba **CBS e IBS – Nota fiscal** preenchida (ver "Simulador CBS/IBS").
5. **Grave no histórico** o faturamento total do mês apurado (e a folha do mês, se houver), Origem `Apuração`. Diga no relatório que gravou.

## Simulador CBS/IBS para a nota fiscal (aba da planilha de apuração)

A aba **CBS e IBS – Nota fiscal** é a réplica do "SIMULADOR COMPLETO – IBS – CBS – SIMPLES NACIONAL 2027–2028" do escritório (mesmos rótulos, fórmulas e tabelas de partilha 2027–2028), preenchida automaticamente pelo script a cada fechamento:

- **Para qual mês**: fechado o mês N, a soma de **N-11 a N** (inclui o mês fechado) é a RBT12 das notas e do DAS de **N+2** (regra de 2027). Ex.: fechou 11/2026 → RBT12 de 12/2025 a 11/2026 → notas de 01/2027; fechou 03/2027 → RBT12 de 04/2026 a 03/2027 → notas de 05/2027.
- **De onde vem a RBT12**: histórico da planilha (meses N-11 a N-1) + faturamento do mês fechado. Folha (Fator R): histórico + `folha_mes`. Sem histórico mês a mês, o script usa a RBT12 do DAS e avisa (ou use `rbt12_simulador` / `folha12_simulador`).
- **O que tem na aba**: no topo, o **resumo para passar ao cliente** (atividade, anexo, alíquota efetiva do DAS, % de CBS e % de IBS para destacar na nota); abaixo, um bloco do simulador para cada anexo das receitas do cliente (Anexo I, II, III, IV ou V; para Fator R, o bloco "Fator R – Anexo III ou V" com as tabelas dos Anexos III e V). Valor da nota de exemplo: R$ 1.000,00 (a Bianca pode mudar na planilha).
- **Fórmulas**: faixa pela RBT12; alíquota efetiva = (RBT12 × nominal − dedução) ÷ RBT12; CBS e IBS da nota = efetiva × % da partilha; na 5ª faixa dos Anexos III e IV, com o teto de 5% do ISS: CBS = (efetiva − 5%) × 23,20% e IBS = (efetiva − 5%) × 0,26% (Anexo III); CBS = (efetiva − 5%) × 36,27% e IBS = (efetiva − 5%) × 0,40% (Anexo IV).
- Se N+2 for antes de 01/2027, a aba avisa que é simulação; se for depois de 2028, avisa que a partilha precisa ser atualizada.
- No relatório da conversa, informe: **"Notas de MM/AAAA: CBS x,xxxx% e IBS x,xxxx%"** para cada atividade.

## Método de cálculo do escritório (obrigatório)

Os cálculos seguem a **PLANILHA SIMPLES NACIONAL 2018** do escritório (abas ANEXO I a V e uma aba por mês), conferida com o e-CAC no centavo (competências 07 a 10/2026). O script faz exatamente isto:

1. **RBT12** = soma do faturamento da janela do ano (ver Rotina B, passo 2). Folha do Fator R: mesma janela.
2. **Alíquota efetiva do SN** = (RBT12 × alíquota nominal − valor a deduzir) ÷ RBT12, **sem arredondar** (a planilha usa o valor cheio, ex.: 7,5079338%).
3. **ICMS pela repartição** = alíquota efetiva × % do ICMS na faixa (Anexo I: 34% nas faixas 1–2 e 33,5% nas 3–5; Anexo II: 32%).
4. **ICMS/PR – só clientes do PR, receitas de comércio e indústria sem ST:**
   - **RBT12 até R$ 360.000,00 → ICMS isento** (Lei PR 15.562/2007, art. 2º). A alíquota do mês fica "SN − ICMS" (ex.: 4% − 1,36% = 2,64%).
   - **RBT12 de R$ 360.000,01 a R$ 3.600.000,00 → redução.** ICMS efetivo PR = (RBT12 × alíquota PR − valor a deduzir PR) ÷ RBT12, pelas tabelas do Anexo XI do RICMS/PR (Decreto 8.660/2018):

     | Faixa RBT12 | Comércio (Tab. I) | Indústria (Tab. II) |
     |---|---|---|
     | até 360.000,00 | Isenção | Isenção |
     | 360.000,01 a 720.000,00 | 3,1825% / 11.457,00 | 3,2000% / 11.520,00 |
     | 720.000,01 a 1.800.000,00 | 3,5845% / 14.351,40 | 3,5840% / 14.284,80 |
     | 1.800.000,01 a 3.600.000,00 | 4,7905% / 36.059,40 | 4,7040% / 34.444,80 |

     **% de redução para o PGDAS-D** = (1 − ICMS efetivo PR ÷ ICMS da repartição) × 100, arredondado em 2 casas (ex.: 38,94%). O ICMS do DAS fica = ICMS da repartição × (1 − redução).
   - **Indústria:** o padrão do escritório (planilha e e-CAC) usa o **mesmo % de redução do comércio** também na indústria. O script faz assim por padrão (`"reducao_industria": "comercio"`) e mostra, na observação, o % que daria pela Tabela II. Se a Bianca pedir, use `"tabela2"`.
   - Receita **c/ ST**: sem ICMS (alíquota = SN − ICMS da repartição), em qualquer faixa.
5. **Serviço c/ retenção**: alíquota = SN − ISS da repartição (ex.: 6% − 2,01% = 3,99%).
6. **Valor do DAS**: cada tributo = faturamento × alíquota efetiva × % de repartição (com a isenção/redução do ICMS), **arredondado no centavo**; DAS = soma dos tributos. É assim que o valor bate com o PGDAS-D.
7. Clientes de **SC e SP**: não há benefício do PR. Calcule o ICMS pela tabela nacional e avise para confirmar benefício estadual.
8. **A partir de 01/2027 (Anexos XVIII a XXII da LC 123, incluídos pela LC 214/2025, vigentes em 2027 e 2028):**
   - **CBS no lugar de PIS e Cofins** e uma fatia pequena de **IBS** na partilha (ex.: Anexo I, 1ª faixa: CBS 15,33% e IBS 0,17%; Anexo III, 1ª faixa: CBS 15,43% e IBS 0,17%). ICMS e ISS com os mesmos percentuais de antes, então a isenção e a redução do ICMS/PR continuam iguais.
   - Faixas 1 a 5: alíquotas nominais e parcelas a deduzir **iguais**. 6ª faixa muda: Anexo I 18,90%, II 29,90%, III 32,90%, IV 32,90%, V 30,40%.
   - **Fim do monofásico de PIS/Cofins**: a segregação `monofasico` deixa de existir (o script ignora e avisa). Combustíveis têm CBS monofásica – conferir.
   - **Destaque de CBS e IBS na nota** (informativo, já incluído no DAS) = alíquota efetiva × % da partilha. O script mostra os percentuais na aba Repartição. Ex.: Anexo III, 1ª faixa, nota de R$ 487,00 → DAS R$ 29,22; CBS 0,9258% (R$ 4,51); IBS 0,0102% (R$ 0,05).
   - Cliente que optou por recolher **IBS/CBS fora do DAS** (regime regular, "Simples híbrido"): tratamento `ibs_cbs_fora` em todas as receitas.
   - Receita reconhecida pela **emissão do documento fiscal** (fim do regime de caixa).
   - Indústria: com o IPI zerado para a maior parte dos produtos, confira se a receita continua no Anexo II.
   - Competências de **2029 em diante**: o script para e avisa que as tabelas precisam ser cadastradas (a partilha muda todo ano até 2033).
9. **Vencimento do DAS**: dia 20 do mês seguinte à competência; se cair em sábado, domingo ou feriado nacional, **prorroga para o 1º dia útil seguinte** (ex.: competência 10/2026 → 20/11/2026 é Consciência Negra → vence 23/11/2026). O script calcula e coloca na planilha com a observação.

## Passo 1 – Achar o cliente no cadastro

O cadastro de clientes fica na skill **analise-notas-de-entrada**, seção "Cadastro de clientes (cartões CNPJ)". Leia sempre a versão atual de lá:

```bash
F=$(find / -path '*analise-notas-de-entrada/SKILL.md' 2>/dev/null | head -1)
awk '/^## Cadastro de clientes/{f=1} f' "$F" | grep -i "TERMO"   # TERMO = CNPJ (só dígitos ou formatado) ou parte do nome
```

Se o `find` não achar o arquivo, carregue a skill `analise-notas-de-entrada` com a ferramenta Skill e leia a tabela do fim dela.

- Compare pelo CNPJ ou por parte do nome/fantasia (várias empresas começam com "DIVA": confirme qual).
- Use o **CNAE principal e os secundários** para definir o anexo (Passo 2).
- Cliente que não está no cadastro: siga o Passo 3 da skill de análise de notas (BrasilAPI no navegador, depois WebSearch) e mostre no fim a linha pronta para o cadastro.

## Passo 2 – Definir o anexo pela atividade

Base legal: LC 123/2006, art. 18 (§§ 5º-B a 5º-M), Resolução CGSN 140/2018. O anexo depende do que gerou a receita, não só do CNAE principal.

**Regras gerais**
- Revenda de mercadorias → **Anexo I**. Restaurante, lanchonete e fornecimento de alimentação também → **Anexo I**.
- Fabricação/industrialização com venda do produto → **Anexo II**.
- Serviços em geral (manutenção, reparação, instalação, beleza, lavanderia, ensino, agência de viagens, corretagem de seguros, apoio administrativo etc.) → **Anexo III**.
- Locação de bens móveis → **Anexo III sem ISS** (tratamento `locacao_bens_moveis`).
- Construção civil (obra, empreitada, subempreitada), vigilância, limpeza, advocacia, decoração de interiores → **Anexo IV** (CPP fora do DAS).
- Serviços intelectuais/técnicos do § 5º-I e as atividades dos §§ 5º-B (XVI, XVIII a XXI) e 5º-D → **Fator R** (≥ 28% Anexo III; < 28% Anexo V): TI, publicidade (exceto promoção de vendas e marketing direto), consultoria, engenharia/arquitetura/desenho técnico, representação comercial e intermediação, fisioterapia, psicologia, nutrição, podologia, medicina, odontologia, academias e escolas de esportes, páginas eletrônicas.
- Transporte intermunicipal/interestadual de carga ou passageiros → Anexo III, com **ICMS no lugar do ISS** (avise; o PGDAS-D faz a troca).
- Cliente com comércio + serviço (ex.: oficina que vende peças, assistência técnica que vende acessórios, salão que vende produtos): cada receita no seu anexo. Se a usuária mandar um total só, calcule tudo pela atividade principal e **avise** que precisa da divisão.

**CNAEs do cadastro → anexo**

| CNAE | Anexo |
|---|---|
| 45.30-7-03, 45.30-7-05, todos os 46.xx de comércio atacadista (46.42 a 46.86), todos os 47.xx | I |
| 56.11-2-01, 56.11-2-03, 56.20-1-01, 56.20-1-02, 56.20-1-04 (alimentação) | I |
| 10.99-6-99, 22.29-3-02, 22.29-3-99, 25.42-0-00, 25.43-8-00, 26.51-5-00, 28.32-1-00, 29.49-2-99, 31.01-2-00 | II |
| 14.12-6-02 Confecção sob medida | II se a empresa fornece o tecido; III se o material é do cliente – CONFERIR |
| 25.39-0-01 Usinagem, tornearia e solda | III (§ 5º-B, IX) |
| 25.99-3-02 Corte e dobra de metais | III se serviço sob encomenda; II se fabrica e vende – CONFERIR |
| 33.13-9-01, 33.21-0-00, 33.29-5-01, 45.20-0-01/03/06/07, 95.11, 95.12, 95.21, 95.29-1-05 (manutenção, reparação, montagem, instalação) | III |
| 41.20-4-00, 43.30-4-02/03/04/05/99, 43.99-1-99 (obras) | IV |
| 43.21-5-00 Instalação elétrica, 43.22-3-02 Ar-condicionado | IV se for obra/empreitada de construção; III se instalação ou manutenção avulsa – CONFERIR |
| 49.29-9-03, 49.30-2-01, 49.30-2-02 (transporte) | III (intermunicipal: ICMS no lugar do ISS) |
| 52.11-7-99 Depósito para terceiros | III |
| 58.11-5-00, 58.19-1-00 (edição) | CONFERIR (livro é imune; edição pode ser I/II) |
| 59.11-1-02 Produção de filmes para publicidade | III – CONFERIR |
| 61.90-6-99 Telecomunicações | CONFERIR |
| 62.01-5-01, 62.04-0-00, 62.09-1-00, 63.19-4-00 (TI, páginas e portais) | Fator R |
| 63.99-2-00 Outros serviços de informação | III – CONFERIR |
| 64.93-0-00 Administração de consórcios | CONFERIR – atividade financeira, pode ser vedada ao Simples |
| 66.19-3-99 Auxiliares dos serviços financeiros | CONFERIR |
| 66.21-5-02 Auditoria e consultoria atuarial | Fator R |
| 66.22-3-00 Corretores de seguros | III |
| 70.20-4-00 Consultoria em gestão | Fator R |
| 71.19-7-03 Desenho técnico | Fator R |
| 73.11-4-00, 73.12-2-00, 73.19-0-04, 73.19-0-99 (publicidade) | Fator R |
| 73.19-0-02 Promoção de vendas, 73.19-0-03 Marketing direto | III (sem Fator R) |
| 74.10-2-02 Design de interiores | IV (decoração de interiores) – CONFERIR |
| 74.10-2-99 Outras atividades de design | Fator R – CONFERIR |
| 74.20-0-01, 74.20-0-04 (fotografia, filmagem) | III |
| 74.90-1-04, 74.90-1-05, 46.18-4-99, 46.19-2-00 (intermediação, agenciamento, representação) | Fator R |
| 77.23-3-00, 77.39-0-03, 77.39-0-99 (aluguel de bens móveis) | III sem ISS |
| 78.10-8-00 Seleção e agenciamento de mão de obra | III – CONFERIR |
| 79.11-2-00 Agência de viagens | III |
| 80.11-1-01 Vigilância | IV |
| 82.11-3-00, 82.19-9-99 (apoio administrativo) | III (sem Fator R) |
| 82.91-1-00 Cobranças e informações cadastrais | III – CONFERIR |
| 85.93-7-00 Idiomas, 85.99-6-03, 85.99-6-04 (treinamentos) | III |
| 85.91-1-00 Ensino de esportes, 93.13-1-00 Condicionamento físico | Fator R (academias e escolas de esportes) |
| 93.11-5-00, 93.12-3-00, 93.19-1-01, 93.29-8-99, 90.01-9-06 | III |
| 86.50-0-02 Nutrição, 86.50-0-03 Psicologia, 86.50-0-04 Fisioterapia, 86.90-9-04 Podologia | Fator R |
| 96.01-7-01 Lavanderia, 96.02-5-01, 96.02-5-02 (beleza e estética), 96.09-2-06 Tatuagem, 96.09-2-08 Pet | III |

CNAE fora desta tabela: aplique as regras gerais e marque CONFERIR. Em qualquer CONFERIR, faça a conta pela opção mais provável, diga qual assumiu e mostre o valor da outra opção numa linha.

## Passo 3 – Segregações (tratamentos)

Use os tratamentos conforme a divisão que a Bianca mandar. Se ela mandar um total só, avise quando a atividade costuma ter receita que paga menos no DAS:

| Tratamento | Quando | O que sai do DAS |
|---|---|---|
| `st` | Revenda com ICMS-ST (empresa substituída) ou antecipação com encerramento | ICMS |
| `st` (base legal) | LC 123, art. 13, §1º, XIII, "a" e art. 18, §4º-A, I; Res. CGSN 140, art. 25, §8º, I. Só o ICMS sai e só para o substituído; a receita continua na RBT12 | — |
| `monofasico` | **Só até 12/2026.** Produtos monofásicos de PIS/Cofins (perfumaria, cosméticos, higiene, farmacêuticos, autopeças, pneus, bebidas frias, combustíveis) | PIS e Cofins |
| `ibs_cbs_fora` | **A partir de 2027**, cliente que optou por recolher IBS e CBS pelo regime regular | CBS e IBS |
| `iss_retido` | Tomador reteve o ISS | ISS (mostre a alíquota de ISS da nota) |
| `exportacao` | Venda/serviço para o exterior | PIS, Cofins, CBS, IBS, ICMS, IPI, ISS |
| `locacao_bens_moveis` | Aluguel de bens móveis | ISS |
| `isencao_icms` / `isencao_iss` | Isenção concedida pelo estado/município | ICMS / ISS |

Alertas obrigatórios:
- **Lojas de cosméticos e perfumaria (47.72, as DIVA e similares)**: grande parte das vendas costuma ter ICMS-ST (e, até 12/2026, PIS/Cofins monofásico). Sem a divisão, o DAS sai **maior** do que o devido. Peça o relatório de vendas por CST/CSOSN (CSOSN 500 = ST; CST PIS/Cofins 04 = monofásico, só até 2026).
- **Autopeças, pneus, bebidas e materiais de construção**: mesmo alerta (ST e/ou monofásico).
- **Serviços com ISS retido**: pergunte se houve retenção no mês (a skill destaque-retencoes-servicos mostra as retenções das notas).

## Passo 4 – Calcular com o script

Grave o script do fim desta skill como `apuracao_simples.py` (com a ferramenta Write) e monte um JSON por cliente:

```json
{"cliente": "RAZÃO SOCIAL", "cnpj": "00.000.000/0001-00", "uf": "PR",
 "atividade": "CNAE principal e descrição", "competencia": "03/2027",
 "historico": [["01/2026", 55571.30, null], ["02/2026", 52846.40, null], "... todo o histórico do cliente na planilha ..."],
 "folha_mes": null,
 "receitas": [
   {"descricao": "Comércio normal", "anexo": "I", "valor": 32406.35},
   {"descricao": "Comércio c/ ST", "anexo": "I", "valor": 20666.61, "tratamentos": ["st"]},
   {"descricao": "Indústria", "anexo": "II", "valor": 9770.00},
   {"descricao": "Serviço s/ retenção", "anexo": "III", "valor": 10578.00},
   {"descricao": "Serviço c/ retenção", "anexo": "III", "valor": 0, "tratamentos": ["iss_retido"]}],
 "valor_ecac": null}
```

- `anexo`: `I`, `II`, `III`, `IV`, `V` ou `FATOR_R`.
- `historico`: lista `[competência, receita, folha ou null]` com **todos** os meses do cliente na planilha; o script escolhe a janela certa do ano. Se a Bianca mandar a RBT12 pronta, use `rbt12` (e `folha12`) no lugar de `historico`.
- `folha_mes`: folha do mês apurado (só Fator R); entra no simulador e vai para o histórico.
- Opcionais: `rbt12_simulador` e `folha12_simulador` (só sem histórico), `reducao_industria` (`comercio` ou `tabela2`), `valor_ecac`, `meses_atividade` (se < 12, os totais são a soma dos meses anteriores e o script proporcionaliza; 0 = primeiro mês).
- Tratamentos: `st`, `monofasico` (só até 2026), `iss_retido`, `exportacao`, `locacao_bens_moveis`, `isencao_icms`, `isencao_iss`, `sem_beneficio_pr`, `ibs_cbs_fora` (2027+).
- Categoria sem faturamento no mês: pode deixar de fora ou mandar com valor 0.

Grave também `ler_extrato_pgdas.py` (fim da skill) quando for carregar extratos.

Rode: `python3 apuracao_simples.py cliente.json "Apuracao Simples - RAZAO SOCIAL - MM.AAAA.xlsx"`

Depois recalcule as fórmulas com o `recalc.py` da skill xlsx (`find / -name recalc.py -path '*xlsx*'`) e confira que não há erros e que o DAS da planilha bate com o resumo do script.

**Como o cálculo funciona**: ver "Método de cálculo do escritório" acima. Também:
- Fator R = folha da janela ÷ RBT12. Folha = FS12 do escritório (Tributável INSS + FGTS, com 13º e FGTS rescisório; sem estagiário, sem multa do FGTS e sem distribuição de lucros) – ver Rotina C.
- 5ª faixa dos Anexos III e IV: ISS limitado a 5%, a diferença vai para os federais.
- RBT12 acima de R$ 3,6 mi: alerta de sublimite (ICMS/ISS podem sair do DAS) – conferir manualmente; acima de R$ 4,8 mi: alerta de exclusão.

## Passo 5 – Entrega

1. **Planilha** (uma por cliente), enviada como arquivo, com as abas:
   - **Apuração** (no formato da planilha do escritório): dados do cliente; meses da janela da RBT12; memória de cálculo do SN (RBT12, período, faixa, tabelas usadas, Fator R); memória do ICMS/PR (alíquota PR, valor a deduzir, ICMS efetivo PR, ICMS da repartição e **% de redução para o PGDAS-D**); imposto do mês por receita (normal, c/ ST, indústria, serviço s/ e c/ retenção) com anexo, faixa, alíquota nominal, parcela a deduzir, alíquota efetiva, ICMS/segregação, alíquota final, faturamento e DAS; total do DAS, alíquota final, **vencimento do DAS** (com a observação de prorrogação, se houver) e conferência com o e-CAC.
   - **Repartição**: cada tributo com % de repartição, alíquota, redução/exclusão e valor arredondado no centavo (a soma é o DAS); a partir de 2027, os % de CBS e IBS para a nota.
   - **CBS e IBS – Nota fiscal** (logo depois da Apuração): simulador do escritório preenchido com a RBT12 de N-11 a N, com os % de CBS e IBS das notas de N+2.
   - **Tabelas**: anexos I a V do período e tabelas do ICMS/PR.
   - Nome: `Apuracao Simples - RAZÃO SOCIAL - MM.AAAA.xlsx`. Valores em R$ com duas casas.
2. **Relatório na conversa**, curto, por cliente:
   - Cliente, CNPJ, CNAE e **anexo** (e se veio do cadastro ou de consulta).
   - **Período da RBT12** (janela do ano), RBT12, **faixa**, folha e **Fator R** (quando aplicável).
   - Alíquota nominal, parcela a deduzir, **alíquota efetiva**, situação do ICMS/PR (**isento** ou **% de redução para o PGDAS-D**) e **alíquota final**.
   - Faturamento do mês e **valor total do DAS**.
   - **% de CBS e IBS para as notas de N+2** (da aba do simulador) e, a partir de 2027, os do próprio mês (aba Repartição).
   - Confirmação de que o faturamento do mês foi **gravado no histórico**.
   - Alertas (ST/monofásico não informados, CONFERIR, sublimite, CPP do Anexo IV fora do DAS, ISS retido, meses faltando).
   - **Vencimento do DAS**: dia 20 do mês seguinte à competência; se cair em sábado, domingo ou feriado nacional, **prorroga para o 1º dia útil seguinte**. O script calcula e coloca na planilha; avise para conferir feriado local e prorrogações do CGSN.
3. Vários clientes de uma vez: uma planilha por cliente e, na conversa, uma tabela-resumo (cliente, anexo, faixa, alíquota final, faturamento, DAS, vencimento) com total geral.

## Checklist antes de entregar
- [ ] RBT12 pela janela do ano (2027+: sem o mês anterior ao PA) e sem mês faltando
- [ ] Faturamento do mês gravado no histórico, sem duplicar
- [ ] Vencimento do DAS na planilha e no relatório
- [ ] Aba CBS e IBS – Nota fiscal com a RBT12 de N-11 a N e o mês das notas (N+2) certo
- [ ] ICMS/PR: isento até R$ 360 mil; redução calculada acima disso; % de redução informado no relatório
- [ ] Conferido com o e-CAC quando a Bianca mandar o valor
- [ ] Anexo de cada receita conferido com CNAE principal e secundários
- [ ] Fator R calculado quando a atividade exige
- [ ] Alertas de ST/monofásico para comércio de cosméticos, autopeças, bebidas, materiais de construção
- [ ] Planilha recalculada sem erros e DAS igual ao do resumo
- [ ] Valores em R$ com duas casas

## Tabelas

O script tem as duas versões e escolhe pelo ano da competência:
- **Até 2026**: LC 123/2006 com redação da LC 155/2016 (PIS e Cofins na partilha).
- **2027 e 2028**: Anexos XVIII a XXII da LC 123/2006, incluídos pela LC 214/2025 (CBS e IBS na partilha; 6ª faixa com nominal nova).

Alíquota nominal / parcela a deduzir (iguais nos dois períodos, exceto a 6ª faixa):

| Faixa | I | II | III | IV | V |
|---|---|---|---|---|---|
| até 180.000 | 4% / 0 | 4,5% / 0 | 6% / 0 | 4,5% / 0 | 15,5% / 0 |
| até 360.000 | 7,3% / 5.940 | 7,8% / 5.940 | 11,2% / 9.360 | 9% / 8.100 | 18% / 4.500 |
| até 720.000 | 9,5% / 13.860 | 10% / 13.860 | 13,5% / 17.640 | 10,2% / 12.420 | 19,5% / 9.900 |
| até 1.800.000 | 10,7% / 22.500 | 11,2% / 22.500 | 16% / 35.640 | 14% / 39.780 | 20,5% / 17.100 |
| até 3.600.000 | 14,3% / 87.300 | 14,7% / 85.500 | 21% / 125.640 | 22% / 183.780 | 23% / 62.100 |
| até 4.800.000 (até 2026) | 19% / 378.000 | 30% / 720.000 | 33% / 648.000 | 33% / 828.000 | 30,5% / 540.000 |
| até 4.800.000 (2027–2028) | 18,9% / 378.000 | 29,9% / 720.000 | 32,9% / 648.000 | 32,9% / 828.000 | 30,4% / 540.000 |

A repartição por tributo de cada período e as tabelas do ICMS/PR estão no script.

## Script `apuracao_simples.py`

```python
#!/usr/bin/env python3
"""Apuração do Simples Nacional – método da planilha do escritório (PLANILHA SIMPLES NACIONAL 2018).

Uso:  python3 apuracao_simples.py entrada.json saida.xlsx      (imprime resumo JSON)

Escolhe as tabelas pelo ano da competência (até 2026: LC 155/2016; 2027-2028: Anexos XVIII a XXII, LC 214/2025)
e a janela da RBT12 pela regra do ano (a partir de 2027 o mês anterior ao PA não entra).
Entrada: historico [[MM/AAAA, receita, folha|null], ...] (ou rbt12/folha12, ou faturamento_12m/folha_12m) e receitas.
anexo: I, II, III, IV, V ou FATOR_R.  tratamentos: st, monofasico (só até 2026), iss_retido, exportacao,
locacao_bens_moveis, isencao_icms, isencao_iss, sem_beneficio_pr, ibs_cbs_fora
"""
import json, sys
from decimal import Decimal, ROUND_HALF_UP
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def rnd(x, n=2):
    return float(Decimal(repr(float(x))).quantize(Decimal(1).scaleb(-n), ROUND_HALF_UP))

FAIXAS = [180000, 360000, 720000, 1800000, 3600000, 4800000]
TRIB = ["IRPJ","CSLL","COFINS","PIS","CBS","CPP","IPI","ICMS","ISS","IBS"]
def R(**k): return {t: k.get(t,0.0) for t in TRIB}

# ---------- Até 2026: LC 123/2006 com redação da LC 155/2016 ----------
NOMINAL_2018 = {
 "I":  [(0.04,0),(0.073,5940),(0.095,13860),(0.107,22500),(0.143,87300),(0.19,378000)],
 "II": [(0.045,0),(0.078,5940),(0.10,13860),(0.112,22500),(0.147,85500),(0.30,720000)],
 "III":[(0.06,0),(0.112,9360),(0.135,17640),(0.16,35640),(0.21,125640),(0.33,648000)],
 "IV": [(0.045,0),(0.09,8100),(0.102,12420),(0.14,39780),(0.22,183780),(0.33,828000)],
 "V":  [(0.155,0),(0.18,4500),(0.195,9900),(0.205,17100),(0.23,62100),(0.305,540000)],
}
REPART_2018 = {
 "I": [R(IRPJ=.055,CSLL=.035,COFINS=.1274,PIS=.0276,CPP=.415,ICMS=.34)]*2 +
      [R(IRPJ=.055,CSLL=.035,COFINS=.1274,PIS=.0276,CPP=.42,ICMS=.335)]*3 +
      [R(IRPJ=.135,CSLL=.10,COFINS=.2827,PIS=.0613,CPP=.421)],
 "II":[R(IRPJ=.055,CSLL=.035,COFINS=.1151,PIS=.0249,CPP=.375,IPI=.075,ICMS=.32)]*5 +
      [R(IRPJ=.085,CSLL=.075,COFINS=.2096,PIS=.0454,CPP=.235,IPI=.35)],
 "III":[R(IRPJ=.04,CSLL=.035,COFINS=.1282,PIS=.0278,CPP=.434,ISS=.335),
        R(IRPJ=.04,CSLL=.035,COFINS=.1405,PIS=.0305,CPP=.434,ISS=.32),
        R(IRPJ=.04,CSLL=.035,COFINS=.1364,PIS=.0296,CPP=.434,ISS=.325),
        R(IRPJ=.04,CSLL=.035,COFINS=.1364,PIS=.0296,CPP=.434,ISS=.325),
        R(IRPJ=.04,CSLL=.035,COFINS=.1282,PIS=.0278,CPP=.434,ISS=.335),
        R(IRPJ=.35,CSLL=.15,COFINS=.1603,PIS=.0347,CPP=.305)],
 "IV":[R(IRPJ=.188,CSLL=.152,COFINS=.1767,PIS=.0383,ISS=.445),
       R(IRPJ=.198,CSLL=.152,COFINS=.2055,PIS=.0445,ISS=.40),
       R(IRPJ=.208,CSLL=.152,COFINS=.1973,PIS=.0427,ISS=.40),
       R(IRPJ=.178,CSLL=.192,COFINS=.189,PIS=.041,ISS=.40),
       R(IRPJ=.188,CSLL=.192,COFINS=.1808,PIS=.0392,ISS=.40),
       R(IRPJ=.535,CSLL=.215,COFINS=.2055,PIS=.0445)],
 "V": [R(IRPJ=.25,CSLL=.15,COFINS=.141,PIS=.0305,CPP=.2885,ISS=.14),
       R(IRPJ=.23,CSLL=.15,COFINS=.141,PIS=.0305,CPP=.2785,ISS=.17),
       R(IRPJ=.24,CSLL=.15,COFINS=.1492,PIS=.0323,CPP=.2385,ISS=.19),
       R(IRPJ=.21,CSLL=.15,COFINS=.1574,PIS=.0341,CPP=.2385,ISS=.21),
       R(IRPJ=.23,CSLL=.125,COFINS=.141,PIS=.0305,CPP=.2385,ISS=.235),
       R(IRPJ=.35,CSLL=.155,COFINS=.1644,PIS=.0356,CPP=.295)],
}
TETO_ISS_2018 = {"III": R(IRPJ=.0602,CSLL=.0526,COFINS=.1928,PIS=.0418,CPP=.6526),
                 "IV":  R(IRPJ=.3133,CSLL=.32,COFINS=.3013,PIS=.0654)}

# ---------- 2027 e 2028: Anexos XVIII a XXII da LC 123/2006 (incluídos pela LC 214/2025) ----------
NOMINAL_2027 = {
 "I":  [(0.04,0),(0.073,5940),(0.095,13860),(0.107,22500),(0.143,87300),(0.189,378000)],
 "II": [(0.045,0),(0.078,5940),(0.10,13860),(0.112,22500),(0.147,85500),(0.299,720000)],
 "III":[(0.06,0),(0.112,9360),(0.135,17640),(0.16,35640),(0.21,125640),(0.329,648000)],
 "IV": [(0.045,0),(0.09,8100),(0.102,12420),(0.14,39780),(0.22,183780),(0.329,828000)],
 "V":  [(0.155,0),(0.18,4500),(0.195,9900),(0.205,17100),(0.23,62100),(0.304,540000)],
}
REPART_2027 = {
 "I": [R(IRPJ=.055,CSLL=.035,CBS=.1533,CPP=.415,ICMS=.34,IBS=.0017)]*2 +
      [R(IRPJ=.055,CSLL=.035,CBS=.1533,CPP=.42,ICMS=.335,IBS=.0017)]*3 +
      [R(IRPJ=.1358,CSLL=.1006,CBS=.3402,CPP=.4234)],
 "II":[R(IRPJ=.055,CSLL=.035,CBS=.1385,CPP=.375,IPI=.075,ICMS=.32,IBS=.0015)]*5 +
      [R(IRPJ=.0853,CSLL=.0753,CBS=.2522,CPP=.2359,IPI=.3513)],
 "III":[R(IRPJ=.04,CSLL=.035,CBS=.1543,CPP=.434,ISS=.335,IBS=.0017),
        R(IRPJ=.04,CSLL=.035,CBS=.1691,CPP=.434,ISS=.32,IBS=.0019),
        R(IRPJ=.04,CSLL=.035,CBS=.1641,CPP=.434,ISS=.325,IBS=.0019),
        R(IRPJ=.04,CSLL=.035,CBS=.1641,CPP=.434,ISS=.325,IBS=.0019),
        R(IRPJ=.04,CSLL=.035,CBS=.1543,CPP=.434,ISS=.335,IBS=.0017),
        R(IRPJ=.3509,CSLL=.1504,CBS=.1929,CPP=.3058)],
 "IV":[R(IRPJ=.188,CSLL=.152,CBS=.2126,ISS=.445,IBS=.0024),
       R(IRPJ=.198,CSLL=.152,CBS=.2473,ISS=.40,IBS=.0027),
       R(IRPJ=.208,CSLL=.152,CBS=.2374,ISS=.40,IBS=.0026),
       R(IRPJ=.178,CSLL=.192,CBS=.2275,ISS=.40,IBS=.0025),
       R(IRPJ=.188,CSLL=.192,CBS=.2176,ISS=.40,IBS=.0024),
       R(IRPJ=.5371,CSLL=.2159,CBS=.2470)],
 "V": [R(IRPJ=.25,CSLL=.15,CBS=.1696,CPP=.2885,ISS=.14,IBS=.0019),
       R(IRPJ=.23,CSLL=.15,CBS=.1696,CPP=.2785,ISS=.17,IBS=.0019),
       R(IRPJ=.24,CSLL=.15,CBS=.1795,CPP=.2385,ISS=.19,IBS=.0020),
       R(IRPJ=.21,CSLL=.15,CBS=.1894,CPP=.2385,ISS=.21,IBS=.0021),
       R(IRPJ=.23,CSLL=.125,CBS=.1696,CPP=.2385,ISS=.235,IBS=.0019),
       R(IRPJ=.351,CSLL=.1554,CBS=.1978,CPP=.2958)],
}
TETO_ISS_2027 = {"III": R(IRPJ=.0602,CSLL=.0526,CBS=.2320,CPP=.6526,IBS=.0026),
                 "IV":  R(IRPJ=.3133,CSLL=.32,CBS=.3627,IBS=.0040)}

def tabelas(ano):
    if ano <= 2026: return NOMINAL_2018, REPART_2018, TETO_ISS_2018, "LC 123/2006 (redação LC 155/2016) – até 2026"
    if ano <= 2028: return NOMINAL_2027, REPART_2027, TETO_ISS_2027, "LC 123/2006, Anexos XVIII a XXII (LC 214/2025) – 2027 e 2028"
    raise SystemExit(f"Tabelas de {ano} não cadastradas no script (a partilha muda a cada ano a partir de 2029). Atualize a skill.")

NOMINAL, REPART, REPART_5_ISS_TETO = NOMINAL_2018, REPART_2018, TETO_ISS_2018   # trocadas em calcular() pelo ano

# ICMS/PR – Anexo XI do RICMS/PR (Dec. 8.660/2018). Faixas 1-2 (até 360 mil): isenção (Lei 15.562/2007, art. 2º)
ICMS_PR = {
 "I":  {2:(0.031825,11457.0), 3:(0.035845,14351.4), 4:(0.047905,36059.4)},   # Tabela I – comércio
 "II": {2:(0.032,11520.0),    3:(0.03584,14284.8),  4:(0.04704,34444.8)},    # Tabela II – indústria
}
ZERA = {"st":["ICMS"],"monofasico":["PIS","COFINS"],"iss_retido":["ISS"],
        "exportacao":["PIS","COFINS","CBS","IBS","ICMS","IPI","ISS"],"locacao_bens_moveis":["ISS"],
        "isencao_icms":["ICMS"],"isencao_iss":["ISS"],"ibs_cbs_fora":["CBS","IBS"]}
NOME_TRAT = {"st":"c/ ICMS-ST","monofasico":"PIS/Cofins monofásico","iss_retido":"ISS retido",
             "exportacao":"Exportação","locacao_bens_moveis":"Locação de bens móveis",
             "isencao_icms":"Isenção ICMS","isencao_iss":"Isenção ISS","sem_beneficio_pr":"Sem benefício ICMS-PR",
             "ibs_cbs_fora":"IBS/CBS fora do DAS (regime regular)"}

def mes_ano(txt):
    m, a = [int(x) for x in str(txt).replace("-", "/").split("/")[-2:]]
    return m, a

def soma_mes(m, a, n):
    t = a*12 + (m-1) + n
    return t % 12 + 1, t // 12

def janela_rbt12(competencia):
    """Meses que formam a RBT12 e a folha de 12 meses (FS12) da competência.
    Até 12/2026: os 12 meses anteriores ao período de apuração.
    A partir de 01/2027: os 12 meses antecedentes ao mês anterior ao PA (LC 123, art. 18, §1º e §1º-A, I e §5º-K,
    redação da LC 214/2025; Res. CGSN 190/2026). Ex.: PA 05/2027 → 04/2026 a 03/2027."""
    m, a = mes_ano(competencia)
    fim = 2 if a >= 2027 else 1
    return [f"{mm:02d}/{aa}" for mm, aa in (soma_mes(m, a, -k) for k in range(fim+11, fim-1, -1))]

def faixa_de(rbt):
    for i, lim in enumerate(FAIXAS):
        if rbt <= lim + 1e-9: return i
    return 5

def efetiva(anexo, rbt, f):
    if rbt <= 0: return NOMINAL[anexo][0][0]
    nom, pd = NOMINAL[anexo][f]
    return (rbt*nom - pd)/rbt

def repart(anexo, f, ef):
    p = dict(REPART[anexo][f])
    if f == 4 and anexo in REPART_5_ISS_TETO and ef*p["ISS"] > 0.05:
        fat = REPART_5_ISS_TETO[anexo]; resto = ef - 0.05
        p = {t: (fat[t]*resto/ef if t != "ISS" else 0.05/ef) for t in TRIB}
    return p

def icms_pr(tabela, rbt, f, icms_rep):
    """Retorna (situação, alíquota efetiva PR, % redução p/ PGDAS-D arredondado em 2 casas do %)."""
    if f <= 1: return ("isento", 0.0, 1.0)
    if f >= 5 or f not in ICMS_PR[tabela]: return ("sem_tabela", None, 0.0)
    nom, pd = ICMS_PR[tabela][f]
    ef_pr = (rbt*nom - pd)/rbt
    red = rnd((1 - ef_pr/icms_rep)*100, 2)/100 if icms_rep else 0.0
    return ("reducao", ef_pr, red)

def _pascoa(a):
    import datetime
    # algoritmo de Meeus/Jones/Butcher
    a1=a%19; b1=a//100; c1=a%100; d1=b1//4; e1=b1%4; f1=(b1+8)//25; g1=(b1-f1+1)//3
    h1=(19*a1+b1-d1-g1+15)%30; i1=c1//4; k1=c1%4; l1=(32+2*e1+2*i1-h1-k1)%7; m1=(a1+11*h1+22*l1)//451
    mes=(h1+l1-7*m1+114)//31; dia=((h1+l1-7*m1+114)%31)+1
    return datetime.date(a, mes, dia)

def feriados_nacionais(ano):
    import datetime
    D=datetime.date; p=_pascoa(ano); td=datetime.timedelta
    fixos={D(ano,1,1):"Confraternização Universal",D(ano,4,21):"Tiradentes",D(ano,5,1):"Dia do Trabalho",
           D(ano,9,7):"Independência",D(ano,10,12):"N. Sra. Aparecida",D(ano,11,2):"Finados",
           D(ano,11,15):"Proclamação da República",D(ano,11,20):"Consciência Negra",D(ano,12,25):"Natal"}
    fixos[p-td(days=48)]="Carnaval"; fixos[p-td(days=47)]="Carnaval"
    fixos[p-td(days=2)]="Sexta-feira Santa"; fixos[p+td(days=60)]="Corpus Christi"
    return fixos

def vencimento_das(competencia):
    """Dia 20 do mês seguinte; se não for dia útil, prorroga para o 1º dia útil seguinte (Res. CGSN 140/2018)."""
    import datetime
    try:
        m, a = [int(x) for x in str(competencia).replace("-", "/").split("/")[-2:]]
    except Exception:
        return None, ""
    m += 1
    if m == 13: m, a = 1, a + 1
    d = datetime.date(a, m, 20); motivo = []
    fer = feriados_nacionais(a)
    while d.weekday() >= 5 or d in fer:
        motivo.append(fer.get(d, "sábado" if d.weekday() == 5 else "domingo"))
        d += datetime.timedelta(days=1)
    obs = f"Vencimento do DAS: {d.strftime('%d/%m/%Y')}"
    if motivo:
        obs += f" (dia 20/{m:02d}/{a} cai em {motivo[0]}; prorrogado para o 1º dia útil seguinte)"
    else:
        obs += " (dia 20 do mês seguinte à competência)"
    obs += ". Confira feriado local e eventual prorrogação do CGSN."
    return d, obs

def soma_meses(v):
    if v is None: return None, []
    if isinstance(v, (int, float)): return float(v), []
    lst = [(str(m), float(x)) for m, x in v]
    return sum(x for _, x in lst), lst

def calcular(d):
    global NOMINAL, REPART, REPART_5_ISS_TETO
    alertas = []
    _, ano = mes_ano(d.get("competencia"))
    NOMINAL, REPART, REPART_5_ISS_TETO, base_tabelas = tabelas(ano)
    janela = janela_rbt12(d["competencia"])
    hist_sim = None
    if d.get("historico") is not None:
        # historico: [["MM/AAAA", receita, folha_ou_null], ...] – o script escolhe os meses da janela
        hist = {}
        for item in d["historico"]:
            mm, aa = mes_ano(item[0]); k = f"{mm:02d}/{aa}"
            hist[k] = (float(item[1] or 0), (float(item[2]) if len(item) > 2 and item[2] is not None else None))
        hist_sim = hist
        faltando = [k for k in janela if k not in hist]
        if faltando and d.get("meses_atividade", 12) >= 12:
            alertas.append("Faltam no histórico os meses " + ", ".join(faltando) + " da janela da RBT12 – calculado como zero. Confira.")
        d = dict(d)
        d["faturamento_12m"] = [[k, hist.get(k, (0.0, None))[0]] for k in janela]
        if any(hist.get(k, (0, None))[1] is not None for k in janela):
            d["folha_12m"] = [[k, hist.get(k, (0, None))[1] or 0.0] for k in janela]
    rbt_soma, meses_fat = soma_meses(d.get("faturamento_12m"))
    folha_soma, meses_folha = soma_meses(d.get("folha_12m"))
    rbt_inf = rbt_soma if rbt_soma is not None else float(d["rbt12"])
    folha_inf = folha_soma if folha_soma is not None else float(d.get("folha12") or 0)
    if meses_fat and len(meses_fat) != 12 and d.get("meses_atividade", 12) >= 12:
        alertas.append(f"Foram informados {len(meses_fat)} meses de faturamento (esperado 12). Confira.")
    meses = d.get("meses_atividade", 12)
    receita_mes_total = sum(float(r["valor"]) for r in d["receitas"])
    if meses < 12:
        if meses == 0:
            rbt = receita_mes_total*12; folha = folha_inf*12
            alertas.append("1º mês de atividade: RBT12 = receita do mês × 12.")
        else:
            rbt = rbt_inf/meses*12; folha = folha_inf/meses*12
            alertas.append(f"Início de atividade ({meses} meses): RBT12 = média dos meses anteriores × 12.")
    else:
        rbt, folha = rbt_inf, folha_inf
    uf = (d.get("uf") or "").upper()
    if rbt > 4800000:
        alertas.append("RBT12 acima de R$ 4,8 mi: verificar exclusão do Simples.")
    f = faixa_de(rbt)
    fator_r = (folha/rbt) if rbt > 0 else None
    usa_fr = any(r["anexo"] == "FATOR_R" for r in d["receitas"])
    if usa_fr and not folha_inf:
        alertas.append("Atividade com Fator R sem folha informada: calculado como Anexo V. Confirme a folha.")
    if f == 5:
        alertas.append("RBT12 na 6ª faixa: ICMS/ISS podem ficar fora do DAS (sublimite R$ 3,6 mi). Conferir manualmente.")
    red_ind = d.get("reducao_industria", "comercio")

    # ICMS PR do comércio (usado também na indústria no padrão do escritório)
    pr_com = None
    if uf == "PR" and f < 5:
        ef1 = efetiva("I", rbt, f); pr_com = icms_pr("I", rbt, f, ef1*REPART["I"][f]["ICMS"])

    linhas = []
    for r in d["receitas"]:
        anexo = r["anexo"]
        if anexo == "FATOR_R":
            anexo = "III" if (fator_r or 0) >= 0.28 else "V"
            origem = f"Fator R {fator_r*100:.2f}% → Anexo {anexo}" if fator_r is not None else "Fator R"
        else:
            origem = f"Anexo {anexo} (atividade)"
        nom, pd = NOMINAL[anexo][f]
        ef = efetiva(anexo, rbt, f)
        p = repart(anexo, f, ef)
        trats = list(r.get("tratamentos", []))
        if ano >= 2027 and "monofasico" in trats:
            trats.remove("monofasico")
            alertas.append(f"{r.get('descricao','')}: o regime monofásico de PIS/Cofins acabou em 2027 (PIS e Cofins viraram CBS); receita tributada normalmente. Combustíveis têm CBS monofásica – conferir.")
        zerados = set(t for tr in trats for t in ZERA.get(tr, []))
        aliq = {t: ef*p[t] for t in TRIB}
        icms_rep = aliq["ICMS"]
        pr = None; red_icms = 0.0; obs_icms = ""
        if uf == "PR" and anexo in ("I","II") and "ICMS" not in zerados and "sem_beneficio_pr" not in trats and f < 5:
            pr = icms_pr(anexo, rbt, f, icms_rep)
            if pr[0] == "isento":
                red_icms = 1.0; obs_icms = "ICMS-PR isento (RBT12 até R$ 360 mil – Lei 15.562/2007, art. 2º)"
            elif pr[0] == "reducao":
                red_icms = pr[2]
                if anexo == "II" and red_ind == "comercio" and pr_com and pr_com[0] == "reducao":
                    red_tab2 = red_icms; red_icms = pr_com[2]
                    obs_icms = (f"Redução ICMS-PR {red_icms*100:.2f}% (padrão do escritório: mesmo % do comércio; "
                                f"pela Tabela II da indústria seria {red_tab2*100:.2f}%)")
                else:
                    obs_icms = f"Redução ICMS-PR {red_icms*100:.2f}% (informar no PGDAS-D)"
        elif uf in ("SC","SP") and anexo in ("I","II") and "ICMS" not in zerados:
            obs_icms = f"Cliente de {uf}: sem benefício de ICMS-PR; ICMS pela tabela nacional (confirmar benefício estadual)."
        final = {}
        for t in TRIB:
            if t in zerados: final[t] = 0.0
            elif t == "ICMS": final[t] = aliq[t]*(1 - red_icms)
            else: final[t] = aliq[t]
        valor = float(r["valor"])
        partes = {t: rnd(valor*final[t]) for t in TRIB}
        das = round(sum(partes.values()), 2)
        aliq_final = sum(final.values())
        cbs_nota = aliq["CBS"] if "CBS" not in zerados else 0.0
        ibs_nota = aliq["IBS"] if "IBS" not in zerados else 0.0
        iss_nota = None
        if anexo in ("III","IV","V"):
            iss_nota = max(aliq["ISS"], 0.02) if aliq["ISS"] else None
        linhas.append(dict(descricao=r.get("descricao",""), anexo=anexo, origem=origem, faixa=f+1,
            nominal=nom, pd=pd, efetiva=ef, perc=p, aliq=aliq, final=final, tratamentos=trats,
            zerados=sorted(zerados), red_icms=red_icms, icms_pr=pr, obs_icms=obs_icms,
            aliq_final=aliq_final, receita=valor, das=das, partes=partes, iss_nota=iss_nota,
            cbs_nota=cbs_nota, ibs_nota=ibs_nota))
        if obs_icms and "Cliente de" in obs_icms: alertas.append(obs_icms)
    das_total = round(sum(l["das"] for l in linhas), 2)
    for l in linhas:
        if "iss_retido" in l["tratamentos"] and l["iss_nota"]:
            alertas.append(f"{l['descricao']}: ISS retido – alíquota de ISS na nota ≈ {l['iss_nota']*100:.2f}%.")
        if l["anexo"] == "IV":
            alertas.append(f"{l['descricao']}: Anexo IV – CPP (20% s/ folha) fora do DAS.")
    if ano >= 2027:
        alertas.append(f"Regras de {ano}: RBT12 e folha de {janela[0]} a {janela[-1]} (12 meses antes do mês anterior ao PA); CBS no lugar de PIS/Cofins e IBS na partilha; receita pela emissão do documento fiscal (fim do regime de caixa).")
        if any(l["anexo"] == "II" for l in linhas):
            alertas.append("Indústria em 2027: com o IPI reduzido a zero para a maior parte dos produtos, confira se a receita continua no Anexo II.")
    # ---- Simulador CBS/IBS para as notas fiscais ----
    # Fechado o mês N, a soma de N-11 a N é a RBT12 (e a folha) das notas e do DAS de N+2 (regra de 2027).
    mN, aN = mes_ano(d["competencia"])
    meses_sim = [f"{mm:02d}/{aa}" for mm, aa in (soma_mes(mN, aN, -k) for k in range(11, -1, -1))]
    tm, ta = soma_mes(mN, aN, 2); alvo = f"{tm:02d}/{ta}"
    sim_obs = []
    if hist_sim is not None:
        falt = [k for k in meses_sim[:-1] if k not in hist_sim]
        rbt_sim = sum(hist_sim.get(k, (0.0, None))[0] for k in meses_sim[:-1]) + receita_mes_total
        folha_mes = d.get("folha_mes")
        folha_sim = sum((hist_sim.get(k, (0, None))[1] or 0.0) for k in meses_sim[:-1]) + float(folha_mes or 0)
        if falt: sim_obs.append("Faltam no histórico " + ", ".join(falt) + " (contados como zero).")
        if usa_fr and folha_mes is None: sim_obs.append("Folha do mês não informada: contada como zero no Fator R do simulador.")
    elif d.get("rbt12_simulador") is not None:
        rbt_sim = float(d["rbt12_simulador"]); folha_sim = float(d.get("folha12_simulador") or folha)
    else:
        rbt_sim, folha_sim = rbt, folha
        sim_obs.append("Sem histórico mês a mês: o simulador usou a mesma RBT12 do DAS. Confira.")
    if ta < 2027: sim_obs.append("CBS e IBS só entram no Simples em 01/2027: estes percentuais são simulação.")
    if ta > 2028: sim_obs.append(f"Partilha de {ta} ainda não cadastrada: o simulador usa a de 2027–2028. Atualize a skill.")
    anexos_sim = []
    for r in d["receitas"]:
        a_ = "FATOR_R" if r["anexo"] == "FATOR_R" else r["anexo"]
        if a_ not in anexos_sim: anexos_sim.append(a_)
    simulador = dict(alvo=alvo, meses=meses_sim, rbt12=rbt_sim, folha12=folha_sim, anexos=anexos_sim, obs=sim_obs)

    venc, venc_obs = vencimento_das(d.get("competencia",""))
    if venc_obs: alertas.insert(0, venc_obs)
    ecac = d.get("valor_ecac")
    if ecac is not None:
        dif = round(das_total - float(ecac), 2)
        alertas.append(f"Conferência e-CAC: calculado R$ {das_total:,.2f} x e-CAC R$ {float(ecac):,.2f} (diferença R$ {dif:,.2f}).")
    return dict(cliente=d.get("cliente",""), cnpj=d.get("cnpj",""), uf=uf, atividade=d.get("atividade",""),
        competencia=d.get("competencia",""), ano=ano, janela=janela, base_tabelas=base_tabelas, rbt12_informado=rbt_inf, rbt12=rbt, folha12=folha,
        meses_fat=meses_fat, meses_folha=meses_folha, fator_r=fator_r, faixa=f+1, receita_mes=receita_mes_total,
        das_total=das_total, aliquota_final=(das_total/receita_mes_total if receita_mes_total else 0),
        linhas=linhas, alertas=alertas, meses=meses, usa_fr=usa_fr, pr_com=pr_com, valor_ecac=ecac,
        vencimento=(venc.strftime("%d/%m/%Y") if venc else ""), vencimento_obs=venc_obs, simulador=simulador)

# ---------------- planilha ----------------
AZUL = PatternFill("solid", fgColor="1F4E78"); CINZA = PatternFill("solid", fgColor="D9E1F2")
AMAR = PatternFill("solid", fgColor="FFF2CC"); VERDE = PatternFill("solid", fgColor="E2EFDA")
B = Side(style="thin", color="999999"); BORDA = Border(left=B, right=B, top=B, bottom=B)
MOEDA = 'R$ #,##0.00'; PCT = '0.00%'; PCT4 = '0.0000%'

def cab(ws, row, cols, col0=1):
    for i, c in enumerate(cols, col0):
        x = ws.cell(row=row, column=i, value=c)
        x.font = Font(bold=True, color="FFFFFF"); x.fill = AZUL; x.border = BORDA
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def titulo(ws, r, txt, size=12):
    ws.cell(row=r, column=1, value=txt).font = Font(bold=True, size=size)

# ---------------- aba do simulador CBS/IBS (réplica do "SIMULADOR COMPLETO - IBS - CBS - SIMPLES NACIONAL 2027-2028") ----------------
NAVY = PatternFill("solid", fgColor="1B365D"); CREME = PatternFill("solid", fgColor="FFFBEB"); VERDE_NF = PatternFill("solid", fgColor="E6F4EA")
TITULOS_SIM = {"I": "ANEXO I - COMÉRCIO", "II": "ANEXO II - INDÚSTRIA", "III": "ANEXO III - SERVIÇOS",
               "IV": "ANEXO IV - SERVIÇOS", "V": "ANEXO V - SERVIÇOS"}
COLS_SIM = {"I": ["IRPJ","CSLL","CBS","CPP","ICMS","IBS"], "II": ["IRPJ","CSLL","CBS","CPP","IPI","ICMS","IBS"],
            "III": ["IRPJ","CSLL","CBS","CPP","ISS","IBS"], "IV": ["IRPJ","CSLL","CBS","ISS","IBS"],
            "V": ["IRPJ","CSLL","CBS","CPP","ISS","IBS"]}
TETO_SIM = {"III": (0.335, 0.232, 0.0026), "IV": (0.40, 0.3627, 0.004)}   # ISS da faixa 5, fator CBS, fator IBS

def _faixa_f(c):
    return f'=IF({c}>4800000,"EXCEDE LIMITE",IF({c}<=180000,1,IF({c}<=360000,2,IF({c}<=720000,3,IF({c}<=1800000,4,IF({c}<=3600000,5,6))))))'

def _barra(ws, r, txt, ncols):
    for c in range(1, ncols + 1):
        ws.cell(row=r, column=c).fill = NAVY
    x = ws.cell(row=r, column=1, value=txt); x.font = Font(bold=True, color="FFFFFF")

def _tabela_sim(ws, r, anexo):
    """Escreve cabeçalho + 6 faixas da partilha 2027-2028; devolve a linha da 1ª faixa."""
    cols = ["Faixa", "RBT12 Até", "Alíq. Nom.", "Dedução"] + COLS_SIM[anexo]
    for i, h in enumerate(cols, 1):
        x = ws.cell(row=r, column=i, value=h); x.font = Font(bold=True, color="FFFFFF"); x.fill = NAVY
        x.alignment = Alignment(horizontal="center")
    for f in range(6):
        rr = r + 1 + f
        nom, pd = NOMINAL_2027[anexo][f]
        ws.cell(row=rr, column=1, value=f + 1)
        ws.cell(row=rr, column=2, value=FAIXAS[f]).number_format = MOEDA
        ws.cell(row=rr, column=3, value=nom).number_format = PCT
        ws.cell(row=rr, column=4, value=pd).number_format = MOEDA
        for j, t in enumerate(COLS_SIM[anexo]):
            ws.cell(row=rr, column=5 + j, value=REPART_2027[anexo][f][t]).number_format = PCT
        for c in range(1, len(cols) + 1): ws.cell(row=rr, column=c).border = BORDA
    return r + 1

def _bloco_anexo(ws, r, anexo, rbt, nota=1000.0):
    """Réplica de uma aba de anexo do simulador. Devolve (próxima linha, célula efetiva, célula CBS, célula IBS)."""
    ncol = 4 + len(COLS_SIM[anexo])
    _barra(ws, r, f"{TITULOS_SIM[anexo]} - SIMPLES NACIONAL 2027/2028 (LC 214/2025 e Res. CGSN 190/2026)", ncol)
    R0 = r + 2
    lab = ["Faturamento Bruto 12 Meses (RBT12):", "Valor da Nota Fiscal Emitida (R$):"]
    for i, t in enumerate(lab): ws.cell(row=R0 + i, column=1, value=t).font = Font(bold=True)
    c = ws.cell(row=R0, column=3, value=rbt); c.number_format = MOEDA; c.fill = CREME; c.font = Font(bold=True)
    c = ws.cell(row=R0 + 1, column=3, value=nota); c.number_format = MOEDA; c.fill = CREME; c.font = Font(bold=True)
    t0 = R0 + 12                       # linha do cabeçalho da tabela
    f1 = t0 + 1; f6 = t0 + 6           # faixas 1 a 6
    cR, cN = f"C{R0}", f"C{R0+1}"
    rf = R0 + 3; cF = f"$C${rf}"
    col_cbs = "G"; col_ibs = get_column_letter(4 + len(COLS_SIM[anexo]))
    linhas = [
        ("Faixa Enquadrada:", _faixa_f(cR), "0"),
        ("Alíquota Nominal:", f'=IF(ISNUMBER({cF}),INDEX(C{f1}:C{f6},C{rf}),"—")', PCT),
        ("Parcela a Deduzir:", f'=IF(ISNUMBER({cF}),INDEX(D{f1}:D{f6},C{rf}),"—")', MOEDA),
        ("Alíquota Efetiva Total do DAS:", f'=IF(ISNUMBER({cF}),IF({cR}>0,(({cR}*C{rf+1})-C{rf+2})/{cR},C{rf+1}),"—")', PCT4),
        ("Valor Total do DAS (R$):", f'=IF(ISNUMBER({cF}),ROUND({cN}*C{rf+3},2),"—")', MOEDA),
    ]
    ce = f"$C${rf+3}"
    if anexo in TETO_SIM:
        iss, kc, ki = TETO_SIM[anexo]
        cbs_f = f'=IF(ISNUMBER({cF}),IF(AND({cF}=5,{ce}>0.05/{iss}),({ce}-0.05)*{kc},C{rf+3}*INDEX({col_cbs}{f1}:{col_cbs}{f6},C{rf})),"—")'
        ibs_f = f'=IF(ISNUMBER({cF}),IF(AND({cF}=5,{ce}>0.05/{iss}),({ce}-0.05)*{ki},C{rf+3}*INDEX({col_ibs}{f1}:{col_ibs}{f6},C{rf})),"—")'
    else:
        cbs_f = f'=IF(ISNUMBER({cF}),C{rf+3}*INDEX({col_cbs}{f1}:{col_cbs}{f6},C{rf}),"—")'
        ibs_f = f'=IF(ISNUMBER({cF}),C{rf+3}*INDEX({col_ibs}{f1}:{col_ibs}{f6},C{rf}),"—")'
    linhas += [
        ("Alíquota Efetiva CBS (Destaque NF):", cbs_f, PCT4),
        ("Valor CBS Destaque NF (R$):", f'=IF(ISNUMBER({cF}),ROUND({cN}*C{rf+5},2),"—")', MOEDA),
        ("Alíquota Efetiva IBS (Destaque NF):", ibs_f, PCT4),
        ("Valor IBS Destaque NF (R$):", f'=IF(ISNUMBER({cF}),ROUND({cN}*C{rf+7},2),"—")', MOEDA),
    ]
    for i, (lbl, fml, fmt) in enumerate(linhas):
        rr = rf + i
        ws.cell(row=rr, column=1, value=lbl).font = Font(bold=True)
        c = ws.cell(row=rr, column=3, value=fml); c.number_format = fmt; c.font = Font(bold=True)
        if "Destaque" in lbl and "Alíquota" in lbl:
            c.fill = VERDE_NF; c.font = Font(bold=True, color="0B4F32")
    ws.cell(row=rf, column=4, value=f'=IF(ISNUMBER(C{rf}),"","RBT12 acima de R$ 4,8 milhões: fora do Simples Nacional. Sem cálculo.")')
    if anexo in TETO_SIM:
        iss = TETO_SIM[anexo][0]
        ws.cell(row=rf + 5, column=4, value=f'=IF(AND(ISNUMBER({cF}),AND({cF}=5,{ce}>0.05/{iss})),"Teto do ISS aplicado: ISS fixo em 5%; CBS e IBS redistribuídos (LC 214/2025)","")')
    _tabela_sim(ws, t0, anexo)
    return t0 + 8, f"C{rf+3}", f"C{rf+5}", f"C{rf+7}"

def _bloco_fator_r(ws, r, rbt, folha, nota=1000.0):
    """Réplica da aba 'Fator R (III ou V)', com as tabelas dos Anexos III e V logo abaixo."""
    _barra(ws, r, "FATOR R – ANEXO III OU V – SIMPLES NACIONAL 2027/2028 (LC 214/2025)", 6)
    R0 = r + 2
    for i, (t, v) in enumerate([("Faturamento Bruto 12 Meses (RBT12):", rbt),
                                ("Folha de salários 12 meses (com pró-labore, INSS patronal e FGTS):", folha),
                                ("Valor da Nota Fiscal Emitida (R$):", nota)]):
        ws.cell(row=R0 + i, column=1, value=t).font = Font(bold=True)
        c = ws.cell(row=R0 + i, column=3, value=v); c.number_format = MOEDA; c.fill = CREME; c.font = Font(bold=True)
    cR, cFo, cN = f"$C${R0}", f"$C${R0+1}", f"$C${R0+2}"
    r7 = R0 + 4
    ws.cell(row=r7, column=1, value="Fator R (folha ÷ RBT12):").font = Font(bold=True)
    c = ws.cell(row=r7, column=3, value=f'=IF({cR}>0,{cFo}/{cR},"—")'); c.number_format = PCT; c.font = Font(bold=True)
    ws.cell(row=r7, column=4, value=f'=IF(ISNUMBER(C{r7}),IF(C{r7}>=0.28,"≥ 28%: tributa pelo Anexo III","< 28%: tributa pelo Anexo V"),"Informe o RBT12")')
    ws.cell(row=r7 + 1, column=1, value="Anexo aplicável:").font = Font(bold=True)
    ws.cell(row=r7 + 1, column=3, value=f'=IF(ISNUMBER(C{r7}),IF(C{r7}>=0.28,"III","V"),"—")').font = Font(bold=True)
    rF = r7 + 2
    ws.cell(row=rF, column=1, value="Faixa Enquadrada:").font = Font(bold=True)
    ws.cell(row=rF, column=3, value=_faixa_f(cR)).font = Font(bold=True)
    ws.cell(row=rF, column=4, value=f'=IF(ISNUMBER(C{rF}),"","RBT12 acima de R$ 4,8 milhões: fora do Simples Nacional. Sem cálculo.")')
    cF, cA = f"$C${rF}", f"$C${r7+1}"
    rs = rF + 2
    for i, h in enumerate(["Simulação", "", "Anexo III", "Anexo V", "APLICÁVEL"], 1):
        x = ws.cell(row=rs, column=i, value=h or None); x.font = Font(bold=True, color="FFFFFF"); x.fill = NAVY
    t3 = rs + 11; t5 = t3 + 9        # tabelas III e V abaixo da simulação
    f3, f3b = t3 + 1, t3 + 6; f5, f5b = t5 + 1, t5 + 6
    linhas = ["Alíquota Nominal", "Parcela a Deduzir", "Alíquota Efetiva Total do DAS", "Valor Total do DAS (R$)",
              "Alíquota Efetiva CBS (Destaque NF)", "Valor CBS Destaque NF (R$)", "Alíquota Efetiva IBS (Destaque NF)", "Valor IBS Destaque NF (R$)"]
    fmts = [PCT, MOEDA, PCT4, MOEDA, PCT4, MOEDA, PCT4, MOEDA]
    for i, lbl in enumerate(linhas):
        rr = rs + 1 + i
        ws.cell(row=rr, column=1, value=lbl).font = Font(bold=True)
    def F(col, i, fb, fe):
        e = f"{col}{rs+3}"
        if i == 0: return f'=IF(ISNUMBER({cF}),INDEX($C${fb}:$C${fe},{cF}),"—")'
        if i == 1: return f'=IF(ISNUMBER({cF}),INDEX($D${fb}:$D${fe},{cF}),"—")'
        if i == 2: return f'=IF(ISNUMBER({cF}),IF({cR}>0,({cR}*{col}{rs+1}-{col}{rs+2})/{cR},{col}{rs+1}),"—")'
        if i == 3: return f'=IF(ISNUMBER({cF}),ROUND({cN}*{e},2),"—")'
        if i in (4, 6):
            tab = "G" if i == 4 else "J"
            base = f"{e}*INDEX(${tab}${fb}:${tab}${fe},{cF})"
            if col == "C":   # Anexo III: teto do ISS na 5ª faixa
                k = 0.232 if i == 4 else 0.0026
                return f'=IF(ISNUMBER({cF}),IF(AND({cF}=5,{e}>0.05/0.335),({e}-0.05)*{k},{base}),"—")'
            return f'=IF(ISNUMBER({cF}),{base},"—")'
        return f'=IF(ISNUMBER({cF}),ROUND({cN}*{col}{rs+i},2),"—")'
    for i in range(8):
        rr = rs + 1 + i
        ws.cell(row=rr, column=3, value=F("C", i, f3, f3b)).number_format = fmts[i]
        ws.cell(row=rr, column=4, value=F("D", i, f5, f5b)).number_format = fmts[i]
        c = ws.cell(row=rr, column=5, value=f'=IF({cA}="III",C{rr},IF({cA}="V",D{rr},"—"))'); c.number_format = fmts[i]; c.font = Font(bold=True)
        if i in (4, 6): c.fill = VERDE_NF; c.font = Font(bold=True, color="0B4F32")
        for col in range(1, 6): ws.cell(row=rr, column=col).border = BORDA
    ws.cell(row=rs + 5, column=6, value=f'=IF(AND(ISNUMBER({cF}),AND({cF}=5,C{rs+3}>0.05/0.335)),"Anexo III: teto do ISS aplicado","")')
    ws.cell(row=t3 - 1, column=1, value="Tabela – Anexo III (partilha 2027–2028)").font = Font(bold=True)
    _tabela_sim(ws, t3, "III")
    ws.cell(row=t5 - 1, column=1, value="Tabela – Anexo V (partilha 2027–2028)").font = Font(bold=True)
    _tabela_sim(ws, t5, "V")
    return t5 + 8, f"E{rs+3}", f"E{rs+5}", f"E{rs+7}", f"C{r7+1}"

def aba_simulador(wb, res):
    sim = res["simulador"]
    ws = wb.create_sheet("CBS e IBS – Nota fiscal")
    titulo(ws, 1, f"CBS E IBS PARA AS NOTAS FISCAIS DE {sim['alvo']} – {res['cliente']}", 14)
    ws.cell(row=2, column=1, value=(f"RBT12 de {sim['meses'][0]} a {sim['meses'][-1]} (inclui o mês fechado, {res['competencia']}). "
             f"Pela regra de 2027, este acumulado vale para as notas e o DAS de {sim['alvo']}.")).font = Font(italic=True)
    r = 3
    for o in sim["obs"]:
        ws.cell(row=r, column=1, value="• " + o).font = Font(italic=True, color="C00000"); r += 1
    r += 1
    titulo(ws, r, "RESUMO PARA PASSAR AO CLIENTE"); r += 1
    for i, h in enumerate(["Atividade", "Anexo", "Alíq. efetiva do DAS", "CBS (destaque NF)", "IBS (destaque NF)"], 1):
        x = ws.cell(row=r, column=i, value=h); x.font = Font(bold=True, color="FFFFFF"); x.fill = NAVY
    r_res = r + 1
    r = r_res + len(sim["anexos"]) + 2
    resumo = []
    for an in sim["anexos"]:
        if an == "FATOR_R":
            r, ce, cc, ci, ca = _bloco_fator_r(ws, r, sim["rbt12"], sim["folha12"])
            resumo.append(("Serviço com Fator R", f"={ca}", ce, cc, ci))
        else:
            r, ce, cc, ci = _bloco_anexo(ws, r, an, sim["rbt12"])
            nomes = {"I": "Comércio (Anexo I)", "II": "Indústria (Anexo II)", "III": "Serviços (Anexo III)",
                     "IV": "Serviços (Anexo IV)", "V": "Serviços (Anexo V)"}
            resumo.append((nomes[an], an, ce, cc, ci))
        r += 2
    for i, (ativ, an, ce, cc, ci) in enumerate(resumo):
        rr = r_res + i
        ws.cell(row=rr, column=1, value=ativ)
        ws.cell(row=rr, column=2, value=an)
        ws.cell(row=rr, column=3, value=f"={ce}").number_format = PCT4
        c = ws.cell(row=rr, column=4, value=f"={cc}"); c.number_format = PCT4; c.fill = VERDE_NF; c.font = Font(bold=True, color="0B4F32")
        c = ws.cell(row=rr, column=5, value=f"={ci}"); c.number_format = PCT4; c.fill = VERDE_NF; c.font = Font(bold=True, color="0B4F32")
        for col in range(1, 6): ws.cell(row=rr, column=col).border = BORDA
    ws.cell(row=r, column=1, value=("Notas: partilha 2027–2028 conforme LC 214/2025 (Anexos XVIII a XXII da LC 123). Destaque na nota = "
             "alíquota efetiva do DAS × % de CBS/IBS da partilha; CBS e IBS já estão dentro do DAS. Na 5ª faixa dos Anexos III e IV, "
             "teto de 5% do ISS com redistribuição. Percentuais sujeitos à regulamentação final do CGSN.")).font = Font(italic=True, color="666666")
    for col, w in zip("ABCDEFGHIJK", [44, 18, 16, 16, 16, 14, 12, 12, 12, 12, 12]):
        ws.column_dimensions[col].width = w

def planilha(res, path):
    wb = Workbook(); ws = wb.active; ws.title = "Apuração"
    titulo(ws, 1, f"APURAÇÃO DO SIMPLES NACIONAL – {res['competencia']}", 14)
    r = 3
    for k, v in [("Cliente", res["cliente"]), ("CNPJ", res["cnpj"]), ("UF", res["uf"]),
                 ("Atividade (CNAE)", res["atividade"]), ("Competência", res["competencia"])]:
        ws.cell(row=r, column=1, value=k).font = Font(bold=True); ws.cell(row=r, column=2, value=v); r += 1
    r += 1
    # 12 meses
    if res["meses_fat"]:
        titulo(ws, r, f"FATURAMENTO E FOLHA DA JANELA DA RBT12 ({res['janela'][0]} a {res['janela'][-1]})"); r += 1
        cab(ws, r, ["Mês", "Faturamento", "Folha c/ encargos"]); r += 1
        f0 = r
        folha_map = dict(res["meses_folha"])
        for m, v in res["meses_fat"]:
            ws.cell(row=r, column=1, value=m)
            ws.cell(row=r, column=2, value=v).number_format = MOEDA
            if folha_map: ws.cell(row=r, column=3, value=folha_map.get(m, 0)).number_format = MOEDA
            for c in range(1, 4): ws.cell(row=r, column=c).border = BORDA
            r += 1
        ws.cell(row=r, column=1, value="Total 12 meses").font = Font(bold=True)
        ws.cell(row=r, column=2, value=f"=SUM(B{f0}:B{r-1})").number_format = MOEDA
        if folha_map: ws.cell(row=r, column=3, value=f"=SUM(C{f0}:C{r-1})").number_format = MOEDA
        tot12 = r; r += 2
        rbt_src = f"=B{tot12}" if res["meses"] >= 12 else res["rbt12"]
        folha_src = f"=C{tot12}" if (folha_map and res["meses"] >= 12) else res["folha12"]
    else:
        ws.cell(row=r, column=1, value=f"RBT12 informada pelo total de {res['janela'][0]} a {res['janela'][-1]}"); r += 2
        rbt_src, folha_src = res["rbt12"], res["folha12"]
    # Memória SN
    titulo(ws, r, "MEMÓRIA DE CÁLCULO – SIMPLES NACIONAL"); r += 1
    lin_rbt = r
    ws.cell(row=r, column=1, value="RBT12"); c = ws.cell(row=r, column=2, value=rbt_src); c.number_format = MOEDA; c.fill = AMAR; r += 1
    ws.cell(row=r, column=1, value="Período da RBT12"); ws.cell(row=r, column=2, value=f"{res['janela'][0]} a {res['janela'][-1]}"); r += 1
    ws.cell(row=r, column=1, value="Faixa"); ws.cell(row=r, column=2, value=f"{res['faixa']}ª faixa"); r += 1
    ws.cell(row=r, column=1, value="Tabelas usadas"); ws.cell(row=r, column=2, value=res["base_tabelas"]); r += 1
    if res["usa_fr"]:
        lin_folha = r
        ws.cell(row=r, column=1, value="Folha 12 meses (Fator R)"); c = ws.cell(row=r, column=2, value=folha_src); c.number_format = MOEDA; c.fill = AMAR; r += 1
        lin_fr = r
        ws.cell(row=r, column=1, value="Fator R (folha ÷ RBT12)")
        ws.cell(row=r, column=2, value=f"=IF(B{lin_rbt}>0,B{lin_folha}/B{lin_rbt},0)").number_format = PCT; r += 1
        ws.cell(row=r, column=1, value="Anexo pelo Fator R (≥ 28% → III; < 28% → V)")
        ws.cell(row=r, column=2, value=f'=IF(B{lin_fr}>=0.28,"III","V")'); r += 1
    r += 1
    # Memória ICMS PR
    if res["uf"] == "PR" and any(l["icms_pr"] for l in res["linhas"]):
        titulo(ws, r, "MEMÓRIA DE CÁLCULO – ICMS/PR (Anexo XI do RICMS/PR)"); r += 1
        cab(ws, r, ["Tabela", "Situação", "Alíq. nominal PR", "Valor a deduzir PR", "ICMS efetivo PR",
                    "ICMS na repartição do SN", "% redução PGDAS-D"]); r += 1
        vistos = set()
        for l in res["linhas"]:
            if not l["icms_pr"] or l["anexo"] in vistos: continue
            vistos.add(l["anexo"])
            sit, efpr, red = l["icms_pr"]
            tab = "I – comércio" if l["anexo"] == "I" else "II – indústria"
            ws.cell(row=r, column=1, value=tab)
            ws.cell(row=r, column=2, value={"isento":"Isento (até R$ 360 mil)","reducao":"Redução","sem_tabela":"Sem tabela"}[sit])
            if sit == "reducao":
                nom, pd = ICMS_PR[l["anexo"]][res["faixa"]-1]
                ws.cell(row=r, column=3, value=nom).number_format = '0.0000%'
                ws.cell(row=r, column=4, value=pd).number_format = MOEDA
                ws.cell(row=r, column=5, value=f"=($B${lin_rbt}*C{r}-D{r})/$B${lin_rbt}").number_format = PCT4
                ws.cell(row=r, column=6, value=l["aliq"]["ICMS"]).number_format = PCT4
                ws.cell(row=r, column=7, value=f"=ROUND((1-E{r}/F{r})*100,2)/100").number_format = PCT
            else:
                ws.cell(row=r, column=6, value=l["aliq"]["ICMS"]).number_format = PCT4
                ws.cell(row=r, column=7, value=1 if sit == "isento" else 0).number_format = PCT
            for c in range(1, 8): ws.cell(row=r, column=c).border = BORDA
            r += 1
        r += 1
    # Cálculo por receita
    titulo(ws, r, "IMPOSTO DO MÊS POR RECEITA"); r += 1
    cols = ["Receita", "Anexo", "Como chegou no anexo", "Faixa", "Alíq. nominal", "Parcela a deduzir",
            "Alíq. efetiva SN", "Segregação / ICMS", "Alíq. final", "Faturamento", "DAS"]
    cab(ws, r, cols); r += 1
    first = r
    lin_receita = []
    for l in res["linhas"]:
        ws.cell(row=r, column=1, value=l["descricao"])
        ws.cell(row=r, column=2, value=l["anexo"])
        ws.cell(row=r, column=3, value=l["origem"])
        ws.cell(row=r, column=4, value=f"{l['faixa']}ª")
        ws.cell(row=r, column=5, value=l["nominal"]).number_format = PCT
        ws.cell(row=r, column=6, value=l["pd"]).number_format = MOEDA
        ws.cell(row=r, column=7, value=f"=($B${lin_rbt}*E{r}-F{r})/$B${lin_rbt}").number_format = PCT4
        seg = [NOME_TRAT.get(t, t) for t in l["tratamentos"]]
        if l["zerados"]: seg.append("sem " + ", ".join(l["zerados"]))
        if l["obs_icms"]: seg.append(l["obs_icms"])
        ws.cell(row=r, column=8, value="; ".join(seg) or "—").alignment = Alignment(wrap_text=True, vertical="top")
        ws.cell(row=r, column=9, value=l["aliq_final"]).number_format = PCT4
        ws.cell(row=r, column=10, value=l["receita"]).number_format = MOEDA
        lin_receita.append(r)
        for c in range(1, 12): ws.cell(row=r, column=c).border = BORDA
        r += 1
    last = r - 1
    ws.cell(row=r, column=1, value="TOTAL").font = Font(bold=True)
    ws.cell(row=r, column=10, value=f"=SUM(J{first}:J{last})").number_format = MOEDA
    c = ws.cell(row=r, column=11, value=f"=SUM(K{first}:K{last})"); c.number_format = MOEDA; c.font = Font(bold=True)
    for cc in range(1, 12): ws.cell(row=r, column=cc).fill = CINZA; ws.cell(row=r, column=cc).border = BORDA
    tot = r; r += 2
    titulo(ws, r, "VALOR TOTAL DO DAS")
    c = ws.cell(row=r, column=2, value=f"=K{tot}"); c.number_format = MOEDA; c.font = Font(bold=True, size=12); c.fill = VERDE
    lin_das = r; r += 1
    ws.cell(row=r, column=1, value="Alíquota final sobre o faturamento do mês").font = Font(bold=True)
    ws.cell(row=r, column=2, value=f"=IF(J{tot}>0,K{tot}/J{tot},0)").number_format = PCT4; r += 1
    if res["vencimento"]:
        ws.cell(row=r, column=1, value="Vencimento do DAS").font = Font(bold=True)
        c = ws.cell(row=r, column=2, value=res["vencimento"]); c.font = Font(bold=True, color="C00000"); c.fill = AMAR
        ws.cell(row=r, column=3, value=res["vencimento_obs"]).font = Font(italic=True); r += 1
    if res["valor_ecac"] is not None:
        ws.cell(row=r, column=1, value="Valor no e-CAC / PGDAS-D")
        ws.cell(row=r, column=2, value=float(res["valor_ecac"])).number_format = MOEDA; r += 1
        ws.cell(row=r, column=1, value="Diferença (calculado − e-CAC)")
        ws.cell(row=r, column=2, value=f"=B{lin_das}-B{r-1}").number_format = MOEDA; r += 1
    r += 1
    if res["alertas"]:
        titulo(ws, r, "OBSERVAÇÕES"); r += 1
        for a in res["alertas"]:
            ws.cell(row=r, column=1, value="• " + a); r += 1
        r += 1
    ws.cell(row=r, column=1, value="Método: alíquota efetiva = (RBT12 × nominal − parcela a deduzir) ÷ RBT12, sem arredondar; "
            "cada tributo = faturamento × efetiva × % repartição (ICMS com isenção/redução PR), arredondado no centavo; DAS = soma dos tributos.").font = Font(italic=True, color="666666")
    for i, w in enumerate([34, 16, 26, 7, 11, 14, 13, 46, 11, 15, 14], 1): ws.column_dimensions[get_column_letter(i)].width = w

    # Repartição
    w2 = wb.create_sheet("Repartição")
    titulo(w2, 1, "REPARTIÇÃO DO DAS POR TRIBUTO", 14)
    rr = 3
    for idx, l in enumerate(res["linhas"]):
        lr = lin_receita[idx]
        w2.cell(row=rr, column=1, value=f"{l['descricao']} – Anexo {l['anexo']}, {l['faixa']}ª faixa").font = Font(bold=True); rr += 1
        cab(w2, rr, ["Tributo", "% repartição", "Alíq. do tributo", "Redução / exclusão", "Alíq. final", "Valor"]); rr += 1
        f0 = rr
        for t in TRIB:
            if l["perc"][t] == 0: continue
            red = 1.0 if t in l["zerados"] else (l["red_icms"] if t == "ICMS" else 0.0)
            w2.cell(row=rr, column=1, value=t)
            w2.cell(row=rr, column=2, value=l["perc"][t]).number_format = PCT
            w2.cell(row=rr, column=3, value=f"=Apuração!G{lr}*B{rr}").number_format = PCT4
            w2.cell(row=rr, column=4, value=red).number_format = PCT
            w2.cell(row=rr, column=5, value=f"=C{rr}*(1-D{rr})").number_format = PCT4
            w2.cell(row=rr, column=6, value=f"=ROUND(Apuração!J{lr}*E{rr},2)").number_format = MOEDA
            for c in range(1, 7): w2.cell(row=rr, column=c).border = BORDA
            rr += 1
        w2.cell(row=rr, column=1, value="Total").font = Font(bold=True)
        w2.cell(row=rr, column=6, value=f"=SUM(F{f0}:F{rr-1})").number_format = MOEDA
        ws.cell(row=lr, column=11, value=f"=Repartição!F{rr}").number_format = MOEDA
        if l["iss_nota"]:
            rr += 1
            w2.cell(row=rr, column=1, value="Alíquota de ISS para a nota")
            w2.cell(row=rr, column=3, value=l["iss_nota"]).number_format = PCT
        if res["ano"] >= 2027 and (l["cbs_nota"] or l["ibs_nota"]):
            rr += 1
            w2.cell(row=rr, column=1, value="Destaque na nota fiscal (informativo, já dentro do DAS)").font = Font(bold=True)
            rr += 1
            w2.cell(row=rr, column=1, value="CBS"); w2.cell(row=rr, column=3, value=l["cbs_nota"]).number_format = PCT4
            rr += 1
            w2.cell(row=rr, column=1, value="IBS"); w2.cell(row=rr, column=3, value=l["ibs_nota"]).number_format = PCT4
        rr += 2
    for i, w in enumerate([12, 13, 15, 16, 13, 14], 1): w2.column_dimensions[get_column_letter(i)].width = w

    # Tabelas
    w3 = wb.create_sheet("Tabelas")
    titulo(w3, 1, f"TABELAS DO SIMPLES NACIONAL – {res['base_tabelas']}")
    cab(w3, 3, ["Faixa", "RBT12 até", "Anexo I", "PD I", "Anexo II", "PD II", "Anexo III", "PD III",
                "Anexo IV", "PD IV", "Anexo V", "PD V"])
    for i in range(6):
        row = 4 + i
        w3.cell(row=row, column=1, value=f"{i+1}ª"); w3.cell(row=row, column=2, value=FAIXAS[i]).number_format = MOEDA
        for j, an in enumerate(["I","II","III","IV","V"]):
            nom, pd = NOMINAL[an][i]
            w3.cell(row=row, column=3+2*j, value=nom).number_format = PCT
            w3.cell(row=row, column=4+2*j, value=pd).number_format = MOEDA
    titulo(w3, 12, "ICMS/PR – Anexo XI do RICMS/PR (Decreto 8.660/2018)")
    cab(w3, 13, ["Faixa", "RBT12 até", "Comércio (Tab. I)", "Deduzir", "Indústria (Tab. II)", "Deduzir"])
    for i in range(5):
        row = 14 + i
        w3.cell(row=row, column=1, value=f"{i+1}ª"); w3.cell(row=row, column=2, value=FAIXAS[i]).number_format = MOEDA
        if i <= 1:
            w3.cell(row=row, column=3, value="Isento"); w3.cell(row=row, column=5, value="Isento")
        else:
            for j, an in enumerate(["I","II"]):
                nom, pd = ICMS_PR[an][i]
                w3.cell(row=row, column=3+2*j, value=nom).number_format = '0.0000%'
                w3.cell(row=row, column=4+2*j, value=pd).number_format = MOEDA
    for i in range(1, 13): w3.column_dimensions[get_column_letter(i)].width = 14
    aba_simulador(wb, res)
    wb.move_sheet("CBS e IBS – Nota fiscal", offset=-(len(wb.sheetnames) - 2))   # logo depois da aba Apuração
    wb.save(path)

if __name__ == "__main__":
    d = json.load(open(sys.argv[1], encoding="utf-8"))
    res = calcular(d)
    planilha(res, sys.argv[2])
    out = {k: res[k] for k in ("cliente","competencia","janela","base_tabelas","vencimento","simulador","uf","rbt12","folha12","fator_r","faixa","receita_mes","das_total","aliquota_final","alertas")}
    out["linhas"] = [{k: l[k] for k in ("descricao","anexo","origem","faixa","nominal","pd","efetiva","red_icms","obs_icms","aliq_final","receita","das","partes","iss_nota","cbs_nota","ibs_nota")} for l in res["linhas"]]
    print(json.dumps(out, ensure_ascii=False, indent=1))
```

## Script `ler_extrato_pgdas.py`

```python
#!/usr/bin/env python3
"""Lê o extrato do PGDAS-D (PDF da Receita Federal) e devolve, em JSON, o faturamento mês a mês.

Uso:  python3 ler_extrato_pgdas.py extrato.pdf [outro.pdf ...]

Para cada extrato:
- cnpj_basico, nome, pa (período de apuração "MM/AAAA");
- rpa: receita bruta do próprio PA (mercado interno + externo);
- rbt12_extrato: RBT12 informada no extrato (para conferência);
- meses: {"MM/AAAA": receita} das "Receitas Brutas Anteriores" (interno + externo) + o próprio PA;
- folha: {"MM/AAAA": valor} das "Folhas de Salários Anteriores", quando houver;
- confere: true se a soma dos meses da RBT12 bate com a RBT12 do extrato
  (até 2026: 12 meses anteriores ao PA; a partir de 2027: 12 meses antes do mês anterior ao PA).
"""
import json, re, subprocess, sys

VAL = r"(\d{1,3}(?:\.\d{3})*,\d{2})"
PAR = re.compile(r"(\d{2}/\d{4})\s+" + VAL)

def num(txt):
    return float(txt.replace(".", "").replace(",", "."))

def texto(pdf):
    return subprocess.run(["pdftotext", "-layout", pdf, "-"], capture_output=True, text=True, check=True).stdout

def secao(t, inicio, fins):
    i = re.search(inicio, t, re.I)
    if not i: return ""
    resto = t[i.end():]
    corte = len(resto)
    for f in fins:
        j = re.search(f, resto, re.I)
        if j: corte = min(corte, j.start())
    return resto[:corte]

def pares(bloco):
    out = {}
    for m, v in PAR.findall(bloco):
        out[m] = out.get(m, 0.0) + num(v)
    return out

def mes_anterior(pa, n):
    m, a = int(pa[:2]), int(pa[3:])
    t = a*12 + m - 1 - n
    return f"{t % 12 + 1:02d}/{t // 12}"

def ler(pdf):
    t = texto(pdf)
    r = {"arquivo": pdf}
    # extrato: "Período de Apuração (PA): 10/2026"; declaração: "Período de Apuração: 01/10/2026 a 31/10/2026"
    m = re.search(r"Per[ií]odo de Apura[cç][aã]o(?: \(PA\))?:?\s*(?:\d{2}/)?(\d{2}/\d{4})", t, re.I)
    r["pa"] = m.group(1) if m else None
    m = re.search(r"CNPJ Matriz:?\s*([\d\./-]{18})", t)
    r["cnpj"] = m.group(1) if m else None
    m = re.search(r"CNPJ Estabelecimento:?\s*([\d\./-]{18})", t)
    if m and not r["cnpj"]: r["cnpj"] = m.group(1)
    m = re.search(r"UF:\s*([A-Z]{2})", t); r["uf"] = m.group(1) if m else None
    m = re.search(r"CNPJ B[aá]sico:?\s*([\d\.]{8,11})", t, re.I)
    r["cnpj_basico"] = re.sub(r"\D", "", m.group(1)) if m else (re.sub(r"\D", "", r["cnpj"])[:8] if r["cnpj"] else None)
    m = re.search(r"Nome [Ee]mpresarial:?\s*(.+)", t)
    r["nome"] = m.group(1).strip() if m else None
    # Receita do PA: primeira linha "Receita Bruta do PA (RPA)" -> pega o último valor da linha (Total)
    m = re.search(r"Receita Bruta do PA \(RPA\)[^\n]*", t, re.I)
    if m:
        vals = re.findall(VAL, m.group(0)); r["rpa"] = num(vals[-1]) if vals else None
    # a linha da RBT12 quebra no PDF: os valores ficam na linha "...doze meses anteriores (ao PA)"
    m = re.search(r"doze meses anteriores(?: ao PA)?\s{2,}[^\n]*", t)
    if m:
        vals = re.findall(VAL, m.group(0)); r["rbt12_extrato"] = num(vals[-1]) if vals else None
    ant = secao(t, r"Receitas? Brutas? Anteriores", [r"Folhas? de Sal[aá]rios? Anteriores", r"2\.4\)", r"Fator r", r"\n\s*3\)"])
    meses = pares(ant)
    folha_bloco = secao(t, r"Folhas? de Sal[aá]rios? Anteriores", [r"2\.4\)", r"Fator r", r"\n\s*3\)", r"Valor Original"])
    r["folha"] = pares(folha_bloco)
    if r.get("pa") and r.get("rpa") is not None:
        meses[r["pa"]] = r["rpa"]
    r["meses"] = dict(sorted(meses.items(), key=lambda kv: (kv[0][3:], kv[0][:2])))
    if r.get("pa") and r.get("rbt12_extrato") is not None:
        ini = 2 if int(r["pa"][3:]) >= 2027 else 1
        soma = sum(meses.get(mes_anterior(r["pa"], k), 0.0) for k in range(ini, ini + 12))
        r["soma_12_anteriores"] = round(soma, 2)
        r["confere"] = abs(soma - r["rbt12_extrato"]) < 0.05
    return r

if __name__ == "__main__":
    print(json.dumps([ler(p) for p in sys.argv[1:]], ensure_ascii=False, indent=1))
```