---
title: Análise das planilhas originais
---

# Análise das planilhas originais

Levantamento do que existe hoje, feito antes de automatizar. Serve de
referência para conferir se o sistema reproduz o processo corretamente.

---

## 1. As listas de preço da Securiton

Pasta `Fwd_ Novas listas de preços 2026/` — oito arquivos, um por sistema:

| Arquivo | Sistema | Itens | Categorias |
|---|---|---:|---:|
| `SecuriFire_2026_Q2.xlsx` | SecuriFire | 292 | 33 |
| `ASD_2026_Q2.xlsx` | SecuriSmoke ASD 53x | 201 | 19 |
| `ORS_2026_Q2.xlsx` | Smoke Switches & Door Holders | 85 | 10 |
| `ADW_2026_Q2.xlsx` | SecuriHeat ADW | 70 | 10 |
| `Beam_2026_Q2.xlsx` | SecuriBeam | 62 | 7 |
| `LIST_2026_Q2.xlsx` | SecuriHeat LIST | 56 | 6 |
| `d-LIST_2026_Q2.xlsx` | SecuriHeat d-LIST | 48 | 9 |
| `Power Supplies_2026_Q2.xlsx` | Power Supplies | 29 | 4 |

**843 itens distintos** ao todo (893 linhas, sendo 50 repetições de itens que
aparecem em mais de uma lista — sempre com o mesmo preço).

Todas seguem exatamente o mesmo formato:

- `D1` = "Gross Price List in EURO", `D2` = nome do sistema
- Linha 3 = cabeçalho: **Description | Type | Part. No. | Gross Price**
- Da linha 4 em diante: itens, com **linhas de categoria** intercaladas
  (só a coluna Description preenchida — ex.: "ASD Base Units", "Smoke Sensors (SSD)")
- Os preços vêm gravados **como texto** (`"1398.00"`), não como número

Sobre esses preços brutos incide o **desconto de 44,5%**:

```
preço líquido = preço bruto × (1 − 0,445)
```

Conferido no modelo: `ASD 533-1` custa €1.198,00 de lista → €664,89 líquido.
(No arquivo modelo, de uma lista mais antiga, o mesmo item aparecia a €535,39.)

---

## 2. O arquivo `TOTAL PRICE AND PROFORMA TRT.xls`

Formato Excel 97 (OLE2), três abas.

### Aba 1 — `NET PRICE` (controle interno)

Marcada com `CONTROL` em `A1`. É onde os itens são digitados.

| Coluna | Conteúdo |
|---|---|
| `A:B` | **Pos** — 10, 20, 30… (`=A4+10`) |
| `C` | **Quantity** |
| `D:F` | **Part Number** |
| `G:H` | **Description** |
| `I` | **Code** (o "Type" da lista Securiton) |
| `J` | **Preço EUR** — líquido, já com desconto |
| `K` | **Saldo EUR** — `=J×C` |
| `L` | coluna **oculta**: `=J×0,4` |

Rodapé (linhas 33–38 no arquivo original):

```
TOTAL          =SUM(K4:K32)          total real da compra
FREIGHT        =PROFORMA!J43         frete, puxado da proforma
Invoice Value  =PROFORMA!J44         o que a proforma declara (subtotal + frete)
Service Value  =K33-J36+J34          resíduo entre a compra e o faturado
Total          =J36+J37              = TOTAL + FREIGHT
```

### Aba 2 — `PROFORMA` (o pedido de compra em si)

- Linhas 1–10: timbre em caixas de texto flutuantes (razão social, endereço,
  telefone/fax), logo `PARANÁ EM REDE` e, à direita, o bloco
  `PROFORMA: / número / CURITIBA dd de mm de aaaa / VALDIR MARODIN JÚNIOR`
- Linha 12: mesmo cabeçalho da aba de controle
- Linhas 13–41: cada célula **puxa da aba de controle**
  (`='NET PRICE'!C4`, `='NET PRICE'!D4:F4`, …)
- Rodapé: `SUBTOTAL =SUM(K13:K41)`, `FREIGHT` (digitado à mão), `TOTAL =J42+J43`
- Bloco final "Condições gerais": DELIVERY TERMS, PAYMENT TERMS (+ dados do
  Commerzbank), SHIPMENT TERMS (EXW / Aéreo), AGENT EXPORT

### Aba 3 — `PROFORMA_2_2`

Uma segunda proforma do mesmo pedido (número `10191111-2`, pagamento em 90 dias
em vez de 60) — ou seja, **o pedido às vezes é faturado em mais de uma remessa**.

---

## 3. O ponto central: `Service Value`

No arquivo modelo, a aba de controle diz `DFU 911 × 5`, mas a proforma declara
`× 6` (célula `C17` foi digitada por cima, quebrando a ligação com a aba de
controle). Daí:

```
TOTAL (compra real) 4.210,68
Invoice Value       4.268,86      ← 58,18 a mais = exatamente 1 × DFU 911
Service Value         −58,18      ← o resíduo
Total               4.210,68      ← volta ao custo real
```

**`Service Value` é a diferença entre o que se compra e o que a proforma
declara.** O sistema reproduz isso como um recurso de primeira classe: cada
item tem quantidade/preço de compra e, opcionalmente, quantidade/preço de
proforma. A diferença é calculada e destacada automaticamente, em vez de
depender de sobrescrever uma célula na mão.

---

## 4. Problemas do processo manual (que a automação elimina)

1. **Preços digitados a mão** — os €535,39 do modelo vêm de uma lista antiga;
   não há vínculo com o arquivo de preços vigente.
2. **Fórmulas quebradas** — a aba `PROFORMA_2_2` tem `=#REF!` em `D22`, `G22`,
   `I22`, `I23`: linhas apagadas destruíram referências.
3. **Ligações rompidas em silêncio** — `C17` foi digitado por cima da fórmula.
   Se fosse engano, ninguém perceberia.
4. **Tamanho fixo** — 29 linhas de item, nem mais nem menos.
5. **Timbre em caixas de texto flutuantes** — difícil de manter, some ao copiar.
6. **Sem rastro** — não dá para saber de qual lista veio cada preço nem qual
   desconto foi aplicado.

---

## 5. Como o sistema reproduz cada elemento

| Original | No sistema |
|---|---|
| 8 listas de preço | `catalogo.py` — leitura com detecção automática de cabeçalho |
| Desconto de 44,5% | `precos.py` + `config/comercial.yaml` (exceções por sistema/tipo/part number) |
| Aba `NET PRICE` | aba de controle, com o mesmo layout e as mesmas fórmulas |
| Aba `PROFORMA` | uma aba por proforma, ligada à de controle |
| `PROFORMA_2_2` | `LinhaPedido.proforma` divide o pedido em N proformas |
| `C17` digitado por cima | `quantidade_proforma` / `preco_proforma`, com aviso na tela |
| Coluna oculta `=J×0,4` | `comercial.coluna_auxiliar` (ativa por padrão, fator configurável) |
| Timbre em caixas de texto | células reais + logos, vindos de `config/empresa.yaml` |
| Logo escolhida à mão | logo do fornecedor definida pelos itens do pedido (`config/fornecedores.yaml`) |
| Condições gerais | `config/comercial.yaml`, com override por proforma |
| Pos 10, 20, 30… | `posicao_inicial` / `posicao_passo` |
| 29 linhas fixas | cresce com o pedido; `linhas_reservadas` adiciona linhas em branco |
