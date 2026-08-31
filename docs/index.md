---
title: Automação de compras — Securiton
---

Sistema que monta a planilha **TOTAL PRICE AND PROFORMA** a partir das listas de
preço do fornecedor: você escolhe os itens, ele aplica o desconto negociado e
gera o arquivo do Excel com as duas abas já ligadas por fórmula.

> Esta página é a documentação do projeto. O sistema em si é um aplicativo
> Python que roda na máquina de quem usa — **não roda aqui no GitHub Pages**,
> que só serve páginas estáticas. Para usá-lo, siga a instalação abaixo.

- [Código no GitHub](https://github.com/pmarodinn/automatizacao_imposto)
- [Análise das planilhas originais](ANALISE_PLANILHAS.md) — o que o processo
  manual fazia, campo a campo, e como cada parte foi automatizada

---

## Como funciona

Quatro passos, na interface web:

1. **Busque os itens** — o catálogo reúne todas as listas do fabricante (843
   itens em 8 sistemas, na versão 2026 Q2). Procure por descrição, tipo ou
   part number e clique para adicionar.
2. **Ajuste o pedido** — quantidades, número da proforma, data e frete. A
   compra pode ser dividida em várias proformas.
3. **Confira os totais** — recalculados a cada mudança, com aviso quando a
   proforma declara algo diferente do que está sendo comprado.
4. **Gere a planilha** — sai o `.xlsx` pronto para enviar.

## O que sai

**Aba `NET PRICE`** — o controle interno: um item por linha, com posição
(10, 20, 30…), part number, descrição, código, preço líquido e total. No pé,
o bloco de conciliação:

| Linha | Significado |
|---|---|
| `TOTAL` | total real da compra |
| `FREIGHT` | frete, somado das proformas |
| `Invoice Value` | o que as proformas declaram |
| `Service Value` | a diferença entre os dois |
| `Total` | `Invoice + Service` |

**Aba `PROFORMA`** (uma por proforma) — timbre com a logo de quem emite e a do
fornecedor, número e data, a tabela de itens ligada por fórmula à aba de
controle, subtotal/frete/total e as condições gerais.

As fórmulas ficam vivas: mudar uma quantidade na aba de controle atualiza a
proforma e todos os totais, direto no Excel.

## Os dois conceitos que não são óbvios

**Service Value.** É a diferença entre o que está sendo comprado e o que a
proforma declara. Comprando 5 peças e declarando 6, a diferença aparece
sozinha ali — e o total geral continua sendo o custo real da compra. Na
planilha antiga isso era feito digitando por cima de uma célula com fórmula;
aqui é um campo próprio, com aviso na tela.

**A logo do fornecedor.** Sai escolhida pelos itens do pedido: cada item sabe
de qual pasta de listas veio, a pasta sabe de qual fornecedor é, e o fornecedor
traz a sua logo. Comprando produtos Securiton, a proforma sai com a logo da
Securiton, sem escolher nada.

## Instalação

```bash
git clone https://github.com/pmarodinn/automatizacao_imposto.git
cd automatizacao_imposto
python3 -m venv .venv && .venv/bin/pip install -e .
cp config/comercial.exemplo.yaml config/comercial.yaml
```

Preencha `config/comercial.yaml` com o desconto negociado e os dados bancários,
copie as listas de preço para a pasta apontada em `config/catalogo.yaml`, e:

```bash
.venv/bin/augeo-compras web
```

A interface abre em `http://127.0.0.1:8000`. No primeiro acesso ela mostra uma
página de apresentação e oferece um guia passo a passo pela tela.

## Configuração

Tudo que muda com o tempo está em `config/`, em YAML comentado:

| Arquivo | Para quê |
|---|---|
| `empresa.yaml` | razão social, endereço, telefone, responsável, logo |
| `comercial.yaml` | desconto, moeda, numeração, condições padrão |
| `catalogo.yaml` | onde ficam as listas e a qual fornecedor pertencem |
| `fornecedores.yaml` | nome e logo de cada fabricante |
| `layout.yaml` | colunas da planilha, rótulos, larguras, fontes, página |

Listas novas a cada trimestre: copie os arquivos para a pasta e clique em
**Recarregar listas**. O leitor encontra o cabeçalho sozinho, desde que a lista
tenha as colunas Description / Type / Part. No. / Gross Price — em qualquer
ordem, em `.xlsx`, `.xls` ou `.csv`.

## Dados fora do repositório

O repositório contém só o sistema. As listas de preço do fabricante, as
planilhas com pedidos reais e o `config/comercial.yaml` (desconto e dados
bancários) ficam de fora por `.gitignore` — vivem apenas na máquina de quem
usa. Não existindo `comercial.yaml`, o sistema usa
`config/comercial.exemplo.yaml`, então um clone novo sobe e funciona, só sem o
desconto real.
