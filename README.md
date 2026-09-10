# Automação das planilhas de compra — Securiton

Gera as planilhas **TOTAL PRICE AND PROFORMA** a partir das listas de preço da
Securiton: você escolhe os itens e as quantidades, o sistema aplica o desconto,
monta a aba de controle e a(s) proforma(s) com todas as fórmulas ligadas.

### ▶ Usar agora, sem instalar nada

**<https://pmarodinn.github.io/automatizacao_imposto/>**

A mesma interface, rodando **inteiramente dentro do navegador**: o código Python
é o mesmo, compilado para WebAssembly. Abra o endereço, carregue suas listas de
preço, informe o desconto e gere a planilha. Nada é enviado para servidor
nenhum — as listas e os pedidos ficam guardados na sua própria máquina.

> Não sabe o que as planilhas originais faziam? Está tudo mapeado em
> [`ANALISE_PLANILHAS.md`](ANALISE_PLANILHAS.md).

---

## Entrada

Dois caminhos:

| | Login mestre | Cadastro |
|---|---|---|
| Quem | `valdirPR` | qualquer e-mail |
| Onde fica | no código (hash SHA-256) | Firebase Auth |
| Validade | não expira | 7 dias |

Quem não tem acesso clica em **Criar cadastro**, informa e-mail e senha, e entra
na hora — com prazo de uma semana. O prazo fica no Firestore
(`usuarios/{uid}.expira_em`) e as regras do banco **recusam alterá-lo**, então
não adianta mexer pelo navegador.

O banco guarda **só o cadastro**. Pedidos, rascunhos e listas de preço continuam
apenas no navegador de quem usa.

O login mestre é uma tranca, não segurança — o código de um site estático é
público. Para trocá-lo, gere o hash de `usuario:senha:augeo-compras` e edite
`MESTRE` em `web/static/index.html`.

### Firebase

Projeto `cadastro-impautomatizado`. Os arquivos `firebase.json`,
`.firebaserc`, `firestore.rules` e `firestore.indexes.json` estão no
repositório. Para publicar as regras:

```bash
firebase deploy --only firestore:rules
```

Para dar mais prazo a alguém, edite `expira_em` na ficha dele pelo console do
Firebase.

## Identificação: quem compra, e de quem

No topo da tela ficam as duas partes do negócio — à esquerda o **emitente**
(nome, CNPJ e logo de quem faz o pedido), à direita o **fornecedor**. Clicar ali
abre a tela de identificação, que também é a primeira coisa que aparece no
primeiro uso.

Dá para cadastrar **quantas empresas quiser** — botão *+ Nova empresa*, e o ✕ no
canto do chip remove. Paraná em Rede e Augeo Engenharia vêm em
`config/empresa.yaml` só como ponto de partida. O emitente escolhido é o que sai
no timbre.

O **CNPJ é opcional** e não vem de lugar nenhum automaticamente: as listas de
preço não têm essa informação, e o modelo original não trazia CNPJ no timbre.

As **duas logos podem ser enviadas pela tela** — emitente à esquerda da
proforma, fornecedor à direita — e valem na hora.

## Os dois modos de uso

A interface é uma só; o que muda é onde o Python roda.

| | Site publicado | Servidor local |
|---|---|---|
| Como abrir | um endereço no navegador | `augeo-compras web` |
| Onde o Python roda | no navegador (WebAssembly) | na sua máquina |
| Listas de preço | você carrega os arquivos; ficam no navegador | lidas da pasta em `catalogo.yaml` |
| Desconto | informado na tela | vem de `config/comercial.yaml` |
| Pedidos salvos | guardados no navegador | em `data/pedidos/` |
| Precisa instalar | nada | Python + `pip install -e .` |

A tela decide sozinha: havendo servidor local, fala com ele; não havendo, sobe o
Python dentro do navegador. Os dois usam exatamente o mesmo código de catálogo,
preços e geração de planilha.

Para republicar o site depois de mexer no sistema:

```bash
python3 ferramentas/gerar_site.py
```

## Instalação do servidor local (uma vez só)

```bash
python3 -m venv .venv && .venv/bin/pip install -e .
```

Depois, dois arquivos que **não ficam no repositório** (veja
[O que fica fora do Git](#o-que-fica-fora-do-git)):

```bash
cp config/comercial.exemplo.yaml config/comercial.yaml
```

Edite-o com o desconto negociado e os dados bancários, e copie as listas de
preço do fabricante para a pasta apontada em `config/catalogo.yaml`.

## Uso no dia a dia

```bash
.venv/bin/augeo-compras web
```

Abra <http://127.0.0.1:8000> no navegador.

No **primeiro acesso** aparece uma página inicial curta explicando o que o
sistema faz, os quatro passos de uso e o significado de cada conceito
(as duas abas, preço líquido, `Service Value`, a logo automática). De lá dá para
começar com um **guia passo a passo**: pop-ups que destacam cada parte da tela e
explicam para que serve. Depois disso o sistema abre direto — o guia continua
disponível no botão **? Como usar**, no topo.

A tela tem duas colunas:

- **esquerda** — busca no catálogo (843 itens, 8 sistemas). Digite parte da
  descrição, do tipo ou do part number e clique no item para adicioná-lo.
- **direita** — o pedido: referência, proformas, itens, condições gerais e os
  totais, recalculados a cada alteração.

O botão **Gerar planilha** baixa o `.xlsx` pronto. Uma cópia fica em
`data/saida/`.

---

## O que a planilha gerada contém

**Aba `NET PRICE`** (controle interno) — um item por linha, com posição
(10, 20, 30…), part number, descrição, código, preço líquido e total. No pé:

| Linha | Significado |
|---|---|
| `TOTAL` | total real da compra |
| `FREIGHT` | frete, somado das proformas |
| `Invoice Value` | o que as proformas declaram |
| `Service Value` | a diferença entre os dois |
| `Total` | `Invoice + Service` |

**Aba `PROFORMA`** (uma por proforma) — timbre com as duas logos, número e data,
a tabela de itens ligada por fórmula à aba de controle, subtotal/frete/total e
o bloco "Condições gerais".

As fórmulas ficam vivas: alterar uma quantidade na aba de controle atualiza a
proforma e todos os totais, direto no Excel.

### A logo sai sozinha, conforme o que está sendo comprado

A logo de quem emite a proforma fica sempre no canto superior esquerdo. A do
**fornecedor** vai no canto direito e é escolhida pelos próprios itens do
pedido: cada item sabe de qual pasta de listas veio, a pasta sabe de qual
fornecedor é, e o fornecedor traz a sua logo. Comprando produtos Securiton, sai
a logo da Securiton — sem precisar escolher nada. A interface mostra, ao lado do
botão de gerar, qual logo será usada.

Num pedido que misture fabricantes, vale a do fornecedor com mais itens (e um
aviso aparece na tela). Para omitir a logo nesse caso, troque `pedido_misto`
para `nenhuma` em `config/fornecedores.yaml`.

### Quando a proforma declara algo diferente da compra

É o caso do `Service Value`. Clique no ⚙ da linha e preencha **Qtd. na
proforma** ou **Preço na proforma**. O sistema:

- grava o valor próprio naquela célula da proforma (desligando a ligação só ali);
- calcula a diferença como `Service Value` na aba de controle;
- marca a linha com o selo *proforma difere* e mostra um aviso.

Deixando em branco, a proforma declara exatamente o que está sendo comprado.

### Dividir em várias proformas

**+ Dividir em outra proforma** cria mais uma. Cada item escolhe em qual entra
(⚙ → Proforma), e cada uma tem seu número, data, frete e *payment terms*
próprios. Sai uma aba para cada, e a aba de controle soma todas.

---

## Configuração

Tudo que muda com o tempo está em `config/`, em arquivos de texto comentados —
não é preciso mexer em código.

| Arquivo | Para quê |
|---|---|
| `empresa.yaml` | os emitentes: razão social, CNPJ, endereço, responsável e logo de cada um |
| `comercial.yaml` | **desconto**, moeda, numeração das posições, condições padrão (fora do Git) |
| `catalogo.yaml` | onde ficam as listas de preço e como lê-las |
| `fornecedores.yaml` | nome e logo de cada fabricante |
| `layout.yaml` | colunas da planilha, rótulos, larguras, fontes, página |

### Trocar as listas de preço (a cada trimestre)

Copie os arquivos novos para a pasta indicada em `catalogo.yaml` (por padrão
`Fwd_ Novas listas de preços 2026/`) e clique em **↻ Recarregar listas**. Não
precisa configurar nada: o leitor acha o cabeçalho sozinho, contanto que a lista
tenha as colunas Description / Type / Part. No. / Gross Price — em qualquer
ordem, com ou sem acento, em `.xlsx`, `.xls` ou `.csv`.

Se a pasta mudar de nome, ajuste `pastas:` em `catalogo.yaml`.

### Acrescentar outro fabricante

Dois passos. Em `config/fornecedores.yaml`:

```yaml
fornecedores:
  - id: securiton
    nome: Securiton AG
    logo: logo/SECURITON.jpg
  - id: outro
    nome: Outro Fabricante S.A.
    logo: logo/outro.png
```

E em `config/catalogo.yaml`, aponte a pasta das listas dele:

```yaml
pastas:
  - caminho: "Fwd_ Novas listas de preços 2026"
    fornecedor: securiton
  - caminho: "Listas Outro Fabricante"
    fornecedor: outro
```

Pronto: os itens dele entram no catálogo e a proforma passa a sair com a logo
certa quando o pedido for dele.

### Mudar o desconto

Em `config/comercial.yaml`:

```yaml
desconto_percentual: 44.5

# Exceções, do mais específico para o mais geral:
descontos_por_sistema:
  SecuriFire: 40.0
descontos_por_tipo:
  "ASD 535-3": 50.0
descontos_por_part_number:
  "11-2000001-01-04": 35.0
```

Vence sempre a regra mais específica: **part number > tipo > sistema > padrão**.

---

## O que fica fora do Git

Este repositório é público e contém **só o sistema**. Ficam de fora, por
`.gitignore`, os arquivos que são material de terceiros ou informação
comercial — eles vivem apenas na máquina de quem usa:

| Fora do repositório | Por quê |
|---|---|
| As listas de preço do fabricante | Material confidencial da Securiton |
| Proformas e planilhas `.xls`/`.xlsx` | Trazem cliente, quantidades e preços praticados |
| `config/comercial.yaml` | Desconto negociado e dados bancários |

No lugar do último vai `config/comercial.exemplo.yaml`, com a mesma estrutura e
valores neutros. **Se `comercial.yaml` não existir, o sistema usa o de exemplo**
— um clone novo sobe e funciona, só sem o desconto real (0%). Os testes que
dependem das listas se pulam sozinhos quando elas não estão presentes.

## Linha de comando

Para uso avançado ou automação, a CLI faz tudo o que a interface faz:

```bash
augeo-compras catalogo --exportar     # catálogo com bruto, desconto e líquido
augeo-compras buscar asd 535          # procura itens
augeo-compras modelo pedido.json      # cria um pedido de exemplo
augeo-compras gerar pedido.json       # gera a planilha
augeo-compras importar lista.xlsx --gerar   # part number + quantidade → planilha
```

`importar` aceita qualquer planilha com uma coluna de part number e uma de
quantidade (opcionalmente uma de proforma) — útil quando a lista de compras
chega pronta de outro lugar.

### Formato do pedido (`.json`)

```json
{
  "referencia": "TRT",
  "cliente": "Securiton AG",
  "data": "2026-08-29",
  "proformas": [
    { "numero": "2608261-1", "data": "2026-08-29", "frete": 0 }
  ],
  "linhas": [
    { "part_number": "11-2000001-01-04", "quantidade": 2 },
    { "part_number": "11-2300030-01-02", "quantidade": 5, "quantidade_proforma": 6 }
  ]
}
```

Campos opcionais por linha: `proforma`, `preco_unitario`, `desconto`,
`quantidade_proforma`, `preco_proforma`, `descricao`, `codigo`, `observacao`.

Para incluir algo que não está nas listas (um serviço, por exemplo), informe
`descricao` e `preco_unitario` na linha — o preço entra sem desconto.

---

## Organização do código

```
config/                 configuração em YAML (é aqui que se ajusta o sistema)
logo/                   logotipos (o da empresa e o de cada fornecedor)
data/pedidos/           pedidos salvos (.json)
data/saida/             planilhas geradas
docs/                   site publicado (gerado por ferramentas/gerar_site.py)
ferramentas/            scripts de apoio
src/augeo_compras/
  config.py             lê os YAML em objetos tipados
  catalogo.py           lê as listas de preço da Securiton
  precos.py             aplica o desconto (com as exceções)
  pedido.py             modelo do pedido e cálculo dos totais
  planilha.py           escreve o .xlsx (controle + proformas)
  estilos.py            fontes, bordas e alinhamentos
  servico.py            junta tudo — usado pela CLI, pela web e pelo navegador
  navegador.py          ponte para quando o Python roda dentro do navegador
  cli.py                linha de comando
  web/                  API FastAPI + interface (página única, sem build)
tests/                  55 testes
```

O fluxo é sempre o mesmo, e cada etapa é testável isoladamente:

```
listas .xlsx ──▶ Catálogo ──▶ TabelaDePrecos ──▶ PedidoResolvido ──▶ .xlsx
                                                        │
                                          totais, Service Value, avisos
```

## Testes

```bash
.venv/bin/python -m pytest
```

Cobrem a leitura das listas reais, as regras de desconto, os totais (incluindo
`Service Value` e a divisão em várias proformas) e a estrutura da planilha. O
último teste reabre a planilha no LibreOffice e confere que as fórmulas
recalculam exatamente os valores esperados (pulado se o LibreOffice não estiver
instalado).
