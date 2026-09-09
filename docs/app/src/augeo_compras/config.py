"""Carrega os arquivos YAML de `config/` em objetos tipados.

Toda a parametrização do sistema passa por aqui: dados da empresa, regras
comerciais, origem das listas de preço e layout da planilha. Nenhum outro
módulo lê YAML diretamente.
"""

from __future__ import annotations

import copy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

import yaml

# Raiz do projeto = duas pastas acima deste arquivo (src/augeo_compras/).
RAIZ = Path(__file__).resolve().parents[2]
PASTA_CONFIG = RAIZ / "config"


def _ler_yaml(caminho: Path) -> dict[str, Any]:
    """Lê um YAML de configuração.

    Não existindo, recorre ao `*.exemplo.yaml` de mesmo nome — é o que faz um
    clone novo do repositório funcionar mesmo sem os arquivos que ficam de fora
    do Git por trazerem informação comercial.
    """
    if not caminho.exists():
        exemplo = caminho.parent / f"{caminho.stem}.exemplo.yaml"
        if not exemplo.exists():
            return {}
        caminho = exemplo
    with caminho.open(encoding="utf-8") as arq:
        return yaml.safe_load(arq) or {}


def resolver(caminho: str | Path) -> Path:
    """Resolve um caminho relativo a partir da raiz do projeto."""
    p = Path(caminho).expanduser()
    return p if p.is_absolute() else (RAIZ / p)


@dataclass
class Empresa:
    """Um emitente da proforma.

    A mesma operação pode faturar por mais de uma pessoa jurídica — o
    `identificador` distingue cada CNPJ, e a tela de identificação escolhe qual
    está em uso.
    """

    identificador: str = ""
    razao_social: str = ""
    cnpj: str = ""
    endereco: list[str] = field(default_factory=list)
    telefone: str = ""
    fax: str = ""
    email: str = ""
    responsavel: str = ""
    logo: str = ""
    largura_logo: int = 110
    cidade: str = ""

    @property
    def caminho_logo(self) -> Path | None:
        if not self.logo:
            return None
        p = resolver(self.logo)
        return p if p.exists() else None

    def linhas_timbre(self) -> list[str]:
        """Bloco de texto do topo da proforma, na ordem de exibição."""
        linhas = [self.razao_social]
        if self.cnpj:
            linhas.append(f"CNPJ: {self.cnpj}")
        linhas.extend(self.endereco)
        if self.telefone:
            linhas.append(f"Phone: {self.telefone}")
        if self.fax:
            linhas.append(f"Fax: {self.fax}")
        if self.email:
            linhas.append(self.email)
        return [linha for linha in linhas if linha]

    def para_dict(self) -> dict[str, Any]:
        return {
            "identificador": self.identificador,
            "razao_social": self.razao_social,
            "cnpj": self.cnpj,
            "endereco": list(self.endereco),
            "telefone": self.telefone,
            "fax": self.fax,
            "email": self.email,
            "responsavel": self.responsavel,
            "logo": self.logo,
            "cidade": self.cidade,
        }

    @classmethod
    def de_dict(cls, dados: dict[str, Any]) -> "Empresa":
        return cls(
            identificador=str(dados.get("identificador") or dados.get("id") or ""),
            razao_social=dados.get("razao_social", "") or "",
            cnpj=dados.get("cnpj", "") or "",
            endereco=[linha for linha in (dados.get("endereco") or []) if linha],
            telefone=dados.get("telefone", "") or "",
            fax=dados.get("fax", "") or "",
            email=dados.get("email", "") or "",
            responsavel=dados.get("responsavel", "") or "",
            logo=dados.get("logo", "") or "",
            largura_logo=int(dados.get("largura_logo", 110)),
            cidade=dados.get("cidade", "") or "",
        )


@dataclass
class ColunaAuxiliar:
    ativa: bool = False
    rotulo: str = "Aux"
    fator: float = 0.0


@dataclass
class Condicoes:
    """Textos do bloco "Condições gerais" da proforma."""

    delivery_terms: str = ""
    payment_terms: str = ""
    banco: list[str] = field(default_factory=list)
    shipment_terms: str = ""
    modal: str = ""
    agent_export: str = ""

    @classmethod
    def de_dict(cls, dados: dict[str, Any] | None) -> "Condicoes":
        dados = dados or {}
        return cls(
            delivery_terms=dados.get("delivery_terms", "") or "",
            payment_terms=dados.get("payment_terms", "") or "",
            banco=list(dados.get("banco") or []),
            shipment_terms=dados.get("shipment_terms", "") or "",
            modal=dados.get("modal", "") or "",
            agent_export=dados.get("agent_export", "") or "",
        )

    def para_dict(self) -> dict[str, Any]:
        return {
            "delivery_terms": self.delivery_terms,
            "payment_terms": self.payment_terms,
            "banco": list(self.banco),
            "shipment_terms": self.shipment_terms,
            "modal": self.modal,
            "agent_export": self.agent_export,
        }

    def mesclar(self, outras: "Condicoes | None") -> "Condicoes":
        """Devolve uma cópia com os campos preenchidos de `outras` aplicados."""
        if outras is None:
            return copy.deepcopy(self)
        base = self.para_dict()
        for chave, valor in outras.para_dict().items():
            if valor:
                base[chave] = valor
        return Condicoes.de_dict(base)


@dataclass
class Comercial:
    desconto_percentual: float = 0.0
    descontos_por_sistema: dict[str, float] = field(default_factory=dict)
    descontos_por_tipo: dict[str, float] = field(default_factory=dict)
    descontos_por_part_number: dict[str, float] = field(default_factory=dict)
    casas_decimais: int = 2
    moeda: str = "EUR"
    simbolo_moeda: str = "€"
    posicao_inicial: int = 10
    posicao_passo: int = 10
    coluna_auxiliar: ColunaAuxiliar = field(default_factory=ColunaAuxiliar)
    linhas_reservadas: int = 0
    condicoes_padrao: Condicoes = field(default_factory=Condicoes)


@dataclass
class Fornecedor:
    """Fabricante de quem se compra — dá nome e logo à proforma."""

    id: str
    nome: str = ""
    logo: str = ""
    largura_logo: int = 190

    @property
    def caminho_logo(self) -> Path | None:
        if not self.logo:
            return None
        p = resolver(self.logo)
        return p if p.exists() else None


@dataclass
class Fornecedores:
    itens: dict[str, Fornecedor] = field(default_factory=dict)
    padrao: str = ""
    pedido_misto: str = "dominante"

    def get(self, identificador: str | None) -> Fornecedor | None:
        if identificador and identificador in self.itens:
            return self.itens[identificador]
        if self.padrao:
            return self.itens.get(self.padrao)
        return None


@dataclass
class FonteListas:
    """Uma pasta (ou arquivo) de listas e o fornecedor a que pertence."""

    caminho: str
    fornecedor: str = ""


@dataclass
class OrigemCatalogo:
    pastas: list[FonteListas] = field(default_factory=list)
    arquivos: list[FonteListas] = field(default_factory=list)
    ignorar: list[str] = field(default_factory=list)
    colunas: dict[str, list[str]] = field(default_factory=dict)
    sobrescritas: dict[str, dict[str, Any]] = field(default_factory=dict)
    sistemas: dict[str, str] = field(default_factory=dict)
    categorias_por_linha_sem_preco: bool = True


@dataclass
class Coluna:
    """Uma coluna da tabela de itens da planilha."""

    chave: str
    rotulo: str
    span: int = 1
    larguras: list[float] = field(default_factory=list)
    alinhamento: str = "center"
    formato: str = "General"
    fonte: str | None = None
    fonte_tamanho: float | None = None
    negrito: bool = False
    quebra_linha: bool = False


@dataclass
class Layout:
    abas: dict[str, str] = field(default_factory=dict)
    colunas: list[Coluna] = field(default_factory=list)
    estilo: dict[str, Any] = field(default_factory=dict)
    totais: dict[str, str] = field(default_factory=dict)
    rodape: dict[str, str] = field(default_factory=dict)
    pagina: dict[str, Any] = field(default_factory=dict)

    @property
    def total_colunas(self) -> int:
        return sum(c.span for c in self.colunas)

    def coluna(self, chave: str) -> Coluna | None:
        return next((c for c in self.colunas if c.chave == chave), None)


@dataclass
class Config:
    empresa: Empresa                      # o emitente em uso
    comercial: Comercial
    catalogo: OrigemCatalogo
    layout: Layout
    fornecedores: Fornecedores = field(default_factory=Fornecedores)
    emitentes: list[Empresa] = field(default_factory=list)
    raiz: Path = RAIZ

    def emitente(self, identificador: str) -> Empresa | None:
        return next((e for e in self.emitentes if e.identificador == identificador), None)

    @property
    def pasta_saida(self) -> Path:
        return self.raiz / "data" / "saida"

    @property
    def pasta_logos(self) -> Path:
        """Onde ficam as logos enviadas pela tela de identificação."""
        return self.raiz / "data" / "logos"

    @property
    def pasta_pedidos(self) -> Path:
        return self.raiz / "data" / "pedidos"


def carregar(pasta: Path | str | None = None) -> Config:
    """Lê `config/*.yaml` e devolve a configuração completa do sistema."""
    pasta = Path(pasta) if pasta else PASTA_CONFIG

    dados_empresa = _ler_yaml(pasta / "empresa.yaml")
    if dados_empresa.get("emitentes"):
        emitentes = [Empresa.de_dict(e) for e in dados_empresa["emitentes"]]
    else:
        # Formato antigo, com um emitente só solto na raiz do arquivo.
        emitentes = [Empresa.de_dict({**dados_empresa, "identificador": "padrao"})]
    escolhido = dados_empresa.get("padrao") or (emitentes[0].identificador if emitentes else "")
    empresa = next(
        (e for e in emitentes if e.identificador == escolhido),
        emitentes[0] if emitentes else Empresa(),
    )

    dc = _ler_yaml(pasta / "comercial.yaml")
    aux = dc.get("coluna_auxiliar") or {}
    comercial = Comercial(
        desconto_percentual=float(dc.get("desconto_percentual", 0) or 0),
        descontos_por_sistema={k: float(v) for k, v in (dc.get("descontos_por_sistema") or {}).items()},
        descontos_por_tipo={k: float(v) for k, v in (dc.get("descontos_por_tipo") or {}).items()},
        descontos_por_part_number={
            str(k): float(v) for k, v in (dc.get("descontos_por_part_number") or {}).items()
        },
        casas_decimais=int(dc.get("casas_decimais", 2)),
        moeda=dc.get("moeda", "EUR"),
        simbolo_moeda=dc.get("simbolo_moeda", "€"),
        posicao_inicial=int(dc.get("posicao_inicial", 10)),
        posicao_passo=int(dc.get("posicao_passo", 10)),
        coluna_auxiliar=ColunaAuxiliar(
            ativa=bool(aux.get("ativa", False)),
            rotulo=aux.get("rotulo", "Aux"),
            fator=float(aux.get("fator", 0) or 0),
        ),
        linhas_reservadas=int(dc.get("linhas_reservadas", 0)),
        condicoes_padrao=Condicoes.de_dict(dc.get("condicoes_padrao")),
    )

    dcat = _ler_yaml(pasta / "catalogo.yaml")

    def _fontes(valores: Any) -> list[FonteListas]:
        """Aceita tanto `- pasta/x` quanto `- {caminho: ..., fornecedor: ...}`."""
        fontes = []
        for valor in valores or []:
            if isinstance(valor, dict):
                fontes.append(
                    FonteListas(
                        caminho=str(valor.get("caminho", "")),
                        fornecedor=str(valor.get("fornecedor", "") or ""),
                    )
                )
            elif valor:
                fontes.append(FonteListas(caminho=str(valor)))
        return [f for f in fontes if f.caminho]

    catalogo = OrigemCatalogo(
        pastas=_fontes(dcat.get("pastas")),
        arquivos=_fontes(dcat.get("arquivos")),
        ignorar=list(dcat.get("ignorar") or []),
        colunas={k: list(v) for k, v in (dcat.get("colunas") or {}).items()},
        sobrescritas=dict(dcat.get("sobrescritas") or {}),
        sistemas=dict(dcat.get("sistemas") or {}),
        categorias_por_linha_sem_preco=bool(dcat.get("categorias_por_linha_sem_preco", True)),
    )

    dl = _ler_yaml(pasta / "layout.yaml")
    colunas = [
        Coluna(
            chave=c["chave"],
            rotulo=c.get("rotulo", c["chave"]),
            span=int(c.get("span", 1)),
            larguras=[float(x) for x in (c.get("larguras") or [])],
            alinhamento=c.get("alinhamento", "center"),
            formato=c.get("formato", "General"),
            fonte=c.get("fonte"),
            fonte_tamanho=c.get("fonte_tamanho"),
            negrito=bool(c.get("negrito", False)),
            quebra_linha=bool(c.get("quebra_linha", False)),
        )
        for c in (dl.get("colunas") or [])
    ]
    layout = Layout(
        abas=dict(dl.get("abas") or {"controle": "NET PRICE", "proforma": "PROFORMA"}),
        colunas=colunas,
        estilo=dict(dl.get("estilo") or {}),
        totais=dict(dl.get("totais") or {}),
        rodape=dict(dl.get("rodape") or {}),
        pagina=dict(dl.get("pagina") or {}),
    )

    df = _ler_yaml(pasta / "fornecedores.yaml")
    fornecedores = Fornecedores(
        itens={
            str(f["id"]): Fornecedor(
                id=str(f["id"]),
                nome=f.get("nome", "") or "",
                logo=f.get("logo", "") or "",
                largura_logo=int(f.get("largura_logo", 190)),
            )
            for f in (df.get("fornecedores") or [])
            if f.get("id")
        },
        padrao=str(df.get("padrao", "") or ""),
        pedido_misto=str(df.get("pedido_misto", "dominante") or "dominante"),
    )

    return Config(
        empresa=empresa,
        comercial=comercial,
        catalogo=catalogo,
        layout=layout,
        fornecedores=fornecedores,
        emitentes=emitentes,
    )
