"""Modelo do pedido de compra e sua resolução contra o catálogo.

Um `Pedido` guarda só o essencial (part number + quantidade + ajustes). Ao ser
resolvido contra o catálogo e a tabela de preços, vira um `PedidoResolvido`,
que já tem descrição, preço líquido e totais — é o que o gerador de planilha
consome.

Um pedido pode ser dividido em várias proformas (o modelo original tinha
PROFORMA e PROFORMA_2_2). Cada linha indica em qual proforma entra.
"""

from __future__ import annotations

import json
from dataclasses import asdict, dataclass, field
from datetime import date
from pathlib import Path
from typing import Any, Iterable

from .catalogo import Catalogo, Item, normalizar_part_number
from .config import Comercial, Condicoes
from .precos import TabelaDePrecos

MESES_PT = [
    "janeiro", "fevereiro", "março", "abril", "maio", "junho",
    "julho", "agosto", "setembro", "outubro", "novembro", "dezembro",
]


class PedidoInvalido(ValueError):
    """Erro de conteúdo do pedido (item inexistente, quantidade inválida...)."""


def _para_data(valor: Any) -> date:
    if isinstance(valor, date):
        return valor
    if not valor:
        return date.today()
    return date.fromisoformat(str(valor)[:10])


def data_por_extenso(quando: date, cidade: str = "") -> str:
    texto = f"{quando.day} de {MESES_PT[quando.month - 1]} de {quando.year}"
    return f"{cidade}, {texto}" if cidade else texto


# --------------------------------------------------------------------------- #
# Entrada
# --------------------------------------------------------------------------- #
@dataclass
class LinhaPedido:
    """Um item do pedido.

    `quantidade` é o que se compra de fato (aba de controle). Quando a proforma
    precisa declarar quantidade ou preço diferentes, use os campos
    `quantidade_proforma` / `preco_proforma`: a diferença aparece
    automaticamente como "Service Value" na aba de controle.
    """

    part_number: str
    quantidade: float = 1
    proforma: int = 1
    descricao: str | None = None
    codigo: str | None = None
    preco_unitario: float | None = None
    desconto: float | None = None
    quantidade_proforma: float | None = None
    preco_proforma: float | None = None
    observacao: str = ""

    def para_dict(self) -> dict[str, Any]:
        obrigatorios = {"part_number", "quantidade", "proforma"}
        return {
            chave: valor
            for chave, valor in asdict(self).items()
            if chave in obrigatorios or valor not in (None, "")
        }

    @classmethod
    def de_dict(cls, dados: dict[str, Any]) -> "LinhaPedido":
        return cls(
            part_number=normalizar_part_number(dados.get("part_number") or dados.get("codigo_peca") or ""),
            quantidade=float(dados.get("quantidade", 1) or 0),
            proforma=int(dados.get("proforma", 1) or 1),
            descricao=dados.get("descricao"),
            codigo=dados.get("codigo"),
            preco_unitario=_opcional_float(dados.get("preco_unitario")),
            desconto=_opcional_float(dados.get("desconto")),
            quantidade_proforma=_opcional_float(dados.get("quantidade_proforma")),
            preco_proforma=_opcional_float(dados.get("preco_proforma")),
            observacao=dados.get("observacao", "") or "",
        )


def _opcional_float(valor: Any) -> float | None:
    if valor is None or valor == "":
        return None
    return float(valor)


@dataclass
class Proforma:
    """Cabeçalho de uma proforma do pedido."""

    numero: str = ""
    data: date = field(default_factory=date.today)
    frete: float = 0.0
    condicoes: Condicoes | None = None
    titulo_aba: str = ""

    def para_dict(self) -> dict[str, Any]:
        return {
            "numero": self.numero,
            "data": self.data.isoformat(),
            "frete": self.frete,
            "condicoes": self.condicoes.para_dict() if self.condicoes else None,
            "titulo_aba": self.titulo_aba,
        }

    @classmethod
    def de_dict(cls, dados: dict[str, Any]) -> "Proforma":
        return cls(
            numero=str(dados.get("numero", "") or ""),
            data=_para_data(dados.get("data")),
            frete=float(dados.get("frete", 0) or 0),
            condicoes=Condicoes.de_dict(dados["condicoes"]) if dados.get("condicoes") else None,
            titulo_aba=str(dados.get("titulo_aba", "") or ""),
        )


@dataclass
class Pedido:
    referencia: str = ""
    cliente: str = ""
    data: date = field(default_factory=date.today)
    proformas: list[Proforma] = field(default_factory=lambda: [Proforma()])
    linhas: list[LinhaPedido] = field(default_factory=list)
    observacoes: str = ""

    # -- persistência ------------------------------------------------------ #
    def para_dict(self) -> dict[str, Any]:
        return {
            "referencia": self.referencia,
            "cliente": self.cliente,
            "data": self.data.isoformat(),
            "observacoes": self.observacoes,
            "proformas": [p.para_dict() for p in self.proformas],
            "linhas": [linha.para_dict() for linha in self.linhas],
        }

    @classmethod
    def de_dict(cls, dados: dict[str, Any]) -> "Pedido":
        proformas = [Proforma.de_dict(p) for p in (dados.get("proformas") or [])] or [Proforma()]
        return cls(
            referencia=str(dados.get("referencia", "") or ""),
            cliente=str(dados.get("cliente", "") or ""),
            data=_para_data(dados.get("data")),
            proformas=proformas,
            linhas=[LinhaPedido.de_dict(linha) for linha in (dados.get("linhas") or [])],
            observacoes=str(dados.get("observacoes", "") or ""),
        )

    def salvar(self, caminho: Path | str) -> Path:
        caminho = Path(caminho)
        caminho.parent.mkdir(parents=True, exist_ok=True)
        caminho.write_text(
            json.dumps(self.para_dict(), ensure_ascii=False, indent=2), encoding="utf-8"
        )
        return caminho

    @classmethod
    def carregar(cls, caminho: Path | str) -> "Pedido":
        return cls.de_dict(json.loads(Path(caminho).read_text(encoding="utf-8")))

    # -- importação de planilha simples ------------------------------------ #
    @classmethod
    def de_planilha(cls, caminho: Path | str, **cabecalho: Any) -> "Pedido":
        """Cria um pedido a partir de uma planilha/CSV com part number + qtd.

        Aceita qualquer arquivo com uma coluna de part number e uma de
        quantidade; o cabeçalho é detectado pelo nome da coluna.
        """
        from .catalogo import _linhas_da_planilha, _normalizar, limpar_texto

        _, linhas = _linhas_da_planilha(Path(caminho))
        alias_pn = ("part", "codigo", "código", "item", "sap")
        alias_qt = ("qtd", "quant", "qty", "quantidade", "menge")
        alias_pf = ("proforma", "remessa")

        idx_pn = idx_qt = idx_pf = None
        inicio = 0
        for numero, linha in enumerate(linhas[:20]):
            mapa = {}
            for col, celula in enumerate(linha):
                texto = _normalizar(celula)
                if not texto:
                    continue
                if idx_pn is None and any(texto.startswith(a) for a in alias_pn):
                    mapa["pn"] = col
                elif idx_qt is None and any(texto.startswith(a) for a in alias_qt):
                    mapa["qt"] = col
                elif any(texto.startswith(a) for a in alias_pf):
                    mapa["pf"] = col
            if "pn" in mapa and "qt" in mapa:
                idx_pn, idx_qt, idx_pf = mapa["pn"], mapa["qt"], mapa.get("pf")
                inicio = numero + 1
                break

        if idx_pn is None or idx_qt is None:
            raise PedidoInvalido(
                "A planilha precisa de uma coluna de part number e uma de quantidade."
            )

        itens: list[LinhaPedido] = []
        for linha in linhas[inicio:]:
            def valor(i: int | None) -> Any:
                return linha[i] if i is not None and i < len(linha) else None

            pn = normalizar_part_number(valor(idx_pn))
            if not pn:
                continue
            try:
                qtd = float(str(limpar_texto(valor(idx_qt))).replace(",", ".") or 0)
            except ValueError:
                continue
            if qtd <= 0:
                continue
            proforma = 1
            if idx_pf is not None:
                try:
                    proforma = int(float(valor(idx_pf) or 1))
                except (TypeError, ValueError):
                    proforma = 1
            itens.append(LinhaPedido(part_number=pn, quantidade=qtd, proforma=proforma))

        pedido = cls(linhas=itens)
        for chave, valor_campo in cabecalho.items():
            if hasattr(pedido, chave) and valor_campo is not None:
                setattr(pedido, chave, valor_campo)
        return pedido


# --------------------------------------------------------------------------- #
# Resolução contra o catálogo
# --------------------------------------------------------------------------- #
@dataclass
class LinhaResolvida:
    posicao: int
    part_number: str
    tipo: str
    descricao: str
    sistema: str
    categoria: str
    quantidade: float
    preco_bruto: float
    desconto_percentual: float
    preco_unitario: float
    quantidade_proforma: float
    preco_proforma: float
    proforma: int
    observacao: str = ""
    fornecedor: str = ""

    @property
    def total(self) -> float:
        return round(self.quantidade * self.preco_unitario, 2)

    @property
    def total_proforma(self) -> float:
        return round(self.quantidade_proforma * self.preco_proforma, 2)

    @property
    def divergente(self) -> bool:
        """A proforma declara algo diferente do que está sendo comprado."""
        return (
            abs(self.quantidade - self.quantidade_proforma) > 1e-9
            or abs(self.preco_unitario - self.preco_proforma) > 1e-9
        )


@dataclass
class ProformaResolvida:
    indice: int
    numero: str
    data: date
    frete: float
    condicoes: Condicoes
    titulo_aba: str
    linhas: list[LinhaResolvida] = field(default_factory=list)

    @property
    def subtotal(self) -> float:
        return round(sum(linha.total_proforma for linha in self.linhas), 2)

    @property
    def total(self) -> float:
        return round(self.subtotal + self.frete, 2)


@dataclass
class PedidoResolvido:
    pedido: Pedido
    linhas: list[LinhaResolvida]
    proformas: list[ProformaResolvida]
    avisos: list[str] = field(default_factory=list)

    @property
    def total_controle(self) -> float:
        """Total real da compra (aba de controle)."""
        return round(sum(linha.total for linha in self.linhas), 2)

    @property
    def frete_total(self) -> float:
        return round(sum(p.frete for p in self.proformas), 2)

    @property
    def valor_faturado(self) -> float:
        """Soma das proformas — o que é declarado na importação."""
        return round(sum(p.total for p in self.proformas), 2)

    @property
    def valor_servico(self) -> float:
        """Resíduo entre a compra real e o que foi faturado na proforma."""
        return round(self.total_controle + self.frete_total - self.valor_faturado, 2)

    @property
    def total_geral(self) -> float:
        return round(self.valor_faturado + self.valor_servico, 2)

    def contagem_por_fornecedor(self) -> dict[str, int]:
        """Quantas linhas vieram de cada fornecedor, da maior para a menor."""
        contagem: dict[str, int] = {}
        for linha in self.linhas:
            if linha.fornecedor:
                contagem[linha.fornecedor] = contagem.get(linha.fornecedor, 0) + 1
        return dict(sorted(contagem.items(), key=lambda par: -par[1]))

    @property
    def fornecedor_dominante(self) -> str:
        """O fornecedor com mais itens no pedido — define a logo da proforma."""
        contagem = self.contagem_por_fornecedor()
        return next(iter(contagem), "")

    @property
    def fornecedores_misturados(self) -> bool:
        return len(self.contagem_por_fornecedor()) > 1


def resolver(
    pedido: Pedido,
    catalogo: Catalogo,
    comercial: Comercial,
) -> PedidoResolvido:
    """Cruza o pedido com o catálogo e calcula preços, posições e totais."""
    tabela = TabelaDePrecos(comercial)
    avisos: list[str] = []
    linhas: list[LinhaResolvida] = []

    posicao = comercial.posicao_inicial
    for entrada in pedido.linhas:
        if entrada.quantidade is None or entrada.quantidade < 0:
            raise PedidoInvalido(f"{entrada.part_number}: quantidade inválida ({entrada.quantidade})")

        item = catalogo.resolver(entrada.part_number)
        if item is None:
            if entrada.preco_unitario is None or not entrada.descricao:
                raise PedidoInvalido(
                    f"Part number {entrada.part_number!r} não existe nas listas de preço. "
                    "Para incluí-lo mesmo assim, informe descrição e preço unitário na linha."
                )
            # Item avulso (serviço, frete interno, item fora de lista).
            item = Item(
                part_number=entrada.part_number,
                tipo=entrada.codigo or "",
                descricao=entrada.descricao,
                preco_bruto=entrada.preco_unitario,
                sistema="Fora de lista",
                origem="pedido",
            )
            avisos.append(f"{entrada.part_number}: item fora das listas, usando preço informado")

        if entrada.preco_unitario is not None:
            preco = tabela.arredondar(entrada.preco_unitario)
            desconto = entrada.desconto if entrada.desconto is not None else 0.0
        else:
            calculo = tabela.calcular(item, desconto=entrada.desconto)
            preco = calculo.preco_liquido
            desconto = calculo.desconto_percentual

        qtd_proforma = (
            entrada.quantidade_proforma if entrada.quantidade_proforma is not None else entrada.quantidade
        )
        preco_proforma = (
            tabela.arredondar(entrada.preco_proforma) if entrada.preco_proforma is not None else preco
        )

        linhas.append(
            LinhaResolvida(
                posicao=posicao,
                part_number=item.part_number,
                tipo=entrada.codigo or item.tipo,
                descricao=entrada.descricao or item.descricao,
                sistema=item.sistema,
                categoria=item.categoria,
                quantidade=float(entrada.quantidade),
                preco_bruto=item.preco_bruto,
                desconto_percentual=desconto,
                preco_unitario=preco,
                quantidade_proforma=float(qtd_proforma),
                preco_proforma=preco_proforma,
                proforma=max(1, int(entrada.proforma or 1)),
                observacao=entrada.observacao,
                fornecedor=item.fornecedor,
            )
        )
        posicao += comercial.posicao_passo

    # Garante uma proforma para cada índice referenciado pelas linhas.
    quantidade_proformas = max([linha.proforma for linha in linhas] + [len(pedido.proformas), 1])
    proformas: list[ProformaResolvida] = []
    for indice in range(1, quantidade_proformas + 1):
        base = pedido.proformas[indice - 1] if indice <= len(pedido.proformas) else Proforma()
        proformas.append(
            ProformaResolvida(
                indice=indice,
                numero=base.numero,
                data=base.data or pedido.data,
                frete=float(base.frete or 0),
                condicoes=comercial.condicoes_padrao.mesclar(base.condicoes),
                titulo_aba=base.titulo_aba,
                linhas=[linha for linha in linhas if linha.proforma == indice],
            )
        )

    for proforma in proformas:
        if not proforma.linhas:
            avisos.append(f"Proforma {proforma.indice} está sem itens")

    return PedidoResolvido(pedido=pedido, linhas=linhas, proformas=proformas, avisos=avisos)


def de_itens(
    itens: Iterable[tuple[str, float]],
    **cabecalho: Any,
) -> Pedido:
    """Atalho: cria um pedido a partir de pares (part number, quantidade)."""
    pedido = Pedido(linhas=[LinhaPedido(part_number=pn, quantidade=q) for pn, q in itens])
    for chave, valor in cabecalho.items():
        if hasattr(pedido, chave) and valor is not None:
            setattr(pedido, chave, valor)
    return pedido
