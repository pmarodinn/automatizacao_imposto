"""Camada de serviço: junta configuração, catálogo, pedido e planilha.

É a porta de entrada única usada tanto pela linha de comando quanto pela
interface web, para que as duas se comportem exatamente da mesma forma.
"""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
from functools import lru_cache
from pathlib import Path
from typing import Any

from . import catalogo as mod_catalogo
from . import pedido as mod_pedido
from . import planilha as mod_planilha
from .catalogo import Catalogo, Item
from .config import Config, carregar as carregar_config
from .pedido import Pedido, PedidoResolvido
from .precos import PrecoCalculado, TabelaDePrecos


@dataclass
class ItemComPreco:
    """Item do catálogo já com o preço líquido calculado — para exibição."""

    item: Item
    preco: PrecoCalculado

    def para_dict(self) -> dict[str, Any]:
        return {
            **self.item.para_dict(),
            "desconto_percentual": self.preco.desconto_percentual,
            "preco_liquido": self.preco.preco_liquido,
            "fonte_desconto": self.preco.fonte_desconto,
        }


class Servico:
    """Ponto único de acesso ao sistema."""

    def __init__(self, config: Config | None = None):
        self.config = config or carregar_config()
        self._catalogo: Catalogo | None = None

    # -- catálogo ---------------------------------------------------------- #
    @property
    def catalogo(self) -> Catalogo:
        if self._catalogo is None:
            self._catalogo = mod_catalogo.carregar(self.config)
        return self._catalogo

    def recarregar_catalogo(self) -> Catalogo:
        self._catalogo = mod_catalogo.carregar(self.config)
        return self._catalogo

    @property
    def tabela(self) -> TabelaDePrecos:
        return TabelaDePrecos(self.config.comercial)

    def com_preco(self, item: Item) -> ItemComPreco:
        return ItemComPreco(item=item, preco=self.tabela.calcular(item))

    def buscar(
        self,
        termo: str = "",
        *,
        sistema: str | None = None,
        categoria: str | None = None,
        limite: int | None = 200,
    ) -> list[ItemComPreco]:
        encontrados = self.catalogo.buscar(
            termo, sistema=sistema, categoria=categoria, limite=limite
        )
        return [self.com_preco(item) for item in encontrados]

    def exportar_catalogo(self, destino: Path | str) -> Path:
        """Gera um Excel com todo o catálogo: preço bruto, desconto e líquido."""
        from openpyxl import Workbook
        from openpyxl.styles import Alignment, Font

        wb = Workbook()
        ws = wb.active
        ws.title = "Catálogo"
        cabecalho = [
            "Sistema", "Categoria", "Part Number", "Type", "Description",
            f"Bruto {self.config.comercial.moeda}", "Desconto %",
            f"Líquido {self.config.comercial.moeda}", "Arquivo",
        ]
        ws.append(cabecalho)
        for coluna, _ in enumerate(cabecalho, start=1):
            celula = ws.cell(row=1, column=coluna)
            celula.font = Font(bold=True)
            celula.alignment = Alignment(horizontal="center")

        tabela = self.tabela
        for item in self.catalogo:
            calculo = tabela.calcular(item)
            ws.append([
                item.sistema, item.categoria, item.part_number, item.tipo, item.descricao,
                item.preco_bruto, calculo.desconto_percentual, calculo.preco_liquido, item.origem,
            ])

        for coluna, largura in zip("ABCDEFGHI", (22, 30, 20, 20, 70, 14, 11, 14, 24)):
            ws.column_dimensions[coluna].width = largura
        for linha in range(2, ws.max_row + 1):
            for coluna in ("F", "H"):
                ws[f"{coluna}{linha}"].number_format = "#,##0.00"
            ws[f"G{linha}"].number_format = "0.0"
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = f"A1:I{ws.max_row}"

        destino = Path(destino)
        destino.parent.mkdir(parents=True, exist_ok=True)
        wb.save(destino)
        return destino

    # -- pedido ------------------------------------------------------------ #
    def resolver(self, pedido: Pedido) -> PedidoResolvido:
        return mod_pedido.resolver(pedido, self.catalogo, self.config.comercial)

    def gerar(
        self,
        pedido: Pedido,
        destino: Path | str | None = None,
    ) -> tuple[Path, PedidoResolvido]:
        """Resolve o pedido e grava a planilha. Devolve (caminho, resolvido)."""
        resolvido = self.resolver(pedido)
        if destino is None:
            destino = self.config.pasta_saida / mod_planilha.nome_sugerido(resolvido)
        caminho = mod_planilha.gerar(self.config, resolvido, destino)
        return caminho, resolvido

    def pedido_de_planilha(self, caminho: Path | str, **cabecalho: Any) -> Pedido:
        return Pedido.de_planilha(caminho, **cabecalho)

    # -- pedidos salvos ---------------------------------------------------- #
    def listar_pedidos(self) -> list[Path]:
        pasta = self.config.pasta_pedidos
        if not pasta.is_dir():
            return []
        return sorted(pasta.glob("*.json"), key=lambda p: p.stat().st_mtime, reverse=True)

    def salvar_pedido(self, pedido: Pedido, nome: str | None = None) -> Path:
        base = nome or pedido.referencia or date.today().isoformat()
        seguro = "".join(c if c.isalnum() or c in "-_." else "-" for c in base).strip("-")
        return pedido.salvar(self.config.pasta_pedidos / f"{seguro or 'pedido'}.json")


@lru_cache(maxsize=1)
def servico_padrao() -> Servico:
    """Instância compartilhada (o catálogo é carregado uma única vez)."""
    return Servico()
