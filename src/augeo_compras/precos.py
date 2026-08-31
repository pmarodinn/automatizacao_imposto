"""Aplicação do desconto sobre a lista bruta da Securiton.

A regra padrão é um desconto único (44,5%), mas o sistema aceita exceções por
sistema, por tipo e por part number. Vence sempre a regra mais específica:

    part number  >  tipo  >  sistema  >  desconto padrão
"""

from __future__ import annotations

from dataclasses import dataclass
from decimal import ROUND_HALF_UP, Decimal

from .catalogo import Item, normalizar_part_number
from .config import Comercial


@dataclass(frozen=True)
class PrecoCalculado:
    preco_bruto: float
    desconto_percentual: float
    preco_liquido: float
    fonte_desconto: str  # de onde veio o percentual, para auditoria

    @property
    def valor_desconto(self) -> float:
        return round(self.preco_bruto - self.preco_liquido, 6)


class TabelaDePrecos:
    """Converte preço bruto em preço líquido conforme as regras comerciais."""

    def __init__(self, comercial: Comercial):
        self.comercial = comercial
        self._por_part_number = {
            normalizar_part_number(k): v for k, v in comercial.descontos_por_part_number.items()
        }
        self._por_tipo = {k.strip().lower(): v for k, v in comercial.descontos_por_tipo.items()}
        self._por_sistema = {k.strip().lower(): v for k, v in comercial.descontos_por_sistema.items()}

    # -- desconto ---------------------------------------------------------- #
    def desconto_de(self, item: Item) -> tuple[float, str]:
        """Percentual de desconto do item e a regra que o originou."""
        chave = normalizar_part_number(item.part_number)
        if chave in self._por_part_number:
            return self._por_part_number[chave], f"part number {item.part_number}"

        tipo = (item.tipo or "").strip().lower()
        if tipo and tipo in self._por_tipo:
            return self._por_tipo[tipo], f"tipo {item.tipo}"

        sistema = (item.sistema or "").strip().lower()
        if sistema and sistema in self._por_sistema:
            return self._por_sistema[sistema], f"sistema {item.sistema}"

        return self.comercial.desconto_percentual, "desconto padrão"

    # -- preço ------------------------------------------------------------- #
    def arredondar(self, valor: float) -> float:
        casas = max(0, int(self.comercial.casas_decimais))
        quantum = Decimal(1).scaleb(-casas)
        return float(Decimal(str(valor)).quantize(quantum, rounding=ROUND_HALF_UP))

    def calcular(self, item: Item, *, desconto: float | None = None) -> PrecoCalculado:
        """Preço líquido do item. `desconto` sobrescreve as regras."""
        if desconto is None:
            percentual, fonte = self.desconto_de(item)
        else:
            percentual, fonte = float(desconto), "informado no pedido"

        liquido = self.arredondar(item.preco_bruto * (1 - percentual / 100.0))
        return PrecoCalculado(
            preco_bruto=item.preco_bruto,
            desconto_percentual=percentual,
            preco_liquido=liquido,
            fonte_desconto=fonte,
        )

    def preco_liquido(self, item: Item, *, desconto: float | None = None) -> float:
        return self.calcular(item, desconto=desconto).preco_liquido
