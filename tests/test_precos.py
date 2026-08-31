"""Regras de desconto."""

from augeo_compras.catalogo import Item
from augeo_compras.config import Comercial
from augeo_compras.precos import TabelaDePrecos


def comercial(**ajustes) -> Comercial:
    base = Comercial(desconto_percentual=44.5, casas_decimais=2)
    for chave, valor in ajustes.items():
        setattr(base, chave, valor)
    return base


def test_desconto_padrao():
    tabela = TabelaDePrecos(comercial())
    item = Item("PN", "T", "d", 1198.0, "Sistema")
    calculo = tabela.calcular(item)
    assert calculo.preco_liquido == 664.89
    assert calculo.desconto_percentual == 44.5
    assert calculo.fonte_desconto == "desconto padrão"


def test_arredondamento_meio_para_cima():
    tabela = TabelaDePrecos(comercial())
    # 45.70 * 0.555 = 25.3635 -> 25.36 ; 12.29 * 0.555 = 6.82095 -> 6.82
    assert tabela.calcular(Item("A", "", "", 45.70, "S")).preco_liquido == 25.36
    assert tabela.calcular(Item("B", "", "", 10.0, "S")).preco_liquido == 5.55
    # meio exato sobe (ROUND_HALF_UP, não bankers rounding)
    assert TabelaDePrecos(comercial(desconto_percentual=0)).arredondar(2.345) == 2.35


def test_precedencia_part_number_sobre_tipo_e_sistema():
    cfg = comercial(
        descontos_por_sistema={"Sistema A": 40.0},
        descontos_por_tipo={"AAA 100": 30.0},
        descontos_por_part_number={"PN-1": 10.0},
    )
    tabela = TabelaDePrecos(cfg)
    item = Item("PN-1", "AAA 100", "d", 100.0, "Sistema A")

    assert tabela.calcular(item).preco_liquido == 90.0

    sem_pn = Item("PN-2", "AAA 100", "d", 100.0, "Sistema A")
    assert tabela.calcular(sem_pn).preco_liquido == 70.0

    so_sistema = Item("PN-3", "ZZZ", "d", 100.0, "Sistema A")
    assert tabela.calcular(so_sistema).preco_liquido == 60.0

    nenhum = Item("PN-4", "ZZZ", "d", 100.0, "Sistema B")
    assert tabela.calcular(nenhum).preco_liquido == 55.5


def test_desconto_informado_no_pedido_vence_tudo():
    cfg = comercial(descontos_por_part_number={"PN-1": 10.0})
    tabela = TabelaDePrecos(cfg)
    item = Item("PN-1", "", "", 100.0, "S")
    assert tabela.calcular(item, desconto=25).preco_liquido == 75.0
