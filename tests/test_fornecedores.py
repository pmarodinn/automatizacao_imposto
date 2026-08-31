"""Escolha automática da logo pelo fornecedor dos itens do pedido."""

import copy

import pytest

from augeo_compras.catalogo import Catalogo, Item
from augeo_compras.config import Comercial, Fornecedor, Fornecedores
from augeo_compras.pedido import LinhaPedido, Pedido, resolver
from augeo_compras.planilha import GeradorPlanilha


@pytest.fixture
def catalogo_dois_fornecedores():
    return Catalogo(
        itens=[
            Item("A-1", "T1", "Item A1", 100.0, "Sistema A", fornecedor="securiton"),
            Item("A-2", "T2", "Item A2", 100.0, "Sistema A", fornecedor="securiton"),
            Item("B-1", "T3", "Item B1", 100.0, "Sistema B", fornecedor="outro"),
        ]
    )


def _config(config, **ajustes):
    cfg = copy.deepcopy(config)
    cfg.fornecedores = Fornecedores(
        itens={
            "securiton": Fornecedor("securiton", "Securiton AG", "logo/SECURITON.jpg"),
            "outro": Fornecedor("outro", "Outro Fabricante", "logo/augeo.png"),
        },
        padrao="securiton",
        **ajustes,
    )
    return cfg


def test_itens_carregam_o_fornecedor_da_pasta(catalogo_real):
    item = catalogo_real.por_part_number("11-2000001-01-04")
    assert item.fornecedor == "securiton"


def test_logo_segue_o_fornecedor_dos_itens(config, catalogo_dois_fornecedores):
    pedido = Pedido(linhas=[LinhaPedido("B-1", 1)])
    resolvido = resolver(pedido, catalogo_dois_fornecedores, Comercial())

    escolhido = GeradorPlanilha(_config(config)).fornecedor_do_pedido(resolvido)
    assert escolhido.id == "outro"


def test_pedido_misto_usa_o_fornecedor_dominante(config, catalogo_dois_fornecedores):
    pedido = Pedido(
        linhas=[LinhaPedido("A-1", 1), LinhaPedido("A-2", 1), LinhaPedido("B-1", 1)]
    )
    resolvido = resolver(pedido, catalogo_dois_fornecedores, Comercial())

    assert resolvido.fornecedores_misturados
    assert resolvido.contagem_por_fornecedor() == {"securiton": 2, "outro": 1}
    assert GeradorPlanilha(_config(config)).fornecedor_do_pedido(resolvido).id == "securiton"


def test_pedido_misto_pode_omitir_a_logo(config, catalogo_dois_fornecedores):
    pedido = Pedido(linhas=[LinhaPedido("A-1", 1), LinhaPedido("B-1", 1)])
    resolvido = resolver(pedido, catalogo_dois_fornecedores, Comercial())

    cfg = _config(config, pedido_misto="nenhuma")
    assert GeradorPlanilha(cfg).fornecedor_do_pedido(resolvido) is None


def test_item_sem_fornecedor_cai_no_padrao(config, catalogo_dois_fornecedores):
    pedido = Pedido(
        linhas=[LinhaPedido("SERVICO", 1, descricao="Serviço", preco_unitario=100.0)]
    )
    resolvido = resolver(pedido, catalogo_dois_fornecedores, Comercial())

    assert resolvido.fornecedor_dominante == ""
    assert GeradorPlanilha(_config(config)).fornecedor_do_pedido(resolvido).id == "securiton"


def test_as_duas_logos_entram_na_proforma(config, catalogo_real, tmp_path):
    """A da empresa à esquerda e a do fornecedor à direita, sem sobrepor texto."""
    import openpyxl

    from augeo_compras.servico import Servico

    servico = Servico(config)
    servico._catalogo = catalogo_real
    caminho, _ = servico.gerar(
        Pedido(referencia="LOGOS", linhas=[LinhaPedido("11-2000001-01-04", 1)]),
        tmp_path / "logos.xlsx",
    )
    proforma = openpyxl.load_workbook(caminho)["PROFORMA"]

    ancoras = sorted(img.anchor._from.col for img in proforma._images)
    assert len(ancoras) == 2
    assert ancoras[0] == 0  # empresa, na coluna A

    # O texto do timbre começa abaixo das logos, sem ficar coberto.
    textos = {c.value for linha in proforma.iter_rows(max_row=14) for c in linha if c.value}
    assert "PARANÁ EM REDE SISTEMA LTDA" in textos
    assert "PROFORMA:" in textos
