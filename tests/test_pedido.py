"""Modelo do pedido, totais e conciliação compra x proforma."""

import pytest

from augeo_compras.config import Comercial
from augeo_compras.pedido import (
    LinhaPedido,
    Pedido,
    PedidoInvalido,
    Proforma,
    resolver,
)


def comercial(**ajustes) -> Comercial:
    base = Comercial(desconto_percentual=50.0, posicao_inicial=10, posicao_passo=10)
    for chave, valor in ajustes.items():
        setattr(base, chave, valor)
    return base


def test_posicoes_seguem_o_passo_configurado(catalogo_falso):
    pedido = Pedido(linhas=[LinhaPedido("11-0000001-01-01", 1), LinhaPedido("11-0000002-01-01", 1)])
    resolvido = resolver(pedido, catalogo_falso, comercial(posicao_inicial=100, posicao_passo=5))
    assert [linha.posicao for linha in resolvido.linhas] == [100, 105]


def test_totais_sem_divergencia(catalogo_falso):
    pedido = Pedido(linhas=[LinhaPedido("11-0000001-01-01", 2), LinhaPedido("22-0000003-01-01", 3)])
    resolvido = resolver(pedido, catalogo_falso, comercial())

    # 1000 * 0,5 * 2 = 1000 ; 33,33 * 0,5 = 16,67 (arredondado) * 3 = 50,01
    assert resolvido.total_controle == 1050.01
    assert resolvido.valor_faturado == 1050.01
    assert resolvido.valor_servico == 0.0
    assert resolvido.total_geral == 1050.01


def test_quantidade_maior_na_proforma_vira_valor_de_servico_negativo(catalogo_falso):
    pedido = Pedido(
        linhas=[LinhaPedido("11-0000002-01-01", quantidade=5, quantidade_proforma=6)]
    )
    resolvido = resolver(pedido, catalogo_falso, comercial())

    unitario = resolvido.linhas[0].preco_unitario  # 125,00
    assert resolvido.total_controle == 5 * unitario
    assert resolvido.valor_faturado == 6 * unitario
    assert resolvido.valor_servico == -unitario
    # O total geral continua sendo o custo real da compra.
    assert resolvido.total_geral == resolvido.total_controle
    assert resolvido.linhas[0].divergente


def test_frete_entra_no_faturado_e_no_total(catalogo_falso):
    pedido = Pedido(
        proformas=[Proforma(numero="1", frete=100.0)],
        linhas=[LinhaPedido("11-0000001-01-01", 1)],
    )
    resolvido = resolver(pedido, catalogo_falso, comercial())
    assert resolvido.total_controle == 500.0
    assert resolvido.frete_total == 100.0
    assert resolvido.valor_faturado == 600.0
    assert resolvido.valor_servico == 0.0
    assert resolvido.total_geral == 600.0


def test_divisao_em_duas_proformas(catalogo_falso):
    pedido = Pedido(
        proformas=[Proforma(numero="A"), Proforma(numero="B")],
        linhas=[
            LinhaPedido("11-0000001-01-01", 1, proforma=1),
            LinhaPedido("11-0000002-01-01", 4, proforma=2),
        ],
    )
    resolvido = resolver(pedido, catalogo_falso, comercial())

    assert [len(p.linhas) for p in resolvido.proformas] == [1, 1]
    assert resolvido.proformas[0].subtotal == 500.0
    assert resolvido.proformas[1].subtotal == 500.0
    assert resolvido.valor_faturado == 1000.0
    assert resolvido.valor_servico == 0.0


def test_item_inexistente_e_recusado(catalogo_falso):
    pedido = Pedido(linhas=[LinhaPedido("NAO-EXISTE", 1)])
    with pytest.raises(PedidoInvalido, match="não existe"):
        resolver(pedido, catalogo_falso, comercial())


def test_item_fora_de_lista_com_preco_informado(catalogo_falso):
    pedido = Pedido(
        linhas=[LinhaPedido("SERVICO-01", 1, descricao="Serviço de comissionamento", preco_unitario=800.0)]
    )
    resolvido = resolver(pedido, catalogo_falso, comercial())
    assert resolvido.linhas[0].preco_unitario == 800.0  # sem desconto
    assert resolvido.total_controle == 800.0
    assert any("fora das listas" in aviso for aviso in resolvido.avisos)


def test_condicoes_da_proforma_herdam_o_padrao(catalogo_falso):
    from augeo_compras.config import Condicoes

    padrao = Condicoes(payment_terms="60 dias", shipment_terms="EXW")
    pedido = Pedido(
        proformas=[Proforma(numero="A"), Proforma(numero="B", condicoes=Condicoes(payment_terms="90 dias"))],
        linhas=[LinhaPedido("11-0000001-01-01", 1, proforma=1), LinhaPedido("11-0000002-01-01", 1, proforma=2)],
    )
    resolvido = resolver(pedido, catalogo_falso, comercial(condicoes_padrao=padrao))

    assert resolvido.proformas[0].condicoes.payment_terms == "60 dias"
    assert resolvido.proformas[1].condicoes.payment_terms == "90 dias"
    assert resolvido.proformas[1].condicoes.shipment_terms == "EXW"


def test_pedido_sobrevive_a_ida_e_volta_em_json(tmp_path, catalogo_falso):
    pedido = Pedido(
        referencia="R1",
        cliente="Cliente",
        proformas=[Proforma(numero="A", frete=12.5)],
        linhas=[LinhaPedido("11-0000001-01-01", 3, quantidade_proforma=2, observacao="parcial")],
    )
    caminho = pedido.salvar(tmp_path / "p.json")
    recuperado = Pedido.carregar(caminho)

    assert recuperado.para_dict() == pedido.para_dict()
    assert recuperado.linhas[0].quantidade_proforma == 2
    assert recuperado.proformas[0].frete == 12.5
