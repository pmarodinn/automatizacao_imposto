"""Leitura das listas de preço."""

import pytest

from augeo_compras.catalogo import converter_preco, normalizar_part_number


@pytest.mark.parametrize(
    "entrada, esperado",
    [
        ("1398.00", 1398.0),
        ("1.398,00", 1398.0),
        ("1,398.00", 1398.0),
        ("45,70", 45.7),
        (121, 121.0),
        ("  978.00  ", 978.0),
        ("", None),
        (None, None),
        ("sob consulta", None),
    ],
)
def test_converter_preco(entrada, esperado):
    assert converter_preco(entrada) == esperado


def test_normalizar_part_number_remove_espacos_e_nbsp():
    assert normalizar_part_number(" 11-2000001-01-04 ") == "11-2000001-01-04"
    assert normalizar_part_number("11-2000001\xa0-01-04") == "11-2000001-01-04"


def test_config_recorre_ao_exemplo_quando_falta_o_arquivo(tmp_path):
    """Um clone novo não tem `comercial.yaml`; o `.exemplo.yaml` assume."""
    from augeo_compras.config import carregar

    (tmp_path / "comercial.exemplo.yaml").write_text(
        "desconto_percentual: 12.5\nmoeda: USD\n", encoding="utf-8"
    )
    config = carregar(tmp_path)
    assert config.comercial.desconto_percentual == 12.5
    assert config.comercial.moeda == "USD"

    # Existindo o arquivo real, é ele que vale.
    (tmp_path / "comercial.yaml").write_text("desconto_percentual: 44.5\n", encoding="utf-8")
    assert carregar(tmp_path).comercial.desconto_percentual == 44.5


def test_listas_reais_sao_lidas_sem_avisos(catalogo_real):
    assert len(catalogo_real) > 800
    assert len(catalogo_real.sistemas) == 8
    assert catalogo_real.avisos == []


def test_categoria_e_herdada_dos_titulos(catalogo_real):
    item = catalogo_real.por_part_number("11-2000001-01-04")
    assert item is not None
    assert item.tipo == "ASD 533-1"
    assert item.categoria == "ASD Base Units"
    assert item.sistema == "SecuriSmoke ASD 53x"


def test_busca_ignora_acento_ordem_e_caixa(catalogo_real):
    resultados = catalogo_real.buscar("535 aspirating")
    assert resultados
    assert all("aspirating" in i.descricao.lower() for i in resultados)


def test_busca_por_sistema(catalogo_real):
    resultados = catalogo_real.buscar("", sistema="Power Supplies")
    assert resultados
    assert {i.sistema for i in resultados} == {"Power Supplies"}


def test_resolver_aceita_part_number_ou_tipo(catalogo_real):
    por_pn = catalogo_real.resolver("11-2300030-01-02")
    por_tipo = catalogo_real.resolver("DFU 911")
    assert por_pn is not None and por_pn == por_tipo
