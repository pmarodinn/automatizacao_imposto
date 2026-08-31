import pytest

from augeo_compras.catalogo import Catalogo, Item
from augeo_compras.config import carregar as carregar_config


@pytest.fixture(scope="session")
def config():
    return carregar_config()


@pytest.fixture(scope="session")
def catalogo_real(config):
    """Catálogo lido das listas de preço reais.

    As listas são material do fabricante e ficam fora do Git, então num clone
    novo elas não existem: os testes que dependem delas são pulados em vez de
    quebrarem.
    """
    from augeo_compras import catalogo as mod

    catalogo = mod.carregar(config)
    if not len(catalogo):
        pytest.skip("listas de preço não disponíveis nesta máquina")
    return catalogo


@pytest.fixture
def catalogo_falso():
    """Catálogo mínimo e previsível, para testes que não dependem das listas."""
    return Catalogo(
        itens=[
            Item("11-0000001-01-01", "AAA 100", "Item alfa", 1000.0, "Sistema A", "Categoria 1"),
            Item("11-0000002-01-01", "BBB 200", "Item beta", 250.0, "Sistema A", "Categoria 2"),
            Item("22-0000003-01-01", "CCC 300", "Item gama", 33.33, "Sistema B", "Categoria 1"),
        ]
    )
