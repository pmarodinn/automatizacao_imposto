"""Emitentes, CNPJ e as logos escolhidas na tela de identificação."""

import base64
import copy

import openpyxl
import pytest

from augeo_compras.config import Empresa
from augeo_compras.pedido import LinhaPedido, Pedido
from augeo_compras.servico import Servico


@pytest.fixture
def servico(config, catalogo_real, tmp_path):
    cfg = copy.deepcopy(config)
    cfg.raiz = tmp_path          # logos enviadas vão para um lugar descartável
    servico = Servico(cfg)
    servico._catalogo = catalogo_real
    return servico


PIXEL_PNG = base64.b64encode(bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4"
    "890000000a49444154789c6360000002000100fef8b6a80000000049454e44ae426082"
)).decode()


# --------------------------------------------------------------------------- #
# Emitentes
# --------------------------------------------------------------------------- #
def test_config_traz_os_emitentes_do_yaml(config):
    identificadores = [e.identificador for e in config.emitentes]
    assert "parana-em-rede" in identificadores
    assert "augeo-engenharia" in identificadores
    assert config.empresa.identificador == "parana-em-rede"


def test_cnpj_entra_no_timbre():
    empresa = Empresa(
        razao_social="AUGEO ENGENHARIA LTDA",
        cnpj="12.345.678/0001-90",
        endereco=["RUA X, 1", "CURITIBA - PR"],
        telefone="+55 41 0000-0000",
    )
    assert empresa.linhas_timbre() == [
        "AUGEO ENGENHARIA LTDA",
        "CNPJ: 12.345.678/0001-90",
        "RUA X, 1",
        "CURITIBA - PR",
        "Phone: +55 41 0000-0000",
    ]


def test_sem_cnpj_o_timbre_nao_ganha_linha_vazia():
    empresa = Empresa(razao_social="EMPRESA", endereco=["RUA X"])
    assert empresa.linhas_timbre() == ["EMPRESA", "RUA X"]


# --------------------------------------------------------------------------- #
# Troca de identidade
# --------------------------------------------------------------------------- #
def test_identidade_comeca_incompleta_sem_cnpj(servico):
    identidade = servico.identidade()
    assert identidade["completo"] is False
    assert len(identidade["emitentes"]) == 2
    assert identidade["logo_emitente"].startswith("data:image/")


def test_trocar_de_emitente(servico):
    resultado = servico.aplicar_identidade({
        "emitente": {
            "identificador": "augeo-engenharia",
            "razao_social": "AUGEO ENGENHARIA LTDA",
            "cnpj": "12.345.678/0001-90",
            "endereco": ["RUA X, 1"],
        }
    })
    assert resultado["completo"] is True
    assert servico.config.empresa.identificador == "augeo-engenharia"
    assert servico.config.empresa.cnpj == "12.345.678/0001-90"
    # O outro emitente continua disponível para voltar depois.
    assert {e["identificador"] for e in resultado["emitentes"]} == {
        "parana-em-rede", "augeo-engenharia"
    }


def test_logo_enviada_substitui_a_do_emitente(servico):
    anterior = servico.config.empresa.logo
    servico.aplicar_identidade({
        "emitente": {
            "identificador": "parana-em-rede",
            "razao_social": "PARANÁ EM REDE SISTEMA LTDA",
            "cnpj": "00.000.000/0001-00",
            "logo_conteudo": PIXEL_PNG,
            "logo_nome": "nova.png",
        }
    })
    assert servico.config.empresa.logo != anterior
    assert servico.config.empresa.caminho_logo.exists()
    assert servico.config.empresa.caminho_logo.name == "nova.png"


def test_nome_e_logo_do_fornecedor(servico):
    resultado = servico.aplicar_identidade({
        "fornecedor": {"id": "securiton", "nome": "Securiton do Brasil",
                       "logo_conteudo": PIXEL_PNG, "logo_nome": "forn.png"}
    })
    assert resultado["fornecedor"]["nome"] == "Securiton do Brasil"
    assert servico.config.fornecedores.itens["securiton"].caminho_logo.name == "forn.png"


def test_fornecedor_novo_vira_o_padrao(servico):
    servico.aplicar_identidade({"fornecedor": {"id": "outro", "nome": "Outro Fabricante"}})
    assert servico.config.fornecedores.padrao == "outro"
    assert servico.config.fornecedores.get(None).nome == "Outro Fabricante"


# --------------------------------------------------------------------------- #
# Efeito na planilha
# --------------------------------------------------------------------------- #
def test_planilha_sai_com_o_emitente_escolhido(servico, tmp_path):
    servico.aplicar_identidade({
        "emitente": {
            "identificador": "augeo-engenharia",
            "razao_social": "AUGEO ENGENHARIA LTDA",
            "cnpj": "12.345.678/0001-90",
            "endereco": ["RUA ISAÍAS REGIS DE MIRANDA, 689", "CURITIBA - PR"],
            "cidade": "CURITIBA",
            "responsavel": "VALDIR MARODIN JÚNIOR",
        }
    })
    caminho, _ = servico.gerar(
        Pedido(referencia="ID", linhas=[LinhaPedido("11-2000001-01-04", 1)]),
        tmp_path / "id.xlsx",
    )
    proforma = openpyxl.load_workbook(caminho)["PROFORMA"]
    textos = [c.value for linha in proforma.iter_rows(max_row=16) for c in linha if c.value]

    assert "AUGEO ENGENHARIA LTDA" in textos
    assert "CNPJ: 12.345.678/0001-90" in textos
    assert "PARANÁ EM REDE SISTEMA LTDA" not in textos
    assert len(proforma._images) == 2   # emitente à esquerda, fornecedor à direita
