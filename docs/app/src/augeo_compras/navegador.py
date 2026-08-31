"""Ponte para quando o sistema roda dentro do navegador.

No site publicado não existe servidor: o mesmo código Python roda em
WebAssembly (Pyodide) e estas funções fazem o papel dos endpoints da API. Elas
trocam JSON com a interface — arquivos binários vão em base64 — para que o
contrato seja idêntico ao do servidor local.

As listas de preço não vêm do repositório: quem usa carrega os arquivos na
página, e eles ficam só na memória do navegador.
"""

from __future__ import annotations

import base64
import json
from pathlib import Path
from typing import Any

from . import catalogo as mod_catalogo
from .catalogo import Catalogo
from .config import carregar as carregar_config
from .pedido import Pedido, PedidoInvalido
from .planilha import gerar as gerar_planilha, nome_sugerido
from .servico import Servico

PASTA_TEMP = Path("/tmp/augeo")

_servico: Servico | None = None


def _ativo() -> Servico:
    if _servico is None:
        raise RuntimeError("chame iniciar() antes")
    return _servico


def _resposta(**dados: Any) -> str:
    return json.dumps(dados, ensure_ascii=False, default=str)


# --------------------------------------------------------------------------- #
# Ciclo de vida
# --------------------------------------------------------------------------- #
def iniciar() -> str:
    """Prepara o sistema. Chamada uma única vez, ao abrir a página."""
    global _servico
    _servico = Servico(carregar_config())
    # Começa sem catálogo: as listas são carregadas por quem usa a página.
    _servico._catalogo = Catalogo()
    PASTA_TEMP.mkdir(parents=True, exist_ok=True)
    return configuracao()


def configuracao() -> str:
    """Estado atual da configuração — pode ser consultada quantas vezes for
    preciso, sem reiniciar nada."""
    return _resposta(**_ativo().descricao_config())


def definir_desconto(percentual: float) -> str:
    """O desconto negociado não vai para o repositório; é informado na tela."""
    servico = _ativo()
    servico.config.comercial.desconto_percentual = float(percentual)
    return _resposta(desconto_percentual=servico.config.comercial.desconto_percentual)


def carregar_listas(caminhos_json: str) -> str:
    """Lê as listas de preço já gravadas no sistema de arquivos do navegador."""
    servico = _ativo()
    caminhos = [Path(c) for c in json.loads(caminhos_json)]
    servico._catalogo = mod_catalogo.carregar(servico.config, caminhos=caminhos)
    return _resposta(
        total_itens=len(servico.catalogo),
        sistemas=servico.catalogo.sistemas,
        avisos=servico.catalogo.avisos,
        arquivos=[c.name for c in caminhos],
    )


# --------------------------------------------------------------------------- #
# Catálogo e pedido
# --------------------------------------------------------------------------- #
def buscar(termo: str = "", sistema: str = "", limite: int = 60) -> str:
    resultados = _ativo().buscar(termo, sistema=sistema or None, limite=limite)
    return _resposta(itens=[r.para_dict() for r in resultados])


def resumo(pedido_json: str) -> str:
    servico = _ativo()
    try:
        resolvido = servico.resolver(Pedido.de_dict(json.loads(pedido_json)))
    except PedidoInvalido as erro:
        return _resposta(erro=str(erro))
    return _resposta(**servico.resumo(resolvido))


def gerar(pedido_json: str) -> str:
    """Gera a planilha e devolve o arquivo em base64, para o navegador baixar."""
    servico = _ativo()
    try:
        resolvido = servico.resolver(Pedido.de_dict(json.loads(pedido_json)))
    except PedidoInvalido as erro:
        return _resposta(erro=str(erro))

    destino = PASTA_TEMP / nome_sugerido(resolvido)
    gerar_planilha(servico.config, resolvido, destino)
    return _resposta(
        nome=destino.name,
        conteudo=base64.b64encode(destino.read_bytes()).decode(),
    )


def exportar_catalogo() -> str:
    servico = _ativo()
    if not len(servico.catalogo):
        return _resposta(erro="Carregue as listas de preço antes de exportar o catálogo.")
    destino = PASTA_TEMP / "catalogo_com_desconto.xlsx"
    servico.exportar_catalogo(destino)
    return _resposta(
        nome=destino.name,
        conteudo=base64.b64encode(destino.read_bytes()).decode(),
    )
