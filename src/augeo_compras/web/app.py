"""API da interface web.

Serve a página única de `static/` e expõe o mesmo comportamento da CLI:
buscar no catálogo, montar o pedido, ver os totais e baixar a planilha.
"""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any

from fastapi import Body, FastAPI, HTTPException, Query
from fastapi.responses import FileResponse, HTMLResponse
from fastapi.staticfiles import StaticFiles

from ..pedido import Pedido, PedidoInvalido
from ..planilha import nome_sugerido
from ..servico import Servico

PASTA_ESTATICA = Path(__file__).parent / "static"


def criar_app(servico: Servico | None = None) -> FastAPI:
    servico = servico or Servico()
    app = FastAPI(title="Automação de compras — Securiton", docs_url="/api/docs")

    # -- página ------------------------------------------------------------ #
    @app.get("/", response_class=HTMLResponse)
    def pagina() -> HTMLResponse:
        return HTMLResponse((PASTA_ESTATICA / "index.html").read_text(encoding="utf-8"))

    # -- configuração ------------------------------------------------------ #
    @app.get("/api/config")
    def config() -> dict[str, Any]:
        return servico.descricao_config()

    @app.post("/api/catalogo/recarregar")
    def recarregar() -> dict[str, Any]:
        catalogo = servico.recarregar_catalogo()
        return {"total_itens": len(catalogo), "sistemas": catalogo.sistemas, "avisos": catalogo.avisos}

    # -- catálogo ---------------------------------------------------------- #
    @app.get("/api/catalogo/buscar")
    def buscar(
        q: str = Query("", description="termos de busca"),
        sistema: str | None = None,
        categoria: str | None = None,
        limite: int = Query(60, ge=1, le=500),
    ) -> dict[str, Any]:
        resultados = servico.buscar(q, sistema=sistema, categoria=categoria, limite=limite)
        return {"itens": [r.para_dict() for r in resultados]}

    @app.get("/api/catalogo/categorias")
    def categorias(sistema: str | None = None) -> dict[str, Any]:
        return {"categorias": servico.catalogo.categorias(sistema)}

    @app.get("/api/catalogo/exportar")
    def exportar_catalogo() -> FileResponse:
        destino = Path(tempfile.gettempdir()) / "catalogo_com_desconto.xlsx"
        servico.exportar_catalogo(destino)
        return FileResponse(
            destino,
            filename="catalogo_com_desconto.xlsx",
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # -- pedido ------------------------------------------------------------ #
    @app.post("/api/pedido/resumo")
    def resumo(dados: dict[str, Any] = Body(...)) -> dict[str, Any]:
        try:
            resolvido = servico.resolver(Pedido.de_dict(dados))
        except PedidoInvalido as erro:
            raise HTTPException(status_code=400, detail=str(erro)) from erro
        return servico.resumo(resolvido)

    @app.post("/api/pedido/gerar")
    def gerar(dados: dict[str, Any] = Body(...)) -> FileResponse:
        pedido = Pedido.de_dict(dados)
        try:
            resolvido = servico.resolver(pedido)
        except PedidoInvalido as erro:
            raise HTTPException(status_code=400, detail=str(erro)) from erro
        nome = nome_sugerido(resolvido)
        caminho = servico.config.pasta_saida / nome
        from ..planilha import gerar as gerar_planilha

        gerar_planilha(servico.config, resolvido, caminho)
        return FileResponse(
            caminho,
            filename=nome,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @app.get("/api/pedidos")
    def listar_pedidos() -> dict[str, Any]:
        return {
            "pedidos": [
                {"nome": p.stem, "arquivo": p.name, "modificado": p.stat().st_mtime}
                for p in servico.listar_pedidos()
            ]
        }

    @app.get("/api/pedidos/{nome}")
    def abrir_pedido(nome: str) -> dict[str, Any]:
        caminho = servico.config.pasta_pedidos / f"{Path(nome).stem}.json"
        if not caminho.is_file():
            raise HTTPException(status_code=404, detail="Pedido não encontrado")
        return Pedido.carregar(caminho).para_dict()

    @app.post("/api/pedidos")
    def salvar_pedido(dados: dict[str, Any] = Body(...)) -> dict[str, Any]:
        pedido = Pedido.de_dict(dados)
        caminho = servico.salvar_pedido(pedido)
        return {"arquivo": caminho.name, "nome": caminho.stem}

    if PASTA_ESTATICA.is_dir():
        app.mount("/static", StaticFiles(directory=PASTA_ESTATICA), name="static")

    return app


app = criar_app()
