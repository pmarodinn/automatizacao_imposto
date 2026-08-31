#!/usr/bin/env python3
"""Monta o site estático publicado no GitHub Pages, em `docs/`.

O site é a mesma interface do servidor local, mas sem servidor: o código Python
roda dentro do navegador, em WebAssembly. Este script junta as peças que ele
precisa baixar — os módulos do pacote, a configuração e as logos — e escreve um
manifesto com a lista.

    python3 ferramentas/gerar_site.py

Nada de confidencial entra aqui: as listas de preço e o `config/comercial.yaml`
(desconto negociado e dados bancários) ficam de fora de propósito. Quem usa o
site carrega as próprias listas e informa o desconto na tela.
"""

from __future__ import annotations

import json
import shutil
import sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parents[1]
DESTINO = RAIZ / "docs"
APP = DESTINO / "app"

# Módulos que o navegador precisa. `cli.py` e `web/` ficam de fora: dependem de
# argparse e FastAPI, que só fazem sentido no servidor local.
MODULOS = [
    "__init__.py",
    "config.py",
    "catalogo.py",
    "precos.py",
    "pedido.py",
    "estilos.py",
    "planilha.py",
    "servico.py",
    "navegador.py",
]

# Configuração que vai para o site. `comercial.yaml` NUNCA entra — o exemplo
# assume o lugar dele, e o desconto real é informado na tela.
CONFIGS = [
    "empresa.yaml",
    "catalogo.yaml",
    "fornecedores.yaml",
    "layout.yaml",
    "comercial.exemplo.yaml",
]

PROIBIDOS = {"comercial.yaml"}


def copiar(origem: Path, relativo: str) -> str:
    destino = APP / relativo
    destino.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(origem, destino)
    return relativo


def main() -> int:
    if not (RAIZ / "src" / "augeo_compras").is_dir():
        print("execute a partir da raiz do projeto", file=sys.stderr)
        return 1

    if APP.exists():
        shutil.rmtree(APP)
    DESTINO.mkdir(exist_ok=True)

    arquivos: list[str] = []

    for modulo in MODULOS:
        arquivos.append(
            copiar(RAIZ / "src" / "augeo_compras" / modulo, f"src/augeo_compras/{modulo}")
        )

    for nome in CONFIGS:
        if nome in PROIBIDOS:
            raise SystemExit(f"{nome} não pode ir para o site")
        origem = RAIZ / "config" / nome
        if origem.exists():
            arquivos.append(copiar(origem, f"config/{nome}"))

    for logo in sorted((RAIZ / "logo").glob("*")):
        if logo.is_file():
            arquivos.append(copiar(logo, f"logo/{logo.name}"))

    wheels = sorted(w.name for w in (DESTINO / "wheels").glob("*.whl"))
    if not wheels:
        print("aviso: nenhuma wheel em docs/wheels — o site não vai gerar planilhas",
              file=sys.stderr)

    (APP / "manifest.json").write_text(
        json.dumps(
            {"arquivos": arquivos, "wheels": [f"wheels/{w}" for w in wheels]},
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    shutil.copy2(RAIZ / "src" / "augeo_compras" / "web" / "static" / "index.html",
                 DESTINO / "index.html")
    # Sem isto o GitHub Pages passa tudo pelo Jekyll e ignora a pasta `app`.
    (DESTINO / ".nojekyll").write_text("", encoding="utf-8")

    # Rede de segurança: nada de lista de preço ou desconto real no site.
    for suspeito in DESTINO.rglob("*"):
        if suspeito.name in PROIBIDOS or suspeito.suffix.lower() in {".xls", ".xlsx"}:
            raise SystemExit(f"arquivo indevido no site: {suspeito}")

    print(f"site em {DESTINO.relative_to(Path.cwd())}/")
    print(f"  {len(arquivos)} arquivos do sistema + {len(wheels)} wheel(s)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
