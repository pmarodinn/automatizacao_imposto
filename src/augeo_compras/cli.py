"""Linha de comando do sistema.

    python -m augeo_compras web                     abre a interface no navegador
    python -m augeo_compras buscar "asd 535"        procura itens no catálogo
    python -m augeo_compras catalogo --exportar     exporta o catálogo com preços
    python -m augeo_compras modelo pedido.json      cria um pedido de exemplo
    python -m augeo_compras gerar pedido.json       gera a planilha do pedido
    python -m augeo_compras importar lista.xlsx     lê part number + quantidade
"""

from __future__ import annotations

import argparse
import sys
from datetime import date
from pathlib import Path

from .config import carregar as carregar_config
from .pedido import LinhaPedido, Pedido, PedidoInvalido, Proforma
from .planilha import nome_sugerido
from .servico import Servico

MOEDA = "{:>10,.2f}"


def _servico(args: argparse.Namespace) -> Servico:
    config = carregar_config(args.config) if getattr(args, "config", None) else carregar_config()
    return Servico(config)


def _mostrar_avisos(avisos: list[str], titulo: str = "Avisos") -> None:
    if not avisos:
        return
    print(f"\n{titulo}:", file=sys.stderr)
    for aviso in avisos:
        print(f"  ! {aviso}", file=sys.stderr)


# --------------------------------------------------------------------------- #
# Comandos
# --------------------------------------------------------------------------- #
def cmd_catalogo(args: argparse.Namespace) -> int:
    servico = _servico(args)
    catalogo = servico.catalogo
    print(f"{len(catalogo)} itens em {len(catalogo.sistemas)} sistemas")
    for sistema in catalogo.sistemas:
        quantidade = sum(1 for i in catalogo if i.sistema == sistema)
        print(f"  {quantidade:>4}  {sistema}")
    _mostrar_avisos(catalogo.avisos)

    if args.exportar:
        destino = Path(args.exportar) if isinstance(args.exportar, str) else (
            servico.config.pasta_saida / "catalogo_com_desconto.xlsx"
        )
        print(f"\nExportado para {servico.exportar_catalogo(destino)}")
    return 0


def cmd_buscar(args: argparse.Namespace) -> int:
    servico = _servico(args)
    resultados = servico.buscar(
        " ".join(args.termo), sistema=args.sistema, limite=args.limite
    )
    if not resultados:
        print("Nenhum item encontrado.")
        return 1
    moeda = servico.config.comercial.moeda
    print(f"{'PART NUMBER':<20} {'TYPE':<18} {'BRUTO':>10} {'LÍQUIDO':>10}  DESCRIÇÃO")
    for resultado in resultados:
        item, preco = resultado.item, resultado.preco
        print(
            f"{item.part_number:<20} {item.tipo[:18]:<18} "
            f"{item.preco_bruto:>10,.2f} {preco.preco_liquido:>10,.2f}  {item.descricao[:60]}"
        )
    print(f"\n{len(resultados)} item(ns) — preços em {moeda}, "
          f"desconto padrão de {servico.config.comercial.desconto_percentual}%")
    return 0


def cmd_modelo(args: argparse.Namespace) -> int:
    """Cria um pedido de exemplo, já preenchido, para servir de ponto de partida."""
    servico = _servico(args)
    exemplos = [item.part_number for item in list(servico.catalogo)[:3]]
    pedido = Pedido(
        referencia="EXEMPLO",
        cliente="Securiton AG",
        data=date.today(),
        proformas=[Proforma(numero="0001-1", data=date.today(), frete=0.0)],
        linhas=[LinhaPedido(part_number=pn, quantidade=1) for pn in exemplos],
        observacoes="Substitua os itens por part numbers reais e ajuste as quantidades.",
    )
    destino = Path(args.destino)
    pedido.salvar(destino)
    print(f"Pedido de exemplo criado em {destino}")
    return 0


def cmd_gerar(args: argparse.Namespace) -> int:
    servico = _servico(args)
    pedido = Pedido.carregar(args.pedido)
    try:
        caminho, resolvido = servico.gerar(pedido, args.saida)
    except PedidoInvalido as erro:
        print(f"Erro no pedido: {erro}", file=sys.stderr)
        return 2

    moeda = servico.config.comercial.moeda
    print(f"Planilha gerada: {caminho}\n")
    print(f"  Itens                {len(resolvido.linhas):>10}")
    print(f"  Total da compra      {MOEDA.format(resolvido.total_controle)} {moeda}")
    print(f"  Frete                {MOEDA.format(resolvido.frete_total)} {moeda}")
    print(f"  Valor faturado       {MOEDA.format(resolvido.valor_faturado)} {moeda}")
    print(f"  Valor de serviço     {MOEDA.format(resolvido.valor_servico)} {moeda}")
    print(f"  Total geral          {MOEDA.format(resolvido.total_geral)} {moeda}")
    divergentes = [linha for linha in resolvido.linhas if linha.divergente]
    if divergentes:
        print(f"\n  {len(divergentes)} linha(s) declaram na proforma valor diferente da compra:")
        for linha in divergentes:
            print(
                f"    pos {linha.posicao}: {linha.part_number} "
                f"compra {linha.quantidade:g} x {linha.preco_unitario:.2f} | "
                f"proforma {linha.quantidade_proforma:g} x {linha.preco_proforma:.2f}"
            )
    _mostrar_avisos(resolvido.avisos)
    return 0


def cmd_importar(args: argparse.Namespace) -> int:
    servico = _servico(args)
    try:
        pedido = servico.pedido_de_planilha(
            args.planilha, referencia=args.referencia or Path(args.planilha).stem
        )
    except PedidoInvalido as erro:
        print(f"Erro ao importar: {erro}", file=sys.stderr)
        return 2

    destino = Path(args.destino) if args.destino else (
        servico.config.pasta_pedidos / f"{pedido.referencia}.json"
    )
    pedido.salvar(destino)
    print(f"{len(pedido.linhas)} item(ns) importado(s) para {destino}")

    if args.gerar:
        caminho, resolvido = servico.gerar(pedido)
        print(f"Planilha gerada: {caminho}")
        _mostrar_avisos(resolvido.avisos)
    return 0


def cmd_web(args: argparse.Namespace) -> int:
    import uvicorn

    from .web.app import criar_app

    print(f"\n  Interface disponível em http://{args.host}:{args.porta}\n")
    uvicorn.run(criar_app(), host=args.host, port=args.porta, log_level="warning")
    return 0


# --------------------------------------------------------------------------- #
def construir_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="augeo-compras",
        description="Gera as planilhas de compra (proforma) a partir das listas Securiton.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    parser.add_argument("--config", help="pasta com os arquivos YAML de configuração")
    sub = parser.add_subparsers(dest="comando", required=True)

    p = sub.add_parser("catalogo", help="resumo das listas de preço carregadas")
    p.add_argument(
        "--exportar", nargs="?", const=True, default=None,
        help="exporta o catálogo com desconto para Excel (caminho opcional)",
    )
    p.set_defaults(func=cmd_catalogo)

    p = sub.add_parser("buscar", help="procura itens no catálogo")
    p.add_argument("termo", nargs="+", help="palavras a procurar")
    p.add_argument("--sistema", help="restringe a um sistema (ex.: SecuriFire)")
    p.add_argument("--limite", type=int, default=30)
    p.set_defaults(func=cmd_buscar)

    p = sub.add_parser("modelo", help="cria um arquivo de pedido de exemplo")
    p.add_argument("destino", nargs="?", default="data/pedidos/exemplo.json")
    p.set_defaults(func=cmd_modelo)

    p = sub.add_parser("gerar", help="gera a planilha a partir de um pedido .json")
    p.add_argument("pedido")
    p.add_argument("--saida", help="caminho do .xlsx de saída")
    p.set_defaults(func=cmd_gerar)

    p = sub.add_parser("importar", help="cria um pedido a partir de uma planilha simples")
    p.add_argument("planilha", help="arquivo com colunas de part number e quantidade")
    p.add_argument("--referencia", help="nome do pedido")
    p.add_argument("--destino", help="onde salvar o pedido .json")
    p.add_argument("--gerar", action="store_true", help="já gera a planilha final")
    p.set_defaults(func=cmd_importar)

    p = sub.add_parser("web", help="abre a interface no navegador")
    p.add_argument("--host", default="127.0.0.1")
    p.add_argument("--porta", type=int, default=8000)
    p.set_defaults(func=cmd_web)

    return parser


def main(argv: list[str] | None = None) -> int:
    args = construir_parser().parse_args(argv)
    try:
        return args.func(args)
    except FileNotFoundError as erro:
        print(f"Arquivo não encontrado: {erro}", file=sys.stderr)
        return 2
    except KeyboardInterrupt:
        return 130


if __name__ == "__main__":  # pragma: no cover
    raise SystemExit(main())
