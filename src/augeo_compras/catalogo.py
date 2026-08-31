"""Leitura das listas de preço da Securiton.

Cada arquivo de lista é um sistema (ASD, SecuriFire, ADW...) e segue o mesmo
padrão: uma linha de cabeçalho com Description / Type / Part. No. / Gross Price
e, abaixo, itens intercalados com linhas de categoria (só o texto na primeira
coluna).

O leitor detecta o cabeçalho sozinho, então listas novas ou com colunas em
outra ordem continuam funcionando sem alterar código.
"""

from __future__ import annotations

import csv
import fnmatch
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Iterator

from .config import Config, OrigemCatalogo, resolver

EXTENSOES = {".xlsx", ".xlsm", ".xls", ".csv"}


# --------------------------------------------------------------------------- #
# Modelo
# --------------------------------------------------------------------------- #
@dataclass(frozen=True)
class Item:
    """Um produto do catálogo, com o preço bruto (de lista) da Securiton."""

    part_number: str
    tipo: str
    descricao: str
    preco_bruto: float
    sistema: str
    categoria: str = ""
    origem: str = ""
    linha: int = 0
    fornecedor: str = ""

    @property
    def chave(self) -> str:
        return normalizar_part_number(self.part_number)

    def para_dict(self) -> dict[str, Any]:
        return {
            "part_number": self.part_number,
            "tipo": self.tipo,
            "descricao": self.descricao,
            "preco_bruto": self.preco_bruto,
            "sistema": self.sistema,
            "categoria": self.categoria,
            "origem": self.origem,
            "fornecedor": self.fornecedor,
        }


@dataclass
class Catalogo:
    """Coleção de itens de todas as listas carregadas."""

    itens: list[Item] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)

    def __len__(self) -> int:
        return len(self.itens)

    def __iter__(self) -> Iterator[Item]:
        return iter(self.itens)

    @property
    def sistemas(self) -> list[str]:
        return sorted({i.sistema for i in self.itens})

    def por_part_number(self, part_number: str) -> Item | None:
        """Busca exata por part number (tolerante a espaços e maiúsculas)."""
        alvo = normalizar_part_number(part_number)
        return next((i for i in self.itens if i.chave == alvo), None)

    def por_tipo(self, tipo: str) -> list[Item]:
        alvo = _normalizar(tipo)
        return [i for i in self.itens if _normalizar(i.tipo) == alvo]

    def resolver(self, referencia: str) -> Item | None:
        """Encontra um item por part number ou, na falta, por tipo exato."""
        item = self.por_part_number(referencia)
        if item:
            return item
        candidatos = self.por_tipo(referencia)
        return candidatos[0] if candidatos else None

    def buscar(
        self,
        termo: str = "",
        *,
        sistema: str | None = None,
        categoria: str | None = None,
        limite: int | None = None,
    ) -> list[Item]:
        """Busca livre por termos em part number, tipo, descrição e categoria.

        Todos os termos precisam aparecer (E lógico), sem depender de acento,
        maiúsculas ou da ordem das palavras.
        """
        termos = [_normalizar(t) for t in termo.split() if t.strip()]
        resultado: list[Item] = []
        for item in self.itens:
            if sistema and item.sistema != sistema:
                continue
            if categoria and item.categoria != categoria:
                continue
            if termos:
                alvo = _normalizar(
                    f"{item.part_number} {item.tipo} {item.descricao} {item.categoria} {item.sistema}"
                )
                if not all(t in alvo for t in termos):
                    continue
            resultado.append(item)
            if limite and len(resultado) >= limite:
                break
        return resultado

    def categorias(self, sistema: str | None = None) -> list[str]:
        return sorted(
            {i.categoria for i in self.itens if i.categoria and (not sistema or i.sistema == sistema)}
        )


# --------------------------------------------------------------------------- #
# Normalização
# --------------------------------------------------------------------------- #
def _normalizar(texto: Any) -> str:
    """Minúsculas, sem acento e com espaços colapsados — para comparações."""
    if texto is None:
        return ""
    txt = unicodedata.normalize("NFKD", str(texto))
    txt = "".join(c for c in txt if not unicodedata.combining(c))
    txt = txt.replace("\xa0", " ").replace("\t", " ")
    return re.sub(r"\s+", " ", txt).strip().lower()


def normalizar_part_number(valor: Any) -> str:
    """Part numbers às vezes vêm com espaço ou NBSP sobrando na planilha."""
    return re.sub(r"\s+", "", str(valor or "")).upper()


def limpar_texto(valor: Any) -> str:
    if valor is None:
        return ""
    txt = str(valor).replace("\xa0", " ").replace("\t", " ")
    return re.sub(r"[ ]+", " ", txt).strip()


def converter_preco(valor: Any) -> float | None:
    """Converte o preço da lista (texto '1.398,00' ou '1398.00') em float."""
    if valor is None or valor == "":
        return None
    if isinstance(valor, (int, float)):
        return float(valor)
    txt = str(valor).strip().replace("\xa0", "").replace(" ", "")
    txt = re.sub(r"[^\d,.\-]", "", txt)
    if not txt:
        return None
    if "," in txt and "." in txt:
        # O separador decimal é o que aparece por último.
        if txt.rfind(",") > txt.rfind("."):
            txt = txt.replace(".", "").replace(",", ".")
        else:
            txt = txt.replace(",", "")
    elif "," in txt:
        txt = txt.replace(",", ".")
    try:
        return float(txt)
    except ValueError:
        return None


# --------------------------------------------------------------------------- #
# Leitura dos arquivos
# --------------------------------------------------------------------------- #
def _linhas_da_planilha(caminho: Path) -> tuple[str, list[list[Any]]]:
    """Devolve (nome_da_aba, matriz de valores) do primeiro sheet do arquivo."""
    if caminho.suffix.lower() == ".csv":
        with caminho.open(encoding="utf-8-sig", newline="") as arq:
            amostra = arq.read(4096)
            arq.seek(0)
            try:
                dialeto = csv.Sniffer().sniff(amostra, delimiters=";,\t")
            except csv.Error:
                dialeto = csv.excel
            return caminho.stem, [list(linha) for linha in csv.reader(arq, dialeto)]

    if caminho.suffix.lower() == ".xls":
        import xlrd  # dependência opcional, só usada em listas antigas

        wb = xlrd.open_workbook(caminho)
        sh = wb.sheet_by_index(0)
        return sh.name, [[sh.cell_value(r, c) for c in range(sh.ncols)] for r in range(sh.nrows)]

    import openpyxl

    wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    ws = wb.worksheets[0]
    linhas = [list(linha) for linha in ws.iter_rows(values_only=True)]
    wb.close()
    return ws.title, linhas


def _detectar_cabecalho(
    linhas: list[list[Any]], sinonimos: dict[str, list[str]]
) -> tuple[int, dict[str, int]] | None:
    """Procura a linha de cabeçalho e mapeia cada campo ao índice da coluna."""
    obrigatorios = {"part_number", "preco"}
    for idx, linha in enumerate(linhas[:30]):
        mapa: dict[str, int] = {}
        for col, celula in enumerate(linha):
            texto = _normalizar(celula)
            if not texto:
                continue
            for campo, opcoes in sinonimos.items():
                if campo in mapa:
                    continue
                if any(texto.startswith(_normalizar(op)) for op in opcoes):
                    mapa[campo] = col
                    break
        if obrigatorios <= set(mapa):
            return idx, mapa
    return None


def _titulo_da_lista(linhas: list[list[Any]], ate_linha: int) -> str:
    """O título do sistema fica ao lado de 'Gross Price List in EURO'."""
    marcador = None
    for linha in linhas[:ate_linha]:
        for celula in linha:
            texto = limpar_texto(celula)
            if not texto:
                continue
            if "price list" in _normalizar(texto):
                marcador = True
                continue
            if marcador:
                return texto
    return ""


def ler_arquivo(
    caminho: Path,
    origem: OrigemCatalogo | None = None,
    fornecedor: str = "",
) -> tuple[list[Item], list[str]]:
    """Lê uma lista de preço e devolve (itens, avisos)."""
    origem = origem or OrigemCatalogo()
    avisos: list[str] = []
    nome = caminho.name

    sinonimos = origem.colunas or {
        "descricao": ["description"],
        "tipo": ["type"],
        "part_number": ["part. no.", "part no", "part number"],
        "preco": ["gross price", "price"],
    }

    try:
        aba, linhas = _linhas_da_planilha(caminho)
    except Exception as erro:  # arquivo corrompido, protegido, formato estranho
        return [], [f"{nome}: não foi possível ler ({erro})"]

    sobrescrita = origem.sobrescritas.get(nome, {})
    if sobrescrita.get("colunas"):
        from openpyxl.utils import column_index_from_string

        mapa = {
            campo: column_index_from_string(str(letra).upper()) - 1
            for campo, letra in sobrescrita["colunas"].items()
        }
        inicio = int(sobrescrita.get("linha_cabecalho", 1))
    else:
        detectado = _detectar_cabecalho(linhas, sinonimos)
        if not detectado:
            return [], [f"{nome}: cabeçalho não encontrado (Part No./Gross Price)"]
        idx, mapa = detectado
        inicio = idx + 1

    sistema = (
        origem.sistemas.get(nome)
        or _titulo_da_lista(linhas, inicio)
        or aba
        or caminho.stem
    )

    itens: list[Item] = []
    categoria = ""
    for numero, linha in enumerate(linhas[inicio:], start=inicio + 1):
        def campo(nome_campo: str) -> Any:
            col = mapa.get(nome_campo)
            return linha[col] if col is not None and col < len(linha) else None

        part_number = normalizar_part_number(campo("part_number"))
        preco = converter_preco(campo("preco"))
        descricao = limpar_texto(campo("descricao"))
        tipo = limpar_texto(campo("tipo"))

        if not part_number and preco is None:
            # Linha de categoria: só descrição preenchida.
            if descricao and origem.categorias_por_linha_sem_preco and not tipo:
                categoria = descricao
            continue

        if not part_number:
            avisos.append(f"{nome}, linha {numero}: item sem part number ({descricao!r}) — ignorado")
            continue
        if preco is None:
            avisos.append(f"{nome}, linha {numero}: {part_number} sem preço — ignorado")
            continue

        itens.append(
            Item(
                part_number=part_number,
                tipo=tipo,
                descricao=descricao,
                preco_bruto=preco,
                sistema=sistema,
                categoria=categoria,
                origem=nome,
                linha=numero,
                fornecedor=fornecedor,
            )
        )

    if not itens:
        avisos.append(f"{nome}: nenhum item encontrado")
    return itens, avisos


def _arquivos_de(origem: OrigemCatalogo) -> list[tuple[Path, str]]:
    """Lista os arquivos de lista e o fornecedor de cada um."""
    encontrados: list[tuple[Path, str]] = []
    for fonte in origem.pastas:
        p = resolver(fonte.caminho)
        if not p.is_dir():
            continue
        encontrados.extend(
            (f, fonte.fornecedor)
            for f in sorted(p.iterdir())
            if f.is_file() and f.suffix.lower() in EXTENSOES
        )
    for fonte in origem.arquivos:
        p = resolver(fonte.caminho)
        if p.is_file():
            encontrados.append((p, fonte.fornecedor))

    def ignorado(f: Path) -> bool:
        return any(fnmatch.fnmatch(f.name, padrao) for padrao in origem.ignorar)

    # dedup preservando a ordem
    unicos: dict[Path, str] = {}
    for arquivo, fornecedor in encontrados:
        if not ignorado(arquivo):
            unicos.setdefault(arquivo.resolve(), fornecedor)
    return list(unicos.items())


def carregar(
    config: Config | None = None,
    *,
    caminhos: Iterable[Path | str] | None = None,
) -> Catalogo:
    """Carrega o catálogo completo a partir da configuração (ou de caminhos)."""
    origem = config.catalogo if config else OrigemCatalogo()
    if caminhos:
        arquivos = [(resolver(c), "") for c in caminhos]
    else:
        arquivos = _arquivos_de(origem)

    catalogo = Catalogo()
    vistos: dict[str, Item] = {}
    for arquivo, fornecedor in arquivos:
        itens, avisos = ler_arquivo(arquivo, origem, fornecedor)
        catalogo.avisos.extend(avisos)
        for item in itens:
            anterior = vistos.get(item.chave)
            if anterior is None:
                vistos[item.chave] = item
                catalogo.itens.append(item)
            elif abs(anterior.preco_bruto - item.preco_bruto) > 0.004:
                # Mesmo part number com preços diferentes entre listas: avisa e
                # mantém o primeiro, para o resultado não depender da ordem.
                catalogo.avisos.append(
                    f"{item.part_number}: preço divergente entre {anterior.origem} "
                    f"({anterior.preco_bruto:.2f}) e {item.origem} ({item.preco_bruto:.2f}) "
                    f"— mantido o de {anterior.origem}"
                )
    return catalogo
