"""Trabalhos em PDF: a estrutura é remontada a partir do layout das páginas.

O PDF não tem parágrafos, estilos nem seções: só caracteres desenhados em posições. Como fazem os extratores
de estrutura de documentos científicos (GROBID, CERMINE), a leitura usa o que a página mostra — posição,
fonte, tamanho e peso de cada caractere — para:

1. juntar caracteres em linhas (pela linha de base, aceitando sobrescritos e versaletes);
2. separar cabeçalho, rodapé, número de página e notas de rodapé (pelo fio separador);
3. reconhecer tabelas com fios e regiões de figura (imagens e desenhos), com o texto que está dentro delas;
4. juntar linhas em parágrafos (espaço entre linhas, recuo, linha que termina antes da margem, mudança de
   fonte, início de alínea, legenda ou título) e medir alinhamento, recuos e entrelinhas;
5. reconhecer títulos de seção (numeração, negrito ou caixa alta, tamanho).

O resultado é um .docx interno com essas medidas, sobre o qual rodam as mesmas verificações dos arquivos do
Word. Cada caractere do .docx interno guarda de onde veio na página: as marcações voltam ao PDF no lugar
exato, tanto no visualizador quanto no PDF com notas que o aluno recebe.
"""
import base64
import hashlib
import html
import io
import re
import statistics
import threading
import unicodedata
from collections import Counter, OrderedDict
from dataclasses import dataclass, field

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt

from .cache import qn

CM = 72 / 2.54  # pontos por centímetro
ACENTOS = {"´": "́", "`": "̀", "^": "̂", "ˆ": "̂", "~": "̃", "˜": "̃",
           "¨": "̈", "¸": "̧", "˚": "̊", "ˇ": "̌"}
TROCAS = {" ": " ", " ": " ", " ": " ", "­": "-", "‐": "-", "‑": "-",
          "−": "-", "": "•", "": "▪", "": "•", "●": "•"}
RE_NUM_PAG = re.compile(r"^\s*(?:p[áa]g(?:ina)?\.?\s*)?(\d{1,3}|[ivxlcdm]{1,7}|[IVXLCDM]{1,7})"
                        r"(?:\s*(?:de|/)\s*\d{1,4})?\s*$")
RE_INICIO_PAR = re.compile(r"^(?:[a-z]\)\s|\(?[a-z]\)\s|[•▪●◦]\s|\d{1,2}[.)]\s|"
                           r"(?:Figura|Tabela|Quadro|Gr[áa]fico|Fonte|Nota|Fotografia|Mapa|Esquema)\s*\d*\s*[:.\-–—]?\s)")
RE_TITULO_NUM = re.compile(r"^(\d{1,2}(?:\.\d{1,2}){0,4})\.?\s+\S")
RE_LEGENDA = re.compile(r"^(Figura|Tabela|Quadro|Gr[áa]fico|Fonte|Nota|Fotografia|Mapa|Esquema)\b", re.I)
SEM_NUMERO = {"ERRATA", "AGRADECIMENTOS", "AGRADECIMENTO", "RESUMO", "ABSTRACT", "RESUMEN", "SUMARIO",
              "REFERENCIAS", "REFERENCIAS BIBLIOGRAFICAS", "BIBLIOGRAFIA", "GLOSSARIO", "INDICE",
              "LISTA DE ILUSTRACOES", "LISTA DE FIGURAS", "LISTA DE TABELAS", "LISTA DE QUADROS",
              "LISTA DE GRAFICOS", "LISTA DE ABREVIATURAS E SIGLAS", "LISTA DE ABREVIATURAS", "LISTA DE SIGLAS",
              "LISTA DE SIMBOLOS", "DEDICATORIA", "EPIGRAFE", "APENDICES", "ANEXOS"}


def eh_pdf(dados):
    return isinstance(dados, (bytes, bytearray)) and bytes(dados[:1024]).lstrip().startswith(b"%PDF")


def _chave(s):
    s = unicodedata.normalize("NFD", s)
    return re.sub(r"\s+", " ", "".join(c for c in s if not unicodedata.combining(c)).upper()).strip()


# ====================================================================== caracteres e fontes
def _familia(nome):
    """Nome da fonte no PDF -> (família, negrito, itálico, símbolo)."""
    n = re.sub(r"^[A-Z]{6}\+", "", nome or "")
    b = n.lower().replace(" ", "")
    neg = bool(re.search(r"bold|black|heavy|semibold|demi|medi\b|-medi|cmbx|cmb\d|sfbx|ptmb|phvb|,b\b|-bd\b", b))
    ita = bool(re.search(r"italic|oblique|cmti|cmsl|sfti|sfsl|ptmri|-it\b|,i\b|ital", b))
    simb = bool(re.search(r"symbol|cmsy|cmmi|cmex|msam|msbm|wingding|dingbat|mtsy|stix.*math|cambria.?math|"
                          r"lmmath|euler|rsfs", b))
    if re.search(r"times|tnr|liberationserif|nimbusrom|termes|tinos|^ptm|utm|stixgeneral|stixtwo", b):
        fam = "Times New Roman"
    elif re.search(r"arial|liberationsans|helvetica|nimbussans|heros|arimo|^phv|uhv", b):
        fam = "Arial"
    elif re.search(r"calibri|carlito", b):
        fam = "Calibri"
    elif re.search(r"cambria|caladea", b):
        fam = "Cambria"
    elif re.search(r"^cm[a-z]*\d|^lmroman|^lmsans|^sf[a-z]{2}\d|^cmu|latinmodern", b):
        fam = "Latin Modern (LaTeX)"
    else:
        base = re.split(r"[-,]", n)[0]
        base = re.sub(r"(PS)?MT$|PS$", "", base)
        fam = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", base).strip() or "?"
    return fam, neg, ita, simb


@dataclass(eq=False)
class Car:
    c: str
    x0: float
    x1: float
    top: float
    bottom: float
    size: float
    fonte: str
    negrito: bool
    italico: bool
    simbolo: bool
    pag: int
    sup: bool = False


def _caracteres(page, pag):
    out = []
    for ch in page.chars:
        if not ch.get("upright", True):
            continue
        c = ch.get("text") or ""
        if not c or c == "\x00":
            continue
        c = TROCAS.get(c, c)
        if "ﬀ" <= c <= "ﬆ":
            c = unicodedata.normalize("NFKC", c)
        fam, neg, ita, simb = _familia(ch.get("fontname"))
        tam = round(float(ch.get("size") or 0) * 2) / 2
        out.append(Car(c, ch["x0"], ch["x1"], ch["top"], ch["bottom"], tam, fam, neg, ita, simb, pag))
    return out


# ====================================================================== linhas
@dataclass(eq=False)
class Linha:
    pag: int
    itens: list               # [(caractere, Car)] — espaços inseridos levam um Car no vão
    x0: float = 0
    x1: float = 0
    top: float = 0
    bottom: float = 0
    base: float = 0
    size: float = 0
    fonte: str = ""
    negrito: bool = False
    italico: bool = False
    tipo: str = "corpo"       # corpo | cab | rod | num | nota
    texto: str = ""

    def medir(self):
        cars = [c for _, c in self.itens if _.strip()]
        if not cars:
            return self
        self.x0, self.x1 = min(c.x0 for c in cars), max(c.x1 for c in cars)
        normais = [c for c in cars if not c.simbolo]
        peso = Counter()
        for c in normais or cars:
            peso[c.size] += 1
        self.size = peso.most_common(1)[0][0]
        prin = [c for c in cars if abs(c.size - self.size) < 0.6] or cars
        self.top = min(c.top for c in prin)
        self.bottom = max(c.bottom for c in prin)
        self.base = statistics.median(c.bottom for c in prin)
        for c in cars:
            if c.size < 0.8 * self.size and c.bottom < self.base - 0.15 * self.size:
                c.sup = True
        fontes = Counter(c.fonte for c in normais) if normais else Counter(c.fonte for c in cars)
        self.fonte = fontes.most_common(1)[0][0]
        letras = [c for c in normais if c.c.isalpha()] or normais or cars
        self.negrito = sum(c.negrito for c in letras) >= 0.7 * len(letras)
        self.italico = sum(c.italico for c in letras) >= 0.7 * len(letras)
        self.texto = "".join(s for s, _ in self.itens).strip()
        return self

    @property
    def caixa_alta(self):
        letras = [c for c in self.texto if c.isalpha()]
        return len(letras) >= 3 and all(c.isupper() for c in letras)

    def vao_medio(self):
        vaos = [(c.x1 - c.x0) / (self.size or 12) for s, c in self.itens if s == " "]
        return statistics.median(vaos) if vaos else None

    def esticada(self, normal=0.27):
        """Espaços entre palavras maiores que o normal (linha justificada)."""
        v = self.vao_medio()
        return v is not None and v > normal * 1.1

    def largura_primeira_palavra(self):
        fim = self.x0
        for s, c in self.itens:
            if s.isspace() and fim > self.x0:
                break
            fim = c.x1
        return fim - self.x0


def _juntar_acentos(cars):
    """Acentos desenhados à parte (comum em PDFs do LaTeX) viram o caractere acentuado."""
    usados = set()
    for k, c in enumerate(cars):
        comb = ACENTOS.get(c.c) or (c.c if len(c.c) == 1 and unicodedata.combining(c.c) else None)
        if not comb:
            continue
        meio = (c.x0 + c.x1) / 2
        for j in (k + 1, k - 1, k + 2, k - 2):
            if 0 <= j < len(cars) and j not in usados and cars[j].c.isalpha() and \
                    cars[j].x0 - 0.5 <= meio <= cars[j].x1 + 0.5:
                cars[j].c = unicodedata.normalize("NFC", cars[j].c + comb)
                usados.add(k)
                break
    return [c for k, c in enumerate(cars) if k not in usados]


def _espaco(antes, depois, s):
    return Car(s, antes.x1, max(antes.x1, depois.x0), min(antes.top, depois.top), max(antes.bottom, depois.bottom),
               antes.size, antes.fonte, antes.negrito, antes.italico, False, antes.pag)


def _montar_linha(cars, pag):
    cars = _juntar_acentos(sorted(cars, key=lambda c: (c.x0, c.top)))
    itens, ant, espaco = [], None, 0
    for c in cars:
        if not c.c.strip():
            espaco += ant is not None
            continue
        if ant is not None:
            gap = c.x0 - ant.x1
            if c.c == ant.c and abs(c.x0 - ant.x0) < 0.3 * max(0.5, ant.x1 - ant.x0) and abs(c.top - ant.top) < 1:
                ant.negrito = True        # negrito "falso": o mesmo caractere desenhado duas vezes
                continue
            tam = max(c.size, ant.size)
            if gap > 2.2 * tam:
                itens.append(("\t", _espaco(ant, c, "\t")))
            elif espaco >= 2 and gap > 0.4 * tam:   # dois espaços digitados (o PDF guarda os caracteres)
                vao = _espaco(ant, c, " ")
                itens += [(" ", vao), (" ", vao)]
            elif espaco or gap > 0.15 * tam:
                itens.append((" ", _espaco(ant, c, " ")))
        for s in c.c:
            itens.append((s, c))
        ant, espaco = c, 0
    # sumário e listas: pontilhado até o número da página vira tabulação (como no Word)
    texto = "".join(s for s, _ in itens)
    m = re.search(r"\s*(?:\.\s*){4,}(?=[\divxlcIVXLC]{1,5}\s*$)|\s*…{2,}\s*(?=[\divxlcIVXLC]{1,5}\s*$)", texto)
    if m and m.start() > 0:
        a, z = itens[m.start()][1], itens[m.end() - 1][1]
        vao = Car("\t", a.x0, z.x1, a.top, a.bottom, a.size, a.fonte, a.negrito, a.italico, False, a.pag)
        itens = itens[:m.start()] + [("\t", vao)] + itens[m.end():]
    # sumário e listas sem pontilhado: o número no fim, depois de um vão bem maior que os outros
    seps = [k for k, (s, _) in enumerate(itens) if s in " \t"]
    if seps:
        k = seps[-1]
        resto = "".join(s for s, _ in itens[k + 1:])
        outros = [itens[j][1].x1 - itens[j][1].x0 for j in seps[:-1]]
        vao = itens[k][1].x1 - itens[k][1].x0
        tam = itens[k][1].size or 12
        if re.fullmatch(r"[\divxlcIVXLC]{1,5}", resto) and vao > 1.2 * tam and \
                (not outros or vao > 2.5 * statistics.median(outros)):
            itens[k] = ("\t", itens[k][1])
    return Linha(pag, itens).medir()


def _linhas(cars, pag):
    cars = sorted((c for c in cars if c.c.strip() or c.c == " "), key=lambda c: (c.bottom, c.x0))
    grupos = []
    for c in cars:
        alvo = None
        for g in reversed(grupos[-6:]):
            if abs(c.bottom - g["base"]) <= 0.45 * max(g["size"], c.size, 4):
                alvo = g
                break
        if alvo is None:
            grupos.append({"base": c.bottom, "size": c.size if c.c.strip() else 0, "cars": [c]})
        else:
            alvo["cars"].append(c)
            if c.c.strip() and c.size > alvo["size"]:
                alvo["size"], alvo["base"] = c.size, c.bottom
    linhas = [_montar_linha(g["cars"], pag) for g in grupos if any(c.c.strip() for c in g["cars"])]
    return sorted((ln for ln in linhas if ln.texto), key=lambda ln: (ln.top, ln.x0))


# ====================================================================== páginas
@dataclass(eq=False)
class Tabela:
    pag: int
    bbox: tuple
    linhas_tab: list          # [[texto da célula]]
    cars: list                # [[ [Car] por célula ]]
    bordas: dict
    top: float = 0


@dataclass(eq=False)
class Figura:
    pag: int
    bbox: tuple
    top: float = 0


@dataclass(eq=False)
class Pagina:
    n: int
    largura: float
    altura: float
    linhas: list = field(default_factory=list)
    tabelas: list = field(default_factory=list)
    figuras: list = field(default_factory=list)
    separador: float | None = None   # topo do fio das notas de rodapé
    numero: dict | None = None        # {"texto", "x0", "x1", "top", "bottom", "size", "onde"}
    tem_imagem: bool = False


def _dentro(ln, bbox, folga=2):
    x0, top, x1, bottom = bbox
    meio = (ln.top + ln.bottom) / 2
    return top - folga <= meio <= bottom + folga and ln.x0 >= x0 - folga and ln.x1 <= x1 + folga


def _tabelas(page, n):
    out = []
    try:
        achadas = page.find_tables()
    except Exception:
        return out
    for t in achadas:
        try:
            linhas = t.extract()
        except Exception:
            continue
        if len(linhas) < 2 or max((len(r) for r in linhas), default=0) < 2:
            continue
        x0, top, x1, bottom = t.bbox
        if x1 - x0 < 3 * CM:
            continue
        cars = []
        for row in t.rows:
            linha = []
            for cel in row.cells:
                if cel is None:
                    linha.append([])
                    continue
                cx0, ctop, cx1, cbot = cel
                linha.append([c for c in page.chars if cx0 - 1 <= (c["x0"] + c["x1"]) / 2 <= cx1 + 1
                              and ctop - 1 <= (c["top"] + c["bottom"]) / 2 <= cbot + 1])
            cars.append(linha)
        verticais = [e for e in page.edges if e["orientation"] == "v" and x0 - 2 <= e["x0"] <= x1 + 2
                     and e["top"] < bottom and e["bottom"] > top and e["bottom"] - e["top"] > 4]
        horizontais = [e for e in page.edges if e["orientation"] == "h" and top - 2 <= e["top"] <= bottom + 2
                       and e["x1"] - e["x0"] > 4]
        bordas = {
            "left": any(abs(e["x0"] - x0) < 3 for e in verticais),
            "right": any(abs(e["x0"] - x1) < 3 for e in verticais),
            "insideV": any(x0 + 3 < e["x0"] < x1 - 3 for e in verticais),
            "top": any(abs(e["top"] - top) < 3 for e in horizontais),
            "bottom": any(abs(e["top"] - bottom) < 3 for e in horizontais),
            "insideH": sum(1 for e in horizontais if top + 3 < e["top"] < bottom - 3) >= max(1, len(linhas) - 2),
        }
        out.append(Tabela(n, t.bbox, [[(c or "").strip() for c in r] for r in linhas], cars, bordas, top))
    return out


def _tabelas_abertas(page, n, ja, linhas):
    """Tabelas sem traços verticais (padrão IBGE): fios horizontais alinhados com texto em colunas entre eles."""
    fios = sorted((e for e in page.edges if e["orientation"] == "h" and e["x1"] - e["x0"] > 3 * CM),
                  key=lambda e: e["top"])
    grupos = []
    for e in fios:
        if any(t.bbox[1] - 2 <= e["top"] <= t.bbox[3] + 2 for t in ja):
            continue
        g = next((g for g in grupos if abs(g[0]["x0"] - e["x0"]) < 4 and abs(g[0]["x1"] - e["x1"]) < 4
                  and e["top"] - g[-1]["top"] < 12 * CM), None)
        if g is None:
            grupos.append([e])
        elif e["top"] - g[-1]["top"] > 1:
            g.append(e)
    # fios do mesmo alinhamento podem ser de tabelas diferentes: separa onde, entre dois fios, há texto que
    # não é de tabela (legenda, fonte, parágrafo comum)
    divididos = []
    for g in grupos:
        atual = [g[0]]
        for a, b in zip(g, g[1:]):
            entre = [ln for ln in linhas if a["top"] < (ln.top + ln.bottom) / 2 < b["top"]]
            fora = any(RE_LEGENDA.match(ln.texto) or ("\t" not in ln.texto and len(ln.texto) > 45) for ln in entre)
            if fora:
                divididos.append(atual)
                atual = [b]
            else:
                atual.append(b)
        divididos.append(atual)
    out = []
    for g in divididos:
        if len(g) < 2 or g[-1]["top"] - g[0]["top"] < 0.8 * CM:
            continue
        bbox = (g[0]["x0"], g[0]["top"], g[0]["x1"], g[-1]["top"])
        try:
            area = page.crop((bbox[0] - 1, bbox[1] - 1, bbox[2] + 1, bbox[3] + 1))
            linhas = area.extract_table({"vertical_strategy": "text", "horizontal_strategy": "text",
                                         "snap_tolerance": 3, "join_tolerance": 3, "intersection_tolerance": 3})
        except Exception:
            continue
        linhas = [r for r in (linhas or []) if any((c or "").strip() for c in r)]
        if len(linhas) < 2 or max(len(r) for r in linhas) < 2:
            continue
        cars = [[[c for c in area.chars]] + [[] for _ in r[1:]] if k == 0 else [[] for _ in r]
                for k, r in enumerate(linhas)]
        internos = sum(1 for e in g[1:-1])
        bordas = {"top": True, "bottom": True, "left": False, "right": False, "insideV": False,
                  "insideH": internos >= max(1, len(linhas) - 2)}
        out.append(Tabela(n, bbox, [[(c or "").strip() for c in r] for r in linhas], cars, bordas, bbox[1]))
    return out


def _segmentos(ln, com_fim=False):
    """Pedaços da linha separados por tabulação: [(x0, texto)] (ou [(x0, x1, texto)])."""
    out, atual, x0, x1 = [], "", None, None
    for s, c in ln.itens + [("\t", None)]:
        if s == "\t":
            if atual.strip():
                out.append((x0, x1, atual.strip()) if com_fim else (x0, atual.strip()))
            atual, x0 = "", None
            continue
        if x0 is None and s.strip():
            x0 = c.x0
        if s.strip():
            x1 = c.x1
        atual += s
    return out


def _tabelas_de_texto(linhas, n, R):
    """Tabelas sem fio nenhum: 3+ linhas seguidas com colunas alinhadas (como o modo 'text' do pdfplumber)."""
    out, k = [], 0
    while k < len(linhas):
        j = k
        while j < len(linhas) and (j == k or linhas[j].top - linhas[j - 1].bottom < 2.2 * linhas[j].size):
            sg = _segmentos(linhas[j])
            if len(sg) >= 2:
                j += 1
                continue
            # continuação de uma célula: começa na mesma posição de uma coluna (não a primeira) da linha de cima
            if j > k and j + 1 < len(linhas) and len(_segmentos(linhas[j + 1])) >= 2 and \
                    any(abs(sg[0][0] - x) < 3 for x, _ in _segmentos(linhas[j - 1])[1:]):
                j += 1
                continue
            break
        grupo = linhas[k:j]
        if len(grupo) >= 2 and len(_segmentos(grupo[0])) >= 2:
            if len(grupo) == 2 and [round(x) for x, _ in _segmentos(grupo[0])] != \
                    [round(x) for x, _ in _segmentos(grupo[1])]:
                k = max(j, k + 1)
                continue
            segs = [_segmentos(ln) for ln in grupo]
            # sumário e listas: o número da página alinhado na margem direita
            sumario = sum(1 for ln in grupo if abs(ln.x1 - R) < 4 and re.search(r"[\divxlcIVXLC]$", ln.texto))
            # colunas alinhadas pelo início, pelo fim (números) ou pelo centro
            sfim = [_segmentos(ln, True) for ln in grupo]
            marcas = Counter()
            for sg in sfim:
                for a, z, _ in sg[1:]:
                    marcas.update({("i", round(a / 4)), ("f", round(z / 4)), ("c", round((a + z) / 8))})

            def apoio(tipo, v, marcas=marcas):
                return marcas[(tipo, v)] + marcas[(tipo, v - 1)] + marcas[(tipo, v + 1)]
            alinhadas = sum(1 for sg in sfim if all(max(apoio("i", round(a / 4)), apoio("f", round(z / 4)),
                                                        apoio("c", round((a + z) / 8))) >= 2 for a, z, _ in sg[1:]))
            ncols = max(len(sg) for sg in segs)
            iguais_x1 = max(Counter(round(ln.x1 / 2) for ln in grupo).values()) >= 0.8 * len(grupo)
            num_fim = sum(1 for sg in segs if re.fullmatch(r"[\divxlcIVXLC]{1,5}", sg[-1][1])) >= 0.8 * len(grupo)
            if ncols == 2 and iguais_x1 and num_fim:
                sumario = len(grupo)               # sumário ou lista de ilustrações
            # 2 colunas sem fio nenhum costumam ser listas (siglas, símbolos); só vira tabela com legenda perto
            legenda = any(re.match(r"^(Tabela|Quadro)\s+\d", ln.texto) for ln in linhas
                          if -3 * CM < grupo[0].top - ln.bottom < 2.5 * CM or -3 * CM < ln.top - grupo[-1].bottom < 2.5 * CM)
            if sumario < 0.8 * len(grupo) and alinhadas >= 0.7 * len(grupo) and (ncols >= 3 or legenda):
                cols = sorted({round(x / 4) for sg in segs for x, _ in sg})
                cols_x = []
                for c in cols:
                    if not cols_x or c - cols_x[-1] > 2:
                        cols_x.append(c)
                tab = []
                for sg in segs:
                    if len(sg) == 1 and tab:   # continuação de célula
                        idx = min(range(len(cols_x)), key=lambda i: abs(cols_x[i] - round(sg[0][0] / 4)))
                        tab[-1][idx] = (tab[-1][idx] + " " + sg[0][1]).strip()
                        continue
                    row = [""] * len(cols_x)
                    for x, t in sg:
                        idx = min(range(len(cols_x)), key=lambda i: abs(cols_x[i] - round(x / 4)))
                        row[idx] = (row[idx] + " " + t).strip()
                    tab.append(row)
                bbox = (min(ln.x0 for ln in grupo), grupo[0].top, max(ln.x1 for ln in grupo), grupo[-1].bottom)
                bordas = {x: False for x in ("top", "bottom", "left", "right", "insideH", "insideV")}
                t = Tabela(n, bbox, tab, [[[] for _ in r] for r in tab], bordas, bbox[1])
                for ln in grupo:
                    ln.tipo = "tabela"
                out.append(t)
        k = max(j, k + 1)
    return out


def _figuras(page, n, tabelas):
    """Regiões de imagem e de desenho (gráficos vetoriais), fora das tabelas."""
    W, H = page.width, page.height
    caixas = []
    for im in page.images:
        if im["x1"] - im["x0"] >= 0.8 * CM and im["bottom"] - im["top"] >= 0.8 * CM:
            caixas.append([im["x0"], im["top"], im["x1"], im["bottom"], 3])
    for ob in list(page.curves) + list(page.rects) + list(page.lines):
        x0, top, x1, bottom = ob["x0"], ob["top"], ob["x1"], ob["bottom"]
        if x1 - x0 > 0.85 * W and bottom - top > 0.5 * H:
            continue  # moldura da página
        if any(t.bbox[0] - 2 <= x0 and x1 <= t.bbox[2] + 2 and t.bbox[1] - 2 <= top and bottom <= t.bbox[3] + 2
               for t in tabelas):
            continue
        caixas.append([x0, top, x1, bottom, 1])
    mudou = True
    while mudou:  # agrupa o que se toca
        mudou = False
        k = 0
        while k < len(caixas):
            a = caixas[k]
            j = k + 1
            while j < len(caixas):
                b = caixas[j]
                if a[0] - 6 <= b[2] and b[0] - 6 <= a[2] and a[1] - 6 <= b[3] and b[1] - 6 <= a[3]:
                    a[:] = [min(a[0], b[0]), min(a[1], b[1]), max(a[2], b[2]), max(a[3], b[3]), a[4] + b[4]]
                    caixas.pop(j)
                    mudou = True
                else:
                    j += 1
            k += 1
    return [Figura(n, tuple(c[:4]), c[1]) for c in caixas
            if c[4] >= 3 and c[2] - c[0] >= 1.5 * CM and c[3] - c[1] >= 1.2 * CM]


def _separador_notas(page, L):
    cands = []
    for ob in list(page.lines) + list(page.rects):
        alt = ob["bottom"] - ob["top"]
        larg = ob["x1"] - ob["x0"]
        if alt <= 1.6 and 1.2 * CM <= larg <= 9 * CM and ob["top"] > 0.4 * page.height and \
                (L is None or abs(ob["x0"] - L) < 1.2 * CM):
            cands.append(ob["top"])
    return max(cands) if cands else None


# ====================================================================== leitura
@dataclass(eq=False)
class Paragrafo:
    linhas: list
    nova_pagina: bool = False      # começa numa página nova depois de uma página que não estava cheia
    props: dict = field(default_factory=dict)
    texto: str = ""
    mapa: list = field(default_factory=list)    # Car de cada caractere de .texto
    fatias: list = field(default_factory=list)  # (linha, início, fim) em .texto
    nivel: int | None = None

    @property
    def pag(self):
        return self.linhas[0].pag if self.linhas else getattr(self, "pag_vazia", None)


class LeituraPDF:
    """Leitura completa de um PDF: páginas, blocos remontados (na ordem do .docx interno) e medidas."""

    def __init__(self, dados):
        import pdfplumber
        self.dados = bytes(dados)
        self.paginas = []
        self.tex = False
        cars_total = 0
        with pdfplumber.open(io.BytesIO(self.dados)) as pdf:
            meta = pdf.metadata or {}
            prod = f"{meta.get('Producer', '')} {meta.get('Creator', '')}"
            self.produtor = prod.strip()
            self.tex = bool(re.search(r"tex\b|latex|pdftex|luatex|xetex|overleaf", prod, re.I))
            for n, page in enumerate(pdf.pages):
                pg = Pagina(n, float(page.width), float(page.height))
                cars = _caracteres(page, n)
                cars_total += len(cars)
                pg.tem_imagem = bool(page.images)
                pg.linhas = _linhas(cars, n)
                pg.tabelas = _tabelas(page, n)
                pg.tabelas += _tabelas_abertas(page, n, pg.tabelas, pg.linhas)
                pg.figuras = _figuras(page, n, pg.tabelas)
                pg._objs = page  # usado só durante a leitura
                self.paginas.append(pg)
            self.sem_texto = cars_total < 50
            self._medir()
            for pg in self.paginas:
                pg.separador = _separador_notas(pg._objs, self.L.get(pg.n))
                del pg._objs
        self._classificar()
        self._medir()
        self.blocos = self._paragrafos()
        self._propriedades()
        self._titulos()
        self.docx = self._gerar_docx()

    # ------------------------------------------------------------------ medidas
    def _medir(self):
        corpo = [ln for pg in self.paginas for ln in pg.linhas if ln.tipo == "corpo"]
        peso = Counter()
        for ln in corpo:
            peso[ln.size] += len(ln.texto)
        self.tam_corpo = peso.most_common(1)[0][0] if peso else 12.0
        self.L, self.R = {}, {}

        def moda(valores, minimo):
            """Posição mais frequente (bins de 1 pt, juntando vizinhos) se tiver apoio suficiente."""
            c = Counter(round(v) for v in valores)
            if not c:
                return None
            melhor = max(c, key=lambda x: (c[x] + c[x - 1] + c[x + 1], c[x]))
            if c[melhor] + c[melhor - 1] + c[melhor + 1] < minimo:
                return None
            perto = [v for v in valores if abs(v - melhor) <= 1.5]
            return float(statistics.median(perto))

        def borda_direita(valores):
            """Maior x1 que tenha ao menos 2 outras linhas terminando até 4 pt antes (margem direita)."""
            vs = sorted(valores, reverse=True)
            for k, v in enumerate(vs):
                if sum(1 for w in vs[k + 1:k + 12] if v - w <= 4) >= 2:
                    return float(v)
            return None

        cheias = {pg.n: [ln for ln in pg.linhas if ln.tipo == "corpo" and abs(ln.size - self.tam_corpo) < 1.1
                         and ln.x1 - ln.x0 > 0.45 * pg.largura] for pg in self.paginas}
        larg0 = self.paginas[0].largura if self.paginas else 595
        por_par = {k: [ln for n, v in cheias.items() if n % 2 == k for ln in v] for k in (0, 1)}
        todas = por_par[0] + por_par[1]
        gL = moda([ln.x0 for ln in todas], 1) or 3 * CM
        gR = borda_direita([ln.x1 for ln in todas]) or larg0 - 2 * CM
        parL = {k: moda([ln.x0 for ln in por_par[k]], 5) or gL for k in (0, 1)}
        parR = {k: borda_direita([ln.x1 for ln in por_par[k]]) or gR for k in (0, 1)}
        if abs(parL[0] - parL[1]) <= 0.5 * CM:      # margens iguais nas páginas pares e ímpares
            parL = {0: gL, 1: gL}
            parR = {0: gR, 1: gR}
        for pg in self.paginas:
            v = cheias[pg.n]
            k = pg.n % 2
            L = moda([ln.x0 for ln in v], 3)
            R = moda([ln.x1 for ln in v], 3)
            self.L[pg.n] = L if L is not None and abs(L - parL[k]) < 0.5 * CM else parL[k]
            if R is not None and any(ln.x1 > R + 1.5 for ln in v):
                R = None                            # linhas passam da "coluna": texto alinhado à esquerda
            self.R[pg.n] = R if R is not None and abs(R - parR[k]) < 0.5 * CM else parR[k]
        # entrelinhas típica por tamanho de fonte
        dif = {}
        for pg in self.paginas:
            ls = [ln for ln in pg.linhas if ln.tipo == "corpo"]
            for a, b, c in zip(ls, ls[1:], ls[2:]):
                d1, d2 = b.base - a.base, c.base - b.base
                if abs(a.size - b.size) < 0.6 and abs(b.size - c.size) < 0.6 and abs(d1 - d2) < 1 and \
                        0.95 * a.size <= d1 <= 2.7 * a.size and abs(a.x0 - self.L[pg.n]) < 2.5 * CM:
                    dif.setdefault(a.size, Counter())[round(d1 * 2) / 2] += 1
        self.passo = {s: c.most_common(1)[0][0] for s, c in dif.items() if sum(c.values()) >= 2}

    def passo_de(self, tam):
        if tam in self.passo:
            return self.passo[tam]
        perto = [s for s in self.passo if abs(s - tam) <= 1]
        if perto:
            s = min(perto, key=lambda s: abs(s - tam))
            return self.passo[s] * tam / s
        base = self.passo.get(self.tam_corpo)
        return base * tam / self.tam_corpo if base else 1.5 * 1.15 * tam

    # ------------------------------------------------------------------ cabeçalho, rodapé, notas, tabelas
    def _classificar(self):
        n_pag = len(self.paginas)
        repet = Counter()
        for pg in self.paginas:
            for ln in pg.linhas[:2] + pg.linhas[-2:]:
                if ln.top < 0.14 * pg.altura or ln.bottom > 0.88 * pg.altura:
                    repet[(re.sub(r"\d+", "#", _chave(ln.texto)), round(ln.top / 4))] += 1
        for pg in self.paginas:
            ls = pg.linhas
            if not ls:
                continue
            # número de página: primeira ou última linha, isolada, perto da borda
            for ln, onde in ((ls[0], "topo"), (ls[-1], "base")):
                perto = ln.top < 0.16 * pg.altura if onde == "topo" else ln.bottom > 0.84 * pg.altura
                if not perto:
                    continue
                m = RE_NUM_PAG.match(ln.texto)
                fim = re.search(r"(?:^|\s)(\d{1,3})\s*$", ln.texto)
                if m or (fim and len(ls) > 1 and repet[(re.sub(r"\d+", "#", _chave(ln.texto)),
                                                        round(ln.top / 4))] >= 3):
                    ln.tipo = "num" if m else ("cab" if onde == "topo" else "rod")
                    pg.numero = {"texto": (m.group(1) if m else fim.group(1)), "x0": ln.x0, "x1": ln.x1,
                                 "top": ln.top, "bottom": ln.bottom, "size": ln.size, "onde": onde}
                    break
            for ln in ls:
                if ln.tipo != "corpo":
                    continue
                if (ln.top < 0.14 * pg.altura or ln.bottom > 0.88 * pg.altura) and n_pag >= 4 and \
                        repet[(re.sub(r"\d+", "#", _chave(ln.texto)), round(ln.top / 4))] >= max(3, n_pag // 4) and \
                        ln in (ls[:2] + ls[-2:]):
                    ln.tipo = "cab" if ln.top < 0.5 * pg.altura else "rod"
            if pg.separador is not None:
                for ln in ls:
                    if ln.tipo == "corpo" and ln.top > pg.separador - 1 and ln.size <= self.tam_corpo - 0.9:
                        ln.tipo = "nota"
            for t in pg.tabelas:
                for ln in ls:
                    if ln.tipo == "corpo" and _dentro(ln, t.bbox):
                        ln.tipo = "tabela"
            pg.tabelas += _tabelas_de_texto([ln for ln in ls if ln.tipo == "corpo"], pg.n,
                                            self.R.get(pg.n, pg.largura - 2 * CM))
            # vão largo que sobrou numa linha justificada de poucas palavras: é espaço, não coluna
            for ln in ls:
                if ln.tipo != "corpo" or "\t" not in ln.texto or re.search(r"\t\s*[\divxlcIVXLC]{1,5}$", ln.texto):
                    continue
                vaos = [c.x1 - c.x0 for s_, c in ln.itens if s_ in " \t"]
                cheia = ln.x0 <= self.L.get(pg.n, 0) + 1.5 * CM and ln.x1 >= self.R.get(pg.n, 1e9) - 3
                if cheia and len(vaos) >= 2 and min(vaos) > 0.6 * max(vaos):
                    ln.itens = [(" " if s_ == "\t" else s_, c) for s_, c in ln.itens]
                    ln.medir()
            larg_texto = (self.R.get(pg.n, pg.largura - 2 * CM) - self.L.get(pg.n, 3 * CM)) if hasattr(self, "L") \
                else pg.largura * 0.7
            boas = []
            for f in pg.figuras:
                dentro = [ln for ln in ls if ln.tipo == "corpo" and _dentro(ln, f.bbox, 1)]
                if any(ln.x1 - ln.x0 > 0.7 * larg_texto and ln.size >= self.tam_corpo - 0.6 for ln in dentro):
                    continue  # caixa de texto ou sombreado, não figura
                for ln in dentro:
                    ln.tipo = "figura"
                boas.append(f)
            pg.figuras = boas

    # ------------------------------------------------------------------ parágrafos
    def _centrada(self, ln):
        L, R = self.L[ln.pag], self.R[ln.pag]
        esq, dir_ = ln.x0 - L, R - ln.x1
        destaque = ln.negrito or ln.caixa_alta or abs(ln.size - self.tam_corpo) > 0.9
        tol = max(0.3 * CM, 0.02 * (R - L)) * (2 if destaque else 1)
        if esq < 0.2 * CM or dir_ < 0.2 * CM or abs(esq - dir_) > tol:
            return False
        # linha comum de texto que, por acaso, ficou com folgas parecidas dos dois lados (recuo de 1ª linha)
        return destaque or (ln.x1 - ln.x0) < 0.75 * (R - L)

    def _titulo_provavel(self, ln):
        t = ln.texto
        if "\t" in t or len(t) > 150 or RE_LEGENDA.match(t) or re.search(r"\s[\divxlcIVXLC]{1,4}$", t):
            return False
        k = _chave(re.sub(r"^\W+|\W+$", "", t))
        if k in SEM_NUMERO or re.match(r"^(APENDICE|ANEXO)\s+[A-Z]{1,2}\b", k):
            return ln.negrito or ln.caixa_alta or ln.size >= self.tam_corpo + 1
        m = RE_TITULO_NUM.match(t)
        if not m:
            return False
        nivel = m.group(1).count(".") + 1
        if ln.negrito or ln.caixa_alta or ln.size >= self.tam_corpo + 1:
            return True
        return nivel >= 2 and len(t) <= 100 and not t.rstrip().endswith((".", ";", ","))

    def _quebra(self, p, c, cur, prox=None):
        if abs(c.size - p.size) > 0.8 or c.negrito != p.negrito:
            return True
        if self._titulo_provavel(c):
            return True
        if RE_INICIO_PAR.match(c.texto):
            # "Figura 2", "Tabela 1"... no meio de uma frase ("como mostra a\nTabela 1.") não é legenda
            meio_de_frase = RE_LEGENDA.match(c.texto) and p.texto[-1:] not in ".:;!?)" and \
                abs(c.size - p.size) < 0.6 and not self._centrada(c)
            if not meio_de_frase:
                return True
        if re.match(r"^[-–—]\s", c.texto) and p.texto[-1:] in ".:;":
            return True                            # item de lista com travessão
        if self._refs and p.texto[-1:] == "." and re.match(r"^[A-ZÀ-Ý][A-ZÀ-Ý'’\-]+(?:[ ,.;:]|$)", c.texto) and \
                p.x0 - self.L[p.pag] < 0.3 * CM and c.x0 - self.L[c.pag] < 0.3 * CM and \
                (len(cur) == 1 or abs(c.x0 - cur[0].x0) < 2):
            return True                            # nova referência (SOBRENOME, ... / ENTIDADE)
        if re.search(r"\t\s*[\divxlcIVXLC]{1,5}\s*$", p.texto):
            return True                            # entrada de sumário ou de lista: uma por linha
        cp, cc = self._centrada(p), self._centrada(c)
        if cp != cc:
            return True
        mesma = c.pag == p.pag
        if mesma:
            gap = c.base - p.base
            vizinhos = []
            if len(cur) >= 2 and cur[-2].pag == cur[-1].pag:
                vizinhos.append(cur[-1].base - cur[-2].base)
            if isinstance(prox, Linha) and prox.pag == c.pag and abs(prox.size - c.size) < 0.6 and \
                    0.9 * c.size < prox.base - c.base < 3 * c.size:
                vizinhos.append(prox.base - c.base)   # a distância até a linha seguinte (o passo deste parágrafo)
            if self._refs:
                vizinhos += self._passo_local(c.pag, c.size)
            passo = min(vizinhos) if vizinhos else self.passo_de(p.size)
            if gap < 0:
                return True
            if gap > passo * 1.28 + 1:
                # frase que claramente continua na linha de baixo (sem pontuação final; minúscula, número...)
                continua = gap <= 2.4 * p.size and p.texto[-1:] not in ".:;!?" and \
                    re.match(r"[a-zà-ÿ0-9(\[]", c.texto)
                if not continua:
                    return True
        if cp and cc:
            return not mesma
        R = self.R[p.pag]
        if R - p.x1 > c.largura_primeira_palavra() + 0.3 * c.size:
            return True                            # a linha anterior terminou antes da margem: fim de parágrafo
        if len(cur) >= 2:
            esq = min(ln.x0 - self.L[ln.pag] for ln in cur[1:])
            if abs((c.x0 - self.L[c.pag]) - esq) > 0.35 * CM:
                return True
        elif (c.x0 - self.L[c.pag]) - (p.x0 - self.L[p.pag]) > 0.35 * CM and \
                not RE_INICIO_PAR.match(p.texto) and not re.match(r"^\S{1,4}\s", p.texto):
            return True                            # (alíneas e listas têm recuo deslocado: 2ª linha mais recuada)
        return False

    def _juntar_tabelas(self, blocos):
        """Tabela no fim de uma página que continua no início da seguinte (mesmas colunas): uma tabela só."""
        out = []
        for b in blocos:
            ant = out[-1] if out else None
            if isinstance(b, Tabela) and isinstance(ant, Tabela) and b.pag == ant.pag + 1 and \
                    len(b.linhas_tab[0]) == len(ant.linhas_tab[0]):
                repetido = [c.strip() for c in b.linhas_tab[0]] == [c.strip() for c in ant.linhas_tab[0]]
                ant.linhas_tab += b.linhas_tab[1:] if repetido else b.linhas_tab
                ant.cars += b.cars[1:] if repetido else b.cars
                ant.partes = getattr(ant, "partes", [ant.bbox and (ant.pag, ant.bbox)]) + [(b.pag, b.bbox)]
                ant.cabecalho_repetido = repetido
                continue
            out.append(b)
        return out

    def _passo_local(self, pag, tam):
        """Menor distância entre linhas (do mesmo tamanho) que se repete na página."""
        chave = (pag, tam)
        if chave not in self._passos:
            ls = [ln for ln in self.paginas[pag].linhas if ln.tipo == "corpo" and abs(ln.size - tam) < 0.6]
            c = Counter(round((b.base - a.base) * 2) / 2 for a, b in zip(ls, ls[1:]) if b.base - a.base >= 0.95 * tam)
            freq = [d for d, n in c.items() if n >= 2]
            self._passos[chave] = [min(freq)] if freq else []
        return self._passos[chave]

    def _paragrafos(self):
        blocos, cur = [], []
        ultimo, cheia_ant = None, True
        self._refs, self._passos = False, {}

        def fechar():
            nonlocal cur
            if cur:
                blocos.append(Paragrafo(cur, nova_pagina=cur[0]._np))
            cur = []

        for pg in self.paginas:
            itens = sorted([ln for ln in pg.linhas if ln.tipo == "corpo"] + pg.tabelas + pg.figuras,
                           key=lambda x: x.top)
            if not itens:
                fechar()
                vazio = Paragrafo([], nova_pagina=True)
                vazio.pag_vazia = pg.n
                blocos.append(vazio)
                ultimo, cheia_ant = None, False
                continue
            for k, x in enumerate(itens):
                np_ = k == 0 and not cheia_ant
                if isinstance(x, Linha):
                    x._np = np_
                    k_ = _chave(re.sub(r"^\W+|\W+$", "", x.texto))
                    if re.match(r"^(REFERENCIAS|BIBLIOGRAFIA|OBRAS CITADAS)\b", k_) and len(k_) < 40:
                        self._refs = True
                    elif self._refs and (re.match(r"^(APENDICE|ANEXO|GLOSSARIO|INDICE)\b", k_) and len(k_) < 90):
                        self._refs = False
                    prox = itens[k + 1] if k + 1 < len(itens) else None
                    if not cur or not isinstance(ultimo, Linha) or np_ or self._quebra(ultimo, x, cur, prox):
                        fechar()
                        cur = [x]
                    else:
                        cur.append(x)
                else:
                    fechar()
                    x.nova_pagina = np_
                    blocos.append(x)
                ultimo = x
            fim = max(x.bottom if isinstance(x, Linha) else x.bbox[3] for x in itens)
            cheia_ant = fim > pg.altura - 2 * CM - 3 * self.passo_de(self.tam_corpo)
        fechar()
        return self._juntar_tabelas(blocos)

    # ------------------------------------------------------------------ medidas de cada parágrafo
    def _entrelinhas(self, passo, tam, fonte):
        if self.tex:
            r = passo / (1.2 * tam)
            for lo, hi, v in ((0.93, 1.09, 1.0), (1.17, 1.31, 1.5), (1.55, 1.75, 2.0)):
                if lo <= r <= hi:
                    return v
            return round(r, 2)
        fator = 1.22 if "Calibri" in fonte else 1.15
        r = passo / (fator * tam)
        for v in (1.0, 1.15, 1.5, 2.0, 2.5, 3.0):
            if abs(r - v) <= 0.06 * v:
                return v
        return round(r, 2)

    def _justificado(self, ls):
        """Linhas (menos a última) até a margem direita. Uma ou outra pode ter quebra de linha manual, que no
        Word não é esticada; com duas linhas só, a primeira precisa ter os espaços esticados."""
        corpo = ls[:-1]
        cheias = [ln for ln in corpo if ln.x1 >= self.R[ln.pag] - 2.5]
        if len(corpo) == 1:   # (o LaTeX justifica sempre, encolhendo ou esticando os espaços)
            return bool(cheias) and (self.tex or cheias[0].esticada(self.vao_normal))
        return len(cheias) >= max(2, 0.6 * len(corpo)) and any(ln.esticada(self.vao_normal) for ln in cheias)

    def _propriedades(self):
        # espaço normal entre palavras: o das últimas linhas dos parágrafos (nunca esticadas)
        ultimas = [b.linhas[-1].vao_medio() for b in self.blocos if isinstance(b, Paragrafo) and len(b.linhas) >= 2
                   and b.linhas[-1].x1 < self.R[b.linhas[-1].pag] - 2 * CM]
        ultimas = [v for v in ultimas if v]
        self.vao_normal = statistics.median(ultimas) if len(ultimas) >= 3 else 0.27
        ant = None
        for b in self.blocos:
            if not isinstance(b, Paragrafo) or not b.linhas:
                ant = b if not isinstance(b, Paragrafo) else None
                continue
            ls = b.linhas
            L0 = self.L[ls[0].pag]
            rel = [ln.x0 - self.L[ln.pag] for ln in ls]
            props = {}
            if all(self._centrada(ln) for ln in ls):
                props["al"] = WD_ALIGN_PARAGRAPH.CENTER
            elif len(ls) >= 2 and self._justificado(ls):
                props["al"] = WD_ALIGN_PARAGRAPH.JUSTIFY
            elif all(abs(ln.x1 - self.R[ln.pag]) < 2.5 for ln in ls) and rel[0] > 2 * CM:
                props["al"] = WD_ALIGN_PARAGRAPH.RIGHT
            elif len(ls) >= 2:
                props["al"] = WD_ALIGN_PARAGRAPH.LEFT
            if props.get("al") is None and len(ls) == 1 and self._titulo_provavel(ls[0]) and rel[0] < 0.3 * CM:
                props["al"] = WD_ALIGN_PARAGRAPH.LEFT
            if props.get("al") == WD_ALIGN_PARAGRAPH.CENTER:
                props["esq"], props["pri"] = 0.0, 0.0
            elif len(ls) >= 2:
                esq = min(rel[1:])
                props["esq"], props["pri"] = max(0.0, esq), rel[0] - max(0.0, esq)
            else:
                off = rel[0]
                if 0.9 * CM <= off <= 2.6 * CM and not self._titulo_provavel(ls[0]):
                    props["esq"], props["pri"] = 0.0, off
                else:
                    props["esq"], props["pri"] = max(0.0, off), 0.0
                if "\t" in ls[0].texto:
                    props["al"] = WD_ALIGN_PARAGRAPH.LEFT
            difs = [b2.base - a.base for a, b2 in zip(ls, ls[1:]) if a.pag == b2.pag and b2.base > a.base]
            passo = statistics.median(difs) if difs else self.passo_de(ls[0].size)
            props["linha"] = self._entrelinhas(passo, ls[0].size, ls[0].fonte)
            props["linha_medida"] = bool(difs)
            antes = 0.0
            if isinstance(ant, Paragrafo) and ant.linhas and ant.linhas[-1].pag == ls[0].pag:
                antes = max(0.0, ls[0].base - ant.linhas[-1].base - passo)
            elif ant is not None and not isinstance(ant, Paragrafo) and ant.pag == ls[0].pag:
                antes = max(0.0, ls[0].top - ant.bbox[3] - (passo - ls[0].size))
            props["antes"] = round(antes) if antes >= 3 else 0
            props["L0"] = L0
            b.props = props
            ant = b
        # parágrafos de uma linha: entrelinhas do parágrafo medido mais próximo, com a mesma fonte e recuo
        pars = [b for b in self.blocos if isinstance(b, Paragrafo) and b.linhas]
        for k, b in enumerate(pars):
            if b.props.get("linha_medida"):
                continue
            for d in range(1, 7):
                viz = [pars[j] for j in (k - d, k + d) if 0 <= j < len(pars)]
                ok = [v for v in viz if v.props.get("linha_medida") and abs(v.linhas[0].size - b.linhas[0].size) < 0.6
                      and abs(v.props["esq"] - b.props["esq"]) < 0.3 * CM and v.nivel is None]
                if ok:
                    b.props["linha"] = ok[0].props["linha"]
                    b.props["linha_medida"] = "vizinho"
                    break
        tipica = {}
        for b in self.blocos:
            if isinstance(b, Paragrafo) and b.linhas and b.props.get("linha_medida") is True:
                tipica.setdefault(b.linhas[0].size, Counter())[b.props["linha"]] += len(b.linhas)
        for b in self.blocos:
            if isinstance(b, Paragrafo) and b.linhas and not b.props.get("linha_medida"):
                c = tipica.get(b.linhas[0].size) or tipica.get(self.tam_corpo)
                if c:
                    b.props["linha"] = c.most_common(1)[0][0]
        # parágrafos de uma linha sem alinhamento definido: o do parágrafo vizinho parecido
        ultimo = {}
        for b in self.blocos:
            if not isinstance(b, Paragrafo) or not b.linhas:
                continue
            k = (b.linhas[0].size, round(b.props["esq"] / CM, 1))
            if b.props.get("al") is not None and len(b.linhas) >= 2:
                ultimo[k] = b.props["al"]
            elif b.props.get("al") is None:
                b.props["al"] = ultimo.get(k, WD_ALIGN_PARAGRAPH.LEFT)
        # texto com o mapa de cada caractere
        for b in self.blocos:
            if isinstance(b, Paragrafo):
                self._texto(b)

    def _texto(self, b):
        itens = []
        fatias = []
        for k, ln in enumerate(b.linhas):
            ini = len(itens)
            li = list(ln.itens)
            while li and li[0][0].isspace():
                li.pop(0)
            while li and li[-1][0].isspace():
                li.pop()
            if itens:
                ant = itens[-1]
                hifen = ant[0] == "-" and li and li[0][0].islower() and len(itens) >= 2 and itens[-2][0].isalpha() \
                    and b.linhas[k - 1].x1 >= self.R[b.linhas[k - 1].pag] - 3
                if hifen and self.tex:   # LaTeX hifeniza sozinho: a palavra é juntada
                    itens.pop()
                    fatias[-1] = (fatias[-1][0], fatias[-1][1], len(itens))
                elif hifen:              # no Word, hífen no fim da linha é de palavra composta: fica, sem espaço
                    pass
                else:
                    itens.append((" ", _espaco(ant[1], li[0][1] if li else ant[1], " ")))
                    fatias[-1] = (fatias[-1][0], fatias[-1][1], len(itens))
                ini = len(itens)
            itens += li
            fatias.append((ln, ini, len(itens)))
        b.texto = "".join(s for s, _ in itens)
        b.mapa = [c for _, c in itens]
        b.fatias = fatias

    def _titulos(self):
        for b in self.blocos:
            if not isinstance(b, Paragrafo) or not b.linhas or len(b.linhas) > 3:
                continue
            ln = b.linhas[0]
            if not self._titulo_provavel(ln):
                continue
            if any(x.negrito != ln.negrito or abs(x.size - ln.size) > 0.6 for x in b.linhas[1:]):
                continue
            m = RE_TITULO_NUM.match(b.texto)
            b.nivel = (m.group(1).count(".") + 1) if m else 1

    # ------------------------------------------------------------------ .docx interno
    def _gerar_docx(self):
        doc = Document()
        st = doc.styles["Normal"]
        st.paragraph_format.space_after = Pt(0)
        st.paragraph_format.space_before = Pt(0)
        st.paragraph_format.line_spacing = 1.0
        st.font.name = "Times New Roman"
        st.font.size = Pt(self.tam_corpo)
        sec = doc.sections[0]
        pg0 = self.paginas[0] if self.paginas else None
        if pg0:
            sec.page_width, sec.page_height = Pt(pg0.largura), Pt(pg0.altura)
            m = self.margens_medidas()
            sec.left_margin, sec.right_margin = Cm(m["esq"]), Cm(m["dir"])
            sec.top_margin, sec.bottom_margin = Cm(m["sup"]), Cm(m["inf"])
        pag_atual = 0
        for b in self.blocos:
            if isinstance(b, Paragrafo):
                p = doc.add_paragraph()
                if not b.linhas:
                    r = p.add_run()
                    r._r.append(OxmlElement("w:lastRenderedPageBreak"))
                    pag_atual += 1
                    continue
                pf = p.paragraph_format
                pf.alignment = b.props["al"]
                if b.props["esq"] > 0.15 * CM:
                    pf.left_indent = Pt(b.props["esq"])
                if abs(b.props["pri"]) > 0.15 * CM:
                    pf.first_line_indent = Pt(b.props["pri"])
                pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
                pf.line_spacing = b.props["linha"]
                pf.space_before = Pt(b.props["antes"])
                pf.space_after = Pt(0)
                if b.nova_pagina:
                    pf.page_break_before = True
                if b.nivel:
                    ppr = p._p.get_or_add_pPr()
                    ol = OxmlElement("w:outlineLvl")
                    ol.set(qn("w:val"), str(min(8, b.nivel - 1)))
                    ppr.append(ol)
                pag_atual = self._runs(p, b.texto, b.mapa, pag_atual)
            elif isinstance(b, Tabela):
                pag_atual = self._tabela(doc, b, pag_atual)
            else:
                p = doc.add_paragraph()
                p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                if b.pag != pag_atual:
                    r = p.add_run()
                    r._r.append(OxmlElement("w:lastRenderedPageBreak"))
                    pag_atual = b.pag
                x0, top, x1, bottom = b.bbox
                p.add_run().add_picture(io.BytesIO(_PNG), width=Pt(max(1, x1 - x0)), height=Pt(max(1, bottom - top)))
        buf = io.BytesIO()
        doc.save(buf)
        return buf.getvalue()

    def _runs(self, p, texto, mapa, pag_atual):
        k = 0
        while k < len(texto):
            c = mapa[k]
            chave = (c.fonte, c.size, c.negrito, c.italico, c.sup, c.pag)
            j = k + 1
            while j < len(texto):
                d = mapa[j]
                if texto[j] == " " and j + 1 < len(texto) and mapa[j + 1].pag == c.pag:
                    j += 1
                    continue
                if (d.fonte, d.size, d.negrito, d.italico, d.sup, d.pag) != chave:
                    break
                j += 1
            r = p.add_run(texto[k:j])
            if c.pag != pag_atual:
                for _ in range(c.pag - pag_atual):
                    r._r.insert(1 if r._r.rPr is not None else 0, OxmlElement("w:lastRenderedPageBreak"))
                pag_atual = c.pag
            f = r.font
            f.name = c.fonte
            rpr = r._r.get_or_add_rPr()
            fontes = rpr.find(qn("w:rFonts"))
            if fontes is not None:
                for a in ("w:hAnsi", "w:cs", "w:eastAsia"):
                    fontes.set(qn(a), c.fonte)
            f.size = Pt(c.size if not c.sup else c.size)
            f.bold = c.negrito
            f.italic = c.italico
            if c.sup:
                f.superscript = True
            k = j
        return pag_atual

    def _tabela(self, doc, t, pag_atual):
        nlin = len(t.linhas_tab)
        ncol = max(len(r) for r in t.linhas_tab)
        tab = doc.add_table(rows=nlin, cols=ncol)
        for i, row in enumerate(t.linhas_tab):
            for j in range(ncol):
                cel = tab.cell(i, j)
                txt = row[j] if j < len(row) else ""
                cars = t.cars[i][j] if i < len(t.cars) and j < len(t.cars[i]) else []
                p = cel.paragraphs[0]
                r = p.add_run(re.sub(r"\s+", " ", txt))
                if i == 0 and j == 0 and t.pag != pag_atual:
                    r._r.insert(0, OxmlElement("w:lastRenderedPageBreak"))
                    pag_atual = t.pag
                if cars:
                    fam, neg, ita, _ = _familia(Counter(c["fontname"] for c in cars).most_common(1)[0][0])
                    r.font.name = fam
                    r.font.size = Pt(round(statistics.median(c["size"] for c in cars) * 2) / 2)
                    r.font.bold = neg
                    r.font.italic = ita
        tblpr = tab._tbl.tblPr
        bordas = OxmlElement("w:tblBorders")
        for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
            e = OxmlElement(f"w:{lado}")
            e.set(qn("w:val"), "single" if t.bordas.get(lado) else "nil")
            if t.bordas.get(lado):
                e.set(qn("w:sz"), "4")
                e.set(qn("w:color"), "000000")
            bordas.append(e)
        tblpr.append(bordas)
        return pag_atual

    # ------------------------------------------------------------------ medidas para as verificações
    def margens_medidas(self):
        """Margens estimadas (cm) pela posição do texto nas páginas com texto corrido."""
        if not self.paginas:
            return {"esq": 3.0, "dir": 2.0, "sup": 3.0, "inf": 2.0}
        W = self.paginas[0].largura
        esq = Counter(round(self.L[p.n] / CM, 1) for p in self.paginas).most_common(1)[0][0]
        dir_ = Counter(round((p.largura - self.R[p.n]) / CM, 1) for p in self.paginas).most_common(1)[0][0]
        tops, bots = [], []
        for p in self.paginas:
            corpo = [ln for ln in p.linhas if ln.tipo in ("corpo", "nota")]
            if len(corpo) >= 8:
                tops.append(corpo[0].top)
                bots.append(p.altura - corpo[-1].bottom)
        sup = round(min(tops) / CM, 1) if tops else 3.0
        inf = round(min(bots) / CM, 1) if bots else 2.0
        return {"esq": esq, "dir": dir_ if W else 2.0, "sup": sup, "inf": inf}

    @property
    def espelhado(self):
        """Margens espelhadas (frente e verso): páginas pares e ímpares com a margem esquerda trocada."""
        imp = [self.L[p.n] for p in self.paginas if p.n % 2 == 0]
        par = [self.L[p.n] for p in self.paginas if p.n % 2 == 1]
        return bool(imp and par) and abs(statistics.median(imp) - statistics.median(par)) > 0.5 * CM

    def lado_trocado(self, n):
        """True se a página n tem a margem maior à direita (verso, nas margens espelhadas)."""
        if not self.espelhado:
            return False
        outra = [self.L[p.n] for p in self.paginas if p.n % 2 != n % 2]
        return self.L[n] < statistics.median(outra) - 0.5 * CM

    def pagina_do_bloco(self, i):
        b = self.blocos[i] if 0 <= i < len(self.blocos) else None
        if isinstance(b, Paragrafo):
            return b.linhas[0].pag if b.linhas else getattr(b, "pag_vazia", None)
        return getattr(b, "pag", None)


def _png():
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (8, 8), (235, 240, 245)).save(buf, "PNG")
    return buf.getvalue()


_PNG = _png()

# ====================================================================== cache da leitura
_LEITURAS = OrderedDict()
_TRAVA = threading.Lock()


def ler(dados):
    """LeituraPDF dos bytes (guardada: a análise, o visualizador e as notas usam a mesma leitura)."""
    chave = hashlib.sha1(bytes(dados)).hexdigest()
    with _TRAVA:
        if chave in _LEITURAS:
            _LEITURAS.move_to_end(chave)
            return _LEITURAS[chave]
    lt = LeituraPDF(dados)
    with _TRAVA:
        _LEITURAS[chave] = lt
        while len(_LEITURAS) > 4:
            _LEITURAS.popitem(last=False)
    return lt


# ====================================================================== visualizador
def html_paginas(lt, escala=1.2, qualidade=60):
    """Páginas do PDF como imagem, com uma camada de texto transparente por cima (como no pdf.js): as
    marcações, a busca e a seleção funcionam sobre o texto, e o aluno vê a página exatamente como ela é."""
    import pypdfium2 as pdfium
    por_pag = {}
    for i, b in enumerate(lt.blocos):
        if isinstance(b, Paragrafo):
            for ln, ini, fim in b.fatias:
                por_pag.setdefault(ln.pag, []).append(
                    f"<p class='lin' data-b='{i}' data-w='{ln.x1 - ln.x0:.1f}' style='left:{ln.x0:.1f}pt;"
                    f"top:{ln.top:.1f}pt;font-size:{ln.size:.1f}pt;height:{ln.bottom - ln.top:.1f}pt'>"
                    f"{html.escape(b.texto[ini:fim])}</p>")
        elif isinstance(b, (Tabela, Figura)):
            x0, top, x1, bottom = b.bbox
            if isinstance(b, Tabela):
                for ln in lt.paginas[b.pag].linhas:
                    if ln.tipo == "tabela" and _dentro(ln, b.bbox):
                        por_pag.setdefault(b.pag, []).append(
                            f"<p class='lin' data-b='{i}' data-w='{ln.x1 - ln.x0:.1f}' style='left:{ln.x0:.1f}pt;"
                            f"top:{ln.top:.1f}pt;font-size:{ln.size:.1f}pt;height:{ln.bottom - ln.top:.1f}pt'>"
                            f"{html.escape(ln.texto)} </p>")
            else:
                por_pag.setdefault(b.pag, []).append(
                    f"<div class='fig-pdf' data-b='{i}' style='left:{x0:.1f}pt;top:{top:.1f}pt;"
                    f"width:{x1 - x0:.1f}pt;height:{bottom - top:.1f}pt'></div>")
    doc = pdfium.PdfDocument(lt.dados)
    partes = []
    from .perfil import atual
    M = atual()["margens"]
    try:
        for n, pg in enumerate(lt.paginas):
            img = doc[n].render(scale=escala).to_pil().convert("RGB")
            buf = io.BytesIO()
            img.save(buf, "JPEG", quality=qualidade, optimize=True)
            src = "data:image/jpeg;base64," + base64.b64encode(buf.getvalue()).decode()
            esq, dir_ = (M["direita"], M["esquerda"]) if lt.lado_trocado(n) else (M["esquerda"], M["direita"])
            partes.append(
                f"<section class='folha pdf' data-sec='{n + 1}' data-paginacao='word' "
                f"style='width:{pg.largura:.1f}pt;height:{pg.altura:.1f}pt;min-height:{pg.altura:.1f}pt'>"
                f"<div class='rotulo-sec'>Página {n + 1} do PDF · guia das margens {_fmt(esq)}/{_fmt(dir_)}/"
                f"{_fmt(M['superior'])}/{_fmt(M['inferior'])} cm</div>"
                f"<img class='pg' alt='' src='{src}'>"
                f"<div class='guia-margens' style='left:{esq}cm;right:{dir_}cm;top:{M['superior']}cm;"
                f"bottom:{M['inferior']}cm'></div>"
                f"<div class='area'>{''.join(por_pag.get(n, []))}</div></section>")
    finally:
        doc.close()
    return "".join(partes)


def _fmt(v):
    return f"{v:g}".replace(".", ",")


# ====================================================================== PDF com notas
CORES = {"Ajustar": (0.98, 0.55, 0.55), "Verificar": (1.0, 0.86, 0.3)}


def _quads(lt, i, m):
    """Retângulos (pág., x0, top, x1, bottom) por linha do trecho 'm' no bloco i (ou do bloco todo)."""
    b = lt.blocos[i] if 0 <= i < len(lt.blocos) else None
    if isinstance(b, Paragrafo) and b.linhas:
        ini, fim = 0, len(b.texto)
        if m:
            pos = b.texto.find(m)
            if pos >= 0:
                ini, fim = pos, pos + len(m)
        out = []
        for ln, a, z in b.fatias[:14] if not m else b.fatias:
            a2, z2 = max(a, ini), min(z, fim)
            cars = [c for c in b.mapa[a2:z2] if c is not None]
            while cars and cars[-1].c.isspace():
                cars.pop()
            if not cars:
                continue
            out.append((ln.pag, min(c.x0 for c in cars), ln.top, max(c.x1 for c in cars), ln.bottom))
        return out, False
    if isinstance(b, (Tabela, Figura)):
        x0, top, x1, bottom = b.bbox
        return [(b.pag, x0, top, x1, bottom)], True
    return [], False


def anotar(dados, lt, serial, texto_comentario, coletivos, limite_coletivo, max_por_item):
    """PDF original com destaques e notas (como no Acrobat). Retorna (bytes, número de notas)."""
    from pypdf import PdfReader, PdfWriter
    from pypdf.generic import ArrayObject, DictionaryObject, FloatObject, NameObject, NumberObject, TextStringObject

    w = PdfWriter(clone_from=PdfReader(io.BytesIO(dados)))
    n = 0

    def num(v):
        return FloatObject(round(v, 2))

    def caixa(pag):
        mb = w.pages[pag].mediabox
        return float(mb.left), float(mb.top)

    def destaque(pag, rets, cor, texto, quadrado=False):
        ox, oy = caixa(pag)
        x0 = min(r[0] for r in rets)
        x1 = max(r[2] for r in rets)
        top = min(r[1] for r in rets)
        bot = max(r[3] for r in rets)
        quads = []
        for a, t, z, bo in rets:
            quads += [a, t, z, t, a, bo, z, bo]
        rect = [ox + x0, oy - bot, ox + x1, oy - top]
        anot = DictionaryObject({
            NameObject("/Type"): NameObject("/Annot"),
            NameObject("/Subtype"): NameObject("/Square" if quadrado else "/Highlight"),
            NameObject("/Rect"): ArrayObject([num(v) for v in rect]),
            NameObject("/C"): ArrayObject([num(v) for v in cor]),
            NameObject("/T"): TextStringObject("Biblioteca EESC (verificador)"),
            NameObject("/Contents"): TextStringObject(texto),
            NameObject("/F"): NumberObject(4),
            NameObject("/CA"): FloatObject(0.55 if not quadrado else 1),
        })
        if not quadrado:
            pts = []
            for k in range(0, len(quads), 2):
                pts += [num(ox + quads[k]), num(oy - quads[k + 1])]
            anot[NameObject("/QuadPoints")] = ArrayObject(pts)
        else:
            anot[NameObject("/BS")] = DictionaryObject({NameObject("/W"): NumberObject(2)})
        w.add_annotation(pag, anot)

    def nota(pag, x, top, cor, texto):
        ox, oy = caixa(pag)
        anot = DictionaryObject({
            NameObject("/Type"): NameObject("/Annot"),
            NameObject("/Subtype"): NameObject("/Text"),
            NameObject("/Rect"): ArrayObject([num(ox + x), num(oy - top - 18), num(ox + x + 18), num(oy - top)]),
            NameObject("/C"): ArrayObject([num(v) for v in cor]),
            NameObject("/T"): TextStringObject("Biblioteca EESC (verificador)"),
            NameObject("/Contents"): TextStringObject(texto),
            NameObject("/Name"): NameObject("/Comment"),
            NameObject("/F"): NumberObject(4),
        })
        w.add_annotation(pag, anot)

    gerais = []
    for a in serial:
        if a["status"] not in CORES:
            continue
        locais = a["locais"][:max_por_item]
        if not locais:
            gerais.append(a)
            continue
        linhas = texto_comentario(a)
        dica = next((v for k, v in coletivos.items() if a["item"].startswith(k)), None)
        if dica and len(a["locais"]) > limite_coletivo:
            linhas.insert(1, f"Este problema se repete em {len(a['locais'])} trechos do documento; marcado só o primeiro.")
            locais = locais[:1]
        elif len(a["locais"]) > max_por_item:
            linhas.append(f"(Este problema aparece em {len(a['locais'])} lugares; marcados os {max_por_item} primeiros.)")
        cor = CORES[a["status"]]
        for k, loc in enumerate(locais):
            if not isinstance(loc["b"], int):
                continue
            rets, quadrado = _quads(lt, loc["b"], loc.get("m"))
            if not rets:
                continue
            texto = "\n".join(linhas) if k == 0 else \
                f"{linhas[0]} (ocorrência {k + 1} de {len(locais)}; veja a primeira nota deste item)"
            por_pag = {}
            for pag, x0, top, x1, bottom in rets:
                por_pag.setdefault(pag, []).append((x0, top, x1, bottom))
            for pag, rs in por_pag.items():
                destaque(pag, rs, cor, texto, quadrado)
                texto = f"{linhas[0]} (continuação)"
            n += 1
    if gerais:
        corpo = ["PONTOS GERAIS DO DOCUMENTO (não ligados a um trecho específico):"]
        for a in gerais:
            corpo.append("")
            corpo += texto_comentario(a)
        pag0 = next((b.linhas[0].pag for b in lt.blocos if isinstance(b, Paragrafo) and b.linhas), 0)
        nota(pag0, 8, 8, CORES["Verificar"], "\n".join(corpo))
        n += 1
    buf = io.BytesIO()
    w.write(buf)
    return buf.getvalue(), n
