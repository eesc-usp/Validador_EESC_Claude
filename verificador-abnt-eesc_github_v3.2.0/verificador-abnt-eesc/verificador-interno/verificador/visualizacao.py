"""Mostra o .docx como páginas (somente leitura) e liga cada achado ao trecho do documento.

- Cada parágrafo/tabela recebe data-b="<índice do bloco>" (o mesmo índice usado pelo Documento), para o
  visualizador rolar até ele e marcar o trecho.
- As páginas seguem a paginação que o próprio Word grava ao salvar (w:lastRenderedPageBreak): a página 12
  daqui é a página 12 que o aluno viu. Arquivos sem essa informação (Google Docs, outros programas) são
  paginados no navegador pela altura do conteúdo.
- Nada aqui altera o documento: a página é só para leitura.
"""
import base64
import html
import re

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.run import Run

from . import cache
from . import docx_util as du
from .cache import qn
from .documento import Documento, chave
from .modelo import OK

ALINHA = {WD_ALIGN_PARAGRAPH.CENTER: "center", WD_ALIGN_PARAGRAPH.RIGHT: "right",
          WD_ALIGN_PARAGRAPH.JUSTIFY: "justify", WD_ALIGN_PARAGRAPH.DISTRIBUTE: "justify"}
LIMITE_IMAGEM = 1_500_000  # bytes; imagens maiores viram um marcador
class _Quebra:
    """Marca de nova página entre os pedaços de um parágrafo/tabela."""

    def __init__(self, manual):
        self.manual = manual  # True: quebra posta pelo aluno · False: paginação gravada pelo Word


QUEBRA = _Quebra(True)       # quebra de página, "quebra antes", quebra de seção
PAGINA_WORD = _Quebra(False)  # onde o Word começou uma página nova da última vez que salvou
MIN_MARCAS_WORD = 5  # mesmo critério de elementos.paginas(): menos que isso, o arquivo não traz a paginação


def _esc(s):
    return html.escape(s, quote=True)


def _intercalar(pedacos, quebras):
    out = [pedacos[0]]
    for q, x in zip(quebras, pedacos[1:]):
        out += [q, x]
    return out


# ====================================================================== runs
class _Render:
    def __init__(self, d):
        self.d = d
        self.doc = d.doc
        self.partes = []
        corpo = self.doc.element.body
        self.paginar_word = len(corpo.findall(".//" + qn("w:lastRenderedPageBreak"))) >= MIN_MARCAS_WORD
        self.pagina_vazia = True
        self.sec = 0
        try:
            from .tabelas import analisar_tabelas
            self.infos = {t.i: t for t in analisar_tabelas(d)}
        except Exception:
            self.infos = {}

    def _img(self, r):
        blips = r.xpath(".//a:blip/@r:embed")
        if not blips:
            return "<span class='img-ph'>[imagem]</span>"
        try:
            parte = self.doc.part.related_parts[blips[0]]
            blob = parte.blob
            ext = r.xpath(".//wp:extent")
            largura = (int(ext[0].get("cx")) / 360000) if ext else 8
            altura = (int(ext[0].get("cy")) / 360000) if ext else None
            # altura conhecida: a página não "pula" quando a imagem termina de carregar
            alt_css = f";height:{altura:.2f}cm" if altura else ""
            if len(blob) > LIMITE_IMAGEM:
                return f"<span class='img-ph' style='width:{largura:.2f}cm{alt_css}'>[imagem grande]</span>"
            uri = f"data:{parte.content_type};base64,{base64.b64encode(blob).decode()}"
            return (f"<img src='{uri}' alt='' style='width:{largura:.2f}cm{alt_css};max-width:100%;"
                    f"object-fit:contain'>")
        except Exception:
            return "<span class='img-ph'>[imagem]</span>"

    def _quebra_aqui(self, ch):
        tag = ch.tag
        if tag == qn("w:br") and ch.get(qn("w:type")) == "page":
            return QUEBRA
        if self.paginar_word and tag == qn("w:lastRenderedPageBreak"):
            return PAGINA_WORD
        return None

    def _run(self, r, par):
        """HTML do trecho em pedaços: [html, quebra, html, ...] (uma quebra a cada nova página dentro dele)."""
        grupos, quebras = [[]], []
        tem = False
        for ch in r:
            tag = ch.tag
            q = self._quebra_aqui(ch)
            if q is not None:
                quebras.append(q)
                grupos.append([])
            elif tag in (qn("w:t"), qn("w:delText")):
                grupos[-1].append(_esc(ch.text or ""))
                tem = True
            elif tag == qn("w:tab"):
                grupos[-1].append("<span class='tab'></span>")
                tem = True
            elif tag == qn("w:br"):
                grupos[-1].append("<br>")
                tem = True
            elif tag in (qn("w:drawing"), qn("w:pict")):
                grupos[-1].append(self._img(r))
                tem = True
        if not tem:
            return _intercalar([""] * len(grupos), quebras)
        run = Run(r, par)
        css = []
        try:
            tam = du.tamanho_run(run, par, self.doc)
            css.append(f"font-size:{tam:g}pt")
        except Exception:
            pass
        neg, ita = du.negrito_run(run, par), du.italico_run(run, par)
        sub_, sup_, und = run.font.subscript, run.font.superscript, run.underline

        def vestir(txt):
            if not txt:
                return ""
            if neg:
                txt = f"<b>{txt}</b>"
            if ita:
                txt = f"<i>{txt}</i>"
            if und:
                txt = f"<u>{txt}</u>"
            if sup_:
                txt = f"<sup>{txt}</sup>"
            elif sub_:
                txt = f"<sub>{txt}</sub>"
            return f"<span style='{';'.join(css)}'>{txt}</span>" if css else txt
        return _intercalar([vestir("".join(g)) for g in grupos], quebras)

    def _conteudo(self, el, par):
        """Conteúdo do parágrafo em pedaços: [html, quebra, html, ...]."""
        pedacos = [""]

        def juntar(novos):
            pedacos[-1] += novos[0]
            pedacos.extend(novos[1:])

        def vestir(novos, abre, fecha):
            return [x if isinstance(x, _Quebra) else (f"{abre}{x}{fecha}" if x else "") for x in novos]

        for ch in el:
            tag = ch.tag
            if tag == qn("w:r"):
                juntar(self._run(ch, par))
            elif tag in (qn("w:ins"), qn("w:moveTo")):
                autor = _esc(ch.get(qn("w:author")) or "")
                juntar(vestir(self._conteudo(ch, par), f"<ins title='{autor}'>", "</ins>"))
            elif tag in (qn("w:del"), qn("w:moveFrom")):
                autor = _esc(ch.get(qn("w:author")) or "")
                juntar(vestir(self._conteudo(ch, par), f"<del title='{autor}'>", "</del>"))
            elif tag in (qn("w:pPr"), qn("w:bookmarkStart"), qn("w:bookmarkEnd"), qn("w:proofErr")):
                continue
            else:  # hyperlink, fldSimple, smartTag, sdt...
                juntar(self._conteudo(ch, par))
        return pedacos

    # ------------------------------------------------------------------ parágrafo e tabela
    def _par_html(self, p, data_b, extra_classe=""):
        css = []
        al = ALINHA.get(du.alinhamento(p))
        if al:
            css.append(f"text-align:{al}")
        esq = du.recuo_esquerdo_cm(p)
        if esq:
            css.append(f"margin-left:{esq:.2f}cm")
        pri = du._pf(p, "first_line_indent")
        if pri:
            css.append(f"text-indent:{du.cm(pri):.2f}cm")
        try:
            css.append(f"line-height:{max(1.0, du.entrelinhas(p, self.doc)) * 1.18:.2f}")
        except Exception:
            pass
        for prop, nome in (("space_before", "margin-top"), ("space_after", "margin-bottom")):
            v = du._pf(p, prop)
            if v is not None:
                css.append(f"{nome}:{v.pt:.1f}pt")
        nivel = du.nivel_titulo_estilo(p)
        classe = "par" + (f" tit tit{nivel}" if nivel else "") + (f" {extra_classe}" if extra_classe else "")
        itens = self._conteudo(p._p, p)
        if extra_classe == "cel":  # dentro de tabela a página não é dividida
            itens = ["".join(x for x in itens if not isinstance(x, _Quebra))]
        out, mostrado = [], False
        for x in itens:
            if isinstance(x, _Quebra):
                out.append(x)
            elif x.strip() or (not mostrado and len(itens) == 1):
                estilo = css if not mostrado else [c for c in css if not c.startswith(("text-indent", "margin-top"))]
                out.append(f"<p class='{classe}{' cont' if mostrado else ''}' data-b='{data_b}' "
                           f"style='{';'.join(estilo)}'>{x or '&nbsp;'}</p>")
                mostrado = True
        return out

    def paragrafo(self, b):
        return self._par_html(b.obj, b.i)

    SELO = {"tabela": ("Tabela de dados", "selo-tab"), "quadro": ("Quadro (texto)", "selo-qua"),
            "diagramação": ("Diagramação — não é tabela", "selo-lay"), "ilustração": ("Ilustração", "selo-qua")}

    def tabela(self, b):
        from docx.text.paragraph import Paragraph
        t = b.obj
        tbl = t._tbl
        info = self.infos.get(b.i)
        natureza = info.natureza if info else "tabela"
        if info and info.natureza in ("tabela", "quadro") and not info.leg_tipo:
            selo, cls = f"Sem título · {'tabela de dados' if natureza == 'tabela' else 'quadro'}", "selo-sem"
        else:
            rot, cls = self.SELO.get(natureza, ("Tabela", "selo-tab"))
            selo = f"{info.nome} · {rot}" if info and info.leg_tipo else rot
        if info and info.leg_tipo == "Tabela" and info.frac_num < 0.3 and info.colunas >= 2:
            selo, cls = f"{info.nome} · conteúdo de texto: seria Quadro", "selo-sem"
        elif info and info.leg_tipo == "Quadro" and info.frac_num >= 0.7:
            selo, cls = f"{info.nome} · dados numéricos: seria Tabela", "selo-sem"
        if info and natureza in ("tabela", "quadro"):
            selo += f" · {info.linhas}×{info.colunas}"
        vert = info.bordas_vert if info else True
        estilo = "vert" if (vert or (info and info.bordas_fechadas)) else "horiz"
        if natureza == "diagramação":
            estilo = "lay"
        cab = bool(info and (info.cab_repete or info.cab_destacado))

        # grade: posição de cada célula, para colspan/rowspan
        grade = []
        for tr in tbl.tr_lst:
            col, linha = 0, []
            for c, tc in enumerate(tr.tc_lst):
                tcpr = tc.tcPr
                span = 1
                vm = None
                if tcpr is not None:
                    gs = tcpr.find(qn("w:gridSpan"))
                    if gs is not None:
                        span = int(gs.get(qn("w:val")) or 1)
                    v = tcpr.find(qn("w:vMerge"))
                    if v is not None:
                        vm = "restart" if v.get(qn("w:val")) == "restart" else "cont"
                linha.append((c, tc, col, span, vm))
                col += span
            grade.append(linha)
        linhas = []
        quebras = set()
        if self.paginar_word:
            for r, tr in enumerate(tbl.tr_lst):
                if r and tr.find(".//" + qn("w:lastRenderedPageBreak")) is not None:
                    quebras.add(r)
        for r, linha in enumerate(grade):
            if r in quebras:
                linhas.append(PAGINA_WORD)
            cels = []
            for c, tc, col, span, vm in linha:
                if vm == "cont":
                    continue
                rowspan = 1
                if vm == "restart":
                    for prox in grade[r + 1:]:
                        if any(x[2] == col and x[4] == "cont" for x in prox):
                            rowspan += 1
                        else:
                            break
                pars = [el for el in tc if el.tag == qn("w:p")]
                corpo = "".join(x for k, pel in enumerate(pars)
                                for x in self._par_html(Paragraph(pel, t), f"{b.i}:{r}:{c}:{k}", "cel"))
                attrs = (f" colspan='{span}'" if span > 1 else "") + (f" rowspan='{rowspan}'" if rowspan > 1 else "")
                cels.append(f"<td{attrs}>{corpo}</td>")
            classe_tr = " class='cab'" if (r == 0 and cab) else ""
            linhas.append(f"<tr{classe_tr}>{''.join(cels)}</tr>")
        out, atual, primeira = [], [], True
        for x in linhas + [PAGINA_WORD]:
            if not isinstance(x, _Quebra):
                atual.append(x)
                continue
            if atual:
                rotulo = _esc(selo) if primeira else f"{_esc(selo)} · continuação"
                out.append(f"<div class='tab-wrap'><div class='tab-selo {cls}'>{rotulo}</div>"
                           f"<table class='tbl {estilo}' data-b='{b.i}'>{''.join(atual)}</table></div>")
                primeira, atual = False, []
            out.append(QUEBRA)
        return out[:-1]


    # ------------------------------------------------------------------ folhas
    def _abrir_folha(self):
        secs = self.doc.sections
        s = secs[min(self.sec, len(secs) - 1)]
        med = du.cm(s.gutter) or 0
        top, bot = abs(du.cm(s.top_margin) or 3), abs(du.cm(s.bottom_margin) or 2)
        esq, dir_ = (du.cm(s.left_margin) or 3) + med, du.cm(s.right_margin) or 2
        larg = du.cm(s.page_width) or 21
        alt = du.cm(s.page_height) or 29.7
        pag = " data-paginacao='word'" if self.paginar_word else ""
        self.partes.append(
            f"<section class='folha' data-sec='{self.sec + 1}'{pag} style='width:{larg:.2f}cm;min-height:{alt:.2f}cm;"
            f"padding:{top:.2f}cm {dir_:.2f}cm {bot:.2f}cm {esq:.2f}cm'>"
            f"<div class='rotulo-sec'>Seção {self.sec + 1} do Word · margens {esq:g}/{dir_:g}/{top:g}/{bot:g} cm"
            f"</div><div class='area'>")
        self.pagina_vazia = True
        self.primeira_folha = len(self.partes) == 1

    def _fechar_folha(self):
        self.partes.append("</div></section>")

    def _nova_pagina(self, quebra=QUEBRA):
        """Abre a página seguinte. A marca do Word logo depois de uma quebra do aluno é a mesma página
        (não conta duas vezes); duas quebras do aluno seguidas fazem uma página em branco, como no Word."""
        if self.pagina_vazia and (not quebra.manual or self.primeira_folha):
            return
        self._fechar_folha()
        self._abrir_folha()

    def _por(self, pedacos):
        for x in pedacos:
            if isinstance(x, _Quebra):
                self._nova_pagina(x)
            else:
                self.partes.append(x)
                self.pagina_vazia = False

    def gerar(self):
        self._abrir_folha()
        primeiro = True
        for b in self.d.blocos:
            if b.tipo == "p":
                p = b.obj
                if not primeiro and du._pf(p, "page_break_before"):
                    self._nova_pagina(PAGINA_WORD)  # "quebra antes" no topo de uma página não faz página nova
                self._por(self.paragrafo(b))
                primeiro = False
                ppr = p._p.pPr
                if ppr is not None and ppr.sectPr is not None:
                    self.sec += 1  # quebra de seção: folha nova com as medidas da seção seguinte
                    self._fechar_folha()
                    self._abrir_folha()
            else:
                self._por(self.tabela(b))
                primeiro = False
        self._fechar_folha()
        return "".join(self.partes)


def renderizar(arquivo):
    with cache.leitura():
        d = Documento(arquivo)
        if d.pdf is not None:  # PDF: as próprias páginas, com uma camada de texto para as marcações
            from .pdf import html_paginas
            return html_paginas(d.pdf), d
        return _Render(d).gerar(), d


# ====================================================================== localização dos achados
# Ordem de confiança (como nos verificadores de manuscritos: cada apontamento ligado ao texto que o gerou):
#   1. trechos que já sabem de qual bloco vieram (modelo.Trecho) e blocos informados pela verificação;
#   2. busca do texto do trecho, escolhendo a ocorrência na zona certa do trabalho (não no sumário ou
#      nas listas, se o problema é do texto; nas referências, se é de referência...);
#   3. o elemento a que o item se refere (título "RESUMO", linha de palavras-chave...), achado de forma
#      tolerante a variações de grafia.
MSG_ALINEA = re.compile(r"^(começa com maiúscula|deveria terminar em '.'|"
                        r"texto antes das alíneas não termina em dois-pontos):\s*")
RE_ENTRADA_LISTA = re.compile(r"(\.{4,}|…{2,}|\t)\s*[\divxlcIVXLC]+\s*$")  # "Figura 3 ... 24" (sumário/listas)
CHAVE_PALAVRAS = r"PALAVRAS?[\s-]*CHAVES?\s*[:.\-–]?.*"
CHAVE_KEYWORDS = r"KEY[\s-]*WORDS?\s*[:.\-–]?.*"
POR_ITEM = [  # itens sem trecho: onde apontar (o mais específico primeiro)
    (r"Palavras-chave", CHAVE_PALAVRAS), (r"Keywords", CHAVE_KEYWORDS),
    (r"^Resumo", r"RESUMO\W*"), (r"^Abstract", r"ABSTRACT\W*"), (r"^Sumário", r"SUMARIO\W*"),
    (r"^Introdução", r"(\d+\s*[-.–]?\s*)?INTRODUCAO\W*"), (r"^Referências$", r"REFERENCIAS( BIBLIOGRAFICAS)?\W*"),
    (r"^Lista de referências|^Título da lista de referências", r"REFERENCIAS( BIBLIOGRAFICAS)?\W*|BIBLIOGRAFIA\W*"),
    (r"orientador", r"(PROF\w*\.?\s*(DR\w*\.?\s*)?)?(CO)?ORIENTADOR.*"),
    (r"versão original", r"VERSAO (ORIGINAL|CORRIGIDA)"),
    (r"^Ordem alfabética|^Destaque uniforme|^Prenomes|^4 ou mais", r"REFERENCIAS( BIBLIOGRAFICAS)?\W*"),
    (r"^Declaração de uso de IA", r".*INTELIGENCIA ARTIFICIAL.*"),
    (r"^Lista de tabelas", r"LISTA DE TABELAS\W*"), (r"^Lista de ilustrações|^Lista de figuras", r"LISTA DE (ILUSTRACOES|FIGURAS)\W*"),
    (r"^Lista de siglas|^Lista de abreviaturas", r"LISTA DE (SIGLAS|ABREVIATURAS).*"),
    (r"^Agradecimentos", r"AGRADECIMENTOS?\W*"), (r"^Epígrafe", r"EPIGRAFE\W*"),
]
# onde cada tipo de problema costuma estar (para escolher entre várias ocorrências do mesmo texto)
ZONA_DA_CATEGORIA = {"Citações": ("textual",), "Ilustrações": ("textual", "pos"), "Tabelas e quadros": ("textual", "pos"),
                     "Seções e títulos": ("textual", "pos"), "Referências": ("refs",),
                     "Texto e tipografia": ("textual",), "Formatação": ("textual",)}


def _normal(s):
    return re.sub(r"\s+", " ", s).strip()


def _sem_acento_baixo(s):
    return chave(s).lower()


class _Localizador:
    def __init__(self, d):
        self.d = d
        self.textos = []
        self.lista = set()  # blocos que são linhas de sumário/listas
        for b in d.blocos:
            if b.tipo == "p":
                self.textos.append((b.i, b.texto))
                if RE_ENTRADA_LISTA.search(b.bruto or b.texto):
                    self.lista.add(b.i)
            else:
                try:
                    t = " ".join(c.text for row in b.obj.rows for c in row.cells)
                except Exception:
                    t = ""
                self.textos.append((b.i, t))
        self.por_bloco = dict(self.textos)
        self._norm = {i: _normal(t) for i, t in self.textos}
        self._fold = None

    def zona(self, i):
        if self.d.i_refs is not None and i > self.d.i_refs and i < getattr(self.d, "i_fim_refs", len(self.d.blocos)):
            return "refs"
        return self.d.regiao(i)

    @staticmethod
    def nucleo(trecho):
        core = MSG_ALINEA.sub("", str(trecho).strip())
        return re.sub(r"^\[[^\]]*\]\s*", "", core).rstrip("…").strip()

    def marca_no_bloco(self, i, trecho):
        """O pedaço do trecho a marcar dentro do bloco, se for um pedaço (e não o parágrafo todo)."""
        if trecho is None:
            return None
        core = _normal(self.nucleo(trecho))
        t = self._norm.get(i, "")
        if 2 < len(core) < 60 and len(t) > len(core) + 10 and core in t:
            return core
        return None

    def _candidatos(self, alvo, categoria):
        alvo_n = _normal(alvo)
        if len(alvo_n) < 3:
            return []
        busca = alvo_n[:70]
        achados = [i for i, t in self._norm.items() if busca in t]
        if not achados:  # tolera acento, maiúsculas e pontuação colada
            if self._fold is None:
                self._fold = {i: _sem_acento_baixo(t) for i, t in self._norm.items()}
            dobra = _sem_acento_baixo(busca)
            if len(dobra) >= 4:
                achados = [i for i, t in self._fold.items() if dobra in t]
        if len(achados) <= 1:
            return achados
        zonas = ZONA_DA_CATEGORIA.get(categoria, ())
        sumario = categoria == "Sumário e listas"

        def nota(i):
            pontos = 0
            if (i in self.lista) != sumario:
                pontos += 4  # linha do sumário/lista só vale para problemas do sumário/listas
            if zonas and self.zona(i) not in zonas:
                pontos += 2
            if self._norm[i] == busca:
                pontos -= 1  # o parágrafo é exatamente o trecho
            return pontos
        return sorted(achados, key=lambda i: (nota(i), i))

    def achar(self, trecho, categoria="", usados=()):
        core = self.nucleo(trecho)
        m = re.match(r"^\S+ → (\S+)$", core)  # sequência de numeração: "2.1 → 2.3"
        if m:
            alvo = m.group(1)
            for i, t in self.textos:
                if i not in self.lista and re.match(rf"^{re.escape(alvo)}\.?\s", t):
                    return {"b": i, "m": None}
            return None
        m = re.match(r"^(.+?) \(([^()]+)\)( ≈ .*)?$", core)  # citação: "Prado (2021)"
        if m and len(core) < 120:
            nome, ano = m.group(1), m.group(2)
            for i, t in self.textos:
                if self.zona(i) == "refs" or i in self.lista:
                    continue
                mm = re.search(rf"{re.escape(nome)}[^()]{{0,60}}?{re.escape(ano)}\)?", t)
                if mm:
                    return {"b": i, "m": mm.group(0)}
        cands = self._candidatos(core, categoria)
        if not cands:
            return None
        livres = [i for i in cands if i not in usados] or cands
        i = livres[0]
        return {"b": i, "m": self.marca_no_bloco(i, core)}

    def por_item(self, item):
        for padrao_item, padrao_txt in POR_ITEM:
            if re.search(padrao_item, item, re.I):
                for i, t in self.textos:
                    if t and len(t) < 300 and i not in self.lista and re.fullmatch(padrao_txt, chave(t)):
                        return {"b": i, "m": None}
        return None


# trecho exato a marcar em negrito, por tipo de problema
MARCAS = [
    (r"^Edição", r"\d+\s*[ªºa°]\.?\s*(?:edi[çc][ãa]o|ed)\b\.?|\d+\.\s*edi[çc][ãa]o"),
    (r"^Páginas", r"\bpp\.|\bp\.\d|\bpág\w*"),
    (r"^Meses", r"(?i)\b(janeiro|fevereiro|março|abril|junho|julho|agosto|setembro|outubro|novembro|dezembro|maio\.)"),
    (r"^Entrada pelo sobrenome", r"^[^,.]+"),
    (r"^'In:'", r"\b(IN|in):"),
    (r"^'Disponível em:'", r"Dispon[ií]vel em"),
    (r"^'Acesso em", r"Acesso.{0,22}"),
    (r"^DOI antes", r"DOI"),
    (r"^Referência termina", r".{1,12}$"),
    (r"^Indicativo numérico", r"^\d+(\.\d+)*\.?\s*[-–—:.]+"),
    (r"^Títulos sem indicativo numérico", r"^\d+(\.\d+)*\.?\s*"),
    (r"^Ilustrações: título no formato|^Tabelas: título", r"^\w+\s+\d+(\.\d+)?\s*[:.\-–—]?"),
    (r"formato de 'Palavras|formato de 'Keywords", r"(?<=:).+"),
    (r"^Título da lista de referências", r".+"),
]


def _marca_do_item(item, texto):
    for padrao_item, padrao_txt in MARCAS:
        if re.search(padrao_item, item):
            m = re.search(padrao_txt, texto)
            if m and m.group(0).strip():
                return m.group(0).strip()
    return None


def serializar(d, achados):
    """Achados -> dicts para a área de trabalho, com os locais no documento."""
    with cache.leitura():
        return _serializar(d, achados)


def _serializar(d, achados):
    loc = _Localizador(d)
    saida = []
    for n, a in enumerate(achados):
        locais, vistos = [], set()

        def por(i, m, locais=locais, vistos=vistos):
            if i is None or (i, m) in vistos or i not in loc.por_bloco:
                return
            vistos.add((i, m))
            locais.append({"b": i, "m": m})

        # 1. trechos ligados ao bloco de origem e blocos informados pela própria verificação
        for tr in a.trechos:
            if getattr(tr, "i", None) is not None:
                por(tr.i, getattr(tr, "m", None) or loc.marca_no_bloco(tr.i, tr))
        pares = list(zip(a.blocos, a.trechos)) if len(a.blocos) == len(a.trechos) else [(i, None) for i in a.blocos]
        for i, tr in pares:
            if not any(x["b"] == i for x in locais):
                por(i, loc.marca_no_bloco(i, tr))
        # 2. trechos sem origem: busca no texto, na zona certa
        if not locais:
            usados = set()
            for tr in a.trechos:
                x = loc.achar(tr, a.categoria, usados)
                if x:
                    usados.add(x["b"])
                    por(x["b"], x["m"])
        # 3. o elemento a que o item se refere
        if not locais and a.status != OK and not a.tabela:
            x = loc.por_item(a.item)
            if x:
                por(x["b"], x["m"])
        for x in locais:
            if not x["m"]:
                x["m"] = _marca_do_item(a.item, loc.por_bloco.get(x["b"], ""))
        saida.append({
            "id": n, "categoria": a.categoria, "item": a.item, "status": a.status, "detalhe": a.detalhe,
            "corrigir": a.como_corrigir, "secao": a.secao, "origem": a.origem, "tabela": a.tabela,
            "trechos": [str(t) for t in a.trechos[:50]], "locais": locais,
            "documento_todo": not locais and (a.categoria == "Formatação" and a.item.startswith(("Margens", "Papel",
                                                                                                    "Paginação"))),
        })
    return saida
