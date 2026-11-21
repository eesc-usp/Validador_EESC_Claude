"""Renderiza o .docx como folhas A4 em HTML e liga cada achado ao trecho do documento.

Cada parágrafo/tabela recebe data-b="<índice do bloco>" (o mesmo índice usado pelo Documento),
para a área de trabalho rolar até ele, marcar o trecho e permitir edição.
"""
import base64
import html
import re

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.text.run import Run

from . import docx_util as du
from .correcao import _texto_run, texto_atual
from .documento import Documento, chave
from .modelo import OK

ALINHA = {WD_ALIGN_PARAGRAPH.CENTER: "center", WD_ALIGN_PARAGRAPH.RIGHT: "right",
          WD_ALIGN_PARAGRAPH.JUSTIFY: "justify", WD_ALIGN_PARAGRAPH.DISTRIBUTE: "justify"}
LIMITE_IMAGEM = 1_500_000  # bytes; imagens maiores viram um marcador


def _esc(s):
    return html.escape(s, quote=True)


# ====================================================================== runs
class _Render:
    def __init__(self, d):
        self.d = d
        self.doc = d.doc
        self.partes = []
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
            if len(blob) > LIMITE_IMAGEM:
                return f"<span class='img-ph' style='width:{largura:.2f}cm'>[imagem grande]</span>"
            uri = f"data:{parte.content_type};base64,{base64.b64encode(blob).decode()}"
            return f"<img src='{uri}' style='width:{largura:.2f}cm;max-width:100%'>"
        except Exception:
            return "<span class='img-ph'>[imagem]</span>"

    def _run(self, r, par):
        out = []
        for ch in r:
            tag = ch.tag
            if tag in (qn("w:t"), qn("w:delText")):
                out.append(_esc(ch.text or ""))
            elif tag == qn("w:tab"):
                out.append("<span class='tab'></span>")
            elif tag == qn("w:br"):
                out.append("<span class='quebra-pag'>quebra de página</span>" if ch.get(qn("w:type")) == "page"
                           else "<br>")
            elif tag in (qn("w:drawing"), qn("w:pict")):
                out.append(self._img(r))
        if not out:
            return ""
        run = Run(r, par)
        css = []
        try:
            tam = du.tamanho_run(run, par, self.doc)
            css.append(f"font-size:{tam:g}pt")
        except Exception:
            pass
        txt = "".join(out)
        if du.negrito_run(run, par):
            txt = f"<b>{txt}</b>"
        if du.italico_run(run, par):
            txt = f"<i>{txt}</i>"
        if run.underline:
            txt = f"<u>{txt}</u>"
        if run.font.superscript:
            txt = f"<sup>{txt}</sup>"
        elif run.font.subscript:
            txt = f"<sub>{txt}</sub>"
        return f"<span style='{';'.join(css)}'>{txt}</span>" if css else txt

    def _conteudo(self, el, par):
        out = []
        for ch in el:
            tag = ch.tag
            if tag == qn("w:r"):
                out.append(self._run(ch, par))
            elif tag in (qn("w:ins"), qn("w:moveTo")):
                out.append(f"<ins title='{_esc(ch.get(qn('w:author')) or '')}'>{self._conteudo(ch, par)}</ins>")
            elif tag in (qn("w:del"), qn("w:moveFrom")):
                out.append(f"<del title='{_esc(ch.get(qn('w:author')) or '')}'>{self._conteudo(ch, par)}</del>")
            elif tag in (qn("w:pPr"), qn("w:bookmarkStart"), qn("w:bookmarkEnd"), qn("w:proofErr")):
                continue
            else:  # hyperlink, fldSimple, smartTag, sdt...
                out.append(self._conteudo(ch, par))
        return "".join(out)

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
        corpo = self._conteudo(p._p, p) or "&nbsp;"
        return (f"<p class='{classe}' data-b='{data_b}' data-t='{_esc(texto_atual(p))}' "
                f"style='{';'.join(css)}'>{corpo}</p>")

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
        for r, linha in enumerate(grade):
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
                corpo = "".join(self._par_html(Paragraph(pel, t), f"{b.i}:{r}:{c}:{k}", "cel")
                                for k, pel in enumerate(pars))
                attrs = (f" colspan='{span}'" if span > 1 else "") + (f" rowspan='{rowspan}'" if rowspan > 1 else "")
                cels.append(f"<td{attrs}>{corpo}</td>")
            classe_tr = " class='cab'" if (r == 0 and cab) else ""
            linhas.append(f"<tr{classe_tr}>{''.join(cels)}</tr>")
        return (f"<div class='tab-wrap'><div class='tab-selo {cls}'>{_esc(selo)}</div>"
                f"<table class='tbl {estilo}' data-b='{b.i}'>{''.join(linhas)}</table></div>")

    # ------------------------------------------------------------------ folhas
    def _abrir_folha(self, sec_idx):
        secs = self.doc.sections
        s = secs[min(sec_idx, len(secs) - 1)]
        med = du.cm(s.gutter) or 0
        top, bot = abs(du.cm(s.top_margin) or 3), abs(du.cm(s.bottom_margin) or 2)
        esq, dir_ = (du.cm(s.left_margin) or 3) + med, du.cm(s.right_margin) or 2
        larg = du.cm(s.page_width) or 21
        alt = du.cm(s.page_height) or 29.7
        self.partes.append(
            f"<section class='folha' data-sec='{sec_idx + 1}' style='width:{larg:.2f}cm;min-height:{alt:.2f}cm;"
            f"padding:{top:.2f}cm {dir_:.2f}cm {bot:.2f}cm {esq:.2f}cm'>"
            f"<div class='rotulo-sec'>Seção {sec_idx + 1} do Word · margens {esq:g}/{dir_:g}/{top:g}/{bot:g} cm"
            f"</div><div class='area'>")

    def _fechar_folha(self):
        self.partes.append("</div></section>")

    def gerar(self):
        sec = 0
        self._abrir_folha(sec)
        anterior = None
        for b in self.d.blocos:
            if b.tipo == "p":
                p = b.obj
                if anterior is not None and du.comeca_em_nova_pagina(p, anterior) and \
                        not (anterior._p.pPr is not None and anterior._p.pPr.sectPr is not None):
                    self._fechar_folha()
                    self._abrir_folha(sec)
                self.partes.append(self.paragrafo(b))
                anterior = p
                ppr = p._p.pPr
                if ppr is not None and ppr.sectPr is not None:
                    sec += 1
                    self._fechar_folha()
                    self._abrir_folha(sec)
            else:
                self.partes.append(self.tabela(b))
        self._fechar_folha()
        return "".join(self.partes)


def renderizar(arquivo):
    d = Documento(arquivo)
    return _Render(d).gerar(), d


# ====================================================================== localização dos achados
MSG_ALINEA = re.compile(r"^(começa com maiúscula|deveria terminar em '.'|"
                        r"texto antes das alíneas não termina em dois-pontos):\s*")
POR_ITEM = [  # itens sem trecho: onde apontar
    (r"Palavras-chave", r"PALAVRAS-CHAVE:.*"), (r"Keywords", r"KEYWORDS:.*"),
    (r"^Resumo", r"RESUMO"), (r"^Abstract", r"ABSTRACT"), (r"^Sumário", r"SUMARIO"),
    (r"^Introdução", r"(1\s*)?INTRODUCAO"), (r"^Referências$", r"REFERENCIAS"),
    (r"^Lista de referências", r"REFERENCIAS"), (r"Palavras-chave", r"PALAVRAS-CHAVE:.*"),
    (r"orientador", r"ORIENTADOR.*"), (r"versão original", r"VERSAO (ORIGINAL|CORRIGIDA)"),
    (r"^Ordem alfabética|^Destaque uniforme|^Prenomes|^4 ou mais", r"REFERENCIAS"),
]


class _Localizador:
    def __init__(self, d):
        self.d = d
        self.textos = []
        for b in d.blocos:
            if b.tipo == "p":
                self.textos.append((b.i, b.texto))
            else:
                try:
                    t = " ".join(c.text for row in b.obj.rows for c in row.cells)
                except Exception:
                    t = ""
                self.textos.append((b.i, t))

    def achar(self, trecho):
        core = MSG_ALINEA.sub("", trecho.strip())
        core = re.sub(r"^\[[^\]]*\]\s*", "", core).rstrip("…").strip()
        m = re.match(r"^\S+ → (\S+)$", core)  # sequência de numeração: "2.1 → 2.3"
        if m:
            alvo = m.group(1)
            for i, t in self.textos:
                if re.match(rf"^{re.escape(alvo)}\.?\s", t):
                    return {"b": i, "m": None}
            return None
        m = re.match(r"^(.+?) \(([^()]+)\)$", core)  # citação sem referência: "Prado (2021)"
        if m and len(core) < 80:
            nome, ano = m.groups()
            for i, t in self.textos:
                mm = re.search(rf"{re.escape(nome)}[^()]{{0,60}}?{re.escape(ano)}", t)
                if mm:
                    return {"b": i, "m": mm.group(0)}
        if len(core) < 3:
            return None
        chave_busca = core[:70]
        for i, t in self.textos:
            pos = t.find(chave_busca)
            if pos >= 0:
                marca = core if len(core) < 45 and len(t) > len(core) + 10 else None
                return {"b": i, "m": marca}
        return None

    def por_item(self, item):
        for padrao_item, padrao_txt in POR_ITEM:
            if re.search(padrao_item, item, re.I):
                for i, t in self.textos:
                    if t and re.fullmatch(padrao_txt, chave(t)):
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
    loc = _Localizador(d)
    saida = []
    for n, a in enumerate(achados):
        locais = [{"b": i, "m": None} for i in dict.fromkeys(a.blocos)]
        for tr in ([] if a.blocos else a.trechos):
            x = loc.achar(tr)
            if x and x not in locais:
                locais.append(x)
        if not locais and a.status != OK and not a.tabela:
            x = loc.por_item(a.item)
            if x:
                locais.append(x)
        textos = dict(loc.textos)
        for x in locais:
            if not x["m"]:
                x["m"] = _marca_do_item(a.item, textos.get(x["b"], ""))
        saida.append({
            "id": n, "categoria": a.categoria, "item": a.item, "status": a.status, "detalhe": a.detalhe,
            "corrigir": a.como_corrigir, "secao": a.secao, "origem": a.origem, "tabela": a.tabela,
            "trechos": a.trechos[:50], "locais": locais,
            "documento_todo": not locais and (a.categoria == "Formatação" and a.item.startswith(("Margens", "Papel",
                                                                                                    "Paginação"))),
        })
    return saida
