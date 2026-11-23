"""Leitura do .docx com propriedades EFETIVAS (run -> estilo -> padrão do documento).

O Word guarda formatação em camadas: um parágrafo pode não ter tamanho de fonte
próprio e herdar do estilo, que herda do estilo base, que herda do padrão do
documento. Estas funções resolvem essa cadeia para a verificação não errar.
"""
import re
from collections import Counter

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.shared import Length
from docx.table import Table
from docx.text.paragraph import Paragraph

from . import cache
from .cache import por_elemento, qn

cache.instalar()

EMU_POR_CM = 360000


def cm(valor):
    return None if valor is None else round(valor / EMU_POR_CM, 2)


# ---------------------------------------------------------------- blocos
def blocos(doc):
    """Percorre o corpo na ordem real: ('p', Paragraph) ou ('t', Table).

    Entra nos controles de conteúdo (w:sdt) do corpo: o sumário automático do Word, por exemplo,
    fica dentro de um — sem isso, as linhas dele não seriam vistas."""
    yield from _blocos_de(doc.element.body, doc)


_W_P, _W_TBL, _W_SDT, _W_SDT_CONTENT, _W_CUSTOM_XML = (
    qn("w:p"), qn("w:tbl"), qn("w:sdt"), qn("w:sdtContent"), qn("w:customXml"))


def _blocos_de(raiz, doc):
    for el in raiz.iterchildren():
        tag = el.tag
        if tag == _W_P:
            yield "p", Paragraph(el, doc)
        elif tag == _W_TBL:
            yield "t", Table(el, doc)
        elif tag == _W_SDT:
            conteudo = el.find(_W_SDT_CONTENT)
            if conteudo is not None:
                yield from _blocos_de(conteudo, doc)
        elif tag == _W_CUSTOM_XML:
            yield from _blocos_de(el, doc)


@por_elemento
def texto(par):
    """Texto visível: inclui o que está em alterações controladas inseridas e exclui o apagado."""
    from .texto import texto_atual
    return texto_atual(par).strip()


# ---------------------------------------------------------------- padrões do documento
def _doc_defaults(doc):
    return cache.memo(("doc_defaults", doc), lambda: _doc_defaults_ler(doc))


def _doc_defaults_ler(doc):
    estilos = doc.styles.element
    dd = estilos.find(qn("w:docDefaults"))
    sz, linha = None, None
    if dd is not None:
        el = dd.find(".//" + qn("w:rPrDefault") + "//" + qn("w:sz"))
        if el is not None:
            sz = int(el.get(qn("w:val"))) / 2
        sp = dd.find(".//" + qn("w:pPrDefault") + "//" + qn("w:spacing"))
        if sp is not None and sp.get(qn("w:line")):
            regra = sp.get(qn("w:lineRule"), "auto")
            v = int(sp.get(qn("w:line")))
            linha = ("mult", v / 240) if regra == "auto" else ("pt", v / 20)
    return sz, linha


def _cadeia(estilo):
    while estilo is not None:
        yield estilo
        estilo = estilo.base_style


def cadeia_par(par):
    """Estilo do parágrafo e seus estilos-base (guardado por documento)."""
    return cache.cadeia(par, lambda p: p.style)


def cadeia_run(run):
    """Estilo de caractere do trecho e seus estilos-base (guardado por documento)."""
    return cache.cadeia(run, lambda r: r.style)


# ---------------------------------------------------------------- fonte
def fonte_nos_estilos(run, par, prop, estilo_do_run=True, so_verdadeiro=False):
    """Primeiro valor da propriedade de fonte 'prop' definido nos estilos: primeiro na cadeia do
    estilo de caractere do trecho (se 'estilo_do_run'), depois na do parágrafo. None se nenhum define.

    Os estilos não mudam durante uma análise, então a resposta é guardada por par de estilos."""
    def calcular():
        for cad in ((cadeia_run(run), cadeia_par(par)) if estilo_do_run else (cadeia_par(par),)):
            for est in cad:
                v = getattr(est.font, prop)
                if (v if so_verdadeiro else v is not None):
                    return v
        return None
    if not cache.ativo():
        return calcular()
    try:
        k = ("fonte_estilos", prop, estilo_do_run, so_verdadeiro, par.part,
             cache.id_do_estilo(run._r) if estilo_do_run else None, cache.id_do_estilo(par._p))
    except Exception:
        return calcular()
    return cache.memo(k, calcular)


def tamanho_run(run, par, doc):
    if run.font.size is not None:
        return run.font.size.pt
    v = fonte_nos_estilos(run, par, "size")
    if v is not None:
        return v.pt
    return _doc_defaults(doc)[0] or 10.0  # Word usa 10 pt se nada for definido


@por_elemento
def tamanho_predominante(par, doc):
    """Tamanho de fonte que cobre mais caracteres do parágrafo."""
    cont = Counter()
    for r in par.runs:
        t = r.text
        if t.strip():
            cont[tamanho_run(r, par, doc)] += len(t)
    return cont.most_common(1)[0][0] if cont else None


def negrito_run(run, par):
    if run.bold is not None:
        return run.bold
    v = fonte_nos_estilos(run, par, "bold")
    return v if v is not None else False


def italico_run(run, par):
    if run.italic is not None:
        return run.italic
    v = fonte_nos_estilos(run, par, "italic", estilo_do_run=False)
    return v if v is not None else False


@por_elemento
def texto_marcado(par):
    """Texto com **negrito** marcado, para as referências não perderem o destaque."""
    partes, buf, em_negrito = [], "", False
    for r in par.runs:
        b = negrito_run(r, par) and r.text.strip() != ""
        if b != em_negrito and buf:
            partes.append(f"**{buf}**" if em_negrito else buf)
            buf = ""
        em_negrito = b if r.text.strip() else em_negrito
        buf += r.text
    if buf:
        partes.append(f"**{buf}**" if em_negrito else buf)
    t = "".join(partes)
    return re.sub(r"\*\*(\s*)\*\*", r"\1", t).strip()


@por_elemento
def tem_negrito(par):
    return any(negrito_run(r, par) for r in par.runs if r.text.strip())


@por_elemento
def tem_italico(par):
    return any(italico_run(r, par) for r in par.runs if r.text.strip())


@por_elemento
def todo_negrito(par):
    runs = [r for r in par.runs if r.text.strip()]
    return bool(runs) and all(negrito_run(r, par) for r in runs)


# ---------------------------------------------------------------- parágrafo
@por_elemento
def _pf(par, attr):
    v = getattr(par.paragraph_format, attr)
    if v is not None:
        return v
    for est in cadeia_par(par):
        v = getattr(est.paragraph_format, attr)
        if v is not None:
            return v
    return None


@por_elemento
def entrelinhas(par, doc):
    """Retorna o espaçamento como múltiplo de linha (1.0, 1.5...) ou None.

    Espaçamento 'exato' em pontos é convertido supondo fonte 12 (1 linha ~ 12 pt).
    """
    v = _pf(par, "line_spacing")
    regra = _pf(par, "line_spacing_rule")
    if v is None:
        d = _doc_defaults(doc)[1]
        if d is None:
            return 1.0
        return d[1] if d[0] == "mult" else round(d[1] / 12, 2)
    if isinstance(v, Length):
        return round(v.pt / 12, 2)
    if regra in (WD_LINE_SPACING.SINGLE,):
        return 1.0
    return round(float(v), 2)


@por_elemento
def alinhamento(par):
    a = _pf(par, "alignment")
    return a if a is not None else WD_ALIGN_PARAGRAPH.LEFT


@por_elemento
def recuo_esquerdo_cm(par):
    v = _pf(par, "left_indent")
    return cm(v) or 0.0


def comeca_em_nova_pagina(par, anterior):
    if _pf(par, "page_break_before"):
        return True
    xml_par = par._p
    primeiro = xml_par.find(".//" + qn("w:br"))
    if primeiro is not None and primeiro.get(qn("w:type")) == "page" and not texto_antes_do_br(par):
        return True
    if anterior is not None:
        pa = anterior._p
        for br in pa.iter(qn("w:br")):
            if br.get(qn("w:type")) == "page":
                return True
        ppr = pa.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
            return True  # quebra de seção
    return False


def texto_antes_do_br(par):
    txt = ""
    for el in par._p.iter():
        if el.tag == qn("w:br") and el.get(qn("w:type")) == "page":
            return txt.strip() != ""
        if el.tag == qn("w:t") and el.text:
            txt += el.text
    return False


@por_elemento
def tem_imagem(par):
    return bool(par._p.xpath(".//w:drawing")) or bool(par._p.xpath(".//w:pict"))


@por_elemento
def nivel_titulo_estilo(par):
    """Nível do título pelo estilo (Heading N / Título N / nível de tópico)."""
    for est in cadeia_par(par):
        nome = (est.name or "").lower()
        m = re.match(r"(heading|título|titulo)\s*(\d)", nome)
        if m:
            return int(m.group(2))
        ppr = est.element.find(qn("w:pPr"))
        if ppr is not None:
            ol = ppr.find(qn("w:outlineLvl"))
            if ol is not None and int(ol.get(qn("w:val"))) < 9:
                return int(ol.get(qn("w:val"))) + 1
    ppr = par._p.find(qn("w:pPr"))
    if ppr is not None:
        ol = ppr.find(qn("w:outlineLvl"))
        if ol is not None and int(ol.get(qn("w:val"))) < 9:
            return int(ol.get(qn("w:val"))) + 1
    return None


@por_elemento
def tem_numeracao_automatica(par):
    ppr = par._p.find(qn("w:pPr"))
    if ppr is not None and ppr.find(qn("w:numPr")) is not None:
        return True
    for est in cadeia_par(par):
        ppr = est.element.find(qn("w:pPr"))
        if ppr is not None and ppr.find(qn("w:numPr")) is not None:
            return True
    return False


def campos_pagina(doc):
    """Onde está o campo PAGE: lista de (seção, 'cabeçalho'|'rodapé', alinhamento)."""
    achados = []
    for i, sec in enumerate(doc.sections):
        for tipo, partes in (("cabeçalho", (sec.header, sec.first_page_header, sec.even_page_header)),
                             ("rodapé", (sec.footer, sec.first_page_footer, sec.even_page_footer))):
            for parte in partes:
                try:
                    if parte.is_linked_to_previous:
                        continue
                    el = parte._element
                except Exception:
                    continue
                xml = el.xml
                if re.search(r"\bPAGE\b", xml):
                    alin = None
                    for p in parte.paragraphs:
                        if re.search(r"\bPAGE\b", p._p.xml):
                            alin = alinhamento(p)
                    achados.append((i, tipo, alin))
    return achados


def margens_espelhadas(doc):
    st = doc.settings.element
    return st.find(qn("w:mirrorMargins")) is not None
