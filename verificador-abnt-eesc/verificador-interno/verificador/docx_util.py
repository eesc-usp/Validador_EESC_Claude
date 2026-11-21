"""Leitura do .docx com propriedades EFETIVAS (run -> estilo -> padrão do documento).

O Word guarda formatação em camadas: um parágrafo pode não ter tamanho de fonte
próprio e herdar do estilo, que herda do estilo base, que herda do padrão do
documento. Estas funções resolvem essa cadeia para a verificação não errar.
"""
import re
from collections import Counter

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
from docx.shared import Length
from docx.table import Table
from docx.text.paragraph import Paragraph

EMU_POR_CM = 360000


def cm(valor):
    return None if valor is None else round(valor / EMU_POR_CM, 2)


# ---------------------------------------------------------------- blocos
def blocos(doc):
    """Percorre o corpo na ordem real: ('p', Paragraph) ou ('t', Table)."""
    for el in doc.element.body.iterchildren():
        if el.tag == qn("w:p"):
            yield "p", Paragraph(el, doc)
        elif el.tag == qn("w:tbl"):
            yield "t", Table(el, doc)


def texto(par):
    """Texto visível: inclui o que está em alterações controladas inseridas e exclui o apagado."""
    from .correcao import texto_atual
    return texto_atual(par).strip()


# ---------------------------------------------------------------- padrões do documento
def _doc_defaults(doc):
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


# ---------------------------------------------------------------- fonte
def tamanho_run(run, par, doc):
    if run.font.size is not None:
        return run.font.size.pt
    if run.style is not None:
        for est in _cadeia(run.style):
            if est.font.size is not None:
                return est.font.size.pt
    for est in _cadeia(par.style):
        if est.font.size is not None:
            return est.font.size.pt
    return _doc_defaults(doc)[0] or 10.0  # Word usa 10 pt se nada for definido


def tamanho_predominante(par, doc):
    """Tamanho de fonte que cobre mais caracteres do parágrafo."""
    cont = Counter()
    for r in par.runs:
        if r.text.strip():
            cont[tamanho_run(r, par, doc)] += len(r.text)
    return cont.most_common(1)[0][0] if cont else None


def negrito_run(run, par):
    if run.bold is not None:
        return run.bold
    if run.style is not None:
        for est in _cadeia(run.style):
            if est.font.bold is not None:
                return est.font.bold
    for est in _cadeia(par.style):
        if est.font.bold is not None:
            return est.font.bold
    return False


def italico_run(run, par):
    if run.italic is not None:
        return run.italic
    for est in _cadeia(par.style):
        if est.font.italic is not None:
            return est.font.italic
    return False


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


def tem_negrito(par):
    return any(negrito_run(r, par) for r in par.runs if r.text.strip())


def tem_italico(par):
    return any(italico_run(r, par) for r in par.runs if r.text.strip())


def todo_negrito(par):
    runs = [r for r in par.runs if r.text.strip()]
    return bool(runs) and all(negrito_run(r, par) for r in runs)


# ---------------------------------------------------------------- parágrafo
def _pf(par, attr):
    v = getattr(par.paragraph_format, attr)
    if v is not None:
        return v
    for est in _cadeia(par.style):
        v = getattr(est.paragraph_format, attr)
        if v is not None:
            return v
    return None


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


def alinhamento(par):
    a = _pf(par, "alignment")
    return a if a is not None else WD_ALIGN_PARAGRAPH.LEFT


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


def tem_imagem(par):
    return bool(par._p.xpath(".//w:drawing")) or bool(par._p.xpath(".//w:pict"))


def nivel_titulo_estilo(par):
    """Nível do título pelo estilo (Heading N / Título N / nível de tópico)."""
    for est in _cadeia(par.style):
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


def tem_numeracao_automatica(par):
    ppr = par._p.find(qn("w:pPr"))
    if ppr is not None and ppr.find(qn("w:numPr")) is not None:
        return True
    for est in _cadeia(par.style):
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
