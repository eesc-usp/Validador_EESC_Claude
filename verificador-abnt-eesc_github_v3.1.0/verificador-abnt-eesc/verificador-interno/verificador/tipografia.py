"""Verificações que não dependem do documento seguir o modelo.

Funcionam com qualquer trabalho, mesmo sem estilos do Word nem a estrutura do template:
- títulos identificados por aproximação (grafia diferente, "CAPÍTULO 1 – ...", pontuação no fim);
- títulos de seção reconhecidos pelo número mesmo sem estilo: ponto no fim, centralizados, espaço/tab
  depois do número, texto na mesma linha do título;
- legendas fora do padrão ("FIGURA 1", "Fig. 1") e imagens sem título;
- texto em cor/realce, mistura de fontes, alinhamento e recuo de parágrafo desiguais;
- espaços duplos, espaço antes da pontuação, parênteses e aspas sem par;
- notas de rodapé e número de página em fonte 10/11 e espaço simples (Diretrizes 2.2 e 2.2.1);
- sigla por extenso na 1ª menção (Diretrizes 2.4);
- sumário sem elementos pré-textuais e "Índice" usado no lugar do sumário (NBR 6027);
- ano de referência posterior ao ano atual.
"""
import datetime as _dt
import re
from collections import Counter

from docx.enum.dml import MSO_THEME_COLOR
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree

from . import docx_util as du
from . import perfil as PF
from .cache import qn
from .documento import chave
from .estrutura import CAT_FORM, CAT_ILU, CAT_SEC, RE_SECAO, RE_SUMARIO, _amostra
from .modelo import AJUSTAR, OK, VERIFICAR, Achado, T

CAT_TIPO = "Texto e tipografia"

RE_LEGENDA_FLEX = re.compile(
    r"^(figura|fig\.|quadro|gr[áa]fico|fotografia|mapa|esquema|fluxograma|desenho|organograma|planta|imagem|"
    r"retrato|tabela|tab\.)\s*(\d+(?:\.\d+)?)\s*([-–—:.]?)\s*(.*)$", re.I)
RE_LISTA_INI = re.compile(r"^\s*([a-z]|\d{1,2}|[ivx]{1,4})\)\s")
SIGLAS_IGNORAR = {"DE", "DA", "DO", "DOS", "DAS", "E", "EM", "NO", "NA", "AO", "OS", "AS", "UM", "OU", "II", "III",
                  "IV", "VI", "VII", "VIII", "IX", "XI", "XII", "XX", "XXI", "OK", "ET", "AL", "CM", "MM", "KG",
                  "USP", "EESC", "ABNT", "NBR", "ISO", "DOI", "URL", "PDF", "ISBN", "ISSN", "EUA", "UK", "BR", "SP",
                  "RJ", "MG", "PR", "RS", "SC", "BA", "PE", "CE", "GO", "DF", "ES", "PA", "AM", "FIGURA", "TABELA"}


def _corpo(d):
    """Parágrafos de texto corrido (parte textual), sem títulos, legendas, fontes e alíneas."""
    out = []
    for b in d.paragrafos("textual"):
        t = b.texto
        if len(t) < 60 or RE_LEGENDA_FLEX.match(t) or re.match(r"^Fonte\b", t) or "\t" in b.bruto:
            continue
        if du.nivel_titulo_estilo(b.obj):
            continue
        out.append(b)
    return out


# ------------------------------------------------------------------ títulos aproximados e títulos flexíveis
def titulos_aproximados(d):
    todos = getattr(d, "aproximados", [])
    out = []
    primeira = [b for c, b in todos if c == "PRIMEIRA SECAO"]
    if primeira:
        out.append(Achado(CAT_SEC, "Início do texto sem o título INTRODUÇÃO", VERIFICAR,
                          f"a parte textual foi considerada a partir de “{primeira[0].texto[:60]}”",
                          "A primeira seção do texto é a Introdução: confira se o título deveria ser "
                          "'1 INTRODUÇÃO'.", "3.2", [T(primeira[0].texto, primeira[0])]))
    lst = [(c, b) for c, b in todos if c != "PRIMEIRA SECAO"]
    if not lst:
        return out
    nomes = {"INTRODUCAO": "INTRODUÇÃO", "REFERENCIAS": "REFERÊNCIAS"}
    return out + [Achado(CAT_SEC, "Títulos identificados por aproximação", VERIFICAR,
                   f"{len(lst)} título(s) com grafia diferente da esperada",
                   "O verificador reconheceu estes títulos pela semelhança. Confira a grafia: o título deve ser "
                   "exatamente o nome do elemento (ex.: 'REFERÊNCIAS', '1 INTRODUÇÃO').", "2.3",
                   [f"{b.texto} (≈ {nomes.get(c, c)})" for c, b in lst], blocos=[b.i for _, b in lst])]


def _titulo_flex(d, b):
    """Título de seção numerado, reconhecido mesmo sem estilo/negrito (o que o modelo não garante)."""
    t = b.texto
    if not t or len(t) > 150 or "\t" in b.bruto or RE_SUMARIO.search(t) or d.regiao(b.i) == "pre":
        return None
    m = RE_SECAO.match(t)
    if not m or len(t.split()) > 18 or re.search(r"[.;!?]\s+\S", t.rstrip(".")):
        return None
    if re.match(r"^\d+(\.\d+)?\s*(%|cm|mm|m|kg|g|s|h|min|°|ºC|MPa|kN|N|x|vezes)\b", t):
        return None
    prox = next((x for x in d.blocos[b.i + 1:b.i + 4] if x.tipo == "t" or x.texto), None)
    if prox is None:
        return None
    destaque = du.nivel_titulo_estilo(b.obj) or du.todo_negrito(b.obj) or \
        m.group(4).upper() == m.group(4) or (prox.tipo == "p" and len(prox.texto) > len(t) * 2)
    return m if destaque else None


def titulos_flexiveis(d):
    ponto_final, centralizados, espaco, mesma_linha = [], [], [], []
    for b in d.paragrafos():
        if d.regiao(b.i) == "pre":
            continue
        m = _titulo_flex(d, b)
        if m:
            if b.texto.rstrip().endswith((".", ":", ";")) and not b.texto.rstrip().endswith("..."):
                ponto_final.append((b.texto, b.i))
            if du.alinhamento(b.obj) == WD_ALIGN_PARAGRAPH.CENTER:
                centralizados.append((b.texto, b.i))
            if re.match(r"^\d{1,2}(?:\.\d{1,2}){0,6}\.?(\t|\s{2,})", b.bruto or b.texto):
                espaco.append((b.texto, b.i))
            continue
        mm = re.match(r"^(\d{1,2}(?:\.\d{1,2}){0,5})\.?\s+([A-ZÀ-Ý][^.:]{2,70})[.:]\s+([A-ZÀ-Ý].{80,})$", b.texto)
        if mm and d.regiao(b.i) == "textual" and not RE_LISTA_INI.match(b.texto) and \
                (du.tem_negrito(b.obj) or mm.group(2).upper() == mm.group(2)):
            mesma_linha.append((b.texto[:120], b.i))
    out = []
    for item, lst, corr, st in [
        ("Título de seção sem ponto final", ponto_final,
         "Título de seção não termina com ponto, dois-pontos ou ponto e vírgula.", AJUSTAR),
        ("Título de seção alinhado à esquerda", centralizados,
         "Títulos com indicativo numérico ficam alinhados à esquerda (só os sem número são centralizados).", AJUSTAR),
        ("Um espaço entre o número e o título", espaco,
         "Separe o indicativo do título por um único espaço (sem tabulação ou vários espaços).", AJUSTAR),
        ("Texto começa em outra linha depois do título", mesma_linha,
         "O título fica sozinho no parágrafo; o texto da seção começa na linha seguinte.", VERIFICAR),
    ]:
        out.append(Achado(CAT_SEC, item, st if lst else OK, f"{len(lst)} título(s)" if lst else "sem problemas",
                          corr if lst else "", "2.3", _amostra([x for x, _ in lst]),
                          blocos=[i for _, i in lst[:300]]))
    return out


# ------------------------------------------------------------------ ilustrações
def legendas(d):
    fora_padrao, sem_titulo = [], []
    pdf = _convertido_de_pdf(d)
    for b in d.paragrafos():
        if d.regiao(b.i) == "pre":
            continue
        m = RE_LEGENDA_FLEX.match(b.texto)
        if m and "\t" not in b.bruto and len(b.texto) < 250:
            palavra = m.group(1)
            padrao = palavra[:1].upper() + palavra[1:].lower()
            if (palavra != padrao or palavra.lower() in ("fig.", "tab.")) and m.group(3):
                fora_padrao.append((b.texto, b.i))
        if du.tem_imagem(b.obj) and d.regiao(b.i) != "pre" and not pdf:
            viz = [x for x in d.blocos[max(0, b.i - 3):b.i + 4] if x.tipo == "p" and x.texto and x.i != b.i]
            if not any(RE_LEGENDA_FLEX.match(x.texto) for x in viz) and \
                    not RE_LEGENDA_FLEX.match(b.texto or ""):
                sem_titulo.append((f"imagem perto de: “{(viz[0].texto if viz else '')[:70]}”", b.i))
    return [
        Achado(CAT_ILU, "Palavra designativa da legenda no padrão", AJUSTAR if fora_padrao else OK,
               f"{len(fora_padrao)} legenda(s)" if fora_padrao else "sem problemas",
               "Escreva a palavra por extenso, só com a inicial maiúscula: 'Figura 1 – Título', "
               "'Tabela 2 – Título' (não 'FIGURA 1', 'Fig. 1').", "2.6",
               _amostra([x for x, _ in fora_padrao]), blocos=[i for _, i in fora_padrao[:300]]),
        Achado(CAT_ILU, "Toda imagem com título (legenda)", VERIFICAR if sem_titulo else OK,
               f"{len(sem_titulo)} imagem(ns) sem 'Figura N – ...' por perto" if sem_titulo else "todas com título",
               "Toda ilustração tem título acima ('Figura 1 – ...') e fonte abaixo. Se for uma equação ou logotipo, "
               "ignore.", "2.6", [x for x, _ in sem_titulo][:60], blocos=[i for _, i in sem_titulo[:300]]),
    ]


# ------------------------------------------------------------------ formatação do texto corrido
def _fonte_tema(doc):
    try:
        for rel in doc.part.rels.values():
            if rel.reltype.endswith("/theme"):
                xml = etree.fromstring(rel.target_part.blob)
                ns = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
                minor = xml.find(".//a:minorFont/a:latin", ns)
                major = xml.find(".//a:majorFont/a:latin", ns)
                return (minor.get("typeface") if minor is not None else None,
                        major.get("typeface") if major is not None else None)
    except Exception:
        pass
    return None, None


def _fonte_run(run, par, doc, tema):
    if run.font.name:
        return run.font.name
    rf = run._r.rPr.find(qn("w:rFonts")) if run._r.rPr is not None else None
    if rf is not None and rf.get(qn("w:asciiTheme")):
        return tema[1] if "major" in rf.get(qn("w:asciiTheme")) else tema[0]
    nome = du.fonte_nos_estilos(run, par, "name", estilo_do_run=False, so_verdadeiro=True)
    if nome:
        return nome
    try:
        d_rf = doc.styles.element.find(qn("w:docDefaults")).find(".//" + qn("w:rFonts"))
        if d_rf is not None:
            if d_rf.get(qn("w:ascii")):
                return d_rf.get(qn("w:ascii"))
            if d_rf.get(qn("w:asciiTheme")):
                return tema[1] if "major" in d_rf.get(qn("w:asciiTheme")) else tema[0]
    except Exception:
        pass
    return None


def _convertido_de_pdf(d):
    from .estrutura import origem_do_arquivo
    try:
        return bool(origem_do_arquivo(d))
    except Exception:
        return False


def formatacao_texto(d):
    corpo = _corpo(d)
    if not corpo:
        return []
    poucos = len(corpo) < 5  # com pouco texto, só a cor é conferida (comparar padrões não faz sentido)
    pdf = _convertido_de_pdf(d)  # vindo de PDF, recuos e alinhamentos são artificiais: não comparar
    tema = _fonte_tema(d.doc)
    fontes, alinh, recuos = Counter(), Counter(), Counter()
    por_par = []
    coloridos = []
    for b in corpo:
        p = b.obj
        f = Counter()
        for r in p.runs:
            t = r.text
            if not t.strip():
                continue
            if r._r.getparent().tag == qn("w:hyperlink") or (r.style is not None and "Hyperlink" in (r.style.name or "")):
                continue
            nome = _fonte_run(r, p, d.doc, tema)
            if nome:
                f[nome] += len(t)
            cor = r.font.color
            preta = cor is None or cor.type is None or (cor.rgb is not None and str(cor.rgb) in ("000000", "auto")) \
                or (cor.theme_color in (MSO_THEME_COLOR.TEXT_1, MSO_THEME_COLOR.DARK_1))
            realce = r.font.highlight_color is not None
            if (not preta or realce) and len(t.strip()) > 2 and b.i not in {i for _, i in coloridos}:
                coloridos.append((f"{'realce' if realce else 'cor'}: “{t.strip()[:60]}”", b.i))
        principal = f.most_common(1)[0][0] if f else None
        if principal:
            fontes[principal] += 1
        a = du.alinhamento(p)
        alinh[a] += 1
        fl = du._pf(p, "first_line_indent")
        recuos[round(fl.cm, 1) if fl is not None else 0.0] += 1
        por_par.append((b, principal, a, round(fl.cm, 1) if fl is not None else 0.0))
    out = [Achado(CAT_FORM, "Texto na cor preta, sem realce", VERIFICAR if coloridos else OK,
                  f"{len(coloridos)} trecho(s) coloridos ou marcados" if coloridos else "sem problemas",
                  "O texto deve ser digitado em preto (outras cores só nas ilustrações). Remova realces e "
                  "cores esquecidos da revisão.", "2.2", [x for x, _ in coloridos][:300],
                  blocos=[i for _, i in coloridos[:300]])]
    if poucos:
        return out
    f_mais = fontes.most_common(1)[0][0] if fontes else None
    outras = [(f"[{fp}] {b.texto}", b.i) for b, fp, _, _ in por_par if fp and f_mais and fp != f_mais]
    out.append(Achado(CAT_FORM, "Mesma família de fonte no texto", VERIFICAR if outras else OK,
                      f"{len(outras)} parágrafo(s) fora da fonte principal ({f_mais})" if outras else (f_mais or "—"),
                      f"Use uma só fonte no texto ({f_mais}). Fontes diferentes costumam vir de trechos colados."
                      if outras else "", "2.2", _amostra([x for x, _ in outras]), blocos=[i for _, i in outras[:300]]))
    a_mais = alinh.most_common(1)[0][0]
    alin_dif = [(b.texto, b.i) for b, _, a, _ in por_par if a != a_mais and
                a in (WD_ALIGN_PARAGRAPH.JUSTIFY, WD_ALIGN_PARAGRAPH.LEFT, None, WD_ALIGN_PARAGRAPH.CENTER,
                      WD_ALIGN_PARAGRAPH.RIGHT)]
    nomes_a = {WD_ALIGN_PARAGRAPH.JUSTIFY: "justificado", WD_ALIGN_PARAGRAPH.LEFT: "à esquerda", None: "à esquerda",
               WD_ALIGN_PARAGRAPH.CENTER: "centralizado", WD_ALIGN_PARAGRAPH.RIGHT: "à direita"}
    if not pdf and len(alin_dif) <= max(3, len(por_par) * 0.3):
        out.append(Achado(CAT_FORM, "Alinhamento igual nos parágrafos do texto", VERIFICAR if alin_dif else OK,
                          f"{len(alin_dif)} parágrafo(s) diferentes do padrão ({nomes_a.get(a_mais, '?')})"
                          if alin_dif else nomes_a.get(a_mais, "uniforme"),
                          "Mantenha o mesmo alinhamento em todo o texto corrido." if alin_dif else "", "2.2",
                          _amostra([x for x, _ in alin_dif]), blocos=[i for _, i in alin_dif[:300]]))
    r_mais = recuos.most_common(1)[0][0]
    rec_dif = [(f"[recuo {r:g} cm] {b.texto}", b.i) for b, _, _, r in por_par if abs(r - r_mais) > 0.2]
    if not pdf and len(rec_dif) <= max(3, len(por_par) * 0.3):
        out.append(Achado(CAT_FORM, "Recuo de primeira linha igual nos parágrafos", VERIFICAR if rec_dif else OK,
                          f"{len(rec_dif)} parágrafo(s) com recuo diferente de {r_mais:g} cm" if rec_dif
                          else f"{r_mais:g} cm", "Use o mesmo recuo de primeira linha em todos os parágrafos."
                          if rec_dif else "", "2.2", _amostra([x for x, _ in rec_dif]),
                          blocos=[i for _, i in rec_dif[:300]]))
    return out


# ------------------------------------------------------------------ tipografia
RE_ESP_DUPLO = re.compile(r"(?<=\S) {2,}(?=\S)")
RE_ESP_PONT = re.compile(r"(?<=[\wÀ-ÿ)\]”\"]) +(?=[,;:.!?](?!\.\.)(?!\d))")
RE_ESP_PAR = re.compile(r"\( +|(?<=\S) +\)")


def tipografia(d):
    duplos, antes_pont, parenteses, aspas = [], [], [], []
    for b in d.paragrafos("textual") + d.paragrafos("pos"):
        t = b.bruto or ""
        if not t.strip() or "\t" in t:
            continue
        if RE_ESP_DUPLO.search(t):
            duplos.append((t.strip(), b.i))
        if RE_ESP_PONT.search(t) or RE_ESP_PAR.search(t):
            antes_pont.append((t.strip(), b.i))
        sem_lista = RE_LISTA_INI.sub("", t)
        sem_lista = re.sub(r"(?<=\s)[a-z0-9]\)(?=\s)", "", sem_lista)
        if sem_lista.count("(") != sem_lista.count(")") and len(t) < 2000:
            parenteses.append((t.strip(), b.i))
        if (t.count("“") != t.count("”")) or (t.count('"') % 2 == 1):
            aspas.append((t.strip(), b.i))
    out = []
    for item, lst, corr, st in [
        ("Sem espaços duplos entre palavras", duplos, "Troque os espaços repetidos por um só (a correção "
                                                      "automática faz isso).", AJUSTAR),
        ("Sem espaço antes da pontuação", antes_pont, "Sem espaço antes de vírgula, ponto, dois-pontos, ponto e "
                                                      "vírgula e fechamento de parênteses; nem depois de '('.", AJUSTAR),
        ("Parênteses abertos e fechados", parenteses, "Confira se algum parêntese ficou sem fechar (ou sobrando).",
         VERIFICAR),
        ("Aspas abertas e fechadas", aspas, "Confira se alguma citação ficou com as aspas sem fechar.", VERIFICAR),
    ]:
        out.append(Achado(CAT_TIPO, item, st if lst else OK, f"{len(lst)} parágrafo(s)" if lst else "sem problemas",
                          corr if lst else "", "2.2", _amostra([x for x, _ in lst]),
                          blocos=[i for _, i in lst[:300]]))
    return out


# ------------------------------------------------------------------ notas de rodapé e número de página
def _estilos_xml(doc):
    est = {}
    for s in doc.styles.element.findall(qn("w:style")):
        est[s.get(qn("w:styleId"))] = s
    return est


def _valor_estilo(estilos, sid, caminho, attr):
    vistos = 0
    while sid and sid in estilos and vistos < 10:
        el = estilos[sid].find(caminho)
        if el is not None and el.get(attr):
            return el.get(attr)
        base = estilos[sid].find(qn("w:basedOn"))
        sid = base.get(qn("w:val")) if base is not None else None
        vistos += 1
    return None


def notas_e_pagina(d):
    out = []
    lim = PF.atual()["fonte_legenda_max"]
    estilos = _estilos_xml(d.doc)
    padrao_sz = None
    try:
        dsz = d.doc.styles.element.find(qn("w:docDefaults")).find(".//" + qn("w:sz"))
        padrao_sz = int(dsz.get(qn("w:val"))) / 2 if dsz is not None else None
    except Exception:
        pass
    notas_xml = None
    for rel in d.doc.part.rels.values():
        if rel.reltype.endswith("/footnotes"):
            try:
                notas_xml = etree.fromstring(rel.target_part.blob)
            except Exception:
                notas_xml = None
    if notas_xml is not None:
        grandes, espacadas, n = [], [], 0
        for fn in notas_xml.findall(qn("w:footnote")):
            if fn.get(qn("w:type")) in ("separator", "continuationSeparator", "continuationNotice"):
                continue
            n += 1
            for p in fn.findall(qn("w:p")):
                texto = "".join(x.text or "" for x in p.iter(qn("w:t"))).strip()
                if not texto:
                    continue
                ppr = p.find(qn("w:pPr"))
                sid = ppr.find(qn("w:pStyle")).get(qn("w:val")) if ppr is not None and ppr.find(qn("w:pStyle")) \
                    is not None else "FootnoteText"
                tams = []
                for r in p.findall(".//" + qn("w:r")):
                    if not "".join(x.text or "" for x in r.iter(qn("w:t"))).strip():
                        continue
                    sz = r.find(qn("w:rPr") + "/" + qn("w:sz")) if r.find(qn("w:rPr")) is not None else None
                    v = sz.get(qn("w:val")) if sz is not None else _valor_estilo(estilos, sid, qn("w:rPr") + "/" +
                                                                               qn("w:sz"), qn("w:val"))
                    tams.append(int(v) / 2 if v else padrao_sz)
                tams = [x for x in tams if x]
                if tams and max(tams) > lim:
                    grandes.append(f"[fonte {max(tams):g}] {texto}")
                linha = None
                if ppr is not None and ppr.find(qn("w:spacing")) is not None:
                    linha = ppr.find(qn("w:spacing")).get(qn("w:line"))
                linha = linha or _valor_estilo(estilos, sid, qn("w:pPr") + "/" + qn("w:spacing"), qn("w:line"))
                if linha and int(linha) > 280:
                    espacadas.append(texto)
        if n:
            probs = grandes + [f"[entrelinhas > simples] {t}" for t in espacadas]
            out.append(Achado(CAT_FORM, "Notas de rodapé em fonte menor e espaço simples", AJUSTAR if probs else OK,
                              f"{len(probs)} nota(s) fora do padrão de {n}" if probs else f"{n} nota(s) conferidas",
                              f"Notas de rodapé em fonte {PF.fmt(min(10, lim))} ou {PF.fmt(lim)} e espaço simples: "
                              "ajuste o estilo 'Texto de nota de rodapé'." if probs else "", "2.2.1", _amostra(probs)))
    # número de página
    grandes = []
    for sec in d.doc.sections:
        for parte in (sec.header, sec.first_page_header, sec.even_page_header,
                      sec.footer, sec.first_page_footer, sec.even_page_footer):
            try:
                if parte.is_linked_to_previous:
                    continue
                for p in parte.paragraphs:
                    if re.search(r"\bPAGE\b", p._p.xml):
                        tam = du.tamanho_predominante(p, d.doc)
                        if tam is None:
                            szs = [int(x) / 2 for x in re.findall(r'<w:sz w:val="(\d+)"', p._p.xml)]
                            tam = max(szs) if szs else None
                        if tam and tam > lim:
                            grandes.append(f"número de página em fonte {tam:g}")
            except Exception:
                continue
    if grandes:
        out.append(Achado(CAT_FORM, "Número de página em fonte menor", VERIFICAR, grandes[0],
                          f"A paginação usa fonte menor (10 ou {PF.fmt(lim)}).", "2.2"))
    return out


# ------------------------------------------------------------------ siglas, sumário, anos
def siglas_primeira_mencao(d):
    corpo = _corpo(d)
    textos = [(b.texto, b.i) for b in corpo if sum(c.isupper() for c in b.texto) < 0.3 * max(1, sum(
        c.isalpha() for c in b.texto))]
    from .referencias import extrair_referencias
    sobrenomes = {tok.upper() for r in extrair_referencias(d) for tok in r.tokens[:8]}
    usos = Counter()
    primeira = {}
    for t, i in textos:
        for m in re.finditer(r"(?<![\wÀ-ÿ])([A-Z]{2,7})(?![\wÀ-ÿ])", t):
            s = m.group(1)
            antes, depois = t[max(0, m.start() - 25):m.start()], t[m.end():m.end() + 25]
            # parte de uma frase em maiúsculas ("NA PARTE SUPERIOR") ou autor numa chamada (SILVA, 2020)
            if re.search(r"[A-ZÀ-Ý]{2,}[\s,;]+$", antes) or re.match(r"^[\s,;]+(e\s+)?[A-ZÀ-Ý]{2,}", depois) or \
                    re.match(r"^,\s*(\d{4}|\[)", depois) or s in sobrenomes or chave(s) in sobrenomes:
                continue
            usos[s] += 1
            if s not in primeira:
                primeira[s] = (t, m.start(), m.end(), i)
    faltam = []
    for s, n in usos.items():
        if n < 2 or s in SIGLAS_IGNORAR or re.fullmatch(r"[IVXLC]+", s):
            continue
        t, a, b, i = primeira[s]
        entre_par = t[max(0, a - 1):a] == "(" and t[b:b + 1] in (")", ";", ",")
        com_extenso = re.match(r"\s*[(–—-]\s*[A-ZÀ-Ý][a-zà-ÿ]", t[b:b + 4])
        if not (entre_par or com_extenso):
            faltam.append((f"{s}: “…{t[max(0, a - 40):b + 20]}…”", i))
    return [Achado(CAT_TIPO, "Sigla por extenso na primeira menção", VERIFICAR if faltam else OK,
                   f"{len(faltam)} sigla(s) sem a forma completa na 1ª vez" if faltam else "sem problemas",
                   "Na primeira menção, escreva o nome completo seguido da sigla entre parênteses: "
                   "'Universidade de São Paulo (USP)'. Depois, use só a sigla.", "2.4",
                   [x for x, _ in faltam][:100], blocos=[i for _, i in faltam[:300]])]


PRE_NO_SUMARIO = r"^(RESUMO|ABSTRACT|RESUMEN|LISTA DE|SUMARIO|AGRADECIMENTO|DEDICATORIA|EPIGRAFE|ERRATA|" \
                 r"FOLHA DE APROVACAO|FICHA CATALOGRAFICA|DECLARACAO)"


def sumario_extra(d):
    out = []
    i_intro = d.i_intro if d.i_intro is not None else len(d.blocos)
    sum_b = d.achar_titulo(r"SUMARIO")
    ind_b = next((b for b in d.paragrafos() if b.i < i_intro and chave(b.texto).strip(" :.") == "INDICE"), None)
    if ind_b is not None and (sum_b is None or sum_b.i > i_intro):
        out.append(Achado(CAT_SEC, "Sumário com o nome 'SUMÁRIO'", AJUSTAR, f"título “{ind_b.texto}”",
                          "No início do trabalho, a lista das seções chama-se SUMÁRIO. Índice é outra coisa (lista "
                          "de assuntos/nomes no final, opcional).", "3.1.2.1.11", blocos=[ind_b.i]))
        sum_b = ind_b
    if sum_b is None:
        return out
    pre = []
    for b in d.blocos[sum_b.i + 1:min(i_intro, sum_b.i + 300)]:
        if b.tipo != "p" or not b.texto:
            continue
        if not (RE_SUMARIO.search(b.texto) or "\t" in b.bruto):
            continue
        k = re.sub(r"^\d+(\.\d+)*\.?\s*", "", chave(b.texto))
        if re.match(PRE_NO_SUMARIO, k):
            pre.append((b.texto.split("\t")[0].strip(" .…"), b.i))
    out.append(Achado(CAT_SEC, "Sumário sem os elementos pré-textuais", AJUSTAR if pre else OK,
                      f"{len(pre)} elemento(s) pré-textual(is) listados" if pre else "sem problemas",
                      "Resumo, abstract, listas, agradecimentos etc. não entram no sumário: ele começa na "
                      "Introdução (NBR 6027)." if pre else "", "3.1.2.1.11", [x for x, _ in pre],
                      blocos=[i for _, i in pre]))
    return out


def anos_referencias(d):
    from .referencias import extrair_referencias
    atual = _dt.date.today().year
    futuras = []
    for r in extrair_referencias(d):
        anos = [int(a[:4]) for a in r.anos if a[:4].isdigit()]
        if anos and min(anos) > atual + 1 or any(a > atual + 1 for a in anos[:1]):
            futuras.append((r.texto, r.i))
    return [Achado("Referências", "Ano da obra não posterior ao atual", VERIFICAR if futuras else OK,
                   f"{len(futuras)} referência(s) com ano depois de {atual}" if futuras else "sem problemas",
                   "Confira o ano: pode ser erro de digitação." if futuras else "", "6",
                   [x[:120] for x, _ in futuras], blocos=[i for _, i in futuras])]


def verificar(d):
    out = []
    for f in (titulos_aproximados, titulos_flexiveis, legendas, formatacao_texto, tipografia, notas_e_pagina,
              siglas_primeira_mencao, sumario_extra, anos_referencias):
        try:
            out += f(d)
        except Exception as e:
            out.append(Achado("Sistema", f"tipografia.{f.__name__}", VERIFICAR, f"verificação não concluída: {e}"))
    return out
