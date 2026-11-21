"""Elementos do trabalho: aparece? está formatado? o conteúdo é verdadeiro?

Não exige que cada elemento esteja na mesma página/posição do template: cada um é procurado no arquivo
inteiro. Para cada elemento o verificador responde três perguntas:

1. **Aparece?** – obrigatório, condicional ou opcional (conforme o tipo de trabalho e a norma escolhida);
2. **Está formatado?** – título sem número, centralizado e em maiúsculas; recuos e alinhamentos próprios;
3. **É verdade?** – o conteúdo confere com o resto do documento:
   autor, título e ano iguais na capa, folha de rosto, ficha e referência do resumo; orientador igual na folha
   de rosto e na ficha; nº de palavras-chave = nº de keywords; toda "Figura 3", "Tabela 2", "seção 2.1",
   "Apêndice A" citada existe; siglas da lista usadas no texto (e vice-versa) e em ordem alfabética;
   páginas do sumário e das listas conferem com a paginação que o Word gravou no arquivo; a epígrafe tem
   a obra nas referências; a errata tem o quadro "Folha / Linha / Onde se lê / Leia-se".
"""
import re
from collections import Counter

from docx.enum.text import WD_ALIGN_PARAGRAPH

from . import docx_util as du
from . import perfil as PF
from .cache import qn
from .documento import chave
from .estrutura import RE_ILU, RE_TAB, _amostra, _candidato_titulo
from .modelo import AJUSTAR, OK, VERIFICAR, Achado

CAT = "Elementos do trabalho"

OBRIG, COND, OPC = "obrigatório", "condicional", "opcional"


# ------------------------------------------------------------------ utilidades
def _pars(d):
    return [b for b in d.blocos if b.tipo == "p" and b.texto]


def _titulo(d, rx):
    """Primeiro parágrafo curto, em qualquer lugar do arquivo, cujo texto (sem número) é o título rx."""
    r = re.compile(rf"^(\d+(\.\d+)*\.?\s*[-–]?\s*)?({rx})\s*:?$")
    for b in _pars(d):
        if len(b.texto) <= 90 and "\t" not in b.bruto and r.match(chave(b.texto).strip()):
            return b
    return None


def _ate_proximo_titulo(d, b, limite=80):
    """Parágrafos da seção que começa em b (até o próximo título em maiúsculas ou limite)."""
    out = []
    for x in d.blocos[b.i + 1:b.i + 1 + limite]:
        if x.tipo != "p" or not x.texto:
            continue
        k = chave(x.texto).strip()
        if len(k) <= 80 and re.fullmatch(r"[A-Z0-9 \-–:.]{5,}", k) and k == k.upper() and \
                not re.match(r"^[A-Z][A-Z' \-]+,\s", x.texto) and "\t" not in x.bruto and out:
            break
        out.append(x)
    return out


def paginas(d):
    """{índice do bloco: página física} pela paginação que o Word grava ao salvar (lastRenderedPageBreak).

    Devolve {} quando o arquivo não tem essa informação (gerado por outro programa ou nunca salvo no Word).
    """
    corpo = d.doc.element.body
    if len(corpo.findall(".//" + qn("w:lastRenderedPageBreak"))) < 5:
        return {}
    pag, out = 1, {}
    for b in d.blocos:
        el = b.obj._p if b.tipo == "p" else b.obj._tbl
        antes_do_texto = True
        inicio = pag
        for x in el.iter(qn("w:lastRenderedPageBreak"), qn("w:t")):
            if x.tag == qn("w:t"):
                if (x.text or "").strip():
                    antes_do_texto = False
            else:
                pag += 1
                if antes_do_texto:
                    inicio = pag
        out[b.i] = inicio
    return out


# ------------------------------------------------------------------ dados do trabalho
def ref_do_resumo(d, titulo_rx=r"RESUMO"):
    tb = _titulo(d, titulo_rx)
    if not tb:
        return None
    for x in _ate_proximo_titulo(d, tb, 12):
        m = re.match(r"^([A-ZÀ-Ý][A-ZÀ-Ý' \-]+),\s*(.+?)\.\s+((?:19|20)\d{2})\.\s+(.+)$", x.texto)
        if m:
            sobrenome, resto, ano, tipo = m.groups()
            mi = re.match(r"^((?:[A-ZÀ-Ý]\.\s*)+)\s*(.+)$", resto)
            if mi:
                prenomes, titulo = mi.group(1).strip(), mi.group(2)
            else:
                partes = resto.split(". ", 1)
                prenomes, titulo = (partes[0], partes[1]) if len(partes) == 2 else ("", resto)
            return {"sobrenome": sobrenome.strip(), "prenomes": prenomes, "titulo": titulo.strip(" ."),
                    "ano": ano, "tipo": tipo, "bloco": x}
    return None


def _norm(t):
    return re.sub(r"[^A-Z0-9 ]", " ", chave(t)).split()


# ------------------------------------------------------------------ "é verdade?"
def consistencia(d):
    out = []
    pre = [b for b in _pars(d) if d.regiao(b.i) == "pre"] or _pars(d)[:150]
    ref = ref_do_resumo(d)
    ref_en = ref_do_resumo(d, r"ABSTRACT")
    excl = {x["bloco"].i for x in (ref, ref_en) if x}
    pre_txt = " ".join(" ".join(_norm(b.texto)) for b in pre if b.i not in excl)
    ficha_b = next((b for b in pre if re.search(r"CATALOGACAO|FICHA CATALOGRAFICA|\bCDD\b|\bCDU\b", chave(b.texto))
                    and not re.search(r"DEVE CONTER|PODE SER ELABORADA|DEVE SER ELABORADA", chave(b.texto))), None)
    ficha_txt = ""
    if ficha_b:
        ficha_txt = " ".join(" ".join(_norm(b.texto)) for b in pre if ficha_b.i <= b.i <= ficha_b.i + 25)

    if ref:
        probs = []
        sob = " ".join(_norm(ref["sobrenome"]))
        if sob and sob not in pre_txt:
            probs.append(f"o autor “{ref['sobrenome']}” não aparece na capa/folha de rosto")
        tit = " ".join(_norm(re.split(r"[:.]", ref["titulo"])[0]))[:60]
        if len(tit) > 8 and tit not in pre_txt:
            probs.append(f"o título “{ref['titulo'][:60]}” não aparece igual na capa/folha de rosto")
        anos_pre = set()
        for b in pre:   # linhas que são só o ano, ou "São Carlos, 2025" (capa, folha de rosto)
            for ln in (b.bruto or b.texto).split("\n"):
                m = re.fullmatch(r"\s*(?:S[ãa]o Carlos\s*[,\-–]?\s*)?((?:19|20)\d{2})\.?\s*", ln)
                if m:
                    anos_pre.add(m.group(1))
        anos = anos_pre | {ref["ano"]} | ({ref_en["ano"]} if ref_en else set())
        if len(anos) > 1:
            probs.append("anos diferentes entre capa, folha de rosto e resumo/abstract: " + ", ".join(sorted(anos)))
        if ref_en:
            if " ".join(_norm(ref_en["sobrenome"])) != sob:
                probs.append(f"autor do abstract (“{ref_en['sobrenome']}”) diferente do resumo")
        if ficha_txt:
            if sob and sob not in ficha_txt:
                probs.append("a ficha catalográfica traz outro autor")
            if ref["ano"] not in ficha_txt:
                probs.append("a ficha catalográfica traz outro ano")
        out.append(Achado(CAT, "Autor, título e ano iguais em todo o trabalho", AJUSTAR if probs else OK,
                          "; ".join(probs) if probs else
                          f"{ref['sobrenome']} · {ref['titulo'][:50]} · {ref['ano']}",
                          "Autor, título e ano devem ser idênticos na capa, folha de rosto, ficha catalográfica e "
                          "nas referências do resumo e do abstract." if probs else "", "3.1",
                          blocos=[ref["bloco"].i]))

    # orientador: folha de rosto × ficha
    ori = next((b for b in pre if re.match(r"^(ORIENTADOR|ORIENTADORA|ORIENTADOR\(A\)|ADVISOR|SUPERVISOR)\b",
                                            chave(b.texto))), None)
    if ori and ficha_txt and re.search(r"ORIENTADOR", ficha_txt):
        nome = [t for t in _norm(ori.texto.split(":", 1)[-1]) if len(t) > 3 and t not in
                ("PROF", "PROFA", "DRA", "DOUTOR", "DOUTORA")]
        if nome and nome[-1] not in ficha_txt:
            out.append(Achado(CAT, "Orientador igual na folha de rosto e na ficha", AJUSTAR,
                              f"“{ori.texto[:60]}” não confere com a ficha catalográfica",
                              "Use o mesmo nome do(a) orientador(a) na folha de rosto e na ficha.", "3.1.2.1.2",
                              blocos=[ori.i]))

    # palavras-chave × keywords
    def termos(rx):
        b = next((x for x in _pars(d) if re.match(rx, chave(x.texto))), None)
        if not b:
            return None, None
        corpo = b.texto.split(":", 1)[-1]
        return [t for t in re.split(r"[.;]\s*", corpo) if t.strip()], b
    pc, bpc = termos(r"PALAVRAS[- ]?CHAVES?\s*:")
    kw, bkw = termos(r"KEY[- ]?WORDS?\s*:")
    if pc is not None and kw is not None:
        out.append(Achado(CAT, "Mesmo número de palavras-chave e keywords", OK if len(pc) == len(kw) else VERIFICAR,
                          f"{len(pc)} palavras-chave e {len(kw)} keywords",
                          "As keywords são a tradução das palavras-chave: mesma quantidade e mesma ordem."
                          if len(pc) != len(kw) else "", "3.1.2.1.9", blocos=[bkw.i]))
    return out


def mencoes(d):
    """Toda 'Figura 3', 'Tabela 2', 'seção 2.1', 'Apêndice A' citada no texto existe."""
    existentes = set()
    for b in _pars(d):
        m = RE_ILU.match(b.texto)
        if m:
            existentes.add((m.group(1).capitalize().replace("Grafico", "Gráfico"), m.group(2)))
        m = RE_TAB.match(b.texto)
        if m:
            existentes.add(("Tabela", m.group(1)))
    secoes = set()
    for b in d.paragrafos():
        m = _candidato_titulo(b) if d.regiao(b.i) != "pre" else None
        if m and m != "auto":
            secoes.add(m.group(1))
        mm = re.match(r"^(\d{1,2}(?:\.\d{1,2}){0,5})\.?\s+\S", b.texto)
        if mm and len(b.texto) < 150 and d.regiao(b.i) != "pre":
            secoes.add(mm.group(1))
    autos = [b for b in d.paragrafos() if _candidato_titulo(b) == "auto"]
    anexos = {(m.group(1), m.group(2)) for b in _pars(d)
              for m in [re.match(r"^(APENDICE|ANEXO)\s+([A-Z]{1,2})\b", chave(b.texto))] if m}
    faltam = []
    for b in d.paragrafos("textual") + d.paragrafos("pos"):
        t = b.texto
        if RE_ILU.match(t) or RE_TAB.match(t) or re.match(r"^Fonte\b", t) or len(t) < 15:
            continue
        for m in re.finditer(r"\b(Figuras?|Fig\.|Gr[áa]ficos?|Quadros?|Tabelas?|Fotografias?|Mapas?)\s+"
                             r"(\d+(?:\.\d+)?)(?:\s*(?:,|e|a)\s*(\d+(?:\.\d+)?))?", t):
            tipo = m.group(1).rstrip("s").replace("Fig.", "Figura").capitalize().replace("Grafico", "Gráfico")
            for n in filter(None, (m.group(2), m.group(3))):
                if (tipo, n) not in existentes and existentes:
                    faltam.append((f"{tipo} {n}: “…{t[max(0, m.start() - 30):m.end() + 20]}…”", b.i))
        if secoes and not autos:
            for m in re.finditer(r"\b(?:[Ss]e[çc][ãa]o|[Ss]ubse[çc][ãa]o|[Cc]ap[íi]tulo|[Ii]tem)\s+"
                                 r"(\d{1,2}(?:\.\d{1,2}){0,5})\b(?!\s*%)", t):
                if m.group(1) not in secoes:
                    faltam.append((f"seção {m.group(1)}: “…{t[max(0, m.start() - 30):m.end() + 20]}…”", b.i))
        for m in re.finditer(r"\b(Ap[êe]ndice|Anexo)\s+([A-Z]{1,2})\b", t):
            alvo = "APENDICE" if m.group(1).lower().startswith("ap") else "ANEXO"
            if (alvo, m.group(2)) not in anexos:
                faltam.append((f"{m.group(1)} {m.group(2)}: “…{t[max(0, m.start() - 30):m.end() + 20]}…”", b.i))
    faltam = list(dict.fromkeys(faltam))
    return [Achado(CAT, "Figuras, tabelas, seções e apêndices citados existem", AJUSTAR if faltam else OK,
                   f"{len(faltam)} menção(ões) a algo que não existe no documento" if faltam else "todas existem",
                   "O texto cita um número que não existe (numeração mudou?). Confira a numeração ou use "
                   "referência cruzada do Word (Referências > Referência cruzada), que se atualiza sozinha."
                   if faltam else "", "2.6", _amostra([x for x, _ in faltam]), blocos=[i for _, i in faltam[:300]])]


SIGLAS_COMUNS = {"DE", "DA", "DO", "DOS", "DAS", "E", "EM", "NO", "NA", "AO", "OS", "AS", "UM", "II", "III", "IV", "VI", "VII", "VIII", "IX", "XI", "XII", "XX", "XXI", "OK", "EUA", "USA", "UK",
                 "ABNT", "NBR", "ISO", "DOI", "URL", "PDF", "ISBN", "ISSN", "ET", "AL", "CM", "MM", "KG"}


def siglas(d):
    tb = _titulo(d, r"LISTA DE (ABREVIATURAS E SIGLAS|ABREVIATURAS|SIGLAS|SIGLAS E ABREVIATURAS)")
    if not tb:
        return []
    entradas = []
    for x in _ate_proximo_titulo(d, tb, 200):
        m = re.match(r"^\s*([A-Za-zÀ-ÿ0-9&./\-]{2,15})\s*(?:\t+|\s{2,}|\s[-–—:]\s)\s*(\S.*)$", x.bruto or x.texto)
        if m:
            entradas.append((m.group(1).strip(), x))
    if not entradas:
        return [Achado(CAT, "Lista de siglas: siglas usadas no texto", VERIFICAR, "não foi possível ler as linhas",
                       "Cada linha: SIGLA, tabulação e o significado por extenso.", "3.1.2.1.11", blocos=[tb.i])]
    corpo = " ".join(b.texto for b in d.paragrafos("textual") + d.paragrafos("pos"))
    # para achar siglas soltas: só parágrafos de texto corrido (não títulos, não referências em caixa alta)
    corrido = " ".join(b.texto for b in d.paragrafos("textual") if len(b.texto) > 60 and
                       sum(c.isupper() for c in b.texto) < 0.3 * sum(c.isalpha() for c in b.texto))
    nao_usadas = [(s, b.i) for s, b in entradas if not re.search(rf"(?<![\wÀ-ÿ]){re.escape(s)}(?![\wÀ-ÿ])", corpo)]
    listadas = {s.upper() for s, _ in entradas}
    usos = Counter(re.findall(r"(?<![\wÀ-ÿ])([A-Z][A-Z0-9]{1,7})(?![\wÀ-ÿ])", corrido))
    fora = sorted(s for s, n in usos.items() if n >= 3 and s not in listadas and s not in SIGLAS_COMUNS and
                  not re.fullmatch(r"[IVXLC]+", s) and not s.isdigit())
    ordem = [s for s in (s for s, _ in entradas)]
    chaves = [chave(s).upper() for s in ordem]
    fora_ordem = [ordem[i] for i in range(1, len(ordem)) if chaves[i] < chaves[i - 1]]
    out = [Achado(CAT, "Lista de siglas: siglas usadas no texto", VERIFICAR if nao_usadas else OK,
                  f"{len(nao_usadas)} de {len(entradas)} sigla(s) da lista não aparecem no texto" if nao_usadas
                  else f"{len(entradas)} siglas, todas usadas",
                  "Liste só as siglas que aparecem no texto." if nao_usadas else "", "3.1.2.1.11",
                  [s for s, _ in nao_usadas], blocos=[i for _, i in nao_usadas])]
    if fora:
        out.append(Achado(CAT, "Lista de siglas: siglas do texto que faltam na lista", VERIFICAR,
                          f"{len(fora)} sigla(s) usadas 3+ vezes e fora da lista",
                          "Inclua na lista as siglas usadas no texto (e escreva por extenso na 1ª menção).",
                          "3.1.2.1.11", fora[:60], blocos=[tb.i]))
    out.append(Achado(CAT, "Lista de siglas em ordem alfabética", AJUSTAR if fora_ordem else OK,
                      f"{len(fora_ordem)} fora de ordem" if fora_ordem else "em ordem",
                      "Ordene a lista alfabeticamente pela sigla." if fora_ordem else "", "3.1.2.1.11",
                      fora_ordem, blocos=[tb.i]))
    return out


def paginas_sumario(d):
    """Os números de página do sumário e das listas conferem com a paginação gravada pelo Word?"""
    pg = paginas(d)
    if not pg:
        return []
    out = []
    for rx_tit, nome, tipo in ((r"SUMARIO", "Sumário", "sec"),
                               (r"LISTA DE (FIGURAS|ILUSTRACOES)", "Lista de figuras", "ilu"),
                               (r"LISTA DE TABELAS", "Lista de tabelas", "tab")):
        tb = _titulo(d, rx_tit)
        if not tb:
            continue
        entradas = []
        for x in _ate_proximo_titulo(d, tb, 300):
            t = x.bruto or x.texto
            m = re.match(r"^(.*?)(?:\.{2,}|…+|\t)+\s*(\d{1,4})\s*$", t.strip())
            if m:
                entradas.append((m.group(1).strip(), int(m.group(2)), x))
        if len(entradas) < 2:
            continue
        alvo = []
        for rot, num_pag, _x in entradas:
            k = chave(rot).strip()
            achado = None
            mnum = re.match(r"^(\d{1,2}(?:\.\d{1,2}){0,6})\.?\s+(.*)$", k)
            for b in _pars(d):
                if b.i <= tb.i + len(entradas) or len(b.texto) > 200:
                    continue
                kb = chave(b.texto).strip()
                if tipo == "sec":
                    ok = (mnum and re.match(rf"^{re.escape(mnum.group(1))}\.?\s", kb)) or \
                         (not mnum and len(k) > 4 and kb == k)
                else:
                    mm = re.match(r"^(FIGURA|TABELA|QUADRO|GRAFICO)\s+(\d+)", k)
                    ok = mm and re.match(rf"^{mm.group(1)}\s+{mm.group(2)}\b", kb)
                if ok:
                    achado = b
                    break
            if achado is not None and achado.i in pg:
                alvo.append((rot, num_pag, pg[achado.i], achado))
        if len(alvo) < 2:
            continue
        desloc = Counter(p - f for _, p, f, _ in alvo).most_common(1)[0][0]
        errados = [(f"{rot[:60]}: diz p. {p}, está na p. {f + desloc}", b.i) for rot, p, f, b in alvo
                   if abs((p - f) - desloc) > 1]
        out.append(Achado(CAT, f"{nome}: números de página conferem", VERIFICAR if errados else OK,
                          f"{len(errados)} de {len(alvo)} com página diferente" if errados else
                          f"{len(alvo)} conferidos",
                          "O sumário/lista parece desatualizado: clique nele e pressione F9 (Atualizar o índice "
                          "inteiro). Conferência feita pela paginação que o Word gravou ao salvar o arquivo."
                          if errados else "", "3.1.2.1.12", _amostra([x for x, _ in errados]),
                          blocos=[i for _, i in errados[:300]]))
    return out


def epigrafe_e_errata(d):
    out = []
    pre = [b for b in _pars(d) if d.regiao(b.i) == "pre"]
    from .referencias import extrair_referencias
    refs = extrair_referencias(d)
    for b in pre:
        if du.recuo_esquerdo_cm(b.obj) < 4 and du.alinhamento(b.obj) != WD_ALIGN_PARAGRAPH.RIGHT:
            continue
        m = re.search(r"\(([A-ZÀ-Ý][\wÀ-ÿ' \-]+?)(?:\s+et al\.)?,\s*((?:19|20)\d{2})", b.texto)
        if not m or re.search(r"ORIENTADOR|APRESENTAD", chave(b.texto)):
            continue
        sob = _norm(m.group(1))[-1].lower() if _norm(m.group(1)) else ""
        ok = any(sob in r.tokens[:6] and m.group(2) in "".join(r.anos) for r in refs)
        out.append(Achado(CAT, "Epígrafe: obra citada nas referências", OK if ok else VERIFICAR,
                          f"({m.group(1)}, {m.group(2)})" + ("" if ok else " não está nas referências"),
                          "" if ok else "A obra da epígrafe deve constar nas referências.", "3.1.2.1.7",
                          blocos=[b.i]))
        break
    er = _titulo(d, r"ERRATA")
    if er:
        seg = d.blocos[er.i + 1:er.i + 15]
        tem_tab = any(x.tipo == "t" and re.search(r"ONDE SE LE", chave(" ".join(c.text for c in x.obj._cells)))
                      for x in seg)
        tem_txt = any(x.tipo == "p" and re.search(r"ONDE SE LE", chave(x.texto)) for x in seg)
        out.append(Achado(CAT, "Errata com o quadro Folha / Linha / Onde se lê / Leia-se",
                          OK if (tem_tab or tem_txt) else VERIFICAR,
                          "presente" if (tem_tab or tem_txt) else "a errata não traz as correções",
                          "" if (tem_tab or tem_txt) else
                          "Se houver errata: referência do trabalho e um quadro com Folha, Linha, Onde se lê e "
                          "Leia-se. Se não houver erros a corrigir, retire a página.", "3.1.2.1.3",
                          blocos=[er.i]))
    return out


# ------------------------------------------------------------------ catálogo e tabela-resumo
def catalogo():
    """[(nome, regex do título ou None, obrigatoriedade, itens de formatação, itens de conteúdo)]."""
    P = PF.atual()
    EL = P["elementos"]
    o = lambda flag, padrao=OBRIG: padrao if EL.get(flag, True) else OPC
    return [
        ("Capa", "@capa", o("capa"), ["Capa: instituição, local e ano"],
         ["Autor, título e ano iguais em todo o trabalho", "Textos de exemplo do template removidos"]),
        ("Folha de rosto", "@natureza", o("natureza"), ["Folha de rosto: nota de natureza do trabalho"],
         ["Tipo de trabalho confere com a norma escolhida", "Folha de rosto: orientador(a)",
          "Folha de rosto: versão original/corrigida"]),
        ("Ficha catalográfica", "@ficha", o("ficha"), [], ["Orientador igual na folha de rosto e na ficha"]),
        ("Errata", r"ERRATA", OPC, [], ["Errata com o quadro"]),
        ("Folha de aprovação", "@aprovacao", o("folha_aprovacao"), [], []),
        ("Dedicatória", "@dedicatoria", OPC, ["Dedicatória/epígrafe com título"], []),
        ("Agradecimentos", r"AGRADECIMENTOS?", OPC, [], ["Agradecimentos: frases obrigatórias de fomento"]),
        ("Declaração de uso de IA generativa", "@ia", OBRIG if EL.get("declaracao_ia") else OPC, [],
         ["Declaração de uso de IA generativa"]),
        ("Epígrafe", "@epigrafe", OPC, [], ["Epígrafe: obra citada nas referências"]),
        ("Resumo", r"RESUMO", OBRIG, ["Resumo: parágrafo único", "Resumo: formato de 'Palavras-chave'"],
         ["Resumo: referência do trabalho acima do texto", "Mesmo número de palavras-chave e keywords"]),
        ("Abstract", r"ABSTRACT", OBRIG if EL.get("abstract", True) else OPC,
         ["Abstract: parágrafo único", "Abstract: formato de 'Keywords'"],
         ["Abstract: referência do trabalho acima do texto"]),
        ("Lista de ilustrações", r"LISTA DE (FIGURAS|ILUSTRACOES|GRAFICOS|QUADROS)", OPC, [],
         ["Lista de ilustrações confere", "Lista de figuras confere", "Lista de figuras: números de página",
          "Lista de quadros confere", "Lista de gráficos confere"]),
        ("Lista de tabelas", r"LISTA DE TABELAS", OPC, [],
         ["Lista de tabelas confere", "Lista de tabelas: números de página"]),
        ("Lista de abreviaturas e siglas", r"LISTA DE (ABREVIATURAS E SIGLAS|ABREVIATURAS|SIGLAS)", OPC,
         ["Lista de siglas em ordem alfabética"],
         ["Lista de siglas: siglas usadas no texto", "Lista de siglas: siglas do texto que faltam na lista"]),
        ("Lista de símbolos", r"LISTA DE SIMBOLOS", OPC, [], []),
        ("Sumário", r"SUMARIO", OBRIG, [],
         ["Sumário com as mesmas palavras", "Todos os títulos aparecem no sumário",
          "Sumário sem títulos que não existem", "Sumário: números de página"]),
        ("Introdução", "@introducao", OBRIG, ["Títulos das seções", "Seção primária em nova página"], []),
        ("Conclusão / Considerações finais", "@conclusao", OBRIG, [], []),
        ("Referências", "@referencias", OBRIG,
         ["Espaço simples nas referências", "Referências alinhadas", "Título em destaque", "Ordem alfabética"],
         ["Toda citação está nas referências", "Toda referência foi citada", "Referência repetida"]),
        ("Glossário", r"GLOSSARIO", OPC, [], []),
        ("Apêndice", r"APENDICES?(\s+[A-Z]{1,2}\b.*)?", OPC, [], ["Figuras, tabelas, seções e apêndices citados"]),
        ("Anexo", r"ANEXOS?(\s+[A-Z]{1,2}\b.*)?", OPC, [], []),
        ("Índice", r"INDICE", OPC, [], []),
    ]


def _localizar(d, rx):
    pars = _pars(d)
    pre = [b for b in pars if d.regiao(b.i) == "pre"] or pars[:150]
    if rx == "@capa":
        return next((b for b in pre if "UNIVERSIDADE DE SAO PAULO" in chave(b.texto)), None)
    if rx == "@natureza":
        from .trabalho import detectar
        return detectar(d)[2]
    if rx == "@ficha":
        return next((b for b in pars if re.search(r"CATALOGACAO NA PUBLICACAO|FICHA CATALOGRAFICA|\bCDD\b|\bCDU\b",
                                                   chave(b.texto))), None)
    if rx == "@aprovacao":
        return next((b for b in pars if re.search(r"FOLHA DE (AVALIACAO|APROVACAO)|APROVAD[OA] EM|BANCA EXAMINADORA",
                                                   chave(b.texto))), None)
    if rx == "@ia":
        return next((b for b in pars if re.search(r"USO DE INTELIGENCIA ARTIFICIAL|DECLARACAO DE USO DE IA",
                                                   chave(b.texto))), None)
    if rx == "@dedicatoria":
        return next((b for b in pre if du.recuo_esquerdo_cm(b.obj) >= 4 and
                     re.search(r"\b(DEDICO|DEDICAD[OA]|DEDICATORIA|A MINHA|AOS MEUS|AS MINHAS|AO MEU)\b", chave(b.texto))),
                    None)
    if rx == "@epigrafe":
        return next((b for b in pre if (du.recuo_esquerdo_cm(b.obj) >= 4 or
                                        du.alinhamento(b.obj) == WD_ALIGN_PARAGRAPH.RIGHT) and
                     re.search(r"\([A-ZÀ-Ý][^()]*,\s*(19|20)\d{2}", b.texto) and
                     not re.search(r"APRESENTAD|ORIENTADOR", chave(b.texto))), None)
    if rx == "@introducao":
        return d.blocos[d.i_intro] if d.i_intro is not None else None
    if rx == "@conclusao":
        return next((b for b in pars if d.regiao(b.i) == "textual" and len(b.texto) < 80 and
                     re.match(r"^(\d+\.?\s*[-–]?\s*)?(CONCLUS(AO|OES)|CONSIDERACOES FINAIS)\b", chave(b.texto))), None)
    if rx == "@referencias":
        return d.blocos[d.i_refs] if d.i_refs is not None else None
    return _titulo(d, rx)


def _formato_titulo(d, b, nome):
    """Título de elemento: sem número, centralizado, em maiúsculas."""
    k = chave(b.texto).strip()
    probs = []
    if nome in ("Capa", "Folha de rosto", "Ficha catalográfica", "Dedicatória", "Epígrafe", "Introdução",
                "Conclusão / Considerações finais", "Folha de aprovação"):
        return probs
    if re.match(r"^\d", k) and nome not in ("Apêndice", "Anexo"):
        probs.append("título numerado")
    if du.alinhamento(b.obj) != WD_ALIGN_PARAGRAPH.CENTER:
        probs.append("título não centralizado")
    letras = re.sub(r"[^A-Za-zÀ-ÿ]", "", b.texto)
    if letras and letras != letras.upper():
        probs.append("título fora de maiúsculas")
    return probs


def resumo_elementos(d, achados):
    """Linhas da tabela 'Elementos do trabalho' (aparece / formatação / conteúdo)."""
    linhas, extras_ = [], []
    por_item = {}
    for a in achados:
        por_item.setdefault(a.item, []).append(a)

    def busca(prefixos):
        res = []
        for p in prefixos:
            for item, lst in por_item.items():
                if item.startswith(p):
                    res += lst
        return res

    pos_intro = d.i_intro
    for nome, rx, obrig, it_fmt, it_cont in catalogo():
        b = _localizar(d, rx)
        if b is None:
            aparece = "❌ falta" if obrig == OBRIG else "— não tem"
            linhas.append({"Elemento": nome, "Exigência": obrig, "Aparece": aparece, "Formatação": "—",
                           "Conteúdo confere": "—"})
            continue
        onde = ""
        pre_el = nome not in ("Introdução", "Conclusão / Considerações finais", "Referências", "Glossário",
                              "Apêndice", "Anexo", "Índice")
        if pre_el and pos_intro is not None and b.i > pos_intro:
            onde = " (fora da parte pré-textual)"
            extras_.append(Achado(CAT, f"{nome}: posição no trabalho", VERIFICAR,
                                  "aparece depois da introdução",
                                  "Elementos pré-textuais vêm antes da introdução (não precisam estar na mesma "
                                  "página do modelo, mas sim nessa parte do trabalho).", "3.1", blocos=[b.i]))
        fmt = _formato_titulo(d, b, nome)
        ja = {t for a in achados if a.item.startswith("Títulos sem indicativo") and a.status != OK for t in a.trechos}
        novos = [f for f in fmt if not (b.texto in ja and f in ("título numerado", "título não centralizado"))]
        if novos:
            extras_.append(Achado(CAT, f"{nome}: título formatado", AJUSTAR, "; ".join(novos),
                                  "Títulos de elementos sem indicativo numérico: centralizados, em maiúsculas e "
                                  "com o mesmo destaque das seções primárias.", "2.3", [b.texto[:80]], blocos=[b.i]))
        af = [a for a in busca(it_fmt) if a.status != OK]
        ac = [a for a in busca(it_cont) if a.status != OK]
        sinal = lambda lst, extra=(): ("✅ ok" if not lst and not extra else
                                       ("⚠️ " if all(x.status == VERIFICAR for x in lst) and not extra else "❌ ")
                                       + "; ".join([*extra, *(x.item for x in lst)])[:160])
        tem_cont = bool(it_cont)
        linhas.append({"Elemento": nome, "Exigência": obrig, "Aparece": "✅ sim" + onde,
                       "Formatação": sinal(af, fmt) if (it_fmt or fmt or nome not in ("Capa",)) else "—",
                       "Conteúdo confere": sinal(ac) if tem_cont else "—"})
    return linhas, extras_


def verificar(d):
    out = []
    for f in (consistencia, mencoes, siglas, paginas_sumario, epigrafe_e_errata):
        try:
            out += f(d)
        except Exception as e:
            out.append(Achado("Sistema", f"elementos.{f.__name__}", VERIFICAR, f"verificação não concluída: {e}"))
    return out
