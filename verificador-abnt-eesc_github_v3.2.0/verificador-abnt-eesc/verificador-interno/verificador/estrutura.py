"""Verificações de apresentação gráfica e estrutura (Diretrizes USP, seções 2 e 3)."""
import re
from collections import Counter, defaultdict
from types import SimpleNamespace

from docx.enum.text import WD_ALIGN_PARAGRAPH

from . import docx_util as du
from . import perfil as PF
from .cache import qn
from .documento import RE_APENDICE, TITULOS_SEM_NUMERO, chave
from .modelo import AJUSTAR, OK, VERIFICAR, Achado, T, Trecho, num

CAT_FORM = "Formatação"
CAT_EST = "Estrutura"
CAT_SEC = "Seções e títulos"
CAT_ILU = "Ilustrações"


def _amostra(textos, n=300, tam=120):
    """Até n exemplos, cortados em 'tam' caracteres (mantém o bloco de origem de cada um)."""
    return [Trecho(t[:tam] + ("…" if len(t) > tam else ""), getattr(t, "i", None), getattr(t, "m", None))
            for t in textos[:n]]


# ------------------------------------------------------------------ papel e margens
def _faixas(indices):
    """[1,2,3,5] -> '1–3, 5'"""
    out, ini, ant = [], None, None
    for i in indices + [None]:
        if ini is None:
            ini = ant = i
        elif i is not None and i == ant + 1:
            ant = i
        else:
            out.append(str(ini) if ini == ant else f"{ini}–{ant}")
            ini = ant = i
    return ", ".join(out)


def papel_e_margens(d):
    out = []
    espelhada = du.margens_espelhadas(d.doc)
    medianiz_no_topo = d.doc.settings.element.find(qn("w:gutterAtTop")) is not None
    papel_errado = []
    grupos = {}  # (esq, dir, sup, inf, medianiz) -> [seções]
    for k, s in enumerate(d.doc.sections, 1):
        w, h = du.cm(s.page_width), du.cm(s.page_height)
        if w and h and not (abs(w - 21) <= 0.1 and abs(h - 29.7) <= 0.1):
            papel_errado.append(k)
        med = du.cm(s.gutter) or 0.0
        esq, dir_ = du.cm(s.left_margin) or 0, du.cm(s.right_margin) or 0
        sup, inf = abs(du.cm(s.top_margin) or 0), abs(du.cm(s.bottom_margin) or 0)
        # a medianiz (margem de encadernação) soma à margem esquerda/interna (ou à superior)
        if medianiz_no_topo:
            sup += med
        else:
            esq += med
        grupos.setdefault((round(esq, 2), round(dir_, 2), round(sup, 2), round(inf, 2), med), []).append(k)

    total = len(d.doc.sections)
    out.append(Achado(CAT_FORM, "Papel A4", AJUSTAR if papel_errado else OK,
                      f"tamanho diferente de A4 nas seções {_faixas(papel_errado)}" if papel_errado
                      else "21 x 29,7 cm", "Layout > Tamanho > A4." if papel_errado else "", "2.2"))

    M = PF.atual()["margens"]
    esperado = (M["esquerda"], M["direita"], M["superior"], M["inferior"])
    rot = ("Interna", "Externa") if espelhada else ("Esquerda", "Direita")
    linhas, erradas = [], 0
    for (esq, dir_, sup, inf, med), secs in grupos.items():
        valores = (esq, dir_, sup, inf)
        ok = all(abs(v - e) <= 0.1 for v, e in zip(valores, esperado))
        if not ok:
            erradas += len(secs)

        def cel(v, e):
            return num(v) + ("" if abs(v - e) <= 0.1 else "  ✗")
        linhas.append({"Seções do Word": _faixas(secs),
                       rot[0]: cel(esq, esperado[0]) + (f" (inclui medianiz {num(med)})" if med else ""),
                       rot[1]: cel(dir_, esperado[1]), "Superior": cel(sup, esperado[2]),
                       "Inferior": cel(inf, esperado[3])})
    tabela = [{"Seções do Word": "Esperado", rot[0]: num(esperado[0]), rot[1]: num(esperado[1]),
               "Superior": num(esperado[2]), "Inferior": num(esperado[3])}] + linhas

    if erradas:
        det = (f"fora do padrão em {erradas} de {total} seção(ões) do Word" if total > 1
               else "fora do padrão")
    else:
        det = PF.margens_txt() + (" (margens espelhadas)" if espelhada else "")
    e3, d2, s3, i2 = (PF.fmt(x) for x in esperado)
    out.append(Achado(CAT_FORM, "Margens (cm)", AJUSTAR if erradas else OK, det,
                      (f"Layout > Margens: esquerda {e3} cm, direita {d2} cm, superior {s3} cm, inferior {i2} cm. "
                       f"Se usar medianiz, a soma 'esquerda + medianiz' deve dar {e3} cm. "
                       f"Para impressão frente e verso, use margens espelhadas (interna {e3} cm, externa {d2} cm). "
                       "Aplique a 'Todo o documento' para corrigir todas as seções de uma vez.") if erradas else "",
                      "2.2", tabela=tabela if (erradas or total > 1) else []))
    return out


# ------------------------------------------------------------------ texto
def _eh_legenda_ou_fonte(t):
    return re.match(r"^(Figura|Quadro|Gr[áa]fico|Tabela|Fotografia|Mapa|Esquema|Fluxograma|Imagem|"
                    r"Desenho|Organograma|Planta|Fonte|Nota|Legenda)\b", t) is not None


RE_LISTA = re.compile(r"^(\w+\s+\d+\s*[:.)\-–]|[a-zA-Z]\)|\d+[.)]\s|[\-–•●▪◦]\s?)")
RE_CHAMADA = re.compile(r"\([^()]{2,80}?,?\s*(?:1[5-9]\d{2}|20\d{2}|\[?s\.\s?d\.\]?)[a-z]?(?:,\s*p\.[^)]*)?\)")


def citacao_longa(p, t, doc):
    """Classifica um parágrafo recuado: True = citação longa; None = recuado, mas sem sinais (conferir);
    False = não é citação longa."""
    recuo = du.recuo_esquerdo_cm(p)
    CL = PF.atual()["citacao_longa"]
    if recuo < max(1.5, CL["recuo"] - 0.5):
        return False
    t = t.strip()
    if RE_LISTA.match(t) or t.endswith(":"):
        return False                      # lista/etapas recuadas, não é citação
    tam = du.tamanho_predominante(p, doc)
    sinais = 0
    if RE_CHAMADA.search(t) or re.search(r"\b(19|20)\d{2}\b[^.]{0,20}p\.\s?\d", t):
        sinais += 1
    if tam and tam <= CL["fonte_max"]:
        sinais += 1
    if t[:1] in "\"“":
        sinais += 1
    if (du.entrelinhas(p, doc) or 1.5) <= 1.15:
        sinais += 1
    return True if sinais else None


def texto_principal(d):
    out = []
    P = PF.atual()
    CL, FT, EL = P["citacao_longa"], P["fonte_texto"], P["entrelinhas"]
    LIM_REC = max(1.5, CL["recuo"] - 0.5)
    fonte_err, esp_err, longas_err, longas_inline, recuados = [], [], [], [], []
    n_corpo = 0
    for b in d.paragrafos("textual"):
        p, t = b.obj, b.texto
        if not t or len(t) < 60 or du.nivel_titulo_estilo(p) or _eh_legenda_ou_fonte(t):
            continue
        recuo = du.recuo_esquerdo_cm(p)
        tam = du.tamanho_predominante(p, d.doc)
        esp = du.entrelinhas(p, d.doc)
        tipo_cit = citacao_longa(p, t, d.doc) if recuo >= LIM_REC else False
        if tipo_cit is None:
            recuados.append(T(t, b))
            continue
        if tipo_cit:  # citação longa
            probs = []
            if tam and tam > CL["fonte_max"]:
                probs.append(f"fonte {tam:g}")
            if esp and esp > CL["entrelinhas"] + 0.15:
                probs.append(f"entrelinhas {esp:g}")
            if abs(recuo - CL["recuo"]) > 0.3:
                probs.append(f"recuo {recuo:g} cm")
            if t[:1] in "\"“":
                probs.append("está entre aspas")
            if probs:
                longas_err.append(T(f"[{', '.join(probs)}] {t}", b))
            continue
        if re.match(r"^[a-z]\)\s", t) or re.match(r"^[-–—•]\s", t) or recuo >= LIM_REC:
            continue  # alíneas são verificadas à parte; listas recuadas não são texto corrido
        n_corpo += 1
        if tam and abs(tam - FT) > 0.1:
            fonte_err.append(T(f"[fonte {tam:g}] {t}", b))
        if esp and abs(esp - EL) > 0.1:
            esp_err.append(T(f"[entrelinhas {esp:g}] {t}", b))
        for m in re.finditer(r"[\"“]([^\"”]{300,})[\"”]", t):
            longas_inline.append(T(m.group(1), b, m.group(0)[:60]))

    out.append(Achado(CAT_FORM, f"Fonte do texto (tamanho {PF.fmt(FT)})",
                      AJUSTAR if fonte_err else OK,
                      f"{len(fonte_err)} de {n_corpo} parágrafos fora do padrão" if fonte_err
                      else f"{n_corpo} parágrafos conferidos",
                      f"Texto em fonte {PF.fmt(FT)}; tamanho menor só para citações longas, notas, legendas e fontes.",
                      "2.2", _amostra(fonte_err)))
    out.append(Achado(CAT_FORM, f"Entrelinhas {PF.fmt(EL)} no texto",
                      AJUSTAR if esp_err else OK,
                      f"{len(esp_err)} de {n_corpo} parágrafos fora do padrão" if esp_err
                      else f"{n_corpo} parágrafos conferidos",
                      f"Parágrafo > Espaçamento entre linhas: {PF.fmt(EL)}.", "2.2.1", _amostra(esp_err)))
    out.append(Achado(CAT_FORM, "Citações com mais de 3 linhas",
                      AJUSTAR if (longas_err or longas_inline) else OK,
                      (f"{len(longas_err)} citação(ões) recuada(s) com problema; " if longas_err else "")
                      + (f"{len(longas_inline)} citação(ões) longa(s) entre aspas no meio do parágrafo"
                         if longas_inline else "") or "sem problemas encontrados",
                      f"Mais de 3 linhas: parágrafo próprio, recuo de {PF.fmt(CL['recuo'])} cm, fonte até "
                      f"{PF.fmt(CL['fonte_max'])}, entrelinhas {PF.fmt(CL['entrelinhas'])}, sem aspas.",
                      "4.1", _amostra(longas_err + longas_inline)))
    if recuados:
        out.append(Achado(CAT_FORM, "Parágrafos recuados sem sinal de citação", VERIFICAR,
                          f"{len(recuados)} parágrafo(s) com recuo ≥ {PF.fmt(LIM_REC)} cm",
                          f"Se for citação direta longa: recuo {PF.fmt(CL['recuo'])} cm, fonte até "
                          f"{PF.fmt(CL['fonte_max'])}, espaço simples e a chamada "
                          "(Autor, ano, p.). Se for texto comum, retire o recuo.", "4.1", _amostra(recuados)))
    return out


# ------------------------------------------------------------------ elementos obrigatórios
def elementos_obrigatorios(d):
    out = []
    pos = {}
    EL_ = PF.atual()["elementos"]
    for nome, padrao in [("Resumo", r"RESUMO"), ("Abstract", r"ABSTRACT"), ("Sumário", r"SUMARIO"),
                         ("Introdução", r"(1\.?\s*[-–]?\s*)?INTRODUCAO"),
                         ("Referências", r"(\d+\.?\s*)?(REFERENCIAS( BIBLIOGRAFICAS)?|BIBLIOGRAFIA)")]:
        b = d.achar_titulo(padrao)
        if nome == "Abstract" and not EL_["abstract"]:
            pos[nome] = b.i if b else None
            continue
        pos[nome] = b.i if b else None
        out.append(Achado(CAT_EST, f"{nome}", OK if b else AJUSTAR,
                          "encontrado" if b else "título não encontrado",
                          "" if b else f"Incluir o elemento obrigatório '{nome.upper()}' com o título "
                                       "isolado numa linha.", "3.1"))

    todo = "\n".join(b.texto for b in d.blocos if b.tipo == "p")
    tudo_k = chave(todo)
    checagens = [
        ("Folha de rosto: versão original/corrigida", r"VERSAO (ORIGINAL|CORRIGIDA)|ORIGINAL VERSION|REVISED VERSION",
         "Indicar 'Versão original' ou 'Versão corrigida' na folha de rosto.", "3.1.2.1.1"),
        ("Folha de rosto: orientador(a)", r"ORIENTADOR(A|\(A\))?:|ORIENTADORES:|ADVISORS?:|SUPERVISOR:",
         "Indicar 'Orientador(a): Prof(a). Dr(a). Nome'.", "3.1.2.1.1"),
        ("Ficha catalográfica", r"CATALOGACAO NA PUBLICACAO|FICHA CATALOGRAFICA|\bCDD\b|\bCDU\b",
         "A ficha é elaborada pela Biblioteca: solicitar ao Serviço de Biblioteca da EESC.", "3.1.2.1.2"),
        ("Folha de aprovação", r"APROVAD[OA] EM|BANCA EXAMINADORA|JULGAMENTO|FOLHA DE (AVALIACAO|APROVACAO)",
         "Incluir folha de aprovação com 'Aprovado em:' e campos da banca em branco.", "3.1.2.1.4"),
        ("Palavras-chave", ROTULO_PALAVRAS, "Logo abaixo do resumo: 'Palavras-chave:'.", "3.1.2.1.8"),
        ("Keywords", ROTULO_KEYWORDS, "Logo abaixo do abstract: 'Keywords:'.", "3.1.2.1.9"),
    ]
    desligar = {"Folha de rosto: versão original/corrigida": "versao", "Folha de rosto: orientador(a)": "orientador",
                "Ficha catalográfica": "ficha", "Folha de aprovação": "folha_aprovacao", "Keywords": "keywords"}
    for nome, padrao, corr, sec in checagens:
        if nome in desligar and not EL_[desligar[nome]]:
            continue
        if nome in ("Palavras-chave", "Keywords"):  # rótulo no começo de um parágrafo (ou de uma linha)
            ok = any(re.match(r"\s*" + padrao, chave(ln)) for b in d.blocos if b.tipo == "p" and b.texto
                     for ln in (b.bruto or b.texto).split("\n"))
        else:
            ok = re.search(padrao, tudo_k) is not None
        out.append(Achado(CAT_EST, nome, OK if ok else VERIFICAR,
                          "encontrado" if ok else "não identificado no texto", "" if ok else corr, sec))

    ordem = [(n, pos[n]) for n in ("Resumo", "Abstract", "Sumário", "Introdução", "Referências") if pos[n] is not None]
    fora = [f"{a} aparece depois de {b}" for (a, ia), (b, ib) in zip(ordem[1:], ordem[:-1]) if ia < ib]
    if ordem:
        out.append(Achado(CAT_EST, "Ordem dos elementos", VERIFICAR if fora else OK,
                          "; ".join(fora) or "Resumo → Abstract → Sumário → Introdução → Referências",
                          "Seguir a ordem da NBR 14724 (Figura 1 do manual).", "3.1"))
    if re.search(r"^\s*(DEDICATORIA|EPIGRAFE)\s*$", tudo_k, re.M):
        out.append(Achado(CAT_EST, "Dedicatória/epígrafe com título", AJUSTAR,
                          "há título 'DEDICATÓRIA' ou 'EPÍGRAFE'",
                          "Dedicatória e epígrafe não levam título.", "3.1.2.1.5"))
    return out


# ------------------------------------------------------------------ resumo
def _partes_rotulo(texto):
    """'Palavras-chave: a. b.' -> ('Palavras-chave', ':', 'a. b.'). Tolera variações do rótulo."""
    m = re.match(r"^\s*((?:palavras?[\s-]*chaves?|key[\s-]*words?))\b\s*([:.\-–—]?)\s*(.*)$", texto, re.I | re.S)
    if not m:
        return texto.strip(), "", ""
    return m.group(1).strip(), m.group(2), m.group(3).strip()


def _resumo(d, titulo, rotulo, nome):
    out = []
    b = d.achar_titulo(titulo)
    if not b:
        return out
    pars, kw, termos_b, kw_texto, colado = [], None, None, "", False
    for x in d.blocos[b.i + 1:]:
        if x.tipo != "p" or not x.texto:
            continue
        if re.match(rotulo, chave(x.texto)):
            kw, kw_texto = x, x.texto
            if not _partes_rotulo(x.texto)[2]:  # rótulo sozinho na linha: termos no parágrafo seguinte
                termos_b = next((y for y in d.blocos[x.i + 1:x.i + 4] if y.tipo == "p" and y.texto), None)
            break
        linhas = (x.bruto or x.texto).split("\n")
        k = next((k for k, ln in enumerate(linhas) if k and re.match(r"\s*" + rotulo, chave(ln))), None)
        if k is not None:  # rótulo depois de uma quebra de linha, dentro do parágrafo do resumo
            antes = "\n".join(linhas[:k]).strip()
            if antes:
                pars.append(SimpleNamespace(i=x.i, texto=antes, obj=x.obj))
            kw, kw_texto, colado = x, "\n".join(linhas[k:]).strip(), True
            break
        if du.nivel_titulo_estilo(x.obj) or chave(x.texto) in TITULOS_SEM_NUMERO:
            break
        pars.append(x)
    if not pars:
        return out
    ref = pars[0] if re.search(r"\b(19|20)\d{2}\b", pars[0].texto) and re.search(
        r"(Disserta|Tese|Trabalho de Conclus|Thesis|Dissertation|Monografia)", pars[0].texto) else None
    corpo = pars[1:] if ref else pars
    out.append(Achado(CAT_EST, f"{nome}: referência do trabalho antes do texto", OK if ref else VERIFICAR,
                      "encontrada" if ref else "não identificada",
                      "" if ref else "Antes do texto do resumo, colocar a referência do próprio trabalho "
                                     "(SOBRENOME, A. B. Título. Ano. Dissertação (Mestrado em ...) – Escola de "
                                     "Engenharia de São Carlos, Universidade de São Paulo, São Carlos, ano.).",
                      "3.1.2.1.8", blocos=[] if ref else [pars[0].i]))
    palavras = sum(len(x.texto.split()) for x in corpo)
    R = PF.atual()["resumo"]
    status = OK if R["min"] <= palavras <= R["max"] and len(corpo) == 1 else AJUSTAR
    det = f"{palavras} palavras em {len(corpo)} parágrafo(s)"
    out.append(Achado(CAT_EST, f"{nome}: parágrafo único, {R['min']} a {R['max']} palavras", status, det,
                      "" if status == OK else f"Resumo em parágrafo único, com {R['min']} a {R['max']} palavras.",
                      "3.1.2.1.8", blocos=[x.i for x in corpo] or [b.i]))
    if kw:
        rot, sinal, t = _partes_rotulo(kw_texto)
        t = t or (termos_b.texto.strip() if termos_b else "")
        probs = []
        if ";" in t or (t.count(",") >= 2 and t.count(".") <= 1):
            probs.append("termos separados por vírgula ou ponto e vírgula")
        if t and not t.endswith("."):
            probs.append("não termina com ponto")
        if sinal != ":":
            probs.append("rótulo sem dois-pontos")
        if colado:
            probs.append("no mesmo parágrafo do resumo (separado só por quebra de linha)")
        if rot not in ("Palavras-chave", "Keywords"):
            probs.append(f"rótulo '{rot}' (use '{'Palavras-chave' if nome == 'Resumo' else 'Keywords'}')")
        rotulo_certo = "Palavras-chave" if nome == "Resumo" else "Keywords"
        out.append(Achado(CAT_EST, f"{nome}: formato de '{rotulo_certo}'",
                          AJUSTAR if probs else OK, "; ".join(probs) or t[:100],
                          "Termos separados entre si por ponto e finalizados por ponto. Ex.: "
                          f"{rotulo_certo}: Engenharia. Materiais. Ensaios.", "3.1.2.1.8",
                          [T(kw_texto, kw, rot if rot not in ("Palavras-chave", "Keywords") else None)],
                          blocos=[kw.i] + ([termos_b.i] if termos_b else [])))
    return out


def resumos(d):
    return (_resumo(d, r"RESUMO", ROTULO_PALAVRAS, "Resumo")
            + (_resumo(d, r"ABSTRACT", ROTULO_KEYWORDS, "Abstract") if PF.atual()["elementos"]["abstract"] else []))


# ------------------------------------------------------------------ seções e títulos
# número da seção + separador (espaço, ou pontuação) + título que começa com letra
# rótulo das palavras-chave, tolerante ao que se vê nos trabalhos: "Palavras-chave:", "Palavras chave -",
# "PALAVRAS-CHAVES.", "Palavra-chave" sozinho na linha (termos no parágrafo seguinte); idem "Keywords"
ROTULO_PALAVRAS = r"PALAVRAS?[\s-]*CHAVES?\b\s*(?:[:.\-–—]|$)"
ROTULO_KEYWORDS = r"KEY[\s-]*WORDS?\b\s*(?:[:.\-–—]|$)"
RE_SECAO = re.compile(r"^(\d{1,2}(?:\.\d{1,2}){0,6})(\.?)(\s*[-–—:]\s*|\s+)([A-Za-zÀ-ÿ].*)$")
RE_SUMARIO = re.compile(r"(\.{4,}|…{2,}|\t)\s*\d+\s*$")  # linha de sumário: pontilhado/tab + nº da página


def _candidato_titulo(b):
    p, t = b.obj, b.texto
    if not t or len(t) > 150 or "\t" in p.text or RE_SUMARIO.search(t):
        return None
    nivel = du.nivel_titulo_estilo(p)
    m = RE_SECAO.match(t)
    if m and (nivel or du.todo_negrito(p) or t.upper() == t) and not t.endswith((".", ";", ",")):
        return m
    if nivel and not m:
        return "auto"
    return None


def secoes(d):
    out = []
    titulos = []
    auto, auto_b = 0, []
    for b in d.paragrafos():
        if d.regiao(b.i) == "pre":
            continue
        m = _candidato_titulo(b)
        if m == "auto":
            if chave(b.texto) in TITULOS_SEM_NUMERO or RE_APENDICE.match(chave(b.texto)):
                continue
            if du.tem_numeracao_automatica(b.obj):
                auto += 1
                auto_b.append(T(b.texto, b))
            continue
        if m:
            titulos.append((b, m))

    if not titulos and not auto:
        out.append(Achado(CAT_SEC, "Numeração progressiva", VERIFICAR,
                          "nenhum título numerado identificado",
                          "Seções numeradas em algarismos arábicos (1, 1.1, 1.1.1...).", "2.3"))
        return out
    if auto:
        out.append(Achado(CAT_SEC, "Numeração automática do Word", VERIFICAR,
                          f"{auto} título(s) numerado(s) automaticamente: conferir no arquivo",
                          "Conferir visualmente: indicativo + espaço + título, sem ponto ou hífen.", "2.3",
                          _amostra(auto_b)))

    pontuacao, profundos, sequencia, sem_texto, sem_pagina = [], [], [], [], []
    estilos = defaultdict(Counter)
    anterior = None
    for b, m in titulos:
        num, ponto, sep, titulo = m.groups()
        partes = [int(x) for x in num.split(".")]
        nivel = len(partes)
        if ponto or sep.strip():
            pontuacao.append(T(b.texto, b, (num + ponto + sep).strip() or None))
        if nivel > 5:
            profundos.append(T(b.texto, b))
        if anterior:
            prox = []
            for n in range(1, len(anterior) + 1):
                base = anterior[:n]
                base = base[:-1] + [base[-1] + 1]
                prox.append(base)
            prox.append(anterior + [1])
            ok_seq = partes in prox
            if not ok_seq:
                sequencia.append(T(f"{'.'.join(map(str, anterior))} → {num}", b, num))
        anterior = partes
        p = b.obj
        estilos[nivel][(du.todo_negrito(p), titulo.upper() == titulo)] += 1
        if nivel == 1 and PF.atual()["elementos"]["secao_nova_pagina"] and \
                not du.comeca_em_nova_pagina(p, d.anterior_paragrafo(b.i)):
            sem_pagina.append(T(b.texto, b))
        prox_bloco = next((x for x in d.blocos[b.i + 1:] if x.tipo == "t" or x.texto), None)
        if prox_bloco is not None and prox_bloco.tipo == "p" and _candidato_titulo(prox_bloco) not in (None,):
            sem_texto.append(T(b.texto, b))

    def add(item, lista, corr, sec="2.3"):
        out.append(Achado(CAT_SEC, item, AJUSTAR if lista else OK,
                          f"{len(lista)} ocorrência(s)" if lista else "sem problemas", corr if lista else "",
                          sec, _amostra(lista)))

    add("Indicativo numérico sem ponto, hífen ou dois-pontos", pontuacao,
        "Formato: '1 INTRODUÇÃO', '2.1 Materiais' — número, um espaço e o título.")
    add("Numeração até a seção quinária", profundos, "Limitar a 5 níveis; abaixo disso usar alíneas.")
    add("Sequência da numeração", sequencia, "Corrigir saltos ou repetições na numeração das seções.")
    add("Seções primárias em nova página", sem_pagina,
        "Inserir quebra de página antes de cada seção primária (ou 'Quebra de página antes' no estilo).")
    add("Toda seção com texto", sem_texto,
        "Toda seção deve ter texto próprio antes da subseção seguinte.")
    misturados = [f"nível {n}: {len(c)} formatos diferentes" for n, c in estilos.items() if len(c) > 1]
    out.append(Achado(CAT_SEC, "Destaque tipográfico uniforme por nível", VERIFICAR if misturados else OK,
                      "; ".join(misturados) or "uniforme",
                      "Cada nível de seção com o mesmo destaque (maiúsculas/negrito), repetido no sumário."
                      if misturados else "", "2.3"))
    return out


def titulos_sem_indicativo(d):
    errados_num, nao_centralizados, ref_bib, centrado_a_mao = [], [], [], []
    for b in d.paragrafos():
        if not b.texto or "\t" in b.bruto or len(b.texto) > 80:
            continue
        k = chave(b.texto)
        base = re.sub(r"^\d+(\.\d+)*\.?\s*[-–]?\s*", "", k)
        alvo = base in TITULOS_SEM_NUMERO or RE_APENDICE.match(base)
        if not alvo or (base != k and not re.match(r"^\d", k)):
            continue
        if base in ("RESUMO", "ABSTRACT") and d.regiao(b.i) != "pre":
            continue
        if base != k:
            errados_num.append(T(b.texto, b))
        if du.alinhamento(b.obj) != WD_ALIGN_PARAGRAPH.CENTER:
            # recuo grande ou espaços no começo: o aluno "centralizou à mão" (comum em arquivos convertidos)
            if du.recuo_esquerdo_cm(b.obj) >= 2.5 or re.match(r"^[ \t\u00a0]{3,}", b.bruto or ""):
                centrado_a_mao.append(T(b.texto, b))
            else:
                nao_centralizados.append(T(b.texto, b))
        if base in ("REFERENCIAS", "REFERENCIAS BIBLIOGRAFICAS", "BIBLIOGRAFIA") and \
                base != chave(PF.atual()["referencias"]["titulo"]):
            ref_bib.append(T(b.texto, b))
    out = [
        Achado(CAT_SEC, "Títulos sem indicativo numérico", AJUSTAR if errados_num else OK,
               f"{len(errados_num)} numerado(s)" if errados_num else "sem numeração",
               "Resumo, sumário, listas, referências, apêndices e anexos não são numerados.", "2.3",
               _amostra(errados_num)),
        Achado(CAT_SEC, "Títulos sem indicativo centralizados", AJUSTAR if nao_centralizados else OK,
               f"{len(nao_centralizados)} não centralizado(s)" if nao_centralizados else "centralizados",
               "Centralizar e usar o mesmo destaque das seções primárias.", "2.3",
               _amostra(nao_centralizados)),
    ]
    if centrado_a_mao:
        out.append(Achado(CAT_SEC, "Títulos centralizados com recuo ou espaços", VERIFICAR,
                          f"{len(centrado_a_mao)} título(s) parecem centralizados, mas com recuo ou espaços",
                          "Usar o alinhamento 'Centralizado' (Ctrl+E) e tirar o recuo/espaços do início: assim o "
                          "título continua no centro se a fonte ou a margem mudar.", "2.3", _amostra(centrado_a_mao)))
    if ref_bib:
        out.append(Achado(CAT_SEC, "Título da lista de referências", AJUSTAR, ref_bib[0],
                          f"Usar o título '{PF.atual()['referencias']['titulo']}'.", "3.1.2.3.1", ref_bib[:1],
                          blocos=[ref_bib[0].i]))
    return out


# ------------------------------------------------------------------ paginação
def paginacao(d):
    campos = du.campos_pagina(d.doc)
    intro = next((b for b in d.blocos if b.i == d.i_intro), None)
    sec_intro = intro.secao_word if intro else None
    if not campos:
        return [Achado(CAT_FORM, "Paginação", AJUSTAR, "nenhum número de página encontrado",
                       "Inserir número de página no canto superior direito a partir da Introdução.", "2.2.2")]
    probs = []
    if any(tipo == "rodapé" for _, tipo, _ in campos):
        probs.append("número no rodapé (deve ser no cabeçalho, canto superior direito)")
    if any(tipo == "cabeçalho" and a not in (WD_ALIGN_PARAGRAPH.RIGHT, None) for _, tipo, a in campos):
        probs.append("número não alinhado à direita")
    if sec_intro is not None and any(s < sec_intro for s, _, _ in campos):
        probs.append("número aparece em páginas pré-textuais")
    if sec_intro == 0 and len(d.doc.sections) == 1:
        probs.append("documento sem quebra de seção: não é possível ocultar o número das páginas pré-textuais")
    return [Achado(CAT_FORM, "Paginação", AJUSTAR if probs else OK, "; ".join(probs) or
                   "número no cabeçalho à direita, a partir da Introdução",
                   "Contar todas as folhas desde a folha de rosto, mas exibir o número só a partir da "
                   "Introdução: inserir 'Quebra de seção (próxima página)' antes da Introdução e desvincular "
                   "o cabeçalho ('Vincular ao anterior' desligado)." if probs else "", "2.2.2")]


# ------------------------------------------------------------------ ilustrações e tabelas
RE_ILU = re.compile(r"^(Figura|Quadro|Gr[áa]fico|Fotografia|Mapa|Esquema|Fluxograma|Desenho|Organograma|"
                    r"Planta|Imagem|Retrato)\s+(\d+(?:\.\d+)?)\s*([-–—:.]?)\s*(.*)$")
RE_TAB = re.compile(r"^Tabela\s+(\d+(?:\.\d+)?)\s*([-–—:.]?)\s*(.*)$")


def _vizinho(d, i, passo, limite=3):
    j, n = i + passo, 0
    while 0 <= j < len(d.blocos) and n < limite:
        b = d.blocos[j]
        if b.tipo == "t" or b.texto or (b.tipo == "p" and du.tem_imagem(b.obj)):
            return b
        j += passo
        n += 1
    return None


def _tem_fonte_apos(d, i, limite=4):
    for b in d.blocos[i + 1:i + 1 + limite]:
        if b.tipo == "p" and re.match(r"^Fonte\b", b.texto):
            return b
    return None


def ilustracoes_tabelas(d):
    out = []
    sep_ilu, abaixo, sem_fonte_ilu, seq_ilu, fonte_no_titulo = [], [], [], [], []
    fontes_grandes = []
    ultimos = {}
    for b in d.paragrafos():
        if d.regiao(b.i) == "pre":
            continue
        m = RE_ILU.match(b.texto)
        if m and m.group(1) == "Quadro":
            ant_q, prox_q = _vizinho(d, b.i, -1), _vizinho(d, b.i, +1)
            if (prox_q is not None and prox_q.tipo == "t") or (ant_q is not None and ant_q.tipo == "t"):
                continue  # quadro feito com tabela do Word: verificado em tabelas.py
        if m and "\t" not in b.bruto:
            tipo, num, sep, _ = m.groups()
            if sep not in ("-", "–", "—"):
                sep_ilu.append(T(b.texto, b, f"{tipo} {num}{sep}".strip() if sep.strip() else None))
            ant, prox = _vizinho(d, b.i, -1), _vizinho(d, b.i, +1)
            img_antes = ant is not None and ant.tipo == "p" and du.tem_imagem(ant.obj)
            img_depois = prox is not None and (prox.tipo == "t" or (prox.tipo == "p" and du.tem_imagem(prox.obj)))
            if img_antes and not img_depois:
                abaixo.append(T(b.texto, b))
            alvo = prox.i if img_depois else b.i
            f = _tem_fonte_apos(d, alvo)
            if not f and re.search(r"\bFonte\s*:", b.texto):
                fonte_no_titulo.append(T(b.texto, b, re.search(r"\bFonte\s*:.*", b.texto).group(0)[:60]))
            elif not f:
                sem_fonte_ilu.append(T(b.texto, b))
            elif (du.tamanho_predominante(f.obj, d.doc) or 10) > PF.atual()["fonte_legenda_max"]:
                fontes_grandes.append(T(f.texto, f))
            k = tipo.lower()
            n = num.split(".")[-1]
            if k in ultimos and int(n) not in (ultimos[k] + 1, 1):
                seq_ilu.append(T(f"{tipo} {ultimos[k]} → {num}", b, f"{tipo} {num}"))
            ultimos[k] = int(n)
            continue
    def add(item, lista, corr, sec):
        out.append(Achado(CAT_ILU, item, AJUSTAR if lista else OK,
                          f"{len(lista)} ocorrência(s)" if lista else "sem problemas",
                          corr if lista else "", sec, _amostra(lista)))

    add("Ilustrações: título no formato 'Figura 1 – Título'", sep_ilu,
        "Palavra designativa + número + travessão + título.", "2.6")
    if fonte_no_titulo:
        out.append(Achado(CAT_ILU, "Fonte em linha própria abaixo da ilustração", VERIFICAR,
                          f"{len(fonte_no_titulo)} ilustração(ões) com a fonte no mesmo parágrafo do título",
                          "A fonte existe, mas deve ficar abaixo da imagem, em linha própria, fonte 10 ou 11 "
                          "(ex.: 'Fonte: Slack (1999).').", "2.6", _amostra(fonte_no_titulo)))
    add("Ilustrações: título acima da imagem", abaixo,
        "O título fica na parte superior; abaixo vêm fonte, legenda e notas.", "2.6")
    add("Ilustrações: fonte indicada abaixo", sem_fonte_ilu,
        "Fonte obrigatória, mesmo se for do autor (ex.: 'Fonte: elaborado pelo autor.').", "2.6")
    add("Fonte das ilustrações/tabelas em tamanho menor", fontes_grandes,
        f"Fonte, legenda e notas em tamanho menor (até {PF.fmt(PF.atual()['fonte_legenda_max'])}).", "2.6")
    add("Numeração sequencial de ilustrações", seq_ilu,
        "Numerar em sequência no documento ou por capítulo (4.1, 4.2...).", "2.6")
    return out


def _tem_bordas_verticais(tabela):
    tblpr = tabela._tbl.tblPr
    bordas = tblpr.find(qn("w:tblBorders")) if tblpr is not None else None

    def ativa(el):
        return el is not None and el.get(qn("w:val")) not in ("nil", "none")

    if bordas is not None:
        return any(ativa(bordas.find(qn(f"w:{lado}"))) for lado in ("insideV", "left", "right", "start", "end"))
    nome = (tabela.style.name if tabela.style is not None else "") or ""
    if re.search(r"grid|grade", nome, re.I):
        return True
    for tc in tabela._tbl.iter(qn("w:tcBorders")):
        if any(ativa(tc.find(qn(f"w:{lado}"))) for lado in ("left", "right", "start", "end")):
            return True
    return False


# ------------------------------------------------------------------ alíneas
def alineas(d):
    probs = []
    pars = [b for b in d.paragrafos() if d.regiao(b.i) != "pre"]
    grupos, atual = [], []
    for b in pars:
        if re.match(r"^[a-z]\)\s", b.texto):
            atual.append(b)
        elif b.texto:
            if atual:
                grupos.append(atual)
                atual = []
    if atual:
        grupos.append(atual)
    for g in grupos:
        ant = next((x for x in reversed(d.blocos[:g[0].i]) if x.tipo == "p" and x.texto), None)
        if ant is not None and not ant.texto.endswith(":"):
            probs.append(T(f"texto antes das alíneas não termina em dois-pontos: {ant.texto[:60]}", ant))
        for k, b in enumerate(g):
            corpo = b.texto[3:].strip()
            fim = "." if k == len(g) - 1 else ";"
            if corpo and corpo[0].isupper() and not re.match(r"^[A-Z]{2,}", corpo):
                probs.append(T(f"começa com maiúscula: {b.texto[:60]}", b, b.texto[:3 + len(corpo.split()[0])]))
            if not (corpo.endswith(fim) or corpo.endswith(":")):
                probs.append(T(f"deveria terminar em '{fim}': {b.texto[:60]}", b, b.texto[-12:].strip() or None))
    return [Achado(CAT_FORM, "Alíneas", AJUSTAR if probs else OK,
                   f"{len(probs)} ocorrência(s)" if probs else
                   ("sem problemas" if grupos else "nenhuma alínea encontrada"),
                   "Texto anterior termina em ':'; alíneas 'a)', iniciando com minúscula e terminando em ';' "
                   "(a última em '.')." if probs else "", "2.3.1", _amostra(probs, 5))]


def origem_do_arquivo(d):
    """Detecta arquivo convertido de PDF (caixas de texto, quadros e quebras de seção em excesso)."""
    corpo = d.doc.element.body
    caixas = len(corpo.findall(".//" + qn("w:txbxContent")))
    quadros = len(corpo.findall(".//" + qn("w:framePr")))
    secoes = len(d.doc.sections)
    n_par = max(1, sum(1 for b in d.blocos if b.tipo == "p" and b.texto))
    sinais = []
    if caixas >= 15:
        sinais.append(f"{caixas} caixas de texto")
    if quadros >= 15:
        sinais.append(f"{quadros} quadros de posicionamento")
    if secoes >= 20 and secoes > n_par / 15:
        sinais.append(f"{secoes} quebras de seção")
    if len(sinais) < 1 or (len(sinais) == 1 and caixas < 40 and secoes < 40):
        return []
    return [Achado(CAT_EST, "Arquivo convertido de PDF", VERIFICAR, "; ".join(sinais),
                   "O arquivo parece ter sido gerado a partir de um PDF. A formatação (margens, recuos, "
                   "espaçamentos) pode não refletir o original, e textos dentro de caixas não são revisados. "
                   "Peça ao aluno o .docx original.", "—")]


def verificar(d):
    achados = []
    fs = (origem_do_arquivo, papel_e_margens, paginacao)
    if getattr(d, "pdf", None) is not None:  # PDF: margens e paginação medidas na página
        from . import pdf_regras
        fs = (pdf_regras.arquivo, pdf_regras.papel_e_margens, pdf_regras.paginacao)
    for f in fs + (texto_principal, alineas, elementos_obrigatorios, resumos,
                   secoes, titulos_sem_indicativo, ilustracoes_tabelas):
        try:
            achados += f(d)
        except Exception as e:  # uma verificação com falha não derruba as outras
            achados.append(Achado("Sistema", f.__name__, VERIFICAR, f"verificação não concluída: {e}"))
    return achados
