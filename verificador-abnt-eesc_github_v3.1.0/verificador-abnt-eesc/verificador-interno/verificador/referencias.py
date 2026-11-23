"""Citações no texto e lista de referências (Diretrizes USP, seções 4 a 8; NBR 10520 e 6023)."""
import difflib
import re
from collections import Counter
from dataclasses import dataclass, field

from docx.enum.text import WD_ALIGN_PARAGRAPH

from . import cache as _cache
from . import docx_util as du
from . import perfil as PF
from .documento import sem_acento
from .modelo import AJUSTAR, OK, VERIFICAR, Achado, T

CAT_CIT = "Citações"
CAT_REF = "Referências"

ANO = r"(?:1[5-9]\d{2}|20\d{2}|\[[^\]]{2,12}\]|s\.\s?d\.)[a-z]?"
PART = r"(?:[Dd]e|[Dd]a|[Dd]o|[Dd]os|[Dd]as|[Vv]an|[Vv]on|[Dd]el|Filho|Júnior|Junior|Neto|Sobrinho|Jr\.?)"
PART_INI = r"(?:De|Da|Do|Dos|Das|Van|Von|Del|Di)"
NOME = rf"(?:{PART_INI}\s)?[A-ZÀ-Ý][\wÀ-ÿ'\-]+(?:\s{PART})*"

# Narrativa: Silva (2020), Ramos e Pestana (2014), Oliveira, Ferreira e Mascarenhas (2013), Sena et al. (2012)
RE_NARRATIVA = re.compile(
    rf"({NOME}(?:(?:,\s|\se\s){NOME})*(?:\set\s?al\.)?)\s\(({ANO}(?:,\s*{ANO})*)(?:,\s*(p\.\s*[^)]*))?\)")
RE_PARENTESE = re.compile(r"\(([^()]*?(?:1[5-9]\d{2}|20\d{2})[a-z]?[^()]*)\)")

MESES_PT = {"jan.", "fev.", "mar.", "abr.", "maio", "jun.", "jul.", "ago.", "set.", "out.", "nov.", "dez."}
MESES_ERRADOS = re.compile(
    r"\b(?:(?:janeiro|fevereiro|março|abril|junho|julho|agosto|setembro|outubro|novembro|dezembro)\b"
    r"|(?:maio|mai)\.|(?:jan|fev|mar|abr|jun|jul|ago|set|out|nov|dez)(?![.\w]))", re.I)


def norm(s):
    s = sem_acento(s).lower()
    s = re.sub(r"\bet\s?al\.?", "", s)
    s = re.sub(r"\[\.\.\.\]|\[…\]", "", s)
    return re.sub(r"[^a-z0-9 ]", " ", s).split()


@dataclass
class Citacao:
    autor: str          # como aparece
    primeiro: list      # tokens normalizados do primeiro autor/entidade/título
    ano: str
    pagina: str
    trecho: str
    direta: bool = False
    caixa_alta: bool = False
    parentese: bool = False
    i: int = None       # bloco onde a chamada está
    marca: str = ""     # a chamada exatamente como aparece no texto


@dataclass
class Referencia:
    bloco: object               # primeiro parágrafo da referência
    texto: str
    marcado: str
    tokens: list
    anos: list = field(default_factory=list)
    tipo: str = "livro"
    blocos: list = field(default_factory=list)  # todos os parágrafos (referência quebrada em linhas)

    @property
    def i(self):
        return self.bloco.i if self.bloco is not None else None


# ------------------------------------------------------------------ citações
CONECTIVOS = {
    "porém", "assim", "portanto", "completando", "segundo", "conforme", "para", "como", "ainda", "já", "então",
    "logo", "além", "também", "contudo", "entretanto", "todavia", "ademais", "inclusive", "finalmente", "enfim",
    "desta", "dessa", "deste", "desse", "neste", "nesse", "nesta", "nessa", "em", "de", "por", "com", "sobre",
    "ou", "e", "o", "a", "os", "as", "um", "uma", "no", "na", "nos", "nas", "do", "da", "dos", "das",
    "mas", "pois", "quando", "onde", "isso", "isto", "aqui", "dessa forma", "desse modo", "adicionalmente",
    "resumindo", "concluindo", "complementando", "analogamente", "similarmente", "igualmente", "outrossim",
    "posteriormente", "anteriormente", "recentemente", "atualmente", "hoje", "nesse sentido", "de acordo",
    "consoante", "acrescenta", "afirma", "explica", "destaca", "ressalta", "observa", "cita",
}


def _limpar_autores(autores, nomes_refs=None):
    """Tira do início da chamada narrativa as palavras que não são autor ("Porém, Fleury (2000)")."""
    nomes = [x.strip() for x in re.split(r",\s|\se\s", autores) if x.strip()]
    sufixo = " et al." if re.search(r"et\s?al", autores) else ""
    nomes = [re.sub(r"\set\s?al\.?$", "", x) for x in nomes]

    def em_refs(n):
        t = norm(n)
        return bool(t) and nomes_refs is not None and t[-1] in nomes_refs

    while len(nomes) > 1 and (nomes[0].lower() in CONECTIVOS or
                              (nomes_refs is not None and not em_refs(nomes[0]) and em_refs(nomes[1]))):
        nomes.pop(0)
    if len(nomes) == 1:
        return nomes[0] + sufixo
    return ", ".join(nomes[:-1]) + " e " + nomes[-1] + sufixo


def _primeiro_autor(s):
    s = re.split(r",\s|\se\s|;\s?|\set\s?al", s.strip())[0]
    return norm(s)


class ListaCitacoes(list):
    """Lista de citações que também guarda as chamadas com 'e' ou ';' entre autores nos parênteses."""

    def __init__(self):
        super().__init__()
        self.e_parenteses = []
        self.pv_parenteses = []


_PEQUENAS = {"de", "da", "do", "dos", "das", "e", "para", "em"}


def _entidade_completa(antes, autor, refs):
    """Se a palavra antes do '(ano)' é o fim do nome de uma entidade ("São Paulo" em "Universidade de São
    Paulo (2014)"), devolve o nome inteiro — o que bate com o começo de uma referência. Senão, o autor."""
    palavras = re.findall(r"[\wÀ-ÿ'\-]+", antes)[-7:]
    melhor = autor
    for k in range(1, len(palavras) + 1):
        trecho = palavras[-k:]
        if not (trecho[0][:1].isupper()):
            continue
        if any(not (w[:1].isupper() or w.lower() in _PEQUENAS) for w in trecho):
            break
        candidato = " ".join(trecho + [autor])
        toks = norm(candidato)
        if any(_mesma_entrada(toks, r.tokens) for r in refs):
            melhor = candidato
    return melhor


def extrair_citacoes(d, refs=None):
    cits = ListaCitacoes()
    e_parenteses = cits.e_parenteses
    pv_parenteses = cits.pv_parenteses
    nomes_refs = None
    if refs:
        nomes_refs = {tok for r in refs for tok in r.tokens[:12] if len(tok) > 2}
    for b in d.texto_do_corpo():
        t = b.texto
        aspas = [(q.start(), q.end()) for q in re.finditer(r"[\"“][^\"”]{15,}[\"”]", t)]

        def junto_de_aspas(ini, fim, narrativa=False, aspas=aspas):
            # (Autor, ano) logo depois do fecha-aspas, ou "Autor (ano)," logo antes do abre-aspas
            return any(0 <= ini - qf <= 12 or (narrativa and 0 <= qi - fim <= 15) for qi, qf in aspas)

        for m in RE_NARRATIVA.finditer(t):
            autores, anos, pag = m.group(1), m.group(2), m.group(3) or ""
            autores = _limpar_autores(autores, nomes_refs)
            marca = m.group(0)
            if refs and "," not in autores and " e " not in autores:
                completo = _entidade_completa(t[:m.start(1)], autores, refs)
                if completo != autores:
                    autores = completo
                    marca = completo + t[m.end(1):m.end()] if completo + t[m.end(1):m.end()] in t else marca
            primeiro = re.split(r",\s|\se\s|\set\s?al", autores)[0].strip()
            caixa = bool(re.fullmatch(r"[A-ZÀ-Ý][A-ZÀ-Ý\-]{2,}", primeiro))
            for ano in re.split(r",\s*", anos):
                cits.append(Citacao(autores, _primeiro_autor(autores), ano, pag, t[:120],
                                    junto_de_aspas(m.start(), m.end(), True), caixa, i=b.i, marca=marca))
        for m in RE_PARENTESE.finditer(t):
            dentro = m.group(1)
            if re.match(r"^\s*(19|20)\d{2}", dentro):
                continue  # é o '(2013)' de uma citação narrativa
            if "apud" in dentro:
                dentro = dentro.split("apud", 1)[1]
            acum = []
            for item in [x.strip() for x in dentro.split(";")]:
                mm = re.match(rf"^(.*?),\s*({ANO}(?:,\s*{ANO})*)(?:,\s*(.*))?$", item)
                if not mm:
                    acum.append(item)
                    continue
                if acum:
                    pv_parenteses.append(T(f"({'; '.join(acum + [item])})", b, m.group(0)))
                nome = (acum[0] if acum else mm.group(1))
                acum = []
                if not re.search(r"[A-Za-zÀ-ÿ]", nome):
                    continue
                pag = mm.group(3) or ""
                if re.search(r"\S\s+e\s+[A-ZÀ-Ý]", nome):
                    e_parenteses.append(T(f"({item})", b, m.group(0)))
                    nome = re.split(r"\s+e\s+", nome)[0]
                caixa = bool(re.fullmatch(r"[A-ZÀ-Ý][A-ZÀ-Ý\s\-]{2,}", nome.strip()))
                for ano in re.split(r",\s*", mm.group(2)):
                    cits.append(Citacao(nome, _primeiro_autor(nome), ano, pag, t[:120],
                                        junto_de_aspas(m.start(), m.end()), caixa, True, i=b.i, marca=m.group(0)))
    return cits


# ------------------------------------------------------------------ referências
def classificar(t):
    k = sem_acento(t)
    if re.search(r"\bIn:\s*[A-Z ]{4,},?\s*\d*\.?,?\s*(19|20)\d{2}", k) or re.search(r"\b(Anais|Proceedings|Resumos)\b", t):
        return "evento"
    if re.search(r"\bIn:", t):
        return "capitulo"
    if re.search(r"(Disserta[cç][aã]o|Tese|Trabalho de Conclus[aã]o|Monografia)\s*\(", t):
        return "academico"
    if re.search(r"\bNBR\b|ISO\s\d", t):
        return "norma"
    if re.search(r"\b(Lei|Decreto|Resolu[cç][aã]o|Portaria|Medida provis[oó]ria|Constitui[cç][aã]o|S[uú]mula)\b", t):
        return "legislacao"
    if re.search(r"Dep[oó]sito:|Depositante:", t):
        return "patente"
    if re.search(r"\bv\.\s*\d+|\bn\.\s*\d+|\be\d{3,}\b", t) and re.search(r"\bpp?\.\s*\d+\s*-\s*\d+|\be\d{3,}", t):
        return "artigo"
    if re.search(r"Dispon[ií]vel em", t) and not re.search(r":\s*[^.]+,\s*(19|20)\d{2}", t):
        return "online"
    return "livro"


def extrair_referencias(d):
    """Referências da lista (linhas quebradas já juntadas). Durante a análise, calculadas uma vez."""
    return list(_cache.memo(("referencias", d), lambda: _extrair_referencias(d)))


def _extrair_referencias(d):
    refs = []
    for grupo in d.referencias_agrupadas():
        t = " ".join(b.texto for b in grupo)
        t = re.sub(r"\s+", " ", t).strip()
        marcado = " ".join(du.texto_marcado(b.obj) for b in grupo)
        anos = re.findall(r"\b((?:1[5-9]\d{2}|20\d{2})[a-z]?)\b", re.split(r"Acesso em", t)[0])
        corpo = re.sub(r"^\[?\d+[\].]\s*", "", t)
        refs.append(Referencia(grupo[0], t, marcado, norm(corpo), anos, classificar(t), list(grupo)))
    return refs


def autor_chave(texto):
    """Elemento de entrada (autor/entidade/título) normalizado, para ordenar a lista."""
    corpo = re.sub(r"^\[?\d+[\].]\s*", "", texto)
    for m in re.finditer(r"\.(?=\s+(\S+))", corpo):
        prox = m.group(1)
        if not re.match(r"^[A-ZÀ-Ý]\.", prox) and len(re.sub(r"\W", "", prox)) >= 2:
            return " ".join(norm(corpo[:m.start()]))
    return " ".join(norm(re.split(r"[.:]", corpo, maxsplit=1)[0]))


def _mesma_entrada(primeiro, tokens):
    """O autor/entidade da chamada é o começo da entrada da referência.

    Tolera espaço faltando ou sobrando em nomes de entidade ("NÍVELSUPERIOR" × "Nível Superior")."""
    if tokens[:len(primeiro)] == primeiro:
        return True
    if len(primeiro) >= 3:
        junto = "".join(primeiro)
        return "".join(tokens[:len(primeiro) + 2]).startswith(junto)
    return False


def _casa(cit, ref):
    if not cit.primeiro or not _mesma_entrada(cit.primeiro, ref.tokens):
        return False
    if cit.ano.startswith("[") or cit.ano.startswith("s"):
        return True
    if cit.ano in ref.anos:
        return True
    return cit.ano.isdigit() and cit.ano in [a[:4] for a in ref.anos]


def _casa_flex(cit, ref):
    """2º método, independente do 1º: o sobrenome aparece no começo da referência e o ano bate.

    Tolera diferenças de grafia na entrada (falta de vírgula, partículas, entidade abreviada).
    """
    toks = [x for x in cit.primeiro if len(x) > 2 and x not in ("del", "dos", "das")]
    if not toks:
        return False
    inicio = ref.tokens[:12]
    if not all(x in inicio for x in toks[-1:]):
        return False
    ano = cit.ano.strip("[]?")
    anos = ref.anos or re.findall(r"\b(?:1[5-9]\d{2}|20\d{2})\b", ref.texto)  # online sem data: ano do acesso
    return cit.ano.startswith("[") or any(a[:4] == ano[:4] for a in anos)


def _casa_parecida(cit, ref):
    """3º método: sobrenome com grafia parecida (erro de digitação) e mesmo ano. Devolve o nome da referência."""
    if not cit.primeiro:
        return None
    nome = cit.primeiro[-1]
    ano = cit.ano.strip("[]?")[:4]
    anos = ref.anos or re.findall(r"\b(?:1[5-9]\d{2}|20\d{2})\b", ref.texto)
    if not any(a[:4] == ano for a in anos):
        return None
    for tok in ref.tokens[:3]:
        if len(tok) > 3 and difflib.SequenceMatcher(None, nome, tok).ratio() >= 0.8:
            return re.split(r"[,.]", ref.texto, maxsplit=1)[0]
    return None


def verificar_citacoes(d, cits, refs):
    out = []
    numerico = sum(1 for r in refs if re.match(r"^\[?\d+[\].]\s", r.texto)) > len(refs) / 2 if refs else False

    # dupla verificação: só é "sem referência" se os DOIS métodos falharem;
    # se só o método exato falhar, a grafia difere entre citação e referência (conferir)
    sem_ref, grafia = [], []
    for c in cits:
        if not refs or any(_casa(c, r) for r in refs):
            continue
        if any(_casa_flex(c, r) for r in refs):
            grafia.append(T(f"{c.autor} ({c.ano})", c.i, c.marca))
            continue
        parecida = next((x for x in (_casa_parecida(c, r) for r in refs) if x), None)
        if parecida:
            grafia.append(T(f"{c.autor} ({c.ano}) ≈ {parecida} ({c.ano}) nas referências", c.i, c.marca))
        else:
            sem_ref.append(T(f"{c.autor} ({c.ano})", c.i, c.marca))
    sem_ref, grafia = list(dict.fromkeys(sem_ref)), list(dict.fromkeys(grafia))

    nao_citadas, nao_citadas_b = [], []
    if not numerico:
        for r in refs:
            if not any(_casa(c, r) or _casa_flex(c, r) or _casa_parecida(c, r) for c in cits):
                nao_citadas.append(r.texto)
                nao_citadas_b.append(r.i)

    # caixa alta só é erro para PESSOAS (entrada 'SOBRENOME, X.'); siglas de entidade (IBGE, CAPES) valem
    modo_caixa = PF.atual()["citacao_autor"]
    caixa, mista_par = [], []
    for c in cits:
        if c.caixa_alta and (modo_caixa == "mista" or (modo_caixa == "maiusculas" and not c.parentese)):
            pessoa = any(_casa(c, r) and re.match(r"^[\wÀ-ÿ' \-]+,\s", r.texto) for r in refs)
            if pessoa or len(c.autor.strip()) >= 8:
                caixa.append(T(f"({c.autor}, {c.ano})" if c.parentese else f"{c.autor} ({c.ano})", c.i, c.marca))
        elif modo_caixa == "maiusculas" and c.parentese and not c.caixa_alta and \
                re.match(r"^[A-ZÀ-Ý][a-zà-ÿ]", c.autor.strip()):
            mista_par.append(T(f"({c.autor}, {c.ano})", c.i, c.marca))
    caixa, mista_par = list(dict.fromkeys(caixa)), list(dict.fromkeys(mista_par))
    diretas_sem_pag = list(dict.fromkeys(T(c.trecho, c.i, c.marca) for c in cits
                                         if c.direta and not re.search(r"p\.\s*\d", c.pagina)))

    out.append(Achado(CAT_CIT, "Citações encontradas", OK if cits else VERIFICAR,
                      f"{len(cits)} chamadas autor-data identificadas" if cits else
                      "nenhuma citação autor-data identificada (sistema numérico ou formato não reconhecido)"))
    if refs and not numerico:
        out.append(Achado(CAT_CIT, "Toda citação está nas referências", AJUSTAR if sem_ref else OK,
                          f"{len(sem_ref)} citação(ões) sem referência correspondente" if sem_ref else "todas encontradas",
                          "Incluir a obra nas Referências ou corrigir autor/ano da chamada (devem ser idênticos)."
                          if sem_ref else "", "4", sem_ref[:300]))
        if grafia:
            out.append(Achado(CAT_CIT, "Citação e referência com grafias diferentes", VERIFICAR,
                              f"{len(grafia)} citação(ões) cuja entrada difere da referência",
                              "A chamada deve repetir a entrada da referência (mesmo sobrenome/entidade e ano). "
                              "Confira a grafia nas duas.", "4", grafia[:300]))
        out.append(Achado(CAT_CIT, "Toda referência foi citada no texto", VERIFICAR if nao_citadas else OK,
                          f"{len(nao_citadas)} referência(s) sem citação localizada" if nao_citadas else "todas citadas",
                          "Toda obra das Referências deve ser citada no texto (conferir: pode ser citação em "
                          "tabela, figura ou formato não reconhecido)." if nao_citadas else "", "4",
                          [T(t[:120], i) for t, i in zip(nao_citadas[:300], nao_citadas_b)]))
    sep = PF.atual()["separador_parenteses"]
    if sep == ";":
        e_par = list(dict.fromkeys(getattr(cits, "e_parenteses", [])))
        out.append(Achado(CAT_CIT, "Autores separados por ';' dentro dos parênteses", AJUSTAR if e_par else OK,
                          f"{len(e_par)} chamada(s) com 'e' entre autores" if e_par else "sem problemas",
                          "Entre parênteses, separe os autores por ponto e vírgula: (Bowersox; Closs, 2001). "
                          "Fora dos parênteses usa-se 'e': Bowersox e Closs (2001)." if e_par else "", "5.1",
                          e_par[:300]))
    elif sep == "e":
        pv = list(dict.fromkeys(getattr(cits, "pv_parenteses", [])))
        out.append(Achado(CAT_CIT, "Autores separados por 'e' dentro dos parênteses", AJUSTAR if pv else OK,
                          f"{len(pv)} chamada(s) com ';' entre autores da mesma obra" if pv else "sem problemas",
                          "Pela norma escolhida, autores da mesma obra entre parênteses são ligados por 'e': "
                          "(Bowersox e Closs, 2001). O ';' separa obras diferentes." if pv else "", "5.1", pv[:300]))
    if modo_caixa == "mista":
        out.append(Achado(CAT_CIT, "Autor em maiúsculas e minúsculas na chamada", AJUSTAR if caixa else OK,
                          f"{len(caixa)} chamada(s) em CAIXA ALTA" if caixa else "sem problemas",
                          "Pela NBR 10520:2023: (Silva, 2020), não (SILVA, 2020). Siglas de entidades podem "
                          "ficar em maiúsculas (IBGE)." if caixa else "", "5", caixa[:300]))
    elif modo_caixa == "maiusculas":
        out.append(Achado(CAT_CIT, "Autor em MAIÚSCULAS dentro dos parênteses", AJUSTAR if mista_par else OK,
                          f"{len(mista_par)} chamada(s) em minúsculas" if mista_par else "sem problemas",
                          "Pela norma escolhida, dentro dos parênteses o sobrenome vai em maiúsculas: "
                          "(SILVA, 2020)." if mista_par else "", "5", mista_par[:300]))
        if caixa:
            out.append(Achado(CAT_CIT, "Autor fora dos parênteses em maiúsculas e minúsculas", AJUSTAR,
                              f"{len(caixa)} chamada(s) narrativa(s) em CAIXA ALTA",
                              "Fora dos parênteses, só a inicial maiúscula: Silva (2020).", "5", caixa[:300]))
    out.append(Achado(CAT_CIT, "Citação direta com página", VERIFICAR if diretas_sem_pag else OK,
                      f"{len(diretas_sem_pag)} parágrafo(s) com trecho entre aspas e chamada sem página"
                      if diretas_sem_pag else "sem problemas",
                      "Em citação direta, a página é obrigatória: (Silva, 2020, p. 15)." if diretas_sem_pag else "",
                      "4.1", diretas_sem_pag[:300]))

    tem_notas = _tem_notas_rodape(d)
    if numerico and tem_notas:
        out.append(Achado(CAT_CIT, "Sistema numérico com notas de rodapé", AJUSTAR,
                          "referências numeradas e o documento tem notas de rodapé",
                          "Com notas de rodapé, usar o sistema autor-data.", "5.2"))
    etal = sum(1 for c in cits if "et al" in c.autor)
    return out, numerico, etal


def _tem_notas_rodape(d):
    for rel in d.doc.part.rels.values():
        if rel.reltype.endswith("/footnotes"):
            xml = rel.target_part.blob.decode("utf8", "ignore")
            return len(re.findall(r"<w:footnote [^>]*w:id=\"[1-9]", xml)) > 0
    return False


def _destaque(par):
    """Recurso de destaque usado no parágrafo: negrito, itálico, sublinhado ou nenhum."""
    tipos = set()
    for r in par.runs:
        if not r.text.strip() or len(r.text.strip()) < 3:
            continue
        if du.negrito_run(r, par):
            tipos.add("negrito")
        if du.italico_run(r, par):
            tipos.add("itálico")
        if r.underline:
            tipos.add("sublinhado")
    return tipos


def verificar_referencias(d, refs, numerico):
    out = []
    if not refs:
        return [Achado(CAT_REF, "Lista de referências", AJUSTAR, "não encontrada ou vazia",
                       "Incluir a seção REFERÊNCIAS ao final do texto.", "3.1.2.3.1")]

    RP = PF.atual()["referencias"]
    p = {k: [] for k in ("entrada", "destaque", "ponto_final", "acesso", "disponivel", "edicao", "paginas",
                         "in", "doi_ordem", "espaco", "alinhamento", "meses", "quebrada")}
    prenomes = Counter()
    destaques = Counter()
    autores4 = Counter()
    for r in refs:
        t = r.texto
        pars = [b.obj for b in r.blocos]
        reg = lambda k, t=t, r=r: p[k].append((t, r.i))
        corpo = re.sub(r"^\[?\d+[\].]\s*", "", t) if numerico else t
        primeiro = re.split(r"[,.:]", corpo, maxsplit=1)[0]
        if primeiro.upper() != primeiro:
            reg("entrada")
        tipos = set().union(*(_destaque(x) for x in pars))
        exigido = {"negrito": "negrito", "italico": "itálico"}.get(RP["destaque"])
        if not tipos or (exigido and exigido not in tipos):
            reg("destaque")
        destaques[" + ".join(sorted(tipos)) or "nenhum"] += 1
        if not t.rstrip().endswith((".", ".]")) and not re.search(r"https?://\S+>?$", t):
            reg("ponto_final")
        if re.search(r"Dispon[ií]vel em", t):
            if not re.search(r"Dispon[ií]vel em:", t):
                reg("disponivel")
            if not re.search(r"Acesso em:\s*\d{1,2}\s+\S+\s+(19|20)\d{2}", t):
                reg("acesso")
        if re.search(r"Acesso\s*(em)?\s*:?", t):
            if MESES_ERRADOS.search(t.split("Acesso", 1)[1][:30]):
                reg("meses")
        if re.search(r"\d+\s*[ªºa]\s*(ed\b|edi[çc][ãa]o)|\d+\.\s*edi[çc][ãa]o|\bEd\.\s", t):
            reg("edicao")
        if re.search(r"\bpp\.|\bp\.\d|\bpág", t):
            reg("paginas")
        if re.search(r"\b(IN|in):\s", t) and not re.search(r"\bIn:\s", t):
            reg("in")
        if "DOI" in t.upper() and re.search(r"Dispon[ií]vel", t) and t.upper().find("DOI") > t.find("Dispon"):
            reg("doi_ordem")
        if any(du.entrelinhas(x, d.doc) > 1.15 for x in pars):
            reg("espaco")
        if RP["alinhar_esquerda"] and any(du.alinhamento(x) == WD_ALIGN_PARAGRAPH.JUSTIFY for x in pars):
            reg("alinhamento")
        if len(pars) > 1:
            reg("quebrada")
        if re.match(r"^[A-ZÀ-Ý' \-]+,\s*([A-ZÀ-Ý][a-zà-ÿ]+)", corpo):
            prenomes["extenso"] += 1
        elif re.match(r"^[A-ZÀ-Ý' \-]+,?\s*[A-ZÀ-Ý]\.", corpo):
            prenomes["abreviado"] += 1
        n_aut = len(re.findall(r";\s*[A-ZÀ-Ý' \-]{2,},", corpo.split(".")[0] if "et al" not in corpo else ""))
        if "et al" in corpo:
            autores4["et al."] += 1
        elif n_aut >= 3:
            autores4["todos"] += 1

    def add(item, lista, corr, sec, status=AJUSTAR):
        out.append(Achado(CAT_REF, item, status if lista else OK,
                          f"{len(lista)} referência(s)" if lista else "sem problemas",
                          corr if lista else "", sec, [x[:120] for x, _ in lista[:300]],
                          blocos=[i for _, i in lista[:300] if i is not None]))

    add("Entrada pelo sobrenome em MAIÚSCULAS", p["entrada"],
        "Iniciar pelo SOBRENOME em maiúsculas (ou entidade/título em maiúsculas): SILVA, J. A.", "6")
    exigido_txt = {"negrito": "negrito", "italico": "itálico"}.get(RP["destaque"])
    add(f"Título em destaque ({exigido_txt})" if exigido_txt else "Título em destaque", p["destaque"],
        f"Destacar o título da obra (ou o nome do periódico/evento) em {exigido_txt}, igual em toda a lista."
        if exigido_txt else
        "Destacar o título da obra (ou o nome do periódico/evento) com negrito — padrão das Diretrizes USP — "
        "ou itálico, igual em toda a lista.", "6", AJUSTAR if exigido_txt else VERIFICAR)
    add("Referência termina com ponto", p["ponto_final"], "Finalizar cada referência com ponto.", "6")
    add("'Disponível em:' com dois-pontos", p["disponivel"], "Usar 'Disponível em: URL.'", "6.6")
    add("'Acesso em: dia mês. ano.'", p["acesso"], "Ex.: 'Acesso em: 5 abr. 2019.'", "6.6")
    add("Meses abreviados (NBR 6023)", p["meses"],
        "jan. fev. mar. abr. maio jun. jul. ago. set. out. nov. dez. (maio sem ponto).", "8")
    add("Edição no formato '2. ed.'", p["edicao"],
        "Edição só a partir da 2ª, no idioma da obra: '2. ed.', '3rd ed.'.", "8")
    add("Páginas com 'p.'", p["paginas"], "Usar 'p. 12-34' (não 'pp.' nem 'p.12').", "6.2.1")
    add("'In:' em capítulos e eventos", p["in"], "Grafar 'In:' (I maiúsculo, n minúsculo).", "6.1.2")
    add("DOI antes de 'Disponível em'", p["doi_ordem"], "Ordem: DOI: ... Disponível em: ... Acesso em: ...", "6.2.1")
    add("Espaço simples nas referências", p["espaco"],
        "Referências em espaço simples, separadas entre si por uma linha em branco.", "2.2.1")
    add("Referências alinhadas à esquerda", p["alinhamento"], "Alinhar à esquerda (não justificar).", "6", VERIFICAR)
    add("Referência inteira num só parágrafo", p["quebrada"],
        "Cada referência deve ser um parágrafo único (sem Enter no meio); a linha em branco separa uma "
        "referência da outra.", "2.2.1", VERIFICAR)

    if len(destaques) > 1 and destaques.most_common(1)[0][1] < len(refs):
        out.append(Achado(CAT_REF, "Destaque uniforme em toda a lista", AJUSTAR,
                          ", ".join(f"{k}: {v}" for k, v in destaques.most_common()),
                          "Usar o mesmo recurso de destaque (de preferência negrito, como nas Diretrizes USP) "
                          "em todas as referências.", "6"))
    if prenomes["extenso"] and prenomes["abreviado"]:
        out.append(Achado(CAT_REF, "Prenomes padronizados", AJUSTAR,
                          f"{prenomes['abreviado']} abreviados e {prenomes['extenso']} por extenso",
                          "Prenomes abreviados ou por extenso, mas iguais em toda a lista.", "6.1"))
    if autores4["et al."] and autores4["todos"]:
        out.append(Achado(CAT_REF, "4 ou mais autores: mesmo critério", VERIFICAR,
                          f"{autores4['et al.']} com et al. e {autores4['todos']} com todos os autores",
                          "Escolher um critério (todos ou et al.) e aplicar em citações e referências.", "6.1.1.4"))

    # duplicadas (mesma obra listada duas vezes)
    dup = []
    for i, r in enumerate(refs):
        for r2 in refs[:i]:
            a1, a2 = " ".join(r.tokens[:25]), " ".join(r2.tokens[:25])
            if a1 and r.anos[:1] == r2.anos[:1] and difflib.SequenceMatcher(None, a1, a2).ratio() >= 0.92:
                dup.append((f"{r.texto[:90]}  ⇄  {r2.texto[:60]}", r.i))
                break
    add("Referência repetida", dup, "A mesma obra aparece duas vezes na lista: mantenha só uma.", "7")

    # mesmo autor e mesmo ano: diferenciar com letras (2020a, 2020b)
    por_chave = {}
    for r in refs:
        anos_r = [a for a in r.anos[:1]]
        if anos_r:
            por_chave.setdefault((autor_chave(r.texto)[:40], anos_r[0][:4]), []).append(r)
    sem_letra = []
    for (_, ano), lista in por_chave.items():
        if len(lista) > 1 and not all(re.search(rf"\b{ano}[a-z]\b", r.texto) for r in lista):
            sem_letra += [(r.texto, r.i) for r in lista]
    add("Mesmo autor e ano diferenciados (2020a, 2020b)", sem_letra,
        "Obras do mesmo autor no mesmo ano: acrescente letras minúsculas após o ano, na referência e na "
        "citação (Silva, 2020a; Silva, 2020b).", "5.1")

    if not numerico and RP["ordem_alfabetica"]:
        fora = []
        chaves = [autor_chave(r.texto) for r in refs]
        for a, b, r in zip(chaves, chaves[1:], refs[1:]):
            if b < a:  # mesmo autor em sequência é aceito (ordem cronológica ou por título)
                fora.append((r.texto, r.i))
        add("Ordem alfabética", fora, "Ordenar alfabeticamente pelo elemento de entrada (autor, entidade ou "
            "título). Obras do mesmo autor: em ordem cronológica.", "7.1")
    return out


def verificar(d):
    refs = extrair_referencias(d)
    cits = extrair_citacoes(d, refs)
    a, numerico, _ = verificar_citacoes(d, cits, refs)
    return a + verificar_referencias(d, refs, numerico), refs
