"""Mapa do documento: blocos em ordem, regiões (pré-textual, textual, pós-textual)."""
import re
import unicodedata
from dataclasses import dataclass, field
from functools import lru_cache

from docx import Document
from docx.text.paragraph import Paragraph

from . import cache
from . import docx_util as du
from .cache import qn


def sem_acento(s):
    return "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")


@lru_cache(maxsize=4096)  # ~600 textos distintos por trabalho; o limite segura a memória do servidor
def chave(s):
    return re.sub(r"\s+", " ", sem_acento(s).upper()).strip()


@dataclass
class Bloco:
    i: int
    tipo: str          # "p" ou "t"
    obj: object
    texto: str = ""
    secao_word: int = 0  # índice da seção do Word (quebras de seção)
    _bruto: str | None = field(default=None, repr=False, compare=False)

    @property
    def bruto(self):
        """Texto do parágrafo como o Word o guarda (com tabulações e quebras de linha).

        Durante a análise é lido uma vez só; fora dela é sempre lido ao vivo."""
        if self.tipo != "p":
            return ""
        if self._bruto is not None:
            return self._bruto
        t = self.obj.text
        if cache.ativo():
            self._bruto = t
        return t


TITULOS_SEM_NUMERO = [
    "ERRATA", "AGRADECIMENTOS", "AGRADECIMENTO", "RESUMO", "ABSTRACT", "RESUMEN", "RESUME",
    "SUMARIO", "REFERENCIAS", "REFERENCIAS BIBLIOGRAFICAS", "BIBLIOGRAFIA", "GLOSSARIO",
    "INDICE", "LISTA DE ILUSTRACOES", "LISTA DE FIGURAS", "LISTA DE TABELAS", "LISTA DE QUADROS",
    "LISTA DE GRAFICOS", "LISTA DE ABREVIATURAS E SIGLAS", "LISTA DE ABREVIATURAS",
    "LISTA DE SIGLAS", "LISTA DE SIMBOLOS",
]
RE_APENDICE = re.compile(r"^(APENDICE|ANEXO)\s+[A-Z]{1,2}\b")

# início típico de uma referência: [1] / SOBRENOME (maiúsculas) / ENTIDADE / Sobrenome, X. / A TÍTULO...
RE_INICIO_REF = re.compile(
    r"^(\[?\d{1,3}[\].)]\s"
    r"|(?:[A-ZÀ-Ý]\s)?[A-ZÀ-Ý][A-ZÀ-Ý'’\-]+(?:\s+(?:DE|DA|DO|DOS|DAS|E|[A-ZÀ-Ý][A-ZÀ-Ý'’\-]+))*[\s,.;:]"
    r"|[A-ZÀ-Ý][\wÀ-ÿ'’\-]+(?:\s[\wÀ-ÿ'’\-]+){0,3},\s*[A-ZÀ-Ý][\wÀ-ÿ]*\.?)")


def continua_referencia(anterior, atual):
    """True se 'atual' é a continuação (linha quebrada) da referência 'anterior'."""
    a, t = anterior.rstrip(), atual.lstrip()
    if not a or not t:
        return False
    if re.match(r"^\[?\d{1,3}[\].)]\s", t):
        return False                       # nova entrada numerada
    if re.search(r"[,;:–\-&]$", a) or a.endswith((" e", " In")):
        return True                        # a anterior claramente não terminou
    if re.match(r"^([a-zà-ÿ0-9<(]|https?:|www\.|Dispon[ií]vel|Acesso|DOI)", t):
        return True                        # começa com minúscula, número, link...
    if re.match(r"^[A-ZÀ-Ý]\.\s", t):
        return True                        # inicial solta ("K. F. (Org.)")
    return not RE_INICIO_REF.match(t)


class Documento:
    def __init__(self, caminho_ou_arquivo):
        self.doc = Document(caminho_ou_arquivo)
        self.blocos = []
        sec = 0
        for i, (tipo, obj) in enumerate(du.blocos(self.doc)):
            txt = du.texto(obj) if tipo == "p" else ""
            self.blocos.append(Bloco(i, tipo, obj, txt, sec))
            if tipo == "p":
                ppr = obj._p.pPr
                if ppr is not None and ppr.sectPr is not None:
                    sec += 1
        self._mapear()

    # -------------------------------------------------------------- regiões
    def _eh_linha_de_titulo(self, b, padrao):
        """Linha igual ao título (exclui linhas do sumário com nº de página).

        Tolera o que é comum fora do modelo: pontuação no fim ("RESUMO:", "Introdução."),
        "CAPÍTULO 1 – INTRODUÇÃO", "Capítulo 1: Introdução" e minúsculas/acentos."""
        if b.tipo != "p" or not b.texto or "\t" in b.bruto or len(b.texto) > 120:
            return False
        k = chave(b.texto)
        if re.fullmatch(padrao, k):
            return True
        k2 = re.sub(r"[\s.:;,–—-]+$", "", k)
        k2 = re.sub(r"^CAPITULO\s+(\d+|[IVXL]+)\s*[-–—:.]?\s*", lambda m: (m.group(1) + " ") if m.group(1).isdigit()
                    else "", k2)
        return k2 != k and re.fullmatch(padrao, k2) is not None

    def _aproximado(self, canon, prefixos=(), depois_de=None, limiar=0.84):
        """Título por aproximação: começa com o nome esperado ("1 INTRODUÇÃO E OBJETIVOS") ou tem grafia
        parecida ("INTRODUCÃO", "REFERENCIAS"). Guarda o achado em self.aproximados."""
        import difflib
        for b in self.blocos:
            if b.tipo != "p" or not b.texto or "\t" in b.bruto or len(b.texto) > 70:
                continue
            if depois_de is not None and b.i <= depois_de:
                continue
            k = re.sub(r"[\s.:;,–—-]+$", "", chave(b.texto))
            k = re.sub(r"^(CAPITULO\s+\S+\s*[-–—:.]?\s*|\d+(\.\d+)*\.?\s*[-–—:]?\s*)", "", k)
            if not k or len(k.split()) > 8:
                continue
            ok = any(k.startswith(x + " ") for x in (canon, *prefixos)) or \
                difflib.SequenceMatcher(None, k, canon).ratio() >= limiar
            if ok and k != canon:
                self.aproximados.append((canon, b))
                return b
        return None

    def _mapear(self):
        self.aproximados = []
        self.i_intro = next((b.i for b in self.blocos
                             if self._eh_linha_de_titulo(b, r"(1\.?\s*[-–]?\s*)?INTRODUCAO")), None)
        if self.i_intro is None:
            b = self._aproximado("INTRODUCAO")
            self.i_intro = b.i if b else None
        rx_refs = (r"(\d+\.?\s*)?(REFERENCIAS( BIBLIOGRAFICAS)?|BIBLIOGRAFIA( CONSULTADA)?|"
                   r"LISTA DE REFERENCIAS|OBRAS CITADAS|FONTES CONSULTADAS)")
        refs = [b.i for b in self.blocos if self._eh_linha_de_titulo(b, rx_refs)]
        refs = [i for i in refs if self.i_intro is None or i > self.i_intro]
        if not refs:
            b = self._aproximado("REFERENCIAS", depois_de=self.i_intro)
            refs = [b.i] if b else []
        self.i_refs = refs[-1] if refs else None
        fim = [b.i for b in self.blocos if self.i_refs is not None and b.i > self.i_refs
               and b.tipo == "p" and (RE_APENDICE.match(chave(b.texto))
                                      or chave(b.texto) in ("GLOSSARIO", "INDICE"))
               and "\t" not in b.bruto]
        self.i_fim_refs = fim[0] if fim else len(self.blocos)

    def regiao(self, i):
        if self.i_intro is not None and i < self.i_intro:
            return "pre"
        if self.i_refs is not None and i >= self.i_refs:
            return "pos"
        return "textual"

    def paragrafos(self, regiao=None):
        lista = cache.memo(("paragrafos", self, regiao), lambda: [
            b for b in self.blocos if b.tipo == "p" and (regiao is None or self.regiao(b.i) == regiao)])
        return list(lista)  # cópia: quem chama pode alterar a lista à vontade

    def achar_titulo(self, padrao, depois_de=None):
        return cache.memo(("achar_titulo", self, padrao, depois_de),
                          lambda: self._achar_titulo(padrao, depois_de))

    def _achar_titulo(self, padrao, depois_de=None):
        for b in self.blocos:
            if depois_de is not None and b.i <= depois_de:
                continue
            if self._eh_linha_de_titulo(b, padrao):
                return b
        return None

    def anterior_paragrafo(self, i):
        for j in range(i - 1, -1, -1):
            if self.blocos[j].tipo == "p":
                return self.blocos[j].obj
        return None

    def entradas_referencias(self):
        """Parágrafos da lista de referências (sem linhas vazias)."""
        if self.i_refs is None:
            return []
        return [b for b in self.blocos[self.i_refs + 1:self.i_fim_refs] if b.tipo == "p" and b.texto]

    def referencias_agrupadas(self):
        """Lista de referências; cada uma é uma lista de Blocos (linhas quebradas são juntadas)."""
        grupos = []
        for b in self.entradas_referencias():
            if grupos and continua_referencia(grupos[-1][-1].texto, b.texto):
                grupos[-1].append(b)
            else:
                grupos.append([b])
        return grupos

    def texto_do_corpo(self):
        """Onde se procuram citações: parágrafos (exceto a lista de referências) e células de tabelas.

        Parágrafos de célula viram Blocos com o índice da própria tabela (para localizar no documento).
        """
        out = []
        for b in self.blocos:
            if self.i_refs is not None and self.i_refs <= b.i < self.i_fim_refs:
                continue
            if b.tipo == "p" and b.texto:
                out.append(b)
            if b.tipo == "p":
                # caixas de texto (comuns em arquivos convertidos de PDF); ignora a cópia em mc:Fallback
                for tx in b.obj._p.iter(qn("w:txbxContent")):
                    anc, dup = tx.getparent(), False
                    while anc is not None:
                        if anc.tag.endswith("}Fallback"):
                            dup = True
                            break
                        anc = anc.getparent()
                    if dup:
                        continue
                    for p_el in tx.iter(qn("w:p")):
                        par = Paragraph(p_el, b.obj._parent)
                        txt = du.texto(par)
                        if txt:
                            out.append(Bloco(b.i, "p", par, txt, b.secao_word))
            elif b.tipo == "t":
                for p_el in b.obj._tbl.iter(qn("w:p")):
                    par = Paragraph(p_el, b.obj)
                    txt = du.texto(par)
                    if txt:
                        out.append(Bloco(b.i, "p", par, txt, b.secao_word))
        return out
