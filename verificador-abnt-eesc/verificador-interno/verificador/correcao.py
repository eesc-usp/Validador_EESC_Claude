"""Gera uma versão corrigida do .docx do aluno.

Princípios:
- Só corrige o que é seguro corrigir sem entender o conteúdo (formatação e padrões de escrita da norma).
- Alterações de TEXTO entram como controle de alterações do Word (aceitar/rejeitar uma a uma).
  Alterações de FORMATAÇÃO (margem, fonte, espaçamento, bordas) são aplicadas direto.
- O que depende do autor (resumo curto, fonte de figura faltando, citação sem referência, página em
  citação direta, negrito do título) NÃO é inventado: continua no relatório como ajuste manual.
"""
import io
import re
from copy import deepcopy
from datetime import datetime, timezone

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.run import Run

from . import docx_util as du
from . import perfil as PF
from . import referencias as R
from .documento import RE_APENDICE, TITULOS_SEM_NUMERO, Documento, chave
from .estrutura import RE_ILU, RE_SECAO, RE_TAB, _candidato_titulo, _eh_legenda_ou_fonte, _vizinho, citacao_longa

GRUPOS = {
    "margens": "Margens e papel A4",
    "texto": "Fonte e entrelinhas do texto; citações longas",
    "titulos": "Títulos de seção",
    "resumo": "Palavras-chave e keywords",
    "ilustracoes": "Figuras (títulos e fontes)",
    "tabelas": "Tabelas e quadros (título, fonte, bordas, cabeçalho)",
    "alineas": "Alíneas",
    "citacoes": "Citações (autor em maiúsculas e minúsculas)",
    "tipografia": "Espaços duplos, espaço antes da pontuação e títulos de seção (ponto final, espaço, alinhamento)",
    "referencias": "Referências (formato, espaçamento e ordem)",
}

AUTOR = "Biblioteca EESC (verificador)"
MESES = {"janeiro": "jan.", "fevereiro": "fev.", "março": "mar.", "marco": "mar.", "abril": "abr.",
         "maio.": "maio", "mai.": "maio", "junho": "jun.", "julho": "jul.", "agosto": "ago.", "setembro": "set.",
         "outubro": "out.", "novembro": "nov.", "dezembro": "dez."}
PARTICULAS = {"de", "da", "do", "dos", "das", "e", "van", "von", "del"}


# ====================================================================== texto e runs
def _texto_run(r):
    out = []
    for ch in r:
        if ch.tag == qn("w:t"):
            out.append(ch.text or "")
        elif ch.tag == qn("w:tab"):
            out.append("\t")
        elif ch.tag in (qn("w:br"), qn("w:cr")):
            out.append("\n")
    return "".join(out)


_MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
_FORA = {qn("w:del"), qn("w:moveFrom"), qn("w:r"), qn("w:txbxContent"), _MC_FALLBACK}


def _runs(p):
    """Runs visíveis do próprio parágrafo, em ordem (inclui hyperlinks e alterações inseridas).

    Ficam de fora: texto apagado (alteração controlada) e texto de caixas de texto, formas e gráficos
    que estão "dentro" de um run (w:drawing/w:pict) — esse texto não pertence ao parágrafo e, em
    arquivos convertidos de PDF, aparecia duplicado (versão nova + versão de compatibilidade).
    """
    res = []
    for r in p.iter(qn("w:r")):
        anc, fora = r.getparent(), False
        while anc is not None and anc is not p:
            if anc.tag in _FORA:
                fora = True
                break
            anc = anc.getparent()
        if not fora:
            res.append(r)
    return res


def texto_atual(par):
    p = getattr(par, "_p", par)
    return "".join(_texto_run(r) for r in _runs(p))


def _simples(r, p):
    return r.getparent() is p and all(ch.tag in (qn("w:rPr"), qn("w:t")) for ch in r)


def _set_texto(r, s):
    for t in r.findall(qn("w:t")):
        r.remove(t)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = s
    r.append(t)


def _titulo_nome(s):
    palavras = s.lower().split()
    return " ".join(w if (w in PARTICULAS and i) else w[:1].upper() + w[1:] for i, w in enumerate(palavras))


# ====================================================================== corretor
class Corretor:
    def __init__(self, arquivo, rastrear=True, grupos=None):
        self.d = Documento(arquivo)
        self.doc = self.d.doc
        self.rastrear = rastrear
        self.grupos = set(grupos or GRUPOS)
        self.log = []          # dicts: Grupo, Alteração, Trecho
        self.falhas = []       # alterações que não puderam ser feitas com segurança
        self._id = 880000
        self._data = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
        self._ref_ia = set()   # parágrafos de referência substituídos pela IA

    # ------------------------------------------------------------------ registro
    def _reg(self, grupo, alteracao, trecho=""):
        self.log.append({"Grupo": GRUPOS[grupo], "Alteração": alteracao,
                         "Trecho": (trecho[:100] + "…") if len(trecho) > 100 else trecho})

    def _marca(self, tag):
        el = OxmlElement(tag)
        self._id += 1
        el.set(qn("w:id"), str(self._id))
        el.set(qn("w:author"), AUTOR)
        el.set(qn("w:date"), self._data)
        return el

    # ------------------------------------------------------------------ troca de texto
    def trocar(self, par, ini, fim, novo):
        """Troca texto[ini:fim] por 'novo' (str ou lista de (texto, negrito)). True se conseguiu."""
        p = getattr(par, "_p", par)
        segs = [(novo, None)] if isinstance(novo, str) else novo
        segs = [(s, b) for s, b in segs if s]

        def mapa():
            pos, m = 0, []
            for r in _runs(p):
                s = _texto_run(r)
                m.append((r, pos, pos + len(s)))
                pos += len(s)
            return m

        for corte in (fim, ini):
            for r, a, b in mapa():
                if a < corte < b:
                    if not _simples(r, p):
                        return False
                    txt = _texto_run(r)
                    r2 = deepcopy(r)
                    _set_texto(r, txt[:corte - a])
                    _set_texto(r2, txt[corte - a:])
                    r.addnext(r2)
                    break
        m = mapa()
        alvo = [r for r, a, b in m if a >= ini and b <= fim and b > a]
        if any(not _simples(r, p) for r in alvo):
            return False
        if alvo:
            modelo, ancora, depois = alvo[0], None, True
        else:
            antes = [r for r, a, b in m if b == ini and _simples(r, p)]
            apos = [r for r, a, b in m if a == ini and _simples(r, p)]
            if antes:
                modelo, ancora, depois = antes[-1], antes[-1], True
            elif apos:
                modelo, ancora, depois = apos[0], apos[0], False
            else:
                return False

        novos = []
        for s, negrito in segs:
            nr = OxmlElement("w:r")
            rpr = modelo.find(qn("w:rPr"))
            if rpr is not None:
                nr.append(deepcopy(rpr))
            _set_texto(nr, s)
            if negrito is not None:
                Run(nr, None).bold = negrito
            novos.append(nr)

        if self.rastrear:
            if alvo:
                dele = self._marca("w:del")
                alvo[0].addprevious(dele)
                for r in alvo:
                    dele.append(r)
                    for t in r.findall(qn("w:t")):
                        t.tag = qn("w:delText")
                ancora, depois = dele, True
            if novos:
                ins = self._marca("w:ins")
                for nr in novos:
                    ins.append(nr)
                (ancora.addnext if depois else ancora.addprevious)(ins)
        else:
            if alvo:
                ancora, depois = alvo[0], False
            for nr in reversed(novos) if depois else novos:
                (ancora.addnext if depois else ancora.addprevious)(nr)
            for r in alvo:
                p.remove(r)
        return True

    def trocar_regex(self, par, padrao, repl, grupo, descricao, flags=0, limite=20, registrar=True):
        """Aplica uma substituição regex repetidamente, recalculando o texto a cada troca."""
        feitos = 0
        for _ in range(limite):
            t = texto_atual(par)
            m = re.search(padrao, t, flags)
            if not m:
                break
            novo = m.expand(repl) if isinstance(repl, str) else repl(m)
            if novo == m.group(0) or not self.trocar(par, m.start(), m.end(), novo):
                if novo != m.group(0):
                    self.falhas.append(f"{descricao}: {t[:80]}")
                break
            feitos += 1
            if registrar:
                self._reg(grupo, f"{descricao}: “{m.group(0).strip()}” → “{novo.strip()}”", t)
        return feitos

    # ------------------------------------------------------------------ formatação
    def _runs_obj(self, par):
        return [Run(r, par) for r in _runs(par._p)]

    def _tamanho(self, par, pt):
        for run in self._runs_obj(par):
            if _texto_run(run._r).strip():
                run.font.size = Pt(pt)

    # ================================================================== grupos
    def margens(self):
        n = 0
        for s in self.doc.sections:
            larg, alt = du.cm(s.page_width) or 21, du.cm(s.page_height) or 29.7
            paisagem = larg > alt
            alvo = (Cm(29.7), Cm(21)) if paisagem else (Cm(21), Cm(29.7))
            mudou = abs(larg - du.cm(alvo[0])) > 0.1 or abs(alt - du.cm(alvo[1])) > 0.1
            s.page_width, s.page_height = alvo
            med = du.cm(s.gutter) or 0
            atual = ((du.cm(s.left_margin) or 0) + med, du.cm(s.right_margin), abs(du.cm(s.top_margin) or 0),
                     abs(du.cm(s.bottom_margin) or 0))
            M = PF.atual()["margens"]
            esp_m = (M["esquerda"], M["direita"], M["superior"], M["inferior"])
            if mudou or any(abs((v or 0) - e) > 0.1 for v, e in zip(atual, esp_m)):
                s.gutter = Cm(0)
                s.left_margin, s.right_margin, s.top_margin, s.bottom_margin = (Cm(x) for x in esp_m)
                n += 1
        if n:
            self._reg("margens", f"Margens {PF.margens_txt()} e papel A4 aplicados em {n} seção(ões) do Word "
                                 "(medianiz incorporada à margem esquerda)")

    def texto(self):
        P = PF.atual()
        CL, FT, EL = P["citacao_longa"], P["fonte_texto"], P["entrelinhas"]
        LIM = max(1.5, CL["recuo"] - 0.5)
        fonte = esp = longas = 0
        for b in self.d.paragrafos("textual"):
            p, t = b.obj, texto_atual(b.obj).strip()
            if not t or len(t) < 60 or du.nivel_titulo_estilo(p) or _eh_legenda_ou_fonte(t):
                continue
            if re.match(r"^[a-z]\)\s", t) or re.match(r"^[-–—•]\s", t):
                continue
            recuo = du.recuo_esquerdo_cm(p)
            tam = du.tamanho_predominante(p, self.doc)
            if recuo >= LIM and not citacao_longa(p, t, self.doc):
                continue  # recuado sem sinais de citação (lista, etapas...): não mexer, fica para conferência
            if recuo >= LIM:
                full = texto_atual(p)
                ini = len(full) - len(full.lstrip())
                com_aspas = full[ini:ini + 1] in "\"“"
                esp_q = du.entrelinhas(p, self.doc) or 1
                if abs(recuo - CL["recuo"]) <= 0.3 and tam and tam <= CL["fonte_max"] and \
                        esp_q <= CL["entrelinhas"] + 0.15 and not com_aspas:
                    continue
                p.paragraph_format.left_indent = Cm(CL["recuo"])
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.line_spacing = CL["entrelinhas"]
                if not tam or tam > CL["fonte_max"]:
                    self._tamanho(p, CL["fonte"])
                if com_aspas:
                    fim_q = full.rstrip()
                    fecha = max(fim_q.rfind("\""), fim_q.rfind("”"))
                    if fecha > ini and self.trocar(p, fecha, fecha + 1, ""):
                        self.trocar(p, ini, ini + 1, "")
                longas += 1
                continue
            if tam and abs(tam - FT) > 0.1:
                self._tamanho(p, FT)
                fonte += 1
            if abs((du.entrelinhas(p, self.doc) or EL) - EL) > 0.1:
                p.paragraph_format.line_spacing = EL
                esp += 1
        if fonte:
            self._reg("texto", f"Fonte {PF.fmt(FT)} aplicada em {fonte} parágrafo(s) do texto")
        if esp:
            self._reg("texto", f"Entrelinhas {PF.fmt(EL)} aplicado em {esp} parágrafo(s) do texto")
        if longas:
            self._reg("texto", f"{longas} citação(ões) longa(s) padronizada(s): recuo {PF.fmt(CL['recuo'])} cm, "
                               f"fonte {PF.fmt(CL['fonte'])}, entrelinhas {PF.fmt(CL['entrelinhas'])}, sem aspas")

    def titulos(self):
        for b in self.d.paragrafos():
            if self.d.regiao(b.i) == "pre":
                continue
            m = _candidato_titulo(b)
            if not m or m == "auto":
                continue
            p = b.obj
            full = texto_atual(p)
            lead = len(full) - len(full.lstrip())
            mm = RE_SECAO.match(full.strip())
            if not mm:
                continue
            num, ponto, sep, _ = mm.groups()
            if ponto or sep.strip():
                ini, fim = lead + len(num), lead + mm.start(4)
                if self.trocar(p, ini, fim, " "):
                    self._reg("titulos", f"Pontuação removida após o indicativo: “{full.strip()[:40]}”", full)
            if "." not in num and not du.comeca_em_nova_pagina(p, self.d.anterior_paragrafo(b.i)):
                p.paragraph_format.page_break_before = True
                self._reg("titulos", f"Seção primária passa a começar em nova página: “{full.strip()[:40]}”")

        for b in self.d.paragrafos():
            full = texto_atual(b.obj)
            if not full.strip() or "\t" in full or len(full.strip()) > 80:
                continue
            k = chave(full)
            base = re.sub(r"^\d+(\.\d+)*\.?\s*[-–]?\s*", "", k)
            if not (base in TITULOS_SEM_NUMERO or RE_APENDICE.match(base)):
                continue
            if base != k and not re.match(r"^\d", k):
                continue
            if base in ("RESUMO", "ABSTRACT") and self.d.regiao(b.i) != "pre":
                continue
            p = b.obj
            if base != k:
                m = re.match(r"^\s*\d+(\.\d+)*\.?\s*[-–]?\s*", full)
                if self.trocar(p, 0, m.end(), ""):
                    self._reg("titulos", f"Número removido de título sem indicativo: “{full.strip()}”")
            titulo_ref = PF.atual()["referencias"]["titulo"]
            if base in ("REFERENCIAS", "REFERENCIAS BIBLIOGRAFICAS", "BIBLIOGRAFIA") and base != chave(titulo_ref):
                t2 = texto_atual(p)
                ini = len(t2) - len(t2.lstrip())
                if self.trocar(p, ini, len(t2.rstrip()), titulo_ref):
                    self._reg("titulos", f"“{t2.strip()}” → “{titulo_ref}”")
            if du.alinhamento(p) != WD_ALIGN_PARAGRAPH.CENTER:
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Cm(0)
                self._reg("titulos", f"Título centralizado: “{full.strip()}”")

    def resumo(self):
        for b in self.d.paragrafos("pre"):
            t = texto_atual(b.obj)
            m = re.match(r"^(\s*(Palavras[- ]?chaves?|Key[- ]?words?)\s*:)(.*)$", t, re.I | re.S)
            if not m:
                continue
            certo = "Palavras-chave" if m.group(2).lower().startswith("pal") else "Keywords"
            ini_rot = len(m.group(1)) - len(m.group(1).lstrip())
            if m.group(2) != certo and self.trocar(b.obj, ini_rot, ini_rot + len(m.group(2)), certo):
                self._reg("resumo", f"Rótulo '{m.group(2)}' → '{certo}'")
                t = texto_atual(b.obj)
                m = re.match(r"^(\s*(Palavras-chave|Keywords)\s*:)(.*)$", t, re.S)
                if not m:
                    continue
            termos_txt = m.group(3).strip()
            ruim = ";" in termos_txt or (termos_txt.count(",") >= 2 and termos_txt.count(".") <= 1) \
                or (termos_txt and not termos_txt.endswith("."))
            if not ruim or not termos_txt:
                continue
            termos = [x.strip() for x in re.split(r"\s*[;,]\s*|\.\s+", termos_txt.rstrip(".")) if x.strip()]
            novo = " " + ". ".join(x[:1].upper() + x[1:] for x in termos) + "."
            if self.trocar(b.obj, len(m.group(1)), len(t), novo):
                self._reg("resumo", f"{m.group(2)}: termos separados por ponto e finalizados por ponto", novo)

    def ilustracoes(self):
        movidas = 0
        for b in list(self.d.paragrafos()):
            if self.d.regiao(b.i) == "pre":
                continue
            p = b.obj
            full = texto_atual(p)
            lead = len(full) - len(full.lstrip())
            t = full.strip()
            if "\t" in t:
                continue
            m = RE_ILU.match(t)
            tab = None  # tabelas e quadros de tabela: grupo "tabelas"
            if m and m.group(1) == "Quadro":
                ant_q, prox_q = _vizinho(self.d, b.i, -1), _vizinho(self.d, b.i, +1)
                if (prox_q is not None and prox_q.tipo == "t") or (ant_q is not None and ant_q.tipo == "t"):
                    continue
            if m:
                if m.group(3) not in ("-", "–", "—"):
                    if self.trocar(p, lead + m.end(2), lead + m.start(4), " – "):
                        self._reg("ilustracoes", f"Título de ilustração com travessão: “{t[:50]}”")
                ant, prox = _vizinho(self.d, b.i, -1), _vizinho(self.d, b.i, +1)
                img_antes = ant is not None and ant.tipo == "p" and du.tem_imagem(ant.obj)
                img_depois = prox is not None and (prox.tipo == "t" or (prox.tipo == "p" and du.tem_imagem(prox.obj)))
                if img_antes and not img_depois:
                    ant.obj._p.addprevious(p._p)
                    movidas += 1
                    self._reg("ilustracoes", f"Título movido para cima da imagem: “{t[:50]}”")
            elif tab:
                if tab.group(2) not in ("-", "–", "—"):
                    if self.trocar(p, lead + tab.end(1), lead + tab.start(3), " - "):
                        self._reg("ilustracoes", f"Título de tabela com hífen: “{t[:50]}”")
                prox = _vizinho(self.d, b.i, +1)
                if prox is not None and prox.tipo == "t" and self._limpar_bordas(prox.obj):
                    self._reg("ilustracoes", f"Traços verticais removidos da tabela: “{t[:50]}”")
            elif re.match(r"^Fonte\s*:", t) and (du.tamanho_predominante(p, self.doc) or 10) > \
                    PF.atual()["fonte_legenda_max"]:
                self._tamanho(p, min(10, PF.atual()["fonte_legenda_max"]))
                p.paragraph_format.line_spacing = 1.0
                self._reg("ilustracoes", f"Fonte de ilustração/tabela em tamanho 10: “{t[:50]}”")

    def tabelas(self):
        from .tabelas import analisar_tabelas
        for t in analisar_tabelas(self.d):
            if t.natureza not in ("tabela", "quadro") or not t.leg_tipo:
                continue
            leg = t.legenda.obj
            nome = t.legenda.texto[:50]
            if t.leg_pos == "abaixo":
                t.bloco.obj._tbl.addprevious(leg._p)
                self._reg("tabelas", f"Título movido para cima da tabela: “{nome}”")
            sep_certo = " - " if t.leg_tipo == "Tabela" else " – "
            if t.leg_sep not in ("-", "–", "—"):
                full = texto_atual(leg)
                lead = len(full) - len(full.lstrip())
                m = re.match(r"^(\w+\s+\d+(?:\.\d+)?)(\s*[-–—:.]?\s*)", full.strip())
                if m and self.trocar(leg, lead + m.end(1), lead + m.end(2), sep_certo):
                    self._reg("tabelas", f"Título no formato “{t.leg_tipo} N{sep_certo}Título”: “{nome}”")
            if t.natureza == "tabela" and t.leg_tipo == "Tabela" and t.bordas_vert and self._limpar_bordas(t.bloco.obj):
                self._reg("tabelas", f"Traços verticais removidos: “{nome}”")
            if t.leg_tipo == "Quadro" and not t.bordas_fechadas:
                self._bordas_fechadas(t.bloco.obj)
                self._reg("tabelas", f"Bordas fechadas no quadro: “{nome}”")
            if t.linhas > 25 and not t.cab_repete:
                tr0 = t.bloco.obj._tbl.tr_lst[0]
                trpr = tr0.get_or_add_trPr()
                trpr.append(OxmlElement("w:tblHeader"))
                self._reg("tabelas", f"Cabeçalho passa a se repetir em cada página: “{nome}” ({t.linhas} linhas)")
            if t.fonte is not None and (du.tamanho_predominante(t.fonte.obj, self.doc) or 10) > \
                    PF.atual()["fonte_legenda_max"]:
                self._tamanho(t.fonte.obj, min(10, PF.atual()["fonte_legenda_max"]))
                t.fonte.obj.paragraph_format.line_spacing = 1.0
                self._reg("tabelas", f"Fonte da {t.leg_tipo.lower()} em tamanho 10: “{t.fonte.texto[:50]}”")

    def _bordas_fechadas(self, tabela):
        tblpr = tabela._tbl.tblPr
        velho = tblpr.find(qn("w:tblBorders"))
        if velho is not None:
            tblpr.remove(velho)
        bordas = OxmlElement("w:tblBorders")
        for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement(f"w:{lado}")
            for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:space", "0"), ("w:color", "000000")):
                el.set(qn(k), v)
            bordas.append(el)
        tblpr.append(bordas)

    def _limpar_bordas(self, tabela):
        from .estrutura import _tem_bordas_verticais
        if not _tem_bordas_verticais(tabela):
            return False
        tbl = tabela._tbl
        tblpr = tbl.tblPr
        velho = tblpr.find(qn("w:tblBorders"))
        if velho is not None:
            tblpr.remove(velho)
        bordas = OxmlElement("w:tblBorders")
        for lado, val in (("top", "single"), ("left", "nil"), ("bottom", "single"), ("right", "nil"),
                          ("insideH", "nil"), ("insideV", "nil")):
            el = OxmlElement(f"w:{lado}")
            el.set(qn("w:val"), val)
            if val == "single":
                el.set(qn("w:sz"), "8")
                el.set(qn("w:space"), "0")
                el.set(qn("w:color"), "000000")
            bordas.append(el)
        tblpr.append(bordas)
        for tcb in list(tbl.iter(qn("w:tcBorders"))):
            tcb.getparent().remove(tcb)
        # traço abaixo do cabeçalho
        for tc in tabela.rows[0].cells:
            tcpr = tc._tc.get_or_add_tcPr()
            tcb = OxmlElement("w:tcBorders")
            el = OxmlElement("w:bottom")
            for k, v in (("w:val", "single"), ("w:sz", "6"), ("w:space", "0"), ("w:color", "000000")):
                el.set(qn(k), v)
            tcb.append(el)
            tcpr.append(tcb)
        return True

    def alineas(self):
        grupos, atual = [], []
        for b in self.d.paragrafos():
            if self.d.regiao(b.i) == "pre":
                continue
            t = texto_atual(b.obj).strip()
            if re.match(r"^[a-z]\)\s", t):
                atual.append(b)
            elif t:
                if atual:
                    grupos.append(atual)
                atual = []
        if atual:
            grupos.append(atual)
        for g in grupos:
            for k, b in enumerate(g):
                p = b.obj
                full = texto_atual(p)
                lead = len(full) - len(full.lstrip())
                corpo_ini = lead + 3 + (len(full[lead + 3:]) - len(full[lead + 3:].lstrip()))
                c = full[corpo_ini:corpo_ini + 2]
                if c[:1].isupper() and not c[1:2].isupper():
                    if self.trocar(p, corpo_ini, corpo_ini + 1, c[0].lower()):
                        self._reg("alineas", f"Alínea iniciada com minúscula: “{full.strip()[:50]}”")
                full = texto_atual(p).rstrip()
                fim = "." if k == len(g) - 1 else ";"
                if full.endswith(":") or full.endswith(fim):
                    continue
                if full[-1:] in ".;,":
                    ok = self.trocar(p, len(full) - 1, len(full), fim)
                else:
                    ok = self.trocar(p, len(full), len(full), fim)
                if ok:
                    self._reg("alineas", f"Alínea terminada em “{fim}”: “{full.strip()[:50]}”")

    def citacoes(self):
        modo = PF.atual()["citacao_autor"]
        if modo == "qualquer":
            return
        refs = R.extrair_referencias(self.d)
        cits = R.extrair_citacoes(self.d, refs)
        nomes = set()
        for c in cits:
            if modo == "maiusculas":
                if c.parentese and not c.caixa_alta and re.match(r"^[A-ZÀ-Ý][a-zà-ÿ]", c.autor.strip()) and \
                        any(R._casa(c, r) and re.match(r"^[\wÀ-ÿ' \-]+,\s", r.texto) for r in refs):
                    nomes.add(c.autor.strip())
                continue
            if c.caixa_alta:
                pessoa = any(R._casa(c, r) and re.match(r"^[\wÀ-ÿ' \-]+,\s", r.texto) for r in refs)
                if pessoa or len(c.autor.strip()) >= 8:
                    nomes.add(c.autor.strip())
        for nome in nomes:
            padrao = rf"(?<=[(;])\s*({re.escape(nome)})(?=,\s*(?:1[5-9]\d{{2}}|20\d{{2}}|\[)|;\s*[A-ZÀ-Ý])"
            for b in self.d.texto_do_corpo():
                for _ in range(20):
                    t = texto_atual(b.obj)
                    m = re.search(padrao, t)
                    if not m:
                        break
                    novo = m.group(1).upper() if modo == "maiusculas" else _titulo_nome(m.group(1))
                    if not self.trocar(b.obj, m.start(1), m.end(1), novo):
                        self.falhas.append(f"citação {nome}: {t[:60]}")
                        break
                    self._reg("citacoes", f"({m.group(1)}, …) → ({novo}, …)", t)

    def aplicar_ia(self, sugestoes):
        """sugestoes: {texto_original_da_referência: sugestão com **negrito**}"""
        if not sugestoes:
            return
        for grupo in self.d.referencias_agrupadas():
            orig = re.sub(r"\s+", " ", " ".join(b.texto for b in grupo)).strip()
            sug = sugestoes.get(orig) or (sugestoes.get(grupo[0].texto) if len(grupo) == 1 else None)
            if not sug:
                continue
            b = grupo[0]
            segs = [(s, i % 2 == 1) for i, s in enumerate(re.split(r"\*\*", sug.strip()))]
            full = texto_atual(b.obj)
            if self.trocar(b.obj, 0, len(full), segs):
                for resto in grupo[1:]:  # o restante da referência quebrada já está na sugestão
                    self.trocar(resto.obj, 0, len(texto_atual(resto.obj)), "")
                    self._ref_ia.add(id(resto.obj._p))
                self._ref_ia.add(id(b.obj._p))
                self._reg("referencias", "Referência substituída pela sugestão da IA (conferir)", sug.replace("**", ""))
            else:
                self.falhas.append(f"sugestão da IA não aplicada: {orig[:60]}")

    def _juntar(self, grupo):
        """Une os parágrafos de uma referência quebrada no primeiro (com marca de revisão, se rastreando)."""
        alvo = grupo[0].obj
        for b in grupo[1:]:
            fim = texto_atual(alvo)
            prox = texto_atual(b.obj)
            if not prox.strip():
                continue
            if fim and not fim.endswith(" ") and not prox.startswith(" "):
                self.trocar(alvo, len(fim), len(fim), " ")
            if self.rastrear:
                # move o conteúdo como inserção e apaga o original, preservando a formatação dos runs
                for r in _runs(b.obj._p):
                    novo = deepcopy(r)
                    ins = self._marca("w:ins")
                    ins.append(novo)
                    alvo._p.append(ins)
                for r in _runs(b.obj._p):  # apaga com marca de revisão (também dentro de hiperlinks)
                    if r.getparent().tag == qn("w:ins"):
                        r.getparent().remove(r)  # inserção feita agora pelo próprio verificador
                        continue
                    dele = self._marca("w:del")
                    r.addprevious(dele)
                    dele.append(r)
                    for t in r.findall(qn("w:t")):
                        t.tag = qn("w:delText")
                ppr = b.obj._p.get_or_add_pPr()
                rpr = ppr.find(qn("w:rPr"))
                if rpr is None:
                    rpr = OxmlElement("w:rPr")
                    ppr.append(rpr)
                rpr.append(self._marca("w:del"))  # marca de parágrafo apagada: Word junta ao aceitar
            else:
                for r in _runs(b.obj._p):
                    alvo._p.append(r)
                b.obj._p.getparent().remove(b.obj._p)
        return len(grupo) > 1

    def referencias(self):
        grupos = self.d.referencias_agrupadas()
        if not grupos:
            return
        numerico = sum(1 for g in grupos if re.match(r"^\[?\d+[\].]\s", g[0].texto)) > len(grupos) / 2
        g = "referencias"
        for grupo in grupos:
            primeiro, ultimo = grupo[0].obj, grupo[-1].obj
            if id(primeiro._p) not in self._ref_ia:
                t = texto_atual(primeiro)
                corpo_ini = re.match(r"^\s*(\[?\d+[\].]\s*)?", t).end() if numerico else 0
                m = re.match(r"^([\wÀ-ÿ' \-]+),", t[corpo_ini:])
                if m and m.group(1) != m.group(1).upper() and len(m.group(1).split()) <= 4:
                    if self.trocar(primeiro, corpo_ini, corpo_ini + len(m.group(1)), m.group(1).upper()):
                        self._reg(g, f"Sobrenome de entrada em maiúsculas: “{m.group(1)}”", t)
                for b in grupo:
                    p = b.obj
                    self.trocar_regex(p, r"(\d+)\s*[ªºa°]\.?\s*(?:edi[çc][ãa]o|ed)\b\.?", r"\1. ed.", g,
                                      "Edição", re.I)
                    self.trocar_regex(p, r"(\d+)\.\s*edi[çc][ãa]o\.?", r"\1. ed.", g, "Edição", re.I)
                    self.trocar_regex(p, r"\bpp\.\s*", "p. ", g, "Páginas")
                    self.trocar_regex(p, r"\bp\.(\d)", r"p. \1", g, "Páginas")
                    self.trocar_regex(p, r"\b(?:IN|in):\s", "In: ", g, "In:")
                    self.trocar_regex(p, r"(Dispon[ií]vel em)(?!:)\s*", r"\1: ", g, "Disponível em")
                    self.trocar_regex(p, r"(Acesso em)(?!:)\s*", r"\1: ", g, "Acesso em")
                    self.trocar_regex(
                        p, r"(?<=Acesso em: )(\d{1,2})\s+(?:de\s+)?(janeiro|fevereiro|março|marco|abril|maio\.|mai\.|junho|"
                           r"julho|agosto|setembro|outubro|novembro|dezembro)\s+(?:de\s+)?",
                        lambda mm: f"{mm.group(1)} {MESES[mm.group(2).lower()]} ", g, "Mês abreviado", re.I)
                t = texto_atual(ultimo).rstrip()
                if t and not t.endswith((".", ".]")) and not re.search(r"https?://\S+>?$", t):
                    if self.trocar(ultimo, len(t), len(t), "."):
                        self._reg(g, "Ponto final incluído", t)
            for b in grupo:
                p = b.obj
                if abs((du.entrelinhas(p, self.doc) or 1) - 1) > 0.15:
                    p.paragraph_format.line_spacing = 1.0
                    self._reg(g, "Espaço simples", texto_atual(p))
                if du.alinhamento(p) == WD_ALIGN_PARAGRAPH.JUSTIFY:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    self._reg(g, "Alinhada à esquerda", texto_atual(p))
                p.paragraph_format.first_line_indent = Cm(0)
                p.paragraph_format.left_indent = Cm(0)
        if not numerico:
            self._ordenar_e_espacar(grupos)
        juntas = 0
        for grupo in grupos:
            try:
                juntas += self._juntar(grupo)
            except Exception as e:
                self.falhas.append(f"referência quebrada não unida: {grupo[0].texto[:60]} ({e})")
        if juntas:
            self._reg(g, f"{juntas} referência(s) quebrada(s) em várias linhas unida(s) num só parágrafo")

    def _ordenar_e_espacar(self, grupos):
        ini, fim = self.d.i_refs + 1, self.d.i_fim_refs
        regiao = self.d.blocos[ini:fim]
        if any(b.tipo != "p" or (b.obj._p.pPr is not None and b.obj._p.pPr.sectPr is not None) for b in regiao):
            return  # tabela ou quebra de seção no meio da lista: não mexe na ordem
        textos = [" ".join(texto_atual(b.obj) for b in gr) for gr in grupos]
        # ordem pelo elemento de entrada; o mesmo autor mantém a ordem em que o aluno colocou (sort estável)
        ordenadas = sorted(zip([R.autor_chave(t) for t in textos], range(len(grupos)), grupos),
                           key=lambda x: (x[0], x[1]))
        mudou_ordem = [k for _, k, _ in ordenadas] != list(range(len(grupos)))
        dentro = {id(b.obj._p) for gr in grupos for b in gr[1:]}  # continuação de referência quebrada
        cheio = [bool(texto_atual(b.obj).strip()) for b in regiao if id(b.obj._p) not in dentro]
        while cheio and not cheio[-1]:
            cheio.pop()  # linhas em branco no fim da lista não contam
        # entre duas referências deve haver exatamente uma linha em branco
        precisa_espaco = any(a and b for a, b in zip(cheio, cheio[1:])) or any(
            not a and not b for a, b in zip(cheio, cheio[1:]))
        if not (mudou_ordem or precisa_espaco):
            return
        entradas = [gr[0] for gr in grupos]
        modelo_vazio = deepcopy(entradas[0].obj._p)
        for r in list(modelo_vazio):
            if r.tag != qn("w:pPr"):
                modelo_vazio.remove(r)
        ancora = self.d.blocos[self.d.i_refs].obj._p
        for b in regiao:
            b.obj._p.getparent().remove(b.obj._p)
        for k, (_, _, gr) in enumerate(ordenadas):
            for b in gr:  # a referência quebrada se move inteira
                ancora.addnext(b.obj._p)
                ancora = b.obj._p
            if k < len(ordenadas) - 1:
                v = deepcopy(modelo_vazio)
                ancora.addnext(v)
                ancora = v
        if mudou_ordem:
            self._reg("referencias", "Lista reordenada em ordem alfabética")
        if precisa_espaco:
            self._reg("referencias", "Uma linha em branco entre as referências")

    # ================================================================== execução
    def tipografia(self):
        from .tipografia import RE_ESP_DUPLO, RE_ESP_PONT, _titulo_flex
        g = "tipografia"
        duplos = pont = 0
        for b in self.d.paragrafos("textual") + self.d.paragrafos("pos"):
            p = b.obj
            t = texto_atual(p)
            if not t.strip() or "\t" in p.text:
                continue
            duplos += self.trocar_regex(p, RE_ESP_DUPLO.pattern, " ", g, "Espaço duplo", limite=50, registrar=False)
            pont += self.trocar_regex(p, RE_ESP_PONT.pattern, "", g, "Espaço antes da pontuação", limite=50,
                                      registrar=False)
            pont += self.trocar_regex(p, r"\( +", "(", g, "Espaço depois de '('", limite=20, registrar=False)
            pont += self.trocar_regex(p, r"(?<=\S) +\)", ")", g, "Espaço antes de ')'", limite=20, registrar=False)
        if duplos:
            self._reg(g, f"{duplos} espaço(s) duplo(s) trocado(s) por um espaço")
        if pont:
            self._reg(g, f"{pont} espaço(s) antes da pontuação ou junto aos parênteses removido(s)")
        for b in self.d.paragrafos():
            if self.d.regiao(b.i) == "pre" or not _titulo_flex(self.d, b):
                continue
            p = b.obj
            t = texto_atual(p)
            m = re.search(r"\s*[.:;]\s*$", t)
            if m and not t.rstrip().endswith("..."):
                if self.trocar(p, m.start(), len(t), ""):
                    self._reg(g, "Ponto final retirado do título", t)
            t = texto_atual(p)
            m = re.match(r"^(\d{1,2}(?:\.\d{1,2}){0,6}\.?)(\t|\s{2,})", t)
            if m:
                if self.trocar(p, m.end(1), m.end(), " "):
                    self._reg(g, "Um espaço entre o número e o título", t)
            if du.alinhamento(p) == WD_ALIGN_PARAGRAPH.CENTER:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                self._reg(g, "Título de seção alinhado à esquerda", t)

    def executar(self, sugestoes_ia=None):
        etapas = [("margens", self.margens), ("texto", self.texto), ("titulos", self.titulos),
                  ("resumo", self.resumo), ("ilustracoes", self.ilustracoes), ("tabelas", self.tabelas),
                  ("alineas", self.alineas),
                  ("citacoes", self.citacoes), ("tipografia", self.tipografia)]
        for nome, f in etapas:
            if nome in self.grupos:
                try:
                    f()
                except Exception as e:
                    self.falhas.append(f"{GRUPOS[nome]}: etapa interrompida ({e})")
        if "referencias" in self.grupos:
            try:
                self.aplicar_ia(sugestoes_ia)
                self.referencias()
            except Exception as e:
                self.falhas.append(f"{GRUPOS['referencias']}: etapa interrompida ({e})")
        buf = io.BytesIO()
        self.doc.save(buf)
        return buf.getvalue()


def corrigir(arquivo_bytes, grupos=None, rastrear=True, sugestoes_ia=None, perfil_norma=None):
    """Retorna (bytes do .docx corrigido, log de alterações, falhas)."""
    with PF.usar(perfil_norma or PF.atual()):
        c = Corretor(io.BytesIO(arquivo_bytes), rastrear=rastrear, grupos=grupos)
        dados = c.executar(sugestoes_ia)
    return dados, c.log, c.falhas


# ====================================================================== revisão dupla
# palavras que a correção automática pode legitimamente tirar ou pôr (todas as outras devem ser preservadas)
_PODE_SAIR = {"edição", "edicao", "ed", "pp", "pág", "pag", "de", "bibliográficas", "bibliograficas", "chaves",
              "key", "words", "a", "ª", "º", "janeiro", "fevereiro", "março", "marco", "abril", "junho", "julho",
              "agosto", "setembro", "outubro", "novembro", "dezembro", "mai", "in"}
_PODE_ENTRAR = {"ed", "p", "in", "chave", "keywords", "jan", "fev", "mar", "abr", "maio", "jun", "jul", "ago",
                "set", "out", "nov", "dez", "fonte", "elaborado", "pelo", "autor", "referências", "referencias"}


def _palavras(dados, sem_refs=False):
    from collections import Counter
    from .edicao import aceitar_tudo
    d = Documento(io.BytesIO(aceitar_tudo(dados)))
    cont = Counter()
    for b in d.texto_do_corpo() if sem_refs else [*d.texto_do_corpo(), *d.entradas_referencias()]:
        for w in re.findall(r"[\wÀ-ÿ]+", b.texto.lower()):
            w = re.sub(r"\d", "", w)  # "3a", "2ª" → ordinal sem o número
            if w:
                cont[w] += 1
    return cont


def integridade(antes, depois, com_ia=False):
    """Confere que a correção não apagou nem inventou texto do aluno. Devolve (ok, problemas)."""
    a, b = _palavras(antes, com_ia), _palavras(depois, com_ia)
    extra = set(re.findall(r"[\wÀ-ÿ]+", PF.atual()["referencias"]["titulo"].lower()))
    sumiram = {w: n for w, n in (a - b).items() if w not in _PODE_SAIR | {"referências", "referencias", "bibliografia"}}
    surgiram = {w: n for w, n in (b - a).items() if w not in _PODE_ENTRAR | extra}
    probs = []
    if sumiram:
        probs.append("palavras que sumiram: " + ", ".join(f"{w} ({n}×)" for w, n in list(sumiram.items())[:8]))
    if surgiram:
        probs.append("palavras que surgiram: " + ", ".join(f"{w} ({n}×)" for w, n in list(surgiram.items())[:8]))
    return not probs, probs


def _pendencias(dados):
    from . import analisar
    from .edicao import aceitar_tudo
    from .modelo import OK
    achados, _ = analisar(io.BytesIO(aceitar_tudo(dados)), PF.atual())
    return sum(1 for x in achados if x.status != OK and x.categoria != "Sistema")


def corrigir_duplo(arquivo_bytes, grupos=None, rastrear=True, sugestoes_ia=None, perfil_norma=None):
    with PF.usar(perfil_norma):
        return _corrigir_duplo(arquivo_bytes, grupos, rastrear, sugestoes_ia)


def _corrigir_duplo(arquivo_bytes, grupos=None, rastrear=True, sugestoes_ia=None):
    """Revisão dupla: 1ª passada → nova análise → 2ª passada (pega o que só aparece depois da 1ª, como a
    ordem das referências depois de juntar linhas quebradas) → nova análise.

    Cada passada só é mantida se (1) não apagar nem inventar palavras do aluno e (2) não aumentar as
    pendências. Retorna (bytes, log, falhas, passadas) — 'passadas' descreve o que aconteceu em cada uma.
    """
    passadas = []
    atual = arquivo_bytes
    pend = _pendencias(atual)
    log_total, falhas_total = [], []
    for n in (1, 2):
        try:
            novo, log, falhas = corrigir(atual, grupos, rastrear, sugestoes_ia if n == 1 else None)
        except Exception as e:
            passadas.append(dict(passada=n, correcoes=0, antes=pend, depois=pend, integra=True,
                                 resultado=f"interrompida: {e}"))
            break
        if not log:
            passadas.append(dict(passada=n, correcoes=0, antes=pend, depois=pend, integra=True,
                                 resultado="nada mais a corrigir"))
            break
        ok, probs = integridade(atual, novo, com_ia=bool(sugestoes_ia) and n == 1)
        depois = _pendencias(novo)
        if not ok:
            passadas.append(dict(passada=n, correcoes=len(log), antes=pend, depois=pend, integra=False,
                                 resultado="desfeita: " + "; ".join(probs)))
            falhas_total.append(f"{n}ª passada desfeita por segurança ({'; '.join(probs)})")
            break
        if depois > pend:
            passadas.append(dict(passada=n, correcoes=len(log), antes=pend, depois=depois, integra=True,
                                 resultado="desfeita: aumentaria as pendências"))
            falhas_total.append(f"{n}ª passada desfeita: as pendências subiriam de {pend} para {depois}")
            break
        for x in log:
            x["Passada"] = f"{n}ª"
        log_total += log
        falhas_total += [f for f in falhas if f not in falhas_total]
        passadas.append(dict(passada=n, correcoes=len(log), antes=pend, depois=depois, integra=True,
                             resultado="mantida"))
        atual, pend = novo, depois
    return atual, log_total, falhas_total, passadas
