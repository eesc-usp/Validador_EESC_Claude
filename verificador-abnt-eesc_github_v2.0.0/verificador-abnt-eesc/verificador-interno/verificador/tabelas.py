"""Tabelas e quadros: identificação e verificação (Diretrizes USP 2.6 e 2.7; Normas de apresentação tabular do IBGE).

Cada tabela do Word é classificada como:
- "tabela": dados (predominantemente numéricos) → título "Tabela N - ...", sem traços verticais;
- "quadro": informação textual → é uma ilustração: "Quadro N – ...", bordas fechadas;
- "diagramação": tabela usada só para montar a página (capa, folha de aprovação, banca, assinatura).
"""
import re
from dataclasses import dataclass, field

from . import docx_util as du
from .cache import qn
from .documento import chave
from .modelo import AJUSTAR, OK, VERIFICAR, Achado

CAT = "Tabelas e quadros"
RE_LEG = re.compile(r"^(Tabela|Quadro|Figura|Gr[áa]fico|Fluxograma|Esquema)\s+(\d+(?:\.\d+)?)\s*([-–—:.]?)\s*(.*)$")
RE_NUM = re.compile(r"^[\s\d.,;:%‰()+\-–−±×x/<>=≤≥*R$€US]+$")
LINHAS_LONGA = 25


def _texto_tc(tc):
    from .correcao import texto_atual
    return "\n".join(texto_atual(p) for p in tc.iter(qn("w:p"))).strip()


@dataclass
class InfoTabela:
    bloco: object
    i: int
    regiao: str
    celulas: list                  # [[texto, ...], ...] (células únicas por linha)
    linhas: int
    colunas: int
    legenda: object = None         # Bloco do título (Tabela/Quadro N - ...)
    leg_tipo: str = ""             # "Tabela", "Quadro", ...
    leg_num: str = ""
    leg_sep: str = ""
    leg_pos: str = ""              # "acima" / "abaixo"
    fonte: object = None
    frac_num: float = 0.0
    cab_repete: bool = False
    cab_destacado: bool = False
    mescladas: int = 0
    bordas_vert: bool = False
    bordas_fechadas: bool = False
    natureza: str = ""             # tabela / quadro / diagramação
    citada: bool = True
    motivo_layout: str = ""
    problemas: list = field(default_factory=list)

    @property
    def nome(self):
        if self.leg_tipo:
            return f"{self.leg_tipo} {self.leg_num}"
        return "Tabela sem título" if self.natureza != "diagramação" else "Tabela de diagramação"

    @property
    def amostra(self):
        cel = [c for linha in self.celulas[:2] for c in linha if c][:3]
        return " | ".join(x[:30] for x in cel)


def _bordas(tabela):
    from .estrutura import _tem_bordas_verticais
    vert = _tem_bordas_verticais(tabela)
    tblpr = tabela._tbl.tblPr
    b = tblpr.find(qn("w:tblBorders")) if tblpr is not None else None

    def on(lado):
        if b is None:
            return None
        el = b.find(qn(f"w:{lado}"))
        return el is not None and el.get(qn("w:val")) not in ("nil", "none")
    nome = (tabela.style.name if tabela.style is not None else "") or ""
    grade = bool(re.search(r"grid|grade", nome, re.I))
    fechada = all(on(x) for x in ("top", "bottom", "left", "right", "insideH", "insideV")) if b is not None else grade
    return vert, fechada


def _vizinho_txt(d, i, passo, limite=2):
    j, n = i + passo, 0
    while 0 <= j < len(d.blocos) and n < limite:
        b = d.blocos[j]
        if b.tipo == "t":
            return None
        if b.texto:
            return b
        j += passo
        n += 1
    return None


def analisar_tabelas(d):
    infos = []
    corpo_txt = [(b.i, b.texto) for b in d.blocos if b.tipo == "p" and b.texto]
    for b in d.blocos:
        if b.tipo != "t":
            continue
        tbl = b.obj._tbl
        cel, mescl, max_cols = [], 0, 0
        for tr in tbl.tr_lst:
            linha = []
            for tc in tr.tc_lst:
                tcpr = tc.tcPr
                if tcpr is not None:
                    if tcpr.find(qn("w:gridSpan")) is not None:
                        mescl += 1
                    vm = tcpr.find(qn("w:vMerge"))
                    if vm is not None and vm.get(qn("w:val")) != "restart":
                        mescl += 1
                        linha.append("")
                        continue
                linha.append(_texto_tc(tc))
            max_cols = max(max_cols, len(linha))
            cel.append(linha)
        info = InfoTabela(b, b.i, d.regiao(b.i), cel, len(cel), max_cols, mescladas=mescl)

        # título acima (preferido) ou abaixo
        for pos, viz in (("acima", _vizinho_txt(d, b.i, -1)), ("abaixo", _vizinho_txt(d, b.i, +1))):
            if viz is not None:
                m = RE_LEG.match(viz.texto)
                if m and "\t" not in viz.bruto:
                    info.legenda, info.leg_tipo, info.leg_num, info.leg_sep = viz, m.group(1), m.group(2), m.group(3)
                    info.leg_pos = pos
                    break
        # fonte logo abaixo
        for x in d.blocos[b.i + 1:b.i + 4]:
            if x.tipo == "p" and re.match(r"^Fonte\b", x.texto):
                info.fonte = x
                break
        # conteúdo: fração de células numéricas (sem a primeira linha e a primeira coluna)
        dados = [c for linha in cel[1:] for c in linha[1:] if c.strip()]
        info.frac_num = (sum(1 for c in dados if RE_NUM.match(c)) / len(dados)) if dados else 0.0
        # cabeçalho
        tr0 = tbl.tr_lst[0] if tbl.tr_lst else None
        if tr0 is not None and tr0.trPr is not None and tr0.trPr.find(qn("w:tblHeader")) is not None:
            info.cab_repete = True
        try:
            primeira = b.obj.rows[0].cells
            info.cab_destacado = all(du.todo_negrito(p) for c in primeira for p in c.paragraphs if p.text.strip())
        except Exception:
            pass
        info.bordas_vert, info.bordas_fechadas = _bordas(b.obj)

        # natureza
        texto_tudo = chave(" ".join(c for linha in cel for c in linha))
        if info.leg_tipo == "Tabela":
            info.natureza = "tabela"
        elif info.leg_tipo:
            info.natureza = "quadro" if info.leg_tipo == "Quadro" else "ilustração"
        elif info.regiao == "pre":
            info.natureza, info.motivo_layout = "diagramação", "elemento pré-textual (capa, folha de aprovação...)"
        elif info.linhas * max(info.colunas, 1) <= 2:
            info.natureza, info.motivo_layout = "diagramação", "uma ou duas células"
        elif re.search(r"\b(BANCA|ASSINATURA|APROVAD|JULGAMENTO|ORIENTADOR)\b", texto_tudo):
            info.natureza, info.motivo_layout = "diagramação", "banca/assinaturas"
        else:
            info.natureza = "tabela" if info.frac_num >= 0.5 else "quadro"

        # mencionada no texto?
        if info.leg_tipo and info.leg_num:
            padrao = re.compile(rf"\b{info.leg_tipo[:3]}\w*\.?\s+{re.escape(info.leg_num)}\b", re.I)
            excluir = {info.legenda.i if info.legenda else -1}
            info.citada = any(padrao.search(t) for i, t in corpo_txt
                              if i not in excluir and d.regiao(i) != "pre" and "\t" not in t)
        infos.append(info)
    return infos


def verificar(d):
    infos = analisar_tabelas(d)
    reais = [t for t in infos if t.natureza in ("tabela", "quadro")]
    out = []
    grupos = {k: [] for k in ("sem_titulo", "abaixo", "formato", "fonte", "vert", "quadro_aberto", "natureza",
                              "cabecalho", "citada", "sequencia")}

    def reg(chave_g, t, texto):
        grupos[chave_g].append((t, texto))

    ultimo = {}
    for t in reais:
        if not t.leg_tipo:
            reg("sem_titulo", t, f"{'Tabela' if t.natureza == 'tabela' else 'Quadro'} sem título: {t.amostra}")
            continue
        if t.leg_pos == "abaixo":
            reg("abaixo", t, t.legenda.texto)
        if t.leg_sep not in ("-", "–", "—"):
            reg("formato", t, t.legenda.texto)
        elif t.leg_tipo == "Quadro" and t.leg_sep == "-":
            pass  # hífen em quadro: tolerado
        if not t.fonte:
            reg("fonte", t, t.legenda.texto)
        if t.natureza == "tabela" and t.bordas_vert:
            reg("vert", t, t.legenda.texto)
        if t.leg_tipo == "Quadro" and not t.bordas_fechadas:
            reg("quadro_aberto", t, t.legenda.texto)
        if t.leg_tipo == "Tabela" and t.frac_num < 0.3 and len(t.celulas) >= 2 and t.colunas >= 2:
            reg("natureza", t, f"{t.legenda.texto} — conteúdo é texto: é um Quadro")
        if t.leg_tipo == "Quadro" and t.frac_num >= 0.7:
            reg("natureza", t, f"{t.legenda.texto} — conteúdo numérico: é uma Tabela")
        if t.linhas > LINHAS_LONGA and not t.cab_repete:
            reg("cabecalho", t, f"{t.legenda.texto} ({t.linhas} linhas)")
        if not t.citada:
            reg("citada", t, t.legenda.texto)
        k = t.leg_tipo
        n = int(t.leg_num.split(".")[-1])
        if k in ultimo and n not in (ultimo[k] + 1, 1):
            reg("sequencia", t, f"{k} {ultimo[k]} → {t.leg_num}")
        ultimo[k] = n

    def add(chave_g, item, corr, sec, status=AJUSTAR):
        lista = grupos[chave_g]
        out.append(Achado(CAT, item, status if lista else OK,
                          f"{len(lista)} ocorrência(s)" if lista else "sem problemas",
                          corr if lista else "", sec, [x for _, x in lista],
                          blocos=[(t.legenda.i if (t.legenda and chave_g not in ("sem_titulo", "vert",
                                                                                   "quadro_aberto", "cabecalho"))
                                   else t.i) for t, _ in lista]))

    n_tab = sum(1 for t in reais if t.natureza == "tabela")
    n_qua = sum(1 for t in reais if t.natureza == "quadro")
    n_lay = sum(1 for t in infos if t.natureza == "diagramação")
    resumo = f"{n_tab} tabela(s), {n_qua} quadro(s)" + (f", {n_lay} de diagramação (ignoradas)" if n_lay else "")
    out.append(Achado(CAT, "Tabelas e quadros identificados", OK if infos else OK,
                      resumo if infos else "nenhuma tabela no documento", "", "2.7",
                      tabela=[{"Identificação": t.nome, "Tipo": t.natureza,
                               "Tamanho": f"{t.linhas}×{t.colunas}",
                               "Dados numéricos": f"{round(t.frac_num * 100)}%",
                               "Título": t.leg_pos or "—", "Fonte": "sim" if t.fonte else "não"}
                              for t in infos] if infos else []))
    if not reais:
        return out, infos

    add("sem_titulo", "Tabela/quadro com título", "Toda tabela e todo quadro têm título acima: "
        "'Tabela 1 - Título' (dados) ou 'Quadro 1 – Título' (texto). Se for só diagramação da página, ignore.",
        "2.7")
    add("abaixo", "Título acima da tabela/quadro", "Mover o título para cima; abaixo ficam fonte e notas.", "2.7")
    add("formato", "Formato do título", "'Tabela 1 - Título' (hífen) e 'Quadro 1 – Título' (travessão), "
        "sem dois-pontos ou ponto.", "2.7")
    add("fonte", "Fonte abaixo da tabela/quadro", "Indicar a fonte logo abaixo, em tamanho 10 ou 11 "
        "(ex.: 'Fonte: elaborada pelo autor.').", "2.7")
    add("vert", "Tabelas sem traços verticais", "Tabela (IBGE): sem traços verticais nem laterais; traços "
        "horizontais no cabeçalho e no fechamento.", "2.7")
    add("quadro_aberto", "Quadros com bordas fechadas", "Quadro é ilustração com informação textual e "
        "costuma ter as bordas fechadas.", "2.6", VERIFICAR)
    add("natureza", "Tabela × quadro (tipo coerente com o conteúdo)", "Tabela apresenta dados (principalmente "
        "numéricos); quadro apresenta texto. Ajuste o nome e o estilo de bordas.", "2.7", VERIFICAR)
    add("cabecalho", "Cabeçalho repetido em tabela longa", "Em tabela que continua em outra folha: repetir o "
        "cabeçalho (Propriedades da tabela > Linha > 'Repetir como linha de cabeçalho') e indicar "
        "'continua', 'continuação' e 'conclusão'.", "2.7", VERIFICAR)
    add("citada", "Tabela/quadro mencionado no texto", "Mencionar no texto (ex.: 'conforme a Tabela 2') e "
        "inserir próximo ao trecho em que é citado.", "2.7", VERIFICAR)
    add("sequencia", "Numeração sequencial de tabelas e quadros", "Numerar tabelas e quadros em sequência, "
        "cada tipo com a sua numeração.", "2.7")

    # listas recomendadas a partir de 5 itens
    titulos = {chave(b.texto) for b in d.blocos if b.tipo == "p" and b.texto and "\t" not in b.texto}
    if n_tab >= 5 and "LISTA DE TABELAS" not in titulos:
        out.append(Achado(CAT, "Lista de tabelas", VERIFICAR, f"{n_tab} tabelas e nenhuma 'LISTA DE TABELAS'",
                          "Recomenda-se elaborar a lista a partir de 5 tabelas.", "3.1.2.1.10"))
    if n_qua >= 5 and not ({"LISTA DE QUADROS", "LISTA DE ILUSTRACOES"} & titulos):
        out.append(Achado(CAT, "Lista de quadros", VERIFICAR, f"{n_qua} quadros e nenhuma lista de quadros",
                          "Recomenda-se elaborar a lista a partir de 5 itens.", "3.1.2.1.10"))
    return out, infos
