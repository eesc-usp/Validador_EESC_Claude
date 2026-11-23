"""Original com apontamentos: o .docx do aluno como ele enviou, com um comentário do Word em cada ponto a
revisar (o que está errado, como corrigir e a seção das Diretrizes) e o trecho destacado em amarelo.

O texto do aluno não é alterado: só entram os comentários e o realce. Quem corrige é o próprio aluno.
"""
import io
from copy import deepcopy

from docx.enum.text import WD_COLOR_INDEX
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from .cache import qn
from .documento import Documento
from .modelo import AJUSTAR, OK, VERIFICAR
from .texto import runs_visiveis as _runs
from .texto import texto_atual
from .texto import texto_run as _texto_run
from .visualizacao import serializar

AUTOR = "Biblioteca EESC (verificador)"
INICIAIS = "BE"
MAX_POR_ITEM = 25  # evita encher o documento quando um problema pontual se repete muito
# problemas de formatação que costumam valer para o texto todo: 1 comentário só, na 1ª ocorrência
COLETIVOS = {
    "Entrelinhas": "Dica: corrija tudo de uma vez pelo estilo 'Normal' (Página Inicial > Estilos > "
                   "Normal > Modificar > Formatar > Parágrafo > Entre linhas).",
    "Fonte do texto": "Dica: corrija tudo de uma vez pelo estilo 'Normal' (Página Inicial > Estilos > "
                      "Normal > Modificar > tamanho da fonte).",
    "Citações com mais de 3 linhas": "Dica: crie um estilo 'Citação longa' (recuo esquerdo, fonte menor, "
                                     "espaço simples) e aplique a todas.",
    "Título em destaque (negrito)": "Confira cada referência: o título da obra (ou o nome do periódico) vai em negrito.",
    "Referências alinhadas à esquerda": "Dica: selecione toda a lista de referências e use 'Alinhar à esquerda' (Ctrl+L).",
    "Espaço simples nas referências": "Dica: selecione toda a lista e aplique espaçamento simples entre linhas.",
    "Seções primárias em nova página": "Dica: no estilo 'Título 1', ative 'Quebra de página antes' (Formatar > "
                                       "Parágrafo > Quebras de linha e de página).",
}
LIMITE_COLETIVO = 6
ROTULO = {AJUSTAR: "AJUSTAR", VERIFICAR: "CONFERIR"}


def _simples(r, p):
    return r.getparent() is p and all(ch.tag in (qn("w:rPr"), qn("w:t")) for ch in r)


def _set_texto(r, s):
    for t in r.findall(qn("w:t")):
        r.remove(t)
    t = OxmlElement("w:t")
    t.set(qn("xml:space"), "preserve")
    t.text = s
    r.append(t)


def _isolar(p_el, ini, fim):
    """Divide os runs para que [ini, fim) fique em runs próprios. Devolve esses runs (ou [] se não der)."""
    for corte in (fim, ini):
        pos = 0
        for r in _runs(p_el):
            s = _texto_run(r)
            a, b = pos, pos + len(s)
            pos = b
            if a < corte < b:
                if not _simples(r, p_el):
                    return []
                r2 = deepcopy(r)
                _set_texto(r, s[:corte - a])
                _set_texto(r2, s[corte - a:])
                r.addnext(r2)
                break
    alvo, pos = [], 0
    for r in _runs(p_el):
        s = _texto_run(r)
        if pos >= ini and pos + len(s) <= fim and s:
            alvo.append(r)
        pos += len(s)
    return alvo


def _paragrafo_do_bloco(d, b):
    """Parágrafo onde ancorar o comentário: o próprio parágrafo, ou o 1º parágrafo com texto da tabela."""
    bloco = d.blocos[b]
    if bloco.tipo == "p":
        return bloco.obj
    for p_el in bloco.obj._tbl.iter(qn("w:p")):
        if texto_atual(p_el).strip():
            return Paragraph(p_el, bloco.obj)
    return None


def _texto_comentario(a):
    linhas = [f"{ROTULO.get(a['status'], a['status'])}: {a['item'] if a['origem'] != 'ia' else 'Referência (revisão por IA)'}"]
    if a["detalhe"]:
        linhas.append(f"Encontrado: {a['detalhe']}")
    if a["corrigir"]:
        rot = "Sugestão da IA (conferir)" if a["origem"] == "ia" else "Como corrigir"
        linhas.append(f"{rot}: {a['corrigir'].replace('**', '')}")
    if a["tabela"]:
        cols = list(a["tabela"][0].keys())
        for reg in a["tabela"][:6]:
            linhas.append(f"{reg[cols[0]]}: " + " | ".join(str(reg[c]) for c in cols[1:]))
    if a["secao"]:
        linhas.append(f"Ver Diretrizes USP (ABNT), seção {a['secao']}.")
    return linhas


def _comentar(doc, runs_el, par, linhas):
    runs = [Run(r, par) for r in runs_el]
    c = doc.add_comment(runs, text=linhas[0], author=AUTOR, initials=INICIAIS)
    for linha in linhas[1:]:
        c.add_paragraph(linha)
    return c


def gerar(dados, achados):
    """dados: bytes do .docx ORIGINAL; achados: lista de Achado da análise desse mesmo arquivo.

    Retorna (bytes do .docx anotado, número de comentários).
    """
    d = Documento(io.BytesIO(dados))
    doc = d.doc
    serial = serializar(d, achados)
    n = 0
    gerais = []
    for a in serial:
        if a["status"] == OK:
            continue
        locais = a["locais"][:MAX_POR_ITEM]
        if not locais:
            gerais.append(a)
            continue
        linhas = _texto_comentario(a)
        dica = next((v for k, v in COLETIVOS.items() if a["item"].startswith(k)), None)
        if dica and len(a["locais"]) > LIMITE_COLETIVO:
            linhas.insert(1, f"Este problema se repete em {len(a['locais'])} trechos do documento; "
                             "marcado só o primeiro.")
            linhas.append(dica)
            locais = locais[:1]
        elif len(a["locais"]) > MAX_POR_ITEM:
            linhas.append(f"(Este problema aparece em {len(a['locais'])} lugares; marcados os {MAX_POR_ITEM} primeiros.)")
        for k, loc in enumerate(locais):
            par = _paragrafo_do_bloco(d, loc["b"]) if isinstance(loc["b"], int) else None
            if par is None:
                continue
            p_el = par._p
            txt = texto_atual(p_el)
            alvo = []
            if loc.get("m"):
                pos = txt.find(loc["m"])
                if pos >= 0:
                    alvo = _isolar(p_el, pos, pos + len(loc["m"]))
            if not alvo:
                alvo = [r for r in _runs(p_el) if _texto_run(r).strip() and r.getparent() is p_el]
            if not alvo:
                continue
            for r in alvo:
                Run(r, par).font.highlight_color = WD_COLOR_INDEX.YELLOW
            # 1º local leva o comentário completo; os demais, uma referência curta
            texto = linhas if k == 0 else [f"{linhas[0]} (ocorrência {k + 1} de {min(len(a['locais']), MAX_POR_ITEM)}; "
                                           "veja o primeiro comentário deste item)"]
            try:
                _comentar(doc, alvo, par, texto)
                n += 1
            except Exception:
                continue

    # problemas do documento como um todo (margens, paginação, elementos ausentes...): no início
    if gerais:
        primeiro = next((b.obj for b in d.blocos if b.tipo == "p" and b.texto), None)
        if primeiro is not None:
            alvo = [r for r in _runs(primeiro._p) if _texto_run(r).strip() and r.getparent() is primeiro._p]
            if alvo:
                linhas = ["PONTOS GERAIS DO DOCUMENTO (não ligados a um trecho específico):"]
                for a in gerais:
                    linhas.append("")
                    linhas.extend(_texto_comentario(a))
                _comentar(doc, alvo, primeiro, linhas)
                n += 1

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue(), n
