"""Relatório .docx editável para o bibliotecário revisar e enviar ao aluno."""
import io
import re
from datetime import date

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Cm, Pt, RGBColor

from .cache import qn
from .modelo import AJUSTAR, OK, VERIFICAR

LARGURAS = (4.2, 1.9, 5.6, 5.3)  # cm; soma = largura útil A4 com margens padrão
CORES = {AJUSTAR: "F8D7DA", VERIFICAR: "FFF3CD", OK: "D4EDDA"}


def _sombrear(celula, cor):
    tcpr = celula._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:fill"), cor)
    tcpr.append(shd)


def _texto_com_negrito(par, texto):
    for i, parte in enumerate(re.split(r"\*\*", texto)):
        if parte:
            par.add_run(parte).bold = (i % 2 == 1)


def gerar(achados, nome_arquivo, incluir_ok=False, observacao=""):
    doc = Document()
    for sec in doc.sections:
        sec.page_width, sec.page_height = Cm(21), Cm(29.7)
        sec.left_margin = sec.right_margin = Cm(2)
    estilo = doc.styles["Normal"]
    estilo.font.name = "Arial"
    estilo.font.size = Pt(10)

    doc.add_heading("Relatório de verificação de normalização", level=1)
    doc.add_paragraph(f"Biblioteca da Escola de Engenharia de São Carlos – USP\n"
                      f"Arquivo: {nome_arquivo}\nData: {date.today():%d/%m/%Y}\n"
                      "Referência: Diretrizes para apresentação de dissertações e teses da USP: parte I (ABNT), "
                      "5. ed., 2024.")
    n = {s: sum(1 for a in achados if a.status == s) for s in (AJUSTAR, VERIFICAR, OK)}
    doc.add_paragraph(f"Resumo: {n[AJUSTAR]} item(ns) a ajustar, {n[VERIFICAR]} a conferir, {n[OK]} em conformidade.")
    if observacao.strip():
        doc.add_paragraph().add_run("Observações da Biblioteca: ").bold = True
        doc.add_paragraph(observacao.strip())

    # tabela-resumo dos elementos (aparece · formatação · conteúdo)
    resumo = next((a for a in achados if a.tabela and a.item.startswith("Elementos: aparece")), None)
    if resumo:
        doc.add_heading("Elementos do trabalho", level=2)
        lin = [x for x in resumo.tabela if not x["Aparece"].startswith("—")]
        cols = list(lin[0].keys()) if lin else []
        if cols:
            tb = doc.add_table(rows=1, cols=len(cols))
            tb.style = "Table Grid"
            for c, titulo in zip(tb.rows[0].cells, cols):
                c.text = ""
                c.paragraphs[0].add_run(titulo).bold = True
            for linha in lin:
                cel = tb.add_row().cells
                for c, k in zip(cel, cols):
                    c.text = str(linha[k])
                    for r in c.paragraphs[0].runs:
                        r.font.size = Pt(8)
        achados = [a for a in achados if a is not resumo]

    categorias = list(dict.fromkeys(a.categoria for a in achados))
    for cat in categorias:
        itens = [a for a in achados if a.categoria == cat and (incluir_ok or a.status != OK)]
        if not itens:
            continue
        doc.add_heading(cat, level=2)
        tab = doc.add_table(rows=1, cols=4)
        tab.style = "Table Grid"
        for c, titulo in zip(tab.rows[0].cells, ("Item", "Situação", "O que foi encontrado", "Como corrigir")):
            c.text = ""
            c.paragraphs[0].add_run(titulo).bold = True
        for a in itens:
            linha = tab.add_row().cells
            linha[0].text = a.item + (f" (Diretrizes §{a.secao})" if a.secao else "")
            linha[1].text = a.status
            _sombrear(linha[1], CORES.get(a.status, "FFFFFF"))
            linha[2].text = a.detalhe
            if a.tabela:
                cols = list(a.tabela[0].keys())
                for reg in [dict(zip(cols, cols))] + a.tabela:
                    pt = linha[2].add_paragraph()
                    r = pt.add_run(f"{reg[cols[0]]}: " + " | ".join(str(reg[c]) for c in cols[1:]))
                    r.font.size = Pt(8)
                    r.bold = reg[cols[0]] in (cols[0], "Esperado")
            for tr in a.trechos[:5]:
                p = linha[2].add_paragraph()
                r = p.add_run("• " + tr)
                r.italic = True
                r.font.size = Pt(8)
                r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
            linha[3].text = ""
            _texto_com_negrito(linha[3].paragraphs[0], a.como_corrigir)
        tab.autofit = False
        for col, larg in zip(tab.columns, LARGURAS):
            col.width = Cm(larg)
        for linha in tab.rows:
            for c, larg in zip(linha.cells, LARGURAS):
                c.width = Cm(larg)

    doc.add_paragraph()
    nota = doc.add_paragraph("Verificação automática de apoio. Itens marcados como 'Verificar' precisam de "
                             "conferência manual. Sugestões geradas por IA devem ser conferidas. "
                             "Este relatório não substitui as Diretrizes USP nem as orientações do Programa.")
    nota.runs[0].font.size = Pt(8)
    nota.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()
