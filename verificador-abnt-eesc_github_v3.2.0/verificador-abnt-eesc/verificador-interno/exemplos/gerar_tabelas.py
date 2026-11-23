"""Gera exemplo_tabelas.docx: vários tipos de tabela para testar a identificação.

Esperado:
 T1  Tabela 1 - correta (dados, título acima, fonte, sem traços verticais, citada)      -> OK
 T2  Tabela 2: título ABAIXO, com dois-pontos, sem fonte, com grade, não citada          -> Ajustar/Conferir
 T3  "Tabela 3" com conteúdo textual                                                     -> Conferir (é Quadro)
 T4  Quadro 1 – sem bordas fechadas                                                      -> Conferir
 T5  tabela de dados sem título                                                          -> Ajustar
 T6  banca examinadora montada em tabela (pré-textual)                                   -> ignorada (diagramação)
 T7  Tabela 5 - longa (30 linhas), sem cabeçalho repetido, salto de numeração 3 -> 5     -> Conferir + Ajustar
 T8  células mescladas com citação (Silva, 2020) dentro da célula                        -> citação reconhecida
"""
import sys
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt


def bordas(tabela, horizontais_so=True):
    tblpr = tabela._tbl.tblPr
    b = OxmlElement("w:tblBorders")
    for lado in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{lado}")
        ativo = lado in ("top", "bottom") if horizontais_so else True
        el.set(qn("w:val"), "single" if ativo else "nil")
        if ativo:
            el.set(qn("w:sz"), "8")
        b.append(el)
    tblpr.append(b)


def gerar(destino):
    doc = Document()
    st = doc.styles["Normal"]
    st.font.name = "Arial"
    st.font.size = Pt(12)
    st.paragraph_format.line_spacing = 1.5
    s = doc.sections[0]
    s.page_width, s.page_height = Cm(21), Cm(29.7)
    s.left_margin, s.top_margin, s.right_margin, s.bottom_margin = Cm(3), Cm(3), Cm(2), Cm(2)

    def p(t, tam=None, centro=False, negrito=False, quebra=False):
        par = doc.add_paragraph()
        if quebra:
            par.add_run().add_break(WD_BREAK.PAGE)
        r = par.add_run(t)
        if tam:
            r.font.size = Pt(tam)
        r.bold = negrito or None
        if centro:
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        return par

    # ---- pré-textual com banca em tabela (diagramação)
    p("FOLHA DE APROVAÇÃO", centro=True, negrito=True)
    p("Aprovado em:")
    t6 = doc.add_table(rows=3, cols=2)
    t6.style = "Table Grid"
    for i, (n, j) in enumerate([("Banca Examinadora", ""), ("Prof. Dr. Fulano", "Julgamento:"),
                                ("Prof. Dr. Beltrano", "Julgamento:")]):
        t6.cell(i, 0).text, t6.cell(i, 1).text = n, j
    doc.add_section(WD_SECTION.NEW_PAGE)

    doc.add_heading("1 INTRODUÇÃO", level=1)
    p("Os resultados estão na Tabela 1 e a comparação de conceitos aparece no Quadro 1. A Tabela 5 "
      "apresenta a série completa de medições realizadas ao longo do experimento.")

    # T1 correta
    p("Tabela 1 - Resistência por lote", tam=10)
    t1 = doc.add_table(rows=4, cols=3)
    bordas(t1)
    for i, linha in enumerate([("Lote", "Média (MPa)", "Desvio"), ("A", "32,4", "1,2"), ("B", "30,1", "0,9"),
                               ("C", "31,7", "1,1")]):
        for j, v in enumerate(linha):
            t1.cell(i, j).text = v
    p("Fonte: elaborada pelo autor.", tam=10)

    # T2 título abaixo, dois-pontos, sem fonte, grade, não citada
    t2 = doc.add_table(rows=3, cols=2)
    t2.style = "Table Grid"
    for i, linha in enumerate([("Ensaio", "Valor"), ("Tração", "45"), ("Compressão", "60")]):
        for j, v in enumerate(linha):
            t2.cell(i, j).text = v
    p("Tabela 2: Ensaios mecânicos", tam=10)
    p("Texto que segue a segunda tabela, sem mencioná-la diretamente em nenhum momento do parágrafo.")

    # T3 "Tabela" com texto
    p("Tabela 3 - Conceitos principais", tam=10)
    t3 = doc.add_table(rows=3, cols=2)
    bordas(t3)
    for i, linha in enumerate([("Conceito", "Definição"),
                               ("Normalização", "Conjunto de regras para padronizar documentos"),
                               ("Referência", "Elementos que identificam uma obra consultada")]):
        for j, v in enumerate(linha):
            t3.cell(i, j).text = v
    p("Fonte: elaborada pelo autor.", tam=10)
    p("A Tabela 3 resume os conceitos usados no trabalho.")

    # T4 quadro sem bordas fechadas
    p("Quadro 1 – Comparação entre normas", tam=10)
    t4 = doc.add_table(rows=3, cols=2)
    bordas(t4)
    for i, linha in enumerate([("Norma", "Objeto"), ("NBR 6023", "Referências"), ("NBR 10520", "Citações")]):
        for j, v in enumerate(linha):
            t4.cell(i, j).text = v
    p("Fonte: elaborado pelo autor.", tam=10)

    # T5 sem título
    p("Os valores medidos em campo foram organizados da seguinte forma, para consulta posterior pela equipe.")
    t5 = doc.add_table(rows=3, cols=3)
    bordas(t5)
    for i, linha in enumerate([("Ponto", "X", "Y"), ("1", "10,5", "3,2"), ("2", "11,0", "3,4")]):
        for j, v in enumerate(linha):
            t5.cell(i, j).text = v
    p("Na sequência, a série completa de medições é apresentada na tabela a seguir.")

    # T7 longa, sem cabeçalho repetido, salto 3 -> 5
    p("Tabela 5 - Série de medições", tam=10)
    t7 = doc.add_table(rows=31, cols=2)
    bordas(t7)
    t7.cell(0, 0).text, t7.cell(0, 1).text = "Dia", "Medição"
    for i in range(1, 31):
        t7.cell(i, 0).text, t7.cell(i, 1).text = str(i), f"{20 + i * 0.3:.1f}".replace(".", ",")
    p("Fonte: elaborada pelo autor.", tam=10)

    # T8 mesclada com citação dentro
    p("Tabela 6 - Fontes consultadas por categoria", tam=10)
    t8 = doc.add_table(rows=3, cols=3)
    bordas(t8)
    t8.cell(0, 0).merge(t8.cell(0, 1)).text = "Categoria e fonte"
    t8.cell(0, 2).text = "Quantidade"
    t8.cell(1, 0).text, t8.cell(1, 1).text, t8.cell(1, 2).text = "Livros", "Normalização (Silva, 2020)", "12"
    t8.cell(2, 0).text, t8.cell(2, 1).text, t8.cell(2, 2).text = "Artigos", "Periódicos nacionais", "30"
    p("Fonte: elaborada pelo autor.", tam=10)
    p("Conforme a Tabela 6, predominam os artigos.")

    p("REFERÊNCIAS", centro=True, negrito=True, quebra=True)
    par = doc.add_paragraph()
    par.paragraph_format.line_spacing = 1.0
    par.add_run("SILVA, J. A. ")
    par.add_run("Normalização documentária").bold = True
    par.add_run(". 2. ed. São Paulo: Atlas, 2020.")
    doc.save(destino)


if __name__ == "__main__":
    destino = Path(sys.argv[1]) if len(sys.argv) > 1 else Path(__file__).with_name("exemplo_tabelas.docx")
    gerar(destino)
    print("gerado:", destino)
