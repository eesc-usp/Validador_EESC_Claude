"""Documentos sintéticos para os testes de robustez, de cache e o retrato (nenhum é trabalho de aluno).

Cada função devolve os bytes de um .docx construído do zero, sempre igual (determinístico):
casos extremos (vazio, só tabela, só imagem...), estruturas raras do Word (controles de conteúdo,
equações, campos, estilos quebrados, tabelas aninhadas) e um trabalho longo para medir desempenho.
"""
import base64
import io
from functools import cache

from docx import Document
from docx.enum.text import WD_BREAK
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
from docx.shared import Cm, Pt

TEXTO = "Este parágrafo tem texto suficiente para representar o texto corrido de um trabalho acadêmico. "


def _bytes(doc):
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# imagem PNG 40x30 fixa (bytes guardados aqui, para não depender da versão da biblioteca de imagens)
_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAACgAAAAeCAIAAADRv8uKAAAALElEQVR4nO3NMQEAAAQAMATRP4pYKvg8W4HldMeHelnFYrFYLBaLxWKx"
    "WHyyHdwBQE1gaCAAAAAASUVORK5CYII=")


def _png():
    return io.BytesIO(_PNG)


def _refs(d, *linhas):
    d.add_paragraph("REFERÊNCIAS")
    for ln in linhas:
        d.add_paragraph(ln)


# ------------------------------------------------------------------ casos extremos
def vazio():
    return _bytes(Document())


def so_tabela():
    d = Document()
    t = d.add_table(rows=3, cols=3)
    for i, r in enumerate(t.rows):
        for j, c in enumerate(r.cells):
            c.text = f"{i}{j}"
    return _bytes(d)


def so_imagem():
    d = Document()
    d.add_picture(_png(), width=Cm(4))
    return _bytes(d)


def so_titulos():
    d = Document()
    for t in ("SUMÁRIO", "1 INTRODUÇÃO", "2 DESENVOLVIMENTO", "2.1 Subseção", "REFERÊNCIAS"):
        d.add_paragraph(t)
    return _bytes(d)


def cabecalho_sem_corpo():
    d = Document()
    fld = OxmlElement("w:fldSimple")
    fld.set(qn("w:instr"), "PAGE")
    d.sections[0].header.paragraphs[0]._p.append(fld)
    return _bytes(d)


def ingles():
    d = Document()
    d.add_paragraph("INTRODUCTION")
    d.add_paragraph("This work presents a study (SMITH, 2020). " * 5)
    _refs(d, "SMITH, J. Book title. London: Pub, 2020.")
    return _bytes(d)


def sem_margens():
    d = Document()
    d.add_paragraph("1 INTRODUÇÃO")
    d.add_paragraph(TEXTO)
    sect = d.sections[0]._sectPr
    for tag in ("w:pgMar", "w:pgSz"):
        el = sect.find(qn(tag))
        if el is not None:
            sect.remove(el)
    return _bytes(d)


def estilos_word():
    d = Document()
    for i in range(1, 4):
        d.add_paragraph(f"Capítulo {i}", style="Heading 1")
        d.add_paragraph("Item de lista automática", style="List Number")
        d.add_paragraph(TEXTO * 3)
    return _bytes(d)


# ------------------------------------------------------------------ estruturas raras do Word
def tabelas_aninhadas():
    d = Document()
    c = d.add_table(rows=1, cols=1).cell(0, 0)
    for n in range(8):
        c = c.add_table(rows=1, cols=1).cell(0, 0)
        c.paragraphs[0].text = f"nível {n} (Silva, 2020)"
    d.add_paragraph("Tabela 1 – Depois das tabelas aninhadas")
    _refs(d, "SILVA, J. Obra. São Paulo: Ed., 2020.")
    return _bytes(d)


def equacoes_e_campos():
    d = Document()
    d.add_paragraph("1 INTRODUÇÃO")
    p = d.add_paragraph("Equação: ")
    p._p.append(parse_xml(f'<m:oMath {nsdecls("m")}><m:r><m:t>x=1</m:t></m:r></m:oMath>'))
    r = p.add_run()
    r._r.append(parse_xml(f'<w:sym {nsdecls("w")} w:font="Symbol" w:char="F061"/>'))
    r = p.add_run("a")
    r._r.append(OxmlElement("w:noBreakHyphen"))
    r._r.append(OxmlElement("w:softHyphen"))
    p = d.add_paragraph()
    for xml in ('<w:r %s><w:fldChar w:fldCharType="begin"/></w:r>',
                '<w:r %s><w:instrText> TOC \\o "1-3" </w:instrText></w:r>',
                '<w:r %s><w:fldChar w:fldCharType="separate"/></w:r>',
                '<w:r %s><w:t>Sumário do campo</w:t></w:r>',
                '<w:r %s><w:fldChar w:fldCharType="end"/></w:r>'):
        p._p.append(parse_xml(xml % nsdecls("w")))
    d.add_paragraph("Emoji 😀, עברית, العربية, 中文 (Silva, 2020).")
    _refs(d, "SILVA, J. Obra. São Paulo: Ed., 2020.")
    return _bytes(d)


def controles_de_conteudo():
    """Sumário automático do Word: fica dentro de um controle de conteúdo (w:sdt)."""
    d = Document()
    titulo = d.add_paragraph("SUMÁRIO")
    sdt = parse_xml(
        f'<w:sdt {nsdecls("w")}><w:sdtPr><w:docPartObj><w:docPartGallery w:val="Table of Contents"/>'
        f'</w:docPartObj></w:sdtPr><w:sdtContent>'
        f'<w:p><w:r><w:t>1 INTRODUÇÃO</w:t></w:r><w:r><w:tab/><w:t>3</w:t></w:r></w:p>'
        f'<w:p><w:r><w:t>2 MÉTODO</w:t></w:r><w:r><w:tab/><w:t>4</w:t></w:r></w:p>'
        f'<w:p><w:r><w:t>REFERÊNCIAS</w:t></w:r><w:r><w:tab/><w:t>5</w:t></w:r></w:p>'
        f'</w:sdtContent></w:sdt>')
    titulo._p.addnext(sdt)
    d.add_paragraph("1 INTRODUÇÃO")
    p = d.add_paragraph(TEXTO + "Antes ")
    p._p.append(parse_xml(f'<w:sdt {nsdecls("w")}><w:sdtContent><w:r><w:t>dentro de um controle '
                          f'(Souza, 2019).</w:t></w:r></w:sdtContent></w:sdt>'))
    d.add_paragraph("2 MÉTODO")
    d.add_paragraph(TEXTO)
    _refs(d, "SOUZA, A. Obra. São Paulo: Ed., 2019.")
    return _bytes(d)


def estilos_quebrados():
    d = Document()
    st = d.styles.add_style("Quebrado", 1)
    st.element.append(parse_xml(f'<w:basedOn {nsdecls("w")} w:val="NaoExiste"/>'))
    d.add_paragraph("1 INTRODUÇÃO", style="Quebrado")
    d.add_paragraph(TEXTO + "(Silva, 2020).", style="Quebrado")
    p = d.add_paragraph("Estilo de parágrafo que não existe")
    p._p.get_or_add_pPr().append(parse_xml(f'<w:pStyle {nsdecls("w")} w:val="Fantasma"/>'))
    r = d.add_paragraph().add_run("trecho com estilo de caractere que não existe")
    r._r.get_or_add_rPr().append(parse_xml(f'<w:rStyle {nsdecls("w")} w:val="FantasmaChar"/>'))
    _refs(d, "SILVA, J. Obra. São Paulo: Ed., 2020.")
    return _bytes(d)


def quebras_e_vazios():
    d = Document()
    d.add_paragraph("1 INTRODUÇÃO")
    r = d.add_paragraph().add_run("antes")
    r.add_break(WD_BREAK.COLUMN)
    r.add_text("depois")
    r.add_break(WD_BREAK.PAGE)
    r.add_tab()
    r.add_text("tab")
    r.font.size = Pt(1)
    d.add_paragraph("Fonte gigante").runs[0].font.size = Pt(400)
    for t in ("\t\t\t", "", "   "):
        d.add_paragraph(t)
    d.add_paragraph("REFERÊNCIAS")
    for _ in range(4):
        d.add_paragraph("")
    return _bytes(d)


def so_referencias():
    d = Document()
    _refs(d, *[f"{n}ULANO, A. Obra {n}. São Paulo: Ed., {2000 + k}." for k, n in enumerate("ZYXWVU")])
    return _bytes(d)


def titulos_repetidos():
    d = Document()
    for t in ("1 INTRODUÇÃO", "1 INTRODUÇÃO", "INTRODUÇÃO.", "1.1.1.1.1.1.1.1 MUITO FUNDO", "99 FIM",
              "REFERÊNCIAS", "REFERÊNCIAS", "APÊNDICE A – X", "ANEXO"):
        d.add_paragraph(t)
        d.add_paragraph("Texto qualquer (Autor, 2001).")
    return _bytes(d)


def paragrafo_gigante():
    d = Document()
    p = d.add_paragraph()
    for i in range(800):
        p.add_run(f"palavra{i} ").bold = i % 7 == 0
    d.add_paragraph("1 INTRODUÇÃO")
    d.add_paragraph("Texto " * 2000)
    _refs(d, "SILVA, J. Título. São Paulo: Ed., 2020.")
    return _bytes(d)


# ------------------------------------------------------------------ trabalho longo (desempenho)
def trabalho_longo(capitulos=4, secoes=3, paragrafos=8, referencias=30):
    d = Document()
    for s in d.sections:
        s.left_margin = s.top_margin = Cm(3)
        s.right_margin = s.bottom_margin = Cm(2)
    d.add_paragraph("UNIVERSIDADE DE SÃO PAULO\nESCOLA DE ENGENHARIA DE SÃO CARLOS")
    d.add_paragraph("SUMÁRIO")
    for c in range(1, capitulos + 1):
        d.add_paragraph(f"{c} CAPÍTULO {c}\t{c * 10}")
    d.add_paragraph("REFERÊNCIAS\t99")
    d.add_paragraph("1 INTRODUÇÃO")
    fig = 0
    for c in range(1, capitulos + 1):
        if c > 1:
            d.add_paragraph(f"{c} CAPÍTULO {c}")
        for sec in range(1, secoes + 1):
            d.add_paragraph(f"{c}.{sec} Seção {c}.{sec}")
            for k in range(paragrafos):
                a = (c * 7 + sec * 3 + k) % referencias
                d.add_paragraph(f"Segundo Autor{a:02d} ({1990 + a % 30}), o processo da Figura {fig + 1} "
                                f"apresenta resultados relevantes (Silva{a:02d}, {2000 + a % 20}). " + TEXTO)
            fig += 1
            d.add_paragraph(f"Figura {fig} – Ilustração {fig}")
            d.add_picture(_png(), width=Cm(3))
            d.add_paragraph("Fonte: Elaborado pelo autor.")
            if sec == 2:
                d.add_paragraph(f"Tabela {c} – Dados do capítulo {c}")
                t = d.add_table(rows=4, cols=3)
                for i, r in enumerate(t.rows):
                    for j, cel in enumerate(r.cells):
                        cel.text = str(i * j * 1.5)
                d.add_paragraph("Fonte: Elaborado pelo autor.")
    d.add_paragraph("REFERÊNCIAS")
    for a in range(referencias):
        d.add_paragraph(f"AUTOR{a:02d}, A. B. Título da obra número {a}. São Paulo: Editora, {1990 + a % 30}.")
        d.add_paragraph("")
    return _bytes(d)


GERADORES = {
    "vazio": vazio, "so_tabela": so_tabela, "so_imagem": so_imagem, "so_titulos": so_titulos,
    "cabecalho_sem_corpo": cabecalho_sem_corpo, "ingles": ingles, "sem_margens": sem_margens,
    "estilos_word": estilos_word, "tabelas_aninhadas": tabelas_aninhadas, "equacoes_e_campos": equacoes_e_campos,
    "controles_de_conteudo": controles_de_conteudo, "estilos_quebrados": estilos_quebrados,
    "quebras_e_vazios": quebras_e_vazios, "so_referencias": so_referencias, "titulos_repetidos": titulos_repetidos,
    "paragrafo_gigante": paragrafo_gigante, "trabalho_longo": trabalho_longo,
}


@cache
def todos():
    """{nome: bytes do .docx} de todos os documentos sintéticos (gerados uma vez por sessão de teste)."""
    return {nome: f() for nome, f in GERADORES.items()}
