"""Trabalhos em PDF: a estrutura remontada pelo layout deve dar a mesma verificação do .docx de origem.

Os PDFs de teste são sintéticos: 'exemplo_tcc_com_erros.pdf' foi exportado pelo LibreOffice a partir de
exemplos/exemplo_tcc_com_erros.docx, e 'exemplo_latex.pdf' foi compilado do 'exemplo_latex.tex' (pdflatex).
"""
import io
from pathlib import Path

import pytest
from conftest import EXEMPLOS

from verificador import analisar, anotacao, pdf
from verificador.visualizacao import renderizar, serializar

DADOS = Path(__file__).resolve().parent / "dados"
PDF_WORD = DADOS / "exemplo_tcc_com_erros.pdf"
PDF_TEX = DADOS / "exemplo_latex.pdf"
SO_NO_PDF = ("Arquivo em PDF", "Número da página segue", "Número de página em fonte menor")


def _status(achados):
    return {(a.categoria, a.item): a.status for a in achados}


def test_pdf_tem_a_mesma_verificacao_do_docx():
    docx, _ = analisar(str(EXEMPLOS / "exemplo_tcc_com_erros.docx"))
    doc_pdf, refs = analisar(str(PDF_WORD))
    a, b = _status(docx), _status(doc_pdf)
    comuns = [k for k in a if k in b]
    iguais = [k for k in comuns if a[k] == b[k]]
    diferentes = {k[1]: (a[k], b[k]) for k in comuns if a[k] != b[k]}
    assert len(refs) == 6
    assert len(iguais) >= 0.97 * len(comuns), diferentes
    assert not [k for k in a if k not in b], "itens do .docx que o PDF não verificou"
    assert all(k[1].startswith(SO_NO_PDF) for k in b if k not in a)


def test_estrutura_remontada():
    lt = pdf.ler(PDF_WORD.read_bytes())
    pars = [b for b in lt.blocos if isinstance(b, pdf.Paragrafo) and b.linhas]
    titulos = {b.texto: b.nivel for b in pars if b.nivel}
    assert titulos.get("1 INTRODUÇÃO") == 1 and titulos.get("2.1 Materiais") == 2
    assert any(type(b).__name__ == "Tabela" for b in lt.blocos)
    assert sum(type(b).__name__ == "Figura" for b in lt.blocos) == 2
    # parágrafo de 3 linhas, com citação que atravessa a quebra de linha, fica inteiro
    par = next(b for b in pars if b.texto.startswith("A normalização de trabalhos acadêmicos"))
    assert len(par.linhas) == 3 and "Ramos e Pestana (2014)" in par.texto
    assert par.props["linha"] == 1.5
    # citação longa: recuo de 4 cm, fonte 10, espaço simples
    longa = next(b for b in pars if b.texto.startswith("A análise e a síntese"))
    assert round(longa.props["esq"] / pdf.CM, 1) == 4.0 and longa.linhas[0].size == 10 and longa.props["linha"] == 1.0
    # cada caractere sabe onde está na página
    assert len(par.mapa) == len(par.texto) and all(c is not None for c in par.mapa)


def test_pdf_do_latex():
    lt = pdf.ler(PDF_TEX.read_bytes())
    assert lt.tex
    corpo = next(b for b in lt.blocos if isinstance(b, pdf.Paragrafo) and b.texto.startswith("Este trabalho avalia"))
    assert corpo.props["linha"] == 1.5                  # \onehalfspacing do LaTeX = 1,5 do Word
    assert corpo.linhas[0].fonte == "Times New Roman"
    assert any(ln.tipo == "nota" for p in lt.paginas for ln in p.linhas)   # nota de rodapé separada do texto
    sumario = [b.texto for b in lt.blocos if isinstance(b, pdf.Paragrafo) and "\t" in b.texto]
    assert "1.1 Objetivos\t4" in sumario                # pontilhado do sumário vira tabulação
    achados, _ = analisar(str(PDF_TEX))
    st = {a.item: a for a in achados}
    assert st["Paginação"].status == "OK"
    assert st["Margens (cm)"].status == "OK"
    assert st["Entrelinhas 1,5 no texto"].status == "OK"
    assert st["Referências alinhadas à esquerda"].status != "OK"   # no exemplo, as referências estão justificadas


def test_marcacoes_no_pdf_e_no_visualizador():
    dados = PDF_WORD.read_bytes()
    achados, _ = analisar(io.BytesIO(dados))
    html, d = renderizar(io.BytesIO(dados))
    assert html.count("class='folha pdf'") == len(d.pdf.paginas)
    serial = serializar(d, achados)
    com_local = [a for a in serial if a["locais"]]
    assert com_local
    # o trecho marcado está no texto do bloco (o visualizador e o PDF com notas procuram o mesmo texto)
    for a in com_local:
        for loc in a["locais"]:
            if loc.get("m"):
                assert loc["m"] in d.blocos[loc["b"]].texto
    saida, n = anotacao.gerar(dados, achados)
    from pypdf import PdfReader
    r = PdfReader(io.BytesIO(saida))
    anots = [a.get_object() for p in r.pages for a in (p.get("/Annots") or [])]
    assert n > 0 and len(anots) >= n
    assert any(a["/Subtype"] == "/Highlight" for a in anots)
    assert len(r.pages) == len(PdfReader(io.BytesIO(dados)).pages)      # nenhuma página a mais ou a menos


def test_pdf_digitalizado():
    pytest.importorskip("PIL")
    from PIL import Image
    buf = io.BytesIO()
    Image.new("RGB", (595, 842), (255, 255, 255)).save(buf, "PDF")
    achados, refs = analisar(io.BytesIO(buf.getvalue()))
    assert [a.item for a in achados] == ["PDF sem texto (digitalizado)"] and refs == []
