"""Identificação tolerante e apontamento no lugar certo.

Como nos verificadores de manuscritos (Penelope, Recite) e nos extratores de estrutura (GROBID, CERMINE):
cada apontamento fica ligado ao trecho que o gerou, e a busca de texto escolhe a ocorrência na zona certa
(o mesmo texto costuma aparecer também no sumário e nas listas).
"""
import io

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH

from verificador import analisar
from verificador import perfil as PF
from verificador.documento import Documento
from verificador.visualizacao import renderizar, serializar

TEXTO = ("Este parágrafo tem texto suficiente para representar o texto corrido de um trabalho acadêmico "
         "e para que o verificador reconheça a parte textual. ")


def _docx(montar):
    d = Document()
    montar(d)
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def _centro(p):
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    return p


def _achados_e_locais(dados, perfil=PF.EESC):
    achados, _ = analisar(io.BytesIO(dados), perfil)
    _, d = renderizar(io.BytesIO(dados))
    textos = {b.i: b.texto for b in d.blocos}
    return {a["item"]: a for a in serializar(d, achados)}, textos


def _trabalho(d, rotulo_palavras="Palavras-chave: Termo um. Termo dois.", titulo_resumo="RESUMO",
              primeira_secao="1 INTRODUÇÃO", legenda="Figura 1: Esquema do processo"):
    _centro(d.add_paragraph(titulo_resumo))
    d.add_paragraph("SILVA, A. Título do trabalho. 2026. Trabalho de Conclusão de Curso (Graduação) – Escola "
                    "de Engenharia de São Carlos, Universidade de São Paulo, São Carlos, 2026.")
    d.add_paragraph(TEXTO * 3)
    for linha in rotulo_palavras.split("\n"):
        d.add_paragraph(linha)
    _centro(d.add_paragraph("LISTA DE FIGURAS"))
    d.add_paragraph(legenda + "\t12")
    _centro(d.add_paragraph("SUMÁRIO"))
    d.add_paragraph(primeira_secao + "\t10")
    d.add_paragraph("2. MÉTODO\t11")
    d.add_paragraph(primeira_secao).runs[0].bold = True
    d.add_paragraph(TEXTO + "Como afirma a literatura (SOUZA, 2019), o tema é recorrente.")
    d.add_paragraph(legenda)
    d.add_paragraph("Fonte: Elaborado pelo autor.")
    d.add_paragraph("2. MÉTODO").runs[0].bold = True
    d.add_paragraph(TEXTO)
    _centro(d.add_paragraph("REFERÊNCIAS"))
    d.add_paragraph("SOUZA, A. Obra. São Paulo: Ed., 2019.")


def test_palavras_chave_sem_dois_pontos_sao_reconhecidas_e_apontadas():
    dados = _docx(lambda d: _trabalho(d, rotulo_palavras="Palavras chave – Termo um, Termo dois, Termo três"))
    itens, textos = _achados_e_locais(dados)
    assert itens["Palavras-chave"]["status"] == "OK"
    fmt = itens["Resumo: formato de 'Palavras-chave'"]
    assert "rótulo sem dois-pontos" in fmt["detalhe"]
    assert textos[fmt["locais"][0]["b"]].startswith("Palavras chave")  # na linha certa, não no título RESUMO


def test_rotulo_sozinho_com_termos_na_linha_seguinte():
    dados = _docx(lambda d: _trabalho(d, rotulo_palavras="Palavras-chave:\nTermo um. Termo dois."))
    itens, textos = _achados_e_locais(dados)
    fmt = itens["Resumo: formato de 'Palavras-chave'"]
    assert fmt["status"] == "OK", fmt["detalhe"]


def test_titulo_com_letras_espacadas():
    dados = _docx(lambda d: _trabalho(d, titulo_resumo="R E S U M O"))
    itens, _ = _achados_e_locais(dados)
    assert itens["Resumo"]["status"] == "OK"


def test_texto_sem_titulo_introducao_comeca_na_primeira_secao():
    dados = _docx(lambda d: _trabalho(d, primeira_secao="1 CONSIDERAÇÕES INICIAIS"))
    d = Documento(io.BytesIO(dados))
    assert d.i_intro is not None and d.blocos[d.i_intro].texto == "1 CONSIDERAÇÕES INICIAIS"
    itens, _ = _achados_e_locais(dados)
    assert itens["Início do texto sem o título INTRODUÇÃO"]["status"] == "Verificar"


def test_legenda_e_titulo_apontados_no_texto_e_nao_nas_listas():
    dados = _docx(_trabalho)
    itens, textos = _achados_e_locais(dados)
    for item in ("Ilustrações: título no formato 'Figura 1 – Título'",
                 "Indicativo numérico sem ponto, hífen ou dois-pontos"):
        locais = itens[item]["locais"]
        assert locais, item
        assert all("\t" not in textos[x["b"]] for x in locais), (item, [textos[x["b"]] for x in locais])


def test_citacao_marcada_exatamente():
    dados = _docx(_trabalho)
    itens, textos = _achados_e_locais(dados)
    loc = itens["Autor em maiúsculas e minúsculas na chamada"]["locais"][0]
    assert loc["m"] == "(SOUZA, 2019)"
    assert "literatura" in textos[loc["b"]]


def test_palavras_chave_coladas_no_resumo_com_quebra_de_linha():
    def montar(d):
        _trabalho(d, rotulo_palavras="")
        p = d.paragraphs[2]  # parágrafo do resumo
        r = p.add_run()
        r.add_break()
        r.add_text("Palavras-chave: Termo um. Termo dois.")
    itens, textos = _achados_e_locais(_docx(montar))
    assert itens["Palavras-chave"]["status"] == "OK"
    fmt = itens["Resumo: formato de 'Palavras-chave'"]
    assert "mesmo parágrafo do resumo" in fmt["detalhe"]
