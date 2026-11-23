"""Documento escrito fora do modelo: a identificação deve funcionar mesmo assim."""
import io

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import RGBColor

from verificador import OK, analisar
from verificador import perfil as PF
from verificador.documento import Documento

LONGO = ("Este parágrafo descreve o trabalho realizado com bastante detalhe para que o verificador reconheça "
         "o texto corrido do documento e possa conferir a formatação. ")


def documento_fora_do_modelo():
    d = Document()
    d.add_paragraph("Universidade de São Paulo")
    d.add_paragraph("Sumario")
    d.add_paragraph("Resumo\t3")
    d.add_paragraph("1 Introdução\t5")
    d.add_paragraph("Capítulo 1 – Introdução:")
    d.add_paragraph(LONGO + "A Associação Brasileira de Normas Técnicas (ABNT) publica normas. O IBGE e o "
                    "IBGE divulgam dados ( conforme Silva , 2020).")
    d.add_paragraph(LONGO + "Texto com  espaço duplo e com o CNPq citado sem nome, e o CNPq de novo.")
    t = d.add_paragraph("2.1 Materiais.")
    t.runs[0].bold = True
    t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p = d.add_paragraph(LONGO + "Trecho esquecido em vermelho.")
    p.runs[0].font.color.rgb = RGBColor(0xFF, 0, 0)
    d.add_paragraph("FIGURA 1 - Esquema do processo")
    d.add_paragraph("Fonte: autor.")
    d.add_paragraph(LONGO)
    d.add_paragraph("Referencias Bibliograficas")
    d.add_paragraph("SILVA, J. Livro. São Paulo: Ed., 2020.")
    buf = io.BytesIO()
    d.save(buf)
    return buf.getvalue()


def test_reconhece_regioes_fora_do_modelo():
    doc = Documento(io.BytesIO(documento_fora_do_modelo()))
    assert doc.i_intro is not None and doc.i_refs is not None


def test_aponta_problemas_tipograficos():
    a, _ = analisar(io.BytesIO(documento_fora_do_modelo()), PF.EESC)
    nomes = {x.item for x in a if x.status != OK}
    for esperado in ("Sem espaços duplos entre palavras", "Sem espaço antes da pontuação",
                     "Título de seção sem ponto final", "Título de seção alinhado à esquerda",
                     "Palavra designativa da legenda no padrão", "Texto na cor preta, sem realce",
                     "Sigla por extenso na primeira menção", "Sumário sem os elementos pré-textuais"):
        assert esperado in nomes, esperado


def test_comentarios_no_documento_fora_do_modelo():
    from verificador import anotacao
    dados = documento_fora_do_modelo()
    achados, _ = analisar(io.BytesIO(dados), PF.EESC)
    _, n = anotacao.gerar(dados, achados)
    assert n >= 5
