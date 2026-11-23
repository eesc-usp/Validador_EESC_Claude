"""Normas personalizadas, conferências de consistência e template EESC."""
import io

from conftest import EXEMPLOS, TEMPLATE
from docx import Document

from verificador import OK, analisar
from verificador import perfil as PF
from verificador.documento import Documento
from verificador.trabalho import detectar


def itens(achados):
    return {a.item: a for a in achados if a.status != OK}


def test_norma_personalizada_muda_a_verificacao():
    p = {"nome": "Teste", "fonte_texto": 11, "regras": [
        {"nome": "Sem 'presente trabalho'", "tipo": "proibir", "padrao": "presente trabalho", "onde": "documento"},
        {"nome": "Tem AGRADECIMENTOS", "tipo": "exigir", "padrao": "AGRADECIMENTOS", "onde": "documento"}]}
    achados, _ = analisar(EXEMPLOS / "exemplo_tcc_correto.docx", p)
    nomes = set(itens(achados))
    assert "Fonte do texto (tamanho 11)" in nomes
    assert "Tem AGRADECIMENTOS" in nomes


def test_sanear_ignora_valores_invalidos():
    p = PF.sanear({"margens": {"esquerda": "2,5", "direita": 99}, "citacao_autor": "xpto", "lixo": 1})
    assert p["margens"]["esquerda"] == 2.5 and p["margens"]["direita"] == 2.0
    assert p["citacao_autor"] == "mista"


def test_padroes_eesc_sem_diferencas():
    assert PF.eh_padrao(PF.EESC) and PF.eh_padrao(PF.EESC_POS)
    assert PF.EESC["tipo"] == "tcc" and PF.EESC_POS["tipo"] == "pos"


def test_consistencia_ano_titulo_e_mencoes():
    d = Document(EXEMPLOS / "exemplo_correto.docx")
    for r in d.paragraphs[3].runs:
        r.text = r.text.replace("2026", "2025")
    d.paragraphs[2].runs[0].text = "Outro título qualquer"
    for p in d.paragraphs:
        if p.text.startswith("Foram analisados"):
            p.runs[-1].text += " Ver a Figura 9 e a seção 5.2."
    buf = io.BytesIO()
    d.save(buf)
    nomes = itens(analisar(io.BytesIO(buf.getvalue()), PF.EESC_POS)[0])
    assert "Autor, título e ano iguais em todo o trabalho" in nomes
    assert "Figuras, tabelas, seções e apêndices citados existem" in nomes


def test_template_eesc():
    assert detectar(Documento(TEMPLATE))[0] == "pos"
    nomes = itens(analisar(TEMPLATE, PF.EESC_POS)[0])
    assert "Textos de exemplo do template removidos" in nomes
