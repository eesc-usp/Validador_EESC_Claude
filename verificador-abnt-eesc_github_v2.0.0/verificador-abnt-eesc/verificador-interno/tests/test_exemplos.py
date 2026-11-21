"""Regressão: os exemplos com erros plantados e os exemplos corretos."""
import io

import pytest
from conftest import EXEMPLOS

from verificador import OK, analisar
from verificador import perfil as PF
from verificador.correcao import GRUPOS, corrigir_duplo
from verificador.documento import Documento
from verificador.edicao import aceitar_tudo
from verificador.trabalho import detectar

CASOS = [
    ("exemplo_com_erros.docx", PF.EESC_POS, 23),
    ("exemplo_correto.docx", PF.EESC_POS, 0),
    ("exemplo_tcc_com_erros.docx", PF.EESC, 23),
    ("exemplo_tcc_correto.docx", PF.EESC, 0),
]


def pendencias(achados):
    return [a for a in achados if a.status != OK]


@pytest.mark.parametrize("arquivo,perfil,esperado", CASOS)
def test_quantidade_de_pendencias(arquivo, perfil, esperado):
    achados, _ = analisar(EXEMPLOS / arquivo, perfil)
    assert not [a for a in achados if a.categoria == "Sistema"], "alguma verificação quebrou"
    assert len(pendencias(achados)) == esperado, [a.item for a in pendencias(achados)]


@pytest.mark.parametrize("arquivo,tipo", [("exemplo_correto.docx", "pos"), ("exemplo_tcc_correto.docx", "tcc")])
def test_detecta_tipo_de_trabalho(arquivo, tipo):
    assert detectar(Documento(EXEMPLOS / arquivo))[0] == tipo


@pytest.mark.parametrize("arquivo,perfil", [("exemplo_correto.docx", PF.EESC_POS),
                                            ("exemplo_tcc_correto.docx", PF.EESC)])
def test_documento_correto_nao_recebe_correcoes(arquivo, perfil):
    _, log, falhas, _ = corrigir_duplo((EXEMPLOS / arquivo).read_bytes(), list(GRUPOS), True, None, perfil)
    assert log == [] and falhas == []


@pytest.mark.parametrize("rastrear", [True, False])
def test_correcao_reduz_pendencias_e_preserva_texto(rastrear):
    dados = (EXEMPLOS / "exemplo_com_erros.docx").read_bytes()
    novo, log, falhas, passadas = corrigir_duplo(dados, list(GRUPOS), rastrear, None, PF.EESC_POS)
    assert log, "nenhuma correção aplicada"
    assert all(p["integra"] for p in passadas)
    antes = len(pendencias(analisar(io.BytesIO(dados), PF.EESC_POS)[0]))
    depois = len(pendencias(analisar(io.BytesIO(aceitar_tudo(novo)), PF.EESC_POS)[0]))
    assert depois < antes
