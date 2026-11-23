"""Regressão: os exemplos com erros plantados e os exemplos corretos."""
import io

import pytest
from conftest import EXEMPLOS

from verificador import OK, analisar, anotacao
from verificador import perfil as PF
from verificador.documento import Documento
from verificador.texto import texto_atual
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


def _textos(dados):
    d = Documento(io.BytesIO(dados))
    return [texto_atual(b.obj) if b.tipo == "p" else [c.text for r in b.obj.rows for c in r.cells] for b in d.blocos]


@pytest.mark.parametrize("arquivo,perfil,esperado", CASOS)
def test_comentarios_nao_alteram_o_texto_do_aluno(arquivo, perfil, esperado):
    """O arquivo com comentários tem exatamente o mesmo texto do enviado: só ganha comentários e realce."""
    dados = (EXEMPLOS / arquivo).read_bytes()
    achados, _ = analisar(io.BytesIO(dados), perfil)
    comentado, n = anotacao.gerar(dados, achados)
    assert _textos(comentado) == _textos(dados)
    assert (n > 0) == (esperado > 0)
