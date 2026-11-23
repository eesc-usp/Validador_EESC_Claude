"""Documentos estranhos não podem quebrar nenhuma função (análise, página, comentários, relatório)."""
import io

import pytest
import sinteticos

from verificador import analisar, anotacao, relatorio
from verificador import perfil as PF
from verificador.documento import Documento
from verificador.visualizacao import renderizar, serializar

NOMES = sorted(sinteticos.GERADORES)


@pytest.mark.parametrize("nome", NOMES)
def test_tudo_funciona(nome):
    dados = sinteticos.todos()[nome]
    for perfil in (PF.EESC, PF.EESC_POS):
        achados, _ = analisar(io.BytesIO(dados), perfil)
        quebradas = [f"{a.item}: {a.detalhe}" for a in achados if a.categoria == "Sistema"]
        assert not quebradas, quebradas
    html_doc, d = renderizar(io.BytesIO(dados))
    assert html_doc
    serializar(d, achados)
    anotacao.gerar(dados, achados)
    assert relatorio.gerar(achados, f"{nome}.docx")


def test_sumario_automatico_do_word_e_lido():
    """O sumário automático fica num controle de conteúdo (w:sdt): as linhas dele precisam ser vistas."""
    d = Documento(io.BytesIO(sinteticos.todos()["controles_de_conteudo"]))
    textos = [b.texto for b in d.blocos]
    assert "2 MÉTODO\t4" in textos
    achados, _ = analisar(io.BytesIO(sinteticos.todos()["controles_de_conteudo"]), PF.EESC)
    sumario = next(a for a in achados if a.item == "Todos os títulos aparecem no sumário")
    assert "MÉTODO" not in " ".join(sumario.trechos)
