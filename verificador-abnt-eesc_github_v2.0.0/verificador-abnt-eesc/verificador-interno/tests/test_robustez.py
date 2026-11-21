"""Documentos estranhos não podem quebrar nenhuma função (análise, página, apontamentos, relatório, correção)."""
import io

import pytest
import sinteticos
from docx.oxml.ns import qn

from verificador import analisar, anotacao, relatorio
from verificador import perfil as PF
from verificador.correcao import GRUPOS, corrigir_duplo, texto_atual
from verificador.documento import Documento
from verificador.edicao import editar
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
    for rastrear in (True, False):
        _, _, falhas, passadas = corrigir_duplo(dados, list(GRUPOS), rastrear, None, PF.EESC)
        assert all(p["integra"] for p in passadas), passadas
        assert not [p for p in passadas if p["resultado"].startswith("interrompida")], passadas


@pytest.mark.parametrize("nome", ["trabalho_longo", "controles_de_conteudo", "equacoes_e_campos"])
def test_edicao_de_paragrafo(nome):
    dados = sinteticos.todos()[nome]
    d = Documento(io.BytesIO(dados))
    alvo = next(b for b in d.blocos if b.tipo == "p" and len(b.texto) > 30 and not set("\n\t") & set(b.bruto)
                and all(r.getparent() is b.obj._p for r in b.obj._p.iter(qn("w:r"))))
    orig = texto_atual(alvo.obj)
    saida, _ = editar(dados, str(alvo.i), orig + " (editado)", orig, True)
    novo = Documento(io.BytesIO(saida))
    assert "(editado)" in texto_atual(novo.blocos[alvo.i].obj)


def test_sumario_automatico_do_word_e_lido():
    """O sumário automático fica num controle de conteúdo (w:sdt): as linhas dele precisam ser vistas."""
    d = Documento(io.BytesIO(sinteticos.todos()["controles_de_conteudo"]))
    textos = [b.texto for b in d.blocos]
    assert "2 MÉTODO\t4" in textos
    achados, _ = analisar(io.BytesIO(sinteticos.todos()["controles_de_conteudo"]), PF.EESC)
    sumario = next(a for a in achados if a.item == "Todos os títulos aparecem no sumário")
    assert "MÉTODO" not in " ".join(sumario.trechos)


def test_edicao_dentro_do_sumario_automatico():
    dados = sinteticos.todos()["controles_de_conteudo"]
    d = Documento(io.BytesIO(dados))
    alvo = next(b for b in d.blocos if b.texto == "2 MÉTODO\t4")
    saida, _ = editar(dados, str(alvo.i), "2 MÉTODOS\t4", texto_atual(alvo.obj), True)
    novo = Documento(io.BytesIO(saida)).blocos[alvo.i]
    assert texto_atual(novo.obj) == "2 MÉTODOS\t4"  # edição rastreada: o texto atual já a inclui
    assert novo.obj._p.getparent().tag == qn("w:sdtContent")  # continua dentro do sumário automático
