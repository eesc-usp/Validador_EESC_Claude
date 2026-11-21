"""As acelerações (verificador/cache.py) não podem mudar nenhum resultado.

Roda os mesmos documentos com o cache ligado e desligado e compara tudo."""
import io
import json
from dataclasses import asdict

import pytest
import sinteticos
from conftest import EXEMPLOS, TEMPLATE
from retrato import xml_canonico

from verificador import analisar, cache
from verificador import perfil as PF
from verificador.correcao import GRUPOS, corrigir_duplo
from verificador.visualizacao import renderizar, serializar

DOCUMENTOS = {
    **{p.name: p.read_bytes() for p in sorted(EXEMPLOS.glob("*.docx"))},
    TEMPLATE.name: TEMPLATE.read_bytes(),
    **{n: sinteticos.todos()[n] for n in ("trabalho_longo", "controles_de_conteudo", "estilos_quebrados",
                                          "tabelas_aninhadas")},
}


def _leitura(dados, perfil):
    achados, refs = analisar(io.BytesIO(dados), perfil)
    html_doc, d = renderizar(io.BytesIO(dados))
    return json.dumps({"achados": [asdict(a) for a in achados], "refs": [r.texto for r in refs],
                       "pagina": html_doc, "locais": serializar(d, achados)}, sort_keys=True, default=str)


@pytest.mark.parametrize("nome", sorted(DOCUMENTOS))
@pytest.mark.parametrize("perfil", [PF.EESC, PF.EESC_POS], ids=["tcc", "pos"])
def test_analise_igual_com_e_sem_cache(nome, perfil):
    dados = DOCUMENTOS[nome]
    com = _leitura(dados, perfil)
    with cache.desligado():
        sem = _leitura(dados, perfil)
    assert com == sem


@pytest.mark.parametrize("nome", ["exemplo_com_erros.docx", "exemplo_tcc_com_erros.docx", "trabalho_longo"])
def test_correcao_igual_com_e_sem_cache(nome):
    dados = DOCUMENTOS[nome]

    def corrigir():
        novo, log, falhas, passadas = corrigir_duplo(dados, list(GRUPOS), True, None, PF.EESC_POS)
        return xml_canonico(novo), log, falhas, passadas

    com = corrigir()
    with cache.desligado():
        sem = corrigir()
    assert com == sem


def test_cache_de_leitura_so_dentro_da_analise():
    assert not cache.ativo()
    with cache.leitura():
        assert cache.ativo()
        with cache.leitura():  # aninhado reaproveita o de fora
            assert cache.ativo()
    assert not cache.ativo()


def test_desligado_restaura_as_aceleracoes():
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.parts.document import DocumentPart
    assert hasattr(DocumentPart.get_style, "__wrapped__") and hasattr(BaseOxmlElement.xpath, "__wrapped__")
    with cache.desligado():
        assert not hasattr(DocumentPart.get_style, "__wrapped__")
        assert not hasattr(BaseOxmlElement.xpath, "__wrapped__")
    assert hasattr(DocumentPart.get_style, "__wrapped__") and hasattr(BaseOxmlElement.xpath, "__wrapped__")
