"""O verificador só mostra e aponta: nada de editar ou corrigir o trabalho do aluno."""
import importlib.util
import io
import re
from pathlib import Path

import sinteticos
from conftest import EXEMPLOS, TEMPLATE

from verificador import analisar
from verificador import perfil as PF
from verificador.visualizacao import renderizar, serializar

RAIZ = Path(__file__).resolve().parent.parent
COMPONENTE = (RAIZ / "componentes" / "area_trabalho" / "index.html").read_text("utf-8")


def _folhas(html_doc):
    return re.findall(r"<section class='folha'.*?</section>", html_doc, re.S)


def test_nao_existe_mais_correcao_nem_edicao():
    for modulo in ("verificador.correcao", "verificador.edicao"):
        assert importlib.util.find_spec(modulo) is None, modulo
    app = (RAIZ / "app.py").read_text("utf-8")
    for termo in ("corrigir_duplo", "edicao.editar", "aceitar_tudo", "Versão limpa", "Documento revisado"):
        assert termo not in app, termo


def test_visualizador_e_somente_leitura():
    for termo in ("contenteditable", "textarea", "abrirEditor", "setComponentValue", "dblclick"):
        assert termo not in COMPONENTE, termo
    html_doc, _ = renderizar(io.BytesIO((EXEMPLOS / "exemplo_com_erros.docx").read_bytes()))
    assert "data-t=" not in html_doc and "contenteditable" not in html_doc


def test_paginas_seguem_a_paginacao_do_word():
    """O template foi salvo no Word com 53 páginas (inclui os versos em branco): a página daqui é a do Word."""
    html_doc, _ = renderizar(io.BytesIO(TEMPLATE.read_bytes()))
    folhas = _folhas(html_doc)
    assert all("data-paginacao='word'" in f for f in folhas)
    assert 50 <= len(folhas) <= 54
    assert sum("<div class='area'></div>" in f for f in folhas) >= 10  # versos em branco continuam lá


def test_arquivo_sem_paginacao_do_word_fica_para_o_navegador():
    html_doc, _ = renderizar(io.BytesIO(sinteticos.todos()["trabalho_longo"]))
    assert "data-paginacao" not in html_doc  # o visualizador pagina pela altura do conteúdo


def test_quebra_de_pagina_no_meio_do_paragrafo_continua_na_folha_seguinte():
    from docx import Document
    from docx.enum.text import WD_BREAK
    d = Document()
    r = d.add_paragraph().add_run("antes da quebra")
    r.add_break(WD_BREAK.PAGE)
    r.add_text("depois da quebra")
    buf = io.BytesIO()
    d.save(buf)
    folhas = _folhas(renderizar(io.BytesIO(buf.getvalue()))[0])
    assert len(folhas) == 2
    assert "antes da quebra" in folhas[0] and "depois da quebra" in folhas[1]
    assert "class='par cont'" in folhas[1]


def test_todo_achado_com_local_aponta_para_um_bloco_desenhado():
    dados = (EXEMPLOS / "exemplo_com_erros.docx").read_bytes()
    achados, _ = analisar(io.BytesIO(dados), PF.EESC_POS)
    html_doc, d = renderizar(io.BytesIO(dados))
    blocos = set(re.findall(r"data-b='(\d+)'", html_doc))
    locais = [loc for a in serializar(d, achados) for loc in a["locais"]]
    assert locais
    assert all(str(loc["b"]) in blocos for loc in locais)
