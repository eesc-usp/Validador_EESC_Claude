"""Leitura do texto dos parágrafos como o aluno vê (somente leitura: o verificador nunca altera o trabalho).

- inclui o texto de hyperlinks, campos e alterações controladas *inseridas*;
- exclui o texto *apagado* em alterações controladas e o de caixas de texto/formas que ficam "dentro" de
  um trecho (esse texto não pertence ao parágrafo e, em arquivos convertidos de PDF, aparecia duplicado).
"""
import io

from docx import Document

from .cache import qn

_W_T, _W_TAB, _W_BR, _W_CR, _W_R = qn("w:t"), qn("w:tab"), qn("w:br"), qn("w:cr"), qn("w:r")
_MC_FALLBACK = "{http://schemas.openxmlformats.org/markup-compatibility/2006}Fallback"
_FORA = {qn("w:del"), qn("w:moveFrom"), qn("w:r"), qn("w:txbxContent"), _MC_FALLBACK}


def texto_run(r):
    """Texto de um trecho (w:r): tabulação vira \\t e quebra de linha vira \\n."""
    out = []
    for ch in r:
        tag = ch.tag
        if tag == _W_T:
            out.append(ch.text or "")
        elif tag == _W_TAB:
            out.append("\t")
        elif tag == _W_BR or tag == _W_CR:
            out.append("\n")
    return "".join(out)


def runs_visiveis(p):
    """Trechos (w:r) visíveis do próprio parágrafo, em ordem (inclui hyperlinks e alterações inseridas)."""
    res = []
    for r in p.iter(_W_R):
        anc, fora = r.getparent(), False
        while anc is not None and anc is not p:
            if anc.tag in _FORA:
                fora = True
                break
            anc = anc.getparent()
        if not fora:
            res.append(r)
    return res


def texto_atual(par):
    """Texto visível de um parágrafo (python-docx Paragraph ou elemento w:p)."""
    p = getattr(par, "_p", par)
    return "".join(texto_run(r) for r in runs_visiveis(p))


def tem_alteracoes(dados):
    """Se o arquivo tem alterações controladas do Word (inseridas ou apagadas) ainda não aceitas."""
    body = Document(io.BytesIO(dados)).element.body
    return any(True for _ in body.iter(qn("w:ins"))) or any(True for _ in body.iter(qn("w:del")))
