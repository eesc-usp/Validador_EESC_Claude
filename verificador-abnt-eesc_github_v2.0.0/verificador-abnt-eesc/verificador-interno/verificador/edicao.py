"""Edição manual de parágrafos feita na área de trabalho, e utilidades de alterações controladas."""
import difflib
import io
import re

from docx import Document

from .cache import qn
from .correcao import Corretor, texto_atual


class EdicaoRecusada(Exception):
    pass


def editar(dados, bloco, novo_texto, texto_original, rastrear=True):
    """Aplica a edição de um parágrafo. Devolve (bytes do novo .docx, aviso).

    Só a parte que mudou é trocada (prefixo e sufixo iguais são preservados), então a formatação
    do resto do parágrafo continua igual. Com 'rastrear', a troca entra como alteração controlada.
    """
    c = Corretor(io.BytesIO(dados), rastrear=rastrear, grupos=[])
    par = _localizar(c.d, bloco)
    atual = texto_atual(par)
    if atual != texto_original:
        raise EdicaoRecusada("O parágrafo mudou desde que foi aberto. Tente de novo.")
    novo = novo_texto.replace("\r\n", "\n")
    if novo == atual:
        return dados, ""
    aviso = ""
    if not _aplicar_diff(c, par, atual, novo):
        # trecho tocava uma alteração controlada anterior: aceita as deste parágrafo e tenta de novo
        _aceitar(par._p)
        atual2 = texto_atual(par)
        if not _aplicar_diff(c, par, atual2, novo):
            raise EdicaoRecusada("Essa parte do parágrafo tem formatação especial (campo, link ou quebra) e "
                                 "não pode ser editada aqui. Edite no Word.")
        aviso = "As alterações controladas anteriores deste parágrafo foram aceitas para permitir a edição."
    buf = io.BytesIO()
    c.doc.save(buf)
    return buf.getvalue(), aviso


def _localizar(d, bloco):
    """'12' = parágrafo do corpo; '12:r:c:k' = parágrafo k da célula c da linha r da tabela 12."""
    from docx.text.paragraph import Paragraph
    partes = str(bloco).split(":")
    try:
        i = int(partes[0])
        b = d.blocos[i]
    except (ValueError, IndexError):
        raise EdicaoRecusada("Trecho não encontrado no documento.") from None
    if len(partes) == 1:
        if b.tipo != "p":
            raise EdicaoRecusada("Para editar uma tabela, dê duplo clique dentro da célula.")
        return b.obj
    if b.tipo != "t" or len(partes) != 4:
        raise EdicaoRecusada("Trecho não encontrado no documento.")
    r, cidx, k = (int(x) for x in partes[1:])
    try:
        tc = b.obj._tbl.tr_lst[r].tc_lst[cidx]
        p_el = [el for el in tc if el.tag == qn("w:p")][k]
    except IndexError:
        raise EdicaoRecusada("Célula não encontrada (a tabela mudou?).") from None
    return Paragraph(p_el, b.obj)


def _tokens(s):
    return [(m.start(), m.end()) for m in re.finditer(r"\s+|\w+|[^\w\s]", s)]


def _aplicar_diff(c, par, atual, novo):
    """Troca só as palavras que mudaram, da última para a primeira (offsets continuam válidos)."""
    ta, tn = _tokens(atual), _tokens(novo)
    sm = difflib.SequenceMatcher(None, [atual[a:b] for a, b in ta], [novo[a:b] for a, b in tn], autojunk=False)
    ops = [op for op in sm.get_opcodes() if op[0] != "equal"]
    if not ops:
        return True
    plano = []
    for _, i1, i2, j1, j2 in ops:
        ini = ta[i1][0] if i1 < len(ta) else len(atual)
        fim = ta[i2 - 1][1] if i2 > i1 else ini
        n_ini = tn[j1][0] if j1 < len(tn) else len(novo)
        n_fim = tn[j2 - 1][1] if j2 > j1 else n_ini
        plano.append((ini, fim, novo[n_ini:n_fim]))
    for ini, fim, txt in reversed(plano):
        if not c.trocar(par, ini, fim, txt):
            return False
    return texto_atual(par) == novo


def _aceitar(p):
    for tag in ("w:del", "w:moveFrom"):
        for el in list(p.iter(qn(tag))):
            el.getparent().remove(el)
    for tag in ("w:ins", "w:moveTo"):
        for el in list(p.iter(qn(tag))):
            pai = el.getparent()
            pos = pai.index(el)
            for filho in list(el):
                pai.insert(pos, filho)
                pos += 1
            pai.remove(el)


def aceitar_tudo(dados):
    """Versão 'limpa': aceita todas as alterações controladas do documento."""
    doc = Document(io.BytesIO(dados))
    partes = [doc.element.body]
    for s in doc.sections:
        for hf in (s.header, s.footer):
            try:
                if not hf.is_linked_to_previous:
                    partes.append(hf._element)
            except Exception:
                pass
    for raiz in partes:
        # marca de parágrafo apagada (w:pPr/w:rPr/w:del): o parágrafo se une ao seguinte
        for p in list(raiz.iter(qn("w:p"))):
            ppr = p.find(qn("w:pPr"))
            rpr = ppr.find(qn("w:rPr")) if ppr is not None else None
            if rpr is None or rpr.find(qn("w:del")) is None:
                continue
            prox = p.getnext()
            if prox is None or prox.tag != qn("w:p") or (ppr.find(qn("w:sectPr")) is not None):
                rpr.remove(rpr.find(qn("w:del")))
                continue
            conteudo = [c for c in p if c.tag != qn("w:pPr")]
            vazio = not "".join(t.text or "" for c in conteudo for t in c.iter(qn("w:t"))
                                if not any(a.tag == qn("w:del") for a in t.iterancestors()))
            if not vazio:
                ref = prox.find(qn("w:pPr"))
                pos = 0 if ref is None else prox.index(ref) + 1
                for c in conteudo:
                    prox.insert(pos, c)
                    pos += 1
            p.getparent().remove(p)
        for tag in ("w:del", "w:moveFrom"):
            for el in list(raiz.iter(qn(tag))):
                if el.getparent() is not None:
                    el.getparent().remove(el)
        for tag in ("w:ins", "w:moveTo"):
            for el in list(raiz.iter(qn(tag))):
                pai = el.getparent()
                if pai is None:
                    continue
                pos = pai.index(el)
                for filho in list(el):
                    pai.insert(pos, filho)
                    pos += 1
                pai.remove(el)
        for tag in ("w:rPrChange", "w:pPrChange", "w:sectPrChange", "w:tblPrChange"):
            for el in list(raiz.iter(qn(tag))):
                el.getparent().remove(el)
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def tem_alteracoes(dados):
    doc = Document(io.BytesIO(dados))
    body = doc.element.body
    return any(True for _ in body.iter(qn("w:ins"))) or any(True for _ in body.iter(qn("w:del")))
