"""Retrato dos resultados: garante que uma mudança no código não altera nada sem querer.

Para cada documento público (exemplos, template da EESC e os sintéticos de ``sinteticos.py``),
registra tudo o que o verificador produz: os achados nas duas normas, as referências, a página
desenhada, os locais dos achados, o original com apontamentos, o relatório e as duas revisões
automáticas (com e sem controle de alterações). O teste ``test_retrato.py`` compara com o
``retrato_esperado.json``.

Se uma mudança nas regras foi **intencional**, atualize o retrato e confira a diferença no commit:

    python tests/retrato.py --atualizar

Sem argumentos, só mostra o que mudou.
"""
import hashlib
import io
import json
import re
import sys
import zipfile
from dataclasses import asdict
from pathlib import Path

PASTA = Path(__file__).resolve().parent
RAIZ = PASTA.parent
sys.path.insert(0, str(RAIZ))
sys.path.insert(0, str(PASTA))

ESPERADO = PASTA / "retrato_esperado.json"
TEMPLATE = RAIZ / "assets" / "modelos" / "TEMPLATE_Dissertacao_Tese_ABNT_EESC.docx"
_VOLATEIS = ("date", "rsidR", "rsidRDefault", "rsidP", "rsidRPr", "rsidDel", "rsidSect", "rsidTr",
             "paraId", "textId", "initials")


def _sha(b):
    return hashlib.sha1(b).hexdigest()[:16]


def xml_canonico(dados, parte="word/document.xml"):
    """Hash do XML de uma parte do .docx sem o que muda a cada execução (datas, rsid, ids de parágrafo).

    Usa a forma canônica do XML (C14N), que não depende da versão da biblioteca que salvou."""
    from lxml import etree
    with zipfile.ZipFile(io.BytesIO(dados)) as z:
        if parte not in z.namelist():
            return None
        raiz = etree.fromstring(z.read(parte))
    for el in raiz.iter():
        for nome in list(el.attrib):
            if nome.rsplit("}", 1)[-1] in _VOLATEIS:
                del el.attrib[nome]
        if el.text and re.fullmatch(r".*Data: \d\d/\d\d/\d{4}.*", el.text, re.S):
            el.text = re.sub(r"\d\d/\d\d/\d{4}", "dd/mm/aaaa", el.text)
    return _sha(etree.tostring(raiz, method="c14n2"))


def corpus():
    """[(nome, bytes)] dos documentos públicos, em ordem fixa."""
    import sinteticos
    from test_fora_do_modelo import documento_fora_do_modelo
    docs = [(p.name, p.read_bytes()) for p in sorted((RAIZ / "exemplos").glob("*.docx"))]
    docs.append((TEMPLATE.name, TEMPLATE.read_bytes()))
    docs.append(("fora_do_modelo.docx", documento_fora_do_modelo()))
    docs += [(f"{nome}.docx", dados) for nome, dados in sinteticos.todos().items()]
    return docs


def _achados(lista):
    """Achados em forma compacta: o que se lê por inteiro e um resumo (hash) do resto."""
    out = []
    for a in lista:
        x = asdict(a)
        resto = json.dumps([x["trechos"], x["blocos"], x["como_corrigir"], x["secao"], x["origem"], x["tabela"]],
                           ensure_ascii=False, sort_keys=True, default=str)
        out.append([x["categoria"], x["item"], x["status"], x["detalhe"][:200], [t[:80] for t in x["trechos"][:2]],
                    len(x["trechos"]), _sha(resto.encode() + x["detalhe"].encode())])
    return out


def retrato_de(nome, dados):
    from verificador import analisar, anotacao, relatorio
    from verificador import perfil as PF
    from verificador.correcao import GRUPOS, corrigir_duplo, texto_atual
    from verificador.documento import Documento
    from verificador.edicao import EdicaoRecusada, aceitar_tudo, editar
    from verificador.trabalho import detectar
    from verificador.visualizacao import renderizar, serializar

    r = {}
    tipo = detectar(Documento(io.BytesIO(dados)))[0]
    P = PF.PADROES.get(tipo, PF.EESC)
    r["tipo"] = tipo
    base = None
    for rotulo, perfil in (("norma_detectada", P), ("outra_norma", PF.EESC_POS if P is PF.EESC else PF.EESC)):
        achados, refs = analisar(io.BytesIO(dados), perfil)
        if base is None:
            base = achados
            r[rotulo] = _achados(achados)
            r["referencias"] = [x.texto for x in refs]
        else:  # a outra norma entra resumida: contagem por situação e um hash de tudo
            contagem = {}
            for a in achados:
                contagem[a.status] = contagem.get(a.status, 0) + 1
            r[rotulo] = [contagem, _sha(json.dumps(_achados(achados), ensure_ascii=False).encode())]
    html_doc, d = renderizar(io.BytesIO(dados))
    r["pagina"] = _sha(html_doc.encode())
    r["locais"] = _sha(json.dumps(serializar(d, base), sort_keys=True, default=str).encode())
    anotado, n = anotacao.gerar(dados, base)
    r["apontamentos"] = [n, xml_canonico(anotado), xml_canonico(anotado, "word/comments.xml")]
    r["relatorio"] = xml_canonico(relatorio.gerar(base, nome))
    for rastrear in (True, False):
        novo, log, falhas, passadas = corrigir_duplo(dados, list(GRUPOS), rastrear, None, P)
        r[f"correcao_{'rastreada' if rastrear else 'direta'}"] = {
            "log": log, "falhas": falhas, "passadas": passadas,
            "arquivo": xml_canonico(novo), "aceito": xml_canonico(aceitar_tudo(novo))}
    alvo = next((b for b in d.blocos if b.tipo == "p" and len(b.texto) > 60), None)
    if alvo is not None:
        orig = texto_atual(alvo.obj)
        try:
            saida, aviso = editar(dados, str(alvo.i), orig.replace(" ", "  ", 1) + " (editado)", orig, True)
            r["edicao"] = [xml_canonico(saida), aviso]
        except EdicaoRecusada as e:  # recusar com segurança também é um resultado
            r["edicao"] = ["recusada", str(e)]
    return r


def calcular():
    return {nome: retrato_de(nome, dados) for nome, dados in corpus()}


def diferencas(esperado, atual, caminho=""):
    """Lista legível das diferenças entre dois retratos."""
    out = []
    if isinstance(esperado, dict) and isinstance(atual, dict):
        for k in sorted(set(esperado) | set(atual), key=str):
            if k not in atual:
                out.append(f"{caminho}/{k}: sumiu")
            elif k not in esperado:
                out.append(f"{caminho}/{k}: novo")
            else:
                out += diferencas(esperado[k], atual[k], f"{caminho}/{k}")
    elif isinstance(esperado, list) and isinstance(atual, list) and len(esperado) == len(atual):
        for i, (a, b) in enumerate(zip(esperado, atual)):
            out += diferencas(a, b, f"{caminho}[{i}]")
    elif esperado != atual:
        out.append(f"{caminho}: {json.dumps(esperado, ensure_ascii=False)[:160]}  →  "
                   f"{json.dumps(atual, ensure_ascii=False)[:160]}")
    return out


def _formatar(retrato):
    """JSON legível e de diff limpo: um documento por bloco, um achado ou item de log por linha."""
    def compacto(v):
        return json.dumps(v, ensure_ascii=False, sort_keys=True)
    linhas = ["{"]
    for i, (doc, campos) in enumerate(sorted(retrato.items())):
        linhas.append(f" {compacto(doc)}: {{")
        itens = sorted(campos.items())
        for j, (k, v) in enumerate(itens):
            fim = "," if j < len(itens) - 1 else ""
            if isinstance(v, list) and v and isinstance(v[0], list) and k == "norma_detectada":
                corpo = ",\n".join(f"   {compacto(x)}" for x in v)
                linhas.append(f"  {compacto(k)}: [\n{corpo}\n  ]{fim}")
            else:
                linhas.append(f"  {compacto(k)}: {compacto(v)}{fim}")
        linhas.append(" }" + ("," if i < len(retrato) - 1 else ""))
    linhas.append("}")
    return "\n".join(linhas) + "\n"


def main():
    atual = json.loads(json.dumps(calcular(), ensure_ascii=False, default=str))
    if "--atualizar" in sys.argv:
        ESPERADO.write_text(_formatar(atual), "utf-8")
        print(f"retrato atualizado: {ESPERADO.name} ({len(atual)} documentos)")
        return 0
    if not ESPERADO.exists():
        print("ainda não há retrato: rode com --atualizar")
        return 1
    difs = diferencas(json.loads(ESPERADO.read_text("utf-8")), atual)
    print("\n".join(difs) if difs else "nenhuma diferença")
    return 1 if difs else 0


if __name__ == "__main__":
    sys.exit(main())
