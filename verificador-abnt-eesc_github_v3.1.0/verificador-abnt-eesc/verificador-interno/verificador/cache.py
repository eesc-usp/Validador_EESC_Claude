"""Aceleração sem mudar resultados.

O verificador lê cada parágrafo várias vezes (uma por regra). O python-docx, a cada leitura, refaz
trabalho caro: procura o estilo varrendo todos os estilos do arquivo, recompila expressões XPath,
remonta nomes de tags. Este módulo guarda essas respostas. Tudo aqui é *transparente*: com ou sem
ele, os achados, as páginas e os arquivos gerados são idênticos — o teste ``test_cache.py`` roda
o mesmo documento com o cache ligado e desligado e compara tudo.

Camadas:

1. **Biblioteca** (sempre ligada, ver ``instalar()``):
   - estilo de cada id guardado por documento (o verificador nunca cria, apaga nem renomeia estilos;
     um id que não existe não é guardado, então um estilo criado depois seria encontrado);
   - expressões XPath compiladas uma vez por thread (mesma expressão e namespaces → mesmo resultado);
   - nomes de tags (``qn("w:p")`` → ``{...}p``) guardados: é uma função pura.
2. **Leituras durante uma análise** (só dentro de ``leitura()``). Enquanto um documento é analisado
   ele não muda, então a formatação efetiva de cada parágrafo, a lista de referências etc. são
   calculadas uma vez e reaproveitadas por todas as verificações. Fora de ``leitura()`` (ex.: ao gerar o
   arquivo com comentários, que mexe numa cópia do documento) nada disso é guardado.

Para depurar, ``VERIFICADOR_SEM_CACHE=1`` no ambiente desliga tudo.
"""
import contextvars
import os
import threading
from contextlib import contextmanager
from functools import lru_cache, wraps

from docx.oxml import ns as _ns

_ATIVO = contextvars.ContextVar("verificador_cache_leitura", default=None)
_LIGADO = os.environ.get("VERIFICADOR_SEM_CACHE", "").strip() not in ("1", "true", "sim")

# nome de tag com namespace: função pura, guardada (chamada centenas de milhares de vezes)
qn = lru_cache(maxsize=None)(_ns.qn)


# ------------------------------------------------------------------ leituras durante a análise
@contextmanager
def leitura():
    """Liga o cache de leituras enquanto o documento não muda (pode ser aninhado)."""
    if not _LIGADO or _ATIVO.get() is not None:
        yield
        return
    tok = _ATIVO.set({})
    try:
        yield
    finally:
        _ATIVO.reset(tok)


def ativo():
    return _ATIVO.get() is not None


def _xml(obj):
    for atributo in ("_p", "_r", "_tbl", "_element"):
        el = getattr(obj, atributo, None)
        if el is not None:
            return el
    return obj


def _hashavel(x):
    """Objetos do python-docx (Document, Paragraph...) não são hasháveis: usa o XML deles."""
    try:
        hash(x)
        return x
    except TypeError:
        return _xml(x)


def memo(chave, calcular):
    """Valor guardado para 'chave' durante a análise (ou calculado na hora, fora dela)."""
    c = _ATIVO.get()
    if c is None:
        return calcular()
    k = tuple(_hashavel(x) for x in chave) if isinstance(chave, tuple) else _hashavel(chave)
    try:
        return c[k]
    except KeyError:
        v = c[k] = calcular()
        return v


def por_elemento(fn):
    """Memoriza fn(paragrafo_ou_run, *args) pela identidade do XML, só durante a análise."""
    nome = fn.__module__ + "." + fn.__qualname__

    @wraps(fn)
    def embrulho(obj, *args):
        c = _ATIVO.get()
        if c is None:
            return fn(obj, *args)
        k = (nome, _xml(obj), *(_hashavel(a) for a in args))
        try:
            return c[k]
        except KeyError:
            v = c[k] = fn(obj, *args)
            return v

    embrulho.sem_cache = fn
    return embrulho


# ------------------------------------------------------------------ estilos do Word (por documento)
def _guardado(part):
    """Dicionário de respostas guardado no próprio objeto do documento (morre junto com ele)."""
    if not _LIGADO:
        return None
    try:
        d = part.__dict__
    except AttributeError:
        return None
    c = d.get("_verificador_cache")
    if c is None:
        c = d.setdefault("_verificador_cache", {})
    return c


def id_do_estilo(el):
    """Id do estilo aplicado direto no parágrafo/trecho (w:pStyle / w:rStyle), ou None."""
    c = _ATIVO.get()
    if c is None:
        return el.style
    k = ("id_estilo", el)
    try:
        return c[k]
    except KeyError:
        v = c[k] = el.style
        return v


def cadeia(obj, estilo_de):
    """Estilo de um parágrafo/trecho e seus estilos-base, guardados por documento.

    'estilo_de' é a função que lê o estilo do objeto (ex.: lambda p: p.style)."""
    from .docx_util import _cadeia
    try:
        part = obj.part
        el = obj._p if hasattr(obj, "_p") else obj._r
        sid = id_do_estilo(el)
    except Exception:
        return tuple(_cadeia(estilo_de(obj)))
    c = _guardado(part)
    if c is None:
        return tuple(_cadeia(estilo_de(obj)))
    k = ("cadeia", el.tag, sid)
    v = c.get(k)
    if v is None:
        v = tuple(_cadeia(estilo_de(obj)))
        if sid is None or (v and getattr(v[0], "style_id", None) == sid):
            c[k] = v
    return v


# ------------------------------------------------------------------ instalação na biblioteca
_ORIGINAIS = {}
_LOCAL = threading.local()


def _get_style_com_cache(original):
    @wraps(original)
    def get_style(self, style_id, style_type):
        c = _guardado(self)
        if c is None:
            return original(self, style_id, style_type)
        k = ("estilo", style_id, style_type)
        try:
            return c[k]
        except KeyError:
            pass
        estilo = original(self, style_id, style_type)
        # guarda o padrão do tipo e os estilos encontrados; um id inexistente não é guardado
        if style_id is None or (estilo is not None and getattr(estilo, "style_id", None) == style_id):
            c[k] = estilo
        return estilo
    return get_style


def _xpath_compilado(original):
    from lxml import etree

    @wraps(original)
    def xpath(self, xpath_str):
        try:
            compiladas = _LOCAL.xpaths
        except AttributeError:
            compiladas = _LOCAL.xpaths = {}
        xp = compiladas.get(xpath_str)
        if xp is None:
            try:
                xp = compiladas[xpath_str] = etree.XPath(xpath_str, namespaces=_ns.nsmap)
            except Exception:
                return original(self, xpath_str)
        return xp(self)
    return xpath


def instalar():
    """Liga as acelerações na biblioteca (idempotente). Sem efeito com VERIFICADOR_SEM_CACHE=1."""
    if not _LIGADO or _ORIGINAIS:
        return
    try:
        from docx.oxml import xmlchemy
        from docx.oxml.xmlchemy import BaseOxmlElement
        from docx.parts.document import DocumentPart
    except Exception:  # pragma: no cover - python-docx muito diferente: segue sem acelerar
        return
    _ORIGINAIS["get_style"] = DocumentPart.get_style
    _ORIGINAIS["xpath"] = BaseOxmlElement.xpath
    DocumentPart.get_style = _get_style_com_cache(DocumentPart.get_style)
    BaseOxmlElement.xpath = _xpath_compilado(BaseOxmlElement.xpath)
    if getattr(xmlchemy, "qn", None) is _ns.qn:
        _ORIGINAIS["xmlchemy.qn"] = xmlchemy.qn
        xmlchemy.qn = qn


def _desinstalar():
    from docx.oxml import xmlchemy
    from docx.oxml.xmlchemy import BaseOxmlElement
    from docx.parts.document import DocumentPart
    if "get_style" in _ORIGINAIS:
        DocumentPart.get_style = _ORIGINAIS.pop("get_style")
    if "xpath" in _ORIGINAIS:
        BaseOxmlElement.xpath = _ORIGINAIS.pop("xpath")
    if "xmlchemy.qn" in _ORIGINAIS:
        xmlchemy.qn = _ORIGINAIS.pop("xmlchemy.qn")


@contextmanager
def desligado():
    """Roda o bloco sem nenhuma aceleração (para os testes compararem). Não usar em produção:
    vale para o processo inteiro enquanto o bloco roda."""
    global _LIGADO
    antes = _LIGADO
    _desinstalar()
    _LIGADO = False
    tok = _ATIVO.set(None)
    try:
        yield
    finally:
        _ATIVO.reset(tok)
        _LIGADO = antes
        instalar()


# compatibilidade com o nome antigo
instalar_cache_de_estilos = instalar
