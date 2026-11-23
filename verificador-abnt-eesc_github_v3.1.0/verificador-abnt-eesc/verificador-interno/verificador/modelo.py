"""Estrutura comum dos achados da verificação."""
from dataclasses import dataclass, field

OK = "OK"
AJUSTAR = "Ajustar"
VERIFICAR = "Verificar"  # o código não tem certeza; o bibliotecário confere


@dataclass
class Achado:
    categoria: str          # ex.: "Formatação", "Referências"
    item: str               # o que foi verificado
    status: str             # OK / Ajustar / Verificar
    detalhe: str = ""       # o que foi encontrado
    como_corrigir: str = ""
    secao: str = ""         # seção das Diretrizes USP
    trechos: list = field(default_factory=list)  # exemplos do documento
    origem: str = "regra"   # "regra" ou "ia"
    tabela: list = field(default_factory=list)   # linhas (dict) para mostrar como tabela
    blocos: list = field(default_factory=list)   # índices de blocos do documento (localização direta)


class Trecho(str):
    """Texto de exemplo de um achado que sabe de onde veio.

    Funciona como um texto comum (listas, relatório, comentários), mas guarda o bloco do documento (``i``)
    e, se for o caso, o pedaço exato a marcar (``m``). Assim o visualizador aponta o lugar certo, sem
    procurar o texto de novo (o mesmo texto pode aparecer no sumário, numa lista ou repetido).
    """

    def __new__(cls, texto, i=None, m=None):
        t = super().__new__(cls, texto)
        t.i, t.m = i, m
        return t

    def __copy__(self):
        return self

    def __deepcopy__(self, memo):
        return self

    def __reduce__(self):
        return (Trecho, (str(self), self.i, self.m))


def T(texto, bloco, marca=None):
    """Trecho ligado ao bloco (objeto com ``.i`` ou o índice) e, opcionalmente, ao pedaço a marcar."""
    i = getattr(bloco, "i", bloco)
    return Trecho(texto, i if isinstance(i, int) else None, marca)


def num(v, casas=2):
    """Número no formato brasileiro, sem zeros sobrando: 2.5 -> '2,5'."""
    if v is None:
        return "-"
    t = f"{v:.{casas}f}".rstrip("0").rstrip(".")
    return t.replace(".", ",")
