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


def num(v, casas=2):
    """Número no formato brasileiro, sem zeros sobrando: 2.5 -> '2,5'."""
    if v is None:
        return "-"
    t = f"{v:.{casas}f}".rstrip("0").rstrip(".")
    return t.replace(".", ",")
