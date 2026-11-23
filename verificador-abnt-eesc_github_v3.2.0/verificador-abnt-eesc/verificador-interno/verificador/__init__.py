"""Verificador ABNT – Biblioteca EESC-USP."""
from . import cache as _cache
from . import elementos, estrutura, extras, referencias, tabelas, tipografia, trabalho
from . import perfil as _perfil
from .documento import Documento
from .modelo import AJUSTAR, OK, VERIFICAR, Achado

__all__ = ["analisar", "Achado", "OK", "AJUSTAR", "VERIFICAR"]


def analisar(arquivo, perfil_norma=None):
    """Roda todas as verificações por regra, segundo o perfil de norma (padrão: EESC).

    Retorna (achados, referências extraídas)."""
    with _perfil.usar(perfil_norma), _cache.leitura():
        return _analisar(arquivo)


def _analisar(arquivo):
    d = Documento(arquivo)
    if d.pdf is not None and d.pdf.sem_texto:
        return [Achado("Estrutura", "PDF sem texto (digitalizado)", AJUSTAR,
                       "o arquivo só tem imagens das páginas: não há texto para verificar",
                       "Peça o arquivo original (.docx) ou o PDF exportado pelo editor de texto, não uma digitalização.",
                       "—")], []
    achados = estrutura.verificar(d)
    achados += extras.verificar(d)
    try:
        achados += trabalho.verificar(d)
    except Exception as e:
        achados.append(Achado("Sistema", "Tipo de trabalho", VERIFICAR, f"verificação não concluída: {e}"))
    try:
        achados += tabelas.verificar(d)[0]
    except Exception as e:
        achados.append(Achado("Sistema", "tabelas", VERIFICAR, f"verificação não concluída: {e}"))
    try:
        a, refs = referencias.verificar(d)
    except Exception as e:
        a, refs = [Achado("Sistema", "referências", VERIFICAR, f"verificação não concluída: {e}")], []
    achados += a
    achados += elementos.verificar(d)
    achados += tipografia.verificar(d)
    try:
        linhas, extra = elementos.resumo_elementos(d, achados)
        achados += extra
        achados.append(Achado(elementos.CAT, RESUMO_ELEMENTOS, OK,
                              f"{sum(x['Aparece'].startswith('✅') for x in linhas)} elementos encontrados",
                              tabela=linhas))
    except Exception as e:
        achados.append(Achado("Sistema", "elementos", VERIFICAR, f"verificação não concluída: {e}"))
    return achados, refs


RESUMO_ELEMENTOS = "Elementos: aparece · formatação · conteúdo"
