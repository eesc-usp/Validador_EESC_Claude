"""Verificações próprias dos trabalhos em PDF.

No PDF não há configurações do Word para ler (margens da seção, campo de número de página, estilo da nota):
o que existe é o resultado na página. Estas verificações substituem as do Word que dependem dessas
configurações e medem o que o leitor vê: onde o texto começa e termina em cada página, onde está e que valor
tem o número da página, o tamanho e o espaçamento das notas de rodapé.
"""
import re
import statistics

from . import perfil as PF
from .modelo import AJUSTAR, OK, VERIFICAR, Achado, num
from .pdf import CM, Paragrafo

CAT_FORM = "Formatação"
CAT_EST = "Estrutura"
TOL_LADO = 0.3      # cm: margens esquerda/direita (bem medidas pela coluna de texto)
TOL_ALTURA = 0.45   # cm: margens superior/inferior (a 1ª e a última linha variam um pouco)


def _faixas(nums):
    """[1, 2, 3, 7] -> '1–3, 7' (páginas do PDF, contadas a partir de 1)."""
    nums = sorted(set(nums))
    out, ini = [], None
    for k, n in enumerate(nums):
        if ini is None:
            ini = n
        if k + 1 == len(nums) or nums[k + 1] != n + 1:
            out.append(str(ini) if ini == n else f"{ini}–{n}")
            ini = None
    return ", ".join(out)


def arquivo(d):
    lt = d.pdf
    n_pars = sum(1 for b in lt.blocos if isinstance(b, Paragrafo) and b.linhas)
    n_tab = sum(1 for b in lt.blocos if type(b).__name__ == "Tabela")
    n_fig = sum(1 for b in lt.blocos if type(b).__name__ == "Figura")
    out = [Achado(CAT_EST, "Arquivo em PDF", OK,
                  f"{len(lt.paginas)} páginas lidas: {n_pars} parágrafos, {n_tab} tabela(s), {n_fig} figura(s)"
                  + (" · produzido pelo LaTeX" if lt.tex else ""),
                  "", "—")]
    sem = [p.n + 1 for p in lt.paginas if not p.linhas and p.tem_imagem]
    if sem:
        out.append(Achado(CAT_EST, "Páginas digitalizadas (imagem, sem texto)", VERIFICAR,
                          f"páginas do PDF: {_faixas(sem)}",
                          "O texto dessas páginas é uma imagem e não pode ser verificado. Se for o trabalho todo, "
                          "peça o arquivo original (.docx ou o PDF exportado pelo editor de texto).", "—"))
    return out


def papel_e_margens(d):
    lt = d.pdf
    M = PF.atual()["margens"]
    out = []
    papel = [p.n + 1 for p in lt.paginas
             if not (abs(p.largura / CM - 21) <= 0.15 and abs(p.altura / CM - 29.7) <= 0.15)]
    out.append(Achado(CAT_FORM, "Papel A4", AJUSTAR if papel else OK,
                      f"tamanho diferente de A4 nas páginas do PDF {_faixas(papel)}" if papel else "21 x 29,7 cm",
                      "Layout > Tamanho > A4." if papel else "", "2.2"))

    grupos, erradas, medidas = {}, [], 0
    for p in lt.paginas:
        corpo = [ln for ln in p.linhas if ln.tipo in ("corpo", "nota", "tabela", "figura")]
        texto = [ln for ln in corpo if ln.tipo == "corpo" and abs(ln.size - lt.tam_corpo) < 1.1]
        if len(texto) < 5:
            continue
        medidas += 1
        trocado = lt.lado_trocado(p.n)
        e_esq, e_dir = (M["direita"], M["esquerda"]) if trocado else (M["esquerda"], M["direita"])
        esq = lt.L[p.n] / CM
        cheias = [ln for ln in texto if abs(ln.x1 - lt.R[p.n]) < 2.5]
        mais = max(ln.x1 for ln in corpo)
        dir_ = (p.largura - lt.R[p.n]) / CM if len(cheias) >= 3 else (p.largura - mais) / CM
        dir_confiavel = len(cheias) >= 3
        sup = min(ln.top for ln in corpo) / CM
        inf = (p.altura - max(ln.bottom for ln in corpo)) / CM
        probs = []
        if abs(esq - e_esq) > TOL_LADO:
            probs.append("esq")
        if (dir_confiavel and abs(dir_ - e_dir) > TOL_LADO) or dir_ < e_dir - TOL_LADO:
            probs.append("dir")
        if sup < M["superior"] - TOL_ALTURA:
            probs.append("sup")
        if inf < M["inferior"] - TOL_ALTURA:
            probs.append("inf")
        if probs:
            erradas.append(p.n + 1)
        chave = (round(esq, 1), round(dir_, 1), "sup" in probs, "inf" in probs, trocado, dir_confiavel)
        grupos.setdefault(chave, {"pags": [], "sup": [], "inf": [], "probs": set()})
        g = grupos[chave]
        g["pags"].append(p.n + 1)
        g["sup"].append(sup)
        g["inf"].append(inf)
        g["probs"].update(probs)

    if not medidas:
        out.append(Achado(CAT_FORM, "Margens (cm)", VERIFICAR, "não há páginas com texto corrido suficiente para medir",
                          "Confira as margens no arquivo original.", "2.2"))
        return out

    def cel(v, e, prob, prefixo=""):
        return prefixo + num(v) + ("  ✗" if prob else "")

    linhas = []
    for (esq, dir_, _, _, trocado, conf), g in sorted(grupos.items(), key=lambda x: x[1]["pags"][0]):
        e_esq, e_dir = (M["direita"], M["esquerda"]) if trocado else (M["esquerda"], M["direita"])
        linhas.append({"Páginas do PDF": _faixas(g["pags"]),
                       "Esquerda": cel(esq, e_esq, "esq" in g["probs"]),
                       "Direita": cel(dir_, e_dir, "dir" in g["probs"], "" if conf else "≥ "),
                       "Superior (início do texto)": cel(min(g["sup"]), M["superior"], "sup" in g["probs"]),
                       "Inferior (fim do texto)": cel(min(g["inf"]), M["inferior"], "inf" in g["probs"], "≥ ")})
    tabela = [{"Páginas do PDF": "Esperado", "Esquerda": num(M["esquerda"]), "Direita": num(M["direita"]),
               "Superior (início do texto)": num(M["superior"]), "Inferior (fim do texto)": num(M["inferior"])}] + linhas
    esp = " (margens espelhadas: frente e verso)" if lt.espelhado else ""
    if erradas:
        det = f"texto fora das margens em {len(erradas)} de {medidas} páginas medidas (p. {_faixas(erradas)[:60]})"
    else:
        det = PF.margens_txt() + esp + " · medidas pela posição do texto"
    e3, d2, s3, i2 = (PF.fmt(M[k]) for k in ("esquerda", "direita", "superior", "inferior"))
    out.append(Achado(CAT_FORM, "Margens (cm)", AJUSTAR if erradas else OK, det,
                      (f"No arquivo original: Layout > Margens: esquerda {e3} cm, direita {d2} cm, superior {s3} cm, "
                       f"inferior {i2} cm (aplicar a 'Todo o documento'). No PDF as margens são medidas pela "
                       "posição do texto; a tabela mostra cada grupo de páginas.") if erradas else "",
                      "2.2", tabela=tabela if (erradas or len(linhas) > 1) else []))
    return out


def _inicio_textual(d):
    lt = d.pdf
    if d.i_intro is not None:
        p = lt.pagina_do_bloco(d.i_intro)
        if p is not None:
            return p
    return None


def paginacao(d):
    lt = d.pdf
    nums = [(p.n, p.numero) for p in lt.paginas if p.numero]
    if not nums:
        return [Achado(CAT_FORM, "Paginação", AJUSTAR, "nenhum número de página encontrado",
                       "Inserir número de página no canto superior direito a partir da Introdução.", "2.2.2")]
    intro = _inicio_textual(d)
    probs = []
    rodape = [n + 1 for n, x in nums if x["onde"] == "base"]
    if rodape:
        probs.append(f"número no rodapé (deve ser no cabeçalho, canto superior direito) — p. {_faixas(rodape)[:40]}")
    def lado(n, x):
        if x["x1"] >= lt.R[n] - 0.5 * CM:
            return "dir"
        return "esq" if x["x0"] <= lt.L[n] + 0.5 * CM else "meio"
    lados = {n: lado(n, x) for n, x in nums}
    # frente e verso: no verso (página par) o número fica à esquerda
    pares = [n for n in lados if n % 2 != (nums[0][0] % 2)]
    verso_esq = bool(pares) and sum(lados[n] == "esq" for n in pares) >= 0.7 * len(pares)
    fora = [n + 1 for n, ld in lados.items()
            if not (ld == "dir" or (ld == "esq" and verso_esq and n in pares))]
    if fora:
        probs.append(f"número não alinhado à direita — p. {_faixas(fora)[:40]}")
    if intro is not None:
        pre = [n + 1 for n, _ in nums if n < intro]
        if pre:
            probs.append(f"número aparece em páginas pré-textuais — p. {_faixas(pre)[:40]}")
        sem = [p.n + 1 for p in lt.paginas if p.n >= intro and not p.numero and p.linhas
               and any(ln.tipo == "corpo" for ln in p.linhas)]
        if sem:
            probs.append(f"páginas do texto sem número — p. {_faixas(sem)[:40]}")
    out = [Achado(CAT_FORM, "Paginação", AJUSTAR if probs else OK, "; ".join(probs) or
                  "número no cabeçalho à direita, a partir da Introdução",
                  "Contar todas as folhas desde a folha de rosto, mas exibir o número só a partir da "
                  "Introdução: no arquivo original, inserir 'Quebra de seção (próxima página)' antes da Introdução "
                  "e desvincular o cabeçalho ('Vincular ao anterior' desligado)." if probs else "", "2.2.2")]
    # o número mostrado deve seguir a contagem das folhas (a capa não conta; a folha de rosto é a 1)
    arabicos = [(n, int(x["texto"])) for n, x in nums if x["texto"].isdigit()]
    if len(arabicos) >= 3:
        difs = [v - n for n, v in arabicos]
        base = statistics.mode(difs)
        saltos = [f"p. {n + 1} do PDF mostra {v}" for n, v in arabicos if v - n != base]
        if saltos or base not in (0, 1):
            det = "; ".join(saltos[:4]) if saltos else \
                f"a p. {arabicos[0][0] + 1} do PDF mostra {arabicos[0][1]}"
            out.append(Achado(CAT_FORM, "Número da página segue a contagem das folhas", VERIFICAR, det,
                              "Todas as folhas são contadas a partir da folha de rosto (a capa não conta), inclusive "
                              "as pré-textuais e os versos. Confira se não faltam ou sobram páginas e se a numeração "
                              "não recomeça no meio do trabalho.", "2.2.2"))
        else:
            out.append(Achado(CAT_FORM, "Número da página segue a contagem das folhas", OK,
                              f"{len(arabicos)} números conferidos", "", "2.2.2"))
    return out


def notas_e_pagina(d):
    lt = d.pdf
    lim = PF.atual()["fonte_legenda_max"]
    out = []
    notas = []
    for p in lt.paginas:
        ls = [ln for ln in p.linhas if ln.tipo == "nota"]
        if ls:
            notas.append((p, ls))
    if notas:
        grandes, espacadas, n = [], [], 0
        for p, ls in notas:
            n += sum(1 for ln in ls if re.match(r"^\s*[\d*†]", ln.texto)) or 1
            for ln in ls:
                if ln.size > lim + 0.1:
                    grandes.append(f"[fonte {ln.size:g}, p. {p.n + 1}] {ln.texto}")
            for a, b in zip(ls, ls[1:]):
                if abs(a.size - b.size) < 0.6 and b.base - a.base > 1.3 * 1.15 * a.size and \
                        not re.match(r"^\s*[\d*†]", b.texto):
                    espacadas.append(f"[p. {p.n + 1}] {a.texto}")
        probs = grandes + [f"[entrelinhas > simples] {t}" for t in espacadas]
        out.append(Achado(CAT_FORM, "Notas de rodapé em fonte menor e espaço simples", AJUSTAR if probs else OK,
                          f"{len(probs)} linha(s) de nota fora do padrão" if probs else f"{n} nota(s) conferidas",
                          f"Notas de rodapé em fonte {PF.fmt(min(10, lim))} ou {PF.fmt(lim)} e espaço simples."
                          if probs else "", "2.2.1", probs[:30]))
    grandes = sorted({x["size"] for p in lt.paginas if (x := p.numero) and x["size"] > lim + 0.1})
    if grandes:
        out.append(Achado(CAT_FORM, "Número de página em fonte menor", VERIFICAR,
                          f"número de página em fonte {num(grandes[-1])}",
                          f"A paginação usa fonte menor (10 ou {PF.fmt(lim)}).", "2.2"))
    return out
