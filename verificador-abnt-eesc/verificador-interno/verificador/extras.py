"""Verificações cruzadas (como fazem os verificadores de tese) e regras personalizadas do perfil.

- Sumário × títulos do texto (mesma numeração e mesmas palavras; títulos faltando no sumário)
- Listas de ilustrações/tabelas × títulos das legendas
- Toda ilustração citada no texto (Figura 3, Gráfico 2...)
- Regras personalizadas do perfil de norma (proibir/exigir textos)
"""
import re

from . import docx_util as du
from . import perfil as PF
from .documento import RE_APENDICE, TITULOS_SEM_NUMERO, chave
from .estrutura import RE_ILU, RE_SECAO, RE_SUMARIO, RE_TAB, _amostra, _candidato_titulo
from .modelo import AJUSTAR, OK, VERIFICAR, Achado

CAT_CRUZ = "Sumário e listas"
CAT_REGRA = "Regras personalizadas"


def _limpo(t):
    """Chave de comparação de título: sem pontilhado/nº de página, sem acento, maiúsculas, sem pontuação."""
    t = re.sub(r"(\.{3,}|…+|\t).*$", "", t)
    t = re.sub(r"\s+\d+\s*$", "", t)
    return re.sub(r"[^A-Z0-9 ]", "", chave(t)).strip()


def _secao_pre(d, titulo_re):
    """Parágrafos da seção pré-textual cujo título casa titulo_re (até o próximo título sem número)."""
    ini = None
    for b in d.blocos:
        if b.tipo == "p" and b.texto and re.fullmatch(titulo_re, chave(b.texto)):
            ini = b.i
            break
    if ini is None:
        return None
    out = []
    for b in d.blocos[ini + 1:]:
        if b.tipo != "p" or not b.texto:
            continue
        k = chave(b.texto)
        if (k in TITULOS_SEM_NUMERO or k.startswith("LISTA DE")) and not RE_SUMARIO.search(b.texto) \
                and "\t" not in b.obj.text:
            break
        if d.i_intro is not None and b.i >= d.i_intro:
            break
        out.append(b)
        if len(out) > 400:
            break
    return out


def sumario(d):
    linhas = _secao_pre(d, r"SUMARIO")
    if linhas is None:
        return []
    entradas = {}
    sem_num = set()
    for b in linhas:
        t = b.obj.text.strip() or b.texto
        m = re.match(r"^(\d{1,2}(?:\.\d{1,2}){0,6})\.?\s*[-–—:]?\s*(.+)$", t)
        if m:
            entradas[m.group(1)] = (_limpo(m.group(2)), b)
        else:
            sem_num.add(_limpo(t))
    if not entradas:
        return _sumario_sem_numeros(d, sem_num)
    titulos = {}
    for b in d.paragrafos():
        if d.regiao(b.i) == "pre":
            continue
        m = _candidato_titulo(b)
        if m and m != "auto":
            titulos[m.group(1)] = (_limpo(m.group(4)), b)
    # títulos que o detector não reconheceu pela formatação, mas que existem como parágrafo "2.1.1 Título"
    for num in entradas:
        if num in titulos:
            continue
        rx = re.compile(rf"^{re.escape(num)}\.?\s+(\S.*)$")
        for b in d.paragrafos():
            if d.regiao(b.i) != "pre" and len(b.texto) < 150:
                m2 = rx.match(b.texto)
                if m2:
                    titulos[num] = (_limpo(m2.group(1)), b)
                    break
    diferentes, faltando_sum, faltando_txt = [], [], []
    for num, (tit, b) in titulos.items():
        if num not in entradas:
            if num.count(".") <= 2:  # até a seção terciária deve aparecer
                faltando_sum.append((f"{num} {b.texto[len(num):].strip()}", b.i))
        elif entradas[num][0] != tit:
            diferentes.append((f"{num}: sumário “{entradas[num][1].texto[:60]}” × texto “{b.texto[:60]}”", b.i))
    for num, (tit, b) in entradas.items():
        if titulos and num not in titulos:
            faltando_txt.append((f"{num} {tit.title()[:70]}", b.i))
    out = []
    for item, lista, corr in [
        ("Sumário com as mesmas palavras dos títulos", diferentes,
         "O título no sumário deve ser idêntico ao do texto (mesma numeração e redação). Atualize o sumário "
         "automático (clique nele > F9 > Atualizar o índice inteiro)."),
        ("Todos os títulos aparecem no sumário", faltando_sum,
         "Incluir no sumário todas as seções (até a terciária, ao menos). Se o sumário é automático, "
         "aplique o estilo de Título correspondente e atualize-o."),
        ("Sumário sem títulos que não existem no texto", faltando_txt,
         "O sumário cita seções que não foram encontradas no texto: provavelmente está desatualizado (F9)."),
    ]:
        out.append(Achado(CAT_CRUZ, item, AJUSTAR if lista else OK,
                          f"{len(lista)} ocorrência(s)" if lista else "confere",
                          corr if lista else "", "3.1.2.1.12", _amostra([x for x, _ in lista]),
                          blocos=[i for _, i in lista[:300]]))
    return out


def _parecido(a, b):
    return bool(a) and bool(b) and (a == b or a.startswith(b) or b.startswith(a))


def _sumario_sem_numeros(d, sem_num):
    """Sumário cujos números vêm da numeração automática do Word: compara só as palavras dos títulos."""
    sem_num = {x for x in sem_num if x and x not in {_limpo(t) for t in TITULOS_SEM_NUMERO} and
               not RE_APENDICE.match(x)}
    if not sem_num:
        return [Achado(CAT_CRUZ, "Sumário confere com os títulos", VERIFICAR,
                       "não foi possível ler as linhas do sumário",
                       "Use o sumário automático do Word (Referências > Sumário) e atualize antes de entregar.",
                       "3.1.2.1.12")]
    titulos = []
    for b in d.paragrafos():
        if d.regiao(b.i) == "pre":
            continue
        m = _candidato_titulo(b)
        if m == "auto":
            titulos.append((_limpo(b.texto), b))
        elif m:
            titulos.append((_limpo(m.group(4)), b))
    textos = [t for t, _ in titulos] + [_limpo(b.texto) for b in d.paragrafos()
                                         if d.regiao(b.i) != "pre" and b.texto and len(b.texto) < 120]
    falta_txt = sorted(x for x in sem_num if not any(_parecido(x, t) for t in textos))
    falta_sum = [(b.texto, b.i) for t, b in titulos
                 if t not in {_limpo(x) for x in TITULOS_SEM_NUMERO} and not any(_parecido(t, x) for x in sem_num)]
    out = [Achado(CAT_CRUZ, "Todos os títulos aparecem no sumário", VERIFICAR if falta_sum else OK,
                  f"{len(falta_sum)} título(s) do texto não achados no sumário" if falta_sum else "confere",
                  "Incluir no sumário todas as seções; se é automático, aplique o estilo de Título e atualize (F9)."
                  if falta_sum else "", "3.1.2.1.12", _amostra([x for x, _ in falta_sum]),
                  blocos=[i for _, i in falta_sum[:300]]),
           Achado(CAT_CRUZ, "Sumário sem títulos que não existem no texto", VERIFICAR if falta_txt else OK,
                  f"{len(falta_txt)} linha(s) do sumário sem título correspondente" if falta_txt else "confere",
                  "O sumário cita títulos não encontrados no texto (ou com outra redação): atualize-o (F9)."
                  if falta_txt else "", "3.1.2.1.12", [x.title() for x in falta_txt[:300]])]
    return out


def _legendas(d):
    """{('Figura','3'): (titulo_limpo, bloco)} das legendas no texto."""
    out = {}
    for b in d.paragrafos():
        if d.regiao(b.i) == "pre":
            continue
        m = RE_ILU.match(b.texto)
        if m:
            out[(m.group(1).capitalize().replace("Grafico", "Gráfico"), m.group(2))] = (_limpo(m.group(4)), b)
            continue
        m = RE_TAB.match(b.texto)
        if m:
            out[("Tabela", m.group(1))] = (_limpo(m.group(3)), b)
    return out


def listas(d):
    legendas = _legendas(d)
    out = []
    for nome_lista, tipos in [("LISTA DE (FIGURAS|ILUSTRACOES)", None), ("LISTA DE TABELAS", {"Tabela"}),
                              ("LISTA DE QUADROS", {"Quadro"}), ("LISTA DE GRAFICOS", {"Gráfico"})]:
        linhas = _secao_pre(d, nome_lista)
        if linhas is None:
            continue
        rotulo = {"Tabela": "tabelas", "Quadro": "quadros", "Gráfico": "gráficos"}.get(next(iter(tipos)) if tipos else "",
                                                                                     "ilustrações")
        entradas = {}
        for b in linhas:
            t = b.obj.text.strip() or b.texto
            m = re.match(r"^(Figura|Quadro|Gr[áa]fico|Tabela|Fotografia|Mapa|Esquema|Fluxograma|Desenho|"
                         r"Organograma|Planta|Imagem)\s+(\d+(?:\.\d+)?)\s*[-–—:.]?\s*(.*)$", t, re.I)
            if m:
                tipo = m.group(1).capitalize().replace("Grafico", "Gráfico")
                entradas[(tipo, m.group(2))] = (_limpo(m.group(3)), b)
        alvo = {k: v for k, v in legendas.items() if (k[0] in tipos if tipos else k[0] not in ("Tabela",))}
        if tipos is None:
            # lista geral de figuras/ilustrações: compara só os tipos que aparecem nela
            tipos_na_lista = {k[0] for k in entradas} or {"Figura"}
            alvo = {k: v for k, v in alvo.items() if k[0] in tipos_na_lista}
        dif, falta = [], []
        for k, (tit, b) in alvo.items():
            if k not in entradas:
                falta.append((f"{k[0]} {k[1]} – {b.texto.split(k[1], 1)[-1].strip(' -–—:.')[:60]}", b.i))
            elif entradas[k][0] and tit and not _parecido(entradas[k][0], tit):
                dif.append((f"{k[0]} {k[1]}: lista “{entradas[k][1].texto[:50]}” × legenda “{b.texto[:50]}”", b.i))
        sobra = [(f"{k[0]} {k[1]}", v[1].i) for k, v in entradas.items() if alvo and k not in alvo]
        lst = falta + dif + sobra
        out.append(Achado(CAT_CRUZ, f"Lista de {rotulo} confere com as legendas", AJUSTAR if lst else OK,
                          (f"{len(falta)} faltando, {len(dif)} com título diferente, {len(sobra)} sem legenda no texto"
                           if lst else f"{len(entradas)} item(ns) conferidos"),
                          "A lista deve ter todos os itens, na ordem, com o mesmo título da legenda e a página. "
                          "Use legendas do Word (Referências > Inserir legenda) e o índice de ilustrações "
                          "automático, atualizando-o com F9." if lst else "", "3.1.2.1.10",
                          _amostra([x for x, _ in lst]), blocos=[i for _, i in lst[:300]]))
    return out


def ilustracoes_citadas(d):
    """Toda ilustração (exceto tabelas/quadros, vistos em tabelas.py) deve ser citada no texto."""
    legendas = {k: v for k, v in _legendas(d).items() if k[0] not in ("Tabela", "Quadro")}
    if not legendas:
        return []
    texto = []
    for b in d.paragrafos("textual"):
        if RE_ILU.match(b.texto) or re.match(r"^Fonte\b", b.texto):
            continue
        texto.append(b.texto)
    corpo = " ".join(texto)
    citadas = set()
    for m in re.finditer(r"\b(Figuras?|Fig\.|Gr[áa]ficos?|Fotografias?|Mapas?|Esquemas?|Fluxogramas?|"
                         r"Desenhos?|Organogramas?|Plantas?|Imagens?)\s+((?:\d+(?:\.\d+)?)(?:\s*(?:,|e|a|-|–)\s*"
                         r"\d+(?:\.\d+)?)*)", corpo, re.I):
        tipo = m.group(1).rstrip("s").replace("Fig.", "Figura").capitalize().replace("Grafico", "Gráfico")
        tipo = {"Gráfico": "Gráfico"}.get(tipo, tipo)
        nums = re.findall(r"\d+(?:\.\d+)?", m.group(2))
        if re.search(r"\s(a|-|–)\s*\d", m.group(2)) and len(nums) == 2 and "." not in "".join(nums):
            nums = [str(n) for n in range(int(nums[0]), int(nums[1]) + 1)]
        for n in nums:
            citadas.add((tipo, n))
    nao = [(f"{k[0]} {k[1]}", b.i) for k, (_, b) in legendas.items() if k not in citadas]
    return [Achado(CAT_CRUZ, "Toda ilustração citada no texto", VERIFICAR if nao else OK,
                   f"{len(nao)} de {len(legendas)} sem menção no texto" if nao else f"{len(legendas)} citadas",
                   "Toda ilustração deve ser mencionada no texto (ex.: 'conforme a Figura 3') e inserida o mais "
                   "perto possível do trecho que a cita." if nao else "", "2.6",
                   [x for x, _ in nao], blocos=[i for _, i in nao])]


# ------------------------------------------------------------------ regras personalizadas
def _unidades(d, onde):
    if onde == "referencias":
        from .referencias import extrair_referencias
        return [(r.texto, r.i) for r in extrair_referencias(d)]
    if onde == "texto":
        return [(b.texto, b.i) for b in d.paragrafos("textual") if b.texto]
    return [(b.texto, b.i) for b in d.blocos if b.tipo == "p" and b.texto]


def regras_personalizadas(d):
    out = []
    for r in PF.atual().get("regras") or []:
        nome = (r.get("nome") or r.get("padrao") or "").strip()
        padrao = (r.get("padrao") or "").strip()
        if not padrao or not nome:
            continue
        try:
            rx = re.compile(padrao if r.get("regex") else re.escape(padrao), 0 if r.get("maiusculas") else re.I)
        except re.error as e:
            out.append(Achado(CAT_REGRA, nome, VERIFICAR, f"padrão inválido: {e}", "Corrija a regra no perfil."))
            continue
        status = AJUSTAR if r.get("status", "Ajustar") == "Ajustar" else VERIFICAR
        corr = r.get("como_corrigir") or ""
        unidades = _unidades(d, r.get("onde", "texto"))
        tipo = r.get("tipo", "proibir")
        if tipo == "proibir":
            achados = [(t, i) for t, i in unidades if rx.search(t)]
            out.append(Achado(CAT_REGRA, nome, status if achados else OK,
                              f"{len(achados)} ocorrência(s)" if achados else "nenhuma ocorrência",
                              corr if achados else "", "perfil", _amostra([t for t, _ in achados]),
                              blocos=[i for _, i in achados[:300]]))
        elif tipo == "exigir":
            ok = any(rx.search(t) for t, _ in unidades)
            out.append(Achado(CAT_REGRA, nome, OK if ok else status, "encontrado" if ok else "não encontrado",
                              "" if ok else corr, "perfil"))
        else:  # exigir_cada
            faltam = [(t, i) for t, i in unidades if not rx.search(t)]
            out.append(Achado(CAT_REGRA, nome, status if faltam else OK,
                              f"{len(faltam)} de {len(unidades)} sem o padrão" if faltam else "todos conferem",
                              corr if faltam else "", "perfil", _amostra([t for t, _ in faltam]),
                              blocos=[i for _, i in faltam[:300]]))
    return out


def verificar(d):
    achados = []
    for f in (sumario, listas, ilustracoes_citadas, regras_personalizadas):
        try:
            achados += f(d)
        except Exception as e:
            achados.append(Achado("Sistema", f.__name__, VERIFICAR, f"verificação não concluída: {e}"))
    return achados
