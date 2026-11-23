"""Revisão das referências por IA, com vários provedores.

Provedores:
- ollama     IA local (nada sai do computador)
- gemini     Google Gemini (chave do Google AI Studio)
- openai     OpenAI / ChatGPT (chave da plataforma OpenAI)
- anthropic  Anthropic / Claude (chave do Claude Console)
- compat     qualquer serviço compatível com a API da OpenAI (OpenRouter, Groq, DeepSeek, Mistral,
             LM Studio...), informando o endereço base

Só o TEXTO DE CADA REFERÊNCIA é enviado ao provedor, junto com as regras da ABNT. O trabalho do aluno
não é enviado.
"""
import json
import re
import time
from concurrent.futures import CancelledError, ThreadPoolExecutor, as_completed

import requests

from .ia_local import MODELOS, PROMPT, REGRAS_GERAIS
from .modelo import AJUSTAR, OK, VERIFICAR, Achado

PROVEDORES = {
    "ollama": "IA local (Ollama) — gratuita, nada sai do computador",
    "gemini": "Google Gemini",
    "openai": "OpenAI (ChatGPT)",
    "anthropic": "Anthropic (Claude)",
    "compat": "Outro compatível com OpenAI (OpenRouter, Groq, DeepSeek, Mistral...)",
}
NUVEM = {"gemini", "openai", "anthropic", "compat"}
URLS_PADRAO = {
    "ollama": "http://localhost:11434",
    "gemini": "https://generativelanguage.googleapis.com/v1beta",
    "openai": "https://api.openai.com/v1",
    "anthropic": "https://api.anthropic.com/v1",
    "compat": "https://openrouter.ai/api/v1",
}
ONDE_PEGAR_CHAVE = {
    "gemini": "https://aistudio.google.com/apikey (a assinatura do app Gemini não vale para a API: sem "
              "faturamento ativado, a chave fica no plano gratuito, com poucos pedidos por dia)",
    "openai": "https://platform.openai.com/api-keys",
    "anthropic": "https://console.anthropic.com/settings/keys",
    "compat": "no site do serviço escolhido (ex.: https://openrouter.ai/keys)",
}
TIMEOUT = 180


class ErroIA(Exception):
    pass


class ErroCota(ErroIA):
    """Limite de uso do provedor. diario=True: só volta depois (não adianta repetir agora)."""

    def __init__(self, msg, diario=False, espera=None, limite=None):
        super().__init__(msg)
        self.diario, self.espera, self.limite = diario, espera, limite


def _segundos(txt):
    """'9h53m44.47s' / '35s' / '1m5s' -> segundos."""
    m = re.search(r"(?:(\d+)h)?\s*(?:(\d+)m(?!s))?\s*(?:([\d.]+)s)?", txt or "")
    if not m or not any(m.groups()):
        return None
    h, mi, se = (float(x) if x else 0 for x in m.groups())
    return h * 3600 + mi * 60 + se


def _erro_cota(r, msg):
    """Interpreta o 429 de cada provedor: diário (para tudo) ou por minuto (espera e repete)."""
    texto = r.text
    espera = None
    m = re.search(r"retry in ([\dhms.]+)", texto, re.I) or re.search(r'"retryDelay":\s*"([\d.]+s)"', texto)
    if m:
        espera = _segundos(m.group(1))
    if espera is None and r.headers.get("retry-after"):
        try:
            espera = float(r.headers["retry-after"])
        except ValueError:
            pass
    lim = re.search(r"limit:\s*(\d+)", texto)
    diario = bool(re.search(r"PerDay|per day|insufficient_quota|exceeded your current quota.*billing", texto, re.I | re.S)) \
        or (espera is not None and espera > 120)
    if "insufficient_quota" in texto:  # OpenAI: sem saldo
        diario = True
    return ErroCota(f"Limite de uso atingido no provedor (429): {msg}", diario=diario, espera=espera,
                    limite=int(lim.group(1)) if lim else None)


def _url(cfg):
    return (cfg.get("url") or URLS_PADRAO[cfg["provedor"]]).rstrip("/")


def _checar(r):
    if r.status_code >= 400:
        try:
            msg = r.json()
            msg = msg.get("error", msg)
            msg = msg.get("message", msg) if isinstance(msg, dict) else msg
        except Exception:
            msg = r.text[:300]
        if r.status_code in (401, 403):
            raise ErroIA(f"Chave inválida ou sem permissão ({r.status_code}): {msg}")
        if r.status_code == 429 or (r.status_code == 503 and "overloaded" in r.text.lower()):
            raise _erro_cota(r, msg)
        raise ErroIA(f"Erro {r.status_code} do provedor: {msg}")
    return r


# ====================================================================== modelos disponíveis
def listar_modelos(cfg):
    """Lista os modelos da conta/servidor. Lança ErroIA se a conexão ou a chave falhar."""
    p, base, chave = cfg["provedor"], _url(cfg), cfg.get("chave", "")
    try:
        if p == "ollama":
            r = _checar(requests.get(f"{base}/api/tags", timeout=5))
            return sorted(m["name"] for m in r.json().get("models", []))
        if p == "gemini":
            r = _checar(requests.get(f"{base}/models", headers={"x-goog-api-key": chave},
                                     params={"pageSize": 200}, timeout=20))
            return sorted(m["name"].split("/", 1)[-1] for m in r.json().get("models", [])
                          if "generateContent" in m.get("supportedGenerationMethods", []))
        if p == "anthropic":
            r = _checar(requests.get(f"{base}/models", timeout=20, params={"limit": 100},
                                     headers={"x-api-key": chave, "anthropic-version": "2023-06-01"}))
            return [m["id"] for m in r.json().get("data", [])]
        r = _checar(requests.get(f"{base}/models", headers={"Authorization": f"Bearer {chave}"}, timeout=20))
        ids = sorted(m["id"] for m in r.json().get("data", []))
        if p == "openai":  # só modelos de texto
            ids = [i for i in ids if re.match(r"^(gpt|o\d|chatgpt)", i)
                   and not re.search(r"audio|realtime|tts|transcribe|image|search|embedding", i)]
        return ids
    except requests.RequestException as e:
        raise ErroIA(f"Sem conexão com {base}: {e.__class__.__name__}") from e


# ====================================================================== chamada
def _json_da_resposta(texto):
    texto = texto.strip()
    m = re.search(r"\{.*\}", texto, re.S)
    if not m:
        raise ValueError("sem JSON")
    return json.loads(m.group(0))


def _chamar(cfg, prompt):
    p, base, chave, modelo = cfg["provedor"], _url(cfg), cfg.get("chave", ""), cfg["modelo"]
    try:
        if p == "ollama":
            opcoes = {"temperature": 0, "num_ctx": 4096}
            if cfg.get("nucleos"):
                opcoes["num_thread"] = int(cfg["nucleos"])
            r = _checar(requests.post(f"{base}/api/chat", timeout=600, json={
                "model": modelo, "messages": [{"role": "user", "content": prompt}], "format": "json",
                "stream": False, "options": opcoes}))
            return r.json()["message"]["content"]
        if p == "gemini":
            r = _checar(requests.post(f"{base}/models/{modelo}:generateContent", timeout=TIMEOUT,
                                      headers={"x-goog-api-key": chave}, json={
                "contents": [{"role": "user", "parts": [{"text": prompt}]}],
                "generationConfig": {"responseMimeType": "application/json", "temperature": 0}}))
            dados = r.json()
            try:
                return "".join(part.get("text", "") for part in dados["candidates"][0]["content"]["parts"])
            except (KeyError, IndexError):
                raise ErroIA(f"Resposta vazia do Gemini: {str(dados)[:200]}") from None
        if p == "anthropic":
            r = _checar(requests.post(f"{base}/messages", timeout=TIMEOUT, headers={
                "x-api-key": chave, "anthropic-version": "2023-06-01", "content-type": "application/json"},
                json={"model": modelo, "max_tokens": 1500,
                      "messages": [{"role": "user", "content": prompt}]}))
            return "".join(b.get("text", "") for b in r.json().get("content", []) if b.get("type") == "text")
        # openai e compatíveis
        corpo = {"model": modelo, "messages": [{"role": "user", "content": prompt}],
                 "response_format": {"type": "json_object"}}
        cab = {"Authorization": f"Bearer {chave}"}
        r = requests.post(f"{base}/chat/completions", headers=cab, json=corpo, timeout=TIMEOUT)
        if r.status_code == 400 and "response_format" in r.text:
            corpo.pop("response_format")  # serviço não aceita o modo JSON: pede JSON só pelo texto
            r = requests.post(f"{base}/chat/completions", headers=cab, json=corpo, timeout=TIMEOUT)
        _checar(r)
        return r.json()["choices"][0]["message"]["content"] or ""
    except requests.RequestException as e:
        raise ErroIA(f"Sem conexão com {base}: {e.__class__.__name__}") from e


def _achado(ref, dados):
    rotulo = ref.texto[:80] + ("…" if len(ref.texto) > 80 else "")
    if dados is None:
        return Achado("Referências (IA)", rotulo, VERIFICAR, "a IA não devolveu resultado para esta referência",
                      origem="ia", trechos=[ref.texto])
    status = OK if str(dados.get("status", "")).upper().startswith("OK") else AJUSTAR
    probs = [str(x) for x in dados.get("problemas", []) if str(x).strip()]
    if status == OK and probs:
        status = VERIFICAR
    return Achado("Referências (IA)", rotulo, status, "; ".join(probs) or "sem problemas",
                  str(dados.get("sugestao", "")).strip() if status != OK else "",
                  f"6 ({ref.tipo})", [ref.texto], origem="ia")


def _chamar_com_espera(cfg, prompt, tentativas=4):
    """Repete quando o limite é por minuto; desiste na hora quando é diário."""
    for k in range(tentativas):
        try:
            return _chamar(cfg, prompt)
        except ErroCota as e:
            if e.diario or k == tentativas - 1:
                raise
            time.sleep(min(90, (e.espera or 0) + 1) if e.espera else 5 * (k + 1))


def _regras(cfg):
    """Regras gerais + instruções do perfil de norma escolhido (se houver)."""
    extra = (cfg.get("instrucoes") or "").strip()
    if not extra:
        return REGRAS_GERAIS
    return (REGRAS_GERAIS + "\n\nRegras ADICIONAIS da norma escolhida pela biblioteca (prevalecem sobre as "
            "anteriores quando houver conflito):\n" + extra[:4000])


def revisar_referencia(ref, cfg):
    prompt = PROMPT.format(regras=_regras(cfg), modelo=MODELOS.get(ref.tipo, MODELOS["livro"]), ref=ref.marcado)
    texto = _chamar_com_espera(cfg, prompt)
    try:
        return _achado(ref, _json_da_resposta(texto))
    except Exception:
        a = _achado(ref, None)
        a.detalhe = "resposta da IA ilegível"
        return a


PROMPT_LOTE = """Você é um bibliotecário revisor de referências segundo as Diretrizes USP (ABNT NBR 6023:2018).
Avalie CADA referência numerada abaixo, independentemente. Use apenas as regras e os modelos abaixo.
Não invente dados que não estão na referência: se faltar um elemento (editora, local, páginas, volume),
aponte a falta em vez de preencher. O negrito está marcado como **texto**.

{regras}

Modelos por tipo:
{modelos}

Referências a avaliar:
{refs}

Responda SOMENTE com JSON, com um item por referência, na mesma numeração:
{{"resultados": [{{"n": 1, "status": "OK" ou "AJUSTAR", "problemas": ["problema curto", ...],
"sugestao": "referência corrigida, com **negrito**, sem inventar dados; vazio se OK"}}, ...]}}"""


def revisar_lote(refs, cfg):
    """Revisa várias referências num único pedido (economiza a cota do provedor)."""
    if len(refs) == 1:
        return [revisar_referencia(refs[0], cfg)]
    tipos = list(dict.fromkeys(r.tipo for r in refs))
    prompt = PROMPT_LOTE.format(
        regras=_regras(cfg),
        modelos="\n\n".join(MODELOS.get(t, MODELOS["livro"]) for t in tipos),
        refs="\n".join(f"[{i}] (tipo provável: {r.tipo}) {r.marcado}" for i, r in enumerate(refs, 1)))
    texto = _chamar_com_espera(cfg, prompt)
    try:
        dados = _json_da_resposta(texto)
        itens = dados.get("resultados", dados if isinstance(dados, list) else [])
        por_n = {int(x.get("n")): x for x in itens if isinstance(x, dict) and str(x.get("n", "")).isdigit()}
        if not por_n and len(itens) == len(refs):
            por_n = {i: x for i, x in enumerate(itens, 1)}
    except Exception:
        por_n = {}
    if not por_n and len(refs) > 3:  # resposta ilegível: tenta em duas metades
        meio = len(refs) // 2
        return revisar_lote(refs[:meio], cfg) + revisar_lote(refs[meio:], cfg)
    return [_achado(r, por_n.get(i)) for i, r in enumerate(refs, 1)]


def testar(cfg):
    """Faz uma revisão curta de teste. Retorna (ok, mensagem)."""
    from .referencias import Referencia
    ref = Referencia(None, "SILVA, J. A. Normalização. 2ª edição. São Paulo: Atlas, 2020.",
                     "SILVA, J. A. **Normalização**. 2ª edição. São Paulo: Atlas, 2020.", [], ["2020"], "livro")
    try:
        a = revisar_referencia(ref, cfg)
    except ErroIA as e:
        return False, str(e)
    if a.detalhe == "resposta da IA ilegível":
        return False, "O modelo respondeu, mas fora do formato esperado. Escolha outro modelo."
    return True, f"Conexão ok. Resposta de teste: {a.status} — {a.detalhe[:120]}"


def pedidos_necessarios(n_refs, cfg):
    lote = tamanho_lote(cfg)
    return -(-n_refs // lote)


def tamanho_lote(cfg):
    if cfg["provedor"] not in NUVEM:
        return max(1, int(cfg.get("lote_local", 1)))
    return max(1, int(cfg.get("lote", 15)))


def mensagem_cota(e, cfg):
    nome = PROVEDORES[cfg["provedor"]].split(" —")[0].split(" (")[0]
    if not e.diario:
        return (f"{nome} está limitando a quantidade de pedidos por minuto. Tente de novo em instantes ou reduza "
                "'Pedidos em paralelo' na configuração.")
    quando = ""
    if e.espera:
        h, m = int(e.espera // 3600), int(e.espera % 3600 // 60)
        quando = f" Libera de novo em cerca de {h}h{m:02d}min." if h else f" Libera de novo em cerca de {m} min."
    lim = f" ({e.limite} pedidos por dia neste modelo)" if e.limite else ""
    extra = ""
    if cfg["provedor"] == "gemini" and "free_tier" in str(e):
        extra = (" Observação: esta chave está no plano gratuito da API. A assinatura do app Gemini (Plus/Pro) "
                 "não aumenta a cota da API; para mais pedidos é preciso ativar o faturamento do projeto no "
                 "Google AI Studio (Set up billing).")
    return (f"A cota do {nome} acabou{lim}.{quando} Enquanto isso: escolha outro modelo na barra lateral "
            "(no Gemini, cada modelo tem a própria cota), use outra chave ou ative o faturamento na conta, "
            f"ou use a IA local.{extra}")


def revisar_varias(refs, cfg, ao_progredir=None):
    """Revisa as referências em lotes (vários por pedido) e, na nuvem, com pedidos em paralelo.

    Retorna (achados, erros). Na cota diária esgotada, para na hora e devolve o que já foi revisado,
    com uma única mensagem explicando o que fazer.
    """
    fila = sorted(refs, key=lambda r: r.tipo)  # mesmo tipo junto: prompt menor e mais estável
    lote = tamanho_lote(cfg)
    lotes = [fila[i:i + lote] for i in range(0, len(fila), lote)]
    paralelo = max(1, int(cfg.get("paralelo", 3))) if cfg["provedor"] in NUVEM else 1
    achados, erros, feitos = [], [], 0
    parar = False
    with ThreadPoolExecutor(max_workers=paralelo) as ex:
        futuros = {ex.submit(revisar_lote, lt, cfg): lt for lt in lotes}
        for fut in as_completed(futuros):
            lt = futuros[fut]
            feitos += len(lt)
            try:
                achados.extend(fut.result())
            except CancelledError:
                continue
            except ErroCota as e:
                if not parar:
                    erros.append(mensagem_cota(e, cfg))
                if e.diario:
                    parar = True
                    for f in futuros:
                        f.cancel()
            except ErroIA as e:
                if str(e) not in erros:
                    erros.append(str(e))
                if "Chave inválida" in str(e):
                    parar = True
                    for f in futuros:
                        f.cancel()
            except Exception as e:
                erros.append(f"Erro inesperado: {e}")
            if ao_progredir:
                ao_progredir(min(feitos, len(fila)), len(fila))
    return achados, erros


# ------------------------------------------------------------------ norma a partir de um manual
PROMPT_PERFIL = """Você recebe trechos do manual de normalização de trabalhos acadêmicos de uma instituição.
Extraia SOMENTE o que o manual diz explicitamente. Se o manual não disser algo, NÃO inclua a chave.
Responda SOMENTE com JSON neste formato (todas as chaves opcionais; números com ponto decimal):
{{"nome": "nome curto da norma/instituição",
 "margens": {{"esquerda": cm, "direita": cm, "superior": cm, "inferior": cm}},
 "fonte_texto": tamanho, "entrelinhas": 1.5,
 "citacao_longa": {{"recuo": cm, "fonte": tamanho, "fonte_max": tamanho, "entrelinhas": 1.0}},
 "fonte_legenda_max": tamanho,
 "resumo": {{"min": palavras, "max": palavras}},
 "citacao_autor": "mista" (Silva, 2020) ou "maiusculas" (SILVA, 2020),
 "separador_parenteses": ";" ou "e",
 "referencias": {{"titulo": "título da lista", "destaque": "negrito" ou "italico" ou "qualquer"}},
 "elementos": {{"abstract": true/false, "ficha": true/false, "folha_aprovacao": true/false}},
 "regras": [{{"nome": "regra curta", "tipo": "proibir" ou "exigir", "padrao": "texto exato",
             "onde": "texto" ou "referencias" ou "documento", "status": "Ajustar", "como_corrigir": "..."}}],
 "instrucoes_ia": "regras de referências do manual que diferem da ABNT, em tópicos curtos"}}

Trechos do manual:
{texto}"""


def texto_do_manual(nome, dados):
    """Texto de um manual enviado (.pdf, .docx, .txt, .md)."""
    import io as _io
    nome = nome.lower()
    if nome.endswith(".pdf"):
        from pypdf import PdfReader
        r = PdfReader(_io.BytesIO(dados))
        return "\n".join((p.extract_text() or "") for p in r.pages[:120])
    if nome.endswith(".docx"):
        from docx import Document
        d = Document(_io.BytesIO(dados))
        return "\n".join(p.text for p in d.paragraphs)
    return dados.decode("utf-8", errors="replace")


def _trechos_relevantes(texto, limite=24000):
    """Fica com os parágrafos que falam de formatação, citações e referências (o manual pode ser longo)."""
    chaves = re.compile(r"margem|margens|fonte|tamanho|espa[çc]amento|entrelinha|recuo|cita[çc][ãa]o|"
                        r"refer[êe]ncia|resumo|palavras|negrito|it[áa]lico|t[íi]tulo|se[çc][ãa]o|legenda|"
                        r"ilustra|tabela|pagina[çc][ãa]o|cm\b", re.I)
    blocos = [b.strip() for b in re.split(r"\n\s*\n|(?<=\.)\n", texto) if b.strip()]
    rel = [b for b in blocos if chaves.search(b)]
    out, total = [], 0
    for b in rel:
        if total + len(b) > limite:
            break
        out.append(b)
        total += len(b)
    return "\n".join(out) or texto[:limite]


def extrair_perfil(texto_manual, cfg):
    """Pede à IA os parâmetros de norma descritos no manual. Devolve um dict parcial de perfil."""
    texto = _chamar_com_espera(cfg, PROMPT_PERFIL.format(texto=_trechos_relevantes(texto_manual)))
    dados = _json_da_resposta(texto)
    if not isinstance(dados, dict):
        raise ValueError("a IA não devolveu um perfil legível")
    return dados
