"""Perfis de norma: os padrões que a verificação e a correção seguem.

O padrão é o da EESC-USP (Diretrizes USP, parte I – ABNT, 5. ed., 2024). Um perfil pode alterar os
valores (margens, fonte, espaçamento...), ligar/desligar elementos e acrescentar regras próprias.

Uso:
    with perfil.usar(meu_perfil):
        analisar(arquivo)          # todas as regras leem perfil.atual()
"""
import contextvars
import copy
import hashlib
import json
import re
from contextlib import contextmanager
from pathlib import Path

PASTA_PERFIS = Path(__file__).resolve().parent.parent / "perfis"

EESC = {
    "id": "eesc",
    "tipo": "tcc",              # "tcc" | "pos" (dissertação/tese)
    "nome": "EESC-USP — TCC (padrão)",
    "descricao": "Diretrizes USP, parte I (ABNT), 5. ed., 2024, aplicadas ao TCC de graduação da EESC.",
    "margens": {"esquerda": 3.0, "direita": 2.0, "superior": 3.0, "inferior": 2.0},
    "fonte_texto": 12.0,
    "entrelinhas": 1.5,
    "citacao_longa": {"recuo": 4.0, "fonte": 10.0, "fonte_max": 11.0, "entrelinhas": 1.0},
    "fonte_legenda_max": 11.0,
    "resumo": {"min": 150, "max": 500},
    # chamada de autor nas citações: "mista" (Silva, 2020) | "maiusculas" (SILVA, 2020) | "qualquer"
    "citacao_autor": "mista",
    # autores dentro dos parênteses: ";" (Silva; Souza, 2020) | "e" (Silva e Souza, 2020) | "qualquer"
    "separador_parenteses": ";",
    "referencias": {
        "titulo": "REFERÊNCIAS",
        "destaque": "qualquer",        # "qualquer" (uniforme) | "negrito" | "italico"
        "ordem_alfabetica": True,
        "alinhar_esquerda": True,
    },
    "elementos": {
        "abstract": True, "keywords": True, "versao": False, "orientador": True,
        "ficha": True, "folha_aprovacao": True, "secao_nova_pagina": True,
        "natureza": True,            # nota de natureza na folha de rosto, recuada à direita
        "referencia_resumo": True,   # referência do trabalho acima do resumo/abstract
        "declaracao_ia": False,      # declaração de uso de IA generativa (template EESC de pós)
        "capa": True,                # capa com USP / EESC / São Carlos / ano
    },
    # regras próprias: {"nome", "tipo": proibir|exigir|exigir_cada, "padrao", "regex": bool,
    #                   "onde": texto|referencias|documento, "status": Ajustar|Verificar, "como_corrigir"}
    "regras": [],
    "instrucoes_ia": "",
}

# Dissertação/tese: baseado no template oficial da EESC (TEMPLATE_Dissertacao_Tese_ABNT) + Diretrizes USP
EESC_POS = None  # definido depois de _mesclar

TIPOS = {"tcc": "TCC (graduação)", "pos": "Dissertação ou Tese"}

ROTULOS = {
    "citacao_autor": {"mista": "Maiúsculas e minúsculas — (Silva, 2020)", "maiusculas": "MAIÚSCULAS — (SILVA, 2020)",
                      "qualquer": "Não verificar"},
    "separador_parenteses": {";": "Ponto e vírgula — (Silva; Souza, 2020)", "e": "'e' — (Silva e Souza, 2020)",
                             "qualquer": "Não verificar"},
    "destaque": {"qualquer": "Qualquer um, desde que uniforme", "negrito": "Negrito", "italico": "Itálico"},
    "onde": {"texto": "Texto (corpo do trabalho)", "referencias": "Cada referência", "documento": "Documento inteiro"},
    "tipo": {"proibir": "Não pode aparecer", "exigir": "Tem que aparecer (ao menos uma vez)",
             "exigir_cada": "Tem que aparecer em cada item"},
}

_ATUAL = contextvars.ContextVar("perfil_atual", default=None)


def _mesclar(base, extra):
    out = copy.deepcopy(base)
    for k, v in (extra or {}).items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _mesclar(out[k], v)
        else:
            out[k] = copy.deepcopy(v)
    return out


EESC_POS = _mesclar(EESC, {
    "id": "eesc-pos",
    "tipo": "pos",
    "nome": "EESC-USP — Dissertação/Tese (template EESC)",
    "descricao": "Template oficial de dissertação/tese da EESC (ABNT) e Diretrizes USP, parte I, 5. ed., 2024.",
    "citacao_longa": {"fonte": 11.0},
    "elementos": {"versao": True, "declaracao_ia": True},
})
PADROES = {"tcc": EESC, "pos": EESC_POS}


def base_de(p):
    return EESC_POS if (p or {}).get("tipo") == "pos" else EESC


def completo(p):
    """Perfil com todos os campos (o que faltar vem do padrão EESC do mesmo tipo de trabalho)."""
    return _mesclar(base_de(p), p or {})


def atual():
    return _ATUAL.get() or EESC


@contextmanager
def usar(p):
    tok = _ATUAL.set(completo(p) if p else EESC)
    try:
        yield atual()
    finally:
        _ATUAL.reset(tok)


def chave(p):
    """Identifica o conteúdo do perfil (para cache)."""
    return hashlib.md5(json.dumps(completo(p), sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:12]


def diferencas(p):
    """Lista legível do que o perfil muda em relação ao padrão EESC."""
    out = []

    def rec(a, b, cam):
        for k in b:
            if k in ("id", "nome", "descricao", "tipo"):
                continue
            if isinstance(b[k], dict) and isinstance(a.get(k), dict):
                rec(a[k], b[k], cam + [k])
            elif a.get(k) != b[k]:
                if k == "regras":
                    out.append(f"{len(b[k])} regra(s) personalizada(s)")
                elif k == "instrucoes_ia":
                    out.append("instruções extras para a IA")
                else:
                    out.append(f"{'.'.join(cam + [k])}: {a.get(k)} → {b[k]}")
    rec(base_de(p), completo(p), [])
    return out


# ------------------------------------------------------------------ guardar / carregar
def _slug(nome):
    import unicodedata
    s = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode().lower()
    return re.sub(r"[^a-z0-9]+", "-", s).strip("-")[:40] or "perfil"


def listar():
    """[EESC TCC, EESC dissertação/tese] + perfis salvos em perfis/*.json."""
    out = [EESC, EESC_POS]
    if PASTA_PERFIS.exists():
        for f in sorted(PASTA_PERFIS.glob("*.json")):
            try:
                p = completo(json.loads(f.read_text(encoding="utf-8-sig")))
                p["id"] = f.stem
                out.append(p)
            except Exception:
                continue
    return out


def salvar(p):
    p = completo(p)
    if p.get("id") in (None, "", "eesc", "eesc-pos"):
        p["id"] = _slug(p.get("nome", "perfil"))
    PASTA_PERFIS.mkdir(exist_ok=True)
    tmp = PASTA_PERFIS / f"{p['id']}.json.tmp"
    tmp.write_text(json.dumps(p, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(PASTA_PERFIS / f"{p['id']}.json")
    return p["id"]


def excluir(pid):
    if pid and pid not in ("eesc", "eesc-pos"):
        (PASTA_PERFIS / f"{pid}.json").unlink(missing_ok=True)


def de_json(texto):
    dados = json.loads(texto)
    if not isinstance(dados, dict):
        raise ValueError("arquivo de perfil inválido")
    p = sanear(dados)
    if dados.get("regras") is not None:  # no arquivo, as regras substituem (não somam)
        p["regras"] = sanear({"regras": dados.get("regras")}, {**EESC, "regras": []})["regras"]
    p["id"] = ""
    return p


def para_json(p):
    return json.dumps(completo(p), ensure_ascii=False, indent=2)


# ------------------------------------------------------------------ ajuda para textos das regras
def fmt(v):
    v = float(v)
    return (f"{v:.2f}".rstrip("0").rstrip(".")).replace(".", ",")


def margens_txt(p=None):
    m = (p or atual())["margens"]
    return f"{fmt(m['esquerda'])} / {fmt(m['direita'])} / {fmt(m['superior'])} / {fmt(m['inferior'])} cm"


def eh_padrao(p=None):
    return not diferencas(p or atual())


def sanear(dados, base=None):
    """Aplica sobre 'base' só os campos conhecidos e com tipo válido (para dados vindos da IA ou de arquivos)."""
    base = completo(base or EESC)
    out = copy.deepcopy(base)

    def numero(v, lo, hi):
        try:
            v = float(str(v).replace(",", "."))
        except Exception:
            return None
        return v if lo <= v <= hi else None

    if dados.get("tipo") in TIPOS and dados.get("tipo") != out.get("tipo"):
        out = _mesclar(PADROES[dados["tipo"]], {k: v for k, v in out.items() if k not in ("elementos", "tipo")})
    if isinstance(dados.get("nome"), str) and dados["nome"].strip():
        out["nome"] = dados["nome"].strip()[:80]
    for k in ("esquerda", "direita", "superior", "inferior"):
        v = numero((dados.get("margens") or {}).get(k), 0.5, 6)
        if v is not None:
            out["margens"][k] = v
    for k, lo, hi in (("fonte_texto", 8, 16), ("entrelinhas", 1, 3), ("fonte_legenda_max", 7, 14)):
        v = numero(dados.get(k), lo, hi)
        if v is not None:
            out[k] = v
    for k, lo, hi in (("recuo", 0.5, 8), ("fonte", 7, 14), ("fonte_max", 7, 14), ("entrelinhas", 1, 2)):
        v = numero((dados.get("citacao_longa") or {}).get(k), lo, hi)
        if v is not None:
            out["citacao_longa"][k] = v
    for k in ("min", "max"):
        v = numero((dados.get("resumo") or {}).get(k), 50, 2000)
        if v is not None:
            out["resumo"][k] = int(v)
    if dados.get("citacao_autor") in ROTULOS["citacao_autor"]:
        out["citacao_autor"] = dados["citacao_autor"]
    if dados.get("separador_parenteses") in ROTULOS["separador_parenteses"]:
        out["separador_parenteses"] = dados["separador_parenteses"]
    refs = dados.get("referencias") or {}
    if isinstance(refs.get("titulo"), str) and refs["titulo"].strip():
        out["referencias"]["titulo"] = refs["titulo"].strip().upper()[:40]
    if refs.get("destaque") in ROTULOS["destaque"]:
        out["referencias"]["destaque"] = refs["destaque"]
    for k in ("ordem_alfabetica", "alinhar_esquerda"):
        if isinstance(refs.get(k), bool):
            out["referencias"][k] = refs[k]
    for k, v in (dados.get("elementos") or {}).items():
        if k in out["elementos"] and isinstance(v, bool):
            out["elementos"][k] = v
    regras = []
    for r in dados.get("regras") or []:
        if isinstance(r, dict) and str(r.get("padrao", "")).strip():
            regras.append({"nome": str(r.get("nome") or r["padrao"])[:80], "padrao": str(r["padrao"])[:200],
                           "tipo": r.get("tipo") if r.get("tipo") in ROTULOS["tipo"] else "proibir",
                           "regex": bool(r.get("regex", False)),
                           "onde": r.get("onde") if r.get("onde") in ROTULOS["onde"] else "texto",
                           "status": "Verificar" if r.get("status") == "Verificar" else "Ajustar",
                           "como_corrigir": str(r.get("como_corrigir") or "")[:300]})
    if regras:
        out["regras"] = (base.get("regras") or []) + regras
    if isinstance(dados.get("instrucoes_ia"), str):
        out["instrucoes_ia"] = dados["instrucoes_ia"][:4000]
    return out
