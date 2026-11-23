"""Configurações do app (config.json): IA, chaves, senhas.

As chaves de API ficam só no computador onde o app roda (servidor). Elas nunca são enviadas ao
navegador de quem usa o app: a tela mostra apenas os 4 últimos caracteres.
"""
import hashlib
import json
import os
import secrets
from pathlib import Path

PASTA = Path(__file__).resolve().parent.parent
ARQ = PASTA / "config.json"


def carregar():
    try:
        cfg = json.loads(ARQ.read_text(encoding="utf-8-sig"))
    except Exception:
        cfg = {}
    ia = cfg.setdefault("ia", {})
    ia.setdefault("provedor", "ollama")
    ia.setdefault("modelos", {})
    ia.setdefault("chaves", {})
    ia.setdefault("urls", {})
    ia.setdefault("paralelo", 3)
    ia.setdefault("lote", 15)
    # compatibilidade com o config.json das versões anteriores (gravado pelo instalador)
    if cfg.get("modelo") and not ia["modelos"].get("ollama"):
        ia["modelos"]["ollama"] = cfg["modelo"]
    if cfg.get("ollama_url") and not ia["urls"].get("ollama"):
        ia["urls"]["ollama"] = cfg["ollama_url"]
    cfg.setdefault("nucleos", max(1, (os.cpu_count() or 4) - 2))
    return cfg


def salvar(cfg):
    tmp = ARQ.with_suffix(".tmp")
    tmp.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, ARQ)


def hash_senha(senha):
    sal = secrets.token_hex(8)
    h = hashlib.pbkdf2_hmac("sha256", senha.encode(), sal.encode(), 200_000).hex()
    return f"{sal}${h}"


def conferir(senha, guardado):
    if not guardado or "$" not in guardado:
        return False
    sal, h = guardado.split("$", 1)
    return secrets.compare_digest(hashlib.pbkdf2_hmac("sha256", senha.encode(), sal.encode(), 200_000).hex(), h)


def mascarar(chave):
    return f"••••{chave[-4:]}" if chave else "não informada"


def cfg_ia(cfg, provedor=None):
    """Monta a configuração usada por ia_provedores para o provedor escolhido."""
    ia = cfg["ia"]
    p = provedor or ia["provedor"]
    return {"provedor": p, "chave": ia["chaves"].get(p, ""), "url": ia["urls"].get(p, ""),
            "modelo": ia["modelos"].get(p, ""), "nucleos": cfg.get("nucleos"), "paralelo": ia.get("paralelo", 3),
            "lote": ia.get("lote", 15)}


def provedores_prontos(cfg):
    """Provedores com o necessário configurado (chave e modelo; o Ollama só precisa do modelo)."""
    ia = cfg["ia"]
    return [p for p in ("ollama", "gemini", "openai", "anthropic", "compat")
            if ia["modelos"].get(p) and (p == "ollama" or ia["chaves"].get(p))]
