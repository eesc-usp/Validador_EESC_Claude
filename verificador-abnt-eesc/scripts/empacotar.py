"""Gera os pacotes de distribuição em dist/.

- verificador_abnt_eesc_vX.Y.Z.zip  → pasta 'verificador_abnt_eesc' pronta para o instalar.bat
- assistente_alunos_vX.Y.Z.zip      → kit do GPT/Gem/Projeto dos alunos

Uso: python scripts/empacotar.py
"""
import fnmatch
import zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
DIST = RAIZ / "dist"
IGNORAR = ["__pycache__", "*.pyc", ".pytest_cache", ".ruff_cache", "config.json", "perfis", "publico.json", "compartilhar",
           "ferramentas", "*.log", "*.pid", "*.parar", "*.tmp", "~$*", ".venv"]


def ignorado(caminho: Path) -> bool:
    return any(fnmatch.fnmatch(parte, padrao) for parte in caminho.parts for padrao in IGNORAR)


def zipar(origem: Path, destino: Path, pasta_no_zip: str) -> int:
    n = 0
    with zipfile.ZipFile(destino, "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted(origem.rglob("*")):
            rel = f.relative_to(origem)
            if f.is_file() and not ignorado(rel):
                z.write(f, f"{pasta_no_zip}/{rel.as_posix()}")
                n += 1
    return n


def main():
    versao = (RAIZ / "verificador-interno" / "VERSION").read_text(encoding="utf-8").strip()
    DIST.mkdir(exist_ok=True)
    for origem, nome, pasta in [
        (RAIZ / "verificador-interno", f"verificador_abnt_eesc_v{versao}.zip", "verificador_abnt_eesc"),
        (RAIZ / "assistente-alunos", f"assistente_alunos_v{versao}.zip", "assistente_alunos_ABNT_EESC"),
    ]:
        n = zipar(origem, DIST / nome, pasta)
        print(f"{nome}: {n} arquivos")


if __name__ == "__main__":
    main()
