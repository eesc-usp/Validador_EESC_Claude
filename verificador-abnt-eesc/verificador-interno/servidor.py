"""Servidor da biblioteca: deixa o Verificador acessível a todos na rede local.

- Roda o app para a rede (0.0.0.0) na porta 8501, sem janela.
- Se o app cair, reinicia sozinho em 5 segundos.
- Grava o próprio PID em servidor.pid (o instalador usa para parar/reiniciar) e o log em servidor.log.

Iniciado pelo instalar.ps1 -Servidor e, a cada login no Windows, pelo atalho da pasta Inicializar.

Com --publico (publicar.ps1): roda uma 2ª cópia só para a internet, em 127.0.0.1:8502 (fora da rede
local), em modo público, e mantém o túnel do Cloudflare ligado (se for o caso), gravando o link em
compartilhar/link_publico.txt. Com Tailscale Funnel o túnel é mantido pelo próprio Tailscale.
"""
import json
import re
import threading
import os
import socket
import subprocess
import sys
import time
from pathlib import Path

PASTA = Path(__file__).resolve().parent
PUBLICO = "--publico" in sys.argv
PORTA = int(os.environ.get("VERIFICADOR_PORTA", "8502" if PUBLICO else "8501"))
PREFIXO = "publico" if PUBLICO else "servidor"
LOG = PASTA / f"{PREFIXO}.log"
PID = PASTA / f"{PREFIXO}.pid"
PARAR = PASTA / f"{PREFIXO}.parar"
CFG_PUB = PASTA / "publico.json"          # {"tunel": "tailscale"|"cloudflare"|"rapido", "token": ..., "endereco": ...}
LINK = PASTA / "compartilhar" / "link_publico.txt"


def enderecos():
    nome = socket.gethostname()
    ips = set()
    try:
        for info in socket.getaddrinfo(nome, None, socket.AF_INET):
            ip = info[4][0]
            if not ip.startswith("127."):
                ips.add(ip)
    except Exception:
        pass
    try:  # IP da interface usada para sair para a rede (não envia nada)
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.connect(("10.255.255.255", 1))
            ips.add(s.getsockname()[0])
    except Exception:
        pass
    return nome, sorted(ips)


def log(msg):
    if LOG.exists() and LOG.stat().st_size > 5_000_000:
        LOG.replace(LOG.with_suffix(".log.1"))
    with open(LOG, "a", encoding="utf-8") as f:
        f.write(f"[{time.strftime('%Y-%m-%d %H:%M:%S')}] {msg}\n")


def gravar_link(url):
    LINK.parent.mkdir(exist_ok=True)
    LINK.write_text(url + "\n", encoding="utf-8")
    (LINK.parent / "Verificador ABNT (site publico).url").write_text(f"[InternetShortcut]\nURL={url}\n",
                                                                     encoding="utf-8")
    try:
        cfg = json.loads(CFG_PUB.read_text(encoding="utf-8-sig"))
        cfg["endereco"] = url
        CFG_PUB.write_text(json.dumps(cfg, indent=2), encoding="utf-8")
    except Exception:
        pass
    log(f"Link público: {url}")


def manter_tunel():
    """Cloudflare: mantém o cloudflared rodando e reinicia se cair (o Tailscale cuida de si mesmo)."""
    try:
        cfg = json.loads(CFG_PUB.read_text(encoding="utf-8-sig"))
    except Exception:
        return
    if cfg.get("tunel") not in ("cloudflare", "rapido"):
        return
    exe = PASTA / "ferramentas" / "cloudflared.exe"
    if not exe.exists():
        log("cloudflared.exe não encontrado em ferramentas\\. Rode publicar.bat de novo.")
        return
    if cfg["tunel"] == "cloudflare":
        cmd = [str(exe), "tunnel", "--no-autoupdate", "run", "--token", cfg.get("token", "")]
        if cfg.get("endereco"):
            gravar_link(cfg["endereco"])
    else:
        cmd = [str(exe), "tunnel", "--no-autoupdate", "--url", f"http://localhost:{PORTA}"]
    flags = 0x08000000 if os.name == "nt" else 0
    while not PARAR.exists():
        proc = subprocess.Popen(cmd, cwd=PASTA, stdout=subprocess.PIPE, stderr=subprocess.STDOUT, text=True,
                                encoding="utf-8", errors="replace", creationflags=flags)
        for linha in proc.stdout:
            m = re.search(r"https://[a-z0-9-]+\.trycloudflare\.com", linha)
            if m:
                gravar_link(m.group(0))
            if PARAR.exists():
                proc.terminate()
                break
        proc.wait()
        if PARAR.exists():
            break
        log("Túnel caiu; religando em 10 s.")
        time.sleep(10)


def main():
    PID.write_text(str(os.getpid()))
    PARAR.unlink(missing_ok=True)
    nome, ips = enderecos()
    endereco = f"http://{nome}:{PORTA}"
    if PUBLICO:
        threading.Thread(target=manter_tunel, daemon=True).start()
    py = Path(sys.executable)
    if py.name.lower() == "pythonw.exe":
        py = py.with_name("python.exe")
    env = dict(os.environ, VERIFICADOR_SERVIDOR="1", VERIFICADOR_ENDERECO=endereco, PYTHONIOENCODING="utf-8")
    if PUBLICO:
        env["VERIFICADOR_PUBLICO"] = "1"
    flags = 0x08000000 if os.name == "nt" else 0  # sem janela
    log("Site público iniciado (só via túnel)." if PUBLICO else
        f"Servidor iniciado. Acesso: {endereco} " + " ".join(f"http://{ip}:{PORTA}" for ip in ips))
    falhas = 0
    while not PARAR.exists():
        inicio = time.time()
        with open(LOG, "a", encoding="utf-8") as saida:
            proc = subprocess.Popen(
                [str(py), "-m", "streamlit", "run", str(PASTA / "app.py"),
                 "--server.address", "127.0.0.1" if PUBLICO else "0.0.0.0",
                 "--server.port", str(PORTA), "--server.headless", "true"]
                + (["--server.maxUploadSize", "30", "--server.fileWatcherType", "none"] if PUBLICO else []),
                cwd=PASTA, env=env, stdout=saida, stderr=subprocess.STDOUT, creationflags=flags)
            while proc.poll() is None:
                if PARAR.exists():
                    proc.terminate()
                    break
                time.sleep(2)
        if PARAR.exists():
            break
        falhas = falhas + 1 if time.time() - inicio < 60 else 0
        espera = min(60, 5 * max(1, falhas))
        log(f"O app parou (código {proc.returncode}). Reiniciando em {espera} s.")
        time.sleep(espera)
    log("Servidor encerrado.")
    PID.unlink(missing_ok=True)
    PARAR.unlink(missing_ok=True)


if __name__ == "__main__":
    main()
