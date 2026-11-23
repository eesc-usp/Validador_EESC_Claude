"""Abre o Verificador ABNT como aplicativo de desktop, numa janela própria.

- Sobe o servidor do app em segundo plano (sem janela de console), numa porta livre.
- Mostra uma janela nativa (WebView2 / Edge do Windows) com tela de carregamento.
- Ao fechar a janela, encerra o servidor.
- Se a janela nativa não estiver disponível, abre no navegador padrão.

Rodar:  pythonw desktop.py    (o atalho da área de trabalho faz isso)
"""
import os
import socket
import subprocess
import sys
import threading
import time
import urllib.request
import webbrowser
from pathlib import Path

PASTA = Path(__file__).resolve().parent
TITULO = "Verificador ABNT – Biblioteca EESC"
LOG = PASTA / "app.log"

CARREGANDO = """<!DOCTYPE html><html><head><meta charset="utf-8"><style>
body{margin:0;height:100vh;display:flex;align-items:center;justify-content:center;
font-family:"Segoe UI",sans-serif;background:linear-gradient(135deg,#1F4E79,#2E6DA4);color:#fff}
.c{text-align:center}.l{font-size:54px;margin-bottom:10px}h1{margin:0 0 6px;font-size:26px}
p{opacity:.85;margin:0 0 22px}.b{width:220px;height:5px;background:rgba(255,255,255,.25);border-radius:9px;
overflow:hidden;margin:auto}.b i{display:block;width:40%;height:100%;background:#fff;border-radius:9px;
animation:a 1.2s infinite ease-in-out}@keyframes a{0%{margin-left:-40%}100%{margin-left:100%}}
</style></head><body><div class="c"><div class="l">📘</div><h1>Verificador ABNT</h1>
<p>Biblioteca da EESC-USP · abrindo...</p><div class="b"><i></i></div></div></body></html>"""

ERRO = """<!DOCTYPE html><html><head><meta charset="utf-8"><style>body{{font-family:"Segoe UI",sans-serif;
padding:40px;color:#1F2937}}code{{background:#F1F5F9;padding:2px 6px;border-radius:4px}}</style></head><body>
<h2>Não foi possível iniciar o Verificador</h2><p>{msg}</p>
<p>Detalhes em <code>{log}</code>. Rode <code>instalar.bat</code> para reparar a instalação.</p></body></html>"""


def porta_livre(preferida=8501):
    for porta in [preferida, *range(8502, 8600)]:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(("127.0.0.1", porta))
                return porta
            except OSError:
                continue
    raise RuntimeError("Nenhuma porta livre entre 8501 e 8599.")


def no_ar(url):
    try:
        urllib.request.urlopen(url + "/_stcore/health", timeout=2)
        return True
    except Exception:
        return False


def iniciar_servidor(porta):
    py = Path(sys.executable)
    if py.name.lower() == "pythonw.exe":  # o servidor precisa de stdout: usa python.exe sem janela
        py = py.with_name("python.exe")
    env = dict(os.environ, VERIFICADOR_DESKTOP="1", PYTHONIOENCODING="utf-8")
    flags = 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW
    log = open(LOG, "w", encoding="utf-8")
    return subprocess.Popen(
        [str(py), "-m", "streamlit", "run", str(PASTA / "app.py"), "--server.port", str(porta),
         "--server.headless", "true", "--server.address", "localhost"],
        cwd=PASTA, env=env, stdout=log, stderr=subprocess.STDOUT, creationflags=flags)


def main():
    porta = porta_livre()
    url = f"http://localhost:{porta}"
    proc = iniciar_servidor(porta)
    try:
        try:
            import webview
        except ImportError:
            webview = None

        if webview is None:
            for _ in range(240):
                if no_ar(url):
                    break
                time.sleep(0.5)
            webbrowser.open(url)
            proc.wait()
            return

        try:
            webview.settings["ALLOW_DOWNLOADS"] = True
        except Exception:
            pass
        janela = webview.create_window(TITULO, html=CARREGANDO, width=1480, height=940, min_size=(1100, 700),
                                       background_color="#1F4E79", text_select=True)

        def carregar():
            for _ in range(240):  # até 2 minutos (primeira abertura pode demorar)
                if proc.poll() is not None:
                    break
                if no_ar(url):
                    janela.load_url(url)
                    return
                time.sleep(0.5)
            janela.load_html(ERRO.format(msg="O servidor do app não respondeu.", log=LOG))

        threading.Thread(target=carregar, daemon=True).start()
        try:
            webview.start(gui="edgechromium" if os.name == "nt" else None)
        except Exception:
            webbrowser.open(url)  # sem WebView2: usa o navegador
            proc.wait()
    finally:
        if proc.poll() is None:
            proc.terminate()
            try:
                proc.wait(timeout=8)
            except Exception:
                proc.kill()


if __name__ == "__main__":
    main()
