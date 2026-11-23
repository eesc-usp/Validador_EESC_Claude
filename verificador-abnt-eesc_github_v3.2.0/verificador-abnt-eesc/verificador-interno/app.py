"""Verificador ABNT – Biblioteca EESC-USP (uso interno).

Rodar:  streamlit run app.py   (ou atalho "Verificador ABNT EESC" / iniciar.bat)
"""
import hashlib
import html
import io
import os
import re
import secrets
import time
import zipfile
from pathlib import Path

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from verificador import (
    AJUSTAR,
    OK,
    RESUMO_ELEMENTOS,
    VERIFICAR,
    analisar,
    anotacao,
    configuracao,
    ia_provedores,
    relatorio,
)
from verificador import perfil as PF
from verificador import trabalho as TRAB
from verificador.cache import leitura
from verificador.documento import Documento
from verificador.texto import tem_alteracoes as _tem_alteracoes
from verificador.visualizacao import renderizar, serializar

PASTA = Path(__file__).parent
VERSAO = (PASTA / "VERSION").read_text(encoding="utf-8-sig").strip() if (PASTA / "VERSION").exists() else "dev"
CONFIG = configuracao.carregar()

DESKTOP = os.environ.get("VERIFICADOR_DESKTOP") == "1"
SERVIDOR = os.environ.get("VERIFICADOR_SERVIDOR") == "1"
PUBLICO = os.environ.get("VERIFICADOR_PUBLICO") == "1"   # site aberto na internet (publicar.bat)
ENDERECO_REDE = os.environ.get("VERIFICADOR_ENDERECO", "")
PASTA_SAIDA = Path(os.environ.get("USERPROFILE", str(Path.home()))) / "Documents" / "Verificador ABNT"
area_trabalho = components.declare_component("area_trabalho", path=str(PASTA / "componentes" / "area_trabalho"))
MIME_DOCX = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
ROTULO = {AJUSTAR: "Ajustar", VERIFICAR: "Conferir", OK: "OK"}

st.set_page_config(page_title="Verificador ABNT – Biblioteca EESC", page_icon="📘", layout="wide",
                   initial_sidebar_state="collapsed")

# ------------------------------------------------------------------ tema (claro/escuro)
try:
    ESCURO = (st.context.theme.type or "light") == "dark"
except Exception:
    ESCURO = False
CORES = ({"borda": "#2B3747", "cartao": "#17212C", "suave": "#9AA6B5", "forte": "#E5E9F0", "tag": "#243142",
          "aviso": "#17212C", "aviso_tx": "#C9D3DF"} if ESCURO else
         {"borda": "#E3E7ED", "cartao": "#FFFFFF", "suave": "#6B7280", "forte": "#1F2937", "tag": "#EEF2F7",
          "aviso": "#F3F6FA", "aviso_tx": "#374151"})
st.markdown("<style>:root {" + "".join(f"--{k.replace('_', '-')}:{v};" for k, v in CORES.items()) + "}"
            + (".topo {background: linear-gradient(135deg, #16324F 0%, #1F4E79 100%) !important;}"
               ".tag.verde {background:#173B26 !important; color:#8FD8A8 !important;}"
               ".tag.amarela {background:#3D3110 !important; color:#F4D47C !important;}" if ESCURO else "")
            + "</style>", unsafe_allow_html=True)

st.markdown("""
<style>
.block-container {padding-top: 1.2rem; padding-bottom: 3rem; max-width: none; padding-left: 2rem; padding-right: 2rem;}
header[data-testid="stHeader"] {background: transparent;}
.topo {display:flex; align-items:center; gap:14px; padding:14px 20px; border-radius:14px; margin-bottom:14px;
       background: linear-gradient(135deg, #1F4E79 0%, #2E6DA4 100%); color:#fff;}
.topo .logo {font-size:30px; background:rgba(255,255,255,.15); width:52px; height:52px; border-radius:12px;
             display:flex; align-items:center; justify-content:center;}
.topo h1 {font-size:1.45rem; margin:0; color:#fff; padding:0; font-weight:700;}
.topo p {margin:2px 0 0; opacity:.85; font-size:.9rem;}
.topo .versao {margin-left:auto; font-size:.78rem; background:rgba(255,255,255,.16); padding:4px 10px; border-radius:999px;}
.kpi {border:1px solid var(--borda); border-radius:12px; padding:12px 16px; background:var(--cartao);}
.kpi .rot {font-size:.82rem; color:var(--suave); display:flex; align-items:center; gap:6px;}
.kpi .val {font-size:1.9rem; font-weight:700; line-height:1.2; color:var(--forte);}
.kpi .sub {font-size:.78rem; color:var(--suave);}
.dot {width:9px; height:9px; border-radius:50%; display:inline-block;}
.arquivo {display:flex; align-items:center; gap:10px; flex-wrap:wrap; margin:2px 0 10px;}
.arquivo .nome {font-weight:700; font-size:1.1rem;}
.tag {font-size:.75rem; padding:3px 9px; border-radius:999px; background:var(--tag); color:var(--forte);}
.tag.verde {background:#DDF3E4; color:#1E6B3A;} .tag.amarela {background:#FFF1C2; color:#7A5600;}
.boas-vindas {display:grid; grid-template-columns:repeat(auto-fit,minmax(230px,1fr)); gap:12px; margin-top:6px;}
.bv {border:1px solid var(--borda); border-radius:12px; padding:14px 16px; background:var(--cartao);}
.bv .ic {font-size:1.4rem;} .bv b {display:block; margin:6px 0 2px;} .bv span {color:var(--suave); font-size:.88rem;}
div[data-testid="stTabs"] button p {font-size:.95rem;}
.pill {display:inline-block; padding:2px 9px; border-radius:999px; font-size:.75rem; font-weight:600;}
.pill-Ajustar {background:#FDE2E4; color:#9B1C2C;} .pill-Verificar {background:#FFF1C2; color:#7A5600;}
.pill-OK {background:#DDF3E4; color:#1E6B3A;}
section[data-testid="stSidebar"][aria-expanded="true"] {min-width: 330px;}
</style>
""", unsafe_allow_html=True)

SELO_DESKTOP = ("<span class='versao'>📂 salva em Documentos/Verificador ABNT</span>" if DESKTOP else
                "" if PUBLICO else
                "<span class='versao'>🌐 servidor da biblioteca</span>" if SERVIDOR else "")
ONDE = ("o arquivo é verificado no computador da Biblioteca e não fica guardado" if PUBLICO
        else "o arquivo é processado neste computador")
c_topo, c_tema = st.columns([14, 1], vertical_alignment="center")
c_topo.markdown(f"""<div class="topo"><div class="logo">📘</div><div><h1>Verificador ABNT</h1>
<p>Biblioteca da EESC-USP · Diretrizes USP (ABNT) · {ONDE}</p></div>
{SELO_DESKTOP}
<span class="versao">versão {VERSAO}</span></div>""", unsafe_allow_html=True)
if c_tema.button("☀️" if ESCURO else "🌙", key="tema", help="Modo claro" if ESCURO else "Modo escuro"):
    st.session_state.trocar_tema = "Light" if ESCURO else "Dark"
if st.session_state.pop("trocar_tema", None) in ("Light", "Dark"):
    novo_tema = "Light" if ESCURO else "Dark"
    # o Streamlit guarda a escolha de tema no navegador; o trabalho aberto é recuperado após recarregar (cofre)
    components.html(f"""<script>
      const w = window.parent;
      w.localStorage.setItem("stActiveTheme-" + w.location.pathname + "-v2", JSON.stringify("{novo_tema}"));
      setTimeout(() => w.location.reload(), 150);
    </script>""", height=0)


# ------------------------------------------------------------------ cache
@st.cache_resource(show_spinner=False, max_entries=8)
def _analisar_cache(dados, chave_perfil, _perfil):
    achados, refs = analisar(io.BytesIO(dados), _perfil)
    return achados, refs


def analisar_bytes(dados):
    return _analisar_cache(dados, PF.chave(PERFIL), PERFIL)


@st.cache_resource(show_spinner=False, max_entries=8)
def visual(chave_cache, dados, _achados):
    with leitura():  # desenhar e localizar os achados no mesmo documento: as leituras são reaproveitadas
        html_doc, d = renderizar(io.BytesIO(dados))
        return html_doc, serializar(d, _achados)


@st.cache_resource(show_spinner=False, max_entries=8)
def apontamentos(chave_cache, dados, _achados):
    return anotacao.gerar(dados, _achados)


@st.cache_resource(show_spinner=False, max_entries=8)
def relatorio_docx(chave_cache, _todos, nome, incluir_ok, observacao):
    """Relatório .docx (guardado pela chave: documento, norma, data e opções)."""
    return relatorio.gerar(_todos, nome, incluir_ok, observacao)


@st.cache_resource(show_spinner=False, max_entries=16)
def tem_alteracoes(dados):
    """Se o arquivo do aluno tem alterações controladas pendentes (guardado: a tela é redesenhada a cada clique)."""
    return _tem_alteracoes(dados)


def _celula_vazia(v):
    """Célula sem valor no editor de tabela (None, NaN ou NA do pandas)."""
    if v is None:
        return True
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def tabela(dados, coluna=None):
    df = pd.DataFrame(dados if coluna is None else {coluna: dados})
    try:
        st.dataframe(df, hide_index=True, width="stretch")
    except TypeError:
        st.dataframe(df, hide_index=True, use_container_width=True)


def botao_arquivo(col, rotulo, dados, nome, mime=None, **kw):
    """Navegador: download. Aplicativo de desktop: salva em Documentos\\Verificador ABNT e mostra no Explorer."""
    if not DESKTOP:
        col.download_button(rotulo, dados, file_name=nome, mime=mime or MIME_DOCX, **kw)
        return
    if col.button(rotulo.replace("⬇", "💾").replace("Baixar", "Salvar"), key=f"salvar_{nome}", **kw):
        try:
            PASTA_SAIDA.mkdir(parents=True, exist_ok=True)
            destino = PASTA_SAIDA / nome
            destino.write_bytes(dados)
            st.toast(f"Salvo em {destino}", icon="💾")
            if os.name == "nt":
                import subprocess
                subprocess.Popen(["explorer", "/select,", str(destino)])
        except Exception as e:
            st.toast(f"Não foi possível salvar: {e}", icon="⚠️")


def pill(status):
    return f"<span class='pill pill-{status}'>{ROTULO.get(status, status)}</span>"


ss = st.session_state

# ------------------------------------------------------------------ acesso (modo servidor)
def _ip_cliente():
    try:
        return st.context.ip_address
    except Exception:
        return None


def _ip_visitante():
    """IP real de quem acessa pelo site público (o túnel informa nos cabeçalhos)."""
    try:
        h = st.context.headers
        for nome in ("Cf-Connecting-Ip", "X-Forwarded-For", "X-Real-Ip"):
            v = h.get(nome)
            if v:
                return v.split(",")[0].strip()
    except Exception:
        pass
    return _ip_cliente() or "?"


@st.cache_resource
def _contadores():
    return {}


def limite_ok(tipo, quantidade, maximo, janela_s, por_visitante=True):
    """Limite de uso (site público), por visitante ou no total. Devolve (ok, quanto_resta)."""
    if not PUBLICO or not maximo:
        return True, None
    agora = time.time()
    chave_lim = (tipo, _ip_visitante() if por_visitante else "*")
    lista = [t for t in _contadores().get(chave_lim, []) if agora - t < janela_s]
    _contadores()[chave_lim] = lista
    if len(lista) + quantidade > maximo:
        return False, max(0, maximo - len(lista))
    lista.extend([agora] * quantidade)
    return True, maximo - len(lista)


def ia_permitida(n):
    """Site público: limita quantas referências cada visitante (e o site todo) manda para a IA por dia."""
    lim = CONFIG.get("publico", {})
    ok1, resta = limite_ok("ia", n, lim.get("ia_refs_por_visitante_dia", 80), 86400)
    if not ok1:
        st.warning(f"Limite diário de revisão por IA atingido para você (restam {resta} referência(s) hoje). "
                   "Diminua a quantidade ou tente amanhã.")
        return False
    ok2, _ = limite_ok("ia", n, lim.get("ia_refs_total_dia", 600), 86400, por_visitante=False)
    if not ok2:
        st.warning("A cota diária de IA do site acabou. As verificações por regra continuam funcionando.")
        return False
    return True


# ------------------------------------------------------------------ cofre: o trabalho aberto sobrevive a
# recarregar a página (F5, troca de tema). Fica só na memória; apaga sozinho.
VALIDADE_COFRE = 3600 if PUBLICO else 8 * 3600
CAMPOS_COFRE = ("nome", "h_orig", "original", "ia", "perfil_id", "perfis_sessao", "tipo_detectado")


@st.cache_resource
def _cofre():
    return {}


def cofre_guardar():
    t = st.query_params.get("t") or ss.get("token")
    if not t:
        t = secrets.token_urlsafe(9)
    ss.token = t
    if st.query_params.get("t") != t:
        st.query_params["t"] = t
    agora = time.time()
    for k in [k for k, v in _cofre().items() if agora - v["_ts"] > VALIDADE_COFRE]:
        _cofre().pop(k, None)
    while len(_cofre()) >= 60:
        _cofre().pop(min(_cofre(), key=lambda k: _cofre()[k]["_ts"]))
    _cofre()[t] = {**{c: ss.get(c) for c in CAMPOS_COFRE}, "_ts": agora}


def cofre_recuperar():
    t = st.query_params.get("t")
    item = _cofre().get(t) if t else None
    if not item or time.time() - item["_ts"] > VALIDADE_COFRE:
        return False
    ss.update({c: item[c] for c in CAMPOS_COFRE if item.get(c) is not None})
    ss.setdefault("msg", None)
    ss.token = t
    return True


def cofre_fechar():
    _cofre().pop(ss.get("token") or st.query_params.get("t"), None)
    for c in CAMPOS_COFRE:
        ss.pop(c, None)
    ss.pop("token", None)
    if "t" in st.query_params:
        del st.query_params["t"]


def _e_local():
    ip = _ip_cliente()
    return ip is None or ip in ("127.0.0.1", "::1", "localhost")


if CONFIG.get("senha_acesso") and not PUBLICO and not ss.get("acesso_ok"):
    with st.form("acesso"):
        st.markdown("#### 🔒 Acesso ao Verificador")
        senha = st.text_input("Senha de acesso da Biblioteca", type="password")
        if st.form_submit_button("Entrar", type="primary"):
            if configuracao.conferir(senha, CONFIG["senha_acesso"]):
                ss.acesso_ok = True
                st.rerun()
            st.error("Senha incorreta.")
    st.stop()

# administrador: pode ver/alterar as configurações de IA e senhas
if CONFIG.get("senha_admin"):
    admin = bool(ss.get("admin_ok"))
elif PUBLICO:
    admin = False  # no site público, todo acesso chega pelo túnel como "local": só com senha de administrador
else:
    admin = (not SERVIDOR) or _e_local()  # servidor sem senha: só o próprio PC do servidor configura


def form_provedor(p):
    """Formulário de configuração de um provedor. A chave digitada não volta para o navegador."""
    ia = CONFIG["ia"]
    tmp = ss.setdefault("chave_tmp", {})
    listados = ss.setdefault("modelos_listados", {})
    salvo = ia["chaves"].get(p, "")
    with st.form(f"form_{p}", clear_on_submit=True, border=False):
        chave_nova = ""
        if p != "ollama":
            chave_nova = st.text_input("Chave de API", type="password",
                                       placeholder=f"atual: {configuracao.mascarar(tmp.get(p) or salvo)}",
                                       help="Cole uma chave nova para trocar. A chave fica só neste computador.")
            st.caption(f"Onde conseguir: {ia_provedores.ONDE_PEGAR_CHAVE[p]}")
        url = ""
        if p in ("ollama", "compat"):
            url = st.text_input("Endereço", value=ia["urls"].get(p) or ia_provedores.URLS_PADRAO[p],
                                help="OpenRouter: https://openrouter.ai/api/v1 · Groq: https://api.groq.com/openai/v1 · "
                                     "DeepSeek: https://api.deepseek.com/v1 · LM Studio: http://localhost:1234/v1")
        atual = ia["modelos"].get(p, "")
        opcoes = listados.get(p) or []
        if opcoes:
            if atual and atual not in opcoes:
                opcoes = [atual] + opcoes
            modelo = st.selectbox("Modelo", opcoes, index=opcoes.index(atual) if atual in opcoes else 0)
        else:
            modelo = st.text_input("Modelo", value=atual, placeholder="clique em 'Carregar modelos'")
        lote = ia.get("lote", 15)
        if p in ia_provedores.NUVEM:
            lote = st.slider("Referências por pedido", 1, 30, int(ia.get("lote", 15)),
                             help="Várias referências num mesmo pedido economizam a cota (planos gratuitos limitam "
                                  "pedidos por dia). Se a IA se confundir, reduza.")
            paralelo = st.slider("Pedidos em paralelo", 1, 6, int(ia.get("paralelo", 3)),
                                 help="Mais rápido; reduza se o provedor limitar pedidos por minuto.")
            nucleos = CONFIG.get("nucleos")
        else:
            total = os.cpu_count() or 4
            nucleos = st.slider("Núcleos da CPU para a IA", 1, total, min(total, int(CONFIG.get("nucleos") or total)),
                                help="Deixe núcleos livres para o BEESC BOT e o uso normal do PC.")
            paralelo = ia.get("paralelo", 4)
        b1, b2 = st.columns(2)
        carregar = b1.form_submit_button("🔄 Modelos", width="stretch", help="Busca os modelos da sua conta")
        testar = b2.form_submit_button("🧪 Testar", width="stretch", help="Revisa uma referência de exemplo")
        gravar = st.form_submit_button("💾 Salvar como padrão", type="primary", width="stretch")

    if not (carregar or testar or gravar):
        return
    if chave_nova.strip():
        tmp[p] = chave_nova.strip()
    cfg_p = {"provedor": p, "chave": tmp.get(p) or salvo, "url": url or ia["urls"].get(p, ""),
             "modelo": modelo, "nucleos": nucleos, "paralelo": paralelo, "lote": lote}
    if p != "ollama" and not cfg_p["chave"]:
        st.error("Informe a chave de API.")
        return
    if carregar:
        try:
            listados[p] = ia_provedores.listar_modelos(cfg_p)
            ss.aviso_cfg = ("ok", f"{len(listados[p])} modelo(s) encontrados.")
        except ia_provedores.ErroIA as e:
            ss.aviso_cfg = ("erro", str(e))
        st.rerun()
    if testar:
        if not modelo:
            st.error("Escolha um modelo antes de testar.")
            return
        with st.spinner("Testando..."):
            ok, msg = ia_provedores.testar(cfg_p)
        ss.aviso_cfg = ("ok" if ok else "erro", msg)
        st.rerun()
    if gravar:
        if not modelo:
            st.error("Escolha um modelo antes de salvar.")
            return
        if tmp.get(p):
            ia["chaves"][p] = tmp.pop(p)
        if url:
            ia["urls"][p] = url
        ia["modelos"][p] = modelo
        ia["paralelo"] = int(paralelo)
        ia["lote"] = int(lote)
        CONFIG["nucleos"] = int(nucleos or CONFIG.get("nucleos") or 1)
        ia["provedor"] = p
        configuracao.salvar(CONFIG)
        ss.provedor = p
        ss.aviso_cfg = ("ok", f"{ia_provedores.PROVEDORES[p].split(' —')[0]} salvo como padrão.")
        st.rerun()


with st.sidebar:
    st.subheader("⚙️ Configurações")
    st.subheader("🤖 IA para as referências")
    prontos = configuracao.provedores_prontos(CONFIG)
    visiveis = list(ia_provedores.PROVEDORES) if admin else (prontos or ["ollama"])
    padrao = ss.get("provedor") or CONFIG["ia"]["provedor"]
    provedor = st.selectbox("Provedor", visiveis, index=visiveis.index(padrao) if padrao in visiveis else 0,
                            format_func=lambda k: ("✓ " if k in prontos else "○ ") +
                            ia_provedores.PROVEDORES[k].split(" —")[0])
    ss.provedor = provedor
    cfg_ativo = configuracao.cfg_ia(CONFIG, provedor)
    if provedor in prontos:
        extra = f" · chave {configuracao.mascarar(cfg_ativo['chave'])}" if provedor != "ollama" else ""
        st.caption(f"Modelo: **{cfg_ativo['modelo']}**{extra}")
    else:
        st.caption("Ainda não configurado." + ("" if admin else " Peça ao administrador."))
    if provedor in ia_provedores.NUVEM:
        st.caption("☁️ Só o texto das referências vai para o provedor; o trabalho do aluno não sai do computador.")

    st.divider()
    st.subheader("📐 Norma usada")
    perfis = {p["id"]: p for p in PF.listar()}
    perfis.update(ss.get("perfis_sessao") or {})
    if ss.get("perfil_id") not in perfis:
        ss.perfil_id = "eesc"
    ids = list(perfis)
    escolhido = st.selectbox("Norma", ids, index=ids.index(ss.perfil_id), label_visibility="collapsed",
                             format_func=lambda k: perfis[k]["nome"] + (" (só nesta sessão)" if k.startswith("sessao")
                                                                        else ""))
    if escolhido != ss.perfil_id:
        ss.perfil_id = escolhido
        st.rerun()
    PERFIL = perfis[ss.perfil_id]
    ss.perfil_tipo = PERFIL.get("tipo", "tcc")
    difs = PF.diferencas(PERFIL)
    st.caption((f"Padrão da EESC para {PF.TIPOS[PERFIL.get('tipo', 'tcc')]}." if not difs else
                f"{len(difs)} diferença(s) em relação ao padrão EESC.") + " Base: " + PF.TIPOS[PERFIL.get("tipo", "tcc")])
    if st.button("✏️ Criar ou editar norma", width="stretch"):
        ss.abrir_norma, ss.rascunho_norma = True, None
        ss.ver_norma = ss.get("ver_norma", 0) + 1

    aviso = ss.pop("aviso_cfg", None)
    if aviso:
        (st.success if aviso[0] == "ok" else st.error)(aviso[1])

    if admin:
        with st.expander(f"Configurar {ia_provedores.PROVEDORES[provedor].split(' (')[0].split(' —')[0]}",
                         expanded=provedor not in prontos):
            form_provedor(provedor)
            if provedor != "ollama" and CONFIG["ia"]["chaves"].get(provedor):
                if st.button("Remover chave salva", key=f"rm_{provedor}"):
                    CONFIG["ia"]["chaves"].pop(provedor, None)
                    configuracao.salvar(CONFIG)
                    st.rerun()
        with st.expander("🔑 Senhas"):
            with st.form("senhas", clear_on_submit=True, border=False):
                nova_admin = st.text_input("Nova senha de administrador", type="password",
                                           help="Protege estas configurações e as chaves.")
                nova_acesso = st.text_input("Senha de acesso (opcional)", type="password",
                                            help="Se definida, todos precisam dela para usar o app na rede.")
                tirar_acesso = st.checkbox("Remover senha de acesso")
                if st.form_submit_button("Salvar senhas"):
                    if nova_admin:
                        CONFIG["senha_admin"] = configuracao.hash_senha(nova_admin)
                        ss.admin_ok = True
                    if nova_acesso:
                        CONFIG["senha_acesso"] = configuracao.hash_senha(nova_acesso)
                        ss.acesso_ok = True
                    if tirar_acesso:
                        CONFIG.pop("senha_acesso", None)
                    configuracao.salvar(CONFIG)
                    ss.aviso_cfg = ("ok", "Senhas atualizadas.")
                    st.rerun()
        if CONFIG.get("senha_admin") and ss.get("admin_ok"):
            if st.button("🔒 Sair do modo administrador"):
                ss.admin_ok = False
                st.rerun()
    else:
        with st.expander("🔒 Configurações do administrador"):
            if not CONFIG.get("senha_admin"):
                st.caption("Ainda não há senha de administrador. Configure abrindo o app no próprio "
                           "computador do servidor (http://localhost:8501).")
            else:
                with st.form("admin_login", clear_on_submit=True, border=False):
                    senha = st.text_input("Senha de administrador", type="password")
                    if st.form_submit_button("Entrar"):
                        if configuracao.conferir(senha, CONFIG["senha_admin"]):
                            ss.admin_ok = True
                            st.rerun()
                        st.error("Senha incorreta.")
    st.divider()
    link_pub = PASTA / "compartilhar" / "link_publico.txt"
    if admin and not PUBLICO and link_pub.exists():
        st.caption(f"🌍 Site público para os alunos: {link_pub.read_text(encoding='utf-8-sig').strip()}")
    if SERVIDOR and ENDERECO_REDE and not PUBLICO:
        st.caption(f"🌐 Endereço na rede: {ENDERECO_REDE}")
    st.caption(f"Versão {VERSAO} · Base: Diretrizes USP parte I (ABNT), 5. ed., 2024.")


# ------------------------------------------------------------------ editor de normas
cfg_ativo = {**cfg_ativo, "instrucoes": PERFIL.get("instrucoes_ia", "")}


@st.dialog("📐 Norma de verificação", width="large")
def editor_norma():
    base = ss.get("rascunho_norma") or PF.completo(PERFIL)
    v = ss.get("ver_norma", 0)  # muda quando um rascunho novo é carregado (renova os campos)
    st.caption("Ajuste os padrões que a verificação segue. A norma da EESC continua "
               "disponível como padrão; aqui você cria uma variação ou uma norma nova.")
    aba_imp, aba_val, aba_reg, aba_ia = st.tabs(["📥 Começar de…", "📏 Valores", "➕ Regras próprias",
                                                  "🤖 Instruções para a IA"])
    with aba_imp:
        c1, c2 = st.columns(2)
        if c1.button("Começar do padrão EESC — TCC", width="stretch"):
            ss.rascunho_norma = PF.completo({"nome": "Minha norma (baseada na EESC — TCC)"})
            ss.ver_norma = v + 1
            st.rerun(scope="fragment")
        if c1.button("Começar do padrão EESC — Dissertação/Tese", width="stretch"):
            ss.rascunho_norma = PF.completo({"tipo": "pos", "nome": "Minha norma (baseada na EESC — pós)"})
            ss.ver_norma = v + 1
            st.rerun(scope="fragment")
        arq_p = c2.file_uploader("Importar norma (.json)", type=["json"], key="imp_norma")
        if arq_p is not None and c2.button("Carregar arquivo", width="stretch"):
            try:
                ss.rascunho_norma = PF.de_json(arq_p.getvalue().decode("utf-8-sig"))
                ss.ver_norma = v + 1
                st.rerun(scope="fragment")
            except Exception as e:
                st.error(f"Arquivo inválido: {e}")
        st.markdown("**Ler o manual da sua instituição com IA**")
        st.caption("Envie o manual de normalização (PDF, Word ou texto). A IA configurada lê os trechos sobre "
                   "formatação, citações e referências e preenche os valores para você conferir. "
                   "Só o texto do manual é enviado ao provedor.")
        man = st.file_uploader("Manual", type=["pdf", "docx", "txt", "md"], key="manual_norma",
                               label_visibility="collapsed")
        pronto_ia = provedor in configuracao.provedores_prontos(CONFIG)
        if not pronto_ia:
            st.caption("⚠️ Configure um provedor de IA na barra lateral para usar esta opção.")
        if man is not None and st.button("Ler manual com IA", type="primary", disabled=not pronto_ia) and \
                ia_permitida(10):
            with st.spinner("Lendo o manual..."):
                try:
                    texto = ia_provedores.texto_do_manual(man.name, man.getvalue())
                    if len(texto.strip()) < 200:
                        raise ValueError("não foi possível extrair texto do arquivo (PDF escaneado?)")
                    dados = ia_provedores.extrair_perfil(texto, cfg_ativo)
                    novo = PF.sanear(dados, base)
                    if not dados.get("nome"):
                        novo["nome"] = f"Norma de {man.name.rsplit('.', 1)[0]}"[:80]
                    ss.rascunho_norma = novo
                    ss.ver_norma = v + 1
                    ss.msg_norma = f"Lido. {len(PF.diferencas(novo))} diferença(s) em relação à EESC: confira as abas."
                    st.rerun(scope="fragment")
                except Exception as e:
                    st.error(f"Não deu para ler o manual: {e}")
        if ss.get("msg_norma"):
            st.success(ss.pop("msg_norma"))

    with aba_val:
        tipos_l = list(PF.TIPOS)
        tipo_n = st.radio("Tipo de trabalho", tipos_l, index=tipos_l.index(base.get("tipo", "tcc")),
                          format_func=PF.TIPOS.get, horizontal=True, key=f"tipo_{v}",
                          help="Define a base: TCC (Diretrizes USP aplicadas ao TCC) ou dissertação/tese "
                               "(template oficial da EESC).")
        nome = st.text_input("Nome da norma", value=base["nome"] if base.get("id") not in ("eesc", "eesc-pos") else
                             "Minha norma (baseada na EESC)", key=f"n13_{v}")
        st.markdown("**Margens (cm)**")
        m = base["margens"]
        c = st.columns(4)
        marg = {k: c[i].number_input(k.capitalize(), 0.5, 6.0, float(m[k]), 0.25, key=f"m_{k}_{v}")
                for i, k in enumerate(("esquerda", "direita", "superior", "inferior"))}
        st.markdown("**Texto**")
        c = st.columns(3)
        fonte = c[0].number_input("Fonte do texto", 8.0, 16.0, float(base["fonte_texto"]), 0.5, key=f"n1_{v}")
        entre = c[1].number_input("Entrelinhas", 1.0, 3.0, float(base["entrelinhas"]), 0.05, key=f"n2_{v}")
        leg = c[2].number_input("Fonte máx. de legendas/fontes", 7.0, 14.0, float(base["fonte_legenda_max"]), 0.5, key=f"n3_{v}")
        st.markdown("**Citação longa (mais de 3 linhas)**")
        cl = base["citacao_longa"]
        c = st.columns(4)
        cl_n = {"recuo": c[0].number_input("Recuo (cm)", 0.5, 8.0, float(cl["recuo"]), 0.25, key=f"n4_{v}"),
                "fonte": c[1].number_input("Fonte aplicada", 7.0, 14.0, float(cl["fonte"]), 0.5, key=f"n5_{v}"),
                "fonte_max": c[2].number_input("Fonte máxima", 7.0, 14.0, float(cl["fonte_max"]), 0.5, key=f"n6_{v}"),
                "entrelinhas": c[3].number_input("Entrelinhas ", 1.0, 2.0, float(cl["entrelinhas"]), 0.05, key=f"n7_{v}")}
        st.markdown("**Resumo**")
        c = st.columns(2)
        rmin = c[0].number_input("Mínimo de palavras", 50, 2000, int(base["resumo"]["min"]), 10, key=f"n8_{v}")
        rmax = c[1].number_input("Máximo de palavras", 50, 2000, int(base["resumo"]["max"]), 10, key=f"n9_{v}")
        st.markdown("**Citações**")
        c = st.columns(2)
        op = list(PF.ROTULOS["citacao_autor"])
        caut = c[0].selectbox("Autor na chamada", op, index=op.index(base["citacao_autor"]),
                              format_func=PF.ROTULOS["citacao_autor"].get, key=f"n10_{v}")
        op = list(PF.ROTULOS["separador_parenteses"])
        csep = c[1].selectbox("Vários autores dentro dos parênteses", op, index=op.index(base["separador_parenteses"]),
                              format_func=PF.ROTULOS["separador_parenteses"].get, key=f"n11_{v}")
        st.markdown("**Referências**")
        rb = base["referencias"]
        c = st.columns(2)
        rtit = c[0].text_input("Título da lista", value=rb["titulo"], key=f"n14_{v}")
        op = list(PF.ROTULOS["destaque"])
        rdest = c[1].selectbox("Destaque do título da obra", op, index=op.index(rb["destaque"]),
                               format_func=PF.ROTULOS["destaque"].get, key=f"n12_{v}")
        c = st.columns(2)
        rord = c[0].toggle("Ordem alfabética", value=rb["ordem_alfabetica"], key=f"n15_{v}")
        resq = c[1].toggle("Alinhadas à esquerda (não justificar)", value=rb["alinhar_esquerda"], key=f"n16_{v}")
        st.markdown("**Elementos conferidos**")
        el = base["elementos"]
        nomes_el = {"abstract": "Abstract", "keywords": "Keywords", "versao": "Versão original/corrigida",
                    "orientador": "Orientador(a) na folha de rosto", "ficha": "Ficha catalográfica",
                    "folha_aprovacao": "Folha de aprovação", "secao_nova_pagina": "Seção primária em nova página",
                    "capa": "Capa (USP, EESC, São Carlos, ano)", "natureza": "Nota de natureza na folha de rosto",
                    "referencia_resumo": "Referência do trabalho acima do resumo",
                    "declaracao_ia": "Declaração de uso de IA generativa"}
        c = st.columns(3)
        el_n = {k: c[i % 3].checkbox(v, value=el[k], key=f"el_{k}_{v}") for i, (k, v) in enumerate(nomes_el.items())}

    with aba_reg:
        st.caption("Regras extras que a sua norma exige. Ex.: proibir 'ou seja' no texto; exigir 'Acesso em:' em "
                   "cada referência; exigir que o documento tenha 'AGRADECIMENTOS'. Marque 'regex' para usar "
                   "expressão regular.")
        regras_df = pd.DataFrame(base.get("regras") or [],
                                 columns=["nome", "tipo", "padrao", "regex", "onde", "status", "como_corrigir"])
        regras_df["regex"] = regras_df["regex"].fillna(False).astype(bool)
        editado = st.data_editor(
            regras_df, num_rows="dynamic", hide_index=True, width="stretch", key=f"ed_regras_{v}",
            column_config={
                "nome": st.column_config.TextColumn("Nome da regra", required=True),
                "tipo": st.column_config.SelectboxColumn("Tipo", options=list(PF.ROTULOS["tipo"]), default="proibir",
                                                         help="; ".join(f"{k}: {v}" for k, v in PF.ROTULOS["tipo"].items())),
                "padrao": st.column_config.TextColumn("Texto/padrão", required=True),
                "regex": st.column_config.CheckboxColumn("regex", default=False),
                "onde": st.column_config.SelectboxColumn("Onde", options=list(PF.ROTULOS["onde"]), default="texto"),
                "status": st.column_config.SelectboxColumn("Gravidade", options=["Ajustar", "Verificar"],
                                                           default="Ajustar"),
                "como_corrigir": st.column_config.TextColumn("Como corrigir"),
            })

    with aba_ia:
        instr = st.text_area("Regras de referências que diferem da ABNT (a IA segue estas na revisão)",
                             value=base.get("instrucoes_ia", ""), height=200,
                             placeholder="Ex.: título da obra em itálico; prenomes sempre por extenso; "
                                         "incluir DOI quando houver.", key=f"n17_{v}")

    regras = []
    for r in editado.to_dict("records"):
        if str(r.get("padrao") or "").strip():
            regras.append({k: (bool(v) if k == "regex" else ("" if _celula_vazia(v) else v)) for k, v in r.items()})
    novo = PF.completo({
        "tipo": tipo_n,
        "nome": nome.strip() or "Minha norma", "margens": marg, "fonte_texto": fonte, "entrelinhas": entre,
        "fonte_legenda_max": leg, "citacao_longa": cl_n, "resumo": {"min": int(rmin), "max": int(rmax)},
        "citacao_autor": caut, "separador_parenteses": csep,
        "referencias": {"titulo": rtit.strip() or "REFERÊNCIAS", "destaque": rdest, "ordem_alfabetica": rord,
                        "alinhar_esquerda": resq},
        "elementos": el_n, "regras": regras, "instrucoes_ia": instr,
        "descricao": base.get("descricao", "") if base.get("id") not in ("eesc", "eesc-pos") else "",
    })
    difs = PF.diferencas(novo)
    st.divider()
    st.caption(f"Diferenças em relação ao padrão EESC de {PF.TIPOS[tipo_n]}: " + ("; ".join(difs) if difs else "nenhuma"))
    c1, c2, c3 = st.columns(3)
    if c1.button("✅ Usar nesta sessão", type="primary", width="stretch", disabled=not difs):
        pid = "sessao-" + PF._slug(novo["nome"])
        novo["id"] = pid
        ss.setdefault("perfis_sessao", {})[pid] = novo
        ss.perfil_id, ss.rascunho_norma, ss.abrir_norma = pid, None, False
        ss.msg = f"Norma '{novo['nome']}' aplicada nesta sessão."
        st.rerun()
    if admin:
        if c2.button("💾 Salvar para todos", width="stretch", disabled=not difs,
                     help="Fica disponível para toda a equipe (e no site, se publicado)."):
            novo["id"] = "" if str(base.get("id", "")).startswith(("eesc", "sessao")) else base.get("id", "")
            pid = PF.salvar(novo)
            ss.perfil_id, ss.rascunho_norma, ss.abrir_norma = pid, None, False
            ss.msg = f"Norma '{novo['nome']}' salva."
            st.rerun()
    c3.download_button("⬇ Baixar (.json)", PF.para_json(novo), file_name=f"norma_{PF._slug(novo['nome'])}.json",
                       mime="application/json", width="stretch")
    if admin and base.get("id") not in (None, "", "eesc", "eesc-pos") and not str(base.get("id")).startswith("sessao"):
        if st.button(f"🗑 Excluir a norma salva '{base['nome']}'"):
            PF.excluir(base["id"])
            ss.perfil_id, ss.rascunho_norma, ss.abrir_norma = "eesc", None, False
            st.rerun()


if ss.pop("abrir_norma", False):
    editor_norma()

# ------------------------------------------------------------------ envio
# ------------------------------------------------------------------ diferenciação inicial: TCC × dissertação/tese
PADRAO_DO_TIPO = {"tcc": "eesc", "pos": "eesc-pos"}
MODELO_POS = PASTA / "assets" / "modelos" / "TEMPLATE_Dissertacao_Tese_ABNT_EESC.docx"


def _ao_escolher_tipo():
    t = ss.get("tipo_escolha") or "auto"
    if t in PADRAO_DO_TIPO and ss.get("perfil_tipo") != t:
        ss.perfil_id = PADRAO_DO_TIPO[t]


ss.setdefault("tipo_escolha", "auto")
c_tipo, c_modelo = st.columns([3, 1], vertical_alignment="bottom")
c_tipo.segmented_control(
    "Que tipo de trabalho é?", ["auto", "tcc", "pos"], key="tipo_escolha", on_change=_ao_escolher_tipo,
    format_func={"auto": "🔎 Detectar pelo arquivo", "tcc": "🎓 TCC (graduação)",
                 "pos": "📚 Dissertação ou Tese"}.get,
    help="TCC usa as Diretrizes USP aplicadas ao TCC; dissertação/tese usa o template oficial da EESC "
         "(folha de rosto, declaração de IA, referência no resumo, agradecimentos CAPES/FAPESP...). "
         "Em 'Detectar', o tipo é lido da folha de rosto.")
if ss.get("perfil_tipo") == "pos" and MODELO_POS.exists():
    c_modelo.download_button("📄 Template EESC (dissertação/tese)", MODELO_POS.read_bytes(),
                             file_name=MODELO_POS.name, width="stretch",
                             mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

st.markdown(
    "<div style='font-size:.9rem;color:var(--aviso-tx);background:var(--aviso);border-radius:8px;"
    "padding:8px 12px;margin:6px 0 6px'>📎 <b>Envie o arquivo do Word (.docx) ou o PDF.</b> No PDF, margens, "
    "fontes e espaçamentos são medidos na própria página; o .docx permite conferir também os estilos e "
    "configurações do Word. .doc (Word antigo), .odt e PDF digitalizado (imagem) não são aceitos. "
    "<span style='opacity:.8'>Para gerar o PDF: no Word, <i>Arquivo > Salvar como > PDF</i> (não imprima nem "
    "digitalize).</span></div>", unsafe_allow_html=True)
arq = st.file_uploader("Arquivo do trabalho (.docx ou .pdf)", type=["docx", "pdf"],
                       help="Arquivo .docx do Word ou PDF gerado pelo editor de texto. .doc, .odt e PDF "
                            "digitalizado não são aceitos.",
                       key=f"envio_{ss.get('n_envio', 0)}")
recuperado = False
if not arq and ss.get("original") is not None and ss.get("nome"):
    recuperado = True                      # já estava aberto nesta sessão
elif not arq:
    recuperado = cofre_recuperar()         # página recarregada: volta ao trabalho aberto
if not arq and not recuperado:
    if PUBLICO:
        st.markdown(
            "<div style='font-size:.86rem;color:var(--aviso-tx);background:var(--aviso);border-radius:8px;"
            "padding:8px 12px;margin:2px 0 10px'>🔒 <b>Privacidade:</b> o arquivo é verificado no computador da "
            "Biblioteca da EESC, fica só na memória enquanto você usa e é apagado em até 1 hora. Não é guardado, "
            "publicado nem usado para outra finalidade. Se você usar a revisão por IA, só o texto das referências "
            "é enviado ao provedor de IA. A verificação é automática e não substitui a revisão da Biblioteca.</div>",
            unsafe_allow_html=True)
    if PUBLICO:
        st.markdown("""<div class="boas-vindas">
      <div class="bv"><div class="ic">📤</div><b>1. Envie o seu arquivo</b><span>Arraste o seu trabalho em .docx (Word) ou em PDF para a área acima.</span></div>
      <div class="bv"><div class="ic">🔎</div><b>2. Veja as marcações</b><span>O trabalho aparece página por página, com cada problema marcado e numerado. Clique numa marcação para ver o que ajustar.</span></div>
      <div class="bv"><div class="ic">💬</div><b>3. Baixe com comentários</b><span>O seu arquivo volta igual, só com comentários (no Word ou no PDF) nos pontos a revisar.</span></div>
      <div class="bv"><div class="ic">✍️</div><b>4. Corrija no original</b><span>Você faz as correções no seu arquivo (Word, LaTeX...). Depois, envie de novo para conferir.</span></div>
    </div>""", unsafe_allow_html=True)
        st.stop()
    st.markdown("""<div class="boas-vindas">
      <div class="bv"><div class="ic">📤</div><b>1. Envie o arquivo</b><span>Arraste o trabalho do aluno em .docx ou PDF para a área acima (.doc não é aceito).</span></div>
      <div class="bv"><div class="ic">🔎</div><b>2. Confira as marcações</b><span>O trabalho aparece como num leitor de PDF, página por página, com cada problema marcado e numerado.</span></div>
      <div class="bv"><div class="ic">💬</div><b>3. Gere os comentários</b><span>Baixe o arquivo do aluno com um comentário em cada ponto: no Word (.docx) ou como notas no PDF. O texto não é alterado.</span></div>
      <div class="bv"><div class="ic">📄</div><b>4. Devolva ao aluno</b><span>Envie o arquivo comentado e o relatório. As correções são feitas pelo aluno.</span></div>
    </div>""", unsafe_allow_html=True)
    st.stop()
if arq is not None:
    original = arq.getvalue()
    try:
        with zipfile.ZipFile(io.BytesIO(original)) as z_:
            docx_ok = "word/document.xml" in z_.namelist()
    except zipfile.BadZipFile:
        docx_ok = False
    eh_pdf_ = arq.name.lower().endswith(".pdf")
    if eh_pdf_ and not original.lstrip()[:5].startswith(b"%PDF"):
        st.error("Este arquivo não é um PDF válido. Gere de novo pelo editor de texto (Arquivo > Salvar como > PDF).")
        st.stop()
    if not eh_pdf_ and (not arq.name.lower().endswith(".docx") or not docx_ok):
        st.error("Este arquivo não é um .docx do Word nem um PDF (pode ser um .doc renomeado). "
                 + ("Envie o arquivo editável do Word" if PUBLICO else "Peça ao aluno o arquivo editável do Word")
                 + " em .docx (Arquivo > Salvar como > Documento do Word) ou o PDF.")
        st.stop()
    h_orig = hashlib.md5(original).hexdigest()
    if ss.get("h_orig") != h_orig:
        ok_lim, _ = limite_ok("analise", 1, CONFIG.get("publico", {}).get("analises_por_hora", 20), 3600)
        if not ok_lim:
            st.warning("Você atingiu o limite de verificações por hora deste site. Tente de novo mais tarde.")
            st.stop()
        ss.update(h_orig=h_orig, original=original, ia=[], nome=arq.name, msg=None)
        try:
            ss.tipo_detectado = TRAB.detectar(Documento(io.BytesIO(original)))[:2]
        except Exception:
            ss.tipo_detectado = (None, None)
        t_det = ss.tipo_detectado[0]
        if ss.tipo_escolha == "auto" and t_det and t_det != ss.get("perfil_tipo") and \
                ss.get("perfil_id", "eesc") in PADRAO_DO_TIPO.values():
            ss.perfil_id = PADRAO_DO_TIPO[t_det]
            ss.msg = f"Detectado: {PF.TIPOS[t_det]}. Usando a norma '{PF.PADROES[t_det]['nome']}'."
            st.rerun()
nome_arq, h_orig = ss.nome, ss.h_orig
cofre_guardar()


# ------------------------------------------------------------------ análise (o arquivo do aluno não é alterado)
try:
    with st.spinner("Analisando o documento..."):
        achados, refs = analisar_bytes(ss.original)
        ia_validos = [a for a in ss.ia if a.trechos and a.trechos[0] in {r.texto for r in refs}]
        todos = achados + ia_validos
        chave_ia = "|".join(a.status + a.item + (a.trechos[0] if a.trechos else "") for a in ia_validos)
        html_doc, serial = visual(h_orig + PF.chave(PERFIL) + "|" + chave_ia, ss.original, todos)
        dados_apont, n_apont = apontamentos(h_orig + PF.chave(PERFIL) + "|" + chave_ia, ss.original, todos)
except Exception as e:
    st.error(f"Não foi possível ler o arquivo: {e}")
    st.stop()

cont = {s: sum(a.status == s for a in todos) for s in (AJUSTAR, VERIFICAR, OK)}

# ------------------------------------------------------------------ cabeçalho do arquivo e downloads
t_det, g_det = ss.get("tipo_detectado") or (None, None)
rot_det = ({"doutorado": "Tese (doutorado)", "mestrado": "Dissertação (mestrado)"}.get(g_det) or
           (PF.TIPOS[t_det] if t_det else "tipo não identificado"))
tags = [f"<span class='tag'>{'🎓' if t_det == 'tcc' else '📚' if t_det else '❔'} {rot_det}</span>",
        f"<span class='tag'>{len(refs)} referências</span>",
        f"<span class='tag'>📐 {html.escape(PERFIL['nome'])}</span>"]
if tem_alteracoes(ss.original):
    tags.append("<span class='tag amarela' title='O aluno enviou o arquivo com alterações controladas ainda não "
                "aceitas. O verificador lê o texto como se elas estivessem aceitas.'>com alterações controladas "
                "do aluno</span>")
c_arq, c_fechar = st.columns([8, 1], vertical_alignment="center")
c_arq.markdown(f"<div class='arquivo'><span class='nome'>📄 {html.escape(nome_arq)}</span>{''.join(tags)}</div>",
               unsafe_allow_html=True)
if c_fechar.button("✖ Fechar", width="stretch", help="Fecha este documento e apaga da memória do verificador"):
    cofre_fechar()
    ss.n_envio = ss.get("n_envio", 0) + 1  # limpa o campo de envio
    st.rerun()

k = st.columns([1, 1, 1, 1, 1.6])
kpis = [("#D9475B", "Ajustar", cont[AJUSTAR], "fora da norma"),
        ("#E0A800", "Conferir", cont[VERIFICAR], "pedem conferência manual"),
        ("#2F9E5B", "Em conformidade", cont[OK], "itens conferidos"),
        ("#1F4E79", "Comentários", n_apont, "no arquivo para o aluno")]
for col, (cor, rot, val, sub) in zip(k, kpis):
    col.markdown(f"<div class='kpi'><div class='rot'><span class='dot' style='background:{cor}'></span>{rot}</div>"
                 f"<div class='val'>{val}</div><div class='sub'>{sub}</div></div>", unsafe_allow_html=True)
base = re.sub(r"(_(limpo|revisado|corrigido|apontamentos|comentado|marcado))+$", "", nome_arq.rsplit(".", 1)[0])
with k[4]:
    if nome_arq.lower().endswith(".pdf"):
        botao_arquivo(st, "⬇ PDF com marcações (.pdf)", dados_apont, f"{base}_marcado.pdf", type="primary",
                      width="stretch", mime="application/pdf",
                      help=f"O PDF do aluno exatamente como foi enviado, com {n_apont} nota(s): cada trecho realçado "
                           "(vermelho: ajustar; amarelo: conferir) traz o que ajustar e como. Abre em qualquer leitor "
                           "de PDF (Acrobat, navegador, Foxit). Nenhuma palavra é alterada.")
        st.caption("O PDF do aluno não é alterado: o verificador só aponta. Quem corrige é o aluno, no arquivo original.")
    else:
        botao_arquivo(st, "⬇ Arquivo com comentários (.docx)", dados_apont, f"{base}_comentado.docx", type="primary",
                      width="stretch", help=f"O arquivo do aluno exatamente como foi enviado, com {n_apont} "
                                            "comentário(s) do Word na margem dizendo o que ajustar e como (trechos "
                                            "realçados em amarelo). Nenhuma palavra do texto é alterada.")
        st.caption("O texto do aluno não é alterado: o verificador só aponta. Quem corrige é o aluno, no Word.")

# ------------------------------------------------------------------ visualizador (somente leitura)
area_trabalho(achados=serial, html=html_doc, versao=h_orig + PF.chave(PERFIL) + chave_ia, mensagem=ss.msg,
              altura=880, escuro=ESCURO, key="area", default=None)
ss.msg = None

st.write("")
aba_el, aba_ia, aba_rel = st.tabs(["🧾 Elementos do trabalho", "🤖 Revisão das referências por IA", "📄 Relatório"])

# ------------------------------------------------------------------ elementos: aparece · formatação · conteúdo
with aba_el:
    resumo_el = next((a for a in todos if a.item == RESUMO_ELEMENTOS), None)
    if resumo_el and resumo_el.tabela:
        lin = resumo_el.tabela
        faltam = [x["Elemento"] for x in lin if x["Aparece"].startswith("❌")]
        c1, c2, c3 = st.columns(3)
        c1.metric("Encontrados", sum(x["Aparece"].startswith("✅") for x in lin))
        c2.metric("Obrigatórios faltando", len(faltam))
        c3.metric("Com algo a conferir", sum(("❌" in x["Formatação"] + x["Conteúdo confere"]) or
                                             ("⚠️" in x["Formatação"] + x["Conteúdo confere"]) for x in lin))
        st.caption("Cada elemento é procurado no arquivo inteiro (não precisa estar na mesma página do modelo). "
                   "**Aparece**: existe no trabalho · **Formatação**: título, alinhamento, recuos e regras de "
                   "forma · **Conteúdo confere**: o que está escrito bate com o resto do trabalho (autor, título e "
                   "ano iguais, sumário e listas com as páginas e títulos reais, siglas usadas, figuras/seções "
                   "citadas que existem, obras nas referências...). Os detalhes de cada ❌/⚠️ estão na lista de "
                   "pontos a revisar, com o trecho marcado no documento.")
        so_tem = st.toggle("Esconder opcionais que o trabalho não tem", value=True)
        mostrar = [x for x in lin if not (so_tem and x["Aparece"].startswith("—"))]
        tabela(mostrar)
    else:
        st.info("Não foi possível montar a tabela de elementos para este arquivo.")

# ------------------------------------------------------------------ IA
with aba_ia:
    pronto = provedor in configuracao.provedores_prontos(CONFIG)
    nome_p = ia_provedores.PROVEDORES[provedor].split(" —")[0]
    if not refs:
        st.caption("Nenhuma referência extraída do documento.")
    elif not pronto:
        st.info(f"{nome_p} ainda não está configurado. Abra a barra lateral (seta no canto superior esquerdo) "
                "e escolha ou configure um provedor de IA.")
    else:
        nuvem = provedor in ia_provedores.NUVEM
        if nuvem:
            n_ped = ia_provedores.pedidos_necessarios(len(refs), cfg_ativo)
            st.caption(f"{len(refs)} referências · {nome_p} · {cfg_ativo['modelo']} · em lotes de "
                       f"{ia_provedores.tamanho_lote(cfg_ativo)}: **{n_ped} pedido(s)** ao provedor. Costuma levar "
                       "poucos segundos. Só o texto das referências é enviado. O uso pode gerar custo ou consumir a "
                       "cota do plano gratuito.")
        else:
            st.caption(f"{len(refs)} referências · IA local · {cfg_ativo['modelo']}. Sem placa de vídeo, conte de 15 "
                       "a 60 s por referência. Pode deixar rodando; não feche esta aba.")
        c1, c2 = st.columns([1, 3], vertical_alignment="bottom")
        limite = c1.number_input("Quantas revisar", 1, len(refs), len(refs) if nuvem else min(len(refs), 20))
        if c2.button(f"Revisar {limite} referência(s) com {cfg_ativo['modelo']}", type="primary") and \
                ia_permitida(limite):
            barra = st.progress(0.0, text="Enviando..." if nuvem else "Carregando o modelo...")
            t0 = time.time()

            def progresso(feitos, total):
                passo = (time.time() - t0) / feitos
                barra.progress(feitos / total, text=f"{feitos}/{total} · faltam cerca de {int(passo * (total - feitos))} s")

            resultados, erros = ia_provedores.revisar_varias(refs[:limite], cfg_ativo, progresso)
            # guarda o que já foi revisado (inclusive de rodadas anteriores) mesmo se a cota acabar no meio
            novos = {a.trechos[0]: a for a in resultados if a.trechos}
            ss.ia = [a for a in ss.ia if a.trechos and a.trechos[0] not in novos] + list(novos.values())
            if erros:
                ss.erros_ia = list(dict.fromkeys(erros))
            ss.msg = f"{len(resultados)} referência(s) revisada(s) pela IA."
            st.rerun()
        for e in ss.pop("erros_ia", []):
            st.warning(e, icon="⚠️")
        if ia_validos:
            n_aj = sum(a.status == AJUSTAR for a in ia_validos)
            st.success(f"{len(ia_validos)} referências revisadas; {n_aj} com sugestão de ajuste. "
                       "Elas aparecem marcadas no documento (categoria 'Referências (IA)') e entram como "
                       "comentário no arquivo do aluno.")

# ------------------------------------------------------------------ relatório
with aba_rel:
    c1, c2 = st.columns([2, 1])
    obs = c1.text_area("Observações da Biblioteca (entram no relatório)", height=120)
    incluir_ok = c2.checkbox("Incluir itens em conformidade", value=False)
    chave_rel = "|".join([h_orig, PF.chave(PERFIL), time.strftime("%Y-%m-%d"), chave_ia])  # leva a data de hoje
    botao_arquivo(c2, "⬇ Baixar relatório (.docx)",
                  relatorio_docx(chave_rel, todos, nome_arq, incluir_ok,
                                 ("" if PF.eh_padrao(PERFIL) else f"Norma usada: {PERFIL['nome']}.\n") + obs),
                  f"Relatorio_ABNT_{base}.docx", type="primary", width="stretch")
    c2.caption("Lista o que precisa ser ajustado, com a seção das Diretrizes. Revise antes de enviar ao aluno.")
    pend = sorted((a for a in todos if a.status != OK), key=lambda a: (a.status != AJUSTAR, a.categoria))
    if pend:
        st.markdown("**Pendências que vão no relatório:**")
        for a in pend:
            st.markdown(f"{pill(a.status)}&nbsp; {html.escape(a.item if a.origem != 'ia' else 'Referência (IA)')} "
                        f"<span style='color:#6B7280'>— {html.escape(a.detalhe)}</span>", unsafe_allow_html=True)
