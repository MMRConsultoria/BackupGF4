# Home.py

import streamlit as st
import time, hashlib, glob, os

# =====================================
# CSS para esconder barra de botões do canto superior direito
# =====================================
st.markdown("""
    <style>
        [data-testid="stToolbar"] {
            visibility: hidden;
            height: 0%;
            position: fixed;
        }
    </style>
""", unsafe_allow_html=True)

# ⚙️ Config da página (sempre no topo)
st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")

# 🔎 Indicadores para provar o deploy
st.sidebar.write("🔄 Build time:", time.strftime("%Y-%m-%d %H:%M:%S"))

def app_version():
    h = hashlib.sha256()
    for p in sorted(glob.glob("**/*.py", recursive=True) + ["requirements.txt"]):
        if os.path.exists(p):
            with open(p, "rb") as f:
                h.update(f.read())
    return h.hexdigest()[:8]

st.sidebar.caption(f"🧩 Versão do app: {app_version()}")

# (Opcional) limpar cache via URL ?nocache=1
# ✅ novo (compatível com 1.49+)
nocache = st.query_params.get("nocache", "0")
if isinstance(nocache, list):  # st.query_params pode retornar lista
    nocache = nocache[0] if nocache else "0"

if nocache == "1":
    st.cache_data.clear()
    st.warning("🧹 Cache limpo via ?nocache=1")

# ✅ Gate de login
if not st.session_state.get("acesso_liberado"):
    st.switch_page("pages/Login.py")
    st.stop()

# ✅ Código da empresa logada
codigo_empresa = st.session_state.get("empresa")

# ✅ Logos por código
LOGOS_CLIENTES = {
    "1825": "https://raw.githubusercontent.com/MMRConsultoria/MMRBackup/main/logo_grupofit.png",
    "3377": "https://raw.githubusercontent.com/MMRConsultoria/MMRBackup/main/rossi_ferramentas_logo.png",
    "0041": "https://raw.githubusercontent.com/MMRConsultoria/MMRBackup/main/logo_empresa3.png",
}

# ✅ Logo na sidebar
logo_cliente = LOGOS_CLIENTES.get(codigo_empresa)
if logo_cliente:
    st.sidebar.markdown(
        f"""
        <div style="text-align: center; padding: 10px 0 30px 0;">
            <img src="{logo_cliente}" width="100">
        </div>
        """,
        unsafe_allow_html=True,
    )

# ✅ Logo principal
st.image(logo_cliente or "https://raw.githubusercontent.com/MMRConsultoria/MMRBackup/main/logo-mmr.png", width=150)

# ✅ Mensagem
st.markdown("## Bem-vindo ao Portal de Relatórios")
st.success(f"✅ Acesso liberado para o código {codigo_empresa}!")

# ======================
# Botão de Logout
# ======================
if "usuario_logado" in st.session_state:
    st.markdown("---")
    st.caption(f"🔑 Logado como: {st.session_state['usuario_logado']}")

    if st.button("Sair"):
        try:
            # importa a função que criamos no Login.py
            from pages.Login import encerrar_sessao  
            encerrar_sessao(st.session_state["usuario_logado"])
        except Exception as e:
            st.warning(f"Não foi possível encerrar sessão no servidor: {e}")

        # limpa session_state local
        for k in ["acesso_liberado", "empresa", "usuario_logado", "sessao_token"]:
            st.session_state.pop(k, None)

        st.rerun()
