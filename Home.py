# Home.py

import streamlit as st
import time, hashlib, glob, os

# ⚙️ Config da página (sempre no topo)
st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")

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
# ✅ compatível com 1.49+
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

# 🔒 Validação de posse da sessão + renovação de timeout
from pages.Login import validar_sessao, atualizar_sessao

email_atual = st.session_state.get("usuario_logado")
token_atual = st.session_state.get("sessao_token")

if not email_atual or not token_atual or not validar_sessao(email_atual, token_atual):
    # Sessão foi assumida por outra máquina (ou não existe mais)
    for k in ["acesso_liberado", "empresa", "usuario_logado", "sessao_token"]:
        st.session_state.pop(k, None)
    st.warning("Sua sessão foi encerrada (acessada em outro dispositivo). Faça login novamente.")
    st.switch_page("pages/Login.py")
    st.stop()

# 🔄 Mantém a sessão viva enquanto o usuário navega
atualizar_sessao(email_atual)

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
