import streamlit as st
import time
import hashlib
import glob
import os
from datetime import datetime
from zoneinfo import ZoneInfo

# =====================================
# Esconder barra de ferramentas do Streamlit
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

# ⚙️ Configuração da página (deve vir antes de qualquer elemento)
st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")

# 🕒 Mostrar hora de build na sidebar
now_br = datetime.now(ZoneInfo("America/Sao_Paulo"))
st.sidebar.write("🔄 Build time (Brasília):", now_br.strftime("%Y-%m-%d %H:%M:%S"))

# 🔢 Função para gerar versão do app com base nos arquivos
def app_version():
    h = hashlib.sha256()
    for p in sorted(glob.glob("**/*.py", recursive=True) + ["requirements.txt"]):
        if os.path.exists(p):
            with open(p, "rb") as f:
                h.update(f.read())
    return h.hexdigest()[:8]

st.sidebar.caption(f"🧩 Versão do app: {app_version()}")

# 🧪 Mostrar versão do Streamlit
st.sidebar.write(f"🐍 Versão do Streamlit: {st.__version__}")

# 🧹 Limpar cache se ?nocache=1 estiver na URL
nocache = st.query_params.get("nocache", "0")
if isinstance(nocache, list):
    nocache = nocache[0] if nocache else "0"

if nocache == "1":
    st.cache_data.clear()
    st.cache_resource.clear()
    st.warning("🧹 Cache limpo via ?nocache=1")

# 🔐 Gate de login
if not st.session_state.get("acesso_liberado"):
    st.switch_page("pages/Login.py")
    st.stop()

# 🏢 Código da empresa logada
codigo_empresa = st.session_state.get("empresa")

# 🖼️ Logos dos clientes
LOGOS_CLIENTES = {
    "1825": "https://raw.githubusercontent.com/MMRConsultoria/MMRBackup/main/logo_grupofit.png",
    "3377": "https://raw.githubusercontent.com/MMRConsultoria/MMRBackup/main/rossi_ferramentas_logo.png",
    "0041": "https://raw.githubusercontent.com/MMRConsultoria/MMRBackup/main/logo_empresa3.png",
}

# 🖼️ Logo na sidebar
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

# 🖼️ Logo principal
st.image(logo_cliente or "https://raw.githubusercontent.com/MMRConsultoria/MMRBackup/main/logo-mmr.png", width=150)

# 🎉 Mensagem de boas-vindas
st.markdown("## Bem-vindo ao Portal de Relatórios")
st.success(f"✅ Acesso liberado para o código {codigo_empresa}!")
