# Home.py
import streamlit as st
import time, hashlib, glob, os, json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from utils.sessoes import validar_sessao, atualizar_sessao

st.set_page_config(page_title="Portal de Relatórios | MMR Consultoria")

# CSS
st.markdown("""
<style>
[data-testid="stToolbar"] { visibility: hidden; height: 0%; position: fixed; }
</style>
""", unsafe_allow_html=True)

# Sidebar build info
st.sidebar.write("🔄 Build time:", time.strftime("%Y-%m-%d %H:%M:%S"))
def app_version():
    h = hashlib.sha256()
    for p in sorted(glob.glob("**/*.py", recursive=True) + ["requirements.txt"]):
        if os.path.exists(p):
            with open(p, "rb") as f: h.update(f.read())
    return h.hexdigest()[:8]
st.sidebar.caption(f"🧩 Versão do app: {app_version()}")

# nocache
nocache = st.query_params.get("nocache", "0")
if isinstance(nocache, list): nocache = nocache[0] if nocache else "0"
if nocache == "1":
    st.cache_data.clear()
    st.warning("🧹 Cache limpo via ?nocache=1")

# Gate de login
if not st.session_state.get("acesso_liberado"):
    st.switch_page("Login")   # porque Login.py está em /pages
    st.stop()
# Google Sheets client
PLANILHA_KEY = "1SZ5R6hcBE6o_qWs0_wx6IGKfIGltxpb9RWiGyF4L5uE"
SHEET_SESSOES = "SessõesAtivas"
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT_ACESSOS"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)

# Valida e renova sessão
email_atual = st.session_state.get("usuario_logado")
token_atual = st.session_state.get("sessao_token")

if not email_atual or not token_atual or not validar_sessao(gc, PLANILHA_KEY, SHEET_SESSOES, email_atual, token_atual):
    for k in ["acesso_liberado", "empresa", "usuario_logado", "sessao_token"]:
        st.session_state.pop(k, None)
    st.warning("Sua sessão foi encerrada. Faça login novamente.")
    st.switch_page("Login")
    st.stop()

# Mantém a sessão viva
atualizar_sessao(gc, PLANILHA_KEY, SHEET_SESSOES, email_atual)

# --- Conteúdo original ---
codigo_empresa = st.session_state.get("empresa")
LOGOS_CLIENTES = {
    "1825": "https://raw.githubusercontent.com/MMRConsultoria/MMRBackup/main/logo_grupofit.png",
    "3377": "https://raw.githubusercontent.com/MMRConsultoria/MMRBackup/main/rossi_ferramentas_logo.png",
    "0041": "https://raw.githubusercontent.com/MMRConsultoria/MMRBackup/main/logo_empresa3.png",
}
logo_cliente = LOGOS_CLIENTES.get(codigo_empresa)
if logo_cliente:
    st.sidebar.markdown(f"<div style='text-align:center;padding:10px 0 30px 0;'><img src='{logo_cliente}' width='100'></div>", unsafe_allow_html=True)

st.image(logo_cliente or "https://raw.githubusercontent.com/MMRConsultoria/MMRBackup/main/logo-mmr.png", width=150)
st.markdown("## Bem-vindo ao Portal de Relatórios")
st.success(f"✅ Acesso liberado para o código {codigo_empresa}!")
