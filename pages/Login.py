# páginas/Login.py
import streamlit as st
import json, gspread, pytz
from datetime import datetime
from oauth2client.service_account import ServiceAccountCredentials

from utils.sessoes import registrar_sessao_assumindo

st.set_page_config(page_title="Login | MMR Consultoria")

# CSS para esconder a barra
st.markdown("""
<style>
[data-testid="stToolbar"] { visibility: hidden; height: 0%; position: fixed; }
</style>
""", unsafe_allow_html=True)
# Se já estiver logado, vai direto pro Home
if st.session_state.get("acesso_liberado"):
    st.switch_page("Home")   # funciona porque Home.py está na raiz
    st.stop()
# Parâmetros opcionais
params = st.query_params
codigo_param = (params.get("codigo") or "").strip()
empresa_param = (params.get("empresa") or "").strip().lower()
if not codigo_param or not empresa_param:
    st.warning("⚠️ Acesso direto sem parâmetros. Você pode logar normalmente abaixo.")

# Usuários permitidos
USUARIOS = [
    {"codigo": "1825", "email": "carlos.soveral@grupofit.com.br", "senha": "$%252M"},
    {"codigo": "1825", "email": "maricelisrossi@gmail.com", "senha": "1825o"},
    {"codigo": "1825", "email": "vanessa.carvalho@grupofit.com.br", "senha": "%6790"},
    {"codigo": "1825", "email": "rosana.rocha@grupofit.com.br", "senha": "hjk&54lmhp"},
    {"codigo": "1825", "email": "debora@grupofit.com.br", "senha": "klom52#@$65"},
    {"codigo": "1825", "email": "renata.favacho@grupofit.com.br", "senha": "Huom63@#$52"},
    {"codigo": "1825", "email": "marcos.bogli@grupofit.com.br", "senha": "Ahlk52@#$81"},
    {"codigo": "1825", "email": "contabilidade@grupofit.com.br", "senha": "hYhIO18@#$21"},
    {"codigo": "1825", "email": "larissa.esthefani@grupofit.com.br", "senha": "OjKlo252@#$%$21"},
    {"codigo": "1825", "email": "contasareceber_01@grupofit.com.br", "senha": "kird*$#@&Mklo*21"},
    {"codigo": "3377", "email": "maricelisrossi@gmail.com", "senha": "1825"}
]

# Google Sheets
PLANILHA_KEY = "1SZ5R6hcBE6o_qWs0_wx6IGKfIGltxpb9RWiGyF4L5uE"
SHEET_SESSOES = "SessõesAtivas"
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT_ACESSOS"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)



# Formulário de login
st.title("🔐 Acesso Restrito")
st.markdown("Informe o código da empresa, e-mail e senha.")

codigo = st.text_input("Código da Empresa:", value=codigo_param)
email = st.text_input("E-mail:")
senha = st.text_input("Senha:", type="password")

if st.button("Entrar"):
    usuario = next((u for u in USUARIOS if u["codigo"] == codigo and u["email"] == email and u["senha"] == senha), None)
    if usuario:
        token = registrar_sessao_assumindo(gc, PLANILHA_KEY, SHEET_SESSOES, email)
        st.session_state["sessao_token"] = token
        st.session_state["acesso_liberado"] = True
        st.session_state["empresa"] = codigo
        st.session_state["usuario_logado"] = email
        st.switch_page("Home")   # ✅ redireciona direto
    else:
        st.error("❌ Código, e-mail ou senha incorretos.")
