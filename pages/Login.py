# pages/Login.py
import streamlit as st
import json
import pytz
from datetime import datetime

import gspread
from oauth2client.service_account import ServiceAccountCredentials

from utils.sessoes import (
    registrar_sessao_assumindo, validar_sessao,
    atualizar_sessao, encerrar_sessao
)

st.set_page_config(page_title="Login | MMR Consultoria")

# CSS – esconder toolbar
st.markdown("""
<style>
  [data-testid="stToolbar"] { visibility: hidden; height: 0%; position: fixed; }
</style>
""", unsafe_allow_html=True)

# Parâmetros (opcionais)
params = st.query_params
codigo_param = (params.get("codigo") or "").strip()
empresa_param = (params.get("empresa") or "").strip().lower()
if not codigo_param or not empresa_param:
    st.warning("⚠️ Acesso direto sem parâmetros. Você pode logar normalmente abaixo.")

# Usuários
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
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT_ACESSOS"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)

# Registro de acesso
def registrar_acesso(nome_usuario, acao="LOGIN"):
    try:
        fuso = pytz.timezone("America/Sao_Paulo")
        agora = datetime.now(fuso)
        data = agora.strftime("%d/%m/%Y"); hora = agora.strftime("%H:%M:%S")
        planilha = gc.open_by_key(PLANILHA_KEY)
        aba = planilha.sheet1
        vals = aba.get_all_values()
        if not vals:
            aba.append_row(["Usuario", "Data", "Hora", "Acao"])
        aba.append_row([nome_usuario, data, hora, acao])
    except Exception as e:
        st.caption(f"ℹ️ Não foi possível registrar acesso: {e}")

# Já logado? Vai pra Home
if st.session_state.get("acesso_liberado"):
    st.switch_page("Home.py")

# Login UI
st.title("🔐 Acesso Restrito")
st.markdown("Informe o código da empresa, e-mail e senha.")

codigo = st.text_input("Código da Empresa:", value=codigo_param)
email = st.text_input("E-mail:")
senha = st.text_input("Senha:", type="password")

if st.button("Entrar"):
    usuario = next(
        (u for u in USUARIOS if u["codigo"] == codigo and u["email"] == email and u["senha"] == senha), None
    )
    if usuario:
        # Login sempre assume: derruba sessão antiga e cria nova
        token = registrar_sessao_assumindo(gc, PLANILHA_KEY, email)
        st.session_state["sessao_token"] = token
        st.session_state["acesso_liberado"] = True
        st.session_state["empresa"] = codigo
        st.session_state["usuario_logado"] = email
        registrar_acesso(email, acao="LOGIN")
        st.switch_page("Home.py")
    else:
        st.error("❌ Código, e-mail ou senha incorretos.")
