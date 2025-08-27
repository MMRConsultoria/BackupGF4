# pages/Login.py
import streamlit as st
import json
import uuid
import pytz
from datetime import datetime

import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Login | MMR Consultoria")

# =====================================
# CSS: esconder apenas a barra superior
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

# =====================================
# Parâmetros opcionais da URL (não bloqueiam mais a tela)
# =====================================
params = st.query_params
codigo_param = (params.get("codigo") or "").strip()
empresa_param = (params.get("empresa") or "").strip().lower()
if not codigo_param or not empresa_param:
    st.warning("⚠️ Acesso direto sem parâmetros. Você pode logar normalmente abaixo.")

# =====================================
# Usuários autorizados
# =====================================
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

# =====================================
# Google Sheets
# =====================================
PLANILHA_KEY = "1SZ5R6hcBE6o_qWs0_wx6IGKfIGltxpb9RWiGyF4L5uE"
SHEET_SESSOES = "SessõesAtivas"

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT_ACESSOS"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)

# =====================================
# Registro de acessos (sheet1)
# =====================================
def registrar_acesso(nome_usuario, acao="LOGIN"):
    try:
        fuso = pytz.timezone("America/Sao_Paulo")
        agora = datetime.now(fuso)
        data = agora.strftime("%d/%m/%Y")
        hora = agora.strftime("%H:%M:%S")

        planilha = gc.open_by_key(PLANILHA_KEY)
        aba = planilha.sheet1
        vals = aba.get_all_values()
        if not vals:
            aba.append_row(["Usuario", "Data", "Hora", "Acao"])
        aba.append_row([nome_usuario, data, hora, acao])
    except Exception as e:
        st.warning(f"Não foi possível registrar acesso: {e}")

# =====================================
# Sessões: login sempre assume (derruba anterior)
# =====================================
def _open_aba_sessoes():
    planilha = gc.open_by_key(PLANILHA_KEY)
    try:
        aba = planilha.worksheet(SHEET_SESSOES)
    except:
        aba = planilha.add_worksheet(title=SHEET_SESSOES, rows=200, cols=6)
        aba.update("A1:E1", [["email", "token", "data", "hora", "ultimo_acesso"]])
    return aba

def _liberar_sessao(email: str):
    """Remove TODAS as linhas desse e-mail (derruba sessão anterior)."""
    aba = _open_aba_sessoes()
    todas = aba.get_all_values()
    if not todas:
        return
    novas = [todas[0]] + [row for row in todas[1:] if row and row[0] != email]
    aba.clear()
    aba.update("A1", novas)

def registrar_sessao_assumindo(email: str):
    """
    Login preemptivo: apaga sessão anterior desse e-mail e cria uma nova.
    Salva token no session_state para validação na Home.
    """
    aba = _open_aba_sessoes()

    # derruba sessão anterior
    _liberar_sessao(email)

    # cria nova sessão
    fuso = pytz.timezone("America/Sao_Paulo")
    agora = datetime.now(fuso)
    token = str(uuid.uuid4())
    nova = [email, token, agora.strftime("%d/%m/%Y"), agora.strftime("%H:%M:%S"), agora.isoformat()]
    aba.append_row(nova)

    st.session_state["sessao_token"] = token
    return True

def atualizar_sessao(email: str):
    """Renova carimbo de tempo da sessão atual (se existir)."""
    try:
        aba = _open_aba_sessoes()
        todas = aba.get_all_values()
        if not todas:
            return
        fuso = pytz.timezone("America/Sao_Paulo")
        agora = datetime.now(fuso)
        for i in range(1, len(todas)):
            row = todas[i]
            if row and row[0] == email:
                if len(row) < 5:
                    row += [""] * (5 - len(row))
                row[2] = agora.strftime("%d/%m/%Y")
                row[3] = agora.strftime("%H:%M:%S")
                row[4] = agora.isoformat()
                todas[i] = row
        aba.clear()
        aba.update("A1", todas)
    except Exception:
        # não travar o app por isso
        pass

def validar_sessao(email: str, token: str) -> bool:
    """Confere se o token na planilha ainda é o mesmo desta máquina."""
    try:
        aba = _open_aba_sessoes()
        registros = aba.get_all_records()
        for r in registros:
            if r.get("email") == email:
                return r.get("token") == token
        return False  # não achou sessão ativa para este e-mail
    except Exception:
        return False

def encerrar_sessao(email: str):
    """Logout explícito (opcional)."""
    try:
        _liberar_sessao(email)
        registrar_acesso(email, acao="LOGOUT")
    except Exception:
        pass

# =====================================
# Já logado? Vai pra Home
# =====================================
if st.session_state.get("acesso_liberado"):
    st.switch_page("Home.py")

# =====================================
# Tela de login
# =====================================
st.title("🔐 Acesso Restrito")
st.markdown("Informe o código da empresa, e-mail e senha.")

codigo = st.text_input("Código da Empresa:", value=codigo_param)
email = st.text_input("E-mail:")
senha = st.text_input("Senha:", type="password")

if st.button("Entrar"):
    usuario = next(
        (u for u in USUARIOS if u["codigo"] == codigo and u["email"] == email and u["senha"] == senha),
        None
    )
    if usuario:
        # Login sempre assume: derruba sessão antiga e cria nova
        if registrar_sessao_assumindo(email):
            st.session_state["acesso_liberado"] = True
            st.session_state["empresa"] = codigo
            st.session_state["usuario_logado"] = email
            registrar_acesso(email, acao="LOGIN")
            st.switch_page("Home.py")
    else:
        st.error("❌ Código, e-mail ou senha incorretos.")
