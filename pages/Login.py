# pages/Login.py
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from datetime import datetime
import pytz
import uuid

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
# Sessão única + timeout + forçar login
# =====================================
NOME_ABA_SESSOES = "SessõesAtivas"
SESSION_TIMEOUT_MIN = 30  # ajuste aqui

def _open_aba_sessoes():
    planilha = gc.open_by_key(PLANILHA_KEY)
    try:
        aba = planilha.worksheet(NOME_ABA_SESSOES)
    except:
        aba = planilha.add_worksheet(title=NOME_ABA_SESSOES, rows=200, cols=6)
        aba.update("A1:E1", [["email", "token", "data", "hora", "ultimo_acesso"]])
    return aba

def get_sessoes_ativas():
    try:
        aba = _open_aba_sessoes()
        registros = aba.get_all_records()
        return aba, registros
    except Exception as e:
        st.error(f"Erro ao acessar sessões ativas: {e}")
        return None, []

def _liberar_sessao(email):
    """Remove todas as linhas da sessão desse e-mail."""
    aba = _open_aba_sessoes()
    todas = aba.get_all_values()
    if not todas:
        return
    novas = [todas[0]] + [row for row in todas[1:] if row and row[0] != email]
    aba.clear()
    aba.update("A1", novas)

def registrar_sessao(email, force=False):
    aba, registros = get_sessoes_ativas()
    if not aba:
        return False

    fuso = pytz.timezone("America/Sao_Paulo")
    agora = datetime.now(fuso)

    # verifica sessão existente
    existente = None
    for r in registros:
        if r.get("email") == email:
            existente = r
            break

    if existente:
        # idade da sessão
        try:
            ultimo = datetime.strptime(f"{existente['data']} {existente['hora']}", "%d/%m/%Y %H:%M:%S")
        except Exception:
            ultimo = agora  # conservador
        diff_min = (agora - ultimo).total_seconds() / 60.0
        if diff_min < SESSION_TIMEOUT_MIN and not force:
            return False  # ainda válida e sem forçar
        # expirou ou force=True => libera antiga
        _liberar_sessao(email)

    # cria nova sessão
    token = str(uuid.uuid4())
    nova = [email, token, agora.strftime("%d/%m/%Y"), agora.strftime("%H:%M:%S"), agora.isoformat()]
    aba.append_row(nova)
    st.session_state["sessao_token"] = token
    return True

def atualizar_sessao(email):
    """Renova data/hora/ultimo_acesso enquanto o usuário navega."""
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
    except Exception as e:
        # não travar app por isso
        st.caption(f"ℹ️ Sessão não renovada: {e}")

def encerrar_sessao(email):
    """Logout explícito (opcional no Home)."""
    try:
        _liberar_sessao(email)
        registrar_acesso(email, acao="LOGOUT")
    except Exception as e:
        st.warning(f"Não foi possível encerrar sessão: {e}")

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

# Guarda tentativa para o botão "Forçar login"
if "pending_login" not in st.session_state:
    st.session_state["pending_login"] = {}

if st.button("Entrar"):
    usuario = next(
        (u for u in USUARIOS if u["codigo"] == codigo and u["email"] == email and u["senha"] == senha),
        None
    )
    if usuario:
        ok = registrar_sessao(email, force=False)
        if ok:
            st.session_state["acesso_liberado"] = True
            st.session_state["empresa"] = codigo
            st.session_state["usuario_logado"] = email
            registrar_acesso(email, acao="LOGIN")
            st.switch_page("Home.py")
        else:
            st.error("⚠️ Esse usuário já está logado em outra máquina.")
            st.info("Se a sessão anterior travou, você pode liberar e entrar agora.")
            st.session_state["pending_login"] = {"codigo": codigo, "email": email, "senha": senha}
    else:
        st.error("❌ Código, e-mail ou senha incorretos.")

# Botão FORÇAR LOGIN (mostra só após bloqueio)
if st.session_state.get("pending_login", {}).get("email"):
    if st.button("⚡ Liberar sessão anterior e entrar agora"):
        pend = st.session_state["pending_login"]
        usuario = next(
            (u for u in USUARIOS if u["codigo"] == pend["codigo"] and u["email"] == pend["email"] and u["senha"] == pend["senha"]),
            None
        )
        if usuario:
            ok = registrar_sessao(pend["email"], force=True)
            if ok:
                st.session_state["acesso_liberado"] = True
                st.session_state["empresa"] = pend["codigo"]
                st.session_state["usuario_logado"] = pend["email"]
                st.session_state["pending_login"] = {}
                registrar_acesso(pend["email"], acao="FORCE_LOGIN")
                st.switch_page("Home.py")
            else:
                st.error("Não foi possível assumir a sessão agora. Tente novamente.")
        else:
            st.error("Credenciais inválidas ao forçar login. Tente novamente.")
