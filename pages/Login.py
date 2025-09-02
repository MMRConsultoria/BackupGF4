import streamlit as st
import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from datetime import datetime
import streamlit as st

st.set_page_config(page_title="Login | MMR Consultoria")

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

# ✅ Captura segura dos parâmetros da URL
params = st.query_params
codigo_param = (params.get("codigo") or "").strip()
empresa_param = (params.get("empresa") or "").strip().lower()

# ✅ Bloqueia acesso direto sem parâmetros
if not codigo_param or not empresa_param:
    st.markdown("""
        <meta charset="UTF-8">
        <style>
        #MainMenu, header, footer, .stSidebar, .stToolbar, .block-container { display: none !important; }
        body {
          background-color: #ffffff;
          font-family: Arial, sans-serif;
          display: flex;
          align-items: center;
          justify-content: center;
          height: 100vh;
          margin: 0;
        }
        </style>
        <div style="text-align: center;">
            <h2 style="color:#555;">🚫 Acesso Negado</h2>
            <p style="color:#888;">Você deve acessar pelo <strong>portal oficial da MMR Consultoria</strong>.</p>
        </div>
    """, unsafe_allow_html=True)
    st.stop()

## 🔍 Descobrir IP externo do usuário
#@st.cache_data(ttl=600)
#def get_ip():
#    try:
#        return requests.get("https://api.ipify.org").text
#    except:
#        return "0.0.0.0"

# Lista de IPs autorizados
#IPS_AUTORIZADOS = ["35.203.187.165", "201.10.22.33"]  # atualize conforme necessário

# 👉 Captura o IP corretamente
#ip_usuario = get_ip()

# ❌ Bloqueia se IP não estiver na lista
#if ip_usuario not in IPS_AUTORIZADOS:
#    st.markdown("""
#        <style>
#        #MainMenu, header, footer, .stSidebar { display: none; }
#        </style>
#        ## 🔐 IP não autorizado
#        Seu IP detectado: """ + ip_usuario + """
#
#        Copie este IP e envie para a equipe da MMR Consultoria para liberar o acesso.
#    """, unsafe_allow_html=True)
#    st.stop()

# ✅ Lista de usuários
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
    {"codigo": "1825", "email": "micaelly@grupofit.com.br", "senha": "bird*$#@&Mklo*22"},
    {"codigo": "3377", "email": "maricelisrossi@gmail.com", "senha": "1825"}
]

# ========================
# 🔐 Autenticação Google Sheets
# ========================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT_ACESSOS"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)

from datetime import datetime
import pytz

def registrar_acesso(nome_usuario):
    try:
        fuso_brasilia = pytz.timezone("America/Sao_Paulo")
        agora = datetime.now(fuso_brasilia)
        data = agora.strftime("%d/%m/%Y")
        hora = agora.strftime("%H:%M:%S")

        planilha = gc.open_by_key("1SZ5R6hcBE6o_qWs0_wx6IGKfIGltxpb9RWiGyF4L5uE")
        aba = planilha.sheet1
        nova_linha = [nome_usuario, data, hora]
        aba.append_row(nova_linha)
    except Exception as e:
        st.error(f"Erro ao registrar acesso: {e}")

# ✅ Redireciona se já estiver logado
if st.session_state.get("acesso_liberado"):
    st.switch_page("Home.py")

# ✅ Exibe o IP do usuário discretamente
#st.markdown(f"<p style='font-size:12px; color:#aaa;'>🛠️ Seu IP: <code>{ip_usuario}</code></p>", unsafe_allow_html=True)

# 🧾 Tela de login
st.title("🔐 Acesso Restrito")
st.markdown("Informe o código da empresa, e-mail e senha.")

codigo = st.text_input("Código da Empresa:")
email = st.text_input("E-mail:")
senha = st.text_input("Senha:", type="password")

# ✅ Botão de login
if st.button("Entrar"):
    usuario_encontrado = next(
        (u for u in USUARIOS if u["codigo"] == codigo and u["email"] == email and u["senha"] == senha),
        None
    )

    if usuario_encontrado:
        st.session_state["acesso_liberado"] = True
        st.session_state["empresa"] = codigo
        st.session_state["usuario_logado"] = email
        registrar_acesso(email)
        st.switch_page("Home.py")

    else:
        st.error("❌ Código, e-mail ou senha incorretos.")
