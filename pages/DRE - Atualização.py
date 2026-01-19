import streamlit as st
import pandas as pd
import json
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# tenta importar Drive API (opcional)
try:
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except Exception:
    build = None
    HttpError = Exception

st.set_page_config(page_title="Atualização e Auditoria - Meio de Pagamento", layout="wide")
st.title("Atualização e Auditoria — Faturamento x Meio de Pagamento")

# -----------------------
# Configurações iniciais
# -----------------------
DEFAULT_FOLDER_IDS = [
    # Cole aqui as IDs das pastas que quer listar
    # Exemplo: "1ptFvtxYjISfB19S7bU9olMLmAxDTBkOh",
]
DEFAULT_ORIGIN_SPREADSHEET = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
DEFAULT_ORIGIN_SHEET = "Fat Sistema Externo"
DEFAULT_DATA_MINIMA = datetime.now() - timedelta(days=365)  # últimos 365 dias

# -----------------------
# Autenticação gspread (+drive opcional)
# -----------------------
@st.cache_resource
def autenticar_gspread():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gc = gspread.authorize(credentials)
    drive_service = None
    if build:
        try:
            drive_service = build("drive", "v3", credentials=credentials)
        except Exception:
            drive_service = None
    return gc, drive_service

gc, drive_service = autenticar_gspread()

# -----------------------
# Funções utilitárias
# -----------------------
def listar_arquivos_pasta(drive_service, pasta_id):
    arquivos = []
    page_token = None
    query = f"'{pasta_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    while True:
        try:
            resp = drive_service.files().list(q=query, spaces="drive", fields="nextPageToken, files(id, name)", pageToken=page_token).execute()
            arquivos.extend(resp.get("files", []))
            page_token = resp.get("nextPageToken", None)
            if not page_token:
                break
        except HttpError as e:
            st.error(f"Erro listando pasta {pasta_id}: {e}")
            break
        except Exception as e:
            st.error(f"Erro listando pasta {pasta_id}: {e}")
            break
    return arquivos

def carregar_origem(gc, origin_spreadsheet_id, origin_sheet_name):
    sh = gc.open_by_key(origin_spreadsheet_id)
    ws = sh.worksheet(origin_sheet_name)
    vals = ws.get_all_values()
    if not vals or len(vals) < 2:
        raise RuntimeError(f"Aba origem '{origin_sheet_name}' vazia ou sem dados.")
    df = pd.DataFrame(vals[1:], columns=vals[0])
    df.columns = [c.strip() for c in df.columns]
    if "Grupo" not in df.columns or "Data" not in df.columns:
        raise RuntimeError("Aba origem precisa conter as colunas 'Grupo' e 'Data'.")
    df["Grupo"] = df["Grupo"].astype(str).str.strip().str.upper()
    df["Data_dt"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
    return df

# -----------------------
# Sidebar: parâmetros
# -----------------------
st.sidebar.header("Parâmetros")
origin_id = st.sidebar.text_input("ID planilha origem", value=DEFAULT_ORIGIN_SPREADSHEET)
origin_sheet = st.sidebar.text_input("Aba origem (na planilha origem)", value=DEFAULT_ORIGIN_SHEET)
data_minima = st.sidebar.date_input("Data mínima (incluir)", value=DEFAULT_DATA_MINIMA.date())

folder_ids_text = st.sidebar.text_area("IDs das pastas (uma por linha) — opcional", value="\n".join(DEFAULT_FOLDER_IDS), height=80)
folder_ids = [s.strip() for s in folder_ids_text.splitlines() if s.strip()]

# -----------------------
# Listar arquivos nas pastas e mostrar
# -----------------------
if drive_service and folder_ids:
    st.header("Arquivos encontrados nas pastas configuradas")
    for fid in folder_ids:
        st.subheader(f"Pasta ID: {fid}")
        try:
            arquivos = listar_arquivos_pasta(drive_service, fid)
            if arquivos:
                for a in arquivos:
                    st.write(f"- {a['name']} (ID: {a['id']})")
            else:
                st.write("Nenhum arquivo encontrado ou sem permissão.")
        except Exception as e:
            st.error(f"Erro ao listar pasta {fid}: {e}")
else:
    st.info("Drive API não disponível ou nenhuma pasta configurada para listar.")

# -----------------------
# Teste rápido de acesso a planilha compartilhada
# -----------------------
st.header("Teste de acesso a planilha compartilhada")
creds = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
st.write("Service account client_email:", creds.get("client_email"))

test_id = st.text_input("Cole aqui o ID da planilha para testar acesso", "")

if st.button("Testar acesso"):
    if not test_id.strip():
        st.warning("Por favor, insira um ID válido.")
    else:
        try:
            sh = gc.open_by_key(test_id.strip())
            st.success(f"Acesso OK: título = {sh.title}")
        except Exception as e:
            st.error(f"Falha ao abrir planilha: {e}")
            st.info("Se falhar, compartilhe a planilha com o e-mail acima (client_email) e tente novamente.")
