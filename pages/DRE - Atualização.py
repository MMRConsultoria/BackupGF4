import streamlit as st
import pandas as pd
import json
import time
from datetime import datetime, timedelta, date
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# Drive API
try:
    from googleapiclient.discovery import build
except Exception:
    build = None

# ---------------- CONFIG ----------------
MAIN_FOLDER_ID = "1LrbcStUAcvZV_dOYKBt-vgBHb9e1d6X-"
ID_PLANILHA_ORIGEM = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
ABA_ORIGEM = "Fat Sistema Externo"
MAPA_ABAS = {"Faturamento": "Importado Fat", "Meio Pagamento": "Meio Pagamento", "Desconto": "Desconto"}

st.set_page_config(page_title="Atualizador DRE", layout="wide")

# --- CSS PARA REDUZIR ESPAÇAMENTO ---
st.markdown("""
    <style>
    /* Reduz o espaço no topo da página */
    .block-container { padding-top: 1rem; padding-bottom: 0rem; }
    /* Reduz o espaço entre os widgets */
    div.stVerticalBlock > div { margin-bottom: -0.8rem; }
    /* Reduz o espaço dos títulos */
    h1 { margin-top: -1rem; margin-bottom: 0.5rem; }
    /* Esconde o rótulo vazio do multiselect para não ocupar espaço */
    label[data-testid="stWidgetLabel"] { min-height: 0px; margin-bottom: 0px; }
    </style>
    """, unsafe_allow_html=True)

st.title("Atualizador DRE")

# ---------------- AUTENTICAÇÃO ----------------
@st.cache_resource
def autenticar():
    scope = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
    creds_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    gc = gspread.authorize(creds)
    drive = None
    if build:
        try:
            drive = build("drive", "v3", credentials=creds)
        except Exception:
            drive = None
    return gc, drive

try:
    gc, drive_service = autenticar()
except Exception as e:
    st.error(f"Erro de autenticação: {e}")
    st.stop()

# ---------------- HELPERS DRIVE ----------------
@st.cache_data(ttl=300)
def list_subfolders(_drive, parent_id):
    folders = []
    page_token = None
    q = f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    while True:
        resp = _drive.files().list(q=q, fields="nextPageToken, files(id, name)", pageToken=page_token).execute()
        files = resp.get("files", [])
        for f in files:
            folders.append({"id": f["id"], "name": f["name"]})
        page_token = resp.get("nextPageToken", None)
        if not page_token: break
    return folders

@st.cache_data(ttl=300)
def list_spreadsheets_in_folders(_drive, folder_ids):
    sheets = []
    for fid in folder_ids:
        page_token = None
        q = f"'{fid}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
        while True:
            resp = _drive.files().list(q=q, fields="nextPageToken, files(id, name)", pageToken=page_token).execute()
            files = resp.get("files", [])
            for f in files:
                sheets.append({"id": f["id"], "name": f["name"], "parent_folder_id": fid})
            page_token = resp.get("nextPageToken", None)
            if not page_token: break
    seen = set()
    unique = []
    for s in sheets:
        if s["id"] not in seen:
            seen.add(s["id"])
            unique.append(s)
    return unique

# ---------------- UI: PASSO 0 - PERÍODO ----------------
col_start, col_end = st.columns(2)
default_end = date.today()
default_start = default_end - timedelta(days=30)

with col_start:
    data_de = st.date_input("De (dd/mm/aaaa)", value=default_start)
with col_end:
    data_ate = st.date_input("Até (dd/mm/aaaa)", value=default_end)

data_de_str = data_de.strftime("%d/%m/%Y")
data_ate_str = data_ate.strftime("%d/%m/%Y")

if data_ate < data_de:
    st.error("Data 'Até' deve ser posterior à 'De'.")
    st.stop()

# ---------------- LISTAR E SELECIONAR SUBPASTAS ----------------
try:
    subfolders = list_subfolders(drive_service, MAIN_FOLDER_ID)
    sub_names = [f"{s['name']} ({s['id']})" for s in subfolders]
    # Rótulo removido ("") para eliminar os dizeres
    selected = st.multiselect("", options=sub_names, default=sub_names)
except Exception as e:
    st.error(f"Erro: {e}")
    st.stop()

selected_folder_ids = [s.split("(")[-1].strip(")") for s in selected if "(" in s]

if not selected_folder_ids:
    st.stop()

# ---------------- BUSCAR PLANILHAS ----------------
with st.spinner("Buscando planilhas..."):
    planilhas = list_spreadsheets_in_folders(drive_service, list(selected_folder_ids))

if not planilhas:
    st.warning("Nenhuma planilha encontrada.")
    st.stop()

df = pd.DataFrame(planilhas)
df = df.rename(columns={"name": "Planilha", "id": "ID_Planilha", "parent_folder_id": "Folder_ID"})
df["Desconto"] = True
df["Meio Pagamento"] = True
df["Faturamento"] = True
df = df[["Planilha", "Folder_ID", "ID_Planilha", "Desconto", "Meio Pagamento", "Faturamento"]].sort_values("Planilha").reset_index(drop=True)

# ---------------- TABELA E FORM ----------------
with st.form("selection_form"):
    edited = st.data_editor(
        df,
        num_rows="fixed",
        use_container_width=True,
        column_config={
            "Planilha": st.column_config.TextColumn("Planilha", disabled=True, width="large"),
            "Folder_ID": None, # Escondido para reduzir espaço
            "ID_Planilha": None, # Escondido para reduzir espaço
            "Desconto": st.column_config.CheckboxColumn("Desconto", default=True),
            "Meio Pagamento": st.column_config.CheckboxColumn("Meio Pagamento", default=True),
            "Faturamento": st.column_config.CheckboxColumn("Faturamento", default=True),
        },
        hide_index=True
    )
    submit = st.form_submit_button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True)

# ---------------- EXECUÇÃO ----------------
if submit:
    # Configurações fixas
    DRY_RUN = True 
    DO_BACKUP = True

    tarefas = []
    for _, row in edited.iterrows():
        if row["Desconto"]: tarefas.append({"planilha": row["Planilha"], "id": row["ID_Planilha"], "operacao": "Desconto", "aba": MAPA_ABAS["Desconto"]})
        if row["Meio Pagamento"]: tarefas.append({"planilha": row["Planilha"], "id": row["ID_Planilha"], "operacao": "Meio Pagamento", "aba": MAPA_ABAS["Meio Pagamento"]})
        if row["Faturamento"]: tarefas.append({"planilha": row["Planilha"], "id": row["ID_Planilha"], "operacao": "Faturamento", "aba": MAPA_ABAS["Faturamento"]})

    if not tarefas:
        st.warning("Nenhuma operação selecionada.")
    else:
        st.write(f"Processando **{len(tarefas)}** tarefas...")
        progresso = st.progress(0)
        for i, t in enumerate(tarefas):
            try:
                if not DRY_RUN:
                    sh = gc.open_by_key(t["id"])
                    # Lógica de backup e gravação aqui
                time.sleep(0.1)
            except Exception as e:
                st.error(f"Erro em {t['planilha']}: {e}")
            progresso.progress((i + 1) / len(tarefas))
        st.success("Concluído!")
