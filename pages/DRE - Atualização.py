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

# --- CSS PARA COMPACTAR LAYOUT ---
st.markdown(
    """
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 0rem; }
    div.stVerticalBlock > div { margin-bottom: -0.1rem; }
    h1 { margin-top: 1rem; margin-bottom: 1rem; font-size: 1.0rem; }
    [data-testid="stTable"] td, [data-testid="stTable"] th { padding: 2px 5px !important; }
    /* Estilo para os botões de seleção global */
    .stCheckbox { margin-bottom: -15px; }
    </style>
    """,
    unsafe_allow_html=True,
)

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
        if not page_token:
            break
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
            if not page_token:
                break
    seen = set()
    unique = []
    for s in sheets:
        if s["id"] not in seen:
            seen.add(s["id"])
            unique.append(s)
    return unique

# ---------------- FILTROS (TOPO) ----------------
col_d1, col_d2 = st.columns(2)
with col_d1:
    data_de = st.date_input("De", value=date.today() - timedelta(days=30))
with col_d2:
    data_ate = st.date_input("Até", value=date.today())

try:
    subfolders = list_subfolders(drive_service, MAIN_FOLDER_ID)
    sub_names = [f"{s['name']} ({s['id']})" for s in subfolders]
    selected = st.multiselect("Selecione as subpastas a incluir:", options=sub_names, default=sub_names)
    selected_folder_ids = [s.split("(")[-1].strip(")") for s in selected if "(" in s]
except:
    selected_folder_ids = []

st.markdown("---")

# ---------------- LOGICA DE SELEÇÃO GLOBAL ----------------
if selected_folder_ids:
    with st.spinner("Buscando planilhas..."):
        planilhas = list_spreadsheets_in_folders(drive_service, list(selected_folder_ids))
        
        if planilhas:
            # Criar DataFrame base
            df_base = pd.DataFrame(planilhas).sort_values("name").reset_index(drop=True)
            
            # Botões de Seleção Global
            st.write("**Marcar/Desmarcar todos:**")
            c1, c2, c3, _ = st.columns([1, 1, 1, 5])
            with c1: sel_desc = st.checkbox("Desconto", value=True)
            with c2: sel_mp = st.checkbox("Meio Pagto", value=True)
            with c3: sel_fat = st.checkbox("Faturamento", value=True)
            
            # Aplicar seleção global ao DataFrame
            df_base["Desconto"] = sel_desc
            df_base["Meio Pagamento"] = sel_mp
            df_base["Faturamento"] = sel_fat
            
            # Divide o DataFrame em dois para as colunas
            meio = len(df_base) // 2 + (len(df_base) % 2)
            df_esq = df_base.iloc[:meio]
            df_dir = df_base.iloc[meio:]

            # Configuração das colunas da tabela
            config_col = {
                "name": st.column_config.TextColumn("Planilha", disabled=True),
                "id": None,
                "parent_folder_id": None,
                "Desconto": st.column_config.CheckboxColumn("Desc."),
                "Meio Pagamento": st.column_config.CheckboxColumn("M.Pag"),
                "Faturamento": st.column_config.CheckboxColumn("Fat."),
            }

            # Renderiza as duas tabelas lado a lado
            col_t1, col_t2 = st.columns(2)
            
            with col_t1:
                edit_esq = st.data_editor(df_esq, key="tab_esq", use_container_width=True, column_config=config_col, hide_index=True)
            
            with col_t2:
                edit_dir = st.data_editor(df_dir, key="tab_dir", use_container_width=True, column_config=config_col, hide_index=True)

            st.markdown("---")
            
            if st.button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True):
                df_final = pd.concat([edit_esq, edit_dir])
                tarefas = []
                for _, row in df_final.iterrows():
                    if row["Desconto"]: tarefas.append({"planilha": row["name"], "id": row["id"], "op": "Desconto"})
                    if row["Meio Pagamento"]: tarefas.append({"planilha": row["name"], "id": row["id"], "op": "Meio Pagamento"})
                    if row["Faturamento"]: tarefas.append({"planilha": row["name"], "id": row["id"], "op": "Faturamento"})
                
                if tarefas:
                    st.success(f"Processando {len(tarefas)} tarefas...")
                    # Lógica de execução aqui...
                else:
                    st.warning("Nenhuma operação selecionada.")
        else:
            st.warning("Nenhuma planilha encontrada.")
