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
PASTA_PRINCIPAL_ID = "0B1owaTi3RZnFfm4tTnhfZ2l0VHo4bWNMdHhKS3ZlZzR1ZjRSWWJSSUFxQTJtUExBVlVTUW8"
ID_PLANILHA_ORIGEM = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
ABA_ORIGEM = "Fat Sistema Externo"
MAPA_ABAS = {"Faturamento": "Importado Fat", "Meio Pagamento": "Meio Pagamento", "Desconto": "Desconto"}

st.set_page_config(page_title="Atualizador DRE", layout="wide")

# --- CSS PARA COMPACTAÇÃO MÁXIMA ---
st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; padding-bottom: 0rem; }
    div.stVerticalBlock > div { margin-bottom: -1.0rem; }
    div[data-testid="stVerticalBlock"] > div { padding-top: 0rem; padding-bottom: 0rem; }
    h1 { margin-top: -1rem; margin-bottom: 0.5rem; font-size: 1.8rem; }
    [data-testid="stTable"] td, [data-testid="stTable"] th { padding: 2px 6px !important; }
    .global-selection-container { padding-top: 0px; padding-bottom: 0px; margin-top: 10px; margin-bottom: 5px; }
    hr { margin: 0.5rem 0px !important; }
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
def list_child_folders(_drive, parent_id, filtro_texto=None):
    folders = []
    if not _drive: return folders
    page_token = None
    q = f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    while True:
        resp = _drive.files().list(q=q, fields="nextPageToken, files(id, name)", pageToken=page_token).execute()
        files = resp.get("files", [])
        for f in files:
            name = f.get("name", "")
            if filtro_texto is None or filtro_texto.lower() in name.lower():
                folders.append({"id": f["id"], "name": name})
        page_token = resp.get("nextPageToken", None)
        if not page_token: break
    return folders

@st.cache_data(ttl=60)
def list_spreadsheets_in_folders(_drive, folder_ids):
    sheets = []
    if not _drive: return sheets
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

# ---------------- FILTROS DE DATA ----------------
col_d1, col_d2 = st.columns(2)
with col_d1:
    data_de = st.date_input("De", value=date.today() - timedelta(days=30))
with col_d2:
    data_ate = st.date_input("Até", value=date.today())

# ---------------- SELEÇÃO DE PASTAS ----------------
try:
    pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, filtro_texto="fechamento")
    map_pasta_nome_id = {p["name"]: p["id"] for p in pastas_fech}
    
    pasta_selecionada_nome = st.selectbox("Pasta principal (fechamento):", options=list(map_pasta_nome_id.keys()), index=0)
    pasta_selecionada_id = map_pasta_nome_id.get(pasta_selecionada_nome)

    subpastas = list_child_folders(drive_service, pasta_selecionada_id)
    map_sub_nome_id = {s["name"]: s["id"] for s in subpastas}
    
    selecionadas_nomes = st.multiselect("Subpastas:", options=list(map_sub_nome_id.keys()), default=list(map_sub_nome_id.keys())[:1])
    selecionadas_ids = [map_sub_nome_id[n] for n in selecionadas_nomes]
except Exception as e:
    st.error(f"Erro ao carregar pastas: {e}")
    st.stop()

# ---------------- TABELAS DE PLANILHAS ----------------
if selecionadas_ids:
    with st.spinner("Buscando planilhas..."):
        planilhas = list_spreadsheets_in_folders(drive_service, selecionadas_ids)
        
        if planilhas:
            df_base = pd.DataFrame(planilhas).sort_values("name").reset_index(drop=True)
            df_base = df_base.rename(columns={"name": "Planilha", "id": "ID_Planilha", "parent_folder_id": "Folder_ID"})
            
            # Seleção global compacta
            st.markdown('<div class="global-selection-container">', unsafe_allow_html=True)
            c1, c2, c3, _ = st.columns([1.2, 1.2, 1.2, 5])
            with c1: sel_desc = st.checkbox("Desconto", value=True, key="global_desc")
            with c2: sel_mp = st.checkbox("Meio Pagto", value=True, key="global_mp")
            with c3: sel_fat = st.checkbox("Faturamento", value=True, key="global_fat")
            st.markdown('</div>', unsafe_allow_html=True)
            
            df_base["Desconto"] = sel_desc
            df_base["Meio Pagamento"] = sel_mp
            df_base["Faturamento"] = sel_fat
            
            meio = len(df_base) // 2 + (len(df_base) % 2)
            df_esq = df_base.iloc[:meio].copy()
            df_dir = df_base.iloc[meio:].copy()

            config_col = {
                "Planilha": st.column_config.TextColumn("Planilha", disabled=True),
                "Folder_ID": None, "ID_Planilha": None,
                "Desconto": st.column_config.CheckboxColumn("Desc."),
                "Meio Pagamento": st.column_config.CheckboxColumn("M.Pag"),
                "Faturamento": st.column_config.CheckboxColumn("Fat."),
            }

            col_t1, col_t2 = st.columns(2)
            with col_t1:
                edit_esq = st.data_editor(df_esq, key="tab_esq", use_container_width=True, column_config=config_col, hide_index=True)
            with col_t2:
                edit_dir = st.data_editor(df_dir, key="tab_dir", use_container_width=True, column_config=config_col, hide_index=True)

            if st.button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True):
                st.info("Processando...")
        else:
            st.warning("Nenhuma planilha encontrada.")
