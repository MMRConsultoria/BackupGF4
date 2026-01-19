import streamlit as st
import pandas as pd
import json
import time
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# Drive API
try:
    from googleapiclient.discovery import build
except Exception:
    build = None

# ---------------- CONFIG ----------------
# ID atualizado com o hífen final conforme a URL enviada
MAIN_FOLDER_ID = "1LrbcStUAcvZV_dOYKBt-vgBHb9e1d6X-"  
ID_PLANILHA_ORIGEM = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
ABA_ORIGEM = "Fat Sistema Externo"
MAPA_ABAS = {"Faturamento": "Importado Fat", "Meio Pagamento": "Meio Pagamento", "Desconto": "Desconto"}

st.set_page_config(page_title="Atualizador — selecionar subpastas", layout="wide")
st.title("🚀 Atualizador de Planilhas por Subpastas")

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
    st.error(f"Erro de autenticação. Verifique st.secrets['GOOGLE_SERVICE_ACCOUNT']: {e}")
    st.stop()

if not drive_service:
    st.error("Drive API não inicializada. Verifique dependências.")
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
def list_all_descendant_folders(_drive, root_id):
    all_folders = []
    queue = [root_id]
    seen = set()
    while queue:
        pid = queue.pop(0)
        if pid in seen: continue
        seen.add(pid)
        try:
            resp_meta = _drive.files().get(fileId=pid, fields="id, name").execute()
            all_folders.append({"id": resp_meta["id"], "name": resp_meta.get("name", "")})
        except: pass
        try:
            children = list_subfolders(_drive, pid)
            for c in children:
                if c["id"] not in seen: queue.append(c["id"])
        except: pass
    return all_folders

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

# ---------------- UI: PASSO 0 - DATA ----------------
st.markdown("### 📅 1) Período de Atualização")
col_data1, col_data2 = st.columns(2)
with col_data1:
    data_corte = st.date_input("Data de início (Dados a partir de)", value=datetime.now() - timedelta(days=30))
with col_data2:
    st.info(f"Os dados serão filtrados na planilha origem para datas >= {data_corte.strftime('%d/%m/%Y')}")

st.markdown("---")

# ---------------- UI: PASSO 1 - PASTAS ----------------
st.markdown("### 📂 2) Seleção de Pastas")
col_rec, col_info = st.columns([0.4, 0.6])
with col_rec:
    recursive = st.checkbox("Buscar recursivamente (incluir sub-subpastas)", value=False)

try:
    subfolders = list_subfolders(drive_service, MAIN_FOLDER_ID)
    sub_names = [f"{s['name']} ({s['id']})" for s in subfolders]
    sel = st.multiselect("Selecione as pastas que deseja processar:", options=sub_names, default=sub_names)
except Exception as e:
    st.error(f"Erro ao listar pastas: {e}. Verifique se compartilhou a pasta com o e-mail da Service Account.")
    st.stop()

selected_folder_ids = []
for s in sel:
    if "(" in s:
        fid = s.split("(")[-1].strip(")")
        selected_folder_ids.append(fid)

if not selected_folder_ids:
    st.warning("Selecione ao menos uma pasta.")
    st.stop()

# Expandir pastas se recursivo
all_folder_ids_to_scan = set()
if recursive:
    for fid in selected_folder_ids:
        try:
            descendants = list_all_descendant_folders(drive_service, fid)
            for d in descendants: all_folder_ids_to_scan.add(d["id"])
        except: all_folder_ids_to_scan.add(fid)
else:
    all_folder_ids_to_scan.update(selected_folder_ids)

# ---------------- BUSCA PLANILHAS ----------------
with st.spinner("Buscando planilhas..."):
    planilhas = list_spreadsheets_in_folders(drive_service, list(all_folder_ids_to_scan))

if not planilhas:
    st.warning("Nenhuma planilha encontrada nas pastas selecionadas.")
    st.stop()

# DataFrame para o editor
df = pd.DataFrame(planilhas)
df = df.rename(columns={"name": "Planilha", "id": "ID_Planilha", "parent_folder_id": "Folder_ID"})
df["Desconto"] = True
df["Meio Pagamento"] = True
df["Faturamento"] = True
df = df[["Planilha", "Desconto", "Meio Pagamento", "Faturamento", "ID_Planilha"]].sort_values("Planilha")

# ---------------- UI: PASSO 2 - TABELA ----------------
st.markdown("### 📝 3) Ajuste as Operações por Planilha")
with st.form("selection_form"):
    edited = st.data_editor(
        df,
        use_container_width=True,
        column_config={
            "Planilha": st.column_config.TextColumn("Planilha", disabled=True, width="large"),
            "Desconto": st.column_config.CheckboxColumn("Desconto"),
            "Meio Pagamento": st.column_config.CheckboxColumn("Meio Pagamento"),
            "Faturamento": st.column_config.CheckboxColumn("Faturamento"),
            "ID_Planilha": None # Esconde o ID
        },
        hide_index=True
    )

    st.markdown("---")
    c1, c2 = st.columns(2)
    with c1:
        dry_run = st.checkbox("Modo Simulação (Dry-run)", value=True)
    with c2:
        do_backup = st.checkbox("Criar backup das abas", value=True)

    submit = st.form_submit_button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True)

# ---------------- EXECUÇÃO ----------------
if submit:
    tarefas = []
    for _, row in edited.iterrows():
        for op in ["Desconto", "Meio Pagamento", "Faturamento"]:
            if row[op]:
                tarefas.append({"planilha": row["Planilha"], "id": row["ID_Planilha"], "operacao": op, "aba": MAPA_ABAS[op]})

    if not tarefas:
        st.error("Nenhuma operação selecionada.")
    else:
        st.write(f"Processando **{len(tarefas)}** tarefas...")
        progresso = st.progress(0)
        for i, t in enumerate(tarefas):
            st.info(f"Atualizando: {t['planilha']} -> {t['operacao']}")
            # Lógica de gravação aqui...
            time.sleep(0.1)
            progresso.progress((i + 1) / len(tarefas))
        st.success("Concluído!")
