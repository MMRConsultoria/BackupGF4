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
# Observação: ID com hífen conforme sua URL
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
    st.error("Drive API não inicializada. Verifique dependências e permissões.")
    st.stop()

# ---------------- HELPERS DRIVE (usar _drive para evitar hashing do objeto) ----------------
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
def list_all_descendant_folders(_drive, root_id):
    all_folders = []
    queue = [root_id]
    seen = set()
    while queue:
        pid = queue.pop(0)
        if pid in seen:
            continue
        seen.add(pid)
        try:
            resp_meta = _drive.files().get(fileId=pid, fields="id, name").execute()
            all_folders.append({"id": resp_meta["id"], "name": resp_meta.get("name", "")})
        except Exception:
            pass
        try:
            children = list_subfolders(_drive, pid)
        except Exception:
            children = []
        for c in children:
            if c["id"] not in seen:
                queue.append(c["id"])
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
            if not page_token:
                break
    # remover duplicatas por id, mantendo primeiro encontro
    seen = set()
    unique = []
    for s in sheets:
        if s["id"] not in seen:
            seen.add(s["id"])
            unique.append(s)
    return unique

# ---------------- UI: PASSO 0 - PERÍODO (DE / ATÉ) ----------------
st.markdown("### 📅 1) Período de atualização (formato exigido: dd/mm/aaaa)")
col_start, col_end = st.columns(2)

# valores default: últimos 30 dias
default_end = date.today()
default_start = default_end - timedelta(days=30)

with col_start:
    data_de = st.date_input("De (dd/mm/aaaa)", value=default_start, help="Escolha a data inicial (formato dd/mm/aaaa)")
with col_end:
    data_ate = st.date_input("Até (dd/mm/aaaa)", value=default_end, help="Escolha a data final (formato dd/mm/aaaa)")

# formatação e validação
data_de_str = data_de.strftime("%d/%m/%Y")
data_ate_str = data_ate.strftime("%d/%m/%Y")

if data_ate < data_de:
    st.error("Data 'Até' deve ser igual ou posterior à data 'De'. Ajuste o período.")
    st.stop()
else:
    #st.info(f"Dados serão filtrados entre {data_de_str} e {data_ate_str} (inclusive).")

st.markdown("---")

# ---------------- UI: PASSO 1 - PASTAS ----------------
st.markdown("### 📂 2) Seleção de Pastas")
col_rec, col_info = st.columns([0.35, 0.65])
with col_rec:
    recursive = st.checkbox("Buscar recursivamente (incluir sub-subpastas)", value=False)
with col_info:
    st.write(f"Pasta principal: `{MAIN_FOLDER_ID}`")

# Listar subpastas imediatas para o usuário escolher
try:
    subfolders = list_subfolders(drive_service, MAIN_FOLDER_ID)
except Exception as e:
    st.error(f"Erro listando subpastas: {e}")
    st.stop()

if not subfolders:
    st.warning("Nenhuma subpasta encontrada dentro da pasta principal. Verifique se a service-account tem acesso ou se a pasta contém subpastas.")
    st.stop()

sub_names = [f"{s['name']} ({s['id']})" for s in subfolders]
selected = st.multiselect("Selecione as subpastas a incluir:", options=sub_names, default=sub_names)

selected_folder_ids = []
for s in selected:
    if "(" in s and s.strip().endswith(")"):
        fid = s.split("(")[-1].strip(")")
        selected_folder_ids.append(fid)

if not selected_folder_ids:
    st.warning("Selecione ao menos uma subpasta para prosseguir.")
    st.stop()

# Expandir se recursivo
all_folder_ids_to_scan = set()
if recursive:
    for fid in selected_folder_ids:
        try:
            descendants = list_all_descendant_folders(drive_service, fid)
            for d in descendants:
                all_folder_ids_to_scan.add(d["id"])
        except Exception:
            all_folder_ids_to_scan.add(fid)
else:
    all_folder_ids_to_scan.update(selected_folder_ids)

st.success(f"Serão escaneadas {len(all_folder_ids_to_scan)} pasta(s).")

# ---------------- BUSCAR PLANILHAS NAS PASTAS SELECIONADAS ----------------
with st.spinner("Buscando planilhas nas pastas selecionadas..."):
    try:
        planilhas = list_spreadsheets_in_folders(drive_service, list(all_folder_ids_to_scan))
    except Exception as e:
        st.error(f"Erro ao listar planilhas: {e}")
        st.stop()

if not planilhas:
    st.warning("Nenhuma planilha encontrada nas subpastas selecionadas.")
    st.stop()

# preparar DataFrame para edição
df = pd.DataFrame(planilhas)
df = df.rename(columns={"name": "Planilha", "id": "ID_Planilha", "parent_folder_id": "Folder_ID"})
df["Desconto"] = True
df["Meio Pagamento"] = True
df["Faturamento"] = True
df = df[["Planilha", "Folder_ID", "ID_Planilha", "Desconto", "Meio Pagamento", "Faturamento"]].sort_values("Planilha").reset_index(drop=True)

st.markdown("### 📝 3) Ajuste as operações por planilha")
if not hasattr(st, "data_editor"):
    st.error("Seu Streamlit não tem `st.data_editor`. Atualize o Streamlit: pip install --upgrade streamlit")
    st.stop()

# ---------------- Form com data_editor ----------------
with st.form("selection_form"):
    edited = st.data_editor(
        df,
        num_rows="fixed",
        use_container_width=True,
        column_config={
            "Planilha": st.column_config.TextColumn("Planilha", disabled=True, width="large"),
            "Folder_ID": st.column_config.TextColumn("Pasta (ID)", disabled=True),
            "ID_Planilha": st.column_config.TextColumn("ID Planilha", disabled=True),
            "Desconto": st.column_config.CheckboxColumn("Desconto", default=True),
            "Meio Pagamento": st.column_config.CheckboxColumn("Meio Pagamento", default=True),
            "Faturamento": st.column_config.CheckboxColumn("Faturamento", default=True),
        },
        hide_index=True
    )

    st.markdown("---")
    col1, col2 = st.columns(2)
    with col1:
        dry_run = st.checkbox("Dry-run (não grava)", value=True)
    with col2:
        do_backup = st.checkbox("Criar backup da aba destino (se existir)", value=True)

    submit = st.form_submit_button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True)

# ---------------- EXECUÇÃO (simulada por padrão) ----------------
if submit:
    # monta tarefas conforme seleção
    tarefas = []
    for _, row in edited.iterrows():
        if row["Desconto"]:
            tarefas.append({"planilha": row["Planilha"], "id": row["ID_Planilha"], "operacao": "Desconto", "aba": MAPA_ABAS["Desconto"]})
        if row["Meio Pagamento"]:
            tarefas.append({"planilha": row["Planilha"], "id": row["ID_Planilha"], "operacao": "Meio Pagamento", "aba": MAPA_ABAS["Meio Pagamento"]})
        if row["Faturamento"]:
            tarefas.append({"planilha": row["Planilha"], "id": row["ID_Planilha"], "operacao": "Faturamento", "aba": MAPA_ABAS["Faturamento"]})

    if not tarefas:
        st.warning("Nenhuma operação selecionada. Marque ao menos uma caixa antes de enviar.")
    else:
        st.write(f"Iniciando processamento de **{len(tarefas)}** tarefas (período: {data_de_str} → {data_ate_str})")
        progresso = st.progress(0)
        logs = []
        for i, t in enumerate(tarefas):
            st.info(f"{i+1}/{len(tarefas)} — {t['planilha']} -> {t['operacao']}")
            try:
                if not dry_run:
                    sh = gc.open_by_key(t["id"])
                    if do_backup:
                        try:
                            ws = sh.worksheet(t["aba"])
                            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                            sh.duplicate_sheet(ws.id, new_sheet_name=f"BACKUP_{t['aba']}_{timestamp}")
                            logs.append(f"{t['planilha']}/{t['aba']}: backup criado")
                        except Exception as e:
                            logs.append(f"{t['planilha']}/{t['aba']}: backup falhou ou aba não existe -> {e}")
                    # Aqui você adiciona a lógica real de leitura da origem (ID_PLANILHA_ORIGEM/ABA_ORIGEM),
                    # filtra com data >= data_de e <= data_ate e grava na aba destino.
                    # Por segurança, esse exemplo só simula a gravação.
                    time.sleep(0.2)
                    logs.append(f"{t['planilha']}/{t['operacao']}: gravaria dados para {data_de_str}→{data_ate_str} (simulado)")
                else:
                    logs.append(f"{t['planilha']}/{t['operacao']}: dry-run (não gravado)")
            except Exception as e:
                logs.append(f"{t['planilha']}/{t['operacao']}: ERRO -> {e}")
            progresso.progress((i + 1) / len(tarefas))

        st.success("Processamento finalizado.")
        st.write("Logs:")
        st.write("\n".join(logs))
