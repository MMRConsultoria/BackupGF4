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

# --- CSS PARA COMPACTAR E AJUSTAR ESPAÇAMENTOS ---
st.markdown(
    """
    <style>
    .block-container { padding-top: 1.5rem; padding-bottom: 0rem; }
    div.stVerticalBlock > div { margin-bottom: -0.2rem; }
    h1 { margin-top: -1rem; margin-bottom: 1rem; font-size: 1.8rem; }
    [data-testid="stTable"] td, [data-testid="stTable"] th { padding: 2px 6px !important; }
    .global-selection-container { padding-top: 10px; padding-bottom: 10px; margin-bottom: 10px; }
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
    """
    Lista pastas-filhas diretas de parent_id.
    Se filtro_texto fornecido, filtra (case-insensitive) apenas nomes que contenham o texto.
    Retorna lista de dicts: {'id': <id>, 'name': <name>}
    """
    folders = []
    if not _drive:
        return folders
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
        if not page_token:
            break
    return folders

@st.cache_data(ttl=60)
def list_spreadsheets_in_folders(_drive, folder_ids):
    """
    Lista planilhas (sheets) que estão diretamente dentro das folder_ids fornecidas.
    Retorna lista de dicts: {'id': <id>, 'name': <name>, 'parent_folder_id': <folder_id>}
    """
    sheets = []
    if not _drive:
        return sheets
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
    # remover duplicados por id (caso apareça)
    seen = set()
    unique = []
    for s in sheets:
        if s["id"] not in seen:
            seen.add(s["id"])
            unique.append(s)
    return unique

# ---------------- FILTROS INICIAIS ----------------
col_d1, col_d2 = st.columns(2)
with col_d1:
    data_de = st.date_input("De", value=date.today() - timedelta(days=30))
with col_d2:
    data_ate = st.date_input("Até", value=date.today())

st.markdown("---")

# ---------------- PASSO 1: Lista de pastas "fechamento" dentro da pasta principal ----------------
st.subheader("1) Escolha a pasta principal (contendo 'fechamento')")
try:
    pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, filtro_texto="fechamento")
except Exception as e:
    st.error(f"Erro ao listar pastas na pasta principal: {e}")
    pastas_fech = []

if not pastas_fech:
    st.info("Nenhuma pasta com 'fechamento' encontrada na pasta principal.")
    st.stop()

# Mostrar somente nomes (sem ID) no selectbox. Usamos mapping interno name->id.
nomes_pastas_fech = [p["name"] for p in pastas_fech]
map_pasta_nome_id = {p["name"]: p["id"] for p in pastas_fech}

# Permitir seleção única (use selectbox). Se preferir múltipla, troque por multiselect.
pasta_selecionada_nome = st.selectbox("Selecione uma pasta (fechamento):", options=nomes_pastas_fech, index=0)
pasta_selecionada_id = map_pasta_nome_id.get(pasta_selecionada_nome)

# ---------------- PASSO 2: Listar subpastas da pasta selecionada ----------------
st.markdown("---")
st.subheader("2) Escolha uma ou mais subpastas dentro da pasta selecionada")

try:
    subpastas = list_child_folders(drive_service, pasta_selecionada_id, filtro_texto=None)
except Exception as e:
    st.error(f"Erro ao listar subpastas da pasta selecionada: {e}")
    subpastas = []

if not subpastas:
    st.info("A pasta selecionada não contém subpastas.")
    st.stop()

# Mostrar apenas nomes no multiselect; mapear para ids internamente.
nomes_subpastas = [s["name"] for s in subpastas]
map_sub_nome_id = {s["name"]: s["id"] for s in subpastas}

selecionadas_nomes = st.multiselect("Selecione as subpastas:", options=nomes_subpastas, default=nomes_subpastas[:1])
selecionadas_ids = [map_sub_nome_id[n] for n in selecionadas_nomes]

if not selecionadas_ids:
    st.info("Selecione ao menos uma subpasta para continuar.")
    st.stop()

# ---------------- PASSO 3: Listar planilhas dentro das subpastas selecionadas ----------------
st.markdown("---")
st.subheader("3) Planilhas encontradas nas subpastas selecionadas")

with st.spinner("Buscando planilhas nas subpastas selecionadas..."):
    try:
        planilhas = list_spreadsheets_in_folders(drive_service, selecionadas_ids)
    except Exception as e:
        st.error(f"Erro ao listar planilhas: {e}")
        planilhas = []

if not planilhas:
    st.info("Nenhuma planilha encontrada nas subpastas selecionadas.")
    st.stop()

# Montar DataFrame base (mantemos 'id' internamente, mas NÃO mostramos para o usuário)
df_base = pd.DataFrame(planilhas).sort_values("name").reset_index(drop=True)
df_base = df_base.rename(columns={"name": "Planilha", "id": "ID_Planilha", "parent_folder_id": "Folder_ID"})
# Colunas de seleção padrão
df_base["Desconto"] = True
df_base["Meio Pagamento"] = True
df_base["Faturamento"] = True

# ---------------- Seleção global (marcar/desmarcar todos) ----------------
st.markdown('<div class="global-selection-container">', unsafe_allow_html=True)
st.write("**Marcar/Desmarcar todos:**")
c1, c2, c3, _ = st.columns([1.2, 1.2, 1.2, 5])
with c1:
    sel_desc = st.checkbox("Desconto", value=True, key="global_desc")
with c2:
    sel_mp = st.checkbox("Meio Pagto", value=True, key="global_mp")
with c3:
    sel_fat = st.checkbox("Faturamento", value=True, key="global_fat")
st.markdown('</div>', unsafe_allow_html=True)

# Aplica seleção global ao DataFrame antes de renderizar
df_base["Desconto"] = sel_desc
df_base["Meio Pagamento"] = sel_mp
df_base["Faturamento"] = sel_fat

# ---------------- Exibir em duas colunas lado a lado (sem IDs visíveis) ----------------
meio = len(df_base) // 2 + (len(df_base) % 2)
df_esq = df_base.iloc[:meio].copy()
df_dir = df_base.iloc[meio:].copy()

config_col = {
    "Planilha": st.column_config.TextColumn("Planilha", disabled=True, width="large"),
    "Folder_ID": None,       # ocultar Pasta(ID)
    "ID_Planilha": None,     # ocultar ID Planilha
    "Desconto": st.column_config.CheckboxColumn("Desc."),
    "Meio Pagamento": st.column_config.CheckboxColumn("M.Pag"),
    "Faturamento": st.column_config.CheckboxColumn("Fat."),
}

col_t1, col_t2 = st.columns(2)
with col_t1:
    edit_esq = st.data_editor(df_esq, key="tab_esq", use_container_width=True, column_config=config_col, hide_index=True)
with col_t2:
    edit_dir = st.data_editor(df_dir, key="tab_dir", use_container_width=True, column_config=config_col, hide_index=True)

st.markdown("---")

# ---------------- Botão de execução (une as tabelas editadas) ----------------
if st.button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True):
    # Concatenar resultados (manter a ordem original)
    df_final = pd.concat([edit_esq, edit_dir], ignore_index=True)
    tarefas = []
    for _, row in df_final.iterrows():
        if row.get("Desconto"):
            tarefas.append({"planilha": row["Planilha"], "id": row["ID_Planilha"], "op": "Desconto"})
        if row.get("Meio Pagamento"):
            tarefas.append({"planilha": row["Planilha"], "id": row["ID_Planilha"], "op": "Meio Pagamento"})
        if row.get("Faturamento"):
            tarefas.append({"planilha": row["Planilha"], "id": row["ID_Planilha"], "op": "Faturamento"})

    if not tarefas:
        st.warning("Nenhuma operação selecionada.")
    else:
        st.success(f"Iniciando processamento de {len(tarefas)} tarefas...")
        progresso = st.progress(0)
        logs = []
        for i, t in enumerate(tarefas):
            try:
                # Aqui você implementa a lógica real: abrir a planilha t['id'], gravar abas, etc.
                # Exemplo (simulado): time.sleep(0.1)
                time.sleep(0.05)
                logs.append(f"{t['planilha']} -> {t['op']}: simulado")
            except Exception as e:
                logs.append(f"{t['planilha']} -> {t['op']}: ERRO -> {e}")
            progresso.progress((i + 1) / len(tarefas))
        st.write("Logs:")
        st.write("\n".join(logs))
