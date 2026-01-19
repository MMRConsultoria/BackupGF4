import streamlit as st
import pandas as pd
import json
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import time

# tenta importar Drive API (opcional)
try:
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except Exception:
    build = None
    HttpError = Exception

st.set_page_config(page_title="Atualização e Auditoria - Meio de Pagamento", layout="wide")
st.title("Atualização e Auditoria — Faturamento x Meio de Pagamento")

# Configurações iniciais
DEFAULT_FOLDER_IDS = [
    "1ptFvtxYjISfB19S7bU9olMLmAxDTBkOh",
    "1F2Py4eeoqxqrHptgoeUODNXDCUddoU1u",
]
DEFAULT_ORIGIN_SPREADSHEET = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
DEFAULT_ORIGIN_SHEET = "Fat Sistema Externo"
DEFAULT_DATA_MINIMA = datetime.now() - timedelta(days=365)

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

# Sidebar parâmetros
st.sidebar.header("Parâmetros")
origin_id = st.sidebar.text_input("ID planilha origem", value=DEFAULT_ORIGIN_SPREADSHEET)
origin_sheet = st.sidebar.text_input("Aba origem (na planilha origem)", value=DEFAULT_ORIGIN_SHEET)
data_minima = st.sidebar.date_input("Data mínima (incluir)", value=DEFAULT_DATA_MINIMA.date())

folder_ids_text = st.sidebar.text_area("IDs das pastas (uma por linha) — opcional", value="\n".join(DEFAULT_FOLDER_IDS), height=80)
folder_ids = [s.strip() for s in folder_ids_text.splitlines() if s.strip()]

# Listar arquivos nas pastas e mostrar
if drive_service and folder_ids:
    st.header("Arquivos encontrados nas pastas configuradas")
    planilhas = []
    for fid in folder_ids:
        st.subheader(f"Pasta ID: {fid}")
        try:
            arquivos = listar_arquivos_pasta(drive_service, fid)
            if arquivos:
                for a in arquivos:
                    st.write(f"- {a['name']} (ID: {a['id']})")
                    planilhas.append({"id": a['id'], "name": a['name']})
            else:
                st.write("Nenhum arquivo encontrado ou sem permissão.")
        except Exception as e:
            st.error(f"Erro ao listar pasta {fid}: {e}")
else:
    st.info("Drive API não disponível ou nenhuma pasta configurada para listar.")

# Seleção das planilhas para atualizar
if planilhas:
    nomes_ids = [f"{p['name']} ({p['id']})" for p in planilhas]
    selecionadas = st.multiselect("Selecione as planilhas para atualizar", nomes_ids)
else:
    selecionadas = []

if selecionadas:
    # Carregar dados origem
    with st.spinner("Carregando planilha origem..."):
        try:
            df_origem = carregar_origem(gc, origin_id, origin_sheet)
        except Exception as e:
            st.error(f"Falha ao carregar origem: {e}")
            st.stop()
    st.success("Planilha origem carregada.")

    # Configurações globais
    data_min = st.date_input("Data mínima (filtrar)", value=DEFAULT_DATA_MINIMA)
    dry_run = st.checkbox("Dry-run (não grava)", value=True)
    do_backup = st.checkbox("Fazer backup da aba destino antes de sobrescrever", value=True)

    for sel in selecionadas:
        pid = sel.split("(")[-1].strip(")")
        try:
            sh = gc.open_by_key(pid)
        except Exception as e:
            st.error(f"Não foi possível abrir planilha {pid}: {e}")
            continue

        st.subheader(f"Configurar planilha: {sh.title}")

        # Detectar grupo
        grupo_detectado = None
        try:
            aba_rel = sh.worksheet("rel comp")
            grupo_detectado = (aba_rel.acell("B4").value or "").strip().upper()
        except Exception:
            grupo_detectado = None
        st.write(f"Grupo detectado (B4 de 'rel comp'): **{grupo_detectado or '— não detectado —'}**")

        # Escolher aba destino
        abas = [ws.title for ws in sh.worksheets()]
        dest_options = abas + ["__CRIAR_NOVA_ABA__"]
        dest_choice = st.selectbox(f"Aba destino para {sh.title}", dest_options, index=0 if abas else len(dest_options)-1, key=f"dest_{pid}")
        new_aba_name = ""
        if dest_choice == "__CRIAR_NOVA_ABA__":
            new_aba_name = st.text_input(f"Nome da nova aba para {sh.title}", value="Importado_Fat", key=f"newname_{pid}")

        # Preparar preview
        df = df_origem.copy()
        if grupo_detectado:
            mask = df["Grupo"].astype(str).str.upper() == grupo_detectado
        else:
            mask = pd.Series([True] * len(df), index=df.index)
        mask = mask & df["Data_dt"].notna() & (df["Data_dt"].dt.date >= data_min)
        df_preview = df.loc[mask].copy()
        st.write(f"Linhas a enviar para {sh.title}: **{len(df_preview)}**")
        if not df_preview.empty:
            st.dataframe(df_preview.head(10).drop(columns=["Data_dt"], errors="ignore"), use_container_width=True)

        # Botão para executar atualização
        if st.button(f"Executar atualização para {sh.title}"):
            try:
                # Backup
                if do_backup and dest_choice != "__CRIAR_NOVA_ABA__":
                    try:
                        ws_dest = sh.worksheet(dest_choice)
                        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
                        backup_name = f"{dest_choice}_backup_{ts}"
                        sh.add_worksheet(title=backup_name, rows=str(ws_dest.row_count), cols=str(ws_dest.col_count))
                        values = ws_dest.get_all_values()
                        sh.worksheet(backup_name).update("A1", values, value_input_option="USER_ENTERED")
                        st.success(f"Backup criado: {backup_name}")
                    except Exception as e:
                        st.warning(f"Falha ao criar backup: {e}")

                if dry_run:
                    st.info("Dry-run ativo: nenhuma alteração foi feita.")
                else:
                    # Criar aba nova se necessário
                    if dest_choice == "__CRIAR_NOVA_ABA__":
                        dest = new_aba_name or "Importado_Fat"
                        try:
                            ws_dest = sh.add_worksheet(title=dest, rows=str(max(1000, len(df_preview)+10)), cols=str(max(20, len(df_preview.columns))))
                        except Exception as e:
                            st.error(f"Erro ao criar aba: {e}")
                            continue
                    else:
                        dest = dest_choice
                        ws_dest = sh.worksheet(dest)

                    # Limpar e atualizar
                    ws_dest.clear()
                    values = [df_preview.columns.tolist()] + df_preview.fillna("").astype(str).values.tolist()
                    ws_dest.update("A1", values, value_input_option="USER_ENTERED")
                    st.success(f"{len(df_preview)} linhas gravadas em '{dest}' da planilha {sh.title}.")

            except Exception as e:
                st.error(f"Erro ao atualizar planilha {sh.title}: {e}")

else:
    st.info("Selecione ao menos uma planilha para atualizar.")
