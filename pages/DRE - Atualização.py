import streamlit as st
import pandas as pd
import json
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import time

try:
    from googleapiclient.discovery import build
except Exception:
    build = None

# Configurações
DEFAULT_FOLDER_IDS = [
    "1ptFvtxYjISfB19S7bU9olMLmAxDTBkOh",
    "1F2Py4eeoqxqrHptgoeUODNXDCUddoU1u",
]
DEFAULT_ORIGIN_SPREADSHEET = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
DEFAULT_ORIGIN_SHEET = "Fat Sistema Externo"
DEFAULT_DATA_MINIMA = (datetime.now() - timedelta(days=365)).date()

OPERACOES = ["Desconto", "Meio Pagamento", "Faturamento"]
ABA_MAP = {"Faturamento": "Importado Fat", "Meio Pagamento": "Meio Pagamento", "Desconto": "Desconto"}

st.set_page_config(page_title="Atualização por Operação", layout="wide")
st.title("Atualização de Planilhas por Operação")

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
    if not drive_service:
        return arquivos
    page_token = None
    query = f"'{pasta_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    while True:
        resp = drive_service.files().list(q=query, spaces="drive", fields="nextPageToken, files(id, name)", pageToken=page_token).execute()
        arquivos.extend(resp.get("files", []))
        page_token = resp.get("nextPageToken", None)
        if not page_token:
            break
    return arquivos

def carregar_origem(gc, origin_spreadsheet_id, origin_sheet_name):
    sh = gc.open_by_key(origin_spreadsheet_id)
    ws = sh.worksheet(origin_sheet_name)
    vals = ws.get_all_values()
    df = pd.DataFrame(vals[1:], columns=vals[0])
    df["Grupo"] = df["Grupo"].astype(str).str.strip().str.upper()
    df["Data_dt"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
    return df

def backup_worksheet(sh, ws_title):
    try:
        ws = sh.worksheet(ws_title)
    except Exception:
        return None, "Worksheet not found"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{ws_title}_backup_{ts}"
    new_ws = sh.add_worksheet(title=backup_name, rows=str(ws.row_count), cols=str(ws.col_count))
    values = ws.get_all_values()
    if values:
        new_ws.update("A1", values, value_input_option="USER_ENTERED")
    return backup_name, None

# Lista planilhas
planilhas = []
if drive_service and DEFAULT_FOLDER_IDS:
    for fid in DEFAULT_FOLDER_IDS:
        arquivos = listar_arquivos_pasta(drive_service, fid)
        for a in arquivos:
            planilhas.append({"id": a["id"], "name": a["name"]})

st.markdown("### Selecione as planilhas e operações para atualizar")

with st.form("form_selecao"):
    selecao = {}
    for p in planilhas:
        cols = st.columns([0.05, 0.6] + [0.12]*len(OPERACOES))
        sel = cols[0].checkbox("", key=f"sel_{p['id']}")
        cols[1].write(p["name"])
        ops_selecionadas = []
        for i, op in enumerate(OPERACOES):
            op_sel = cols[2+i].checkbox("", key=f"sel_{p['id']}_{op.replace(' ','_')}")
            if op_sel:
                ops_selecionadas.append(op)
        if sel and ops_selecionadas:
            selecao[p['id']] = {"name": p["name"], "ops": ops_selecionadas}

    data_min = st.date_input("Data mínima (incluir)", value=DEFAULT_DATA_MINIMA)
    dry_run = st.checkbox("Dry-run (não grava)", value=True)
    do_backup = st.checkbox("Fazer backup antes de sobrescrever", value=True)

    submitted = st.form_submit_button("Atualizar")

if submitted:
    if not selecao:
        st.warning("Selecione ao menos uma planilha e operação para atualizar.")
    else:
        with st.spinner("Carregando dados origem..."):
            df_origem = carregar_origem(gc, DEFAULT_ORIGIN_SPREADSHEET, DEFAULT_ORIGIN_SHEET)
        st.success("Dados origem carregados.")

        resultados = []
        total = sum(len(v["ops"]) for v in selecao.values())
        progresso = st.progress(0)
        i = 0

        for pid, info in selecao.items():
            try:
                sh = gc.open_by_key(pid)
            except Exception as e:
                st.error(f"Erro abrindo planilha {info['name']}: {e}")
                continue

            df = df_origem.copy()
            df = df[df["Data_dt"].notna() & (df["Data_dt"].dt.date >= data_min)]

            for op in info["ops"]:
                i += 1
                aba = ABA_MAP.get(op, op)
                try:
                    try:
                        ws = sh.worksheet(aba)
                        aba_existe = True
                    except gspread.exceptions.WorksheetNotFound:
                        ws = None
                        aba_existe = False

                    if do_backup and aba_existe:
                        backup_worksheet(sh, aba)

                    if dry_run:
                        resultados.append((info["name"], op, "DRY-RUN", len(df)))
                        progresso.progress(int(i/total*100))
                        continue

                    if not aba_existe:
                        ws = sh.add_worksheet(title=aba, rows=str(len(df)+10), cols=str(len(df.columns)))

                    ws.clear()
                    valores = [df.columns.tolist()] + df.fillna("").astype(str).values.tolist()
                    ws.update("A1", valores, value_input_option="USER_ENTERED")
                    resultados.append((info["name"], op, "OK", len(df)))
                except Exception as e:
                    resultados.append((info["name"], op, "ERRO", str(e)))
                progresso.progress(int(i/total*100))

        st.success("Atualização concluída")
        df_res = pd.DataFrame(resultados, columns=["Planilha", "Operação", "Status", "Linhas"])
        st.dataframe(df_res)
