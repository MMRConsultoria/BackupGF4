import streamlit as st
import pandas as pd
import json
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import time

# opcional: Drive API para listar pastas
try:
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except Exception:
    build = None
    HttpError = Exception

# -----------------------
# CONFIGURAÇÕES
# -----------------------
DEFAULT_FOLDER_IDS = [
    "1ptFvtxYjISfB19S7bU9olMLmAxDTBkOh",
    "1F2Py4eeoqxqrHptgoeUODNXDCUddoU1u",
]
DEFAULT_ORIGIN_SPREADSHEET = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
DEFAULT_ORIGIN_SHEET = "Fat Sistema Externo"
DEFAULT_DATA_MINIMA = (datetime.now() - timedelta(days=365)).date()

# Operações/abas fixas que o usuário quer controlar (rótulos visuais no grid)
OPERACOES = ["Desconto", "Meio Pagamento", "Faturamento"]
# Mapeamento visual -> aba real (Faturamento -> Importado Fat)
ABA_MAP = {"Faturamento": "Importado Fat", "Meio Pagamento": "Meio Pagamento", "Desconto": "Desconto"}

# -----------------------
# UI
# -----------------------
st.set_page_config(page_title="Atualização por Operação — Faturamento x Meio de Pagamento", layout="wide")
st.title("📋 Seleção por Operação — Atualização de Planilhas")

st.markdown("""
<style>
.row-name { font-weight: 600; }
.small-muted { color:#6c757d; font-size:0.9em; }
.header-cell { font-weight:700; }
</style>
""", unsafe_allow_html=True)

# -----------------------
# AUTENTICAÇÃO GSPREAD (+ Drive opcional)
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
    return gc, drive_service, credentials_dict.get("client_email")

try:
    gc, drive_service, service_account_email = autenticar_gspread()
except Exception as e:
    st.error("Erro na autenticação com Google. Verifique st.secrets['GOOGLE_SERVICE_ACCOUNT'].")
    st.stop()

st.markdown(f"<div class='small-muted'>Service account: <b>{service_account_email}</b></div>", unsafe_allow_html=True)

# -----------------------
# FUNÇÕES AUX
# -----------------------
def listar_arquivos_pasta(drive_service, pasta_id):
    arquivos = []
    if not drive_service:
        return arquivos
    page_token = None
    query = f"'{pasta_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    while True:
        try:
            resp = drive_service.files().list(q=query, spaces="drive", fields="nextPageToken, files(id, name)", pageToken=page_token).execute()
            arquivos.extend(resp.get("files", []))
            page_token = resp.get("nextPageToken", None)
            if not page_token:
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

def detectar_grupo_relcomp(sh):
    try:
        aba_rel = next((a for a in sh.worksheets() if "rel comp" in a.title.lower()), None)
        if not aba_rel:
            return None
        v = aba_rel.acell("B4").value
        return (v or "").strip().upper()
    except Exception:
        return None

def backup_worksheet(sh, ws_title):
    try:
        ws = sh.worksheet(ws_title)
    except Exception:
        return None, "Worksheet not found"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{ws_title}_backup_{ts}"
    rows = max(1000, ws.row_count)
    cols = max(20, ws.col_count)
    try:
        new_ws = sh.add_worksheet(title=backup_name, rows=str(rows), cols=str(cols))
        values = ws.get_all_values()
        if values:
            new_ws.update("A1", values, value_input_option="USER_ENTERED")
        return backup_name, None
    except Exception as e:
        return None, str(e)

# -----------------------
# SIDEBAR: parâmetros
# -----------------------
st.sidebar.header("Parâmetros")
origin_id = st.sidebar.text_input("ID planilha origem", value=DEFAULT_ORIGIN_SPREADSHEET)
origin_sheet = st.sidebar.text_input("Aba origem (na planilha origem)", value=DEFAULT_ORIGIN_SHEET)
data_minima = st.sidebar.date_input("Data mínima (incluir)", value=DEFAULT_DATA_MINIMA)
folder_ids_text = st.sidebar.text_area("IDs das pastas (uma por linha) — opcional", value="\n".join(DEFAULT_FOLDER_IDS), height=120)
folder_ids = [s.strip() for s in folder_ids_text.splitlines() if s.strip()]

# -----------------------
# CARREGA A LISTA DE PLANILHAS (a partir das pastas)
# -----------------------
planilhas = []
if drive_service and folder_ids:
    for fid in folder_ids:
        arquivos = listar_arquivos_pasta(drive_service, fid)
        if arquivos:
            for a in arquivos:
                planilhas.append({"id": a["id"], "name": a["name"], "folder_id": fid})
else:
    if not drive_service:
        st.warning("Drive API não disponível — verifique googleapiclient/credentials.")
    if not folder_ids:
        st.info("Insira IDs de pasta no sidebar para listar planilhas automaticamente.")

# -----------------------
# CONTROLES GLOBAIS (selecionar tudo por coluna)
# -----------------------
st.markdown("### Escolha as operações por planilha")
col_headers = st.columns([0.05, 0.45] + [0.15]*len(OPERACOES))

with col_headers[0]:
    st.write("")  # espaço para checkbox de linha
with col_headers[1]:
    st.markdown("**Planilha**")
# criar checkboxes globais para cada operação
global_ops_state = {}
for i, op in enumerate(OPERACOES):
    key = f"global_op_{op.replace(' ','_')}"
    default = True
    # se já existir estado anterior, mantemos; caso contrário, default True
    val = st.checkbox(op, value=st.session_state.get(key, default), key=key)
    global_ops_state[op] = val

st.write("")  # separador visual

# -----------------------
# PREPARAR KEYS NA session_state PARA OS CHECKBOXES (inicializar)
# -----------------------
# Inicializa chave por planilha/operação com o valor global se não existir
for p in planilhas:
    pid = p["id"]
    # checkbox linha selecionada
    row_key = f"row_sel_{pid}"
    if row_key not in st.session_state:
        st.session_state[row_key] = True
    for op in OPERACOES:
        op_key = f"{pid}__{op.replace(' ','_')}"
        if op_key not in st.session_state:
            st.session_state[op_key] = global_ops_state.get(op, True)

# -----------------------
# RENDERIZA GRID (tabela de checkboxes)
# -----------------------
st.markdown("#### Seleção por planilha / operação")
rows_container = st.container()
with rows_container:
    for p in planilhas:
        pid = p["id"]
        cols = st.columns([0.05, 0.45] + [0.15]*len(OPERACOES))
        # checkbox para selecionar a linha inteira
        with cols[0]:
            row_checked = st.checkbox("", value=st.session_state.get(f"row_sel_{pid}", True), key=f"row_sel_{pid}")
        with cols[1]:
            st.markdown(f"<div class='row-name'>{p['name']}</div>", unsafe_allow_html=True)

        # checkboxes por operação (visuais)
        for j, op in enumerate(OPERACOES):
            op_key = f"{pid}__{op.replace(' ','_')}"
            with cols[2 + j]:
                # se a linha não estiver selecionada, mostramos unchecked — mas mantemos estado no session_state
                if row_checked:
                    checked = st.checkbox("", value=st.session_state.get(op_key, True), key=op_key)
                else:
                    # ainda coloca o widget mas forçado para False visualmente; grava False no state
                    # Para evitar que a checkbox reapareça true quando row é desmarcada, atualizamos state:
                    st.session_state[op_key] = False
                    checked = st.checkbox("", value=False, key=op_key)
        st.write("")  # espaçamento leve

# resumo de quantos pares selecionados (opções ativas)
def contar_selecionados():
    total_pairs = 0
    planilhas_selecionadas = {}
    for p in planilhas:
        pid = p["id"]
        row_key = f"row_sel_{pid}"
        if not st.session_state.get(row_key, False):
            continue
        ops = []
        for op in OPERACOES:
            op_key = f"{pid}__{op.replace(' ','_')}"
            if st.session_state.get(op_key, False):
                ops.append(op)
        if ops:
            planilhas_selecionadas[pid] = ops
            total_pairs += len(ops)
    return total_pairs, planilhas_selecionadas

total_pairs, planilhas_selecionadas = contar_selecionados()
st.write(f"Total de (planilha × operação) selecionados: **{total_pairs}**")

# -----------------------
# EXECUÇÃO: carregar origem e executar atualizações por par planilha/aba
# -----------------------
if total_pairs > 0:
    with st.spinner("Carregando planilha origem..."):
        try:
            df_origem = carregar_origem(gc, origin_id, origin_sheet)
        except Exception as e:
            st.error(f"Falha ao carregar origem: {e}")
            st.stop()
    st.success("Planilha origem carregada.")

    col_a, col_b, col_c = st.columns([2,1,1])
    with col_a:
        data_min = st.date_input("Data mínima (filtrar)", value=data_minima)
    with col_b:
        dry_run = st.checkbox("Dry-run (não grava)", value=True)
    with col_c:
        do_backup = st.checkbox("Fazer backup da aba destino antes de sobrescrever", value=True)

    if st.checkbox("Confirmo e desejo executar a operação", key="confirm_exec"):
        if st.button("Executar agora"):
            resultados = []
            logs = []
            total_tasks = sum(len(ops) for ops in planilhas_selecionadas.values())
            progress = st.progress(0)
            i_task = 0

            for pid, ops in planilhas_selecionadas.items():
                # abrir planilha
                try:
                    sh = gc.open_by_key(pid)
                    sheet_name = sh.title
                except Exception as e:
                    logs.append(f"{pid}: erro abrindo planilha -> {e}")
                    for op in ops:
                        resultados.append((pid, f"(ID) {pid}", op, 0, "ERROR", f"Erro ao abrir planilha: {e}"))
                    continue

                # detectar grupo (para filtro)
                grupo_detectado = detectar_grupo_relcomp(sh)

                # preview df filtrado por grupo/data
                df = df_origem.copy()
                if grupo_detectado:
                    mask = df["Grupo"].astype(str).str.upper() == grupo_detectado
                else:
                    mask = pd.Series([True] * len(df), index=df.index)
                mask = mask & df["Data_dt"].notna() & (df["Data_dt"].dt.date >= data_min)
                df_preview = df.loc[mask].copy()

                if df_preview.empty:
                    logs.append(f"{sheet_name}: sem linhas depois do filtro (grupo/data)")
                    for op in ops:
                        resultados.append((pid, sheet_name, op, 0, "SKIP", "Sem linhas após filtro"))
                    # atualizar progress
                    i_task += len(ops)
                    progress.progress(int(i_task/total_tasks*100))
                    continue

                # para cada operação, gravar em sua aba correspondente
                for op in ops:
                    i_task += 1
                    dest_aba = ABA_MAP.get(op, op)
                    try:
                        # verificar aba
                        try:
                            ws_dest = sh.worksheet(dest_aba)
                            aba_existed = True
                        except gspread.exceptions.WorksheetNotFound:
                            ws_dest = None
                            aba_existed = False

                        # backup opcional
                        if do_backup and aba_existed:
                            bname, berr = backup_worksheet(sh, dest_aba)
                            if berr:
                                logs.append(f"{sheet_name}/{dest_aba}: falha backup -> {berr}")
                            else:
                                logs.append(f"{sheet_name}/{dest_aba}: backup criado -> {bname}")

                        # dry-run
                        if dry_run:
                            resultados.append((pid, sheet_name, op, len(df_preview), "DRY-RUN", "Não gravado (dry-run)"))
                            logs.append(f"{sheet_name}/{dest_aba}: dry-run -> {len(df_preview)} linhas preparadas.")
                            progress.progress(int(i_task/total_tasks*100))
                            continue

                        # criar aba se não existir
                        if not aba_existed:
                            ws_dest = sh.add_worksheet(title=dest_aba, rows=str(max(1000, len(df_preview)+10)), cols=str(max(20, len(df_preview.columns))))
                            time.sleep(0.4)

                        # escrever dados
                        ws_dest.clear()
                        values = [df_preview.columns.tolist()] + df_preview.fillna("").astype(str).values.tolist()
                        ws_dest.update("A1", values, value_input_option="USER_ENTERED")
                        resultados.append((pid, sheet_name, op, len(df_preview), "OK", f"Gravado em '{dest_aba}'"))
                        logs.append(f"{sheet_name}/{dest_aba}: {len(df_preview)} linhas gravadas.")
                    except Exception as e:
                        resultados.append((pid, sheet_name, op, 0, "ERROR", str(e)))
                        logs.append(f"{sheet_name}/{dest_aba}: ERRO -> {e}")
                    progress.progress(int(i_task/total_tasks*100))

            st.success("Operação finalizada")
            df_res = pd.DataFrame(resultados, columns=["ID","Planilha","Operação","Linhas","Status","Detalhes"])
            st.dataframe(df_res, use_container_width=True)
            with st.expander("Logs detalhados"):
                for l in logs:
                    st.write(l)
else:
    st.info("Nenhuma operação selecionada. Marque as operações desejadas ao lado das planilhas.")
