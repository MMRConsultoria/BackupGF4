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

# -----------------------
# CONFIGURAÇÃO
# -----------------------
DEFAULT_FOLDER_IDS = [
    "1ptFvtxYjISfB19S7bU9olMLmAxDTBkOh",
    "1F2Py4eeoqxqrHptgoeUODNXDCUddoU1u",
]
DEFAULT_ORIGIN_SPREADSHEET = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
DEFAULT_ORIGIN_SHEET = "Fat Sistema Externo"
DEFAULT_DATA_MINIMA = (datetime.now() - timedelta(days=365)).date()  # padrão 365 dias

# Abas fixas para atualizar
ABAS_FIXAS = [
    "Meio Pagamento",
    "Desconto",
    "Volumetria",
    "Importado Fat",  # visual será "Faturamento"
]

# -----------------------
# UI
# -----------------------
st.set_page_config(page_title="Atualização e Auditoria - Meio de Pagamento", layout="wide")
st.title("📊 Atualização e Auditoria — Faturamento x Meio de Pagamento")

st.markdown("""
<style>
.card { background: #ffffff; border-radius: 10px; padding: 12px; box-shadow: 0 2px 8px rgba(0,0,0,0.04); margin-bottom:12px; }
.small-muted { color:#6c757d; font-size:0.9em; }
.planilha-row { display: flex; align-items: center; margin-bottom: 6px; }
.planilha-name { flex: 1; }
.aba-select { width: 200px; margin-left: 12px; }
</style>
""", unsafe_allow_html=True)

# -----------------------
# AUTENTICAÇÃO gspread (+ Drive opcional)
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
# FUNÇÕES AUXILIARES
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

def detectar_grupo_relcomp(sh):
    try:
        abas = sh.worksheets()
        aba_rel = next((a for a in abas if "rel comp" in a.title.lower()), None)
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
# CARREGA PLANILHAS DAS PASTAS
# -----------------------
planilhas = []
if drive_service and folder_ids:
    for fid in folder_ids:
        try:
            arquivos = listar_arquivos_pasta(drive_service, fid)
            if arquivos:
                for a in arquivos:
                    planilhas.append({"id": a["id"], "name": a["name"], "folder_id": fid})
        except Exception as e:
            st.error(f"Erro listando pasta {fid}: {e}")
else:
    if not drive_service:
        st.warning("Drive API não disponível — verifique googleapiclient/credentials.")
    if not folder_ids:
        st.info("Insira IDs de pasta no sidebar para listar planilhas automaticamente.")

# -----------------------
# SELEÇÃO: checkbox + selectbox para aba fixa
# -----------------------
st.markdown("### Selecione as planilhas para atualizar e escolha a aba destino")

selecionadas = []
abas_selecionadas = {}

for p in planilhas:
    cols = st.columns([0.05, 0.6, 0.35])
    with cols[0]:
        checked = st.checkbox("", value=True, key=f"chk_{p['id']}")
    with cols[1]:
        st.markdown(f"**{p['name']}**")
    with cols[2]:
        # substitui label "Importado Fat" por "Faturamento" visualmente
        opcoes_aba = [aba if aba != "Importado Fat" else "Faturamento" for aba in ABAS_FIXAS]
        aba_escolhida = st.selectbox(f"Aba destino {p['id']}", opcoes_aba, key=f"aba_{p['id']}")
    if checked:
        selecionadas.append(p)
        # guarda aba escolhida real (troca "Faturamento" por "Importado Fat")
        abas_selecionadas[p['id']] = "Importado Fat" if aba_escolhida == "Faturamento" else aba_escolhida

st.write(f"Total selecionadas: {len(selecionadas)}")

# -----------------------
# EXECUÇÃO PARA SELECIONADAS
# -----------------------
if selecionadas:
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

    planilhas_config = {}
    for p in selecionadas:
        pid = p["id"]
        try:
            sh = gc.open_by_key(pid)
        except Exception as e:
            st.error(f"Não foi possível abrir planilha {pid}: {e}")
            continue

        with st.expander(f"Configurar: {sh.title}", expanded=False):
            st.markdown(f"**Planilha:** {sh.title} — ID: {pid}")
            grupo_detectado = detectar_grupo_relcomp(sh)
            st.write(f"Grupo detectado (B4 de 'rel comp'): **{grupo_detectado or '— não detectado —'}**")

            dest_aba = abas_selecionadas.get(pid, "Importado Fat")

            # preview
            df = df_origem.copy()
            if grupo_detectado:
                mask = df["Grupo"].astype(str).str.upper() == grupo_detectado
            else:
                mask = pd.Series([True] * len(df), index=df.index)
            mask = mask & df["Data_dt"].notna() & (df["Data_dt"].dt.date >= data_min)
            df_preview = df.loc[mask].copy()
            st.write(f"Linhas a enviar: **{len(df_preview)}**")
            if not df_preview.empty:
                st.dataframe(df_preview.head(10).drop(columns=["Data_dt"], errors="ignore"), use_container_width=True)

            planilhas_config[pid] = {
                "spreadsheet": sh,
                "dest_aba": dest_aba,
                "backup": do_backup,
                "dry_run": dry_run,
                "df_preview": df_preview,
                "grupo": grupo_detectado
            }

    if planilhas_config:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Executar atualização para planilhas configuradas")
        confirm = st.checkbox("Confirmo e desejo executar a operação", key="confirm_exec")
        if st.button("Executar agora") and confirm:
            resultados = []
            logs = []
            total = len(planilhas_config)
            progress = st.progress(0)
            i = 0
            for pid, cfg in planilhas_config.items():
                i += 1
                progress.progress(int(i/total*100))
                sh = cfg["spreadsheet"]
                dest = cfg["dest_aba"] or "Importado Fat"
                df_send = cfg["df_preview"]
                dry = cfg["dry_run"]
                do_bkp = cfg["backup"]
                try:
                    if df_send is None or df_send.empty:
                        resultados.append((pid, sh.title, 0, "SKIP", "Sem linhas"))
                        logs.append(f"{sh.title}: Sem linhas para enviar.")
                        continue

                    try:
                        ws_dest = sh.worksheet(dest)
                        aba_existed = True
                    except gspread.exceptions.WorksheetNotFound:
                        ws_dest = None
                        aba_existed = False

                    if do_bkp and aba_existed:
                        bname, berr = backup_worksheet(sh, dest)
                        if berr:
                            logs.append(f"{sh.title}: Falha backup -> {berr}")
                        else:
                            logs.append(f"{sh.title}: Backup criado -> {bname}")

                    if dry:
                        resultados.append((pid, sh.title, len(df_send), "DRY-RUN", "Não gravado"))
                        logs.append(f"{sh.title}: Dry-run -> {len(df_send)} linhas preparadas.")
                        continue

                    if not aba_existed:
                        ws_dest = sh.add_worksheet(title=dest, rows=str(max(1000, len(df_send)+10)), cols=str(max(20, len(df_send.columns))))
                        time.sleep(0.5)

                    ws_dest.clear()
                    values = [df_send.columns.tolist()] + df_send.fillna("").astype(str).values.tolist()
                    ws_dest.update("A1", values, value_input_option="USER_ENTERED")
                    resultados.append((pid, sh.title, len(df_send), "OK", f"Gravado em '{dest}'"))
                    logs.append(f"{sh.title}: {len(df_send)} linhas gravadas em '{dest}'.")
                except Exception as e:
                    resultados.append((pid, sh.title, 0, "ERROR", str(e)))
                    logs.append(f"{sh.title}: ERRO -> {e}")
            progress.progress(100)
            st.success("Operação finalizada")
            df_res = pd.DataFrame(resultados, columns=["ID", "Nome", "Linhas Enviadas", "Status", "Detalhes"])
            st.dataframe(df_res, use_container_width=True)
            with st.expander("Logs"):
                for line in logs:
                    st.write(line)
        else:
            st.info("Marque a confirmação e clique em 'Executar agora' para aplicar as alterações.")
        st.markdown('</div>', unsafe_allow_html=True)
else:
    st.info("Nenhuma planilha selecionada. Desmarque as que não quiser atualizar na lista acima.")
