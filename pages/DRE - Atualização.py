# streamlit_app.py
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
# CONFIGURAÇÃO (edite se necessário)
# -----------------------
DEFAULT_FOLDER_IDS = [
    "1ptFvtxYjISfB19S7bU9olMLmAxDTBkOh",  # coloque aqui as pastas que quer listar automaticamente
]

DEFAULT_ORIGIN_SPREADSHEET = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
DEFAULT_ORIGIN_SHEET = "Fat Sistema Externo"
DEFAULT_DATA_MINIMA = (datetime.now() - timedelta(days=365)).date()  # padrão 365 dias

# -----------------------
# UI & CSS
# -----------------------
st.set_page_config(page_title="Atualização e Auditoria - Meio de Pagamento", layout="wide")
st.markdown("""
<style>
.card { background: #ffffff; border-radius: 10px; padding: 16px; box-shadow: 0 2px 8px rgba(0,0,0,0.06); margin-bottom:16px; }
.kv { color:#6c757d; font-size:0.9em; }
.small-muted { color:#6c757d; font-size:0.9em; }
</style>
""", unsafe_allow_html=True)
st.title("📊 Atualização e Auditoria — Faturamento x Meio de Pagamento")

# -----------------------
# AUTENTICAÇÃO
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
            drive_service = build("drive", "v3", credentials=credentials, cache_discovery=False)
        except Exception:
            drive_service = None
    return gc, drive_service, credentials_dict.get("client_email")

try:
    gc, drive_service, service_account_email = autenticar_gspread()
except Exception as e:
    st.error("Erro na autenticação com Google. Verifique st.secrets['GOOGLE_SERVICE_ACCOUNT'].")
    st.stop()

# -----------------------
# UTILITÁRIAS
# -----------------------
def listar_arquivos_pasta(drive_service, pasta_id):
    arquivos = []
    if not drive_service:
        return arquivos
    page_token = None
    query = f"'{pasta_id}' in parents and trashed=false"
    while True:
        try:
            resp = drive_service.files().list(
                q=query,
                spaces="drive",
                fields="nextPageToken, files(id, name, mimeType, shortcutDetails)",
                pageToken=page_token,
                includeItemsFromAllDrives=True,
                supportsAllDrives=True
            ).execute()
            items = resp.get("files", [])
            arquivos.extend(items)
            page_token = resp.get("nextPageToken", None)
            if not page_token:
                break
        except HttpError as e:
            st.error(f"Drive API error listing folder {pasta_id}: {e}")
            break
        except Exception as e:
            st.error(f"Error listing folder {pasta_id}: {e}")
            break
    return arquivos

def testar_abrir_planilha(gc, file_id):
    try:
        sh = gc.open_by_key(file_id)
        return True, sh.title
    except Exception as e:
        return False, str(e)

def carregar_origem_df(gc, spreadsheet_id, sheet_name):
    sh = gc.open_by_key(spreadsheet_id)
    ws = sh.worksheet(sheet_name)
    vals = ws.get_all_values()
    if not vals or len(vals) < 2:
        raise RuntimeError(f"Aba origem '{sheet_name}' vazia ou sem dados.")
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
    """
    Faz backup simples: copia conteúdo para uma nova aba com sufixo timestamp.
    Retorna o título da aba de backup.
    """
    try:
        ws = sh.worksheet(ws_title)
    except Exception:
        return None, "Worksheet not found"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{ws_title}_backup_{ts}"
    # cria backup com mesmas col/linhas aproximadas
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
# AUTO-LISTAGEM COM DIAGNÓSTICO
# -----------------------
def auto_listar_planilhas():
    candidatas = []
    diag = []
    if not DEFAULT_FOLDER_IDS:
        return candidatas, diag
    for fid in DEFAULT_FOLDER_IDS:
        info = {"folder_id": fid, "drive_ok": bool(drive_service), "listed": [], "errors": []}
        if not drive_service:
            info["errors"].append("Drive API não disponível (googleapiclient não inicializada).")
            diag.append(info)
            continue
        arquivos = listar_arquivos_pasta(drive_service, fid)
        if not arquivos:
            info["errors"].append("Nenhum arquivo listado — verifique permissão/ID.")
        else:
            for f in arquivos:
                # se for shortcut, get target
                if f.get("shortcutDetails") and f["shortcutDetails"].get("targetId"):
                    fid_target = f["shortcutDetails"].get("targetId")
                else:
                    fid_target = f["id"]
                ok, title_or_err = testar_abrir_planilha(gc, fid_target)
                entry = {"id": fid_target, "listed_name": f.get("name"), "mimeType": f.get("mimeType"), "gspread_ok": ok, "gspread_title_or_error": title_or_err}
                info["listed"].append(entry)
                if ok:
                    candidatas.append({"id": fid_target, "name": title_or_err, "folder_id": fid})
        diag.append(info)
    return candidatas, diag

if "candidatas" not in st.session_state:
    with st.spinner("Listando planilhas nas pastas..."):
        st.session_state.candidatas, st.session_state.diag = auto_listar_planilhas()

# -----------------------
# TOP BAR (diagnóstico / reload)
# -----------------------
col1, col2, col3 = st.columns([3, 2, 1])
with col1:
    st.markdown(f"<div class='small-muted'>Service account: <b>{service_account_email}</b></div>", unsafe_allow_html=True)
with col2:
    if st.button("🔍 Ver diagnóstico"):
        st.session_state.candidatas, st.session_state.diag = auto_listar_planilhas()
        st.experimental_rerun()
with col3:
    if st.button("🔄 Recarregar lista"):
        st.session_state.candidatas, st.session_state.diag = auto_listar_planilhas()
        st.experimental_rerun()

# -----------------------
# EXIBE DIAGNÓSTICO (opcional)
# -----------------------
if st.session_state.get("diag"):
    with st.expander("Ver diagnóstico detalhado", expanded=False):
        for d in st.session_state.diag:
            st.markdown(f"**Pasta ID:** {d['folder_id']}")
            st.write(f"- Drive API disponível: {d['drive_ok']}")
            if d.get("errors"):
                for e in d["errors"]:
                    st.error(e)
            if d.get("listed"):
                df_diag = pd.DataFrame(d["listed"])
                st.dataframe(df_diag, use_container_width=True)

# -----------------------
# ABAS PRINCIPAIS
# -----------------------
tab1, tab2 = st.tabs(["Atualização (Operacional)", "Auditoria / Logs"])

# -----------------------
# ABA Atualização
# -----------------------
with tab1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("1) Planilhas candidatas (seleção)")

    selecionadas = []  # inicializa para evitar erro
    if not st.session_state.candidatas:
        st.info("Nenhuma planilha encontrada automaticamente. Verifique permissões e DEFAULT_FOLDER_IDS.")
        st.markdown("Compartilhe as pastas com o service account exibido no topo.")
    else:
        df_c = pd.DataFrame(st.session_state.candidatas)[["name", "id", "folder_id"]].rename(columns={"name":"Nome","id":"ID","folder_id":"Pasta ID"})
        st.dataframe(df_c, use_container_width=True)
        options = [f"{r['name']} ({r['id']})" for r in st.session_state.candidatas]
        selecionadas = st.multiselect("Selecione as planilhas para atualizar", options, default=[], key="sel_planilhas")
    st.markdown('</div>', unsafe_allow_html=True)

    if selecionadas:
        # carrega origem (apenas uma vez)
        with st.spinner("Carregando planilha origem..."):
            try:
                df_origem = carregar_origem_df(gc, DEFAULT_ORIGIN_SPREADSHEET, DEFAULT_ORIGIN_SHEET)
            except Exception as e:
                st.error(f"Falha ao carregar origem: {e}")
                st.stop()

        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("2) Configurar por planilha")
        st.write("Para cada planilha selecionada você pode escolher a aba destino (ou criar uma nova). Os dados serão filtrados por Grupo detectado em 'rel comp' (B4) e por data mínima definida abaixo.")
        # parâmetros globais
        col_a, col_b, col_c = st.columns([2,1,1])
        with col_a:
            data_min = st.date_input("Data mínima (filtrar)", value=DEFAULT_DATA_MINIMA)
        with col_b:
            dry_run = st.checkbox("Dry-run (não grava)", value=True)
        with col_c:
            do_backup = st.checkbox("Fazer backup da aba destino antes de sobrescrever", value=True)

        planilhas_config = {}
        for opt in selecionadas:
            pid = opt.split("(")[-1].strip(")")
            try:
                sh = gc.open_by_key(pid)
            except Exception as e:
                st.error(f"Não foi possível abrir planilha {pid}: {e}")
                continue

            with st.expander(f"Configurar: {sh.title}", expanded=False):
                st.markdown(f"**Planilha:** {sh.title} — ID: {pid}")
                grupo_detectado = detectar_grupo_relcomp(sh)
                st.write(f"Grupo detectado (B4 de 'rel comp'): **{grupo_detectado or '— não detectado —'}**")

                # abas existentes + opção criar nova
                abas = [ws.title for ws in sh.worksheets()]
                dest_options = abas + ["__CRIAR_NOVA_ABA__"]
                dest_choice = st.selectbox("Escolha a aba destino", dest_options, index=0 if abas else len(dest_options)-1, key=f"dest_{pid}")
                new_aba_name = ""
                if dest_choice == "__CRIAR_NOVA_ABA__":
                    new_aba_name = st.text_input("Nome da nova aba", value="Importado_Fat", key=f"newname_{pid}")

                # gerar preview
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
                    "dest_aba": (new_aba_name if dest_choice == "__CRIAR_NOVA_ABA__" else dest_choice),
                    "backup": do_backup,
                    "dry_run": dry_run,
                    "df_preview": df_preview,
                    "grupo": grupo_detectado
                }
        st.markdown('</div>', unsafe_allow_html=True)

        # confirmação e execução
        if planilhas_config:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.subheader("3) Confirmar e executar")
            st.write("Revise as configurações acima. Se estiver tudo ok, confirme e execute.")
            confirm = st.checkbox("Confirmo e desejo executar a operação", key="confirm_exec")
            if st.button("Executar agora") and confirm:
                resultados = []
                log_text = []
                total = len(planilhas_config)
                progress = st.progress(0)
                i = 0
                for pid, cfg in planilhas_config.items():
                    i += 1
                    progress.progress(int(i/total*100))
                    sh = cfg["spreadsheet"]
                    dest = cfg["dest_aba"] or "Importado_Fat"
                    df_send = cfg["df_preview"]
                    dry = cfg["dry_run"]
                    do_bkp = cfg["backup"]
                    status = "OK"
                    details = ""
                    try:
                        if df_send is None or df_send.empty:
                            status = "SKIP"
                            details = "Sem linhas para enviar."
                            resultados.append((pid, sh.title, 0, status, details))
                            log_text.append(f"{sh.title}: {details}")
                            continue

                        # criar aba se não existir
                        try:
                            ws_dest = sh.worksheet(dest)
                            aba_existed = True
                        except gspread.exceptions.WorksheetNotFound:
                            ws_dest = None
                            aba_existed = False

                        # backup
                        if do_bkp and aba_existed:
                            bname, berr = backup_worksheet(sh, dest)
                            if berr:
                                log_text.append(f"{sh.title}: Falha backup -> {berr}")
                            else:
                                log_text.append(f"{sh.title}: Backup criado -> {bname}")

                        # se dry-run apenas registra
                        if dry:
                            details = "Dry-run: não foram feitas alterações."
                            resultados.append((pid, sh.title, len(df_send), "DRY-RUN", details))
                            log_text.append(f"{sh.title}: Dry-run -> {len(df_send)} linhas preparadas.")
                            continue

                        # criar aba se necessário
                        if not aba_existed:
                            ws_dest = sh.add_worksheet(title=dest, rows=str(max(1000, len(df_send)+10)), cols=str(max(20, len(df_send.columns))))
                            time.sleep(0.5)  # dar tempo para criação

                        # limpar e escrever
                        ws_dest.clear()
                        values = [df_send.columns.tolist()] + df_send.fillna("").astype(str).values.tolist()
                        ws_dest.update("A1", values, value_input_option="USER_ENTERED")
                        resultados.append((pid, sh.title, len(df_send), "OK", "Dados gravados"))
                        log_text.append(f"{sh.title}: {len(df_send)} linhas gravadas em '{dest}'.")
                    except Exception as e:
                        resultados.append((pid, sh.title, 0, "ERROR", str(e)))
                        log_text.append(f"{sh.title}: ERRO -> {e}")
                progress.progress(100)
                st.success("Operação finalizada")
                df_res = pd.DataFrame(resultados, columns=["ID", "Nome", "Linhas Enviadas", "Status", "Detalhes"])
                st.dataframe(df_res, use_container_width=True)
                with st.expander("Logs"):
                    for line in log_text:
                        st.write(line)
            else:
                st.info("Marque a confirmação e clique em 'Executar agora' para aplicar as alterações.")
            st.markdown('</div>', unsafe_allow_html=True)

# -----------------------
# ABA Auditoria / Logs
# -----------------------
with tab2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Auditoria e logs")
    st.write("- Use a aba Atualização para preparar e executar a atualização.")
    st.write("- Use o diagnóstico (🔍 Ver diagnóstico) se nenhuma planilha aparecer.")
    st.write("- Posso acrescentar relatórios de divergência automaticamente (Faturamento x Meio Pagamento) se desejar.")
    st.markdown('</div>', unsafe_allow_html=True)
