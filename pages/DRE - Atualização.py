# streamlit_app.py
import streamlit as st
import pandas as pd
import json
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# tenta importar Drive API (opcional)
try:
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except Exception:
    build = None
    HttpError = Exception

# -----------------------------------------
# Configurações (não exigir input do usuário)
# -----------------------------------------
# IDs das pastas a listar automaticamente
DEFAULT_FOLDER_IDS = [
    "1ptFvtxYjISfB19S7bU9olMLmAxDTBkOh",  # sua pasta
    # adicione outros IDs de pasta se necessário
]

# Planilha origem (padrão)
DEFAULT_ORIGIN_SPREADSHEET = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
DEFAULT_ORIGIN_SHEET = "Fat Sistema Externo"
DEFAULT_DATA_MINIMA = (datetime.now() - timedelta(days=365)).date()

# -----------------------------------------
# Streamlit page config & CSS
# -----------------------------------------
st.set_page_config(page_title="Atualização e Auditoria", layout="wide")
st.markdown("""
<style>
.card {
    background: #ffffff;
    border-radius: 8px;
    padding: 16px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.06);
    margin-bottom: 16px;
}
.small-muted { color:#6c757d; font-size:0.9em; }
.bad { color: #a94442; }
.good { color: #3c763d; }
</style>
""", unsafe_allow_html=True)

st.title("📊 Atualização e Auditoria — Faturamento x Meio Pagamento")

# -----------------------------------------
# Autenticação gspread + Drive (opcional)
# -----------------------------------------
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

# -----------------------------------------
# Funções utilitárias (melhoradas)
# -----------------------------------------
def listar_arquivos_pasta(drive_service, pasta_id):
    """
    Lista arquivos do Drive dentro de uma pasta (suporta Shared Drives).
    Retorna lista de dicts {id, name, mimeType, owners (if available), shortcutTargetId (if shortcut)}.
    """
    arquivos = []
    if not drive_service:
        return arquivos
    page_token = None
    # inclui campos úteis e shortcutDetails para tratar atalhos
    fields = "nextPageToken, files(id, name, mimeType, parents, shortcutDetails)"
    query = f"'{pasta_id}' in parents and trashed=false"
    while True:
        try:
            resp = drive_service.files().list(
                q=query,
                spaces="drive",
                fields=fields,
                pageToken=page_token,
                includeItemsFromAllDrives=True,
                supportsAllDrives=True
            ).execute()
            items = resp.get("files", [])
            for f in items:
                arquivos.append(f)
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
    """
    Tenta abrir a planilha via gspread.open_by_key para confirmar permissão.
    Retorna (ok:bool, title_or_error:str)
    """
    try:
        sh = gc.open_by_key(file_id)
        return True, sh.title
    except Exception as e:
        return False, str(e)

def carregar_origem(gc, spreadsheet_id, sheet_name):
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

# -----------------------------------------
# Auto-listar planilhas candidatas com diagnóstico
# -----------------------------------------
def auto_listar_e_diagnosticar():
    candidatas = []
    diag = []  # lista de diagnósticos por pasta
    if not DEFAULT_FOLDER_IDS:
        return candidatas, diag

    for fid in DEFAULT_FOLDER_IDS:
        info = {"folder_id": fid, "drive_api_ok": bool(drive_service), "listed_files": [], "errors": []}
        if not drive_service:
            info["errors"].append("Drive API cliente não disponível (biblioteca googleapiclient não inicializada).")
            diag.append(info)
            continue

        # 1) tenta listar arquivos na pasta
        try:
            arquivos = listar_arquivos_pasta(drive_service, fid)
            info["listed_files_raw_count"] = len(arquivos)
            if not arquivos:
                info["errors"].append("Nenhum arquivo listado — pode ser permissão ou pasta incorreta.")
            else:
                # filtra apenas planilhas (ou tenta aplicar)
                for f in arquivos:
                    # se for atalho, tenta recuperar targetId
                    if f.get("shortcutDetails") and f["shortcutDetails"].get("targetId"):
                        target_id = f["shortcutDetails"].get("targetId")
                        f_id = target_id
                    else:
                        f_id = f["id"]
                    # testa abrir via gspread para confirmar permissão
                    ok, title_or_err = testar_abrir_planilha(gc, f_id)
                    entry = {
                        "id": f_id,
                        "listed_name": f.get("name"),
                        "mimeType": f.get("mimeType"),
                        "gspread_ok": ok,
                        "gspread_title_or_error": title_or_err
                    }
                    info["listed_files"].append(entry)
                    if ok:
                        # considerar apenas arquivos que são planilha do Google (mimeType check opcional)
                        candidatas.append({"id": f_id, "name": title_or_err, "folder_id": fid})
        except Exception as e:
            info["errors"].append(str(e))
        diag.append(info)
    return candidatas, diag

if "candidatas" not in st.session_state:
    with st.spinner("Listando automaticamente planilhas nas pastas configuradas..."):
        st.session_state.candidatas, st.session_state.diag = auto_listar_e_diagnosticar()

# -----------------------------------------
# Topbar com ações (refresh) e diagnóstico
# -----------------------------------------
col1, col2, col3 = st.columns([3, 2, 1])
with col1:
    st.markdown(f"<div class='small-muted'>Service account: <b>{service_account_email}</b></div>", unsafe_allow_html=True)
with col2:
    if st.button("🔍 Ver diagnóstico"):
        st.session_state.candidatas, st.session_state.diag = auto_listar_e_diagnosticar()
        st.experimental_rerun()
with col3:
    if st.button("🔄 Recarregar lista"):
        st.session_state.candidatas, st.session_state.diag = auto_listar_e_diagnosticar()
        st.experimental_rerun()

# -----------------------------------------
# Mostrar diagnóstico (se solicitado)
# -----------------------------------------
if st.session_state.get("diag"):
    st.markdown("<​div class='card'>", unsafe_allow_html=True)
    st.subheader("Diagnóstico de listagem por pasta")
    for d in st.session_state.diag:
        st.markdown(f"**Pasta ID:** {d['folder_id']}")
        if d.get("errors"):
            for e in d["errors"]:
                st.markdown(f"- <span class='bad'>Erro:</span> {e}", unsafe_allow_html=True)
        st.markdown(f"- Drive API disponível: {'✅' if d.get('drive_api_ok') else '❌'}")
        st.markdown(f"- Arquivos listados (raw): {d.get('listed_files_raw_count', 0)}")
        if d.get("listed_files"):
            df_listed = pd.DataFrame(d["listed_files"])
            st.dataframe(df_listed, use_container_width=True)
    st.markdown("<​/div>", unsafe_allow_html=True)

# -----------------------------------------
# Abas principais
# -----------------------------------------
tab1, tab2 = st.tabs(["Atualização", "Auditoria Faturamento X Meio Pagamento"])

# -------------------------
# ABA 1: Atualização
# -------------------------
with tab1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Passo 1 — Planilhas candidatas (listagem automática)")
    if not st.session_state.candidatas:
        st.info("Nenhuma planilha candidata foi encontrada automaticamente.")
        st.markdown("Possíveis causas e ações:")
        st.markdown("""
        - Verifique se o ID da pasta em DEFAULT_FOLDER_IDS está correto.<br>
        - Compartilhe a pasta com o service account: <b>{}</b> (papel Visualizador ou Editor).<br>
        - Se as planilhas estiverem em um Shared Drive (Unidade), verifique permissão e se o service account tem acesso à Unidade.<br>
        - Certifique-se de que a Drive API esteja ativada no projeto Google Cloud da service account.<br>
        - Use o botão <b>🔍 Ver diagnóstico</b> para ver o que a API retornou (erros / arquivos listados / tentativas de abertura com gspread).
        """.format(service_account_email), unsafe_allow_html=True)
    else:
        df_c = pd.DataFrame(st.session_state.candidatas)
        df_display = df_c[["name", "id", "folder_id"]].rename(columns={"name": "Nome", "id": "ID", "folder_id": "Pasta ID"})
        st.dataframe(df_display, use_container_width=True)
        # seleção automática vazia por padrão; usuário não precisa digitar (seleção opcional)
        selecionadas = st.multiselect("Selecione as planilhas para preparar atualização (opcional)", [f"{r['name']} ({r['id']})" for r in st.session_state.candidatas], default=[], key="sel_planilhas")
    st.markdown("<​/div>", unsafe_allow_html=True)

    # carregar origem automaticamente (apenas quando houver seleção)
    if st.session_state.candidatas and st.session_state.get("sel_planilhas"):
        with st.spinner("Carregando planilha origem..."):
            try:
                df_origem = carregar_origem(gc, DEFAULT_ORIGIN_SPREADSHEET, DEFAULT_ORIGIN_SHEET)
            except Exception as e:
                st.error(f"Falha ao carregar origem: {e}")
                st.stop()
        st.success("Planilha origem carregada.")

        # Configurar cada planilha selecionada
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("Passo 2 — Configurar e revisar por planilha")
        planilhas_config = {}
        for opt in st.session_state.sel_planilhas:
            pid = opt.split("(")[-1].strip(")")
            try:
                sh = gc.open_by_key(pid)
            except Exception as e:
                st.error(f"Erro abrindo planilha {pid}: {e}")
                continue

            st.markdown(f"### {sh.title}")
            # detectar grupo por rel comp (B4)
            grupo_detectado = None
            try:
                grupo_detectado, extra_detectado = None, None
                abas = sh.worksheets()
                aba_rel = next((a for a in abas if "rel comp" in a.title.lower()), None)
                if aba_rel:
                    grupo_detectado = (aba_rel.acell("B4").value or "").strip().upper()
            except Exception:
                grupo_detectado = None

            grupo_final = (grupo_detectado or "").strip().upper()
            st.write(f"Grupo detectado (B4 de 'rel comp'): **{grupo_final or '— não detectado —'}**")

            # escolher aba destino automaticamente: tenta achar "Importado_Fat", senão primeira aba
            abas = [ws.title for ws in sh.worksheets()]
            preferred = "Importado_Fat"
            dest_aba = preferred if preferred in abas else abas[0] if abas else "Importado_Fat"
            st.write(f"Aba destino selecionada automaticamente: **{dest_aba}**")

            # preview filtrado
            df = df_origem.copy()
            if grupo_final:
                mask = df["Grupo"].astype(str).str.upper() == grupo_final
            else:
                mask = pd.Series([True] * len(df), index=df.index)
            mask = mask & df["Data_dt"].notna() & (df["Data_dt"].dt.date >= DEFAULT_DATA_MINIMA)
            df_filtrado = df.loc[mask].copy()
            st.write(f"Linhas que seriam enviadas: **{len(df_filtrado)}**")
            if not df_filtrado.empty:
                with st.expander("Ver amostra (10 linhas)"):
                    st.dataframe(df_filtrado.head(10).drop(columns=["Data_dt"], errors="ignore"), use_container_width=True)

            planilhas_config[pid] = {
                "spreadsheet": sh,
                "grupo": grupo_final,
                "dest_aba": dest_aba,
                "df_preview": df_filtrado
            }
        st.markdown('</div>', unsafe_allow_html=True)

        # confirmação e execução
        if planilhas_config:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.subheader("Passo 3 — Confirmar e executar")
            confirm = st.checkbox("Confirmo que desejo enviar os dados filtrados para as planilhas/abas selecionadas", key="confirm_send_auto")
            if st.button("Executar Atualização Agora") and confirm:
                resultados = []
                with st.spinner("Enviando dados para planilhas..."):
                    for pid, conf in planilhas_config.items():
                        sh = conf["spreadsheet"]
                        df_send = conf["df_preview"]
                        dest_aba = conf["dest_aba"] or "Importado_Fat"
                        try:
                            if df_send is None or df_send.empty:
                                resultados.append((pid, sh.title, 0, "Sem linhas para enviar"))
                                continue
                            if "Data_dt" in df_send.columns:
                                df_send = df_send.drop(columns=["Data_dt"])
                            # criar aba se necessário
                            try:
                                ws_dest = sh.worksheet(dest_aba)
                            except gspread.exceptions.WorksheetNotFound:
                                ws_dest = sh.add_worksheet(title=dest_aba, rows=str(max(1000, len(df_send)+10)), cols=str(len(df_send.columns)))
                            ws_dest.clear()
                            valores = [df_send.columns.tolist()] + df_send.fillna("").astype(str).values.tolist()
                            ws_dest.update("A1", valores, value_input_option="USER_ENTERED")
                            resultados.append((pid, sh.title, len(df_send), "OK"))
                        except Exception as e:
                            resultados.append((pid, sh.title, 0, f"ERRO: {e}"))
                st.success("Processo concluído. Resumo:")
                df_res = pd.DataFrame(resultados, columns=["ID", "Nome", "Linhas Enviadas", "Status"])
                st.dataframe(df_res, use_container_width=True)
            else:
                st.info("Marque a confirmação e clique em 'Executar Atualização Agora' para prosseguir.")
            st.markdown('</div>', unsafe_allow_html=True)

# -------------------------
# ABA 2: Auditoria
# -------------------------
with tab2:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Auditoria — Faturamento x Meio Pagamento")
    st.write("Nesta aba você pode estender a implementação para comparar o Faturamento (origem) x Meio de Pagamento (planilha específica).")
    st.info("Use o botão 🔍 'Ver diagnóstico' para ver detalhes da listagem automática.")
    st.markdown('</div>', unsafe_allow_html=True)
