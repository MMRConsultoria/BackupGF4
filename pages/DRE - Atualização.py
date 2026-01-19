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
# Cole aqui os IDs das pastas que devem ser listadas automaticamente
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
.kv { color:#6c757d; font-size:0.9em; }
.header-row { display:flex; gap:16px; align-items:center; justify-content:space-between; }
.small-muted { color:#6c757d; font-size:0.9em; }
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
# Funções utilitárias
# -----------------------------------------
def listar_arquivos_pasta(drive_service, pasta_id):
    arquivos = []
    if not drive_service:
        return arquivos
    page_token = None
    query = f"'{pasta_id}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
    while True:
        try:
            resp = drive_service.files().list(
                q=query,
                spaces="drive",
                fields="nextPageToken, files(id, name)",
                pageToken=page_token
            ).execute()
            arquivos.extend(resp.get("files", []))
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

def detectar_grupo_por_relcomp(sh):
    try:
        abas = sh.worksheets()
        aba_rel = next((a for a in abas if "rel comp" in a.title.lower()), None)
        if not aba_rel:
            return None, None
        grupo = aba_rel.acell("B4").value
        extra = aba_rel.acell("B6").value if aba_rel.row_count >= 6 else None
        grupo = grupo.strip().upper() if grupo else None
        extra = extra.strip().upper() if extra else None
        return grupo, extra
    except Exception:
        return None, None

# -----------------------------------------
# Auto-listar planilhas candidatas (executa ao carregar o app)
# -----------------------------------------
if "candidatas" not in st.session_state:
    st.session_state.candidatas = []
    # Tenta listar automaticamente das pastas configuradas
    if drive_service and DEFAULT_FOLDER_IDS:
        with st.spinner("Listando planilhas nas pastas configuradas..."):
            for fid in DEFAULT_FOLDER_IDS:
                arquivos = listar_arquivos_pasta(drive_service, fid)
                for a in arquivos:
                    st.session_state.candidatas.append({"id": a["id"], "name": a["name"], "folder_id": fid})
    # Se não encontrou nada, tentamos apenas deixar a lista vazia (o app mostrará instruções)
    # Você pode adicionar aqui DEFAULT_MANUAL_SPREADSHEET_IDS se quiser incluir alguns hardcoded.

# -----------------------------------------
# Topbar com ações (refresh)
# -----------------------------------------
col1, col2 = st.columns([3, 1])
with col1:
    st.markdown(f"<div class='small-muted'>Service account: <b>{service_account_email}</b></div>", unsafe_allow_html=True)
with col2:
    if st.button("🔄 Recarregar lista"):
        # reload candidatas
        st.session_state.candidatas = []
        if drive_service and DEFAULT_FOLDER_IDS:
            with st.spinner("Relendo planilhas nas pastas..."):
                for fid in DEFAULT_FOLDER_IDS:
                    arquivos = listar_arquivos_pasta(drive_service, fid)
                    for a in arquivos:
                        st.session_state.candidatas.append({"id": a["id"], "name": a["name"], "folder_id": fid})
        st.experimental_rerun()

# -----------------------------------------
# Abas principais
# -----------------------------------------
tab1, tab2 = st.tabs(["Atualização", "Auditoria Faturamento X Meio Pagamento"])

# -------------------------
# ABA 1: Atualização
# -------------------------
with tab1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Passo 1 — Planilhas candidatas")
    if not st.session_state.candidatas:
        st.info("Nenhuma planilha candidata foi encontrada automaticamente. Verifique: 1) DEFAULT_FOLDER_IDS; 2) se o service account tem acesso às PASTAS configuradas; 3) clique em '🔄 Recarregar lista'.")
        st.markdown("Se preferir, edite `DEFAULT_FOLDER_IDS` no código e adicione os IDs das pastas que deseja listar automaticamente.")
    else:
        df_c = pd.DataFrame(st.session_state.candidatas)
        df_display = df_c[["name", "id", "folder_id"]].rename(columns={"name": "Nome", "id": "ID", "folder_id": "Pasta ID"})
        st.dataframe(df_display, use_container_width=True)
        options = [f"{r['name']} ({r['id']})" for r in st.session_state.candidatas]
        selecionadas = st.multiselect("Selecione as planilhas para preparar atualização", options, default=[], key="sel_planilhas")
    st.markdown('</div>', unsafe_allow_html=True)

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
            grupo_detectado, extra_detectado = detectar_grupo_por_relcomp(sh)
            colA, colB = st.columns([2, 1])
            with colA:
                st.write(f"Grupo detectado (B4 de 'rel comp'): **{grupo_detectado or '— não detectado —'}**")
                # se não detectado, usamos texto detectado (não há input conforme pedido)
                grupo_final = (grupo_detectado or "").strip().upper()
            with colB:
                st.write(f"Filtro extra (B6): **{extra_detectado or '— não detectado —'}**")

            # escolher aba destino automaticamente: tenta achar "Importado_Fat", senão primeira aba
            abas = [ws.title for ws in sh.worksheets()]
            preferred = "Importado_Fat"
            dest_aba = preferred if preferred in abas else abas[0] if abas else "Importado_Fat"

            st.write(f"Aba destino selecionada automaticamente: **{dest_aba}** (se quiser mudar, edite o nome aqui no código)")

            # preview dos dados filtrados
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
                "extra": extra_detectado,
                "dest_aba": dest_aba,
                "df_preview": df_filtrado
            }
        st.markdown('</div>', unsafe_allow_html=True)

        # confirmação e execução
        if planilhas_config:
            st.markdown('<div class="card">', unsafe_allow_html=True)
            st.subheader("Passo 3 — Confirmar e executar")
            st.write("Aviso: esta ação irá sobrescrever a aba destino selecionada em cada planilha.")
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
                            # opcional: fazer backup da aba antes de limpar (pode ser adicionado)
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
    st.write("- Carregar aqui os dados de 'Faturamento Meio Pagamento' e 'Tabela Meio Pagamento' e gerar relatórios de divergência.")
    st.info("Se quiser, eu adiciono as rotinas de auditoria (somas por data/PDV, diferenças, linhas inválidas) conforme suas regras.")
    st.markdown('</div>', unsafe_allow_html=True)
