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

# Page config
st.set_page_config(page_title="Atualização e Auditoria", layout="wide")

# CSS básico para "cards"
st.markdown(
    """
    <style>
    .card {
        background: #ffffff;
        border-radius: 8px;
        padding: 16px;
        box-shadow: 0 2px 6px rgba(0,0,0,0.08);
        margin-bottom: 12px;
    }
    .small-muted { color: #6c757d; font-size: 0.9em }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 Atualização e Auditoria — Faturamento x Meio de Pagamento")

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

gc, drive_service, service_account_email = autenticar_gspread()

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
            resp = drive_service.files().list(q=query, spaces="drive", fields="nextPageToken, files(id, name)", pageToken=page_token).execute()
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
# Sidebar - configurações
# -----------------------------------------
with st.sidebar:
    st.header("⚙️ Configurações")
    st.markdown(f"<div class='small-muted'>Service account: <b>{service_account_email}</b></div>", unsafe_allow_html=True)
    st.write("")
    origin_id = st.text_input("ID da planilha origem", value="1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU")
    origin_sheet = st.text_input("Aba origem (na planilha origem)", value="Fat Sistema Externo")
    data_minima = st.date_input("Data mínima (incluir)", value=(datetime.now() - timedelta(days=365)).date())
    st.write("---")
    st.markdown("Pastas para listar (uma ID por linha)")
    folder_ids_text = st.text_area("IDs das pastas (opcional)", value="", height=120)
    folder_ids = [x.strip() for x in folder_ids_text.splitlines() if x.strip()]
    st.write("---")
    st.markdown("IDs/URLs de planilhas (manuais) — uma por linha")
    manual_text = st.text_area("Planilhas manuais (opcional)", value="", height=120)
    st.write("")
    st.button("Listar planilhas candidatas", key="btn_list_cand")

# -----------------------------------------
# Session state inicialização
# -----------------------------------------
if "candidatas" not in st.session_state:
    st.session_state.candidatas = []  # lista de dicts {'id','name','folder_id'}

# -----------------------------------------
# Ações ao clicar em listar candidatas
# -----------------------------------------
if st.session_state.get("btn_list_cand", False) or st.button("Atualizar lista (debug)", key="btn_debug_refresh"):
    st.session_state.candidatas = []
    # 1) listar por pastas (se Drive API estiver disponível)
    if drive_service and folder_ids:
        with st.spinner("Listando arquivos nas pastas..."):
            for fid in folder_ids:
                arquivos = listar_arquivos_pasta(drive_service, fid)
                for a in arquivos:
                    st.session_state.candidatas.append({"id": a["id"], "name": a["name"], "folder_id": fid})
    # 2) adicionar manuais (se fornecidos)
    for line in manual_text.splitlines():
        line = line.strip()
        if not line:
            continue
        if "docs.google.com/spreadsheets" in line:
            parts = line.split("/d/")
            if len(parts) > 1:
                mid = parts[1].split("/")[0]
            else:
                mid = line
        else:
            mid = line
        try:
            sh = gc.open_by_key(mid)
            st.session_state.candidatas.append({"id": mid, "name": sh.title, "folder_id": None})
        except Exception as e:
            st.warning(f"Falha abrindo planilha manual '{mid}': {e}")

# -----------------------------------------
# Abas principais
# -----------------------------------------
tab1, tab2 = st.tabs(["Atualização", "Auditoria Faturamento X Meio Pagamento"])

# -------------------------
# ABA 1: Atualização
# -------------------------
with tab1:
    st.markdown("<​div class='card'>", unsafe_allow_html=True)
    st.subheader("Passo 1 — Planilhas candidatas")
    if not st.session_state.candidatas:
        st.info("Nenhuma planilha candidata listada ainda. Use a sidebar para listar por pastas ou colar IDs/URLs manuais e clique 'Listar planilhas candidatas'.")
    else:
        df_c = pd.DataFrame(st.session_state.candidatas)
        df_display = df_c[["name", "id", "folder_id"]].rename(columns={"name": "Nome", "id": "ID", "folder_id": "Pasta ID"})
        st.dataframe(df_display, use_container_width=True)
        # seleção
        options = [f"{r['name']} ({r['id']})" for r in st.session_state.candidatas]
        selecionadas = st.multiselect("Selecione as planilhas que deseja preparar para atualização", options, default=options[:0], key="sel_planilhas")
    st.markdown("<​/div>", unsafe_allow_html=True)

    # carregar origem (antes de configurar cada planilha)
    if st.session_state.candidatas and selecionadas:
        with st.spinner("Carregando planilha origem..."):
            try:
                df_origem = carregar_origem(gc, origin_id, origin_sheet)
                st.success("Planilha origem carregada com sucesso.")
            except Exception as e:
                st.error(f"Falha ao carregar origem: {e}")
                st.stop()

        # configurar cada planilha selecionada
        st.markdown("<​div class='card'>", unsafe_allow_html=True)
        st.subheader("Passo 2 — Configurar cada planilha selecionada")
        planilhas_config = {}
        for opt in selecionadas:
            pid = opt.split("(")[-1].strip(")")
            try:
                sh = gc.open_by_key(pid)
            except Exception as e:
                st.error(f"Não foi possível abrir {pid}: {e}")
                continue

            st.markdown(f"### {sh.title}")
            grupo_detectado, extra_detectado = detectar_grupo_por_relcomp(sh)
            col1, col2 = st.columns([2, 1])
            with col1:
                st.write(f"Grupo detectado (B4 de 'rel comp'): **{grupo_detectado or '— não detectado —'}**")
                grupo_override = st.text_input("Grupo a usar (se vazio usa detectado)", value=grupo_detectado or "", key=f"grupo_{pid}")
            with col2:
                st.write(f"Filtro extra (B6): **{extra_detectado or '— não detectado —'}**")
                extra_override = st.text_input("Filtro extra (opcional)", value=extra_detectado or "", key=f"extra_{pid}")

            # listar abas existentes e escolha de aba destino
            abas = [ws.title for ws in sh.worksheets()]
            abas_choice = abas + ["__CRIAR_NOVA_ABA__"]
            dest_aba = st.selectbox("Escolha a aba de destino para atualizar", abas_choice, key=f"dest_{pid}")
            new_aba_name = ""
            if dest_aba == "__CRIAR_NOVA_ABA__":
                new_aba_name = st.text_input("Nome da nova aba", value="Importado_Fat", key=f"newaba_{pid}")

            # gerar preview filtrado
            grupo_final = (grupo_override.strip().upper() if grupo_override and grupo_override.strip() else (grupo_detectado or "")).strip().upper()
            df = df_origem.copy()
            if grupo_final:
                mask = df["Grupo"].astype(str).str.upper() == grupo_final
            else:
                mask = pd.Series([True] * len(df), index=df.index)
            mask = mask & df["Data_dt"].notna() & (df["Data_dt"].dt.date >= data_minima)
            df_filtrado = df.loc[mask].copy()
            st.write(f"Linhas que seriam enviadas: **{len(df_filtrado)}**")
            if not df_filtrado.empty:
                with st.expander("Ver amostra (10 linhas)"):
                    st.dataframe(df_filtrado.head(10).drop(columns=["Data_dt"], errors="ignore"), use_container_width=True)

            planilhas_config[pid] = {
                "spreadsheet": sh,
                "grupo": grupo_final,
                "extra": extra_override.strip().upper() if extra_override else extra_detectado,
                "dest_aba": new_aba_name.strip() if dest_aba == "__CRIAR_NOVA_ABA__" else dest_aba,
                "df_preview": df_filtrado
            }
        st.markdown("<​/div>", unsafe_allow_html=True)

        # passo final: confirmação e execução
        if planilhas_config:
            st.markdown("<​div class='card'>", unsafe_allow_html=True)
            st.subheader("Passo 3 — Confirmar e executar")
            confirm = st.checkbox("Confirmo que desejo enviar os dados selecionados para as planilhas/abas escolhidas", key="confirm_send")
            if st.button("Executar Atualização Agora") and confirm:
                resultados = []
                with st.spinner("Executando envios..."):
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
                st.success("Processo concluído. Resumo abaixo:")
                df_res = pd.DataFrame(resultados, columns=["ID", "Nome", "Linhas Enviadas", "Status"])
                st.dataframe(df_res, use_container_width=True)
            else:
                st.info("Marque a caixa de confirmação e clique em 'Executar Atualização Agora' para enviar.")
            st.markdown("<​/div>", unsafe_allow_html=True)

# -------------------------
# ABA 2: Auditoria
# -------------------------
with tab2:
    st.markdown("<​div class='card'>", unsafe_allow_html=True)
    st.subheader("Auditoria — Faturamento x Meio Pagamento")
    st.write("Aqui você pode implementar comparações entre os dados de faturamento e os meios de pagamento.")
    st.write("- Carregue os dados fonte (origem) e os dados de 'Faturamento Meio Pagamento' / 'Tabela Meio Pagamento'.")
    st.write("- Realize validações: somas por data, por PDV, divergências, linhas com valores não numéricos, etc.")
    st.write("")
    st.info("Implementação de auditoria personalizada pode ser adicionada conforme regras do seu negócio.")
    st.markdown("<​/div>", unsafe_allow_html=True)
