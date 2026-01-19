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

st.set_page_config(page_title="Atualização e Auditoria - Meio de Pagamento", layout="wide")
st.title("Atualização e Auditoria — Faturamento x Meio de Pagamento")

# -----------------------
# Configurações iniciais
# -----------------------
# Valores padrão — ajuste conforme seu ambiente
DEFAULT_FOLDER_IDS = [
    # cole aqui as IDs das pastas a listar (opcional)
    # "1ptFvtxYjISfB19S7bU9olMLmAxDTBkOh",
    # "1F2Py4eeoqxqrHptgoeUODNXDCUddoU1u",
]
DEFAULT_ORIGIN_SPREADSHEET = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
DEFAULT_ORIGIN_SHEET = "Fat Sistema Externo"
DEFAULT_DATA_MINIMA = datetime.now() - timedelta(days=365)  # exemplo: últimos 365 dias

# -----------------------
# Autenticação gspread (+drive opcional)
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
    return gc, drive_service

gc, drive_service = autenticar_gspread()

# -----------------------
# Funções utilitárias
# -----------------------
def listar_arquivos_pasta(drive_service, pasta_id):
    """Retorna lista de dicts {'id','name'} de spreadsheets na pasta."""
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

def obter_candidatas_de_pastas(drive_service, folder_ids):
    """Retorna lista de candidatos (id, name) consultando cada pasta."""
    resultados = []
    for fid in folder_ids:
        if drive_service:
            arquivos = listar_arquivos_pasta(drive_service, fid)
            for a in arquivos:
                resultados.append({"id": a["id"], "name": a["name"], "folder_id": fid})
        else:
            # sem drive_service, não conseguimos listar automaticamente
            pass
    return resultados

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

def detectar_grupo_por_relcomp(spreadsheet):
    """Tenta encontrar aba 'rel comp' e ler B4 (Grupo) e B6 (filtro extra)."""
    try:
        abas = spreadsheet.worksheets()
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

# -----------------------
# Sidebar: parâmetros
# -----------------------
st.sidebar.header("Parâmetros")
origin_id = st.sidebar.text_input("ID planilha origem", value=DEFAULT_ORIGIN_SPREADSHEET)
origin_sheet = st.sidebar.text_input("Aba origem (na planilha origem)", value=DEFAULT_ORIGIN_SHEET)
data_minima = st.sidebar.date_input("Data mínima (incluir)", value=DEFAULT_DATA_MINIMA.date())
# receber lista de folder ids (multiline) ou manual entry de spreadsheet IDs caso Drive API não disponível
use_drive = build is not None and drive_service is not None
if use_drive:
    st.sidebar.markdown("Drive API: disponível — será listado automaticamente o conteúdo das pastas.")
else:
    st.sidebar.markdown("Drive API: indisponível — insira manualmente IDs de planilhas destino (uma por linha).")

folder_ids_text = st.sidebar.text_area("IDs das pastas (uma por linha) — opcional", value="\n".join(DEFAULT_FOLDER_IDS), height=80)
folder_ids = [s.strip() for s in folder_ids_text.splitlines() if s.strip()]

manual_spreadsheets_text = st.sidebar.text_area("IDs ou URLs de planilhas destino (uma por linha) — usado se Drive API indisponível ou para adicionar manualmente", height=120)
manual_ids = []
for line in manual_spreadsheets_text.splitlines():
    line = line.strip()
    if not line:
        continue
    # extrai id se for URL
    if "docs.google.com/spreadsheets" in line:
        parts = line.split("/d/")
        if len(parts) > 1:
            id_part = parts[1].split("/")[0]
            manual_ids.append(id_part)
        else:
            manual_ids.append(line)
    else:
        manual_ids.append(line)

# botão para listar candidatos
if st.sidebar.button("Listar planilhas candidatas"):
    st.session_state.candidatas = []
    # 1) obtém a lista por pastas (se drive disponível)
    if folder_ids and use_drive:
        with st.spinner("Listando planilhas nas pastas..."):
            candidatas = obter_candidatas_de_pastas(drive_service, folder_ids)
            st.session_state.candidatas = candidatas
    else:
        st.session_state.candidatas = []

    # 2) adiciona manuais (ids)
    for mid in manual_ids:
        try:
            sh = gc.open_by_key(mid)
            st.session_state.candidatas.append({"id": mid, "name": sh.title, "folder_id": None})
        except Exception as e:
            st.warning(f"Falha abrindo planilha manual '{mid}': {e}")

# inicializa lista vazia se não existir
if "candidatas" not in st.session_state:
    st.session_state.candidatas = []

# mostra candidatos encontrados
st.header("Passo 1 — Candidatas à atualização")
if not st.session_state.candidatas:
    st.info("Nenhuma planilha candidata listada ainda. Use a sidebar para listar por pastas ou colar IDs/URLs manualmente e clique 'Listar planilhas candidatas'.")
else:
    df_cand = pd.DataFrame(st.session_state.candidatas)
    df_cand_display = df_cand[["name", "id", "folder_id"]].rename(columns={"name": "Nome", "id": "ID", "folder_id": "Pasta ID"})
    st.dataframe(df_cand_display, use_container_width=True)

    # seletor de quais planilhas atualizar
    options_for_select = [f"{r['name']} || {r['id']}" for r in st.session_state.candidatas]
    selected = st.multiselect("Selecione as planilhas que deseja preparar para atualização", options_for_select, default=options_for_select)
    selected_ids = [opt.split("||")[-1].strip() for opt in selected]

    # Carrega origem (uma vez)
    loaded_origin = None
    try:
        with st.spinner("Carregando planilha origem..."):
            loaded_origin = carregar_origem(gc, origin_id, origin_sheet)
        st.success("Planilha origem carregada.")
    except Exception as e:
        st.error(f"Falha ao carregar origem: {e}")
        st.stop()

    # Para cada planilha selecionada, obter abas e detectar grupo
    st.markdown("## Passo 2 — Configurar cada planilha selecionada")
    planilhas_config = {}
    for pid in selected_ids:
        try:
            sh = gc.open_by_key(pid)
            st.subheader(f"{sh.title}  —  {pid}")
            # detectar grupo via rel comp
            grupo_detectado, extra_detectado = detectar_grupo_por_relcomp(sh)
            col1, col2 = st.columns([2, 2])
            with col1:
                st.write(f"Grupo detectado (B4 de 'rel comp'): {grupo_detectado or '— não detectado —'}")
                grupo_override = st.text_input(f"Grupo a usar (se vazio usa detectado) — {pid}", value=grupo_detectado or "", key=f"grupo_override_{pid}")
            with col2:
                st.write(f"Filtro extra detectado (B6): {extra_detectado or '— não detectado —'}")
                extra_override = st.text_input(f"Filtro extra (se necessário) — {pid}", value=extra_detectado or "", key=f"extra_override_{pid}")

            # listar abas existentes e escolher aba destino
            abas = [ws.title for ws in sh.worksheets()]
            abas_display = abas + ["__CRIAR_NOVA_ABA__"]
            chosen_aba = st.selectbox(f"Escolha aba de destino para atualizar em {sh.title}", abas_display, key=f"dest_aba_{pid}")
            new_aba_name = ""
            if chosen_aba == "__CRIAR_NOVA_ABA__":
                new_aba_name = st.text_input(f"Nome da nova aba para criar em {sh.title}", value="Importado_Fat", key=f"nova_aba_{pid}")

            # calcular prévia: filtrar origem usando Grupo/extra/data_minima
            grupo_final = (grupo_override.strip().upper() if grupo_override and grupo_override.strip() else (grupo_detectado or "")).strip().upper()
            extra_final = extra_override.strip().upper() if extra_override and extra_override.strip() else (extra_detectado or "")
            df = loaded_origin.copy()
            mask = df["Grupo"].astype(str).str.upper() == grupo_final if grupo_final else pd.Series([True]*len(df), index=df.index)
            mask = mask & df["Data_dt"].notna() & (df["Data_dt"].dt.date >= data_minima)
            df_filtrado = df.loc[mask].copy()
            st.write(f"Linhas que seriam enviadas para essa planilha: {len(df_filtrado)}")
            if len(df_filtrado) > 0:
                with st.expander("Ver amostra (10 linhas)"):
                    st.dataframe(df_filtrado.head(10).drop(columns=["Data_dt"], errors="ignore"), use_container_width=True)

            planilhas_config[pid] = {
                "spreadsheet": sh,
                "grupo": grupo_final,
                "extra": extra_final,
                "dest_aba": new_aba_name.strip() if chosen_aba == "__CRIAR_NOVA_ABA__" else chosen_aba,
                "df_preview": df_filtrado
            }
        except Exception as e:
            st.error(f"Erro abrindo planilha {pid}: {e}")

    # Botão para executar atualização (somente se existir ao menos 1 selecionada)
    if planilhas_config:
        st.markdown("---")
        confirmar = st.checkbox("Eu confirmo que desejo enviar os dados selecionados para as planilhas/abas escolhidas", key="confirm_update")
        if st.button("Executar Atualização Agora") and confirmar:
            resultados = []
            with st.spinner("Enviando atualizações..."):
                for pid, conf in planilhas_config.items():
                    sh = conf["spreadsheet"]
                    df_send = conf["df_preview"]
                    dest_aba = conf["dest_aba"] or "Importado_Fat"
                    try:
                        if df_send is None or df_send.empty:
                            resultados.append((pid, sh.title, 0, "Sem linhas para enviar"))
                            continue
                        # garante colunas (remove coluna Data_dt auxiliar)
                        if "Data_dt" in df_send.columns:
                            df_send = df_send.drop(columns=["Data_dt"])
                        # criar aba se não existir
                        try:
                            ws_dest = sh.worksheet(dest_aba)
                        except gspread.exceptions.WorksheetNotFound:
                            ws_dest = sh.add_worksheet(title=dest_aba, rows=str(max(1000, len(df_send)+10)), cols=str(len(df_send.columns)))
                        # limpar e enviar
                        ws_dest.clear()
                        valores = [df_send.columns.tolist()] + df_send.fillna("").astype(str).values.tolist()
                        ws_dest.update("A1", valores, value_input_option="USER_ENTERED")
                        resultados.append((pid, sh.title, len(df_send), "OK"))
                    except Exception as e:
                        resultados.append((pid, sh.title, 0, f"ERRO: {e}"))

            # Exibir resumo
            st.success("Processo concluído. Resumo:")
            df_res = pd.DataFrame(resultados, columns=["ID", "Nome", "Linhas Enviadas", "Status"])
            st.dataframe(df_res, use_container_width=True)
            # limpa seleção se quiser
            st.session_state.candidatas = st.session_state.candidatas  # mantém lista
    else:
        st.info("Nenhuma configuração de planilhas pronta para atualização. Selecione ao menos uma planilha acima.")
