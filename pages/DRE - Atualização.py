import streamlit as st
import pandas as pd
import json
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
import gspread

try:
    from googleapiclient.discovery import build
    from googleapiclient.errors import HttpError
except Exception:
    build = None
    HttpError = Exception

st.set_page_config(page_title="Atualização e Auditoria Profissional", layout="wide")

# CSS para estilizar caixas e títulos
st.markdown("""
<style>
.card {
    background-color: #f9f9f9;
    padding: 15px;
    border-radius: 8px;
    box-shadow: 0 2px 5px rgb(0 0 0 / 0.1);
    margin-bottom: 20px;
}
h2 {
    color: #2c3e50;
}
</style>
""", unsafe_allow_html=True)

st.title("📊 Atualização e Auditoria — Faturamento x Meio de Pagamento")

# Autenticação
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

# Sidebar com configurações
with st.sidebar:
    st.header("⚙️ Configurações")
    origin_id = st.text_input("ID da planilha origem", value="1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU")
    origin_sheet = st.text_input("Aba origem", value="Fat Sistema Externo")
    data_minima = st.date_input("Data mínima", value=datetime.now() - timedelta(days=365))
    folder_ids_text = st.text_area("IDs das pastas (uma por linha)", height=100)
    folder_ids = [x.strip() for x in folder_ids_text.splitlines() if x.strip()]
    manual_ids_text = st.text_area("IDs ou URLs de planilhas manuais (uma por linha)", height=100)
    manual_ids = []
    for line in manual_ids_text.splitlines():
        line = line.strip()
        if not line:
            continue
        if "docs.google.com/spreadsheets" in line:
            parts = line.split("/d/")
            if len(parts) > 1:
                manual_ids.append(parts[1].split("/")[0])
            else:
                manual_ids.append(line)
        else:
            manual_ids.append(line)
    listar_btn = st.button("🔍 Listar planilhas candidatas")

# Função para listar arquivos na pasta
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

# Listar planilhas candidatas
candidatas = []

if listar_btn:
    with st.spinner("Buscando planilhas nas pastas..."):
        if drive_service and folder_ids:
            for fid in folder_ids:
                arquivos = listar_arquivos_pasta(drive_service, fid)
                for a in arquivos:
                    candidatas.append({"id": a["id"], "name": a["name"], "folder_id": fid})
        # Adiciona manuais
        for mid in manual_ids:
            try:
                sh = gc.open_by_key(mid)
                candidatas.append({"id": mid, "name": sh.title, "folder_id": None})
            except Exception as e:
                st.warning(f"Falha abrindo planilha manual '{mid}': {e}")

    if not candidatas:
        st.warning("Nenhuma planilha encontrada.")
    else:
        st.success(f"{len(candidatas)} planilhas encontradas.")

# Exibir resultados em uma caixa estilizada
if candidatas:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Planilhas candidatas")
    df_cand = pd.DataFrame(candidatas)
    df_cand_display = df_cand[["name", "id", "folder_id"]].rename(columns={"name": "Nome", "id": "ID", "folder_id": "Pasta ID"})
    st.dataframe(df_cand_display, use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)

# Aqui você pode continuar com a lógica de seleção, pré-visualização e atualização, usando colunas e expanders para organizar melhor.

# Exemplo simples de seleção
if candidatas:
    st.subheader("Selecione as planilhas para atualizar")
    options = [f"{c['name']} ({c['id']})" for c in candidatas]
    selecionadas = st.multiselect("Planilhas", options, default=options)

    if selecionadas:
        st.info(f"Você selecionou {len(selecionadas)} planilhas para atualizar.")
        # Aqui você pode adicionar botões para avançar, mostrar prévias, etc.

# Rodar o app com:
# streamlit run seu_arquivo.py
