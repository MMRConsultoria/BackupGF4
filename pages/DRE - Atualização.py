import streamlit as st
import pandas as pd
import json
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# Opcional: Drive API para listar planilhas em pastas
try:
    from googleapiclient.discovery import build
except Exception:
    build = None

# ---------- CONFIG ----------
DEFAULT_FOLDER_IDS = [
    # coloque aqui as folder IDs que você usa para listar planilhas (opcional)
    # "PASTA_ID_1",
    # "PASTA_ID_2",
]
st.set_page_config(page_title="Seleção — Tabela Simples", layout="wide")
st.title("Tabela simples — marcar operações por planilha")

OPERACOES = ["Desconto", "Meio Pagamento", "Faturamento"]

# ---------- AUTENTICAÇÃO ----------
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

try:
    gc, drive_service = autenticar_gspread()
except Exception as e:
    st.error("Erro na autenticação com Google. Verifique st.secrets['GOOGLE_SERVICE_ACCOUNT'].")
    st.stop()

# ---------- FUNÇÃO DE LISTAGEM ----------
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

# ---------- CARREGA LISTA DE PLANILHAS ----------
planilhas = []
# tenta listar a partir de pastas (se configuradas e Drive API disponível)
if drive_service and DEFAULT_FOLDER_IDS:
    for fid in DEFAULT_FOLDER_IDS:
        try:
            arquivos = listar_arquivos_pasta(drive_service, fid)
            for a in arquivos:
                planilhas.append({"id": a["id"], "name": a["name"]})
        except Exception:
            pass

# campo para colar IDs manuais caso a listagem automática não encontre tudo
st.markdown("Se a listagem automática não encontrar suas planilhas, cole IDs (um por linha) abaixo e pressione 'Adicionar IDs'.")
manual_ids_text = st.text_area("IDs manuais (opcional)", height=120, placeholder="cole IDs de planilhas aqui, 1 por linha")
if st.button("Adicionar IDs"):
    for line in manual_ids_text.splitlines():
        line = line.strip()
        if not line:
            continue
        try:
            sh = gc.open_by_key(line)
            if not any(p["id"] == sh.id for p in planilhas):
                planilhas.append({"id": sh.id, "name": sh.title})
        except Exception as e:
            st.warning(f"Não foi possível abrir ID '{line}': {e}")

if not planilhas:
    st.info("Nenhuma planilha encontrada. Cole IDs manuais ou configure DEFAULT_FOLDER_IDS.")
else:
    st.markdown(f"Encontradas {len(planilhas)} planilhas.")

# ---------- FORMULÁRIO: tabela simples com checkboxes ----------
st.markdown("### Seleções (tabela simples)")
with st.form("tabela_form"):
    # cabeçalho
    header_cols = st.columns([0.6, 0.13, 0.13, 0.13])
    header_cols[0].markdown("**Planilha**")
    header_cols[1].markdown("**Desconto**")
    header_cols[2].markdown("**Meio Pagamento**")
    header_cols[3].markdown("**Faturamento**")

    # para cada planilha, renderiza uma "linha" com nome e 3 checkboxes (todos True por padrão)
    for p in planilhas:
        pid = p["id"]
        name = p["name"]
        row_cols = st.columns([0.6, 0.13, 0.13, 0.13])
        row_cols[0].write(name)
        # keys estáveis baseadas no id da planilha + operação
        key_des = f"{pid}__Desconto"
        key_mp = f"{pid}__Meio_Pagamento"
        key_fat = f"{pid}__Faturamento"
        row_cols[1].checkbox("", value=True, key=key_des)
        row_cols[2].checkbox("", value=True, key=key_mp)
        row_cols[3].checkbox("", value=True, key=key_fat)

    submitted = st.form_submit_button("Enviar seleções")

# ---------- APÓS SUBMIT: mostrar resumo ----------
if 'submitted' in locals() and submitted:
    resultados = []
    for p in planilhas:
        pid = p["id"]
        name = p["name"]
        des = st.session_state.get(f"{pid}__Desconto", False)
        mp = st.session_state.get(f"{pid}__Meio_Pagamento", False)
        fat = st.session_state.get(f"{pid}__Faturamento", False)
        resultados.append({
            "Planilha": name,
            "ID": pid,
            "Desconto": des,
            "Meio Pagamento": mp,
            "Faturamento": fat
        })
    df_res = pd.DataFrame(resultados)
    st.markdown("### Resumo das seleções")
    st.dataframe(df_res, use_container_width=True)
    # aqui você pode prosseguir para executar as ações conforme as seleções,
    # por enquanto só mostramos o resumo (dry-run).
    st.success("Seleções capturadas (dry-run). Se quiser que eu adicione execução, me diga o que deseja que aconteça ao confirmar.")
