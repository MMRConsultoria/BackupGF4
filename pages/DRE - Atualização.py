import streamlit as st
import pandas as pd
import json
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# opcional: Drive API para listar planilhas em pastas
try:
    from googleapiclient.discovery import build
except Exception:
    build = None

# ---------------- CONFIG ----------------
DEFAULT_FOLDER_IDS = [
    # coloque aqui as folder IDs que contêm suas planilhas (deixe vazio se não usar)
    "1ptFvtxYjISfB19S7bU9olMLmAxDTBkOh",
]

st.set_page_config(page_title="Tabela com Checkboxes", layout="wide")
st.title("Planilhas — tabela com checkboxes (Desconto / Meio Pagamento / Faturamento)")

OPERACOES = ["Desconto", "Meio Pagamento", "Faturamento"]

# ---------------- Autenticação gspread (usar st.secrets['GOOGLE_SERVICE_ACCOUNT']) ----------------
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

# ---------------- Função para listar planilhas em pastas (Drive API) ----------------
def listar_arquivos_pasta(drive_service, pasta_id):
    arquivos = []
    if not drive_service or not pasta_id:
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

# ---------------- Carregar lista de planilhas automaticamente ----------------
planilhas = []
if drive_service and DEFAULT_FOLDER_IDS:
    for fid in DEFAULT_FOLDER_IDS:
        try:
            files = listar_arquivos_pasta(drive_service, fid)
            for f in files:
                planilhas.append({"Name": f["name"], "ID": f["id"]})
        except Exception as e:
            st.warning(f"Falha listando pasta {fid}: {e}")
else:
    if not DEFAULT_FOLDER_IDS:
        st.info("Nenhuma DEFAULT_FOLDER_IDS configurada — configure as pastas para listagem automática.")
    elif not drive_service:
        st.warning("Drive API indisponível (não foi possível inicializar googleapiclient).")

if not planilhas:
    st.info("Nenhuma planilha encontrada automaticamente. Se quiser, me forneça as folder IDs ou eu adapto para outra estratégia.")
    st.stop()

# ---------------- Montar DataFrame com colunas de seleção (todos True por padrão) ----------------
df = pd.DataFrame(planilhas)
# pré-cria colunas de seleção com True por padrão
df["Desconto"] = True
df["Meio Pagamento"] = True
df["Faturamento"] = True

# Reordenar colunas para ficar parecido com sua imagem (nome -> checkboxes -> id escondido)
display_cols = ["Name", "Desconto", "Meio Pagamento", "Faturamento", "ID"]
df = df[display_cols]

st.markdown("Edite as caixas na tabela abaixo e depois clique em 'Enviar seleções'. As alterações só serão aplicadas ao clicar no botão.")

# ---------------- Usar st.data_editor dentro de um form (evita reexecuções enquanto usuário edita) ----------------
if not hasattr(st, "data_editor"):
    st.error("Seu Streamlit não tem `st.data_editor`. Atualize o Streamlit: pip install --upgrade streamlit")
    st.stop()

with st.form("selection_form"):
    edited = st.data_editor(
        df,
        num_rows="fixed",
        use_container_width=True
    )
    submitted = st.form_submit_button("Enviar seleções")

# ---------------- Após submit: processar as linhas selecionadas ----------------
if submitted:
    # edited é um DataFrame com alterações do usuário
    sel_rows = edited[
        (edited["Desconto"] == True) |
        (edited["Meio Pagamento"] == True) |
        (edited["Faturamento"] == True)
    ].copy()

    if sel_rows.empty:
        st.warning("Nenhuma operação marcada em nenhuma planilha.")
    else:
        # mostra resumo parecido com sua tabela
        st.markdown("### Resumo das seleções")
        st.dataframe(sel_rows.reset_index(drop=True), use_container_width=True)

        # Exemplo: montar lista de tarefas a executar (apenas visual; não grava nada ainda)
        tarefas = []
        for _, row in sel_rows.iterrows():
            pid = row["ID"]
            nome = row["Name"]
            if row["Desconto"]:
                tarefas.append((nome, pid, "Desconto"))
            if row["Meio Pagamento"]:
                tarefas.append((nome, pid, "Meio Pagamento"))
            if row["Faturamento"]:
                tarefas.append((nome, pid, "Faturamento"))

        st.markdown(f"Total (planilha × operação) selecionados: **{len(tarefas)}**")
        # por enquanto exibimos as tarefas; se quiser eu adiciono a execução (escrita nas abas)
        df_tarefas = pd.DataFrame(tarefas, columns=["Planilha", "ID", "Operação"])
        st.dataframe(df_tarefas, use_container_width=True)
        st.success("Seleções prontas — diga se você quer que eu adicione o passo de gravação nas planilhas (com backup / dry-run).")
