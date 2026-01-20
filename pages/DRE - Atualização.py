import streamlit as st
import pandas as pd
import json
import time
from datetime import datetime, timedelta, date
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# Drive API (opcional)
try:
    from googleapiclient.discovery import build
except Exception:
    build = None

# ---------------- CONFIG ----------------
MAIN_FOLDER_ID = "1LrbcStUAcvZV_dOYKBt-vgBHb9e1d6X-"  # defina sua pasta raiz
ID_PLANILHA_ORIGEM = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
ABA_ORIGEM = "Fat Sistema Externo"
MAPA_ABAS = {
    "Faturamento": "Importado Fat",
    "Meio Pagamento": "Meio Pagamento",
    "Desconto": "Desconto",
}

st.set_page_config(page_title="Atualizador DRE", layout="wide")

# --- CSS PARA REDUZIR ESPAÇAMENTO ---
st.markdown(
    """
    <style>
    .block-container { padding-top: 3rem; padding-bottom: 0rem; }
    div.stVerticalBlock > div { margin-bottom: -0.1rem; }
    h1 { margin-top: -1rem; margin-bottom: 0.5rem; }
    label[data-testid="stWidgetLabel"] { min-height: 0px; margin-bottom: 0px; }
    [data-testid="stTable"] td, [data-testid="stTable"] th { padding: 4px 8px !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Atualizador DRE")

# ---------------- AUTENTICAÇÃO ----------------
@st.cache_resource
def autenticar():
    scope = [
        "https://www.googleapis.com/auth/drive",
        "https://www.googleapis.com/auth/spreadsheets",
    ]
    creds_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    gc = gspread.authorize(creds)
    drive = None
    if build:
        try:
            drive = build("drive", "v3", credentials=creds)
        except Exception:
            drive = None
    return gc, drive

# BOTÃO PARA FORÇAR ATUALIZAÇÃO (limpa caches e rerun)
if st.button("🔄 Forçar Atualização do Drive (limpa cache)"):
    # tenta limpar os caches registrados
    try:
        st.cache_data.clear()
    except Exception:
        pass
    try:
        st.cache_resource.clear()
    except Exception:
        pass
    st.experimental_rerun()

# Tenta autenticar (cached)
try:
    gc, drive_service = autenticar()
except Exception as e:
    st.error(f"Erro de autenticação: {e}")
    gc, drive_service = None, None

# Se o drive_service não existir, mostramos aviso, mas NÃO paramos a app inteira
if not drive_service:
    st.warning("Drive API não inicializada. Algumas funcionalidades (listagem automática) podem não funcionar. Verifique dependências/permissões.")
    # Não usamos st.stop() para permitir interação com a UI (ex.: upload manual)

# ---------------- HELPERS DRIVE (SEM CACHE PARA REFLEXÃO INSTANTÂNEA) ----------------
def list_subfolders(_drive, parent_id):
    """Sempre consulta o Drive — sem cache — para refletir mudanças imediatamente."""
    folders = []
    if not _drive:
        return folders
    page_token = None
    q = f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    while True:
        resp = _drive.files().list(q=q, fields="nextPageToken, files(id, name)", pageToken=page_token).execute()
        files = resp.get("files", [])
        for f in files:
            folders.append({"id": f["id"], "name": f["name"]})
        page_token = resp.get("nextPageToken", None)
        if not page_token:
            break
    return folders


def list_spreadsheets_in_folders(_drive, folder_ids):
    """Sempre consulta o Drive — sem cache — para refletir mudanças imediatamente."""
    sheets = []
    if not _drive:
        return sheets
    for fid in folder_ids:
        page_token = None
        q = f"'{fid}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
        while True:
            resp = _drive.files().list(q=q, fields="nextPageToken, files(id, name)", pageToken=page_token).execute()
            files = resp.get("files", [])
            for f in files:
                sheets.append({"id": f["id"], "name": f["name"], "parent_folder_id": fid})
            page_token = resp.get("nextPageToken", None)
            if not page_token:
                break
    # remover duplicatas por id, mantendo primeiro encontro
    seen = set()
    unique = []
    for s in sheets:
        if s["id"] not in seen:
            seen.add(s["id"])
            unique.append(s)
    return unique


# ---------------- Layout com Abas ----------------
tab_atualizador, tab_auditoria = st.tabs(["Atualizador", "Auditoria (em desenvolvimento)"])

with tab_atualizador:
    # ---------------- UI: PASSO 0 - PERÍODO ----------------
    col_start, col_end = st.columns(2)
    default_end = date.today()
    default_start = default_end - timedelta(days=30)

    with col_start:
        data_de = st.date_input("De (dd/mm/aaaa)", value=default_start)
    with col_end:
        data_ate = st.date_input("Até (dd/mm/aaaa)", value=default_end)

    data_de_str = data_de.strftime("%d/%m/%Y")
    data_ate_str = data_ate.strftime("%d/%m/%Y")

    if data_ate < data_de:
        st.error("Data 'Até' deve ser posterior à 'De'.")
    else:
        # ---------------- LISTAR E SELECIONAR SUBPASTAS ----------------
        try:
            subfolders = list_subfolders(drive_service, MAIN_FOLDER_ID)
        except Exception as e:
            st.error(f"Erro listando subpastas: {e}")
            subfolders = []

        if not subfolders:
            st.warning("Nenhuma subpasta encontrada dentro da pasta principal. Verifique permissões ou clique em '🔄 Forçar Atualização do Drive' após ajustar no Drive.")
        else:
            sub_names = [f"{s['name']} ({s['id']})" for s in subfolders]
            selected = st.multiselect("", options=sub_names, default=sub_names)
            selected_folder_ids = [s.split("(")[-1].strip(")") for s in selected if "(" in s]

            if not selected_folder_ids:
                st.info("Nenhuma subpasta selecionada. Selecione para prosseguir.")
            else:
                # ---------------- BUSCAR PLANILHAS ----------------
                with st.spinner("Buscando planilhas..."):
                    try:
                        planilhas = list_spreadsheets_in_folders(drive_service, list(selected_folder_ids))
                    except Exception as e:
                        st.error(f"Erro ao listar planilhas: {e}")
                        planilhas = []

                if not planilhas:
                    st.warning("Nenhuma planilha encontrada.")
                else:
                    df = pd.DataFrame(planilhas)
                    df = df.rename(columns={"name": "Planilha", "id": "ID_Planilha", "parent_folder_id": "Folder_ID"})
                    df["Desconto"] = True
                    df["Meio Pagamento"] = True
                    df["Faturamento"] = True
                    df = df[["Planilha", "Folder_ID", "ID_Planilha", "Desconto", "Meio Pagamento", "Faturamento"]].sort_values("Planilha").reset_index(drop=True)

                    with st.form("selection_form"):
                        edited = st.data_editor(
                            df,
                            num_rows="fixed",
                            use_container_width=True,
                            column_config={
                                "Planilha": st.column_config.TextColumn("Planilha", disabled=True, width="large"),
                                "Folder_ID": None,
                                "ID_Planilha": None,
                                "Desconto": st.column_config.CheckboxColumn("Desconto", default=True),
                                "Meio Pagamento": st.column_config.CheckboxColumn("Meio Pagamento", default=True),
                                "Faturamento": st.column_config.CheckboxColumn("Faturamento", default=True),
                            },
                            hide_index=True,
                        )
                        submit = st.form_submit_button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True)

                    if submit:
                        DRY_RUN = True
                        DO_BACKUP = True

                        tarefas = []
                        for _, row in edited.iterrows():
                            if row["Desconto"]:
                                tarefas.append({"planilha": row["Planilha"], "id": row["ID_Planilha"], "operacao": "Desconto", "aba": MAPA_ABAS["Desconto"]})
                            if row["Meio Pagamento"]:
                                tarefas.append({"planilha": row["Planilha"], "id": row["ID_Planilha"], "operacao": "Meio Pagamento", "aba": MAPA_ABAS["Meio Pagamento"]})
                            if row["Faturamento"]:
                                tarefas.append({"planilha": row["Planilha"], "id": row["ID_Planilha"], "operacao": "Faturamento", "aba": MAPA_ABAS["Faturamento"]})

                        if not tarefas:
                            st.warning("Nenhuma operação selecionada.")
                        else:
                            st.write(f"Processando **{len(tarefas)}** tarefas...")
                            progresso = st.progress(0)
                            logs = []
                            for i, t in enumerate(tarefas):
                                try:
                                    if not DRY_RUN:
                                        sh = gc.open_by_key(t["id"])
                                        if DO_BACKUP:
                                            try:
                                                ws = sh.worksheet(t["aba"])
                                                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                                                sh.duplicate_sheet(ws.id, new_sheet_name=f"BACKUP_{t['aba']}_{timestamp}")
                                                logs.append(f"{t['planilha']}/{t['aba']}: backup criado")
                                            except Exception as e:
                                                logs.append(f"{t['planilha']}/{t['aba']}: backup falhou ou aba não existe -> {e}")
                                    else:
                                        logs.append(f"{t['planilha']}/{t['operacao']}: dry-run (não gravado)")

                                    time.sleep(0.1)
                                    progresso.progress((i + 1) / len(tarefas))
                                except Exception as e:
                                    logs.append(f"{t['planilha']}: ERRO -> {e}")
                                    st.error(f"Erro em {t['planilha']}: {e}")
                                    progresso.progress((i + 1) / len(tarefas))

                            st.success("Concluído!")
                            if logs:
                                st.write("Logs:")
                                st.write("\n".join(logs))

with tab_auditoria:
    st.header("Auditoria (em desenvolvimento)")
    st.write("Área de auditoria em construção. Use este espaço para:")
    st.write("- Fazer uploads de arquivos para auditoria")
    st.write("- Visualizar logs detalhados")
    st.write("- Rodar verificações automatizadas")
    uploaded = st.file_uploader("Enviar arquivo de auditoria (opcional)", type=["csv", "xlsx", "txt"])
    if uploaded is not None:
        try:
            if uploaded.type == "text/csv" or uploaded.name.lower().endswith(".csv"):
                df_audit = pd.read_csv(uploaded)
            else:
                df_audit = pd.read_excel(uploaded)
            st.write("Pré-visualização do arquivo enviado:")
            st.dataframe(df_audit.head(200))
        except Exception as e:
            st.error(f"Erro ao ler o arquivo: {e}")

    st.write("---")
    st.write("Notas rápidas para desenvolvedor:")
    st.code(
        """# Aqui você pode implementar:
# - funções de comparação entre df_relatorio_base e df_audit
# - rotinas para marcar divergências e exportar CSVs de relatório
# - painel de filtros (data, loja, tipo) para reproduzir problemas
"""
    )
