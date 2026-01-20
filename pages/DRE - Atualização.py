import streamlit as st
import pandas as pd
import numpy as np
import json
import time
from io import BytesIO
from datetime import datetime, timedelta, date
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# try import optional formatting helpers (não obrigatório)
try:
    from gspread_formatting import format_cell_range, CellFormat, NumberFormat
except Exception:
    format_cell_range = None
    CellFormat = None
    NumberFormat = None

# ---------------- Config/layout (modelo) ----------------
st.set_page_config(page_title="Atualizador DRE", layout="wide")

# 🔥 CSS para estilizar as abas e reduzir espaçamento (do seu modelo + ajustes)
st.markdown(
    """
    <style>
    .stApp { background-color: #f9f9f9; }
    div[data-baseweb="tab-list"] { margin-top: 20px; }
    button[data-baseweb="tab"] {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 10px 20px;
        margin-right: 10px;
        transition: all 0.3s ease;
        font-size: 16px;
        font-weight: 600;
    }
    button[data-baseweb="tab"]:hover { background-color: #dce0ea; color: black; }
    button[data-baseweb="tab"][aria-selected="true"] { background-color: #0366d6; color: white; }

    /* Oculta toolbar do Streamlit */
    [data-testid="stToolbar"] { visibility: hidden; height: 0%; position: fixed; }

    /* Espaçamento geral */
    .block-container { padding-top: 1.5rem; padding-bottom: 0.5rem; }
    div.stVerticalBlock > div { margin-bottom: 0.0rem; }
    h1 { margin-top: -1rem; margin-bottom: 0.5rem; }

    /* Tabela: reduzir padding célula */
    [data-testid="stTable"] td, [data-testid="stTable"] th { padding: 6px 8px !important; }

    /* Esconde rótulos vazios de widgets */
    label[data-testid="stWidgetLabel"] { min-height: 0px; margin-bottom: 0px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# 🔒 Bloqueio — mantenha seu fluxo de autenticação separado; o app para aqui se não autorizado
if not st.session_state.get("acesso_liberado", True):
    # Se você usa controle de acesso real, defina st.session_state["acesso_liberado"] = True após login
    st.error("Acesso não liberado. Contate o administrador.")
    st.stop()

NOME_SISTEMA = "Colibri"

# ---------------- Helpers ----------------
@st.cache_data(ttl=300)
def list_subfolders(_drive, parent_id):
    folders = []
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

@st.cache_data(ttl=300)
def list_spreadsheets_in_folders(_drive, folder_ids):
    sheets = []
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

# ---------------- AUTENTICAÇÃO Google Sheets ----------------
@st.cache_resource
def autenticar():
    scope = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
    creds_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    gc = gspread.authorize(creds)
    drive = None
    try:
        from googleapiclient.discovery import build
        drive = build("drive", "v3", credentials=creds)
    except Exception:
        drive = None
    return gc, drive

try:
    gc, drive_service = autenticar()
except Exception as e:
    st.error(f"Erro de autenticação: {e}")
    st.stop()

# ---------------- Abas do app ----------------
tab_atualizador, tab_auditoria = st.tabs(["Atualizador", "Auditoria (em desenvolvimento)"])

# ---------------- Aba: Atualizador ----------------
with tab_atualizador:
    st.header("Atualizador DRE")

    # Período
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
        # Carrega subpastas (pasta principal configurada no topo)
        try:
            subfolders = list_subfolders(drive_service, MAIN_FOLDER_ID)
        except Exception as e:
            st.error(f"Erro listando subpastas: {e}")
            subfolders = []

        if not subfolders:
            st.warning("Nenhuma subpasta encontrada dentro da pasta principal. Verifique permissões.")
        else:
            sub_names = [f"{s['name']} ({s['id']})" for s in subfolders]
            # Rótulo vazio para não mostrar texto
            selected = st.multiselect("", options=sub_names, default=sub_names)
            selected_folder_ids = [s.split("(")[-1].strip(")") for s in selected if "(" in s]

            if not selected_folder_ids:
                st.info("Nenhuma subpasta selecionada. Selecione as subpastas para prosseguir.")
            else:
                # Buscar planilhas nas pastas selecionadas
                with st.spinner("Buscando planilhas nas pastas selecionadas..."):
                    try:
                        planilhas = list_spreadsheets_in_folders(drive_service, list(selected_folder_ids))
                    except Exception as e:
                        st.error(f"Erro ao listar planilhas: {e}")
                        planilhas = []

                if not planilhas:
                    st.warning("Nenhuma planilha encontrada nas subpastas selecionadas.")
                else:
                    # Preparar DataFrame
                    df = pd.DataFrame(planilhas)
                    df = df.rename(columns={"name": "Planilha", "id": "ID_Planilha", "parent_folder_id": "Folder_ID"})
                    df["Desconto"] = True
                    df["Meio Pagamento"] = True
                    df["Faturamento"] = True
                    df = df[["Planilha", "Folder_ID", "ID_Planilha", "Desconto", "Meio Pagamento", "Faturamento"]].sort_values("Planilha").reset_index(drop=True)

                    # Form para editar e submeter
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
                            hide_index=True
                        )
                        submit = st.form_submit_button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True)

                    if submit:
                        # Configurações internas fixas (pode tornar toggles visíveis se quiser)
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
                            st.write(f"Iniciando processamento de **{len(tarefas)}** tarefas (período: {data_de_str} → {data_ate_str})")
                            progresso = st.progress(0)
                            logs = []
                            for i, t in enumerate(tarefas):
                                st.info(f"{i+1}/{len(tarefas)} — {t['planilha']} -> {t['operacao']}")
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
                                        # Aqui deve entrar a lógica de leitura da origem (ID_PLANILHA_ORIGEM/ABA_ORIGEM),
                                        # filtragem por data e gravação na aba destino.
                                        # Por enquanto mantemos simulação:
                                        time.sleep(0.1)
                                        logs.append(f"{t['planilha']}/{t['operacao']}: gravaria dados para {data_de_str}→{data_ate_str} (simulado)")
                                    else:
                                        logs.append(f"{t['planilha']}/{t['operacao']}: dry-run (não gravado)")
                                except Exception as e:
                                    logs.append(f"{t['planilha']}/{t['operacao']}: ERRO -> {e}")
                                progresso.progress((i + 1) / len(tarefas))

                            st.success("Processamento finalizado.")
                            if logs:
                                st.write("Logs:")
                                st.write("\n".join(logs))

# ---------------- Aba: Auditoria (esqueleto) ----------------
with tab_auditoria:
    st.header("Auditoria (em desenvolvimento)")
    st.write("Área de auditoria em construção. Use este espaço para:")
    st.write("- Fazer uploads de arquivos para auditoria")
    st.write("- Visualizar logs detalhados")
    st.write("- Rodar verificações automatizadas")
    st.info("Funcionalidades planejadas: comparação de dados entre planilhas, validação de tipos, registros de divergência e relatório de auditoria exportável.")

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
