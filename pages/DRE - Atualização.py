import streamlit as st
import pandas as pd
import json
from datetime import datetime, timedelta
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import time

# opcional: Drive API para listar arquivos em pastas
try:
    from googleapiclient.discovery import build
except Exception:
    build = None

# -----------------------
# CONFIGURAÇÃO — ajuste conforme necessário
# -----------------------
DEFAULT_FOLDER_IDS = [
    "1ptFvtxYjISfB19S7bU9olMLmAxDTBkOh",
    "1F2Py4eeoqxqrHptgoeUODNXDCUddoU1u",
]
DEFAULT_ORIGIN_SPREADSHEET = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
DEFAULT_ORIGIN_SHEET = "Fat Sistema Externo"
DEFAULT_DATA_MINIMA = (datetime.now() - timedelta(days=365)).date()

OPERACOES = ["Desconto", "Meio Pagamento", "Faturamento"]
ABA_MAP = {"Faturamento": "Importado Fat", "Meio Pagamento": "Meio Pagamento", "Desconto": "Desconto"}

st.set_page_config(page_title="Atualização por Operação", layout="wide")
st.title("Atualização de Planilhas por Operação — Fluxo Estável")

# -----------------------
# AUTENTICAÇÃO (gspread + Drive opcional)
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

try:
    gc, drive_service = autenticar_gspread()
except Exception as e:
    st.error("Erro na autenticação com Google. Verifique st.secrets['GOOGLE_SERVICE_ACCOUNT'].")
    st.stop()

# -----------------------
# FUNÇÕES AUXILIARES
# -----------------------
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

def detectar_grupo_relcomp(sh):
    try:
        aba_rel = next((a for a in sh.worksheets() if "rel comp" in a.title.lower()), None)
        if not aba_rel:
            return None
        v = aba_rel.acell("B4").value
        return (v or "").strip().upper()
    except Exception:
        return None

def backup_worksheet(sh, ws_title):
    try:
        ws = sh.worksheet(ws_title)
    except Exception:
        return None, "Worksheet not found"
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"{ws_title}_backup_{ts}"
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
# CARREGA LISTA DE PLANILHAS (a partir das pastas configuradas)
# -----------------------
planilhas = []
if drive_service and DEFAULT_FOLDER_IDS:
    for fid in DEFAULT_FOLDER_IDS:
        try:
            arquivos = listar_arquivos_pasta(drive_service, fid)
            for a in arquivos:
                planilhas.append({"id": a["id"], "name": a["name"]})
        except Exception as e:
            st.warning(f"Não foi possível listar pasta {fid}: {e}")
else:
    if not drive_service:
        st.warning("Drive API não disponível — verifique dependências.")
if not planilhas:
    st.info("Nenhuma planilha encontrada automaticamente. Você pode colar IDs manualmente abaixo.")

# -----------------------
# PASSO 1: seleção de planilhas (multiselect)
# -----------------------
st.markdown("### 1) Escolha as planilhas que deseja atualizar")

names = [p["name"] for p in planilhas]
# por usabilidade, selecionamos nenhum por padrão; usuário escolhe
sel_names = st.multiselect("Planilhas (selecione uma ou mais)", options=names, default=[])

st.markdown("IDs manuais (opcional). Cole um ID por linha para abrir planilhas diretamente.")
manual_ids = st.text_area("IDs de planilha (opcional)", height=80, placeholder="Cole IDs de planilha do Google Sheets, 1 por linha")

# integrar manual IDs à lista (tentativa de abrir)
if manual_ids.strip():
    for line in manual_ids.splitlines():
        line = line.strip()
        if not line:
            continue
        # tenta abrir por id
        try:
            sh = gc.open_by_key(line)
            # adiciona à lista se não existir
            if sh.title not in names:
                planilhas.append({"id": sh.id, "name": sh.title})
                names.append(sh.title)
            if sh.title not in sel_names:
                sel_names.append(sh.title)
        except Exception as e:
            st.warning(f"Não foi possível abrir ID '{line}': {e}")

# reconstruir mapeamento nome->id
name_to_id = {p["name"]: p["id"] for p in planilhas}

if not sel_names:
    st.info("Selecione ao menos uma planilha (ou cole IDs manuais) para prosseguir.")
else:
    st.markdown("### 2) Como deseja escolher operações para as planilhas?")
    escolha_modo = st.radio("Modo", options=["Mesmas operações para todas (recomendado)", "Escolher operações por planilha (aviso: muitos widgets)"])

    # -----------------------
    # FORMULÁRIO PRINCIPAL — agrupa inputs para evitar renderizações repetidas
    # -----------------------
    with st.form("main_form"):
        ops_global = []
        ops_por_planilha = {}
        if escolha_modo == "Mesmas operações para todas (recomendado)":
            ops_global = st.multiselect("Escolha as operações que serão aplicadas a todas as planilhas selecionadas", OPERACOES, default=[])
            st.caption("As operações selecionadas serão aplicadas a todas as planilhas escolhidas na etapa 1.")
        else:
            st.warning("Você escolheu selecionar por planilha. Se houver muitas planilhas selecionadas, a interface terá vários widgets.")
            for nm in sel_names:
                key = f"ops__{nm}"
                escolha = st.multiselect(f"{nm}", options=OPERACOES, default=[])
                ops_por_planilha[nm] = escolha

        data_min = st.date_input("Data mínima (incluir)", value=DEFAULT_DATA_MINIMA)
        dry_run = st.checkbox("Dry-run (não grava)", value=True)
        do_backup = st.checkbox("Fazer backup antes de sobrescrever", value=True)

        submitted = st.form_submit_button("Enviar / Atualizar")

    # -----------------------
    # PROCESSAMENTO APÓS SUBMIT
    # -----------------------
    if submitted:
        # montar tarefas: nome -> lista de operações
        tarefas = {}
        if escolha_modo == "Mesmas operações para todas (recomendado)":
            if not ops_global:
                st.warning("Nenhuma operação selecionada. Marque ao menos uma operação ou escolha modo por planilha.")
            else:
                for nm in sel_names:
                    tarefas[nm] = ops_global.copy()
        else:
            for nm in sel_names:
                ops = ops_por_planilha.get(nm, [])
                if ops:
                    tarefas[nm] = ops

        total_pairs = sum(len(v) for v in tarefas.values())
        st.write(f"Total de (planilha × operação) selecionados: **{total_pairs}**")

        if total_pairs == 0:
            st.info("Nenhuma operação marcada. Selecione operações antes de executar.")
        else:
            # carregar origem (somente uma vez)
            with st.spinner("Carregando planilha origem..."):
                try:
                    df_origem = carregar_origem(gc, DEFAULT_ORIGIN_SPREADSHEET, DEFAULT_ORIGIN_SHEET)
                except Exception as e:
                    st.error(f"Falha ao carregar origem: {e}")
                    st.stop()
            st.success("Planilha origem carregada.")

            # confirmar e executar
            confirm = st.checkbox("Confirmo e desejo executar a operação", key="confirm_exec_final")
            if confirm and st.button("Executar agora"):
                resultados = []
                logs = []
                total_tasks = total_pairs
                progress = st.progress(0)
                i_task = 0

                for nm, ops in tarefas.items():
                    pid = name_to_id.get(nm)
                    if not pid:
                        logs.append(f"{nm}: ID não encontrado - pulando")
                        for op in ops:
                            resultados.append((nm, op, 0, "ERROR", "ID não encontrado"))
                            i_task += 1
                            progress.progress(int(i_task/total_tasks*100))
                        continue

                    # tenta abrir planilha
                    try:
                        sh = gc.open_by_key(pid)
                    except Exception as e:
                        logs.append(f"{nm}: erro abrindo planilha -> {e}")
                        for op in ops:
                            resultados.append((nm, op, 0, "ERROR", f"Erro ao abrir: {e}"))
                            i_task += 1
                            progress.progress(int(i_task/total_tasks*100))
                        continue

                    grupo_detectado = detectar_grupo_relcomp(sh)

                    # aplicar filtro de data e grupo
                    df = df_origem.copy()
                    if grupo_detectado:
                        mask = df["Grupo"].astype(str).str.upper() == grupo_detectado
                    else:
                        mask = pd.Series([True] * len(df), index=df.index)
                    mask = mask & df["Data_dt"].notna() & (df["Data_dt"].dt.date >= data_min)
                    df_preview = df.loc[mask].copy()

                    if df_preview.empty:
                        logs.append(f"{nm}: sem linhas após filtro (grupo/data)")
                        for op in ops:
                            resultados.append((nm, op, 0, "SKIP", "Sem linhas após filtro"))
                            i_task += 1
                            progress.progress(int(i_task/total_tasks*100))
                        continue

                    # para cada operação grava na aba mapeada
                    for op in ops:
                        i_task += 1
                        dest_aba = ABA_MAP.get(op, op)
                        try:
                            # verificar existência da aba
                            try:
                                ws_dest = sh.worksheet(dest_aba)
                                aba_existed = True
                            except gspread.exceptions.WorksheetNotFound:
                                ws_dest = None
                                aba_existed = False

                            # backup se solicitado
                            if do_backup and aba_existed:
                                bname, berr = backup_worksheet(sh, dest_aba)
                                if berr:
                                    logs.append(f"{nm}/{dest_aba}: falha backup -> {berr}")
                                else:
                                    logs.append(f"{nm}/{dest_aba}: backup criado -> {bname}")

                            # dry-run
                            if dry_run:
                                resultados.append((nm, op, len(df_preview), "DRY-RUN", "Não gravado (dry-run)"))
                                logs.append(f"{nm}/{dest_aba}: dry-run -> {len(df_preview)} linhas.")
                                progress.progress(int(i_task/total_tasks*100))
                                continue

                            # criar aba se não existir
                            if not aba_existed:
                                ws_dest = sh.add_worksheet(title=dest_aba, rows=str(max(1000, len(df_preview)+10)), cols=str(max(20, len(df_preview.columns))))
                                time.sleep(0.3)

                            # escrever dados
                            ws_dest.clear()
                            values = [df_preview.columns.tolist()] + df_preview.fillna("").astype(str).values.tolist()
                            ws_dest.update("A1", values, value_input_option="USER_ENTERED")
                            resultados.append((nm, op, len(df_preview), "OK", f"Gravado em '{dest_aba}'"))
                            logs.append(f"{nm}/{dest_aba}: {len(df_preview)} linhas gravadas.")
                        except Exception as e:
                            resultados.append((nm, op, 0, "ERROR", str(e)))
                            logs.append(f"{nm}/{dest_aba}: ERRO -> {e}")
                        progress.progress(int(i_task/total_tasks*100))

                progress.progress(100)
                st.success("Operação finalizada")
                df_res = pd.DataFrame(resultados, columns=["Planilha", "Operação", "Linhas", "Status", "Detalhes"])
                st.dataframe(df_res, use_container_width=True)
                with st.expander("Logs detalhados"):
                    for l in logs:
                        st.write(l)
            else:
                st.info("Marque a confirmação e clique em 'Executar agora' para aplicar as alterações.")
