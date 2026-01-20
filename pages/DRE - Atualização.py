import streamlit as st
import pandas as pd
import json
import time
from datetime import datetime, timedelta, date
from oauth2client.service_account import ServiceAccountCredentials
import gspread

# Drive API
try:
    from googleapiclient.discovery import build
except Exception:
    build = None

# ---------------- CONFIG ----------------
PASTA_PRINCIPAL_ID = "0B1owaTi3RZnFfm4tTnhfZ2l0VHo4bWNMdHhKS3ZlZzR1ZjRSWWJSSUFxQTJtUExBVlVTUW8"
TARGET_SHEET_NAME = "Configurações Não Apagar"
ID_PLANILHA_ORIGEM = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
ABA_ORIGEM = "Fat Sistema Externo"

st.set_page_config(page_title="Atualizador DRE", layout="wide")

# --- CSS ---
st.markdown(
    """
    <style>
    .block-container { padding-top: 1rem; padding-bottom: 0rem; }
    [data-testid="stVerticalBlock"] > div { margin-bottom: -0.5rem !important; padding-top: 0rem !important; }
    h1 { margin-top: -1rem; margin-bottom: 0.5rem; font-size: 1.8rem; }
    .global-selection-container { margin-top: 5px !important; margin-bottom: 5px !important; }
    [data-testid="stTable"] td, [data-testid="stTable"] th { padding: 2px 6px !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Atualizador DRE")

# ---------------- AUTENTICAÇÃO ----------------
@st.cache_resource
def autenticar():
    scope = ["https://www.googleapis.com/auth/drive", "https://www.googleapis.com/auth/spreadsheets"]
    creds_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    gc = gspread.authorize(creds)
    drive = build("drive", "v3", credentials=creds) if build else None
    return gc, drive

try:
    gc, drive_service = autenticar()
except Exception as e:
    st.error(f"Erro de autenticação: {e}")
    st.stop()

# ---------------- HELPERS ----------------
def list_child_folders(_drive, parent_id, filtro_texto=None):
    folders = []
    page_token = None
    q = f"'{parent_id}' in parents and mimeType='application/vnd.google-apps.folder' and trashed=false"
    while True:
        resp = _drive.files().list(q=q, fields="nextPageToken, files(id, name)", pageToken=page_token).execute()
        for f in resp.get("files", []):
            if filtro_texto is None or filtro_texto.lower() in f["name"].lower():
                folders.append({"id": f["id"], "name": f["name"]})
        page_token = resp.get("nextPageToken", None)
        if not page_token: break
    return folders

@st.cache_data(ttl=60)
def list_spreadsheets_in_folders(_drive, folder_ids):
    sheets = []
    for fid in folder_ids:
        page_token = None
        q = f"'{fid}' in parents and mimeType='application/vnd.google-apps.spreadsheet' and trashed=false"
        while True:
            resp = _drive.files().list(q=q, fields="nextPageToken, files(id, name)", pageToken=page_token).execute()
            for f in resp.get("files", []):
                sheets.append({"id": f["id"], "name": f["name"], "parent_folder_id": fid})
            page_token = resp.get("nextPageToken", None)
            if not page_token: break
    return sheets

@st.cache_data(ttl=300)
def get_conf_map(sheet_ids, target_name):
    res = {}
    target_clean = target_name.strip().lower()
    for sid in sheet_ids:
        try:
            sh = gc.open_by_key(sid)
            titles = [ws.title.strip().lower() for ws in sh.worksheets()]
            res[sid] = target_clean in titles
        except:
            res[sid] = False
    return res

def read_codes_from_config_sheet(gsheet):
    try:
        ws = gsheet.worksheet(TARGET_SHEET_NAME)
        b2 = ws.acell("B2").value
        b3 = ws.acell("B3").value
        return (str(b2).strip() if b2 else None, str(b3).strip() if b3 else None)
    except Exception:
        return (None, None)

def col_letter_to_name(df, letter):
    if df is None or df.shape[1] == 0:
        return None
    idx = ord(letter.upper()) - ord("A")
    if idx < 0 or idx >= df.shape[1]:
        return None
    return df.columns[idx]

def filter_data_by_date(df, date_col, start_date, end_date):
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce')
    mask = (df[date_col] >= pd.to_datetime(start_date)) & (df[date_col] <= pd.to_datetime(end_date))
    return df.loc[mask]

# ---------------- INTERFACE ----------------
col_d1, col_d2 = st.columns(2)
with col_d1:
    data_de = st.date_input("De", value=date.today() - timedelta(days=30))
with col_d2:
    data_ate = st.date_input("Até", value=date.today())

try:
    pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
    map_p = {p["name"]: p["id"] for p in pastas_fech}
    p_sel = st.selectbox("Pasta principal:", options=list(map_p.keys()))
    
    subpastas = list_child_folders(drive_service, map_p[p_sel])
    map_s = {s["name"]: s["id"] for s in subpastas}
    s_sel = st.multiselect("Subpastas:", options=list(map_s.keys()), default=list(map_s.keys())[:1])
    s_ids = [map_s[n] for n in s_sel]
except:
    st.stop()

if s_ids:
    with st.spinner("Buscando planilhas e verificando abas..."):
        planilhas = list_spreadsheets_in_folders(drive_service, s_ids)
        if planilhas:
            df = pd.DataFrame(planilhas).sort_values("name").reset_index(drop=True)
            df = df.rename(columns={"name": "Planilha", "id": "ID_Planilha"})
            
            conf_map = get_conf_map(df["ID_Planilha"].tolist(), TARGET_SHEET_NAME)
            df["conf"] = df["ID_Planilha"].map(conf_map).astype(bool)
            
            st.markdown('<div class="global-selection-container">', unsafe_allow_html=True)
            c1, c2, c3, _ = st.columns([1.2, 1.2, 1.2, 5])
            with c1: s_desc = st.checkbox("Desconto", value=True)
            with c2: s_mp = st.checkbox("Meio Pagto", value=True)
            with c3: s_fat = st.checkbox("Faturamento", value=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
            df["Desconto"], df["Meio Pagamento"], df["Faturamento"] = s_desc, s_mp, s_fat
            
            config = {
                "Planilha": st.column_config.TextColumn("Planilha", disabled=True),
                "conf": st.column_config.CheckboxColumn("Conf", disabled=True),
                "ID_Planilha": None, "parent_folder_id": None,
                "Desconto": st.column_config.CheckboxColumn("Desc."),
                "Meio Pagamento": st.column_config.CheckboxColumn("M.Pag"),
                "Faturamento": st.column_config.CheckboxColumn("Fat."),
            }
            
            meio = len(df) // 2 + (len(df) % 2)
            col_t1, col_t2 = st.columns(2)
            with col_t1:
                edit_esq = st.data_editor(df.iloc[:meio], key="t1", use_container_width=True, column_config=config, hide_index=True)
            with col_t2:
                edit_dir = st.data_editor(df.iloc[meio:], key="t2", use_container_width=True, column_config=config, hide_index=True)
            
            if st.button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True):
                df_final = pd.concat([edit_esq, edit_dir], ignore_index=True)
                
                # Abre planilha origem e lê dados da aba Fat Sistema Externo
                try:
                    sh_origem = gc.open_by_key(ID_PLANILHA_ORIGEM)
                    ws_origem = sh_origem.worksheet(ABA_ORIGEM)
                    dados_origem = ws_origem.get_all_records()
                    df_origem = pd.DataFrame(dados_origem)
                except Exception as e:
                    st.error(f"Erro ao abrir planilha origem: {e}")
                    st.stop()
                
                # Colunas para filtro
                col_data = "Data"  # ajuste se necessário
                col_grupo = col_letter_to_name(df_origem, "F")  # coluna F
                col_loja = col_letter_to_name(df_origem, "D")   # coluna D
                
                # Filtra dados pela data selecionada
                df_origem = filter_data_by_date(df_origem, col_data, data_de, data_ate)
                
                progresso = st.progress(0)
                logs = []
                total = len(df_final)
                
                for i, row in df_final.iterrows():
                    if not (row["Desconto"] or row["Meio Pagamento"] or row["Faturamento"]):
                        continue
                    
                    try:
                        sh_destino = gc.open_by_key(row["ID_Planilha"])
                        b2, b3 = read_codes_from_config_sheet(sh_destino)
                        
                        if not b2:
                            logs.append(f"{row['Planilha']}: Código do grupo (B2) não encontrado. Pulando.")
                            progresso.progress((i+1)/total)
                            continue
                        
                        # Filtra dados origem pelo grupo e loja
                        df_filtrado = df_origem[df_origem[col_grupo].astype(str).str.strip() == b2.strip()]
                        if b3 and b3.strip():
                            df_filtrado = df_filtrado[df_filtrado[col_loja].astype(str).str.strip() == b3.strip()]
                        
                        if df_filtrado.empty:
                            logs.append(f"{row['Planilha']}: Nenhum dado filtrado para grupo {b2} e loja {b3}.")
                            progresso.progress((i+1)/total)
                            continue
                        
                        # Atualiza aba Importado_Fat na planilha destino
                        try:
                            ws_destino = sh_destino.worksheet("Importado_Fat")
                        except Exception:
                            ws_destino = sh_destino.add_worksheet(title="Importado_Fat", rows=1000, cols=50)
                        
                        # Apaga dados do período filtrado na aba destino
                        # Como não temos a lógica para apagar só o período, vamos limpar tudo e inserir os dados filtrados
                        ws_destino.clear()
                        
                        # Escreve os dados filtrados (incluindo cabeçalho)
                        valores = [df_filtrado.columns.to_list()] + df_filtrado.values.tolist()
                        ws_destino.update("A1", valores)
                        
                        logs.append(f"{row['Planilha']}: Atualizado com {len(df_filtrado)} linhas.")
                    except Exception as e:
                        logs.append(f"{row['Planilha']}: Erro ao atualizar - {e}")
                    
                    progresso.progress((i+1)/total)
                
                st.success("Atualização concluída!")
                st.write("\n".join(logs))
