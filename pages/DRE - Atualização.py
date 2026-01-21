import streamlit as st
import pandas as pd
import json
import time
import re
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
            res[sid] = any(target_clean in t for t in titles)
        except:
            res[sid] = False
    return res

def read_codes_from_config_sheet(gsheet):
    try:
        ws = None
        for w in gsheet.worksheets():
            if TARGET_SHEET_NAME.strip().lower() in w.title.strip().lower():
                ws = w
                break
        if ws is None: return None, None
        b2 = ws.acell("B2").value
        b3 = ws.acell("B3").value
        return (str(b2).strip() if b2 else None, str(b3).strip() if b3 else None)
    except:
        return None, None

def get_headers_and_df_raw(ws):
    vals = ws.get_all_values()
    if not vals: return [], pd.DataFrame()
    headers = [str(h).strip() for h in vals[0]]
    df = pd.DataFrame(vals[1:], columns=headers)
    return headers, df

def detect_date_col(headers):
    for h in headers:
        if "data" in h.lower(): return h
    return None

def _parse_currency_like(s):
    """Tenta converter uma string de formato monetário para float.
    Retorna float se conseguir, None se string vazia/nula, ou None se falhar (o chamador decide manter original)."""
    if s is None: 
        return None
    s = str(s).strip()
    if s == "" or s in ["-", "–"]:
        return None
    # Detecta parênteses (negativo)
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    # Remove prefixos tipo R$, espaços e outros símbolos
    s = s.replace("R$", "").replace("r$", "").replace(" ", "")
    # Remove tudo que não seja dígito, vírgula, ponto ou sinal de menos
    s = re.sub(r"[^0-9,.\-]", "", s)
    if s == "" or s == "-" or s == ".": 
        return None
    # Se tem '.' e ',', assumir '.' milhares e ',' decimal -> remover '.' e trocar ',' por '.'
    if s.count(".") > 0 and s.count(",") > 0:
        s = s.replace(".", "").replace(",", ".")
    else:
        # Se só tem vírgula, trocar por ponto (ex: '1234,56' -> '1234.56')
        if s.count(",") > 0 and s.count(".") == 0:
            s = s.replace(",", ".")
        # Se só tem pontos, mas muitos pontos (ex: '1.234.567'), pode ser milhares.
        # Se houver mais de 1 ponto e nenhum vírgula, remover pontos (assumir milhares)
        if s.count(".") > 1 and s.count(",") == 0:
            s = s.replace(".", "")
    # Tenta converter
    try:
        val = float(s)
        if neg:
            val = -val
        return val
    except:
        return None

def tratar_numericos(df, headers):
    """Converte colunas G, H, I, J (índices 6, 7, 8, 9) para numérico quando possível.
    Se não for possível converter uma célula, mantemos o valor original (ou string vazia)."""
    indices_valor = [6, 7, 8, 9]  # G, H, I, J
    for idx in indices_valor:
        if idx < len(headers):
            col_name = headers[idx]
            # Aplicar parse robusto, sem sobrescrever com 0 quando falhar
            orig_series = df[col_name].astype(object).copy()
            parsed = orig_series.apply(_parse_currency_like)
            # Montar nova coluna: se parsed não for None -> número (float). Se None:
            # - se orig vazio ou '-', colocar '' (string vazia) para não inserir 0
            # - senão manter string original (ex: texto que não é número)
            new_col = []
            for p, o in zip(parsed, orig_series):
                if p is not None:
                    new_col.append(p)  # float
                else:
                    o_str = "" if pd.isna(o) else str(o).strip()
                    if o_str == "" or o_str in ["-", "–"]:
                        new_col.append("")  # vazio
                    else:
                        new_col.append(o_str)  # mantém texto original
            df[col_name] = new_col
    return df

# ---------------- INTERFACE ----------------
col_d1, col_d2 = st.columns(2)
with col_d1: data_de = st.date_input("De", value=date.today() - timedelta(days=30))
with col_d2: data_ate = st.date_input("Até", value=date.today())

try:
    pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
    map_p = {p["name"]: p["id"] for p in pastas_fech}
    p_sel = st.selectbox("Pasta principal:", options=list(map_p.keys()))
    subpastas = list_child_folders(drive_service, map_p[p_sel])
    map_s = {s["name"]: s["id"] for s in subpastas}
    s_sel = st.multiselect("Subpastas:", options=list(map_s.keys()), default=[])
    s_ids = [map_s[n] for n in s_sel]
except: st.stop()

if not s_ids:
    st.info("Selecione as subpastas para listar as planilhas.")
    st.stop()

if s_ids:
    with st.spinner("Buscando planilhas..."):
        planilhas = list_spreadsheets_in_folders(drive_service, s_ids)
        if planilhas:
            df_list = pd.DataFrame(planilhas).sort_values("name").reset_index(drop=True)
            df_list = df_list.rename(columns={"name": "Planilha", "id": "ID_Planilha"})
            conf_map = get_conf_map(df_list["ID_Planilha"].tolist(), TARGET_SHEET_NAME)
            df_list["conf"] = df_list["ID_Planilha"].map(conf_map).astype(bool)
            
            st.markdown('<div class="global-selection-container">', unsafe_allow_html=True)
            c1, c2, c3, _ = st.columns([1.2, 1.2, 1.2, 5])
            with c1: s_desc = st.checkbox("Desconto", value=True)
            with c2: s_mp = st.checkbox("Meio Pagto", value=True)
            with c3: s_fat = st.checkbox("Faturamento", value=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
            df_list["Desconto"], df_list["Meio Pagamento"], df_list["Faturamento"] = s_desc, s_mp, s_fat
            config = {
                "Planilha": st.column_config.TextColumn("Planilha", disabled=True),
                "conf": st.column_config.CheckboxColumn("Conf", disabled=True),
                "ID_Planilha": None, "parent_folder_id": None,
                "Desconto": st.column_config.CheckboxColumn("Desc."),
                "Meio Pagamento": st.column_config.CheckboxColumn("M.Pag"),
                "Faturamento": st.column_config.CheckboxColumn("Fat."),
            }
            meio = len(df_list) // 2 + (len(df_list) % 2)
            col_t1, col_t2 = st.columns(2)
            with col_t1: edit_esq = st.data_editor(df_list.iloc[:meio], key="t1", use_container_width=True, column_config=config, hide_index=True)
            with col_t2: edit_dir = st.data_editor(df_list.iloc[meio:], key="t2", use_container_width=True, column_config=config, hide_index=True)
            
            if st.button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True):
                df_final_edit = pd.concat([edit_esq, edit_dir], ignore_index=True)
                df_marcadas = df_final_edit[(df_final_edit["Desconto"]) | (df_final_edit["Meio Pagamento"]) | (df_final_edit["Faturamento"])].copy()
                
                if df_marcadas.empty:
                    st.warning("Nenhuma planilha marcada.")
                    st.stop()

                try:
                    sh_origem = gc.open_by_key(ID_PLANILHA_ORIGEM)
                    ws_origem = sh_origem.worksheet(ABA_ORIGEM)
                    headers_orig, df_orig = get_headers_and_df_raw(ws_origem)
                    
                    # Trata colunas G,H,I,J como números quando possível
                    df_orig = tratar_numericos(df_orig, headers_orig)
                    
                    col_data_orig = detect_date_col(headers_orig)
                    df_orig_temp = df_orig.copy()
                    # Converter coluna de data para date para comparar com st.date_input
                    df_orig_temp['_dt'] = pd.to_datetime(df_orig_temp[col_data_orig], dayfirst=True, errors='coerce').dt.date
                    mask_orig = (df_orig_temp['_dt'] >= data_de) & (df_orig_temp['_dt'] <= data_ate)
                    df_orig_filtrado = df_orig.loc[mask_orig].copy()
                except Exception as e:
                    st.error(f"Erro na origem: {e}"); st.stop()

                progresso = st.progress(0)
                logs = []
                total = len(df_marcadas)

                for i, (_, row) in enumerate(df_marcadas.iterrows()):
                    try:
                        sh_dest = gc.open_by_key(row["ID_Planilha"])
                        b2, b3 = read_codes_from_config_sheet(sh_dest)
                        if not b2:
                            logs.append(f"{row['Planilha']}: B2 não encontrado."); continue

                        col_f_name = headers_orig[5] # Grupo
                        col_d_name = headers_orig[3] # Loja
                        
                        df_para_inserir = df_orig_filtrado[df_orig_filtrado[col_f_name].astype(str).str.strip() == b2].copy()
                        if b3:
                            df_para_inserir = df_para_inserir[df_para_inserir[col_d_name].astype(str).str.strip() == b3]

                        if df_para_inserir.empty:
                            logs.append(f"{row['Planilha']}: Sem dados para o período."); continue

                        try:
                            ws_dest = sh_dest.worksheet("Importado_Fat")
                        except:
                            ws_dest = sh_dest.add_worksheet("Importado_Fat", 1000, 30)
                        
                        headers_dest, df_dest = get_headers_and_df_raw(ws_dest)
                        df_dest = tratar_numericos(df_dest, headers_dest)
                        
                        if df_dest.empty:
                            df_final_ws = df_para_inserir
                            h_final = headers_orig
                        else:
                            col_dt_dest = detect_date_col(headers_dest) or col_data_orig
                            df_dest_temp = df_dest.copy()
                            df_dest_temp['_dt'] = pd.to_datetime(df_dest_temp[col_dt_dest], dayfirst=True, errors='coerce').dt.date
                            
                            to_remove = (df_dest_temp['_dt'] >= data_de) & (df_dest_temp['_dt'] <= data_ate)
                            if col_f_name in df_dest.columns:
                                to_remove &= (df_dest[col_f_name].astype(str).str.strip() == b2)
                            if b3 and col_d_name in df_dest.columns:
                                to_remove &= (df_dest[col_d_name].astype(str).str.strip() == b3)
                            
                            df_restante = df_dest.loc[~to_remove]
                            df_final_ws = pd.concat([df_restante, df_para_inserir], ignore_index=True)
                            h_final = headers_dest if headers_dest else headers_orig

                        # Antes de enviar para o Sheets, substituímos valores pandas.NA/NaN por string vazia,
                        # e mantemos floats como floats para que o Sheets receba números (com USER_ENTERED).
                        send_df = df_final_ws[h_final].copy()
                        send_df = send_df.where(pd.notna(send_df), "")
                        final_vals = [h_final] + send_df.values.tolist()

                        ws_dest.clear()
                        ws_dest.update("A1", final_vals, value_input_option='USER_ENTERED')
                        logs.append(f"{row['Planilha']}: Sucesso ({len(df_para_inserir)} linhas)")
                    except Exception as e:
                        logs.append(f"{row['Planilha']}: Erro: {e}")
                    progresso.progress(min((i + 1) / total, 1.0))

                st.success("Concluído!"); st.write("\n".join(logs))
