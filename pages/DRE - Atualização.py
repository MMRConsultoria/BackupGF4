import streamlit as st
import pandas as pd
import json
import re
from datetime import datetime, timedelta, date
from oauth2client.service_account import ServiceAccountCredentials
import gspread

try:
    from googleapiclient.discovery import build
except Exception:
    build = None

# ---------------- CONFIG ----------------
PASTA_PRINCIPAL_ID = "0B1owaTi3RZnFfm4tTnhfZ2l0VHo4bWNMdHhKS3ZlZzR1ZjRSWWJSSUFxQTJtUExBVlVTUW8"
TARGET_SHEET_NAME = "Configurações Não Apagar"

# Origem FATURAMENTO
ID_PLANILHA_ORIGEM_FAT = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
ABA_ORIGEM_FAT = "Fat Sistema Externo"

# Origem MEIO DE PAGAMENTO
ID_PLANILHA_ORIGEM_MP = "1GSI291SEeeU9MtOWkGwsKGCGMi_xXMSiQnL_9GhXxfU"
ABA_ORIGEM_MP = "Faturamento Meio Pagamento"

st.set_page_config(page_title="Atualizador DRE", layout="wide")

st.markdown(
    """
    <style>
    .block-container { padding-top: 1.2rem; padding-bottom: 1.2rem; }
    [data-testid="stTable"] td, [data-testid="stTable"] th { padding: 8px 12px !important; }
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
@st.cache_data(ttl=300)
def list_child_folders(_drive, parent_id, filtro_texto=None):
    if _drive is None: return []
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
    if _drive is None: return []
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
    except Exception:
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
    if s is None: return None
    s = str(s).strip()
    if s == "" or s in ["-", "–"]: return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    s = s.replace("R$", "").replace("r$", "").replace(" ", "")
    s = re.sub(r"[^0-9,.\-]", "", s)
    if s == "" or s == "-" or s == ".": return None
    if s.count(".") > 0 and s.count(",") > 0:
        s = s.replace(".", "").replace(",", ".")
    else:
        if s.count(",") > 0 and s.count(".") == 0: s = s.replace(",", ".")
        if s.count(".") > 1 and s.count(",") == 0: s = s.replace(".", "")
    try:
        val = float(s)
        if neg: val = -val
        return val
    except: return None

def tratar_numericos(df, headers):
    indices_valor = [6, 7, 8, 9]
    for idx in indices_valor:
        if idx < len(headers):
            col_name = headers[idx]
            df[col_name] = df[col_name].apply(_parse_currency_like).fillna(0.0)
    return df

def format_brl(val):
    try: return f"R$ {float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return val

# ---------------- TABS ----------------
tab_audit, tab_atual = st.tabs(["Auditoria", "Atualização"])

with tab_atual:
    col_d1, col_d2 = st.columns(2)
    with col_d1: data_de = st.date_input("De", value=date.today() - timedelta(days=30), key="at_de")
    with col_d2: data_ate = st.date_input("Até", value=date.today(), key="at_ate")

    try:
        pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
        map_p = {p["name"]: p["id"] for p in pastas_fech}
        p_sel = st.selectbox("Pasta principal:", options=list(map_p.keys()), key="at_p")
        subpastas = list_child_folders(drive_service, map_p[p_sel])
        map_s = {s["name"]: s["id"] for s in subpastas}
        s_sel = st.multiselect("Subpastas:", options=list(map_s.keys()), default=[], key="at_s")
        s_ids = [map_s[n] for n in s_sel]
    except:
        st.error("Erro ao listar pastas."); st.stop()

    if not s_ids:
        st.info("Selecione as subpastas.")
    else:
        planilhas = list_spreadsheets_in_folders(drive_service, s_ids)
        if not planilhas: st.warning("Nenhuma planilha.")
        else:
            df_list = pd.DataFrame(planilhas).sort_values("name").reset_index(drop=True)
            df_list = df_list.rename(columns={"name": "Planilha", "id": "ID_Planilha"})
            c1, c2, c3, _ = st.columns([1.2, 1.2, 1.2, 5])
            with c1: s_desc = st.checkbox("Desconto", value=False, key="at_chk1")
            with c2: s_mp = st.checkbox("Meio Pagto", value=True, key="at_chk2")
            with c3: s_fat = st.checkbox("Faturamento", value=True, key="at_chk3")
            df_list["Desconto"], df_list["Meio Pagamento"], df_list["Faturamento"] = s_desc, s_mp, s_fat
            config = {"Planilha": st.column_config.TextColumn("Planilha", disabled=True), "ID_Planilha": None, "parent_folder_id": None}
            meio = len(df_list)//2 + (len(df_list)%2)
            col_t1, col_t2 = st.columns(2)
            with col_t1: edit_esq = st.data_editor(df_list.iloc[:meio], key="at_t1", use_container_width=True, column_config=config, hide_index=True)
            with col_t2: edit_dir = st.data_editor(df_list.iloc[meio:], key="at_t2", use_container_width=True, column_config=config, hide_index=True)

            if st.button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True):
                df_final_edit = pd.concat([edit_esq, edit_dir], ignore_index=True)
                df_marcadas = df_final_edit[(df_final_edit["Desconto"]) | (df_final_edit["Meio Pagamento"]) | (df_final_edit["Faturamento"])].copy()
                if df_marcadas.empty:
                    st.warning("Nada marcado.")
                    st.stop()
            
                status_placeholder = st.empty()  # placeholder para status
                status_placeholder.info("Executando atualização, por favor aguarde...")
            
                logs = []
                prog = st.progress(0)
                log_placeholder = st.empty()
            
                total = len(df_marcadas)
                for i, (_, row) in enumerate(df_marcadas.iterrows()):
                    try:
                        # ... seu código de atualização ...
                        logs.append(f"{row['Planilha']}:.")
                    except Exception as e:
                        logs.append(f"{row['Planilha']}: Erro {e}")
                    prog.progress((i+1)/total)
                    log_placeholder.text("\n".join(logs))
            
                status_placeholder.success("Em Execução!")
                
                # Carregar Origem Faturamento
                try:
                    sh_orig_fat = gc.open_by_key(ID_PLANILHA_ORIGEM_FAT)
                    ws_orig_fat = sh_orig_fat.worksheet(ABA_ORIGEM_FAT)
                    h_orig_fat, df_orig_fat = get_headers_and_df_raw(ws_orig_fat)
                    c_dt_fat = detect_date_col(h_orig_fat)
                    df_orig_fat['_dt'] = pd.to_datetime(df_orig_fat[c_dt_fat], dayfirst=True, errors='coerce').dt.date
                    df_orig_fat_f = df_orig_fat[(df_orig_fat['_dt'] >= data_de) & (df_orig_fat['_dt'] <= data_ate)].copy()
                except Exception as e: st.error(f"Erro origem Fat: {e}"); st.stop()

                # Carregar Origem Meio Pagamento
                try:
                    sh_orig_mp = gc.open_by_key(ID_PLANILHA_ORIGEM_MP)
                    ws_orig_mp = sh_orig_mp.worksheet(ABA_ORIGEM_MP)
                    h_orig_mp, df_orig_mp = get_headers_and_df_raw(ws_orig_mp)
                    c_dt_mp = detect_date_col(h_orig_mp)
                    df_orig_mp['_dt'] = pd.to_datetime(df_orig_mp[c_dt_mp], dayfirst=True, errors='coerce').dt.date
                    df_orig_mp_f = df_orig_mp[(df_orig_mp['_dt'] >= data_de) & (df_orig_mp['_dt'] <= data_ate)].copy()
                except Exception as e: st.error(f"Erro origem MP: {e}"); st.stop()

                prog = st.progress(0)
                logs = []
                log_placeholder = st.empty()  # placeholder para logs
                
                total = len(df_marcadas)
                for i, (_, row) in enumerate(df_marcadas.iterrows()):
                    try:
                        sid = row["ID_Planilha"]
                        sh_dest = gc.open_by_key(sid)
                        b2, b3 = read_codes_from_config_sheet(sh_dest)
                        if not b2:
                            logs.append(f"{row['Planilha']}: Sem B2.")
                            log_placeholder.text("\n".join(logs))
                            prog.progress((i+1)/total)
                            continue
                        
                        # --- ATUALIZAR FATURAMENTO ---
                        if row["Faturamento"]:
                            c_f, c_d = h_orig_fat[5], h_orig_fat[3]
                            df_ins = df_orig_fat_f[df_orig_fat_f[c_f].astype(str).str.strip() == b2].copy()
                            if b3:
                                df_ins = df_ins[df_ins[c_d].astype(str).str.strip() == b3]
                            
                            if not df_ins.empty:
                                try:
                                    ws_dest = sh_dest.worksheet("Importado_Fat")
                                except:
                                    ws_dest = sh_dest.add_worksheet("Importado_Fat", 1000, 30)
                                
                                h_dest, df_dest = get_headers_and_df_raw(ws_dest)
                                if df_dest.empty:
                                    df_f_ws, h_f = df_ins, h_orig_fat
                                else:
                                    c_dt_d = detect_date_col(h_dest) or c_dt_fat
                                    df_dest['_dt'] = pd.to_datetime(df_dest[c_dt_d], dayfirst=True, errors='coerce').dt.date
                                    rem = (df_dest['_dt'] >= data_de) & (df_dest['_dt'] <= data_ate)
                                    if c_f in df_dest.columns:
                                        rem &= (df_dest[c_f].astype(str).str.strip() == b2)
                                    df_f_ws = pd.concat([df_dest.loc[~rem], df_ins], ignore_index=True)
                                    h_f = h_dest if h_dest else h_orig_fat
                                
                                send = df_f_ws[h_f].fillna("")
                                ws_dest.clear()
                                ws_dest.update("A1", [h_f] + send.values.tolist(), value_input_option='USER_ENTERED')
                                logs.append(f"{row['Planilha']}: Fat OK.")
                                log_placeholder.text("\n".join(logs))
                
                        # --- ATUALIZAR MEIO DE PAGAMENTO ---
                        if row["Meio Pagamento"]:
                            c_f_mp, c_d_mp = h_orig_mp[8], h_orig_mp[6]
                            df_ins_mp = df_orig_mp_f[df_orig_mp_f[c_f_mp].astype(str).str.strip() == b2].copy()
                            if b3:
                                df_ins_mp = df_ins_mp[df_ins_mp[c_d_mp].astype(str).str.strip() == b3]
                            
                            if not df_ins_mp.empty:
                                try:
                                    ws_dest_mp = sh_dest.worksheet("Meio de Pagamento")
                                except:
                                    ws_dest_mp = sh_dest.add_worksheet("Meio de Pagamento", 1000, 30)
                                
                                h_dest_mp, df_dest_mp = get_headers_and_df_raw(ws_dest_mp)
                                if df_dest_mp.empty:
                                    df_f_mp, h_f_mp = df_ins_mp, h_orig_mp
                                else:
                                    c_dt_d_mp = detect_date_col(h_dest_mp) or c_dt_mp
                                    df_dest_mp['_dt'] = pd.to_datetime(df_dest_mp[c_dt_d_mp], dayfirst=True, errors='coerce').dt.date
                                    rem_mp = (df_dest_mp['_dt'] >= data_de) & (df_dest_mp['_dt'] <= data_ate)
                                    if c_f_mp in df_dest_mp.columns:
                                        rem_mp &= (df_dest_mp[c_f_mp].astype(str).str.strip() == b2)
                                    df_f_mp = pd.concat([df_dest_mp.loc[~rem_mp], df_ins_mp], ignore_index=True)
                                    h_f_mp = h_dest_mp if h_dest_mp else h_orig_mp
                                
                                send_mp = df_f_mp[h_f_mp].fillna("")
                                ws_dest_mp.clear()
                                ws_dest_mp.update("A1", [h_f_mp] + send_mp.values.tolist(), value_input_option='USER_ENTERED')
                                logs.append(f"{row['Planilha']}: MP OK.")
                                log_placeholder.text("\n".join(logs))
                
                    except Exception as e:
                        logs.append(f"{row['Planilha']}: Erro {e}")
                        log_placeholder.text("\n".join(logs))
                    prog.progress((i+1)/total)
                
                st.success("Concluido!")

with tab_audit:
    st.header("Auditoria")
    try:
        pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
        map_p = {p["name"]: p["id"] for p in pastas_fech}
        p_sel = st.selectbox("Pasta principal:", options=list(map_p.keys()), key="au_p")
        subpastas = list_child_folders(drive_service, map_p[p_sel])
        map_s = {s["name"]: s["id"] for s in subpastas}
        s_sel = st.multiselect("Subpastas:", options=list(map_s.keys()), default=[], key="au_s")
        s_ids_audit = [map_s[n] for n in s_sel]
    except:
        st.error("Erro ao listar pastas.")
        st.stop()

    c_ano, c_mes = st.columns(2)
    with c_ano:
        ano_sel = st.selectbox("Ano:", list(range(2020, date.today().year + 1)), index=date.today().year - 2020, key="au_ano")
    with c_mes:
        mes_sel = st.selectbox("Mês (Opcional):", ["Todos"] + list(range(1, 13)), key="au_mes")

    if st.button("📊 EXECUTAR AUDITORIA", use_container_width=True):
        if not s_ids_audit:
            st.warning("Selecione subpastas.")
            st.stop()

        if mes_sel == "Todos":
            d_ini, d_fim = date(ano_sel, 1, 1), date(ano_sel, 12, 31)
        else:
            d_ini = date(ano_sel, int(mes_sel), 1)
            d_fim = (date(ano_sel, int(mes_sel), 28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)

        try:
            # Origem FATURAMENTO
            sh_o_fat = gc.open_by_key(ID_PLANILHA_ORIGEM_FAT)
            ws_o_fat = sh_o_fat.worksheet(ABA_ORIGEM_FAT)
            h_o_fat, df_o_fat = get_headers_and_df_raw(ws_o_fat)
            df_o_fat = tratar_numericos(df_o_fat, h_o_fat)
            c_dt_o_fat = detect_date_col(h_o_fat)
            df_o_fat['_dt'] = pd.to_datetime(df_o_fat[c_dt_o_fat], dayfirst=True, errors='coerce').dt.date
            df_o_fat_p = df_o_fat[(df_o_fat['_dt'] >= d_ini) & (df_o_fat['_dt'] <= d_fim)].copy()

            # Origem MEIO DE PAGAMENTO
            sh_o_mp = gc.open_by_key(ID_PLANILHA_ORIGEM_MP)
            ws_o_mp = sh_o_mp.worksheet(ABA_ORIGEM_MP)
            h_o_mp, df_o_mp = get_headers_and_df_raw(ws_o_mp)
            df_o_mp = tratar_numericos(df_o_mp, h_o_mp)
            c_dt_o_mp = detect_date_col(h_o_mp)
            df_o_mp['_dt'] = pd.to_datetime(df_o_mp[c_dt_o_mp], dayfirst=True, errors='coerce').dt.date
            df_o_mp_p = df_o_mp[(df_o_mp['_dt'] >= d_ini) & (df_o_mp['_dt'] <= d_fim)].copy()

            planilhas = list_spreadsheets_in_folders(drive_service, s_ids_audit)
            results = []
            prog = st.progress(0)
            total = len(planilhas)

            for idx, p in enumerate(planilhas):
                sid, p_name = p["id"], p["name"]
                sh_d = gc.open_by_key(sid)
                b2, b3 = read_codes_from_config_sheet(sh_d)

                if not b2:
                    results.append({
                        "Planilha": p_name,
                        "Origem": 0,
                        "DRE": 0,
                        "MP Origem": 0,
                        "MP DRE": 0,
                        "Dif": 0,
                        "Dif MP": 0,
                        "Status": "Sem Config"
                    })
                else:
                    # FATURAMENTO
                    df_o_f = df_o_fat_p[df_o_fat_p[h_o_fat[5]].astype(str).str.strip() == b2]
                    if b3:
                        df_o_f = df_o_f[df_o_fat_p[h_o_fat[3]].astype(str).str.strip() == b3]
                    v_o = df_o_f[h_o_fat[6]].sum()

                    v_d = 0.0
                    v_mp_o = 0.0
                    v_mp_d = 0.0

                    try:
                        ws_d = sh_d.worksheet("Importado_Fat")
                        h_d, df_d = get_headers_and_df_raw(ws_d)
                        df_d = tratar_numericos(df_d, h_d)
                        c_dt_d = detect_date_col(h_d)
                        df_d['_dt'] = pd.to_datetime(df_d[c_dt_d], dayfirst=True, errors='coerce').dt.date

                        v_d = df_d[(df_d['_dt'] >= d_ini) & (df_d['_dt'] <= d_fim)][h_d[6]].sum()

                        # MEIO DE PAGAMENTO na aba Importado_Fat
                        idx_mp_col = 7 if len(h_d) > 7 else None
                        if idx_mp_col is not None:
                            col_mp = h_d[idx_mp_col]
                            v_mp_d = df_d[(df_d['_dt'] >= d_ini) & (df_d['_dt'] <= d_fim)][col_mp].sum()
                    except:
                        pass

                    # MEIO DE PAGAMENTO ORIGEM
                    df_o_mp_f = df_o_mp_p[df_o_mp_p[h_o_mp[8]].astype(str).str.strip() == b2]
                    if b3:
                        df_o_mp_f = df_o_mp_f[df_o_mp_p[h_o_mp[6]].astype(str).str.strip() == b3]
                    v_mp_o = df_o_mp_f[h_o_mp[9]].sum() if len(h_o_mp) > 9 else 0

                    diff = v_o - v_d
                    diff_mp = v_mp_o - v_mp_d

                    results.append({
                        "Planilha": p_name,
                        "Origem": v_o,
                        "DRE": v_d,
                        "MP Origem": v_mp_o,
                        "MP DRE": v_mp_d,
                        "Dif": diff,
                        "Dif MP": diff_mp,
                        "Status": "✅ OK" if abs(diff) < 0.01 and abs(diff_mp) < 0.01 else "❌ Erro"
                    })

                prog.progress((idx + 1) / total)

            df_res = pd.DataFrame(results)
            for c in ["Origem", "DRE", "MP Origem", "MP DRE", "Dif", "Dif MP"]:
                df_res[c] = df_res[c].apply(format_brl)
            st.table(df_res)

        except Exception as e:
            st.error(f"Erro: {e}")
