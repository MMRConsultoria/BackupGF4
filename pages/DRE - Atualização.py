

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
        if ws is None: return None, None, None, None
        b2 = ws.acell("B2").value
        b3 = ws.acell("B3").value
        b4 = ws.acell("B4").value
        b5 = ws.acell("B5").value
        return (
            str(b2).strip() if b2 else None,
            str(b3).strip() if b3 else None,
            str(b4).strip() if b4 else None,
            str(b5).strip() if b5 else None
        )
    except Exception:
        return None, None, None, None

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
                    df_orig_fat["_dt"] = pd.to_datetime(df_orig_fat[c_dt_fat], dayfirst=True, errors="coerce").dt.date
                    df_orig_fat_f = df_orig_fat[(df_orig_fat["_dt"] >= data_de) & (df_orig_fat["_dt"] <= data_ate)].copy()
                except Exception as e: st.error(f"Erro origem Fat: {e}"); st.stop()

                # Carregar Origem Meio Pagamento
                try:
                    sh_orig_mp = gc.open_by_key(ID_PLANILHA_ORIGEM_MP)
                    ws_orig_mp = sh_orig_mp.worksheet(ABA_ORIGEM_MP)
                    h_orig_mp, df_orig_mp = get_headers_and_df_raw(ws_orig_mp)
                    c_dt_mp = detect_date_col(h_orig_mp)
                    df_orig_mp["_dt"] = pd.to_datetime(df_orig_mp[c_dt_mp], dayfirst=True, errors="coerce").dt.date
                    df_orig_mp_f = df_orig_mp[(df_orig_mp["_dt"] >= data_de) & (df_orig_mp["_dt"] <= data_ate)].copy()
                except Exception as e: st.error(f"Erro origem MP: {e}"); st.stop()

                prog = st.progress(0)
                logs = []
                log_placeholder = st.empty()  # placeholder para logs

                total = len(df_marcadas)
                for i, (_, row) in enumerate(df_marcadas.iterrows()):
                    try:
                        sid = row["ID_Planilha"]
                        sh_dest = gc.open_by_key(sid)
                        b2, b3, b4, b5 = read_codes_from_config_sheet(sh_dest)
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
                            if b4 and len(h_orig_fat) > 4:
                                df_ins = df_ins[df_ins[h_orig_fat[4]].astype(str).str.strip() == b4]
                            if b5 and len(h_orig_fat) > 2:
                                df_ins = df_ins[df_ins[h_orig_fat[2]].astype(str).str.strip() == b5]

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
                                    df_dest["_dt"] = pd.to_datetime(df_dest[c_dt_d], dayfirst=True, errors="coerce").dt.date
                                    rem = (df_dest["_dt"] >= data_de) & (df_dest["_dt"] <= data_ate)
                                    if c_f in df_dest.columns:
                                        rem &= (df_dest[c_f].astype(str).str.strip() == b2)
                                    df_f_ws = pd.concat([df_dest.loc[~rem], df_ins], ignore_index=True)
                                    h_f = h_dest if h_dest else h_orig_fat

                                send = df_f_ws[h_f].fillna("")
                                ws_dest.clear()
                                ws_dest.update("A1", [h_f] + send.values.tolist(), value_input_option="USER_ENTERED")
                                logs.append(f"{row['Planilha']}: Fat OK.")
                                log_placeholder.text("\n".join(logs))

                        # --- ATUALIZAR MEIO DE PAGAMENTO ---
                        if row["Meio Pagamento"]:
                            c_f_mp, c_d_mp = h_orig_mp[8], h_orig_mp[6]
                            df_ins_mp = df_orig_mp_f[df_orig_mp_f[c_f_mp].astype(str).str.strip() == b2].copy()
                            if b3:
                                df_ins_mp = df_ins_mp[df_ins_mp[c_d_mp].astype(str).str.strip() == b3]
                            if b4 and len(h_orig_mp) > 7:
                                df_ins_mp = df_ins_mp[df_ins_mp[h_orig_mp[7]].astype(str).str.strip() == b4]
                            if b5 and len(h_orig_mp) > 5:
                                df_ins_mp = df_ins_mp[df_ins_mp[h_orig_mp[5]].astype(str).str.strip() == b5]

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
                                    df_dest_mp["_dt"] = pd.to_datetime(df_dest_mp[c_dt_d_mp], dayfirst=True, errors="coerce").dt.date
                                    rem_mp = (df_dest_mp["_dt"] >= data_de) & (df_dest_mp["_dt"] <= data_ate)
                                    if c_f_mp in df_dest_mp.columns:
                                        rem_mp &= (df_dest_mp[c_f_mp].astype(str).str.strip() == b2)
                                    df_f_mp = pd.concat([df_dest_mp.loc[~rem_mp], df_ins_mp], ignore_index=True)
                                    h_f_mp = h_dest_mp if h_dest_mp else h_orig_mp

                                send_mp = df_f_mp[h_f_mp].fillna("")
                                ws_dest_mp.clear()
                                ws_dest_mp.update("A1", [h_f_mp] + send_mp.values.tolist(), value_input_option="USER_ENTERED")
                                logs.append(f"{row['Planilha']}: MP OK.")
                                log_placeholder.text("\n".join(logs))

                    except Exception as e:
                        logs.append(f"{row['Planilha']}: Erro {e}")
                        log_placeholder.text("\n".join(logs))
                    prog.progress((i+1)/total)

                st.success("Concluido!")

# Aba Auditoria completa (cole onde tab_audit está definido)
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from st_aggrid.shared import JsCode
import pandas as pd
from datetime import date, timedelta
import streamlit as st

with tab_audit:
    st.header("Auditoria")

    # -----------------------
    # Helpers
    # -----------------------
    def format_brl(v):
        try:
            v = float(v)
        except Exception:
            return ""
        s = f"{v:,.2f}"
        s = s.replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {s}"

    def detect_column_by_keywords(headers, keywords_list):
        for kw in keywords_list:
            for h in headers:
                if kw in str(h).lower():
                    return h
        return None

    def normalize_code(val):
        try:
            f = float(val)
            i = int(f)
            return str(i) if f == i else str(f)
        except Exception:
            return str(val).strip()

    def to_bool_like(x):
        if isinstance(x, bool):
            return x
        s = str(x).strip().lower()
        return s in ("true", "t", "1", "yes", "y", "sim", "s")

    # -----------------------
    # Pastas / Subpastas
    # -----------------------
    try:
        pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
        if not pastas_fech:
            st.error("Nenhuma pasta de fechamento encontrada na pasta principal.")
            st.stop()
        map_p = {p["name"]: p["id"] for p in pastas_fech}
        p_sel = st.selectbox("Pasta principal:", options=list(map_p.keys()), key="au_p")
        subpastas = list_child_folders(drive_service, map_p[p_sel])
        map_s = {s["name"]: s["id"] for s in subpastas}
        s_sel = st.multiselect("Subpastas (se nenhuma, trará todas):", options=list(map_s.keys()), default=[], key="au_s")
        s_ids_audit = [map_s[n] for n in s_sel] if s_sel else list(map_s.values())
    except Exception as e:
        st.error(f"Erro ao listar pastas/subpastas: {e}")
        st.stop()

    # -----------------------
    # Filtros de período
    # -----------------------
    c1, c2 = st.columns(2)
    with c1:
        ano_sel = st.selectbox("Ano:", list(range(2020, date.today().year + 1)),
                               index=max(0, date.today().year - 2020), key="au_ano")
    with c2:
        mes_sel = st.selectbox("Mês (Opcional):", ["Todos"] + list(range(1, 13)), key="au_mes")

    # -----------------------
    # Carregar planilhas (recarrega se subpastas mudarem)
    # -----------------------
    need_reload = ("au_last_subpastas" not in st.session_state) or (st.session_state.get("au_last_subpastas") != s_ids_audit)
    if need_reload:
        try:
            planilhas = list_spreadsheets_in_folders(drive_service, s_ids_audit)
        except Exception as e:
            st.error(f"Erro ao listar planilhas nas subpastas: {e}")
            st.stop()

        df_init = pd.DataFrame([{
            "Planilha": p["name"],
            "Flag": False,
            "Planilha_id": p["id"],
            "Origem": "",
            "DRE": "",
            "MP DRE": "",
            "Dif": "",
            "Dif MP": "",
            "Status": ""
        } for p in planilhas])

        st.session_state.au_last_subpastas = s_ids_audit
        st.session_state.au_planilhas_df = df_init
        st.session_state.au_resultados = {}
        st.session_state.au_flags_temp = {}

    # garantir chaves no session_state
    if "au_planilhas_df" not in st.session_state:
        st.session_state.au_planilhas_df = pd.DataFrame(columns=["Planilha", "Flag", "Planilha_id", "Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"])
    if "au_flags_temp" not in st.session_state:
        st.session_state.au_flags_temp = {}
    if "au_resultados" not in st.session_state:
        st.session_state.au_resultados = {}

    df_table = st.session_state.au_planilhas_df.copy()
    if df_table.empty:
        st.info("Nenhuma planilha encontrada nas subpastas selecionadas.")
        st.stop()

    # -----------------------
    # Preparar display_df (garantir colunas)
    # NOTA: NÃO atualizamos st.session_state durante a edição do grid.
    # -----------------------
    expected_cols = ["Planilha", "Flag", "Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"]
    for c in expected_cols:
        if c not in df_table.columns:
            df_table[c] = False if c == "Flag" else ""
    display_df = df_table[expected_cols].copy()

    # -----------------------
    # AgGrid config (NO_UPDATE evita re-renders automáticos)
    # -----------------------
    row_style_js = JsCode("""
    function(params) {
        if (params.data && (params.data.Flag === true || params.data.Flag === 'true')) {
            return {'background-color': '#e9f7ee'};
        }
    }
    """)

    gb = GridOptionsBuilder.from_dataframe(display_df)
    gb.configure_column("Planilha", headerName="Planilha", editable=False, width=420)
    gb.configure_column("Flag",
                        headerName="",
                        editable=True,
                        cellEditor="agCheckboxCellEditor",
                        cellRenderer="agCheckboxCellRenderer",
                        width=80)
    for col in ["Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"]:
        if col in display_df.columns:
            gb.configure_column(col, editable=False)
    grid_options = gb.build()
    grid_options["getRowStyle"] = row_style_js

    st.markdown("Marque as planilhas (checkbox). As alterações só serão aplicadas quando clicar em 'EXECUTAR AUDITORIA' ou ao usar os botões de limpeza.")

    # Exibir grid (NO_UPDATE)
    grid_response = AgGrid(
        display_df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.NO_UPDATE,
        allow_unsafe_jscode=True,
        theme="alpine",
        height=480,
        fit_columns_on_grid_load=True,
    )

    # -----------------------
    # Botões: EXECUTAR, DESMARCAR TUDO, LIMPAR MARCADAS, LIMPAR TUDO
    # -----------------------
    c_run, c_clear_marked = st.columns([2, 1])
    run = c_run.button("📊 EXECUTAR AUDITORIA (aplicar flags do grid)")

    clear_marked = c_clear_marked.button("🧹 Limpar dados das marcadas")




    # 2) Limpar dados das marcadas (lê o grid atual; se grid vazio, usa master como fallback)
    if clear_marked:
        df_from_grid = pd.DataFrame(grid_response.get("data", []))
        planilhas_marcadas = []
        if not df_from_grid.empty and "Planilha" in df_from_grid.columns:
            planilhas_marcadas = df_from_grid[df_from_grid["Flag"].apply(to_bool_like) == True]["Planilha"].tolist()

        # fallback: usar master flags caso grid não retorne dados válidos
        if not planilhas_marcadas:
            mask_master = st.session_state.au_planilhas_df["Flag"] == True
            if mask_master.any():
                planilhas_marcadas = st.session_state.au_planilhas_df.loc[mask_master, "Planilha"].tolist()

        if planilhas_marcadas:
            mask = st.session_state.au_planilhas_df["Planilha"].isin(planilhas_marcadas)
            cols_limpar = ["Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"]
            for col in cols_limpar:
                st.session_state.au_planilhas_df.loc[mask, col] = ""
            # desmarcar após limpar
            st.session_state.au_planilhas_df.loc[mask, "Flag"] = False
            st.session_state.au_flags_temp = {}
            st.success(f"Dados de {len(planilhas_marcadas)} planilhas limpos.")
            try:
                st.experimental_rerun()
            except Exception:
                pass
        else:
            st.warning("Marque as planilhas no checkbox primeiro!")



    # -----------------------
    # Função: carregar origem faturamento
    # -----------------------
    def carregar_origem_faturamento(d_ini, d_fim):
        try:
            sh_o_fat = gc.open_by_key(ID_PLANILHA_ORIGEM_FAT)
            ws_o_fat = sh_o_fat.worksheet(ABA_ORIGEM_FAT)
            h_o_fat, df_o_fat = get_headers_and_df_raw(ws_o_fat)
            if not df_o_fat.empty:
                df_o_fat = tratar_numericos(df_o_fat, h_o_fat)

            c_dt_o_fat = detect_date_col(h_o_fat) or (h_o_fat[0] if h_o_fat else None)
            if c_dt_o_fat and not df_o_fat.empty:
                df_o_fat["_dt"] = pd.to_datetime(df_o_fat[c_dt_o_fat], dayfirst=True, errors="coerce")
                parsed_pct = df_o_fat["_dt"].notna().mean()
                if parsed_pct == 0:
                    df_o_fat["_dt"] = pd.to_datetime(df_o_fat[c_dt_o_fat], dayfirst=False, errors="coerce")
                df_o_fat["_dt"] = df_o_fat["_dt"].dt.date
                df_o_fat_p = df_o_fat[(df_o_fat["_dt"] >= d_ini) & (df_o_fat["_dt"] <= d_fim)].copy()
            else:
                df_o_fat_p = df_o_fat.copy()

            return h_o_fat, df_o_fat_p
        except Exception as e:
            st.error(f"Erro ao carregar origem de faturamento: {e}")
            return None, None

    # -----------------------
    # Intervalo de datas
    # -----------------------
    if mes_sel == "Todos":
        d_ini, d_fim = date(ano_sel, 1, 1), date(ano_sel, 12, 31)
    else:
        d_ini = date(ano_sel, int(mes_sel), 1)
        d_fim = (date(ano_sel, int(mes_sel), 28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)

    # -----------------------
    # Ao clicar em EXECUTAR: ler o grid, aplicar flags e executar auditoria
    # -----------------------
    if run:
        df_from_grid = pd.DataFrame(grid_response.get("data", []))
        st.session_state.au_flags_temp = {}
        if not df_from_grid.empty and "Planilha" in df_from_grid.columns:
            for _, row in df_from_grid.iterrows():
                pname = row.get("Planilha")
                if pname is None:
                    continue
                st.session_state.au_flags_temp[pname] = to_bool_like(row.get("Flag", False))

        # Aplicar as flags na master
        for i, row in st.session_state.au_planilhas_df.iterrows():
            pname = row["Planilha"]
            st.session_state.au_planilhas_df.at[i, "Flag"] = bool(st.session_state.au_flags_temp.get(pname, False))

        selecionadas = st.session_state.au_planilhas_df[st.session_state.au_planilhas_df["Flag"] == True]
        if selecionadas.empty:
            st.warning("Nenhuma planilha marcada. Marque ao menos uma antes de executar.")
        else:
            h_o_fat, df_o_fat_p = carregar_origem_faturamento(d_ini, d_fim)
            if h_o_fat is None and df_o_fat_p is None:
                st.stop()

            total = len(selecionadas)
            prog = st.progress(0)
            logs = []

            for idx, row in selecionadas.reset_index(drop=True).iterrows():
                sid = row["Planilha_id"]
                pname = row["Planilha"]
                v_o = v_d = v_mp_d = 0.0
                status = "Erro desconhecido"

                # abrir planilha destino
                try:
                    sh_d = gc.open_by_key(sid)
                except Exception as e:
                    status = f"Erro ao abrir planilha ({e})"
                    logs.append(f"{pname}: {status}")
                    st.session_state.au_resultados[sid] = {"Planilha": pname, "Origem": 0.0, "DRE": 0.0, "MP DRE": 0.0, "Dif": 0.0, "Dif MP": 0.0, "Status": status}
                    prog.progress((idx + 1) / total)
                    continue

                # ler codes B2/B3/B4/B5
                try:
                    b2, b3, b4, b5 = read_codes_from_config_sheet(sh_d)
                except Exception:
                    b2, b3, b4, b5 = None, None, None, None

                if not b2:
                    status = "Sem B2 (Config)"
                    logs.append(f"{pname}: {status}")
                    st.session_state.au_resultados[sid] = {"Planilha": pname, "Origem": 0.0, "DRE": 0.0, "MP DRE": 0.0, "Dif": 0.0, "Dif MP": 0.0, "Status": status}
                    prog.progress((idx + 1) / total)
                    continue

                # FATURAMENTO ORIGEM
                try:
                    if h_o_fat and len(h_o_fat) > 5 and (df_o_fat_p is not None) and (not df_o_fat_p.empty):
                        col_b2_fat = h_o_fat[5]
                        df_filter = df_o_fat_p[df_o_fat_p[col_b2_fat].astype(str).str.strip() == str(b2).strip()]
                        if b3 and len(h_o_fat) > 3:
                            col_b3_fat = h_o_fat[3]
                            df_filter = df_filter[df_filter[col_b3_fat].astype(str).str.strip() == str(b3).strip()]
                        if b4 and len(h_o_fat) > 4:
                            col_b4_fat = h_o_fat[4]
                            df_filter = df_filter[df_filter[col_b4_fat].astype(str).str.strip() == str(b4).strip()]
                        if b5 and len(h_o_fat) > 2:
                            col_b5_fat = h_o_fat[2]
                            df_filter = df_filter[df_filter[col_b5_fat].astype(str).str.strip() == str(b5).strip()]
                        if len(h_o_fat) > 6:
                            v_o = float(df_filter[h_o_fat[6]].sum()) if not df_filter.empty else 0.0
                except Exception:
                    v_o = 0.0

                # FATURAMENTO DESTINO (Importado_Fat)
                try:
                    ws_d = sh_d.worksheet("Importado_Fat")
                    h_d, df_d = get_headers_and_df_raw(ws_d)
                    if not df_d.empty:
                        df_d = tratar_numericos(df_d, h_d)

                    c_dt_d = detect_date_col(h_d) or (h_d[0] if h_d else None)
                    if c_dt_d and not df_d.empty:
                        df_d["_dt"] = pd.to_datetime(df_d[c_dt_d], dayfirst=True, errors="coerce")
                        if df_d["_dt"].isna().all():
                            df_d["_dt"] = pd.to_datetime(df_d[c_dt_d], dayfirst=False, errors="coerce")
                        df_d["_dt"] = df_d["_dt"].dt.date
                        df_d_periodo = df_d[(df_d["_dt"] >= d_ini) & (df_d["_dt"] <= d_fim)]
                    else:
                        df_d_periodo = df_d.copy()

                    if len(h_d) > 6 and not df_d_periodo.empty:
                        v_d = float(df_d_periodo[h_d[6]].sum())
                    else:
                        v_d = 0.0
                except Exception:
                    v_d = 0.0

                # MEIO DE PAGAMENTO
                try:
                    ws_mp_d = sh_d.worksheet("Meio de Pagamento")
                    h_mp_d, df_mp_d = get_headers_and_df_raw(ws_mp_d)
                    if not df_mp_d.empty:
                        df_mp_d = tratar_numericos(df_mp_d, h_mp_d)

                    c_dt_mp_d = detect_date_col(h_mp_d) or (h_mp_d[0] if h_mp_d else None)
                    if c_dt_mp_d and not df_mp_d.empty:
                        df_mp_d["_dt"] = pd.to_datetime(df_mp_d[c_dt_mp_d], dayfirst=True, errors="coerce")
                        if df_mp_d["_dt"].isna().all():
                            df_mp_d["_dt"] = pd.to_datetime(df_mp_d[c_dt_mp_d], dayfirst=False, errors="coerce")
                        if "_dt" in df_mp_d.columns:
                            df_mp_d["_dt"] = df_mp_d["_dt"].dt.date
                        df_mp_periodo = df_mp_d[(df_mp_d.get("_dt") >= d_ini) & (df_mp_d.get("_dt") <= d_fim)] if "_dt" in df_mp_d.columns else df_mp_d.copy()
                    else:
                        df_mp_periodo = df_mp_d.copy()

                    v_mp_d = 0.0
                    if len(h_mp_d) > 9 and not df_mp_periodo.empty:
                        col_b2_mp = h_mp_d[8]
                        col_b3_mp = h_mp_d[6]
                        col_b4_mp = h_mp_d[7] if len(h_mp_d) > 7 else None
                        col_b5_mp = h_mp_d[5] if len(h_mp_d) > 5 else None
                        col_val_mp = h_mp_d[9]

                        b2_norm = normalize_code(b2)
                        b3_norm = normalize_code(b3) if b3 else None
                        b4_norm = normalize_code(b4) if b4 else None
                        b5_norm = normalize_code(b5) if b5 else None

                        mask = df_mp_periodo[col_b2_mp].apply(normalize_code) == b2_norm
                        if b3_norm:
                            mask &= df_mp_periodo[col_b3_mp].apply(normalize_code) == b3_norm
                        if b4_norm and col_b4_mp:
                            mask &= df_mp_periodo[col_b4_mp].apply(normalize_code) == b4_norm
                        if b5_norm and col_b5_mp:
                            mask &= df_mp_periodo[col_b5_mp].apply(normalize_code) == b5_norm

                        df_mp_dest_f = df_mp_periodo[mask]

                        if not df_mp_dest_f.empty:
                            v_mp_d = float(df_mp_dest_f[col_val_mp].sum())
                        else:
                            col_val_guess = detect_column_by_keywords(h_mp_d, ["valor", "soma", "total", "amount"])
                            if col_val_guess and col_val_guess in df_mp_periodo.columns:
                                df_guess = df_mp_periodo
                                col_b2_guess = h_mp_d[8] if len(h_mp_d) > 8 else None
                                col_b3_guess = h_mp_d[6] if len(h_mp_d) > 6 else None
                                if col_b2_guess:
                                    df_guess = df_guess[df_guess[col_b2_guess].astype(str).str.strip() == str(b2).strip()]
                                if b3 and col_b3_guess:
                                    df_guess = df_guess[df_guess[col_b3_guess].astype(str).str.strip() == str(b3).strip()]
                                if not df_guess.empty:
                                    v_mp_d = float(df_guess[col_val_guess].sum())
                    else:
                        v_mp_d = 0.0
                except Exception:
                    v_mp_d = 0.0

                # Diferenças e status
                diff = v_o - v_d
                diff_mp = v_d - v_mp_d
                status = "✅ OK" if (abs(diff) < 0.01 and abs(diff_mp) < 0.01) else "❌ Erro"

                # Salvar resultado e atualizar apenas a linha correspondente na master
                st.session_state.au_resultados[sid] = {"Planilha": pname, "Origem": v_o, "DRE": v_d, "MP DRE": v_mp_d, "Dif": diff, "Dif MP": diff_mp, "Status": status}

                mask = st.session_state.au_planilhas_df["Planilha_id"] == sid
                if mask.any():
                    st.session_state.au_planilhas_df.loc[mask, "Origem"] = format_brl(v_o)
                    st.session_state.au_planilhas_df.loc[mask, "DRE"] = format_brl(v_d)
                    st.session_state.au_planilhas_df.loc[mask, "MP DRE"] = format_brl(v_mp_d)
                    st.session_state.au_planilhas_df.loc[mask, "Dif"] = format_brl(diff)
                    st.session_state.au_planilhas_df.loc[mask, "Dif MP"] = format_brl(diff_mp)
                    st.session_state.au_planilhas_df.loc[mask, "Status"] = status
                    # desmarcar a Flag para indicar concluído (se preferir, remova esta linha)
                    st.session_state.au_planilhas_df.loc[mask, "Flag"] = False

                logs.append(f"{pname}: {status if status != "✅ OK" else "OK"}")
                prog.progress((idx + 1) / total)

            # limpar temporário (já aplicado) e mostrar logs
            st.session_state.au_flags_temp = {}
            st.markdown("### Log de processamento")
            st.text("\n".join(logs))
            st.success("Auditoria concluída.")

            # atualizar a tela para mostrar novos dados
            try:
                st.experimental_rerun()
            except Exception:
                pass
    import io

    def to_excel_bytes(df):
        # Copiar e preparar df para exportação
        df_export = df.copy()

        # Remover colunas que não quer no Excel
        cols_to_drop = ["Flag", "Planilha_id", "Status"]
        df_export = df_export.drop(columns=[c for c in cols_to_drop if c in df_export.columns], errors="ignore")

        # Converter colunas de valores para numérico
        valor_cols = ["Origem", "DRE", "MP DRE", "Dif", "Dif MP"]
        for col in valor_cols:
            if col in df_export.columns:
                # Remover "R$ ", pontos e vírgulas para converter corretamente
                df_export[col] = df_export[col].astype(str).str.replace(r"[R$\s\.]", "", regex=True).str.replace(",", ".", regex=False)
                df_export[col] = pd.to_numeric(df_export[col], errors="coerce")

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df_export.to_excel(writer, index=False, sheet_name="Auditoria")
        return output.getvalue()
        # Gerar o arquivo Excel a partir do DataFrame atual da auditoria
    excel_data = to_excel_bytes(st.session_state.au_planilhas_df)

    # Botão para download do Excel
    st.download_button(
        label="📥 Exportar tabela para Excel",
        data=excel_data,
        file_name="auditoria_dre.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
