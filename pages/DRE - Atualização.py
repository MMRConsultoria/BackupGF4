import streamlit as st
import pandas as pd
import json
import re
import io
from datetime import datetime, timedelta, date
from oauth2client.service_account import ServiceAccountCredentials
import gspread

from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode
from st_aggrid.shared import JsCode

try:
    from googleapiclient.discovery import build
except Exception:
    build = None

# ================= BLOQUEIO DE ACESSO – RH (simples, EM-CÓDIGO) =================
USUARIOS_AUTORIZADOS_CONTROLADORIA = {
    "maricelisrossi@gmail.com",
    "alex.komatsu@grupofit.com.br",
}

usuario_logado = st.session_state.get("usuario_logado")

if not usuario_logado:
    st.stop()

if str(usuario_logado).strip().lower() not in {e.lower() for e in USUARIOS_AUTORIZADOS_CONTROLADORIA}:
    st.warning("⛔ Acesso restrito ao CONTROLADORIA")
    st.stop()
# ============================================================================

# ---- CONFIG ----
PASTA_PRINCIPAL_ID = "0B1owaTi3RZnFfm4tTnhfZ2l0VHo4bWNMdHhKS3ZlZzR1ZjRSWWJSSUFxQTJtUExBVlVTUW8"
TARGET_SHEET_NAME = "Configurações Não Apagar"
ID_PLANILHA_ORIGEM_FAT = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
ABA_ORIGEM_FAT = "Fat Sistema Externo"
ID_PLANILHA_ORIGEM_MP = "1GSI291SEeeU9MtOWkGwsKGCGMi_xXMSiQnL_9GhXxfU"
ABA_ORIGEM_MP = "Faturamento Meio Pagamento"

st.set_page_config(page_title="Atualizador DRE", layout="wide")

# ================= PERFUMARIA: CSS CUSTOMIZADO =================
# ================= SUBSTITUA O SEU BLOCO DE CSS POR ESTE =================
st.markdown(
    """
    <style>
    /* 1. REMOVE O ESPAÇO VAZIO NO TOPO DA PÁGINA (HEADER DO STREAMLIT) */
    header {visibility: hidden;}
    .main .block-container {
        padding-top: 3rem !important; /* Dá um espaço de segurança no topo */
        padding-left: 2rem !important;
        padding-right: 2rem !important;
    }
    
    /* 2. ESTILO DO TÍTULO (PARA NÃO CORTAR) */
    .main-title {
        font-size: 32px;
        font-weight: bold;
        color: #1e3d59;
        margin-bottom: 30px;
        display: block; /* Garante que ocupe a linha toda */
        width: 100%;
    }
    
    /* 3. ESTILO DAS ABAS (CORRIGINDO A COR DO TEXTO) */
    button[data-baseweb="tab"] {
        font-size: 18px !important;
        font-weight: bold !important;
        color: #555 !important;
        background-color: #f0f2f6 !important;
        border-radius: 8px 8px 0px 0px !important;
        padding: 10px 25px !important;
    }
    
    button[data-baseweb="tab"][aria-selected="true"] {
        background-color: #0066cc !important;
    }

    button[data-baseweb="tab"][aria-selected="true"] p {
        color: white !important; /* Texto branco na aba azul */
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# O Título agora com uma margem melhor
st.markdown('<div class="main-title">📈 Atualizar DRE</div>', unsafe_allow_html=True)
# =========================================================================
# --- INICIO: CSS ESPECÍFICO PARA A TABELA DE AUDITORIA ---
st.markdown("""
<style>
#auditoria .ag-theme-alpine .ag-root-wrapper {
    border: 1px solid #c9d2da !important;
    border-radius: 6px !important;
    overflow: hidden !important;
}
#auditoria .ag-theme-alpine .ag-header-cell,
#auditoria .ag-theme-alpine .ag-cell {
    border-right: 1px solid #e6edf3 !important;
    border-bottom: 1px solid #e6edf3 !important;
}
#auditoria .ag-theme-alpine .ag-header {
    border-bottom: 2px solid #cfd8e3 !important;
    background: #f7fafc !important;
}
#auditoria .ag-theme-alpine .ag-cell,
#auditoria .ag-theme-alpine .ag-header-cell {
    padding: 6px 8px !important;
}
#auditoria .ag-theme-alpine .ag-header-cell-label {
    color: #203040 !important;
    font-weight: 600 !important;
}
#auditoria .ag-theme-alpine .ag-cell-focus {
    outline: none !important;
    box-shadow: none !important;
}
#auditoria .ag-theme-alpine .ag-center-cols-container .ag-row .ag-cell:last-child {
    border-right: 1px solid #e6edf3 !important;
}
#auditoria .ag-theme-alpine .ag-body-viewport {
    background: #ffffff !important;
}
</style>
""", unsafe_allow_html=True)
# --- FIM: CSS ESPECÍFICO PARA A TABELA DE AUDITORIA ---
# ---- AUTENTICAÇÃO ----
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

# ---- HELPERS GLOBAIS (Lógica Intacta) ----
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
    if not headers: return None
    if len(headers) > 0 and "data" in headers[0].lower():
        return headers[0]
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
            try:
                df[col_name] = df[col_name].apply(_parse_currency_like).fillna(0.0)
            except Exception:
                pass
    return df

def format_brl(val):
    try: return f"R$ {float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return val

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

# ---- TABS ESTILIZADAS ----
tab_atual, tab_audit  = st.tabs(["🔄 Atualização","🔍 Auditoria" ])

# -----------------------------
# ABA: ATUALIZAÇÃO (Lógica Intacta)
# -----------------------------
with tab_atual:
    st.subheader("Atualizar DRE Google Sheets")
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        data_de = st.date_input("De", value=date.today() - timedelta(days=30), key="at_de")
    with col_d2:
        data_ate = st.date_input("Até", value=date.today(), key="at_ate")

    try:
        pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
        map_p = {p["name"]: p["id"] for p in pastas_fech}
        p_sel = st.selectbox("Pasta principal:", options=list(map_p.keys()), key="at_p")
        subpastas = list_child_folders(drive_service, map_p[p_sel])
        map_s = {s["name"]: s["id"] for s in subpastas}
        s_sel = st.multiselect("Subpastas:", options=list(map_s.keys()), default=[], key="at_s")
        s_ids = [map_s[n] for n in s_sel]
    except Exception:
        st.error("Erro ao listar pastas."); st.stop()

    if not s_ids:
        st.info("Selecione as subpastas.")
    else:
        planilhas = list_spreadsheets_in_folders(drive_service, s_ids)
        if not planilhas:
            st.warning("Nenhuma planilha.")
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
            with col_t1:
                edit_esq = st.data_editor(df_list.iloc[:meio], key="at_t1", use_container_width=True, column_config=config, hide_index=True)
            with col_t2:
                edit_dir = st.data_editor(df_list.iloc[meio:], key="at_t2", use_container_width=True, column_config=config, hide_index=True)

            if st.button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True):
                df_final_edit = pd.concat([edit_esq, edit_dir], ignore_index=True)
                df_marcadas = df_final_edit[(df_final_edit["Desconto"]) | (df_final_edit["Meio Pagamento"]) | (df_final_edit["Faturamento"])].copy()
                if df_marcadas.empty:
                    st.warning("Nada marcado.")
                    st.stop()

                status_placeholder = st.empty()
                status_placeholder.info("Carregando dados de origem...")

                try:
                    sh_orig_fat = gc.open_by_key(ID_PLANILHA_ORIGEM_FAT)
                    ws_orig_fat = sh_orig_fat.worksheet(ABA_ORIGEM_FAT)
                    h_orig_fat, df_orig_fat = get_headers_and_df_raw(ws_orig_fat)
                    c_dt_fat = detect_date_col(h_orig_fat)
                    if c_dt_fat:
                        df_orig_fat["_dt"] = pd.to_datetime(df_orig_fat[c_dt_fat], dayfirst=True, errors="coerce").dt.date
                        df_orig_fat_f = df_orig_fat[(df_orig_fat["_dt"] >= data_de) & (df_orig_fat["_dt"] <= data_ate)].copy()
                    else:
                        df_orig_fat_f = df_orig_fat.copy()
                except Exception as e:
                    st.error(f"Erro origem Fat: {e}"); st.stop()

                try:
                    sh_orig_mp = gc.open_by_key(ID_PLANILHA_ORIGEM_MP)
                    ws_orig_mp = sh_orig_mp.worksheet(ABA_ORIGEM_MP)
                    h_orig_mp, df_orig_mp = get_headers_and_df_raw(ws_orig_mp)
                    c_dt_mp = detect_date_col(h_orig_mp)
                    if c_dt_mp:
                        df_orig_mp["_dt"] = pd.to_datetime(df_orig_mp[c_dt_mp], dayfirst=True, errors="coerce").dt.date
                        df_orig_mp_f = df_orig_mp[(df_orig_mp["_dt"] >= data_de) & (df_orig_mp["_dt"] <= data_ate)].copy()
                    else:
                        df_orig_mp_f = df_orig_mp.copy()
                except Exception as e:
                    st.error(f"Erro origem MP: {e}"); st.stop()

                prog = st.progress(0)
                logs = []
                log_placeholder = st.empty()

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

                        lojas_filtro = []
                        if b3: lojas_filtro.append(str(b3).strip())
                        if b4: lojas_filtro.append(str(b4).strip())
                        if b5: lojas_filtro.append(str(b5).strip())

                        if row["Faturamento"]:
                            df_ins = df_orig_fat_f.copy()
                            if len(h_orig_fat) > 5:
                                c_b2 = h_orig_fat[5]
                                df_ins = df_ins[df_ins[c_b2].astype(str).str.strip() == b2]
                            if lojas_filtro and not df_ins.empty:
                                if len(h_orig_fat) > 3:
                                    c_loja = h_orig_fat[3]
                                    df_ins = df_ins[df_ins[c_loja].astype(str).str.strip().isin(lojas_filtro)]
                            if not df_ins.empty:
                                try:
                                    ws_dest = sh_dest.worksheet("Importado_Fat")
                                except:
                                    ws_dest = sh_dest.add_worksheet("Importado_Fat", 1000, 30)
                                h_dest, df_dest = get_headers_and_df_raw(ws_dest)
                                if df_dest.empty:
                                    df_f_ws, h_f = df_ins, h_orig_fat
                                else:
                                    c_dt_d = detect_date_col(h_dest)
                                    if c_dt_d:
                                        df_dest["_dt"] = pd.to_datetime(df_dest[c_dt_d], dayfirst=True, errors="coerce").dt.date
                                        rem = (df_dest["_dt"] >= data_de) & (df_dest["_dt"] <= data_ate)
                                    else:
                                        rem = pd.Series([False] * len(df_dest))
                                    if len(h_orig_fat) > 5 and c_b2 in df_dest.columns:
                                        rem &= (df_dest[c_b2].astype(str).str.strip() == b2)
                                    df_f_ws = pd.concat([df_dest.loc[~rem], df_ins], ignore_index=True)
                                    h_f = h_dest if h_dest else h_orig_fat
                                if "_dt" in df_f_ws.columns: df_f_ws = df_f_ws.drop(columns=["_dt"])
                                send = df_f_ws[h_f].fillna("")
                                ws_dest.clear()
                                ws_dest.update("A1", [h_f] + send.values.tolist(), value_input_option="USER_ENTERED")
                                logs.append(f"{row['Planilha']}: Fat OK.")
                            else:
                                logs.append(f"{row['Planilha']}: Fat Sem dados.")

                        if row["Meio Pagamento"]:
                            df_ins_mp = df_orig_mp_f.copy()
                            if len(h_orig_mp) > 8:
                                c_b2_mp = h_orig_mp[8]
                                df_ins_mp = df_ins_mp[df_ins_mp[c_b2_mp].astype(str).str.strip() == b2]
                            if lojas_filtro and not df_ins_mp.empty:
                                if len(h_orig_mp) > 6:
                                    c_loja_mp = h_orig_mp[6]
                                    df_ins_mp = df_ins_mp[df_ins_mp[c_loja_mp].astype(str).str.strip().isin(lojas_filtro)]
                            if not df_ins_mp.empty:
                                try:
                                    ws_dest_mp = sh_dest.worksheet("Meio de Pagamento")
                                except:
                                    ws_dest_mp = sh_dest.add_worksheet("Meio de Pagamento", 1000, 30)
                                h_dest_mp, df_dest_mp = get_headers_and_df_raw(ws_dest_mp)
                                if df_dest_mp.empty:
                                    df_f_mp, h_f_mp = df_ins_mp, h_orig_mp
                                else:
                                    c_dt_d_mp = detect_date_col(h_dest_mp)
                                    if c_dt_d_mp:
                                        df_dest_mp["_dt"] = pd.to_datetime(df_dest_mp[c_dt_d_mp], dayfirst=True, errors="coerce").dt.date
                                        rem_mp = (df_dest_mp["_dt"] >= data_de) & (df_dest_mp["_dt"] <= data_ate)
                                    else:
                                        rem_mp = pd.Series([False] * len(df_dest_mp))
                                    if len(h_orig_mp) > 8 and c_b2_mp in df_dest_mp.columns:
                                        rem_mp &= (df_dest_mp[c_b2_mp].astype(str).str.strip() == b2)
                                    df_f_mp = pd.concat([df_dest_mp.loc[~rem_mp], df_ins_mp], ignore_index=True)
                                    h_f_mp = h_dest_mp if h_dest_mp else h_orig_mp
                                if "_dt" in df_f_mp.columns: df_f_mp = df_f_mp.drop(columns=["_dt"])
                                send_mp = df_f_mp[h_f_mp].fillna("")
                                ws_dest_mp.clear()
                                ws_dest_mp.update("A1", [h_f_mp] + send_mp.values.tolist(), value_input_option="USER_ENTERED")
                                logs.append(f"{row['Planilha']}: MP OK.")
                            else:
                                logs.append(f"{row['Planilha']}: MP Sem dados.")
                    except Exception as e:
                        logs.append(f"{row['Planilha']}: Erro {e}")
                    prog.progress((i+1)/total)
                    log_placeholder.text("\n".join(logs))
                st.success("Concluído!")

# -----------------------------
# -----------------------------
# ABA: AUDITORIA (COMPLETA E CORRIGIDA)
# -----------------------------
with tab_audit:
    st.subheader("Auditoria Faturamento X Meio de Pagamento")
    try:
        pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
        if not pastas_fech:
            st.error("Nenhuma pasta de fechamento encontrada.")
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

    c1, c2 = st.columns(2)
    with c1:
        ano_sel = st.selectbox("Ano:", list(range(2020, date.today().year + 1)), index=max(0, date.today().year - 2020), key="au_ano")
    with c2:
        mes_sel = st.selectbox("Mês (Opcional):", ["Todos"] + list(range(1, 13)), key="au_mes")

    need_reload = ("au_last_subpastas" not in st.session_state) or (st.session_state.get("au_last_subpastas") != s_ids_audit)
    if need_reload:
        try:
            planilhas = list_spreadsheets_in_folders(drive_service, s_ids_audit)
        except Exception as e:
            st.error(f"Erro ao listar planilhas: {e}")
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

    if "au_planilhas_df" not in st.session_state:
        st.session_state.au_planilhas_df = pd.DataFrame(columns=["Planilha", "Flag", "Planilha_id", "Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"])
    
    df_table = st.session_state.au_planilhas_df.copy()
    expected_cols = ["Planilha", "Planilha_id", "Flag", "Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"]
    for c in expected_cols:
        if c not in df_table.columns:
            df_table[c] = False if c == "Flag" else ""

    display_df = df_table[expected_cols].copy()

    gb = GridOptionsBuilder.from_dataframe(display_df)
    gb.configure_column("Planilha", headerName="Planilha", editable=False, width=420)
    gb.configure_column("Planilha_id", headerName="Planilha_id", editable=False, hide=True)
    gb.configure_column("Flag", editable=True, cellEditor="agCheckboxCellEditor", cellRenderer="agCheckboxCellRenderer", width=80)
    grid_options = gb.build()

    st.markdown('<div id="auditoria">', unsafe_allow_html=True)    
    grid_response = AgGrid(
        display_df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        allow_unsafe_jscode=True,
        theme='alpine',
        height=420,
        fit_columns_on_grid_load=True,
    )
    st.markdown('</div>', unsafe_allow_html=True)

    # --- PREPARAÇÃO DE VARIÁVEIS PARA OS BOTÕES ---
    currency_cols = ["Origem", "DRE", "MP DRE", "Dif", "Dif MP"]
    cols_for_excel = ["Planilha"] + [c for c in currency_cols if c in st.session_state.au_planilhas_df.columns]
    df_para_excel_btn = st.session_state.au_planilhas_df[cols_for_excel].copy()
    is_empty_btn = df_para_excel_btn.empty

    def _to_numeric_or_nan(x):
        if pd.isna(x) or str(x).strip() == "": return pd.NA
        if isinstance(x, (int, float)): return float(x)
        try:
            n = _parse_currency_like(x)
            if n is not None: return float(n)
        except: pass
        try: return float(str(x).replace(".", "").replace(",", "."))
        except: return pd.NA

    # --- BOTÕES EM 4 COLUNAS ---
    c1, c2, c3, c4 = st.columns([1, 1, 1, 1])
    
    with c1:
        executar_clicado = st.button("📊 Atualizar", key="au_exec", use_container_width=True)
    
    with c2:
        limpar_clicadas = st.button("🧹 Limpar marcadas", key="au_limpar", use_container_width=True)
    
    # Preparação do Excel da Auditoria
    processed_btn = b""
    if not is_empty_btn:
        df_to_write = df_para_excel_btn.copy()
        for col in currency_cols:
            if col in df_to_write.columns:
                df_to_write[col] = df_to_write[col].apply(_to_numeric_or_nan)
        output_btn = io.BytesIO()
        with pd.ExcelWriter(output_btn, engine="xlsxwriter") as writer:
            df_to_write.to_excel(writer, index=False, sheet_name="Auditoria")
        processed_btn = output_btn.getvalue()

    with c3:
        st.download_button(
            label="⬇️ Excel Auditoria",
            data=processed_btn,
            file_name=f"auditoria_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            disabled=is_empty_btn,
            key="au_download"
        )
    
    with c4:
        verificar_btn = st.button("🔎 Verificar Lojas", use_container_width=True, key="au_verif_simple")

    # --- LÓGICA: VERIFICAR LOJAS FALTANTES ---
    if verificar_btn:
        placeholder_msg = st.empty()
        placeholder_msg.info("Verificando Tabela Empresa vs Subpastas...")
        try:
            sh_origem = gc.open_by_key(ID_PLANILHA_ORIGEM_FAT)
            ws_empresa = sh_origem.worksheet("Tabela Empresa")
            dados_empresa = ws_empresa.get_all_values()
            nomes_codigos = []
            for r in dados_empresa[1:]:
                nome = r[0].strip() if len(r) > 0 else ""
                codigo_raw = r[2] if len(r) > 2 else ""
                if str(codigo_raw).strip() != "":
                    nomes_codigos.append((nome, normalize_code(codigo_raw)))

            planilhas_pasta = st.session_state.au_planilhas_df.copy()
            mapa_codigos = {}
            prog = st.progress(0)
            total = len(planilhas_pasta)
            for i, prow in planilhas_pasta.reset_index(drop=True).iterrows():
                sid = prow.get("Planilha_id")
                try:
                    if sid:
                        sh_dest = gc.open_by_key(sid)
                        _, b3, b4, b5 = read_codes_from_config_sheet(sh_dest)
                        for val in (b3, b4, b5):
                            if val: mapa_codigos.setdefault(normalize_code(val), []).append(prow["Planilha"])
                except: pass
                prog.progress((i + 1) / total)

            relatorio = []
            for nome, cod in nomes_codigos:
                vinc = mapa_codigos.get(cod, [])
                relatorio.append({
                    "Nome Empresa": nome, "Código": cod,
                    "Status": "✅ OK" if vinc else "❌ FALTANDO",
                    "Planilhas": ", ".join(vinc) if vinc else "NENHUMA"
                })
            df_rel = pd.DataFrame(relatorio)
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                df_rel.to_excel(writer, index=False, sheet_name="Relatorio")
            
            placeholder_msg.success("Verificação concluída!")
            st.download_button("⬇️ Baixar Relatório de Lojas", buf.getvalue(), "relatorio_lojas.xlsx", use_container_width=True)
        except Exception as e:
            st.error(f"Erro na verificação: {e}")

    # --- LÓGICA: LIMPAR MARCADAS ---
    if limpar_clicadas:
        df_grid_now = pd.DataFrame(grid_response.get("data", []))
        if not df_grid_now.empty:
            planilhas_marcadas = df_grid_now[df_grid_now["Flag"].apply(to_bool_like) == True]["Planilha"].tolist()
            if planilhas_marcadas:
                mask = st.session_state.au_planilhas_df["Planilha"].isin(planilhas_marcadas)
                for col in ["Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"]:
                    st.session_state.au_planilhas_df.loc[mask, col] = ""
                st.session_state.au_planilhas_df.loc[mask, "Flag"] = False
                st.rerun()

    # --- LÓGICA: ATUALIZAR (PROCESSAMENTO) ---
    if executar_clicado:
        df_grid = pd.DataFrame(grid_response.get("data", []))
        selecionadas = df_grid[df_grid["Flag"].apply(to_bool_like) == True].copy()
        if selecionadas.empty:
            st.warning("Marque ao menos uma planilha.")
        else:
            # Define período
            if mes_sel == "Todos":
                d_ini, d_fim = date(ano_sel, 1, 1), date(ano_sel, 12, 31)
            else:
                d_ini = date(ano_sel, int(mes_sel), 1)
                d_fim = (date(ano_sel, int(mes_sel), 28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)

            # Carrega Origem Faturamento
            try:
                sh_o_fat = gc.open_by_key(ID_PLANILHA_ORIGEM_FAT)
                ws_o_fat = sh_o_fat.worksheet(ABA_ORIGEM_FAT)
                h_o_fat, df_o_fat = get_headers_and_df_raw(ws_o_fat)
                if not df_o_fat.empty:
                    df_o_fat = tratar_numericos(df_o_fat, h_o_fat)
                    c_dt_o = detect_date_col(h_o_fat)
                    if c_dt_o:
                        df_o_fat["_dt"] = pd.to_datetime(df_o_fat[c_dt_o], dayfirst=True, errors="coerce").dt.date
                        df_o_fat_p = df_o_fat[(df_o_fat["_dt"] >= d_ini) & (df_o_fat["_dt"] <= d_fim)].copy()
                    else: df_o_fat_p = df_o_fat.copy()
            except Exception as e:
                st.error(f"Erro origem: {e}"); st.stop()

            prog = st.progress(0)
            logs = []
            for idx, row in selecionadas.reset_index(drop=True).iterrows():
                sid = row.get("Planilha_id")
                pname = row.get("Planilha", "Sem Nome")
                v_o = v_d = v_mp = 0.0
                try:
                    sh_d = gc.open_by_key(sid)
                    b2, b3, b4, b5 = read_codes_from_config_sheet(sh_d)
                    lojas_audit = [normalize_code(x) for x in [b3, b4, b5] if x]

                    # Valor Origem
                    if not df_o_fat_p.empty and len(h_o_fat) > 6:
                        df_f = df_o_fat_p[df_o_fat_p[h_o_fat[5]].astype(str).str.strip() == str(b2).strip()]
                        if lojas_audit: df_f = df_f[df_f[h_o_fat[3]].apply(normalize_code).isin(lojas_audit)]
                        v_o = float(df_f[h_o_fat[6]].sum())

                    # Valor DRE (Importado_Fat)
                    ws_d = sh_d.worksheet("Importado_Fat")
                    h_d, df_d = get_headers_and_df_raw(ws_d)
                    if not df_d.empty:
                        df_d = tratar_numericos(df_d, h_d)
                        c_dt_d = detect_date_col(h_d) or h_d[0]
                        df_d["_dt"] = pd.to_datetime(df_d[c_dt_d], dayfirst=True, errors="coerce").dt.date
                        df_dp = df_d[(df_d["_dt"] >= d_ini) & (df_d["_dt"] <= d_fim)]
                        v_d = float(df_dp[h_d[6]].sum()) if len(h_d) > 6 else 0.0

                    # Valor MP (Lógica Restaurada)
                    # --- BLOCO CORRIGIDO: Valor MP (Meio de Pagamento) ---
                    ws_mp = sh_d.worksheet("Meio de Pagamento")
                    h_mp, df_mp = get_headers_and_df_raw(ws_mp)
                    if not df_mp.empty:
                        df_mp = tratar_numericos(df_mp, h_mp)
                        # Força a data a ser lida da coluna A (índice 0)
                        c_dt_mp = h_mp[0]
                        df_mp["_dt"] = pd.to_datetime(df_mp[c_dt_mp], dayfirst=True, errors="coerce").dt.date
                        
                        # Filtra o período
                        df_mpp = df_mp[(df_mp["_dt"] >= d_ini) & (df_mp["_dt"] <= d_fim)].copy()
                        
                        if not df_mpp.empty and len(h_mp) > 9:
                            # Coluna 8: B2 (Código Empresa)
                            # Coluna 6: Loja (B3, B4, B5)
                            # Coluna 9: Valor
                            
                            # Criamos a máscara comparando sempre com normalize_code
                            mask_b2 = df_mpp[h_mp[8]].apply(normalize_code) == normalize_code(b2)
                            
                            if lojas_audit:
                                mask_loja = df_mpp[h_mp[6]].apply(normalize_code).isin(lojas_audit)
                                mask_final = mask_b2 & mask_loja
                            else:
                                mask_final = mask_b2
                                
                            # Soma os valores que batem no filtro
                            v_mp = float(df_mpp.loc[mask_final, h_mp[9]].sum())

                    diff, diff_mp = v_o - v_d, v_d - v_mp
                    status = "✅ OK" if (abs(diff) < 0.01 and abs(diff_mp) < 0.01) else "❌ Erro"
                    
                    mask_m = st.session_state.au_planilhas_df["Planilha_id"] == sid
                    st.session_state.au_planilhas_df.loc[mask_m, ["Origem","DRE","MP DRE","Dif","Dif MP","Status"]] = [
                        format_brl(v_o), format_brl(v_d), format_brl(v_mp), format_brl(diff), format_brl(diff_mp), status
                    ]
                    st.session_state.au_planilhas_df.loc[mask_m, "Flag"] = False
                    logs.append(f"{pname}: {status}")
                except Exception as e: logs.append(f"{pname}: Erro {e}")
                prog.progress((idx + 1) / len(selecionadas))
            st.success("Concluído!"); st.rerun()
