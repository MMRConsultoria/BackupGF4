
import streamlit as st
import pandas as pd
import json
import re
import io
from datetime import datetime, timedelta, date
from oauth2client.service_account import ServiceAccountCredentials
import gspread

try:
    from googleapiclient.discovery import build
except Exception:
    build = None

# ---- CONFIG ----
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

st.title("Atualizador DRE - Multi-Lojas")

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

# ---- HELPERS ----
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
    # Prioriza a coluna A (índice 0) se ela tiver "data" no nome
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
            df[col_name] = df[col_name].apply(_parse_currency_like).fillna(0.0)
    return df

def format_brl(val):
    try: return f"R$ {float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except: return val

# ---- TABS ----
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

                status_placeholder = st.empty()
                status_placeholder.info("Carregando dados de origem...")

                # Carregar Origem Faturamento
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

                # Carregar Origem Meio Pagamento
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

                        # --- ATUALIZAR FATURAMENTO ---
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
                                try: ws_dest = sh_dest.worksheet("Importado_Fat")
                                except: ws_dest = sh_dest.add_worksheet("Importado_Fat", 1000, 30)
                                h_dest, df_dest = get_headers_and_df_raw(ws_dest)
                                if df_dest.empty:
                                    df_f_ws, h_f = df_ins, h_orig_fat
                                else:
                                    c_dt_d = detect_date_col(h_dest)
                                    if c_dt_d:
                                        df_dest["_dt"] = pd.to_datetime(df_dest[c_dt_d], dayfirst=True, errors="coerce").dt.date
                                        rem = (df_dest["_dt"] >= data_de) & (df_dest["_dt"] <= data_ate)
                                    else: rem = pd.Series([False] * len(df_dest))
                                    if len(h_orig_fat) > 5 and c_b2 in df_dest.columns:
                                        rem &= (df_dest[c_b2].astype(str).str.strip() == b2)
                                    df_f_ws = pd.concat([df_dest.loc[~rem], df_ins], ignore_index=True)
                                    h_f = h_dest if h_dest else h_orig_fat
                                if "_dt" in df_f_ws.columns: df_f_ws = df_f_ws.drop(columns=["_dt"])
                                send = df_f_ws[h_f].fillna("")
                                ws_dest.clear()
                                ws_dest.update("A1", [h_f] + send.values.tolist(), value_input_option="USER_ENTERED")
                                logs.append(f"{row['Planilha']}: Fat OK.")
                            else: logs.append(f"{row['Planilha']}: Fat Sem dados.")

                        # --- ATUALIZAR MEIO DE PAGAMENTO ---
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
                                try: ws_dest_mp = sh_dest.worksheet("Meio de Pagamento")
                                except: ws_dest_mp = sh_dest.add_worksheet("Meio de Pagamento", 1000, 30)
                                h_dest_mp, df_dest_mp = get_headers_and_df_raw(ws_dest_mp)
                                if df_dest_mp.empty:
                                    df_f_mp, h_f_mp = df_ins_mp, h_orig_mp
                                else:
                                    c_dt_d_mp = detect_date_col(h_dest_mp)
                                    if c_dt_d_mp:
                                        df_dest_mp["_dt"] = pd.to_datetime(df_dest_mp[c_dt_d_mp], dayfirst=True, errors="coerce").dt.date
                                        rem_mp = (df_dest_mp["_dt"] >= data_de) & (df_dest_mp["_dt"] <= data_ate)
                                    else: rem_mp = pd.Series([False] * len(df_dest_mp))
                                    if len(h_orig_mp) > 8 and c_b2_mp in df_dest_mp.columns:
                                        rem_mp &= (df_dest_mp[c_b2_mp].astype(str).str.strip() == b2)
                                    df_f_mp = pd.concat([df_dest_mp.loc[~rem_mp], df_ins_mp], ignore_index=True)
                                    h_f_mp = h_dest_mp if h_dest_mp else h_orig_mp
                                if "_dt" in df_f_mp.columns: df_f_mp = df_f_mp.drop(columns=["_dt"])
                                send_mp = df_f_mp[h_f_mp].fillna("")
                                ws_dest_mp.clear()
                                ws_dest_mp.update("A1", [h_f_mp] + send_mp.values.tolist(), value_input_option="USER_ENTERED")
                                logs.append(f"{row['Planilha']}: MP OK.")
                            else: logs.append(f"{row['Planilha']}: MP Sem dados.")
                    except Exception as e:
                        logs.append(f"{row['Planilha']}: Erro {e}")
                    prog.progress((i+1)/total)
                    log_placeholder.text("\n".join(logs))
                st.success("Concluído!")

# Botões lado a lado: Executar | Limpar marcadas | Download
    col_btn1, col_btn2, col_btn3, _ = st.columns([2, 2, 1, 6])

    with col_btn1:
        executar_clicado = st.button("📊 EXECUTAR AUDITORIA", key="au_exec", use_container_width=True)

    with col_btn2:
        limpar_clicadas = st.button("🧹 Limpar marcadas", key="au_limpar", use_container_width=True)

    # preparar dados para o botão de download (pequeno, ao lado dos botões)
    # garante que a estrutura exista
    if "au_planilhas_df" not in st.session_state:
        st.session_state.au_planilhas_df = pd.DataFrame(
            columns=["Planilha", "Flag", "Planilha_id", "Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"]
        )

    with col_btn3:
        df_para_excel_btn = st.session_state.au_planilhas_df[["Planilha", "Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"]].copy()
        is_empty_btn = df_para_excel_btn.empty
        if not is_empty_btn:
            output_btn = io.BytesIO()
            with pd.ExcelWriter(output_btn, engine="xlsxwriter") as writer:
                df_para_excel_btn.to_excel(writer, index=False, sheet_name="Auditoria")
            processed_btn = output_btn.getvalue()
        else:
            processed_btn = b""

        st.download_button(
            label="⬇️ Excel",
            data=processed_btn,
            file_name=f"auditoria_dre_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            disabled=is_empty_btn,
            key="au_download"
        )

    # --- Lógica de Limpar Marcadas (executa quando usuário clica no botão Limpar marcadas) ---
    if limpar_clicadas:
        # obter os dados visíveis no grid
        df_grid_now = pd.DataFrame(grid_response.get("data", []))
        planilhas_marcadas = []
        if not df_grid_now.empty and "Planilha" in df_grid_now.columns:
            planilhas_marcadas = df_grid_now[df_grid_now["Flag"].apply(to_bool_like) == True]["Planilha"].tolist()

        # fallback: usar master flags caso grid não retorne dados válidos
        if not planilhas_marcadas:
            if "au_planilhas_df" in st.session_state:
                mask_master = st.session_state.au_planilhas_df["Flag"] == True
                if mask_master.any():
                    planilhas_marcadas = st.session_state.au_planilhas_df.loc[mask_master, "Planilha"].tolist()

        if not planilhas_marcadas:
            st.warning("Marque as planilhas no checkbox primeiro!")
        else:
            mask = st.session_state.au_planilhas_df["Planilha"].isin(planilhas_marcadas)
            cols_limpar = ["Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"]
            for col in cols_limpar:
                st.session_state.au_planilhas_df.loc[mask, col] = ""
            # desmarcar após limpar
            st.session_state.au_planilhas_df.loc[mask, "Flag"] = False
            st.success(f"Dados de {len(planilhas_marcadas)} planilhas limpos.")
            try:
                st.experimental_rerun()
            except Exception:
                st.info("As flags foram limpas. Atualize a página se necessário para ver a alteração.")

    # --- Lógica de EXECUTAR (substitui o antigo if st.button(...) por este gatilho) ---
    if executar_clicado:
        df_grid = pd.DataFrame(grid_response.get("data", []))
        selecionadas = df_grid[df_grid["Flag"].apply(to_bool_like) == True]

        if selecionadas.empty:
            st.warning("Marque as planilhas.")
        else:
            # (aqui entra o seu bloco original de processamento)
            if mes_sel == "Todos":
                d_ini, d_fim = date(ano_sel, 1, 1), date(ano_sel, 12, 31)
            else:
                d_ini = date(ano_sel, int(mes_sel), 1)
                d_fim = (date(ano_sel, int(mes_sel), 28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)

            sh_o_fat = gc.open_by_key(ID_PLANILHA_ORIGEM_FAT)
            ws_o_fat = sh_o_fat.worksheet(ABA_ORIGEM_FAT)
            h_o_fat, df_o_fat = get_headers_and_df_raw(ws_o_fat)
            df_o_fat = tratar_numericos(df_o_fat, h_o_fat)
            c_dt_o = detect_date_col(h_o_fat)
            if c_dt_o:
                df_o_fat["_dt"] = pd.to_datetime(df_o_fat[c_dt_o], dayfirst=True, errors="coerce").dt.date
                df_o_fat_p = df_o_fat[(df_o_fat["_dt"] >= d_ini) & (df_o_fat["_dt"] <= d_fim)].copy()
            else:
                df_o_fat_p = df_o_fat.copy()

            prog = st.progress(0)
            results_excel = []
            n = len(selecionadas)
            for idx, row in selecionadas.reset_index(drop=True).iterrows():
                sid = row["Planilha_id"]
                try:
                    sh_d = gc.open_by_key(sid)
                except:
                    continue

                b2, b3, b4, b5 = read_codes_from_config_sheet(sh_d)
                lojas_audit = []
                if b3: lojas_audit.append(normalize_code(b3))
                if b4: lojas_audit.append(normalize_code(b4))
                if b5: lojas_audit.append(normalize_code(b5))

                # Valor Origem
                try:
                    df_f = df_o_fat_p[df_o_fat_p[h_o_fat[5]].astype(str).str.strip() == b2]
                    if lojas_audit:
                        df_f = df_f[df_f[h_o_fat[3]].apply(normalize_code).isin(lojas_audit)]
                    v_o = float(df_f[h_o_fat[6]].sum())
                except:
                    v_o = 0.0

                # Valor Destino (DRE)
                try:
                    ws_d = sh_d.worksheet("Importado_Fat")
                    h_d, df_d = get_headers_and_df_raw(ws_d)
                    df_d = tratar_numericos(df_d, h_d)
                    c_dt_d = detect_date_col(h_d)
                    if c_dt_d:
                        df_d["_dt"] = pd.to_datetime(df_d[c_dt_d], dayfirst=True, errors="coerce").dt.date
                        df_d_p = df_d[(df_d["_dt"] >= d_ini) & (df_d["_dt"] <= d_fim)]
                    else:
                        df_d_p = df_d
                    v_d = float(df_d_p[h_d[6]].sum())
                except:
                    v_d = 0.0

                # Valor MP (mantém a lógica que você já ajustou)
                try:
                    ws_mp = sh_d.worksheet("Meio de Pagamento")
                    h_mp, df_mp = get_headers_and_df_raw(ws_mp)
                    if not df_mp.empty:
                        df_mp = tratar_numericos(df_mp, h_mp)

                    c_dt_mp = (h_mp[0] if h_mp and len(h_mp) > 0 else None)
                    if not c_dt_mp:
                        c_dt_mp = detect_date_col(h_mp)

                    if c_dt_mp and not df_mp.empty:
                        df_mp["_dt"] = pd.to_datetime(df_mp[c_dt_mp], dayfirst=True, errors="coerce")
                        if df_mp["_dt"].isna().all():
                            df_mp["_dt"] = pd.to_datetime(df_mp[c_dt_mp], dayfirst=False, errors="coerce")
                        df_mp["_dt"] = df_mp["_dt"].dt.date
                        df_mp_periodo = df_mp[(df_mp["_dt"] >= d_ini) & (df_mp["_dt"] <= d_fim)]
                    else:
                        df_mp_periodo = df_mp.copy()

                    v_mp = 0.0
                    if not df_mp_periodo.empty:
                        col_b2_mp = h_mp[8] if len(h_mp) > 8 else None
                        col_loja_mp = h_mp[6] if len(h_mp) > 6 else None
                        col_val_mp = h_mp[9] if len(h_mp) > 9 else None

                        ok_b2 = (col_b2_mp in df_mp_periodo.columns) if col_b2_mp else False
                        ok_loja = (col_loja_mp in df_mp_periodo.columns) if col_loja_mp else False
                        ok_val = (col_val_mp in df_mp_periodo.columns) if col_val_mp else False

                        if ok_b2:
                            b2_norm = normalize_code(b2)
                            mask = df_mp_periodo[col_b2_mp].apply(normalize_code) == b2_norm
                            if lojas_audit and ok_loja:
                                mask &= df_mp_periodo[col_loja_mp].apply(normalize_code).isin(lojas_audit)

                            df_mp_dest_f = df_mp_periodo[mask]
                            if not df_mp_dest_f.empty and ok_val:
                                v_mp = float(df_mp_dest_f[col_val_mp].sum())
                            else:
                                col_val_guess = None
                                try:
                                    col_val_guess = detect_column_by_keywords(h_mp, ["valor", "soma", "total", "amount", "receita", "vl"])
                                except Exception:
                                    col_val_guess = None
                                if col_val_guess and col_val_guess in df_mp_periodo.columns:
                                    df_guess = df_mp_periodo.copy()
                                    df_guess = df_guess[df_guess[col_b2_mp].astype(str).str.strip() == str(b2).strip()]
                                    if lojas_audit and ok_loja:
                                        df_guess = df_guess[df_guess[col_loja_mp].apply(normalize_code).isin(lojas_audit)]
                                    if not df_guess.empty:
                                        v_mp = float(df_guess[col_val_guess].sum())
                                else:
                                    v_mp = 0.0
                        else:
                            v_mp = 0.0
                    else:
                        v_mp = 0.0
                except Exception:
                    v_mp = 0.0

                status_text = "✅ OK" if abs(v_o - v_d) < 0.1 else "❌ Erro"
                results_excel.append({"Planilha": row["Planilha"], "Origem": v_o, "DRE": v_d, "MP DRE": v_mp, "Dif": v_o-v_d, "Dif MP": v_d-v_mp, "Status": status_text})

                mask = st.session_state.au_planilhas_df["Planilha_id"] == sid
                st.session_state.au_planilhas_df.loc[mask, ["Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"]] = [
                    format_brl(v_o), format_brl(v_d), format_brl(v_mp), format_brl(v_o-v_d), format_brl(v_d-v_mp), status_text
                ]
                # desmarca a planilha para indicar concluído
                st.session_state.au_planilhas_df.loc[mask, "Flag"] = False
                prog.progress((idx+1)/n)

            # fim do loop de processamento
            st.success("Auditoria finalizada.")
            # tentar forçar a atualização para refletir flags limpas
            try:
                st.experimental_rerun()
            except Exception:
                st.info("As flags foram limpas. Atualize a página se necessário para ver a alteração.")
