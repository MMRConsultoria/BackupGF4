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
ID_PLANILHA_ORIGEM = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
ABA_ORIGEM = "Fat Sistema Externo"

st.set_page_config(page_title="Atualizador DRE", layout="wide")

# --- CSS pequeno para espaçamento ---
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
    if _drive is None:
        raise RuntimeError("drive_service não está disponível (googleapiclient não carregado).")
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
    if _drive is None:
        raise RuntimeError("drive_service não está disponível (googleapiclient não carregado).")
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
        if ws is None:
            return None, None
        b2 = ws.acell("B2").value
        b3 = ws.acell("B3").value
        return (str(b2).strip() if b2 else None, str(b3).strip() if b3 else None)
    except Exception:
        return None, None

def get_headers_and_df_raw(ws):
    vals = ws.get_all_values()
    if not vals:
        return [], pd.DataFrame()
    headers = [str(h).strip() for h in vals[0]]
    df = pd.DataFrame(vals[1:], columns=headers)
    return headers, df

def detect_date_col(headers):
    for h in headers:
        if "data" in h.lower():
            return h
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
        if s.count(",") > 0 and s.count(".") == 0:
            s = s.replace(",", ".")
        if s.count(".") > 1 and s.count(",") == 0:
            s = s.replace(".", "")
    try:
        val = float(s)
        if neg: val = -val
        return val
    except:
        return None

def tratar_numericos(df, headers):
    indices_valor = [6, 7, 8, 9]
    for idx in indices_valor:
        if idx < len(headers):
            col_name = headers[idx]
            df[col_name] = df[col_name].apply(_parse_currency_like).fillna(0.0)
    return df

def format_brl(val):
    try:
        return f"R$ {float(val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return val

# ---------------- UI GLOBAL ----------------
if "sheet_codes" not in st.session_state:
    st.session_state["sheet_codes"] = {}

# ---------------- TABS ----------------
tab_atual, tab_verif, tab_audit = st.tabs(["Atualização", "Verificar Configurações", "Auditoria"])

with tab_atual:
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        data_de = st.date_input("De", value=date.today() - timedelta(days=30), key="atual_data_de")
    with col_d2:
        data_ate = st.date_input("Até", value=date.today(), key="atual_data_ate")

    try:
        pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
        map_p = {p["name"]: p["id"] for p in pastas_fech}
        p_sel = st.selectbox("Pasta principal:", options=list(map_p.keys()), key="atual_pasta_principal")
        subpastas = list_child_folders(drive_service, map_p[p_sel])
        map_s = {s["name"]: s["id"] for s in subpastas}
        s_sel = st.multiselect("Subpastas:", options=list(map_s.keys()), default=[], key="atual_subpastas")
        s_ids = [map_s[n] for n in s_sel]
    except Exception:
        st.error("Erro ao listar pastas.")
        st.stop()

    if not s_ids:
        st.info("Selecione as subpastas para listar as planilhas.")
    else:
        with st.spinner("Buscando planilhas..."):
            planilhas = list_spreadsheets_in_folders(drive_service, s_ids)
        if not planilhas:
            st.warning("Nenhuma planilha encontrada.")
        else:
            df_list = pd.DataFrame(planilhas).sort_values("name").reset_index(drop=True)
            df_list = df_list.rename(columns={"name": "Planilha", "id": "ID_Planilha"})
            
            c1, c2, c3, _ = st.columns([1.2, 1.2, 1.2, 5])
            with c1: s_desc = st.checkbox("Desconto", value=True, key="atual_chk_desc")
            with c2: s_mp = st.checkbox("Meio Pagto", value=True, key="atual_chk_mp")
            with c3: s_fat = st.checkbox("Faturamento", value=True, key="atual_chk_fat")

            df_list["Desconto"], df_list["Meio Pagamento"], df_list["Faturamento"] = s_desc, s_mp, s_fat
            config = {
                "Planilha": st.column_config.TextColumn("Planilha", disabled=True),
                "ID_Planilha": None, "parent_folder_id": None,
                "Desconto": st.column_config.CheckboxColumn("Desc."),
                "Meio Pagamento": st.column_config.CheckboxColumn("M.Pag"),
                "Faturamento": st.column_config.CheckboxColumn("Fat."),
            }
            meio = len(df_list) // 2 + (len(df_list) % 2)
            col_t1, col_t2 = st.columns(2)
            with col_t1: edit_esq = st.data_editor(df_list.iloc[:meio], key="atual_t1", use_container_width=True, column_config=config, hide_index=True)
            with col_t2: edit_dir = st.data_editor(df_list.iloc[meio:], key="atual_t2", use_container_width=True, column_config=config, hide_index=True)

            if st.button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True, key="btn_iniciar_atualizacao"):
                df_final_edit = pd.concat([edit_esq, edit_dir], ignore_index=True)
                df_marcadas = df_final_edit[(df_final_edit["Desconto"]) | (df_final_edit["Meio Pagamento"]) | (df_final_edit["Faturamento"])].copy()
                if df_marcadas.empty:
                    st.warning("Nenhuma planilha marcada.")
                    st.stop()

                try:
                    sh_origem = gc.open_by_key(ID_PLANILHA_ORIGEM)
                    ws_origem = sh_origem.worksheet(ABA_ORIGEM)
                    headers_orig, df_orig = get_headers_and_df_raw(ws_origem)
                    df_orig = tratar_numericos(df_orig, headers_orig)
                    col_data_orig = detect_date_col(headers_orig)
                    df_orig_temp = df_orig.copy()
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
                        sid = row["ID_Planilha"]
                        cached = st.session_state["sheet_codes"].get(sid)
                        if cached: b2, b3 = cached
                        else:
                            sh_dest = gc.open_by_key(sid)
                            b2, b3 = read_codes_from_config_sheet(sh_dest)
                            st.session_state["sheet_codes"][sid] = (b2, b3)

                        if not b2:
                            logs.append(f"{row['Planilha']}: Sem config."); continue

                        col_f_name, col_d_name = headers_orig[5], headers_orig[3]
                        df_para_inserir = df_orig_filtrado[df_orig_filtrado[col_f_name].astype(str).str.strip() == b2].copy()
                        if b3: df_para_inserir = df_para_inserir[df_para_inserir[col_d_name].astype(str).str.strip() == b3]

                        if df_para_inserir.empty:
                            logs.append(f"{row['Planilha']}: Sem dados."); continue

                        sh_dest = gc.open_by_key(sid)
                        try: ws_dest = sh_dest.worksheet("Importado_Fat")
                        except: ws_dest = sh_dest.add_worksheet("Importado_Fat", 1000, 30)

                        headers_dest, df_dest = get_headers_and_df_raw(ws_dest)
                        df_dest = tratar_numericos(df_dest, headers_dest)

                        if df_dest.empty:
                            df_final_ws, h_final = df_para_inserir, headers_orig
                        else:
                            col_dt_dest = detect_date_col(headers_dest) or col_data_orig
                            df_dest_temp = df_dest.copy()
                            df_dest_temp['_dt'] = pd.to_datetime(df_dest_temp[col_dt_dest], dayfirst=True, errors='coerce').dt.date
                            to_remove = (df_dest_temp['_dt'] >= data_de) & (df_dest_temp['_dt'] <= data_ate)
                            if col_f_name in df_dest.columns: to_remove &= (df_dest[col_f_name].astype(str).str.strip() == b2)
                            if b3 and col_d_name in df_dest.columns: to_remove &= (df_dest[col_d_name].astype(str).str.strip() == b3)
                            df_final_ws = pd.concat([df_dest.loc[~to_remove], df_para_inserir], ignore_index=True)
                            h_final = headers_dest if headers_dest else headers_orig

                        send_df = df_final_ws[h_final].copy().where(pd.notna(df_final_ws[h_final]), "")
                        ws_dest.clear()
                        ws_dest.update("A1", [h_final] + send_df.values.tolist(), value_input_option='USER_ENTERED')
                        logs.append(f"{row['Planilha']}: Sucesso.")
                    except Exception as e: logs.append(f"{row['Planilha']}: Erro: {e}")
                    progresso.progress(min((i + 1) / total, 1.0))
                st.success("Concluído!"); st.write("\n".join(logs))

with tab_verif:
    st.markdown("Verifique a presença da aba de configuração e os códigos B2/B3.")
    try:
        pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
        map_p = {p["name"]: p["id"] for p in pastas_fech}
        p_sel = st.selectbox("Pasta principal:", options=list(map_p.keys()), key="verif_pasta_principal")
        subpastas = list_child_folders(drive_service, map_p[p_sel])
        map_s = {s["name"]: s["id"] for s in subpastas}
        s_sel = st.multiselect("Subpastas:", options=list(map_s.keys()), default=[], key="verif_subpastas")
        s_ids = [map_s[n] for n in s_sel]
    except Exception:
        st.error("Erro ao listar pastas.")
        st.stop()

    if not s_ids:
        st.info("Selecione as subpastas primeiro.")
    else:
        with st.spinner("Listando..."):
            planilhas = list_spreadsheets_in_folders(drive_service, s_ids)
        if planilhas:
            df_list_ver = pd.DataFrame(planilhas).sort_values("name").reset_index(drop=True)
            df_list_ver = df_list_ver.rename(columns={"name": "Planilha", "id": "ID_Planilha"})
            
            data_display = []
            for _, row in df_list_ver.iterrows():
                sid = row["ID_Planilha"]
                b2b3 = st.session_state["sheet_codes"].get(sid, (None, None))
                data_display.append({
                    "Planilha": row["Planilha"],
                    "Config": "Sim" if b2b3[0] else "Não",
                    "B2": b2b3[0] or "", "B3": b2b3[1] or ""
                })
            st.dataframe(pd.DataFrame(data_display), use_container_width=True)

            if st.button("🔎 Verificar configurações", key="btn_verificar_config"):
                prog = st.progress(0)
                total = len(df_list_ver)
                for i, r in df_list_ver.iterrows():
                    try:
                        sh = gc.open_by_key(r["ID_Planilha"])
                        b2, b3 = read_codes_from_config_sheet(sh)
                        st.session_state["sheet_codes"][r["ID_Planilha"]] = (b2, b3)
                    except: pass
                    prog.progress(min((i + 1) / total, 1.0))
                st.experimental_rerun()

with tab_audit:
    st.header("Auditoria (independente)")
    st.markdown("Escolha pasta principal, subpastas, ano e mês — a auditoria será executada só no período selecionado.")

    try:
        pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
    except Exception as e:
        st.error(f"Erro ao listar pastas principais no Drive: {e}")
        st.stop()

    if not pastas_fech:
        st.warning("Nenhuma pasta encontrada com o filtro 'fechamento'. Verifique o PASTA_PRINCIPAL_ID e permissões.")
        st.stop()

    map_p = {p["name"]: p["id"] for p in pastas_fech}
    p_sel = st.selectbox("Pasta principal:", options=list(map_p.keys()), key="audit_pasta_principal")

    try:
        subpastas = list_child_folders(drive_service, map_p[p_sel])
    except Exception as e:
        st.error(f"Erro ao listar subpastas: {e}")
        st.stop()

    if not subpastas:
        st.info("Nenhuma subpasta encontrada nesta pasta principal.")
        s_ids_audit = []
    else:
        map_s = {s["name"]: s["id"] for s in subpastas}
        s_sel = st.multiselect("Subpastas:", options=list(map_s.keys()), default=[], key="audit_subpastas")
        s_ids_audit = [map_s[n] for n in s_sel]

    anos_disponiveis = list(range(2018, datetime.now().year + 1))
    ano_sel = st.selectbox("Ano:", anos_disponiveis, index=len(anos_disponiveis) - 1, key="audit_ano")
    meses_disponiveis = list(range(1, 13))
    mes_sel = st.selectbox("Mês:", meses_disponiveis, index=datetime.now().month - 1, key="audit_mes")

    if not s_ids_audit:
        st.info("Selecione ao menos uma subpasta para listar planilhas e executar auditoria.")
    else:
        if st.button("🔎 Listar planilhas nas subpastas selecionadas", key="audit_listar_planilhas"):
            try:
                planilhas = list_spreadsheets_in_folders(drive_service, s_ids_audit)
                if not planilhas:
                    st.warning("Nenhuma planilha encontrada nas subpastas selecionadas.")
                else:
                    df_list = pd.DataFrame(planilhas).sort_values("name").reset_index(drop=True)
                    df_list = df_list.rename(columns={"name": "Planilha", "id": "ID_Planilha"})
                    st.dataframe(df_list[["Planilha", "ID_Planilha"]], use_container_width=True)
            except Exception as e:
                st.error(f"Erro ao listar planilhas: {e}")

        if st.button("📊 Executar Auditoria para o período selecionado", key="audit_executar_auditoria"):
            try:
                data_inicio = date(ano_sel, mes_sel, 1)
                if mes_sel == 12:
                    data_fim = date(ano_sel + 1, 1, 1) - timedelta(days=1)
                else:
                    data_fim = date(ano_sel, mes_sel + 1, 1) - timedelta(days=1)

                st.info(f"Executando auditoria de {data_inicio} até {data_fim} ...")

                sh_origem = gc.open_by_key(ID_PLANILHA_ORIGEM)
                ws_origem = sh_origem.worksheet(ABA_ORIGEM)
                headers_orig, df_orig = get_headers_and_df_raw(ws_origem)
                if df_orig.empty:
                    st.warning("Planilha de origem vazia.")
                df_orig = tratar_numericos(df_orig, headers_orig)
                col_data_orig = detect_date_col(headers_orig)
                if col_data_orig is None:
                    st.error("Não foi possível detectar coluna de data na planilha origem.")
                    st.stop()
                col_fat_orig = headers_orig[6] if len(headers_orig) > 6 else None
                col_grupo_orig = headers_orig[5] if len(headers_orig) > 5 else None
                col_loja_orig = headers_orig[3] if len(headers_orig) > 3 else None
                if not col_fat_orig or not col_grupo_orig or not col_loja_orig:
                    st.error("Estrutura da planilha origem inesperada (colunas F/G/D). Verifique o layout.")
                    st.stop()

                df_orig['_dt'] = pd.to_datetime(df_orig[col_data_orig], dayfirst=True, errors='coerce').dt.date
                df_orig_periodo = df_orig[(df_orig['_dt'] >= data_inicio) & (df_orig['_dt'] <= data_fim)].copy()

                planilhas = list_spreadsheets_in_folders(drive_service, s_ids_audit)
                if not planilhas:
                    st.warning("Nenhuma planilha encontrada para auditoria.")
                    st.stop()

                audit_results = []
                prog = st.progress(0)
                total_planilhas = len(planilhas)

                for idx, p in enumerate(planilhas):
                    sid = p["id"]
                    p_name = p["name"]

                    cached = st.session_state["sheet_codes"].get(sid)
                    if not cached:
                        try:
                            sh_dest = gc.open_by_key(sid)
                            b2, b3 = read_codes_from_config_sheet(sh_dest)
                            st.session_state["sheet_codes"][sid] = (b2, b3)
                        except Exception:
                            b2, b3 = None, None
                    else:
                        b2, b3 = cached

                    if not b2:
                        audit_results.append({"Planilha": p_name, "Faturamento Origem": 0.0, "Faturamento DRE": 0.0, "Diferença": 0.0, "Status": "Sem Config", "df_o_raw": pd.DataFrame(), "df_d_raw": pd.DataFrame()})
                        prog.progress((idx + 1) / total_planilhas)
                        continue

                    df_o_f = df_orig_periodo[df_orig_periodo[col_grupo_orig].astype(str).str.strip() == b2]
                    if b3:
                        df_o_f = df_o_f[df_o_f[col_loja_orig].astype(str).str.strip() == b3]
                    total_orig = df_o_f[col_fat_orig].sum()

                    total_dest = 0.0
                    df_d_periodo = pd.DataFrame()
                    try:
                        sh_dest = gc.open_by_key(sid)
                        ws_dest = sh_dest.worksheet("Importado_Fat")
                        h_dest, df_d = get_headers_and_df_raw(ws_dest)
                        df_d = tratar_numericos(df_d, h_dest)
                        c_dt_d = detect_date_col(h_dest)
                        if c_dt_d is None or len(h_dest) <= 6:
                            df_d_periodo = pd.DataFrame()
                        else:
                            c_ft_d = h_dest[6]
                            df_d['_dt'] = pd.to_datetime(df_d[c_dt_d], dayfirst=True, errors='coerce').dt.date
                            df_d_periodo = df_d[(df_d['_dt'] >= data_inicio) & (df_d['_dt'] <= data_fim)].copy()
                            total_dest = df_d_periodo[c_ft_d].sum()
                    except Exception:
                        df_d_periodo = pd.DataFrame()

                    diff = total_orig - total_dest
                    status = "✅ OK" if abs(diff) < 0.01 else "❌ Divergente"

                    audit_results.append({
                        "Planilha": p_name,
                        "Faturamento Origem": float(total_orig),
                        "Faturamento DRE": float(total_dest),
                        "Diferença": float(diff),
                        "Status": status,
                        "df_o_raw": df_o_f,
                        "df_d_raw": df_d_periodo
                    })

                    prog.progress((idx + 1) / total_planilhas)

                df_main = pd.DataFrame(audit_results).drop(columns=["df_o_raw", "df_d_raw"])
                for col in ["Faturamento Origem", "Faturamento DRE", "Diferença"]:
                    df_main[col] = df_main[col].apply(format_brl)
                st.subheader("Resumo por planilha")
                st.dataframe(df_main, use_container_width=True, hide_index=True)

                st.markdown("---")
                st.subheader("Detalhamento (apenas planilhas divergentes)")
                for i, res in enumerate(audit_results):
                    if res["Status"] == "❌ Divergente":
                        with st.expander(f"🔍 {res['Planilha']} - Detalhes"):
                            df_o = res["df_o_raw"]
                            df_d = res["df_d_raw"]
                            if df_o.empty and df_d.empty:
                                st.write("Sem dados para detalhamento (origem e/ou destino vazios).")
                                continue

                            if not df_o.empty:
                                df_o['Mes_Ano'] = pd.to_datetime(df_o['_dt']).dt.strftime('%Y-%m')
                                fat_orig_mes = df_o.groupby('Mes_Ano')[col_fat_orig].sum()
                            else:
                                fat_orig_mes = pd.Series(dtype=float)

                            if not df_d.empty:
                                df_d['Mes_Ano'] = pd.to_datetime(df_d['_dt']).dt.strftime('%Y-%m')
                                h_dest_name = h_dest[6] if 'h_dest' in locals() and len(h_dest) > 6 else None
                                if h_dest_name:
                                    fat_dest_mes = df_d.groupby('Mes_Ano')[h_dest_name].sum()
                                else:
                                    fat_dest_mes = pd.Series(dtype=float)
                            else:
                                fat_dest_mes = pd.Series(dtype=float)

                            meses = sorted(set(list(fat_orig_mes.index) + list(fat_dest_mes.index)))
                            detalhes_mes = []
                            for m in meses:
                                vo = float(fat_orig_mes.get(m, 0.0))
                                vd = float(fat_dest_mes.get(m, 0.0))
                                diff_m = vo - vd
                                status_m = "✅ OK" if abs(diff_m) < 0.01 else "❌ Divergente"
                                detalhes_mes.append({
                                    "Mês": m,
                                    "Faturamento Origem": format_brl(vo),
                                    "Faturamento DRE": format_brl(vd),
                                    "Diferença": format_brl(diff_m),
                                    "Status": status_m
                                })
                            if detalhes_mes:
                                st.write("Resumo (por mês):")
                                st.table(pd.DataFrame(detalhes_mes))
                            else:
                                st.write("Sem detalhes mensais disponíveis.")

                            meses_opcoes = [d["Mês"] for d in detalhes_mes] if detalhes_mes else []
                            if meses_opcoes:
                                mes_sel_local = st.selectbox(f"Selecionar mês para detalhar por dia - {res['Planilha']}", options=meses_opcoes, key=f"audit_mes_dia_{i}")
                                d_o = res["df_o_raw"]
                                d_d = res["df_d_raw"]
                                if not d_o.empty:
                                    d_o_sel = d_o[pd.to_datetime(d_o['_dt']).dt.strftime('%Y-%m') == mes_sel_local].groupby('_dt')[col_fat_orig].sum()
                                else:
                                    d_o_sel = pd.Series(dtype=float)
                                if not d_d.empty and ('h_dest' in locals() and len(h_dest) > 6):
                                    d_d_sel = d_d[pd.to_datetime(d_d['_dt']).dt.strftime('%Y-%m') == mes_sel_local].groupby('_dt')[h_dest[6]].sum()
                                else:
                                    d_d_sel = pd.Series(dtype=float)

                                dias = sorted(set(list(d_o_sel.index) + list(d_d_sel.index)))
                                detalhes_dia = []
                                for d in dias:
                                    vo = float(d_o_sel.get(d, 0.0))
                                    vd = float(d_d_sel.get(d, 0.0))
                                    diff_d = vo - vd
                                    status_d = "✅ OK" if abs(diff_d) < 0.01 else "❌ Divergente"
                                    detalhes_dia.append({
                                        "Dia": d.strftime('%d/%m/%Y'),
                                        "Faturamento Origem": format_brl(vo),
                                        "Faturamento DRE": format_brl(vd),
                                        "Diferença": format_brl(diff_d),
                                        "Status": status_d
                                    })
                                if detalhes_dia:
                                    st.write(f"Detalhamento diário ({mes_sel_local}):")
                                    st.table(pd.DataFrame(detalhes_dia))
                                else:
                                    st.write("Sem dados diários para o mês selecionado.")

            except Exception as e:
                st.error(f"Erro ao executar auditoria: {e}")
