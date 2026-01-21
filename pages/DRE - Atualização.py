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
    s = re.sub(r"[^0-9,.-]", "", s)
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
tab_audit, tab_atual = st.tabs(["Auditoria", "Atualização"])

with tab_atual:
    st.info("A aba Atualização está aqui, mas sem alterações para este ajuste.")

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

    col_ano, col_mes = st.columns(2)
    anos_disponiveis = list(range(2018, datetime.now().year + 1))
    with col_ano:
        ano_sel = st.selectbox("Ano:", anos_disponiveis, index=len(anos_disponiveis) - 1, key="audit_ano")
    with col_mes:
        meses_disponiveis = [""] + list(range(1, 13))  # mês vazio para "não selecionar"
        mes_sel = st.selectbox("Mês (opcional):", meses_disponiveis, index=0, key="audit_mes")

    if not s_ids_audit:
        st.info("Selecione ao menos uma subpasta para listar planilhas e executar auditoria.")
    else:
        if st.button("📊 Executar Auditoria para o período selecionado", key="audit_executar_auditoria"):
            try:
                data_inicio = date(ano_sel, 1, 1)
                data_fim = date(ano_sel, 12, 31)
                if mes_sel != "" and mes_sel is not None:
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
