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

# --- CSS ---
st.markdown(
    """
    <style>
    .block-container { padding-top: 2rem; padding-bottom: 2rem; }
    h1 { margin-bottom: 1.5rem; font-size: 2.2rem; line-height: 1.2; }
    .stSelectbox, .stMultiSelect, .stDateInput { margin-bottom: 1rem; }
    [data-testid="stTable"] td, [data-testid="stTable"] th { padding: 8px 12px !important; }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("Atualizador DRE")

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

@st.cache_data(ttl=300)
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
    except: return None

def tratar_numericos(df, headers):
    indices_valor = [6, 7, 8, 9]
    for idx in indices_valor:
        if idx < len(headers):
            col_name = headers[idx]
            df[col_name] = df[col_name].apply(_parse_currency_like).fillna(0.0)
    return df

def format_brl(val):
    return f"R$ {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")

tab_atual, tab_verif, tab_audit = st.tabs(["Atualização", "Verificar Configurações", "Auditoria"])

with tab_atual:
    # (Aqui seu código atual da aba Atualização, com filtros e atualizações)

    st.write("Aqui vai o conteúdo da aba Atualização...")

with tab_verif:
    # (Aqui seu código atual da aba Verificar Configurações)

    st.write("Aqui vai o conteúdo da aba Verificar Configurações...")

with tab_audit:
    st.markdown("### Auditoria de Faturamento - Independente")

    # Seleção pasta principal e subpastas
    try:
        pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
        map_p = {p["name"]: p["id"] for p in pastas_fech}
        p_sel = st.selectbox("Pasta principal:", options=list(map_p.keys()))
        subpastas = list_child_folders(drive_service, map_p[p_sel])
        map_s = {s["name"]: s["id"] for s in subpastas}
        s_sel = st.multiselect("Subpastas:", options=list(map_s.keys()), default=[])
        s_ids_audit = [map_s[n] for n in s_sel]
    except Exception:
        st.error("Erro ao listar pastas.")
        st.stop()

    # Seleção ano e mês (opcional)
    anos_disponiveis = list(range(2020, datetime.now().year + 1))
    ano_sel = st.selectbox("Ano:", anos_disponiveis, index=anos_disponiveis.index(datetime.now().year))
    meses_disponiveis = list(range(1, 13))
    mes_sel = st.selectbox("Mês:", meses_disponiveis, index=datetime.now().month - 1)

    if not s_ids_audit:
        st.info("Selecione as subpastas para listar as planilhas.")
    else:
        if st.button("📊 Executar Auditoria"):
            try:
                data_inicio = date(ano_sel, mes_sel, 1)
                if mes_sel == 12:
                    data_fim = date(ano_sel + 1, 1, 1) - timedelta(days=1)
                else:
                    data_fim = date(ano_sel, mes_sel + 1, 1) - timedelta(days=1)

                # Ler Origem
                sh_origem = gc.open_by_key(ID_PLANILHA_ORIGEM)
                ws_origem = sh_origem.worksheet(ABA_ORIGEM)
                headers_orig, df_orig = get_headers_and_df_raw(ws_origem)
                df_orig = tratar_numericos(df_orig, headers_orig)
                col_data_orig = detect_date_col(headers_orig)
                col_fat_orig = headers_orig[6]
                col_grupo_orig = headers_orig[5]
                col_loja_orig = headers_orig[3]

                df_orig['_dt'] = pd.to_datetime(df_orig[col_data_orig], dayfirst=True, errors='coerce').dt.date
                df_orig_periodo = df_orig[(df_orig['_dt'] >= data_inicio) & (df_orig['_dt'] <= data_fim)].copy()

                # Listar planilhas destino
                planilhas = list_spreadsheets_in_folders(drive_service, s_ids_audit)
                audit_results = []
                prog = st.progress(0)

                for idx, p in enumerate(planilhas):
                    sid = p["id"]
                    p_name = p["name"]

                    cached = st.session_state.get("sheet_codes", {}).get(sid)
                    if not cached:
                        try:
                            sh_dest = gc.open_by_key(sid)
                            b2, b3 = read_codes_from_config_sheet(sh_dest)
                            st.session_state.setdefault("sheet_codes", {})[sid] = (b2, b3)
                        except:
                            b2, b3 = None, None
                    else:
                        b2, b3 = cached

                    if not b2:
                        audit_results.append({"Planilha": p_name, "Faturamento Origem": 0, "Faturamento DRE": 0, "Diferença": 0, "Status": "Sem Config"})
                        prog.progress((idx + 1) / len(planilhas))
                        continue

                    df_o_f = df_orig_periodo[df_orig_periodo[col_grupo_orig].astype(str).str.strip() == b2]
                    if b3:
                        df_o_f = df_o_f[df_o_f[col_loja_orig].astype(str).str.strip() == b3]
                    total_orig = df_o_f[col_fat_orig].sum()

                    total_dest = 0
                    try:
                        sh_dest = gc.open_by_key(sid)
                        ws_dest = sh_dest.worksheet("Importado_Fat")
                        h_dest, df_d = get_headers_and_df_raw(ws_dest)
                        df_d = tratar_numericos(df_d, h_dest)
                        c_dt_d = detect_date_col(h_dest)
                        c_ft_d = h_dest[6]
                        df_d['_dt'] = pd.to_datetime(df_d[c_dt_d], dayfirst=True, errors='coerce').dt.date
                        df_d_periodo = df_d[(df_d['_dt'] >= data_inicio) & (df_d['_dt'] <= data_fim)].copy()
                        total_dest = df_d_periodo[c_ft_d].sum()
                    except:
                        df_d_periodo = pd.DataFrame()

                    diff = total_orig - total_dest
                    status = "✅ OK" if abs(diff) < 0.01 else "❌ Divergente"

                    audit_results.append({
                        "Planilha": p_name,
                        "Faturamento Origem": total_orig,
                        "Faturamento DRE": total_dest,
                        "Diferença": diff,
                        "Status": status,
                        "df_o_raw": df_o_f,
                        "df_d_raw": df_d_periodo
                    })
                    prog.progress((idx + 1) / len(planilhas))

                df_main = pd.DataFrame(audit_results).drop(columns=["df_o_raw", "df_d_raw"])

                for col in ["Faturamento Origem", "Faturamento DRE", "Diferença"]:
                    df_main[col] = df_main[col].apply(format_brl)

                st.dataframe(df_main, use_container_width=True, hide_index=True)

                st.markdown("---")
                st.subheader("Detalhamento de Divergências")
                for res in audit_results:
                    if res["Status"] == "❌ Divergente":
                        with st.expander(f"🔍 Ver detalhes: {res['Planilha']}"):
                            df_o = res["df_o_raw"]
                            df_d = res["df_d_raw"]
                            if df_o.empty or df_d.empty:
                                st.write("Dados insuficientes para detalhamento.")
                                continue

                            df_o['Mes_Ano'] = pd.to_datetime(df_o['_dt']).dt.strftime('%Y-%m')
                            df_d['Mes_Ano'] = pd.to_datetime(df_d['_dt']).dt.strftime('%Y-%m')

                            fat_orig_mes = df_o.groupby('Mes_Ano')[col_fat_orig].sum()
                            fat_dest_mes = df_d.groupby('Mes_Ano')[h_dest[6]].sum()
                            meses = sorted(set(fat_orig_mes.index) | set(fat_dest_mes.index))
                            detalhes_mes = []
                            for m in meses:
                                vo = fat_orig_mes.get(m, 0)
                                vd = fat_dest_mes.get(m, 0)
                                diff_m = vo - vd
                                status_m = "✅ OK" if abs(diff_m) < 0.01 else "❌ Divergente"
                                detalhes_mes.append({
                                    "Mês": m,
                                    "Faturamento Origem": format_brl(vo),
                                    "Faturamento DRE": format_brl(vd),
                                    "Diferença": format_brl(diff_m),
                                    "Status": status_m
                                })
                            st.write("**Resumo Mensal:**")
                            st.table(pd.DataFrame(detalhes_mes))

                            mes_sel = st.selectbox(f"Selecionar mês para detalhar por dia - {res['Planilha']}", options=[d["Mês"] for d in detalhes_mes], key=f"mes_dia_{res['Planilha']}")

                            fat_orig_dia = df_o[df_o['Mes_Ano'] == mes_sel].groupby('_dt')[col_fat_orig].sum()
                            fat_dest_dia = df_d[df_d['Mes_Ano'] == mes_sel].groupby('_dt')[h_dest[6]].sum()
                            dias = sorted(set(fat_orig_dia.index) | set(fat_dest_dia.index))
                            detalhes_dia = []
                            for d in dias:
                                vo = fat_orig_dia.get(d, 0)
                                vd = fat_dest_dia.get(d, 0)
                                diff_d = vo - vd
                                status_d = "✅ OK" if abs(diff_d) < 0.01 else "❌ Divergente"
                                detalhes_dia.append({
                                    "Dia": d.strftime('%d/%m/%Y'),
                                    "Faturamento Origem": format_brl(vo),
                                    "Faturamento DRE": format_brl(vd),
                                    "Diferença": format_brl(diff_d),
                                    "Status": status_d
                                })
                            st.write(f"**Detalhamento Diário ({mes_sel}):**")
                            st.table(pd.DataFrame(detalhes_dia))

            except Exception as e:
                st.error(f"Erro na auditoria: {e}")
