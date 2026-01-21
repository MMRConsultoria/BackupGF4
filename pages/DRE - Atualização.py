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
    .block-container { padding-top: 1rem; padding-bottom: 5rem; }
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

@st.cache_data(ttl=300)
def get_conf_map(sheet_ids, target_name):
    # Função auxiliar (não usada automaticamente no fluxo de atualização)
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
    if s is None:
        return None
    s = str(s).strip()
    if s == "" or s in ["-", "–"]:
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1].strip()
    s = s.replace("R$", "").replace("r$", "").replace(" ", "")
    s = re.sub(r"[^0-9,.\-]", "", s)
    if s == "" or s == "-" or s == ".": 
        return None
    if s.count(".") > 0 and s.count(",") > 0:
        s = s.replace(".", "").replace(",", ".")
    else:
        if s.count(",") > 0 and s.count(".") == 0:
            s = s.replace(",", ".")
        if s.count(".") > 1 and s.count(",") == 0:
            s = s.replace(".", "")
    try:
        val = float(s)
        if neg:
            val = -val
        return val
    except:
        return None

def tratar_numericos(df, headers):
    indices_valor = [6, 7, 8, 9]  # G, H, I, J
    for idx in indices_valor:
        if idx < len(headers):
            col_name = headers[idx]
            orig_series = df[col_name].astype(object).copy()
            parsed = orig_series.apply(_parse_currency_like)
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

# ---------------- UI GLOBAL ----------------
col_d1, col_d2 = st.columns(2)
with col_d1:
    data_de = st.date_input("De", value=date.today() - timedelta(days=30))
with col_d2:
    data_ate = st.date_input("Até", value=date.today())

# Se não houver session_state para sheet_codes, inicializa
if "sheet_codes" not in st.session_state:
    st.session_state["sheet_codes"] = {}  # map sheet_id -> (b2,b3) or (None,None) se não existir

# Seleção de pastas (com cache)
try:
    pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
    map_p = {p["name"]: p["id"] for p in pastas_fech}
    p_sel = st.selectbox("Pasta principal:", options=list(map_p.keys()))
    subpastas = list_child_folders(drive_service, map_p[p_sel])
    map_s = {s["name"]: s["id"] for s in subpastas}
    s_sel = st.multiselect("Subpastas:", options=list(map_s.keys()), default=[])
    s_ids = [map_s[n] for n in s_sel]
except Exception:
    st.error("Erro ao listar pastas. Verifique autenticação/permissões.")
    st.stop()

if not s_ids:
    st.info("Selecione as subpastas para listar as planilhas (inicie em branco, escolha quando quiser).")
    # Não st.stop() aqui para permitir ir na aba Verificar e fazer a verificação manualmente
# ---------------- TABS ----------------
tab_atual, tab_verif = st.tabs(["Atualização", "Verificar Configurações"])

with tab_atual:
    if not s_ids:
        st.info("Escolha ao menos uma subpasta para listar planilhas e habilitar a atualização.")
    else:
        with st.spinner("Buscando planilhas nas subpastas selecionadas..."):
            planilhas = list_spreadsheets_in_folders(drive_service, s_ids)
        if not planilhas:
            st.warning("Nenhuma planilha encontrada nas subpastas selecionadas.")
        else:
            df_list = pd.DataFrame(planilhas).sort_values("name").reset_index(drop=True)
            df_list = df_list.rename(columns={"name": "Planilha", "id": "ID_Planilha"})
            # Não verificar automaticamente existência da aba de configuração aqui (economia de API)
            st.markdown("Selecione as planilhas que deseja atualizar e marque as opções desejadas:")
            st.markdown('<div class="global-selection-container">', unsafe_allow_html=True)
            c1, c2, c3, _ = st.columns([1.2, 1.2, 1.2, 5])
            with c1:
                s_desc = st.checkbox("Desconto", value=True, key="chk_desc")
            with c2:
                s_mp = st.checkbox("Meio Pagto", value=True, key="chk_mp")
            with c3:
                s_fat = st.checkbox("Faturamento", value=True, key="chk_fat")
            st.markdown('</div>', unsafe_allow_html=True)

            df_list["Desconto"], df_list["Meio Pagamento"], df_list["Faturamento"] = s_desc, s_mp, s_fat
            config = {
                "Planilha": st.column_config.TextColumn("Planilha", disabled=True),
                "ID_Planilha": None,
                "parent_folder_id": None,
                "Desconto": st.column_config.CheckboxColumn("Desc."),
                "Meio Pagamento": st.column_config.CheckboxColumn("M.Pag"),
                "Faturamento": st.column_config.CheckboxColumn("Fat."),
            }
            meio = len(df_list) // 2 + (len(df_list) % 2)
            col_t1, col_t2 = st.columns(2)
            with col_t1:
                edit_esq = st.data_editor(df_list.iloc[:meio], key="t1", use_container_width=True, column_config=config, hide_index=True)
            with col_t2:
                edit_dir = st.data_editor(df_list.iloc[meio:], key="t2", use_container_width=True, column_config=config, hide_index=True)

            if st.button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True):
                df_final_edit = pd.concat([edit_esq, edit_dir], ignore_index=True)
                df_marcadas = df_final_edit[(df_final_edit["Desconto"]) | (df_final_edit["Meio Pagamento"]) | (df_final_edit["Faturamento"])].copy()

                if df_marcadas.empty:
                    st.warning("Nenhuma planilha marcada.")
                    st.stop()

                # Ler origem apenas uma vez
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
                    st.error(f"Erro ao ler planilha origem: {e}")
                    st.stop()

                progresso = st.progress(0)
                logs = []
                total = len(df_marcadas)

                for i, (_, row) in enumerate(df_marcadas.iterrows()):
                    try:
                        sheet_id = row["ID_Planilha"]
                        # Reusar b2/b3 se já foram obtidos pela aba de verificação
                        cached = st.session_state.get("sheet_codes", {}).get(sheet_id)
                        if cached is not None:
                            b2, b3 = cached
                        else:
                            # Buscar na hora (único acesso necessário se não houver cache)
                            sh_dest = gc.open_by_key(sheet_id)
                            b2, b3 = read_codes_from_config_sheet(sh_dest)
                            # guardamos o resultado para evitar nova leitura se o usuário repetir
                            st.session_state["sheet_codes"][sheet_id] = (b2, b3)

                        if not b2:
                            logs.append(f"{row['Planilha']}: B2 não encontrado (pulei).")
                            progresso.progress(min((i + 1) / total, 1.0))
                            continue

                        # Carregar origem filtrada por códigos
                        col_f_name = headers_orig[5]  # Grupo
                        col_d_name = headers_orig[3]  # Loja

                        df_para_inserir = df_orig_filtrado[df_orig_filtrado[col_f_name].astype(str).str.strip() == b2].copy()
                        if b3:
                            df_para_inserir = df_para_inserir[df_para_inserir[col_d_name].astype(str).str.strip() == b3]

                        if df_para_inserir.empty:
                            logs.append(f"{row['Planilha']}: Sem dados para o período/filtro.")
                            progresso.progress(min((i + 1) / total, 1.0))
                            continue

                        # Abrir/pegar worksheet destino
                        sh_dest = gc.open_by_key(sheet_id)
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

                        send_df = df_final_ws[h_final].copy()
                        send_df = send_df.where(pd.notna(send_df), "")
                        final_vals = [h_final] + send_df.values.tolist()

                        ws_dest.clear()
                        ws_dest.update("A1", final_vals, value_input_option='USER_ENTERED')
                        logs.append(f"{row['Planilha']}: Sucesso ({len(df_para_inserir)} linhas).")
                    except Exception as e:
                        logs.append(f"{row['Planilha']}: Erro: {e}")
                    progresso.progress(min((i + 1) / total, 1.0))

                st.success("Processo concluído.")
                st.write("\n".join(logs))

with tab_verif:
    st.markdown("Aqui você pode verificar isoladamente quais planilhas possuem a aba de configuração e **buscar os códigos B2/B3** (essa ação fará leituras nas planilhas selecionadas).")
    if not s_ids:
        st.info("Escolha ao menos uma subpasta para listar planilhas antes de verificar.")
    else:
        with st.spinner("Listando planilhas..."):
            planilhas = list_spreadsheets_in_folders(drive_service, s_ids)
        if not planilhas:
            st.warning("Nenhuma planilha encontrada nas subpastas selecionadas.")
        else:
            df_list_ver = pd.DataFrame(planilhas).sort_values("name").reset_index(drop=True)
            df_list_ver = df_list_ver.rename(columns={"name": "Planilha", "id": "ID_Planilha"})

            # Prepara dados para exibir tabela com cache atual
            data_display = []
            for _, row in df_list_ver.iterrows():
                sid = row["ID_Planilha"]
                nm = row["Planilha"]
                b2b3 = st.session_state.get("sheet_codes", {}).get(sid, (None, None))
                has_conf = b2b3[0] is not None  # se B2 existe, considera que tem config
                data_display.append({
                    "Planilha": nm,
                    "Config": "Sim" if has_conf else "Não",
                    "B2": b2b3[0] or "",
                    "B3": b2b3[1] or "",
                })
            df_display = pd.DataFrame(data_display)

            st.dataframe(df_display, use_container_width=True)

            if st.button("🔎 Verificar presença da aba de configuração e buscar B2/B3"):
                progresso = st.progress(0)
                results = []
                total = len(df_list_ver)
                for i, r in df_list_ver.iterrows():
                    sid = r["ID_Planilha"]
                    nm = r["Planilha"]
                    try:
                        sh = gc.open_by_key(sid)
                        b2, b3 = read_codes_from_config_sheet(sh)
                        has_conf = b2 is not None
                        results.append({"Planilha": nm, "ID_Planilha": sid, "has_config": has_conf, "B2": b2, "B3": b3})
                        st.session_state["sheet_codes"][sid] = (b2, b3)
                    except Exception as e:
                        results.append({"Planilha": nm, "ID_Planilha": sid, "has_config": False, "B2": None, "B3": None, "error": str(e)})
                    progresso.progress(min((i + 1) / total, 1.0))
                df_res = pd.DataFrame(results)
                st.success("Verificação concluída. Resultados atualizados:")
                # Atualiza a tabela exibida com os novos dados
                df_display = df_res[["Planilha", "has_config", "B2", "B3"]].copy()
                df_display.rename(columns={"has_config": "Config"}, inplace=True)
                df_display["Config"] = df_display["Config"].apply(lambda x: "Sim" if x else "Não")
                st.dataframe(df_display.fillna(""), use_container_width=True)
