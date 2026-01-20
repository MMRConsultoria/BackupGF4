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

# Planilha origem (Fat Sistema Externo)
ID_PLANILHA_ORIGEM = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
ABA_ORIGEM = "Fat Sistema Externo"

st.set_page_config(page_title="Atualizador DRE", layout="wide")

# --- CSS compactação leve ---
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
    """
    Verifica se a planilha contém (em qualquer aba) a substring target_name.
    Retorna dict {sheet_id: bool}
    """
    res = {}
    target_clean = target_name.strip().lower()
    for sid in sheet_ids:
        try:
            sh = gc.open_by_key(sid)
            titles = [ws.title.strip().lower() for ws in sh.worksheets()]
            # marca True se qualquer aba contiver a substring target_clean
            res[sid] = any(target_clean in t for t in titles)
        except Exception:
            res[sid] = False
    return res

def read_codes_from_config_sheet(gsheet):
    """
    Lê B2 (grupo) e B3 (loja) da aba 'Configurações Não Apagar' da planilha destino.
    Retorna tupla (b2, b3) com strings ou (None, None).
    """
    try:
        # busca worksheet com nome que contenha TARGET_SHEET_NAME (mais robusto)
        ws = None
        for w in gsheet.worksheets():
            if TARGET_SHEET_NAME.strip().lower() in w.title.strip().lower():
                ws = w
                break
        if ws is None:
            return None, None
        b2 = ws.acell("B2").value
        b3 = ws.acell("B3").value
        b2 = str(b2).strip() if b2 and str(b2).strip() != "" else None
        b3 = str(b3).strip() if b3 and str(b3).strip() != "" else None
        return b2, b3
    except Exception:
        return None, None

def get_headers_and_df_from_ws(ws):
    """
    Retorna (headers_list, df) para uma worksheet gspread.
    Usa get_all_values() para preservar ordem de colunas.
    """
    vals = ws.get_all_values()
    if not vals:
        return [], pd.DataFrame()
    headers = [str(h).strip() for h in vals[0]]
    rows = vals[1:]
    df = pd.DataFrame(rows, columns=headers)
    return headers, df

def get_colname_by_letter_from_values_header(headers, letter):
    """
    Dado headers (lista) retorna o header correspondente à letra A..Z -> index 0..
    """
    if not headers:
        return None
    idx = ord(letter.upper()) - ord("A")
    if idx < 0 or idx >= len(headers):
        return None
    return headers[idx]

def detect_date_column_name(headers):
    """
    Tenta detectar coluna de data pelo nome: procura 'data' (case-insensitive).
    Retorna header encontrado ou None.
    """
    for h in headers:
        if isinstance(h, str) and "data" in h.strip().lower():
            return h
    return None

def safe_to_datetime_series(s):
    return pd.to_datetime(s, errors='coerce', dayfirst=True)

def filter_df_by_date_range(df, date_col_name, start_date, end_date):
    if date_col_name is None or date_col_name not in df.columns:
        return df  # sem filtro possível
    s = safe_to_datetime_series(df[date_col_name])
    mask = (s >= pd.to_datetime(start_date)) & (s <= pd.to_datetime(end_date))
    return df.loc[mask].copy()

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
except Exception:
    st.stop()

if s_ids:
    with st.spinner("Buscando planilhas e verificando abas..."):
        planilhas = list_spreadsheets_in_folders(drive_service, s_ids)
        if planilhas:
            df = pd.DataFrame(planilhas).sort_values("name").reset_index(drop=True)
            df = df.rename(columns={"name": "Planilha", "id": "ID_Planilha"})
            
            # preenche coluna 'conf' apenas para exibição (não revalidar na atualização)
            conf_map = get_conf_map(df["ID_Planilha"].tolist(), TARGET_SHEET_NAME)
            df["conf"] = df["ID_Planilha"].map(conf_map).astype(bool)
            
            # Seleção global (checkboxes que afetam as colunas)
            st.markdown('<div class="global-selection-container">', unsafe_allow_html=True)
            c1, c2, c3, _ = st.columns([1.2, 1.2, 1.2, 5])
            with c1: s_desc = st.checkbox("Desconto", value=True)
            with c2: s_mp = st.checkbox("Meio Pagto", value=True)
            with c3: s_fat = st.checkbox("Faturamento", value=True)
            st.markdown('</div>', unsafe_allow_html=True)
            
            df["Desconto"], df["Meio Pagamento"], df["Faturamento"] = s_desc, s_mp, s_fat
            
            # Configuração de colunas do data_editor (conf fica disabled)
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
            
            # -------- BOTÃO DE ATUALIZAÇÃO (LÓGICA PRINCIPAL) --------
            if st.button("🚀 INICIAR ATUALIZAÇÃO", use_container_width=True):
                # Monta df_final pelas duas metades do editor
                df_final = pd.concat([edit_esq, edit_dir], ignore_index=True)
                
                # Filtra só as planilhas marcadas para atualização (qualquer uma das 3 colunas)
                df_marcadas = df_final[
                    (df_final["Desconto"] == True) |
                    (df_final["Meio Pagamento"] == True) |
                    (df_final["Faturamento"] == True)
                ].copy()
                
                total = len(df_marcadas)
                if total == 0:
                    st.warning("Nenhuma planilha marcada para atualização.")
                    st.stop()
                
                # Abre planilha origem uma vez e obtém headers + df
                try:
                    sh_origem = gc.open_by_key(ID_PLANILHA_ORIGEM)
                    ws_origem = sh_origem.worksheet(ABA_ORIGEM)
                    headers_origem, df_origem = get_headers_and_df_from_ws(ws_origem)
                except Exception as e:
                    st.error(f"Erro ao abrir planilha origem: {e}")
                    st.stop()
                
                # Detectar coluna de data na origem
                col_data = detect_date_column_name(headers_origem) or "Data"
                # Mapear colunas por letra: F (grupo) e D (loja) a partir do header da origem
                col_grupo = get_colname_by_letter_from_values_header(headers_origem, "F")
                col_loja = get_colname_by_letter_from_values_header(headers_origem, "D")
                
                if not col_grupo:
                    st.error("Não foi possível identificar a coluna de Grupo (coluna F) na origem. Verifique o header da aba origem.")
                    st.stop()
                
                # Filtra pela data escolhida no UI
                df_origem = filter_df_by_date_range(df_origem, col_data, data_de, data_ate)
                
                progresso = st.progress(0)
                logs = []
                
                # Use enumerate para índice sequencial 0..total-1 ao atualizar progresso
                for i, (_, row) in enumerate(df_marcadas.iterrows()):
                    try:
                        nome_planilha = row["Planilha"]
                        id_dest = row["ID_Planilha"]
                        
                        # Abre planilha destino
                        try:
                            sh_destino = gc.open_by_key(id_dest)
                        except Exception as e:
                            logs.append(f"{nome_planilha}: Erro abrindo planilha destino -> {e}")
                            progresso.progress(min((i+1)/total, 1.0))
                            continue
                        
                        # Lê B2/B3 da aba Configurações Não Apagar da planilha destino
                        b2, b3 = read_codes_from_config_sheet(sh_destino)
                        if not b2:
                            logs.append(f"{nome_planilha}: B2 (código do grupo) não encontrado em '{TARGET_SHEET_NAME}'. Pulando.")
                            progresso.progress(min((i+1)/total, 1.0))
                            continue
                        
                        # Filtra df_origem por grupo (col_grupo) e opcionalmente por loja (col_loja)
                        df_filtro = df_origem[df_origem[col_grupo].astype(str).str.strip().str.upper() == str(b2).strip().upper()].copy()
                        if b3 and col_loja:
                            df_filtro = df_filtro[df_filtro[col_loja].astype(str).str.strip().str.upper() == str(b3).strip().upper()].copy()
                        
                        if df_filtro.empty:
                            logs.append(f"{nome_planilha}: Nenhum registro encontrado para grupo '{b2}'{(' e loja ' + b3) if b3 else ''}.")
                            progresso.progress(min((i+1)/total, 1.0))
                            continue
                        
                        # --- Atualiza aba Importado_Fat na planilha destino (APAGA SÓ PERÍODO) ---
                        try:
                            # obtém worksheet destino (cria se não existir)
                            try:
                                ws_dest = sh_destino.worksheet("Importado_Fat")
                            except Exception:
                                ws_dest = sh_destino.add_worksheet(title="Importado_Fat", rows=max(1000, len(df_filtro)+10), cols=max(10, len(df_filtro.columns)))
                            
                            # lê dados atuais da aba destino
                            headers_dest, df_dest = get_headers_and_df_from_ws(ws_dest)
                            
                            if df_dest.empty:
                                # Se não há dados, apenas escreve o filtrado
                                headers_to_write = df_filtro.columns.tolist()
                                values = [headers_to_write] + df_filtro[headers_to_write].values.tolist()
                                ws_dest.clear()
                                ws_dest.update("A1", values)
                                logs.append(f"{nome_planilha}: Importado_Fat criado com {len(df_filtro)} linhas.")
                            else:
                                # detecta coluna de data na planilha destino (preferir header do destino)
                                col_data_dest = detect_date_column_name(headers_dest) or col_data
                                
                                # Se existir coluna de grupo/loja no destino, usá-las para condicional de remoção
                                has_col_grupo_dest = col_grupo in headers_dest
                                has_col_loja_dest = col_loja in headers_dest
                                
                                # converte colunas de data para datetime para filtragem
                                df_dest[col_data_dest] = safe_to_datetime_series(df_dest[col_data_dest])
                                
                                # cria máscara para manter linhas fora do período ou que não correspondam ao grupo/loja alvo
                                mask_keep = pd.Series([True]*len(df_dest), index=df_dest.index)
                                
                                # condição para linhas que estão no período
                                cond_period = (df_dest[col_data_dest] >= pd.to_datetime(data_de)) & (df_dest[col_data_dest] <= pd.to_datetime(data_ate))
                                
                                # condição para matching por grupo/loja quando colunas existirem
                                if has_col_grupo_dest:
                                    cond_group = df_dest[col_grupo].astype(str).str.strip().str.upper() == str(b2).strip().upper()
                                else:
                                    cond_group = pd.Series([True]*len(df_dest), index=df_dest.index)
                                if b3 and has_col_loja_dest:
                                    cond_store = df_dest[col_loja].astype(str).str.strip().str.upper() == str(b3).strip().upper()
                                else:
                                    cond_store = pd.Series([True]*len(df_dest), index=df_dest.index)
                                
                                # linhas que devem ser removidas = in period AND matches group AND matches store
                                to_remove = cond_period & cond_group & cond_store
                                
                                # keep = not to_remove
                                df_dest_keep = df_dest.loc[~to_remove].copy()
                                
                                # Agora prepare df_filtro para concatenar: alinhar colunas com destino
                                # Preservar ordem de colunas do destino se possível
                                headers_to_write = headers_dest.copy()
                                # adicionar colunas do filtro que não existam ainda, mantendo ordem
                                for c in df_filtro.columns:
                                    if c not in headers_to_write:
                                        headers_to_write.append(c)
                                
                                # reindex ambos dataframes para headers_to_write e preencher vazios
                                df_dest_keep = df_dest_keep.reindex(columns=headers_to_write).fillna("")
                                df_filtro_prep = df_filtro.reindex(columns=headers_to_write).fillna("")
                                
                                # concatenar: manter df_dest_keep (linhas fora do período) + df_filtro_prep (novas linhas)
                                df_combined = pd.concat([df_dest_keep, df_filtro_prep], ignore_index=True)
                                
                                # opcional: ordenar por data se col_data_dest existir nas colunas
                                if col_data_dest in df_combined.columns:
                                    try:
                                        df_combined[col_data_dest] = safe_to_datetime_series(df_combined[col_data_dest])
                                        df_combined = df_combined.sort_values(by=col_data_dest).reset_index(drop=True)
                                        # ao escrever, transformar datas para string no formato ISO ou dd/mm/YYYY conforme preferência
                                        df_combined[col_data_dest] = df_combined[col_data_dest].dt.strftime("%Y-%m-%d").fillna("")
                                    except Exception:
                                        pass
                                
                                # escrever de volta na planilha destino
                                values = [headers_to_write] + df_combined[headers_to_write].astype(str).values.tolist()
                                ws_dest.clear()
                                ws_dest.update("A1", values)
                                
                                logs.append(f"{nome_planilha}: Importado_Fat atualizado - {len(df_filtro)} linhas substituídas no período; total agora {len(df_combined)}.")
                        except Exception as e:
                            logs.append(f"{nome_planilha}: Erro escrevendo Importado_Fat -> {e}")
                    except Exception as e:
                        logs.append(f"{row.get('Planilha', '??')}: Erro geral -> {e}")
                    
                    progresso.progress(min((i+1)/total, 1.0))
                
                st.success("Atualização concluída!")
                st.write("\n".join(logs))
        else:
            st.warning("Nenhuma planilha encontrada nas subpastas selecionadas.")
