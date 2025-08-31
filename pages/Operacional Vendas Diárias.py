# pages/OperacionalVendasDiarias.py



import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import plotly.express as px
from datetime import date
st.set_page_config(page_title="Vendas Diarias", layout="wide")

# 🔒 Bloqueia o acesso caso o usuário não esteja logado
if not st.session_state.get("acesso_liberado"):
    st.stop()
from contextlib import contextmanager
st.set_page_config(page_title="Spinner personalizado | MMR Consultoria")
import streamlit as st

# ======================
# CSS para esconder só a barra superior
# ======================
st.markdown("""
    <style>
        [data-testid="stToolbar"] {
            visibility: hidden;
            height: 0%;
            position: fixed;
        }
        .stSpinner {
            visibility: visible !important;
        }
    </style>
""", unsafe_allow_html=True)

# ======================
# Spinner durante todo o processamento
# ======================
with st.spinner("⏳ Processando..."):

    # ================================
    # 1. Conexão com Google Sheets
    # ================================
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gc = gspread.authorize(credentials)
    planilha_empresa = gc.open("Vendas diarias")
    df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())
    
    # ================================
    # 2. Configuração inicial do app
    # ================================
    
    #st.title("📋 Relatório de Vendas Diarias")
    
    # 🎨 Estilizar abas
    st.markdown("""
        <style>
        .stApp { background-color: #f9f9f9; }
        div[data-baseweb="tab-list"] { margin-top: 20px; }
        button[data-baseweb="tab"] {
            background-color: #f0f2f6;
            border-radius: 10px;
            padding: 10px 20px;
            margin-right: 10px;
            transition: all 0.3s ease;
            font-size: 16px;
            font-weight: 600;
        }
        button[data-baseweb="tab"]:hover { background-color: #dce0ea; color: black; }
        button[data-baseweb="tab"][aria-selected="true"] { background-color: #0366d6; color: white; }
        </style>
    """, unsafe_allow_html=True)
    
    try:
        planilha = gc.open("Vendas diarias")
        aba_fat = planilha.worksheet("Fat Sistema Externo")
        data_raw = aba_fat.get_all_values()
    
        # Converte para DataFrame e define o cabeçalho
        if len(data_raw) > 1:
            df = pd.DataFrame(data_raw[1:], columns=data_raw[0])  # usa a primeira linha como header
    
            # Limpa espaços extras nos nomes de colunas
            df.columns = df.columns.str.strip()
    
            # Verifica se coluna "Data" está presente
            if "Data" in df.columns:
                df["Data"] = pd.to_datetime(df["Data"].astype(str).str.strip(), dayfirst=True, errors="coerce")
    
    
                ultima_data_valida = df["Data"].dropna()
    
                if not ultima_data_valida.empty:
                    ultima_data = ultima_data_valida.max().strftime("%d/%m/%Y")
    
                    # Corrige coluna Grupo
                    df["Grupo"] = df["Grupo"].astype(str).str.strip().str.lower()
                    df["GrupoExibicao"] = df["Grupo"].apply(
                        lambda g: "Bares" if g in ["amata", "aurora"]
                        else "Kopp" if g == "kopp"
                        else "GF4"
                    )
    
                    # Contagem de lojas únicas por grupo
                    df_ultima_data = df[df["Data"] == df["Data"].max()]
                    contagem = df_ultima_data.groupby("GrupoExibicao")["Loja"].nunique().to_dict()
                    qtde_bares = contagem.get("Bares", 0)
                    qtde_kopp = contagem.get("Kopp", 0)
                    qtde_gf4 = contagem.get("GF4", 0)
    
                    resumo_msg = f"""
                    <div style='font-size:13px; color:gray; margin-bottom:10px;'>
                    📅 Última atualização: {ultima_data} — Bares ({qtde_bares}), Kopp ({qtde_kopp}), GF4 ({qtde_gf4})
                    </div>
                    """
                    st.markdown(resumo_msg, unsafe_allow_html=True)
                else:
                    st.info("⚠️ Nenhuma data válida encontrada.")
            else:
                st.info("⚠️ Coluna 'Data' não encontrada no Google Sheets.")
        else:
            st.info("⚠️ Tabela vazia.")
    except Exception as e:
        st.error(f"❌ Erro ao processar dados do Google Sheets: {e}")
    
    # Cabeçalho bonito (depois do estilo)
    st.markdown("""
        <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
            <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
            <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatório Vendas Diarias</h1>
        </div>
    """, unsafe_allow_html=True)
    
    
    # ================================
    # 3. Separação em ABAS
    # ================================
    aba1, aba3, aba4 = st.tabs(["📄 Upload e Processamento", "🔄 Atualizar Google Sheets","📊 Auditar integração Everest"])
    
    # ================================
    # 📄 Aba 1 - Upload e Processamento
    # ================================
    
    with aba1:
        uploaded_file = st.file_uploader(
            "📁 Clique para selecionar ou arraste aqui o arquivo Excel com os dados de faturamento",
            type=["xls", "xlsx"]
        )    
    
        if uploaded_file:
            try:
                xls = pd.ExcelFile(uploaded_file)
                abas = xls.sheet_names
    
                if "FaturamentoDiarioPorLoja" in abas:
                    df_raw = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None)
                    texto_b1 = str(df_raw.iloc[0, 1]).strip().lower()
                    if texto_b1 != "faturamento diário sintético multi-loja":
                        st.error(f"❌ A célula B1 está com '{texto_b1}'. Corrija para 'Faturamento diário sintético multi-loja'.")
                        st.stop()
    
                    df = pd.read_excel(xls, sheet_name="FaturamentoDiarioPorLoja", header=None, skiprows=4)
                    df.iloc[:, 2] = pd.to_datetime(df.iloc[:, 2], dayfirst=True, errors='coerce')
    
                    registros = []
                    col = 3
                    while col < df.shape[1]:
                        nome_loja = str(df_raw.iloc[3, col]).strip()
                        if re.match(r"^\d+\s*-?\s*", nome_loja):
                            nome_loja = nome_loja.split("-", 1)[-1].strip()
                            header_col = str(df.iloc[0, col]).strip().lower()
                            if "fat.total" in header_col:
                                for i in range(1, df.shape[0]):
                                    linha = df.iloc[i]
                                    valor_data = df.iloc[i, 2]
                                    valor_check = str(df.iloc[i, 1]).strip().lower()
                                    if pd.isna(valor_data) or valor_check in ["total", "subtotal"]:
                                        continue
                                    valores = linha[col:col+5].values
                                    if pd.isna(valores).all():
                                        continue
                                    registros.append([
                                        valor_data, nome_loja, *valores,
                                        valor_data.strftime("%b"), valor_data.year
                                    ])
                            col += 5
                        else:
                            col += 1
    
                    if len(registros) == 0:
                        st.warning("⚠️ Nenhum registro encontrado.")
    
                    df_final = pd.DataFrame(registros, columns=[
                        "Data", "Loja", "Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas", "Ticket", "Mês", "Ano"
                    ])
    
                elif "Relatório 100132" in abas:
                    df = pd.read_excel(xls, sheet_name="Relatório 100132")
                    df["Loja"] = df["Código - Nome Empresa"].astype(str).str.split("-", n=1).str[-1].str.strip().str.lower()
                    df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
                    df["Fat.Total"] = pd.to_numeric(df["Valor Total"], errors="coerce")
                    df["Serv/Tx"] = pd.to_numeric(df["Taxa de Serviço"], errors="coerce")
                    df["Fat.Real"] = df["Fat.Total"] - df["Serv/Tx"]
                    df["Ticket"] = pd.to_numeric(df["Ticket Médio"], errors="coerce")
    
                    df_agrupado = df.groupby(["Data", "Loja"]).agg({
                        "Fat.Total": "sum",
                        "Serv/Tx": "sum",
                        "Fat.Real": "sum",
                        "Ticket": "mean"
                    }).reset_index()
    
                    df_agrupado["Mês"] = df_agrupado["Data"].dt.strftime("%b").str.lower()
                    df_agrupado["Ano"] = df_agrupado["Data"].dt.year
                    df_final = df_agrupado
    
                else:
                    st.error("❌ O arquivo enviado não contém uma aba reconhecida. Esperado: 'FaturamentoDiarioPorLoja' ou 'Relatório 100113'.")
                    st.stop()
    
                dias_traducao = {
                    "Monday": "segunda-feira", "Tuesday": "terça-feira", "Wednesday": "quarta-feira",
                    "Thursday": "quinta-feira", "Friday": "sexta-feira", "Saturday": "sábado", "Sunday": "domingo"
                }
                df_final.insert(1, "Dia da Semana", pd.to_datetime(df_final["Data"], dayfirst=True, errors='coerce').dt.day_name().map(dias_traducao))
                df_final["Data"] = pd.to_datetime(df_final["Data"], dayfirst=True, errors='coerce').dt.strftime("%d/%m/%Y")
    
                for col_val in ["Fat.Total", "Serv/Tx", "Fat.Real", "Pessoas", "Ticket"]:
                    if col_val in df_final.columns:
                        df_final[col_val] = pd.to_numeric(df_final[col_val], errors="coerce").round(2)
    
                meses = {"jan": "jan", "feb": "fev", "mar": "mar", "apr": "abr", "may": "mai", "jun": "jun",
                         "jul": "jul", "aug": "ago", "sep": "set", "oct": "out", "nov": "nov", "dec": "dez"}
                df_final["Mês"] = df_final["Mês"].str.lower().map(meses)
    
                df_final["Data_Ordenada"] = pd.to_datetime(df_final["Data"], format="%d/%m/%Y", errors='coerce')
                df_final = df_final.sort_values(by=["Data_Ordenada", "Loja"]).drop(columns="Data_Ordenada")
    
                df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower()
                df_final["Loja"] = df_final["Loja"].astype(str).str.strip().str.lower()
                df_final = pd.merge(df_final, df_empresa, on="Loja", how="left")
    
                colunas_finais = [
                    "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
                    "Código Grupo Everest", "Fat.Total", "Serv/Tx", "Fat.Real",
                    "Ticket", "Mês", "Ano"
                ]
                df_final = df_final[colunas_finais]
    
                st.session_state.df_final = df_final
                st.session_state.atualizou_google = False
    
                datas_validas = pd.to_datetime(df_final["Data"], format="%d/%m/%Y", errors='coerce').dropna()
                if not datas_validas.empty:
                    data_inicial = datas_validas.min().strftime("%d/%m/%Y")
                    data_final_str = datas_validas.max().strftime("%d/%m/%Y")
                    valor_total = df_final["Fat.Total"].sum().round(2)
                    valor_total_formatado = f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown(f"""
                            <div style='font-size:24px; font-weight: bold; margin-bottom:10px;'>🗓️ Período processado</div>
                            <div style='font-size:30px; color:#000;'>{data_inicial} até {data_final_str}</div>
                        """, unsafe_allow_html=True)
                    with col2:
                        st.markdown(f"""
                            <div style='font-size:24px; font-weight: bold; margin-bottom:10px;'>💰 Valor total</div>
                            <div style='font-size:30px; color:green;'>{valor_total_formatado}</div>
                        """, unsafe_allow_html=True)
                else:
                    st.warning("⚠️ Não foi possível identificar o período de datas.")
    
                empresas_nao_localizadas = df_final[df_final["Código Everest"].isna()]["Loja"].unique()
                if len(empresas_nao_localizadas) > 0:
                    empresas_nao_localizadas_str = "<br>".join(empresas_nao_localizadas)
                    mensagem = f"""
                    ⚠️ {len(empresas_nao_localizadas)} empresa(s) não localizada(s), cadastre e reprocesse novamente! <br>{empresas_nao_localizadas_str}
                    <br>✏️ Atualize a tabela clicando 
                    <a href='https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU/edit?usp=drive_link' target='_blank'><strong>aqui</strong></a>.
                    """
                    st.markdown(mensagem, unsafe_allow_html=True)
                else:
                    st.success("✅ Todas as empresas foram localizadas na Tabela_Empresa!")
    
                    def to_excel(df):
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                            df.to_excel(writer, index=False, sheet_name='Faturamento Servico')
                        output.seek(0)
                        return output
    
                    excel_data = to_excel(df_final)
    
                    st.download_button(
                        label="📥 Baixar Relatório Excel",
                        data=excel_data,
                        file_name="faturamento_servico.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    
            except Exception as e:
                st.error(f"❌ Erro ao processar o arquivo: {e}")
    
    
    
    
# =======================================
# 🔄 Atualizar Google Sheets (ABA 3)
# =======================================

with aba3:
    import streamlit as st
    import pandas as pd
    import numpy as np
    import json, re, unicodedata, uuid
    from datetime import date, datetime, timedelta
    import requests
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry

    import gspread
    from oauth2client.service_account import ServiceAccountCredentials
    from gspread_dataframe import get_as_dataframe
    from gspread_formatting import CellFormat, NumberFormat, format_cell_range

    # ====== CONFIG DO SHEET (use SEMPRE o ID) ======
    SHEET_ID = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
    WS_FAT   = "Fat Sistema Externo"
    WS_TBEMP = "Tabela Empresa"

    # ====== ESTILO ======
    def _inject_button_css():
        st.markdown("""
        <style>
          div.stButton > button, div.stLinkButton > a {
            background-color: #e0e0e0 !important;
            color: #000 !important;
            border: 1px solid #b3b3b3 !important;
            border-radius: 6px !important;
            padding: 0.35em 0.6em !important;
            font-size: 0.85rem !important;
            font-weight: 600 !important;
            min-height: 30px !important;
            height: 30px !important;
            width: 100% !important;
            box-shadow: none !important;
          }
          div.stButton > button:hover, div.stLinkButton > a:hover { background-color: #d6d6d6 !important; }
          div.stButton > button:active, div.stLinkButton > a:active { background-color: #c2c2c2 !important; }
          div.stButton > button:disabled { background-color: #f0f0f0 !important; color:#666 !important; }
        </style>
        """, unsafe_allow_html=True)
    if "css_buttons_applied" not in st.session_state:
        _inject_button_css()
        st.session_state["css_buttons_applied"] = True

    # ====== DEBUG / STATUS ======
    MODO_DEBUG = st.sidebar.toggle("🔍 Modo debug (Aba Google Sheets)", value=False, key="dbg_operacional_aba3")
    def dlog(msg, data=None):
        if MODO_DEBUG:
            st.caption(f"🧪 {msg}")
            if data is not None:
                try:
                    import json as _json
                    st.code(_json.dumps(data, ensure_ascii=False, indent=2) if not isinstance(data, str) else data, language="json")
                except Exception:
                    st.code(str(data))

    def _show_status_banner():
        status = st.session_state.get("_gs_update_status")
        if not status: return
        (st.success if status.get("ok") else st.error)(status.get("msg", ""))
        extra = status.get("extra")
        if extra:
            try:
                import json as _json
                st.code(_json.dumps(extra, ensure_ascii=False, indent=2), language="json")
            except Exception:
                st.write(extra)
    def _set_status(ok: bool, msg: str, extra: dict|None=None):
        st.session_state["_gs_update_status"] = {"ok": bool(ok), "msg": str(msg), "extra": extra or {}}
    _show_status_banner()

    # ====== HELPERS GOOGLE ======
    def fetch_with_retry(url, connect_timeout=10, read_timeout=180, retries=3, backoff=1.5):
        s = requests.Session()
        retry = Retry(total=retries, connect=retries, read=retries, backoff_factor=backoff,
                      status_forcelist=[429, 500, 502, 503, 504], allowed_methods=["GET"], raise_on_status=False)
        s.mount("https://", HTTPAdapter(max_retries=retry))
        try:
            return s.get(url, timeout=(connect_timeout, read_timeout), headers={"Accept": "text/plain"})
        finally:
            s.close()

    def get_gc():
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
        credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
        return gspread.authorize(credentials)

    def get_service_account_email():
        try:
            return json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"]).get("client_email", "")
        except Exception:
            return ""

    def open_dest(gc, ws_name=WS_FAT):
        sh = gc.open_by_key(SHEET_ID)
        ws = sh.worksheet(ws_name)
        dlog("Destino (Sheets)", {
            "spreadsheet_title": sh.title,
            "spreadsheet_id": sh.id,
            "worksheet_title": ws.title,
            "worksheet_id": ws.id,
            "svc_account": get_service_account_email()
        })
        return sh, ws

    # ====== NORMALIZAÇÕES ======
    def _norm_simple(s: str) -> str:
        s = str(s or "").strip().lower()
        s = unicodedata.normalize("NFD", s)
        s = "".join(c for c in s if unicodedata.category(c) != "Mn")
        s = re.sub(r"[^a-z0-9]+", " ", s).strip()
        return s
    def _norm_key(s: str) -> str:
        s = str(s or "")
        s = unicodedata.normalize("NFD", s)
        s = "".join(c for c in s if unicodedata.category(c) != "Mn")
        s = s.strip().lower()
        s = re.sub(r"[^a-z0-9]+", " ", s)
        return s.strip()
    def _fmt_serial_to_br(x):
        try:
            return pd.to_datetime(pd.Series([x]), origin="1899-12-30", unit="D", errors="coerce").dt.strftime("%d/%m/%Y").iloc[0]
        except Exception:
            return x
    def _normN(x):
        return str(x).strip().replace(".0", "")
    ALIASES = {
        "codigo everest": {"codigo everest", "cod everest", "codigo ev"},
        "cod grupo empresas": {"codigo grupo everest", "codigo grupo empresas", "cod grupo empresas"},
        "fat total": {"fat total", "fat.total"},
        "serv tx": {"serv tx", "serv/tx"},
        "fat real": {"fat real", "fat.real"},
        "mes": {"mes", "mês"},
    }
    def _alias_target(norm_name: str) -> str:
        for canon, variants in ALIASES.items():
            if norm_name == canon or norm_name in variants:
                return canon
        return norm_name
    def build_row_values(headers_raw, registro_dict):
        dnorm = {_norm_key(k): v for k, v in registro_dict.items()}
        out = []
        for h in headers_raw:
            h_norm = _alias_target(_norm_key(h))
            val = dnorm.get(h_norm, None)
            if val is None:
                val = registro_dict.get(str(h).strip(), "")
            out.append(val)
        return out

    # ====== CATÁLOGO ======
    def carregar_catalogo_codigos(gc):
        try:
            _, ws = open_dest(gc, WS_TBEMP)
            df = get_as_dataframe(ws, evaluate_formulas=True, dtype=str).fillna("")
            if df.empty:
                return pd.DataFrame(columns=["Loja","Loja_norm","Grupo","Código Everest","Código Grupo Everest"])
            df.columns = df.columns.str.strip()
            cols_norm = {c: _norm_simple(c) for c in df.columns}
            loja_col  = next((c for c,n in cols_norm.items() if "loja" in n), None)
            grupo_col = next((c for c,n in cols_norm.items() if n == "grupo" or "grupo" in n), None)
            cod_col   = next((c for c,n in cols_norm.items() if "codigo" in n and "everest" in n and "grupo" not in n), None)
            codg_col  = next((c for c,n in cols_norm.items() if "codigo" in n and "grupo" in n and "everest" in n), None)
            out = pd.DataFrame()
            if not loja_col:
                return pd.DataFrame(columns=["Loja","Loja_norm","Grupo","Código Everest","Código Grupo Everest"])
            out["Loja"] = df[loja_col].astype(str).str.strip()
            out["Loja_norm"] = out["Loja"].str.lower()
            out["Grupo"] = df[grupo_col].astype(str).str.strip() if grupo_col else ""
            out["Código Everest"] = pd.to_numeric(df[cod_col], errors="coerce") if cod_col else pd.NA
            out["Código Grupo Everest"] = pd.to_numeric(df[codg_col], errors="coerce") if codg_col else pd.NA
            return out
        except Exception as e:
            st.error(f"❌ Não foi possível carregar o catálogo de códigos: {e}")
            return pd.DataFrame(columns=["Loja","Loja_norm","Grupo","Código Everest","Código Grupo Everest"])

    def preencher_codigos_por_loja(df_manuais: pd.DataFrame, catalogo: pd.DataFrame) -> pd.DataFrame:
        df = df_manuais.copy()
        if df.empty or catalogo.empty or "Loja" not in df.columns:
            return df
        look = catalogo.set_index("Loja_norm")
        lojakey = df["Loja"].astype(str).str.strip().str.lower()
        if "Grupo" in look.columns:
            df["Grupo"] = lojakey.map(look["Grupo"]).fillna(df.get("Grupo", ""))
        if "Código Everest" in look.columns:
            df["Código Everest"] = lojakey.map(look["Código Everest"])
        if "Código Grupo Everest" in look.columns:
            df["Código Grupo Everest"] = lojakey.map(look["Código Grupo Everest"])
        return df

    # ====== TEMPLATE MANUAL ======
    def template_manuais(n: int = 10) -> pd.DataFrame:
        d0 = pd.Timestamp(date.today() - timedelta(days=1))
        df = pd.DataFrame({
            "Data":      pd.Series([d0]*n, dtype="datetime64[ns]"),
            "Loja":      pd.Series([""]*n, dtype="object"),
            "Fat.Total": pd.Series([0.0]*n, dtype="float"),
            "Serv/Tx":   pd.Series([0.0]*n, dtype="float"),
            "Fat.Real":  pd.Series([0.0]*n, dtype="float"),
            "Ticket":    pd.Series([0.0]*n, dtype="float"),
        })
        return df[["Data","Loja","Fat.Total","Serv/Tx","Fat.Real","Ticket"]]

    _DIA_PT = {0:"segunda-feira",1:"terça-feira",2:"quarta-feira",3:"quinta-feira",4:"sexta-feira",5:"sábado",6:"domingo"}
    def _mes_label_pt(dt: pd.Series) -> pd.Series:
        nomes = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"]
        return dt.dt.month.map(lambda m: nomes[m-1] if pd.notnull(m) else "")

    def preparar_manuais_para_envio(edited_df: pd.DataFrame, catalogo: pd.DataFrame) -> pd.DataFrame:
        if edited_df is None or edited_df.empty: return pd.DataFrame()
        df = edited_df.copy()
        df["Loja"] = df["Loja"].fillna("").astype(str).str.strip()
        df = df[df["Loja"] != ""]
        if df.empty: return df
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
        for c in ["Fat.Total","Serv/Tx","Fat.Real","Ticket"]:
            if c in df.columns:
                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
        df["Dia da Semana"] = df["Data"].dt.dayofweek.map(_DIA_PT).str.title()
        df["Mês"] = _mes_label_pt(df["Data"])
        df["Ano"] = df["Data"].dt.year
        df = preencher_codigos_por_loja(df, catalogo)
        cols_preferidas = ["Data","Dia da Semana","Loja","Código Everest","Grupo","Código Grupo Everest",
                           "Fat.Total","Serv/Tx","Fat.Real","Ticket","Mês","Ano"]
        cols = [c for c in cols_preferidas if c in df.columns] + [c for c in df.columns if c not in cols_preferidas]
        return df[cols]

    # ====== PAINEL DE TESTE DE ESCRITA (diagnóstico rápido) ======
    with st.expander("🔧 Teste direto de escrita no Google Sheets (diagnóstico)"):
        st.caption("Este teste faz 1 append e em seguida exclui a linha de teste. Use para checar permissões do service account.")
        if st.button("▶️ Rodar teste de escrita (append + delete)", key="btn_teste_escrita"):
            try:
                _, ws = open_dest(get_gc(), WS_FAT)
                antes = len(ws.col_values(1))
                # faz append de uma linha de teste
                marca = f"__TEST__ {datetime.now().isoformat(timespec='seconds')}"
                ws.append_row([marca, "ping"], value_input_option="USER_ENTERED")
                depois_append = len(ws.col_values(1))
                # encontra a linha do teste pelo valor da 1ª coluna
                colA = ws.col_values(1)
                idxs = [i+1 for i,v in enumerate(colA) if v == marca]  # 1-indexed
                if idxs:
                    ws.delete_rows(idxs[-1])
                depois_delete = len(ws.col_values(1))
                st.success("✅ Teste concluído.")
                dlog("Teste escrita", {"linhas_antes": antes, "apos_append": depois_append, "apos_delete": depois_delete, "svc_account": get_service_account_email()})
            except Exception as e:
                st.error(f"❌ Falha no teste de escrita: {e}")
                dlog("Erro teste escrita", str(e))

    # ====== ENVIO ======
    def enviar_para_sheets(df_input: pd.DataFrame, titulo_origem: str = "dados") -> bool:
        if df_input.empty:
            st.info("ℹ️ Nada a enviar."); return True

        with st.spinner(f"🔄 Processando {titulo_origem} e verificando duplicidades..."):
            df_final = df_input.copy()

            # Chave M
            try:
                df_final['M'] = pd.to_datetime(df_final['Data'], format='%d/%m/%Y').dt.strftime('%Y-%m-%d') \
                                + df_final['Fat.Total'].astype(str) + df_final['Loja'].astype(str)
            except Exception:
                _dt = pd.to_datetime(df_final['Data'], origin="1899-12-30", unit='D', errors="coerce")
                df_final['M'] = _dt.dt.strftime('%Y-%m-%d') + df_final['Fat.Total'].astype(str) + df_final['Loja'].astype(str)
            df_final['M'] = df_final['M'].astype(str).str.strip()

            # Numerificação
            for coln in ['Fat.Total','Serv/Tx','Fat.Real','Ticket']:
                if coln in df_final.columns:
                    df_final[coln] = pd.to_numeric(df_final[coln], errors="coerce").fillna(0.0)
            dt_parsed = pd.to_datetime(df_final['Data'].astype(str).replace("'", "", regex=True).str.strip(), dayfirst=True, errors="coerce")
            if dt_parsed.notna().any():
                df_final['Data'] = (dt_parsed - pd.Timestamp("1899-12-30")).dt.days
            def to_int_safe(x):
                try:
                    x_clean = str(x).replace("'", "").strip()
                    return int(float(x_clean)) if x_clean not in ("", "nan", "None") else ""
                except: return ""
            for c in ['Código Everest','Código Grupo Everest','Ano']:
                if c in df_final.columns:
                    df_final[c] = df_final[c].apply(to_int_safe)

            # Abre destino
            gc = get_gc()
            _, aba_destino = open_dest(gc, WS_FAT)

            # Lê existentes
            valores_existentes_df = get_as_dataframe(aba_destino, evaluate_formulas=True, dtype=str).fillna("")
            cols_exist = valores_existentes_df.columns.str.strip().tolist()
            dados_existentes   = set(valores_existentes_df["M"].astype(str).str.strip()) if "M" in cols_exist else set()
            dados_n_existentes = set(valores_existentes_df["N"].astype(str).str.strip()) if "N" in cols_exist else set()

            # Chave N
            df_final['Data_Formatada'] = pd.to_datetime(df_final['Data'], origin="1899-12-30", unit='D', errors="coerce").dt.strftime('%Y-%m-%d')
            if 'Código Everest' not in df_final.columns: df_final['Código Everest'] = ""
            df_final['N'] = (df_final['Data_Formatada'] + df_final['Código Everest'].astype(str)).astype(str).str.strip()
            df_final = df_final.drop(columns=['Data_Formatada'])

            # Cabeçalho real
            headers = aba_destino.row_values(1)
            lookup = {_norm_simple(h): h for h in headers}
            aliases = {
                "codigo everest": ["codigo everest", "codigo ev", "cod everest", "cod ev"],
                "codigo grupo everest": ["cod grupo empresas", "codigo grupo empresas", "codigo grupo", "cod grupo"],
                "fat total": ["fat.total", "fat total"],
                "serv tx": ["serv/tx", "serv tx"],
                "fat real": ["fat.real", "fat real"],
                "mes": ["mes", "mês"],
            }
            # Renomeia p/ casar com cabeçalho
            rename_map = {}
            for col in list(df_final.columns):
                k = _norm_simple(col)
                if k in lookup: rename_map[col] = lookup[k]; continue
                found = False
                for canonical, variations in aliases.items():
                    if k == canonical or k in variations:
                        for v in [canonical] + variations:
                            kv = _norm_simple(v)
                            if kv in lookup: rename_map[col] = lookup[kv]; found = True; break
                    if found: break
                if not found and ("codigo grupo" in k or "cod grupo" in k):
                    for cand in ("codigo grupo everest", "cod grupo empresas", "codigo grupo empresas"):
                        kc = _norm_simple(cand)
                        if kc in lookup: rename_map[col] = lookup[kc]; break
            if rename_map: df_final = df_final.rename(columns=rename_map)

            # Reindex
            extras = [c for c in df_final.columns if c not in headers]
            df_final = df_final.reindex(columns=headers + extras, fill_value="")

            # Classificação
            M_in = df_final["M"].astype(str).str.strip()
            N_in = df_final["N"].astype(str).str.strip()
            is_dup_M = M_in.isin(dados_existentes)
            is_dup_N = N_in.isin(dados_n_existentes)
            df_suspeitos = df_final.loc[(~is_dup_M) & is_dup_N].copy()
            df_novos     = df_final.loc[(~is_dup_M) & (~is_dup_N)].copy()
            df_dup_M     = df_final.loc[is_dup_M].copy()

            st.markdown("<div style='color:#a33; font-weight:500; margin-top:10px;'>🔴 Possíveis duplicidades (chave N)</div>", unsafe_allow_html=True)

            # ========== SUBSTITUIÇÃO POR N ==========
            if len(df_suspeitos) > 0:
                valores_existentes_df = valores_existentes_df.copy()
                cN = "N" if "N" in valores_existentes_df.columns else None
                if cN: valores_existentes_df[cN] = valores_existentes_df[cN].map(_normN)

                def _col_sheet(humano):
                    k = _norm_simple(humano)
                    return lookup[k] if k in lookup else None
                cData = _col_sheet("Data"); cLoja = _col_sheet("Loja")
                cCod  = _col_sheet("codigo everest"); cFat = _col_sheet("fat total")
                cM    = "M" if "M" in valores_existentes_df.columns else None

                entrada_por_n = {}
                for _, row in df_suspeitos.iterrows():
                    d = row.fillna("").to_dict()
                    nkey = _normN(d.get("N", ""))
                    if "Data" in d: d["Data"] = _fmt_serial_to_br(d["Data"])
                    entrada_por_n[nkey] = d

                conflitos_linhas = []
                for nkey, d_in in sorted(entrada_por_n.items()):
                    d_view = d_in.copy()
                    d_view["Origem"] = "🟢 Novo Arquivo"
                    d_view["N"] = nkey
                    conflitos_linhas.append(d_view)
                    if cN:
                        df_sh = valores_existentes_df[valores_existentes_df[cN] == nkey].copy()
                    else:
                        df_sh = valores_existentes_df.iloc[0:0].copy()
                    ren = {}
                    if cData: ren[cData] = "Data"
                    if cLoja: ren[cLoja] = "Loja"
                    if cCod:  ren[cCod]  = "Codigo Everest"
                    if cFat:  ren[cFat]  = "Fat.Total"
                    if cM:    ren[cM]    = "M"
                    if cN:    ren[cN]    = "N"
                    df_sh = df_sh.rename(columns=ren)
                    if "Data" in df_sh.columns:
                        try:
                            ser = pd.to_numeric(df_sh["Data"], errors="coerce")
                            if ser.notna().any():
                                df_sh["Data"] = pd.to_datetime(ser, origin="1899-12-30", unit="D", errors="coerce").dt.strftime("%d/%m/%Y")
                            else:
                                df_sh["Data"] = pd.to_datetime(df_sh["Data"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")
                        except Exception:
                            pass
                    for idx, r in df_sh.iterrows():
                        d_sh = r.fillna("").to_dict()
                        d_sh["Origem"] = "🔴 Google Sheets"
                        d_sh["N"] = nkey
                        d_sh["__sheet_row"] = int(idx) + 2
                        conflitos_linhas.append(d_sh)

                cols_show = ["Origem","N","Data","Loja","Codigo Everest","Fat.Total","M"]
                df_conf = pd.DataFrame(conflitos_linhas)
                df_conf_show = df_conf.reindex(columns=[c for c in cols_show if c in df_conf.columns], fill_value="")
                st.markdown("### 🔴 Revisão antes de substituir")
                st.dataframe(df_conf_show, use_container_width=True)

                if st.button("🧹 Excluir existentes e inserir 'Novo Arquivo' (substituir por N)", use_container_width=True, key="btn_substituir_por_n"):
                    try:
                        adicionados = 0
                        deletados   = 0

                        # reabre destino e lê estado atual
                        _, aba_destino = open_dest(get_gc(), WS_FAT)
                        headers = aba_destino.row_values(1)
                        num_cols = len(headers)

                # 1) EXCLUIR com 3 estratégias: __sheet_row, DF e varredura direta na COLUNA N
                rows_to_delete = set()
                
                # (a) linhas pré-mapeadas no grid
                if "__sheet_row" in df_conf.columns:
                    rows_to_delete.update(
                        df_conf.loc[df_conf["Origem"]=="🔴 Google Sheets","__sheet_row"]
                        .dropna().astype(int).tolist()
                    )
                
                # (b) fallback pelo DataFrame de existentes (índice + 2)
                if cN and not valores_existentes_df.empty:
                    for nkey in entrada_por_n.keys():
                        idxs = valores_existentes_df.index[valores_existentes_df[cN].astype(str).str.strip() == nkey].tolist()
                        rows_to_delete.update(i+2 for i in idxs)
                
                # (c) varredura DIRETA na COLUNA N do Sheet (mais robusto)
                try:
                    headers = aba_destino.row_values(1)
                    if "N" in headers:
                        n_col_idx = headers.index("N") + 1  # 1-indexed
                        colN = aba_destino.col_values(n_col_idx)
                        for nkey in entrada_por_n.keys():
                            # normaliza e compara
                            matches = [i+1 for i, v in enumerate(colN) if _normN(v) == _normN(nkey)]
                            rows_to_delete.update(matches)
                        dlog("Varredura direta na coluna N", {
                            "n_col_idx": n_col_idx,
                            "rows_found": sorted(rows_to_delete)
                        })
                except Exception as e:
                    st.warning(f"Não consegui varrer a coluna N diretamente: {e}")
                
                # remove cabeçalho caso tenha entrado por engano
                rows_to_delete.discard(1)
                
                # executa exclusão em ordem decrescente
                for row_idx in sorted(rows_to_delete, reverse=True):
                    try:
                        aba_destino.delete_rows(row_idx)
                        deletados += 1
                    except Exception as e:
                        st.error(f"❌ Erro ao excluir linha {row_idx}: {e}")


                        # 2) inserir: “Novo Arquivo” por N
                        for nkey, d_in in entrada_por_n.items():
                            row_values = build_row_values(headers, d_in)
                            if len(row_values) < num_cols: row_values += [""]*(num_cols - len(row_values))
                            elif len(row_values) > num_cols: row_values = row_values[:num_cols]
                            try:
                                aba_destino.append_row(row_values, value_input_option="USER_ENTERED")
                                adicionados += 1
                            except Exception as e:
                                st.error(f"❌ Erro ao inserir (N={nkey}): {e}")

                        # 3) enviar NOVOS sem conflito
                        enviados_novos = 0
                        if len(df_novos) > 0:
                            headers_envio = aba_destino.row_values(1)
                            payload = df_novos.reindex(columns=headers_envio).fillna("").astype(object).values.tolist()
                            if payload:
                                aba_destino.append_rows(payload, value_input_option="USER_ENTERED")
                                enviados_novos = len(payload)

                        st.success(
                            f"✅ Substituição concluída: {adicionados} inserido(s) | {deletados} excluído(s) | {enviados_novos} novo(s) sem conflito."
                        )
                        _set_status(True, "Google Sheets atualizado com sucesso.", {
                            "substituidos": adicionados, "excluidos": deletados, "novos_sem_conflito": enviados_novos
                        })
                    except Exception as e:
                        st.error(f"❌ Falha ao substituir: {e}")
                        _set_status(False, f"Falha ao substituir: {e}")
            else:
                st.caption("Sem suspeitos por N. Nada a substituir.")

            # ====== ENVIO DIRETO dos NOVOS (sem conflitos) ======
            def _is_na_code(x):
                s = str(x).strip()
                return (s == "" or s.lower() == "nan" or s == "0")
            if "Código Everest" in df_final.columns:
                lojas_nao_cadastradas = df_final.loc[df_final["Código Everest"].apply(_is_na_code), "Loja"].dropna().unique().tolist()
            else:
                lojas_nao_cadastradas = []
            todas_lojas_ok = len(lojas_nao_cadastradas) == 0
            pode_enviar_direto = (len(df_suspeitos) == 0)

            if todas_lojas_ok and pode_enviar_direto and len(df_novos) > 0:
                headers_envio = aba_destino.row_values(1)
                dados_para_enviar = df_novos.reindex(columns=headers_envio).fillna("").astype(object).values.tolist()
                try:
                    inicio = len(aba_destino.col_values(1)) + 1
                    aba_destino.append_rows(dados_para_enviar, value_input_option='USER_ENTERED')
                    fim = inicio + len(dados_para_enviar) - 1
                    if inicio <= fim:
                        data_format   = CellFormat(numberFormat=NumberFormat(type='DATE',   pattern='dd/mm/yyyy'))
                        numero_format = CellFormat(numberFormat=NumberFormat(type='NUMBER', pattern='0'))
                        format_cell_range(aba_destino, f"A{inicio}:A{fim}", data_format)
                        format_cell_range(aba_destino, f"D{inicio}:D{fim}", numero_format)
                        format_cell_range(aba_destino, f"F{inicio}:F{fim}", numero_format)
                        format_cell_range(aba_destino, f"L{inicio}:L{fim}", numero_format)
                    st.success(f"✅ {len(dados_para_enviar)} novo(s) enviados. ❌ {len(df_dup_M)} duplicado(s) por M.")
                    _set_status(True, f"{len(dados_para_enviar)} registro(s) enviados para o Google Sheets.", {
                        "enviados": len(dados_para_enviar), "duplicados_M": int(len(df_dup_M))
                    })
                except Exception as e:
                    st.error(f"❌ Erro ao fazer append_rows: {e}")
                    _set_status(False, f"Falha ao enviar para o Google Sheets: {e}")
            elif len(df_novos) == 0:
                st.info(f"ℹ️ 0 novos para enviar. ❌ {len(df_dup_M)} duplicado(s) por M.")
            elif not todas_lojas_ok:
                st.error("🚫 Há lojas sem **Código Everest** cadastradas.")
                _set_status(False, "Há lojas sem Código Everest cadastrado.")
            elif len(df_suspeitos) > 0:
                st.warning("⚠️ Existem suspeitos (chave N). Use o botão de substituição acima.")
                _set_status(False, "Há suspeitos por N — substitua antes do envio direto.")

            return True

    # ====== ESTADO / CONTROLES DA ABA ======
    if st.session_state.get("_last_tab") != "atualizar_google_sheets":
        st.session_state["show_manual_editor"] = False
    st.session_state["_last_tab"] = "atualizar_google_sheets"
    if "show_manual_editor" not in st.session_state:
        st.session_state.show_manual_editor = False
    if "manual_df" not in st.session_state:
        st.session_state.manual_df = template_manuais(10)

    LINK_SHEET = f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/edit?usp=sharing"
    has_df = ('df_final' in st.session_state and isinstance(st.session_state.df_final, pd.DataFrame) and not st.session_state.df_final.empty)

    c1, c2, c3, c4 = st.columns([1, 1, 1, 1])
    with c1:
        enviar_auto = st.button("Atualizar SheetsS", use_container_width=True, disabled=not has_df,
                                help=None if has_df else "Carregue os dados para habilitar", key="btn_enviar_auto_header")
    with c2:
        aberto = st.session_state.get("show_manual_editor", False)
        label_toggle = "❌ Fechar lançamentos" if aberto else "Lançamentos manuais"
        if st.button(label_toggle, key="btn_toggle_manual", use_container_width=True):
            novo_estado = not aberto
            st.session_state["show_manual_editor"] = novo_estado
            st.session_state.manual_df = template_manuais(10)
            st.rerun()
    with c3:
        try:
            st.link_button("Abrir Google Sheets", LINK_SHEET, use_container_width=True)
        except Exception:
            st.markdown(f"""<a href="{LINK_SHEET}" target="_blank">
                <button style="width:100%;background:#e0e0e0;color:#000;border:1px solid #b3b3b3;
                padding:0.45em;border-radius:6px;font-weight:600;cursor:pointer;width:100%;">
                Abrir Google Sheets</button></a>""", unsafe_allow_html=True)
    with c4:
        atualizar_dre = st.button("Atualizar DRE", use_container_width=True, key="btn_atualizar_dre",
                                  help="Dispara a atualização do DRE agora")
    if atualizar_dre:
        SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw-gK_KYcSyqyfimHTuXFLEDxKvWdW4k0o_kOPE-r-SWxL-SpogE2U9wiZt7qCZoH-gqQ/exec"
        try:
            with st.spinner("Atualizando DRE..."):
                resp = fetch_with_retry(SCRIPT_URL, connect_timeout=10, read_timeout=180, retries=3, backoff=1.5)
            if resp is None:
                st.error("❌ Sem resposta do servidor.")
            elif resp.status_code == 200:
                st.success("✅ DRE atualizada com sucesso!")
                st.caption(resp.text[:1000] if resp.text else "OK")
            else:
                st.error(f"❌ Erro HTTP {resp.status_code} ao executar o script.")
                if resp.text: st.caption(resp.text[:1000])
        except requests.exceptions.ReadTimeout:
            st.error("❌ Tempo limite de leitura atingido. Tente novamente.")
        except requests.exceptions.ConnectTimeout:
            st.error("❌ Tempo limite de conexão atingido.")
        except Exception as e:
            st.error(f"❌ Falha ao conectar: {e}")

    # ====== EDITOR MANUAL ======
    if st.session_state.get("show_manual_editor", False):
        st.subheader("Lançamentos manuais")
        gc_ = get_gc()
        catalogo = carregar_catalogo_codigos(gc_)
        lojas_options = sorted(catalogo["Loja"].dropna().astype(str).str.strip().unique().tolist()) if not catalogo.empty else []
        PLACEHOLDER_LOJA = "— selecione a loja —"
        lojas_options_ui = [PLACEHOLDER_LOJA] + lojas_options
        df_disp = st.session_state.manual_df.copy()
        df_disp["Loja"] = df_disp["Loja"].fillna("").astype(str).str.strip()
        df_disp.loc[df_disp["Loja"] == "", "Loja"] = PLACEHOLDER_LOJA
        df_disp["Data"] = pd.to_datetime(df_disp["Data"], errors="coerce")
        for c in ["Fat.Total","Serv/Tx","Fat.Real","Ticket"]:
            df_disp[c] = pd.to_numeric(df_disp[c], errors="coerce")
        df_disp = df_disp[["Data","Loja","Fat.Total","Serv/Tx","Fat.Real","Ticket"]]
        edited_df = st.data_editor(
            df_disp, num_rows="dynamic", use_container_width=True,
            column_config={
                "Data": st.column_config.DateColumn(format="DD/MM/YYYY"),
                "Loja": st.column_config.SelectboxColumn(options=lojas_options_ui, default=PLACEHOLDER_LOJA,
                                                         help="Clique e escolha a loja (digite para filtrar)"),
                "Fat.Total": st.column_config.NumberColumn(step=0.01),
                "Serv/Tx":   st.column_config.NumberColumn(step=0.01),
                "Fat.Real":  st.column_config.NumberColumn(step=0.01),
                "Ticket":    st.column_config.NumberColumn(step=0.01),
            }, key="editor_manual",
        )
        col_esq, _ = st.columns([2, 8])
        with col_esq:
            enviar_manuais = st.button("Salvar Lançamentos", key="btn_enviar_manual", use_container_width=True)
        if enviar_manuais:
            edited_df["Loja"] = edited_df["Loja"].replace({PLACEHOLDER_LOJA: ""}).astype(str).str.strip()
            df_pronto = preparar_manuais_para_envio(edited_df, catalogo)
            if df_pronto.empty:
                st.warning("Nenhuma linha com Loja preenchida para enviar.")
            else:
                ok = enviar_para_sheets(df_pronto, titulo_origem="manuais")
                if ok:
                    st.session_state.manual_df = template_manuais(10)
                    st.rerun()

    # ====== ENVIO AUTOMÁTICO (usa enviar_para_sheets) ======
    if enviar_auto:
        if 'df_final' not in st.session_state or st.session_state.df_final.empty:
            st.error("Não há dados para enviar.")
        else:
            df_auto = st.session_state.df_final.copy()
            ok = enviar_para_sheets(df_auto, titulo_origem="automático")
            if ok:
                st.success("✅ Processo finalizado.")


        
    
    
    
    # =======================================
    # Aba 4 - Integração Everest (independente do upload)
    # =======================================
    
    from datetime import date
    import streamlit as st
    import pandas as pd
    
    # =======================================
    # Aba 4 - Integração Everest (independente do upload)
    # =======================================
    
    with aba4:
        try:
            planilha = gc.open("Vendas diarias")
            aba_everest = planilha.worksheet("Everest")
            aba_externo = planilha.worksheet("Fat Sistema Externo")
    
            df_everest = pd.DataFrame(aba_everest.get_all_values()[1:])
            df_externo = pd.DataFrame(aba_externo.get_all_values()[1:])
    
            df_everest.columns = [f"col{i}" for i in range(df_everest.shape[1])]
            df_externo.columns = [f"col{i}" for i in range(df_externo.shape[1])]
    
            df_everest["col0"] = pd.to_datetime(df_everest["col0"], dayfirst=True, errors="coerce")
            df_externo["col0"] = pd.to_datetime(df_externo["col0"], dayfirst=True, errors="coerce")
    
            datas_validas = df_everest["col0"].dropna()
    
            if not datas_validas.empty:
               # Garantir objetos do tipo date
                datas_validas = pd.to_datetime(df_everest["col0"], errors="coerce").dropna()
                datas_validas = datas_validas.dt.date
    
                if not datas_validas.empty:
                   from datetime import date
    
                # Garantir tipo date para todas as datas
                datas_validas = pd.to_datetime(df_everest["col0"], errors="coerce").dropna().dt.date
    
                if not datas_validas.empty:
                    datas_validas = df_everest["col0"].dropna()
    
                    if not datas_validas.empty:
                        min_data = datas_validas.min().date()
                        max_data_planilha = datas_validas.max().date()
                        sugestao_data = max_data_planilha
                    
                        data_range = st.date_input(
                            label="Selecione o intervalo de datas:",
                            value=(sugestao_data, sugestao_data),
                            min_value=min_data,
                            max_value=max_data_planilha
                        )
                    
                        if isinstance(data_range, tuple) and len(data_range) == 2:
                            data_inicio, data_fim = data_range
                            # Aqui já segue direto o processamento normal
    
    
               
                    def tratar_valor(valor):
                        try:
                            return float(str(valor).replace("R$", "").replace(".", "").replace(",", ".").strip())
                        except:
                            return None
    
                    ev = df_everest.rename(columns={
                        "col0": "Data", "col1": "Codigo",
                        "col7": "Valor Bruto (Everest)", "col6": "Impostos (Everest)"
                    })
                    
                    # 🔥 Remove linhas do Everest que são Total/Subtotal
                    ev = ev[~ev["Codigo"].astype(str).str.lower().str.contains("total", na=False)]
                    ev = ev[~ev["Codigo"].astype(str).str.lower().str.contains("subtotal", na=False)]
                    
                    ex = df_externo.rename(columns={
                        "col0": "Data",
                        "col2": "Nome Loja Sistema Externo",
                        "col3": "Codigo",
                        "col6": "Valor Bruto (Externo)",
                        "col8": "Valor Real (Externo)"
                    })
    
                    ev["Data"] = pd.to_datetime(ev["Data"], errors="coerce").dt.date
                    ex["Data"] = pd.to_datetime(ex["Data"], errors="coerce").dt.date
    
                    ev = ev[(ev["Data"] >= data_inicio) & (ev["Data"] <= data_fim)].copy()
                    ex = ex[(ex["Data"] >= data_inicio) & (ex["Data"] <= data_fim)].copy()
    
                    for col in ["Valor Bruto (Everest)", "Impostos (Everest)"]:
                        ev[col] = ev[col].apply(tratar_valor)
                    for col in ["Valor Bruto (Externo)", "Valor Real (Externo)"]:
                        ex[col] = ex[col].apply(tratar_valor)
    
                    if "Impostos (Everest)" in ev.columns:
                        ev["Impostos (Everest)"] = pd.to_numeric(ev["Impostos (Everest)"], errors="coerce").fillna(0)
                        ev["Valor Real (Everest)"] = ev["Valor Bruto (Everest)"] - ev["Impostos (Everest)"]
                    else:
                        ev["Valor Real (Everest)"] = ev["Valor Bruto (Everest)"]
    
                    ev["Valor Bruto (Everest)"] = pd.to_numeric(ev["Valor Bruto (Everest)"], errors="coerce").round(2)
                    ev["Valor Real (Everest)"] = pd.to_numeric(ev["Valor Real (Everest)"], errors="coerce").round(2)
                    ex["Valor Bruto (Externo)"] = pd.to_numeric(ex["Valor Bruto (Externo)"], errors="coerce").round(2)
                    ex["Valor Real (Externo)"] = pd.to_numeric(ex["Valor Real (Externo)"], errors="coerce").round(2)
    
                    mapa_nome_loja = ex.drop_duplicates(subset="Codigo")[["Codigo", "Nome Loja Sistema Externo"]]\
                        .set_index("Codigo").to_dict()["Nome Loja Sistema Externo"]
                    ev["Nome Loja Everest"] = ev["Codigo"].map(mapa_nome_loja)
    
                    df_comp = pd.merge(ev, ex, on=["Data", "Codigo"], how="outer", suffixes=("_Everest", "_Externo"))
    
                    # 🔄 Comparação
                    df_comp["Valor Bruto Iguais"] = df_comp["Valor Bruto (Everest)"] == df_comp["Valor Bruto (Externo)"]
                    df_comp["Valor Real Iguais"] = df_comp["Valor Real (Everest)"] == df_comp["Valor Real (Externo)"]
                    
                    # 🔄 Criar coluna auxiliar só para lógica interna
                    df_comp["_Tem_Diferenca"] = ~(df_comp["Valor Bruto Iguais"] & df_comp["Valor Real Iguais"])
                    
                    # 🔥 Filtro para ignorar as diferenças do grupo Kopp (apenas nas diferenças)
                    df_comp["_Ignorar_Kopp"] = df_comp["Nome Loja Sistema Externo"].str.contains("kop", case=False, na=False)
                    df_comp_filtrado = df_comp[~(df_comp["_Tem_Diferenca"] & df_comp["_Ignorar_Kopp"])].copy()
                    
                    # 🔧 Filtro no Streamlit
                    opcao = st.selectbox("Filtro de diferenças:", ["Todas", "Somente com diferenças", "Somente sem diferenças"])
                    
                    if opcao == "Todas":
                        df_resultado = df_comp_filtrado.copy()
                    elif opcao == "Somente com diferenças":
                        df_resultado = df_comp_filtrado[df_comp_filtrado["_Tem_Diferenca"]].copy()
                    else:
                        df_resultado = df_comp_filtrado[~df_comp_filtrado["_Tem_Diferenca"]].copy()
                    
                    # 🔧 Remover as colunas auxiliares antes de exibir
                    df_resultado = df_resultado.drop(columns=["Valor Bruto Iguais", "Valor Real Iguais", "_Tem_Diferenca", "_Ignorar_Kopp"], errors='ignore')
                    
                    # 🔧 Ajuste de colunas para exibição
                    df_resultado = df_resultado[[
                        "Data",
                        "Nome Loja Everest", "Codigo", "Valor Bruto (Everest)", "Valor Real (Everest)",
                        "Nome Loja Sistema Externo", "Valor Bruto (Externo)", "Valor Real (Externo)"
                    ]].sort_values("Data")
                    
                    df_resultado.columns = [
                        "Data",
                        "Nome (Everest)", "Código", "Valor Bruto (Everest)", "Valor Real (Everest)",
                        "Nome (Externo)", "Valor Bruto (Externo)", "Valor Real (Externo)"
                    ]
                    
                    colunas_texto = ["Nome (Everest)", "Nome (Externo)"]
                    df_resultado[colunas_texto] = df_resultado[colunas_texto].fillna("")
                    df_resultado = df_resultado.fillna(0)
    
                    df_resultado = df_resultado.reset_index(drop=True)
    
                    # ✅ Aqui adiciona o Total do dia logo após cada dia
                    dfs_com_totais = []
                    for data, grupo in df_resultado.groupby("Data", sort=False):
                        dfs_com_totais.append(grupo)
                    
                        total_dia = {
                            "Data": data,
                            "Nome (Everest)": "Total do dia",
                            "Código": "",
                            "Valor Bruto (Everest)": grupo["Valor Bruto (Everest)"].sum(),
                            "Valor Real (Everest)": grupo["Valor Real (Everest)"].sum(),
                            "Nome (Externo)": "",
                            "Valor Bruto (Externo)": grupo["Valor Bruto (Externo)"].sum(),
                            "Valor Real (Externo)": grupo["Valor Real (Externo)"].sum(),
                        }
                        dfs_com_totais.append(pd.DataFrame([total_dia]))
                    
                    df_resultado_final = pd.concat(dfs_com_totais, ignore_index=True)
                    
                    # 🔄 E continua com seu Total Geral normalmente
                    linha_total = pd.DataFrame([{
                        "Data": "",
                        "Nome (Everest)": "Total Geral",
                        "Código": "",
                        "Valor Bruto (Everest)": ev["Valor Bruto (Everest)"].sum(),
                        "Valor Real (Everest)": ev["Valor Real (Everest)"].sum(),
                        "Nome (Externo)": "",
                        "Valor Bruto (Externo)": ex["Valor Bruto (Externo)"].sum(),
                        "Valor Real (Externo)": ex["Valor Real (Externo)"].sum()
                    }])
                    
                    df_resultado_final = pd.concat([df_resultado_final, linha_total], ignore_index=True)
    
                                    
                    st.session_state.df_resultado = df_resultado
                                          
                    # 🔹 Estilo linha: destacar se tiver diferença (em vermelho)
                    def highlight_diferenca(row):
                        if (row["Valor Bruto (Everest)"] != row["Valor Bruto (Externo)"]) or (row["Valor Real (Everest)"] != row["Valor Real (Externo)"]):
                            return ["background-color: #ff9999"] * len(row)  # vermelho claro
                        else:
                            return [""] * len(row)
                    
                    # 🔹 Estilo colunas: manter azul e rosa padrão
                    def destacar_colunas_por_origem(col):
                        if "Everest" in col:
                            return "background-color: #e6f2ff"
                        elif "Externo" in col:
                            return "background-color: #fff5e6"
                        else:
                            return ""
                    
                    # 🔹 Aplicar estilos
                    st.dataframe(
                        df_resultado_final.style
                            .apply(highlight_diferenca, axis=1)
                            .set_properties(subset=["Valor Bruto (Everest)", "Valor Real (Everest)"], **{"background-color": "#e6f2ff"})
                            .set_properties(subset=["Valor Bruto (Externo)", "Valor Real (Externo)"], **{"background-color": "#fff5e6"})
                            .format({
                                "Valor Bruto (Everest)": "R$ {:,.2f}",
                                "Valor Real (Everest)": "R$ {:,.2f}",
                                "Valor Bruto (Externo)": "R$ {:,.2f}",
                                "Valor Real (Externo)": "R$ {:,.2f}"
                            }),
                        use_container_width=True,
                        height=600
                    )
    
    
                    
            else:
                st.warning("⚠️ Nenhuma data válida encontrada nas abas do Google Sheets.")
    
        except Exception as e:
            st.error(f"❌ Erro ao carregar ou comparar dados: {e}")
    
        # ==================================
        # Botão download Excel estilizado
        # ==================================
        
        def to_excel_com_estilo(df):
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                df.to_excel(writer, index=False, sheet_name='Comparativo')
                workbook  = writer.book
                worksheet = writer.sheets['Comparativo']
        
                # Formatos
                formato_everest = workbook.add_format({'bg_color': '#e6f2ff'})
                formato_externo = workbook.add_format({'bg_color': '#fff5e6'})
                formato_dif     = workbook.add_format({'bg_color': '#ff9999'})
        
                # Formatar colunas Everest e Externo
                worksheet.set_column('D:E', 15, formato_everest)
                worksheet.set_column('G:H', 15, formato_externo)
        
                # Destacar linhas com diferença
                for row_num, row_data in enumerate(df.itertuples(index=False)):
                    if (row_data[3] != row_data[6]) or (row_data[4] != row_data[7]):
                        worksheet.set_row(row_num+1, None, formato_dif)
        
            output.seek(0)
            return output
        
            # botão de download
            excel_bytes = to_excel_com_estilo(df_resultado_final)
            st.download_button(
                label="📥 Baixar Excel",
                data=excel_bytes,
                file_name="comparativo_everest_externo.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
