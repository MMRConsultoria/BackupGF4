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
    aba1, aba3, aba4, aba5 = st.tabs(["📄 Upload e Processamento", "🔄 Atualizar Google Sheets","📊 Auditar integração Everest","📊 Auditar Faturamento X Meio Pagamento"])
    # ---- Helper para saber qual aba está ativa ----
    def marcar_aba_ativa(tab_key: str):
        st.session_state["_aba_ativa"] = tab_key
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
    # Atualizar Google Sheets (Evitar duplicação)
    # =======================================
    with aba3:
        # ----- IMPORTS -----
        import streamlit as st
        import pandas as pd
        import numpy as np
        import json, re, unicodedata
        from datetime import date, datetime, timedelta
        import requests
        from requests.adapters import HTTPAdapter
        from urllib3.util.retry import Retry
    
        import gspread
        from oauth2client.service_account import ServiceAccountCredentials
        from gspread_dataframe import get_as_dataframe
        from gspread_formatting import CellFormat, NumberFormat, format_cell_range
        
        
        # --- estado para o modo de conflitos + df/ids persistidos ---
        if "modo_conflitos" not in st.session_state:
            st.session_state.modo_conflitos = False
        if "conflitos_df_conf" not in st.session_state:
            st.session_state.conflitos_df_conf = None
        if "conflitos_spreadsheet_id" not in st.session_state:
            st.session_state.conflitos_spreadsheet_id = None
        if "conflitos_sheet_id" not in st.session_state:
            st.session_state.conflitos_sheet_id = None
        # --- estado da CONFERÊNCIA do sistema ---
        st.session_state.setdefault("conf_pendente", False)
        st.session_state.setdefault("conf_ok", False)
        st.session_state.setdefault("conf_params", {})   # {sistema, mes_num, ano}
        st.session_state.setdefault("show_conf_panel", False)
        # ======= AUTH =======
        def get_gc():
            scope = [
                "https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/drive",
            ]
            credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
            credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
            return gspread.authorize(credentials)

        
        


    
        # ------------------------ ESTILO (botões pequenos, cinza) ------------------------
        def _inject_button_css():
            st.markdown("""
            <style>
              div.stButton > button, div.stLinkButton > a {
                background-color: #e0e0e0 !important;
                color: #000 !important;
                border: 1px solid #b3b3b3 !important;
                border-radius: 4px !important;
                padding: 0.25em 0.5em !important;
                font-size: 0.8rem !important;
                font-weight: 500 !important;
                min-height: 28px !important;
                height: 28px !important;
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


        def _inject_conf_css():
            st.markdown("""
            <style>
              #conf-btn-area div.stButton > button {
                background-color: #d32f2f !important;
                color: #fff !important;
                border: 1px solid #b71c1c !important;
                border-radius: 6px !important;
                font-weight: 700 !important;
              }
              #conf-panel {
                padding:12px;border:1px solid #c62828;background:#ffebee;border-radius:8px;margin-top:8px;
              }
            </style>
            """, unsafe_allow_html=True)
        
        if "css_conf_applied" not in st.session_state:
            _inject_conf_css()
            st.session_state["css_conf_applied"] = True

        # ------------------------ RETRY helper ------------------------
        def fetch_with_retry(url, connect_timeout=10, read_timeout=180, retries=3, backoff=1.5):
            s = requests.Session()
            retry = Retry(
                total=retries, connect=retries, read=retries,
                backoff_factor=backoff,
                status_forcelist=[429, 500, 502, 503, 504],
                allowed_methods=["GET"], raise_on_status=False,
            )
            s.mount("https://", HTTPAdapter(max_retries=retry))
            try:
                return s.get(url, timeout=(connect_timeout, read_timeout), headers={"Accept": "text/plain"})
            finally:
                s.close()
    
        # ------------------------ Helpers p/ catálogo/manuais (iguais aos seus) ------------------------
        def _norm(s: str) -> str:
            s = str(s or "").strip()
            s = unicodedata.normalize("NFD", s)
            s = "".join(c for c in s if unicodedata.category(c) != "Mn")
            s = s.lower()
            s = re.sub(r"[^a-z0-9]+", " ", s).strip()
            return s
    
        def carregar_catalogo_codigos(gc, nome_planilha="Vendas diarias", aba_catalogo="Tabela Empresa"):
            try:
                ws = gc.open(nome_planilha).worksheet(aba_catalogo)
                df = get_as_dataframe(ws, evaluate_formulas=True, dtype=str).fillna("")
                if df.empty:
                    return pd.DataFrame(columns=["Loja","Loja_norm","Grupo","Código Everest","Código Grupo Everest"])
    
                df.columns = df.columns.str.strip()
                cols_norm = {c: _norm(c) for c in df.columns}
    
                loja_col  = next((c for c,n in cols_norm.items() if "loja" in n), None)
                if not loja_col:
                    return pd.DataFrame(columns=["Loja","Loja_norm","Grupo","Código Everest","Código Grupo Everest"])
    
                grupo_col = next((c for c,n in cols_norm.items() if n == "grupo" or "grupo" in n), None)
                cod_col   = next((c for c,n in cols_norm.items() if "codigo" in n and "everest" in n and "grupo" not in n), None)
                codg_col  = next((c for c,n in cols_norm.items() if "codigo" in n and "grupo" in n and "everest" in n), None)
    
                out = pd.DataFrame()
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
    
        def template_manuais(n: int = 1) -> pd.DataFrame:
            # sempre iniciar com 1 linha e Data em branco (NaT)
            df = pd.DataFrame({
                "Data":      pd.Series([pd.NaT]*n, dtype="datetime64[ns]"),
                "Loja":      pd.Series([""]*n, dtype="object"),
                "Fat.Total": pd.Series([0.0]*n, dtype="float"),
                "Serv/Tx":   pd.Series([0.0]*n, dtype="float"),
                "Fat.Real":  pd.Series([0.0]*n, dtype="float"),
                "Ticket":    pd.Series([0.0]*n, dtype="float"),
            })
            return df[["Data","Loja","Fat.Total","Serv/Tx","Fat.Real","Ticket"]]

        # ======== Conferência: helpers ========
        MESES_PT = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"]
        
        def _parse_money(txt: str) -> float:
            import re
            s = str(txt or "").strip()
            s = s.replace("R$", "").replace("\u00A0","")
            s = re.sub(r"[^\d,.\-]", "", s)
            if s.count(",") == 1 and s.count(".") >= 1:
                s = s.replace(".", "").replace(",", ".")
            elif s.count(",") == 1 and s.count(".") == 0:
                s = s.replace(",", ".")
            try:
                return float(s)
            except:
                return float("nan")
        
        def _coerce_float(x):
            import re
            s = str(x or "").strip().replace("\u00A0", "")
            s = re.sub(r"[^\d,.\-]", "", s)
            if s == "":
                return float("nan")
            if s.count(",") == 1 and s.count(".") >= 1:
                s = s.replace(".", "").replace(",", ".")
            elif s.count(",") == 1 and s.count(".") == 0:
                s = s.replace(",", ".")
            try:
                return float(s)
            except:
                return float("nan")
        
        def _ns_header(s: str) -> str:
            import unicodedata, re
            s = str(s or "").strip().lower()
            s = unicodedata.normalize("NFD", s)
            s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
            return re.sub(r"[^a-z0-9]+", " ", s).strip()
        def _fmt_brl(v: float) -> str:
            try:
                v = float(v)
            except:
                return "R$ 0,00"
            s = f"{v:,.2f}"
            # vira padrão brasileiro: milhar com ponto e decimal com vírgula
            s = s.replace(",", "_").replace(".", ",").replace("_", ".")
            return f"R$ {s}"

        def inferir_sistema_mes_ano(df: pd.DataFrame):
            # Sistema
            if "Sistema" in df.columns and df["Sistema"].astype(str).str.strip().ne("").any():
                sistema = df["Sistema"].astype(str).str.strip().mode().iloc[0]
            else:
                grp = df.get("Grupo", pd.Series([], dtype="object")).astype(str).str.lower()
                sistema = "CISS" if grp.str.contains(r"\bkopp\b", regex=True).any() else "Colibri"
        
            # Mês/Ano
            dt = pd.to_datetime(df.get("Data", pd.Series([], dtype="object")).astype(str), dayfirst=True, errors="coerce")
            if dt.notna().any():
                mes_num = int(dt.dt.month.mode().iloc[0])
                ano     = int(dt.dt.year.mode().iloc[0])
            else:
                mes_num, ano = datetime.now().month, datetime.now().year
            return sistema, mes_num, ano
        
        def obter_total_sheet_por_sistema_mes(gc, sistema: str, ano: int, mes_num: int) -> float:
            """Soma Fat.Total na aba 'Fat Sistema Externo' filtrando por Sistema + Mês + Ano (não exibe o valor)."""
            sh = gc.open("Vendas diarias")
            ws = sh.worksheet("Fat Sistema Externo")
            df = get_as_dataframe(ws, evaluate_formulas=True, dtype=str).fillna("")
            if df.empty:
                return 0.0
        
            df.columns = df.columns.str.strip()
            col_data = next((c for c in df.columns if _ns_header(c) == "data"), None)
            col_sis  = next((c for c in df.columns if _ns_header(c) == "sistema"), None)
            col_fat  = next((c for c in df.columns if _ns_header(c) in ("fat total","fat total")), None)
            col_mes  = next((c for c in df.columns if _ns_header(c) in ("mes","mês")), None)
            col_ano  = next((c for c in df.columns if _ns_header(c) == "ano"), None)
        
            # Deriva mês/ano
            if col_data:
                ser = pd.to_numeric(df[col_data], errors="coerce")
                dt  = pd.to_datetime(ser, origin="1899-12-30", unit="D", errors="coerce")
                if not dt.notna().any():
                    dt = pd.to_datetime(df[col_data], dayfirst=True, errors="coerce")
                df["_mes_num"] = dt.dt.month
                df["_ano_num"] = dt.dt.year
            else:
                mapa_mes = {"jan":1,"fev":2,"mar":3,"abr":4,"mai":5,"jun":6,"jul":7,"ago":8,"set":9,"out":10,"nov":11,"dez":12,
                            "janeiro":1,"fevereiro":2,"marco":3,"março":3,"abril":4,"maio":5,"junho":6,"julho":7,"agosto":8,"setembro":9,"outubro":10,"novembro":11,"dezembro":12}
                df["_mes_num"] = df[col_mes].astype(str).str.lower().map(mapa_mes) if col_mes else pd.NA
                df["_ano_num"] = pd.to_numeric(df[col_ano], errors="coerce") if col_ano else pd.NA
        
            if not col_sis or not col_fat:
                return 0.0
        
            sis_norm = df[col_sis].astype(str).str.strip().str.lower()
            alvo = str(sistema or "").strip().lower()
        
            mask = (sis_norm == alvo) & (df["_mes_num"] == int(mes_num)) & (df["_ano_num"] == int(ano))
            sub = df.loc[mask, col_fat].apply(_coerce_float)
            total = float(np.nansum(sub.values))
            return round(total, 2)

    
        _DIA_PT = {0:"segunda-feira",1:"terça-feira",2:"quarta-feira",3:"quinta-feira",4:"sexta-feira",5:"sábado",6:"domingo"}
        def _mes_label_pt(dt: pd.Series) -> pd.Series:
            nomes = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"]
            return dt.dt.month.map(lambda m: nomes[m-1] if pd.notnull(m) else "")
    
        def preparar_manuais_para_envio(edited_df: pd.DataFrame, catalogo: pd.DataFrame) -> pd.DataFrame:
            if edited_df is None or edited_df.empty:
                return pd.DataFrame()
            df = edited_df.copy()
        
            # Limpeza básica
            df["Loja"] = df["Loja"].fillna("").astype(str).str.strip()
            df = df[df["Loja"] != ""]
            if df.empty:
                return df
        
            # PARSE de data com dayfirst=True para não inverter mês/dia
            # (se já vier datetime do DateColumn, mantemos)
            if not np.issubdtype(df["Data"].dtype, np.datetime64):
                df["Data"] = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce")
        
            for c in ["Fat.Total","Serv/Tx","Fat.Real","Ticket"]:
                if c in df.columns:
                    df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
        
            # Derivados de data
            _DIA_PT = {0:"segunda-feira",1:"terça-feira",2:"quarta-feira",3:"quinta-feira",4:"sexta-feira",5:"sábado",6:"domingo"}
            nomes_mes = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"]
            df["Dia da Semana"] = df["Data"].dt.dayofweek.map(_DIA_PT).str.title()
            df["Mês"] = df["Data"].dt.month.map(lambda m: nomes_mes[m-1] if pd.notnull(m) else "")
            df["Ano"] = df["Data"].dt.year
        
            # Completar códigos a partir do catálogo
            df = preencher_codigos_por_loja(df, catalogo)
        
            cols_preferidas = [
                "Data","Dia da Semana","Loja","Código Everest","Grupo","Código Grupo Everest",
                "Fat.Total","Serv/Tx","Fat.Real","Ticket","Mês","Ano"
            ]
            cols = [c for c in cols_preferidas if c in df.columns] + [c for c in df.columns if c not in cols_preferidas]
            return df[cols]


        def enviar_para_sheets(df_input: pd.DataFrame, titulo_origem: str = "dados") -> bool:
            if df_input.empty:
                st.info("ℹ️ Nada a enviar.")
                return True
        
            with st.spinner(f""):
        
                df_final = df_input.copy()
            
                # >>> SISTEMA (preencher sempre que ausente OU vazio)
                if ("Sistema" not in df_final.columns) or df_final["Sistema"].astype(str).str.strip().eq("").all():
                    if str(titulo_origem).lower() == "manuais":
                        df_final["Sistema"] = "Lançamento manual"
                    else:
                        grp_norm = df_final.get("Grupo", "").astype(str).str.strip().str.lower()
                        df_final["Sistema"] = np.where(grp_norm.str.contains(r"\bkopp\b", regex=True), "CISS", "Colibri")
                # <<< fim SISTEMA
            
                # ===== 1) Preparos =====
                # Numéricos
                for coln in ["Fat.Total","Serv/Tx","Fat.Real","Ticket"]:
                    if coln in df_final.columns:
                        df_final[coln] = pd.to_numeric(df_final[coln], errors="coerce").fillna(0.0)
            
                # ==== DATA: interpretar como dd/mm/YYYY (se vier string) e converter para serial (1899-12-30)
                if not np.issubdtype(df_final["Data"].dtype, np.datetime64):
                    dt_parsed = pd.to_datetime(
                        df_final["Data"].astype(str).str.replace("'", "").str.strip(),
                        dayfirst=True, errors="coerce"
                    )
                else:
                    dt_parsed = df_final["Data"]
            
                # Guarda formato YYYY-MM-DD para M/N
                df_final["_Data_ymd"] = dt_parsed.dt.strftime("%Y-%m-%d")
            
                # Converte Data para serial do Google Sheets
                df_final["Data"] = (dt_parsed - pd.Timestamp("1899-12-30")).dt.days
            
                # ==== Normalização de códigos para string inteira (sem .0)
                if "Codigo Everest" in df_final.columns:
                    cod_col = "Codigo Everest"
                elif "Código Everest" in df_final.columns:
                    cod_col = "Código Everest"
                else:
                    df_final["Código Everest"] = ""
                    cod_col = "Código Everest"
            
                def to_intstr(x):
                    try:
                        s = str(x).strip()
                        if s in ("", "nan", "None"):
                            return ""
                        return str(int(float(s)))
                    except:
                        return ""
            
                df_final[cod_col] = df_final[cod_col].apply(to_intstr)
                if "Código Grupo Everest" in df_final.columns:
                    df_final["Código Grupo Everest"] = df_final["Código Grupo Everest"].apply(to_intstr)
                if "Ano" in df_final.columns:
                    df_final["Ano"] = df_final["Ano"].apply(to_intstr)
            
                # ==== M e N com a data correta (YYYY-MM-DD)
                fat_col = "Fat.Total" if "Fat.Total" in df_final.columns else "Fat Total"
                df_final[fat_col] = pd.to_numeric(df_final[fat_col], errors="coerce").fillna(0)
                loja_str = df_final["Loja"].astype(str)
            
                # M = yyyy-mm-dd + Fat.Total + Loja
                df_final["M"] = (df_final["_Data_ymd"].fillna("") +
                                 df_final[fat_col].astype(str) +
                                 loja_str).str.strip()
            
                # N = yyyy-mm-dd + Codigo
                df_final["N"] = (df_final["_Data_ymd"].fillna("") +
                                 df_final[cod_col].astype(str)).str.strip()
            
                # Limpa coluna auxiliar
                df_final.drop(columns=["_Data_ymd"], errors="ignore", inplace=True)

                # ===== 2) Conecta planilha =====
                gc = get_gc()
                planilha_destino = gc.open("Vendas diarias")
                aba_destino = planilha_destino.worksheet("Fat Sistema Externo")
        
                valores_existentes_df = get_as_dataframe(aba_destino, evaluate_formulas=True, dtype=str).fillna("")
                colunas_df_existente = valores_existentes_df.columns.str.strip().tolist()
                dados_existentes   = set(valores_existentes_df["M"].astype(str).str.strip()) if "M" in colunas_df_existente else set()
                dados_n_existentes = set(valores_existentes_df["N"].astype(str).str.strip()) if "N" in colunas_df_existente else set()
        
                # ===== 3) Garantir N (yyyy-mm-dd + Código) =====
                df_final['Data_Formatada'] = pd.to_datetime(
                    df_final['Data'], origin="1899-12-30", unit='D', errors="coerce"
                ).dt.strftime('%Y-%m-%d')
                if 'Código Everest' not in df_final.columns and 'Codigo Everest' not in df_final.columns:
                    df_final['Código Everest'] = ""
                # normaliza nome da coluna de código
                cod_col = 'Codigo Everest' if 'Codigo Everest' in df_final.columns else 'Código Everest'
                df_final['N'] = (df_final['Data_Formatada'] + df_final[cod_col].astype(str)).astype(str).str.strip()
                df_final = df_final.drop(columns=['Data_Formatada'], errors='ignore')
        
                # ===== 4) Alinhar ao cabeçalho do Sheet =====
                headers_raw = aba_destino.row_values(1)
                headers = [h.strip() for h in headers_raw]
        
                def _ns(s: str) -> str:
                    s = str(s or "").strip().lower()
                    s = unicodedata.normalize("NFD", s)
                    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
                    s = re.sub(r"[^a-z0-9]+", " ", s).strip()
                    return s
        
                lookup = {_ns(h): h for h in headers}
                aliases = {
                    "codigo everest": ["codigo ev", "cod everest", "código everest"],
                    "codigo grupo everest": ["cod grupo empresas", "codigo grupo", "cód grupo empresas"],
                    "fat total": ["fat.total", "fat total"],
                    "serv tx": ["serv/tx", "serv tx"],
                    "fat real": ["fat.real", "fat real"],
                    "mes": ["mês", "mes"],
                }
                rename_map = {}
                for col in list(df_final.columns):
                    k = _ns(col)
                    if k in lookup:
                        rename_map[col] = lookup[k]
                        continue
                    for canonical, variations in aliases.items():
                        if k == canonical or k in variations:
                            for v in [canonical] + variations:
                                kv = _ns(v)
                                if kv in lookup:
                                    rename_map[col] = lookup[kv]
                                    break
                            break
                if rename_map:
                    df_final = df_final.rename(columns=rename_map)
        
                # M/N finais (garantia)
                fat_col = "Fat.Total" if "Fat.Total" in df_final.columns else "Fat Total"
                df_final["Data_Formatada"] = pd.to_datetime(
                    df_final["Data"], origin="1899-12-30", unit="D", errors="coerce"
                ).dt.strftime("%Y-%m-%d")
                df_final[fat_col] = pd.to_numeric(df_final[fat_col], errors="coerce").fillna(0)
                df_final["M"] = (df_final["Data_Formatada"].fillna("") +
                                 df_final[fat_col].astype(str) +
                                 df_final["Loja"].astype(str)).str.strip()
                if "Codigo Everest" in df_final.columns:
                    df_final["Codigo Everest"] = pd.to_numeric(df_final["Codigo Everest"], errors="coerce").fillna(0).astype(int).astype(str)
                elif "Código Everest" in df_final.columns:
                    df_final["Código Everest"] = pd.to_numeric(df_final["Código Everest"], errors="coerce").fillna(0).astype(int).astype(str)
                cod_col = "Codigo Everest" if "Codigo Everest" in df_final.columns else "Código Everest"
                df_final["N"] = (df_final["Data_Formatada"].fillna("") + df_final[cod_col].astype(str)).str.strip()
                df_final = df_final.drop(columns=["Data_Formatada"], errors="ignore")
        
                # reindex
                extras = [c for c in df_final.columns if c not in headers]
                df_final = df_final.reindex(columns=headers + extras, fill_value="")
        
                # ===== 5) Classificar: NOVOS / DUP(M) / SUS(N) =====
                M_in = df_final["M"].astype(str).str.strip()
                N_in = df_final["N"].astype(str).str.strip()
                is_dup_M = M_in.isin(dados_existentes)
                is_dup_N = N_in.isin(dados_n_existentes)
        

                # classificar
                M_in = df_final["M"].astype(str).str.strip()
                N_in = df_final["N"].astype(str).str.strip()
                is_dup_M = M_in.isin(dados_existentes)
                is_dup_N = N_in.isin(dados_n_existentes)
                mask_suspeitos = (~is_dup_M) & is_dup_N
                mask_novos     = (~is_dup_M) & (~is_dup_N)
                
                df_suspeitos = df_final.loc[mask_suspeitos].copy()
                df_novos     = df_final.loc[mask_novos].copy()
                df_dup_M     = df_final.loc[is_dup_M].copy()  # <- para contar duplicados por M
                
                # === CONTAGENS ===
                q_novos = int(len(df_novos))
                q_dup_m = int(len(df_dup_M))
                q_sus_n = int(len(df_suspeitos))
                
                # === ENVIA NOVOS MESMO HAVENDO SUSPEITOS ===
                if q_novos > 0:
                    try:
                        dados_para_enviar = df_novos.fillna("").values.tolist()
                        inicio = len(aba_destino.col_values(1)) + 1
                        aba_destino.append_rows(dados_para_enviar, value_input_option='USER_ENTERED')
                        fim = inicio + q_novos - 1
                
                        if inicio <= fim:
                            data_format   = CellFormat(numberFormat=NumberFormat(type='DATE',   pattern='dd/mm/yyyy'))
                            numero_format = CellFormat(numberFormat=NumberFormat(type='NUMBER', pattern='0'))
                            format_cell_range(aba_destino, f"A{inicio}:A{fim}", data_format)
                            format_cell_range(aba_destino, f"D{inicio}:D{fim}", numero_format)
                            format_cell_range(aba_destino, f"F{inicio}:F{fim}", numero_format)
                            format_cell_range(aba_destino, f"L{inicio}:L{fim}", numero_format)
                
                        #st.success(f"✅ {q_novos} novo(s) enviado(s).")
                    except Exception as e:
                        st.error(f"❌ Erro ao enviar novos: {e}")
                #else:
                #    st.info("ℹ️ Nenhum novo para enviar.")
                
                # === SE HÁ SUSPEITOS POR N, PREPARA O PAINEL DE CONFLITOS ===
                if q_sus_n > 0:
                    #st.warning("🔎 Existem possíveis duplicados por N. Revise-os na seção de conflitos.")
                
                    # normaliza N do sheet para casar
                    def _normN(x): return str(x).strip().replace(".0", "")
                    valores_existentes_df2 = valores_existentes_df.copy()
                    if "N" in valores_existentes_df2.columns:
                        valores_existentes_df2["N"] = valores_existentes_df2["N"].map(_normN)
                
                    entrada_por_n = {}
                    for _, r in df_suspeitos.iterrows():
                        nkey = _normN(r.get("N",""))
                        entrada_por_n[nkey] = r.to_dict()
                
                    sheet_por_n = {nkey: valores_existentes_df2[valores_existentes_df2["N"] == nkey].copy()
                                   for nkey in entrada_por_n.keys()}
                
                    # canonização leve
                    def _norm_key(s: str) -> str:
                        s = str(s or "").strip().lower()
                        s = unicodedata.normalize("NFD", s)
                        s = "".join(c for c in s if unicodedata.category(c) != "Mn")
                        s = re.sub(r"[^a-z0-9]+", " ", s).strip()
                        return s
                
                    CANON_MAP = {
                        "codigo everest": "Codigo Everest",
                        "cod everest": "Codigo Everest",
                        "codigo ev": "Codigo Everest",
                        "código everest": "Codigo Everest",
                        "codigo grupo everest": "Cod Grupo Empresas",
                        "codigo grupo empresas": "Cod Grupo Empresas",
                        "cod grupo empresas": "Cod Grupo Empresas",
                        "código grupo everest": "Cod Grupo Empresas",
                        "cód grupo empresas": "Cod Grupo Empresas",
                        "cod grupo": "Cod Grupo Empresas",
                        "serv tx": "Serv/Tx",
                        "servico": "Serv/Tx",
                        "serv": "Serv/Tx",
                        "fat real": "Fat.Real",
                        "fat.real": "Fat.Real",
                        "loja": "Loja",
                        "grupo": "Grupo",
                        "data": "Data",
                        "mes": "Mês",
                        "mês": "Mês",
                        "ano": "Ano",
                        "dia da semana": "Dia da Semana",
                        "m": "M",
                        "n": "N",
                    }
                    def canon_colname(name: str) -> str:
                        if str(name).strip().lower() in ("fat.total","fat total"): return "Fat. Total"
                        k = _norm_key(name)
                        if k == "fat total": return "Fat. Total"
                        return CANON_MAP.get(k, name)
                    def canonize_cols_df(df: pd.DataFrame) -> pd.DataFrame:
                        return df.rename(columns={c: canon_colname(c) for c in df.columns}) if (df is not None and not df.empty) else df
                    def canonize_dict(d: dict) -> dict:
                        return {canon_colname(k): v for k, v in d.items()} if isinstance(d, dict) else d
                    def _fmt_serial_to_br(x):
                        try:
                            return pd.to_datetime(pd.Series([x]), origin="1899-12-30", unit="D", errors="coerce")\
                                     .dt.strftime("%d/%m/%Y").iloc[0]
                        except Exception:
                            try:
                                return pd.to_datetime(pd.Series([x]), dayfirst=True, errors="coerce")\
                                         .dt.strftime("%d/%m/%Y").iloc[0]
                            except Exception:
                                return x
                
                    conflitos_linhas = []
                    for nkey in sorted(entrada_por_n.keys()):
                        # entrada
                        d_in = canonize_dict(entrada_por_n[nkey].copy())
                        d_in["_origem_"] = "Nova Arquivo"
                        d_in["N"] = _normN(d_in.get("N",""))
                        if "Data" in d_in: d_in["Data"] = _fmt_serial_to_br(d_in["Data"])
                        conflitos_linhas.append(d_in)
                
                        # sheet
                        df_sh = sheet_por_n[nkey].copy()
                        df_sh = canonize_cols_df(df_sh)
                        if "Data" in df_sh.columns:
                            try:
                                ser = pd.to_numeric(df_sh["Data"], errors="coerce")
                                if ser.notna().any():
                                    df_sh["Data"] = pd.to_datetime(ser, origin="1899-12-30", unit="D", errors="coerce").dt.strftime("%d/%m/%Y")
                                else:
                                    df_sh["Data"] = pd.to_datetime(df_sh["Data"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")
                            except Exception:
                                pass
                        for idx, row in df_sh.iterrows():
                            d_sh = canonize_dict(row.to_dict())
                            d_sh["_origem_"] = "Google Sheets"
                            d_sh["Linha Sheet"] = idx + 2   # 1 = cabeçalho
                            conflitos_linhas.append(d_sh)
                
                    df_conf = pd.DataFrame(conflitos_linhas).copy()
                    df_conf = canonize_cols_df(df_conf)
                
                    # derivados de data
                    if "Data" in df_conf.columns:
                        _dt = pd.to_datetime(df_conf["Data"], dayfirst=True, errors="coerce")
                        nomes_dia = ["segunda-feira","terça-feira","quarta-feira","quinta-feira","sexta-feira","sábado","domingo"]
                        df_conf["Dia da Semana"] = _dt.dt.dayofweek.map(lambda i: nomes_dia[i].title() if pd.notna(i) else "")
                        nomes_mes = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"]
                        df_conf["Mês"] = _dt.dt.month.map(lambda m: nomes_mes[m-1] if pd.notna(m) else "")
                        df_conf["Ano"] = _dt.dt.year.fillna("").astype(str).replace("nan","")
                
                    if "_origem_" in df_conf.columns:
                        df_conf["_origem_"] = df_conf["_origem_"].replace({
                            "Nova Arquivo":"🟢 Nova Arquivo",
                            "Google Sheets":"🔴 Google Sheets"
                        })
                    if "Manter" not in df_conf.columns:
                        df_conf.insert(0,"Manter",False)
                
                    ordem_final = [
                        "Manter","_origem_","Linha Sheet","Data","Dia da Semana","Loja",
                        "Codigo Everest","Grupo","Cod Grupo Empresas",
                        "Fat. Total","Serv/Tx","Fat.Real","Ticket","Mês","Ano","M","N"
                    ]
                    df_conf = df_conf.reindex(
                        columns=[c for c in ordem_final if c in df_conf.columns] +
                                [c for c in df_conf.columns if c not in ordem_final],
                        fill_value=""
                    )
                
                    # ——— salvar estado p/ FASE 2 + resumo e rerun ———
                    st.session_state.conflitos_df_conf = df_conf
                    st.session_state.conflitos_spreadsheet_id = planilha_destino.id
                    st.session_state.conflitos_sheet_id = int(aba_destino.id)
                    st.session_state.modo_conflitos = True
                
                    # guarda o resumo para exibir no topo após o rerun
                    st.session_state._resumo_envio = {"enviados": q_novos, "dup_m": q_dup_m, "sus_n": q_sus_n}
                
                    st.rerun()
                # ✅ Defaults após as funções estarem definidas
                st.session_state.setdefault("modo_conflitos", False)
                st.session_state.setdefault("conflitos_df_conf", None)
                st.session_state.setdefault("conflitos_spreadsheet_id", None)
                st.session_state.setdefault("conflitos_sheet_id", None)
                st.session_state.setdefault("show_manual_editor", False)
                
                if "manual_df" not in st.session_state:
                    st.session_state["manual_df"] = template_manuais(5)
                # === SEM SUSPEITOS: limpar estado e concluir ===
                # === SEM SUSPEITOS: limpar estado e concluir ===
                st.session_state.modo_conflitos = False
                st.session_state.conflitos_df_conf = None
                st.session_state.conflitos_spreadsheet_id = None
                st.session_state.conflitos_sheet_id = None
                
                # resumo
                st.session_state._resumo_envio = {"enviados": q_novos, "dup_m": q_dup_m, "sus_n": q_sus_n}
                
                # >>> Agendar conferência do sistema (não mostrar o valor do Sheets)
                try:
                    # use df_final (já normalizado) para inferir
                    sis_hint, mes_hint, ano_hint = inferir_sistema_mes_ano(df_final)
                except Exception:
                    sis_hint, mes_hint, ano_hint = "Colibri", datetime.now().month, datetime.now().year
                
                st.session_state.update({
                    "conf_pendente": True,
                    "conf_ok": False,
                    "conf_params": {"sistema": str(sis_hint), "mes_num": int(mes_hint), "ano": int(ano_hint)},
                    "show_conf_panel": False,  # começa mostrando só o botão vermelho
                })
                # <<< fim agendamento
                
                return True



                   
              
                # Envia NOVOS (se existirem) e informa contagens de NOVOS e DUPLICADOS por M
                # Funciona tanto com df_novos/df_dup_M (DataFrames) quanto com novos_dados/duplicados (listas)
                

                # === ENVIO DE NOVOS + RESUMO (Enviados / Duplicados M / Possíveis duplicados N) ===
                # Suporta tanto DataFrames (df_novos/df_dup_M/df_suspeitos) quanto listas (novos_dados/duplicados/suspeitos_n)
                
                # --- 1) Preparar dados e contagens de forma resiliente ---
                # Novos
                if 'df_novos' in locals() and isinstance(df_novos, pd.DataFrame):
                    dados_para_enviar = df_novos.fillna("").values.tolist()
                    q_novos = len(dados_para_enviar)
                elif 'novos_dados' in locals():
                    dados_para_enviar = novos_dados
                    q_novos = len(novos_dados)
                else:
                    dados_para_enviar = []
                    q_novos = 0
                
                # Duplicados por M (ignorados)
                if 'df_dup_M' in locals() and isinstance(df_dup_M, pd.DataFrame):
                    q_dup_m = len(df_dup_M)
                elif 'duplicados' in locals():
                    q_dup_m = len(duplicados)
                else:
                    q_dup_m = 0
                
                # Possíveis duplicados por N (para revisão)
                if 'df_suspeitos' in locals() and isinstance(df_suspeitos, pd.DataFrame):
                    q_sus_n = len(df_suspeitos)
                elif 'suspeitos_n' in locals():
                    q_sus_n = len(suspeitos_n)
                else:
                    q_sus_n = 0
                
                # --- 2) Enviar automaticamente os NOVOS (se houver) ---
                # Se você usa a checagem de lojas, respeita-a; caso contrário, remove o bloco 'todas_lojas_ok'
                if 'todas_lojas_ok' in locals() and not todas_lojas_ok:
                    st.error("🚫 Há lojas sem **Código Everest** cadastradas. Corrija e tente novamente.")
                else:
                    if q_novos > 0:
                        try:
                            inicio = len(aba_destino.col_values(1)) + 1
                            aba_destino.append_rows(dados_para_enviar, value_input_option="USER_ENTERED")
                            fim = inicio + q_novos - 1
                
                            # Formatação (ajuste as colunas conforme seu Sheet)
                            if inicio <= fim:
                                data_format   = CellFormat(numberFormat=NumberFormat(type="DATE",   pattern="dd/mm/yyyy"))
                                numero_format = CellFormat(numberFormat=NumberFormat(type="NUMBER", pattern="0"))
                                format_cell_range(aba_destino, f"A{inicio}:A{fim}", data_format)   # Data
                                format_cell_range(aba_destino, f"D{inicio}:D{fim}", numero_format) # ex.: Fat.Total
                                format_cell_range(aba_destino, f"F{inicio}:F{fim}", numero_format)
                                format_cell_range(aba_destino, f"L{inicio}:L{fim}", numero_format)
                
                            st.success(f"✅ {q_novos} novo(s) enviado(s).")
                        except Exception as e:
                            st.error(f"❌ Erro ao enviar novos: {e}")
                    else:
                        st.info("ℹ️ Nenhum novo para enviar.")
                
                # --- 3) Resumo consolidado (sem listar os incluídos) ---
                #st.markdown(
                 #   f"**Resumo:** 🟢 Enviados: **{q_novos}** &nbsp;&nbsp;|&nbsp;&nbsp; "
                  #  f"❌ Duplicados : **{q_dup_m}** &nbsp;&nbsp;|&nbsp;&nbsp; "
                   # f"🔴 Possíveis duplicados : **{q_sus_n}**"
                #)
                
                
        # ------------------------ HEADER / BOTÕES ------------------------
        LINK_SHEET = "https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU/edit?usp=sharing"
        df_sess = st.session_state.get("df_final")
        has_df = isinstance(df_sess, pd.DataFrame) and not df_sess.empty  # ✅ simples e robusto

    
        c1, c2, c3, c4 = st.columns([1, 1, 1, 1])
    
        with c1:
            enviar_auto = st.button(
                "Atualizar Google Sheets",
                use_container_width=True,
                disabled=not has_df,
                key="btn_enviar_auto_header",
            )
        
        if enviar_auto:
            if not has_df:
                st.error("Não há dados para enviar.")
            else:
                ok = enviar_para_sheets(st.session_state.df_final.copy(), titulo_origem="upload")
                # Se houver suspeitos por N, a função faz st.rerun() e abrirá o painel.
                # Se não houver, cai aqui e mostramos o resumo.
        
        # ✅ Mostra o resumo sempre que existir (fora do if enviar_auto)
        r = st.session_state.get("_resumo_envio")
        if r:
            st.success(
                f"🟢 Enviados: **{r['enviados']}**  |  ❌ Duplicados: **{r['dup_m']}**  |  🔴 Suspeitos: **{r['sus_n']}**"
            )
           
        # badge de pendência (só visual)
        if st.session_state.get("conf_pendente", False) and not st.session_state.get("conf_ok", False):
            st.caption(":red[Conferência pendente do sistema]")
        
        # botão vermelho "Conferir sistema" que apenas abre o painel
        st.markdown('<div id="conf-btn-area">', unsafe_allow_html=True)
        abrir_conf = st.button("🔴 Conciliar Sistema PDV X GoogleSheets", use_container_width=True, key="btn_conf_sistema")
        st.markdown('</div>', unsafe_allow_html=True)
        if abrir_conf:
            st.session_state["show_conf_panel"] = True
        
        # painel (abre quando solicitado OU quando ainda pendente)
        if (st.session_state.get("conf_pendente", False) and not st.session_state.get("conf_ok", False)) \
           or st.session_state.get("show_conf_panel", False):
        
            st.markdown("""
            <div id="conf-panel">
              <div style="font-weight:700;color:#b71c1c;font-size:16px; margin-bottom:6px;">
                Conciliação do sistema PDV
              </div>
              <div style="color:#444;">Informe o total que está no sistema PDV</div>
            </div>
            """, unsafe_allow_html=True)

        
            params = st.session_state.get("conf_params", {}) or {}
            sis_default  = params.get("sistema", "Colibri")
            mes_default  = int(params.get("mes_num", datetime.now().month))
            ano_default  = int(params.get("ano", datetime.now().year))
        
            colA, colB, colC = st.columns([1,1,2])
            with colA:
                sistema_escolhido = st.selectbox("Sistema", ["Colibri","CISS","Lançamento manual"],
                    index=["Colibri","CISS","Lançamento manual"].index(sis_default) if sis_default in ["Colibri","CISS","Lançamento manual"] else 0)
            with colB:
                mes_num = st.selectbox("Mês", list(range(1,13)), index=max(0, mes_default-1),
                    format_func=lambda m: f"{m:02d} - {MESES_PT[m-1].title()}")
            with colC:
                ano_sel = st.number_input("Ano", min_value=2020, max_value=2100, value=ano_default, step=1)
        
            with st.form("form_conf_sistema", clear_on_submit=False):
                valor_usuario = st.text_input(
                    "Digite o TOTAL do sistema (use vírgula para centavos)",
                    key="conf_total_usuario",
                    placeholder="ex.: 123.456,78 ou 123456,78"
                )
                col_ok, col_skip = st.columns([1,1])
                confirmar = col_ok.form_submit_button("Conciliar Sistema PDV X GoogleSheets")
                fechar   = col_skip.form_submit_button("Fechar painel")
        
            if confirmar:
                if not str(valor_usuario or "").strip():
                    st.error("Informe o total para conferência.")
                else:
                    try:
                        total_user = _parse_money(valor_usuario)
                        gc_tmp = get_gc()
                        total_sheet = obter_total_sheet_por_sistema_mes(gc_tmp, sistema_escolhido, int(ano_sel), int(mes_num))
                        # Não mostramos o total do Sheets; apenas conferimos
                        # exige igualdade exata a 2 casas (0,01 já reprova)
                        if np.isfinite(total_user) and round(total_user, 2) == round(total_sheet, 2):
                            st.session_state.update({"conf_ok": True, "conf_pendente": False, "show_conf_panel": True})
                            st.success("✅ Conferência concluída.")
                        else:
                            # mostra a diferença sem revelar o total do Sheets
                            if np.isfinite(total_user) and np.isfinite(total_sheet):
                                diff = round(abs(total_user - total_sheet), 2)
                                st.error(
                                    f"🚫 Não confere com o registrado na planilha. "
                                    f"Verifique o total no sistema e tente novamente. Diferença apurada: {_fmt_brl(diff)}."
                                )
                            else:
                                st.error("🚫 Valor digitado inválido. Revise o total informado e tente novamente.")


                    except Exception as e:
                        st.error(f"Erro ao conferir: {e}")
        
            if fechar:
                st.session_state["show_conf_panel"] = False
                st.rerun()

    
        with c2:
            aberto = st.session_state.get("show_manual_editor", False)
            label_toggle = "❌ Fechar lançamentos" if aberto else "Lançamentos manuais"
            if st.button(label_toggle, key="btn_toggle_manual", use_container_width=True):
                novo_estado = not aberto
                st.session_state["show_manual_editor"] = novo_estado
                st.session_state.manual_df = template_manuais(5)
                st.rerun()
    
        with c3:
            try:
                st.link_button("Abrir Google Sheets", LINK_SHEET, use_container_width=True)
            except Exception:
                st.markdown(
                    f"""
                    <a href="{LINK_SHEET}" target="_blank">
                        <button style="width:100%;background:#e0e0e0;color:#000;border:1px solid #b3b3b3;
                        padding:0.45em;border-radius:6px;font-weight:600;cursor:pointer;width:100%;">
                        Abrir Google Sheets
                        </button>
                    </a>
                    """, unsafe_allow_html=True
                )
    
        with c4:
            atualizar_dre = st.button(
                "Atualizar DRE",
                use_container_width=True,
                key="btn_atualizar_dre",
                help="Dispara a atualização do DRE agora",
            )
        if atualizar_dre:
            SCRIPT_URL = "https://script.google.com/macros/s/AKfycbw-gK_KYcSyqyfimHTuXFLEDxKvWdW4k0o_kOPE-r-SWxL-SpogE2U9wiZt7qCZoH-gqQ/exec"
            try:
                with st.spinner("Atualizando DRE..."):
                    resp = fetch_with_retry(SCRIPT_URL, connect_timeout=10, read_timeout=180, retries=3, backoff=1.5)
                if resp is None:
                    st.error("❌ Falha inesperada: sem resposta do servidor.")
                elif resp.status_code == 200:
                    st.success("✅ DRE atualizada com sucesso!")
                    st.caption(resp.text[:1000] if resp.text else "OK")
                else:
                    st.error(f"❌ Erro HTTP {resp.status_code} ao executar o script.")
                    if resp.text:
                        st.caption(resp.text[:1000])
            except requests.exceptions.ReadTimeout:
                st.error("❌ Tempo limite de leitura atingido. Tente novamente.")
            except requests.exceptions.ConnectTimeout:
                st.error("❌ Tempo limite de conexão atingido. Verifique sua rede e tente novamente.")
            except Exception as e:
                st.error(f"❌ Falha ao conectar: {e}")

        # === Handler do botão superior "Atualizar SheetsS" ===
        if enviar_auto:
            if not has_df:
                st.error("Não há dados para enviar.")
            else:
                ok = enviar_para_sheets(st.session_state.df_final.copy(), titulo_origem="upload")
                # Se a função abriu o painel de conflitos, ela já dá st.rerun().
                # Só mostramos "concluído" quando NÃO ficou em modo de conflitos.
                #if ok and not st.session_state.get("modo_conflitos", False):
                    #st.success("✅ Processo concluído.")
        if "_resumo_envio" in st.session_state:
                    r = st.session_state._resumo_envio
                    #st.markdown(
                    #    f"**Resumo:** 🟢 Enviados: **{r['enviados']}** &nbsp;|&nbsp; "
                    #    f"❌ Duplicados (M): **{r['dup_m']}** &nbsp;|&nbsp; "
                    #    f"🔴 Possíveis duplicados (N): **{r['sus_n']}**"
                    #)
                    del st.session_state._resumo_envio
        # ========================== FASE 2: FORM DE CONFLITOS ==========================
        if st.session_state.get("modo_conflitos", False) and st.session_state.get("conflitos_df_conf") is not None:
            df_conf = st.session_state.conflitos_df_conf.copy()
    
            # sanear tipos p/ st.data_editor
            num_cols = ["Fat. Total","Serv/Tx","Fat.Real","Ticket"]
            for c in num_cols:
                if c in df_conf.columns:
                    df_conf[c] = pd.to_numeric(df_conf[c], errors="coerce").astype("Float64")
            if "Linha Sheet" in df_conf.columns:
                df_conf["Linha Sheet"] = pd.to_numeric(df_conf["Linha Sheet"], errors="coerce").astype("Int64")
            if "Manter" in df_conf.columns:
                df_conf["Manter"] = (
                    df_conf["Manter"].astype(str).str.strip().str.lower()
                    .isin(["true","1","yes","y","sim","verdadeiro"])
                ).astype("boolean")
            _proteger = set(num_cols + ["Linha Sheet","Manter"])
            for c in df_conf.columns:
                if c not in _proteger:
                    df_conf[c] = df_conf[c].astype("string").fillna("")
    
            st.markdown(
                "<div style='color:#555; font-size:0.9rem; font-weight:500; margin:10px 0;'>"
                "🔴 Possíveis duplicados — <b>marque as linhas do lado Google Sheets que você quer EXCLUIR</b>, 🟢 Marque para incluir os valores corretamente "
                "</div>",
                unsafe_allow_html=True
            )
    
            with st.form("form_conflitos_globais"):
                edited_conf = st.data_editor(
                    df_conf,
                    use_container_width=True,
                    hide_index=True,
                    key="editor_conflitos",
                    column_config={
                        "Manter": st.column_config.CheckboxColumn(
                            help="Marque as linhas 🔴 Google Sheets que deseja excluir",
                            default=False
                        )
                    }
                )
                aplicar_tudo = st.form_submit_button("🗑️ Google Sheets")
    
            if aplicar_tudo:
                try:
                    # --- IDs salvos quando você preparou os conflitos ---
                    spreadsheet_id = st.session_state.conflitos_spreadsheet_id
                    sheet_id = st.session_state.conflitos_sheet_id
                    if not spreadsheet_id or sheet_id is None:
                        st.error("❌ Ids da planilha não encontrados. Clique em 'Atualizar SheetsS' novamente.")
                        st.stop()
    
                    # --- Normalizações auxiliares ---
                    def _ns(s: str) -> str:
                        s = str(s or "").strip().lower()
                        s = unicodedata.normalize("NFD", s)
                        s = "".join(c for c in s if unicodedata.category(c) != "Mn")
                        s = re.sub(r"[^a-z0-9]+", " ", s).strip()
                        return s
    
                    def _to_serial_1899_12_30(date_str: str):
                        dt = pd.to_datetime(str(date_str), dayfirst=True, errors="coerce")
                        if pd.isna(dt):
                            return ""
                        return int((dt - pd.Timestamp("1899-12-30")).days)
    
                    def _yyyy_mm_dd(date_str: str):
                        dt = pd.to_datetime(str(date_str), dayfirst=True, errors="coerce")
                        return "" if pd.isna(dt) else dt.strftime("%Y-%m-%d")
    
                    def _to_num(x, default=0):
                        try:
                            v = pd.to_numeric(x, errors="coerce")
                            return default if pd.isna(v) else float(v)
                        except:
                            return default
    
                    def _to_intstr(x):
                        try:
                            if str(x).strip() == "":
                                return ""
                            return str(int(float(str(x).strip())))
                        except:
                            return ""
    
                    # --- Reabre planilha/aba e lê cabeçalho real ---
                    gc2 = get_gc()
                    sh2 = gc2.open_by_key(spreadsheet_id)
                    try:
                        ws2 = sh2.get_worksheet_by_id(int(sheet_id))
                    except Exception:
                        ws2 = None
                        for w in sh2.worksheets():
                            if int(getattr(w, "id", -1)) == int(sheet_id):
                                ws2 = w
                                break
                    if ws2 is None:
                        st.error("❌ Não consegui acessar a worksheet pelo sheetId.")
                        st.stop()
    
                    headers_raw = ws2.row_values(1)
                    headers = [h.strip() for h in headers_raw]
                    lookup = {_ns(h): h for h in headers}
    
                    # mapeamento de nomes “canônicos” do df_conf -> cabeçalho real do Sheet
                    def map_col(name: str):
                        k = _ns(name)
                        if k in lookup:
                            return lookup[k]
                        aliases = {
                            "codigo everest": ["código everest", "codigo ev", "cod everest", "cod ev"],
                            "codigo grupo everest": ["cod grupo empresas", "codigo grupo", "codigo grupo empresas"],
                            "fat total": ["fat.total", "fat total"],
                            "serv tx": ["serv/tx", "servico", "serv"],
                            "fat real": ["fat.real", "fat real"],
                            "mes": ["mês", "mes"],
                            "dia da semana": ["dia da semana"],
                        }
                        for canonical, vars_ in aliases.items():
                            if k == _ns(canonical) or k in [_ns(v) for v in vars_]:
                                for v in [canonical] + vars_:
                                    nv = _ns(v)
                                    if nv in lookup:
                                        return lookup[nv]
                        # casos especiais do seu df_conf
                        if k == _ns("Cod Grupo Empresas"):
                            nv = _ns("codigo grupo everest")
                            return lookup.get(nv)
                        if k == _ns("Codigo Everest"):
                            nv = _ns("codigo everest")
                            return lookup.get(nv)
                        return None  # sem mapeamento direto
    
                    # --- Separa marcados: excluir (Google) e incluir (Nova Arquivo) ---
                    manter = edited_conf["Manter"]
                    if manter.dtype != bool:
                        manter = manter.astype(str).str.strip().str.lower().isin(["true","1","yes","y","sim","verdadeiro"])
    
                    is_google = edited_conf["_origem_"].astype(str).str.contains("google", case=False, na=False)
                    is_novo   = edited_conf["_origem_"].astype(str).str.contains("nova",   case=False, na=False)
    
                    # 1) Linhas reais (1-based) a excluir (lado Google)
                    linhas_excluir = (
                        pd.to_numeric(edited_conf.loc[is_google & manter, "Linha Sheet"], errors="coerce")
                        .dropna().astype(int).tolist()
                    )
                    linhas_excluir = sorted({ln for ln in linhas_excluir if ln >= 2}, reverse=True)
    
                    # 2) Registros a incluir (lado Nova Arquivo)
                    novos_marcados = edited_conf.loc[is_novo & manter].copy()
                    # mapeia a coluna "Sistema" no cabeçalho do Sheet
                    col_sis = map_col("Sistema")
                    if not col_sis:
                        # fallback se o cabeçalho tiver variação/espacos/caixa
                        col_sis = next((h for h in headers if _ns(h) == "sistema"), None)

                    # Constrói linhas no formato do Sheet (ordem = headers)
                    rows_to_append = []
                    for _, r in novos_marcados.iterrows():
                        d = r.to_dict()
    
                        # pega/normaliza campos básicos
                        data_br   = d.get("Data", "")
                        loja      = str(d.get("Loja", "") or "")
                        grupo     = str(d.get("Grupo", "") or "")
                        fat_total = _to_num(d.get("Fat. Total", d.get("Fat.Total", 0)), 0)
                        serv_tx   = _to_num(d.get("Serv/Tx", 0), 0)
                        fat_real  = _to_num(d.get("Fat.Real", 0), 0)
                        ticket    = _to_num(d.get("Ticket", 0), 0)
                        mes_lbl   = str(d.get("Mês", "") or "")
                        ano_val   = _to_intstr(d.get("Ano", ""))
    
                        cod_emp   = _to_intstr(d.get("Codigo Everest", d.get("Código Everest", "")))
                        cod_grp   = _to_intstr(d.get("Cod Grupo Empresas", d.get("Código Grupo Everest", "")))
    
                        # Data: serial + yyyy-mm-dd para M/N
                        data_serial = _to_serial_1899_12_30(data_br)
                        data_ymd    = _yyyy_mm_dd(data_br)
    
                        # Calcula M/N se existirem no cabeçalho
                        col_M = lookup.get(_ns("M"))
                        col_N = lookup.get(_ns("N"))
                        M_val = (data_ymd + str(fat_total) + loja) if col_M else None
                        N_val = (data_ymd + cod_emp)             if col_N else None
    
                        # monta dict seguindo cabeçalho
                        row_out = {h: "" for h in headers}
                        # Data
                        col_data = map_col("Data")
                        if col_data:
                            row_out[col_data] = data_serial
                        # Dia da Semana
                        col_dia = map_col("Dia da Semana")
                        if col_dia:
                            row_out[col_dia] = d.get("Dia da Semana", "")
                        # Loja
                        col_loja = map_col("Loja")
                        if col_loja:
                            row_out[col_loja] = loja
                        # Grupo
                        col_grupo = map_col("Grupo")
                        if col_grupo:
                            row_out[col_grupo] = grupo
                        # Códigos
                        col_cod = map_col("Codigo Everest")
                        if col_cod:
                            row_out[col_cod] = cod_emp
                        col_codg = map_col("Cod Grupo Empresas")
                        if col_codg:
                            row_out[col_codg] = cod_grp
                        # Valores
                        col_fat = map_col("Fat. Total")
                        if col_fat:
                            row_out[col_fat] = fat_total
                        col_srv = map_col("Serv/Tx")
                        if col_srv:
                            row_out[col_srv] = serv_tx
                        col_real = map_col("Fat.Real")
                        if col_real:
                            row_out[col_real] = fat_real
                        col_tic = map_col("Ticket")
                        if col_tic:
                            row_out[col_tic] = ticket
                        # Mês/Ano
                        col_mes = map_col("Mês")
                        if col_mes:
                            row_out[col_mes] = mes_lbl
                        col_ano = map_col("Ano")
                        if col_ano:
                            row_out[col_ano] = ano_val
                        # M / N
                        if col_M:
                            row_out[col_M] = M_val
                        if col_N:
                            row_out[col_N] = N_val
                        # --- Sistema ---
                        if col_sis:
                            sis_val = str(d.get("Sistema", "") or "").strip()
                            if not sis_val:
                                # fallback: se vier vazio do df_conf, deduz a partir do Grupo
                                grp = str(d.get("Grupo", "") or "")
                                sis_val = "CISS" if re.search(r"kopp", grp, flags=re.I) else "Colibri"
                            row_out[col_sis] = sis_val
                        # garante que a ordem é a do headers
                        rows_to_append.append([row_out[h] for h in headers])
                    
                    # --- Executa: 1) excluir, 2) incluir ---
                    removidos = 0
                    inseridos = 0
    
                    # Exclusão (se houver)
                    if linhas_excluir:
                        reqs_del = [{
                            "deleteDimension": {
                                "range": {
                                    "sheetId": int(sheet_id),
                                    "dimension": "ROWS",
                                    "startIndex": ln - 1,
                                    "endIndex": ln
                                }
                            }
                        } for ln in linhas_excluir]
                        sh2.batch_update({"requests": reqs_del})
                        removidos = len(linhas_excluir)
    
                    # Inclusão (se houver)
                    if rows_to_append:
                        ws2.append_rows(rows_to_append, value_input_option='USER_ENTERED')
                        inseridos = len(rows_to_append)
    
                    if removidos == 0 and inseridos == 0:
                        st.info("ℹ️ Nada a fazer: nenhuma linha marcada para excluir ou incluir.")
                    #else:
                        #st.success(f"✅ Concluído: {removidos} excluída(s) e {inseridos} incluída(s).")
                        #st.markdown(
                        #    f"[Abrir a aba no Google Sheets](https://docs.google.com/spreadsheets/d/{sh2.id}/edit#gid={sheet_id})"
                        #)
    
                    # limpa estado e encerra
                    # >>> Agendar conferência do sistema após aplicar conflitos
                    try:
                        if 'novos_marcados' in locals() and isinstance(novos_marcados, pd.DataFrame) and not novos_marcados.empty:
                            sis_hint, mes_hint, ano_hint = inferir_sistema_mes_ano(novos_marcados)
                        else:
                            p = st.session_state.get("conf_params", {}) or {}
                            sis_hint = p.get("sistema", "Colibri")
                            mes_hint = int(p.get("mes_num", datetime.now().month))
                            ano_hint = int(p.get("ano", datetime.now().year))
                    except Exception:
                        sis_hint, mes_hint, ano_hint = "Colibri", datetime.now().month, datetime.now().year
                    
                    st.session_state.update({
                        "conf_pendente": True,
                        "conf_ok": False,
                        "conf_params": {"sistema": str(sis_hint), "mes_num": int(mes_hint), "ano": int(ano_hint)},
                        "show_conf_panel": False,  # começa com botão; painel abre ao clicar
                    })
                    # <<< fim agendamento
                    
                    # limpa estado e encerra
                    st.session_state.modo_conflitos = False
                    st.session_state.conflitos_df_conf = None
                    st.session_state.conflitos_spreadsheet_id = None
                    st.session_state.conflitos_sheet_id = None
                    st.stop()

    
                except Exception as e:
                    st.error(f"❌ Erro ao aplicar exclusões/inclusões: {e}")
                    st.stop()

        
        # ------------------------ EDITOR MANUAL (somente processa no final) ------------------------
        # ------------------------ EDITOR MANUAL (1 linha por vez) ------------------------
        if "show_manual_editor" not in st.session_state:
            st.session_state.show_manual_editor = False
        if "manual_df" not in st.session_state or st.session_state.manual_df.shape[0] != 5:
            st.session_state.manual_df = template_manuais(5)
        
        if st.session_state.get("show_manual_editor", False):
            #st.subheader("Lançamentos manuais (1 linha por vez)")
        
            # Catálogo de lojas para preencher códigos automaticamente
            gc_ = get_gc()
            catalogo = carregar_catalogo_codigos(gc_, nome_planilha="Vendas diarias", aba_catalogo="Tabela Empresa")
            lojas_options = sorted(
                catalogo["Loja"].dropna().astype(str).str.strip().unique().tolist()
            ) if not catalogo.empty else []
        
            PLACEHOLDER_LOJA = "— selecione a loja —"
            lojas_options_ui = [PLACEHOLDER_LOJA] + lojas_options
        
            # Base exibida
            df_disp = st.session_state.manual_df.copy()
            df_disp["Loja"] = df_disp["Loja"].fillna("").astype(str).str.strip()
            for c in ["Fat.Total","Serv/Tx","Fat.Real","Ticket"]:
                df_disp[c] = pd.to_numeric(df_disp[c], errors="coerce")
        
            df_disp = df_disp[["Data","Loja","Fat.Total","Serv/Tx","Fat.Real","Ticket"]]
        
            with st.form("form_lancamentos_manuais_5"):
                edited_df = st.data_editor(
                    df_disp,
                    num_rows="fixed",  # trava em 1 linha
                    use_container_width=True,
                    hide_index=True,
                    column_config={
                        "Data": st.column_config.DateColumn(format="DD/MM/YYYY"),  # calendário, sem valor sugerido (NaT)
                        "Loja": st.column_config.SelectboxColumn(
                            options=lojas_options_ui,
                            default=PLACEHOLDER_LOJA,
                            help="Escolha a loja (digite para filtrar)"
                        ),
                        "Fat.Total": st.column_config.NumberColumn(step=0.01),
                        "Serv/Tx":   st.column_config.NumberColumn(step=0.01),
                        "Fat.Real":  st.column_config.NumberColumn(step=0.01),
                        "Ticket":    st.column_config.NumberColumn(step=0.01),
                    },
                    key="editor_manual_5",
                )
                c_esq, c_dir = st.columns([1,1])
                salvar = c_esq.form_submit_button("💾 Salvar linha", use_container_width=True)
                limpar = c_dir.form_submit_button("🧹 Limpar", use_container_width=True)
        

            if salvar:
                edited_df = edited_df.copy()
        
                # Normaliza placeholder de loja
                edited_df["Loja"] = edited_df["Loja"].replace({PLACEHOLDER_LOJA: ""}).astype(str).str.strip()
        
                # Mantém apenas linhas com Data e Loja preenchidos
                mask_validas = (~edited_df["Data"].isna()) & (edited_df["Loja"] != "")
                df_validas = edited_df.loc[mask_validas].copy()
        
                if df_validas.empty:
                    st.error("⚠️ Preencha **Data** e **Loja** em pelo menos uma linha.")
                    st.stop()
        
                # Prepara (parse de data, numéricos, derivados) e completa códigos
                df_pronto = preparar_manuais_para_envio(df_validas, catalogo)
        
                if df_pronto.empty:
                    st.warning("Nada para enviar.")
                else:
                    ok = enviar_para_sheets(df_pronto, titulo_origem="manuais")
                    if ok:
                        # reset para próxima edição com 5 linhas
                        st.session_state.manual_df = template_manuais(5)
                        st.rerun()
        
            if limpar:
                st.session_state.manual_df = template_manuais(5)
                st.rerun()

        # ---------- ENVIO AUTOMÁTICO (botão principal) ----------
        #if st.button("Atualizar SheetsS (usar df do Upload)", use_container_width=True, disabled=('df_final' not in st.session_state or st.session_state.df_final.empty), key="btn_enviar_auto_footer"):
        #    if 'df_final' not in st.session_state or st.session_state.df_final.empty:
        #        st.error("Não há dados para enviar.")
        #    else:
        #        # Reaproveita a função enviar_para_sheets com o DF do upload (aba 1)
        #        ok = enviar_para_sheets(st.session_state.df_final.copy(), titulo_origem="upload")
        #        if ok:
        #            st.success("✅ Processo concluído.")

        
    
    
    
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

    # =======================================
    # =======================================
    # Aba 5 - Auditoria PDV x Faturamento Meio Pagamento (tabela única)
    # =======================================
    with aba5:
        import re
        import json
        import pandas as pd
        import numpy as np
        import unicodedata
    
        #st.subheader("📊 Auditoria Mensal — Sistema × Meio de Pagamento (Mês/Ano das abas)")
    
        # ---------------- Config ----------------
        ANO_MIN = 2025       # considera 2025 em diante (>= 2025)
        TOL = 0.01           # tolerância de 1 centavo
    
        # ---------------- Helpers ----------------
        def _ns(s: str) -> str:
            s = str(s or "").strip().lower()
            s = unicodedata.normalize("NFD", s)
            s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
            return re.sub(r"[^a-z0-9]+", " ", s).strip()
    
        def _norm_key(s: str) -> str:
            return _ns(s).replace(" ", "")
    
        def pick_exact_column(cols, targets):
            norm_map = {_norm_key(c): c for c in cols}
            for t in targets:
                t_norm = _norm_key(t)
                if t_norm in norm_map:
                    return norm_map[t_norm]
            return None
    
        _MAP_MES = {
            "jan":1,"fev":2,"mar":3,"abr":4,"mai":5,"jun":6,"jul":7,"ago":8,"set":9,"out":10,"nov":11,"dez":12,
            "janeiro":1,"fevereiro":2,"março":3,"marco":3,"abril":4,"maio":5,"junho":6,"julho":7,"agosto":8,
            "setembro":9,"outubro":10,"novembro":11,"dezembro":12,
            "jan.":1,"feb":2,"feb.":2,"mar.":3,"apr":4,"apr.":4,"may":5,"jun.":6,"jul.":7,"aug":8,"aug.":8,
            "sep":9,"sept":9,"sep.":9,"oct":10,"oct.":10,"nov.":11,"dec":12,"dec.":12
        }
        def _month_to_num(x):
            s = str(x or "").strip().lower()
            if s in _MAP_MES: return _MAP_MES[s]
            try:
                n = int(float(s))
                if 1 <= n <= 12: return n
            except: pass
            return None
    
        def _mk_messtr(m, a):
            try:
                m = int(m); a = int(a)
                if 1 <= m <= 12:
                    return f"{m:02d}/{a}"
            except:
                pass
            return ""
    
        def _to_float_brl_ext(x):
            s = str(x or "").strip()
            s = re.sub(r"[^\d,.\-]", "", s)
            if s == "": return float("nan")
            if s.count(",") == 1 and s.count(".") >= 1:
                s = s.replace(".", "").replace(",", ".")
            elif s.count(",") == 1 and s.count(".") == 0:
                s = s.replace(",", ".")
            try:
                return float(s)
            except:
                return float("nan")
    
        def _parse_date_series(ser):
            """
            Converte uma Series que pode conter datas como texto (dd/mm/aaaa) OU
            números seriais do Excel (dias desde 1899-12-30) para datetime64[ns],
            de forma resiliente (evita FloatingPointError).
            """
            import numpy as np
            import pandas as pd
        
            # Garante Series
            if not isinstance(ser, pd.Series):
                ser = pd.Series(ser)
        
            # 1) Tenta parsear como texto (dd/mm/aaaa, etc.)
            dt_text = pd.to_datetime(ser, dayfirst=True, errors="coerce")
        
            # 2) Tenta como serial do Excel, mas SOMENTE para números plausíveis
            num = pd.to_numeric(ser, errors="coerce")
        
            # máscara: numérico, finito e dentro de uma faixa plausível de dias do Excel
            # (1 até 100000 cobre ~até ano 2173; ajuste se quiser mais curto)
            mask = num.notna() & np.isfinite(num) & (num >= 1) & (num <= 100000)
        
            # Evita passar floats para o parser com unit="D" (que aciona np.round internamente).
            # Forçamos inteiro por truncamento (ou use np.floor se preferir).
            num_i = pd.Series(index=ser.index, dtype="Int64")
            if mask.any():
                num_i.loc[mask] = np.floor(num.loc[mask].astype(float)).astype("Int64")
        
            dt_excel = pd.Series(pd.NaT, index=ser.index, dtype="datetime64[ns]")
            if mask.any():
                dt_excel.loc[mask] = pd.to_datetime(
                    num_i.loc[mask], origin="1899-12-30", unit="D", errors="coerce"
                )
        
            # 3) Prioriza texto; onde NaT, usa Excel; resto fica NaT
            return dt_text.where(dt_text.notna(), dt_excel)

    
        # === parser do Meio de Pagamento (mesmo da Aba Vendas)
        def _valor_meio_pagamento(ser: pd.Series) -> pd.Series:
            ser = ser.astype(str)
            ser = ser.str.replace("R$", "", regex=False)
            ser = ser.str.replace("(", "-", regex=False).str.replace(")", "", regex=False)
            ser = ser.str.replace(" ", "", regex=False)
            ser = ser.str.replace(".", "", regex=False)
            ser = ser.str.replace(",", ".", regex=False)
            return pd.to_numeric(ser, errors="coerce")
    
        # ✅ BRL: R$ 246.931,52
        def _fmt_brl(v):
            try:
                v = float(v)
            except:
                return "R$ 0,00"
            s = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            return f"R$ {s}"
    
        # ---------------- Conexão (sem vazar gc) ----------------
        if "get_gc" in globals() and callable(get_gc):
            _gc = get_gc()
        else:
            import gspread
            from oauth2client.service_account import ServiceAccountCredentials
            scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
            credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
            credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
            _gc = gspread.authorize(credentials)
    
        sh = _gc.open("Vendas diarias")
        del _gc  # remove o cliente para não “aparecer”
    
        # ---------------- Leitura ----------------
        ws_ext = sh.worksheet("Fat Sistema Externo")
        vals_ext = ws_ext.get_all_values()
        if len(vals_ext) < 2:
            st.error("Aba 'Fat Sistema Externo' vazia.")
            st.stop()
        header_ext, rows_ext = vals_ext[0], vals_ext[1:]
        df_ext = pd.DataFrame(rows_ext, columns=[c.strip() for c in header_ext])
    
        ws_mp = sh.worksheet("Faturamento Meio Pagamento")
        vals_mp = ws_mp.get_all_values()
        if len(vals_mp) < 2:
            st.error("Aba 'Faturamento Meio Pagamento' vazia.")
            st.stop()
        header_mp, rows_mp = vals_mp[0], vals_mp[1:]
        df_mp = pd.DataFrame(rows_mp, columns=[c.strip() for c in header_mp])
    
        # ---------------- Colunas ----------------
        col_ext_data = next((c for c in df_ext.columns if _norm_key(c) == _norm_key("Data")), None)
        col_ext_cod  = next((c for c in df_ext.columns if "codigo" in _ns(c) and "everest" in _ns(c)), None)
        col_ext_sis  = next((c for c in df_ext.columns if _norm_key(c) == _norm_key("Sistema")), None)
        col_ext_mes  = next((c for c in df_ext.columns if _norm_key(c) in (_norm_key("Mês"), _norm_key("Mes"))), None)
        col_ext_ano  = next((c for c in df_ext.columns if _norm_key(c) == _norm_key("Ano")), None)
        col_ext_fat  = pick_exact_column(df_ext.columns, ["Fat. Total"])
    
        if not all([col_ext_data, col_ext_cod, col_ext_fat, col_ext_sis]):
            st.error("Em 'Fat Sistema Externo' preciso de: Data, Código Everest, **Fat. Total**, Sistema (ideal Mês/Ano).")
            st.stop()
    
        col_mp_data = next((c for c in df_mp.columns if _norm_key(c) == _norm_key("Data")), None)
        col_mp_cod  = next((c for c in df_mp.columns if "codigo" in _ns(c) and "everest" in _ns(c)), None)
        col_mp_sis  = next((c for c in df_mp.columns if _norm_key(c) == _norm_key("Sistema")), None)
        col_mp_mes  = next((c for c in df_mp.columns if _norm_key(c) in (_norm_key("Mês"), _norm_key("Mes"))), None)
        col_mp_ano  = next((c for c in df_mp.columns if _norm_key(c) == _norm_key("Ano")), None)
        col_mp_val  = pick_exact_column(df_mp.columns, ["Valor (R$)", "Valor R$"])
    
        if not all([col_mp_data, col_mp_cod, col_mp_val, col_mp_sis]):
            st.error("Em 'Faturamento Meio Pagamento' preciso de: Data, Código Everest, **Valor (R$)**, Sistema (ideal Mês/Ano).")
            st.stop()
    
        # ---------------- Preparar EXTERNO ----------------
        ext = pd.DataFrame({
            "Data":           _parse_date_series(df_ext[col_ext_data]),
            "Código Everest": pd.to_numeric(df_ext[col_ext_cod], errors="coerce"),
            "Sistema":        df_ext[col_ext_sis].astype(str).str.strip().str.upper(),
            "Fat.Total":      df_ext[col_ext_fat].map(_to_float_brl_ext),
            "MesNum":         df_ext[col_ext_mes] if col_ext_mes else None,
            "Ano":            df_ext[col_ext_ano] if col_ext_ano else None,
        })
        ext["MesNum"] = ext["MesNum"].apply(lambda x: _month_to_num(x) if pd.notna(x) else None)
        ext["Ano"]    = pd.to_numeric(ext["Ano"], errors="coerce")
        ext = ext[ext["Fat.Total"].notna()].copy()
        if ext["MesNum"].isna().any():
            ext["MesNum"] = ext["MesNum"].fillna(pd.to_datetime(ext["Data"], errors="coerce").dt.month)
        if ext["Ano"].isna().any():
            ext["Ano"] = ext["Ano"].fillna(pd.to_datetime(ext["Data"], errors="coerce").dt.year)
        # 2025+
        ext = ext[ext["Ano"] >= ANO_MIN].copy()
        ext["Mês"] = ext.apply(lambda r: _mk_messtr(r["MesNum"], r["Ano"]), axis=1)
    
        # ---------------- Preparar MEIO DE PAGAMENTO ----------------
        mp = pd.DataFrame({
            "Data":           _parse_date_series(df_mp[col_mp_data]),
            "Código Everest": pd.to_numeric(df_mp[col_mp_cod], errors="coerce"),
            "Sistema":        df_mp[col_mp_sis].astype(str).str.strip().str.upper(),
            "Valor_MP":       _valor_meio_pagamento(df_mp[col_mp_val]),
            "MesNum":         df_mp[col_mp_mes] if col_mp_mes else None,
            "Ano":            df_mp[col_mp_ano] if col_mp_ano else None,
        })
        mp["MesNum"] = mp["MesNum"].apply(lambda x: _month_to_num(x) if pd.notna(x) else None)
        mp["Ano"]    = pd.to_numeric(mp["Ano"], errors="coerce")
        mp = mp[mp["Valor_MP"].notna()].copy()
        if mp["MesNum"].isna().any():
            mp["MesNum"] = mp["MesNum"].fillna(pd.to_datetime(mp["Data"], errors="coerce").dt.month)
        if mp["Ano"].isna().any():
            mp["Ano"] = mp["Ano"].fillna(pd.to_datetime(mp["Data"], errors="coerce").dt.year)
        # 2025+
        mp = mp[mp["Ano"] >= ANO_MIN].copy()
        mp["Mês"] = mp.apply(lambda r: _mk_messtr(r["MesNum"], r["Ano"]), axis=1)
    
        # ---------------- RESUMO (Mês + Sistema) ----------------
        ext_mes = (ext.groupby(["Mês","Sistema"], as_index=False)["Fat.Total"].sum()
                      .rename(columns={"Fat.Total":"Total_Faturamento"}))
        mp_mes  = (mp .groupby(["Mês","Sistema"], as_index=False)["Valor_MP"].sum()
                      .rename(columns={"Valor_MP":"Total_MeioPagamento"}))
    
        resumo = ext_mes.merge(mp_mes, on=["Mês","Sistema"], how="outer")
        resumo[["Total_Faturamento","Total_MeioPagamento"]] = resumo[["Total_Faturamento","Total_MeioPagamento"]].fillna(0.0).round(2)
        resumo["Diferença"] = (resumo["Total_Faturamento"] - resumo["Total_MeioPagamento"]).round(2)
    
        # 🔢 Ordena por ano/mês + sistema
        def _ordem_mes(mes_str):
            try:
                mm, aa = mes_str.split("/")
                return int(aa)*100 + int(mm)
            except:
                return 0
        resumo["__ordem__"] = resumo["Mês"].apply(_ordem_mes)
        resumo = resumo.sort_values(["__ordem__","Sistema"]).reset_index(drop=True)
    
        # ----------- TABELA ÚNICA (com checkbox e BRL em texto) -----------
        tabela = resumo[["Mês","Sistema","Total_Faturamento","Total_MeioPagamento","Diferença"]].copy()
    
        # Aparência do Sistema (title) apenas visual; manteremos UPPER nas comparações
        tabela["Sistema_View"] = tabela["Sistema"].str.title()
        # Formata BRL em texto
        for col in ["Total_Faturamento","Total_MeioPagamento","Diferença"]:
            tabela[col] = tabela[col].apply(_fmt_brl)
    
        # Coluna de seleção
        tabela.insert(0, "Selecionar", False)
    
        # Ordena e organiza colunas para exibição
        tabela_view = tabela[["Selecionar","Mês","Sistema_View","Total_Faturamento","Total_MeioPagamento","Diferença"]].rename(
            columns={"Sistema_View":"Sistema"}
        )
    
        edited = st.data_editor(
            tabela_view,
            hide_index=True,
            use_container_width=True,
            column_config={
                "Selecionar": st.column_config.CheckboxColumn(help="Marque o(s) Mês/Sistema para ver o detalhe por loja"),
                "Total_Faturamento":    st.column_config.TextColumn(label="Total Faturamento (R$)"),
                "Total_MeioPagamento":  st.column_config.TextColumn(label="Total Meio Pagamento (R$)"),
                "Diferença":            st.column_config.TextColumn(label="Diferença (R$)"),
            },
        )
    
        # Pares selecionados (convertendo Sistema de volta para UPPER p/ casar com dados)
        selected_pairs = set(
            zip(edited.loc[edited["Selecionar"] == True, "Mês"],
                edited.loc[edited["Selecionar"] == True, "Sistema"].astype(str).str.upper())
        )
    
        # ---------------- Detalhe por loja (apenas dos selecionados) ----------------
        #st.markdown("**Lojas com diferença (por Data), somente dos Mês/Sistema selecionados**")
        if not selected_pairs:
            st.info("Marque ao menos um Mês/Sistema na tabela para ver o detalhe.")
        else:
            try:
                df_emp_map = df_empresa.copy()
            except NameError:
                df_emp_map = pd.DataFrame(sh.worksheet("Tabela Empresa").get_all_records())
            df_emp_map.columns = df_emp_map.columns.str.strip()
            col_cod_emp = next((c for c in df_emp_map.columns if "everest" in _ns(c) and "grupo" not in _ns(c)), None)
            col_loja    = next((c for c in df_emp_map.columns if _ns(c) == "loja"), None)
            mapa_loja = {}
            if col_cod_emp and col_loja:
                tmp = df_emp_map.dropna(subset=[col_cod_emp]).copy()
                tmp[col_cod_emp] = pd.to_numeric(tmp[col_cod_emp], errors="coerce")
                tmp = tmp.dropna(subset=[col_cod_emp])
                mapa_loja = tmp.set_index(col_cod_emp)[col_loja].to_dict()
    
            def _mask_sel(df):
                mask = False
                for m, s in selected_pairs:
                    mask = mask | ((df["Mês"] == m) & (df["Sistema"] == s))
                return df[mask]
    
            ext_det = (ext.pipe(_mask_sel)
                         .groupby(["Data","Mês","Sistema","Código Everest"], as_index=False)["Fat.Total"].sum()
                         .rename(columns={"Fat.Total":"Ext"}))
            mp_det  = (mp.pipe(_mask_sel)
                         .groupby(["Data","Mês","Sistema","Código Everest"], as_index=False)["Valor_MP"].sum()
                         .rename(columns={"Valor_MP":"MP"}))
    
            det = (ext_det.merge(mp_det, on=["Data","Mês","Sistema","Código Everest"], how="outer").fillna(0.0))
            det["Diferença"] = (det["Ext"] - det["MP"]).round(2)
            det = det.loc[det["Diferença"].abs() > TOL].copy()
    
            if det.empty:
                st.success("✅ Nenhuma diferença por loja para os pares selecionados.")
            else:
                det["Loja"] = det["Código Everest"].map(mapa_loja).fillna("")
                det = det.sort_values(["Mês","Sistema","Data","Loja","Código Everest"]).reset_index(drop=True)
                st.dataframe(
                    det[["Mês","Sistema","Data","Loja","Código Everest","Ext","MP","Diferença"]]
                      .style.format({"Ext": _fmt_brl, "MP": _fmt_brl, "Diferença": _fmt_brl})
                      .format_index(na_rep=""),
                    use_container_width=True, hide_index=True
                )
