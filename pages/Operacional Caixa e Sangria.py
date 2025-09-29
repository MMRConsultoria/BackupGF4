import streamlit as st
import pandas as pd
import numpy as np
import json
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from gspread_formatting import format_cell_range, CellFormat, NumberFormat




st.set_page_config(page_title="Relatório de Sangria", layout="wide")
# 🔥 CSS para estilizar as abas
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

# 🔒 Bloqueio
if not st.session_state.get("acesso_liberado"):
    st.stop()

# 🔕 Oculta toolbar
st.markdown("""
    <style>
        [data-testid="stToolbar"] { visibility: hidden; height: 0%; position: fixed; }
        .stSpinner { visibility: visible !important; }
    </style>
""", unsafe_allow_html=True)

NOME_SISTEMA = "Colibri"

# -----------------------
# Helpers
# -----------------------
def auto_read_first_or_sheet(uploaded, preferred="Sheet"):
    """Lê a guia 'preferred' se existir; senão, lê a primeira guia."""
    xls = pd.ExcelFile(uploaded)
    sheets = xls.sheet_names
    sheet_to_read = preferred if preferred in sheets else sheets[0]
    df0 = pd.read_excel(xls, sheet_name=sheet_to_read)
    return df0, sheet_to_read, sheets

def normalize_dates(s):
    """Para comparar datas (remove horário)."""
    return pd.to_datetime(s, errors="coerce", dayfirst=True).dt.normalize()

with st.spinner("⏳ Processando..."):
    # 🔌 Conexão Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gc = gspread.authorize(credentials)
    planilha = gc.open("Vendas diarias")

    df_empresa = pd.DataFrame(planilha.worksheet("Tabela Empresa").get_all_records())
    ws_tab = planilha.worksheet("Tabela Sangria")
    dados = ws_tab.get_all_records()  # lê usando a primeira linha como cabeçalho
    df_descricoes = pd.DataFrame(dados)
    # Garante as colunas
    df_descricoes.columns = [c.strip() for c in df_descricoes.columns]
    if not {"Palavra-chave", "Descrição Agrupada"}.issubset(df_descricoes.columns):
        st.error("A aba 'Tabela Sangria' precisa ter as colunas 'Palavra-chave' e 'Descrição Agrupada'.")
        st.stop()

    # 🔥 Título
    st.markdown("""
        <div style='display: flex; align-items: center; gap: 10px;'>
            <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
            <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Controle de Caixa e Sangria</h1>
        </div>
    """, unsafe_allow_html=True)

    # 🗂️ Abas
    tab1, tab2 = st.tabs(["📥 Upload e Processamento", "🔄 Atualizar Google Sheets"])

    # ================
    # 📥 Aba 1 — Upload e Processamento (detecção Colibri × Everest)
    # ================
    with tab1:
        uploaded_file = st.file_uploader(
            label="📁 Clique para selecionar ou arraste aqui o arquivo Excel",
            type=["xlsx", "xlsm"],
            help="Somente arquivos .xlsx ou .xlsm. Tamanho máximo: 200MB."
        )
    
        if uploaded_file:
            def auto_read_first_or_sheet(uploaded, preferred="Sheet"):
                xls = pd.ExcelFile(uploaded)
                sheets = xls.sheet_names
                sheet_to_read = preferred if preferred in sheets else sheets[0]
                df0 = pd.read_excel(xls, sheet_name=sheet_to_read)
                df0.columns = [str(c).strip() for c in df0.columns]
                return df0, sheet_to_read, sheets
    
            try:
                df_dados, guia_lida, lista_guias = auto_read_first_or_sheet(uploaded_file, preferred="Sheet")
                #st.caption(f"Guia lida: **{guia_lida}** (disponíveis: {', '.join(lista_guias)})")
            except Exception as e:
                st.error(f"❌ Não foi possível ler o arquivo enviado. Detalhes: {e}")
            else:
                df = df_dados.copy()
                primeira_col = df.columns[0] if len(df.columns) else ""
                is_everest = primeira_col.lower() in ["lançamento", "lancamento"] or ("Lançamento" in df.columns) or ("Lancamento" in df.columns)
    
                if is_everest:

                    # ---------------- MODO EVEREST ----------------
                    st.session_state.mode = "everest"
                    st.session_state.df_everest = df.copy()
                
                    import unicodedata, re
                    def _norm(s: str) -> str:
                        s = unicodedata.normalize('NFKD', str(s)).encode('ASCII','ignore').decode('ASCII')
                        s = s.lower()
                        s = re.sub(r'[^a-z0-9]+', ' ', s)
                        return re.sub(r'\s+', ' ', s).strip()
                
                    # 1) DATA => "D. Lançamento" (variações)
                    date_col = None
                    for cand in ["D. Lançamento", "D.Lançamento", "D. Lancamento", "D.Lancamento"]:
                        if cand in df.columns:
                            date_col = cand
                            break
                    if date_col is None:
                        for col in df.columns:
                            if _norm(col) in ["d lancamento", "data lancamento", "d lancamento data"]:
                                date_col = col
                                break
                    st.session_state.everest_date_col = date_col
                
                    # 2) VALOR => "Valor Lançamento" (variações) com fallback seguro
                    def detect_valor_col(_df, avoid_col=None):
                        aliases = [
                            "valor lancamento", "valor lançamento",
                            "valor do lancamento", "valor de lancamento",
                            "valor do lançamento", "valor de lançamento",
                            "valor"
                        ]
                        # preferir match por nome normalizado (exato)
                        targets = {a: _norm(a) for a in aliases}
                        for c in _df.columns:
                            if c == avoid_col: 
                                continue
                            if _norm(c) in targets.values():
                                return c
                        # fallback: escolher coluna (≠ data) com mais células contendo dígitos
                        best, score = None, -1
                        for c in _df.columns:
                            if c == avoid_col: 
                                continue
                            sc = _df[c].astype(str).str.contains(r"\d").sum()
                            if sc > score:
                                best, score = c, sc
                        return best
                
                    valor_col = detect_valor_col(df, avoid_col=date_col)
                    st.session_state.everest_value_col = valor_col
                    # Conversor pt-BR robusto: R$, parênteses, sinal no final (1.234,56-)
                    def to_number_br(series):
                        def _one(x):
                            if pd.isna(x):
                                return 0.0
                            s = str(x).strip()
                            if s == "":
                                return 0.0
                            neg = False
                            # parênteses => negativo
                            if s.startswith("(") and s.endswith(")"):
                                neg = True
                                s = s[1:-1].strip()
                            # remove R$
                            s = s.replace("R$", "").replace("r$", "").strip()
                            # sinal no final (ex.: 1.234,56-)
                            if s.endswith("-"):
                                neg = True
                                s = s[:-1].strip()
                            # separadores pt-BR
                            s = s.replace(".", "").replace(",", ".")
                            s_clean = re.sub(r"[^0-9.\-]", "", s)
                            if s_clean in ["", "-", "."]:
                                return 0.0
                            try:
                                val = float(s_clean)
                            except:
                                s_fallback = re.sub(r"[^0-9.]", "", s_clean)
                                val = float(s_fallback) if s_fallback else 0.0
                            return -abs(val) if neg else val
                        return series.apply(_one)
                
                    # 3) Métricas
                    periodo_txt = "—"
                    total_txt = "—"
                
                    # Período a partir de D. Lançamento
                    if date_col is not None:
                        dt = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True)
                        valid = dt.dropna()
                        if not valid.empty:
                            periodo_min = valid.min().strftime("%d/%m/%Y")
                            periodo_max = valid.max().strftime("%d/%m/%Y")
                            periodo_txt = f"{periodo_min} até {periodo_max}"
                            st.session_state.everest_dates = valid.dt.normalize().unique().tolist()
                        else:
                            st.warning("⚠️ A coluna 'D. Lançamento' existe, mas não tem datas válidas.")
                    else:
                        st.error("❌ Não encontrei a coluna **'D. Lançamento'**.")
                
                    # Total pela coluna de valor (preservando o sinal real)
                    if valor_col is not None:
                        if pd.api.types.is_numeric_dtype(df[valor_col]):
                            serie_val = pd.to_numeric(df[valor_col], errors="coerce").fillna(0.0)
                        else:
                            serie_val = to_number_br(df[valor_col])
                        total_liquido = float(serie_val.sum())
                        st.session_state.everest_total_liquido = total_liquido
                
                        sinal = "-" if total_liquido < 0 else ""
                        total_fmt = f"{abs(total_liquido):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        total_txt = f"{sinal}R$ {total_fmt}"
                    else:
                        st.warning("⚠️ Não encontrei a coluna de **valor** (ex.: 'Valor Lançamento').")
                
                    # 4) Métricas (sem preview)
                    if periodo_txt != "—":
                        c1, c2, c3 = st.columns(3)
                        c1.metric("📅 Período processado", periodo_txt)
                        #c2.metric("🧾 Linhas lidas", f"{len(df)}")
                        c3.metric("💰 Total (Valor Lançamento)", total_txt)
                    else:
                        c1, c2 = st.columns(2)
                        #c1.metric("🧾 Linhas lidas", f"{len(df)}")
                        c2.metric("💰 Total (Valor Lançamento)", total_txt)
                
                    # 5) Download do arquivo como veio
                    output_ev = BytesIO()
                    with pd.ExcelWriter(output_ev, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="Sangria Everest")
                    output_ev.seek(0)
                    st.download_button(
                        "📥 Sangria Everest",
                        data=output_ev,
                        file_name="Sangria_Everest.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

    
                else:
                    # ---------------- MODO COLIBRI (seu fluxo atual) ----------------
                    try:
                        df["Loja"] = np.nan
                        df["Data"] = np.nan
                        df["Funcionário"] = np.nan
    
                        data_atual = None
                        funcionario_atual = None
                        loja_atual = None
                        linhas_validas = []
    
                        for i, row in df.iterrows():
                            valor = str(row["Hora"]).strip()
                            if valor.startswith("Loja:"):
                                loja = valor.split("Loja:")[1].split("(Total")[0].strip()
                                if "-" in loja:
                                    loja = loja.split("-", 1)[1].strip()
                                loja_atual = loja or "Loja não cadastrada"
                            elif valor.startswith("Data:"):
                                try:
                                    data_atual = pd.to_datetime(
                                        valor.split("Data:")[1].split("(Total")[0].strip(), dayfirst=True
                                    )
                                except Exception:
                                    data_atual = pd.NaT
                            elif valor.startswith("Funcionário:"):
                                funcionario_atual = valor.split("Funcionário:")[1].split("(Total")[0].strip()
                            else:
                                if pd.notna(row["Valor(R$)"]) and pd.notna(row["Hora"]):
                                    df.at[i, "Data"] = data_atual
                                    df.at[i, "Funcionário"] = funcionario_atual
                                    df.at[i, "Loja"] = loja_atual
                                    linhas_validas.append(i)
    
                        df = df.loc[linhas_validas].copy()
                        df.ffill(inplace=True)
                        # ===== FIX: preencher Descrição quando Hora é horário e Valor(R$) está preenchido =====
                        def _is_blank(series: pd.Series) -> pd.Series:
                            s = series.astype(str).str.strip().str.lower()
                            return series.isna() | s.isin(["", "nan", "none", "null"])
                        
                        # Hora válida: tenta converter para horário (ou timestamp) → não NaT
                        mask_hora_valida = pd.to_datetime(df["Hora"], errors="coerce").notna()
                        
                        # Valor(R$) preenchido (qualquer coisa diferente de vazio/NAN/NONE)
                        mask_valor_preenchido = ~_is_blank(df["Valor(R$)"])
                        
                        # Descrição vazia
                        mask_desc_vazia = _is_blank(df["Descrição"])
                        
                        # Aplica a regra
                        df.loc[mask_hora_valida & mask_valor_preenchido & mask_desc_vazia, "Descrição"] = "sem preenchimento"
                        # =======================================================================

                        # Limpeza e conversões
                        df["Descrição"] = (
                            df["Descrição"].astype(str).str.strip().str.lower().str.replace(r"\s+", " ", regex=True)
                        )
                        df["Funcionário"] = df["Funcionário"].astype(str).str.strip()
                        df["Valor(R$)"] = pd.to_numeric(df["Valor(R$)"], errors="coerce").fillna(0.0).round(2)
    
                        # Dia semana / mês / ano
                        dias_semana = {0: 'segunda-feira', 1: 'terça-feira', 2: 'quarta-feira',
                                       3: 'quinta-feira', 4: 'sexta-feira', 5: 'sábado', 6: 'domingo'}
                        df["Dia da Semana"] = df["Data"].dt.dayofweek.map(dias_semana)
                        df["Mês"] = df["Data"].dt.month.map({
                            1: 'jan', 2: 'fev', 3: 'mar', 4: 'abr', 5: 'mai', 6: 'jun',
                            7: 'jul', 8: 'ago', 9: 'set', 10: 'out', 11: 'nov', 12: 'dez'
                        })
                        df["Ano"] = df["Data"].dt.year
                        df["Data"] = df["Data"].dt.strftime("%d/%m/%Y")
    
                        # Merge com cadastro de lojas
                        df["Loja"] = df["Loja"].astype(str).str.strip().str.lower()
                        df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower()
                        df = pd.merge(df, df_empresa, on="Loja", how="left")
    
                        # Agrupamento de descrição
                        def mapear_descricao(desc):
                            desc_lower = str(desc).lower()
                            for _, r in df_descricoes.iterrows():
                                if str(r["Palavra-chave"]).lower() in desc_lower:
                                    return r["Descrição Agrupada"]
                            return "Outros"
    
                        df["Descrição Agrupada"] = df["Descrição"].apply(mapear_descricao)
    
                        # ➕ Colunas adicionais
                        df["Sistema"] = NOME_SISTEMA
    
                        # 🔑 DUPLICIDADE
                        data_key = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d")
                        hora_key = pd.to_datetime(df["Hora"], errors="coerce").dt.strftime("%H:%M:%S")
                        valor_centavos = (df["Valor(R$)"].astype(float) * 100).round().astype(int).astype(str)
                        desc_key = df["Descrição"].fillna("").astype(str)
                        df["Duplicidade"] = (
                            data_key.fillna("") + "|" +
                            hora_key.fillna("") + "|" +
                            df["Código Everest"].fillna("").astype(str) + "|" +
                            valor_centavos + "|" +
                            desc_key
                        )
    
                        if "Meio de recebimento" not in df.columns:
                            df["Meio de recebimento"] = ""
    
                        colunas_ordenadas = [
                            "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
                            "Código Grupo Everest", "Funcionário", "Hora", "Descrição",
                            "Descrição Agrupada", "Meio de recebimento", "Valor(R$)",
                            "Mês", "Ano", "Duplicidade", "Sistema"
                        ]
                        df = df[colunas_ordenadas].sort_values(by=["Data", "Loja"])
    
                        # Métricas
                        periodo_min = pd.to_datetime(df["Data"], dayfirst=True).min().strftime("%d/%m/%Y")
                        periodo_max = pd.to_datetime(df["Data"], dayfirst=True).max().strftime("%d/%m/%Y")
                        valor_total = float(df["Valor(R$)"].sum())
    
                        col1, col2 = st.columns(2)
                        col1.metric("📅 Período processado", f"{periodo_min} até {periodo_max}")
                        col2.metric(
                            "💰 Valor total de sangria",
                            f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        )
    
                        st.success("✅ Relatório gerado com sucesso!")
    
                        lojas_sem_codigo = df[df["Código Everest"].isna()]["Loja"].unique()
                        if len(lojas_sem_codigo) > 0:
                            st.warning(
                                f"⚠️ Lojas sem Código Everest cadastrado: {', '.join(lojas_sem_codigo)}\n\n"
                                "🔗 Atualize na planilha de empresas."
                            )
    
                        # Guarda para a Tab2 (fluxo antigo)
                        st.session_state.mode = "colibri"
                        st.session_state.df_sangria = df.copy()
    
                        # Download Excel local (sem formatação especial)
                        # --- ORDENAR por Data antes do download local (Colibri) ---
                        df_download = df.copy()
                        
                        # A coluna "Data" no fluxo Colibri já está como string dd/mm/aaaa.
                        # Vamos criar uma coluna auxiliar datetime para ordenar corretamente.
                        df_download["_Data_dt"] = pd.to_datetime(df_download["Data"], dayfirst=True, errors="coerce")
                        
                        # Se quiser desempatar por loja dentro do dia, use: by=["_Data_dt","Loja"]
                        df_download = df_download.sort_values(by=["_Data_dt"]).drop(columns=["_Data_dt"])
                        
                        # Gera o Excel já ordenado
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            df_download.to_excel(writer, index=False, sheet_name="Sangria")
                        output.seek(0)
                        
                        st.download_button("📥Sangria Colibri",
                                           data=output,
                                           file_name="Sangria_estruturada.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

                    except KeyError as e:
                        st.error(f"❌ Coluna obrigatória ausente para o padrão Colibri: {e}")


    # ================
    # 🔄 Aba 2 — Atualizar Google Sheets (Sangria × Sangria Everest)
    # ================
    with tab2:
        st.markdown("🔗 [Abrir planilha Vendas diarias](https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU)")
        from gspread_formatting import format_cell_range, CellFormat, NumberFormat

        # Formato contábil com símbolo R$ (positivo; negativo entre parênteses; zero como "-"; texto)
        ACCOUNTING_RS = CellFormat(
            numberFormat=NumberFormat(
                type="CURRENCY",
                pattern="R$ * #,##0.00_);R$ * (#,##0.00);R$ * -_;@"
            )
        )

        # ✅ defina o mode ANTES de usá-lo
        mode = st.session_state.get("mode", None)
        def _excel_col_letter(idx_zero_based: int) -> str:
            """Converte índice 0-based em letra de coluna (A..Z, AA..)."""
            n = idx_zero_based + 1
            s = ""
            while n:
                n, r = divmod(n - 1, 26)
                s = chr(65 + r) + s
            return s

        # --- modo Everest: substituir apenas as datas presentes no arquivo e enviar valor com vírgula ---
        # --- MODO EVEREST: remover só as datas do arquivo; inserir novas; formatar valores com vírgula/2 casas ---
        if mode == "everest" and "df_everest" in st.session_state:
            df_file = st.session_state.df_everest.copy()
        
            import re, unicodedata
        
            def _norm(s: str) -> str:
                s = unicodedata.normalize('NFKD', str(s)).encode('ASCII','ignore').decode('ASCII')
                s = s.lower()
                s = re.sub(r'[^a-z0-9]+', ' ', s)
                return re.sub(r'\s+', ' ', s).strip()
        
            # Detecta colunas no ARQUIVO
            date_file_col = st.session_state.get("everest_date_col")
            if not date_file_col or date_file_col not in df_file.columns:
                for cand in ["D. Lançamento", "D.Lançamento", "D. Lancamento", "D.Lancamento"]:
                    if cand in df_file.columns:
                        date_file_col = cand
                        break
            if not date_file_col or date_file_col not in df_file.columns:
                st.error("❌ Para atualizar a aba **Sangria Everest**, preciso da coluna **'D. Lançamento'** no arquivo.")
                st.stop()
        
            # Detecta colunas de valor no ARQUIVO
            def detect_valor_col(cols):
                aliases = {
                    "valor lancamento", "valor lançamento",
                    "valor do lancamento", "valor de lancamento",
                    "valor do lançamento", "valor de lançamento"
                }
                for c in cols:
                    if _norm(c) in aliases:
                        return c
                return None
        
            def detect_rateio_col(cols):
                aliases = {"v rateio", "v. rateio", "valor rateio"}
                for c in cols:
                    if _norm(c) in aliases:
                        return c
                return None
        
            valor_file_col  = detect_valor_col(df_file.columns)
            rateio_file_col = detect_rateio_col(df_file.columns)
        
            # Conversor robusto pt-BR → número
            def to_number_br(series):
                def _one(x):
                    if pd.isna(x): return 0.0
                    s = str(x).strip()
                    if s == "": return 0.0
                    neg = False
                    if s.startswith("(") and s.endswith(")"):
                        neg = True; s = s[1:-1].strip()
                    s = s.replace("R$", "").replace("r$", "").strip()
                    if s.endswith("-"):
                        neg = True; s = s[:-1].strip()
                    s = s.replace(".", "").replace(",", ".")
                    s_clean = re.sub(r"[^0-9.\-]", "", s)
                    if s_clean in ["", "-", "."]: return 0.0
                    try:
                        val = float(s_clean)
                    except:
                        s_fallback = re.sub(r"[^0-9.]", "", s_clean)
                        val = float(s_fallback) if s_fallback else 0.0
                    return -abs(val) if neg else val
                return series.apply(_one)
        
            # ▶️ Datas como STRING dd/mm/aaaa (aceita texto e serial Excel)
            def date_to_str(series):
                def _parse_one(x):
                    if pd.isna(x): return ""
                    s = str(x).strip()
                    if s == "": return ""
                    # número/float → serial Excel (sistema 1900)
                    if re.fullmatch(r"\d+([.,]\d+)?", s):
                        try:
                            f = float(s.replace(",", "."))
                            dt = pd.Timestamp("1899-12-30") + pd.to_timedelta(f, unit="D")
                            return dt.strftime("%d/%m/%Y")
                        except Exception:
                            pass
                    # tenta dd/mm/aaaa etc. (dayfirst) e ISO
                    dt = pd.to_datetime(s, dayfirst=True, errors="coerce")
                    if pd.isna(dt):
                        dt = pd.to_datetime(s, errors="coerce")
                    return "" if pd.isna(dt) else dt.strftime("%d/%m/%Y")
                return series.apply(_parse_one)
        
            # Conjunto de datas do ARQUIVO (strings dd/mm/aaaa)
            file_dates_str = set([d for d in date_to_str(df_file[date_file_col]).tolist() if d])
        
            if not file_dates_str:
                st.error("❌ A coluna **'D. Lançamento'** do arquivo não possui datas válidas.")
                st.stop()
        
            # Abre a aba destino
            try:
                ws = planilha.worksheet("Sangria Everest")
            except Exception as e:
                st.error(f"❌ Não consegui abrir a aba 'Sangria Everest': {e}")
                st.stop()
        
            rows = ws.get_all_values()
            if not rows:
                # planilha vazia → escreve arquivo já formatando
                header_sheet = list(df_file.columns)
                df_insert = df_file.copy()
                # Data do arquivo como dd/mm/aaaa para manter consistência
                df_insert[date_file_col] = date_to_str(df_insert[date_file_col])
        
                def to_str_comma(series_like):
                    if pd.api.types.is_numeric_dtype(series_like):
                        nums = pd.to_numeric(series_like, errors="coerce").fillna(0.0)
                    else:
                        nums = to_number_br(series_like)
                    return nums.apply(lambda v: f"{float(v):.2f}".replace(".", ","))
        
                if valor_file_col and valor_file_col in df_insert.columns:
                    df_insert[valor_file_col] = to_str_comma(df_insert[valor_file_col])
                if rateio_file_col and rateio_file_col in df_insert.columns:
                    df_insert[rateio_file_col] = to_str_comma(df_insert[rateio_file_col])
        
                values = [header_sheet] + df_insert.fillna("").astype(str).values.tolist()
                ws.clear()

                ws.update("A1", values, value_input_option="USER_ENTERED")
                
                # ⬇️ ADICIONE ESTE TRECHO
                # Descobre as colunas "Valor Lançamento" e "V. Rateio" pelo cabeçalho
                valor_sheet_col  = None
                rateio_sheet_col = None
                for c in header_sheet:
                    if _norm(c) in {"valor lancamento", "valor lançamento",
                                    "valor do lancamento", "valor de lancamento",
                                    "valor do lançamento", "valor de lançamento"}:
                        valor_sheet_col = c
                    if _norm(c) in {"v rateio", "v. rateio", "valor rateio"}:
                        rateio_sheet_col = c
                
                last_row = 1 + len(df_insert)  # 1 (cabeçalho) + linhas de dados
                if valor_sheet_col:
                    col_letter = _excel_col_letter(header_sheet.index(valor_sheet_col))
                    format_cell_range(ws, f"{col_letter}2:{col_letter}{last_row}", ACCOUNTING_RS)
                if rateio_sheet_col:
                    col_letter = _excel_col_letter(header_sheet.index(rateio_sheet_col))
                    format_cell_range(ws, f"{col_letter}2:{col_letter}{last_row}", ACCOUNTING_RS)
                # ⬆️ FIM DO TRECHO
                
                # (já existente)
                st.success(f"✅ 'Sangria Everest' criada com {len(df_insert)} linhas.")
                st.balloons()
                st.stop()

                
                
                
                
                
                st.success(f"✅ 'Sangria Everest' criada com {len(df_insert)} linhas.")
                st.balloons()
                st.stop()
        
            # Já existe conteúdo no sheet → REMOVER APENAS AS DATAS DO ARQUIVO e inserir as novas
            header_sheet = rows[0]
            data_sheet   = rows[1:]
            df_sheet     = pd.DataFrame(data_sheet, columns=header_sheet)
        
            # Detecta coluna de data no SHEET (equivalente a D. Lançamento)
            date_sheet_col = None
            for c in df_sheet.columns:
                if _norm(c) in {"d lancamento", "data lancamento", "d lancamento data", "d lancamento.", "d. lancamento", "d.lancamento"}:
                    date_sheet_col = c
                    break
            if not date_sheet_col:
                st.error("❌ A aba **Sangria Everest** não tem uma coluna de data equivalente a 'D. Lançamento'. Nada foi alterado.")
                st.stop()
        
            # Detecta colunas de valor no SHEET (para formatar)
            valor_sheet_col  = None
            rateio_sheet_col = None
            for c in df_sheet.columns:
                if _norm(c) in {"valor lancamento", "valor lançamento", "valor do lancamento", "valor de lancamento", "valor do lançamento", "valor de lançamento"}:
                    valor_sheet_col = c
                if _norm(c) in {"v rateio", "v. rateio", "valor rateio"}:
                    rateio_sheet_col = c
        
            # 1) Mantém SOMENTE as linhas cujas datas NÃO estão no arquivo (comparação por string dd/mm/aaaa)
            sheet_dates_str = date_to_str(df_sheet[date_sheet_col])
            mask_remove = sheet_dates_str.isin(file_dates_str) & sheet_dates_str.ne("")
            kept = df_sheet.loc[~mask_remove].copy()
            removidas = int(mask_remove.sum())
        
            # 2) Alinhar as colunas do ARQUIVO à ORDEM do SHEET
            df_insert = pd.DataFrame({col: (df_file[col] if col in df_file.columns else "") for col in header_sheet})
        
            # 3) Harmonizar a coluna de data inserida para dd/mm/aaaa (se existir nos dois lados)
            if date_file_col in df_file.columns and date_sheet_col in df_insert.columns:
                df_insert[date_sheet_col] = date_to_str(df_file[date_file_col])
        
            # 4) Formatar valores (vírgula e 2 casas) p/ Valor Lançamento e V. Rateio
            def to_str_comma(series_like):
                if pd.api.types.is_numeric_dtype(series_like):
                    nums = pd.to_numeric(series_like, errors="coerce").fillna(0.0)
                else:
                    nums = to_number_br(series_like)
                return nums.apply(lambda v: f"{float(v):.2f}".replace(".", ","))
        
            if valor_sheet_col:
                src_val = df_file[valor_file_col] if (valor_file_col and valor_file_col in df_file.columns) else df_insert.get(valor_sheet_col, "")
                df_insert[valor_sheet_col] = to_str_comma(src_val)
            if rateio_sheet_col:
                src_rat = df_file[rateio_file_col] if (rateio_file_col and rateio_file_col in df_file.columns) else df_insert.get(rateio_sheet_col, "")
                df_insert[rateio_sheet_col] = to_str_comma(src_rat)
        
            # 5) Final = OUTRAS DATAS (kept) + NOVAS (df_insert)
            df_final = pd.concat([kept, df_insert], ignore_index=True)
        
            # 6) Atualiza (mesmo layout: um botão)
            # ⏩ Botão super-rápido: deleta em FAIXAS + append único
            if st.button("📥 Atualizar Google Sheets Sangria Everest"):
                with st.spinner("🔄 Enviando..."):
                    import pandas as _pd
            
                    # 1) Quais linhas do SHEET remover? (datas que estão no arquivo)
                    sheet_dates_str = date_to_str(df_sheet[date_sheet_col])  # 'dd/mm/aaaa'
                    mask_remove = sheet_dates_str.isin(file_dates_str) & sheet_dates_str.ne("")
                    rows_to_delete = [i + 2 for i, rm in enumerate(mask_remove) if bool(rm)]  # 1-based; +1 p/ header, +1 p/ 0-index
            
                    # 1a) Agrupa linhas contíguas em FAIXAS para deletar por intervalo (muito mais rápido)
                    def compress_ranges(rows_1based):
                        if not rows_1based:
                            return []
                        rows = sorted(rows_1based)
                        start = prev = rows[0]
                        out = []
                        for r in rows[1:]:
                            if r == prev + 1:
                                prev = r
                            else:
                                out.append((start, prev))  # inclusive
                                start = prev = r
                        out.append((start, prev))
                        return out
            
                    ranges_1based = compress_ranges(rows_to_delete)
                    # Converte para índices 0-based/exclusivos exigidos pelo Sheets API
                    # Lembrando: linha 1 = header → já não entra; aqui estão linhas ≥2.
                    reqs = []
                    for (r1, r2) in sorted(ranges_1based, key=lambda t: t[0], reverse=True):
                        start_idx_0 = r1 - 1  # 0-based
                        end_idx_0   = r2      # exclusivo
                        reqs.append({
                            "deleteDimension": {
                                "range": {
                                    "sheetId": ws.id,           # id da worksheet (gid)
                                    "dimension": "ROWS",
                                    "startIndex": start_idx_0,
                                    "endIndex": end_idx_0
                                }
                            }
                        })
            
                    # 2) Dispara UM batch_update com todas as deleções (ordem descendente evita shift)
                    if reqs:
                        ws.spreadsheet.batch_update({"requests": reqs})
            
                    # 3) Prepara as NOVAS linhas (alinhadas ao cabeçalho do SHEET)
                    df_insert = _pd.DataFrame({col: (df_file[col] if col in df_file.columns else "") for col in header_sheet})
            
                    # 3a) Data no padrão dd/mm/aaaa na coluna do SHEET
                    if date_file_col in df_file.columns and date_sheet_col in df_insert.columns:
                        df_insert[date_sheet_col] = date_to_str(df_file[date_file_col])
            
                    # 3b) Valores com vírgula e 2 casas: Valor Lançamento e V. Rateio
                    def to_str_comma(series_like):
                        if _pd.api.types.is_numeric_dtype(series_like):
                            nums = _pd.to_numeric(series_like, errors="coerce").fillna(0.0)
                        else:
                            nums = to_number_br(series_like)
                        return nums.apply(lambda v: f"{float(v):.2f}".replace(".", ","))
            
                    if valor_sheet_col:
                        src_val = (df_file[valor_file_col] if (valor_file_col and valor_file_col in df_file.columns)
                                   else df_insert.get(valor_sheet_col, ""))
                        df_insert[valor_sheet_col] = to_str_comma(src_val)
            
                    if rateio_sheet_col:
                        src_rat = (df_file[rateio_file_col] if (rateio_file_col and rateio_file_col in df_file.columns)
                                   else df_insert.get(rateio_sheet_col, ""))
                        df_insert[rateio_sheet_col] = to_str_comma(src_rat)
            
                    # 4) Append ÚNICO das novas linhas (sem limpar o sheet)
                    novas_linhas = df_insert.fillna("").astype(str).values.tolist()

                    if novas_linhas:
                        ws.append_rows(novas_linhas, value_input_option="USER_ENTERED")
                    
                        # ⬇️ ADICIONE ESTE TRECHO
                        # Linhas novas começam após header + linhas mantidas
                        inicio = len(kept) + 2                # primeira linha nova (1=header)
                        fim    = inicio + len(novas_linhas) - 1
                    
                        # Usa os nomes reais no header_sheet detectados antes
                        if valor_sheet_col:
                            col_letter = _excel_col_letter(header_sheet.index(valor_sheet_col))
                            format_cell_range(ws, f"{col_letter}{inicio}:{col_letter}{fim}", ACCOUNTING_RS)
                        if rateio_sheet_col:
                            col_letter = _excel_col_letter(header_sheet.index(rateio_sheet_col))
                            format_cell_range(ws, f"{col_letter}{inicio}:{col_letter}{fim}", ACCOUNTING_RS)
                        # ⬆️ FIM DO TRECHO

                    st.success(
                        f"✅Atualização Concluida"
                    )


    
        # --- caso contrário, fluxo Colibri original ---
        else:
            if "df_sangria" not in st.session_state:
                st.warning("⚠️ Primeiro faça o upload e o processamento na Aba 1.")
            else:
                df_final = st.session_state.df_sangria.copy()
    
                # Colunas na ordem do destino
                destino_cols = [
                    "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
                    "Código Grupo Everest", "Funcionário", "Hora", "Descrição",
                    "Descrição Agrupada", "Meio de recebimento", "Valor(R$)",
                    "Mês", "Ano", "Duplicidade", "Sistema"
                ]
                faltantes = [c for c in destino_cols if c not in df_final.columns]
                if faltantes:
                    st.error(f"❌ Colunas ausentes para envio: {faltantes}")
                    st.stop()
    
                # Recalcula Duplicidade (Data + Hora + Código + Valor + Descrição)
                df_final["Descrição"] = (
                    df_final["Descrição"].astype(str).str.strip().str.lower().str.replace(r"\s+", " ", regex=True)
                )
                data_key = pd.to_datetime(df_final["Data"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d")
                hora_key = pd.to_datetime(df_final["Hora"], errors="coerce").dt.strftime("%H:%M:%S")
                df_final["Valor(R$)"] = pd.to_numeric(df_final["Valor(R$)"], errors="coerce").fillna(0.0).round(2)
                valor_centavos = (df_final["Valor(R$)"].astype(float) * 100).round().astype(int).astype(str)
                desc_key = df_final["Descrição"].fillna("").astype(str)
                df_final["Duplicidade"] = (
                    data_key.fillna("") + "|" +
                    hora_key.fillna("") + "|" +
                    df_final["Código Everest"].fillna("").astype(str) + "|" +
                    valor_centavos + "|" +
                    desc_key
                )
    
                # Inteiros opcionais
                for col in ["Código Everest", "Código Grupo Everest", "Ano"]:
                    df_final[col] = df_final[col].apply(lambda x: int(x) if pd.notnull(x) and str(x).strip() != "" else "")

                # Acessa a aba de destino
                aba_destino = planilha.worksheet("Sangria")
                valores_existentes = aba_destino.get_all_values()
                
                if not valores_existentes:
                    st.error("❌ A aba 'Sangria' está vazia. Crie o cabeçalho antes de enviar.")
                    st.stop()
                
                header_raw = valores_existentes[0]  # cabeçalho como está no Sheets (linha 1)
                
                # =========================
                # Normalização/matching
                # =========================
                import unicodedata
                import re
                
                def _normalize_name(s: str) -> str:
                    s = str(s or "").strip()
                    # remove acentos
                    s = unicodedata.normalize("NFKD", s).encode("ASCII", "ignore").decode("ASCII")
                    s = s.lower()
                    # troca separadores por espaço
                    s = re.sub(r"[_\-]+", " ", s)
                    # remove múltiplos espaços
                    s = re.sub(r"\s+", " ", s).strip()
                    return s
                
                # nomes "canônicos" (o que seu código espera)
                destino_cols = [
                    "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
                    "Código Grupo Everest", "Funcionário", "Hora", "Descrição",
                    "Descrição Agrupada", "Meio de recebimento", "Valor(R$)",
                    "Mês", "Ano", "Duplicidade", "Sistema"
                ]
                
                # versões normalizadas dos canônicos
                canon_norm = [_normalize_name(c) for c in destino_cols]
                
                # cria um índice do header existente por normalização
                header_norm_map = {}  # normalizado -> nome original do sheet
                for col_name in header_raw:
                    header_norm_map[_normalize_name(col_name)] = col_name
                
                # tenta mapear cada canônico para um nome real existente no sheet
                col_map = {}         # nome canonico -> nome existente no sheet (original)
                faltando = []        # canônicos que não foram encontrados
                for canon, canon_n in zip(destino_cols, canon_norm):
                    if canon_n in header_norm_map:
                        col_map[canon] = header_norm_map[canon_n]
                    else:
                        faltando.append(canon)
                
                if faltando:
                    # Diagnóstico amigável
                    st.error("❌ O cabeçalho da aba 'Sangria' não corresponde ao esperado.")
                    with st.expander("Ver diagnóstico"):
                        st.write("**Esperado (canônico):**", destino_cols)
                        st.write("**Encontrado (linha 1 do Sheet):**", header_raw)
                        # quais equivalências foram reconhecidas
                        reconhecidas = [f"{k} → {v}" for k, v in col_map.items()]
                        if reconhecidas:
                            st.write("**Equivalências reconhecidas:**", reconhecidas)
                        st.write("**Faltando no Sheet:**", faltando)
                        sobras = [h for h in header_raw if _normalize_name(h) not in set(canon_norm)]
                        if sobras:
                            st.write("**Colunas extras no Sheet (não usadas):**", sobras)
                
                    # 👉 BOTÃO OPCIONAL para corrigir só o cabeçalho (mantendo dados)
                    # Use APENAS se tiver certeza de que as linhas abaixo já estão na ordem das colunas canônicas!
                    if st.button("⚠️ Corrigir cabeçalho (linha 1) para o padrão esperado"):
                        # atualiza somente a linha 1 com os nomes canônicos
                        aba_destino.update("A1", [destino_cols])
                        st.success("✅ Cabeçalho atualizado. Rode o envio novamente.")
                    st.stop()
                
                # Se chegou aqui, todas as colunas canônicas existem no Sheet (mesmo que com outro nome/ordem)
                # Vamos alinhar o DataFrame à ORDEM ATUAL do Sheet, preservando layout existente.
                # Monta a ordem final de colunas para enviar, com base no header do Sheet:
                # - se a coluna do sheet está entre as que mapeiam para canônicas, usamos o nome canônico
                # - se for uma coluna extra do sheet, vamos preenchê-la com vazio para as novas linhas
                sheet_order_canon = []       # nomes CANÔNICOS na ordem do sheet
                sheet_order_real = []        # nomes REAIS do sheet na mesma ordem (útil para log)
                sheet_extras = []
                
                norm_to_canon = {_normalize_name(v): k for k, v in col_map.items()}  # nome real(normalizado) -> canônico
                
                for h in header_raw:
                    hn = _normalize_name(h)
                    if hn in norm_to_canon:
                        sheet_order_canon.append(norm_to_canon[hn])  # canônico correspondente
                        sheet_order_real.append(h)                   # nome real no sheet
                    else:
                        sheet_extras.append(h)
                
                # Reindexa df_final para a ordem do Sheet:
                # - primeiras colunas: os canônicos na ordem em que aparecem no sheet
                # - para colunas extra do sheet (que não existem no df_final), preenche com ""
                df_final = df_final.copy()
                
                for extra in sheet_extras:
                    # cria coluna vazia para cobrir colunas excedentes do sheet (se houverem)
                    df_final[extra] = ""
                
                # O df deve conter todas as colunas canônicas; garantimos isso:
                for c in destino_cols:
                    if c not in df_final.columns:
                        df_final[c] = ""
                
                # Monta a lista de colunas finais na ordem do sheet:
                colunas_finais = []
                for h in header_raw:
                    hn = _normalize_name(h)
                    if hn in norm_to_canon:
                        # pega o canônico correspondente
                        ccanon = norm_to_canon[hn]
                        colunas_finais.append(ccanon)
                    else:
                        # coluna extra do sheet
                        colunas_finais.append(h)
                
                # aplica a ordem
                df_final = df_final[colunas_finais].fillna("")
                
                # =========================
                # Duplicidade (usando a coluna 'Duplicidade' CANÔNICA)
                # =========================
                # encontra a posição da coluna 'Duplicidade' na ORDEM DO SHEET
                try:
                    dup_idx = header_raw.index(col_map["Duplicidade"])  # nome real do sheet para 'Duplicidade'
                except Exception:
                    # fallback: tenta achar 'Duplicidade' literal no header
                    try:
                        dup_idx = header_raw.index("Duplicidade")
                    except ValueError:
                        st.error("❌ Cabeçalho não contém a coluna 'Duplicidade'.")
                        st.stop()
                
                # ⚠️ CHAVES JÁ EXISTENTES (apenas do Google Sheets!)
                dados_existentes = set([
                    linha[dup_idx] for linha in valores_existentes[1:]
                    if len(linha) > dup_idx and linha[dup_idx] != ""
                ])
                # 🔽 AQUI: garanta que Valor(R$) é número antes de virar lista
                df_final["Valor(R$)"] = pd.to_numeric(df_final["Valor(R$)"], errors="coerce").fillna(0.0)
                # ✅ Ignorar duplicidade interna do arquivo, checar só com o Sheets
                novos_dados, duplicados_sheet = [], []
                for linha in df_final.values.tolist():
                    chave = linha[dup_idx]
                    if chave in dados_existentes:
                        duplicados_sheet.append(linha)
                    else:
                        novos_dados.append(linha)
                
                if st.button("📥 Atualizar Google Sheets Sangria"):
                    with st.spinner("🔄 Enviando..."):
                        if novos_dados:
                            aba_destino.append_rows(novos_dados, value_input_option="USER_ENTERED")
                
                            # Descobre índices (1-based) das colunas Data e Valor para formatar
                            try:
                                col_data_letter = chr(ord('A') + header_raw.index(col_map["Data"]))
                            except Exception:
                                col_data_letter = None
                
                            try:
                                col_valor_letter = chr(ord('A') + header_raw.index(col_map["Valor(R$)"]))
                            except Exception:
                                col_valor_letter = None
                

                            inicio = len(valores_existentes) + 1
                            fim = inicio + len(novos_dados) - 1
                            
                            if fim >= inicio:
                                # Data (mantém como está)
                                if col_data_letter:
                                    format_cell_range(
                                        aba_destino, f"{col_data_letter}{inicio}:{col_data_letter}{fim}",
                                        CellFormat(numberFormat=NumberFormat(type="DATE", pattern="dd/mm/yyyy"))
                                    )
                                # Valor(R$) em CONTÁBIL com R$
                                if col_valor_letter:
                                    format_cell_range(
                                        aba_destino, f"{col_valor_letter}{inicio}:{col_valor_letter}{fim}",
                                        ACCOUNTING_RS   # ✅ use o mesmo nome definido
                                    )

                
                            st.success(f"✅ {len(novos_dados)} registros enviados!")
                        if duplicados_sheet:
                            st.warning("⚠️ Alguns registros já existiam no Google Sheets e não foram enviados.")
