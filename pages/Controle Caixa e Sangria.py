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
    df_descricoes = pd.DataFrame(
        planilha.worksheet("Tabela Sangria").get_all_values(),
        columns=["Palavra-chave", "Descrição Agrupada"]
    )

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
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine="openpyxl") as writer:
                            df.to_excel(writer, index=False, sheet_name="Sangria")
                        output.seek(0)
                        st.download_button("📥Sangria Colibri",
                                           data=output, file_name="Sangria_estruturada.xlsx")
                    except KeyError as e:
                        st.error(f"❌ Coluna obrigatória ausente para o padrão Colibri: {e}")


    # ================
    # ================
    # 🔄 Aba 2 — Atualizar Google Sheets (layout unificado)
    # ================
    with tab2:
        st.markdown("🔗 [Abrir planilha Vendas diarias](https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU)")
    
        mode = st.session_state.get("mode")
    
        def normalize_dates(s):
            return pd.to_datetime(s, errors="coerce", dayfirst=True).dt.normalize()
    
        # ---------- MODO EVEREST ----------
        if mode == "everest" and "df_everest" in st.session_state:
            ws_name = "Sangria Everest"
            df_file = st.session_state.df_everest.copy()
            header_file = list(df_file.columns)
    
            # Colunas detectadas na Aba 1
            date_col = st.session_state.get("everest_date_col")
            valor_col = st.session_state.get("everest_value_col")
    
            # Fallbacks mínimos (caso algo não tenha ficado no estado)
            if date_col is None:
                for cand in ["D. Lançamento", "D.Lançamento", "D. Lancamento", "D.Lancamento"]:
                    if cand in df_file.columns:
                        date_col = cand; break
            if valor_col is None:
                for cand in ["Valor Lancamento ", "Valor Lançamento ", "Valor Lancamento", "Valor Lançamento"]:
                    if cand in df_file.columns:
                        valor_col = cand; break
    
            # Período
            periodo_txt = "—"
            if date_col and date_col in df_file.columns:
                dt = pd.to_datetime(df_file[date_col], errors="coerce", dayfirst=True).dropna()
                if not dt.empty:
                    periodo_txt = f"{dt.min():%d/%m/%Y} até {dt.max():%d/%m/%Y}"
                    datas_set = set(dt.dt.normalize().unique())
                else:
                    datas_set = set()
            else:
                datas_set = set()
                st.warning("⚠️ Coluna de data 'D. Lançamento' não encontrada no arquivo.")
    
            # Total (Valor Lançamento) com sinal real
            total_txt = "—"
            if valor_col and valor_col in df_file.columns:
                # Conversor pt-BR robusto
                import re
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
    
                serie_val = pd.to_numeric(df_file[valor_col], errors="coerce") if pd.api.types.is_numeric_dtype(df_file[valor_col]) else to_number_br(df_file[valor_col])
                total_liquido = float(serie_val.sum())
                sinal = "-" if total_liquido < 0 else ""
                total_fmt = f"{abs(total_liquido):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                total_txt = f"{sinal}R$ {total_fmt}"
            else:
                st.info("ℹ️ Não encontrei coluna de valor (ex.: 'Valor Lançamento'). O total ficará vazio.")
    
            # ---- Métricas (layout unificado) ----
            m1, m2, m3 = st.columns(3)
            m1.metric("📅 Período processado", periodo_txt)
            m2.metric("🧾 Linhas do arquivo", f"{len(df_file)}")
            m3.metric("💰 Total (Valor Lançamento)", total_txt)
    
            # ---- Alcance da atualização (datas a substituir) ----
            if not datas_set:
                st.warning("⚠️ Não há datas válidas em 'D. Lançamento' para substituir na planilha.")
            else:
                st.caption(f"Alcance da atualização: **{len(datas_set)}** data(s) em 'D. Lançamento'.")
    
            # ---- Pré-cálculo de impacto na planilha (mantém o mesmo visual para os dois modos) ----
            try:
                ws = planilha.worksheet(ws_name)
            except Exception as e:
                st.error(f"❌ Não consegui abrir a aba '{ws_name}': {e}")
                ws = None
    
            linhas_removidas = 0
            linhas_finais = None
            if ws:
                rows = ws.get_all_values()
                if not rows:
                    linhas_removidas = 0
                    linhas_finais = len(df_file)
                else:
                    header_sheet = rows[0]
                    data_sheet = rows[1:]
                    import pandas as _pd
                    df_sheet = _pd.DataFrame(data_sheet, columns=header_sheet)
                    # alinhar colunas ao cabeçalho do arquivo
                    for c in header_file:
                        if c not in df_sheet.columns:
                            df_sheet[c] = ""
                    df_sheet = df_sheet[header_file]
                    if date_col in df_sheet.columns:
                        datas_sheet_norm = normalize_dates(df_sheet[date_col])
                        linhas_removidas = int(datas_sheet_norm.isin(datas_set).sum())
                        linhas_finais = int(len(df_sheet) - linhas_removidas + len(df_file))
                    else:
                        # Se o destino não tiver a coluna -> reescreve inteiro
                        linhas_removidas = len(df_sheet)
                        linhas_finais = len(df_file)
    
            s1, s2 = st.columns(2)
            s1.metric("🧹 Linhas a substituir/remover", f"{linhas_removidas}" if linhas_finais is not None else "—")
            s2.metric("📊 Linhas totais após atualização", f"{linhas_finais}" if linhas_finais is not None else "—")
    
            # ---- Botão único (ação Everest) ----
            if st.button("🚀 Atualizar Google Sheets", type="primary", use_container_width=True):
                if not ws:
                    st.stop()
                # Recarrega conteúdo e executa a substituição
                rows = ws.get_all_values()
                if not rows:
                    values = [header_file] + df_file.fillna("").astype(str).values.tolist()
                    ws.clear()
                    ws.update("A1", values, value_input_option="USER_ENTERED")
                    st.success(f"✅ '{ws_name}' criada com {len(df_file)} linhas.")
                else:
                    header_sheet = rows[0]
                    data_sheet = rows[1:]
                    df_sheet = pd.DataFrame(data_sheet, columns=header_sheet)
                    for c in header_file:
                        if c not in df_sheet.columns:
                            df_sheet[c] = ""
                    df_sheet = df_sheet[header_file]
                    if date_col in df_sheet.columns and datas_set:
                        datas_sheet_norm = normalize_dates(df_sheet[date_col])
                        kept = df_sheet.loc[~datas_sheet_norm.isin(datas_set)].copy()
                        df_final = pd.concat([kept, df_file[header_file].copy()], ignore_index=True)
                    else:
                        df_final = df_file[header_file].copy()
                    values = [header_file] + df_final.fillna("").astype(str).values.tolist()
                    ws.clear()
                    ws.update("A1", values, value_input_option="USER_ENTERED")
                    st.success(f"✅ '{ws_name}' atualizada! Linhas finais: {len(df_final)}")
                st.balloons()
    
        # ---------- MODO COLIBRI ----------
        elif "df_sangria" in st.session_state:
            ws_name = "Sangria"
            df_final = st.session_state.df_sangria.copy()
    
            # Período / Total
            dt = pd.to_datetime(df_final["Data"], dayfirst=True, errors="coerce").dropna()
            periodo_txt = f"{dt.min():%d/%m/%Y} até {dt.max():%d/%m/%Y}" if not dt.empty else "—"
            total_val = float(pd.to_numeric(df_final["Valor(R$)"], errors="coerce").fillna(0.0).sum())
            sinal = "-" if total_val < 0 else ""
            total_fmt = f"{abs(total_val):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            total_txt = f"{sinal}R$ {total_fmt}"
    
            # Métricas (layout igual ao Everest)
            m1, m2, m3 = st.columns(3)
            m1.metric("📅 Período processado", periodo_txt)
            m2.metric("🧾 Linhas do arquivo", f"{len(df_final)}")
            m3.metric("💰 Total (Valor R$)", total_txt)
    
            # Alcance da atualização = registros novos (não duplicados)
            try:
                aba_destino = planilha.worksheet(ws_name)
                valores_existentes = aba_destino.get_all_values()
            except Exception as e:
                st.error(f"❌ Não consegui abrir a aba '{ws_name}': {e}")
                valores_existentes = []
    
            novos_dados = []
            duplicados_sheet = []
    
            if valores_existentes:
                header = valores_existentes[0]
                destino_cols = [
                    "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
                    "Código Grupo Everest", "Funcionário", "Hora", "Descrição",
                    "Descrição Agrupada", "Meio de recebimento", "Valor(R$)",
                    "Mês", "Ano", "Duplicidade", "Sistema"
                ]
                if header[:len(destino_cols)] != destino_cols:
                    st.error("❌ O cabeçalho da aba 'Sangria' não corresponde ao esperado.")
                    st.stop()
                try:
                    dup_idx = header.index("Duplicidade")
                except ValueError:
                    st.error("❌ Cabeçalho da aba 'Sangria' não contém a coluna 'Duplicidade'.")
                    st.stop()
    
                dados_existentes = set([
                    linha[dup_idx] for linha in valores_existentes[1:]
                    if len(linha) > dup_idx and linha[dup_idx] != ""
                ])
    
                df_envio = df_final[destino_cols].fillna("")
                for linha in df_envio.values.tolist():
                    chave = linha[dup_idx]
                    (duplicados_sheet if chave in dados_existentes else novos_dados).append(linha)
            else:
                # Planilha vazia → tudo é novo
                destino_cols = [
                    "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
                    "Código Grupo Everest", "Funcionário", "Hora", "Descrição",
                    "Descrição Agrupada", "Meio de recebimento", "Valor(R$)",
                    "Mês", "Ano", "Duplicidade", "Sistema"
                ]
                df_envio = df_final[destino_cols].fillna("")
                novos_dados = df_envio.values.tolist()
    
            st.caption(f"Alcance da atualização: **{len(novos_dados)}** novo(s) registro(s) "
                       f"(duplicados ignorados: {len(duplicados_sheet)}).")
    
            # Métricas adicionais para manter o layout coerente
            s1, s2 = st.columns(2)
            s1.metric("🧹 Linhas a substituir/remover", "—")  # Colibri não remove por data
            s2.metric("📊 Linhas totais após atualização",
                      f"{(len(valores_existentes)-1 if valores_existentes else 0) + len(novos_dados)}")
    
            # Botão único (ação Colibri)
            if st.button("🚀 Atualizar Google Sheets", type="primary", use_container_width=True):
                if novos_dados:
                    aba_destino.append_rows(novos_dados, value_input_option="USER_ENTERED")
                    inicio = (len(valores_existentes) if valores_existentes else 1) + 1
                    fim = inicio + len(novos_dados) - 1
                    if valores_existentes and fim >= inicio:
                        format_cell_range(
                            aba_destino, f"A{inicio}:A{fim}",
                            CellFormat(numberFormat=NumberFormat(type="DATE", pattern="dd/mm/yyyy"))
                        )
                        format_cell_range(
                            aba_destino, f"L{inicio}:L{fim}",
                            CellFormat(numberFormat=NumberFormat(type="NUMBER", pattern="#,##0.00"))
                        )
                st.success(f"✅ '{ws_name}' atualizada!")
                st.balloons()
    
        else:
            st.warning("⚠️ Primeiro faça o upload na Aba 1.")
