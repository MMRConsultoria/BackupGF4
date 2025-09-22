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
                    #st.success("🔎 Detectado **padrão Everest** (cabeçalho do arquivo será mantido).")
    
                    # Guarda no estado para a Tab2 decidir a atualização
                    st.session_state.mode = "everest"
                    st.session_state.df_everest = df.copy()
                    st.session_state.everest_date_col = "D. Lançamento" if "D. Lançamento" in df.columns else None
    
                    # Download do arquivo com cabeçalho original
                    output_ev = BytesIO()
                    with pd.ExcelWriter(output_ev, engine="openpyxl") as writer:
                        df.to_excel(writer, index=False, sheet_name="Sangria Everest")
                    output_ev.seek(0)
                    st.download_button(
                        "📥 Baixar arquivo (Everest)",
                        data=output_ev,
                        file_name="Sangria_Everest.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
    
                    st.info("ℹ️ A atualização no Google Sheets será feita na aba **“🔄 Atualizar Google Sheets”** "
                            "usando a coluna **'D. Lançamento'** para substituir as datas correspondentes.")
    
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
                        st.download_button("📥 Baixar relatório de sangria",
                                           data=output, file_name="Sangria_estruturada.xlsx")
                    except KeyError as e:
                        st.error(f"❌ Coluna obrigatória ausente para o padrão Colibri: {e}")


    # ================
    # 🔄 Aba 2 — Atualizar Google Sheets (aba: sangria — fluxo existente)
    # ================
    with tab2:
        st.markdown("🔗 [Abrir planilha Vendas diarias](https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU)")

        if "df_sangria" not in st.session_state:
            st.warning("⚠️ Primeiro faça o upload e o processamento na Aba 1 (modo Colibri).")
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

            # Inteiros opcionais (mantém string vazia quando não há número)
            for col in ["Código Everest", "Código Grupo Everest", "Ano"]:
                df_final[col] = df_final[col].apply(lambda x: int(x) if pd.notnull(x) and str(x).strip() != "" else "")

            # Acessa a aba de destino
            aba_destino = planilha.worksheet("Sangria")
            valores_existentes = aba_destino.get_all_values()
            if not valores_existentes:
                st.error("❌ A aba 'sangria' está vazia ou sem cabeçalho. Crie o cabeçalho antes de enviar.")
                st.stop()

            header = valores_existentes[0]
            if header[:len(destino_cols)] != destino_cols:
                st.error("❌ O cabeçalho da aba 'sangria' não corresponde ao esperado.")
                st.stop()

            # Índice da coluna 'Duplicidade' no destino
            try:
                dup_idx = header.index("Duplicidade")
            except ValueError:
                st.error("❌ Cabeçalho da aba 'sangria' não contém a coluna 'Duplicidade'.")
                st.stop()

            # ⚠️ CHAVES JÁ EXISTENTES (apenas do Google Sheets!)
            dados_existentes = set([
                linha[dup_idx] for linha in valores_existentes[1:]
                if len(linha) > dup_idx and linha[dup_idx] != ""
            ])

            # Prepara linhas na ordem do destino
            df_final = df_final[destino_cols].fillna("")

            # ✅ Ignorar duplicidade interna do arquivo, checar só com o Sheets
            novos_dados, duplicados_sheet = [], []
            for linha in df_final.values.tolist():
                chave = linha[dup_idx]
                if chave in dados_existentes:
                    duplicados_sheet.append(linha)
                else:
                    novos_dados.append(linha)

            if st.button("📥 Enviar dados para a aba 'sangria'"):
                with st.spinner("🔄 Enviando..."):
                    if novos_dados:
                        # USER_ENTERED => Sheets interpreta Data e Hora, valor numérico sem texto
                        aba_destino.append_rows(novos_dados, value_input_option="USER_ENTERED")

                        # ▸ Formatação das novas linhas
                        inicio = len(valores_existentes) + 1
                        fim = inicio + len(novos_dados) - 1

                        if fim >= inicio:
                            # Data (coluna A) -> dd/mm/yyyy
                            format_cell_range(
                                aba_destino, f"A{inicio}:A{fim}",
                                CellFormat(numberFormat=NumberFormat(type="DATE", pattern="dd/mm/yyyy"))
                            )
                            # Valor(R$) (coluna L) -> padrão locale: 1.000,00 em pt-BR
                            format_cell_range(
                                aba_destino, f"L{inicio}:L{fim}",
                                CellFormat(numberFormat=NumberFormat(type="NUMBER", pattern="#,##0.00"))
                            )

                        st.success(f"✅ {len(novos_dados)} registros enviados!")
                    if duplicados_sheet:
                        st.warning("⚠️ Alguns registros já existiam no Google Sheets e não foram enviados.")
