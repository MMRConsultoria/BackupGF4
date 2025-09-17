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

NOME_SISTEMA = "Sangria"

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
    # 📥 Aba 1
    # ================
    with tab1:
        uploaded_file = st.file_uploader(
            label="📁 Clique para selecionar ou arraste aqui o arquivo Excel com os dados de sangria",
            type=["xlsx", "xlsm"],
            help="Somente arquivos .xlsx ou .xlsm. Tamanho máximo: 200MB."
        )

        if uploaded_file:
            try:
                xls = pd.ExcelFile(uploaded_file)
                df_dados = pd.read_excel(xls, sheet_name="Sheet")
            except Exception as e:
                st.error(f"❌ Não foi possível ler o arquivo enviado. Detalhes: {e}")
            else:
                df = df_dados.copy()

                # Campos que serão preenchidos durante o parsing
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
                df["Descrição"] = df["Descrição"].astype(str).str.strip().str.lower()
                df["Funcionário"] = df["Funcionário"].astype(str).str.strip()
                df["Valor(R$)"] = pd.to_numeric(df["Valor(R$)"], errors="coerce").fillna(0.0)

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

                # 🔑 DUPLICIDADE = Data(YYYY-MM-DD) + Hora(HH:MM:SS) + Código Everest + Valor em centavos (inteiro)
                data_key = pd.to_datetime(df["Data"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d")
                hora_key = pd.to_datetime(df["Hora"], errors="coerce").dt.strftime("%H:%M:%S")
                valor_centavos = (df["Valor(R$)"].astype(float).round(2) * 100).astype(int).astype(str)
                df["Duplicidade"] = (
                    data_key.fillna("") + "|" +
                    hora_key.fillna("") + "|" +
                    df["Código Everest"].fillna("").astype(str) + "|" +
                    valor_centavos
                )

                # Garante coluna opcional
                if "Meio de recebimento" not in df.columns:
                    df["Meio de recebimento"] = ""

                # Ordenação conforme cabeçalho da aba "sangria"
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
                valor_total = df["Valor(R$)"].sum()

                col1, col2 = st.columns(2)
                col1.metric("📅 Período processado", f"{periodo_min} até {periodo_max}")
                col2.metric("💰 Valor total de sangria",
                            f"R$ {valor_total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

                st.success("✅ Relatório gerado com sucesso!")

                # Aviso de lojas sem código
                lojas_sem_codigo = df[df["Código Everest"].isna()]["Loja"].unique()
                if len(lojas_sem_codigo) > 0:
                    st.warning(
                        f"⚠️ Lojas sem Código Everest cadastrado: {', '.join(lojas_sem_codigo)}\n\n"
                        "🔗 Atualize na planilha de empresas."
                    )

                # Guarda para Aba 2
                st.session_state.df_sangria = df.copy()

                # Download
                output = BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df.to_excel(writer, index=False, sheet_name="Sangria")
                output.seek(0)
                st.download_button("📥 Baixar relatório de sangria",
                                   data=output, file_name="Sangria_estruturada.xlsx")

    # ================
    # 🔄 Aba 2 — Atualizar Google Sheets (aba: sangria)
    # ================
    with tab2:
        st.markdown("🔗 [Abrir planilha Vendas diarias](https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU)")

        if "df_sangria" not in st.session_state:
            st.warning("⚠️ Primeiro faça o upload e o processamento na Aba 1.")
        else:
            df_final = st.session_state.df_sangria.copy()

            # Valida colunas necessárias (conforme cabeçalho da aba 'sangria')
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

            # Recalcula Duplicidade por garantia (Data + Hora + Código + Valor em centavos)
            data_key = pd.to_datetime(df_final["Data"], dayfirst=True, errors="coerce").dt.strftime("%Y-%m-%d")
            hora_key = pd.to_datetime(df_final["Hora"], errors="coerce").dt.strftime("%H:%M:%S")
            df_final["Valor(R$)"] = pd.to_numeric(df_final["Valor(R$)"], errors="coerce").fillna(0.0)
            valor_centavos = (df_final["Valor(R$)"].astype(float).round(2) * 100).astype(int).astype(str)
            df_final["Duplicidade"] = (
                data_key.fillna("") + "|" +
                hora_key.fillna("") + "|" +
                df_final["Código Everest"].fillna("").astype(str) + "|" +
                valor_centavos
            )

            # Inteiros opcionais (mantém strings vazias quando não há número)
            for col in ["Código Everest", "Código Grupo Everest", "Ano"]:
                df_final[col] = df_final[col].apply(lambda x: int(x) if pd.notnull(x) and str(x).strip() != "" else "")

            # Verifica lojas sem código
            lojas_nao_cadastradas = df_final[df_final["Código Everest"].isin(["", np.nan])]["Loja"].unique()
            todas_lojas_ok = len(lojas_nao_cadastradas) == 0
            if not todas_lojas_ok:
                st.warning(f"⚠️ Existem lojas sem Código Everest: {', '.join(lojas_nao_cadastradas)}")

            # Acessa a aba de destino
            aba_destino = planilha.worksheet("sangria")
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

            # Chaves já existentes
            dados_existentes = set([linha[dup_idx] for linha in valores_existentes[1:]
                                    if len(linha) > dup_idx and linha[dup_idx] != ""])

            # Prepara linhas na ordem do destino
            df_final = df_final[destino_cols].fillna("")

            novos_dados, duplicados = [], []
            for linha in df_final.values.tolist():
                chave = linha[dup_idx]
                if chave not in dados_existentes:
                    novos_dados.append(linha)
                    dados_existentes.add(chave)
                else:
                    duplicados.append(linha)

            st.write(f"🧮 Prontos para envio: {len(novos_dados)} | Duplicados detectados: {len(duplicados)}")

            if todas_lojas_ok and st.button("📥 Enviar dados para a aba 'sangria'"):
                with st.spinner("🔄 Enviando..."):
                    if novos_dados:
                        # USER_ENTERED => Sheets interpreta Data (dd/mm/yyyy) como data
                        aba_destino.append_rows(novos_dados, value_input_option="USER_ENTERED")

                        # Formatação das novas linhas (Data e Valor) para exibir como 1.000,00
                        inicio = len(valores_existentes) + 1  # primeira linha dos novos dados
                        fim = inicio + len(novos_dados) - 1

                        # Data dd/mm/yyyy
                        format_cell_range(
                            aba_destino, f"A{inicio}:A{fim}",
                            CellFormat(numberFormat=NumberFormat(type="DATE", pattern="dd/mm/yyyy"))
                        )
                        # Valor(R$) com separador BR
                        format_cell_range(
                            aba_destino, f"L{inicio}:L{fim}",
                            CellFormat(numberFormat=NumberFormat(type="NUMBER", pattern="#.##0,00"))
                        )

                        st.success(f"✅ {len(novos_dados)} registros enviados!")
                    if duplicados:
                        st.warning("⚠️ Alguns registros duplicados não foram enviados (chave: Data+Hora+Código+Valor).")
