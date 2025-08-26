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

# ================================
# 1. Conexão com Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Vendas diarias")

# ========================
# 🔐 Autenticação Google Sheets
# ========================
#from google.oauth2.service_account import Credentials
#import gspread
#import streamlit as st

#scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

#credentials_dict = st.secrets["GOOGLE_SERVICE_ACCOUNT"]  # já vem como dict
#credentials = Credentials.from_service_account_info(credentials_dict, scopes=scope)
#gc = gspread.authorize(credentials)

# Sempre use open_by_key (mais seguro que open pelo nome)
planilha_empresa = gc.open_by_key("1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU")



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
# Atualizar Google Sheets (Evitar duplicação)
# =======================================

with aba3:
        
    # 🔗 Link sempre visível
    st.markdown("""
      🔗 [Link  **Faturamento Sistema Externo**](https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU/edit?usp=sharing)
    """, unsafe_allow_html=True)
       
 
    #st.header("📤 Atualizar Banco de Dados (Evitar duplicação usando coluna M)")

    if 'df_final' in st.session_state:
        df_final = st.session_state.df_final.copy()


        # Mostra botão SEMPRE que houver dados na aba upload
        if st.button("📥 Enviar dados para o Google Sheets"):

            with st.spinner("🔄 Processando dados e verificando duplicidades..."):

            # Verifica se há lojas sem código Everest
                lojas_nao_cadastradas = df_final[df_final["Código Everest"].isna()]["Loja"].unique()

                # Só continua se todas estiverem cadastradas
                todas_lojas_ok = len(lojas_nao_cadastradas) == 0
                
                
                #🔗 Links úteis
                #st.markdown("""
                #  🔗 [Link  **Faturamento Sistema Externo**](https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU/edit?usp=sharing)
                #""", unsafe_allow_html=True)

                # Criar a coluna "M" com a concatenação de "Data", "Fat.Total" e "Loja" como string para verificação de duplicação
                df_final['M'] = pd.to_datetime(df_final['Data'], format='%d/%m/%Y').dt.strftime('%Y-%m-%d') + \
                                df_final['Fat.Total'].astype(str) + df_final['Loja'].astype(str)

                #df_final['M'] = df_final['Data'] + df_final['Fat.Total'].astype(str) + df_final['Loja'].astype(str)


                # Não converter para string, apenas utilizar "M" para verificação de duplicação
                df_final['M'] = df_final['M'].apply(str)

                # Converter o restante do DataFrame para string, mas mantendo as colunas numéricas com seu formato correto
                df_final = df_final.applymap(str)
                

            

                #TIRAR ASPAS DOS VALORES, DATA E NUMEROS

                
                
                # Formatando os valores monetários (não convertendo para string, mantendo como numérico)
                df_final['Fat.Total'] = df_final['Fat.Total'].apply(lambda x: float(x.replace(',', '.')) if isinstance(x, str) else x)
                df_final['Serv/Tx'] = df_final['Serv/Tx'].apply(lambda x: float(x.replace(',', '.')) if isinstance(x, str) else x)
                df_final['Fat.Real'] = df_final['Fat.Real'].apply(lambda x: float(x.replace(',', '.')) if isinstance(x, str) else x)
                df_final['Ticket'] = df_final['Ticket'].apply(lambda x: float(x.replace(',', '.')) if isinstance(x, str) else x)

                # Garantir datetime sem aspas
                df_final['Data'] = pd.to_datetime(df_final['Data'].astype(str).str.replace("'", "").str.strip(), dayfirst=True)

                # Converter para número serial (dias desde 1899-12-30, padrão do Excel/Sheets)
                df_final['Data'] = (df_final['Data'] - pd.Timestamp("1899-12-30")).dt.days
            
                # Corrigir coluna Ano: remover aspas, espaços e garantir que seja inteiro
                df_final['Ano'] = df_final['Ano'].apply(
                lambda x: int(str(x).replace("'", "").strip()) if pd.notnull(x) and str(x).strip() != "" else ""
                )

                # ✅ Função segura para conversão para inteiro
                def to_int_safe(x):
                    try:
                        x_clean = str(x).replace("'", "").strip()
                        return int(x_clean)
                    except:
                        return ""

                # ✅ Aplica conversão segura nas colunas de códigos
                df_final['Código Everest'] = df_final['Código Everest'].apply(to_int_safe)
                df_final['Código Grupo Everest'] = df_final['Código Grupo Everest'].apply(to_int_safe)
                
                
                # Conectar ao Google Sheets
                scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
                credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
                gc = gspread.authorize(credentials)

                planilha_destino = gc.open("Vendas diarias")
                aba_destino = planilha_destino.worksheet("Fat Sistema Externo")

                # Obter dados já existentes na aba
                from gspread_dataframe import get_as_dataframe

                valores_existentes_df = get_as_dataframe(
                    aba_destino, evaluate_formulas=True, dtype=str
                ).fillna("")
                valores_existentes = valores_existentes_df.values.tolist()  # ✅ Correção necessária aqui
                colunas_df_existente = valores_existentes_df.columns.str.strip().tolist()
                
                # 🔍 Verificação da coluna N
                dados_n_existentes = set()
                if "N" in colunas_df_existente:
                    dados_n_existentes = set(valores_existentes_df["N"].astype(str).str.strip())
                else:
                    st.warning("⚠️ A coluna 'N' não foi encontrada na planilha. Nenhuma checagem de duplicidade será feita com base nela.")
                
                # 🔍 Verificação da coluna M
                dados_existentes = set()
                if "M" in colunas_df_existente:
                    dados_existentes = set(valores_existentes_df["M"].astype(str).str.strip())
                else:
                    st.warning("⚠️ A coluna 'M' não foi encontrada na planilha. Nenhuma checagem de duplicidade será feita com base nela.")
                
                        
            
                #novos_dados = []
            
                # ✅ Cria a coluna N diretamente, sem deixar a Data_Formatada como coluna a ser exportada
                df_final['Código Everest'] = df_final['Código Everest'].apply(to_int_safe)

                df_final['Data_Formatada'] = pd.to_datetime(
                    df_final['Data'], origin="1899-12-30", unit='D'
                ).dt.strftime('%Y-%m-%d')
                df_final['N'] = df_final['Data_Formatada'] + df_final['Código Everest'].astype(str)
                df_final['N'] = df_final['N'].astype(str).str.strip()
                df_final = df_final[df_final['Código Everest'].notna() & (df_final['Código Everest'] != 0)]
                
                # ✅ Remove a coluna auxiliar antes de montar os dados
                if 'Data_Formatada' in df_final.columns:
                    df_final = df_final.drop(columns=['Data_Formatada'])
                colunas_df = df_final.columns.tolist()
                # ✅ Garante que vai usar o índice exato da coluna N
                # Garante que vai usar o índice exato da coluna N
                idx_coluna_n_df = colunas_df.index("N")
                idx_coluna_m_df = colunas_df.index("M")

                # Obter linhas do DataFrame como lista de listas
                rows = df_final.fillna("").values.tolist()

                
            
                duplicados = []  # Duplicados pela M
                suspeitos_n = []  # ⚠️ Possíveis duplicados pela N
                novos_dados = []



                # Diagnóstico: Ver interseções entre as chaves novas e existentes
                chaves_novas = set(df_final["N"].tolist())
                intersecao_n = chaves_novas & dados_n_existentes
               
                # =========================================
                # ✅ Verifica duplicidade pela M e depois N
                # =========================================
                for linha in rows:
                    linha_dict = dict(zip(colunas_df, linha))

                    chave_m = str(linha_dict["M"]).strip()
                    chave_n = str(linha_dict["N"]).strip()

                    if chave_m not in dados_existentes:
                        if chave_n in dados_n_existentes:
                            suspeitos_n.append(linha)  # ⚠️ Duplicado pela N
                        else:
                            novos_dados.append(linha)  # ✅ Livre para envio
                        dados_existentes.add(chave_m)
                    else:
                        duplicados.append(linha)  # ❌ Duplicado pela M

                # ==================================================
                # ✅ Mostra alerta para duplicidade pela coluna N
                # ==================================================
                pode_enviar = True  # Variável de controle

               
                # ================================
                # 🚨 Verifica duplicidade pela coluna N
                # ================================
                if suspeitos_n:
                    st.warning("❌ Existem registros possivelmente duplicados. Corrija antes de continuar.")
                    
                    df_exibir = pd.DataFrame(suspeitos_n, columns=colunas_df).copy()
                
                    # 🗓️ Converte o número serial para data legível (sem alterar o original)
                    df_exibir["Data"] = pd.to_datetime(df_exibir["Data"], origin="1899-12-30", unit="D").dt.strftime("%d/%m/%Y")
                
                    # 📊 Exibe a tabela com a data formatada
                    st.dataframe(df_exibir, use_container_width=True)
                
                    pode_enviar = False
                else:
                    pode_enviar = True

                    

                # =============================================
                # 🟢 Só mostra o botão se permitido pelo checkbox
                # =============================================
                if todas_lojas_ok and pode_enviar:
                    try:
                        dados_para_enviar = novos_dados + suspeitos_n
                
                        # Descobre a primeira linha livre
                        inicio = len(aba_destino.col_values(1)) + 1
                        aba_destino.append_rows(dados_para_enviar, value_input_option='USER_ENTERED')
                        fim = inicio + len(dados_para_enviar) - 1
                
                        # 🔹 Declara os formatos
                        from gspread_formatting import CellFormat, NumberFormat, format_cell_range
                
                        data_format = CellFormat(
                            numberFormat=NumberFormat(type='DATE', pattern='dd/mm/yyyy')
                        )
                
                        numero_format = CellFormat(
                            numberFormat=NumberFormat(type='NUMBER', pattern='0')
                        )
                
                        # 🔹 Aplica os formatos
                        format_cell_range(aba_destino, f"A{inicio}:A{fim}", data_format)
                        format_cell_range(aba_destino, f"L{inicio}:L{fim}", numero_format)
                        format_cell_range(aba_destino, f"D{inicio}:D{fim}", numero_format)
                        format_cell_range(aba_destino, f"F{inicio}:F{fim}", numero_format)
                
                        st.success(f"✅ {len(dados_para_enviar)} registro(s) enviado(s) com sucesso para o Google Sheets!")
                
                        if duplicados:
                            st.warning(f"⚠️ {len(duplicados)} registro(s) duplicados na google sheets, não foram enviados.")
                    except Exception as e:
                        st.error(f"❌ Erro ao atualizar o Google Sheets: {e}")



       

        
        
    
    from datetime import datetime
    import requests

    # 🔘 Botão que chama o Apps Script (após as 9h)
    def pode_executar_agora():
        agora = datetime.now()
        hora_local = agora.hour
        return hora_local >= 12


    #st.subheader("🚀 Atualização DRE")

    if pode_executar_agora():
        if st.button("📤 Atualizar DRE Após as 10h"):
            try:
                url_script = "https://script.google.com/macros/s/AKfycbw-gK_KYcSyqyfimHTuXFLEDxKvWdW4k0o_kOPE-r-SWxL-SpogE2U9wiZt7qCZoH-gqQ/exec"  # Substituir pelo seu link
                resposta = requests.get(url_script)

                if resposta.status_code == 200:
                    st.success("✅ Atualização realizada com sucesso!")
                    st.info(resposta.text)
                else:
                    st.error(f"❌ Erro ao executar o script: {resposta.status_code}")
            except Exception as e:
                st.error(f"❌ Falha ao conectar: {e}")
    else:
        st.warning("⏰ A atualização externa só está disponível após às 9h (horário de Brasília).")



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
