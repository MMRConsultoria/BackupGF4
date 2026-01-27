    

# pages/OperacionalVendasDiarias.py

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import date, datetime, timedelta
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import plotly.express as px
from datetime import date
import psycopg2

st.set_page_config(page_title="Vendas Diarias", layout="wide")

# 🔒 Bloqueia o acesso caso o usuário não esteja logado
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ✅ Inicializar controle de modo
if "modo_3s" not in st.session_state:
    st.session_state.modo_3s = False

# ✅ Função para limpar estado da Aba 2
def limpar_estado_aba_google():
    chaves = [
        "df_google",
        "df_google_sheets",
        "df_para_atualizar",
        "mensagem_atualizacao",
        "mensagem_sucesso",
        "atualizou_google",
        "preview_google",
    ]
    for k in chaves:
        if k in st.session_state:
            del st.session_state[k]

# ======================
# CSS para esconder só a barra superior
# ======================
# ADICIONE AQUI:
# ======= CSS QUE TRAVA O BOTÃO (COLE NO TOPO DO ARQUIVO) =======
st.markdown(
    """
    <style>
    /* Container fixo que NÃO encolhe */
    div.botao-vermelho {
        display: inline-flex !important;
        flex: 0 0 220px !important;    /* NÃO encolher; largura fixa */
        min-width: 220px !important;
        max-width: 220px !important;
        box-sizing: border-box !important;
        margin: 8px 0 !important;
    }

    /* Botão ocupa 100% do container fixo e não quebra texto */
    div.botao-vermelho > button {
        width: 100% !important;
        height: 40px !important;
        min-height: 40px !important;
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        box-sizing: border-box !important;

        /* Estética (opcional) */
        background-color: #ff4b4b !important;
        color: white !important;
        font-weight: 600 !important;
        font-size: 14px !important;
        border-radius: 6px !important;
        border: none !important;
        padding: 0 12px !important;
        box-shadow: none !important;
    }

    /* Se Streamlit usar elementos internos (<p>, <span>) */
    div.botao-vermelho > button * {
        white-space: nowrap !important;
        overflow: hidden !important;
        text-overflow: ellipsis !important;
    }

    div.botao-vermelho > button:hover {
        background-color: #ff3333 !important;
        color: white !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)
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
    # ================================
    # 1. Conexão com Google Sheets - OTIMIZADO
    # ================================
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gc = gspread.authorize(credentials)
    planilha_empresa = gc.open("Vendas diarias")
    
    # ✅ OTIMIZAÇÃO: Carrega valores brutos e normaliza ANTES de criar DataFrame
    aba_empresa = planilha_empresa.worksheet("Tabela Empresa")
    valores_empresa = aba_empresa.get_all_values()
    
    if len(valores_empresa) > 1:
        # Cria DataFrame com cabeçalho
        df_empresa = pd.DataFrame(valores_empresa[1:], columns=valores_empresa[0])
        df_empresa.columns = df_empresa.columns.str.strip()
        
        # ✅ Força Loja em minúsculo IMEDIATAMENTE após carregar
        if "Loja" in df_empresa.columns:
            df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.lower().str.strip()
    else:
        df_empresa = pd.DataFrame()
    # ================================
    # 2. Função de conexão com PostgreSQL
    # ================================
    CERT_PATH = "aws-us-east-2-bundle.pem"
    
    if "cert_written" not in st.session_state:
        with open(CERT_PATH, "w", encoding="utf-8") as f:
            f.write(st.secrets["certs"]["aws_rds_us_east_2"])
        st.session_state["cert_written"] = True
    
    def get_db_conn():
        return psycopg2.connect(
            host=st.secrets["db"]["host"],
            port=st.secrets["db"]["port"],
            dbname=st.secrets["db"]["database"],
            user=st.secrets["db"]["user"],
            password=st.secrets["db"]["password"],
            sslmode="verify-full",
            sslrootcert=CERT_PATH,
        )
    
    
    
    def buscar_dados_3s_checkout():
        """Busca dados do 3S Checkout direto do banco e processa"""
        conn = get_db_conn()
        try:
            # Ajuste para fuso horário de Brasília (UTC-3) e define "ontem" como limite máximo
            from datetime import datetime, timedelta
            agora_brasil = datetime.utcnow() - timedelta(hours=3)
            ontem = (agora_brasil - timedelta(days=1)).date()
            data_inicio = ontem - timedelta(days=29)  # últimos 30 dias incluindo ontem
            
            # ✅ FILTRO SQL: Adicionado "AND business_dt <= %s"
            query = """
                SELECT store_code, business_dt, total_gross, custom_properties, order_code, state_id
                FROM public.order_picture
                WHERE business_dt >= %s
                  AND business_dt <= %s
                  AND store_code NOT IN ('0000', '0001', '9999')
                  AND state_id = 5
            """
            df = pd.read_sql(query, conn, params=(data_inicio, ontem))
            
            
            # 1. Converter datas
            df['business_dt'] = pd.to_datetime(df['business_dt'], errors='coerce')
            
            # 2. ✅ REMOVER ZEROS À ESQUERDA do store_code
            df['store_code'] = df['store_code'].astype(str).str.lstrip('0')
            
            # 3. Extrair campos de custom_properties
            def parse_props(x):
                if pd.isna(x): return {}
                try:
                    if isinstance(x, str):
                        return json.loads(x)
                except:
                    try:
                        import ast
                        return ast.literal_eval(x)
                    except:
                        return {}
                return x if isinstance(x, dict) else {}
            
            props = df['custom_properties'].apply(parse_props)
            df['TIP_AMOUNT'] = pd.to_numeric(props.apply(lambda x: x.get('TIP_AMOUNT')), errors='coerce').fillna(0)
            df['VOID_TYPE'] = props.apply(lambda x: x.get('VOID_TYPE'))
            
            # 4. Desconsiderar registros com VOID_TYPE preenchido
            df = df[df['VOID_TYPE'].isna() | (df['VOID_TYPE'] == "") | (df['VOID_TYPE'] == 0)].copy()
            
            # 5. Criar coluna de data sem hora para agrupamento
            df['data'] = df['business_dt'].dt.date
            
            # 6. Agrupar totais por store_code e data
            resumo = df.groupby(['store_code', 'data']).agg(
                Fat_Real=('total_gross', 'sum'),
                Serv_Tx=('TIP_AMOUNT', 'sum')
            ).reset_index()
            
            # 7. Calcular Fat.Total
            resumo['Fat_Total'] = resumo['Fat_Real'] + resumo['Serv_Tx']
            
            # 8. Renomear colunas para o formato correto
            resumo.columns = ['Código Everest', 'Data', 'Fat.Real', 'Serv/Tx', 'Fat.Total']
            
            # 9. Formatar Data
            resumo['Data'] = pd.to_datetime(resumo['Data']).dt.strftime('%d/%m/%Y')
            
            # 10. Adicionar Dia da Semana
            dias_traducao = {
                "Monday": "segunda-feira", "Tuesday": "terça-feira", "Wednesday": "quarta-feira",
                "Thursday": "quinta-feira", "Friday": "sexta-feira", "Saturday": "sábado", "Sunday": "domingo"
            }
            resumo.insert(1, 'Dia da Semana', pd.to_datetime(resumo['Data'], format='%d/%m/%Y').dt.day_name().map(dias_traducao))
            
            # 11. ✅ Buscar informações da Tabela Empresa (também sem zeros à esquerda)
            df_empresa["Código Everest"] = (
                df_empresa["Código Everest"]
                .astype(str)
                .str.replace(r"\D", "", regex=True)
                .str.lstrip("0")
            )
            resumo["Código Everest"] = resumo["Código Everest"].astype(str).str.strip()
            
            resumo = pd.merge(resumo, df_empresa[["Código Everest", "Loja", "Grupo", "Código Grupo Everest"]], 
                             on="Código Everest", how="left")
            # ✅ Converte nome da loja para MAIÚSCULO
            resumo["Loja"] = resumo["Loja"].astype(str).str.strip().str.lower()
            # 12. Adicionar colunas adicionais
            resumo['Ticket'] = 0  # Não temos essa informação no agrupamento
            resumo['Mês'] = pd.to_datetime(resumo['Data'], format='%d/%m/%Y').dt.strftime('%b').str.lower()
            
            meses = {"jan": "jan", "feb": "fev", "mar": "mar", "apr": "abr", "may": "mai", "jun": "jun",
                     "jul": "jul", "aug": "ago", "sep": "set", "oct": "out", "nov": "nov", "dec": "dez"}
            resumo["Mês"] = resumo["Mês"].map(meses)
            
            resumo['Ano'] = pd.to_datetime(resumo['Data'], format='%d/%m/%Y').dt.year
            resumo['Sistema'] = '3SCheckout'
            
            # 13. Ordenar colunas no formato padrão
            colunas_finais = [
                "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
                "Código Grupo Everest", "Fat.Total", "Serv/Tx", "Fat.Real",
                "Ticket", "Mês", "Ano", "Sistema"
            ]
            
            resumo = resumo[[c for c in colunas_finais if c in resumo.columns]]
            
            # 14. Arredondar valores
            for col in ["Fat.Total", "Serv/Tx", "Fat.Real", "Ticket"]:
                if col in resumo.columns:
                    resumo[col] = resumo[col].round(2)
            
            # 15. Ordenar por Data e Loja
            resumo['Data_Ordenada'] = pd.to_datetime(resumo['Data'], format='%d/%m/%Y')
            resumo = resumo.sort_values(by=['Data_Ordenada', 'Loja']).drop(columns='Data_Ordenada')
            
            return resumo, None, len(df)
        except Exception as e:
            return None, str(e), 0
        finally:
            conn.close()
    
    # ================================
    # Configuração inicial do app
    # ================================
    
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

        if len(data_raw) > 1:
            df = pd.DataFrame(data_raw[1:], columns=data_raw[0])
            df.columns = df.columns.str.strip()

            if "Data" in df.columns:
                df["Data"] = pd.to_datetime(df["Data"].astype(str).str.strip(), dayfirst=True, errors="coerce")

                ultima_data_valida = df["Data"].dropna()

                if not ultima_data_valida.empty:
                    ultima_data = ultima_data_valida.max().strftime("%d/%m/%Y")

                    df["Grupo"] = df["Grupo"].astype(str).str.strip().str.lower()
                    df["GrupoExibicao"] = df["Grupo"].apply(
                        lambda g: "Bares" if g in ["amata", "aurora"]
                        else "Kopp" if g == "kopp"
                        else "GF4"
                    )

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
    
    # Cabeçalho bonito
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
    
    # ================================
    # 📄 Aba 1 - Upload e Processamento
    # ================================
    
    with aba1:
        # ========== BOTÃO 3S CHECKOUT ==========
        st.markdown("### 🔄 Atualização Automática 3S Checkout")
    
        # wrapper flex para manter o botão à esquerda e o conteúdo à direita
        st.markdown('<div style="display:flex; align-items:flex-start; gap:20px;">', unsafe_allow_html=True)
    
        # container fixo para o botão — largura fixa definida pelo CSS global (.botao-vermelho)
        st.markdown('<div class="botao-vermelho">', unsafe_allow_html=True)
        if st.button("🔄 Atualizar 3S Checkout", key="btn_3s_vermelho"):
            st.session_state.modo_3s = True
            st.session_state.df_final = None  # limpa upload manual
            limpar_estado_aba_google()
    
            with st.spinner("Buscando dados do banco..."):
                resumo_3s, erro_3s, total_registros = buscar_dados_3s_checkout()
                if erro_3s:
                    st.error(f"❌ Erro ao buscar dados: {erro_3s}")
                elif resumo_3s is not None and not resumo_3s.empty:
                    st.session_state.resumo_3s = resumo_3s
                    st.session_state.total_registros_3s = total_registros
                    st.rerun()
                else:
                    st.warning("⚠️ Nenhum dado encontrado para o período.")
        st.markdown('</div>', unsafe_allow_html=True)
    
        # Conteúdo principal da aba (ocupando o resto do espaço)
        st.markdown('<div style="flex:1 1 auto;">', unsafe_allow_html=True)
        # -------------------- coloque aqui o resto do conteúdo da aba (tabelas, filtros etc.) --------------------
        st.markdown('</div>', unsafe_allow_html=True)
        
        # ========== EXIBIR RESULTADO 3S ==========
        if st.session_state.modo_3s and "resumo_3s" in st.session_state:
            resumo_3s = st.session_state.resumo_3s
            total_registros = st.session_state.total_registros_3s
            
            st.success(f"✅ {total_registros} registros processados com sucesso!")
            
            # Verificar empresas não localizadas
            empresas_nao_localizadas = resumo_3s[resumo_3s["Loja"].isna()]["Código Everest"].unique()
            if len(empresas_nao_localizadas) > 0:
                empresas_nao_localizadas_str = "<br>".join(empresas_nao_localizadas)
                mensagem = f"""
                ⚠️ {len(empresas_nao_localizadas)} código(s) não localizado(s) na Tabela Empresa! <br>{empresas_nao_localizadas_str}
                <br>✏️ Atualize a tabela clicando 
                <a href='https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU/edit?usp=drive_link' target='_blank'><strong>aqui</strong></a>.
                """
                st.markdown(mensagem, unsafe_allow_html=True)
            else:
                st.success("✅ Todas as lojas foram localizadas na Tabela_Empresa!")
            
            # Mostrar resumo do período
            datas_validas = pd.to_datetime(resumo_3s["Data"], format="%d/%m/%Y", errors='coerce').dropna()
            if not datas_validas.empty:
                data_inicial = datas_validas.min().strftime("%d/%m/%Y")
                data_final_str = datas_validas.max().strftime("%d/%m/%Y")
                valor_total = resumo_3s["Fat.Total"].sum()
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
            
            # Gera Excel no formato padrão
            output = BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                resumo_3s.to_excel(writer, sheet_name='Faturamento Servico', index=False)
            output.seek(0)
            
            st.download_button(
                label="📥 Baixar Excel 3S Checkout",
                data=output,
                file_name=f"3s_checkout_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # Botão para voltar ao upload manual
            st.markdown("---")
            if st.button("↩️ Voltar para Upload Manual"):
                st.session_state.modo_3s = False
                if "resumo_3s" in st.session_state:
                    del st.session_state.resumo_3s
                if "total_registros_3s" in st.session_state:
                    del st.session_state.total_registros_3s
                st.rerun()
        
        # ========== UPLOAD MANUAL (só aparece se não estiver no modo 3S) ==========
        if not st.session_state.modo_3s:
            st.markdown("---")
            st.markdown("### 📁 Upload Manual")
           
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
                    
                    # =====================================================
                    # FORMATO 3 — PRIMEIRA ABA | ID LOJA DINÂMICO
                    # =====================================================
                    else:
                        nome_aba = abas[0]
                        df_bruto = pd.read_excel(xls, sheet_name=nome_aba, header=None)
        
                        # localizar linha do cabeçalho (ID LOJA na coluna A)
                        linha_header = None
                        for i in range(len(df_bruto)):
                            if "ID LOJA" in str(df_bruto.iloc[i, 0]).upper():
                                linha_header = i
                                break
        
                        if linha_header is None:
                            st.error("❌ Arquivo não reconhecido. Não encontrei 'ID LOJA' na coluna A.")
                            st.stop()
        
                        # leitura dos dados (SEM cabeçalho)
                        df = pd.read_excel(
                            xls,
                            sheet_name=nome_aba,
                            skiprows=linha_header + 1,
                            header=None
                        )
        
                        # colunas fixas
                        # A=0 ID LOJA | C=2 DATA | G=6 Ticket | H=7 Fat.Total | L=11 Serv/Tx | M=12 Fat.Real
                        df = df.iloc[:, [0, 2, 6, 7, 11, 12]]
                        df = df.dropna(how="all")
        
                        # normalizações
                        df[0] = (
                            df[0]
                            .astype(str)
                            .str.replace(r"\D", "", regex=True)
                            .str.lstrip("0")
                        )
                        df = df[df[0] != "9999"]
                        df[2] = pd.to_datetime(df[2], dayfirst=True, errors="coerce")
                        df[7] = pd.to_numeric(df[7], errors="coerce")
                        df[11] = pd.to_numeric(df[11], errors="coerce")
                        df[12] = pd.to_numeric(df[12], errors="coerce")
                        df[6] = pd.to_numeric(df[6], errors="coerce")
                        df = df[~((df[7] == 0) & (df[12] == 0))]
                        df = df.dropna(subset=[0, 2])
        
                        # merge com Tabela Empresa (Código Everest = coluna C)
                        df_empresa["Código Everest"] = (
                            df_empresa["Código Everest"]
                            .astype(str)
                            .str.replace(r"\D", "", regex=True)
                            .str.lstrip("0")
                        )
        
                        df = df.merge(
                            df_empresa[["Código Everest", "Loja"]],
                            left_on=0,
                            right_on="Código Everest",
                            how="left"
                        )
        
                        df_final = (
                            df.groupby([2, "Loja"], as_index=False)
                            .agg({
                                7: "sum",
                                11: "sum",
                                12: "sum",
                                6: "mean"
                            })
                        )
        
                        df_final.columns = ["Data", "Loja", "Fat.Total", "Serv/Tx", "Fat.Real", "Ticket"]
                        df_final["Mês"] = df_final["Data"].dt.strftime("%b").str.lower()
                        df_final["Ano"] = df_final["Data"].dt.year
                        df_final["Sistema"] = "3SCheckout"

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
                    df_final["Loja"] = df_final["Loja"].astype(str).str.strip().str.lower()
                    colunas_finais = [
                        "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
                        "Código Grupo Everest", "Fat.Total", "Serv/Tx", "Fat.Real",
                        "Ticket", "Mês", "Ano","Sistema"
                    ]
                    df_final = df_final[[c for c in colunas_finais if c in df_final.columns]]
                    
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
                out["Loja"] = df[loja_col].astype(str).str.strip().str.lower()  # ✅ Já em minúsculo
                out["Loja_norm"] = out["Loja"]  # Já está normalizado
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
           
            # 1️⃣ Se existir Sistema explícito, respeita
            if "Sistema" in df.columns:
                sistema_series = df["Sistema"].astype(str).str.strip()
                sistema_series = sistema_series[sistema_series != ""]
        
                if not sistema_series.empty:
                    return sistema_series.mode().iloc[0]
        
            # 2️⃣ Fallback por Grupo
            grp = df.get("Grupo", pd.Series([], dtype="object")).astype(str).str.lower()
        
            if grp.str.contains(r"\bkopp\b", regex=True).any():
                return "CISS"
        
            # 3️⃣ Fallback final
            return "Colibri"

        
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
            # ✅ Converte nome da loja para MAIÚSCULO
            df["Loja"] = df["Loja"].astype(str).str.strip().str.lower()
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
            
                # >>> SISTEMA (regra de precedência)
                if str(titulo_origem).lower() == "manuais":
                    # 🔒 Regra máxima: lançamento manual sempre vence
                    df_final["Sistema"] = "Lançamento manual"
                
                elif "Sistema" not in df_final.columns:
                    # 🔁 Fallback apenas se ninguém definiu
                    grp_norm = df_final.get("Grupo", "").astype(str).str.strip().str.lower()
                    df_final["Sistema"] = np.where(
                        grp_norm.str.contains(r"\bkopp\b", regex=True),
                        "CISS",
                        "Colibri"
                    )
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
                # ==== M e N com a data correta (YYYY-MM-DD) + Sistema
                fat_col = "Fat.Total" if "Fat.Total" in df_final.columns else "Fat Total"
                df_final[fat_col] = pd.to_numeric(df_final[fat_col], errors="coerce").fillna(0)
                loja_str = df_final["Loja"].astype(str)
                sistema_str = df_final["Sistema"].astype(str).str.strip()
                
                # M = yyyy-mm-dd + Fat.Total + Loja + Sistema
                df_final["M"] = (df_final["_Data_ymd"].fillna("") +
                                 df_final[fat_col].astype(str) +
                                 loja_str +
                                 sistema_str).str.strip()
                
                # N = yyyy-mm-dd + Codigo + Sistema
                df_final["N"] = (df_final["_Data_ymd"].fillna("") +
                                 df_final[cod_col].astype(str) +
                                 sistema_str).str.strip()
            
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
                # M/N finais (garantia) + Sistema
                fat_col = "Fat.Total" if "Fat.Total" in df_final.columns else "Fat Total"
                df_final["Data_Formatada"] = pd.to_datetime(
                    df_final["Data"], origin="1899-12-30", unit="D", errors="coerce"
                ).dt.strftime("%Y-%m-%d")
                df_final[fat_col] = pd.to_numeric(df_final[fat_col], errors="coerce").fillna(0)
                sistema_str = df_final["Sistema"].astype(str).str.strip()
                
                df_final["M"] = (df_final["Data_Formatada"].fillna("") +
                                 df_final[fat_col].astype(str) +
                                 df_final["Loja"].astype(str) +
                                 sistema_str).str.strip()
                
                if "Codigo Everest" in df_final.columns:
                    df_final["Codigo Everest"] = pd.to_numeric(df_final["Codigo Everest"], errors="coerce").fillna(0).astype(int).astype(str)
                elif "Código Everest" in df_final.columns:
                    df_final["Código Everest"] = pd.to_numeric(df_final["Código Everest"], errors="coerce").fillna(0).astype(int).astype(str)
                
                cod_col = "Codigo Everest" if "Codigo Everest" in df_final.columns else "Código Everest"
                df_final["N"] = (df_final["Data_Formatada"].fillna("") + 
                                 df_final[cod_col].astype(str) +
                                 sistema_str).str.strip()
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
                # ✅ ADICIONE AQUI (LOGO APÓS AS CONTAGENS):
                st.session_state._resumo_envio = {
                    "enviados": q_novos, 
                    "dup_m": q_dup_m, 
                    "sus_n": q_sus_n
                }
                
                # ✅ RENDERIZA A BARRA SEMPRE
                st.markdown(f"""
                    <div style="padding:10px; border-radius:8px; background-color:#e8f5e9; border:1px solid #c8e6c9; margin-bottom:15px;">
                        🟢 Enviados: <b>{q_novos}</b> | 
                        ❌ Duplicados: <b>{q_dup_m}</b> | 
                        🔴 Suspeitos: <b>{q_sus_n}</b>
                    </div>
                """, unsafe_allow_html=True)
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
        # ✅ Aceita TANTO df_final (upload) QUANTO resumo_3s (3S Checkout)
        df_final = st.session_state.get("df_final")
        resumo_3s = st.session_state.get("resumo_3s")
        
        if isinstance(df_final, pd.DataFrame) and not df_final.empty:
            df_sess = df_final
        elif isinstance(resumo_3s, pd.DataFrame) and not resumo_3s.empty:
            df_sess = resumo_3s
        else:
            df_sess = None
        
        has_df = df_sess is not None

    
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
                # ✅ Usa o DataFrame disponível (prioriza resumo_3s se existir)
                # --- Lógica correta para selecionar o DataFrame sem erro de ValueError ---
                df_para_enviar = None
                origem = ""
                
                # 1. Tenta pegar o do 3S Checkout primeiro
                if "resumo_3s" in st.session_state and st.session_state.resumo_3s is not None:
                    if isinstance(st.session_state.resumo_3s, pd.DataFrame) and not st.session_state.resumo_3s.empty:
                        df_para_enviar = st.session_state.resumo_3s
                        origem = "3s_checkout"
                
                # 2. Se não tiver 3S, tenta o do Upload Manual
                if df_para_enviar is None:
                    if "df_final" in st.session_state and st.session_state.df_final is not None:
                        if isinstance(st.session_state.df_final, pd.DataFrame) and not st.session_state.df_final.empty:
                            df_para_enviar = st.session_state.df_final
                            origem = "upload"
        
        # ✅ Mostra o resumo sempre que existir (fora do if enviar_auto)
        r = st.session_state.get("_resumo_envio")
        if r:
            st.success(
                f"🟢 Enviados: **{r['enviados']}**  |  ❌ Duplicados: **{r['dup_m']}**  |  🔴 Suspeitos: **{r['sus_n']}**"
            )
           
        # badge de pendência (só visual)
        if st.session_state.get("conf_pendente", False) and not st.session_state.get("conf_ok", False):
            st.caption(":red[Conferência pendente do sistema]")
        
        

    
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
        # === Handler do botão superior "Atualizar Sheets" ===
        if enviar_auto:
            if not has_df:
                st.error("Não há dados para enviar.")
            else:
                df_para_enviar = None
                origem = None
                
                # Prioriza 3S
                df_3s = st.session_state.get("resumo_3s", None)
                if isinstance(df_3s, pd.DataFrame) and not df_3s.empty:
                    df_para_enviar = df_3s
                    origem = "3s_checkout"
                
                # Fallback: upload manual
                if df_para_enviar is None:
                    df_up = st.session_state.get("df_final", None)
                    if isinstance(df_up, pd.DataFrame) and not df_up.empty:
                        df_para_enviar = df_up
                        origem = "upload"
                
                if df_para_enviar is None:
                    st.error("❌ Nenhum dado disponível para enviar (faça upload ou rode o 3S Checkout).")
                else:
                    ok = enviar_para_sheets(df_para_enviar.copy(), titulo_origem=origem)
                    # Se a função abriu o painel de conflitos, ela já dá st.rerun().
                    # Só mostramos "concluído" quando NÃO ficou em modo de conflitos.
                    # if ok and not st.session_state.get("modo_conflitos", False):
                    #     st.success("✅ Processo concluído.")
        
        # ✅ Mostra resumo se existir
        if "_resumo_envio" in st.session_state:
            r = st.session_state._resumo_envio
            # st.markdown(
            #     f"**Resumo:** 🟢 Enviados: **{r['enviados']}** &nbsp;|&nbsp; "
            #     f"❌ Duplicados (M): **{r['dup_m']}** &nbsp;|&nbsp; "
            #     f"🔴 Possíveis duplicados (N): **{r['sus_n']}**"
            # )
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
                        # M / N + Sistema
                        if col_M:
                            sis_val = str(d.get("Sistema", "") or "").strip()
                            row_out[col_M] = (data_ymd + str(fat_total) + loja + sis_val) if col_M else None
                        if col_N:
                            sis_val = str(d.get("Sistema", "") or "").strip()
                            row_out[col_N] = (data_ymd + cod_emp + sis_val) if col_N else None
                        # --- Sistema ---
                        if col_sis:
                            sis_val = str(d.get("Sistema", "") or "").strip()
                        
                            # 🔒 protege sistemas explícitos (ex: 3SCheckout)
                            if not sis_val:
                                grp = str(d.get("Grupo", "") or "").strip().lower()
                        
                                if grp:
                                    sis_val = "CISS" if re.search(r"\bkopp\b", grp, flags=re.I) else "Colibri"
                                else:
                                    # sem Grupo → não força Colibri
                                    sis_val = ""
                        
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
    # Aba 4 - Auditoria Everest
    # =======================================
    
    # app_streamlit_3fontes.py
    from datetime import date, timedelta, datetime
    import streamlit as st
    import pandas as pd
    import unicodedata, re
    import gspread
    from google.oauth2.service_account import Credentials
    
    # ---------------- Config ----------------
    SERVICE_ACCOUNT_FILE = "JsonImportadorEverest.json"  # ajuste se necessário
    
    # Fonte 1: Vendas Dárias (Fat Sistema Externo)
    ID_FAT_EXT = "1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU"
    ABA_FAT_EXT = "Fat Sistema Externo"
    COL_DATA_FAT_EXT = "Coluna A"
    COL_CODIGO_FAT_EXT = "Coluna D"
    COL_VALOR_FAT_EXT = "Coluna G"
    
    # Fonte 2: 507 venda
    ID_507 = "1tqmql1aL6M6A6yZ1QSOiifDjas5Bs_AziI7IkqwmAOM"
    ABA_507 = "Faturamento"
    COL_DATA_507 = "Coluna A"
    COL_CODIGO_507 = "Coluna B"
    COL_VALOR_507 = "Coluna D"
    
    # Fonte 3: Faturamento Meio Pagamento
    ID_MEIO = "1GSI291SEeeU9MtOWkGwsKGCGMi_xXMSiQnL_9GhXxfU"
    ABA_MEIO = "Faturamento Meio Pagamento"
    COL_DATA_MEIO = "Coluna A"
    COL_CODIGO_MEIO = "Coluna G"
    COL_VALOR_MEIO = "Coluna J"
    # ----------------------------------------
    
    st.set_page_config(page_title="Comparar 507 / Faturamento / Meio", layout="wide")
    st.title("Comparação: 507 / Faturamento (Vendas Dárias) / Meio Pagamento")
    
    # ---------- Helpers ----------
    def _norm_key(s):
        if not isinstance(s, str):
            return ""
        s = s.lower()
        s = unicodedata.normalize('NFKD', s)
        s = "".join(c for c in s if not unicodedata.combining(c))
        s = re.sub(r'\W+', '', s)
        return s
    
    def make_unique_headers(headers):
        seen = {}
        out = []
        for h in headers:
            h0 = (h or "").strip()
            if h0 == "": h0 = "col"
            cnt = seen.get(h0, 0) + 1
            seen[h0] = cnt
            out.append(h0 if cnt == 1 else f"{h0}_{cnt}")
        return out
    
    def safe_df_from_sheet(vals):
        """Converte get_all_values() -> DataFrame usando primeira linha como cabeçalho."""
        if not vals or len(vals) < 1:
            return pd.DataFrame()
        ncols = max(len(r) for r in vals)
        headers_raw = [ (vals[0][i] if i < len(vals[0]) else "") for i in range(ncols) ]
        headers = make_unique_headers([h.strip() for h in headers_raw])
        rows = [ [(r[i] if i < len(r) else "") for i in range(ncols)] for r in vals[1:] ]
        return pd.DataFrame(rows, columns=headers)
    
    def tratar_valor(valor):
        try:
            if pd.isna(valor) or str(valor).strip() == "":
                return None
            s = str(valor).strip().replace("R$", "").replace("\xa0", "").strip()
            # tratar milhar/decimal BR/US
            if s.count(",") == 1 and s.count(".") >= 1:
                s = s.replace(".", "").replace(",", ".")
            else:
                s = s.replace(",", ".")
            return float(s)
        except:
            return None
    
    def normalize_codigo(x):
        try:
            if pd.isna(x):
                return ""
            sx = str(x).strip().replace("\xa0", "")
            if re.fullmatch(r"\d+(\.0+)?", sx):
                return str(int(float(sx)))
            return sx
        except:
            return str(x)
    
    def fmt_brl(v):
        try:
            v = float(v)
            return f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except:
            return "R$ 0,00"
    
    def find_column_by_candidates(df, candidates, fallback_idx=None):
        if df is None or df.empty:
            return None
        cmap = {_norm_key(c): c for c in df.columns}
        for cand in candidates:
            k = _norm_key(cand)
            if k in cmap:
                return cmap[k]
        if fallback_idx is not None and 0 <= fallback_idx < len(df.columns):
            return df.columns[fallback_idx]
        return None
    
    def detect_most_numeric_column(df):
        best = None
        best_count = -1
        for col in df.columns:
            sample = df[col].dropna().astype(str)
            if sample.empty:
                continue
            count = 0
            for v in sample.head(200):
                if tratar_valor(v) is not None:
                    count += 1
            if count > best_count:
                best_count = count
                best = col
        return best
    
    def parse_date_col(series):
        # tenta interpretar data/serial
        def _parse(v):
            if pd.isna(v) or str(v).strip() == "":
                return pd.NaT
            # tenta float (excel serial)
            try:
                f = float(v)
                if f > 31:
                    try:
                        return (datetime(1899,12,30) + pd.Timedelta(days=int(f))).date()
                    except:
                        pass
            except:
                pass
            try:
                dt = pd.to_datetime(v, dayfirst=True, errors="coerce")
                if pd.isna(dt):
                    dt = pd.to_datetime(v, errors="coerce")
                if pd.isna(dt):
                    return pd.NaT
                return dt.date()
            except:
                return pd.NaT
        return series.apply(_parse)
    
    def detect_name_column(cols):
        for c in cols:
            k = _norm_key(c)
            if k in ("nome","nomefantasia","nomefantasia(razao)","razao","cliente","empresa","nome_fantasia"):
                return c
        return None
    
    # ---------- Autenticação Google ----------
    def gspread_auth():
        scopes = ["https://www.googleapis.com/auth/spreadsheets.readonly","https://www.googleapis.com/auth/drive.readonly"]
        creds = Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=scopes)
        return gspread.authorize(creds)
    
    # ---------- Leitura e normalização por fonte ----------
    def read_and_normalize(gc, ssid, aba, col_data_ref, col_codigo_ref, col_valor_ref, sample_n=None):
        sh = gc.open_by_key(ssid)
        try:
            ws = sh.worksheet(aba)
        except Exception:
            # tentativa alternativa por título case-insensitive
            ws = next((w for w in sh.worksheets() if w.title.strip().lower() == aba.strip().lower()), None)
            if ws is None:
                raise
        values = ws.get_all_values()
        df_raw = safe_df_from_sheet(values)
        if df_raw.empty:
            return pd.DataFrame(columns=["Data","Codigo","Nome","Valor"])
        if sample_n:
            df_raw = df_raw.head(sample_n)
    
        date_col = find_column_by_candidates(df_raw, [col_data_ref, "Data","Data do Movimento","data"], fallback_idx=0)
        code_col = find_column_by_candidates(df_raw, [col_codigo_ref, "Codigo","Código da Empresa","Código","codigo","empresa"], fallback_idx=1)
        val_col  = find_column_by_candidates(df_raw, [col_valor_ref, "Valor","Valor Total","Valor (R$)","V. Total"], fallback_idx=None)
        if val_col is None:
            val_col = detect_most_numeric_column(df_raw)
    
        name_col = detect_name_column(df_raw.columns)
    
        # construir df normalizado
        out = pd.DataFrame()
        out["Data"] = parse_date_col(df_raw[date_col]) if date_col else pd.NaT
        out["Codigo"] = df_raw[code_col].apply(normalize_codigo) if code_col else ""
        out["Nome"] = df_raw[name_col] if name_col else ""
        out["Valor"] = df_raw[val_col].apply(tratar_valor).fillna(0.0) if val_col else 0.0
    
        # remover linhas sem data
        out = out.dropna(subset=["Data"])
        return out[["Data","Codigo","Nome","Valor"]]
    
    # ---------- UI e execução ----------
    preview_mode = st.sidebar.selectbox("Modo", ["Prévia (100 linhas por fonte)", "Completo (tudo)"])
    sample_n = 100 if preview_mode.startswith("Prévia") else None
    
    st.sidebar.write("Arquivo de credenciais:", SERVICE_ACCOUNT_FILE)
    if st.sidebar.button("Executar comparação"):
        run = True
    else:
        run = False
    
    if run:
        try:
            gc = gspread_auth()
        except Exception as e:
            st.error(f"Erro na autenticação Google: {e}")
            st.stop()
    
        # Ler as 3 fontes
        try:
            df_fat_ext = read_and_normalize(gc, ID_FAT_EXT, ABA_FAT_EXT, COL_DATA_FAT_EXT, COL_CODIGO_FAT_EXT, COL_VALOR_FAT_EXT, sample_n=sample_n)
        except Exception as e:
            st.warning(f"Erro lendo Vendas Dárias (Fat Sistema Externo): {e}")
            df_fat_ext = pd.DataFrame(columns=["Data","Codigo","Nome","Valor"])
    
        try:
            df_507 = read_and_normalize(gc, ID_507, ABA_507, COL_DATA_507, COL_CODIGO_507, COL_VALOR_507, sample_n=sample_n)
        except Exception as e:
            st.warning(f"Erro lendo 507 venda: {e}")
            df_507 = pd.DataFrame(columns=["Data","Codigo","Nome","Valor"])
    
        try:
            df_meio = read_and_normalize(gc, ID_MEIO, ABA_MEIO, COL_DATA_MEIO, COL_CODIGO_MEIO, COL_VALOR_MEIO, sample_n=sample_n)
        except Exception as e:
            st.warning(f"Erro lendo Faturamento Meio Pagamento: {e}")
            df_meio = pd.DataFrame(columns=["Data","Codigo","Nome","Valor"])
    
        st.write("Exemplo - 507 (primeiras linhas):")
        st.dataframe(df_507.head(10))
        st.write("Exemplo - Vendas Dárias (primeiras linhas):")
        st.dataframe(df_fat_ext.head(10))
        st.write("Exemplo - Meio (primeiras linhas):")
        st.dataframe(df_meio.head(10))
    
        # Agregar por Data + Codigo
        agg_507 = df_507.groupby(["Data","Codigo"], as_index=False).agg({"Valor":"sum","Nome":"first"}).rename(columns={"Valor":"Valor_507","Nome":"Nome_507"})
        agg_fat = df_fat_ext.groupby(["Data","Codigo"], as_index=False).agg({"Valor":"sum","Nome":"first"}).rename(columns={"Valor":"Valor_Faturamento","Nome":"Nome_Faturamento"})
        agg_meio = df_meio.groupby(["Data","Codigo"], as_index=False).agg({"Valor":"sum","Nome":"first"}).rename(columns={"Valor":"Valor_Meio","Nome":"Nome_Meio"})
    
        # união das chaves (Data + Codigo)
        keys = pd.concat([agg_507[["Data","Codigo"]], agg_fat[["Data","Codigo"]], agg_meio[["Data","Codigo"]]]).drop_duplicates().reset_index(drop=True)
    
        # Data range picker baseado nos dados disponíveis
        all_dates = pd.to_datetime(keys["Data"], errors="coerce").dropna().dt.date
        if all_dates.empty:
            st.error("Nenhuma data válida encontrada nas fontes.")
            st.stop()
        data_min = all_dates.min()
        data_max = (date.today() - timedelta(days=1))
        col1, col2 = st.columns(2)
        with col1:
            start_date = st.date_input("Data Inicial:", value=data_min, min_value=data_min, max_value=data_max)
        with col2:
            end_date = st.date_input("Data Final:", value=data_max, min_value=data_min, max_value=data_max)
    
        # filtrar chaves
        mask = (keys["Data"] >= pd.to_datetime(start_date).date()) & (keys["Data"] <= pd.to_datetime(end_date).date())
        keys = keys.loc[mask].reset_index(drop=True)
    
        # merge seguro
        res = keys.merge(agg_507, on=["Data","Codigo"], how="left")
        res = res.merge(agg_fat, on=["Data","Codigo"], how="left")
        res = res.merge(agg_meio, on=["Data","Codigo"], how="left")
    
        # preencher nulos e coerção
        res["Valor_507"] = pd.to_numeric(res["Valor_507"].fillna(0.0), errors="coerce").fillna(0.0)
        res["Valor_Faturamento"] = pd.to_numeric(res["Valor_Faturamento"].fillna(0.0), errors="coerce").fillna(0.0)
        res["Valor_Meio"] = pd.to_numeric(res["Valor_Meio"].fillna(0.0), errors="coerce").fillna(0.0)
    
        # Nome Fantasia final: preferir 507, depois Faturamento, depois Meio
        def coalesce_nome(row):
            for c in ("Nome_507","Nome_Faturamento","Nome_Meio"):
                v = row.get(c)
                if pd.notna(v) and str(v).strip() != "":
                    return v
            return "Não Identificado"
        res["Nome Fantasia"] = res.apply(coalesce_nome, axis=1)
    
        # diferenças numéricas
        res["Diff_507_Faturamento"] = (res["Valor_507"] - res["Valor_Faturamento"]).round(2)
        res["Diff_507_Meio"] = (res["Valor_507"] - res["Valor_Meio"]).round(2)
    
        # formatação para exibição
        disp = res.copy()
        disp["Data"] = pd.to_datetime(disp["Data"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
        disp["Valor (507)"] = disp["Valor_507"].apply(fmt_brl)
        disp["Valor (Faturamento)"] = disp["Valor_Faturamento"].apply(fmt_brl)
        disp["Valor (Meio)"] = disp["Valor_Meio"].apply(fmt_brl)
        disp["Diferença (507 - Faturamento)"] = disp["Diff_507_Faturamento"].apply(fmt_brl)
        disp["Diferença (507 - Meio)"] = disp["Diff_507_Meio"].apply(fmt_brl)
    
        # ordenar por Data e código (tentando numérico)
        def try_numeric_sort_key(s):
            try:
                return int(s)
            except:
                try:
                    return float(s)
                except:
                    return str(s)
        disp["__cod_key"] = disp["Codigo"].apply(try_numeric_sort_key)
        disp = disp.sort_values(["Data","__cod_key"], ascending=[True,True]).drop(columns="__cod_key").reset_index(drop=True)
    
        final = disp[["Data","Codigo","Nome Fantasia","Valor (507)","Valor (Faturamento)","Valor (Meio)","Diferença (507 - Faturamento)","Diferença (507 - Meio)"]].copy()
        st.subheader("📊 Comparativo (507 vs Faturamento vs Meio)")
        st.dataframe(final, use_container_width=True, height=700)
    
        st.markdown("### Dados numéricos (para conferência)")
        st.dataframe(res[["Data","Codigo","Nome Fantasia","Valor_507","Valor_Faturamento","Valor_Meio","Diff_507_Faturamento","Diff_507_Meio"]].head(200))
    
    else:
        st.write("Clique em 'Executar comparação' na barra lateral para iniciar.")
   
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
    
        # Abre a planilha pelo ID (substitua pelo seu ID)
        sh = _gc.open_by_key("1GSI291SEeeU9MtOWkGwsKGCGMi_xXMSiQnL_9GhXxfU")

        # 1. Abre a planilha antiga para pegar o Sistema Externo
        sh_vendas = _gc.open("Vendas diarias")
        ws_ext = sh_vendas.worksheet("Fat Sistema Externo")

        # 2. Abre a NOVA planilha para pegar o Meio de Pagamento
        sh_meio_pagto = _gc.open_by_key("1GSI291SEeeU9MtOWkGwsKGCGMi_xXMSiQnL_9GhXxfU")
        ws_mp = sh_meio_pagto.worksheet("Faturamento Meio Pagamento")

        del _gc
    
        # ---------------- Leitura ----------------
        #ws_ext = sh.worksheet("Fat Sistema Externo")
        vals_ext = ws_ext.get_all_values()
        if len(vals_ext) < 2:
            st.error("Aba 'Fat Sistema Externo' vazia.")
            st.stop()
        header_ext, rows_ext = vals_ext[0], vals_ext[1:]
        df_ext = pd.DataFrame(rows_ext, columns=[c.strip() for c in header_ext])
    
        #ws_mp = sh.worksheet("Faturamento Meio Pagamento")
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
        tabela["Sistema_View"] = tabela["Sistema"].str.lower()
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
