import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
import json
import psycopg2

st.set_page_config(page_title="Atualização 3S Checkout", layout="wide")

# ================================
# 1. Certificado AWS
# ================================
CERT_PATH = "aws-us-east-2-bundle.pem"

if "cert_written" not in st.session_state:
    with open(CERT_PATH, "w", encoding="utf-8") as f:
        f.write(st.secrets["certs"]["aws_rds_us_east_2"])
    st.session_state["cert_written"] = True

# ================================
# 2. Conexão com PostgreSQL
# ================================
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

# ================================
# 3. Conexão com Google Sheets (Tabela Empresa)
# ================================
import gspread
from oauth2client.service_account import ServiceAccountCredentials

scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Vendas diarias")

# Carrega Tabela Empresa
aba_empresa = planilha_empresa.worksheet("Tabela Empresa")
valores_empresa = aba_empresa.get_all_values()

if len(valores_empresa) > 1:
    df_empresa = pd.DataFrame(valores_empresa[1:], columns=valores_empresa[0])
    df_empresa.columns = df_empresa.columns.str.strip()
    
    if "Loja" in df_empresa.columns:
        df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.lower().str.strip()
else:
    df_empresa = pd.DataFrame()

# ================================
# 4. Função auxiliar para parse JSON
# ================================
def parse_props(x):
    if pd.isna(x): 
        return {}
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

# ================================
# 5. Função de busca e processamento
# ================================
def buscar_dados_3s_checkout():
    """Busca dados do 3S Checkout - MODELO EXCEL (sem duplicação)"""
    conn = get_db_conn()
    try:
        ontem = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        
        # ========================================
        # QUERY 1: order_picture (BASE - PEDIDOS)
        # ========================================
        query_op = """
            SELECT 
                order_picture_id,
                store_code, 
                business_dt, 
                total_gross, 
                custom_properties,
                state_id
            FROM public.order_picture
            WHERE business_dt >= '2024-12-01'
              AND business_dt <= %s
              AND store_code NOT IN ('0000', '0001', '9999')
              AND state_id = 5
        """
        df_op = pd.read_sql(query_op, conn, params=(ontem,))
        
        # ========================================
        # QUERY 2: order_picture_tender (PAGAMENTOS)
        # ========================================
        query_tender = """
            SELECT 
                order_picture_id,
                tender_amount,
                details
            FROM public.order_picture_tender
            WHERE order_picture_id IN (
                SELECT order_picture_id 
                FROM public.order_picture
                WHERE business_dt >= '2024-12-01'
                  AND business_dt <= %s
                  AND store_code NOT IN ('0000', '0001', '9999')
                  AND state_id = 5
            )
        """
        df_tender = pd.read_sql(query_tender, conn, params=(ontem,))
        
        # ========================================
        # PROCESSAR order_picture (ABA VENDAS)
        # ========================================
        df_op['business_dt'] = pd.to_datetime(df_op['business_dt'], errors='coerce')
        df_op['store_code'] = df_op['store_code'].astype(str).str.lstrip('0')
        
        # Extrair TIP_AMOUNT e VOID_TYPE
        props = df_op['custom_properties'].apply(parse_props)
        df_op['TIP_AMOUNT'] = pd.to_numeric(props.apply(lambda x: x.get('TIP_AMOUNT')), errors='coerce').fillna(0)
        df_op['VOID_TYPE'] = props.apply(lambda x: x.get('VOID_TYPE'))
        
        # Filtrar VOID_TYPE
        df_op = df_op[df_op['VOID_TYPE'].isna() | (df_op['VOID_TYPE'] == "") | (df_op['VOID_TYPE'] == 0)].copy()
        
        # Criar colunas da aba VENDAS
        resumo_vendas = df_op.copy()
        resumo_vendas['Código Everest'] = resumo_vendas['store_code']
        resumo_vendas['Data'] = resumo_vendas['business_dt'].dt.strftime('%d/%m/%Y')
        resumo_vendas['Fat.Real'] = resumo_vendas['total_gross']
        resumo_vendas['Serv/Tx'] = resumo_vendas['TIP_AMOUNT']
        resumo_vendas['Fat.Total'] = resumo_vendas['Fat.Real'] + resumo_vendas['Serv/Tx']
        
        # Dia da Semana
        dias_traducao = {
            "Monday": "segunda-feira", "Tuesday": "terça-feira", "Wednesday": "quarta-feira",
            "Thursday": "quinta-feira", "Friday": "sexta-feira", "Saturday": "sábado", "Sunday": "domingo"
        }
        resumo_vendas['Dia da Semana'] = resumo_vendas['business_dt'].dt.day_name().map(dias_traducao)
        
        # PROCV com Tabela Empresa
        df_empresa["Código Everest"] = (
            df_empresa["Código Everest"]
            .astype(str)
            .str.replace(r"\D", "", regex=True)
            .str.lstrip("0")
        )
        resumo_vendas["Código Everest"] = resumo_vendas["Código Everest"].astype(str).str.strip()
        
        resumo_vendas = pd.merge(
            resumo_vendas, 
            df_empresa[["Código Everest", "Loja", "Grupo", "Código Grupo Everest"]], 
            on="Código Everest", 
            how="left"
        )
        
        resumo_vendas["Loja"] = resumo_vendas["Loja"].astype(str).str.strip().str.lower()
        
        # Colunas adicionais
        resumo_vendas['Ticket'] = 0
        resumo_vendas['Mês'] = resumo_vendas['business_dt'].dt.strftime('%b').str.lower()
        
        meses = {"jan": "jan", "feb": "fev", "mar": "mar", "apr": "abr", "may": "mai", "jun": "jun",
                 "jul": "jul", "aug": "ago", "sep": "set", "oct": "out", "nov": "nov", "dec": "dez"}
        resumo_vendas["Mês"] = resumo_vendas["Mês"].map(meses)
        resumo_vendas['Ano'] = resumo_vendas['business_dt'].dt.year
        resumo_vendas['Sistema'] = '3SCheckout'
        
        # Selecionar e ordenar colunas
        colunas_vendas = [
            "order_picture_id", "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
            "Código Grupo Everest", "Fat.Total", "Serv/Tx", "Fat.Real",
            "Ticket", "Mês", "Ano", "Sistema"
        ]
        resumo_vendas = resumo_vendas[[c for c in colunas_vendas if c in resumo_vendas.columns]]
        
        # Arredondar valores
        for col in ["Fat.Total", "Serv/Tx", "Fat.Real", "Ticket"]:
            if col in resumo_vendas.columns:
                resumo_vendas[col] = resumo_vendas[col].round(2)
        
        # Ordenar
        resumo_vendas['Data_Ordenada'] = pd.to_datetime(resumo_vendas['Data'], format='%d/%m/%Y')
        resumo_vendas = resumo_vendas.sort_values(by=['Data_Ordenada', 'Loja']).drop(columns='Data_Ordenada')
        
        # ========================================
        # PROCESSAR order_picture_tender (ABA PAGAMENTO)
        # ========================================
        if not df_tender.empty:
            # Extrair dados do JSON details
            tender_props = df_tender['details'].apply(parse_props)
            
            df_tender['Meio de Pagamento'] = tender_props.apply(
                lambda x: x.get("tenderDescr") if isinstance(x, dict) else None
            )
            df_tender['tender_tip_amount'] = pd.to_numeric(
                tender_props.apply(lambda x: x.get('tipAmount', 0) if isinstance(x, dict) else 0), 
                errors='coerce'
            ).fillna(0)
            
            # Criar colunas da aba PAGAMENTO
            resumo_pagamento = df_tender.copy()
            resumo_pagamento['Fat.Real'] = pd.to_numeric(resumo_pagamento['tender_amount'], errors='coerce').fillna(0)
            resumo_pagamento['Serv/Tx'] = resumo_pagamento['tender_tip_amount']
            resumo_pagamento['Fat.Total'] = resumo_pagamento['Fat.Real'] + resumo_pagamento['Serv/Tx']
            
            # ✅ PROCV - Buscar dados do pedido (store_code, data, etc)
            resumo_pagamento = resumo_pagamento.merge(
                df_op[['order_picture_id', 'store_code', 'business_dt']],
                on='order_picture_id',
                how='left'
            )
            
            resumo_pagamento['Código Everest'] = resumo_pagamento['store_code'].astype(str).str.lstrip('0').str.strip()
            resumo_pagamento['Data'] = pd.to_datetime(resumo_pagamento['business_dt'], errors='coerce').dt.strftime('%d/%m/%Y')
            resumo_pagamento['Dia da Semana'] = pd.to_datetime(resumo_pagamento['business_dt'], errors='coerce').dt.day_name().map(dias_traducao)
            
            # PROCV com Tabela Empresa
            resumo_pagamento = pd.merge(
                resumo_pagamento, 
                df_empresa[["Código Everest", "Loja", "Grupo", "Código Grupo Everest"]], 
                on="Código Everest", 
                how="left"
            )
            
            resumo_pagamento["Loja"] = resumo_pagamento["Loja"].astype(str).str.strip().str.lower()
            
            # Colunas adicionais
            resumo_pagamento['Mês'] = pd.to_datetime(resumo_pagamento['business_dt'], errors='coerce').dt.strftime('%b').str.lower()
            resumo_pagamento["Mês"] = resumo_pagamento["Mês"].map(meses)
            resumo_pagamento['Ano'] = pd.to_datetime(resumo_pagamento['business_dt'], errors='coerce').dt.year
            resumo_pagamento['Sistema'] = '3SCheckout'
            
            # Selecionar e ordenar colunas
            colunas_pagamento = [
                "order_picture_id", "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
                "Código Grupo Everest", "Meio de Pagamento", "Fat.Total", "Serv/Tx", "Fat.Real",
                "Mês", "Ano", "Sistema"
            ]
            resumo_pagamento = resumo_pagamento[[c for c in colunas_pagamento if c in resumo_pagamento.columns]]
            
            # Arredondar valores
            for col in ["Fat.Total", "Serv/Tx", "Fat.Real"]:
                if col in resumo_pagamento.columns:
                    resumo_pagamento[col] = resumo_pagamento[col].round(2)
            
            # Ordenar
            resumo_pagamento['Data_Ordenada'] = pd.to_datetime(resumo_pagamento['Data'], format='%d/%m/%Y', errors='coerce')
            resumo_pagamento = resumo_pagamento.sort_values(by=['Data_Ordenada', 'Loja', 'Meio de Pagamento']).drop(columns='Data_Ordenada')
        else:
            resumo_pagamento = pd.DataFrame()
        
        return resumo_vendas, resumo_pagamento, None, len(df_op)
        
    except Exception as e:
        return None, None, str(e), 0
    finally:
        conn.close()

# ================================
# 6. Interface do Streamlit
# ================================
st.title("🔄 Atualização 3S Checkout")

# Inicializar session_state
if "resumo_vendas" not in st.session_state:
    st.session_state["resumo_vendas"] = None
if "resumo_pagamento" not in st.session_state:
    st.session_state["resumo_pagamento"] = None
if "total_registros" not in st.session_state:
    st.session_state["total_registros"] = 0

if st.button("🔄 Atualizar 3S Checkout", type="primary", use_container_width=True):
    with st.spinner("Buscando dados do banco..."):
        resumo_vendas, resumo_pagamento, erro_3s, total_registros = buscar_dados_3s_checkout()
    
    if erro_3s:
        st.error(f"❌ Erro ao buscar dados: {erro_3s}")
    elif resumo_vendas is not None and not resumo_vendas.empty:
        st.session_state["resumo_vendas"] = resumo_vendas
        st.session_state["resumo_pagamento"] = resumo_pagamento
        st.session_state["total_registros"] = total_registros
        st.success(f"✅ {total_registros} registros processados com sucesso!")
    else:
        st.warning("⚠️ Nenhum dado encontrado para o período.")

# ✅ Se já tiver dados no estado, mostra informações e botão de download
if st.session_state["resumo_vendas"] is not None and not st.session_state["resumo_vendas"].empty:
    resumo_vendas = st.session_state["resumo_vendas"]
    resumo_pagamento = st.session_state["resumo_pagamento"]
    total_registros = st.session_state["total_registros"]
    
    # Verificar empresas não localizadas
    empresas_nao_localizadas = resumo_vendas[resumo_vendas["Loja"].isna()]["Código Everest"].unique()
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
    datas_validas = pd.to_datetime(resumo_vendas["Data"], format="%d/%m/%Y", errors='coerce').dropna()
    if not datas_validas.empty:
        data_inicial = datas_validas.min().strftime("%d/%m/%Y")
        data_final_str = datas_validas.max().strftime("%d/%m/%Y")
        valor_total = resumo_vendas["Fat.Total"].sum()
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
    
    # ✅ Gera Excel com 2 abas
    try:
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            resumo_vendas.to_excel(writer, sheet_name='Faturamento Servico', index=False)
            
            if resumo_pagamento is not None and not resumo_pagamento.empty:
                resumo_pagamento.to_excel(writer, sheet_name='Meio de Pagamento', index=False)
            else:
                pd.DataFrame().to_excel(writer, sheet_name='Meio de Pagamento', index=False)
        
        output.seek(0)
        
        st.download_button(
            label="📥 Baixar Excel 3S Checkout",
            data=output.getvalue(),
            file_name=f"3s_checkout_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    except Exception as e:
        st.error("❌ Erro ao gerar o Excel:")
        st.exception(e)
