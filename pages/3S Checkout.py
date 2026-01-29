import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime, timedelta
import json
import psycopg2

st.set_page_config(page_title="Atualização 3S", layout="wide")

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
    """Busca dados do 3S Checkout - Apenas Meio de Pagamento Agrupado"""
    conn = get_db_conn()
    try:
        ontem = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        
        # QUERY 1: order_picture (BASE)
        query_op = """
            SELECT 
                order_picture_id,
                store_code, 
                business_dt,
                custom_properties
            FROM public.order_picture
            WHERE business_dt >= '2024-12-01'
              AND business_dt <= %s
              AND store_code NOT IN ('0000', '0001', '9999')
              AND state_id = 5
        """
        df_op = pd.read_sql(query_op, conn, params=(ontem,))
        
        # Filtrar VOID_TYPE na base de pedidos
        props = df_op['custom_properties'].apply(parse_props)
        df_op['VOID_TYPE'] = props.apply(lambda x: x.get('VOID_TYPE'))
        df_op = df_op[df_op['VOID_TYPE'].isna() | (df_op['VOID_TYPE'] == "") | (df_op['VOID_TYPE'] == 0)].copy()

        # QUERY 2: order_picture_tender (PAGAMENTOS)
        query_tender = """
            SELECT 
                order_picture_id,
                tender_amount,
                change_amount,
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
        
        if df_tender.empty:
            return pd.DataFrame(), None, 0

        # Processar Pagamentos
        df_tender["tender_amount"] = pd.to_numeric(df_tender["tender_amount"], errors="coerce").fillna(0)
        df_tender["change_amount"] = pd.to_numeric(df_tender["change_amount"], errors="coerce").fillna(0)
        
        tender_props = df_tender['details'].apply(parse_props)
        df_tender['Meio de Pagamento'] = tender_props.apply(lambda x: x.get("tenderDescr") if isinstance(x, dict) else None)
        df_tender['tender_tip_amount'] = pd.to_numeric(
            tender_props.apply(lambda x: x.get('tipAmount', 0) if isinstance(x, dict) else 0), 
            errors='coerce'
        ).fillna(0)
        
        # Cálculo Líquido (Descontando Troco)
        df_tender["Fat.Real"] = (df_tender["tender_amount"] - df_tender["change_amount"]).clip(lower=0)
        df_tender["Fat.Total"] = df_tender["Fat.Real"] + df_tender["tender_tip_amount"]
        
        # Merge com dados do pedido
        df_tender = df_tender.merge(df_op[['order_picture_id', 'store_code', 'business_dt']], on='order_picture_id', how='inner')
        
        # Formatação de Datas e Tradução
        dias_traducao = {
            "Monday": "segunda-feira", "Tuesday": "terça-feira", "Wednesday": "quarta-feira",
            "Thursday": "quinta-feira", "Friday": "sexta-feira", "Saturday": "sábado", "Sunday": "domingo"
        }
        df_tender['Data'] = pd.to_datetime(df_tender['business_dt']).dt.strftime('%d/%m/%Y')
        df_tender['Dia da Semana'] = pd.to_datetime(df_tender['business_dt']).dt.day_name().map(dias_traducao)
        
        # PROCV Tabela Empresa
        df_empresa["Código Everest"] = df_empresa["Código Everest"].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0")
        df_tender['Código Everest'] = df_tender['store_code'].astype(str).str.lstrip('0').str.strip()
        
        df_tender = pd.merge(df_tender, df_empresa[["Código Everest", "Loja", "Grupo", "Código Grupo Everest"]], on="Código Everest", how="left")
        df_tender["Loja"] = df_tender["Loja"].astype(str).str.strip().str.lower()
        
        # Mês e Ano
        meses = {"jan": "jan", "feb": "fev", "mar": "mar", "apr": "abr", "may": "mai", "jun": "jun",
                 "jul": "jul", "aug": "ago", "sep": "set", "oct": "out", "nov": "nov", "dec": "dez"}
        df_tender['Mês'] = pd.to_datetime(df_tender['business_dt']).dt.strftime('%b').str.lower().map(meses)
        df_tender['Ano'] = pd.to_datetime(df_tender['business_dt']).dt.year
        df_tender['Sistema'] = '3SCheckout'
        
        # AGRUPAMENTO FINAL
        resumo_pagamento = df_tender.groupby(
            ['Data', 'Dia da Semana', 'Meio de Pagamento', 'Loja', 'Código Everest', 'Grupo', 'Código Grupo Everest', 'Mês', 'Ano', 'Sistema'],
            dropna=False
        ).agg({'Fat.Total': 'sum'}).reset_index()
        
        resumo_pagamento['Fat.Total'] = resumo_pagamento['Fat.Total'].round(2)
        
        # Ordenação de Colunas
        colunas_finais = [
            "Data", "Dia da Semana", "Meio de Pagamento", "Loja", "Código Everest", 
            "Grupo", "Código Grupo Everest", "Fat.Total", "Mês", "Ano", "Sistema"
        ]
        resumo_pagamento = resumo_pagamento[colunas_finais]
        
        # Ordenação Cronológica
        resumo_pagamento['Data_Sort'] = pd.to_datetime(resumo_pagamento['Data'], format='%d/%m/%Y')
        resumo_pagamento = resumo_pagamento.sort_values(by=['Data_Sort', 'Loja', 'Meio de Pagamento']).drop(columns='Data_Sort')
        
        return resumo_pagamento, None, len(df_op)
        
    except Exception as e:
        return None, str(e), 0
    finally:
        conn.close()

# ================================
# 6. Interface do Streamlit
# ================================
st.title("🔄 Atualização 3S")

if "resumo_pagamento" not in st.session_state:
    st.session_state["resumo_pagamento"] = None

if st.button("🔄 Atualizar 3S Checkout", type="primary", use_container_width=True):
    with st.spinner("Buscando dados do banco..."):
        resumo_pagamento, erro_3s, total_pedidos = buscar_dados_3s_checkout()
    
    if erro_3s:
        st.error(f"❌ Erro: {erro_3s}")
    elif resumo_pagamento is not None and not resumo_pagamento.empty:
        st.session_state["resumo_pagamento"] = resumo_pagamento
        st.success(f"✅ {total_pedidos} pedidos processados!")
    else:
        st.warning("⚠️ Nenhum dado encontrado.")

if st.session_state["resumo_pagamento"] is not None:
    df = st.session_state["resumo_pagamento"]
    
    # Alerta de Lojas não localizadas
    lojas_nulas = df[df["Loja"].isna() | (df["Loja"] == "nan")]["Código Everest"].unique()
    if len(lojas_nulas) > 0:
        st.warning(f"⚠️ Lojas não localizadas para os códigos: {', '.join(lojas_nulas)}")

    # Resumo Financeiro no Topo
    total_geral = df["Fat.Total"].sum()
    st.metric("💰 Faturamento Total (Líquido de Troco)", f"R$ {total_geral:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    
    # Download
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Meio de Pagamento', index=False)
    output.seek(0)
    
    st.download_button(
        label="📥 Baixar Excel (Aba Única: Meio de Pagamento)",
        data=output.getvalue(),
        file_name=f"3s_pagamentos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
   
