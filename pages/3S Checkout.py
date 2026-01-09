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
# 4. Função de busca e processamento
# ================================
def buscar_dados_3s_checkout():
    """Busca dados do 3S Checkout direto do banco e processa"""
    conn = get_db_conn()
    try:
        # ✅ CALCULA A DATA DE ONTEM
        ontem = (datetime.now() - timedelta(days=1)).strftime('%Y-%m-%d')
        
        # ✅ FILTRO SQL
        query = """
            SELECT store_code, business_dt, total_gross, custom_properties, order_code, state_id
            FROM public.order_picture
            WHERE business_dt >= '2024-12-01'
              AND business_dt <= %s
              AND store_code NOT IN ('0000', '0001', '9999')
              AND state_id = 5
        """
        df = pd.read_sql(query, conn, params=(ontem,))
        
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
        
        # 11. ✅ Buscar informações da Tabela Empresa
        df_empresa["Código Everest"] = (
            df_empresa["Código Everest"]
            .astype(str)
            .str.replace(r"\D", "", regex=True)
            .str.lstrip("0")
        )
        resumo["Código Everest"] = resumo["Código Everest"].astype(str).str.strip()
        
        resumo = pd.merge(resumo, df_empresa[["Código Everest", "Loja", "Grupo", "Código Grupo Everest"]], 
                         on="Código Everest", how="left")
        
        # ✅ Converte nome da loja para minúsculo
        resumo["Loja"] = resumo["Loja"].astype(str).str.strip().str.lower()
        
        # 12. Adicionar colunas adicionais
        resumo['Ticket'] = 0
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
# 5. Interface do Streamlit
# ================================
st.title("🔄 Atualização 3S Checkout")

if st.button("🔄 Atualizar 3S Checkout", type="primary", use_container_width=True):
    with st.spinner("Buscando dados do banco..."):
        resumo_3s, erro_3s, total_registros = buscar_dados_3s_checkout()
        
        if erro_3s:
            st.error(f"❌ Erro ao buscar dados: {erro_3s}")
        elif resumo_3s is not None and not resumo_3s.empty:
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
        else:
            st.warning("⚠️ Nenhum dado encontrado para o período.")
