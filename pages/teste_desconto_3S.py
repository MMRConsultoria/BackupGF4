import os
import pandas as pd
import psycopg2
import re
from datetime import datetime, timedelta

# ESTE CÓDIGO NÃO POSSUI CONEXÃO COM GOOGLE SHEETS
# ELE APENAS GERA UM ARQUIVO EXCEL LOCAL

def _parse_money(x):
    if pd.isna(x): return 0.0
    s = str(x).replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
    try:
        return float(re.sub(r"[^\d.]", "", s))
    except:
        return 0.0

def gerar_apenas_excel_local():
    # 1. Conexão com o Banco (Postgres)
    try:
        db = st.secrets["db"] # Se estiver no Streamlit
        conn = psycopg2.connect(
            host=db["host"],
            database=db["database"],
            user=db["user"],
            password=db["password"],
            port=db.get("port", 5432)
        )
    except:
        # Fallback para variáveis de ambiente locais se não houver st.secrets
        conn = psycopg2.connect(
            host=os.environ.get("PGHOST"),
            database=os.environ.get("PGDATABASE"),
            user=os.environ.get("PGUSER"),
            password=os.environ.get("PGPASSWORD")
        )

    # 2. Busca de Dados
    data_ate = datetime.now().date()
    data_de = data_ate - timedelta(days=30)
    
    query = """
        SELECT store_code, business_dt, order_discount_amount
        FROM public.order_picture
        WHERE business_dt >= %s AND business_dt <= %s
        AND store_code NOT IN ('0000','0001','9999')
        AND state_id = 5
    """
    
    df = pd.read_sql(query, conn, params=(data_de, data_ate))
    conn.close()

    # 3. Processamento dos Dados
    # Remove zeros à esquerda do código da loja
    df["store_code"] = df["store_code"].astype(str).str.lstrip("0")
    # Formata data para MM/AAAA
    df["Mes_Referencia"] = pd.to_datetime(df["business_dt"]).dt.strftime("%m/%Y")
    # Converte desconto para número real
    df["Valor_Desconto"] = df["order_discount_amount"].apply(_parse_money)

    # 4. Geração do Arquivo Excel
    nome_arquivo = "CONFERENCIA_DESCONTOS.xlsx"
    df.to_excel(nome_arquivo, index=False)
    
    print(f"ARQUIVO GERADO COM SUCESSO: {nome_arquivo}")
    print("Nenhuma conexão com Google Sheets foi realizada.")

if __name__ == "__main__":
    gerar_apenas_excel_local()
