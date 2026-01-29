import os
import re
from io import BytesIO
from datetime import datetime, timedelta

import streamlit as st
import pandas as pd
import psycopg2
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ----------------- Helpers -----------------
def _parse_money_to_float(x):
    if pd.isna(x):
        return None
    s = str(x).strip()
    s = s.replace("R$", "").replace("\u00A0", "").replace(" ", "")
    s = re.sub(r"[^\d,\-\.]", "", s)
    if s == "":
        return None
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        try:
            return float(s.replace(",", "."))
        except Exception:
            return None

def _format_brl(v):
    try:
        v = float(v)
    except Exception:
        return "R$ 0,00"
    s = f"{v:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return f"R$ {s}"

def _get_db_params():
    try:
        db = st.secrets["db"]
        return {
            "host": db["host"],
            "port": int(db.get("port", 5432)),
            "dbname": db["database"],
            "user": db["user"],
            "password": db["password"]
        }
    except Exception:
        return {
            "host": os.environ.get("PGHOST", "localhost"),
            "port": int(os.environ.get("PGPORT", 5432)),
            "dbname": os.environ.get("PGDATABASE", ""),
            "user": os.environ.get("PGUSER", ""),
            "password": os.environ.get("PGPASSWORD", "")
        }

def create_db_conn(params):
    return psycopg2.connect(
        host=params["host"],
        port=params["port"],
        dbname=params["dbname"],
        user=params["user"],
        password=params["password"]
    )

def create_gspread_client():
    creds_json = st.secrets["GOOGLE_SERVICE_ACCOUNT"]
    if isinstance(creds_json, str):
        creds_dict = json.loads(creds_json)
    else:
        creds_dict = creds_json
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    return gspread.authorize(credentials)

# ----------------- Funções -----------------
@st.cache_data(ttl=300)
def fetch_tabela_empresa():
    gc = create_gspread_client()
    sh = gc.open("Vendas diarias")
    ws = sh.worksheet("Tabela Empresa")
    data = ws.get_all_records()
    df = pd.DataFrame(data)
    # Normaliza nomes colunas para evitar erros
    df.columns = [c.strip() for c in df.columns]
    return df

@st.cache_data(ttl=300)
def fetch_order_picture(data_de, data_ate, excluir_stores=("0000", "0001", "9999"), estado_filtrar=5):
    params = _get_db_params()
    if not params["dbname"] or not params["user"] or not params["password"]:
        raise RuntimeError("Credenciais do banco não encontradas. Configure st.secrets['db'] ou variáveis de ambiente PG*.")

    conn = create_db_conn(params)
    try:
        sql = """
            SELECT store_code, business_dt, order_discount_amount
            FROM public.order_picture
            WHERE business_dt >= %s
              AND business_dt <= %s
              AND store_code NOT IN %s
              AND state_id = %s
            ORDER BY business_dt, store_code
        """
        df = pd.read_sql(sql, conn, params=(data_de, data_ate, tuple(excluir_stores), estado_filtrar))
    finally:
        conn.close()
    return df

def process_and_merge(df_orders, df_empresa):
    if df_orders is None or df_orders.empty:
        return pd.DataFrame()

    # Limpa store_code e remove zeros à esquerda
    df_orders["store_code"] = df_orders["store_code"].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0").replace("", "0")
    df_orders["business_dt"] = pd.to_datetime(df_orders["business_dt"], errors="coerce")
    df_orders["business_month"] = df_orders["business_dt"].dt.strftime("%m/%Y").fillna("")

    df_orders["order_discount_amount_val"] = df_orders["order_discount_amount"].apply(_parse_money_to_float)
    df_orders["order_discount_amount_fmt"] = df_orders["order_discount_amount_val"].apply(lambda x: _format_brl(x if pd.notna(x) else 0.0))

    # Normaliza colunas da tabela empresa para facilitar merge
    df_empresa.columns = [c.strip() for c in df_empresa.columns]
    # Ajuste os nomes abaixo conforme sua planilha
    # Supondo:
    # Col A: Nome da loja (ex: "Loja Nome")
    # Col C: Código da loja (ex: "Store Code")
    # Col D: Código do grupo (ex: "Grupo")
    nome_loja_col = df_empresa.columns[0]  # Col A
    codigo_loja_col = df_empresa.columns[2]  # Col C
    codigo_grupo_col = df_empresa.columns[3]  # Col D

    # Faz merge para trazer Grupo e Loja Nome
    df_merged = pd.merge(
        df_orders,
        df_empresa[[codigo_loja_col, codigo_grupo_col, nome_loja_col]],
        how="left",
        left_on="store_code",
        right_on=codigo_loja_col
    )

    # Monta DataFrame final com as colunas na ordem pedida
    df_final = pd.DataFrame({
        "3S Checkout": "3S Checkout",
        "Business Month": df_merged["business_month"],
        "Loja": df_merged[nome_loja_col],
        "Grupo": df_merged[codigo_grupo_col],
        "Loja Nome": df_merged[nome_loja_col],
        "Order Discount Amount (BRL)": df_merged["order_discount_amount_fmt"],
        "Store Code": df_merged["store_code"],
        "Código do Grupo": df_merged[codigo_grupo_col]
    })

    return df_final

def to_excel_bytes(df: pd.DataFrame, sheet_name="Desconto"):
    output = BytesIO()
    try:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max() if not df.empty else 0, len(col)) + 2
                worksheet.set_column(i, i, max_len)
    except Exception:
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

# ----------------- UI Streamlit -----------------
st.title("Relatório Desconto com dados da Tabela Empresa")

dias_default = st.number_input("Últimos quantos dias", min_value=1, max_value=365, value=30)
data_ate = st.date_input("Data até", value=(datetime.utcnow() - timedelta(hours=3) - timedelta(days=1)).date())
data_de = st.date_input("Data de", value=(data_ate - timedelta(days=dias_default - 1)))

nome_arquivo = st.text_input("Nome do arquivo para download", value="relatorio_desconto_completo.xlsx")

if st.button("🔁 Gerar relatório completo"):
    try:
        with st.spinner("Buscando dados da Tabela Empresa..."):
            df_empresa = fetch_tabela_empresa()
        with st.spinner("Buscando dados do banco..."):
            df_orders = fetch_order_picture(data_de, data_ate)
        with st.spinner("Processando e juntando dados..."):
            df_final = process_and_merge(df_orders, df_empresa)
        if df_final.empty:
            st.warning("Nenhum dado encontrado para o período selecionado.")
        else:
            st.dataframe(df_final.head(200))
            excel_bytes = to_excel_bytes(df_final, sheet_name="Desconto")
            st.download_button(
                label="⬇️ Baixar relatório Excel",
                data=excel_bytes,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Erro ao gerar relatório: {e}")
        st.exception(e)
