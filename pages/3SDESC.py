# streamlit_app_desconto.py
import os
import re
from io import BytesIO
from datetime import datetime, timedelta

import streamlit as st
import pandas as pd
import psycopg2

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
    # tenta st.secrets['db'] (quando rodando no Streamlit cloud)
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
        # fallback para variáveis de ambiente locais
        return {
            "host": os.environ.get("PGHOST", "localhost"),
            "port": int(os.environ.get("PGPORT", 5432)),
            "dbname": os.environ.get("PGDATABASE", ""),
            "user": os.environ.get("PGUSER", ""),
            "password": os.environ.get("PGPASSWORD", "")
        }

# ----------------- Funções de negócio -----------------
@st.cache_data(ttl=300)
def fetch_order_picture(data_de, data_ate, excluir_stores=("0000", "0001", "9999"), estado_filtrar=5):
    params = _get_db_params()
    if not params["dbname"] or not params["user"] or not params["password"]:
        raise RuntimeError("Credenciais do banco não encontradas. Configure st.secrets['db'] ou variáveis de ambiente PG*.")

    conn = psycopg2.connect(
        host=params["host"],
        port=params["port"],
        dbname=params["dbname"],
        user=params["user"],
        password=params["password"]
    )
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

def process_df(df):
    if df is None or df.empty:
        return pd.DataFrame(columns=[
            "Store Code", "Business Date", "Business Month",
            "Order Discount Amount (num)", "Order Discount Amount (BRL)"
        ])
    df["store_code"] = df["store_code"].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0").replace("", "0")
    df["business_dt"] = pd.to_datetime(df["business_dt"], errors="coerce")
    df["business_month"] = df["business_dt"].dt.strftime("%m/%Y").fillna("")
    df["order_discount_amount_val"] = df["order_discount_amount"].apply(_parse_money_to_float)
    df["order_discount_amount_fmt"] = df["order_discount_amount_val"].apply(lambda x: _format_brl(x if pd.notna(x) else 0.0))
    df_out = df[[
        "store_code", "business_dt", "business_month",
        "order_discount_amount_val", "order_discount_amount_fmt"
    ]].rename(columns={
        "store_code": "Store Code",
        "business_dt": "Business Date",
        "business_month": "Business Month",
        "order_discount_amount_val": "Order Discount Amount (num)",
        "order_discount_amount_fmt": "Order Discount Amount (BRL)"
    })
    return df_out

def to_excel_bytes(df: pd.DataFrame, sheet_name="Desconto"):
    output = BytesIO()
    # tenta usar xlsxwriter se disponível
    try:
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
            workbook = writer.book
            worksheet = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max() if not df.empty else 0, len(col)) + 2
                worksheet.set_column(i, i, max_len)
    except Exception:
        # fallback sem formatação
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name=sheet_name)
    output.seek(0)
    return output

# ----------------- UI Streamlit -----------------
st.title("Gerar Desconto (.xlsx) — Somente leitura")

st.markdown(
    "Este botão consulta public.order_picture no banco, processa os dados e gera um arquivo Excel em memória. "
    "Nada é gravado no disco e nenhuma planilha Google será atualizada automaticamente."
)

# Parâmetros do período
dias_default = st.number_input("Últimos quantos dias", min_value=1, max_value=365, value=30)
data_ate = st.date_input("Data até", value=(datetime.utcnow() - timedelta(hours=3) - timedelta(days=1)).date())
data_de = st.date_input("Data de", value=(data_ate - timedelta(days=dias_default - 1)))

col1, col2 = st.columns([3, 1])
with col1:
    nome_arquivo = st.text_input("Nome do arquivo para download (ex.: relatorio_desconto.xlsx)", value="relatorio_desconto.xlsx")
with col2:
    st.write("")  # alinhamento
    st.write("") 

if st.button("🔁 Gerar relatório (somente Excel)"):
    try:
        with st.spinner("Consultando banco e gerando relatório..."):
            df_raw = fetch_order_picture(data_de, data_ate)
            df_proc = process_df(df_raw)
            if df_proc.empty:
                st.info("Nenhum registro encontrado no período selecionado.")
            else:
                st.success(f"{len(df_proc)} linhas processadas.")
            excel_bytes = to_excel_bytes(df_proc, sheet_name="Desconto")
            # exibe uma pré-visualização (primeiras linhas)
            st.dataframe(df_proc.head(200))
            # botão de download — sem gravar no disco e sem acionar reload
            st.download_button(
                label="⬇️ Baixar arquivo .xlsx",
                data=excel_bytes,
                file_name=nome_arquivo,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Erro ao gerar relatório: {e}")
        st.exception(e)
