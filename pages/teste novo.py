import streamlit as st
import pandas as pd
import psycopg2
from datetime import date, timedelta

st.set_page_config(layout="wide", page_title="Diagnóstico Meio de Pagamento 3S")

CERT_PATH = "aws-us-east-2-bundle.pem"

def ensure_cert_written():
    if "cert_written_diag" not in st.session_state:
        with open(CERT_PATH, "w", encoding="utf-8") as f:
            f.write(st.secrets["certs"]["aws_rds_us_east_2"])
        st.session_state["cert_written_diag"] = True

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

ensure_cert_written()

st.title("Diagnóstico: Meio de Pagamento vazio por loja (3S)")

loja = st.text_input("Store code (ex: 0087 ou 87)", value="")
data_inicio = st.date_input("Desde quando?", value=date.today() - timedelta(days=30))
data_fim = st.date_input("Até quando?", value=date.today())

if st.button("Consultar", type="primary"):
    if not loja.strip():
        st.error("Informe o store_code.")
        st.stop()

    loja_norm = loja.strip().lstrip("0")  # igual seu código faz

    conn = get_db_conn()
    try:
        # 1) Pega order_picture_ids da loja no período
        q_op = """
            SELECT order_picture_id, store_code, business_dt, custom_properties
            FROM public.order_picture
            WHERE business_dt >= %s
              AND business_dt <= %s
              AND state_id = 5
        """
        df_op = pd.read_sql(q_op, conn, params=(data_inicio, data_fim))
        df_op["store_code"] = df_op["store_code"].astype(str).str.lstrip("0").str.strip()
        df_op = df_op[df_op["store_code"] == loja_norm].copy()

        st.write("### order_picture (filtrado)")
        st.write(f"Linhas: {len(df_op)}")
        st.dataframe(df_op.head(50), use_container_width=True)

        if df_op.empty:
            st.warning("Nenhum order_picture para essa loja no período.")
            st.stop()

        ids = df_op["order_picture_id"].dropna().astype(int).tolist()

        # 2) Busca tenders desses pedidos
        q_t = """
            SELECT order_picture_id, tender_amount, change_amount, details
            FROM public.order_picture_tender
            WHERE order_picture_id = ANY(%s)
        """
        df_t = pd.read_sql(q_t, conn, params=(ids,))

        st.write("### order_picture_tender (RAW)")
        st.write(f"Linhas: {len(df_t)}")
        st.dataframe(df_t.head(50), use_container_width=True)

        # 3) Diagnóstico rápido: quantos details vazios / nulos
        st.write("### Diagnóstico de details")
        st.write("details nulo:", int(df_t["details"].isna().sum()))
        st.write("details vazio (string):", int((df_t["details"].astype(str).str.strip() == "").sum()))

        # 4) Mostra exemplos de details pra achar a chave do meio de pagamento
        st.write("### Amostras de `details` (copie e cole aqui se precisar)")
        amostra = df_t["details"].dropna().astype(str).head(20).tolist()
        for i, d in enumerate(amostra, start=1):
            st.code(d, language="json")

    finally:
        conn.close()
