# streamlit_relatorio_desconto.py
import os
import re
import json
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
    # normaliza separadores
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

def create_db_conn(params):
    return psycopg2.connect(
        host=params["host"],
        port=params["port"],
        dbname=params["dbname"],
        user=params["user"],
        password=params["password"]
    )

def create_gspread_client():
    # Verifica se as credenciais estão em st.secrets
    if "GOOGLE_SERVICE_ACCOUNT" not in st.secrets:
        raise RuntimeError(
            "Credenciais do Google não encontradas em st.secrets['GOOGLE_SERVICE_ACCOUNT']. "
            "Adicione as credenciais da service account (JSON) em st.secrets antes de usar esta função."
        )
    creds_json = st.secrets["GOOGLE_SERVICE_ACCOUNT"]
    if isinstance(creds_json, str):
        creds_dict = json.loads(creds_json)
    else:
        creds_dict = creds_json
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    return gspread.authorize(credentials)

# ----------------- Fetch / Cache -----------------
@st.cache_data(ttl=300)
def fetch_tabela_empresa():
    gc = create_gspread_client()
    try:
        sh = gc.open("Vendas diarias")
    except Exception as e:
        raise RuntimeError(f"Erro ao abrir a planilha 'Vendas diarias': {e}")
    try:
        ws = sh.worksheet("Tabela Empresa")
    except Exception as e:
        raise RuntimeError(f"Erro ao abrir a aba 'Tabela Empresa': {e}")

    values = ws.get_all_values()
    if not values:
        return pd.DataFrame()

    max_cols = max(len(r) for r in values)
    rows = [r + [""] * (max_cols - len(r)) for r in values]
    cols = [chr(ord("A") + i) for i in range(max_cols)]
    data_rows = rows[1:] if len(rows) > 1 else []
    df = pd.DataFrame(data_rows, columns=cols)
    df = df.loc[~(df[cols].apply(lambda r: all(str(x).strip() == "" for x in r), axis=1))]
    return df

@st.cache_data(ttl=300)
def fetch_order_picture(data_de, data_ate, excluir_stores=("0000", "0001", "9999"), estado_filtrar=5):
    """
    Tenta aplicar filtro por coluna VOID_TYPE; se coluna não existir, tenta pod_type;
    se também não existir, busca sem filtro de void/pod.
    """
    params = _get_db_params()
    if not params["dbname"] or not params["user"] or not params["password"]:
        raise RuntimeError("Credenciais do banco não encontradas. Configure st.secrets['db'] ou variáveis de ambiente PG*.")

    conn = create_db_conn(params)
    try:
        base_sql = """
            SELECT store_code, business_dt, order_discount_amount
            FROM public.order_picture
            WHERE business_dt >= %s
              AND business_dt <= %s
              AND store_code NOT IN %s
              AND state_id = %s
        """
        # tentativas de coluna que podem representar o "void"
        try_cols = [
            ("VOID_TYPE", "AND (VOID_TYPE IS NULL OR VOID_TYPE = '' OR LOWER(VOID_TYPE) NOT LIKE %s)"),
            ("pod_type", "AND (pod_type IS NULL OR pod_type = '' OR LOWER(pod_type) NOT LIKE %s)")
        ]
        like_void = "%void%"

        last_exc = None
        for col_name, cond_sql in try_cols:
            sql = f"{base_sql} {cond_sql} ORDER BY business_dt, store_code"
            try:
                df = pd.read_sql(sql, conn, params=(data_de, data_ate, tuple(excluir_stores), estado_filtrar, like_void))
                return df
            except Exception as e:
                last_exc = e
                msg = str(e).lower()
                # se o erro é por coluna inexistente, tenta próxima opção; caso contrário, re-raise
                if "does not exist" in msg or "column" in msg and col_name.lower() in msg:
                    continue
                else:
                    raise

        # fallback: buscar sem filtro de void/pod (para não quebrar o processo)
        sql = f"{base_sql} ORDER BY business_dt, store_code"
        df = pd.read_sql(sql, conn, params=(data_de, data_ate, tuple(excluir_stores), estado_filtrar))
        return df
    finally:
        conn.close()

# ----------------- Processamento -----------------
def process_and_build_report(df_orders: pd.DataFrame, df_empresa: pd.DataFrame) -> pd.DataFrame:
    if df_orders is None or df_orders.empty:
        return pd.DataFrame(columns=[
            "3S Checkout", "Business Month", "Loja", "Grupo",
            "Loja Nome", "Order Discount Amount (BRL)", "Store Code", "Código do Grupo"
        ])

    df = df_orders.copy()
    df["store_code"] = df["store_code"].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0").replace("", "0")
    df["business_dt"] = pd.to_datetime(df["business_dt"], errors="coerce")
    df["business_month"] = df["business_dt"].dt.strftime("%m/%Y").fillna("")
    df["order_discount_amount_val"] = df["order_discount_amount"].apply(_parse_money_to_float)
    df["order_discount_amount_fmt"] = df["order_discount_amount_val"].apply(lambda x: _format_brl(x if pd.notna(x) else 0.0))

    # prepara mapas a partir da Tabela Empresa (usando índices fixos das colunas)
    if df_empresa is None or df_empresa.empty:
        mapa_codigo_para_nome = {}
        mapa_codigo_para_colB = {}
        mapa_codigo_para_grupo = {}
    else:
        for col in ["A", "B", "C", "D"]:
            if col not in df_empresa.columns:
                df_empresa[col] = ""
        codigo_col = df_empresa["C"].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0").replace("", "0")
        mapa_codigo_para_nome = dict(zip(codigo_col, df_empresa["A"].astype(str)))
        mapa_codigo_para_colB = dict(zip(codigo_col, df_empresa["B"].astype(str)))
        mapa_codigo_para_grupo = dict(zip(codigo_col, df_empresa["D"].astype(str)))

    df["Loja Nome (lookup)"] = df["store_code"].map(mapa_codigo_para_nome)
    df["ColB (lookup)"] = df["store_code"].map(mapa_codigo_para_colB)
    df["Grupo (lookup)"] = df["store_code"].map(mapa_codigo_para_grupo)

    df_final = pd.DataFrame({
        "3S Checkout": ["3S Checkout"] * len(df),
        "Business Month": df["business_month"],
        "Loja": df["Loja Nome (lookup)"],
        "Grupo": df["ColB (lookup)"],                       # Coluna D -> coluna B da Tabela Empresa
        "Loja Nome": df["Loja Nome (lookup)"],
        "Order Discount Amount (BRL)": df["order_discount_amount_fmt"],
        "Store Code": df["store_code"],
        "Código do Grupo": df["Grupo (lookup)"]
    })

    col_order = [
        "3S Checkout", "Business Month", "Loja", "Grupo",
        "Loja Nome", "Order Discount Amount (BRL)", "Store Code", "Código do Grupo"
    ]
    df_final = df_final[col_order]
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

# ----------------- Streamlit UI -----------------
st.title("Relatório Desconto (lookup: Tabela Empresa) — Somente leitura")

st.markdown(
    "Gera um relatório Excel com as colunas na ordem solicitada, usando a aba 'Tabela Empresa' da planilha 'Vendas diarias' para lookup. "
    "Nenhuma planilha será atualizada."
)

dias_default = st.number_input("Últimos quantos dias", min_value=1, max_value=365, value=30)
data_ate = st.date_input("Data até", value=(datetime.utcnow() - timedelta(hours=3) - timedelta(days=1)).date())
data_de = st.date_input("Data de", value=(data_ate - timedelta(days=dias_default - 1)))

nome_arquivo = st.text_input("Nome do arquivo para download", value="relatorio_desconto_completo.xlsx")

if st.button("🔁 Gerar relatório completo"):
    try:
        with st.spinner("Buscando Tabela Empresa (Google Sheets)..."):
            df_empresa = fetch_tabela_empresa()
        with st.spinner("Buscando dados do banco..."):
            df_orders = fetch_order_picture(data_de, data_ate)
        with st.spinner("Processando relatório..."):
            df_final = process_and_build_report(df_orders, df_empresa)

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
