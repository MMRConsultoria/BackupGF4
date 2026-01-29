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
    # normaliza separators
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
    """
    Lê a aba 'Tabela Empresa' da planilha 'Vendas diarias' e retorna um DataFrame
    com colunas nomeadas por letras 'A','B','C',... correspondendo às colunas da planilha.
    Usa get_all_values() para preservar posições fixas (A=0, C=2, D=3).
    """
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
        return pd.DataFrame()  # vazio

    # Normalizar número de colunas
    max_cols = max(len(r) for r in values)
    rows = [r + [""] * (max_cols - len(r)) for r in values]

    # Primeira linha pode ser cabeçalho; mas usaremos colunas por índice fixo,
    # então criamos colunas 'A','B','C',...
    cols = [chr(ord("A") + i) for i in range(max_cols)]
    # Ignorar a primeira linha (supondo que seja header) e retornar o restante.
    # Se preferir incluir a primeira linha como dado, remova [1:].
    data_rows = rows[1:] if len(rows) > 1 else []
    df = pd.DataFrame(data_rows, columns=cols)
    # Remove possíveis linhas em branco
    df = df.loc[~(df[cols].apply(lambda r: all(str(x).strip() == "" for x in r), axis=1))]
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

# ----------------- Processamento e montagem do relatório -----------------
def process_and_build_report(df_orders: pd.DataFrame, df_empresa: pd.DataFrame) -> pd.DataFrame:
    """
    Recebe:
      - df_orders com colunas: store_code, business_dt, order_discount_amount
      - df_empresa com colunas indexadas por letras 'A','B','C',... (A=nome loja, C=codigo loja, D=codigo grupo)
    Retorna DataFrame final com colunas na ordem solicitada.
    """
    if df_orders is None or df_orders.empty:
        return pd.DataFrame(columns=[
            "3S Checkout", "Business Month", "Loja", "Grupo",
            "Loja Nome", "Order Discount Amount (BRL)", "Store Code", "Código do Grupo"
        ])

    # processa orders
    df = df_orders.copy()
    df["store_code"] = df["store_code"].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0").replace("", "0")
    df["business_dt"] = pd.to_datetime(df["business_dt"], errors="coerce")
    df["business_month"] = df["business_dt"].dt.strftime("%m/%Y").fillna("")
    df["order_discount_amount_val"] = df["order_discount_amount"].apply(_parse_money_to_float)
    df["order_discount_amount_fmt"] = df["order_discount_amount_val"].apply(lambda x: _format_brl(x if pd.notna(x) else 0.0))

    # prepara mapas a partir da Tabela Empresa (usando índices fixos)
    # Coluna A -> index 0 -> letra 'A'
    # Coluna C -> index 2 -> letra 'C'
    # Coluna D -> index 3 -> letter 'D'
    if df_empresa is None or df_empresa.empty:
        mapa_codigo_para_nome = {}
        mapa_codigo_para_grupo = {}
    else:
        # garantir que temos colunas 'A','C','D' no df_empresa
        cols_present = set(df_empresa.columns.tolist())
        # cria colunas faltantes vazias se necessário
        for col in ["A", "C", "D"]:
            if col not in cols_present:
                df_empresa[col] = ""
        mapa_codigo_para_nome = dict(zip(df_empresa["C"].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0").replace("", "0"),
                                         df_empresa["A"].astype(str)))
        mapa_codigo_para_grupo = dict(zip(df_empresa["C"].astype(str).str.replace(r"\D", "", regex=True).str.lstrip("0").replace("", "0"),
                                          df_empresa["D"].astype(str)))

    # aplica lookup
    df["Loja Nome (lookup)"] = df["store_code"].map(mapa_codigo_para_nome)
    df["Grupo (lookup)"] = df["store_code"].map(mapa_codigo_para_grupo)

    # Monta DataFrame final
    df_final = pd.DataFrame({
        "3S Checkout": ["3S Checkout"]  len(df),
        "Business Month": df["business_month"],
        "Loja": df["Loja Nome (lookup)"],                # Col C: Loja (usando nome da Tabela Empresa Col A)
        "Grupo": df["Grupo (lookup)"],                   # Col D: Grupo (código do grupo da Tabela Empresa Col D)
        "Loja Nome": df["Loja Nome (lookup)"],           # Col E: Loja Nome (idem)
        "Order Discount Amount (BRL)": df["order_discount_amount_fmt"],  # Col F
        "Store Code": df["store_code"],                  # Col G
        "Código do Grupo": df["Grupo (lookup)"]          # Col H (mesmo que D)
    })

    # Opcional: reordenar colunas (já na ordem desejada)
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
st.title("Relatório Desconto (com lookup em Tabela Empresa) — Somente leitura")

st.markdown(
    "Gera um relatório Excel com as colunas na ordem solicitada, fazendo lookup na aba 'Tabela Empresa' da planilha 'Vendas diarias'. "
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
