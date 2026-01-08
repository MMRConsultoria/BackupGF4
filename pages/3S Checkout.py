import streamlit as st
import psycopg2
import pandas as pd
from io import BytesIO
from datetime import datetime
import json
import ast

CERT_PATH = "aws-us-east-2-bundle.pem"

# Grava o certificado em arquivo só uma vez por sessão
if "cert_written" not in st.session_state:
    with open(CERT_PATH, "w", encoding="utf-8") as f:
        f.write(st.secrets["certs"]["aws_rds_us_east_2"])
    st.session_state["cert_written"] = True

def get_conn():
    return psycopg2.connect(
        host=st.secrets["db"]["host"],
        port=st.secrets["db"]["port"],
        dbname=st.secrets["db"]["database"],
        user=st.secrets["db"]["user"],
        password=st.secrets["db"]["password"],
        sslmode="verify-full",
        sslrootcert=CERT_PATH,
    )

def fetch_table_data(conn, schema, table):
    query = f'SELECT * FROM "{schema}"."{table}"'
    return pd.read_sql(query, conn)

def sanitize_for_excel(df: pd.DataFrame, target_tz: str = "America/Sao_Paulo") -> pd.DataFrame:
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_datetime64tz_dtype(df[col]):
            df[col] = df[col].dt.tz_convert(target_tz).dt.tz_localize(None)
        elif df[col].dtype == "object":
            def _fix(x):
                if isinstance(x, (pd.Timestamp, datetime)) and getattr(x, "tzinfo", None) is not None:
                    ts = pd.Timestamp(x).tz_convert(target_tz)
                    return ts.tz_localize(None).to_pydatetime()
                return x
            df[col] = df[col].map(_fix)
    return df

# --- FUNÇÕES DE PARSE 3SCHECKOUT ---

def _to_dict(x):
    """Converte a string da coluna 'details' em um dicionário Python."""
    if not x or pd.isna(x): return {}
    if isinstance(x, dict): return x
    try:
        # Tenta JSON padrão
        return json.loads(x)
    except:
        try:
            # Tenta formato literal do Python (comum quando o log usa aspas simples)
            return ast.literal_eval(x)
        except:
            return {}

def parse_3s_details(row_value):
    """Extrai os campos específicos do 3SCheckout da coluna details."""
    d = _to_dict(row_value)
    if not d: return {}

    res = {
        "3s_setor": d.get("SECTOR"),
        "3s_pdv_tipo": d.get("POS_TYPE"),
        "3s_mesa_id": d.get("TABLE_ID"),
        "3s_mesa_nome": d.get("TABLE_NAME"),
        "3s_taxa_servico_pct": d.get("TIP_RATE"),
        "3s_taxa_servico_valor": d.get("TIP_AMOUNT"),
        "3s_fiscal_id": d.get("FISCAL_ID"),
        "3s_data_fiscalizacao": d.get("FISCALIZATION_DATE"),
    }

    # Processamento do BENEFIT_LIST (Descontos)
    benefit_raw = d.get("BENEFIT_LIST")
    if benefit_raw:
        # O BENEFIT_LIST costuma ser uma string JSON dentro do JSON
        b_dict = _to_dict(benefit_raw)
        # Dentro dele tem uma chave 'benefitList' que é uma string de lista
        b_list_str = b_dict.get("benefitList") or b_dict.get("benefit_list")
        
        try:
            b_list = json.loads(b_list_str) if isinstance(b_list_str, str) else b_list_str
            if isinstance(b_list, list) and len(b_list) > 0:
                main_benefit = b_list[0]
                res["3s_desconto_nome"] = main_benefit.get("benefit_id")
                res["3s_desconto_valor"] = main_benefit.get("benefit_total_value")
                res["3s_desconto_autorizado_por"] = main_benefit.get("authorized_by")
        except:
            pass
            
    return res

def export_db_to_excel(target_tz: str = "America/Sao_Paulo"):
    conn = get_conn()
    try:
        tables_to_export = [
            ("public", "order_picture"),
            ("public", "order_picture_tender")
        ]

        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            for schema, table in tables_to_export:
                df = fetch_table_data(conn, schema, table)
                
                # Se a coluna 'details' existir, aplica o detalhamento
                if "details" in df.columns:
                    details_expanded = df["details"].apply(parse_3s_details)
                    details_df = pd.DataFrame(details_expanded.tolist(), index=df.index)
                    df = pd.concat([df, details_df], axis=1)

                df = sanitize_for_excel(df, target_tz=target_tz)
                sheet_name = f"{schema}_{table}"[:31]
                df.to_excel(writer, sheet_name=sheet_name, index=False)

        output.seek(0)
        return output, None
    finally:
        conn.close()

# --- INTERFACE STREAMLIT ---

st.title("Exportador 3SCheckout (Detalhado)")

target_tz = st.selectbox("Fuso horário", ["America/Sao_Paulo", "UTC"])

if st.button("Gerar Excel", type="primary"):
    st.session_state["exporting"] = True
    with st.status("Processando dados...") as status:
        try:
            excel_bytes, err = export_db_to_excel(target_tz=target_tz)
            if err:
                st.error(err)
            else:
                status.update(label="Excel Gerado!", state="complete")
                st.download_button(
                    "Baixar Arquivo",
                    data=excel_bytes,
                    file_name=f"export_3s_{datetime.now().strftime('%Y%m%d')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Erro: {e}")
