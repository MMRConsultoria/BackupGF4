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

# Inicializa o estado de exportação
if "exporting" not in st.session_state:
    st.session_state["exporting"] = False


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


def fetch_filtered_data(conn):
    """Busca dados da tabela order_picture de forma simples"""
    query = "SELECT store_code, business_dt, total_gross, custom_properties, order_code FROM public.order_picture"
    return pd.read_sql(query, conn)


def sanitize_for_excel(df: pd.DataFrame) -> pd.DataFrame:
    """Remove timezones para compatibilidade com Excel."""
    df = df.copy()
    for col in df.columns:
        if pd.api.types.is_datetime64tz_dtype(df[col]):
            df[col] = df[col].dt.tz_localize(None)
        elif df[col].dtype == "object":
            def _fix(x):
                if isinstance(x, (pd.Timestamp, datetime)) and getattr(x, "tzinfo", None) is not None:
                    return pd.Timestamp(x).tz_localize(None).to_pydatetime()
                return x
            df[col] = df[col].map(_fix)
    return df


def parse_props(x):
    """Converte JSON/str em dict Python."""
    if pd.isna(x):
        return {}
    try:
        if isinstance(x, str):
            return json.loads(x)
    except:
        try:
            return ast.literal_eval(x)
        except:
            return {}
    return x if isinstance(x, dict) else {}


def export_order_picture_to_excel():
    conn = get_conn()
    try:
        # Busca dados do banco
        df = fetch_filtered_data(conn)
        
        # 1. Converter datas (SEM FILTRO DE DATA)
        df['business_dt'] = pd.to_datetime(df['business_dt'], errors='coerce')
        
        # 2. Filtrar lojas (Excluir 0000, 0001, 9999)
        df['store_code'] = df['store_code'].astype(str).str.zfill(4)
        excluir = ['0000', '0001', '9999']
        df = df[~df['store_code'].isin(excluir)].copy()
        
        # 3. Extrair campos de custom_properties (TIP_AMOUNT e VOID_TYPE)
        props = df['custom_properties'].apply(parse_props)
        df['TIP_AMOUNT'] = pd.to_numeric(props.apply(lambda x: x.get('TIP_AMOUNT')), errors='coerce').fillna(0)
        df['VOID_TYPE'] = props.apply(lambda x: x.get('VOID_TYPE'))
        
        # 4. Desconsiderar registros com VOID_TYPE preenchido
        df = df[df['VOID_TYPE'].isna() | (df['VOID_TYPE'] == "") | (df['VOID_TYPE'] == 0)].copy()
        
        # 5. Criar coluna de data sem hora para agrupamento
        df['data'] = df['business_dt'].dt.date
        
        # 6. Agrupar totais por store_code e data
        resumo = df.groupby(['store_code', 'data']).agg(
            qtd_pedidos=('order_code', 'count'),
            total_gross=('total_gross', 'sum'),
            total_tip=('TIP_AMOUNT', 'sum')
        ).reset_index()
        
        # Limpa datas para o Excel
        resumo = sanitize_for_excel(resumo)
        
        # Gera Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            resumo.to_excel(writer, sheet_name="Resumo_Loja_Dia", index=False)

        output.seek(0)
        return output, None, len(df), len(resumo)
    except Exception as e:
        return None, str(e), 0, 0
    finally:
        conn.close()


# -------------------------
# UI
# -------------------------
st.title("Exportar order_picture - Resumo por Loja e Dia")
st.subheader("Filtros aplicados:")
st.markdown("""
- **Período**: Todas as datas
- **Lojas excluídas**: 0000, 0001, 9999
- **Registros válidos**: sem VOID_TYPE preenchido
- **Colunas**: store_code, business_dt, total_gross, TIP_AMOUNT
- **Resultado**: Totais agrupados por loja e dia
""")

# Botão de reset (caso fique travado)
if st.button("🔄 Resetar Página", type="secondary"):
    st.session_state["exporting"] = False
    st.rerun()

st.write("Clique no botão abaixo para gerar o Excel com o resumo.")

if st.button("Gerar Excel", type="primary", disabled=st.session_state["exporting"]):
    st.session_state["exporting"] = True
    status = st.status("Processando dados...", expanded=True)

    try:
        status.write("Conectando ao banco e lendo tabela...")
        excel_bytes, err, total_rows, summary_rows = export_order_picture_to_excel()

        if err:
            status.update(label="Falhou", state="error")
            st.error(f"Erro no banco: {err}")
        else:
            status.update(label="Concluído ✅", state="complete")
            st.success(f"""
            **Processamento concluído!**
            - Registros válidos processados: {total_rows:,}
            - Linhas no resumo (loja + dia): {summary_rows:,}
            """)
            st.download_button(
                "📥 Baixar Excel",
                data=excel_bytes,
                file_name=f"resumo_vendas_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    except Exception as e:
        status.update(label="Falhou", state="error")
        st.error(f"Erro inesperado: {e}")
    finally:
        st.session_state["exporting"] = False
