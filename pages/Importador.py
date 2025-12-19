import streamlit as st
import pandas as pd
import pdfplumber
import re
from io import BytesIO

# =========================
# UTILITÁRIOS
# =========================

def clean_company_name(nome):
    return re.sub(r";+", "", nome).strip()

def extrair_mes_ano(periodo):
    if not periodo:
        return "", ""
    m = re.search(r"\d{2}/(\d{2})/(\d{4})", periodo)
    if m:
        return m.group(1), m.group(2)
    return "", ""

def _to_float_br(valor):
    try:
        return float(
            valor.replace(".", "")
                 .replace(",", ".")
                 .replace("R$", "")
                 .strip()
        )
    except:
        return 0.0


# =========================
# PDF (BASE – NÃO ALTERADO)
# =========================

def extrair_dados_pdf(texto):
    linhas = texto.splitlines()

    empresa_raw = ""
    cnpj = ""
    periodo = ""

    for l in linhas:
        if "Empresa:" in l:
            empresa_raw = l.split("Empresa:")[-1].strip()
        if "CNPJ:" in l:
            cnpj = l.split("CNPJ:")[-1].strip()
        if "Período:" in l:
            periodo = l.split("Período:")[-1].strip()

    codigo_empresa = ""
    nome_empresa = empresa_raw
    if "-" in empresa_raw:
        codigo_empresa, nome_empresa = empresa_raw.split("-", 1)

    nome_empresa = clean_company_name(nome_empresa)
    mes, ano = extrair_mes_ano(periodo)

    rows = []

    for l in linhas:
        parts = l.split()
        if len(parts) >= 3 and parts[-1].replace(",", "").replace(".", "").isdigit():
            try:
                valor = parts[-1]
                codigo = parts[0]
                descricao = " ".join(parts[1:-1])
                rows.append({
                    "Codigo Empresa": codigo_empresa,
                    "Empresa": nome_empresa,
                    "CNPJ": cnpj,
                    "Período": periodo,
                    "Mês": mes,
                    "Ano": ano,
                    "Tipo": "",
                    "Codigo da Descrição": codigo,
                    "Descrição": descricao,
                    "Valor": valor,
                    "Valor_num": _to_float_br(valor),
                    "Sistema": "PDF"
                })
            except:
                pass

    df = pd.DataFrame(rows)

    return df


# =========================
# CSV QUESTOR
# =========================

def extrair_dados_csv_questor(uploaded_file):
    df_raw = pd.read_csv(
        uploaded_file,
        sep=";",
        header=None,
        dtype=str,
        encoding="latin1"
    ).fillna("")

    # ---- Cabeçalho
    empresa_raw = df_raw.iloc[2, 1]
    sistema = df_raw.iloc[2, -1]

    codigo_empresa = ""
    nome_empresa = empresa_raw
    if "-" in empresa_raw:
        codigo_empresa, nome_empresa = empresa_raw.split("-", 1)

    nome_empresa = clean_company_name(nome_empresa)
    cnpj = df_raw.iloc[3, 1]

    periodo_raw = df_raw.iloc[7, 0]
    m = re.search(r"(\d{2}/\d{2}/\d{4}\s*a\s*\d{2}/\d{2}/\d{4})", periodo_raw)
    periodo = m.group(1) if m else ""

    mes, ano = extrair_mes_ano(periodo)

    # ---- Encontrar início
    start = None
    for i in range(len(df_raw)):
        if df_raw.iloc[i, 0].strip().lower() == "resumo contrato":
            start = i + 1
            break

    rows = []

    tipo_map = {
        "1": "Proventos",
        "2": "Vantagens",
        "3": "Descontos",
        "4": "Informativo",
        "5": "Informativo"
    }

    for i in range(start, len(df_raw)):
        if df_raw.iloc[i, 0].strip().lower() == "totais":
            break

        # 🔵 Quadro esquerdo
        cod = df_raw.iloc[i, 0].strip()
        tipo = df_raw.iloc[i, 1].strip()
        desc = df_raw.iloc[i, 2].strip()
        valor = df_raw.iloc[i, 4].strip()

        if cod and valor:
            rows.append({
                "Codigo Empresa": codigo_empresa,
                "Empresa": nome_empresa,
                "CNPJ": cnpj,
                "Período": periodo,
                "Mês": mes,
                "Ano": ano,
                "Tipo": tipo_map.get(tipo, ""),
                "Codigo da Descrição": cod,
                "Descrição": desc,
                "Valor": valor,
                "Valor_num": _to_float_br(valor),
                "Sistema": sistema
            })

        # 🔴 Quadro direito
        cod = df_raw.iloc[i, 5].strip()
        tipo = df_raw.iloc[i, 6].strip()
        desc = df_raw.iloc[i, 7].strip()
        valor = df_raw.iloc[i, 9].strip()

        if cod and valor:
            rows.append({
                "Codigo Empresa": codigo_empresa,
                "Empresa": nome_empresa,
                "CNPJ": cnpj,
                "Período": periodo,
                "Mês": mes,
                "Ano": ano,
                "Tipo": tipo_map.get(tipo, ""),
                "Codigo da Descrição": cod,
                "Descrição": desc,
                "Valor": valor,
                "Valor_num": _to_float_br(valor),
                "Sistema": sistema
            })

    return pd.DataFrame(rows)


# =========================
# STREAMLIT APP
# =========================

st.set_page_config(layout="wide")
st.title("Leitura GF4 – PDF + CSV Questor")

files = st.file_uploader(
    "Envie PDFs ou CSVs",
    accept_multiple_files=True
)

dfs = []

if files:
    for f in files:
        if f.name.lower().endswith(".pdf"):
            with pdfplumber.open(f) as pdf:
                texto = ""
                for p in pdf.pages:
                    texto += (p.extract_text() or "") + "\n"
            df = extrair_dados_pdf(texto)

        elif f.name.lower().endswith(".csv"):
            df = extrair_dados_csv_questor(f)

        dfs.append(df)

    df_final = pd.concat(dfs, ignore_index=True)

    df_final = df_final[[
        "Codigo Empresa", "Empresa", "CNPJ", "Período",
        "Mês", "Ano", "Tipo",
        "Codigo da Descrição", "Descrição", "Valor", "Sistema"
    ]]

    st.dataframe(df_final, use_container_width=True)

    # Excel
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Resumo")

    st.download_button(
        "📥 Baixar Excel",
        data=buffer.getvalue(),
        file_name="GF4_RESUMO.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
