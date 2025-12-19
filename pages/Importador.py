import streamlit as st
import pandas as pd
import re
from io import BytesIO
import pdfplumber

st.set_page_config(page_title="Resumo Proventos", layout="wide")

# =========================================================
# FUNÇÕES AUXILIARES
# =========================================================

def limpar_texto(txt):
    if pd.isna(txt):
        return ""
    return str(txt).replace(";", "").strip()


def parse_valor(valor):
    if pd.isna(valor):
        return None
    v = str(valor).strip()
    if not v:
        return None
    v = v.replace(".", "").replace(",", ".")
    try:
        return float(v)
    except ValueError:
        return None


# =========================================================
# PARSER QUESTOR (CSV / EXCEL)
# =========================================================

def ler_questor(uploaded_file):

    if uploaded_file.name.lower().endswith(".csv"):
        df = pd.read_csv(
            uploaded_file,
            sep=";",
            header=None,
            dtype=str,
            encoding="latin1"
        )
    else:
        df = pd.read_excel(uploaded_file, header=None, dtype=str)

    df = df.fillna("")

    registros = []
    empresa = ""
    cnpj = ""
    periodo = ""
    dentro_resumo = False

    for _, row in df.iterrows():

        linha = [limpar_texto(c) for c in row.tolist()]
        texto_linha = " ".join(linha).upper()

        # Empresa
        if not empresa and linha[0]:
            empresa = linha[0]

        # CNPJ
        if not cnpj:
            for c in linha:
                if re.search(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", c):
                    cnpj = c
                    break

        # Período
        if "PERÍODO" in texto_linha:
            periodo = texto_linha.replace("PERÍODO", "").strip()

        # Início / fim do resumo
        if "RESUMO CONTRATO" in texto_linha:
            dentro_resumo = True
            continue

        if dentro_resumo and ("TOTAL" in texto_linha or "LÍQUIDO" in texto_linha):
            dentro_resumo = False
            continue

        if not dentro_resumo:
            continue

        # 🔵 QUADRO ESQUERDO
        cod_esq = linha[0]
        tipo_esq = linha[1]
        desc_esq = linha[2]
        valor_esq = parse_valor(linha[4])

        if tipo_esq and desc_esq and valor_esq is not None:
            registros.append({
                "Sistema": "Questor",
                "Empresa": empresa,
                "CNPJ": cnpj,
                "Período": periodo,
                "Código": cod_esq,
                "Tipo": tipo_esq,
                "Descrição": desc_esq,
                "Valor": valor_esq
            })

        # 🔴 QUADRO DIREITO
        cod_dir = linha[5]
        tipo_dir = linha[6]
        desc_dir = linha[7]
        valor_dir = parse_valor(linha[9])

        if tipo_dir and desc_dir and valor_dir is not None:
            registros.append({
                "Sistema": "Questor",
                "Empresa": empresa,
                "CNPJ": cnpj,
                "Período": periodo,
                "Código": cod_dir,
                "Tipo": tipo_dir,
                "Descrição": desc_dir,
                "Valor": valor_dir
            })

    return pd.DataFrame(registros)


# =========================================================
# PARSER PDF ANTIGO (SIMPLIFICADO)
# =========================================================

def ler_pdf_antigo(uploaded_file):

    registros = []

    with pdfplumber.open(uploaded_file) as pdf:
        for page in pdf.pages:
            texto = page.extract_text()
            if not texto:
                continue

            linhas = texto.split("\n")
            empresa = linhas[0].strip() if linhas else ""
            periodo = ""

            for ln in linhas:
                if "PERÍODO" in ln.upper():
                    periodo = ln.replace("PERÍODO", "").strip()

                m = re.search(r"(.+?)\s+([\d\.]+,\d{2})$", ln)
                if m:
                    registros.append({
                        "Sistema": "PDF Antigo",
                        "Empresa": empresa,
                        "CNPJ": "",
                        "Período": periodo,
                        "Código": "",
                        "Tipo": "",
                        "Descrição": m.group(1).strip(),
                        "Valor": parse_valor(m.group(2))
                    })

    return pd.DataFrame(registros)


# =========================================================
# INTERFACE STREAMLIT
# =========================================================

st.title("📊 Resumo – Questor + PDF Antigo")

files = st.file_uploader(
    "Envie arquivos Questor (CSV/XLSX) e/ou PDF antigo",
    type=["csv", "xlsx", "pdf"],
    accept_multiple_files=True
)

if files:
    dfs = []

    for f in files:
        if f.name.lower().endswith((".csv", ".xlsx")):
            df_q = ler_questor(f)
            dfs.append(df_q)

        elif f.name.lower().endswith(".pdf"):
            df_p = ler_pdf_antigo(f)
            dfs.append(df_p)

    df_final = pd.concat(dfs, ignore_index=True)

    st.subheader("📋 Dados consolidados")
    st.dataframe(df_final, use_container_width=True)

    # =====================================================
    # EXPORTAÇÃO EXCEL
    # =====================================================

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Resumo")

    st.download_button(
        "📥 Baixar Excel",
        data=buffer.getvalue(),
        file_name="resumo_consolidado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
