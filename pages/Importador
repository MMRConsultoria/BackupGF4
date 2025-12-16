# pages/Leitor_Folha_Ponto.py

import streamlit as st
import pandas as pd
from io import BytesIO
import re

# Função para extrair texto do PDF
def parse_pdf_content(file_path):
    from PyPDF2 import PdfReader

    reader = PdfReader(file_path)
    text = ""
    for page in reader.pages:
        text += page.extract_text()

    # Extraindo dados principais com regex
    empresa_match = re.search(r"Empresa:\s*(\d+)\s*-\s*(.+)", text)
    cnpj_match = re.search(r"Inscrição Federal:\s*([\d\./-]+)", text)
    periodo_match = re.search(r"Período:\s*([\d/]+)\s*a\s*([\d/]+)", text)

    empresa_codigo = empresa_match.group(1) if empresa_match else ""
    empresa_nome = empresa_match.group(2) if empresa_match else ""
    cnpj = cnpj_match.group(1) if cnpj_match else ""
    periodo_inicio = periodo_match.group(1) if periodo_match else ""
    periodo_fim = periodo_match.group(2) if periodo_match else ""

    # Extraindo totais de provisões, descontos e líquidos
    provisoes_match = re.search(r"Proventos:\s*([\d\.,]+)", text)
    descontos_match = re.search(r"Descontos:\s*([\d\.,]+)", text)
    liquido_match = re.search(r"Líquido:\s*([\d\.,]+)", text)

    provisoes = provisoes_match.group(1) if provisoes_match else "0,00"
    descontos = descontos_match.group(1) if descontos_match else "0,00"
    liquido = liquido_match.group(1) if liquido_match else "0,00"

    # Convertendo valores para float
    def br_to_float(val):
        return float(val.replace(".", "").replace(",", "."))

    df = pd.DataFrame([{
        "Código Empresa": empresa_codigo,
        "Nome Empresa": empresa_nome,
        "CNPJ": cnpj,
        "Período Início": periodo_inicio,
        "Período Fim": periodo_fim,
        "Proventos": br_to_float(provisoes),
        "Descontos": br_to_float(descontos),
        "Líquido": br_to_float(liquido)
    }])
    return df

# Interface Streamlit
st.title("📄 Leitor de Resumo da Folha de Pagamento (PDF)")
uploaded_file = st.file_uploader("Faça upload do PDF", type="pdf")

if uploaded_file:
    with open("temp.pdf", "wb") as f:
        f.write(uploaded_file.getbuffer())

    df_dados = parse_pdf_content("temp.pdf")
    st.subheader("Dados Extraídos")
    st.dataframe(df_dados)

    # Botão para download em Excel
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_dados.to_excel(writer, index=False, sheet_name='Resumo')
    output.seek(0)

    st.download_button(
        label="📥 Baixar Excel",
        data=output,
        file_name="resumo_folha.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
