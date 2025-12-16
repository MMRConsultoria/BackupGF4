import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO

def extrair_dados(texto):
    empresa_match = re.search(r"Empresa:\s*\d+\s*-\s*(.+)", texto)
    nome_empresa = empresa_match.group(1).strip() if empresa_match else ""

    cnpj_match = re.search(r"Inscrição Federal:\s*([\d./-]+)", texto)
    cnpj = cnpj_match.group(1).strip() if cnpj_match else ""

    periodo_match = re.search(r"Período:\s*([\d/]+)\s*a\s*([\d/]+)", texto)
    periodo = f"{periodo_match.group(1)} a {periodo_match.group(2)}" if periodo_match else ""

    tabela_match = re.search(r"Resumo Contrato(.*?)Totais", texto, re.DOTALL)
    tabela_texto = tabela_match.group(1).strip() if tabela_match else ""

    linhas = [l.strip() for l in tabela_texto.split('\n') if l.strip()]
    dados = []
    for linha in linhas:
        cols = re.split(r'\s{2,}', linha)
        dados.append(cols)

    df_tabela = pd.DataFrame(dados)

    valores_match = re.search(
        r"Proventos:\s*([\d.,]+)\s*Vantagens:\s*([\d.,]+)\s*Descontos:\s*([\d.,]+)\s*Líquido:\s*([\d.,]+)",
        texto
    )
    proventos, vantagens, descontos, liquido = ("", "", "", "")
    if valores_match:
        proventos = valores_match.group(1)
        vantagens = valores_match.group(2)
        descontos = valores_match.group(3)
        liquido = valores_match.group(4)

    return {
        "nome_empresa": nome_empresa,
        "cnpj": cnpj,
        "periodo": periodo,
        "tabela": df_tabela,
        "proventos": proventos,
        "vantagens": vantagens,
        "descontos": descontos,
        "liquido": liquido
    }

st.title("📄 Extrator de Dados do Resumo da Folha (PDF)")

uploaded_file = st.file_uploader("Faça upload do arquivo PDF", type="pdf")

if uploaded_file:
    with pdfplumber.open(uploaded_file) as pdf:
        texto_completo = ""
        for page in pdf.pages:
            texto_completo += page.extract_text() + "\n"

    dados = extrair_dados(texto_completo)

    st.subheader("Informações extraídas")
    st.markdown(f"**Nome da Empresa:** {dados['nome_empresa']}")
    st.markdown(f"**CNPJ:** {dados['cnpj']}")
    st.markdown(f"**Período:** {dados['periodo']}")

    st.subheader("Tabela - Resumo Contrato")
    st.dataframe(dados["tabela"])

    st.subheader("Totais")
    st.markdown(f"- Proventos: {dados['proventos']}")
    st.markdown(f"- Vantagens: {dados['vantagens']}")
    st.markdown(f"- Descontos: {dados['descontos']}")
    st.markdown(f"- Líquido: {dados['liquido']}")

    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dados["tabela"].to_excel(writer, index=False, sheet_name='Resumo_Contrato')
    output.seek(0)

    st.download_button(
        label="📥 Baixar tabela em Excel",
        data=output,
        file_name="resumo_contrato.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
