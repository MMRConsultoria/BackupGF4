import streamlit as st
import pandas as pd
import tabula
from io import BytesIO

st.title("Leitor Completo de Tabelas do PDF")

uploaded_file = st.file_uploader("Faça upload do PDF", type="pdf")

if uploaded_file:
    # Salva temporariamente o arquivo
    with open("temp.pdf", "wb") as f:
        f.write(uploaded_file.getbuffer())

    # Extrai todas as tabelas do PDF em uma lista de DataFrames
    try:
        dfs = tabula.read_pdf("temp.pdf", pages='all', multiple_tables=True)
    except Exception as e:
        st.error(f"Erro ao ler tabelas do PDF: {e}")
        dfs = []

    if dfs:
        st.write(f"Foram encontradas {len(dfs)} tabelas no PDF.")
        for i, df in enumerate(dfs):
            st.subheader(f"Tabela {i+1}")
            st.dataframe(df)

        # Exemplo: concatenar todas as tabelas em um único DataFrame (se fizer sentido)
        df_all = pd.concat(dfs, ignore_index=True)
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df_all.to_excel(writer, index=False, sheet_name='Todas_Tabelas')
        output.seek(0)

        st.download_button(
            label="📥 Baixar todas as tabelas em Excel",
            data=output,
            file_name="tabelas_extraidas.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.info("Nenhuma tabela encontrada no PDF.")
