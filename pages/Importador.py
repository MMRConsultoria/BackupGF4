import streamlit as st
import pdfplumber
import pandas as pd
from io import BytesIO

uploaded_file = st.file_uploader("Faça upload do PDF", type="pdf")

if uploaded_file:
    with pdfplumber.open(uploaded_file) as pdf:
        all_text = ""
        all_tables = []
        for page in pdf.pages:
            all_text += page.extract_text() or ""
            tables = page.extract_tables()
            for table in tables:
                df = pd.DataFrame(table[1:], columns=table[0])
                all_tables.append(df)

    st.subheader("Texto extraído")
    st.text_area("Texto completo", all_text, height=300)

    if all_tables:
        st.subheader(f"{len(all_tables)} tabelas extraídas")
        for i, df in enumerate(all_tables):
            st.write(f"Tabela {i+1}")
            st.dataframe(df)

        # Concatenar todas as tabelas e permitir download Excel
        df_all = pd.concat(all_tables, ignore_index=True)
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
