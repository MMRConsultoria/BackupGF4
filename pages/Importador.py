import streamlit as st
import camelot

uploaded_file = st.file_uploader("Faça upload do PDF", type="pdf")

if uploaded_file:
    with open("temp.pdf", "wb") as f:
        f.write(uploaded_file.getbuffer())

    tables = camelot.read_pdf("temp.pdf", pages='all')

    st.write(f"Encontradas {tables.n} tabelas.")

    for i, table in enumerate(tables):
        st.subheader(f"Tabela {i+1}")
        st.dataframe(table.df)

    # Exportar todas as tabelas para Excel
    import pandas as pd
    from io import BytesIO

    df_all = pd.concat([t.df for t in tables], ignore_index=True)
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
