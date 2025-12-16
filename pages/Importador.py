# pages/Extrair_Tabela_Imagem.py

import streamlit as st
import pytesseract
from PIL import Image
import pandas as pd
import re
from io import BytesIO

st.title("📷 Extrator de Tabela de Imagem via OCR")

uploaded_file = st.file_uploader("Faça upload da imagem da tabela (PNG, JPG, etc.)", type=["png", "jpg", "jpeg"])

if uploaded_file:
    img = Image.open(uploaded_file)
    st.image(img, caption="Imagem carregada", use_column_width=True)

    # Extrair texto com pytesseract
    try:
        text = pytesseract.image_to_string(img, lang='por')  # Use 'por' para português, se instalado
    except Exception as e:
        st.error(f"Erro ao executar OCR: {e}")
        st.stop()

    st.subheader("Texto extraído (preview)")
    st.text_area("Texto OCR", text, height=300)

    # Processar texto para extrair tabela
    lines = text.split('\n')
    data = []
    for line in lines:
        # Ignorar linhas vazias
        if not line.strip():
            continue
        # Dividir por múltiplos espaços ou tabulação
        cols = re.split(r'\s{2,}|\t', line.strip())
        # Filtrar linhas que parecem ter colunas suficientes (ajuste conforme necessário)
        if len(cols) >= 5:
            data.append(cols)

    if not data:
        st.warning("Não foi possível extrair dados tabulares do texto OCR.")
    else:
        # Criar DataFrame
        df = pd.DataFrame(data)

        st.subheader("Tabela extraída")
        st.dataframe(df)

        # Botão para download em Excel
        output = BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Tabela_Extraida')
        output.seek(0)

        st.download_button(
            label="📥 Baixar tabela em Excel",
            data=output,
            file_name="tabela_extraida.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
