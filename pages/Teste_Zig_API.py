import streamlit as st
import requests
import pandas as pd

st.title("Teste ZIG")

token = st.secrets["zig"]["token"]
rede = st.secrets["zig"]["rede"]

if st.button("Buscar lojas"):

    resp = requests.get(
        "https://api.zigcore.com.br/integration/erp/lojas",
        headers={
            "Authorization": token
        },
        params={
            "rede": rede
        },
        timeout=30
    )

    st.write("Status:", resp.status_code)
    st.write("URL:", resp.url)

    try:
        dados = resp.json()

        st.json(dados)

        if isinstance(dados, list):
            df = pd.DataFrame(dados)
            st.success(f"{len(df)} lojas encontradas")
            st.dataframe(df, use_container_width=True)

    except Exception:
        st.text(resp.text)
