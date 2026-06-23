import streamlit as st
import requests
import pandas as pd

st.title("Teste ZIG")

token = st.secrets["zig"]["token"]

headers = {
    "Authorization": token
}

if st.button("Buscar lojas"):
    resp = requests.get(
        "https://api.zigcore.com.br/integration/erp/lojas",
        headers=headers
    )

    st.write("Status:", resp.status_code)

    try:
        st.json(resp.json())
    except:
        st.text(resp.text)
