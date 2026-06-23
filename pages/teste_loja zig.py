import streamlit as st
import requests
import pandas as pd
from datetime import date, timedelta

st.set_page_config(page_title="Teste ZIG", layout="wide")

st.title("🧪 Teste ZIG API")

token = st.secrets["zig"]["token"]
rede = st.secrets["zig"]["rede"]

headers = {
    "Authorization": token,
    "Accept": "application/json"
}

base_url = "https://api.zigcore.com.br/integration"

# =========================
# 1. BUSCAR LOJAS
# =========================
st.markdown("### 1. Buscar lojas")

if st.button("Buscar lojas"):
    resp = requests.get(
        f"{base_url}/erp/lojas",
        headers=headers,
        params={"rede": rede},
        timeout=30
    )

    st.write("Status:", resp.status_code)
    st.write("URL:", resp.url)

    try:
        dados = resp.json()
        st.json(dados)

        if isinstance(dados, list):
            df_lojas = pd.DataFrame(dados)
            st.success(f"{len(df_lojas)} lojas encontradas")
            st.dataframe(df_lojas, use_container_width=True)
            st.session_state["zig_lojas"] = df_lojas

    except Exception:
        st.text(resp.text)


# =========================
# 2. TESTAR FATURAMENTO
# =========================
st.markdown("---")
st.markdown("### 2. Testar faturamento")

lojas_fixas = {
    "HEINEKEN STAGING": "5fa1ff35-6145-4130-8707-8fd62e8822a5",
    "HEINEKEN HOUSE GRUPO FIT": "0e377e7a-9920-4674-916f-2ab196f8b691"
}

loja_nome = st.selectbox("Escolha a loja", list(lojas_fixas.keys()))
loja_id = lojas_fixas[loja_nome]

col1, col2 = st.columns(2)

with col1:
    dtinicio = st.date_input("Data início", value=date.today() - timedelta(days=7))

with col2:
    dtfim = st.date_input("Data fim", value=date.today())

if st.button("Buscar faturamento"):
    resp = requests.get(
        f"{base_url}/erp/faturamento",
        headers=headers,
        params={
            "dtinicio": dtinicio.strftime("%Y-%m-%d"),
            "dtfim": dtfim.strftime("%Y-%m-%d"),
            "loja": loja_id
        },
        timeout=60
    )

    st.write("Status:", resp.status_code)
    st.write("URL:", resp.url)

    try:
        dados = resp.json()
        st.json(dados)

        if isinstance(dados, list):
            df = pd.DataFrame(dados)

            if not df.empty and "value" in df.columns:
                df["valor_reais"] = df["value"] / 100

            st.dataframe(df, use_container_width=True)

            if not df.empty and "valor_reais" in df.columns:
                total = df["valor_reais"].sum()
                st.success(f"Total faturamento: R$ {total:,.2f}")

    except Exception:
        st.text(resp.text)
