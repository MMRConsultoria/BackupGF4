import streamlit as st
import requests
import pandas as pd
from datetime import date, timedelta, datetime
from io import BytesIO

st.set_page_config(page_title="ZIG - Teste Gorjeta", layout="wide")
st.title("🧪 ZIG - Teste Gorjeta por Compradores")

token = st.secrets["zig"]["token"]
rede = st.secrets["zig"]["rede"]

headers = {
    "Authorization": token,
    "Accept": "application/json"
}

base_url = "https://api.zigcore.com.br/integration"


col1, col2 = st.columns(2)

with col1:
    dtinicio = st.date_input("Data início", value=date.today() - timedelta(days=1))

with col2:
    dtfim = st.date_input("Data fim", value=date.today() - timedelta(days=1))


if st.button("🔍 Buscar Gorjetas"):

    resp_lojas = requests.get(
        f"{base_url}/erp/lojas",
        headers=headers,
        params={"rede": rede},
        timeout=30
    )

    if resp_lojas.status_code != 200:
        st.error("Erro ao buscar lojas ZIG")
        st.text(resp_lojas.text)
        st.stop()

    lojas = resp_lojas.json()

    registros_compradores = []

    for loja in lojas:
        loja_id = loja.get("id")
        loja_nome = loja.get("name")

        resp = requests.get(
            f"{base_url}/erp/compradores",
            headers=headers,
            params={
                "dtinicio": dtinicio.strftime("%Y-%m-%d"),
                "dtfim": dtfim.strftime("%Y-%m-%d"),
                "loja": loja_id
            },
            timeout=60
        )

        if resp.status_code != 200:
            continue

        dados = resp.json()

        if not isinstance(dados, list):
            continue

        for item in dados:
            registros_compradores.append({
                "Loja": loja_nome,
                "Loja ID": loja_id,
                "transactionId": item.get("transactionId"),
                "userName": item.get("userName"),
                "userDocument": item.get("userDocument"),
                "productsValue_centavos": item.get("productsValue"),
                "tipValue_centavos": item.get("tipValue"),
                "productsValue": float(item.get("productsValue", 0) or 0) / 100,
                "tipValue": float(item.get("tipValue", 0) or 0) / 100,
                "total_com_gorjeta": (
                    float(item.get("productsValue", 0) or 0)
                    + float(item.get("tipValue", 0) or 0)
                ) / 100
            })

    if not registros_compradores:
        st.warning("Nenhum comprador encontrado no período.")
        st.stop()

    df = pd.DataFrame(registros_compradores)

    df["productsValue"] = pd.to_numeric(df["productsValue"], errors="coerce").fillna(0)
    df["tipValue"] = pd.to_numeric(df["tipValue"], errors="coerce").fillna(0)
    df["total_com_gorjeta"] = pd.to_numeric(df["total_com_gorjeta"], errors="coerce").fillna(0)

    resumo_loja = (
        df.groupby(["Loja", "Loja ID"], as_index=False)
        .agg({
            "productsValue": "sum",
            "tipValue": "sum",
            "total_com_gorjeta": "sum",
            "transactionId": "nunique"
        })
        .rename(columns={
            "productsValue": "Total Produtos",
            "tipValue": "Total Gorjeta",
            "total_com_gorjeta": "Total Produtos + Gorjeta",
            "transactionId": "Qtde Transações"
        })
    )

    total_produtos = df["productsValue"].sum()
    total_gorjeta = df["tipValue"].sum()
    total_com_gorjeta = df["total_com_gorjeta"].sum()

    def brl(v):
        return (
            f"R$ {v:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    col1, col2, col3 = st.columns(3)

    col1.metric("Produtos", brl(total_produtos))
    col2.metric("Gorjeta", brl(total_gorjeta))
    col3.metric("Produtos + Gorjeta", brl(total_com_gorjeta))

    st.subheader("Resumo por Loja")
    st.dataframe(resumo_loja, use_container_width=True, hide_index=True)

    st.subheader("Detalhe por Comprador / Transação")
    st.dataframe(df, use_container_width=True, hide_index=True)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False, sheet_name="Compradores Detalhe")
        resumo_loja.to_excel(writer, index=False, sheet_name="Resumo Gorjeta Loja")

    output.seek(0)

    st.download_button(
        label="📥 Baixar Teste Gorjeta",
        data=output,
        file_name=f"zig_teste_gorjeta_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
