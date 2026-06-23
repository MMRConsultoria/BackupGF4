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

# ======================================
# PERÍODO
# ======================================

col1, col2 = st.columns(2)

with col1:
    dtinicio = st.date_input(
        "Data início",
        value=date.today() - timedelta(days=30)
    )

with col2:
    dtfim = st.date_input(
        "Data fim",
        value=date.today()
    )

# ======================================
# BUSCAR FATURAMENTO
# ======================================

if st.button("🔄 Buscar faturamento ZIG"):

    with st.spinner("Buscando lojas..."):

        resp_lojas = requests.get(
            f"{base_url}/erp/lojas",
            headers=headers,
            params={"rede": rede},
            timeout=30
        )

        st.write("Status lojas:", resp_lojas.status_code)

        if resp_lojas.status_code != 200:
            st.error("Erro ao buscar lojas")
            st.text(resp_lojas.text)
            st.stop()

        lojas = resp_lojas.json()

        st.success(f"{len(lojas)} lojas encontradas")

    todos_registros = []
    erros = []

    barra = st.progress(0)

    for idx, loja in enumerate(lojas):

        loja_id = loja.get("id")
        loja_nome = loja.get("name")

        try:

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

            if resp.status_code != 200:

                erros.append({
                    "Loja": loja_nome,
                    "Status": resp.status_code,
                    "Erro": resp.text
                })

                continue

            dados = resp.json()

            if isinstance(dados, list):

                for item in dados:

                    todos_registros.append({
                        "Data": item.get("eventDate"),
                        "Loja": loja_nome,
                        "Faturamento": float(item.get("value", 0)) / 100
                    })

        except Exception as e:

            erros.append({
                "Loja": loja_nome,
                "Status": "EXCEPTION",
                "Erro": str(e)
            })

        barra.progress((idx + 1) / len(lojas))

    # ======================================
    # RESUMO
    # ======================================

    if len(todos_registros) > 0:

        df = pd.DataFrame(todos_registros)

        tabela_resumo = (
            df.groupby(
                ["Data", "Loja"],
                as_index=False
            )["Faturamento"]
            .sum()
        )

        tabela_resumo["Faturamento"] = (
            tabela_resumo["Faturamento"]
            .round(2)
        )

        total = tabela_resumo["Faturamento"].sum()

        st.success(
            f"{len(tabela_resumo)} linhas encontradas | Total: R$ {total:,.2f}"
        )

        st.dataframe(
            tabela_resumo,
            use_container_width=True,
            hide_index=True
        )

        # Excel

        excel = tabela_resumo.to_csv(
            index=False,
            sep=";"
        ).encode("utf-8-sig")

        st.download_button(
            "📥 Baixar CSV",
            excel,
            file_name="zig_faturamento.csv",
            mime="text/csv"
        )

    else:

        st.warning("Nenhum faturamento encontrado.")

    # ======================================
    # ERROS
    # ======================================

    if len(erros) > 0:

        st.warning(
            f"{len(erros)} loja(s) retornaram erro."
        )

        st.dataframe(
            pd.DataFrame(erros),
            use_container_width=True,
            hide_index=True
        )
