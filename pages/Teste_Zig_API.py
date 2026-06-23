import streamlit as st
import requests
import pandas as pd
from datetime import date, timedelta

st.set_page_config(page_title="Teste Campos Faturamento ZIG", layout="wide")

st.title("🧪 ZIG - Ver todos os campos do faturamento")

token = st.secrets["zig"]["token"]
rede = st.secrets["zig"]["rede"]

headers = {
    "Authorization": token,
    "Accept": "application/json"
}

base_url = "https://api.zigcore.com.br/integration"


def gerar_periodos_5_dias(data_inicio, data_fim):
    periodos = []
    atual = data_inicio

    while atual <= data_fim:
        fim_bloco = min(atual + timedelta(days=4), data_fim)
        periodos.append((atual, fim_bloco))
        atual = fim_bloco + timedelta(days=1)

    return periodos


col1, col2 = st.columns(2)

with col1:
    dtinicio = st.date_input("Data início", value=date.today() - timedelta(days=5))

with col2:
    dtfim = st.date_input("Data fim", value=date.today() - timedelta(days=1))


if st.button("🔍 Buscar todos os campos do faturamento"):

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
    periodos = gerar_periodos_5_dias(dtinicio, dtfim)

    todos_registros = []
    lojas_sem_movimento = []

    for loja in lojas:
        loja_id = loja.get("id")
        loja_nome = loja.get("name")
        teve_movimento = False

        for inicio_bloco, fim_bloco in periodos:
            resp = requests.get(
                f"{base_url}/erp/faturamento",
                headers=headers,
                params={
                    "dtinicio": inicio_bloco.strftime("%Y-%m-%d"),
                    "dtfim": fim_bloco.strftime("%Y-%m-%d"),
                    "loja": loja_id
                },
                timeout=60
            )

            if resp.status_code != 200:
                continue

            dados = resp.json()

            if not isinstance(dados, list):
                continue

            if len(dados) > 0:
                teve_movimento = True

            for item in dados:
                item["loja_nome_consulta"] = loja_nome
                item["loja_id_consulta"] = loja_id
                item["periodo_inicio_consulta"] = inicio_bloco.strftime("%Y-%m-%d")
                item["periodo_fim_consulta"] = fim_bloco.strftime("%Y-%m-%d")
                todos_registros.append(item)

        if not teve_movimento:
            lojas_sem_movimento.append(loja_nome)

    if lojas_sem_movimento:
        st.info(
            f"ℹ️ {len(lojas_sem_movimento)} loja(s) sem movimentação: "
            + ", ".join(lojas_sem_movimento)
        )

    if not todos_registros:
        st.warning("Nenhum registro encontrado.")
        st.stop()

    df = pd.json_normalize(todos_registros)

    st.success(f"{len(df)} registros encontrados.")

    st.markdown("### Colunas retornadas pela API")
    st.write(list(df.columns))

    st.markdown("### Tabela completa")
    st.dataframe(df, use_container_width=True, hide_index=True)

    csv = df.to_csv(index=False, sep=";").encode("utf-8-sig")

    st.download_button(
        label="📥 Baixar CSV com todos os campos",
        data=csv,
        file_name="zig_faturamento_todos_campos.csv",
        mime="text/csv"
    )
