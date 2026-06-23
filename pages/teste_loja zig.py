import streamlit as st
import requests
import pandas as pd
from datetime import date, timedelta
from io import BytesIO

st.set_page_config(page_title="Teste ZIG", layout="wide")

st.title("🧪 Teste ZIG API - Padrão Vendas Diárias")

token = st.secrets["zig"]["token"]
rede = st.secrets["zig"]["rede"]

headers = {
    "Authorization": token,
    "Accept": "application/json"
}

base_url = "https://api.zigcore.com.br/integration"

col1, col2 = st.columns(2)

with col1:
    dtinicio = st.date_input("Data início", value=date.today() - timedelta(days=30))

with col2:
    dtfim = st.date_input("Data fim", value=date.today())

if st.button("🔄 Buscar faturamento ZIG no padrão"):

    resp_lojas = requests.get(
        f"{base_url}/erp/lojas",
        headers=headers,
        params={"rede": rede},
        timeout=30
    )

    if resp_lojas.status_code != 200:
        st.error("Erro ao buscar lojas")
        st.text(resp_lojas.text)
        st.stop()

    lojas = resp_lojas.json()

    todos_registros = []
    erros = []

    for loja in lojas:
        loja_id = loja.get("id")
        loja_nome = loja.get("name")

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
                    "Fat.Total": float(item.get("value", 0)) / 100
                })

    if not todos_registros:
        st.warning("Nenhum faturamento encontrado.")
        st.stop()

    df = pd.DataFrame(todos_registros)

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df["Loja"] = df["Loja"].astype(str).str.strip().str.lower()
    df["Fat.Total"] = pd.to_numeric(df["Fat.Total"], errors="coerce").fillna(0)

    df_final = (
        df.groupby(["Data", "Loja"], as_index=False)
        .agg({"Fat.Total": "sum"})
    )

    df_final["Serv/Tx"] = 0
    df_final["Fat.Real"] = df_final["Fat.Total"]
    df_final["Ticket"] = 0

    dias_traducao = {
        "Monday": "segunda-feira",
        "Tuesday": "terça-feira",
        "Wednesday": "quarta-feira",
        "Thursday": "quinta-feira",
        "Friday": "sexta-feira",
        "Saturday": "sábado",
        "Sunday": "domingo"
    }

    df_final.insert(
        1,
        "Dia da Semana",
        df_final["Data"].dt.day_name().map(dias_traducao)
    )

    meses = {
        "jan": "jan", "feb": "fev", "mar": "mar", "apr": "abr",
        "may": "mai", "jun": "jun", "jul": "jul", "aug": "ago",
        "sep": "set", "oct": "out", "nov": "nov", "dec": "dez"
    }

    df_final["Mês"] = df_final["Data"].dt.strftime("%b").str.lower().map(meses)
    df_final["Ano"] = df_final["Data"].dt.year
    df_final["Sistema"] = "ZIG"

    # Ainda sem merge com Tabela Empresa
    df_final["Código Everest"] = ""
    df_final["Grupo"] = ""
    df_final["Código Grupo Everest"] = ""

    df_final["Data"] = df_final["Data"].dt.strftime("%d/%m/%Y")

    colunas_finais = [
        "Data",
        "Dia da Semana",
        "Loja",
        "Código Everest",
        "Grupo",
        "Código Grupo Everest",
        "Fat.Total",
        "Serv/Tx",
        "Fat.Real",
        "Ticket",
        "Mês",
        "Ano",
        "Sistema"
    ]

    df_final = df_final[colunas_finais]

    for col in ["Fat.Total", "Serv/Tx", "Fat.Real", "Ticket"]:
        df_final[col] = pd.to_numeric(df_final[col], errors="coerce").fillna(0).round(2)

    st.success(f"{len(df_final)} linhas no padrão geradas.")

    st.dataframe(df_final, use_container_width=True, hide_index=True)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_final.to_excel(writer, index=False, sheet_name="Faturamento Servico")

    output.seek(0)

    st.download_button(
        label="📥 Baixar Excel no padrão",
        data=output,
        file_name="zig_faturamento_padrao.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if erros:
        st.warning("Algumas lojas retornaram erro:")
        st.dataframe(pd.DataFrame(erros), use_container_width=True, hide_index=True)
