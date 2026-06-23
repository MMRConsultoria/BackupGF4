import streamlit as st
import requests
import pandas as pd
from datetime import date, datetime, timedelta
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

st.set_page_config(page_title="Teste ZIG Final", layout="wide")

st.title("🧪 Teste ZIG - Padrão Final")

# Google Sheets
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)

planilha_empresa = gc.open("Vendas diarias")
aba_empresa = planilha_empresa.worksheet("Tabela Empresa")
valores_empresa = aba_empresa.get_all_values()

df_empresa = pd.DataFrame(valores_empresa[1:], columns=valores_empresa[0])
df_empresa.columns = df_empresa.columns.str.strip()
df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.lower().str.strip()

# ZIG
token = st.secrets["zig"]["token"]
rede = st.secrets["zig"]["rede"]

headers = {
    "Authorization": token,
    "Accept": "application/json"
}

base_url = "https://api.zigcore.com.br/integration"

col1, col2 = st.columns(2)

with col1:
    dtinicio = st.date_input("Data início", value=date.today() - timedelta(days=99))

with col2:
    dtfim = st.date_input("Data fim", value=date.today() - timedelta(days=1))

if st.button("🔄 Atualizar ZIG - Teste Final"):

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

    registros = []
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
                registros.append({
                    "Data": item.get("eventDate"),
                    "Loja": loja_nome,
                    "Fat.Total": float(item.get("value", 0)) / 100
                })

    if not registros:
        st.warning("Nenhum faturamento encontrado.")
        if erros:
            st.warning("Algumas lojas retornaram erro:")
            st.write(erros)
        st.stop()

    df = pd.DataFrame(registros)

    df["Data"] = pd.to_datetime(df["Data"], errors="coerce")
    df["Loja"] = df["Loja"].astype(str).str.strip().str.lower()
    df["Fat.Total"] = pd.to_numeric(df["Fat.Total"], errors="coerce").fillna(0)

    resumo = (
        df.groupby(["Data", "Loja"], as_index=False)
        .agg({"Fat.Total": "sum"})
    )

    resumo["Serv/Tx"] = 0
    resumo["Fat.Real"] = resumo["Fat.Total"]
    resumo["Ticket"] = 0

    dias_traducao = {
        "Monday": "segunda-feira",
        "Tuesday": "terça-feira",
        "Wednesday": "quarta-feira",
        "Thursday": "quinta-feira",
        "Friday": "sexta-feira",
        "Saturday": "sábado",
        "Sunday": "domingo"
    }

    resumo.insert(
        1,
        "Dia da Semana",
        resumo["Data"].dt.day_name().map(dias_traducao)
    )

    meses = {
        "jan": "jan", "feb": "fev", "mar": "mar", "apr": "abr",
        "may": "mai", "jun": "jun", "jul": "jul", "aug": "ago",
        "sep": "set", "oct": "out", "nov": "nov", "dec": "dez"
    }

    resumo["Mês"] = resumo["Data"].dt.strftime("%b").str.lower().map(meses)
    resumo["Ano"] = resumo["Data"].dt.year
    resumo["Sistema"] = "ZIG"

    resumo = resumo.merge(
        df_empresa[["Loja", "Código Everest", "Grupo", "Código Grupo Everest"]],
        on="Loja",
        how="left"
    )

    resumo["Data"] = resumo["Data"].dt.strftime("%d/%m/%Y")

    colunas_finais = [
        "Data", "Dia da Semana", "Loja", "Código Everest", "Grupo",
        "Código Grupo Everest", "Fat.Total", "Serv/Tx", "Fat.Real",
        "Ticket", "Mês", "Ano", "Sistema"
    ]

    resumo = resumo[colunas_finais]

    for col in ["Fat.Total", "Serv/Tx", "Fat.Real", "Ticket"]:
        resumo[col] = pd.to_numeric(resumo[col], errors="coerce").fillna(0).round(2)

    resumo["Data_Ordenada"] = pd.to_datetime(resumo["Data"], format="%d/%m/%Y")
    resumo = resumo.sort_values(["Data_Ordenada", "Loja"]).drop(columns="Data_Ordenada")

    lojas_nao_localizadas = resumo[
        resumo["Código Everest"].isna() |
        (resumo["Código Everest"].astype(str).str.strip() == "")
    ]["Loja"].unique()

    if len(lojas_nao_localizadas) > 0:
        st.error("❌ Lojas ZIG não localizadas na Tabela Empresa:")
        st.write(lojas_nao_localizadas)
        st.stop()

    st.success("✅ Todas as lojas foram localizadas na Tabela Empresa.")

    datas_validas = pd.to_datetime(resumo["Data"], format="%d/%m/%Y", errors="coerce").dropna()

    if not datas_validas.empty:
        data_inicial = datas_validas.min().strftime("%d/%m/%Y")
        data_final = datas_validas.max().strftime("%d/%m/%Y")
        total = resumo["Fat.Total"].sum()

        total_formatado = (
            f"R$ {total:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

        col1, col2 = st.columns(2)

        with col1:
            st.metric("Período", f"{data_inicial} até {data_final}")

        with col2:
            st.metric("Valor total", total_formatado)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        resumo.to_excel(writer, index=False, sheet_name="Faturamento Servico")

    output.seek(0)

    st.download_button(
        label="📥 Baixar Excel ZIG",
        data=output,
        file_name=f"zig_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    if erros:
        st.warning("Algumas lojas retornaram erro:")
        st.write(erros)
