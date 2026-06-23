import streamlit as st
import requests
import pandas as pd
import numpy as np
from datetime import date, datetime, timedelta
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json

st.set_page_config(page_title="Teste ZIG Produtos", layout="wide")
st.title("🧪 Teste ZIG - Faturamento + Produtos + Ticket")

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
    dtinicio = st.date_input("Data início", value=date.today() - timedelta(days=30))

with col2:
    dtfim = st.date_input("Data fim", value=date.today() - timedelta(days=1))

if st.button("🔄 Testar Faturamento + Produtos"):

    if dtinicio > dtfim:
        st.error("A data inicial não pode ser maior que a data final.")
        st.stop()

    resp_lojas = requests.get(
        f"{base_url}/erp/lojas",
        headers=headers,
        params={"rede": rede},
        timeout=30
    )

    if resp_lojas.status_code != 200:
        st.error("Erro ao buscar lojas ZIG")
        st.stop()

    lojas = resp_lojas.json()
    periodos = gerar_periodos_5_dias(dtinicio, dtfim)

    registros_faturamento = []
    registros_produtos = []
    lojas_sem_movimento = []

    for loja in lojas:
        loja_id = loja.get("id")
        loja_nome = loja.get("name")
        teve_movimento = False

        for inicio_bloco, fim_bloco in periodos:

            resp_fat = requests.get(
                f"{base_url}/erp/faturamento",
                headers=headers,
                params={
                    "dtinicio": inicio_bloco.strftime("%Y-%m-%d"),
                    "dtfim": fim_bloco.strftime("%Y-%m-%d"),
                    "loja": loja_id
                },
                timeout=60
            )

            if resp_fat.status_code == 200:
                dados_fat = resp_fat.json()

                if isinstance(dados_fat, list) and len(dados_fat) > 0:
                    teve_movimento = True

                    for item in dados_fat:
                        registros_faturamento.append({
                            "Data": item.get("eventDate"),
                            "Loja": loja_nome,
                            "Fat.Total": float(item.get("value", 0) or 0) / 100
                        })

            resp_prod = requests.get(
                f"{base_url}/erp/saida-produtos",
                headers=headers,
                params={
                    "dtinicio": inicio_bloco.strftime("%Y-%m-%d"),
                    "dtfim": fim_bloco.strftime("%Y-%m-%d"),
                    "loja": loja_id
                },
                timeout=60
            )

            if resp_prod.status_code == 200:
                dados_prod = resp_prod.json()

                if isinstance(dados_prod, list) and len(dados_prod) > 0:
                    teve_movimento = True

                    for item in dados_prod:
                        unit_value = float(item.get("unitValue", 0) or 0)
                        count = float(item.get("count", 0) or 0)
                        fractional_amount = item.get("fractionalAmount")
                        discount_value = float(item.get("discountValue", 0) or 0)

                        qtd = count
                        if fractional_amount not in [None, "", 0]:
                            try:
                                qtd = float(fractional_amount)
                            except Exception:
                                qtd = count

                        valor = ((unit_value * qtd) - discount_value) / 100

                        registros_produtos.append({
                            "Data": item.get("eventDate"),
                            "Loja": loja_nome,
                            "Tipo": item.get("type"),
                            "Produto": item.get("productName"),
                            "Categoria": item.get("productCategory"),
                            "TransactionId": item.get("transactionId"),
                            "Valor Produto": valor
                        })

        if not teve_movimento:
            lojas_sem_movimento.append(loja_nome)

    if lojas_sem_movimento:
        st.info(
            f"ℹ️ {len(lojas_sem_movimento)} loja(s) sem movimentação no período: "
            + ", ".join(lojas_sem_movimento)
        )

    if not registros_faturamento:
        st.warning("⚠️ Nenhum faturamento encontrado.")
        st.stop()

    df_fat = pd.DataFrame(registros_faturamento)
    df_fat["Data"] = pd.to_datetime(df_fat["Data"], errors="coerce").dt.strftime("%Y-%m-%d")
    df_fat["Loja"] = df_fat["Loja"].astype(str).str.strip().str.lower()
    df_fat["Fat.Total"] = pd.to_numeric(df_fat["Fat.Total"], errors="coerce").fillna(0)

    resumo_fat = (
        df_fat.groupby(["Data", "Loja"], as_index=False)
        .agg({"Fat.Total": "sum"})
    )

    if registros_produtos:
        df_prod = pd.DataFrame(registros_produtos)
        df_prod["Data"] = pd.to_datetime(df_prod["Data"], errors="coerce").dt.strftime("%Y-%m-%d")
        df_prod["Loja"] = df_prod["Loja"].astype(str).str.strip().str.lower()
        df_prod["Tipo"] = df_prod["Tipo"].astype(str).str.strip().str.lower()
        df_prod["Valor Produto"] = pd.to_numeric(df_prod["Valor Produto"], errors="coerce").fillna(0)
        df_prod["TransactionId"] = df_prod["TransactionId"].astype(str).str.strip()

        df_serv = df_prod[
            df_prod["Tipo"].isin(["tip", "couvert", "entrance"])
        ].copy()

        resumo_serv = (
            df_serv.groupby(["Data", "Loja"], as_index=False)
            .agg({"Valor Produto": "sum"})
            .rename(columns={"Valor Produto": "Serv/Tx"})
        )

        ticket_df = (
            df_prod.groupby(["Data", "Loja"], as_index=False)
            .agg(Qtde_Transacoes=("TransactionId", "nunique"))
        )

        resumo_tipo = (
            df_prod.groupby(["Data", "Loja", "Tipo"], as_index=False)
            .agg({"Valor Produto": "sum"})
        )

    else:
        df_prod = pd.DataFrame()
        resumo_serv = pd.DataFrame(columns=["Data", "Loja", "Serv/Tx"])
        ticket_df = pd.DataFrame(columns=["Data", "Loja", "Qtde_Transacoes"])
        resumo_tipo = pd.DataFrame()

    resumo_serv["Data"] = pd.to_datetime(resumo_serv["Data"], errors="coerce").dt.strftime("%Y-%m-%d")
    resumo_serv["Loja"] = resumo_serv["Loja"].astype(str).str.strip().str.lower()

    ticket_df["Data"] = pd.to_datetime(ticket_df["Data"], errors="coerce").dt.strftime("%Y-%m-%d")
    ticket_df["Loja"] = ticket_df["Loja"].astype(str).str.strip().str.lower()

    resumo = resumo_fat.merge(
        resumo_serv,
        on=["Data", "Loja"],
        how="left"
    )

    resumo = resumo.merge(
        ticket_df,
        on=["Data", "Loja"],
        how="left"
    )

    resumo["Serv/Tx"] = pd.to_numeric(resumo["Serv/Tx"], errors="coerce").fillna(0)
    resumo["Qtde_Transacoes"] = pd.to_numeric(resumo["Qtde_Transacoes"], errors="coerce").fillna(0)

    resumo["Fat.Real"] = resumo["Fat.Total"] - resumo["Serv/Tx"]

    resumo["Ticket"] = np.where(
        resumo["Qtde_Transacoes"] > 0,
        resumo["Fat.Total"] / resumo["Qtde_Transacoes"],
        0
    )

    resumo["Data"] = pd.to_datetime(resumo["Data"], errors="coerce")

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

    datas_validas = pd.to_datetime(resumo["Data"], format="%d/%m/%Y", errors="coerce").dropna()

    if not datas_validas.empty:
        data_inicial = datas_validas.min().strftime("%d/%m/%Y")
        data_final = datas_validas.max().strftime("%d/%m/%Y")

        total_fat = resumo["Fat.Total"].sum()
        total_serv = resumo["Serv/Tx"].sum()
        total_real = resumo["Fat.Real"].sum()
        ticket_medio = total_fat / resumo["Ticket"].replace(0, np.nan).count() if resumo["Ticket"].replace(0, np.nan).count() > 0 else 0

        def brl(valor):
            return (
                f"R$ {valor:,.2f}"
                .replace(",", "X")
                .replace(".", ",")
                .replace("X", ".")
            )

        col1, col2, col3, col4 = st.columns(4)

        with col1:
            st.metric("Período", f"{data_inicial} até {data_final}")

        with col2:
            st.metric("Fat.Total", brl(total_fat))

        with col3:
            st.metric("Serv/Tx", brl(total_serv))

        with col4:
            st.metric("Fat.Real", brl(total_real))

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        resumo.to_excel(writer, index=False, sheet_name="Faturamento Servico")

        if not df_prod.empty:
            df_prod.to_excel(writer, index=False, sheet_name="Produtos Detalhe")

        if not resumo_tipo.empty:
            resumo_tipo.to_excel(writer, index=False, sheet_name="Resumo Tipo Produto")

        if not ticket_df.empty:
            ticket_df.to_excel(writer, index=False, sheet_name="Transacoes")

    output.seek(0)

    st.download_button(
        label="📥 Baixar Excel ZIG Teste",
        data=output,
        file_name=f"zig_teste_produtos_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
