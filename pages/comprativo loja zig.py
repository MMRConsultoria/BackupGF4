import streamlit as st
import requests
import pandas as pd
from datetime import date, datetime, timedelta
from io import BytesIO
import unicodedata

st.set_page_config(page_title="Comparativo ZIG", layout="wide")
st.title("🧪 ZIG - Meio de Pagamento ZIG")

token = st.secrets["zig"]["token"]
rede = st.secrets["zig"]["rede"]

headers = {
    "Authorization": token,
    "Accept": "application/json"
}

base_url = "https://api.zigcore.com.br/integration"


def gerar_periodos_1_dia(data_inicio, data_fim):
    periodos = []
    atual = data_inicio

    while atual <= data_fim:
        periodos.append((atual, atual))
        atual = atual + timedelta(days=1)

    return periodos


def normalizar(txt):
    if txt is None:
        return ""

    txt = str(txt).strip().lower()

    txt = unicodedata.normalize("NFKD", txt)
    txt = "".join(c for c in txt if not unicodedata.combining(c))

    return txt


def eh_credito_debito(payment_name):
    nome = normalizar(payment_name)

    palavras = [
        "credito",
        "credit",
        "debito",
        "debit",
        "cartao",
        "card"
    ]

    return any(p in nome for p in palavras)


def brl(v):
    return (
        f"R$ {v:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


col1, col2 = st.columns(2)

with col1:
    dtinicio = st.date_input("Data início", value=date.today() - timedelta(days=7))

with col2:
    dtfim = st.date_input("Data fim", value=date.today() - timedelta(days=1))


if st.button("🔍 Gerar Meio de Pagamento ZIG"):

    resp_lojas = requests.get(
        f"{base_url}/erp/lojas",
        headers=headers,
        params={"rede": rede},
        timeout=30
    )

    if resp_lojas.status_code != 200:
        st.error("Erro ao buscar lojas")
        st.stop()

    lojas = resp_lojas.json()
    periodos = gerar_periodos_1_dia(dtinicio, dtfim)

    faturamento_geral = []
    maquina_integrada = []

    for loja in lojas:
        loja_id = loja.get("id")
        loja_nome = loja.get("name")

        for inicio, fim in periodos:

            resp_fat = requests.get(
                f"{base_url}/erp/faturamento",
                headers=headers,
                params={
                    "dtinicio": inicio.strftime("%Y-%m-%d"),
                    "dtfim": fim.strftime("%Y-%m-%d"),
                    "loja": loja_id
                },
                timeout=60
            )

            if resp_fat.status_code == 200:
                dados_fat = resp_fat.json()

                if isinstance(dados_fat, list):
                    for item in dados_fat:
                        valor = float(item.get("value", 0) or 0) / 100

                        faturamento_geral.append({
                            "Data": inicio.strftime("%d/%m/%Y"),
                            "Loja": loja_nome,
                            "Loja ID": loja_id,
                            "Origem": "Faturamento Geral",
                            "paymentId": item.get("paymentId"),
                            "Meio de Pagamento": item.get("paymentName"),
                            "Bandeira": "",
                            "Valor": valor,
                            "redeId": item.get("redeId"),
                            "lojaId": item.get("lojaId"),
                            "eventId": item.get("eventId"),
                            "eventDate": item.get("eventDate")
                        })

            resp_maq = requests.get(
                f"{base_url}/erp/faturamento/detalhesMaquinaIntegrada",
                headers=headers,
                params={
                    "dtinicio": inicio.strftime("%Y-%m-%d"),
                    "dtfim": fim.strftime("%Y-%m-%d"),
                    "loja": loja_id
                },
                timeout=60
            )

            if resp_maq.status_code == 200:
                dados_maq = resp_maq.json()

                if isinstance(dados_maq, list):
                    for item in dados_maq:
                        values = item.get("values", [])

                        if isinstance(values, list):
                            for v in values:
                                valor = float(v.get("totalValue", 0) or 0) / 100

                                maquina_integrada.append({
                                    "Data": inicio.strftime("%d/%m/%Y"),
                                    "Loja": loja_nome,
                                    "Loja ID": loja_id,
                                    "Origem": "Máquina Integrada",
                                    "paymentId": item.get("paymentId"),
                                    "Meio de Pagamento": item.get("paymentName"),
                                    "Bandeira": v.get("cardBrand"),
                                    "Valor": valor,
                                    "lojaId": item.get("lojaId"),
                                    "eventId": item.get("eventId")
                                })

    if not faturamento_geral:
        st.warning("Nenhum dado encontrado em Faturamento Geral.")
        st.stop()

    df_fat = pd.DataFrame(faturamento_geral)

    if maquina_integrada:
        df_maq = pd.DataFrame(maquina_integrada)
    else:
        df_maq = pd.DataFrame(columns=[
            "Data", "Loja", "Loja ID", "Origem", "paymentId",
            "Meio de Pagamento", "Bandeira", "Valor", "lojaId", "eventId"
        ])

    # ==========================================
    # REGRA FINAL:
    # Crédito/Débito vem da Máquina Integrada
    # Demais meios vêm do Faturamento Geral
    # ==========================================

    df_fat["Eh Cartao"] = df_fat["Meio de Pagamento"].apply(eh_credito_debito)
    df_maq["Eh Cartao"] = df_maq["Meio de Pagamento"].apply(eh_credito_debito)

    # Cartão pela máquina
    df_cartao = df_maq[df_maq["Eh Cartao"] == True].copy()

    # Não cartão pelo faturamento geral
    df_nao_cartao = df_fat[df_fat["Eh Cartao"] == False].copy()

    df_meio_pagamento = pd.concat(
        [df_cartao, df_nao_cartao],
        ignore_index=True
    )

    df_meio_pagamento["Valor"] = pd.to_numeric(
        df_meio_pagamento["Valor"],
        errors="coerce"
    ).fillna(0).round(2)

    df_meio_pagamento = df_meio_pagamento[
        [
            "Data",
            "Loja",
            "Loja ID",
            "Origem",
            "paymentId",
            "Meio de Pagamento",
            "Bandeira",
            "Valor"
        ]
    ]

    resumo_meio_pagamento = (
        df_meio_pagamento
        .groupby(
            [
                "Data",
                "Loja",
                "Origem",
                "Meio de Pagamento",
                "Bandeira"
            ],
            as_index=False
        )
        .agg({"Valor": "sum"})
    )

    resumo_meio_pagamento["Valor"] = resumo_meio_pagamento["Valor"].round(2)

    # Comparativo de conferência
    resumo_fat = (
        df_fat.groupby(["Data", "Loja"], as_index=False)
        .agg({"Valor": "sum"})
        .rename(columns={"Valor": "Valor Geral"})
    )

    resumo_final = (
        df_meio_pagamento.groupby(["Data", "Loja"], as_index=False)
        .agg({"Valor": "sum"})
        .rename(columns={"Valor": "Valor Meio Pagamento"})
    )

    comparativo = resumo_fat.merge(
        resumo_final,
        on=["Data", "Loja"],
        how="left"
    )

    comparativo["Valor Meio Pagamento"] = pd.to_numeric(
        comparativo["Valor Meio Pagamento"],
        errors="coerce"
    ).fillna(0)

    comparativo["Diferença"] = (
        comparativo["Valor Geral"] - comparativo["Valor Meio Pagamento"]
    ).round(2)

    comparativo["Valor Geral"] = comparativo["Valor Geral"].round(2)
    comparativo["Valor Meio Pagamento"] = comparativo["Valor Meio Pagamento"].round(2)

    total_geral = comparativo["Valor Geral"].sum()
    total_mp = comparativo["Valor Meio Pagamento"].sum()
    total_diferenca = comparativo["Diferença"].sum()

    col1, col2, col3 = st.columns(3)

    col1.metric("Faturamento Geral", brl(total_geral))
    col2.metric("Tabela Meio Pagamento", brl(total_mp))
    col3.metric("Diferença", brl(total_diferenca))

    st.subheader("Tabela Meio de Pagamento Final")
    st.dataframe(resumo_meio_pagamento, use_container_width=True)

    st.subheader("Comparativo de Conferência")
    st.dataframe(comparativo, use_container_width=True)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_fat.drop(columns=["Eh Cartao"], errors="ignore").to_excel(
            writer,
            index=False,
            sheet_name="Faturamento Geral"
        )

        df_maq.drop(columns=["Eh Cartao"], errors="ignore").to_excel(
            writer,
            index=False,
            sheet_name="Maquina Integrada"
        )

        df_meio_pagamento.to_excel(
            writer,
            index=False,
            sheet_name="Base Meio Pagamento"
        )

        resumo_meio_pagamento.to_excel(
            writer,
            index=False,
            sheet_name="Resumo Meio Pagamento"
        )

        comparativo.to_excel(
            writer,
            index=False,
            sheet_name="Comparativo"
        )

    output.seek(0)

    st.download_button(
        label="📥 Baixar Meio de Pagamento ZIG",
        data=output,
        file_name=f"zig_meio_pagamento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
