import streamlit as st
import requests
import pandas as pd
from datetime import date, datetime, timedelta
from io import BytesIO

st.set_page_config(page_title="Comparativo ZIG", layout="wide")
st.title("🧪 ZIG - Comparativo Faturamento x Máquina Integrada")

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


col1, col2 = st.columns(2)

with col1:
    dtinicio = st.date_input("Data início", value=date.today() - timedelta(days=7))

with col2:
    dtfim = st.date_input("Data fim", value=date.today() - timedelta(days=1))


if st.button("🔍 Comparar ZIG"):

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

            # ==========================
            # 1. FATURAMENTO GERAL
            # ==========================
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
                        faturamento_geral.append({
                            "Data Consulta": inicio.strftime("%d/%m/%Y"),
                            "Loja": loja_nome,
                            "Loja ID Consulta": loja_id,
                            "paymentId": item.get("paymentId"),
                            "paymentName": item.get("paymentName"),
                            "value_centavos": item.get("value"),
                            "Valor Geral": float(item.get("value", 0) or 0) / 100,
                            "redeId": item.get("redeId"),
                            "lojaId": item.get("lojaId"),
                            "eventId": item.get("eventId"),
                            "eventDate": item.get("eventDate")
                        })

            # ==========================
            # 2. DETALHE MÁQUINA
            # ==========================
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
                                maquina_integrada.append({
                                    "Data Consulta": inicio.strftime("%d/%m/%Y"),
                                    "Loja": loja_nome,
                                    "Loja ID Consulta": loja_id,
                                    "paymentId": item.get("paymentId"),
                                    "paymentName": item.get("paymentName"),
                                    "lojaId": item.get("lojaId"),
                                    "eventId": item.get("eventId"),
                                    "cardBrand": v.get("cardBrand"),
                                    "totalValue_centavos": v.get("totalValue"),
                                    "Valor Máquina": float(v.get("totalValue", 0) or 0) / 100
                                })

    if not faturamento_geral:
        st.warning("Nenhum dado encontrado em Faturamento Geral.")
        st.stop()

    df_fat = pd.DataFrame(faturamento_geral)

    if maquina_integrada:
        df_maq = pd.DataFrame(maquina_integrada)
    else:
        df_maq = pd.DataFrame(columns=[
            "Data Consulta", "Loja", "paymentId", "paymentName",
            "cardBrand", "Valor Máquina"
        ])

    resumo_fat = (
        df_fat.groupby(["Data Consulta", "Loja"], as_index=False)
        .agg({"Valor Geral": "sum"})
    )

    resumo_maq = (
        df_maq.groupby(["Data Consulta", "Loja"], as_index=False)
        .agg({"Valor Máquina": "sum"})
    )

    comparativo = resumo_fat.merge(
        resumo_maq,
        on=["Data Consulta", "Loja"],
        how="left"
    )

    comparativo["Valor Máquina"] = pd.to_numeric(
        comparativo["Valor Máquina"],
        errors="coerce"
    ).fillna(0)

    comparativo["Diferença"] = (
        comparativo["Valor Geral"] - comparativo["Valor Máquina"]
    ).round(2)

    comparativo["Valor Geral"] = comparativo["Valor Geral"].round(2)
    comparativo["Valor Máquina"] = comparativo["Valor Máquina"].round(2)

    total_geral = comparativo["Valor Geral"].sum()
    total_maquina = comparativo["Valor Máquina"].sum()
    total_diferenca = comparativo["Diferença"].sum()

    def brl(v):
        return (
            f"R$ {v:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

    col1, col2, col3 = st.columns(3)

    col1.metric("Faturamento Geral", brl(total_geral))
    col2.metric("Máquina Integrada", brl(total_maquina))
    col3.metric("Diferença", brl(total_diferenca))

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_fat.to_excel(writer, index=False, sheet_name="Faturamento Geral")
        df_maq.to_excel(writer, index=False, sheet_name="Maquina Integrada")
        comparativo.to_excel(writer, index=False, sheet_name="Comparativo")

    output.seek(0)

    st.download_button(
        label="📥 Baixar Comparativo ZIG",
        data=output,
        file_name=f"zig_comparativo_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
