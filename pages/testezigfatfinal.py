import streamlit as st
import requests
import pandas as pd
from datetime import date, datetime, timedelta
from io import BytesIO

st.set_page_config(page_title="ZIG - Conferência Máquina x Faturamento", layout="wide")
st.title("🧪 ZIG - Faturamento Geral x Máquina Integrada")

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
        atual += timedelta(days=1)

    return periodos


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


if st.button("🔍 Comparar Faturamento x Máquina"):

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
    periodos = gerar_periodos_1_dia(dtinicio, dtfim)

    faturamento_geral = []
    maquina_integrada = []

    for loja in lojas:
        loja_id = loja.get("id")
        loja_nome = loja.get("name")

        for inicio, fim in periodos:

            # FATURAMENTO GERAL
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
                            "Data": inicio.strftime("%d/%m/%Y"),
                            "Loja": loja_nome,
                            "Loja ID": loja_id,
                            "Origem": "Faturamento Geral",
                            "paymentId": item.get("paymentId"),
                            "paymentName": item.get("paymentName"),
                            "Valor Faturamento": float(item.get("value", 0) or 0) / 100,
                            "value_centavos": item.get("value"),
                            "redeId": item.get("redeId"),
                            "lojaId": item.get("lojaId"),
                            "eventId": item.get("eventId"),
                            "eventDate": item.get("eventDate")
                        })

            # MÁQUINA INTEGRADA
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
                                total_value = float(v.get("totalValue", 0) or 0) / 100

                                maquina_integrada.append({
                                    "Data": inicio.strftime("%d/%m/%Y"),
                                    "Loja": loja_nome,
                                    "Loja ID": loja_id,
                                    "Origem": "Máquina Integrada",
                                    "paymentId": item.get("paymentId"),
                                    "paymentName": item.get("paymentName"),
                                    "cardBrand": v.get("cardBrand"),
                                    "Valor Máquina": total_value,
                                    "totalValue_centavos": v.get("totalValue"),
                                    "lojaId": item.get("lojaId"),
                                    "eventId": item.get("eventId")
                                })

    if not faturamento_geral:
        st.warning("Nenhum dado encontrado no faturamento geral.")
        st.stop()

    df_fat = pd.DataFrame(faturamento_geral)

    if maquina_integrada:
        df_maq = pd.DataFrame(maquina_integrada)
    else:
        df_maq = pd.DataFrame(columns=[
            "Data", "Loja", "Loja ID", "Origem", "paymentId",
            "paymentName", "cardBrand", "Valor Máquina"
        ])

    resumo_fat_pagamento = (
        df_fat
        .groupby(["Data", "Loja", "paymentId", "paymentName"], as_index=False)
        .agg({"Valor Faturamento": "sum"})
    )

    resumo_maq_pagamento = (
        df_maq
        .groupby(["Data", "Loja", "paymentId", "paymentName"], as_index=False)
        .agg({"Valor Máquina": "sum"})
    )

    comparativo_pagamento = resumo_fat_pagamento.merge(
        resumo_maq_pagamento,
        on=["Data", "Loja", "paymentId", "paymentName"],
        how="outer"
    )

    comparativo_pagamento["Valor Faturamento"] = pd.to_numeric(
        comparativo_pagamento["Valor Faturamento"],
        errors="coerce"
    ).fillna(0)

    comparativo_pagamento["Valor Máquina"] = pd.to_numeric(
        comparativo_pagamento["Valor Máquina"],
        errors="coerce"
    ).fillna(0)

    comparativo_pagamento["Diferença Fat - Máquina"] = (
        comparativo_pagamento["Valor Faturamento"] -
        comparativo_pagamento["Valor Máquina"]
    ).round(2)

    resumo_fat_loja = (
        df_fat
        .groupby(["Data", "Loja"], as_index=False)
        .agg({"Valor Faturamento": "sum"})
    )

    resumo_maq_loja = (
        df_maq
        .groupby(["Data", "Loja"], as_index=False)
        .agg({"Valor Máquina": "sum"})
    )

    comparativo_loja = resumo_fat_loja.merge(
        resumo_maq_loja,
        on=["Data", "Loja"],
        how="outer"
    )

    comparativo_loja["Valor Faturamento"] = pd.to_numeric(
        comparativo_loja["Valor Faturamento"],
        errors="coerce"
    ).fillna(0)

    comparativo_loja["Valor Máquina"] = pd.to_numeric(
        comparativo_loja["Valor Máquina"],
        errors="coerce"
    ).fillna(0)

    comparativo_loja["Diferença Fat - Máquina"] = (
        comparativo_loja["Valor Faturamento"] -
        comparativo_loja["Valor Máquina"]
    ).round(2)

    total_fat = comparativo_loja["Valor Faturamento"].sum()
    total_maq = comparativo_loja["Valor Máquina"].sum()
    total_dif = comparativo_loja["Diferença Fat - Máquina"].sum()

    col1, col2, col3 = st.columns(3)

    col1.metric("Faturamento Geral", brl(total_fat))
    col2.metric("Máquina Integrada", brl(total_maq))
    col3.metric("Diferença", brl(total_dif))

    st.subheader("Comparativo por loja e dia")
    st.dataframe(comparativo_loja, use_container_width=True, hide_index=True)

    st.subheader("Comparativo por meio de pagamento")
    st.dataframe(comparativo_pagamento, use_container_width=True, hide_index=True)

    st.subheader("Base Faturamento Geral")
    st.dataframe(df_fat, use_container_width=True, hide_index=True)

    st.subheader("Base Máquina Integrada")
    st.dataframe(df_maq, use_container_width=True, hide_index=True)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_fat.to_excel(writer, index=False, sheet_name="Base Faturamento Geral")
        df_maq.to_excel(writer, index=False, sheet_name="Base Maquina Integrada")
        comparativo_loja.to_excel(writer, index=False, sheet_name="Comparativo Loja Dia")
        comparativo_pagamento.to_excel(writer, index=False, sheet_name="Comparativo Pagamento")

    output.seek(0)

    st.download_button(
        label="📥 Baixar Conferência ZIG",
        data=output,
        file_name=f"zig_conferencia_fat_x_maquina_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
