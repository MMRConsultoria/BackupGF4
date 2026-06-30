import streamlit as st
import requests
import pandas as pd
from datetime import date, datetime, timedelta
from io import BytesIO
import unicodedata

st.set_page_config(page_title="ZIG - Faturamento Meio Pagamento", layout="wide")
st.title("🧪 ZIG - Faturamento Meio de Pagamento")

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


if st.button("🔍 Gerar Faturamento Meio Pagamento"):

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
    periodos = gerar_periodos_1_dia(dtinicio, dtfim)

    faturamento_geral = []
    maquina_integrada = []

    for loja in lojas:
        loja_id = loja.get("id")
        loja_nome = loja.get("name")

        for inicio, fim in periodos:

            # ==========================
            # FATURAMENTO GERAL
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

            # ==========================
            # MÁQUINA INTEGRADA
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
            "Data",
            "Loja",
            "Loja ID",
            "Origem",
            "paymentId",
            "Meio de Pagamento",
            "Bandeira",
            "Valor",
            "lojaId",
            "eventId"
        ])

    # ==========================================
    # REGRA FINAL
    # Crédito/Débito = Máquina Integrada
    # Demais meios = Faturamento Geral
    # ==========================================

    df_fat["Eh Cartao"] = df_fat["Meio de Pagamento"].apply(eh_credito_debito)
    df_maq["Eh Cartao"] = df_maq["Meio de Pagamento"].apply(eh_credito_debito)

    df_cartao = df_maq[df_maq["Eh Cartao"] == True].copy()
    df_outros = df_fat[df_fat["Eh Cartao"] == False].copy()

    df_cartao["Origem"] = "Máquina Integrada"
    df_outros["Origem"] = "Faturamento Geral"

    df_final = pd.concat(
        [df_cartao, df_outros],
        ignore_index=True
    )

    if df_final.empty:
        st.warning("Nenhum dado encontrado para montar a tabela de meio de pagamento.")
        st.stop()

    df_final["Valor"] = pd.to_numeric(
        df_final["Valor"],
        errors="coerce"
    ).fillna(0)

    df_final["Bandeira"] = df_final["Bandeira"].fillna("").astype(str).str.upper()

    df_final["Data Convertida"] = pd.to_datetime(
        df_final["Data"],
        format="%d/%m/%Y",
        errors="coerce"
    )

    mapa_dias = {
        "Monday": "Segunda-feira",
        "Tuesday": "Terça-feira",
        "Wednesday": "Quarta-feira",
        "Thursday": "Quinta-feira",
        "Friday": "Sexta-feira",
        "Saturday": "Sábado",
        "Sunday": "Domingo"
    }

    df_final["Dia da Semana"] = df_final["Data Convertida"].dt.day_name().map(mapa_dias)
    df_final["Mês"] = df_final["Data Convertida"].dt.month
    df_final["Ano"] = df_final["Data Convertida"].dt.year

    # ==========================================
    # TABELA FINAL NO PADRÃO MEIO DE PAGAMENTO
    # ==========================================

    tabela_meio_pagamento = (
        df_final
        .groupby(
            [
                "Data",
                "Dia da Semana",
                "Mês",
                "Ano",
                "Loja",
                "Loja ID",
                "Meio de Pagamento",
                "Bandeira",
                "Origem"
            ],
            as_index=False
        )
        .agg({"Valor": "sum"})
    )

    tabela_meio_pagamento["Valor"] = tabela_meio_pagamento["Valor"].round(2)

    tabela_meio_pagamento = tabela_meio_pagamento.rename(columns={
        "Loja ID": "Cód Loja"
    })

    tabela_meio_pagamento = tabela_meio_pagamento[
        [
            "Data",
            "Dia da Semana",
            "Mês",
            "Ano",
            "Loja",
            "Cód Loja",
            "Meio de Pagamento",
            "Bandeira",
            "Origem",
            "Valor"
        ]
    ]

    total_mp = tabela_meio_pagamento["Valor"].sum()

    st.metric("Total Faturamento Meio Pagamento", brl(total_mp))

    st.subheader("Faturamento Meio de Pagamento")
    st.dataframe(tabela_meio_pagamento, use_container_width=True, hide_index=True)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        tabela_meio_pagamento.to_excel(
            writer,
            index=False,
            sheet_name="Faturamento Meio Pagamento"
        )

        df_cartao.drop(columns=["Eh Cartao"], errors="ignore").to_excel(
            writer,
            index=False,
            sheet_name="Credito Debito Maquina"
        )

        df_outros.drop(columns=["Eh Cartao"], errors="ignore").to_excel(
            writer,
            index=False,
            sheet_name="Outros Faturamento"
        )

        df_fat.drop(columns=["Eh Cartao"], errors="ignore").to_excel(
            writer,
            index=False,
            sheet_name="Base Faturamento Geral"
        )

        df_maq.drop(columns=["Eh Cartao"], errors="ignore").to_excel(
            writer,
            index=False,
            sheet_name="Base Maquina Integrada"
        )

    output.seek(0)

    st.download_button(
        label="📥 Baixar Faturamento Meio Pagamento",
        data=output,
        file_name=f"zig_faturamento_meio_pagamento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
