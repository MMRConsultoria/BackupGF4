import streamlit as st
import requests
import pandas as pd
from datetime import date, datetime, timedelta
from io import BytesIO

st.set_page_config(page_title="ZIG - Venda Completa", layout="wide")
st.title("🧪 ZIG - Venda Completa Detalhada")

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
        periodos.append(atual)
        atual += timedelta(days=1)

    return periodos


def buscar_get(endpoint, params, timeout=60):
    try:
        resp = requests.get(
            f"{base_url}{endpoint}",
            headers=headers,
            params=params,
            timeout=timeout
        )

        if resp.status_code != 200:
            return [], f"Erro {resp.status_code}: {resp.text}"

        dados = resp.json()

        if isinstance(dados, list):
            return dados, None

        if isinstance(dados, dict):
            return [dados], None

        return [], "Retorno inválido"

    except Exception as e:
        return [], str(e)


def buscar_paginado(endpoint, params_base, max_paginas=100):
    todos = []
    erros = []

    for page in range(1, max_paginas + 1):
        params = params_base.copy()
        params["page"] = page

        dados, erro = buscar_get(endpoint, params)

        if erro:
            erros.append(erro)
            break

        if not dados:
            break

        todos.extend(dados)

    return todos, erros


def centavos(valor):
    return float(valor or 0) / 100


def brl(v):
    return (
        f"R$ {v:,.2f}"
        .replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
    )


col1, col2 = st.columns(2)

with col1:
    dtinicio = st.date_input("Data início", value=date.today() - timedelta(days=1))

with col2:
    dtfim = st.date_input("Data fim", value=date.today() - timedelta(days=1))


if st.button("🔍 Gerar Venda Completa"):

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
        st.error("Erro ao buscar lojas")
        st.text(resp_lojas.text)
        st.stop()

    lojas = resp_lojas.json()
    datas = gerar_periodos_1_dia(dtinicio, dtfim)

    produtos = []
    compradores = []
    faturamento = []
    maquina = []
    notas = []
    erros = []

    total_consultas = len(lojas) * len(datas)
    progresso = st.progress(0)
    contador = 0

    for loja in lojas:
        loja_id = loja.get("id")
        loja_nome = loja.get("name")

        for data_ref in datas:
            data_str = data_ref.strftime("%Y-%m-%d")

            contador += 1
            progresso.progress(contador / total_consultas)

            params = {
                "dtinicio": data_str,
                "dtfim": data_str,
                "loja": loja_id
            }

            # PRODUTOS
            dados, erro = buscar_get("/erp/saida-produtos", params)

            if erro:
                erros.append({
                    "Data": data_str,
                    "Loja": loja_nome,
                    "Endpoint": "saida-produtos",
                    "Erro": erro
                })
            else:
                for item in dados:
                    unit_value = centavos(item.get("unitValue"))
                    count = float(item.get("count", 0) or 0)
                    fractional = item.get("fractionalAmount")
                    desconto = centavos(item.get("discountValue"))

                    qtd = count

                    if fractional not in [None, "", 0]:
                        try:
                            qtd = float(fractional)
                        except Exception:
                            qtd = count

                    valor_produto = (unit_value * qtd) - desconto

                    produtos.append({
                        "Data Consulta": data_str,
                        "Loja": loja_nome,
                        "Loja ID": loja_id,
                        "transactionId": item.get("transactionId"),
                        "transactionDate": item.get("transactionDate"),
                        "eventId": item.get("eventId"),
                        "eventDate": item.get("eventDate"),
                        "invoiceId": item.get("invoiceId"),
                        "employeeName": item.get("employeeName"),
                        "productId": item.get("productId"),
                        "productSku": item.get("productSku"),
                        "Produto": item.get("productName"),
                        "Categoria": item.get("productCategory"),
                        "Tipo Produto": item.get("type"),
                        "unitValue_R$": unit_value,
                        "count": item.get("count"),
                        "fractionalAmount": item.get("fractionalAmount"),
                        "fractionUnit": item.get("fractionUnit"),
                        "discountValue_R$": desconto,
                        "Quantidade Calculada": qtd,
                        "Valor Produto R$": valor_produto,
                        "additions": item.get("additions")
                    })

            # COMPRADORES / GORJETA
            dados, erro = buscar_get("/erp/compradores", params)

            if erro:
                erros.append({
                    "Data": data_str,
                    "Loja": loja_nome,
                    "Endpoint": "compradores",
                    "Erro": erro
                })
            else:
                for item in dados:
                    compradores.append({
                        "Data Consulta": data_str,
                        "Loja": loja_nome,
                        "Loja ID": loja_id,
                        "transactionId": item.get("transactionId"),
                        "userName": item.get("userName"),
                        "userDocument": item.get("userDocument"),
                        "userDocumentType": item.get("userDocumentType"),
                        "userPhone": item.get("userPhone"),
                        "userEmail": item.get("userEmail"),
                        "productsValue_R$": centavos(item.get("productsValue")),
                        "tipValue_R$": centavos(item.get("tipValue")),
                        "Total Comprador c/ Gorjeta R$": (
                            centavos(item.get("productsValue")) +
                            centavos(item.get("tipValue"))
                        )
                    })

            # FATURAMENTO
            dados, erro = buscar_get("/erp/faturamento", params)

            if erro:
                erros.append({
                    "Data": data_str,
                    "Loja": loja_nome,
                    "Endpoint": "faturamento",
                    "Erro": erro
                })
            else:
                for item in dados:
                    faturamento.append({
                        "Data Consulta": data_str,
                        "Loja": loja_nome,
                        "Loja ID": loja_id,
                        "paymentId": item.get("paymentId"),
                        "paymentName": item.get("paymentName"),
                        "Valor Faturamento R$": centavos(item.get("value")),
                        "redeId": item.get("redeId"),
                        "lojaId": item.get("lojaId"),
                        "eventId": item.get("eventId"),
                        "eventDate": item.get("eventDate")
                    })

            # MÁQUINA INTEGRADA
            dados, erro = buscar_get("/erp/faturamento/detalhesMaquinaIntegrada", params)

            if erro:
                erros.append({
                    "Data": data_str,
                    "Loja": loja_nome,
                    "Endpoint": "maquina-integrada",
                    "Erro": erro
                })
            else:
                for item in dados:
                    values = item.get("values", [])

                    if isinstance(values, list):
                        for v in values:
                            maquina.append({
                                "Data Consulta": data_str,
                                "Loja": loja_nome,
                                "Loja ID": loja_id,
                                "paymentId": item.get("paymentId"),
                                "paymentName": item.get("paymentName"),
                                "eventId": item.get("eventId"),
                                "lojaId": item.get("lojaId"),
                                "Bandeira": v.get("cardBrand"),
                                "Valor Máquina R$": centavos(v.get("totalValue"))
                            })

            # NOTAS FISCAIS
            dados, erros_nf = buscar_paginado("/erp/invoice", params)

            if erros_nf:
                for erro_nf in erros_nf:
                    erros.append({
                        "Data": data_str,
                        "Loja": loja_nome,
                        "Endpoint": "invoice",
                        "Erro": erro_nf
                    })

            for item in dados:
                notas.append({
                    "Data Consulta": data_str,
                    "Loja": loja_nome,
                    "Loja ID": loja_id,
                    "invoiceId": item.get("id"),
                    "invoice_eventId": item.get("eventId"),
                    "invoice_eventDate": item.get("eventDate"),
                    "invoice_mode": item.get("mode"),
                    "invoice_isCanceled": item.get("isCanceled"),
                    "invoice_xml": item.get("xml"),
                    "invoice_canceledXml": item.get("canceledXml")
                })

    df_prod = pd.DataFrame(produtos)
    df_comp = pd.DataFrame(compradores)
    df_fat = pd.DataFrame(faturamento)
    df_maq = pd.DataFrame(maquina)
    df_nf = pd.DataFrame(notas)
    df_erros = pd.DataFrame(erros)

    # ==========================
    # RESUMO PRODUTOS POR VENDA
    # ==========================

    if not df_prod.empty:
        resumo_prod = (
            df_prod
            .groupby(
                [
                    "Data Consulta",
                    "Loja",
                    "Loja ID",
                    "transactionId",
                    "eventId",
                    "eventDate",
                    "invoiceId"
                ],
                dropna=False,
                as_index=False
            )
            .agg({
                "Valor Produto R$": "sum",
                "Produto": lambda x: " | ".join(x.dropna().astype(str).unique()),
                "Categoria": lambda x: " | ".join(x.dropna().astype(str).unique()),
                "Tipo Produto": lambda x: " | ".join(x.dropna().astype(str).unique()),
                "employeeName": lambda x: " | ".join(x.dropna().astype(str).unique())
            })
            .rename(columns={
                "Valor Produto R$": "Total Produtos Detalhe R$",
                "Produto": "Produtos",
                "Categoria": "Categorias",
                "Tipo Produto": "Tipos Produto",
                "employeeName": "Funcionários"
            })
        )
    else:
        resumo_prod = pd.DataFrame()

    # ==========================
    # COMPRADORES POR VENDA
    # ==========================

    if not df_comp.empty:
        resumo_comp = (
            df_comp
            .groupby(
                [
                    "Data Consulta",
                    "Loja",
                    "Loja ID",
                    "transactionId"
                ],
                dropna=False,
                as_index=False
            )
            .agg({
                "userName": lambda x: " | ".join(x.dropna().astype(str).unique()),
                "userDocument": lambda x: " | ".join(x.dropna().astype(str).unique()),
                "userPhone": lambda x: " | ".join(x.dropna().astype(str).unique()),
                "userEmail": lambda x: " | ".join(x.dropna().astype(str).unique()),
                "productsValue_R$": "sum",
                "tipValue_R$": "sum",
                "Total Comprador c/ Gorjeta R$": "sum"
            })
        )
    else:
        resumo_comp = pd.DataFrame()

    # ==========================
    # BASE VENDA COMPLETA
    # ==========================

    if not resumo_prod.empty and not resumo_comp.empty:
        venda_completa = resumo_prod.merge(
            resumo_comp,
            on=["Data Consulta", "Loja", "Loja ID", "transactionId"],
            how="outer"
        )
    elif not resumo_prod.empty:
        venda_completa = resumo_prod.copy()
    elif not resumo_comp.empty:
        venda_completa = resumo_comp.copy()
    else:
        venda_completa = pd.DataFrame()

    # ==========================
    # NOTAS FISCAIS POR INVOICE
    # ==========================

    if not venda_completa.empty and not df_nf.empty and "invoiceId" in venda_completa.columns:
        venda_completa = venda_completa.merge(
            df_nf,
            on=["Data Consulta", "Loja", "Loja ID", "invoiceId"],
            how="left"
        )

    # ==========================
    # RESUMOS POR EVENTO
    # Faturamento e Máquina não vêm por transactionId,
    # então juntamos por Data + Loja + eventId.
    # ==========================

    if not df_fat.empty:
        resumo_fat = (
            df_fat
            .groupby(
                ["Data Consulta", "Loja", "Loja ID", "eventId"],
                dropna=False,
                as_index=False
            )
            .agg({
                "paymentName": lambda x: " | ".join(x.dropna().astype(str).unique()),
                "Valor Faturamento R$": "sum"
            })
            .rename(columns={
                "paymentName": "Meios Faturamento",
                "Valor Faturamento R$": "Total Faturamento Evento R$"
            })
        )
    else:
        resumo_fat = pd.DataFrame()

    if not df_maq.empty:
        resumo_maq = (
            df_maq
            .groupby(
                ["Data Consulta", "Loja", "Loja ID", "eventId"],
                dropna=False,
                as_index=False
            )
            .agg({
                "paymentName": lambda x: " | ".join(x.dropna().astype(str).unique()),
                "Bandeira": lambda x: " | ".join(x.dropna().astype(str).unique()),
                "Valor Máquina R$": "sum"
            })
            .rename(columns={
                "paymentName": "Meios Máquina",
                "Bandeira": "Bandeiras Máquina",
                "Valor Máquina R$": "Total Máquina Evento R$"
            })
        )
    else:
        resumo_maq = pd.DataFrame()

    if not venda_completa.empty and "eventId" in venda_completa.columns:
        if not resumo_fat.empty:
            venda_completa = venda_completa.merge(
                resumo_fat,
                on=["Data Consulta", "Loja", "Loja ID", "eventId"],
                how="left"
            )

        if not resumo_maq.empty:
            venda_completa = venda_completa.merge(
                resumo_maq,
                on=["Data Consulta", "Loja", "Loja ID", "eventId"],
                how="left"
            )

    # ==========================
    # COLUNAS DE CONFERÊNCIA
    # ==========================

    for col in [
        "Total Produtos Detalhe R$",
        "productsValue_R$",
        "tipValue_R$",
        "Total Comprador c/ Gorjeta R$",
        "Total Faturamento Evento R$",
        "Total Máquina Evento R$"
    ]:
        if col in venda_completa.columns:
            venda_completa[col] = pd.to_numeric(
                venda_completa[col],
                errors="coerce"
            ).fillna(0).round(2)

    if (
        "Total Comprador c/ Gorjeta R$" in venda_completa.columns
        and "Total Produtos Detalhe R$" in venda_completa.columns
    ):
        venda_completa["Dif. Comprador c/ Gorjeta x Produtos R$"] = (
            venda_completa["Total Comprador c/ Gorjeta R$"]
            - venda_completa["Total Produtos Detalhe R$"]
        ).round(2)

    if (
        "Total Máquina Evento R$" in venda_completa.columns
        and "Total Faturamento Evento R$" in venda_completa.columns
    ):
        venda_completa["Dif. Máquina x Faturamento Evento R$"] = (
            venda_completa["Total Máquina Evento R$"]
            - venda_completa["Total Faturamento Evento R$"]
        ).round(2)

    if (
        "Total Máquina Evento R$" in venda_completa.columns
        and "tipValue_R$" in venda_completa.columns
        and "Total Faturamento Evento R$" in venda_completa.columns
    ):
        venda_completa["Possível Dif. Gorjeta R$"] = (
            venda_completa["Total Máquina Evento R$"]
            - venda_completa["Total Faturamento Evento R$"]
            - venda_completa["tipValue_R$"]
        ).round(2)

    if "invoice_isCanceled" in venda_completa.columns:
        venda_completa["Nota Cancelada?"] = venda_completa["invoice_isCanceled"]

    # ==========================
    # ORDENAR
    # ==========================

    if not venda_completa.empty:
        colunas_prioritarias = [
            "Data Consulta",
            "Loja",
            "Loja ID",
            "transactionId",
            "eventId",
            "eventDate",
            "invoiceId",
            "invoice_isCanceled",
            "userName",
            "userDocument",
            "userPhone",
            "productsValue_R$",
            "tipValue_R$",
            "Total Comprador c/ Gorjeta R$",
            "Total Produtos Detalhe R$",
            "Dif. Comprador c/ Gorjeta x Produtos R$",
            "Produtos",
            "Categorias",
            "Tipos Produto",
            "Funcionários",
            "Meios Faturamento",
            "Total Faturamento Evento R$",
            "Meios Máquina",
            "Bandeiras Máquina",
            "Total Máquina Evento R$",
            "Dif. Máquina x Faturamento Evento R$",
            "Possível Dif. Gorjeta R$",
            "invoice_mode",
            "invoice_canceledXml"
        ]

        colunas_existentes = [
            c for c in colunas_prioritarias
            if c in venda_completa.columns
        ]

        outras_colunas = [
            c for c in venda_completa.columns
            if c not in colunas_existentes
        ]

        venda_completa = venda_completa[colunas_existentes + outras_colunas]

    total_vendas = len(venda_completa)
    total_prod = venda_completa["Total Produtos Detalhe R$"].sum() if "Total Produtos Detalhe R$" in venda_completa.columns else 0
    total_gorjeta = venda_completa["tipValue_R$"].sum() if "tipValue_R$" in venda_completa.columns else 0
    total_com_gorjeta = venda_completa["Total Comprador c/ Gorjeta R$"].sum() if "Total Comprador c/ Gorjeta R$" in venda_completa.columns else 0

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Vendas", total_vendas)
    col2.metric("Produtos", brl(total_prod))
    col3.metric("Gorjeta", brl(total_gorjeta))
    col4.metric("Total c/ Gorjeta", brl(total_com_gorjeta))

    st.subheader("Venda Completa")
    st.dataframe(venda_completa, use_container_width=True, hide_index=True)

    if not df_erros.empty:
        st.subheader("Erros")
        st.dataframe(df_erros, use_container_width=True, hide_index=True)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        venda_completa.to_excel(
            writer,
            index=False,
            sheet_name="Venda Completa"
        )

        if not df_erros.empty:
            df_erros.to_excel(
                writer,
                index=False,
                sheet_name="Erros"
            )

    output.seek(0)

    st.download_button(
        label="📥 Baixar Venda Completa",
        data=output,
        file_name=f"zig_venda_completa_{dtinicio.strftime('%Y%m%d')}_{dtfim.strftime('%Y%m%d')}_{datetime.now().strftime('%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
