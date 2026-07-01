import streamlit as st
import requests
import pandas as pd
from datetime import date, datetime, timedelta
from io import BytesIO

st.set_page_config(page_title="ZIG - Auditoria Todas as Tabelas", layout="wide")
st.title("🧪 ZIG - Auditoria de Todas as Tabelas por Período")

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

        return [], "Retorno não é lista nem dicionário"

    except Exception as e:
        return [], str(e)


def buscar_paginado(endpoint, params_base, campo_page="page", max_paginas=100):
    todos = []
    erros = []

    for page in range(1, max_paginas + 1):
        params = params_base.copy()
        params[campo_page] = page

        dados, erro = buscar_get(endpoint, params)

        if erro:
            erros.append(f"Página {page}: {erro}")
            break

        if not dados:
            break

        todos.extend(dados)

    return todos, erros


def normalizar_lista(registros):
    if not registros:
        return pd.DataFrame()

    return pd.json_normalize(registros)


def adicionar_contexto(lista, loja_id, loja_nome, data_consulta, endpoint):
    registros = []

    for item in lista:
        if isinstance(item, dict):
            novo = item.copy()
        else:
            novo = {"valor_retorno": item}

        novo["_data_consulta"] = data_consulta
        novo["_loja_id_consulta"] = loja_id
        novo["_loja_nome_consulta"] = loja_nome
        novo["_endpoint"] = endpoint

        registros.append(novo)

    return registros


def abrir_maquina_por_bandeira(lista, loja_id, loja_nome, data_consulta, endpoint):
    registros = []

    for item in lista:
        values = item.get("values", [])

        if isinstance(values, list):
            for v in values:
                linha = item.copy()
                linha.pop("values", None)

                for chave, valor in v.items():
                    linha[f"values.{chave}"] = valor

                linha["_data_consulta"] = data_consulta
                linha["_loja_id_consulta"] = loja_id
                linha["_loja_nome_consulta"] = loja_nome
                linha["_endpoint"] = endpoint

                registros.append(linha)

    return registros


col1, col2 = st.columns(2)

with col1:
    dtinicio = st.date_input(
        "Data início",
        value=date.today() - timedelta(days=7)
    )

with col2:
    dtfim = st.date_input(
        "Data fim",
        value=date.today() - timedelta(days=1)
    )


if st.button("🔍 Buscar todas as tabelas ZIG"):

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
        st.text(resp_lojas.text)
        st.stop()

    lojas = resp_lojas.json()

    if not isinstance(lojas, list) or not lojas:
        st.warning("Nenhuma loja encontrada.")
        st.stop()

    periodos = gerar_periodos_1_dia(dtinicio, dtfim)

    registros_lojas = []
    registros_saida_produtos = []
    registros_compradores = []
    registros_faturamento = []
    registros_maquina = []
    registros_maquina_aberta = []
    registros_invoice = []
    registros_checkins = []
    registros_recharges = []
    erros_gerais = []

    total_consultas = len(lojas) * len(periodos)
    progresso = st.progress(0)
    contador = 0

    for loja in lojas:
        loja_id = loja.get("id")
        loja_nome = loja.get("name")

        registros_lojas.append({
            "id": loja_id,
            "name": loja_nome
        })

        for inicio, fim in periodos:
            data_str = inicio.strftime("%Y-%m-%d")

            contador += 1
            progresso.progress(contador / total_consultas)

            params_periodo = {
                "dtinicio": data_str,
                "dtfim": data_str,
                "loja": loja_id
            }

            # ==========================
            # 1. SAÍDA DE PRODUTOS
            # ==========================
            endpoint = "/erp/saida-produtos"
            dados, erro = buscar_get(endpoint, params_periodo)

            if erro:
                erros_gerais.append({
                    "Data": data_str,
                    "Loja": loja_nome,
                    "Loja ID": loja_id,
                    "Endpoint": endpoint,
                    "Erro": erro
                })
            else:
                registros_saida_produtos.extend(
                    adicionar_contexto(dados, loja_id, loja_nome, data_str, endpoint)
                )

            # ==========================
            # 2. COMPRADORES
            # ==========================
            endpoint = "/erp/compradores"
            dados, erro = buscar_get(endpoint, params_periodo)

            if erro:
                erros_gerais.append({
                    "Data": data_str,
                    "Loja": loja_nome,
                    "Loja ID": loja_id,
                    "Endpoint": endpoint,
                    "Erro": erro
                })
            else:
                registros_compradores.extend(
                    adicionar_contexto(dados, loja_id, loja_nome, data_str, endpoint)
                )

            # ==========================
            # 3. FATURAMENTO
            # ==========================
            endpoint = "/erp/faturamento"
            dados, erro = buscar_get(endpoint, params_periodo)

            if erro:
                erros_gerais.append({
                    "Data": data_str,
                    "Loja": loja_nome,
                    "Loja ID": loja_id,
                    "Endpoint": endpoint,
                    "Erro": erro
                })
            else:
                registros_faturamento.extend(
                    adicionar_contexto(dados, loja_id, loja_nome, data_str, endpoint)
                )

            # ==========================
            # 4. MÁQUINA INTEGRADA
            # ==========================
            endpoint = "/erp/faturamento/detalhesMaquinaIntegrada"
            dados, erro = buscar_get(endpoint, params_periodo)

            if erro:
                erros_gerais.append({
                    "Data": data_str,
                    "Loja": loja_nome,
                    "Loja ID": loja_id,
                    "Endpoint": endpoint,
                    "Erro": erro
                })
            else:
                registros_maquina.extend(
                    adicionar_contexto(dados, loja_id, loja_nome, data_str, endpoint)
                )

                registros_maquina_aberta.extend(
                    abrir_maquina_por_bandeira(
                        dados,
                        loja_id,
                        loja_nome,
                        data_str,
                        endpoint
                    )
                )

            # ==========================
            # 5. NOTAS FISCAIS
            # ==========================
            endpoint = "/erp/invoice"
            dados, erros = buscar_paginado(endpoint, params_periodo)

            if erros:
                for erro in erros:
                    erros_gerais.append({
                        "Data": data_str,
                        "Loja": loja_nome,
                        "Loja ID": loja_id,
                        "Endpoint": endpoint,
                        "Erro": erro
                    })

            registros_invoice.extend(
                adicionar_contexto(dados, loja_id, loja_nome, data_str, endpoint)
            )

            # ==========================
            # 6. CHECK-INS
            # ==========================
            endpoint = "/erp/checkins"

            params_checkins = {
                "desde": data_str,
                "dtfim": data_str,
                "loja": loja_id
            }

            dados, erros = buscar_paginado(endpoint, params_checkins)

            if erros:
                for erro in erros:
                    erros_gerais.append({
                        "Data": data_str,
                        "Loja": loja_nome,
                        "Loja ID": loja_id,
                        "Endpoint": endpoint,
                        "Erro": erro
                    })

            registros_checkins.extend(
                adicionar_contexto(dados, loja_id, loja_nome, data_str, endpoint)
            )

            # ==========================
            # 7. RECARGAS
            # ==========================
            endpoint = "/erp/recharges"
            dados, erro = buscar_get(endpoint, params_periodo)

            if erro:
                erros_gerais.append({
                    "Data": data_str,
                    "Loja": loja_nome,
                    "Loja ID": loja_id,
                    "Endpoint": endpoint,
                    "Erro": erro
                })
            else:
                registros_recharges.extend(
                    adicionar_contexto(dados, loja_id, loja_nome, data_str, endpoint)
                )

    df_lojas = normalizar_lista(registros_lojas)
    df_saida_produtos = normalizar_lista(registros_saida_produtos)
    df_compradores = normalizar_lista(registros_compradores)
    df_faturamento = normalizar_lista(registros_faturamento)
    df_maquina = normalizar_lista(registros_maquina)
    df_maquina_aberta = normalizar_lista(registros_maquina_aberta)
    df_invoice = normalizar_lista(registros_invoice)
    df_checkins = normalizar_lista(registros_checkins)
    df_recharges = normalizar_lista(registros_recharges)
    df_erros = normalizar_lista(erros_gerais)

    st.success("Consulta finalizada.")

    col1, col2, col3, col4 = st.columns(4)

    col1.metric("Lojas", len(df_lojas))
    col2.metric("Saída Produtos", len(df_saida_produtos))
    col3.metric("Compradores", len(df_compradores))
    col4.metric("Máquina Aberta", len(df_maquina_aberta))

    st.subheader("Máquina Integrada Aberta por Bandeira")
    st.dataframe(df_maquina_aberta, use_container_width=True, hide_index=True)

    st.subheader("Compradores")
    st.dataframe(df_compradores, use_container_width=True, hide_index=True)

    st.subheader("Saída de Produtos")
    st.dataframe(df_saida_produtos, use_container_width=True, hide_index=True)

    st.subheader("Faturamento")
    st.dataframe(df_faturamento, use_container_width=True, hide_index=True)

    if not df_erros.empty:
        st.subheader("Erros")
        st.dataframe(df_erros, use_container_width=True, hide_index=True)

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_lojas.to_excel(writer, index=False, sheet_name="Lojas")
        df_saida_produtos.to_excel(writer, index=False, sheet_name="Saida Produtos")
        df_compradores.to_excel(writer, index=False, sheet_name="Compradores")
        df_faturamento.to_excel(writer, index=False, sheet_name="Faturamento")
        df_maquina.to_excel(writer, index=False, sheet_name="Maquina Original")
        df_maquina_aberta.to_excel(writer, index=False, sheet_name="Maquina Aberta")
        df_invoice.to_excel(writer, index=False, sheet_name="Notas Fiscais")
        df_checkins.to_excel(writer, index=False, sheet_name="Checkins")
        df_recharges.to_excel(writer, index=False, sheet_name="Recargas")

        if not df_erros.empty:
            df_erros.to_excel(writer, index=False, sheet_name="Erros")

    output.seek(0)

    st.download_button(
        label="📥 Baixar Excel Auditoria ZIG",
        data=output,
        file_name=f"zig_auditoria_todas_tabelas_{dtinicio.strftime('%Y%m%d')}_{dtfim.strftime('%Y%m%d')}_{datetime.now().strftime('%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
