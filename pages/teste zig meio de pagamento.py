import streamlit as st
import requests
import pandas as pd
import numpy as np
import re
import json
import unicodedata
from datetime import date, datetime, timedelta
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Teste ZIG Meio Pagamento", layout="wide")
st.title("🧪 Teste ZIG - Meio de Pagamento")


def _strip_accents_keep_case(s: str) -> str:
    return unicodedata.normalize("NFKD", str(s or "")).encode("ASCII", "ignore").decode("ASCII")


def _norm(s: str) -> str:
    s = _strip_accents_keep_case(s)
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


def gerar_periodos_5_dias(data_inicio, data_fim):
    periodos = []
    atual = data_inicio

    while atual <= data_fim:
        fim_bloco = min(atual + timedelta(days=4), data_fim)
        periodos.append((atual, fim_bloco))
        atual = fim_bloco + timedelta(days=1)

    return periodos


# ======================
# Google Sheets
# ======================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)

planilha = gc.open("Tabelas")

df_empresa = pd.DataFrame(planilha.worksheet("Tabela Empresa").get_all_records())
df_empresa.columns = [str(c).strip() for c in df_empresa.columns]

df_meio_pgto_google = pd.DataFrame(planilha.worksheet("Tabela Meio Pagamento").get_all_records())
df_meio_pgto_google.columns = [str(c).strip() for c in df_meio_pgto_google.columns]

for col in ["Meio de Pagamento", "Tipo de Pagamento", "Tipo DRE"]:
    if col not in df_meio_pgto_google.columns:
        df_meio_pgto_google[col] = ""
    df_meio_pgto_google[col] = df_meio_pgto_google[col].astype(str).str.strip()

df_meio_pgto_google["__meio_norm__"] = df_meio_pgto_google["Meio de Pagamento"].map(_norm)
df_meio_pgto_google = df_meio_pgto_google.drop_duplicates(subset=["__meio_norm__"], keep="first")

if "Loja" in df_empresa.columns:
    df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.lower()


# ======================
# ZIG
# ======================
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
    dtfim = st.date_input("Data fim", value=date.today() - timedelta(days=1))


if st.button("🔄 Buscar ZIG Meio de Pagamento"):

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

    registros = []
    lojas_sem_movimento = []

    for loja in lojas:
        loja_id = loja.get("id")
        loja_nome = loja.get("name")
        teve_movimento = False

        for inicio_bloco, fim_bloco in periodos:

            resp = requests.get(
                f"{base_url}/erp/faturamento/detalhesMaquinaIntegrada",
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
                payment_name = str(item.get("paymentName") or "").strip()
                loja_id_retorno = item.get("lojaId")
                event_id = item.get("eventId")
                values = item.get("values", [])

                if not isinstance(values, list):
                    continue

                for v in values:
                    card_brand = str(v.get("cardBrand") or "").strip()
                    total_value = float(v.get("totalValue", 0) or 0) / 100

                    if total_value == 0:
                        continue

                    if card_brand:
                        meio_pagamento = f"{payment_name} {card_brand}".strip()
                    else:
                        meio_pagamento = payment_name

                    registros.append({
                        "Data_raw": item.get("eventDate"),
                        "Loja": loja_nome,
                        "Loja ID": loja_id_retorno,
                        "Event ID": event_id,
                        "Meio de Pagamento": meio_pagamento,
                        "Valor (R$)": total_value
                    })

        if not teve_movimento:
            lojas_sem_movimento.append(loja_nome)

    if lojas_sem_movimento:
        st.info(
            f"ℹ️ {len(lojas_sem_movimento)} loja(s) sem movimentação no período: "
            + ", ".join(lojas_sem_movimento)
        )

    if not registros:
        st.warning("⚠️ Nenhum meio de pagamento encontrado no período.")
        st.stop()

    df = pd.DataFrame(registros)

    df["Data_dt"] = pd.to_datetime(df["Data_raw"], errors="coerce")
    df["Loja"] = df["Loja"].astype(str).str.strip().str.lower()
    df["Meio de Pagamento"] = (
        df["Meio de Pagamento"]
        .astype(str)
        .str.strip()
        .str.replace(r"\s+", " ", regex=True)
    )
    df["Valor (R$)"] = pd.to_numeric(df["Valor (R$)"], errors="coerce").fillna(0)

    resumo = (
        df.groupby(["Data_dt", "Loja", "Meio de Pagamento"], as_index=False)
        .agg({"Valor (R$)": "sum"})
    )

    dias_traducao = {
        "Monday": "segunda-feira",
        "Tuesday": "terça-feira",
        "Wednesday": "quarta-feira",
        "Thursday": "quinta-feira",
        "Friday": "sexta-feira",
        "Saturday": "sábado",
        "Sunday": "domingo"
    }

    resumo["Data"] = resumo["Data_dt"].dt.strftime("%d/%m/%Y")
    resumo["Dia da Semana"] = resumo["Data_dt"].dt.day_name().map(dias_traducao)

    meses = {
        1: "jan", 2: "fev", 3: "mar", 4: "abr",
        5: "mai", 6: "jun", 7: "jul", 8: "ago",
        9: "set", 10: "out", 11: "nov", 12: "dez"
    }

    resumo["Mês"] = resumo["Data_dt"].dt.month.map(meses)
    resumo["Ano"] = resumo["Data_dt"].dt.year
    resumo["Sistema"] = "ZIG"

    resumo = resumo.merge(
        df_empresa[["Loja", "Código Everest", "Grupo", "Código Grupo Everest"]],
        on="Loja",
        how="left"
    )

    resumo["__meio_norm__"] = resumo["Meio de Pagamento"].map(_norm)

    pgto_map = dict(
        zip(
            df_meio_pgto_google["__meio_norm__"],
            df_meio_pgto_google["Tipo de Pagamento"]
        )
    )

    dre_map = dict(
        zip(
            df_meio_pgto_google["__meio_norm__"],
            df_meio_pgto_google["Tipo DRE"]
        )
    )

    resumo["Tipo de Pagamento"] = resumo["__meio_norm__"].map(pgto_map).fillna("")
    resumo["Tipo DRE"] = resumo["__meio_norm__"].map(dre_map).fillna("")
    resumo.drop(columns=["__meio_norm__"], inplace=True, errors="ignore")

    col_order = [
        "Data",
        "Dia da Semana",
        "Meio de Pagamento",
        "Tipo de Pagamento",
        "Tipo DRE",
        "Loja",
        "Código Everest",
        "Grupo",
        "Código Grupo Everest",
        "Sistema",
        "Valor (R$)",
        "Mês",
        "Ano"
    ]

    for c in col_order:
        if c not in resumo.columns:
            resumo[c] = ""

    resumo = resumo[col_order].copy()
    resumo["Valor (R$)"] = pd.to_numeric(resumo["Valor (R$)"], errors="coerce").fillna(0).round(2)

    resumo["_DataOrdenada"] = pd.to_datetime(resumo["Data"], dayfirst=True, errors="coerce")
    resumo = resumo.sort_values(["_DataOrdenada", "Loja", "Meio de Pagamento"]).drop(columns="_DataOrdenada")

    datas_validas = pd.to_datetime(resumo["Data"], dayfirst=True, errors="coerce").dropna()

    if not datas_validas.empty:
        data_inicial = datas_validas.min().strftime("%d/%m/%Y")
        data_final = datas_validas.max().strftime("%d/%m/%Y")
        valor_total = resumo["Valor (R$)"].sum()

        valor_total_formatado = (
            f"R$ {valor_total:,.2f}"
            .replace(",", "X")
            .replace(".", ",")
            .replace("X", ".")
        )

        col1, col2 = st.columns(2)

        with col1:
            st.markdown(
                f"<div style='font-size:1.2rem;'>📅 Período processado<br>{data_inicial} até {data_final}</div>",
                unsafe_allow_html=True
            )

        with col2:
            st.markdown(
                f"<div style='font-size:1.2rem;'>💰 Valor total<br><span style='color:green;'>{valor_total_formatado}</span></div>",
                unsafe_allow_html=True
            )

    meios_norm_tabela = set(df_meio_pgto_google["__meio_norm__"])
    meios_nao_localizados = resumo[
        ~resumo["Meio de Pagamento"].astype(str).str.strip().map(_norm).isin(meios_norm_tabela)
    ]["Meio de Pagamento"].astype(str).unique()

    if len(meios_nao_localizados) > 0:
        meios_nao_localizados_str = "<br>".join(meios_nao_localizados)
        st.markdown(
            f"""
            ⚠️ {len(meios_nao_localizados)} meio(s) de pagamento não localizado(s):<br>
            {meios_nao_localizados_str}
            """,
            unsafe_allow_html=True
        )
    else:
        st.success("✅ Todos os meios de pagamento foram localizados!")

    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        resumo.to_excel(writer, index=False, sheet_name="FaturamentoPorMeio")

    output.seek(0)

    st.download_button(
        label="📥 Baixar Excel ZIG Meio Pagamento",
        data=output,
        file_name=f"zig_meio_pagamento_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
