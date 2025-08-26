# pages/PainelResultados.py
import streamlit as st
st.set_page_config(page_title="Vendas Diarias", layout="wide")  # ✅ Escolha um título só

import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
import plotly.express as px
import io
from st_aggrid import AgGrid, GridOptionsBuilder
from datetime import datetime, date
from datetime import datetime, date, timedelta
from calendar import monthrange
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Image as RLImage, Spacer
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet


#st.set_page_config(page_title="Painel Agrupado", layout="wide")
#st.set_page_config(page_title="Vendas Diarias", layout="wide")
# 🔒 Bloqueia o acesso caso o usuário não esteja logado
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ================================
# 1. Conexão com Google Sheets
# ================================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
gc = gspread.authorize(credentials)
planilha_empresa = gc.open("Vendas diarias")


df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())

# ================================
# 2. Configuração inicial do app
# ================================


# 🎨 Estilizar abas
st.markdown("""
    <style>
    .stApp { background-color: #f9f9f9; }
    div[data-baseweb="tab-list"] { margin-top: 20px; }
    button[data-baseweb="tab"] {
        background-color: #f0f2f6;
        border-radius: 10px;
        padding: 10px 20px;
        margin-right: 10px;
        transition: all 0.3s ease;
        font-size: 16px;
        font-weight: 600;
    }
    button[data-baseweb="tab"]:hover { background-color: #dce0ea; color: black; }
    button[data-baseweb="tab"][aria-selected="true"] { background-color: #0366d6; color: white; }
    </style>
""", unsafe_allow_html=True)

# Cabeçalho bonito
st.markdown("""
    <div style='display: flex; align-items: center; gap: 10px; margin-bottom: 20px;'>
        <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
        <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Rateio</h1>
    </div>
""", unsafe_allow_html=True)

  # 🎨 Remove fundo e borda do multiselect, deixa só texto preto
st.markdown("""
    <style>
    /* Itens selecionados */
    div[data-testid="stMultiSelect"] [data-baseweb="tag"] {
        background-color: transparent !important; /* sem fundo */
        border: none !important;                   /* sem borda */
        color: black !important;                   /* texto preto */
    }
    /* Texto e ícone dentro do item selecionado */
    div[data-testid="stMultiSelect"] [data-baseweb="tag"] * {
        color: black !important;
        fill: black !important;                    /* ícone X */
    }
    /* Área interna do multiselect sem fundo */
    div[data-testid="stMultiSelect"] > div {
        background-color: transparent !important;
    }
    </style>
""", unsafe_allow_html=True)
# ================================
# 3. Separação em ABAS
# ================================


# ================================
# Aba 3: Relatórios Vendas
# ================================


import pandas as pd
from datetime import datetime
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Carrega dados
df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())
df_vendas = pd.DataFrame(planilha_empresa.worksheet("Fat Sistema Externo").get_all_records())

# Normalização
df_empresa["Loja"] = df_empresa["Loja"].str.strip().str.upper()
df_empresa["Grupo"] = df_empresa["Grupo"].str.strip()
df_vendas.columns = df_vendas.columns.str.strip()
df_vendas["Data"] = pd.to_datetime(df_vendas["Data"], dayfirst=True, errors="coerce")
df_vendas["Loja"] = df_vendas["Loja"].astype(str).str.strip().str.upper()
df_vendas["Grupo"] = df_vendas["Grupo"].astype(str).str.strip()

# Merge com Tipo
df_vendas = df_vendas.merge(df_empresa[["Loja", "Tipo"]], on="Loja", how="left")

# Ajusta Fat.Total
df_vendas["Fat.Total"] = (
    df_vendas["Fat.Total"]
    .astype(str)
    .str.replace("R$", "", regex=False)
    .str.replace("(", "-", regex=False)
    .str.replace(")", "", regex=False)
    .str.replace(" ", "", regex=False)
    .str.replace(".", "", regex=False)
    .str.replace(",", ".", regex=False)
)
df_vendas["Fat.Total"] = pd.to_numeric(df_vendas["Fat.Total"], errors="coerce")


# ==== Filtros lado a lado ====

col1, col2, col3 = st.columns([1, 1, 2])

with col1:
    tipos_disponiveis = sorted(df_vendas["Tipo"].dropna().unique())
    tipos_disponiveis.insert(0, "Todos")
    tipo_selecionado = st.selectbox("🏪 Tipo:", options=tipos_disponiveis, index=0)

with col2:
    grupos_disponiveis = sorted(df_vendas["Grupo"].dropna().unique())
    grupos_disponiveis.insert(0, "Todos")
    grupo_selecionado = st.selectbox("👥 Grupo:", options=grupos_disponiveis, index=0)

with col3:
    df_vendas["Mes/Ano"] = df_vendas["Data"].dt.strftime("%m/%Y")
    meses_disponiveis = sorted(
        df_vendas["Mes/Ano"].unique(),
        key=lambda x: datetime.strptime("01/" + x, "%d/%m/%Y")
    )
    meses_selecionados = st.multiselect(
        "🗓️ Selecione os meses:",
        options=meses_disponiveis,
        default=[datetime.today().strftime("%m/%Y")]
    )

# ==== Aplica filtros ====
df_filtrado = df_vendas[df_vendas["Mes/Ano"].isin(meses_selecionados)]
df_filtrado["Período"] = df_filtrado["Data"].dt.strftime("%m/%Y")

if tipo_selecionado != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Tipo"] == tipo_selecionado]

if grupo_selecionado != "Todos":
    df_filtrado = df_filtrado[df_filtrado["Grupo"] == grupo_selecionado]



# ==== Agrupamento dinâmico ====
if grupo_selecionado == "Todos":
    chaves = ["Tipo", "Grupo"]
else:
    chaves = ["Grupo", "Loja"]

df_agrupado = df_filtrado.groupby(chaves + ["Período"], as_index=False)["Fat.Total"].sum()
df_final = df_agrupado.groupby(chaves, as_index=False)["Fat.Total"].sum()
df_final.rename(columns={"Fat.Total": "Total"}, inplace=True)
df_final["Rateio"] = 0.0

# ==== Cálculo de % e Subtotais ====
if grupo_selecionado == "Todos":
    # Caso geral: Tipo + Grupo
    total_geral = df_final["Total"].sum()
    df_final["% Total"] = df_final["Total"] / total_geral

    # Ordena por subtotal do Tipo
    subtotais_tipo = df_final.groupby("Tipo")["Total"].sum().reset_index()
    subtotais_tipo = subtotais_tipo.sort_values(by="Total", ascending=False)
    ordem_tipos = subtotais_tipo["Tipo"].tolist()
    df_final["ord_tipo"] = df_final["Tipo"].apply(lambda x: ordem_tipos.index(x) if x in ordem_tipos else 999)
    df_final = df_final.sort_values(by=["ord_tipo", "Total"], ascending=[True, False]).drop(columns="ord_tipo")

    # Monta subtotais por Tipo
    linhas_com_subtotal = []
    for tipo in ordem_tipos:
        bloco_tipo = df_final[df_final["Tipo"] == tipo].copy()
        linhas_com_subtotal.append(bloco_tipo)
        subtotal = bloco_tipo.drop(columns=["Tipo", "Grupo"]).sum(numeric_only=True)
        subtotal["Tipo"] = tipo
        subtotal["Grupo"] = f"Subtotal {tipo}"
        linhas_com_subtotal.append(pd.DataFrame([subtotal]))
    df_final = pd.concat(linhas_com_subtotal, ignore_index=True)

else:
    # Caso específico: Grupo + Loja
    total_geral = df_final["Total"].sum()
    df_final["% Total"] = df_final["Total"] / total_geral

    # Ordena por Total dentro do Grupo
    df_final = df_final.sort_values(by=["Grupo", "Total"], ascending=[True, False])

    # Monta subtotais por Grupo
    linhas_com_subtotal = []
    for grupo in df_final["Grupo"].unique():
        bloco_grupo = df_final[df_final["Grupo"] == grupo].copy()
        linhas_com_subtotal.append(bloco_grupo)
        subtotal = bloco_grupo.drop(columns=["Grupo", "Loja"]).sum(numeric_only=True)
        subtotal["Grupo"] = grupo
        subtotal["Loja"] = f"Subtotal {grupo}"
        linhas_com_subtotal.append(pd.DataFrame([subtotal]))
    df_final = pd.concat(linhas_com_subtotal, ignore_index=True)

# ==== Linha TOTAL no topo ====
cols_drop = [c for c in ["Tipo","Grupo","Loja"] if c in df_final.columns]
apenas = df_final.copy()
for col in cols_drop:
    apenas = apenas[~apenas[col].astype(str).str.startswith("Subtotal", na=False)]
linha_total = apenas.drop(columns=cols_drop, errors="ignore").sum(numeric_only=True)
for col in cols_drop:
    linha_total[col] = ""  # limpa
linha_total[cols_drop[0] if cols_drop else "Grupo"] = "TOTAL"  # garante nome TOTAL
df_final = pd.concat([pd.DataFrame([linha_total]), df_final], ignore_index=True)

# ==== RATEIO ====
df_final["% Total"] = 0.0
df_final["Rateio"] = 0.0

if grupo_selecionado == "Todos":
    # === Caso geral: Rateio por Tipo ===
    def moeda_para_float(valor_str: str) -> float:
        try:
            return float(valor_str.replace(".", "").replace(",", "."))
        except:
            return 0.0

    tipos_unicos = [
        t for t in df_final["Tipo"].dropna().unique()
        if str(t).strip() not in ["", "TOTAL"] and not str(t).startswith("Subtotal")
    ]
    valores_rateio_por_tipo = {}

    COLS_POR_LINHA = 3
    for i in range(0, len(tipos_unicos), COLS_POR_LINHA):
        linha = tipos_unicos[i:i+COLS_POR_LINHA]
        cols = st.columns(len(linha))
        for c, tipo in zip(cols, linha):
            with c:
                valor_str = st.text_input(
                    f"💰 Rateio — {tipo}",
                    value="0,00",
                    key=f"rateio_{tipo}"
                )
                valores_rateio_por_tipo[tipo] = moeda_para_float(valor_str)

    for tipo in df_final["Tipo"].unique():
        mask_tipo = (
            (df_final["Tipo"] == tipo) &
            (~df_final["Grupo"].str.startswith("Subtotal")) &
            (df_final["Grupo"] != "TOTAL")
        )
        subtotal_tipo = df_final.loc[df_final["Grupo"] == f"Subtotal {tipo}", "Total"].sum()

        if subtotal_tipo > 0:
            df_final.loc[mask_tipo, "% Total"] = (df_final.loc[mask_tipo, "Total"] / subtotal_tipo) * 100

        df_final.loc[df_final["Grupo"] == f"Subtotal {tipo}", "% Total"] = 100

        valor_rateio = valores_rateio_por_tipo.get(tipo, 0.0)
        df_final.loc[mask_tipo, "Rateio"] = df_final.loc[mask_tipo, "% Total"] / 100 * valor_rateio

        rateio_tipo = df_final.loc[mask_tipo, "Rateio"].sum()
        df_final.loc[df_final["Grupo"] == f"Subtotal {tipo}", "Rateio"] = rateio_tipo

else:
    # === Caso específico: Rateio por Grupo (distribuído nas Lojas) ===
    total_rateio = st.number_input(
        f"💰 Rateio — {grupo_selecionado}",
        min_value=0.0, step=100.0, format="%.2f",
        key=f"rateio_{grupo_selecionado}"
    )

    mask_lojas = (
        (df_final["Grupo"] == grupo_selecionado) &
        (~df_final["Loja"].astype(str).str.startswith("Subtotal")) &
        (df_final["Loja"] != "TOTAL")
    )

    subtotal_grupo = df_final.loc[df_final["Loja"] == f"Subtotal {grupo_selecionado}", "Total"].sum()

    if subtotal_grupo > 0:
        df_final.loc[mask_lojas, "% Total"] = (df_final.loc[mask_lojas, "Total"] / subtotal_grupo) * 100
        df_final.loc[df_final["Loja"] == f"Subtotal {grupo_selecionado}", "% Total"] = 100

        df_final.loc[mask_lojas, "Rateio"] = df_final.loc[mask_lojas, "% Total"] / 100 * total_rateio

        df_final.loc[df_final["Loja"] == f"Subtotal {grupo_selecionado}", "Rateio"] = df_final.loc[mask_lojas, "Rateio"].sum()

# ==== Reordenar colunas ====
colunas_existentes = [c for c in ["Tipo","Grupo","Loja","Total","% Total","Rateio"] if c in df_final.columns]
df_final = df_final[colunas_existentes]

# ==== Formatação para exibição ====
df_view = df_final.copy()

def formatar(valor):
    try:
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return valor

for col in ["Total", "Rateio"]:
    if col in df_view.columns:
        df_view[col] = df_view[col].apply(lambda x: formatar(x) if pd.notnull(x) and x != "" else x)

if "% Total" in df_view.columns:
    df_view["% Total"] = pd.to_numeric(df_view["% Total"], errors="coerce").apply(
        lambda x: f"{x:.2f}%" if pd.notnull(x) else ""
    )

def aplicar_estilo(df):
    def estilo_linha(row):
        if "Grupo" in df.columns and row["Grupo"] == "TOTAL":
            return ["background-color: #f4b084; font-weight: bold"] * len(row)
        elif "Loja" in df.columns and row["Loja"].startswith("Subtotal"):
            return ["background-color: #d9d9d9; font-weight: bold"] * len(row)
        elif "Grupo" in df.columns and isinstance(row["Grupo"], str) and row["Grupo"].startswith("Subtotal"):
            return ["background-color: #d9d9d9; font-weight: bold"] * len(row)
        else:
            return ["" for _ in row]
    return df.style.apply(estilo_linha, axis=1)

st.dataframe(aplicar_estilo(df_view), use_container_width=True, height=700)

# ==== Exporta Excel ====


if "% Total" in df_final.columns:
    df_final["% Total"] = pd.to_numeric(df_final["% Total"], errors="coerce") / 100

df_exportar = df_final.copy()
output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    df_exportar.to_excel(writer, index=False, sheet_name="Relatório")
output.seek(0)

wb = load_workbook(output)
ws = wb["Relatório"]

header_font = Font(bold=True, color="FFFFFF")
header_fill = PatternFill("solid", fgColor="305496")
center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

for cell in ws[1]:
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = center_alignment
    cell.border = border

for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
    grupo_valor = row[1].value
    estilo_fundo = None
    if isinstance(grupo_valor, str):
        if grupo_valor.strip().upper() == "TOTAL":
            estilo_fundo = PatternFill("solid", fgColor="F4B084")
        elif "SUBTOTAL" in grupo_valor.strip().upper():
            estilo_fundo = PatternFill("solid", fgColor="D9D9D9")
    for cell in row:
        cell.border = border
        cell.alignment = center_alignment
        if estilo_fundo:
            cell.fill = estilo_fundo
        col_name = ws.cell(row=1, column=cell.column).value
        if isinstance(cell.value, (int, float)):
            if col_name == "% Total":
                if cell.value not in (None, ""):
                    cell.number_format = '0.00%'  # Formato percentual com 2 casas
                else:
                    cell.number_format = '0.00%'
            else:
                cell.number_format = '"R$" #,##0.00'

for i, col_cells in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row), start=1):
    max_length = 0
    for cell in col_cells:
        if cell.value:
            max_length = max(max_length, len(str(cell.value)))
    ws.column_dimensions[get_column_letter(i)].width = max_length + 2

for col_nome in ["Tipo", "Grupo"]:
    if col_nome in df_exportar.columns:
        col_idx = df_exportar.columns.get_loc(col_nome) + 1
        for cell in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
            for c in cell:
                c.alignment = Alignment(horizontal="left")

output_final = BytesIO()
wb.save(output_final)
output_final.seek(0)

st.download_button(
    label="📥 Baixar Excel",
    data=output_final,
    file_name="Resumo_Grupos_Mensal.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet
from datetime import datetime
import pytz
import io

def gerar_pdf(df, mes_rateio, usuario):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=A4,
        topMargin=30,
        bottomMargin=30,
        leftMargin=20,
        rightMargin=20
    )

    elementos = []
    estilos = getSampleStyleSheet()
    estilo_normal = estilos["Normal"]
    estilo_titulo = estilos["Heading1"]

    # Logo
    try:
        logo_url = "https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo_grupofit.png"
        img = Image(logo_url, width=100, height=40)
        elementos.append(img)
    except:
        pass

    # Título
    elementos.append(Paragraph(f"<b>Rateio - {mes_rateio}</b>", estilo_titulo))

    # Data no fuso correto
    fuso_brasilia = pytz.timezone("America/Sao_Paulo")
    data_geracao = datetime.now(fuso_brasilia).strftime("%d/%m/%Y %H:%M")

    elementos.append(Paragraph(f"<b>Usuário:</b> {usuario}", estilo_normal))
    elementos.append(Paragraph(f"<b>Data de Geração:</b> {data_geracao}", estilo_normal))
    elementos.append(Spacer(1, 12))

    # Converte DataFrame para lista
    dados_tabela = [df.columns.tolist()] + df.values.tolist()

    # Tabela
    tabela = Table(dados_tabela, repeatRows=1)

    # Estilo inicial
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003366")),  # Cabeçalho azul escuro
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
    ]))

    # Aplica cor fixa nas linhas
    # Aplica cor fixa nas linhas
    for i in range(1, len(dados_tabela)):
        linha_texto = str(dados_tabela[i][1]).strip().lower()

        if "subtotal" in linha_texto or "total" == linha_texto:
            # Cinza mais escuro para subtotal e total
            tabela.setStyle(TableStyle([("BACKGROUND", (0, i), (-1, i), colors.HexColor("#BFBFBF"))]))
            tabela.setStyle(TableStyle([("FONTNAME", (0, i), (-1, i), "Helvetica-Bold")]))
        else:
            # Cinza claro para linhas normais
            tabela.setStyle(TableStyle([("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F2F2F2"))]))

    elementos.append(tabela)

    doc.build(elementos)
    pdf_value = buffer.getvalue()
    buffer.close()
    return pdf_value


# ====== Chamada no seu Streamlit ======
# ====== Chamada no seu Streamlit ======
usuario_logado = st.session_state.get("usuario_logado", "Usuário Desconhecido")

# monta o título a partir do multiselect existente: `meses_selecionados` (ex.: "08/2025")
sele = meses_selecionados if meses_selecionados else [datetime.today().strftime("%m/%Y")]
if len(sele) == 1:
    mes_rateio = sele[0]
elif len(sele) == 2:
    mes_rateio = f"{sele[0]} e {sele[1]}"
else:
    mes_rateio = f"{', '.join(sele[:-1])} e {sele[-1]}"

pdf_bytes = gerar_pdf(df_view, mes_rateio=mes_rateio, usuario=usuario_logado)

st.download_button(
    label="📄 Baixar PDF",
    data=pdf_bytes,
    file_name=f"Rateio_{datetime.now().strftime('%Y%m%d')}.pdf",
    mime="application/pdf"
)
