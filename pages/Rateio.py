# pages/PainelResultados.py
import streamlit as st
st.set_page_config(page_title="Vendas Diarias", layout="wide")

import pandas as pd
import numpy as np
from io import BytesIO
from datetime import datetime, date, timedelta
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json
from calendar import monthrange

# ReportLab / Excel
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.platypus import Image as RLImg
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
import pytz
import io

# 🔒 Bloqueia o acesso caso o usuário não esteja logado
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ======================
# CSS
# ======================
st.markdown("""
    <style>
        [data-testid="stToolbar"] { visibility: hidden; height: 0%; position: fixed; }
        .stSpinner { visibility: visible !important; }

        /* Tabs bonitas */
        .stApp { background-color: #f9f9f9; }
        div[data-baseweb="tab-list"] { margin-top: 20px; }
        button[data-baseweb="tab"] {
            background-color: #f0f2f6; border-radius: 10px;
            padding: 10px 20px; margin-right: 10px;
            transition: all 0.3s ease; font-size: 16px; font-weight: 600;
        }
        button[data-baseweb="tab"]:hover { background-color: #dce0ea; color: black; }
        button[data-baseweb="tab"][aria-selected="true"] { background-color: #0366d6; color: white; }

        /* Multiselect clean */
        div[data-testid="stMultiSelect"] [data-baseweb="tag"] { background-color: transparent !important; border: none !important; color: black !important; }
        div[data-testid="stMultiSelect"] [data-baseweb="tag"] * { color: black !important; fill: black !important; }
        div[data-testid="stMultiSelect"] > div { background-color: transparent !important; }
    </style>
""", unsafe_allow_html=True)

# ======================
# Helpers
# ======================
def formatar_moeda_br(valor):
    try:
        return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except:
        return valor

def aplicar_estilo_df(df):
    def estilo_linha(row):
        if "Grupo" in df.columns and isinstance(row.get("Grupo", ""), str) and row["Grupo"] == "TOTAL":
            return ["background-color: #f4b084; font-weight: bold"] * len(row)
        if "Loja" in df.columns and isinstance(row.get("Loja", ""), str) and row["Loja"].startswith("Subtotal"):
            return ["background-color: #d9d9d9; font-weight: bold"] * len(row)
        if "Grupo" in df.columns and isinstance(row.get("Grupo", ""), str) and row["Grupo"].startswith("Subtotal"):
            return ["background-color: #d9d9d9; font-weight: bold"] * len(row)
        return ["" for _ in row]
    return df.style.apply(estilo_linha, axis=1)

def exportar_excel(df_exportar):
    output = BytesIO()
    df_exportar = df_exportar.copy()
    # Ajuste de % antes de exportar
    if "% Total" in df_exportar.columns and df_exportar["% Total"].dtype != float:
        # se vier em string tipo "12,34%" não mexe; se vier numérico (0-100) normalizo
        try:
            df_exportar["% Total"] = pd.to_numeric(df_exportar["% Total"], errors="coerce") / 100.0
        except Exception:
            pass

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_exportar.to_excel(writer, index=False, sheet_name="Relatório")
    output.seek(0)

    wb = load_workbook(output)
    ws = wb["Relatório"]

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="305496")
    center_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))

    # Header
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
        cell.border = border

    # Linhas
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, max_col=ws.max_column):
        # pinta subtotais/total
        grupo_valor = None
        try:
            grupo_valor = row[1].value  # coluna 2 geralmente é "Grupo"
        except:
            grupo_valor = None

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
                    cell.number_format = '0.00%'
                else:
                    cell.number_format = '"R$" #,##0.00'

    # Largura das colunas
    for i, col_cells in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row), start=1):
        max_length = 0
        for cell in col_cells:
            if cell.value:
                max_length = max(max_length, len(str(cell.value)))
        ws.column_dimensions[get_column_letter(i)].width = max_length + 2

    # Alinhar texto à esquerda em colunas de texto (se existirem)
    for col_nome in ["Tipo", "Grupo", "Loja"]:
        try:
            if col_nome in df_exportar.columns:
                col_idx = df_exportar.columns.get_loc(col_nome) + 1
                for cell in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
                    for c in cell:
                        c.alignment = Alignment(horizontal="left")
        except Exception:
            pass

    output_final = BytesIO()
    wb.save(output_final)
    output_final.seek(0)
    return output_final

def gerar_pdf(df, mes_rateio, usuario):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, pagesize=A4, topMargin=30, bottomMargin=30, leftMargin=20, rightMargin=20
    )
    elementos = []
    estilos = getSampleStyleSheet()
    estilo_normal = estilos["Normal"]
    estilo_titulo = estilos["Heading1"]

    # Logo (opcional)
    try:
        logo_url = "https://raw.githubusercontent.com/MMRConsultoria/mmr-site/main/logo_grupofit.png"
        img = RLImg(logo_url, width=100, height=40)
        elementos.append(img)
    except:
        pass

    # Título + metadados
    elementos.append(Paragraph(f"<b>Rateio - {mes_rateio}</b>", estilo_titulo))
    fuso_brasilia = pytz.timezone("America/Sao_Paulo")
    data_geracao = datetime.now(fuso_brasilia).strftime("%d/%m/%Y %H:%M")
    elementos.append(Paragraph(f"<b>Usuário:</b> {usuario or 'Usuário Desconhecido'}", estilo_normal))
    elementos.append(Paragraph(f"<b>Data de Geração:</b> {data_geracao}", estilo_normal))
    elementos.append(Spacer(1, 12))

    # Tabela
    dados_tabela = [df.columns.tolist()] + df.values.tolist()
    tabela = Table(dados_tabela, repeatRows=1)
    tabela.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#003366")),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"),
        ("ALIGN", (0, 0), (0, -1), "LEFT"),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
    ]))
    # Cores por linha
    for i in range(1, len(dados_tabela)):
        txt = ""
        try:
            txt = str(dados_tabela[i][1]).strip().lower()  # coluna 1 (Grupo/Loja)
        except:
            pass
        if "subtotal" in txt or txt == "total":
            tabela.setStyle(TableStyle([("BACKGROUND", (0, i), (-1, i), colors.HexColor("#BFBFBF"))]))
            tabela.setStyle(TableStyle([("FONTNAME", (0, i), (-1, i), "Helvetica-Bold")]))
        else:
            tabela.setStyle(TableStyle([("BACKGROUND", (0, i), (-1, i), colors.HexColor("#F2F2F2"))]))

    elementos.append(tabela)
    doc.build(elementos)
    pdf_value = buffer.getvalue()
    buffer.close()
    return pdf_value

def render_aba(titulo_aba: str, tab_key: str, df_empresa: pd.DataFrame, df_vendas: pd.DataFrame, metric_col: str = "Fat.Total"):
    # Cabeçalho
    st.markdown(f"""
        <div style='display:flex; align-items:center; gap:10px; margin-bottom: 10px;'>
            <img src='https://img.icons8.com/color/48/graph.png' width='32'/>
            <h2 style='display:inline; margin:0;'>{titulo_aba}</h2>
        </div>
    """, unsafe_allow_html=True)

    # ==== Filtros ====
    col1, col2, col3 = st.columns([1, 1, 2])

    with col1:
        tipos_disponiveis = sorted(df_vendas["Tipo"].dropna().unique())
        tipos_disponiveis.insert(0, "Todos")
        tipo_sel = st.selectbox("🏪 Tipo:", options=tipos_disponiveis, index=0, key=f"tipo_{tab_key}")

    with col2:
        grupos_disponiveis = sorted(df_vendas["Grupo"].dropna().unique())
        grupos_disponiveis.insert(0, "Todos")
        grupo_sel = st.selectbox("👥 Grupo:", options=grupos_disponiveis, index=0, key=f"grupo_{tab_key}")

    with col3:
        df_vendas["Mes/Ano"] = df_vendas["Data"].dt.strftime("%m/%Y")

        def _ord_key(mmyyyy: str):
            try:
                return datetime.strptime("01/" + str(mmyyyy), "%d/%m/%Y")
            except Exception:
                return datetime.min

        meses_disponiveis = sorted([m for m in df_vendas["Mes/Ano"].dropna().unique()], key=_ord_key)
        mes_atual = datetime.today().strftime("%m/%Y")
        if meses_disponiveis:
            default_meses = [mes_atual] if mes_atual in meses_disponiveis else [meses_disponiveis[-1]]
        else:
            default_meses = []

        if meses_disponiveis:
            meses_sel = st.multiselect(
                "🗓️ Selecione os meses:",
                options=meses_disponiveis,
                default=default_meses,
                key=f"ms_meses_{tab_key}"
            )
        else:
            st.warning("⚠️ Nenhum mês disponível nos dados (verifique a coluna 'Data').")
            meses_sel = []

    # ==== Aplica filtros ====
    if meses_sel:
        df_filtrado = df_vendas[df_vendas["Mes/Ano"].isin(meses_sel)].copy()
    else:
        df_filtrado = df_vendas.iloc[0:0].copy()

    df_filtrado["Período"] = df_filtrado["Data"].dt.strftime("%m/%Y")

    if tipo_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Tipo"] == tipo_sel]
    if grupo_sel != "Todos":
        df_filtrado = df_filtrado[df_filtrado["Grupo"] == grupo_sel]

    # ==== Agrupamento dinâmico ====
    if metric_col not in df_filtrado.columns:
        df_filtrado[metric_col] = 0.0

    if grupo_sel == "Todos":
        chaves = ["Tipo", "Grupo"]
    else:
        chaves = ["Grupo", "Loja"]

    df_agrupado = df_filtrado.groupby(chaves + ["Período"], as_index=False)[metric_col].sum()
    df_final = df_agrupado.groupby(chaves, as_index=False)[metric_col].sum()
    df_final.rename(columns={metric_col: "Total"}, inplace=True)
    df_final["Rateio"] = 0.0

    # ==== % e Subtotais ====
    if grupo_sel == "Todos":
        total_geral = df_final["Total"].sum()
        df_final["% Total"] = (df_final["Total"] / total_geral) if total_geral else 0.0

        subtotais_tipo = df_final.groupby("Tipo")["Total"].sum().reset_index().sort_values(by="Total", ascending=False)
        ordem_tipos = subtotais_tipo["Tipo"].tolist()

        df_final["ord_tipo"] = df_final["Tipo"].apply(lambda x: ordem_tipos.index(x) if x in ordem_tipos else 999)
        df_final = df_final.sort_values(by=["ord_tipo", "Total"], ascending=[True, False]).drop(columns="ord_tipo")

        linhas = []
        for tipo in ordem_tipos:
            bloco = df_final[df_final["Tipo"] == tipo].copy()
            linhas.append(bloco)
            subtotal = bloco.drop(columns=["Tipo", "Grupo"]).sum(numeric_only=True)
            subtotal["Tipo"] = tipo
            subtotal["Grupo"] = f"Subtotal {tipo}"
            linhas.append(pd.DataFrame([subtotal]))
        df_final = pd.concat(linhas, ignore_index=True)

    else:
        total_geral = df_final["Total"].sum()
        df_final["% Total"] = (df_final["Total"] / total_geral) if total_geral else 0.0

        df_final = df_final.sort_values(by=["Grupo", "Total"], ascending=[True, False])

        linhas = []
        for g in df_final["Grupo"].unique():
            bloco = df_final[df_final["Grupo"] == g].copy()
            linhas.append(bloco)
            subtotal = bloco.drop(columns=["Grupo", "Loja"]).sum(numeric_only=True)
            subtotal["Grupo"] = g
            subtotal["Loja"] = f"Subtotal {g}"
            linhas.append(pd.DataFrame([subtotal]))
        df_final = pd.concat(linhas, ignore_index=True)

    # ==== Linha TOTAL no topo ====
    cols_drop = [c for c in ["Tipo", "Grupo", "Loja"] if c in df_final.columns]
    apenas = df_final.copy()
    for col in cols_drop:
        apenas = apenas[~apenas[col].astype(str).str.startswith("Subtotal", na=False)]
    linha_total = apenas.drop(columns=cols_drop, errors="ignore").sum(numeric_only=True)
    for col in cols_drop:
        linha_total[col] = ""
    linha_total[cols_drop[0] if cols_drop else "Grupo"] = "TOTAL"
    df_final = pd.concat([pd.DataFrame([linha_total]), df_final], ignore_index=True)

    # ==== Rateio ====
    df_final["% Total"] = 0.0
    df_final["Rateio"] = 0.0

    if grupo_sel == "Todos":
        def moeda_para_float(s: str) -> float:
            try:
                return float(s.replace(".", "").replace(",", "."))
            except:
                return 0.0

        tipos_unicos = [
            t for t in df_final["Tipo"].dropna().unique()
            if str(t).strip() not in ["", "TOTAL"] and not str(t).startswith("Subtotal")
        ]
        valores_rateio = {}
        COLS_POR_LINHA = 3
        for i in range(0, len(tipos_unicos), COLS_POR_LINHA):
            linha = tipos_unicos[i:i+COLS_POR_LINHA]
            cols = st.columns(len(linha))
            for c, tipo in zip(cols, linha):
                with c:
                    valor_str = st.text_input(f"💰 Rateio — {tipo}", value="0,00", key=f"rateio_{tipo}_{tab_key}")
                    valores_rateio[tipo] = moeda_para_float(valor_str)

        for tipo in df_final["Tipo"].unique():
            mask = (
                (df_final["Tipo"] == tipo) &
                (~df_final["Grupo"].astype(str).str.startswith("Subtotal")) &
                (df_final["Grupo"] != "TOTAL")
            )
            subtotal_tipo = df_final.loc[df_final["Grupo"] == f"Subtotal {tipo}", "Total"].sum()

            if subtotal_tipo > 0:
                df_final.loc[mask, "% Total"] = (df_final.loc[mask, "Total"] / subtotal_tipo) * 100
            df_final.loc[df_final["Grupo"] == f"Subtotal {tipo}", "% Total"] = 100

            valor_rateio = valores_rateio.get(tipo, 0.0)
            df_final.loc[mask, "Rateio"] = df_final.loc[mask, "% Total"] / 100 * valor_rateio

            df_final.loc[df_final["Grupo"] == f"Subtotal {tipo}", "Rateio"] = df_final.loc[mask, "Rateio"].sum()

    else:
        total_rateio = st.number_input(
            f"💰 Rateio — {grupo_sel}",
            min_value=0.0, step=100.0, format="%.2f",
            key=f"rateio_{grupo_sel}_{tab_key}"
        )

        mask_lojas = (
            (df_final["Grupo"] == grupo_sel) &
            (~df_final["Loja"].astype(str).str.startswith("Subtotal")) &
            (df_final["Loja"] != "TOTAL")
        )

        subtotal_grupo = df_final.loc[df_final["Loja"] == f"Subtotal {grupo_sel}", "Total"].sum()

        if subtotal_grupo > 0:
            df_final.loc[mask_lojas, "% Total"] = (df_final.loc[mask_lojas, "Total"] / subtotal_grupo) * 100
            df_final.loc[df_final["Loja"] == f"Subtotal {grupo_sel}", "% Total"] = 100
            df_final.loc[mask_lojas, "Rateio"] = df_final.loc[mask_lojas, "% Total"] / 100 * total_rateio
            df_final.loc[df_final["Loja"] == f"Subtotal {grupo_sel}", "Rateio"] = df_final.loc[mask_lojas, "Rateio"].sum()

    # ==== Reordenar colunas e formatar para exibição ====
    colunas_existentes = [c for c in ["Tipo", "Grupo", "Loja", "Total", "% Total", "Rateio"] if c in df_final.columns]
    df_final = df_final[colunas_existentes]
    df_view = df_final.copy()

    for col in ["Total", "Rateio"]:
        if col in df_view.columns:
            df_view[col] = df_view[col].apply(lambda x: formatar_moeda_br(x) if pd.notnull(x) and x != "" else x)

    if "% Total" in df_view.columns:
        df_view["% Total"] = pd.to_numeric(df_view["% Total"], errors="coerce").apply(lambda x: f"{x:.2f}%" if pd.notnull(x) else "")

    st.dataframe(aplicar_estilo_df(df_view), use_container_width=True, height=700)

    # ==== Exportações ====
    output_final = exportar_excel(df_final.copy())
    st.download_button(
        label="📥 Baixar Excel",
        data=output_final,
        file_name=f"Resumo_{titulo_aba.replace(' ', '_')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"dl_excel_{tab_key}"
    )

    usuario_logado = st.session_state.get("usuario_logado", "Usuário Desconhecido")
    sele = st.session_state.get(f"ms_meses_{tab_key}", [])
    if not sele:
        mes_rateio = "(sem dados)"
    elif len(sele) == 1:
        mes_rateio = sele[0]
    elif len(sele) == 2:
        mes_rateio = f"{sele[0]} e {sele[1]}"
    else:
        mes_rateio = f"{', '.join(sele[:-1])} e {sele[-1]}"

    pdf_bytes = gerar_pdf(df_view, mes_rateio=mes_rateio, usuario=usuario_logado)
    st.download_button(
        label="📄 Baixar PDF",
        data=pdf_bytes,
        file_name=f"{titulo_aba.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d')}.pdf",
        mime="application/pdf",
        key=f"dl_pdf_{tab_key}"
    )

# ======================
# Carga e normalização dos dados
# ======================
with st.spinner("⏳ Processando..."):
    # 1) Conexão com Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gc = gspread.authorize(credentials)
    planilha_empresa = gc.open("Vendas diarias")

    # 2) Dados base
    df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())
    df_vendas  = pd.DataFrame(planilha_empresa.worksheet("Fat Sistema Externo").get_all_records())

    # 3) Normalização
    df_empresa.columns = df_empresa.columns.str.strip()
    df_vendas.columns  = df_vendas.columns.str.strip()

    if "Loja" in df_empresa.columns:
        df_empresa["Loja"] = df_empresa["Loja"].astype(str).str.strip().str.upper()
    if "Grupo" in df_empresa.columns:
        df_empresa["Grupo"] = df_empresa["Grupo"].astype(str).str.strip()

    if "Data" in df_vendas.columns:
        df_vendas["Data"] = pd.to_datetime(df_vendas["Data"], dayfirst=True, errors="coerce")
    if "Loja" in df_vendas.columns:
        df_vendas["Loja"] = df_vendas["Loja"].astype(str).str.strip().str.upper()
    if "Grupo" in df_vendas.columns:
        df_vendas["Grupo"] = df_vendas["Grupo"].astype(str).str.strip()

    # Merge com Tipo
    if "Tipo" in df_empresa.columns and "Loja" in df_empresa.columns and "Loja" in df_vendas.columns:
        df_vendas = df_vendas.merge(df_empresa[["Loja", "Tipo"]], on="Loja", how="left")
    else:
        df_vendas["Tipo"] = df_vendas.get("Tipo", "")

    # Ajusta Fat.Total (moeda para float)
    if "Fat.Total" in df_vendas.columns:
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
        df_vendas["Fat.Total"] = pd.to_numeric(df_vendas["Fat.Total"], errors="coerce").fillna(0.0)
    else:
        df_vendas["Fat.Total"] = 0.0

    # ======================
    # Tabs
    # ======================
    aba1, aba2 = st.tabs(["📄 %Faturamento", "🔄 Volumetria"])

    with aba1:
        render_aba("📄 % Faturamento", "fat", df_empresa, df_vendas, metric_col="Fat.Total")

    with aba2:
        # Por enquanto, mesma métrica (vamos ajustar depois para a coluna de volume)
        render_aba("🔄 Volumetria", "vol", df_empresa, df_vendas, metric_col="Fat.Total")
