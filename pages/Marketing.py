# pages/Importar_Vendas_Materiais.py
import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
from io import BytesIO

# ====== Dependências Google Sheets (ajuste conforme seu projeto) ======
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Vendas por Grupo e Loja (com Tabela Empresa)", layout="wide")

# ======================
# Helpers
# ======================
def _ns(s: str) -> str:
    """normaliza texto: minúsculo, sem acento e sem espaços duplos"""
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("utf-8")
    s = re.sub(r"\s+", " ", s)
    return s

def _strip_invisible_cols(cols):
    return [re.sub(r"[\u200B-\u200D\uFEFF]", "", str(c)).strip() for c in cols]

def detectar_linha_cabecalho(df_raw: pd.DataFrame):
    """
    Tenta encontrar a linha que contém nomes de colunas (ex.: 'Código Everest', 'Loja', etc.)
    Retorna índice da linha de cabeçalho; se não achar, assume linha 0.
    """
    alvos = [
        "codigo everest", "cod everest", "codigo loja", "cod loja",
        "loja", "grupo", "valor", "venda", "material", "descricao", "data"
    ]
    for idx, row in df_raw.iterrows():
        linha = [_ns(x) for x in row.astype(str).tolist()]
        score = sum(any(a in c for a in alvos) for c in linha)
        # heurística simples: se >= 2 termos relevantes, considero cabeçalho
        if score >= 2:
            return idx
    return 0

def limpar_dataframe(df: pd.DataFrame):
    # remove colunas "Unnamed"
    keep = [c for c in df.columns if _ns(c) and not _ns(c).startswith("unnamed")]
    df = df.loc[:, keep].copy()
    # tira espaços invisíveis
    df.columns = _strip_invisible_cols(df.columns)
    return df

def carregar_tabela_empresa_gsheets(nome_planilha="Vendas diarias", aba="Tabela Empresa"):
    # >>>> IMPORTANTE <<<<
    # Ajuste o caminho do JSON ou use st.secrets["gcp_service_account"]
    # Exemplo com arquivo local:
    # creds = ServiceAccountCredentials.from_json_keyfile_name("service_account.json", scopes=[
    #     "https://spreadsheets.google.com/feeds",
    #     "https://www.googleapis.com/auth/drive"
    # ])
    # Exemplo com st.secrets:
    creds = ServiceAccountCredentials.from_json_keyfile_dict(st.secrets["gcp_service_account"], scopes=[
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ])
    gc = gspread.authorize(creds)
    sh = gc.open(nome_planilha)
    ws = sh.worksheet(aba)
    df_emp = pd.DataFrame(ws.get_all_records())
    # limpeza
    df_emp.columns = _strip_invisible_cols(df_emp.columns)
    # normaliza nomes esperados
    ren = {}
    for c in df_emp.columns:
        cn = _ns(c)
        if cn in ["codigo everest", "cod everest", "codigo loja", "cod loja"]:
            ren[c] = "Código Everest"
        elif cn == "loja":
            ren[c] = "Loja"
        elif cn == "grupo":
            ren[c] = "Grupo"
        elif cn == "tipo":
            ren[c] = "Tipo"
    df_emp = df_emp.rename(columns=ren)
    return df_emp

def merge_com_tabela_empresa(df_base: pd.DataFrame, df_emp: pd.DataFrame):
    cols_emp = [c for c in ["Código Everest", "Loja", "Grupo", "Tipo"] if c in df_emp.columns]
    df_emp = df_emp[cols_emp].drop_duplicates()

    tem_cod_base = "Código Everest" in df_base.columns
    tem_loja_base = "Loja" in df_base.columns

    if tem_cod_base and "Código Everest" in df_emp.columns:
        out = df_base.merge(df_emp, on="Código Everest", how="left")
    elif tem_loja_base and "Loja" in df_emp.columns:
        out = df_base.merge(df_emp, on="Loja", how="left")
    else:
        # tenta mapear por colunas parecidas
        poss_cod = [c for c in df_base.columns if _ns(c) in ["codigo everest","cod everest","codigo loja","cod loja"]]
        if poss_cod and "Código Everest" in df_emp.columns:
            out = df_base.rename(columns={poss_cod[0]: "Código Everest"}).merge(df_emp, on="Código Everest", how="left")
        else:
            poss_loja = [c for c in df_base.columns if _ns(c) == "loja"]
            if poss_loja and "Loja" in df_emp.columns:
                out = df_base.rename(columns={poss_loja[0]: "Loja"}).merge(df_emp, on="Loja", how="left")
            else:
                out = df_base.copy()
                out["(Aviso)"] = "Não foi possível juntar com Tabela Empresa (sem colunas compatíveis)."
    return out

def exportar_excel_formatado(df: pd.DataFrame, nome="vendas_materiais_com_lojas.xlsx"):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Dados", index=False)
        wb = writer.book
        ws = writer.sheets["Dados"]

        # formatos
        fmt_header = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
        fmt_border = wb.add_format({"border": 1})
        # auto largura
        for i, col in enumerate(df.columns):
            largura = max(10, min(45, int(df[col].astype(str).str.len().fillna(0).quantile(0.95)) + 2))
            ws.set_column(i, i, largura)
        # header
        ws.set_row(0, None, fmt_header)
        # bordas
        ws.conditional_format(0, 0, len(df), len(df.columns)-1,
                              {"type": "no_blanks", "format": fmt_border})
        ws.conditional_format(0, 0, len(df), len(df.columns)-1,
                              {"type": "blanks", "format": fmt_border})
    buffer.seek(0)
    st.download_button("⬇️ Baixar Excel (com lojas)", data=buffer, file_name=nome, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ======================
# UI
# ======================
st.title("📦 Vendas de Materiais por Grupo e Loja — com Tabela Empresa")
st.caption("Lê o Excel exportado, detecta o cabeçalho, limpa colunas e junta com a aba **Tabela Empresa** do seu Google Sheets.")

up = st.file_uploader("Envie o relatório Excel (ex.: venda-de-materiais-por-grupo-e-loja.xlsx)", type=["xlsx","xls"])

col_cfg1, col_cfg2 = st.columns([1,1])
with col_cfg1:
    nome_planilha = st.text_input("Nome da planilha no Google Sheets", value="Vendas diarias")
with col_cfg2:
    aba_empresa = st.text_input("Aba com a Tabela Empresa", value="Tabela Empresa")

if up is not None:
    try:
        df_raw = pd.read_excel(up, sheet_name=0, header=None, dtype=object)
    except Exception as e:
        st.error(f"Não consegui ler o Excel: {e}")
        st.stop()

    # detectar e aplicar cabeçalho
    idx_header = detectar_linha_cabecalho(df_raw)
    df = df_raw.iloc[idx_header:].copy()
    df.columns = df.iloc[0].astype(str).tolist()
    df = df.iloc[1:].reset_index(drop=True)

    # limpeza colunas
    df.columns = _strip_invisible_cols(df.columns)
    df = limpar_dataframe(df)

    # normalizações comuns
    # tenta converter Data (se existir)
    poss_data = [c for c in df.columns if _ns(c) in ["data","dt","competencia","d.competencia","d.lancamento","d.lançamento"]]
    for c in poss_data:
        df[c] = pd.to_datetime(df[c], errors="coerce")

    # carrega Tabela Empresa
    try:
        df_emp = carregar_tabela_empresa_gsheets(nome_planilha, aba_empresa)
    except Exception as e:
        st.error(f"Erro ao carregar a Tabela Empresa do Google Sheets: {e}")
        st.stop()

    # merge
    df_merge = merge_com_tabela_empresa(df, df_emp)

    st.subheader("Prévia dos dados (com lojas/empresa)")
    st.dataframe(df_merge.head(100), use_container_width=True, hide_index=True)

    st.info(f"Linhas totais: {len(df_merge):,}".replace(",", "."))

    exportar_excel_formatado(df_merge)
else:
    st.info("Envie o arquivo Excel para começar.")
