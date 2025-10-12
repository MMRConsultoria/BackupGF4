# pages/Importar_Vendas_Materiais.py
import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
from io import BytesIO

# ====== Google Sheets ======
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import json  # p/ st.secrets["GOOGLE_SERVICE_ACCOUNT"]

st.set_page_config(page_title="Vendas por Grupo e Loja (com Tabela Empresa)", layout="wide")

# ======================
# Helpers base
# ======================
def _ns(s: str) -> str:
    """normaliza texto: minúsculo, sem acento e sem espaços duplos"""
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("utf-8")
    s = re.sub(r"\s+", " ", s)
    return s

def _strip_invisible_cols(cols):
    return [re.sub(r"[\u200B-\u200D\uFEFF]", "", str(c)).strip() for c in cols]

def _dedupe_cols(df: pd.DataFrame) -> pd.DataFrame:
    """garante nomes únicos: Loja, Loja.1, Loja.2 ..."""
    new_cols, seen = [], {}
    for c in [str(x) for x in df.columns]:
        if c not in seen:
            seen[c] = 0
            new_cols.append(c)
        else:
            seen[c] += 1
            new_cols.append(f"{c}.{seen[c]}")
    df.columns = new_cols
    return df

def detectar_linha_cabecalho(df_raw: pd.DataFrame):
    """
    Acha a linha que parece cabeçalho (2+ hits de termos conhecidos). Se não achar, 0.
    """
    alvos = [
        "codigo everest", "cod everest", "codigo loja", "cod loja",
        "loja", "grupo", "valor", "venda", "material", "descricao", "data"
    ]
    for idx, row in df_raw.iterrows():
        linha = [_ns(x) for x in row.astype(str).tolist()]
        score = sum(any(a in c for a in alvos) for c in linha)
        if score >= 2:
            return idx
    return 0

def limpar_dataframe(df: pd.DataFrame):
    keep = [c for c in df.columns if _ns(c) and not _ns(c).startswith("unnamed")]
    df = df.loc[:, keep].copy()
    df.columns = _strip_invisible_cols(df.columns)
    df = _dedupe_cols(df)
    return df

# ======================
# Conexão / Tabela Empresa (GS)
# ======================
def carregar_tabela_empresa_gsheets(nome_planilha="Vendas diarias", aba="Tabela Empresa"):
    """
    Lê a 'Tabela Empresa' do Google Sheets via st.secrets["GOOGLE_SERVICE_ACCOUNT"].
    Renomeia colunas usuais para: Código Everest, Loja, Grupo, Tipo, Código Grupo Everest.
    """
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive",
    ]
    try:
        credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    except KeyError:
        raise RuntimeError("st.secrets['GOOGLE_SERVICE_ACCOUNT'] não encontrado.")

    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gc = gspread.authorize(creds)

    sh = gc.open(nome_planilha)
    ws = sh.worksheet(aba)
    df_emp = pd.DataFrame(ws.get_all_records())

    df_emp.columns = _strip_invisible_cols(df_emp.columns)
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
        elif cn in ["codigo grupo everest", "cod grupo everest"]:
            ren[c] = "Código Grupo Everest"
    df_emp = df_emp.rename(columns=ren)
    df_emp = _dedupe_cols(df_emp)
    return df_emp

# ======================
# Merge genérico com Tabela Empresa (fallback)
# ======================
def merge_com_tabela_empresa(df_base: pd.DataFrame, df_emp: pd.DataFrame):
    cols_emp = [c for c in ["Código Everest", "Loja", "Grupo", "Tipo", "Código Grupo Everest"] if c in df_emp.columns]
    df_emp2 = df_emp[cols_emp].drop_duplicates().copy()

    tem_cod_base = "Código Everest" in df_base.columns
    tem_loja_base = "Loja" in df_base.columns

    join_col = None
    if tem_cod_base and "Código Everest" in df_emp2.columns:
        join_col = "Código Everest"
    elif tem_loja_base and "Loja" in df_emp2.columns:
        join_col = "Loja"
    else:
        poss_cod = [c for c in df_base.columns if _ns(c) in ["codigo everest","cod everest","codigo loja","cod loja"]]
        if poss_cod and "Código Everest" in df_emp2.columns:
            df_base = df_base.rename(columns={poss_cod[0]: "Código Everest"})
            join_col = "Código Everest"
        else:
            poss_loja = [c for c in df_base.columns if _ns(c) == "loja"]
            if poss_loja and "Loja" in df_emp2.columns:
                df_base = df_base.rename(columns={poss_loja[0]: "Loja"})
                join_col = "Loja"

    if join_col is None:
        out = df_base.copy()
        out["(Aviso)"] = "Não foi possível juntar com Tabela Empresa (sem colunas compatíveis)."
        return _dedupe_cols(out)

    merged = df_base.merge(df_emp2, on=join_col, how="left", suffixes=("", "_tbemp"))
    for col in ["Loja", "Grupo", "Tipo", "Código Grupo Everest", "Código Everest"]:
        a, b = col, f"{col}_tbemp"
        if a in merged.columns and b in merged.columns:
            merged[a] = merged[a].where(merged[a].notna() & (merged[a].astype(str).str.strip() != ""), merged[b])
            merged.drop(columns=[b], inplace=True)
    return _dedupe_cols(merged)

# ======================
# Helpers do layout "matriz" (Loja no header; pares Qtde/Valor)
# ======================
def _coerce_number(x):
    if pd.isna(x):
        return 0.0
    s = str(x).strip()
    if s == "":
        return 0.0
    s = (s.replace("R$", "")
           .replace(" ", "")
           .replace(".", "")
           .replace(",", "."))
    try:
        return float(s)
    except Exception:
        try:
            return float(re.sub(r"[^0-9\.\-]", "", s))
        except Exception:
            return 0.0

def _is_layout_matriz(df_raw: pd.DataFrame) -> bool:
    """
    Há uma linha com 'Qtde' e 'Valor' (subcabeçalho por loja)?
    """
    for r in range(min(40, len(df_raw))):
        row_vals = [str(x).strip().lower() for x in df_raw.iloc[r].tolist()]
        if "qtde" in row_vals and any("valor" in x for x in row_vals):
            return True
    return False

def _find_store_columns(df_raw: pd.DataFrame):
    """
    Retorna (stores_info, header_sub_idx)
    stores_info = [{'loja': <nome>, 'qtde_col': j, 'valor_col': j+1}, ...]
    """
    header_sub_idx = None
    for r in range(min(40, len(df_raw))):
        row_vals = [str(x).strip().lower() for x in df_raw.iloc[r].tolist()]
        if "qtde" in row_vals and any("valor" in x for x in row_vals):
            header_sub_idx = r
            break
    if header_sub_idx is None:
        return [], None

    header_sub = [str(x).strip().lower() for x in df_raw.iloc[header_sub_idx].tolist()]

    def _lookup_loja_name(col_idx: int) -> str:
        # busca 1-3 linhas acima (células mescladas, etc.)
        for up in range(1, 4):
            rr = header_sub_idx - up
            if rr < 0:
                break
            val = str(df_raw.iat[rr, col_idx]).strip()
            if val and val.lower() not in ["qtde", "valor", "valor (r$)"]:
                return val
        return ""

    stores = []
    j = 0
    while j < len(header_sub):
        cell = header_sub[j]
        if cell == "qtde":
            if j + 1 < len(header_sub) and ("valor" in header_sub[j+1]):
                loja = _lookup_loja_name(j) or _lookup_loja_name(j+1)
                if loja:
                    stores.append({"loja": loja, "qtde_col": j, "valor_col": j+1})
                j += 2
                continue
        j += 1

    return stores, header_sub_idx

def processar_layout_matriz(df_raw: pd.DataFrame, df_emp: pd.DataFrame) -> pd.DataFrame:
    """
    Converte o relatório "matriz" em formato longo:
      - Grupo em coluna B; linha com 'subtotal' em B encerra o grupo
      - Produto em coluna D
      - Para cada loja (pares Qtde/Valor), gera linhas com Qtde/Valor
      - Junta com Tabela Empresa por Loja (normalizada)
    """
    stores_info, header_sub_idx = _find_store_columns(df_raw)
    if not stores_info:
        raise ValueError("Não consegui identificar os pares (Qtde/Valor) por loja no cabeçalho.")

    start_row = header_sub_idx + 1
    current_group = None
    rows = []
    ncols = df_raw.shape[1]
    has_colB, has_colD = ncols > 1, ncols > 3

    for i in range(start_row, len(df_raw)):
        # controla grupo via coluna B
        bval = str(df_raw.iat[i, 1]).strip() if has_colB else ""
        if bval:
            if "subtotal" in bval.lower():
                continue  # fim do grupo atual; o próximo texto em B será novo grupo
            else:
                current_group = bval

        # produto em D
        produto = str(df_raw.iat[i, 3]).strip() if has_colD else ""
        if not produto:
            continue

        # valores por loja
        for info in stores_info:
            loja_name = str(info["loja"]).strip()
            q = df_raw.iat[i, info["qtde_col"]] if info["qtde_col"] < ncols else None
            v = df_raw.iat[i, info["valor_col"]] if info["valor_col"] < ncols else None
            qtde = _coerce_number(q)
            valor = _coerce_number(v)
            if qtde == 0 and valor == 0:
                continue
            rows.append({
                "Loja": loja_name,
                "Grupo (Produto)": current_group if current_group else "",
                "Produto": produto,
                "Qtde": qtde,
                "Valor (R$)": valor,
            })

    df_long = pd.DataFrame(rows)
    if df_long.empty:
        return df_long

    # normaliza Loja e junta com Tabela Empresa
    df_long["__loja_norm__"] = df_long["Loja"].astype(str).str.strip().str.lower()
    emp = df_emp.copy()
    if "Loja" not in emp.columns:
        emp["Loja"] = ""
    emp["__loja_norm__"] = emp["Loja"].astype(str).str.strip().str.lower()
    cols_emp = [c for c in ["Loja","Código Everest","Grupo","Código Grupo Everest","Tipo","__loja_norm__"] if c in emp.columns]
    df_long = df_long.merge(emp[cols_emp].drop_duplicates(), on="__loja_norm__", how="left", suffixes=("", "_tbemp"))
    df_long.drop(columns=["__loja_norm__"], inplace=True, errors="ignore")

    # ordenação amigável
    ord_cols = [c for c in ["Loja","Grupo (Produto)","Produto"] if c in df_long.columns]
    if ord_cols:
        df_long = df_long.sort_values(by=ord_cols, kind="stable")

    return df_long

# ======================
# Exportação Excel
# ======================
def exportar_excel_formatado(df: pd.DataFrame, nome="vendas_materiais_com_lojas.xlsx"):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Dados", index=False)
        wb = writer.book
        ws = writer.sheets["Dados"]

        fmt_header = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
        fmt_border = wb.add_format({"border": 1})

        for i, col in enumerate(df.columns):
            try:
                largura = max(10, min(45, int(df[col].astype(str).str.len().fillna(0).quantile(0.95)) + 2))
            except Exception:
                largura = 18
            ws.set_column(i, i, largura)

        ws.set_row(0, None, fmt_header)
        ws.conditional_format(0, 0, len(df), len(df.columns)-1, {"type": "no_blanks", "format": fmt_border})
        ws.conditional_format(0, 0, len(df), len(df.columns)-1, {"type": "blanks", "format": fmt_border})

    buffer.seek(0)
    st.download_button("⬇️ Baixar Excel (com lojas)", data=buffer, file_name=nome,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ======================
# UI
# ======================
st.title("📦 Vendas de Materiais por Grupo e Loja — com Tabela Empresa")
st.caption("Lê o Excel no layout matriz (Loja em colunas com Qtde/Valor, Grupo em B, Produto em D) e junta com a aba **Tabela Empresa** do seu Google Sheets.")

up = st.file_uploader(
    "Envie o relatório Excel (ex.: venda-de-materiais-por-grupo-e-loja.xlsx)",
    type=["xlsx","xls","xlsm"]
)

col_cfg1, col_cfg2 = st.columns([1,1])
with col_cfg1:
    nome_planilha = st.text_input("Nome da planilha no Google Sheets", value="Vendas diarias")
with col_cfg2:
    aba_empresa = st.text_input("Aba com a Tabela Empresa", value="Tabela Empresa")

if up is not None:
    try:
        # lê SEM header para detectar subcabeçalho Qtde/Valor
        df_raw = pd.read_excel(up, sheet_name=0, header=None, dtype=object, engine="openpyxl")
    except Exception as e:
        st.error(f"Não consegui ler o Excel: {e}")
        st.stop()

    # carrega Tabela Empresa (GS)
    try:
        df_emp = carregar_tabela_empresa_gsheets(nome_planilha, aba_empresa)
    except Exception as e:
        st.error(f"Erro ao carregar a Tabela Empresa do Google Sheets: {e}")
        st.stop()

    # ——— Parser matriz (preferencial) ———
    if _is_layout_matriz(df_raw):
        try:
            df_merge = processar_layout_matriz(df_raw, df_emp)
        except Exception as e:
            st.error(f"Erro no parser do layout matriz: {e}")
            st.stop()
    else:
        # ——— Fallback genérico ———
        idx_header = detectar_linha_cabecalho(df_raw)
        df = df_raw.iloc[idx_header:].copy()
        df.columns = _strip_invisible_cols(df.iloc[0].astype(str).tolist())
        df = df.iloc[1:].reset_index(drop=True)
        df = limpar_dataframe(df)

        poss_data = [c for c in df.columns if _ns(c) in ["data","dt","competencia","d.competencia","d.lancamento","d.lançamento"]]
        for c in poss_data:
            df[c] = pd.to_datetime(df[c], errors="coerce")

        df_merge = merge_com_tabela_empresa(df, df_emp)

    # (opcional) consolidar duplicados por Loja + Grupo(Produto) + Produto
    if not df_merge.empty and all(c in df_merge.columns for c in ["Loja","Grupo (Produto)","Produto"]):
        df_merge = (
            df_merge
            .groupby(["Loja","Grupo (Produto)","Produto","Código Everest","Grupo","Código Grupo Everest","Tipo"], dropna=False, as_index=False)
            .agg({"Qtde":"sum","Valor (R$)":"sum"})
        )

    st.subheader("Prévia dos dados (com lojas/empresa)")
    st.dataframe(df_merge.head(100), use_container_width=True, hide_index=True)
    st.info(f"Linhas totais: {len(df_merge):,}".replace(",", "."))

    exportar_excel_formatado(df_merge, nome="vendas_materiais_com_lojas.xlsx")
else:
    st.info("Envie o arquivo Excel para começar.")
