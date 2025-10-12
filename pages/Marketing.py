# pages/Importar_Materiais_por_Loja.py
import streamlit as st
import pandas as pd
import numpy as np
import re, unicodedata, json
from io import BytesIO

import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Materiais × Lojas (com Operação)", layout="wide")

# ---------------- Helpers ----------------
def _strip_invis(s: str) -> str:
    return re.sub(r"[\u200B-\u200D\uFEFF]", "", str(s or ""))

def _ns(s: str) -> str:
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return re.sub(r"\s+", " ", s)

def _norm_loja(s: str) -> str:
    # remove prefixo "123 - " e normaliza
    s = str(s or "").strip()
    s = re.sub(r"^\s*\d+\s*[-–]?\s*", "", s)
    s = unicodedata.normalize("NFD", s)
    s = "".join(ch for ch in s if unicodedata.category(ch) != "Mn")
    return s.strip().lower()

# --- Google Sheets auth (lê de st.secrets) ---
def _get_service_account_dict() -> dict:
    for k in ("GOOGLE_SERVICE_ACCOUNT", "gcp_service_account"):
        if k in st.secrets:
            raw = st.secrets[k]
            return json.loads(raw) if isinstance(raw, str) else raw
    raise RuntimeError("Defina a credencial no st.secrets (GOOGLE_SERVICE_ACCOUNT ou gcp_service_account).")

def _get_gspread_client():
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(_get_service_account_dict(), scope)
    return gspread.authorize(creds)

def carregar_tabela_empresa(nome_planilha="Vendas diarias", aba="Tabela Empresa") -> pd.DataFrame:
    gc = _get_gspread_client()
    ws = gc.open(nome_planilha).worksheet(aba)
    df = pd.DataFrame(ws.get_all_records())
    df.columns = [_strip_invis(c).strip() for c in df.columns]

    ren = {}
    for c in df.columns:
        cn = _ns(c)
        if cn == "loja": ren[c] = "Loja"
        elif cn == "grupo": ren[c] = "Operação"         # <- já renomeio para Operação
        elif ("codigo" in cn and "everest" in cn and "grupo" not in cn): ren[c] = "Código Everest"
        elif ("codigo" in cn and "grupo" in cn and "everest" in cn):     ren[c] = "Código Grupo Everest"
    df = df.rename(columns=ren)
    if "Loja" not in df.columns:
        df["Loja"] = ""
    df["Loja_norm"] = df["Loja"].astype(str).map(_norm_loja)
    return df[["Loja","Loja_norm","Operação"]]

# ---------------- Parser do Excel ----------------
def detectar_blocos_loja(df_raw: pd.DataFrame):
    """
    Lojas na linha 4 (index 3). Na linha 5 (index 4) ficam 'Qtde' e 'Valor(R$)'.
    Retorna lista de blocos: [{'col_qtde':c, 'col_valor':c+1, 'loja_raw': ... , 'loja_norm': ...}, ...]
    """
    r_lojas = 3  # linha 4
    r_sub    = 4  # linha 5
    header_row = [_ns(df_raw.iat[r_sub, c]) for c in range(df_raw.shape[1])]
    blocos = []
    c = 0
    while c < len(header_row):
        eh_qtde = header_row[c] == "qtde"
        eh_val  = (c+1 < len(header_row)) and ("valor" in header_row[c+1])
        if eh_qtde and eh_val:
            loja_raw = str(df_raw.iat[r_lojas, c]).strip() or str(df_raw.iat[r_lojas, c+1]).strip()
            if loja_raw:
                loja_norm = _norm_loja(loja_raw)
                if loja_norm and not re.search(r"\b(total|subtotal)\b", _ns(loja_raw)):
                    blocos.append({"col_qtde": c, "col_valor": c+1, "loja_raw": loja_raw, "loja_norm": loja_norm})
            c += 2
        else:
            c += 1
    return blocos

def extrair_registros(df_raw: pd.DataFrame, blocos: list) -> pd.DataFrame:
    """
    A partir da linha 6 (index 5):
      - Grupo do produto: coluna B (index 1)
      - Código material:  coluna C (index 2)
      - Material:         coluna D (index 3)
      - Para cada bloco (loja): Qtde/Valor nas colunas do bloco
    Ignora linhas cuja coluna B contenha 'total' ou 'subtotal'.
    """
    registros = []
    for r in range(5, df_raw.shape[0]):
        grupo_prod = str(df_raw.iat[r, 1]).strip()
        if not grupo_prod:
            # linha vazia — pode pular
            continue
        if re.search(r"\b(total|subtotal)\b", _ns(grupo_prod)):
            # linha de total/subtotal (fim de grupo) — ignora
            continue
        cod_mat = df_raw.iat[r, 2]
        mat     = df_raw.iat[r, 3]

        for b in blocos:
            qtde  = df_raw.iat[r, b["col_qtde"]]
            valor = df_raw.iat[r, b["col_valor"]]

            # considera só quando houver quantidade ou valor
            if (pd.isna(qtde) or str(qtde).strip() == "") and (pd.isna(valor) or str(valor).strip() == ""):
                continue

            try:
                qtde_num = float(str(qtde).replace(",", "."))
            except:
                qtde_num = pd.to_numeric(qtde, errors="coerce")
            try:
                valor_num = float(
                    str(valor).replace("R$", "").replace(".", "").replace(",", ".").strip()
                )
            except:
                valor_num = pd.to_numeric(valor, errors="coerce")

            registros.append({
                "Loja_norm": b["loja_norm"],
                "Grupo do Produto": grupo_prod,
                "Código Material": str(cod_mat).strip(),
                "Material": str(mat).strip(),
                "Qtde": float(qtde_num) if pd.notna(qtde_num) else np.nan,
                "Valor (R$)": float(valor_num) if pd.notna(valor_num) else np.nan,
            })
    return pd.DataFrame(registros)

# ---------------- UI ----------------
st.title("📦 Materiais por Loja — com Operação (Tabela Empresa)")
st.caption("Lê o Excel (lojas na linha 4; Qtde/Valor na linha 5) e cruza com a Tabela Empresa para trazer **Operação** e **Loja** (oficial).")

col1, col2 = st.columns(2)
with col1:
    nome_planilha = st.text_input("Nome da planilha no Google Sheets", value="Vendas diarias")
with col2:
    aba_empresa = st.text_input("Aba da Tabela Empresa", value="Tabela Empresa")

uploaded = st.file_uploader("Envie o Excel", type=["xlsx","xls","xlsm"])

if uploaded:
    # 1) Lê o Excel sem header
    try:
        df_raw = pd.read_excel(uploaded, sheet_name=0, header=None, dtype=object)
        for c in range(df_raw.shape[1]):
            df_raw.iloc[:, c] = df_raw.iloc[:, c].map(_strip_invis)
    except Exception as e:
        st.error(f"Não consegui ler o Excel: {e}")
        st.stop()

    # 2) Detecta blocos de lojas
    blocos = detectar_blocos_loja(df_raw)
    if not blocos:
        st.error("Não encontrei pares 'Qtde'/'Valor(R$)' na linha 5.")
        st.stop()

    # 3) Extrai os registros (uma linha por material/loja com Qtde/Valor)
    df_itens = extrair_registros(df_raw, blocos)
    if df_itens.empty:
        st.warning("Nenhum item com Qtde/Valor encontrado.")
        st.stop()

    # 4) Carrega Tabela Empresa e cruza para obter Operação e Loja oficial
    try:
        df_emp = carregar_tabela_empresa(nome_planilha, aba_empresa)
    except Exception as e:
        st.error(f"Erro ao carregar Tabela Empresa do Google Sheets: {e}")
        st.stop()

    df_final = df_itens.merge(df_emp, on="Loja_norm", how="left")

    # 5) Seleção e ordem de colunas finais
    cols_final = ["Operação", "Loja", "Grupo do Produto", "Código Material", "Material", "Qtde", "Valor (R$)"]
    for c in cols_final:
        if c not in df_final.columns:
            df_final[c] = np.nan
    df_final = df_final[cols_final]

    # 6) Exibe + download
    st.subheader("Prévia")
    st.dataframe(df_final.head(200), use_container_width=True, hide_index=True)

    st.info(f"Linhas: {len(df_final):,}".replace(",", "."))

    faltando = df_final[df_final["Operação"].isna() | (df_final["Operação"].astype(str).str.strip() == "")]
    if not faltando.empty:
        st.warning(
            "Existem lojas sem correspondência na **Tabela Empresa**. "
            "Atualize a aba e reenvie. Lojas afetadas (amostra): "
            + ", ".join(sorted(faltando["Loja"].dropna().astype(str).unique())[:10])
        )

    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as wr:
        df_final.to_excel(wr, index=False, sheet_name="MateriaisPorLoja")
    buf.seek(0)
    st.download_button(
        "⬇️ Baixar Excel (Materiais × Lojas)",
        data=buf,
        file_name="materiais_por_loja_com_operacao.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.info("Envie o arquivo Excel para começar.")
