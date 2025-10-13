# pages/Importar_Materiais.py
import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
from io import BytesIO

# Google Sheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Materiais por Loja (com Tabela Empresa)", layout="wide")

# =============== Helpers ===============
def _ns(s: str) -> str:
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"\s+", " ", s)
    return s

def _is_empty_cell(x) -> bool:
    if x is None:
        return True
    s = str(x).strip()
    return s == "" or s.lower() in {"nan", "none"}

def _to_float_brl(x):
    if x is None:
        return np.nan
    s = str(x).strip()
    s = s.replace("R$", "").replace("\u00A0", "")
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        try:
            return float(str(x))
        except:
            return np.nan

def _to_float_qtde(x):
    if x is None:
        return np.nan
    s = str(x).strip()
    s = re.sub(r"[^\d,.\-]", "", s)
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    elif s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        try:
            return float(str(x))
        except:
            return np.nan

def _norm_loja(s: str) -> str:
    s = str(s or "").strip()
    # remove prefixo "123 - " se existir
    s = re.sub(r"^\d+\s*-\s*", "", s).strip()
    s = s.lower()
    return s

# ======= Google Sheets: Tabela Empresa =======
def carregar_tabela_empresa(nome_planilha="Vendas diarias", aba="Tabela Empresa") -> pd.DataFrame:
    import json
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    # 1) Lê as credenciais do st.secrets (pode ser dict ou string JSON)
    creds_any = None
    if "GOOGLE_SERVICE_ACCOUNT" in st.secrets:
        creds_any = st.secrets["GOOGLE_SERVICE_ACCOUNT"]
    elif "gcp_service_account" in st.secrets:
        # fallback para outro nome que você já usou
        creds_any = st.secrets["gcp_service_account"]
    else:
        st.error("🚫 Não encontrei credenciais em st.secrets['GOOGLE_SERVICE_ACCOUNT'] nem 'gcp_service_account'.")
        st.stop()

    if isinstance(creds_any, str):
        try:
            creds_dict = json.loads(creds_any)
        except Exception:
            st.error("🚫 As credenciais do Google vieram como string mas não são um JSON válido.")
            st.stop()
    elif isinstance(creds_any, dict):
        creds_dict = creds_any
    else:
        st.error("🚫 Formato de credenciais inválido em st.secrets.")
        st.stop()

    # 2) Autentica e lê a planilha
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    gc = gspread.authorize(credentials)
    sh = gc.open(nome_planilha)
    ws = sh.worksheet(aba)
    df = pd.DataFrame(ws.get_all_records())
    if df.empty:
        return pd.DataFrame(columns=["Loja","Grupo","Código Everest","Código Grupo Everest"])

    # 3) Normaliza colunas (aceitando variações de nomes)
    def _ns(s: str) -> str:
        import unicodedata, re
        s = str(s or "").strip().lower()
        s = unicodedata.normalize("NFD", s)
        s = "".join(c for c in s if unicodedata.category(c) != "Mn")
        return re.sub(r"\s+", " ", s)

    def pick(colnames, targets):
        m = {_ns(c): c for c in colnames}
        for t in targets:
            k = _ns(t)
            if k in m:
                return m[k]
        return None

    cols = df.columns.tolist()
    col_loja  = pick(cols, ["Loja"])
    col_grupo = pick(cols, ["Grupo","Operação"])
    col_cod   = pick(cols, ["Código Everest","Codigo Everest","Cod Everest"])
    col_codg  = pick(cols, [
        "Código Grupo Everest","Codigo Grupo Everest","Cod Grupo Empresas","Código Grupo Empresas"
    ])

    out = pd.DataFrame()
    out["Loja"] = df[col_loja].astype(str).str.strip() if col_loja else ""
    out["Loja_norm"] = out["Loja"].str.replace(r"^\d+\s*-\s*", "", regex=True).str.strip().str.lower()
    out["Grupo"] = df[col_grupo].astype(str).str.strip() if col_grupo else ""
    out["Código Everest"] = pd.to_numeric(df[col_cod], errors="coerce") if col_cod else pd.NA
    out["Código Grupo Everest"] = pd.to_numeric(df[col_codg], errors="coerce") if col_codg else pd.NA
    return out


# ======= Detectar blocos de loja (linha 4 = nomes; linha 5 = Qtde/Valor) =======
RE_QTDE  = re.compile(r"\bqtde\b|\bqtd\b", re.I)
RE_VALOR = re.compile(r"valor|\bvalor\s*\(r\$\)", re.I)

def detectar_blocos_lojas(df_raw: pd.DataFrame):
    """
    Lê a linha 4 (index 3) para nomes de loja e a linha 5 (index 4) para localizar
    pares de colunas Qtde / Valor (R$). Retorna lista de dicts:
      {"loja_raw": "...", "loja_norm": "...", "col_qtde": X, "col_valor": Y}
    """
    if df_raw.shape[0] < 5:
        return []

    header_lojas = df_raw.iloc[3, :]  # linha 4
    header_tipos = df_raw.iloc[4, :]  # linha 5

    blocos = []
    ncols = df_raw.shape[1]
    c = 0
    while c < ncols:
        h_tipo = str(header_tipos.iloc[c]) if c < len(header_tipos) else ""
        if RE_QTDE.search(_ns(h_tipo)):
            # queremos encontrar VALOR à direita (mesma loja)
            j = c + 1
            while j < ncols:
                h_tipo_j = str(header_tipos.iloc[j])
                if RE_VALOR.search(_ns(h_tipo_j)):
                    loja_raw = str(header_lojas.iloc[j] if j < len(header_lojas) else header_lojas.iloc[c]).strip()
                    if _is_empty_cell(loja_raw):
                        # às vezes o nome está na coluna do Qtde
                        loja_raw = str(header_lojas.iloc[c]).strip()
                    loja_norm = _norm_loja(loja_raw)
                    blocos.append({
                        "loja_raw": loja_raw,
                        "loja_norm": loja_norm,
                        "col_qtde": c,
                        "col_valor": j,
                    })
                    c = j + 1
                    break
                j += 1
            else:
                c += 1
        else:
            c += 1
    return blocos

# ======= Parser de materiais =======
RE_SUBTOTAL = re.compile(r"\bsub\.?total\b|\bsubtotal\b", re.I)

def extrair_registros(df_raw: pd.DataFrame, blocos: list) -> pd.DataFrame:
    """
    Dados:
      - Linha 5 (index 4): cabeçalhos "GRUPO", "CÓDIGO", "MATERIAL", "Qtde"/"Valor (R$)"
      - A partir da linha 6 (index 5): dados
      - Col B (index 1): grupo (só na primeira linha de cada grupo; depois vazio) → carry-forward
        * Se contiver 'Sub.Total'/'Subtotal': ignora a linha e reseta grupo/código
      - Col C (index 2): código do material; se vazio, herda da linha anterior dentro do mesmo grupo
      - Col D (index 3): nome do material (obrigatório)
      - Só registra se Valor (R$) > 0 para a loja
    """
    registros = []
    if df_raw.shape[0] < 6:
        return pd.DataFrame(registros)

    grupo_atual = None
    last_code = None

    # linhas de dados começam no index 5
    for r in range(5, df_raw.shape[0]):
        # --- coluna B: grupo / subtotal / vazio
        b_raw = df_raw.iat[r, 1] if df_raw.shape[1] > 1 else ""
        b_txt = "" if _is_empty_cell(b_raw) else str(b_raw).strip()
        b_ns  = _ns(b_txt)

        # header "grupo" (na linha 5) não deve ser tratado como grupo de dados
        if b_txt.lower() == "grupo":
            continue

        # subtotal → reseta e pula
        if b_txt and RE_SUBTOTAL.search(b_ns):
            grupo_atual = None
            last_code = None
            continue

        # se apareceu texto em B (não subtotal), é o nome do grupo desta “seção”
        if b_txt:
            grupo_atual = b_txt
            last_code = None  # novo grupo → zera o carry de código
            # segue para próxima linha (os itens do grupo começam nas linhas seguintes)
            continue

        # Se não temos grupo atual, não é linha de item ainda
        if not grupo_atual:
            continue

        # --- coluna C: código (pode estar vazio e deve herdar)
        cod_raw = df_raw.iat[r, 2] if df_raw.shape[1] > 2 else ""
        codigo  = "" if _is_empty_cell(cod_raw) else str(cod_raw).strip()
        if codigo == "" and last_code:
            codigo = last_code
        if codigo != "":
            last_code = codigo

        # --- coluna D: material
        mat_raw = df_raw.iat[r, 3] if df_raw.shape[1] > 3 else ""
        material = "" if _is_empty_cell(mat_raw) else str(mat_raw).strip()

        # se não tem material, pula (é obrigatório)
        if material == "":
            continue

        # Para cada loja (pares Qtde/Valor)
        for b in blocos:
            qtde_raw  = df_raw.iat[r, b["col_qtde"]] if b["col_qtde"] < df_raw.shape[1] else None
            valor_raw = df_raw.iat[r, b["col_valor"]] if b["col_valor"] < df_raw.shape[1] else None
            qtde = _to_float_qtde(qtde_raw)
            valor = _to_float_brl(valor_raw)

            if pd.isna(valor) or float(valor) <= 0:
                continue

            registros.append({
                "Loja_norm": b["loja_norm"],
                "Grupo do Produto": str(grupo_atual).strip(),
                "Código Material": str(codigo),
                "Material": material,
                "Qtde": float(qtde) if pd.notna(qtde) else np.nan,
                "Valor (R$)": float(valor),
            })

    return pd.DataFrame(registros)

# ======= Merge com Tabela Empresa =======
def juntar_tabela_empresa(df_items: pd.DataFrame, df_emp: pd.DataFrame) -> pd.DataFrame:
    if df_items.empty:
        return df_items
    look = df_emp.set_index("Loja_norm")
    df = df_items.copy()
    df["Loja"] = df["Loja_norm"]  # inicial
    # recuperar campos
    for col_src, col_dst in [("Loja","Loja"), ("Grupo","Operação"),
                             ("Código Everest","Código Everest"),
                             ("Código Grupo Everest","Código Grupo Everest")]:
        if col_src in look.columns:
            df[col_dst] = df["Loja_norm"].map(look[col_src] if col_src=="Loja" else look[col_src])
        else:
            if col_dst not in df.columns:
                df[col_dst] = ""

    # Loja (exibição) deve ser o nome “bonito” da tabela
    df["Loja"] = df["Loja_norm"].map(look["Loja"]).fillna(df["Loja"])
    # Operação é o Grupo da Tabela Empresa
    if "Operação" not in df.columns and "Grupo" in look.columns:
        df["Operação"] = df["Loja_norm"].map(look["Grupo"])

    # ordenar e limpar
    cols_final = [
        "Operação","Loja","Grupo do Produto",
        "Código Material","Material","Qtde","Valor (R$)",
        "Código Everest","Código Grupo Everest"
    ]
    for c in cols_final:
        if c not in df.columns:
            df[c] = ""
    df = df[cols_final].copy()
    return df

# =============== UI ===============
st.title("📦 Materiais por Loja — Upload + Tabela Empresa")
st.caption("Lê o Excel, herda Grupo (coluna B) e Código (coluna C), ignora Sub.Total e traz Operação/lojas/códigos da Tabela Empresa.")

col1, col2 = st.columns([1,1])
with col1:
    nome_planilha = st.text_input("Nome da planilha (Google Sheets)", value="Vendas diarias")
with col2:
    aba_empresa   = st.text_input("Aba da Tabela Empresa", value="Tabela Empresa")

up = st.file_uploader("Envie o Excel original", type=["xlsx","xls"])

if up is not None:
    try:
        df_raw = pd.read_excel(up, sheet_name=0, header=None, dtype=object, engine="openpyxl")
    except Exception as e:
        st.error(f"❌ Não consegui ler o Excel: {e}")
        st.stop()

    # detectar lojas e pares Qtde/Valor na linha 4/5
    blocos = detectar_blocos_lojas(df_raw)
    if not blocos:
        st.error("❌ Não encontrei pares de colunas Qtde/Valor (R$) na linha 5. Confirme o layout.")
        st.stop()

    # extrair itens
    df_itens = extrair_registros(df_raw, blocos)
    if df_itens.empty:
        st.warning("Nenhum item elegível encontrado (somente valor > 0 por loja).")
        st.stop()

    # carregar Tabela Empresa
    try:
        df_emp = carregar_tabela_empresa(nome_planilha, aba_empresa)
    except Exception as e:
        st.error(f"❌ Erro ao carregar Tabela Empresa: {e}")
        st.stop()

    # normalizar chave de loja para join
    df_emp["Loja_norm"] = df_emp["Loja"].map(_norm_loja)

    df_final = juntar_tabela_empresa(df_itens, df_emp)

    st.subheader("Prévia")
    st.dataframe(df_final.head(100), use_container_width=True, hide_index=True)

    st.info(f"Linhas totais: {len(df_final):,}".replace(",", "."))

    # download
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
        df_final.to_excel(w, index=False, sheet_name="MateriaisPorLoja")
    buf.seek(0)
    st.download_button("⬇️ Baixar Excel (Materiais por Loja)", data=buf,
                       file_name="materiais_por_loja.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Envie o arquivo Excel para começar.")
