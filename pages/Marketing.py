# pages/Importar_Vendas_Materiais.py
import streamlit as st
import pandas as pd
import numpy as np
import re
import unicodedata
from io import BytesIO

# ===== Google Sheets =====
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Materiais por Grupo e Loja", layout="wide")

# ---------- Helpers ----------
def _strip_invisibles(s: str) -> str:
    return re.sub(r"[\u200B-\u200D\uFEFF]", "", str(s or ""))

def _ns(s: str) -> str:
    s = str(s or "").strip().lower()
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"\s+", " ", s)
    return s

def _norm_loja(s: str) -> str:
    s = str(s or "").strip()
    # remove prefixo numérico "123 - ..." (com ou sem hífen)
    s = re.sub(r"^\s*\d+\s*[-–]?\s*", "", s)
    return s.strip().lower()

def _to_float_brl(x):
    s = str(x or "").strip()
    s = s.replace("R$", "")
    s = s.replace("\u00A0", "")
    s = s.replace(".", "")
    s = s.replace(",", ".")
    s = re.sub(r"[^\d\.\-]", "", s)
    try:
        return float(s)
    except:
        return np.nan

def _to_float(x):
    try:
        return float(str(x).replace(",", "."))
    except:
        return np.nan

# ---------- Detecta a linha de subcabeçalho (onde aparecem "Qtde" e "Valor") ----------
def localizar_linha_qtde_valor(df: pd.DataFrame) -> int | None:
    # olha só as 60 primeiras linhas para acelerar
    lim = min(60, len(df))
    for r in range(lim):
        vals = [ _ns(df.iat[r, c]) for c in range(df.shape[1]) ]
        # precisamos de pares (qtde, valor)
        if "qtde" in vals and any("valor" in v for v in vals):
            return r
    return None

# ---------- Mapeia colunas de lojas (pares Qtde/Valor) ----------
def mapear_lojas(df: pd.DataFrame, r_sub: int):
    header = [ _ns(df.iat[r_sub, c]) for c in range(df.shape[1]) ]

    def captura_loja_acima(c):
        # tenta nas 3 linhas acima pegar o "título" da loja
        for up in (1,2,3):
            r = r_sub - up
            if r < 0: break
            raw = str(df.iat[r, c]).strip()
            if raw and _ns(raw) not in ("qtde","valor","valor (r$)"):
                return raw
        return ""

    lojas = []
    c = 0
    while c < len(header):
        eh_qtde = header[c] == "qtde"
        eh_val  = c+1 < len(header) and ("valor" in header[c+1])
        if eh_qtde and eh_val:
            loja_raw = captura_loja_acima(c) or captura_loja_acima(c+1)
            if loja_raw:
                lojas.append({
                    "loja_raw": loja_raw,
                    "loja": _norm_loja(loja_raw),
                    "col_qtde": c,
                    "col_valor": c+1,
                })
            c += 2
        else:
            c += 1
    return lojas

# ---------- Parser principal conforme suas regras ----------
def parse_materiais_por_grupo_loja(df_raw: pd.DataFrame) -> pd.DataFrame:
    """
    Regras:
      - Loja está na linha (acima dos pares 'Qtde' e 'Valor(R$)').
      - Grupo do produto sempre na coluna B.
      - Uma linha com 'subtotal' (coluna B) indica que o grupo terminou e um novo grupo virá.
      - Produto referente ao grupo está na coluna D.
      - Quantidade e Valor(R$) ficam nas colunas abaixo de cada Loja (pares Qtde/Valor).
    Saída: linhas por (Loja, Grupo, Produto) com Qtde e Valor (R$).
    """
    # limpar invisíveis e padronizar colunas
    df = df_raw.copy()
    for c in range(df.shape[1]):
        df.iloc[:, c] = df.iloc[:, c].map(_strip_invisibles)

    r_sub = localizar_linha_qtde_valor(df)
    if r_sub is None:
        raise ValueError("Não encontrei a linha com subcabeçalho 'Qtde' / 'Valor'.")

    lojas = mapear_lojas(df, r_sub)
    if not lojas:
        raise ValueError("Não identifiquei pares de colunas (Qtde/Valor) por Loja.")

    # índices fixos conforme especificação
    IDX_B = 1   # coluna B → Grupo
    IDX_D = 3   # coluna D → Produto

    resultados = []
    grupo_atual = None

    for r in range(r_sub + 1, df.shape[0]):
        grupo_b = str(df.iat[r, IDX_B]).strip() if IDX_B < df.shape[1] else ""
        if grupo_b:  # sempre olhar B primeiro
            if "subtotal" in _ns(grupo_b):   # sinaliza fim de um grupo
                # não registrar nada nesta linha; próximo B não-vazio será um novo grupo
                continue
            # B não contém "subtotal" → é nome de grupo
            grupo_atual = grupo_b

        produto = str(df.iat[r, IDX_D]).strip() if IDX_D < df.shape[1] else ""

        # se não houver grupo ainda ou produto vazio, não há item; segue
        if not grupo_atual or not produto:
            continue

        # para cada loja, capturar qtde/valor
        for lj in lojas:
            qtde  = _to_float(df.iat[r, lj["col_qtde"]]) if lj["col_qtde"] < df.shape[1] else np.nan
            valor = _to_float_brl(df.iat[r, lj["col_valor"]]) if lj["col_valor"] < df.shape[1] else np.nan

            # descarta linhas completamente vazias
            if (pd.isna(qtde) or qtde == 0) and (pd.isna(valor) or abs(valor) < 1e-9):
                continue

            resultados.append({
                "Loja": lj["loja"],
                "Loja (original)": lj["loja_raw"],
                "Grupo": grupo_atual,
                "Produto": produto,
                "Qtde": qtde if not pd.isna(qtde) else 0.0,
                "Valor (R$)": valor if not pd.isna(valor) else 0.0,
            })

    out = pd.DataFrame(resultados)
    # ordena para facilitar leitura
    if not out.empty:
        out = out.sort_values(["Loja","Grupo","Produto"]).reset_index(drop=True)
    return out

# ---------- Google Sheets: carregar Tabela Empresa ----------
def carregar_tabela_empresa_gsheets(nome_planilha="Vendas diarias", aba="Tabela Empresa") -> pd.DataFrame:
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials_dict = st.secrets["GOOGLE_SERVICE_ACCOUNT"]  # mesmo nome que você já usa
    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gc = gspread.authorize(creds)
    sh = gc.open(nome_planilha)
    ws = sh.worksheet(aba)
    df_emp = pd.DataFrame(ws.get_all_records())
    # normaliza cabeçalho
    cols = {c: _strip_invisibles(c).strip() for c in df_emp.columns}
    df_emp = df_emp.rename(columns=cols)
    # renomeia campos mais usados
    ren = {}
    for c in df_emp.columns:
        cn = _ns(c)
        if "codigo" in cn and "everest" in cn and "grupo" not in cn:
            ren[c] = "Código Everest"
        elif cn == "loja":
            ren[c] = "Loja"
        elif cn == "grupo":
            ren[c] = "Grupo (Empresa)"
        elif "codigo" in cn and "grupo" in cn and "everest" in cn:
            ren[c] = "Código Grupo Everest"
    df_emp = df_emp.rename(columns=ren)
    # chave normalizada da loja
    if "Loja" in df_emp.columns:
        df_emp["Loja_norm"] = df_emp["Loja"].astype(str).str.strip().str.lower().map(_norm_loja)
    else:
        df_emp["Loja_norm"] = ""
    return df_emp

def juntar_com_tabela_empresa(df_items: pd.DataFrame, df_emp: pd.DataFrame) -> pd.DataFrame:
    if df_items.empty or df_emp.empty:
        return df_items
    look = df_emp.set_index("Loja_norm")
    key = df_items["Loja"].astype(str).str.strip().str.lower().map(_norm_loja)
    df = df_items.copy()
    for col in ["Loja","Grupo (Empresa)","Código Everest","Código Grupo Everest"]:
        if col not in look.columns:
            look[col] = ""
    df["Loja (Cadastro)"] = key.map(look["Loja"])
    df["Grupo (Empresa)"] = key.map(look["Grupo (Empresa)"])
    df["Código Everest"] = key.map(look["Código Everest"])
    df["Código Grupo Everest"] = key.map(look["Código Grupo Everest"])
    return df

# ---------- Exportar Excel ----------
def exportar_excel(df: pd.DataFrame, nome="materiais_por_grupo_loja.xlsx"):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="Dados", index=False)
        wb = writer.book
        ws = writer.sheets["Dados"]
        fmt_header = wb.add_format({"bold": True, "bg_color": "#F2F2F2", "border": 1})
        fmt_border = wb.add_format({"border": 1})
        # auto largura
        for i, col in enumerate(df.columns):
            try:
                largura = max(10, min(45, int(df[col].astype(str).str.len().quantile(0.95)) + 2))
            except Exception:
                largura = 18
            ws.set_column(i, i, largura)
        ws.set_row(0, None, fmt_header)
        ws.conditional_format(0, 0, len(df), len(df.columns)-1, {"type":"no_blanks","format":fmt_border})
        ws.conditional_format(0, 0, len(df), len(df.columns)-1, {"type":"blanks","format":fmt_border})
    buffer.seek(0)
    st.download_button(
        "⬇️ Baixar Excel (Materiais x Loja/Grupo)",
        data=buffer,
        file_name=nome,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# ================= UI =================
st.title("📦 Materiais por Grupo e Loja (layout com Grupo em B e Produto em D)")
st.caption("Detecta pares Qtde/Valor por loja, usa Grupo (col. B) e Produto (col. D), e junta com a Tabela Empresa do Google Sheets.")

uploaded = st.file_uploader("Envie o Excel (ex.: venda-de-materiais-por-grupo-e-loja.xlsx)", type=["xlsx","xls"])

col1, col2 = st.columns(2)
with col1:
    nome_planilha = st.text_input("Planilha no Google Sheets", value="Vendas diarias")
with col2:
    aba_empresa = st.text_input("Aba com Tabela Empresa", value="Tabela Empresa")

if uploaded is not None:
    try:
        # Lemos SEM cabeçalho para trabalhar com posições fixas (B, D, etc.)
        df_raw = pd.read_excel(uploaded, sheet_name=0, header=None, dtype=object)
    except Exception as e:
        st.error(f"❌ Não consegui ler o Excel: {e}")
        st.stop()

    try:
        df_itens = parse_materiais_por_grupo_loja(df_raw)
        if df_itens.empty:
            st.warning("⚠️ Nenhum item encontrado conforme as regras fornecidas.")
        else:
            # carrega Tabela Empresa e junta
            try:
                df_emp = carregar_tabela_empresa_gsheets(nome_planilha, aba_empresa)
            except Exception as e:
                st.error(f"⚠️ Li os itens, mas falhei ao carregar Tabela Empresa: {e}")
                df_emp = pd.DataFrame()

            df_final = juntar_com_tabela_empresa(df_itens, df_emp) if not df_emp.empty else df_itens

            st.subheader("Prévia")
            st.dataframe(df_final.head(200), use_container_width=True, hide_index=True)
            st.info(f"Linhas totais: {len(df_final):,}".replace(",", "."))

            # lista lojas/grupos que não casaram
            if "Código Everest" in df_final.columns:
                lojas_sem_codigo = df_final[df_final["Código Everest"].isna()]["Loja"].dropna().unique().tolist()
                if lojas_sem_codigo:
                    st.warning("⚠️ Lojas sem Código Everest cadastrado: " + ", ".join(sorted(set(lojas_sem_codigo))))

            exportar_excel(df_final)

    except Exception as e:
        st.error(f"❌ Erro ao processar pelas regras do relatório: {e}")

else:
    st.info("Envie o arquivo para começar.")
