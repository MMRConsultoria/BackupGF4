# pages/Importar_Vendas_Materiais.py
import streamlit as st
import pandas as pd
import numpy as np
import re
from io import BytesIO
import json

# Google Sheets
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Vendas por Grupo e Loja (Materiais)", layout="wide")

# ======================
# Helpers de normalização
# ======================
def _ns(s: str) -> str:
    s = str(s or "").strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s

def _limpa_loja(loja: str) -> str:
    """remove prefixo 'NN - ' e espaços dobrados."""
    s = str(loja or "").strip()
    s = re.sub(r"^\s*\d+\s*-\s*", "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s

def _to_number(x):
    """Converte (br/pt/en) para float. Vazio vira NaN."""
    s = str(x or "").strip()
    if s == "" or s.lower() == "nan":
        return np.nan
    s = s.replace("R$", "").replace("\u00A0", "")
    s = re.sub(r"[^\d,.\-]", "", s)
    # Se tem vírgula e ponto, assume pt-BR -> remove pontos, troca vírgula por ponto
    if s.count(",") == 1 and s.count(".") >= 1:
        s = s.replace(".", "").replace(",", ".")
    # Se tem vírgula e não tem ponto, assume pt-BR -> vírgula vira ponto
    elif s.count(",") == 1 and s.count(".") == 0:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        return np.nan

# ======================
# Google Sheets: Tabela Empresa
# ======================
def carregar_tabela_empresa(nome_planilha="Vendas diarias", aba="Tabela Empresa") -> pd.DataFrame:
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]

    # aceita GOOGLE_SERVICE_ACCOUNT como dict ou string JSON
    # e também gcp_service_account (alguns projetos usam esse nome)
    raw = None
    try:
        raw = st.secrets.get("GOOGLE_SERVICE_ACCOUNT", None)
        if raw is None:
            raw = st.secrets.get("gcp_service_account", None)
    except Exception:
        raw = None

    if raw is None:
        raise RuntimeError(
            "Credenciais não encontradas em st.secrets. "
            "Defina GOOGLE_SERVICE_ACCOUNT (ou gcp_service_account)."
        )

    # se vier string, fazer json.loads; se já for dict, usa direto
    if isinstance(raw, str):
        credentials_dict = json.loads(raw)
    else:
        credentials_dict = dict(raw)

    creds = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gc = gspread.authorize(creds)

    ws = gc.open(nome_planilha).worksheet(aba)
    df = pd.DataFrame(ws.get_all_records())
    if df.empty:
        return pd.DataFrame(columns=["Loja","Grupo","Código Everest","Código Grupo Everest"])

    df.columns = [str(c).strip() for c in df.columns]

    # localizar colunas com tolerância de grafia
    def _ns(s: str) -> str:
        return re.sub(r"\s+", " ", str(s or "").strip().lower())

    def pick(cols, alvo, *alts):
        alvo_norm = _ns(alvo)
        cand = [alvo, *alts]
        for wanted in cand:
            wn = _ns(wanted)
            for c in cols:
                if _ns(c) == wn:
                    return c
        return None

    c_loja  = pick(df.columns, "Loja")
    c_grupo = pick(df.columns, "Grupo")
    c_cod   = pick(df.columns, "Código Everest", "Codigo Everest", "Cod Everest", "Codigo Ev", "Cód Everest")
    c_codg  = pick(df.columns, "Código Grupo Everest", "Codigo Grupo Everest", "Cod Grupo Empresas", "Codigo Grupo Empresas")

    keep = {c_loja:"Loja", c_grupo:"Grupo", c_cod:"Código Everest", c_codg:"Código Grupo Everest"}
    # remove None que não foram achados
    keep = {k:v for k,v in keep.items() if k is not None}

    df = df[list(keep.keys())].rename(columns=keep)

    # normalização para o join
    df["Loja"] = df["Loja"].astype(str).str.strip()
    df["Loja_norm"] = df["Loja"].str.lower()
    return df


# ======================
# Parser do Excel de Upload
# ======================
def parse_excel_materiais(file) -> pd.DataFrame:
    """
    Regras:
      - Lojas estão na linha 4 (células mescladas), cada loja ocupa 2 colunas (Qtde / Valor(R$)) na linha 5
      - Grupo de material: coluna B (apenas em algumas linhas) -> forward-fill
      - Código: coluna C (nem todas as linhas têm) -> forward-fill quando Material preenchido
      - Material: coluna D
      - Ignorar linhas 'Sub.Total' (coluna C), 'Total' em lojas, 'Total Geral'
      - Itens: trazer somente se Qtde e Valor forem > 0 (ou pelo menos Valor > 0, conforme sua regra)
    Retorna: DataFrame com colunas base (antes do merge):
      ['Loja','GrupoProduto','Codigo','Material','Qtde','Valor']
    """
    # Lê planilha principal (primeira aba)
    df_raw = pd.read_excel(file, sheet_name=0, header=None, dtype=object)

    # Linha 4 (index 3) = nomes das lojas (mescladas)
    # Linha 5 (index 4) = cabeçalho ("Qtde" / "Valor(R$)")
    if df_raw.shape[0] < 6:
        return pd.DataFrame(columns=["Loja","GrupoProduto","Codigo","Material","Qtde","Valor"])

    header_row = 4  # linha 5 (0-based = 4)
    lojas_row  = 3  # linha 4 (0-based = 3)

    # Identificar os pares de colunas (Qtde / Valor)
    col_pairs = []  # [(col_qtde, col_valor, loja_name), ...]
    c = 0
    ncols = df_raw.shape[1]

    # Primeiro, extrai os nomes de loja da linha lojas_row.
    # Em planilhas com mesclagem, só uma das duas colunas do par pode carregar o nome;
    # por isso usamos "preencher para a direita" (ffill axis=1) numa cópia dessa linha.
    lojas_line = df_raw.iloc[lojas_row:lojas_row+1, :].copy()
    lojas_line = lojas_line.ffill(axis=1).iloc[0].tolist()  # lista de nomes repetidos

    while c < ncols - 1:
        h1 = str(df_raw.iat[header_row, c]).strip().lower() if c < ncols else ""
        h2 = str(df_raw.iat[header_row, c+1]).strip().lower() if c+1 < ncols else ""
        if "qtde" in h1 and "valor" in h2:
            loja_bruta = lojas_line[c] if c < len(lojas_line) else ""
            loja = _limpa_loja(loja_bruta).strip()
            if loja == "" or _ns(loja).startswith("total"):
                # ignora colunas cuja "loja" é Total / Total Geral
                c += 2
                continue
            col_pairs.append((c, c+1, loja))
            c += 2
        else:
            c += 1

    # Colunas de base (A,B,C,D) => 0,1,2,3
    COL_GRUPO   = 1
    COL_CODIGO  = 2
    COL_MATERIAL= 3

    # Faixa de dados começa depois do cabeçalho (linha 6 para baixo: index 5+)
    start_row = header_row + 1

    # Vamos construir itens linha a linha (para cada linha do produto, replicar por loja)
    registros = []
    grupo_atual = None
    cod_atual = None

    for r in range(start_row, df_raw.shape[0]):
        grupo_raw = str(df_raw.iat[r, COL_GRUPO]) if COL_GRUPO < ncols else ""
        cod_raw   = df_raw.iat[r, COL_CODIGO] if COL_CODIGO < ncols else ""
        mat_raw   = df_raw.iat[r, COL_MATERIAL] if COL_MATERIAL < ncols else ""

        grupo_txt = str(grupo_raw or "").strip()
        cod_txt   = str(cod_raw   or "").strip()
        mat_txt   = str(mat_raw   or "").strip()

        # Ignora linhas “Total Geral” explícitas
        if re.search(r"total\s*geral", f"{grupo_txt} {cod_txt} {mat_txt}", flags=re.I):
            continue

        # Detecta linha de Sub.Total (fica na coluna C, segundo seu relato)
        if re.search(r"sub\.?\s*total", cod_txt, flags=re.I):
            # Ao encontrar Sub.Total, apenas marca “fronteira” de grupo e PULA o registro
            # (grupo_atual só troca quando B tiver novo nome; aqui só ignora a linha)
            continue

        # Quando coluna B (grupo) não está vazia, atualiza grupo_atual
        grp_clean = grupo_txt.strip()
        if grp_clean != "" and not re.search(r"sub\.?\s*total", grp_clean, flags=re.I):
            grupo_atual = grp_clean

        # Material vazio => nada para registrar
        if mat_txt == "" or _ns(mat_txt) == "nan":
            continue

        # Código: se vier vazio, mantém o último (forward-fill)
        if cod_txt != "" and not re.search(r"sub\.?\s*total", cod_txt, flags=re.I):
            cod_atual = cod_txt
        # Se ainda assim não temos código, skipa — (ou mantém vazio se preferir)
        codigo_final = cod_atual if (cod_atual is not None and str(cod_atual).strip() != "") else ""

        # Para cada par de loja (Qtde/Valor), cria uma linha
        for (c_q, c_v, loja) in col_pairs:
            qtde_v = _to_number(df_raw.iat[r, c_q]) if c_q < ncols else np.nan
            val_v  = _to_number(df_raw.iat[r, c_v]) if c_v < ncols else np.nan

            # Ignora lojas "Total" (já tratadas ao montar col_pairs), mas reforça aqui
            if _ns(loja).startswith("total"):
                continue

            # Regras: se qtde vazia/0 ou valor vazio/0 -> não traz
            if pd.isna(qtde_v) or pd.isna(val_v) or qtde_v == 0 or val_v == 0:
                continue

            registros.append([
                loja,                  # Loja
                grupo_atual or "",     # GrupoProduto (Grupo do material)
                codigo_final,          # Codigo
                mat_txt,               # Material
                float(qtde_v),         # Qtde
                float(val_v),          # Valor
            ])

    df_items = pd.DataFrame(registros, columns=["Loja","GrupoProduto","Codigo","Material","Qtde","Valor"])
    return df_items

# ======================
# Enriquecimento com Tabela Empresa
# ======================
def enriquecer_com_tabela_empresa(df_items: pd.DataFrame, df_emp: pd.DataFrame) -> pd.DataFrame:
    # Normaliza lojas
    df_items = df_items.copy()
    df_items["Loja"] = df_items["Loja"].astype(str).map(_limpa_loja)
    df_items["Loja_norm"] = df_items["Loja"].str.lower()

    df_emp = df_emp.copy()
    if "Loja_norm" not in df_emp.columns:
        df_emp["Loja"] = df_emp["Loja"].astype(str).str.strip()
        df_emp["Loja_norm"] = df_emp["Loja"].str.lower()

    merged = df_items.merge(
        df_emp[["Loja_norm","Loja","Grupo","Código Everest","Código Grupo Everest"]],
        on="Loja_norm", how="left", suffixes=("_x","_y")
    )

    # Loja preferindo a da Tabela Empresa
    if "Loja_x" in merged.columns and "Loja_y" in merged.columns:
        merged["Loja"] = merged["Loja_y"].where(
            merged["Loja_y"].astype(str).str.strip() != "", merged["Loja_x"]
        )
        merged.drop(columns=["Loja_x","Loja_y"], inplace=True)
    elif "Loja_y" in merged.columns:
        merged["Loja"] = merged["Loja_y"]; merged.drop(columns=["Loja_y"], inplace=True)
    elif "Loja_x" in merged.columns:
        merged["Loja"] = merged["Loja_x"]; merged.drop(columns=["Loja_x"], inplace=True)

    # Renomeia para modelo final
    merged = merged.rename(columns={
        "Grupo": "Operação",            # Grupo da Tabela Empresa = Operação
        "GrupoProduto": "Grupo Material",
        "Codigo": "Codigo Material",
    })

    final_cols = [
        "Operação","Loja","Código Everest","Código Grupo Everest",
        "Grupo Material","Codigo Material","Material","Qtde","Valor"
    ]
    for c in final_cols:
        if c not in merged.columns:
            merged[c] = ""

    df_final = merged[final_cols].copy()
    return df_final

# ======================
# UI
# ======================
st.title("📦 Importar Vendas de Materiais por Grupo e Loja")
st.caption("Lê o Excel do PDV (lojas na linha 4; Qtde/Valor na linha 5), elimina Sub.Total/Total, e completa com a Tabela Empresa do Google Sheets.")

colA, colB = st.columns([1,1])
with colA:
    nome_planilha = st.text_input("Planilha do Google Sheets", value="Vendas diarias")
with colB:
    aba_empresa   = st.text_input("Aba da Tabela Empresa", value="Tabela Empresa")

up = st.file_uploader("Envie o Excel (venda-de-materiais-por-grupo-e-loja.xlsx)", type=["xlsx","xls"])

if up is not None:
    try:
        with st.spinner("🔎 Lendo arquivo e montando itens..."):
            df_items = parse_excel_materiais(up)

        if df_items.empty:
            st.warning("Nenhum item elegível foi encontrado (verifique se há Qtde/Valor > 0 e se o layout segue a regra da linha 4/5).")
            st.stop()

        # Carrega Tabela Empresa
        with st.spinner("🔗 Carregando Tabela Empresa..."):
            df_emp = carregar_tabela_empresa(nome_planilha, aba_empresa)

        # Enriquecer com Tabela Empresa
        with st.spinner("🧭 Normalizando lojas e anexando códigos..."):
            df_final = enriquecer_com_tabela_empresa(df_items, df_emp)

        st.success(f"✅ Itens processados: {len(df_final):,}".replace(",", "."))

        st.subheader("Prévia")
        st.dataframe(df_final.head(200), use_container_width=True, hide_index=True)

        # Download Excel
        def _to_excel(df: pd.DataFrame) -> BytesIO:
            buf = BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
                df.to_excel(w, index=False, sheet_name="Materiais")
                ws = w.sheets["Materiais"]
                for i, col in enumerate(df.columns):
                    width = max(12, min(45, int(df[col].astype(str).str.len().quantile(0.95)) + 2))
                    ws.set_column(i, i, width)
            buf.seek(0)
            return buf

        excel_bytes = _to_excel(df_final)
        st.download_button(
            "⬇️ Baixar Excel (Materiais por Loja)",
            data=excel_bytes,
            file_name="materiais_por_loja.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Avisos úteis
        faltantes = df_final[df_final["Código Everest"].astype(str).str.strip().isin(["","nan"])].copy()
        if not faltantes.empty:
            lojas_faltantes = sorted(faltantes["Loja"].dropna().astype(str).unique().tolist())
            if lojas_faltantes:
                st.warning(
                    "⚠️ Algumas lojas não foram localizadas na **Tabela Empresa**. "
                    "Atualize a planilha e reprocese:\n\n- " + "\n- ".join(lojas_faltantes)
                )

    except KeyError as e:
        st.error(f"❌ Erro de colunas: {e}")
    except Exception as e:
        st.error(f"❌ Erro ao processar: {e}")
else:
    st.info("Envie o arquivo Excel para começar.")
