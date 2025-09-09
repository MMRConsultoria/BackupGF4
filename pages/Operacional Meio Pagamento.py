import streamlit as st
import pandas as pd
import numpy as np
import re
import json
import unicodedata
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(page_title="Meio de Pagamento", layout="wide")

# 🔥 CSS para estilizar as abas
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

# 🔒 Bloqueia o acesso caso o usuário não esteja logado
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ======================
# CSS para esconder só a barra superior
# ======================
st.markdown("""
    <style>
        [data-testid="stToolbar"] {
            visibility: hidden;
            height: 0%;
            position: fixed;
        }
        .stSpinner {
            visibility: visible !important;
        }
    </style>
""", unsafe_allow_html=True)


# ======================
# Helpers
# ======================
def _norm(s: str) -> str:
    """Normaliza texto (sem acento, espacos extras, caixa baixa)."""
    s = unicodedata.normalize("NFKD", str(s or "")).encode("ASCII", "ignore").decode("ASCII")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s

def _is_formato2(df_headed: pd.DataFrame) -> bool:
    """Detecta arquivo 'plano' (A:F): Cod Empresa, Data, Forma, Bandeira, Tipo, Total."""
    cols = {_norm(c) for c in df_headed.columns}
    return ("data" in cols and "total" in cols
            and any("cod" in c and "empresa" in c for c in cols)
            and any("forma" in c and "pag" in c for c in cols))

def _rename_cols_formato2(df: pd.DataFrame) -> pd.DataFrame:
    """Renomeia as colunas para canônicas."""
    new_names = {}
    for c in df.columns:
        n = _norm(c)
        if "cod" in n and "empresa" in n:
            new_names[c] = "cod_empresa"
        elif n == "data":
            new_names[c] = "data"
        elif "forma" in n and "pag" in n:
            new_names[c] = "forma_pgto"
        elif "bandeira" in n:
            new_names[c] = "bandeira"
        elif "tipo" in n and "cart" in n:
            new_names[c] = "tipo_cartao"
        elif n == "total" or "valor" in n:
            new_names[c] = "total"
    return df.rename(columns=new_names)

def processar_formato2(
    df_src: pd.DataFrame,
    df_empresa: pd.DataFrame,
    df_meio_pgto_google: pd.DataFrame,
) -> pd.DataFrame:
    """Processa o Formato 2 e entrega no layout final."""
    df = _rename_cols_formato2(df_src.copy())

    req = {"cod_empresa", "data", "forma_pgto", "bandeira", "tipo_cartao", "total"}
    faltando = [c for c in req if c not in df.columns]
    if faltando:
        raise ValueError(f"Colunas obrigatórias ausentes no arquivo: {faltando}")

    # Datas e valores
    df["data"] = pd.to_datetime(df["data"], dayfirst=True, errors="coerce")
    df["total"] = pd.to_numeric(df["total"], errors="coerce").fillna(0.0)

    # Meio de Pagamento
    ban = df["bandeira"].fillna("").astype(str).str.strip()
    tip = df["tipo_cartao"].fillna("").astype(str).str.strip()
    meio_from_de = (ban + " " + tip).str.strip().map(_strip_accents_keep_case)

    meio_from_c = df["forma_pgto"].astype(str).str.strip()
    meio_from_c = meio_from_c.str.replace(r"^\d+\s*-\s*", "", regex=True).map(_strip_accents_keep_case)


    df["Meio de Pagamento"] = np.where(
        (ban != "") | (tip != ""), meio_from_de, meio_from_c
    )

    # Mapeios Tabela Meio Pagamento
    df_meio = df_meio_pgto_google.copy()
    df_meio["__meio_norm__"] = df_meio["Meio de Pagamento"].astype(str).str.strip().str.lower()
    tipo_pgto_map = dict(zip(df_meio["__meio_norm__"], df_meio["Tipo de Pagamento"].astype(str)))
    tipo_dre_map  = dict(zip(df_meio["__meio_norm__"], df_meio["Tipo DRE"].astype(str)))

    df["__meio_norm__"] = df["Meio de Pagamento"].astype(str).str.strip().str.lower()
    df["Tipo de Pagamento"] = df["__meio_norm__"].map(tipo_pgto_map).fillna("")
    df["Tipo DRE"]          = df["__meio_norm__"].map(tipo_dre_map).fillna("")
    df.drop(columns=["__meio_norm__"], inplace=True, errors="ignore")

    # Join com Tabela Empresa (Código Everest)
    emp = df_empresa.copy()
    emp["Código Everest"] = emp["Código Everest"].astype(str).str.strip()
    df["Código Everest"]  = df["cod_empresa"].astype(str).str.strip()

    df = df.merge(
        emp[["Código Everest", "Loja", "Grupo", "Código Grupo Everest"]],
        on="Código Everest", how="left"
    )

    # Datas derivadas
    dias_semana = {
        'Monday': 'segunda-feira','Tuesday': 'terça-feira','Wednesday': 'quarta-feira',
        'Thursday': 'quinta-feira','Friday': 'sexta-feira','Saturday': 'sábado','Sunday': 'domingo'
    }
    df["Dia da Semana"] = df["data"].dt.day_name().map(dias_semana)
    df["Mês"] = df["data"].dt.month.map({1:'jan',2:'fev',3:'mar',4:'abr',5:'mai',6:'jun',7:'jul',8:'ago',9:'set',10:'out',11:'nov',12:'dez'})
    df["Ano"] = df["data"].dt.year
    df["Data"] = df["data"].dt.strftime("%d/%m/%Y")

    # Valor e colunas finais
    df.rename(columns={"total": "Valor (R$)"}, inplace=True)

    df_final = df[[
        "Data", "Dia da Semana",
        "Meio de Pagamento", "Tipo de Pagamento", "Tipo DRE",
        "Loja", "Código Everest",
        "Grupo", "Código Grupo Everest",
        "Valor (R$)", "Mês", "Ano"
    ]].copy()

    try:
        df_final.sort_values(by=["Data", "Loja"], inplace=True)
    except Exception:
        pass

    return df_final
def _strip_accents_keep_case(s: str) -> str:
    """Remove acentos sem mexer em maiúsculas/minúsculas."""
    return unicodedata.normalize("NFKD", str(s or "")).encode("ASCII", "ignore").decode("ASCII")

# ======================
# Spinner + cargas do Google
# ======================
with st.spinner("⏳ Processando..."):
    # 🔌 Conexão com Google Sheets
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gc = gspread.authorize(credentials)
    planilha = gc.open("Vendas diarias")

    df_empresa = pd.DataFrame(planilha.worksheet("Tabela Empresa").get_all_records())

    # === Tabela Meio Pagamento ===
    df_meio_pgto_google = pd.DataFrame(planilha.worksheet("Tabela Meio Pagamento").get_all_records())
    df_meio_pgto_google.columns = [str(c).strip() for c in df_meio_pgto_google.columns]
    for col in ["Meio de Pagamento", "Tipo de Pagamento", "Tipo DRE"]:
        if col not in df_meio_pgto_google.columns:
            df_meio_pgto_google[col] = ""
        df_meio_pgto_google[col] = df_meio_pgto_google[col].astype(str).str.strip()
    df_meio_pgto_google["__meio_norm__"] = df_meio_pgto_google["Meio de Pagamento"].str.lower()
    df_meio_pgto_google = df_meio_pgto_google.drop_duplicates(subset=["__meio_norm__"], keep="first")

    tipo_pgto_map = dict(zip(df_meio_pgto_google["__meio_norm__"], df_meio_pgto_google["Tipo de Pagamento"]))
    tipo_dre_map  = dict(zip(df_meio_pgto_google["__meio_norm__"], df_meio_pgto_google["Tipo DRE"]))

    # 🔥 Título
    st.markdown("""
        <div style='display: flex; align-items: center; gap: 10px;'>
            <img src='https://img.icons8.com/color/48/graph.png' width='40'/>
            <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Meio de Pagamento</h1>
        </div>
    """, unsafe_allow_html=True)

    # ========================
    # 🗂️ Abas
    # ========================
    tab1, tab2 = st.tabs([
        "📥 Upload e Processamento",
        "🔄 Atualizar Google Sheets"
    ])

    # ======================
    # 📥 Aba 1
    # ======================
    with tab1:
        uploaded_file = st.file_uploader(
            label="📁 Clique para selecionar ou arraste aqui o arquivo Excel",
            type=["xlsx", "xlsm"]
        )

        if uploaded_file:
            try:
                # Tenta detectar Formato 2
                df_head = pd.read_excel(uploaded_file, sheet_name=0)  # header=0
                if _is_formato2(df_head):
                    # ➜ Processa Formato 2 (plano)
                    df_meio_pagamento = processar_formato2(
                        df_head, df_empresa, df_meio_pgto_google
                    )
                else:
                    # ➜ Não é Formato 2: volta ao fluxo antigo (Formato 1)
                    uploaded_file.seek(0)
                    xls = pd.ExcelFile(uploaded_file)
                    abas_disponiveis = xls.sheet_names
                    aba_escolhida = abas_disponiveis[0] if len(abas_disponiveis) == 1 else st.selectbox(
                        "Escolha a aba para processar", abas_disponiveis)

                    df_raw = pd.read_excel(xls, sheet_name=aba_escolhida, header=None)
                    df_raw = df_raw[~df_raw.iloc[:, 1].astype(str).str.lower().str.contains("total|subtotal", na=False)]

                    if str(df_raw.iloc[0, 1]).strip().lower() != "faturamento diário por meio de pagamento":
                        st.error("❌ A célula B1 deve conter 'Faturamento diário por meio de pagamento'.")
                        st.stop()

                    linha_inicio_dados, blocos, col, loja_atual = 5, [], 3, None

                    while col < df_raw.shape[1]:
                        valor_linha4 = str(df_raw.iloc[3, col]).strip()
                        match = re.match(r"^\d+\s*-\s*(.+)$", valor_linha4)
                        if match:
                            loja_atual = match.group(1).strip().lower()

                        meio_pgto = str(df_raw.iloc[4, col]).strip()
                        if not loja_atual or not meio_pgto or meio_pgto.lower() in ["nan", ""]:
                            col += 1
                            continue

                        linha3 = str(df_raw.iloc[2, col]).strip().lower()
                        linha5 = meio_pgto.lower()
                        if any(p in t for t in [linha3, valor_linha4.lower(), linha5] for p in ["total", "serv/tx", "total real"]):
                            col += 1
                            continue

                        try:
                            df_temp = df_raw.iloc[linha_inicio_dados:, [2, col]].copy()
                            df_temp.columns = ["Data", "Valor (R$)"]
                            df_temp = df_temp[~df_temp["Data"].astype(str).str.lower().str.contains("total|subtotal")]
                            df_temp.insert(1, "Meio de Pagamento", meio_pgto.lower())
                            df_temp.insert(2, "Loja", loja_atual)
                            blocos.append(df_temp)
                        except Exception as e:
                            st.warning(f"⚠️ Erro ao processar coluna {col}: {e}")
                        col += 1

                    if not blocos:
                        st.error("❌ Nenhum dado válido encontrado.")
                        st.stop()

                    df_meio_pagamento = pd.concat(blocos, ignore_index=True).dropna()
                    df_meio_pagamento = df_meio_pagamento[~df_meio_pagamento["Data"].astype(str).str.lower().str.contains("total|subtotal")]
                    df_meio_pagamento["Data"] = pd.to_datetime(df_meio_pagamento["Data"], dayfirst=True, errors="coerce")
                    df_meio_pagamento = df_meio_pagamento[df_meio_pagamento["Data"].notna()]

                    # Derivados
                    dias_semana = {'Monday':'segunda-feira','Tuesday':'terça-feira','Wednesday':'quarta-feira',
                                   'Thursday':'quinta-feira','Friday':'sexta-feira','Saturday':'sábado','Sunday':'domingo'}
                    df_meio_pagamento["Dia da Semana"] = df_meio_pagamento["Data"].dt.day_name().map(dias_semana)
                    df_meio_pagamento = df_meio_pagamento.sort_values(by=["Data", "Loja"])
                    df_meio_pagamento["Data"] = df_meio_pagamento["Data"].dt.strftime("%d/%m/%Y")

                    # Join com Tabela Empresa por Loja (fluxo antigo)
                    df_meio_pagamento["Loja"] = df_meio_pagamento["Loja"].str.strip().str.replace(r"^\d+\s*-\s*", "", regex=True).str.lower()
                    df_empresa["Loja"] = df_empresa["Loja"].str.strip().str.lower()
                    df_meio_pagamento = pd.merge(df_meio_pagamento, df_empresa, on="Loja", how="left")

                    # Mapeia Tipo de Pagamento / Tipo DRE
                    if "Meio de Pagamento" not in df_meio_pagamento.columns:
                        df_meio_pagamento["Meio de Pagamento"] = ""
                    df_meio_pagamento["__meio_norm__"] = (
                        df_meio_pagamento["Meio de Pagamento"].astype(str).str.strip().str.lower()
                    )
                    col_meio_idx = df_meio_pagamento.columns.get_loc("Meio de Pagamento")
                    df_meio_pagamento.insert(
                        loc=col_meio_idx + 1,
                        column="Tipo de Pagamento",
                        value=df_meio_pagamento["__meio_norm__"].map(tipo_pgto_map).fillna("")
                    )
                    df_meio_pagamento.insert(
                        loc=col_meio_idx + 2,
                        column="Tipo DRE",
                        value=df_meio_pagamento["__meio_norm__"].map(tipo_dre_map).fillna("")
                    )
                    df_meio_pagamento.drop(columns=["__meio_norm__"], inplace=True, errors="ignore")

                    # Mês/Ano
                    df_meio_pagamento["Mês"] = pd.to_datetime(df_meio_pagamento["Data"], dayfirst=True).dt.month.map({
                        1:'jan',2:'fev',3:'mar',4:'abr',5:'mai',6:'jun',7:'jul',8:'ago',9:'set',10:'out',11:'nov',12:'dez'})
                    df_meio_pagamento["Ano"] = pd.to_datetime(df_meio_pagamento["Data"], dayfirst=True).dt.year

                # Fim dos dois caminhos -> df_meio_pagamento está pronto
                st.session_state.df_meio_pagamento = df_meio_pagamento

                # KPIs de período e total
                dts = pd.to_datetime(df_meio_pagamento["Data"], dayfirst=True, errors="coerce")
                periodo_min = dts.min().strftime("%d/%m/%Y") if not dts.empty else ""
                periodo_max = dts.max().strftime("%d/%m/%Y") if not dts.empty else ""

                col1, col2 = st.columns(2)
                col1.markdown(f"<div style='font-size:1.2rem;'>📅 Período processado<br>{periodo_min} até {periodo_max}</div>", unsafe_allow_html=True)
                valor_total = f"R$ {df_meio_pagamento['Valor (R$)'].sum():,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                col2.markdown(f"<div style='font-size:1.2rem;'>💰 Valor total<br><span style='color:green;'>{valor_total}</span></div>", unsafe_allow_html=True)

                # Validação: empresas não localizadas (mostra códigos sem match)
                empresas_nao_localizadas = df_meio_pagamento[
                    df_meio_pagamento["Loja"].astype(str).str.strip().isin(["", "nan"])
                ]["Código Everest"].astype(str).unique()

                # Validação: meios não localizados na tabela
                meios_norm_tabela = set(df_meio_pgto_google["__meio_norm__"])
                meios_nao_localizados = df_meio_pagamento[
                    ~df_meio_pagamento["Meio de Pagamento"].astype(str).str.strip().str.lower().isin(meios_norm_tabela)
                ]["Meio de Pagamento"].astype(str).unique()

                if len(empresas_nao_localizadas) == 0 and len(meios_nao_localizados) == 0:
                    st.success("✅ Todas as empresas e todos os meios de pagamento foram localizados!")

                    output = BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        df_meio_pagamento.to_excel(writer, index=False, sheet_name="FaturamentoPorMeio")
                    output.seek(0)

                    st.download_button(
                        "📥 Baixar relatório Excel",
                        data=output,
                        file_name="FaturamentoPorMeio.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                else:
                    if len(empresas_nao_localizadas) > 0:
                        empresas_nao_localizadas_str = "<br>".join(empresas_nao_localizadas)
                        st.markdown(f"""
                        ⚠️ {len(empresas_nao_localizadas)} Código(s) Everest sem correspondência:<br>{empresas_nao_localizadas_str}
                        <br>✏️ Atualize a tabela clicando 
                        <a href='https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU' target='_blank'><strong>aqui</strong></a>.
                        """, unsafe_allow_html=True)
                    if len(meios_nao_localizados) > 0:
                        meios_nao_localizados_str = "<br>".join(meios_nao_localizados)
                        st.markdown(f"""
                        ⚠️ {len(meios_nao_localizados)} meio(s) de pagamento não localizado(s):<br>{meios_nao_localizados_str}
                        <br>✏️ Atualize a tabela clicando 
                        <a href='https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU' target='_blank'><strong>aqui</strong></a>.
                        """, unsafe_allow_html=True)

            except Exception as e:
                st.error(f"❌ Erro ao processar: {e}")

    # ======================
    # 🔄 Aba 2
    # ======================
    with tab2:
        st.markdown("🔗 [Abrir planilha Faturamento Meio Pagamento](https://docs.google.com/spreadsheets/d/1AVacOZDQT8vT-E8CiD59IVREe3TpKwE_25wjsj--qTU)")
    
        if 'df_meio_pagamento' not in st.session_state:
            st.warning("⚠️ Primeiro faça o upload e o processamento na Aba 1.")
        elif not all(c in st.session_state.df_meio_pagamento.columns for c in ["Meio de Pagamento","Loja","Data"]):
            st.warning("⚠️ O arquivo processado não tem as colunas necessárias.")
        else:
          
            # =======================================================
            # 1) Base: copia do que foi montado na Aba 1
            # =======================================================
            df_final = st.session_state.df_meio_pagamento.copy()
        
            # =======================================================
            # 2) Garantir "Tipo de Pagamento" e "Tipo DRE"
            #    (se por algum motivo não vieram da Aba 1, mapeamos aqui)
            # =======================================================
            # normaliza "Meio de Pagamento" para casarmos com a tabela
            df_final["Meio de Pagamento"] = (
                df_final["Meio de Pagamento"].astype(str).str.strip().str.lower()
            )
        
            # df_meio_pgto_google já foi carregado no topo do app (Aba "Tabela Meio Pagamento")
            df_meio_pgto_google.columns = [str(c).strip() for c in df_meio_pgto_google.columns]
            for c in ["Meio de Pagamento", "Tipo de Pagamento", "Tipo DRE"]:
                if c not in df_meio_pgto_google.columns:
                    df_meio_pgto_google[c] = ""
        
            df_meio_pgto_google["Meio de Pagamento"] = (
                df_meio_pgto_google["Meio de Pagamento"].astype(str).str.strip().str.lower()
            )
            df_meio_pgto_google["Tipo de Pagamento"] = (
                df_meio_pgto_google["Tipo de Pagamento"].astype(str).str.strip()
            )
            df_meio_pgto_google["Tipo DRE"] = (
                df_meio_pgto_google["Tipo DRE"].astype(str).str.strip()
            )
        
            # mapas
            pgto_map = dict(zip(df_meio_pgto_google["Meio de Pagamento"], df_meio_pgto_google["Tipo de Pagamento"]))
            dre_map  = dict(zip(df_meio_pgto_google["Meio de Pagamento"], df_meio_pgto_google["Tipo DRE"]))
        
            # chave normalizada no df_final
            df_final["__meio_norm__"] = df_final["Meio de Pagamento"].astype(str).str.strip().str.lower()
        
            # ---- Tipo de Pagamento (logo após Meio) ----
            if "Tipo de Pagamento" not in df_final.columns:
                pos = df_final.columns.get_loc("Meio de Pagamento") + 1
                df_final.insert(pos, "Tipo de Pagamento", df_final["__meio_norm__"].map(pgto_map))
            else:
                df_final["Tipo de Pagamento"] = df_final["Tipo de Pagamento"].fillna(df_final["__meio_norm__"].map(pgto_map))
        
            # ---- Tipo DRE (logo após Tipo de Pagamento) ----
            if "Tipo DRE" not in df_final.columns:
                pos = df_final.columns.get_loc("Tipo de Pagamento") + 1
                df_final.insert(pos, "Tipo DRE", df_final["__meio_norm__"].map(dre_map))
            else:
                df_final["Tipo DRE"] = df_final["Tipo DRE"].fillna(df_final["__meio_norm__"].map(dre_map))
        
            df_final.drop(columns=["__meio_norm__"], inplace=True, errors="ignore")
        
            # =======================================================
            # 3) Ordenar colunas (deixa Tipo imediatamente após Meio)
            # =======================================================
            cols_preferidas = [
                "Data", "Dia da Semana",
                "Meio de Pagamento", "Tipo de Pagamento", "Tipo DRE",   # <-- acrescentado aqui
                "Loja", "Código Everest", "Grupo", "Código Grupo Everest",
                "Valor (R$)", "Mês", "Ano"
            ]
            # preserva ordem preferida + quaisquer extras que porventura existam
            cols_ordenadas = [c for c in cols_preferidas if c in df_final.columns] + \
                             [c for c in df_final.columns if c not in cols_preferidas and c != "M"]
            df_final = df_final[cols_ordenadas]
        
            # =======================================================
            # 4) Chave M + conversões (igual antes)
            # =======================================================
            df_final['M'] = pd.to_datetime(df_final['Data'], format='%d/%m/%Y').dt.strftime('%Y-%m-%d') + \
                            df_final['Meio de Pagamento'] + df_final['Loja']
        
            # valor para float
            df_final['Valor (R$)'] = df_final['Valor (R$)'].apply(lambda x: float(str(x).replace(',', '.')))
            # data para serial do Sheets
            df_final['Data'] = (pd.to_datetime(df_final['Data'], dayfirst=True) - pd.Timestamp("1899-12-30")).dt.days
            # inteiros (quando houver)
            for col in ["Código Everest", "Código Grupo Everest", "Ano"]:
                if col in df_final.columns:
                    df_final[col] = df_final[col].apply(lambda x: int(x) if pd.notnull(x) and str(x).strip() != "" else "")
        
            # garante M no final
            if "M" in df_final.columns:
                df_final = df_final[[c for c in df_final.columns if c != "M"] + ["M"]]
        
            # =======================================================
            # 5) Planilha destino + duplicidade usando índice dinâmico de 'M'
            # =======================================================
            aba_destino = gc.open("Vendas diarias").worksheet("Faturamento Meio Pagamento")
            valores_existentes = aba_destino.get_all_values()
        
            m_idx = -1
            if valores_existentes:
                header = [h.strip() for h in valores_existentes[0]]
                header_lower = [h.lower() for h in header]
                if "m" in header_lower:
                    m_idx = header_lower.index("m")
        
            # coleção de chaves M existentes
            if len(valores_existentes) > 1:
                if m_idx >= 0:
                    dados_existentes = set(
                        linha[m_idx] for linha in valores_existentes[1:] if len(linha) > m_idx
                    )
                else:
                    # fallback: considera última coluna como M
                    dados_existentes = set(
                        linha[-1] for linha in valores_existentes[1:] if len(linha) > 0
                    )
            else:
                dados_existentes = set()
        
            novos_dados, duplicados = [], []
            for linha in df_final.fillna("").values.tolist():
                chave_m = linha[-1]  # M é a última coluna do df_final
                if chave_m not in dados_existentes:
                    novos_dados.append(linha)
                    dados_existentes.add(chave_m)
                else:
                    duplicados.append(linha)
        
            # =======================================================
            # 6) Envio
            # =======================================================
            lojas_nao_cadastradas = df_final[df_final["Código Everest"].astype(str).isin(["", "nan"])]['Loja'].unique() \
                if "Código Everest" in df_final.columns else []
            todas_lojas_ok = len(lojas_nao_cadastradas) == 0
        
            if todas_lojas_ok and st.button("📥 Enviar dados para o Google Sheets"):
                with st.spinner("🔄 Atualizando..."):
                    if novos_dados:
                        aba_destino.append_rows(novos_dados)
                        st.success(f"✅ {len(novos_dados)} novos registros enviados!")
                    else:
                        st.info("ℹ️ Nenhum novo registro para enviar.")
                    if duplicados:
                        st.warning(f"⚠️ {len(duplicados)} registros duplicados não foram enviados.")

    
    # # ======================
    # # 📝 Aba 3
    # # ======================
    
    # with tab3:
    #     try:
    #         import pandas as pd
    #         pd.set_option('display.max_colwidth', 20)
    #         pd.set_option('display.width', 1000)
    
    #         aba_relatorio = planilha.worksheet("Faturamento Meio Pagamento")
    #         df_relatorio = pd.DataFrame(aba_relatorio.get_all_records())
    #         df_relatorio.columns = df_relatorio.columns.str.strip()
    
    #         aba_meio_pagamento = planilha.worksheet("Tabela Meio Pagamento")
    #         df_meio_pagamento = pd.DataFrame(aba_meio_pagamento.get_all_records())
    #         df_meio_pagamento.columns = df_meio_pagamento.columns.str.strip()
    
    #         # Corrige valores
    #         df_relatorio["Valor (R$)"] = (
    #             df_relatorio["Valor (R$)"]
    #             .astype(str)
    #             .str.replace("R$", "", regex=False)
    #             .str.replace("(", "-")
    #             .str.replace(")", "")
    #             .str.replace(" ", "")
    #             .str.replace(".", "")
    #             .str.replace(",", ".")
    #             .astype(float)
    #         )
    
    #         df_relatorio["Data"] = pd.to_datetime(df_relatorio["Data"], dayfirst=True, errors="coerce")
    #         df_relatorio = df_relatorio[df_relatorio["Data"].notna()]
    
    #         from unidecode import unidecode
    #         for col in ["Loja", "Grupo", "Meio de Pagamento"]:
    #             df_relatorio[col] = df_relatorio[col].astype(str).str.strip().str.upper().map(unidecode)
    #             if col in df_meio_pagamento.columns:
    #                 df_meio_pagamento[col] = df_meio_pagamento[col].astype(str).str.strip().str.upper().map(unidecode)
    
    #         min_data = df_relatorio["Data"].min().date()
    #         max_data = df_relatorio["Data"].max().date()
    
    #         col1, col2, col3 = st.columns(3)
    
    #         with col1:
    #             data_inicio, data_fim = st.date_input(
    #                 "Período:",
    #                 value=(max_data, max_data),
    #                 min_value=min_data,
    #                 max_value=max_data
    #             )
    
    #         with col2:
    #             modo_relatorio = st.selectbox(
    #                 "Tipo de análise:",
    #                 ["Vendas", "Financeiro", "Vendas + Prazo e Taxas"]
    #             )
    
    #         with col3:
    #             if modo_relatorio == "Vendas":
    #                 tipo_relatorio = st.selectbox(
    #                     "Relatório:",
    #                     ["Meio de Pagamento", "Loja", "Grupo"]
    #                 )
    #             else:
    #                 tipo_relatorio = None
    #         if data_inicio > data_fim:
    #             st.warning("🚫 A data inicial não pode ser maior que a data final.")
    #         else:
    #             df_filtrado = df_relatorio[
    #                 (df_relatorio["Data"].dt.date >= data_inicio) &
    #                 (df_relatorio["Data"].dt.date <= data_fim)
    #             ]
    
    #             if df_filtrado.empty:
    #                 st.info("🔍 Não há dados para o período selecionado.")
    #             else:
    #                 if modo_relatorio == "Vendas":
    
    #                     if tipo_relatorio == "Meio de Pagamento":
    #                         index_cols = ["Meio de Pagamento"]
    #                     elif tipo_relatorio == "Loja":
    #                         index_cols = ["Loja", "Grupo", "Meio de Pagamento"]
    #                     elif tipo_relatorio == "Grupo":
    #                         index_cols = ["Grupo", "Meio de Pagamento"]
    
    #                     df_pivot = pd.pivot_table(
    #                         df_filtrado,
    #                         index=index_cols,
    #                         columns=df_filtrado["Data"].dt.strftime("%d/%m/%Y"),
    #                         values="Valor (R$)",
    #                         aggfunc="sum",
    #                         fill_value=0
    #                     ).reset_index()
    
    #                     novo_nome_datas = {col: f"Vendas - {col}" for col in df_pivot.columns if "/" in str(col)}
    #                     df_pivot.rename(columns=novo_nome_datas, inplace=True)
    
    #                     df_pivot["Total Vendas"] = df_pivot[[c for c in df_pivot.columns if "Vendas -" in str(c)]].sum(axis=1)
    
    #                     linha_total_dict = {df_pivot.columns[0]: "TOTAL GERAL"}
    #                     for col in df_pivot.columns[1:]:
    #                         if "Vendas -" in str(col) or col == "Total Vendas":
    #                             linha_total_dict[col] = df_pivot[col].sum()
    #                         else:
    #                             linha_total_dict[col] = np.nan
    #                     linha_total = pd.DataFrame([linha_total_dict])
    
    #                     df_pivot_total = pd.concat([linha_total, df_pivot], ignore_index=True)
    
    #                     df_pivot_exibe = df_pivot_total.copy()
    #                     for col in df_pivot_exibe.select_dtypes(include=[np.number]).columns:
    #                         df_pivot_exibe[col] = df_pivot_exibe[col].map(
    #                             lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    #                             if pd.notna(x) else ""
    #                         )
    
    #                     st.dataframe(df_pivot_exibe, use_container_width=True)
    
    #                 elif modo_relatorio == "Financeiro":
    #                     df_completo = df_filtrado.merge(
    #                         df_meio_pagamento[["Meio de Pagamento", "Prazo", "Antecipa S/N"]],
    #                         on="Meio de Pagamento",
    #                         how="left"
    #                     )
    #                     df_completo["Prazo"] = pd.to_numeric(df_completo["Prazo"], errors="coerce").fillna(0).astype(int)
    #                     df_completo["Antecipa S/N"] = df_completo["Antecipa S/N"].astype(str).str.strip().str.upper()
    
    #                     from pandas.tseries.offsets import BDay
    #                     df_completo["Data Recebimento"] = df_completo.apply(
    #                         lambda row: row["Data"] + BDay(1) if row["Antecipa S/N"] == "SIM" else row["Data"] + BDay(row["Prazo"]),
    #                         axis=1
    #                     )
    
    #                     df_financeiro = df_completo.groupby(df_completo["Data Recebimento"].dt.date)["Valor (R$)"].sum().reset_index()
    #                     df_financeiro = df_financeiro.rename(columns={"Data Recebimento": "Data"}).sort_values("Data")
    
    #                     total_geral = df_financeiro["Valor (R$)"].sum()
    #                     linha_total = pd.DataFrame([["TOTAL GERAL", total_geral]], columns=df_financeiro.columns)
    #                     df_financeiro_total = pd.concat([linha_total, df_financeiro], ignore_index=True)
    
    #                     df_financeiro_total["Valor (R$)"] = df_financeiro_total["Valor (R$)"].map(
    #                         lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    #                         if pd.notna(x) else ""
    #                     )
    
    #                     st.dataframe(df_financeiro_total, use_container_width=True)
    
    #                 elif modo_relatorio == "Vendas + Prazo e Taxas":
    #                     df_completo = df_filtrado.merge(
    #                         df_meio_pagamento[["Meio de Pagamento", "Prazo", "Antecipa S/N", "Taxa Bandeira", "Taxa Antecipação"]],
    #                         on="Meio de Pagamento",
    #                         how="left"
    #                     )
    
    #                     df_pivot = pd.pivot_table(
    #                         df_completo,
    #                         index=["Meio de Pagamento", "Prazo", "Antecipa S/N", "Taxa Bandeira", "Taxa Antecipação"],
    #                         columns=df_completo["Data"].dt.strftime("%d/%m/%Y"),
    #                         values="Valor (R$)",
    #                         aggfunc="sum",
    #                         fill_value=0
    #                     ).reset_index()
    
    #                     colunas_datas = [col for col in df_pivot.columns if "/" in col]
    #                     novo_nome_datas = {col: f"Vendas - {col}" for col in colunas_datas}
    #                     df_pivot.rename(columns=novo_nome_datas, inplace=True)
    #                     df_pivot.rename(columns={"Vendas - Antecipa S/N": "Antecipa S/N"}, inplace=True)
    
    #                     colunas_vendas = [col for col in df_pivot.columns if "Vendas" in col]
    #                     cols_fixas = ["Meio de Pagamento", "Prazo", "Antecipa S/N", "Taxa Bandeira", "Taxa Antecipação"]
    #                     novas_cols = []
    
    #                     for col_vendas in colunas_vendas:
    #                         data_col = col_vendas.split(" - ")[1]
    #                         col_taxa_bandeira = f"Vlr Taxa Bandeira - {data_col}"
    #                         taxa_bandeira = (
    #                             pd.to_numeric(df_pivot["Taxa Bandeira"].astype(str)
    #                                           .str.replace("%","")
    #                                           .str.replace(",","."),
    #                                           errors="coerce").fillna(0) / 100
    #                         )
    #                         df_pivot[col_taxa_bandeira] = df_pivot[col_vendas] * taxa_bandeira
    
    #                         col_taxa_antecipacao = f"Vlr Taxa Antecipação - {data_col}"
    #                         taxa_antecipacao = (
    #                             pd.to_numeric(df_pivot["Taxa Antecipação"].astype(str)
    #                                           .str.replace("%","")
    #                                           .str.replace(",","."),
    #                                           errors="coerce").fillna(0) / 100
    #                         )
    #                         df_pivot[col_taxa_antecipacao] = df_pivot[col_vendas] * taxa_antecipacao
    
    #                         novas_cols.extend([col_vendas, col_taxa_bandeira, col_taxa_antecipacao])
    
    #                     df_pivot = df_pivot[cols_fixas + novas_cols]
    
    #                     df_pivot["Total Vendas"] = df_pivot[colunas_vendas].sum(axis=1)
    #                     df_pivot["Total Tx Bandeira"] = df_pivot[[col for col in df_pivot.columns if "Vlr Taxa Bandeira" in col]].sum(axis=1)
    #                     df_pivot["Total Tx Antecipação"] = df_pivot[[col for col in df_pivot.columns if "Vlr Taxa Antecipação" in col]].sum(axis=1)
    #                     df_pivot["Total a Receber"] = df_pivot["Total Vendas"] - df_pivot["Total Tx Bandeira"] - df_pivot["Total Tx Antecipação"]
    
    #                     linha_total_dict = {col: "" for col in df_pivot.columns}
    #                     linha_total_dict["Meio de Pagamento"] = "TOTAL GERAL"
    #                     for col in df_pivot.columns:
    #                         if "Vendas" in col or "Vlr Taxa Bandeira" in col or "Vlr Taxa Antecipação" in col \
    #                             or "Total Tx" in col or col in ["Total Vendas", "Total a Receber"]:
    #                             linha_total_dict[col] = df_pivot[col].sum()
    
    #                     linha_total = pd.DataFrame([linha_total_dict])
    #                     df_pivot_total = pd.concat([linha_total, df_pivot], ignore_index=True)
    
    #                     df_pivot_exibe = df_pivot_total.copy()
    #                     for col in [c for c in df_pivot_exibe.columns if "Vendas" in c or "Vlr Taxa Bandeira" in c 
    #                                 or "Vlr Taxa Antecipação" in c or "Total Tx" in c or c in ["Total Vendas", "Total a Receber"]]:
    #                         df_pivot_exibe[col] = df_pivot_exibe[col].map(
    #                             lambda x: f"R$ {x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    #                         )
    
    #                     st.dataframe(df_pivot_exibe, use_container_width=True)
    
    #                     from openpyxl import load_workbook
    
    #                     output = BytesIO()
    #                     df_exportar = df_pivot_total.copy()
    #                     df_exportar["Taxa Bandeira"] = (
    #                         pd.to_numeric(df_exportar["Taxa Bandeira"].astype(str)
    #                                       .str.replace("%", "")
    #                                       .str.replace(",", "."),
    #                                       errors="coerce") / 100
    #                     )
    #                     df_exportar["Taxa Antecipação"] = (
    #                         pd.to_numeric(df_exportar["Taxa Antecipação"].astype(str)
    #                                       .str.replace("%", "")
    #                                       .str.replace(",", "."),
    #                                       errors="coerce") / 100
    #                     )
    
    #                     with pd.ExcelWriter(output, engine='openpyxl') as writer:
    #                         df_exportar.to_excel(writer, index=False, sheet_name="PrazoTaxas")
    #                     output.seek(0)
    
    #                     wb = load_workbook(output)
    #                     ws = wb["PrazoTaxas"]
    #                     header = [cell.value for cell in ws[1]]
    
    #                     for col_name in ["Taxa Bandeira", "Taxa Antecipação"]:
    #                         if col_name in header:
    #                             col_idx = header.index(col_name) + 1
    #                             for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
    #                                 for cell in row:
    #                                     cell.number_format = "0.00%"
    
    #                     for col_name in header:
    #                         if ("Vendas" in col_name or "Vlr Taxa Bandeira" in col_name 
    #                             or "Vlr Taxa Antecipação" in col_name or "Total Tx" in col_name 
    #                             or col_name in ["Total Vendas", "Total a Receber"]):
    #                             col_idx = header.index(col_name) + 1
    #                             for row in ws.iter_rows(min_row=2, min_col=col_idx, max_col=col_idx):
    #                                 for cell in row:
    #                                     cell.number_format = '"R$" #,##0.00'
    
    #                     output_final = BytesIO()
    #                     wb.save(output_final)
    #                     output_final.seek(0)
    
    #                     st.download_button(
    #                         "📥 Baixar Excel",
    #                         data=output_final,
    #                         file_name=f"Vendas_Prazo_Taxas_{data_inicio.strftime('%d-%m-%Y')}_a_{data_fim.strftime('%d-%m-%Y')}.xlsx",
    #                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    #                     )
    
     #    except Exception as e:
     #        st.error(f"❌ Erro ao acessar Google Sheets: {e}")
    
    
     #   except Exception as e:
     #       st.error(f"❌ Erro ao acessar Google Sheets: {e}")
