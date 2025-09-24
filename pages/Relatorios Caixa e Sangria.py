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


#st.set_page_config(page_title="Painel Agrupado", layout="wide")
#st.set_page_config(page_title="Vendas Diarias", layout="wide")
# 🔒 Bloqueia o acesso caso o usuário não esteja logado
if not st.session_state.get("acesso_liberado"):
    st.stop()
import streamlit as st

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
# Spinner durante todo o processamento
# ======================
with st.spinner("⏳ Processando..."):
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
            <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatórios</h1>
        </div>
    """, unsafe_allow_html=True)
    
    # ================================
    # 3. Separação em ABAS
    # ================================
    aba5 = st.tabs([
      
        "🧾 Relatórios Caixa e Sangria"
    ])


    # ================================
  
    # ================================
    # 🧾 Relatórios Caixa e Sangria (com sub-abas)
    # ================================
    
    import re
    import numpy as np
    from io import BytesIO
    import pandas as pd
    import streamlit as st
    
    # ---------- helpers (iguais aos que você validou) ----------
    def parse_valor_brl_sheets(x):
        """
        Normaliza valores vindo do Sheets:
          • Com vírgula: força 2 casas.
          • Sem vírgula:
              - 1 a 3 dígitos -> reais inteiros      (548   -> 548,00)
              - 4 dígitos:
                  * termina com '00' -> ÷100         (1200  -> 12,00)
                  * termina com '0'  -> ÷10          (5480  -> 548,00)
                  * outros           -> mantém       (1234  -> 1.234,00)
              - 5+ dígitos -> ÷100                    (54800 -> 548,00)
        Aceita negativos '(...)' ou '-'. Remove 'R$', espaços e pontos de milhar.
        Retorna float.
        """
        if isinstance(x, (int, float)):
            try:
                return float(x)
            except Exception:
                return 0.0
    
        s = str(x).strip()
        if s == "" or s.lower() in {"nan", "none"}:
            return 0.0
    
        neg = False
        if s.startswith("(") and s.endswith(")"):
            neg = True
            s = s[1:-1].strip()
        if s.startswith("-"):
            neg = True
            s = s[1:].strip()
    
        s = (s.replace("R$", "")
             .replace("\u00A0", "")
             .replace(" ", "")
             .replace(".", ""))
    
        if "," in s:
            inteiro, dec = s.rsplit(",", 1)
            inteiro = re.sub(r"\D", "", inteiro)
            dec     = re.sub(r"\D", "", dec)
            if dec == "":
                dec = "00"
            elif len(dec) == 1:
                dec = dec + "0"
            else:
                dec = dec[:2]
            num_str = f"{inteiro}.{dec}" if inteiro != "" else f"0.{dec}"
            try:
                val = float(num_str)
            except Exception:
                val = 0.0
        else:
            digits = re.sub(r"\D", "", s)
            if digits == "":
                val = 0.0
            else:
                n = len(digits)
                if n <= 3:
                    val = float(digits)
                elif n == 4:
                    if digits.endswith("00"):
                        val = float(digits) / 100.0
                    elif digits.endswith("0"):
                        val = float(digits) / 10.0
                    else:
                        val = float(digits)
                else:  # n >= 5
                    val = float(digits) / 100.0
    
        return -val if neg else val
    
    
    def _render_df(df, *, height=480):
        df = df.copy().reset_index(drop=True)
        seen, new_cols = {}, []
        for c in df.columns:
            s = "" if c is None else str(c)
            if s in seen:
                seen[s] += 1
                s = f"{s}_{seen[s]}"
            else:
                seen[s] = 0
            new_cols.append(s)
        df.columns = new_cols
        st.dataframe(df, use_container_width=True, height=height, hide_index=True)
        return df
    
    
    def pick_valor_col(cols):
        def norm(s):
            return re.sub(r"[\s\u00A0]+", " ", str(s)).strip().lower()
        nm = {c: norm(c) for c in cols}
    
        prefer = ["valor(r$)", "valor (r$)", "valor", "valor r$"]
        for want in prefer:
            for c, n in nm.items():
                if n == want:
                    return c
    
        for c, n in nm.items():
            if ("valor" in n
                and "valores" not in n
                and "google"  not in n
                and "sheet"   not in n):
                return c
        return None
    
    
    # ================================
    # ABA 5
    # ================================
    with aba5:
        # --- carrega aba Sangria ---
        df_sangria = None
        try:
            ws_sangria = planilha_empresa.worksheet("Sangria")
            df_sangria = pd.DataFrame(ws_sangria.get_all_records())
            df_sangria.columns = [c.strip() for c in df_sangria.columns]
            if "Data" in df_sangria.columns:
                df_sangria["Data"] = pd.to_datetime(df_sangria["Data"], dayfirst=True, errors="coerce")
        except Exception as e:
            st.warning(f"⚠️ Não foi possível carregar a aba 'Sangria': {e}")
    
        sub_sangria, sub_caixa, sub_evx = st.tabs(["💸 Movimentação de Caixa", "🧰 Controle de Sangria", "🗂️ Everest x Sangria"])

        # -------------------------------
        # Sub-aba: 💸  Movimentação de Caixa
        # -------------------------------
        with sub_sangria:
            if df_sangria is None or df_sangria.empty:
                st.info("Sem dados de **sangria** disponíveis.")
            else:
                from io import BytesIO
        
                # Base e colunas
                df_sangria = df_sangria.copy()
                df_sangria.columns = [str(c).strip() for c in df_sangria.columns]
        
                # Data obrigatória e normalizada
                if "Data" not in df_sangria.columns:
                    st.error("A aba 'Sangria' precisa da coluna **Data**.")
                    st.stop()
                df_sangria["Data"] = pd.to_datetime(df_sangria["Data"], errors="coerce", dayfirst=True)
        
                # Coluna de valor e parsing BRL -> float
                col_valor = pick_valor_col(df_sangria.columns)
                if not col_valor:
                    st.error("Não encontrei a coluna de **valor** (ex.: 'Valor(R$)').")
                    st.stop()
                df_sangria[col_valor] = df_sangria[col_valor].map(parse_valor_brl_sheets).astype(float)
        
                # ------------- Filtros robustos -------------
                top1, top2, top3, top4 = st.columns([1.2, 1.2, 1.6, 1.6])
                with top1:
                    dmin = pd.to_datetime(df_sangria["Data"].min(), errors="coerce")
                    dmax = pd.to_datetime(df_sangria["Data"].max(), errors="coerce")
                    today = pd.Timestamp.today().normalize()
                    if pd.isna(dmin): dmin = today
                    if pd.isna(dmax): dmax = today
                    dt_inicio, dt_fim = st.date_input(
                        "Período",
                        value=(dmax.date(), dmax.date()),
                        min_value=dmin.date(),
                        max_value=(dmax.date() if dmax >= dmin else dmin.date()),
                        key="periodo_sangria"
                    )
                with top2:
                    lojas = sorted(df_sangria.get("Loja", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
                    lojas_sel = st.multiselect("Lojas", options=lojas, default=[], key="lojas_sangria")
                with top3:
                    descrs = sorted(df_sangria.get("Descrição Agrupada", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
                    descrs_sel = st.multiselect("Descrição Agrupada", options=descrs, default=[], key="descr_sangria")
                with top4:
                    visao = st.selectbox(
                        "Visão do Relatório",
                        options=["Analítico", "Sintético"],  # mantemos só essas duas aqui
                        index=0,
                        key="visao_sangria"
                    )
        
                # Aplica filtros
                df_fil = df_sangria.copy()
                df_fil = df_fil[(df_fil["Data"].dt.date >= dt_inicio) & (df_fil["Data"].dt.date <= dt_fim)]
                if lojas_sel:
                    df_fil = df_fil[df_fil["Loja"].astype(str).isin(lojas_sel)]
                if descrs_sel:
                    df_fil = df_fil[df_fil["Descrição Agrupada"].astype(str).isin(descrs_sel)]
        
                # Helper de formatação apenas visual
                def _fmt_brl_df(df, col):
                    df[col] = df[col].apply(
                        lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        if isinstance(v, (int, float)) else v
                    )
                    return df
        
                # ===================== Visões =====================
                df_exibe = pd.DataFrame()
        
                if visao == "Analítico":
                    grid = st.empty()
        
                    df_base = df_fil.copy()
                    df_base["Data"] = pd.to_datetime(df_base["Data"], errors="coerce").dt.normalize()
                    df_base = df_base.sort_values(["Data"], na_position="last")
        
                    total_val = df_base[col_valor].sum(min_count=1)
                    total_row = {c: "" for c in df_base.columns}
                    if "Loja" in total_row: total_row["Loja"] = "TOTAL"
                    if "Data" in total_row: total_row["Data"] = pd.NaT
                    if "Descrição Agrupada" in total_row: total_row["Descrição Agrupada"] = ""
                    total_row[col_valor] = total_val
        
                    df_exibe = pd.concat([pd.DataFrame([total_row]), df_base], ignore_index=True)
        
                    # Datas p/ exibição (TOTAL vazio)
                    df_exibe["Data"] = pd.to_datetime(df_exibe["Data"], errors="coerce").dt.strftime("%d/%m/%Y")
                    df_exibe.loc[df_exibe.index == 0, "Data"] = ""
        
                    # Valor p/ exibição
                    df_exibe = _fmt_brl_df(df_exibe, col_valor)
        
                    # Remove colunas “técnicas”/ruído
                    aliases_remover = [
                        "Código Everest", "Codigo Everest", "Cod Everest",
                        "Código grupo Everest", "Codigo grupo Everest", "Cod Grupo Everest", "Código Grupo Everest",
                        "Mês", "Mes", "Ano", "Duplicidade", "Possível Duplicidade", "Duplicado", "Sistema"
                    ]
                    df_exibe = df_exibe.drop(columns=[c for c in aliases_remover if c in df_exibe.columns], errors="ignore")
        
                    grid.dataframe(df_exibe, use_container_width=True, hide_index=True)
        
                    # Export Excel mantendo tipos
                    df_export = pd.concat([pd.DataFrame([total_row]), df_base], ignore_index=True)
                    df_export = df_export.drop(columns=[c for c in aliases_remover if c in df_export.columns], errors="ignore")
                    df_export["Data"] = pd.to_datetime(df_export["Data"], errors="coerce")
                    df_export[col_valor] = pd.to_numeric(df_export[col_valor], errors="coerce")
        
                    buf = BytesIO()
                    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                        sheet_name = "Analítico"
                        df_export.to_excel(writer, sheet_name=sheet_name, index=False)
                        wb, ws = writer.book, writer.sheets[sheet_name]
                        header_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#F2F2F2", "border": 1})
                        date_fmt   = wb.add_format({"num_format": "dd/mm/yyyy", "border": 1})
                        money_fmt  = wb.add_format({"num_format": "R$ #,##0.00", "border": 1})
                        text_fmt   = wb.add_format({"border": 1})
                        total_row_fmt   = wb.add_format({"bold": True, "bg_color": "#FCE5CD", "border": 1})
                        total_money_fmt = wb.add_format({"bold": True, "bg_color": "#FCE5CD", "border": 1, "num_format": "R$ #,##0.00"})
        
                        # Cabeçalho
                        for j, name in enumerate(df_export.columns):
                            ws.write(0, j, name, header_fmt)
        
                        # Larguras e formatos
                        for j, name in enumerate(df_export.columns):
                            largura, fmt = 18, text_fmt
                            if name.lower() == "data":
                                largura, fmt = 12, date_fmt
                            elif name == col_valor:
                                largura, fmt = 16, money_fmt
                            elif "loja" in name.lower():
                                largura = 28
                            elif "grupo" in name.lower():
                                largura = 22
                            ws.set_column(j, j, largura, fmt)
        
                        # Destaca TOTAL (linha 1 de dados no Excel)
                        ws.set_row(1, None, total_row_fmt)
                        if pd.notna(df_export.iloc[0][col_valor]):
                            ws.write_number(1, list(df_export.columns).index(col_valor), float(df_export.iloc[0][col_valor]), total_money_fmt)
                        if "Loja" in df_export.columns:
                            ws.write_string(1, list(df_export.columns).index("Loja"), "TOTAL", total_row_fmt)
        
                        ws.freeze_panes(1, 0)
        
                    st.download_button(
                        "⬇️ Baixar Excel (Analítico)",
                        data=buf.getvalue(),
                        file_name="Relatorio_Analitico.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
        
                elif visao == "Sintético":
                    if "Loja" not in df_fil.columns:
                        st.warning("Para 'Sintético', preciso da coluna **Loja**.")
                    else:
                        tmp = df_fil.copy()
                        tmp["Data"] = pd.to_datetime(tmp["Data"], errors="coerce").dt.normalize()
        
                        # Encontra/garante 'Grupo'
                        col_grupo = None
                        for c in tmp.columns:
                            if str(c).strip().lower() == "grupo":
                                col_grupo = c; break
                        if not col_grupo:
                            col_grupo = next((c for c in tmp.columns if "grupo" in str(c).lower() and "everest" not in str(c).lower()), None)
                        if not col_grupo and "Loja" in tmp.columns:
                            mapa = df_empresa[["Loja", "Grupo"]].drop_duplicates()
                            tmp = tmp.merge(mapa, on="Loja", how="left")
                            col_grupo = "Grupo"
        
                        group_cols = [c for c in [col_grupo, "Loja", "Data"] if c]
                        df_agg = tmp.groupby(group_cols, as_index=False)[col_valor].sum()
        
                        ren = {col_valor: "Sangria"}
                        if col_grupo and col_grupo != "Grupo": ren[col_grupo] = "Grupo"
                        df_agg = df_agg.rename(columns=ren)
        
                        df_agg = df_agg.sort_values(["Data", "Grupo", "Loja"], na_position="last")
        
                        total_sangria = df_agg["Sangria"].sum(min_count=1)
                        linha_total = pd.DataFrame({"Grupo":["TOTAL"], "Loja":[""], "Data":[pd.NaT], "Sangria":[total_sangria]})
                        df_exibe = pd.concat([linha_total, df_agg], ignore_index=True)
        
                        # Exibição bonita
                        df_show = df_exibe.copy()
                        df_show["Data"] = pd.to_datetime(df_show["Data"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
                        df_show["Sangria"] = df_show["Sangria"].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
        
                        st.dataframe(df_show[["Grupo","Loja","Data","Sangria"]], use_container_width=True, hide_index=True)
        
                        # Excel com tipos corretos
                        df_export = df_exibe[["Grupo","Loja","Data","Sangria"]].copy()
                        df_export["Data"] = pd.to_datetime(df_export["Data"], errors="coerce")
                        df_export["Sangria"] = pd.to_numeric(df_export["Sangria"], errors="coerce")
        
                        buf = BytesIO()
                        with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                            sheet_name = "Sintético"
                            df_export.to_excel(writer, sheet_name=sheet_name, index=False)
                            wb, ws = writer.book, writer.sheets[sheet_name]
                            header_fmt = wb.add_format({"bold": True, "align": "center", "valign": "vcenter", "bg_color": "#F2F2F2", "border": 1})
                            date_fmt   = wb.add_format({"num_format": "dd/mm/yyyy", "border": 1})
                            money_fmt  = wb.add_format({"num_format": "R$ #,##0.00", "border": 1})
                            text_fmt   = wb.add_format({"border": 1})
                            total_row_fmt   = wb.add_format({"bold": True, "bg_color": "#FCE5CD", "border": 1})
                            total_money_fmt = wb.add_format({"bold": True, "bg_color": "#FCE5CD", "border": 1, "num_format": "R$ #,##0.00"})
        
                            # Cabeçalho
                            for j, name in enumerate(["Grupo","Loja","Data","Sangria"]):
                                ws.write(0, j, name, header_fmt)
        
                            ws.set_column("A:A", 20, text_fmt)
                            ws.set_column("B:B", 28, text_fmt)
                            ws.set_column("C:C", 12, date_fmt)
                            ws.set_column("D:D", 14, money_fmt)
        
                            # TOTAL destacado (linha 1 no Excel)
                            ws.set_row(1, None, total_row_fmt)
                            if pd.notna(df_export.iloc[0]["Sangria"]):
                                ws.write_number(1, 3, float(df_export.iloc[0]["Sangria"]), total_money_fmt)
                            ws.write_string(1, 0, "TOTAL", total_row_fmt)
        
                            ws.freeze_panes(1, 0)
        
                        st.download_button(
                            "⬇️ Baixar Excel (Sintético)",
                            data=buf.getvalue(),
                            file_name="Relatorio_Sintetico.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
        
                # Oculta colunas técnicas se montamos df_exibe (apenas Analítico usa esse bloco comum)
                if not df_exibe.empty and visao == "Analítico":
                    ocultar = ["Código Grupo Everest","Duplicidade","Sistema","Mês","Mes","Ano"]
                    df_show = df_exibe.drop(columns=ocultar, errors="ignore")
                    # (já exibimos acima no grid; export feito também)

        # -------------------------------
        # Sub-aba: 🧰 CONTROLE DE CAIXA
        # -------------------------------
        with sub_caixa:
            if df_sangria is None or df_sangria.empty:
                st.info("Sem dados de **sangria** disponíveis.")
            else:
                from io import BytesIO
        
                # ===== base =====
                df = df_sangria.copy()
                df.columns = [str(c).strip() for c in df.columns]
        
                if "Data" not in df.columns:
                    st.error("A aba 'Sangria' precisa da coluna **Data**.")
                    st.stop()
                df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)
        
                col_valor = pick_valor_col(df.columns)
                if not col_valor:
                    st.error("Não encontrei a coluna de **valor** (ex.: 'Valor(R$)').")
                    st.stop()
                df[col_valor] = df[col_valor].map(parse_valor_brl_sheets).astype(float)
        
                # ===== filtros =====
                c1, c2, c3, c4 = st.columns([1.2, 1.2, 1.6, 1.6])
                with c1:
                    dmin = pd.to_datetime(df["Data"].min(), errors="coerce")
                    dmax = pd.to_datetime(df["Data"].max(), errors="coerce")
                    today = pd.Timestamp.today().normalize()
                    if pd.isna(dmin): dmin = today
                    if pd.isna(dmax): dmax = today
                    dt_inicio, dt_fim = st.date_input(
                        "Período",
                        value=(dmax.date(), dmax.date()),
                        min_value=dmin.date(),
                        max_value=(dmax.date() if dmax >= dmin else dmin.date()),
                        key="caixa_periodo",
                    )
                with c2:
                    lojas = sorted(df.get("Loja", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
                    lojas_sel = st.multiselect("Lojas", options=lojas, default=[], key="caixa_lojas")
                with c3:
                    descrs = sorted(df.get("Descrição Agrupada", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
                    descrs_sel = st.multiselect("Descrição Agrupada", options=descrs, default=[], key="caixa_descr")
                with c4:
                    visao = st.selectbox(
                        "Visão do Relatório",
                        options=["Comparativa Everest", "Diferenças Everest"],
                        index=0,
                        key="caixa_visao",
                    )
        
                # aplica filtros
                df_fil = df[(df["Data"].dt.date >= dt_inicio) & (df["Data"].dt.date <= dt_fim)].copy()
                if lojas_sel:
                    df_fil = df_fil[df_fil["Loja"].astype(str).isin(lojas_sel)]
                if descrs_sel:
                    df_fil = df_fil[df_fil["Descrição Agrupada"].astype(str).isin(descrs_sel)]
        
                # helpers
                def _fmt_brl_df(_df, col):
                    _df[col] = _df[col].apply(
                        lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                        if isinstance(v, (int, float)) else v
                    )
                    return _df
        
                df_exibe = pd.DataFrame()
        
                
                # ======= Comparativa / Diferenças (mantidas) =======
                if visao in ("Comparativa Everest", "Diferenças Everest"):
                    base = df_fil.copy()
                    if "Data" not in base.columns or "Código Everest" not in base.columns or not col_valor:
                        st.error("❌ Preciso de 'Data', 'Código Everest' e coluna de valor na aba Sangria.")
                        df_exibe = pd.DataFrame()
                    else:
                        base["Data"] = pd.to_datetime(base["Data"], dayfirst=True, errors="coerce").dt.normalize()
                        base[col_valor] = pd.to_numeric(base[col_valor], errors="coerce").fillna(0.0)
                        base["Código Everest"] = base["Código Everest"].astype(str).str.extract(r"(\d+)")
        
                        df_sys = (
                            base.groupby(["Código Everest","Data"], as_index=False)[col_valor]
                                .sum()
                                .rename(columns={col_valor:"Sangria (Sistema)"})
                        )
        
                        # Everest
                        ws_ev = planilha_empresa.worksheet("Sangria Everest")
                        df_ev = pd.DataFrame(ws_ev.get_all_records())
                        df_ev.columns = [c.strip() for c in df_ev.columns]
        
                        def _norm(s): return re.sub(r"[^a-z0-9]", "", str(s).lower())
                        cmap = {_norm(c): c for c in df_ev.columns}
                        col_emp   = cmap.get("empresa")
                        col_dt_ev = next((orig for norm, orig in cmap.items()
                                          if norm in ("dlancamento","dlancament","dlanamento","datadelancamento","data")), None)
                        col_val_ev= next((orig for norm, orig in cmap.items()
                                          if norm in ("valorlancamento","valorlancament","valorlcto","valor")), None)
                        col_fant  = next((orig for norm, orig in cmap.items()
                                          if norm in ("fantasiaempresa","fantasia")), None)
        
                        if not all([col_emp, col_dt_ev, col_val_ev]):
                            st.error("❌ Na 'Sangria Everest' preciso de 'Empresa', 'D. Lançamento' e 'Valor Lancamento'.")
                            df_exibe = pd.DataFrame()
                        else:
                            de = df_ev.copy()
                            de["Código Everest"]   = de[col_emp].astype(str).str.extract(r"(\d+)")
                            de["Fantasia Everest"] = de[col_fant] if col_fant else ""
                            de["Data"]             = pd.to_datetime(de[col_dt_ev], dayfirst=True, errors="coerce").dt.normalize()
                            de["Valor Lancamento"] = de[col_val_ev].map(parse_valor_brl_sheets).astype(float)
                            # restringe por período do filtro
                            de = de[(de["Data"].dt.date >= dt_inicio) & (de["Data"].dt.date <= dt_fim)]
                            de["Sangria Everest"]  = de["Valor Lancamento"].abs()
        
                            def _pick_first(s):
                                s = s.dropna().astype(str).str.strip()
                                s = s[s != ""]
                                return s.iloc[0] if not s.empty else ""
                            de_agg = (
                                de.groupby(["Código Everest","Data"], as_index=False)
                                  .agg({"Sangria Everest":"sum","Fantasia Everest": _pick_first})
                            )
        
                            cmp = df_sys.merge(de_agg, on=["Código Everest","Data"], how="outer", indicator=True)
                            cmp["Sangria (Sistema)"] = cmp["Sangria (Sistema)"].fillna(0.0)
                            cmp["Sangria Everest"]   = cmp["Sangria Everest"].fillna(0.0)
        
                            # mapeia Loja/Grupo via Tabela Empresa
                            mapa = df_empresa.copy()
                            mapa.columns = [str(c).strip() for c in mapa.columns]
                            if "Código Everest" in mapa.columns:
                                mapa["Código Everest"] = mapa["Código Everest"].astype(str).str.extract(r"(\d+)")
                                cmp = cmp.merge(mapa[["Código Everest","Loja","Grupo"]].drop_duplicates(),
                                                on="Código Everest", how="left")
        
                            # fallback LOJA = Fantasia (linhas só do Everest)
                            cmp["Loja"] = cmp["Loja"].astype(str)
                            so_everest = (cmp["_merge"] == "right_only") & (cmp["Loja"].isin(["", "nan"]))
                            cmp.loc[so_everest, "Loja"] = cmp.loc[so_everest, "Fantasia Everest"]
                            cmp["Nao Mapeada?"] = so_everest
        
                            cmp["Diferença"] = cmp["Sangria (Sistema)"] - cmp["Sangria Everest"]
                            if visao == "Diferenças Everest":
                                cmp = cmp[np.isclose(cmp["Diferença"], 0.0) == False]
        
                            cmp = cmp[["Grupo","Loja","Código Everest","Data",
                                       "Sangria (Sistema)","Sangria Everest","Diferença","Nao Mapeada?"]
                                     ].sort_values(["Grupo","Loja","Código Everest","Data"])
        
                            total = {
                                "Grupo":"TOTAL","Loja":"","Código Everest":"","Data":pd.NaT,
                                "Sangria (Sistema)": cmp["Sangria (Sistema)"].sum(),
                                "Sangria Everest":   cmp["Sangria Everest"].sum(),
                                "Diferença":         cmp["Diferença"].sum(),
                                "Nao Mapeada?": False
                            }
                            df_exibe = pd.concat([pd.DataFrame([total]), cmp], ignore_index=True)
        
                            df_exibe["Data"] = pd.to_datetime(df_exibe["Data"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
                            for c in ["Sangria (Sistema)","Sangria Everest","Diferença"]:
                                df_exibe[c] = df_exibe[c].apply(
                                    lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
                                    if isinstance(v,(int,float)) else v
                                )
                            st.session_state.__cmp_has_red = True
        
                # ===== render comum + export =====
                if not df_exibe.empty:
                    if visao in ("Comparativa Everest", "Diferenças Everest"):
                        ocultar = []  # mantém Código Everest
                    else:
                        ocultar = ["Código Grupo Everest","Duplicidade","Sistema","Mês","Mes","Ano"]
        
                    df_show = df_exibe.drop(columns=ocultar, errors="ignore").copy()
        
                    if st.session_state.get("__cmp_has_red") and "Nao Mapeada?" in df_show.columns and "Loja" in df_show.columns:
                        def _paint_row(row):
                            styles = [""] * len(df_show.columns)
                            if bool(row.get("Nao Mapeada?", False)):
                                styles[df_show.columns.get_loc("Loja")] = "color: red; font-weight: 700"
                            return styles
                        st.dataframe(df_show.style.apply(_paint_row, axis=1),
                                     use_container_width=True, height=480)
                    else:
                        _render_df(df_show, height=480)
        
                    # export genérico da visão exibida
                    df_exp = df_show.drop(columns=["Nao Mapeada?"], errors="ignore")
                    bufx = BytesIO()
                    with pd.ExcelWriter(bufx, engine="openpyxl") as w:
                        df_exp.to_excel(w, index=False, sheet_name="ControleCaixa")
                    bufx.seek(0)
                    st.download_button(f"⬇️ Baixar Excel ({visao})", bufx, f"ControleCaixa_{visao.replace(' ','_')}.xlsx")

    
        # -------------------------------
        # Sub-aba: 🗂️ EVEREST x SANGRIA
        # -------------------------------
        with sub_evx:
            if df_sangria is None or df_sangria.empty:
                st.info("Sem dados para comparação.")
            else:
                col_valor_evx = pick_valor_col(df_sangria.columns)
                if ("Código Everest" in df_sangria.columns) and col_valor_evx:
                    df_top = (
                        df_sangria.groupby(["Código Everest", "Loja"], dropna=False)[col_valor_evx]
                                  .sum()
                                  .reset_index()
                                  .sort_values(col_valor_evx, ascending=False)
                                  .head(50)
                    )
                    df_top[col_valor_evx] = df_top[col_valor_evx].apply(
                        lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                    )
                    st.markdown("#### Top 50 — Sangria por Código Everest / Loja")
                    _render_df(df_top, height=480)
    
                    buf2 = BytesIO()
                    with pd.ExcelWriter(buf2, engine="openpyxl") as w:
                        df_top.to_excel(w, index=False, sheet_name="Everest_x_Sangria")
                    buf2.seek(0)
                    st.download_button("⬇️ Baixar Excel (Everest x Sangria)", buf2, "everest_x_sangria.xlsx")
                else:
                    st.info("Para esta comparação, é preciso ter **Código Everest** e a coluna de valor (ex.: 'Valor(R$)').")
