# pages/05_Relatorios_Caixa_Sangria.py
import streamlit as st


import pandas as pd
import numpy as np
import re, json
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Bloqueio opcional (mantenha se você já usa login/sessão)
if not st.session_state.get("acesso_liberado"):
    st.stop()

# ======================
# CSS (opcional)
# ======================
st.markdown("""
<style>
[data-testid="stToolbar"]{visibility:hidden;height:0;position:fixed}
</style>
""", unsafe_allow_html=True)

with st.spinner("⏳ Carregando dados..."):
    # ============ Conexão com Google Sheets ============
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    credentials_dict = json.loads(st.secrets["GOOGLE_SERVICE_ACCOUNT"])
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(credentials_dict, scope)
    gc = gspread.authorize(credentials)
    planilha_empresa = gc.open("Vendas diarias")

    # Tabela Empresa (para mapear Grupo/Loja/Tipo/PDV etc.)
    df_empresa = pd.DataFrame(planilha_empresa.worksheet("Tabela Empresa").get_all_records())
    df_empresa.columns = [str(c).strip() for c in df_empresa.columns]
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
            <h1 style='display: inline; margin: 0; font-size: 2.4rem;'>Relatórios Caixa Sangria</h1>
        </div>
    """, unsafe_allow_html=True)
    # ============ Helpers ============



    import unicodedata
    import re
    
    def _norm_txt(s: str) -> str:
        s = str(s or "").strip().lower()
        s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("utf-8")
        return s
    
    def eh_deposito_mask(df, cols_texto=None):
        """
        Retorna uma Series booleana marcando linhas de depósito.
        Critérios por texto em colunas comuns (ajuste a lista se quiser).
        """
        if cols_texto is None:
            cols_texto = [
                "Descrição Agrupada", "Descrição", "Historico", "Histórico",
                "Categoria", "Obs", "Observação", "Tipo", "Tipo Movimento"
            ]
        cols_texto = [c for c in cols_texto if c in df.columns]
        if not cols_texto:
            return pd.Series(False, index=df.index)
    
        txt = df[cols_texto].astype(str).agg(" ".join, axis=1).map(_norm_txt)
    
        padrao = r"""
            \bdeposito\b        |   # 'deposito'/'depósito'
            \bdepsito\b         |   # variações sem acento
            \bdep\b             |   # abreviação comum
            credito\s+em\s+conta|
            transf(erencia)?\s*(p/?\s*banco|banco) |
            envio\s*para\s*banco|
            remessa\s* banco
        """
        return txt.str.contains(padrao, flags=re.IGNORECASE | re.VERBOSE, regex=True, na=False)
    
    def brl(v):
        try:
            return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        except Exception:
            return "R$ 0,00"

    def parse_valor_brl_sheets(x):
        """
        Normaliza valores vindos do Sheets para float (BRL):
        Aceita negativos '(...)' ou '-'. Remove 'R$', espaços e pontos de milhar.
        Regras para quando não há vírgula: heurísticas de casas decimais.
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

    # ============ Carrega aba Sangria ============
    df_sangria = None
    try:
        ws_sangria = planilha_empresa.worksheet("Sangria")
        df_sangria = pd.DataFrame(ws_sangria.get_all_records())
        df_sangria.columns = [c.strip() for c in df_sangria.columns]
        if "Data" in df_sangria.columns:
            df_sangria["Data"] = pd.to_datetime(df_sangria["Data"], dayfirst=True, errors="coerce")
    except Exception as e:
        st.warning(f"⚠️ Não foi possível carregar a aba 'Sangria': {e}")

# ============ Cabeçalho ============
#st.markdown("""
#<div style='display:flex;align-items:center;gap:10px;margin-bottom:12px;'>
#  <img src='https://img.icons8.com/color/48/cash-register.png' width='36'/>
#  <h1 style='margin:0;font-size:1.8rem;'>Relatórios Caixa & Sangria</h1>
#</div>
#""", unsafe_allow_html=True)

# ============ Sub-Abas ============
sub_sangria, sub_caixa = st.tabs([
    "💸 Movimentação de Caixa",   # Analítico/Sintético da Sangria
    "🧰 Controle de Sangria"      # Comparativa Everest / Diferenças
   
])

# -------------------------------
# Sub-aba: 💸  Movimentação de Caixa (Analítico / Sintético)
# -------------------------------
with sub_sangria:
    if df_sangria is None or df_sangria.empty:
        st.info("Sem dados de **sangria** disponíveis.")
    else:
        from io import BytesIO

        # Base e colunas
        df = df_sangria.copy()
        df.columns = [str(c).strip() for c in df.columns]

        # Data obrigatória e normalizada
        if "Data" not in df.columns:
            st.error("A aba 'Sangria' precisa da coluna **Data**.")
            st.stop()
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)

        # Coluna de valor e parsing BRL -> float
        col_valor = pick_valor_col(df.columns)
        if not col_valor:
            st.error("Não encontrei a coluna de **valor** (ex.: 'Valor(R$)').")
            st.stop()
        df[col_valor] = df[col_valor].map(parse_valor_brl_sheets).astype(float)

        # Filtros
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
                key="periodo_sangria_movi"
            )
        with c2:
            lojas = sorted(df.get("Loja", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
            lojas_sel = st.multiselect("Lojas", options=lojas, default=[], key="lojas_sangria_movi")
        with c3:
            descrs = sorted(df.get("Descrição Agrupada", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
            descrs_sel = st.multiselect("Descrição Agrupada", options=descrs, default=[], key="descr_sangria_movi")
        with c4:
            visao = st.selectbox(
                "Visão do Relatório",
                options=["Analítico", "Sintético"],
                index=0,
                key="visao_sangria_movi"
            )

        # Aplica filtros
        df_fil = df[(df["Data"].dt.date >= dt_inicio) & (df["Data"].dt.date <= dt_fim)].copy()
        if lojas_sel:
            df_fil = df_fil[df_fil["Loja"].astype(str).isin(lojas_sel)]
        if descrs_sel:
            df_fil = df_fil[df_fil["Descrição Agrupada"].astype(str).isin(descrs_sel)]

        # Helper de formatação BRL (apenas visual)
        def _fmt_brl_df(_df, col):
            _df[col] = _df[col].apply(
                lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
                if isinstance(v, (int, float)) else v
            )
            return _df

        df_exibe = pd.DataFrame()

        # ====== Analítico ======
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

            # Datas (TOTAL vazio) e valor formatado
            df_exibe["Data"] = pd.to_datetime(df_exibe["Data"], errors="coerce").dt.strftime("%d/%m/%Y")
            df_exibe.loc[df_exibe.index == 0, "Data"] = ""
            df_exibe = _fmt_brl_df(df_exibe, col_valor)

            # Remove colunas técnicas/ruído
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
                sh = "Analítico"
                df_export.to_excel(writer, sheet_name=sh, index=False)
                wb, ws = writer.book, writer.sheets[sh]
                header = wb.add_format({"bold": True,"align":"center","valign":"vcenter","bg_color":"#F2F2F2","border":1})
                date_f = wb.add_format({"num_format":"dd/mm/yyyy","border":1})
                money  = wb.add_format({"num_format":"R$ #,##0.00","border":1})
                text   = wb.add_format({"border":1})
                tot    = wb.add_format({"bold": True,"bg_color":"#FCE5CD","border":1})
                totm   = wb.add_format({"bold": True,"bg_color":"#FCE5CD","border":1,"num_format":"R$ #,##0.00"})

                for j, name in enumerate(df_export.columns):
                    ws.write(0, j, name, header)
                    width, fmt = 18, text
                    if name.lower() == "data": width, fmt = 12, date_f
                    if name == col_valor:      width, fmt = 16, money
                    if "loja"  in name.lower(): width = 28
                    if "grupo" in name.lower(): width = 22
                    ws.set_column(j, j, width, fmt)

                ws.set_row(1, None, tot)
                if pd.notna(df_export.iloc[0][col_valor]):
                    ws.write_number(1, list(df_export.columns).index(col_valor), float(df_export.iloc[0][col_valor]), totm)
                if "Loja" in df_export.columns:
                    ws.write_string(1, list(df_export.columns).index("Loja"), "TOTAL", tot)
                ws.freeze_panes(1, 0)

            buf.seek(0)
            st.download_button(
                label="⬇️ Baixar Excel",
                data=buf,  # ou buf.getvalue()
                file_name="Relatorio_Analitico_Sangria.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_sangria_analitico"
            )


        # ====== Sintético ======
        elif visao == "Sintético":
            if "Loja" not in df_fil.columns:
                st.warning("Para 'Sintético', preciso da coluna **Loja**.")
            else:
                tmp = df_fil.copy()
                tmp["Data"] = pd.to_datetime(tmp["Data"], errors="coerce").dt.normalize()

                # Garante 'Grupo'
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
                if col_grupo and col_grupo != "Grupo":
                    ren[col_grupo] = "Grupo"
                df_agg = df_agg.rename(columns=ren).sort_values(["Data", "Grupo", "Loja"], na_position="last")

                total_sangria = df_agg["Sangria"].sum(min_count=1)
                linha_total = pd.DataFrame({"Grupo":["TOTAL"], "Loja":[""], "Data":[pd.NaT], "Sangria":[total_sangria]})
                df_exibe = pd.concat([linha_total, df_agg], ignore_index=True)

                # Exibição
                df_show = df_exibe.copy()
                df_show["Data"] = pd.to_datetime(df_show["Data"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
                df_show["Sangria"] = df_show["Sangria"].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

                st.dataframe(df_show[["Grupo","Loja","Data","Sangria"]], use_container_width=True, hide_index=True)

                # Export Excel
                df_exp = df_exibe[["Grupo","Loja","Data","Sangria"]].copy()
                df_exp["Data"] = pd.to_datetime(df_exp["Data"], errors="coerce")
                df_exp["Sangria"] = pd.to_numeric(df_exp["Sangria"], errors="coerce")

                buf = BytesIO()
                with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                    sh = "Sintético"
                    df_exp.to_excel(writer, sheet_name=sh, index=False)
                    wb, ws = writer.book, writer.sheets[sh]
                    header = wb.add_format({"bold": True,"align":"center","valign":"vcenter","bg_color":"#F2F2F2","border":1})
                    date_f = wb.add_format({"num_format":"dd/mm/yyyy","border":1})
                    money  = wb.add_format({"num_format":"R$ #,##0.00","border":1})
                    text   = wb.add_format({"border":1})
                    tot    = wb.add_format({"bold": True,"bg_color":"#FCE5CD","border":1})
                    totm   = wb.add_format({"bold": True,"bg_color":"#FCE5CD","border":1,"num_format":"R$ #,##0.00"})

                    for j, name in enumerate(["Grupo","Loja","Data","Sangria"]):
                        ws.write(0, j, name, header)
                    ws.set_column("A:A", 20, text)
                    ws.set_column("B:B", 28, text)
                    ws.set_column("C:C", 12, date_f)
                    ws.set_column("D:D", 14, money)

                    ws.set_row(1, None, tot)
                    if pd.notna(df_exp.iloc[0]["Sangria"]):
                        ws.write_number(1, 3, float(df_exp.iloc[0]["Sangria"]), totm)
                    ws.write_string(1, 0, "TOTAL", tot)
                    ws.freeze_panes(1, 0)

                buf.seek(0)
                st.download_button(
                    label="⬇️ Baixar Excel",
                    data=buf,
                    file_name="Relatorio_Sintetico_Sangria.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_sangria_sintetico"
                )

# -------------------------------
# Sub-aba: 🧰 CONTROLE DE SANGRIA (Comparativa Everest / Diferenças)
# -------------------------------
with sub_caixa:
    if df_sangria is None or df_sangria.empty:
        st.info("Sem dados de **sangria** disponíveis.")
    else:
        from io import BytesIO
        import unicodedata, re, os
        import pandas as pd
        from datetime import datetime

        # ===== helpers =====
        def _norm_txt(s: str) -> str:
            s = str(s or "").strip().lower()
            s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("utf-8")
            return s

        def eh_deposito_mask(df, cols_texto=None):
            if cols_texto is None:
                cols_texto = [
                    "Descrição Agrupada","Descrição","Historico","Histórico",
                    "Categoria","Obs","Observação","Tipo","Tipo Movimento"
                ]
            cols_texto = [c for c in cols_texto if c in df.columns]
            if not cols_texto:
                return pd.Series(False, index=df.index)
            txt = df[cols_texto].astype(str).agg(" ".join, axis=1).map(_norm_txt)
            padrao = r"""
                \bdeposito\b | \bdepsito\b | \bdep\b |
                credito\s+em\s+conta | envio\s*para\s*banco |
                transf(erencia)?\s*(p/?\s*banco|banco)
            """
            return txt.str.contains(padrao, flags=re.IGNORECASE | re.VERBOSE, regex=True, na=False)

        def brl(v):
            try:
                return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except Exception:
                return "R$ 0,00"

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

        # Filtros
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
                key="caixa_periodo_cmp",
            )
        with c2:
            lojas = sorted(df.get("Loja", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
            lojas_sel = st.multiselect("Lojas", options=lojas, default=[], key="caixa_lojas_cmp")
        with c3:
            descrs = sorted(df.get("Descrição Agrupada", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
            descrs_sel = st.multiselect("Descrição Agrupada", options=descrs, default=[], key="caixa_descr_cmp")
        with c4:
            visao = st.selectbox(
                "Visão do Relatório",
                options=["Comparativa Everest"],  # foquei na comparativa
                index=0,
                key="caixa_visao_cmp",
            )

        # aplica filtros
        df_fil = df[(df["Data"].dt.date >= dt_inicio) & (df["Data"].dt.date <= dt_fim)].copy()
        if lojas_sel:
            df_fil = df_fil[df_fil["Loja"].astype(str).isin(lojas_sel)]
        if descrs_sel:
            df_fil = df_fil[df_fil["Descrição Agrupada"].astype(str).isin(descrs_sel)]

        df_exibe = pd.DataFrame()

        # ======= Comparativa =======
        if visao == "Comparativa Everest":
            base = df_fil.copy()

            if "Data" not in base.columns or "Código Everest" not in base.columns or not col_valor:
                st.error("❌ Preciso de 'Data', 'Código Everest' e coluna de valor na aba Sangria.")
            else:
                # normalização
                base["Data"] = pd.to_datetime(base["Data"], dayfirst=True, errors="coerce").dt.normalize()
                base[col_valor] = pd.to_numeric(base[col_valor], errors="coerce").fillna(0.0)
                base["Código Everest"] = base["Código Everest"].astype(str).str.extract(r"(\d+)")

                # --- EXCLUI DEPÓSITOS (somente lado Sistema/Colibri) ---
                mask_dep_sys = eh_deposito_mask(base)
                with st.expander("🔎 Ver depósitos removidos (Colibri/CISS)"):
                    audit = base.loc[mask_dep_sys, :].copy()
                    if col_valor in audit.columns:
                        audit[col_valor] = audit[col_valor].map(brl)
                    st.dataframe(audit, use_container_width=True, hide_index=True)

                base = base.loc[~mask_dep_sys].copy()

                # agrega Sistema (já sem depósitos)
                df_sys = (
                    base.groupby(["Código Everest","Data"], as_index=False)[col_valor]
                        .sum()
                        .rename(columns={col_valor:"Sangria (Colibri/CISS)"})
                )

                # --- Everest ---
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
                else:
                    de = df_ev.copy()
                    de["Código Everest"]   = de[col_emp].astype(str).str.extract(r"(\d+)")
                    de["Fantasia Everest"] = de[col_fant] if col_fant else ""
                    de["Data"]             = pd.to_datetime(de[col_dt_ev], dayfirst=True, errors="coerce").dt.normalize()
                    de["Valor Lancamento"] = de[col_val_ev].map(parse_valor_brl_sheets).astype(float)
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
                    cmp["Sangria (Colibri/CISS)"] = cmp["Sangria (Colibri/CISS)"].fillna(0.0)
                    cmp["Sangria Everest"]        = cmp["Sangria Everest"].fillna(0.0)

                    # mapeamento Loja/Grupo
                    mapa = df_empresa.copy()
                    mapa.columns = [str(c).strip() for c in mapa.columns]
                    if "Código Everest" in mapa.columns:
                        mapa["Código Everest"] = mapa["Código Everest"].astype(str).str.extract(r"(\d+)")
                        cmp = cmp.merge(mapa[["Código Everest","Loja","Grupo"]].drop_duplicates(),
                                        on="Código Everest", how="left")

                    # fallback LOJA = Fantasia (linhas apenas do Everest)
                    cmp["Loja"] = cmp["Loja"].astype(str)
                    so_everest = (cmp["_merge"] == "right_only") & (cmp["Loja"].isin(["", "nan"]))
                    cmp.loc[so_everest, "Loja"] = cmp.loc[so_everest, "Fantasia Everest"]
                    cmp["Nao Mapeada?"] = so_everest

                    cmp["Diferença"] = cmp["Sangria (Colibri/CISS)"] - cmp["Sangria Everest"]

                    cmp = cmp[["Grupo","Loja","Código Everest","Data",
                               "Sangria (Colibri/CISS)","Sangria Everest","Diferença","Nao Mapeada?"]
                             ].sort_values(["Grupo","Loja","Código Everest","Data"])

                    total = {
                        "Grupo":"TOTAL","Loja":"","Código Everest":"","Data":pd.NaT,
                        "Sangria (Colibri/CISS)": cmp["Sangria (Colibri/CISS)"].sum(),
                        "Sangria Everest":        cmp["Sangria Everest"].sum(),
                        "Diferença":              cmp["Diferença"].sum(),
                        "Nao Mapeada?": False
                    }
                    df_exibe = pd.concat([pd.DataFrame([total]), cmp], ignore_index=True)

                    # ---- render no app
                    df_show = df_exibe.copy()
                    df_show["Data"] = pd.to_datetime(df_show["Data"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
                    for c in ["Sangria (Colibri/CISS)","Sangria Everest","Diferença"]:
                        df_show[c] = df_show[c].apply(
                            lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
                            if isinstance(v,(int,float)) else v
                        )

                    if "Nao Mapeada?" in df_show.columns and "Loja" in df_show.columns:
                        view = df_show.drop(columns=["Nao Mapeada?"], errors="ignore").copy()
                        mask_nm = (
                            df_show["Nao Mapeada?"].astype(bool)
                            if "Nao Mapeada?" in df_show.columns
                            else pd.Series(False, index=df_show.index)
                        )
                        def _paint_row(row: pd.Series):
                            styles = [""] * len(row.index)
                            try:
                                if mask_nm.loc[row.name] and "Loja" in row.index:
                                    idx = list(row.index).index("Loja")
                                    styles[idx] = "color: red; font-weight: 700"
                            except Exception:
                                pass
                            return styles
                        st.dataframe(view.style.apply(_paint_row, axis=1), use_container_width=True, height=520)
                    else:
                        st.dataframe(df_show.drop(columns=["Nao Mapeada?"], errors="ignore"),
                                     use_container_width=True, height=520)

                    # ========= EXPORTAÇÃO (com slicers quando possível) =========
                    import os
                    from io import BytesIO
                    import pandas as pd
                    
                    # ---------- 1) Monte o Excel base com TABELA 'tbl_dados' ----------
                    def preparar_df_export(cmp: pd.DataFrame) -> pd.DataFrame:
                        df = cmp.copy()
                    
                        df = df.drop(columns=["Nao Mapeada?"], errors="ignore")
                        if "Sangria (Sistema)" in df.columns:
                            df = df.rename(columns={"Sangria (Sistema)": "Sangria (Colibri/CISS)"})
                    
                        df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.normalize()
                        df["Ano"]  = df["Data"].dt.year
                        df["Mês"]  = df["Data"].dt.month
                    
                        for c in ["Sangria (Colibri/CISS)", "Sangria Everest", "Diferença"]:
                            if c in df.columns:
                                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
                    
                        ordem = [
                            "Data","Grupo","Loja","Código Everest",
                            "Sangria (Colibri/CISS)","Sangria Everest","Diferença",
                            "Mês","Ano"
                        ]
                        df = df[[c for c in ordem if c in df.columns]].copy()
                        return df
                    
                    
                    def escrever_base_xlsx_com_tabela(df: pd.DataFrame, path_out: str):
                        """Cria planilha 'Dados' com TABELA 'tbl_dados' e dados com tipos corretos (via XlsxWriter)."""
                        with pd.ExcelWriter(path_out, engine="xlsxwriter") as writer:
                            wb = writer.book
                            ws = wb.add_worksheet("Dados")
                            writer.sheets["Dados"] = ws
                    
                            fmt_header = wb.add_format({"bold": True, "align":"center", "valign":"vcenter",
                                                        "bg_color":"#F2F2F2", "border":1})
                            fmt_text   = wb.add_format({"border":1})
                            fmt_int    = wb.add_format({"border":1, "num_format":"0"})
                            fmt_date   = wb.add_format({"border":1, "num_format":"dd/mm/yyyy"})
                            fmt_money  = wb.add_format({"border":1, "num_format":"R$ #,##0.00"})
                    
                            # cabeçalho
                            for j, col in enumerate(df.columns):
                                ws.write(0, j, col, fmt_header)
                    
                            # linhas (mantendo tipos)
                            for i, row in df.iterrows():
                                r = i + 1
                                for j, col in enumerate(df.columns):
                                    val = row[col]
                                    if col == "Data" and pd.notna(val):
                                        ws.write_datetime(r, j, pd.to_datetime(val).to_pydatetime(), fmt_date)
                                    elif col in ("Ano","Mês","Código Everest"):
                                        ws.write_number(r, j, int(val) if pd.notna(val) else 0, fmt_int)
                                    elif col in ("Sangria (Colibri/CISS)","Sangria Everest","Diferença"):
                                        ws.write_number(r, j, float(val), fmt_money)
                                    else:
                                        ws.write(r, j, "" if pd.isna(val) else val, fmt_text)
                    
                            last_row = len(df)          # header = 0
                            last_col = len(df.columns)-1
                    
                            ws.add_table(0, 0, last_row, last_col, {
                                "name": "tbl_dados",
                                "style": "TableStyleMedium9",
                                "columns": [{"header": c} for c in df.columns],
                            })
                    
                            # larguras
                            idx = {c:i for i,c in enumerate(df.columns)}
                            if "Data" in idx:               ws.set_column(idx["Data"], idx["Data"], 12, fmt_date)
                            if "Grupo" in idx:              ws.set_column(idx["Grupo"], idx["Grupo"], 10, fmt_text)
                            if "Loja" in idx:               ws.set_column(idx["Loja"],  idx["Loja"],  28, fmt_text)
                            if "Código Everest" in idx:     ws.set_column(idx["Código Everest"], idx["Código Everest"], 14, fmt_int)
                            for c in ("Sangria (Colibri/CISS)","Sangria Everest","Diferença"):
                                if c in idx:                ws.set_column(idx[c], idx[c], 18, fmt_money)
                            if "Mês" in idx:                ws.set_column(idx["Mês"], idx["Mês"], 6, fmt_int)
                            if "Ano" in idx:                ws.set_column(idx["Ano"], idx["Ano"], 8, fmt_int)
                            ws.freeze_panes(1, 0)
                    
                        return os.path.abspath(path_out)
                    
                    # ---------- 2) Abra com Excel (COM) e crie as SEGMENTAÇÕES ----------
                    def criar_slicers_via_excel(path_xlsx: str, salvar_em: str = None):
                        """
                        Abre o arquivo no Excel Desktop via COM e cria segmentações para: Ano, Mês (ou Mes), Grupo, Loja.
                        Salva como .xlsx (sem macro). Não requer habilitar macros.
                        """
                        import pythoncom
                        import win32com.client as win32
                    
                        pythoncom.CoInitialize()
                        excel = win32.Dispatch("Excel.Application")
                        excel.Visible = False
                        excel.DisplayAlerts = False
                    
                        try:
                            wb = excel.Workbooks.Open(os.path.abspath(path_xlsx))
                            ws = wb.Worksheets("Dados")
                    
                            # garante que a tabela existe
                            try:
                                lo = ws.ListObjects("tbl_dados")
                            except Exception:
                                # cria tabela se não existir (do A1 até a última célula usada contínua)
                                used = ws.UsedRange
                                last_row = used.Row + used.Rows.Count - 1
                                last_col = used.Column + used.Columns.Count - 1
                                rng = ws.Range(ws.Cells(1,1), ws.Cells(last_row, last_col))
                                lo = ws.ListObjects.Add(SourceType=1, Source=rng, XlListObjectHasHeaders=1)
                                lo.Name = "tbl_dados"
                    
                            # remove slicers antigos
                            for shp in list(ws.Shapes):
                                try:
                                    # msoSlicer = 66 (constante)
                                    if shp.Type == 66:
                                        shp.Delete()
                                except Exception:
                                    pass
                    
                            def add_slicer(field_name, cell_addr, width, height):
                                # tenta Add2 (Excel mais novo), cai para Add se necessário
                                left = ws.Range(cell_addr).Left
                                top  = ws.Range(cell_addr).Top
                                sc = None
                                try:
                                    sc = wb.SlicerCaches.Add2(lo, field_name)
                                except Exception:
                                    try:
                                        sc = wb.SlicerCaches.Add(lo, field_name)
                                    except Exception:
                                        # tentativa com "Mes" se "Mês" falhar
                                        if field_name == "Mês":
                                            try:
                                                sc = wb.SlicerCaches.Add2(lo, "Mes")
                                            except Exception:
                                                sc = wb.SlicerCaches.Add(lo, "Mes")
                                if sc is not None:
                                    sc.Slicers.Add(ws, Name=f"slc_{field_name}", Caption=field_name,
                                                   Top=top, Left=left, Width=width, Height=height)
                    
                            # cria as segmentações (ajuste posições/dimensões se quiser)
                            add_slicer("Ano",  "L2", 130, 110)
                            add_slicer("Mês",  "L8", 130, 130)   # cai para "Mes" se necessário
                            add_slicer("Grupo","N2", 180, 180)
                            add_slicer("Loja", "N12",260, 320)
                    
                            # salva
                            out = salvar_em or path_xlsx
                            wb.SaveAs(os.path.abspath(out), FileFormat=51)  # 51 = xlOpenXMLWorkbook (.xlsx)
                            wb.Close(SaveChanges=False)
                            return os.path.abspath(out)
                    
                        finally:
                            excel.Quit()
                    
                    
                    # ---------- 3) Função principal: gera o XLSX com slicers ----------
                    def gerar_excel_com_slicers(cmp: pd.DataFrame, caminho_saida: str):
                        df = preparar_df_export(cmp)
                        base_tmp = os.path.splitext(caminho_saida)[0] + "_base.xlsx"
                        escrever_base_xlsx_com_tabela(df, base_tmp)
                        final = criar_slicers_via_excel(base_tmp, caminho_saida)
                        try:
                            os.remove(base_tmp)
                        except Exception:
                            pass
                        return final
                    
                    
                    # ========== EXEMPLO DE USO LOCAL ==========
                    if __name__ == "__main__":
                        # EXEMPLO: simular seu 'cmp'
                        dados = {
                            "Data": ["01/09/2025","01/09/2025","02/09/2025","02/09/2025"],
                            "Grupo": ["A","A","B","B"],
                            "Loja": ["Loja 1","Loja 2","Loja 1","Loja 3"],
                            "Código Everest": [123,124,125,126],
                            "Sangria (Colibri/CISS)": [100.0, 200.5, 50, 300],
                            "Sangria Everest": [90.0, 210.0, 50, 280],
                            "Diferença": [10.0, -9.5, 0, 20],
                        }
                        cmp = pd.DataFrame(dados)
                    
                        out = gerar_excel_com_slicers(cmp, "Sangria_Controle.xlsx")
                        print("Arquivo gerado:", out)
