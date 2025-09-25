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
                   # ========= EXPORTAÇÃO VIA TEMPLATE (preserva slicers sem openpyxl) =========
                    from io import BytesIO
                    import os, re, zipfile
                    import pandas as pd
                    import streamlit as st
                    import xml.etree.ElementTree as ET
                    
                    TEMPLATE_NOME = "modelo_segmentacao_sangria.xlsx"   # deixe no mesmo diretório do .py
                    
                    NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                    NS_REL  = "http://schemas.openxmlformats.org/package/2006/relationships"
                    ET.register_namespace("", NS_MAIN)
                    ET.register_namespace("r", NS_REL)
                    
                    def _col_letter(idx: int) -> str:
                        # 0->A, 1->B, ...
                        s = ""
                        idx += 1
                        while idx:
                            idx, r = divmod(idx-1, 26)
                            s = chr(65 + r) + s
                        return s
                    
                    def _find_sheet_and_table_paths(z: zipfile.ZipFile, sheet_name: str, table_name: str):
                        # 1) workbook -> achar sheet "Dados" e arquivo sheetX.xml
                        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
                        # pega r:id da sheet desejada
                        r_id = None
                        for sh in wb_xml.findall(f".//{{{NS_MAIN}}}sheet"):
                            if sh.get("name") == sheet_name:
                                r_id = sh.get(f"{{{NS_REL}}}id")
                                break
                        if not r_id:
                            raise RuntimeError(f"Planilha '{sheet_name}' não encontrada no template.")
                    
                        # 2) workbook rels -> target do r:id
                        wb_rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
                        sheet_target = None
                        for rel in wb_rels.findall(f".//{{{NS_REL}}}Relationship"):
                            if rel.get("Id") == r_id:
                                sheet_target = rel.get("Target")
                                break
                        if not sheet_target:
                            raise RuntimeError("Não foi possível resolver o arquivo da planilha de 'Dados' no template.")
                    
                        sheet_path = "xl/" + sheet_target.lstrip("/")
                        # 3) sheet rels -> tabela(s) referenciadas pela aba
                        rels_path = sheet_path.replace("worksheets/", "worksheets/_rels/") + ".rels"
                        if rels_path not in z.namelist():
                            raise RuntimeError("A planilha 'Dados' não possui relacionamentos para Tabelas.")
                    
                        sheet_rels = ET.fromstring(z.read(rels_path))
                        table_paths = []
                        for rel in sheet_rels.findall(f".//{{{NS_REL}}}Relationship"):
                            if rel.get("Type", "").endswith("/table"):
                                table_paths.append("xl/" + rel.get("Target").lstrip("/").replace("../", ""))
                    
                        # 4) achar qual table file tem displayName == tbl_dados
                        table_path = None
                        for p in table_paths:
                            tbl_xml = ET.fromstring(z.read(p))
                            if tbl_xml.tag.endswith("table") and tbl_xml.get("displayName") == table_name:
                                table_path = p
                                break
                        if not table_path:
                            raise RuntimeError(f"Tabela '{table_name}' não encontrada na planilha '{sheet_name}'.")
                    
                        return sheet_path, table_path
                    
                    def _headers_from_table(z: zipfile.ZipFile, table_path: str) -> list[str]:
                        tbl = ET.fromstring(z.read(table_path))
                        cols = []
                        for col in tbl.findall(f".//{{{NS_MAIN}}}tableColumn"):
                            cols.append(col.get("name"))
                        return cols
                    
                    def _prep_df(cmp: pd.DataFrame, headers: list[str]) -> pd.DataFrame:
                        df = cmp.copy()
                        df = df.drop(columns=["Nao Mapeada?"], errors="ignore")
                        if "Sangria (Sistema)" in df.columns and "Sangria (Colibri/CISS)" not in df.columns:
                            df = df.rename(columns={"Sangria (Sistema)": "Sangria (Colibri/CISS)"})
                    
                        df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.normalize()
                        df["Ano"]  = df["Data"].dt.year
                        df["Mês"]  = df["Data"].dt.month
                    
                        for c in ["Sangria (Colibri/CISS)","Sangria Everest","Diferença"]:
                            if c in df.columns:
                                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
                    
                        # alinhar "Mês"/"Mes" com o template
                        if "Mes" in headers and "Mês" in df.columns:
                            df = df.rename(columns={"Mês":"Mes"})
                        if "Mês" in headers and "Mes" in df.columns:
                            df = df.rename(columns={"Mes":"Mês"})
                    
                        # manter somente e na ordem dos headers da tabela
                        keep = [h for h in headers if h in df.columns]
                        return df[keep].copy()
                    
                    def exportar_via_template_zip(cmp: pd.DataFrame, template_path: str,
                                                  sheet_name="Dados", table_name="tbl_dados") -> BytesIO:
                        with open(template_path, "rb") as f:
                            tpl_bytes = f.read()
                    
                        zin  = zipfile.ZipFile(BytesIO(tpl_bytes), "r")
                        zout_buffer = BytesIO()
                        zout = zipfile.ZipFile(zout_buffer, "w", zipfile.ZIP_DEFLATED)
                    
                        # localizar arquivos
                        sheet_path, table_path = _find_sheet_and_table_paths(zin, sheet_name, table_name)
                        headers = _headers_from_table(zin, table_path)
                        df = _prep_df(cmp, headers)
                    
                        # 1) atualizar sheet xml (linhas da tabela a partir da linha do cabeçalho)
                        sheet_xml = ET.fromstring(zin.read(sheet_path))
                        # <dimension> para A1:... (opcional)
                        dim = sheet_xml.find(f".//{{{NS_MAIN}}}dimension")
                        # <sheetData>
                        sheetData = sheet_xml.find(f".//{{{NS_MAIN}}}sheetData")
                        if sheetData is None:
                            sheetData = ET.SubElement(sheet_xml, f"{{{NS_MAIN}}}sheetData")
                    
                        # descobrir linha do header da tabela olhando para o ref atual da tabela
                        tbl_xml = ET.fromstring(zin.read(table_path))
                        ref = tbl_xml.get("ref")  # ex: A1:I2
                        m = re.match(r"([A-Z]+)(\d+):([A-Z]+)(\d+)", ref)
                        col_ini_letters, row_ini = m.group(1), int(m.group(2))
                    
                        # reconstruir sheetData mantendo a linha do cabeçalho (row_ini)
                        # pega a row do header existente, se houver; senão cria uma.
                        new_sheetData = ET.Element(f"{{{NS_MAIN}}}sheetData")
                        header_row = None
                        for r in sheetData.findall(f"./{{{NS_MAIN}}}row"):
                            if int(r.get("r", "0")) == row_ini:
                                header_row = r
                                break
                        if header_row is None:
                            header_row = ET.Element(f"{{{NS_MAIN}}}row", {"r": str(row_ini)})
                    
                        # garante que os textos do cabeçalho batem com os headers do template
                        # escreve valores como inlineStr
                        for j, h in enumerate(headers):
                            col = _col_letter(j)  # 0->A
                            cell_ref = f"{col}{row_ini}"
                            # checa se célula existe
                            cell = None
                            for c in header_row.findall(f"./{{{NS_MAIN}}}c"):
                                if c.get("r") == cell_ref:
                                    cell = c; break
                            if cell is None:
                                cell = ET.SubElement(header_row, f"{{{NS_MAIN}}}c", {"r": cell_ref, "t": "inlineStr"})
                            else:
                                cell.clear(); cell.attrib.update({"r": cell_ref, "t": "inlineStr"})
                            is_ = ET.SubElement(cell, f"{{{NS_MAIN}}}is")
                            t_  = ET.SubElement(is_,  f"{{{NS_MAIN}}}t")
                            t_.text = h
                    
                        new_sheetData.append(header_row)
                    
                        # agora dados: linhas a partir de row_ini+1
                        for i, (_, row) in enumerate(df.iterrows(), start=row_ini+1):
                            r_el = ET.Element(f"{{{NS_MAIN}}}row", {"r": str(i)})
                            for j, h in enumerate(headers):
                                col = _col_letter(j)
                                cell_ref = f"{col}{i}"
                                val = row[h]
                                c = ET.SubElement(r_el, f"{{{NS_MAIN}}}c", {"r": cell_ref})
                                if pd.isna(val):
                                    continue
                                if h in ("Ano", "Mês", "Mes", "Código Everest", "Sangria (Colibri/CISS)", "Sangria Everest", "Diferença"):
                                    # numérico
                                    v = ET.SubElement(c, f"{{{NS_MAIN}}}v")
                                    v.text = str(float(val)) if isinstance(val, float) else str(int(val))
                                elif h == "Data":
                                    # texto (dd/mm/aaaa) para não depender de estilos
                                    c.set("t", "inlineStr")
                                    is_ = ET.SubElement(c, f"{{{NS_MAIN}}}is")
                                    t_  = ET.SubElement(is_, f"{{{NS_MAIN}}}t")
                                    try:
                                        t_.text = pd.to_datetime(val).strftime("%d/%m/%Y")
                                    except Exception:
                                        t_.text = str(val)
                                else:
                                    # texto genérico
                                    c.set("t", "inlineStr")
                                    is_ = ET.SubElement(c, f"{{{NS_MAIN}}}is")
                                    t_  = ET.SubElement(is_, f"{{{NS_MAIN}}}t")
                                    t_.text = "" if pd.isna(val) else str(val)
                            new_sheetData.append(r_el)
                    
                        # troca o sheetData antigo pelo novo
                        parent = sheet_xml
                        parent.remove(sheetData)
                        parent.append(new_sheetData)
                    
                        # atualiza <dimension>
                        last_row = row_ini + len(df)
                        last_col_letter = _col_letter(len(headers)-1)
                        new_dim_ref = f"{col_ini_letters}{row_ini}:{last_col_letter}{last_row}"
                        if dim is None:
                            dim = ET.SubElement(sheet_xml, f"{{{NS_MAIN}}}dimension", {"ref": new_dim_ref})
                        else:
                            dim.set("ref", new_dim_ref)
                    
                        # 2) atualizar table xml (ref e colunas)
                        tbl = tbl_xml  # já lido
                        tbl.set("ref", new_dim_ref)
                        # autoFilter
                        af = tbl.find(f".//{{{NS_MAIN}}}autoFilter")
                        if af is None:
                            af = ET.SubElement(tbl, f"{{{NS_MAIN}}}autoFilter")
                        af.set("ref", new_dim_ref)
                        # tableColumns
                        tcols = tbl.find(f".//{{{NS_MAIN}}}tableColumns")
                        if tcols is not None:
                            tbl.remove(tcols)
                        tcols = ET.SubElement(tbl, f"{{{NS_MAIN}}}tableColumns", {"count": str(len(headers))})
                        for idx, name in enumerate(headers, start=1):
                            ET.SubElement(tcols, f"{{{NS_MAIN}}}tableColumn", {"id": str(idx), "name": name})
                    
                        # 3) reempacotar: escrevemos sheet e table alterados; o resto copia igual
                        for name in zin.namelist():
                            if name == sheet_path:
                                zout.writestr(name, ET.tostring(sheet_xml, encoding="UTF-8", xml_declaration=True))
                            elif name == table_path:
                                zout.writestr(name, ET.tostring(tbl, encoding="UTF-8", xml_declaration=True))
                            else:
                                zout.writestr(name, zin.read(name))
                    
                        zin.close(); zout.close()
                        zout_buffer.seek(0)
                        return zout_buffer
                    
                    # ===== uso =====
                    template_path = os.path.join(os.getcwd(), TEMPLATE_NOME)
                    if not os.path.exists(template_path):
                        st.error(
                            "Template não encontrado. Coloque **modelo_segmentacao_sangria.xlsx** ao lado do seu `.py`.\n"
                            "Requisitos: planilha 'Dados', Tabela 'tbl_dados' em A1, slicers (Ano, Mês/Mes, Grupo, Loja)."
                        )
                    else:
                        try:
                            arquivo = exportar_via_template_zip(cmp, template_path)
                            st.download_button(
                                "⬇️ Baixar Excel",
                                data=arquivo,
                                file_name="Sangria_Controle.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_sangria_controle_excel",
                            )
                            st.caption("Exportei preenchendo o **template ZIP** (slicers preservados). Abra no Excel Desktop.")
                        except Exception as e:
                            st.error(f"Falha ao preencher template preservando slicers: {type(e).__name__}: {e}")
