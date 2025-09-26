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
        import re
        import pandas as pd
    
        if cols_texto is None:
            cols_texto = [
                "Descrição Agrupada","Descrição","Historico","Histórico",
                "Categoria","Obs","Observação","Tipo","Tipo Movimento"
            ]
    
        # mantém somente colunas existentes
        cols_texto = [c for c in cols_texto if c in df.columns]
        if not cols_texto:
            # nenhuma coluna de texto -> ninguém é depósito
            return pd.Series(False, index=df.index)
    
        # junta textos de forma segura e normaliza
        def _norm_txt(s: str) -> str:
            s = str(s or "").strip().lower()
            import unicodedata
            return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("utf-8")
    
        # Series com o texto combinado por linha
        txt = (
            df[cols_texto]
            .astype(str)                # garante string
            .fillna("")                 # zera nulos
            .agg(" ".join, axis=1)      # junta numa só string por linha
            .map(_norm_txt)             # normaliza
        )
    
        padrao = r"""
            \bdeposito\b | \bdepsito\b | \bdep\b |
            credito\s+em\s+conta | envio\s*para\s*banco |
            transf(erencia)?\s*(p/?\s*banco|banco)
        """
        rx = re.compile(padrao, re.IGNORECASE | re.VERBOSE)
    
        # evita usar .str.contains; usa search do regex diretamente
        return txt.apply(lambda s: bool(rx.search(s)))

    
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
        st.dataframe(audit.reset_index(drop=True), use_container_width=True, hide_index=True, height=480)
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
# -------------------------------
# Sub-aba: 🧰 CONTROLE DE SANGRIA (Comparativa Everest / Diferenças)
# -------------------------------
with sub_caixa:
    if df_sangria is None or df_sangria.empty:
        st.info("Sem dados de **sangria** disponíveis.")
    else:
        from io import BytesIO
        import unicodedata, re
        import pandas as pd
        import numpy as np

        # ===== helpers =====
        def _norm_txt(s: str) -> str:
            s = str(s or "").strip().lower()
            return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("utf-8")

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

        # ===== Filtros =====
        c1, c2, c3, c4, c5 = st.columns([1.2, 1.2, 1.2, 1.2, 1.2])

        with c2:
            try:
                grupos_df = sorted(df.get("Grupo", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
            except Exception:
                grupos_df = []
            try:
                grupos_emp = sorted(df_empresa.get("Grupo", pd.Series([], dtype=str)).dropna().astype(str).unique().tolist())
            except Exception:
                grupos_emp = []
            opcoes_grupo = sorted({*grupos_df, *grupos_emp})
            grupos_sel = st.multiselect("Grupos", options=opcoes_grupo, default=[], key="caixa_grupos_cmp")

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

        with c3:
            df_opt = df[(df["Data"].dt.date >= dt_inicio) & (df["Data"].dt.date <= dt_fim)].copy()
            if grupos_sel and "Grupo" in df_opt.columns:
                df_opt = df_opt[df_opt["Grupo"].astype(str).isin(grupos_sel)]
            opcoes_lojas = sorted(df_opt.get("Loja", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
            prev_sel = st.session_state.get("caixa_lojas_cmp", [])
            default_sel = [x for x in prev_sel if x in opcoes_lojas]
            lojas_sel = st.multiselect("Lojas", options=opcoes_lojas, default=default_sel, key="caixa_lojas_cmp")

        with c4:
            visao = st.selectbox("Visão do Relatório", options=["Comparativa Everest"], index=0, key="caixa_visao_cmp")

        with c5:
            filtro_dif = st.selectbox("Filtro por Diferença", options=["Todas", "Diferenças", "Sem diferença"], index=0, key="caixa_filtro_diferenca")

        # aplica filtros
        df_fil = df[(df["Data"].dt.date >= dt_inicio) & (df["Data"].dt.date <= dt_fim)].copy()
        if lojas_sel:
            df_fil = df_fil[df_fil["Loja"].astype(str).isin(lojas_sel)]
        if grupos_sel and "Grupo" in df_fil.columns:
            df_fil = df_fil[df_fil["Grupo"].astype(str).isin(grupos_sel)]

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
                base_raw = base.copy()  # preserva para INCLUÍDOS

                mask_extra = (
                    base["Descrição Agrupada"].astype(str).str.contains(r"\b(maionese|Moeda Estrangeira)\b", regex=True, na=False)
                    if "Descrição Agrupada" in base.columns else pd.Series(False, index=base.index)
                )
                mask_dep_sys = eh_deposito_mask(base) | mask_extra

                with st.expander("🔎 Ver depósitos/termos removidos (Colibri/CISS)"):
                    audit = base.loc[mask_dep_sys, :].copy()
                    if "Data" in audit.columns:
                        audit["Data"] = pd.to_datetime(audit["Data"], errors="coerce").dt.strftime("%d/%m/%Y")
                    if col_valor in audit.columns:
                        audit[col_valor] = audit[col_valor].map(brl)
                    st.dataframe(audit, use_container_width=True, hide_index=True)

                # ===== INCLUÍDOS (mostra Descrição e Descrição Agrupada) =====
                inc = base_raw.loc[~mask_dep_sys, :].copy()
                cols_inc = ["Grupo","Loja","Código Everest","Data"]
                if "Descrição" in inc.columns: cols_inc.append("Descrição")
                if "Descrição Agrupada" in inc.columns: cols_inc.append("Descrição Agrupada")
                if col_valor in inc.columns: cols_inc.append(col_valor)
                cols_inc = [c for c in cols_inc if c in inc.columns]
                if "Data" in inc.columns:
                    inc["Data"] = pd.to_datetime(inc["Data"], errors="coerce").dt.strftime("%d/%m/%Y")
                if col_valor in inc.columns:
                    inc[col_valor] = inc[col_valor].map(brl)
                with st.expander("🧾 INCLUÍDOS (lado Colibri/CISS)"):
                    st.dataframe(inc[cols_inc], use_container_width=True, hide_index=True)

                # ===== agrega Sistema (já sem depósitos) =====
                base = base_raw.loc[~mask_dep_sys].copy()
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
                pref_comp      = ["dcompetencia","datacompetencia","datadecompetencia","competencia","dtcompetencia"]
                fallback_lcto  = ["dlancamento","datadelancamento","data"]
                col_dt_ev = next((cmap[k] for k in pref_comp if k in cmap), next((cmap[k] for k in fallback_lcto if k in cmap), None))
                col_val_ev= next((orig for norm, orig in cmap.items() if norm in ("valorlancamento","valorlcto","valor")), None)
                col_fant  = next((orig for norm, orig in cmap.items() if norm in ("fantasiaempresa","fantasia")), None)

                if not all([col_emp, col_dt_ev, col_val_ev]):
                    st.error("❌ Na 'Sangria Everest' preciso de 'Empresa', 'D. Competência' (ou 'D. Lançamento') e 'Valor Lancamento'.")
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

                    # mapeamento Loja/Grupo (1 loja por Código Everest)
                    mapa = df_empresa.copy()
                    mapa.columns = [str(c).strip() for c in mapa.columns]
                    if "Código Everest" in mapa.columns:
                        mapa["Código Everest"] = mapa["Código Everest"].astype(str).str.extract(r"(\d+)")
                        mapa["__prio__"] = mapa["Loja"].astype(str).str.contains(r"(embarque|checkin)", case=False, na=False).astype(int)
                        mapa_unico = (
                            mapa.sort_values(["Código Everest","__prio__","Loja"])
                                .drop_duplicates(subset=["Código Everest"], keep="first")
                                [["Código Everest","Loja","Grupo"]]
                        )
                        cmp = cmp.merge(mapa_unico, on="Código Everest", how="left")

                    # fallback LOJA = Fantasia (linhas só do Everest)
                    cmp["Loja"] = cmp["Loja"].astype(str)
                    so_everest = (cmp["_merge"] == "right_only") & (cmp["Loja"].isin(["", "nan"]))
                    cmp.loc[so_everest, "Loja"] = cmp.loc[so_everest, "Fantasia Everest"]
                    cmp["Nao Mapeada?"] = so_everest

                    # diferença + filtro
                    cmp["Diferença"] = pd.to_numeric(cmp["Sangria (Colibri/CISS)"] - cmp["Sangria Everest"], errors="coerce").fillna(0.0)
                    TOL = 0.0099
                    eh_zero = np.isclose(cmp["Diferença"].to_numpy(dtype=float), 0.0, atol=TOL)
                    if filtro_dif == "Diferenças":
                        cmp = cmp[~eh_zero]
                        st.caption("Mostrando apenas linhas com diferença (|Diferença| > R$ 0,01).")
                    elif filtro_dif == "Sem diferença":
                        cmp = cmp[eh_zero]
                        st.caption("Mostrando apenas linhas sem diferença (|Diferença| ≤ R$ 0,01).")
                    if grupos_sel:
                        cmp = cmp[cmp["Grupo"].astype(str).isin(grupos_sel)]

                    # ordenação + TOTAL
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

                    # ---- render no app (coluna Selecionado à direita de Diferença) ----
                    df_show = df_exibe.copy()
                    df_show["Data"] = pd.to_datetime(df_show["Data"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
                    for c in ["Sangria (Colibri/CISS)","Sangria Everest","Diferença"]:
                        df_show[c] = df_show[c].apply(
                            lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
                            if isinstance(v,(int,float)) else v
                        )

                    df_view = df_show.drop(columns=["Nao Mapeada?"], errors="ignore").copy()
                    is_total = df_view["Grupo"].astype(str).str.upper().eq("TOTAL") if "Grupo" in df_view.columns else pd.Series(False, index=df_view.index)
                    df_view.loc[~is_total, "Selecionado"] = False
                    df_view.loc[ is_total, "Selecionado"] = ""  # TOTAL sem checkbox

                    # posiciona após Diferença
                    cols = list(df_view.columns)
                    if "Selecionado" in cols and "Diferença" in cols:
                        cols.remove("Selecionado")
                        pos = cols.index("Diferença") + 1
                        cols = cols[:pos] + ["Selecionado"] + cols[pos:]

                    st.data_editor(
                        df_view[cols],
                        use_container_width=True,
                        hide_index=True,
                        num_rows="fixed",
                        column_config={"Selecionado": st.column_config.CheckboxColumn("Selecionado")},
                        height=520,
                        key="cmp_editor_selecionado",
                    )

                    # ========= EXPORTAÇÃO básica =========
                    def _prep_df_export(cmp_df: pd.DataFrame) -> pd.DataFrame:
                        d = cmp_df.drop(columns=["Nao Mapeada?"], errors="ignore").copy()
                        d["Data"] = pd.to_datetime(d["Data"], errors="coerce").dt.normalize()
                        for c in ["Sangria (Colibri/CISS)","Sangria Everest","Diferença"]:
                            if c in d.columns:
                                d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0.0)
                        return d[["Data","Grupo","Loja","Código Everest",
                                  "Sangria (Colibri/CISS)","Sangria Everest","Diferença"]]

                    def exportar_xlsxwriter(cmp_df: pd.DataFrame) -> BytesIO:
                        from xlsxwriter import Workbook
                        d = _prep_df_export(cmp_df)
                        buf = BytesIO()
                        wb = Workbook(buf, {"in_memory": True})
                        ws = wb.add_worksheet("Dados")
                        fmt_header = wb.add_format({"bold": True, "align": "center", "valign": "vcenter",
                                                    "bg_color": "#F2F2F2", "border": 1})
                        fmt_text = wb.add_format({"border": 1})
                        fmt_date = wb.add_format({"border": 1, "num_format": "dd/mm/yyyy"})
                        fmt_money = wb.add_format({"border": 1, "num_format": "R$ #,##0.00"})

                        headers = list(d.columns)
                        for j, c in enumerate(headers):
                            ws.write(0, j, c, fmt_header)
                        for i, row in d.iterrows():
                            r = i + 1
                            for j, c in enumerate(headers):
                                v = row[c]
                                if c == "Data" and pd.notna(v):
                                    ws.write_datetime(r, j, pd.to_datetime(v).to_pydatetime(), fmt_date)
                                elif c in ("Sangria (Colibri/CISS)", "Sangria Everest", "Diferença"):
                                    ws.write_number(r, j, float(v), fmt_money)
                                else:
                                    ws.write(r, j, "" if pd.isna(v) else v, fmt_text)
                        ws.add_table(0, 0, len(d), len(headers) - 1, {
                            "name": "tbl_dados",
                            "style": "TableStyleMedium9",
                            "columns": [{"header": h} for h in headers],
                        })
                        wb.close()
                        buf.seek(0)
                        return buf

                    st.download_button(
                        label="⬇️ Baixar Excel",
                        data=exportar_xlsxwriter(cmp),
                        file_name="Sangria_Controle.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_sangria_controle_excel",
                    )
