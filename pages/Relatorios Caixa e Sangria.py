# pages/05_Relatorios_Caixa_Sangria.py
import streamlit as st
import pandas as pd
import numpy as np
import re, json
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials

st.set_page_config(
    page_title="Relatórios Caixa e Sangria",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)
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
        # Filtros
        c1, c2, c3, c4, c5 = st.columns([1.2, 1.2, 1.2, 1.2, 1.2])
        with c2:
            # tenta pegar grupos do df_sangria; se não houver, usa df_empresa
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
            
            # opções de lojas dependem do período e (se houver) do(s) Grupo(s) selecionado(s)
            df_opt = df[(df["Data"].dt.date >= dt_inicio) & (df["Data"].dt.date <= dt_fim)].copy()
        
            if grupos_sel and "Grupo" in df_opt.columns:
                df_opt = df_opt[df_opt["Grupo"].astype(str).isin(grupos_sel)]
        
            # lista final de lojas possíveis (somente as que aparecem na tela com os filtros acima)
            opcoes_lojas = sorted(
                df_opt.get("Loja", pd.Series(dtype=str)).dropna().astype(str).unique().tolist()
            )
        
            # preserva seleção válida se o usuário já tinha escolhido lojas
            prev_sel = st.session_state.get("caixa_lojas_cmp", [])
            default_sel = [x for x in prev_sel if x in opcoes_lojas]
        
            lojas_sel = st.multiselect(
                "Lojas",
                options=opcoes_lojas,
                default=default_sel,
                key="caixa_lojas_cmp",
            )

        
        
        with c4:
            visao = st.selectbox(
                "Visão do Relatório",
                options=["Comparativa Everest"],
                index=0,
                key="caixa_visao_cmp",
            )
        
        with c5:
            # 🔎 NOVO filtro por diferença (atua depois que 'cmp' é calculado)
            filtro_dif = st.selectbox(
                "Filtro por Diferença",
                options=["Todas", "Diferenças", "Sem diferença"],
                index=0,
                key="caixa_filtro_diferenca",
            )
       

        # aplica filtros
        df_fil = df[(df["Data"].dt.date >= dt_inicio) & (df["Data"].dt.date <= dt_fim)].copy()
        if lojas_sel:
            df_fil = df_fil[df_fil["Loja"].astype(str).isin(lojas_sel)]
        # 🔎 se a aba Sangria já tiver "Grupo"
        if grupos_sel and "Grupo" in df_fil.columns:
            df_fil = df_fil[df_fil["Grupo"].astype(str).isin(grupos_sel)]

       
        df_exibe = pd.DataFrame()

        # ======= Comparativa =======
        # ======= Comparativa =======
        if visao == "Comparativa Everest":
            base = df_fil.copy()
        
            if "Data" not in base.columns or "Código Everest" not in base.columns or not col_valor:
                st.error("❌ Preciso de 'Data', 'Código Everest' e coluna de valor na aba Sangria.")
            else:
                import re
                from streamlit import column_config as cc
                
                # Import robusto do Cell (funciona em várias versões do gspread)
                try:
                    from gspread.cell import Cell          # gspread >= 5.x
                except Exception:
                    try:
                        from gspread.models import Cell    # versões antigas
                    except Exception:
                        Cell = None  # sem Cell → usamos fallback célula a célula
                
                def _excel_col_letter(idx_zero_based: int) -> str:
                    n = idx_zero_based + 1
                    s = ""
                    while n:
                        n, r = divmod(n - 1, 26)
                        s = chr(65 + r) + s
                    return s
        
                # ====================== normalização ======================
                base["Data"] = pd.to_datetime(base["Data"], dayfirst=True, errors="coerce").dt.normalize()
                base[col_valor] = pd.to_numeric(base[col_valor], errors="coerce").fillna(0.0)
                base["Código Everest"] = base["Código Everest"].astype(str).str.extract(r"(\d+)")
        
                # ====================== opções globais de "Descrição Agrupada" a partir do Google Sheets ======================
                # Tentamos primeiro uma aba dedicada ("Tabela Sangria"); se não existir, usamos a própria "Sangria".
                try:
                    ws_tab_desc = planilha_empresa.worksheet("Tabela Sangria")
                except Exception:
                    ws_tab_desc = planilha_empresa.worksheet("Sangria")
                
                df_descricoes = pd.DataFrame(ws_tab_desc.get_all_records())
                df_descricoes.columns = [c.strip() for c in df_descricoes.columns]
                
                def _limpo(s: str) -> str:
                    s = str(s or "").strip()
                    s = re.sub(r"\s+", " ", s)
                    return s
                
                def _norm_acento(s: str) -> str:
                    import unicodedata
                    s0 = _limpo(s)
                    s0 = unicodedata.normalize("NFKD", s0).encode("ASCII","ignore").decode("ASCII")
                    return s0.lower()
                
                # pega TODAS as opções únicas da coluna "Descrição Agrupada" do Sheets
                opcoes_desc_global = []
                if "Descrição Agrupada" in df_descricoes.columns:
                    bruto = df_descricoes["Descrição Agrupada"].dropna().astype(str).map(_limpo)
                    bruto = bruto[~bruto.isin(["", "nan", "none"])]
                    if not bruto.empty:
                        s = pd.Series(bruto.tolist())
                        norm = s.map(_norm_acento)
                        df_opts = pd.DataFrame({"norm": norm, "orig": s})
                        # mantém a grafia mais frequente para cada chave normalizada
                        escolha = df_opts.groupby("norm")["orig"].agg(lambda col: col.value_counts().idxmax())
                        opcoes_desc_global = sorted(escolha.tolist(), key=lambda x: x.lower())
                
                # fallback mínimo
                if not opcoes_desc_global:
                    opcoes_desc_global = ["Outros"]
        
                # ====================== filtros/removidos/incluídos ======================
                # Classificar SOMENTE pela "Descrição Agrupada"
                desc_norm = (
                    base.get("Descrição Agrupada", pd.Series("", index=base.index))
                        .astype(str).fillna("").map(_norm_acento).str.strip()
                )
                rotulos_deposito = {"depósito", "deposito", "dep", "moeda estrangeira", "maionese"}
                rotulos_deposito_norm = {_norm_acento(x) for x in rotulos_deposito}
                # True = REMOVIDO/DEPÓSITO | False = INCLUÍDO
                mask_dep_sys = desc_norm.isin(rotulos_deposito_norm)
        
                # colunas técnicas para esconder
                _cols_hide = ["Mês", "Mes", "Ano", "Duplicidade", "Sistema"]
        
                # ====================== Estado (CÓDIGOS) aplicado para filtrar os expanders ======================
                def _only_digits(x):
                    x = "" if x is None else str(x)
                    return re.sub(r"\D+", "", x)
        
                codigos_aplicados = set(map(_only_digits, st.session_state.get("cmp_codigos_selecionados", set())))
                codigos_aplicados = set(filter(None, codigos_aplicados))
                tem_filtro_codigo = bool(codigos_aplicados)
        
                # ✅ Versor do editor – precisa estar fora de funções e antes dos editores
                if "rev_desc" not in st.session_state:
                    st.session_state["rev_desc"] = 0
        
                # Opções de "Descrição Agrupada" vindas do Google Sheets (Tabela Sangria),
                # removendo duplicatas (insensível a maiúsc/minúsc, acentos e espaços extras).
                def _desc_options_from_sheet(df_extra: pd.DataFrame | None = None) -> list[str]:
                    import unicodedata, re
                    def _clean(s: str) -> str:
                        s = str(s or "").strip()
                        s = re.sub(r"\s+", " ", s)  # normaliza espaços
                        return s
                    def _norm_key(s: str) -> str:
                        s0 = _clean(s)
                        s0 = unicodedata.normalize("NFKD", s0).encode("ASCII", "ignore").decode("ASCII")
                        return s0.lower()
                    base_opts = []
                    try:
                        if "Descrição Agrupada" in df_descricoes.columns:
                            base_opts = df_descricoes["Descrição Agrupada"].dropna().astype(str).tolist()
                    except Exception:
                        pass
                    extras = []
                    if df_extra is not None and "Descrição Agrupada" in df_extra.columns:
                        extras = df_extra["Descrição Agrupada"].dropna().astype(str).tolist()
                    candidatos = [x for x in map(_clean, base_opts + extras) if x and x.lower() not in ("nan", "none")]
                    if not candidatos:
                        return ["Outros"]
                    s = pd.Series(candidatos)
                    norm = s.map(_norm_key)
                    df_opts = pd.DataFrame({"norm": norm, "orig": s})
                    escolha = df_opts.groupby("norm")["orig"].agg(lambda col: col.value_counts().idxmax())
                    return sorted(escolha.tolist(), key=lambda x: x.lower())
        
                # helpers p/ sheets
                WS_SISTEMA = "Sangria"  # ⬅️ ajuste se necessário
        
                def _sheet_df_with_row(ws):
                    vals = ws.get_all_values()
                    if not vals:
                        return pd.DataFrame(), {}
                    header = vals[0]
                    rows = vals[1:]
                    df_ws = pd.DataFrame(rows, columns=header)
                    df_ws["_row"] = (df_ws.index + 2).astype(int)  # linha real no Sheets
                    col_map = {c: i + 1 for i, c in enumerate(header)}  # coluna -> índice 1-based
                    return df_ws, col_map
        
                def _desc_options_flex(fallback_df: pd.DataFrame) -> list[str]:
                    try:
                        base_opts = set(df_descricoes["Descrição Agrupada"].astype(str).dropna().unique())
                    except Exception:
                        base_opts = set()
                    cur_opts = set()
                    if "Descrição Agrupada" in fallback_df.columns:
                        cur_opts = set(fallback_df["Descrição Agrupada"].astype(str).dropna().unique())
                    opts = sorted({o.strip() for o in (base_opts | cur_opts) if o and o.strip()})
                    return opts if opts else ["Outros"]
        
                # 👇 UTIL: encontra a coluna Observação ignorando acento/caixa/espaços
                def _find_col(cols, candidates=("Observação", "Observacao", "Obs")):
                    import unicodedata, re
                    def norm(s):
                        s = unicodedata.normalize("NFKD", str(s)).encode("ASCII","ignore").decode("ASCII")
                        return re.sub(r"\s+", " ", s).strip().lower()
                    cmap = {norm(c): c for c in cols}
                    for cand in candidates:
                        key = norm(cand)
                        if key in cmap:
                            return cmap[key]
                    for k, orig in cmap.items():  # fallback: qualquer coisa que comece com "observa"
                        if k.startswith("observa"):
                            return orig
                    return None
        
                # ====================== EXPANDERS (com edição) ======================
                # -------- INCLUÍDOS --------
                with st.expander("Sangria(Colibri/CISS)"):
                    audit_in_raw = base.loc[~mask_dep_sys, :].copy()  # mantém Duplicidade
                    if tem_filtro_codigo and "Código Everest" in audit_in_raw.columns:
                        audit_in_raw["_cod"] = audit_in_raw["Código Everest"].astype(str).str.extract(r"(\d+)")
                        audit_in_raw = audit_in_raw[audit_in_raw["_cod"].isin(codigos_aplicados)].drop(columns=["_cod"])
        
                    # Opções para o dropdown
                    opcoes_desc_in = _desc_options_from_sheet(audit_in_raw)
        
                    # Visão para tela
                    audit_in_view = audit_in_raw.drop(columns=_cols_hide, errors="ignore").copy()
                    if col_valor in audit_in_view.columns:
                        audit_in_view[col_valor] = audit_in_view[col_valor].map(brl)
                    if "Data" in audit_in_view.columns:
                        audit_in_view["Data"] = pd.to_datetime(audit_in_view["Data"], errors="coerce").dt.date
                    if tem_filtro_codigo:
                        st.caption(f"Filtrando por {len(codigos_aplicados)} código(s) selecionado(s).")
                    if tem_filtro_codigo and audit_in_view.empty:
                        st.info("Nenhum item incluído para os códigos selecionados.")
        
                    col_cfg_in = {c: cc.TextColumn(disabled=True, label=c) for c in audit_in_view.columns}
        
                    # 🔓 Observação (texto livre) – habilitar a edição
                    obs_col_in = _find_col(audit_in_view.columns)
                    if obs_col_in:
                        audit_in_view[obs_col_in] = audit_in_view[obs_col_in].astype(str).replace({"nan": ""}).fillna("")
                        col_cfg_in[obs_col_in] = cc.TextColumn(
                            label=obs_col_in,
                            help="Digite livremente; será salvo no Google Sheets.",
                            disabled=False,
                        )
        
                    # Descrição Agrupada como select editável
                    if "Descrição Agrupada" in audit_in_view.columns:
                        presentes_in = audit_in_view["Descrição Agrupada"].dropna().astype(str).map(_limpo).tolist()
                        options_in = sorted(set(opcoes_desc_global) | set(presentes_in), key=lambda x: x.lower())
                        col_cfg_in["Descrição Agrupada"] = cc.SelectboxColumn(
                            label="Descrição Agrupada",
                            options=options_in,
                            help="Escolha a descrição agrupada para esta linha."
                        )
        
                    if "Data" in audit_in_view.columns:
                        col_cfg_in["Data"] = cc.DateColumn(label="Data", format="DD/MM/YYYY", disabled=True)
        
                    with st.form("form_editar_desc_incluidos", clear_on_submit=False):
                        edited_in_view = st.data_editor(
                            audit_in_view,
                            use_container_width=True,
                            hide_index=True,
                            height=320,
                            column_config=col_cfg_in,
                            key=f"editor_incluidos_desc_{st.session_state['rev_desc']}",
                        )
                        c_save_in, _ = st.columns([1, 6])
                        salvar_in = c_save_in.form_submit_button("Atualizar Google Sheets")
        
                    if salvar_in:
                        try:
                            # Antes / Depois - Descrição Agrupada
                            antes_desc  = audit_in_raw.get("Descrição Agrupada", pd.Series("", index=audit_in_raw.index)).astype(str).fillna("").reset_index(drop=True)
                            depois_desc = edited_in_view.get("Descrição Agrupada", pd.Series("", index=audit_in_raw.index)).astype(str).fillna("").reset_index(drop=True)
                            # Antes / Depois - Observação
                            if obs_col_in:
                                antes_obs  = audit_in_raw.get(obs_col_in, pd.Series("", index=audit_in_raw.index)).astype(str).fillna("").reset_index(drop=True)
                                depois_obs = edited_in_view.get(obs_col_in, pd.Series("", index=audit_in_raw.index)).astype(str).fillna("").reset_index(drop=True)
                            else:
                                antes_obs = depois_obs = pd.Series("", index=antes_desc.index)
        
                            mask_changed = (antes_desc != depois_desc) | (antes_obs != depois_obs)
                            if not mask_changed.any():
                                st.success("Nada para atualizar — nenhuma alteração em incluídos.")
                            else:
                                ws_sys = planilha_empresa.worksheet(WS_SISTEMA)
                                df_ws, col_map = _sheet_df_with_row(ws_sys)
                                if "Duplicidade" not in df_ws.columns:
                                    st.error("A aba do Sheets precisa ter a coluna 'Duplicidade'.")
                                else:
                                    df_ws["Duplicidade"] = df_ws["Duplicidade"].astype(str)
                                    col_idx_desc = col_map.get("Descrição Agrupada")
                                    obs_sheet_name = _find_col(col_map.keys())
                                    col_idx_obs = col_map.get(obs_sheet_name) if obs_sheet_name else None
        
                                    keys = audit_in_raw["Duplicidade"].astype(str).reset_index(drop=True)
        
                                    for i in mask_changed[mask_changed].index:
                                        dup_key   = keys.iloc[i]
                                        nova_desc = depois_desc.iloc[i].strip()
                                        nova_obs  = depois_obs.iloc[i].strip()
        
                                        hits = df_ws.index[df_ws["Duplicidade"] == dup_key].tolist()
                                        if not hits:
                                            continue
        
                                        updates = []
                                        for h in hits:
                                            row_num = int(df_ws.loc[h, "_row"])
                                            if Cell is not None:
                                                if col_idx_desc:
                                                    updates.append(Cell(row=row_num, col=col_idx_desc, value=nova_desc))
                                                if col_idx_obs is not None:
                                                    updates.append(Cell(row=row_num, col=col_idx_obs,  value=nova_obs))
                                            else:
                                                if col_idx_desc:
                                                    updates.append((row_num, col_idx_desc, nova_desc))
                                                if col_idx_obs is not None:
                                                    updates.append((row_num, col_idx_obs,  nova_obs))
                                        if updates:
                                            if Cell is not None:
                                                ws_sys.update_cells(updates, value_input_option="USER_ENTERED")
                                            else:
                                                for row_num, col_idx, value in updates:
                                                    a1 = f"{_excel_col_letter(col_idx-1)}{row_num}"
                                                    ws_sys.update(a1, value, value_input_option="USER_ENTERED")
        
                                st.success("Alterações salvas (incluídos).")
                                st.session_state["rev_desc"] += 1
                                st.rerun()
        
                        except Exception as e:
                            st.error(f"Falha ao atualizar (incluídos): {type(e).__name__}: {e}")
        
                # -------- REMOVIDOS --------
                with st.expander("Depósitos(Colibri/CISS)"):
                    audit_out_raw = base.loc[mask_dep_sys, :].copy()  # mantém Duplicidade
                    if tem_filtro_codigo and "Código Everest" in audit_out_raw.columns:
                        audit_out_raw["_cod"] = audit_out_raw["Código Everest"].astype(str).str.extract(r"(\d+)")
                        audit_out_raw = audit_out_raw[audit_out_raw["_cod"].isin(codigos_aplicados)].drop(columns=["_cod"])
        
                    opcoes_desc_out = _desc_options_from_sheet(audit_out_raw)
        
                    audit_out_view = audit_out_raw.drop(columns=_cols_hide, errors="ignore").copy()
                    if col_valor in audit_out_view.columns:
                        audit_out_view[col_valor] = audit_out_view[col_valor].map(brl)
                    if "Data" in audit_out_view.columns:
                        audit_out_view["Data"] = pd.to_datetime(audit_out_view["Data"], errors="coerce").dt.date
                    if tem_filtro_codigo:
                        st.caption(f"Filtrando por {len(codigos_aplicados)} código(s) selecionado(s).")
                    if tem_filtro_codigo and audit_out_view.empty:
                        st.info("Nenhum depósito/remoção para os códigos selecionados.")
        
                    col_cfg_out = {c: cc.TextColumn(disabled=True, label=c) for c in audit_out_view.columns}
        
                    if "Descrição Agrupada" in audit_out_view.columns:
                        presentes_out = audit_out_view["Descrição Agrupada"].dropna().astype(str).map(_limpo).tolist()
                        options_out = sorted(set(opcoes_desc_global) | set(presentes_out), key=lambda x: x.lower())
                        col_cfg_out["Descrição Agrupada"] = cc.SelectboxColumn(
                            label="Descrição Agrupada", options=options_out, help="Escolha a descrição agrupada para esta linha."
                        )
        
                    if "Data" in audit_out_view.columns:
                        col_cfg_out["Data"] = cc.DateColumn(label="Data", format="DD/MM/YYYY", disabled=True)
        
                    # 🔓 Observação (texto livre) – habilitar a edição
                    obs_col_out = _find_col(audit_out_view.columns)
                    if obs_col_out:
                        audit_out_view[obs_col_out] = audit_out_view[obs_col_out].astype(str).replace({"nan": ""}).fillna("")
                        col_cfg_out[obs_col_out] = cc.TextColumn(
                            label=obs_col_out,
                            help="Digite livremente; será salvo no Google Sheets.",
                            disabled=False,
                        )
        
                    with st.form("form_editar_desc_removidos", clear_on_submit=False):
                        edited_out_view = st.data_editor(
                            audit_out_view,
                            use_container_width=True,
                            hide_index=True,
                            height=320,
                            column_config=col_cfg_out,
                            key=f"editor_removidos_desc_{st.session_state['rev_desc']}",
                        )
                        c_save_out, _ = st.columns([1, 6])
                        salvar_out = c_save_out.form_submit_button("Atualizar Google Sheets")
        
                    if salvar_out:
                        try:
                            antes_desc  = audit_out_raw.get("Descrição Agrupada", pd.Series("", index=audit_out_raw.index)).astype(str).fillna("").reset_index(drop=True)
                            depois_desc = edited_out_view.get("Descrição Agrupada", pd.Series("", index=audit_out_raw.index)).astype(str).fillna("").reset_index(drop=True)
        
                            if obs_col_out:
                                antes_obs  = audit_out_raw.get(obs_col_out, pd.Series("", index=audit_out_raw.index)).astype(str).fillna("").reset_index(drop=True)
                                depois_obs = edited_out_view.get(obs_col_out, pd.Series("", index=audit_out_raw.index)).astype(str).fillna("").reset_index(drop=True)
                            else:
                                antes_obs = depois_obs = pd.Series("", index=antes_desc.index)
        
                            mask_changed = (antes_desc != depois_desc) | (antes_obs != depois_obs)
                            if not mask_changed.any():
                                st.success("Nada para atualizar — nenhuma alteração em removidos.")
                            else:
                                ws_sys = planilha_empresa.worksheet(WS_SISTEMA)
                                df_ws, col_map = _sheet_df_with_row(ws_sys)
                                if "Duplicidade" not in df_ws.columns:
                                    st.error("A aba do Sheets precisa ter a coluna 'Duplicidade'.")
                                else:
                                    df_ws["Duplicidade"] = df_ws["Duplicidade"].astype(str)
                                    col_idx_desc = col_map.get("Descrição Agrupada")
                                    obs_sheet_name = _find_col(col_map.keys())
                                    col_idx_obs = col_map.get(obs_sheet_name) if obs_sheet_name else None
        
                                    keys = audit_out_raw["Duplicidade"].astype(str).reset_index(drop=True)
        
                                    for i in mask_changed[mask_changed].index:
                                        dup_key   = keys.iloc[i]
                                        nova_desc = depois_desc.iloc[i].strip()
                                        nova_obs  = depois_obs.iloc[i].strip()
        
                                        hits = df_ws.index[df_ws["Duplicidade"] == dup_key].tolist()
                                        if not hits:
                                            continue
        
                                        updates = []
                                        for h in hits:
                                            row_num = int(df_ws.loc[h, "_row"])
                                            if Cell is not None:
                                                if col_idx_desc:
                                                    updates.append(Cell(row=row_num, col=col_idx_desc, value=nova_desc))
                                                if col_idx_obs is not None:
                                                    updates.append(Cell(row=row_num, col=col_idx_obs,  value=nova_obs))
                                            else:
                                                if col_idx_desc:
                                                    updates.append((row_num, col_idx_desc, nova_desc))
                                                if col_idx_obs is not None:
                                                    updates.append((row_num, col_idx_obs,  nova_obs))
                                        if updates:
                                            if Cell is not None:
                                                ws_sys.update_cells(updates, value_input_option="USER_ENTERED")
                                            else:
                                                for row_num, col_idx, value in updates:
                                                    a1 = f"{_excel_col_letter(col_idx-1)}{row_num}"
                                                    ws_sys.update(a1, value, value_input_option="USER_ENTERED")
        
                                st.success("Alterações salvas (removidos).")
                                st.session_state["rev_desc"] += 1
                                st.rerun()
        
                        except Exception as e:
                            st.error(f"Falha ao atualizar (removidos): {type(e).__name__}: {e}")
        
                # ====================== segue o fluxo normal usando apenas os incluídos ======================
                base = base.loc[~mask_dep_sys].copy()
        
                # ====================== agrega Sistema (sem depósitos) ======================
                df_sys = (
                    base.groupby(["Código Everest","Data"], as_index=False)[col_valor]
                        .sum()
                        .rename(columns={col_valor:"Sangria (Colibri/CISS)"})
                )
        
                # ====================== Everest ======================
                ws_ev = planilha_empresa.worksheet("Sangria Everest")
                df_ev = pd.DataFrame(ws_ev.get_all_records())
                df_ev.columns = [c.strip() for c in df_ev.columns]
        
                def _norm(s): return re.sub(r"[^a-z0-9]", "", str(s).lower())
                cmap = {_norm(c): c for c in df_ev.columns}
                col_emp   = cmap.get("empresa")
                # ✅ PRIORIDADE: D. Competência → fallback para D. Lançamento/Data
                pref_comp      = ["dcompetencia", "datacompetencia", "datadecompetencia", "competencia", "dtcompetencia"]
                fallback_lcto  = ["dlancamento", "dlancament", "dlanamento", "datadelancamento", "data"]
        
                col_dt_ev = next((cmap[k] for k in pref_comp if k in cmap), None)
                if col_dt_ev is None:
                    col_dt_ev = next((cmap[k] for k in fallback_lcto if k in cmap), None)
        
                col_val_ev= next((orig for norm, orig in cmap.items()
                                  if norm in ("valorlancamento","valorlancament","valorlcto","valor")), None)
                col_fant  = next((orig for norm, orig in cmap.items()
                                  if norm in ("fantasiaempresa","fantasia")), None)
        
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
        
                    # ========== mapeamento Loja/Grupo (garante 1 loja por Código Everest) ==========
                    mapa = df_empresa.copy()
                    mapa.columns = [str(c).strip() for c in mapa.columns]
        
                    if "Código Everest" in mapa.columns:
                        mapa["Código Everest"] = mapa["Código Everest"].astype(str).str.extract(r"(\d+)")
                        mapa["__prio__"] = mapa["Loja"].astype(str).str.contains(r"(embarque|checkin)", case=False, na=False).astype(int)
                        mapa_unico = (
                            mapa.sort_values(["Código Everest", "__prio__", "Loja"])
                                .drop_duplicates(subset=["Código Everest"], keep="first")
                                [["Código Everest", "Loja", "Grupo"]]
                        )
                        cmp = cmp.merge(mapa_unico, on="Código Everest", how="left")
        
                    # fallback LOJA = Fantasia (linhas apenas do Everest)
                    cmp["Loja"] = cmp["Loja"].astype(str)
                    so_everest = (cmp["_merge"] == "right_only") & (cmp["Loja"].isin(["", "nan"]))
                    cmp.loc[so_everest, "Loja"] = cmp.loc[so_everest, "Fantasia Everest"]
                    cmp["Nao Mapeada?"] = so_everest
        
                    # ====================== diferença ======================
                    cmp["Diferença"] = cmp["Sangria (Colibri/CISS)"] - cmp["Sangria Everest"]
        
                    # ====================== filtro Diferenças/Sem diferença ======================
                    cmp["Diferença"] = pd.to_numeric(cmp["Diferença"], errors="coerce").fillna(0.0)
                    TOL = 0.0099
                    eh_zero = np.isclose(cmp["Diferença"].to_numpy(dtype=float), 0.0, atol=TOL)
        
                    if filtro_dif == "Diferenças":
                        cmp = cmp[~eh_zero]
                    elif filtro_dif == "Sem diferença":
                        cmp = cmp[eh_zero]
        
                    if grupos_sel:
                        cmp = cmp[cmp["Grupo"].astype(str).isin(grupos_sel)]
        
                    cmp = cmp[["Grupo","Loja","Código Everest","Data",
                               "Sangria (Colibri/CISS)","Sangria Everest","Diferença","Nao Mapeada?"]
                             ].sort_values(["Grupo","Loja","Código Everest","Data"])
        
                    if filtro_dif == "Diferenças":
                        st.caption("Mostrando apenas linhas com diferença (|Diferença| > R$ 0,01).")
                    elif filtro_dif == "Sem diferença":
                        st.caption("Mostrando apenas linhas sem diferença (|Diferença| ≤ R$ 0,01).")
        
                    total = {
                        "Grupo":"TOTAL","Loja":"","Código Everest":"","Data":pd.NaT,
                        "Sangria (Colibri/CISS)": cmp["Sangria (Colibri/CISS)"].sum(),
                        "Sangria Everest":        cmp["Sangria Everest"].sum(),
                        "Diferença":              cmp["Diferença"].sum(),
                        "Nao Mapeada?": False
                    }
                    df_exibe = pd.concat([pd.DataFrame([total]), cmp], ignore_index=True)
        
                    # ====================== BOTÕES Selecionar/Limpar + TABELA (depois dos depósitos) ======================
                    with st.form("form_selecao_codigos", clear_on_submit=False):
                        c_sel, c_limpar, _ = st.columns([0.8, 0.8, 8], gap="small")
                        aplicar = c_sel.form_submit_button("Selecionar", help="Aplicar o filtro pelos códigos marcados na tabela")
                        limpar  = c_limpar.form_submit_button("Limpar", help="Remover o filtro aplicado e desmarcar tudo")
        
                        df_show = df_exibe.copy()
                        df_show["Data"] = pd.to_datetime(df_show["Data"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
                        for c in ["Sangria (Colibri/CISS)","Sangria Everest","Diferença"]:
                            df_show[c] = df_show[c].apply(
                                lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
                                if isinstance(v,(int,float)) else v
                            )
        
                        view = df_show.drop(columns=["Nao Mapeada?"], errors="ignore").copy()
                        insert_pos = (list(view.columns).index("Diferença") + 1) if "Diferença" in view.columns else len(view.columns)
                        if "Selecionado" not in view.columns:
                            view.insert(insert_pos, "Selecionado", False)
        
                        col_cfg = {}
                        for col in view.columns:
                            if col == "Selecionado":
                                col_cfg[col] = cc.CheckboxColumn(
                                    label="Selecionado",
                                    help="Marque as linhas e depois clique em ✅ Selecionar.",
                                    default=False
                                )
                            elif col == "Data":
                                col_cfg[col] = cc.TextColumn(label=col, disabled=True)
                            elif col in ("Sangria (Colibri/CISS)","Sangria Everest","Diferença"):
                                col_cfg[col] = cc.TextColumn(label=col, disabled=True)
                            else:
                                col_cfg[col] = cc.TextColumn(label=col, disabled=True)
        
                        # Pré-marcar pelos códigos atualmente aplicados (não recarrega)
                        if tem_filtro_codigo and {"Código Everest","Grupo"}.issubset(set(view.columns)):
                            cod_series = view["Código Everest"].astype(str).str.extract(r"(\d+)")[0]
                            mask_normais = view["Grupo"].astype(str).str.upper() != "TOTAL"
                            view.loc[mask_normais, "Selecionado"] = cod_series[mask_normais].isin(codigos_aplicados).values
        
                        edited_view = st.data_editor(
                            view,
                            use_container_width=True,
                            hide_index=True,
                            height=520,
                            column_config=col_cfg,
                            key="cmp_editor_com_checkbox",
                        )
        
                    # ===== AÇÃO PÓS-SUBMIT =====
                    if aplicar:
                        try:
                            sel_mask = (edited_view["Selecionado"] == True) & (edited_view["Grupo"].astype(str).str.upper() != "TOTAL")
                            sel_codigos = (
                                edited_view.loc[sel_mask, "Código Everest"]
                                .astype(str).str.extract(r"(\d+)")[0]
                                .dropna().tolist()
                            )
                            st.session_state["cmp_codigos_selecionados"] = set(sel_codigos)
                        except Exception:
                            st.session_state["cmp_codigos_selecionados"] = set()
                        st.rerun()  # aplica o filtro nos expanders só agora
        
                    if limpar:
                        st.session_state["cmp_codigos_selecionados"] = set()
                        st.rerun()
        
                    # ====================== EXPORTAÇÃO (com slicers quando possível) ======================
                    from io import BytesIO
                    import os
        
                    def _prep_df_export(cmp: pd.DataFrame, usar_mes_sem_acento: bool = False) -> pd.DataFrame:
                        df = cmp.copy()
                        df = df.drop(columns=["Nao Mapeada?"], errors="ignore")
                        if "Sangria (Sistema)" in df.columns:
                            df = df.rename(columns={"Sangria (Sistema)":"Sangria (Colibri/CISS)"})
        
                        df["Data"] = pd.to_datetime(df["Data"], errors="coerce").dt.normalize()
                        df["Ano"]  = df["Data"].dt.year
                        df["Mês"]  = df["Data"].dt.month
        
                        for c in ["Sangria (Colibri/CISS)","Sangria Everest","Diferença"]:
                            if c in df.columns:
                                df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
        
                        ordem = ["Data","Grupo","Loja","Código Everest",
                                 "Sangria (Colibri/CISS)","Sangria Everest","Diferença",
                                 "Mês","Ano"]
                        df = df[[c for c in ordem if c in df.columns]].copy()
        
                        if usar_mes_sem_acento and "Mês" in df.columns:
                            df = df.rename(columns={"Mês":"Mes"})
                        return df
        
                    
                    def exportar_xlsxwriter_tentando_slicers(cmp: pd.DataFrame, usar_mes_sem_acento: bool=False) -> tuple[BytesIO,bool]:
                        df = _prep_df_export(cmp, usar_mes_sem_acento=usar_mes_sem_acento)
                        try:
                            import xlsxwriter as xw
                            ver_tuple = tuple(int(p) for p in xw.__version__.split(".")[:3])
                        except Exception:
                            ver_tuple = (0,0,0)
                    
                        from xlsxwriter import Workbook
                        buf = BytesIO()
                        wb  = Workbook(buf, {"in_memory": True})
                        ws  = wb.add_worksheet("Dados")
                    
                        fmt_header = wb.add_format({"bold":True,"align":"center","valign":"vcenter","bg_color":"#F2F2F2","border":1})
                        fmt_text   = wb.add_format({"border":1})
                        fmt_int    = wb.add_format({"border":1,"num_format":"0"})
                        fmt_date   = wb.add_format({"border":1,"num_format":"dd/mm/yyyy"})
                        fmt_money  = wb.add_format({"border":1,"num_format":'R$ #,##0.00'})
                    
                        headers = list(df.columns)
                        for j,c in enumerate(headers):
                            ws.write(0,j,c,fmt_header)
                    
                        for i,row in df.iterrows():
                            r = i+1
                            for j,c in enumerate(headers):
                                v = row[c]
                                if c=="Data" and pd.notna(v):
                                    ws.write_datetime(r,j,pd.to_datetime(v).to_pydatetime(),fmt_date)
                                elif c in ("Ano","Mês","Mes","Código Everest"):
                                    ws.write_number(r,j,int(v) if pd.notna(v) else 0,fmt_int)
                                elif c in ("Sangria (Colibri/CISS)","Sangria Everest","Diferença"):
                                    ws.write_number(r,j,float(v),fmt_money)
                                else:
                                    ws.write(r,j,("" if pd.isna(v) else v),fmt_text)
                    
                        last_row = len(df)
                        last_col = len(headers)-1
                        ws.add_table(0,0,last_row,last_col,{
                            "name":"tbl_dados",
                            "style":"TableStyleMedium9",
                            "columns":[{"header":h} for h in headers],
                        })
                    
                        col_idx = {c:i for i,c in enumerate(headers)}
                        if "Data" in col_idx:           ws.set_column(col_idx["Data"], col_idx["Data"], 12, fmt_date)
                        if "Grupo" in col_idx:          ws.set_column(col_idx["Grupo"],col_idx["Grupo"],10,fmt_text)
                        if "Loja" in col_idx:           ws.set_column(col_idx["Loja"], col_idx["Loja"], 28,fmt_text)
                        if "Código Everest" in col_idx: ws.set_column(col_idx["Código Everest"],col_idx["Código Everest"],14,fmt_int)
                        for c in ("Sangria (Colibri/CISS)","Sangria Everest","Diferença"):
                            if c in col_idx:            ws.set_column(col_idx[c],col_idx[c],18,fmt_money)
                        if "Mês" in col_idx:            ws.set_column(col_idx["Mês"],6,6,fmt_int)
                        if "Mes" in col_idx:            ws.set_column(col_idx["Mes"],6,6,fmt_int)
                        if "Ano" in col_idx:            ws.set_column(col_idx["Ano"],8,8,fmt_int)
                        ws.freeze_panes(1,0)
                    
                        slicers_ok = False
                        if ver_tuple >= (3,2,0) and hasattr(wb, "add_slicer"):
                            try:
                                col_mes = "Mes" if ("Mes" in headers) else ("Mês" if "Mês" in headers else None)
                                if "Ano" in headers:
                                    wb.add_slicer({"table":"tbl_dados","column":"Ano","cell":"L2","width":130,"height":100})
                                if col_mes:
                                    wb.add_slicer({"table":"tbl_dados","column":col_mes,"cell":"L8","width":130,"height":130})
                                if "Grupo" in headers:
                                    wb.add_slicer({"table":"tbl_dados","column":"Grupo","cell":"N2","width":180,"height":180})
                                if "Loja" in headers:
                                    wb.add_slicer({"table":"tbl_dados","column":"Loja","cell":"N12","width":260,"height":320})
                                slicers_ok = True
                            except Exception:
                                slicers_ok = False  # silencioso
                    
                        wb.close()
                        buf.seek(0)
                        return buf, slicers_ok
        
                    xlsx_out, ok = exportar_xlsxwriter_tentando_slicers(cmp, usar_mes_sem_acento=True)
                    if not ok:
                        try:
                            xlsx_out = exportar_via_template_preservando_slicers(cmp)
                        except Exception:
                            pass  # segue sem mensagens
        
                    st.download_button(
                        label="⬇️ Baixar Excel",
                        data=xlsx_out,
                        file_name="Sangria_Controle.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_sangria_controle_excel"
                    )
