# -------------------------------
# Sub-aba: 🧰 CONTROLE DE SANGRIA (Comparativa Everest / Diferenças)
# -------------------------------
with sub_caixa:
    if df_sangria is None or df_sangria.empty:
        st.info("Sem dados de **sangria** disponíveis.")
    else:
        import pandas as pd, numpy as np, re, unicodedata
        from io import BytesIO

        # -------- helpers --------
        def _norm_txt(s: str) -> str:
            s = str(s or "").strip().lower()
            return unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("utf-8")

        def eh_deposito_mask(df, cols_texto=None):
            if cols_texto is None:
                cols_texto = ["Descrição Agrupada","Descrição","Historico","Histórico",
                              "Categoria","Obs","Observação","Tipo","Tipo Movimento"]
            cols_texto = [c for c in cols_texto if c in df.columns]
            if not cols_texto:
                return pd.Series(False, index=df.index)
            txt = df[cols_texto].astype(str).agg(" ".join, axis=1).map(_norm_txt)
            padrao = r"""
                \bdeposito\b | \bdepsito\b | \bdep\b |
                credito\s+em\s+conta | envio\s*para\s*banco |
                transf(erencia)?\s*(p/?\s*banco|banco)
            """
            return txt.str.contains(padrao, flags=re.IGNORECASE|re.VERBOSE, regex=True, na=False)

        def brl(v):
            try:   return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            except: return "R$ 0,00"

        # -------- base --------
        df = df_sangria.copy()
        df.columns = [str(c).strip() for c in df.columns]
        if "Data" not in df.columns:
            st.error("A aba 'Sangria' precisa da coluna **Data**."); st.stop()
        df["Data"] = pd.to_datetime(df["Data"], errors="coerce", dayfirst=True)

        col_valor = pick_valor_col(df.columns)
        if not col_valor:
            st.error("Não encontrei a coluna de **valor** (ex.: 'Valor(R$)')."); st.stop()
        df[col_valor] = df[col_valor].map(parse_valor_brl_sheets).astype(float)

        # -------- filtros --------
        c1, c2, c3, c4, c5 = st.columns([1.2,1.2,1.2,1.2,1.2])
        with c2:
            grupos_df  = sorted(df.get("Grupo", pd.Series([],dtype=str)).dropna().astype(str).unique().tolist())
            grupos_emp = sorted(df_empresa.get("Grupo", pd.Series([],dtype=str)).dropna().astype(str).unique().tolist())
            grupos_sel = st.multiselect("Grupos", options=sorted({*grupos_df,*grupos_emp}), key="caixa_grupos_cmp")
        with c1:
            dmin = pd.to_datetime(df["Data"].min(), errors="coerce"); dmax = pd.to_datetime(df["Data"].max(), errors="coerce")
            today = pd.Timestamp.today().normalize()
            if pd.isna(dmin): dmin = today
            if pd.isna(dmax): dmax = today
            dt_inicio, dt_fim = st.date_input("Período",
                value=(dmax.date(), dmax.date()),
                min_value=dmin.date(),
                max_value=(dmax.date() if dmax>=dmin else dmin.date()),
                key="caixa_periodo_cmp")
        with c3:
            df_opt = df[(df["Data"].dt.date>=dt_inicio)&(df["Data"].dt.date<=dt_fim)].copy()
            if grupos_sel and "Grupo" in df_opt.columns:
                df_opt = df_opt[df_opt["Grupo"].astype(str).isin(grupos_sel)]
            opcoes_lojas = sorted(df_opt.get("Loja", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
            prev = st.session_state.get("caixa_lojas_cmp", [])
            lojas_sel = st.multiselect("Lojas", options=opcoes_lojas, default=[x for x in prev if x in opcoes_lojas], key="caixa_lojas_cmp")
        with c4:
            visao = st.selectbox("Visão do Relatório", ["Comparativa Everest"], key="caixa_visao_cmp")
        with c5:
            filtro_dif = st.selectbox("Filtro por Diferença", ["Todas","Diferenças","Sem diferença"], key="caixa_filtro_diferenca")

        df_fil = df[(df["Data"].dt.date>=dt_inicio)&(df["Data"].dt.date<=dt_fim)].copy()
        if lojas_sel: df_fil = df_fil[df_fil["Loja"].astype(str).isin(lojas_sel)]
        if grupos_sel and "Grupo" in df_fil.columns: df_fil = df_fil[df_fil["Grupo"].astype(str).isin(grupos_sel)]

        if visao == "Comparativa Everest":
            base = df_fil.copy()
            base["Data"] = pd.to_datetime(base["Data"], dayfirst=True, errors="coerce").dt.normalize()
            base[col_valor] = pd.to_numeric(base[col_valor], errors="coerce").fillna(0.0)
            if "Código Everest" not in base.columns:
                st.error("❌ Falta a coluna **Código Everest** na aba Sangria."); st.stop()
            base["Código Everest"] = base["Código Everest"].astype(str).str.extract(r"(\d+)")

            # --- listas Excluídos / Incluídos ---
            base_raw = base.copy()
            mask_extra = (base["Descrição Agrupada"].astype(str).str.contains(r"\b(maionese|Moeda Estrangeira)\b", regex=True, na=False)
                          if "Descrição Agrupada" in base.columns else pd.Series(False, index=base.index))
            mask_dep_sys = eh_deposito_mask(base) | mask_extra

            with st.expander("🔎 Ver depósitos/termos removidos (Colibri/CISS)"):
                audit = base.loc[mask_dep_sys].copy()
                if "Data" in audit.columns:
                    audit["Data"] = pd.to_datetime(audit["Data"], errors="coerce").dt.strftime("%d/%m/%Y")
                if col_valor in audit.columns:
                    audit[col_valor] = audit[col_valor].map(brl)
                st.dataframe(audit, use_container_width=True, hide_index=True, height=260)

            # INCLUÍDOS (mostra Descrição e Descrição Agrupada; sem checkboxes aqui)
            inc = base_raw.loc[~mask_dep_sys].copy()
            col_desc = "Descrição" if "Descrição" in inc.columns else None
            col_desc_agr = "Descrição Agrupada" if "Descrição Agrupada" in inc.columns else None
            cols_inc = ["Grupo","Loja","Código Everest","Data"]
            if col_desc: cols_inc.append(col_desc)
            if col_desc_agr: cols_inc.append(col_desc_agr)
            if col_valor in inc.columns: cols_inc.append(col_valor)
            cols_inc = [c for c in cols_inc if c in inc.columns]
            if "Data" in inc.columns:
                inc["Data"] = pd.to_datetime(inc["Data"], errors="coerce").dt.strftime("%d/%m/%Y")
            if col_valor in inc.columns:
                inc[col_valor] = inc[col_valor].map(brl)
            with st.expander("🧾 INCLUÍDOS (lado Colibri/CISS)"):
                st.dataframe(inc[cols_inc], use_container_width=True, hide_index=True, height=260)

            # --- agrega Sistema (sem depósitos) ---
            base = base_raw.loc[~mask_dep_sys].copy()
            df_sys = (base.groupby(["Código Everest","Data"], as_index=False)[col_valor]
                          .sum()
                          .rename(columns={col_valor:"Sangria (Colibri/CISS)"}))

            # --- Everest ---
            ws_ev = planilha_empresa.worksheet("Sangria Everest")
            df_ev = pd.DataFrame(ws_ev.get_all_records()); df_ev.columns=[c.strip() for c in df_ev.columns]
            def _n(s): return re.sub(r"[^a-z0-9]","",str(s).lower())
            cmap = {_n(c):c for c in df_ev.columns}
            col_emp = cmap.get("empresa")
            pref_comp = ["dcompetencia","datacompetencia","datadecompetencia","competencia","dtcompetencia"]
            fallback  = ["dlancamento","datadelancamento","data"]
            col_dt_ev = next((cmap[k] for k in pref_comp if k in cmap), next((cmap[k] for k in fallback if k in cmap), None))
            col_val_ev= next((orig for k,orig in cmap.items() if k in ("valorlancamento","valorlcto","valor")), None)
            col_fant  = next((orig for k,orig in cmap.items() if k in ("fantasiaempresa","fantasia")), None)
            if not all([col_emp,col_dt_ev,col_val_ev]):
                st.error("❌ Na 'Sangria Everest' preciso de 'Empresa', 'D. Competência' (ou 'D. Lançamento') e 'Valor Lancamento'."); st.stop()

            de = df_ev.copy()
            de["Código Everest"]   = de[col_emp].astype(str).str.extract(r"(\d+)")
            de["Fantasia Everest"] = de[col_fant] if col_fant else ""
            de["Data"]             = pd.to_datetime(de[col_dt_ev], dayfirst=True, errors="coerce").dt.normalize()
            de["Valor Lancamento"] = de[col_val_ev].map(parse_valor_brl_sheets).astype(float)
            de = de[(de["Data"].dt.date>=dt_inicio)&(de["Data"].dt.date<=dt_fim)]
            de["Sangria Everest"]  = de["Valor Lancamento"].abs()
            def _pick_first(s):
                s = s.dropna().astype(str).str.strip(); s = s[s!=""]; return s.iloc[0] if not s.empty else ""
            de_agg = (de.groupby(["Código Everest","Data"], as_index=False)
                        .agg({"Sangria Everest":"sum","Fantasia Everest":_pick_first}))

            cmp = df_sys.merge(de_agg, on=["Código Everest","Data"], how="outer", indicator=True)
            cmp["Sangria (Colibri/CISS)"] = cmp["Sangria (Colibri/CISS)"].fillna(0.0)
            cmp["Sangria Everest"]        = cmp["Sangria Everest"].fillna(0.0)

            # mapeamento Loja/Grupo (1 por Código Everest)
            mapa = df_empresa.copy(); mapa.columns=[str(c).strip() for c in mapa.columns]
            if "Código Everest" in mapa.columns:
                mapa["Código Everest"] = mapa["Código Everest"].astype(str).str.extract(r"(\d+)")
                mapa["__prio__"] = mapa["Loja"].astype(str).str.contains(r"(embarque|checkin)", case=False, na=False).astype(int)
                mapa_unico = (mapa.sort_values(["Código Everest","__prio__","Loja"])
                                   .drop_duplicates(subset=["Código Everest"], keep="first")[["Código Everest","Loja","Grupo"]])
                cmp = cmp.merge(mapa_unico, on="Código Everest", how="left")

            # fallback loja
            cmp["Loja"] = cmp["Loja"].astype(str)
            so_everest = (cmp["_merge"]=="right_only") & (cmp["Loja"].isin(["","nan"]))
            cmp.loc[so_everest,"Loja"] = cmp.loc[so_everest,"Fantasia Everest"]
            cmp["Nao Mapeada?"] = so_everest

            # diferença + filtro
            cmp["Diferença"] = pd.to_numeric(cmp["Sangria (Colibri/CISS)"] - cmp["Sangria Everest"], errors="coerce").fillna(0.0)
            TOL = 0.0099
            eh_zero = np.isclose(cmp["Diferença"].to_numpy(dtype=float), 0.0, atol=TOL)
            if filtro_dif == "Diferenças":
                cmp = cmp[~eh_zero]; st.caption("Mostrando apenas linhas com diferença (|Diferença| > R$ 0,01).")
            elif filtro_dif == "Sem diferença":
                cmp = cmp[eh_zero]; st.caption("Mostrando apenas linhas sem diferença (|Diferença| ≤ R$ 0,01).")
            if grupos_sel: cmp = cmp[cmp["Grupo"].astype(str).isin(grupos_sel)]

            cmp = cmp[["Grupo","Loja","Código Everest","Data",
                       "Sangria (Colibri/CISS)","Sangria Everest","Diferença","Nao Mapeada?"]
                     ].sort_values(["Grupo","Loja","Código Everest","Data"])

            # ---------- linha TOTAL + RENDER COM CHECKBOX NA PRÓPRIA TABELA ----------
            total = {
                "Grupo":"TOTAL","Loja":"","Código Everest":"","Data":pd.NaT,
                "Sangria (Colibri/CISS)": cmp["Sangria (Colibri/CISS)"].sum(),
                "Sangria Everest":        cmp["Sangria Everest"].sum(),
                "Diferença":              cmp["Diferença"].sum(),
                "Nao Mapeada?": False
            }
            df_exibe = pd.concat([pd.DataFrame([total]), cmp], ignore_index=True)

            df_show = df_exibe.copy()
            df_show["Data"] = pd.to_datetime(df_show["Data"], errors="coerce").dt.strftime("%d/%m/%Y").fillna("")
            for c in ["Sangria (Colibri/CISS)","Sangria Everest","Diferença"]:
                df_show[c] = df_show[c].apply(lambda v: f"R$ {v:,.2f}".replace(",", "X").replace(".", ",").replace("X",".")
                                              if isinstance(v,(int,float)) else v)

            # remove col técnica para exibir, mas preserva para estilizar se quiser
            df_to_show = df_show.drop(columns=["Nao Mapeada?"], errors="ignore").copy()

            # NÃO criar segunda tabela: apenas adiciona a coluna de seleção
            is_total = df_to_show["Grupo"].astype(str).eq("TOTAL") if "Grupo" in df_to_show.columns else pd.Series(False, index=df_to_show.index)
            df_total = df_to_show[is_total].copy()
            df_main  = df_to_show[~is_total].copy()
            df_main["✅ Selecionar"] = False  # só nas linhas normais

            # editor único com checkbox
            edited = st.data_editor(
                pd.concat([df_total, df_main], ignore_index=True),
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                column_config={
                    "✅ Selecionar": st.column_config.CheckboxColumn("✅ Selecionar", help="Marque livremente (sem filtros).")
                },
                key="cmp_editor_checks",
                height=520
            )

            # ---------- (restante: export) ----------
            def _prep_df_export(cmp_df: pd.DataFrame, usar_mes_sem_acento: bool=False) -> pd.DataFrame:
                d = cmp_df.copy()
                d = d.drop(columns=["Nao Mapeada?"], errors="ignore")
                if "Sangria (Sistema)" in d.columns:
                    d = d.rename(columns={"Sangria (Sistema)":"Sangria (Colibri/CISS)"})
                d["Data"] = pd.to_datetime(d["Data"], errors="coerce").dt.normalize()
                d["Ano"]  = d["Data"].dt.year; d["Mês"] = d["Data"].dt.month
                for c in ["Sangria (Colibri/CISS)","Sangria Everest","Diferença"]:
                    if c in d.columns: d[c] = pd.to_numeric(d[c], errors="coerce").fillna(0.0)
                ordem = ["Data","Grupo","Loja","Código Everest","Sangria (Colibri/CISS)","Sangria Everest","Diferença","Mês","Ano"]
                d = d[[c for c in ordem if c in d.columns]]
                if usar_mes_sem_acento and "Mês" in d.columns: d = d.rename(columns={"Mês":"Mes"})
                return d

            def exportar_xlsxwriter_tentando_slicers(cmp_df: pd.DataFrame, usar_mes_sem_acento: bool=False):
                d = _prep_df_export(cmp_df, usar_mes_sem_acento)
                from xlsxwriter import Workbook
                buf = BytesIO(); wb = Workbook(buf, {"in_memory": True}); ws = wb.add_worksheet("Dados")
                fmt_header = wb.add_format({"bold":True,"align":"center","valign":"vcenter","bg_color":"#F2F2F2","border":1})
                fmt_text = wb.add_format({"border":1}); fmt_int = wb.add_format({"border":1,"num_format":"0"})
                fmt_date = wb.add_format({"border":1,"num_format":"dd/mm/yyyy"})
                fmt_money= wb.add_format({"border":1,"num_format":"R$ #,##0.00"})
                headers = list(d.columns)
                for j,c in enumerate(headers): ws.write(0,j,c,fmt_header)
                for i,row in d.iterrows():
                    r=i+1
                    for j,c in enumerate(headers):
                        v=row[c]
                        if c=="Data" and pd.notna(v): ws.write_datetime(r,j,pd.to_datetime(v).to_pydatetime(),fmt_date)
                        elif c in ("Ano","Mês","Mes","Código Everest"): ws.write_number(r,j,int(v) if pd.notna(v) else 0,fmt_int)
                        elif c in ("Sangria (Colibri/CISS)","Sangria Everest","Diferença"): ws.write_number(r,j,float(v),fmt_money)
                        else: ws.write(r,j,("" if pd.isna(v) else v),fmt_text)
                last_row=len(d); last_col=len(headers)-1
                ws.add_table(0,0,last_row,last_col,{"name":"tbl_dados","style":"TableStyleMedium9","columns":[{"header":h} for h in headers]})
                wb.close(); buf.seek(0); return buf

            xlsx_out = exportar_xlsxwriter_tentando_slicers(cmp, usar_mes_sem_acento=True)
            st.download_button("⬇️ Baixar Excel", data=xlsx_out, file_name="Sangria_Controle.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               key="dl_sangria_controle_excel")
