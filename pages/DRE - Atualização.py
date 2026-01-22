# -----------------------------
# ABA: AUDITORIA (Lógica Intacta)
# -----------------------------
with tab_audit:
    st.subheader("Auditoria Faturamento X Meio de Pagamento")
    try:
        pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
        if not pastas_fech:
            st.error("Nenhuma pasta de fechamento encontrada.")
            st.stop()
        map_p = {p["name"]: p["id"] for p in pastas_fech}
        p_sel = st.selectbox("Pasta principal:", options=list(map_p.keys()), key="au_p")
        subpastas = list_child_folders(drive_service, map_p[p_sel])
        map_s = {s["name"]: s["id"] for s in subpastas}
        s_sel = st.multiselect("Subpastas (se nenhuma, trará todas):", options=list(map_s.keys()), default=[], key="au_s")
        s_ids_audit = [map_s[n] for n in s_sel] if s_sel else list(map_s.values())
    except Exception as e:
        st.error(f"Erro ao listar pastas/subpastas: {e}")
        st.stop()

    c1, c2 = st.columns(2)
    with c1:
        ano_sel = st.selectbox("Ano:", list(range(2020, date.today().year + 1)), index=max(0, date.today().year - 2020), key="au_ano")
    with c2:
        mes_sel = st.selectbox("Mês (Opcional):", ["Todos"] + list(range(1, 13)), key="au_mes")

    need_reload = ("au_last_subpastas" not in st.session_state) or (st.session_state.get("au_last_subpastas") != s_ids_audit)
    if need_reload:
        try:
            planilhas = list_spreadsheets_in_folders(drive_service, s_ids_audit)
        except Exception as e:
            st.error(f"Erro ao listar planilhas: {e}")
            st.stop()

        df_init = pd.DataFrame([{
            "Planilha": p["name"],
            "Flag": False,
            "Planilha_id": p["id"],
            "Origem": "",
            "DRE": "",
            "MP DRE": "",
            "Dif": "",
            "Dif MP": "",
            "Status": ""
        } for p in planilhas])

        st.session_state.au_last_subpastas = s_ids_audit
        st.session_state.au_planilhas_df = df_init
        st.session_state.au_resultados = {}
        st.session_state.au_flags_temp = {}

    if "au_planilhas_df" not in st.session_state:
        st.session_state.au_planilhas_df = pd.DataFrame(columns=["Planilha", "Flag", "Planilha_id", "Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"])
    
    df_table = st.session_state.au_planilhas_df.copy()
    if df_table.empty:
        st.info("Nenhuma planilha encontrada.")
    
    expected_cols = ["Planilha", "Planilha_id", "Flag", "Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"]
    for c in expected_cols:
        if c not in df_table.columns:
            df_table[c] = False if c == "Flag" else ("" if c != "Planilha_id" else "")

    display_df = df_table[expected_cols].copy()

    row_style_js = JsCode("""
    function(params) {
        if (params.data && (params.data.Flag === true || params.data.Flag === 'true')) {
            return {'background-color': '#e9f7ee'};
        }
    }
    """)
    gb = GridOptionsBuilder.from_dataframe(display_df)
    gb.configure_column("Planilha", headerName="Planilha", editable=False, width=420)
    gb.configure_column("Planilha_id", headerName="Planilha_id", editable=False, hide=True)
    gb.configure_column("Flag", editable=True, cellEditor="agCheckboxCellEditor", cellRenderer="agCheckboxCellRenderer", width=80)
    for col in ["Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"]:
        if col in display_df.columns:
            gb.configure_column(col, editable=False)
    grid_options = gb.build()
    grid_options['getRowStyle'] = row_style_js

 
    st.markdown('<div id="auditoria">', unsafe_allow_html=True)    
    
    grid_response = AgGrid(
        display_df,
        gridOptions=grid_options,
        update_mode=GridUpdateMode.MODEL_CHANGED,
        allow_unsafe_jscode=True,
        theme='alpine',
        height=420,
        fit_columns_on_grid_load=True,
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    # use the fourth column for the verification button so everything stays aligned
    col_btn1, col_btn2, col_btn3, col_btn4 = st.columns([2, 2, 1, 6])

    with col_btn1:
        executar_clicado = st.button("📊 Atualizar", key="au_exec", use_container_width=True)

    with col_btn2:
        limpar_clicadas = st.button("🧹 Limpar marcadas", key="au_limpar", use_container_width=True)

    currency_cols = ["Origem", "DRE", "MP DRE", "Dif", "Dif MP"]
    cols_for_excel = ["Planilha"] + [c for c in currency_cols if c in st.session_state.au_planilhas_df.columns]
    df_para_excel_btn = st.session_state.au_planilhas_df[cols_for_excel].copy()
    is_empty_btn = df_para_excel_btn.empty

    def _to_numeric_or_nan(x):
        if pd.isna(x) or str(x).strip() == "": return pd.NA
        if isinstance(x, (int, float)): return float(x)
        n = _parse_currency_like(x)
        if n is None:
            try: return float(str(x).replace(".", "").replace(",", "."))
            except: return pd.NA
        return float(n)

    with col_btn3:
        if not is_empty_btn:
            df_to_write = df_para_excel_btn.copy()
            for col in currency_cols:
                if col in df_to_write.columns:
                    df_to_write[col] = df_to_write[col].apply(_to_numeric_or_nan)

            output_btn = io.BytesIO()
            with pd.ExcelWriter(output_btn, engine="xlsxwriter") as writer:
                df_to_write.to_excel(writer, index=False, sheet_name="Auditoria")
                workbook = writer.book
                worksheet = writer.sheets["Auditoria"]
                currency_fmt = workbook.add_format({'num_format': u'R$ #,##0.00'})
                for i, col in enumerate(df_to_write.columns):
                    if col in currency_cols:
                        worksheet.set_column(i, i, 18, currency_fmt)
                    else:
                        worksheet.set_column(i, i, 40)
            processed_btn = output_btn.getvalue()
        else:
            processed_btn = b""

        st.download_button(
            label="⬇️ Excel",
            data=processed_btn,
            file_name=f"auditoria_dre_{date.today()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            disabled=is_empty_btn,
            key="au_download"
        )

    # place the verification button in the 4th column (aligned)
    with col_btn4:
        verificar_btn = st.button("🔎 Verificar Lojas", use_container_width=True, key="au_verif_simple")

    # --- VERIFICAÇÃO DE LOJAS (mantida) ---
    if verificar_btn:
        st.info("Executando verificação — gerando arquivo para download quando concluir...")
        try:
            # --- PASSO 1: Ler Tabela Empresa (Origem) col A (nome) e col C (código) ---
            sh_origem = gc.open_by_key(ID_PLANILHA_ORIGEM_FAT)
            ws_empresa = sh_origem.worksheet("Tabela Empresa")
            dados_empresa = ws_empresa.get_all_values()

            nomes_codigos = []  # lista de tuples (nome, codigo_normalizado)
            for r in dados_empresa[1:]:  # pula cabeçalho
                nome = r[0].strip() if len(r) > 0 and r[0] is not None else ""
                codigo_raw = r[2] if len(r) > 2 else ""
                if str(codigo_raw).strip() != "":
                    cod_norm = normalize_code(codigo_raw)
                    nomes_codigos.append((nome, cod_norm))

            if not nomes_codigos:
                st.error("Nenhum código encontrado na coluna C da aba 'Tabela Empresa'.")
                st.stop()

            codigos_origem = set(c for _, c in nomes_codigos)

            # --- PASSO 2: Varre as planilhas da pasta e coleta todos os códigos em B3/B4/B5 ---
            planilhas_pasta = st.session_state.get("au_planilhas_df", pd.DataFrame()).copy()
            mapa_codigos_nas_planilhas = {}  # {codigo_normalizado: [nomes_das_planilhas]}

            prog = st.progress(0)
            total = len(planilhas_pasta) if not planilhas_pasta.empty else 0

            for i, prow in planilhas_pasta.reset_index(drop=True).iterrows():
                pname = prow.get("Planilha", "Sem Nome")
                sid = prow.get("Planilha_id")
                try:
                    if sid and str(sid).strip() != "":
                        sh_dest = gc.open_by_key(sid)
                        _, b3, b4, b5 = read_codes_from_config_sheet(sh_dest)
                        for val in (b3, b4, b5):
                            if val and str(val).strip() != "":
                                cod_norm = normalize_code(val)
                                mapa_codigos_nas_planilhas.setdefault(cod_norm, []).append(pname)
                except Exception:
                    # ignora falhas em planilhas individuais (não interrompe todo processo)
                    pass
                if total:
                    prog.progress((i + 1) / total)

            # --- PASSO 3: Monta relatório com nome, código e onde foi encontrado ---
            relatorio = []
            for nome, cod in nomes_codigos:
                planilhas_onde_esta = mapa_codigos_nas_planilhas.get(cod, [])
                relatorio.append({
                    "Nome Empresa (Origem)": nome,
                    "Código Loja (Origem)": cod,
                    "Status": "✅ OK" if planilhas_onde_esta else "❌ FALTANDO PLANILHA",
                    "Planilhas Vinculadas": ", ".join(planilhas_onde_esta) if planilhas_onde_esta else "NENHUMA"
                })

            df_relatorio = pd.DataFrame(relatorio)

            # --- PASSO 4: Gera Excel e disponibiliza para download SEM mostrar tabela abaixo ---
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                df_relatorio.to_excel(writer, index=False, sheet_name="Lojas_Faltantes")
                # formatação básica de colunas
                workbook = writer.book
                worksheet = writer.sheets["Lojas_Faltantes"]
                worksheet.set_column(0, 0, 40)  # Nome Empresa
                worksheet.set_column(1, 1, 20)  # Código
                worksheet.set_column(2, 2, 18)  # Status
                worksheet.set_column(3, 3, 60)  # Planilhas Vinculadas

            excel_bytes = buf.getvalue()
            faltam = int((df_relatorio["Status"] == "❌ FALTANDO PLANILHA").sum())
            st.success(f"Verificação concluída — {faltam} lojas sem planilha. Faça o download do relatório abaixo.")
            st.download_button(
                label="⬇️ Baixar Relatório de Lojas Faltantes",
                data=excel_bytes,
                file_name=f"lojas_sem_planilha_{date.today()}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="au_verif_download_simple"
            )

        except Exception as e:
            st.error(f"Erro na verificação: {e}")

    if limpar_clicadas:
        df_grid_now = pd.DataFrame(grid_response.get("data", []))
        planilhas_marcadas = []
        if not df_grid_now.empty and "Planilha" in df_grid_now.columns:
            planilhas_marcadas = df_grid_now[df_grid_now["Flag"].apply(to_bool_like) == True]["Planilha"].tolist()

        if not planilhas_marcadas:
            mask_master = st.session_state.au_planilhas_df["Flag"] == True
            if mask_master.any():
                planilhas_marcadas = st.session_state.au_planilhas_df.loc[mask_master, "Planilha"].tolist()

        if not planilhas_marcadas:
            st.warning("Marque as planilhas primeiro!")
        else:
            mask = st.session_state.au_planilhas_df["Planilha"].isin(planilhas_marcadas)
            for col in ["Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"]:
                st.session_state.au_planilhas_df.loc[mask, col] = ""
            st.session_state.au_planilhas_df.loc[mask, "Flag"] = False
            st.success(f"Dados de {len(planilhas_marcadas)} planilhas limpos.")
            st.experimental_rerun()

    if executar_clicado:
        df_grid = pd.DataFrame(grid_response.get("data", []))
        if df_grid.empty:
            st.warning("Nenhuma linha para processar.")
        else:
            selecionadas = df_grid[df_grid["Flag"].apply(to_bool_like) == True].copy()
            if "Planilha_id" not in selecionadas.columns:
                selecionadas = selecionadas.merge(st.session_state.au_planilhas_df[["Planilha", "Planilha_id"]], on="Planilha", how="left")

            if selecionadas.empty:
                st.warning("Marque ao menos uma planilha.")
            else:
                if mes_sel == "Todos":
                    d_ini, d_fim = date(ano_sel, 1, 1), date(ano_sel, 12, 31)
                else:
                    d_ini = date(ano_sel, int(mes_sel), 1)
                    d_fim = (date(ano_sel, int(mes_sel), 28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)

                try:
                    sh_o_fat = gc.open_by_key(ID_PLANILHA_ORIGEM_FAT)
                    ws_o_fat = sh_o_fat.worksheet(ABA_ORIGEM_FAT)
                    h_o_fat, df_o_fat = get_headers_and_df_raw(ws_o_fat)
                    if not df_o_fat.empty: df_o_fat = tratar_numericos(df_o_fat, h_o_fat)
                    c_dt_o = detect_date_col(h_o_fat)
                    if c_dt_o and not df_o_fat.empty:
                        df_o_fat["_dt"] = pd.to_datetime(df_o_fat[c_dt_o], dayfirst=True, errors="coerce").dt.date
                        df_o_fat_p = df_o_fat[(df_o_fat["_dt"] >= d_ini) & (df_o_fat["_dt"] <= d_fim)].copy()
                    else:
                        df_o_fat_p = df_o_fat.copy()
                except Exception as e:
                    st.error(f"Erro origem fat: {e}"); st.stop()

                total = len(selecionadas)
                prog = st.progress(0)
                logs = []

                for idx, row in selecionadas.reset_index(drop=True).iterrows():
                    sid = row.get("Planilha_id")
                    if not sid:
                        pname = row.get("Planilha")
                        match = st.session_state.au_planilhas_df.loc[st.session_state.au_planilhas_df["Planilha"] == pname, "Planilha_id"]
                        if not match.empty: sid = match.iloc[0]
                    
                    if not sid:
                        logs.append(f"{row.get('Planilha')}: ID não encontrado.")
                        continue

                    pname = row.get("Planilha", "(sem nome)")
                    v_o = v_d = v_mp = 0.0

                    try:
                        sh_d = gc.open_by_key(sid)
                        b2, b3, b4, b5 = read_codes_from_config_sheet(sh_d)
                        if not b2:
                            logs.append(f"{pname}: Sem B2.")
                            continue

                        lojas_audit = []
                        if b3: lojas_audit.append(normalize_code(b3))
                        if b4: lojas_audit.append(normalize_code(b4))
                        if b5: lojas_audit.append(normalize_code(b5))

                        if h_o_fat and len(h_o_fat) > 5 and not df_o_fat_p.empty:
                            col_b2_fat = h_o_fat[5]
                            df_filter = df_o_fat_p[df_o_fat_p[col_b2_fat].astype(str).str.strip() == str(b2).strip()]
                            if lojas_audit and len(h_o_fat) > 3:
                                col_b3_fat = h_o_fat[3]
                                df_filter = df_filter[df_filter[col_b3_fat].apply(normalize_code).isin(lojas_audit)]
                            v_o = float(df_filter[h_o_fat[6]].sum()) if not df_filter.empty else 0.0

                        ws_d = sh_d.worksheet("Importado_Fat")
                        h_d, df_d = get_headers_and_df_raw(ws_d)
                        if not df_d.empty:
                            df_d = tratar_numericos(df_d, h_d)
                            c_dt_d = detect_date_col(h_d) or (h_d[0] if h_d else None)
                            if c_dt_d:
                                df_d["_dt"] = pd.to_datetime(df_d[c_dt_d], dayfirst=True, errors="coerce").dt.date
                                df_d_periodo = df_d[(df_d["_dt"] >= d_ini) & (df_d["_dt"] <= d_fim)]
                                v_d = float(df_d_periodo[h_d[6]].sum()) if len(h_d) > 6 and not df_d_periodo.empty else 0.0

                        
                        try:
                            ws_mp = sh_d.worksheet("Meio de Pagamento")
                            h_mp, df_mp = get_headers_and_df_raw(ws_mp)
                            if not df_mp.empty:
                                df_mp = tratar_numericos(df_mp, h_mp)
    
                            c_dt_mp = (h_mp[0] if h_mp and len(h_mp) > 0 else None)
                            if not c_dt_mp:
                                c_dt_mp = detect_date_col(h_mp)
    
                            if c_dt_mp and not df_mp.empty:
                                df_mp["_dt"] = pd.to_datetime(df_mp[c_dt_mp], dayfirst=True, errors="coerce")
                                if df_mp["_dt"].isna().all():
                                    df_mp["_dt"] = pd.to_datetime(df_mp[c_dt_mp], dayfirst=False, errors="coerce")
                                df_mp["_dt"] = df_mp["_dt"].dt.date
                                df_mp_periodo = df_mp[(df_mp["_dt"] >= d_ini) & (df_mp["_dt"] <= d_fim)]
                            else:
                                df_mp_periodo = df_mp.copy()
    
                            v_mp_calc = 0.0
                            if not df_mp_periodo.empty:
                                col_b2_mp = h_mp[8] if len(h_mp) > 8 else None
                                col_loja_mp = h_mp[6] if len(h_mp) > 6 else None
                                col_val_mp = h_mp[9] if len(h_mp) > 9 else None
    
                                ok_b2 = (col_b2_mp in df_mp_periodo.columns) if col_b2_mp else False
                                ok_loja = (col_loja_mp in df_mp_periodo.columns) if col_loja_mp else False
                                ok_val = (col_val_mp in df_mp_periodo.columns) if col_val_mp else False
    
                                if ok_b2:
                                    b2_norm = normalize_code(b2)
                                    mask = df_mp_periodo[col_b2_mp].apply(normalize_code) == b2_norm
                                    if lojas_audit and ok_loja:
                                        mask &= df_mp_periodo[col_loja_mp].apply(normalize_code).isin(lojas_audit)
    
                                    df_mp_dest_f = df_mp_periodo[mask]
                                    if not df_mp_dest_f.empty and ok_val:
                                        v_mp_calc = float(df_mp_dest_f[col_val_mp].sum())
                                    else:
                                        col_val_guess = detect_column_by_keywords(h_mp, ["valor", "soma", "total", "amount", "receita", "vl"])
                                        if col_val_guess and col_val_guess in df_mp_periodo.columns:
                                            df_guess = df_mp_periodo.copy()
                                            if col_b2_mp in df_guess.columns:
                                                df_guess = df_guess[df_guess[col_b2_mp].astype(str).str.strip() == str(b2).strip()]
                                            if lojas_audit and ok_loja:
                                                df_guess = df_guess[df_guess[col_loja_mp].apply(normalize_code).isin(lojas_audit)]
                                            if not df_guess.empty:
                                                v_mp_calc = float(df_guess[col_val_guess].sum())
                                v_mp = v_mp_calc
                            else:
                                v_mp = 0.0
                        except Exception:
                            v_mp = 0.0

                        diff = v_o - v_d
                        diff_mp = v_d - v_mp
                        status = "✅ OK" if (abs(diff) < 0.01 and abs(diff_mp) < 0.01) else "❌ Erro"

                        mask_master = st.session_state.au_planilhas_df["Planilha_id"] == sid
                        if mask_master.any():
                            st.session_state.au_planilhas_df.loc[mask_master, "Origem"] = format_brl(v_o)
                            st.session_state.au_planilhas_df.loc[mask_master, "DRE"] = format_brl(v_d)
                            st.session_state.au_planilhas_df.loc[mask_master, "MP DRE"] = format_brl(v_mp)
                            st.session_state.au_planilhas_df.loc[mask_master, "Dif"] = format_brl(diff)
                            st.session_state.au_planilhas_df.loc[mask_master, "Dif MP"] = format_brl(diff_mp)
                            st.session_state.au_planilhas_df.loc[mask_master, "Status"] = status
                            st.session_state.au_planilhas_df.loc[mask_master, "Flag"] = False
                        logs.append(f"{pname}: {status}")
                    except Exception as e:
                        logs.append(f"{pname}: Erro {e}")
                    prog.progress((idx + 1) / total)

                st.markdown("### Log de processamento")
                st.text("\n".join(logs))
                st.success("Auditoria concluída.")
                st.rerun()
