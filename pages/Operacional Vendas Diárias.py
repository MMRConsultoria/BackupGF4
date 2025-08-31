 # sheet
                            df_sh = sheet_por_n[nkey].copy()
                            df_sh = canonize_cols_df(df_sh)
                            if "Data" in df_sh.columns:
                                try:
                                    ser = pd.to_numeric(df_sh["Data"], errors="coerce")
                                    if ser.notna().any():
                                        df_sh["Data"] = pd.to_datetime(ser, origin="1899-12-30", unit="D", errors="coerce").dt.strftime("%d/%m/%Y")
                                    else:
                                        df_sh["Data"] = pd.to_datetime(df_sh["Data"], dayfirst=True, errors="coerce").dt.strftime("%d/%m/%Y")
                                except Exception:
                                    pass
                            for idx, row in df_sh.iterrows():
                                d_sh = canonize_dict(row.to_dict())
                                d_sh["_origem_"] = "Google Sheets"
                                d_sh["Linha Sheet"] = idx + 2   # linha real do Google Sheets
                                conflitos_linhas.append(d_sh)


                        # origem (só se existir)
                        if "_origem_" in df_conf.columns:
                            df_conf["_origem_"] = df_conf["_origem_"].replace({
                                "Nova Arquivo": "🟢 Nova Arquivo",
                                "Google Sheets": "🔴 Google Sheets",
                            })
                        
                        # === SANEAR TIPOS p/ evitar ArrowTypeError no st.data_editor ===
                        num_cols = ["Fat. Total", "Serv/Tx", "Fat.Real", "Ticket"]
                        for c in num_cols:
                            if c in df_conf.columns:
                                df_conf[c] = pd.to_numeric(df_conf[c], errors="coerce").astype("Float64")
                        
                        if "Linha Sheet" in df_conf.columns:
                            df_conf["Linha Sheet"] = pd.to_numeric(df_conf["Linha Sheet"], errors="coerce").astype("Int64")
                        
                        if "Manter" in df_conf.columns:
                            df_conf["Manter"] = (
                                df_conf["Manter"].astype(str).str.strip().str.lower()
                                .isin(["true","1","yes","y","sim","verdadeiro"])
                            ).astype("boolean")
                        
                        _proteger = set(num_cols + ["Linha Sheet", "Manter"])
                        for c in df_conf.columns:
                            if c not in _proteger:
                                df_conf[c] = df_conf[c].astype("string").fillna("")
                        # === FIM SANEAR TIPOS ===
                        
                        with st.form("form_conflitos_globais"):
                            edited_conf = st.data_editor(
                                df_conf,
                                use_container_width=True,
                                hide_index=True,
                                key="editor_conflitos",
                                column_config={
                                    "Manter": st.column_config.CheckboxColumn(
                                        help="Marque quais linhas (de cada N) deseja manter",
                                        default=False
                                    )
                                }
                            )
                            aplicar_tudo = st.form_submit_button("🗑️ Excluir linhas do Google Sheets")

                        # (2) DataFrame consolidado
                        df_conf = pd.DataFrame(conflitos_linhas).copy()
                        df_conf = canonize_cols_df(df_conf)
                        
                        # (3) derivados de data
                        if "Data" in df_conf.columns:
                            _dt = pd.to_datetime(df_conf["Data"], dayfirst=True, errors="coerce")
                            nomes_dia = ["segunda-feira","terça-feira","quarta-feira","quinta-feira","sexta-feira","sábado","domingo"]
                            df_conf["Dia da Semana"] = _dt.dt.dayofweek.map(lambda i: nomes_dia[i].title() if pd.notna(i) else "")
                            nomes_mes = ["jan","fev","mar","abr","mai","jun","jul","ago","set","out","nov","dez"]
                            df_conf["Mês"] = _dt.dt.month.map(lambda m: nomes_mes[m-1] if pd.notna(m) else "")
                            df_conf["Ano"] = _dt.dt.year.fillna("").astype(str).replace("nan","")
                        
                        # (4) origem com emojis
                        df_conf["_origem_"] = df_conf["_origem_"].replace({
                            "Nova Arquivo": "🟢 Nova Arquivo",
                            "Google Sheets": "🔴 Google Sheets"
                        })
                        
                        # (5) coluna Manter
                        if "Manter" not in df_conf.columns:
                            df_conf.insert(0,"Manter",False)
                        
                        # (6) reordena
                        ordem_final = [
                            "Manter","_origem_","Linha Sheet","Data","Dia da Semana","Loja",
                            "Codigo Everest","Grupo","Cod Grupo Empresas",
                            "Fat. Total","Serv/Tx","Fat.Real","Ticket","Mês","Ano","M","N"
                        ]

                        cols_final = [c for c in ordem_final if c in df_conf.columns] + [c for c in df_conf.columns if c not in ordem_final]
                        df_conf = df_conf.reindex(columns=cols_final, fill_value="")
                        
                        st.markdown("<div style='color:#555; font-size:0.9rem; font-weight:500; margin:10px 0;'>🔴 Possíveis duplicados — marque o(s) que deseja manter</div>", unsafe_allow_html=True)
                        

                        # ... você já montou df_conf e fez o reindex:
                        # df_conf = df_conf.reindex(columns=cols_final, fill_value="")
                        
                        # === SANEAR TIPOS p/ evitar ArrowTypeError no st.data_editor ===
                        num_cols = ["Fat. Total", "Serv/Tx", "Fat.Real", "Ticket"]
                        for c in num_cols:
                            if c in df_conf.columns:
                                df_conf[c] = pd.to_numeric(df_conf[c], errors="coerce").astype("Float64")
                        
                        if "Linha Sheet" in df_conf.columns:
                            df_conf["Linha Sheet"] = pd.to_numeric(df_conf["Linha Sheet"], errors="coerce").astype("Int64")
                        
                        if "Manter" in df_conf.columns:
                            df_conf["Manter"] = (
                                df_conf["Manter"].astype(str).str.strip().str.lower()
                                .isin(["true","1","yes","y","sim","verdadeiro"])
                            ).astype("boolean")
                        
                        # Demais colunas em string para não sobrar 'object' com tipos mistos
                        _proteger = set(num_cols + ["Linha Sheet", "Manter"])
                        for c in df_conf.columns:
                            if c not in _proteger:
                                df_conf[c] = df_conf[c].astype(str)
                        # === FIM SANEAR TIPOS ===
                        
                        # (agora vem exatamente o seu form)
                        with st.form("form_conflitos_globais"):
                            edited_conf = st.data_editor(
                                df_conf,
                                use_container_width=True,
                                hide_index=True,
                                key="editor_conflitos",
                                column_config={
                                    "Manter": st.column_config.CheckboxColumn(
                                        help="Marque quais linhas (de cada N) deseja manter",
                                        default=False
                                    )
                                }
                            )
                            aplicar_tudo = st.form_submit_button("✅ Atualizar planilha")

                        
                    
                        if aplicar_tudo:
               
                            try:
                                # Garante que a aba está acessível
                                try:
                                    _ = aba_destino.id
                                except Exception:
                                    gc = get_gc()
                                    planilha_destino = gc.open("Vendas diarias")
                                    aba_destino = planilha_destino.worksheet("Fat Sistema Externo")
                        
                                # Normaliza 'Manter' e filtra somente origem Google
                                manter_series = edited_conf["Manter"]
                                if manter_series.dtype != bool:
                                    manter_series = manter_series.astype(str).str.strip().str.lower().isin(
                                        ["true","1","yes","y","sim","verdadeiro"]
                                    )
                                mask_google = edited_conf["_origem_"].astype(str).str.contains("google", case=False, na=False)
                        
                                # Linhas reais (1-based) vindas da coluna "Linha Sheet"
                                linhas = (
                                    pd.to_numeric(edited_conf.loc[mask_google & manter_series, "Linha Sheet"], errors="coerce")
                                    .dropna().astype(int).tolist()
                                )
                                linhas = sorted({ln for ln in linhas if ln >= 2}, reverse=True)
                        
                                st.warning(f"📄 Planilha: {aba_destino.spreadsheet.title} / Aba: {aba_destino.title} (sheetId={aba_destino.id})")
                                st.warning(f"🧮 Linhas a excluir (1-based): {linhas}")
                        
                                if not linhas:
                                    st.error("Nenhuma linha do Google Sheets marcada/identificada para exclusão (confira 'Manter' e 'Linha Sheet').")
                                    st.stop()
                        
                                # Exclusão robusta via batchUpdate (deleteDimension)
                                sheet_id = int(aba_destino.id)
                                requests = [
                                    {
                                        "deleteDimension": {
                                            "range": {
                                                "sheetId": sheet_id,
                                                "dimension": "ROWS",
                                                "startIndex": ln - 1,  # 0-based inclusivo
                                                "endIndex": ln        # 0-based exclusivo
                                            }
                                        }
                                    }
                                    for ln in linhas   # já em ordem DESC
                                ]
                                aba_destino.spreadsheet.batch_update({"requests": requests})
                        
                                st.success(f"🗑️ {len(linhas)} linha(s) excluída(s) do Google Sheets. Atualize a planilha no navegador para ver.")
                                st.stop()
                        
                            except Exception as e:
                                st.error(f"❌ Erro ao excluir linhas do Google Sheets: {e}")
                                st.stop()




                        
                        pode_enviar=False
