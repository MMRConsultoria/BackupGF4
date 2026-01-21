# ---------------- ABA: AUDITORIA (sem st_aggrid) ----------------
import pandas as pd
from datetime import date, timedelta

with tab_audit:
    st.header("Auditoria")

    # Seleção de pastas
    try:
        pastas_fech = list_child_folders(drive_service, PASTA_PRINCIPAL_ID, "fechamento")
        map_p = {p["name"]: p["id"] for p in pastas_fech}
        p_sel = st.selectbox("Pasta principal:", options=list(map_p.keys()), key="au_p")
        subpastas = list_child_folders(drive_service, map_p[p_sel])
        map_s = {s["name"]: s["id"] for s in subpastas}
        s_sel = st.multiselect("Subpastas:", options=list(map_s.keys()), default=[], key="au_s")
        s_ids_audit = [map_s[n] for n in s_sel]
    except Exception as e:
        st.error(f"Erro ao listar pastas: {e}")
        st.stop()

    # Período
    c1, c2 = st.columns(2)
    with c1:
        ano_sel = st.selectbox("Ano:", list(range(2020, date.today().year + 1)),
                               index=date.today().year - 2020, key="au_ano")
    with c2:
        mes_sel = st.selectbox("Mês (Opcional):", ["Todos"] + list(range(1, 13)), key="au_mes")

    # Helpers locais
    def detect_column_by_keywords(headers, keywords_list):
        for kw in keywords_list:
            for h in headers:
                if kw in str(h).lower():
                    return h
        return None

    def normalize_code(val):
        try:
            f = float(val)
            i = int(f)
            return str(i) if f == i else str(f)
        except Exception:
            return str(val).strip()

    # Carrega lista de planilhas para as subpastas (recarrega se mudarem subpastas)
    if "au_last_subpastas" not in st.session_state or st.session_state.au_last_subpastas != s_ids_audit:
        try:
            planilhas = list_spreadsheets_in_folders(drive_service, s_ids_audit)
        except Exception as e:
            st.error(f"Erro ao listar planilhas nas subpastas: {e}")
            st.stop()

        df_init = pd.DataFrame([{
            "Planilha": p["name"],
            "Planilha_id": p["id"],
            "Auditar": False,
            "Origem": "",
            "DRE": "",
            "MP DRE": "",
            "Dif": "",
            "Dif MP": "",
            "Status": ""
        } for p in planilhas])

        st.session_state.au_last_subpastas = s_ids_audit
        st.session_state.au_planilhas_df = df_init
        st.session_state.au_resultados = {}  # id -> result

    # DataFrame atual
    df_table = st.session_state.au_planilhas_df.copy()

    st.markdown("**1) Selecione as planilhas que deseja auditar**")

    # Se poucas planilhas, mostrar checkbox por linha (interface parecida com sua imagem).
    MAX_CHECKBOX_ROWS = 80
    if len(df_table) <= MAX_CHECKBOX_ROWS:
        sel_cols = st.columns([0.08, 0.92])  # coluna para checkboxes + nomes
        sel_container = st.container()
        selected_ids = []
        for i, row in df_table.iterrows():
            chk_key = f"au_chk_{row['Planilha_id']}"
            # preserva estado anterior
            prev = st.session_state.get(chk_key, False)
            cols = sel_container.columns([0.08, 0.92])
            checked = cols[0].checkbox("", value=prev, key=chk_key)
            cols[1].markdown(f"**{row['Planilha']}** - {row['Status'] or ''}")
            if checked:
                selected_ids.append(row["Planilha_id"])
    else:
        # Se muitas planilhas, usar multiselect por nome (mais performático)
        st.info("Muitas planilhas — use a seleção múltipla abaixo.")
        names = df_table["Planilha"].tolist()
        selected_names = st.multiselect("Planilhas para auditar:", options=names, key="au_multisel")
        selected_ids = [r["Planilha_id"] for r in df_table.to_dict("records") if r["Planilha"] in selected_names]

    # Botões de ação
    c_run, c_clear = st.columns([1, 1])
    run = c_run.button("📊 EXECUTAR AUDITORIA (somente marcadas)")
    clear = c_clear.button("🔁 Desmarcar todas")

    if clear:
        # reset flags
        if len(df_table) <= MAX_CHECKBOX_ROWS:
            for _, r in df_table.iterrows():
                st.session_state[f"au_chk_{r['Planilha_id']}"] = False
        else:
            st.session_state["au_multisel"] = []
        # Also update session_state table
        st.session_state.au_planilhas_df["Auditar"] = False
        st.experimental_rerun()

    # Mostra tabela acumulada de resultados ao lado/abaixo
    def show_accumulated_table():
        # Build display DF from session_state.au_planilhas_df
        display_df = st.session_state.au_planilhas_df.copy()
        # garantir colunas na ordem desejada
        cols_order = ["Planilha", "Origem", "DRE", "MP DRE", "Dif", "Dif MP", "Status"]
        display_df = display_df[cols_order]
        st.markdown("**Resultados acumulados**")
        st.dataframe(display_df, height=300)

    show_accumulated_table()

    # Roda auditoria nas selecionadas
    if run:
        # Recalcula selected_ids in case of multiselect
        if len(df_table) > MAX_CHECKBOX_ROWS:
            # already built selected_ids from multiselect
            pass
        else:
            # rebuild from session_state checkboxes
            selected_ids = [r["Planilha_id"] for _, r in df_table.iterrows() if st.session_state.get(f"au_chk_{r['Planilha_id']}", False)]

        if not selected_ids:
            st.warning("Nenhuma planilha selecionada para auditar.")
        else:
            # intervalo
            if mes_sel == "Todos":
                d_ini, d_fim = date(ano_sel, 1, 1), date(ano_sel, 12, 31)
            else:
                d_ini = date(ano_sel, int(mes_sel), 1)
                d_fim = (date(ano_sel, int(mes_sel), 28) + timedelta(days=4)).replace(day=1) - timedelta(days=1)

            # Carregar origem de faturamento (uma vez)
            try:
                sh_o_fat = gc.open_by_key(ID_PLANILHA_ORIGEM_FAT)
                ws_o_fat = sh_o_fat.worksheet(ABA_ORIGEM_FAT)
                h_o_fat, df_o_fat = get_headers_and_df_raw(ws_o_fat)
                if not df_o_fat.empty:
                    df_o_fat = tratar_numericos(df_o_fat, h_o_fat)

                c_dt_o_fat = detect_date_col(h_o_fat) or (h_o_fat[0] if h_o_fat else None)
                if c_dt_o_fat and not df_o_fat.empty:
                    df_o_fat["_dt"] = pd.to_datetime(df_o_fat[c_dt_o_fat], dayfirst=True, errors="coerce")
                    parsed_pct = df_o_fat["_dt"].notna().mean()
                    if parsed_pct == 0:
                        df_o_fat["_dt"] = pd.to_datetime(df_o_fat[c_dt_o_fat], dayfirst=False, errors="coerce")
                    df_o_fat["_dt"] = df_o_fat["_dt"].dt.date
                    df_o_fat_p = df_o_fat[(df_o_fat["_dt"] >= d_ini) & (df_o_fat["_dt"] <= d_fim)].copy()
                else:
                    df_o_fat_p = df_o_fat.copy()
            except Exception as e:
                st.error(f"Erro ao carregar origem de faturamento: {e}")
                st.stop()

            planilhas_map = {r["Planilha_id"]: r["Planilha"] for _, r in df_table.iterrows()}
            total = len(selected_ids)
            prog = st.progress(0)
            log_lines = []

            for idx, sid in enumerate(selected_ids):
                pname = planilhas_map.get(sid, "Desconhecido")
                v_o = v_d = v_mp_d = 0.0
                status = "Erro desconhecido"

                # abrir planilha
                try:
                    sh_d = gc.open_by_key(sid)
                except Exception as e:
                    status = f"Erro ao abrir planilha ({e})"
                    log_lines.append(f"{pname}: {status}")
                    # salvar resultado parcial
                    st.session_state.au_resultados[sid] = {
                        "Planilha": pname, "Origem": 0.0, "DRE": 0.0, "MP DRE": 0.0,
                        "Dif": 0.0, "Dif MP": 0.0, "Status": status
                    }
                    prog.progress((idx + 1) / total)
                    continue

                # ler códigos
                b2, b3 = read_codes_from_config_sheet(sh_d)
                if not b2:
                    status = "Sem B2 (Config)"
                    log_lines.append(f"{pname}: {status}")
                    st.session_state.au_resultados[sid] = {
                        "Planilha": pname, "Origem": 0.0, "DRE": 0.0, "MP DRE": 0.0,
                        "Dif": 0.0, "Dif MP": 0.0, "Status": status
                    }
                    prog.progress((idx + 1) / total)
                    continue

                # FATURAMENTO ORIGEM
                try:
                    if len(h_o_fat) > 5 and not df_o_fat_p.empty:
                        col_b2_fat = h_o_fat[5]
                        df_filter = df_o_fat_p[df_o_fat_p[col_b2_fat].astype(str).str.strip() == str(b2).strip()]
                        if b3 and len(h_o_fat) > 3:
                            col_b3_fat = h_o_fat[3]
                            df_filter = df_filter[df_filter[col_b3_fat].astype(str).str.strip() == str(b3).strip()]
                        if len(h_o_fat) > 6:
                            v_o = df_filter[h_o_fat[6]].sum() if not df_filter.empty else 0.0
                except Exception:
                    v_o = 0.0

                # FATURAMENTO DESTINO
                try:
                    ws_d = sh_d.worksheet("Importado_Fat")
                    h_d, df_d = get_headers_and_df_raw(ws_d)
                    if not df_d.empty:
                        df_d = tratar_numericos(df_d, h_d)

                    c_dt_d = detect_date_col(h_d) or (h_d[0] if h_d else None)
                    if c_dt_d and not df_d.empty:
                        df_d["_dt"] = pd.to_datetime(df_d[c_dt_d], dayfirst=True, errors="coerce")
                        if df_d["_dt"].isna().all():
                            df_d["_dt"] = pd.to_datetime(df_d[c_dt_d], dayfirst=False, errors="coerce")
                        df_d["_dt"] = df_d["_dt"].dt.date
                        df_d_periodo = df_d[(df_d["_dt"] >= d_ini) & (df_d["_dt"] <= d_fim)]
                    else:
                        df_d_periodo = df_d.copy()

                    if len(h_d) > 6 and not df_d_periodo.empty:
                        v_d = df_d_periodo[h_d[6]].sum()
                    else:
                        v_d = 0.0
                except Exception:
                    v_d = 0.0

                # MEIO DE PAGAMENTO
                try:
                    ws_mp_d = sh_d.worksheet("Meio de Pagamento")
                    h_mp_d, df_mp_d = get_headers_and_df_raw(ws_mp_d)
                    if not df_mp_d.empty:
                        df_mp_d = tratar_numericos(df_mp_d, h_mp_d)

                    c_dt_mp_d = detect_date_col(h_mp_d) or (h_mp_d[0] if h_mp_d else None)
                    if c_dt_mp_d and not df_mp_d.empty:
                        df_mp_d["_dt"] = pd.to_datetime(df_mp_d[c_dt_mp_d], dayfirst=True, errors="coerce")
                        if df_mp_d["_dt"].isna().all():
                            df_mp_d["_dt"] = pd.to_datetime(df_mp_d[c_dt_mp_d], dayfirst=False, errors="coerce")
                        if "_dt" in df_mp_d.columns:
                            df_mp_d["_dt"] = df_mp_d["_dt"].dt.date
                        df_mp_periodo = df_mp_d[(df_mp_d.get("_dt") >= d_ini) & (df_mp_d.get("_dt") <= d_fim)] if "_dt" in df_mp_d.columns else df_mp_d.copy()
                    else:
                        df_mp_periodo = df_mp_d.copy()

                    v_mp_d = 0.0
                    if len(h_mp_d) > 9 and not df_mp_periodo.empty:
                        col_b2_mp = h_mp_d[8]
                        col_b3_mp = h_mp_d[6]
                        col_val_mp = h_mp_d[9]

                        b2_norm = normalize_code(b2)
                        b3_norm = normalize_code(b3) if b3 else None

                        mask = df_mp_periodo[col_b2_mp].apply(normalize_code) == b2_norm
                        if b3_norm:
                            mask &= df_mp_periodo[col_b3_mp].apply(normalize_code) == b3_norm

                        df_mp_dest_f = df_mp_periodo[mask]

                        if not df_mp_dest_f.empty:
                            v_mp_d = df_mp_dest_f[col_val_mp].sum()
                        else:
                            col_val_guess = detect_column_by_keywords(h_mp_d, ["valor", "soma", "total", "amount"])
                            if col_val_guess and col_val_guess in df_mp_periodo.columns:
                                df_guess = df_mp_periodo
                                col_b2_guess = h_mp_d[8] if len(h_mp_d) > 8 else None
                                col_b3_guess = h_mp_d[6] if len(h_mp_d) > 6 else None
                                if col_b2_guess:
                                    df_guess = df_guess[df_guess[col_b2_guess].astype(str).str.strip() == str(b2).strip()]
                                if b3 and col_b3_guess:
                                    df_guess = df_guess[df_guess[col_b3_guess].astype(str).str.strip() == str(b3).strip()]
                                if not df_guess.empty:
                                    v_mp_d = df_guess[col_val_guess].sum()
                    else:
                        v_mp_d = 0.0
                except Exception:
                    v_mp_d = 0.0

                diff = v_o - v_d
                diff_mp = v_d - v_mp_d
                status = "✅ OK" if (abs(diff) < 0.01 and abs(diff_mp) < 0.01) else "❌ Erro"

                # salvar resultado acumulado
                st.session_state.au_resultados[sid] = {
                    "Planilha": pname, "Origem": v_o, "DRE": v_d, "MP DRE": v_mp_d,
                    "Dif": diff, "Dif MP": diff_mp, "Status": status
                }

                # atualizar tabela em session_state (formatar com format_brl)
                mask = st.session_state.au_planilhas_df["Planilha_id"] == sid
                if mask.any():
                    st.session_state.au_planilhas_df.loc[mask, "Origem"] = format_brl(v_o)
                    st.session_state.au_planilhas_df.loc[mask, "DRE"] = format_brl(v_d)
                    st.session_state.au_planilhas_df.loc[mask, "MP DRE"] = format_brl(v_mp_d)
                    st.session_state.au_planilhas_df.loc[mask, "Dif"] = format_brl(diff)
                    st.session_state.au_planilhas_df.loc[mask, "Dif MP"] = format_brl(diff_mp)
                    st.session_state.au_planilhas_df.loc[mask, "Status"] = status
                    st.session_state.au_planilhas_df.loc[mask, "Auditar"] = False  # desmarcar

                log_lines.append(f"{pname}: {status if status != '✅ OK' else 'OK'}")
                prog.progress((idx + 1) / total)

            # mostrar resultados e logs ao final do lote
            resultados_list = list(st.session_state.au_resultados.values())
            if resultados_list:
                df_res = pd.DataFrame(resultados_list)
                for c in ["Origem", "DRE", "MP DRE", "Dif", "Dif MP"]:
                    if c in df_res.columns:
                        df_res[c] = df_res[c].apply(format_brl)
                st.markdown("### Resultados acumulados")
                st.table(df_res)

            st.markdown("---")
            st.subheader("Relatório de Processamento")
            st.text("\n".join(log_lines) if log_lines else "Sem mensagens de processamento.")
            st.success("Auditoria concluída.")

    # fim da aba
