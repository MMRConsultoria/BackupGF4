if drive_service and folder_ids:
    for fid in folder_ids:
        st.write(f"Arquivos na pasta {fid}:")
        try:
            arquivos = listar_arquivos_pasta(drive_service, fid)
            if arquivos:
                for a in arquivos:
                    st.write(f"- {a['name']} (ID: {a['id']})")
            else:
                st.write("Nenhum arquivo encontrado ou sem permissão.")
        except Exception as e:
            st.error(f"Erro ao listar pasta {fid}: {e}")
else:
    st.info("Drive API não disponível ou nenhuma pasta configurada.")
