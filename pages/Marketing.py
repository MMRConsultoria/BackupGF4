import streamlit as st
 import psycopg2
 import pandas as pd
 from io import BytesIO

CERT_PATH = "aws-us-east-2-bundle.pem"

# Grava o certificado em arquivo apenas uma vez por sessão 
if  "cert_write"  not  in st.session_state:
     with  open (CERT_PATH, "w" ) as f:
        f.write(st.secrets[ "certs" ][ "aws_rds_us_east_2" ])
    st.session_state[ "cert_written" ] = True

def  get_conn ():
    conexão = psycopg2.connect(
        host=st.secrets[ "db" ][ "host" ],
        porta=st.secrets[ "db" ][ "porta" ],
        dbname=st.secrets[ "db" ][ "database" ],
        usuário=st.secrets[ "db" ][ "user" ],
        senha=st.secrets[ "db" ][ "senha" ],
        sslmode= "verify-full" ,
        sslrootcert=CERT_PATH,
    )
    retornar conexão

def  get_all_tables ( conn ):
    consulta = """
    SELECIONE o esquema da tabela, o nome da tabela
    A PARTIR DE information_schema.tables
    ONDE table_type = 'BASE TABLE' AND table_schema NOT IN ('pg_catalog', 'information_schema');
    """
    df = pd.read_sql(query, conn)
    retornar df

def  fetch_table_data ( conn, schema, table ):
    consulta = f'SELECT * FROM " {esquema} "." {tabela} "'
    df = pd.read_sql(query, conn)
    retornar df

st.title( "Exportar todas as tabelas do banco para Excel" )

if st.button( "Gerar Excel" ):
     try :
        conexão = obter_conexão()
        tables_df = get_all_tables(conn)
        se tables_df.empty:
            st.warning( "Nenhuma tabela encontrada no banco." )
         else :
            saída = BytesIO()
            com pd.ExcelWriter(output, engine= 'openpyxl' ) como escritor:
                 para idx, linha em tables_df.iterrows():
                    esquema = linha[ 'esquema_da_tabela' ]
                    tabela = linha[ 'nome_da_tabela' ]
                    st.write( f"Lendo tabela: {schema} . {table} " )
                    df = buscar_dados_da_tabela(conn, esquema, tabela)
                    # Nome da aba: schema_table (máximo 31 caracteres) 
                    sheet_name = f" {schema} _ {table} " [: 31 ]
                    df.to_excel(writer, sheet_name=sheet_name, index= False )
            conexão.fechar()
            saída.buscar( 0 )
            st.success( "Arquivo Excel gerado com sucesso!" )
            st.download_button(
                rótulo= "Baixar Excel" ,
                dados=saída,
                nome_do_arquivo= "banco_completo.xlsx" ,
                mime= "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    exceto Exception como e:
        st.error( f"Erro ao gerar Excel: {e} " )
