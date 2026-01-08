import pandas as pd
import numpy as np
import json
import ast

# Carregar o arquivo original para processar conforme as novas regras
input_file = 'banco_completo.xlsx'
df = pd.read_excel(input_file)

# 1. Converter datas e filtrar (A partir de 01/12/2025)
df['business_dt'] = pd.to_datetime(df['business_dt'], errors='coerce')
data_corte = pd.Timestamp('2025-12-01')
df = df[df['business_dt'] >= data_corte].copy()

# 2. Filtrar lojas (Excluir 0000, 0001, 9999)
# Garantir que store_code seja string para comparar corretamente
df['store_code'] = df['store_code'].astype(str).str.zfill(4)
excluir = ['0000', '0001', '9999']
df = df[~df['store_code'].isin(excluir)].copy()

# 3. Extrair campos de custom_properties (TIP_AMOUNT e VOID_TYPE)
def parse_props(x):
    if pd.isna(x): return {}
    try:
        if isinstance(x, str):
            return json.loads(x)
    except:
        try:
            return ast.literal_eval(x)
        except:
            return {}
    return x if isinstance(x, dict) else {}

props = df['custom_properties'].apply(parse_props)
df['TIP_AMOUNT'] = pd.to_numeric(props.apply(lambda x: x.get('TIP_AMOUNT')), errors='coerce').fillna(0)
df['VOID_TYPE'] = props.apply(lambda x: x.get('VOID_TYPE'))

# 4. Desconsiderar registros com VOID_TYPE preenchido
df = df[df['VOID_TYPE'].isna() | (df['VOID_TYPE'] == "") | (df['VOID_TYPE'] == 0)].copy()

# 5. Agrupar totais por store_code e business_dt
resumo = df.groupby(['store_code', df['business_dt'].dt.date]).agg(
    total_gross=('total_gross', 'sum'),
    total_tip=('TIP_AMOUNT', 'sum'),
    qtd_pedidos=('order_code', 'count')
).reset_index()

# Salvar o novo Excel
output_name = 'resumo_vendas_dezembro.xlsx'
resumo.to_excel(output_name, index=False)
