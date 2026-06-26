import pandas as pd
from sqlalchemy import create_engine
import os
from dotenv import load_dotenv

load_dotenv('.env')

db_user = os.getenv('DB_USER', 'postgres')
db_pass = os.getenv('DB_PASS', '')
db_host = os.getenv('DB_HOST', 'localhost')
db_port = os.getenv('DB_PORT', '5432')
db_name = os.getenv('DB_NAME', 'powerbi_reports')
conn_string = f'postgresql://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}'

engine = create_engine(conn_string)
query = 'SELECT DISTINCT "Key" FROM "03_entrega"."salescalc_tower_sales"'
df = pd.read_sql(query, con=engine)
keys_unicas = sorted(df['Key'].dropna().unique())

print(f'Total de torres únicas en PostgreSQL: {len(keys_unicas)}')
print('\n10 primeros ejemplos de nombres normalizados:')
for k in keys_unicas[:10]:
    print(f'  - [{k}]')
    
ts90 = [k for k in keys_unicas if 'TS90' in k]
print('\nTorres que contienen TS90:')
for k in ts90:
    print(f'  - [{k}]')
    
tcs = [k for k in keys_unicas if 'TCS164' in k][:3]
print('\nEjemplos de torres TCS:')
for k in tcs:
    print(f'  - [{k}]')
