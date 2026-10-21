import dotenv, os, hdbcli.dbapi as dbapi

dotenv.load_dotenv(r'c:\Users\frederico.magalhaes\Documents\script\automacao_relatorio\.env')
conn = dbapi.connect(
    address=os.getenv('HANA_HOST'),
    port=int(os.getenv('HANA_PORT', 30015)),
    user=os.getenv('HANA_USER'),
    password=os.getenv('HANA_PASSWORD')
)
cur = conn.cursor()
cur.execute('SET SCHEMA "SBOGPHARMA"')

with open(r'c:\Users\frederico.magalhaes\Documents\db_script\atributos_lote.sql', encoding='utf-8') as f:
    sql = f.read()

cleaned = sql.replace('"{SCHEMA}".', '')
select_part = cleaned.split('CREATE OR REPLACE VIEW "SBOGPHARMA"."VW_ATRIBUTOS_LOTE_VEICULOS" AS')[1]

try:
    cur.execute('EXPLAIN PLAN FOR ' + select_part)
    print('EXPLAIN PLAN SUCCESS: Query is 100% syntactically and semantically valid in SAP HANA!')
except Exception as e:
    print('VALIDATION ERROR:', e)

cur.close()
conn.close()
