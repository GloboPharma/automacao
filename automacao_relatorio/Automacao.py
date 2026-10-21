import os
import warnings
import pandas as pd
from hdbcli import dbapi
from datetime import date, datetime
from dotenv import load_dotenv
from tqdm import tqdm  # <<< BIBLIOTECA DA BARRA DE PROGRESSO

# Ignora o aviso chato do Pandas sobre SQLAlchemy (UserWarning)
warnings.filterwarnings('ignore', category=UserWarning)

# Carrega as variáveis do arquivo .env (com override para garantir leitura)
load_dotenv(override=True)

# ==========================================
# 1. CONFIGURAÇÕES (Lendo do .env)
# ==========================================
# Credenciais
HANA_HOST = os.getenv("HANA_HOST")
HANA_PORT = int(os.getenv("HANA_PORT", 30015))
HANA_USER = os.getenv("HANA_USER")
HANA_PASSWORD = os.getenv("HANA_PASSWORD")

# Schema dinâmico (Usa o valor do .env ou padrão 'GLOBOPHARMA')
SCHEMA = os.getenv("HANA_SCHEMA", "SBOGPHARMA")

# Caminhos dos diretórios (Adicione os novos no seu .env)
PATH_VENDAS = os.getenv("PATH_VENDAS", "./Vendas")
PATH_ESTOQUE_BD = os.getenv("PATH_ESTOQUE_BD", "./Estoque")
PATH_ESTOQUE_POSICAO = os.getenv("PATH_ESTOQUE_POSICAO", "./Estoque_Posicao")
PATH_DEVOLUCOES = os.getenv("PATH_DEVOLUCOES", "./Devolucoes")
PATH_LOGS = os.getenv("PATH_LOGS", "./Logs")
PATH_ESTOQUE_BD_DEVOLUCOES = os.getenv("PATH_ESTOQUE_BD_DEVOLUCOES", "./Devolucoes")
PATH_PROJ = os.getenv("PATH_PROJ", "./Projetos")

# Define as datas dinâmicas (mantido caso você use nas outras queries)
data_inicio = date(date.today().year, 1, 1).strftime('%Y-%m-%d')
data_fim = date.today().strftime('%Y-%m-%d')

# ==========================================
# 2. CONSULTAS SQL (ORIGINAIS)
# ==========================================

QUERY_BOLETIM_VENDAS = f""" 
SELECT *
FROM "{SCHEMA}"."VW_RELATORIO_BONIFICACOES_DEVOLUCOES_VOLUMES"
WHERE YEAR("Data de Lançamento") = YEAR(CURRENT_DATE)
  AND "Data de Lançamento" <= CURRENT_DATE;
"""

QUERY_ESTOQUE_GERAL = f"""
SELECT * FROM "{SCHEMA}"."PBI_FAT_BEAS_OP"
"""

QUERY_ESTOQUE_LOTE = f"""
SELECT * FROM "{SCHEMA}"."View Estoque"
"""

QUERY_ESTOQUE_COM_E_SEM = f"""SELECT * 
FROM "{SCHEMA}"."VW_ESTOQUE_COM_E_SEM"
ORDER BY
    "ItemCode",
    "StatusLote",
    "DiasSemMovimentacao" DESC,
    "WhsCode",
    "Posicao",
    "Lote";
"""

# ==========================================
# 2.1 CONSULTAS SQL (NOVAS COM SCHEMA APLICADO)
# ==========================================

QUERY_ACOMPANHAMENTO_DEV = f"""
SELECT * 
FROM "{SCHEMA}"."VW_ACOMPANHAMENTO_DEVOLUCAO"
ORDER BY "Data do Pedido de Devolução" DESC;
"""

QUERY_LOG_BLOQUEIO = f"""
SELECT * 
FROM "{SCHEMA}"."VW_RASTREABILIDADE_LOTES"
WHERE "Tipo_Documento_SAP" IN (20, 18, 59, 16, 14, 67, 69)
ORDER BY "Codigo_Item", "Numero_Lote", "Data_Transacao";
"""

QUERY_IBT1_DEV = f"""
SELECT * FROM "{SCHEMA}"."IBT1" WHERE ("WhsCode"='01.08' OR "WhsCode"='01.07' OR "WhsCode"='01.09' OR "WhsCode"='01.21')
"""

QUERY_PROCESSO_DEV = f"""
SELECT * 
FROM "{SCHEMA}"."VW_PROC_DEVOLUCAO"
ORDER BY
    "Situação do Processo",
    "Data Pedido Devolução",
    "Nº Esboço",
    "Nº Pedido Devolução",
    "Nº Devolução",
    "Nº Doc. NF Devolução",
    "Item";
"""


QUERY_RELATORIO_PEDIDOS_V3 = f"""
SELECT * 
FROM "{SCHEMA}"."VW_RELATORIO_PEDIDOS_V3"
ORDER BY "Data de Lançamento Esboço" DESC, "Key2" DESC;
"""

# ==========================================
# 2.2 NOVAS CONSULTAS (ADICIONADA AGORA)
# ==========================================

QUERY_ATRIBUTOS_LOTE = f"""
SELECT * 
FROM "{SCHEMA}"."VW_ATRIBUTOS_LOTE_VEICULOS";
"""

# ==========================================
# 3. DEFINIÇÃO DOS RELATÓRIOS
# ==========================================
relatorios = [
    {
        "nome_relatorio": "Boletim de Vendas - Versão 3",
        "caminho": PATH_VENDAS,
        "query": QUERY_BOLETIM_VENDAS
    },
    {
        "nome_relatorio": "Estoque Geral",
        "caminho": PATH_ESTOQUE_BD,
        "query": QUERY_ESTOQUE_GERAL
    },
    {
        "nome_relatorio": "Estoque Por Lote",
        "caminho": PATH_ESTOQUE_BD,
        "query": QUERY_ESTOQUE_LOTE 
    },
    {
        "nome_relatorio": "Estoque Geral Com e Sem Lotes Com Posicao",
        "caminho": PATH_ESTOQUE_POSICAO,
        "query": QUERY_ESTOQUE_COM_E_SEM
    },
    # --- NOVAS CONSULTAS ---
    {
        "nome_relatorio": "Acompanhamento Devolução Teste",
        "caminho": PATH_DEVOLUCOES,
        "query": QUERY_ACOMPANHAMENTO_DEV
    },
    {
        "nome_relatorio": "Log de Bloqueio de Lotes (Status 1 e 2)",
        "caminho": PATH_LOGS,
        "query": QUERY_LOG_BLOQUEIO
    },
    {
        "nome_relatorio": "IBT1 (Devoluções 01.08)",
        "caminho": PATH_DEVOLUCOES,
        "query": QUERY_IBT1_DEV
    },
    {
        "nome_relatorio": "Processo_Dev_Aberto v2 (Com Concluídas)",
        "caminho": PATH_DEVOLUCOES,
        "query": QUERY_PROCESSO_DEV
    },
    {
        "nome_relatorio": "Relatório de Pedidos de Devolução V3",
        "caminho": PATH_DEVOLUCOES,
        "query": QUERY_RELATORIO_PEDIDOS_V3
    },
    {
        "nome_relatorio": "Atributos de Lote",
        "caminho": PATH_ESTOQUE_BD,
        "query": QUERY_ATRIBUTOS_LOTE
    }
]

# ==========================================
# 4. CONEXÃO E EXECUÇÃO
# ==========================================
print("Conectando ao SAP HANA externo...", flush=True)
try:
    conn = dbapi.connect(
        address=HANA_HOST,
        port=HANA_PORT,
        user=HANA_USER,
        password=HANA_PASSWORD
    )
    print("Conexão estabelecida com sucesso!", flush=True)
except Exception as e:
    print(f"\n❌ Erro ao conectar ao SAP HANA: {e}", flush=True)
    exit(1)

# Data para o nome do arquivo
data_hoje = datetime.now().strftime("%Y-%m-%d")

# Percorre a lista de relatórios
for item in relatorios:
    try:
        print(f"\n--- Processando: {item['nome_relatorio']} ---")
        
        print("Executando SQL no SAP HANA e baixando dados...")
        
        # 1. Download em blocos com Barra de Progresso Animada
        chunksize = 100000 # Baixa 100.000 linhas por vez
        chunks = []
        
        # O tqdm vai criar uma barra de progresso contando os blocos conforme o HANA for enviando os dados
        for chunk in tqdm(pd.read_sql(item['query'], conn, chunksize=chunksize), desc="Baixando dados"):
            chunks.append(chunk)
            
        # Concatena todos os blocos recebidos
        if chunks:
            df = pd.concat(chunks, ignore_index=True)
        else:
            df = pd.DataFrame()
            
        # Mostra o tamanho exato final
        print(f"\nDownload concluído! A consulta retornou {len(df):,} linhas e {len(df.columns)} colunas.")
        
        # 2. Garante que a pasta existe
        if item['caminho']:
            os.makedirs(item['caminho'], exist_ok=True)
        else:
            # Se a variável estiver vazia/nula, usa o diretório atual
            item['caminho'] = "."
        
        # 3. Validação de Limite do Excel (1.048.576 linhas)
        if len(df) > 1048000:
            print("AVISO: O arquivo ultrapassa o limite do Excel! Salvando como CSV para não perder dados.")
            nome_arquivo = f"{item['nome_relatorio']}_{data_hoje}.csv"
            caminho_final = os.path.join(item['caminho'], nome_arquivo)
            print(f"Salvando CSV em: {caminho_final}")
            df.to_csv(caminho_final, index=False, sep=';', encoding='utf-8-sig')
        else:
            nome_arquivo = f"{item['nome_relatorio']}_{data_hoje}.xlsx"
            caminho_final = os.path.join(item['caminho'], nome_arquivo)
            print(f"Salvando XLSX em: {caminho_final} ...")
            df.to_excel(caminho_final, index=False, engine='xlsxwriter')
            
        print("Sucesso!")

    except Exception as e:
        print(f"ERRO ao gerar '{item['nome_relatorio']}': {e}")

# Fecha a conexão
conn.close()
print("\n==========================================")
print("Processamento concluído para todos os itens.")
print("==========================================")