# 📊 Automação de Relatórios — SAP HANA

Script Python para extração automática de dados do **SAP HANA** e geração de relatórios em formato `.xlsx` ou `.csv`, organizados por diretório conforme configuração no arquivo `.env`.

---

## 📋 Visão Geral

O script conecta-se ao banco de dados SAP HANA, executa um conjunto de consultas SQL pré-definidas e salva os resultados como arquivos Excel (`.xlsx`) ou CSV (quando o volume ultrapassa o limite do Excel de 1.048.576 linhas). Os arquivos são nomeados automaticamente com a data de execução.

---

## 🗂️ Estrutura do Projeto

```
automacao_relatorio/
├── Automacao.py                    # Script principal de extração e geração de relatórios
├── test_query_atributos_lote.py    # Script de validação de query via EXPLAIN PLAN no HANA
├── requirements.txt                # Dependências Python do projeto
├── .env                            # Variáveis de ambiente (credenciais e caminhos) ⚠️ não versionar
└── Validacao/                      # Diretório com subpastas de saída por categoria
    ├── 02. HEIJUNKA COMERCIAL/
    ├── 04. KPIS/
    ├── 06. DEVOLUÇÕES/
    ├── BD_POSIÇÕES_ESTOQUE/
    └── LOGS/
```

---

## ⚙️ Configuração

### Pré-requisitos

- Python 3.8+
- Acesso à rede onde o SAP HANA está hospedado
- Credenciais de acesso ao banco SAP HANA

### Instalação das dependências

```bash
pip install -r requirements.txt
```

### Arquivo `.env`

Crie (ou edite) o arquivo `.env` na raiz do projeto com as seguintes variáveis:

```env
# Credenciais do SAP HANA
HANA_HOST=seu_host_hana
HANA_PORT=30015
HANA_USER=seu_usuario
HANA_PASSWORD=sua_senha

# Schema do banco de dados
HANA_SCHEMA=SBOGPHARMA

# Caminhos de saída dos relatórios
PATH_VENDAS=./Vendas
PATH_ESTOQUE_BD=./Estoque
PATH_ESTOQUE_POSICAO=./Estoque_Posicao
PATH_DEVOLUCOES=./Devolucoes
PATH_LOGS=./Logs
PATH_ESTOQUE_BD_DEVOLUCOES=./Devolucoes
PATH_PROJ=./Projetos
```

> ⚠️ **Nunca versione o arquivo `.env`** — ele contém credenciais sensíveis.

---

## 📑 Relatórios Gerados

| # | Nome do Relatório | Destino | View/Tabela SAP HANA |
|---|---|---|---|
| 1 | Boletim de Vendas - Versão 3 | `PATH_VENDAS` | `VW_RELATORIO_BONIFICACOES_DEVOLUCOES_VOLUMES` |
| 2 | Estoque Geral | `PATH_ESTOQUE_BD` | `PBI_FAT_BEAS_OP` |
| 3 | Estoque Por Lote | `PATH_ESTOQUE_BD` | `View Estoque` |
| 4 | Estoque Geral Com e Sem Lotes Com Posição | `PATH_ESTOQUE_POSICAO` | `VW_ESTOQUE_COM_E_SEM` |
| 5 | Acompanhamento Devolução | `PATH_DEVOLUCOES` | `VW_ACOMPANHAMENTO_DEVOLUCAO` |
| 6 | Log de Bloqueio de Lotes | `PATH_LOGS` | `VW_RASTREABILIDADE_LOTES` |
| 7 | IBT1 (Devoluções 01.08) | `PATH_DEVOLUCOES` | `IBT1` |
| 8 | Processo Dev Aberto v2 | `PATH_DEVOLUCOES` | `VW_PROC_DEVOLUCAO` |
| 9 | Relatório de Pedidos de Devolução V3 | `PATH_DEVOLUCOES` | `VW_RELATORIO_PEDIDOS_V3` |
| 10 | Atributos de Lote | `PATH_ESTOQUE_BD` | `VW_ATRIBUTOS_LOTE_VEICULOS` |

Os arquivos são salvos com o padrão de nomenclatura:

```
{Nome do Relatório}_{AAAA-MM-DD}.xlsx
```

ou, quando ultrapassar 1.048.000 linhas:

```
{Nome do Relatório}_{AAAA-MM-DD}.csv
```

---

## 🚀 Como Executar

```bash
python Automacao.py
```

O script irá:
1. Conectar ao SAP HANA com as credenciais do `.env`
2. Executar cada consulta SQL com download em blocos de 100.000 linhas (com barra de progresso)
3. Salvar os arquivos nos diretórios configurados
4. Fechar a conexão ao finalizar

### Exemplo de saída no terminal

```
Conectando ao SAP HANA externo...
Conexão estabelecida com sucesso!

--- Processando: Boletim de Vendas - Versão 3 ---
Executando SQL no SAP HANA e baixando dados...
Baixando dados: 3it [00:05,  1.82s/it]

Download concluído! A consulta retornou 320.000 linhas e 42 colunas.
Salvando XLSX em: ./Vendas/Boletim de Vendas - Versão 3_2026-10-21.xlsx ...
Sucesso!

==========================================
Processamento concluído para todos os itens.
==========================================
```

---

## 🧪 Validação de Queries

O arquivo `test_query_atributos_lote.py` permite validar a sintaxe e semântica de uma query SQL diretamente no SAP HANA via `EXPLAIN PLAN`, sem executar a consulta completa.

```bash
python test_query_atributos_lote.py
```

---

## 📦 Dependências

As principais bibliotecas utilizadas (ver `requirements.txt`):

| Biblioteca | Finalidade |
|---|---|
| `hdbcli` | Driver de conexão com o SAP HANA |
| `pandas` | Manipulação e exportação de dados |
| `python-dotenv` | Leitura de variáveis do `.env` |
| `tqdm` | Barra de progresso durante o download |
| `xlsxwriter` | Geração de arquivos `.xlsx` |
| `openpyxl` | Engine alternativa para Excel |

---

## 🔒 Segurança

- Jamais exponha o arquivo `.env` em repositórios públicos.
- Adicione `.env` ao seu `.gitignore`:

```gitignore
.env
.venv/
```

---

## 🛠️ Comportamento Especial

| Situação | Comportamento |
|---|---|
| Pasta de destino inexistente | Criada automaticamente pelo script |
| Resultado > 1.048.000 linhas | Salvo como `.csv` (separador `;`, encoding `utf-8-sig`) |
| Resultado vazio | DataFrame vazio salvo normalmente |
| Erro de conexão | Script encerra com mensagem de erro e código de saída `1` |
| Erro em relatório específico | Erro logado no terminal; os demais relatórios continuam sendo processados |
