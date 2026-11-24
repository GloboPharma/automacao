# Gerador DTW – Transferência de Estoque por Lote (SAP B1 HANA)

Lê a planilha de ajuste, consulta o saldo dos lotes no SAP Business One (HANA)
e gera os arquivos de importação do DTW **somente com as linhas aprovadas**:

| Arquivo | Objeto DTW |
|---|---|
| `OWTR - Transferência de Estoque.txt` / `.xlsx` | Cabeçalho (StockTransfer) |
| `WTR1 - Linhas Transferencia de Estoque.txt` / `.xlsx` | Linhas |
| `BTNT - Lotes.txt` / `.xlsx` | Lotes (BatchNumbers) |
| `Relatorio_Validacao.xlsx` | Todas as linhas, saldo e motivo de bloqueio |

O script **só faz SELECT** no banco (`OITM`, `OBTN`, `OBTQ`, `OWHS`). Nada é gravado no SAP.

## 1. Instalação (uma vez)

Pré-requisito: Python 3.9+ ([python.org](https://www.python.org/downloads/) — marque *Add Python to PATH*).

```bat
cd C:\caminho\da\pasta
pip install -r requirements.txt
```

## 2. Configuração – arquivo `.env`

Na mesma pasta do script, preencha o `.env`:

```ini
SAP_HANA_HOST=10.0.0.10
SAP_HANA_PORTA=30015
SAP_HANA_USUARIO=USR_LEITURA
SAP_HANA_SENHA=
SAP_HANA_SCHEMA=SBO_GLOBO_TESTE
```

- Deixe `SAP_HANA_SENHA` vazio: a senha é pedida no terminal.
- Para alternar entre **teste** e **produção**, troque só `SAP_HANA_SCHEMA`.
- O `.env` contém dados de acesso: não compartilhe nem versione.

## 3. Planilha de entrada

Qualquer nome de arquivo. São lidas as abas que tenham as colunas:

| Coluna | Uso |
|---|---|
| `Código do Item` | ItemCode |
| `Nome do Item` | Descrição (opcional) |
| `Lote` | DistNumber |
| `Qtd 01.08` | Quantidade; o código no título (`01.08`) é o **depósito de origem** |
| `Depósito` | **Depósito de destino** |

Abas de resultado de consulta (com colunas `Saldo`, `Situacao` etc.) são ignoradas.

## 4. Execução

```bat
python gerar_dtw_transferencia.py "Ajuste Estoque PA v2-validacao.xlsx"
```

Saída numa subpasta `DTW_aaaammdd_hhmm` ao lado da planilha.

### Opções

| Opção | Efeito |
|---|---|
| `--schema BANCO` | Sobrepõe o schema do `.env` |
| `--origem 01.08` | Força o depósito de origem |
| `--formato txt\|xlsx\|ambos` | Formato dos arquivos DTW (padrão `ambos`) |
| `--documento-unico` | Uma só transferência (padrão: uma por par origem→destino) |
| `--serie 21` / `--data 20261123` | Série e data do documento |
| `--comentario "texto"` | Observação da OWTR |
| `--bloquear-qualquer-reserva` | Bloqueia lote com qualquer quantidade comprometida |
| `--bloquear-status` | Bloqueia lotes Inacessível / Bloqueado (padrão: só alerta) |
| `--decimal ","` | Separador decimal no `.txt` (padrão `.`) |
| `--codificacao utf-16` | `.txt` em Unicode (padrão ANSI/cp1252) |
| `--saida PASTA` | Pasta de saída |

## 5. Regras de bloqueio

A linha **não** entra na transferência quando:

1. Lote sem saldo no depósito de origem;
2. Saldo menor que a quantidade pedida;
3. Saldo reservado/comprometido (`OBTQ.CommitQty`) deixa o disponível menor que o pedido;
4. Lote não existe para o item, item inexistente ou não controlado por lote;
5. Quantidade zero/vazia;
6. Depósito de origem/destino inexistente ou origem = destino;
7. (com `--bloquear-status`) lote Inacessível ou Bloqueado.

Se o mesmo lote/origem aparece em várias linhas ou abas, as quantidades são **somadas** antes de comparar com o saldo.

## 6. Importação no DTW

1. Confira o `Relatorio_Validacao.xlsx` (aba *Resumo* e coluna *Motivo*).
2. DTW → *Estoque e Produção* → **Transferência de estoque** (oStockTransfer).
3. Associe `OWTR…txt` ao cabeçalho, `WTR1…txt` a *StockTransfer_Lines* e `BTNT…txt` a *BatchNumbers*.
4. Rode primeiro em **Simular** e depois na base de **teste**; só então em produção.
5. Entre a validação e a importação o saldo pode mudar — gere os arquivos e importe em seguida.
