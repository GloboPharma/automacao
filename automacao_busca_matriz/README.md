# Consulta da Matriz de Acessos x Cargos (INF-000007)

App Streamlit para consultar quais acessos cada cargo possui.

## Como rodar
```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python -m streamlit run app.py --server.address localhost
```

Use `python -m streamlit` (e não `streamlit.exe`): em alguns computadores o Windows
bloqueia o executável do `.venv` com "Acesso negado". O `--server.address localhost`
mantém o app acessível só neste computador, pois a matriz lista grupos de acesso internos.

## Modos de consulta
- **Cargo**: escolha departamento e cargo e veja os sistemas e perfis de acesso.
- **Sistema**: veja quais cargos têm acesso e com qual perfil.
- **Comparar cargos**: diferenças de acesso entre dois cargos.
- **Busca livre**: texto em cargo, sistema, perfil ou departamento, com filtros.

Todos os resultados podem ser baixados em CSV ou Excel.

## Dados
O app lê **somente a cópia local** em `dados/`.

A pasta `dados/` já vem vazia no repositório, mas a planilha **não é versionada**
(`dados/*.xlsx` está no `.gitignore`). Depois de clonar, copie o arquivo
`INF-000007-MATRIZ DE ACESSOS X CARGOS.xlsx` (fornecido pela TI) para dentro de `dados/`
**antes de rodar o app**, mantendo exatamente esse nome. Sem ele, o app mostra
"Planilha não encontrada".

Para atualizar a matriz, substitua o arquivo em `dados/` e recarregue a página do navegador.

## Observações
- A chave de um cargo é departamento + cargo (o mesmo cargo existe em vários departamentos).
- Cargo repetido no mesmo departamento (colunas distintas) aparece com o sufixo `[col. XX]`.
- Células `N/A` ou vazias são tratadas como "sem acesso".

## Testes
```
python -m pytest
```
