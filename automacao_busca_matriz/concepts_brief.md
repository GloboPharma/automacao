# Consulta da Matriz de Acessos x Cargos: avaliação e evolução

* **Author:** FREDERICO GUSTAVO MAGALHAES
* **Date:** 2026-11-05

## Overview

O app é uma ferramenta Streamlit local e somente leitura. Ele lê a planilha INF-000007 (sistemas nas linhas, cargos nas colunas) e oferece quatro consultas: por cargo, por sistema, comparação entre dois cargos e busca livre, com exportação para CSV e Excel.

Objetivos:
- Reduzir o tempo para localizar quais acessos cada cargo tem.
- Reduzir os erros da TI ao criar usuários nos ambientes.

A planilha é a fonte de verdade. Quando ela muda, o arquivo inteiro é substituído em `dados/`.

Metas de crescimento:
- Mais pessoas usando o app.
- Possivelmente, cadastrar novas atribuições de acesso pelo próprio app.

Sucesso significa dados corretos e mais usuários.

## Analysis

Os revisores (7 lentes) concluíram que o desenho de base está certo: somente leitura, planilha como fonte de verdade e tabela em formato longo. O maior risco é o app mostrar dados errados com aparência de corretos. Isso anula o objetivo de reduzir erros na criação de usuários. A ordem recomendada é: primeiro confiar nos dados, depois publicar com segurança, e só então decidir sobre edição.

Os pontos abaixo estão em ordem de gravidade. Os revisores não abriram a planilha real, então afirmações sobre linhas ocultas, células mescladas e fórmulas sem valor em cache vêm só da leitura do código e precisam ser verificadas.

### 1. Layout fixo sem validação (bloqueia confiança nos dados)
- O loader assume posições fixas: departamento na linha 5, cargo na linha 6, dados a partir da linha 7 e da coluna G. Ele também sempre lê a primeira aba.
- Se alguém inserir uma linha ou coluna, renomear um cabeçalho ou adicionar uma aba na frente, o app lê as células erradas sem dar erro.
- Linhas sem nome de sistema e colunas sem nome de cargo são descartadas sem aviso.
- Correção proposta: validar antes de aceitar o arquivo.
  - Conferir os rótulos esperados.
  - Exigir departamento em toda coluna de cargo.
  - Comparar contagens com a carga anterior.
  - Detectar sistemas duplicados.
  - Se a validação falhar, continuar servindo a última versão boa e mostrar um aviso com o motivo.

### 2. Sem autenticação (bloqueia a publicação)
- A matriz lista nomes de grupos de acesso internos, como `G_PCP_RW`.
- Hoje a única proteção é rodar em `localhost`. Publicar remove essa proteção.
- É preciso login (SSO da empresa, proxy reverso ou login OIDC do Streamlit), política para os botões de download e registro de acessos.
- Ideia de segmentação: todos veem o mapa cargo x sistema, e só a TI vê nomes de perfil e de grupo.

### 3. Edição cria uma segunda fonte de verdade (condicional à meta de cadastrar atribuições)
- Como a planilha é substituída inteira, qualquer edição feita no app é sobrescrita na próxima atualização ou passa a divergir dela.
- Editar exigiria:
  - identificadores estáveis para cargos e sistemas (hoje não existem);
  - um armazenamento diferente do xlsx;
  - login e papéis;
  - trilha de auditoria (quem, o quê, quando, antes e depois, motivo);
  - regras de conciliação a cada nova versão da planilha;
  - controle de edição simultânea.
- Gravar de volta no xlsx com openpyxl perderia formatação, mesclagens e validações, e conflita com quem estiver com o arquivo aberto.
- Não comece a edição antes dos itens 1, 5 e 2.

### 4. Testes não protegem nenhuma mudança
- Os testes rodam só contra a planilha real, que não é versionada, e são pulados quando ela falta. Um clone novo ou uma CI "passa" sem testar nada.
- Eles fixam números exatos (248 cargos, 32 departamentos, 101 sistemas, `[col. AT]`), então quebram a cada atualização legítima.
- Não há teste para:
  - mesclagem de departamentos;
  - N/A;
  - células com várias linhas;
  - cargos duplicados;
  - falha de layout;
  - a lógica de comparação, que está embutida em `app.py` e não pode ser importada sem o Streamlit.
- Correção proposta: montar uma planilha sintética pequena nos testes, escrever testes que fixam o comportamento atual antes de qualquer refatoração e mover a classificação da comparação para uma função pura.

### 5. Identidade instável do cargo
- O sufixo `[col. XX]` é a letra da coluna no Excel. Se uma coluna for adicionada ou removida, o mesmo cargo ganha outro rótulo. Isso quebra links, arquivos exportados, referências em chamados e o teste.
- Para o usuário, o sufixo não diz qual duplicata vale para a pessoa que está sendo criada. Isso pode levar a acesso errado.
- Correção proposta:
  - Perguntar ao dono da planilha por que existem cargos repetidos no mesmo departamento. Podem ser variantes reais (nível, unidade) ou erros.
  - Separar um identificador interno estável do rótulo exibido.
  - Desambiguar por um atributo real da planilha, não pela letra da coluna.

### 6. Sistemas sem acesso somem, e a planilha é lida duas vezes
- A tabela longa só tem linhas com perfil. Um sistema que nenhum cargo acessa não aparece em "Sistema". O usuário não distingue "ninguém tem acesso" de "não está na matriz".
- `carregar_matriz` e `listar_cargos` abrem a planilha duas vezes e repetem a lógica de departamentos. As duas cópias podem divergir.
- `sel.iloc[0]` no modo Sistema falha com `IndexError` se a seleção for vazia. Isso aparece assim que o seletor for corrigido de forma direta.
- A planilha não é fechada em `try/finally`. No Windows, um arquivo aberto pode bloquear a próxima cópia por cima.
- Correção proposta: uma leitura única que devolve três tabelas (cargos, sistemas, atribuições), com teste provando que a saída é igual à do loader atual.

### 7. Atualização manual e sem como saber se os dados estão atuais
- A legenda mostra a data do arquivo, não a revisão da matriz.
- O cache usa só a data de modificação. Uma cópia que preserve o horário antigo serve dados velhos. Um arquivo copiado pela metade pode mostrar erro técnico ao usuário.
- O nome do arquivo é exato. Em Linux, onde uma versão publicada provavelmente rodaria, o `x` minúsculo do nome real não seria encontrado.
- Não há versão anterior guardada, diferença entre versões nem volta atrás.
- Correção proposta, em ordem:
  - cache por conteúdo (hash) mais data e tamanho;
  - mostrar a revisão e a data da matriz;
  - escolher o arquivo `INF-000007*.xlsx` mais recente;
  - trocar o arquivo só depois de validar, mantendo o anterior;
  - gerar um relatório de mudanças ("3 cargos novos, 5 células alteradas").

### 8. Ambiguidade de significado na tela (maior fonte provável de erro na TI)
- Célula vazia e `N/A` viram "sem acesso", e o modo Cargo diz "não possui acesso cadastrado". Isso pode ser uma coluna ainda não preenchida, não uma ordem de não conceder nada.
- Termos sem definição: impacto, perfil, RO/RW, tipo de sistema, POPS.
- Não se sabe se vários perfis na mesma célula são todos necessários ou alternativas.
- A cor é o único indicador de significado.
- Rótulos enganosos:
  - "Limpar filtros" na comparação restaura o padrão em vez de limpar;
  - "Resultados" e "Perfis atribuídos" contam linhas de perfil, não pessoas nem sistemas.
- É possível comparar um cargo com ele mesmo sem aviso.

### 9. Leitura estreita descarta ou distorce dados
- Perfis só são separados por quebra de linha. "A; B" ou "A, B" vira um perfil só.
- Valores vazios como "N.A." ou "SEM ACESSO" não são reconhecidos, e marcadores como "X" viram perfil. Um perfil real chamado "NA" seria descartado.
- Perfis duplicados inflam contagens e podem gerar um falso "perfil diferente".
- Com `data_only=True`, uma fórmula sem valor em cache vira vazio, ou seja, "sem acesso". Um erro como `#REF!` aparece como perfil.
- Hipóteses ainda sem verificação na planilha real: lacuna na linha de departamento, células de sistema mescladas na vertical, coluna com cabeçalho vazio mas com dados, linhas e colunas ocultas.
- Correção proposta: registrar e mostrar quantas linhas, colunas e células foram descartadas, e mostrar valores curtos não reconhecidos num painel de qualidade dos dados.

### 10. Falta de rastreabilidade e telemetria
- Nada registra qual arquivo foi carregado, seu hash, as contagens ou o que foi ignorado. Uma resposta errada não pode ser reconstruída.
- Os resultados não apontam de volta para a linha e a coluna da planilha.
- Sem telemetria de uso (modo, filtros, buscas sem resultado), não dá para medir "mais pessoas usando".
- Correção proposta: o loader devolve um relatório de validação, cada registro guarda linha e coluna de origem, e um painel "Sobre os dados" mostra o resumo.

### 11. `app.py` monolítico
- Tema, formatação, consultas e quatro modos estão no mesmo arquivo, executados na importação. Textos como "IMPACTO DIRETO" se repetem como literais.
- Há duas funções que fazem a mesma coisa com colunas de lista (`_para_arquivo` e `_quebrar_listas`).
- Há divergência entre as lentes. Uma lente defende separar consultas puras, um render por modo e constantes compartilhadas. A lente de simplicidade acha que, com cerca de 275 linhas, basta extrair a comparação e unificar o helper de listas, sem novos pacotes, bancos de dados ou sistemas de configuração.

### 12. Desempenho
- Com cerca de 250 cargos por 100 sistemas, não há problema para dezenas de usuários. A Busca livre envia todo o resultado ao navegador, sem limite de linhas.
- Não otimizar antes de medir. Meça a carga a frio e a quente. A meta "localizar mais rápido" hoje não tem medida de base.

### 13. Defeitos menores
- Modo Sistema usa a primeira linha para impacto, tipo e área. Se esses atributos variarem entre linhas, a primeira vence sem aviso.
- A ordenação ignora acentos e diferença de maiúsculas ("É" fica depois de "Z"). A lista de cargos segue a ordem da planilha.
- Opções de filtro podem incluir texto vazio. `index=min(cargo_inicial, len(lista)-1)` quebra com lista vazia.
- A Busca livre diferencia acentos ("gestao" não acha "Gestão") e trata uma frase como um texto literal ("analista qualidade" não acha nada). Ela também não pesquisa área dona nem POPS.
- Exportações:
  - valores que começam com `=`, `+`, `-` ou `@` abrem como fórmula no Excel e devem ser escapados;
  - o nome de arquivo pode ficar vazio;
  - as chaves dos botões de download podem colidir;
  - as exportações não registram a data da matriz.
- Arquivos temporários `~$` do Excel em `dados/` não são tratados.
- `streamlit.log` não versionado e sem `*.log` no `.gitignore`.
- A docstring recomenda `streamlit run app.py`, o que expõe o app na rede. O README orienta o contrário.
- O README não descreve o contrato de layout, a versão mínima do Python (3.10+) nem as limitações.

### Contradições de projeto e como resolvê-las

| Contradição | Resolução proposta | Custo |
| :-- | :-- | :-- |
| Consulta somente leitura x edição controlada | O visualizador continua somente leitura. Um componente separado grava pedidos de mudança num log só de acréscimo, com aprovação, e a TI aplica na planilha. Versão mínima: link de chamado ou e-mail preenchido com cargo e sistema. A alternativa é inverter e tornar um banco a fonte de verdade, com a planilha gerada a partir dele. | Aprovador a manter. A inversão exige migração e que o time dono pare de editar a planilha à mão. |
| Abertura (mais usuários) x sensibilidade | Segmentar por público, SSO fora do app, negar por padrão e registrar acessos. | Integração com SSO. Restrição demais leva as pessoas de volta a perguntar à TI. |
| Layout fixo x liberdade de quem edita a planilha | Detecção dos cabeçalhos por texto mais validação que, se falhar, mantém a última versão boa. | Detecção quebra se cabeçalhos forem renomeados. Posições fixas são mais simples para um modelo estável. |
| Nome legível x identidade estável | ID interno separado do rótulo, mais dados de origem limpos. | Pode exigir uma linha de ID na planilha. |
| Fatos compactos x dimensões completas | Modelo com tabelas de cargos, sistemas e atribuições, e estado explícito para "sem acesso". | Baixo. |
| Atualização manual x frescor | Buscar do local oficial depois da validação, com cópias versionadas. | Conta de serviço e retenção de versões. |
| Localizar rápido x transcrever certo | Lista de provisionamento por cargo, pronta para copiar, com a data da matriz. | A automação futura aumenta o risco se o dado estiver errado. |

A ideia de produto mais citada: uma lista de provisionamento por cargo, pronta para copiar, ordenada por impacto e com a data da matriz. Ela é basicamente formatação sobre `_perfis_por` e `_download`. Combine-a com uma visão "o que mudou desde a última versão".

### Resultado final ideal
O dono da matriz publica a planilha uma vez, no local oficial. O sistema a busca, carrega só se for válida e, se não for, recusa com explicação e a diferença. Cada cargo tem identidade estável. Cada pessoa vê só o que seu papel exige, pelo login da empresa. A TI escolhe um cargo (ou abre a partir de um chamado) e recebe uma lista copiável e versionada. Pedidos de mudança passam por aprovação registrada e chegam à mesma fonte, então nunca existe uma segunda verdade. Toda tela informa a versão da matriz e a idade dos dados.

## Decisões do autor (rodada 2)

- **Edição de atribuições:** fora de escopo. O app é somente leitura por enquanto.
- **Cargos duplicados:** podem ser erro de cadastro na planilha. O app deve tratá-los como alerta de qualidade dos dados, em vez de só acrescentar `[col. XX]` sem explicação.
- **Troca da planilha:** deve ser feita dentro do app, e não copiando o arquivo à mão para `dados/`. Isso é um recurso novo: enviar a nova planilha pela tela, validar e só então substituir a versão em uso. Ele substitui a atualização manual dos itens 1 e 7. Quando a validação falha, o app mantém a versão anterior e mostra o motivo.
- **Acesso:** local e de um único usuário (o autor). Login e segmentação (item 2) ficam adiados. Eles voltam a ser obrigatórios antes de qualquer publicação, e a tela de envio da planilha precisará de controle de quem pode usá-la.

## Goals vs. Non-Goals

* **In-Scope Goals:**
  - Enviar a nova planilha pelo app, validar o layout, mostrar um resumo do que mudou e substituir a versão em uso mantendo a anterior (atende o item 1 e parte do item 7).
  - Avisar sobre cargos duplicados no mesmo departamento como possível erro da planilha, listando onde estão (item 5).
  - Garantir que o app mostre dados corretos ou avise claramente quando não puder (itens 1, 4, 6, 7, 9, 10).
  - Dar identidade estável aos cargos (item 5).
  - Esclarecer o significado dos dados na tela (item 8).
  - Lista de provisionamento por cargo.
  - Segurança mínima antes de publicar para outros usuários (item 2): adiada enquanto o uso for local e único.
  - Medir o sucesso: tempo para localizar, usuários distintos por semana, erros de criação de usuário.
* **Out-of-Scope Non-Goals:**
  - Edição direta de acessos dentro do app (decisão do autor: somente leitura por enquanto).
  - Login e segmentação por público enquanto o uso for local e único.
  - Otimização de desempenho sem medida de base.
  - Automação de provisionamento nos ambientes.
  - Reescrita completa em outra tecnologia (sem restrição de tecnologia hoje, mas nada indica essa necessidade).

## Alternatives Weighed

- **Posições fixas com falha explícita x detecção de cabeçalhos por texto.** A maioria das lentes prefere detecção de cabeçalhos mais validação. Posições fixas são mais simples se a TI raramente muda o layout. Decisão do dono: pendente.
- **Gravar edições no xlsx com openpyxl.** Rejeitado pelos revisores: perde formatação, mesclagens e validações, e conflita com quem tem o arquivo aberto.
- **Edição direta no app.** Adiada: cria segunda fonte de verdade e exige login, papéis e auditoria.
- **Pedidos de mudança com log e aprovação (caminho a).** Recomendado como primeiro passo se a demanda por cadastrar atribuições for real.
- **Banco de dados como fonte de verdade, planilha gerada (caminho b).** Só faz sentido se o time dono parar de editar a planilha à mão.
- **Dividir `app.py` em pacote completo.** A lente de simplicidade rejeita: com cerca de 275 linhas, basta extrair a comparação e unificar o helper de listas.

## Key Decisions & Staging for Next Steps

Decididas:
- O app é somente leitura (sem edição de atribuições).
- A troca da planilha acontece dentro do app, com validação.
- O uso é local e de um único usuário por enquanto.
- Cargos duplicados são tratados como possível erro da planilha, com alerta.

- O layout da planilha não costuma mudar, então a validação usa **posições fixas com falha clara**, sem detecção de cabeçalhos por texto.
- O que muda é a **entrada de novos cargos** dentro das áreas (novas colunas de cargo). A validação deve aceitar mais colunas de cargo, desde que cada uma tenha departamento e nome. Aumento de cargos é uma mudança normal e aparece no relatório de mudanças, não como erro.
- Falha na validação **bloqueia** o envio: a versão anterior continua em uso e o app mostra o motivo.

Decisões pendentes do dono:
2. Quando uma duplicata é erro e quando é variante real? Existe atributo na planilha que as distingue? Com o bloqueio decidido, falta definir se uma duplicata bloqueia o envio ou só gera alerta.
3. A planilha tem número de revisão para exibir?
4. Quem corrige a planilha quando o envio é bloqueado, e com que frequência ela muda?
5. O que significam impacto, perfil, RO/RW e POPS? Vários perfis numa célula são todos necessários ou alternativos?
6. O que significa célula vazia, em contraste com `N/A`?
7. Antes de publicar para outros: quem pode ver os dados, e quem pode enviar uma nova planilha?

Sequência recomendada:
1. Planilha de teste sintética e testes que fixam o comportamento atual, sem mudar o comportamento.
2. Leitura única com três tabelas, provada igual à antiga. Validação de layout com falha segura, aviso de dados descartados e colunas de origem.
3. Identidade estável do cargo, depois de confirmar com a TI a causa das duplicatas.
4. Comparação e consultas em funções puras com testes. Cache por conteúdo, versão exibida e **envio da nova planilha pelo app**: valida, troca de forma atômica, guarda a versão anterior e mostra o relatório de mudanças.
5. Semântica da interface (vazio x N/A, glossário, legenda, rótulos). Lista de provisionamento. Medir tempo de carga.
6. Antes de publicar: SSO ou proxy, definição do público, política de download, registro de acessos, caminho do arquivo por configuração.
7. Edição, só depois dos itens anteriores e da decisão sobre a fonte de verdade, começando por pedidos de mudança.
