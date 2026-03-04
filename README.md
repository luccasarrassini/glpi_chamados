# glpi_chamados
Aplicativo desktop para operacoes em lote no GLPI via planilha (`.ods`, `.xlsx`, `.xls`).

## O que o app faz
- Importa chamados em lote.
- Fecha chamados em lote.
- Adiciona solucao em lote sem fechar chamados.
- Valida IDs de usuario/categoria/localizacao na API (opcional).
- Mostra preview (primeiras 100 linhas), progresso e log.

## Formato das planilhas
O app normaliza nomes de colunas para minusculo e sem espacos extras.

### 1) Importar chamados
Colunas obrigatorias:
- `titulo`
- `descricao`
- `categoria_id`
- `localizacao_id`

Colunas opcionais:
- `tecnico_id`
- `requerente_id`
- `tipo`

Regras:
- `titulo` nao pode ser vazio.
- `categoria_id` e `localizacao_id` devem ser numericos.
- Se `tecnico_id` e/ou `requerente_id` forem enviados, devem ser numericos.
- Se `tecnico_id` e `requerente_id` estiverem vazios (ou sem coluna), o chamado e criado sem atribuicao.
- `tipo` aceita `incidente` ou `requisicao` (tambem `1` ou `2`).
- Se `tipo` estiver vazio, o padrao e `incidente` (`1`).

### 2) Fechar chamados
Colunas obrigatorias:
- `ticket_id`

Coluna opcional:
- `solucao`

Regras:
- Se `solucao` existir e tiver valor, ela e enviada antes do fechamento.

### 3) Solucionar chamados (sem fechar)
Colunas obrigatorias:
- `ticket_id`
- `solucao`

Regras:
- `ticket_id` deve ser numerico.
- `solucao` nao pode ser vazia.
- Esta rotina apenas adiciona solucao, sem alterar status do ticket.

## Dependencias
Principais libs de planilha usadas pelo app:
- `openpyxl` (`.xlsx`, `.xlsm`)
- `xlrd` (`.xls`)
- `odfpy` (`.ods`)

## Relatorios gerados
Ao finalizar a importacao de chamados, o app salva automaticamente:
- Planilha Excel com os `ticket_id` criados.
- Arquivo de log texto da execucao.

Local padrao:
- `Documentos\logchamados`.
