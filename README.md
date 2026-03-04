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
- `tecnico_id`
- `requerente_id`

Regras:
- `titulo` nao pode ser vazio.
- IDs devem ser numericos.

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
