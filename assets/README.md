# Estrutura modular do painel

Esta pasta guarda os arquivos extraidos do HTML original para facilitar manutencao.

## CSS

- `css/base.css`: reset e variaveis visuais.
- `css/layout.css`: estrutura principal, sidebar, topbar, cards, formularios, setup, erro, overview e ambientes.
- `css/criar-campo.css`: tela de criacao de campo.
- `css/shared.css`: estilos compartilhados entre paineis.
- `css/pipelines.css`: tela de pipelines e estagios.
- `css/campos.css`: consulta, edicao e exclusao de campos.
- `css/massa.css`: criacao em massa e importacao CSV/XLSX.
- `css/card-editor.css`: configuracao visual do card.

## JavaScript

- `js/core/api.js`: estado global basico e chamadas Bitrix24.
- `js/core/utils.js`: helpers compartilhados, labels, HTML escape e toasts.
- `js/core/connection.js`: teste de conexao e overview.
- `js/core/navigation.js`: navegacao lateral e acao do topbar.
- `js/features/ambientes.js`: setup e gerenciamento de ambientes.
- `js/features/criar-campo.js`: criacao de campos individuais.
- `js/features/pipelines.js`: pipelines e estagios.
- `js/features/campos.js`: consulta, exportacao, edicao e exclusao de campos.
- `js/features/massa.js`: criacao em massa, CSV e XLSX.
- `js/features/card-editor.js`: editor de configuracao do card.
- `js/app.js`: chamada final de inicializacao.

O arquivo `Painel de Implementação Bitrix24_V1.3_modular.html` carrega estes arquivos mantendo o HTML original praticamente intacto.