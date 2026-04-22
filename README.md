# Painel de Implementacao Bitrix24

Painel web estatico para apoiar a configuracao e manutencao de recursos do Bitrix24, incluindo campos, criacao em massa, pipelines e configuracao de card.

## Como usar

1. Publique este projeto como site estatico.
2. Abra o painel no navegador.
3. Informe o webhook REST do Bitrix24 em `Webhook REST`.
4. Teste a conexao antes de criar ou alterar dados.

## Configuracao do webhook Bitrix24

Crie um webhook de entrada no Bitrix24 em **Desenvolvedor > Webhooks de entrada** e conceda apenas as permissoes necessarias para as operacoes que serao usadas no painel.

O webhook deve seguir o formato:

```text
https://seudominio.bitrix24.com.br/rest/1/token/
```

## Seguranca

O webhook informado no painel fica salvo apenas no navegador, usando `localStorage`. Ele nao e enviado para nenhum servidor externo alem das chamadas feitas diretamente para a API REST do Bitrix24.

Cuidados recomendados:

- Evite usar webhook de usuario administrador quando nao for necessario.
- Use permissoes minimas para o objetivo do painel.
- Revogue o webhook se ele for exposto, compartilhado por engano ou usado em computador nao confiavel.
- Nao publique tokens, senhas ou webhooks reais no repositorio.

## Publicacao

Este projeto e um app estatico. Para publicar no GitHub Pages, use a branch `main` e a pasta raiz do repositorio.
