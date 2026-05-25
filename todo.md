# Portal Saúde de Atleta - TODO

## Implementação de Backend e Persistência

- [x] Criar tabelas no banco de dados para Antropometria, FPM e ISAK
- [x] Implementar rotas tRPC para salvar avaliações (createAntropometria, createFPM, createIsak)
- [x] Implementar geração automática de Excel ao finalizar cada avaliação
- [x] Implementar upload de Excel para S3 (nuvem)
- [x] Reorganizar campos ISAK com nova ordem (dobras primeiro, depois perímetros)
- [x] Adicionar novos campos ISAK: perímetro de tórax e coxa média
- [x] Integrar frontend com API para salvar dados automaticamente
- [x] Implementar notificações de sucesso/erro ao salvar dados
- [x] Criar interface de visualização de dados salvos
- [x] Testar fluxo completo de coleta e salvamento

## Melhorias de UX (Opcionais)

- [x] Adicionar indicador de sincronização com servidor
- [x] Implementar retry automático em caso de falha
- [x] Adicionar histórico de avaliações do usuário
- [x] Criar dashboard com resumo de dados coletados

## Testes

- [x] Testar coleta de Antropometria
- [x] Testar coleta de FPM
- [x] Testar coleta de ISAK com nova ordem
- [x] Testar geração de Excel
- [x] Testar upload para nuvem
- [x] Testar sincronização de dados

## Correções Solicitadas

- [x] Salvar TUDO que for preenchido em cada rodada (não apenas na rodada 3)
- [x] Implementar proteção contra atualização de página sem salvar
- [x] Adicionar botão para resetar dados com confirmação
- [x] Fazer teste completo com dados aleatórios para ID 0
- [x] Verificar se tudo está funcionando corretamente
- [ ] Publicar atualização no link
