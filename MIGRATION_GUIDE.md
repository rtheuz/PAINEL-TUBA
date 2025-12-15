# Guia de Migração - Unificação de Orçamentos e Pedidos

## Visão Geral

Este guia descreve as mudanças implementadas para unificar as abas "Orçamentos" e "Pedidos" em uma única aba chamada "Projetos", e como realizar a migração dos dados existentes.

## O Que Mudou

### Nova Estrutura - Aba "Projetos"

A nova aba "Projetos" possui 14 colunas que consolidam informações de orçamentos e pedidos:

1. **CLIENTE** - Nome do cliente
2. **DESCRIÇÃO** - Descrição do projeto
3. **RESPONSÁVEL CLIENTE** - Pessoa responsável do cliente
4. **PROJETO** - Número único do projeto (chave primária)
5. **VALOR TOTAL** - Valor total do orçamento
6. **DATA** - Data de criação
7. **PROCESSOS** - Tempo estimado por processo
8. **LINK DO PDF** - Link para o PDF do orçamento
9. **LINK DA MEMÓRIA DE CÁLCULO** - Link para memória de cálculo
10. **STATUS_ORCAMENTO** - Status do orçamento (Rascunho, Enviado, Convertido em Pedido, Expirado/Perdido)
11. **STATUS_PEDIDO** - Status do pedido (vazio, Preparação MP/CAD/CAM, Corte, Dobra, Processos Adicionais, Envio/Coleta, Finalizado)
12. **PRAZO** - Prazo de entrega
13. **OBSERVAÇÕES** - Observações do pedido
14. **JSON_DADOS** - Dados completos do formulário em JSON

### Lógica de Funcionamento

- **Novo projeto**: STATUS_ORCAMENTO = "Rascunho", STATUS_PEDIDO = vazio
- **Orçamento enviado**: STATUS_ORCAMENTO = "Enviado", STATUS_PEDIDO = vazio
- **Convertido em pedido**: STATUS_ORCAMENTO = "Convertido em Pedido", STATUS_PEDIDO = "Processo de Preparação MP / CAD / CAM"
- **Progresso do pedido**: STATUS_ORCAMENTO permanece "Convertido em Pedido", STATUS_PEDIDO muda conforme etapa

## Como Migrar os Dados

### Passo 1: Fazer Backup

**IMPORTANTE**: Antes de iniciar a migração, faça um backup completo da planilha!

1. Abra a planilha no Google Sheets
2. Vá em **Arquivo > Fazer uma cópia**
3. Nomeie como "TUBA_BACKUP_[DATA]"

### Passo 2: Executar a Função de Migração

1. No Google Sheets, vá em **Extensões > Apps Script**
2. No editor de scripts, localize a função `migrarDadosParaProjetosUnificados()`
3. Clique em **Executar** (ícone de play)
4. Autorize as permissões necessárias quando solicitado
5. Aguarde a conclusão da migração (você verá logs no console)

### Passo 3: Verificar a Migração

Após a migração, verifique:

1. A aba "Projetos" foi criada
2. Todos os orçamentos foram migrados
3. Os orçamentos "Convertidos em Pedido" têm o STATUS_PEDIDO preenchido
4. As abas "Orçamentos" e "Pedidos" originais permanecem como backup

### Passo 4: Testar o Sistema

1. Acesse a nova página de Projetos: `?page=projetos`
2. Verifique se todos os projetos aparecem corretamente
3. Teste os filtros de busca e status
4. Teste a edição de campos
5. Verifique se o Kanban continua funcionando corretamente
6. Verifique se o Dashboard mostra os contadores corretos

### Passo 5: Renomear Abas Antigas (Opcional)

Se tudo estiver funcionando corretamente, você pode renomear as abas antigas para indicar que são backup:

1. **"Orçamentos"** → **"Orçamentos_backup"**
2. **"Pedidos"** → **"Pedidos_backup"**

⚠️ **Não delete as abas antigas imediatamente!** Mantenha-as por algumas semanas como backup.

## Funcionalidades Implementadas

### 1. Validação de Projeto Duplicado

O sistema agora valida se um número de projeto já existe antes de criar um novo:

```javascript
verificarProjetoDuplicado(numeroProjeto)
// Retorna: { duplicado: boolean, linha: number, onde: string }
```

Esta validação é aplicada em:
- Criação de novos rascunhos
- Registro de novos orçamentos
- Interface do formulário (feedback visual imediato)

### 2. Página Projetos Unificada

Nova interface em `projetos.html` com:

- **Filtros avançados**: Busca por cliente/projeto/descrição, filtro por STATUS_ORCAMENTO e STATUS_PEDIDO
- **Edição inline**: Clique em campos editáveis para modificar diretamente
- **Gestão de status**: Selects para alterar status de orçamento e pedido
- **Indicadores visuais**: Cores diferentes para cada status
- **Responsivo**: Adaptado para desktop e mobile
- **Links rápidos**: Acesso direto a PDFs e memórias de cálculo

### 3. Funções Atualizadas

#### Code.js - Principais Mudanças

**Novas Constantes:**
```javascript
const SHEET_PROJ = ss.getSheetByName("Projetos");
const PROJETOS_NUM_COLUNAS = 14;
```

**Novas Funções:**
- `verificarProjetoDuplicado(numeroProjeto)` - Valida duplicidade
- `migrarDadosParaProjetosUnificados()` - Migra dados existentes
- `getProjetos()` - Retorna todos os projetos
- `atualizarProjetoNaPlanilha(linha, dataObj)` - Atualiza projeto
- `excluirProjeto(linha)` - Exclui projeto

**Funções Atualizadas:**
- `registrarOrcamento()` - Suporta nova estrutura
- `salvarRascunho()` - Valida duplicidade e usa nova estrutura
- `carregarRascunho()` - Carrega de ambas estruturas
- `getListaRascunhos()` - Lista de ambas estruturas
- `getKanbanData()` - Lê da aba Projetos
- `atualizarStatusKanban()` - Atualiza STATUS_PEDIDO
- `getDashboardStats()` - Conta da aba Projetos

### 4. Retrocompatibilidade

O sistema foi desenvolvido com retrocompatibilidade em mente:

- Todas as funções verificam primeiro se a aba "Projetos" existe
- Se não existir, usam a aba "Orçamentos" como fallback
- Isso permite uma transição gradual e segura
- As rotas antigas (`?page=orcamentos` e `?page=pedidos`) continuam funcionando

## Impacto no Kanban

O Kanban continua funcionando normalmente com a nova estrutura:

- **Coluna "Processo de Orçamento"**: Projetos com STATUS_ORCAMENTO em ('Rascunho', 'Enviado') e STATUS_PEDIDO vazio
- **Demais colunas de pedido**: Projetos com STATUS_PEDIDO correspondente ao nome da coluna
- Os logs de tempo real continuam sendo aplicados normalmente

## Impacto no Dashboard

O Dashboard foi atualizado para contar corretamente da nova estrutura:

- **Orçamentos**: Conta projetos que não foram convertidos nem perdidos
- **Pedidos**: Conta projetos com STATUS_PEDIDO não vazio
- **Kanban**: Conta projetos com STATUS_PEDIDO != "Finalizado"

## Solução de Problemas

### Erro: "Aba 'Projetos' não encontrada"

Se você vê este erro mas ainda não migrou os dados:
1. Execute a função `migrarDadosParaProjetosUnificados()` como descrito acima
2. Ou continue usando as abas antigas - o sistema tem fallback

### Projetos não aparecem na nova página

1. Verifique se a migração foi concluída com sucesso
2. Verifique os logs no Apps Script (Ver > Logs)
3. Recarregue a página (Ctrl+F5 ou Cmd+Shift+R)

### Erro ao editar projeto

1. Verifique suas permissões de usuário
2. Verifique se a linha ainda existe na planilha
3. Tente recarregar a página e editar novamente

### Kanban não mostra projetos corretos

1. Verifique se os STATUS_ORCAMENTO e STATUS_PEDIDO estão preenchidos corretamente
2. Verifique se a migração foi concluída com sucesso
3. Limpe o cache do navegador e recarregue

## Suporte

Se encontrar problemas durante a migração:

1. Consulte os logs no Apps Script (**Ver > Logs** ou **View > Logs**)
2. Verifique se há mensagens de erro no console do navegador (F12)
3. Reverta para o backup se necessário
4. Entre em contato com o suporte técnico

## Cronograma Recomendado de Migração

1. **Semana 1**: Fazer backup e executar migração em ambiente de teste
2. **Semana 2**: Validar todos os dados migrados e funcionalidades
3. **Semana 3**: Executar migração em produção
4. **Semana 4**: Monitorar sistema e resolver eventuais problemas
5. **Após 1 mês**: Renomear/arquivar abas antigas se tudo estiver OK

## Conclusão

Esta unificação simplifica significativamente o gerenciamento de projetos, eliminando a duplicação de dados e facilitando o acompanhamento do ciclo completo de orçamento até entrega. A nova estrutura é mais intuitiva e permite melhor rastreabilidade dos projetos.

Para dúvidas ou suporte adicional, consulte a documentação técnica ou entre em contato com a equipe de desenvolvimento.
