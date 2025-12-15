# Guia de Depuração - Problemas com Projetos

## Problemas Relatados
1. Projetos não aparecem na página /projetos
2. Modal de adicionar projeto aparece em branco
3. Contadores do dashboard estão zerados
4. Contador do Kanban mostra apenas quantidade de cards

## Como Depurar

### 1. Verificar a Aba "Projetos" na Planilha

Abra a planilha do Google Sheets e verifique:

**a) A aba "Projetos" existe?**
- Se SIM: Continue para o próximo passo
- Se NÃO: Execute novamente a função `migrarDadosParaProjetosUnificados()` no Apps Script

**b) A aba "Projetos" tem dados?**
- Abra a aba "Projetos"
- Verifique se há pelo menos 2 linhas (cabeçalho + dados)
- Se a aba está vazia ou só tem o cabeçalho, a migração não funcionou

**c) Os cabeçalhos estão corretos?**
Os cabeçalhos da linha 1 devem ser EXATAMENTE (incluindo acentos):
```
CLIENTE | DESCRIÇÃO | RESPONSÁVEL CLIENTE | PROJETO | VALOR TOTAL | DATA | PROCESSOS | LINK DO PDF | LINK DA MEMÓRIA DE CÁLCULO | STATUS_ORCAMENTO | STATUS_PEDIDO | PRAZO | OBSERVAÇÕES | JSON_DADOS
```

### 2. Verificar Logs do Apps Script

1. No Apps Script, vá em **Execuções** (menu lateral)
2. Execute manualmente a função `getProjetos()`
3. Verifique os logs:
   - `getProjetos: Sheet name=...` - deve mostrar "Projetos"
   - `getProjetos: lastRow=...` - deve ser maior que 1
   - `getProjetos: Headers=...` - deve mostrar os cabeçalhos
   - `getProjetos: Retornando X projetos` - deve mostrar quantos projetos foram encontrados

4. Execute manualmente a função `getDashboardStats()`
5. Verifique os logs:
   - `getDashboardStats: Aba Projetos encontrada, totalProjetos=...`
   - `getDashboardStats: Headers da aba Projetos: ...`
   - `getDashboardStats: Contagem final - orcamentos=..., pedidos=..., kanban=...`

### 3. Verificar Console do Navegador

1. Abra a página `/projetos` no navegador
2. Abra o Console do Desenvolvedor (F12)
3. Procure por mensagens de erro ou logs:
   - `Iniciando loadProjetos...`
   - `getProjetos success, recebido: ...`
   - `Tipo de data: ... Length: ...`
   - `renderTable: projetos.length= ...`
   - `renderTable: filtered.length= ...`

### 4. Problema: "Modal de Adicionar Projeto em Branco"

O botão "Novo Projeto" redireciona para a página `/orcamento` (formulário de orçamento), não é um modal.

**Se a página aparece em branco:**
1. Verifique o console do navegador (F12) para erros JavaScript
2. Verifique se o token de autenticação está válido
3. Tente fazer logout e login novamente
4. Verifique se a função `doGet` no Code.js está processando corretamente o `case 'orcamento'`

### 5. Problema: "Contadores do Dashboard Zerados"

**Causa mais provável:** A aba "Projetos" existe mas está vazia, ou os cabeçalhos estão incorretos.

**Solução:**
1. Verifique se a aba "Projetos" tem dados (veja passo 1.b acima)
2. Execute os logs do `getDashboardStats()` (veja passo 2 acima)
3. Se os logs mostram `totalProjetos=0`, a aba está vazia
4. Se os logs mostram `idxStatusOrc=-1` ou `idxStatusPed=-1`, os cabeçalhos estão errados

### 6. Re-executar a Migração

Se a aba "Projetos" está vazia ou com problemas:

1. No Apps Script, execute a função `migrarDadosParaProjetosUnificados()`
2. Verifique os logs:
   - Deve mostrar "Migrando dados de 'Orçamentos'..."
   - Deve mostrar quantas linhas foram migradas
   - Deve mostrar "Migração concluída com sucesso"
3. Se houver erros, copie-os e informe

### 7. Verificar Estrutura das Abas Antigas

Se você ainda tem as abas "Orçamentos" e "Pedidos" originais:

**Aba "Orçamentos" deve ter as colunas:**
```
CLIENTE | DESCRIÇÃO | Responsável Cliente | PROJETO | VALOR TOTAL | DATA | Processos | LINK DO PDF | LINK DA MEMÓRIA DE CÁLCULO | STATUS | PRAZO | JSON_DADOS
```

**Aba "Pedidos" deve ter as colunas:**
```
Cliente | Número do Projeto | Status | Observações | DESCRIÇÃO | Tempo estimado por processo | PRAZO
```

## Ações Corretivas Rápidas

### Se a aba "Projetos" não existe:
```javascript
// Execute no Apps Script
migrarDadosParaProjetosUnificados();
```

### Se a aba "Projetos" existe mas está vazia:
1. Delete a aba "Projetos" manualmente
2. Execute novamente: `migrarDadosParaProjetosUnificados()`

### Se nenhum projeto aparece mas você sabe que existem dados:
1. Verifique os logs conforme instruções acima
2. Copie os logs completos e envie para análise
3. Faça um screenshot da aba "Projetos" mostrando os cabeçalhos e primeira linha de dados

## Informações para Suporte

Se o problema persistir, forneça:
1. Screenshot da aba "Projetos" (cabeçalhos e primeiras linhas)
2. Logs completos da função `getProjetos()`
3. Logs completos da função `getDashboardStats()`
4. Logs do console do navegador ao acessar `/projetos`
5. Quantas linhas existem em cada aba (Orçamentos, Pedidos, Projetos)
