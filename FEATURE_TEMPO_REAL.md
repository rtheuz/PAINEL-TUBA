# Sistema de Rastreamento de Tempo Real - Kanban

## Visão Geral

Este sistema permite que operadores registrem o tempo **real** de execução de cada processo no Kanban, separado do tempo total que o card fica na coluna (que inclui tempo de espera na fila).

## Como Funciona

### Para Operadores

1. **Iniciar Processo**: Quando começar a trabalhar em um pedido, clique no botão **"▶ Iniciar"** no card
2. **Durante Execução**: 
   - O card ficará destacado com uma borda animada
   - Um timer mostrará o tempo decorrido em tempo real
   - O botão mudará para **"⏹ Finalizar"**
3. **Finalizar Processo**: Quando terminar o trabalho, clique em **"⏹ Finalizar"**
   - O sistema calculará automaticamente a duração
   - O botão mostrará **"✓ XXmin"** com o tempo total

### Características

- ✅ **Botões grandes e fáceis de clicar** - funcionam bem em touch screens
- ✅ **Destaque visual** - card em execução tem borda pulsante e animada
- ✅ **Timer em tempo real** - mostra tempo decorrido durante execução
- ✅ **Persistência** - dados salvos automaticamente no Google Sheets
- ✅ **Compatível com drag & drop** - não interfere com movimentação dos cards

## Estrutura de Dados

### Frontend (kanban.html)

Os dados de tempo real são armazenados em `card.temposReais`:

```javascript
{
  "processo-de-corte": {
    "iniciadoEm": "2026-01-12T09:30:00.000Z",
    "finalizadoEm": "2026-01-12T10:15:00.000Z",
    "duracaoMinutos": 45
  },
  "processo-de-dobra": {
    "iniciadoEm": null,
    "finalizadoEm": null,
    "duracaoMinutos": null
  }
}
```

### Backend (Código.gs)

Função `salvarTempoReal(cliente, projeto, processoSlug, tipo, timestamp, duracaoMinutos)`:
- Salva os dados na coluna `JSON_DADOS` da planilha Projetos
- Tipos: `'INICIO'` ou `'FIM'`
- Mantém estrutura separada para cada processo

## Exemplo de Uso

1. Card "ACME Corp - Projeto 123" está em "Processo de Corte"
2. Operador clica **"▶ Iniciar"** às 09:30
3. Sistema salva: `iniciadoEm: "2026-01-12T09:30:00Z"`
4. Timer mostra tempo decorrido em tempo real
5. Operador clica **"⏹ Finalizar"** às 10:15
6. Sistema calcula: `duracaoMinutos: 45` e salva `finalizadoEm`
7. Botão mostra **"✓ 45min"**

## Benefícios

- 📊 **Dados reais de produtividade** - saber quanto tempo cada processo realmente leva
- ⏱️ **Separação de tempo de fila** - não conta tempo de espera
- 📈 **Melhor planejamento** - dados históricos para estimativas futuras
- 👀 **Visibilidade** - saber qual pedido está sendo trabalhado no momento

## Notas Técnicas

### Arquivos Modificados

1. **kanban.html**: 
   - Adicionados estilos CSS para botões e animações
   - Função `createCardElement()` cria botões baseado em `temposReais`
   - Handlers `startProcess()` e `finishProcess()`
   - Timer de atualização em tempo real

2. **Código.gs** e **Código.js**:
   - Nova função `salvarTempoReal()`
   - Atualização em `getKanbanData()` para carregar tempos reais

### Compatibilidade

- ✅ Funciona com drag & drop existente
- ✅ Compatível com dispositivos móveis
- ✅ Dados sobrevivem refresh da página
- ✅ Não afeta tempo estimado ou logs existentes
