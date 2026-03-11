## Changelog

### Implementação de RF001-RF007

**Data:** 2026-03-11

---

### [RF001] Criar novo projeto a partir de projeto carregado

**Arquivos:** `formulario.html`

- Adicionadas variáveis globais `_originalProjetoData`, `_originalProjetoIndice`, `_originalProjetoIniciais`
- Em `preencherFormulario()`: ao carregar um projeto, os valores originais são armazenados e o banner de aviso é ocultado
- Adicionados listeners nos campos `projetoData`, `projetoIndice`, `projetoIniciais` que chamam `verificarMudancaCampoProjeto()` a cada input
- `verificarMudancaCampoProjeto()`: exibe banner `#bannerNovoProjeto` quando algum campo mudou em relação ao original
- Em `calcular()` e `salvarComoPedido()`: failsafe com `confirm()` antes de gerar PDF — informa se está criando novo projeto ou sobrescrevendo existente
- Banner HTML adicionado na seção "Informações do Projeto"

---

### [RF002] Refatorar memória de cálculo

**Arquivos:** `formulario.html`, `Código.js`

**Removido:**
- Seção HTML `#memoriaCalculoSection` completa
- Funções JS: `addMemoriaCalculo()`, `calcularMemoriaCalculo()`, `addProcessoAdicionalMemoria()`, `addOutroCustoMemoria()`, `salvarMemoriaCalculo()` (frontend)
- Coleta de `formData.memoriasCalculo` em `coletarDadosFormulario()`
- Preenchimento de memórias em `preencherFormulario()`
- Funções backend: `gerarPdfMemoriaCalculo()` e `salvarMemoriaCalculo()` (backend)

**Adicionado:**
- Em `addProdutoCadastrado()`: campo `div.produtoDescricoesProcessos` dentro da caixa de processos
- Função `atualizarDescricoesProcessos()` — cria/remove campos de texto para cada processo marcado/desmarcado
- Listeners nos checkboxes de processo chamam `atualizarDescricoesProcessos()`
- `coletarProdutosCadastrados()` agora inclui campo `descricoesProcessos` (objeto `{sigla: texto}`)
- `preencherFormulario()` preenche os campos de descrição ao carregar um projeto
- `gerarPdfOrdemProducao()` (backend): exibe descrições por processo abaixo da descrição do item

---

### [RF003] Número sequencial no nome do arquivo da proposta

**Arquivo:** `Código.js`

- `gerarPdfOrcamento()`: nome do arquivo mudado para `Proposta_codigoBase_numeroSequencial[_vN].pdf`
- Exemplo: `Proposta_260310aMS_1705.pdf`, v2: `Proposta_260310aMS_1705_v2.pdf`
- `detectarProximaVersao()`: atualizado para reconhecer formatos antigo e novo

---

### [RF004] Botão de Data de Entrega em projetos.html

**Arquivos:** `projetos.html`, `Código.js`

- Botão 📦 adicionado na coluna de ações, visível apenas para projetos "Convertido em Pedido"
- Amarelo se pendente, verde se preenchida; `informarDataEntrega(linha)` via prompt
- Filtro "Entrega preenchida / Entrega a preencher" adicionado ao lado do filtro de NF
- Nova função backend `informarDataEntregaProjeto(linha, dataEntrega)`
- `getProjetos()`: extrai `dataEntrega` do `JSON_DADOS`

---

### [RF005] "Gerar nova proposta"

**Arquivo:** `formulario.html`

- Label renomeado de "Gerar como v2, v3..." para "Gerar nova proposta"
- Sem pasta nova; salva com sufixo _v2, _v3 na mesma pasta atual
- Nome do arquivo inclui número sequencial (RF003)

---

### [RF006] Botões "Criar pasta 01_IN" e "Abrir pasta 01_IN"

**Arquivos:** `formulario.html`, `Código.js`

- Botões reativados: "📁 Criar pasta 01_IN" e "🔗 Abrir pasta 01_IN"
- `criarPasta01INAction()` / `abrirPasta01INAction()` no frontend
- Novas funções backend: `criarPasta01IN()` e `abrirPasta01IN()`

---

### [RF007] Campo "Nome Abreviado" do cliente

**Arquivos:** `formulario.html`, `projetos.html`, `Código.js`

- Campo `#clienteNomeAbreviado` no formulário; campo no modal "Novo Projeto Rápido"
- `getTodosClientes()`: inclui `nomeAbreviado: dados[i][5]`
- `salvarClienteSeNovo()`: salva na 6ª coluna; atualiza se estava vazio
- `gerarNomePasta()`: usa nome abreviado quando fornecido (fallback para nome completo)
- Todas as funções de criação de pasta recebem e passam `nomeAbreviado`
