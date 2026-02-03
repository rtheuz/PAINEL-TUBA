/************* Code.gs *************/
const ss = SpreadsheetApp.openById("1wMIbd8r2HeniFLTYaG8Yhhl8CWmaHaW5oXBVnxj0xos");
const SHEET_CALC = ss.getSheetByName("Tabelas para cálculos");
const SHEET_VEIC = ss.getSheetByName('Controle de Veículos');
const SHEET_MANU_NAME = ss.getSheetByName("Registro de Manutenções");
const SHEET_PED = ss.getSheetByName("Pedidos");
const SHEET_MAT = ss.getSheetByName("Controle de Materiais");
const SHEET_AVAL = ss.getSheetByName("Avaliações");
const SHEET_LOGS = ss.getSheetByName("Logs");
const SHEET_CLIENTES = ss.getSheetByName("Cadastro de Clientes");
const SHEET_PRODUTOS = ss.getSheetByName("Relação de produtos");
const SHEET_PROJ = ss.getSheetByName("Projetos"); // Nova aba unificada
const ID_PASTA_PRINCIPAL = "1jqIVHbThV3SPBM8MOHek4r5tr2DoHbqz";
const ID_LOGO = "1pnRLV6YZYMD6Yhv1cUb4FXVr0ol_Zzzf";
const FAVICON = "https://i.imgur.com/C0dSTyE.png"

// ==================== PRODUTOS CADASTRADOS ====================
/**
 * Busca produtos cadastrados da aba "Relação de produtos"
 * @returns {Array} Array de objetos com dados dos produtos
 */
function getProdutosCadastrados() {
  try {
    const SHEET_PRODUTOS = ss.getSheetByName("Relação de produtos");
    if (!SHEET_PRODUTOS) {
      Logger.log("Aba 'Relação de produtos' não encontrada");
      return [];
    }

    const dados = SHEET_PRODUTOS.getDataRange().getValues();
    if (dados.length < 2) return [];

    // Estrutura da planilha:
    // A=Código do Produto, B=Descrição do Produto, C=Código da Família, D=Código EAN (GTIN), 
    // E=Código NCM, F=Preço Unitário de Venda, G=Unidade, H=Características, I=Estoque, J=Local de estoque
    const produtos = [];
    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      if (row[0]) { // se tem código (coluna A)
        produtos.push({
          codigo: row[0],                    // Coluna A - Código do Produto
          descricao: row[1] || "",           // Coluna B - Descrição do Produto
          codigoFamilia: row[2] || "",       // Coluna C - Código da Família
          codigoEAN: row[3] || "",           // Coluna D - Código EAN (GTIN)
          ncm: row[4] || "",                 // Coluna E - Código NCM
          preco: parseFloat(row[5]) || 0,    // Coluna F - Preço Unitário de Venda
          unidade: row[6] || "UN"            // Coluna G - Unidade
        });
      }
    }
    return produtos;
  } catch (err) {
    Logger.log("Erro ao buscar produtos cadastrados: " + err);
    return [];
  }
}

/**
 * Obtém o próximo código PRD disponível
 * @returns {string} Próximo código no formato PRD00001, PRD00002, etc.
 */
function getProximoCodigoPRD() {
  try {
    const SHEET_PRODUTOS = ss.getSheetByName("Relação de produtos");
    if (!SHEET_PRODUTOS) {
      return "PRD00001"; // Primeiro código se a aba não existe
    }

    const dados = SHEET_PRODUTOS.getDataRange().getValues();
    if (dados.length < 2) {
      return "PRD00001"; // Primeiro código se não há produtos
    }

    // Encontra o maior número PRD
    let maxNumero = 0;
    for (let i = 1; i < dados.length; i++) {
      const codigo = String(dados[i][0] || "");
      if (codigo.startsWith("PRD")) {
        const numero = parseInt(codigo.substring(3), 10);
        if (!isNaN(numero) && numero > maxNumero) {
          maxNumero = numero;
        }
      }
    }

    // Retorna o próximo número formatado
    const proximoNumero = maxNumero + 1;
    return "PRD" + String(proximoNumero).padStart(5, "0");
  } catch (err) {
    Logger.log("Erro ao obter próximo código PRD: " + err);
    return "PRD00001";
  }
}

/**
 * Atribui códigos PRD aos produtos que não possuem código
 * @param {Array} produtos - Array de objetos de produtos
 * @returns {Array} Array de produtos com códigos atribuídos
 */
function atribuirCodigosPRDAutomaticos(produtos) {
  if (!produtos || produtos.length === 0) return produtos;
  
  // Conta quantos produtos precisam de código
  const produtosSemCodigo = produtos.filter(p => !p.codigo || p.codigo.trim() === "");
  
  if (produtosSemCodigo.length === 0) {
    return produtos; // Todos já têm código
  }
  
  // Obtém o próximo código PRD disponível
  const proximoCodigo = getProximoCodigoPRD();
  let numeroBase = parseInt(proximoCodigo.substring(3), 10);
  
  // Atribui códigos aos produtos que não têm
  produtos.forEach(produto => {
    if (!produto.codigo || produto.codigo.trim() === "") {
      produto.codigo = "PRD" + String(numeroBase).padStart(5, "0");
      numeroBase++;
    }
  });
  
  return produtos;
}

/**
 * Insere um produto na aba "Relação de produtos"
 * @param {Object} produto - Objeto com os dados do produto
 */
function inserirProdutoNaRelacao(produto) {
  try {
    Logger.log("Tentando inserir produto: " + JSON.stringify(produto));

    const SHEET_PRODUTOS = ss.getSheetByName("Relação de produtos");
    if (!SHEET_PRODUTOS) {
      Logger.log("ERRO: Aba 'Relação de produtos' não encontrada");
      return false;
    }

    Logger.log("Aba 'Relação de produtos' encontrada. Verificando duplicatas...");

    // Verifica se o produto já existe
    const dados = SHEET_PRODUTOS.getDataRange().getValues();
    Logger.log("Total de linhas na planilha: " + dados.length);

    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0] === produto.codigo) {
        Logger.log("Produto " + produto.codigo + " já existe na relação (linha " + (i + 1) + ")");
        return false; // Produto já existe
      }
    }

    // Estrutura da planilha:
    // A=Código do Produto, B=Descrição, C=Código da Família, D=Código EAN (GTIN), 
    // E=Código NCM, F=Preço Unitário de Venda, G=Unidade, H=Características, I=Estoque, J=Local de estoque
    const novaLinha = [
      produto.codigo || "",           // A - Código do Produto
      produto.descricao || "",        // B - Descrição do Produto
      "",                             // C - Código da Família (vazio)
      "",                             // D - Código EAN (vazio)
      produto.ncm || "",              // E - Código NCM
      produto.preco || 0,             // F - Preço Unitário de Venda
      produto.unidade || "UN",        // G - Unidade
      produto.caracteristicas || ""   // H - Características
    ];

    Logger.log("Inserindo nova linha: " + JSON.stringify(novaLinha));
    SHEET_PRODUTOS.appendRow(novaLinha);
    Logger.log("✓ Produto " + produto.codigo + " inserido com sucesso na relação");
    return true;
  } catch (err) {
    Logger.log("ERRO ao inserir produto na relação: " + err);
    Logger.log("Stack trace: " + err.stack);
    return false;
  }
}

/**
 * Atualiza um PRD no catálogo e salva um log com os dados antigos
 * @param {Object} dadosNovos - Objeto com os novos dados do produto
 * @returns {Object} - Resultado da operação
 */
function atualizarPRDNoCatalogo(dadosNovos) {
  try {
    if (!dadosNovos || !dadosNovos.codigo) {
      throw new Error("Código do produto é obrigatório");
    }

    const SHEET_PRODUTOS = ss.getSheetByName("Relação de produtos");
    if (!SHEET_PRODUTOS) {
      throw new Error("Aba 'Relação de produtos' não encontrada");
    }

    const dados = SHEET_PRODUTOS.getDataRange().getValues();
    let linhaEncontrada = -1;
    let dadosAntigos = null;

    // Busca o produto pelo código
    for (let i = 1; i < dados.length; i++) {
      if (String(dados[i][0]).trim() === dadosNovos.codigo.trim()) {
        linhaEncontrada = i + 1; // +1 porque índice começa em 1 na planilha
        // Salva dados antigos para o log
        dadosAntigos = {
          codigo: dados[i][0] || "",
          descricao: dados[i][1] || "",
          ncm: dados[i][4] || "",
          preco: dados[i][5] || 0,
          unidade: dados[i][6] || "UN"
        };
        break;
      }
    }

    if (linhaEncontrada === -1) {
      throw new Error(`Produto com código ${dadosNovos.codigo} não encontrado no catálogo`);
    }

    // Atualiza os dados na planilha
    // Estrutura: A=Código, B=Descrição, C=Código Família, D=EAN, E=NCM, F=Preço, G=Unidade, H=Características
    SHEET_PRODUTOS.getRange(linhaEncontrada, 2).setValue(dadosNovos.descricao || ""); // B - Descrição
    SHEET_PRODUTOS.getRange(linhaEncontrada, 5).setValue(dadosNovos.ncm || ""); // E - NCM
    SHEET_PRODUTOS.getRange(linhaEncontrada, 6).setValue(dadosNovos.preco || 0); // F - Preço
    SHEET_PRODUTOS.getRange(linhaEncontrada, 7).setValue(dadosNovos.unidade || "UN"); // G - Unidade

    return {
      success: true,
      mensagem: `PRD atualizado no catálogo.`
    };
  } catch (err) {
    Logger.log("Erro ao atualizar PRD no catálogo: " + err.message);
    throw new Error("Erro ao atualizar PRD: " + err.message);
  }
}

/**
 * Insere produtos com código PRD das chapas na "Relação de produtos"
 * @param {Array} chapas - Array com dados das chapas e peças
 */
function inserirProdutosDasChapas(chapas) {
  try {
    if (!Array.isArray(chapas)) {
      Logger.log("inserirProdutosDasChapas: chapas não é um array");
      return;
    }

    Logger.log("inserirProdutosDasChapas: Processando " + chapas.length + " chapas");

    let produtosInseridos = 0;
    let produtosPulados = 0;

    chapas.forEach((chapa, chapaIdx) => {
      if (chapa.pecas && Array.isArray(chapa.pecas)) {
        Logger.log("Chapa " + chapaIdx + ": " + chapa.pecas.length + " peças encontradas");
        chapa.pecas.forEach((peca, pecaIdx) => {
          // Só insere se tiver código PRD
          if (peca.codigo && String(peca.codigo).startsWith("PRD")) {
            Logger.log("Peça " + pecaIdx + " tem código PRD: " + peca.codigo);
            const produto = {
              codigo: peca.codigo,
              descricao: peca.descricao || "",
              ncm: "",  // Peças não têm NCM específico
              preco: peca.precoUnitario || 0,
              unidade: "UN",
              caracteristicas: `${chapa.material} - ${peca.comprimento}x${peca.largura} - ${chapa.espessura}mm`
            };
            const resultado = inserirProdutoNaRelacao(produto);
            if (resultado) {
              produtosInseridos++;
            } else {
              produtosPulados++;
            }
          } else {
            Logger.log("Peça " + pecaIdx + " não tem código PRD válido: " + (peca.codigo || "sem código"));
            produtosPulados++;
          }
        });
      } else {
        Logger.log("Chapa " + chapaIdx + ": sem peças ou peças não é array");
      }
    });

    Logger.log("Total: " + produtosInseridos + " produtos inseridos, " + produtosPulados + " pulados");
  } catch (err) {
    Logger.log("Erro ao inserir produtos das chapas: " + err);
  }
}

// ==================== HELPERS DE OTIMIZAÇÃO ====================
/**
 * Retorna índice (0-based) do material na ordem do objeto MATERIAIS.
 * Usado para calcular offsets (por exemplo linhas de corte/dobra baseadas em um índice)
 */
function _getMaterialIndexMap() {
  const keys = Object.keys(MATERIAIS);
  const map = {};
  keys.forEach((k, i) => map[k] = i);
  return { keys, map };
}

/**
 * Lê preços (colunas L e M) para todas as entradas de MATERIAIS de uma só vez.
 * Retorna objeto { "NOME_MAT": { precoUnit: x, precoTotalPlanilha: y } }
 */
function _lerPrecosMateriais() {
  const matKeys = Object.keys(MATERIAIS);
  // assumindo que linhaPreco em MATERIAIS é sequencial por material (4,5,6..)
  const linhas = matKeys.map(k => MATERIAIS[k].linhaPreco);
  const min = Math.min.apply(null, linhas);
  const max = Math.max.apply(null, linhas);
  const count = max - min + 1;
  const valores = SHEET_CALC.getRange(min, 12, count, 2).getValues(); // col L=12, M=13 -> here cols 12,13
  const res = {};
  matKeys.forEach((k, i) => {
    const rowIndex = linhas[i] - min; // offset
    const v = valores[rowIndex] || [0, 0];
    res[k] = { precoUnit: parseFloat(v[0]) || 0, precoTotalPlanilha: parseFloat(v[1]) || 0 };
  });
  return res;
}

function _preencherInputsCalcParaPeca(mat, chapa, peca) {
  const linhaChapa = mat.linhaChapa;
  const linhaPeca = mat.linhaPeca;

  // 1) Preenche C/D/E da linhaChapa (comprimento, largura, espessura) — contíguo
  SHEET_CALC.getRange(linhaChapa, 3, 1, 3)
    .setValues([[chapa.comprimento, chapa.largura, chapa.espessura]]); // C,D,E

  // 2) Preenche B/C (col 2,3) da linhaPeca
  SHEET_CALC.getRange(linhaPeca, 2, 1, 2)
    .setValues([[peca.comprimento, peca.largura]]); // B,C

  // 3) Preenche E/F (col 5,6) da linhaPeca (numPecasLote, numPecasChapa)
  SHEET_CALC.getRange(linhaPeca, 5, 1, 2)
    .setValues([[Number(peca.numPecasLote || 0), Number(peca.numPecasChapa || 0)]]); // E,F

  // 4) Tempo de corte (linhaCorte col E = 5)
  SHEET_CALC.getRange(mat.linhaCorte, 5).setValue(Number(peca.tempoCorte || 0));

  // 5) Dobra: D (col4) = numDobras, E (col5) = tempoDobra (col5 may be used differently per sheet row)
  SHEET_CALC.getRange(mat.linhaDobra, 4, 1, 2)
    .setValues([[Number(peca.numDobras || 0), Number(peca.tempoDobra || 0)]]); // D,E

  // 6) setupDobra (col G = 7)
  SHEET_CALC.getRange(mat.linhaDobra, 7).setValue(Number(peca.setupDobra || 0));
}

// Formata número BR (R$) (mesma lógica que você tinha)
function _formatBR(n) {
  const num = Number(n) || 0;
  const parts = num.toFixed(2).split('.');
  parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
  return "R$ " + parts[0] + ',' + parts[1];
}

// Formata data BR (dd/mm/yyyy)
function formatarDataBrasil(date) {
  if (!date) return "";
  const d = new Date(date);
  const dia = String(d.getUTCDate()).padStart(2, "0");
  const mes = String(d.getUTCMonth() + 1).padStart(2, "0");
  const ano = d.getUTCFullYear();
  return `${dia}/${mes}/${ano}`;
}

const MATERIAIS = {
  "AÇO CARBONO": { linhaChapa: 4, linhaPeca: 12, linhaCorte: 20, linhaDobra: 28, linhaPreco: 4 },
  "ALUMÍNIO": { linhaChapa: 5, linhaPeca: 13, linhaCorte: 21, linhaDobra: 29, linhaPreco: 5 },
  "INOX 200 OU 300": { linhaChapa: 6, linhaPeca: 14, linhaCorte: 22, linhaDobra: 30, linhaPreco: 6 },
  "INOX 400": { linhaChapa: 7, linhaPeca: 15, linhaCorte: 23, linhaDobra: 31, linhaPreco: 7 },
  "LATÃO": { linhaChapa: 8, linhaPeca: 16, linhaCorte: 24, linhaDobra: 32, linhaPreco: 8 },
  "COBRE": { linhaChapa: 9, linhaPeca: 17, linhaCorte: 25, linhaDobra: 33, linhaPreco: 9 }
};

// ========================= CÁLCULO DE ORÇAMENTO =========================
function calcularOrcamento(chapas) {
  const resultados = [];
  if (!chapas || !chapas.length) return resultados;

  chapas.forEach(chapa => {
    const mat = MATERIAIS[chapa.material];
    if (!mat) return;

    chapa.pecas.forEach(peca => {
      // escreve inputs corretos na planilha
      _preencherInputsCalcParaPeca(mat, chapa, peca);

      // força recálculo para que as fórmulas (coluna L etc.) sejam atualizadas
      SpreadsheetApp.flush();

      // lê o preço atualizado relativo a esse material (coluna L = 12)
      const precoUnitario = parseFloat(SHEET_CALC.getRange(mat.linhaPreco, 12).getValue()) || 0;
      // (opcional) se precisar do total na planilha leia M = col 13
      // const precoTotalPlanilha = parseFloat(SHEET_CALC.getRange(mat.linhaPreco, 13).getValue()) || 0;

      const adicionaisPorUnidade = parseFloat(peca.adicionaisTotal) || 0;
      const precoUnitarioFinal = precoUnitario + adicionaisPorUnidade;
      const quantidade = parseFloat(peca.numPecasLote) || 0;
      const precoTotalFinal = precoUnitarioFinal * quantidade;

      resultados.push({
        descricao: peca.descricao,
        codigo: peca.codigo,
        quantidade: quantidade,
        precoUnitario: precoUnitarioFinal,
        precoTotal: precoTotalFinal
      });
    });
  });

  // tratar conjuntos (mantendo sua lógica, mas lendo preço quando necessário)
  const agrupados = [];
  chapas.forEach(chapa => {
    if (!chapa.isConjunto) return;
    const mat = MATERIAIS[chapa.material];
    if (!mat) return;

    let somaTotal = 0;
    chapa.pecas.forEach(p => {
      const qtd = parseFloat(p.numPecasLote) || 0;

      const encontrado = resultados.find(r =>
        r.descricao === p.descricao &&
        r.codigo === p.codigo &&
        r.quantidade === qtd
      );

      if (encontrado) {
        somaTotal += encontrado.precoTotal;
        const index = resultados.indexOf(encontrado);
        if (index > -1) resultados.splice(index, 1);
      } else {
        // fallback: escreve inputs específicos para essa peça, força recálculo e lê preço
        try {
          _preencherInputsCalcParaPeca(mat, chapa, p);
          SpreadsheetApp.flush();
          const precoUnitario = parseFloat(SHEET_CALC.getRange(mat.linhaPreco, 12).getValue()) || 0;
          const adicionais = parseFloat(p.adicionaisTotal) || 0;
          somaTotal += (precoUnitario + adicionais) * qtd;
        } catch (e) {
          Logger.log("Erro fallback conjunto: " + e);
        }
      }
    });

    const descricaoConj = chapa.descricaoConjunto || "Conjunto";
    const codigoConj = chapa.codigoConjunto || "";
    const qtdConj = parseFloat(chapa.quantidadeConjunto) || 1;

    agrupados.push({
      descricao: descricaoConj,
      codigo: codigoConj,
      quantidade: qtdConj,
      precoUnitario: somaTotal,
      precoTotal: somaTotal * qtdConj
    });
  });

  return resultados.concat(agrupados);
}

// ========================= PREVIEW DE ORÇAMENTO =========================
/**
 * Calcula preview do orçamento em tempo real
 * @param {Object} dados - Dados do formulário (chapas, produtosCadastrados, processosPedido)
 * @returns {Object} { total, detalhamento, timestamp }
 */
function calcularPreviewOrcamento(dados) {
  try {
    let totalGeral = 0;
    const detalhamento = [];

    // 1. Produtos Cadastrados
    if (dados.produtosCadastrados && Array.isArray(dados.produtosCadastrados)) {
      dados.produtosCadastrados.forEach(prod => {
        const precoTotal = (parseFloat(prod.precoUnitario) || 0) * (parseFloat(prod.quantidade) || 0);
        totalGeral += precoTotal;
        detalhamento.push({
          tipo: 'produto',
          descricao: prod.descricao || prod.codigo,
          quantidade: prod.quantidade,
          precoUnitario: prod.precoUnitario,
          precoTotal: precoTotal
        });
      });
    }

    // 2. Chapas/Peças (usa função existente)
    if (dados.chapas && Array.isArray(dados.chapas)) {
      const resultadosChapas = calcularOrcamento(dados.chapas);
      resultadosChapas.forEach(res => {
        totalGeral += res.precoTotal || 0;
        detalhamento.push({
          tipo: 'peca',
          descricao: res.descricao,
          codigo: res.codigo,
          quantidade: res.quantidade,
          precoUnitario: res.precoUnitario,
          precoTotal: res.precoTotal
        });
      });
    }

    // 3. Processos do Pedido
    if (dados.processosPedido && Array.isArray(dados.processosPedido)) {
      dados.processosPedido.forEach(proc => {
        const valorHora = parseFloat(proc.valorHora) || 0;
        const horas = parseFloat(proc.horas) || 0;
        const valorMat = parseFloat(proc.valorMat) || 0;
        const qtdMat = parseFloat(proc.qtdMat) || 0;
        const valorFixo = parseFloat(proc.valorFixo) || 0;
        const preco = valorHora * horas + valorMat * qtdMat + valorFixo;
        totalGeral += preco;
        detalhamento.push({
          tipo: 'processo',
          descricao: proc.descricao || 'Processo adicional',
          precoTotal: preco
        });
      });
    }

    return {
      total: totalGeral,
      detalhamento: detalhamento,
      timestamp: new Date().toISOString()
    };
  } catch (err) {
    Logger.log("Erro calcularPreviewOrcamento: " + err.message);
    return { total: 0, detalhamento: [], erro: err.message };
  }
}

// ========================= CLIENTES =========================
function getTodosClientes() {
  const dados = SHEET_CLIENTES.getDataRange().getValues();
  const clientes = [];
  for (let i = 1; i < dados.length; i++) {
    clientes.push({
      nome: dados[i][0],
      cpf: dados[i][1],
      endereco: dados[i][2],
      telefone: dados[i][3],
      email: dados[i][4]
    });
  }
  return clientes;
}

function salvarClienteSeNovo(cliente) {
  const dados = SHEET_CLIENTES.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0].toString().trim().toLowerCase() === cliente.nome.trim().toLowerCase()) return;
  }
  SHEET_CLIENTES.appendRow([cliente.nome, cliente.cpf, cliente.endereco, cliente.telefone, cliente.email]);
}

// ========================= PASTAS =========================

/**
 * Remove caracteres inválidos do nome da pasta
 * @param {string} texto - Texto a ser limpo
 * @returns {string} - Texto limpo
 */
function limparNomePasta(texto) {
  if (!texto) return "";
  return String(texto)
    .replace(/[\/\\:*?"<>|]/g, "") // Remove caracteres inválidos do Drive
    .replace(/\s+/g, " ")           // Normaliza espaços múltiplos
    .trim();
}

/**
 * Gera o nome formatado da pasta
 * @param {string} codigoProjeto - Código do projeto (ex: "260202aBR")
 * @param {string} nomeCliente - Nome do cliente
 * @param {string} descricao - Descrição do projeto
 * @param {boolean} isPedido - Se true, usa prefixo PED; se false, usa COT
 * @returns {string} - Nome formatado (ex: "260202aBR COT CLIENTE - DESCRICAO")
 */
function gerarNomePasta(codigoProjeto, nomeCliente, descricao, isPedido) {
  const prefixo = isPedido ? "PED" : "COT";
  const clienteLimpo = limparNomePasta(nomeCliente || "");
  const descricaoLimpa = limparNomePasta(descricao || "");
  
  let nomeFinal = codigoProjeto + " " + prefixo;
  if (clienteLimpo) {
    nomeFinal += " " + clienteLimpo;
  }
  if (descricaoLimpa) {
    nomeFinal += " - " + descricaoLimpa;
  }
  
  return nomeFinal;
}

/**
 * Detecta se uma pasta para o projeto já existe e retorna seu tipo (COT/PED)
 * @param {string} codigoProjeto - Código do projeto
 * @param {string} data - Data no formato YYMMDD
 * @returns {Object|null} - {pasta: Folder, tipo: "COT"|"PED", estrutura: "PROJ"|"COM"} ou null
 */
function detectarPastaExistente(codigoProjeto, data) {
  try {
    const root = DriveApp.getFolderById(ID_PASTA_PRINCIPAL);
    const ano = data.substring(0, 2);
    const mes = data.substring(0, 4);
    const dia = data;
    
    Logger.log("🔍 Buscando pasta para projeto: " + codigoProjeto + " na data: " + dia);
    
    // Tenta estrutura nova: PROJ
    try {
      const anoFolder = root.getFoldersByName("20" + ano);
      if (anoFolder.hasNext()) {
        const ano_f = anoFolder.next();
        const mesFolder = ano_f.getFoldersByName(mes);
        if (mesFolder.hasNext()) {
          const mes_f = mesFolder.next();
          const diaFolder = mes_f.getFoldersByName(dia);
          if (diaFolder.hasNext()) {
            const dia_f = diaFolder.next();
            const projFolder = dia_f.getFoldersByName("PROJ");
            if (projFolder.hasNext()) {
              const proj_f = projFolder.next();
              
              // Busca pasta que COMEÇA com o código do projeto
              const pastas = proj_f.getFolders();
              while (pastas.hasNext()) {
                const pasta = pastas.next();
                const nomePasta = pasta.getName();
                
                // Verifica se o nome começa com o código do projeto seguido de espaço
                if (nomePasta.startsWith(codigoProjeto + " ")) {
                  const tipo = nomePasta.includes(" PED ") ? "PED" : "COT";
                  Logger.log("✅ Pasta encontrada: " + nomePasta + " (tipo: " + tipo + ")");
                  return { pasta: pasta, tipo: tipo, estrutura: "PROJ" };
                }
              }
            }
          }
        }
      }
    } catch (e) {
      Logger.log("Estrutura PROJ não encontrada: " + e.message);
    }
    
    // Tenta estrutura antiga: COM
    try {
      const anoFolder = root.getFoldersByName("20" + ano);
      if (anoFolder.hasNext()) {
        const ano_f = anoFolder.next();
        const mesFolder = ano_f.getFoldersByName(mes);
        if (mesFolder.hasNext()) {
          const mes_f = mesFolder.next();
          const diaFolder = mes_f.getFoldersByName(dia);
          if (diaFolder.hasNext()) {
            const dia_f = diaFolder.next();
            const comFolder = dia_f.getFoldersByName("COM");
            if (comFolder.hasNext()) {
              const com_f = comFolder.next();
              
              // Busca pasta que COMEÇA com o código do projeto
              const pastas = com_f.getFolders();
              while (pastas.hasNext()) {
                const pasta = pastas.next();
                const nomePasta = pasta.getName();
                
                if (nomePasta.startsWith(codigoProjeto + " ")) {
                  const tipo = nomePasta.includes(" PED ") ? "PED" : "COT";
                  Logger.log("✅ Pasta encontrada (estrutura antiga): " + nomePasta + " (tipo: " + tipo + ")");
                  return { pasta: pasta, tipo: tipo, estrutura: "COM" };
                }
              }
            }
          }
        }
      }
    } catch (e) {
      Logger.log("Estrutura COM não encontrada: " + e.message);
    }
    
    Logger.log("❌ Pasta não encontrada para: " + codigoProjeto);
    return null;
  } catch (e) {
    Logger.log("Erro ao detectar pasta: " + e.message);
    return null;
  }
}

/**
 * Atualiza o nome da pasta mantendo o prefixo atual (COT ou PED)
 * @param {Folder} pasta - Pasta a ser renomeada
 * @param {string} codigoProjeto - Código do projeto
 * @param {string} nomeCliente - Nome do cliente
 * @param {string} descricao - Descrição do projeto
 */
function atualizarNomePasta(pasta, codigoProjeto, nomeCliente, descricao) {
  try {
    const nomeAtual = pasta.getName();
    const isPedido = nomeAtual.includes(" PED ");
    const novoNome = gerarNomePasta(codigoProjeto, nomeCliente, descricao, isPedido);
    
    if (nomeAtual !== novoNome) {
      pasta.setName(novoNome);
      Logger.log("Pasta renomeada de '" + nomeAtual + "' para '" + novoNome + "'");
    }
  } catch (e) {
    Logger.log("Erro ao atualizar nome da pasta: " + e.message);
  }
}

/**
 * Atualiza o prefixo da pasta de COT para PED quando um orçamento é convertido em pedido
 * @param {string} codigoProjeto - Código do projeto
 * @param {string} data - Data no formato YYMMDD
 * @param {string} nomeCliente - Nome do cliente
 * @param {string} descricao - Descrição do projeto
 * @returns {boolean} - True se renomeou com sucesso, false caso contrário
 */
function atualizarPrefixoPastaParaPedido(codigoProjeto, data, nomeCliente, descricao) {
  try {
    Logger.log("🔄 Iniciando conversão de COT para PED: " + codigoProjeto);
    
    const pastaInfo = detectarPastaExistente(codigoProjeto, data);
    if (!pastaInfo) {
      Logger.log("❌ Pasta não encontrada para converter para PED: " + codigoProjeto);
      return false;
    }
    
    if (pastaInfo.tipo === "PED") {
      Logger.log("✅ Pasta já é PED: " + codigoProjeto);
      return true; // Já está como PED
    }
    
    // Renomeia para PED
    const novoNome = gerarNomePasta(codigoProjeto, nomeCliente, descricao, true);
    pastaInfo.pasta.setName(novoNome);
    Logger.log("✅ Pasta convertida de COT para PED: " + novoNome);
    return true;
  } catch (e) {
    Logger.log("❌ Erro ao converter pasta para PED: " + e.message);
    return false;
  }
}
/**
 * Cria ou usa pasta existente do projeto na estrutura PROJ
 * @param {string} codigoProjeto - Código do projeto
 * @param {string} nomeCliente - Nome do cliente
 * @param {string} descricao - Descrição do projeto
 * @param {string} data - Data no formato YYMMDD
 * @param {boolean} isPedido - Se true, usa prefixo PED; se false, usa COT
 * @returns {Folder} - Pasta do projeto
 */
function criarOuUsarPastaProjeto(codigoProjeto, nomeCliente, descricao, data, isPedido) {
  Logger.log("📁 criarOuUsarPastaProjeto - Código: " + codigoProjeto + ", isPedido: " + isPedido);
  
  // Valida descrição obrigatória
  if (!descricao || descricao.trim() === "") {
    throw new Error("Descrição do projeto é obrigatória para criar a pasta.");
  }
  
  // Detecta se pasta já existe
  const pastaInfo = detectarPastaExistente(codigoProjeto, data);
  
  if (pastaInfo) {
    Logger.log("✅ Usando pasta existente: " + pastaInfo.pasta.getName());
    // Pasta existe - atualiza o nome se necessário
    const nomeDesejado = gerarNomePasta(codigoProjeto, nomeCliente, descricao, isPedido);
    
    // Se mudou de COT para PED, atualiza
    if (isPedido && pastaInfo.tipo === "COT") {
      pastaInfo.pasta.setName(nomeDesejado);
      Logger.log("🔄 Pasta convertida de COT para PED: " + nomeDesejado);
    } 
    // Se o nome mudou (cliente ou descrição), atualiza mantendo o tipo atual
    else if (!isPedido || pastaInfo.tipo === "PED") {
      const tipoAtual = isPedido ? "PED" : pastaInfo.tipo;
      const nomeAtualizado = gerarNomePasta(codigoProjeto, nomeCliente, descricao, tipoAtual === "PED");
      if (pastaInfo.pasta.getName() !== nomeAtualizado) {
        pastaInfo.pasta.setName(nomeAtualizado);
        Logger.log("📝 Pasta atualizada: " + nomeAtualizado);
      }
    }
    
    // Se a pasta está na estrutura antiga (COM), move para PROJ
    if (pastaInfo.estrutura === "COM") {
      try {
        const root = DriveApp.getFolderById(ID_PASTA_PRINCIPAL);
        const anoFolder = getOrCreateSubFolder(root, "20" + data.substring(0, 2));
        const mesFolder = getOrCreateSubFolder(anoFolder, data.substring(0, 4));
        const diaFolder = getOrCreateSubFolder(mesFolder, data);
        const projFolder = getOrCreateSubFolder(diaFolder, "PROJ");
        
        // Move a pasta de COM para PROJ
        const pastaCom = pastaInfo.pasta.getParents().next();
        projFolder.addFolder(pastaInfo.pasta);
        pastaCom.removeFolder(pastaInfo.pasta);
        Logger.log("📦 Pasta migrada de COM para PROJ: " + pastaInfo.pasta.getName());
      } catch (e) {
        Logger.log("⚠️ Erro ao migrar pasta de COM para PROJ: " + e.message);
        // Continua usando a pasta na localização antiga
      }
    }
    
    return pastaInfo.pasta;
  }
  
  // Pasta não existe - cria nova na estrutura PROJ
  Logger.log("📁 Criando nova pasta para projeto: " + codigoProjeto);
  const root = DriveApp.getFolderById(ID_PASTA_PRINCIPAL);
  const anoFolder = getOrCreateSubFolder(root, "20" + data.substring(0, 2));
  const mesFolder = getOrCreateSubFolder(anoFolder, data.substring(0, 4));
  const diaFolder = getOrCreateSubFolder(mesFolder, data);
  const projFolder = getOrCreateSubFolder(diaFolder, "PROJ");
  
  const nomePasta = gerarNomePasta(codigoProjeto, nomeCliente, descricao, isPedido);
  const novaPasta = projFolder.createFolder(nomePasta);
  Logger.log("✅ Nova pasta criada: " + nomePasta);
  
  return novaPasta;
}
// Função legada para compatibilidade - redireciona para nova estrutura
function criarOuUsarPasta(codigoProjeto, nomePasta, data) {
  // Tenta detectar pasta existente primeiro (suporta COM e PROJ)
  const pastaInfo = detectarPastaExistente(codigoProjeto, data);
  if (pastaInfo) {
    return pastaInfo.pasta;
  }
  
  // Se não existe e nomePasta está vazio, usa código como descrição temporária
  const descricao = nomePasta || codigoProjeto;
  
  // Cria nova pasta usando estrutura PROJ
  // isPedido = false (COT) por padrão para compatibilidade
  return criarOuUsarPastaProjeto(codigoProjeto, "", descricao, data, false);
}

function buscarNomePastaPorCodigo(codigoProjeto) {
  const ano = codigoProjeto.slice(0, 2);
  const mes = codigoProjeto.slice(0, 4);
  const dia = codigoProjeto.slice(0, 6);
  
  try {
    // Tenta detectar pasta usando nova função
    const pastaInfo = detectarPastaExistente(codigoProjeto, dia);
    if (pastaInfo) {
      const nomePasta = pastaInfo.pasta.getName();
      // Remove o código e o prefixo (COT ou PED) para retornar apenas a parte customizada
      const prefixo = pastaInfo.tipo === "PED" ? " PED " : " COT ";
      const nomeCustomizado = nomePasta.replace(codigoProjeto + prefixo, "");
      return nomeCustomizado;
    }
    return "";
  } catch (e) {
    Logger.log("Erro ao buscar nome da pasta: " + e.message);
    return "";
  }
}

/**
 * Detecta a próxima versão disponível para um projeto baseado nos arquivos PDF existentes na pasta
 * @param {string} codigoProjeto - Código do projeto (ex: "260202cBR")
 * @param {string} data - Data no formato YYMMDD
 * @returns {string} - Próxima versão disponível (ex: "", "v2", "v3")
 */
function detectarProximaVersao(codigoProjeto, data) {
  try {
    if (!codigoProjeto || !data) return "";
    
    // Usa a nova função para detectar a pasta
    const pastaInfo = detectarPastaExistente(codigoProjeto, data);
    if (!pastaInfo) return ""; // Primeira versão (sem sufixo)
    
    const pastaProjeto = pastaInfo.pasta;
    
    // Busca na pasta 02_WORK/COM
    let workFolder = null;
    try {
      const workFolders = pastaProjeto.getFoldersByName("02_WORK");
      if (workFolders.hasNext()) {
        workFolder = workFolders.next();
        const comFolders = workFolder.getFoldersByName("COM");
        if (comFolders.hasNext()) {
          const comFolder = comFolders.next();
          const arquivos = comFolder.getFiles();
          const prefixo = "Proposta_" + codigoProjeto;
            const versoesEncontradas = [];
            
            while (arquivos.hasNext()) {
              const arquivo = arquivos.next();
              const nomeArquivo = arquivo.getName();
              if (nomeArquivo.startsWith(prefixo) && nomeArquivo.endsWith(".pdf")) {
                // Verifica se é exatamente "Proposta_<codigo>.pdf" (sem sufixo de versão)
                if (nomeArquivo === prefixo + ".pdf") {
                  versoesEncontradas.push(1); // Sem sufixo = v1
                } else {
                  // Extrai a versão do nome: Proposta_260202cBR_v2.pdf -> "v2"
                  const match = nomeArquivo.match(new RegExp(prefixo + "_v(\\d+)\\.pdf"));
                  if (match && match[1]) {
                    versoesEncontradas.push(parseInt(match[1], 10));
                  }
                }
              }
            }
            
            if (versoesEncontradas.length === 0) return ""; // Primeira versão (sem sufixo)
            
            // Encontra a próxima versão disponível
            const maiorVersao = Math.max(...versoesEncontradas);
            return "v" + (maiorVersao + 1);
          }
        }
      } catch (e) {
        Logger.log("Erro ao buscar versões na pasta 02_WORK/COM: " + e.message);
      }
      
      return ""; // Primeira versão se não encontrar pasta
    } catch (e) {
      Logger.log("Erro ao detectar próxima versão: " + e.message);
      return ""; // Retorna primeira versão em caso de erro
    }
  }

/**
 * Detecta o próximo índice disponível para um usuário em um determinado dia
 * @param {string} data - Data no formato YYMMDD
 * @param {string} iniciais - Iniciais do usuário (ex: "AB")
 * @returns {string} - Próximo índice disponível (ex: "a", "b", "c")
 */
function detectarProximoIndice(data, iniciais) {
  try {
    if (!data || !iniciais) return "a";
    
    const sheetProj = ss.getSheetByName("Projetos");
    if (!sheetProj) return "a";
    
    const lastRow = sheetProj.getLastRow();
    if (lastRow < 2) return "a"; // Primeiro projeto do dia
    
    // Lê todas as linhas da planilha
    const numCols = PROJETOS_NUM_COLUNAS;
    const dados = sheetProj.getRange(2, 1, lastRow - 1, numCols).getValues();
    
    // Lista de índices já usados neste dia para estas iniciais
    const indicesUsados = [];
    
    dados.forEach((row) => {
      const projeto = String(row[3] || ""); // Coluna PROJETO (índice 3)
      const dataProjeto = String(row[5] || ""); // Coluna DATA (índice 5)
      
      // Verifica se é do mesmo dia e tem as mesmas iniciais
      if (projeto.length >= 8 && projeto.substring(0, 6) === data) {
        const resto = projeto.substring(6);
        if (resto.length > 0) {
          const indice = resto.charAt(0);
          const iniciaisProjeto = resto.substring(1);
          
          if (iniciaisProjeto === iniciais) {
            indicesUsados.push(indice.toLowerCase());
          }
        }
      }
    });
    
    // Se não há índices usados, retorna "a"
    if (indicesUsados.length === 0) return "a";
    
    // Encontra o próximo índice disponível
    const letras = "abcdefghijklmnopqrstuvwxyz";
    for (let i = 0; i < letras.length; i++) {
      const letra = letras[i];
      if (!indicesUsados.includes(letra)) {
        return letra;
      }
    }
    
    // Se todas as letras foram usadas (improvável), retorna "z"
    return "z";
  } catch (e) {
    Logger.log("Erro ao detectar próximo índice: " + e.message);
    return "a"; // Retorna "a" em caso de erro
  }
}

// Cria (ou retorna) a pasta do orçamento SEM criar a subpasta 01_IN.
// A subpasta 01_IN só será criada quando arquivos forem enviados.
// Usa a mesma lógica de criação de pasta utilizada no gerarPdfOrcamento.
// Modificado para aceitar nomeCliente e isPedido
function criarPastaOrcamento(codigoProjeto, descricao, data, nomeCliente, isPedido) {
  if (!codigoProjeto || !data) {
    throw new Error("Dados do projeto incompletos para criar a pasta (código ou data ausentes).");
  }
  
  if (!descricao || descricao.trim() === "") {
    throw new Error("Descrição do projeto é obrigatória para criar a pasta.");
  }

  const pastaProjeto = criarOuUsarPastaProjeto(
    codigoProjeto,
    nomeCliente || "",
    descricao,
    data,
    isPedido || false
  );

  return {
    pastaId: pastaProjeto.getId(),
    pastaNome: pastaProjeto.getName(),
    pastaUrl: pastaProjeto.getUrl()
  };
}

// Busca apenas a pasta do orçamento SEM criar (retorna erro se não existir)
// Usado pelo botão "Abrir Pasta" que só deve abrir pastas já existentes
// Modificado para retornar o tipo da pasta (COT/PED)
function buscarPastaOrcamento(codigoProjeto, data) {
  if (!codigoProjeto || !data) {
    throw new Error("Dados do projeto incompletos para buscar a pasta (código ou data ausentes).");
  }

  const pastaInfo = detectarPastaExistente(codigoProjeto, data);
  
  if (!pastaInfo) {
    throw new Error("Pasta do orçamento não encontrada. Crie a pasta primeiro usando o botão 'Criar/Confirmar Pasta do Orçamento'.");
  }

  // Busca pasta 01_IN se existir
  let inFolder = null;
  try {
    const inFolders = pastaInfo.pasta.getFoldersByName("01_IN");
    if (inFolders.hasNext()) {
      inFolder = inFolders.next();
    }
  } catch (e) {
    // 01_IN pode não existir ainda, mas a pasta principal existe
  }

  return {
    pastaId: pastaInfo.pasta.getId(),
    pastaNome: pastaInfo.pasta.getName(),
    pastaUrl: pastaInfo.pasta.getUrl(),
    inFolderId: inFolder ? inFolder.getId() : null,
    inFolderNome: inFolder ? inFolder.getName() : null,
    inFolderUrl: inFolder ? inFolder.getUrl() : null,
    existe: true,
    tipo: pastaInfo.tipo // Retorna COT ou PED
  };
}

// Recebe arquivos enviados pelo formulário e salva dentro da pasta 01_IN do projeto.
// A pasta do projeto é criada/obtida usando a mesma lógica do orçamento calculado.
// IMPORTANTE: Quando há file inputs, o formulário deve ser o único parâmetro.
// Os dados do projeto (codigoProjeto, nomePasta, data) vêm em campos hidden do formulário.
function salvarArquivosCliente(formObject) {
  if (!formObject) {
    throw new Error("Formulário inválido ao salvar arquivos do cliente.");
  }

  // Extrai dados do projeto dos campos hidden do formulário
  const codigoProjeto = formObject.codigoProjeto || "";
  const nomePasta = formObject.nomePasta || "";
  const data = formObject.dataProjeto || "";

  if (!codigoProjeto || !data) {
    throw new Error("Dados do projeto incompletos para salvar arquivos (código ou data ausentes). Verifique se os campos do projeto estão preenchidos.");
  }

  const pastaProjeto = criarOuUsarPasta(codigoProjeto, nomePasta || "", data);
  const inFolder = getOrCreateSubFolder(pastaProjeto, "01_IN");

  // Campo de arquivos no formulário (name="arquivosCliente")
  let arquivos = formObject.arquivosCliente;
  if (!arquivos) {
    // Nada para salvar
    return {
      ok: true,
      quantidade: 0,
      pastaNome: pastaProjeto.getName(),
      inFolderNome: inFolder.getName(),
      arquivos: []
    };
  }

  // Garante que seja um array
  if (!Array.isArray(arquivos)) {
    arquivos = [arquivos];
  }

  const salvos = [];

  arquivos.forEach(function (blob) {
    if (!blob) return;

    // Mantém o nome original do arquivo, se disponível
    let nomeArquivo = "";
    try {
      if (typeof blob.getName === "function") {
        nomeArquivo = blob.getName();
      }
    } catch (e) {
      // fallback silencioso
    }

    const file = inFolder.createFile(blob);
    if (nomeArquivo && file.getName() !== nomeArquivo) {
      file.setName(nomeArquivo);
    }

    salvos.push({
      id: file.getId(),
      nome: file.getName(),
      url: file.getUrl()
    });
  });

  return {
    ok: true,
    quantidade: salvos.length,
    pastaNome: pastaProjeto.getName(),
    inFolderNome: inFolder.getName(),
    arquivos: salvos
  };
}

// ========================= GERAR PDF (VERSÃO AJUSTADA) =========================
function gerarPdfOrcamento(
  chapas, cliente, observacoes, codigoProjeto, nomePasta, dataProjeto, versao, somaProcessosPedido, descricaoProcessosPedido, produtosCadastrados, dadosFormularioCompleto, infoPagamento, isPedido
) {
  try {

    // Incrementa contador de propostas
    incrementarContador("totalPropostas");
    
    // Obtém e incrementa o número sequencial do orçamento
    const numeroSequencial = obterEIncrementarNumeroOrcamento();

    // Persiste numeroSequencial em dadosFormularioCompleto
    if (dadosFormularioCompleto) {
      dadosFormularioCompleto.numeroSequencial = numeroSequencial;
    }

    const resultados = calcularOrcamento(chapas);

    // Atribui códigos PRD a produtos cadastrados que não têm código
    if (produtosCadastrados && Array.isArray(produtosCadastrados)) {
      atribuirCodigosPRDAutomaticos(produtosCadastrados);
    }

    // Adiciona produtos cadastrados aos resultados
    if (produtosCadastrados && Array.isArray(produtosCadastrados)) {
      produtosCadastrados.forEach(prod => {
        resultados.push({
          codigo: prod.codigo || "",
          descricao: prod.descricao || "",
          quantidade: prod.quantidade || 0,
          precoUnitario: prod.precoUnitario || 0,
          precoTotal: prod.precoTotal || 0
        });
      });
    }

    // Usa nova estrutura de pastas
    const nomeCliente = cliente.nome || "";
    const descricao = observacoes.descricao || nomePasta || codigoProjeto;
    const pasta = criarOuUsarPastaProjeto(codigoProjeto, nomeCliente, descricao, dataProjeto, isPedido || false);
    const workFolder = getOrCreateSubFolder(pasta, "02_WORK");
    const comSubFolder = getOrCreateSubFolder(workFolder, "COM");

    // Logo
    const logoFile = DriveApp.getFileById(ID_LOGO);
    const logoBlob = logoFile.getBlob();
    const logoBase64 = Utilities.base64Encode(logoBlob.getBytes());
    const logoMime = logoBlob.getContentType();

    // Totais
    const totalPecas = resultados.reduce((sum, p) => sum + (Number(p.precoTotal) || 0), 0);
    const totalFinal = totalPecas + (Number(somaProcessosPedido) || 0);

    // Data/hora
    const agora = new Date();
    const dataBrasil = formatarDataBrasil(agora);
    const horaBrasil = agora.toLocaleTimeString("pt-BR");

    // Detecta automaticamente a versão se não foi fornecida
    let versaoFinal = versao || "";
    if (!versaoFinal) {
      versaoFinal = "_" + detectarProximaVersao(codigoProjeto, dataProjeto);
    }

    const numeroProposta = (codigoProjeto || "") + (versaoFinal || "");

    // cores
    const headerColor = "#FF9933"; // cabeçalho (laranja médio)
    const rowColor = "#FDF5E6";    // linhas / totais (laranja claro)

    // helpers
    function esc(v) {
      if (v === null || v === undefined) return "";
      return String(v)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }

    function formatBR(n) {
      const num = Number(n) || 0;
      const parts = num.toFixed(2).split('.');
      parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ".");
      return "R$ " + parts[0] + ',' + parts[1];
    }

    // NOVA FUNÇÃO: Calcula parcelas baseado no texto de pagamento
    function calcularParcelas(textoPagamento, valorTotal) {
      if (!textoPagamento || textoPagamento.trim() === "") {
        return null;
      }

      const texto = textoPagamento.trim().toUpperCase();

      // Se for "À vista" ou "30 dias" (parcela única), retorna null (não precisa de tabela)
      if (texto.includes("VISTA") || texto === "30 DIAS" || !texto.includes("/")) {
        return null;
      }

      // Extrai os números de dias (ex: "30 / 45 / 60" -> [30, 45, 60])
      const diasMatch = texto.match(/\d+/g);
      if (!diasMatch || diasMatch.length === 0) {
        return null;
      }

      const dias = diasMatch.map(d => parseInt(d, 10));
      const numParcelas = dias.length;
      const valorParcela = valorTotal / numParcelas;

      // Retorna array de objetos com dia e valor
      return dias.map((dia, idx) => ({
        numero: idx + 1,
        dias: dia,
        valor: valorParcela
      }));
    }

    const itensHtml = resultados.map(function (p) {
      return ''
        + '<tr>'
        + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt;">${esc(p.codigo || "")}</td>`
        + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt;">${esc(p.descricao || "")}</td>`
        + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">${esc(p.quantidade || 0)}</td>`
        + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">${formatBR(p.precoUnitario || 0)}</td>`
        + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">${formatBR(p.precoTotal || 0)}</td>`
        + '</tr>';
    }).join('');

    const processosPedidoRow = (somaProcessosPedido && Number(somaProcessosPedido) > 0)
      ? ''
      + '<tr>'
      + `<td colspan="4" bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;"><strong>${esc(descricaoProcessosPedido || "")}</strong></td>`
      + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">${formatBR(somaProcessosPedido)}</td>`
      + '</tr>'
      : '';

    // NOVO: Gera tabela de parcelas se houver múltiplas parcelas
    let tabelaParcelasHtml = "";
    if (infoPagamento && infoPagamento.texto) {
      const parcelas = calcularParcelas(infoPagamento.texto, totalFinal);

      if (parcelas && parcelas.length > 1) {
        tabelaParcelasHtml = `
    <table cellpadding="1" cellspacing="1" style="width:auto; max-width:200px; border-collapse:collapse; margin-top:10px; margin-right:auto; font-size:7pt;">
      <tr>
        <th colspan="3" bgcolor="${headerColor}" style="background:${headerColor}; color:#fff; padding:2px; text-align:center; font-size:9pt; font-weight:bold;">
           Pagamento
        </th>
      </tr>
      <tr>
        <th bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt; text-align:center;">Parc.</th>
        <th bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt; text-align:center;">Dias</th>
        <th bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt; text-align:center;">Valor</th>
      </tr>
      ${parcelas.map(p => `
        <tr>
          <td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt; text-align:center;">${p.numero}/${parcelas.length}</td>
          <td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt; text-align:center;">${p.dias}</td>
          <td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt; text-align:center;">${formatBR(p.valor)}</td>
        </tr>
      `).join('')}
    </table>
    `;
      }
    }

    const htmlContent = `
      <html>
      <head>
        <meta charset="utf-8">
        <style>
          body, table, th, td { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
          body { font-family: Arial, sans-serif; font-size: 8pt; color: #000; margin: 0px; line-height:1.2; -webkit-font-smoothing:antialiased; } /* margem ainda menor */
          .header { display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; }
          .logo { max-height:160px; }
          .company-info { text-align:right; font-size:8pt; }
          h2 { text-align:left; margin:15px 0 20px 0; font-size:12pt; } /* reduzido */
          h3 { margin-top:12px; margin-bottom:4px; font-size:10pt; }
          table { width:100%; border-collapse:collapse; border-spacing:0; font-size:6pt; }
        </style>
      </head>
      <body style="-webkit-print-color-adjust: exact; print-color-adjust: exact;">
        <div class="header">
          <img class="logo" src="data:${logoMime};base64,${logoBase64}">
          <div class="company-info">
            <strong>TUBA FERRAMENTARIA LTDA</strong><br>
            CNPJ: 10.684.825/0001-26<br>
            Inscrição Estadual: 635592888110<br>
            Endereço: Estrada Dos Alvarengas, 4101 - Assunção<br>
            São Bernardo do Campo - SP - CEP: 09850-550<br>
            Site: www.tb4.com.br<br>
            <b>Email:</b> tubaferram@gmail.com<br>
            <b>Telefone:</b> (11) 91285-4204
            
          </div>
        </div>

        <h2>Proposta Comercial Nº ${esc(numeroProposta)}</h2><br>
        <h2><strong>Orçamento Nº ${numeroSequencial}</strong></h2>
        </p>

        <h3>Informações do Cliente:</h3>
        <p style="margin-bottom:12px; font-size:9pt; line-height:1.3;">
          <p><strong>${esc(cliente.nome)}</strong><br></p>
            CNPJ/CPF: ${esc(cliente.cpf)}<br>
            ${esc(cliente.endereco)}<br>
            <b>Telefone:</b> ${esc(cliente.telefone)}<br>
            <b>Email:</b> ${esc(cliente.email)}<br>
            <b>Responsável:</b> ${esc(cliente.responsavel || "-")}
        </p>

        <h3>Itens da Proposta Comercial</h3>
        <table style="margin-top:8px;">
          <tr>
            <th bgcolor="${headerColor}" style="background:${headerColor}; color:#ffffff; padding:3px; text-align:left; border:0.1px solid #fff; font-size:9pt;">Código</th>
            <th bgcolor="${headerColor}" style="background:${headerColor}; color:#ffffff; padding:3px; text-align:left; border:0.1px solid #fff; font-size:9pt;">Descrição</th>
            <th bgcolor="${headerColor}" style="background:${headerColor}; color:#ffffff; padding:3px; text-align:right; border:0.1px solid #fff; font-size:9pt;">Quant.</th>
            <th bgcolor="${headerColor}" style="background:${headerColor}; color:#ffffff; padding:3px; text-align:right; border:0.1px solid #fff; font-size:9pt;">Unit.</th>
            <th bgcolor="${headerColor}" style="background:${headerColor}; color:#ffffff; padding:3px; text-align:right; border:0.1px solid #fff; font-size:9pt;">Valor Total</th>
          </tr>
          ${itensHtml}
          ${processosPedidoRow}
        </table>

        <!-- Totais alinhados com a coluna Valor Total -->
<div style="width:100%; text-align:right; margin-top:5px;">
  <table style="display:inline-block; border-collapse:collapse; width:100%; max-width:280px;">
    <tr>
      <td style="border:none; text-align:right; width:120px; background:#fff; padding:3px; font-weight:bold; font-size:8pt;">Subtotal:</td>
      <td style="border:none; text-align:right; background:${rowColor}; padding:3px; width:100px; font-weight:bold; font-size:8pt;">${formatBR(totalPecas)}</td>
    </tr>
    <tr>
      <td style="border:none; text-align:right; background:#fff; padding:3px; font-weight:bold; font-size:8pt;">Total:</td>
      <td style="border:none; text-align:right; background:${rowColor}; padding:3px; width:100px; font-weight:bold; font-size:8pt;">${formatBR(totalFinal)}</td>
    </tr>
  </table>
</div>

        ${tabelaParcelasHtml}

        <h3 style="margin-top:12px;">Outras Informações</h3>
        <p style="font-size:8pt; line-height:1.25;">
          <b>Proposta Comercial - incluído em:</b> ${esc(dataBrasil)} às ${esc(horaBrasil)}<br>
          <b>Validade da Proposta:</b> 30 dias
        </p>

        <p style="font-size:8pt; line-height:1.25;">
          <b>Previsão de Faturamento:</b> ${esc(formatarDataBrasil(observacoes.faturamento) || "-")}<br>
          <b>Pagamento:</b> ${esc(observacoes.pagamento || "-")}<br>
          <b>Vendedor:</b> ${esc(observacoes.vendedor || "-")}<br>
        </p>

        <p style="font-size:8pt; line-height:1.25;">
          <b>PROJ:</b> ${esc(observacoes.projeto || "-")}<br>
          <b>Condições do Material:</b> ${esc(observacoes.materialCond || "-")}<br>
        </p>

        ${observacoes.adicional ? `<p style="font-size:8pt; line-height:1.25;"><b>Observações adicionais:</b><br>${esc(observacoes.adicional)}</p>` : ""}

      </body>
      </html>
    `;

    const blob = Utilities.newBlob(htmlContent, "text/html", "orcamento.html");
    const pdf = blob.getAs("application/pdf").setName("Proposta_" + numeroProposta + ".pdf");
    const file = comSubFolder.createFile(pdf);

    let memoriaUrl = null;
    try {
      const memoria = gerarPdfMemoriaCalculo(chapas, cliente, codigoProjeto, comSubFolder, file.getName(), produtosCadastrados);
      memoriaUrl = memoria && memoria.url ? memoria.url : null;
    } catch (eMem) {
      Logger.log("Erro ao gerar memoria de calculo: " + eMem.toString());
    }

    registrarOrcamento(cliente, codigoProjeto, totalFinal, dataBrasil, file.getUrl(), memoriaUrl, chapas, observacoes, produtosCadastrados, dadosFormularioCompleto, isPedido);
    return { url: file.getUrl(), nome: file.getName(), memoriaUrl: memoriaUrl };
  } catch (err) {
    Logger.log("ERRO gerarPdfOrcamento: " + err.toString());
    throw err;
  }
}

/* ======= GERAR PDF ORDEM DE PRODUÇÃO (sem valores) ======= */
function gerarPdfOrdemProducao(linhaOuKey) {
  try {
    // Carrega os dados do orçamento
    const dados = carregarRascunho(linhaOuKey);
    if (!dados) {
      throw new Error("Não foi possível carregar os dados do orçamento");
    }

    // Extrai dados necessários
    const chapas = dados.chapas || [];
    const cliente = dados.cliente || {};
    const observacoes = dados.observacoes || {};
    const projeto = dados.projeto || {};
    const processosPedido = dados.processosPedido || [];
    const produtosCadastrados = dados.produtosCadastrados || [];
    const numeroSequencial = dados.numeroSequencial || null;

    const codigoProjeto = (projeto.data || "") + (projeto.indice || "") + (projeto.iniciais || "");
    const data = projeto.data || "";
    // Detecta automaticamente a versão se não foi fornecida
    let versaoFinal = projeto.versao || "";
    if (!versaoFinal) {
      versaoFinal = detectarProximaVersao(codigoProjeto, data);
    }
    const numeroProposta = (codigoProjeto || "") + (versaoFinal || "");

    // Calcula resultados (mas sem mostrar valores)
    const resultados = calcularOrcamento(chapas);

    // Adiciona produtos cadastrados aos resultados (sem valores)
    if (produtosCadastrados && Array.isArray(produtosCadastrados)) {
      produtosCadastrados.forEach(prod => {
        resultados.push({
          codigo: prod.codigo || "",
          descricao: prod.descricao || "",
          quantidade: prod.quantidade || 0,
          precoUnitario: 0,
          precoTotal: 0
        });
      });
    }

    // Busca pasta (data já definido acima como projeto.data)
    const nomePasta = projeto.pasta || "";
    const pasta = criarOuUsarPasta(codigoProjeto, nomePasta, data);
    const workFolder = getOrCreateSubFolder(pasta, "02_WORK");
    const comSubFolder = getOrCreateSubFolder(workFolder, "COM");

    // Logo
    const logoFile = DriveApp.getFileById(ID_LOGO);
    const logoBlob = logoFile.getBlob();
    const logoBase64 = Utilities.base64Encode(logoBlob.getBytes());
    const logoMime = logoBlob.getContentType();

    // Data/hora
    const agora = new Date();
    const dataBrasil = formatarDataBrasil(agora);
    const horaBrasil = agora.toLocaleTimeString("pt-BR");

    // cores
    const headerColor = "#FF9933";
    const rowColor = "#FDF5E6";

    // helpers
    function esc(v) {
      if (v === null || v === undefined) return "";
      return String(v)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }

    // Gera HTML dos itens SEM valores
    const itensHtml = resultados.map(function (p) {
      return ''
        + '<tr>'
        + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:8pt;">${esc(p.codigo || "")}</td>`
        + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:8pt;">${esc(p.descricao || "")}</td>`
        + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; text-align:center; font-size:8pt;">${esc(p.quantidade || 0)}</td>`
        + '</tr>';
    }).join('');

    const htmlContent = `
      <html>
      <head>
        <meta charset="utf-8">
        <style>
          body, table, th, td { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
          body { font-family: Arial, sans-serif; font-size: 9pt; color: #000; margin: 5px; line-height:1.2; -webkit-font-smoothing:antialiased; }
          .header { display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; }
          .logo { max-height:180px; }
          .company-info { text-align:right; font-size:9pt; }
          h2 { text-align:left; margin:20px 0 30px 0; font-size:14pt; }
          h3 { margin-top:15px; margin-bottom:5px; font-size:11pt; }
          table { width:100%; border-collapse:collapse; border-spacing:0; font-size:8pt; }
        </style>
      </head>
      <body style="-webkit-print-color-adjust: exact; print-color-adjust: exact;">
        <div class="header">
          <img class="logo" src="data:${logoMime};base64,${logoBase64}">
          <div class="company-info">
            <strong>TUBA FERRAMENTARIA LTDA</strong><br>
            CNPJ: 10.684.825/0001-26<br>
            Inscrição Estadual: 635592888110<br>
            Endereço: Estrada Dos Alvarengas, 4101 - Assunção<br>
            São Bernardo do Campo - SP - CEP: 09850-550<br>
            Site: www.tb4.com.br<br>
            <b>Email:</b> tubaferram@gmail.com<br>
            <b>Telefone:</b> (11) 91285-4204
          </div>
        </div>

        <h2>Ordem de Produção Nº ${esc(numeroProposta)}</h2>
        ${numeroSequencial ? `<h2><strong>Orçamento Nº ${numeroSequencial}</strong></h2>` : ''}

        <h3>Informações do Cliente:</h3>
        <p style="margin-bottom:12px; font-size:9pt; line-height:1.3;">
          <p><strong>${esc(cliente.nome)}</strong><br></p>
            CNPJ/CPF: ${esc(cliente.cpf)}<br>
            ${esc(cliente.endereco)}<br>
            <b>Telefone:</b> ${esc(cliente.telefone)}<br>
            <b>Email:</b> ${esc(cliente.email)}<br>
            <b>Responsável:</b> ${esc(cliente.responsavel || "-")}
        </p>

        <h3>Itens da Ordem de Produção</h3>
        <table style="margin-top:8px;">
          <tr>
            <th bgcolor="${headerColor}" style="background:${headerColor}; color:#ffffff; padding:3px; text-align:left; border:0.1px solid #fff; font-size:8pt;">Código</th>
            <th bgcolor="${headerColor}" style="background:${headerColor}; color:#ffffff; padding:3px; text-align:left; border:0.1px solid #fff; font-size:8pt;">Descrição</th>
            <th bgcolor="${headerColor}" style="background:${headerColor}; color:#ffffff; padding:3px; text-align:center; border:0.1px solid #fff; font-size:8pt;">Quantidade</th>
          </tr>
          ${itensHtml}
        </table>

        <h3 style="margin-top:12px;">Outras Informações</h3>
        <p style="font-size:8pt; line-height:1.25;">
          <b>Ordem de Produção - gerado em:</b> ${esc(dataBrasil)} às ${esc(horaBrasil)}
        </p>

        <p style="font-size:8pt; line-height:1.25;">
          <b>Previsão de Faturamento:</b> ${esc(formatarDataBrasil(observacoes.faturamento) || "-")}<br>
          <b>Vendedor:</b> ${esc(observacoes.vendedor || "-")}<br>
        </p>

        <p style="font-size:8pt; line-height:1.25;">
          <b>PROJ:</b> ${esc(observacoes.projeto || codigoProjeto || "-")}<br>
          <b>Condições do Material:</b> ${esc(observacoes.materialCond || "-")}<br>
        </p>

        ${observacoes.adicional ? `<p style="font-size:8pt; line-height:1.25;"><b>Observações adicionais:</b><br>${esc(observacoes.adicional)}</p>` : ""}

      </body>
      </html>
    `;

    const blob = Utilities.newBlob(htmlContent, "text/html", "ordem_producao.html");
    const pdf = blob.getAs("application/pdf").setName("Ordem_Producao_" + numeroProposta + ".pdf");
    const file = comSubFolder.createFile(pdf);

    return { url: file.getUrl(), nome: file.getName() };
  } catch (err) {
    Logger.log("ERRO gerarPdfOrdemProducao: " + err.toString());
    throw err;
  }
}

/* ======= gerarPdfMemoriaCalculo corrigido: lê linha de referência APÓS flush ======= */
function gerarPdfMemoriaCalculo(chapas, cliente, codigoProjeto, pastaDestino, nomePdfOrcamento, produtosCadastrados) {
  function formatarNumero(v) {
    if (v === null || v === undefined || v === "") return "";
    const n = Number(v);
    if (isNaN(n)) return String(v);
    return Number.isInteger(n) ? n.toString() : n.toFixed(2).replace(".", ",");
  }

  let htmlMemoria = `

  <html>
  <head>
  <style>
  @page { size: A4; margin: 1mm; }
  body { font-family: Arial; margin: 0; font-size: 10pt; color: #333; }
  table { width: 100%; border-collapse: collapse; margin-bottom: 15px; font-size: 9pt; }
  th, td { border: 1px solid #999; padding: 4px; text-align: center; }
  th { background-color: #eee; }
  .titulo-material { font-weight: bold; font-size: 11pt; margin-top: 20px; }
  .subtitulo-peca { margin-left: 20px; font-weight: bold; font-size: 10pt; margin-bottom: 5px; }
  .titulo-produtos-cadastrados { font-weight: bold; font-size: 12pt; margin-top: 30px; margin-bottom: 10px; background-color: #f0f0f0; padding: 8px; }
  .produto-cadastrado-item { margin-left: 20px; font-size: 10pt; margin-bottom: 8px; }
  </style>
  </head>
  <body>
  <h2>Memória de Cálculo - ${nomePdfOrcamento}</h2>`;

  const capturaCols = 18; // O..AD

  chapas.forEach(chapa => {
    const mat = MATERIAIS[chapa.material];
    if (!mat) return;

    htmlMemoria += `<div class="titulo-material">MATERIAL: ${chapa.material} - Chapa: ${chapa.comprimento}x${chapa.largura}x${chapa.espessura}</div><br>`;

    chapa.pecas.forEach(peca => {
      let processosHtml = "";
      if (peca.processos && peca.processos.length > 0) {
        peca.processos.forEach(processo => {
          const descricaoProcesso = processo.descproc || "-";
          const precoProcesso = formatarNumero(processo.precoProc || 0);

          processosHtml += `<span class="processo-item">&nbsp;&nbsp;&nbsp;&nbsp;- ${descricaoProcesso}: R$ ${precoProcesso}</span><br>`;
        });
      } else {
        processosHtml = "Nenhum processo adicional.";
      }
      htmlMemoria += `<div class="subtitulo-peca">
    Descrição: ${peca.descricao}<br>
    Dimensões: ${peca.comprimento}x${peca.largura}<br> 
    Quantidade do Lote: ${peca.numPecasLote}<br>
    Peças por Chapa: ${peca.numPecasChapa}<br>
    Informações de Processos Adicionais:<br>${processosHtml}<br>
    Totais Adicionais da Peça: R$ ${formatarNumero(peca.adicionaisTotal || 0)}
  </div><br>`;

      // Preenche inputs e força recálculo
      try {
        _preencherInputsCalcParaPeca(mat, chapa, peca);
      } catch (e) {
        Logger.log("Erro preencher inputs (memoria): " + e);
      }
      SpreadsheetApp.flush();

      // Lê a linha de referência O:AD PARA A LINHA ATUAL (após flush)
      let linhaRef = [];
      try {
        linhaRef = SHEET_CALC.getRange(mat.linhaChapa, 15, 1, capturaCols).getValues()[0];
      } catch (e) {
        linhaRef = new Array(capturaCols).fill("");
      }

      htmlMemoria += `<table>
    <tr>
      <th>Preço Kg / Material</th><th>Peso Peça / Chapa</th><th>Peso Lote</th><th>Preço Material Lote</th>
      <th>Nº Trocas Chapa</th><th>Tempo Corte (h)</th><th>Tempo Setup (min)</th>
      <th>Tempo Corte + Setup (h)</th><th>Hora Corte (R$/h)</th><th>Corte Lote (R$)</th><th>Nº Dobras</th>
      <th>Tempo de cada dobra (s)</th><th>Nº Troca de peças</th><th>Total Dobra (h)</th>
      <th>Hora Dobra (R$)</th><th>Total Dobra (R$)</th><th>Preço Unit (R$)</th><th>Preço Total (R$)</th>
    </tr>
    <tr>
      ${linhaRef.map(formatarNumero).map(v => `<td>${v}</td>`).join("")}
    </tr>
  </table>`;
    });
  });

  // Adiciona seção de produtos cadastrados se houver
  if (produtosCadastrados && Array.isArray(produtosCadastrados) && produtosCadastrados.length > 0) {
    htmlMemoria += `<div class="titulo-produtos-cadastrados">PRODUTOS CADASTRADOS</div>`;

    htmlMemoria += `<table>
      <tr>
        <th>Código</th>
        <th>Descrição</th>
        <th>NCM</th>
        <th>Unidade</th>
        <th>Quantidade</th>
        <th>Preço Unitário (R$)</th>
        <th>Preço Total (R$)</th>
      </tr>`;

    produtosCadastrados.forEach(produto => {
      htmlMemoria += `<tr>
        <td>${produto.codigo || "-"}</td>
        <td>${produto.descricao || "-"}</td>
        <td>${produto.ncm || "-"}</td>
        <td>${produto.unidade || "UN"}</td>
        <td>${formatarNumero(produto.quantidade || 0)}</td>
        <td>${formatarNumero(produto.precoUnitario || 0)}</td>
        <td>${formatarNumero(produto.precoTotal || 0)}</td>
      </tr>`;
    });

    htmlMemoria += `</table><br>`;

    // Calcula total dos produtos cadastrados
    const totalProdutosCadastrados = produtosCadastrados.reduce((sum, p) => {
      return sum + (parseFloat(p.precoTotal) || 0);
    }, 0);

    htmlMemoria += `<div class="produto-cadastrado-item">
      <strong>Total de Produtos Cadastrados: R$ ${formatarNumero(totalProdutosCadastrados)}</strong>
    </div><br>`;
  }

  htmlMemoria += `</body></html>`;

  const blobMemoria = Utilities.newBlob(htmlMemoria, "text/html", "memoria.html");
  const pdfMemoria = blobMemoria.getAs("application/pdf").setName("Memoria de Cálculo - " + nomePdfOrcamento);
  const file = pastaDestino.createFile(pdfMemoria);
  return { url: file.getUrl(), nome: file.getName() };
}

function findRowByColumnValue(sheet, colHeader, value) {
  if (!sheet || !colHeader) return null;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(colHeader);
  if (colIndex === -1) return null;
  // lê somente a coluna necessária
  const values = sheet.getRange(2, colIndex + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]) === String(value)) {
      return i + 2; // retorna linha real (considerando header)
    }
  }
  return null;
}


// ----------------- MODIFICAÇÃO: registrarOrcamento -----------------
function registrarOrcamento(cliente, codigoProjeto, valorTotal, dataOrcamento, urlPdf, urlMemoria, chapas, observacoes, produtosCadastrados, dadosFormularioCompleto, isPedido) {
  // Leitura em bloco das colunas H para as faixas de corte/dobra que você utiliza
  const matKeys = Object.keys(MATERIAIS);
  const idxMap = _getMaterialIndexMap().map; // não usado diretamente, mantido por compatibilidade
  // As linhas de corte começam em 20 e vão até 20 + n-1 (conforme seu schema)
  const linhaCorteMin = 20;
  const linhaDobraMin = 28;
  const qtdMat = matKeys.length;

  let valoresCorte = [];
  let valoresDobra = [];
  try {
    valoresCorte = SHEET_CALC.getRange(linhaCorteMin, 8, qtdMat, 1).getValues().flat(); // coluna H
    valoresDobra = SHEET_CALC.getRange(linhaDobraMin, 8, qtdMat, 1).getValues().flat(); // coluna H
  } catch (e) {
    // fallback arrays vazias
    valoresCorte = new Array(qtdMat).fill(0);
    valoresDobra = new Array(qtdMat).fill(0);
  }

  let totalCorte = 0;
  let totalDobra = 0;
  let totalAdicionais = 0;

  chapas.forEach(chapa => {
    const matIdx = Object.keys(MATERIAIS).indexOf(chapa.material);
    if (matIdx < 0) return;
    const corteVal = parseFloat(valoresCorte[matIdx]) || 0;
    const dobraVal = parseFloat(valoresDobra[matIdx]) || 0;

    chapa.pecas.forEach(peca => {
      totalCorte += corteVal;
      totalDobra += dobraVal;
      totalAdicionais += parseFloat(peca.tempoTotal) || 0;
    });
  });

  const processosArray = [];
  if (totalCorte > 0) processosArray.push(`Corte: ${totalCorte.toFixed(2)}h`);
  if (totalDobra > 0) processosArray.push(`Dobra: ${totalDobra.toFixed(2)}h`);
  if (totalAdicionais > 0) processosArray.push(`Adicionais: ${totalAdicionais.toFixed(2)}h`);
  const processosStr = processosArray.join(", ");

  // Extrai descrição e prazo das observações
  const descricao = (observacoes && observacoes.descricao) || "";
  const prazo = (observacoes && observacoes.prazo) || "";

  // Atribui PRD a peças sem código e sincroniza em dadosFormularioCompleto para evitar duplicidade
  chapas.forEach((chapa, chapaIdx) => {
    if (chapa.pecas && Array.isArray(chapa.pecas)) {
      chapa.pecas.forEach((peca) => {
        const codigo = (peca.codigo && String(peca.codigo).trim()) || "";
        if (!codigo || !String(codigo).toUpperCase().startsWith("PRD")) {
          peca.codigo = getProximoCodigoPRD();
        }
      });
    }
  });
  if (dadosFormularioCompleto && dadosFormularioCompleto.chapas && Array.isArray(dadosFormularioCompleto.chapas)) {
    dadosFormularioCompleto.chapas.forEach((chapaDados, chapaIdx) => {
      if (chapas[chapaIdx] && chapaDados.pecas && Array.isArray(chapaDados.pecas)) {
        chapaDados.pecas.forEach((pecaDados, pecaIdx) => {
          if (chapas[chapaIdx].pecas[pecaIdx]) {
            pecaDados.codigo = chapas[chapaIdx].pecas[pecaIdx].codigo;
          }
        });
      }
    });
  }

  // Atribui PRD a produtos cadastrados sem código e sincroniza em dadosFormularioCompleto
  const listaProds = produtosCadastrados || [];
  listaProds.forEach((prod, idx) => {
    const codigo = (prod.codigo && String(prod.codigo).trim()) || "";
    if (!codigo || !String(codigo).toUpperCase().startsWith("PRD")) {
      prod.codigo = getProximoCodigoPRD();
      if (dadosFormularioCompleto && dadosFormularioCompleto.produtosCadastrados && dadosFormularioCompleto.produtosCadastrados[idx]) {
        dadosFormularioCompleto.produtosCadastrados[idx].codigo = prod.codigo;
      }
    }
  });

  // ----- Aqui fazíamos appendRow; agora vamos checar existência e atualizar se necessário -----
  try {
    // Extrai numeroSequencial de dadosFormularioCompleto se disponível
    const numeroSequencial = (dadosFormularioCompleto && dadosFormularioCompleto.numeroSequencial) || null;
    
    // Serializa TODOS os dados do formulário para JSON (para poder reabrir e editar depois)
    const agora = new Date();
    const dadosJson = JSON.stringify({
      nome: codigoProjeto,
      dataSalvo: agora.toISOString(),
      numeroSequencial: numeroSequencial,
      dados: dadosFormularioCompleto || {
        chapas: chapas,
        cliente: cliente,
        observacoes: observacoes,
        produtosCadastrados: produtosCadastrados || []
      }
    });

    // usar: Projetos 
    const sheetProj = ss.getSheetByName("Projetos");
    const targetSheet = sheetProj;

    let rowValues, linhaExistente;

    if (sheetProj) {
      // Nova estrutura: 14 colunas com STATUS_ORCAMENTO e STATUS_PEDIDO separados
      // CLIENTE, DESCRIÇÃO, RESPONSÁVEL CLIENTE, PROJETO, VALOR TOTAL, DATA, PROCESSOS,
      // LINK DO PDF, LINK DA MEMÓRIA DE CÁLCULO, STATUS_ORCAMENTO, STATUS_PEDIDO, PRAZO, OBSERVAÇÕES, JSON_DADOS
      rowValues = [
        cliente.nome || "",
        descricao,
        cliente.responsavel || "",
        codigoProjeto || "",
        valorTotal || "",
        dataOrcamento || "",
        processosStr || "",
        urlPdf || "",
        urlMemoria || "",
        "Enviado",  // STATUS_ORCAMENTO
        "",         // STATUS_PEDIDO (vazio inicialmente)
        prazo,
        "",         // OBSERVAÇÕES (vazio inicialmente)
        dadosJson
      ];
      linhaExistente = findRowByColumnValue(sheetProj, "PROJETO", codigoProjeto);
    }
    if (linhaExistente) {
      // Preserva STATUS_ORCAMENTO, STATUS_PEDIDO e OBSERVAÇÕES ao atualizar (ex.: pedido já convertido)
      const linhaAtual = targetSheet.getRange(linhaExistente, 1, 1, rowValues.length).getValues()[0];
      if (linhaAtual[9]) rowValues[9] = linhaAtual[9]; // STATUS_ORCAMENTO
      if (linhaAtual[10]) rowValues[10] = linhaAtual[10]; // STATUS_PEDIDO
      if (linhaAtual[12]) rowValues[12] = linhaAtual[12]; // OBSERVAÇÕES
      targetSheet.getRange(linhaExistente, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      targetSheet.appendRow(rowValues);
    }

    // Insere produtos com código PRD na "Relação de produtos" ao criar o orçamento (peças das chapas)
    inserirProdutosDasChapas(chapas);

    // Insere também os produtos cadastrados (lista do formulário) que tenham PRD
    if (produtosCadastrados && Array.isArray(produtosCadastrados)) {
      produtosCadastrados.forEach(function (prod) {
        const codigo = (prod.codigo && String(prod.codigo).trim()) || "";
        if (codigo && String(codigo).toUpperCase().startsWith("PRD")) {
          const produtoRelacao = {
            codigo: codigo,
            descricao: prod.descricao || "",
            ncm: prod.ncm || "",
            preco: Number(prod.precoUnitario) || 0,
            unidade: prod.unidade || "UN",
            caracteristicas: ""
          };
          inserirProdutoNaRelacao(produtoRelacao);
        }
      });
    }

  } catch (err) {
    Logger.log("Erro ao registrarOrcamento (atualizar/inserir): " + err);
    // fallback: tentar appendRow (comportamento antigo) se algo falhar
    try {
      // Extrai numeroSequencial de dadosFormularioCompleto se disponível
      const numeroSequencial = (dadosFormularioCompleto && dadosFormularioCompleto.numeroSequencial) || null;
      
      const agora = new Date();
      const dadosJson = JSON.stringify({
        nome: codigoProjeto,
        dataSalvo: agora.toISOString(),
        numeroSequencial: numeroSequencial,
        dados: dadosFormularioCompleto || {
          chapas: chapas,
          cliente: cliente,
          observacoes: observacoes,
          produtosCadastrados: produtosCadastrados || []
        }
      });

      const sheetProj = ss.getSheetByName("Projetos");
      if (sheetProj) {
        // Nova estrutura com 14 colunas
        sheetProj.appendRow([
          cliente.nome || "",
          descricao,
          cliente.responsavel || "",
          codigoProjeto || "",
          valorTotal || "",
          dataOrcamento || "",
          processosStr || "",
          urlPdf || "",
          urlMemoria || "",
          "Enviado",  // STATUS_ORCAMENTO
          "",         // STATUS_PEDIDO
          prazo,
          "",         // OBSERVAÇÕES
          dadosJson
        ]);
      }
      // Insere produtos mesmo no fallback (peças e produtos cadastrados)
      inserirProdutosDasChapas(chapas);
      if (produtosCadastrados && Array.isArray(produtosCadastrados)) {
        produtosCadastrados.forEach(function (prod) {
          const codigo = (prod.codigo && String(prod.codigo).trim()) || "";
          if (codigo && String(codigo).toUpperCase().startsWith("PRD")) {
            inserirProdutoNaRelacao({
              codigo: codigo,
              descricao: prod.descricao || "",
              ncm: prod.ncm || "",
              preco: Number(prod.precoUnitario) || 0,
              unidade: prod.unidade || "UN",
              caracteristicas: ""
            });
          }
        });
      }

    } catch (e2) {
      Logger.log("Erro fallback appendRow em registrarOrcamento: " + e2);
      throw e2;
    }
  }
}

function incrementarContador(tipo) {
  const props = PropertiesService.getScriptProperties();
  const valorAtual = Number(props.getProperty(tipo)) || 0;
  props.setProperty(tipo, valorAtual + 1);
}

/**
 * Obtém e incrementa o número sequencial de orçamentos
 * Começa em 1464 se ainda não existe
 * @returns {number} - Número sequencial do orçamento
 */
function obterEIncrementarNumeroOrcamento() {
  const props = PropertiesService.getScriptProperties();
  const numeroAtual = Number(props.getProperty("numeroOrcamento")) || 1463; // Se não existe, começa em 1463 para que o próximo seja 1464
  const proximoNumero = numeroAtual + 1;
  props.setProperty("numeroOrcamento", proximoNumero);
  return proximoNumero;
}

// ========================= DASHBOARD STATS =========================
function getDashboardStats() {
  const props = PropertiesService.getScriptProperties();

  // Contadores baseados em eventos (propostas e etiquetas)
  const propostas = Number(props.getProperty("totalPropostas")) || 0;
  const etiquetas = Number(props.getProperty("totalEtiquetas")) || 0;

  // Materiais cadastrados
  const materiais = SHEET_MAT ? Math.max(SHEET_MAT.getLastRow() - 1, 0) : 0;

  // Produtos cadastrados
  const produtos = SHEET_PRODUTOS ? Math.max(SHEET_PRODUTOS.getLastRow() - 1, 0) : 0;

  // Logs
  const logs = SHEET_LOGS ? Math.max(SHEET_LOGS.getLastRow() - 1, 0) : 0;

  // Verifica se existe aba Projetos unificada
  const sheetProj = ss.getSheetByName("Projetos");
  let projetos = 0;
  let kanban = 0;

  if (sheetProj) {
    // === NOVA LÓGICA: Conta da aba Projetos ===
    const totalProjetos = Math.max(sheetProj.getLastRow() - 1, 0);
    Logger.log("getDashboardStats: Aba Projetos encontrada, totalProjetos=%s", totalProjetos);

    if (totalProjetos > 0) {
      try {
        const dados = sheetProj.getDataRange().getValues();
        const headers = dados[0];
        Logger.log("getDashboardStats: Headers da aba Projetos: %s", JSON.stringify(headers));
        const idxStatusOrc = _findHeaderIndex(headers, "STATUS_ORCAMENTO");
        const idxStatusPed = _findHeaderIndex(headers, "STATUS_PEDIDO");
        Logger.log("getDashboardStats: idxStatusOrc=%s, idxStatusPed=%s", idxStatusOrc, idxStatusPed);

        for (let i = 1; i < dados.length; i++) {
          const row = dados[i];
          const statusOrc = idxStatusOrc >= 0 ? row[idxStatusOrc] : "";
          const statusPed = idxStatusPed >= 0 ? row[idxStatusPed] : "";

          // Conta orçamentos: projetos que não foram convertidos nem perdidos
          if (statusOrc !== "Expirado/Perdido") {
            projetos++;
          }
          // Kanban: pedidos que não estão finalizados
          if (statusPed !== "Finalizado" && statusOrc !== "Rascunho" && statusOrc !== "Expirado/Perdido" && statusOrc !== "Enviado") {
            kanban++;
          }
        }
        Logger.log("getDashboardStats: Contagem final - projetos=%s, kanban=%s", projetos, kanban);
      } catch (e) {
        Logger.log("Erro ao contar stats da aba Projetos: " + e.message);
      }
    }
  }

  return { propostas, kanban, etiquetas, materiais, logs, produtos, projetos };
}

// --- helper para achar índice de cabeçalho de forma robusta ---
function _normalizeHeader(s) {
  return String(s || '')
    .toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // remove acentos
    .replace(/[^a-z0-9]/g, ''); // remove tudo que não é alfanumérico
}

function _findHeaderIndex(headers, name) {
  const target = _normalizeHeader(name);
  for (let i = 0; i < headers.length; i++) {
    if (_normalizeHeader(headers[i]) === target) return i;
  }
  return -1;
}

function normalizePrazo(value) {
  if (value == null || value === '') return '';
  // Date vindo do getValues()
  if (Object.prototype.toString.call(value) === '[object Date]') {
    try {
      return value.toISOString(); // formato ISO é seguro para serialização
    } catch (e) {
      try { // fallback: format usando timezone da planilha
        const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'UTC';
        return Utilities.formatDate(value, tz, "yyyy-MM-dd'T'HH:mm:ss'Z'");
      } catch (e2) {
        return String(value);
      }
    }
  }
  // se for número -> potencial serial do Sheets (dias desde 1899-12-30)
  if (typeof value === 'number' && !isNaN(value)) {
    try {
      const ms = (value - 25569) * 86400 * 1000;
      const d = new Date(ms);
      if (!isNaN(d.getTime())) return d.toISOString();
    } catch (e) { /* ignore */ }
  }
  // se for string que represente data ISO ou dd/mm/yyyy, tentamos normalizar um pouco
  const s = String(value).trim();
  // tenta interpretar dd/mm/yyyy
  const m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m1) {
    const d = new Date(Number(m1[3]), Number(m1[2]) - 1, Number(m1[1]));
    if (!isNaN(d.getTime())) return d.toISOString();
  }
  // se já parecer ISO, devolve tal qual (ou tenta Date.parse)
  const iso = Date.parse(s);
  if (!isNaN(iso)) return new Date(iso).toISOString();
  // fallback: apenas retornar a string bruta (segura)
  return s;
}

// ===== Atualizada: getKanbanData (usa busca robusta de cabeçalhos) =====
function getKanbanData() {
  try {
    const data = {
      "Processo de Orçamento": [],
      "Processo de Preparação MP / CAD / CAM": [],
      "Processo de Corte": [],
      "Processo de Dobra": [],
      "Processos Adicionais": [],
      "Envio / Coleta": []
    };

    // Verifica se existe a aba Projetos unificada
    const sheetProj = ss.getSheetByName("Projetos");

    if (sheetProj) {
      // === NOVA LÓGICA: Aba Projetos Unificada ===
      const valsProj = sheetProj.getDataRange().getValues();
      if (valsProj && valsProj.length > 1) {
        const headersProj = valsProj[0];
        const idxCliente = _findHeaderIndex(headersProj, "CLIENTE");
        const idxProjeto = _findHeaderIndex(headersProj, "PROJETO");
        const idxDescricao = _findHeaderIndex(headersProj, "DESCRIÇÃO");
        const idxStatusOrc = _findHeaderIndex(headersProj, "STATUS_ORCAMENTO");
        const idxStatusPed = _findHeaderIndex(headersProj, "STATUS_PEDIDO");
        const idxPrazo = _findHeaderIndex(headersProj, "PRAZO");
        const idxProcessos = _findHeaderIndex(headersProj, "PROCESSOS");
        const idxObs = _findHeaderIndex(headersProj, "OBSERVAÇÕES");
        const idxJsonDados = _findHeaderIndex(headersProj, "JSON_DADOS");

        for (let i = 1; i < valsProj.length; i++) {
          const row = valsProj[i];
          const cliente = idxCliente >= 0 ? row[idxCliente] : "";
          const projeto = idxProjeto >= 0 ? row[idxProjeto] : "";
          const descricao = idxDescricao >= 0 ? row[idxDescricao] : "";
          const statusOrc = idxStatusOrc >= 0 ? row[idxStatusOrc] : "";
          const statusPed = idxStatusPed >= 0 ? row[idxStatusPed] : "";
          let prazo = idxPrazo >= 0 ? row[idxPrazo] : "";
          prazo = normalizePrazo(prazo);

          // Cards de orçamento: Somente STATUS_ORCAMENTO = 'RASCUNHO' ou 'Rascunho'
          // e STATUS_PEDIDO vazio
          if (statusOrc && (statusOrc === "RASCUNHO" || statusOrc === "Rascunho") && !statusPed) {
            data["Processo de Orçamento"].push({
              cliente: cliente,
              projeto: projeto,
              descricao: descricao,
              status: statusOrc,
              prazo: prazo
            });
          }

          // Cards de pedido: WHERE STATUS_PEDIDO IS NOT NULL AND STATUS_PEDIDO != '' AND STATUS_PEDIDO != 'Finalizado'
          if (statusPed && statusPed !== "" && statusPed !== "Finalizado") {
            const obs = idxObs >= 0 ? row[idxObs] : "";
            const processosStr = idxProcessos >= 0 ? String(row[idxProcessos] || "") : "";
            const jsonDados = idxJsonDados >= 0 ? row[idxJsonDados] : "";

            // Extrai tempo estimado do campo PROCESSOS
            let tempoEstimado = "";
            if (/Preparação/i.test(statusPed)) {
              tempoEstimado = processosStr.match(/preparação\s*:?\s*([\d.,]+h?)/i)?.[1] || "";
            } else if (/Corte/i.test(statusPed)) {
              tempoEstimado = processosStr.match(/corte\s*:?\s*([\d.,]+h?)/i)?.[1] || "";
            } else if (/Dobra/i.test(statusPed)) {
              tempoEstimado = processosStr.match(/dobra\s*:?\s*([\d.,]+h?)/i)?.[1] || "";
            } else if (/Adicion/i.test(statusPed)) {
              tempoEstimado = processosStr.match(/adici.*:?\s*([\d.,]+h?)/i)?.[1] || "";
            }

            // Extrai temposReais do JSON_DADOS se existir
            let temposReais = {};
            if (jsonDados) {
              try {
                const parsed = JSON.parse(jsonDados);
                if (parsed && parsed.dados && parsed.dados.temposReais) {
                  temposReais = parsed.dados.temposReais;
                }
              } catch (e) {
                // Ignora erros de parse
              }
            }

            // Busca tempo real dos logs (se disponível) - mantido para compatibilidade
            let tempoReal = "";
            const chave = cliente + "|" + projeto;

            if (Array.isArray(data[statusPed])) {
              data[statusPed].push({
                cliente: cliente,
                projeto: projeto,
                descricao: descricao,
                observacoes: obs,
                tempoEstimado: tempoEstimado,
                tempoReal: tempoReal,  // Será preenchido pelos logs abaixo
                temposReais: temposReais, // Novos tempos reais detalhados
                prazo: prazo
              });
            }
          }
        }
      }
    }
    // --- Logs (mapa) - Processa logs para ambas estruturas ---
    const mapaLogs = {};
    if (typeof SHEET_LOGS !== 'undefined' && SHEET_LOGS) {
      const valsLogs = SHEET_LOGS.getDataRange().getValues();
      if (valsLogs && valsLogs.length > 0) {
        const headersLogs = valsLogs[0];
        const idxClienteL = _findHeaderIndex(headersLogs, "Cliente");
        const idxProjetoL = _findHeaderIndex(headersLogs, "Número do Projeto");
        const idxPrep = (_findHeaderIndex(headersLogs, "Tempo estimado / tempo real preparação") >= 0)
          ? _findHeaderIndex(headersLogs, "Tempo estimado / tempo real preparação")
          : (_findHeaderIndex(headersLogs, "Tempo estimado / tempo real de preparação") >= 0)
            ? _findHeaderIndex(headersLogs, "Tempo estimado / tempo real de preparação")
            : _findHeaderIndex(headersLogs, "tempo estimado e tempo real preparação");
        const idxCorte = _findHeaderIndex(headersLogs, "Tempo estimado / tempo real Corte");
        const idxDobra = _findHeaderIndex(headersLogs, "Tempo estimado / tempo real Dobra");
        const idxAdic = _findHeaderIndex(headersLogs, "Tempo estimado / tempo real Adicionais");

        for (let i = 1; i < valsLogs.length; i++) {
          const row = valsLogs[i];
          const chave = String(idxClienteL >= 0 ? row[idxClienteL] : "") + "|" + String(idxProjetoL >= 0 ? row[idxProjetoL] : "");
          mapaLogs[chave] = {
            preparacao_mp_cad_com: idxPrep >= 0 ? row[idxPrep] || "" : "",
            corte: idxCorte >= 0 ? row[idxCorte] || "" : "",
            dobra: idxDobra >= 0 ? row[idxDobra] || "" : "",
            adicionais: idxAdic >= 0 ? row[idxAdic] || "" : ""
          };
        }
      }
    }
    // Aplica tempos reais dos logs aos cards de pedido (para nova estrutura)
    if (sheetProj && Object.keys(mapaLogs).length > 0) {
      Object.keys(data).forEach(coluna => {
        if (coluna !== "Processo de Orçamento" && Array.isArray(data[coluna])) {
          data[coluna].forEach(card => {
            const chave = card.cliente + "|" + card.projeto;
            if (mapaLogs[chave]) {
              if (/Preparação/i.test(coluna)) {
                card.tempoReal = mapaLogs[chave].preparacao_mp_cad_com || "";
              } else if (/Corte/i.test(coluna)) {
                card.tempoReal = mapaLogs[chave].corte || "";
              } else if (/Dobra/i.test(coluna)) {
                card.tempoReal = mapaLogs[chave].dobra || "";
              } else if (/Adicion/i.test(coluna)) {
                card.tempoReal = mapaLogs[chave].adicionais || "";
              }
            }
          });
        }
      });
    }

    return data;
  } catch (e) {
    Logger.log('getKanbanData ERRO (geral): %s\n%s', e && e.message, e && e.stack);
    return {
      "Processo de Orçamento": [],
      "Processo de Preparação MP / CAD / CAM": [],
      "Processo de Corte": [],
      "Processo de Dobra": [],
      "Processos Adicionais": [],
      "Envio / Coleta": []
    };
  }
}

const USUARIOS = {
  "Ivan": { senha: "P4Z", nivel: "admin" },
  "Matheus": { senha: "117082mat", nivel: "admin" },
  "Ana": { senha: "Linda", nivel: "mod" },
  "BrunoMacedo": { senha: "bm4821", nivel: "mod" },
  "BrunoSena": { senha: "bs9374", nivel: "usuario" },
  "IcaroFerreira": { senha: "if6258", nivel: "usuario" },
  "AndreGomes": { senha: "ag7043", nivel: "mod" },
  "Bruna": { senha: "bbbraga123", nivel: "mod" },
  "TV": { senha: "tv123", nivel: "usuario" },
  "Visitante": { senha: "visitante", nivel: "visitante" }
};

// =================== LOGIN ===================
function autenticarUsuario(usuario, senha) {
  if (USUARIOS[usuario] && USUARIOS[usuario].senha === senha) {
    const token = Utilities.getUuid();
    // Armazena usuário e nível no token
    PropertiesService.getScriptProperties().setProperty(token, JSON.stringify({
      usuario: usuario,
      nivel: USUARIOS[usuario].nivel
    }));
    return { success: true, token: token };
  }
  return { success: false };
}

// Retorna nome completo do usuário logado pelo token
function getUsuarioLogadoPorToken(token) {
  const data = PropertiesService.getScriptProperties().getProperty(token);
  if (!data) return null;

  const { usuario, nivel } = JSON.parse(data);

  // Usa o mesmo dicionário da outra função
  const NOMES_COMPLETOS = {
    "BrunoMacedo": "Bruno Macedo Silva",
    "Ivan": "Ivan Braga Ramos",
    "AndreGomes": "André Gomes da Silva",
    "Ana": "Adriana Brauer Braga",
    "Bruna": "Bruna Brauer Braga",
    "Matheus": "Matheus Rodrigues",
    "BrunoSena": "Bruno Sena",
    "IcaroFerreira": "Icaro Ferreira",
    "visitante": "Visitante"
  };

  const INICIAIS = {
    "BrunoMacedo": "MS",
    "Ivan": "BR",
    "AndreGomes": "GS",
    "Ana": "AB",
    "Bruna": "BB",
    "Matheus": "SR",
    "BrunoSena": "SN",
    "IcaroFerreira": "FR",
  };

  const nomeCompleto = NOMES_COMPLETOS[usuario] || usuario;

  const iniciais =
    INICIAIS[usuario] ||
    (() => {
      const partes = nomeCompleto.trim().split(" ");
      if (partes.length === 1) return partes[0].slice(0, 2).toUpperCase();
      const primeira = partes[0][0].toUpperCase();
      const ultima = partes[partes.length - 1][0].toUpperCase();
      return primeira + ultima;
    })();

  return { usuario: nomeCompleto, iniciais, nivel };
}

// =================== AVALIAÇÕES ===================
// Retorna nomes para avaliação, já filtrando o usuário logado
function getAvaliacoesPorUsuario(token) {
  const usuarioLogado = getUsuarioLogadoPorToken(token);

  const equipe = ["Matheus Rodrigues", "Bruno Sena", "Icaro Ferreira"];
  const chefia = ["André Gomes da Silva", "Ivan Braga Ramos", "Bruno Macedo Silva", "Adriana Brauer Braga"];

  return {
    usuarioLogado: usuarioLogado,
    autoavaliacao: [usuarioLogado],                   // só o próprio usuário
    equipe: equipe.filter(nome => nome !== usuarioLogado), // remove usuário logado
    chefia: chefia.filter(nome => nome !== usuarioLogado)  // remove usuário logado
  };
}

// Retorna avaliações já salvas
function getAvaliacoesSalvas(token) {
  const values = SHEET_AVAL.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

// Salva avaliações no Google Sheet
function salvarAvaliacao(avaliacoes, token) {
  try {
    // Pega o usuário logado
    const usuarioObj = getUsuarioLogadoPorToken(token);
    if (!usuarioObj) throw new Error("Usuário não encontrado ou token inválido");

    const usuario = usuarioObj.usuario; // nome do avaliador
    const aval = SHEET_AVAL;

    // Cria cabeçalho se ainda não existir
    if (aval.getLastRow() === 0) {
      aval.appendRow([
        "Avaliador",
        "Tipo",
        "Avaliado",
        "Desempenho",
        "TrabalhoEquipe",
        "Pontualidade",
        "Organizacao",
        "Lideranca",
        "Comunicacao",
        "Observacoes"
      ]);
    }
    // Salva cada avaliação
    avaliacoes.forEach(av => {
      // Avaliado deve sempre ser string
      let avaliado = av.avaliado;

      if (typeof avaliado === "object" && avaliado !== null) {
        // tenta pegar a propriedade 'usuario' ou 'nome'
        avaliado = avaliado.usuario || avaliado.nome || JSON.stringify(avaliado);
      }

      aval.appendRow([
        usuario,               // Avaliador
        av.tipo || "",         // Tipo
        avaliado || "",        // Avaliado como string
        av.desempenho || "",
        av.trabalhoEquipe || "",
        av.pontualidade || "",
        av.organizacao || "",
        av.lideranca || "",
        av.comunicacao || "",
        av.observacoes || ""
      ]);
    });

    return { success: true };

  } catch (e) {
    return { success: false, message: e.message };
  }
}
function doGet(e) {
  let page = e?.parameter?.page || 'login';
  let token = e?.parameter?.token || null;

  const paginasProtegidas = {
    'dashboard': ['admin', 'mod', 'usuario'],
    'formulario': ['admin', 'mod', 'usuario'],
    'materiais': ['admin', 'mod', 'usuario'],
    'geradoretiquetas': ['admin', 'mod', 'usuario'],
    'kanban': ['admin', 'mod', 'usuario'],
    'avaliacoes': ['admin'],
    'projetos': ['admin', 'mod', 'usuario'],
    'avaliacoespage': ['admin'],
    'pedidos': ['admin', 'mod', 'usuario'],
    'logs': ['admin', 'mod'],
    'manutencao': ['admin', 'mod', 'usuario'],
    'manu_registros': ['admin', 'mod', 'usuario'],
    'paginasprotegidas': ['admin'],
    'veiculos': ['admin', 'mod', 'usuario', 'visitante'],
    'veiculos_list': ['admin', 'mod', 'usuario', 'visitante'],
    'produtos': ['admin', 'mod', 'usuario'],

  };

  // Helper que constrói a query de redirecionamento,
  // preservando outros parâmetros além de "page" (se houver)
  function _buildRedirectPath(params, targetPage) {
    const p = Object.assign({}, params || {});
    delete p.page;  // evitar duplicar
    delete p.token; // token será anexado após login
    const qs = Object.keys(p)
      .map(k => encodeURIComponent(k) + '=' + encodeURIComponent(p[k]))
      .join('&');
    return '?page=' + encodeURIComponent(targetPage) + (qs ? '&' + qs : '');
  }

  // ==================== PÁGINAS PROTEGIDAS ====================
  if (paginasProtegidas[page]) {
    const usuarioLogado = getUsuarioLogadoPorToken(token);

    // Se NÃO está logado, servir a página de login e informar para onde redirecionar após login.
    if (!usuarioLogado) {
      const templateLogin = HtmlService.createTemplateFromFile('login');
      // rota de retorno (ex.: ?page=kanban&foo=bar)
      templateLogin.redirectTo = _buildRedirectPath(e?.parameter, page);
      templateLogin.postLoginMsg = "Faça login para acessar: " + page;

      // NOVO: se veio do app com embedded=1, sinalizamos para o login.html
      templateLogin.embedded = (e?.parameter?.embedded === '1');

      return templateLogin.evaluate()
        .setFaviconUrl(FAVICON)
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }

    // Se está logado mas não tem permissão, negar acesso
    if (!paginasProtegidas[page].includes(usuarioLogado.nivel)) {
      return HtmlService.createHtmlOutput("Acesso negado. Você não tem permissão para esta página.");
    }
  }

  // ==================== ROTAS PÚBLICAS / PRINCIPAIS ====================
  try {
    switch (page) {
      case 'login': {
        const templateLoginDefault = HtmlService.createTemplateFromFile('login');
        templateLoginDefault.redirectTo = e?.parameter?.redirectTo || null;

        // NOVO: login "padrão" também pode ser embedido se vier com embedded=1
        templateLoginDefault.embedded = (e?.parameter?.embedded === '1');

        return templateLoginDefault.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      case 'dashboard':
        const templateDashboard = HtmlService.createTemplateFromFile('dashboard');
        templateDashboard.token = token;
        return templateDashboard.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'produtos':
        if (!SHEET_PRODUTOS) throw new Error("Aba 'Relação de produtos' não encontrada");

        const produtosResult = getProdutos();

        const templateProdutos = HtmlService.createTemplateFromFile('produtos');
        templateProdutos.headers = produtosResult.headers;
        templateProdutos.dados = produtosResult.data;
        templateProdutos.token = token;
        return templateProdutos.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'projetos':
        const templateProjetos = HtmlService.createTemplateFromFile('projetos');
        templateProjetos.token = token;
        return templateProjetos.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'paginasprotegidas':
        const templatePaginasProtegidas = HtmlService.createTemplateFromFile('paginasprotegidas');
        templatePaginasProtegidas.token = token;
        return templatePaginasProtegidas.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'formulario':
        const template = HtmlService.createTemplateFromFile('formulario');
        template.token = token;
        return template.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'veiculos':
        const usuario = getUsuarioLogadoPorToken(token);
        const templateVeicForm = HtmlService.createTemplateFromFile('veiculos');
        templateVeicForm.token = token;
        templateVeicForm.usuario = usuario ? usuario.usuario : "Usuário não identificado";
        return templateVeicForm.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'manutencao':
        const templateManutencao = HtmlService.createTemplateFromFile('manutencao');
        templateManutencao.token = token;
        return templateManutencao.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'avaliacoes':
        const templateAval = HtmlService.createTemplateFromFile('avaliacoes');
        templateAval.token = token;
        return templateAval.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'avaliacoespage':
        if (!SHEET_AVAL) throw new Error("Aba 'Registro de Avaliações' não encontrada");

        const avalValues = SHEET_AVAL.getDataRange().getDisplayValues();
        const avalHeaders = avalValues[0];
        const avalData = avalValues.slice(1).map((row, index) => {
          let obj = {};
          avalHeaders.forEach((h, i) => {
            let valor = row[i];
            if (valor instanceof Date) {
              valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
            }
            obj[h] = valor;
          });
          obj["_linhaPlanilha"] = index + 2;
          return obj;
        });

        const templateAvalReg = HtmlService.createTemplateFromFile('avaliacoespage');
        templateAvalReg.dados = avalData;
        templateAvalReg.token = token;
        return templateAvalReg.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'kanban':
        return HtmlService.createTemplateFromFile('kanban')
          .evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'veiculos_list':
        const templateVeiculosList = HtmlService.createTemplateFromFile('veiculos_list');
        templateVeiculosList.token = token;
        return templateVeiculosList.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'geradoretiquetas':
        const templateEtiquetas = HtmlService.createTemplateFromFile('geradoretiquetas');
        templateEtiquetas.token = token;
        return templateEtiquetas.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'materiais':
        if (!SHEET_MAT) throw new Error("Aba 'Controle de Materiais' não encontrada");

        const valuesMat = SHEET_MAT.getDataRange().getDisplayValues();
        const headersMat = valuesMat[0];
        const dataMat = valuesMat.slice(1).map((row) => {
          let obj = {};
          headersMat.forEach((h, i) => {
            let valor = row[i];

            if (h === "DATA" && valor) {
              const dataObj = new Date(valor);
              if (!isNaN(dataObj)) {
                valor = Utilities.formatDate(dataObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
              }
            }

            if (h === "PESO APROXIMADO") {
              valor = parseFloat(valor.toString().replace(',', '.')) || 0;
            }

            obj[h] = valor;
          });
          return obj;
        });

        const templateEtiqTable = HtmlService.createTemplateFromFile('materiais');
        templateEtiqTable.dados = dataMat;
        templateEtiqTable.token = token;
        return templateEtiqTable.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'manutencaologs': // ← NOVO CASE
        if (!SHEET_MANU_NAME) throw new Error("Aba 'Registro de Manutenções' não encontrada");

        const manuValues = SHEET_MANU_NAME.getDataRange().getDisplayValues();
        const manuHeaders = manuValues[0];
        const manuData = manuValues.slice(1).map((row, index) => {
          let obj = {};
          manuHeaders.forEach((h, i) => {
            let valor = row[i];
            if (valor instanceof Date) {
              valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
            }
            obj[h] = valor;
          });
          obj["_linhaPlanilha"] = index + 2;
          return obj;
        });

        const templateManuReg = HtmlService.createTemplateFromFile('manutencaologs');
        templateManuReg.dados = manuData;
        templateManuReg.token = token;
        return templateManuReg.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'seguranca':
        const templateSeguranca = HtmlService.createTemplateFromFile('seguranca');
        return templateSeguranca.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'apresentacao':
        const templateApresentacao = HtmlService.createTemplateFromFile('apresentacao');
        templateApresentacao.token = token;
        return templateApresentacao.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      default:
        return HtmlService.createHtmlOutput("Página não encontrada");
    }

  } catch (err) {
    return HtmlService.createHtmlOutput("Erro ao carregar a página: " + err.message);
  }
}

// ===== Nova função para salvar tempos reais de execução =====
/**
 * Salva o tempo real de início ou fim de um processo no card
 * @param {string} cliente - Nome do cliente
 * @param {string} projeto - Número do projeto
 * @param {string} processoSlug - Slug do processo (ex: "processo-de-corte")
 * @param {string} tipo - 'INICIO' ou 'FIM'
 * @param {string} timestamp - ISO timestamp
 * @param {number} duracaoMinutos - Duração em minutos (apenas para FIM)
 */
function salvarTempoReal(cliente, projeto, processoSlug, tipo, timestamp, duracaoMinutos) {
  try {
    Logger.log('salvarTempoReal: cliente=%s, projeto=%s, processo=%s, tipo=%s', cliente, projeto, processoSlug, tipo);

    // === APENAS salva na aba "TemposReais" ===
    // Removido: salvamento em JSON_DADOS (não é mais necessário)
    salvarTempoRealNaAba(cliente, projeto, processoSlug, tipo, timestamp, duracaoMinutos);

    Logger.log('salvarTempoReal: Sucesso');
    return { success: true };

  } catch (err) {
    Logger.log('salvarTempoReal ERROR: %s\n%s', err.message, err.stack);
    return { success: false, error: err.message };
  }
}

// === Nova função para salvar tempos em aba separada ===
function salvarTempoRealNaAba(cliente, projeto, processoSlug, tipo, timestamp, duracaoMinutos) {
  try {
    // Obtém ou cria a aba TemposReais
    let sheetTempos = ss.getSheetByName("TemposReais");

    if (!sheetTempos) {
      // Cria a aba com cabeçalhos
      sheetTempos = ss.insertSheet("TemposReais");
      sheetTempos.appendRow([
        "CLIENTE",
        "PROJETO",
        "PROCESSO",
        "DATA_HORA_INICIO",
        "DATA_HORA_FIM",
        "DURACAO_MINUTOS",
        "STATUS"
      ]);
      // Formata cabeçalho
      const headerRange = sheetTempos.getRange(1, 1, 1, 7);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#1a73e8");
      headerRange.setFontColor("#ffffff");
    }

    // Converte slug para nome legível
    const nomeProcesso = processoSlug
      .replace(/-/g, ' ')
      .replace(/\b\w/g, l => l.toUpperCase());

    // Converte timestamp ISO para horário local do Brasil (GMT-3)
    function converterParaHorarioBrasil(isoTimestamp) {
      if (!isoTimestamp) return '';
      try {
        const data = new Date(isoTimestamp);
        // Formata no fuso horário de São Paulo (America/Sao_Paulo)
        return Utilities.formatDate(data, 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss');
      } catch (e) {
        Logger.log('Erro ao converter timestamp: ' + e);
        return isoTimestamp; // Retorna original se houver erro
      }
    }

    const timestampFormatado = converterParaHorarioBrasil(timestamp);

    // Busca linha existente para este cliente + projeto + processo
    const dados = sheetTempos.getDataRange().getValues();
    let linhaExistente = null;

    for (let i = 1; i < dados.length; i++) {
      const rowCliente = String(dados[i][0] || '').trim();
      const rowProjeto = String(dados[i][1] || '').trim();
      const rowProcesso = String(dados[i][2] || '').trim();
      const rowStatus = String(dados[i][6] || '').trim();

      // Procura linha com mesmo cliente, projeto, processo e status "EM_EXECUCAO"
      if (rowCliente === String(cliente).trim() &&
        rowProjeto === String(projeto).trim() &&
        rowProcesso === nomeProcesso &&
        rowStatus === 'EM_EXECUCAO') {
        linhaExistente = i + 1;
        break;
      }
    }

    if (tipo === 'INICIO') {
      // Cria nova linha com início
      const novaLinha = [
        cliente,
        projeto,
        nomeProcesso,
        timestampFormatado, // Horário local do Brasil
        '', // DATA_HORA_FIM vazio
        '', // DURACAO_MINUTOS vazio
        'EM_EXECUCAO'
      ];
      sheetTempos.appendRow(novaLinha);

    } else if (tipo === 'FIM' && linhaExistente) {
      // Atualiza linha existente com fim e duração
      sheetTempos.getRange(linhaExistente, 5).setValue(timestampFormatado); // DATA_HORA_FIM
      sheetTempos.getRange(linhaExistente, 6).setValue(duracaoMinutos); // DURACAO_MINUTOS
      sheetTempos.getRange(linhaExistente, 7).setValue('FINALIZADO'); // STATUS
    }

    Logger.log('salvarTempoRealNaAba: Sucesso');

  } catch (err) {
    Logger.log('salvarTempoRealNaAba ERROR: %s\n%s', err.message, err.stack);
    // Não falha a operação principal se houver erro na aba secundária
  }
}

// =================== ETIQUETAS ===================
function gerarEtiqueta(dados, token) {

  const NOMES_COMPLETOS = {
    "BrunoMacedo": "Bruno Macedo Silva",
    "Ivan": "Ivan Braga Ramos",
    "AndreGomes": "André Gomes da Silva",
    "Ana": "Adriana Brauer Braga",
    "Bruna": "Bruna Brauer Braga",
    "Matheus": "Matheus Rodrigues",
    "BrunoSena": "Bruno Sena",
    "IcaroFerreira": "Icaro Ferreira",
  };

  // Incrementa contador de etiquetas e pega o número atualizado
  const props = PropertiesService.getScriptProperties();
  let numEtiqueta = Number(props.getProperty("totalEtiquetas")) || 0;
  numEtiqueta++;
  props.setProperty("totalEtiquetas", numEtiqueta);

  // Adiciona o número da chapa/etiqueta para o template
  dados.numeroChapa = numEtiqueta;

  // Descobre o usuário pelo token
  let usuario = PropertiesService.getScriptProperties().getProperty(token) || "Desconhecido";
  try { usuario = JSON.parse(usuario).usuario; } catch (e) { }
  usuario = usuario.replace(/([a-z])([A-Z])/g, '$1 $2');
  usuario = NOMES_COMPLETOS[usuario] || usuario;

  const agora = new Date();
  const data = Utilities.formatDate(agora, "America/Sao_Paulo", "dd/MM/yy");

  if (dados.dataEntrada && /^\d{4}-\d{2}-\d{2}$/.test(dados.dataEntrada)) {
    const [y, m, d] = dados.dataEntrada.split("-");
    dados.dataEntrada = `${d}/${m}/${y}`;
  }

  // Cria PDF
  const template = HtmlService.createTemplateFromFile("etiqueta");
  template.dados = dados;
  const pdf = template.evaluate()
    .setWidth(105 * 3.78)
    .setHeight(54 * 3.78)
    .getAs("application/pdf");

  // Pastas no Drive
  const ano = agora.getFullYear();
  const ano2 = String(ano).slice(-2);
  const mes = String(agora.getMonth() + 1).padStart(2, "0");
  const dia = String(agora.getDate()).padStart(2, "0");

  const raiz = DriveApp.getFolderById(ID_PASTA_PRINCIPAL);
  const subAno = getOrCreateSubFolder(raiz, String(ano));
  const subMes = getOrCreateSubFolder(subAno, `${ano2}${mes}`);
  const subDia = getOrCreateSubFolder(subMes, `${ano2}${mes}${dia}`);
  const subAdm = getOrCreateSubFolder(subDia, "ADM");
  const subEtiquetas = getOrCreateSubFolder(subAdm, "Etiquetas");

  const nomeArquivo = `ETIQUETA  ${dados.prop || ""} - NFº ${dados.nf || ""} - ${dados.esp || ""} mm - CHAPA #${dados.numeroChapa || ""} - ${usuario}.pdf`;
  const arquivo = subEtiquetas.createFile(pdf.setName(nomeArquivo));
  const urlPdf = arquivo.getUrl();

  // === SALVA NA PLANILHA ===
  const novaLinha = [
    numEtiqueta,
    dados.dataEntrada || "",
    usuario,
    dados.prop || "",
    dados.tipo || "",
    dados.dim || "",
    dados.esp || "",
    dados.material || "",
    dados.qtde || "",
    dados.fornecedor || "",
    dados.nf || "",
    urlPdf,
    "",                     // PESO APROXIMADO (fórmula será inserida depois)


  ];

  SHEET_MAT.appendRow(novaLinha);

  // =================== FÓRMULA DO PESO ===================
  const ultimaLinha = SHEET_MAT.getLastRow();
  const colunaPeso = 13; // coluna M = PESO APROXIMADO
  const f = ultimaLinha; // linha nova
  const formulaNova = `=IF(OR(F${f}="";G${f}="";H${f}="";I${f}="");"";(VALUE(INDEX(SPLIT(REGEXREPLACE(F${f};"[^\\d]+";"x");"x");1))/1000)*(VALUE(INDEX(SPLIT(REGEXREPLACE(F${f};"[^\\d]+";"x");"x");2))/1000)*G${f}*IF(REGEXMATCH(UPPER(H${f});"AÇO|ACO");7,86;IF(REGEXMATCH(UPPER(H${f});"ALUM");2,7;IF(REGEXMATCH(UPPER(H${f});"LAT");8,73;IF(REGEXMATCH(UPPER(H${f});"COBRE");8,96;0))))*I${f})`;

  SHEET_MAT.getRange(ultimaLinha, colunaPeso).setFormula(formulaNova);

  return urlPdf;
}

function gerarNovaEtiqueta(dadosLinha, token) {

  // Acessa a planilha
  const materiais = SHEET_MAT;
  // Lê a linha atual para pegar os valores originais
  const rowIndex = dadosLinha.rowIndex;
  const linhaValores = materiais.getRange(rowIndex, 1, 1, materiais.getLastColumn()).getValues()[0];

  // Monta objeto com os dados da etiqueta
  const dadosEtiqueta = {
    numeroChapa: linhaValores[0],  // Coluna A
    dataEntrada: Utilities.formatDate(new Date(linhaValores[1]), Session.getScriptTimeZone(), "dd/MM/yy"), // Coluna B
    usuario: linhaValores[2],   // Coluna C
    prop: linhaValores[3],      // Coluna D
    tipo: dadosLinha.tipo || linhaValores[4],  // Coluna E
    dim1: (dadosLinha.dim || linhaValores[5]).split('x')[0].trim(), // Coluna F
    dim2: (dadosLinha.dim || linhaValores[5]).split('x')[1]?.trim() || '', // Coluna F
    esp: linhaValores[6],       // Coluna G
    material: linhaValores[7],  // Coluna H
    qtde: linhaValores[8],      // Coluna I
    fornecedor: linhaValores[9],// Coluna J
    nf: linhaValores[10],       // Coluna K
    etiqueta: linhaValores[11]  // Coluna L

  };

  // Cria PDF a partir do template
  const template = HtmlService.createTemplateFromFile("etiqueta");
  template.dados = dadosEtiqueta;
  const pdf = template.evaluate()
    .setWidth(105 * 3.78)
    .setHeight(54 * 3.78)
    .getAs("application/pdf");

  // Cria pastas no Drive
  const agora = new Date();
  const ano = agora.getFullYear();
  const ano2 = String(ano).slice(-2);
  const mes = String(agora.getMonth() + 1).padStart(2, "0");
  const dia = String(agora.getDate()).padStart(2, "0");

  const raiz = DriveApp.getFolderById(ID_PASTA_PRINCIPAL);
  const subAno = getOrCreateSubFolder(raiz, String(ano));
  const subMes = getOrCreateSubFolder(subAno, `${ano2}${mes}`);
  const subDia = getOrCreateSubFolder(subMes, `${ano2}${mes}${dia}`);
  const subAdm = getOrCreateSubFolder(subDia, "ADM");
  const subEtiquetas = getOrCreateSubFolder(subAdm, "Etiquetas");

  // Nome do arquivo
  const nomeArquivo = `ETIQUETA  ${dadosEtiqueta.prop || ""} - NFº ${dadosEtiqueta.nf || ""} - ${dadosEtiqueta.esp || ""} mm - CHAPA #${dadosEtiqueta.numeroChapa || ""} - ${dadosEtiqueta.usuario}.pdf`;

  const arquivo = subEtiquetas.createFile(pdf.setName(nomeArquivo));
  const urlPdf = arquivo.getUrl();

  // Atualiza a coluna ETIQUETA da linha existente
  materiais.getRange(rowIndex, 12).setValue(urlPdf); // Coluna L = ETIQUETA

  return urlPdf;
}

function atualizarCelulaNaPlanilha(linha, campo, novoValor) {
  const colunas = {
    codigo: 1,
    prop: 4,
    tipo: 5,
    dim: 6,
    esp: 7,
    material: 8,
    qtde: 9,
    fornecedor: 10,
    nf: 11,
    etiqueta: 12
  };

  const coluna = colunas[campo];
  if (!coluna) throw new Error("Campo inválido: " + campo);

  if (campo === "qtde") {
    novoValor = Number(novoValor);
    if (novoValor <= 0 || isNaN(novoValor)) {
      SHEET_MAT.deleteRow(linha);
      return "Linha removida";
    } else {
      SHEET_MAT.getRange(linha, coluna).setValue(novoValor);
      return "Quantidade atualizada";
    }
  }

  SHEET_MAT.getRange(linha, coluna).setValue(novoValor);
  return "Valor atualizado";
}

// Cria ou retorna subpasta dentro da pasta pai
function getOrCreateSubFolder(pastaPai, nome) {
  const subPastas = pastaPai.getFoldersByName(nome);
  if (subPastas.hasNext()) {
    return subPastas.next();
  } else {
    return pastaPai.createFolder(nome);
  }
}

// Pega os dados da aba "avaliacoes"
function getAvaliacoes() {
  const valores = SHEET_AVAL.getDataRange().getValues();
  const cabecalho = valores.shift();
  const dados = valores.map((linha, i) => {
    let obj = {};
    cabecalho.forEach((col, j) => obj[col] = linha[j]);
    obj["_linhaPlanilha"] = i + 2; // linha real da planilha
    return obj;
  });
  return { cabecalho, dados };
}

// Excluir linha pelo número
function excluirAvaliacao(linha) {
  SHEET_AVAL.deleteRow(linha);
}

// Exportar tabela para PDF
function exportarAvaliacoesPdf() {
  const url = ss.getUrl().replace(/edit$/, '') +
    'export?format=pdf&gid=' + SHEET_AVAL.getSheetId() +
    '&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false' +
    '&pagenumbers=false&gridlines=true&fzr=false&top_margin=0.25' +
    '&bottom_margin=0.25&left_margin=0.25&right_margin=0.25';
  const token = ScriptApp.getOAuthToken();
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + token }
  });

  const blob = response.getBlob().setName("avaliacoes.pdf");
  const arquivo = DriveApp.getRootFolder().createFile(blob);
  return arquivo.getUrl(); // retorna link para o PDF
}

function getProjetos() {
  try {
    // Tenta usar aba Projetos primeiro
    const sheetProj = ss.getSheetByName("Projetos");
    const targetSheet = sheetProj;

    if (!targetSheet) {
      Logger.log('getProjetos: Nenhuma aba encontrada');
      throw new Error("Nenhuma aba de projetos encontrada");
    }

    const lastRow = targetSheet.getLastRow();
    const lastCol = targetSheet.getLastColumn();
    Logger.log('getProjetos: Sheet name=%s, lastRow=%s, lastCol=%s', targetSheet.getName(), lastRow, lastCol);

    if (lastRow < 2) {
      Logger.log('getProjetos: Nenhum dado encontrado (lastRow < 2)');
      return [];
    }

    const values = targetSheet.getRange(1, 1, lastRow, lastCol).getValues();
    if (!values || values.length === 0) {
      Logger.log('getProjetos: getValues retornou vazio ou null');
      return [];
    }

    const headers = values[0];
    Logger.log('getProjetos: Headers count=%s, first few=%s', headers.length, headers.slice(0, 5).join(','));

    // Formata timezone para datas
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'UTC';

    const data = values.slice(1).map((row, index) => {
      let obj = {};
      headers.forEach((h, i) => {
        let value = row[i];
        // Converte Date para string para evitar problemas de serialização
        if (Object.prototype.toString.call(value) === '[object Date]') {
          try {
            value = Utilities.formatDate(value, tz, 'dd/MM/yyyy');
          } catch (e) {
            value = value.toString();
          }
        }
        // Converte null/undefined para string vazia
        if (value === null || value === undefined) {
          value = '';
        }
        obj[h] = value;
      });
      obj["_linhaPlanilha"] = index + 2;
      return obj;
    });

    Logger.log('getProjetos: Retornando %s projetos', data.length);
    if (data.length > 0) {
      Logger.log('getProjetos: Exemplo primeiro projeto: %s', JSON.stringify(data[0]));
    }

    // Garante que sempre retorna um array
    if (!Array.isArray(data)) {
      Logger.log('getProjetos: AVISO - data não é array, retornando array vazio');
      return [];
    }

    return data;
  } catch (e) {
    Logger.log('getProjetos error: %s\n%s', e.message, e.stack);
    // Em caso de erro, retorna array vazio em vez de lançar exceção
    // para evitar quebrar a interface
    Logger.log('getProjetos: Retornando array vazio devido a erro');
    return [];
  }
}

/**
 * Atualiza dados de um projeto na planilha
 * @param {number} linha - Número da linha na planilha
 * @param {Object} dadosAtualizacao - Objeto com campos a atualizar
 * @returns {Object} - {sucesso: boolean}
 */
function atualizarProjetoNaPlanilha(linha, dadosAtualizacao) {
  try {
    const sheetProj = ss.getSheetByName("Projetos");
    if (!sheetProj) {
      throw new Error("Aba 'Projetos' não encontrada");
    }
    
    // Busca cabeçalhos
    const headers = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, idx) => {
      if (h) headerMap[h.toString().trim().toUpperCase()] = idx;
    });
    
    // Atualiza cada campo
    for (let campo in dadosAtualizacao) {
      const colIdx = headerMap[campo.toUpperCase()];
      if (colIdx !== undefined) {
        sheetProj.getRange(linha, colIdx + 1).setValue(dadosAtualizacao[campo]);
      }
    }
    
    // Se mudou para "Convertido em Pedido", renomeia pasta
    if (dadosAtualizacao.STATUS_ORCAMENTO === "Convertido em Pedido") {
      try {
        Logger.log("🔄 Detectado conversão para pedido, tentando renomear pasta...");
        
        const numCols = sheetProj.getLastColumn();
        const rowData = sheetProj.getRange(linha, 1, 1, numCols).getValues()[0];
        
        const idxProjeto = headerMap["PROJETO"];
        const idxCliente = headerMap["CLIENTE"];
        const idxDescricao = headerMap["DESCRIÇÃO"] || headerMap["DESCRICAO"];
        
        if (idxProjeto !== undefined && rowData[idxProjeto]) {
          const codigoProjeto = String(rowData[idxProjeto]).trim();
          const cliente = idxCliente !== undefined ? String(rowData[idxCliente] || "").trim() : "";
          const descricao = idxDescricao !== undefined ? String(rowData[idxDescricao] || "").trim() : "";
          const dataProj = codigoProjeto.substring(0, 6);
          
          // Renomear pasta de COT para PED
          const sucesso = atualizarPrefixoPastaParaPedido(codigoProjeto, dataProj, cliente, descricao);
          if (sucesso) {
            Logger.log("✅ Pasta convertida de COT para PED: " + codigoProjeto);
          } else {
            Logger.log("⚠️ Não foi possível converter pasta de COT para PED: " + codigoProjeto);
          }
        }
      } catch (e) {
        Logger.log("⚠️ Erro ao renomear pasta de COT para PED: " + e.message);
      }
    }
    
    return { sucesso: true };
  } catch (e) {
    Logger.log("Erro ao atualizar projeto: " + e.message);
    throw new Error("Erro ao atualizar projeto: " + e.message);
  }
}

/**
 * Exclui um projeto da planilha
 * @param {number} linha - Número da linha na planilha
 */
function excluirProjeto(linha) {
  try {
    linha = Number(linha);
    if (!linha || linha < 2) {
      throw new Error('Índice de linha inválido para exclusão: ' + linha);
    }

    // Tenta usar aba Projetos primeiro
    const sheetProj = ss.getSheetByName("Projetos");
    const targetSheet = sheetProj;

    if (!targetSheet) {
      throw new Error("Nenhuma aba de projetos encontrada");
    }

    targetSheet.deleteRow(linha);
    return { success: true };
  } catch (e) {
    Logger.log('excluirProjeto error (linha=%s): %s', linha, e.message);
    throw new Error('excluirProjeto failed: ' + (e.message || 'erro desconhecido'));
  }
}

/**
 * Adiciona um novo projeto na planilha (usado quando projeto já virou pedido externamente)
 * @param {Object} projeto - Objeto com os dados do projeto
 */
function adicionarNovoProjetoNaPlanilha(projeto) {
  try {
    Logger.log('adicionarNovoProjetoNaPlanilha: Iniciando para projeto %s', projeto.PROJETO);

    // Tenta usar aba Projetos primeiro
    const sheetProj = ss.getSheetByName("Projetos");
    const targetSheet = sheetProj;

    if (!targetSheet) {
      throw new Error("Nenhuma aba de projetos encontrada");
    }

    // Verifica se o projeto já existe
    const dados = targetSheet.getDataRange().getValues();
    const headers = dados[0];
    const idxProjeto = headers.indexOf('PROJETO');

    if (idxProjeto >= 0) {
      for (let i = 1; i < dados.length; i++) {
        if (String(dados[i][idxProjeto]).trim() === String(projeto.PROJETO).trim()) {
          throw new Error('Já existe um projeto com este número: ' + projeto.PROJETO);
        }
      }
    }

    // Se é aba Projetos (14 colunas), usa estrutura nova
    if (sheetProj) {
      const novaLinha = [
        projeto.CLIENTE || '',
        projeto['DESCRIÇÃO'] || '',
        projeto['RESPONSÁVEL CLIENTE'] || '',
        projeto.PROJETO || '',
        projeto['VALOR TOTAL'] || '',
        projeto.DATA || new Date().toLocaleDateString('pt-BR'),
        projeto.PROCESSOS || '',
        projeto['LINK DO PDF'] || '',
        projeto['LINK DA MEMÓRIA DE CÁLCULO'] || '',
        projeto.STATUS_ORCAMENTO || 'Convertido em Pedido',
        projeto.STATUS_PEDIDO !== undefined ? projeto.STATUS_PEDIDO : 'Processo de Preparação MP / CAD / CAM',
        projeto.PRAZO || '',
        projeto['OBSERVAÇÕES'] || '',
        projeto.JSON_DADOS || ''
      ];

      targetSheet.appendRow(novaLinha);
      Logger.log('adicionarNovoProjetoNaPlanilha: Projeto adicionado com sucesso na aba Projetos');
    } else {
      throw new Error("Aba Projetos não encontrada");
    }

    return { success: true };
  } catch (e) {
    Logger.log('adicionarNovoProjetoNaPlanilha error: %s\n%s', e.message, e.stack);
    throw new Error('Erro ao adicionar projeto: ' + (e.message || 'erro desconhecido'));
  }
}

function getProdutos() {
  try {
    if (!SHEET_PRODUTOS) throw new Error("Aba 'Relação de produtos' não encontrada");

    const values = SHEET_PRODUTOS.getDataRange().getDisplayValues();
    if (values.length === 0) return { headers: [], data: [] };

    const headers = values[0];
    const data = values.slice(1).map((row, index) => {
      let obj = {};
      headers.forEach((h, i) => {
        let valor = row[i];

        // Formatação de data se a coluna for DATA
        if (h === "DATA" && valor) {
          const dataObj = new Date(valor);
          if (!isNaN(dataObj)) {
            valor = Utilities.formatDate(dataObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
          }
        }

        obj[h] = valor;
      });
      obj["_linhaPlanilha"] = index + 2;
      return obj;
    });

    return { headers: headers, data: data };
  } catch (err) {
    Logger.log("Erro getProdutos: " + err.message);
    throw err;
  }
}

function atualizarStatusKanban(cliente, projeto, novoStatus) {
  try {
    let statusAntigo = '';
    let processosStr = '';

    // Verifica se existe a aba Projetos unificada
    const sheetProj = ss.getSheetByName("Projetos");

    if (sheetProj) {
      // === NOVA LÓGICA: Atualiza STATUS_PEDIDO na aba Projetos ===
      const dadosProj = sheetProj.getDataRange().getValues();
      if (!dadosProj || dadosProj.length < 2) return;

      const headers = dadosProj[0];
      const idxCliente = _findHeaderIndex(headers, "CLIENTE");
      const idxProjeto = _findHeaderIndex(headers, "PROJETO");
      const idxStatusPed = _findHeaderIndex(headers, "STATUS_PEDIDO");
      const idxStatusOrc = _findHeaderIndex(headers, "STATUS_ORCAMENTO");
      const idxDescricao = _findHeaderIndex(headers, "DESCRIÇÃO");
      const idxProcessos = _findHeaderIndex(headers, "PROCESSOS");

      // Valida índices
      if (idxCliente < 0 || idxProjeto < 0 || idxStatusPed < 0) {
        Logger.log('atualizarStatusKanban (Projetos): cabeçalhos não encontrados');
        return;
      }

      for (let i = 1; i < dadosProj.length; i++) {
        const row = dadosProj[i];
        const valCliente = String(row[idxCliente] || '').trim();
        const valProjeto = String(row[idxProjeto] || '').trim();

        if (valCliente === String(cliente).trim() && valProjeto === String(projeto).trim()) {
          statusAntigo = String(row[idxStatusPed] || '').trim();
          processosStr = idxProcessos >= 0 ? String(row[idxProcessos] || '').trim() : '';
          const descricao = idxDescricao >= 0 ? String(row[idxDescricao] || '').trim() : '';

          // Se estava em orçamento e está mudando para um status de pedido, atualiza STATUS_ORCAMENTO também
          if (!statusAntigo && idxStatusOrc >= 0) {
            const statusOrc = String(row[idxStatusOrc] || '').trim();
            if (statusOrc !== "Convertido em Pedido") {
              sheetProj.getRange(i + 1, idxStatusOrc + 1).setValue("Convertido em Pedido");
              
              // Renomeia a pasta de COT para PED
              try {
                const dataProj = valProjeto.substring(0, 6);
                atualizarPrefixoPastaParaPedido(valProjeto, dataProj, valCliente, descricao);
                Logger.log("Pasta convertida de COT para PED: " + valProjeto);
              } catch (e) {
                Logger.log("Erro ao renomear pasta de COT para PED: " + e.message);
              }
            }
          }

          // Atualiza STATUS_PEDIDO
          sheetProj.getRange(i + 1, idxStatusPed + 1).setValue(novoStatus);
          break;
        }
      }
    }
  } catch (e) {
    Logger.log('atualizarStatusKanban error: %s\n%s', e.message, e.stack);
    throw new Error('atualizarStatusKanban failed: ' + (e.message || 'erro desconhecido'));
  }
}

// Número de colunas na nova planilha Projetos unificada
// CLIENTE, DESCRIÇÃO, RESPONSÁVEL CLIENTE, PROJETO, VALOR TOTAL, DATA, PROCESSOS, 
// LINK DO PDF, LINK DA MEMÓRIA DE CÁLCULO, STATUS_ORCAMENTO, STATUS_PEDIDO, PRAZO, OBSERVAÇÕES, JSON_DADOS
const PROJETOS_NUM_COLUNAS = 14;

// ==================== FUNÇÕES DE VALIDAÇÃO E MIGRAÇÃO ====================

/**
 * Verifica se um projeto já existe na aba Projetos
 * @param {string} numeroProjeto - Número do projeto a verificar
 * @returns {Object} { duplicado: boolean, linha: number|null, onde: string }
 */
function verificarProjetoDuplicado(numeroProjeto) {
  try {

    if (!numeroProjeto) {
      return { duplicado: false, linha: null, onde: "" };
    }

    // Busca na aba Projetos (se existir)
    const sheetProj = ss.getSheetByName("Projetos");
    if (sheetProj) {
      const linha = findRowByColumnValue(sheetProj, "PROJETO", numeroProjeto);
      if (linha) {
        return { duplicado: true, linha: linha, onde: "Projetos" };
      }
    }
    return { duplicado: false, linha: null, onde: "" };
  } catch (err) {
    Logger.log("Erro ao verificar projeto duplicado: " + err.message);
    return { duplicado: false, linha: null, onde: "", erro: err.message };
  }
}

function salvarRascunho(nomeRascunho, dados) {
  try {
    // Extrai dados relevantes do formulário
    const clienteNome = (dados.cliente && dados.cliente.nome) || "";
    const descricao = (dados.observacoes && dados.observacoes.descricao) || "";
    const prazo = (dados.observacoes && dados.observacoes.prazo) || "";
    const clienteResponsavel = (dados.cliente && dados.cliente.responsavel) || "";
    const codigoProjeto = (dados.observacoes && dados.observacoes.projeto) || "";
    
    // Validação: Descrição obrigatória
    if (!descricao || descricao.trim() === "") {
      throw new Error("A descrição do projeto é obrigatória para salvar o rascunho.");
    }

    // Validação de duplicidade antes de salvar
    if (codigoProjeto) {
      const validacao = verificarProjetoDuplicado(codigoProjeto);
      // Se existe e não é um rascunho sendo editado, retorna erro
      if (validacao.duplicado) {
        // Verifica se é edição do mesmo projeto (mesma linha)
        const sheetProj = ss.getSheetByName("Projetos");
        const targetSheet = sheetProj;
        const linhaExistente = findRowByColumnValue(targetSheet, "PROJETO", codigoProjeto);

        // Se a linha existe, verifica o status
        if (linhaExistente) {
          const numCols = PROJETOS_NUM_COLUNAS;
          const statusIdx = 9; // STATUS_ORCAMENTO ou STATUS (ambos índice 9)
          const rowData = targetSheet.getRange(linhaExistente, 1, 1, numCols).getValues()[0];
          const status = rowData[statusIdx];

          // Se não é um rascunho, não permite sobrescrever
          if (status !== "RASCUNHO") {
            throw new Error(`Projeto ${codigoProjeto} já existe com status "${status}". Use outra numeração ou edite o projeto existente.`);
          }
        }
      }
    }

    // Garante que a pasta do orçamento já exista para este rascunho (SEM criar 01_IN)
    if (codigoProjeto) {
      try {
        // Extrai componentes do código YYMMDD + índice + iniciais
        const dataProj = codigoProjeto.substring(0, 6); // YYMMDD
        criarPastaOrcamento(codigoProjeto, descricao, dataProj, clienteNome, false);
      } catch (e) {
        Logger.log("Aviso ao criar pasta para rascunho: " + e.message);
      }
    }

    // Data formatada para exibição
    const agora = new Date();
    const dataBrasil = formatarDataBrasil(agora);

    // Atribui códigos PRD a produtos cadastrados que não têm código
    if (dados.produtosCadastrados && Array.isArray(dados.produtosCadastrados)) {
      atribuirCodigosPRDAutomaticos(dados.produtosCadastrados);
    }

    // Serializa todos os dados do formulário para JSON
    const dadosJson = JSON.stringify({
      nome: nomeRascunho,
      dataSalvo: agora.toISOString(),
      dados: dados
    });

    const sheetProj = ss.getSheetByName("Projetos");
    const targetSheet = sheetProj;

    if (!targetSheet) throw new Error("Nenhuma aba de projetos encontrada");

    let rowValues, linhaExistente;

    if (sheetProj) {
      // Nova estrutura: 14 colunas
      // CLIENTE, DESCRIÇÃO, RESPONSÁVEL CLIENTE, PROJETO, VALOR TOTAL, DATA, PROCESSOS,
      // LINK DO PDF, LINK DA MEMÓRIA DE CÁLCULO, STATUS_ORCAMENTO, STATUS_PEDIDO, PRAZO, OBSERVAÇÕES, JSON_DADOS
      rowValues = [
        clienteNome,
        descricao,
        clienteResponsavel,
        codigoProjeto,
        "",  // VALOR TOTAL
        dataBrasil,
        "",  // PROCESSOS
        "",  // LINK DO PDF
        "",  // LINK DA MEMÓRIA DE CÁLCULO
        "RASCUNHO",  // STATUS_ORCAMENTO
        "",          // STATUS_PEDIDO
        prazo,
        "",          // OBSERVAÇÕES
        dadosJson
      ];
      linhaExistente = findRowByColumnValue(sheetProj, "PROJETO", codigoProjeto);
    }

    if (linhaExistente) {
      // Atualiza a linha existente
      targetSheet.getRange(linhaExistente, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      // Cria novo rascunho
      targetSheet.appendRow(rowValues);
    }

    return { success: true };
  } catch (e) {
    Logger.log("Erro ao salvar rascunho: " + e.message);
    throw new Error("Erro ao salvar rascunho: " + e.message);
  }
}

// Nova função: Atualiza apenas os dados do formulário sem mudar o status
// Usada quando o usuário quer atualizar um rascunho sem calcular o orçamento
function atualizarRascunho(linhaOuKey, dados) {
  try {
    const sheetProj = ss.getSheetByName("Projetos");
    const targetSheet = sheetProj;

    if (!targetSheet) throw new Error("Nenhuma aba de projetos encontrada");

    // linhaOuKey é o número da linha na planilha
    const linha = parseInt(linhaOuKey, 10);
    if (isNaN(linha) || linha < 2) {
      throw new Error("Linha inválida: " + linhaOuKey);
    }

    // Verifica se a linha existe
    const lastRow = targetSheet.getLastRow();
    if (linha > lastRow) {
      throw new Error("Orçamento não encontrado");
    }

    // Lê a linha atual para preservar o status e outros campos
    const numCols = PROJETOS_NUM_COLUNAS;
    const rowData = targetSheet.getRange(linha, 1, 1, numCols).getValues()[0];

    // Preserva o status atual (índice 9)
    const statusAtual = rowData[9] || "RASCUNHO";

    // Recalcula o valor total a partir dos dados do formulário (para refletir alterações de preço, etc.)
    let valorTotal = rowData[4] || "";
    try {
      const preview = calcularPreviewOrcamento(dados);
      if (preview && typeof preview.total === "number") {
        valorTotal = preview.total;
      }
    } catch (e) {
      Logger.log("Aviso: não foi possível recalcular total na atualização: " + e.message);
    }

    // Preserva PROCESSOS (índice 6), LINK PDF (índice 7), LINK MEMÓRIA (índice 8), STATUS_PEDIDO (índice 10), OBSERVAÇÕES (índice 12)
    const processos = rowData[6] || "";
    const linkPdf = rowData[7] || "";
    const linkMemoria = rowData[8] || "";
    const statusPedido = rowData[10] || "";
    const observacoes = rowData[12] || "";

    // Extrai dados relevantes do formulário para atualizar
    const clienteNome = (dados.cliente && dados.cliente.nome) || "";
    const descricao = (dados.observacoes && dados.observacoes.descricao) || "";
    const prazo = (dados.observacoes && dados.observacoes.prazo) || "";
    const clienteResponsavel = (dados.cliente && dados.cliente.responsavel) || "";
    const codigoProjeto = (dados.observacoes && dados.observacoes.projeto) || "";
    
    // Validação: Descrição obrigatória
    if (!descricao || descricao.trim() === "") {
      throw new Error("A descrição do projeto é obrigatória para atualizar o rascunho.");
    }

    // Garante que a pasta do orçamento já exista para este rascunho atualizado (SEM criar 01_IN)
    if (codigoProjeto) {
      try {
        const dataProj = codigoProjeto.substring(0, 6); // YYMMDD
        criarPastaOrcamento(codigoProjeto, descricao, dataProj, clienteNome, false);
      } catch (e) {
        Logger.log("Aviso ao criar pasta para atualização de rascunho: " + e.message);
      }
    }

    // Data formatada para exibição
    const agora = new Date();
    const dataBrasil = formatarDataBrasil(agora);

    // Atribui códigos PRD a produtos cadastrados que não têm código
    if (dados.produtosCadastrados && Array.isArray(dados.produtosCadastrados)) {
      atribuirCodigosPRDAutomaticos(dados.produtosCadastrados);
    }

    // Serializa todos os dados do formulário para JSON
    const dadosJson = JSON.stringify({
      nome: codigoProjeto,
      dataSalvo: agora.toISOString(),
      dados: dados
    });

    // Atualiza apenas os campos editáveis, preservando status e outros campos importantes
    const rowValues = [
      clienteNome,           // CLIENTE (0)
      descricao,             // DESCRIÇÃO (1)
      clienteResponsavel,    // RESPONSÁVEL CLIENTE (2)
      codigoProjeto,         // PROJETO (3)
      valorTotal,            // VALOR TOTAL (4) - preservado
      dataBrasil,            // DATA (5) - atualizada
      processos,             // PROCESSOS (6) - preservado
      linkPdf,               // LINK DO PDF (7) - preservado
      linkMemoria,           // LINK DA MEMÓRIA DE CÁLCULO (8) - preservado
      statusAtual,           // STATUS_ORCAMENTO (9) - preservado
      statusPedido,          // STATUS_PEDIDO (10) - preservado
      prazo,                 // PRAZO (11) - atualizado
      observacoes,           // OBSERVAÇÕES (12) - preservado
      dadosJson              // JSON_DADOS (13) - atualizado
    ];

    // Atualiza a linha existente
    targetSheet.getRange(linha, 1, 1, rowValues.length).setValues([rowValues]);

    // Se o orçamento já foi convertido em pedido, gera nova versão do PDF
    if (statusAtual === "Convertido em Pedido" && codigoProjeto) {
      try {
        const chapas = dados.chapas || [];
        const cliente = dados.cliente || {};
        const observacoes = dados.observacoes || {};
        const nomePasta = (dados.projeto && dados.projeto.pasta) || "";
        const dataProjeto = (dados.projeto && dados.projeto.data) ? String(dados.projeto.data).replace(/-/g, "").substring(0, 6) : codigoProjeto.substring(0, 6);
        let somaProcessosPedido = 0;
        const descricoesProcessos = [];
        if (dados.processosPedido && Array.isArray(dados.processosPedido)) {
          dados.processosPedido.forEach(function (proc) {
            const vh = parseFloat(proc.valorHora) || 0;
            const h = parseFloat(proc.horas) || 0;
            const vm = parseFloat(proc.valorMat) || 0;
            const qm = parseFloat(proc.qtdMat) || 0;
            const vf = parseFloat(proc.valorFixo) || 0;
            somaProcessosPedido += vh * h + vm * qm + vf;
            if (proc.descricao) descricoesProcessos.push(proc.descricao);
          });
        }
        const descricaoProcessosPedido = descricoesProcessos.join(" / ");
        const produtosCadastrados = dados.produtosCadastrados || [];
        const infoPagamento = {
          texto: (observacoes.pagamento || "").trim(),
          valorTotal: valorTotal
        };
        const resultPdf = gerarPdfOrcamento(
          chapas,
          cliente,
          observacoes,
          codigoProjeto,
          nomePasta,
          dataProjeto,
          "",
          somaProcessosPedido,
          descricaoProcessosPedido,
          produtosCadastrados,
          dados,
          infoPagamento
        );
        // Atualiza a linha com os novos links do PDF (nova versão)
        if (resultPdf && (resultPdf.url || resultPdf.memoriaUrl)) {
          const newLinkPdf = resultPdf.url || linkPdf;
          const newLinkMemoria = (resultPdf.memoriaUrl != null && resultPdf.memoriaUrl !== "") ? resultPdf.memoriaUrl : linkMemoria;
          targetSheet.getRange(linha, 8, linha, 9).setValues([[newLinkPdf, newLinkMemoria]]);
        }
      } catch (errPdf) {
        Logger.log("Aviso: não foi possível gerar nova versão do PDF ao atualizar pedido: " + errPdf.message);
      }
    }

    return { success: true };
  } catch (e) {
    Logger.log("Erro ao atualizar rascunho: " + e.message);
    throw new Error("Erro ao atualizar rascunho: " + e.message);
  }
}

/**
 * Salva o formulário diretamente como pedido (sem passar por orçamento enviado).
 * O projeto é registrado já com STATUS_ORCAMENTO = "Convertido em Pedido" e STATUS_PEDIDO definido.
 */
function salvarComoPedido(dados) {
  try {
    const cliente = dados.cliente || {};
    const observacoes = dados.observacoes || {};
    const codigoProjeto = (observacoes.projeto || "").trim();
    const clienteNome = (cliente.nome || "").trim();
    const descricao = (observacoes.descricao || "").trim();
    const prazo = (observacoes.prazo || "").trim();
    const clienteResponsavel = (cliente.responsavel || "").trim();

    if (!codigoProjeto || codigoProjeto.length < 8) {
      throw new Error("Código do projeto inválido. Preencha Data, Índice e Iniciais.");
    }
    if (!clienteNome) {
      throw new Error("Nome do cliente é obrigatório.");
    }
    if (!descricao || descricao.trim() === "") {
      throw new Error("Descrição do projeto é obrigatória.");
    }

    const validacao = verificarProjetoDuplicado(codigoProjeto);
    if (validacao.duplicado) {
      throw new Error("Já existe um projeto com o número " + codigoProjeto + ". Use outro número ou carregue o projeto existente para atualizar.");
    }
    
    // Cria pasta com prefixo PED (isPedido=true)
    try {
      const dataProj = codigoProjeto.substring(0, 6);
      criarPastaOrcamento(codigoProjeto, descricao, dataProj, clienteNome, true);
    } catch (e) {
      Logger.log("Aviso ao criar pasta como pedido: " + e.message);
    }

    let valorTotal = 0;
    try {
      const preview = calcularPreviewOrcamento(dados);
      if (preview && typeof preview.total === "number") {
        valorTotal = preview.total;
      }
    } catch (e) {
      Logger.log("Aviso ao calcular total em salvarComoPedido: " + e.message);
    }

    const sheetProj = ss.getSheetByName("Projetos");
    if (!sheetProj) throw new Error("Aba Projetos não encontrada");

    const agora = new Date();
    const dataBrasil = formatarDataBrasil(agora);
    const dadosJson = JSON.stringify({
      nome: codigoProjeto,
      dataSalvo: agora.toISOString(),
      dados: dados
    });

    const rowValues = [
      clienteNome,
      descricao,
      clienteResponsavel,
      codigoProjeto,
      valorTotal,
      dataBrasil,
      "",  // PROCESSOS
      "",  // LINK DO PDF
      "",  // LINK DA MEMÓRIA DE CÁLCULO
      "Convertido em Pedido",  // STATUS_ORCAMENTO
      "Processo de Preparação MP / CAD / CAM",  // STATUS_PEDIDO
      prazo,
      "",  // OBSERVAÇÕES
      dadosJson
    ];

    sheetProj.appendRow(rowValues);
    return { success: true };
  } catch (e) {
    Logger.log("Erro salvarComoPedido: " + e.message);
    throw new Error(e.message || "Erro ao salvar como pedido");
  }
}

// Carrega qualquer orçamento (rascunho ou enviado) pelo número da linha
function carregarRascunho(linhaOuKey) {
  try {
    // Decide qual aba usar
    const sheetProj = ss.getSheetByName("Projetos");
    const targetSheet = sheetProj;

    if (!targetSheet) throw new Error("Nenhuma aba de projetos encontrada");

    // linhaOuKey é o número da linha na planilha
    const linha = parseInt(linhaOuKey, 10);
    if (isNaN(linha) || linha < 2) {
      throw new Error("Linha inválida: " + linhaOuKey);
    }

    // Verifica se a linha existe
    const lastRow = targetSheet.getLastRow();
    if (linha > lastRow) {
      throw new Error("Orçamento não encontrado");
    }

    // Lê a linha da planilha usando a constante apropriada
    const numCols = sheetProj ? PROJETOS_NUM_COLUNAS : ORCAMENTOS_NUM_COLUNAS;
    const rowData = targetSheet.getRange(linha, 1, 1, numCols).getValues()[0];

    // STATUS está no índice 9 em ambas estruturas (STATUS ou STATUS_ORCAMENTO)
    const status = rowData[9];

    // JSON_DADOS está no último índice em ambas estruturas
    const jsonIdx = numCols - 1;
    const dadosJson = rowData[jsonIdx];

    // Se tiver JSON_DADOS, usa os dados completos do formulário
    if (dadosJson) {
      try {
        const dadosParsed = JSON.parse(dadosJson);
        // Incluir numeroSequencial nos dados retornados
        const dadosRetorno = dadosParsed.dados;
        dadosRetorno.numeroSequencial = dadosParsed.numeroSequencial || null;
        return dadosRetorno;
      } catch (parseErr) {
        Logger.log("Erro ao parsear JSON_DADOS na linha " + linha + ": " + parseErr.message);
        // Se falhar o parse, continua para construir dados básicos
      }
    }

    // Se não tiver JSON_DADOS, constrói dados básicos a partir das colunas da planilha
    // Ambas estruturas têm os mesmos índices para campos básicos:
    // CLIENTE(0), DESCRIÇÃO(1), RESPONSÁVEL(2), PROJETO(3), VALOR TOTAL(4), DATA(5), etc.
    const clienteNome = rowData[0] || "";
    const descricao = rowData[1] || "";
    const responsavel = rowData[2] || "";
    const projeto = rowData[3] || "";
  const valorTotal = rowData[4] || "";
  const dataOrcamentoRaw = rowData[5] || "";
  const processos = rowData[6] || "";
    // PRAZO está no índice 11 (nova estrutura) ou 10 (antiga)
  const prazoRaw = sheetProj ? (rowData[11] || "") : (rowData[10] || "");

  // Converte datas do Sheets para string ISO (yyyy-mm-dd) quando forem objetos Date
  let dataOrcamento = "";
  if (dataOrcamentoRaw instanceof Date) {
    dataOrcamento = Utilities.formatDate(dataOrcamentoRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } else {
    dataOrcamento = String(dataOrcamentoRaw || "");
  }

  let prazo = "";
  if (prazoRaw instanceof Date) {
    prazo = Utilities.formatDate(prazoRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
  } else {
    prazo = String(prazoRaw || "");
  }

    // Extrai código do projeto (assumindo formato padrão YYMMDD + índice + iniciais)
    const codigoProjeto = projeto || "";
    let projetoData = "";
    let projetoIndice = "";
    let projetoIniciais = "";

    if (codigoProjeto.length >= 6) {
      projetoData = codigoProjeto.substring(0, 6);
      // Tenta extrair índice (letra) e iniciais
      const resto = codigoProjeto.substring(6);
      if (resto.length > 0) {
        projetoIndice = resto.charAt(0);
        projetoIniciais = resto.substring(1);
      }
    }

  // Busca dados completos do cliente na aba "Cadastro de Clientes"
  let clienteCpf = "";
  let clienteEndereco = "";
  let clienteTelefone = "";
  let clienteEmail = "";

  if (clienteNome) {
    try {
      const dadosClientes = SHEET_CLIENTES.getDataRange().getValues();
      for (let i = 1; i < dadosClientes.length; i++) {
        const rowCli = dadosClientes[i];
        if (rowCli[0] && String(rowCli[0]).trim().toLowerCase() === clienteNome.trim().toLowerCase()) {
          clienteCpf = rowCli[1] || "";
          clienteEndereco = rowCli[2] || "";
          clienteTelefone = rowCli[3] || "";
          clienteEmail = rowCli[4] || "";
          break;
        }
      }
    } catch (e) {
      Logger.log("Erro ao buscar cliente em Cadastro de Clientes: " + e.message);
    }
  }

  // Constrói estrutura básica compatível com o formulário
    const dadosBasicos = {
      projeto: {
        data: projetoData,
        indice: projetoIndice,
        iniciais: projetoIniciais,
      versao: ""
      // Removido campo "pasta" - não é mais usado
      },
      cliente: {
        select: clienteNome,
        nome: clienteNome,
      cpf: clienteCpf,
      endereco: clienteEndereco,
      telefone: clienteTelefone,
      email: clienteEmail,
        responsavel: responsavel,
        data: dataOrcamento
      },
      chapas: [],
      processosPedido: [],
      observacoes: {
      // Usa a data do orçamento como previsão de faturamento padrão (pode ser editada no formulário)
      faturamento: dataOrcamento,
        prazo: prazo,
        vendedor: "",
        materialCond: "",
        pagamento: "",
        adicional: "",
        projeto: codigoProjeto,
        descricao: descricao
      },
      produtosCadastrados: []
    };

    return dadosBasicos;
  } catch (e) {
    Logger.log("Erro ao carregar orçamento: " + e.message);
    throw new Error("Erro ao carregar orçamento: " + e.message);
  }
}

// Retorna lista de orçamentos (rascunhos e/ou enviados) para seleção
// incluirEnviados: se true, inclui também os orçamentos já enviados
// MODIFICADO: Agora inclui TODOS os projetos com número de projeto, mesmo sem JSON_DADOS
function getListaRascunhos(incluirEnviados) {
  try {
    // Decide qual aba usar
    const sheetProj = ss.getSheetByName("Projetos");
    const targetSheet = sheetProj;

    if (!targetSheet) throw new Error("Nenhuma aba de projetos");

    const lastRow = targetSheet.getLastRow();
    if (lastRow < 2) return []; // Sem dados

    // Lê todas as linhas da planilha usando a constante apropriada
    const numCols = sheetProj ? PROJETOS_NUM_COLUNAS : ORCAMENTOS_NUM_COLUNAS;
    const data = targetSheet.getRange(2, 1, lastRow - 1, numCols).getValues();

    const orcamentos = [];
    data.forEach((row, index) => {
      // STATUS_ORCAMENTO ou STATUS está sempre no índice 9
      const status = row[9];
      // JSON_DADOS está sempre no último índice
      const dadosJson = row[numCols - 1];

      const isRascunho = status === "RASCUNHO";

      // Número do projeto (obrigatório para aparecer na lista)
      const projeto = row[3];
      if (!projeto) {
        // Sem número de projeto, não entra na lista
        return;
      }

      // Se incluirEnviados for false, mostra apenas rascunhos
      if (!incluirEnviados && !isRascunho) {
        return;
      }

      const clienteNome = row[0] || "Sem cliente";
      const descricao = row[1] || ""; // Coluna DESCRIÇÃO (índice 1)
      const dataOrcamento = row[5] || ""; // DATA (índice 5)
      // PRAZO está no índice 11 (Projetos)
      const prazo = sheetProj ? (row[11] || "") : (row[10] || "");

      // Tenta extrair o nome do rascunho do JSON (mantido apenas se você quiser usar em futuro ajuste)
      let nomeRascunho = "";
      try {
        if (dadosJson) {
          const parsed = JSON.parse(dadosJson);
          nomeRascunho = parsed.nome || "";
        }
      } catch (e) {
        // Ignora erros de parse
      }

      const linhaReal = index + 2; // +2 porque índice começa em 0 e há cabeçalho

      // Formata a data em formato brasileiro quando for objeto Date
      let dataFormatada = "";
      if (dataOrcamento instanceof Date) {
        dataFormatada = formatarDataBrasil(dataOrcamento);
      } else if (typeof dataOrcamento === "string") {
        dataFormatada = dataOrcamento;
      }

      // Formato: número do projeto + data BR + nome do cliente + descrição (para permitir busca por descrição)
      // Ex: 260112aAB - 12/01/2026 - João da Silva - CORTE DE TUBOS 7mm
      const parteCliente = clienteNome && clienteNome !== "Sem cliente" ? clienteNome : (descricao || "Sem cliente");
      let nomeExibicao = `${projeto} - ${dataFormatada || ""} - ${parteCliente}`;
      if (descricao && parteCliente !== descricao) {
        nomeExibicao += " - " + descricao;
      }

      orcamentos.push({
        key: linhaReal.toString(),
        nome: nomeExibicao,
        status: status
      });
    });

    // Ordena pelo mais recente (maior número de linha = mais recente)
    return orcamentos.sort((a, b) => parseInt(b.key) - parseInt(a.key));
  } catch (e) {
    Logger.log("Erro ao obter lista de orçamentos: " + e.message);
    // Retorna array vazio em caso de erro para não quebrar a UI
    return [];
  }
}

function deletarRascunho(linhaOuKey) {
  const sheetProj = ss.getSheetByName("Projetos");
  try {
    if (!sheetProj) throw new Error("Aba 'Projetos' não encontrada");

    const linha = parseInt(linhaOuKey, 10);
    if (isNaN(linha) || linha < 2) {
      throw new Error("Linha inválida: " + linhaOuKey);
    }

    const lastRow = sheetProj.getLastRow();
    if (linha > lastRow) {
      throw new Error("Rascunho não encontrado");
    }

    // ALTERADO: Permite deletar qualquer orçamento (não apenas rascunhos)
    // A confirmação extra para orçamentos enviados é feita no frontend

    // Remove a linha da planilha
    sheetProj.deleteRow(linha);
    return { success: true };
  } catch (e) {
    Logger.log("Erro ao deletar rascunho: " + e.message);
    throw new Error("Erro ao deletar rascunho: " + e.message);
  }
}

// Retorna a planilha ativa
function getMaintenanceSheet() {

  const sheet = SHEET_MANU_NAME;
  if (!sheet) throw new Error("Aba '" + SHEET_MANU_NAME + "' não encontrada.");
  return sheet;
}

// Registra as manutenções enviadas pelo formulário
function recordMaintenance(tasks) {
  if (!tasks || tasks.length === 0) {
    Logger.log("Nenhuma tarefa para registrar.");
    return { status: "AVISO", totalTasks: 0, row: null };
  }

  const sheet = getMaintenanceSheet();

  const rows = tasks.map(task => {
    // Corrige o formato da data para ISO
    let executionDateTime;
    if (task.date && typeof task.date === "string") {
      const fixedDateStr = task.date.replace(" ", "T");
      executionDateTime = new Date(fixedDateStr);
    } else {
      executionDateTime = new Date();
    }

    // Validação extra
    if (isNaN(executionDateTime)) {
      Logger.log("Data inválida recebida: " + task.date);
      executionDateTime = new Date();
    }

    return [
      task.planName || "",
      executionDateTime,
      task.responsible || "",
      task.frequency || "",
      task.componente || "",
      task.acao || "",
      task.responsavelSugerido || ""
    ];
  });

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);

  return { status: "OK", totalTasks: rows.length, row: startRow };
}

/**
 * Retorna o histórico completo de manutenções
 * @returns {Array<Object>}
 */
function getMaintenanceHistory() {
  const sheet = getMaintenanceSheet();

  if (sheet.getLastRow() < 2) return [];

  const dataRange = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
  const values = dataRange.getValues();

  return values.map(row => {
    const dateObj = row[1]; // Coluna B (data)
    let formattedDate = "";

    // Formata data como string "YYYY-MM-DDTHH:mm"
    if (dateObj instanceof Date) {
      const yyyy = dateObj.getFullYear();
      const mm = String(dateObj.getMonth() + 1).padStart(2, "0");
      const dd = String(dateObj.getDate()).padStart(2, "0");
      const hh = String(dateObj.getHours()).padStart(2, "0");
      const min = String(dateObj.getMinutes()).padStart(2, "0");
      formattedDate = `${yyyy}-${mm}-${dd}T${hh}:${min}`;
    }

    return {
      planName: row[0],
      date: formattedDate,
      responsible: row[2],
      frequency: row[3],
      componente: row[4],
      acao: row[5],
      responsavelSugerido: row[6]
    };
  });
}
/* Save the order for a given status (called from client) */
function saveKanbanOrder(status, ids) {
  if (!status) return { success: false, message: 'status missing' };
  const props = PropertiesService.getScriptProperties();
  let raw = props.getProperty('KANBAN_ORDERS');
  let map = raw ? JSON.parse(raw) : {};
  map[status] = Array.isArray(ids) ? ids : [];
  props.setProperty('KANBAN_ORDERS', JSON.stringify(map));
  return { success: true };
}

/* Return all saved orders as object { status: [ids...] } */
function getKanbanOrders() {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('KANBAN_ORDERS');
    if (!raw) return {};
    try {
      const parsed = JSON.parse(raw);
      return (parsed && typeof parsed === 'object') ? parsed : {};
    } catch (e) {
      Logger.log('getKanbanOrders: JSON.parse falhou para KANBAN_ORDERS; valor cortado: %s', String(raw).slice(0, 1000));
      return {};
    }
  } catch (e) {
    Logger.log('getKanbanOrders ERRO: %s\n%s', e && e.message, e && e.stack);
    return {};
  }
}
function getKanbanDataWithOrders() {
  try {
    const data = (function () {
      try { return getKanbanData(); } catch (e) { Logger.log('getKanbanData lançou: %s\n%s', e && e.message, e && e.stack); return null; }
    })();

    if (!data || typeof data !== 'object') {
      Logger.log('getKanbanDataWithOrders: getKanbanData retornou inválido: %s', String(data));
      return {
        "Processo de Orçamento": [],
        "Processo de Preparação MP / CAD / CAM": [],
        "Processo de Corte": [],
        "Processo de Dobra": [],
        "Processos Adicionais": [],
        "Envio / Coleta": []
      };
    }

    const orders = (function () { try { return getKanbanOrders(); } catch (e) { Logger.log('getKanbanOrders lançou: %s\n%s', e && e.message, e && e.stack); return {}; } })() || {};

    // garante colunas mínimas
    const cols = [
      "Processo de Orçamento",
      "Processo de Preparação MP / CAD / CAM",
      "Processo de Corte",
      "Processo de Dobra",
      "Processos Adicionais",
      "Envio / Coleta"
    ];
    cols.forEach(c => { if (!Array.isArray(data[c])) data[c] = []; });

    Object.keys(data).forEach(status => {
      try {
        const saved = orders[status];
        if (Array.isArray(saved) && saved.length && Array.isArray(data[status])) {
          const map = {};
          data[status].forEach(item => {
            const key = String(item.cliente || '') + '|' + String(item.projeto || '');
            map[key] = item;
          });
          const reordered = [];
          saved.forEach(k => { if (map[k]) { reordered.push(map[k]); delete map[k]; } });
          Object.keys(map).forEach(k => reordered.push(map[k]));
          data[status] = reordered;
        }
      } catch (eStatus) {
        Logger.log('Erro mesclando orders para status %s: %s\n%s', status, eStatus && eStatus.message, eStatus && eStatus.stack);
      }
    });

    return data;
  } catch (e) {
    Logger.log('getKanbanDataWithOrders ERRO (geral): %s\n%s', e && e.message, e && e.stack);
    return {
      "Processo de Orçamento": [],
      "Processo de Preparação MP / CAD / CAM": [],
      "Processo de Corte": [],
      "Processo de Dobra": [],
      "Processos Adicionais": [],
      "Envio / Coleta": []
    };
  }
}
function registrarSaidaVeiculo(dados, token) {
  const user = getUsuarioLogadoPorToken(token);
  if (!user) throw new Error("Usuário não autenticado.");

  // Abra a planilha e aba correta (substitua o ID se necessário)
  const sheet = SHEET_VEIC;
  if (!sheet) throw new Error("Aba 'Controle de Veículos' não encontrada.");

  // Parse do datetime-local enviado pelo cliente (ex: "2025-11-04T13:45")
  // Se o campo vier vazio ou inválido, lidamos de forma segura.
  let saidaDt = null;
  if (dados["HORA SAÍDA"]) {
    // new Date(string) funciona para ISO-like "YYYY-MM-DDTHH:MM"
    saidaDt = new Date(dados["HORA SAÍDA"]);
    if (isNaN(saidaDt.getTime())) {
      // tentativa alternativa: substituir espaço por T (caso)
      const alt = ('' + dados["HORA SAÍDA"]).replace(' ', 'T');
      saidaDt = new Date(alt);
    }
  }

  const nowTz = saidaDt && !isNaN(saidaDt.getTime()) ? saidaDt : new Date();

  const dataFormatada = Utilities.formatDate(nowTz, Session.getScriptTimeZone(), "dd/MM/yyyy"); // DATA
  const horaFormatada = Utilities.formatDate(nowTz, Session.getScriptTimeZone(), "HH:mm"); // HORA SAÍDA

  // Previsão de retorno (opcional) - formatar como "dd/MM/yyyy HH:mm" quando presente e válida
  let previsaoTexto = "";
  if (dados["PREVISÃO RETORNO"]) {
    let retornoDt = new Date(dados["PREVISÃO RETORNO"]);
    if (isNaN(retornoDt.getTime())) {
      const altR = ('' + dados["PREVISÃO RETORNO"]).replace(' ', 'T');
      retornoDt = new Date(altR);
    }
    if (!isNaN(retornoDt.getTime())) {
      previsaoTexto = Utilities.formatDate(retornoDt, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");
    } else {
      // deixa vazio se inválido
      previsaoTexto = "";
    }
  }

  const motivo = dados["MOTIVO"] || "";
  const veiculo = dados["VEÍCULO"] || "";

  const novaLinha = [
    dataFormatada,      // DATA (coluna 1)
    user.usuario,       // FUNCIONÁRIO (coluna 2)
    veiculo,            // VEÍCULO (coluna 3)
    horaFormatada,      // HORA SAÍDA (coluna 4)
    previsaoTexto,      // PREVISÃO RETORNO (coluna 5)
    motivo,             // MOTIVO (coluna 6)
    "Em uso"            // STATUS inicial (coluna 7)
  ];

  sheet.appendRow(novaLinha);
}
function getControleVeiculos() {
  try {
    // Tente usar a variável global se existir
    let sheet = (typeof SHEET_VEIC !== 'undefined' && SHEET_VEIC) ? SHEET_VEIC : null;

    // Se não houver SHEET_VEIC, abra pela ID (substitua 'ID_DA_PLANILHA' pelo ID real)
    if (!sheet) {
      const SPREADSHEET_ID = '1wMIbd8r2HeniFLTYaG8Yhhl8CWmaHaW5oXBVnxj0xos'; // <-- substitua pelo seu ID real
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      sheet = ss.getSheetByName('Controle de Veículos');
    }

    if (!sheet) {
      throw new Error("Aba 'Controle de Veículos' não encontrada (verifique o nome/ID da planilha).");
    }

    const lastRow = sheet.getLastRow();
    const lastCol = Math.max(sheet.getLastColumn(), 8); // garantimos pelo menos 8 colunas
    if (lastRow < 2) return [];

    const range = sheet.getRange(2, 1, lastRow - 1, lastCol);
    const values = range.getDisplayValues ? range.getDisplayValues() : range.getValues();

    const result = values.map((rowVals, idx) => {
      return {
        row: idx + 2,
        values: rowVals
      };
    });

    return result;

  } catch (err) {
    // Log pra você inspecionar nas Execuções do Apps Script
    Logger.log('getControleVeiculos erro: ' + (err && err.message ? err.message : err));
    throw err; // devolve o erro para o cliente (google.script.run.withFailureHandler)
  }
}
function registrarRetornoVeiculo(rowNumber) {
  const sheet = SHEET_VEIC;
  if (!sheet) throw new Error("Aba 'Controle de Veículos' não encontrada.");

  const lastRow = sheet.getLastRow();
  if (rowNumber < 2 || rowNumber > lastRow) {
    throw new Error('Número de linha inválido: ' + rowNumber);
  }

  // Colunas: 1=DATA,2=FUNCIONÁRIO,3=VEÍCULO,4=HORA SAÍDA,5=PREVISÃO RETORNO,6=MOTIVO,7=STATUS,8=HORA RETORNO
  const statusCol = 7;
  const retornoCol = 8;

  const now = new Date();
  const tz = Session.getScriptTimeZone();
  const retornoTexto = Utilities.formatDate(now, tz, "dd/MM/yyyy HH:mm");

  // atualizar status e hora de retorno
  sheet.getRange(rowNumber, statusCol).setValue('Finalizado');
  sheet.getRange(rowNumber, retornoCol).setValue(retornoTexto);

  return { success: true, row: rowNumber, retorno: retornoTexto };
}

// ==================== CONFIGURAÇÕES DA APRESENTAÇÃO ====================

function getConfiguracoesApresentacao() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Configuracoes');

    if (!sheet) {
      // Criar aba se não existir
      sheet = ss.insertSheet('Configuracoes');
      sheet.getRange('A1:B1').setValues([['chave', 'valor']]);
      sheet.getRange('A2:B5').setValues([
        ['timeKanban', '10'],
        ['timeSeguranca', '1'],
        ['transitionTime', '1.5'],
        ['messagePosition', 'bottom']
      ]);
    }

    const data = sheet.getDataRange().getValues();
    const config = {};

    for (let i = 1; i < data.length; i++) {
      const chave = data[i][0];
      const valor = data[i][1];

      if (chave === 'timeKanban' || chave === 'timeSeguranca') {
        config[chave] = parseInt(valor) || 5;
      } else if (chave === 'transitionTime') {
        config[chave] = parseFloat(valor) || 1.5;
      } else {
        config[chave] = valor;
      }
    }

    return { success: true, config: config };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function salvarConfiguracoesApresentacao(config) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Configuracoes');

    if (!sheet) {
      sheet = ss.insertSheet('Configuracoes');
      sheet.getRange('A1:B1').setValues([['chave', 'valor']]);
    }

    // Limpar dados antigos (exceto cabeçalho)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 2).clear();
    }

    // Salvar novas configurações
    const configArray = Object.entries(config).map(([chave, valor]) => [chave, valor.toString()]);
    if (configArray.length > 0) {
      sheet.getRange(2, 1, configArray.length, 2).setValues(configArray);
    }

    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// Atualizar função de mensagem para incluir destaque
function salvarMensagemApresentacao(texto, cor, tamanho, destaque) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('MensagensApresentacao');

    if (!sheet) {
      sheet = ss.insertSheet('MensagensApresentacao');
      sheet.getRange('A1:E1').setValues([['id', 'texto', 'cor', 'tamanho', 'destaque']]);
    }

    const id = Utilities.getUuid();
    const lastRow = sheet.getLastRow() + 1;

    sheet.getRange(lastRow, 1, 1, 5).setValues([[id, texto, cor, tamanho, destaque || false]]);

    return { success: true, id: id };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

function getMensagensApresentacao() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('MensagensApresentacao');

    if (!sheet || sheet.getLastRow() <= 1) {
      return [];
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();

    return data.map(row => ({
      id: row[0],
      texto: row[1],
      cor: row[2],
      tamanho: row[3],
      destaque: row[4] === true || row[4] === 'true'
    })).filter(msg => msg.texto);
  } catch (error) {
    console.error('Erro ao buscar mensagens:', error);
    return [];
  }
}

// Funções para sincronizar confirmação de notificação de orçamento
function confirmarNotificacaoOrcamento(timestamp) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('notificacao_orcamento_confirmada', timestamp.toString());
    // Limpa a lista de orçamentos pendentes quando confirma
    props.deleteProperty('notificacao_orcamentos_pendentes');
    return { success: true, timestamp: timestamp };
  } catch (error) {
    Logger.log('Erro ao confirmar notificação de orçamento: ' + error.message);
    return { success: false, error: error.message };
  }
}

function verificarConfirmacaoNotificacaoOrcamento() {
  try {
    const props = PropertiesService.getScriptProperties();
    const timestampStr = props.getProperty('notificacao_orcamento_confirmada');

    if (timestampStr) {
      const timestamp = parseInt(timestampStr);
      return { confirmado: true, timestamp: timestamp };
    }

    return { confirmado: false, timestamp: null };
  } catch (error) {
    Logger.log('Erro ao verificar confirmação de notificação: ' + error.message);
    return { confirmado: false, timestamp: null, error: error.message };
  }
}

// Salva lista de orçamentos que precisam de notificação
function salvarOrcamentosPendentesNotificacao(orcamentosIds) {
  try {
    const props = PropertiesService.getScriptProperties();
    const timestamp = new Date().getTime();
    const data = {
      timestamp: timestamp,
      orcamentos: orcamentosIds
    };
    props.setProperty('notificacao_orcamentos_pendentes', JSON.stringify(data));
    return { success: true, timestamp: timestamp };
  } catch (error) {
    Logger.log('Erro ao salvar orçamentos pendentes: ' + error.message);
    return { success: false, error: error.message };
  }
}

// Verifica se há orçamentos pendentes de notificação
function verificarOrcamentosPendentesNotificacao() {
  try {
    const props = PropertiesService.getScriptProperties();
    const dataStr = props.getProperty('notificacao_orcamentos_pendentes');

    if (dataStr) {
      const data = JSON.parse(dataStr);
      // Verifica se a notificação ainda não foi confirmada
      const confirmacaoStr = props.getProperty('notificacao_orcamento_confirmada');
      const timestampConfirmacao = confirmacaoStr ? parseInt(confirmacaoStr) : 0;

      // Se a confirmação é mais recente que a notificação, não há pendências
      if (timestampConfirmacao >= data.timestamp) {
        return { pendente: false, orcamentos: [] };
      }

      return { pendente: true, timestamp: data.timestamp, orcamentos: data.orcamentos || [] };
    }

    return { pendente: false, orcamentos: [] };
  } catch (error) {
    Logger.log('Erro ao verificar orçamentos pendentes: ' + error.message);
    return { pendente: false, orcamentos: [], error: error.message };
  }
}

function deletarMensagemApresentacao(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('MensagensApresentacao');

    // Se a planilha não existe, retorna erro
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: false, error: "Nenhuma mensagem encontrada" };
    }

    // Busca a mensagem pelo ID na coluna A (coluna 1)
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    let linhaEncontrada = -1;

    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === id) {
        linhaEncontrada = i + 2; // +2 porque começa na linha 2 (linha 1 é cabeçalho)
        break;
      }
    }

    // Se não encontrou a mensagem, retorna erro
    if (linhaEncontrada === -1) {
      Logger.log("Mensagem não encontrada com ID: " + id);
      return { success: false, error: "Mensagem não encontrada com ID: " + id };
    }

    // Deleta a linha encontrada
    sheet.deleteRow(linhaEncontrada);

    Logger.log("Mensagem deletada com sucesso. ID: " + id);
    return { success: true };
  } catch (e) {
    Logger.log("Erro ao deletar mensagem: " + e.message);
    return { success: false, error: e.message };
  }
}
// Função para listar TODAS as mensagens (incluindo inativas) - útil para debug
function getTodasMensagensApresentacao() {
  try {
    const props = PropertiesService.getScriptProperties();
    const raw = props.getProperty('APRESENTACAO_MENSAGENS');
    if (!raw) return [];

    return JSON.parse(raw);
  } catch (e) {
    Logger.log("Erro ao carregar todas as mensagens: " + e.message);
    return [];
  }
}

// Função para limpar TODAS as mensagens (use com cuidado!)
function limparTodasMensagensApresentacao() {
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('APRESENTACAO_MENSAGENS');
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

