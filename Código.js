/************* Code.gs *************/
const ss = SpreadsheetApp.openById("1wMIbd8r2HeniFLTYaG8Yhhl8CWmaHaW5oXBVnxj0xos");
const SHEET_CALC = ss.getSheetByName("Tabelas para cálculos");
const SHEET_VEIC = ss.getSheetByName('Controle de Veículos');
const SHEET_MANU_NAME = ss.getSheetByName("Registro de Manutenções");
let SHEET_PED = ss.getSheetByName("Pedidos");
const PEDIDOS_HEADERS = ["PROJETO", "LINHA_PROJETO", "CLIENTE", "NF", "DATA_COMPETENCIA", "VALOR_TOTAL", "CONDICOES_PAGAMENTO", "DATA_ENTREGA", "DATA_VENCIMENTO", "VALOR_PAGO", "STATUS_PAGAMENTO", "PARCELAS_E_PGTOS", "NUMERO_SEQUENCIAL", "OBS"];
/** Campos que pertencem apenas à aba Pedidos; na Projetos ficam só em JSON_DADOS até conversão. Não criar nem gravar essas colunas na aba Projetos. */
const CAMPOS_APENAS_PEDIDOS = { "CONDICOES_PAGAMENTO": 1, "NUMERO_SEQUENCIAL": 1, "DATA_ENTREGA": 1, "DATA_VENCIMENTO": 1, "DATA_COMPETENCIA": 1, "HISTORICO_PAGAMENTOS": 1, "NOTA_FISCAL": 1, "VALOR_PAGO": 1, "STATUS_PAGAMENTO": 1 };
const SHEET_MAT = ss.getSheetByName("Controle de Materiais");
const SHEET_AVAL = ss.getSheetByName("Avaliações");
const SHEET_LOGS = ss.getSheetByName("Logs");
const SHEET_CLIENTES = ss.getSheetByName("Cadastro de Clientes");
const SHEET_FORNECEDORES = ss.getSheetByName("Cadastro de Fornecedores");
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
    // A=Código do Produto, B=Descrição, C=Família, D=EAN, E=NCM, F=Preço, G=Unidade, H=Características, I=Projeto, J=Cliente, K=MP, L=CL, M=D, N=S, O=Pin
    const produtos = [];
    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      if (row[0]) { // se tem código (coluna A)
        var processos = [];
        if (row[10] && String(row[10]).trim().toUpperCase() === "X") processos.push("MP");
        if (row[11] && String(row[11]).trim().toUpperCase() === "X") processos.push("CL");
        if (row[12] && String(row[12]).trim().toUpperCase() === "X") processos.push("D");
        if (row[13] && String(row[13]).trim().toUpperCase() === "X") processos.push("S");
        if (row[14] && String(row[14]).trim().toUpperCase() === "X") processos.push("Pin");
        if (row[15] && String(row[15]).trim().toUpperCase() === "X") processos.push("CAD");
        if (row[16] && String(row[16]).trim().toUpperCase() === "X") processos.push("ACB");
        produtos.push({
          codigo: row[0],                    // Coluna A - Código do Produto
          descricao: row[1] || "",           // Coluna B - Descrição do Produto
          codigoFamilia: row[2] || "",       // Coluna C - Código da Família
          codigoEAN: row[3] || "",           // Coluna D - Código EAN (GTIN)
          ncm: row[4] || "",                 // Coluna E - Código NCM
          preco: parseFloat(row[5]) || 0,    // Coluna F - Preço Unitário de Venda
          unidade: row[6] || "UN",            // Coluna G - Unidade
          processos: processos               // Colunas K-O (MP, CL, D, S, Pin)
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
 * Atribui códigos PRD às peças das chapas que não possuem código
 * @param {Array} chapas - Array de chapas com peças
 */
function atribuirCodigosPRDEmChapas(chapas) {
  if (!chapas || !Array.isArray(chapas)) return;
  var numeroPRD = parseInt(getProximoCodigoPRD().substring(3), 10) || 0;
  chapas.forEach(function (chapa) {
    if (chapa.pecas && Array.isArray(chapa.pecas)) {
      chapa.pecas.forEach(function (peca) {
        var codigo = (peca.codigo && String(peca.codigo).trim()) || "";
        if (!codigo || String(codigo).toUpperCase().indexOf("PRD") !== 0) {
          peca.codigo = "PRD" + String(numeroPRD).padStart(5, "0");
          numeroPRD++;
        }
      });
    }
  });
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

    var processosRel = produto.processos && Array.isArray(produto.processos) ? produto.processos : [];
    function temProc(sigla) { return processosRel.indexOf(sigla) >= 0; }
    // Estrutura da planilha:
    // A=Código do Produto, B=Descrição do Produto, C=Código da Família, D=Código EAN (GTIN), 
    // E=Código NCM, F=Preço Unitário de Venda, G=Unidade, H=Características, I=Projeto, J=Cliente
    // K=MP, L=CL, M=D, N=S, O=Pin, P=CAD, Q=ACB (processos: X ou vazio)
    const novaLinha = [
      produto.codigo || "",           // A - Código do Produto
      produto.descricao || "",        // B - Descrição do Produto
      "",                             // C - Código da Família (vazio)
      "",                             // D - Código EAN (vazio)
      produto.ncm || "",              // E - Código NCM
      produto.preco || 0,             // F - Preço Unitário de Venda
      produto.unidade || "UN",        // G - Unidade
      produto.caracteristicas || "",  // H - Características
      produto.projeto || "",          // I - Projeto
      produto.cliente || "",          // J - Cliente
      temProc("MP") ? "X" : "",       // K - MP
      temProc("CL") ? "X" : "",       // L - CL
      temProc("D") ? "X" : "",        // M - D
      temProc("S") ? "X" : "",        // N - S
      temProc("Pin") ? "X" : "",      // O - Pin
      temProc("CAD") ? "X" : "",      // P - CAD
      temProc("ACB") ? "X" : ""       // Q - ACB
    ];

    var numCols = SHEET_PRODUTOS.getLastColumn();
    if (numCols < 11 && novaLinha.length > 10) {
      SHEET_PRODUTOS.getRange(1, 11).setValue("MP");
      SHEET_PRODUTOS.getRange(1, 12).setValue("CL");
      SHEET_PRODUTOS.getRange(1, 13).setValue("D");
      SHEET_PRODUTOS.getRange(1, 14).setValue("S");
      SHEET_PRODUTOS.getRange(1, 15).setValue("Pin");
      SHEET_PRODUTOS.getRange(1, 16).setValue("CAD");
      SHEET_PRODUTOS.getRange(1, 17).setValue("ACB");
    }
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
    // Estrutura: A=Código, B=Descrição, C=Código Família, D=EAN, E=NCM, F=Preço, G=Unidade, H=Características, I=Projeto, J=Cliente, K=MP, L=CL, M=D, N=S, O=Pin
    SHEET_PRODUTOS.getRange(linhaEncontrada, 2).setValue(dadosNovos.descricao || ""); // B - Descrição
    SHEET_PRODUTOS.getRange(linhaEncontrada, 5).setValue(dadosNovos.ncm || ""); // E - NCM
    SHEET_PRODUTOS.getRange(linhaEncontrada, 6).setValue(dadosNovos.preco || 0); // F - Preço
    SHEET_PRODUTOS.getRange(linhaEncontrada, 7).setValue(dadosNovos.unidade || "UN"); // G - Unidade

    var processosArr = dadosNovos.processos && Array.isArray(dadosNovos.processos) ? dadosNovos.processos : [];
    function temP(sigla) { return processosArr.indexOf(sigla) >= 0; }
    var numCols = SHEET_PRODUTOS.getLastColumn();
    if (numCols < 11) {
      SHEET_PRODUTOS.getRange(1, 11).setValue("MP");
      SHEET_PRODUTOS.getRange(1, 12).setValue("CL");
      SHEET_PRODUTOS.getRange(1, 13).setValue("D");
      SHEET_PRODUTOS.getRange(1, 14).setValue("S");
      SHEET_PRODUTOS.getRange(1, 15).setValue("Pin");
      SHEET_PRODUTOS.getRange(1, 16).setValue("CAD");
      SHEET_PRODUTOS.getRange(1, 17).setValue("ACB");
    }
    SHEET_PRODUTOS.getRange(linhaEncontrada, 11).setValue(temP("MP") ? "X" : "");
    SHEET_PRODUTOS.getRange(linhaEncontrada, 12).setValue(temP("CL") ? "X" : "");
    SHEET_PRODUTOS.getRange(linhaEncontrada, 13).setValue(temP("D") ? "X" : "");
    SHEET_PRODUTOS.getRange(linhaEncontrada, 14).setValue(temP("S") ? "X" : "");
    SHEET_PRODUTOS.getRange(linhaEncontrada, 15).setValue(temP("Pin") ? "X" : "");
    SHEET_PRODUTOS.getRange(linhaEncontrada, 16).setValue(temP("CAD") ? "X" : "");
    SHEET_PRODUTOS.getRange(linhaEncontrada, 17).setValue(temP("ACB") ? "X" : "");

    return {
      success: true,
      mensagem: "PRD atualizado no catálogo."
    };
  } catch (err) {
    Logger.log("Erro ao atualizar PRD no catálogo: " + err.message);
    throw new Error("Erro ao atualizar PRD: " + err.message);
  }
}

/**
 * Insere produtos com código PRD das chapas na "Relação de produtos"
 * @param {Array} chapas - Array com dados das chapas e peças
 * @param {string} codigoProjeto - Código do projeto (para coluna Projeto)
 * @param {string} nomeCliente - Nome do cliente (para coluna Cliente)
 */
function inserirProdutosDasChapas(chapas, codigoProjeto, nomeCliente) {
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
              caracteristicas: `${chapa.material} - ${peca.comprimento}x${peca.largura} - ${chapa.espessura}mm`,
              projeto: codigoProjeto || "",
              cliente: nomeCliente || ""
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

    // Processos do Pedido (descontos e custos extras: nova estrutura ou antiga)
    if (dados.processosPedido && Array.isArray(dados.processosPedido)) {
      var subtotalAntesProcessos = totalGeral;
      dados.processosPedido.forEach(proc => {
        var preco = 0;
        if (proc.tipo === "desconto" || proc.tipo === "custo") {
          var tipoValor = proc.tipoValor === "percentual" || proc.tipoValor === "fixo" ? proc.tipoValor : "fixo";
          if (proc.tipo === "desconto") {
            if (tipoValor === "percentual") {
              var pct = parseFloat(proc.percentual) || 0;
              preco = -(subtotalAntesProcessos * pct / 100);
            } else {
              preco = -(parseFloat(proc.valorFixo) || 0);
            }
          } else {
            preco = parseFloat(proc.valorFixo) || 0;
          }
          detalhamento.push({
            tipo: 'processo',
            descricao: proc.descricao || (proc.tipo === "desconto" ? "Desconto" : "Custo extra"),
            precoTotal: preco
          });
        } else {
          var vh = parseFloat(proc.valorHora) || 0, h = parseFloat(proc.horas) || 0;
          var vm = parseFloat(proc.valorMat) || 0, qm = parseFloat(proc.qtdMat) || 0, vf = parseFloat(proc.valorFixo) || 0;
          preco = vh * h + vm * qm + vf;
          totalGeral += preco;
          detalhamento.push({
            tipo: 'processo',
            descricao: proc.descricao || 'Processo adicional',
            precoTotal: preco
          });
          return;
        }
        totalGeral += preco;
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
      email: dados[i][4],
      nomeAbreviado: dados[i][5] || ""
    });
  }
  return clientes;
}

function salvarClienteSeNovo(cliente) {
  const dados = SHEET_CLIENTES.getDataRange().getValues();
  for (let i = 1; i < dados.length; i++) {
    if (dados[i][0].toString().trim().toLowerCase() === cliente.nome.trim().toLowerCase()) {
      // Atualiza nomeAbreviado se ainda não estiver preenchido
      if (cliente.nomeAbreviado && !dados[i][5]) {
        SHEET_CLIENTES.getRange(i + 1, 6).setValue(cliente.nomeAbreviado);
      }
      return;
    }
  }
  SHEET_CLIENTES.appendRow([cliente.nome, cliente.cpf, cliente.endereco, cliente.telefone, cliente.email, cliente.nomeAbreviado || ""]);
}

// ========================= FORNECEDORES (mesma estrutura de clientes: nome, cpf, endereco, telefone, email) =========================
/** Retorna a aba "Cadastro de Fornecedores"; cria com cabeçalhos se não existir. */
function getSheetFornecedores() {
  var sheet = ss.getSheetByName("Cadastro de Fornecedores");
  if (!sheet) {
    sheet = ss.insertSheet("Cadastro de Fornecedores");
    sheet.getRange(1, 1, 1, 5).setValues([["Nome", "CPF/CNPJ", "Endereço", "Telefone", "Email"]]);
    sheet.getRange(1, 1, 1, 5).setFontWeight("bold");
  }
  return sheet;
}

function getTodosFornecedores() {
  var sheet = SHEET_FORNECEDORES || ss.getSheetByName("Cadastro de Fornecedores");
  if (!sheet) return [];
  var dados = sheet.getDataRange().getValues();
  var lista = [];
  for (var i = 1; i < dados.length; i++) {
    lista.push({
      nome: dados[i][0] || "",
      cpf: dados[i][1] || "",
      endereco: dados[i][2] || "",
      telefone: dados[i][3] || "",
      email: dados[i][4] || ""
    });
  }
  return lista;
}

function salvarFornecedorSeNovo(fornecedor) {
  var sheet = getSheetFornecedores();
  var dados = sheet.getDataRange().getValues();
  var nome = (fornecedor.nome || "").trim();
  if (!nome) return;
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0] || "").trim().toLowerCase() === nome.toLowerCase()) return;
  }
  sheet.appendRow([
    nome,
    fornecedor.cpf || "",
    fornecedor.endereco || "",
    fornecedor.telefone || "",
    fornecedor.email || ""
  ]);
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
    .replace(/\bCOT\b/gi, "")       // Não usar prefixos COT/PED na descrição/nome (reservados)
    .replace(/\bPED\b/gi, "")
    .replace(/\s+/g, " ")           // Normaliza espaços múltiplos
    .trim();
}

/**
 * Gera o nome formatado da pasta
 * @param {string} codigoProjeto - Código do projeto (ex: "260202aBR")
 * @param {string} nomeCliente - Nome do cliente (completo ou abreviado)
 * @param {string} descricao - Descrição do projeto
 * @param {boolean} isPedido - Se true, usa prefixo PED; se false, usa COT
 * @param {string} [nomeAbreviado] - Nome abreviado do cliente (se fornecido, prevalece sobre nomeCliente)
 * @returns {string} - Nome formatado (ex: "260202aBR COT ABREV - DESCRICAO")
 */
function gerarNomePasta(codigoProjeto, nomeCliente, descricao, isPedido, nomeAbreviado) {
  const prefixo = isPedido ? "PED" : "COT";
  const nomeParaPasta = (nomeAbreviado && String(nomeAbreviado).trim()) ? nomeAbreviado : nomeCliente;
  const clienteLimpo = limparNomePasta(nomeParaPasta || "");
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
 * Renomeia a pasta do projeto no Drive quando a descrição ou o cliente é alterado (ex.: pela página de projetos).
 * @param {string} codigoProjeto - Código do projeto
 * @param {string} data - Data no formato YYMMDD (primeiros 6 caracteres do código)
 * @param {string} cliente - Nome do cliente
 * @param {string} descricao - Descrição do projeto
 * @param {boolean} isPedido - Se true, usa prefixo PED; se false, COT
 * @returns {boolean} - true se a pasta foi renomeada
 */
function renomearPastaProjeto(codigoProjeto, data, cliente, descricao, isPedido) {
  try {
    var pastaInfo = detectarPastaExistente(codigoProjeto, data);
    if (!pastaInfo) return false;
    var nome = gerarNomePasta(codigoProjeto, cliente || "", descricao || "", !!isPedido);
    pastaInfo.pasta.setName(nome);
    Logger.log("📁 Pasta renomeada: " + nome);
    return true;
  } catch (e) {
    Logger.log("Erro ao renomear pasta: " + e.message);
    return false;
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
 * @param {string} [nomeAbreviado] - Nome abreviado do cliente (opcional)
 * @returns {Folder} - Pasta do projeto
 */
function criarOuUsarPastaProjeto(codigoProjeto, nomeCliente, descricao, data, isPedido, nomeAbreviado) {
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
    const nomeDesejado = gerarNomePasta(codigoProjeto, nomeCliente, descricao, isPedido, nomeAbreviado);
    
    // Se mudou de COT para PED, atualiza
    if (isPedido && pastaInfo.tipo === "COT") {
      pastaInfo.pasta.setName(nomeDesejado);
      Logger.log("🔄 Pasta convertida de COT para PED: " + nomeDesejado);
    } 
    // Se o nome mudou (cliente ou descrição), atualiza mantendo o tipo atual
    else if (!isPedido || pastaInfo.tipo === "PED") {
      const tipoAtual = isPedido ? "PED" : pastaInfo.tipo;
      const nomeAtualizado = gerarNomePasta(codigoProjeto, nomeCliente, descricao, tipoAtual === "PED", nomeAbreviado);
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
  
  const nomePasta = gerarNomePasta(codigoProjeto, nomeCliente, descricao, isPedido, nomeAbreviado);
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
                // Formato antigo: "Proposta_<codigo>.pdf" (sem sufixo de versão)
                if (nomeArquivo === prefixo + ".pdf") {
                  versoesEncontradas.push(1); // Sem sufixo = v1
                } else {
                  // Formato antigo: Proposta_260202cBR_v2.pdf -> "v2"
                  const matchOld = nomeArquivo.match(new RegExp(prefixo + "_v(\\d+)\\.pdf"));
                  if (matchOld && matchOld[1]) {
                    versoesEncontradas.push(parseInt(matchOld[1], 10));
                  } else {
                    // Formato novo: Proposta_260202cBR_1705.pdf (seq, sem versão) → v1
                    const matchNewBase = nomeArquivo.match(new RegExp(prefixo + "_\\d+\\.pdf"));
                    if (matchNewBase) {
                      versoesEncontradas.push(1);
                    } else {
                      // Formato novo: Proposta_260202cBR_1705_v2.pdf → v2
                      const matchNew = nomeArquivo.match(new RegExp(prefixo + "_\\d+_v(\\d+)\\.pdf"));
                      if (matchNew && matchNew[1]) {
                        versoesEncontradas.push(parseInt(matchNew[1], 10));
                      }
                    }
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
 * Detecta a próxima versão para Ordem de Produção - busca "Ordem_Producao_" na pasta.
 * Só adiciona v2, v3... se já existir arquivo com o mesmo nome base.
 * @param {string} codigoProjeto - Código do projeto (ex: "260202cBR")
 * @param {string} data - Data no formato YYMMDD
 * @returns {string} - Próxima versão (ex: "", "v2", "v3") - "" = primeira, sem sufixo
 */
function detectarProximaVersaoOrdemProducao(codigoProjeto, data) {
  try {
    if (!codigoProjeto || !data) return "";
    const pastaInfo = detectarPastaExistente(codigoProjeto, data);
    if (!pastaInfo) return "";
    const pastaProjeto = pastaInfo.pasta;
    try {
      const workFolders = pastaProjeto.getFoldersByName("02_WORK");
      if (workFolders.hasNext()) {
        const workFolder = workFolders.next();
        const comFolders = workFolder.getFoldersByName("COM");
        if (comFolders.hasNext()) {
          const comFolder = comFolders.next();
          const arquivos = comFolder.getFiles();
          const prefixo = "Ordem_Producao_" + codigoProjeto;
          const versoesEncontradas = [];
          while (arquivos.hasNext()) {
            const arquivo = arquivos.next();
            const nomeArquivo = arquivo.getName();
            if (nomeArquivo.startsWith(prefixo) && nomeArquivo.endsWith(".pdf")) {
              if (nomeArquivo === prefixo + ".pdf") {
                versoesEncontradas.push(1);
              } else {
                const match = nomeArquivo.match(new RegExp(prefixo + "_v(\\d+)\\.pdf"));
                if (match && match[1]) {
                  versoesEncontradas.push(parseInt(match[1], 10));
                }
              }
            }
          }
          if (versoesEncontradas.length === 0) return "";
          const maiorVersao = Math.max(...versoesEncontradas);
          return "v" + (maiorVersao + 1);
        }
      }
    } catch (e) {
      Logger.log("Erro ao buscar versões Ordem de Produção: " + e.message);
    }
    return "";
  } catch (e) {
    Logger.log("Erro detectarProximaVersaoOrdemProducao: " + e.message);
    return "";
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
// Modificado para aceitar nomeCliente, isPedido e nomeAbreviado
function criarPastaOrcamento(codigoProjeto, descricao, data, nomeCliente, isPedido, nomeAbreviado) {
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
    isPedido || false,
    nomeAbreviado || ""
  );

  return {
    pastaId: pastaProjeto.getId(),
    pastaNome: pastaProjeto.getName(),
    pastaUrl: pastaProjeto.getUrl()
  };
}

// Cria a pasta 01_IN dentro da pasta do projeto (ou cria a pasta do projeto se não existir).
// Retorna a URL da pasta 01_IN.
function criarPasta01IN(codigoProjeto, descricao, data, nomeCliente, nomeAbreviado) {
  if (!codigoProjeto || !data) {
    throw new Error("Dados do projeto incompletos (código ou data ausentes).");
  }
  if (!descricao || descricao.trim() === "") {
    throw new Error("Descrição do projeto é obrigatória para criar a pasta.");
  }
  const pastaProjeto = criarOuUsarPastaProjeto(
    codigoProjeto,
    nomeCliente || "",
    descricao,
    data,
    false,
    nomeAbreviado || ""
  );
  const inFolder = getOrCreateSubFolder(pastaProjeto, "01_IN");
  return {
    pastaId: pastaProjeto.getId(),
    pastaNome: pastaProjeto.getName(),
    pastaUrl: pastaProjeto.getUrl(),
    inFolderId: inFolder.getId(),
    inFolderUrl: inFolder.getUrl()
  };
}

// Busca a pasta 01_IN do projeto e retorna sua URL. Lança erro se não existir.
function abrirPasta01IN(codigoProjeto, data) {
  if (!codigoProjeto || !data) {
    throw new Error("Dados do projeto incompletos (código ou data ausentes).");
  }
  const pastaInfo = detectarPastaExistente(codigoProjeto, data);
  if (!pastaInfo) {
    throw new Error("Pasta do projeto não encontrada. Crie a pasta 01_IN primeiro.");
  }
  const inFolders = pastaInfo.pasta.getFoldersByName("01_IN");
  if (!inFolders.hasNext()) {
    throw new Error("Pasta 01_IN não encontrada. Clique em 'Criar pasta 01_IN' primeiro.");
  }
  const inFolder = inFolders.next();
  return { inFolderUrl: inFolder.getUrl(), inFolderId: inFolder.getId() };
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

// Salva UM único arquivo na pasta 01_IN. Usado quando o cliente envia múltiplos arquivos.
// O cliente NÃO pode enviar File/Blob pelo google.script.run (erro "illegal value in property"),
// então envia fileData: { base64: string, name: string, mimeType: string }.
function salvarArquivoClienteUnico(codigoProjeto, nomePasta, dataProjeto, fileData) {
  if (!codigoProjeto || !dataProjeto) {
    throw new Error("Dados do projeto incompletos para salvar arquivos (código ou data ausentes).");
  }
  if (!fileData || !fileData.base64) {
    throw new Error("Nenhum arquivo recebido (envie em base64).");
  }

  var blob;
  if (typeof fileData.base64 === "string") {
    var base64 = fileData.base64.replace(/^data:[^;]+;base64,/, "");
    blob = Utilities.newBlob(Utilities.base64Decode(base64), fileData.mimeType || "application/octet-stream", fileData.name || "arquivo");
  } else {
    throw new Error("Formato de arquivo inválido.");
  }

  const pastaProjeto = criarOuUsarPasta(codigoProjeto, nomePasta || "", dataProjeto);
  const inFolder = getOrCreateSubFolder(pastaProjeto, "01_IN");

  const file = inFolder.createFile(blob);
  var nomeArquivo = fileData.name || file.getName();
  if (nomeArquivo && file.getName() !== nomeArquivo) {
    file.setName(nomeArquivo);
  }

  return {
    ok: true,
    quantidade: 1,
    pastaNome: pastaProjeto.getName(),
    inFolderNome: inFolder.getName(),
    arquivo: { id: file.getId(), nome: file.getName(), url: file.getUrl() }
  };
}

// ========================= GERAR PDF (VERSÃO AJUSTADA) =========================
// sobrescreverVersao: se true, sobrescreve o PDF existente em vez de criar _v2, _v3
function gerarPdfOrcamento(
  chapas, cliente, observacoes, codigoProjeto, nomePasta, dataProjeto, versao, somaProcessosPedido, descricaoProcessosPedido, produtosCadastrados, dadosFormularioCompleto, infoPagamento, isPedido, sobrescreverVersao, apenasPreview
) {
  sobrescreverVersao = !!sobrescreverVersao;
  apenasPreview = !!apenasPreview;
  try {

    if (!apenasPreview) incrementarContador("totalPropostas");

    // Número sequencial: ao sobrescrever mantém o existente; novas versões recebem número novo (em preview usa dados ou "—")
    let numeroSequencial;
    if (apenasPreview) {
      numeroSequencial = (dadosFormularioCompleto && dadosFormularioCompleto.numeroSequencial != null && dadosFormularioCompleto.numeroSequencial !== "") ? dadosFormularioCompleto.numeroSequencial : "—";
    } else if (sobrescreverVersao) {
      // Prioridade: valor que está NA PLANILHA (o que o usuário editou manualmente), depois formulário, depois incrementar.
      const sheetProj = ss.getSheetByName("Projetos");
      const linhaProj = sheetProj ? findRowByColumnValue(sheetProj, "PROJETO", codigoProjeto) : null;

      // 1) Ler da aba Pedidos (coluna Nº / NUMERO_SEQUENCIAL) — onde o usuário costuma editar manualmente
      var sheetPed = ss.getSheetByName("Pedidos");
      if (sheetPed && sheetPed.getLastRow() >= 2) {
        var linhaPed = findRowByColumnValue(sheetPed, "PROJETO", codigoProjeto);
        if (linhaPed) {
          var headersPed = sheetPed.getRange(1, 1, 1, sheetPed.getLastColumn()).getValues()[0];
          var rowPed = sheetPed.getRange(linhaPed, 1, linhaPed, sheetPed.getLastColumn()).getValues()[0];
          var aliasesNum = ["NUMERO_SEQUENCIAL", "NUMERO SEQUENCIAL", "NÚMERO SEQUENCIAL", "Nº", "N"];
          for (var a = 0; a < aliasesNum.length && (numeroSequencial == null || numeroSequencial === undefined); a++) {
            for (var c = 0; c < headersPed.length; c++) {
              if (String(headersPed[c] || "").trim() === String(aliasesNum[a] || "").trim()) {
                var valPed = rowPed[c];
                if (valPed != null && String(valPed).trim() !== "") {
                  numeroSequencial = valPed;
                  break;
                }
              }
            }
          }
        }
      }

      // 2) Se não achou em Pedidos, ler do JSON_DADOS na aba Projetos
      if ((numeroSequencial == null || numeroSequencial === undefined) && linhaProj && sheetProj) {
        var numCols = sheetProj.getLastColumn();
        var headers = sheetProj.getRange(1, 1, 1, numCols).getValues()[0];
        var jsonIdx = _findHeaderIndex(headers, "JSON_DADOS");
        if (jsonIdx >= 0) {
          var jsonCell = sheetProj.getRange(linhaProj, jsonIdx + 1).getValue();
          try {
            var parsed = jsonCell ? JSON.parse(String(jsonCell).trim()) : null;
            if (parsed && parsed.numeroSequencial != null) numeroSequencial = parsed.numeroSequencial;
            else if (parsed && parsed.dados && parsed.dados.numeroSequencial != null) numeroSequencial = parsed.dados.numeroSequencial;
          } catch (e) {}
        }
      }

      // 3) Se ainda não tem, usar o que veio do formulário
      if ((numeroSequencial == null || numeroSequencial === undefined) && dadosFormularioCompleto && dadosFormularioCompleto.numeroSequencial != null && String(dadosFormularioCompleto.numeroSequencial).trim() !== "")
        numeroSequencial = dadosFormularioCompleto.numeroSequencial;

      // 4) Só então incrementar
      if (numeroSequencial == null || numeroSequencial === undefined || String(numeroSequencial).trim() === "")
        numeroSequencial = obterEIncrementarNumeroOrcamento();

      if (dadosFormularioCompleto) dadosFormularioCompleto.numeroSequencial = numeroSequencial;
    } else {
      numeroSequencial = obterEIncrementarNumeroOrcamento();
      if (dadosFormularioCompleto) {
        dadosFormularioCompleto.numeroSequencial = numeroSequencial;
      }
    }

    // Atribui códigos PRD apenas a produtos cadastrados (orçamento não usa mais chapas/peças)
    var proximoPRD = getProximoCodigoPRD();
    var numeroPRD = parseInt(proximoPRD.substring(3), 10) || 0;
    if (produtosCadastrados && Array.isArray(produtosCadastrados)) {
      produtosCadastrados.forEach(function (prod) {
        var codigo = (prod.codigo && String(prod.codigo).trim()) || "";
        if (!codigo || String(codigo).toUpperCase().indexOf("PRD") !== 0) {
          prod.codigo = "PRD" + String(numeroPRD).padStart(5, "0");
          numeroPRD++;
        }
      });
    }

    const resultados = [];

    // Adiciona produtos cadastrados aos resultados (precoTotal usado para valor na planilha)
    if (produtosCadastrados && Array.isArray(produtosCadastrados)) {
      produtosCadastrados.forEach(prod => {
        var precoUnit = Number(prod.precoUnitario) || 0;
        var qtd = Number(prod.quantidade) || 0;
        var precoTotal = Number(prod.precoTotal) || 0;
        if (!precoTotal && (precoUnit || qtd)) {
          precoTotal = precoUnit * qtd;
        }
        resultados.push({
          codigo: prod.codigo || "",
          descricao: prod.descricao || "",
          quantidade: qtd,
          precoUnitario: precoUnit,
          precoTotal: precoTotal,
          processos: prod.processos && Array.isArray(prod.processos) ? prod.processos : [],
          descricoesProcessos: prod.descricoesProcessos && typeof prod.descricoesProcessos === 'object' ? prod.descricoesProcessos : {}
        });
      });
    }

    // Código base (sem _v2) para pasta - todas as versões vão na mesma pasta
    const matchBase = String(codigoProjeto || "").match(/^(.+?)(_v\d+)$/);
    const codigoBase = matchBase ? matchBase[1] : (codigoProjeto || "");

    // Usa nova estrutura de pastas (quando isPedido=true, o prefixo da pasta deve ser PED, não COT) — em preview não cria pasta
    var pasta, workFolder, comSubFolder;
    if (!apenasPreview) {
      const nomeCliente = cliente.nome || "";
      const nomeAbreviado = (dadosFormularioCompleto && dadosFormularioCompleto.cliente && dadosFormularioCompleto.cliente.nomeAbreviado) || "";
      const descricaoPasta = observacoes.descricao || nomePasta || codigoBase;
      pasta = criarOuUsarPastaProjeto(codigoBase, nomeCliente, descricaoPasta, dataProjeto, isPedido || false, nomeAbreviado);
      if (isPedido && pasta && pasta.getName().indexOf(" PED ") === -1) {
        const nomePED = gerarNomePasta(codigoBase, nomeCliente, descricaoPasta, true, nomeAbreviado);
        pasta.setName(nomePED);
        Logger.log("📁 Pasta garantida com prefixo PED: " + nomePED);
      }
      workFolder = getOrCreateSubFolder(pasta, "02_WORK");
      comSubFolder = getOrCreateSubFolder(workFolder, "COM");
    }

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

    // Lógica de versão: nova versão (_v2, _v3) ou sobrescrever atual
    let versaoFinal = "";
    let codigoParaPlanilha = codigoBase;

    if (sobrescreverVersao) {
      // Sobrescrever: usar a versão informada (ex: _v2) ou v1 (sem sufixo)
      versaoFinal = (versao && String(versao).trim()) ? (versao.startsWith("_") ? versao : "_" + versao) : "";
      codigoParaPlanilha = codigoBase + versaoFinal;
    } else {
      // Nova versão: detectar próxima disponível
      const proximaVersao = detectarProximaVersao(codigoBase, dataProjeto);
      versaoFinal = proximaVersao ? "_" + proximaVersao : "";
      codigoParaPlanilha = codigoBase + versaoFinal;
    }

    const numeroProposta = codigoParaPlanilha;

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

    // Calcula parcelas baseado no texto de pagamento.
    // Suporta: "30/60/90" (dias), "Á Vista / 30 / 45", ou "70% no pedido, 30% na entrega" (percentuais por condição).
    function calcularParcelas(textoPagamento, valorTotal) {
      if (!textoPagamento || textoPagamento.trim() === "") return null;
      const texto = textoPagamento.trim().replace(/\s+/g, " ");
      const textoUpper = texto.toUpperCase();

      // --- Detectar percentuais por condição (ex: 70% no pedido, 30% na entrega) ---
      const temPedidoOuEntrega = /pedido|entrega/i.test(texto);
      if (temPedidoOuEntrega && textoUpper.indexOf("%") >= 0) {
        const partes = texto.split(/(\d+\s*%)/);
        const percentuais = [];
        for (let i = 1; i < partes.length; i += 2) {
          const pctStr = partes[i] || "";
          const textoApos = (partes[i + 1] || "").toLowerCase();
          const pct = parseInt(pctStr.replace(/\D/g, ""), 10);
          if (isNaN(pct)) continue;
          const condicao = textoApos.indexOf("entrega") >= 0 ? "Na entrega" : "No pedido";
          percentuais.push({ pct: pct, condicao: condicao });
        }
        if (percentuais.length >= 1) {
          return percentuais.map((item, idx) => ({
            numero: idx + 1,
            dias: null,
            condicao: item.condicao,
            valor: valorTotal * (item.pct / 100)
          }));
        }
      }

      // --- Fluxo original: dias (ex: 30/60/90) ---
      if (textoUpper.includes("VISTA") && !textoUpper.includes("/")) return null;
      if (textoUpper === "30 DIAS") return null;
      let dias = [];
      if (textoUpper.includes("/")) {
        const partes = textoUpper.split(/\s*\/\s*/);
        for (let i = 0; i < partes.length; i++) {
          const p = (partes[i] || "").trim();
          if (/vista/i.test(p)) dias.push(0);
          else {
            const num = parseInt(p.replace(/\D/g, ""), 10);
            if (!isNaN(num)) dias.push(num);
          }
        }
      }
      if (dias.length === 0) {
        const diasMatch = textoUpper.match(/\d+/g);
        if (!diasMatch || diasMatch.length === 0) return null;
        dias = diasMatch.map(d => parseInt(d, 10));
      }
      const numParcelas = dias.length;
      const valorParcela = valorTotal / numParcelas;
      return dias.map((dia, idx) => ({
        numero: idx + 1,
        dias: dia,
        condicao: null,
        valor: valorParcela
      }));
    }

    const itensHtml = resultados.map(function (p) {
      return ''
        + '<tr>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; font-size:7pt;">' + esc(p.codigo || "") + '</td>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; font-size:7pt;">' + esc(p.descricao || "") + '</td>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">' + esc(p.quantidade || 0) + '</td>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">' + formatBR(p.precoUnitario || 0) + '</td>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">' + formatBR(p.precoTotal || 0) + '</td>'
        + '</tr>';
    }).join('');

    // Processos do pedido: descontos e custos extras (nova estrutura com lista detalhada)
    var processosPedidoRow = "";
    var processosPedidoArray = (dadosFormularioCompleto && dadosFormularioCompleto.processosPedido && Array.isArray(dadosFormularioCompleto.processosPedido)) ? dadosFormularioCompleto.processosPedido : [];
    var subtotalParaPercentual = totalPecas;
    if (processosPedidoArray.length > 0 && processosPedidoArray[0].tipo !== undefined) {
      processosPedidoArray.forEach(function (proc) {
        var tipo = proc.tipo === "desconto" || proc.tipo === "custo" ? proc.tipo : "custo";
        var tipoValor = proc.tipoValor === "percentual" || proc.tipoValor === "fixo" ? proc.tipoValor : "fixo";
        var valor = 0;
        if (tipo === "desconto") {
          if (tipoValor === "percentual") {
            var pct = parseFloat(proc.percentual) || 0;
            valor = -(subtotalParaPercentual * pct / 100);
          } else {
            valor = -(parseFloat(proc.valorFixo) || 0);
          }
        } else {
          valor = parseFloat(proc.valorFixo) || 0;
        }
        var descricaoLinha = (proc.descricao || "").trim();
        if (!descricaoLinha) descricaoLinha = tipo === "desconto" ? "Desconto" : "Custo extra";
        processosPedidoRow += '<tr>'
          + '<td colspan="4" bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">' + esc(descricaoLinha) + '</td>'
          + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">' + formatBR(valor) + '</td>'
          + '</tr>';
      });
    } else if (somaProcessosPedido && Number(somaProcessosPedido) !== 0) {
      processosPedidoRow = '<tr>'
        + '<td colspan="4" bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;"><strong>' + esc(descricaoProcessosPedido || "") + '</strong></td>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">' + formatBR(somaProcessosPedido) + '</td>'
        + '</tr>';
    }

    // NOVO: Gera tabela de parcelas se houver múltiplas parcelas
    let tabelaParcelasHtml = "";
    if (infoPagamento && infoPagamento.texto) {
      const parcelas = calcularParcelas(infoPagamento.texto, totalFinal);

      if (parcelas && parcelas.length > 1) {
        const porCondicao = parcelas.some(function (p) { return p.condicao != null && p.condicao !== ""; });
        const headerSegundaCol = porCondicao ? "Condição" : "Dias";
        const celulaSegundaCol = function (p) {
          if (p.condicao != null && p.condicao !== "") return p.condicao;
          return p.dias === 0 ? "À vista" : p.dias;
        };
        tabelaParcelasHtml = `
    <table cellpadding="1" cellspacing="1" style="width:auto; max-width:200px; border-collapse:collapse; margin-top:10px; margin-right:auto; font-size:7pt;">
      <tr>
        <th colspan="3" bgcolor="${headerColor}" style="background:${headerColor}; color:#fff; padding:2px; text-align:center; font-size:9pt; font-weight:bold;">
           Pagamento
        </th>
      </tr>
      <tr>
        <th bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt; text-align:center;">Parc.</th>
        <th bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt; text-align:center;">${headerSegundaCol}</th>
        <th bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt; text-align:center;">Valor</th>
      </tr>
      ${parcelas.map(p => `
        <tr>
          <td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt; text-align:center;">${p.numero}/${parcelas.length}</td>
          <td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt; text-align:center;">${celulaSegundaCol(p)}</td>
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

        <h3 style="margin-top:12px;">Informações da proposta</h3>
        <p style="font-size:8pt; line-height:1.25;">
          <b>Proposta Comercial - incluído em:</b> ${esc(dataBrasil)} às ${esc(horaBrasil)}<br>
            <b>Validade da Proposta:</b> 30 dias
          </p>

        <p style="font-size:8pt; line-height:1.25;">
          <b>Prazo de entrega:</b> ${esc(formatarDataBrasil(observacoes.prazo) || "-")}<br>
          <b>Pagamento:</b> ${esc(observacoes.pagamento || "-")}<br>
          <b>Vendedor:</b> ${esc(observacoes.vendedor || "-")}<br>
          <b>Condições do Material:</b> ${esc(observacoes.materialCond || "-")}<br>
          <b>Transporte:</b> ${esc(observacoes.transporte || "-")}<br>
        </p>

        ${observacoes.adicional ? `<p style="font-size:8pt; line-height:1.25;"><b>Observações:</b><br>${esc(observacoes.adicional)}</p>` : ""}

      </body>
      </html>
    `;

    if (apenasPreview) return { html: htmlContent };

    const blob = Utilities.newBlob(htmlContent, "text/html", "orcamento.html");
    const nomeArquivoPdf = "Proposta_" + codigoBase + "_" + numeroSequencial + versaoFinal + ".pdf";
    const pdf = blob.getAs("application/pdf").setName(nomeArquivoPdf);

    let file;
    if (sobrescreverVersao) {
      // Sobrescrever: move o antigo para lixeira e cria novo (setContent corrompe PDF)
      const arquivos = comSubFolder.getFilesByName(nomeArquivoPdf);
      if (arquivos.hasNext()) {
        const arquivoAntigo = arquivos.next();
        arquivoAntigo.setTrashed(true);
        Logger.log("📄 Arquivo antigo movido para lixeira: " + nomeArquivoPdf);
      }
      file = comSubFolder.createFile(pdf);
      Logger.log("📄 PDF sobrescrito (novo arquivo): " + nomeArquivoPdf);
    } else {
      file = comSubFolder.createFile(pdf);
    }

    const memoriaUrl = null; // Memória de Cálculo é gerada separadamente pelo botão "Salvar Memória de Cálculo"

    var urlPdfRetorno = file.getUrl();
    try {
      registrarOrcamento(cliente, codigoParaPlanilha, totalFinal, dataBrasil, urlPdfRetorno, memoriaUrl, chapas || [], observacoes, produtosCadastrados, dadosFormularioCompleto, isPedido);
    } catch (errReg) {
      Logger.log("Aviso gerarPdfOrcamento: PDF criado mas falha ao registrar na planilha - " + (errReg && errReg.message ? errReg.message : errReg));
      // Retorna mesmo assim para o cliente poder abrir o PDF e o botão não ficar travado
    }
    return { url: urlPdfRetorno, nome: file.getName(), memoriaUrl: memoriaUrl };
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
    // Ordem de Produção: sempre sobrescreve (sem v2, v3) - usa sempre o mesmo nome base
    const numeroProposta = codigoProjeto || "";

    const resultados = [];
    if (produtosCadastrados && Array.isArray(produtosCadastrados)) {
      produtosCadastrados.forEach(prod => {
        resultados.push({
          codigo: prod.codigo || "",
          descricao: prod.descricao || "",
          quantidade: prod.quantidade || 0,
          precoUnitario: 0,
          precoTotal: 0,
          processos: prod.processos && Array.isArray(prod.processos) ? prod.processos : [],
          descricoesProcessos: prod.descricoesProcessos && typeof prod.descricoesProcessos === 'object' ? prod.descricoesProcessos : {}
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

    var todosProcessosOp = ["MP", "CL", "D", "S", "Pin", "CAD", "ACB"];
    var processosPresentes = [];
    resultados.forEach(function (p) {
      var arr = p.processos;
      if (arr && Array.isArray(arr)) {
        arr.forEach(function (sigla) {
          if (todosProcessosOp.indexOf(sigla) >= 0 && processosPresentes.indexOf(sigla) < 0) {
            processosPresentes.push(sigla);
          }
        });
      }
    });

    function temProcessoOp(p, sigla) {
      var arr = p.processos;
      return arr && Array.isArray(arr) && arr.indexOf(sigla) >= 0;
    }

    const itensHtml = resultados.map(function (p) {
      var cellsProc = processosPresentes.map(function (sigla) {
        // Ordem de Produção: marca X automaticamente se o item tiver o processo
        return '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:center; font-size:8pt;">' + (temProcessoOp(p, sigla) ? "X" : "") + '</td>';
      }).join("");
      // Descrições por processo
      var descProcsHtml = "";
      if (p.processos && p.processos.length > 0 && p.descricoesProcessos) {
        var descItems = p.processos.filter(function(s) { return p.descricoesProcessos[s]; }).map(function(sigla) {
          return '<span style="font-size:7pt; color:#555;"><b>' + esc(sigla) + ':</b> ' + esc(p.descricoesProcessos[sigla]) + '</span>';
        });
        if (descItems.length > 0) {
          descProcsHtml = '<br><div style="margin-top:2px;">' + descItems.join(' | ') + '</div>';
        }
      }
      return ''
        + '<tr>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; font-size:8pt;">' + esc(p.codigo || "") + '</td>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; font-size:8pt;">' + esc(p.descricao || "") + descProcsHtml + '</td>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:center; font-size:8pt;">' + esc(p.quantidade || 0) + '</td>'
        + cellsProc
        + '</tr>';
    }).join('');

    var headerProcessosHtml = processosPresentes.map(function (sigla) {
      return '<th bgcolor="' + headerColor + '" style="background:' + headerColor + '; color:#ffffff; padding:3px; text-align:center; border:0.1px solid #fff; font-size:8pt;">' + sigla + '</th>';
    }).join("");

    var PROCESSOS_DESCRICOES = {
      "MP": "Matéria Prima",
      "CL": "Corte Laser",
      "D": "Dobra",
      "S": "Solda",
      "Pin": "Pintura",
      "CAD": "Projeto CAD",
      "ACB": "Acabamento"
    };

    var legendaProcessosHtml = "";
    if (processosPresentes.length > 0) {
      var linhasLegenda = processosPresentes.map(function (sigla) {
        var desc = PROCESSOS_DESCRICOES[sigla] || sigla;
        return '<tr>'
          + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:4px; border:0.1px solid #fff; font-size:9pt; font-weight:bold; text-align:center; width:60px;">' + sigla + '</td>'
          + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:4px; border:0.1px solid #fff; font-size:9pt;">' + desc + '</td>'
          + '</tr>';
      }).join("");

      legendaProcessosHtml = '<h3 style="margin-top:15px;">Legenda dos Processos</h3>'
        + '<table style="margin-top:8px; width:auto; max-width:350px;">'
        + '<tr>'
        + '<th bgcolor="' + headerColor + '" style="background:' + headerColor + '; color:#ffffff; padding:3px; text-align:center; border:0.1px solid #fff; font-size:9pt; width:60px;">Sigla</th>'
        + '<th bgcolor="' + headerColor + '" style="background:' + headerColor + '; color:#ffffff; padding:3px; text-align:left; border:0.1px solid #fff; font-size:9pt;">Descrição</th>'
        + '</tr>'
        + linhasLegenda
        + '</table>';
    }

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
            ${headerProcessosHtml}
          </tr>
          ${itensHtml}
        </table>

        <h3 style="margin-top:12px;">Informações do PDF</h3>
        <p style="font-size:8pt; line-height:1.25;">
          <b>Ordem de Produção - gerado em:</b> ${esc(dataBrasil)} às ${esc(horaBrasil)}
        </p>

        <p style="font-size:8pt; line-height:1.25;">
          <b>Prazo de entrega:</b> ${esc(formatarDataBrasil(observacoes.prazo) || "-")}<br>
          <b>Vendedor:</b> ${esc(observacoes.vendedor || "-")}<br>
          <b>Condições do Material:</b> ${esc(observacoes.materialCond || "-")}<br>
          <b>Transporte:</b> ${esc(observacoes.transporte || "-")}<br>
        </p>

        ${observacoes.adicional ? `<p style="font-size:8pt; line-height:1.25;"><b>Observações para o cliente:</b><br>${esc(observacoes.adicional)}</p>` : ""}

        ${legendaProcessosHtml}

      </body>
      </html>
    `;

    const blob = Utilities.newBlob(htmlContent, "text/html", "ordem_producao.html");
    const nomePdf = "Ordem_Producao_" + numeroProposta + ".pdf";
    const pdf = blob.getAs("application/pdf").setName(nomePdf);
    // Sobrescreve: remove versão anterior se existir (sempre mesmo nome, sem v2/v3)
    const arquivosCom = comSubFolder.getFilesByName(nomePdf);
    while (arquivosCom.hasNext()) {
      arquivosCom.next().setTrashed(true);
    }
    const file = comSubFolder.createFile(pdf);
    const fileUrl = file.getUrl();

    // Salva o link da Ordem de Produção na aba Projetos (coluna LINK ORDEM DE PRODUÇÃO) para exibir na página
    try {
      const linha = typeof linhaOuKey === "number" ? linhaOuKey : parseInt(linhaOuKey, 10);
      if (linha >= 2) {
        const sheetProj = ss.getSheetByName("Projetos");
        if (sheetProj) {
          const headers = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
          let colIdx = headers.indexOf("LINK ORDEM DE PRODUÇÃO");
          if (colIdx < 0) colIdx = headers.indexOf("LINK ORDEM DE PRODUCAO");
          if (colIdx < 0) {
            sheetProj.getRange(1, sheetProj.getLastColumn() + 1).setValue("LINK ORDEM DE PRODUÇÃO");
            colIdx = sheetProj.getLastColumn() - 1;
          }
          sheetProj.getRange(linha, colIdx + 1).setValue(fileUrl);
        }
      }
    } catch (e) {
      Logger.log("Aviso: não foi possível salvar link da Ordem de Produção na planilha: " + (e.message || e));
    }

    return { url: fileUrl, nome: file.getName() };
  } catch (err) {
    Logger.log("ERRO gerarPdfOrdemProducao: " + err.toString());
    throw err;
  }
}

/* ======= ORDEM DE COMPRA (PDF com itens selecionados e preços informados pelo usuário) ======= */
/**
 * Retorna a lista de itens (PRDs) do projeto para o usuário escolher quais entram na Ordem de Compra e informar preços.
 * @param {number|string} linhaOuKey - Número da linha do projeto na planilha
 * @returns {{ itens: Array<{codigo:string, descricao:string, quantidade:number}>, codigoProjeto: string }}
 */
function getItensProjetoParaOrdemCompra(linhaOuKey) {
  try {
    const dados = carregarRascunho(linhaOuKey);
    if (!dados) throw new Error("Não foi possível carregar os dados do orçamento");
    const produtosCadastrados = dados.produtosCadastrados || [];
    const projeto = dados.projeto || {};
    const codigoProjeto = (projeto.data || "") + (projeto.indice || "") + (projeto.iniciais || "");
    const itens = (produtosCadastrados || []).map(function (p) {
      return {
        codigo: p.codigo || "",
        descricao: p.descricao || "",
        quantidade: parseFloat(p.quantidade) || 0
      };
    });
    return { itens: itens, codigoProjeto: codigoProjeto };
  } catch (err) {
    Logger.log("ERRO getItensProjetoParaOrdemCompra: " + err.toString());
    throw err;
  }
}

/**
 * Gera PDF da Ordem de Compra com apenas os itens que possuem valor (unitário ou total).
 * Inclui informações do fornecedor (destinatário da ordem), no mesmo formato do cliente no orçamento.
 * @param {number|string} linhaOuKey - Linha do projeto
 * @param {Array<{codigo:string, descricao:string, quantidade:number, valorUnitario?:number, valorTotal?:number}>} itensComValor - Itens com pelo menos valorUnitario ou valorTotal
 * @param {{ nome:string, cpf:string, endereco:string, telefone:string, email:string, responsavel?:string }} fornecedor - Dados do fornecedor (mesma estrutura de cliente)
 * @returns {{ url: string, nome: string }}
 */
function gerarPdfOrdemCompra(linhaOuKey, itensComValor, fornecedor) {
  try {
    var itensFiltrados = (itensComValor || []).filter(function (item) {
      var q = parseFloat(item.quantidade) || 0;
      var vu = parseFloat(item.valorUnitario);
      var vt = parseFloat(item.valorTotal);
      if (!isNaN(vu) && vu > 0) return true;
      if (!isNaN(vt) && vt > 0) return true;
      return false;
    });
    if (itensFiltrados.length === 0) {
      throw new Error("Inclua pelo menos um item com valor (unitário ou total) para gerar a Ordem de Compra.");
    }

    var dados = carregarRascunho(linhaOuKey);
    if (!dados) throw new Error("Não foi possível carregar os dados do orçamento");
    var projeto = dados.projeto || {};
    var cliente = dados.cliente || {};
    var codigoProjeto = (projeto.data || "") + (projeto.indice || "") + (projeto.iniciais || "");
    var data = projeto.data || "";
    // Não usar projeto.pasta aqui: Ordem de Compra não deve alterar o nome da pasta do projeto (evita usar nome do fornecedor por engano). Apenas localizar/criar pela estrutura código+data.
    var pasta = criarOuUsarPasta(codigoProjeto, "", data);
    var workFolder = getOrCreateSubFolder(pasta, "02_WORK");
    var comSubFolder = getOrCreateSubFolder(workFolder, "COM");

    var logoFile = DriveApp.getFileById(ID_LOGO);
    var logoBlob = logoFile.getBlob();
    var logoBase64 = Utilities.base64Encode(logoBlob.getBytes());
    var logoMime = logoBlob.getContentType();
    var agora = new Date();
    var dataBrasil = formatarDataBrasil(agora);
    var horaBrasil = agora.toLocaleTimeString("pt-BR");
    var headerColor = "#FF9933";
    var rowColor = "#FDF5E6";

    function esc(v) {
      if (v === null || v === undefined) return "";
      return String(v)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#39;");
    }

    function formatarMoeda(num) {
      if (num == null || isNaN(num)) return "";
      return Utilities.formatString("%.2f", num).replace(".", ",");
    }

    var linhas = itensFiltrados.map(function (p) {
      var qtd = parseFloat(p.quantidade) || 0;
      var vu = parseFloat(p.valorUnitario);
      var vt = parseFloat(p.valorTotal);
      if (isNaN(vu) || vu <= 0) vu = (qtd > 0 && !isNaN(vt) && vt > 0) ? vt / qtd : 0;
      if (isNaN(vt) || vt <= 0) vt = vu * qtd;
      return (
        "<tr>" +
        "<td bgcolor=\"" + rowColor + "\" style=\"background:" + rowColor + "; padding:2px; border:0.1px solid #fff; font-size:8pt;\">" + esc(p.codigo || "") + "</td>" +
        "<td bgcolor=\"" + rowColor + "\" style=\"background:" + rowColor + "; padding:2px; border:0.1px solid #fff; font-size:8pt;\">" + esc(p.descricao || "") + "</td>" +
        "<td bgcolor=\"" + rowColor + "\" style=\"background:" + rowColor + "; padding:2px; border:0.1px solid #fff; text-align:center; font-size:8pt;\">" + esc(qtd) + "</td>" +
        "<td bgcolor=\"" + rowColor + "\" style=\"background:" + rowColor + "; padding:2px; border:0.1px solid #fff; text-align:right; font-size:8pt;\">R$ " + formatarMoeda(vu) + "</td>" +
        "<td bgcolor=\"" + rowColor + "\" style=\"background:" + rowColor + "; padding:2px; border:0.1px solid #fff; text-align:right; font-size:8pt;\">R$ " + formatarMoeda(vt) + "</td>" +
        "</tr>"
      );
    }).join("");

    var htmlContent =
      "<html><head><meta charset=\"utf-8\"><style>" +
      "body, table, th, td { -webkit-print-color-adjust: exact; print-color-adjust: exact; }" +
      "body { font-family: Arial, sans-serif; font-size: 9pt; color: #000; margin: 5px; line-height:1.2; }" +
      ".header { display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; }" +
      ".logo { max-height:180px; } .company-info { text-align:right; font-size:9pt; }" +
      "h2 { text-align:left; margin:20px 0 30px 0; font-size:14pt; } h3 { margin-top:15px; margin-bottom:5px; font-size:11pt; }" +
      "table { width:100%; border-collapse:collapse; font-size:8pt; }" +
      "</style></head><body style=\"-webkit-print-color-adjust: exact; print-color-adjust: exact;\">" +
      "<div class=\"header\">" +
      "<img class=\"logo\" src=\"data:" + logoMime + ";base64," + logoBase64 + "\">" +
      "<div class=\"company-info\"><strong>TUBA FERRAMENTARIA LTDA</strong><br>CNPJ: 10.684.825/0001-26<br>Inscrição Estadual: 635592888110<br>Endereço: Estrada Dos Alvarengas, 4101 - Assunção<br>São Bernardo do Campo - SP - CEP: 09850-550<br>Site: www.tb4.com.br<br><b>Email:</b> tubaferram@gmail.com<br><b>Telefone:</b> (11) 91285-4204</div>" +
      "</div>" +
      "<h2>Ordem de Compra Nº " + esc(codigoProjeto) + "</h2>" +
      "<h3>Informações do Fornecedor (destinatário da ordem):</h3>" +
      "<p style=\"margin-bottom:12px; font-size:9pt; line-height:1.3;\">" +
      "<strong>" + esc((fornecedor && fornecedor.nome) ? fornecedor.nome : "") + "</strong><br>" +
      "CNPJ/CPF: " + esc((fornecedor && fornecedor.cpf) ? fornecedor.cpf : "-") + "<br>" +
      esc((fornecedor && fornecedor.endereco) ? fornecedor.endereco : "-") + "<br>" +
      "<b>Telefone:</b> " + esc((fornecedor && fornecedor.telefone) ? fornecedor.telefone : "-") + "<br>" +
      "<b>Email:</b> " + esc((fornecedor && fornecedor.email) ? fornecedor.email : "-") + "<br>" +
      (fornecedor && fornecedor.responsavel ? "<b>Responsável:</b> " + esc(fornecedor.responsavel) : "") + "</p>" +
      "<h3>Itens da Ordem de Compra</h3>" +
      "<table style=\"margin-top:8px;\">" +
      "<tr>" +
      "<th bgcolor=\"" + headerColor + "\" style=\"background:" + headerColor + "; color:#fff; padding:3px; text-align:left; border:0.1px solid #fff; font-size:8pt;\">Código</th>" +
      "<th bgcolor=\"" + headerColor + "\" style=\"background:" + headerColor + "; color:#fff; padding:3px; text-align:left; border:0.1px solid #fff; font-size:8pt;\">Descrição</th>" +
      "<th bgcolor=\"" + headerColor + "\" style=\"background:" + headerColor + "; color:#fff; padding:3px; text-align:center; border:0.1px solid #fff; font-size:8pt;\">Quantidade</th>" +
      "<th bgcolor=\"" + headerColor + "\" style=\"background:" + headerColor + "; color:#fff; padding:3px; text-align:right; border:0.1px solid #fff; font-size:8pt;\">Valor Unitário</th>" +
      "<th bgcolor=\"" + headerColor + "\" style=\"background:" + headerColor + "; color:#fff; padding:3px; text-align:right; border:0.1px solid #fff; font-size:8pt;\">Valor Total</th>" +
      "</tr>" + linhas + "</table>" +
      "<p style=\"font-size:8pt; margin-top:12px;\"><b>Ordem de Compra - gerado em:</b> " + esc(dataBrasil) + " às " + esc(horaBrasil) + "</p>" +
      "</body></html>";

    var blob = Utilities.newBlob(htmlContent, "text/html", "ordem_compra.html");
    var nomePdf = "Ordem_Compra_" + codigoProjeto + ".pdf";
    var pdf = blob.getAs("application/pdf").setName(nomePdf);
    var arquivosCom = comSubFolder.getFilesByName(nomePdf);
    while (arquivosCom.hasNext()) { arquivosCom.next().setTrashed(true); }
    var file = comSubFolder.createFile(pdf);
    var fileUrl = file.getUrl();

    // Salva o link da Ordem de Compra na aba Projetos (coluna LINK ORDEM DE COMPRA) para exibir na página
    try {
      var linhaNum = typeof linhaOuKey === "number" ? linhaOuKey : parseInt(linhaOuKey, 10);
      if (linhaNum >= 2) {
        var sheetProj = ss.getSheetByName("Projetos");
        if (sheetProj) {
          var headers = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
          var colIdx = headers.indexOf("LINK ORDEM DE COMPRA");
          if (colIdx < 0) {
            sheetProj.getRange(1, sheetProj.getLastColumn() + 1).setValue("LINK ORDEM DE COMPRA");
            colIdx = sheetProj.getLastColumn() - 1;
          }
          sheetProj.getRange(linhaNum, colIdx + 1).setValue(fileUrl);
        }
      }
    } catch (e) {
      Logger.log("Aviso: não foi possível salvar link da Ordem de Compra na planilha: " + (e.message || e));
    }

    return { url: fileUrl, nome: file.getName() };
  } catch (err) {
    Logger.log("ERRO gerarPdfOrdemCompra: " + err.toString());
    throw err;
  }
}

// ----------------- MODIFICAÇÃO: registrarOrcamento -----------------
function registrarOrcamento(cliente, codigoProjeto, valorTotal, dataOrcamento, urlPdf, urlMemoria, chapas, observacoes, produtosCadastrados, dadosFormularioCompleto, isPedido) {
  // Calcula a união dos processos de todos os itens (produtos cadastrados)
  const todosProcessos = [];
  const ordemProcessos = ["MP", "CL", "D", "S", "Pin", "CAD", "ACB"];
  (produtosCadastrados || []).forEach(function(prod) {
    var procs = prod.processos;
    if (procs && Array.isArray(procs)) {
      procs.forEach(function(sigla) {
        if (todosProcessos.indexOf(sigla) < 0) todosProcessos.push(sigla);
      });
    }
  });
  // Ordena na ordem padrão
  todosProcessos.sort(function(a, b) {
    return ordemProcessos.indexOf(a) - ordemProcessos.indexOf(b);
  });
  // Se o usuário digitou algo manual no campo obsProcessos, usa isso; senão usa a união dos itens
  const processosManual = (observacoes && observacoes.processos) ? String(observacoes.processos).trim() : "";
  const processosStr = processosManual || todosProcessos.join(", ");

  // Extrai descrição e prazo das observações
  const descricao = (observacoes && observacoes.descricao) || "";
  const prazo = (observacoes && observacoes.prazo) || "";

  // Atribui PRD apenas a produtos cadastrados (chapas/peças não são mais usados no orçamento)
  chapas = chapas || [];

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
    
    const dadosParaJson = dadosFormularioCompleto ? { ...dadosFormularioCompleto } : {};
    dadosParaJson.chapas = chapas || [];
    dadosParaJson.produtosCadastrados = produtosCadastrados || [];
    if (numeroSequencial != null) dadosParaJson.numeroSequencial = numeroSequencial;
    if (!dadosParaJson.observacoes) dadosParaJson.observacoes = {};
    dadosParaJson.observacoes.projeto = codigoProjeto; // Garante PROJETO com versão (_v2) no JSON
    
    const agora = new Date();
    const dadosJson = JSON.stringify({
      nome: codigoProjeto,
      dataSalvo: agora.toISOString(),
      numeroSequencial: numeroSequencial,
      dados: dadosParaJson
    });

    // usar: Projetos 
    const sheetProj = ss.getSheetByName("Projetos");
    const targetSheet = sheetProj;

    // 1) Se o formulário enviou a linha do projeto carregado, usar essa linha (evita duplicar ao gerar PDF)
    var linhaExistente = 0;
    const linhaProjetoForm = (dadosFormularioCompleto && dadosFormularioCompleto.linhaProjeto != null) ? parseInt(dadosFormularioCompleto.linhaProjeto, 10) : NaN;
    if (sheetProj && !isNaN(linhaProjetoForm) && linhaProjetoForm >= 2 && linhaProjetoForm <= sheetProj.getLastRow()) {
      const headersProj = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
      const idxProjeto = _findHeaderIndexProjetos(headersProj, "PROJETO");
      if (idxProjeto >= 0) {
        const valorNaLinha = sheetProj.getRange(linhaProjetoForm, idxProjeto + 1).getValue();
        if (String(valorNaLinha || "").trim() === String(codigoProjeto || "").trim()) {
          linhaExistente = linhaProjetoForm;
        }
      }
    }
    // 2) Senão, buscar por código do projeto na coluna PROJETO
    if (!linhaExistente && sheetProj) {
      linhaExistente = findRowByColumnValue(sheetProj, "PROJETO", codigoProjeto) || 0;
    }
    const statusOrcamento = isPedido ? "Convertido em Pedido" : "Enviado";
    const statusPedidoInicial = isPedido ? "Processo de Preparação MP / CAD / CAM" : "";
    const observacoesKanban = (observacoes && observacoes.observacoesKanban != null) ? String(observacoes.observacoesKanban).trim() : "";
    const prazoProposta = (observacoes && observacoes.prazoProposta != null) ? String(observacoes.prazoProposta).trim() : "";
    var dadosObjReg = {
      "CLIENTE": cliente.nome || "",
      "DESCRIÇÃO": descricao,
      "RESPONSÁVEL CLIENTE": cliente.responsavel || "",
      "PROJETO": codigoProjeto || "",
      "VALOR TOTAL": valorTotal || "",
      "DATA": dataOrcamento || "",
      "PROCESSOS": processosStr || "",
      "LINK DO PDF": urlPdf || "",
      "LINK DA MEMÓRIA DE CÁLCULO": urlMemoria || "",
      "STATUS_ORCAMENTO": statusOrcamento,
      "STATUS_PEDIDO": statusPedidoInicial,
      "PRAZO": prazo,
      "PRAZO_PROPOSTA": prazoProposta,
      "OBSERVAÇÕES": observacoesKanban,
      "JSON_DADOS": dadosJson
    };
    if (linhaExistente) {
      var rowAtualReg = targetSheet.getRange(linhaExistente, 1, linhaExistente, targetSheet.getLastColumn()).getValues()[0];
      var headersReg = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
      var idxData = _findHeaderIndexProjetos(headersReg, "DATA");
      var statusAtualReg = "";
      var idxOrc = _findHeaderIndexProjetos(headersReg, "STATUS_ORCAMENTO");
      if (idxOrc >= 0 && rowAtualReg[idxOrc]) statusAtualReg = String(rowAtualReg[idxOrc] || "").trim();
      // Projeto já convertido em pedido: não alterar DATA ao gerar novo PDF (data de competência só na aba Pedidos e editável só no modal de Pedidos)
      if (statusAtualReg === "Convertido em Pedido" && idxData >= 0 && rowAtualReg[idxData] != null && String(rowAtualReg[idxData]).trim() !== "") {
        dadosObjReg["DATA"] = rowAtualReg[idxData];
      }
      if (!isPedido) {
        var idxPed = _findHeaderIndexProjetos(headersReg, "STATUS_PEDIDO");
        var idxObs = _findHeaderIndexProjetos(headersReg, "OBSERVAÇÕES");
        if (idxOrc >= 0 && rowAtualReg[idxOrc]) dadosObjReg["STATUS_ORCAMENTO"] = rowAtualReg[idxOrc];
        if (idxPed >= 0 && rowAtualReg[idxPed]) dadosObjReg["STATUS_PEDIDO"] = rowAtualReg[idxPed];
        if (idxObs >= 0 && observacoesKanban === "" && rowAtualReg[idxObs]) dadosObjReg["OBSERVAÇÕES"] = rowAtualReg[idxObs];
      }
      _escreverLinhaProjetosPorCabecalho(targetSheet, linhaExistente, dadosObjReg, true);
    } else {
      // Evitar duplicata: garantir que não existe linha com mesmo PROJETO (busca por nome de coluna)
      var codigoNorm = String(codigoProjeto || "").trim();
      if (sheetProj && codigoNorm) {
        var headersProj2 = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
        var idxProj2 = _findHeaderIndexProjetos(headersProj2, "PROJETO");
        if (idxProj2 >= 0) {
          var lastR = sheetProj.getLastRow();
          for (var ri = 2; ri <= lastR; ri++) {
            var cellVal = sheetProj.getRange(ri, idxProj2 + 1).getValue();
            if (String(cellVal || "").trim() === codigoNorm) {
              linhaExistente = ri;
              var rowAtualDup = sheetProj.getRange(ri, 1, ri, sheetProj.getLastColumn()).getValues()[0];
              var idxDataDup = _findHeaderIndexProjetos(headersProj2, "DATA");
              var idxOrcDup = _findHeaderIndexProjetos(headersProj2, "STATUS_ORCAMENTO");
              var statusDup = (idxOrcDup >= 0 && rowAtualDup[idxOrcDup]) ? String(rowAtualDup[idxOrcDup] || "").trim() : "";
              if (statusDup === "Convertido em Pedido" && idxDataDup >= 0 && rowAtualDup[idxDataDup] != null && String(rowAtualDup[idxDataDup]).trim() !== "") dadosObjReg["DATA"] = rowAtualDup[idxDataDup];
              if (!isPedido && idxOrcDup >= 0 && rowAtualDup[idxOrcDup]) dadosObjReg["STATUS_ORCAMENTO"] = rowAtualDup[idxOrcDup];
              _escreverLinhaProjetosPorCabecalho(targetSheet, ri, dadosObjReg, true);
              break;
            }
          }
        }
      }
      if (!linhaExistente) _escreverLinhaProjetosPorCabecalho(targetSheet, targetSheet.getLastRow() + 1, dadosObjReg, false);
    }

    // Sincronizar aba Pedidos quando atualizamos uma linha existente que já é pedido (ex.: gerar PDF de pedido existente)
    if (linhaExistente && sheetProj && (codigoProjeto || "").trim() !== "") {
      try {
        var headersSync = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
        var rowSync = sheetProj.getRange(linhaExistente, 1, linhaExistente, sheetProj.getLastColumn()).getValues()[0];
        var idxStatusSync = _findHeaderIndexProjetos(headersSync, "STATUS_ORCAMENTO");
        var statusSync = (idxStatusSync >= 0 && rowSync[idxStatusSync] != null) ? String(rowSync[idxStatusSync]).trim() : "";
        if (statusSync === "Convertido em Pedido") {
          ensurePedidoRow(linhaExistente);
        }
      } catch (errSync) {
        Logger.log("Aviso ao sincronizar Projetos→Pedidos em registrarOrcamento: " + (errSync && errSync.message));
      }
    }

    // Quando salvou como pedido: renomear pasta COT→PED e sincronizar aba Pedidos
    if (isPedido && sheetProj && (codigoProjeto || "").length >= 6) {
      try {
        const codigoBase = String(codigoProjeto).replace(/_v\d+$/, "").trim();
        const dataProj = codigoBase.substring(0, 6);
        atualizarPrefixoPastaParaPedido(codigoBase, dataProj, cliente.nome || "", descricao);
      } catch (ePasta) {
        Logger.log("Aviso ao renomear pasta COT→PED em registrarOrcamento: " + (ePasta && ePasta.message));
      }
      const linhaPedido = linhaExistente || (targetSheet ? targetSheet.getLastRow() : 0);
      if (linhaPedido >= 2) ensurePedidoRow(linhaPedido);
    }

    // Insere os produtos cadastrados na "Relação de produtos" apenas quando é projeto NOVO.
    // Ao editar e gerar PDF de projeto existente, não reinsere (evita timeout: cada inserirProdutoNaRelacao lê a planilha inteira).
    if (!linhaExistente && produtosCadastrados && Array.isArray(produtosCadastrados)) {
      produtosCadastrados.forEach(function (prod) {
        const codigo = (prod.codigo && String(prod.codigo).trim()) || "";
        if (codigo && String(codigo).toUpperCase().startsWith("PRD")) {
          const produtoRelacao = {
            codigo: codigo,
            descricao: prod.descricao || "",
            ncm: prod.ncm || "",
            preco: Number(prod.precoUnitario) || 0,
            unidade: prod.unidade || "UN",
            caracteristicas: "",
            projeto: codigoProjeto || "",
            cliente: cliente.nome || "",
            processos: prod.processos && Array.isArray(prod.processos) ? prod.processos : []
          };
          inserirProdutoNaRelacao(produtoRelacao);
        }
      });
    }

  } catch (err) {
    Logger.log("Erro ao registrarOrcamento (atualizar/inserir): " + err);
    // fallback: gravar por cabeçalho (mesma lógica, evita coluna errada)
    try {
      const numeroSequencial = (dadosFormularioCompleto && dadosFormularioCompleto.numeroSequencial) || null;
      const dadosParaJson = dadosFormularioCompleto ? { ...dadosFormularioCompleto } : {};
      dadosParaJson.chapas = chapas || [];
      dadosParaJson.produtosCadastrados = produtosCadastrados || [];
      if (numeroSequencial != null) dadosParaJson.numeroSequencial = numeroSequencial;
      const agora = new Date();
      const dadosJson = JSON.stringify({
        nome: codigoProjeto,
        dataSalvo: agora.toISOString(),
        numeroSequencial: numeroSequencial,
        dados: dadosParaJson
      });
      const observacoesKanbanFallback = (observacoes && observacoes.observacoesKanban != null) ? String(observacoes.observacoesKanban).trim() : "";
      const prazoPropostaFallback = (observacoes && observacoes.prazoProposta != null) ? String(observacoes.prazoProposta).trim() : "";
      const sheetProj = ss.getSheetByName("Projetos");
      if (sheetProj) {
        var dadosObjFallback = {
          "CLIENTE": cliente.nome || "",
          "DESCRIÇÃO": descricao,
          "RESPONSÁVEL CLIENTE": cliente.responsavel || "",
          "PROJETO": codigoProjeto || "",
          "VALOR TOTAL": valorTotal || "",
          "DATA": dataOrcamento || "",
          "PROCESSOS": processosStr || "",
          "LINK DO PDF": urlPdf || "",
          "LINK DA MEMÓRIA DE CÁLCULO": urlMemoria || "",
          "STATUS_ORCAMENTO": "Rascunho",
          "STATUS_PEDIDO": "",
          "PRAZO": prazo,
          "PRAZO_PROPOSTA": prazoPropostaFallback,
          "OBSERVAÇÕES": observacoesKanbanFallback,
          "JSON_DADOS": dadosJson
        };
        var linhaFallback = (codigoProjeto && findRowByColumnValue(sheetProj, "PROJETO", codigoProjeto)) || null;
        if (linhaFallback) {
          _escreverLinhaProjetosPorCabecalho(sheetProj, linhaFallback, dadosObjFallback, true);
        } else {
          _escreverLinhaProjetosPorCabecalho(sheetProj, sheetProj.getLastRow() + 1, dadosObjFallback, false);
        }
      }
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
              caracteristicas: "",
              projeto: codigoProjeto || "",
              cliente: cliente.nome || ""
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

/**
 * Retorna resumo operacional para o dashboard: orçamentos em aberto, entregas esta semana, pedidos pendentes.
 * @returns {{ orcamentosEmAberto: number, entregasEstaSemana: number, pedidosPendentes: number }}
 */
function getDashboardResumoOperacional() {
  try {
    var orcamentosEmAberto = 0;
    var entregasEstaSemana = 0;
    var pedidosPendentes = 0;

    var projetos = getProjetos();
    for (var i = 0; i < projetos.length; i++) {
      var st = (projetos[i].STATUS_ORCAMENTO || projetos[i]["STATUS ORCAMENTO"] || "").toString().trim().toLowerCase();
      if (st === "enviado" || st === "rascunho") orcamentosEmAberto++;
    }

    var pedidos = getPedidos();
    var hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    var fimSemana = new Date(hoje);
    fimSemana.setDate(fimSemana.getDate() + 7);

    function parseDataBr(s) {
      if (!s) return null;
      var m = (s + "").trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) {
        var d = new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
        return isNaN(d.getTime()) ? null : d;
      }
      return null;
    }

    for (var j = 0; j < pedidos.length; j++) {
      var p = pedidos[j];
      var statusPag = (p["STATUS_PAGAMENTO"] || p["STATUS PAGAMENTO"] || "Pendente").toString().trim();
      if (statusPag !== "Pago") pedidosPendentes++;

      var dataEntregaStr = (p["DATA_ENTREGA"] || p["DATA DE ENTREGA"] || p.PRAZO || "").toString().trim();
      if (!dataEntregaStr && p._dataVencimento) {
        var first = (p._dataVencimento + "").split(",")[0];
        dataEntregaStr = first ? first.trim() : "";
      }
      var dataEnt = parseDataBr(dataEntregaStr);
      if (dataEnt) {
        dataEnt.setHours(0, 0, 0, 0);
        if (dataEnt.getTime() >= hoje.getTime() && dataEnt.getTime() <= fimSemana.getTime()) entregasEstaSemana++;
      }
    }

    return { orcamentosEmAberto: orcamentosEmAberto, entregasEstaSemana: entregasEstaSemana, pedidosPendentes: pedidosPendentes };
  } catch (e) {
    Logger.log("getDashboardResumoOperacional error: " + e.message);
    return { orcamentosEmAberto: 0, entregasEstaSemana: 0, pedidosPendentes: 0 };
  }
}

/**
 * Retorna resumo financeiro (total recebido e total a receber) para o card do dashboard.
 * @returns {{ recebido: number, aReceber: number }}
 */
function getFinanceiroResumo() {
  try {
    var pedidos = getPedidos();
    var recebido = 0;
    var aReceber = 0;
    for (var i = 0; i < pedidos.length; i++) {
      var p = pedidos[i];
      recebido += _parseCurrency(p["VALOR_PAGO"] || p["VALOR PAGO"]);
      if (p._valorRestante != null) {
        aReceber += p._valorRestante;
      }
    }
    return { recebido: recebido, aReceber: aReceber };
  } catch (e) {
    Logger.log("getFinanceiroResumo error: " + e.message);
    return { recebido: 0, aReceber: 0 };
  }
}

/**
 * Retorna as últimas atividades para exibir no dashboard.
 * 1) Se existir aba "Log de Atividades" ou "Atividades" com colunas Data e Descrição/Texto, usa as últimas 5 linhas.
 * 2) Caso contrário, monta a partir dos projetos mais recentes (Projetos).
 * @returns {Array<{ texto: string, data?: string }>}
 */
function getUltimasAtividades() {
  try {
    var sheetAtiv = ss.getSheetByName("Log de Atividades") || ss.getSheetByName("Atividades");
    if (sheetAtiv && sheetAtiv.getLastRow() >= 2) {
      var dados = sheetAtiv.getDataRange().getValues();
      var headers = dados[0].map(function (h) { return (h || "").toString().trim(); });
      var idxData = _findHeaderIndex(headers, "Data");
      if (idxData < 0) idxData = _findHeaderIndex(headers, "DataHora") >= 0 ? _findHeaderIndex(headers, "DataHora") : 0;
      var idxTexto = _findHeaderIndex(headers, "Descrição");
      if (idxTexto < 0) idxTexto = _findHeaderIndex(headers, "Texto");
      if (idxTexto < 0) idxTexto = _findHeaderIndex(headers, "Descricao");
      if (idxTexto < 0) idxTexto = 1;
      var lista = [];
      for (var i = dados.length - 1; i >= 1 && lista.length < 5; i--) {
        var row = dados[i];
        var texto = (row[idxTexto] != null && row[idxTexto] !== "") ? String(row[idxTexto]).trim() : "";
        if (!texto) continue;
        var dataVal = row[idxData];
        var dataStr = "";
        if (dataVal != null && dataVal !== "") {
          if (Object.prototype.toString.call(dataVal) === "[object Date]") {
            try {
              dataStr = Utilities.formatDate(dataVal, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || "UTC", "dd/MM/yyyy HH:mm");
            } catch (e) { dataStr = String(dataVal); }
          } else {
            dataStr = String(dataVal);
          }
        }
        lista.push({ texto: texto, data: dataStr });
      }
      return lista;
    }

    // Fallback: últimos projetos — ordena por data decrescente aqui para manter "5 mais recentes" no dashboard
    var projetos = getProjetos();
    var parseBr = function (s) {
      if (!s) return 0;
      var m = (s + '').trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) return parseInt(m[3] + ('0' + m[2]).slice(-2) + ('0' + m[1]).slice(-2), 10);
      return 0;
    };
    projetos = projetos.slice().sort(function (a, b) {
      var ta = parseBr(a.DATA);
      var tb = parseBr(b.DATA);
      if (ta !== tb) return tb - ta;
      return (b._linhaPlanilha || 0) - (a._linhaPlanilha || 0);
    });
    var out = [];
    for (var j = 0; j < projetos.length && out.length < 5; j++) {
      var proj = projetos[j];
      var cod = (proj.PROJETO || proj["Número do Projeto"] || "").toString().trim();
      var cliente = (proj.CLIENTE || proj.Cliente || "").toString().trim();
      var status = (proj.STATUS_ORCAMENTO || proj["STATUS ORCAMENTO"] || "").toString().trim();
      var dataProj = (proj.DATA || "").toString().trim();
      var texto = "Projeto " + (cod || "—") + (cliente ? " - " + cliente : "") + (status ? " - " + status : "");
      out.push({ texto: texto, data: dataProj });
    }
    return out;
  } catch (e) {
    Logger.log("getUltimasAtividades error: " + e.message);
    return [];
  }
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

/** Nomes das colunas da aba Projetos (ordem lógica; a escrita usa índice por cabeçalho) */
var PROJETOS_CAMPOS = ["CLIENTE", "DESCRIÇÃO", "RESPONSÁVEL CLIENTE", "PROJETO", "VALOR TOTAL", "DATA", "PROCESSOS", "LINK DO PDF", "LINK DA MEMÓRIA DE CÁLCULO", "STATUS_ORCAMENTO", "STATUS_PEDIDO", "PRAZO", "PRAZO_PROPOSTA", "OBSERVAÇÕES", "JSON_DADOS"];
/** Nome do cabeçalho na planilha quando a coluna for criada (ex.: PRAZO_PROPOSTA -> "PRAZO PROPOSTA") */
var PROJETOS_HEADER_DISPLAY = { "PRAZO_PROPOSTA": "PRAZO PROPOSTA", "RESPONSÁVEL CLIENTE": "RESPONSÁVEL CLIENTE", "VALOR TOTAL": "VALOR TOTAL", "LINK DO PDF": "LINK DO PDF", "LINK DA MEMÓRIA DE CÁLCULO": "LINK DA MEMÓRIA DE CÁLCULO", "STATUS_ORCAMENTO": "STATUS_ORCAMENTO", "STATUS_PEDIDO": "STATUS_PEDIDO", "JSON_DADOS": "JSON_DADOS" };
/** Sinônimos para achar coluna na planilha (cabeçalho pode estar com nome diferente) */
var PROJETOS_HEADER_SINONIMOS = {
  "PROJETO": ["PROJETO", "Projeto", "Número do Projeto"],
  "OBSERVAÇÕES": ["OBSERVAÇÕES", "OBSERVAÇÕES INTERNAS", "Observações", "Observacoes"],
  "PRAZO_PROPOSTA": ["PRAZO_PROPOSTA", "PRAZO PROPOSTA"],
  "JSON_DADOS": ["JSON_DADOS", "JSON DADOS"]
};

/**
 * Busca índice do cabeçalho, tentando o nome e sinônimos (evita gravar em coluna errada quando a planilha tem nome diferente).
 */
function _findHeaderIndexProjetos(headers, campo) {
  var idx = _findHeaderIndex(headers, campo);
  if (idx >= 0) return idx;
  var sinonimos = PROJETOS_HEADER_SINONIMOS[campo];
  if (sinonimos) {
    for (var s = 0; s < sinonimos.length; s++) {
      idx = _findHeaderIndex(headers, sinonimos[s]);
      if (idx >= 0) return idx;
    }
  }
  return -1;
}

/**
 * Escreve na aba Projetos por nome de coluna: grava célula a célula (linha, coluna do cabeçalho) para não deslocar colunas.
 * Estrutura esperada: CLIENTE, DESCRIÇÃO, RESPONSÁVEL CLIENTE, PROJETO, VALOR TOTAL, DATA, PROCESSOS, LINK DO PDF, LINK DA MEMÓRIA DE CÁLCULO, STATUS_ORCAMENTO, STATUS_PEDIDO, PRAZO, OBSERVAÇÕES, JSON_DADOS, ... (PRAZO PROPOSTA no final).
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheetProj - Aba Projetos
 * @param {number} linha - Número da linha (1-based). Se isUpdate=false, usa getLastRow()+1.
 * @param {Object} dadosObj - Objeto com chaves = nomes de coluna e valores a gravar
 * @param {boolean} isUpdate - Se true, atualiza só as células dos campos enviados; se false, escreve em nova linha só as células enviadas
 */
function _escreverLinhaProjetosPorCabecalho(sheetProj, linha, dadosObj, isUpdate) {
  var lastCol = sheetProj.getLastColumn();
  var headers = sheetProj.getRange(1, 1, 1, lastCol).getValues()[0];
  var rowToWrite = isUpdate ? linha : (sheetProj.getLastRow() + 1);
  for (var campo in dadosObj) {
    if (!Object.prototype.hasOwnProperty.call(dadosObj, campo)) continue;
    var idx = _findHeaderIndexProjetos(headers, campo);
    // Evitar gravar JSON_DADOS na coluna de CONDIÇÕES DE PAGAMENTO (planilha pode ter ordem diferente)
    if (campo === "JSON_DADOS" && idx >= 0 && headers[idx] != null) {
      var normCol = _normalizeHeader(headers[idx]);
      if (normCol === "condicoesdepagamento") idx = -1;
    }
    if (idx < 0) {
      var newCol = sheetProj.getLastColumn() + 1;
      var headerName = PROJETOS_HEADER_DISPLAY[campo] || campo;
      sheetProj.getRange(1, newCol).setValue(headerName);
      headers = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
      idx = headers.length - 1;
    }
    var val = dadosObj[campo] != null ? dadosObj[campo] : "";
    sheetProj.getRange(rowToWrite, idx + 1).setValue(val);
  }
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
        const idxPrazoProposta = _findHeaderIndex(headersProj, "PRAZO_PROPOSTA");
        const idxProcessos = _findHeaderIndex(headersProj, "PROCESSOS");
        const idxObs = _findHeaderIndex(headersProj, "OBSERVAÇÕES");
        const idxJsonDados = _findHeaderIndex(headersProj, "JSON_DADOS");

        for (let i = 1; i < valsProj.length; i++) {
          const row = valsProj[i];
          const cliente = idxCliente >= 0 ? row[idxCliente] : "";
          const projeto = idxProjeto >= 0 ? row[idxProjeto] : "";
          const descricao = idxDescricao >= 0 ? row[idxDescricao] : "";
          const statusOrc = idxStatusOrc >= 0 ? String(row[idxStatusOrc] || "").trim() : "";
          const statusPed = idxStatusPed >= 0 ? String(row[idxStatusPed] || "").trim() : "";
          let prazo = idxPrazo >= 0 ? row[idxPrazo] : "";
          prazo = normalizePrazo(prazo);
          let prazoProposta = idxPrazoProposta >= 0 ? row[idxPrazoProposta] : "";
          prazoProposta = normalizePrazo(prazoProposta);

          // Cards de orçamento: somente Rascunho; usa PRAZO_PROPOSTA (prazo da proposta), não PRAZO (prazo de entrega)
          const statusOrcNorm = statusOrc.toLowerCase();
          const soRascunho = statusOrcNorm === "rascunho";
          const excluidos = statusOrcNorm === "enviado" || statusOrcNorm === "convertido em pedido" || statusOrcNorm.indexOf("expirado") >= 0 || statusOrcNorm.indexOf("perdido") >= 0;
          const pedidoVazio = !statusPed || statusPed === "-";
          if (soRascunho && !excluidos && pedidoVazio) {
            data["Processo de Orçamento"].push({
              cliente: cliente,
              projeto: projeto,
              descricao: descricao,
              status: statusOrc,
              PRAZO_PROPOSTA: prazoProposta,
              prazoProposta: prazoProposta
            });
          }

          // Cards de pedido: STATUS_PEDIDO preenchido e não "-" (revertido para orçamento) nem Finalizado
          if (statusPed && statusPed !== "" && statusPed !== "-" && statusPed !== "Finalizado") {
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

            // Extrai temposReais, temNotaFiscal e valorNF do JSON_DADOS se existir
            let temposReais = {};
            let temNotaFiscal = false;
            let valorNF = "";
            if (jsonDados) {
              try {
                const parsed = JSON.parse(jsonDados);
                if (parsed && parsed.dados) {
                  if (parsed.dados.temposReais) temposReais = parsed.dados.temposReais;
                  const obs = parsed.dados.observacoes || {};
                  temNotaFiscal = !!obs.temNotaFiscal;
                  valorNF = (obs.valorNF != null && obs.valorNF !== "") ? String(obs.valorNF).trim() : "";
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
                prazo: prazo,
                temNotaFiscal: temNotaFiscal,
                valorNF: valorNF
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
  /** Quando o app é aberto no seu domínio (iframe com baseUrl), links ficam no seu domínio em vez de script.google.com */
  const baseUrl = (e && e.parameter && e.parameter.baseUrl) ? String(e.parameter.baseUrl).trim() : null;

  // Alias: PWA usa page=painelfinanceiro-dashboard; servidor usa dashboardfinanceiro
  if (page === 'painelfinanceiro-dashboard') page = 'dashboardfinanceiro';

  const paginasProtegidas = {
    'dashboard': ['admin', 'mod', 'usuario'],
    'formulario': ['admin', 'mod', 'usuario'],
    'materiais': ['admin', 'mod', 'usuario'],
    'geradoretiquetas': ['admin', 'mod', 'usuario'],
    'kanban': ['admin', 'mod', 'usuario'],
    'avaliacoes': ['admin'],
    'projetos': ['admin', 'mod', 'usuario'],
    'projetodetalhe': ['admin', 'mod', 'usuario'],
    'avaliacoespage': ['admin'],
    'pedidos': ['admin', 'mod'],
    'logs': ['admin', 'mod'],
    'manutencao': ['admin', 'mod', 'usuario'],
    'manu_registros': ['admin', 'mod', 'usuario'],
    'paginasprotegidas': ['admin'],
    'veiculos': ['admin', 'mod', 'usuario', 'visitante'],
    'veiculos_list': ['admin', 'mod', 'usuario', 'visitante'],
    'produtos': ['admin', 'mod', 'usuario'],
    'painelfinanceiro': ['admin', 'mod'],
    'dashboardfinanceiro': ['admin', 'mod']
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
      templateLogin.baseUrl = baseUrl;

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
        templateLoginDefault.baseUrl = baseUrl;

        return templateLoginDefault.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      case 'dashboard':
        const templateDashboard = HtmlService.createTemplateFromFile('dashboard');
        templateDashboard.token = token;
        templateDashboard.baseUrl = baseUrl;
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
        templateProdutos.baseUrl = baseUrl;
        return templateProdutos.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'projetos':
        const templateProjetos = HtmlService.createTemplateFromFile('projetos');
        templateProjetos.token = token;
        templateProjetos.baseUrl = baseUrl;
        return templateProjetos.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'projetodetalhe': {
        const templateProjetoDetalhe = HtmlService.createTemplateFromFile('projetodetalhe');
        templateProjetoDetalhe.token = token;
        templateProjetoDetalhe.linha = e?.parameter?.linha || '';
        templateProjetoDetalhe.baseUrl = baseUrl;
        return templateProjetoDetalhe.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      case 'pedidos': {
        const templatePedidos = HtmlService.createTemplateFromFile('pedidos');
        templatePedidos.token = token;
        templatePedidos.baseUrl = baseUrl;
        try {
          templatePedidos.pedidosData = JSON.stringify(getPedidos());
        } catch (err) {
          Logger.log('getPedidos ao servir página pedidos: ' + (err.message || err));
          templatePedidos.pedidosData = '[]';
        }
        return templatePedidos.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      case 'paginasprotegidas':
        const templatePaginasProtegidas = HtmlService.createTemplateFromFile('paginasprotegidas');
        templatePaginasProtegidas.token = token;
        templatePaginasProtegidas.baseUrl = baseUrl;
        return templatePaginasProtegidas.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'formulario': {
        const templateForm = HtmlService.createTemplateFromFile('formulario');
        templateForm.token = token;
        templateForm.baseUrl = baseUrl;
        const linhaParam = (e && e.parameter && e.parameter.linha) ? String(e.parameter.linha).trim() : '';
        templateForm.linhaUrl = linhaParam;
        templateForm.linhaUrlInicial = linhaParam;
        templateForm.config = { token: token, linha: linhaParam };
        return templateForm.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      case 'veiculos':
        const usuario = getUsuarioLogadoPorToken(token);
        const templateVeicForm = HtmlService.createTemplateFromFile('veiculos');
        templateVeicForm.token = token;
        templateVeicForm.baseUrl = baseUrl;
        templateVeicForm.usuario = usuario ? usuario.usuario : "Usuário não identificado";
        return templateVeicForm.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      
      case 'painelfinanceiro':
        const templatePainelFinanceiro = HtmlService.createTemplateFromFile('painelfinanceiro');
        templatePainelFinanceiro.token = token;
        templatePainelFinanceiro.baseUrl = baseUrl;
        return templatePainelFinanceiro.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'dashboardfinanceiro': {
        const templateDashboardFin = HtmlService.createTemplateFromFile('painelfinanceiro-dashboard');
        templateDashboardFin.token = token;
        templateDashboardFin.baseUrl = baseUrl;
        try {
          templateDashboardFin.pedidosData = JSON.stringify(getPedidos());
        } catch (err) {
          Logger.log('getPedidos ao servir dashboard financeiro: ' + (err.message || err));
          templateDashboardFin.pedidosData = '[]';
        }
        try {
          templateDashboardFin.clientesData = JSON.stringify(getTodosClientes());
        } catch (err) {
          templateDashboardFin.clientesData = '[]';
        }
        return templateDashboardFin.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      case 'manutencao':
        const templateManutencao = HtmlService.createTemplateFromFile('manutencao');
        templateManutencao.token = token;
        templateManutencao.baseUrl = baseUrl;
        return templateManutencao.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'avaliacoes':
        const templateAval = HtmlService.createTemplateFromFile('avaliacoes');
        templateAval.token = token;
        templateAval.baseUrl = baseUrl;
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
        templateAvalReg.baseUrl = baseUrl;
        return templateAvalReg.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'kanban': {
        const templateKanban = HtmlService.createTemplateFromFile('kanban');
        templateKanban.baseUrl = baseUrl;
        return templateKanban
          .evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }

      case 'veiculos_list':
        const templateVeiculosList = HtmlService.createTemplateFromFile('veiculos_list');
        templateVeiculosList.token = token;
        templateVeiculosList.baseUrl = baseUrl;
        return templateVeiculosList.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'geradoretiquetas':
        const templateEtiquetas = HtmlService.createTemplateFromFile('geradoretiquetas');
        templateEtiquetas.token = token;
        templateEtiquetas.baseUrl = baseUrl;
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
        templateEtiqTable.baseUrl = baseUrl;
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
        templateManuReg.baseUrl = baseUrl;
        return templateManuReg.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'seguranca':
        const templateSeguranca = HtmlService.createTemplateFromFile('seguranca');
        templateSeguranca.baseUrl = baseUrl;
        return templateSeguranca.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'apresentacao':
        const templateApresentacao = HtmlService.createTemplateFromFile('apresentacao');
        templateApresentacao.token = token;
        templateApresentacao.baseUrl = baseUrl;
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

  // Usuário que está editando/regenerando a etiqueta (não o criador original)
  let usuario = "Desconhecido";
  if (token) {
    usuario = PropertiesService.getScriptProperties().getProperty(token) || "Desconhecido";
    try { usuario = JSON.parse(usuario).usuario; } catch (e) { }
    usuario = usuario.replace(/([a-z])([A-Z])/g, '$1 $2');
    usuario = NOMES_COMPLETOS[usuario] || usuario;
  }

  // Acessa a planilha
  const materiais = SHEET_MAT;
  // Lê a linha atual para pegar os valores originais
  const rowIndex = dadosLinha.rowIndex;
  const linhaValores = materiais.getRange(rowIndex, 1, 1, materiais.getLastColumn()).getValues()[0];

  if (!token) usuario = linhaValores[2] || usuario; // fallback se token não vier

  // Monta objeto com os dados da etiqueta (usuario = quem está gerando a nova etiqueta)
  const dadosEtiqueta = {
    numeroChapa: linhaValores[0],  // Coluna A
    dataEntrada: Utilities.formatDate(new Date(linhaValores[1]), Session.getScriptTimeZone(), "dd/MM/yy"), // Coluna B
    usuario: usuario,   // Coluna C - quem editou/regenerou
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

  // Atualiza a coluna ETIQUETA e a coluna FEITO POR (quem editou/regenerou)
  materiais.getRange(rowIndex, 12).setValue(urlPdf); // Coluna L = ETIQUETA
  materiais.getRange(rowIndex, 3).setValue(usuario);  // Coluna C = FEITO POR (usuário que gerou a nova etiqueta)

  return { urlPdf, usuario };
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
      // Extrai temNotaFiscal, valorNF e dataEntrega do JSON_DADOS para exibição (botão NF, botão Entrega)
      const jsonDados = obj["JSON_DADOS"];
      if (jsonDados && typeof jsonDados === "string") {
        try {
          const parsed = JSON.parse(jsonDados);
          if (parsed && parsed.dados && parsed.dados.observacoes) {
            const obs = parsed.dados.observacoes;
            obj.temNotaFiscal = !!obs.temNotaFiscal;
            obj.valorNF = obs.valorNF != null && obs.valorNF !== "" ? String(obs.valorNF).trim() : "";
            obj.dataEntrega = obs.dataEntrega != null && obs.dataEntrega !== "" ? String(obs.dataEntrega).trim() : "";
          }
        } catch (e) { /* ignora */ }
      }
      return obj;
    });

    // Ordem: maior número de linha primeiro = último adicionado no topo. Novos projetos são inseridos ao final (appendRow), então sempre aparecem no topo na página. Não usa DATA, assim ao editar ou gerar PDF o projeto não pula para o topo.
    data.sort(function (a, b) {
      var linhaA = a._linhaPlanilha || 0;
      var linhaB = b._linhaPlanilha || 0;
      return linhaB - linhaA; // descendente: linha maior primeiro
    });
    Logger.log('getProjetos: Retornando %s projetos (última linha = topo; novos projetos no final da planilha aparecem no topo)', data.length);
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
 * Converte valor de moeda (número ou string BR/US) para número.
 * Evita erro quando a planilha retorna número (ex: 1174.76) e o código removia o ponto.
 */
function _parseCurrency(val) {
  if (val == null || val === "") return 0;
  if (typeof val === "number" && !isNaN(val)) return val;
  var s = (val + "").trim().replace(/[^\d,.]/g, "");
  if (!s) return 0;
  var lastComma = s.lastIndexOf(",");
  var lastDot = s.lastIndexOf(".");
  if (lastComma > lastDot) {
    return parseFloat(s.replace(/\./g, "").replace(",", ".")) || 0;
  }
  if (lastDot > lastComma) {
    return parseFloat(s.replace(/,/g, "")) || 0;
  }
  return parseFloat(s.replace(",", ".")) || 0;
}

/**
 * Calcula parcelas a partir do texto de pagamento (ex: "30/60/90", "Á Vista / 30 / 45", ou "70% no pedido, 30% na entrega").
 * dataBase = data de entrega (prioridade) ou data de competência para cálculo dos vencimentos.
 * Retorna array de { numero, dias, valor, dataVencimento, condicao? } ou null se à vista / parcela única.
 */
function _calcularParcelasPedidos(textoPagamento, valorTotal, dataBase) {
  if (!textoPagamento || !textoPagamento.trim()) return null;
  var texto = textoPagamento.trim().replace(/\s+/g, " ");
  var textoUpper = texto.toUpperCase();

  function parseBr(s) {
    if (!s) return null;
    var m = (s + "").trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (m) return new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    return null;
  }
  function formatBr(d) {
    if (!d || !d.getDate) return "";
    var dd = ("0" + d.getDate()).slice(-2);
    var mm = ("0" + (d.getMonth() + 1)).slice(-2);
    return dd + "/" + mm + "/" + d.getFullYear();
  }
  var dBase = parseBr(dataBase);

  // --- Percentual por condição (ex: 70% no pedido, 30% na entrega) ---
  var temPedidoOuEntrega = /pedido|entrega/i.test(texto);
  if (temPedidoOuEntrega && textoUpper.indexOf("%") >= 0) {
    var partes = texto.split(/(\d+\s*%)/);
    var percentuais = [];
    for (var i = 1; i < partes.length; i += 2) {
      var pctStr = partes[i] || "";
      var textoApos = (partes[i + 1] || "").toLowerCase();
      var pct = parseInt(pctStr.replace(/\D/g, ""), 10);
      if (isNaN(pct)) continue;
      var condicao = textoApos.indexOf("entrega") >= 0 ? "Na entrega" : "No pedido";
      percentuais.push({ pct: pct, condicao: condicao });
    }
    if (percentuais.length >= 1) {
      return percentuais.map(function (item, idx) {
        return {
          numero: idx + 1,
          dias: null,
          valor: valorTotal * (item.pct / 100),
          dataVencimento: item.condicao,
          condicao: item.condicao
        };
      });
    }
  }

  // À vista único (sem barra ou "30 dias" único): não gera parcelas
  if ((textoUpper.indexOf("VISTA") >= 0 && textoUpper.indexOf("/") < 0) || textoUpper === "30 DIAS") return null;
  if (!dBase) return null;
  // Suporte "Á Vista / 30 / 45": partes separadas por / ; "Á Vista" ou "A VISTA" = 0 dias, números = dias
  var dias = [];
  if (textoUpper.indexOf("/") >= 0) {
    var partes = textoUpper.split(/\s*\/\s*/);
    for (var i = 0; i < partes.length; i++) {
      var p = (partes[i] || "").trim();
      if (/vista/i.test(p)) dias.push(0);
      else {
        var num = parseInt(p.replace(/\D/g, ""), 10);
        if (!isNaN(num)) dias.push(num);
      }
    }
  }
  if (dias.length === 0) {
    var diasMatch = textoUpper.match(/\d+/g);
    if (!diasMatch || diasMatch.length === 0) return null;
    dias = diasMatch.map(function (x) { return parseInt(x, 10); });
  }
  if (dias.length === 0) return null;
  var numParcelas = dias.length;
  var valorParcela = valorTotal / numParcelas;
  return dias.map(function (dia, idx) {
    var d = new Date(dBase.getTime());
    d.setDate(d.getDate() + dia);
    return { numero: idx + 1, dias: dia, valor: valorParcela, dataVencimento: formatBr(d) };
  });
}

/**
 * Retorna apenas projetos com STATUS_ORCAMENTO = "Convertido em Pedido" (pedidos),
 * ordenados por data de competência (mais recentes primeiro).
 * Condições de pagamento e data de vencimento vêm da aba Pedidos (prioridade) ou JSON_DADOS.
 * @returns {Array} Lista de pedidos com _condicoesPagamento, _dataVencimento, _parcelas, _valorRestante, etc.
 */
function ensurePedidosSheet() {
  var sheet = ss.getSheetByName("Pedidos");
  if (!sheet) {
    sheet = ss.insertSheet("Pedidos");
    sheet.getRange(1, 1, 1, PEDIDOS_HEADERS.length).setValues([PEDIDOS_HEADERS]);
    SHEET_PED = sheet;
  } else {
    if (sheet.getLastRow() < 1) {
      sheet.getRange(1, 1, 1, PEDIDOS_HEADERS.length).setValues([PEDIDOS_HEADERS]);
    }
  }
  return sheet;
}

var PEDIDOS_HEADER_ALIASES = {
  "PROJETO": ["PROJETO", "Projeto", "Código", "Codigo", "CÓDIGO", "CODIGO", "Código do projeto"],
  "NF": ["NF", "N", "NOTA FISCAL", "NOTA_FISCAL"],
  "CONDICOES_PAGAMENTO": ["CONDICOES_PAGAMENTO", "CONDICOES PAGAMENTO", "CONDIÇÕES DE PAGAMENTO", "Condições de pagamento"],
  "DATA_ENTREGA": ["DATA_ENTREGA", "DATA ENTREGA", "DATA DE ENTREGA"],
  "DATA_VENCIMENTO": ["DATA_VENCIMENTO", "DATA VENCIMENTO", "DATA DE VENCIMENTO"],
  "VALOR_PAGO": ["VALOR_PAGO", "VALOR PAGO"],
  "STATUS_PAGAMENTO": ["STATUS_PAGAMENTO", "STATUS PAGAMENTO"],
  "PARCELAS_E_PGTOS": ["PARCELAS_E_PGTOS", "PARCELAS E PGTOS", "HISTORICO PAGAMENTOS", "PARCELAS PAGAS"],
  "NUMERO_SEQUENCIAL": ["NUMERO_SEQUENCIAL", "NUMERO SEQUENCIAL", "NÚMERO SEQUENCIAL", "Nº", "N"],
  "OBS": ["OBS", "OBSERVAÇÕES", "OBSERVACOES", "OBSERVACAO"],
  "DATA_COMPETENCIA": ["DATA_COMPETENCIA", "DATA COMPETENCIA", "DATA COMPETÊNCIA", "DATA DE COMPETENCIA"],
  "VALOR_TOTAL": ["VALOR_TOTAL", "VALOR TOTAL"],
  "LINHA_PROJETO": ["LINHA_PROJETO", "LINHA PROJETO"],
  "CLIENTE": ["CLIENTE", "Cliente"]
};

/**
 * Lê a aba Pedidos e retorna um mapa por PROJETO (código do projeto).
 * Aceita nomes alternativos de coluna (Projeto, Código, etc.) para planilhas já existentes.
 * @returns {Object} { "codigoProjeto": { rowIndex, NOTA_FISCAL, CONDICOES_PAGAMENTO, ... }, ... }
 */
function getPedidosSheetMap() {
  try {
    var sheet = ss.getSheetByName("Pedidos");
    if (!sheet || sheet.getLastRow() < 2) return {};
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var idx = {};
    function findCol(headerName) {
      var col = headers.indexOf(headerName);
      if (col >= 0) return col;
      var aliases = PEDIDOS_HEADER_ALIASES[headerName];
      if (aliases) {
        for (var a = 0; a < aliases.length; a++) {
          col = headers.indexOf(aliases[a]);
          if (col >= 0) return col;
        }
        for (var i = 0; i < headers.length; i++) {
          if (headers[i] && (headers[i] + "").toLowerCase().replace(/\s/g, " ") === (headerName + "").toLowerCase().replace(/_/g, " ")) return i;
        }
      }
      return -1;
    }
    PEDIDOS_HEADERS.forEach(function (h) {
      var col = findCol(h);
      if (col >= 0) idx[h] = col;
    });
    if (idx.PROJETO === undefined) {
      if (headers.length > 0 && (headers[0] != null && String(headers[0]).trim() !== "")) idx.PROJETO = 0;
      else return {};
    }
    var byProjeto = {};
    for (var r = 1; r < data.length; r++) {
      var row = data[r];
      var codigo = (row[idx.PROJETO] || "").toString().trim();
      if (!codigo) continue;
      if (byProjeto[codigo]) continue;
      var obj = { _linhaPedidos: r + 1 };
      if (idx.NF !== undefined) obj["NF"] = (row[idx.NF] != null && row[idx.NF] !== "") ? String(row[idx.NF]).trim() : "";
      if (idx.CONDICOES_PAGAMENTO !== undefined) obj["CONDICOES_PAGAMENTO"] = (row[idx.CONDICOES_PAGAMENTO] != null && row[idx.CONDICOES_PAGAMENTO] !== "") ? String(row[idx.CONDICOES_PAGAMENTO]).trim() : "";
      if (idx.DATA_ENTREGA !== undefined) obj["DATA_ENTREGA"] = (row[idx.DATA_ENTREGA] != null && row[idx.DATA_ENTREGA] !== "") ? String(row[idx.DATA_ENTREGA]).trim() : "";
      if (idx.DATA_VENCIMENTO !== undefined) obj["DATA_VENCIMENTO"] = (row[idx.DATA_VENCIMENTO] != null && row[idx.DATA_VENCIMENTO] !== "") ? String(row[idx.DATA_VENCIMENTO]).trim() : "";
      if (idx.VALOR_PAGO !== undefined) obj["VALOR_PAGO"] = row[idx.VALOR_PAGO];
      if (idx.STATUS_PAGAMENTO !== undefined) obj["STATUS_PAGAMENTO"] = (row[idx.STATUS_PAGAMENTO] != null && row[idx.STATUS_PAGAMENTO] !== "") ? String(row[idx.STATUS_PAGAMENTO]).trim() : "";
      if (idx.PARCELAS_E_PGTOS !== undefined) obj["PARCELAS_E_PGTOS"] = (row[idx.PARCELAS_E_PGTOS] != null && row[idx.PARCELAS_E_PGTOS] !== "") ? String(row[idx.PARCELAS_E_PGTOS]).trim() : "";
      if (idx.NUMERO_SEQUENCIAL !== undefined) obj["NUMERO_SEQUENCIAL"] = (row[idx.NUMERO_SEQUENCIAL] != null && row[idx.NUMERO_SEQUENCIAL] !== "") ? row[idx.NUMERO_SEQUENCIAL] : "";
      if (idx.OBS !== undefined) obj["OBS"] = (row[idx.OBS] != null && row[idx.OBS] !== "") ? String(row[idx.OBS]).trim() : "";
      if (idx.DATA_COMPETENCIA !== undefined) obj["DATA_COMPETENCIA"] = (row[idx.DATA_COMPETENCIA] != null && row[idx.DATA_COMPETENCIA] !== "") ? String(row[idx.DATA_COMPETENCIA]).trim() : "";
      if (idx.VALOR_TOTAL !== undefined) obj["VALOR_TOTAL"] = row[idx.VALOR_TOTAL];
      byProjeto[codigo] = obj;
    }
    return byProjeto;
  } catch (e) {
    Logger.log("getPedidosSheetMap error: " + e.message);
    return {};
  }
}

/**
 * Atualiza ou insere uma linha na aba Pedidos pelo código do projeto.
 * @param {string} projeto - Código do projeto
 * @param {Object} dadosAtualizacao - Campos a gravar (NOTA_FISCAL, CONDICOES_PAGAMENTO, DATA_ENTREGA, DATA_VENCIMENTO, VALOR_PAGO, STATUS_PAGAMENTO, PARCELAS_E_PGTOS, NUMERO_SEQUENCIAL, etc.)
 * @param {Object} [dadosProjeto] - Dados do projeto para preencher nova linha (CLIENTE, VALOR_TOTAL, DATA_COMPETENCIA, LINHA_PROJETO)
 * @returns {Object} { sucesso: boolean }
 */
function atualizarPedidoNaPlanilha(projeto, dadosAtualizacao, dadosProjeto) {
  try {
    projeto = (projeto || "").toString().trim();
    if (!projeto) throw new Error("Código do projeto é obrigatório.");
    var sheet = ensurePedidosSheet();
    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var colIdx = {};
    PEDIDOS_HEADERS.forEach(function (h) {
      var c = headers.indexOf(h);
      if (c >= 0) colIdx[h] = c;
    });
    for (var i = 0; i < PEDIDOS_HEADERS.length; i++) {
      if (colIdx[PEDIDOS_HEADERS[i]] === undefined) {
        var newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(PEDIDOS_HEADERS[i]);
        colIdx[PEDIDOS_HEADERS[i]] = newCol - 1;
      }
    }
    var rowNum = -1;
    for (var r = 1; r < data.length; r++) {
      if ((data[r][colIdx.PROJETO] || "").toString().trim() === projeto) {
        rowNum = r + 1;
        break;
      }
    }
    var campoParaColuna = {
      NOTA_FISCAL: "NF",
      CONDICOES_PAGAMENTO: "CONDICOES_PAGAMENTO",
      DATA_ENTREGA: "DATA_ENTREGA",
      DATA_VENCIMENTO: "DATA_VENCIMENTO",
      DATA_COMPETENCIA: "DATA_COMPETENCIA",
      VALOR_PAGO: "VALOR_PAGO",
      STATUS_PAGAMENTO: "STATUS_PAGAMENTO",
      HISTORICO_PAGAMENTOS: "PARCELAS_E_PGTOS",
      NUMERO_SEQUENCIAL: "NUMERO_SEQUENCIAL",
      OBS: "OBS"
    };
    var updates = {};
    for (var campo in dadosAtualizacao) {
      if (!Object.prototype.hasOwnProperty.call(dadosAtualizacao, campo)) continue;
      var colName = campoParaColuna[campo] || campo;
      if (colIdx[colName] !== undefined) updates[colName] = dadosAtualizacao[campo];
    }
    if (rowNum >= 2) {
      for (var col in updates) {
        var c = colIdx[col];
        if (c !== undefined) sheet.getRange(rowNum, c + 1).setValue(updates[col] == null ? "" : updates[col]);
      }
      if (dadosProjeto) {
        if (dadosProjeto.CLIENTE != null && colIdx.CLIENTE !== undefined) sheet.getRange(rowNum, colIdx.CLIENTE + 1).setValue(dadosProjeto.CLIENTE);
        if (dadosProjeto["VALOR TOTAL"] != null && colIdx.VALOR_TOTAL !== undefined) sheet.getRange(rowNum, colIdx.VALOR_TOTAL + 1).setValue(dadosProjeto["VALOR TOTAL"]);
        if (dadosProjeto._linhaPlanilha != null && colIdx.LINHA_PROJETO !== undefined) sheet.getRange(rowNum, colIdx.LINHA_PROJETO + 1).setValue(dadosProjeto._linhaPlanilha);
        // DATA_COMPETENCIA em linha existente não é sobrescrita por dadosProjeto (só pelo modal da página Pedidos via dadosAtualizacao)
      }
    } else {
      // Regra: só insere nova linha quando não existe nenhuma com este PROJETO (garante uma linha por código)
      var newRow = [];
      for (var j = 0; j < PEDIDOS_HEADERS.length; j++) {
        var h = PEDIDOS_HEADERS[j];
        var val = "";
        if (h === "PROJETO") val = projeto;
        else if (updates[h] !== undefined) val = updates[h];
        else if (dadosProjeto) {
          if (h === "CLIENTE") val = dadosProjeto.CLIENTE || "";
          if (h === "VALOR_TOTAL") val = dadosProjeto["VALOR TOTAL"] != null ? dadosProjeto["VALOR TOTAL"] : "";
          if (h === "DATA_COMPETENCIA") val = dadosProjeto._dataCompetencia || dadosProjeto.DATA || "";
          if (h === "LINHA_PROJETO") val = dadosProjeto._linhaPlanilha || "";
        }
        newRow.push(val == null ? "" : val);
      }
      sheet.appendRow(newRow);
    }

    // Sincronizar com a aba Projetos: alterações feitas na página/planilha de pedidos também atualizam a linha do projeto
    try {
      var linhaProj = null;
      if (dadosProjeto && dadosProjeto._linhaPlanilha != null) {
        var parsed = parseInt(dadosProjeto._linhaPlanilha, 10);
        if (!isNaN(parsed) && parsed >= 2) linhaProj = parsed;
      }
      if (linhaProj == null && projeto) {
        var sheetProj = ss.getSheetByName("Projetos");
        if (sheetProj) linhaProj = findRowByColumnValue(sheetProj, "PROJETO", projeto);
      }
      if (linhaProj >= 2) {
        // Edição veio da página de Pedidos: não criar nem escrever colunas na aba Projetos (evita criar Data competência, Status pagamento, etc.). Apenas atualizar JSON_DADOS.
        atualizarProjetoNaPlanilha(linhaProj, dadosAtualizacao, { apenasJsonDados: true });
      }
    } catch (syncErr) {
      Logger.log("Sync Pedidos->Projetos: " + syncErr.message);
    }
    return { sucesso: true };
  } catch (e) {
    Logger.log("atualizarPedidoNaPlanilha error: " + e.message);
    throw new Error("Erro ao atualizar pedido: " + (e.message || "erro desconhecido"));
  }
}

/**
 * Monta objeto de atualização para a aba Pedidos a partir dos dados de uma linha da aba Projetos.
 * Usado para backfill (projetos já convertidos antes da nova versão) e para ensurePedidoRow.
 * @param {Object} p - Objeto projeto (linha da aba Projetos)
 * @returns {Object} Objeto com NOTA_FISCAL, CONDICOES_PAGAMENTO, DATA_ENTREGA, etc.
 */
function updatesPedidoDesdeProjeto(p) {
  var updates = {};
  var j = p.JSON_DADOS || p["JSON_DADOS"] || "";
  var parsed = null;
  if (j && typeof j === "string" && j.trim()) {
    try {
      parsed = JSON.parse(j);
    } catch (e) {}
  }
  var dados = parsed && (parsed.dados || parsed) ? (parsed.dados || parsed) : {};
  var obs = dados.observacoes || {};

  var exigeNF = !!obs.temNotaFiscal;
  var nf = (p["NF"] || p["NOTA_FISCAL"] || p["NOTA FISCAL"] || "").toString().trim();
  if (!nf && (obs.valorNF != null && obs.valorNF !== "")) nf = String(obs.valorNF).trim();
  // Regra:
  // - Se o projeto NÃO exige NF (temNotaFiscal=false) e ainda não há NF definida, usar "SN" (Sem Nota) como valor padrão.
  // - Se o projeto EXIGE NF, não considerar "SN" como NF feita (mantém pendente).
  if (!exigeNF) {
    if (!nf) nf = "SN";
  } else if (nf && nf.toUpperCase() === "SN") {
    nf = "";
  }
  if (nf) updates.NOTA_FISCAL = nf;

  var cond = (p["CONDICOES_PAGAMENTO"] || p["CONDIÇÕES DE PAGAMENTO"] || p["CONDICOES DE PAGAMENTO"] || "").toString().trim();
  if (cond && (cond.indexOf("{") === 0 || cond.indexOf("[") === 0)) cond = "";
  if (!cond && (obs.pagamento != null && obs.pagamento !== "")) cond = String(obs.pagamento).trim();
  if (cond) updates.CONDICOES_PAGAMENTO = cond;

  var dataEnt = (p["DATA_ENTREGA"] || p["DATA DE ENTREGA"] || p["DATA ENTREGA"] || "").toString().trim();
  if (dataEnt) updates.DATA_ENTREGA = dataEnt;

  var dataVenc = (p["DATA_VENCIMENTO"] || p["DATA VENCIMENTO"] || p["DATA DE VENCIMENTO"] || "").toString().trim();
  if (dataVenc) updates.DATA_VENCIMENTO = dataVenc;

  var vp = p["VALOR_PAGO"] || p["VALOR PAGO"];
  if (vp != null && vp !== "") updates.VALOR_PAGO = vp;

  var st = (p["STATUS_PAGAMENTO"] || p["STATUS PAGAMENTO"] || p["STATUS DE PAGAMENTO"] || "").toString().trim();
  if (st) updates.STATUS_PAGAMENTO = st;

  var hist = (p["PARCELAS E PGTOS"] || p["HISTORICO_PAGAMENTOS"] || p["HISTORICO PAGAMENTOS"] || p["HISTÓRICO PAGAMENTOS"] || p["PARCELAS PAGAS"] || "").toString().trim();
  if (hist) updates.HISTORICO_PAGAMENTOS = hist;

  var numSeq = p["NUMERO_SEQUENCIAL"] || p["NÚMERO SEQUENCIAL"] || p["NUMERO SEQUENCIAL"] || p._numeroSequencial;
  if ((numSeq == null || String(numSeq).trim() === "") && parsed) {
    if (parsed.numeroSequencial != null && String(parsed.numeroSequencial).trim() !== "") numSeq = parsed.numeroSequencial;
    else if (dados.numeroSequencial != null && String(dados.numeroSequencial).trim() !== "") numSeq = dados.numeroSequencial;
  }
  if (numSeq != null && String(numSeq).trim() !== "") updates.NUMERO_SEQUENCIAL = numSeq;

  return updates;
}

/**
 * Migração única: copia para a aba Pedidos os dados que estão preenchidos na aba Projetos
 * (nas colunas de pedido) e estão vazios em Pedidos. Execute uma vez antes de excluir as colunas de pedido da Projetos.
 * Não sobrescreve células já preenchidas em Pedidos.
 * @returns {{ atualizados: number, semLinhaPedidos: number, erros: Array<string> }}
 */
function migrarDadosPedidoProjetosParaPedidos() {
  var resultado = { atualizados: 0, semLinhaPedidos: 0, erros: [] };
  try {
    var todos = getProjetos();
    if (!todos || !Array.isArray(todos)) return resultado;

    var pedidos = todos.filter(function (p) {
      var status = (p.STATUS_ORCAMENTO || p["STATUS ORCAMENTO"] || "").toString().trim();
      return status.toLowerCase() === "convertido em pedido";
    });

    var pedidosSheetMap = getPedidosSheetMap();
    // Mapa: chave em dadosAtualizacao (atualizarPedidoNaPlanilha) -> chave no rowPed (getPedidosSheetMap)
    var campoParaChavePedidos = {
      NOTA_FISCAL: "NF",
      CONDICOES_PAGAMENTO: "CONDICOES_PAGAMENTO",
      DATA_ENTREGA: "DATA_ENTREGA",
      DATA_VENCIMENTO: "DATA_VENCIMENTO",
      VALOR_PAGO: "VALOR_PAGO",
      STATUS_PAGAMENTO: "STATUS_PAGAMENTO",
      HISTORICO_PAGAMENTOS: "PARCELAS_E_PGTOS",
      NUMERO_SEQUENCIAL: "NUMERO_SEQUENCIAL",
      DATA_COMPETENCIA: "DATA_COMPETENCIA"
    };

    for (var i = 0; i < pedidos.length; i++) {
      var p = pedidos[i];
      var codigo = (p.PROJETO || "").toString().trim();
      if (!codigo) continue;

      var rowPed = pedidosSheetMap[codigo];
      if (!rowPed) {
        resultado.semLinhaPedidos++;
        ensurePedidoRow(p._linhaPlanilha);
        pedidosSheetMap = getPedidosSheetMap();
        rowPed = pedidosSheetMap[codigo];
        if (!rowPed) continue;
      }

      var updates = updatesPedidoDesdeProjeto(p);
      var dataComp = (p["DATA_COMPETENCIA"] || p["DATA COMPETÊNCIA"] || p.DATA || "").toString().trim();
      if (dataComp) updates.DATA_COMPETENCIA = dataComp;

      var toWrite = {};
      for (var campo in updates) {
        if (!Object.prototype.hasOwnProperty.call(updates, campo)) continue;
        var pedidosKey = campoParaChavePedidos[campo] || campo;
        var valorPedidos = rowPed[pedidosKey];
        var valorProjetos = updates[campo];
        var pedidosVazio = (valorPedidos == null || String(valorPedidos).trim() === "");
        var projetosPreenchido = (valorProjetos != null && String(valorProjetos).trim() !== "");
        if (pedidosVazio && projetosPreenchido) toWrite[campo] = valorProjetos;
      }

      if (Object.keys(toWrite).length > 0) {
        try {
          var dadosProjeto = {
            CLIENTE: p.CLIENTE,
            "VALOR TOTAL": p["VALOR TOTAL"],
            _dataCompetencia: (p["DATA_COMPETENCIA"] || p["DATA COMPETÊNCIA"] || p.DATA || "").toString().trim(),
            _linhaPlanilha: p._linhaPlanilha
          };
          atualizarPedidoNaPlanilha(codigo, toWrite, dadosProjeto);
          resultado.atualizados++;
        } catch (err) {
          resultado.erros.push(codigo + ": " + (err.message || err));
        }
      }
    }

    return resultado;
  } catch (e) {
    resultado.erros.push("migrarDadosPedidoProjetosParaPedidos: " + (e.message || e));
    Logger.log("migrarDadosPedidoProjetosParaPedidos error: " + e.message);
    return resultado;
  }
}

/**
 * Adiciona em lote na aba Pedidos todas as linhas dos projetos que ainda não têm registro.
 * Uma única leitura/escrita para evitar timeout e "um por vez".
 * @param {Array} pedidos - Lista de objetos projeto (status Convertido em Pedido)
 * @param {Object} pedidosSheetMap - Mapa por PROJETO já existente na aba Pedidos (de getPedidosSheetMap)
 */
function backfillPedidosEmLote(pedidos, pedidosSheetMap) {
  if (!pedidos || !pedidos.length) return;
  var sheet = ensurePedidosSheet();
  var rowsToAdd = [];
  for (var i = 0; i < pedidos.length; i++) {
    var p = pedidos[i];
    var codigoProjeto = (p.PROJETO || "").toString().trim();
    if (!codigoProjeto || pedidosSheetMap[codigoProjeto]) continue;
    var updates = updatesPedidoDesdeProjeto(p);
    var dataCompP = (p["DATA_COMPETENCIA"] || p["DATA COMPETÊNCIA"] || p.DATA || "").toString().trim();
    var newRow = [];
    for (var j = 0; j < PEDIDOS_HEADERS.length; j++) {
      var h = PEDIDOS_HEADERS[j];
      var val = "";
      if (h === "PROJETO") val = codigoProjeto;
      else if (h === "NF") val = updates.NOTA_FISCAL != null ? updates.NOTA_FISCAL : "";
      else if (h === "CONDICOES_PAGAMENTO") val = updates.CONDICOES_PAGAMENTO != null ? updates.CONDICOES_PAGAMENTO : "";
      else if (h === "DATA_ENTREGA") val = updates.DATA_ENTREGA != null ? updates.DATA_ENTREGA : "";
      else if (h === "DATA_VENCIMENTO") val = updates.DATA_VENCIMENTO != null ? updates.DATA_VENCIMENTO : "";
      else if (h === "VALOR_PAGO") val = updates.VALOR_PAGO != null ? updates.VALOR_PAGO : "";
      else if (h === "STATUS_PAGAMENTO") val = updates.STATUS_PAGAMENTO != null ? updates.STATUS_PAGAMENTO : "";
      else if (h === "PARCELAS_E_PGTOS") val = updates.HISTORICO_PAGAMENTOS != null ? updates.HISTORICO_PAGAMENTOS : "";
      else if (h === "NUMERO_SEQUENCIAL") val = updates.NUMERO_SEQUENCIAL != null ? updates.NUMERO_SEQUENCIAL : "";
      else if (h === "CLIENTE") val = p.CLIENTE || "";
      else if (h === "VALOR_TOTAL") val = p["VALOR TOTAL"] != null ? p["VALOR TOTAL"] : "";
      else if (h === "DATA_COMPETENCIA") val = dataCompP || "";
      else if (h === "LINHA_PROJETO") val = p._linhaPlanilha != null ? p._linhaPlanilha : "";
      newRow.push(val == null ? "" : val);
    }
    rowsToAdd.push(newRow);
  }
  if (rowsToAdd.length === 0) return;
  try {
    var startRow = sheet.getLastRow() + 1;
    var numRows = rowsToAdd.length;
    sheet.getRange(startRow, 1, startRow + numRows - 1, PEDIDOS_HEADERS.length).setValues(rowsToAdd);
  } catch (e) {
    Logger.log("backfillPedidosEmLote error: " + e.message);
  }
}

/**
 * Garante que existe uma linha na aba Pedidos para o projeto da linha indicada (Projetos).
 * Chamado ao converter orçamento em pedido. Se já existir dados de pedido na linha Projetos, copia para Pedidos.
 * Data de competência: só é preenchida quando a linha em Pedidos é NOVA (conversão). Se a linha já existe (ex.: novo PDF),
 * não sobrescreve DATA_COMPETENCIA (editável apenas pelo modal da página Pedidos).
 * @param {number} linhaProjetos - Número da linha na aba Projetos
 */
function ensurePedidoRow(linhaProjetos) {
  try {
    var sheetProj = ss.getSheetByName("Projetos");
    if (!sheetProj || sheetProj.getLastRow() < linhaProjetos) return;
    var headers = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
    var row = sheetProj.getRange(linhaProjetos, 1, 1, sheetProj.getLastColumn()).getValues()[0];
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    var codigo = (obj.PROJETO || "").toString().trim();
    if (!codigo) return;
    var dataComp = (obj["DATA_COMPETENCIA"] || obj["DATA COMPETÊNCIA"] || obj["DATA COMPETENCIA"] || obj.DATA || "").toString().trim();
    var valorTotal = (obj["VALOR TOTAL"] != null && obj["VALOR TOTAL"] !== "") ? obj["VALOR TOTAL"] : (obj["VALOR_TOTAL"] != null && obj["VALOR_TOTAL"] !== "") ? obj["VALOR_TOTAL"] : "";
    var linkPdf = (obj["LINK DO PDF"] != null && obj["LINK DO PDF"] !== "") ? obj["LINK DO PDF"] : (obj["LINK_DO_PDF"] != null && obj["LINK_DO_PDF"] !== "") ? obj["LINK_DO_PDF"] : "";
    var cliente = (obj.CLIENTE != null && obj.CLIENTE !== "") ? obj.CLIENTE : "";
    var updates = updatesPedidoDesdeProjeto(obj);
    var sheetPed = ensurePedidosSheet();
    var dataPed = sheetPed.getDataRange().getValues();
    var headersPed = dataPed[0];
    var colIdxProjeto = headersPed.indexOf("PROJETO");
    var pedidoJaExiste = false;
    if (colIdxProjeto >= 0) {
      for (var r = 1; r < dataPed.length; r++) {
        if ((dataPed[r][colIdxProjeto] || "").toString().trim() === codigo) {
          pedidoJaExiste = true;
          break;
        }
      }
    }
    var dadosProjeto = {
      CLIENTE: cliente,
      "VALOR TOTAL": valorTotal,
      _linhaPlanilha: linhaProjetos
    };
    if (!pedidoJaExiste && dataComp) dadosProjeto._dataCompetencia = dataComp;
    atualizarPedidoNaPlanilha(codigo, updates, dadosProjeto);
  } catch (e) {
    Logger.log("ensurePedidoRow error: " + e.message);
  }
}

function getPedidos() {
  try {
    function parseBr(s) {
      if (!s) return null;
      var str = (s || "").toString().trim();
      var m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) {
        var d = new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
        return isNaN(d.getTime()) ? null : d;
      }
      return null;
    }
    function formatBr(date) {
      if (!date || !date.getDate) return "";
      var dd = ("0" + date.getDate()).slice(-2);
      var mm = ("0" + (date.getMonth() + 1)).slice(-2);
      var yyyy = date.getFullYear();
      return dd + "/" + mm + "/" + yyyy;
    }
    function diasDasCondicoes(cond) {
      if (!cond) return 0;
      var c = (cond || "").toString().toLowerCase();
      if (c.indexOf("vista") >= 0) return 0;
      var nums = c.match(/\d+/g);
      if (!nums || nums.length === 0) return 0;
      var max = 0;
      for (var i = 0; i < nums.length; i++) {
        var n = parseInt(nums[i], 10);
        if (!isNaN(n) && n > max) max = n;
      }
      return max;
    }

    ensurePedidosSheet();
    var pedidosSheetMap = getPedidosSheetMap();
    var todos = getProjetos();
    var convertidos = (todos && Array.isArray(todos)) ? todos.filter(function (p) {
      var status = (p.STATUS_ORCAMENTO || p["STATUS ORCAMENTO"] || "").toString().trim();
      return status.toLowerCase() === "convertido em pedido";
    }) : [];
    backfillPedidosEmLote(convertidos, pedidosSheetMap);
    pedidosSheetMap = getPedidosSheetMap();

    var byProjetoProj = {};
    if (todos && Array.isArray(todos)) {
      for (var i = 0; i < todos.length; i++) {
        var c = (todos[i].PROJETO || "").toString().trim();
        if (c) byProjetoProj[c] = todos[i];
      }
    }

    var pedidos = [];
    for (var codigo in pedidosSheetMap) {
      var rowPed = pedidosSheetMap[codigo];
      var p = {
        PROJETO: codigo,
        _linhaPedidos: rowPed._linhaPedidos,
        "NF": rowPed["NF"],
        "CONDICOES_PAGAMENTO": rowPed["CONDICOES_PAGAMENTO"],
        "DATA_ENTREGA": rowPed["DATA_ENTREGA"],
        "DATA_VENCIMENTO": rowPed["DATA_VENCIMENTO"],
        "VALOR_PAGO": rowPed["VALOR_PAGO"],
        "STATUS_PAGAMENTO": rowPed["STATUS_PAGAMENTO"],
        "PARCELAS E PGTOS": rowPed["PARCELAS_E_PGTOS"],
        "NUMERO_SEQUENCIAL": rowPed["NUMERO_SEQUENCIAL"],
        "OBS": rowPed["OBS"],
        "DATA_COMPETENCIA": rowPed["DATA_COMPETENCIA"],
        "VALOR TOTAL": rowPed["VALOR_TOTAL"],
        CLIENTE: "",
        DATA: "",
        "LINK DO PDF": "",
        JSON_DADOS: ""
      };
      var proj = byProjetoProj[codigo];
      if (proj) {
        p.CLIENTE = proj.CLIENTE || "";
        if (p["VALOR TOTAL"] == null || p["VALOR TOTAL"] === "") p["VALOR TOTAL"] = proj["VALOR TOTAL"];
        p.DATA = proj.DATA || "";
        p._linhaPlanilha = proj._linhaPlanilha;
        p["LINK DO PDF"] = proj["LINK DO PDF"] || "";
        p.JSON_DADOS = proj.JSON_DADOS || proj["JSON_DADOS"] || "";
        p.PROCESSOS = proj.PROCESSOS || "";
      } else {
        p._linhaPlanilha = 50000 + (rowPed._linhaPedidos || 0);
      }
      pedidos.push(p);
    }

    // Regra: todo projeto Convertido em Pedido deve aparecer na página. Se a aba Pedidos não retornou linhas (planilha vazia, coluna diferente, etc.), montar lista a partir dos convertidos e garantir linha na aba.
    if (pedidos.length === 0 && convertidos.length > 0) {
      convertidos.forEach(function (proj) {
        var codigo = (proj.PROJETO || "").toString().trim();
        if (!codigo) return;
        var p = {
          PROJETO: codigo,
          _linhaPedidos: null,
          "NF": proj["NF"] || proj["NOTA_FISCAL"] || "",
          "CONDICOES_PAGAMENTO": proj["CONDICOES_PAGAMENTO"] || proj["CONDIÇÕES DE PAGAMENTO"] || "",
          "DATA_ENTREGA": proj["DATA_ENTREGA"] || proj["DATA DE ENTREGA"] || "",
          "DATA_VENCIMENTO": proj["DATA_VENCIMENTO"] || proj["DATA DE VENCIMENTO"] || "",
          "VALOR_PAGO": proj["VALOR_PAGO"] || proj["VALOR PAGO"],
          "STATUS_PAGAMENTO": proj["STATUS_PAGAMENTO"] || proj["STATUS PAGAMENTO"] || "Pendente",
          "PARCELAS E PGTOS": proj["PARCELAS E PGTOS"] || proj["HISTORICO_PAGAMENTOS"] || "",
          "NUMERO_SEQUENCIAL": proj["NUMERO_SEQUENCIAL"] || proj["NUMERO SEQUENCIAL"] || "",
          "OBS": proj["OBS"] || proj["OBSERVAÇÕES"] || "",
          "DATA_COMPETENCIA": proj["DATA_COMPETENCIA"] || proj["DATA COMPETÊNCIA"] || proj.DATA || "",
          "VALOR TOTAL": proj["VALOR TOTAL"] || proj["VALOR_TOTAL"],
          CLIENTE: proj.CLIENTE || "",
          DATA: proj.DATA || "",
          PROCESSOS: proj.PROCESSOS || "",
          "LINK DO PDF": proj["LINK DO PDF"] || "",
          JSON_DADOS: proj.JSON_DADOS || proj["JSON_DADOS"] || "",
          _linhaPlanilha: proj._linhaPlanilha
        };
        pedidos.push(p);
      });
    }

    pedidos.forEach(function (p) {
      var dataComp = (p["DATA_COMPETENCIA"] != null && String(p["DATA_COMPETENCIA"]).trim() !== "") ? String(p["DATA_COMPETENCIA"]).trim() : (p["DATA COMPETÊNCIA"] != null && String(p["DATA COMPETÊNCIA"]).trim() !== "" ? String(p["DATA COMPETÊNCIA"]).trim() : "");
      if (!dataComp) dataComp = (p.DATA != null && String(p.DATA).trim() !== "") ? String(p.DATA).trim() : "";
      p._dataCompetencia = dataComp;

      // Normalizar VALOR TOTAL e LINK DO PDF (planilha pode ter "Valor Total", "VALOR TOTAL", etc.)
      if ((p["VALOR TOTAL"] == null || p["VALOR TOTAL"] === "") && (p["Valor Total"] != null && p["Valor Total"] !== "")) p["VALOR TOTAL"] = p["Valor Total"];
      if ((p["VALOR TOTAL"] == null || p["VALOR TOTAL"] === "") && (p["VALOR_TOTAL"] != null && p["VALOR_TOTAL"] !== "")) p["VALOR TOTAL"] = p["VALOR_TOTAL"];
      if ((p["LINK DO PDF"] == null || p["LINK DO PDF"] === "") && (p["Link do PDF"] != null && p["Link do PDF"] !== "")) p["LINK DO PDF"] = p["Link do PDF"];
      if ((p["LINK DO PDF"] == null || p["LINK DO PDF"] === "") && (p["LINK_DO_PDF"] != null && p["LINK_DO_PDF"] !== "")) p["LINK DO PDF"] = p["LINK_DO_PDF"];

      // Condições de pagamento e número sequencial: prioridade aba Pedidos, depois JSON_DADOS
      var condicoes = (p["CONDICOES_PAGAMENTO"] || p["CONDIÇÕES DE PAGAMENTO"] || "").toString().trim();
      var numSeqCol = p["NUMERO_SEQUENCIAL"] || p["NÚMERO SEQUENCIAL"] || p["NUMERO SEQUENCIAL"];
      p._numeroSequencial = (numSeqCol != null && String(numSeqCol).trim() !== "") ? numSeqCol : null;
      try {
        var jsonDados = p.JSON_DADOS || p["JSON_DADOS"] || "";
        if (jsonDados) {
          var parsed = JSON.parse(jsonDados);
          if (p._numeroSequencial == null && parsed.numeroSequencial != null) p._numeroSequencial = parsed.numeroSequencial;
          var dados = parsed.dados || parsed;
          var obs = dados.observacoes || {};
          var pag = (obs.pagamento || "").toString().trim();
          if (pag && !condicoes) condicoes = pag;
        }
      } catch (e) {}
      p._condicoesPagamento = condicoes;

      // Data de vencimento: a partir da DATA DE ENTREGA (não mais competência). À vista = data de entrega; parcelado = parcelas a partir da entrega.
      var dataEntrega = (p["DATA_ENTREGA"] || p["DATA DE ENTREGA"] || p.PRAZO || "").toString().trim();
      var dataVenc = (p["DATA_VENCIMENTO"] || p["DATA VENCIMENTO"] || "").toString().trim();
      var dataBaseVenc = dataEntrega || dataComp;
      if (!dataVenc && dataBaseVenc) {
        var valorTotal = _parseCurrency(p["VALOR TOTAL"]);
        var texto = (condicoes || "").toLowerCase();
        if (texto.indexOf("vista") >= 0 && texto.indexOf("/") < 0) {
          dataVenc = dataEntrega || dataComp;
          p._dataVencimentoOrdenar = dataVenc;
        } else {
          var parcelas = _calcularParcelasPedidos(condicoes, valorTotal, dataBaseVenc);
          if (parcelas && parcelas.length > 0) {
            p._parcelas = parcelas;
            var todasDatas = parcelas.map(function (parc) { return parc.dataVencimento; });
            dataVenc = todasDatas.join(", ");
            p._dataVencimentoOrdenar = (parcelas[0].condicao ? dataBaseVenc : parcelas[0].dataVencimento) || dataBaseVenc;
          } else {
            var dias = diasDasCondicoes(condicoes);
            var d = parseBr(dataBaseVenc);
            if (d) {
              d.setDate(d.getDate() + dias);
              dataVenc = formatBr(d);
            }
            p._dataVencimentoOrdenar = dataVenc;
          }
        }
      }
      p._dataVencimento = dataVenc || dataComp;
      if (p._dataVencimento && !p._dataVencimentoOrdenar) {
        var first = (p._dataVencimento + "").split(",")[0];
        p._dataVencimentoOrdenar = first ? first.trim() : p._dataVencimento;
      }

      // Histórico de parcelas/pagamentos: nova coluna "PARCELAS E PGTOS"; mantém leitura dos nomes antigos
      var historicoStr = (p["PARCELAS E PGTOS"] || p["HISTORICO_PAGAMENTOS"] || p["HISTORICO PAGAMENTOS"] || p["HISTÓRICO PAGAMENTOS"] || p["PARCELAS PAGAS"] || "").toString().trim();
      p._historicoPagamentos = [];
      if (historicoStr) {
        try {
          var arr = JSON.parse(historicoStr);
          if (Array.isArray(arr)) {
            p._historicoPagamentos = arr.map(function (x) {
              return { data: x.data || "", valor: _parseCurrency(x.valor), parcela: x.parcela != null ? x.parcela : null };
            });
          }
        } catch (e) {}
      }

      // Valor restante: se Pago = null (mostrar "Pago"); senão VALOR TOTAL - VALOR_PAGO
      var statusPag = (p["STATUS_PAGAMENTO"] || p["STATUS PAGAMENTO"] || p["STATUS DE PAGAMENTO"] || "Pendente").toString().trim();
      var valorPago = _parseCurrency(p["VALOR_PAGO"] || p["VALOR PAGO"]);
      var vtNum = _parseCurrency(p["VALOR TOTAL"]);
      if (statusPag === "Pago") {
        p._valorRestante = null;
      } else {
        p._valorRestante = Math.max(0, vtNum - valorPago);
      }
    });

    // Mesma ordem da página Projetos: por DATA do projeto decrescente, depois por linha (maior = mais recente)
    var parseBrNum = function (s) {
      if (!s) return 0;
      var m = (s + "").trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
      if (m) return parseInt(m[3] + ("0" + m[2]).slice(-2) + ("0" + m[1]).slice(-2), 10);
      return 0;
    };
    pedidos.sort(function (a, b) {
      var dataA = (a.DATA != null ? a.DATA : a._dataCompetencia || "").toString().trim();
      var dataB = (b.DATA != null ? b.DATA : b._dataCompetencia || "").toString().trim();
      var ta = parseBrNum(dataA);
      var tb = parseBrNum(dataB);
      if (ta !== tb) return tb - ta;
      return (b._linhaPlanilha || 0) - (a._linhaPlanilha || 0);
    });

    return pedidos;
  } catch (e) {
    Logger.log("getPedidos error: " + e.message);
    return [];
  }
}

/**
 * Retorna a URL da primeira pasta do Drive cujo nome começa com o código do projeto
 * (ex: 260204cMS -> pasta "260204cMS PED Cliente..." ou "260204cMS COT ...").
 * Usa a mesma lógica de detectarPastaExistente (estrutura PROJ ou COM por data YYMMDD).
 * @param {string} codigoProjeto - Código do projeto (ex: 260204cMS)
 * @returns {string|null} URL da pasta ou null se não encontrada
 */
function getUrlPastaProjeto(codigoProjeto) {
  if (!codigoProjeto || codigoProjeto.length < 6) return null;
  try {
    var data = codigoProjeto.substring(0, 6);
    var info = detectarPastaExistente(codigoProjeto, data);
    if (info && info.pasta) return info.pasta.getUrl();
    return null;
  } catch (e) {
    Logger.log("getUrlPastaProjeto error: " + e.message);
    return null;
  }
}

/**
 * Atualiza dados de um projeto na planilha
 * @param {number} linha - Número da linha na planilha
 * @param {Object} dadosAtualizacao - Objeto com campos a atualizar
 * @param {Object} [opcoes] - { apenasJsonDados: boolean } quando true (ex.: edição pela página Pedidos), não cria nem escreve colunas na aba Projetos
 * @returns {Object} - {sucesso: boolean}
 */
function atualizarProjetoNaPlanilha(linha, dadosAtualizacao, opcoes) {
  try {
    var rowNum = parseInt(linha, 10);
    if (isNaN(rowNum) || rowNum < 2) {
      throw new Error("Linha da planilha inválida: " + linha);
    }
    linha = rowNum;

    const sheetProj = ss.getSheetByName("Projetos");
    if (!sheetProj) {
      throw new Error("Aba 'Projetos' não encontrada");
    }
    if (sheetProj.getLastRow() < linha) {
      throw new Error("Linha " + linha + " não existe na planilha (última linha: " + sheetProj.getLastRow() + ")");
    }

    var apenasJsonDados = opcoes && opcoes.apenasJsonDados === true;
    var statusOrcAntes = ""; // usado mais abaixo; quando apenasJsonDados=true não é alterado (evita conversão)

    // Busca cabeçalhos (só necessário se formos criar/escrever colunas)
    const headers = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
    const headerMap = {};
    headers.forEach((h, idx) => {
      if (h) headerMap[h.toString().trim().toUpperCase()] = idx;
    });
    // Aliases para colunas da página Pedidos (nomes possíveis na planilha)
    const aliasesPedidos = {
      "DATA_COMPETENCIA": ["DATA COMPETENCIA", "DATA COMPETÊNCIA", "DATA DE COMPETENCIA", "DATA DE COMPETÊNCIA"],
      "CONDICOES_PAGAMENTO": ["CONDICOES PAGAMENTO", "CONDIÇÕES DE PAGAMENTO", "CONDICOES DE PAGAMENTO"],
      "DATA_ENTREGA": ["DATA ENTREGA", "DATA DE ENTREGA"],
      "DATA_VENCIMENTO": ["DATA VENCIMENTO", "DATA DE VENCIMENTO"],
      "STATUS_PAGAMENTO": ["STATUS PAGAMENTO", "STATUS DE PAGAMENTO"],
      "VALOR_PAGO": ["VALOR PAGO", "VALOR PAGO"],
      "NUMERO_SEQUENCIAL": ["NUMERO SEQUENCIAL", "NÚMERO SEQUENCIAL", "Nº"],
      "HISTORICO_PAGAMENTOS": ["PARCELAS E PGTOS", "HISTORICO PAGAMENTOS", "HISTÓRICO PAGAMENTOS", "PARCELAS PAGAS"],
      "NOTA_FISCAL": ["NF", "NOTA FISCAL", "NOTA_FISCAL"]
    };
    // Nome do cabeçalho a criar quando a coluna não existir (aba Projetos pode ter só 14 colunas iniciais)
    const headerNameForCampo = {
      "DATA_COMPETENCIA": "DATA COMPETÊNCIA",
      "CONDICOES_PAGAMENTO": "CONDIÇÕES DE PAGAMENTO",
      "DATA_ENTREGA": "DATA ENTREGA",
      "DATA_VENCIMENTO": "DATA VENCIMENTO",
      "STATUS_PAGAMENTO": "STATUS PAGAMENTO",
      "VALOR_PAGO": "VALOR PAGO",
      "NUMERO_SEQUENCIAL": "NÚMERO SEQUENCIAL",
      "HISTORICO_PAGAMENTOS": "PARCELAS E PGTOS",
      "NOTA_FISCAL": "NF",
      "PRAZO_PROPOSTA": "PRAZO PROPOSTA"
    };
    function acharColuna(campo) {
      var key = campo.toUpperCase();
      if (headerMap[key] !== undefined) return headerMap[key];
      // Planilha pode ter cabeçalho com espaço (ex: "STATUS ORCAMENTO") em vez de underscore
      var keyComEspaco = key.replace(/_/g, " ");
      if (headerMap[keyComEspaco] !== undefined) return headerMap[keyComEspaco];
      var aliases = aliasesPedidos[campo];
      if (aliases) {
        for (var a = 0; a < aliases.length; a++) {
          var up = (aliases[a] || "").toString().toUpperCase();
          if (headerMap[up] !== undefined) return headerMap[up];
        }
      }
      // Fallback: planilha pode ter "Status Orçamento" (com acento) — usa mesma normalização do resto do código
      if (campo === "STATUS_ORCAMENTO" || campo === "STATUS_PEDIDO") {
        var idx = _findHeaderIndex(headers, campo);
        if (idx >= 0) return idx;
      }
      return undefined;
    }
    function valorParaPlanilha(campo, valor) {
      var v = (valor === null || valor === undefined) ? '' : valor;
      if (campo === 'DATA_ENTREGA' || campo === 'DATA_VENCIMENTO') {
        if (v && typeof v === 'string' && v.indexOf('-') >= 0) {
          var parts = v.split('-');
          if (parts.length === 3) v = parts[2] + '/' + parts[1] + '/' + parts[0];
        }
      }
      return v;
    }

    if (!apenasJsonDados) {
      // 1) Garantir que todas as colunas existem (criar cabeçalho quando não existir). Não criar colunas que são apenas da aba Pedidos.
      for (var campo in dadosAtualizacao) {
        if (!Object.prototype.hasOwnProperty.call(dadosAtualizacao, campo)) continue;
        if (CAMPOS_APENAS_PEDIDOS[campo]) continue;
        if (acharColuna(campo) !== undefined) continue;
        if (headerNameForCampo[campo]) {
          var newCol = sheetProj.getLastColumn() + 1;
          var headerName = headerNameForCampo[campo];
          sheetProj.getRange(1, newCol).setValue(headerName);
          headerMap[headerName.toString().trim().toUpperCase()] = newCol - 1;
        }
      }
      // Status de orçamento ANTES da escrita (para saber se é conversão nova ou edição de pedido existente)
      var rowAntes = sheetProj.getRange(linha, 1, linha, sheetProj.getLastColumn()).getValues()[0];
      var idxStatusOrcCol = headerMap["STATUS_ORCAMENTO"] !== undefined ? headerMap["STATUS_ORCAMENTO"] : headerMap["STATUS ORCAMENTO"];
      if (idxStatusOrcCol === undefined) { var idxOrc = _findHeaderIndex(headers, "STATUS_ORCAMENTO"); idxStatusOrcCol = idxOrc >= 0 ? idxOrc : undefined; }
      statusOrcAntes = (idxStatusOrcCol !== undefined && rowAntes[idxStatusOrcCol] !== undefined) ? String(rowAntes[idxStatusOrcCol]).trim() : "";
      // 2) Montar mapa colIdx -> valor e escrever na linha. Não gravar na Projetos os campos que são apenas da aba Pedidos (ficam em JSON_DADOS e na aba Pedidos).
      var colToVal = {};
      for (var campo2 in dadosAtualizacao) {
        if (!Object.prototype.hasOwnProperty.call(dadosAtualizacao, campo2)) continue;
        if (CAMPOS_APENAS_PEDIDOS[campo2]) continue;
        var colIdx = acharColuna(campo2);
        if (colIdx !== undefined) {
          colToVal[colIdx] = valorParaPlanilha(campo2, dadosAtualizacao[campo2]);
        }
      }
      // Escrever apenas as células que estão sendo atualizadas (não sobrescrever o intervalo inteiro para não apagar VALOR TOTAL, LINK DO PDF, etc.)
      for (var colIdx in colToVal) {
        if (!Object.prototype.hasOwnProperty.call(colToVal, colIdx)) continue;
        sheetProj.getRange(linha, parseInt(colIdx, 10) + 1).setValue(colToVal[colIdx] != null ? colToVal[colIdx] : '');
      }
    }

    // Atualizar JSON_DADOS para o formulário refletir alterações (condições de pagamento, número sequencial, NF, etc.)
    try {
      var headersAtual = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
      var idxJson = _findHeaderIndexProjetos(headersAtual, "JSON_DADOS");
      if (idxJson >= 0) {
        var jsonCell = sheetProj.getRange(linha, idxJson + 1).getValue();
        var parsed = null;
        if (jsonCell && typeof jsonCell === "string" && jsonCell.trim()) {
          try {
            parsed = JSON.parse(jsonCell);
          } catch (e) {}
        }
        if (!parsed) parsed = { dados: {}, numeroSequencial: null };
        if (!parsed.dados) parsed.dados = {};
        if (!parsed.dados.observacoes) parsed.dados.observacoes = {};
        var alterado = false;
        if (dadosAtualizacao.CONDICOES_PAGAMENTO !== undefined) {
          parsed.dados.observacoes.pagamento = dadosAtualizacao.CONDICOES_PAGAMENTO;
          alterado = true;
        }
        if (dadosAtualizacao.NUMERO_SEQUENCIAL !== undefined) {
          parsed.numeroSequencial = dadosAtualizacao.NUMERO_SEQUENCIAL;
          parsed.dados.numeroSequencial = dadosAtualizacao.NUMERO_SEQUENCIAL;
          alterado = true;
        }
        if (dadosAtualizacao.NOTA_FISCAL !== undefined) {
          parsed.dados.observacoes.valorNF = dadosAtualizacao.NOTA_FISCAL;
          alterado = true;
        }
        if (alterado) {
          sheetProj.getRange(linha, idxJson + 1).setValue(JSON.stringify(parsed));
        }
      }
    } catch (eJson) {
      Logger.log("Aviso ao atualizar JSON_DADOS em atualizarProjetoNaPlanilha: " + (eJson && eJson.message));
    }

    // Se a chamada veio da página de Pedidos (apenasJsonDados = true), não alterar colunas da aba Projetos
    // nem disparar sincronização Projetos->Pedidos ou lógica de conversão. Apenas o JSON_DADOS foi atualizado.
    if (apenasJsonDados) {
      return { sucesso: true };
    }

    // Sincronizar com aba Pedidos: se este projeto é pedido (Convertido em Pedido), refletir alterações na aba Pedidos
    try {
      var rowData = sheetProj.getRange(linha, 1, 1, sheetProj.getLastColumn()).getValues()[0];
      var idxStatusOrcRead = headerMap["STATUS_ORCAMENTO"] !== undefined ? headerMap["STATUS_ORCAMENTO"] : headerMap["STATUS ORCAMENTO"];
      if (idxStatusOrcRead === undefined) idxStatusOrcRead = _findHeaderIndex(headers, "STATUS_ORCAMENTO");
      var statusOrc = (idxStatusOrcRead >= 0 && rowData[idxStatusOrcRead] !== undefined ? rowData[idxStatusOrcRead] : "").toString().trim();
      if (statusOrc === "Convertido em Pedido") {
        var codigoProjeto = (rowData[headerMap["PROJETO"]] || "").toString().trim();
        if (codigoProjeto) {
          var syncUpdates = {};
          if (dadosAtualizacao.CONDICOES_PAGAMENTO !== undefined) syncUpdates.CONDICOES_PAGAMENTO = dadosAtualizacao.CONDICOES_PAGAMENTO;
          if (dadosAtualizacao.DATA_ENTREGA !== undefined) syncUpdates.DATA_ENTREGA = dadosAtualizacao.DATA_ENTREGA;
          if (dadosAtualizacao.DATA_VENCIMENTO !== undefined) syncUpdates.DATA_VENCIMENTO = dadosAtualizacao.DATA_VENCIMENTO;
          if (dadosAtualizacao.VALOR_PAGO !== undefined) syncUpdates.VALOR_PAGO = dadosAtualizacao.VALOR_PAGO;
          if (dadosAtualizacao.STATUS_PAGAMENTO !== undefined) syncUpdates.STATUS_PAGAMENTO = dadosAtualizacao.STATUS_PAGAMENTO;
          if (dadosAtualizacao.HISTORICO_PAGAMENTOS !== undefined) syncUpdates.HISTORICO_PAGAMENTOS = dadosAtualizacao.HISTORICO_PAGAMENTOS;
          if (dadosAtualizacao.NOTA_FISCAL !== undefined) syncUpdates.NOTA_FISCAL = dadosAtualizacao.NOTA_FISCAL;
          if (dadosAtualizacao.NUMERO_SEQUENCIAL !== undefined) syncUpdates.NUMERO_SEQUENCIAL = dadosAtualizacao.NUMERO_SEQUENCIAL;
          if (Object.keys(syncUpdates).length > 0) {
            var projObj = {};
            for (var i = 0; i < headers.length && i < rowData.length; i++) projObj[headers[i]] = rowData[i];
            atualizarPedidoNaPlanilha(codigoProjeto, syncUpdates, { CLIENTE: projObj.CLIENTE, "VALOR TOTAL": projObj["VALOR TOTAL"], _dataCompetencia: (projObj["DATA COMPETÊNCIA"] || projObj["DATA COMPETENCIA"] || projObj.DATA || ""), _linhaPlanilha: linha });
          }
        }
      }
    } catch (syncErr) {
      Logger.log("Sync Projetos->Pedidos: " + syncErr.message);
    }

    // SEMPRE que o projeto está como Convertido em Pedido: garantir pasta PED (qualquer circunstância: select, modal, formulário, Kanban)
    (function () {
      try {
        var rowAtual = sheetProj.getRange(linha, 1, 1, sheetProj.getLastColumn()).getValues()[0];
        var idxOrc = _findHeaderIndex(headers, "STATUS_ORCAMENTO");
        var statusAtual = (idxOrc >= 0 && rowAtual[idxOrc] !== undefined) ? String(rowAtual[idxOrc]).trim() : "";
        if (statusAtual !== "Convertido em Pedido") return;
        var idxProjeto = _findHeaderIndex(headers, "PROJETO");
        var idxCliente = _findHeaderIndex(headers, "CLIENTE");
        var idxDescricao = _findHeaderIndex(headers, "DESCRIÇÃO");
        if (idxProjeto < 0 || !rowAtual[idxProjeto]) return;
        var codigoProjeto = String(rowAtual[idxProjeto]).trim();
        var codigoBase = codigoProjeto.replace(/_v\d+$/i, "").trim();
        var cliente = idxCliente >= 0 ? String(rowAtual[idxCliente] || "").trim() : "";
        var descricao = idxDescricao >= 0 ? String(rowAtual[idxDescricao] || "").trim() : "";
        var dataProj = codigoBase.length >= 6 ? codigoBase.substring(0, 6) : "";
        if (dataProj) atualizarPrefixoPastaParaPedido(codigoBase, dataProj, cliente, descricao);
      } catch (ePasta) {
        Logger.log("Aviso ao garantir pasta PED (Convertido em Pedido): " + (ePasta && ePasta.message));
      }
    })();

    // Só faz conversão (data competência, status Prep MP, ensurePedidoRow, ordem de produção) quando é conversão NOVA (antes não era pedido)
    var ehConversaoNova = (dadosAtualizacao.STATUS_ORCAMENTO === "Convertido em Pedido" && statusOrcAntes !== "Convertido em Pedido");
    if (ehConversaoNova) {
      try {
        const hoje = new Date();
        const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || "America/Sao_Paulo";
        const dataCompetenciaStr = Utilities.formatDate(hoje, tz, "dd/MM/yyyy");
        // Todo projeto convertido em pedido tem primeiro status = Preparação MP / CAD / CAM (coluna STATUS_PEDIDO permanece na Projetos)
        var idxStatusPedido = acharColuna("STATUS_PEDIDO");
        if (idxStatusPedido !== undefined) {
          sheetProj.getRange(linha, idxStatusPedido + 1).setValue("Processo de Preparação MP / CAD / CAM");
        }
        ensurePedidoRow(linha);
        // DATA_COMPETENCIA (data da conversão) grava apenas na aba Pedidos, não na Projetos
        var rowConv = sheetProj.getRange(linha, 1, linha, sheetProj.getLastColumn()).getValues()[0];
        var headersConv = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
        var idxProjConv = _findHeaderIndexProjetos(headersConv, "PROJETO");
        var codigoConv = (idxProjConv >= 0 && rowConv[idxProjConv]) ? String(rowConv[idxProjConv]).trim() : "";
        if (codigoConv) {
          var idxClienteConv = _findHeaderIndexProjetos(headersConv, "CLIENTE");
          var idxValorConv = _findHeaderIndexProjetos(headersConv, "VALOR TOTAL");
          atualizarPedidoNaPlanilha(codigoConv, { DATA_COMPETENCIA: dataCompetenciaStr }, {
            CLIENTE: (idxClienteConv >= 0 && rowConv[idxClienteConv] != null) ? rowConv[idxClienteConv] : "",
            "VALOR TOTAL": (idxValorConv >= 0 && rowConv[idxValorConv] != null) ? rowConv[idxValorConv] : "",
            _dataCompetencia: dataCompetenciaStr,
            _linhaPlanilha: linha
          });
        }

        // Pasta já garantida acima no bloco "SEMPRE que status = Convertido em Pedido". Gerar Ordem de Produção ao converter em pedido
        try {
          const resultOp = gerarPdfOrdemProducao(linha);
          if (resultOp && resultOp.url) {
            Logger.log("Ordem de Produção gerada na conversão: " + resultOp.url);
          }
        } catch (errOp) {
          Logger.log("Aviso: não foi possível gerar Ordem de Produção na conversão: " + (errOp.message || errOp));
        }
      } catch (e) {
        Logger.log("⚠️ Erro ao renomear pasta de COT para PED: " + e.message);
      }
    }

    // Se DESCRIÇÃO ou CLIENTE foi alterado: atualizar JSON_DADOS (para formulário ver a mesma descrição) e renomear pasta no Drive
    if (dadosAtualizacao["DESCRIÇÃO"] !== undefined || dadosAtualizacao["CLIENTE"] !== undefined) {
      try {
        const numCols = sheetProj.getLastColumn();
        const rowData = sheetProj.getRange(linha, 1, 1, numCols).getValues()[0];
        const idxProjeto = headerMap["PROJETO"];
        const idxCliente = headerMap["CLIENTE"];
        const idxDescricao = headerMap["DESCRIÇÃO"] || headerMap["DESCRICAO"];
        var idxStatusOrc = headerMap["STATUS_ORCAMENTO"] || headerMap["STATUS ORCAMENTO"];
        if (idxStatusOrc === undefined) { var iOrc = _findHeaderIndex(headers, "STATUS_ORCAMENTO"); idxStatusOrc = iOrc >= 0 ? iOrc : undefined; }
        const idxJson = headerMap["JSON_DADOS"];
        const codigoProjeto = (idxProjeto !== undefined && rowData[idxProjeto]) ? String(rowData[idxProjeto]).trim() : "";
        const codigoBase = codigoProjeto ? codigoProjeto.replace(/_v\d+$/i, "").trim() : "";
        const cliente = (idxCliente !== undefined && rowData[idxCliente]) ? String(rowData[idxCliente]).trim() : "";
        const descricao = (idxDescricao !== undefined && rowData[idxDescricao]) ? String(rowData[idxDescricao] || "").trim() : "";
        const statusOrc = (idxStatusOrc !== undefined && rowData[idxStatusOrc]) ? String(rowData[idxStatusOrc] || "").trim() : "";
        const isPedido = statusOrc === "Convertido em Pedido";
        const dataProj = codigoBase.length >= 6 ? codigoBase.substring(0, 6) : (codigoProjeto.length >= 6 ? codigoProjeto.substring(0, 6) : "");

        if (idxJson !== undefined && rowData[idxJson]) {
          try {
            var parsed = JSON.parse(rowData[idxJson]);
            if (parsed.dados) {
              if (parsed.dados.observacoes == null) parsed.dados.observacoes = {};
              parsed.dados.observacoes.descricao = descricao;
              if (parsed.dados.cliente == null) parsed.dados.cliente = {};
              parsed.dados.cliente.nome = cliente;
              sheetProj.getRange(linha, idxJson + 1).setValue(JSON.stringify(parsed));
            }
          } catch (jsonErr) {
            Logger.log("Atualizar JSON_DADOS (descrição/cliente): " + jsonErr.message);
          }
        }

        if (codigoBase && dataProj) {
          renomearPastaProjeto(codigoBase, dataProj, cliente, descricao, isPedido);
        }
      } catch (e) {
        Logger.log("Erro ao sincronizar descrição/cliente (JSON e pasta): " + e.message);
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
 * Informa o valor da Nota Fiscal para um projeto que possui NF. Atualiza JSON_DADOS na aba Projetos
 * e sincroniza o valor para a aba Pedidos (coluna NOTA_FISCAL / NF).
 * @param {number} linha - Número da linha na aba Projetos
 * @param {string|number} valorNF - Valor da NF informado pelo usuário
 */
function informarValorNFProjeto(linha, valorNF) {
  try {
    linha = Number(linha);
    if (!linha || linha < 2) throw new Error("Linha inválida.");
    var sheetProj = ss.getSheetByName("Projetos");
    if (!sheetProj || sheetProj.getLastRow() < linha) throw new Error("Projeto não encontrado.");
    var headers = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
    var idxJson = headers.indexOf("JSON_DADOS");
    var idxProjeto = headers.indexOf("PROJETO");
    if (idxJson < 0 || idxProjeto < 0) throw new Error("Estrutura da planilha Projetos não encontrada.");
    var row = sheetProj.getRange(linha, 1, linha, headers.length).getValues()[0];
    var codigoProjeto = (row[idxProjeto] || "").toString().trim();
    if (!codigoProjeto) throw new Error("Código do projeto não encontrado.");
    var jsonStr = row[idxJson];
    var parsed = jsonStr ? (function () { try { return JSON.parse(jsonStr); } catch (e) { return null; } })() : null;
    if (!parsed || !parsed.dados) parsed = { dados: {} };
    if (!parsed.dados.observacoes) parsed.dados.observacoes = {};
    parsed.dados.observacoes.valorNF = valorNF == null ? "" : String(valorNF).trim();
    var newJson = JSON.stringify(parsed);
    sheetProj.getRange(linha, idxJson + 1).setValue(newJson);
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    obj["JSON_DADOS"] = newJson;
    var dataComp = (obj["DATA_COMPETENCIA"] || obj["DATA COMPETÊNCIA"] || obj.DATA || "").toString().trim();
    atualizarPedidoNaPlanilha(codigoProjeto, { NOTA_FISCAL: parsed.dados.observacoes.valorNF }, {
      CLIENTE: obj.CLIENTE,
      "VALOR TOTAL": obj["VALOR TOTAL"],
      _dataCompetencia: dataComp,
      _linhaPlanilha: linha
    });
    return { success: true };
  } catch (e) {
    Logger.log("informarValorNFProjeto error: " + e.message);
    throw new Error(e.message || "Erro ao informar valor da NF");
  }
}

/**
 * Informa a data de entrega de um projeto/pedido e sincroniza com a aba de Pedidos.
 * @param {number} linha - Linha na aba Projetos
 * @param {string} dataEntrega - Data de entrega (string)
 * @returns {{ success: boolean }}
 */
function informarDataEntregaProjeto(linha, dataEntrega) {
  try {
    linha = Number(linha);
    if (!linha || linha < 2) throw new Error("Linha inválida.");
    var sheetProj = ss.getSheetByName("Projetos");
    if (!sheetProj || sheetProj.getLastRow() < linha) throw new Error("Projeto não encontrado.");
    var headers = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
    var idxJson = headers.indexOf("JSON_DADOS");
    var idxProjeto = headers.indexOf("PROJETO");
    if (idxJson < 0 || idxProjeto < 0) throw new Error("Estrutura da planilha Projetos não encontrada.");
    var row = sheetProj.getRange(linha, 1, linha, headers.length).getValues()[0];
    var codigoProjeto = (row[idxProjeto] || "").toString().trim();
    if (!codigoProjeto) throw new Error("Código do projeto não encontrado.");
    var jsonStr = row[idxJson];
    var parsed = jsonStr ? (function () { try { return JSON.parse(jsonStr); } catch (e) { return null; } })() : null;
    if (!parsed || !parsed.dados) parsed = { dados: {} };
    if (!parsed.dados.observacoes) parsed.dados.observacoes = {};
    parsed.dados.observacoes.dataEntrega = dataEntrega == null ? "" : String(dataEntrega).trim();
    var newJson = JSON.stringify(parsed);
    sheetProj.getRange(linha, idxJson + 1).setValue(newJson);
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    obj["JSON_DADOS"] = newJson;
    var dataComp = (obj["DATA_COMPETENCIA"] || obj["DATA COMPETÊNCIA"] || obj.DATA || "").toString().trim();
    atualizarPedidoNaPlanilha(codigoProjeto, { DATA_ENTREGA: parsed.dados.observacoes.dataEntrega }, {
      CLIENTE: obj.CLIENTE,
      "VALOR TOTAL": obj["VALOR TOTAL"],
      _dataCompetencia: dataComp,
      _linhaPlanilha: linha
    });
    return { success: true };
  } catch (e) {
    Logger.log("informarDataEntregaProjeto error: " + e.message);
    throw new Error(e.message || "Erro ao informar data de entrega");
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
    // Regra: nunca duplicar código de projeto — se já existir linha com este PROJETO, atualizar em vez de inserir
    const codigoProjetoTrim = String(projeto.PROJETO || '').trim();
    var linhaExistenteProj = codigoProjetoTrim ? findRowByColumnValue(targetSheet, "PROJETO", codigoProjetoTrim) : null;

    const dadosObj = {
      CLIENTE: projeto.CLIENTE || '',
      'DESCRIÇÃO': projeto['DESCRIÇÃO'] || '',
      'RESPONSÁVEL CLIENTE': projeto['RESPONSÁVEL CLIENTE'] || '',
      PROJETO: projeto.PROJETO || '',
      'VALOR TOTAL': projeto['VALOR TOTAL'] || '',
      DATA: projeto.DATA || new Date().toLocaleDateString('pt-BR'),
      PROCESSOS: projeto.PROCESSOS || '',
      'LINK DO PDF': projeto['LINK DO PDF'] || '',
      'LINK DA MEMÓRIA DE CÁLCULO': projeto['LINK DA MEMÓRIA DE CÁLCULO'] || '',
      STATUS_ORCAMENTO: projeto.STATUS_ORCAMENTO || 'RASCUNHO',
      STATUS_PEDIDO: projeto.STATUS_PEDIDO !== undefined ? projeto.STATUS_PEDIDO : '',
      PRAZO: projeto.PRAZO || '',
      PRAZO_PROPOSTA: projeto.PRAZO_PROPOSTA || projeto['PRAZO PROPOSTA'] || '',
      OBSERVAÇÕES: projeto['OBSERVAÇÕES'] || '',
      JSON_DADOS: projeto.JSON_DADOS || ''
    };

    if (sheetProj) {
      if (linhaExistenteProj) {
        _escreverLinhaProjetosPorCabecalho(targetSheet, linhaExistenteProj, dadosObj, true);
        Logger.log('adicionarNovoProjetoNaPlanilha: Projeto já existia — linha atualizada (sem duplicar): ' + projeto.PROJETO);
      } else {
        _escreverLinhaProjetosPorCabecalho(targetSheet, targetSheet.getLastRow() + 1, dadosObj, false);
        Logger.log('adicionarNovoProjetoNaPlanilha: Projeto adicionado com sucesso na aba Projetos');
      }

      // Cria a pasta do projeto no Drive (estrutura COT)
      const codigoProjeto = String(projeto.PROJETO || '').trim();
      const descricao = String(projeto['DESCRIÇÃO'] || '').trim();
      const nomeCliente = String(projeto.CLIENTE || '').trim();
      if (codigoProjeto.length >= 6 && descricao) {
        try {
          const dataProj = codigoProjeto.substring(0, 6); // YYMMDD
          criarPastaOrcamento(codigoProjeto, descricao, dataProj, nomeCliente, false);
          Logger.log('adicionarNovoProjetoNaPlanilha: Pasta do projeto criada no Drive');
        } catch (errPasta) {
          Logger.log('adicionarNovoProjetoNaPlanilha: Aviso ao criar pasta no Drive: ' + (errPasta.message || errPasta));
          // Não falha a operação: projeto já foi salvo na planilha
        }
      }
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

          // Voltar para Processo de Orçamento: orçamento = Rascunho, pedido = "-"
          if (novoStatus === "Processo de Orçamento") {
            if (idxStatusOrc >= 0) sheetProj.getRange(i + 1, idxStatusOrc + 1).setValue("Rascunho");
            sheetProj.getRange(i + 1, idxStatusPed + 1).setValue("-");
            break;
          }

          // Se estava em orçamento e está mudando para um status de pedido, atualiza STATUS_ORCAMENTO também
          if (!statusAntigo && idxStatusOrc >= 0) {
            const statusOrc = String(row[idxStatusOrc] || '').trim();
            if (statusOrc !== "Convertido em Pedido") {
              sheetProj.getRange(i + 1, idxStatusOrc + 1).setValue("Convertido em Pedido");
              const codigoBase = String(valProjeto).replace(/_v\d+$/i, "").trim();
              const dataProj = codigoBase.length >= 6 ? codigoBase.substring(0, 6) : valProjeto.substring(0, 6);
              try {
                atualizarPrefixoPastaParaPedido(codigoBase, dataProj, valCliente, descricao);
                Logger.log("Pasta convertida de COT para PED (Kanban): " + codigoBase);
              } catch (e) {
                Logger.log("Erro ao renomear pasta de COT para PED: " + e.message);
              }
              ensurePedidoRow(i + 1);
            }
          }

          // Atualiza STATUS_PEDIDO: ao converter em pedido o primeiro status é sempre Preparação MP / CAD / CAM
          const statusPedidoFinal = !statusAntigo ? "Processo de Preparação MP / CAD / CAM" : novoStatus;
          sheetProj.getRange(i + 1, idxStatusPed + 1).setValue(statusPedidoFinal);
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
// LINK DO PDF, LINK DA MEMÓRIA DE CÁLCULO, STATUS_ORCAMENTO, STATUS_PEDIDO, PRAZO, PRAZO_PROPOSTA, OBSERVAÇÕES, JSON_DADOS
const PROJETOS_NUM_COLUNAS = 15;

/**
 * Parseia o código do projeto no formato YYMMDD + índice (a-z) + iniciais.
 * Ex: 251125dMS = ano 25, mês 11, dia 25, índice d (4º do dia da pessoa MS).
 * 260102cBR = ano 26, mês 01, dia 02, índice c, iniciais BR.
 * @param {string} codigo - Código do projeto (ex: 251125dMS, 260102cBR)
 * @returns {{ dateNum: number, indexNum: number, initials: string }} para ordenação
 */
function _parseCodigoProjetoParaOrdenacao(codigo) {
  var s = (codigo || '').toString().trim().replace(/_v\d+$/i, '');
  var m = s.match(/^(\d{6})([a-zA-Z])(.*)$/);
  if (!m) return { dateNum: 0, indexNum: 0, initials: s };
  var dateStr = m[1];
  var yy = parseInt(dateStr.substring(0, 2), 10);
  var mm = parseInt(dateStr.substring(2, 4), 10);
  var dd = parseInt(dateStr.substring(4, 6), 10);
  var year = yy >= 0 && yy <= 99 ? 2000 + yy : yy;
  var dateNum = year * 10000 + mm * 100 + dd;
  var indexLetter = (m[2] || 'a').toLowerCase();
  var indexNum = indexLetter.charCodeAt(0) - 97;
  if (indexNum < 0 || indexNum > 25) indexNum = 0;
  var initials = (m[3] || '').trim().toUpperCase();
  return { dateNum: dateNum, indexNum: indexNum, initials: initials };
}

/**
 * Ordena a aba Projetos pelo código do projeto (YYMMDD + índice + iniciais).
 * Ordem na planilha: mais antigos primeiro (data crescente), depois índice (a, b, c...), depois iniciais.
 * Assim a linha 2 fica com o projeto mais antigo e a última linha com o mais recente. Novos projetos
 * são adicionados ao final (appendRow); a página de Projetos ordena por última linha primeiro, então
 * os existentes ficam em ordem por data e todo novo projeto aparece no topo.
 */
function ordenarPlanilhaProjetosPorCodigo() {
  var sheet = ss.getSheetByName("Projetos");
  if (!sheet) {
    Logger.log("ordenarPlanilhaProjetosPorCodigo: Aba Projetos não encontrada");
    throw new Error("Aba 'Projetos' não encontrada.");
  }
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) {
    Logger.log("ordenarPlanilhaProjetosPorCodigo: Nenhum dado para ordenar");
    return;
  }
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0];
  var dataRows = values.slice(1);
  var idxProjeto = -1;
  for (var h = 0; h < headers.length; h++) {
    var hv = (headers[h] || '').toString().trim().toUpperCase();
    if (hv === 'PROJETO' || hv === 'NÚMERO DO PROJETO' || hv === 'NUMERO DO PROJETO' || hv === 'CÓDIGO' || hv === 'CODIGO') {
      idxProjeto = h;
      break;
    }
  }
  if (idxProjeto < 0) {
    Logger.log("ordenarPlanilhaProjetosPorCodigo: Coluna PROJETO não encontrada. Headers: " + headers.join(', '));
    throw new Error("Coluna do código do projeto (PROJETO) não encontrada na aba Projetos.");
  }
  dataRows.sort(function (rowA, rowB) {
    var codA = (rowA[idxProjeto] != null ? rowA[idxProjeto] : '').toString().trim();
    var codB = (rowB[idxProjeto] != null ? rowB[idxProjeto] : '').toString().trim();
    var pa = _parseCodigoProjetoParaOrdenacao(codA);
    var pb = _parseCodigoProjetoParaOrdenacao(codB);
    if (pa.dateNum !== pb.dateNum) return pa.dateNum - pb.dateNum; // mais antigos primeiro na planilha
    if (pa.indexNum !== pb.indexNum) return pa.indexNum - pb.indexNum;
    return (pa.initials || '').localeCompare(pb.initials || '');
  });
  var newValues = [headers].concat(dataRows);
  sheet.getRange(1, 1, lastRow, lastCol).setValues(newValues);
  Logger.log("ordenarPlanilhaProjetosPorCodigo: Planilha ordenada por código (mais antigos primeiro). Novos projetos no final aparecem no topo da página.");
}

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
        const nomeAbreviado = (dados.cliente && dados.cliente.nomeAbreviado) || "";
        criarPastaOrcamento(codigoProjeto, descricao, dataProj, clienteNome, false, nomeAbreviado);
      } catch (e) {
        Logger.log("Aviso ao criar pasta para rascunho: " + e.message);
      }
    }

    // Data formatada para exibição
    const agora = new Date();
    const dataBrasil = formatarDataBrasil(agora);

    if (dados.produtosCadastrados && Array.isArray(dados.produtosCadastrados)) {
      atribuirCodigosPRDAutomaticos(dados.produtosCadastrados);
    }

    const dadosJson = JSON.stringify({
      nome: nomeRascunho,
      dataSalvo: agora.toISOString(),
      dados: dados
    });

    const sheetProj = ss.getSheetByName("Projetos");
    const targetSheet = sheetProj;

    if (!targetSheet) throw new Error("Nenhuma aba de projetos encontrada");

    var linhaExistente;
    const observacoesKanban = (dados.observacoes && dados.observacoes.observacoesKanban != null) ? String(dados.observacoes.observacoesKanban).trim() : "";
    const prazoProposta = (dados.observacoes && dados.observacoes.prazoProposta != null) ? String(dados.observacoes.prazoProposta).trim() : "";
    const processos = (dados.observacoes && dados.observacoes.processos != null) ? String(dados.observacoes.processos).trim() : "";
    linhaExistente = sheetProj ? findRowByColumnValue(sheetProj, "PROJETO", codigoProjeto) : 0;
    var dadosObj = {
      "CLIENTE": clienteNome,
      "DESCRIÇÃO": descricao,
      "RESPONSÁVEL CLIENTE": clienteResponsavel,
      "PROJETO": codigoProjeto,
      "VALOR TOTAL": "",
      "DATA": dataBrasil,
      "PROCESSOS": processos,
      "LINK DO PDF": "",
      "LINK DA MEMÓRIA DE CÁLCULO": "",
      "STATUS_ORCAMENTO": "RASCUNHO",
      "STATUS_PEDIDO": "",
      "PRAZO": prazo,
      "PRAZO_PROPOSTA": prazoProposta,
      "OBSERVAÇÕES": observacoesKanban,
      "JSON_DADOS": dadosJson
    };
    if (linhaExistente) {
      _escreverLinhaProjetosPorCabecalho(targetSheet, linhaExistente, dadosObj, true);
    } else {
      _escreverLinhaProjetosPorCabecalho(targetSheet, targetSheet.getLastRow() + 1, dadosObj, false);
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

    // Lê a linha por cabeçalhos (evita valor/status errados quando a planilha tem colunas em ordem diferente)
    const numCols = targetSheet.getLastColumn();
    const headers = targetSheet.getRange(1, 1, 1, numCols).getValues()[0];
    const rowData = targetSheet.getRange(linha, 1, linha, numCols).getValues()[0];
    const rowObj = {};
    headers.forEach(function (h, i) { rowObj[h] = rowData[i]; });
    function val(obj, keys) {
      for (var k = 0; k < keys.length; k++) if (obj[keys[k]] != null && String(obj[keys[k]]).trim() !== "") return obj[keys[k]];
      return "";
    }

    // Preserva o status atual (por nome de coluna)
    const statusAtual = (val(rowObj, ["STATUS_ORCAMENTO", "STATUS ORCAMENTO"]) || "RASCUNHO").toString().trim();

    // Recalcula o valor total a partir dos dados do formulário (para refletir alterações de preço, etc.)
    let valorTotal = val(rowObj, ["VALOR TOTAL", "VALOR_TOTAL"]);
    try {
      const preview = calcularPreviewOrcamento(dados);
      if (preview && typeof preview.total === "number") {
        valorTotal = preview.total;
      }
    } catch (e) {
      Logger.log("Aviso: não foi possível recalcular total na atualização: " + e.message);
    }

    // PROCESSOS, links e status pedido: por nome de coluna
    const processos = (dados.observacoes && dados.observacoes.processos != null) ? String(dados.observacoes.processos).trim() : (val(rowObj, ["PROCESSOS"]) || "");
    const linkPdf = val(rowObj, ["LINK DO PDF"]);
    const linkMemoria = val(rowObj, ["LINK DA MEMÓRIA DE CÁLCULO", "LINK DA MEMORIA DE CALCULO"]);
    const statusPedido = val(rowObj, ["STATUS_PEDIDO", "STATUS PEDIDO"]);

    // Extrai dados relevantes do formulário para atualizar (OBSERVAÇÕES vem do campo Kanban/Projetos)
    const clienteNome = (dados.cliente && dados.cliente.nome) || "";
    const descricao = (dados.observacoes && dados.observacoes.descricao) || "";
    const prazo = (dados.observacoes && dados.observacoes.prazo) || "";
    const prazoProposta = (dados.observacoes && dados.observacoes.prazoProposta != null) ? String(dados.observacoes.prazoProposta).trim() : "";
    const observacoes = (dados.observacoes && dados.observacoes.observacoesKanban != null) ? String(dados.observacoes.observacoesKanban).trim() : (val(rowObj, ["OBSERVAÇÕES", "OBSERVACOES"]) || "");
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
        const nomeAbreviado = (dados.cliente && dados.cliente.nomeAbreviado) || "";
        criarPastaOrcamento(codigoProjeto, descricao, dataProj, clienteNome, false, nomeAbreviado);
      } catch (e) {
        Logger.log("Aviso ao criar pasta para atualização de rascunho: " + e.message);
      }
    }

    // Data formatada para exibição (não alterar DATA se projeto já é pedido: só editável via modal Pedidos)
    const agora = new Date();
    const dataBrasil = formatarDataBrasil(agora);

    if (dados.produtosCadastrados && Array.isArray(dados.produtosCadastrados)) {
      atribuirCodigosPRDAutomaticos(dados.produtosCadastrados);
    }

    // Preservar numeroSequencial ao atualizar: usar do formulário ou do JSON já salvo na linha (evita perder ao "Atualizar versão atual")
    var numeroSequencialSalvar = (dados.numeroSequencial != null && String(dados.numeroSequencial).trim() !== "") ? dados.numeroSequencial : null;
    if (numeroSequencialSalvar == null) {
      const jsonIdx = _findHeaderIndex(headers, "JSON_DADOS");
      if (jsonIdx >= 0 && rowData[jsonIdx] != null && String(rowData[jsonIdx]).trim() !== "") {
        try {
          const parsed = JSON.parse(String(rowData[jsonIdx]).trim());
          numeroSequencialSalvar = (parsed && parsed.numeroSequencial != null) ? parsed.numeroSequencial : ((parsed && parsed.dados && parsed.dados.numeroSequencial != null) ? parsed.dados.numeroSequencial : null);
        } catch (e) {}
      }
    }
    if (numeroSequencialSalvar != null) dados.numeroSequencial = numeroSequencialSalvar;
    const dadosJson = JSON.stringify({
      nome: codigoProjeto,
      dataSalvo: agora.toISOString(),
      numeroSequencial: numeroSequencialSalvar,
      dados: dados
    });

    var dataParaGravar = dataBrasil;
    if (statusAtual === "Convertido em Pedido") {
      var idxData = _findHeaderIndexProjetos(headers, "DATA");
      if (idxData >= 0 && rowData[idxData] != null && String(rowData[idxData]).trim() !== "") dataParaGravar = rowData[idxData];
    }

    // Atualiza por nome de coluna para não gravar JSON_DADOS em coluna errada (ex.: Condições de pagamento)
    var dadosObj = {
      "CLIENTE": clienteNome,
      "DESCRIÇÃO": descricao,
      "RESPONSÁVEL CLIENTE": clienteResponsavel,
      "PROJETO": codigoProjeto,
      "VALOR TOTAL": valorTotal,
      "DATA": dataParaGravar,
      "PROCESSOS": processos,
      "LINK DO PDF": linkPdf,
      "LINK DA MEMÓRIA DE CÁLCULO": linkMemoria,
      "STATUS_ORCAMENTO": statusAtual,
      "STATUS_PEDIDO": statusPedido,
      "PRAZO": prazo,
      "PRAZO_PROPOSTA": prazoProposta,
      "OBSERVAÇÕES": observacoes,
      "JSON_DADOS": dadosJson
    };
    _escreverLinhaProjetosPorCabecalho(targetSheet, linha, dadosObj, true);

    // Sincronizar com aba Pedidos quando o projeto já é pedido (condições de pagamento, número sequencial, etc.)
    if (statusAtual === "Convertido em Pedido" && codigoProjeto) {
      try {
        ensurePedidoRow(linha);
      } catch (errSync) {
        Logger.log("Aviso ao sincronizar Projetos→Pedidos em atualizarRascunho: " + (errSync && errSync.message));
      }
    }

    // Não gera PDF aqui ao salvar rascunho: evita travar em "Atualizando..." (timeout). Para gerar/atualizar PDF use "Finalizar / Gerar PDF".

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
      const nomeAbreviado = (cliente.nomeAbreviado || "").trim();
      criarPastaOrcamento(codigoProjeto, descricao, dataProj, clienteNome, true, nomeAbreviado);
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

    if (dados.produtosCadastrados && Array.isArray(dados.produtosCadastrados)) {
      atribuirCodigosPRDAutomaticos(dados.produtosCadastrados);
    }

    const agora = new Date();
    const dataBrasil = formatarDataBrasil(agora);
    const dadosJson = JSON.stringify({
      nome: codigoProjeto,
      dataSalvo: agora.toISOString(),
      dados: dados
    });

    var dadosObjPedido = {
      "CLIENTE": clienteNome,
      "DESCRIÇÃO": descricao,
      "RESPONSÁVEL CLIENTE": clienteResponsavel,
      "PROJETO": codigoProjeto,
      "VALOR TOTAL": valorTotal,
      "DATA": dataBrasil,
      "PROCESSOS": "",
      "LINK DO PDF": "",
      "LINK DA MEMÓRIA DE CÁLCULO": "",
      "STATUS_ORCAMENTO": "Convertido em Pedido",
      "STATUS_PEDIDO": "Processo de Preparação MP / CAD / CAM",
      "PRAZO": prazo,
      "PRAZO_PROPOSTA": "",
      "OBSERVAÇÕES": "",
      "JSON_DADOS": dadosJson
    };
    // Regra: nunca duplicar código de projeto — se já existir linha com este PROJETO, atualizar em vez de inserir
    var codigoPedido = String((dados.observacoes && dados.observacoes.projeto) || "").trim();
    var linhaExistentePed = codigoPedido ? findRowByColumnValue(sheetProj, "PROJETO", codigoPedido) : null;
    if (linhaExistentePed) {
      _escreverLinhaProjetosPorCabecalho(sheetProj, linhaExistentePed, dadosObjPedido, true);
    } else {
      _escreverLinhaProjetosPorCabecalho(sheetProj, sheetProj.getLastRow() + 1, dadosObjPedido, false);
    }
    // Gerar Ordem de Produção ao criar projeto como pedido
    try {
      const linhaParaOp = linhaExistentePed || sheetProj.getLastRow();
      const resultOp = gerarPdfOrdemProducao(linhaParaOp);
      if (resultOp && resultOp.url) {
        Logger.log("Ordem de Produção gerada ao salvar como pedido: " + resultOp.url);
      }
    } catch (errOp) {
      Logger.log("Aviso: não foi possível gerar Ordem de Produção ao salvar como pedido: " + (errOp.message || errOp));
    }
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

    // Lê a linha da planilha (todas as colunas) para não depender da ordem fixa (evitar ler JSON_DADOS da coluna Condições de pagamento)
    const numCols = Math.max(sheetProj ? PROJETOS_NUM_COLUNAS : ORCAMENTOS_NUM_COLUNAS, targetSheet.getLastColumn());
    const rowData = targetSheet.getRange(linha, 1, linha, numCols).getValues()[0];
    const headers = targetSheet.getRange(1, 1, 1, numCols).getValues()[0];

    // STATUS está no índice 9 em ambas estruturas (STATUS ou STATUS_ORCAMENTO)
    const status = rowData[9];

    // JSON_DADOS: usar coluna por nome, nunca índice fixo (a planilha pode ter Condições de pagamento etc.)
    const jsonIdx = _findHeaderIndex(headers, "JSON_DADOS");
    const dadosJson = (jsonIdx >= 0 && rowData[jsonIdx] != null && String(rowData[jsonIdx]).trim() !== "") ? String(rowData[jsonIdx]).trim() : "";

    // Índices por nome de coluna (planilha pode ter colunas em ordem diferente)
    const idxClienteCol = _findHeaderIndexProjetos(headers, "CLIENTE");
    const idxDescricaoCol = _findHeaderIndexProjetos(headers, "DESCRIÇÃO");
    const idxProjetoCol = _findHeaderIndexProjetos(headers, "PROJETO");
    const idxDataCol = _findHeaderIndexProjetos(headers, "DATA");

    // Se tiver JSON_DADOS, usa os dados completos do formulário
    if (dadosJson) {
      try {
        const dadosParsed = JSON.parse(dadosJson);
        // Incluir numeroSequencial nos dados retornados
        const dadosRetorno = dadosParsed.dados;
        dadosRetorno.numeroSequencial = dadosParsed.numeroSequencial || null;
        // Planilha é fonte da verdade para DESCRIÇÃO, CLIENTE, PROJETO e OBSERVAÇÕES (Kanban/Projetos)
        if (dadosRetorno.observacoes == null) dadosRetorno.observacoes = {};
        const descricaoPlanilha = (idxDescricaoCol >= 0 && rowData[idxDescricaoCol] != null && String(rowData[idxDescricaoCol]).trim() !== "") ? String(rowData[idxDescricaoCol]).trim() : "";
        dadosRetorno.observacoes.descricao = descricaoPlanilha || (dadosRetorno.observacoes.descricao || "");
        const codigoProjetoPlanilha = (idxProjetoCol >= 0 && rowData[idxProjetoCol] != null && String(rowData[idxProjetoCol]).trim() !== "") ? String(rowData[idxProjetoCol]).trim() : "";
        dadosRetorno.observacoes.projeto = codigoProjetoPlanilha || (dadosRetorno.observacoes.projeto || "");
        // Observações internas: usar SEMPRE a coluna OBSERVAÇÕES por nome, nunca JSON_DADOS (evita preencher com o JSON inteiro)
        const idxObsCol = _findHeaderIndex(headers, "OBSERVAÇÕES");
        if (idxObsCol >= 0 && idxObsCol !== jsonIdx && rowData[idxObsCol] != null && String(rowData[idxObsCol]).trim() !== "") {
          dadosRetorno.observacoes.observacoesKanban = String(rowData[idxObsCol]).trim();
        } else {
          dadosRetorno.observacoes.observacoesKanban = dadosRetorno.observacoes.observacoesKanban || "";
        }
        const idxProcessosCol = _findHeaderIndex(headers, "PROCESSOS");
        if (idxProcessosCol >= 0 && idxProcessosCol !== jsonIdx && rowData[idxProcessosCol] != null && String(rowData[idxProcessosCol]).trim() !== "") {
          dadosRetorno.observacoes.processos = String(rowData[idxProcessosCol]).trim();
        } else {
          dadosRetorno.observacoes.processos = dadosRetorno.observacoes.processos || "";
        }
        if (dadosRetorno.cliente == null) dadosRetorno.cliente = {};
        const clientePlanilha = (idxClienteCol >= 0 && rowData[idxClienteCol] != null && String(rowData[idxClienteCol]).trim() !== "") ? String(rowData[idxClienteCol]).trim() : "";
        dadosRetorno.cliente.nome = clientePlanilha || (dadosRetorno.cliente.nome || "");
        // Projeto: preencher data, indice, iniciais a partir da coluna PROJETO para o formulário exibir e salvar corretamente
        if (codigoProjetoPlanilha.length >= 6) {
          const matchV = codigoProjetoPlanilha.match(/^(.+?)(_v\d+)$/);
          const codigoBase = matchV ? matchV[1] : codigoProjetoPlanilha;
          if (codigoBase.length >= 9) {
            dadosRetorno.projeto = { data: codigoBase.substring(0, 6), indice: codigoBase.substring(6, 7), iniciais: codigoBase.substring(7, 9), versao: matchV ? matchV[2] : "" };
          } else if (codigoBase.length >= 6) {
            const resto = codigoBase.substring(6);
            dadosRetorno.projeto = { data: codigoBase.substring(0, 6), indice: resto.length > 0 ? resto.charAt(0) : "", iniciais: resto.length > 1 ? resto.substring(1) : "", versao: matchV ? matchV[2] : "" };
          }
        }
        // Data do cliente (coluna DATA): converter para yyyy-MM-dd para input type="date"
        const dataRaw = idxDataCol >= 0 ? rowData[idxDataCol] : null;
        if (dataRaw != null && String(dataRaw).trim() !== "") {
          if (Object.prototype.toString.call(dataRaw) === "[object Date]") {
            dadosRetorno.cliente.data = Utilities.formatDate(dataRaw, Session.getScriptTimeZone(), "yyyy-MM-dd");
          } else {
            const dataStr = String(dataRaw).trim();
            const matchBr = dataStr.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
            if (matchBr) {
              dadosRetorno.cliente.data = matchBr[3] + "-" + matchBr[2].padStart(2, "0") + "-" + matchBr[1].padStart(2, "0");
            } else {
              dadosRetorno.cliente.data = dataStr;
            }
          }
        }
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
    // PRAZO (prazo de entrega) no índice 11; PRAZO_PROPOSTA no 12 (estrutura 15 colunas); 14 colunas = antiga
  const prazoRaw = sheetProj ? (rowData[11] || "") : (rowData[10] || "");
  const temColPrazoProposta = sheetProj && rowData.length >= 15;
  const prazoPropostaRaw = temColPrazoProposta ? (rowData[12] || "") : "";
  const idxObsBasico = _findHeaderIndex(headers, "OBSERVAÇÕES");
  const observacoesKanbanBasico = (idxObsBasico >= 0 && idxObsBasico !== jsonIdx && rowData[idxObsBasico] != null && String(rowData[idxObsBasico]).trim() !== "") ? String(rowData[idxObsBasico]).trim() : "";

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

    // Extrai código do projeto (formato YYMMDD + índice + iniciais, opcional _v2, _v3)
    const codigoProjetoRaw = String(projeto || "").trim();
    const matchVersao = codigoProjetoRaw.match(/^(.+?)(_v\d+)$/);
    const codigoProjeto = matchVersao ? matchVersao[1] : codigoProjetoRaw;
    let projetoData = "";
    let projetoIndice = "";
    let projetoIniciais = "";

    if (codigoProjeto.length >= 9) {
      projetoData = codigoProjeto.substring(0, 6);
      projetoIndice = codigoProjeto.substring(6, 7);
      projetoIniciais = codigoProjeto.substring(7, 9);
    } else if (codigoProjeto.length >= 6) {
      projetoData = codigoProjeto.substring(0, 6);
      const resto = codigoProjeto.substring(6);
      if (resto.length > 0) {
        projetoIndice = resto.charAt(0);
        projetoIniciais = resto.substring(1).replace(/_v\d+$/, ""); // remove _v2 se sobrou
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
        versao: matchVersao ? matchVersao[2] : "" // _v2, _v3 quando existir
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
        prazo: prazo,
        prazoProposta: (prazoPropostaRaw != null && String(prazoPropostaRaw).trim() !== "") ? String(prazoPropostaRaw).trim() : "",
        vendedor: "",
        materialCond: "",
        pagamento: "",
        adicional: "",
        observacoesKanban: observacoesKanbanBasico,
        projeto: codigoProjetoRaw, // Inclui _v2, _v3 quando existir
        descricao: descricao,
        processos: processos
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

/**
 * Menu exibido ao abrir a planilha. Inclui opção para ordenar a aba Projetos por código.
 */
function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('TUBA')
      .addItem('Ordenar Projetos por código (permanente)', 'ordenarPlanilhaProjetosPorCodigo')
      .addSeparator()
      .addItem('Migrar dados de pedido (Projetos → Pedidos)', 'executarMigracaoProjetosParaPedidos')
      .addToUi();
  } catch (e) {
    Logger.log('onOpen menu TUBA: ' + (e && e.message));
  }
}

/**
 * Executa migrarDadosPedidoProjetosParaPedidos e exibe o resultado em um alerta.
 * Use antes de excluir as colunas de pedido da aba Projetos.
 */
function executarMigracaoProjetosParaPedidos() {
  var ui = SpreadsheetApp.getUi();
  var msg = "Migração: copiando para Pedidos os dados preenchidos em Projetos (onde Pedidos está vazio).\nNão sobrescreve dados já existentes em Pedidos.";
  if (ui.alert('Migrar dados Projetos → Pedidos', msg, ui.ButtonSet.OK_CANCEL) !== ui.Button.OK) return;
  var r = migrarDadosPedidoProjetosParaPedidos();
  var texto = "Linhas de Pedidos atualizadas: " + r.atualizados + "\nProjetos que ganharam linha em Pedidos: " + r.semLinhaPedidos;
  if (r.erros && r.erros.length > 0) texto += "\n\nErros:\n" + r.erros.slice(0, 10).join("\n") + (r.erros.length > 10 ? "\n... e mais " + (r.erros.length - 10) : "");
  ui.alert("Migração concluída", texto, ui.ButtonSet.OK);
}