/************* Code.gs *************/
const ss = SpreadsheetApp.openById("1wMIbd8r2HeniFLTYaG8Yhhl8CWmaHaW5oXBVnxj0xos");
const SHEET_CALC = ss.getSheetByName("Tabelas para cálculos");
const SHEET_VEIC = ss.getSheetByName('Controle de Veículos');
const SHEET_ORC = ss.getSheetByName("Orçamentos");
const SHEET_MANU_NAME = ss.getSheetByName("Registro de Manutenções");
const SHEET_PED = ss.getSheetByName("Pedidos");
const SHEET_MAT = ss.getSheetByName("Controle de Materiais");
const SHEET_AVAL = ss.getSheetByName("Avaliações");
const SHEET_LOGS = ss.getSheetByName("Logs");
const SHEET_CLIENTES = ss.getSheetByName("Cadastro de Clientes");
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
    // A=Código do Produto, B=Descrição do Produto, H=Preço Unitário de Venda, I=Unidade
    const produtos = [];
    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      if (row[0]) { // se tem código (coluna A)
        produtos.push({
          codigo: row[0],                    // Coluna A - Código do Produto
          descricao: row[1] || "",           // Coluna B - Descrição do Produto
          familia: row[3] || "",             // Coluna D - Família de Produto
          tipo: row[4] || "",                // Coluna E - Tipo do Produto
          preco: parseFloat(row[7]) || 0,    // Coluna H - Preço Unitário de Venda
          unidade: row[8] || "UN"            // Coluna I - Unidade
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
 * Insere um produto na aba "Relação de produtos"
 * @param {Object} produto - Objeto com os dados do produto
 */
function inserirProdutoNaRelacao(produto) {
  try {
    const SHEET_PRODUTOS = ss.getSheetByName("Relação de produtos");
    if (!SHEET_PRODUTOS) {
      Logger.log("Aba 'Relação de produtos' não encontrada");
      return false;
    }
    
    // Verifica se o produto já existe
    const dados = SHEET_PRODUTOS.getDataRange().getValues();
    for (let i = 1; i < dados.length; i++) {
      if (dados[i][0] === produto.codigo) {
        Logger.log("Produto " + produto.codigo + " já existe na relação");
        return false; // Produto já existe
      }
    }
    
    // Estrutura da planilha:
    // A=Código do Produto, B=Descrição, C=Código da Família, D=Família de Produto, 
    // E=Tipo do Produto, F=Código EAN, G=Código NCM, H=Preço Unitário de Venda, 
    // I=Unidade, J=Características
    const novaLinha = [
      produto.codigo || "",           // A - Código do Produto
      produto.descricao || "",        // B - Descrição do Produto
      "",                             // C - Código da Família (vazio)
      produto.familia || "",          // D - Família de Produto
      produto.tipo || "",             // E - Tipo do Produto
      "",                             // F - Código EAN (vazio)
      "",                             // G - Código NCM (vazio)
      produto.preco || 0,             // H - Preço Unitário de Venda
      produto.unidade || "UN",        // I - Unidade
      produto.caracteristicas || ""   // J - Características
    ];
    
    SHEET_PRODUTOS.appendRow(novaLinha);
    Logger.log("Produto " + produto.codigo + " inserido na relação");
    return true;
  } catch (err) {
    Logger.log("Erro ao inserir produto na relação: " + err);
    return false;
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
function criarOuUsarPasta(codigoProjeto, nomePasta, data) {
  const root = DriveApp.getFolderById(ID_PASTA_PRINCIPAL);
  const anoFolder = getOrCreateSubFolder(root, "20" + data.substring(0, 2));
  const mesFolder = getOrCreateSubFolder(anoFolder, data.substring(0, 4));
  const diaFolder = getOrCreateSubFolder(mesFolder, data);
  const comFolder = getOrCreateSubFolder(diaFolder, "COM");

  const folders = comFolder.getFolders();
  while (folders.hasNext()) {
    const f = folders.next();
    if (f.getName().startsWith(codigoProjeto)) return f;
  }
  const novaPastaNome = codigoProjeto + " COT " + nomePasta;
  return comFolder.createFolder(novaPastaNome);
}

function buscarNomePastaPorCodigo(codigoProjeto) {
  const root = DriveApp.getFolderById(ID_PASTA_PRINCIPAL);
  const ano = codigoProjeto.slice(0, 2);
  const mes = codigoProjeto.slice(0, 4);
  const dia = codigoProjeto.slice(0, 6);
  try {
    const pasta = DriveApp.getFolderById(ID_PASTA_PRINCIPAL)
      .getFoldersByName("20" + ano).next()
      .getFoldersByName(mes).next()
      .getFoldersByName(dia).next()
      .getFoldersByName("COM").next();
    const folders = pasta.getFolders();
    while (folders.hasNext()) {
      const f = folders.next();
      if (f.getName().startsWith(codigoProjeto)) return f.getName().replace(codigoProjeto + " COT ", "");
    }
    return "";
  } catch (e) {
    return "";
  }
}

// ========================= GERAR PDF (VERSÃO AJUSTADA) =========================
function gerarPdfOrcamento(
  chapas, cliente, observacoes, codigoProjeto, nomePasta, data, versao, somaProcessosPedido, descricaoProcessosPedido, produtosCadastrados
) {
  try {

    // Incrementa contador de propostas
    incrementarContador("totalPropostas");

    const resultados = calcularOrcamento(chapas);
    
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

    const pasta = criarOuUsarPasta(codigoProjeto, nomePasta, data);
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

    const numeroProposta = (codigoProjeto || "") + (versao || "");

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

    const itensHtml = resultados.map(function (p) {
      return ''
        + '<tr>'
        + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:3px; border:0.1px solid #fff; font-size:9pt;">${esc(p.codigo || "")}</td>`
        + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:3px; border:0.1px solid #fff; font-size:9pt;">${esc(p.descricao || "")}</td>`
        + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:3px; border:0.1px solid #fff; text-align:right; font-size:9pt;">${esc(p.quantidade || 0)}</td>`
        + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:3px; border:0.1px solid #fff; text-align:right; font-size:9pt;">${formatBR(p.precoUnitario || 0)}</td>`
        + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:3px; border:0.1px solid #fff; text-align:right; font-size:9pt;">${formatBR(p.precoTotal || 0)}</td>`
        + '</tr>';
    }).join('');

    const processosPedidoRow = (somaProcessosPedido && Number(somaProcessosPedido) > 0)
      ? ''
      + '<tr>'
      + `<td colspan="4" bgcolor="${rowColor}" style="background:${rowColor}; padding:3px; border:0.1px solid #fff; text-align:right; font-size:9pt;"><strong>${esc(descricaoProcessosPedido || "")}</strong></td>`
      + `<td bgcolor="${rowColor}" style="background:${rowColor}; padding:3px; border:0.1px solid #fff; text-align:right; font-size:9pt;">${formatBR(somaProcessosPedido)}</td>`
      + '</tr>'
      : '';

    const htmlContent = `
      <html>
      <head>
        <meta charset="utf-8">
        <style>
          body, table, th, td { -webkit-print-color-adjust: exact; print-color-adjust: exact; }
          body { font-family: Arial, sans-serif; font-size: 10pt; color: #000; margin: 10px; line-height:1.3; -webkit-font-smoothing:antialiased; } /* margem menor */
          .header { display:flex; justify-content:space-between; align-items:center; margin-bottom:15px; }
          .logo { max-height:200px; }
          .company-info { text-align:right; font-size:10pt; }
          h2 { text-align:left; margin:30px 0 50px 0; } /* mais espaço abaixo */
          h3 { margin-top:25px; margin-bottom:5px; }
          table { width:100%; border-collapse:collapse; border-spacing:0; font-size:9pt; }
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

        <h2>Proposta Comercial Nº ${esc(numeroProposta)}</h2>

        <h3>Informações do Cliente:</h3>
        <p style="margin-bottom:25px;">
          <p><strong>${esc(cliente.nome)}</strong><br></p>
            CNPJ/CPF: ${esc(cliente.cpf)}<br>
            ${esc(cliente.endereco)}<br>
            <b>Telefone:</b> ${esc(cliente.telefone)}<br>
            <b>Email:</b> ${esc(cliente.email)}<br>
            <b>Responsável:</b> ${esc(cliente.responsavel || "-")}
        </p>

        <h3>Itens da Proposta Comercial</h3>
        <table style="margin-top:12px;">
          <tr>
            <th bgcolor="${headerColor}" style="background:${headerColor}; color:#ffffff; padding:4px; text-align:left; border:0.1px solid #fff; font-size:9pt;">Código</th>
            <th bgcolor="${headerColor}" style="background:${headerColor}; color:#ffffff; padding:4px; text-align:left; border:0.1px solid #fff; font-size:9pt;">Descrição</th>
            <th bgcolor="${headerColor}" style="background:${headerColor}; color:#ffffff; padding:4px; text-align:right; border:0.1px solid #fff; font-size:9pt;">Quant.</th>
            <th bgcolor="${headerColor}" style="background:${headerColor}; color:#ffffff; padding:4px; text-align:right; border:0.1px solid #fff; font-size:9pt;">Unit.</th>
            <th bgcolor="${headerColor}" style="background:${headerColor}; color:#ffffff; padding:4px; text-align:right; border:0.1px solid #fff; font-size:9pt;">Valor Total</th>
          </tr>
          ${itensHtml}
          ${processosPedidoRow}
        </table>

        <!-- Totais alinhados com a coluna Valor Total -->
<div style="width:100%; text-align:right; margin-top:8px;">
  <table style="display:inline-block; border-collapse:collapse; width:100%; max-width:320px;">
    <tr>
      <td style="border:none; text-align:right; width:140px; background:#fff; padding:4px; font-weight:bold;">Subtotal:</td>
      <td style="border:none; text-align:right; background:${rowColor}; padding:4px; width:110px; font-weight:bold;">${formatBR(totalPecas)}</td>
    </tr>
    <tr>
      <td style="border:none; text-align:right; background:#fff; padding:4px; font-weight:bold;">Total:</td>
      <td style="border:none; text-align:right; background:${rowColor}; padding:4px; width:110px; font-weight:bold;">${formatBR(totalFinal)}</td>
    </tr>
  </table>
</div>

        <h3 style="margin-top:18px;">Outras Informações</h3>
        <p style="font-size:10pt; line-height:1.35;">
          <b>Proposta Comercial - incluído em:</b> ${esc(dataBrasil)} às ${esc(horaBrasil)}<br>
          <b>Validade da Proposta:</b> 30 dias
        </p>

        <p style="font-size:10pt; line-height:1.35;">
          <b>Previsão de Faturamento:</b> ${esc(formatarDataBrasil(observacoes.faturamento) || "-")}<br>
          <b>Pagamento:</b> ${esc(observacoes.pagamento || "-")}<br>
          <b>Vendedor:</b> ${esc(observacoes.vendedor || "-")}<br>
        </p>

        <p style="font-size:10pt; line-height:1.35;">
          <b>PROJ:</b> ${esc(observacoes.projeto || "-")}<br>
          <b>Condições do Material:</b> ${esc(observacoes.materialCond || "-")}<br>
        </p>

        ${observacoes.adicional ? `<p style="font-size:10pt; line-height:1.35;"><b>Observações adicionais:</b><br>${esc(observacoes.adicional)}</p>` : ""}

      </body>
      </html>
    `;

    const blob = Utilities.newBlob(htmlContent, "text/html", "orcamento.html");
    const pdf = blob.getAs("application/pdf").setName("Proposta_" + numeroProposta + ".pdf");
    const file = comSubFolder.createFile(pdf);

    let memoriaUrl = null;
    try {
      const memoria = gerarPdfMemoriaCalculo(chapas, cliente, codigoProjeto, comSubFolder, file.getName());
      memoriaUrl = memoria && memoria.url ? memoria.url : null;
    } catch (eMem) {
      Logger.log("Erro ao gerar memoria de calculo: " + eMem.toString());
    }

    registrarOrcamento(cliente, codigoProjeto, totalFinal, dataBrasil, file.getUrl(), memoriaUrl, chapas);
    return { url: file.getUrl(), nome: file.getName(), memoriaUrl: memoriaUrl };
  } catch (err) {
    Logger.log("ERRO gerarPdfOrcamento: " + err.toString());
    throw err;
  }
}

/* ======= gerarPdfMemoriaCalculo corrigido: lê linha de referência APÓS flush ======= */
function gerarPdfMemoriaCalculo(chapas, cliente, codigoProjeto, pastaDestino, nomePdfOrcamento) {
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
function registrarOrcamento(cliente, codigoProjeto, valorTotal, dataOrcamento, urlPdf, urlMemoria, chapas) {
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

  // ----- Aqui fazíamos appendRow; agora vamos checar existência e atualizar se necessário -----
  try {
    // Serializa os dados das chapas para JSON (para armazenar e recuperar depois)
    const chapasJson = JSON.stringify(chapas || []);
    
    // definir as colunas que vamos gravar (mesma ordem que estava no appendRow)
    const rowValues = [
      cliente.nome || "",
      cliente.responsavel || "",
      codigoProjeto || "",
      valorTotal || "",
      dataOrcamento || "",
      processosStr || "",
      urlPdf || "",
      urlMemoria || "",
      "Enviado",
      chapasJson  // Nova coluna para armazenar dados das chapas/peças
    ];

    // tenta encontrar linha existente com o mesmo PROJETO (coluna "PROJETO")
    const linhaExistente = findRowByColumnValue(SHEET_ORC, "PROJETO", codigoProjeto);

    if (linhaExistente) {

      SHEET_ORC.getRange(linhaExistente, 1, 1, rowValues.length).setValues([rowValues]);
    } else {

      SHEET_ORC.appendRow(rowValues);
    }
  } catch (err) {
    Logger.log("Erro ao registrarOrcamento (atualizar/inserir): " + err);
    // fallback: tentar appendRow (comportamento antigo) se algo falhar
    try {
      const chapasJson = JSON.stringify(chapas || []);
      SHEET_ORC.appendRow([
        cliente.nome || "",
        cliente.responsavel || "",
        codigoProjeto || "",
        valorTotal || "",
        dataOrcamento || "",
        processosStr || "",
        urlPdf || "",
        urlMemoria || "",
        "Enviado",
        chapasJson
      ]);
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

function getDashboardStats() {
  const props = PropertiesService.getScriptProperties();

  // Contadores baseados em eventos (propostas e etiquetas)
  const propostas = Number(props.getProperty("totalPropostas")) || 0;
  const etiquetas = Number(props.getProperty("totalEtiquetas")) || 0;

  // Materiais cadastrados
  const materiais = SHEET_MAT ? Math.max(SHEET_MAT.getLastRow() - 1, 0) : 0;

  // Pedidos
  const pedidos = SHEET_PED ? Math.max(SHEET_PED.getLastRow() - 1, 0) : 0;

  // Logs
  const logs = SHEET_LOGS ? Math.max(SHEET_LOGS.getLastRow() - 1, 0) : 0;

  // Kanban = pedidos que não estão finalizados
  let kanban = 0;
  if (SHEET_PED && pedidos > 0) {
    const data = SHEET_PED.getRange(2, 3, pedidos).getValues();
    kanban = data.filter(r => r[0] && r[0] !== "Finalizado").length;
  }

  // Orçamentos ativos
  const orcamentos = SHEET_ORC ? Math.max(SHEET_ORC.getLastRow() - 1, 0) : 0;


  return { propostas, orcamentos, pedidos, kanban, etiquetas, materiais, logs };
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

    // --- Orçamentos ---
    if (typeof SHEET_ORC !== 'undefined' && SHEET_ORC) {
      const valsOrc = SHEET_ORC.getDataRange().getValues();
      if (valsOrc && valsOrc.length > 0) {
        const headersOrc = valsOrc[0];
        const idxCliente = _findHeaderIndex(headersOrc, "CLIENTE");
        const idxProjeto = _findHeaderIndex(headersOrc, "PROJETO");
        const idxStatus = _findHeaderIndex(headersOrc, "STATUS");

        for (let i = 1; i < valsOrc.length; i++) {
          const row = valsOrc[i];
          const status = idxStatus >= 0 ? row[idxStatus] : row[2];
          if (status && !["Expirado/Perdido", "Convertido em Pedido", "Enviado"].includes(status)) {
            data["Processo de Orçamento"].push({
              cliente: idxCliente >= 0 ? row[idxCliente] : "",
              projeto: idxProjeto >= 0 ? row[idxProjeto] : "",
              descricao: status || ""
            });
          }
        }
      }
    }

    // --- Logs (mapa) ---
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

    // --- Pedidos ---
    if (typeof SHEET_PED !== 'undefined' && SHEET_PED) {
      const valsPed = SHEET_PED.getDataRange().getValues();
      if (valsPed && valsPed.length > 0) {
        const headersPed = valsPed[0];
        const idxClienteP = _findHeaderIndex(headersPed, "Cliente");
        const idxProjetoP = _findHeaderIndex(headersPed, "Número do Projeto");
        const idxStatusP = _findHeaderIndex(headersPed, "Status");
        const idxObsP = _findHeaderIndex(headersPed, "Observações") >= 0 ? _findHeaderIndex(headersPed, "Observações") : _findHeaderIndex(headersPed, "Observacoes");
        const idxTempoP = _findHeaderIndex(headersPed, "Tempo estimado por processo");
        const idxPrazoP = _findHeaderIndex(headersPed, "PRAZO");

        for (let i = 1; i < valsPed.length; i++) {
          try {
            const row = valsPed[i];
            const status = idxStatusP >= 0 ? row[idxStatusP] : row[2];
            if (!status || status === "Finalizado") continue;

            const cliente = idxClienteP >= 0 ? row[idxClienteP] : "";
            const projeto = idxProjetoP >= 0 ? row[idxProjetoP] : "";
            const obs = idxObsP >= 0 ? row[idxObsP] : "";
            const tempoRaw = String(idxTempoP >= 0 ? row[idxTempoP] : "");
            let prazo = idxPrazoP >= 0 ? row[idxPrazoP] : "";
            // --- NORMALIZA PRAZO para string segura ---
            prazo = normalizePrazo(prazo);

            const chave = cliente + "|" + projeto;

            let tempoEstimado = "";
            let tempoReal = "";

            if (/Preparação/i.test(status)) {
              tempoEstimado = tempoRaw.match(/preparação\s*:? ?([\d.,]+h?)/i)?.[1] || "";
              tempoReal = mapaLogs[chave]?.preparacao_mp_cad_com || "";
            } else if (/Corte/i.test(status)) {
              tempoEstimado = tempoRaw.match(/corte\s*:? ?([\d.,]+h?)/i)?.[1] || "";
              tempoReal = mapaLogs[chave]?.corte || "";
            } else if (/Dobra/i.test(status)) {
              tempoEstimado = tempoRaw.match(/dobra\s*:? ?([\d.,]+h?)/i)?.[1] || "";
              tempoReal = mapaLogs[chave]?.dobra || "";
            } else if (/Adicion/i.test(status)) {
              tempoEstimado = tempoRaw.match(/adici.*:? ?([\d.,]+h?)/i)?.[1] || "";
              tempoReal = mapaLogs[chave]?.adicionais || "";
            }

            // push para a coluna correta (se existir)
            if (Array.isArray(data[status])) {
              data[status].push({
                cliente: cliente,
                projeto: projeto,
                observacoes: obs,
                tempoEstimado: tempoEstimado,
                tempoReal: tempoReal,
                prazo: prazo
              });
            } else {
              // se status novo, cria array e empurra
              data[status] = [{
                cliente: cliente,
                projeto: projeto,
                observacoes: obs,
                tempoEstimado: tempoEstimado,
                tempoReal: tempoReal,
                prazo: prazo
              }];
            }
          } catch (eRow) {
            Logger.log('Erro processando linha %s da aba Pedidos: %s\n%s', i + 1, eRow && eRow.message, eRow && eRow.stack);
          }
        }
      }
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
  "GuilhermeGomes": { senha: "gg5619", nivel: "usuario" },
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
    "GuilhermeGomes": "Guilherme Gomes",
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
    "GuilhermeGomes": "GS"
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

  const equipe = ["Matheus Rodrigues", "Bruno Sena", "Icaro Ferreira", "Guilherme Gomes"];
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
    'orcamento': ['admin', 'mod'],
    'materiais': ['admin', 'mod', 'usuario'],
    'geradoretiquetas': ['admin', 'mod', 'usuario'],
    'kanban': ['admin', 'mod', 'usuario'],
    'avaliacoes': ['admin'],
    'orcamentos': ['admin', 'mod'],
    'avaliacoespage': ['admin'],
    'pedidos': ['admin', 'mod'],
    'logs': ['admin', 'mod'],
    'manutencao': ['admin', 'mod', 'usuario'],
    'manu_registros': ['admin', 'mod', 'usuario'],
    'paginasprotegidas': ['admin'],
    'veiculos': ['admin', 'mod', 'usuario', 'visitante'],
    'veiculos_list': ['admin', 'mod']
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

      case 'paginasprotegidas':
        const templatePaginasProtegidas = HtmlService.createTemplateFromFile('paginasprotegidas');
        templatePaginasProtegidas.token = token;
        return templatePaginasProtegidas.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'orcamento':
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

      case 'pedidos':
        const templatePedidos = HtmlService.createTemplateFromFile('pedidos');
        templatePedidos.token = token;
        return templatePedidos.evaluate()
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

      case 'logs':
        const templateLogs = HtmlService.createTemplateFromFile('logs');
        templateLogs.token = token;
        return templateLogs.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      case 'orcamentos':
        if (!SHEET_ORC) throw new Error("Aba 'Orçamentos' não encontrada");

        const values = SHEET_ORC.getDataRange().getValues();
        const headers = values[0];
        const data = values.slice(1).map((row, index) => {
          let obj = {};
          headers.forEach((h, i) => {
            let valor = row[i];
            if (h === "VALOR TOTAL" && typeof valor === "number") {
              valor = valor.toLocaleString('pt-BR', { style: 'currency', currency: 'BRL' });
            }
            if (h === "DATA" && valor instanceof Date) {
              valor = Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy");
            }
            obj[h] = valor;
          });
          obj["_linhaPlanilha"] = index + 2;
          return obj;
        });

        const templateOrcamentos = HtmlService.createTemplateFromFile('orcamentos');
        templateOrcamentos.token = token;
        templateOrcamentos.dados = data;
        templateOrcamentos.logo = "https://i.imgur.com/F8X7ZMs.png";
        templateOrcamentos.empresa = {
          nome: "TUBA FERRAMENTARIA LTDA",
          endereco: "Estrada dos Alvarengas, 4101 - Assunção, São Bernardo do Campo - SP",
          cnpj: "10.684.825/0001-26",
          email: "tubaferram@gmail.com",
          tel: "(11) 91285-4204"
        };
        return templateOrcamentos.evaluate()
          .setFaviconUrl(FAVICON)
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

      default:
        return HtmlService.createHtmlOutput("Página não encontrada");
    }

  } catch (err) {
    return HtmlService.createHtmlOutput("Erro ao carregar a página: " + err.message);
  }
}

function adicionarOrcamentoNaPlanilha(dados) {
  try {
    if (!SHEET_ORC) throw new Error("Aba 'Orçamentos' não encontrada");
    if (!dados || !dados.CLIENTE || !dados.PROJETO) {
      throw new Error('CLIENTE e PROJETO são obrigatórios.');
    }

    // Lê cabeçalhos atuais
    const lastCol = SHEET_ORC.getLastColumn();
    if (lastCol < 1) throw new Error('Cabeçalho da planilha "Orçamentos" não encontrado.');
    const headers = SHEET_ORC.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());

    // Prepara array com valores na ordem dos headers
    const rowValues = new Array(headers.length).fill('');

    // Campos que esperamos receber (mapear conforme sua planilha)
    const campoMap = [
      'CLIENTE',
      'RESPONSÁVEL',
      'PROJETO',
      'VALOR TOTAL',
      'DATA',
      'Processos',
      'LINK DO PDF',
      'LINK DA MEMÓRIA DE CÁLCULO',
      'STATUS'
    ];

    campoMap.forEach(campo => {
      const colIdx = headers.indexOf(campo);
      if (colIdx !== -1) {
        let valor = dados[campo] !== undefined ? dados[campo] : (dados[campo.toUpperCase()] !== undefined ? dados[campo.toUpperCase()] : '');
        // Se for DATA e estiver no formato dd/mm/yyyy, converte para Date para manter formatação na planilha
        if (campo === 'DATA' && typeof valor === 'string' && /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/.test(valor)) {
          const parts = valor.split('/');
          const d = parseInt(parts[0], 10);
          const m = parseInt(parts[1], 10) - 1;
          const y = parseInt(parts[2], 10);
          valor = new Date(y, m, d);
        }
        // Mantém VALOR TOTAL como texto recebido (front-end envia string formatada). Se quiser gravar número, converta aqui.
        rowValues[colIdx] = valor;
      }
    });

    // Escreve a linha na próxima linha disponível
    const targetRow = SHEET_ORC.getLastRow() + 1;
    SHEET_ORC.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);

    // Incrementa contador de propostas (se aplicável)
    // try { incrementarContador("totalPropostas"); } catch (e) { /* não crítico */ }

    return { linha: targetRow };
  } catch (err) {
    // Propaga erro para o front-end
    throw new Error('Erro adicionarOrcamentoNaPlanilha: ' + (err.message || err));
  }
}

function excluirOrcamento(linha) {
  try {
    if (!SHEET_ORC) throw new Error("Aba 'Orçamentos' não encontrada");
    if (!linha || typeof linha !== 'number') throw new Error("Parâmetro 'linha' inválido.");

    const lastRow = SHEET_ORC.getLastRow();
    if (linha < 2 || linha > lastRow) {
      throw new Error("Número da linha inválido ou fora do intervalo de dados: " + linha);
    }

    // Opcional: ler dados antes de excluir para registrar
    // const dadosLinha = SHEET_ORC.getRange(linha, 1, 1, SHEET_ORC.getLastColumn()).getValues()[0];

    SHEET_ORC.deleteRow(linha);

    // Opcional: ajustar registros dependentes (ex.: logs ou contadores) — aqui apenas retorna sucesso
    return { ok: true };
  } catch (err) {
    throw new Error("Erro ao excluir orçamento: " + (err.message || err));
  }
}

function atualizarStatusNaPlanilha(linha, novoStatus) {
  if (!SHEET_ORC) throw new Error("Aba 'Orçamentos' não encontrada");

  const headers = SHEET_ORC.getRange(1, 1, 1, SHEET_ORC.getLastColumn()).getValues()[0];
  const colunaStatus = headers.indexOf("STATUS") + 1;
  if (colunaStatus === 0) throw new Error("Coluna STATUS não encontrada");

  // Atualiza o status
  SHEET_ORC.getRange(linha, colunaStatus).setValue(novoStatus);

  // Se for "Convertido em Pedido", faz a conversão imediatamente
  if (novoStatus === "Convertido em Pedido") {
    const linhaDados = SHEET_ORC.getRange(linha, 1, 1, SHEET_ORC.getLastColumn()).getValues()[0];

    // --- Aba Pedidos ---
    const headersPedidos = SHEET_PED.getRange(1, 1, 1, SHEET_PED.getLastColumn()).getValues()[0];
    const idxCliente = headers.indexOf("CLIENTE");
    const idxProjeto = headers.indexOf("PROJETO");
    const idxProcessos = headers.indexOf("Processos");
    const idxObs = headers.indexOf("OBSERVAÇÕES") >= 0 ? headers.indexOf("OBSERVAÇÕES") : headers.indexOf("Observações");



    const cliente = linhaDados[idxCliente] || "";
    const projeto = linhaDados[idxProjeto] || "";
    const processos = String(linhaDados[idxProcessos] || "");
    const observacoes = idxObs >= 0 ? (linhaDados[idxObs] || "") : "";


    // Determina o primeiro processo
    let statusInicial = "Processo de Preparação MP / CAD / CAM";
    if (/preparação|preparacao|mp|cad|cam/i.test(processos)) statusInicial = "Processo de Preparação MP / CAD / CAM";
    else if (/corte/i.test(processos)) statusInicial = "Processo de Corte";
    else if (/dobra/i.test(processos)) statusInicial = "Processo de Dobra";
    else if (/adicio/i.test(processos)) statusInicial = "Processos Adicionais";
    else if (/envio|coleta/i.test(processos)) statusInicial = "Envio / Coleta";

    // Insere na aba Pedidos
    const novaLinha = [];
    headersPedidos.forEach(h => {
      switch (String(h).trim()) {
        case "Cliente": novaLinha.push(cliente); break;
        case "Número do Projeto": novaLinha.push(projeto); break;
        case "Status": novaLinha.push(statusInicial); break;
        case "Observações": novaLinha.push(observacoes); break;
        case "Tempo estimado por processo": novaLinha.push(processos); break;
        default: novaLinha.push("");
      }
    });
    SHEET_PED.appendRow(novaLinha);

    // Registrar log
    registrarLog(cliente, projeto, null, statusInicial, processos);
    
    // --- Inserir produtos na "Relação de produtos" ---
    try {
      // Tenta recuperar os dados das chapas da última coluna (JSON)
      const idxChapas = linhaDados.length - 1; // Última coluna onde armazenamos o JSON
      const chapasJson = linhaDados[idxChapas];
      
      if (chapasJson && typeof chapasJson === 'string') {
        const chapas = JSON.parse(chapasJson);
        
        // Percorre todas as chapas e peças para inserir na relação de produtos
        if (Array.isArray(chapas)) {
          chapas.forEach(chapa => {
            if (chapa.pecas && Array.isArray(chapa.pecas)) {
              chapa.pecas.forEach(peca => {
                // Só insere se tiver código PRD
                if (peca.codigo && String(peca.codigo).startsWith("PRD")) {
                  const produto = {
                    codigo: peca.codigo,
                    descricao: peca.descricao || "",
                    familia: chapa.material || "",
                    tipo: "Peça",
                    preco: peca.precoUnitario || 0,
                    unidade: "UN",
                    caracteristicas: `${chapa.material} - ${peca.comprimento}x${peca.largura} - ${chapa.espessura}mm`
                  };
                  inserirProdutoNaRelacao(produto);
                }
              });
            }
          });
        }
      }
    } catch (err) {
      Logger.log("Erro ao inserir produtos na relação: " + err);
      // Não interrompe a conversão se houver erro ao inserir produtos
    }
  }
}

// ===== Atualizada: registrarLog (usa busca robusta de cabeçalhos) =====
function registrarLog(cliente, projeto, statusAntigo, statusNovo, processosStr, tipoParam) {
  try {
    if (!SHEET_LOGS) return;

    const headers = SHEET_LOGS.getRange(1, 1, 1, SHEET_LOGS.getLastColumn()).getValues()[0];
    const idxCliente = _findHeaderIndex(headers, "Cliente");
    const idxProjeto = _findHeaderIndex(headers, "Número do Projeto");
    const idxObs = _findHeaderIndex(headers, "Observações");
    const idxPrep = _findHeaderIndex(headers, "Tempo estimado / tempo real preparação") >= 0 ? _findHeaderIndex(headers, "Tempo estimado / tempo real preparação") : _findHeaderIndex(headers, "Tempo estimado / tempo real de preparação");
    const idxCorte = _findHeaderIndex(headers, "Tempo estimado / tempo real corte");
    const idxDobra = _findHeaderIndex(headers, "Tempo estimado / tempo real dobra");
    const idxAdic = _findHeaderIndex(headers, "Tempo estimado / tempo real adicionais");

    // Localiza linha existente para este cliente+projeto
    const vals = SHEET_LOGS.getDataRange().getValues();
    let linhaExistente = -1;
    for (let i = 1; i < vals.length; i++) {
      const valCliente = idxCliente >= 0 ? String(vals[i][idxCliente]) : "";
      const valProjeto = idxProjeto >= 0 ? String(vals[i][idxProjeto]) : "";
      if (valCliente === String(cliente) && valProjeto === String(projeto)) {
        linhaExistente = i + 1;
        break;
      }
    }

    if (linhaExistente === -1) {
      const nova = Array(headers.length).fill("");
      if (idxCliente >= 0) nova[idxCliente] = cliente;
      if (idxProjeto >= 0) nova[idxProjeto] = projeto;
      if (idxObs >= 0) nova[idxObs] = "";
      SHEET_LOGS.appendRow(nova);
      linhaExistente = SHEET_LOGS.getLastRow();
    }

    const agora = new Date();
    const tempoStr = Utilities.formatDate(agora, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

    // Função para extrair tempo estimado
    function extraiEstimado(processosText, chave) {
      if (!processosText || !chave) return "";
      const regex = new RegExp(chave + "\\s*:?\\s*([0-9]+(?:[.,][0-9]+)?h?)", "i");
      const m = processosText.match(regex);
      return m ? m[1] : "";
    }

    // Função auxiliar para encerrar um processo
    function encerrarProcesso(status, idxCol) {
      if (!status || idxCol < 0) return;
      const atual = SHEET_LOGS.getRange(linhaExistente, idxCol + 1).getValue() || "";
      SHEET_LOGS.getRange(linhaExistente, idxCol + 1).setValue((atual ? atual + " | " : "") + "Fim: " + tempoStr);
    }

    // Função auxiliar para iniciar um processo
    function iniciarProcesso(status, idxCol) {
      if (!status || idxCol < 0) return;
      const estim = extraiEstimado(processosStr || "", status);
      const atual = SHEET_LOGS.getRange(linhaExistente, idxCol + 1).getValue() || "";
      SHEET_LOGS.getRange(linhaExistente, idxCol + 1)
        .setValue((estim ? "Estimado: " + estim + " | " : "") + "Início: " + tempoStr);
    }

    // Encerrar processo antigo
    if (statusAntigo) {
      if (/Preparação/i.test(statusAntigo)) encerrarProcesso("Preparação", idxPrep);
      if (/Corte/i.test(statusAntigo)) encerrarProcesso("Corte", idxCorte);
      if (/Dobra/i.test(statusAntigo)) encerrarProcesso("Dobra", idxDobra);
      if (/Adicion/i.test(statusAntigo)) encerrarProcesso("Adicion", idxAdic);
    }

    // Iniciar novo processo
    if (statusNovo) {
      if (/Preparação/i.test(statusNovo)) iniciarProcesso("Preparação", idxPrep);
      if (/Corte/i.test(statusNovo)) iniciarProcesso("Corte", idxCorte);
      if (/Dobra/i.test(statusNovo)) iniciarProcesso("Dobra", idxDobra);
      if (/Adicion/i.test(statusNovo)) iniciarProcesso("Adicion", idxAdic);
    }

  } catch (err) {
    Logger.log("Erro registrarLog: " + err);
  }
}

// Incrementa contadores de eventos como propostas e etiquetas
function incrementarContador(tipo) {
  const props = PropertiesService.getScriptProperties();
  const valorAtual = Number(props.getProperty(tipo)) || 0;
  props.setProperty(tipo, valorAtual + 1);
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
    "GuilhermeGomes": "Guilherme Gomes"
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
  materiais.getRange(rowIndex, 13).setValue(urlPdf); // Coluna M = ETIQUETA

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
    nf: 11
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

function getLogs() {
  if (!SHEET_LOGS) return [];

  const values = SHEET_LOGS.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];
  const data = values.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => {
      obj[h] = row[i];
    });
    return obj;
  });
  return data;
}

function getPedidos() {
  try {
    const valuesRaw = SHEET_PED.getDataRange().getValues();
    if (!valuesRaw || valuesRaw.length === 0) {
      return { dados: [], opcoesStatus: [] };
    }

    // Normaliza valores: null/undefined -> '', Date -> string formatada
    const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'UTC';
    const values = valuesRaw.map(row => row.map(cell => {
      if (cell === null || cell === undefined) return '';
      if (Object.prototype.toString.call(cell) === '[object Date]') {
        // Formato de data: ajuste se necessário
        return Utilities.formatDate(cell, tz, 'dd/MM/yy');
      }
      // Tenta converter outros tipos para string de forma segura
      try {
        return String(cell);
      } catch (e) {
        return '';
      }
    }));

    // Determina índice da coluna de "Status" (caso-insensitivo)
    const header = values[0] || [];
    const statusIndex = header.findIndex(h => /status/i.test(String(h)));

    // Tenta obter validação (opções) de maneira segura:
    //  - se houver validação VALUE_IN_LIST -> usa a lista
    //  - se houver validação VALUE_IN_RANGE -> lê os valores do intervalo referenciado
    //  - senão, fallback para lista estática
    let opcoesStatus = [];
    try {
      if (statusIndex >= 0) {
        // procura validação na 2ª linha da coluna de status (se existir)
        const rowParaValidacao = Math.min(2, values.length); // evita request inválida se a sheet pequena
        const range = SHEET_PED.getRange(rowParaValidacao, statusIndex + 1);
        const rule = range.getDataValidation();
        if (rule) {
          const criteria = rule.getCriteriaType();
          if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
            const critValues = rule.getCriteriaValues() && rule.getCriteriaValues()[0];
            if (Array.isArray(critValues) && critValues.length) {
              opcoesStatus = critValues.map(String);
            }
          } else if (criteria === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
            // se for validação por intervalo, pega os valores desse intervalo
            const critRange = rule.getCriteriaValues() && rule.getCriteriaValues()[0];
            if (critRange && critRange.getA1Notation) {
              const vals = critRange.getValues().flat().map(v => v === null || v === undefined ? '' : String(v).trim()).filter(v => v !== '');
              // remove duplicatas
              opcoesStatus = Array.from(new Set(vals));
            }
          }
        }
      }
    } catch (e) {
      // não falhar: apenas loga e segue para fallback
      Logger.log('getPedidos: falha ao obter validação de dados: %s', e.message);
    }

    // Fallback se não obteve opções válidas
    if (!opcoesStatus || opcoesStatus.length === 0) {
      opcoesStatus = [
        "Processo de Preparação MP / CAD / CAM",
        "Processo de Corte",
        "Processo de Dobra",
        "Processos Adicionais",
        "Envio / Coleta",
        "Finalizado"
      ];
    }

    return { dados: values, opcoesStatus: opcoesStatus };
  } catch (e) {
    Logger.log('getPedidos error: %s\n%s', e.message, e.stack);
    // lança erro para que o client receba via withFailureHandler e possamos debugar
    throw new Error('getPedidos failed: ' + (e.message || 'erro desconhecido'));
  }
}

function updatePedido(row, col, value) {
  try {
    row = Number(row);
    col = Number(col);
    if (!row || !col || row < 1 || col < 1) {
      throw new Error('Índice de linha/coluna inválido: ' + row + ',' + col);
    }

    // Converte datas de string ISO para Date se necessário (opcional)
    // Se value for string no formato yyyy-MM-dd, tenta converter para Date
    if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}$/.test(value)) {
      const parts = value.split('-').map(Number);
      // month é 0-based
      const dt = new Date(parts[0], parts[1] - 1, parts[2]);
      SHEET_PED.getRange(row, col).setValue(dt);
    } else {
      SHEET_PED.getRange(row, col).setValue(value);
    }
  } catch (e) {
    Logger.log('updatePedido error (row=%s col=%s): %s', row, col, e.message);
    throw new Error('updatePedido failed: ' + (e.message || 'erro desconhecido'));
  }
}

function deletePedido(row) {
  try {
    row = Number(row);
    if (!row || row < 2) {
      // evita deletar cabeçalho (linha 1) por engano, e valida índice
      throw new Error('Índice de linha inválido para exclusão: ' + row);
    }
    SHEET_PED.deleteRow(row);
  } catch (e) {
    Logger.log('deletePedido error (row=%s): %s', row, e.message);
    throw new Error('deletePedido failed: ' + (e.message || 'erro desconhecido'));
  }
}


function atualizarStatusKanban(cliente, projeto, novoStatus) {
  try {
    if (!SHEET_PED) return;

    let statusAntigo = '';
    let processosStr = '';

    // --- Aba Pedidos ---
    const dadosRaw = SHEET_PED.getDataRange().getValues();
    if (!dadosRaw || dadosRaw.length < 1) return;

    const headers = dadosRaw[0].map(h => String(h || '').trim());
    // busca índices de forma case-insensitive (tolerante a espaçamento)
    const idxCliente = headers.findIndex(h => /^cliente$/i.test(h) || /cliente/i.test(h));
    const idxProjeto = headers.findIndex(h => /n[uú]mero do projeto/i.test(h) || /n[oº]mero do projeto/i.test(h) || /projeto/i.test(h));
    const idxStatus = headers.findIndex(h => /status/i.test(h));
    const idxTempo = headers.findIndex(h => /tempo estimado/i.test(h) || /tempo/i.test(h));
    const idxPrazoP = headers.findIndex(h => /prazo/i.test(h));

    // valida índices
    if (idxCliente < 0 || idxProjeto < 0 || idxStatus < 0) {
      Logger.log('atualizarStatusKanban: cabeçalhos não encontrados. cliente:%s projeto:%s status:%s', idxCliente, idxProjeto, idxStatus);
      return;
    }

    for (let i = 1; i < dadosRaw.length; i++) {
      const row = dadosRaw[i];
      const valCliente = String(row[idxCliente] || '').trim();
      const valProjeto = String(row[idxProjeto] || '').trim();
      if (valCliente === String(cliente).trim() && valProjeto === String(projeto).trim()) {
        statusAntigo = String(row[idxStatus] || '').trim();
        // normaliza processosStr (se for Date converte para string)
        const tempoCell = row[idxTempo];
        if (Object.prototype.toString.call(tempoCell) === '[object Date]') {
          const tz = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone() || 'GMT';
          processosStr = Utilities.formatDate(tempoCell, tz, 'yyyy-MM-dd');
        } else {
          processosStr = String(tempoCell || '').trim();
        }
        // atualiza célula de status (linha no sheet = i+1)
        SHEET_PED.getRange(i + 1, idxStatus + 1).setValue(novoStatus);
        break;
      }
    }

    // Atualiza log com processosStr
    registrarLog(cliente, projeto, statusAntigo, novoStatus, processosStr, 'INICIO');
  } catch (e) {
    Logger.log('atualizarStatusKanban error: %s\n%s', e.message, e.stack);
    throw new Error('atualizarStatusKanban failed: ' + (e.message || 'erro desconhecido'));
  }
}

function salvarRascunho(nomeRascunho, dados) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    // Tenta obter e analisar o JSON, usa um objeto vazio se falhar (primeira vez)
    const rascunhosSalvos = JSON.parse(userProperties.getProperty('rascunhos_orcamento') || '{}');

    // Cria uma chave única baseada no tempo atual
    const key = new Date().getTime().toString();

    // Salva o rascunho. JSON.stringify() é importante, pois PropertiesService só aceita strings.
    rascunhosSalvos[key] = {
      nome: nomeRascunho,
      data: new Date().toLocaleString('pt-BR'), // Data e hora para exibir
      dados: dados
    };

    userProperties.setProperty('rascunhos_orcamento', JSON.stringify(rascunhosSalvos));
  } catch (e) {
    Logger.log("Erro ao salvar rascunho: " + e.message);
    throw new Error("Erro ao salvar rascunho: " + e.message);
  }
}

function carregarRascunho(key) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const rascunhosSalvos = JSON.parse(userProperties.getProperty('rascunhos_orcamento') || '{}');

    if (rascunhosSalvos[key]) {
      return rascunhosSalvos[key].dados;
    }
    return null;
  } catch (e) {
    Logger.log("Erro ao carregar rascunho: " + e.message);
    throw new Error("Erro ao carregar rascunho: " + e.message);
  }
}

function getListaRascunhos() {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const rascunhosSalvos = JSON.parse(userProperties.getProperty('rascunhos_orcamento') || '{}');

    // Converte o objeto de rascunhos em um array, formatando o nome para o dropdown
    return Object.keys(rascunhosSalvos).map(key => ({
      key: key,
      nome: `${rascunhosSalvos[key].nome} (${rascunhosSalvos[key].data})`
    })).sort((a, b) => b.key - a.key); // Ordena pelo mais recente
  } catch (e) {
    Logger.log("Erro ao obter lista de rascunhos: " + e.message);
    // Retorna array vazio em caso de erro para não quebrar a UI
    return [];
  }
}

function deletarRascunho(key) {
  try {
    const userProperties = PropertiesService.getUserProperties();
    const rascunhosSalvos = JSON.parse(userProperties.getProperty('rascunhos_orcamento') || '{}');

    delete rascunhosSalvos[key];

    userProperties.setProperty('rascunhos_orcamento', JSON.stringify(rascunhosSalvos));
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

function atualizarOrcamentoNaPlanilha(linha, dataObj) {
  linha = Number(linha);
  if (!linha || linha < 2) {
    throw new Error('Parâmetro "linha" inválido. Deve ser número de linha da planilha (>= 2).');
  }

  // Ajuste o nome da aba se necessário
  var SHEET_NAME = 'Orçamentos';
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    throw new Error("Aba '" + SHEET_NAME + "' não encontrada.");
  }

  // Cabeçalhos
  var lastCol = sh.getLastColumn();
  var headers = sh.getRange(1, 1, 1, lastCol).getValues()[0] || [];

  // Função utilitária para normalizar strings (lowercase, sem acentos, sem espaços/pontuação)
  function normalizeKey(s) {
    if (s === null || s === undefined) return '';
    // transformar em string, lower, decompor acentos e remover diacríticos, remover não-alnum
    return String(s).trim().toLowerCase()
      .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
      .replace(/[^a-z0-9]/g, '');
  }

  // Normaliza as chaves do dataObj para lookup rápido
  var normalizedData = {};
  // Também guarda mapa de chave original -> valor para detectar propriedade mesmo vazia
  var originalKeys = Object.keys(dataObj || {});
  originalKeys.forEach(function (k) {
    var nk = normalizeKey(k);
    normalizedData[nk] = dataObj[k];
  });

  // Lê a linha atual para preservar valores não enviados
  var currentRow = sh.getRange(linha, 1, 1, lastCol).getValues()[0] || [];

  // Monta nova linha: se header correspondente existe em normalizedData (mesmo que valor seja ''), usa-o; senão mantém currentRow
  var newRow = headers.map(function (h, idx) {
    var hk = normalizeKey(h);
    if (normalizedData.hasOwnProperty(hk)) {
      // usar valor enviado (pode ser vazio)
      return normalizedData[hk];
    }
    // tentar também correspondência por header sem acentos/sem espaços: já feito por hk
    // se não há correspondência, manter o valor atual da célula
    return currentRow[idx];
  });

  // Gravar nova linha
  try {
    sh.getRange(linha, 1, 1, newRow.length).setValues([newRow]);
    return { success: true, linha: linha };
  } catch (err) {
    throw new Error('Erro ao escrever na planilha: ' + (err && err.message ? err.message : err));
  }
}