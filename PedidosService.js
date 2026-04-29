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
      LINK_RELATORIO_PROJETO: "LINK_RELATORIO_PROJETO",
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
      }
    } else {
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
      SheetCache.invalidate(sheet);
    }

    try {
      var linhaProj = null;
      if (dadosProjeto && dadosProjeto._linhaPlanilha != null) {
        var parsed = parseInt(dadosProjeto._linhaPlanilha, 10);
        if (!isNaN(parsed) && parsed >= 2) linhaProj = parsed;
      }
      if (linhaProj == null && projeto) {
        var sheetProj = SHEET_PROJ;
        if (sheetProj) linhaProj = findRowByColumnValue(sheetProj, "PROJETO", projeto);
      }
      if (linhaProj >= 2) {
        atualizarProjetoNaPlanilha(linhaProj, dadosAtualizacao, { apenasJsonDados: true });
      }
    } catch (syncErr) {
      Logger.log("Sync Pedidos->Projetos: " + syncErr.message);
    }
    try {
      // Mantém Livro Diário em espelho do Pedido (datas, parcelas, status, valores).
      if (typeof gerarLancamentosLivroDiarioParaPedido === "function") {
        // Modo rápido para não bloquear fluxos críticos (ex.: geração de PDF).
        gerarLancamentosLivroDiarioParaPedido(projeto, null, { skipPosProcess: true });
      }
    } catch (eLivroSync) {
      Logger.log("Sync Pedidos->LivroDiario: " + (eLivroSync && eLivroSync.message ? eLivroSync.message : eLivroSync));
    }
    return { sucesso: true };
  } catch (e) {
    Logger.log("atualizarPedidoNaPlanilha error: " + e.message);
    throw new Error("Erro ao atualizar pedido: " + (e.message || "erro desconhecido"));
  }
}

function updatesPedidoDesdeProjeto(p) {
  var updates = {};
  var j = p.JSON_DADOS || p["JSON_DADOS"] || "";
  var parsed = null;
  if (j && typeof j === "string" && j.trim()) {
    try {
      parsed = JSON.parse(j);
    } catch (e) { }
  }
  var dados = parsed && (parsed.dados || parsed) ? (parsed.dados || parsed) : {};
  var obs = dados.observacoes || {};

  var exigeNF = !!obs.temNotaFiscal;
  var nf = (p["NF"] || p["NOTA_FISCAL"] || p["NOTA FISCAL"] || "").toString().trim();
  if (!nf && (obs.valorNF != null && obs.valorNF !== "")) nf = String(obs.valorNF).trim();
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

function ensurePedidoRow(linhaProjetos) {
  try {
    var sheetProj = SHEET_PROJ;
    if (!sheetProj || sheetProj.getLastRow() < linhaProjetos) return;
    var headers = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
    var row = sheetProj.getRange(linhaProjetos, 1, 1, sheetProj.getLastColumn()).getValues()[0];
    var obj = {};
    headers.forEach(function (h, i) { obj[h] = row[i]; });
    var codigo = (obj.PROJETO || "").toString().trim();
    if (!codigo) return;
    var dataComp = (obj["DATA_COMPETENCIA"] || obj["DATA COMPETÊNCIA"] || obj["DATA COMPETENCIA"] || obj.DATA || "").toString().trim();
    var valorTotal = (obj["VALOR TOTAL"] != null && obj["VALOR TOTAL"] !== "") ? obj["VALOR TOTAL"] : (obj["VALOR_TOTAL"] != null && obj["VALOR_TOTAL"] !== "") ? obj["VALOR_TOTAL"] : "";
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

function converterProjetoParaPedido(linha, codigoVersaoSelecionado) {
  linha = Number(linha);
  if (!linha || linha < 2) throw new Error("Linha inválida.");
  codigoVersaoSelecionado = String(codigoVersaoSelecionado || "").trim();
  if (!codigoVersaoSelecionado) throw new Error("Selecione uma versão.");

  const sheetProj = SHEET_PROJ;
  if (!sheetProj) throw new Error("Aba 'Projetos' não encontrada.");

  const lastCol = sheetProj.getLastColumn();
  const headers = sheetProj.getRange(1, 1, 1, lastCol).getValues()[0];
  const row = sheetProj.getRange(linha, 1, linha, lastCol).getValues()[0];

  const idxProjeto = _findHeaderIndexProjetos(headers, "PROJETO");
  const idxJson = _findHeaderIndexProjetos(headers, "JSON_DADOS");
  if (idxJson < 0) throw new Error("Coluna JSON_DADOS não encontrada.");

  const codigoProjetoCol = idxProjeto >= 0 ? String(row[idxProjeto] || "").trim() : "";
  const codigoBase = codigoProjetoCol.replace(/_v\d+$/i, "").trim() || codigoProjetoCol.trim();
  if (!codigoBase) throw new Error("Código base do projeto não encontrado.");

  let parsed;
  try { parsed = JSON.parse(String(row[idxJson] || "{}")); } catch (e) { parsed = null; }
  if (!parsed || !parsed.dados) throw new Error("JSON_DADOS inválido.");

  const versoesArr = Array.isArray(parsed.versoes) ? parsed.versoes : [];
  const versoesBackup = JSON.parse(JSON.stringify(versoesArr || []));
  let selected = null;
  if (codigoVersaoSelecionado === "__base__") {
    selected = {
      dados: parsed.dados || {},
      numeroSequencial: parsed.numeroSequencial,
      nomeVersao: parsed.dados && parsed.dados.nomeVersao != null ? parsed.dados.nomeVersao : "",
    };
  } else {
    const v = versoesArr.find(function (x) { return x && String(x.codigo || "").trim() === codigoVersaoSelecionado; });
    if (!v) throw new Error("Versão selecionada não encontrada no JSON_DADOS.");
    selected = {
      dados: v.dados || {},
      numeroSequencial: v.numeroSequencial,
      nomeVersao: v.nomeVersao != null ? v.nomeVersao : (v.dados ? v.dados.nomeVersao : ""),
    };
  }

  if (!selected || !selected.dados) throw new Error("Dados da versão inválidos.");

  const dadosVersao = selected.dados || {};
  const nomeVersaoSel = String(selected.nomeVersao || dadosVersao.nomeVersao || "").trim();
  const nomeVersaoSelFinal = nomeVersaoSel || "";

  const numeroSequencialSel = selected.numeroSequencial != null ? selected.numeroSequencial : dadosVersao.numeroSequencial;
  if (numeroSequencialSel == null || String(numeroSequencialSel).trim() === "") {
    throw new Error("Número sequencial da versão não encontrado.");
  }

  const dadosFormularioCompleto = JSON.parse(JSON.stringify(dadosVersao));
  dadosFormularioCompleto.linhaProjeto = linha;
  dadosFormularioCompleto.numeroSequencial = numeroSequencialSel;
  if (!dadosFormularioCompleto.observacoes) dadosFormularioCompleto.observacoes = {};
  dadosFormularioCompleto.observacoes.projeto = codigoBase;
  dadosFormularioCompleto.tipoPdf = "Pedido";
  dadosFormularioCompleto.nomeVersao = nomeVersaoSelFinal;

  const clienteRaw = dadosFormularioCompleto.cliente || {};
  const cliente = {
    nome: clienteRaw.nome || "",
    nomeAbreviado: clienteRaw.nomeAbreviado || "",
    cpf: clienteRaw.cpf || "",
    endereco: clienteRaw.endereco || "",
    telefone: clienteRaw.telefone || "",
    email: clienteRaw.email || "",
    responsavel: clienteRaw.responsavel || "",
    dataCliente: clienteRaw.dataCliente || ""
  };

  const procPedidoArr = Array.isArray(dadosFormularioCompleto.processosPedido) ? dadosFormularioCompleto.processosPedido : [];
  const produtosCadastrados = Array.isArray(dadosFormularioCompleto.produtosCadastrados) ? dadosFormularioCompleto.produtosCadastrados : [];

  let totalPecas = 0;
  let totalProdutos = 0;
  produtosCadastrados.forEach(function (p) {
    const pt = (p && p.precoTotal != null && p.precoTotal !== "") ? Number(p.precoTotal) :
      ((p && p.precoUnitario != null) ? Number(p.precoUnitario) : 0) * ((p && p.quantidade != null) ? Number(p.quantidade) : 0);
    if (!isNaN(pt) && isFinite(pt)) totalProdutos += pt;
  });
  const subtotalAntesProcessos = totalPecas + totalProdutos;

  let somaProcessosPedido = 0;
  const processosPedidoComValores = [];
  procPedidoArr.forEach(function (procData) {
    if (!procData) return;
    const tipo = (procData.tipo === "desconto" || procData.tipo === "custo") ? procData.tipo : "custo";
    const tipoValor = (procData.tipoValor === "percentual" || procData.tipoValor === "fixo") ? procData.tipoValor : "fixo";
    let valor = 0;
    if (tipo === "desconto") {
      if (tipoValor === "percentual" || procData.percentual != null) {
        const pct = parseFloat(procData.percentual) || 0;
        valor = -(subtotalAntesProcessos * pct / 100);
      } else {
        valor = -(parseFloat(procData.valorFixo) || 0);
      }
    } else {
      valor = parseFloat(procData.valorFixo) || 0;
    }
    somaProcessosPedido += valor;
    processosPedidoComValores.push({
      tipo: tipo,
      descricao: procData.descricao || (tipo === "desconto" ? "Desconto" : "Custo extra"),
      valor: valor
    });
  });

  const valorTotalOrcamento = subtotalAntesProcessos + somaProcessosPedido;
  const descricaoProcessosPedido = processosPedidoComValores.map(function (p) { return p.descricao; }).join(" / ");

  const observacoesRaw = dadosFormularioCompleto.observacoes || {};
  const observacoes = {
    projeto: codigoBase,
    descricao: observacoesRaw.descricao || "",
    prazo: observacoesRaw.prazo || "",
    vendedor: observacoesRaw.vendedor || "",
    materialCond: observacoesRaw.materialCond || "",
    transporte: observacoesRaw.transporte || "",
    pagamento: observacoesRaw.pagamento || "",
    adicional: observacoesRaw.adicional || "",
    observacoesKanban: observacoesRaw.observacoesKanban || ""
  };

  const dataProjeto = dadosFormularioCompleto.projeto ? dadosFormularioCompleto.projeto.data : "";
  const infoPagamento = {
    texto: observacoesRaw.pagamento || "",
    valorTotal: valorTotalOrcamento
  };

  const resultPdf = gerarPdfOrcamento(
    [],
    cliente,
    observacoes,
    codigoBase,
    "",
    dataProjeto,
    "",
    somaProcessosPedido,
    descricaoProcessosPedido,
    produtosCadastrados,
    dadosFormularioCompleto,
    infoPagamento,
    true,
    nomeVersaoSelFinal,
    true,
    false
  );

  try {
    const jsonApos = sheetProj.getRange(linha, idxJson + 1).getValue();
    let parsedApos = null;
    try { parsedApos = JSON.parse(String(jsonApos || "{}")); } catch (eP) { parsedApos = null; }
    if (parsedApos && typeof parsedApos === "object") {
      if (!Array.isArray(parsedApos.versoes)) parsedApos.versoes = [];
      if (versoesBackup.length > 0 && parsedApos.versoes.length < versoesBackup.length) {
        parsedApos.versoes = versoesBackup;
        sheetProj.getRange(linha, idxJson + 1).setValue(JSON.stringify(parsedApos));
      }
    }
  } catch (eKeepV) {
    Logger.log("Aviso: falha ao preservar versoes após converterProjetoParaPedido: " + (eKeepV && eKeepV.message ? eKeepV.message : eKeepV));
  }

  try {
    const tz = ss.getSpreadsheetTimeZone() || "America/Sao_Paulo";
    const hoje = new Date();
    const dataCompetenciaStr = Utilities.formatDate(hoje, tz, "dd/MM/yyyy");
    try {
      const idxClienteCol = _findHeaderIndexProjetos(headers, "CLIENTE");
      const idxValorTotalCol = _findHeaderIndexProjetos(headers, "VALOR TOTAL");
      const clienteAtual = idxClienteCol >= 0 ? String(row[idxClienteCol] || "").trim() : (cliente.nome || "");
      const valorTotalAtual = idxValorTotalCol >= 0 ? row[idxValorTotalCol] : valorTotalOrcamento;
      atualizarPedidoNaPlanilha(codigoBase, { DATA_COMPETENCIA: dataCompetenciaStr }, {
        CLIENTE: clienteAtual,
        "VALOR TOTAL": valorTotalAtual,
        _dataCompetencia: dataCompetenciaStr,
        _linhaPlanilha: linha
      });
    } catch (ePed) {
      Logger.log("Aviso: não foi possível definir DATA_COMPETENCIA: " + (ePed && ePed.message ? ePed.message : ePed));
    }

    const idxJsonNow = _findHeaderIndexProjetos(headers, "JSON_DADOS");
    if (idxJsonNow >= 0) {
      const jsonCellNow = sheetProj.getRange(linha, idxJsonNow + 1).getValue();
      let parsedNow = null;
      try { parsedNow = JSON.parse(String(jsonCellNow || "{}")); } catch (eJ) { parsedNow = null; }
      if (!parsedNow) parsedNow = { dados: {} };
      if (!parsedNow.dados) parsedNow.dados = {};
      if (!parsedNow.dados.infoPedido) parsedNow.dados.infoPedido = {};
      if (!parsedNow.dados.infoPedido.statusDates) parsedNow.dados.infoPedido.statusDates = {};
      if (!parsedNow.dados.infoPedido.dataVirouPedido) parsedNow.dados.infoPedido.dataVirouPedido = dataCompetenciaStr;
      if (!parsedNow.dados.infoPedido.statusDates.preparacao) parsedNow.dados.infoPedido.statusDates.preparacao = dataCompetenciaStr;
      sheetProj.getRange(linha, idxJsonNow + 1).setValue(JSON.stringify(parsedNow));
    }
  } catch (eInfo) {
    Logger.log("Aviso converterProjetoParaPedido/infoPedido: " + (eInfo && eInfo.message ? eInfo.message : eInfo));
  }

  try {
    if (typeof gerarLancamentosLivroDiarioParaPedido === "function") {
      // Modo rápido para evitar timeout no retorno ao formulário.
      gerarLancamentosLivroDiarioParaPedido(codigoBase, null, { skipPosProcess: true });
    }
  } catch (eLivro) {
    Logger.log("Aviso converterProjetoParaPedido/livroDiario: " + (eLivro && eLivro.message ? eLivro.message : eLivro));
  }

  return {
    sucesso: true,
    url: resultPdf && resultPdf.url ? resultPdf.url : "",
    nomeVersao: nomeVersaoSelFinal,
    numeroSequencial: numeroSequencialSel
  };
}
