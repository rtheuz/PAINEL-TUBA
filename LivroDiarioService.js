const LIVRO_DIARIO_HEADERS = [
  "DATA COMPETÊNCIA",
  "DATA VENCIMENTO",
  "DATA PAGAMENTO",
  "CLIENTE",
  "CÓDIGO DO PROJETO",
  "DESCRIÇÃO DO PROJETO",
  "CONTA CONTÁBIL",
  "DESCRICAO ABREVIADA da CONTA CONTÁBIL",
  "CONTA FINANCEIRA",
  "MEIO de PAGAMENTO",
  "VALIDADE FISCAL",
  "VALOR",
  "SALDO CONTA FINANCEIRA",
  "SALDO GERAL",
  "STATUS DO PAGAMENTO",
  "RESPONSÁVEL DO LANÇAMENTO",
  "DATA DA ÚLTIMA MOFIFICAÇÃO",
  "OBSERVAÇÕES"
];

const LIVRO_DIARIO_CAD_HEADERS = [
  "CONTA CONTÁBIL",
  "DESCRICAO ABREVIADA",
  "CONTA FINANCEIRA",
  "MEIO de PAGAMENTO",
  "VALIDADE FISCAL",
  "STATUS DO PAGAMENTO"
];

const LIVRO_DIARIO_PREFS_KEY = "LIVRO_DIARIO_PREFS_V1";
const LIVRO_DIARIO_CONTA_CONTABIL_PADRAO_PEDIDO = "TUBA Laser _ Receita _ Operacional _ Indústria";
const LIVRO_DIARIO_DESC_PADRAO_PEDIDO = "TuLa Receita Indústria";

function _findHeaderIndexGeneric(headers, targetName) {
  if (!headers || !headers.length) return -1;
  if (typeof _findHeaderIndex === "function") return _findHeaderIndex(headers, targetName);
  var wanted = String(targetName || "").trim().toLowerCase();
  for (var i = 0; i < headers.length; i++) {
    if (String(headers[i] || "").trim().toLowerCase() === wanted) return i;
  }
  return -1;
}

function _normalizeSheetKey_(name) {
  return String(name || "")
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function _getSheetByNames_(names) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var all = spreadsheet.getSheets();
  var wanted = {};
  (names || []).forEach(function (n) { wanted[_normalizeSheetKey_(n)] = true; });
  for (var i = 0; i < all.length; i++) {
    var sh = all[i];
    if (wanted[_normalizeSheetKey_(sh.getName())]) return sh;
  }
  return null;
}

function _ensureSheetWithHeaders(sheetName, headers, aliases) {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var namesToFind = [sheetName].concat(aliases || []);
  var sheet = _getSheetByNames_(namesToFind);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    return sheet;
  }
  var currentHeaders = sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getValues()[0];
  var changed = false;
  headers.forEach(function (h) {
    if (_findHeaderIndexGeneric(currentHeaders, h) < 0) {
      currentHeaders.push(h);
      changed = true;
    }
  });
  if (changed) {
    sheet.getRange(1, 1, 1, currentHeaders.length).setValues([currentHeaders]);
  }
  return sheet;
}

function ensureLivroDiarioSheet() {
  return _ensureSheetWithHeaders("Livro Diário", LIVRO_DIARIO_HEADERS, ["Livro Diario"]);
}

function ensureLivroDiarioCadastroSheet() {
  return _ensureSheetWithHeaders("Livro Diário Cadastro", LIVRO_DIARIO_CAD_HEADERS, ["Livro Diario Cadastro"]);
}

function _asDateLivro(value) {
  if (value == null || value === "") return null;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return isNaN(value.getTime()) ? null : new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  if (typeof value === "number" && !isNaN(value)) {
    var ms = (value - 25569) * 86400 * 1000;
    var dNum = new Date(ms);
    return isNaN(dNum.getTime()) ? null : new Date(dNum.getFullYear(), dNum.getMonth(), dNum.getDate());
  }
  var s = String(value).trim();
  if (!s) return null;
  var mBr4 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (mBr4) return new Date(parseInt(mBr4[3], 10), parseInt(mBr4[2], 10) - 1, parseInt(mBr4[1], 10));
  var mBr2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2})$/);
  if (mBr2) {
    var yy = parseInt(mBr2[3], 10);
    var year = yy >= 70 ? 1900 + yy : 2000 + yy;
    return new Date(year, parseInt(mBr2[2], 10) - 1, parseInt(mBr2[1], 10));
  }
  var mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (mIso) return new Date(parseInt(mIso[1], 10), parseInt(mIso[2], 10) - 1, parseInt(mIso[3], 10));
  return null;
}

function _formatDateLivro2y(value) {
  var d = _asDateLivro(value);
  if (!d) return "";
  var dd = ("0" + d.getDate()).slice(-2);
  var mm = ("0" + (d.getMonth() + 1)).slice(-2);
  var yy = String(d.getFullYear()).slice(-2);
  return dd + "/" + mm + "/" + yy;
}

function _dateKeyLivro(value) {
  var d = _asDateLivro(value);
  if (!d) return null;
  return d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
}

function _numLivro(v) {
  if (typeof _parseCurrency === "function") return _parseCurrency(v);
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

function _normLivro(v) {
  return String(v == null ? "" : v).trim();
}

function _getLivroDiarioPrefs() {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(LIVRO_DIARIO_PREFS_KEY);
    if (!raw) return { modoOrdenacao: "data_pagamento", contaFinanceiraSaldo: "" };
    var parsed = JSON.parse(raw);
    return {
      modoOrdenacao: parsed && parsed.modoOrdenacao ? String(parsed.modoOrdenacao) : "data_pagamento",
      contaFinanceiraSaldo: parsed && parsed.contaFinanceiraSaldo ? String(parsed.contaFinanceiraSaldo) : ""
    };
  } catch (e) {
    return { modoOrdenacao: "data_pagamento", contaFinanceiraSaldo: "" };
  }
}

function _setLivroDiarioPrefs(prefs) {
  var current = _getLivroDiarioPrefs();
  var next = {
    modoOrdenacao: prefs && prefs.modoOrdenacao ? String(prefs.modoOrdenacao) : current.modoOrdenacao,
    contaFinanceiraSaldo: prefs && prefs.contaFinanceiraSaldo != null ? String(prefs.contaFinanceiraSaldo) : current.contaFinanceiraSaldo
  };
  PropertiesService.getScriptProperties().setProperty(LIVRO_DIARIO_PREFS_KEY, JSON.stringify(next));
  return next;
}

function getLivroDiarioPreferencias() {
  return _getLivroDiarioPrefs();
}

function salvarLivroDiarioPreferencias(prefs) {
  var next = _setLivroDiarioPrefs(prefs || {});
  ordenarLivroDiario(next.modoOrdenacao);
  recalcularSaldosLivroDiario(next.contaFinanceiraSaldo);
  return next;
}

function _usuarioLancamentoPorToken(token) {
  try {
    if (!token || typeof getUsuarioLogadoPorToken !== "function") return "Sistema";
    var user = getUsuarioLogadoPorToken(token);
    if (!user) return "Sistema";
    return user.usuario || user.nome || "Sistema";
  } catch (e) {
    return "Sistema";
  }
}

function getLivroDiarioCadastro() {
  var sh = ensureLivroDiarioCadastroSheet();
  if (sh.getLastRow() < 2) {
    return {
      linhas: [],
      contasContabeis: [],
      descricoesAbreviadas: [],
      contasFinanceiras: [],
      meiosPagamento: [],
      validadesFiscais: [],
      statusPagamento: [],
      mapContaParaDescricao: {},
      mapDescricaoParaConta: {}
    };
  }
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var idxConta = 0, idxDesc = 1, idxFin = 2, idxMeio = 3, idxVal = 4, idxStatus = 5;
  var contas = {};
  var descs = {};
  var fins = {};
  var meios = {};
  var vals = {};
  var status = {};
  var mapContaParaDescricao = {};
  var mapDescricaoParaConta = {};
  var linhas = data.map(function (r, i) {
    var obj = {
      rowIndex: i + 2,
      contaContabil: _normLivro(r[idxConta]),
      descricaoAbreviada: _normLivro(r[idxDesc]),
      contaFinanceira: _normLivro(r[idxFin]),
      meioPagamento: _normLivro(r[idxMeio]),
      validadeFiscal: _normLivro(r[idxVal]),
      statusPagamento: _normLivro(r[idxStatus])
    };
    if (obj.contaContabil) contas[obj.contaContabil] = true;
    if (obj.descricaoAbreviada) descs[obj.descricaoAbreviada] = true;
    if (obj.contaFinanceira) fins[obj.contaFinanceira] = true;
    if (obj.meioPagamento) meios[obj.meioPagamento] = true;
    if (obj.validadeFiscal) vals[obj.validadeFiscal] = true;
    if (obj.statusPagamento) status[obj.statusPagamento] = true;
    if (obj.contaContabil && obj.descricaoAbreviada && !mapContaParaDescricao[obj.contaContabil]) {
      mapContaParaDescricao[obj.contaContabil] = obj.descricaoAbreviada;
    }
    if (obj.descricaoAbreviada && obj.contaContabil && !mapDescricaoParaConta[obj.descricaoAbreviada]) {
      mapDescricaoParaConta[obj.descricaoAbreviada] = obj.contaContabil;
    }
    return obj;
  });
  return {
    linhas: linhas,
    contasContabeis: Object.keys(contas).sort(),
    descricoesAbreviadas: Object.keys(descs).sort(),
    contasFinanceiras: Object.keys(fins).sort(),
    meiosPagamento: Object.keys(meios).sort(),
    validadesFiscais: Object.keys(vals).sort(),
    statusPagamento: Object.keys(status).sort(),
    mapContaParaDescricao: mapContaParaDescricao,
    mapDescricaoParaConta: mapDescricaoParaConta
  };
}

function upsertLivroDiarioCadastroItem(item, token) {
  var sh = ensureLivroDiarioCadastroSheet();
  var payload = item || {};
  var row = [
    _normLivro(payload.contaContabil),
    _normLivro(payload.descricaoAbreviada),
    _normLivro(payload.contaFinanceira),
    _normLivro(payload.meioPagamento),
    _normLivro(payload.validadeFiscal),
    _normLivro(payload.statusPagamento)
  ];
  if (!row.some(function (x) { return !!x; })) throw new Error("Informe ao menos um campo para cadastro.");
  if (sh.getLastRow() >= 2) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues();
    for (var i = 0; i < data.length; i++) {
      var same = true;
      for (var c = 0; c < 6; c++) {
        if (_normLivro(data[i][c]) !== row[c]) {
          same = false;
          break;
        }
      }
      if (same) {
        return { ok: true, rowIndex: i + 2, duplicado: true, usuario: _usuarioLancamentoPorToken(token) };
      }
    }
  }
  sh.appendRow(row);
  return { ok: true, rowIndex: sh.getLastRow(), duplicado: false, usuario: _usuarioLancamentoPorToken(token) };
}

function _getDescricaoProjetoPorCodigo(codigoProjeto) {
  try {
    if (!SHEET_PROJ || !codigoProjeto) return "";
    var linha = findRowByColumnValue(SHEET_PROJ, "PROJETO", codigoProjeto);
    if (!linha) {
      var base = String(codigoProjeto).replace(/_v\d+$/i, "");
      linha = findRowByColumnValue(SHEET_PROJ, "PROJETO", base);
    }
    if (!linha) return "";
    var headers = SHEET_PROJ.getRange(1, 1, 1, SHEET_PROJ.getLastColumn()).getValues()[0];
    var idxDesc = _findHeaderIndexGeneric(headers, "DESCRIÇÃO");
    if (idxDesc < 0) return "";
    return _normLivro(SHEET_PROJ.getRange(linha, idxDesc + 1).getValue());
  } catch (e) {
    return "";
  }
}

function _getProjetoDadosBasicosPorCodigo(codigoProjeto) {
  var out = { cliente: "", descricaoProjeto: "", temNotaFiscal: null };
  try {
    if (!SHEET_PROJ || !codigoProjeto) return out;
    var linha = findRowByColumnValue(SHEET_PROJ, "PROJETO", codigoProjeto);
    if (!linha) {
      var base = String(codigoProjeto).replace(/_v\d+$/i, "");
      linha = findRowByColumnValue(SHEET_PROJ, "PROJETO", base);
    }
    if (!linha) return out;
    var headers = SHEET_PROJ.getRange(1, 1, 1, SHEET_PROJ.getLastColumn()).getValues()[0];
    var idxCliente = _findHeaderIndexGeneric(headers, "CLIENTE");
    var idxDesc = _findHeaderIndexGeneric(headers, "DESCRIÇÃO");
    var idxJson = _findHeaderIndexGeneric(headers, "JSON_DADOS");
    if (idxCliente >= 0) out.cliente = _normLivro(SHEET_PROJ.getRange(linha, idxCliente + 1).getValue());
    if (idxDesc >= 0) out.descricaoProjeto = _normLivro(SHEET_PROJ.getRange(linha, idxDesc + 1).getValue());
    if (idxJson >= 0) {
      var cell = SHEET_PROJ.getRange(linha, idxJson + 1).getValue();
      if (cell && String(cell).trim()) {
        var parsed = null;
        try { parsed = JSON.parse(String(cell).trim()); } catch (e) { parsed = null; }
        if (parsed) {
          var dados = parsed.dados || parsed;
          var obs = dados && dados.observacoes ? dados.observacoes : {};
          if (obs && obs.temNotaFiscal != null) out.temNotaFiscal = !!obs.temNotaFiscal;
          if (!out.cliente && dados && dados.cliente && dados.cliente.nome) out.cliente = _normLivro(dados.cliente.nome);
          if (!out.descricaoProjeto && obs && obs.descricao) out.descricaoProjeto = _normLivro(obs.descricao);
        }
      }
    }
  } catch (e) {
    // ignore and keep fallback object
  }
  return out;
}

function _buildLancamentoFromPedido(codigoProjeto, parcela, projetoInfo) {
  var dataComp = _formatDateLivro2y(projetoInfo.dataCompetencia || projetoInfo.dataEntrega || "");
  var dataVenc = _formatDateLivro2y(parcela.dataVencimento || projetoInfo.dataVencimento || "");
  if (!dataVenc && parcela.condicao) {
    var cond = String(parcela.condicao).toLowerCase();
    if (cond.indexOf("pedido") >= 0) dataVenc = dataComp;
    else if (cond.indexOf("entrega") >= 0) dataVenc = _formatDateLivro2y(projetoInfo.dataEntrega || projetoInfo.dataCompetencia || "");
  }
  var obsAuto = "[AUTO_PEDIDO_PARCELA_" + parcela.numero + "_DE_" + parcela.total + "]";
  return {
    "DATA COMPETÊNCIA": dataComp,
    "DATA VENCIMENTO": dataVenc,
    "DATA PAGAMENTO": "",
    "CLIENTE": _normLivro(projetoInfo.cliente),
    "CÓDIGO DO PROJETO": _normLivro(codigoProjeto),
    "DESCRIÇÃO DO PROJETO": _normLivro(projetoInfo.descricaoProjeto),
    "CONTA CONTÁBIL": LIVRO_DIARIO_CONTA_CONTABIL_PADRAO_PEDIDO,
    "DESCRICAO ABREVIADA da CONTA CONTÁBIL": LIVRO_DIARIO_DESC_PADRAO_PEDIDO,
    "CONTA FINANCEIRA": "",
    "MEIO de PAGAMENTO": "",
    "VALIDADE FISCAL": _normLivro(projetoInfo.validadeFiscal || ""),
    "VALOR": Number(parcela.valor) || 0,
    "SALDO CONTA FINANCEIRA": "",
    "SALDO GERAL": "",
    "STATUS DO PAGAMENTO": _normLivro(projetoInfo.statusPagamento || "Pendente"),
    "RESPONSÁVEL DO LANÇAMENTO": "Sistema",
    "DATA DA ÚLTIMA MOFIFICAÇÃO": _formatDateLivro2y(new Date()),
    "OBSERVAÇÕES": obsAuto
  };
}

function gerarLancamentosLivroDiarioParaPedido(codigoProjeto, token) {
  if (!codigoProjeto) throw new Error("Código do projeto é obrigatório.");
  var pedidoMap = (typeof getPedidosSheetMap === "function") ? getPedidosSheetMap() : {};
  var rowPed = pedidoMap[codigoProjeto] || pedidoMap[String(codigoProjeto).replace(/_v\d+$/i, "")];
  if (!rowPed) return { ok: true, inseridos: 0, motivo: "Pedido não encontrado" };

  var valorTotal = _numLivro(rowPed.VALOR_TOTAL);
  var condicoes = _normLivro(rowPed.CONDICOES_PAGAMENTO);
  var dataComp = _normLivro(rowPed.DATA_COMPETENCIA);
  var dataEntrega = _normLivro(rowPed.DATA_ENTREGA);
  var dataVenc = _normLivro(rowPed.DATA_VENCIMENTO);
  var dataBase = dataEntrega || dataComp;
  var parcelas = null;
  if (typeof _calcularParcelasPedidos === "function") {
    try { parcelas = _calcularParcelasPedidos(condicoes, valorTotal, dataBase); } catch (e) { parcelas = null; }
  }
  if (!parcelas || !parcelas.length) {
    parcelas = [{ numero: 1, valor: valorTotal, dataVencimento: dataVenc || dataBase, total: 1 }];
  } else {
    parcelas = parcelas.map(function (p) { return Object.assign({}, p); });
    parcelas.forEach(function (p, i) { p.numero = i + 1; p.total = parcelas.length; });
  }
  if (parcelas.length === 1 && !parcelas[0].total) parcelas[0].total = 1;

  var projetoInfo = {
    cliente: rowPed.CLIENTE || "",
    descricaoProjeto: _getDescricaoProjetoPorCodigo(codigoProjeto),
    dataCompetencia: dataComp,
    dataEntrega: dataEntrega,
    dataVencimento: dataVenc,
    statusPagamento: rowPed.STATUS_PAGAMENTO || "Pendente"
  };
  var metaProj = _getProjetoDadosBasicosPorCodigo(codigoProjeto);
  if (!projetoInfo.cliente && metaProj.cliente) projetoInfo.cliente = metaProj.cliente;
  if (!projetoInfo.descricaoProjeto && metaProj.descricaoProjeto) projetoInfo.descricaoProjeto = metaProj.descricaoProjeto;
  if (metaProj.temNotaFiscal === true) projetoInfo.validadeFiscal = "SIM";
  else if (metaProj.temNotaFiscal === false) projetoInfo.validadeFiscal = "SN";
  else {
    var nf = _normLivro(rowPed.NOTA_FISCAL || rowPed.NF || "");
    projetoInfo.validadeFiscal = (nf && nf.toUpperCase() !== "SN") ? "SIM" : "SN";
  }

  var sh = ensureLivroDiarioSheet();
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idxProjeto = _findHeaderIndexGeneric(headers, "CÓDIGO DO PROJETO");
  var idxObs = _findHeaderIndexGeneric(headers, "OBSERVAÇÕES");
  var existing = {};
  if (sh.getLastRow() >= 2 && idxProjeto >= 0 && idxObs >= 0) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    data.forEach(function (r) {
      var p = _normLivro(r[idxProjeto]);
      var o = _normLivro(r[idxObs]);
      if (!p || !o) return;
      existing[p + "|" + o] = true;
    });
  }

  var userName = _usuarioLancamentoPorToken(token);
  var toInsert = [];
  parcelas.forEach(function (parc) {
    if (!parc.total) parc.total = parcelas.length || 1;
    var rowObj = _buildLancamentoFromPedido(codigoProjeto, parc, projetoInfo);
    rowObj["RESPONSÁVEL DO LANÇAMENTO"] = userName || "Sistema";
    var key = rowObj["CÓDIGO DO PROJETO"] + "|" + rowObj["OBSERVAÇÕES"];
    if (existing[key]) return;
    toInsert.push(LIVRO_DIARIO_HEADERS.map(function (h) { return rowObj[h] != null ? rowObj[h] : ""; }));
  });

  if (toInsert.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, toInsert.length, LIVRO_DIARIO_HEADERS.length).setValues(toInsert);
  }

  var prefs = _getLivroDiarioPrefs();
  ordenarLivroDiario(prefs.modoOrdenacao);
  recalcularSaldosLivroDiario(prefs.contaFinanceiraSaldo);

  return { ok: true, inseridos: toInsert.length, parcelas: parcelas.length };
}

function _validarVinculoConta(payload) {
  var conta = _normLivro(payload["CONTA CONTÁBIL"]);
  var desc = _normLivro(payload["DESCRICAO ABREVIADA da CONTA CONTÁBIL"]);
  if (!conta && !desc) return;
  var cad = getLivroDiarioCadastro();
  if (conta && desc) {
    var expected = cad.mapContaParaDescricao[conta];
    if (expected && expected !== desc) {
      throw new Error("Descrição abreviada não corresponde à conta contábil selecionada.");
    }
    var expectedConta = cad.mapDescricaoParaConta[desc];
    if (expectedConta && expectedConta !== conta) {
      throw new Error("Conta contábil não corresponde à descrição abreviada selecionada.");
    }
  } else if (conta) {
    payload["DESCRICAO ABREVIADA da CONTA CONTÁBIL"] = cad.mapContaParaDescricao[conta] || "";
  } else if (desc) {
    payload["CONTA CONTÁBIL"] = cad.mapDescricaoParaConta[desc] || "";
  }
}

function salvarLivroDiarioLancamento(lancamento, token) {
  var sh = ensureLivroDiarioSheet();
  var payloadIn = lancamento || {};
  var userName = _usuarioLancamentoPorToken(token);
  var payload = {};
  LIVRO_DIARIO_HEADERS.forEach(function (h) {
    payload[h] = payloadIn[h] != null ? payloadIn[h] : "";
  });

  payload["DATA COMPETÊNCIA"] = _formatDateLivro2y(payload["DATA COMPETÊNCIA"]);
  payload["DATA VENCIMENTO"] = _formatDateLivro2y(payload["DATA VENCIMENTO"]);
  payload["DATA PAGAMENTO"] = _formatDateLivro2y(payload["DATA PAGAMENTO"]);
  payload["DATA DA ÚLTIMA MOFIFICAÇÃO"] = _formatDateLivro2y(new Date());
  payload["RESPONSÁVEL DO LANÇAMENTO"] = userName || "Sistema";
  payload["VALOR"] = _numLivro(payload["VALOR"]);
  payload["STATUS DO PAGAMENTO"] = _normLivro(payload["STATUS DO PAGAMENTO"]) || "Pendente";

  _validarVinculoConta(payload);

  var rowIndex = parseInt(payloadIn.rowIndex, 10);
  var values = LIVRO_DIARIO_HEADERS.map(function (h) { return payload[h] != null ? payload[h] : ""; });
  if (!isNaN(rowIndex) && rowIndex >= 2 && rowIndex <= sh.getLastRow()) {
    sh.getRange(rowIndex, 1, 1, values.length).setValues([values]);
  } else {
    sh.appendRow(values);
    rowIndex = sh.getLastRow();
  }

  var prefs = _getLivroDiarioPrefs();
  ordenarLivroDiario(prefs.modoOrdenacao);
  recalcularSaldosLivroDiario(prefs.contaFinanceiraSaldo);
  return { ok: true, rowIndex: rowIndex };
}

function getLivroDiarioLancamentos(filtros) {
  var sh = ensureLivroDiarioSheet();
  var prefs = _getLivroDiarioPrefs();
  if (sh.getLastRow() < 2) return { rows: [], preferencias: prefs };
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var f = filtros || {};
  var fCliente = _normLivro(f.cliente).toLowerCase();
  var fProjeto = _normLivro(f.projeto).toLowerCase();
  var fStatus = _normLivro(f.status).toLowerCase();
  var fConta = _normLivro(f.contaFinanceira).toLowerCase();
  var rows = data.map(function (r, i) {
    var obj = { rowIndex: i + 2 };
    headers.forEach(function (h, c) { obj[h] = r[c]; });
    obj["DATA COMPETÊNCIA"] = _formatDateLivro2y(obj["DATA COMPETÊNCIA"]);
    obj["DATA VENCIMENTO"] = _formatDateLivro2y(obj["DATA VENCIMENTO"]);
    obj["DATA PAGAMENTO"] = _formatDateLivro2y(obj["DATA PAGAMENTO"]);
    obj["DATA DA ÚLTIMA MOFIFICAÇÃO"] = _formatDateLivro2y(obj["DATA DA ÚLTIMA MOFIFICAÇÃO"]);
    obj["VALOR"] = _numLivro(obj["VALOR"]);
    obj["SALDO CONTA FINANCEIRA"] = _numLivro(obj["SALDO CONTA FINANCEIRA"]);
    obj["SALDO GERAL"] = _numLivro(obj["SALDO GERAL"]);
    return obj;
  }).filter(function (obj) {
    if (fCliente && _normLivro(obj["CLIENTE"]).toLowerCase().indexOf(fCliente) < 0) return false;
    if (fProjeto && _normLivro(obj["CÓDIGO DO PROJETO"]).toLowerCase().indexOf(fProjeto) < 0) return false;
    if (fStatus && _normLivro(obj["STATUS DO PAGAMENTO"]).toLowerCase().indexOf(fStatus) < 0) return false;
    if (fConta && _normLivro(obj["CONTA FINANCEIRA"]).toLowerCase().indexOf(fConta) < 0) return false;
    return true;
  });
  return { rows: rows, preferencias: prefs };
}

function ordenarLivroDiario(modo) {
  var sh = ensureLivroDiarioSheet();
  if (sh.getLastRow() < 3) return { ok: true, rows: Math.max(0, sh.getLastRow() - 1) };
  var selectedMode = (modo === "data_vencimento") ? "data_vencimento" : "data_pagamento";
  _setLivroDiarioPrefs({ modoOrdenacao: selectedMode });

  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idxDataVenc = _findHeaderIndexGeneric(headers, "DATA VENCIMENTO");
  var idxDataPag = _findHeaderIndexGeneric(headers, "DATA PAGAMENTO");
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();

  data.sort(function (a, b) {
    var kpA = _dateKeyLivro(a[idxDataPag]);
    var kpB = _dateKeyLivro(b[idxDataPag]);
    var kvA = _dateKeyLivro(a[idxDataVenc]);
    var kvB = _dateKeyLivro(b[idxDataVenc]);
    var hasPagA = kpA != null;
    var hasPagB = kpB != null;

    if (selectedMode === "data_pagamento") {
      if (hasPagA && hasPagB) {
        if (kpA !== kpB) return kpA - kpB;
        if ((kvA || 99999999) !== (kvB || 99999999)) return (kvA || 99999999) - (kvB || 99999999);
        return 0;
      }
      if (hasPagA && !hasPagB) return -1;
      if (!hasPagA && hasPagB) return 1;
      if ((kvA || 99999999) !== (kvB || 99999999)) return (kvA || 99999999) - (kvB || 99999999);
      return 0;
    }

    if ((kvA || 99999999) !== (kvB || 99999999)) return (kvA || 99999999) - (kvB || 99999999);
    if ((kpA || 99999999) !== (kpB || 99999999)) return (kpA || 99999999) - (kpB || 99999999);
    return 0;
  });

  sh.getRange(2, 1, data.length, sh.getLastColumn()).setValues(data);
  return { ok: true, rows: data.length, modoOrdenacao: selectedMode };
}

function recalcularSaldosLivroDiario(contaFinanceiraSelecionada) {
  var sh = ensureLivroDiarioSheet();
  if (sh.getLastRow() < 2) return { ok: true, saldoContaHoje: 0, saldoGeralHoje: 0, saldoFuturoGeral: 0 };
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var idxConta = _findHeaderIndexGeneric(headers, "CONTA FINANCEIRA");
  var idxValor = _findHeaderIndexGeneric(headers, "VALOR");
  var idxSaldoConta = _findHeaderIndexGeneric(headers, "SALDO CONTA FINANCEIRA");
  var idxSaldoGeral = _findHeaderIndexGeneric(headers, "SALDO GERAL");
  var idxDataPag = _findHeaderIndexGeneric(headers, "DATA PAGAMENTO");

  var prefs = _getLivroDiarioPrefs();
  var contaSel = _normLivro(contaFinanceiraSelecionada != null ? contaFinanceiraSelecionada : prefs.contaFinanceiraSaldo);
  if (contaFinanceiraSelecionada != null) _setLivroDiarioPrefs({ contaFinanceiraSaldo: contaSel });

  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var saldoGeral = 0;
  var saldoConta = 0;
  var todayKey = _dateKeyLivro(new Date());
  var saldoGeralHoje = 0;
  var saldoContaHoje = 0;
  var saldoFuturoGeral = 0;

  data.forEach(function (r) {
    var valor = _numLivro(r[idxValor]);
    saldoGeral += valor;
    if (contaSel && _normLivro(r[idxConta]) === contaSel) {
      saldoConta += valor;
      r[idxSaldoConta] = saldoConta;
    } else {
      r[idxSaldoConta] = contaSel ? r[idxSaldoConta] || "" : "";
    }
    r[idxSaldoGeral] = saldoGeral;
    saldoFuturoGeral += valor;

    var kp = _dateKeyLivro(r[idxDataPag]);
    if (kp != null && kp <= todayKey) {
      saldoGeralHoje += valor;
      if (contaSel && _normLivro(r[idxConta]) === contaSel) saldoContaHoje += valor;
    }
  });

  if (data.length > 0) {
    sh.getRange(2, 1, data.length, sh.getLastColumn()).setValues(data);
  }
  return {
    ok: true,
    contaFinanceiraSelecionada: contaSel,
    saldoContaHoje: saldoContaHoje,
    saldoGeralHoje: saldoGeralHoje,
    saldoFuturoGeral: saldoFuturoGeral
  };
}

function getLivroDiarioResumo(contaFinanceiraSelecionada) {
  var result = recalcularSaldosLivroDiario(contaFinanceiraSelecionada);
  return {
    contaFinanceiraSelecionada: result.contaFinanceiraSelecionada || "",
    saldoContaHoje: result.saldoContaHoje || 0,
    saldoGeralHoje: result.saldoGeralHoje || 0,
    saldoFuturoGeral: result.saldoFuturoGeral || 0
  };
}
