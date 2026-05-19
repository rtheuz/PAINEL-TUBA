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
  "STATUS DO PAGAMENTO",
  // Cadastro complementar usado no Livro Diário (mantido mínimo para evitar expansão indevida de colunas)
  "CÓDIGO DO PROJETO",
  "DESCRIÇÃO do PROJETO",
  "PARCEIRO/CLIENTE"
];

const LIVRO_DIARIO_PREFS_KEY = "LIVRO_DIARIO_PREFS_V1";
const LIVRO_DIARIO_CONTA_CONTABIL_PADRAO_PEDIDO = "TUBA Laser _ Receita _ Operacional _ Indústria";
const LIVRO_DIARIO_DESC_PADRAO_PEDIDO = "TuLa Receita Indústria";
const LIVRO_DIARIO_DATA_VENC_FICTICIA = "31/12/99";

// Layout da aba "Livro Diário":
// Linha 1 = topo de menus/saldo (Saldo Geral Hoje + dropdown de Conta Financeira)
// Linha 2 = cabeçalho
// Linha 3 = primeira linha de dados
const LIVRO_DIARIO_TOPO_ROW = 1;
const LIVRO_DIARIO_HEADER_ROW = 2;
const LIVRO_DIARIO_FIRST_DATA_ROW = 3;
const LIVRO_DIARIO_TOPO_LABEL_SALDO = "Saldo Geral Hoje";
const LIVRO_DIARIO_TOPO_LABEL_CONTA = "Conta Financeira";

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
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = _getSheetByNames_(["Livro Diário", "Livro Diario"]);
  var criadaAgora = false;
  if (!sheet) {
    sheet = spreadsheet.insertSheet("Livro Diário");
    sheet.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, LIVRO_DIARIO_HEADERS.length).setValues([LIVRO_DIARIO_HEADERS]);
    criadaAgora = true;
  } else {
    _migrarLivroDiarioParaTopoMenus_(sheet);
    _ensureLivroDiarioHeaderColumns_(sheet);
  }
  try {
    if (criadaAgora) {
      _aplicarTopoMenusLivroDiarioCompleto_(sheet);
    } else {
      _aplicarTopoMenusLivroDiario_(sheet);
    }
  } catch (eTopo) {
    Logger.log("Aplicar topo LivroDiario: " + (eTopo && eTopo.message ? eTopo.message : eTopo));
  }
  return sheet;
}

function _migrarLivroDiarioParaTopoMenus_(sheet) {
  if (!sheet) return;
  if (sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) {
    sheet.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, LIVRO_DIARIO_HEADERS.length).setValues([LIVRO_DIARIO_HEADERS]);
    return;
  }
  var row1 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var row1HasHeaders =
    _findHeaderIndexGeneric(row1, "DATA COMPETÊNCIA") >= 0 ||
    _findHeaderIndexGeneric(row1, "DATA VENCIMENTO") >= 0 ||
    _findHeaderIndexGeneric(row1, "CÓDIGO DO PROJETO") >= 0 ||
    _findHeaderIndexGeneric(row1, "VALOR") >= 0;
  if (row1HasHeaders) {
    sheet.insertRowBefore(1);
  }
}

function _ensureLivroDiarioHeaderColumns_(sheet) {
  if (!sheet) return;
  var currentHeaders = [];
  if (sheet.getLastColumn() >= 1) {
    currentHeaders = sheet.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
  }
  var hasAny = currentHeaders.some(function (v) { return _normLivro(v); });
  if (!hasAny) {
    sheet.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, LIVRO_DIARIO_HEADERS.length).setValues([LIVRO_DIARIO_HEADERS]);
    return;
  }
  var changed = false;
  LIVRO_DIARIO_HEADERS.forEach(function (h) {
    if (_findHeaderIndexGeneric(currentHeaders, h) < 0) {
      currentHeaders.push(h);
      changed = true;
    }
  });
  if (changed) {
    sheet.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, currentHeaders.length).setValues([currentHeaders]);
  }
}

function _coletarContasFinanceirasLivroDiario_(sheet) {
  var contas = {};
  try {
    if (typeof getLivroDiarioCadastro === "function") {
      var cad = getLivroDiarioCadastro();
      (cad && cad.contasFinanceiras ? cad.contasFinanceiras : []).forEach(function (c) {
        var v = _normLivro(c);
        if (v) contas[v] = true;
      });
    }
  } catch (eCad) {
    Logger.log("Coletar contas (cadastro): " + (eCad && eCad.message ? eCad.message : eCad));
  }
  try {
    if (sheet && sheet.getLastRow() >= LIVRO_DIARIO_FIRST_DATA_ROW && sheet.getLastColumn() >= 1) {
      var headers = sheet.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
      var idxConta = _findHeaderIndexGeneric(headers, "CONTA FINANCEIRA");
      if (idxConta >= 0) {
        var nRows = sheet.getLastRow() - LIVRO_DIARIO_FIRST_DATA_ROW + 1;
        var values = sheet.getRange(LIVRO_DIARIO_FIRST_DATA_ROW, idxConta + 1, nRows, 1).getValues();
        values.forEach(function (r) {
          var v = _normLivro(r[0]);
          if (v) contas[v] = true;
        });
      }
    }
  } catch (eAba) {
    Logger.log("Coletar contas (aba): " + (eAba && eAba.message ? eAba.message : eAba));
  }
  return Object.keys(contas).sort();
}

// ----------------------------------------------------------------------------------
// Detecta dinamicamente o layout do topo (linha 1), inspecionando A1:F1:
// - Procura célula com data validation de lista -> coluna do dropdown da Conta Financeira
// - Procura células com texto "Saldo Geral" / "Saldo da Conta" como labels
// - O valor de cada saldo é escrito na célula imediatamente à direita do label/dropdown
// ----------------------------------------------------------------------------------
function _detectarTopoLivroDiario_(sheet) {
  var info = {
    saldoGeralLabelCol: -1,
    saldoGeralValueCol: -1,
    saldoContaLabelCol: -1,
    contaDropdownCol: -1,
    saldoContaValueCol: -1
  };
  if (!sheet) return info;
  try {
    var maxCol = Math.max(6, sheet.getLastColumn() || 6);
    if (maxCol > 26) maxCol = 26;
    var rng = sheet.getRange(LIVRO_DIARIO_TOPO_ROW, 1, 1, maxCol);
    var vals = rng.getValues()[0];
    var rules = [];
    try { rules = rng.getDataValidations()[0]; } catch (e) { rules = []; }

    function _normTopo(v) {
      return String(v == null ? "" : v)
        .toLowerCase()
        .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
        .replace(/\s+/g, " ")
        .trim();
    }

    for (var i = 0; i < maxCol; i++) {
      var rule = rules ? rules[i] : null;
      if (rule && info.contaDropdownCol < 0) {
        try {
          var ctype = String(rule.getCriteriaType());
          if (ctype === "VALUE_IN_LIST" || ctype === "VALUE_IN_RANGE") {
            info.contaDropdownCol = i + 1;
          }
        } catch (eR) {}
      }
      var t = _normTopo(vals[i]);
      if (info.saldoGeralLabelCol < 0 && t.indexOf("saldo geral") >= 0) {
        info.saldoGeralLabelCol = i + 1;
      }
      if (info.saldoContaLabelCol < 0 && t.indexOf("saldo da conta") >= 0) {
        info.saldoContaLabelCol = i + 1;
      }
    }
  } catch (eDet) {
    Logger.log("Detectar topo: " + (eDet && eDet.message ? eDet.message : eDet));
  }

  // Fallbacks compatíveis com layout padrão (Saldo Geral em A1/B1, dropdown em D1/E1):
  if (info.saldoGeralValueCol < 0 && info.saldoGeralLabelCol > 0) {
    info.saldoGeralValueCol = info.saldoGeralLabelCol + 1;
  }
  if (info.saldoGeralValueCol < 0) info.saldoGeralValueCol = 2; // B1
  if (info.contaDropdownCol > 0) {
    info.saldoContaValueCol = info.contaDropdownCol + 1;
  }
  return info;
}

// ----------------------------------------------------------------------------------
// Aplicação "silenciosa" do topo:
// - Só garante o rótulo "Saldo Geral Hoje" em A1 quando estiver em branco
// - Mantém qualquer validação/formatação que o usuário tenha criado
// ----------------------------------------------------------------------------------
function _aplicarTopoMenusLivroDiario_(sheet) {
  if (!sheet) return;
  try {
    var a1 = sheet.getRange(LIVRO_DIARIO_TOPO_ROW, 1);
    if (_normLivro(a1.getValue()) === "") {
      a1.setValue(LIVRO_DIARIO_TOPO_LABEL_SALDO);
    }
  } catch (e) {
    Logger.log("Topo silencioso: " + (e && e.message ? e.message : e));
  }
}

// Aplicação completa: invocada manualmente (menu) ou na criação inicial.
// Cria layout padrão: A1 = "Saldo Geral Hoje", B1 = valor, C1 = dropdown contas, D1 = saldo conta.
function _aplicarTopoMenusLivroDiarioCompleto_(sheet) {
  if (!sheet) return;
  var a1 = sheet.getRange(LIVRO_DIARIO_TOPO_ROW, 1);
  if (_normLivro(a1.getValue()) === "") a1.setValue(LIVRO_DIARIO_TOPO_LABEL_SALDO);
  a1.setFontWeight("bold").setHorizontalAlignment("right");

  var b1 = sheet.getRange(LIVRO_DIARIO_TOPO_ROW, 2);
  b1.setNumberFormat('"R$" #,##0.00').setHorizontalAlignment("left");

  // C1 = dropdown da Conta Financeira
  var c1 = sheet.getRange(LIVRO_DIARIO_TOPO_ROW, 3);
  c1.setHorizontalAlignment("left");
  _atualizarValidationContas_(c1, /*forcar=*/true);

  // D1 = saldo da conta selecionada
  var d1 = sheet.getRange(LIVRO_DIARIO_TOPO_ROW, 4);
  d1.setNumberFormat('"R$" #,##0.00').setHorizontalAlignment("left");

  try { sheet.setFrozenRows(LIVRO_DIARIO_HEADER_ROW); } catch (eFr) {}
}

function _atualizarValidationContas_(cell, forcar) {
  try {
    var sheet = cell.getSheet();
    var contas = _coletarContasFinanceirasLivroDiario_(sheet);
    if (!contas || !contas.length) {
      if (forcar) {
        try { cell.clearDataValidations(); } catch (eClr) {}
      }
      return;
    }
    if (!forcar) {
      var current = cell.getDataValidation();
      var sameList = false;
      try {
        if (current) {
          var crit = current.getCriteriaValues();
          var listaAtual = (crit && crit[0]) ? crit[0] : [];
          if (Array.isArray(listaAtual) && listaAtual.length === contas.length) {
            sameList = true;
            for (var i = 0; i < listaAtual.length; i++) {
              if (_normLivro(listaAtual[i]) !== _normLivro(contas[i])) { sameList = false; break; }
            }
          }
        }
      } catch (eC) { sameList = false; }
      if (sameList) return;
    }
    var rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(contas, true)
      .setAllowInvalid(false)
      .build();
    cell.setDataValidation(rule);
  } catch (eDv) {
    Logger.log("Atualizar validation contas: " + (eDv && eDv.message ? eDv.message : eDv));
  }
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
    // Converte serial de data do Sheets sem deriva de timezone (evita +/-1 dia).
    var excelEpochUtc = Date.UTC(1899, 11, 30);
    var msUtc = excelEpochUtc + Math.round(value * 86400 * 1000);
    var dNum = new Date(msUtc);
    return isNaN(dNum.getTime()) ? null : new Date(dNum.getUTCFullYear(), dNum.getUTCMonth(), dNum.getUTCDate());
  }
  var s = String(value).trim();
  if (!s) return null;
  // Remove hora quando vier "dd/MM/yyyy HH:mm:ss" ou "yyyy-MM-ddTHH:mm:ss".
  if (s.length > 10) {
    var mIsoDateTime = s.match(/^(\d{4})-(\d{2})-(\d{2})[T\s].*$/);
    if (mIsoDateTime) s = mIsoDateTime[1] + "-" + mIsoDateTime[2] + "-" + mIsoDateTime[3];
    else if (s.indexOf(" ") > -1) s = s.split(" ")[0];
  }
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

function _parseMesPtBr_(s) {
  var t = String(s || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  var map = {
    janeiro: 1, jan: 1,
    fevereiro: 2, fev: 2,
    marco: 3, mar: 3,
    abril: 4, abr: 4,
    maio: 5, mai: 5,
    junho: 6, jun: 6,
    julho: 7, jul: 7,
    agosto: 8, ago: 8,
    setembro: 9, set: 9,
    outubro: 10, out: 10,
    novembro: 11, nov: 11,
    dezembro: 12, dez: 12
  };
  var ks = Object.keys(map);
  for (var i = 0; i < ks.length; i++) {
    if (t.indexOf(ks[i]) >= 0) return map[ks[i]];
  }
  return null;
}

function _safeSetRowValues_(sheet, rowIndex, values) {
  var rg = sheet.getRange(rowIndex, 1, 1, values.length);
  try {
    rg.setValues([values]);
  } catch (e) {
    var msg = String((e && e.message) || e || "");
    if (msg.toLowerCase().indexOf("valida") >= 0) {
      // Em algumas planilhas há validações herdadas em colunas editáveis do Livro Diário.
      // Para gravação vinda da UI, removemos validação da linha-alvo e gravamos.
      rg.clearDataValidations();
      rg.setValues([values]);
      return;
    }
    throw e;
  }
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

function _normHeaderKeyLivro_(v) {
  return String(v == null ? "" : v)
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

function _resolveLivroValueByHeader_(header, rowObj) {
  var src = rowObj || {};
  if (src[header] != null) return src[header];
  var hk = _normHeaderKeyLivro_(header);
  var aliases = {
    datacompetencia: ["DATA COMPETÊNCIA", "DATA COMPETENCIA", "DATA_COMPETENCIA"],
    datavencimento: ["DATA VENCIMENTO", "DATA_VENCIMENTO", "DATA VENC"],
    datapagamento: ["DATA PAGAMENTO", "DATA_PAGAMENTO", "DATA CAIXA"],
    cliente: ["CLIENTE", "PARCEIRO/CLIENTE", "PARCEIRO CLIENTE"],
    codigodoprojeto: ["CÓDIGO DO PROJETO", "CODIGO DO PROJETO", "PROJETO", "CÓDIGO", "CODIGO"],
    descricaodoprojeto: ["DESCRIÇÃO DO PROJETO", "DESCRICAO DO PROJETO"],
    contacontabil: ["CONTA CONTÁBIL", "CONTA CONTABIL"],
    descricaoabreviadadacontacontabil: ["DESCRICAO ABREVIADA da CONTA CONTÁBIL", "DESCRICAO ABREVIADA DA CONTA CONTABIL", "DESCRICAO ABREVIADA"],
    contafinanceira: ["CONTA FINANCEIRA"],
    meiodepagamento: ["MEIO de PAGAMENTO", "MEIO DE PAGAMENTO"],
    validadefiscal: ["VALIDADE FISCAL"],
    valor: ["VALOR", "VALOR TOTAL", "VALOR_TOTAL"],
    saldocontafinanceira: ["SALDO CONTA FINANCEIRA"],
    saldogeral: ["SALDO GERAL"],
    statusdopagamento: ["STATUS DO PAGAMENTO", "STATUS PAGAMENTO"],
    responsaveldolancamento: ["RESPONSÁVEL DO LANÇAMENTO", "RESPONSAVEL DO LANCAMENTO", "LANÇAMENTO RESPONSÁVEL"],
    datadaultimamofificacao: ["DATA DA ÚLTIMA MOFIFICAÇÃO", "DATA DA ULTIMA MODIFICACAO", "DATA da ÚLTIMA MOFIFICAÇÃO (DUM)"],
    observacoes: ["OBSERVAÇÕES", "OBSERVACOES", "OBS"]
  };
  var cands = aliases[hk] || [];
  for (var i = 0; i < cands.length; i++) {
    var k = cands[i];
    if (src[k] != null) return src[k];
  }
  var srcKeys = Object.keys(src);
  for (var j = 0; j < srcKeys.length; j++) {
    if (_normHeaderKeyLivro_(srcKeys[j]) === hk) return src[srcKeys[j]];
  }
  return "";
}

function _buildLivroRowBySheetHeaders_(sheetHeaders, rowObj) {
  var headers = Array.isArray(sheetHeaders) ? sheetHeaders : [];
  var src = rowObj || {};
  return headers.map(function (h) {
    return _resolveLivroValueByHeader_(h, src);
  });
}

function _numLivro(v) {
  if (typeof _parseCurrency === "function") return _parseCurrency(v);
  var n = Number(v);
  return isNaN(n) ? 0 : n;
}

function _saldoDeltaLivro(v) {
  // Regra de negócio:
  // VALOR negativo = receita (aumenta saldo)
  // VALOR positivo = despesa (reduz saldo)
  // impacto no saldo = -VALOR
  return -_numLivro(v);
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
  var anterior = _getLivroDiarioPrefs();
  var next = _setLivroDiarioPrefs(prefs || {});
  // Só reordena (operação destrutiva para filtros nativos) quando o usuário
  // explicitamente alterou o modo de ordenação. Mudar apenas a conta para
  // ver saldo não deve mexer na ordem das linhas nem desfazer filtros.
  var modoMudou = String(anterior.modoOrdenacao || "") !== String(next.modoOrdenacao || "");
  if (modoMudou) ordenarLivroDiario(next.modoOrdenacao);
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
      mapDescricaoParaConta: {},
      codigosProjetos: [],
      descricoesProjetos: [],
      mapCodigoParaDescricaoProjeto: {},
      mapDescricaoParaCodigoProjeto: {}
    };
  }
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  var idxConta = _findHeaderIndexGeneric(headers, "CONTA CONTÁBIL");
  var idxDesc = _findHeaderIndexGeneric(headers, "DESCRICAO ABREVIADA");
  var idxFin = _findHeaderIndexGeneric(headers, "CONTA FINANCEIRA");
  var idxMeio = _findHeaderIndexGeneric(headers, "MEIO de PAGAMENTO");
  var idxVal = _findHeaderIndexGeneric(headers, "VALIDADE FISCAL");
  var idxStatus = _findHeaderIndexGeneric(headers, "STATUS DO PAGAMENTO");
  var idxCodProj = _findHeaderIndexGeneric(headers, "CÓDIGO DO PROJETO");
  var idxDescProj = _findHeaderIndexGeneric(headers, "DESCRIÇÃO do PROJETO");

  function getAt_(row, idx) {
    if (idx == null || idx < 0) return "";
    return row[idx];
  }
  var contas = {};
  var descs = {};
  var fins = {};
  var meios = {};
  var vals = {};
  var status = {};
  var mapContaParaDescricao = {};
  var mapDescricaoParaConta = {};
  var codigosProj = {};
  var descricoesProj = {};
  var mapCodigoParaDescricaoProjeto = {};
  var mapDescricaoParaCodigoProjeto = {};
  var linhas = data.map(function (r, i) {
    var obj = {
      rowIndex: i + 2,
      contaContabil: _normLivro(getAt_(r, idxConta)),
      descricaoAbreviada: _normLivro(getAt_(r, idxDesc)),
      contaFinanceira: _normLivro(getAt_(r, idxFin)),
      meioPagamento: _normLivro(getAt_(r, idxMeio)),
      validadeFiscal: _normLivro(getAt_(r, idxVal)),
      statusPagamento: _normLivro(getAt_(r, idxStatus)),
      codigoProjeto: _normLivro(getAt_(r, idxCodProj)),
      descricaoProjeto: _normLivro(getAt_(r, idxDescProj))
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
    if (obj.codigoProjeto) codigosProj[obj.codigoProjeto] = true;
    if (obj.descricaoProjeto) descricoesProj[obj.descricaoProjeto] = true;
    if (obj.codigoProjeto && obj.descricaoProjeto && !mapCodigoParaDescricaoProjeto[obj.codigoProjeto]) {
      mapCodigoParaDescricaoProjeto[obj.codigoProjeto] = obj.descricaoProjeto;
    }
    if (obj.descricaoProjeto && obj.codigoProjeto && !mapDescricaoParaCodigoProjeto[obj.descricaoProjeto]) {
      mapDescricaoParaCodigoProjeto[obj.descricaoProjeto] = obj.codigoProjeto;
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
    mapDescricaoParaConta: mapDescricaoParaConta,
    codigosProjetos: Object.keys(codigosProj).sort(),
    descricoesProjetos: Object.keys(descricoesProj).sort(),
    mapCodigoParaDescricaoProjeto: mapCodigoParaDescricaoProjeto,
    mapDescricaoParaCodigoProjeto: mapDescricaoParaCodigoProjeto
  };
}

function upsertLivroDiarioCadastroItem(item, token) {
  var sh = ensureLivroDiarioCadastroSheet();
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  var payload = item || {};
  var vConta = _normLivro(payload.contaContabil);
  var vDesc = _normLivro(payload.descricaoAbreviada);
  var vFin = _normLivro(payload.contaFinanceira);
  var vMeio = _normLivro(payload.meioPagamento);
  var vVal = _normLivro(payload.validadeFiscal);
  var vStatus = _normLivro(payload.statusPagamento);
  var vCodProj = _normLivro(payload.codigoProjeto || payload["CÓDIGO DO PROJETO"]);
  var vDescProj = _normLivro(payload.descricaoProjeto || payload["DESCRIÇÃO do PROJETO"] || payload["DESCRIÇÃO DO PROJETO"]);

  var any = vConta || vDesc || vFin || vMeio || vVal || vStatus || vCodProj || vDescProj;
  if (!any) throw new Error("Informe ao menos um campo para cadastro.");

  var idxConta = _findHeaderIndexGeneric(headers, "CONTA CONTÁBIL");
  var idxDescAb = _findHeaderIndexGeneric(headers, "DESCRICAO ABREVIADA");
  var idxFin = _findHeaderIndexGeneric(headers, "CONTA FINANCEIRA");
  var idxMeio = _findHeaderIndexGeneric(headers, "MEIO de PAGAMENTO");
  var idxVal = _findHeaderIndexGeneric(headers, "VALIDADE FISCAL");
  var idxStatus = _findHeaderIndexGeneric(headers, "STATUS DO PAGAMENTO");
  var idxCodProj = _findHeaderIndexGeneric(headers, "CÓDIGO DO PROJETO");
  var idxDescProj = _findHeaderIndexGeneric(headers, "DESCRIÇÃO do PROJETO");

  // Monta a linha com o mesmo tamanho do cabeçalho
  var row = new Array(headers.length);
  for (var k = 0; k < row.length; k++) row[k] = "";
  if (idxConta >= 0) row[idxConta] = vConta;
  if (idxDescAb >= 0) row[idxDescAb] = vDesc;
  if (idxFin >= 0) row[idxFin] = vFin;
  if (idxMeio >= 0) row[idxMeio] = vMeio;
  if (idxVal >= 0) row[idxVal] = vVal;
  if (idxStatus >= 0) row[idxStatus] = vStatus;
  if (idxCodProj >= 0) row[idxCodProj] = vCodProj;
  if (idxDescProj >= 0) row[idxDescProj] = vDescProj;

  // Evita duplicar: se já existir uma linha com mesmo Código ou mesma Descrição (quando presentes)
  if (sh.getLastRow() >= 2 && (idxCodProj >= 0 || idxDescProj >= 0)) {
    var data = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    for (var i = 0; i < data.length; i++) {
      var curCod = idxCodProj >= 0 ? _normLivro(data[i][idxCodProj]) : "";
      var curDesc = idxDescProj >= 0 ? _normLivro(data[i][idxDescProj]) : "";
      if (vCodProj && curCod && curCod === vCodProj) {
        // se faltava descrição, completa na mesma linha
        if (vDescProj && idxDescProj >= 0 && !curDesc) {
          sh.getRange(i + 2, idxDescProj + 1).setValue(vDescProj);
          _invalidateLivroCadastroMapsCache_();
        }
        return { ok: true, rowIndex: i + 2, duplicado: true, usuario: _usuarioLancamentoPorToken(token) };
      }
      if (vDescProj && curDesc && curDesc === vDescProj) {
        if (vCodProj && idxCodProj >= 0 && !curCod) {
          sh.getRange(i + 2, idxCodProj + 1).setValue(vCodProj);
          _invalidateLivroCadastroMapsCache_();
        }
        return { ok: true, rowIndex: i + 2, duplicado: true, usuario: _usuarioLancamentoPorToken(token) };
      }
    }
  }

  sh.appendRow(row);
  _invalidateLivroCadastroMapsCache_();
  return { ok: true, rowIndex: sh.getLastRow(), duplicado: false, usuario: _usuarioLancamentoPorToken(token) };
}

/**
 * Lê CLIENTE, DESCRIÇÃO, PROCESSOS e temNotaFiscal da aba Projetos em um único
 * findRow + uma leitura de linha (evita N getValue e dois findRow no fluxo do Livro).
 */
function _getProjetoLivroMeta_(codigoProjeto) {
  var out = { cliente: "", descricaoProjeto: "", processos: "", temNotaFiscal: null };
  try {
    if (!SHEET_PROJ || !codigoProjeto) return out;
    var linha = findRowByColumnValue(SHEET_PROJ, "PROJETO", codigoProjeto);
    if (!linha) {
      var base = String(codigoProjeto).replace(/_v\d+$/i, "");
      linha = findRowByColumnValue(SHEET_PROJ, "PROJETO", base);
    }
    if (!linha) return out;
    var lastCol = SHEET_PROJ.getLastColumn();
    var headers = SHEET_PROJ.getRange(1, 1, 1, lastCol).getValues()[0];
    var row = SHEET_PROJ.getRange(linha, 1, 1, lastCol).getValues()[0];
    var idxCliente = _findHeaderIndexGeneric(headers, "CLIENTE");
    var idxDesc = _findHeaderIndexGeneric(headers, "DESCRIÇÃO");
    var idxProc = _findHeaderIndexGeneric(headers, "PROCESSOS");
    var idxJson = _findHeaderIndexGeneric(headers, "JSON_DADOS");
    if (idxCliente >= 0) out.cliente = _normLivro(row[idxCliente]);
    if (idxDesc >= 0) out.descricaoProjeto = _normLivro(row[idxDesc]);
    if (idxProc >= 0) out.processos = _normLivro(row[idxProc]);
    if (idxJson >= 0) {
      var cell = row[idxJson];
      if (cell && String(cell).trim()) {
        var parsed = null;
        try { parsed = JSON.parse(String(cell).trim()); } catch (e) { parsed = null; }
        if (parsed) {
          var dados = parsed.dados || parsed;
          var obs = dados && dados.observacoes ? dados.observacoes : {};
          if (obs && obs.temNotaFiscal != null) out.temNotaFiscal = !!obs.temNotaFiscal;
          if (!out.cliente && dados && dados.cliente && dados.cliente.nome) out.cliente = _normLivro(dados.cliente.nome);
          if (!out.descricaoProjeto && obs && obs.descricao) out.descricaoProjeto = _normLivro(obs.descricao);
          if (!out.processos && obs && obs.processos) out.processos = _normLivro(obs.processos);
          if (!out.processos && dados && Array.isArray(dados.produtosCadastrados)) {
            var procSet = {};
            dados.produtosCadastrados.forEach(function (p) {
              if (!p || typeof p !== "object") return;
              var descProc = p.descricoesProcessos || {};
              Object.keys(descProc || {}).forEach(function (k) {
                var sig = _normLivro(k).toUpperCase();
                if (sig) procSet[sig] = true;
              });
              var outros = Array.isArray(p.outrosProcessos) ? p.outrosProcessos : [];
              outros.forEach(function (op) {
                var sig2 = _normLivro(op && op.sigla).toUpperCase();
                if (sig2) procSet[sig2] = true;
              });
            });
            var siglas = Object.keys(procSet).sort();
            if (siglas.length) out.processos = siglas.join(", ");
          }
        }
      }
    }
  } catch (e) {
    // ignore
  }
  return out;
}

function _getDescricaoProjetoPorCodigo(codigoProjeto) {
  return _getProjetoLivroMeta_(codigoProjeto).descricaoProjeto || "";
}

function _getProjetoDadosBasicosPorCodigo(codigoProjeto) {
  return _getProjetoLivroMeta_(codigoProjeto);
}

function _prefixoProcessosProjeto_(processosRaw) {
  var s = _normLivro(processosRaw);
  if (!s) return "";
  // Ex.: "CL, D" -> "CL, D"
  // Mantém formato legível com vírgula no Livro Diário.
  var parts = s
    .split(/[,;|\/]+|\s{2,}/)
    .map(function (p) { return String(p || "").trim(); })
    .filter(function (p) { return !!p; });
  if (!parts.length) return "";
  var joined = parts.join(", ").toUpperCase();
  return joined;
}

function _codigoBaseProjeto_(codigo) {
  return _normLivro(codigo).replace(/_v\d+$/i, "");
}

/** Despesas manuais (material, despejo, etc.) entram com valor positivo; receita de pedido é negativa. */
function _isLivroLinhaDespesaManual_(row, idxValor) {
  if (idxValor < 0 || !row) return false;
  return _numLivro(row[idxValor]) > 0;
}

function _extrairProcessosPedidoOuProjeto_(rowPed, metaProj) {
  // 1) Prioriza o que já vem do projeto (JSON_DADOS)
  var p = _normLivro(metaProj && metaProj.processos);
  if (p) return p;

  // 2) Tenta campos explícitos da linha de pedido
  var candidatos = [
    rowPed && rowPed.PROCESSOS,
    rowPed && rowPed.PROCESSO,
    rowPed && rowPed["PROCESSOS"],
    rowPed && rowPed["PROCESSO"]
  ];
  for (var i = 0; i < candidatos.length; i++) {
    var c = _normLivro(candidatos[i]);
    if (c) return c;
  }

  // 3) Tenta extrair de OBS / OBSERVACOES no formato "processos: CL, D"
  var obs = _normLivro((rowPed && (rowPed.OBS || rowPed.OBSERVACOES || rowPed["OBSERVAÇÕES"])) || "");
  if (obs) {
    var m = obs.match(/processos?\s*[:\-]\s*([A-Za-z0-9,\s\/;|]+)/i);
    if (m && m[1]) return _normLivro(m[1]);
  }
  return "";
}

function _calcularParcelasLivroDiario_(condicoes, valorTotal, dataEntrega, dataCompetencia) {
  var cond = _normLivro(condicoes);
  var total = _numLivro(valorTotal);
  if (!cond) return null;

  // 1) tenta reaproveitar o cálculo central já existente
  if (typeof _calcularParcelasPedidos === "function") {
    try {
      var baseTry = _normLivro(dataEntrega) || _normLivro(dataCompetencia) || "";
      var p = _calcularParcelasPedidos(cond, total, baseTry);
      if (p && p.length) return p.map(function (x, i) {
        return {
          numero: i + 1,
          total: p.length,
          valor: _numLivro(x.valor),
          dataVencimento: x.dataVencimento || "",
          condicao: x.condicao || ""
        };
      });
    } catch (e) {}
  }

  // 2) fallback robusto para formatos "30 / 45 / 60"
  var dias = [];
  if (cond.indexOf("/") >= 0) {
    cond.split(/\s*\/\s*/).forEach(function (part) {
      var t = _normLivro(part).toLowerCase();
      if (!t) return;
      if (t.indexOf("vista") >= 0) dias.push(0);
      else {
        var n = parseInt(t.replace(/\D/g, ""), 10);
        if (!isNaN(n)) dias.push(n);
      }
    });
  }
  if (!dias.length) {
    var nums = cond.match(/\d+/g);
    if (nums && nums.length > 1) dias = nums.map(function (n) { return parseInt(n, 10); });
  }
  if (!dias.length) return null;

  var adiantado = /adiant/i.test(cond);
  var qtd = dias.length;
  var valorParcela = qtd > 0 ? (total / qtd) : total;
  return dias.map(function (d, i) {
    var base = (adiantado && i === 0) ? (_normLivro(dataCompetencia) || _normLivro(dataEntrega))
                                       : (_normLivro(dataEntrega) || _normLivro(dataCompetencia));
    var dt = _asDateLivro(base);
    var venc = "";
    if (dt) {
      dt.setDate(dt.getDate() + (isNaN(d) ? 0 : d));
      venc = _formatDateLivro2y(dt);
    }
    return { numero: i + 1, total: qtd, valor: valorParcela, dataVencimento: venc };
  });
}

function _parseHistoricoParcelasPedido_(rowPed) {
  var out = { byParcela: {}, totalPago: 0 };
  try {
    var raw = _normLivro(
      rowPed.PARCELAS_E_PGTOS ||
      rowPed["PARCELAS E PGTOS"] ||
      rowPed.HISTORICO_PAGAMENTOS ||
      rowPed["HISTORICO PAGAMENTOS"] ||
      rowPed["HISTÓRICO PAGAMENTOS"] ||
      rowPed["PARCELAS PAGAS"] ||
      ""
    );
    if (!raw) return out;
    var arr = JSON.parse(raw);
    if (!Array.isArray(arr)) return out;
    arr.forEach(function (x) {
      if (!x) return;
      var parcelaNum = parseInt(x.parcela, 10);
      if (isNaN(parcelaNum) || parcelaNum <= 0) return;
      var valor = _numLivro(x.valor);
      var data = _formatDateLivro2y(x.data || "");
      if (!out.byParcela[parcelaNum]) {
        out.byParcela[parcelaNum] = { valorPago: 0, dataPagamento: "" };
      }
      out.byParcela[parcelaNum].valorPago += valor;
      if (data) out.byParcela[parcelaNum].dataPagamento = data; // mantém a última data válida
      out.totalPago += valor;
    });
  } catch (e) {
    // ignora histórico inválido
  }
  return out;
}

function _buildLancamentoFromPedido(codigoProjeto, parcela, projetoInfo) {
  var dataComp = _formatDateLivro2y(projetoInfo.dataCompetencia || projetoInfo.dataEntrega || "");
  // Regra: vencimento só existe quando há data de entrega.
  var dataVenc = "";
  if (projetoInfo.dataEntrega) {
    dataVenc = _formatDateLivro2y(parcela.dataVencimento || projetoInfo.dataVencimento || projetoInfo.dataEntrega || "");
  }
  if (projetoInfo.dataEntrega && !dataVenc && parcela.condicao) {
    var cond = String(parcela.condicao).toLowerCase();
    // À vista/pedido: vencimento na própria entrega
    if (cond.indexOf("pedido") >= 0 || cond.indexOf("avista") >= 0 || cond.indexOf("a vista") >= 0) {
      dataVenc = _formatDateLivro2y(projetoInfo.dataEntrega);
    } else if (cond.indexOf("entrega") >= 0) {
      dataVenc = _formatDateLivro2y(projetoInfo.dataEntrega);
    }
  }
  // Garante visibilidade no Livro Diário mesmo sem vencimento real.
  // Quando a DATA_ENTREGA/vencto real for preenchida no Pedido,
  // a sincronização automática substitui este valor fictício.
  if (!dataVenc) dataVenc = LIVRO_DIARIO_DATA_VENC_FICTICIA;
  var obsAuto = "[AUTO_PEDIDO_PARCELA_" + parcela.numero + "_DE_" + parcela.total + "]";
  var prefix = _prefixoProcessosProjeto_(projetoInfo.processos);
  var desc = _normLivro(projetoInfo.descricaoProjeto);
  if (prefix && desc) {
    // Evita duplicar o prefixo em sincronizações subsequentes
    var pref = prefix.toUpperCase() + " - ";
    if (String(desc).toUpperCase().indexOf(pref) !== 0) {
      desc = prefix + " - " + desc;
    }
  }
  // Regra de negócio: pedido (receita) deve entrar NEGATIVO no Livro Diário.
  var valorParcelaBruto = Number(parcela.valor) || 0;
  var valorParcela = -Math.abs(valorParcelaBruto);
  var pagoInfo = parcela.pagamentoInfo || { valorPago: 0, dataPagamento: "" };
  var pagoAbs = Math.abs(_numLivro(pagoInfo.valorPago));
  var parcelaAbs = Math.abs(valorParcela);
  var pagoCompleto = parcelaAbs > 0 ? (pagoAbs >= (parcelaAbs - 0.01)) : false;
  var pagoParcial = !pagoCompleto && pagoAbs > 0;
  var statusParcela = pagoCompleto ? "Pago" : (pagoParcial ? "Pago parcialmente" : _normLivro(projetoInfo.statusPagamento || "À pagar"));
  var dataPag = pagoInfo.dataPagamento || "";
  return {
    "DATA COMPETÊNCIA": dataComp,
    "DATA VENCIMENTO": dataVenc,
    "DATA PAGAMENTO": dataPag,
    "CLIENTE": _normLivro(projetoInfo.cliente),
    "CÓDIGO DO PROJETO": _normLivro(codigoProjeto),
    "DESCRIÇÃO DO PROJETO": desc,
    "CONTA CONTÁBIL": LIVRO_DIARIO_CONTA_CONTABIL_PADRAO_PEDIDO,
    "DESCRICAO ABREVIADA da CONTA CONTÁBIL": LIVRO_DIARIO_DESC_PADRAO_PEDIDO,
    "CONTA FINANCEIRA": "",
    "MEIO de PAGAMENTO": "",
    "VALIDADE FISCAL": _normLivro(projetoInfo.validadeFiscal || ""),
    "VALOR": valorParcela,
    "SALDO CONTA FINANCEIRA": "",
    "SALDO GERAL": "",
    "STATUS DO PAGAMENTO": statusParcela,
    "RESPONSÁVEL DO LANÇAMENTO": "Sistema",
    "DATA DA ÚLTIMA MOFIFICAÇÃO": _formatDateLivro2y(new Date()),
    "OBSERVAÇÕES": obsAuto
  };
}

function sincronizarPedidosSemLivroDiarioPorMes(mesCompetencia, token, options) {
  var mes = parseInt(mesCompetencia, 10);
  if (isNaN(mes) || mes < 1 || mes > 12) mes = 4; // padrão: abril
  var opts = options || {};
  var offset = parseInt(opts.offset, 10);
  if (isNaN(offset) || offset < 0) offset = 0;
  var batchSize = parseInt(opts.batchSize, 10);
  if (isNaN(batchSize) || batchSize <= 0) batchSize = 5;
  if (batchSize > 20) batchSize = 20;

  function _mesCompetenciaFromValue_(v) {
    if (v == null || v === "") return null;
    var dNorm = _asDateLivro(v);
    if (dNorm) return dNorm.getMonth() + 1;
    if (Object.prototype.toString.call(v) === "[object Date]") {
      return isNaN(v.getTime()) ? null : (v.getMonth() + 1);
    }
    var s = String(v).trim();
    if (!s) return null;
    // Aceita datas com sufixo de hora, ex.: "24/04/2026 00:00:00"
    var mBrPrefix = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
    if (mBrPrefix) return parseInt(mBrPrefix[2], 10);
    var mBr = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (mBr) return parseInt(mBr[2], 10);
    var mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (mIso) return parseInt(mIso[2], 10);
    var d = new Date(s);
    if (!isNaN(d.getTime())) return d.getMonth() + 1;
    return null;
  }

  var pedidoMap = (typeof getPedidosSheetMap === "function") ? getPedidosSheetMap() : {};
  if (!pedidoMap || Object.keys(pedidoMap).length === 0) {
    return {
      ok: false,
      mesCompetencia: mes,
      offset: offset,
      nextOffset: offset,
      batchSize: batchSize,
      hasMore: false,
      totalPedidos: 0,
      processados: 0,
      inseridosTotal: 0,
      ignoradosExistentes: 0,
      ignoradosMes: 0,
      erros: [{ codigoProjeto: "", erro: "Nenhum pedido encontrado no mapa de pedidos. Verifique se getPedidosSheetMap() está retornando dados válidos." }]
    };
  }
  var processados = 0;
  var inseridosTotal = 0;
  var ignoradosExistentes = 0;
  var ignoradosMes = 0;
  var erros = [];

  var codigos = Object.keys(pedidoMap || {}).sort();
  var end = Math.min(offset + batchSize, codigos.length);
  for (var i = offset; i < end; i++) {
    var codigoProjeto = codigos[i];
    var rowPed = pedidoMap[codigoProjeto] || {};
    var dataComp = rowPed.DATA_COMPETENCIA || rowPed["DATA COMPETÊNCIA"] || rowPed["DATA COMPETENCIA"];
    var mesComp = _mesCompetenciaFromValue_(dataComp);
    if (mesComp == null || mesComp !== mes) {
      ignoradosMes++;
      continue;
    }

    var cod = _normLivro(codigoProjeto);
    if (!cod) continue;

    try {
      var r = gerarLancamentosLivroDiarioParaPedido(cod, token, { skipPosProcess: true });
      processados++;
      inseridosTotal += Number((r && r.inseridos) || 0);
      if (r && Number(r.inseridos || 0) === 0) ignoradosExistentes++;
    } catch (e) {
      erros.push({ codigoProjeto: cod, erro: (e && e.message) ? e.message : String(e) });
    }
  }

  var hasMore = end < codigos.length;
  if (processados > 0) {
    try {
      var prefsPos = _getLivroDiarioPrefs();
      ordenarLivroDiario(prefsPos.modoOrdenacao);
      recalcularSaldosLivroDiario(prefsPos.contaFinanceiraSaldo);
    } catch (ePos) {
      erros.push({ codigoProjeto: "", erro: "Pos-processamento do lote: " + (ePos && ePos.message ? ePos.message : ePos) });
    }
  }

  return {
    ok: true,
    mesCompetencia: mes,
    offset: offset,
    nextOffset: end,
    batchSize: batchSize,
    hasMore: hasMore,
    totalPedidos: codigos.length,
    processados: processados,
    inseridosTotal: inseridosTotal,
    ignoradosExistentes: ignoradosExistentes,
    ignoradosMes: ignoradosMes,
    erros: erros
  };
}

function gerarLancamentosLivroDiarioParaPedido(codigoProjeto, token, options) {
  var opts = options || {};
  if (!codigoProjeto) throw new Error("Código do projeto é obrigatório.");
  var rowPed = opts.rowPedOverride || null;
  if (!rowPed) {
    var pedidoMap = (typeof getPedidosSheetMap === "function") ? getPedidosSheetMap() : {};
    rowPed = pedidoMap[codigoProjeto] || pedidoMap[String(codigoProjeto).replace(/_v\d+$/i, "")];
  }
  if (!rowPed) return { ok: true, inseridos: 0, motivo: "Pedido não encontrado" };

  var valorTotal = _numLivro(
    rowPed.VALOR_TOTAL != null && rowPed.VALOR_TOTAL !== "" ? rowPed.VALOR_TOTAL :
    (rowPed["VALOR TOTAL"] != null && rowPed["VALOR TOTAL"] !== "" ? rowPed["VALOR TOTAL"] :
      (rowPed.valorTotal != null ? rowPed.valorTotal : 0))
  );
  var condicoes = _normLivro(rowPed.CONDICOES_PAGAMENTO);
  var dataComp = _normLivro(rowPed.DATA_COMPETENCIA);
  var dataEntrega = _normLivro(rowPed.DATA_ENTREGA);
  var dataVenc = _normLivro(rowPed.DATA_VENCIMENTO);
  // Regra: parcelas/vencimento baseadas na data de entrega.
  var dataBase = dataEntrega || "";
  var parcelas = _calcularParcelasLivroDiario_(condicoes, valorTotal, dataEntrega, dataComp);
  if (!parcelas || !parcelas.length) {
    parcelas = [{ numero: 1, valor: valorTotal, dataVencimento: dataEntrega ? (dataVenc || dataEntrega) : "", total: 1 }];
  } else {
    parcelas = parcelas.map(function (p) { return Object.assign({}, p); });
    parcelas.forEach(function (p, i) { p.numero = i + 1; p.total = parcelas.length; });
    // Sem data de entrega, não preenche vencimento ainda.
    if (!dataEntrega) {
      parcelas.forEach(function (p) { p.dataVencimento = ""; });
    }
  }
  if (parcelas.length === 1 && !parcelas[0].total) parcelas[0].total = 1;

  // Aplica pagamentos por parcela a partir do histórico da linha de Pedido.
  var hist = _parseHistoricoParcelasPedido_(rowPed);
  parcelas.forEach(function (p) {
    var num = parseInt(p.numero, 10);
    if (!isNaN(num) && hist.byParcela[num]) {
      p.pagamentoInfo = hist.byParcela[num];
    } else {
      p.pagamentoInfo = { valorPago: 0, dataPagamento: "" };
    }
  });

  var statusPed = _normLivro(rowPed.STATUS_PAGAMENTO || rowPed["STATUS PAGAMENTO"] || rowPed["STATUS DE PAGAMENTO"] || "");
  var metaProj = _getProjetoLivroMeta_(codigoProjeto);
  var projetoInfo = {
    cliente: rowPed.CLIENTE || "",
    descricaoProjeto: metaProj.descricaoProjeto || "",
    dataCompetencia: dataComp,
    dataEntrega: dataEntrega,
    dataVencimento: dataVenc,
    // Regra fixa para automáticos
    statusPagamento: statusPed || "À pagar"
  };
  if (!projetoInfo.cliente && metaProj.cliente) projetoInfo.cliente = metaProj.cliente;
  if (!projetoInfo.descricaoProjeto && metaProj.descricaoProjeto) projetoInfo.descricaoProjeto = metaProj.descricaoProjeto;
  if (!projetoInfo.processos) projetoInfo.processos = _extrairProcessosPedidoOuProjeto_(rowPed, metaProj);
  if (metaProj.temNotaFiscal === true) projetoInfo.validadeFiscal = "Sim";
  else if (metaProj.temNotaFiscal === false) projetoInfo.validadeFiscal = "Não";
  else {
    var nf = _normLivro(rowPed.NOTA_FISCAL || rowPed.NF || "");
    projetoInfo.validadeFiscal = (nf && nf.toUpperCase() !== "SN") ? "Sim" : "Não";
  }

  var sh = ensureLivroDiarioSheet();
  var headers = sh.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0];
  var idxProjeto = _findHeaderIndexGeneric(headers, "CÓDIGO DO PROJETO");
  var idxObs = _findHeaderIndexGeneric(headers, "OBSERVAÇÕES");
  var idxDataComp = _findHeaderIndexGeneric(headers, "DATA COMPETÊNCIA");
  var idxDataVenc = _findHeaderIndexGeneric(headers, "DATA VENCIMENTO");
  var idxDataPag = _findHeaderIndexGeneric(headers, "DATA PAGAMENTO");
  var idxCliente = _findHeaderIndexGeneric(headers, "CLIENTE");
  var idxDescProjeto = _findHeaderIndexGeneric(headers, "DESCRIÇÃO DO PROJETO");
  var idxValidade = _findHeaderIndexGeneric(headers, "VALIDADE FISCAL");
  var idxValor = _findHeaderIndexGeneric(headers, "VALOR");
  var idxStatus = _findHeaderIndexGeneric(headers, "STATUS DO PAGAMENTO");
  var idxResp = _findHeaderIndexGeneric(headers, "RESPONSÁVEL DO LANÇAMENTO");
  var idxDum = _findHeaderIndexGeneric(headers, "DATA DA ÚLTIMA MOFIFICAÇÃO");
  var existing = {};
  var manualExisting = [];
  var codigoAtual = _normLivro(codigoProjeto);
  var baseAtual = _codigoBaseProjeto_(codigoAtual);
  var allRows = [];
  if (sh.getLastRow() >= LIVRO_DIARIO_FIRST_DATA_ROW && idxProjeto >= 0 && idxObs >= 0) {
    allRows = sh.getRange(LIVRO_DIARIO_FIRST_DATA_ROW, 1, sh.getLastRow() - LIVRO_DIARIO_HEADER_ROW, sh.getLastColumn()).getValues();
    allRows.forEach(function (r, i) {
      var p = _normLivro(r[idxProjeto]);
      var pBase = _codigoBaseProjeto_(p);
      var o = _normLivro(r[idxObs]);
      if (!p || !o) return;
      // IMPORTANT: isola apenas as linhas automáticas do projeto atual.
      // Sem esse filtro, sincronizar um projeto pode remover autos de outros.
      if (!(p === codigoAtual || pBase === baseAtual)) return;
      if (o.indexOf("[AUTO_PEDIDO_PARCELA_") === 0) {
        existing[p + "|" + o] = { rowIndex: i + LIVRO_DIARIO_FIRST_DATA_ROW, row: r };
      } else {
        manualExisting.push({ rowIndex: i + LIVRO_DIARIO_FIRST_DATA_ROW, row: r });
      }
    });
  }

  var userName = _usuarioLancamentoPorToken(token);
  var toInsert = [];
  var toUpdate = [];
  var usedKeys = {};
  parcelas.forEach(function (parc) {
    if (!parc.total) parc.total = parcelas.length || 1;
    var rowObj = _buildLancamentoFromPedido(codigoProjeto, parc, projetoInfo);
    rowObj["RESPONSÁVEL DO LANÇAMENTO"] = userName || "Sistema";
    var key = rowObj["CÓDIGO DO PROJETO"] + "|" + rowObj["OBSERVAÇÕES"];
    usedKeys[key] = true;
    if (existing[key]) {
      toUpdate.push({ key: key, rowObj: rowObj, existing: existing[key] });
    } else {
      toInsert.push(_buildLivroRowBySheetHeaders_(headers, rowObj));
    }
  });

  // Compatibilidade com lançamentos manuais legados de RECEITA (sem marcador AUTO).
  // Despesas manuais (valor > 0) não entram aqui — convivem com linhas AUTO do pedido.
  var hasAutoRows = Object.keys(existing).length > 0;
  var manualReceitaLegacy = manualExisting.filter(function (m) {
    return !_isLivroLinhaDespesaManual_(m.row, idxValor);
  });
  if (!hasAutoRows && manualReceitaLegacy.length > 0) {
    toUpdate = [];
    toInsert = [];
    var limit = Math.min(parcelas.length, manualReceitaLegacy.length);
    for (var mi = 0; mi < limit; mi++) {
      var parcManual = parcelas[mi];
      if (!parcManual.total) parcManual.total = parcelas.length || 1;
      var rowObjManual = _buildLancamentoFromPedido(codigoProjeto, parcManual, projetoInfo);
      rowObjManual["RESPONSÁVEL DO LANÇAMENTO"] = userName || "Sistema";
      toUpdate.push({ rowObj: rowObjManual, existing: manualReceitaLegacy[mi], preserveObs: true, manualLegacy: true });
    }
  }

  // Atualiza linhas automáticas existentes (espelho do pedido),
  // preservando campos que o usuário completa manualmente.
  var patchedRows = [];
  toUpdate.forEach(function (u) {
    var r = u.existing.row;
    var rowObj = u.rowObj;
    // Em linhas manuais legadas, altera apenas status/pagamento e metadados.
    var isManualLegacy = !!u.manualLegacy;
    var isDespesaManual = isManualLegacy && _isLivroLinhaDespesaManual_(r, idxValor);
    if (!isManualLegacy && idxDataComp >= 0) r[idxDataComp] = rowObj["DATA COMPETÊNCIA"];
    if (!isManualLegacy && idxDataVenc >= 0) r[idxDataVenc] = rowObj["DATA VENCIMENTO"];
    // DATA PAGAMENTO só sobrescreve se vier preenchida do pedido.
    if (idxDataPag >= 0 && rowObj["DATA PAGAMENTO"]) r[idxDataPag] = rowObj["DATA PAGAMENTO"];
    if (!isManualLegacy && idxCliente >= 0) r[idxCliente] = rowObj["CLIENTE"];
    if (!isManualLegacy && idxDescProjeto >= 0) r[idxDescProjeto] = rowObj["DESCRIÇÃO DO PROJETO"];
    if (!isManualLegacy && idxValidade >= 0) r[idxValidade] = rowObj["VALIDADE FISCAL"];
    if (!isManualLegacy && idxValor >= 0) r[idxValor] = rowObj["VALOR"];
    if (idxStatus >= 0 && !isDespesaManual) r[idxStatus] = rowObj["STATUS DO PAGAMENTO"];
    if (idxResp >= 0 && !isDespesaManual) r[idxResp] = rowObj["RESPONSÁVEL DO LANÇAMENTO"];
    if (idxDum >= 0) r[idxDum] = _formatDateLivro2y(new Date());
    while (r.length < headers.length) r.push("");
    patchedRows.push({ rowIndex: u.existing.rowIndex, r: r });
  });

  // Grava atualizações em blocos contíguos (1 setValues por bloco em vez de 1 por linha).
  patchedRows.sort(function (a, b) { return a.rowIndex - b.rowIndex; });
  var writeW = headers.length;
  var pr = 0;
  while (pr < patchedRows.length) {
    var startRow = patchedRows[pr].rowIndex;
    var block = [patchedRows[pr].r];
    pr++;
    while (pr < patchedRows.length && patchedRows[pr].rowIndex === startRow + block.length) {
      block.push(patchedRows[pr].r);
      pr++;
    }
    sh.getRange(startRow, 1, block.length, writeW).setValues(block);
  }

  if (toInsert.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, toInsert.length, headers.length).setValues(toInsert);
  }

  // Remove parcelas automáticas antigas que não existem mais no pedido atual.
  var toDelete = [];
  if (hasAutoRows) {
    Object.keys(existing).forEach(function (k) {
      if (!usedKeys[k]) toDelete.push(existing[k].rowIndex);
    });
    toDelete.sort(function (a, b) { return b - a; }).forEach(function (rowIdx) {
      sh.deleteRow(rowIdx);
    });
  }

  if (!opts.skipPosProcess) {
    var prefs = _getLivroDiarioPrefs();
    // Por padrão, NÃO reordena toda a aba (operações automáticas como
    // atualização de pedido não devem mexer na ordem visível ao usuário,
    // o que resetaria filtros aplicados manualmente na planilha). O
    // recalcular saldos é cirúrgico (só M/N que mudaram).
    if (opts.fullSort) {
      ordenarLivroDiario(prefs.modoOrdenacao);
    }
    recalcularSaldosLivroDiario(prefs.contaFinanceiraSaldo);
  }

  return { ok: true, inseridos: toInsert.length, atualizados: toUpdate.length, removidos: toDelete.length, parcelas: parcelas.length };
}

function sincronizarProjetosFaltantesPedidosNoLivroDiario(token, options) {
  var opts = options || {};
  var offset = parseInt(opts.offset, 10);
  if (isNaN(offset) || offset < 0) offset = 0;
  var batchSize = parseInt(opts.batchSize, 10);
  if (isNaN(batchSize) || batchSize <= 0) batchSize = 5;
  if (batchSize > 20) batchSize = 20;
  var mesFiltro = parseInt(opts.mesCompetencia, 10);
  if (isNaN(mesFiltro) || mesFiltro < 1 || mesFiltro > 12) mesFiltro = null; // sem filtro por padrão

  function _mesCompetenciaFromValue_(v) {
    if (v == null || v === "") return null;
    if (Object.prototype.toString.call(v) === "[object Date]") {
      return isNaN(v.getTime()) ? null : (v.getMonth() + 1);
    }
    var s = String(v).trim();
    if (!s) return null;
    var mBr = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})$/);
    if (mBr) return parseInt(mBr[2], 10);
    var mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (mIso) return parseInt(mIso[2], 10);
    var d = new Date(s);
    if (!isNaN(d.getTime())) return d.getMonth() + 1;
    var mPt = _parseMesPtBr_(s);
    if (mPt != null) return mPt;
    return null;
  }

  var shLivro = ensureLivroDiarioSheet();
  var headersLivro = shLivro.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, shLivro.getLastColumn()).getValues()[0];
  var idxProjLivro = _findHeaderIndexGeneric(headersLivro, "CÓDIGO DO PROJETO");
  var existentes = {};
  if (idxProjLivro >= 0 && shLivro.getLastRow() >= LIVRO_DIARIO_FIRST_DATA_ROW) {
    var dataLivro = shLivro.getRange(LIVRO_DIARIO_FIRST_DATA_ROW, 1, shLivro.getLastRow() - LIVRO_DIARIO_HEADER_ROW, shLivro.getLastColumn()).getValues();
    dataLivro.forEach(function (r) {
      var cod = _normLivro(r[idxProjLivro]);
      var base = _codigoBaseProjeto_(cod);
      if (cod) existentes[cod] = true;
      if (base) existentes[base] = true;
    });
  }

  // Usa a mesma base da página de pedidos (getPedidos), com fallback para aba Pedidos.
  var pedRows = [];
  var seenCod = {};
  if (typeof getPedidos === "function") {
    try {
      var pedidosLista = getPedidos() || [];
      pedidosLista.forEach(function (p) {
        var codP = _normLivro(p.PROJETO || p["PROJETO"] || "");
        if (!codP) return;
        if (seenCod[codP]) return;
        seenCod[codP] = true;
        var dataCompP =
          p.DATA_COMPETENCIA ||
          p["DATA COMPETÊNCIA"] ||
          p["DATA_COMPETENCIA"] ||
          p.DATA ||
          p.DATA_ENTREGA ||
          p["DATA ENTREGA"] ||
          p.DATA_VENCIMENTO ||
          p["DATA VENCIMENTO"] ||
          "";
        pedRows.push({
          codigo: codP,
          mesCompetencia: _mesCompetenciaFromValue_(dataCompP),
          rowPed: p
        });
      });
    } catch (eGetP) {
      // segue para fallback abaixo
    }
  }

  // Sempre complementa com a aba Pedidos (fonte canônica de colunas),
  // mesmo quando getPedidos() já retornou dados.
  {
    var shPed = (typeof ensurePedidosSheet === "function") ? ensurePedidosSheet() : null;
    if (!shPed || shPed.getLastRow() < 2) {
      if (pedRows.length) {
        // segue apenas com o que veio de getPedidos()
      } else {
      return {
        ok: true, offset: offset, nextOffset: offset, batchSize: batchSize, hasMore: false, totalPedidos: 0,
        processados: 0, inseridosTotal: 0, ignoradosExistentes: 0, ignoradosMes: 0,
        erros: [{ codigoProjeto: "", erro: "Sem dados em getPedidos() e aba Pedidos vazia." }]
      };
      }
    } else {
      var pedData = shPed.getDataRange().getValues();
      var pedHeaders = pedData[0] || [];
      function _rowPedFromSheetRow_(rr) {
        function norm(x) {
          return String(x || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
        }
        var out = {};
        for (var i = 0; i < pedHeaders.length; i++) {
          out[String(pedHeaders[i] || "").trim()] = rr[i];
        }
        // campos canônicos usados no LivroDiarioService
        var map = {};
        for (var j = 0; j < pedHeaders.length; j++) map[norm(pedHeaders[j])] = rr[j];
        out.PROJETO = out.PROJETO || map["projeto"] || "";
        out.CLIENTE = out.CLIENTE || map["cliente"] || "";
        out.CONDICOES_PAGAMENTO = out.CONDICOES_PAGAMENTO || map["condicoespagamento"] || "";
        out.DATA_COMPETENCIA = out.DATA_COMPETENCIA || map["datacompetencia"] || "";
        out.DATA_ENTREGA = out.DATA_ENTREGA || map["dataentrega"] || "";
        out.DATA_VENCIMENTO = out.DATA_VENCIMENTO || map["datavencimento"] || map["datavenc"] || "";
        out.VALOR_TOTAL = out.VALOR_TOTAL || map["valortotal"] || 0;
        out.PROCESSOS = out.PROCESSOS || map["processos"] || "";
        out.STATUS_PAGAMENTO = out.STATUS_PAGAMENTO || map["statuspagamento"] || "";
        out.PARCELAS_E_PGTOS = out.PARCELAS_E_PGTOS || map["parcelasepgtos"] || map["historicopagamentos"] || "";
        out.NF = out.NF || map["nf"] || "";
        return out;
      }
      function _findPedCol_(names) {
        var norm = function (x) {
          return String(x || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").replace(/[^a-z0-9]/g, "");
        };
        var map = {};
        for (var i = 0; i < pedHeaders.length; i++) map[norm(pedHeaders[i])] = i;
        for (var j = 0; j < names.length; j++) {
          var idx = map[norm(names[j])];
          if (idx != null) return idx;
        }
        return -1;
      }
      var idxProj = _findPedCol_(["PROJETO", "Código", "Código do projeto"]);
      var idxDataComp = _findPedCol_(["DATA_COMPETENCIA", "DATA COMPETENCIA", "DATA COMPETÊNCIA"]);
      if (idxProj < 0) {
        if (!pedRows.length) {
          return {
            ok: true, offset: offset, nextOffset: offset, batchSize: batchSize, hasMore: false, totalPedidos: 0,
            processados: 0, inseridosTotal: 0, ignoradosExistentes: 0, ignoradosMes: 0,
            erros: [{ codigoProjeto: "", erro: "Coluna PROJETO não encontrada em Pedidos." }]
          };
        }
      } else {
        for (var r0 = 1; r0 < pedData.length; r0++) {
          var rr = pedData[r0];
          var codRow = _normLivro(rr[idxProj]);
          if (!codRow) continue;
          if (seenCod[codRow]) continue;
          seenCod[codRow] = true;
          var dataCompRaw = (idxDataComp >= 0 ? rr[idxDataComp] : "");
          if (!dataCompRaw) {
            var idxDataEnt = _findPedCol_(["DATA_ENTREGA", "DATA ENTREGA"]);
            var idxDataVenc0 = _findPedCol_(["DATA_VENCIMENTO", "DATA VENCIMENTO", "DATA VENC"]);
            dataCompRaw = idxDataEnt >= 0 ? rr[idxDataEnt] : (idxDataVenc0 >= 0 ? rr[idxDataVenc0] : "");
          }
          pedRows.push({
            codigo: codRow,
            mesCompetencia: _mesCompetenciaFromValue_(dataCompRaw),
            rowPed: _rowPedFromSheetRow_(rr)
          });
        }
      }
    }
  }
  pedRows.sort(function (a, b) { return a.codigo.localeCompare(b.codigo); });
  var end = Math.min(offset + batchSize, pedRows.length);
  var processados = 0;
  var inseridosTotal = 0;
  var ignoradosExistentes = 0;
  var ignoradosMes = 0;
  var erros = [];

  for (var i = offset; i < end; i++) {
    var codRaw = _normLivro(pedRows[i].codigo);
    if (!codRaw) continue;
    var baseRaw = _codigoBaseProjeto_(codRaw);
    if (mesFiltro != null) {
      var mesComp = pedRows[i].mesCompetencia;
      if (mesComp == null || mesComp !== mesFiltro) {
        ignoradosMes++;
        continue;
      }
    }

    if (existentes[codRaw] || existentes[baseRaw]) {
      ignoradosExistentes++;
      continue;
    }

    try {
      var r = gerarLancamentosLivroDiarioParaPedido(codRaw, token, { skipPosProcess: true, rowPedOverride: pedRows[i].rowPed });
      processados++;
      inseridosTotal += Number((r && r.inseridos) || 0);
      existentes[codRaw] = true;
      existentes[baseRaw] = true;
    } catch (e) {
      erros.push({ codigoProjeto: codRaw, erro: (e && e.message) ? e.message : String(e) });
    }
  }

  var hasMore = end < pedRows.length;
  if (processados > 0) {
    try {
      var prefsPos = _getLivroDiarioPrefs();
      ordenarLivroDiario(prefsPos.modoOrdenacao);
      recalcularSaldosLivroDiario(prefsPos.contaFinanceiraSaldo);
    } catch (ePos) {
      erros.push({ codigoProjeto: "", erro: "Pos-processamento do lote: " + (ePos && ePos.message ? ePos.message : ePos) });
    }
  }

  return {
    ok: true,
    offset: offset,
    nextOffset: end,
    batchSize: batchSize,
    hasMore: hasMore,
    totalPedidos: pedRows.length,
    processados: processados,
    inseridosTotal: inseridosTotal,
    ignoradosExistentes: ignoradosExistentes,
    ignoradosMes: ignoradosMes,
    erros: erros
  };
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

function salvarLivroDiarioLancamento(lancamento, token, options) {
  var sh = ensureLivroDiarioSheet();
  var sheetHeaders = sh.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0];
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
  payload["STATUS DO PAGAMENTO"] = _normLivro(payload["STATUS DO PAGAMENTO"]) || "À pagar";

  _validarVinculoConta(payload);

  var rowIndex = parseInt(payloadIn.rowIndex, 10);
  var values = _buildLivroRowBySheetHeaders_(sheetHeaders, payload);
  if (!isNaN(rowIndex) && rowIndex >= LIVRO_DIARIO_FIRST_DATA_ROW && rowIndex <= sh.getLastRow()) {
    _safeSetRowValues_(sh, rowIndex, values);
  } else {
    var nextRow = Math.max(sh.getLastRow() + 1, LIVRO_DIARIO_FIRST_DATA_ROW);
    _safeSetRowValues_(sh, nextRow, values);
    rowIndex = nextRow;
  }

  var opts = options || {};
  // Importante: NÃO disparar ordenarLivroDiario automaticamente aqui — o
  // ordenar reescreve toda a aba e reseta filtros que o usuário esteja
  // aplicando diretamente na planilha. O recalcularSaldosLivroDiario é
  // cirúrgico (atualiza só M/N que mudaram) e preserva filtros nativos.
  // Para reordenar, use o menu "Livro Diario" da planilha ou o botão
  // de recalcular M/N na página web.
  if (opts.fullSort) {
    var prefsFS = _getLivroDiarioPrefs();
    ordenarLivroDiario(prefsFS.modoOrdenacao);
    recalcularSaldosLivroDiario(prefsFS.contaFinanceiraSaldo);
  } else if (!opts.fast) {
    var prefs = _getLivroDiarioPrefs();
    recalcularSaldosLivroDiario(prefs.contaFinanceiraSaldo);
  }
  return { ok: true, rowIndex: rowIndex, row: payload };
}

function excluirLivroDiarioLancamento(rowIndex, options) {
  var sh = ensureLivroDiarioSheet();
  var idx = parseInt(rowIndex, 10);
  if (isNaN(idx) || idx < LIVRO_DIARIO_FIRST_DATA_ROW || idx > sh.getLastRow()) {
    throw new Error("Linha inválida para exclusão no Livro Diário.");
  }
  sh.deleteRow(idx);

  var opts = options || {};
  if (opts.fullSort) {
    var prefsFS = _getLivroDiarioPrefs();
    ordenarLivroDiario(prefsFS.modoOrdenacao);
    recalcularSaldosLivroDiario(prefsFS.contaFinanceiraSaldo);
  } else if (!opts.fast) {
    var prefs = _getLivroDiarioPrefs();
    recalcularSaldosLivroDiario(prefs.contaFinanceiraSaldo);
  }
  return { ok: true, rowIndex: idx };
}

function getLivroDiarioLancamentos(filtros) {
  var sh = ensureLivroDiarioSheet();
  var prefs = _getLivroDiarioPrefs();
  if (sh.getLastRow() < LIVRO_DIARIO_FIRST_DATA_ROW) return { rows: [], preferencias: prefs };
  var headers = sh.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0];
  var data = sh.getRange(LIVRO_DIARIO_FIRST_DATA_ROW, 1, sh.getLastRow() - LIVRO_DIARIO_HEADER_ROW, sh.getLastColumn()).getValues();
  var f = filtros || {};
  var fCliente = _normLivro(f.cliente).toLowerCase();
  var fProjeto = _normLivro(f.projeto).toLowerCase();
  var fStatus = _normLivro(f.status).toLowerCase();
  var fConta = _normLivro(f.contaFinanceira).toLowerCase();
  var fResp = _normLivro(f.responsavel).toLowerCase();
  var rows = data.map(function (r, i) {
    var obj = { rowIndex: i + LIVRO_DIARIO_FIRST_DATA_ROW };
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
    if (fResp && _normLivro(obj["RESPONSÁVEL DO LANÇAMENTO"]).toLowerCase().indexOf(fResp) < 0) return false;
    return true;
  });
  return { rows: rows, preferencias: prefs };
}

function ordenarLivroDiario(modo) {
  var sh = ensureLivroDiarioSheet();
  if (sh.getLastRow() < LIVRO_DIARIO_FIRST_DATA_ROW + 1) {
    return { ok: true, rows: Math.max(0, sh.getLastRow() - LIVRO_DIARIO_HEADER_ROW) };
  }
  var selectedMode = String(modo || "data_pagamento");
  if (selectedMode !== "data_vencimento" && selectedMode !== "ultima_modificacao" && selectedMode !== "ultima_modificacao_asc") {
    selectedMode = "data_pagamento";
  }
  _setLivroDiarioPrefs({ modoOrdenacao: selectedMode });

  var headers = sh.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0];
  var idxDataComp = _findHeaderIndexGeneric(headers, "DATA COMPETÊNCIA");
  var idxDataVenc = _findHeaderIndexGeneric(headers, "DATA VENCIMENTO");
  var idxDataPag = _findHeaderIndexGeneric(headers, "DATA PAGAMENTO");
  var idxDum = _findHeaderIndexGeneric(headers, "DATA DA ÚLTIMA MOFIFICAÇÃO");
  var idxProjeto = _findHeaderIndexGeneric(headers, "CÓDIGO DO PROJETO");
  var idxObs = _findHeaderIndexGeneric(headers, "OBSERVAÇÕES");
  var data = sh.getRange(LIVRO_DIARIO_FIRST_DATA_ROW, 1, sh.getLastRow() - LIVRO_DIARIO_HEADER_ROW, sh.getLastColumn()).getValues();

  // Blindagem: remove duplicados de lançamentos automáticos de parcela
  // (mesmo CÓDIGO DO PROJETO + mesmo marcador [AUTO_PEDIDO_PARCELA_X_Y] em OBSERVAÇÕES).
  if (idxProjeto >= 0 && idxObs >= 0) {
    var bestByAutoKey = {};
    var orderedKeys = [];
    function _rowScore_(r) {
      var s = 0;
      for (var c = 0; c < r.length; c++) {
        if (r[c] !== "" && r[c] != null) s++;
      }
      // Prioriza manter linha com data de pagamento preenchida.
      if (idxDataPag >= 0 && _normLivro(r[idxDataPag])) s += 10;
      // Prioriza linha com DUM mais recente.
      if (idxDum >= 0) {
        var k = _dateKeyLivro(r[idxDum]);
        if (k != null) s += (k / 1000000);
      }
      return s;
    }
    data.forEach(function (r) {
      var cod = _normLivro(r[idxProjeto]);
      var obs = _normLivro(r[idxObs]);
      var isAuto = obs.indexOf("[AUTO_PEDIDO_PARCELA_") === 0;
      if (!cod || !isAuto) return;
      var key = cod + "|" + obs;
      if (!bestByAutoKey[key]) {
        bestByAutoKey[key] = { row: r, score: _rowScore_(r) };
        orderedKeys.push(key);
        return;
      }
      var scoreNew = _rowScore_(r);
      if (scoreNew > bestByAutoKey[key].score) {
        bestByAutoKey[key] = { row: r, score: scoreNew };
      }
    });
    if (orderedKeys.length) {
      var used = {};
      var deduped = [];
      data.forEach(function (r) {
        var cod = _normLivro(r[idxProjeto]);
        var obs = _normLivro(r[idxObs]);
        var isAuto = cod && obs.indexOf("[AUTO_PEDIDO_PARCELA_") === 0;
        if (!isAuto) {
          deduped.push(r);
          return;
        }
        var key = cod + "|" + obs;
        if (used[key]) return;
        used[key] = true;
        deduped.push(bestByAutoKey[key].row);
      });
      data = deduped;
    }
  }

  data.sort(function (a, b) {
    if (selectedMode === "ultima_modificacao" || selectedMode === "ultima_modificacao_asc") {
      var kdA = idxDum >= 0 ? _dateKeyLivro(a[idxDum]) : null;
      var kdB = idxDum >= 0 ? _dateKeyLivro(b[idxDum]) : null;
      var va = kdA != null ? kdA : 0;
      var vb = kdB != null ? kdB : 0;
      if (selectedMode === "ultima_modificacao") return vb - va;
      return va - vb;
    }
    var kpA = _dateKeyLivro(a[idxDataPag]);
    var kpB = _dateKeyLivro(b[idxDataPag]);
    var kvA = _dateKeyLivro(a[idxDataVenc]);
    var kvB = _dateKeyLivro(b[idxDataVenc]);
    var hasPagA = kpA != null;
    var hasPagB = kpB != null;

    if (selectedMode === "data_vencimento") {
      if ((kvA || 99999999) !== (kvB || 99999999)) return (kvA || 99999999) - (kvB || 99999999);
      if (hasPagA && hasPagB) return kpA - kpB;
      if (hasPagA && !hasPagB) return -1;
      if (!hasPagA && hasPagB) return 1;
      return 0;
    } else {
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
  });

  // Evita deriva de timezone ao regravar linhas: persiste datas em formato BR textual.
  data.forEach(function (r) {
    if (idxDataComp >= 0) r[idxDataComp] = _formatDateLivro2y(r[idxDataComp]);
    if (idxDataVenc >= 0) r[idxDataVenc] = _formatDateLivro2y(r[idxDataVenc]);
    if (idxDataPag >= 0) r[idxDataPag] = _formatDateLivro2y(r[idxDataPag]);
    if (idxDum >= 0) r[idxDum] = _formatDateLivro2y(r[idxDum]);
  });

  var existingDataRows = sh.getLastRow() - LIVRO_DIARIO_HEADER_ROW;
  if (existingDataRows > data.length) {
    sh.getRange(data.length + LIVRO_DIARIO_FIRST_DATA_ROW, 1, existingDataRows - data.length, sh.getLastColumn()).clearContent();
  }
  if (data.length > 0) {
    sh.getRange(LIVRO_DIARIO_FIRST_DATA_ROW, 1, data.length, sh.getLastColumn()).setValues(data);
  }
  return { ok: true, rows: data.length, modoOrdenacao: selectedMode };
}

function ordenarLivroDiarioPorDataPagamento() {
  var r = ordenarLivroDiario("data_pagamento");
  var prefs = _getLivroDiarioPrefs();
  recalcularSaldosLivroDiario(prefs.contaFinanceiraSaldo);
  return r;
}

function ordenarLivroDiarioPorDataVencimento() {
  var r = ordenarLivroDiario("data_vencimento");
  var prefs = _getLivroDiarioPrefs();
  recalcularSaldosLivroDiario(prefs.contaFinanceiraSaldo);
  return r;
}

function recalcularSaldosLivroDiario(contaFinanceiraSelecionada) {
  var sh = ensureLivroDiarioSheet();
  var topo = _detectarTopoLivroDiario_(sh);
  if (sh.getLastRow() < LIVRO_DIARIO_FIRST_DATA_ROW) {
    _aplicarTopoSaldos_(sh, topo, 0, 0);
    return { ok: true, saldoContaHoje: 0, saldoGeralHoje: 0, saldoFuturoGeral: 0 };
  }
  var headers = sh.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0];
  var idxConta = _findHeaderIndexGeneric(headers, "CONTA FINANCEIRA");
  var idxValor = _findHeaderIndexGeneric(headers, "VALOR");
  var idxSaldoConta = _findHeaderIndexGeneric(headers, "SALDO CONTA FINANCEIRA");
  var idxSaldoGeral = _findHeaderIndexGeneric(headers, "SALDO GERAL");
  var idxDataPag = _findHeaderIndexGeneric(headers, "DATA PAGAMENTO");

  var prefs = _getLivroDiarioPrefs();
  var contaTopo = "";
  if (topo.contaDropdownCol > 0) {
    contaTopo = _normLivro(sh.getRange(LIVRO_DIARIO_TOPO_ROW, topo.contaDropdownCol).getValue());
  }
  // Importante: o dropdown do topo (C1) é a verdade primária. Só usamos o
  // parâmetro explícito quando o topo estiver vazio. Isso evita que chamadas
  // internas como recalcularSaldosLivroDiario(prefs.contaFinanceiraSaldo)
  // limpem a coluna M caso prefs esteja vazio.
  var paramConta = _normLivro(contaFinanceiraSelecionada);
  var contaSel = contaTopo || paramConta || _normLivro(prefs.contaFinanceiraSaldo);
  var contaSelKey = _normKeyLivro_(contaSel);
  // Persiste para que outras telas/funções vejam a mesma conta selecionada.
  if (contaSel) _setLivroDiarioPrefs({ contaFinanceiraSaldo: contaSel });

  var data = sh.getRange(LIVRO_DIARIO_FIRST_DATA_ROW, 1, sh.getLastRow() - LIVRO_DIARIO_HEADER_ROW, sh.getLastColumn()).getValues();
  var saldoGeral = 0;
  var saldoConta = 0;
  var todayKey = _dateKeyLivro(new Date());
  var saldoGeralHoje = 0;
  var saldoContaHoje = 0;
  var saldoFuturoGeral = 0;
  var outSaldoConta = [];
  var outSaldoGeral = [];

  data.forEach(function (r) {
    var delta = _saldoDeltaLivro(r[idxValor]);
    saldoGeral += delta;
    var contaRowKey = _normKeyLivro_(r[idxConta]);
    if (contaSelKey && contaRowKey === contaSelKey) {
      saldoConta += delta;
      outSaldoConta.push([saldoConta]);
    } else {
      // Evita manter saldo antigo/stale de outro filtro de conta selecionada.
      outSaldoConta.push([""]);
    }
    // Coluna N: saldo geral acumulado linha a linha (impacto = -VALOR).
    outSaldoGeral.push([saldoGeral]);
    saldoFuturoGeral += delta;

    var kp = _dateKeyLivro(r[idxDataPag]);
    if (kp != null && kp <= todayKey) {
      saldoGeralHoje += delta;
      if (contaSelKey && contaRowKey === contaSelKey) saldoContaHoje += delta;
    }
  });

  // Escrita cirúrgica: só atualiza células das colunas M (SALDO CONTA FINANCEIRA)
  // e N (SALDO GERAL) que realmente mudaram. Isso preserva filtros básicos /
  // views de filtro que o usuário aplicar na aba: o Sheets só re-aplica o
  // filtro quando os valores das células dentro do range filtrado mudam,
  // então minimizar escritas evita resets de filtro perceptíveis.
  if (data.length > 0) {
    var atualConta = null;
    var atualGeral = null;
    if (idxSaldoConta >= 0 || idxSaldoGeral >= 0) {
      atualConta = idxSaldoConta >= 0 ? [] : null;
      atualGeral = idxSaldoGeral >= 0 ? [] : null;
      for (var ri = 0; ri < data.length; ri++) {
        var rw = data[ri];
        if (atualConta) atualConta.push([rw[idxSaldoConta]]);
        if (atualGeral) atualGeral.push([rw[idxSaldoGeral]]);
      }
    }
    if (idxSaldoConta >= 0 && atualConta) {
      _gravarColunaSeMudou_(sh, LIVRO_DIARIO_FIRST_DATA_ROW, idxSaldoConta + 1, atualConta, outSaldoConta);
    }
    if (idxSaldoGeral >= 0 && atualGeral) {
      _gravarColunaSeMudou_(sh, LIVRO_DIARIO_FIRST_DATA_ROW, idxSaldoGeral + 1, atualGeral, outSaldoGeral);
    }
  }

  _aplicarTopoSaldos_(sh, topo, saldoGeralHoje, saldoContaHoje);

  return {
    ok: true,
    contaFinanceiraSelecionada: contaSel,
    saldoContaHoje: saldoContaHoje,
    saldoGeralHoje: saldoGeralHoje,
    saldoFuturoGeral: saldoFuturoGeral
  };
}

/**
 * Grava em uma coluna apenas as células cujo valor mudou, em "runs" contínuos
 * (intervalos consecutivos de linhas alteradas). Evita o setValues massivo
 * que poderia desfazer/reaplicar filtros nativos do Sheets na aba quando o
 * usuário está com filtros aplicados manualmente.
 *
 * @param {Sheet} sh - aba alvo
 * @param {number} firstRow - linha da primeira célula do range
 * @param {number} col - coluna 1-based
 * @param {Array<Array<*>>} atual - valores atualmente lidos do range (Nx1)
 * @param {Array<Array<*>>} novo - valores desejados (Nx1)
 */
function _gravarColunaSeMudou_(sh, firstRow, col, atual, novo) {
  if (!sh || !Array.isArray(atual) || !Array.isArray(novo)) return;
  var n = Math.min(atual.length, novo.length);
  function normalize(v) {
    if (v == null || v === "") return "";
    if (typeof v === "number") return v;
    var num = Number(v);
    if (!isNaN(num) && isFinite(num) && String(v).trim() !== "") return num;
    return String(v).trim();
  }
  var i = 0;
  while (i < n) {
    var atualVal = normalize(atual[i] && atual[i][0]);
    var novoVal = normalize(novo[i] && novo[i][0]);
    if (atualVal === novoVal) { i++; continue; }
    var runStart = i;
    var runValues = [];
    while (i < n) {
      var a = normalize(atual[i] && atual[i][0]);
      var b = normalize(novo[i] && novo[i][0]);
      if (a === b) break;
      runValues.push([novo[i][0]]);
      i++;
    }
    if (runValues.length > 0) {
      sh.getRange(firstRow + runStart, col, runValues.length, 1).setValues(runValues);
    }
  }
}

function _aplicarTopoSaldos_(sh, topo, saldoGeralHoje, saldoContaHoje) {
  if (!sh) return;
  try {
    if (topo && topo.saldoGeralValueCol > 0) {
      var cellGeral = sh.getRange(LIVRO_DIARIO_TOPO_ROW, topo.saldoGeralValueCol);
      cellGeral.setValue(Number(saldoGeralHoje) || 0);
      var nfG = String(cellGeral.getNumberFormat() || "");
      if (!nfG || nfG === "0" || nfG === "General" || nfG.toLowerCase().indexOf("text") >= 0) {
        cellGeral.setNumberFormat('"R$" #,##0.00');
      }
    }
  } catch (e1) {
    Logger.log("Aplicar topo saldo geral: " + (e1 && e1.message ? e1.message : e1));
  }
  try {
    if (topo && topo.saldoContaValueCol > 0) {
      var cellConta = sh.getRange(LIVRO_DIARIO_TOPO_ROW, topo.saldoContaValueCol);
      cellConta.setValue(Number(saldoContaHoje) || 0);
      var nfC = String(cellConta.getNumberFormat() || "");
      if (!nfC || nfC === "0" || nfC === "General" || nfC.toLowerCase().indexOf("text") >= 0) {
        cellConta.setNumberFormat('"R$" #,##0.00');
      }
    }
  } catch (e2) {
    Logger.log("Aplicar topo saldo conta: " + (e2 && e2.message ? e2.message : e2));
  }
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

/**
 * Extrato realizado de uma conta financeira: só lançamentos com DATA PAGAMENTO,
 * até a última data de pagamento preenchida na conta (sem ultrapassar hoje).
 * Mesma regra do KPI "Saldo conta até fim" na tela do Livro Diário.
 *
 * @param {string} contaFinanceira
 * @param {number} [ano] - se informado, retorna breakdown mês a mês do ano
 */
function calcularExtratoContaFinanceira(contaFinanceira, ano) {
  var contaSel = _normLivro(contaFinanceira);
  var contaKey = _normKeyLivro_(contaSel);
  if (!contaKey) throw new Error("Informe a conta financeira.");

  var sh = ensureLivroDiarioSheet();
  if (sh.getLastRow() < LIVRO_DIARIO_FIRST_DATA_ROW) {
    return {
      ok: true,
      contaFinanceira: contaSel,
      ultimaDataPagamento: "",
      ultimaDataPagamentoKey: null,
      saldoRealizado: 0,
      saldoAntesAno: 0,
      meses: []
    };
  }

  var headers = sh.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0];
  var idxConta = _findHeaderIndexGeneric(headers, "CONTA FINANCEIRA");
  var idxValor = _findHeaderIndexGeneric(headers, "VALOR");
  var idxDataPag = _findHeaderIndexGeneric(headers, "DATA PAGAMENTO");
  var idxDataVenc = _findHeaderIndexGeneric(headers, "DATA VENCIMENTO");
  var data = sh.getRange(LIVRO_DIARIO_FIRST_DATA_ROW, 1, sh.getLastRow() - LIVRO_DIARIO_HEADER_ROW, sh.getLastColumn()).getValues();

  var maxK = null;
  data.forEach(function (r) {
    if (_normKeyLivro_(r[idxConta]) !== contaKey) return;
    var kp = _dateKeyLivro(r[idxDataPag]);
    if (kp != null && (maxK == null || kp > maxK)) maxK = kp;
  });

  var todayKey = _dateKeyLivro(new Date());
  var cutoff = maxK;
  if (cutoff != null && todayKey != null && cutoff > todayKey) cutoff = todayKey;

  var realizadas = [];
  data.forEach(function (r, i) {
    if (_normKeyLivro_(r[idxConta]) !== contaKey) return;
    var kp = _dateKeyLivro(r[idxDataPag]);
    if (kp == null || cutoff == null || kp > cutoff) return;
    realizadas.push({
      rowIndex: i + LIVRO_DIARIO_FIRST_DATA_ROW,
      dataPagamentoKey: kp,
      dataVencimentoKey: _dateKeyLivro(r[idxDataVenc]),
      delta: _saldoDeltaLivro(r[idxValor])
    });
  });

  realizadas.sort(function (a, b) {
    if (a.dataPagamentoKey !== b.dataPagamentoKey) return a.dataPagamentoKey - b.dataPagamentoKey;
    var kvA = a.dataVencimentoKey != null ? a.dataVencimentoKey : 99999999;
    var kvB = b.dataVencimentoKey != null ? b.dataVencimentoKey : 99999999;
    return kvA - kvB;
  });

  var anoNum = parseInt(ano, 10);
  var temAno = !isNaN(anoNum) && anoNum >= 2000 && anoNum <= 2100;
  var keyJanAno = temAno ? (anoNum * 10000 + 101) : null;

  var saldoAntesAno = 0;
  var saldoRealizado = 0;
  realizadas.forEach(function (item) {
    saldoRealizado += item.delta;
    if (keyJanAno != null && item.dataPagamentoKey < keyJanAno) {
      saldoAntesAno += item.delta;
    }
  });

  var mesNomes = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"];
  var meses = [];
  var acum = saldoAntesAno;

  if (temAno) {
    for (var m = 1; m <= 12; m++) {
      var ent = 0, sai = 0, qtd = 0;
      realizadas.forEach(function (item) {
        var y = Math.floor(item.dataPagamentoKey / 10000);
        var mo = Math.floor((item.dataPagamentoKey % 10000) / 100);
        if (y !== anoNum || mo !== m) return;
        qtd++;
        if (item.delta >= 0) ent += item.delta;
        else sai += Math.abs(item.delta);
        acum += item.delta;
      });
      meses.push({
        label: mesNomes[m - 1] + "/" + anoNum,
        mes: m,
        qtd: qtd,
        entradas: ent,
        saidas: sai,
        liquido: ent - sai,
        saldoAcumulado: acum
      });
    }
  }

  return {
    ok: true,
    contaFinanceira: contaSel,
    ultimaDataPagamentoKey: cutoff,
    ultimaDataPagamento: _formatDateKeyLivro_(cutoff),
    saldoRealizado: saldoRealizado,
    saldoAntesAno: saldoAntesAno,
    saldoFimAno: temAno ? acum : saldoRealizado,
    meses: meses
  };
}

function _formatDateKeyLivro_(k) {
  if (k == null) return "";
  var y = Math.floor(k / 10000);
  var mo = Math.floor((k % 10000) / 100);
  var d = k % 100;
  if (!y || !mo || !d) return "";
  return Utilities.formatDate(new Date(y, mo - 1, d), "America/Sao_Paulo", "dd/MM/yyyy");
}

/**
 * Lançamentos realizados (DATA PAGAMENTO preenchida), até a última data de
 * pagamento da conta ou geral, sem ultrapassar hoje.
 *
 * @param {Array<Object>} rows - objetos de getLivroDiarioLancamentos
 * @param {Object} [opts] - { contaFinanceira, ateDataKey }
 */
function prepararLancamentosRealizadosLinhas(rows, opts) {
  opts = opts || {};
  var contaKey = opts.contaFinanceira ? _normKeyLivro_(_normLivro(opts.contaFinanceira)) : null;
  var todayKey = _dateKeyLivro(new Date());
  var maxK = null;
  (rows || []).forEach(function (row) {
    if (contaKey && _normKeyLivro_(row["CONTA FINANCEIRA"]) !== contaKey) return;
    var kp = _dateKeyLivro(row["DATA PAGAMENTO"]);
    if (kp != null && (maxK == null || kp > maxK)) maxK = kp;
  });
  var cutoff = maxK;
  if (cutoff != null && todayKey != null && cutoff > todayKey) cutoff = todayKey;
  if (opts.ateDataKey != null && cutoff != null && opts.ateDataKey < cutoff) cutoff = opts.ateDataKey;

  var linhas = (rows || []).filter(function (row) {
    if (contaKey && _normKeyLivro_(row["CONTA FINANCEIRA"]) !== contaKey) return false;
    var kp = _dateKeyLivro(row["DATA PAGAMENTO"]);
    return kp != null && cutoff != null && kp <= cutoff;
  });
  linhas.sort(function (a, b) {
    var kpA = _dateKeyLivro(a["DATA PAGAMENTO"]);
    var kpB = _dateKeyLivro(b["DATA PAGAMENTO"]);
    if (kpA !== kpB) return kpA - kpB;
    var kvA = _dateKeyLivro(a["DATA VENCIMENTO"]) || 99999999;
    var kvB = _dateKeyLivro(b["DATA VENCIMENTO"]) || 99999999;
    return kvA - kvB;
  });
  return {
    linhas: linhas,
    cutoffKey: cutoff,
    ultimaDataPagamento: _formatDateKeyLivro_(cutoff)
  };
}

function somarSaldoRealizadoLinhas(linhas) {
  var s = 0;
  (linhas || []).forEach(function (row) {
    s += _saldoDeltaLivro(row["VALOR"]);
  });
  return s;
}

function _normKeyLivro_(v) {
  return String(v == null ? "" : v)
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function _buildLivroCadastroMapsNormalized_() {
  var cad = getLivroDiarioCadastro();
  var maps = {
    codToDesc: {},
    descToCod: {},
    contaToDescAbrev: {},
    descAbrevToConta: {}
  };
  (cad.linhas || []).forEach(function (l) {
    var cod = _normLivro(l.codigoProjeto);
    var descProj = _normLivro(l.descricaoProjeto);
    var conta = _normLivro(l.contaContabil);
    var descAbrev = _normLivro(l.descricaoAbreviada);

    if (cod && descProj && !maps.codToDesc[_normKeyLivro_(cod)]) {
      maps.codToDesc[_normKeyLivro_(cod)] = descProj;
    }
    if (descProj && cod && !maps.descToCod[_normKeyLivro_(descProj)]) {
      maps.descToCod[_normKeyLivro_(descProj)] = cod;
    }
    if (conta && descAbrev && !maps.contaToDescAbrev[_normKeyLivro_(conta)]) {
      maps.contaToDescAbrev[_normKeyLivro_(conta)] = descAbrev;
    }
    if (descAbrev && conta && !maps.descAbrevToConta[_normKeyLivro_(descAbrev)]) {
      maps.descAbrevToConta[_normKeyLivro_(descAbrev)] = conta;
    }
  });
  return maps;
}

function _invalidateLivroCadastroMapsCache_() {
  try {
    CacheService.getScriptCache().remove("LIVRO_DIARIO_CAD_MAPS_V1");
  } catch (eInv) { /* noop */ }
}

function _getLivroCadastroMapsCached_() {
  try {
    var cache = CacheService.getScriptCache();
    var key = "LIVRO_DIARIO_CAD_MAPS_V1";
    var cached = cache.get(key);
    if (cached) return JSON.parse(cached);
    var maps = _buildLivroCadastroMapsNormalized_();
    // TTL curto: reduz leituras e ainda se atualiza rápido após mudanças
    cache.put(key, JSON.stringify(maps), 30);
    return maps;
  } catch (e) {
    return _buildLivroCadastroMapsNormalized_();
  }
}

function _syncLivroDiarioLinhaComCadastro_(sheet, row, headers, maps, editedColStart, editedColEnd) {
  if (row < 2) return false;
  var idxCod = _findHeaderIndexGeneric(headers, "CÓDIGO DO PROJETO");
  var idxDescProj = _findHeaderIndexGeneric(headers, "DESCRIÇÃO DO PROJETO");
  var idxConta = _findHeaderIndexGeneric(headers, "CONTA CONTÁBIL");
  var idxDescAbrev = _findHeaderIndexGeneric(headers, "DESCRICAO ABREVIADA da CONTA CONTÁBIL");
  if (idxCod < 0 && idxDescProj < 0 && idxConta < 0 && idxDescAbrev < 0) return false;

  var relevantCols = [idxCod, idxDescProj, idxConta, idxDescAbrev]
    .filter(function (i) { return i >= 0; })
    .map(function (i) { return i + 1; }); // 1-based
  if (!relevantCols.length) return false;

  // Só processa se a edição tocou colunas relevantes
  var touchedRelevant = relevantCols.some(function (c) {
    return c >= editedColStart && c <= editedColEnd;
  });
  if (!touchedRelevant) return false;

  var minCol = Math.min.apply(null, relevantCols);
  var maxCol = Math.max.apply(null, relevantCols);
  var seg = sheet.getRange(row, minCol, 1, maxCol - minCol + 1).getValues()[0];

  function getCellByIdx0(idx0) {
    if (idx0 < 0) return "";
    return seg[idx0 + 1 - minCol];
  }

  var cod = idxCod >= 0 ? _normLivro(getCellByIdx0(idxCod)) : "";
  var descProj = idxDescProj >= 0 ? _normLivro(getCellByIdx0(idxDescProj)) : "";
  var conta = idxConta >= 0 ? _normLivro(getCellByIdx0(idxConta)) : "";
  var descAbrev = idxDescAbrev >= 0 ? _normLivro(getCellByIdx0(idxDescAbrev)) : "";

  var updates = [];

  // Projeto: Código -> Descrição
  if (idxCod >= 0 && idxDescProj >= 0 && cod) {
    var d = maps.codToDesc[_normKeyLivro_(cod)];
    if (d && d !== descProj) {
      descProj = d;
      updates.push({ col: idxDescProj + 1, value: d });
    }
  }
  // Projeto: Descrição -> Código
  if (idxCod >= 0 && idxDescProj >= 0 && descProj) {
    var c = maps.descToCod[_normKeyLivro_(descProj)];
    if (c && c !== cod) {
      cod = c;
      updates.push({ col: idxCod + 1, value: c });
    }
  }

  // Conta: Conta Contábil -> Descrição abreviada
  if (idxConta >= 0 && idxDescAbrev >= 0 && conta) {
    var da = maps.contaToDescAbrev[_normKeyLivro_(conta)];
    if (da && da !== descAbrev) {
      descAbrev = da;
      updates.push({ col: idxDescAbrev + 1, value: da });
    }
  }
  // Conta: Descrição abreviada -> Conta Contábil
  if (idxConta >= 0 && idxDescAbrev >= 0 && descAbrev) {
    var cc = maps.descAbrevToConta[_normKeyLivro_(descAbrev)];
    if (cc && cc !== conta) {
      conta = cc;
      updates.push({ col: idxConta + 1, value: cc });
    }
  }

  // Remove duplicidade de update na mesma coluna (fica o último valor calculado)
  var dedup = {};
  updates.forEach(function (u) { dedup[u.col] = u.value; });
  var colsNum = Object.keys(dedup).map(Number).sort(function (a, b) { return a - b; });
  if (!colsNum.length) return false;
  // 1 setValues por bloco de colunas consecutivas (em vez de 1 setValue por célula).
  var ci = 0;
  while (ci < colsNum.length) {
    var startCol = colsNum[ci];
    var blockVals = [dedup[startCol]];
    ci++;
    while (ci < colsNum.length && colsNum[ci] === colsNum[ci - 1] + 1) {
      blockVals.push(dedup[colsNum[ci]]);
      ci++;
    }
    sheet.getRange(row, startCol, 1, blockVals.length).setValues([blockVals]);
  }
  return true;
}

function onEdit(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var name = _normalizeSheetKey_(sheet.getName());
    if (name !== _normalizeSheetKey_("Livro Diário") && name !== _normalizeSheetKey_("Livro Diario")) return;

    var r0 = e.range.getRow();
    var numRows = e.range.getNumRows();
    var c0 = e.range.getColumn();
    var c1 = c0 + e.range.getNumColumns() - 1;

    // Topo: edição na coluna do dropdown de Conta Financeira recalcula saldos para a conta selecionada.
    if (r0 === LIVRO_DIARIO_TOPO_ROW) {
      var topoEdit = _detectarTopoLivroDiario_(sheet);
      if (topoEdit.contaDropdownCol > 0 && c0 <= topoEdit.contaDropdownCol && c1 >= topoEdit.contaDropdownCol) {
        var contaTopo = _normLivro(sheet.getRange(LIVRO_DIARIO_TOPO_ROW, topoEdit.contaDropdownCol).getValue());
        recalcularSaldosLivroDiario(contaTopo);
        return;
      }
      // Outras edições na linha 1 (formatação, labels, etc.) — não processa demais regras.
      return;
    }

    // Bloqueia processamento na linha de topo e linha do cabeçalho.
    if (r0 < LIVRO_DIARIO_FIRST_DATA_ROW) return;

    var headers = sheet.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, sheet.getLastColumn()).getValues()[0];
    var idxDataComp = _findHeaderIndexGeneric(headers, "DATA COMPETÊNCIA");
    var idxDataVenc = _findHeaderIndexGeneric(headers, "DATA VENCIMENTO");
    var idxDataPag = _findHeaderIndexGeneric(headers, "DATA PAGAMENTO");
    var idxContaFin = _findHeaderIndexGeneric(headers, "CONTA FINANCEIRA");
    var idxValor = _findHeaderIndexGeneric(headers, "VALOR");
    var idxSaldoConta = _findHeaderIndexGeneric(headers, "SALDO CONTA FINANCEIRA");
    var maps = _getLivroCadastroMapsCached_();
    for (var r = r0; r < r0 + numRows; r++) {
      _syncLivroDiarioLinhaComCadastro_(sheet, r, headers, maps, c0, c1);
    }

    // Formatação rápida de data digitada sem separador (ex.: 100525 -> 10/05/25)
    if (numRows === 1 && e.range.getNumColumns() === 1) {
      var col0 = c0 - 1;
      var isDateCol = (col0 === idxDataComp || col0 === idxDataVenc || col0 === idxDataPag);
      if (isDateCol && e.value != null && e.value !== "") {
        var raw = String(e.value).replace(/\D/g, "");
        // Quando o usuário digita data sem separador iniciando com zero
        // (ex.: 020226), o Sheets pode entregar e.value sem o zero à esquerda (20226).
        // Recompõe o formato esperado para manter a conversão automática.
        if (raw.length === 5) raw = ("0" + raw);      // ddmmyy
        if (raw.length === 7) raw = ("0" + raw);      // ddmmyyyy
        if (raw.length >= 4 && raw.length <= 8) {
          var dia = "", mes = "", anoNum = null;
          if (raw.length === 4) {
            dia = raw.slice(0, 2); mes = raw.slice(2, 4); anoNum = new Date().getFullYear();
          } else if (raw.length === 6) {
            dia = raw.slice(0, 2); mes = raw.slice(2, 4);
            var yy = parseInt(raw.slice(4, 6), 10);
            anoNum = yy < 80 ? 2000 + yy : 1900 + yy;
          } else if (raw.length === 8) {
            dia = raw.slice(0, 2); mes = raw.slice(2, 4); anoNum = parseInt(raw.slice(4, 8), 10);
          }
          var d0 = parseInt(dia, 10), m0 = parseInt(mes, 10);
          if (!isNaN(d0) && !isNaN(m0) && d0 >= 1 && d0 <= 31 && m0 >= 1 && m0 <= 12 && anoNum) {
            var dt = new Date(anoNum, m0 - 1, d0, 12, 0, 0);
            if (!isNaN(dt.getTime())) {
              e.range.setValue(dt).setNumberFormat("dd/MM/yy");
            }
          }
        }
      }
    }

    // Edição direta na planilha que afeta saldos: recalcula automaticamente.
    var touchedSaldoCols = [idxContaFin, idxValor, idxDataPag, idxSaldoConta]
      .filter(function (i) { return i >= 0; })
      .map(function (i) { return i + 1; });
    var touchedSaldo = touchedSaldoCols.some(function (c) { return c >= c0 && c <= c1; });
    if (touchedSaldo) {
      var prefs = _getLivroDiarioPrefs();
      recalcularSaldosLivroDiario(prefs.contaFinanceiraSaldo);
    }
  } catch (err) {
    Logger.log("onEdit LivroDiario: " + (err && err.message ? err.message : err));
  }
}

function removerLancamentosLivroDiarioPorProjeto(codigoProjeto) {
  var cod = _normLivro(codigoProjeto);
  if (!cod) return { ok: true, removidos: 0 };
  var base = _codigoBaseProjeto_(cod);
  var sh = ensureLivroDiarioSheet();
  if (sh.getLastRow() < LIVRO_DIARIO_FIRST_DATA_ROW) return { ok: true, removidos: 0 };
  var headers = sh.getRange(LIVRO_DIARIO_HEADER_ROW, 1, 1, sh.getLastColumn()).getValues()[0];
  var idxProjeto = _findHeaderIndexGeneric(headers, "CÓDIGO DO PROJETO");
  if (idxProjeto < 0) return { ok: true, removidos: 0 };
  var data = sh.getRange(LIVRO_DIARIO_FIRST_DATA_ROW, 1, sh.getLastRow() - LIVRO_DIARIO_HEADER_ROW, sh.getLastColumn()).getValues();
  var toDelete = [];
  data.forEach(function (r, i) {
    var p = _normLivro(r[idxProjeto]);
    if (!p) return;
    var pBase = _codigoBaseProjeto_(p);
    if (p === cod || pBase === base) toDelete.push(i + LIVRO_DIARIO_FIRST_DATA_ROW);
  });
  toDelete.sort(function (a, b) { return b - a; }).forEach(function (rowIdx) { sh.deleteRow(rowIdx); });
  return { ok: true, removidos: toDelete.length };
}

function recalcularLivroDiarioAgora() {
  var prefs = _getLivroDiarioPrefs();
  return recalcularSaldosLivroDiario(prefs.contaFinanceiraSaldo);
}

function onOpenMenuLivroDiario() {
  try {
    var ui = SpreadsheetApp.getUi();
    ui.createMenu("Livro Diario")
      .addItem("📅 Ordenar por data de pagamento", "ordenarLivroDiarioPorDataPagamento")
      .addItem("📅 Ordenar por data de vencimento", "ordenarLivroDiarioPorDataVencimento")
      .addSeparator()
      .addItem("💰 Recalcular saldos M/N agora", "recalcularLivroDiarioAgora")
      .addItem("⬇️ Aplicar/atualizar topo de menus (linha 1)", "aplicarTopoMenusLivroDiarioManual")
      .addToUi();
  } catch (e) {
    Logger.log("onOpen LivroDiario menu: " + (e && e.message ? e.message : e));
  }
}

function aplicarTopoMenusLivroDiarioManual() {
  var sh = ensureLivroDiarioSheet();
  _aplicarTopoMenusLivroDiarioCompleto_(sh);
  var prefs = _getLivroDiarioPrefs();
  recalcularSaldosLivroDiario(prefs.contaFinanceiraSaldo);
  try {
    SpreadsheetApp.getActive().toast("✅ Topo do Livro Diário atualizado.", "Livro Diario", 4);
  } catch (e) {}
  return { ok: true, sheet: sh.getName() };
}

function onOpen(e) {
  onOpenMenuLivroDiario(e);
}

function onInstall(e) {
  onOpenMenuLivroDiario(e);
}

function instalarTriggerOnOpenLivroDiario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error("Planilha ativa não encontrada para instalar trigger.");
  var ssId = ss.getId();
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    var t = triggers[i];
    var h = t.getHandlerFunction();
    if (h === "onOpen" || h === "onOpenMenuLivroDiario") {
      try { ScriptApp.deleteTrigger(t); } catch (eDel) {}
    }
  }
  ScriptApp.newTrigger("onOpenMenuLivroDiario").forSpreadsheet(ssId).onOpen().create();
  return { ok: true, spreadsheetId: ssId, triggersCriados: 1 };
}

function diagnosticarTriggerOnOpenLivroDiario() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssId = ss ? ss.getId() : "";
  var list = ScriptApp.getProjectTriggers().map(function (t) {
    return {
      funcao: t.getHandlerFunction(),
      evento: String(t.getEventType()),
      origem: String(t.getTriggerSource())
    };
  });
  return {
    spreadsheetIdAtiva: ssId,
    triggers: list
  };
}