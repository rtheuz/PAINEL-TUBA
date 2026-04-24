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
  // Cadastro complementar de projetos (para autocomplete e vínculo Código↔Descrição)
  "CÓDIGO DO PROJETO",
  "DESCRIÇÃO do PROJETO",
  "PARCEIRO/CLIENTE",
  "NATUREZA do PROJETO",
  "OBSERVAÇÕES",
  "STATUS",
  "LANÇAMENTO RESPONSÁVEL",
  "LANÇAMENTO DATA",
  "DATA da ÚLTIMA MOFIFICAÇÃO (DUM)"
];

const LIVRO_DIARIO_PREFS_KEY = "LIVRO_DIARIO_PREFS_V1";
const LIVRO_DIARIO_CONTA_CONTABIL_PADRAO_PEDIDO = "TUBA Laser _ Receita _ Operacional _ Indústria";
const LIVRO_DIARIO_DESC_PADRAO_PEDIDO = "TuLa Receita Indústria";
const LIVRO_DIARIO_DATA_VENC_FICTICIA = "31/12/99";

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
        }
        return { ok: true, rowIndex: i + 2, duplicado: true, usuario: _usuarioLancamentoPorToken(token) };
      }
      if (vDescProj && curDesc && curDesc === vDescProj) {
        if (vCodProj && idxCodProj >= 0 && !curCod) {
          sh.getRange(i + 2, idxCodProj + 1).setValue(vCodProj);
        }
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
  var out = { cliente: "", descricaoProjeto: "", processos: "", temNotaFiscal: null };
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
          if (!out.processos && obs && obs.processos) out.processos = _normLivro(obs.processos);
          // Fallback: extrai siglas de processos a partir dos produtos do JSON.
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
    // ignore and keep fallback object
  }
  return out;
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
  var pedidoMap = (typeof getPedidosSheetMap === "function") ? getPedidosSheetMap() : {};
  var rowPed = pedidoMap[codigoProjeto] || pedidoMap[String(codigoProjeto).replace(/_v\d+$/i, "")];
  if (!rowPed) return { ok: true, inseridos: 0, motivo: "Pedido não encontrado" };

  var valorTotal = _numLivro(rowPed.VALOR_TOTAL);
  var condicoes = _normLivro(rowPed.CONDICOES_PAGAMENTO);
  var dataComp = _normLivro(rowPed.DATA_COMPETENCIA);
  var dataEntrega = _normLivro(rowPed.DATA_ENTREGA);
  var dataVenc = _normLivro(rowPed.DATA_VENCIMENTO);
  // Regra: parcelas/vencimento baseadas na data de entrega.
  var dataBase = dataEntrega || "";
  var parcelas = null;
  if (typeof _calcularParcelasPedidos === "function") {
    try { parcelas = _calcularParcelasPedidos(condicoes, valorTotal, dataBase); } catch (e) { parcelas = null; }
  }
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
  var projetoInfo = {
    cliente: rowPed.CLIENTE || "",
    descricaoProjeto: _getDescricaoProjetoPorCodigo(codigoProjeto),
    dataCompetencia: dataComp,
    dataEntrega: dataEntrega,
    dataVencimento: dataVenc,
    // Regra fixa para automáticos
    statusPagamento: statusPed || "À pagar"
  };
  var metaProj = _getProjetoDadosBasicosPorCodigo(codigoProjeto);
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
  var headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
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
  var allRows = [];
  if (sh.getLastRow() >= 2 && idxProjeto >= 0 && idxObs >= 0) {
    allRows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    allRows.forEach(function (r, i) {
      var p = _normLivro(r[idxProjeto]);
      var o = _normLivro(r[idxObs]);
      if (!p || !o) return;
      if (o.indexOf("[AUTO_PEDIDO_PARCELA_") === 0) {
        existing[p + "|" + o] = { rowIndex: i + 2, row: r };
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
      toInsert.push(LIVRO_DIARIO_HEADERS.map(function (h) { return rowObj[h] != null ? rowObj[h] : ""; }));
    }
  });

  // Atualiza linhas automáticas existentes (espelho do pedido),
  // preservando campos que o usuário completa manualmente.
  toUpdate.forEach(function (u) {
    var r = u.existing.row;
    var rowObj = u.rowObj;
    if (idxDataComp >= 0) r[idxDataComp] = rowObj["DATA COMPETÊNCIA"];
    if (idxDataVenc >= 0) r[idxDataVenc] = rowObj["DATA VENCIMENTO"];
    // DATA PAGAMENTO só sobrescreve se vier preenchida do pedido.
    if (idxDataPag >= 0 && rowObj["DATA PAGAMENTO"]) r[idxDataPag] = rowObj["DATA PAGAMENTO"];
    if (idxCliente >= 0) r[idxCliente] = rowObj["CLIENTE"];
    if (idxDescProjeto >= 0) r[idxDescProjeto] = rowObj["DESCRIÇÃO DO PROJETO"];
    if (idxValidade >= 0) r[idxValidade] = rowObj["VALIDADE FISCAL"];
    if (idxValor >= 0) r[idxValor] = rowObj["VALOR"];
    if (idxStatus >= 0) r[idxStatus] = rowObj["STATUS DO PAGAMENTO"];
    if (idxResp >= 0) r[idxResp] = rowObj["RESPONSÁVEL DO LANÇAMENTO"];
    if (idxDum >= 0) r[idxDum] = _formatDateLivro2y(new Date());
    sh.getRange(u.existing.rowIndex, 1, 1, r.length).setValues([r]);
  });

  if (toInsert.length > 0) {
    sh.getRange(sh.getLastRow() + 1, 1, toInsert.length, LIVRO_DIARIO_HEADERS.length).setValues(toInsert);
  }

  // Remove parcelas automáticas antigas que não existem mais no pedido atual.
  var toDelete = [];
  Object.keys(existing).forEach(function (k) {
    if (!usedKeys[k]) toDelete.push(existing[k].rowIndex);
  });
  toDelete.sort(function (a, b) { return b - a; }).forEach(function (rowIdx) {
    sh.deleteRow(rowIdx);
  });

  if (!opts.skipPosProcess) {
    var prefs = _getLivroDiarioPrefs();
    ordenarLivroDiario(prefs.modoOrdenacao);
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
    return null;
  }

  var shLivro = ensureLivroDiarioSheet();
  var headersLivro = shLivro.getRange(1, 1, 1, shLivro.getLastColumn()).getValues()[0];
  var idxProjLivro = _findHeaderIndexGeneric(headersLivro, "CÓDIGO DO PROJETO");
  var existentes = {};
  if (idxProjLivro >= 0 && shLivro.getLastRow() >= 2) {
    var dataLivro = shLivro.getRange(2, 1, shLivro.getLastRow() - 1, shLivro.getLastColumn()).getValues();
    dataLivro.forEach(function (r) {
      var cod = _normLivro(r[idxProjLivro]);
      var base = _codigoBaseProjeto_(cod);
      if (cod) existentes[cod] = true;
      if (base) existentes[base] = true;
    });
  }

  // Usa a mesma base da página de pedidos (getPedidos), com fallback para aba Pedidos.
  var pedRows = [];
  if (typeof getPedidos === "function") {
    try {
      var pedidosLista = getPedidos() || [];
      pedidosLista.forEach(function (p) {
        var codP = _normLivro(p.PROJETO || p["PROJETO"] || "");
        if (!codP) return;
        var dataCompP = p.DATA_COMPETENCIA || p["DATA COMPETÊNCIA"] || p["DATA_COMPETENCIA"] || p.DATA || "";
        pedRows.push({
          codigo: codP,
          mesCompetencia: _mesCompetenciaFromValue_(dataCompP)
        });
      });
    } catch (eGetP) {
      // segue para fallback abaixo
    }
  }

  if (!pedRows.length) {
    var shPed = (typeof ensurePedidosSheet === "function") ? ensurePedidosSheet() : null;
    if (!shPed || shPed.getLastRow() < 2) {
      return {
        ok: true, offset: offset, nextOffset: offset, batchSize: batchSize, hasMore: false, totalPedidos: 0,
        processados: 0, inseridosTotal: 0, ignoradosExistentes: 0, ignoradosMes: 0,
        erros: [{ codigoProjeto: "", erro: "Sem dados em getPedidos() e aba Pedidos vazia." }]
      };
    }
    var pedData = shPed.getDataRange().getValues();
    var pedHeaders = pedData[0] || [];
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
      return {
        ok: true, offset: offset, nextOffset: offset, batchSize: batchSize, hasMore: false, totalPedidos: 0,
        processados: 0, inseridosTotal: 0, ignoradosExistentes: 0, ignoradosMes: 0,
        erros: [{ codigoProjeto: "", erro: "Coluna PROJETO não encontrada em Pedidos." }]
      };
    }
    for (var r0 = 1; r0 < pedData.length; r0++) {
      var rr = pedData[r0];
      var codRow = _normLivro(rr[idxProj]);
      if (!codRow) continue;
      pedRows.push({
        codigo: codRow,
        mesCompetencia: _mesCompetenciaFromValue_(idxDataComp >= 0 ? rr[idxDataComp] : "")
      });
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
      var r = gerarLancamentosLivroDiarioParaPedido(codRaw, token, { skipPosProcess: true });
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

  var opts = options || {};
  // Modo rápido para reduzir latência do modal (UI já recalcula localmente)
  if (!opts.fast) {
    var prefs = _getLivroDiarioPrefs();
    ordenarLivroDiario(prefs.modoOrdenacao);
    recalcularSaldosLivroDiario(prefs.contaFinanceiraSaldo);
  }
  return { ok: true, rowIndex: rowIndex, row: payload };
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
    var delta = _saldoDeltaLivro(r[idxValor]);
    saldoGeral += delta;
    if (contaSel && _normLivro(r[idxConta]) === contaSel) {
      saldoConta += delta;
      r[idxSaldoConta] = saldoConta;
    } else {
      r[idxSaldoConta] = contaSel ? r[idxSaldoConta] || "" : "";
    }
    r[idxSaldoGeral] = saldoGeral;
    saldoFuturoGeral += delta;

    var kp = _dateKeyLivro(r[idxDataPag]);
    if (kp != null && kp <= todayKey) {
      saldoGeralHoje += delta;
      if (contaSel && _normLivro(r[idxConta]) === contaSel) saldoContaHoje += delta;
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
  var cols = Object.keys(dedup);
  if (!cols.length) return false;
  cols.forEach(function (c) {
    sheet.getRange(row, Number(c)).setValue(dedup[c]);
  });
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
    if (r0 < 2) return;
    var c0 = e.range.getColumn();
    var c1 = c0 + e.range.getNumColumns() - 1;

    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var maps = _getLivroCadastroMapsCached_();
    for (var r = r0; r < r0 + numRows; r++) {
      _syncLivroDiarioLinhaComCadastro_(sheet, r, headers, maps, c0, c1);
    }
  } catch (err) {
    Logger.log("onEdit LivroDiario: " + (err && err.message ? err.message : err));
  }
}