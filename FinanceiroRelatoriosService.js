/**
 * Relatórios financeiros em PDF a partir do Livro Diário.
 * Salvos em: {raiz}/{AAAA}/{AAMM}/{AAMMDD}/FIN/
 */

var FIN_RELATORIOS_CACHE_LISTA_KEY = "FIN_RELATORIOS_RECENTES_V1";

function getTiposRelatorioFinanceiro() {
  return [
    { id: "resumo_executivo", nome: "Resumo executivo", grupo: "Geral", descricao: "KPIs consolidados do Livro Diário", precisaPeriodo: false, precisaReferenciaData: false, precisaContas: false, precisaContaUnica: false, precisaAno: false, envolveSaldo: true, saldoFuturo: false },
    { id: "projetos_carro_chefe", nome: "Projetos carro-chefe", grupo: "Ranking", descricao: "Top projetos — receita indústria (pedidos)", precisaPeriodo: true, precisaReferenciaData: true, precisaContas: false, precisaContaUnica: false, precisaAno: false, topN: true, somenteReceitaIndustria: true },
    { id: "clientes_carro_chefe", nome: "Clientes carro-chefe", grupo: "Ranking", descricao: "Top clientes — receita indústria (pedidos)", precisaPeriodo: true, precisaReferenciaData: true, precisaContas: false, precisaContaUnica: false, precisaAno: false, topN: true, somenteReceitaIndustria: true },
    { id: "saldo_futuro_periodo", nome: "Saldo futuro (período)", grupo: "Projeção", descricao: "Saldo realizado + projeção por vencimentos em aberto no período", precisaPeriodo: true, precisaReferenciaData: false, precisaContas: false, precisaContaUnica: false, precisaAno: false, envolveSaldo: true, saldoFuturo: true, campoDataFixo: "vencimento" },
    { id: "saldo_conta_mensal", nome: "Saldo da conta (mês a mês)", grupo: "Conta financeira", descricao: "Extrato realizado: só pagos, até a última data de pagamento da conta", precisaPeriodo: false, precisaReferenciaData: false, precisaContas: false, precisaContaUnica: true, precisaAno: true, envolveSaldo: true, saldoFuturo: false },
    { id: "contas_pagar_receber", nome: "A pagar × A receber", grupo: "Posição", descricao: "Totais em aberto (futuro) — sem filtro de período", precisaPeriodo: false, precisaReferenciaData: false, precisaContas: false, precisaContaUnica: false, precisaAno: false, somenteEmAberto: true },
    { id: "por_status_pagamento", nome: "Por status de pagamento", grupo: "Análise", descricao: "Valores agrupados por status (todos os lançamentos)", precisaPeriodo: false, precisaReferenciaData: false, precisaContas: false, precisaContaUnica: false, precisaAno: false, campoDataFixo: "pagamento" },
    { id: "por_conta_contabil", nome: "Por conta contábil", grupo: "Análise", descricao: "Receitas e despesas por conta contábil", precisaPeriodo: true, precisaReferenciaData: true, precisaContas: false, precisaContaUnica: false, precisaAno: false },
    { id: "por_conta_financeira", nome: "Por conta financeira", grupo: "Análise", descricao: "Movimentação por conta financeira", precisaPeriodo: true, precisaReferenciaData: true, precisaContas: false, precisaContaUnica: false, precisaAno: false },
    { id: "inadimplencia", nome: "Inadimplência / atrasados", grupo: "Cobrança", descricao: "Lançamentos vencidos e não quitados", precisaPeriodo: false, precisaReferenciaData: false, precisaContas: false, precisaContaUnica: false, precisaAno: false },
    { id: "despesas_por_projeto", nome: "Despesas por projeto", grupo: "Projetos", descricao: "Despesas por código e descrição do projeto", precisaPeriodo: true, precisaReferenciaData: true, precisaContas: false, precisaContaUnica: false, precisaAno: false },
    { id: "receitas_periodo", nome: "Receitas no período", grupo: "Receitas", descricao: "Detalhamento de receitas (pedidos)", precisaPeriodo: true, precisaReferenciaData: true, precisaContas: false, precisaContaUnica: false, precisaAno: false, somenteReceitaIndustria: true },
    { id: "extrato_lancamentos", nome: "Extrato de lançamentos", grupo: "Detalhe", descricao: "Lista completa filtrada do período", precisaPeriodo: true, precisaReferenciaData: true, precisaContas: false, precisaContaUnica: false, precisaAno: false }
  ];
}

function _finNormKey_(v) {
  return String(v == null ? "" : v)
    .toLowerCase()
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function _finContaReceitaIndustria_(row) {
  var conta = _finNormKey_(row["CONTA CONTÁBIL"]);
  var desc = _finNormKey_(row["DESCRICAO ABREVIADA da CONTA CONTÁBIL"] || row["DESCRICAO ABREVIADA DA CONTA CONTABIL"] || row["DESCRICAO ABREVIADA"]);
  var alvoConta = _finNormKey_(typeof LIVRO_DIARIO_CONTA_CONTABIL_PADRAO_PEDIDO !== "undefined"
    ? LIVRO_DIARIO_CONTA_CONTABIL_PADRAO_PEDIDO : "TUBA Laser _ Receita _ Operacional _ Indústria");
  var alvoDesc = _finNormKey_(typeof LIVRO_DIARIO_DESC_PADRAO_PEDIDO !== "undefined"
    ? LIVRO_DIARIO_DESC_PADRAO_PEDIDO : "TuLa Receita Indústria");
  return conta === alvoConta || desc === alvoDesc;
}

function _finFiltrarReceitaIndustria_(rows) {
  return rows.filter(_finContaReceitaIndustria_);
}

/** Chave yyyyMMdd da data de pagamento ou null. */
function _finKeyDataPagamento_(row) {
  var d = _asDateLivro(row["DATA PAGAMENTO"]);
  return d ? _finDateKey_(d) : null;
}

/**
 * Lançamentos realizados (mesma regra do Livro Diário / extrato).
 * @param {Array} rows
 * @param {Object} [opts] - { contaFinanceira, ateDataKey }
 */
function _finPrepararRealizados_(rows, opts) {
  if (typeof prepararLancamentosRealizadosLinhas === "function") {
    return prepararLancamentosRealizadosLinhas(rows, opts || {});
  }
  opts = opts || {};
  var contaKey = opts.contaFinanceira ? _finNormKey_(opts.contaFinanceira) : null;
  var todayKey = _finDateKey_(new Date());
  var maxK = null;
  (rows || []).forEach(function (r) {
    if (contaKey && _finNormKey_(r["CONTA FINANCEIRA"]) !== contaKey) return;
    var k = _finKeyDataPagamento_(r);
    if (k != null && (maxK == null || k > maxK)) maxK = k;
  });
  var cutoff = maxK;
  if (cutoff != null && todayKey != null && cutoff > todayKey) cutoff = todayKey;
  if (opts.ateDataKey != null && cutoff != null && opts.ateDataKey < cutoff) cutoff = opts.ateDataKey;
  var linhas = (rows || []).filter(function (r) {
    if (contaKey && _finNormKey_(r["CONTA FINANCEIRA"]) !== contaKey) return false;
    var k = _finKeyDataPagamento_(r);
    return k != null && cutoff != null && k <= cutoff;
  });
  linhas.sort(function (a, b) {
    var ka = _finKeyDataPagamento_(a);
    var kb = _finKeyDataPagamento_(b);
    if (ka !== kb) return ka - kb;
    var va = _finDateKey_(_asDateLivro(a["DATA VENCIMENTO"])) || 99999999;
    var vb = _finDateKey_(_asDateLivro(b["DATA VENCIMENTO"])) || 99999999;
    return va - vb;
  });
  return { linhas: linhas, cutoffKey: cutoff, ultimaDataPagamento: cutoff ? String(cutoff) : "" };
}

function _finSomarSaldoLinhas_(linhas) {
  if (typeof somarSaldoRealizadoLinhas === "function") {
    return somarSaldoRealizadoLinhas(linhas);
  }
  var s = 0;
  (linhas || []).forEach(function (r) { s += _finDelta_(r); });
  return s;
}

function _finSaldosRealizadosLivro_(contaFinanceira) {
  try {
    var r = getLivroDiarioResumo(contaFinanceira || "");
    return {
      saldoGeralHoje: Number(r.saldoGeralHoje) || 0,
      saldoContaHoje: Number(r.saldoContaHoje) || 0,
      saldoFuturoGeral: Number(r.saldoFuturoGeral) || 0
    };
  } catch (e) {
    return { saldoGeralHoje: 0, saldoContaHoje: 0, saldoFuturoGeral: 0 };
  }
}

function _finSaldoInicialRealizado_(contaFinanceira) {
  var s = _finSaldosRealizadosLivro_(contaFinanceira);
  return contaFinanceira ? s.saldoContaHoje : s.saldoGeralHoje;
}

function _finNotaSaldoRealizadoHtml_() {
  return "<p class=\"sub\" style=\"margin:8px 0 12px;font-size:11px;color:#64748b;\">" +
    "Saldo realizado: somente lançamentos com data de pagamento, até a última data de pagamento preenchida (sem futuro)." +
    "</p>";
}

function _finNotaSaldoFuturoHtml_() {
  return "<p class=\"sub\" style=\"margin:8px 0 12px;font-size:11px;color:#64748b;\">" +
    "Saldo inicial = realizado (data de pagamento). Saldo projetado = realizado + lançamentos em aberto no período (por vencimento)." +
    "</p>";
}

function getPastaFinDoDia_(dataRef) {
  var agora = dataRef instanceof Date && !isNaN(dataRef.getTime()) ? dataRef : new Date();
  var ano = agora.getFullYear();
  var ano2 = String(ano).slice(-2);
  var mes = String(agora.getMonth() + 1).padStart(2, "0");
  var dia = String(agora.getDate()).padStart(2, "0");
  var raiz = _getRootFolder_();
  var subAno = getOrCreateSubFolder(raiz, String(ano));
  var subMes = getOrCreateSubFolder(subAno, ano2 + mes);
  var subDia = getOrCreateSubFolder(subMes, ano2 + mes + dia);
  return getOrCreateSubFolder(subDia, "FIN");
}

function _finFormatBrl_(n) {
  var v = Number(n);
  if (isNaN(v)) v = 0;
  return v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

function _finEsc_(s) {
  return String(s == null ? "" : s)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function _finParseIso_(iso) {
  if (!iso) return null;
  var m = String(iso).trim().match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;
  return new Date(parseInt(m[1], 10), parseInt(m[2], 10) - 1, parseInt(m[3], 10));
}

function _finDateKey_(d) {
  if (!d || isNaN(d.getTime())) return null;
  return d.getFullYear() * 10000 + (d.getMonth() + 1) * 100 + d.getDate();
}

function _finLerLancamentos_() {
  if (typeof getLivroDiarioLancamentos !== "function") return [];
  var r = getLivroDiarioLancamentos({});
  return (r && r.rows) ? r.rows : [];
}

function _finDelta_(row) {
  var v = Number(row["VALOR"]);
  if (isNaN(v)) v = 0;
  return -v;
}

function _finIsReceita_(row) {
  return Number(row["VALOR"]) < 0;
}

function _finIsDespesa_(row) {
  return Number(row["VALOR"]) > 0;
}

function _finIsPago_(row) {
  var st = String(row["STATUS DO PAGAMENTO"] || "").toLowerCase();
  if (st === "pago") return true;
  return !!String(row["DATA PAGAMENTO"] || "").trim();
}

function _finIsDataVencimentoFicticia_(row) {
  var v = String(row["DATA VENCIMENTO"] || "").trim();
  if (!v) return false;
  var alvo = (typeof LIVRO_DIARIO_DATA_VENC_FICTICIA !== "undefined")
    ? String(LIVRO_DIARIO_DATA_VENC_FICTICIA).trim() : "31/12/99";
  if (v === alvo || v === "31/12/1999" || v === "31/12/99") return true;
  var d = _asDateLivro(v);
  if (!d) return false;
  var y = d.getFullYear();
  if (y === 2099) return true;
  if (y === 1999 && d.getMonth() === 11 && d.getDate() === 31) return true;
  return false;
}

function _finIsLinhaIgnoradaRelatorio_(row) {
  return _finIsDataVencimentoFicticia_(row);
}

function _finIsAtrasado_(row) {
  if (_finIsPago_(row)) return false;
  if (_finIsDataVencimentoFicticia_(row)) return false;
  var dv = _asDateLivro(row["DATA VENCIMENTO"]);
  if (!dv) return false;
  var hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  dv.setHours(0, 0, 0, 0);
  return dv.getTime() < hoje.getTime();
}

function _finFiltrarEmAberto_(rows) {
  return (rows || []).filter(function (row) {
    return !_finIsPago_(row) && !_finIsLinhaIgnoradaRelatorio_(row);
  });
}

function _finDataRefLinha_(row, campo) {
  if (campo === "pagamento") return _asDateLivro(row["DATA PAGAMENTO"]);
  if (campo === "competencia") return _asDateLivro(row["DATA COMPETÊNCIA"]);
  return _asDateLivro(row["DATA VENCIMENTO"]);
}

function _finFiltrarPeriodo_(rows, dataInicio, dataFim, campoData) {
  var di = _finParseIso_(dataInicio);
  var df = _finParseIso_(dataFim);
  if (df) df.setHours(23, 59, 59, 999);
  var ki = di ? _finDateKey_(di) : null;
  var kf = df ? _finDateKey_(df) : null;
  campoData = campoData || "vencimento";
  return rows.filter(function (row) {
    var d = _finDataRefLinha_(row, campoData);
    if (!d && (ki || kf)) return false;
    if (!d) return true;
    var k = _finDateKey_(d);
    if (ki != null && k < ki) return false;
    if (kf != null && k > kf) return false;
    return true;
  });
}

function _finFiltrarContas_(rows, contas) {
  if (!contas || !contas.length) return rows;
  var set = {};
  contas.forEach(function (c) {
    var k = String(c || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
    if (k) set[k] = true;
  });
  if (!Object.keys(set).length) return rows;
  return rows.filter(function (row) {
    var cf = String(row["CONTA FINANCEIRA"] || "").toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "").trim();
    return set[cf];
  });
}

function _finNomeArquivo_(tipo, params, sufixo) {
  var p = params || {};
  var di = (p.dataInicio || "").replace(/-/g, "");
  var df = (p.dataFim || "").replace(/-/g, "");
  var partes = ["Relatorio", tipo];
  if (di) partes.push("de" + di);
  if (df) partes.push("ate" + df);
  if (p.ano) partes.push("Ano" + String(p.ano));
  if (p.contaFinanceira) {
    var cf = String(p.contaFinanceira).replace(/[^\w]+/g, "").slice(0, 16);
    if (cf) partes.push(cf);
  }
  if (p.topN) partes.push("Top" + p.topN);
  if (sufixo) partes.push(sufixo);
  var base = partes.join("_").replace(/[^\w\-]+/g, "_").replace(/_+/g, "_");
  if (base.length > 180) base = base.slice(0, 180);
  return base + ".pdf";
}

function _finHtmlPdfWrap_(titulo, subtitulo, bodyHtml) {
  var tz = "America/Sao_Paulo";
  var gerado = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm");
  return "<!DOCTYPE html><html><head><meta charset=\"utf-8\"><style>" +
    "body{font-family:Arial,Helvetica,sans-serif;font-size:11px;color:#0f172a;margin:24px;}" +
    "h1{font-size:18px;margin:0 0 4px;color:#0f766e;}" +
    ".sub{font-size:11px;color:#64748b;margin-bottom:16px;}" +
    ".kpi{display:flex;flex-wrap:wrap;gap:10px;margin:12px 0 18px;}" +
    ".kpi div{background:#f0fdfa;border:1px solid #99f6e4;border-radius:8px;padding:10px 14px;min-width:140px;}" +
    ".kpi .l{font-size:9px;text-transform:uppercase;color:#64748b;font-weight:700;}" +
    ".kpi .v{font-size:14px;font-weight:700;margin-top:4px;}" +
    "table{width:100%;border-collapse:collapse;margin-top:8px;}" +
    "th,td{border:1px solid #e2e8f0;padding:6px 8px;text-align:left;vertical-align:top;}" +
    "th{background:#f1f5f9;font-size:9px;text-transform:uppercase;color:#475569;}" +
    "tr:nth-child(even){background:#fafafa;}" +
    ".num{text-align:right;white-space:nowrap;}" +
    ".foot{margin-top:20px;font-size:9px;color:#94a3b8;}" +
    "</style></head><body>" +
    "<h1>" + _finEsc_(titulo) + "</h1>" +
    "<p class=\"sub\">" + _finEsc_(subtitulo) + " · Gerado em " + gerado + "</p>" +
    bodyHtml +
    "<p class=\"foot\">Sistema TUBA — Livro Diário · Relatório financeiro</p>" +
    "</body></html>";
}

function _finSalvarPdfHtml_(html, nomeArquivo) {
  var pasta = getPastaFinDoDia_(new Date());
  var nome = String(nomeArquivo || "relatorio.pdf").trim();
  if (!/\.pdf$/i.test(nome)) nome += ".pdf";
  var existentes = pasta.getFilesByName(nome);
  while (existentes.hasNext()) {
    try { existentes.next().setTrashed(true); } catch (e) {}
  }
  var blob = HtmlService.createHtmlOutput(html)
    .setWidth(794)
    .setHeight(1123)
    .getAs("application/pdf")
    .setName(nome);
  var file = pasta.createFile(blob);
  var meta = {
    nome: file.getName(),
    url: file.getUrl(),
    dataGeracao: new Date().toISOString(),
    pastaUrl: pasta.getUrl()
  };
  _finRegistrarRelatorioRecente_(meta);
  return meta;
}

function _finRegistrarRelatorioRecente_(meta) {
  try {
    var raw = PropertiesService.getScriptProperties().getProperty(FIN_RELATORIOS_CACHE_LISTA_KEY);
    var list = raw ? JSON.parse(raw) : [];
    if (!Array.isArray(list)) list = [];
    list.unshift(meta);
    list = list.slice(0, 80);
    PropertiesService.getScriptProperties().setProperty(FIN_RELATORIOS_CACHE_LISTA_KEY, JSON.stringify(list));
  } catch (e) { /* noop */ }
}

function listarRelatoriosFinanceirosRecentes(limite) {
  var max = parseInt(limite, 10);
  if (isNaN(max) || max <= 0) max = 30;
  var out = [];
  var seen = {};

  try {
    var raw = PropertiesService.getScriptProperties().getProperty(FIN_RELATORIOS_CACHE_LISTA_KEY);
    if (raw) {
      var cached = JSON.parse(raw);
      if (Array.isArray(cached)) {
        cached.forEach(function (m) {
          if (!m || !m.url || seen[m.url]) return;
          seen[m.url] = true;
          out.push(m);
        });
      }
    }
  } catch (e) { /* noop */ }

  try {
    var raiz = _getRootFolder_();
    var hoje = new Date();
    for (var d = 0; d < 60 && out.length < max; d++) {
      var dt = new Date(hoje.getTime());
      dt.setDate(dt.getDate() - d);
      var pastaFin = null;
      try { pastaFin = getPastaFinDoDia_(dt); } catch (eP) { continue; }
      if (!pastaFin) continue;
      var files = pastaFin.getFiles();
      while (files.hasNext() && out.length < max) {
        var f = files.next();
        var url = f.getUrl();
        if (seen[url]) continue;
        seen[url] = true;
        out.push({
          nome: f.getName(),
          url: url,
          dataGeracao: f.getLastUpdated().toISOString(),
          pastaUrl: pastaFin.getUrl()
        });
      }
    }
  } catch (eDrive) {
    Logger.log("listarRelatoriosFinanceirosRecentes: " + (eDrive && eDrive.message));
  }

  out.sort(function (a, b) {
    return new Date(b.dataGeracao || 0).getTime() - new Date(a.dataGeracao || 0).getTime();
  });
  return out.slice(0, max);
}

function getOpcoesRelatorioFinanceiro() {
  var rows = _finLerLancamentos_();
  var contasFin = {};
  rows.forEach(function (r) {
    var c = String(r["CONTA FINANCEIRA"] || "").trim();
    if (c) contasFin[c] = true;
  });
  var cadContas = [];
  try {
    if (typeof getLivroDiarioCadastro === "function") {
      cadContas = (getLivroDiarioCadastro().contasFinanceiras || []);
    }
  } catch (e) { /* noop */ }
  cadContas.forEach(function (c) { if (c) contasFin[c] = true; });
  var resumo = { saldoGeralHoje: 0, saldoFuturoGeral: 0 };
  try {
    if (typeof getLivroDiarioResumo === "function") {
      var rs = getLivroDiarioResumo("");
      resumo.saldoGeralHoje = rs.saldoGeralHoje || 0;
      resumo.saldoFuturoGeral = rs.saldoFuturoGeral || 0;
    }
  } catch (e2) { /* noop */ }
  return {
    tipos: getTiposRelatorioFinanceiro(),
    contasFinanceiras: Object.keys(contasFin).sort(),
    resumo: resumo,
    totalLancamentos: rows.length
  };
}

function _finTabelaSimples_(headers, linhas) {
  var h = "<table><thead><tr>";
  headers.forEach(function (x) { h += "<th>" + _finEsc_(x) + "</th>"; });
  h += "</tr></thead><tbody>";
  if (!linhas.length) {
    h += "<tr><td colspan=\"" + headers.length + "\">Sem dados para os filtros informados.</td></tr>";
  } else {
    linhas.forEach(function (ln) {
      h += "<tr>";
      ln.forEach(function (cell, i) {
        var cls = (i > 0 && typeof cell === "string" && cell.indexOf("R$") === 0) ? " class=\"num\"" : (typeof cell === "number" ? " class=\"num\"" : "");
        var val = typeof cell === "number" ? _finFormatBrl_(cell) : cell;
        h += "<td" + cls + ">" + _finEsc_(val) + "</td>";
      });
      h += "</tr>";
    });
  }
  h += "</tbody></table>";
  return h;
}

function _finKpisHtml_(items) {
  var h = "<div class=\"kpi\">";
  items.forEach(function (it) {
    h += "<div><div class=\"l\">" + _finEsc_(it.label) + "</div><div class=\"v\">" + _finEsc_(it.value) + "</div></div>";
  });
  h += "</div>";
  return h;
}

function _finSubtituloPeriodo_(params, def) {
  var p = params || {};
  var parts = [];
  if (def && def.precisaAno && p.ano) parts.push("Ano " + p.ano);
  if (def && def.precisaPeriodo) {
    if (p.dataInicio) parts.push("De " + p.dataInicio);
    if (p.dataFim) parts.push("até " + p.dataFim);
  }
  if (def && def.precisaContaUnica && p.contaFinanceira) {
    parts.push("Conta: " + p.contaFinanceira);
    if (def.id === "saldo_conta_mensal") {
      parts.push("Realizado até última data de pagamento");
    }
  }
  if (def && def.somenteReceitaIndustria) {
    parts.push("Conta contábil: Receita Indústria (pedidos)");
  }
  if (def && def.campoDataFixo === "pagamento") {
    parts.push("Referência: data de pagamento");
  } else if (def && def.precisaReferenciaData && p.campoData) {
    var labels = { vencimento: "vencimento", pagamento: "pagamento", competencia: "competência" };
    parts.push("Referência: data de " + (labels[p.campoData] || p.campoData));
  }
  if (def && def.envolveSaldo && !def.saldoFuturo) {
    parts.push("Saldos: data de pagamento (realizado)");
  } else   if (def && def.envolveSaldo && def.saldoFuturo) {
    parts.push("Saldo inicial realizado · projeção em aberto");
  }
  if (def && def.somenteEmAberto) {
    parts.push("Posição em aberto (vencimento futuro)");
  }
  if (def && def.campoDataFixo === "vencimento" && def.precisaPeriodo) {
    parts.push("Período por data de vencimento");
  }
  return parts.join(" · ") || "Todos os lançamentos";
}

function _finGerarCorpoRelatorio_(tipo, params, rows, ctx) {
  var p = params || {};
  ctx = ctx || {};
  var todas = ctx.todas || rows;
  var saldos = ctx.saldos || _finSaldosRealizadosLivro_(p.contaFinanceira || "");
  var topN = parseInt(p.topN, 10);
  if (isNaN(topN) || topN <= 0) topN = 10;

  if (tipo === "resumo_executivo") {
    var receita = 0, despesa = 0, aReceber = 0, aPagar = 0, qPago = 0, qAberto = 0;
    rows.forEach(function (r) {
      if (_finIsReceita_(r)) receita += Math.abs(Number(r["VALOR"]) || 0);
      if (_finIsDespesa_(r)) despesa += Math.abs(Number(r["VALOR"]) || 0);
      if (_finIsPago_(r)) qPago++;
      else {
        qAberto++;
        if (_finIsReceita_(r)) aReceber += Math.abs(Number(r["VALOR"]) || 0);
        if (_finIsDespesa_(r)) aPagar += Math.abs(Number(r["VALOR"]) || 0);
      }
    });
    return _finKpisHtml_([
      { label: "Saldo geral realizado", value: _finFormatBrl_(saldos.saldoGeralHoje) },
      { label: "Saldo futuro (livro)", value: _finFormatBrl_(saldos.saldoFuturoGeral) },
      { label: "Receitas (abs.)", value: _finFormatBrl_(receita) },
      { label: "Despesas", value: _finFormatBrl_(despesa) },
      { label: "A receber (aberto)", value: _finFormatBrl_(aReceber) },
      { label: "A pagar (aberto)", value: _finFormatBrl_(aPagar) },
      { label: "Lançamentos pagos", value: String(qPago) },
      { label: "Lançamentos em aberto", value: String(qAberto) }
    ]) + _finNotaSaldoRealizadoHtml_();
  }

  if (tipo === "projetos_carro_chefe") {
    rows = _finFiltrarReceitaIndustria_(rows);
    var map = {};
    rows.forEach(function (r) {
      if (!_finIsReceita_(r)) return;
      var cod = String(r["CÓDIGO DO PROJETO"] || "").trim() || "(sem código)";
      if (!map[cod]) map[cod] = { cod: cod, desc: String(r["DESCRIÇÃO DO PROJETO"] || ""), total: 0 };
      map[cod].total += Math.abs(Number(r["VALOR"]) || 0);
    });
    var arr = Object.keys(map).map(function (k) { return map[k]; });
    arr.sort(function (a, b) { return b.total - a.total; });
    arr = arr.slice(0, topN);
    return _finTabelaSimples_(["Projeto", "Descrição", "Receita"],
      arr.map(function (x) { return [x.cod, x.desc, _finFormatBrl_(x.total)]; }));
  }

  if (tipo === "clientes_carro_chefe") {
    rows = _finFiltrarReceitaIndustria_(rows);
    var mapC = {};
    rows.forEach(function (r) {
      if (!_finIsReceita_(r)) return;
      var cli = String(r["CLIENTE"] || "").trim() || "(sem cliente)";
      if (!mapC[cli]) mapC[cli] = 0;
      mapC[cli] += Math.abs(Number(r["VALOR"]) || 0);
    });
    var arrC = Object.keys(mapC).map(function (k) { return { nome: k, total: mapC[k] }; });
    arrC.sort(function (a, b) { return b.total - a.total; });
    arrC = arrC.slice(0, topN);
    return _finTabelaSimples_(["Cliente", "Receita"],
      arrC.map(function (x) { return [x.nome, _finFormatBrl_(x.total)]; }));
  }

  if (tipo === "saldo_futuro_periodo") {
    var porDia = {};
    var contaFluxo = String(p.contaFinanceira || "").trim();
    var saldoInicialRealizado = _finSaldoInicialRealizado_(contaFluxo);
    var saldoAcum = saldoInicialRealizado;
    rows.forEach(function (r) {
      if (_finIsPago_(r)) return;
      if (_finIsLinhaIgnoradaRelatorio_(r)) return;
      var dv = _asDateLivro(r["DATA VENCIMENTO"]);
      if (!dv) return;
      var k = Utilities.formatDate(dv, "America/Sao_Paulo", "dd/MM/yyyy");
      if (!porDia[k]) porDia[k] = { ent: 0, sai: 0 };
      var abs = Math.abs(Number(r["VALOR"]) || 0);
      if (_finIsReceita_(r)) porDia[k].ent += abs;
      else porDia[k].sai += abs;
    });
    var keys = Object.keys(porDia).sort(function (a, b) {
      var pa = a.split("/"); var pb = b.split("/");
      return new Date(pa[2], pa[1] - 1, pa[0]) - new Date(pb[2], pb[1] - 1, pb[0]);
    });
    var linhasFluxo = [];
    keys.forEach(function (k) {
      var item = porDia[k];
      saldoAcum += item.ent - item.sai;
      linhasFluxo.push([k, _finFormatBrl_(item.ent), _finFormatBrl_(item.sai), _finFormatBrl_(saldoAcum)]);
    });
    var intro = _finKpisHtml_([
      { label: "Saldo inicial realizado", value: _finFormatBrl_(saldoInicialRealizado) },
      { label: "Saldo projetado (final)", value: _finFormatBrl_(saldoAcum) }
    ]);
    return intro + _finNotaSaldoFuturoHtml_() +
      _finTabelaSimples_(["Vencimento", "A receber", "A pagar", "Saldo projetado"], linhasFluxo);
  }

  if (tipo === "saldo_conta_mensal") {
    var conta = String(p.contaFinanceira || "").trim();
    var ano = parseInt(p.ano, 10);
    if (!conta) return "<p>Selecione a conta financeira.</p>";
    if (isNaN(ano) || ano < 2000 || ano > 2100) return "<p>Selecione o ano.</p>";
    var extrato = null;
    try {
      if (typeof calcularExtratoContaFinanceira === "function") {
        extrato = calcularExtratoContaFinanceira(conta, ano);
      }
    } catch (eExt) {
      return "<p>Erro ao calcular extrato: " + _finEsc_(eExt.message || eExt) + "</p>";
    }
    if (!extrato || !extrato.ok) return "<p>Não foi possível calcular o extrato da conta.</p>";
    var linhasM = (extrato.meses || []).map(function (mx) {
      return [
        mx.label,
        String(mx.qtd),
        _finFormatBrl_(mx.entradas),
        _finFormatBrl_(mx.saidas),
        _finFormatBrl_(mx.liquido),
        _finFormatBrl_(mx.saldoAcumulado)
      ];
    });
    return _finKpisHtml_([
      { label: "Conta", value: conta },
      { label: "Ano", value: String(ano) },
      { label: "Última data pagamento", value: extrato.ultimaDataPagamento || "—" },
      { label: "Saldo antes de " + ano, value: _finFormatBrl_(extrato.saldoAntesAno) },
      { label: "Saldo realizado (extrato)", value: _finFormatBrl_(extrato.saldoRealizado) },
      { label: "Saldo ao fim do ano (na grade)", value: _finFormatBrl_(extrato.saldoFimAno) }
    ]) + _finNotaSaldoRealizadoHtml_() +
      _finTabelaSimples_(["Mês", "Lanç.", "Entradas", "Saídas", "Líquido mês", "Saldo acumulado"], linhasM);
  }

  if (tipo === "contas_pagar_receber") {
    var abertos = _finFiltrarEmAberto_(rows);
    var rec = 0, pag = 0;
    abertos.forEach(function (r) {
      var abs = Math.abs(Number(r["VALOR"]) || 0);
      if (_finIsReceita_(r)) rec += abs;
      if (_finIsDespesa_(r)) pag += abs;
    });
    return _finKpisHtml_([
      { label: "Total a receber", value: _finFormatBrl_(rec) },
      { label: "Total a pagar", value: _finFormatBrl_(pag) },
      { label: "Posição líquida", value: _finFormatBrl_(rec - pag) },
      { label: "Lançamentos em aberto", value: String(abertos.length) }
    ]) + "<p class=\"sub\" style=\"margin:8px 0 12px;font-size:11px;color:#64748b;\">" +
      "Todos os lançamentos não quitados (sem filtro de período ou data de pagamento). Exclui vencimento fictício 31/12/99.</p>" +
      _finTabelaSimples_(["Natureza", "Valor"], [
        ["A receber", _finFormatBrl_(rec)],
        ["A pagar", _finFormatBrl_(pag)]
      ]);
  }

  if (tipo === "por_status_pagamento") {
    var stMap = {};
    rows.forEach(function (r) {
      var st = String(r["STATUS DO PAGAMENTO"] || "—").trim() || "—";
      if (!stMap[st]) stMap[st] = { rec: 0, desp: 0, qtd: 0 };
      stMap[st].qtd++;
      var abs = Math.abs(Number(r["VALOR"]) || 0);
      if (_finIsReceita_(r)) stMap[st].rec += abs;
      else stMap[st].desp += abs;
    });
    var lnSt = Object.keys(stMap).sort().map(function (st) {
      var x = stMap[st];
      return [st, String(x.qtd), _finFormatBrl_(x.rec), _finFormatBrl_(x.desp)];
    });
    return _finTabelaSimples_(["Status", "Qtd", "Receitas", "Despesas"], lnSt);
  }

  if (tipo === "por_conta_contabil") {
    var ccMap = {};
    rows.forEach(function (r) {
      var cc = String(r["CONTA CONTÁBIL"] || "—").trim() || "—";
      if (!ccMap[cc]) ccMap[cc] = { rec: 0, desp: 0 };
      var abs = Math.abs(Number(r["VALOR"]) || 0);
      if (_finIsReceita_(r)) ccMap[cc].rec += abs;
      else ccMap[cc].desp += abs;
    });
    var lnCc = Object.keys(ccMap).sort().map(function (cc) {
      var x = ccMap[cc];
      return [cc, _finFormatBrl_(x.rec), _finFormatBrl_(x.desp), _finFormatBrl_(x.rec - x.desp)];
    });
    return _finTabelaSimples_(["Conta contábil", "Receitas", "Despesas", "Líquido"], lnCc);
  }

  if (tipo === "por_conta_financeira") {
    var cfMap = {};
    rows.forEach(function (r) {
      var cf = String(r["CONTA FINANCEIRA"] || "—").trim() || "—";
      if (!cfMap[cf]) cfMap[cf] = { rec: 0, desp: 0, qtd: 0 };
      cfMap[cf].qtd++;
      var abs = Math.abs(Number(r["VALOR"]) || 0);
      if (_finIsReceita_(r)) cfMap[cf].rec += abs;
      else cfMap[cf].desp += abs;
    });
    var lnCf = Object.keys(cfMap).sort().map(function (cf) {
      var x = cfMap[cf];
      return [cf, String(x.qtd), _finFormatBrl_(x.rec), _finFormatBrl_(x.desp)];
    });
    return _finTabelaSimples_(["Conta financeira", "Qtd", "Receitas", "Despesas"], lnCf);
  }

  if (tipo === "inadimplencia") {
    var atras = rows.filter(_finIsAtrasado_);
    var total = 0;
    atras.forEach(function (r) {
      if (_finIsReceita_(r)) total += Math.abs(Number(r["VALOR"]) || 0);
    });
    var lnA = atras.map(function (r) {
      return [
        String(r["CLIENTE"] || ""),
        String(r["CÓDIGO DO PROJETO"] || ""),
        String(r["DATA VENCIMENTO"] || ""),
        String(r["STATUS DO PAGAMENTO"] || ""),
        _finFormatBrl_(Math.abs(Number(r["VALOR"]) || 0))
      ];
    });
    return _finKpisHtml_([{ label: "Total em atraso (receitas)", value: _finFormatBrl_(total) }, { label: "Linhas", value: String(atras.length) }]) +
      _finTabelaSimples_(["Cliente", "Projeto", "Vencimento", "Status", "Valor"], lnA);
  }

  if (tipo === "despesas_por_projeto") {
    var dMap = {};
    rows.forEach(function (r) {
      if (!_finIsDespesa_(r)) return;
      var cod = String(r["CÓDIGO DO PROJETO"] || "").trim() || "(sem código)";
      if (!dMap[cod]) dMap[cod] = { total: 0, desc: String(r["DESCRIÇÃO DO PROJETO"] || "").trim() };
      dMap[cod].total += Math.abs(Number(r["VALOR"]) || 0);
      if (!dMap[cod].desc && r["DESCRIÇÃO DO PROJETO"]) dMap[cod].desc = String(r["DESCRIÇÃO DO PROJETO"]).trim();
    });
    var lnD = Object.keys(dMap).sort(function (a, b) { return dMap[b].total - dMap[a].total; }).map(function (cod) {
      return [cod, dMap[cod].desc, _finFormatBrl_(dMap[cod].total)];
    });
    return _finTabelaSimples_(["Projeto", "Descrição", "Despesas"], lnD);
  }

  if (tipo === "receitas_periodo") {
    rows = _finFiltrarReceitaIndustria_(rows);
    var recRows = rows.filter(_finIsReceita_);
    var lnR = recRows.map(function (r) {
      return [
        String(r["DATA COMPETÊNCIA"] || ""),
        String(r["CLIENTE"] || ""),
        String(r["CÓDIGO DO PROJETO"] || ""),
        String(r["DATA VENCIMENTO"] || ""),
        String(r["STATUS DO PAGAMENTO"] || ""),
        _finFormatBrl_(Math.abs(Number(r["VALOR"]) || 0))
      ];
    });
    var totR = recRows.reduce(function (s, r) { return s + Math.abs(Number(r["VALOR"]) || 0); }, 0);
    return _finKpisHtml_([{ label: "Total receitas", value: _finFormatBrl_(totR) }]) +
      _finTabelaSimples_(["Competência", "Cliente", "Projeto", "Vencimento", "Status", "Valor"], lnR);
  }

  if (tipo === "extrato_lancamentos") {
    var lnE = rows.slice(0, 500).map(function (r) {
      return [
        String(r["DATA VENCIMENTO"] || ""),
        String(r["CLIENTE"] || ""),
        String(r["CÓDIGO DO PROJETO"] || ""),
        String(r["CONTA FINANCEIRA"] || ""),
        String(r["STATUS DO PAGAMENTO"] || ""),
        _finFormatBrl_(Number(r["VALOR"]) || 0)
      ];
    });
    var nota = rows.length > 500 ? "<p><em>Exibindo 500 de " + rows.length + " lançamentos.</em></p>" : "";
    return nota + _finTabelaSimples_(["Venc.", "Cliente", "Projeto", "Conta fin.", "Status", "Valor"], lnE);
  }

  return "<p>Tipo de relatório não implementado.</p>";
}

function gerarRelatorioFinanceiroLivro(tipo, params) {
  tipo = String(tipo || "").trim();
  if (!tipo) throw new Error("Informe o tipo de relatório.");
  var tipos = getTiposRelatorioFinanceiro();
  var def = tipos.filter(function (t) { return t.id === tipo; })[0];
  if (!def) throw new Error("Tipo de relatório inválido: " + tipo);

  var p = params || {};
  var todas = _finLerLancamentos_();
  var rows = todas;
  var campoData = def.campoDataFixo || p.campoData || "vencimento";

  if (def.precisaAno) {
    var anoNum = parseInt(p.ano, 10);
    if (isNaN(anoNum) || anoNum < 2000 || anoNum > 2100) throw new Error("Selecione o ano.");
    p.ano = anoNum;
  }

  if (def.precisaContaUnica) {
    var contaU = String(p.contaFinanceira || "").trim();
    if (!contaU) throw new Error("Selecione a conta financeira para este relatório.");
    rows = _finFiltrarContas_(rows, [contaU]);
  }

  if (def.somenteEmAberto) {
    rows = _finFiltrarEmAberto_(todas);
  }

  if (def.precisaPeriodo) {
    if (p.dataInicio || p.dataFim) {
      rows = _finFiltrarPeriodo_(rows, p.dataInicio, p.dataFim, campoData);
    } else {
      var hoje = new Date();
      var ini = new Date(hoje.getFullYear(), hoje.getMonth(), 1);
      var fim = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0);
      rows = _finFiltrarPeriodo_(rows,
        Utilities.formatDate(ini, "America/Sao_Paulo", "yyyy-MM-dd"),
        Utilities.formatDate(fim, "America/Sao_Paulo", "yyyy-MM-dd"),
        campoData);
      p.dataInicio = Utilities.formatDate(ini, "America/Sao_Paulo", "yyyy-MM-dd");
      p.dataFim = Utilities.formatDate(fim, "America/Sao_Paulo", "yyyy-MM-dd");
    }
  }

  if (tipo === "inadimplencia") {
    rows = todas.filter(_finIsAtrasado_);
  }

  var titulo = def.nome;
  var subtitulo = _finSubtituloPeriodo_(p, def);
  var saldos = _finSaldosRealizadosLivro_(p.contaFinanceira || "");
  var corpo = _finGerarCorpoRelatorio_(tipo, p, rows, { def: def, todas: todas, saldos: saldos });
  var html = _finHtmlPdfWrap_(titulo, subtitulo, corpo);
  var nomeArquivo = _finNomeArquivo_(tipo, p);
  var arquivo = _finSalvarPdfHtml_(html, nomeArquivo);

  return {
    ok: true,
    tipo: tipo,
    nome: arquivo.nome,
    url: arquivo.url,
    pastaUrl: arquivo.pastaUrl,
    linhasAnalisadas: rows.length,
    dataInicio: p.dataInicio || "",
    dataFim: p.dataFim || ""
  };
}
