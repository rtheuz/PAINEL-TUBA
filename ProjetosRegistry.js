function registrarOrcamento(cliente, codigoProjeto, valorTotal, dataOrcamento, urlPdf, urlMemoria, chapas, observacoes, produtosCadastrados, dadosFormularioCompleto, isPedido) {
  const todosProcessos = [];
  const ordemProcessos = ["MP", "CL", "D", "S", "Pin", "CAD", "ACB"];
  (produtosCadastrados || []).forEach(function (prod) {
    var procs = prod.processos;
    if (procs && Array.isArray(procs)) {
      procs.forEach(function (sigla) {
        if (todosProcessos.indexOf(sigla) < 0) todosProcessos.push(sigla);
      });
    }
  });
  todosProcessos.sort(function (a, b) {
    var ia = ordemProcessos.indexOf(a);
    var ib = ordemProcessos.indexOf(b);
    if (ia === -1) ia = ordemProcessos.length;
    if (ib === -1) ib = ordemProcessos.length;
    if (ia !== ib) return ia - ib;
    return a < b ? -1 : (a > b ? 1 : 0);
  });
  const processosManual = (observacoes && observacoes.processos) ? String(observacoes.processos).trim() : "";
  const processosStr = processosManual || todosProcessos.join(", ");

  const descricao = (observacoes && observacoes.descricao) || "";
  const prazoOriginal = (observacoes && observacoes.prazo) || "";

  let prazoParaPlanilha = prazoOriginal;
  try {
    if (isPedido && prazoOriginal) {
      const s = String(prazoOriginal).trim();
      const mDias = s.match(/(\d+)\s*dias?/i);
      if (mDias && !/\/\d{2}\/\d{4}/.test(s)) {
        const qtdDias = parseInt(mDias[1], 10);
        if (!isNaN(qtdDias)) {
          const tz = ss.getSpreadsheetTimeZone ? ss.getSpreadsheetTimeZone() : Session.getScriptTimeZone();
          const base = new Date();
          base.setDate(base.getDate() + qtdDias);
          prazoParaPlanilha = Utilities.formatDate(base, tz || "America/Sao_Paulo", "yyyy-MM-dd");
        }
      }
    }
  } catch (ePrazo) {
    Logger.log("Aviso registrarOrcamento: falha ao calcular prazo em dias: " + (ePrazo && ePrazo.message ? ePrazo.message : ePrazo));
  }

  chapas = chapas || [];

  const listaProds = (produtosCadastrados || []).map(function (prod) {
    if (!prod || typeof prod !== "object") return {};
    return JSON.parse(JSON.stringify(prod));
  });
  // Mantém PRD apenas quando o item bate com catálogo;
  // caso contrário, gera novo no fluxo de gravação do PDF.
  atribuirPRDsUnicos(listaProds);
  produtosCadastrados = listaProds;
  if (dadosFormularioCompleto && Array.isArray(dadosFormularioCompleto.produtosCadastrados)) {
    dadosFormularioCompleto.produtosCadastrados = JSON.parse(JSON.stringify(listaProds));
  }

  function _sincronizarProdutosNaRelacao(produtos, codigoProj, nomeCliente) {
    try {
      if (!produtos || !Array.isArray(produtos) || produtos.length === 0) return;
      if (!SHEET_PRODUTOS) return;

      const existentes = _coletarCodigosPRDDoCatalogo();
      produtos.forEach(function (prod) {
        const codigo = _normalizarCodigoPRD(prod && prod.codigo);
        if (!_ehCodigoPRDValido(codigo)) return;
        if (existentes.has(codigo)) return;

        const produtoRelacao = {
          codigo: codigo,
          descricao: prod.descricao || "",
          ncm: prod.ncm || "",
          preco: Number(prod.precoUnitario) || 0,
          unidade: prod.unidade || "UN",
          caracteristicas: (prod.descricoesProcessos && typeof prod.descricoesProcessos === "object") ? JSON.stringify(prod.descricoesProcessos) : "",
          projeto: codigoProj || "",
          cliente: nomeCliente || "",
          processos: prod.processos && Array.isArray(prod.processos) ? prod.processos : [],
          pecasPorChapa: prod.pecasPorChapa != null ? prod.pecasPorChapa : "",
          precosProcessos: prod.precosProcessos && typeof prod.precosProcessos === "object" ? prod.precosProcessos : {}
        };
        inserirProdutoNaRelacao(produtoRelacao);
        existentes.add(codigo);
      });
    } catch (eSyncRel) {
      Logger.log("Aviso _sincronizarProdutosNaRelacao: " + (eSyncRel && eSyncRel.message ? eSyncRel.message : eSyncRel));
    }
  }

  try {
    if (codigoProjeto && /_v\d+$/.test(codigoProjeto)) {
      const sheetProj = SHEET_PROJ;
      if (sheetProj) {
        const codigoBase = codigoProjeto.replace(/_v\d+$/, "");
        let linhaBase = 0;
        const linhaProjetoForm = (dadosFormularioCompleto && dadosFormularioCompleto.linhaProjeto != null) ? parseInt(dadosFormularioCompleto.linhaProjeto, 10) : NaN;
        if (!isNaN(linhaProjetoForm) && linhaProjetoForm >= 2 && linhaProjetoForm <= sheetProj.getLastRow()) {
          const hdrs = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
          const idxP = _findHeaderIndexProjetos(hdrs, "PROJETO");
          if (idxP >= 0 && String(sheetProj.getRange(linhaProjetoForm, idxP + 1).getValue() || "").trim() === codigoBase) {
            linhaBase = linhaProjetoForm;
          }
        }
        if (!linhaBase) linhaBase = findRowByColumnValue(sheetProj, "PROJETO", codigoBase) || 0;
        if (linhaBase) {
          const hdrs = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
          const idxJson = _findHeaderIndex(hdrs, "JSON_DADOS");
          const idxLink = _findHeaderIndex(hdrs, "LINK DO PDF");
          if (idxJson >= 0) {
            let jsonData = {};
            try { jsonData = JSON.parse(String(sheetProj.getRange(linhaBase, idxJson + 1).getValue() || "{}")); } catch (e) { jsonData = {}; }
            if (!jsonData.urlPdf && idxLink >= 0) {
              jsonData.urlPdf = String(sheetProj.getRange(linhaBase, idxLink + 1).getValue() || "");
            }
            if (!Array.isArray(jsonData.versoes)) jsonData.versoes = [];

            const nomeVersaoNovo = (dadosFormularioCompleto && dadosFormularioCompleto.nomeVersao != null)
              ? String(dadosFormularioCompleto.nomeVersao).trim()
              : "";
            let tipoPdfNovo = isPedido ? "Pedido" : "Proposta";
            try {
              const idxStatusOrcBase = _findHeaderIndexProjetos(hdrs, "STATUS_ORCAMENTO");
              const statusOrcBase = idxStatusOrcBase >= 0 ? String(sheetProj.getRange(linhaBase, idxStatusOrcBase + 1).getValue() || "").trim() : "";
              if (String(statusOrcBase).toLowerCase() === "convertido em pedido") tipoPdfNovo = "Pedido";
            } catch (eTipoV) { }

            if (nomeVersaoNovo) {
              const nomeBaseExistente = (jsonData && jsonData.dados && jsonData.dados.nomeVersao != null)
                ? String(jsonData.dados.nomeVersao).trim()
                : "";

              if (nomeBaseExistente && nomeBaseExistente === nomeVersaoNovo) {
                throw new Error("Já existe uma versão com o mesmo Nome da Versão neste projeto.");
              }

              let existIdxTmp = jsonData.versoes.findIndex(function (v) { return v && v.codigo === codigoProjeto; });
              for (let i = 0; i < jsonData.versoes.length; i++) {
                const v = jsonData.versoes[i];
                if (!v) continue;
                const nomeExist = (v.nomeVersao != null) ? String(v.nomeVersao).trim() : "";
                if (!nomeExist) continue;
                if (nomeExist === nomeVersaoNovo && i !== existIdxTmp) {
                  throw new Error("Já existe uma versão com o mesmo Nome da Versão neste projeto.");
                }
              }
            }

            const numSeqV = (dadosFormularioCompleto && dadosFormularioCompleto.numeroSequencial != null) ? dadosFormularioCompleto.numeroSequencial : null;
            const dadosV = dadosFormularioCompleto ? JSON.parse(JSON.stringify(dadosFormularioCompleto)) : {};
            if (dadosV.observacoes) dadosV.observacoes.projeto = codigoProjeto;
            try {
              const idxStV = _findHeaderIndexProjetos(hdrs, "STATUS_ORCAMENTO");
              const stLinhaV = idxStV >= 0 ? String(sheetProj.getRange(linhaBase, idxStV + 1).getValue() || "").trim() : "";
              if (isPedido && stLinhaV === "Convertido em Pedido") {
                var infoBaseV = (jsonData.dados && jsonData.dados.infoPedido) ? jsonData.dados.infoPedido : null;
                if (infoBaseV && typeof infoBaseV === "object") {
                  if (!dadosV.infoPedido) dadosV.infoPedido = {};
                  var chavesV = ["dataVirouPedido", "dataEntrega", "notaFiscal", "valorPedido", "comprovanteEntregaUrl", "temComprovanteEntrega", "dataFimProducao"];
                  for (var vi = 0; vi < chavesV.length; vi++) {
                    var cv = chavesV[vi];
                    if (infoBaseV[cv] != null && infoBaseV[cv] !== "") dadosV.infoPedido[cv] = infoBaseV[cv];
                  }
                  if (infoBaseV.statusDates && typeof infoBaseV.statusDates === "object") {
                    if (!dadosV.infoPedido.statusDates) dadosV.infoPedido.statusDates = {};
                    Object.keys(infoBaseV.statusDates).forEach(function (sk) {
                      if (infoBaseV.statusDates[sk] != null && infoBaseV.statusDates[sk] !== "") dadosV.infoPedido.statusDates[sk] = infoBaseV.statusDates[sk];
                    });
                  }
                }
              }
            } catch (eInfoV) {
              Logger.log("Aviso registrarOrcamento (versão): " + (eInfoV && eInfoV.message ? eInfoV.message : eInfoV));
            }
            const novaEntrada = {
              codigo: codigoProjeto,
              dataSalvo: new Date().toISOString(),
              numeroSequencial: numSeqV,
              urlPdf: urlPdf || "",
              nomeVersao: nomeVersaoNovo,
              tipoPdf: tipoPdfNovo,
              dados: dadosV
            };
            const existIdx = jsonData.versoes.findIndex(function (v) { return v && v.codigo === codigoProjeto; });
            if (existIdx >= 0) { jsonData.versoes[existIdx] = novaEntrada; } else { jsonData.versoes.push(novaEntrada); }
            jsonData.nome = codigoBase;
            sheetProj.getRange(linhaBase, idxJson + 1).setValue(JSON.stringify(jsonData));
            if (idxLink >= 0 && urlPdf) sheetProj.getRange(linhaBase, idxLink + 1).setValue(urlPdf);
            Logger.log("✅ Versão " + codigoProjeto + " armazenada em JSON_DADOS linha " + linhaBase + " (sem nova linha).");
            return;
          }
        }
        throw new Error("Projeto base '" + codigoBase + "' não encontrado. Nova versão não pode criar uma nova linha; recarregue o projeto base e tente novamente.");
      }
    }

    const numeroSequencial = (dadosFormularioCompleto && dadosFormularioCompleto.numeroSequencial) || null;

    const dadosParaJson = dadosFormularioCompleto ? { ...dadosFormularioCompleto } : {};
    dadosParaJson.chapas = chapas || [];
    dadosParaJson.produtosCadastrados = produtosCadastrados || [];
    if (numeroSequencial != null) dadosParaJson.numeroSequencial = numeroSequencial;
    if (!dadosParaJson.observacoes) dadosParaJson.observacoes = {};
    dadosParaJson.observacoes.projeto = codigoProjeto;

    const agora = new Date();

    try {
      const agoraIso = agora.toISOString();
      dadosParaJson.observacoes.dataUltimoOrcamento = agoraIso;
      if (!dadosParaJson.infoPedido) dadosParaJson.infoPedido = {};
      if (!dadosParaJson.infoPedido.dataUltimoOrcamento) {
        dadosParaJson.infoPedido.dataUltimoOrcamento = agoraIso;
      }
    } catch (eDataUlt) {
      Logger.log("Aviso registrarOrcamento: falha ao registrar dataUltimoOrcamento: " + (eDataUlt && eDataUlt.message ? eDataUlt.message : eDataUlt));
    }

    try {
      let descProc = "";
      if (dadosParaJson.processosPedido && Array.isArray(dadosParaJson.processosPedido)) {
        descProc = dadosParaJson.processosPedido
          .map(function (p) { return (p && p.descricao) ? String(p.descricao).trim() : ""; })
          .filter(function (s) { return s && s.length > 0; })
          .join(" / ");
      }
      if (!dadosParaJson.infoPedido) dadosParaJson.infoPedido = dadosParaJson.infoPedido || {};
      if (descProc && !dadosParaJson.infoPedido.descricoesProcessos) {
        dadosParaJson.infoPedido.descricoesProcessos = descProc;
      }
    } catch (eDescProc) {
      Logger.log("Aviso registrarOrcamento: falha ao montar descricoesProcessos: " + (eDescProc && eDescProc.message ? eDescProc.message : eDescProc));
    }

    const sheetProj = SHEET_PROJ;
    const targetSheet = sheetProj;

    var linhaExistente = 0;
    const linhaProjetoForm = (dadosFormularioCompleto && dadosFormularioCompleto.linhaProjeto != null) ? parseInt(dadosFormularioCompleto.linhaProjeto, 10) : NaN;
    if (sheetProj && !isNaN(linhaProjetoForm) && linhaProjetoForm >= 2 && linhaProjetoForm <= sheetProj.getLastRow()) {
      linhaExistente = linhaProjetoForm;
    }
    if (!linhaExistente && sheetProj) {
      linhaExistente = findRowByColumnValue(sheetProj, "PROJETO", codigoProjeto) || 0;
    }

    try {
      if (linhaExistente && sheetProj && isPedido) {
        const hdrsMerge = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
        const idxOrcMerge = _findHeaderIndexProjetos(hdrsMerge, "STATUS_ORCAMENTO");
        const rowMerge = sheetProj.getRange(linhaExistente, 1, linhaExistente, sheetProj.getLastColumn()).getValues()[0];
        const stMerge = (idxOrcMerge >= 0 && rowMerge[idxOrcMerge] != null) ? String(rowMerge[idxOrcMerge]).trim() : "";
        if (stMerge === "Convertido em Pedido") {
          const idxJsonMerge = _findHeaderIndex(hdrsMerge, "JSON_DADOS");
          if (idxJsonMerge >= 0) {
            const cellMerge = sheetProj.getRange(linhaExistente, idxJsonMerge + 1).getValue();
            let parsedMerge = null;
            try {
              parsedMerge = (cellMerge && String(cellMerge).trim()) ? JSON.parse(String(cellMerge).trim()) : null;
            } catch (eM) { parsedMerge = null; }
            const dadosAnt = parsedMerge && parsedMerge.dados ? parsedMerge.dados : null;
            const infoAnt = (dadosAnt && dadosAnt.infoPedido) ? dadosAnt.infoPedido : (parsedMerge && parsedMerge.infoPedido ? parsedMerge.infoPedido : null);
            if (infoAnt && typeof infoAnt === "object") {
              if (!dadosParaJson.infoPedido) dadosParaJson.infoPedido = {};
              var chavesPreservarInfoPedido = ["dataVirouPedido", "dataEntrega", "notaFiscal", "valorPedido", "comprovanteEntregaUrl", "temComprovanteEntrega", "dataFimProducao"];
              for (var ki = 0; ki < chavesPreservarInfoPedido.length; ki++) {
                var ck = chavesPreservarInfoPedido[ki];
                if (infoAnt[ck] != null && infoAnt[ck] !== "") {
                  dadosParaJson.infoPedido[ck] = infoAnt[ck];
                }
              }
              if (infoAnt.statusDates && typeof infoAnt.statusDates === "object") {
                if (!dadosParaJson.infoPedido.statusDates) dadosParaJson.infoPedido.statusDates = {};
                Object.keys(infoAnt.statusDates).forEach(function (sk) {
                  if (infoAnt.statusDates[sk] != null && infoAnt.statusDates[sk] !== "") {
                    dadosParaJson.infoPedido.statusDates[sk] = infoAnt.statusDates[sk];
                  }
                });
              }
            }
          }
        }
      }
    } catch (eMergeIp) {
      Logger.log("Aviso registrarOrcamento: falha ao preservar infoPedido (pedido já convertido): " + (eMergeIp && eMergeIp.message ? eMergeIp.message : eMergeIp));
    }

    let dadosJson = JSON.stringify({
      nome: codigoProjeto,
      dataSalvo: agora.toISOString(),
      numeroSequencial: numeroSequencial,
      dados: dadosParaJson
    });

    if (sheetProj && codigoProjeto && !/_v\d+$/i.test(String(codigoProjeto))) {
      const linhasMesmoCodigo = _listarLinhasProjetoPorCodigo(sheetProj, codigoProjeto, false);
      const linhaAtualConhecida = linhaExistente ? Number(linhaExistente) : 0;
      const conflito = linhasMesmoCodigo.filter(function (ln) { return ln !== linhaAtualConhecida; });
      if (conflito.length > 0) {
        throw new Error("Já existe projeto com o código '" + codigoProjeto + "' nas linhas " + conflito.join(", ") + ". O sistema bloqueou a gravação para evitar duplicidade.");
      }
    }

    try {
      if (linhaExistente && sheetProj) {
        const hdrsNow = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
        const idxJsonNow = _findHeaderIndex(hdrsNow, "JSON_DADOS");
        if (idxJsonNow >= 0) {
          const cellOld = sheetProj.getRange(linhaExistente, idxJsonNow + 1).getValue();
          if (cellOld && typeof cellOld === "string" && cellOld.trim()) {
            let parsedOld = null;
            try { parsedOld = JSON.parse(String(cellOld).trim()); } catch (eOld) { parsedOld = null; }
            if (parsedOld) {
              const versoesTop = parsedOld.versoes;
              const versoesAlt = (parsedOld.dados && parsedOld.dados.versoes !== undefined) ? parsedOld.dados.versoes : undefined;
              const versoesParaPreservar = (versoesTop !== undefined) ? versoesTop : versoesAlt;

              if (versoesParaPreservar !== undefined) {
                let dadosJsonObj = {};
                try { dadosJsonObj = JSON.parse(dadosJson); } catch (eParseJson) { dadosJsonObj = {}; }
                dadosJsonObj.versoes = versoesParaPreservar;
                if (!dadosJsonObj.urlPdf && urlPdf) dadosJsonObj.urlPdf = urlPdf;
                dadosJson = JSON.stringify(dadosJsonObj);
              }
            }
          }
        }
      }
    } catch (ePres) {
      Logger.log("Aviso: falha ao preservar versoes ao registrarOrcamento: " + (ePres && ePres.message ? ePres.message : ePres));
    }

    const statusOrcamento = isPedido ? "Convertido em Pedido" : "Enviado";
    const statusPedidoInicial = isPedido ? "Processo de Preparação MP / CAD / CAM" : "";
    const observacoesKanban = (observacoes && observacoes.observacoesKanban != null) ? String(observacoes.observacoesKanban).trim() : "";
    const prazoProposta = (observacoes && observacoes.prazoProposta != null) ? String(observacoes.prazoProposta).trim() : "";

    const nomeVersaoNovoBase = (dadosFormularioCompleto && dadosFormularioCompleto.nomeVersao != null)
      ? String(dadosFormularioCompleto.nomeVersao).trim()
      : "";
    if (linhaExistente && nomeVersaoNovoBase) {
      try {
        const headersBase = sheetProj.getRange(1, 1, 1, sheetProj.getLastColumn()).getValues()[0];
        const idxJsonBase = _findHeaderIndex(headersBase, "JSON_DADOS");
        if (idxJsonBase >= 0) {
          const cellJson = sheetProj.getRange(linhaExistente, idxJsonBase + 1).getValue();
          let parsedBase = {};
          try { parsedBase = JSON.parse(String(cellJson || "{}")); } catch (e) { parsedBase = {}; }
          const versoesBase = Array.isArray(parsedBase.versoes) ? parsedBase.versoes : [];
          for (let i = 0; i < versoesBase.length; i++) {
            const v = versoesBase[i];
            if (!v) continue;
            const nomeExist = (v.nomeVersao != null) ? String(v.nomeVersao).trim() : "";
            if (nomeExist && nomeExist === nomeVersaoNovoBase) {
              throw new Error("Já existe uma versão com o mesmo Nome da Versão neste projeto.");
            }
          }
        }
      } catch (eDup) {
        throw eDup;
      }
    }

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
      "PRAZO": prazoParaPlanilha,
      "PRAZO_PROPOSTA": prazoProposta,
      "OBSERVAÇÕES": observacoesKanban,
      "JSON_DADOS": dadosJson
    };
    var jaEraConvertidoEmPedido = false;
    if (linhaExistente) {
      var rowAtualReg = targetSheet.getRange(linhaExistente, 1, linhaExistente, targetSheet.getLastColumn()).getValues()[0];
      var headersReg = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
      var idxData = _findHeaderIndexProjetos(headersReg, "DATA");
      var statusAtualReg = "";
      var idxOrc = _findHeaderIndexProjetos(headersReg, "STATUS_ORCAMENTO");
      if (idxOrc >= 0 && rowAtualReg[idxOrc]) statusAtualReg = String(rowAtualReg[idxOrc] || "").trim();
      if (statusAtualReg === "Convertido em Pedido") jaEraConvertidoEmPedido = true;
      if (statusAtualReg === "Convertido em Pedido" && idxData >= 0 && rowAtualReg[idxData] != null && String(rowAtualReg[idxData]).trim() !== "") {
        dadosObjReg["DATA"] = rowAtualReg[idxData];
      }
      if (!isPedido) {
        var idxPed = _findHeaderIndexProjetos(headersReg, "STATUS_PEDIDO");
        var idxObs = _findHeaderIndexProjetos(headersReg, "OBSERVAÇÕES");
        if (idxOrc >= 0 && rowAtualReg[idxOrc]) dadosObjReg["STATUS_ORCAMENTO"] = rowAtualReg[idxOrc];
        if (idxPed >= 0 && rowAtualReg[idxPed]) dadosObjReg["STATUS_PEDIDO"] = rowAtualReg[idxPed];
        if (idxObs >= 0 && observacoesKanban === "" && rowAtualReg[idxObs]) dadosObjReg["OBSERVAÇÕES"] = rowAtualReg[idxObs];
      } else if (statusAtualReg === "Convertido em Pedido") {
        var idxPedEx = _findHeaderIndexProjetos(headersReg, "STATUS_PEDIDO");
        var idxObsEx = _findHeaderIndexProjetos(headersReg, "OBSERVAÇÕES");
        var idxPrazoEx = _findHeaderIndexProjetos(headersReg, "PRAZO");
        var idxPrazoPropEx = _findHeaderIndexProjetos(headersReg, "PRAZO_PROPOSTA");
        if (idxOrc >= 0 && rowAtualReg[idxOrc]) dadosObjReg["STATUS_ORCAMENTO"] = rowAtualReg[idxOrc];
        if (idxPedEx >= 0 && rowAtualReg[idxPedEx]) dadosObjReg["STATUS_PEDIDO"] = rowAtualReg[idxPedEx];
        if (idxObsEx >= 0 && observacoesKanban === "" && rowAtualReg[idxObsEx]) dadosObjReg["OBSERVAÇÕES"] = rowAtualReg[idxObsEx];
        if (idxPrazoEx >= 0 && rowAtualReg[idxPrazoEx] != null && String(rowAtualReg[idxPrazoEx]).trim() !== "") dadosObjReg["PRAZO"] = rowAtualReg[idxPrazoEx];
        if (idxPrazoPropEx >= 0 && rowAtualReg[idxPrazoPropEx] != null && String(rowAtualReg[idxPrazoPropEx]).trim() !== "") dadosObjReg["PRAZO_PROPOSTA"] = rowAtualReg[idxPrazoPropEx];
      }
      _escreverLinhaProjetosPorCabecalho(targetSheet, linhaExistente, dadosObjReg, true);
    } else {
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
              if (statusDup === "Convertido em Pedido") jaEraConvertidoEmPedido = true;
              if (statusDup === "Convertido em Pedido" && idxDataDup >= 0 && rowAtualDup[idxDataDup] != null && String(rowAtualDup[idxDataDup]).trim() !== "") dadosObjReg["DATA"] = rowAtualDup[idxDataDup];
              if (!isPedido && idxOrcDup >= 0 && rowAtualDup[idxOrcDup]) dadosObjReg["STATUS_ORCAMENTO"] = rowAtualDup[idxOrcDup];
              if (isPedido && statusDup === "Convertido em Pedido") {
                var idxPedDup = _findHeaderIndexProjetos(headersProj2, "STATUS_PEDIDO");
                var idxObsDup = _findHeaderIndexProjetos(headersProj2, "OBSERVAÇÕES");
                var idxPrazoDup = _findHeaderIndexProjetos(headersProj2, "PRAZO");
                var idxPrazoPropDup = _findHeaderIndexProjetos(headersProj2, "PRAZO_PROPOSTA");
                if (idxOrcDup >= 0 && rowAtualDup[idxOrcDup]) dadosObjReg["STATUS_ORCAMENTO"] = rowAtualDup[idxOrcDup];
                if (idxPedDup >= 0 && rowAtualDup[idxPedDup]) dadosObjReg["STATUS_PEDIDO"] = rowAtualDup[idxPedDup];
                if (idxObsDup >= 0 && observacoesKanban === "" && rowAtualDup[idxObsDup]) dadosObjReg["OBSERVAÇÕES"] = rowAtualDup[idxObsDup];
                if (idxPrazoDup >= 0 && rowAtualDup[idxPrazoDup] != null && String(rowAtualDup[idxPrazoDup]).trim() !== "") dadosObjReg["PRAZO"] = rowAtualDup[idxPrazoDup];
                if (idxPrazoPropDup >= 0 && rowAtualDup[idxPrazoPropDup] != null && String(rowAtualDup[idxPrazoPropDup]).trim() !== "") dadosObjReg["PRAZO_PROPOSTA"] = rowAtualDup[idxPrazoPropDup];
              }
              _escreverLinhaProjetosPorCabecalho(targetSheet, ri, dadosObjReg, true);
              break;
            }
          }
        }
      }
      if (!linhaExistente) _escreverLinhaProjetosPorCabecalho(targetSheet, targetSheet.getLastRow() + 1, dadosObjReg, false);
    }

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

    if (isPedido && sheetProj && (codigoProjeto || "").length >= 6) {
      try {
        if (!jaEraConvertidoEmPedido) {
          const codigoBase = String(codigoProjeto).replace(/_v\d+$/, "").trim();
          const dataProj = codigoBase.substring(0, 6);
          const nomeAbrev = (dadosFormularioCompleto && dadosFormularioCompleto.cliente && dadosFormularioCompleto.cliente.nomeAbreviado) || "";
          atualizarPrefixoPastaParaPedido(codigoBase, dataProj, cliente.nome || "", descricao, nomeAbrev);
        }
      } catch (ePasta) {
        Logger.log("Aviso ao renomear pasta COT→PED em registrarOrcamento: " + (ePasta && ePasta.message));
      }
      const linhaPedido = linhaExistente || (targetSheet ? targetSheet.getLastRow() : 0);
      if (linhaPedido >= 2) ensurePedidoRow(linhaPedido);
    }

    _sincronizarProdutosNaRelacao(produtosCadastrados, codigoProjeto, cliente && cliente.nome);

    try {
      if (cliente && cliente.nome) atualizarClienteSalvo(cliente);
    } catch (eCliente) {
      Logger.log("Aviso registrarOrcamento: falha ao atualizar cadastro do cliente: " + (eCliente && eCliente.message));
    }

  } catch (err) {
    Logger.log("Erro ao registrarOrcamento (atualizar/inserir): " + err);
    try {
      const numeroSequencial = (dadosFormularioCompleto && dadosFormularioCompleto.numeroSequencial) || null;
      const dadosParaJson = dadosFormularioCompleto ? { ...dadosFormularioCompleto } : {};
      const codigoBaseFallback = String(codigoProjeto || "").replace(/_v\d+$/i, "").trim() || String(codigoProjeto || "").trim();
      dadosParaJson.chapas = chapas || [];
      dadosParaJson.produtosCadastrados = produtosCadastrados || [];
      if (numeroSequencial != null) dadosParaJson.numeroSequencial = numeroSequencial;
      const agora = new Date();
      const dadosJson = JSON.stringify({
        nome: codigoBaseFallback,
        dataSalvo: agora.toISOString(),
        numeroSequencial: numeroSequencial,
        dados: dadosParaJson
      });
      const observacoesKanbanFallback = (observacoes && observacoes.observacoesKanban != null) ? String(observacoes.observacoesKanban).trim() : "";
      const prazoPropostaFallback = (observacoes && observacoes.prazoProposta != null) ? String(observacoes.prazoProposta).trim() : "";
      const sheetProj = SHEET_PROJ;
      if (sheetProj) {
        var dadosObjFallback = {
          "CLIENTE": cliente.nome || "",
          "DESCRIÇÃO": descricao,
          "RESPONSÁVEL CLIENTE": cliente.responsavel || "",
          "PROJETO": codigoBaseFallback || "",
          "VALOR TOTAL": valorTotal || "",
          "DATA": dataOrcamento || "",
          "PROCESSOS": processosStr || "",
          "LINK DO PDF": urlPdf || "",
          "LINK DA MEMÓRIA DE CÁLCULO": urlMemoria || "",
          "STATUS_ORCAMENTO": isPedido ? "Convertido em Pedido" : "Enviado",
          "STATUS_PEDIDO": isPedido ? "Processo de Preparação MP / CAD / CAM" : "",
          "PRAZO": prazoParaPlanilha,
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
      _sincronizarProdutosNaRelacao(produtosCadastrados, codigoProjeto, cliente && cliente.nome);

    } catch (e2) {
      Logger.log("Erro fallback appendRow em registrarOrcamento: " + e2);
      throw e2;
    }
  }
}

function _listarLinhasProjetoPorCodigo(sheetProj, numeroProjeto, compararCodigoBase) {
  if (!sheetProj || !numeroProjeto) return [];
  const lastRow = sheetProj.getLastRow();
  if (lastRow < 2) return [];
  const lastCol = sheetProj.getLastColumn();
  const headers = sheetProj.getRange(1, 1, 1, lastCol).getValues()[0];
  const idxProjeto = _findHeaderIndexProjetos(headers, "PROJETO");
  if (idxProjeto < 0) return [];
  const alvo = String(numeroProjeto || "").trim();
  const alvoBase = alvo.replace(/_v\d+$/i, "").trim().toLowerCase();
  const valores = sheetProj.getRange(2, idxProjeto + 1, lastRow - 1, 1).getValues();
  const linhas = [];
  for (var i = 0; i < valores.length; i++) {
    const codigoLinha = String(valores[i][0] || "").trim();
    if (!codigoLinha) continue;
    const bate = compararCodigoBase
      ? codigoLinha.replace(/_v\d+$/i, "").trim().toLowerCase() === alvoBase
      : codigoLinha.toLowerCase() === alvo.toLowerCase();
    if (bate) linhas.push(i + 2);
  }
  return linhas;
}

function verificarProjetoDuplicado(numeroProjeto) {
  try {
    if (!numeroProjeto) {
      return { duplicado: false, linha: null, onde: "" };
    }

    const sheetProj = SHEET_PROJ;
    if (sheetProj) {
      const linhas = _listarLinhasProjetoPorCodigo(sheetProj, numeroProjeto, false);
      if (linhas.length > 0) {
        return { duplicado: true, linha: linhas[0], linhas: linhas, onde: "Projetos" };
      }
    }
    return { duplicado: false, linha: null, linhas: [], onde: "" };
  } catch (err) {
    Logger.log("Erro ao verificar projeto duplicado: " + err.message);
    return { duplicado: false, linha: null, linhas: [], onde: "", erro: err.message };
  }
}
