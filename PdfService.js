function gerarPdfOrcamento(
  chapas, cliente, observacoes, codigoProjeto, nomePasta, dataProjeto, versao,
  somaProcessosPedido, descricaoProcessosPedido, produtosCadastrados,
  dadosFormularioCompleto, infoPagamento,
  isPedido, nomeVersao, sobrescreverVersao, apenasPreview
) {
  sobrescreverVersao = !!sobrescreverVersao;
  apenasPreview = !!apenasPreview;
  try {

    if (!apenasPreview) incrementarContador("totalPropostas");

    let numeroSequencial;
    if (apenasPreview) {
      numeroSequencial = (dadosFormularioCompleto && dadosFormularioCompleto.numeroSequencial != null && dadosFormularioCompleto.numeroSequencial !== "") ? dadosFormularioCompleto.numeroSequencial : "—";
    } else if (sobrescreverVersao) {
      const sheetProj = SHEET_PROJ;
      const linhaProj = sheetProj ? findRowByColumnValue(sheetProj, "PROJETO", codigoProjeto) : null;

      if (
        dadosFormularioCompleto &&
        dadosFormularioCompleto.numeroSequencial != null &&
        String(dadosFormularioCompleto.numeroSequencial).trim() !== ""
      ) {
        numeroSequencial = dadosFormularioCompleto.numeroSequencial;
      }

      if ((numeroSequencial == null || String(numeroSequencial).trim() === "") && linhaProj && sheetProj) {
        var numCols = sheetProj.getLastColumn();
        var headers = sheetProj.getRange(1, 1, 1, numCols).getValues()[0];
        var jsonIdx = _findHeaderIndex(headers, "JSON_DADOS");
        if (jsonIdx >= 0) {
          var jsonCell = sheetProj.getRange(linhaProj, jsonIdx + 1).getValue();
          try {
            var parsed = jsonCell ? JSON.parse(String(jsonCell).trim()) : null;
            if (parsed && parsed.numeroSequencial != null && String(parsed.numeroSequencial).trim() !== "") numeroSequencial = parsed.numeroSequencial;
            else if (parsed && parsed.dados && parsed.dados.numeroSequencial != null && String(parsed.dados.numeroSequencial).trim() !== "") numeroSequencial = parsed.dados.numeroSequencial;
          } catch (e) { }
        }
      }

      if (numeroSequencial == null || numeroSequencial === undefined || String(numeroSequencial).trim() === "") {
        var sheetPed = SHEET_PED;
        if (sheetPed && sheetPed.getLastRow() >= 2) {
          var linhaPed = findRowByColumnValue(sheetPed, "PROJETO", codigoProjeto);
          if (linhaPed) {
            var headersPed = sheetPed.getRange(1, 1, 1, sheetPed.getLastColumn()).getValues()[0];
            var rowPed = sheetPed.getRange(linhaPed, 1, linhaPed, sheetPed.getLastColumn()).getValues()[0];
            var aliasesNum = ["NUMERO_SEQUENCIAL", "NUMERO SEQUENCIAL", "NÚMERO SEQUENCIAL", "Nº", "N"];
            for (var a = 0; a < aliasesNum.length && (numeroSequencial == null || numeroSequencial === undefined || String(numeroSequencial).trim() === ""); a++) {
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
      }

      if (numeroSequencial == null || numeroSequencial === undefined || String(numeroSequencial).trim() === "")
        numeroSequencial = obterEIncrementarNumeroOrcamento();

      if (dadosFormularioCompleto) dadosFormularioCompleto.numeroSequencial = numeroSequencial;
    } else {
      numeroSequencial = obterEIncrementarNumeroOrcamento();
      if (dadosFormularioCompleto) {
        dadosFormularioCompleto.numeroSequencial = numeroSequencial;
      }
    }

    if (produtosCadastrados && Array.isArray(produtosCadastrados)) {
      atribuirPRDsUnicos(produtosCadastrados);
    }

    const resultados = [];

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

    const matchBase = String(codigoProjeto || "").match(/^(.+?)(_v\d+)$/);
    const codigoBase = matchBase ? matchBase[1] : (codigoProjeto || "");

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

    const logoFile = DriveApp.getFileById(ID_LOGO);
    const logoBlob = logoFile.getBlob();
    const logoBase64 = Utilities.base64Encode(logoBlob.getBytes());
    const logoMime = logoBlob.getContentType();

    const totalPecas = resultados.reduce((sum, p) => sum + (Number(p.precoTotal) || 0), 0);
    const totalFinal = totalPecas + (Number(somaProcessosPedido) || 0);

    const agora = new Date();
    const dataBrasil = formatarDataBrasil(agora);
    const horaBrasil = agora.toLocaleTimeString("pt-BR");

    let versaoFinal = "";
    let codigoParaPlanilha = codigoBase;

    if (sobrescreverVersao) {
      versaoFinal = (versao && String(versao).trim()) ? (versao.startsWith("_") ? versao : "_" + versao) : "";
      codigoParaPlanilha = codigoBase + versaoFinal;
    } else {
      const proximaVersao = detectarProximaVersao(codigoBase, dataProjeto);
      versaoFinal = proximaVersao ? "_" + proximaVersao : "";
      codigoParaPlanilha = codigoBase + versaoFinal;
    }

    const numeroProposta = codigoParaPlanilha;

    const nomeVersaoLimpo = limparNomeArquivo(nomeVersao || "");
    const nomeVersaoVazio = !nomeVersaoLimpo;
    const nomeVersaoSanitizadoParaArquivo = nomeVersaoVazio ? "" : nomeVersaoLimpo;
    let usarPrefixoPedido = !!isPedido;
    try {
      if (!usarPrefixoPedido) {
        const sheetProjPrefixo = SHEET_PROJ;
        if (sheetProjPrefixo) {
          const linhaHint = (dadosFormularioCompleto && dadosFormularioCompleto.linhaProjeto != null) ? parseInt(dadosFormularioCompleto.linhaProjeto, 10) : NaN;
          let linhaProjetoPrefixo = 0;
          if (!isNaN(linhaHint) && linhaHint >= 2 && linhaHint <= sheetProjPrefixo.getLastRow()) {
            const hHint = sheetProjPrefixo.getRange(1, 1, 1, sheetProjPrefixo.getLastColumn()).getValues()[0];
            const idxProjetoHint = _findHeaderIndexProjetos(hHint, "PROJETO");
            const codigoLinhaHint = idxProjetoHint >= 0 ? String(sheetProjPrefixo.getRange(linhaHint, idxProjetoHint + 1).getValue() || "").trim() : "";
            const codigoLinhaBase = codigoLinhaHint.replace(/_v\d+$/i, "").trim();
            if (codigoLinhaBase && codigoLinhaBase === codigoBase) {
              linhaProjetoPrefixo = linhaHint;
            }
          }
          if (!linhaProjetoPrefixo) {
            linhaProjetoPrefixo = findRowByColumnValue(sheetProjPrefixo, "PROJETO", codigoBase) || 0;
          }
          if (!linhaProjetoPrefixo && codigoParaPlanilha) {
            linhaProjetoPrefixo = findRowByColumnValue(sheetProjPrefixo, "PROJETO", codigoParaPlanilha) || 0;
          }
          if (linhaProjetoPrefixo) {
            const hPrefixo = sheetProjPrefixo.getRange(1, 1, 1, sheetProjPrefixo.getLastColumn()).getValues()[0];
            const idxStatusOrcPrefixo = _findHeaderIndexProjetos(hPrefixo, "STATUS_ORCAMENTO");
            const statusOrcPrefixo = idxStatusOrcPrefixo >= 0 ? String(sheetProjPrefixo.getRange(linhaProjetoPrefixo, idxStatusOrcPrefixo + 1).getValue() || "").trim() : "";
            if (String(statusOrcPrefixo).toLowerCase() === "convertido em pedido") usarPrefixoPedido = true;
          }
        }
      }
    } catch (ePrefixo) { }
    if (dadosFormularioCompleto) dadosFormularioCompleto.tipoPdf = usarPrefixoPedido ? "Pedido" : "Proposta";
    const prefixoPdf = usarPrefixoPedido ? "Pedido_" : "Proposta_";

    const headerColor = "#FF9933";
    const rowColor = "#FDF5E6";

    function calcularParcelas(textoPagamento, valorTotal) {
      if (!textoPagamento || textoPagamento.trim() === "") return null;
      const texto = textoPagamento.trim().replace(/\s+/g, " ");
      const textoUpper = texto.toUpperCase();
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
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; font-size:7pt;">' + _escHtml(p.codigo || "") + '</td>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; font-size:7pt;">' + _escHtml(p.descricao || "") + '</td>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">' + _escHtml(p.quantidade || 0) + '</td>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">' + _formatBrCurrency(p.precoUnitario || 0) + '</td>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">' + _formatBrCurrency(p.precoTotal || 0) + '</td>'
        + '</tr>';
    }).join('');

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
          + '<td colspan="4" bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">' + _escHtml(descricaoLinha) + '</td>'
          + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">' + _formatBrCurrency(valor) + '</td>'
          + '</tr>';
      });
    } else if (somaProcessosPedido && Number(somaProcessosPedido) !== 0) {
      processosPedidoRow = '<tr>'
        + '<td colspan="4" bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;"><strong>' + _escHtml(descricaoProcessosPedido || "") + '</strong></td>'
        + '<td bgcolor="' + rowColor + '" style="background:' + rowColor + '; padding:2px; border:0.1px solid #fff; text-align:right; font-size:7pt;">' + _formatBrCurrency(somaProcessosPedido) + '</td>'
        + '</tr>';
    }

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
          <td bgcolor="${rowColor}" style="background:${rowColor}; padding:2px; border:0.1px solid #fff; font-size:7pt; text-align:center;">${_formatBrCurrency(p.valor)}</td>
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
          body { font-family: Arial, sans-serif; font-size: 8pt; color: #000; margin: 0px; line-height:1.2; -webkit-font-smoothing:antialiased; }
          .header { display:flex; justify-content:space-between; align-items:center; margin-bottom:8px; }
          .logo { max-height:160px; }
          .company-info { text-align:right; font-size:8pt; }
          h2 { text-align:left; margin:15px 0 20px 0; font-size:12pt; }
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

        <h2>Proposta Comercial Nº ${_escHtml(numeroProposta)}_${_escHtml(numeroSequencial)}</h2><br>

        <h3>Informações do Cliente:</h3>
        <p style="margin-bottom:12px; font-size:9pt; line-height:1.3;">
          <p><strong>${_escHtml(cliente.nome)}</strong><br></p>
            CNPJ/CPF: ${_escHtml(cliente.cpf)}<br>
            ${_escHtml(cliente.endereco)}<br>
            <b>Telefone:</b> ${_escHtml(cliente.telefone)}<br>
            <b>Email:</b> ${_escHtml(cliente.email)}<br>
            <b>Responsável:</b> ${_escHtml(cliente.responsavel || "-")}
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

<div style="width:100%; text-align:right; margin-top:5px;">
  <table style="display:inline-block; border-collapse:collapse; width:100%; max-width:280px;">
    <tr>
      <td style="border:none; text-align:right; width:120px; background:#fff; padding:3px; font-weight:bold; font-size:8pt;">Subtotal:</td>
      <td style="border:none; text-align:right; background:${rowColor}; padding:3px; width:100px; font-weight:bold; font-size:8pt;">${_formatBrCurrency(totalPecas)}</td>
    </tr>
    <tr>
      <td style="border:none; text-align:right; background:#fff; padding:3px; font-weight:bold; font-size:8pt;">Total:</td>
      <td style="border:none; text-align:right; background:${rowColor}; padding:3px; width:100px; font-weight:bold; font-size:8pt;">${_formatBrCurrency(totalFinal)}</td>
    </tr>
  </table>
</div>

        ${tabelaParcelasHtml}

        <h3 style="margin-top:12px;">Informações da proposta</h3>
        <p style="font-size:8pt; line-height:1.25;">
          <b>Proposta Comercial - incluído em:</b> ${_escHtml(dataBrasil)} às ${_escHtml(horaBrasil)}<br>
            <b>Validade da Proposta:</b> 30 dias
          </p>

        <p style="font-size:8pt; line-height:1.25;">
          <b>Prazo de entrega:</b> ${_escHtml(formatarDataBrasil(observacoes.prazo) || "-")}<br>
          <b>Pagamento:</b> ${_escHtml(observacoes.pagamento || "-")}<br>
          <b>Vendedor:</b> ${_escHtml(observacoes.vendedor || "-")}<br>
          <b>Condições do Material:</b> ${_escHtml(observacoes.materialCond || "-")}<br>
          <b>Transporte:</b> ${_escHtml(observacoes.transporte || "-")}<br>
        </p>

        ${observacoes.adicional ? `<p style="font-size:8pt; line-height:1.25;"><b>Observações:</b><br>${_escHtml(observacoes.adicional)}</p>` : ""}

      </body>
      </html>
    `;

    if (apenasPreview) return { html: htmlContent };

    const blob = Utilities.newBlob(htmlContent, "text/html", "orcamento.html");
    const sufixoNomeVersaoArquivo = nomeVersaoVazio ? "" : "_" + nomeVersaoSanitizadoParaArquivo;
    const nomeArquivoPdf = prefixoPdf + codigoBase + "_" + numeroSequencial + sufixoNomeVersaoArquivo + ".pdf";
    const pdf = blob.getAs("application/pdf").setName(nomeArquivoPdf);

    let file;
    if (sobrescreverVersao) {
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

    const memoriaUrl = null;

    var urlPdfRetorno = file.getUrl();
    try {
      if (dadosFormularioCompleto && typeof dadosFormularioCompleto === "object") {
        dadosFormularioCompleto.nomeVersao = nomeVersaoVazio ? "" : nomeVersaoLimpo;
        dadosFormularioCompleto.tipoPdf = isPedido ? "Pedido" : "Proposta";
      }
      registrarOrcamento(cliente, codigoParaPlanilha, totalFinal, dataBrasil, urlPdfRetorno, memoriaUrl, chapas || [], observacoes, produtosCadastrados, dadosFormularioCompleto, isPedido);
    } catch (errReg) {
      Logger.log("Aviso gerarPdfOrcamento: PDF criado mas falha ao registrar na planilha - " + (errReg && errReg.message ? errReg.message : errReg));
    }
    return { url: urlPdfRetorno, nome: file.getName(), memoriaUrl: memoriaUrl };
  } catch (err) {
    Logger.log("ERRO gerarPdfOrcamento: " + err.toString());
    throw err;
  }
}
