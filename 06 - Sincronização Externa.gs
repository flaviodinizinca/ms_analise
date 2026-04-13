// FUNÇÃO: Sincronização Externa (Aplica risco, anexa novos processos e alimenta aba Comparativo)
function sincronizarDadosGerais() {
  var ui = SpreadsheetApp.getUi();
  try {
    var idPlanilhaExterna = '1s44YD2ozLAbBdGQbBE5iW7HcUzvQULZqd4ynYlV_HXA';
    var planilhaExterna = SpreadsheetApp.openById(idPlanilhaExterna);
    var guiaDados = planilhaExterna.getSheetByName('DadosEstoque');
    
    if (!guiaDados) return;
    var dadosExternos = guiaDados.getDataRange().getValues();
    var mapaDados = {};
    
    // 1. Cria o mapa de dados a partir da planilha externa (Estoque) - Nova Estrutura 37 Colunas
    for (var i = 1; i < dadosExternos.length; i++) {
      var itemBruto = dadosExternos[i][1]; // Coluna B (Índice 1) - Item
      if (itemBruto) {
        var itemKey = String(itemBruto).trim().toLowerCase();
        var saldoAtaAtual = dadosExternos[i][21] !== undefined ? dadosExternos[i][21] : "";  // Coluna V (Índice 21) - SaldoAta
        var vencAtaAtual = dadosExternos[i][23] !== undefined ? dadosExternos[i][23] : "";   // Coluna X (Índice 23) - DataVencAta
        var procS = dadosExternos[i][18] !== undefined ? String(dadosExternos[i][18]).trim() : ""; // Coluna S (Índice 18) - NumProcAtaVig
        var procY = dadosExternos[i][25] !== undefined ? String(dadosExternos[i][25]).trim() : ""; // Coluna Z (Índice 25) - NumProcAnalise
        
        var arrayProc = [];
        if (procS !== "" && procS !== "-") arrayProc.push(procS);
        if (procY !== "" && procY !== "-" && procY !== procS) arrayProc.push(procY);
        var novoProcessoExterno = arrayProc.join('\n');
        
        if (!mapaDados[itemKey]) {
          mapaDados[itemKey] = {
            descricao: dadosExternos[i][2] !== undefined ? dadosExternos[i][2] : "",       // Coluna C (Índice 2)
            saldo: dadosExternos[i][7] !== undefined ? dadosExternos[i][7] : 0,            // Coluna H (Índice 7)
            cmm: dadosExternos[i][8] !== undefined ? dadosExternos[i][8] : 0,              // Coluna I (Índice 8)
            saldoDias: dadosExternos[i][10] !== undefined ? dadosExternos[i][10] : "",     // Coluna K (Índice 10)
            cobertura: dadosExternos[i][11] !== undefined ? dadosExternos[i][11] : "",     // Coluna L (Índice 11)
            saldoAta: saldoAtaAtual,      
            vencAta: vencAtaAtual,       
            processo: novoProcessoExterno
          };
        } else {
          if (String(saldoAtaAtual).trim() !== "") {
            mapaDados[itemKey].saldoAta = saldoAtaAtual;
          }
          if (String(vencAtaAtual).trim() !== "") {
            mapaDados[itemKey].vencAta = vencAtaAtual;
          }
          if (novoProcessoExterno !== "") {
            var procExistente = mapaDados[itemKey].processo;
            if (procExistente) {
              var processosJuntos = procExistente.split('\n');
              for (var p = 0; p < arrayProc.length; p++) {
                if (processosJuntos.indexOf(arrayProc[p]) === -1) {
                  processosJuntos.push(arrayProc[p]);
                }
              }
              mapaDados[itemKey].processo = processosJuntos.join('\n');
            } else {
              mapaDados[itemKey].processo = novoProcessoExterno;
            }
          }
        }
      }
    }

    // 1.5 Cria o mapa de dados a partir da planilha "Compilados" (Recebimento Provisório)
    var idPlanilhaCompilados = '1ZLebBqhR1bMZgrnr_dfXikyIY22oi0B2pqXDz1UdRZM';
    var planilhaCompilados = SpreadsheetApp.openById(idPlanilhaCompilados);
    var guiaCompilados = planilhaCompilados.getSheetByName('Compilados');
    var mapaRecebimento = {};
    if (guiaCompilados) {
      var dadosCompilados = guiaCompilados.getDataRange().getValues();
      for (var c = 1; c < dadosCompilados.length; c++) {
        var statusComp = String(dadosCompilados[c][18]).trim(); // Coluna S (índice 18)
        if (statusComp === "Recebimento Provisório") {
          var itemCompBruto = dadosCompilados[c][5]; // Coluna F (índice 5)
          if (itemCompBruto) {
            var itemCompKey = String(itemCompBruto).trim().toLowerCase();
            var qtdRecebida = parseBRNumber(dadosCompilados[c][15]); // Coluna P (índice 15)
            if (qtdRecebida > 0) {
              mapaRecebimento[itemCompKey] = (mapaRecebimento[itemCompKey] || 0) + qtdRecebida;
            }
          }
        }
      }
    }
    
    var planilhaLocal = SpreadsheetApp.getActiveSpreadsheet();
    // Carrega o mapa de Prazos lendo a guia CadastroStatus
    var guiaCadastro = planilhaLocal.getSheetByName('CadastroStatus');
    var mapaPrazos = {};
    if (guiaCadastro) {
      var dadosCad = guiaCadastro.getDataRange().getValues();
      for (var idxCad = 0; idxCad < dadosCad.length; idxCad++) {
        var statusChave = String(dadosCad[idxCad][0]).trim().toLowerCase();
        if (statusChave !== "") {
          mapaPrazos[statusChave] = dadosCad[idxCad][1];
        }
      }
    }
    
    // 2. ATUALIZA A GUIA "ITENS"
    var guiaItens = planilhaLocal.getSheetByName('Itens');
    var dadosItens = guiaItens.getDataRange().getValues();
    var itensAtualizados = 0;
    
    for (var j = 2; j < dadosItens.length; j++) { // Inicia no índice 2 (Linha 3)
      var itemLocalBruto = dadosItens[j][0];
      if (itemLocalBruto) {
        var itemLocalKey = String(itemLocalBruto).trim().toLowerCase();
        // Aplica o Prazo na Coluna N (Índice 13), lendo o Status da Coluna J (Índice 9)
        var statusAtualItem = String(dadosItens[j][9]).trim();
        if (statusAtualItem !== "") {
          var prazoCalculado = mapaPrazos[statusAtualItem.toLowerCase()] !== undefined ? mapaPrazos[statusAtualItem.toLowerCase()] : "";
          guiaItens.getRange(j + 1, 14).setValue(prazoCalculado); // Escreve na Coluna N
        } else {
          guiaItens.getRange(j + 1, 14).setValue("");
        }
        
        if (mapaDados[itemLocalKey] !== undefined) {
          var info = mapaDados[itemLocalKey];
          var saldoEstoqueFormat = parseBRNumber(info.saldo);
          var cmmFormat = parseBRNumber(info.cmm);
          
          // SOMA o saldo do estoque com a quantidade em "Recebimento Provisório"
          var qtdRecebimentoProvisorio = mapaRecebimento[itemLocalKey] || 0;
          var saldoFinalReal = saldoEstoqueFormat + qtdRecebimentoProvisorio;

          if (cmmFormat === 0) {
            var cmmLocal = parseBRNumber(dadosItens[j][3]);
            if (cmmLocal > 0) { cmmFormat = cmmLocal; info.cmm = cmmLocal; }
          }
          
          guiaItens.getRange(j + 1, 2).setValue(info.descricao);
          guiaItens.getRange(j + 1, 3).setValue(saldoFinalReal);     // Saldo = Estoque + Recebimento Provisório
          guiaItens.getRange(j + 1, 4).setValue(info.cmm);
          guiaItens.getRange(j + 1, 5).setValue(info.saldoDias); 
          guiaItens.getRange(j + 1, 6).setValue(info.cobertura);
          
          var riscoCalc = obterClassificacaoRisco(saldoFinalReal, cmmFormat);
          guiaItens.getRange(j + 1, 7).setValue(riscoCalc.texto).setBackground(riscoCalc.cor);
          var celulaProcesso = guiaItens.getRange(j + 1, 8);
          var processoAtual = String(dadosItens[j][7]).trim();
          var novoProcesso = String(info.processo).trim();
          if (novoProcesso !== "") {
            if (processoAtual === "") {
              celulaProcesso.setValue(novoProcesso).setFontWeight("bold").setFontColor("red");
            } else {
              var processosAtuaisArr = processoAtual.split('\n');
              var novosProcessosArr = novoProcesso.split('\n');
              var processosParaAdicionar = [];

              for (var np = 0; np < novosProcessosArr.length; np++) {
                var pNovo = novosProcessosArr[np].trim();
                var jaExiste = false;
                for (var pa = 0; pa < processosAtuaisArr.length; pa++) {
                  if (processosAtuaisArr[pa].trim() === pNovo) { jaExiste = true; break; }
                }
                if (!jaExiste && pNovo !== "") processosParaAdicionar.push(pNovo);
              }

              if (processosParaAdicionar.length > 0) {
                var textoAdicional = processosParaAdicionar.join('\n');
                var textoFinal = processoAtual + '\n' + textoAdicional;
                var estiloPadrao = SpreadsheetApp.newTextStyle().setBold(false).setForegroundColor("black").build();
                var estiloDestaque = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("red").build();
                var richText = SpreadsheetApp.newRichTextValue()
                    .setText(textoFinal)
                    .setTextStyle(0, processoAtual.length, estiloPadrao) 
                    .setTextStyle(processoAtual.length + 1, textoFinal.length, estiloDestaque)
                    .build();
                celulaProcesso.setRichTextValue(richText);
              }
            }
          }
          
          guiaItens.getRange(j + 1, 11).setValue(info.vencAta);
          guiaItens.getRange(j + 1, 12).setValue(info.saldoAta);
          
          var celulaDestino = guiaItens.getRange(j + 1, 9);
          var valorAtual = String(celulaDestino.getValue());
          if (saldoFinalReal > 0) {
            celulaDestino.setValue('Sim').setBackground('#d9ead3');
          } else {
             if (valorAtual === 'Sim' || valorAtual === 'Sim Órgão') {
                celulaDestino.setValue('Não').setBackground(null);
             } else if (valorAtual !== 'Sim Órgão' && valorAtual.indexOf('Parcial') === -1 && valorAtual !== 'Não') {
                celulaDestino.setValue('Não').setBackground(null);
             }
          }
          itensAtualizados++;
        }
      }
    }
    
    // 3. ATUALIZA A GUIA "COMPARATIVO" COM DADOS DA ATA
    var guiaComp = planilhaLocal.getSheetByName('Comparativo-Guia.Itens X listagem.total');
    var compAtualizados = 0;
    
    if (guiaComp) {
      var dadosComp = guiaComp.getDataRange().getValues();
      for (var k = 1; k < dadosComp.length; k++) {
        var itemCompBruto = dadosComp[k][0];
        if (itemCompBruto) {
          var itemCompKey = String(itemCompBruto).trim().toLowerCase();
          if (mapaDados[itemCompKey] !== undefined) {
            var infoComp = mapaDados[itemCompKey];
            guiaComp.getRange(k + 1, 10).setValue(infoComp.vencAta);
            guiaComp.getRange(k + 1, 11).setValue(infoComp.saldoAta);
            compAtualizados++;
          }
        }
      }
      aplicarRiscoNaGuia(guiaComp, 1);
    }

    if (itensAtualizados > 0 || compAtualizados > 0) {
      ui.alert('Sincronização Concluída', 
        itensAtualizados + ' itens foram atualizados na aba Itens (Estoque + Recebimentos Provisórios).\n' + 
        compAtualizados + ' itens receberam dados de Ata na aba Comparativo.', ui.ButtonSet.OK);
    }

    // 4. CHAMA O ARQUIVAMENTO AUTOMÁTICO SE O SALDO (Estoque+Recebimento) SUPERA O CMM
    // Força a planilha a aplicar as edições feitas acima antes de rodar a varredura
    SpreadsheetApp.flush();
    // Executa a função nativa para varrer os itens e arquivar os que bateram a meta de CMM
    arquivarItensResolvidos();
  } catch (erro) {
    ui.alert('Aviso', 'Erro: ' + erro.message, ui.ButtonSet.OK);
  }
}