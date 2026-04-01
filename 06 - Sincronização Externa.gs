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
    
    // 1. Cria o mapa de dados a partir da planilha externa
    for (var i = 1; i < dadosExternos.length; i++) {
      var itemBruto = dadosExternos[i][1];
      if (itemBruto) {
        var itemKey = String(itemBruto).trim().toLowerCase();
        var saldoAtaAtual = dadosExternos[i][17] !== undefined ? dadosExternos[i][17] : ""; 
        var vencAtaAtual = dadosExternos[i][19] !== undefined ? dadosExternos[i][19] : "";
        
        if (!mapaDados[itemKey]) {
          mapaDados[itemKey] = {
            descricao: dadosExternos[i][2] !== undefined ? dadosExternos[i][2] : "",       
            saldo: dadosExternos[i][6] !== undefined ? dadosExternos[i][6] : 0,            
            cmm: dadosExternos[i][7] !== undefined ? dadosExternos[i][7] : 0,              
            saldoDias: dadosExternos[i][8] !== undefined ? dadosExternos[i][8] : "",       
            cobertura: dadosExternos[i][9] !== undefined ? dadosExternos[i][9] : "",       
            saldoAta: saldoAtaAtual,      
            vencAta: vencAtaAtual,       
            processo: dadosExternos[i][38] !== undefined ? dadosExternos[i][38] : ""       
          };
        } else {
          if (String(saldoAtaAtual).trim() !== "") {
            mapaDados[itemKey].saldoAta = saldoAtaAtual;
          }
          if (String(vencAtaAtual).trim() !== "") {
            mapaDados[itemKey].vencAta = vencAtaAtual;
          }
        }
      }
    }
    
    var planilhaLocal = SpreadsheetApp.getActiveSpreadsheet();
    
    // 2. ATUALIZA A GUIA "ITENS"
    var guiaItens = planilhaLocal.getSheetByName('Itens');
    var dadosItens = guiaItens.getDataRange().getValues();
    var itensAtualizados = 0;
    
    for (var j = 2; j < dadosItens.length; j++) { // Inicia no índice 2 (Linha 3)
      var itemLocalBruto = dadosItens[j][0];
      if (itemLocalBruto) {
        var itemLocalKey = String(itemLocalBruto).trim().toLowerCase();
        
        if (mapaDados[itemLocalKey] !== undefined) {
          var info = mapaDados[itemLocalKey];
          var saldoFormat = parseBRNumber(info.saldo);
          var cmmFormat = parseBRNumber(info.cmm);
          
          if (cmmFormat === 0) {
            var cmmLocal = parseBRNumber(dadosItens[j][3]);
            if (cmmLocal > 0) { cmmFormat = cmmLocal; info.cmm = cmmLocal; }
          }
          
          guiaItens.getRange(j + 1, 2).setValue(info.descricao);
          guiaItens.getRange(j + 1, 3).setValue(info.saldo);     
          guiaItens.getRange(j + 1, 4).setValue(info.cmm);       
          guiaItens.getRange(j + 1, 5).setValue(info.saldoDias); 
          guiaItens.getRange(j + 1, 6).setValue(info.cobertura);
          
          var riscoCalc = obterClassificacaoRisco(saldoFormat, cmmFormat);
          guiaItens.getRange(j + 1, 7).setValue(riscoCalc.texto).setBackground(riscoCalc.cor);

          var celulaProcesso = guiaItens.getRange(j + 1, 8);
          var processoAtual = String(dadosItens[j][7]).trim();
          var novoProcesso = String(info.processo).trim();
          
          if (novoProcesso !== "") {
            if (processoAtual === "") {
              celulaProcesso.setValue(novoProcesso).setFontWeight("bold").setFontColor("red");
            } else if (processoAtual.indexOf(novoProcesso) === -1) {
              var textoFinal = processoAtual + '\n' + novoProcesso;
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
          
          guiaItens.getRange(j + 1, 11).setValue(info.vencAta);
          guiaItens.getRange(j + 1, 12).setValue(info.saldoAta);
          
          var celulaDestino = guiaItens.getRange(j + 1, 9);
          var valorAtual = String(celulaDestino.getValue());
          
          if (saldoFormat > 0) {
            celulaDestino.setValue('Sim').setBackground('#d9ead3');
          } else {
             if (valorAtual === 'Sim' || valorAtual === 'Sim') {
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
        itensAtualizados + ' itens foram atualizados na aba Itens.\n' + 
        compAtualizados + ' itens receberam dados de Ata na aba Comparativo.', ui.ButtonSet.OK);
    }
    
  } catch (erro) {
    ui.alert('Aviso', 'Erro: ' + erro.message, ui.ButtonSet.OK);
  }
}