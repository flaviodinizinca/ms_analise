// FUNÇÃO: Consolida processos da listagem com proteção visual RichText
function sincronizarProcessosListagem() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaListagem = planilha.getSheetByName('listagem.total');
  var guiaItens = planilha.getSheetByName('Itens');
  var ui = SpreadsheetApp.getUi();

  if (!guiaListagem || !guiaItens) return;
  var dadosListagem = guiaListagem.getDataRange().getValues();
  var mapaProcessos = {};

  for (var i = 1; i < dadosListagem.length; i++) {
    var itemBruto = dadosListagem[i][0];
    if (itemBruto) {
      var itemKey = String(itemBruto).toLowerCase().trim();
      var processoBruto = String(dadosListagem[i][6]).trim();
      if (processoBruto !== '' && processoBruto !== '-') {
        if (!mapaProcessos[itemKey]) mapaProcessos[itemKey] = [];
        if (mapaProcessos[itemKey].indexOf(processoBruto) === -1) {
          mapaProcessos[itemKey].push(processoBruto);
        }
      }
    }
  }

  var dadosItens = guiaItens.getDataRange().getValues();
  var itensAtualizados = 0;

  for (var j = 2; j < dadosItens.length; j++) { // Inicia no índice 2 (Linha 3)
    var itemItensBruto = dadosItens[j][0];
    if (itemItensBruto) {
      var itemItensKey = String(itemItensBruto).toLowerCase().trim();
      if (mapaProcessos[itemItensKey] && mapaProcessos[itemItensKey].length > 0) {
        var processosAtuaisStr = String(dadosItens[j][7]).trim();
        var celulaDestino = guiaItens.getRange(j + 1, 8);
        var arrayAtuais = processosAtuaisStr !== "" ? processosAtuaisStr.split('\n') : [];
        var arrayNovosParaDestacar = [];

        var novosProcessos = mapaProcessos[itemItensKey];
        for (var p = 0; p < novosProcessos.length; p++) {
          var proc = novosProcessos[p].trim();
          var jaExiste = false;
          for (var a = 0; a < arrayAtuais.length; a++) {
            if (arrayAtuais[a].trim() === proc) { jaExiste = true; break; }
          }
          if (!jaExiste) {
            arrayNovosParaDestacar.push(proc);
          }
        }

        if (arrayNovosParaDestacar.length > 0) {
          var textoNovo = arrayNovosParaDestacar.join('\n');
          if (processosAtuaisStr === "") {
            celulaDestino.setValue(textoNovo).setFontWeight("bold").setFontColor("red");
          } else {
            var textoFinal = processosAtuaisStr + '\n' + textoNovo;
            var estiloPadrao = SpreadsheetApp.newTextStyle().setBold(false).setForegroundColor("black").build();
            var estiloDestaque = SpreadsheetApp.newTextStyle().setBold(true).setForegroundColor("red").build();
            
            var richText = SpreadsheetApp.newRichTextValue()
                .setText(textoFinal)
                .setTextStyle(0, processosAtuaisStr.length, estiloPadrao)
                .setTextStyle(processosAtuaisStr.length + 1, textoFinal.length, estiloDestaque)
                .build();
            celulaDestino.setRichTextValue(richText);
          }
          itensAtualizados++;
        }
      }
    }
  }
  
  if (itensAtualizados > 0) {
    ui.alert('Consolidação Concluída', itensAtualizados + ' itens receberam novos processos (destacados em vermelho na Coluna H).', ui.ButtonSet.OK);
  } else {
    ui.alert('Consolidação Concluída', 'Nenhum novo processo precisou ser atualizado.', ui.ButtonSet.OK);
  }
}