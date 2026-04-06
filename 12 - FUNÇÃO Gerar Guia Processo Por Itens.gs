// FUNÇÃO: Gerar Guia "Processo Por Itens"
function gerarProcessoPorItens() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var nomeGuiaNova = 'Processo Por Itens';
  var guiaNova = planilha.getSheetByName(nomeGuiaNova);
  var ui = SpreadsheetApp.getUi();
  
  if (!guiaNova) return;
  var guiaListagem = planilha.getSheetByName('listagem.total');
  var guiaItens = planilha.getSheetByName('Itens');
  if (!guiaListagem || !guiaItens) return;
  
  function padLines(str, count) {
    var lines = (str || "").split('\n');
    while (lines.length < count) { lines.push(""); }
    return lines.join('\n');
  }

  var dadosItens = guiaItens.getDataRange().getValues();
  var mapItens = {};
  for (var i = 2; i < dadosItens.length; i++) { // Inicia no índice 2 (Linha 3)
    var itemBruto = dadosItens[i][0];
    if (itemBruto) {
      var itemKey = String(itemBruto).toLowerCase().trim();
      mapItens[itemKey] = String(dadosItens[i][1] || "").trim();
    }
  }

  var dadosListagem = guiaListagem.getDataRange().getValues();
  var mapProcessos = {};
  for (var j = 1; j < dadosListagem.length; j++) {
    var itemOriginal = dadosListagem[j][0];
    var processoOriginal = dadosListagem[j][6];
    var statusOriginal = dadosListagem[j][7];

    if (itemOriginal && processoOriginal && String(processoOriginal).trim() !== '-') {
      var itemKey = String(itemOriginal).toLowerCase().trim();
      var processos = String(processoOriginal).split('\n');
      var statusTexto = String(statusOriginal || "").trim();
      
      for (var p = 0; p < processos.length; p++) {
        var pRaw = processos[p].trim();
        if (pRaw) {
          if (!mapProcessos[pRaw]) mapProcessos[pRaw] = [];
          var itemJaExiste = false;
          for (var k = 0; k < mapProcessos[pRaw].length; k++) {
            if (mapProcessos[pRaw][k].key === itemKey) {
              itemJaExiste = true;
              break;
            }
          }
          if (!itemJaExiste) {
            mapProcessos[pRaw].push({
              key: itemKey, original: itemOriginal,
              desc: mapItens[itemKey] || "", status: statusTexto
            });
          }
        }
      }
    }
  }

  var output = [];
  for (var pRaw in mapProcessos) {
    var itensDoProcesso = mapProcessos[pRaw];
    var count = itensDoProcesso.length;
    var colB_arr = [], colD_arr = [];
    
    for (var i = 0; i < count; i++) {
      var it = itensDoProcesso[i];
      var linhaB = it.original + (it.desc ? " - " + it.desc : "");
      var linhaD = it.status;
      var linesB = linhaB.split('\n').length;
      var linesD = linhaD.split('\n').length;
      var maxLines = Math.max(linesB, linesD);
      colB_arr.push(padLines(linhaB, maxLines));
      colD_arr.push(padLines(linhaD, maxLines));
    }
    output.push([pRaw, colB_arr.join('\n\n'), count, colD_arr.join('\n\n')]);
  }

  output.sort(function(a, b) {
    if (a[0] < b[0]) return -1;
    if (a[0] > b[0]) return 1;
    return 0;
  });
  
  var ultimaLinha = guiaNova.getLastRow();
  if (ultimaLinha > 1) {
    guiaNova.getRange(2, 1, ultimaLinha - 1, 4).clearContent();
  }
  
  if (output.length > 0) {
    guiaNova.getRange(2, 1, output.length, 4).setValues(output);
    guiaNova.getRange(2, 1, output.length, 4).setVerticalAlignment('top');
    guiaNova.getRange(2, 2, output.length, 1).setWrap(true);
    guiaNova.getRange(2, 4, output.length, 1).setWrap(true);
  }
  ui.alert('Relatório Atualizado', 'A guia "Processo Por Itens" foi gerada.', ui.ButtonSet.OK);
}