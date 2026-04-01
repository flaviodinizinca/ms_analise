// FUNÇÃO TRANSFERIR STATUS ENTRE GUIAS
function transferirStatusPorProcesso() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaItens = planilha.getSheetByName('Itens');
  var guiaComparativo = planilha.getSheetByName('Comparativo-Guia.Itens X listagem.total');
  var guiaListagem = planilha.getSheetByName('listagem.total');
  
  if (!guiaItens || !guiaComparativo || !guiaListagem) return;
  
  var dadosItens = guiaItens.getDataRange().getValues();
  var mapaStatusPorProcesso = {};
  var processosLidos = 0;
  
  for (var i = 2; i < dadosItens.length; i++) { // Inicia no índice 2 (Linha 3)
    var processosBruto = String(dadosItens[i][7]).trim(); // Coluna H
    var statusObservacao = String(dadosItens[i][9]).trim(); // Coluna J
    
    if (processosBruto !== '' && statusObservacao !== '') {
      var matchProcessos = processosBruto.match(/\d{6}\/\d{4}/g);
      if (matchProcessos && matchProcessos.length > 0) {
        if (matchProcessos.length === 1) {
          mapaStatusPorProcesso[matchProcessos[0]] = statusObservacao;
          processosLidos++;
        } else {
          for (var p = 0; p < matchProcessos.length; p++) {
            if (matchProcessos[p].indexOf('/2024') !== -1) {
              mapaStatusPorProcesso[matchProcessos[p]] = statusObservacao;
              processosLidos++; break;
            }
          }
        }
      }
    }
  }
  
  function aplicarStatusNaGuiaDestino(guia, colProcesso, colStatus) {
    var dadosDestino = guia.getDataRange().getValues();
    var linhasAtualizadas = 0;
    for (var j = 1; j < dadosDestino.length; j++) { 
      var processosDestinoBruto = String(dadosDestino[j][colProcesso - 1]).trim();
      if (processosDestinoBruto !== '') {
        var matchDestino = processosDestinoBruto.match(/\d{6}\/\d{4}/g);
        if (matchDestino && matchDestino.length > 0) {
          for (var k = 0; k < matchDestino.length; k++) {
            var processoChave = matchDestino[k];
            if (mapaStatusPorProcesso[processoChave] !== undefined) {
              guia.getRange(j + 1, colStatus).setValue(mapaStatusPorProcesso[processoChave]);
              linhasAtualizadas++; break; 
            }
          }
        }
      }
    }
    return linhasAtualizadas;
  }
  
  var linhasAtualizadasComp = aplicarStatusNaGuiaDestino(guiaComparativo, 8, 9);
  var linhasAtualizadasListagem = aplicarStatusNaGuiaDestino(guiaListagem, 7, 8);
  
  if (linhasAtualizadasComp > 0 || linhasAtualizadasListagem > 0) {
    var msg = 'Mapeados ' + processosLidos + ' processos válidos na guia Itens.\n\n';
    msg += 'Aplicados na guia Comparativo: ' + linhasAtualizadasComp + ' linhas.\n';
    msg += 'Aplicados na guia listagem.total: ' + linhasAtualizadasListagem + ' linhas.';
    SpreadsheetApp.getUi().alert('Sucesso!', msg, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}