// FUNÇÃO: Roda o comparativo de Entradas x Itens em lote
function atualizarStatusEntradasLote() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaEntradas = planilha.getSheetByName('Entradas');
  var guiaItens = planilha.getSheetByName('Itens');
  var ui = SpreadsheetApp.getUi();

  if (!guiaEntradas || !guiaItens) return;
  var dadosEntradas = guiaEntradas.getDataRange().getValues();
  var somaPorItem = {};

  for (var i = 1; i < dadosEntradas.length; i++) {
    var itemEntradaBruto = dadosEntradas[i][0];
    if (itemEntradaBruto) {
      var itemKey = String(itemEntradaBruto).toLowerCase().trim();
      var qtd = parseBRNumber(dadosEntradas[i][3]);
      if (qtd > 0) somaPorItem[itemKey] = (somaPorItem[itemKey] || 0) + qtd;
    }
  }

  var dadosItens = guiaItens.getDataRange().getValues();
  var itensAtualizados = 0;

  for (var j = 2; j < dadosItens.length; j++) { // Inicia no índice 2 (Linha 3)
    var itemItensBruto = dadosItens[j][0];
    if (itemItensBruto) {
      var itemItensKey = String(itemItensBruto).toLowerCase().trim();
      var somaEntradas = somaPorItem[itemItensKey] || 0;
      var cmm = parseBRNumber(dadosItens[j][3]); 

      var celulaDestino = guiaItens.getRange(j + 1, 9);
      var valorAtual = String(celulaDestino.getValue());
      
      if (cmm > 0) {
        if (somaEntradas >= cmm) {
          celulaDestino.setValue('Sim Órgão').setBackground('#d9ead3');
          itensAtualizados++;
        } else if (somaEntradas > 0 && somaEntradas < cmm) {
          var dias = Math.round((somaEntradas / cmm) * 30);
          celulaDestino.setValue('Parcial ' + dias + ' d').setBackground('#fff2cc');
          itensAtualizados++;
        } else {
          if (valorAtual === 'Sim Órgão' || (valorAtual.indexOf('Parcial') !== -1 && valorAtual !== 'Sim')) {
            celulaDestino.setValue('Não').setBackground(null);
          } else if (valorAtual !== 'Sim' && valorAtual !== 'Sim Empenho' && valorAtual !== 'Não') {
            celulaDestino.setValue('Não').setBackground(null);
          }
        }
      }
    }
  }
  ui.alert('Comparativo Concluído', 'O status da coluna I (Guia Itens) foi atualizado.', ui.ButtonSet.OK);
}