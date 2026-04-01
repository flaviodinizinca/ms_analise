// FUNÇÃO ONEDIT
function onEdit(e) {
  if (!e || !e.range) return;
  var guiaEditada = e.range.getSheet();
  if (guiaEditada.getName() !== 'Entradas') return;
  var linhaEditada = e.range.getRow();
  var colunaEditada = e.range.getColumn();
  
  if (linhaEditada > 1 && (colunaEditada === 1 || colunaEditada === 4)) {
    var planilha = e.source;
    var guiaEntradas = planilha.getSheetByName('Entradas');
    var guiaItens = planilha.getSheetByName('Itens');
    
    var itemEditadoBruto = guiaEntradas.getRange(linhaEditada, 1).getValue();
    if (!itemEditadoBruto) return;
    var itemEditado = String(itemEditadoBruto).toLowerCase().trim();
    var dadosEntradas = guiaEntradas.getDataRange().getValues();
    var somaQuantidades = 0;
    
    for (var i = 1; i < dadosEntradas.length; i++) {
      var itemLinhaAtual = String(dadosEntradas[i][0]).toLowerCase().trim();
      if (itemLinhaAtual === itemEditado) {
        var qtd = parseBRNumber(dadosEntradas[i][3]);
        if (qtd > 0) somaQuantidades += qtd;
      }
    }
    
    var dadosItens = guiaItens.getDataRange().getValues();
    for (var j = 2; j < dadosItens.length; j++) { // Inicia no índice 2 (Linha 3)
      var itemGuiaItens = String(dadosItens[j][0]).toLowerCase().trim();
      if (itemGuiaItens === itemEditado) {
        var valorAlvo = parseBRNumber(dadosItens[j][3]);
        var celulaDestino = guiaItens.getRange(j + 1, 9); // Coluna I
        var valorAtual = String(celulaDestino.getValue());
        
        if (valorAlvo > 0) {
          if (somaQuantidades >= valorAlvo) {
            celulaDestino.setValue('Sim Órgão').setBackground('#d9ead3');
          } else if (somaQuantidades > 0 && somaQuantidades < valorAlvo) {
            var dias = Math.round((somaQuantidades / valorAlvo) * 30);
            celulaDestino.setValue('Parcial ' + dias + ' d').setBackground('#fff2cc'); 
          } else {
            if (valorAtual !== 'Sim' && valorAtual !== 'Sim Empenho' && valorAtual !== 'Não') {
              celulaDestino.setValue('Não').setBackground(null);
            }
          }
        }
        break;
      }
    }
  }
}