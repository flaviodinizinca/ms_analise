// FUNÇÃO MANUAL: Classifica Itens e Comparativo
function classificarRiscoItens() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaItens = planilha.getSheetByName('Itens');
  var guiaComp = planilha.getSheetByName('Comparativo-Guia.Itens X listagem.total');
  var ui = SpreadsheetApp.getUi();

  // Guia Itens agora tem dados começando no índice 2 (Linha 3). Guia Comp no índice 1 (Linha 2).
  var atualizadasItens = aplicarRiscoNaGuia(guiaItens, 2);
  var atualizadasComp = aplicarRiscoNaGuia(guiaComp, 1);
  
  if (atualizadasItens > 0 || atualizadasComp > 0) {
    ui.alert('Classificação Concluída', 'Risco re-calculado:\n- Guia Itens: ' + atualizadasItens + ' linhas.\n- Guia Comparativo: ' + atualizadasComp + ' linhas.', ui.ButtonSet.OK);
  }
}