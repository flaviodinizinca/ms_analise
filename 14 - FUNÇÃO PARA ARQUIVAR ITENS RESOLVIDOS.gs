// 1. FUNÇÃO PARA ARQUIVAR ITENS RESOLVIDOS (Saldo >= CMM)
function arquivarItensResolvidos() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var guiaItens = ss.getSheetByName('Itens');
  var guiaResolvidos = ss.getSheetByName('ItensResolvidos');
  var ui = SpreadsheetApp.getUi();
  
  if (!guiaItens) {
    ui.alert('Erro', 'A guia "Itens" não foi encontrada.', ui.ButtonSet.OK);
    return;
  }
  
  if (!guiaResolvidos) {
    guiaResolvidos = ss.insertSheet('ItensResolvidos');
    var cabecalho = guiaItens.getRange(2, 1, 1, guiaItens.getLastColumn()).getValues(); // Copia cabeçalho da linha 2
    guiaResolvidos.getRange(1, 1, 1, cabecalho[0].length).setValues(cabecalho);
  }

  var dados = guiaItens.getDataRange().getValues();
  var linhasParaExcluir = [];
  var itensParaMover = [];
  
  for (var i = dados.length - 1; i >= 2; i--) { // Inicia a verificação pelo final, parando no índice 2 (Linha 3)
    var saldo = parseBRNumber(dados[i][2]);
    var cmm = parseBRNumber(dados[i][3]);   

    if (cmm > 0 && saldo >= cmm) {
      itensParaMover.push(dados[i]);
      linhasParaExcluir.push(i + 1); 
    }
  }

  if (itensParaMover.length > 0) {
    for (var j = itensParaMover.length - 1; j >= 0; j--) {
      guiaResolvidos.appendRow(itensParaMover[j]);
    }
    
    for (var k = 0; k < linhasParaExcluir.length; k++) {
      guiaItens.deleteRow(linhasParaExcluir[k]);
    }
    
    atualizarResumoCabecalho(); // Atualiza a frase dinâmica do resumo automaticamente
    ui.alert("Meta Alcançada!", itensParaMover.length + " itens saíram do risco, superaram o CMM e foram arquivados na aba ItensResolvidos.", ui.ButtonSet.OK);
  } else {
    atualizarResumoCabecalho(); // Atualiza a contagem mesmo se não houver novos envios
    ui.alert("Aviso", "Nenhum item atingiu a meta (Saldo >= CMM) para ser arquivado nesta execução.", ui.ButtonSet.OK);
  }
}