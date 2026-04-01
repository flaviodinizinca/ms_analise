// APLICA O RISCO EM MASSA PARA QUALQUER GUIA (Atualizado com controle de linha inicial)
function aplicarRiscoNaGuia(guia, linhaInicioData) {
  if (!guia) return 0;
  var dados = guia.getDataRange().getValues();
  var linhasAtualizadas = 0;
  
  // Se linhaInicioData for informado, usa ele, senão usa 1 (linha 2)
  var inicio = (linhaInicioData !== undefined) ? linhaInicioData : 1; 
  
  for (var i = inicio; i < dados.length; i++) {
    var itemBruto = dados[i][0];
    if (itemBruto) {
      var saldo = parseBRNumber(dados[i][2]); // Coluna C
      var cmm = parseBRNumber(dados[i][3]);   // Coluna D
      var riscoCalc = obterClassificacaoRisco(saldo, cmm);
      
      guia.getRange(i + 1, 7).setValue(riscoCalc.texto).setBackground(riscoCalc.cor); // Coluna G
      linhasAtualizadas++;
    }
  }
  return linhasAtualizadas;
}