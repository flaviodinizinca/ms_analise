// MOTOR CENTRAL DE CLASSIFICAÇÃO DE RISCO (Universal)
function obterClassificacaoRisco(saldo, cmm) {
  var txtRisco0 = "Risco Crítico (Risco 0)\nItens com estoque zerado ou com cobertura inferior a um mês, com potencial impacto imediato na assistência.";
  var txtRisco1 = "Risco Alto (Risco 1)\nItens com cobertura entre um e três meses e processo de aquisição ainda em andamento.";
  var txtRisco2 = "Risco Moderado (Risco 2)\nItens com cobertura entre três e seis meses, com processo de aquisição em fase avançada.";
  var txtRiscoBaixo = "Estoque Confortável\nCobertura superior a 6 meses.";
  var txtSemCmm = "Sem Histórico\nCMM zerado, impossível calcular cobertura.";
  
  if (saldo <= 0) {
    return { texto: txtRisco0, cor: '#ea9999' }; // Vermelho
  } else if (cmm > 0) {
    var coberturaMeses = saldo / cmm;
    if (coberturaMeses < 1) {
      return { texto: txtRisco0, cor: '#ea9999' }; // Vermelho
    } else if (coberturaMeses >= 1 && coberturaMeses < 3) {
      return { texto: txtRisco1, cor: '#fce5cd' }; // Laranja claro
    } else if (coberturaMeses >= 3 && coberturaMeses <= 6) {
      return { texto: txtRisco2, cor: '#fff2cc' }; // Amarelo
    } else {
      return { texto: txtRiscoBaixo, cor: '#d9ead3' }; // Verde claro
    }
  } else {
    return { texto: txtSemCmm, cor: '#f3f3f3' }; // Cinza
  }
}