// NOVA FUNÇÃO: Atualiza o texto dinâmico da Linha 1 na Guia Itens
function atualizarResumoCabecalho() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var guiaItens = ss.getSheetByName('Itens');
  var guiaResolvidos = ss.getSheetByName('ItensResolvidos');

  if (!guiaItens) return;

  var ultimaLinhaItens = guiaItens.getLastRow();
  // Pendentes são as linhas de dados na guia Itens (excluindo linha 1 do resumo e linha 2 do cabeçalho)
  var qtdPendentes = ultimaLinhaItens >= 3 ? (ultimaLinhaItens - 2) : 0; 
  
  var qtdResolvidos = 0;
  if (guiaResolvidos) {
    var ultimaLinhaResolvidos = guiaResolvidos.getLastRow();
    // Exclui a linha 1 de cabeçalho da guia de Resolvidos
    qtdResolvidos = ultimaLinhaResolvidos >= 2 ? (ultimaLinhaResolvidos - 1) : 0;
  }
  
  var qtdTotal = qtdPendentes + qtdResolvidos;
  
  var textoResumo = "📊 RESUMO GERAL: De " + qtdTotal + " itens acompanhados, " + qtdResolvidos + " já foram resolvidos (regularizados) e " + qtdPendentes + " ainda constam como pendentes.";
  
  var rangeResumo = guiaItens.getRange("A1");
  rangeResumo.setValue(textoResumo)
             .setFontWeight("bold")
             .setFontSize(14)
             .setFontColor("#1e8e3e") 
             .setBackground("#e6f4ea"); 
}