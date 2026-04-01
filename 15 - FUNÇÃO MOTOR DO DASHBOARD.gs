// 2. FUNÇÃO MOTOR DO DASHBOARD: COLETA HISTÓRICO DINÂMICO
function gravarHistoricoDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var guiaItens = ss.getSheetByName('Itens');
  var guiaHistorico = ss.getSheetByName('Base_Historico');
  var ui = SpreadsheetApp.getUi();

  if (!guiaItens) return;

  if (!guiaHistorico) {
    guiaHistorico = ss.insertSheet('Base_Historico');
    guiaHistorico.appendRow(['Data', 'Situação do Processo', 'Quantidade de Itens']);
    guiaHistorico.getRange("A1:C1").setFontWeight("bold").setBackground("#4CAF50").setFontColor("white");
  }

  var dados = guiaItens.getDataRange().getValues();
  var contagemStatus = {};
  
  for (var i = 2; i < dados.length; i++) { // Inicia no índice 2 (Linha 3)
    var status = String(dados[i][9]).trim();
    if (status !== "") {
      if (!contagemStatus[status]) {
        contagemStatus[status] = 0;
      }
      contagemStatus[status]++;
    }
  }

  var hoje = new Date();
  hoje.setHours(0,0,0,0);
  var novasLinhas = [];
  for (var chaveStatus in contagemStatus) {
    novasLinhas.push([hoje, chaveStatus, contagemStatus[chaveStatus]]);
  }

  if (novasLinhas.length > 0) {
    var ultimaLinha = guiaHistorico.getLastRow();
    guiaHistorico.getRange(ultimaLinha + 1, 1, novasLinhas.length, 3).setValues(novasLinhas);
    guiaHistorico.getRange(ultimaLinha + 1, 1, novasLinhas.length, 1).setNumberFormat('dd/MM/yyyy');
    
    ui.alert("Base Atualizada", "O resumo de hoje foi gravado com sucesso na guia 'Base_Historico' e já pode alimentar seu Dashboard.", ui.ButtonSet.OK);
  }
}