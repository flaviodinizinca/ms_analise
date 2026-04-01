// FUNÇÃO: Comparar Itens Faltantes (Blindada contra Itens Resolvidos)
function compararItensFaltantes() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaItens = planilha.getSheetByName('Itens');
  var guiaResolvidos = planilha.getSheetByName('ItensResolvidos'); // NOVO: Lê a guia de resolvidos
  var guiaListagem = planilha.getSheetByName('listagem.total');
  var nomeGuiaResultado = 'Comparativo-Guia.Itens X listagem.total';
  var guiaResultado = planilha.getSheetByName(nomeGuiaResultado);
  var ui = SpreadsheetApp.getUi();
  
  if (!guiaItens || !guiaListagem || !guiaResultado) return;
  
  var conjuntoItens = new Set();
  
  // 1. Mapeia os itens que já estão na guia principal (Itens)
  var dadosItens = guiaItens.getRange('A:A').getValues();
  for (var i = 2; i < dadosItens.length; i++) { // Inicia no índice 2 (Linha 3)
    var itemBruto = String(dadosItens[i][0]).trim();
    if (itemBruto !== '') conjuntoItens.add(itemBruto.toLowerCase());
  }

  // 2. Mapeia os itens que já foram resolvidos e arquivados (Para não pedirem para voltar)
  if (guiaResolvidos) {
    var dadosResolvidos = guiaResolvidos.getRange('A:A').getValues();
    for (var r = 1; r < dadosResolvidos.length; r++) { 
      var itemRes = String(dadosResolvidos[r][0]).trim();
      if (itemRes !== '') conjuntoItens.add(itemRes.toLowerCase());
    }
  }
  
  // 3. Verifica a Listagem Total contra o que já mapeamos
  var ultimaLinhaListagem = guiaListagem.getLastRow();
  if (ultimaLinhaListagem === 0) return;
  var dadosListagem = guiaListagem.getRange(1, 1, ultimaLinhaListagem, 8).getValues();
  
  var itensFaltantes = [];
  if (dadosListagem.length > 0) {
    var cabecalho = dadosListagem[0];
    itensFaltantes.push([cabecalho[0], cabecalho[1], cabecalho[2], cabecalho[3], cabecalho[4], cabecalho[5], "Risco", cabecalho[6], cabecalho[7]]);
  }
  
  var contagem = 0;
  for (var j = 1; j < dadosListagem.length; j++) { 
    var itemListaBruto = String(dadosListagem[j][0]).trim();
    if (itemListaBruto !== '') {
      // Se não está nem na guia Itens e nem na Resolvidos, aí sim é Faltante
      if (!conjuntoItens.has(itemListaBruto.toLowerCase())) {
        var rowListagem = dadosListagem[j];
        itensFaltantes.push([
          rowListagem[0], rowListagem[1], rowListagem[2], rowListagem[3], 
          rowListagem[4], rowListagem[5], "", rowListagem[6], rowListagem[7]
        ]);
        contagem++;
      }
    }
  }
  
  guiaResultado.clear();
  if (contagem === 0) {
    guiaResultado.getRange(1, 1).setValue('Todos os itens da listagem.total já estão presentes ou resolvidos.');
    ui.alert('Aviso', 'Nenhum item faltante.', ui.ButtonSet.OK); return;
  }
  
  guiaResultado.getRange(1, 1, itensFaltantes.length, 9).setValues(itensFaltantes);
  for (var c = 1; c <= 9; c++) { guiaResultado.autoResizeColumn(c); }
  
  aplicarRiscoNaGuia(guiaResultado, 1);
  ui.alert('Concluído', 'Foram encontrados ' + contagem + ' itens faltantes e os riscos foram pintados.', ui.ButtonSet.OK);
}