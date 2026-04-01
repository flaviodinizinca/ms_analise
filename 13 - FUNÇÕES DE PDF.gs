// FUNÇÕES DE PDF
function gerarRelatorioRegularizadosPDF() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaItens = planilha.getSheetByName('Itens');
  var dadosItens = guiaItens.getDataRange().getValues();
  
  var html = "<!DOCTYPE html><html><head><style>body{font-family:Arial,sans-serif;font-size:12px;color:#333}h2{text-align:center;color:#2c3e50;margin-bottom:20px}table{width:100%;border-collapse:collapse;margin-bottom:20px}th,td{border:1px solid #555;padding:8px;text-align:left}thead{display:table-header-group}tr{page-break-inside:avoid}th{background-color:#2980b9;color:#fff;font-weight:bold}tr:nth-child(even){background-color:#f2f2f2}tr:nth-child(odd){background-color:#fff}</style></head><body>";
  html += "<h2>Relatório de Itens Regularizados</h2><table><thead><tr><th>Item</th><th>Descrição</th><th>Saldo Atual</th><th>Status do Processo</th></tr></thead><tbody>";
  
  var temItens = false;
  for (var k = 2; k < dadosItens.length; k++) { // Inicia no índice 2 (Linha 3)
    var linha = dadosItens[k];
    var itemOriginal = String(linha[0]).trim();
    if (itemOriginal) { 
      var saldo = parseBRNumber(linha[2]);
      var cmm = parseBRNumber(linha[3]);
      if (saldo >= cmm && cmm > 0) {
        temItens = true;
        var desc = linha[1] ? linha[1] : "-"; 
        var statusProc = linha[9] ? linha[9] : "-";
        html += "<tr><td>" + itemOriginal + "</td><td>" + desc + "</td><td>" + saldo.toString().replace('.',',') + "</td><td>" + statusProc + "</td></tr>";
      }
    }
  }
  if (!temItens) html += "<tr><td colspan='4' style='text-align: center;'>Nenhum item regularizado.</td></tr>";
  html += "</tbody></table></body></html>";
  
  var nomeArq = "Relatorio_Itens_Regularizados.pdf";
  DriveApp.getRootFolder().createFile(Utilities.newBlob(html, MimeType.HTML).setName(nomeArq).getAs(MimeType.PDF));
  SpreadsheetApp.getUi().alert('Relatório Gerado!', 'O PDF foi salvo no Meu Drive.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function gerarRelatorioPDF() {
  var planilha = SpreadsheetApp.getActiveSpreadsheet();
  var guiaItens = planilha.getSheetByName('Itens');
  var guiaEntradas = planilha.getSheetByName('Entradas');
  var dadosItens = guiaItens.getDataRange().getValues();
  var dadosEntradas = guiaEntradas.getDataRange().getValues();
  
  var orgaosPorItem = {}, analiticoOrgao = {}, descItens = {};
  for (var j = 2; j < dadosItens.length; j++) { // Inicia no índice 2 (Linha 3)
    var itemBruto = String(dadosItens[j][0]).trim();
    if (itemBruto) descItens[itemBruto.toLowerCase()] = dadosItens[j][1] ? dadosItens[j][1] : "-";
  }
  
  for (var i = 1; i < dadosEntradas.length; i++) {
    var itemOriginal = dadosEntradas[i][0];
    var orgaoOriginal = dadosEntradas[i][2]; 
    var qtd = parseBRNumber(dadosEntradas[i][3]);      
    if (!itemOriginal || !orgaoOriginal) continue;
    var itemKey = String(itemOriginal).trim().toLowerCase();
    var orgao = String(orgaoOriginal).trim();
    
    if (!orgaosPorItem[itemKey]) orgaosPorItem[itemKey] = [];
    if (orgaosPorItem[itemKey].indexOf(orgao) === -1) orgaosPorItem[itemKey].push(orgao);
    if (!analiticoOrgao[orgao]) analiticoOrgao[orgao] = {};
    if (!analiticoOrgao[orgao][itemKey]) analiticoOrgao[orgao][itemKey] = { display: String(itemOriginal).trim(), qtd: 0 };
    analiticoOrgao[orgao][itemKey].qtd += qtd;
  }
  
  var html = "<!DOCTYPE html><html><head><style>body{font-family:Arial,sans-serif;font-size:12px;color:#333}h2{text-align:center;color:#2c3e50;margin-bottom:20px}h3{color:#2c3e50;margin-top:30px;border-bottom:2px solid #4CAF50;padding-bottom:5px}table{width:100%;border-collapse:collapse;margin-bottom:20px}th,td{border:1px solid #555;padding:8px;text-align:left}thead{display:table-header-group}tr{page-break-inside:avoid}th{background-color:#4CAF50;color:#fff;font-weight:bold}tr:nth-child(even){background-color:#f2f2f2}tr:nth-child(odd){background-color:#fff}.quebra{page-break-before:always}</style></head><body>";
  html += "<h2>Relatório Sintético: Visão Geral</h2><table><thead><tr><th>Item</th><th>Descrição</th><th>CMM</th><th>Órgãos Origem</th><th>Processo SEI</th><th>Status Processo</th><th>Item Coberto</th></tr></thead><tbody>";
  
  for (var k = 2; k < dadosItens.length; k++) { // Inicia no índice 2 (Linha 3)
    var l = dadosItens[k];
    var item = String(l[0]).trim();
    if (item) { 
      var cmm = l[3] !== "" ? l[3] : "0";      
      var processo = l[7] ? l[7] : "-"; 
      var statusProc = l[9] ? l[9] : "-";
      var status = l[8] ? l[8] : "Não"; 
      var orgs = orgaosPorItem[item.toLowerCase()] ? orgaosPorItem[item.toLowerCase()].join(", ") : "-";
      html += "<tr><td>"+item+"</td><td>"+(l[1]?l[1]:"-")+"</td><td>"+cmm+"</td><td>"+orgs+"</td><td>"+processo+"</td><td>"+statusProc+"</td><td>"+status+"</td></tr>";
    }
  }
  html += "</tbody></table><div class='quebra'></div><h2>Relatório Analítico: Entradas por Órgão</h2>";
  
  var nomesOrgaos = Object.keys(analiticoOrgao).sort();
  if (nomesOrgaos.length === 0) { html += "<p>Nenhuma entrada.</p>"; } 
  else {
    for (var o = 0; o < nomesOrgaos.length; o++) {
      var orgaoAtual = nomesOrgaos[o];
      html += "<h3>Órgão: " + orgaoAtual + "</h3><table><thead><tr><th>Item</th><th>Descrição</th><th>Total Recebido</th></tr></thead><tbody>";
      var itensOrg = analiticoOrgao[orgaoAtual];
      for (var ch in itensOrg) {
        var desc = descItens[ch] ? descItens[ch] : "-";
        html += "<tr><td>"+itensOrg[ch].display+"</td><td>"+desc+"</td><td>"+itensOrg[ch].qtd.toString().replace('.',',')+"</td></tr>";
      }
      html += "</tbody></table>";
    }
  }
  html += "</body></html>";
  
  var nArq = "Relatorio_Estoque_Completo.pdf";
  DriveApp.getRootFolder().createFile(Utilities.newBlob(html, MimeType.HTML).setName(nArq).getAs(MimeType.PDF));
  SpreadsheetApp.getUi().alert('Relatório Gerado!', 'O PDF foi salvo no Meu Drive.', SpreadsheetApp.getUi().ButtonSet.OK);
}