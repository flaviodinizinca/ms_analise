// 3. CONSTRUTOR DE DASHBOARDS AUTOMÁTICOS (Visão MS)
function gerarPainelPresidencial() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var guiaHistorico = ss.getSheetByName('Base_Historico');
  var guiaResolvidos = ss.getSheetByName('ItensResolvidos');
  var ui = SpreadsheetApp.getUi();

  if (!guiaHistorico) {
    ui.alert('Atenção', 'A guia Base_Historico ainda não existe. Rode a gravação diária primeiro.', ui.ButtonSet.OK);
    return;
  }

  var dadosHist = guiaHistorico.getDataRange().getValues();
  if (dadosHist.length <= 1) {
    ui.alert('Atenção', 'Não há dados suficientes no histórico para gerar evolução.', ui.ButtonSet.OK);
    return;
  }

  // 1. Descobrir a Data Inicial e a Data Atual
  var dataInicial = dadosHist[1][0];
  var dataAtual = dadosHist[dadosHist.length - 1][0];
  
  function formatarData(data) {
    if (!data || !(data instanceof Date)) return String(data);
    var dia = ("0" + data.getDate()).slice(-2);
    var mes = ("0" + (data.getMonth() + 1)).slice(-2);
    return dia + "/" + mes;
  }
  
  var strDataInicial = formatarData(dataInicial);
  var strDataAtual = formatarData(dataAtual);

  // Status alvo para o Gráfico Analítico em Linhas
  var statusAlvoAnalitico = [
    "DISUP - Adesão - Pesquisa",
    "DISUP - Assinatura da ATA de SRP",
    "SEAL - Licitação marcada",
    "SEAL - Pregão em andamento"
  ];

  // 2. Processar os Dados (Sintético, Analítico e Temporal)
  var mapaSintetico = {};
  var mapaAnalitico = {};
  var datasUnicas = [];
  var setoresUnicos = [];
  var mapaEvolucaoTemporal = {}; 
  var mapaEvolucaoAnalitico = {}; 

  for (var i = 1; i < dadosHist.length; i++) {
    var dataLinhaStr = formatarData(dadosHist[i][0]);
    var statusCompleto = String(dadosHist[i][1]).trim();
    var qtd = parseBRNumber(dadosHist[i][2]);
    
    // Extrai o Setor (tudo antes do primeiro hífen)
    var setor = statusCompleto;
    if (statusCompleto.indexOf('-') !== -1) {
      setor = statusCompleto.split('-')[0].trim();
    }
    
    // --- Lógica das Visões Sintética e Analítica (Tabelas) ---
    if (!mapaSintetico[setor]) mapaSintetico[setor] = { inicio: 0, atual: 0 };
    if (!mapaAnalitico[statusCompleto]) mapaAnalitico[statusCompleto] = { inicio: 0, atual: 0 };
    
    if (formatarData(dadosHist[i][0]) === strDataInicial) {
      mapaSintetico[setor].inicio += qtd;
      mapaAnalitico[statusCompleto].inicio += qtd;
    }
    if (formatarData(dadosHist[i][0]) === strDataAtual) {
      mapaSintetico[setor].atual += qtd;
      mapaAnalitico[statusCompleto].atual += qtd;
    }

    // --- Lógica da Visão Temporal (Dia a Dia Macro) ---
    if (datasUnicas.indexOf(dataLinhaStr) === -1) datasUnicas.push(dataLinhaStr);
    if (setoresUnicos.indexOf(setor) === -1) setoresUnicos.push(setor);

    if (!mapaEvolucaoTemporal[dataLinhaStr]) mapaEvolucaoTemporal[dataLinhaStr] = {};
    if (!mapaEvolucaoTemporal[dataLinhaStr][setor]) mapaEvolucaoTemporal[dataLinhaStr][setor] = 0;
    
    mapaEvolucaoTemporal[dataLinhaStr][setor] += qtd;

    // --- Lógica da Visão Temporal Analítica (Específica) ---
    if (!mapaEvolucaoAnalitico[dataLinhaStr]) mapaEvolucaoAnalitico[dataLinhaStr] = {};
    if (statusAlvoAnalitico.indexOf(statusCompleto) !== -1) {
      if (mapaEvolucaoAnalitico[dataLinhaStr][statusCompleto] === undefined) {
        mapaEvolucaoAnalitico[dataLinhaStr][statusCompleto] = 0;
      }
      mapaEvolucaoAnalitico[dataLinhaStr][statusCompleto] += qtd;
    }
  }

  // Monta a Matriz de Dados para o Gráfico de Linhas Macro (Sintético)
  var matrizLinhas = [];
  var cabecalhoLinhas = ["Data"].concat(setoresUnicos);
  matrizLinhas.push(cabecalhoLinhas);

  for (var d = 0; d < datasUnicas.length; d++) {
    var linhaArr = [datasUnicas[d]];
    for (var s = 0; s < setoresUnicos.length; s++) {
      linhaArr.push(mapaEvolucaoTemporal[datasUnicas[d]][setoresUnicos[s]] || 0);
    }
    matrizLinhas.push(linhaArr);
  }

  // Monta a Matriz de Dados para o Gráfico de Linhas Analítico (Micro)
  var matrizLinhasAnalitico = [];
  var cabecalhoAnalitico = ["Data"].concat(statusAlvoAnalitico);
  matrizLinhasAnalitico.push(cabecalhoAnalitico);

  for (var d2 = 0; d2 < datasUnicas.length; d2++) {
    var linhaArrAnalitico = [datasUnicas[d2]];
    for (var s2 = 0; s2 < statusAlvoAnalitico.length; s2++) {
      linhaArrAnalitico.push(mapaEvolucaoAnalitico[datasUnicas[d2]][statusAlvoAnalitico[s2]] || 0);
    }
    matrizLinhasAnalitico.push(linhaArrAnalitico);
  }

  // 3. Contar Itens Resolvidos
  var totalResolvidos = 0;
  if (guiaResolvidos) {
    totalResolvidos = guiaResolvidos.getLastRow() - 1;
    if (totalResolvidos < 0) totalResolvidos = 0;
  }

  // 4. Construtor das guias de Colunas e Linhas (Sintética e Analítica)
  function construirGuiaVisao(nomeGuia, mapaDados, titulo) {
    var guia = ss.getSheetByName(nomeGuia);
    if (!guia) { guia = ss.insertSheet(nomeGuia); } 
    else { 
      guia.clear();
      var charts = guia.getCharts();
      for (var c = 0; c < charts.length; c++) { guia.removeChart(charts[c]); }
    }
    guia.setHiddenGridlines(true);

    guia.getRange("B2").setValue("🏆 ITENS RESOLVIDOS (Regularizados):").setFontWeight("bold").setFontSize(14).setFontColor("#2c3e50");
    guia.getRange("C2").setValue(totalResolvidos).setFontWeight("bold").setFontSize(16).setFontColor("#27ae60").setHorizontalAlignment("left");
    guia.getRange("B4").setValue(titulo).setFontWeight("bold").setFontSize(14).setFontColor("#2c3e50");

    var linhaAtual = 6;
    var cabecalhoTabela = ["Departamento / Status", "Início (" + strDataInicial + ")", "Atualmente (" + strDataAtual + ")", "Evolução"];
    guia.getRange(linhaAtual, 2, 1, 4).setValues([cabecalhoTabela]).setFontWeight("bold").setBackground("#4CAF50").setFontColor("white");
    linhaAtual++;

    var dadosTabela = [];
    for (var chave in mapaDados) {
      var inicio = mapaDados[chave].inicio;
      var atual = mapaDados[chave].atual;
      var evolucao = atual - inicio; 
      var textoEvolucao = evolucao > 0 ? "+" + evolucao + " (Aumentou)" : (evolucao < 0 ? evolucao + " (Reduziu)" : "Estável");
      
      if (inicio > 0 || atual > 0) dadosTabela.push([chave, inicio, atual, textoEvolucao]);
    }
    dadosTabela.sort(function(a, b) { return a[0].localeCompare(b[0]); });
    
    if (dadosTabela.length > 0) {
      guia.getRange(linhaAtual, 2, dadosTabela.length, 4).setValues(dadosTabela);
      guia.getRange(linhaAtual, 2, dadosTabela.length, 4).setBorder(true, true, true, true, true, true, '#cccccc', SpreadsheetApp.BorderStyle.SOLID);
      
      if (nomeGuia === "Visão_Analítica_MS") {
        var rangeDadosGrafico = guia.getRange(50, 7, matrizLinhasAnalitico.length, matrizLinhasAnalitico[0].length);
        rangeDadosGrafico.setValues(matrizLinhasAnalitico);

        var graficoLinhaAnalitico = guia.newChart()
          .setChartType(Charts.ChartType.LINE)
          .addRange(rangeDadosGrafico)
          .setPosition(5, 7, 0, 0)
          .setOption('title', 'Evolução Temporal dos Status Selecionados')
          .setOption('width', 900)
          .setOption('height', 450)
          .setOption('legend', {position: 'right', textStyle: {color: '#2c3e50', fontSize: 12}})
          .setOption('pointSize', 7)
          .setOption('curveType', 'function')
          .build();
        guia.insertChart(graficoLinhaAnalitico);
      } else {
        var rangeGrafico = guia.getRange(5, 2, dadosTabela.length + 1, 3);
        var graficoColuna = guia.newChart()
          .setChartType(Charts.ChartType.COLUMN)
          .addRange(rangeGrafico)
          .setPosition(5, 7, 0, 0)
          .setOption('title', 'Comparativo de Pendências: Início vs Atual')
          .setOption('width', 700)
          .setOption('height', 400)
          .setOption('legend', {position: 'bottom'})
          .setOption('colors', ['#95a5a6', '#e74c3c'])
          .setOption('series', { 0: { dataLabel: 'value' }, 1: { dataLabel: 'value' } })
          .build();
        guia.insertChart(graficoColuna);
      }
    }
    guia.autoResizeColumn(2); guia.autoResizeColumn(3); guia.autoResizeColumn(4); guia.autoResizeColumn(5);
  }

  function construirGuiaTemporal() {
    var nomeGuia = "Visão_Temporal_MS";
    var guia = ss.getSheetByName(nomeGuia);
    if (!guia) { guia = ss.insertSheet(nomeGuia); } 
    else { 
      guia.clear(); 
      var charts = guia.getCharts();
      for (var c = 0; c < charts.length; c++) { guia.removeChart(charts[c]); }
    }
    guia.setHiddenGridlines(true);
    
    guia.getRange("B2").setValue("📈 HISTÓRICO DIA A DIA (Curva de Processos por Setor)").setFontWeight("bold").setFontSize(14).setFontColor("#2c3e50");
    
    var rangeDados = guia.getRange(50, 2, matrizLinhas.length, matrizLinhas[0].length);
    rangeDados.setValues(matrizLinhas); 
    
    var graficoLinhas = guia.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(rangeDados)
        .setPosition(4, 2, 0, 0)
        .setOption('title', 'Evolução de Pendências (Macro Setores)')
        .setOption('width', 900)
        .setOption('height', 450)
        .setOption('legend', {position: 'right', textStyle: {color: '#2c3e50', fontSize: 12}}) 
        .setOption('pointSize', 7) 
        .setOption('curveType', 'function') 
        .build();
        
    guia.insertChart(graficoLinhas);
  }

  construirGuiaVisao("Visão_Sintética_MS", mapaSintetico, "📈 EVOLUÇÃO MACRO: Por Setor (Resumo Executivo)");
  construirGuiaVisao("Visão_Analítica_MS", mapaAnalitico, "🔍 EVOLUÇÃO MICRO: Por Status Detalhado");
  construirGuiaTemporal();

  ui.alert("Painel Presidencial Completo!", "As 3 guias foram geradas. Confira a legenda e os números agora!", ui.ButtonSet.OK);
}