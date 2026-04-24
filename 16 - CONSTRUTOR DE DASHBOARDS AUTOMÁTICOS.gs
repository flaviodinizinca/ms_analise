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

  // Filtrar dados para considerar apenas a partir de 27 de Março de 2026
  var dadosHistFiltrados = [];
  var dataCorte = new Date(2026, 2, 27); 
  dataCorte.setHours(0,0,0,0);

  for (var i = 1; i < dadosHist.length; i++) {
    var dataLinha = new Date(dadosHist[i][0]);
    if (!(dataLinha instanceof Date) || isNaN(dataLinha)) continue;
    
    var dataComparacao = new Date(dataLinha.getTime());
    dataComparacao.setHours(0,0,0,0);
    if (dataComparacao >= dataCorte) {
      dadosHistFiltrados.push(dadosHist[i]);
    }
  }

  if (dadosHistFiltrados.length === 0) {
    ui.alert('Atenção', 'Não há dados no histórico a partir de 27/03/2026 para gerar o painel.', ui.ButtonSet.OK);
    return;
  }

  // 1. Descobrir a Data Inicial e a Data Atual
  var dataInicial = dadosHistFiltrados[0][0];
  var dataAtual = dadosHistFiltrados[dadosHistFiltrados.length - 1][0];
  
  function formatarData(data) {
    if (!data || !(data instanceof Date)) return String(data);
    var dia = ("0" + data.getDate()).slice(-2);
    var mes = ("0" + (data.getMonth() + 1)).slice(-2);
    return dia + "/" + mes;
  }
  
  var strDataInicial = formatarData(dataInicial);
  var strDataAtual = formatarData(dataAtual);

  // ====================================================================================
  // PASSO 1: UNIFICAR MAIÚSCULAS E MINÚSCULAS (Mapeamento Canônico)
  // ====================================================================================
  var mapaCanonical = {};
  
  // Status alvo bruto para o Gráfico Analítico em Linhas (com a capitalização desejada)
  var statusAlvoAnaliticoBrutos = [
    "DISUP - Adesão - Pesquisa",
    "DISUP - Assinatura da ATA de SRP",
    "SEAL - Licitação marcada",
    "SEAL - Marcação de Licitação",
    "SEAL - Pregão em andamento",
    "SECOM - Pesquisa de preços",
    "SECOM - Pesquisa em conclusao"
  ];

  // Força a prioridade de exibição para os itens oficiais do gráfico
  for (var x = 0; x < statusAlvoAnaliticoBrutos.length; x++) {
      var upperAlvo = statusAlvoAnaliticoBrutos[x].toUpperCase();
      mapaCanonical[upperAlvo] = statusAlvoAnaliticoBrutos[x];
  }

  // Escaneia o histórico para achar todas as variações e priorizar as que têm as letras Maiúsculas corretas
  for (var y = 0; y < dadosHistFiltrados.length; y++) {
      var sttBruto = String(dadosHistFiltrados[y][1]).trim();
      if (sttBruto === "") continue;

      var sttUpper = sttBruto.toUpperCase();
      
      if (!mapaCanonical[sttUpper]) {
          mapaCanonical[sttUpper] = sttBruto;
      } else {
          // Se o novo registro tiver mais letras maiúsculas que o salvo, adota o novo como o "bonito/oficial"
          var maiusculasAtuais = (mapaCanonical[sttUpper].match(/[A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ]/g) || []).length;
          var maiusculasNovas = (sttBruto.match(/[A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ]/g) || []).length;
          if (maiusculasNovas > maiusculasAtuais) {
              mapaCanonical[sttUpper] = sttBruto;
          }
      }
      
      // Cadastra também a padronização do nome do departamento (Macro Setor)
      var sTrBruto = sttBruto;
      if (sttBruto.indexOf('-') !== -1) {
        sTrBruto = sttBruto.split('-')[0].trim();
      }
      var sTrUpper = sTrBruto.toUpperCase();
      if (!mapaCanonical[sTrUpper]) {
          mapaCanonical[sTrUpper] = sTrBruto;
      } else {
          var mAtuais = (mapaCanonical[sTrUpper].match(/[A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ]/g) || []).length;
          var mNovas = (sTrBruto.match(/[A-ZÁÀÂÃÉÈÊÍÏÓÔÕÖÚÇÑ]/g) || []).length;
          if (mNovas > mAtuais) {
              mapaCanonical[sTrUpper] = sTrBruto;
          }
      }
  }

  // Função que retorna sempre a versão unificada do status
  function getNomeCanonical(nome) {
      var strTratada = String(nome).trim();
      if (strTratada === "") return "";
      return mapaCanonical[strTratada.toUpperCase()] || strTratada;
  }

  // Oficializa a matriz de status que irão para o gráfico de linhas já com o nome unificado
  var statusAlvoAnalitico = [];
  for (var z = 0; z < statusAlvoAnaliticoBrutos.length; z++) {
      statusAlvoAnalitico.push(getNomeCanonical(statusAlvoAnaliticoBrutos[z]));
  }
  // ====================================================================================

  // 2. Processar os Dados (Sintético, Analítico e Temporal)
  var mapaSintetico = {};
  var mapaAnalitico = {};
  var datasUnicas = [];
  var setoresUnicos = [];
  var mapaEvolucaoTemporal = {}; 
  var mapaEvolucaoAnalitico = {};

  for (var j = 0; j < dadosHistFiltrados.length; j++) {
    var dataLinhaStr = formatarData(dadosHistFiltrados[j][0]);
    var statusSujo = String(dadosHistFiltrados[j][1]).trim();
    
    // Transforma a variação num único texto consolidado, ex: funde "licitação" e "Licitação"
    var statusCompleto = getNomeCanonical(statusSujo); 
    
    var qtd = parseBRNumber(dadosHistFiltrados[j][2]);
    
    var setorSujo = statusCompleto;
    if (statusCompleto.indexOf('-') !== -1) {
      setorSujo = statusCompleto.split('-')[0].trim();
    }
    var setor = getNomeCanonical(setorSujo);
    
    if (!mapaSintetico[setor]) mapaSintetico[setor] = { inicio: 0, atual: 0 };
    if (!mapaAnalitico[statusCompleto]) mapaAnalitico[statusCompleto] = { inicio: 0, atual: 0 };

    if (formatarData(dadosHistFiltrados[j][0]) === strDataInicial) {
      mapaSintetico[setor].inicio += qtd;
      mapaAnalitico[statusCompleto].inicio += qtd;
    }
    if (formatarData(dadosHistFiltrados[j][0]) === strDataAtual) {
      mapaSintetico[setor].atual += qtd;
      mapaAnalitico[statusCompleto].atual += qtd;
    }

    if (datasUnicas.indexOf(dataLinhaStr) === -1) datasUnicas.push(dataLinhaStr);
    if (setoresUnicos.indexOf(setor) === -1) setoresUnicos.push(setor);

    if (!mapaEvolucaoTemporal[dataLinhaStr]) mapaEvolucaoTemporal[dataLinhaStr] = {};
    if (!mapaEvolucaoTemporal[dataLinhaStr][setor]) mapaEvolucaoTemporal[dataLinhaStr][setor] = 0;
    
    mapaEvolucaoTemporal[dataLinhaStr][setor] += qtd;

    if (!mapaEvolucaoAnalitico[dataLinhaStr]) mapaEvolucaoAnalitico[dataLinhaStr] = {};
    if (statusAlvoAnalitico.indexOf(statusCompleto) !== -1) {
      if (mapaEvolucaoAnalitico[dataLinhaStr][statusCompleto] === undefined) {
        mapaEvolucaoAnalitico[dataLinhaStr][statusCompleto] = 0;
      }
      mapaEvolucaoAnalitico[dataLinhaStr][statusCompleto] += qtd;
    }
  }

  // Monta a Matriz de Dados
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

  var totalResolvidos = 0;
  if (guiaResolvidos) {
    totalResolvidos = guiaResolvidos.getLastRow() - 1;
    if (totalResolvidos < 0) totalResolvidos = 0;
  }

  // 4. Construtor das guias
  function construirGuiaVisao(nomeGuia, mapaDados, titulo) {
    var guia = ss.getSheetByName(nomeGuia);
    if (!guia) { 
        guia = ss.insertSheet(nomeGuia); 
    } else { 
      guia.clear();
      var charts = guia.getCharts();
      for (var c = 0; c < charts.length; c++) { 
          guia.removeChart(charts[c]);
      }
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
        
        // Configura rótulos de dados forçando a injeção do nome da legenda de volta
        var seriesLabelAnalitico = {};
        for (var sa = 0; sa < statusAlvoAnalitico.length; sa++) {
          seriesLabelAnalitico[sa] = { 
            dataLabel: 'value',
            labelInLegend: statusAlvoAnalitico[sa],
            label: statusAlvoAnalitico[sa]
          };
        }

        var graficoLinhaAnalitico = guia.newChart()
          .asLineChart() // Usando construtor nativo de linha
          .addRange(rangeDadosGrafico)
          .setPosition(5, 7, 0, 0)
          .setTitle('Evolução Temporal dos Status Selecionados')
          .setOption('width', 900)
          .setOption('height', 450)
          .setLegendPosition(Charts.Position.TOP) // Força legenda pro TOPO nativamente
          .setPointStyle(Charts.PointStyle.MEDIUM) // Bolinhas nativas
          .setCurveStyle(Charts.CurveStyle.SMOOTH) // Curva suave nativa
          .setOption('series', seriesLabelAnalitico) // Ativa os números
          .setOption('useFirstColumnAsDomain', true) // Força o reconhecimento do eixo X e Cabeçalhos
          .setOption('hAxis', { showTextEvery: 1, slantedText: true, slantedTextAngle: 45 }) 
          .build();
          
        guia.insertChart(graficoLinhaAnalitico);
      } else {
        var rangeGrafico = guia.getRange(5, 2, dadosTabela.length + 1, 3);
        var graficoColuna = guia.newChart()
          .asColumnChart() // Usando construtor nativo de coluna
          .addRange(rangeGrafico)
          .setPosition(5, 7, 0, 0)
          .setTitle('Comparativo de Pendências: Início vs Atual')
          .setOption('width', 700)
          .setOption('height', 400)
          .setLegendPosition(Charts.Position.TOP) // Força legenda pro TOPO
          .setColors(['#95a5a6', '#e74c3c'])
          .setOption('series', { 0: { dataLabel: 'value' }, 1: { dataLabel: 'value' } })
          .setOption('useFirstColumnAsDomain', true)
          .build();
          
        guia.insertChart(graficoColuna);
      }
    }
    guia.autoResizeColumn(2); guia.autoResizeColumn(3); guia.autoResizeColumn(4); guia.autoResizeColumn(5);
  }

  function construirGuiaTemporal() {
    var nomeGuia = "Visão_Temporal_MS";
    var guia = ss.getSheetByName(nomeGuia);
    if (!guia) { 
        guia = ss.insertSheet(nomeGuia); 
    } else { 
      guia.clear();
      var charts = guia.getCharts();
      for (var c = 0; c < charts.length; c++) { 
          guia.removeChart(charts[c]);
      }
    }
    guia.setHiddenGridlines(true);
    
    guia.getRange("B2").setValue("📈 HISTÓRICO DIA A DIA (Curva de Processos por Setor)").setFontWeight("bold").setFontSize(14).setFontColor("#2c3e50");
    var rangeDados = guia.getRange(50, 2, matrizLinhas.length, matrizLinhas[0].length);
    rangeDados.setValues(matrizLinhas); 
    
    // Configura rótulos de dados forçando a injeção do nome da legenda
    var seriesLabelTemporal = {};
    for (var st = 0; st < setoresUnicos.length; st++) {
      seriesLabelTemporal[st] = { 
        dataLabel: 'value',
        labelInLegend: setoresUnicos[st],
        label: setoresUnicos[st]
      };
    }

    var graficoLinhas = guia.newChart()
        .asLineChart() // Usando construtor nativo de linha
        .addRange(rangeDados)
        .setPosition(4, 2, 0, 0)
        .setTitle('Evolução de Pendências (Macro Setores)')
        .setOption('width', 900)
        .setOption('height', 450)
        .setLegendPosition(Charts.Position.TOP) // Força legenda pro TOPO
        .setPointStyle(Charts.PointStyle.MEDIUM) // Bolinhas
        .setCurveStyle(Charts.CurveStyle.SMOOTH) // Curva suave
        .setOption('series', seriesLabelTemporal) // Numeração em cada ponto
        .setOption('useFirstColumnAsDomain', true)
        .setOption('hAxis', { showTextEvery: 1, slantedText: true, slantedTextAngle: 45 }) 
        .build();
        
    guia.insertChart(graficoLinhas);
  }

  construirGuiaVisao("Visão_Sintética_MS", mapaSintetico, "📈 EVOLUÇÃO MACRO: Por Setor (Resumo Executivo)");
  construirGuiaVisao("Visão_Analítica_MS", mapaAnalitico, "🔍 EVOLUÇÃO MICRO: Por Status Detalhado");
  construirGuiaTemporal();

  ui.alert("Painel Presidencial Completo!", "As 3 guias foram geradas. Variações de maiúsculas/minúsculas foram consolidadas em linhas únicas com sucesso!", ui.ButtonSet.OK);
}