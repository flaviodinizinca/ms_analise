// MENU DE INICIALIZAÇÃO
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('Gestão de Estoque')
      .addSubMenu(ui.createMenu('1. Sincronização Diária (Rodar Primeiro)')
          .addItem('1.1 Sincronizar Dados (Planilha Externa)', 'sincronizarDadosGerais')
          .addItem('1.2 Atualizar Status Entradas (Em Lote)', 'atualizarStatusEntradasLote'))
      .addSubMenu(ui.createMenu('2. Atualização de Processos SEI')
          .addItem('2.1 Consolidar Processos (Listagem Total)', 'sincronizarProcessosListagem')
          .addItem('2.2 Transferir Status p/ Listagem e Comparativo', 'transferirStatusPorProcesso'))
      .addSubMenu(ui.createMenu('3. Auditoria e Risco')
          .addItem('3.1 Comparar Itens Faltantes', 'compararItensFaltantes')
          .addItem('3.2 Classificar Nível de Risco (Itens e Comparativo)', 'classificarRiscoItens'))
      .addSubMenu(ui.createMenu('4. Fechamento e Dashboard (Fim do Dia)')
          .addItem('4.1 Arquivar Itens Resolvidos (Limpar Pauta)', 'arquivarItensResolvidos')
          .addItem('4.2 Gravar Dados Diários no Histórico', 'gravarHistoricoDashboard')
          .addSeparator()
          .addItem('4.3 👑 GERAR PAINEL (MS)', 'gerarPainelPresidencial')
          .addSeparator()
          .addItem('4.4 Atualizar Resumo (Cabeçalho da Guia Itens)', 'atualizarResumoCabecalho')) 
      .addSubMenu(ui.createMenu('5. Relatórios em PDF e Guias')
          .addItem('5.1 Gerar Guia "Processo Por Itens"', 'gerarProcessoPorItens')
          .addItem('5.2 Relatório Analítico Completo (PDF)', 'gerarRelatorioPDF')
          .addItem('5.3 Relatório de Regularizados (PDF)', 'gerarRelatorioRegularizadosPDF'))
      .addToUi();
}