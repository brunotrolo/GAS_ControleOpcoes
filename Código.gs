/**
 * ═══════════════════════════════════════════════════════════════
 * CÓDIGO.GS - PONTO DE ENTRADA E MENU PRINCIPAL
 * ═══════════════════════════════════════════════════════════════
 * RESPONSABILIDADES:
 * - Criar menu (onOpen) para interação do usuário com a interface (004).
 * - Servir a aplicação Web (doGet / include).
 * - Manter utilitários isolados de autorização.
 * ═══════════════════════════════════════════════════════════════
 */

/**
 * Cria o menu "⚙️ Automação" quando a planilha é aberta.
 * Conecta diretamente com as Pontes (Bridges) do arquivo 004_CoreServiceUI.
 */
function onOpen(e) {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('⚙️ Automação')
        .addItem('🚀 Rodar Fluxo Mestre (Planilha)', 'executarFluxoSequencial')
        .addItem('🤖 Scanner Completo (Sequência OPLab)', 'executarSequenciaScanner')
        .addSeparator()
        .addItem('📥 1. Atualizar Necton (Portfólio)', 'AtualizarNecton_Menu')
        .addItem('📈 2. Atualizar Dados Ativos (Ações)', 'AtualizarDadosAtivos_Menu')
        .addItem('🔍 3. Atualizar Detalhes (Opções)', 'AtualizarDetalhes_Menu')
        .addSeparator()
        .addItem('🧮 4a. Calcular Gregas (API OPLab)', 'AtualizarGregasAPI_Menu')
        .addItem('🔬 4b. Calcular Gregas (Nativo BS)', 'CalcularGregasNativo_Menu')
        .addSeparator()
        .addItem('📡 5. Scanner de Opções (SCANNER_OPCOES)', 'AtualizarScannerOpcoes_Menu')
        .addSeparator()
        .addItem('💰 6. Melhores Taxas Cobertas (PUT/CALL)', 'SyncBestCoveredOptionsRates_Menu')
        .addItem('📊 7. Maiores Volumes em Opções',           'SyncHighestOptionsVolume_Menu')
        .addItem('📈 8. Ranking Tendência M9M21',             'SyncM9M21Ranking_Menu')
        .addItem('🔗 9. Ranking Correlação IBOV',             'SyncCorrelIbovRanking_Menu')
        .addSeparator()
        .addItem('🎯 10. Screener Quantitativo (Trava de Alta PUT)', 'ScreenerQuantitativo_Menu')
        .addSeparator()
        .addItem('📉 11. Histórico de Preços (Spot 250D)', 'SyncHistoricoPrecos_Menu')
        .addToUi();
  } catch (err) {
    console.warn("[onOpen] Interface indisponível.");
  }
}

// ============================================================================
// SERVIDOR WEB (HTML SERVICE)
// ============================================================================

/**
 * Ponto de entrada para o Web App (Dashboard HTML).
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Stock Options | Intelligence')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Função vital para o sistema de slots e componentes HTML.
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    return ``;
  }
}


// ============================================================================
// API EXTERNA (WEBHOOK PARA O CLAUDE MCP)
// ============================================================================

function doPost(e) {
  try {
    // 1. Lê os dados enviados pelo servidor MCP
    const payload = JSON.parse(e.postData.contents);
    
    // 2. Trava de Segurança
    if (payload.token !== "TOKEN_SECRETO_OPLAB_2026") {
      return ContentService.createTextOutput(JSON.stringify({ status: "Erro", message: "Acesso Negado. Token inválido." }))
                           .setMimeType(ContentService.MimeType.JSON);
    }

    // 3. Roteador de Funções (Delega o comando para a função correta do seu onOpen)
    switch (payload.funcao) {
      case "executarFluxoSequencial": executarFluxoSequencial(); break;
      case "executarSequenciaScanner": executarSequenciaScanner(); break;
      case "AtualizarNecton_Menu": AtualizarNecton_Menu(); break;
      case "AtualizarDadosAtivos_Menu": AtualizarDadosAtivos_Menu(); break;
      case "AtualizarDetalhes_Menu": AtualizarDetalhes_Menu(); break;
      case "AtualizarGregasAPI_Menu": AtualizarGregasAPI_Menu(); break;
      case "CalcularGregasNativo_Menu": CalcularGregasNativo_Menu(); break;
      case "AtualizarScannerOpcoes_Menu": AtualizarScannerOpcoes_Menu(); break;
      case "SyncBestCoveredOptionsRates_Menu": SyncBestCoveredOptionsRates_Menu(); break;
      case "SyncHighestOptionsVolume_Menu": SyncHighestOptionsVolume_Menu(); break;
      case "SyncM9M21Ranking_Menu": SyncM9M21Ranking_Menu(); break;
      case "SyncCorrelIbovRanking_Menu": SyncCorrelIbovRanking_Menu(); break;
      case "ScreenerQuantitativo_Menu": ScreenerQuantitativo_Menu(); break;
      case "SyncHistoricoPrecos_Menu": SyncHistoricoPrecos_Menu(); break;
      default:
        throw new Error(`A função '${payload.funcao}' não está mapeada no roteador.`);
    }

    // 4. Retorna sucesso para o Claude saber que a planilha foi atualizada
    return ContentService.createTextOutput(JSON.stringify({ 
      status: "Sucesso", 
      message: `A função ${payload.funcao} foi executada na planilha OPLab.` 
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "Erro", message: err.toString() }))
                         .setMimeType(ContentService.MimeType.JSON);
  }
}
