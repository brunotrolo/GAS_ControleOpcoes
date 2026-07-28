/**
 * @fileoverview CoreOrchestrator - v4.2 (The Brain - Gold Standard)
 * RESPONSABILIDADE: Gerenciar a execução, validar ambiente e impor a ordem lógica.
 * INTEGRAÇÃO: Orquestra os Motores usando a Infraestrutura (000-004).
 */

// Mapa de nomes de exibição para cada função da sequência
const _STEP_DISPLAY_NAMES = {
  'AtualizarNecton_Menu':              'Atualizar Portfolio (Necton)',
  'AtualizarDadosAtivos_Menu':         'Sincronizar Ativos',
  'AtualizarDetalhes_Menu':            'Sincronizar Detalhes de Opções',
  'AtualizarGregasAPI_Menu':           'Gregas via API OPLab',
  'CalcularGregasNativo_Menu':         'Gregas Nativo (Black-Scholes)',
  'AtualizarScannerOpcoes_Menu':       'Scanner de Opções',
  'SyncHighestOptionsVolume_Menu':     'Maiores Volumes (PUT/CALL)',
  'SyncM9M21Ranking_Menu':             'Ranking Tendência M9M21',
  'SyncCorrelIbovRanking_Menu':        'Correlação IBOV',
  'SyncBestCoveredOptionsRates_Menu':  'Melhores Taxas Cobertas',
  'ScreenerQuantitativo_Menu':         'Screener Quantitativo (PUT)',
};

const CoreOrchestrator = {
  _serviceName: "CoreOrchestrator",

  /**
   * Registro Central de Serviços.
   * Centraliza onde cada motor de cálculo está localizado e seus nomes limpos.
   */
  get REGISTRY() {
    return {
      "ATUALIZAR_NECTON": {
        nome: "Atualizar Necton (Portfólio)",
        exec: () => typeof PortfolioUpdater !== 'undefined' ? PortfolioUpdater.syncPortfolioData() : console.warn("Motor 006 não carregado."),
        requer_token: true
      },
      "ATUALIZAR_ATIVOS": {
        nome: "Atualizar Dados Ativos",
        exec: () => typeof StockDataSync !== 'undefined' ? StockDataSync.run() : console.warn("Motor 007 não carregado."),
        requer_token: true
      },
      "ATUALIZAR_SCANNER": {
        nome: "Scanner de Opções (SCANNER_OPCOES)",
        exec: () => typeof CoreScannerOptions !== 'undefined' ? CoreScannerOptions.run() : console.warn("Motor 014 não carregado."),
        requer_token: true
      },
      "ATUALIZAR_BEST_RATES": {
        nome: "Melhores Taxas Cobertas PUT/CALL",
        exec: () => orquestrarSyncBestRates(),
        requer_token: true
      },
      "ATUALIZAR_HIGHEST_VOL": {
        nome: "Maiores Volumes em Opções",
        exec: () => orquestrarSyncHighestVolume(),
        requer_token: true
      },
      "ATUALIZAR_RANK_M9M21": {
        nome: "Ranking Tendência M9M21",
        exec: () => orquestrarSyncM9M21(),
        requer_token: true
      },
      "ATUALIZAR_RANK_CORREL": {
        nome: "Ranking Correlação IBOV",
        exec: () => orquestrarSyncCorrelIbov(),
        requer_token: true
      },
      "ATUALIZAR_SCREENER": {
        nome: "Screener Quantitativo (Trava de Alta PUT)",
        exec: () => typeof orquestrarScreener !== 'undefined' ? orquestrarScreener() : console.warn("Motor 019 não carregado."),
        requer_token: false
      },
      "ATUALIZAR_HISTORICO_PRECOS": {
        nome: "Histórico de Preços (Spot 250D)",
        exec: () => typeof orquestrarSyncHistoricoPrecos !== 'undefined' ? orquestrarSyncHistoricoPrecos() : console.warn("Motor 020 não carregado."),
        requer_token: true
      },
    };
  },

  /**
   * Checklist de Voo: Verifica se o sistema pode operar antes de rodar qualquer motor.
   */
  validarAmbiente() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const erros = [];

    // 1. Verifica abas essenciais
    [SYS_CONFIG.SHEETS.IMPORT, SYS_CONFIG.SHEETS.LOGS, SYS_CONFIG.SHEETS.CONFIG].forEach(aba => {
    if (!ss.getSheetByName(aba)) erros.push("Aba '" + (aba) + "' ausente.");
    });

    // 2. Verifica Token da API
    const token = PropertiesService.getScriptProperties().getProperty("OPLAB_ACCESS_TOKEN");
    if (!token) erros.push("Token OPLAB ausente no PropertiesService.");

    if (erros.length > 0) {
      const msg = "🚨 Falha de Ambiente:\n" + erros.join("\n");
      UIHandler.alert("Erro de Configuração", msg);
      
      // Uso de JSON.stringify para proteger a coluna de Timestamp no 003
      SysLogger.log(this._serviceName, "CRITICO", "Ambiente inválido para execução.", JSON.stringify({ falhas: erros }));
      return false;
    }
    return true;
  },


/**
   * Lê a aba Config_Global e extrai a sequência de funções a serem executadas.
   */
  getSequenciaDinamica(chave) {
    chave = chave || "Orquestrador_Sequencia_Padrao";
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const abaConfig = ss.getSheetByName(SYS_CONFIG.SHEETS.CONFIG);

      if (!abaConfig) {
        throw new Error("Aba 'Config_Global' não encontrada.");
      }

      // Lê as duas primeiras colunas (Chave e Valor) da aba de configuração
      const data = abaConfig.getDataRange().getValues();

      for (let i = 0; i < data.length; i++) {
        if (String(data[i][0]).trim() === chave) {
          const sequenciaRaw = String(data[i][1]).trim();
          
          if (!sequenciaRaw) return [];
          
          // Divide pelo ponto-e-vírgula e remove espaços extras
          return sequenciaRaw.split(';').map(f => f.trim()).filter(f => f.length > 0);
        }
      }
      return []; // Se não achar a chave, retorna vazio
    } catch (e) {
      SysLogger.log(this._serviceName, "ERRO", "Falha ao ler Config_Global", String(e.message));
      return [];
    }
  },



  /**
   * Executa um serviço individual.
   * @return {boolean} true se sucesso, false se falha.
   */
  executarServico(chave) {
    const servico = this.REGISTRY[chave];
    if (!servico) {
      SysLogger.log(this._serviceName, "ERRO", "Serviço '" + (chave) + "' não encontrado no Registro.", "");
      return false;
    }

    if (!this.validarAmbiente()) return false;

    try {
      SysLogger.log(this._serviceName, "INFO", "Delegando execução para: " + (servico.nome), "");
      servico.exec();
      SysLogger.flush();
      return true;
    } catch (e) {
      // e.message garante que enviaremos uma String ao Logger
      SysLogger.log(this._serviceName, "ERRO", "Falha catastrófica no motor: " + (servico.nome), String(e.message));
      SysLogger.flush();
      UIHandler.notify("Falha em " + (servico.nome), "Erro ❌");
      return false;
    }
  },

  /**
   * FLUXO MESTRE DINÂMICO: Lê a configuração da planilha e executa a sequência.
   */
  runFluxoMestre() {
    const sequencia = this.getSequenciaDinamica();
    
    if (sequencia.length === 0) {
      UIHandler.alert("Orquestrador", "Nenhuma sequência definida na chave 'Orquestrador_Sequencia_Padrao' da aba Config_Global.");
      return;
    }

    // MARCADOR DE TERRITÓRIO: INÍCIO DO FLUXO
    SysLogger.log(this._serviceName, "START", ">>> INICIANDO FLUXO MESTRE DINÂMICO <<<", JSON.stringify({ sequencia_encontrada: sequencia }));
    this._limparProgresso();

    // Captura o ambiente global do Apps Script para conseguir chamar funções pelo nome em texto
    const contextoGlobal = (function() { return this; })();

    for (let i = 0; i < sequencia.length; i++) {
      const nomeFuncao = sequencia[i];
      const nomeDisplay = _STEP_DISPLAY_NAMES[nomeFuncao] || nomeFuncao;
      this._anotarProgresso(i + 1, sequencia.length, nomeDisplay);
      // Exibe no canto da tela: "Passo 1/2: atualizarNecton..."
      UIHandler.notify("Passo " + (i+1) + "/" + (sequencia.length) + ": Executando " + (nomeFuncao) + "...", "Orquestrador");
      
      try {
        const funcaoAlvo = contextoGlobal[nomeFuncao];
        
        // Verifica se o texto digitado na planilha realmente é o nome de uma função no código
        if (typeof funcaoAlvo === 'function') {
          
          SysLogger.log(this._serviceName, "INFO", "Invocando passo: " + (nomeFuncao), "");
          funcaoAlvo(); // Executa a função magicamente aqui!
          SysLogger.flush();
          
        } else {
          throw new Error("A função '" + (nomeFuncao) + "' não existe no código. Verifique a ortografia na planilha.");
        }
        
      } catch (e) {
        SysLogger.log(this._serviceName, "ERRO", "Fluxo interrompido no passo: " + (nomeFuncao), String(e.message));
        SysLogger.flush();
        UIHandler.alert("Fluxo Interrompido", "Falha ao executar \"" + (nomeFuncao) + "\".\n\nErro: " + (e.message) + "\n\nO processo foi parado por segurança.");
        return; // Curto-circuito: para tudo!
      }
    }

    // MARCADOR DE TERRITÓRIO: FINALIZAÇÃO
    SysLogger.log(this._serviceName, "FINISH", ">>> FLUXO MESTRE CONCLUÍDO COM SUCESSO <<<", JSON.stringify({ total_passos: sequencia.length }));
    SysLogger.flush();
    this._limparProgresso();
    UIHandler.alert("Fluxo Concluído", "Sincronização global concluída!\nPassos executados: " + (sequencia.join(", ")));
  },

  /**
   * SEQUÊNCIA SCANNER: Lê 'Orquestrador_Sequencia_Scanner' e executa em cascata.
   */
  runFluxoScanner() {
    const sequencia = this.getSequenciaDinamica("Orquestrador_Sequencia_Scanner");

    if (sequencia.length === 0) {
      UIHandler.alert("Orquestrador", "Nenhuma sequência definida na chave 'Orquestrador_Sequencia_Scanner' da aba Config_Global.");
      return;
    }

    SysLogger.log(this._serviceName, "START", ">>> INICIANDO SEQUÊNCIA SCANNER <<<", JSON.stringify({ sequencia_encontrada: sequencia }));
    this._limparProgresso();

    const contextoGlobal = (function() { return this; })();

    for (let i = 0; i < sequencia.length; i++) {
      const nomeFuncao = sequencia[i];
      const nomeDisplay = _STEP_DISPLAY_NAMES[nomeFuncao] || nomeFuncao;
      this._anotarProgresso(i + 1, sequencia.length, nomeDisplay);
      UIHandler.notify("Scanner " + (i+1) + "/" + sequencia.length + ": " + nomeFuncao + "...", "Scanner");
      try {
        const funcaoAlvo = contextoGlobal[nomeFuncao];
        if (typeof funcaoAlvo === 'function') {
          SysLogger.log(this._serviceName, "INFO", "Invocando: " + nomeFuncao, "");
          funcaoAlvo();
          SysLogger.flush();
        } else {
          throw new Error("Função '" + nomeFuncao + "' não existe no código.");
        }
      } catch (e) {
        SysLogger.log(this._serviceName, "ERRO", "Scanner interrompido em: " + nomeFuncao, String(e.message));
        SysLogger.flush();
        UIHandler.alert("Scanner Interrompido", "Falha em \"" + nomeFuncao + "\".\n\nErro: " + e.message + "\n\nProcesso parado por segurança.");
        return;
      }
    }

    SysLogger.log(this._serviceName, "FINISH", ">>> SEQUÊNCIA SCANNER CONCLUÍDA <<<", JSON.stringify({ total_passos: sequencia.length }));
    SysLogger.flush();
    this._limparProgresso();
    UIHandler.alert("Scanner Concluído", "Sequência concluída!\nPassos: " + sequencia.join(", "));
  },

  _anotarProgresso(etapa, total, nomeEtapa) {
    try {
      PropertiesService.getScriptProperties().setProperty('ORQUESTRADOR_PROGRESSO', JSON.stringify({
        ativo: true, etapa: etapa, total: total, nomeEtapa: nomeEtapa || '', ts: new Date().toISOString()
      }));
    } catch(_) {}
  },

  _limparProgresso() {
    try { PropertiesService.getScriptProperties().deleteProperty('ORQUESTRADOR_PROGRESSO'); } catch(_) {}
  },

};

// ============================================================================
// PONTES DE COMPATIBILIDADE E FLUXO MESTRE
// ============================================================================

/** Função acionada pelo menu para rodar tudo na sequência correta */
function executarFluxoSequencial() {
  CoreOrchestrator.runFluxoMestre();
}

/** Roda a sequência definida em Orquestrador_Sequencia_Scanner no Config_Global */
function executarSequenciaScanner() {
  CoreOrchestrator.runFluxoScanner();
}

/**
 * Retorna a lista de etapas de um fluxo a partir do Config_Global.
 * @param {string} fluxo - 'mestre' ou 'scanner'
 * @returns {Array<{nome: string, displayName: string}>}
 */
function getFluxoSteps(fluxo) {
  const chave = fluxo === 'scanner'
    ? 'Orquestrador_Sequencia_Scanner'
    : 'Orquestrador_Sequencia_Padrao';
  const sequencia = CoreOrchestrator.getSequenciaDinamica(chave);
  return sequencia.map(function(fn) {
    return { nome: fn, displayName: _STEP_DISPLAY_NAMES[fn] || fn };
  });
}

/**
 * Lê o progresso atual do orquestrador gravado nas ScriptProperties.
 * @returns {{ativo: boolean, etapa: number, total: number, nomeEtapa: string, ts: string}|null}
 */
function getProgressoAtual() {
  try {
    const raw = PropertiesService.getScriptProperties().getProperty('ORQUESTRADOR_PROGRESSO');
    if (!raw) return null;
    return JSON.parse(raw);
  } catch(_) {
    return null;
  }
}

