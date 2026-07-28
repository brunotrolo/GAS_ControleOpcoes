/**
 * @fileoverview ConfigManager - v5.0 (Data Dictionary & Clean Architecture)
 * RESPONSABILIDADE: Centralizar o Dicionário Universal de Dados (DUD) e gerenciar o cache.
 * PADRÃO: Nomes de Planilha em UPPER_SNAKE_CASE | Chaves JSON em camelCase.
 */

const SYS_CONFIG = {
  // 1. MAPEAMENTO DE ABAS (Nomes exatos no Google Sheets)
  SHEETS: {
    IMPORT:         "NECTON_IMPORT",
    COCKPIT:        "COCKPIT",
    LOGS:           "LOGS",
    DETAILS:        "DADOS_DETALHES",
    ASSETS:         "DADOS_ATIVOS",
    GREEKS_API:     "DADOS_GREEKS",
    GREEKS_CALC:    "CALC_GREEKS",
    CONFIG:         "CONFIG_GLOBAL",
    SELECTION_OPT:    "SCANNER_OPCOES",
    BEST_RATES:       "SELECAO_OPCOES_MAIORES_LUCROS",
    HIGHEST_VOL:      "SELECAO_MAIORES_VOLUMES",
    RANK_M9M21:       "RANKING_TENDENCIA_M9M21",
    RANK_CORREL_IBOV: "RANKING_CORREL_IBOV",
    SCREENER_QUANT:   "SCREENER_QUANTITATIVO",
    HIST_PRECOS:      "HISTORICO_PRECOS_ATIVOS"
  },

  // 2. DICIONÁRIO UNIVERSAL DE DADOS (DUD)
  // Mapeia o [Rótulo da Planilha] -> [Chave JSON para o Web App]
  DUD: {
    "ID_TRADE":       "tradeId",
    "ID_STRATEGY":    "strategyId",
    "TICKER":         "ticker",         // Ação (ex: PETR4)
    "OPTION_TICKER":  "optionTicker",   // Opção (ex: PETRC425)
    "SPOT":           "spot",           // Preço da Ação
    "STRIKE":         "strike",         // Preço de Exercício
    "EXPIRY":         "expiry",         // Vencimento
    "STATUS_OP":      "status",         // Status Operacional
    "SIDE":           "side",           // C ou V
    "QUANTITY":       "quantity",
    "ENTRY_PRICE":    "entryPrice",
    "LAST_PREMIUM":   "lastPremium",
    "UPDATED_AT":     "updatedAt",
    "DELTA":          "delta",
    "GAMMA":          "gamma",
    "THETA":          "theta",
    "VEGA":           "vega",
    "IV":             "iv",
    "IV_RANK":        "ivRank",
    "DTE":            "dte",
    "PL_PCT":         "plPct",
    "PL_VALUE":       "plValue",
    "MONEYNESS":      "moneyness",
    "SCORE":          "score",

    // ── Campos de Volume e Mercado (complementares ao DUD) ───────────────────
    "VOLUME_QTY":   "volumeQty",    // Volume em contratos
    "VOLUME_FIN":   "volumeFin",    // Volume financeiro
    "BID_VOL":      "bidVol",       // Volume da melhor oferta de compra
    "ASK_VOL":      "askVol",       // Volume da melhor oferta de venda
    "BETA_IBOV":    "betaIbov",     // Beta vs IBOV
    "OPEN":         "open",         // Abertura do dia
    "HIGH":         "high",         // Máxima do dia
    "LOW":          "low",          // Mínima do dia
    "CLOSE":        "close",        // Fechamento / último preço
    "VARIATION":    "variation",    // Variação % do dia
    "CATEGORY":     "category",     // CALL ou PUT
    "SECTOR":       "sector",        // Setor da empresa

  }
};

/**
 * Gerenciador de Configurações com Cache de 3 Camadas
 */
const ConfigManager = {
  _memoryCache: null,
  _cacheKey: "APP_GLOBAL_CONFIGS_V5",
  _cacheTime: 21600, // 6 horas

  /**
   * Obtém configurações dinâmicas da aba CONFIG_GLOBAL.
   */
  get() {
    if (this._memoryCache) return this._memoryCache;

    const cache = CacheService.getScriptCache();
    const cachedData = cache.get(this._cacheKey);
    if (cachedData) {
      this._memoryCache = JSON.parse(cachedData);
      return this._memoryCache;
    }

    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheet = ss.getSheetByName(SYS_CONFIG.SHEETS.CONFIG);
      if (!sheet) return {};

      const data = sheet.getDataRange().getValues();
      const configs = {};
      
      for (let i = 1; i < data.length; i++) {
        const key = String(data[i][0]).trim();
        let val = data[i][1];
        if (key && !key.startsWith("//")) {
          if (typeof val === 'string' && val.includes(',')) {
            const parsed = parseFloat(String(val).replace(/\./g, '').replace(',', '.'));
            configs[key] = isNaN(parsed) ? val : parsed;
          } else {
            configs[key] = val;
          }
        }
      }

      this._memoryCache = configs;
      cache.put(this._cacheKey, JSON.stringify(configs), this._cacheTime);
      return configs;
    } catch (e) {
      console.error("[ConfigManager] Erro no I/O: " + (e.message));
      return {};
    }
  },

  /**
   * Invalida os caches.
   */
  clearCache() {
    CacheService.getScriptCache().remove(this._cacheKey);
    this._memoryCache = null;
  }
};
