/**
 * @fileoverview 020_SyncHistoricoPrecos.gs - v1.0
 * ═══════════════════════════════════════════════════════════════
 * RESPONSABILIDADE: Manter um histórico diário do preço de fechamento (spot)
 *   de cada ticker de DADOS_ATIVOS, na aba HISTORICO_PRECOS_ATIVOS.
 *   Alimenta o gráfico de evolução do spot no CardStrategiesBook (frontend).
 *
 * ENDPOINT:
 *   GET /market/historical/{ticker}/1d?amount=N&smooth=true&df=iso
 *   (via 000_CoreServiceAPIClient.gs → OplabService.getHistoricalData)
 *
 * ESTRATÉGIA:
 *   - Ticker NOVO (sem histórico salvo)  → backfill de 250 dias (1 chamada).
 *   - Ticker JÁ conhecido                → janela incremental de 7 dias
 *     (cobre fins de semana/feriados sem perder dados), deduplicado por
 *     TICKER+DATA contra o que já existe na aba.
 *   - Retenção: mantém só os últimos ~260 dias por ticker (poda no mesmo run),
 *     via reescrita completa da aba (1 leitura + 1 gravação, nunca célula a célula).
 *
 * EXECUÇÃO DIÁRIA:
 *   Este motor não se autoagenda — roda via menu, doPost (webhook) ou
 *   incluído na sequência do CoreOrchestrator (Config_Global).
 *
 * INTEGRAÇÃO COM A INFRAESTRUTURA:
 *   000 → OplabService.getHistoricalData()
 *   001 → SYS_CONFIG.SHEETS.HIST_PRECOS / SYS_CONFIG.SHEETS.ASSETS
 *   004 → SysLogger.log() / SysLogger.flush()
 *   005 → _menuBridge()
 * ═══════════════════════════════════════════════════════════════
 */

// ─── Configuração ─────────────────────────────────────────────────────────────
const HIST_PRECOS_CONFIG = {
  SHEET_NAME:       SYS_CONFIG.SHEETS.HIST_PRECOS,
  DIAS_BACKFILL:    250,  // histórico inicial p/ ticker novo
  DIAS_INCREMENTAL: 7,    // janela de segurança p/ ticker já conhecido
  DIAS_RETENCAO:    260   // poda o que passar disso (buffer acima dos 250 pedidos)
};

const HIST_PRECOS_HEADERS = ['TICKER', 'DATA', 'FECHAMENTO'];

// ─── Ponto de entrada (padrão _menuBridge de 005) ─────────────────────────────
function SyncHistoricoPrecos_Menu() {
  _menuBridge('Histórico de Preços', orquestrarSyncHistoricoPrecos);
}

// ─── Orquestrador ─────────────────────────────────────────────────────────────
function orquestrarSyncHistoricoPrecos() {
  var tInicio = Date.now();

  SysLogger.log('HistPrecos', 'START',
    '>>> INICIANDO SYNC HISTÓRICO DE PREÇOS (OPLab) <<<',
    JSON.stringify({ aba_destino: HIST_PRECOS_CONFIG.SHEET_NAME, timestamp: new Date().toISOString() })
  );

  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = _garantirAbaHistoricoPrecos(ss);

  // ── 1. Tickers alvo: reaproveita DADOS_ATIVOS (já mantido pelo motor 008) ──
  var tickers = _lerTickersAtivos(ss);
  if (tickers.length === 0) {
    SysLogger.log('HistPrecos', 'AVISO', 'Nenhum ticker encontrado em DADOS_ATIVOS. Nada a fazer.');
    SysLogger.flush();
    return;
  }

  // ── 2. Histórico já existente na aba (map ticker -> [{ts, close}] ordenado) ─
  var mapaExistente = _lerHistoricoExistente(sheet);

  // ── 3. Busca em paralelo: 250d para tickers novos, 7d para os já conhecidos ─
  var tApi = Date.now();
  var resultadosApi = _buscarHistoricoParalelo(tickers, mapaExistente);
  SysLogger.log('HistPrecos', 'INFO',
    tickers.length + ' tickers consultados na API (' + ((Date.now() - tApi) / 1000).toFixed(2) + 's).'
  );

  // ── 4. Mescla (dedup por TICKER+DATA) e poda para os últimos N dias ─────────
  var mapaFinal  = _mesclarEAparar(mapaExistente, resultadosApi, HIST_PRECOS_CONFIG.DIAS_RETENCAO);
  var linhas     = _flatten(mapaFinal);

  // ── 5. Reescreve a aba inteira de uma vez (nunca célula a célula) ───────────
  _limparDadosHistoricoPrecos(sheet);
  if (linhas.length > 0) {
    var tGs = Date.now();
    sheet.getRange(2, 1, linhas.length, HIST_PRECOS_HEADERS.length).setValues(linhas);
    SysLogger.log('HistPrecos', 'INFO',
      linhas.length + ' linhas gravadas no Sheets (' + ((Date.now() - tGs) / 1000).toFixed(2) + 's).'
    );
  }
  SpreadsheetApp.flush();

  var duracaoTotal = ((Date.now() - tInicio) / 1000).toFixed(1);
  SysLogger.log('HistPrecos', 'FINISH',
    '>>> SYNC CONCLUÍDO: ' + tickers.length + ' tickers | ' + linhas.length + ' linhas totais | ' + duracaoTotal + 's <<<',
    JSON.stringify({ tickers: tickers.length, linhas_totais: linhas.length, duracao_total_s: duracaoTotal })
  );
  SysLogger.flush();
}

// ─── Tickers alvo: reaproveita a lista já mantida por 008_CoreSyncStockData ───
function _lerTickersAtivos(ss) {
  var sheet = getPlanilhaDinamica(ss, SYS_CONFIG.SHEETS.ASSETS);
  if (!sheet) return [];
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var valores = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var vistos  = {};
  var out     = [];
  valores.forEach(function(row) {
    var t = String(row[0] || '').trim().toUpperCase();
    if (t && t !== 'ERRO_API' && t !== 'N/A' && !vistos[t]) { vistos[t] = true; out.push(t); }
  });
  return out;
}

// ─── Lê o histórico já salvo na aba, agrupado por ticker e ordenado por data ──
function _lerHistoricoExistente(sheet) {
  var mapa    = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return mapa;

  var valores = sheet.getRange(2, 1, lastRow - 1, HIST_PRECOS_HEADERS.length).getValues();
  valores.forEach(function(row) {
    var ticker = String(row[0] || '').trim().toUpperCase();
    var data   = row[1];
    var close  = parseFloat(row[2]) || 0;
    if (!ticker || !(data instanceof Date) || isNaN(data.getTime())) return;
    if (!mapa[ticker]) mapa[ticker] = [];
    mapa[ticker].push({ ts: data.getTime(), close: close });
  });

  Object.keys(mapa).forEach(function(t) {
    mapa[t].sort(function(a, b) { return a.ts - b.ts; });
  });
  return mapa;
}

// ─── Busca paralela: amount=250 p/ ticker novo, amount=7 p/ ticker conhecido ──
function _buscarHistoricoParalelo(tickers, mapaExistente) {
  var resultado = {};
  var token = PropertiesService.getScriptProperties().getProperty('OPLAB_ACCESS_TOKEN');
  if (!token) {
    SysLogger.log('HistPrecos', 'ERRO', 'OPLAB_ACCESS_TOKEN não configurado — sync abortado.');
    return resultado;
  }

  var hdrs    = { 'Access-Token': token.trim(), 'Accept': 'application/json' };
  var baseUrl = OplabService._baseUrl + '/market/historical/';

  var requests = tickers.map(function(t) {
    var amount = (mapaExistente[t] && mapaExistente[t].length > 0)
      ? HIST_PRECOS_CONFIG.DIAS_INCREMENTAL
      : HIST_PRECOS_CONFIG.DIAS_BACKFILL;
    var url = baseUrl + encodeURIComponent(t) + '/1d?amount=' + amount + '&smooth=true&df=iso';
    return { url: url, headers: hdrs, muteHttpExceptions: true };
  });

  var respostas;
  try {
    respostas = UrlFetchApp.fetchAll(requests);
  } catch (e) {
    SysLogger.log('HistPrecos', 'ERRO', 'fetchAll histórico falhou: ' + e.message);
    return resultado;
  }

  var ok = 0, falhas = [];
  for (var i = 0; i < respostas.length; i++) {
    var ticker = tickers[i];
    try {
      if (respostas[i].getResponseCode() !== 200) { falhas.push(ticker); continue; }
      var body = JSON.parse(respostas[i].getContentText());
      var itens = (body && Array.isArray(body.data)) ? body.data : [];
      resultado[ticker] = itens
        .filter(function(it) { return it && it.time != null && it.close != null; })
        .map(function(it) {
          return { ts: new Date(it.time).getTime(), close: parseFloat(it.close) || 0 };
        })
        .filter(function(it) { return !isNaN(it.ts); });
      ok++;
    } catch (e) {
      falhas.push(ticker);
    }
  }

  SysLogger.log('HistPrecos', 'INFO',
    'fetchAll histórico: ' + ok + '/' + tickers.length + ' tickers OK.' +
    (falhas.length > 0 ? ' Falharam: ' + falhas.join(', ') : '')
  );
  return resultado;
}

// ─── Mescla API + existente (dedup por dia), ordena e poda por retenção ───────
function _mesclarEAparar(mapaExistente, resultadosApi, diasRetencao) {
  var mapaFinal = {};

  function dayKey(ts) {
    var d = new Date(ts);
    return d.getFullYear() + '-' + d.getMonth() + '-' + d.getDate();
  }

  var tickers = {};
  Object.keys(mapaExistente).forEach(function(t) { tickers[t] = true; });
  Object.keys(resultadosApi).forEach(function(t) { tickers[t] = true; });

  Object.keys(tickers).forEach(function(t) {
    var vistos = {};
    var combinado = [];
    (mapaExistente[t] || []).forEach(function(item) {
      var k = dayKey(item.ts);
      if (!vistos[k]) { vistos[k] = true; combinado.push(item); }
    });
    (resultadosApi[t] || []).forEach(function(item) {
      var k = dayKey(item.ts);
      if (!vistos[k]) { vistos[k] = true; combinado.push(item); }
    });
    combinado.sort(function(a, b) { return a.ts - b.ts; });
    if (combinado.length > diasRetencao) {
      combinado = combinado.slice(combinado.length - diasRetencao);
    }
    if (combinado.length > 0) mapaFinal[t] = combinado;
  });

  return mapaFinal;
}

// ─── Achata o mapa final em linhas [TICKER, DATA(Date), FECHAMENTO] ───────────
function _flatten(mapaFinal) {
  var linhas = [];
  Object.keys(mapaFinal).sort().forEach(function(t) {
    mapaFinal[t].forEach(function(item) {
      linhas.push([t, new Date(item.ts), item.close]);
    });
  });
  return linhas;
}

// ─── Aba: garante existência + cabeçalho ──────────────────────────────────────
function _garantirAbaHistoricoPrecos(ss) {
  var nome  = HIST_PRECOS_CONFIG.SHEET_NAME;
  var sheet = ss.getSheetByName(nome);

  if (!sheet) {
    sheet = ss.insertSheet(nome);
    SysLogger.log('HistPrecos', 'INFO', 'Aba "' + nome + '" criada automaticamente.');
  }

  sheet.getRange(1, 1, 1, HIST_PRECOS_HEADERS.length).setValues([HIST_PRECOS_HEADERS]);
  return sheet;
}

function _limparDadosHistoricoPrecos(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
  }
}

// ─── Teste de homologação ─────────────────────────────────────────────────────
function testHistoricoPrecos() {
  console.log('=== HOMOLOGAÇÃO 020_SyncHistoricoPrecos v1.0 ===');
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var tickers = _lerTickersAtivos(ss);
  console.log('Tickers em DADOS_ATIVOS: ' + tickers.length);
  console.log(tickers.slice(0, 10).join(', ') + (tickers.length > 10 ? ', ...' : ''));

  if (tickers.length > 0) {
    var amostra = tickers.slice(0, 2);
    var resultado = _buscarHistoricoParalelo(amostra, {});
    amostra.forEach(function(t) {
      var itens = resultado[t] || [];
      console.log(t + ': ' + itens.length + ' dias retornados.');
      if (itens.length > 0) {
        var ult = itens[itens.length - 1];
        console.log('  Último: ' + new Date(ult.ts).toISOString().split('T')[0] + ' close=' + ult.close);
      }
    });
  }

  SysLogger.flush();
  console.log('=== FIM ===');
}
