/**
 * ============================================================================
 * API.gs - The "State of the Art" Universal Backend Repository (Refatorado)
 * FOCO ATUAL: 10 Abas Core (Ingestão, Base, Cálculo e Interface)
 * ============================================================================
 */

// ==========================================
// 1. READ (Leitura Universal de Todo o Banco)
// ==========================================

/**
 * 🚀 FASE 1: VOO RÁPIDO
 * Carrega apenas o essencial para a tela acender imediatamente.
 *
 * 🔧 FIX: Strings literais substituídas por SYS_CONFIG.SHEETS.*
 *
 * ⚠️ ATENÇÃO: Mantido getDisplayValues() intencionalmente.
 *    O GAS não serializa Date objects via google.script.run — getValues() retorna
 *    Dates nativos que quebram a comunicação com o frontend com response: null.
 *    A conversão de tipos fica exclusivamente no Tradutor.html.
 */
function getDadosLight() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abasEssenciais = [
      SYS_CONFIG.SHEETS.COCKPIT,
      SYS_CONFIG.SHEETS.CONFIG
    ];

    const data = {
      success: true,
      timestamp: new Date().toLocaleString('pt-BR'),
      raw: {}
    };

    // Sprint 4: single getSheets() call → eliminates N redundant getSheets() + getLastRow() per sheet
    const sheetsMap = _buildSheetsMap(ss);
    abasEssenciais.forEach(function(nomeAba) {
      const sheet = sheetsMap.get(String(nomeAba).toUpperCase());
      if (sheet) {
        const values = sheet.getDataRange().getDisplayValues();
        data.raw[sheet.getName()] = (values.length === 1 && values[0][0] === '') ? [] : values;
      }
    });

    return data;
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

/**
 * 🚚 FASE 2: CARGA PESADA (Background)
 * Otimizado APENAS para as 10 abas do Core de Operações atuais.
 * Séries legadas (400, 500, 600) foram expurgadas para economizar memória e tempo de requisição.
 *
 * 🔧 FIX: Strings literais substituídas por SYS_CONFIG.SHEETS.*
 * ⚠️ ATENÇÃO: Mantido getDisplayValues() intencionalmente — mesma razão do getDadosLight.
 */
function getAbasPesadas() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const abasPesadas = [
      // 1. Ingestão & Logs
      SYS_CONFIG.SHEETS.IMPORT,
      SYS_CONFIG.SHEETS.LOGS,

      // 2. Base de Dados & Histórico
      SYS_CONFIG.SHEETS.ASSETS,
      SYS_CONFIG.SHEETS.DETAILS,

      // 3. Motor Matemático
      SYS_CONFIG.SHEETS.GREEKS_CALC,
      SYS_CONFIG.SHEETS.GREEKS_API
    ];

    const data = { success: true, timestamp: new Date().toLocaleString('pt-BR'), raw: {} };

    // Sprint 4: single getSheets() call → eliminates N redundant getSheets() + getLastRow() per sheet
    const sheetsMap = _buildSheetsMap(ss);
    abasPesadas.forEach(function(nomeAba) {
      const sheet = sheetsMap.get(String(nomeAba).toUpperCase());
      if (sheet) {
        const values = sheet.getDataRange().getDisplayValues();
        data.raw[sheet.getName()] = (values.length === 1 && values[0][0] === '') ? [] : values;
      }
    });

    return data;
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// FUNÇÕES AUXILIARES (Core de Busca)
// ==========================================

/**
 * 🛡️ BUSCA DE PLANILHA DINÂMICA
 * Encontra a aba independentemente de Case Sensitive (maiúsculas/minúsculas).
 */
function getPlanilhaDinamica(planilhaAtiva, nomeProcurado) {
  const abas = planilhaAtiva.getSheets();
  const nomeProcuradoUpper = String(nomeProcurado).toUpperCase();

  // Otimização: For loop tradicional é levemente mais rápido que .find() no V8 do Apps Script
  for (let i = 0; i < abas.length; i++) {
    if (abas[i].getName().toUpperCase() === nomeProcuradoUpper) {
      return abas[i];
    }
  }
  return null;
}

/**
 * Sprint 4 optimization: builds a Map<UPPERNAME, Sheet> from a single getSheets() call.
 * Used by getDadosLight() and getAbasPesadas() to avoid per-sheet getSheets() overhead.
 * getPlanilhaDinamica() is kept unchanged for all other callers.
 */
function _buildSheetsMap(ss) {
  const map = new Map();
  ss.getSheets().forEach(function(sheet) {
    map.set(sheet.getName().toUpperCase(), sheet);
  });
  return map;
}

// ==========================================
// 1b. LEITURA DIRETA DE ABA
// ==========================================

/**
 * Lê qualquer aba do Sheets e retorna como array de objetos JSON.
 * Linha 1 = cabeçalho. Linhas vazias são ignoradas.
 * Retorna [] se a aba não existir ou estiver vazia.
 *
 * Uso no frontend:
 *   google.script.run
 *     .withSuccessHandler(resolve)
 *     .lerAbaComoJSON('DADOS_ATIVOS');
 */
function lerAbaComoJSON(nomeAba) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getPlanilhaDinamica(ss, nomeAba);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const valores = sheet.getDataRange().getDisplayValues();
    const cabecalho = valores[0].map(function(h) { return String(h).trim(); });

    var resultado = [];
    for (var i = 1; i < valores.length; i++) {
      var linha = valores[i];
      // Ignora linhas completamente vazias
      if (linha.every(function(cel) { return String(cel).trim() === ''; })) continue;

      var obj = {};
      obj._linhaReal = i + 1; // número real da linha na planilha (linha 1 = cabeçalho)
      cabecalho.forEach(function(col, j) {
        if (col) obj[col] = linha[j] !== undefined ? linha[j] : '';
      });
      resultado.push(obj);
    }

    return resultado;
  } catch (e) {
    console.error('[lerAbaComoJSON] Erro ao ler aba ' + nomeAba + ': ' + e.message);
    return [];
  }
}

// ==========================================
// 1b-2. HISTÓRICO DE PREÇOS (sob demanda — não entra em getAbasPesadas
//        para não inflar o payload principal com N tickers x 250 dias)
// ==========================================

/**
 * Retorna a série de fechamento (spot) de um ticker a partir de uma data,
 * para alimentar o gráfico de evolução do spot no CardStrategiesBook.
 * Alimentado pela aba HISTORICO_PRECOS_ATIVOS (motor 020_SyncHistoricoPrecos).
 *
 * @param {string} ticker  - ex: "PETR4"
 * @param {string} desdeISO - "YYYY-MM-DD" (opcional). Sem filtro se omitido.
 * @returns {Array<{data:string, fechamento:number}>} ordenado por data crescente.
 *
 * ⚠️ Lê com getValues() (não getDisplayValues()) de propósito: a coluna DATA
 *    precisa ser comparada como Date para o filtro "desde". Cada Date é
 *    convertida para string ISO plana ANTES de retornar — nunca devolve um
 *    objeto Date cru pela fronteira do google.script.run (ele vira null).
 */
function getHistoricoSpotEstrutura(ticker, desdeISO) {
  try {
    if (!ticker) return [];
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getPlanilhaDinamica(ss, SYS_CONFIG.SHEETS.HIST_PRECOS);
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return [];

    const tickerAlvo = String(ticker).trim().toUpperCase();
    const desde       = desdeISO ? new Date(desdeISO + 'T00:00:00') : null;

    const valores = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // TICKER | DATA | FECHAMENTO
    const out = [];
    for (var i = 0; i < valores.length; i++) {
      var row = valores[i];
      if (String(row[0] || '').trim().toUpperCase() !== tickerAlvo) continue;
      var d = row[1];
      if (!(d instanceof Date) || isNaN(d.getTime())) continue;
      if (desde && d.getTime() < desde.getTime()) continue;

      var iso = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
      out.push({ data: iso, fechamento: parseFloat(row[2]) || 0, _ts: d.getTime() });
    }

    out.sort(function(a, b) { return a._ts - b._ts; });
    out.forEach(function(o) { delete o._ts; });
    return out;
  } catch (e) {
    console.error('[getHistoricoSpotEstrutura] Erro para ' + ticker + ': ' + e.message);
    return [];
  }
}

/**
 * Painel de Ações (estilo Google Finance) para o CardRadar.
 * Junta DADOS_ATIVOS (ticker, nome, preço, variação) com a série de
 * fechamento dos últimos 30 dias (HISTORICO_PRECOS_ATIVOS, motor 020).
 *
 * @returns {Array<{ticker,nome,preco,variacao,serie:number[]}>} ordenado por ticker.
 *
 * Uma única chamada (sob demanda, quando a subaba é aberta) — não entra no
 * getAbasPesadas para não pesar a carga principal. A série vai como array de
 * números (fechamentos ordenados por data), nunca objetos Date.
 */
function getPainelAcoes() {
  try {
    const ss        = SpreadsheetApp.getActiveSpreadsheet();
    const abaAtivos = getPlanilhaDinamica(ss, SYS_CONFIG.SHEETS.ASSETS);
    if (!abaAtivos) return [];

    const lastRow = abaAtivos.getLastRow();
    const lastCol = abaAtivos.getLastColumn();
    if (lastRow < 2) return [];

    const headers = abaAtivos.getRange(1, 1, 1, lastCol).getValues()[0];
    const col = {};
    headers.forEach(function(h, i) { col[String(h).trim().toUpperCase()] = i; });
    const cTicker = col['TICKER'], cNome = col['COMPANY_NAME'], cSpot = col['SPOT'], cVar = col['VARIATION'];
    if (cTicker == null) return [];

    const serieMap = _lerSeries30d(ss);
    const valores  = abaAtivos.getRange(2, 1, lastRow - 1, lastCol).getValues();

    const out = [];
    valores.forEach(function(row) {
      const ticker = String(row[cTicker] || '').trim().toUpperCase();
      if (!ticker || ticker === 'ERRO_API' || ticker === 'N/A') return;
      out.push({
        ticker:   ticker,
        nome:     cNome != null ? String(row[cNome] || '') : '',
        preco:    cSpot != null ? (parseFloat(row[cSpot]) || 0) : 0,
        variacao: cVar  != null ? (parseFloat(row[cVar])  || 0) : 0,
        serie:    serieMap[ticker] || []
      });
    });

    out.sort(function(a, b) { return a.ticker.localeCompare(b.ticker); });
    return out;
  } catch (e) {
    console.error('[getPainelAcoes] ' + e.message);
    return [];
  }
}

/**
 * Lê HISTORICO_PRECOS_ATIVOS e devolve { ticker: [fechamento, ...] } com os
 * últimos 30 dias corridos, ordenado por data crescente.
 */
function _lerSeries30d(ss) {
  const mapa  = {};
  const sheet = getPlanilhaDinamica(ss, SYS_CONFIG.SHEETS.HIST_PRECOS);
  if (!sheet) return mapa;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return mapa;

  const corte = new Date();
  corte.setDate(corte.getDate() - 30);
  corte.setHours(0, 0, 0, 0);
  const corteTs = corte.getTime();

  const valores = sheet.getRange(2, 1, lastRow - 1, 3).getValues(); // TICKER | DATA | FECHAMENTO
  const tmp = {};
  valores.forEach(function(row) {
    const t = String(row[0] || '').trim().toUpperCase();
    const d = row[1];
    const c = parseFloat(row[2]) || 0;
    if (!t || !(d instanceof Date) || isNaN(d.getTime()) || d.getTime() < corteTs) return;
    if (!tmp[t]) tmp[t] = [];
    tmp[t].push({ ts: d.getTime(), close: c });
  });

  Object.keys(tmp).forEach(function(t) {
    tmp[t].sort(function(a, b) { return a.ts - b.ts; });
    mapa[t] = tmp[t].map(function(x) { return x.close; });
  });
  return mapa;
}

// ==========================================
// 1c. EXPORTAR CSV (Download de qualquer aba)
// ==========================================

/**
 * Exporta qualquer aba como array 2D de display values para o frontend gerar o CSV.
 * COCKPIT usa linha 10 como cabeçalho; todas as demais abas usam linha 1.
 */
function exportarAbaCSV(nomeAba) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getPlanilhaDinamica(ss, nomeAba);
    if (!sheet) return { success: false, error: 'Aba "' + nomeAba + '" não encontrada.' };

    const linhaHeader = (nomeAba === SYS_CONFIG.SHEETS.COCKPIT) ? 10 : 1;
    const lastRow  = sheet.getLastRow();
    const lastCol  = sheet.getLastColumn();

    if (lastRow < linhaHeader || lastCol === 0) return { success: true, rows: [] };

    const numLinhas = lastRow - linhaHeader + 1;
    const rows = sheet
      .getRange(linhaHeader, 1, numLinhas, lastCol)
      .getDisplayValues();

    return { success: true, rows: rows };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

// ==========================================
// 2. CREATE (Inserção em Lote)
// ==========================================

/**
 * 🔧 FIX: Lógica de busca de linha vazia substituída por getLastRow() + 1.
 *         A abordagem anterior era frágil — qualquer célula vazia no meio dos
 *         dados causava inserção na posição errada.
 */
function apiAdicionarLinhas(nomeAba, dadosMatriz) {
  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getPlanilhaDinamica(ss, nomeAba);
    if (!sheet) throw new Error(`Aba [${nomeAba}] não existe no banco de dados.`);
    if (!dadosMatriz || dadosMatriz.length === 0) return { success: true, message: "Nenhum dado para inserir." };

    // Determina o limite seguro de colunas para escrita: para antes da primeira coluna
    // que contenha fórmula na linha 2 (colunas de fórmula não devem ser sobrescritas).
    // Se a aba estiver vazia ou a linha 2 não tiver fórmulas, usa o comprimento do array.
    const lastCol   = sheet.getLastColumn();
    let safeMaxCols = dadosMatriz[0].length;
    if (lastCol > 0 && sheet.getLastRow() >= 2) {
      const formulas2 = sheet.getRange(2, 1, 1, lastCol).getFormulas()[0];
      for (let c = 0; c < formulas2.length; c++) {
        if (formulas2[c] !== '') {
          safeMaxCols = Math.min(safeMaxCols, c); // para ANTES da coluna com fórmula
          break;
        }
      }
    }

    // Se a 1ª coluna já for fórmula, não há onde gravar com segurança
    if (safeMaxCols < 1) throw new Error("Coluna 1 contém fórmula; nada pode ser gravado com segurança nesta aba.");

    // Normaliza todas as linhas para exatamente safeMaxCols colunas (trim/pad)
    const writeCols  = safeMaxCols;
    const normalizado = dadosMatriz.map(function(row) {
      const r = row.slice(0, writeCols);
      while (r.length < writeCols) r.push('');
      return r;
    });

    const lastDataRow = sheet.getLastRow();
    let startRow = lastDataRow + 1;
    if (lastDataRow > 0) {
      const colA = sheet.getRange(1, 1, lastDataRow, 1).getValues();
      for (let i = 0; i < colA.length; i++) {
        if (String(colA[i][0]).trim() === '') {
          startRow = i + 1;
          break;
        }
      }
    }

    sheet.getRange(startRow, 1, normalizado.length, writeCols).setValues(normalizado);
    SpreadsheetApp.flush();

    return { success: true, message: `${normalizado.length} linhas adicionadas em [${nomeAba}] a partir da linha ${startRow}.` };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// 3. UPDATE (Atualização de Chave-Valor)
// ==========================================

function apiAtualizarChaveValor(nomeAba, payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getPlanilhaDinamica(ss, nomeAba);
    if (!sheet) throw new Error(`Aba de configurações [${nomeAba}] não encontrada.`);

    const data = sheet.getDataRange().getValues();
    const chavesNovas = Object.keys(payload);
    let atualizacoes = 0;

    // Copia coluna 2 para edição em memória — Config_Global usa apenas valores
    const colB = data.map(row => [row[1]]);
    for (let i = 0; i < data.length; i++) {
      const chavePlanilha = String(data[i][0]).trim();
      if (chavesNovas.includes(chavePlanilha)) {
        colB[i][0] = payload[chavePlanilha];
        atualizacoes++;
      }
    }

    // Escreve coluna 2 inteira em um único setValues (batch)
    if (atualizacoes > 0) {
      sheet.getRange(1, 2, colB.length, 1).setValues(colB);
    }

    SpreadsheetApp.flush();
    return { success: true, message: `${atualizacoes} chaves atualizadas com sucesso.` };
  } catch (e) {
    return { success: false, error: e.toString() };
  }
}

function apiSetCellValue(nomeAba, linha, coluna, valor) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 🔧 FIX: busca a aba ignorando maiúsculas/minúsculas
    // getSheetByName é case-sensitive — 'Cockpit' != 'COCKPIT'
    const sheets = ss.getSheets();
    const sheet  = sheets.find(s => s.getName().toUpperCase() === String(nomeAba).toUpperCase());
    if (!sheet) throw new Error(`Aba [${nomeAba}] não encontrada. Abas disponíveis: ${sheets.map(s => s.getName()).join(', ')}`);

    if (!linha || linha < 1) throw new Error(`Linha inválida: ${linha}`);
    if (!coluna || coluna < 1) throw new Error(`Coluna inválida: ${coluna}`);

    sheet.getRange(linha, coluna).setValue(valor);
    SpreadsheetApp.flush();
    return { success: true, timestamp: new Date().toLocaleTimeString() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// 4. DELETE & TRUNCATE (Exclusão e Limpeza)
// ==========================================

function apiExcluirLinhaSegura(nomeAba, numeroLinha, valorEsperadoColunaA) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getPlanilhaDinamica(ss, nomeAba);
    if (!sheet) throw new Error(`Aba [${nomeAba}] não existe no banco de dados.`);

    const valorPlanilha = String(sheet.getRange(numeroLinha, 1).getDisplayValue() || "").trim().toUpperCase();
    const valorSeguro = String(valorEsperadoColunaA || "").trim().toUpperCase();

    // Trava anti-dessincronização (Double Check)
    if (valorPlanilha !== valorSeguro) {
      return { success: false, error: `Falha de sincronia: Esperava encontrar [${valorSeguro}], mas encontrou [${valorPlanilha}] na linha ${numeroLinha}. Exclusão abortada.` };
    }

    sheet.deleteRow(numeroLinha);
    SpreadsheetApp.flush();
    return { success: true, message: `Registro [${valorSeguro}] removido de [${nomeAba}].` };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiExcluirLinhasEmLote(nomeAba, listaLinhas) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getPlanilhaDinamica(ss, nomeAba);
    if (!sheet) throw new Error("Aba não encontrada: " + nomeAba);

    // FILTRO DE SEGURANÇA: Remove nulos, converte para inteiro e ordena de baixo para cima
    const linhasOrdenadas = listaLinhas
      .filter(l => l !== null && l !== undefined && !isNaN(l))
      .map(l => parseInt(l, 10))
      .filter(l => l > 0) // Impede deleção de linha 0 ou negativa
      .sort((a, b) => b - a); // ⚠️ OBRIGATÓRIO: Deletar de baixo para cima para não mudar o índice das linhas de cima

    linhasOrdenadas.forEach(linha => {
      sheet.deleteRow(linha);
    });

    SpreadsheetApp.flush();
    return { success: true, count: linhasOrdenadas.length };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

function apiLimparAba(nomeAba, manterLinhasTop = 1, mensagemAuditoria = null) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getPlanilhaDinamica(ss, nomeAba); // case-insensitive (consistente com o resto)
    if (!sheet) throw new Error(`Aba [${nomeAba}] não existe no banco de dados.`);

    const lastRow = sheet.getLastRow();
    if (lastRow > manterLinhasTop) {
      sheet.getRange(manterLinhasTop + 1, 1, lastRow - manterLinhasTop, sheet.getLastColumn()).clearContent();
    }

    // Rastro de auditoria
    if (mensagemAuditoria) {
      const ts = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss");
      sheet.getRange(manterLinhasTop + 1, 1, 1, 4).setValues([[ts, "SYSTEM", "AVISO", mensagemAuditoria]]);
    }

    SpreadsheetApp.flush();
    return { success: true, message: `Aba [${nomeAba}] foi limpa, mantendo ${manterLinhasTop} linha(s) de cabeçalho.` };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// ==========================================
// 5. EXTERNAL API BRIDGE (Integrações de Terceiros)
// ==========================================

function apiIntegracaoOpLab(ticker) {
  if (!ticker || String(ticker).trim() === '') {
    return { success: false, error: 'Ticker nao fornecido.' };
  }

  try {
    var cleanTicker = String(ticker).toUpperCase().trim();
    var data = OplabService.getOptionDetails(cleanTicker);

    if (!data) return { success: false, error: 'Opcao [' + cleanTicker + '] nao encontrada ou sem liquidez na API OpLab.' };

    // Extrai a data de vencimento no formato YYYY-MM-DD
    var expiry = '';
    if (data.due_date) {
      // due_date pode vir como "2026-05-15T03:00:00.000Z" ou "2026-05-15"
      expiry = String(data.due_date).substring(0, 10);
    } else if (data.days_to_maturity) {
      // Fallback: calcula a data aproximada somando dte a hoje
      var dataVenc = new Date();
      dataVenc.setDate(dataVenc.getDate() + parseInt(data.days_to_maturity || 0));
      var y  = dataVenc.getFullYear();
      var m  = String(dataVenc.getMonth() + 1).padStart(2, '0');
      var d  = String(dataVenc.getDate()).padStart(2, '0');
      expiry = y + '-' + m + '-' + d;
    }

    // Busca a Selic atual da Config_Global para pre-popular o campo de taxa
    var taxaJuros = 14.75; // fallback padrao
    try {
      var cfg = ConfigManager.get();
      if (cfg && cfg['Taxa_Selic_Anual']) {
        const selicParsed = parseFloat(String(cfg['Taxa_Selic_Anual']).replace(',', '.'));
        // taxaJuros é exibido em PONTOS no campo da UI. Aceita "0,1475" (fração ⇒ *100)
        // ou "14,75" (já em pontos ⇒ usa como está).
        taxaJuros = isNaN(selicParsed) ? 14.75 : (selicParsed > 1 ? selicParsed : selicParsed * 100);
      }
    } catch (eCfg) {
      // silencioso -- usa o fallback
    }

    // Busca IV (volatilidade implicita) do ativo mae na aba DADOS_ATIVOS
    // Coluna IV = volatilidade implicita atual (calculada pelo 008_CoreSyncStockData)
    var ivAtivo = null;
    try {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var abaAtivos = getPlanilhaDinamica(ss, SYS_CONFIG.SHEETS.ASSETS);
      if (abaAtivos && abaAtivos.getLastRow() > 1) {
        var headersAtivos = abaAtivos.getRange(1, 1, 1, abaAtivos.getLastColumn()).getValues()[0];
        var colTicker = -1, colIV = -1;
        for (var hi = 0; hi < headersAtivos.length; hi++) {
          var h = String(headersAtivos[hi]).trim().toUpperCase();
          if (h === 'TICKER') colTicker = hi;
          if (h === 'IV')     colIV     = hi;
        }
        if (colTicker >= 0 && colIV >= 0) {
          // Campo correto na API OpLab: 'parent_symbol' (confirmado via debug)
          // 'underlying' nao existe neste endpoint -- vem undefined
          var tickerBusca = String(
            data.parent_symbol ||
            data.underlying    ||
            data.stock         ||
            ''
          ).trim().toUpperCase();

          // Fallback: extrair ticker mae removendo sufixo de opcao do symbol
          // Ex: 'SANBJ349' -> tenta encontrar 'SANB11' via lookup direto
          // Nao tentamos regex aqui pois e dificil de acertar para todos os casos --
          // confiamos no campo 'underlying' da API

          if (tickerBusca) {
            var valoresAtivos = abaAtivos.getRange(2, 1, abaAtivos.getLastRow() - 1, abaAtivos.getLastColumn()).getValues();
            for (var ri = 0; ri < valoresAtivos.length; ri++) {
              if (String(valoresAtivos[ri][colTicker]).trim().toUpperCase() === tickerBusca) {
                var ivBruto = valoresAtivos[ri][colIV];

                // FIX 2: DADOS_ATIVOS usa getDisplayValues() -- IV pode vir como:
                //   numero puro  : 0.3359  (decimal, ja na forma correta)
                //   string BR    : "33,59" (percentual com virgula)
                //   string US    : "33.59" (percentual com ponto)
                // Regra: se o valor numerico for > 1, e percentual -> dividir por 100
                var ivNum = parseFloat(String(ivBruto).replace('%', '').replace(',', '.').trim());
                if (!isNaN(ivNum) && ivNum > 0) {
                  ivAtivo = ivNum > 1 ? ivNum / 100 : ivNum;
                }
                break;
              }
            }
          }
        }
      }
    } catch (eIV) {
      // silencioso -- ivAtivo permanece null, frontend exibe campo manual
      SysLogger.log('apiIntegracaoOpLab', 'AVISO', 'Falha ao buscar IV de DADOS_ATIVOS: ' + eIV.message);
    }

    // Calcula DELTA via Black-Scholes usando OptionMath (011_CoreCalcGreeks)
    var delta = null;
    try {
      var S2     = parseFloat(data.spot_price || data.spot || 0);
      var K2     = parseFloat(data.strike || 0);
      var dte2   = parseInt(data.days_to_maturity || 0);
      var T2     = Math.max(dte2, 1) / OptionMath.DIAS_ANO;
      var r2     = taxaJuros / 100;
      var flag2  = String(data.category || data.type || 'CALL').toUpperCase().indexOf('PUT') >= 0 ? 'p' : 'c';
      var mktPrc = parseFloat(data.close > 0 ? data.close : (data.bid || 0));
      var ivBS   = ivAtivo;
      if (!ivBS && mktPrc > 0.01 && S2 > 0 && K2 > 0) {
        ivBS = OptionMath.estimateIV(S2, K2, T2, r2, mktPrc, flag2);
      }
      if (ivBS > 0 && S2 > 0 && K2 > 0) {
        var gk = OptionMath.calculate(S2, K2, T2, r2, ivBS, flag2);
        delta = parseFloat((gk.delta || 0).toFixed(4));
      }
    } catch (eDelta) { /* silencioso */ }

    return {
      success: true,
      data: {
        symbol:      data.symbol      || cleanTicker,
        category:    String(data.category || data.type || 'CALL').toUpperCase(),
        strike:      parseFloat(data.strike            || 0),
        premioAtual: parseFloat(data.close > 0 ? data.close : (data.bid || 0)),
        spotPrice:   parseFloat(data.spot_price || data.spot || 0),
        dte:         parseInt(data.days_to_maturity    || 0),
        expiry:      expiry,
        taxaJuros:   taxaJuros,
        ivAtivo:     ivAtivo,
        delta:       delta            // null se IV indisponivel
      }
    };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// Converte "dd/MM/yyyy HH:mm[:ss]" (coluna Data do print) para ISO
// "yyyy-MM-ddTHH:mm:ss" — mesmo formato que o Ctrl+V envia, para o Sheets
// interpretar como data/hora e o Sanitizador preservar a hora. Sem match,
// devolve o fallback (hoje em ISO).
function _nectonDataParaISO(dataStr, fallbackISO) {
  var m = String(dataStr || '').trim().match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})(?::(\d{2}))?$/);
  if (m) return m[3] + '-' + m[2] + '-' + m[1] + 'T' + m[4] + ':' + m[5] + ':' + (m[6] || '00');
  return fallbackISO;
}

// ─────────────────────────────────────────────────────────────────────────────
// parsearImagemOrdens — Claude Vision: extrai ordens de um print da corretora
// Recebe base64 puro (sem prefixo data:) e o mime type da imagem.
// Retorna { success, linhas: [[14 campos], ...], total }
// ─────────────────────────────────────────────────────────────────────────────
function parsearImagemOrdens(base64, mimeType) {
  try {
    var claudeKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
    if (!claudeKey) return { success: false, error: 'CLAUDE_API_KEY não configurada nas Script Properties' };

    var prompt = [
      'Analise esta imagem de ordens de opções de uma corretora brasileira (Necton Home Broker).',
      'Extraia TODAS as ordens visíveis e retorne SOMENTE um JSON válido, sem markdown, sem texto extra:',
      '{"ordens":[{"option_ticker":"VALES795","status":"EXECUTADA","side":"V","quantity":1000,"entry_price":1.34,"last_price":1.45,"order_date":"05/06/2026 14:14"}]}',
      '',
      'Regras de extração:',
      '- option_ticker: código exato da opção (ex: VALES795, PRIOR635, CSNAF702)',
      '- status: normalize para exatamente EXECUTADA, CANCELADA ou EXPIRADA',
      '- side: "V" para venda (badge vermelho com V), "C" para compra (badge azul com C)',
      '- quantity: número inteiro sem separador de milhar (ex: 1.000 ou 1,000 → 1000)',
      '- entry_price: campo "Preço médio" (P. médio) como decimal com ponto (ex: R$ 1,3400 → 1.34). Se não houver use "Preço". Se nenhum, use null',
      '- last_price: campo "Último Preço" (Último) como decimal com ponto. Se não houver, repita o entry_price',
      '- order_date: data e hora da coluna "Data" EXATAMENTE no formato "dd/MM/yyyy HH:mm" (ex: 05/06/2026 14:14). Se não houver, omita o campo',
      '- Inclua TODAS as ordens visíveis, independente do status',
      '- Retorne APENAS o JSON, sem nenhum texto adicional'
    ].join('\n');

    var payload = {
      model: 'claude-sonnet-4-6',
      max_tokens: 4096,
      messages: [{ role: 'user', content: [
        { type: 'text', text: prompt },
        { type: 'image', source: { type: 'base64', media_type: mimeType || 'image/jpeg', data: base64 } }
      ]}]
    };

    var response = null;
    var fetchOpts = {
      method: 'post',
      headers: { 'x-api-key': claudeKey, 'anthropic-version': '2023-06-01', 'content-type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    for (var attempt = 0; attempt < 3; attempt++) {
      try {
        response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', fetchOpts);
        var rc = response.getResponseCode();
        if (rc !== 429 && rc < 500) break; // sucesso ou erro não-retentável (4xx exceto 429)
      } catch (eFetch) {
        if (attempt === 2) throw new Error('Falha de conexão com Claude API: ' + eFetch.message);
      }
      if (attempt < 2) Utilities.sleep(Math.pow(2, attempt + 1) * 1000); // 2s, 4s — sem dormir após a última
    }
    if (!response) return { success: false, error: 'Sem resposta da Claude API após 3 tentativas.' };

    var code = response.getResponseCode();
    if (code !== 200) return { success: false, error: 'Erro na API Claude (HTTP ' + code + '): ' + response.getContentText().substring(0, 200) };

    var respJson = JSON.parse(response.getContentText());
    var texto = (respJson.content && respJson.content[0] && respJson.content[0].text) ? respJson.content[0].text.trim() : '';
    texto = texto.replace(/^```json\s*/i, '').replace(/^```\s*/, '').replace(/\s*```$/, '');
    // Robustez: isola do primeiro { ao último } caso o modelo adicione texto extra
    var ini = texto.indexOf('{'), fim = texto.lastIndexOf('}');
    if (ini >= 0 && fim > ini) texto = texto.substring(ini, fim + 1);

    var parsed;
    try { parsed = JSON.parse(texto); }
    catch (eParse) { return { success: false, error: 'JSON inválido do Claude (possível resposta truncada). Tente um print com menos ordens.' }; }
    var ordens = parsed.ordens || [];
    if (!Array.isArray(ordens)) return { success: false, error: 'Resposta inesperada: sem array de ordens' };

    // Fallback de data: hoje em ISO (mesmo formato do Ctrl+V) para o Sanitizador preservar a hora
    var hojeISO = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd'T'HH:mm:ss");
    var linhas = ordens.map(function(o) {
      // Tolerante a vírgula decimal caso o modelo desobedeça o formato ("1,34" → 1.34)
      var parsePreco = function(v) {
        if (v === null || v === undefined || v === '') return '';
        var n = parseFloat(String(v).replace(',', '.'));
        return isNaN(n) ? '' : n;
      };
      var qty    = parseInt(o.quantity) || 0;
      var price  = parsePreco(o.entry_price);
      var last   = (o.last_price !== null && o.last_price !== undefined && o.last_price !== '') ? parsePreco(o.last_price) : price;
      // Title case: "Executada" (1ª maiúscula, resto minúsculo)
      var statusRaw = String(o.status || 'Executada');
      var status = statusRaw.charAt(0).toUpperCase() + statusRaw.slice(1).toLowerCase();
      var side   = String(o.side   || '').toUpperCase();
      var orderDate = _nectonDataParaISO(o.order_date, hojeISO); // data REAL do print
      return [
        String(o.option_ticker || '').toUpperCase(), // [0]  ATIVO
        status,                                        // [1]  STATUS
        'Limitada',                                    // [2]  ORDER_TYPE (idêntico ao Ctrl+V)
        side,                                          // [3]  SIDE (V / C)
        qty, qty, qty,                                 // [4-6] QTY_OFFER / QTY_DISPLAY / QUANTITY
        status.toUpperCase() === 'EXECUTADA' ? 0 : qty, // [7]  QTY_REMAINING
        price, '', '', price, last,                    // [8-12] LIMIT/DISP/ENTRY/LAST (Último separado)
        orderDate                                      // [13] ORDER_DATE (data/hora real, ISO)
      ];
    });

    SysLogger.log('NectonVision', 'SUCESSO', ordens.length + ' ordens extraídas via Claude Vision');
    SysLogger.flush();
    return { success: true, linhas: linhas, total: linhas.length };

  } catch(e) {
    SysLogger.log('NectonVision', 'ERRO', e.message);
    SysLogger.flush();
    return { success: false, error: e.message };
  }
}

