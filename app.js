/* ===== REBALANCE DASHBOARD - APP.JS ===== */
/* Single-page app with hash routing, XLSX parsing, interactive dashboard */
/* Supports BOTH MSCI and S&P rebalance sheet formats */

// ============================================================
// GLOBAL STATE (in-memory only)
// ============================================================
const STATE = {
  format: null,     // 'msci' or 'sp'
  overall: [],      // Main data from Overall sheet
  summary: {},      // Parsed summary data
  topNames: {},     // 4 lists from TopNames
  comparison: [],   // Comparison sheet data
  effectiveDate: '',
  title: '',        // Rebalance title
  allTickers: [],   // For autocomplete
  mergeNote: '',    // e.g. dual S&P file merge description
  // Filters / sort / pagination
  allNamesFilter: { search: '', country: 'all', industry: 'all', direction: 'all', index: 'all' },
  allNamesSort: { col: null, dir: 'asc' },
  allNamesPage: 1,
  compFilter: { search: '', direction: 'all' },
  compSort: { col: null, dir: 'asc' },
  compPage: 1,
  vlookupSymbols: [],
  charts: {}  // Chart.js instances for cleanup
};

const PER_PAGE = 50;

// ============================================================
// S&P CHANGE → EVENT TYPE MAPPING
// ============================================================
const SP_CHANGE_MAP = {
  'addition': 'ADDITION',
  'deletion': 'DELETION',
  'share increase': 'INCREASED',
  'share decrease': 'DECREASED',
  'mid to spx': 'MID\u2192SPX',
  'spx to sml': 'SPX\u2192SML',
  'mid to sml': 'MID\u2192SML',
  'sml to mid': 'SML\u2192MID',
  'reinvestment': 'REINVEST',
  'funding': 'FUNDING',
  'float decrease': 'FLOAT\u2193'
};

// ============================================================
// UTILITY FUNCTIONS
// ============================================================
function fmt(n, dec = 2) {
  if (n == null || isNaN(n)) return '\u2014';
  const abs = Math.abs(n);
  const s = abs.toLocaleString('en-US', { minimumFractionDigits: dec, maximumFractionDigits: dec });
  return n < 0 ? '-' + s : s;
}

function fmtMM(n) {
  if (n == null || isNaN(n)) return '\u2014';
  const sign = n < 0 ? '-' : '';
  return sign + '$' + fmt(Math.abs(n), 1) + 'MM';
}

function fmtPct(n) {
  if (n == null || isNaN(n)) return '\u2014';
  return fmt(n * 100, 2) + '%';
}

function fmtInt(n) {
  if (n == null || isNaN(n)) return '\u2014';
  return Math.round(n).toLocaleString('en-US');
}

function fmtPrice(n) {
  if (n == null || isNaN(n)) return '\u2014';
  return '$' + fmt(n, 2);
}

function pn(v) { return parseFloat(v) || 0; }

function safeStr(v) { return (v == null ? '' : String(v)).trim(); }

function dirClass(val) {
  if (val > 0) return 'green-text';
  if (val < 0) return 'red-text';
  return '';
}

function eventBadge(eventType) {
  if (!eventType) return '<span class="event-badge">\u2014</span>';
  const cls = {
    // MSCI event types
    'INCREASED': 'event-increased',
    'DECREASED': 'event-decreased',
    'BUY\u2192SELL': 'event-dir-change',
    'SELL\u2192BUY': 'event-dir-change',
    'NO CHANGE': 'event-no-change',
    'NEW': 'event-new',
    // S&P event types
    'ADDITION': 'event-new',
    'DELETION': 'event-decreased',
    'MID\u2192SPX': 'event-increased',
    'SML\u2192MID': 'event-increased',
    'SPX\u2192SML': 'event-dir-change',
    'MID\u2192SML': 'event-dir-change',
    'REINVEST': 'event-no-change',
    'FUNDING': 'event-no-change',
    'FLOAT\u2193': 'event-decreased'
  }[eventType] || 'event-no-change';
  return `<span class="event-badge ${cls}">${eventType}</span>`;
}

function destroyChart(key) {
  if (STATE.charts[key]) {
    STATE.charts[key].destroy();
    delete STATE.charts[key];
  }
}

// ============================================================
// FORMAT DETECTION
// ============================================================
function detectFormat(wb) {
  const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes('overall')) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  for (let i = 0; i < Math.min(20, json.length); i++) {
    const row = json[i];
    if (!row) continue;
    const rowStr = row.map(c => String(c).toUpperCase()).join('|');
    if (rowStr.includes('SEDOL')) return 'msci';
    if (rowStr.includes('CUSIP') || rowStr.includes('|CHANGE') || rowStr.includes('|INDEX')) return 'sp';
  }

  // S&P client net-flows workbooks: TICKER + ISSUER, no SEDOL/CUSIP (BasketAnalysis is not MSCI-only)
  for (let i = 4; i < Math.min(18, json.length); i++) {
    const row = json[i];
    if (!row) continue;
    const upper = row.map(c => String(c).toUpperCase());
    const hasTicker = upper.some(c => c.includes('TICKER'));
    const hasIssuer = upper.some(c => c.includes('ISSUER'));
    const hasCusip = upper.some(c => c.includes('CUSIP'));
    const hasSedol = upper.some(c => c.includes('SEDOL'));
    if (hasTicker && hasIssuer && !hasCusip && !hasSedol) return 'sp';
  }

  if (wb.SheetNames.some(n => n.toLowerCase().includes('comparison'))) return 'msci';
  return 'sp';
}

// ============================================================
// COLUMN DEFINITIONS (format-aware)
// ============================================================
function getAllNamesCols() {
  if (STATE.format === 'sp') {
    return [
      { key: 'cleanTicker', label: 'Ticker' },
      { key: 'issuer', label: 'Name' },
      { key: 'index', label: 'Index' },
      { key: 'industry', label: 'Industry' },
      { key: 'eventType', label: 'Event' },
      { key: 'price', label: 'Price', num: true },
      { key: 'avgVol', label: '20D Avg Vol', num: true },
      { key: 'netShares', label: 'Net Shares', num: true },
      { key: 'netValue', label: 'Net Value ($MM)', num: true },
      { key: 'netLiq', label: 'Net Liq %', num: true },
      { key: 'absLiq', label: 'Abs Liq %', num: true }
    ];
  }
  return [
    { key: 'cleanTicker', label: 'Ticker' },
    { key: 'issuer', label: 'Issuer' },
    { key: 'country', label: 'Country' },
    { key: 'industry', label: 'Industry' },
    { key: 'eventType', label: 'Event' },
    { key: 'price', label: 'Price', num: true },
    { key: 'avgVol', label: '20D Avg Vol', num: true },
    { key: 'netShares', label: 'Net Shares', num: true },
    { key: 'netValue', label: 'Net Value ($MM)', num: true },
    { key: 'netLiq', label: 'Net Liq %', num: true },
    { key: 'grossShares', label: 'Gross Shares', num: true },
    { key: 'grossValue', label: 'Gross Value ($MM)', num: true },
    { key: 'grossLiq', label: 'Gross Liq %', num: true }
  ];
}

// ============================================================
// FILE UPLOAD & PARSING
// ============================================================
function readFileAsArrayBuffer(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(new Uint8Array(reader.result));
    reader.onerror = () => reject(reader.error);
    reader.readAsArrayBuffer(file);
  });
}

document.addEventListener('DOMContentLoaded', () => {
  const zone = document.getElementById('upload-zone');
  const input = document.getElementById('file-input');
  const overlay = document.getElementById('loading-overlay');
  const progress = document.getElementById('loading-progress');
  const loadingFormat = document.getElementById('loading-format');

  zone.addEventListener('click', () => input.click());
  zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
  zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('dragover');
    if (e.dataTransfer.files.length) handleFiles(e.dataTransfer.files);
  });
  input.addEventListener('change', () => { if (input.files.length) handleFiles(input.files); });

  window.handleFile = function handleFile(file) {
    handleFiles(file ? [file] : []);
  };

  async function handleFiles(fileList) {
    let files = Array.from(fileList || []).filter(f => f.name.match(/\.xlsx?$/i));
    if (!files.length) {
      alert('Please choose one or two .xlsx files');
      return;
    }
    if (files.length > 2) {
      alert('More than two files selected; using the first two.');
      files = files.slice(0, 2);
    }

    overlay.classList.add('active');
    if (loadingFormat) loadingFormat.textContent = '';

    try {
      if (files.length === 1) {
        progress.textContent = 'Reading file...';
        const data = await readFileAsArrayBuffer(files[0]);
        progress.textContent = 'Parsing workbook...';
        const wb = XLSX.read(data, { type: 'array', cellDates: true });

        progress.textContent = 'Detecting format...';
        STATE.format = detectFormat(wb);
        STATE.mergeNote = '';
        if (loadingFormat) loadingFormat.textContent = STATE.format === 'sp' ? 'S&P' : 'MSCI';

        progress.textContent = 'Extracting Overall sheet...';
        parseOverall(wb);

        progress.textContent = 'Extracting Summary...';
        parseSummary(wb);

        progress.textContent = 'Extracting Top Names...';
        parseTopNames(wb);

        if (STATE.format === 'msci') {
          progress.textContent = 'Extracting Comparison...';
          parseComparison(wb);
        }
      } else {
        progress.textContent = 'Reading two workbooks...';
        const [d1, d2] = await Promise.all([
          readFileAsArrayBuffer(files[0]),
          readFileAsArrayBuffer(files[1])
        ]);
        progress.textContent = 'Parsing workbooks...';
        const wb1 = XLSX.read(d1, { type: 'array', cellDates: true });
        const wb2 = XLSX.read(d2, { type: 'array', cellDates: true });

        if (loadingFormat) loadingFormat.textContent = 'S&P · dual workbook';
        progress.textContent = 'Merging S&P data...';
        if (!loadDualWorkbooks(wb1, wb2)) {
          overlay.classList.remove('active');
          return;
        }
        STATE.comparison = [];
      }

      progress.textContent = 'Building dashboard...';
      setTimeout(() => {
        initDashboard();
        overlay.classList.remove('active');
        document.getElementById('upload-screen').style.display = 'none';
        document.getElementById('dashboard').classList.add('active');
        handleHashChange();
      }, 100);
    } catch (err) {
      console.error(err);
      alert('Error parsing file: ' + err.message);
      overlay.classList.remove('active');
    }
  }

  window.handleFiles = handleFiles;

  // Hash routing
  window.addEventListener('hashchange', handleHashChange);

  // Key modal
  document.getElementById('open-key-modal').addEventListener('click', () => {
    document.getElementById('key-modal').classList.add('show');
  });
  document.getElementById('close-key-modal').addEventListener('click', () => {
    document.getElementById('key-modal').classList.remove('show');
  });
  document.getElementById('key-modal').addEventListener('click', (e) => {
    if (e.target === e.currentTarget) document.getElementById('key-modal').classList.remove('show');
  });
});

// ============================================================
// PARSE OVERALL SHEET
// ============================================================
function parseOverall(wb) {
  const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes('overall')) || wb.SheetNames[0];
  const ws = wb.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  if (STATE.format === 'sp') {
    parseOverallSP(json);
  } else {
    parseOverallMSCI(json);
  }
}

function parseOverallMSCI(json) {
  // Find header row (row 10 = index 9)
  let headerIdx = -1;
  for (let i = 7; i < Math.min(15, json.length); i++) {
    const row = json[i];
    if (row && row.some(c => String(c).toUpperCase().includes('SEDOL'))) {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) headerIdx = 9; // fallback

  const headers = json[headerIdx].map(h => safeStr(h).toUpperCase());

  function findCol(keywords) {
    return headers.findIndex(h => keywords.every(k => h.includes(k)));
  }

  const colMap = {
    sedol: findCol(['SEDOL']),
    ticker: findCol(['PRIMARY', 'TICKER']) >= 0 ? findCol(['PRIMARY', 'TICKER']) : findCol(['TICKER']),
    issuer: findCol(['ISSUER']),
    country: findCol(['COUNTRY']),
    avgVol: findCol(['20D', 'AVG']) >= 0 ? findCol(['20D', 'AVG']) : findCol(['AVG', 'VOL']),
    price: findCol(['USD', 'PRICE']) >= 0 ? findCol(['USD', 'PRICE']) : findCol(['PRICE']),
    netShares: findCol(['NET', 'SHARES', 'TRADE']),
    netValue: findCol(['NET', 'TRADE', 'VALUE']),
    netLiq: findCol(['NET', 'LIQUIDITY']),
    grossShares: findCol(['GROSS', 'CROSSING', 'SHARES']) >= 0 ? findCol(['GROSS', 'CROSSING', 'SHARES']) : findCol(['GROSS', 'SHARES']),
    grossValue: findCol(['GROSS', 'CROSSING', 'TRADE', 'VALUE']) >= 0 ? findCol(['GROSS', 'CROSSING', 'TRADE', 'VALUE']) : findCol(['GROSS', 'TRADE', 'VALUE']),
    grossLiq: findCol(['GROSS', 'CROSSING', 'LIQUIDITY']) >= 0 ? findCol(['GROSS', 'CROSSING', 'LIQUIDITY']) : findCol(['GROSS', 'LIQUIDITY']),
    industry: findCol(['INDUSTRY']),
    effectiveDate: findCol(['EFFECTIVE'])
  };

  STATE.overall = [];
  STATE.allTickers = [];

  for (let i = headerIdx + 1; i < json.length; i++) {
    const row = json[i];
    if (!row || !row[colMap.sedol]) continue;

    const sedol = safeStr(row[colMap.sedol]);
    if (!sedol || sedol === '' || sedol.toUpperCase() === 'SEDOL') continue;

    const ticker = safeStr(row[colMap.ticker]);
    const issuer = safeStr(row[colMap.issuer]);
    const country = safeStr(row[colMap.country]);
    const netVal = pn(row[colMap.netValue]);

    let cleanTicker = ticker.split(' ')[0] || ticker;
    const exchange = ticker.split(' ').slice(1).join(' ') || '';

    const item = {
      sedol, ticker, cleanTicker, exchange, issuer, country,
      index: '',
      cusip: '',
      avgVol: pn(row[colMap.avgVol]),
      price: pn(row[colMap.price]),
      netShares: pn(row[colMap.netShares]),
      netValue: netVal,
      netLiq: pn(row[colMap.netLiq]),
      grossShares: pn(row[colMap.grossShares]),
      grossValue: pn(row[colMap.grossValue]),
      grossLiq: pn(row[colMap.grossLiq]),
      absLiq: 0,
      industry: safeStr(row[colMap.industry]),
      direction: netVal >= 0 ? 'BUY' : 'SELL'
    };

    // Effective date
    if (!STATE.effectiveDate && colMap.effectiveDate >= 0 && row[colMap.effectiveDate]) {
      const d = row[colMap.effectiveDate];
      if (d instanceof Date) {
        STATE.effectiveDate = d.toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' });
      } else {
        STATE.effectiveDate = safeStr(d);
      }
    }

    STATE.overall.push(item);
    STATE.allTickers.push({ ticker: cleanTicker, fullTicker: ticker, issuer, country });
  }
}

/**
 * Parse S&P Overall into a bucket (used for single file or dual-file merge).
 * variant: 'standard' (CUSIP + CHANGE) or 'netflows' (ISSUER/COUNTRY, no event column).
 */
function parseOverallSPInto(json, bucket) {
  bucket.overall = [];
  bucket.allTickers = [];

  for (let r = 1; r < Math.min(10, json.length); r++) {
    const row = json[r];
    if (!row) continue;
    const a = safeStr(row[1]) || safeStr(row[0]);
    if (!a) continue;
    if (!bucket.title && (/S&P|Estimated|Rebalance|Quarterly|Share\s*Changes/i.test(a) || a.length > 24)) {
      bucket.title = a;
    }
    if (/effective/i.test(a)) {
      const m = a.match(/(?:on|at)\s+(.+)/i);
      bucket.effectiveDate = m ? m[1].trim() : a;
    }
  }

  let headerIdx = -1;
  let variant = 'standard';
  for (let i = 4; i < Math.min(16, json.length); i++) {
    const row = json[i];
    if (!row) continue;
    const rowUpper = row.map(c => String(c).toUpperCase());
    const hasT = rowUpper.some(c => c.includes('TICKER'));
    const hasC = rowUpper.some(c => c.includes('CUSIP'));
    const hasIss = rowUpper.some(c => c.includes('ISSUER'));
    if (hasT && hasC) {
      headerIdx = i;
      variant = 'standard';
      break;
    }
    if (hasT && hasIss && !hasC) {
      headerIdx = i;
      variant = 'netflows';
      break;
    }
  }
  if (headerIdx === -1) headerIdx = 6;
  bucket.variant = variant;

  const headers = json[headerIdx].map(h => safeStr(h).toUpperCase());

  function findCol(keywords) {
    return headers.findIndex(h => keywords.every(k => h.includes(k)));
  }

  if (variant === 'standard') {
    const colMap = {
      ticker: findCol(['TICKER']),
      cusip: findCol(['CUSIP']),
      name: findCol(['NAME']),
      industry: findCol(['INDUSTRY']),
      index: findCol(['INDEX']),
      change: findCol(['CHANGE']),
      price: findCol(['PRICE']),
      netValue: findCol(['NET', 'VALUE']),
      netShares: findCol(['NET', 'SHARES']),
      avgVol: findCol(['20D']) >= 0 ? findCol(['20D']) : findCol(['AVG', 'VOL']),
      liquidity: findCol(['LIQUIDITY']),
      absLiq: findCol(['ABS'])
    };

    for (let i = headerIdx + 1; i < json.length; i++) {
      const row = json[i];
      if (!row) continue;

      const ticker = safeStr(row[colMap.ticker]);
      if (!ticker || ticker.toUpperCase() === 'TICKER') continue;

      const cusip = safeStr(row[colMap.cusip]);
      const name = safeStr(row[colMap.name >= 0 ? colMap.name : colMap.ticker + 1]);
      const netVal = pn(row[colMap.netValue]);
      const changeRaw = colMap.change >= 0 ? safeStr(row[colMap.change]) : '';
      const eventType = changeRaw
        ? (SP_CHANGE_MAP[changeRaw.toLowerCase()] || changeRaw.toUpperCase())
        : '';

      const item = {
        sedol: cusip,
        ticker,
        cleanTicker: ticker,
        exchange: '',
        issuer: name,
        country: '',
        index: safeStr(row[colMap.index]),
        cusip,
        avgVol: pn(row[colMap.avgVol]),
        price: pn(row[colMap.price]),
        netShares: pn(row[colMap.netShares]),
        netValue: netVal,
        netLiq: pn(row[colMap.liquidity]),
        grossShares: 0,
        grossValue: 0,
        grossLiq: 0,
        absLiq: pn(row[colMap.absLiq]),
        industry: safeStr(row[colMap.industry]),
        direction: netVal >= 0 ? 'BUY' : 'SELL',
        eventType
      };

      bucket.overall.push(item);
      bucket.allTickers.push({ ticker, fullTicker: ticker, issuer: name, country: '' });
    }
    return;
  }

  // Net-flows / client layout (TICKER + ISSUER, optional gross columns)
  const colMap = {
    ticker: findCol(['TICKER']),
    issuer: findCol(['ISSUER']),
    country: findCol(['COUNTRY']),
    avgVol: findCol(['20D', 'AVG']) >= 0 ? findCol(['20D', 'AVG']) : findCol(['AVG', 'VOL']),
    price: findCol(['USD', 'PRICE']) >= 0 ? findCol(['USD', 'PRICE']) : findCol(['PRICE']),
    netShares: findCol(['NET', 'SHARES']),
    netValue: findCol(['NET', 'TRADE', 'VALUE']) >= 0 ? findCol(['NET', 'TRADE', 'VALUE']) : findCol(['NET', 'VALUE']),
    netLiq: findCol(['NET', 'LIQUIDITY']) >= 0 ? findCol(['NET', 'LIQUIDITY']) : findCol(['LIQUIDITY']),
    grossShares: findCol(['GROSS', 'CROSSING', 'SHARES']) >= 0 ? findCol(['GROSS', 'CROSSING', 'SHARES']) : findCol(['GROSS', 'SHARES']),
    grossValue: findCol(['GROSS', 'CROSSING', 'TRADE', 'VALUE']) >= 0 ? findCol(['GROSS', 'CROSSING', 'TRADE', 'VALUE']) : findCol(['GROSS', 'TRADE', 'VALUE']),
    grossLiq: findCol(['GROSS', 'CROSSING', 'LIQUIDITY']) >= 0 ? findCol(['GROSS', 'CROSSING', 'LIQUIDITY']) : findCol(['GROSS', 'LIQUIDITY']),
    industry: findCol(['INDUSTRY'])
  };

  for (let i = headerIdx + 1; i < json.length; i++) {
    const row = json[i];
    if (!row) continue;

    const ticker = safeStr(row[colMap.ticker]);
    if (!ticker || ticker.toUpperCase() === 'TICKER') continue;

    const issuer = colMap.issuer >= 0 ? safeStr(row[colMap.issuer]) : '';
    const country = colMap.country >= 0 ? safeStr(row[colMap.country]) : '';
    const netVal = pn(row[colMap.netValue]);

    let cleanTicker = ticker.split(' ')[0] || ticker;
    const exchange = ticker.split(' ').slice(1).join(' ') || '';

    const item = {
      sedol: '',
      ticker,
      cleanTicker,
      exchange,
      issuer,
      country,
      index: '',
      cusip: '',
      avgVol: pn(row[colMap.avgVol]),
      price: pn(row[colMap.price]),
      netShares: pn(row[colMap.netShares]),
      netValue: netVal,
      netLiq: pn(row[colMap.netLiq]),
      grossShares: colMap.grossShares >= 0 ? pn(row[colMap.grossShares]) : 0,
      grossValue: colMap.grossValue >= 0 ? pn(row[colMap.grossValue]) : 0,
      grossLiq: colMap.grossLiq >= 0 ? pn(row[colMap.grossLiq]) : 0,
      absLiq: Math.abs(pn(row[colMap.netLiq])),
      industry: colMap.industry >= 0 ? safeStr(row[colMap.industry]) : '',
      direction: netVal >= 0 ? 'BUY' : 'SELL',
      eventType: ''
    };

    bucket.overall.push(item);
    bucket.allTickers.push({ ticker: cleanTicker, fullTicker: ticker, issuer, country });
  }
}

function parseOverallSP(json) {
  const bucket = { overall: [], allTickers: [], title: '', effectiveDate: '', variant: 'standard' };
  parseOverallSPInto(json, bucket);
  STATE.overall = bucket.overall;
  STATE.allTickers = bucket.allTickers;
  if (bucket.title) STATE.title = bucket.title;
  if (bucket.effectiveDate) STATE.effectiveDate = bucket.effectiveDate;
}

function getOverallJson(wb) {
  const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes('overall')) || wb.SheetNames[0];
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: '' });
}

function spOverallBundleFromWorkbook(wb) {
  const bucket = { overall: [], allTickers: [], title: '', effectiveDate: '', variant: 'standard' };
  parseOverallSPInto(getOverallJson(wb), bucket);
  return bucket;
}

/** Merge S&P net-flows list (larger) with standard share-change file (event + CUSIP). Order-independent. */
function mergeSpDualBundles(flowsBundle, changeBundle) {
  const byTicker = new Map();
  for (const row of changeBundle.overall) {
    const k = row.cleanTicker.toUpperCase();
    if (!byTicker.has(k)) byTicker.set(k, row);
    if (row.cusip) byTicker.set('_' + row.cusip, row);
  }

  for (const row of flowsBundle.overall) {
    const k = row.cleanTicker.toUpperCase();
    const src = byTicker.get(k) || (row.cusip && byTicker.get('_' + row.cusip));
    if (!src || !src.eventType) continue;
    row.eventType = src.eventType;
    if (!row.index && src.index) row.index = src.index;
    if (!row.cusip && src.cusip) row.cusip = src.cusip;
  }

  STATE.overall = flowsBundle.overall;
  STATE.allTickers = flowsBundle.overall.map(d => ({
    ticker: d.cleanTicker,
    fullTicker: d.ticker,
    issuer: d.issuer,
    country: d.country || ''
  }));
  STATE.title = flowsBundle.title || changeBundle.title || STATE.title;
  STATE.effectiveDate = flowsBundle.effectiveDate || changeBundle.effectiveDate || STATE.effectiveDate;
}

function scoreSpSummary(summary) {
  const stats = summary.spIndexStats || {};
  const nStats = Object.keys(stats).length;
  const ind = summary.industryNetFlows || [];
  return nStats * 100 + ind.length;
}

function pickRicherSpSummary(wbA, wbB) {
  parseSummary(wbA);
  const sA = JSON.parse(JSON.stringify(STATE.summary));
  parseSummary(wbB);
  const sB = JSON.parse(JSON.stringify(STATE.summary));
  STATE.summary = scoreSpSummary(sA) >= scoreSpSummary(sB) ? sA : sB;
}

function pickRicherTopNames(wbA, wbB) {
  parseTopNames(wbA);
  const tA = JSON.parse(JSON.stringify(STATE.topNames));
  const countA = ['illiquidBuys', 'illiquidSells', 'largestBuys', 'largestSells'].reduce((n, k) => n + (tA[k] || []).length, 0);
  parseTopNames(wbB);
  const tB = JSON.parse(JSON.stringify(STATE.topNames));
  const countB = ['illiquidBuys', 'illiquidSells', 'largestBuys', 'largestSells'].reduce((n, k) => n + (tB[k] || []).length, 0);
  STATE.topNames = countA >= countB ? tA : tB;
}

function loadDualWorkbooks(wb1, wb2) {
  const f1 = detectFormat(wb1);
  const f2 = detectFormat(wb2);
  if (f1 !== 'sp' || f2 !== 'sp') {
    alert('Dual upload is for two S&P-format workbooks (net flows + share-change). For MSCI, use one file.');
    return false;
  }

  STATE.format = 'sp';
  STATE.mergeNote = '';

  const b1 = spOverallBundleFromWorkbook(wb1);
  const b2 = spOverallBundleFromWorkbook(wb2);

  let flows;
  let change;
  if (b1.variant === 'netflows' && b2.variant === 'standard') {
    flows = b1;
    change = b2;
    STATE.mergeNote = 'Net-flows list (larger universe) plus share-change file: events, index, and CUSIP filled in by ticker match.';
  } else if (b2.variant === 'netflows' && b1.variant === 'standard') {
    flows = b2;
    change = b1;
    STATE.mergeNote = 'Net-flows list (larger universe) plus share-change file: events, index, and CUSIP filled in by ticker match.';
  } else if (b1.variant === 'standard' && b2.variant === 'standard') {
    const larger = b1.overall.length >= b2.overall.length ? b1 : b2;
    const smaller = larger === b1 ? b2 : b1;
    flows = larger;
    change = smaller;
    STATE.mergeNote = 'Both files use the share-change layout; events from the smaller file applied where tickers match.';
  } else {
    flows = b1.overall.length >= b2.overall.length ? b1 : b2;
    change = flows === b1 ? b2 : b1;
    STATE.mergeNote = 'Both files look like net-flows exports (no CHANGE column); using the larger name list. Add the share-change XLSX for additions/deletions.';
  }

  mergeSpDualBundles(flows, change);
  pickRicherSpSummary(wb1, wb2);
  pickRicherTopNames(wb1, wb2);
  return true;
}

// ============================================================
// PARSE SUMMARY SHEET
// ============================================================
function parseSummary(wb) {
  const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes('summary'));
  if (!sheetName) return;
  const ws = wb.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  if (STATE.format === 'sp') {
    parseSummarySP(json);
  } else {
    parseSummaryMSCI(json);
  }
}

function parseSummaryMSCI(json) {
  STATE.summary = {
    flowBreakdown: [],
    twoWayTotal: 0,
    beforeCrossing: 0,
    crossingFlow: 0,
    afterCrossing: 0,
    industryNetFlows: [],
    industryTwoWayFlows: []
  };

  for (let i = 0; i < json.length; i++) {
    const row = json[i];
    const label = safeStr(row[0]) || safeStr(row[1]);

    if (label && (label.includes('Two Way Total') || label.includes('Two-Way Total') || label.includes('2 Way Total'))) {
      for (let c = 1; c < row.length; c++) {
        if (pn(row[c]) !== 0) { STATE.summary.twoWayTotal = pn(row[c]); break; }
      }
    }

    if (label && label.includes('Before Crossing') && !label.includes('Total')) {
      for (let c = 1; c < row.length; c++) {
        if (pn(row[c]) !== 0) { STATE.summary.beforeCrossing = pn(row[c]); break; }
      }
    }

    if (label && (label.includes('Crossing Flow') || label.includes('Crossing Est'))) {
      for (let c = 1; c < row.length; c++) {
        if (pn(row[c]) !== 0) { STATE.summary.crossingFlow = pn(row[c]); break; }
      }
    }

    if (label && label.includes('After Crossing')) {
      for (let c = 1; c < row.length; c++) {
        if (pn(row[c]) !== 0) { STATE.summary.afterCrossing = pn(row[c]); break; }
      }
    }
  }

  // Parse index type breakdown (DM, EM, etc.)
  const indexTypes = ['DM', 'EM', 'DM Small Cap', 'EM Small Cap', 'Min Volatility', 'Minimum Volatility',
    'Momentum', 'Quality', 'High Yield', 'High Div', 'Enhanced Value', 'ESG'];

  for (let i = 0; i < json.length; i++) {
    const row = json[i];
    const label = safeStr(row[0]) || safeStr(row[1]);
    for (const idx of indexTypes) {
      if (label && label.toLowerCase().includes(idx.toLowerCase()) && !label.includes('Total') && !label.includes('INDUSTRY')) {
        let val = 0, pct = 0;
        for (let c = 1; c < Math.min(row.length, 8); c++) {
          const v = pn(row[c]);
          if (v !== 0 && val === 0) val = v;
          else if (v !== 0 && v < 1 && v > -1 && pct === 0) pct = v;
          else if (v !== 0 && pct === 0 && val !== 0) pct = v;
        }
        if (val !== 0) {
          STATE.summary.flowBreakdown.push({ name: idx.replace('Minimum Volatility', 'Min Volatility'), value: val, pct: pct });
        }
        break;
      }
    }
  }

  // Parse Industry Net Flows and Two-Way Flows
  let inNetFlows = false, inTwoWayFlows = false;
  for (let i = 0; i < json.length; i++) {
    const row = json[i];
    const label = safeStr(row[0]) || safeStr(row[1]);

    if (label.includes('INDUSTRY NET FLOWS') || label.includes('Industry Net Flows')) {
      inNetFlows = true; inTwoWayFlows = false; continue;
    }
    if (label.includes('INDUSTRY TWO-WAY') || label.includes('Industry Two-Way') || label.includes('INDUSTRY 2-WAY')) {
      inNetFlows = false; inTwoWayFlows = true; continue;
    }

    if (inNetFlows && label && !label.includes('INDUSTRY') && !label.includes('USD') && !label.includes('Industry Group') && !label.includes('Source') && !label.includes('Total')) {
      let val = 0;
      for (let c = 1; c < row.length; c++) {
        if (pn(row[c]) !== 0) { val = pn(row[c]); break; }
      }
      if (val !== 0) {
        STATE.summary.industryNetFlows.push({ name: label, value: val });
      }
    }

    if (inTwoWayFlows && label && !label.includes('INDUSTRY') && !label.includes('USD') && !label.includes('Industry Group') && !label.includes('Source') && !label.includes('Total')) {
      let val = 0;
      for (let c = 1; c < row.length; c++) {
        if (pn(row[c]) !== 0) { val = pn(row[c]); break; }
      }
      if (val !== 0) {
        STATE.summary.industryTwoWayFlows.push({ name: label, value: val });
      }
    }
  }

  if (STATE.summary.beforeCrossing === 0 && STATE.summary.twoWayTotal !== 0) {
    STATE.summary.beforeCrossing = STATE.summary.twoWayTotal;
  }
}

function parseSummarySP(json) {
  STATE.summary = {
    flowBreakdown: [],       // Will hold S&P index buy/sell stats
    twoWayTotal: 0,
    beforeCrossing: 0,
    crossingFlow: 0,
    afterCrossing: 0,
    industryNetFlows: [],
    industryTwoWayFlows: [],
    spIndexStats: {}         // S&P specific: per-index buy/sell/total stats
  };

  // S&P Summary: Industry net flows from rows 9+ (col B=industry, C=S&P500, D=S&P400, E=S&P600, F=NET TOTAL)
  // Look for the "Industry Group" header row
  let industryHeaderRow = -1;
  for (let i = 0; i < Math.min(20, json.length); i++) {
    const row = json[i];
    if (!row) continue;
    const rowStr = row.map(c => safeStr(c).toUpperCase()).join('|');
    if (rowStr.includes('INDUSTRY GROUP') || (rowStr.includes('INDUSTRY') && rowStr.includes('S&P'))) {
      industryHeaderRow = i;
      break;
    }
  }

  if (industryHeaderRow >= 0) {
    // Parse industry net flows
    for (let i = industryHeaderRow + 1; i < json.length; i++) {
      const row = json[i];
      if (!row) continue;
      const label = safeStr(row[1]); // col B = industry name
      if (!label || label.toUpperCase().includes('TOTAL') || label.toUpperCase().includes('INDUSTRY') || label.toUpperCase().includes('SOURCE')) {
        if (label.toUpperCase().includes('TOTAL')) break; // stop at total row
        continue;
      }
      const netTotal = pn(row[5]); // col F = NET TOTAL
      if (netTotal !== 0) {
        STATE.summary.industryNetFlows.push({
          name: label,
          value: netTotal,
          sp500: pn(row[2]),
          sp400: pn(row[3]),
          sp600: pn(row[4])
        });
      }
    }
  }

  // Parse buy/sell/total stats from cols H-K (S&P 500), M-P (S&P 400), R-U (S&P 600)
  // Look for rows with "Buy", "Sell", "Total" labels
  for (let i = 0; i < json.length; i++) {
    const row = json[i];
    if (!row) continue;
    const label = safeStr(row[7]).toLowerCase(); // col H label
    if (!label) continue;

    // S&P 500 stats (cols H-K)
    if (label.includes('buy') || label.includes('sell') || label.includes('total')) {
      if (!STATE.summary.spIndexStats['S&P 500']) {
        STATE.summary.spIndexStats['S&P 500'] = {};
      }
      if (!STATE.summary.spIndexStats['S&P 400']) {
        STATE.summary.spIndexStats['S&P 400'] = {};
      }
      if (!STATE.summary.spIndexStats['S&P 600']) {
        STATE.summary.spIndexStats['S&P 600'] = {};
      }

      const key = label.includes('buy') ? 'buys' : label.includes('sell') ? 'sells' : 'total';

      // S&P 500: cols H(7)=label, I(8)=names, J(9)=value, K(10)=something
      STATE.summary.spIndexStats['S&P 500'][key] = {
        count: pn(row[8]),
        value: pn(row[9])
      };
      // S&P 400: cols M(12)=label, N(13)=names, O(14)=value, P(15)=something
      STATE.summary.spIndexStats['S&P 400'][key] = {
        count: pn(row[13]),
        value: pn(row[14])
      };
      // S&P 600: cols R(17)=label, S(18)=names, T(19)=value, U(20)=something
      STATE.summary.spIndexStats['S&P 600'][key] = {
        count: pn(row[18]),
        value: pn(row[19])
      };
    }
  }
}

// ============================================================
// PARSE TOP NAMES SHEET
// ============================================================
function parseTopNames(wb) {
  const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes('top'));
  if (!sheetName) return;
  const ws = wb.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  STATE.topNames = {
    illiquidBuys: [],
    illiquidSells: [],
    largestBuys: [],
    largestSells: []
  };

  if (STATE.format === 'sp') {
    parseTopNamesSP(json);
  } else {
    parseTopNamesMSCI(json);
  }
}

function parseTopNamesMSCI(json) {
  // Find header rows - look for rows containing "TICKER" multiple times
  let headerRow = -1;
  for (let i = 0; i < Math.min(15, json.length); i++) {
    const row = json[i];
    const tickerCount = row.filter(c => safeStr(c).toUpperCase().includes('TICKER')).length;
    if (tickerCount >= 2) { headerRow = i; break; }
  }

  if (headerRow === -1) {
    for (let i = 0; i < Math.min(15, json.length); i++) {
      const row = json[i];
      if (row.some(c => safeStr(c).toUpperCase().includes('ILLIQUID'))) {
        headerRow = i + 1; break;
      }
    }
  }

  if (headerRow === -1) headerRow = 9;

  const headers = json[headerRow].map(c => safeStr(c).toUpperCase());
  const tickerCols = [];
  headers.forEach((h, idx) => { if (h.includes('TICKER')) tickerCols.push(idx); });

  const tables = [STATE.topNames.illiquidBuys, STATE.topNames.illiquidSells,
                  STATE.topNames.largestBuys, STATE.topNames.largestSells];

  if (tickerCols.length >= 4) {
    for (let t = 0; t < 4 && t < tickerCols.length; t++) {
      const baseCol = tickerCols[t];
      for (let i = headerRow + 1; i < json.length; i++) {
        const row = json[i];
        const ticker = safeStr(row[baseCol]);
        if (!ticker) continue;
        tables[t].push({
          ticker,
          name: safeStr(row[baseCol + 1]),
          country: safeStr(row[baseCol + 2]),
          tradeValue: pn(row[baseCol + 3]),
          liquidity: pn(row[baseCol + 4])
        });
      }
    }
  } else if (tickerCols.length >= 2) {
    for (let t = 0; t < tickerCols.length && t < 4; t++) {
      const baseCol = tickerCols[t];
      for (let i = headerRow + 1; i < json.length; i++) {
        const row = json[i];
        const ticker = safeStr(row[baseCol]);
        if (!ticker) continue;
        tables[t].push({
          ticker,
          name: safeStr(row[baseCol + 1]),
          country: safeStr(row[baseCol + 2]),
          tradeValue: pn(row[baseCol + 3]),
          liquidity: pn(row[baseCol + 4])
        });
      }
    }
  }
}

function parseTopNamesSP(json) {
  // S&P TopNames: Has section dividers ("S&P 500", "S&P 400", "S&P 600") between data groups
  // Header at row 13 (index 12): B=TICKER, C=NAME, D=NET TRADE VALUE, E=LIQUIDITY
  // Tables at column groups B(1), G(6), L(11), Q(16) - 5 cols per group, no COUNTRY

  let headerRow = -1;
  for (let i = 0; i < Math.min(20, json.length); i++) {
    const row = json[i];
    if (!row) continue;
    const tickerCount = row.filter(c => safeStr(c).toUpperCase().includes('TICKER')).length;
    if (tickerCount >= 2) { headerRow = i; break; }
  }

  if (headerRow === -1) headerRow = 12;

  const headers = json[headerRow].map(c => safeStr(c).toUpperCase());
  const tickerCols = [];
  headers.forEach((h, idx) => { if (h.includes('TICKER')) tickerCols.push(idx); });

  const tables = [STATE.topNames.illiquidBuys, STATE.topNames.illiquidSells,
                  STATE.topNames.largestBuys, STATE.topNames.largestSells];

  const numTables = Math.min(4, tickerCols.length);
  for (let t = 0; t < numTables; t++) {
    const baseCol = tickerCols[t];
    for (let i = headerRow + 1; i < json.length; i++) {
      const row = json[i];
      if (!row) continue;
      const ticker = safeStr(row[baseCol]);
      // Skip section dividers like "S&P 500", "S&P 400", "S&P 600"
      if (!ticker || ticker.startsWith('S&P')) continue;
      tables[t].push({
        ticker,
        name: safeStr(row[baseCol + 1]),
        country: '', // No country in S&P format
        tradeValue: pn(row[baseCol + 2]),
        liquidity: pn(row[baseCol + 3])
      });
    }
  }
}

// ============================================================
// PARSE COMPARISON SHEET (MSCI only)
// ============================================================
function parseComparison(wb) {
  const sheetName = wb.SheetNames.find(n => n.toLowerCase().includes('compar'));
  if (!sheetName) return;
  const ws = wb.Sheets[sheetName];
  const json = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

  let headerIdx = -1;
  for (let i = 8; i < Math.min(16, json.length); i++) {
    const row = json[i];
    if (row && row.some(c => safeStr(c).toUpperCase().includes('SEDOL'))) {
      headerIdx = i; break;
    }
  }
  if (headerIdx === -1) headerIdx = 10;

  const headers = json[headerIdx].map(h => safeStr(h).toUpperCase());

  function findCol(keywords) {
    return headers.findIndex(h => keywords.every(k => h.includes(k)));
  }

  const colMap = {
    sedol: findCol(['SEDOL']),
    ticker: findCol(['TICKER']),
    issuer: findCol(['ISSUER']),
    country: findCol(['COUNTRY']),
    dirChange: findCol(['DIRECTION', 'CHANGE']) >= 0 ? findCol(['DIRECTION', 'CHANGE']) : findCol(['NET', 'DIRECTION']),
    pctChange: findCol(['%', 'CHANGE']) >= 0 ? findCol(['%', 'CHANGE']) : findCol(['CHANGE', 'NET', 'SHARES'])
  };

  let dataStartCol = -1;
  for (let c = 0; c < headers.length; c++) {
    if (headers[c].includes('STANDARD') || headers[c].includes('SMALLCAP')) {
      dataStartCol = c;
      break;
    }
  }

  if (dataStartCol === -1) {
    dataStartCol = colMap.pctChange >= 0 ? colMap.pctChange + 1 : 7;
  }

  const d = dataStartCol;

  STATE.comparison = [];
  for (let i = headerIdx + 1; i < json.length; i++) {
    const row = json[i];
    const sedol = safeStr(row[colMap.sedol]);
    if (!sedol) continue;

    const item = {
      sedol,
      ticker: safeStr(row[colMap.ticker]),
      issuer: safeStr(row[colMap.issuer]),
      country: safeStr(row[colMap.country]),
      dirChange: safeStr(row[colMap.dirChange >= 0 ? colMap.dirChange : 4]),
      pctChange: pn(row[colMap.pctChange >= 0 ? colMap.pctChange : 5]),
      netStandard: { shares: pn(row[d]), value: pn(row[d+2]), liq: pn(row[d+4]) },
      netFull: { shares: pn(row[d+1]), value: pn(row[d+3]), liq: pn(row[d+5]) },
      grossStandard: { shares: pn(row[d+6]), value: pn(row[d+8]), liq: pn(row[d+10]) },
      grossFull: { shares: pn(row[d+7]), value: pn(row[d+9]), liq: pn(row[d+11]) }
    };

    STATE.comparison.push(item);
  }
}

// ============================================================
// HASH ROUTING
// ============================================================
function handleHashChange() {
  const hash = (location.hash || '#overview').replace('#', '');
  const validTabs = ['overview', 'allnames', 'topmovers', 'comparison', 'vlookup', 'analytics'];
  const tab = validTabs.includes(hash) ? hash : 'overview';

  document.querySelectorAll('.sidebar-nav a').forEach(a => {
    a.classList.toggle('active', a.dataset.tab === tab);
  });

  document.querySelectorAll('.tab-content').forEach(el => {
    el.classList.toggle('active', el.id === 'tab-' + tab);
  });
}

// ============================================================
// INITIALIZE DASHBOARD
// ============================================================
function initDashboard() {
  // Update sidebar branding based on format
  const sidebarLogo = document.querySelector('.sidebar-logo');
  if (STATE.format === 'sp') {
    sidebarLogo.innerHTML = 'S&P <span>Rebalance</span>';
  } else {
    sidebarLogo.innerHTML = 'MSCI <span>Rebalance</span>';
  }

  document.getElementById('sidebar-date').textContent = STATE.effectiveDate || 'Effective Date: N/A';

  // Update comparison nav item
  const compNavItem = document.querySelector('.sidebar-nav a[data-tab="comparison"]');
  if (STATE.format === 'sp') {
    compNavItem.innerHTML = '<span class="nav-icon">&#8644;</span> Index Breakdown';
  } else {
    compNavItem.innerHTML = '<span class="nav-icon">&#8644;</span> Comparison';
  }

  // Enrich event types (MSCI uses Comparison sheet; S&P already has them from CHANGE column)
  if (STATE.format === 'msci') {
    enrichEventTypes();
  }

  renderOverview();
  renderAllNames();
  renderTopMovers();
  renderComparison();
  renderVLookup();
  renderAnalytics();
}

function enrichEventTypes() {
  // MSCI only: Build a lookup map from comparison data by ticker
  const compMap = {};
  STATE.comparison.forEach(c => {
    compMap[c.ticker.toUpperCase()] = c;
  });

  STATE.overall.forEach(item => {
    const comp = compMap[item.ticker.toUpperCase()];
    if (!comp) {
      item.eventType = 'NEW';
      return;
    }

    const dirChange = (comp.dirChange || '').trim();
    const pctChange = comp.pctChange || 0;

    if (dirChange === 'Buy to Sell' || dirChange === 'Sell to Buy') {
      item.eventType = dirChange === 'Buy to Sell' ? 'BUY\u2192SELL' : 'SELL\u2192BUY';
    } else if (pctChange === 0) {
      item.eventType = 'NO CHANGE';
    } else if (pctChange > 0) {
      item.eventType = 'INCREASED';
    } else {
      item.eventType = 'DECREASED';
    }

    item.pctChangeShares = pctChange;
  });
}

// ============================================================
// TAB: OVERVIEW
// ============================================================
function renderOverview() {
  const data = STATE.overall;
  const summary = STATE.summary;

  const totalNames = data.length;
  const buys = data.filter(d => d.direction === 'BUY');
  const sells = data.filter(d => d.direction === 'SELL');
  const totalNetBuyValue = buys.reduce((s, d) => s + d.netValue, 0);
  const totalNetSellValue = sells.reduce((s, d) => s + d.netValue, 0);

  // Update overview badge
  const badge = document.getElementById('overview-badge');
  badge.textContent = STATE.format === 'sp' ? 'QUARTERLY REVIEW' : 'REBALANCE DAY';

  // Top banner
  const banner = document.getElementById('top-banner');
  if (STATE.format === 'sp') {
    banner.innerHTML = `
      <div class="banner-item">
        <div class="banner-label">Rebalance</div>
        <div class="banner-value cyan">${STATE.title || 'S&P 500/400/600 Quarterly Review'}</div>
      </div>
      <div class="banner-item">
        <div class="banner-label">Effective Date</div>
        <div class="banner-value cyan">${STATE.effectiveDate || 'N/A'}</div>
      </div>
      <div class="banner-item">
        <div class="banner-label">Total Names</div>
        <div class="banner-value">${fmtInt(totalNames)}</div>
      </div>
      <div class="banner-item">
        <div class="banner-label">Buys</div>
        <div class="banner-value" style="color:var(--green)">${fmtInt(buys.length)}</div>
      </div>
      <div class="banner-item">
        <div class="banner-label">Sells</div>
        <div class="banner-value" style="color:var(--red)">${fmtInt(sells.length)}</div>
      </div>
      ${STATE.mergeNote ? `<div class="banner-merge-note">${STATE.mergeNote}</div>` : ''}
    `;
  } else {
    banner.innerHTML = `
      <div class="banner-item">
        <div class="banner-label">Effective Date</div>
        <div class="banner-value cyan">${STATE.effectiveDate || 'N/A'}</div>
      </div>
      <div class="banner-item">
        <div class="banner-label">Total Names</div>
        <div class="banner-value">${fmtInt(totalNames)}</div>
      </div>
      <div class="banner-item">
        <div class="banner-label">Two-Way Flow (Before Crossing)</div>
        <div class="banner-value">${summary.beforeCrossing ? '$' + fmt(summary.beforeCrossing, 0) + ' MM' : fmtMM(buys.reduce((s,d)=>s+Math.abs(d.grossValue),0) + sells.reduce((s,d)=>s+Math.abs(d.grossValue),0))}</div>
      </div>
    `;
  }

  // Summary cards
  const cards = document.getElementById('summary-cards');
  let cardData;
  if (STATE.format === 'sp') {
    cardData = [
      { title: 'Total Names', value: fmtInt(totalNames), cls: 'cyan' },
      { title: 'Net Buys', value: fmtInt(buys.length), sub: fmtMM(totalNetBuyValue), cls: 'green', glow: 'card-glow-green' },
      { title: 'Net Sells', value: fmtInt(sells.length), sub: fmtMM(totalNetSellValue), cls: 'red', glow: 'card-glow-red' },
      { title: 'Total Net Buy Value', value: '$' + fmt(totalNetBuyValue, 1) + ' MM', cls: 'green', glow: 'card-glow-green' },
      { title: 'Total Net Sell Value', value: '$' + fmt(Math.abs(totalNetSellValue), 1) + ' MM', cls: 'red', glow: 'card-glow-red' }
    ];
  } else {
    cardData = [
      { title: 'Two-Way Flow (Before Crossing)', value: summary.beforeCrossing ? '$' + fmt(summary.beforeCrossing, 0) + ' MM' : '\u2014', cls: 'cyan' },
      { title: 'Crossing Flow', value: summary.crossingFlow ? '$' + fmt(summary.crossingFlow, 0) + ' MM' : '\u2014', cls: '' },
      { title: 'After Crossing', value: summary.afterCrossing ? '$' + fmt(summary.afterCrossing, 0) + ' MM' : '\u2014', cls: 'cyan' },
      { title: 'Net Buys', value: fmtInt(buys.length), sub: fmtMM(totalNetBuyValue), cls: 'green', glow: 'card-glow-green' },
      { title: 'Net Sells', value: fmtInt(sells.length), sub: fmtMM(totalNetSellValue), cls: 'red', glow: 'card-glow-red' },
      { title: 'Total Net Buy Value', value: '$' + fmt(totalNetBuyValue, 1) + ' MM', cls: 'green', glow: 'card-glow-green' },
      { title: 'Total Net Sell Value', value: '$' + fmt(Math.abs(totalNetSellValue), 1) + ' MM', cls: 'red', glow: 'card-glow-red' }
    ];
  }

  cards.innerHTML = cardData.map(c => `
    <div class="card ${c.glow || ''}">
      <div class="card-title">${c.title}</div>
      <div class="card-value ${c.cls}">${c.value}</div>
      ${c.sub ? `<div class="card-sub">${c.sub}</div>` : ''}
    </div>
  `).join('');

  // Flow breakdown
  renderFlowBreakdown();

  // Industry charts
  renderIndustryNetChart();

  // Two-Way chart: show for MSCI always, for S&P show index distribution instead
  const twoWaySection = document.getElementById('chart-industry-twoway').closest('.section-card');
  if (STATE.format === 'sp') {
    twoWaySection.querySelector('.section-title').textContent = 'Index Distribution';
    renderIndexDistributionChart();
  } else {
    twoWaySection.querySelector('.section-title').textContent = 'Industry Two-Way Flows (USD MM)';
    renderIndustryTwoWayChart();
  }

  // Country/Index donut
  const donutSection = document.getElementById('chart-country-donut').closest('.section-card');
  if (STATE.format === 'sp') {
    donutSection.querySelector('.section-title').textContent = 'Index Distribution';
    renderIndexDonut();
  } else {
    donutSection.querySelector('.section-title').textContent = 'Country Distribution';
    renderCountryDonut();
  }
}

function renderFlowBreakdown() {
  const body = document.getElementById('flow-breakdown-body');
  const titleEl = body.closest('.section-card').querySelector('.section-title');

  if (STATE.format === 'sp') {
    titleEl.textContent = 'S&P Index Buy/Sell Stats';
    const stats = STATE.summary.spIndexStats || {};
    const indices = Object.keys(stats);
    if (indices.length === 0) {
      // Compute from overall data
      const byIndex = {};
      STATE.overall.forEach(d => {
        if (!d.index) return;
        if (!byIndex[d.index]) byIndex[d.index] = { buys: 0, sells: 0, buyVal: 0, sellVal: 0 };
        if (d.direction === 'BUY') {
          byIndex[d.index].buys++;
          byIndex[d.index].buyVal += d.netValue;
        } else {
          byIndex[d.index].sells++;
          byIndex[d.index].sellVal += Math.abs(d.netValue);
        }
      });
      body.innerHTML = `
        <table class="flow-table">
          <thead><tr><th>Index</th><th class="num-cell">Buys</th><th class="num-cell">Sells</th><th class="num-cell">Net Buy Value ($MM)</th><th class="num-cell">Net Sell Value ($MM)</th></tr></thead>
          <tbody>
            ${Object.entries(byIndex).map(([idx, v]) => `
              <tr>
                <td style="color:var(--text-primary)">${idx}</td>
                <td class="num-cell green-text">${v.buys}</td>
                <td class="num-cell red-text">${v.sells}</td>
                <td class="num-cell green-text">${fmtMM(v.buyVal)}</td>
                <td class="num-cell red-text">${fmtMM(v.sellVal)}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      `;
    } else {
      body.innerHTML = `
        <table class="flow-table">
          <thead><tr><th>Index</th><th class="num-cell">Buys</th><th class="num-cell">Buy Value ($MM)</th><th class="num-cell">Sells</th><th class="num-cell">Sell Value ($MM)</th><th class="num-cell">Total</th></tr></thead>
          <tbody>
            ${indices.map(idx => {
              const s = stats[idx];
              return `
                <tr>
                  <td style="color:var(--text-primary)">${idx}</td>
                  <td class="num-cell green-text">${s.buys ? fmtInt(s.buys.count) : '\u2014'}</td>
                  <td class="num-cell green-text">${s.buys ? fmtMM(s.buys.value) : '\u2014'}</td>
                  <td class="num-cell red-text">${s.sells ? fmtInt(s.sells.count) : '\u2014'}</td>
                  <td class="num-cell red-text">${s.sells ? fmtMM(s.sells.value) : '\u2014'}</td>
                  <td class="num-cell">${s.total ? fmtInt(s.total.count) : '\u2014'}</td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>
      `;
    }
    return;
  }

  // MSCI flow breakdown
  titleEl.textContent = 'Flow Breakdown by Index Type';
  const breakdown = STATE.summary.flowBreakdown;
  if (!breakdown.length) {
    body.innerHTML = '<p style="color:var(--text-muted);">No flow breakdown data found in Summary sheet.</p>';
    return;
  }

  const maxVal = Math.max(...breakdown.map(b => Math.abs(b.value)));

  body.innerHTML = `
    <table class="flow-table">
      <thead><tr><th>Index Type</th><th class="num-cell">USD MM</th><th class="num-cell">% of Total</th><th class="bar-cell">Distribution</th></tr></thead>
      <tbody>
        ${breakdown.map(b => `
          <tr>
            <td style="color:var(--text-primary)">${b.name}</td>
            <td class="num-cell" style="font-family:var(--font-mono)">${fmt(b.value, 0)}</td>
            <td class="num-cell" style="font-family:var(--font-mono)">${b.pct ? fmtPct(b.pct) : '\u2014'}</td>
            <td class="bar-cell">
              <div class="flow-bar-wrap">
                <div class="flow-bar positive" style="width:${(Math.abs(b.value)/maxVal*100).toFixed(1)}%"></div>
              </div>
            </td>
          </tr>
        `).join('')}
      </tbody>
    </table>
  `;
}

function renderIndustryNetChart() {
  let industries = STATE.summary.industryNetFlows;
  if (!industries.length) {
    const byIndustry = {};
    STATE.overall.forEach(d => {
      if (!d.industry) return;
      byIndustry[d.industry] = (byIndustry[d.industry] || 0) + d.netValue;
    });
    industries = Object.entries(byIndustry).map(([name, value]) => ({ name, value }));
  }
  industries.sort((a, b) => b.value - a.value);

  destroyChart('industryNet');
  const ctx = document.getElementById('chart-industry-net').getContext('2d');
  STATE.charts.industryNet = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: industries.map(i => i.name),
      datasets: [{
        data: industries.map(i => i.value),
        backgroundColor: industries.map(i => i.value >= 0 ? '#00ff8866' : '#ff335566'),
        borderColor: industries.map(i => i.value >= 0 ? '#00ff88' : '#ff3355'),
        borderWidth: 1
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => '$' + fmt(ctx.raw, 0) + ' MM' } } },
      scales: {
        x: { grid: { color: '#1e1e30' }, ticks: { color: '#8888a0', font: { family: 'JetBrains Mono', size: 10 } } },
        y: { grid: { display: false }, ticks: { color: '#e8e8f0', font: { family: 'DM Sans', size: 11 }, autoSkip: false } }
      }
    }
  });
  document.getElementById('chart-industry-net').parentElement.style.height = Math.max(400, industries.length * 24) + 'px';
}

function renderIndustryTwoWayChart() {
  let industries = STATE.summary.industryTwoWayFlows;
  if (!industries.length) {
    const byIndustry = {};
    STATE.overall.forEach(d => {
      if (!d.industry) return;
      byIndustry[d.industry] = (byIndustry[d.industry] || 0) + Math.abs(d.grossValue);
    });
    industries = Object.entries(byIndustry).map(([name, value]) => ({ name, value }));
  }
  industries.sort((a, b) => b.value - a.value);

  destroyChart('industryTwoWay');
  const ctx = document.getElementById('chart-industry-twoway').getContext('2d');
  STATE.charts.industryTwoWay = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: industries.map(i => i.name),
      datasets: [{
        data: industries.map(i => Math.abs(i.value)),
        backgroundColor: '#00d4ff44',
        borderColor: '#00d4ff',
        borderWidth: 1
      }]
    },
    options: {
      indexAxis: 'y',
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { callbacks: { label: ctx => '$' + fmt(ctx.raw, 0) + ' MM' } } },
      scales: {
        x: { grid: { color: '#1e1e30' }, ticks: { color: '#8888a0', font: { family: 'JetBrains Mono', size: 10 } } },
        y: { grid: { display: false }, ticks: { color: '#e8e8f0', font: { family: 'DM Sans', size: 11 }, autoSkip: false } }
      }
    }
  });
  document.getElementById('chart-industry-twoway').parentElement.style.height = Math.max(400, industries.length * 24) + 'px';
}

function renderIndexDistributionChart() {
  // S&P: Show a bar chart of names per index
  const byIndex = {};
  STATE.overall.forEach(d => {
    if (!d.index) return;
    if (!byIndex[d.index]) byIndex[d.index] = { buys: 0, sells: 0 };
    if (d.direction === 'BUY') byIndex[d.index].buys++;
    else byIndex[d.index].sells++;
  });

  const indices = Object.keys(byIndex).sort();
  const buyData = indices.map(i => byIndex[i].buys);
  const sellData = indices.map(i => byIndex[i].sells);

  destroyChart('industryTwoWay');
  const ctx = document.getElementById('chart-industry-twoway').getContext('2d');
  STATE.charts.industryTwoWay = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: indices,
      datasets: [
        {
          label: 'Buys',
          data: buyData,
          backgroundColor: '#00ff8866',
          borderColor: '#00ff88',
          borderWidth: 1
        },
        {
          label: 'Sells',
          data: sellData,
          backgroundColor: '#ff335566',
          borderColor: '#ff3355',
          borderWidth: 1
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { labels: { color: '#8888a0', font: { family: 'DM Sans', size: 11 } } },
        tooltip: { callbacks: { label: ctx => `${ctx.dataset.label}: ${ctx.raw} names` } }
      },
      scales: {
        x: { grid: { color: '#1e1e30' }, ticks: { color: '#e8e8f0', font: { family: 'DM Sans', size: 11 } } },
        y: { grid: { color: '#1e1e30' }, ticks: { color: '#8888a0', font: { family: 'JetBrains Mono', size: 10 } } }
      }
    }
  });
  document.getElementById('chart-industry-twoway').parentElement.style.height = '350px';
}

function renderCountryDonut() {
  const byCountry = {};
  STATE.overall.forEach(d => {
    if (!d.country) return;
    byCountry[d.country] = (byCountry[d.country] || 0) + 1;
  });

  const sorted = Object.entries(byCountry).sort((a, b) => b[1] - a[1]);
  const colors = generateColors(sorted.length);

  destroyChart('countryDonut');
  const ctx = document.getElementById('chart-country-donut').getContext('2d');
  STATE.charts.countryDonut = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: sorted.map(s => s[0]),
      datasets: [{
        data: sorted.map(s => s[1]),
        backgroundColor: colors,
        borderColor: '#12121a',
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => `${ctx.label}: ${ctx.raw} stocks (${((ctx.raw / STATE.overall.length) * 100).toFixed(1)}%)`
          }
        }
      }
    }
  });

  const legend = document.getElementById('country-legend');
  legend.innerHTML = sorted.map((s, i) => `
    <div class="legend-item">
      <div class="legend-dot" style="background:${colors[i]}"></div>
      <span class="legend-label">${s[0]}</span>
      <span class="legend-count">${s[1]}</span>
    </div>
  `).join('');
}

function renderIndexDonut() {
  const byIndex = {};
  STATE.overall.forEach(d => {
    if (!d.index) return;
    byIndex[d.index] = (byIndex[d.index] || 0) + 1;
  });

  const sorted = Object.entries(byIndex).sort((a, b) => b[1] - a[1]);
  const colors = generateColors(sorted.length);

  destroyChart('countryDonut');
  const ctx = document.getElementById('chart-country-donut').getContext('2d');
  STATE.charts.countryDonut = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels: sorted.map(s => s[0]),
      datasets: [{
        data: sorted.map(s => s[1]),
        backgroundColor: colors,
        borderColor: '#12121a',
        borderWidth: 2
      }]
    },
    options: {
      responsive: true,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: ctx => `${ctx.label}: ${ctx.raw} names (${((ctx.raw / STATE.overall.length) * 100).toFixed(1)}%)`
          }
        }
      }
    }
  });

  const legend = document.getElementById('country-legend');
  legend.innerHTML = sorted.map((s, i) => `
    <div class="legend-item">
      <div class="legend-dot" style="background:${colors[i]}"></div>
      <span class="legend-label">${s[0]}</span>
      <span class="legend-count">${s[1]}</span>
    </div>
  `).join('');
}

function generateColors(n) {
  const base = [
    '#00d4ff', '#00ff88', '#ff3355', '#ffaa00', '#8855ff',
    '#ff6600', '#00ffd4', '#ff55aa', '#44aaff', '#aaff00',
    '#ff0088', '#00aaff', '#ff8800', '#55ff88', '#aa55ff',
    '#ff4444', '#44ff44', '#4444ff', '#ffff44', '#ff44ff'
  ];
  const colors = [];
  for (let i = 0; i < n; i++) {
    if (i < base.length) colors.push(base[i]);
    else colors.push(`hsl(${(i * 37) % 360}, 70%, 55%)`);
  }
  return colors;
}

// ============================================================
// TAB: ALL NAMES
// ============================================================
function renderAllNames() {
  const filters = document.getElementById('allnames-filters');
  const cols = getAllNamesCols();

  if (STATE.format === 'sp') {
    // Build unique indices and industries
    const indices = [...new Set(STATE.overall.map(d => d.index).filter(Boolean))].sort();
    const industries = [...new Set(STATE.overall.map(d => d.industry).filter(Boolean))].sort();

    filters.innerHTML = `
      <input type="text" class="filter-input" id="an-search" placeholder="Search ticker or name...">
      <select class="filter-select" id="an-index">
        <option value="all">All Indices</option>
        ${indices.map(i => `<option value="${i}">${i}</option>`).join('')}
      </select>
      <select class="filter-select" id="an-industry">
        <option value="all">All Industries</option>
        ${industries.map(i => `<option value="${i}">${i}</option>`).join('')}
      </select>
      <button class="filter-btn" id="an-buy" data-dir="BUY">Buys</button>
      <button class="filter-btn" id="an-sell" data-dir="SELL">Sells</button>
      <button class="filter-btn active" id="an-all" data-dir="all">All</button>
      <button class="export-btn" id="an-export">&#x2B73; Export CSV</button>
    `;

    document.getElementById('an-index').addEventListener('change', e => {
      STATE.allNamesFilter.index = e.target.value;
      STATE.allNamesPage = 1;
      renderAllNamesTable();
    });
  } else {
    const countries = [...new Set(STATE.overall.map(d => d.country).filter(Boolean))].sort();
    const industries = [...new Set(STATE.overall.map(d => d.industry).filter(Boolean))].sort();

    filters.innerHTML = `
      <input type="text" class="filter-input" id="an-search" placeholder="Search ticker or issuer...">
      <select class="filter-select" id="an-country">
        <option value="all">All Countries</option>
        ${countries.map(c => `<option value="${c}">${c}</option>`).join('')}
      </select>
      <select class="filter-select" id="an-industry">
        <option value="all">All Industries</option>
        ${industries.map(i => `<option value="${i}">${i}</option>`).join('')}
      </select>
      <button class="filter-btn" id="an-buy" data-dir="BUY">Buys</button>
      <button class="filter-btn" id="an-sell" data-dir="SELL">Sells</button>
      <button class="filter-btn active" id="an-all" data-dir="all">All</button>
      <button class="export-btn" id="an-export">&#x2B73; Export CSV</button>
    `;

    document.getElementById('an-country').addEventListener('change', e => {
      STATE.allNamesFilter.country = e.target.value;
      STATE.allNamesPage = 1;
      renderAllNamesTable();
    });
  }

  // Headers
  const thead = document.getElementById('allnames-thead');
  thead.innerHTML = '<tr>' + cols.map(c =>
    `<th class="${c.num ? 'num-cell' : ''}" data-col="${c.key}">${c.label} <span class="sort-arrow"></span></th>`
  ).join('') + '</tr>';

  // Common events
  document.getElementById('an-search').addEventListener('input', e => {
    STATE.allNamesFilter.search = e.target.value.toLowerCase();
    STATE.allNamesPage = 1;
    renderAllNamesTable();
  });

  document.getElementById('an-industry').addEventListener('change', e => {
    STATE.allNamesFilter.industry = e.target.value;
    STATE.allNamesPage = 1;
    renderAllNamesTable();
  });

  ['an-buy', 'an-sell', 'an-all'].forEach(id => {
    document.getElementById(id).addEventListener('click', e => {
      ['an-buy', 'an-sell', 'an-all'].forEach(bid => document.getElementById(bid).classList.remove('active'));
      e.target.classList.add('active');
      STATE.allNamesFilter.direction = e.target.dataset.dir;
      STATE.allNamesPage = 1;
      renderAllNamesTable();
    });
  });

  // Sort
  thead.querySelectorAll('th').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.col;
      if (STATE.allNamesSort.col === col) {
        STATE.allNamesSort.dir = STATE.allNamesSort.dir === 'asc' ? 'desc' : 'asc';
      } else {
        STATE.allNamesSort.col = col;
        STATE.allNamesSort.dir = 'desc';
      }
      thead.querySelectorAll('th').forEach(t => t.classList.remove('sort-asc', 'sort-desc'));
      th.classList.add('sort-' + STATE.allNamesSort.dir);
      renderAllNamesTable();
    });
  });

  // Export
  document.getElementById('an-export').addEventListener('click', exportAllNamesCSV);

  renderAllNamesTable();
}

function getFilteredAllNames() {
  let data = [...STATE.overall];
  const f = STATE.allNamesFilter;

  if (f.search) {
    data = data.filter(d =>
      d.cleanTicker.toLowerCase().includes(f.search) ||
      d.ticker.toLowerCase().includes(f.search) ||
      d.issuer.toLowerCase().includes(f.search)
    );
  }
  if (f.country !== 'all') data = data.filter(d => d.country === f.country);
  if (f.industry !== 'all') data = data.filter(d => d.industry === f.industry);
  if (f.direction !== 'all') data = data.filter(d => d.direction === f.direction);
  if (f.index !== 'all') data = data.filter(d => d.index === f.index);

  const s = STATE.allNamesSort;
  if (s.col) {
    data.sort((a, b) => {
      let va = a[s.col], vb = b[s.col];
      if (typeof va === 'string') va = va.toLowerCase();
      if (typeof vb === 'string') vb = vb.toLowerCase();
      if (va < vb) return s.dir === 'asc' ? -1 : 1;
      if (va > vb) return s.dir === 'asc' ? 1 : -1;
      return 0;
    });
  }

  return data;
}

function renderAllNamesTable() {
  const filtered = getFilteredAllNames();
  const totalPages = Math.ceil(filtered.length / PER_PAGE);
  if (STATE.allNamesPage > totalPages) STATE.allNamesPage = Math.max(1, totalPages);

  const start = (STATE.allNamesPage - 1) * PER_PAGE;
  const page = filtered.slice(start, start + PER_PAGE);

  document.getElementById('allnames-count').textContent = fmtInt(filtered.length) + ' STOCKS';

  const tbody = document.getElementById('allnames-tbody');

  if (STATE.format === 'sp') {
    tbody.innerHTML = page.map(d => `
      <tr class="${d.direction === 'BUY' ? 'row-buy' : 'row-sell'}">
        <td><strong style="color:var(--cyan)">${d.cleanTicker}</strong></td>
        <td>${d.issuer}</td>
        <td>${d.index}</td>
        <td>${d.industry}</td>
        <td>${eventBadge(d.eventType)}</td>
        <td class="num-cell">${fmtPrice(d.price)}</td>
        <td class="num-cell">${fmtInt(d.avgVol)}</td>
        <td class="num-cell ${dirClass(d.netShares)}">${fmtInt(d.netShares)}</td>
        <td class="num-cell ${dirClass(d.netValue)}">${fmtMM(d.netValue)}</td>
        <td class="num-cell">${fmtPct(d.netLiq)}</td>
        <td class="num-cell">${fmtPct(d.absLiq)}</td>
      </tr>
    `).join('');
  } else {
    tbody.innerHTML = page.map(d => `
      <tr class="${d.direction === 'BUY' ? 'row-buy' : 'row-sell'}">
        <td><strong style="color:var(--cyan)">${d.cleanTicker}</strong></td>
        <td>${d.issuer}</td>
        <td>${d.country}</td>
        <td>${d.industry}</td>
        <td>${eventBadge(d.eventType)}</td>
        <td class="num-cell">${fmtPrice(d.price)}</td>
        <td class="num-cell">${fmtInt(d.avgVol)}</td>
        <td class="num-cell ${dirClass(d.netShares)}">${fmtInt(d.netShares)}</td>
        <td class="num-cell ${dirClass(d.netValue)}">${fmtMM(d.netValue)}</td>
        <td class="num-cell">${fmtPct(d.netLiq)}</td>
        <td class="num-cell">${fmtInt(d.grossShares)}</td>
        <td class="num-cell">${fmtMM(d.grossValue)}</td>
        <td class="num-cell">${fmtPct(d.grossLiq)}</td>
      </tr>
    `).join('');
  }

  renderPagination('allnames-pagination', STATE.allNamesPage, totalPages, p => {
    STATE.allNamesPage = p;
    renderAllNamesTable();
  });
}

function renderPagination(containerId, currentPage, totalPages, onPageChange) {
  const container = document.getElementById(containerId);
  if (totalPages <= 1) { container.innerHTML = ''; return; }

  let html = '';
  html += `<button ${currentPage === 1 ? 'disabled' : ''} data-page="${currentPage - 1}">&laquo; Prev</button>`;

  const range = [];
  const delta = 2;
  for (let i = Math.max(1, currentPage - delta); i <= Math.min(totalPages, currentPage + delta); i++) {
    range.push(i);
  }
  if (range[0] > 1) { html += `<button data-page="1">1</button>`; if (range[0] > 2) html += `<span class="page-info">...</span>`; }
  range.forEach(p => {
    html += `<button class="${p === currentPage ? 'active' : ''}" data-page="${p}">${p}</button>`;
  });
  if (range[range.length - 1] < totalPages) { if (range[range.length - 1] < totalPages - 1) html += `<span class="page-info">...</span>`; html += `<button data-page="${totalPages}">${totalPages}</button>`; }

  html += `<button ${currentPage === totalPages ? 'disabled' : ''} data-page="${currentPage + 1}">Next &raquo;</button>`;
  html += `<span class="page-info">${currentPage} of ${totalPages}</span>`;

  container.innerHTML = html;
  container.querySelectorAll('button').forEach(btn => {
    btn.addEventListener('click', () => {
      const p = parseInt(btn.dataset.page);
      if (p >= 1 && p <= totalPages) onPageChange(p);
    });
  });
}

function exportAllNamesCSV() {
  const data = getFilteredAllNames();
  let headers, rows;

  if (STATE.format === 'sp') {
    headers = ['Ticker', 'Name', 'Index', 'Industry', 'Event', 'Price', '20D Avg Vol', 'Net Shares', 'Net Value ($MM)', 'Net Liq %', 'Abs Liq %'];
    rows = data.map(d => [
      d.cleanTicker, d.issuer, d.index, d.industry, d.eventType || '', d.price, d.avgVol,
      d.netShares, d.netValue, d.netLiq, d.absLiq
    ]);
  } else {
    headers = ['Ticker', 'Issuer', 'Country', 'Industry', 'Event', 'Price', '20D Avg Vol', 'Net Shares', 'Net Value ($MM)', 'Net Liq %', 'Gross Shares', 'Gross Value ($MM)', 'Gross Liq %'];
    rows = data.map(d => [
      d.cleanTicker, d.issuer, d.country, d.industry, d.eventType || '', d.price, d.avgVol,
      d.netShares, d.netValue, d.netLiq, d.grossShares, d.grossValue, d.grossLiq
    ]);
  }

  let csv = headers.join(',') + '\n';
  rows.forEach(r => {
    csv += r.map(v => {
      const s = String(v);
      return s.includes(',') ? `"${s}"` : s;
    }).join(',') + '\n';
  });

  const blob = new Blob([csv], { type: 'text/csv' });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = (STATE.format === 'sp' ? 'sp_rebalance' : 'msci_rebalance') + '_allnames.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}

// ============================================================
// TAB: TOP MOVERS
// ============================================================
function renderTopMovers() {
  const grid = document.getElementById('topmovers-grid');
  const hasCountry = STATE.format !== 'sp';
  const panels = [
    { title: 'Most Illiquid Net Buys', data: STATE.topNames.illiquidBuys, colorClass: 'green' },
    { title: 'Most Illiquid Net Sells', data: STATE.topNames.illiquidSells, colorClass: 'red' },
    { title: 'Largest Net Buys by Trade Value', data: STATE.topNames.largestBuys, colorClass: 'green' },
    { title: 'Largest Net Sells by Trade Value', data: STATE.topNames.largestSells, colorClass: 'red' }
  ];

  grid.innerHTML = panels.map(p => `
    <div class="section-card ${p.colorClass === 'green' ? 'card-glow-green' : 'card-glow-red'}">
      <div class="section-header">
        <div class="section-title">${p.title}</div>
        <div class="page-badge" style="font-size:0.6rem;">${p.data.length} names</div>
      </div>
      <div class="section-body" style="max-height:500px; overflow-y:auto;">
        ${p.data.length > 0 ? `
          <table class="data-table">
            <thead><tr>
              <th>#</th><th>Ticker</th><th>Name</th>${hasCountry ? '<th>Country</th>' : ''}
              <th class="num-cell">Trade Value ($MM)</th><th class="num-cell">Liquidity (Days ADV)</th>
            </tr></thead>
            <tbody>
              ${p.data.map((d, i) => `
                <tr class="${i < 5 ? 'top-highlight' : ''}">
                  <td style="color:var(--text-muted)">${i + 1}</td>
                  <td><strong style="color:var(--cyan)">${d.ticker}</strong></td>
                  <td>${d.name}</td>
                  ${hasCountry ? `<td>${d.country}</td>` : ''}
                  <td class="num-cell ${p.colorClass === 'green' ? 'green-text' : 'red-text'}">${fmt(d.tradeValue, 1)}</td>
                  <td class="num-cell">${fmt(Math.abs(d.liquidity), 2)}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
        ` : '<p style="color:var(--text-muted);">No data available</p>'}
      </div>
    </div>
  `).join('');
}

// ============================================================
// TAB: COMPARISON (MSCI) / INDEX BREAKDOWN (S&P)
// ============================================================
function renderComparison() {
  if (STATE.format === 'sp') {
    renderComparisonSP();
  } else {
    renderComparisonMSCI();
  }
}

function renderComparisonMSCI() {
  const filters = document.getElementById('comparison-filters');
  filters.innerHTML = `
    <input type="text" class="filter-input" id="comp-search" placeholder="Search ticker or issuer...">
    <button class="filter-btn active" id="comp-all" data-dir="all">All</button>
    <button class="filter-btn" id="comp-changed" data-dir="changed">Direction Changed</button>
    <button class="filter-btn" id="comp-new" data-dir="new">New Additions</button>
  `;

  const thead = document.getElementById('comparison-thead');
  thead.innerHTML = `<tr>
    <th data-col="ticker">Ticker</th>
    <th data-col="issuer">Issuer</th>
    <th data-col="country">Country</th>
    <th data-col="dirChange">Direction Change</th>
    <th class="num-cell" data-col="pctChange">% Change Net Shares</th>
    <th class="num-cell" data-col="netStdVal">Net Std Value</th>
    <th class="num-cell" data-col="netFullVal">Net Full Value</th>
    <th class="num-cell" data-col="grossStdVal">Gross Std Value</th>
    <th class="num-cell" data-col="grossFullVal">Gross Full Value</th>
  </tr>`;

  document.getElementById('comp-search').addEventListener('input', e => {
    STATE.compFilter.search = e.target.value.toLowerCase();
    STATE.compPage = 1;
    renderComparisonTable();
  });

  ['comp-all', 'comp-changed', 'comp-new'].forEach(id => {
    document.getElementById(id).addEventListener('click', e => {
      ['comp-all', 'comp-changed', 'comp-new'].forEach(bid => document.getElementById(bid).classList.remove('active'));
      e.target.classList.add('active');
      STATE.compFilter.direction = e.target.dataset.dir;
      STATE.compPage = 1;
      renderComparisonTable();
    });
  });

  thead.querySelectorAll('th').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.col;
      if (STATE.compSort.col === col) {
        STATE.compSort.dir = STATE.compSort.dir === 'asc' ? 'desc' : 'asc';
      } else {
        STATE.compSort.col = col;
        STATE.compSort.dir = 'desc';
      }
      thead.querySelectorAll('th').forEach(t => t.classList.remove('sort-asc', 'sort-desc'));
      th.classList.add('sort-' + STATE.compSort.dir);
      renderComparisonTable();
    });
  });

  renderComparisonTable();
}

function renderComparisonSP() {
  // S&P: Show INDEX x CHANGE breakdown
  const filters = document.getElementById('comparison-filters');
  filters.innerHTML = `
    <input type="text" class="filter-input" id="comp-search" placeholder="Search ticker or name...">
    <button class="filter-btn active" id="comp-all" data-dir="all">All</button>
    <button class="filter-btn" id="comp-additions" data-dir="additions">Additions</button>
    <button class="filter-btn" id="comp-deletions" data-dir="deletions">Deletions</button>
  `;

  const thead = document.getElementById('comparison-thead');
  thead.innerHTML = `<tr>
    <th data-col="cleanTicker">Ticker</th>
    <th data-col="issuer">Name</th>
    <th data-col="index">Index</th>
    <th data-col="eventType">Change Type</th>
    <th class="num-cell" data-col="price">Price</th>
    <th class="num-cell" data-col="netValue">Net Value ($MM)</th>
    <th class="num-cell" data-col="netShares">Net Shares</th>
    <th class="num-cell" data-col="netLiq">Liquidity %</th>
  </tr>`;

  // Update page header
  const pageTitle = document.querySelector('#tab-comparison .page-title');
  if (pageTitle) pageTitle.textContent = 'Index Breakdown';
  const pageBadge = document.querySelector('#tab-comparison .page-badge');
  if (pageBadge) pageBadge.textContent = 'BY INDEX & CHANGE TYPE';

  document.getElementById('comp-search').addEventListener('input', e => {
    STATE.compFilter.search = e.target.value.toLowerCase();
    STATE.compPage = 1;
    renderComparisonTableSP();
  });

  ['comp-all', 'comp-additions', 'comp-deletions'].forEach(id => {
    document.getElementById(id).addEventListener('click', e => {
      ['comp-all', 'comp-additions', 'comp-deletions'].forEach(bid => document.getElementById(bid).classList.remove('active'));
      e.target.classList.add('active');
      STATE.compFilter.direction = e.target.dataset.dir;
      STATE.compPage = 1;
      renderComparisonTableSP();
    });
  });

  thead.querySelectorAll('th').forEach(th => {
    th.addEventListener('click', () => {
      const col = th.dataset.col;
      if (STATE.compSort.col === col) {
        STATE.compSort.dir = STATE.compSort.dir === 'asc' ? 'desc' : 'asc';
      } else {
        STATE.compSort.col = col;
        STATE.compSort.dir = 'desc';
      }
      thead.querySelectorAll('th').forEach(t => t.classList.remove('sort-asc', 'sort-desc'));
      th.classList.add('sort-' + STATE.compSort.dir);
      renderComparisonTableSP();
    });
  });

  renderComparisonTableSP();
}

function getFilteredComparisonSP() {
  let data = [...STATE.overall];
  const f = STATE.compFilter;

  if (f.search) {
    data = data.filter(d =>
      d.cleanTicker.toLowerCase().includes(f.search) ||
      d.issuer.toLowerCase().includes(f.search)
    );
  }

  if (f.direction === 'additions') {
    data = data.filter(d => d.eventType === 'ADDITION' || (d.eventType && d.eventType.includes('\u2192')));
  }
  if (f.direction === 'deletions') {
    data = data.filter(d => d.eventType === 'DELETION');
  }

  const s = STATE.compSort;
  if (s.col) {
    data.sort((a, b) => {
      let va = a[s.col], vb = b[s.col];
      if (typeof va === 'string') va = va.toLowerCase();
      if (typeof vb === 'string') vb = vb.toLowerCase();
      if (va < vb) return s.dir === 'asc' ? -1 : 1;
      if (va > vb) return s.dir === 'asc' ? 1 : -1;
      return 0;
    });
  }

  return data;
}

function renderComparisonTableSP() {
  const filtered = getFilteredComparisonSP();
  const totalPages = Math.ceil(filtered.length / PER_PAGE);
  if (STATE.compPage > totalPages) STATE.compPage = Math.max(1, totalPages);

  const start = (STATE.compPage - 1) * PER_PAGE;
  const page = filtered.slice(start, start + PER_PAGE);

  const tbody = document.getElementById('comparison-tbody');
  tbody.innerHTML = page.map(d => `
    <tr class="${d.direction === 'BUY' ? 'row-buy' : 'row-sell'}">
      <td><strong style="color:var(--cyan)">${d.cleanTicker}</strong></td>
      <td>${d.issuer}</td>
      <td>${d.index}</td>
      <td>${eventBadge(d.eventType)}</td>
      <td class="num-cell">${fmtPrice(d.price)}</td>
      <td class="num-cell ${dirClass(d.netValue)}">${fmtMM(d.netValue)}</td>
      <td class="num-cell ${dirClass(d.netShares)}">${fmtInt(d.netShares)}</td>
      <td class="num-cell">${fmtPct(d.netLiq)}</td>
    </tr>
  `).join('');

  renderPagination('comparison-pagination', STATE.compPage, totalPages, p => {
    STATE.compPage = p;
    renderComparisonTableSP();
  });
}

function getFilteredComparison() {
  let data = [...STATE.comparison];
  const f = STATE.compFilter;

  if (f.search) {
    data = data.filter(d =>
      d.ticker.toLowerCase().includes(f.search) ||
      d.issuer.toLowerCase().includes(f.search)
    );
  }

  if (f.direction === 'changed') {
    data = data.filter(d => d.dirChange && !d.dirChange.toLowerCase().includes('no change') && d.dirChange !== '' && d.dirChange !== '0');
  }
  if (f.direction === 'new') {
    data = data.filter(d => d.dirChange && (d.dirChange.toLowerCase().includes('new') || d.dirChange.toLowerCase().includes('add')));
  }

  const s = STATE.compSort;
  if (s.col) {
    const getVal = (item) => {
      switch(s.col) {
        case 'netStdVal': return item.netStandard.value;
        case 'netFullVal': return item.netFull.value;
        case 'grossStdVal': return item.grossStandard.value;
        case 'grossFullVal': return item.grossFull.value;
        default: return item[s.col];
      }
    };
    data.sort((a, b) => {
      let va = getVal(a), vb = getVal(b);
      if (typeof va === 'string') va = va.toLowerCase();
      if (typeof vb === 'string') vb = vb.toLowerCase();
      if (va < vb) return s.dir === 'asc' ? -1 : 1;
      if (va > vb) return s.dir === 'asc' ? 1 : -1;
      return 0;
    });
  }

  return data;
}

function renderComparisonTable() {
  const filtered = getFilteredComparison();
  const totalPages = Math.ceil(filtered.length / PER_PAGE);
  if (STATE.compPage > totalPages) STATE.compPage = Math.max(1, totalPages);

  const start = (STATE.compPage - 1) * PER_PAGE;
  const page = filtered.slice(start, start + PER_PAGE);

  const tbody = document.getElementById('comparison-tbody');
  tbody.innerHTML = page.map(d => {
    const dirBadge = getDirBadge(d.dirChange);
    return `
      <tr>
        <td><strong style="color:var(--cyan)">${d.ticker}</strong></td>
        <td>${d.issuer}</td>
        <td>${d.country}</td>
        <td>${dirBadge}</td>
        <td class="num-cell ${d.pctChange > 0 ? 'green-text' : d.pctChange < 0 ? 'red-text' : ''}">${d.pctChange ? fmtPct(d.pctChange) : '\u2014'}</td>
        <td class="num-cell ${dirClass(d.netStandard.value)}">${fmt(d.netStandard.value, 1)}</td>
        <td class="num-cell ${dirClass(d.netFull.value)}">${fmt(d.netFull.value, 1)}</td>
        <td class="num-cell">${fmt(d.grossStandard.value, 1)}</td>
        <td class="num-cell">${fmt(d.grossFull.value, 1)}</td>
      </tr>
    `;
  }).join('');

  renderPagination('comparison-pagination', STATE.compPage, totalPages, p => {
    STATE.compPage = p;
    renderComparisonTable();
  });
}

function getDirBadge(dirChange) {
  if (!dirChange) return '<span class="direction-badge no-change">\u2014</span>';
  const d = dirChange.toLowerCase();
  if (d.includes('buy') && d.includes('sell')) return '<span class="direction-badge buy-to-sell">Buy \u2192 Sell</span>';
  if (d.includes('sell') && d.includes('buy')) return '<span class="direction-badge sell-to-buy">Sell \u2192 Buy</span>';
  if (d.includes('new') || d.includes('add')) return '<span class="direction-badge new">New</span>';
  if (d.includes('no change') || d === '0' || d === '') return '<span class="direction-badge no-change">No Change</span>';
  return `<span class="direction-badge no-change">${dirChange}</span>`;
}

// ============================================================
// TAB: V-LOOKUP
// ============================================================
function renderVLookup() {
  const container = document.getElementById('vlookup-inputs');
  container.innerHTML = '';

  for (let i = 0; i < 10; i++) {
    const field = document.createElement('div');
    field.className = 'vlookup-field';
    field.innerHTML = `
      <input type="text" placeholder="Symbol ${i + 1}" id="vl-input-${i}" autocomplete="off">
      <div class="autocomplete-list" id="vl-ac-${i}"></div>
    `;
    container.appendChild(field);

    const input = field.querySelector('input');
    const acList = field.querySelector('.autocomplete-list');

    input.addEventListener('input', () => {
      const val = input.value.toUpperCase().trim();
      if (val.length < 1) { acList.classList.remove('show'); return; }

      const matches = STATE.allTickers.filter(t =>
        t.ticker.toUpperCase().startsWith(val) ||
        t.fullTicker.toUpperCase().startsWith(val) ||
        t.issuer.toUpperCase().includes(val)
      ).slice(0, 10);

      if (matches.length === 0) { acList.classList.remove('show'); return; }

      acList.innerHTML = matches.map(m => `
        <div class="autocomplete-item" data-ticker="${m.ticker}" data-full="${m.fullTicker}">
          <span class="ac-ticker">${m.ticker}</span>
          <span class="ac-name">${m.issuer.substring(0, 30)}</span>
        </div>
      `).join('');
      acList.classList.add('show');

      acList.querySelectorAll('.autocomplete-item').forEach(item => {
        item.addEventListener('click', () => {
          input.value = item.dataset.ticker;
          acList.classList.remove('show');
          updateVLookupResults();
        });
      });
    });

    input.addEventListener('keydown', e => {
      if (e.key === 'Enter') {
        acList.classList.remove('show');
        updateVLookupResults();
      }
    });

    input.addEventListener('blur', () => {
      setTimeout(() => acList.classList.remove('show'), 200);
    });
  }
}

function updateVLookupResults() {
  const symbols = [];
  for (let i = 0; i < 10; i++) {
    const val = document.getElementById('vl-input-' + i).value.toUpperCase().trim();
    if (val) symbols.push(val);
  }

  if (symbols.length === 0) {
    document.getElementById('vlookup-table-section').style.display = 'none';
    return;
  }

  const matches = symbols.map(sym => {
    return STATE.overall.find(d =>
      d.cleanTicker.toUpperCase() === sym ||
      d.ticker.toUpperCase() === sym ||
      d.ticker.toUpperCase().startsWith(sym + ' ')
    );
  }).filter(Boolean);

  const cols = getAllNamesCols();
  const colCount = cols.length;

  if (matches.length === 0) {
    document.getElementById('vlookup-table-section').style.display = 'block';
    document.getElementById('vlookup-count').textContent = '0 FOUND';
    document.getElementById('vlookup-thead').innerHTML = '';
    document.getElementById('vlookup-tbody').innerHTML = `<tr><td colspan="${colCount}" style="text-align:center;color:var(--text-muted);padding:2rem;">No matching symbols found</td></tr>`;
    return;
  }

  document.getElementById('vlookup-table-section').style.display = 'block';
  document.getElementById('vlookup-count').textContent = matches.length + ' FOUND';

  // Dynamic columns based on format
  document.getElementById('vlookup-thead').innerHTML = '<tr>' + cols.map(c =>
    `<th${c.num ? ' class="num-cell"' : ''}>${c.label}</th>`
  ).join('') + '</tr>';

  if (STATE.format === 'sp') {
    document.getElementById('vlookup-tbody').innerHTML = matches.map(d => `
      <tr class="${d.direction === 'BUY' ? 'row-buy' : 'row-sell'}">
        <td><strong style="color:var(--cyan)">${d.cleanTicker}</strong></td>
        <td>${d.issuer}</td>
        <td>${d.index}</td>
        <td>${d.industry}</td>
        <td>${eventBadge(d.eventType)}</td>
        <td class="num-cell">${fmtPrice(d.price)}</td>
        <td class="num-cell">${fmtInt(d.avgVol)}</td>
        <td class="num-cell ${dirClass(d.netShares)}">${fmtInt(d.netShares)}</td>
        <td class="num-cell ${dirClass(d.netValue)}">${fmtMM(d.netValue)}</td>
        <td class="num-cell">${fmtPct(d.netLiq)}</td>
        <td class="num-cell">${fmtPct(d.absLiq)}</td>
      </tr>
    `).join('');
  } else {
    document.getElementById('vlookup-tbody').innerHTML = matches.map(d => `
      <tr class="${d.direction === 'BUY' ? 'row-buy' : 'row-sell'}">
        <td><strong style="color:var(--cyan)">${d.cleanTicker}</strong></td>
        <td>${d.issuer}</td>
        <td>${d.country}</td>
        <td>${d.industry}</td>
        <td>${eventBadge(d.eventType)}</td>
        <td class="num-cell">${fmtPrice(d.price)}</td>
        <td class="num-cell">${fmtInt(d.avgVol)}</td>
        <td class="num-cell ${dirClass(d.netShares)}">${fmtInt(d.netShares)}</td>
        <td class="num-cell ${dirClass(d.netValue)}">${fmtMM(d.netValue)}</td>
        <td class="num-cell">${fmtPct(d.netLiq)}</td>
        <td class="num-cell">${fmtInt(d.grossShares)}</td>
        <td class="num-cell">${fmtMM(d.grossValue)}</td>
        <td class="num-cell">${fmtPct(d.grossLiq)}</td>
      </tr>
    `).join('');
  }
}

// ============================================================
// TAB: ANALYTICS
// ============================================================
function renderAnalytics() {
  const data = STATE.overall;
  if (data.length === 0) return;

  const isSP = STATE.format === 'sp';

  // 1. Market Impact Scores
  const impactScores = data.map(d => {
    const adv = d.price * d.avgVol;
    return adv > 0 ? Math.abs(d.netValue * 1e6) / adv : 0;
  }).filter(v => v > 0);
  const avgImpact = impactScores.length > 0 ? impactScores.reduce((s, v) => s + v, 0) / impactScores.length : 0;

  // 2. Concentration - Top 10 as % of total
  const sortedByAbsValue = [...data].sort((a, b) => Math.abs(b.netValue) - Math.abs(a.netValue));
  const totalAbsValue = data.reduce((s, d) => s + Math.abs(d.netValue), 0);
  const top10Value = sortedByAbsValue.slice(0, 10).reduce((s, d) => s + Math.abs(d.netValue), 0);
  const concentration = totalAbsValue > 0 ? top10Value / totalAbsValue : 0;

  // 3. Buy/Sell ratio
  const buys = data.filter(d => d.direction === 'BUY');
  const sells = data.filter(d => d.direction === 'SELL');
  const bsRatio = sells.length > 0 ? buys.length / sells.length : 0;

  // 4. Liquidity buckets
  const buckets = { '<1%': 0, '1-5%': 0, '5-10%': 0, '10-25%': 0, '>25%': 0 };
  data.forEach(d => {
    const liq = isSP ? Math.abs(d.netLiq) * 100 : Math.abs(d.grossLiq) * 100;
    if (liq < 1) buckets['<1%']++;
    else if (liq < 5) buckets['1-5%']++;
    else if (liq < 10) buckets['5-10%']++;
    else if (liq < 25) buckets['10-25%']++;
    else buckets['>25%']++;
  });

  // 5. VWAP Impact
  const vwapImpacts = data.map(d => {
    const shares = isSP ? Math.abs(d.netShares) : Math.abs(d.grossShares);
    const pctAdv = d.avgVol > 0 ? shares / d.avgVol : 0;
    return { ticker: d.cleanTicker, vwap: 0.1 * Math.sqrt(pctAdv) * 100, pctAdv };
  }).sort((a, b) => b.vwap - a.vwap);
  const avgVwap = vwapImpacts.length > 0 ? vwapImpacts.reduce((s, v) => s + v.vwap, 0) / vwapImpacts.length : 0;

  // 6. Sector rotation
  const industryBuySell = {};
  data.forEach(d => {
    if (!d.industry) return;
    if (!industryBuySell[d.industry]) industryBuySell[d.industry] = { buys: 0, sells: 0, netVal: 0 };
    if (d.direction === 'BUY') industryBuySell[d.industry].buys += d.netValue;
    else industryBuySell[d.industry].sells += Math.abs(d.netValue);
    industryBuySell[d.industry].netVal += d.netValue;
  });

  const sectorRotation = Object.entries(industryBuySell)
    .map(([name, v]) => ({ name, imbalance: v.netVal, buys: v.buys, sells: v.sells }))
    .sort((a, b) => b.imbalance - a.imbalance);

  // === RENDER TOP CARDS ===
  const topCards = document.getElementById('analytics-top');
  topCards.innerHTML = `
    <div class="section-card">
      <div class="section-header"><div class="section-title">Market Impact Score</div></div>
      <div class="section-body">
        <div class="metric-big">${fmtPct(avgImpact)}</div>
        <div class="metric-label">Avg Net Trade Value / ADV</div>
        <p style="color:var(--text-muted);font-size:0.75rem;margin-top:0.5rem;">
          Higher = more market impact. Top impact: <strong style="color:var(--cyan)">${sortedByAbsValue[0]?.cleanTicker}</strong>
          at ${fmtPct(sortedByAbsValue[0]?.avgVol && sortedByAbsValue[0]?.price ? Math.abs(sortedByAbsValue[0].netValue*1e6)/(sortedByAbsValue[0].price*sortedByAbsValue[0].avgVol) : 0)}
        </p>
      </div>
    </div>
    <div class="section-card">
      <div class="section-header"><div class="section-title">Concentration Risk</div></div>
      <div class="section-body">
        <div class="metric-big">${fmtPct(concentration)}</div>
        <div class="metric-label">Top 10 Names as % of Total Flow</div>
        <p style="color:var(--text-muted);font-size:0.75rem;margin-top:0.5rem;">
          Top 10 by absolute value: ${sortedByAbsValue.slice(0, 5).map(d => `<span style="color:var(--cyan)">${d.cleanTicker}</span>`).join(', ')}...
        </p>
      </div>
    </div>
    <div class="section-card">
      <div class="section-header"><div class="section-title">Buy/Sell Ratio</div></div>
      <div class="section-body">
        <div class="metric-big">${fmt(bsRatio, 2)}x</div>
        <div class="metric-label">${buys.length} Buys / ${sells.length} Sells</div>
      </div>
    </div>
    <div class="section-card">
      <div class="section-header"><div class="section-title">Estimated VWAP Impact</div></div>
      <div class="section-body">
        <div class="metric-big">${fmt(avgVwap, 1)} bps</div>
        <div class="metric-label">Avg Est. VWAP Impact (0.1 \u00d7 \u221a%ADV)</div>
        <p style="color:var(--text-muted);font-size:0.75rem;margin-top:0.5rem;">
          Highest: <strong style="color:var(--cyan)">${vwapImpacts[0]?.ticker}</strong> at ${fmt(vwapImpacts[0]?.vwap, 1)} bps
        </p>
      </div>
    </div>
  `;

  // === CHARTS ROW ===
  const chartsRow = document.getElementById('analytics-charts-row');
  chartsRow.innerHTML = `
    <div class="section-card">
      <div class="section-header"><div class="section-title">Liquidity Buckets</div></div>
      <div class="section-body" id="analytics-liquidity-buckets"></div>
    </div>
    <div class="section-card">
      <div class="section-header"><div class="section-title">Trade Value Distribution</div></div>
      <div class="section-body"><div class="chart-container" style="height:300px;"><canvas id="chart-trade-dist"></canvas></div></div>
    </div>
  `;

  // Liquidity buckets bars
  const maxBucket = Math.max(...Object.values(buckets));
  const bucketsContainer = document.getElementById('analytics-liquidity-buckets');
  const bucketColors = ['#00ff88', '#00d4ff', '#ffaa00', '#ff6600', '#ff3355'];
  bucketsContainer.innerHTML = Object.entries(buckets).map(([label, count], i) => `
    <div class="bucket-bar">
      <div class="bucket-label">${label}</div>
      <div class="bucket-fill-wrap">
        <div class="bucket-fill" style="width:${(count/maxBucket*100).toFixed(1)}%;background:${bucketColors[i]}">${count}</div>
      </div>
    </div>
  `).join('');

  // Trade value distribution histogram
  const values = data.map(d => d.netValue).filter(v => v !== 0);
  const minVal = Math.min(...values);
  const maxVal = Math.max(...values);
  const numBins = 30;
  const binWidth = (maxVal - minVal) / numBins;
  const bins = Array(numBins).fill(0);
  const binLabels = [];
  for (let i = 0; i < numBins; i++) {
    const lo = minVal + i * binWidth;
    binLabels.push(fmt(lo, 0));
    values.forEach(v => {
      if (v >= lo && v < lo + binWidth) bins[i]++;
    });
  }

  destroyChart('tradeDist');
  const ctx = document.getElementById('chart-trade-dist').getContext('2d');
  STATE.charts.tradeDist = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: binLabels,
      datasets: [{
        data: bins,
        backgroundColor: bins.map((_, i) => {
          const lo = minVal + i * binWidth;
          return lo >= 0 ? '#00ff8844' : '#ff335544';
        }),
        borderColor: bins.map((_, i) => {
          const lo = minVal + i * binWidth;
          return lo >= 0 ? '#00ff88' : '#ff3355';
        }),
        borderWidth: 1
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: { legend: { display: false } },
      scales: {
        x: {
          grid: { color: '#1e1e30' },
          ticks: { color: '#8888a0', font: { family: 'JetBrains Mono', size: 9 }, maxRotation: 45, autoSkip: true, maxTicksLimit: 10 }
        },
        y: {
          grid: { color: '#1e1e30' },
          ticks: { color: '#8888a0', font: { family: 'JetBrains Mono', size: 10 } },
          title: { display: true, text: 'Count', color: '#8888a0', font: { family: 'DM Sans', size: 11 } }
        }
      }
    }
  });

  // === HEATMAP ===
  renderHeatmap();

  // === BOTTOM SECTION ===
  const bottomCards = document.getElementById('analytics-bottom');

  const topBuySectors = sectorRotation.slice(0, 5);
  const topSellSectors = sectorRotation.slice(-5).reverse();

  // Buy/Sell ratio by Country (MSCI) or Index (S&P)
  const groupField = isSP ? 'index' : 'country';
  const groupLabel = isSP ? 'Index' : 'Country';
  const groupBS = {};
  data.forEach(d => {
    const key = d[groupField];
    if (!key) return;
    if (!groupBS[key]) groupBS[key] = { buys: 0, sells: 0 };
    if (d.direction === 'BUY') groupBS[key].buys++;
    else groupBS[key].sells++;
  });
  const groupRatios = Object.entries(groupBS)
    .map(([group, v]) => ({ group, ratio: v.sells > 0 ? v.buys / v.sells : v.buys, buys: v.buys, sells: v.sells }))
    .sort((a, b) => b.ratio - a.ratio);

  bottomCards.innerHTML = `
    <div class="section-card">
      <div class="section-header"><div class="section-title">Sector Rotation Signal</div></div>
      <div class="section-body">
        <h4 style="color:var(--green);font-family:var(--font-mono);font-size:0.75rem;margin-bottom:0.5rem;">BIGGEST NET BUYS (INFLOWS)</h4>
        <table class="data-table">
          <thead><tr><th>Industry</th><th class="num-cell">Net Flow</th><th class="num-cell">Buys</th><th class="num-cell">Sells</th></tr></thead>
          <tbody>${topBuySectors.map(s => `
            <tr>
              <td>${s.name}</td>
              <td class="num-cell green-text">${fmtMM(s.imbalance)}</td>
              <td class="num-cell green-text">${fmtMM(s.buys)}</td>
              <td class="num-cell red-text">${fmtMM(s.sells)}</td>
            </tr>
          `).join('')}</tbody>
        </table>
        <h4 style="color:var(--red);font-family:var(--font-mono);font-size:0.75rem;margin:1rem 0 0.5rem;">BIGGEST NET SELLS (OUTFLOWS)</h4>
        <table class="data-table">
          <thead><tr><th>Industry</th><th class="num-cell">Net Flow</th><th class="num-cell">Buys</th><th class="num-cell">Sells</th></tr></thead>
          <tbody>${topSellSectors.map(s => `
            <tr>
              <td>${s.name}</td>
              <td class="num-cell red-text">${fmtMM(s.imbalance)}</td>
              <td class="num-cell green-text">${fmtMM(s.buys)}</td>
              <td class="num-cell red-text">${fmtMM(s.sells)}</td>
            </tr>
          `).join('')}</tbody>
        </table>
      </div>
    </div>
    <div class="section-card">
      <div class="section-header"><div class="section-title">Buy/Sell Ratio by ${groupLabel}</div></div>
      <div class="section-body" style="max-height:500px;overflow-y:auto;">
        <table class="data-table">
          <thead><tr><th>${groupLabel}</th><th class="num-cell">Buys</th><th class="num-cell">Sells</th><th class="num-cell">B/S Ratio</th></tr></thead>
          <tbody>${groupRatios.slice(0, 20).map(c => `
            <tr>
              <td>${c.group}</td>
              <td class="num-cell green-text">${c.buys}</td>
              <td class="num-cell red-text">${c.sells}</td>
              <td class="num-cell" style="color:var(--cyan)">${fmt(c.ratio, 2)}x</td>
            </tr>
          `).join('')}</tbody>
        </table>
      </div>
    </div>
  `;
}

function renderHeatmap() {
  const data = STATE.overall;
  const container = document.getElementById('analytics-heatmap-body');
  const heatmapSection = document.getElementById('analytics-heatmap-section');
  const isSP = STATE.format === 'sp';

  // Update heatmap title
  const heatmapTitle = heatmapSection.querySelector('.section-title');
  heatmapTitle.textContent = isSP ? 'Industry x Index Heatmap (Net Flow)' : 'Industry x Country Heatmap (Net Flow)';

  // Get grouping dimension: country (MSCI) or index (S&P)
  const groupField = isSP ? 'index' : 'country';
  const groupCount = {};
  data.forEach(d => {
    const key = d[groupField];
    if (key) groupCount[key] = (groupCount[key] || 0) + 1;
  });
  const topGroups = Object.entries(groupCount).sort((a, b) => b[1] - a[1]).slice(0, 12).map(e => e[0]);

  const industries = [...new Set(data.map(d => d.industry).filter(Boolean))].sort();

  // Build matrix
  const matrix = {};
  let maxAbs = 0;
  industries.forEach(ind => {
    matrix[ind] = {};
    topGroups.forEach(g => { matrix[ind][g] = 0; });
  });

  data.forEach(d => {
    if (!d.industry || !topGroups.includes(d[groupField])) return;
    matrix[d.industry][d[groupField]] += d.netValue;
  });

  industries.forEach(ind => {
    topGroups.forEach(g => {
      maxAbs = Math.max(maxAbs, Math.abs(matrix[ind][g]));
    });
  });

  if (industries.length === 0 || topGroups.length === 0) {
    container.innerHTML = '<p style="color:var(--text-muted);">Insufficient data for heatmap.</p>';
    return;
  }

  let html = '<table class="data-table" style="font-size:0.65rem;">';
  html += '<thead><tr><th style="min-width:120px;">Industry</th>';
  topGroups.forEach(g => { html += `<th class="num-cell" style="min-width:60px;">${g}</th>`; });
  html += '</tr></thead><tbody>';

  industries.forEach(ind => {
    html += `<tr><td style="white-space:nowrap;">${ind}</td>`;
    topGroups.forEach(g => {
      const val = matrix[ind][g];
      const intensity = maxAbs > 0 ? Math.abs(val) / maxAbs : 0;
      let bg;
      if (val > 0) bg = `rgba(0,255,136,${0.08 + intensity * 0.4})`;
      else if (val < 0) bg = `rgba(255,51,85,${0.08 + intensity * 0.4})`;
      else bg = 'transparent';
      html += `<td class="num-cell" style="background:${bg};color:${val !== 0 ? 'var(--text-primary)' : 'var(--text-muted)'}">${val !== 0 ? fmt(val, 0) : '\u2014'}</td>`;
    });
    html += '</tr>';
  });

  html += '</tbody></table>';
  container.innerHTML = html;
}
