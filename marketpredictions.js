/**
 * MarketPredictions.gs
 * Coffee & Cocoa Market Intelligence Dashboard for CSS
 * Based on market data as of February 14, 2026
 * 
 * Creates a "Market Intel" sheet with current prices, trends,
 * predictions, and CSS volume impact estimates.
 */

// ============================================================
// CONFIGURATION
// ============================================================

const MARKET_CONFIG = {
  sheetName: 'Market Intel',
  lastUpdated: '2026-02-14',
  
  // Current market snapshot (Feb 14, 2026)
  currentPrices: {
    arabica: { price: 2.96, unit: '$/lb', contract: 'KCH26', exchange: 'ICE' },
    robusta: { price: 3.81, unit: '$/kg', contract: 'RMH26', exchange: 'ICE London' },
    cocoa_ny: { price: 3581, unit: '$/MT', contract: 'CCH26', exchange: 'ICE NY' },
    cocoa_london: { price: 2570, unit: '£/MT', contract: 'CAH26', exchange: 'ICE London' }
  },
  
  // 52-week ranges
  ranges52w: {
    arabica: { low: 2.77, high: 4.38 },
    robusta: { low: 3.20, high: 5.80 },
    cocoa_ny: { low: 3563, high: 11280 },
  },
  
  // Key fundamentals
  fundamentals: {
    brazil2026Crop: { bags: 66.2, unit: 'M bags', yoyChange: '+17.2%', source: 'CONAB' },
    brazil2026CropAlt: { bags: 75.8, unit: 'M bags', source: 'Eisa Brokerage' },
    arabicaProduction: { bags: 44.1, unit: 'M bags', yoyChange: '+23.2%' },
    robustaProduction: { bags: 22.1, unit: 'M bags', yoyChange: '+6.3%' },
    vietnamRobusta: { bags: 29.4, unit: 'M bags', note: 'projected 2025/26' },
    cocoaSurplus2526: { tons: 287000, unit: 'MT', source: 'StoneX' },
    cocoaSurplus2627: { tons: 267000, unit: 'MT', source: 'StoneX' },
    cocoaGlobalStocks: { tons: 1.1, unit: 'MMT', yoyChange: '+4.2%', source: 'ICCO' },
    ivoryCoastArrivals: { tons: 1.263, unit: 'MMT', yoyChange: '-4.5%', note: 'through Feb 8' }
  }
};

// ============================================================
// PREDICTIONS ENGINE
// ============================================================

/**
 * Market predictions with confidence levels and CSS impact
 */
function getMarketPredictions_() {
  return [
    // COFFEE PREDICTIONS
    {
      commodity: 'ARABICA COFFEE',
      timeframe: 'Q1 2026 (Now → Mar)',
      prediction: 'Continued downward pressure toward $2.50-2.80/lb',
      confidence: 'HIGH',
      keyDrivers: 'Record Brazil crop (66-76M bags), favorable weather, 6-month price lows',
      riskFactors: 'Real strength causing short covering, roaster buying at lows',
      cssImpact: 'VOLUME UP — Lower prices = more imports clearing, more sampling orders',
      priceRange: '$2.50 - $3.20'
    },
    {
      commodity: 'ARABICA COFFEE',
      timeframe: 'Q2-Q3 2026 (Apr → Sep)',
      prediction: 'Stabilization with potential bounce as Brazil harvest begins',
      confidence: 'MEDIUM',
      keyDrivers: 'World Bank projects -13% for full year 2026, harvest execution risk',
      riskFactors: 'Weather disruption during flowering, frost risk Jun-Aug',
      cssImpact: 'STEADY — Normal seasonal flow, watch for rush orders if frost scares',
      priceRange: '$2.40 - $3.00'
    },
    {
      commodity: 'ARABICA COFFEE',
      timeframe: '2027 Outlook',
      prediction: 'Modest further decline, -5% from 2026 average',
      confidence: 'LOW',
      keyDrivers: 'Two consecutive surplus years projected, inventory rebuild',
      riskFactors: 'Climate change wildcards, biennial cycle could reduce 2027 Brazil crop',
      cssImpact: 'VOLUME UP — Cheaper coffee = more trading activity = more samples',
      priceRange: '$2.30 - $2.80'
    },
    {
      commodity: 'ROBUSTA COFFEE',
      timeframe: 'Q1-Q2 2026',
      prediction: 'Gradual decline as Vietnam production recovers',
      confidence: 'MEDIUM',
      keyDrivers: 'Vietnam exports +17.5% in 2025, projected 29.4M bags 2025/26',
      riskFactors: 'Vietnam weather during harvest, Arabica-Robusta spread normalization',
      cssImpact: 'MODERATE — Robusta sampling volume typically lower for CSS',
      priceRange: '$3.20 - $4.00/kg'
    },
    
    // COCOA PREDICTIONS
    {
      commodity: 'COCOA',
      timeframe: 'Q1-Q2 2026',
      prediction: 'Sharp correction continues — approaching pre-crisis levels',
      confidence: 'HIGH',
      keyDrivers: '287K MT surplus projected, 65% price drop YoY, demand destruction, unsold inventory in Ivory Coast & Ghana (50K MT)',
      riskFactors: 'Short squeeze risk (heavy spec short positioning), any West Africa supply shock',
      cssImpact: 'COCOA SAMPLING DOWN — Major buyers hedged at higher prices, less spot activity',
      priceRange: '$3,000 - $4,200/MT'
    },
    {
      commodity: 'COCOA',
      timeframe: 'H2 2026 → 2027',
      prediction: 'Stabilization at structurally higher floor vs pre-2023',
      confidence: 'MEDIUM',
      keyDrivers: 'Consecutive surpluses (287K then 267K MT), but aging trees and climate risks persist',
      riskFactors: 'West Africa disease outbreaks, structural underinvestment in plantations',
      cssImpact: 'RECOVERY — Lower prices should stimulate demand, cocoa sampling normalizes',
      priceRange: '$3,500 - $5,000/MT'
    },
    
    // INDUSTRY / CSS-SPECIFIC
    {
      commodity: 'WAREHOUSE CONSOLIDATION',
      timeframe: '2026-2027',
      prediction: 'RPM/Continental merger leads to fee increases and service bundling pressure',
      confidence: 'HIGH',
      keyDrivers: 'Reduced competition in NY/NJ corridor, vertical integration trend',
      riskFactors: 'New entrants, customer pushback, regulatory scrutiny',
      cssImpact: 'OPPORTUNITY — Independent sampling becomes more valuable as warehouses consolidate',
      priceRange: 'N/A'
    },
    {
      commodity: 'CSS VOLUME OUTLOOK',
      timeframe: 'Full Year 2026',
      prediction: 'Net positive — coffee volume up, cocoa volume temporarily soft',
      confidence: 'MEDIUM',
      keyDrivers: 'Record Brazil crop = record imports through NY/NJ, lower prices boost trading activity',
      riskFactors: 'Tariff policy changes, warehouse access changes post-merger',
      cssImpact: 'PLAN FOR +10-15% coffee sampling volume, -5-10% cocoa through H1',
      priceRange: 'N/A'
    }
  ];
}

// ============================================================
// SHEET BUILDER
// ============================================================

/**
 * Main entry point — creates or refreshes the Market Intel sheet
 */
function buildMarketIntelSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(MARKET_CONFIG.sheetName);
  
  if (sheet) {
    sheet.clear();
  } else {
    sheet = ss.insertSheet(MARKET_CONFIG.sheetName);
  }
  
  let row = 1;
  
  // ---- HEADER ----
  row = writeHeader_(sheet, row);
  
  // ---- CURRENT PRICES ----
  row = writePriceSnapshot_(sheet, row);
  
  // ---- FUNDAMENTALS ----
  row = writeFundamentals_(sheet, row);
  
  // ---- PREDICTIONS TABLE ----
  row = writePredictions_(sheet, row);
  
  // ---- CSS IMPACT SUMMARY ----
  row = writeCSSImpact_(sheet, row);
  
  // ---- FORMATTING ----
  formatSheet_(sheet, row);
  
  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Market Intel updated for ' + MARKET_CONFIG.lastUpdated, 
    'Market Predictions', 
    5
  );
}

// ============================================================
// SECTION WRITERS
// ============================================================

function writeHeader_(sheet, startRow) {
  const r = startRow;
  sheet.getRange(r, 1).setValue('☕ CSS MARKET INTELLIGENCE DASHBOARD');
  sheet.getRange(r, 1, 1, 8).merge().setBackground('#1a1a2e').setFontColor('#ffffff')
    .setFontSize(16).setFontWeight('bold').setHorizontalAlignment('center');
  
  sheet.getRange(r + 1, 1).setValue('Last Updated: ' + MARKET_CONFIG.lastUpdated + '  |  Disclaimer: Predictions only — not financial advice');
  sheet.getRange(r + 1, 1, 1, 8).merge().setBackground('#16213e').setFontColor('#a0a0a0')
    .setFontSize(9).setHorizontalAlignment('center').setFontStyle('italic');
  
  return r + 3;
}

function writePriceSnapshot_(sheet, startRow) {
  let r = startRow;
  
  // Section header
  sheet.getRange(r, 1).setValue('📊 CURRENT PRICE SNAPSHOT');
  sheet.getRange(r, 1, 1, 8).merge().setBackground('#0f3460').setFontColor('#ffffff')
    .setFontSize(12).setFontWeight('bold');
  r++;
  
  // Column headers
  const headers = ['Commodity', 'Price', 'Unit', 'Contract', 'Exchange', '52W Low', '52W High', 'Position in Range'];
  sheet.getRange(r, 1, 1, 8).setValues([headers])
    .setBackground('#1a1a2e').setFontColor('#e0e0e0').setFontWeight('bold').setFontSize(9);
  r++;
  
  const prices = MARKET_CONFIG.currentPrices;
  const ranges = MARKET_CONFIG.ranges52w;
  
  const priceRows = [
    ['Arabica Coffee', prices.arabica.price, prices.arabica.unit, prices.arabica.contract, 
     prices.arabica.exchange, ranges.arabica.low, ranges.arabica.high,
     getPositionInRange_(prices.arabica.price, ranges.arabica.low, ranges.arabica.high)],
    ['Robusta Coffee', prices.robusta.price, prices.robusta.unit, prices.robusta.contract,
     prices.robusta.exchange, '', '', ''],
    ['Cocoa (NY)', prices.cocoa_ny.price, prices.cocoa_ny.unit, prices.cocoa_ny.contract,
     prices.cocoa_ny.exchange, ranges.cocoa_ny.low, ranges.cocoa_ny.high,
     getPositionInRange_(prices.cocoa_ny.price, ranges.cocoa_ny.low, ranges.cocoa_ny.high)],
    ['Cocoa (London)', prices.cocoa_london.price, prices.cocoa_london.unit, prices.cocoa_london.contract,
     prices.cocoa_london.exchange, '', '', '']
  ];
  
  sheet.getRange(r, 1, priceRows.length, 8).setValues(priceRows);
  
  // Color-code position in range
  for (let i = 0; i < priceRows.length; i++) {
    const pos = priceRows[i][7];
    if (pos && typeof pos === 'string') {
      const cell = sheet.getRange(r + i, 8);
      if (pos.includes('LOW')) {
        cell.setBackground('#1b4332').setFontColor('#95d5b2');
      } else if (pos.includes('HIGH')) {
        cell.setBackground('#6b2020').setFontColor('#f5a0a0');
      } else {
        cell.setBackground('#4a3800').setFontColor('#ffd166');
      }
      cell.setFontWeight('bold');
    }
  }
  
  r += priceRows.length + 1;
  return r;
}

function writeFundamentals_(sheet, startRow) {
  let r = startRow;
  
  sheet.getRange(r, 1).setValue('🌍 KEY FUNDAMENTALS');
  sheet.getRange(r, 1, 1, 8).merge().setBackground('#0f3460').setFontColor('#ffffff')
    .setFontSize(12).setFontWeight('bold');
  r++;
  
  const fund = MARKET_CONFIG.fundamentals;
  
  const fundHeaders = ['Metric', 'Value', 'Unit', 'YoY Change', 'Source', 'Signal', '', ''];
  sheet.getRange(r, 1, 1, 8).setValues([fundHeaders])
    .setBackground('#1a1a2e').setFontColor('#e0e0e0').setFontWeight('bold').setFontSize(9);
  r++;
  
  const fundRows = [
    ['Brazil 2026 Total Crop (CONAB)', fund.brazil2026Crop.bags, fund.brazil2026Crop.unit, 
     fund.brazil2026Crop.yoyChange, fund.brazil2026Crop.source, '🐻 BEARISH (record supply)', '', ''],
    ['Brazil 2026 Total Crop (Eisa)', fund.brazil2026CropAlt.bags, fund.brazil2026CropAlt.unit, 
     '', fund.brazil2026CropAlt.source, '🐻 VERY BEARISH', '', ''],
    ['Brazil Arabica Production', fund.arabicaProduction.bags, fund.arabicaProduction.unit,
     fund.arabicaProduction.yoyChange, 'CONAB', '🐻 BEARISH', '', ''],
    ['Brazil Robusta Production', fund.robustaProduction.bags, fund.robustaProduction.unit,
     fund.robustaProduction.yoyChange, 'CONAB', '🐻 Mildly Bearish', '', ''],
    ['Vietnam Robusta (projected)', fund.vietnamRobusta.bags, fund.vietnamRobusta.unit,
     '+6%', 'Vietnam NSO', '🐻 Bearish for Robusta', '', ''],
    ['Cocoa Global Surplus 25/26', formatNumber_(fund.cocoaSurplus2526.tons), fund.cocoaSurplus2526.unit,
     '', fund.cocoaSurplus2526.source, '🐻 VERY BEARISH', '', ''],
    ['Cocoa Global Surplus 26/27', formatNumber_(fund.cocoaSurplus2627.tons), fund.cocoaSurplus2627.unit,
     '', fund.cocoaSurplus2627.source, '🐻 BEARISH', '', ''],
    ['Cocoa Global Stocks', fund.cocoaGlobalStocks.tons, fund.cocoaGlobalStocks.unit,
     fund.cocoaGlobalStocks.yoyChange, fund.cocoaGlobalStocks.source, '🐻 Bearish', '', ''],
    ['Ivory Coast Arrivals (to Feb 8)', fund.ivoryCoastArrivals.tons, fund.ivoryCoastArrivals.unit,
     fund.ivoryCoastArrivals.yoyChange, 'Port Data', '🐂 Slightly Bullish', '', ''],
  ];
  
  sheet.getRange(r, 1, fundRows.length, 8).setValues(fundRows);
  
  // Color the signal column
  for (let i = 0; i < fundRows.length; i++) {
    const signal = fundRows[i][5];
    const cell = sheet.getRange(r + i, 6);
    if (signal.includes('VERY BEARISH')) {
      cell.setBackground('#6b2020').setFontColor('#f5a0a0');
    } else if (signal.includes('BEARISH') || signal.includes('Bearish')) {
      cell.setBackground('#4a2020').setFontColor('#e0a0a0');
    } else if (signal.includes('Bullish')) {
      cell.setBackground('#1b4332').setFontColor('#95d5b2');
    }
    cell.setFontWeight('bold').setFontSize(9);
  }
  
  r += fundRows.length + 1;
  return r;
}

function writePredictions_(sheet, startRow) {
  let r = startRow;
  const predictions = getMarketPredictions_();
  
  sheet.getRange(r, 1).setValue('🔮 MARKET PREDICTIONS & CSS IMPACT');
  sheet.getRange(r, 1, 1, 8).merge().setBackground('#0f3460').setFontColor('#ffffff')
    .setFontSize(12).setFontWeight('bold');
  r++;
  
  const predHeaders = ['Commodity', 'Timeframe', 'Prediction', 'Confidence', 'Key Drivers', 'Risk Factors', 'CSS Impact', 'Price Range'];
  sheet.getRange(r, 1, 1, 8).setValues([predHeaders])
    .setBackground('#1a1a2e').setFontColor('#e0e0e0').setFontWeight('bold').setFontSize(9);
  r++;
  
  predictions.forEach(function(p) {
    const row = [p.commodity, p.timeframe, p.prediction, p.confidence, 
                 p.keyDrivers, p.riskFactors, p.cssImpact, p.priceRange];
    sheet.getRange(r, 1, 1, 8).setValues([row]);
    
    // Color confidence
    const confCell = sheet.getRange(r, 4);
    if (p.confidence === 'HIGH') {
      confCell.setBackground('#1b4332').setFontColor('#95d5b2');
    } else if (p.confidence === 'MEDIUM') {
      confCell.setBackground('#4a3800').setFontColor('#ffd166');
    } else {
      confCell.setBackground('#4a2020').setFontColor('#e0a0a0');
    }
    confCell.setFontWeight('bold').setHorizontalAlignment('center');
    
    // Color CSS impact
    const impactCell = sheet.getRange(r, 7);
    const impact = p.cssImpact.toUpperCase();
    if (impact.includes('VOLUME UP') || impact.includes('OPPORTUNITY') || impact.includes('+10')) {
      impactCell.setBackground('#1b4332').setFontColor('#95d5b2');
    } else if (impact.includes('DOWN')) {
      impactCell.setBackground('#4a2020').setFontColor('#e0a0a0');
    } else {
      impactCell.setBackground('#2a2a3a').setFontColor('#c0c0e0');
    }
    impactCell.setFontWeight('bold').setFontSize(9);
    
    r++;
  });
  
  r++;
  return r;
}

function writeCSSImpact_(sheet, startRow) {
  let r = startRow;
  
  sheet.getRange(r, 1).setValue('🎯 CSS 2026 OPERATIONAL OUTLOOK');
  sheet.getRange(r, 1, 1, 8).merge().setBackground('#0f3460').setFontColor('#ffffff')
    .setFontSize(12).setFontWeight('bold');
  r++;
  
  const outlook = [
    ['COFFEE SAMPLING VOLUME', 'UP 10-15%', 
     'Record Brazil crop means record imports through NY/NJ. Lower prices stimulate trading. Expect heavier flow Q2-Q3 as harvest ships.',
     'Watch for: Tariff changes on Brazilian coffee (lifted recently), warehouse access changes post RPM/Continental merger', '', '', '', ''],
    ['COCOA SAMPLING VOLUME', 'DOWN 5-10% (H1), RECOVERY H2',
     'Demand destruction from 2024-25 price spike still rippling. Majors hedged at high prices, less spot activity. 50K MT unsold in Ghana ports.',
     'Watch for: Short squeeze potential (heavy spec shorts), Ivory Coast buyback program impact on flows', '', '', '', ''],
    ['WAREHOUSE DYNAMICS', 'DISRUPTION RISK',
     'RPM/Continental merger consolidates major NY/NJ corridor. Expect procedure changes, potential fee increases, possible push toward in-house services.',
     'Action: Strengthen relationships at both legacy operations. Independence value prop becomes stronger.', '', '', '', ''],
    ['PRICING STRATEGY', 'HOLD RATES',
     'Higher coffee volume offsets softer cocoa. No need for rate cuts. Independence premium justified by consolidation.',
     'Consider: Volume-based incentives for top accounts to lock in business ahead of warehouse changes', '', '', '', ''],
    ['KEY DATES TO WATCH', '—',
     'Apr-Sep: Brazil flowering/fruit dev (weather critical) | Jun-Aug: Frost risk window | Oct: New Ivory Coast cocoa season | Q2: RPM/Continental integration milestones',
     '', '', '', '', '']
  ];
  
  const olHeaders = ['Category', 'Direction', 'Analysis', 'Action Items', '', '', '', ''];
  sheet.getRange(r, 1, 1, 8).setValues([olHeaders])
    .setBackground('#1a1a2e').setFontColor('#e0e0e0').setFontWeight('bold').setFontSize(9);
  r++;
  
  outlook.forEach(function(row) {
    sheet.getRange(r, 1, 1, 8).setValues([row]);
    
    // Color direction
    const dirCell = sheet.getRange(r, 2);
    const dir = row[1].toUpperCase();
    if (dir.includes('UP')) {
      dirCell.setBackground('#1b4332').setFontColor('#95d5b2');
    } else if (dir.includes('DOWN')) {
      dirCell.setBackground('#4a2020').setFontColor('#e0a0a0');
    } else if (dir.includes('DISRUPTION')) {
      dirCell.setBackground('#4a3800').setFontColor('#ffd166');
    } else if (dir.includes('HOLD')) {
      dirCell.setBackground('#1b4332').setFontColor('#95d5b2');
    }
    dirCell.setFontWeight('bold');
    
    r++;
  });
  
  return r + 1;
}

// ============================================================
// UTILITIES
// ============================================================

function getPositionInRange_(current, low, high) {
  if (!low || !high) return '';
  const pct = ((current - low) / (high - low) * 100).toFixed(0);
  if (pct <= 20) return '▼ LOW (' + pct + '%)';
  if (pct >= 80) return '▲ HIGH (' + pct + '%)';
  return '● MID (' + pct + '%)';
}

function formatNumber_(num) {
  return num.toLocaleString ? num.toLocaleString() : num.toString();
}

function formatSheet_(sheet, lastRow) {
  // Column widths
  sheet.setColumnWidth(1, 200);  // Commodity/Category
  sheet.setColumnWidth(2, 150);  // Price/Direction
  sheet.setColumnWidth(3, 280);  // Prediction/Analysis
  sheet.setColumnWidth(4, 100);  // Confidence
  sheet.setColumnWidth(5, 280);  // Key Drivers
  sheet.setColumnWidth(6, 250);  // Risk/Signal
  sheet.setColumnWidth(7, 280);  // CSS Impact
  sheet.setColumnWidth(8, 120);  // Price Range
  
  // Global styling
  const allRange = sheet.getRange(1, 1, lastRow, 8);
  allRange.setFontFamily('Roboto Mono');
  allRange.setVerticalAlignment('top');
  allRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  
  // Data rows — dark theme base
  sheet.getRange(1, 1, lastRow, 8).setBackground('#0d1117');
  sheet.getRange(1, 1, lastRow, 8).setFontColor('#c9d1d9');
  sheet.getRange(1, 1, lastRow, 8).setFontSize(10);
  
  // Borders
  allRange.setBorder(true, true, true, true, true, true, '#30363d', SpreadsheetApp.BorderStyle.SOLID);
  
  // Freeze header
  sheet.setFrozenRows(2);
  
  // Tab color
  sheet.setTabColor('#0f3460');
}

