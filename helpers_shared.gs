// FILE: helpers_shared.gs
// ==========================================
// SHARED HELPER FUNCTIONS
// Einmalige Definitionen - überall nutzbar
//
// ENTHÄLT (aus Duplikaten bereinigt):
//   - buildAnalysisJsonForReview()   [war: batchPipeline.gs + geminiSheets.gs]
//   - getRowDataByIndex()            [war: batchPipeline.gs + geminiSheets.gs]
//   - getAllPaperUUIDs()             [war: Claude.gs + workspace_flow_core.gs]
//   - getConfig()                    [war: prompt_automation.gs - undefiniert]
//   - writeResultsToSheetExtended()  [war: sidebar_backend.gs - undefiniert]
//   - setCell()                      [war: sidebar_backend.gs - undefiniert]
//   - getRowData()                   [war: sidebar_backend.gs - undefiniert]
//   - getCategoriesAndTags()         [war: sidebar_backend.gs - undefiniert]
//   - applyDashboardFilters()        [war: Cockpit.html callback - undefiniert]
//   - clearDashboardFilters()        [war: Cockpit.html callback - undefiniert]
// ==========================================

// ==========================================
// ANALYSE-JSON FÜR REVIEW (bisher 2× definiert)
// ==========================================

/**
 * ✅ Baut das Analyse-JSON für den Review-Prompt
 * Ersetzt Duplikate in batchPipeline.gs und geminiSheets.gs
 */
function buildAnalysisJsonForReview(rowData) {
  return JSON.stringify({
    relevanz: rowData.Relevanz,
    produkt_fokus: rowData['Produkt-Fokus'],
    kategorie: {
      hauptkategorie: rowData.Hauptkategorie,
      unterkategorien: (rowData.Unterkategorien || '').split(', ')
    },
    haupterkenntnis: rowData.Haupterkenntnis,
    kernaussagen: (rowData.Kernaussagen || '').split('\n').filter(k => k.trim()),
    zusammenfassung: rowData.Zusammenfassung,
    schlagwoerter: (rowData.Schlagwörter || '').split(', '),
    kritische_bewertung: rowData['Kritische Bewertung']
  }, null, 2);
}

// ==========================================
// ROW DATA HELPERS (bisher 2× definiert)
// ==========================================

/**
 * ✅ Liest eine Dashboard-Zeile als Key-Value-Objekt
 * Ersetzt Duplikate in batchPipeline.gs und geminiSheets.gs
 */
function getRowDataByIndex(sheet, rowIndex) {
  const rowValues = sheet
    .getRange(rowIndex, 1, 1, DASHBOARD_HEADERS.length)
    .getValues()[0];

  const data = {};
  DASHBOARD_HEADERS.forEach((header, index) => {
    data[header] = rowValues[index] !== undefined ? rowValues[index] : '';
  });
  return data;
}

/**
 * ✅ Alias für sidebar_backend.gs Kompatibilität
 * sidebar_backend.gs nutzt getRowData(sheet, rowIndex) ohne DASHBOARD_HEADERS
 */
function getRowData(sheet, rowIndex) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const values = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

  const data = {};
  headers.forEach((header, i) => {
    data[header] = values[i] !== undefined ? values[i] : '';
  });
  return data;
}

// ==========================================
// UUID & PAPER HELPERS (bisher 2× definiert)
// ==========================================

/**
 * ✅ Holt alle Paper-UUIDs aus dem Dashboard
 * Ersetzt Duplikate in Claude.gs und workspace_flow_core.gs
 */
function getAllPaperUUIDs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('Dashboard');

  if (!dashboard || dashboard.getLastRow() <= 1) return [];

  return dashboard
    .getRange(2, 1, dashboard.getLastRow() - 1, 1)
    .getValues()
    .map(row => row[0])
    .filter(uuid => uuid && uuid !== '');
}

// ==========================================
// CONFIG HELPER
// ==========================================

/**
 * ✅ Liest einen Wert aus dem Konfiguration-Sheet
 * War in prompt_automation.gs als getConfig() referenziert aber undefiniert
 */
function getConfig(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);

  if (!configSheet) {
    Logger.log(`getConfig: Sheet '${CONFIG_SHEET_NAME}' nicht gefunden`);
    return null;
  }

  // Bekannte Keys → direkte Zellen-Referenz
  const cellMap = {
    PROMPT_TEMPLATE_RESEARCHER: 'B15',
    PROMPT_TEMPLATE_ANALYSIS: 'B16',
    PROMPT_TEMPLATE_METADATA_EXTRACT: 'B17',
    PROMPT_TEMPLATE_REDAKTION: 'B18',
    PROMPT_TEMPLATE_TRIAGE: 'B19',
    PROMPT_TEMPLATE_FAKTENCHECK: 'B20',
    PROMPT_TEMPLATE_REVIEW: 'B21'
  };

  if (cellMap[key]) {
    return configSheet.getRange(cellMap[key]).getValue() || null;
  }

  // Suche in Spalte A/B (Key-Value Tabelle)
  const data = configSheet.getRange('A:B').getValues();
  for (const row of data) {
    if (String(row[0]).trim() === key) return row[1] || null;
  }

  return null;
}

// ==========================================
// SIDEBAR / SHEET WRITE HELPERS
// ==========================================

/**
 * ✅ Schreibt einen Wert in eine benannte Spalte der aktiven Zeile
 * War in sidebar_backend.gs als setCell() referenziert aber undefiniert
 */
function setCell(sheet, row, headerName, value) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(headerName);

  if (colIndex >= 0) {
    sheet.getRange(row, colIndex + 1).setValue(value);
  } else {
    Logger.log(`setCell: Header '${headerName}' nicht gefunden`);
  }
}

/**
 * ✅ Schreibt strukturierte Analyse-Ergebnisse ins Sheet
 * War in sidebar_backend.gs als writeResultsToSheetExtended() referenziert
 * Nutzt applyJSONToDashboardPhaseAware() aus geminiSheets.gs
 */
function writeResultsToSheetExtended(sheet, row, data) {
  if (!data || typeof data !== 'object') {
    throw new Error('Ungültige Daten für writeResultsToSheetExtended');
  }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const uuidColIndex = headers.indexOf('UUID');

  if (uuidColIndex < 0) {
    throw new Error('UUID-Spalte nicht gefunden');
  }

  const uuid = sheet.getRange(row, uuidColIndex + 1).getValue();

  if (!uuid) {
    throw new Error(`Keine UUID in Zeile ${row}`);
  }

  // Bestimme Phase aus Data-Struktur
  let phase = 'MASTER_ANALYSE'; // Default

  if (data.hauptkategorie || data.kategorie) phase = 'TRIAGE';
  else if (data.pmid || data.doi || data.artikeltyp) phase = 'METADATEN';
  else if (data.is_supported !== undefined) phase = 'FAKTENCHECK';
  else if (data.ist_korrekt !== undefined) phase = 'REVIEW';
  else if (data.wichtige_abbildungen) phase = 'REDAKTION';

  // Nutzt Funktion aus geminiSheets.gs
  applyJSONToDashboardPhaseAware(uuid, phase, data);

  // Zeitstempel
  setCell(sheet, row, 'Letzte Änderung', new Date());
}

// ==========================================
// KATEGORIE HELPER (Alias für sidebar_backend.gs)
// ==========================================

/**
 * ✅ Alias: getCategoriesAndTags() → loadCategoriesAndKeywords()
 * War in sidebar_backend.gs als getCategoriesAndTags() referenziert
 * Die eigentliche Funktion heißt loadCategoriesAndKeywords() in utils.gs
 */
function getCategoriesAndTags() {
  const lists = loadCategoriesAndKeywords(); // aus utils.gs

  return {
    categories: lists.categories.join(', '),
    tags: lists.keywords.join(', '),
    fullMapping: lists.fullMapping
  };
}

// ==========================================
// DASHBOARD FILTER HELPERS
// ==========================================

/**
 * ✅ Wendet Filter auf das Dashboard an
 * Wird von Cockpit.html via google.script.run aufgerufen
 */
function applyDashboardFilters(criteria) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  if (!sheet || sheet.getLastRow() <= 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const statusCol = headers.indexOf('Status');
  const relevanzCol = headers.indexOf('Relevanz');
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  // Zeige erst alle
  sheet.showRows(2, sheet.getLastRow() - 1);

  let hiddenCount = 0;

  data.forEach((row, index) => {
    const rowIndex = index + 2;
    let hide = false;

    if (criteria.status && statusCol >= 0) {
      if (String(row[statusCol]).trim() !== criteria.status) hide = true;
    }

    if (criteria.relevance && relevanzCol >= 0) {
      if (String(row[relevanzCol]).trim() !== criteria.relevance) hide = true;
    }

    if (hide) {
      sheet.hideRows(rowIndex);
      hiddenCount++;
    }
  });

  logAction('Dashboard Filter', `${hiddenCount} Zeilen ausgeblendet`);
}

/**
 * ✅ Setzt Dashboard-Filter zurück
 * Wird von Cockpit.html via google.script.run aufgerufen
 */
function clearDashboardFilters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);

  if (!sheet || sheet.getLastRow() <= 1) return;

  sheet.showRows(1, sheet.getMaxRows());
  logAction('Dashboard Filter', 'Alle Filter zurückgesetzt');
}
