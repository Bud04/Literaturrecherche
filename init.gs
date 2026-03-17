// FILE: init.gs

/**
 * Initialisierung und Header-Wiederherstellung
 */

function initialSetup() {
  logAction("System", "Initialisierung gestartet");
  
  try {
    ensureSheetsAndHeaders();
    initializeConfigSheet();
    logAction("System", "Initialisierung erfolgreich abgeschlossen");
    SpreadsheetApp.getUi().alert("Initialisierung erfolgreich abgeschlossen!");
  } catch (e) {
    logError("initialSetup", e);
    SpreadsheetApp.getUi().alert("Fehler bei der Initialisierung: " + e.message);
  }
}

function ensureSheetsAndHeaders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetConfigs = [
    { name: DASHBOARD_SHEET_NAME, headers: DASHBOARD_HEADERS },
    { name: CONFIG_SHEET_NAME, headers: ["Schlüssel", "Wert"] },
    { name: KATEGORIEN_SHEET_NAME, headers: ["Hauptkategorie", "Unterkategorie", "Schlagwort"] },
    { name: MANUAL_PROMPTS_SHEET_NAME, headers: ["UUID", "Titel", "Phase", "Prompt 1", "Prompt 2", "Prompt 3", "Prompt Drive Link", "Antwort", "Timestamp"] },
    { name: FULLTEXT_SEARCH_SHEET_NAME, headers: ["UUID", "Titel", "Primary Link", "Kandidaten", "Volltext-Status", "Volltext-Link", "Timestamp"] },
    { name: MANUAL_FULLTEXT_SHEET_NAME, headers: ["UUID", "Titel", "Original Alert Link", "Primary Link", "PDF-Link/URL", "Status", "Notizen", "Timestamp"] },
    { name: MANUAL_IMPORT_SHEET_NAME, headers: ["Titel", "Autoren", "Jahr", "Journal", "DOI", "PMID", "Link", "Abstract", "Notizen", "Importiert"] },
    { name: EXPORT_HISTORY_SHEET_NAME, headers: ["UUID", "Titel", "Export-Datum", "Export-Typ", "Fingerprint", "RIS-File-Link"] },
    { name: ONEPAGER_HISTORY_SHEET_NAME, headers: ["UUID", "Titel", "Erstellt am", "Doc-Link", "Action", "Fehler"] },
    { name: ERROR_LOG_SHEET_NAME, headers: ["Timestamp", "Funktion", "Fehler", "Stack"] },
    { name: ERROR_LIST_SHEET_NAME, headers: ["Timestamp", "UUID", "Titel", "Fehlertyp", "Details", "Handlungsempfehlung", "Original Link"] },
    { name: LOGBOOK_SHEET_NAME, headers: ["Timestamp", "Aktion", "Details", "User"] },
    { name: IMPORT_REPORT_SHEET_NAME, headers: ["Timestamp", "Source", "Gmail Subject", "Gmail MessageId", "Gmail ThreadId", "Extracted Title", "Extracted Primary Link", "Parse Status", "Parse Reason", "Volltext Status", "Volltext Reason", "UUID", "Next Action"] }
  ];
  
  sheetConfigs.forEach(config => {
    let sheet = ss.getSheetByName(config.name);
    if (!sheet) {
      sheet = ss.insertSheet(config.name);
    }
    
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(config.headers);
      sheet.getRange(1, 1, 1, config.headers.length).setFontWeight("bold").setBackground("#4a86e8").setFontColor("#ffffff");
    } else {
      const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      let needsUpdate = false;
      
      if (currentHeaders.length !== config.headers.length) {
        needsUpdate = true;
      } else {
        for (let i = 0; i < config.headers.length; i++) {
          if (currentHeaders[i] !== config.headers[i]) {
            needsUpdate = true;
            break;
          }
        }
      }
      
      if (needsUpdate) {
        sheet.getRange(1, 1, 1, config.headers.length).setValues([config.headers]);
        sheet.getRange(1, 1, 1, config.headers.length).setFontWeight("bold").setBackground("#4a86e8").setFontColor("#ffffff");
      }
    }
  });
}

function restoreHeaders() {
  ensureSheetsAndHeaders();
  SpreadsheetApp.getUi().alert("Header wurden wiederhergestellt!");
  logAction("System", "Header wiederhergestellt");
}

function initializeConfigSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  if (sheet.getLastRow() <= 1) {
    Object.keys(CONFIG_DEFAULTS).forEach(key => {
      sheet.appendRow([key, CONFIG_DEFAULTS[key]]);
    });
  } else {
    const existingKeys = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    Object.keys(CONFIG_DEFAULTS).forEach(key => {
      if (!existingKeys.includes(key)) {
        sheet.appendRow([key, CONFIG_DEFAULTS[key]]);
      }
    });
  }
}

function selfCheck() {
  const results = [];
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  results.push("=== SELF CHECK ===\n");
  
  const requiredSheets = [
    DASHBOARD_SHEET_NAME, CONFIG_SHEET_NAME, KATEGORIEN_SHEET_NAME,
    MANUAL_PROMPTS_SHEET_NAME, FULLTEXT_SEARCH_SHEET_NAME, MANUAL_FULLTEXT_SHEET_NAME,
    MANUAL_IMPORT_SHEET_NAME, EXPORT_HISTORY_SHEET_NAME, ONEPAGER_HISTORY_SHEET_NAME,
    ERROR_LOG_SHEET_NAME, ERROR_LIST_SHEET_NAME, LOGBOOK_SHEET_NAME, IMPORT_REPORT_SHEET_NAME
  ];
  
  results.push("Sheets:");
  requiredSheets.forEach(name => {
    const exists = ss.getSheetByName(name) !== null;
    results.push(`  ${name}: ${exists ? "✓" : "✗ FEHLT"}`);
  });
  
  results.push("\nKonfiguration:");
  const requiredKeys = [
    "Gmail Label Name", "Onepager Hauptordner ID", "PDF Hauptordner ID",
    "Citavi Export Ordner ID", "PROMPT_TEMPLATE_ANALYSIS"
  ];
  requiredKeys.forEach(key => {
    const value = getConfig(key);
    results.push(`  ${key}: ${value ? "✓" : "✗ FEHLT"}`);
  });
  
  results.push("\nGmail Label:");
  const labelName = getConfig("Gmail Label Name");
  if (labelName) {
    const label = GmailApp.getUserLabelByName(labelName.replace(/^label:/, ""));
    results.push(`  ${labelName}: ${label ? "✓" : "✗ NICHT GEFUNDEN"}`);
  }
  
  results.push("\nDrive Ordner:");
  const driveIds = [
    ["Onepager Hauptordner ID", getConfig("Onepager Hauptordner ID")],
    ["PDF Hauptordner ID", getConfig("PDF Hauptordner ID")],
    ["Citavi Export Ordner ID", getConfig("Citavi Export Ordner ID")]
  ];
  driveIds.forEach(([name, id]) => {
    if (id) {
      try {
        const folder = DriveApp.getFolderById(id);
        results.push(`  ${name}: ✓ ${folder.getName()}`);
      } catch (e) {
        results.push(`  ${name}: ✗ NICHT ZUGREIFBAR`);
      }
    }
  });
  
  const message = results.join("\n");
  Logger.log(message);
  SpreadsheetApp.getUi().alert(message);
  logAction("System", "Self-Check durchgeführt");
}
