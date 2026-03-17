// FILE: ui.gs
/**
 * ==========================================
 * UI & MENÜ
 * ==========================================
 * Nur: onOpen(), Menü-Callbacks, Initialisierung
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📚 Literatur-Workflow')
    .addItem('🚀 Cockpit Sidebar öffnen', 'showSidebar')
    .addItem('🎨 Markierung (Farbe) entfernen', 'resetCurrentRowColor')
    .addSeparator()
    
    .addSubMenu(ui.createMenu('⚙️ Batch-Prozesse')
      .addItem('Nächste relevante Zeile finden', 'jumpToNextRelevant')
      .addItem('Volltext-Booster (Markierte Zeilen)', 'runBoosterForSelected')
      .addSeparator()
      .addItem('📊 Batch Progress anzeigen', 'showBatchProgress')
      .addItem('🎨 Batch-Phase Farben aktualisieren', 'updateBatchPhaseColors')
      .addItem('🔍 Filter: Nach Batch-Phase', 'filterByBatchPhase')
      .addItem('🔄 Batch-Phase zurücksetzen (Auswahl)', 'resetBatchPhaseForSelected'))
    
    .addSeparator()
    
    .addSubMenu(ui.createMenu('📥 Import')
      .addItem('🌟 ALLE THEMEN (Smart Import)', 'importAllTopicsSmart')
      .addSeparator()
      .addItem('⚡ Quick: Letzte 7 Tage', 'quickImportLast7Days')
      .addItem('⚡ Quick: Letzte 30 Tage', 'quickImportLast30Days')
      .addSeparator()
      .addItem('🔍 Vordefinierte Suchen (25+ Optionen)', 'showPredefinedPubMedSearches')
      .addItem('📚 Freie Suche (individuell)', 'importFromPubMedDialog')
      .addSeparator()
      .addItem('📅 Import-Historie anzeigen', 'showSmartImportHistory')
      .addItem('🔧 Import-Tracking zurücksetzen', 'resetSmartImportTracking')
      .addItem('🔧 Kaputte Volltexte reparieren', 'repairBrokenVolltexte')
      .addSeparator()
      .addItem('🔄 Letzten Import fortsetzen (alt)', 'resumePubMedImport'))
      
    
    .addSubMenu(ui.createMenu('📄 Volltext')
      .addItem('Volltextsuche Batch (50 Items)', 'batchFulltextSearchMenu')
      .addItem('Batch fortsetzen', 'resumeFulltextBatch')
      .addItem('Manuelle Volltextsuche → Dashboard', 'syncManualFulltextToDashboard')
      .addSeparator()
      .addItem('🔄 Gekürzte Abstracts vervollständigen (alle)', 'fixTruncatedAbstracts')
      .addItem('🔄 Gekürzte Abstracts vervollständigen (Auswahl)', 'fixTruncatedAbstractsForSelected')
      .addSeparator()
      .addItem('🧹 Alle Volltexte bereinigen', 'cleanAllVolltexts')
      .addItem('🧪 Volltext-Bereinigung testen', 'testVolltextCleaning')
      .addItem('📊 Volltext-Statistik anzeigen', 'showVolltextStats')
      .addSeparator()
      .addItem('➕ Volltext-Spalten hinzufügen', 'addVolltextManagementColumns')
      .addItem('ℹ️ Volltext-Spalten Info', 'showVolltextColumnInfo'))
    
    .addSeparator()
    
    .addSubMenu(ui.createMenu('🤖 Gemini Batch (Formeln)')
      .addItem('🚀 Automatisch mit Gemini API', 'analyzeWithGeminiAPI')
      .addItem('📝 Prompts vorbereiten (manuell)', 'prepareAllGeminiPrompts')
      .addSeparator()
      .addItem('📋 Schritt 0: Researcher vorbereiten', 'prepareStep0Researcher')
      .addItem('📋 Schritt 1: Triage vorbereiten', 'prepareStep1Triage')
      .addItem('📋 Schritt 2: Metadaten vorbereiten', 'prepareStep2Metadaten')
      .addItem('📋 Schritt 3: Analyse vorbereiten', 'prepareStep3Analyse')
      .addItem('📋 Schritt 4: Redaktion vorbereiten', 'prepareStep4Redaktion')
      .addItem('📋 Schritt 5: Faktencheck vorbereiten', 'prepareStep5Faktencheck')
      .addItem('📋 Schritt 6: Review vorbereiten', 'prepareStep6Review')
      .addSeparator()
      .addItem('✅ Ergebnisse übernehmen', 'processGeminiResults')
      .addSeparator()
      .addItem('🗑️ Alle Gemini-Spalten löschen', 'clearAllGeminiColumns')
      .addItem('📝 Prompts bearbeiten', 'showPromptEditor')
      .addItem('📜 Prompt-Historie', 'showPromptHistory')
      .addItem('🔧 Prompt-System initialisieren', 'initializePromptSystem')
      .addSeparator()
      .addItem('⚙️ Gemini API einrichten', 'showGeminiAPISetupInstructions')
      .addItem('❓ Hilfe', 'showGeminiHelp'))
    
    .addSubMenu(ui.createMenu('⚡ Workspace Flows (Auto)')
      .addItem('🚀 Flow für Auswahl starten', 'startWorkspaceFlowForSelection')
      .addItem('🧪 Flow-Integration testen', 'testWorkspaceFlowIntegration')
      .addItem('📊 Flow-Status anzeigen', 'showWorkspaceFlowStatus')
      .addSeparator()
      .addItem('📋 Manual Flow starten', 'openManualFlowDialog')
      .addSeparator()
      .addItem('📊 Pipeline Übersicht', 'openFlowStatusDialog')
      .addItem('⚠️ Lücken-Übersicht & Schließen', 'showGapOverview')
      .addItem('Duplikate prüfen', 'checkExistingDuplicates')
      .addItem('❓ Hilfe zu Workspace Flows', 'showWorkspaceFlowHelp'))
    
    .addSubMenu(ui.createMenu('✅ Review')
      .addItem('JSON-Antwort verarbeiten', 'processJSONResponse'))
    
    .addSeparator()
    
    .addSubMenu(ui.createMenu('📋 OnePager')
      .addItem('OnePager erstellen (Auswahl)', 'createOnePagerForSelected')
      .addItem('OnePager neu erstellen (Auswahl, Force)', 'recreateOnePagerForSelected')
      .addItem('Batch: Sehr hoch', 'batchOnePagerSehrHoch')
      .addItem('Batch: Hoch', 'batchOnePagerHoch')
      .addItem('Batch: Mittel', 'batchOnePagerMittel')
      .addItem('Batch: Niedrig', 'batchOnePagerNiedrig'))
    
    .addSubMenu(ui.createMenu('💾 Citavi Export')
      .addItem('RIS exportieren (Auswahl)', 'exportRISForSelected')
      .addItem('RIS exportieren (Auswahl, Force)', 'forceExportRISForSelected')
      .addItem('Batch: Sehr hoch', 'batchExportSehrHoch')
      .addItem('Batch: Hoch', 'batchExportHoch')
      .addItem('Batch: Mittel', 'batchExportMittel')
      .addItem('Batch: Niedrig', 'batchExportNiedrig'))
    
    .addSeparator()
    
    .addSubMenu(ui.createMenu('🔍 Qualitätscheck')
      .addItem('OnePager Readiness', 'qualityCheckOnePager')
      .addItem('Citavi Readiness', 'qualityCheckCitavi'))
    
    .addSeparator()
    
    .addSubMenu(ui.createMenu('🔧 System')
      .addItem('Initialisierung', 'initialSetup')
      .addItem('Header wiederherstellen', 'restoreHeaders')
      .addItem('Self-Check', 'selfCheck')
      .addSeparator()
      .addItem('📊 OCR Nutzungs-Statistik', 'showOcrUsageStats')
      .addItem('📊 System-Status anzeigen', 'showSystemStatus'))
    
    .addToUi();
}



/**
 * ✅ Zeigt PubMed Import Status
 */
function showPubMedImportStatus() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const lastQuery = props.getProperty('LAST_PUBMED_QUERY');
  const lastDate = props.getProperty('LAST_PUBMED_DATE');
  const lastTimestamp = props.getProperty('LAST_PUBMED_TIMESTAMP');
  
  let message = '=== PUBMED IMPORT STATUS ===\n\n';
  
  if (lastQuery) {
    message += `Letzter Import:\n`;
    message += `• Suchbegriff: "${lastQuery}"\n`;
    message += `• Datum: ${lastDate || 'neueste Papers'}\n`;
    message += `• Zeitpunkt: ${lastTimestamp ? new Date(lastTimestamp).toLocaleString('de-DE') : 'unbekannt'}\n\n`;
    message += `Nächster Import:\n`;
    message += `• Nutze "Letzten Import fortsetzen"\n`;
    message += `• Oder "Neuer Import" für anderen Suchbegriff`;
  } else {
    message += `Noch kein PubMed Import durchgeführt.\n\n`;
    message += `Starte mit: "Neuer Import mit Datum"`;
  }
  
  ui.alert('PubMed Status', message, ui.ButtonSet.OK);
}


// ==========================================
// SIDEBAR
// ==========================================

function showSidebar() {
  var html = HtmlService.createTemplateFromFile('Sidebar')
    .evaluate()
    .setTitle('Literatur Research Cockpit')
    .setWidth(360);
  SpreadsheetApp.getUi().showSidebar(html);
}

// ==========================================
// MENÜ CALLBACKS
// ==========================================

function resetCurrentRowColor() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveCell().getRow();
  if (row < 2) return;
  sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#ffffff");
}

function batchFulltextSearchMenu() {
  if(typeof batchFulltextSearch === 'function') {
    batchFulltextSearch(50);
    SpreadsheetApp.getUi().alert("Volltextsuche für 50 Items gestartet");
  } else {
    SpreadsheetApp.getUi().alert("Funktion 'batchFulltextSearch' nicht gefunden");
  }
}

function batchOnePagerSehrHoch() { batchCreateOnePagerByRelevance("Sehr hoch"); }
function batchOnePagerHoch() { batchCreateOnePagerByRelevance("Hoch"); }
function batchOnePagerMittel() { batchCreateOnePagerByRelevance("Mittel"); }
function batchOnePagerNiedrig() { batchCreateOnePagerByRelevance("Niedrig"); }

function batchExportSehrHoch() { batchExportRISByRelevance("Sehr hoch"); }
function batchExportHoch() { batchExportRISByRelevance("Hoch"); }
function batchExportMittel() { batchExportRISByRelevance("Mittel"); }
function batchExportNiedrig() { batchExportRISByRelevance("Niedrig"); }

function runBoosterForSelected() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  const selection = sheet.getActiveRange();
  
  if (!selection) {
    SpreadsheetApp.getUi().alert("Bitte Zeilen im Dashboard auswählen");
    return;
  }
  
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  if (startRow < 2) {
    SpreadsheetApp.getUi().alert("Bitte Datenzeilen auswählen (nicht Header)");
    return;
  }
  
  const volltextStatusCol = DASHBOARD_HEADERS.indexOf("Volltext-Status") + 1;
  const linkCol = DASHBOARD_HEADERS.indexOf("Link") + 1;
  const titelCol = DASHBOARD_HEADERS.indexOf("Titel") + 1;
  
  let boostedCount = 0;
  let skippedCount = 0;
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const status = sheet.getRange(row, volltextStatusCol).getValue();
    
    if (status === "VOLLTEXT_GEFUNDEN" || status === "MANUELL_HINZUGEFÜGT" || status === "VOLLTEXT_MIT_OCR") {
      skippedCount++;
      continue;
    }
    
    boostedCount++;
  }
  
  let message = `Volltext-Booster Analyse:\n\n`;
  message += `📊 ${numRows} Zeilen ausgewählt\n`;
  message += `✅ ${boostedCount} benötigen Volltext\n`;
  message += `⏭️ ${skippedCount} bereits mit Volltext\n\n`;
  
  if (boostedCount > 0) {
    message += `Öffne die Sidebar (Tab "Booster") um Volltexte hinzuzufügen.`;
  } else {
    message += `Alle ausgewählten Zeilen haben bereits Volltext!`;
  }
  
  SpreadsheetApp.getUi().alert(message);
  logAction("Booster", `${boostedCount} Zeilen für Volltext-Booster identifiziert`);
}

function debugParseLatestLabelMail() {
  try {
    let labelName = getConfig("Gmail Label Name");
    if (!labelName) {
      SpreadsheetApp.getUi().alert("Gmail Label Name nicht konfiguriert");
      return;
    }
    
    labelName = labelName.replace(/^label:/, "");
    const query = `label:"${labelName}"`;
    const threads = GmailApp.search(query, 0, 1);
    
    if (threads.length === 0) {
      SpreadsheetApp.getUi().alert(`Keine Mails mit Label "${labelName}" gefunden`);
      return;
    }
    
    const thread = threads[0];
    const messages = thread.getMessages();
    const message = messages[messages.length - 1];
    
    const subject = message.getSubject();
    const from = message.getFrom();
    const date = message.getDate();
    const body = message.getPlainBody();
    
    let result = `=== DEBUG: NEUESTE MAIL ===\n\n`;
    result += `📧 Betreff: ${subject}\n`;
    result += `👤 Von: ${from}\n`;
    result += `📅 Datum: ${date}\n`;
    result += `📝 Body-Länge: ${body.length} Zeichen\n\n`;
    
    if (subject.toLowerCase().includes("scholar") || body.includes("scholar.google.com")) {
      result += `🔍 Erkannt als: Google Scholar Alert\n\n`;
      const items = parseScholarAlertStrict(body);
      result += `📊 Gefundene Items: ${items.length}\n\n`;
      
      if (items.length > 0) {
        result += `--- ERSTES ITEM ---\n`;
        result += `Titel: ${items[0].title}\n`;
        result += `Autoren: ${items[0].authors}\n`;
        result += `Link: ${items[0].link}\n`;
      }
    } else if (subject.toLowerCase().includes("ncbi") || subject.toLowerCase().includes("pubmed")) {
      result += `🔍 Erkannt als: PubMed/NCBI Alert\n\n`;
      const pmids = extractPmidsFromText(body);
      result += `📊 Gefundene PMIDs: ${pmids.length}\n`;
      if (pmids.length > 0) {
        result += `PMIDs: ${pmids.slice(0, 5).join(", ")}${pmids.length > 5 ? "..." : ""}\n`;
      }
    } else {
      result += `⚠️ Unbekanntes Mail-Format\n`;
    }
    
    result += `\n--- BODY PREVIEW (erste 500 Zeichen) ---\n`;
    result += body.substring(0, 500);
    
    Logger.log(result);
    SpreadsheetApp.getUi().alert(result.substring(0, 2000));
    
  } catch (e) {
    logError("debugParseLatestLabelMail", e);
    SpreadsheetApp.getUi().alert("Fehler: " + e.message);
  }
}

// ==========================================
// BATCH IMPORT WRAPPER
// ==========================================

function importPhase0() { showBatchImportResult(importBatchAnswersForPhase("RESEARCHER")); }
function importPhase1() { showBatchImportResult(importBatchAnswersForPhase("TRIAGE")); }
function importPhase2() { showBatchImportResult(importBatchAnswersForPhase("METADATEN")); }
function importPhase3() { showBatchImportResult(importBatchAnswersForPhase("MASTER_ANALYSE")); }
function importPhase4() { showBatchImportResult(importBatchAnswersForPhase("REDAKTION")); }
function importPhase5() { showBatchImportResult(importBatchAnswersForPhase("FAKTENCHECK")); }
function importPhase6() { showBatchImportResult(importBatchAnswersForPhase("REVIEW")); }

function showBatchImportResult(result) {
  if (result.error) {
    SpreadsheetApp.getUi().alert("Fehler: " + result.error);
    return;
  }
  
  let msg = `✅ ${result.imported} Antworten importiert für ${result.phase}\n`;
  msg += `⏭️ ${result.skipped} übersprungen (keine Antwort)\n`;
  
  if (result.errors && result.errors.length > 0) {
    msg += `\n⚠️ Fehler bei ${result.errors.length} Zeilen:\n`;
    msg += result.errors.slice(0, 5).join("\n");
    if (result.errors.length > 5) {
      msg += `\n... und ${result.errors.length - 5} weitere`;
    }
  }
  
  SpreadsheetApp.getUi().alert(msg);
}

// ==========================================
// PROMPT GENERATION (alte Funktion für Kompatibilität)
// ==========================================

function generatePromptsForSelected() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  const promptSheet = ss.getSheetByName(MANUAL_PROMPTS_SHEET_NAME);
  
  const selection = dashSheet.getActiveRange();
  if (!selection) {
    SpreadsheetApp.getUi().alert("Bitte Zeilen auswählen");
    return;
  }
  
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  if (startRow === 1) {
    SpreadsheetApp.getUi().alert("Bitte Datenzeilen auswählen");
    return;
  }
  
  let generated = 0;
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const data = getDashboardDataByUUID(dashSheet.getRange(row, 1).getValue());
    
    if (!data) continue;
    
    try {
      const phase = determinePhaseForRow(data);
      const prompt = buildPromptForPhase(data, phase);
      const currentVersion = getCurrentPromptVersion();
      
      let prompt1 = "", prompt2 = "", prompt3 = "", driveLink = "";
      
      if (prompt.length > 50000) {
        driveLink = uploadPromptToDrive(data.UUID, data.Titel, prompt);
        prompt1 = "[PROMPT ZU LANG - IN GOOGLE DOC GESPEICHERT]";
      } else if (prompt.length > 30000) {
        const chunkSize = Math.ceil(prompt.length / 3);
        prompt1 = prompt.substring(0, chunkSize);
        prompt2 = prompt.substring(chunkSize, chunkSize * 2);
        prompt3 = prompt.substring(chunkSize * 2);
      } else {
        prompt1 = prompt;
      }
      
      promptSheet.appendRow([
        data.UUID,
        data.Titel,
        phase,
        prompt1,
        prompt2,
        prompt3,
        driveLink,
        "",
        new Date()
      ]);
      
      updateDashboardField(data.UUID, "Prompt-Version", currentVersion);
      generated++;
      
    } catch (e) {
      logError("generatePromptsForSelected", e);
    }
  }
  
  SpreadsheetApp.getUi().alert(`${generated} Prompts generiert`);
  logAction("Prompts", `${generated} Prompts generiert`);
}

function uploadPromptToDrive(uuid, title, promptText) {
  try {
    const folderId = getConfig("PDF Hauptordner ID");
    const folder = folderId ? DriveApp.getFolderById(folderId) : DriveApp.getRootFolder();
    
    const fileName = `Prompt_${title.substring(0, 30)}_${uuid}`;
    const doc = DocumentApp.create(fileName);
    doc.getBody().setText(promptText);
    doc.saveAndClose();
    
    const file = DriveApp.getFileById(doc.getId());
    file.moveTo(folder);
    
    return file.getUrl();
  } catch (e) {
    logError("uploadPromptToDrive", e);
    return "ERROR: " + e.message;
  }
}

function showOldPromptVersions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  if (sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("Keine Daten im Dashboard");
    return;
  }
  
  const currentVersion = getCurrentPromptVersion();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, DASHBOARD_HEADERS.length).getValues();
  const versionCol = DASHBOARD_HEADERS.indexOf("Prompt-Version");
  
  const oldVersions = [];
  
  data.forEach((row, index) => {
    const rowVersion = row[versionCol];
    if (rowVersion && rowVersion !== currentVersion && rowVersion !== "N/A") {
      oldVersions.push({
        row: index + 2,
        title: row[DASHBOARD_HEADERS.indexOf("Titel")],
        version: rowVersion
      });
    }
  });
  
  if (oldVersions.length === 0) {
    SpreadsheetApp.getUi().alert("Alle Prompts sind aktuell!");
    return;
  }
  
  let message = `Gefunden: ${oldVersions.length} Einträge mit alter Prompt-Version\n\n`;
  message += `Aktuelle Version: ${currentVersion}\n\n`;
  message += oldVersions.slice(0, 10).map(v =>
    `Zeile ${v.row}: ${v.title} (Version: ${v.version})`
  ).join("\n");
  
  if (oldVersions.length > 10) {
    message += `\n... und ${oldVersions.length - 10} weitere`;
  }
  
  SpreadsheetApp.getUi().alert(message);
}

function regeneratePromptsForOldVersions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  const currentVersion = getCurrentPromptVersion();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, DASHBOARD_HEADERS.length).getValues();
  const versionCol = DASHBOARD_HEADERS.indexOf("Prompt-Version");
  
  let regenerated = 0;
  
  data.forEach((row, index) => {
    const rowVersion = row[versionCol];
    if (rowVersion && rowVersion !== currentVersion && rowVersion !== "N/A") {
      const uuid = row[0];
      archiveOldAnalysis(uuid);
      updateDashboardField(uuid, "Status", "Analyse ausstehend");
      updateDashboardField(uuid, "Review-Status", "UNGEPRÜFT");
      regenerated++;
    }
  });
  
  SpreadsheetApp.getUi().alert(`${regenerated} Einträge für Neuanalyse vorbereitet.\n\nBitte Prompts neu generieren.`);
  logAction("Prompts", `${regenerated} alte Versionen markiert für Neuanalyse`);
}

function archiveOldAnalysis(uuid) {
  try {
    const data = getDashboardDataByUUID(uuid);
    if (!data) return;
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let archiveSheet = ss.getSheetByName("Analyse-Archiv");
    
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet("Analyse-Archiv");
      archiveSheet.appendRow([
        "Archiviert am", "UUID", "Titel", "Alte Version",
        "Haupterkenntnis", "Zusammenfassung", "Relevanz"
      ]);
    }
    
    archiveSheet.appendRow([
      new Date(),
      uuid,
      data.Titel,
      data["Prompt-Version"],
      data.Haupterkenntnis,
      data.Zusammenfassung,
      data.Relevanz
    ]);
    
  } catch (e) {
    logError("archiveOldAnalysis", e);
  }
}

// ==========================================
// UTILITIES
// ==========================================

function getPendingFulltexts() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  if(!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  
  const idxRel = headers.indexOf("Relevanz");
  const idxStatus = headers.indexOf("Volltext-Status");
  const idxTitle = headers.indexOf("Titel");
  const idxLink = headers.indexOf("Link");
  
  if (idxRel === -1) return [];
  const items = [];
  
  for (let i = 1; i < data.length; i++) {
    try {
      if (sheet.isRowHiddenByFilter(i + 1)) continue;
    } catch (e) {}
    
    const row = data[i];
    const rel = String(row[idxRel]).toLowerCase();
    const stat = String(row[idxStatus]);
    
    if ((rel.includes("hoch") || rel.includes("mittel")) && 
        stat !== "VOLLTEXT_GEFUNDEN" && 
        stat !== "MANUELL_HINZUGEFÜGT" &&
        stat !== "VOLLTEXT_MIT_OCR") {
      items.push({
        rowIndex: i+1,
        title: row[idxTitle],
        link: row[idxLink],
        relevanz: row[idxRel]
      });
    }
  }
  return items;
}

function applyDashboardFilters(criteria) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const idxStatus = headers.indexOf("Status");
  const idxRel = headers.indexOf("Relevanz");
  
  sheet.showRows(1, sheet.getMaxRows());
  
  for (let i = 1; i < data.length; i++) {
    let hide = false;
    const row = data[i];
    
    if (criteria.status && row[idxStatus] !== criteria.status) hide = true;
    if (!hide && criteria.relevance && row[idxRel] !== criteria.relevance) hide = true;
    
    if (hide) sheet.hideRows(i + 1);
  }
}

function clearDashboardFilters() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  sheet.showRows(1, sheet.getMaxRows());
}

function saveManualFulltext(rowIndex, text) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const colIdx = headers.indexOf("Volltext/Extrakt") + 1;
  const colStat = headers.indexOf("Volltext-Status") + 1;
  
  if (colIdx > 0) {
    sheet.getRange(rowIndex, colIdx).setValue(text);
    if(colStat > 0) sheet.getRange(rowIndex, colStat).setValue("MANUELL_HINZUGEFÜGT");
    
    return "✅ Gespeichert! Jetzt Sidebar Tab 1 → Analyse starten";
  }
  return "Fehler: Spalte 'Volltext/Extrakt' fehlt.";
}

function processBoosterImages(rowIndex, base64Images) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
    const uuid = sheet.getRange(rowIndex, 1).getValue();
    const title = sheet.getRange(rowIndex, getDashboardColumnIndex("Titel")).getValue();
    
    if (!base64Images || base64Images.length === 0) {
      return { error: "Keine Bilder übergeben" };
    }
    
    const geminiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    if (!geminiKey) {
      return { error: "Gemini API Key nicht konfiguriert. Bitte in Script Properties hinterlegen." };
    }
    
    let combinedText = "";
    let processedCount = 0;
    
    for (let i = 0; i < base64Images.length; i++) {
      try {
        const imageBase64 = base64Images[i];
        
        const payload = {
          contents: [{
            parts: [
              {
                inline_data: {
                  mime_type: "image/png",
                  data: imageBase64
                }
              },
              {
                text: "Extrahiere den gesamten Text aus diesem Screenshot/Bild einer wissenschaftlichen Publikation. " +
                      "Behalte die Struktur bei (Überschriften, Absätze, Tabellen). " +
                      "Bei Tabellen: Formatiere sie lesbar mit | als Trennzeichen. " +
                      "Antworte NUR mit dem extrahierten Text, keine Erklärungen."
              }
            ]
          }],
          generationConfig: {
            temperature: 0.1,
            maxOutputTokens: 4000
          }
        };
        
        const options = {
          method: "post",
          contentType: "application/json",
          headers: {
            "x-goog-api-key": geminiKey
          },
          payload: JSON.stringify(payload),
          muteHttpExceptions: true
        };
        
        const response = UrlFetchApp.fetch(
          "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent",
          options
        );
        
        if (response.getResponseCode() === 200) {
          const result = JSON.parse(response.getContentText());
          if (result.candidates && result.candidates.length > 0) {
            const extractedText = result.candidates[0].content.parts[0].text;
            combinedText += `\n\n--- Screenshot ${i + 1} ---\n${extractedText}`;
            processedCount++;
          }
        }
        
        Utilities.sleep(1000);
        
      } catch (imgError) {
        logError("processBoosterImages - Image " + (i + 1), imgError);
      }
    }
    
    if (processedCount === 0) {
      return { error: "Konnte keinen Text aus den Bildern extrahieren" };
    }
    
    const fullText = `[📸 Aus ${processedCount} Screenshot(s) extrahiert]\n${combinedText}`;
    
    updateDashboardField(uuid, "Volltext/Extrakt", fullText);
    updateDashboardField(uuid, "Volltext-Status", "MANUELL_HINZUGEFÜGT");
    
    return {
      message: `✅ ${processedCount}/${base64Images.length} Screenshots verarbeitet!\n\n` +
               `📝 Extrahierter Text: ${fullText.length} Zeichen\n\n` +
               `Jetzt: Sidebar Tab 1 → Analyse starten`
    };
    
  } catch (e) {
    logError("processBoosterImages", e);
    return { error: e.message };
  }
}

/**
 * 🌟 Import: ALLE THEMEN auf einmal (Diabetes Typ 1+2 + Immu/Onko)
 */
function importAllTopics() {
  importPredefinedSearch(
    '((continuous glucose monitoring OR CGM OR insulin pump OR CSII OR ' +
    'continuous subcutaneous insulin infusion OR diabetes technology OR diabetes device OR ' +
    'diabetes devices OR automated insulin delivery OR AID OR closed loop OR hybrid closed loop OR ' +
    'flash glucose monitoring OR FGM OR real-time CGM OR rtCGM OR sensor-augmented pump OR SAP) ' +
    'AND (diabetes OR diabetic OR type 1 diabetes OR type 2 diabetes OR T1D OR T2D OR T1DM OR T2DM OR ' +
    'diabetes mellitus OR insulin dependent diabetes OR IDDM OR non-insulin dependent diabetes OR NIDDM OR ' +
    'insulin-requiring diabetes OR insulin therapy)) ' +
    'OR ' +
    '(primary immunodeficiency OR primary immunodeficiencies OR PID OR PIDD OR ' +
    'secondary immunodeficiency OR secondary immunodeficiencies OR acquired immunodeficiency OR ' +
    'immunoglobulin OR immunoglobulins OR IVIG OR SCIG OR subcutaneous immunoglobulin OR ' +
    'IgG replacement OR immunoglobulin therapy OR immunoglobulin treatment)',
    'Alle Themen (Diabetes Typ 1+2 + Immu/Onko)'
  );
}

/**
 * ⚡ Import: Diabetes (Typ 1 + Typ 2)
 */
function importDiabetes() {
  importPredefinedSearch(
    '(continuous glucose monitoring OR CGM OR insulin pump OR insulin pumps OR CSII OR ' +
    'continuous subcutaneous insulin infusion OR diabetes technology OR diabetes device OR ' +
    'diabetes devices OR automated insulin delivery OR AID OR closed loop OR hybrid closed loop OR ' +
    'flash glucose monitoring OR FGM OR real-time CGM OR rtCGM OR sensor-augmented pump OR SAP) ' +
    'AND ' +
    '(diabetes OR diabetic OR type 1 diabetes OR type 2 diabetes OR T1D OR T2D OR T1DM OR T2DM OR ' +
    'diabetes mellitus OR insulin dependent diabetes OR IDDM OR non-insulin dependent diabetes OR ' +
    'NIDDM OR insulin-requiring diabetes OR insulin therapy)',
    'Diabetes Typ 1+2 (CGM, Pumps, Technology)'
  );
}

/**
 * ⚡ Import: PAH
 */
function importPAH() {
  importPredefinedSearch(
    '(pulmonary hypertension OR pulmonary arterial hypertension OR PAH) AND ' +
    '(treatment OR therapy OR therapeutic OR biomarker OR biomarkers OR diagnostic OR ' +
    'diagnosis OR prognosis OR management OR medication OR drug)',
    'PAH (Hypertension, Treatment, Biomarkers)'
  );
}

/**
 * ⚡ Import: Immu/Onko
 */
function importImmuOnko() {
  importPredefinedSearch(
    '(primary immunodeficiency OR primary immunodeficiencies OR PID OR PIDD OR ' +
    'secondary immunodeficiency OR secondary immunodeficiencies OR acquired immunodeficiency OR ' +
    'immunoglobulin OR immunoglobulins OR IVIG OR SCIG OR subcutaneous immunoglobulin OR ' +
    'IgG replacement OR immunoglobulin therapy OR immunoglobulin treatment)',
    'Immu/Onko (Immunodeficiencies, IgG)'
  );
}

/**
 * ✅ Führt vordefinierten Import durch
 */
function importPredefinedSearch(searchQuery, displayName) {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const lastQuery = props.getProperty('LAST_PUBMED_QUERY');
  const lastDate = props.getProperty('LAST_PUBMED_DATE');
  
  let message = `📚 Import: ${displayName}\n\n`;
  
  if (lastQuery === searchQuery && lastDate) {
    message += `✅ Bereits importiert!\n`;
    message += `Letztes Paper: ${lastDate}\n\n`;
    message += `─────────────────────────\n\n`;
    message += `Was möchtest du tun?\n\n`;
    message += `JA = Import fortsetzen (ab ${lastDate})\n`;
    message += `NEIN = Neuer Import (anderes Startdatum)\n`;
    message += `ABBRECHEN = Abbrechen`;
    
    const response = ui.alert('Import fortsetzen?', message, ui.ButtonSet.YES_NO_CANCEL);
    
    if (response === ui.Button.CANCEL) {
      return;
    } else if (response === ui.Button.YES) {
      askForCountAndImport(searchQuery, lastDate, displayName);
      return;
    }
  } else {
    message += `Wie möchtest du importieren?\n\n`;
    message += `JA = Neueste Papers (ohne Datumsfilter)\n`;
    message += `NEIN = Ab Startdatum importieren\n`;
    message += `ABBRECHEN = Abbrechen`;
    
    const response = ui.alert('Import-Modus', message, ui.ButtonSet.YES_NO_CANCEL);
    
    if (response === ui.Button.CANCEL) {
      return;
    } else if (response === ui.Button.YES) {
      askForCountAndImport(searchQuery, null, displayName);
      return;
    }
  }
  
  const dateResponse = ui.prompt(
    `${displayName} - Startdatum`,
    'Ab welchem Datum importieren?\n\n' +
    'Format: YYYY/MM/DD\n' +
    'Beispiel: 2024/01/01\n\n' +
    'Leer lassen = neueste Papers\n' +
    '"alle" = alle verfügbaren Papers',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (dateResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  let startDate = dateResponse.getResponseText().trim();
  
  if (startDate.toLowerCase() === 'alle') {
    startDate = null;
  } else if (!startDate) {
    startDate = null;
  }
  
  askForCountAndImport(searchQuery, startDate, displayName);
}

/**
 * ✅ Fragt nach Anzahl und startet Import
 */
function askForCountAndImport(searchQuery, startDate, displayName) {
  const ui = SpreadsheetApp.getUi();
  
  const countResponse = ui.prompt(
    `${displayName} - Anzahl`,
    `Wie viele Papers importieren? (1-500)\n\n` +
    `Empfehlung:\n` +
    `• Erster Import: 100-200\n` +
    `• Routine-Update: 50-100`,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (countResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const maxResults = parseInt(countResponse.getResponseText()) || 50;
  
  if (maxResults < 1 || maxResults > 500) {
    ui.alert('❌ Ungültige Anzahl!\n\nBitte Zahl zwischen 1 und 500 eingeben.');
    return;
  }
  
  let confirmMsg = `=== IMPORT STARTEN ===\n\n`;
  confirmMsg += `Thema: ${displayName}\n`;
  confirmMsg += `Anzahl: ${maxResults} Papers\n`;
  confirmMsg += `Startdatum: ${startDate || 'neueste Papers'}\n\n`;
  confirmMsg += `Dies kann einige Minuten dauern.\n\n`;
  confirmMsg += `Jetzt starten?`;
  
  const confirm = ui.alert(confirmMsg, ui.ButtonSet.YES_NO);
  
  if (confirm !== ui.Button.YES) {
    return;
  }
  
  importFromPubMedWithDate(searchQuery, startDate, maxResults);
}

/**
 * ✅ Setzt PubMed Import fort
 */
function resumePubMedImport() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const lastQuery = props.getProperty('LAST_PUBMED_QUERY');
  const lastDate = props.getProperty('LAST_PUBMED_DATE');
  
  if (!lastQuery) {
    ui.alert(
      'Kein vorheriger Import gefunden',
      'Es wurde noch kein PubMed Import durchgeführt.\n\n' +
      'Starte mit: "Von PubMed importieren (Neu)"',
      ui.ButtonSet.OK
    );
    return;
  }
  
  const response = ui.alert(
    'Import fortsetzen',
    `Letzter Import:\n\n` +
    `Suchbegriff: "${lastQuery}"\n` +
    `Letztes Datum: ${lastDate || 'neueste Papers'}\n\n` +
    `Fortsetzen ab diesem Punkt?`,
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  const countResponse = ui.prompt(
    'Anzahl Papers',
    'Wie viele weitere Papers importieren? (1-500)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (countResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  const maxResults = parseInt(countResponse.getResponseText()) || 50;
  
  if (maxResults < 1 || maxResults > 500) {
    ui.alert('Bitte Zahl zwischen 1 und 500 eingeben!');
    return;
  }
  
  ui.alert(
    'Import gestartet',
    `Importiere ${maxResults} weitere Papers...\n\nDies kann einige Minuten dauern.`,
    ui.ButtonSet.OK
  );
  
  importFromPubMedWithDate(lastQuery, lastDate, maxResults);
}

/**
 * ✅ Zeigt Import-Historie
 */
function showImportHistory() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  
  const lastQuery = props.getProperty('LAST_PUBMED_QUERY');
  const lastDate = props.getProperty('LAST_PUBMED_DATE');
  const lastTimestamp = props.getProperty('LAST_PUBMED_TIMESTAMP');
  
  let message = '=== PUBMED IMPORT HISTORIE ===\n\n';
  
  if (lastQuery) {
    message += `📚 Letzter Import:\n\n`;
    message += `Suchbegriff: "${lastQuery.substring(0, 100)}..."\n`;
    message += `Letztes Paper-Datum: ${lastDate || 'keine Angabe'}\n`;
    message += `Import-Zeitpunkt: ${lastTimestamp ? new Date(lastTimestamp).toLocaleString('de-DE') : 'unbekannt'}\n\n`;
    message += `─────────────────────────────\n\n`;
    message += `💡 Nächster Import:\n\n`;
    message += `Option 1: "Import fortsetzen"\n`;
    message += `  → Holt weitere Papers ab ${lastDate || 'neuesten Papers'}\n\n`;
    message += `Option 2: "Von PubMed importieren (Neu)"\n`;
    message += `  → Neuer Suchbegriff oder anderes Datum`;
  } else {
    message += `Noch kein PubMed Import durchgeführt.\n\n`;
    message += `Starte mit:\n`;
    message += `📥 Import → 📚 Freie Suche oder wähle ein Thema`;
  }
  
  ui.alert('Import-Historie', message, ui.ButtonSet.OK);
}

/**
 * ✅ Setzt Import-Tracking zurück
 */
function resetImportTracking() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Import-Tracking zurücksetzen',
    'Sollen alle gespeicherten Import-Daten gelöscht werden?\n\n' +
    'Dies betrifft:\n' +
    '• Letzten Suchbegriff\n' +
    '• Letztes Datum\n' +
    '• Import-Zeitpunkt\n\n' +
    'Papers im Dashboard bleiben erhalten!',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  try {
    const props = PropertiesService.getScriptProperties();
    props.deleteProperty('LAST_PUBMED_QUERY');
    props.deleteProperty('LAST_PUBMED_DATE');
    props.deleteProperty('LAST_PUBMED_TIMESTAMP');
    
    ui.alert('✅ Import-Tracking zurückgesetzt!\n\nDer nächste Import startet neu.');
    
    if (typeof logAction === 'function') {
      logAction("Import-Tracking", "Zurückgesetzt");
    }
    
  } catch (e) {
    ui.alert('❌ Fehler beim Zurücksetzen:\n\n' + e.message);
  }
}

/**
 * ✅ Zeigt System-Status
 */
function showSystemStatus() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName("Dashboard");
  
  let message = '=== VITALAIRE SYSTEM STATUS ===\n\n';
  
  if (dashboard && dashboard.getLastRow() > 1) {
    const totalPapers = dashboard.getLastRow() - 1;
    message += `📄 Papers im Dashboard: ${totalPapers}\n`;
    
    const headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
    const data = dashboard.getRange(2, 1, dashboard.getLastRow() - 1, headers.length).getValues();
    const statusCol = headers.indexOf("Status");
    
    if (statusCol >= 0) {
      const statusCount = {};
      data.forEach(row => {
        const status = String(row[statusCol] || "").trim() || "LEER";
        statusCount[status] = (statusCount[status] || 0) + 1;
      });
      
      message += '\nStatus-Verteilung:\n';
      Object.keys(statusCount).sort().forEach(status => {
        message += `  • ${status}: ${statusCount[status]}\n`;
      });
    }
    
  } else {
    message += '📄 Papers im Dashboard: 0\n';
  }
  
  message += '\n──────────────────────────\n';
  message += '🔌 APIs & Services:\n\n';
  
  const geminiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  message += geminiKey ? '✅ Gemini API: Eingerichtet\n' : '⚠️ Gemini API: Nicht eingerichtet\n';
  
  message += '✅ PubMed API: Verfügbar (kostenlos)\n';
  message += '✅ Europe PMC: Verfügbar (kostenlos)\n';
  message += '✅ Crossref: Verfügbar (kostenlos)\n';
  message += '✅ Semantic Scholar: Verfügbar (kostenlos)\n';
  
  message += '\n──────────────────────────\n';
  message += '🚀 Features:\n\n';
  message += '✅ PubMed Import (mit Datums-Filter)\n';
  message += '✅ Abstract-Vervollständigung (12 Publisher)\n';
  message += '✅ PDF-Verarbeitung & OCR\n';
  message += geminiKey ? '✅ Gemini Auto-Analyse\n' : '⚠️ Gemini Auto-Analyse (API Key fehlt)\n';
  message += '✅ Gemini Manuell (via Sheets Extension)\n';
  message += '✅ OnePager Generation\n';
  message += '✅ Citavi Export\n';
  
  const lastQuery = PropertiesService.getScriptProperties().getProperty('LAST_PUBMED_QUERY');
  if (lastQuery) {
    message += '\n──────────────────────────\n';
    message += '📥 Letzter Import:\n\n';
    message += `"${lastQuery.substring(0, 50)}..."\n`;
    const lastDate = PropertiesService.getScriptProperties().getProperty('LAST_PUBMED_DATE');
    if (lastDate) {
      message += `ab ${lastDate}\n`;
    }
  }
  
  ui.alert('System-Status', message, ui.ButtonSet.OK);
}

// ==========================================
// SMART IMPORT ALIASES (für Menü-Kompatibilität)
// ==========================================

/**
 * ✅ Zeigt Smart Import Historie
 * Alias für showImportHistory() – wird vom Menü aufgerufen
 */
function showSmartImportHistory() {
  showImportHistory(); // Definiert weiter oben in ui.gs
}

/**
 * ✅ Setzt Smart Import Tracking zurück
 * Alias für resetImportTracking() – wird vom Menü aufgerufen
 */
function resetSmartImportTracking() {
  resetImportTracking(); // Definiert weiter oben in ui.gs
}
