// FILE: batchPipeline.gs

/**
 * ==========================================
 * BATCH-PIPELINE FÜR PHASEN-BASIERTE ANALYSE
 * ==========================================
 * 
 * Workflow:
 * 1. Zeilen im Dashboard auswählen
 * 2. Menü: "Phase X Prompts generieren" 
 *    → Prompts werden in "Manuelle Prompts" Sheet geschrieben
 * 3. User kopiert Prompts → ChatGPT → Antworten zurück
 * 4. Menü: "Phase X Antworten importieren"
 *    → JSON wird validiert & ins Dashboard geschrieben
 * 5. Nächste Phase nur für erfolgreiche Zeilen
 */

// ==========================================
// PHASE-DEFINITIONEN
// ==========================================

const PHASES = {
  RESEARCHER: {
    id: "RESEARCHER",
    name: "Phase 0: Volltext-Recherche",
    promptCell: "B15",
    requiredFields: ["Titel"], // Nur Titel nötig (DOI optional)
    targetFields: ["Volltext/Extrakt", "PICO Population"], 
    nextStatus: "Researcher fertig",
    description: "⭐ MUSS ZUERST LAUFEN! Sucht Volltext via DOI, extrahiert PICO."
  },
  TRIAGE: {
    id: "TRIAGE",
    name: "Phase 1: Kategorisierung & Triage",
    promptCell: "B19",
    requiredFields: ["Titel", "Inhalt/Abstract"], // Braucht zumindest Abstract
    optionalFields: ["Volltext/Extrakt"], // Besser mit Volltext
    targetFields: ["Hauptkategorie", "Unterkategorien", "Schlagwörter"],
    nextStatus: "Triage fertig",
    description: "Ordnet Paper in Kategorie ein. Braucht Abstract ODER Volltext."
  },
  METADATEN: {
    id: "METADATEN",
    name: "Phase 2: Metadaten-Extraktion",
    promptCell: "B17",
    requiredFields: ["Titel"], 
    optionalFields: ["Volltext/Extrakt", "Inhalt/Abstract"],
    targetFields: ["Jahr", "Artikeltyp/Studientyp", "Autoren"],
    nextStatus: "Metadaten fertig",
    description: "Extrahiert Jahr, Studientyp, Journal. Funktioniert auch mit Abstract."
  },
  MASTER_ANALYSE: {
    id: "MASTER_ANALYSE",
    name: "Phase 3: Strategische Analyse",
    promptCell: "B16",
    requiredFields: ["Titel", "Hauptkategorie"], // Kategorie muss da sein!
    criticalFields: ["Volltext/Extrakt"], // VOLLTEXT ZWINGEND für gute Analyse!
    targetFields: ["Relevanz", "Haupterkenntnis", "Zusammenfassung"],
    nextStatus: "Analyse fertig",
    description: "⭐⭐⭐ HAUPTPHASE! BRAUCHT VOLLTEXT! Wenn nur Abstract → schlechte Qualität."
  },
  REDAKTION: {
    id: "REDAKTION",
    name: "Phase 4: Redaktions-Check",
    promptCell: "B18",
    requiredFields: ["Haupterkenntnis", "Volltext/Extrakt"], // Volltext für Figures/Tables
    targetFields: ["Praktische Implikationen"],
    nextStatus: "Redaktion fertig",
    description: "Identifiziert wichtige Abbildungen. Braucht Volltext."
  },
  FAKTENCHECK: {
    id: "FAKTENCHECK",
    name: "Phase 5: Faktencheck",
    promptCell: "B20",
    requiredFields: ["Haupterkenntnis", "Volltext/Extrakt"], // Volltext zum Validieren
    targetFields: ["Kritische Bewertung"],
    nextStatus: "Faktencheck fertig",
    description: "Validiert Kernaussagen gegen Originaltext. Braucht Volltext."
  },
  REVIEW: {
    id: "REVIEW",
    name: "Phase 6: Senior Review",
    promptCell: "B21",
    requiredFields: ["Relevanz", "Haupterkenntnis"],
    targetFields: ["Review-Status"],
    nextStatus: "Abgeschlossen",
    description: "Senior prüft Phase 3 Ergebnis. Kann auch mit Abstract arbeiten."
  }
};

// ==========================================
// 1. PROMPT-GENERIERUNG FÜR PHASE
// ==========================================

function showBatchPromptDialog() {
  const html = HtmlService.createHtmlOutputFromFile('BatchPromptDialog')
    .setWidth(450)
    .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Batch Prompt-Generierung');
}

/**
 * Wird vom Dialog aufgerufen
 */
function generateBatchPromptsForPhase(phaseId, useSelection, maxCount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  const promptSheet = ss.getSheetByName(MANUAL_PROMPTS_SHEET_NAME);
  
  const phase = PHASES[phaseId];
  if (!phase) throw new Error("Ungültige Phase: " + phaseId);
  
  // 1. Zeilen sammeln
  let rowsToProcess = [];
  
  if (useSelection) {
    // Nur markierte Zeilen
    const selection = dashSheet.getActiveRange();
    if (!selection) {
      return { error: "Bitte Zeilen im Dashboard auswählen" };
    }
    
    const startRow = selection.getRow();
    const numRows = selection.getNumRows();
    
    for (let i = 0; i < numRows; i++) {
      if (startRow + i < 2) continue; // Skip Header
      rowsToProcess.push(startRow + i);
    }
  } else {
    // Alle Zeilen, die für diese Phase bereit sind
    const data = dashSheet.getRange(2, 1, dashSheet.getLastRow() - 1, DASHBOARD_HEADERS.length).getValues();
    
    data.forEach((row, index) => {
      const rowData = {};
      DASHBOARD_HEADERS.forEach((header, i) => {
        rowData[header] = row[i];
      });
      
      // Prüfe ob Zeile bereit ist
      if (isRowReadyForPhase(rowData, phase)) {
        rowsToProcess.push(index + 2);
      }
    });
  }
  
  // Limit anwenden
  if (maxCount && maxCount > 0) {
    rowsToProcess = rowsToProcess.slice(0, maxCount);
  }
  
  if (rowsToProcess.length === 0) {
    return { error: "Keine Zeilen für Phase " + phase.name + " gefunden" };
  }
  
  // 2. Prompts generieren
  let generated = 0;
  const errors = [];
  
  rowsToProcess.forEach(rowIndex => {
    try {
      const rowData = getRowDataByIndex(dashSheet, rowIndex);
      const prompt = buildPromptForPhaseWithData(phase, rowData);
      const currentVersion = getCurrentPromptVersion();
      
      // Prompt-Länge prüfen
      let prompt1 = "", prompt2 = "", prompt3 = "", driveLink = "";
      
      if (prompt.length > 50000) {
        // Drive Upload
        driveLink = uploadPromptToDrive(rowData.UUID, rowData.Titel, prompt);
        prompt1 = "[PROMPT ZU LANG - IN GOOGLE DOC GESPEICHERT]";
      } else if (prompt.length > 30000) {
        // 3-Teil Splitting
        const chunkSize = Math.ceil(prompt.length / 3);
        prompt1 = prompt.substring(0, chunkSize);
        prompt2 = prompt.substring(chunkSize, chunkSize * 2);
        prompt3 = prompt.substring(chunkSize * 2);
      } else {
        prompt1 = prompt;
      }
      
      // In Manual Prompts Sheet schreiben
      promptSheet.appendRow([
        rowData.UUID,
        rowData.Titel,
        phase.id,
        prompt1,
        prompt2,
        prompt3,
        driveLink,
        "", // Antwort leer
        new Date()
      ]);
      
      // Version im Dashboard aktualisieren
      updateDashboardField(rowData.UUID, "Prompt-Version", currentVersion);
      
      // ✅ NEU: Batch-Phase markieren
      updateDashboardField(rowData.UUID, "Batch-Phase", `${phase.name} - Prompt generiert`);
      
      generated++;
      
    } catch (e) {
      errors.push(`Zeile ${rowIndex}: ${e.message}`);
      logError("generateBatchPromptsForPhase", e);
    }
  });
  
  logAction("Batch Prompts", `${generated} Prompts für ${phase.name} generiert`);
  
  return {
    success: true,
    generated: generated,
    phase: phase.name,
    errors: errors
  };
}

/**
 * Prüft ob eine Zeile bereit für die Phase ist
 * ENHANCED: Berücksichtigt criticalFields (Warnungen bei fehlenden Volltext)
 */
function isRowReadyForPhase(rowData, phase) {
  // 1. Required Fields müssen vorhanden sein
  for (const field of phase.requiredFields) {
    const value = rowData[field];
    if (!value || value === "" || value === "N/A") {
      return false;
    }
  }
  
  // 2. Critical Fields WARNUNG (aber nicht blockierend)
  if (phase.criticalFields) {
    for (const field of phase.criticalFields) {
      const value = rowData[field];
      if (!value || value === "" || value === "N/A") {
        // Logge Warnung, aber lasse Zeile durch
        Logger.log(`⚠️ WARNUNG: ${rowData.Titel} - ${field} fehlt! Phase ${phase.name} wird schlechte Qualität haben.`);
      }
    }
  }
  
  // 3. Target Fields dürfen NICHT schon befüllt sein (sonst Phase schon erledigt)
  const firstTargetField = phase.targetFields[0];
  if (rowData[firstTargetField] && rowData[firstTargetField] !== "" && rowData[firstTargetField] !== "N/A") {
    return false;
  }
  
  return true;
}

/**
 * Baut Prompt für eine Phase mit konkreten Daten
 */
function buildPromptForPhaseWithData(phase, rowData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  let template = configSheet.getRange(phase.promptCell).getValue();
  
  if (!template) {
    throw new Error(`Prompt für ${phase.name} nicht gefunden in ${phase.promptCell}`);
  }
  
  // Kategorien laden
  const lists = loadCategoriesAndKeywords();
  
  // Platzhalter ersetzen
  template = template
    .replace(/{PUBLIKATIONSTEXT}/g, rowData["Volltext/Extrakt"] || rowData["Inhalt/Abstract"] || "N/A")
    .replace(/{TITEL}/g, rowData.Titel || "N/A")
    .replace(/{DOI}/g, rowData.DOI || "N/A")
    .replace(/{AUTOREN}/g, rowData.Autoren || "N/A")
    .replace(/{JAHR}/g, rowData.Jahr || "N/A")
    .replace(/{JOURNAL}/g, rowData["Journal/Quelle"] || "N/A")
    .replace(/{ABSTRACT}/g, rowData["Inhalt/Abstract"] || "N/A")
    .replace(/{VOLLTEXT}/g, rowData["Volltext/Extrakt"] || rowData["Inhalt/Abstract"] || "N/A")
    .replace(/{PDF_LINK}/g, rowData["Volltext-Datei-Link"] || "Kein PDF verfügbar")
    .replace(/{KATEGORIEN_LISTE}/g, lists.categories.join(", "))
    .replace(/{KATEGORIE_MAPPING}/g, lists.fullMapping)
    .replace(/{SCHLAGWÖRTER_LISTE}/g, lists.keywords.join(", "))
    .replace(/{KERNAUSSAGE}/g, rowData.Haupterkenntnis || "N/A")
    .replace(/{JSON_ANALYSE_VON_SCHRITT_1}/g, buildAnalysisJsonForReview(rowData))
    .replace(/{UUID}/g, rowData.UUID);
  
  return template;
}

/**
 * Baut das Analyse-JSON für den Review-Prompt
 */
function buildAnalysisJsonForReview(rowData) {
  return JSON.stringify({
    relevanz: rowData.Relevanz,
    produkt_fokus: rowData["Produkt-Fokus"],
    kategorie: {
      hauptkategorie: rowData.Hauptkategorie,
      unterkategorien: (rowData.Unterkategorien || "").split(", ")
    },
    haupterkenntnis: rowData.Haupterkenntnis,
    kernaussagen: (rowData.Kernaussagen || "").split("\n").filter(k => k.trim()),
    zusammenfassung: rowData.Zusammenfassung,
    schlagwoerter: (rowData.Schlagwörter || "").split(", "),
    kritische_bewertung: rowData["Kritische Bewertung"]
  }, null, 2);
}

// ==========================================
// 2. ANTWORTEN-IMPORT FÜR PHASE
// ==========================================

function importBatchAnswersForPhase(phaseId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const promptSheet = ss.getSheetByName(MANUAL_PROMPTS_SHEET_NAME);
  
  const phase = PHASES[phaseId];
  if (!phase) throw new Error("Ungültige Phase: " + phaseId);
  
  if (promptSheet.getLastRow() <= 1) {
    return { error: "Keine Prompts im Manual Prompts Sheet" };
  }
  
  const data = promptSheet.getRange(2, 1, promptSheet.getLastRow() - 1, 9).getValues();
  
  let imported = 0;
  let skipped = 0;
  const errors = [];
  
  data.forEach((row, index) => {
    const [uuid, titel, promptPhase, p1, p2, p3, driveLink, answer, timestamp] = row;
    
    // Nur Zeilen für diese Phase
    if (promptPhase !== phase.id) return;
    
    // Nur Zeilen mit Antwort
    if (!answer || answer === "") {
      skipped++;
      return;
    }
    
    try {
      // JSON validieren & parsen
      const jsonData = parseAndValidateJSON(answer, phase.id);
      
      // Ins Dashboard schreiben
      applyJSONToDashboard(uuid, phase.id, jsonData);
      
      // Status updaten
      updateDashboardField(uuid, "Status", phase.nextStatus);
      updateDashboardField(uuid, "Review-Status", "GEPRÜFT");
      
      // ✅ NEU: Batch-Phase aktualisieren
      updateDashboardField(uuid, "Batch-Phase", `${phase.name} - Abgeschlossen ✓`);
      
      // Zeile im Prompt Sheet als verarbeitet markieren
      promptSheet.getRange(index + 2, 8).setValue("[IMPORTIERT " + new Date().toLocaleString() + "]");
      
      imported++;
      
    } catch (e) {
      errors.push(`${titel}: ${e.message}`);
      logError("importBatchAnswersForPhase", e);
    }
  });
  
  logAction("Batch Import", `${imported} Antworten für ${phase.name} importiert`);
  
  return {
    success: true,
    imported: imported,
    skipped: skipped,
    phase: phase.name,
    errors: errors
  };
}

// ==========================================
// HELPER
// ==========================================

function getRowDataByIndex(sheet, rowIndex) {
  const rowValues = sheet.getRange(rowIndex, 1, 1, DASHBOARD_HEADERS.length).getValues()[0];
  const data = {};
  DASHBOARD_HEADERS.forEach((header, index) => {
    data[header] = rowValues[index] || "";
  });
  return data;
}

/**
 * Zeigt Statistik welche Zeilen für welche Phase bereit sind
 */
function showBatchReadiness() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  if (sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("Keine Daten im Dashboard");
    return;
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, DASHBOARD_HEADERS.length).getValues();
  
  const stats = {};
  Object.keys(PHASES).forEach(key => {
    stats[key] = 0;
  });
  
  data.forEach(row => {
    const rowData = {};
    DASHBOARD_HEADERS.forEach((header, i) => {
      rowData[header] = row[i];
    });
    
    // Prüfe für welche Phase die Zeile bereit ist
    Object.keys(PHASES).forEach(key => {
      if (isRowReadyForPhase(rowData, PHASES[key])) {
        stats[key]++;
      }
    });
  });
  
  let message = "=== BATCH READINESS ===\n\n";
  Object.keys(PHASES).forEach(key => {
    const phase = PHASES[key];
    message += `${phase.name}: ${stats[key]} Zeilen bereit\n`;
  });
  
  SpreadsheetApp.getUi().alert(message);
}
