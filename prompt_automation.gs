// FILE: prompt_automation.gs
/**
 * ==========================================
 * PROMPT AUTOMATION
 * ==========================================
 * Auto-Regenerierung von Prompts bei neuem Volltext
 */

/**
 * ✅ Wird aufgerufen wenn Volltext hinzugefügt wird
 * Triggert automatische Prompt-Regenerierung
 */
function onVolltextAdded(uuid) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Finde Zeile
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === uuid) {
        rowIndex = i + 1;
        break;
      }
    }
    
    if (rowIndex === -1) return;
    
    // Hole aktuelle Daten
    const haupterkenntnisCol = headers.indexOf("Haupterkenntnis") + 1;
    const statusCol = headers.indexOf("Status") + 1;
    const reviewStatusCol = headers.indexOf("Review-Status") + 1;
    const haupterkenntnis = sheet.getRange(rowIndex, haupterkenntnisCol).getValue();
    
    if (haupterkenntnis && String(haupterkenntnis).length > 10) {
      // ===== ALTE ANALYSE GEFUNDEN =====
      
      Logger.log("🔄 Volltext hinzugefügt für bereits analysiertes Paper: " + uuid);
      
      // 1. Archiviere alte Analyse
      archiveOldAnalysisOnVolltextUpdate(uuid);
      
      // 2. Setze Status zurück
      if (statusCol > 0) {
        sheet.getRange(rowIndex, statusCol).setValue("🔄 Prompts werden neu generiert...");
      }
      
      if (reviewStatusCol > 0) {
        sheet.getRange(rowIndex, reviewStatusCol).setValue("VOLLTEXT_AKTUALISIERT");
      }
      
      // 3. Färbe Zeile orange (Warnung)
      sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).setBackground("#fff3cd");
      
      // 4. ✅ PROMPTS AUTOMATISCH NEU GENERIEREN
      try {
        Logger.log("🤖 Starte automatische Prompt-Regenerierung...");
        
        const promptsGenerated = autoRegeneratePrompts(rowIndex);
        
        if (promptsGenerated) {
          // Erfolg
          if (statusCol > 0) {
            sheet.getRange(rowIndex, statusCol).setValue("✅ Prompts neu generiert - Bitte Gemini starten!");
          }
          
          Logger.log("✅ Prompts erfolgreich neu generiert");
          
        } else {
          // Fehler
          if (statusCol > 0) {
            sheet.getRange(rowIndex, statusCol).setValue("⚠️ Volltext hinzugefügt - Bitte Prompts manuell neu generieren");
          }
          
          Logger.log("⚠️ Prompt-Regenerierung fehlgeschlagen");
        }
        
      } catch (promptError) {
        Logger.log("❌ Fehler bei Prompt-Regenerierung: " + promptError.message);
        
        if (statusCol > 0) {
          sheet.getRange(rowIndex, statusCol).setValue("⚠️ Volltext hinzugefügt - Bitte Prompts manuell neu generieren");
        }
      }
      
      // 5. Log
      if (typeof logAction === 'function') {
        logAction("Volltext-Update", "Prompts neu generiert für: " + uuid);
      }
      
    } else {
      // ===== KEINE ALTE ANALYSE =====
      Logger.log("✅ Volltext hinzugefügt für neues Paper: " + uuid);
    }
    
  } catch (e) {
    Logger.log("❌ Fehler in onVolltextAdded: " + e.message);
  }
}

/**
 * ✅ Automatische Prompt-Regenerierung
 */
function autoRegeneratePrompts(rowIndex) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Hole Row-Daten
    const rowData = {};
    headers.forEach((header, index) => {
      rowData[header] = sheet.getRange(rowIndex, index + 1).getValue();
    });
    
    const uuid = rowData["UUID"];
    const titel = rowData["Titel"];
    
    Logger.log("📝 Generiere Prompts für: " + titel);
    
    // Bestimme welche Prompts generiert werden sollen
    const promptsToGenerate = determinePromptsForRow(rowData);
    
    Logger.log("🎯 Zu generierende Prompts: " + promptsToGenerate.join(", "));
    
    let successCount = 0;
    let errorCount = 0;
    
    // Generiere jeden Prompt
    for (const phase of promptsToGenerate) {
      try {
        const prompt = buildPromptForPhase(rowData, phase);
        
        // Speichere Prompt
        const saved = savePromptToSheet(uuid, phase, prompt);
        
        if (saved) {
          successCount++;
          Logger.log(`✅ ${phase} prompt generiert`);
        } else {
          errorCount++;
          Logger.log(`⚠️ ${phase} prompt konnte nicht gespeichert werden`);
        }
        
      } catch (phaseError) {
        errorCount++;
        Logger.log(`❌ Fehler bei ${phase}: ` + phaseError.message);
      }
    }
    
    Logger.log(`📊 Ergebnis: ${successCount} erfolgreich, ${errorCount} Fehler`);
    
    return successCount > 0;
    
  } catch (e) {
    Logger.log("❌ autoRegeneratePrompts Fehler: " + e.message);
    return false;
  }
}

/**
 * ✅ Bestimmt welche Prompts generiert werden sollen
 */
function determinePromptsForRow(rowData) {
  const prompts = [];
  
  // Hauptkategorie vorhanden?
  if (!rowData["Hauptkategorie"] || rowData["Hauptkategorie"] === "") {
    prompts.push("triage");
  }
  
  // Metadaten vollständig?
  if (!rowData["Jahr"] || rowData["Jahr"] === "N/A" || rowData["Jahr"] === "") {
    prompts.push("metadata_extract");
  }
  
  // Analyse immer neu generieren (da Volltext jetzt vorhanden)
  prompts.push("analysis");
  
  // Review-Prompt hinzufügen
  prompts.push("analysis_review");
  
  return prompts;
}

/**
 * ✅ Baut Prompt für eine Phase
 * Mit intelligenter Volltext-Extraktion (Google Doc Support)
 */
function buildPromptForPhase(data, phase) {
  const templateKey = {
    "triage": "PROMPT_TEMPLATE_TRIAGE",
    "metadata_extract": "PROMPT_TEMPLATE_METADATA_EXTRACT",
    "analysis": "PROMPT_TEMPLATE_ANALYSIS",
    "analysis_review": "PROMPT_TEMPLATE_REVIEW"
  }[phase] || "PROMPT_TEMPLATE_ANALYSIS";
  
  let template = getConfig(templateKey);
  if (!template) {
    throw new Error(`Prompt-Template ${templateKey} nicht gefunden`);
  }
  
  const lists = loadCategoriesAndKeywords();
  
  // ✅ Intelligente Volltext-Extraktion
  let volltextContent = getVolltextForPrompt(data);
  
  template = template
    .replace(/{TITEL}/g, data.Titel || "N/A")
    .replace(/{AUTOREN}/g, data.Autoren || "N/A")
    .replace(/{JAHR}/g, data.Jahr || "N/A")
    .replace(/{JOURNAL}/g, data["Journal/Quelle"] || "N/A")
    .replace(/{ABSTRACT}/g, data["Inhalt/Abstract"] || "N/A")
    .replace(/{VOLLTEXT}/g, volltextContent)
    .replace(/{PUBLIKATIONSTEXT}/g, volltextContent)
    .replace(/{KATEGORIEN}/g, lists.categories.join(", "))
    .replace(/{SCHLAGWOERTER}/g, lists.keywords.join(", "))
    .replace(/{UUID}/g, data.UUID);
  
  return template;
}

/**
 * ✅ Intelligente Volltext-Extraktion
 * Liest Google Docs automatisch aus
 */
function getVolltextForPrompt(data) {
  const volltextField = data["Volltext/Extrakt"] || "";
  const volltextLink = data["Volltext-Datei-Link"] || "";
  
  // Fall 1: Kein Volltext
  if (!volltextField && !volltextLink) {
    return data["Inhalt/Abstract"] || "N/A";
  }
  
  // Fall 2: Volltext direkt in Zelle
  if (volltextField && volltextField.length > 200 && !volltextField.includes("[📄 Text zu lang")) {
    return volltextField;
  }
  
  // Fall 3: Text als Google Doc gespeichert
  if (volltextField.includes("[📄 Text zu lang") || volltextField.includes("Volltext-Doc:")) {
    const docLinkMatch = volltextLink.match(/docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/);
    
    if (docLinkMatch) {
      const docId = docLinkMatch[1];
      
      try {
        // Lese Text aus Google Doc
        const doc = DocumentApp.openById(docId);
        const fullText = doc.getBody().getText();
        
        Logger.log("📄 Volltext aus Google Doc geladen: " + fullText.length + " Zeichen");
        return fullText;
        
      } catch (e) {
        Logger.log("⚠️ Konnte Google Doc nicht öffnen: " + e.message);
        return data["Inhalt/Abstract"] || "N/A";
      }
    }
  }
  
  // Fallback: Abstract
  return data["Inhalt/Abstract"] || "N/A";
}

/**
 * ✅ Speichert Prompt ins Sheet
 */
function savePromptToSheet(uuid, phase, promptText) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let promptSheet = ss.getSheetByName("Manuelle Prompts");
    
    if (!promptSheet) {
      Logger.log("⚠️ 'Manuelle Prompts' Sheet nicht gefunden");
      return false;
    }
    
    // Prüfe ob bereits vorhanden
    const data = promptSheet.getDataRange().getValues();
    let existingRow = -1;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === uuid && data[i][1] === phase) {
        existingRow = i + 1;
        break;
      }
    }
    
    const timestamp = new Date();
    const status = "NEU_GENERIERT";
    
    if (existingRow > 0) {
      // Update
      promptSheet.getRange(existingRow, 3).setValue(promptText);
      promptSheet.getRange(existingRow, 4).setValue(timestamp);
      promptSheet.getRange(existingRow, 5).setValue(status);
      
      Logger.log("✏️ Prompt aktualisiert in Zeile " + existingRow);
      
    } else {
      // Neu
      promptSheet.appendRow([
        uuid,
        phase,
        promptText,
        timestamp,
        status
      ]);
      
      Logger.log("➕ Neuer Prompt hinzugefügt");
    }
    
    return true;
    
  } catch (e) {
    Logger.log("❌ savePromptToSheet Fehler: " + e.message);
    return false;
  }
}

/**
 * ✅ Archiviert alte Analyse
 */
function archiveOldAnalysisOnVolltextUpdate(uuid) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashSheet = ss.getSheetByName("Dashboard");
    const headers = dashSheet.getRange(1, 1, 1, dashSheet.getLastColumn()).getValues()[0];
    
    // Finde Zeile
    const data = dashSheet.getDataRange().getValues();
    let rowData = null;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === uuid) {
        rowData = {};
        headers.forEach((header, idx) => {
          rowData[header] = data[i][idx];
        });
        break;
      }
    }
    
    if (!rowData) return;
    
    // Archiv-Sheet
    let archiveSheet = ss.getSheetByName("Analyse-Archiv-Volltext");
    
    if (!archiveSheet) {
      archiveSheet = ss.insertSheet("Analyse-Archiv-Volltext");
      archiveSheet.appendRow([
        "Archiviert am", "Grund", "UUID", "Titel",
        "Haupterkenntnis (alt)", "Zusammenfassung (alt)", "Relevanz (alt)"
      ]);
    }
    
    archiveSheet.appendRow([
      new Date(),
      "Volltext nachträglich hinzugefügt - Prompts neu generiert",
      uuid,
      rowData["Titel"],
      rowData["Haupterkenntnis"],
      rowData["Zusammenfassung"],
      rowData["Relevanz"]
    ]);
    
    Logger.log("📦 Alte Analyse archiviert für UUID: " + uuid);
    
  } catch (e) {
    Logger.log("❌ Fehler in archiveOldAnalysisOnVolltextUpdate: " + e.message);
  }
}

/**
 * ✅ Bestimmt Phase basierend auf Row-Daten
 */
function determinePhaseForRow(data) {
  if (!data.Hauptkategorie) return "triage";
  if (!data.Jahr || data.Jahr === "N/A") return "metadata_extract";
  if (!data.Haupterkenntnis) return "analysis";
  if (data["Review-Status"] === "UNGEPRÜFT") return "analysis_review";
  return "analysis";
}
