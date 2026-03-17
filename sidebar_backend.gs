// FILE: sidebar_backend.gs
/**
 * ==========================================
 * SIDEBAR BACKEND
 * ==========================================
 * Prompt-Generierung und Daten-Verwaltung für Sidebar
 */

/**
 * ✅ Haupt-Funktion für Sidebar-Daten
 * Mit intelligenter Volltext-Erkennung
 */
function getSidebarData(forcedStep) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rowIndex = sheet.getActiveCell().getRow();
  
  if (rowIndex < 2) return { error: "Bitte eine Zeile im Dashboard wählen." };
  
  try {
    const rowData = getRowData(sheet, rowIndex);
    const lists = getCategoriesAndTags();
    
    // ✅ Intelligente Volltext-Erkennung
    let sourceText = "";
    let sourceInfo = "";
    
    const volltextExtraktCell = String(rowData["Volltext/Extrakt"] || "").trim();
    const volltextFileLink = String(rowData["Volltext-Datei-Link"] || "").trim();
    const manuell = String(rowData["Manueller Volltext"] || "").trim();
    const abstract = String(rowData["Inhalt/Abstract"] || "").trim();
    
    // PRIORISIERUNG:
    // 1. Google Doc Link (bei großen PDFs)
    // 2. Volltext/Extrakt direkt (bei kleinen PDFs)
    // 3. Manueller Volltext
    // 4. Abstract
    
    if (volltextExtraktCell.includes("Text zu lang") && volltextFileLink.includes("docs.google.com")) {
      // Fall: Großes PDF → Text ist in Google Doc
      sourceText = extractTextFromGoogleDoc(volltextFileLink);
      sourceInfo = "VOLLTEXT (aus Google Doc)";
      
    } else if (volltextExtraktCell.length > 500) {
      // Fall: Normaler Volltext direkt in Zelle
      sourceText = volltextExtraktCell;
      sourceInfo = "VOLLTEXT (Dashboard)";
      
    } else if (manuell.length > 500) {
      // Fall: Manuell hinzugefügter Volltext
      sourceText = manuell;
      sourceInfo = "VOLLTEXT (Booster)";
      
    } else if (abstract.length > 50) {
      // Fall: Nur Abstract
      sourceText = abstract;
      sourceInfo = "ABSTRACT (unvollständig)";
      
    } else {
      // Fall: Nur Titel
      sourceText = rowData["Titel"] || "";
      sourceInfo = "NUR TITEL (sehr unvollständig)";
    }
    
    // Schritt bestimmen
    let promptKey = forcedStep;
    if (!promptKey) {
      promptKey = determineCurrentStep(rowData, sourceInfo);
    }
    
    const steps = getAvailableSteps();
    const stepObj = steps.find(s => s.id === promptKey);
    const stepTitle = stepObj ? stepObj.label : "Manueller Schritt";
    let promptTemplate = getPrompt(promptKey);
    
    // Prompt füllen
    let filledPrompt = promptTemplate
      .replace(/{PUBLIKATIONSTEXT}/g, sourceText)
      .replace(/{VOLLTEXT}/g, sourceText)
      .replace(/{KATEGORIEN_LISTE}/g, lists.categories)
      .replace(/{KATEGORIE_MAPPING}/g, lists.fullMapping)
      .replace(/{SCHLAGWÖRTER_LISTE}/g, lists.tags)
      .replace(/{DOI}/g, rowData["DOI"] || "N/A")
      .replace(/{TITEL}/g, rowData["Titel"] || "N/A")
      .replace(/{AUTOREN}/g, rowData["Autoren"] || "N/A")
      .replace(/{JAHR}/g, rowData["Jahr"] || "N/A")
      .replace(/{JOURNAL}/g, rowData["Journal/Quelle"] || "N/A")
      .replace(/{LINK}/g, rowData["Link zum Volltext"] || rowData["Link"] || "N/A");
    
    // PICO hinzufügen wenn relevant
    const picoData = `\n\n### BEREITS EXTRAHIERTE PICO-DATEN ###\n` +
                     `Population: ${rowData["PICO Population"] || "N/A"}\n` +
                     `Intervention: ${rowData["PICO Intervention"] || "N/A"}\n` +
                     `Comparator: ${rowData["PICO Comparator"] || "N/A"}\n` +
                     `Outcomes: ${rowData["PICO Outcomes"] || "N/A"}`;
    
    if (promptKey === "TRIAGE" || promptKey === "MASTER_ANALYSE") {
      filledPrompt += picoData;
    }
    
    // Quelleninfo
    filledPrompt += `\n\n(HINWEIS: Textquelle: ${sourceInfo}, ${sourceText.length} Zeichen)`;
    
    return {
      row: rowIndex,
      title: rowData["Titel"],
      stepInfo: stepTitle,
      currentStepId: promptKey,
      prompt: filledPrompt,
      detectedSource: sourceInfo
    };
  } catch (e) {
    return { error: "Fehler in getSidebarData: " + e.message };
  }
}

/**
 * ✅ Extrahiert Text aus Google Doc
 */
function extractTextFromGoogleDoc(docUrl) {
  try {
    // Extrahiere Doc ID aus URL
    const docIdMatch = docUrl.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!docIdMatch) {
      return "[FEHLER: Ungültige Google Doc URL]";
    }
    
    const docId = docIdMatch[1];
    
    // Öffne und lese Doc
    const doc = DocumentApp.openById(docId);
    const text = doc.getBody().getText();
    
    Logger.log("Google Doc gelesen: " + text.length + " Zeichen");
    
    return text;
    
  } catch (e) {
    Logger.log("Fehler beim Lesen des Google Docs: " + e.message);
    return "[FEHLER: Konnte Google Doc nicht lesen - " + e.message + "]";
  }
}

/**
 * ✅ Bestimmt aktuellen Workflow-Schritt
 */
function determineCurrentStep(rowData, sourceInfo) {
  const hasValue = (field) => {
    const val = String(rowData[field] || "").trim();
    return val.length > 0 && val !== "N/A" && val !== "undefined";
  };
  
  const reviewStatus = String(rowData["Review-Status"] || "").trim().toUpperCase();
  
  // Schritt 0: Kein Text vorhanden → Researcher
  if (sourceInfo.includes("NUR TITEL")) {
    return "RESEARCHER";
  }
  
  // Schritt 1: Keine Kategorie → Triage
  if (!hasValue("Hauptkategorie")) {
    return "TRIAGE";
  }
  
  // Schritt 2: Keine/unvollständige Metadaten → Metadaten
  if (!hasValue("Jahr") || !hasValue("Autoren")) {
    return "METADATEN";
  }
  
  // Schritt 3: Keine Haupterkenntnis → Master-Analyse
  if (!hasValue("Haupterkenntnis")) {
    return "MASTER_ANALYSE";
  }
  
  // Ab hier: Haupterkenntnis existiert
  
  // Schritt 4: Nach Analyse, vor Redaktion
  if (reviewStatus === "" || reviewStatus === "UNGEPRÜFT" || reviewStatus === "ANALYSE_FERTIG") {
    return "REDAKTION";
  }
  
  // Schritt 5: Nach Redaktion → Faktencheck
  if (reviewStatus === "REDAKTION_OK") {
    return "FAKTENCHECK";
  }
  
  // Schritt 6: Nach Faktencheck → Senior Review
  if (reviewStatus === "CHECK_ERFOLGT") {
    return "REVIEW";
  }
  
  // Fertig
  if (reviewStatus === "FERTIG") {
    return "REVIEW";
  }
  
  // Fallback
  return "MASTER_ANALYSE";
}

/**
 * ✅ Verfügbare Schritte
 */
function getAvailableSteps() {
  return [
    {id: "RESEARCHER", label: "Schritt 0: Text-Suche (B15)"},
    {id: "TRIAGE", label: "Schritt 1: Triage (B19)"},
    {id: "METADATEN", label: "Schritt 2: Metadaten (B17)"},
    {id: "MASTER_ANALYSE", label: "Schritt 3: Strategische Analyse (B16)"},
    {id: "REDAKTION", label: "Schritt 4: Redaktion (B18)"},
    {id: "FAKTENCHECK", label: "Schritt 5: Faktencheck (B20)"},
    {id: "REVIEW", label: "Schritt 6: Senior Review (B21)"}
  ];
}

/**
 * ✅ Speichert Sidebar-Daten zurück ins Dashboard
 */
function saveSidebarData(jsonString, sourceInfo) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const row = sheet.getActiveRange().getRow();
  
  try {
    const data = parseJsonFromLlm(jsonString);
    writeResultsToSheetExtended(sheet, row, data);
    
    if (sourceInfo) setCell(sheet, row, "Volltext-Status", sourceInfo);
    
    // Hole aktuelle Daten NACH dem Schreiben
    const rd = getRowData(sheet, row);
    const reviewStatus = String(rd["Review-Status"] || "").trim().toUpperCase();
    const hasHaupterkenntnis = String(rd["Haupterkenntnis"] || "").trim().length > 10;
    const kritischeBewertung = String(rd["Kritische Bewertung"] || "");
    
    // Status-Fortschritt
    if (hasHaupterkenntnis && (reviewStatus === "" || reviewStatus === "UNGEPRÜFT")) {
      setCell(sheet, row, "Review-Status", "ANALYSE_FERTIG");
      return { message: "✅ Analyse gespeichert! Weiter mit Redaktion (Schritt 4)" };
      
    } else if (reviewStatus === "ANALYSE_FERTIG") {
      setCell(sheet, row, "Review-Status", "REDAKTION_OK");
      return { message: "✅ Redaktion gespeichert! Weiter mit Faktencheck (Schritt 5)" };
      
    } else if (reviewStatus === "REDAKTION_OK") {
      setCell(sheet, row, "Review-Status", "CHECK_ERFOLGT");
      return { message: "✅ Faktencheck gespeichert! Weiter mit Senior Review (Schritt 6)" };
      
    } else if (reviewStatus === "CHECK_ERFOLGT") {
      setCell(sheet, row, "Review-Status", "FERTIG");
      setCell(sheet, row, "Status", "Abgeschlossen");
      
      const hatFaktenWarnung = kritischeBewertung.includes("FAKTEN-WARNUNG");
      const warKorrekturNoetig = (data.ist_korrekt === false);
      
      if (hatFaktenWarnung || warKorrekturNoetig) {
        sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#f4cccc");
      } else {
        sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground("#d9ead3");
      }
      
      return { message: "✅ Review abgeschlossen! Paper ist fertig." };
    }
    
    return { message: "✅ Schritt gespeichert!" };
    
  } catch (e) {
    return { error: "Fehler beim Speichern: " + e.message };
  }
}
