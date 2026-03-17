// FILE: geminiSheets.gs

/**
 * ==========================================
 * GEMINI SHEETS BATCH WORKFLOW
 * ==========================================
 * 
 * Workflow:
 * 1. Zeilen im Dashboard auswählen
 * 2. Menü: "Schritt X vorbereiten" → Formeln in Spalte schreiben
 * 3. Auf "Generieren" klicken (Gemini führt aus)
 * 4. Menü: "Ergebnisse übernehmen" → JSON ins Dashboard verteilen
 * 
 * SPALTEN-MAPPING:
 * - AY (51) = Researcher (Phase 0)
 * - AZ (52) = Triage (Phase 1)
 * - BA (53) = Metadaten (Phase 2)
 * - BB (54) = Master Analyse (Phase 3)
 * - BC (55) = Redaktion (Phase 4)
 * - BD (56) = Faktencheck (Phase 5)
 * - BE (57) = Review (Phase 6)
 */

// ==========================================
// KONFIGURATION
// ==========================================

const GEMINI_COLUMNS = {
  RESEARCHER:    56,   // BD
  TRIAGE:        57,   // BE
  METADATEN:     58,   // BF
  MASTER_ANALYSE:59,   // BG
  REDAKTION:     60,   // BH
  FAKTENCHECK:   61,   // BI
  REVIEW:        62    // BJ
};

// PROMPT_CELLS bleibt unverändert:
const PROMPT_CELLS = {
  RESEARCHER:    "B15",
  TRIAGE:        "B19",
  METADATEN:     "B17",
  MASTER_ANALYSE:"B16",
  REDAKTION:     "B18",
  FAKTENCHECK:   "B20",
  REVIEW:        "B21"
};

// ==========================================
// PHASE 0: RESEARCHER
// ==========================================

function prepareStep0Researcher() {
  prepareGeminiStep("RESEARCHER", GEMINI_COLUMNS.RESEARCHER, PROMPT_CELLS.RESEARCHER);
}

// ==========================================
// PHASE 1: TRIAGE
// ==========================================

function prepareStep1Triage() {
  prepareGeminiStep("TRIAGE", GEMINI_COLUMNS.TRIAGE, PROMPT_CELLS.TRIAGE);
}

// ==========================================
// PHASE 2: METADATEN
// ==========================================

function prepareStep2Metadaten() {
  prepareGeminiStep("METADATEN", GEMINI_COLUMNS.METADATEN, PROMPT_CELLS.METADATEN);
}

// ==========================================
// PHASE 3: MASTER ANALYSE
// ==========================================

function prepareStep3Analyse() {
  prepareGeminiStep("MASTER_ANALYSE", GEMINI_COLUMNS.MASTER_ANALYSE, PROMPT_CELLS.MASTER_ANALYSE);
}

// ==========================================
// PHASE 4: REDAKTION
// ==========================================

function prepareStep4Redaktion() {
  prepareGeminiStep("REDAKTION", GEMINI_COLUMNS.REDAKTION, PROMPT_CELLS.REDAKTION);
}

// ==========================================
// PHASE 5: FAKTENCHECK
// ==========================================

function prepareStep5Faktencheck() {
  prepareGeminiStep("FAKTENCHECK", GEMINI_COLUMNS.FAKTENCHECK, PROMPT_CELLS.FAKTENCHECK);
}

// ==========================================
// PHASE 6: REVIEW
// ==========================================

function prepareStep6Review() {
  prepareGeminiStep("REVIEW", GEMINI_COLUMNS.REVIEW, PROMPT_CELLS.REVIEW);
}

// ==========================================
// CORE: PROMPT VORBEREITUNG
// ==========================================

function prepareGeminiStep(phaseId, columnIndex, promptCell) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  const configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  const selection = dashSheet.getActiveRange();
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
  
  const promptTemplate = configSheet.getRange(promptCell).getValue();
  if (!promptTemplate) {
    SpreadsheetApp.getUi().alert(`Prompt für ${phaseId} nicht gefunden in ${promptCell}`);
    return;
  }
  
  const lists = loadCategoriesAndKeywords();
  
  let prepared = 0;
  const errors = [];
  
  for (let i = 0; i < numRows; i++) {
    const rowIndex = startRow + i;
    
    try {
      const rowData = getRowDataByIndex(dashSheet, rowIndex);
      const filledPrompt = buildPromptForRow(promptTemplate, rowData, lists, phaseId);
      const geminiFormula = buildGeminiFormula(filledPrompt);
      
      dashSheet.getRange(rowIndex, columnIndex).setFormula(geminiFormula);
      updateDashboardField(rowData.UUID, "Batch-Phase", `${phaseId} - Prompt bereit`);
      
      prepared++;
      
    } catch (e) {
      errors.push(`Zeile ${rowIndex}: ${e.message}`);
      logError("prepareGeminiStep", e);
    }
  }
  
  let message = `✅ ${prepared} Prompts für ${phaseId} vorbereitet!\n\n`;
  message += `Die Formeln stehen in Spalte ${getColumnLetter(columnIndex)}.\n\n`;
  message += `Nächste Schritte:\n`;
  message += `1. Klicke auf erste Zelle in Spalte ${getColumnLetter(columnIndex)}\n`;
  message += `2. Klicke "Generieren und einfügen"\n`;
  message += `3. Warte bis alle Zeilen fertig sind\n`;
  message += `4. Menü → Gemini Workflow → "Ergebnisse übernehmen"`;
  
  if (errors.length > 0) {
    message += `\n\n⚠️ Fehler bei ${errors.length} Zeilen:\n`;
    message += errors.slice(0, 3).join("\n");
  }
  
  SpreadsheetApp.getUi().alert(message);
  logAction("Gemini Batch", `${prepared} Prompts für ${phaseId} vorbereitet`);
}

function buildPromptForRow(template, rowData, lists, phaseId) {
  let prompt = template;
  
  prompt = prompt
    .replace(/{PUBLIKATIONSTEXT}/g, getCombinedVolltext(rowData))
    .replace(/{TITEL}/g, rowData.Titel || "N/A")
    .replace(/{DOI}/g, rowData.DOI || "N/A")
    .replace(/{AUTOREN}/g, rowData.Autoren || "N/A")
    .replace(/{JAHR}/g, rowData["Publikationsdatum"] || rowData.Jahr || "N/A")
    .replace(/{JOURNAL}/g, rowData["Journal/Quelle"] || "N/A")
    .replace(/{ABSTRACT}/g, rowData["Inhalt/Abstract"] || "N/A")
    .replace(/{VOLLTEXT}/g, rowData["Volltext/Extrakt"] || rowData["Inhalt/Abstract"] || "N/A")
    .replace(/{PDF_LINK}/g, rowData["Volltext-Datei-Link"] || "Kein PDF verfügbar")
    .replace(/{KATEGORIEN_LISTE}/g, lists.categories.join(", "))
    .replace(/{KATEGORIE_MAPPING}/g, lists.fullMapping)
    .replace(/{SCHLAGWÖRTER_LISTE}/g, lists.keywords.join(", "))
    .replace(/{UUID}/g, rowData.UUID);
  
  if (phaseId === "FAKTENCHECK" || phaseId === "REDAKTION") {
    prompt = prompt.replace(/{KERNAUSSAGE}/g, rowData.Haupterkenntnis || "N/A");
  }
  
  if (phaseId === "REVIEW") {
    const analysisJson = buildAnalysisJsonForReview(rowData);
    prompt = prompt.replace(/{JSON_ANALYSE_VON_SCHRITT_1}/g, analysisJson);
  }
  
  return prompt;
}

function buildGeminiFormula(prompt) {
  const cleanPrompt = prompt
    .replace(/\r\n/g, ' ')
    .replace(/\n/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/"/g, "'")
    .replace(/\\/g, '')
    .replace(/\t/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .substring(0, 25000);
  
  return `=GEMINI("${cleanPrompt}")`;
}

// ==========================================
// ERGEBNISSE VERARBEITEN
// ==========================================

function processGeminiResults() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Gemini Ergebnisse übernehmen',
    'Welche Phase?\n\n' +
    '0 = Researcher (BD)\n' +
    '1 = Triage (BE)\n' +
    '2 = Metadaten (BF)\n' +
    '3 = Master Analyse (BG)\n' +
    '4 = Redaktion (BH)\n' +
    '5 = Faktencheck (BI)\n' +
    '6 = Review (BJ)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const phaseNum = parseInt(response.getResponseText().trim());
  
  const phaseMap = {
    0: { id: "RESEARCHER", col: GEMINI_COLUMNS.RESEARCHER },
    1: { id: "TRIAGE", col: GEMINI_COLUMNS.TRIAGE },
    2: { id: "METADATEN", col: GEMINI_COLUMNS.METADATEN },
    3: { id: "MASTER_ANALYSE", col: GEMINI_COLUMNS.MASTER_ANALYSE },
    4: { id: "REDAKTION", col: GEMINI_COLUMNS.REDAKTION },
    5: { id: "FAKTENCHECK", col: GEMINI_COLUMNS.FAKTENCHECK },
    6: { id: "REVIEW", col: GEMINI_COLUMNS.REVIEW }
  };
  
  const phase = phaseMap[phaseNum];
  if (!phase) {
    ui.alert("Ungültige Phase: " + phaseNum);
    return;
  }
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  if (sheet.getLastRow() <= 1) {
    ui.alert("Keine Daten im Dashboard");
    return;
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, DASHBOARD_HEADERS.length).getValues();
  
  let processed = 0;
  let skipped = 0;
  let rejected = 0;
  const errors = [];
  
  data.forEach((row, index) => {
    const rowIndex = index + 2;
    const uuid = row[0];
    
    const geminiResponse = sheet.getRange(rowIndex, phase.col).getValue();
    
    if (!geminiResponse || geminiResponse === "") {
      skipped++;
      return;
    }
    
    const responseStr = geminiResponse.toString();
    
    if (responseStr.startsWith("=GEMINI")) {
      skipped++;
      return;
    }
    
    if (responseStr.includes("I'm still learning") || 
        responseStr.includes("I\"m still learning") ||
        responseStr.includes("can't help with that")) {
      
      sheet.getRange(rowIndex, phase.col).setBackground("#ffe599");
      updateDashboardField(uuid, "Fehler-Details", "Gemini hat Anfrage abgelehnt - bitte manuell bearbeiten");
      updateDashboardField(uuid, "Batch-Phase", `${phase.id} - Gemini Ablehnung ⚠️`);
      
      rejected++;
      return;
    }
    
    try {
      const jsonData = parseJsonFromLlm(responseStr);
      applyJSONToDashboardPhaseAware(uuid, phase.id, jsonData);
      updateDashboardField(uuid, "Batch-Phase", `${phase.id} - Abgeschlossen ✓`);
      sheet.getRange(rowIndex, phase.col).setBackground("#d9ead3");
      processed++;
      
    } catch (e) {
      errors.push(`Zeile ${rowIndex}: ${e.message}`);
      logError("processGeminiResults", e);
      sheet.getRange(rowIndex, phase.col).setBackground("#f4cccc");
    }
  });
  
  let message = `✅ ${processed} Ergebnisse verarbeitet\n`;
  message += `⏭️ ${skipped} übersprungen (leer/Formel)\n`;
  
  if (errors.length > 0) {
    message += `\n❌ Fehler bei ${errors.length} Zeilen:\n`;
    message += errors.slice(0, 5).join("\n");
  }
  
  ui.alert(message);
  logAction("Gemini Batch", `${processed} Ergebnisse für ${phase.id} verarbeitet`);
}

// ==========================================
// HELPER
// ==========================================

function getColumnLetter(columnNumber) {
  let letter = '';
  while (columnNumber > 0) {
    const remainder = (columnNumber - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    columnNumber = Math.floor((columnNumber - 1) / 26);
  }
  return letter;
}

function showGeminiHelp() {
  const ui = SpreadsheetApp.getUi();
  
  const helpText = 
    '🤖 GEMINI ANALYSE-SYSTEM\n\n' +
    '═══════════════════════════════════════════════════════════\n' +
    'ABLAUF:\n' +
    '═══════════════════════════════════════════════════════════\n\n' +
    '1️⃣ ANALYSE STARTEN\n' +
    '   → Menü: Gemini → Analyse durchführen\n' +
    '   → Analysiert Papers mit Status "Volltext vorhanden"\n' +
    '   → Extrahiert Kerninformationen mit KI\n\n' +
    '2️⃣ REVIEW-PHASE\n' +
    '   → Menü: Gemini → Review starten\n' +
    '   → Zeigt analysierte Papers zur Kontrolle\n' +
    '   → Ermöglicht manuelle Korrekturen\n\n' +
    '3️⃣ FINAL ÜBERNEHMEN\n' +
    '   → Menü: Gemini → Final übernehmen\n' +
    '   → Schreibt geprüfte Daten ins Dashboard\n\n' +
    '═══════════════════════════════════════════════════════════\n' +
    'EXTRAHIERTE FELDER:\n' +
    '═══════════════════════════════════════════════════════════\n\n' +
    '📊 METADATEN\n' +
    '   • Artikeltyp (RCT, Review, Case Study...)\n' +
    '   • Studiendesign\n' +
    '   • Teilnehmerzahl\n' +
    '   • Publikationsjahr\n\n' +
    '🎯 CLINICAL CONTEXT\n' +
    '   • Hauptdiagnosen (ICD-Codes)\n' +
    '   • Interventionen\n' +
    '   • Outcomes & Effektgrößen\n\n' +
    '🔬 TECHNOLOGIE\n' +
    '   • Medizinprodukte\n' +
    '   • Algorithmen & KI-Modelle\n' +
    '   • Software & Plattformen\n\n' +
    '💡 SCHLÜSSELERKENNTNISSE\n' +
    '   • Zusammenfassung (2-3 Sätze)\n' +
    '   • Hauptergebnisse\n' +
    '   • Limitationen\n\n' +
    '═══════════════════════════════════════════════════════════\n' +
    'QUALITÄTSCHECKS:\n' +
    '═══════════════════════════════════════════════════════════\n\n' +
    '✅ Validierung:\n' +
    '   • JSON-Struktur wird automatisch geprüft\n' +
    '   • Fehlende Pflichtfelder werden erkannt\n' +
    '   • Inkonsistenzen werden markiert\n\n' +
    '🔄 Retry-Logik:\n' +
    '   • Bis zu 3 Versuche bei Fehlern\n' +
    '   • Automatische Fehlerbehandlung\n\n' +
    '═══════════════════════════════════════════════════════════\n' +
    'TIPPS:\n' +
    '═══════════════════════════════════════════════════════════\n\n' +
    '• Starte mit kleinen Batches (5-10 Papers)\n' +
    '• Prüfe Review-Ergebnisse sorgfältig\n' +
    '• Nutze die Logs für Fehlersuche\n' +
    '• API-Key muss in Script Properties gesetzt sein\n\n' +
    '═══════════════════════════════════════════════════════════\n';
  
  ui.alert('📚 Gemini Analyse-System Hilfe', helpText, ui.ButtonSet.OK);
}


function clearAllGeminiColumns() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Alle Gemini-Spalten löschen?',
    'Dies löscht ALLE Formeln und Antworten in den Spalten BD-BJ.\n\nFortfahren?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  if (sheet.getLastRow() <= 1) return;
  
  const range = sheet.getRange(2, 56, sheet.getLastRow() - 1, 7);
  range.clear();
  range.setBackground(null);
  
  ui.alert('Alle Gemini-Spalten gelöscht');
  logAction("Gemini Batch", "Alle Spalten gelöscht");
}

function applyJSONToDashboardPhaseAware(uuid, phaseId, jsonData) {
  const data = jsonData;
  
  if (!data || typeof data !== 'object') {
    throw new Error("JSON ist ungültig");
  }
  
  Logger.log(`Verarbeite ${phaseId} für ${uuid}, Felder: ${Object.keys(data).join(", ")}`);
  
  if (phaseId === "TRIAGE") {
    const allowedCategories = ["Diabetes Typ 1", "Diabetes Typ 2", "Immunologie/Onkologie"];
    
    if (data.kategorie) {
      const kat = data.kategorie;
      
      if (kat.hauptkategorie) {
        let category = kat.hauptkategorie.trim();
        
        if (!allowedCategories.includes(category)) {
          Logger.log(`⚠️ WARNUNG: Ungültige Kategorie "${category}"`);
          
          const lower = category.toLowerCase();
            if (lower.includes("t1d") || (lower.includes("typ") && lower.includes("1"))) {
            category = "Diabetes Typ 1";
          } else if (lower.includes("t2d") || (lower.includes("typ") && lower.includes("2"))) {
            category = "Diabetes Typ 2";
          } else if (lower.includes("diabetes") && !lower.includes("typ")) {
            category = "Diabetes Typ 1";
            updateDashboardField(uuid, "Fehler-Details", `Generische Diabetes-Kategorie - bitte Typ 1/2 manuell prüfen`);
          } else if (lower.includes("immu") || lower.includes("onko")) {
            category = "Immunologie/Onkologie";
          }
          
          Logger.log(`   → Auto-Korrektur zu: "${category}"`);
        }
        
        updateDashboardField(uuid, "Hauptkategorie", category);
      }
      
      if (kat.unterkategorien) {
        const subs = Array.isArray(kat.unterkategorien) ? kat.unterkategorien.join(", ") : kat.unterkategorien;
        updateDashboardField(uuid, "Unterkategorien", subs);
      }
      
      if (kat.schlagwoerter) {
        const tags = Array.isArray(kat.schlagwoerter) ? kat.schlagwoerter.join(", ") : kat.schlagwoerter;
        updateDashboardField(uuid, "Schlagwörter", tags);
      }
    }
    
    if (data.relevanz) {
      updateDashboardField(uuid, "Relevanz", data.relevanz);
    }
    
    if (data.relevanz_begruendung) {
      updateDashboardField(uuid, "Relevanz-Begründung", data.relevanz_begruendung);
    }
    
    if (data.evidenzgrad) {
      updateDashboardField(uuid, "Evidenzgrad", data.evidenzgrad);
    }
    
    if (data.produkt_erwähnt) {
      const produkte = data.produkt_erwähnt;
      let produktListe = "";
      
      if (produkte.kernprodukte && produkte.kernprodukte.length > 0) {
        produktListe += "Kernprodukte: " + produkte.kernprodukte.join(", ") + "\n";
      }
      if (produkte.konkurrenz && produkte.konkurrenz.length > 0) {
        produktListe += "Konkurrenz: " + produkte.konkurrenz.join(", ");
      }
      
      if (produktListe) {
        updateDashboardField(uuid, "Produkt-Fokus", produktListe.trim());
      }
    }
  }
  
  else if (phaseId === "METADATEN") {
    if (data.pmid) updateDashboardField(uuid, "PMID", data.pmid);
    if (data.doi || data.doi_string) updateDashboardField(uuid, "DOI", data.doi || data.doi_string);
    if (data.jahr || data.publikationsjahr) updateDashboardField(uuid, "Publikationsdatum", data.jahr || data.publikationsjahr);
    if (data.journal || data.journal_name) updateDashboardField(uuid, "Journal/Quelle", data.journal || data.journal_name);
    if (data.autoren || data.autoren_kurz) updateDashboardField(uuid, "Autoren", data.autoren || data.autoren_kurz);
    if (data.artikeltyp || data.studientyp) updateDashboardField(uuid, "Artikeltyp/Studientyp", data.artikeltyp || data.studientyp);
    if (data.volume) updateDashboardField(uuid, "Volume", data.volume);
    if (data.issue) updateDashboardField(uuid, "Issue", data.issue);
    if (data.pages) updateDashboardField(uuid, "Pages", data.pages);
  }
  
  else if (phaseId === "MASTER_ANALYSE") {
    if (data.produkt_fokus) updateDashboardField(uuid, "Produkt-Fokus", data.produkt_fokus);
    if (data.haupterkenntnis) updateDashboardField(uuid, "Haupterkenntnis", data.haupterkenntnis);
    
    if (data.kernaussagen) {
      const ka = Array.isArray(data.kernaussagen) ? data.kernaussagen.map(k => "• " + k).join("\n") : data.kernaussagen;
      updateDashboardField(uuid, "Kernaussagen", ka);
    }
    
    if (data.zusammenfassung) updateDashboardField(uuid, "Zusammenfassung", data.zusammenfassung);
    if (data.praktische_implikationen) updateDashboardField(uuid, "Praktische Implikationen", data.praktische_implikationen);
    if (data.kritische_bewertung) updateDashboardField(uuid, "Kritische Bewertung", data.kritische_bewertung);
    if (data.evidenzgrad) updateDashboardField(uuid, "Evidenzgrad", data.evidenzgrad);
    
    const pico = data.pico || {};
    if (pico.population) updateDashboardField(uuid, "PICO Population", pico.population);
    if (pico.intervention) updateDashboardField(uuid, "PICO Intervention", pico.intervention);
    if (pico.comparator) updateDashboardField(uuid, "PICO Comparator", pico.comparator);
    if (pico.outcomes) updateDashboardField(uuid, "PICO Outcomes", pico.outcomes);
  }
  
  else if (phaseId === "REDAKTION") {
    if (data.wichtige_abbildungen) {
      updateDashboardField(uuid, "Praktische Implikationen", data.wichtige_abbildungen);
    }
  }
  
  else if (phaseId === "FAKTENCHECK") {
    if (data.is_supported === false && data.reasoning) {
      const currentKritik = getDashboardDataByUUID(uuid)["Kritische Bewertung"] || "";
      const warning = `[⚠️ FAKTEN-WARNUNG: ${data.reasoning}]\n\n${currentKritik}`;
      updateDashboardField(uuid, "Kritische Bewertung", warning);
    }
  }
  
  else if (phaseId === "REVIEW") {
    if (data.ist_korrekt === false && data.korrigiertes_json) {
      applyJSONToDashboardPhaseAware(uuid, "MASTER_ANALYSE", data.korrigiertes_json);
    }
    updateDashboardField(uuid, "Review-Status", "GEPRÜFT");
  }
  
  updateDashboardField(uuid, "Letzte Änderung", new Date());
}

function parseJsonFromLlm(text) {
  if (!text) throw new Error("Leere Antwort");
  
  try {
    let cleanText = text.trim();
    
    if (cleanText.includes("I'm still learning") || cleanText.includes("can't help")) {
      throw new Error("Gemini hat Anfrage abgelehnt");
    }
    
    cleanText = cleanText.replace(/```json/gi, "").replace(/```/g, "").trim();
    
    const firstBrace = cleanText.indexOf('{');
    const lastBrace = cleanText.lastIndexOf('}');
    
    if (firstBrace === -1 || lastBrace === -1) {
      throw new Error("Kein JSON-Block gefunden");
    }
    
    cleanText = cleanText.substring(firstBrace, lastBrace + 1);
    
    try {
      const directParse = JSON.parse(cleanText);
      Logger.log(`✅ JSON direkt geparst (Strategie 1)`);
      return directParse;
    } catch (e1) {
      Logger.log(`Strategie 1 fehlgeschlagen: ${e1.message}`);
    }
    
    try {
      let converted = cleanText
        .replace(/True/g, 'true')
        .replace(/False/g, 'false')
        .replace(/None/g, 'null');
      
      const parsed2 = JSON.parse(converted);
      Logger.log(`✅ JSON geparst (Strategie 2: Python-Bool)`);
      return parsed2;
    } catch (e2) {
      Logger.log(`Strategie 2 fehlgeschlagen: ${e2.message}`);
    }
    
    try {
      let converted = cleanText
        .replace(/True/g, 'true')
        .replace(/False/g, 'false')
        .replace(/None/g, 'null');
      
      converted = converted.replace(/'([^']+?)'(\s*:)/g, '"$1"$2');
      
      const parsed3 = JSON.parse(converted);
      Logger.log(`✅ JSON geparst (Strategie 3: Property-Namen)`);
      return parsed3;
    } catch (e3) {
      Logger.log(`Strategie 3 fehlgeschlagen: ${e3.message}`);
    }
    
    try {
      let aggressive = cleanText
        .replace(/True/g, 'true')
        .replace(/False/g, 'false')
        .replace(/None/g, 'null')
        .replace(/'/g, '"');
      
      const parsed4 = JSON.parse(aggressive);
      Logger.log(`✅ JSON geparst (Strategie 4: Aggressive Konvertierung)`);
      return parsed4;
    } catch (e4) {
      Logger.log(`Strategie 4 fehlgeschlagen: ${e4.message}`);
    }
    
    try {
      Logger.log(`Versuche Strategie 5: Nuclear Quote Replacement`);
      
      let nuclear = cleanText
        .replace(/True/g, 'true')
        .replace(/False/g, 'false')
        .replace(/None/g, 'null');
      
      nuclear = nuclear.replace(/:\s*"((?:[^"\\]|\\.)*)"/g, function(match, content) {
        const fixed = content
          .replace(/\\"/g, '___ESCAPED_QUOTE___')
          .replace(/"/g, "'")
          .replace(/___ESCAPED_QUOTE___/g, '\\"');
        return `: "${fixed}"`;
      });
      
      const parsed5 = JSON.parse(nuclear);
      Logger.log(`✅ JSON geparst (Strategie 5: Nuclear)`);
      return parsed5;
    } catch (e5) {
      Logger.log(`Strategie 5 fehlgeschlagen: ${e5.message}`);
    }
    
    throw new Error("Konnte JSON mit keiner der 5 Strategien parsen");
    
  } catch (e) {
    Logger.log(`❌ FINALER FEHLER: ${e.message}`);
    Logger.log(`Originaler Text (erste 1000 Zeichen):\n${text.substring(0, 1000)}`);
    
    throw new Error(`Das JSON-Format der KI ist ungültig. Fehler: ${e.message}\n\nVersuche, die Antwort der KI erneut zu kopieren.`);
  }
}

function getCombinedVolltext(rowData) {
  const teil1 = rowData["Volltext/Extrakt"] || "";
  const teil2 = rowData["Volltext_Teil2"] || "";
  const teil3 = rowData["Volltext_Teil3"] || "";
  
  if (teil1 || teil2 || teil3) {
    return (teil1 + teil2 + teil3).trim();
  }
  
  return rowData["Inhalt/Abstract"] || "N/A";
}
