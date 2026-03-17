// FILE: workflow.gs

// ==========================================
// 1. KONFIGURATION
// ==========================================
const _API_KEY = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY"); 
const _MODEL_NAME = "gpt-4o"; 

const _SHEET_DASHBOARD = "Dashboard";
const _SHEET_PROMPTS = "Konfiguration"; 
const _SHEET_CATEGORIES = "Kategorien"; 

// PROMPT MAPPING (Feste Zeilen im Konfiguration-Sheet)
const _PROMPT_MAP = {
  RESEARCHER: "B15",
  MASTER_ANALYSE: "B16",
  METADATEN: "B17",
  REDAKTION: "B18",
  TRIAGE: "B19",
  FAKTENCHECK: "B20",
  REVIEW: "B21"
};

// ✅ FIX #1: VOLLSTÄNDIGE 50 SPALTEN (inkl. Mail-Felder + Batch-Phase)
const _COLS_DASHBOARD = [
  "UUID", "PMID", "DOI", "Titel", "Autoren", "Jahr", "Journal/Quelle", "Volume", "Issue", "Pages", 
  "Artikeltyp/Studientyp", "Quelle", "Link", "Link zum Volltext", "Volltext-Status", 
  "Volltext-Datei-Link", "Inhalt/Abstract", "Volltext/Extrakt", "Hauptkategorie", 
  "Unterkategorien", "Schlagwörter", "Relevanz", "Relevanz-Begründung", "Produkt-Fokus", 
  "Haupterkenntnis", "Kernaussagen", "Zusammenfassung", "Praktische Implikationen", 
  "Kritische Bewertung", "Evidenzgrad", "PICO Population", "PICO Intervention", 
  "PICO Comparator", "PICO Outcomes", "Review-Status", "Status", "Batch-Phase",
  "Für Citavi-Export", "Export-Status", "Exportiert am", "OnePager Link", "OnePager Status", 
  "OnePager erstellt am", "Fehler-Details", "Import-Timestamp", "Letzte Änderung", 
  "Fingerprint", "Prompt-Version", "Mail-Betreff", "Mail-Absender", "Mail-Datum"
];

// ==========================================
// 2. HAUPTFUNKTION (DIE PIPELINE)
// ==========================================

function analyzeSelectedRows() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(_SHEET_DASHBOARD);
  const ui = SpreadsheetApp.getUi();
  
  if (!_API_KEY) {
    ui.alert("API Key fehlt in den Skripteigenschaften.");
    return;
  }

  const selection = sheet.getSelection();
  const ranges = selection.getActiveRangeList().getRanges();
  
  if (!ranges || ranges.length === 0) {
    ui.alert("Bitte wähle Zeilen im Dashboard aus.");
    return;
  }
  
  ranges.forEach(range => {
    const startRow = range.getRow();
    const numRows = range.getNumRows();
    
    for (let i = 0; i < numRows; i++) {
      const rowIndex = startRow + i;
      if (rowIndex < 2) continue; 
      
      try {
        processFullPipeline(sheet, rowIndex);
      } catch (e) {
        setCell(sheet, rowIndex, "Status", "Fehler: " + e.message);
      }
    }
  });
  ui.alert("Pipeline-Analyse abgeschlossen.");
}

function processFullPipeline(sheet, rowIndex) {
  const rowData = getRowData(sheet, rowIndex);
  const lists = getCategoriesAndTags();
  
  let sourceText = rowData["Manueller Volltext"] || "";
  let usedSource = "Manueller Volltext";

  if (sourceText.length < 100) {
    const pdfUrl = rowData["Link zum Volltext"];
    if (pdfUrl && pdfUrl.toLowerCase().includes(".pdf")) {
      setCell(sheet, rowIndex, "Status", "Lese PDF...");
      const pdfText = getFullTextFromPdf(pdfUrl);
      if (pdfText && pdfText.length > 500) {
        sourceText = pdfText;
        usedSource = "PDF Volltext";
        setCell(sheet, rowIndex, "Volltext-Status", "VOLLTEXT_GEFUNDEN");
      }
    }
  }

  if (sourceText.length < 100) {
    sourceText = rowData["Inhalt/Abstract"] || "";
    usedSource = "Abstract";
  }

  if (sourceText.length < 50) {
    throw new Error("Kein ausreichender Text für die Analyse gefunden.");
  }

  // SCHRITT 1: METADATEN (B17)
  setCell(sheet, rowIndex, "Status", "Schritt 1/4: Metadaten...");
  const promptMeta = getPrompt("METADATEN").replace("{PUBLIKATIONSTEXT}", sourceText);
  const metaJson = parseJsonFromLlm(callLlmApi(promptMeta));
  
  setCell(sheet, rowIndex, "Jahr", metaJson.publikationsjahr);
  setCell(sheet, rowIndex, "DOI", metaJson.doi_string);
  setCell(sheet, rowIndex, "Autoren", metaJson.autoren_kurz);
  setCell(sheet, rowIndex, "Journal/Quelle", metaJson.journal_name);
  setCell(sheet, rowIndex, "PMID", metaJson.pmid || "");

  // SCHRITT 2: MASTER-ANALYSE (B16)
  setCell(sheet, rowIndex, "Status", "Schritt 2/4: Master-Analyse...");
  const promptMaster = getPrompt("MASTER_ANALYSE")
    .replace("{PUBLIKATIONSTEXT}", sourceText)
    .replace("{KATEGORIEN_LISTE}", lists.categories)
    .replace("{SCHLAGWÖRTER_LISTE}", lists.tags);
  const analysisJson = parseJsonFromLlm(callLlmApi(promptMaster));

  // SCHRITT 3: FAKTENCHECK (B20)
  setCell(sheet, rowIndex, "Status", "Schritt 3/4: Faktencheck...");
  const promptCheck = getPrompt("FAKTENCHECK")
    .replace("{PUBLIKATIONSTEXT}", sourceText)
    .replace("{KERNAUSSAGE}", analysisJson.haupterkenntnis);
  const checkResult = parseJsonFromLlm(callLlmApi(promptCheck));

  // SCHRITT 4: FINAL REVIEW (B21)
  setCell(sheet, rowIndex, "Status", "Schritt 4/4: Finales Review...");
  const promptReview = getPrompt("REVIEW")
    .replace("{PUBLIKATIONSTEXT}", sourceText)
    .replace("{JSON_ANALYSE_VON_SCHRITT_1}", JSON.stringify(analysisJson));
  const finalReview = parseJsonFromLlm(callLlmApi(promptReview));

  const finalData = finalReview.ist_korrekt ? analysisJson : finalReview.korrigiertes_json;
  writeResultsToSheet(sheet, rowIndex, finalData);
  
  if (checkResult.is_supported === false) {
    const currentKritik = getRowData(sheet, rowIndex)["Kritische Bewertung"];
    setCell(sheet, rowIndex, "Kritische Bewertung", "[⚠️ FAKTEN-WARNUNG: " + checkResult.reasoning + "]\n\n" + currentKritik);
  }

  setCell(sheet, rowIndex, "Status", "Abgeschlossen");
  setCell(sheet, rowIndex, "Letzte Änderung", new Date());
}

// ==========================================
// 3. HELPER (API, PDF, PARSING)
// ==========================================

function callLlmApi(prompt) {
  const payload = {
    model: _MODEL_NAME,
    messages: [
      { role: "system", content: "Antworte AUSSCHLIESSLICH im JSON-Format." },
      { role: "user", content: prompt }
    ],
    temperature: 0.1, max_tokens: 3000
  };
  const options = {
    method: "post", contentType: "application/json",
    headers: { "Authorization": "Bearer " + _API_KEY },
    payload: JSON.stringify(payload), muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", options);
  if (response.getResponseCode() !== 200) throw new Error("OpenAI API Fehler: " + response.getContentText());
  return JSON.parse(response.getContentText()).choices[0].message.content;
}

function getFullTextFromPdf(pdfUrl) {
  try {
    const response = UrlFetchApp.fetch(pdfUrl);
    const blob = response.getBlob();
    const file = Drive.Files.insert({ title: "Temp_Extract", mimeType: "application/vnd.google-apps.document" }, blob, { convert: true });
    const doc = DocumentApp.openById(file.id);
    const text = doc.getBody().getText();
    Drive.Files.remove(file.id);
    return text;
  } catch (e) { return null; }
}

function getPrompt(key) {
  const cellAddress = _PROMPT_MAP[key];
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_SHEET_PROMPTS);
  const val = sheet.getRange(cellAddress).getValue();
  if (!val) throw new Error("Prompt " + key + " nicht gefunden.");
  return val;
}

// ==========================================
// 4. SHEET OPERATIONS
// ==========================================

function getRowData(sheet, rowIndex) {
  const data = {};
  const maxCols = sheet.getLastColumn();
  const rowValues = sheet.getRange(rowIndex, 1, 1, maxCols).getValues()[0];
  _COLS_DASHBOARD.forEach((header, index) => {
    data[header] = rowValues[index] || "";
  });
  return data;
}

// ✅ FIX #2: NUR NOCH EINE setCell() FUNKTION (Dopplung entfernt)
function setCell(sheet, rowIndex, headerName, value) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf(headerName) + 1;
  
  if (colIndex > 0) {
    sheet.getRange(rowIndex, colIndex).setValue(value);
  } else {
    console.warn("Spalte '" + headerName + "' wurde im Dashboard nicht gefunden!");
  }
}

/**
 * Schreibt alle Analyse-Ergebnisse (inkl. PICO) in das Dashboard.
 */
function writeResultsToSheet(sheet, rowIndex, data) {
  if (!data) return;

  // 1. Basis-Metadaten
  if (data.studientyp)      setCell(sheet, rowIndex, "Artikeltyp/Studientyp", data.studientyp);
  if (data.journal_name)    setCell(sheet, rowIndex, "Journal/Quelle", data.journal_name);
  if (data.publikationsjahr) setCell(sheet, rowIndex, "Jahr", data.publikationsjahr);
  if (data.doi_string)      setCell(sheet, rowIndex, "DOI", data.doi_string);
  if (data.pmid)            setCell(sheet, rowIndex, "PMID", data.pmid);

  // 2. Strategische Analyse
  if (data.relevanz)            setCell(sheet, rowIndex, "Relevanz", data.relevanz);
  if (data.relevanz_begruendung) setCell(sheet, rowIndex, "Relevanz-Begründung", data.relevanz_begruendung);
  if (data.produkt_fokus)       setCell(sheet, rowIndex, "Produkt-Fokus", data.produkt_fokus);
  if (data.haupterkenntnis)     setCell(sheet, rowIndex, "Haupterkenntnis", data.haupterkenntnis);
  if (data.zusammenfassung)     setCell(sheet, rowIndex, "Zusammenfassung", data.zusammenfassung);
  if (data.praktischeImplikationen) setCell(sheet, rowIndex, "Praktische Implikationen", data.praktischeImplikationen);
  if (data.kritische_bewertung)  setCell(sheet, rowIndex, "Kritische Bewertung", data.kritische_bewertung);

  // 3. PICO SCHEMA
  if (data.pico) {
    setCell(sheet, rowIndex, "PICO Population", data.pico.population || "N/A");
    setCell(sheet, rowIndex, "PICO Intervention", data.pico.intervention || "N/A");
    setCell(sheet, rowIndex, "PICO Comparator", data.pico.comparator || "N/A");
    setCell(sheet, rowIndex, "PICO Outcomes", data.pico.outcomes || "N/A");
  }

  // 4. Kategorien & Schlagwörter
  if (data.kategorie) {
    if (data.kategorie.hauptkategorie) setCell(sheet, rowIndex, "Hauptkategorie", data.kategorie.hauptkategorie);
    if (data.kategorie.unterkategorien) {
      const subs = Array.isArray(data.kategorie.unterkategorien) ? data.kategorie.unterkategorien.join(", ") : data.kategorie.unterkategorien;
      setCell(sheet, rowIndex, "Unterkategorien", subs);
    }
  }
  if (data.schlagwoerter) {
    const tagsStr = Array.isArray(data.schlagwoerter) ? data.schlagwoerter.join(", ") : data.schlagwoerter;
    setCell(sheet, rowIndex, "Schlagwörter", tagsStr);
  }

  // 5. Kernaussagen (Aufzählung)
  if (data.kernaussagen) {
    const kaStr = Array.isArray(data.kernaussagen) ? data.kernaussagen.map(k => "• " + k).join("\n") : data.kernaussagen;
    setCell(sheet, rowIndex, "Kernaussagen", kaStr);
  }

  // 6. Zeitstempel & Status
  setCell(sheet, rowIndex, "Letzte Änderung", new Date());
}

function getCategoriesAndTags() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(_SHEET_CATEGORIES);
  if (!sheet) return { categories: "", fullMapping: "", tags: "" };
  
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  let mapping = {};

  values.forEach(row => {
    let cat = row[0] ? row[0].toString().trim() : "";
    let sub = row[1] ? row[1].toString().trim() : "";
    let tag = row[2] ? row[2].toString().trim() : "";
    
    if (cat) {
      if (!mapping[cat]) mapping[cat] = { subs: new Set(), tags: new Set() };
      if (sub) mapping[cat].subs.add(sub);
      if (tag) mapping[cat].tags.add(tag);
    }
  });

  let mappingText = Object.keys(mapping).map(cat => {
    let s = Array.from(mapping[cat].subs).join(", ");
    let t = Array.from(mapping[cat].tags).join(", ");
    return `HAUPTKATEGORIE: "${cat}"\n` +
           `  -> Wählbare UNTERKATEGORIEN (B): ${s || "Keine"}\n` +
           `  -> Wählbare SCHLAGWÖRTER (C): ${t || "Keine"}`;
  }).join("\n\n");

  const allTags = new Set();
  Object.values(mapping).forEach(m => m.tags.forEach(t => allTags.add(t)));

  return { 
    categories: Object.keys(mapping).join(", "), 
    fullMapping: mappingText,
    tags: Array.from(allTags).join(", ")
  };
}

function checkDashboardColumns() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const expected = _COLS_DASHBOARD;
  
  Logger.log("--- START SPALTEN-CHECK ---");
  expected.forEach(name => {
    if (headers.indexOf(name) === -1) {
      Logger.log("❌ FEHLT ODER TIPPFEHLER: '" + name + "'");
    } else {
      Logger.log("✅ OK: '" + name + "'");
    }
  });
  Logger.log("--- CHECK BEENDET ---");
  
  SpreadsheetApp.getUi().alert("Check beendet. Schau ins Protokoll (unten im Editor), um Fehler zu sehen.");
}

/**
 * Wrapper für Rückwärtskompatibilität
 */
function writeResultsToSheetExtended(sheet, rowIndex, data) {
  writeResultsToSheet(sheet, rowIndex, data);
}
