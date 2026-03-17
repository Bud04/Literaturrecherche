// FILE: review.gs

/**
 * JSON Validierung und Review
 */

function processJSONResponse() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(MANUAL_PROMPTS_SHEET_NAME);
  const selection = sheet.getActiveRange();
  
  if (!selection) {
    SpreadsheetApp.getUi().alert("Bitte eine Zeile mit JSON-Antwort auswählen");
    return;
  }
  
  const row = selection.getRow();
  if (row === 1) {
    SpreadsheetApp.getUi().alert("Bitte Datenzeile auswählen");
    return;
  }
  
  const data = sheet.getRange(row, 1, 1, 9).getValues()[0];
  const [uuid, title, phase, prompt1, prompt2, prompt3, driveLink, answer, timestamp] = data;
  
  if (!answer) {
    SpreadsheetApp.getUi().alert("Keine JSON-Antwort in dieser Zeile");
    return;
  }
  
  try {
    const jsonData = parseAndValidateJSON(answer, phase);
    
    if (!jsonData) {
      throw new Error("JSON-Validierung fehlgeschlagen");
    }
    
    applyJSONToDashboard(uuid, phase, jsonData);
    
    updateDashboardField(uuid, "Review-Status", "GEPRÜFT");
    advanceWorkflowStatus(uuid, phase);
    
    SpreadsheetApp.getUi().alert("JSON erfolgreich verarbeitet und ins Dashboard übertragen");
    logAction("Review", `JSON verarbeitet für ${title} (${phase})`);
    
  } catch (e) {
    logError("processJSONResponse", e);
    updateDashboardField(uuid, "Review-Status", "FEHLER");
    updateDashboardField(uuid, "Fehler-Details", e.message);
    logToErrorList(uuid, title, "JSON-Verarbeitung", e.message, "JSON prüfen und korrigieren", "");
    SpreadsheetApp.getUi().alert("Fehler: " + e.message);
  }
}

function parseAndValidateJSON(jsonString, phase) {
  let cleaned = jsonString.trim();
  
  if (cleaned.startsWith("```json")) {
    cleaned = cleaned.replace(/```json\s*/g, '').replace(/```\s*$/g, '');
  } else if (cleaned.startsWith("```")) {
    cleaned = cleaned.replace(/```\s*/g, '');
  }
  
  cleaned = cleaned.trim();
  
  let jsonData;
  try {
    jsonData = JSON.parse(cleaned);
  } catch (e) {
    throw new Error("Ungültiges JSON-Format: " + e.message);
  }
  
  if (!jsonData.schema_version) {
    throw new Error("Fehlendes Feld: schema_version");
  }
  
  if (!jsonData.phase) {
    throw new Error("Fehlendes Feld: phase");
  }
  
  if (jsonData.phase !== phase) {
    throw new Error(`Phase Mismatch: erwartet ${phase}, gefunden ${jsonData.phase}`);
  }
  
  if (!jsonData.data) {
    throw new Error("Fehlendes Feld: data");
  }
  
  validatePhaseData(jsonData.data, phase);
  
  return jsonData;
}

function validatePhaseData(data, phase) {
  const catData = loadCategoriesAndKeywords();
  
  if (phase === "analysis" || phase === "triage") {
    if (data.relevanz) {
      const allowedRelevanz = ["Sehr hoch", "Hoch", "Mittel", "Niedrig", "N/A"];
      validateEnumValue(data.relevanz, allowedRelevanz, "Relevanz");
    }
    
    if (data.hauptkategorie && data.hauptkategorie !== "N/A") {
      if (!catData.categories.includes(data.hauptkategorie)) {
        throw new Error(`Ungültige Hauptkategorie: ${data.hauptkategorie}`);
      }
    }
    
    if (data.unterkategorien && Array.isArray(data.unterkategorien)) {
      data.unterkategorien.forEach(subcat => {
        let valid = false;
        Object.values(catData.subcategories).forEach(subcats => {
          if (subcats.includes(subcat)) valid = true;
        });
        if (!valid && subcat !== "N/A") {
          throw new Error(`Ungültige Unterkategorie: ${subcat}`);
        }
      });
    }
    
    if (data.schlagwoerter && Array.isArray(data.schlagwoerter)) {
      data.schlagwoerter.forEach(kw => {
        if (!catData.keywords.includes(kw) && kw !== "N/A") {
          throw new Error(`Ungültiges Schlagwort: ${kw}`);
        }
      });
    }
  }
  
  if (data.autoren) {
    validateAuthors(data.autoren);
  }
  
  return true;
}

function applyJSONToDashboard(uuid, phase, jsonData) {
  const data = jsonData.data;
  
  const fieldMapping = {
    pmid: "PMID",
    doi: "DOI",
    titel: "Titel",
    autoren: "Autoren",
    jahr: "Jahr",
    journal: "Journal/Quelle",
    volume: "Volume",
    issue: "Issue",
    pages: "Pages",
    artikeltyp: "Artikeltyp/Studientyp",
    hauptkategorie: "Hauptkategorie",
    relevanz: "Relevanz",
    relevanz_begruendung: "Relevanz-Begründung",
    produkt_fokus: "Produkt-Fokus",
    haupterkenntnis: "Haupterkenntnis",
    kernaussagen: "Kernaussagen",
    zusammenfassung: "Zusammenfassung",
    praktische_implikationen: "Praktische Implikationen",
    kritische_bewertung: "Kritische Bewertung",
    evidenzgrad: "Evidenzgrad",
    pico_population: "PICO Population",
    pico_intervention: "PICO Intervention",
    pico_comparator: "PICO Comparator",
    pico_outcomes: "PICO Outcomes"
  };
  
  Object.keys(fieldMapping).forEach(jsonKey => {
    if (data[jsonKey] !== undefined && data[jsonKey] !== null) {
      const dashboardField = fieldMapping[jsonKey];
      let value = data[jsonKey];
      
      if (Array.isArray(value)) {
        value = value.join(", ");
      }
      
      updateDashboardField(uuid, dashboardField, value);
    }
  });
  
  if (data.unterkategorien && Array.isArray(data.unterkategorien)) {
    updateDashboardField(uuid, "Unterkategorien", data.unterkategorien.join(", "));
  }
  
  if (data.schlagwoerter && Array.isArray(data.schlagwoerter)) {
    updateDashboardField(uuid, "Schlagwörter", data.schlagwoerter.join(", "));
  }
  
  checkForRelevantChanges(uuid);
}

function advanceWorkflowStatus(uuid, phase) {
  const currentStatus = getDashboardDataByUUID(uuid)?.Status || "";
  
  const statusFlow = {
    "triage": {
      from: "Neu",
      to: "Analyse ausstehend"
    },
    "analysis": {
      from: "Analyse ausstehend",
      to: "Analyse fertig"
    },
    "metadata_extract": {
      from: "Analyse fertig",
      to: "Faktencheck ausstehend"
    },
    "factcheck_batch": {
      from: "Faktencheck ausstehend",
      to: "Faktencheck fertig"
    },
    "analysis_review": {
      from: "Faktencheck fertig",
      to: "Bewertung ausstehend"
    }
  };
  
  if (statusFlow[phase] && currentStatus === statusFlow[phase].from) {
    updateDashboardField(uuid, "Status", statusFlow[phase].to);
  }
}

function checkForRelevantChanges(uuid) {
  const data = getDashboardDataByUUID(uuid);
  if (!data) return;
  
  const onePagerStatus = data["OnePager Status"];
  const exportStatus = data["Export-Status"];
  
  if (onePagerStatus === "ERSTELLT") {
    updateDashboardField(uuid, "OnePager Status", "GEÄNDERT_SEIT_ONEPAGER");
  }
  
  if (exportStatus === "EXPORTIERT") {
    updateDashboardField(uuid, "Export-Status", "GEÄNDERT_SEIT_EXPORT");
  }
}

function validateCompleteness(uuid, requiredFields) {
  const data = getDashboardDataByUUID(uuid);
  if (!data) return { valid: false, missing: ["UUID nicht gefunden"] };
  
  const missing = [];
  
  requiredFields.forEach(field => {
    const value = data[field];
    if (!value || value === "" || value === "N/A") {
      missing.push(field);
    }
  });
  
  return {
    valid: missing.length === 0,
    missing: missing
  };
}
