// FILE: quality_checks.gs
/**
 * ==========================================
 * QUALITY CHECKS & NAVIGATION
 * ==========================================
 * Readiness Checks für OnePager und Citavi Export
 */

/**
 * ✅ OnePager Readiness Check
 */
function qualityCheckOnePager() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  if (sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("Keine Daten im Dashboard");
    return;
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, DASHBOARD_HEADERS.length).getValues();
  
  const requiredFields = [
    "Titel", "Autoren", "Publikationsdatum", "Journal/Quelle", "Zusammenfassung",
    "Kernaussagen", "Haupterkenntnis", "Artikeltyp/Studientyp"
  ];
  
  const reviewStatusCol = DASHBOARD_HEADERS.indexOf("Review-Status");
  const statusCol = DASHBOARD_HEADERS.indexOf("Status");
  const onepagerStatusCol = DASHBOARD_HEADERS.indexOf("OnePager Status");
  
  let readyCount = 0;
  let notReadyCount = 0;
  const issues = [];
  
  data.forEach((row, index) => {
    const uuid = row[0];
    const titel = row[DASHBOARD_HEADERS.indexOf("Titel")];
    const reviewStatus = row[reviewStatusCol];
    const status = row[statusCol];
    const onepagerStatus = row[onepagerStatusCol];
    
    // Überspringe bereits erstellte
    if (onepagerStatus === "ERSTELLT") {
      return;
    }
    
    const missingFields = [];
    
    // Prüfe Review-Status
    if (reviewStatus !== "GEPRÜFT" && reviewStatus !== "FERTIG") {
      missingFields.push("Review-Status (nicht GEPRÜFT)");
    }
    
    // Prüfe Status
    if (status !== "Abgeschlossen") {
      missingFields.push("Status (nicht Abgeschlossen)");
    }
    
    // Prüfe Pflichtfelder
    requiredFields.forEach(field => {
      const colIndex = DASHBOARD_HEADERS.indexOf(field);
      const value = row[colIndex];
      if (!value || value === "" || value === "N/A") {
        missingFields.push(field);
      }
    });
    
    if (missingFields.length === 0) {
      readyCount++;
    } else {
      notReadyCount++;
      if (issues.length < 10) {
        issues.push({
          row: index + 2,
          titel: titel,
          missing: missingFields
        });
      }
    }
  });
  
  let message = `=== ONEPAGER READINESS CHECK ===\n\n`;
  message += `✅ Bereit für OnePager: ${readyCount}\n`;
  message += `❌ Nicht bereit: ${notReadyCount}\n\n`;
  
  if (issues.length > 0) {
    message += `--- DETAILS (erste 10) ---\n`;
    issues.forEach(issue => {
      message += `\nZeile ${issue.row}: ${issue.titel}\n`;
      message += `  Fehlt: ${issue.missing.join(", ")}\n`;
    });
  }
  
  // Auch in Fehlerliste schreiben
  issues.forEach(issue => {
    logToErrorList(
      "",
      issue.titel,
      "OnePager Readiness",
      `Fehlende Felder: ${issue.missing.join(", ")}`,
      "Felder ausfüllen oder Analyse wiederholen",
      ""
    );
  });
  
  SpreadsheetApp.getUi().alert(message);
  logAction("Qualitätscheck", `OnePager Readiness: ${readyCount} bereit, ${notReadyCount} nicht bereit`);
}

/**
 * ✅ Citavi Export Readiness Check
 */
function qualityCheckCitavi() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  if (sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("Keine Daten im Dashboard");
    return;
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, DASHBOARD_HEADERS.length).getValues();
  
  const requiredFields = [
    "Titel", "Autoren", "Publikationsdatum", "Journal/Quelle", 
    "Schlagwörter", "Hauptkategorie", "Zusammenfassung", "Kernaussagen", "Relevanz"
  ];
  
  const doiCol = DASHBOARD_HEADERS.indexOf("DOI");
  const pmidCol = DASHBOARD_HEADERS.indexOf("PMID");
  const reviewStatusCol = DASHBOARD_HEADERS.indexOf("Review-Status");
  const exportStatusCol = DASHBOARD_HEADERS.indexOf("Export-Status");
  
  let readyCount = 0;
  let notReadyCount = 0;
  const issues = [];
  
  data.forEach((row, index) => {
    const uuid = row[0];
    const titel = row[DASHBOARD_HEADERS.indexOf("Titel")];
    const reviewStatus = row[reviewStatusCol];
    const exportStatus = row[exportStatusCol];
    const doi = row[doiCol];
    const pmid = row[pmidCol];
    
    // Überspringe bereits exportierte
    if (exportStatus === "EXPORTIERT") {
      return;
    }
    
    const missingFields = [];
    
    // Prüfe Review-Status
    if (reviewStatus !== "GEPRÜFT" && reviewStatus !== "FERTIG") {
      missingFields.push("Review-Status (nicht GEPRÜFT)");
    }
    
    // Prüfe DOI oder PMID
    if ((!doi || doi === "") && (!pmid || pmid === "")) {
      missingFields.push("DOI oder PMID");
    }
    
    // Prüfe Pflichtfelder
    requiredFields.forEach(field => {
      const colIndex = DASHBOARD_HEADERS.indexOf(field);
      const value = row[colIndex];
      if (!value || value === "" || value === "N/A") {
        missingFields.push(field);
      }
    });
    
    if (missingFields.length === 0) {
      readyCount++;
    } else {
      notReadyCount++;
      if (issues.length < 10) {
        issues.push({
          row: index + 2,
          titel: titel,
          missing: missingFields
        });
      }
    }
  });
  
  let message = `=== CITAVI EXPORT READINESS CHECK ===\n\n`;
  message += `✅ Bereit für Export: ${readyCount}\n`;
  message += `❌ Nicht bereit: ${notReadyCount}\n\n`;
  
  if (issues.length > 0) {
    message += `--- DETAILS (erste 10) ---\n`;
    issues.forEach(issue => {
      message += `\nZeile ${issue.row}: ${issue.titel}\n`;
      message += `  Fehlt: ${issue.missing.join(", ")}\n`;
    });
  }
  
  // Auch in Fehlerliste schreiben
  issues.forEach(issue => {
    logToErrorList(
      "",
      issue.titel,
      "Citavi Readiness",
      `Fehlende Felder: ${issue.missing.join(", ")}`,
      "Felder ausfüllen oder Metadaten ergänzen",
      ""
    );
  });
  
  SpreadsheetApp.getUi().alert(message);
  logAction("Qualitätscheck", `Citavi Readiness: ${readyCount} bereit, ${notReadyCount} nicht bereit`);
}

/**
 * ✅ Navigation: Nächste relevante Zeile
 */
function jumpToNextRelevant() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  const currentRow = sheet.getActiveCell().getRow();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const colRel = headers.indexOf("Relevanz");
  
  for (let i = currentRow; i < data.length; i++) {
    try {
      if (sheet.isRowHiddenByFilter(i + 1)) continue;
    } catch (e) {
      // Filter nicht aktiv
    }
    
    const val = String(data[i][colRel]).toLowerCase();
    if (val.includes("hoch") || val.includes("sehr hoch")) {
      sheet.getRange(i + 1, 1).activate();
      return { found: true, row: i + 1, title: data[i][headers.indexOf("Titel")] };
    }
  }
  return { found: false, message: "Keine weiteren Highlights gefunden." };
}
