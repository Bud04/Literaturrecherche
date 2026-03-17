// FILE: exportCitavi.gs

/**
 * Exportiert Datensätze im RIS-Format für Citavi
 */

function exportRISForSelected() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  const selection = sheet.getActiveRange();
  
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
  
  logAction("Export", `Starte Export für ${numRows} Zeilen`);
  
  const uuidsToExport = [];
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const uuid = sheet.getRange(row, 1).getValue();
    if (uuid) {
      uuidsToExport.push(uuid);
    }
  }
  
  exportUUIDsToRIS(uuidsToExport, false);
}

function forceExportRISForSelected() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  const selection = sheet.getActiveRange();
  
  if (!selection) return;
  
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  const uuidsToExport = [];
  
  for (let i = 0; i < numRows; i++) {
    const uuid = sheet.getRange(startRow + i, 1).getValue();
    if (uuid) uuidsToExport.push(uuid);
  }
  
  exportUUIDsToRIS(uuidsToExport, true);
}

function batchExportRISByRelevance(relevance) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  if (sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("Keine Daten im Dashboard");
    return;
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, DASHBOARD_HEADERS.length).getValues();
  const relevanceCol = DASHBOARD_HEADERS.indexOf("Relevanz");
  const exportStatusCol = DASHBOARD_HEADERS.indexOf("Export-Status");
  const reviewStatusCol = DASHBOARD_HEADERS.indexOf("Review-Status");
  
  const uuidsToExport = [];
  
  data.forEach(row => {
    const uuid = row[0];
    const rowRelevance = row[relevanceCol];
    const exportStatus = row[exportStatusCol];
    const reviewStatus = row[reviewStatusCol];
    
    // Exportiere nur, wenn Relevanz stimmt, Review OK ist und noch nicht exportiert wurde
    if (rowRelevance === relevance && 
        exportStatus === "NICHT_EXPORTIERT" && 
        reviewStatus === "GEPRÜFT") {
      uuidsToExport.push(uuid);
    }
  });
  
  if (uuidsToExport.length === 0) {
    SpreadsheetApp.getUi().alert(`Keine exportierbaren Datensätze für Relevanz '${relevance}' gefunden.`);
    return;
  }
  
  exportUUIDsToRIS(uuidsToExport, false);
}

function exportUUIDsToRIS(uuids, force) {
  try {
    let risContent = "";
    let count = 0;
    const errors = [];
    const exportedTitles = [];
    
    const outputFolderId = getConfig("Citavi Export Ordner ID");
    if (!outputFolderId) {
      SpreadsheetApp.getUi().alert("Citavi Export Ordner ID fehlt in Konfiguration");
      return;
    }
    
    uuids.forEach(uuid => {
      const data = getDashboardDataByUUID(uuid);
      if (!data) return;
      
      // Prüfungen
      if (!force) {
        if (data["Export-Status"] === "EXPORTIERT") return;
        
        const reviewStatus = String(data["Review-Status"] || "").toUpperCase();
        if (reviewStatus !== "FERTIG" && reviewStatus !== "GEPRÜFT") {
          errors.push(`${data.Titel}: Review noch nicht abgeschlossen`);
          return;
        }
      }
      
      // RIS Mapping bauen
      risContent += createRISEntry(data);
      risContent += "\r\n";
      
      // Status Update im Dashboard
      updateDashboardField(uuid, "Export-Status", "EXPORTIERT");
      updateDashboardField(uuid, "Exportiert am", new Date());
      
      exportedTitles.push(data.Titel);
      count++;
    });
    
    if (count === 0) {
      const msg = errors.length > 0 ? "Fehler:\n" + errors.join("\n") : "Keine Datensätze exportiert.";
      SpreadsheetApp.getUi().alert(msg);
      return;
    }
    
    // Dateiname basierend auf Titel(n)
    let fileName;
    if (count === 1) {
      // Einzelner Export: Titel als Dateiname
      const cleanTitle = exportedTitles[0]
        .substring(0, 80)
        .replace(/[^a-zA-Z0-9äöüÄÖÜß\s\-]/g, '')
        .replace(/\s+/g, '_')
        .trim();
      fileName = `${cleanTitle}.ris`;
    } else {
      // Batch-Export: Anzahl + Datum
      fileName = `Citavi_Export_${count}_Publikationen_${new Date().toISOString().slice(0,10)}.ris`;
    }
    
    // Datei speichern
    const folder = DriveApp.getFolderById(outputFolderId);
    const file = folder.createFile(fileName, risContent, MimeType.PLAIN_TEXT);
    
    // Historie schreiben
    uuids.forEach(uuid => {
      const data = getDashboardDataByUUID(uuid);
      if (data && (force || data["Export-Status"] === "EXPORTIERT")) {
        logExportHistory(uuid, data.Titel, "RIS", file.getUrl());
      }
    });
    
    logAction("Export", `${count} Datensätze in ${fileName} exportiert`);
    
    // Link anzeigen
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      "✅ Export erfolgreich!",
      `${count} Datensätze exportiert\n\n` +
      `📄 Datei: ${fileName}\n\n` +
      `📁 Link: ${file.getUrl()}`,
      ui.ButtonSet.OK
    );
    
  } catch (e) {
    logError("exportUUIDsToRIS", e);
    SpreadsheetApp.getUi().alert("Export Fehler: " + e.message);
  }
}

function createRISEntry(data) {
  const ris = [];
  ris.push("TY  - JOUR"); // Standard: Journal Article
  
  // Titel
  if (data.Titel) ris.push(`TI  - ${data.Titel}`);
  
  // Autoren
  if (data.Autoren) {
    // Versuche Autoren zu splitten (z.B. "Smith, J.; Doe, A." oder "Smith J, Doe A")
    let authors = data.Autoren.split(/;|,(?=\s*[A-Z])/); 
    authors.forEach(a => {
      ris.push(`AU  - ${a.trim()}`);
    });
  }
  
  // Jahr
  if (data.Jahr) ris.push(`PY  - ${data.Jahr}`);
  
  // Journal
  if (data["Journal/Quelle"]) {
    ris.push(`JO  - ${data["Journal/Quelle"]}`);
    ris.push(`T2  - ${data["Journal/Quelle"]}`);
  }
  
  // DOI / PMID
  if (data.DOI) ris.push(`DO  - ${data.DOI}`);
  if (data.PMID) {
    ris.push(`ID  - ${data.PMID}`); // ID oft als interne Nummer
    ris.push(`AN  - ${data.PMID}`); // Accession Number
  }
  
  // Abstract
  if (data.Zusammenfassung) {
    ris.push(`AB  - ${data.Zusammenfassung}`);
  } else if (data["Inhalt/Abstract"]) {
    ris.push(`AB  - ${data["Inhalt/Abstract"]}`);
  }
  
  // Link
  if (data.Link) ris.push(`UR  - ${data.Link}`);
  if (data["Link zum Volltext"]) ris.push(`L1  - ${data["Link zum Volltext"]}`); // PDF Link in L1
  
  // Keywords
  if (data.Schlagwörter) {
    const kws = data.Schlagwörter.split(/,|;/);
    kws.forEach(kw => ris.push(`KW  - ${kw.trim()}`));
  }
  
  // Custom Fields (Notes) für Bewertung
  let notes = "";
  if (data.Relevanz) notes += `Relevanz: ${data.Relevanz}; `;
  if (data.Haupterkenntnis) notes += `Learning: ${data.Haupterkenntnis}; `;
  if (data.Bewertung) notes += `Bewertung: ${data.Bewertung}; `;
  
  if (notes) ris.push(`N1  - ${notes}`);
  
  ris.push("ER  -");
  return ris.join("\r\n");
}

function logExportHistory(uuid, title, type, link) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(EXPORT_HISTORY_SHEET_NAME);
    if (!sheet) return;
    
    // Fingerprint aus Dashboard holen
    const data = getDashboardDataByUUID(uuid);
    const fp = data ? data.Fingerprint : "";
    
    sheet.appendRow([
      uuid,
      title,
      new Date(),
      type,
      fp,
      link
    ]);
  } catch (e) {
    logError("logExportHistory", e);
  }
}
