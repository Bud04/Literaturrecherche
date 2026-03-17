// FILE: manualImport.gs

/**
 * Manueller Import aus Sheet
 */

function syncManualImportToDashboard() {
  logAction("Manueller Import", "Starte Synchronisation");
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(MANUAL_IMPORT_SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      SpreadsheetApp.getUi().alert("Keine Daten im Manueller Import Sheet");
      return;
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    let importedCount = 0;
    
    data.forEach((row, index) => {
      const [titel, autoren, jahr, journal, doi, pmid, link, abstract, notizen, importiert] = row;
      
      if (importiert === true || importiert === "TRUE" || importiert === "Ja") {
        return;
      }
      
      if (!titel) return;
      
      const item = {
        title: titel,
        authors: autoren || "",
        year: jahr || "",
        journal: journal || "",
        doi: doi || "",
        pmid: pmid || "",
        link: link || "",
        abstract: abstract || ""
      };
      
      const imported = importPublicationToDashboard(item, "Manuell");
      
      if (imported) {
        sheet.getRange(index + 2, 10).setValue("Ja");
        importedCount++;
      }
    });
    
    logAction("Manueller Import", `${importedCount} Publikationen importiert`);
    SpreadsheetApp.getUi().alert(`${importedCount} Publikationen erfolgreich importiert`);
    
  } catch (e) {
    logError("syncManualImportToDashboard", e);
    SpreadsheetApp.getUi().alert("Fehler beim Import: " + e.message);
  }
}
