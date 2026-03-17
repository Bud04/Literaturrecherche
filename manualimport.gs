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

function importPublicationToDashboard(item, source) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName(DASHBOARD_SHEET_NAME);
    if (!dashboard) return false;

    const headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];

    // Duplikat-Check via DOI oder PMID
    if (item.doi || item.pmid) {
      const lastRow = dashboard.getLastRow();
      if (lastRow > 1) {
        const doiCol = headers.indexOf('DOI');
        const pmidCol = headers.indexOf('PMID');
        const existingData = dashboard.getRange(2, 1, lastRow - 1, headers.length).getValues();
        for (const row of existingData) {
          if (item.doi && doiCol >= 0 && row[doiCol] === item.doi) return false;
          if (item.pmid && pmidCol >= 0 && row[pmidCol] === item.pmid) return false;
        }
      }
    }

    const uuid = Utilities.getUuid();
    const now = new Date();

    const dataMap = {
      'UUID':              uuid,
      'PMID':              item.pmid || '',
      'DOI':               item.doi || '',
      'Titel':             item.title || '',
      'Autoren':           item.authors || '',
      'Publikationsdatum': item.year || '',
      'Journal/Quelle':    item.journal || '',
      'Link':              item.link || '',
      'Inhalt/Abstract':   item.abstract || '',
      'Quelle':            source || 'Manuell',
      'Volltext-Status':   'OFFEN',
      'Status':            'Neu importiert | Manuell',
      'Import-Timestamp':  now,
      'Letzte Änderung':   now,
      'Flow_Trigger_Researcher': 'PENDING'
    };

    const row = headers.map(h => dataMap[String(h).trim()] ?? '');
    const nextRow = dashboard.getLastRow() + 1;
    dashboard.getRange(nextRow, 1, 1, row.length).setValues([row]);

    return true;

  } catch (e) {
    logError('importPublicationToDashboard', e);
    return false;
  }
}
