// FILE: onepager.gs

/**
 * OnePager Erstellung nach Vorlage
 * VERSION: 2.0 (mit korrigiertem Review-Status Check)
 */

function createOnePagerForSelected() {
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
  
  logAction("OnePager", `Erstelle OnePager für ${numRows} Publikationen`);
  
  let created = 0;
  let skipped = 0;
  const errors = [];
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const uuid = sheet.getRange(row, 1).getValue();
    
    if (!uuid) {
      skipped++;
      continue;
    }
    
    try {
      const result = createOnePagerForUUID(uuid, false);
      if (result.success) {
        created++;
      } else {
        skipped++;
        errors.push(`${result.title}: ${result.error}`);
      }
    } catch (e) {
      skipped++;
      errors.push(`UUID ${uuid}: ${e.message}`);
    }
  }
  
  let message = `OnePager erstellt: ${created}\nÜbersprungen: ${skipped}`;
  if (errors.length > 0) {
    message += `\n\nFehler:\n${errors.slice(0, 5).join("\n")}`;
    if (errors.length > 5) message += `\n... und ${errors.length - 5} weitere`;
  }
  
  SpreadsheetApp.getUi().alert(message);
  logAction("OnePager", message);
}

function recreateOnePagerForSelected() {
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
  
  let created = 0;
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const uuid = sheet.getRange(row, 1).getValue();
    
    if (uuid) {
      const result = createOnePagerForUUID(uuid, true);
      if (result.success) created++;
    }
  }
  
  SpreadsheetApp.getUi().alert(`${created} OnePager neu erstellt`);
  logAction("OnePager", `${created} OnePager neu erstellt (Force)`);
}

function batchCreateOnePagerByRelevance(relevance) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  if (sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("Keine Daten im Dashboard");
    return;
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, DASHBOARD_HEADERS.length).getValues();
  const relevanceCol = DASHBOARD_HEADERS.indexOf("Relevanz");
  const onepagerStatusCol = DASHBOARD_HEADERS.indexOf("OnePager Status");
  
  let created = 0;
  let skipped = 0;
  
  data.forEach(row => {
    const uuid = row[0];
    const rowRelevance = row[relevanceCol];
    const onepagerStatus = row[onepagerStatusCol];
    
    if (rowRelevance === relevance && onepagerStatus === "NICHT_ERSTELLT") {
      const result = createOnePagerForUUID(uuid, false);
      if (result.success) {
        created++;
      } else {
        skipped++;
      }
    }
  });
  
  SpreadsheetApp.getUi().alert(`OnePager Batch (${relevance}):\nErstellt: ${created}\nÜbersprungen: ${skipped}`);
  logAction("OnePager Batch", `Relevanz ${relevance}: ${created} erstellt, ${skipped} übersprungen`);
}

function createOnePagerForUUID(uuid, force) {
  try {
    const data = getDashboardDataByUUID(uuid);
    if (!data) {
      return { success: false, error: "UUID nicht gefunden", title: "" };
    }
    
    if (!force) {
      if (data["OnePager Status"] === "ERSTELLT") {
        return { success: false, error: "Bereits erstellt", title: data.Titel };
      }
      
      // FIX: Akzeptiere "FERTIG" und "GEPRÜFT"
      const reviewStatus = String(data["Review-Status"] || "").toUpperCase();
      if (reviewStatus !== "FERTIG" && reviewStatus !== "GEPRÜFT") {
        return { success: false, error: `Review-Status ist "${data["Review-Status"]}", erwartet: FERTIG`, title: data.Titel };
      }
      
      if (data.Status !== "Abgeschlossen") {
        return { success: false, error: "Status nicht Abgeschlossen", title: data.Titel };
      }
      
      // Pflichtfelder ohne Artikeltyp/Studientyp (da oft nicht ausgefüllt)
      const requiredFields = [
        "Titel", "Autoren", "Publikationsdatum", "Journal/Quelle", "Zusammenfassung",
        "Kernaussagen", "Haupterkenntnis"
      ];
      
      const validation = validateCompleteness(uuid, requiredFields);
      if (!validation.valid) {
        return { 
          success: false, 
          error: `Fehlende Pflichtfelder: ${validation.missing.join(", ")}`,
          title: data.Titel 
        };
      }
    }
    
    const templateId = getConfig("ONEPAGER_TEMPLATE_FILE_ID") || getConfig("Onepager Vorlagen ID");
    if (!templateId) {
      throw new Error("OnePager Template ID nicht konfiguriert");
    }
    
    const templateDoc = DriveApp.getFileById(templateId);
    const outputFolderId = getConfig("ONEPAGER_OUTPUT_FOLDER_ID") || getConfig("Onepager Hauptordner ID");
    const outputFolder = outputFolderId ? DriveApp.getFolderById(outputFolderId) : DriveApp.getRootFolder();
    
    const newFileName = `OnePager_${data.Titel.substring(0, 50).replace(/[^a-zA-Z0-9]/g, '_')}_${uuid}`;
    const newDoc = templateDoc.makeCopy(newFileName, outputFolder);
    const doc = DocumentApp.openById(newDoc.getId());
    const body = doc.getBody();
    
    const tables = body.getTables();
    if (tables.length === 0) {
      throw new Error("Keine Tabelle in der Vorlage gefunden");
    }
    
    const table = tables[0];
    
    fillOnePagerTable(table, data);
    
    doc.saveAndClose();
    
    updateDashboardField(uuid, "OnePager Link", newDoc.getUrl());
    updateDashboardField(uuid, "OnePager Status", "ERSTELLT");
    updateDashboardField(uuid, "OnePager erstellt am", new Date());
    
    logOnePagerHistory(uuid, data.Titel, newDoc.getUrl(), "ERSTELLT", "");
    
    logAction("OnePager", `OnePager erstellt: ${data.Titel}`);
    
    return { success: true, title: data.Titel };
    
  } catch (e) {
    logError("createOnePagerForUUID", e);
    updateDashboardField(uuid, "OnePager Status", "FEHLER");
    updateDashboardField(uuid, "Fehler-Details", e.message);
    logOnePagerHistory(uuid, data?.Titel || "", "", "FEHLER", e.message);
    return { success: false, error: e.message, title: data?.Titel || "" };
  }
}

function fillOnePagerTable(table, data) {
  const numRows = table.getNumRows();
  
  for (let i = 0; i < numRows; i++) {
    const row = table.getRow(i);
    if (row.getNumCells() < 2) continue;
    
    const labelCell = row.getCell(0);
    const contentCell = row.getCell(1);
    const labelText = labelCell.getText().trim();
    
    if (labelText.includes("Studieninfo")) {
      const studieninfo = buildStudieninfo(data);
      contentCell.clear();
      contentCell.setText(studieninfo);
    } else if (labelText.includes("Key Learning")) {
      contentCell.clear();
      contentCell.setText(data.Haupterkenntnis || "N/A");
    } else if (labelText.includes("Studientyp") || labelText.includes("Patientenanzahl")) {
      const studientypInfo = buildStudientypInfo(data);
      contentCell.clear();
      contentCell.setText(studientypInfo);
    } else if (labelText.includes("Besonderheit") || labelText.includes("Ziel")) {
      const besonderheit = buildBesonderheit(data);
      contentCell.clear();
      contentCell.setText(besonderheit);
    } else if (labelText.includes("Ergebnisse")) {
      const ergebnisse = buildErgebnisse(data);
      contentCell.clear();
      contentCell.setText(ergebnisse);
    } else if (labelText.includes("Links")) {
      const links = buildLinks(data);
      contentCell.clear();
      contentCell.setText(links);
    }
  }
}

function buildStudieninfo(data) {
  let text = "";
  text += `Titel: ${data.Titel || "N/A"}\n`;
  text += `Autoren: ${data.Autoren || "N/A"}\n`;
  text += `Zeitschrift & Jahr: ${data["Journal/Quelle"] || "N/A"} ${data.Publikationsdatum || ""}\n`;
  text += `Sponsor: N/A`;
  return text;
}

function buildStudientypInfo(data) {
  let text = "";
  text += `Studientyp: ${data["Artikeltyp/Studientyp"] || "N/A"}\n`;
  text += `Patientenanzahl: ${data["PICO Population"] || "N/A"}`;
  return text;
}

function buildBesonderheit(data) {
  let text = "";
  if (data["Praktische Implikationen"]) {
    text += data["Praktische Implikationen"];
  } else if (data.Zusammenfassung) {
    text += data.Zusammenfassung.substring(0, 200);
  } else {
    text += "N/A";
  }
  return text;
}

function buildErgebnisse(data) {
  let text = "";
  if (data.Kernaussagen) {
    text += data.Kernaussagen;
  } else if (data["PICO Outcomes"]) {
    text += data["PICO Outcomes"];
  } else {
    text += "N/A";
  }
  return text;
}

function buildLinks(data) {
  if (data.DOI) {
    return `DOI: ${data.DOI}`;
  } else if (data.PMID) {
    return `PMID: ${data.PMID}`;
  } else {
    return "N/A";
  }
}

function logOnePagerHistory(uuid, title, docLink, action, error) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ONEPAGER_HISTORY_SHEET_NAME);
    if (!sheet) return;
    
    sheet.appendRow([
      uuid,
      title,
      new Date(),
      docLink,
      action,
      error
    ]);
  } catch (e) {
    logError("logOnePagerHistory", e);
  }
}
