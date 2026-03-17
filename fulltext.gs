// FILE: fulltext.gs

/**
 * Volltextsuche automatisch und manuell
 * Features: Unpaywall API Integration, Direkter PDF-Check, PMC Fallback
 */

function batchFulltextSearch(maxItems) {
  const resumeKey = "FULLTEXT_BATCH_RESUME";
  const props = PropertiesService.getScriptProperties();
  let startIndex = parseInt(props.getProperty(resumeKey)) || 0;
  
  logAction("Volltextsuche Batch", `Starte ab Index ${startIndex}, max ${maxItems || 'alle'}`);
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
    
    if (sheet.getLastRow() <= 1) {
      props.deleteProperty(resumeKey);
      return "Keine Daten im Dashboard";
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, DASHBOARD_HEADERS.length).getValues();
    const statusColIndex = DASHBOARD_HEADERS.indexOf("Volltext-Status");
    
    let processedCount = 0;
    let foundCount = 0;
    
    // Iteriere durch Zeilen
    for (let i = startIndex; i < data.length; i++) {
      if (maxItems && processedCount >= maxItems) {
        props.setProperty(resumeKey, String(i));
        logAction("Volltextsuche Batch", `Pausiert bei Index ${i}, verarbeitet: ${processedCount}`);
        return `Batch pausiert nach ${processedCount} Items. Fortsetzen möglich.`;
      }
      
      const row = data[i];
      const status = row[statusColIndex];
      
      // Überspringe bereits gefundene oder manuell bearbeitete, es sei denn Status ist OFFEN/NUR_ABSTRACT
      if (status !== "OFFEN" && status !== "" && status !== "NUR_ABSTRACT" && status !== "PAYWALL") {
        continue;
      }
      
      const uuid = row[0];
      const title = row[DASHBOARD_HEADERS.indexOf("Titel")];
      const primaryLink = row[DASHBOARD_HEADERS.indexOf("Link")];
      const doi = row[DASHBOARD_HEADERS.indexOf("DOI")];
      const pmid = row[DASHBOARD_HEADERS.indexOf("PMID")];
      
      logAction("Volltextsuche", `Suche für: ${title}`);
      
      const result = searchFulltext(uuid, title, primaryLink, doi, pmid);
      
      if (result.status === "VOLLTEXT_GEFUNDEN") {
        foundCount++;
      }
      
      processedCount++;
      // Kurze Pause, um Server nicht zu überlasten
      Utilities.sleep(500);
    }
    
    // Wenn fertig, Property löschen
    props.deleteProperty(resumeKey);
    logAction("Volltextsuche Batch", `Abgeschlossen: ${processedCount} verarbeitet, ${foundCount} Volltexte gefunden`);
    return `Volltextsuche abgeschlossen:\n${processedCount} verarbeitet\n${foundCount} Volltexte gefunden`;
    
  } catch (e) {
    logError("batchFulltextSearch", e);
    throw e;
  }
}

function resumeFulltextBatch() {
  return batchFulltextSearch(50);
}

function searchFulltext(uuid, title, primaryLink, doi, pmid) {
  let result = {
    status: "OFFEN",
    link: "",
    candidates: "",
    reason: ""
  };

  // 1. STRATEGIE: Unpaywall API (Der "Goldstandard" für Open Access)
  // Dies findet legale Versionen auf Uni-Servern, die per Crawler schwer zu finden sind.
  if (doi) {
    const unpaywall = checkUnpaywall(doi);
    if (unpaywall.found) {
      result.status = "VOLLTEXT_GEFUNDEN";
      result.link = unpaywall.url;
      result.reason = "Unpaywall (Open Access)";
      saveResult(uuid, result);
      return result;
    }
  }

  // 2. STRATEGIE: Kandidaten-Links direkt prüfen
  const candidates = [];
  
  if (primaryLink) candidates.push(primaryLink);
  if (doi) candidates.push(`https://doi.org/${doi}`);
  if (pmid) {
    candidates.push(`https://pubmed.ncbi.nlm.nih.gov/${pmid}/`);
    // Versuch PMC PDF direkt (falls im Import übersehen oder erst später verfügbar)
    candidates.push(`https://www.ncbi.nlm.nih.gov/pmc/articles/PMC${pmid}/pdf/`);
  }
  
  result.candidates = candidates.join("; ");
  
  for (const url of candidates) {
    try {
      const response = UrlFetchApp.fetch(url, { 
        followRedirects: true, 
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0 (Scientific Research Bot)' }
      });
      
      const code = response.getResponseCode();
      const contentType = response.getHeaders()['Content-Type'] || '';
      const finalUrl = code === 200 ? url : "";
      
      // Ist es direkt ein PDF?
      if (code === 200 && (contentType.includes('application/pdf') || url.endsWith('.pdf'))) {
        result.status = "VOLLTEXT_GEFUNDEN";
        result.link = finalUrl;
        result.reason = "PDF Direkt-Link";
        break;
      }
      
      // HTML Content scannen
      const content = response.getContentText();
      
      // Suche nach Meta-Tags für PDF
      // <meta name="citation_pdf_url" content="...">
      const metaMatch = content.match(/name="citation_pdf_url"\s+content="(.*?)"/i) || content.match(/content="(.*?)"\s+name="citation_pdf_url"/i);
      
      if (metaMatch) {
        result.status = "VOLLTEXT_GEFUNDEN";
        result.link = metaMatch[1];
        result.reason = "Meta-Tag PDF";
        break;
      }
      
      // Einfache Textsuche im Content
      if (content.includes('full text') || content.includes('PMC') || content.includes('free article')) {
        // Hinweis: Wir speichern hier den Link zur HTML-Seite, da wir das PDF nicht direkt extrahiert haben
        result.status = "VOLLTEXT_GEFUNDEN";
        result.link = finalUrl;
        result.reason = "Volltext-Seite (HTML)";
        break;
      }
      
      if (content.includes('paywall') || content.includes('purchase') || content.includes('subscribe')) {
        result.status = "PAYWALL";
        result.reason = "Paywall erkannt";
      }
      
    } catch (e) {
      logError("searchFulltext candidate", e);
    }
    
    Utilities.sleep(200);
  }
  
  // Abschluss-Bewertung
  if (result.status === "OFFEN") {
    // Wenn nichts gefunden wurde, setzen wir auf NUR_ABSTRACT
    // und tragen es in die manuelle Liste ein
    result.status = "NUR_ABSTRACT";
    result.reason = "Kein Volltext gefunden";
    
    addToManualFulltextSearch(uuid, title, primaryLink, result.candidates);
  }
  
  // Ergebnisse speichern
  saveResult(uuid, result);
  
  // Reporting
  logToImportReport({
    uuid: uuid,
    volltextStatus: result.status,
    volltextReason: result.reason + " | Candidates: " + result.candidates
  });
  
  return result;
}

/**
 * Prüft DOI gegen Unpaywall API
 */
function checkUnpaywall(doi) {
  try {
    // Unpaywall benötigt eine E-Mail Adresse im Request
    const email = Session.getActiveUser().getEmail(); 
    const url = `https://api.unpaywall.org/v2/${doi}?email=${email}`;
    
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    
    if (response.getResponseCode() !== 200) {
      return { found: false };
    }
    
    const data = JSON.parse(response.getContentText());
    
    // Prüfe auf beste Open Access Location
    if (data.is_oa && data.best_oa_location) {
      // Bevorzuge direkten PDF Link
      if (data.best_oa_location.url_for_pdf) {
        return { 
          found: true, 
          url: data.best_oa_location.url_for_pdf 
        };
      }
      // Fallback auf HTML Landing Page
      if (data.best_oa_location.url) {
        return { 
          found: true, 
          url: data.best_oa_location.url 
        };
      }
    }
    
  } catch (e) {
    logError("checkUnpaywall", e);
  }
  
  return { found: false };
}

function saveResult(uuid, result) {
  updateDashboardField(uuid, "Volltext-Status", result.status);
  updateDashboardField(uuid, "Link zum Volltext", result.link);
  
  logFulltextSearch(uuid, "", "", result);
}

function logFulltextSearch(uuid, title, primaryLink, result) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FULLTEXT_SEARCH_SHEET_NAME);
    if (!sheet) return;
    
    sheet.appendRow([
      uuid,
      title || "N/A", // Titel wird hier ggf. nicht übergeben, ist aber okay für Log
      primaryLink || "N/A",
      result.candidates || "",
      result.status,
      result.link,
      new Date()
    ]);
  } catch (e) {
    logError("logFulltextSearch", e);
  }
}

function addToManualFulltextSearch(uuid, title, originalLink, candidatesChecked) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(MANUAL_FULLTEXT_SHEET_NAME);
    if (!sheet) return;
    
    sheet.appendRow([
      uuid,
      title,
      originalLink,
      candidatesChecked,
      "", // PDF Link leer lassen für manuellen Eintrag
      "TODO",
      "",
      new Date()
    ]);
    
    logAction("Manuelle Volltextsuche", `Eingetragen: ${title}`);
  } catch (e) {
    logError("addToManualFulltextSearch", e);
  }
}

function syncManualFulltextToDashboard() {
  logAction("Manuelle Volltextsuche", "Starte Sync zu Dashboard");
  
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(MANUAL_FULLTEXT_SHEET_NAME);
    
    if (!sheet || sheet.getLastRow() <= 1) {
      SpreadsheetApp.getUi().alert("Keine Daten vorhanden");
      return;
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 8).getValues();
    let syncedCount = 0;
    
    data.forEach((row, index) => {
      const [uuid, titel, originalLink, primaryLink, pdfLinkUrl, status, notizen, timestamp] = row;
      
      // Nur synchronisieren wenn UUID da ist, ein Link eingetragen wurde und es noch nicht erledigt ist
      if (!uuid || !pdfLinkUrl || status === "ERLEDIGT") return;
      
      updateDashboardField(uuid, "Link zum Volltext", pdfLinkUrl);
      updateDashboardField(uuid, "Volltext-Status", "MANUELL");
      
      // Markiere als Erledigt im manuellen Sheet
      sheet.getRange(index + 2, 6).setValue("ERLEDIGT");
      syncedCount++;
    });
    
    logAction("Manuelle Volltextsuche", `${syncedCount} Einträge synchronisiert`);
    SpreadsheetApp.getUi().alert(`${syncedCount} Volltexte erfolgreich übertragen`);
    
  } catch (e) {
    logError("syncManualFulltextToDashboard", e);
    SpreadsheetApp.getUi().alert("Fehler: " + e.message);
  }
}
