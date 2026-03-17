// FILE: utils.gs

/**
 * Utility-Funktionen
 */

function normalizeTitle(title) {
  if (!title) return "";
  return title.toLowerCase()
    .replace(/[*\[\]]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function createFingerprint(pmid, doi, title, year) {
  if (pmid) return "PMID:" + pmid;
  if (doi) return "DOI:" + doi;
  const normalized = normalizeTitle(title) + (year || "");
  return "TITLE:" + computeHash(normalized);
}

function findDashboardRowByUUID(uuid) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return -1;
  
  const uuidCol = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < uuidCol.length; i++) {
    if (uuidCol[i][0] === uuid) {
      return i + 2;
    }
  }
  return -1;
}

function findDashboardRowByFingerprint(fingerprint) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) return -1;
  
  const fpColIndex = DASHBOARD_HEADERS.indexOf("Fingerprint") + 1;
  const fpCol = sheet.getRange(2, fpColIndex, sheet.getLastRow() - 1, 1).getValues();
  for (let i = 0; i < fpCol.length; i++) {
    if (fpCol[i][0] === fingerprint) {
      return i + 2;
    }
  }
  return -1;
}

/**
 * Findet die erste leere Zeile im Dashboard (ab Zeile 2)
 * Prüft ob UUID-Spalte leer ist
 */
function getNextEmptyRowInDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  const lastRow = sheet.getLastRow();
  
  // Nur Header vorhanden oder Sheet leer
  if (lastRow <= 1) {
    return 2;
  }
  
  // Hole alle UUIDs (Spalte A, ab Zeile 2)
  const uuids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  
  // Finde erste leere Zeile
  for (let i = 0; i < uuids.length; i++) {
    if (!uuids[i][0] || uuids[i][0] === "") {
      return i + 2; // +2 weil Array bei 0 startet und Header bei Zeile 1
    }
  }
  
  // Keine leere Zeile gefunden → nach der letzten Zeile
  return lastRow + 1;
}

function getDashboardDataByUUID(uuid) {
  const row = findDashboardRowByUUID(uuid);
  if (row === -1) return null;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  const data = sheet.getRange(row, 1, 1, DASHBOARD_HEADERS.length).getValues()[0];
  
  const obj = {};
  DASHBOARD_HEADERS.forEach((header, index) => {
    obj[header] = data[index];
  });
  return obj;
}

function updateDashboardField(uuid, fieldName, value) {
  const row = findDashboardRowByUUID(uuid);
  if (row === -1) return false;
  
  const colIndex = getDashboardColumnIndex(fieldName);
  if (colIndex === 0) return false;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  sheet.getRange(row, colIndex).setValue(value);
  
  const lastChangeCol = getDashboardColumnIndex("Letzte Änderung");
  if (lastChangeCol > 0) {
    sheet.getRange(row, lastChangeCol).setValue(new Date());
  }
  
  return true;
}

// ROBUSTER CLEANER (WICHTIG!)
function cleanEmailBody(body) {
  if (!body) return "";
  
  let cleaned = body;
  
  // 1. Weiterleitungs-Header entfernen (verschiedene Formate)
  cleaned = cleaned.replace(/---------- Forwarded message ---------[\s\S]*?Subject:.*?\n/g, '');
  cleaned = cleaned.replace(/-----Original Message-----[\s\S]*?Subject:.*?\n/g, '');
  cleaned = cleaned.replace(/Von:.*?\nGesendet:.*?\nAn:.*?\nBetreff:.*?\n/g, '');
  
  // 2. Zitat-Zeichen (>) am Zeilenanfang entfernen
  cleaned = cleaned.replace(/^[\s>]+/gm, '');

  // 3. Bilder und Social Media entfernen
  cleaned = cleaned.replace(/\[image:[^\]]*\]/g, '');
  cleaned = cleaned.replace(/<img[^>]*>/gi, '');
  
  const socialPatterns = [
    /https?:\/\/(www\.)?(twitter|facebook|linkedin|fb)\.com\/[^\s]*/gi,
    /Share on (Twitter|Facebook|LinkedIn)/gi,
    /Save to (My Library|Reading List)/gi
  ];
  socialPatterns.forEach(pattern => {
    cleaned = cleaned.replace(pattern, '');
  });
  
  // 4. Mehrfache Leerzeilen normalisieren
  cleaned = cleaned.replace(/\n{3,}/g, '\n\n');
  cleaned = cleaned.trim();
  
  return cleaned;
}

function extractPMID(text) {
  if (!text) return null;
  const match = text.match(/PMID[:\s]*(\d+)/i);
  return match ? match[1] : null;
}

function extractDOI(text) {
  if (!text) return null;
  const match = text.match(/10\.\d{4,}\/[^\s]+/);
  return match ? match[0] : null;
}

function safeGetRange(sheet, row, col, numRows, numCols) {
  if (numRows <= 0 || numCols <= 0) return null;
  if (row <= 0 || col <= 0) return null;
  try {
    return sheet.getRange(row, col, numRows, numCols);
  } catch (e) {
    logError("safeGetRange", e);
    return null;
  }
}

function loadCategoriesAndKeywords() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KATEGORIEN_SHEET_NAME);
  if (!sheet || sheet.getLastRow() <= 1) {
    return { categories: [], subcategories: {}, keywords: [] };
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const categories = new Set();
  const subcategories = {};
  const keywords = new Set();
  
  data.forEach(row => {
    const [cat, subcat, kw] = row;
    if (cat) {
      categories.add(cat);
      if (!subcategories[cat]) subcategories[cat] = new Set();
      if (subcat) subcategories[cat].add(subcat);
    }
    if (kw) keywords.add(kw);
  });
  
  return {
    categories: Array.from(categories),
    subcategories: Object.fromEntries(
      Object.entries(subcategories).map(([k, v]) => [k, Array.from(v)])
    ),
    keywords: Array.from(keywords)
  };
}

function validateEnumValue(value, allowedValues, fieldName) {
  if (!value || value === "N/A") return true;
  if (allowedValues.includes(value)) return true;
  logToErrorList("", "", "Validation", `Ungültiger Wert für ${fieldName}: ${value}`, `Erlaubt: ${allowedValues.join(", ")}`, "");
  return false;
}

function validateAuthors(authors) {
  if (!authors) return true;
  return true;
}
