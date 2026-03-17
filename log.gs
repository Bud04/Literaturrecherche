// FILE: log.gs

/**
 * Logging-Funktionen
 */

function logAction(action, details, user) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(LOGBOOK_SHEET_NAME);
    if (!sheet) return;
    
    const timestamp = new Date();
    const userName = user || Session.getActiveUser().getEmail();
    sheet.appendRow([timestamp, action, details, userName]);
  } catch (e) {
    Logger.log("logAction failed: " + e.message);
  }
}

function logError(functionName, error) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ERROR_LOG_SHEET_NAME);
    if (!sheet) return;
    
    const timestamp = new Date();
    const message = error.message || String(error);
    const stack = error.stack || "";
    sheet.appendRow([timestamp, functionName, message, stack]);
    Logger.log(`ERROR in ${functionName}: ${message}`);
  } catch (e) {
    Logger.log("logError failed: " + e.message);
  }
}

function logToErrorList(uuid, title, errorType, details, recommendation, originalLink) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ERROR_LIST_SHEET_NAME);
    if (!sheet) return;
    
    const timestamp = new Date();
    sheet.appendRow([timestamp, uuid, title, errorType, details, recommendation, originalLink]);
  } catch (e) {
    Logger.log("logToErrorList failed: " + e.message);
  }
}

function logToImportReport(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(IMPORT_REPORT_SHEET_NAME);
    if (!sheet) return;
    
    const timestamp = new Date();
    sheet.appendRow([
      timestamp,
      data.source || "",
      data.gmailSubject || "",
      data.gmailMessageId || "",
      data.gmailThreadId || "",
      data.extractedTitle || "",
      data.extractedPrimaryLink || "",
      data.parseStatus || "",
      data.parseReason || "",
      data.volltextStatus || "",
      data.volltextReason || "",
      data.uuid || "",
      data.nextAction || ""
    ]);
  } catch (e) {
    Logger.log("logToImportReport failed: " + e.message);
  }
}
