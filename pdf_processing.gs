// FILE: pdf_processing.gs
/**
 * ==========================================
 * PDF VERARBEITUNG
 * ==========================================
 * Funktionen für PDF Upload und Konvertierung
 */

/**
 * ✅ PDF Upload Basic (ohne OCR)
 */
function processBoosterPdf(rowIndex, base64PdfData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
    const uuid = sheet.getRange(rowIndex, 1).getValue();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const titleCol = headers.indexOf("Titel") + 1;
    const title = sheet.getRange(rowIndex, titleCol).getValue() || "Unnamed";
    
    Logger.log("PDF processing for: " + title);
    
    // 1. PDF Blob erstellen
    const pdfBlob = Utilities.newBlob(
      Utilities.base64Decode(base64PdfData), 
      'application/pdf', 
      sanitizeFilename(title) + '.pdf'
    );
    
    // 2. Original PDF speichern
    const pdfFolder = DriveApp.getRootFolder();
    const savedPdf = pdfFolder.createFile(pdfBlob);
    const pdfUrl = savedPdf.getUrl();
    
    Logger.log("PDF saved: " + pdfUrl);
    
    // 3. PDF konvertieren mit Drive API v3
    const conversionResult = convertPdfToTextV3(pdfBlob, 'TEMP_' + uuid);
    
    if (!conversionResult.success) {
      return {
        error: "❌ PDF-Konvertierung fehlgeschlagen!\n\n" +
               "Fehler: " + conversionResult.error + "\n\n" +
               "💡 Versuche stattdessen:\n" +
               "1. 🧠 Smart Modus (mit OCR)\n" +
               "2. 📝 Text manuell kopieren\n\n" +
               "PDF gespeichert: " + pdfUrl
      };
    }
    
    const extractedText = conversionResult.text;
    Logger.log("Text extracted: " + extractedText.length + " chars");
    
    // 4. Validierung
    if (!extractedText || extractedText.length < 100) {
      return { 
        error: "PDF enthält zu wenig Text (< 100 Zeichen).\n\n" +
               "💡 Nutze: 🧠 Smart Modus (mit OCR)\n\n" +
               "PDF gespeichert: " + pdfUrl 
      };
    }
    
    // 5. Warnung hinzufügen
    const warning = "\n\n[⚠️ HINWEIS: Tabellen fehlen. Original-PDF: " + pdfUrl + "]";
    const fullText = extractedText + warning;
    
    // 6. Bei großen Texten: Google Doc erstellen
    let textToSave = fullText;
    let volltextDocLink = null;
    
    if (fullText.length > 45000) {
      Logger.log("Text too long - creating Google Doc");
      
      const volltextDoc = DocumentApp.create('Volltext_' + sanitizeFilename(title) + '_' + uuid);
      volltextDoc.getBody().setText(fullText);
      
      const volltextDocFile = DriveApp.getFileById(volltextDoc.getId());
      volltextDocLink = volltextDocFile.getUrl();
      
      textToSave = "[📄 Text zu lang - " + fullText.length + " Zeichen]\n\n" +
                   "📄 Volltext-Doc: " + volltextDocLink + "\n" +
                   "📎 Original-PDF: " + pdfUrl;
    }
    
    // 7. Dashboard speichern
    const volltextCol = headers.indexOf("Volltext/Extrakt") + 1;
    const statusCol = headers.indexOf("Volltext-Status") + 1;
    const linkCol = headers.indexOf("Volltext-Datei-Link") + 1;
    
    if (volltextCol > 0) sheet.getRange(rowIndex, volltextCol).setValue(textToSave);
    if (statusCol > 0) sheet.getRange(rowIndex, statusCol).setValue("MANUELL_HINZUGEFÜGT");
    if (linkCol > 0) sheet.getRange(rowIndex, linkCol).setValue(volltextDocLink || pdfUrl);
    
    // 8. Log
    if (typeof logAction === 'function') {
      logAction("Volltext-Booster", "PDF hochgeladen: " + title, Session.getActiveUser().getEmail());
    }
    
    // ✅ Auto-Refresh Trigger
    onVolltextAdded(uuid);
    
    // 9. Rückmeldung
    let resultMessage = "✅ PDF verarbeitet!\n\n";
    resultMessage += "📄 " + extractedText.length + " Zeichen extrahiert\n";
    
    if (volltextDocLink) {
      resultMessage += "📄 Als Google Doc gespeichert\n\n";
    }
    
    resultMessage += "🔗 Original-PDF: " + pdfUrl + "\n\n";
    resultMessage += "⚠️ WICHTIG: Tabellen fehlen!\n";
    resultMessage += "💡 Für Tabellen: 🧠 Smart Modus\n\n";
    resultMessage += "➡️ Jetzt Analyse starten!";
    
    return { message: resultMessage };
    
  } catch (e) {
    Logger.log("ERROR: " + e.message);
    Logger.log("Stack: " + e.stack);
    
    if (typeof logError === 'function') {
      logError("processBoosterPdf", e);
    }
    
    return { error: "Fehler: " + e.message };
  }
}

/**
 * ✅ PDF zu Text mit Drive API v3
 */
function convertPdfToTextV3(pdfBlob, tempName) {
  let tempPdf = null;
  let convertedDoc = null;
  
  try {
    // 1. Original PDF in Drive speichern
    tempPdf = DriveApp.createFile(pdfBlob);
    tempPdf.setName(tempName);
    const pdfId = tempPdf.getId();
    
    Logger.log("PDF saved: " + pdfId);
    
    // 2. Mit Drive API v3 als Google Doc kopieren
    const resource = {
      name: 'CONVERTED_' + tempName,
      mimeType: 'application/vnd.google-apps.document'
    };
    
    // Drive API v3 Syntax
    convertedDoc = Drive.Files.copy(resource, pdfId);
    
    Logger.log("Converted to Doc: " + convertedDoc.id);
    
    // 3. Text aus Google Doc extrahieren
    const doc = DocumentApp.openById(convertedDoc.id);
    const text = doc.getBody().getText();
    
    Logger.log("Text extracted: " + text.length + " chars");
    
    // 4. Cleanup
    tempPdf.setTrashed(true);
    Drive.Files.remove(convertedDoc.id);
    
    return {
      success: true,
      text: text,
      method: 'Drive API v3'
    };
    
  } catch (e) {
    Logger.log("ERROR: " + e.message);
    
    // Cleanup
    try {
      if (tempPdf) tempPdf.setTrashed(true);
      if (convertedDoc) Drive.Files.remove(convertedDoc.id);
    } catch (cleanupError) {
      // Ignore
    }
    
    return {
      success: false,
      error: e.message,
      method: 'Drive API v3 failed'
    };
  }
}

/**
 * ✅ Dateinamen bereinigen
 */
function sanitizeFilename(text) {
  if (!text) return "unnamed";
  return text
    .substring(0, 50)
    .replace(/[^a-zA-Z0-9\-_]/g, '_')
    .replace(/_+/g, '_');
}

/**
 * ✅ OCR Usage Stats anzeigen
 */
function showOcrUsageStats() {
  const stats = getOcrUsageStats();
  
  const message = `📊 OCR.space Nutzung (${stats.monthName})\n\n` +
                  `Verwendet: ${stats.currentMonth} / ${stats.limit} Requests\n` +
                  `Das sind ${stats.percentUsed}% des kostenlosen Kontingents.\n\n` +
                  `💡 Smart PDF nutzt ~3-5 Requests pro Paper\n` +
                  `   (nur für Tabellen/Figures, nicht für ganzes PDF)\n\n` +
                  `Limit Reset: Am 1. des nächsten Monats`;
  
  SpreadsheetApp.getUi().alert(message);
}

// ==========================================
// DEBUG & TESTING
// ==========================================

function testCompleteUpload() {
  try {
    const testPdfContent = "%PDF-1.4\n1 0 obj\n<<\n/Type /Catalog\n/Pages 2 0 R\n>>\nendobj\nxref\n0 5\ntrailer\n<<\n/Size 5\n/Root 1 0 R\n>>\nstartxref\n408\n%%EOF";
    const testBlob = Utilities.newBlob(testPdfContent, 'application/pdf', 'test.pdf');
    const base64 = Utilities.base64Encode(testBlob.getBytes());
    
    Logger.log("1. Test-PDF erstellt ✓");
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
    if (sheet.getLastRow() < 2) {
      SpreadsheetApp.getUi().alert("Keine Datenzeile im Dashboard vorhanden.");
      return;
    }
    
    Logger.log("2. Dashboard gefunden ✓");
    Logger.log("3. Rufe processBoosterPdf auf...");
    
    const result = processBoosterPdf(2, base64);
    
    if (result.error) {
      Logger.log("❌ FEHLER: " + result.error);
      SpreadsheetApp.getUi().alert("❌ Fehler:\n\n" + result.error);
    } else {
      Logger.log("✅ SUCCESS: " + result.message);
      SpreadsheetApp.getUi().alert("✅ Erfolg!\n\n" + result.message);
    }
    
  } catch (e) {
    Logger.log("❌ Exception: " + e.message);
    SpreadsheetApp.getUi().alert("❌ Exception:\n\n" + e.message);
  }
}

function forceAuthorization() {
  DriveApp.getRootFolder().getName();
  DocumentApp.create("TEST").getId();
  SpreadsheetApp.getActiveSpreadsheet().getName();
  UrlFetchApp.fetch("https://www.google.com");
  
  SpreadsheetApp.getUi().alert("✅ Alle Berechtigungen autorisiert!");
}
