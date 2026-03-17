// FILE: selectiveOcr.gs
/**
 * ==========================================
 * SELECTIVE OCR - INTELLIGENTE TABELLEN-EXTRAKTION
 * ==========================================
 * Workflow:
 * 1. PDF → Google Doc (kostenlos)
 * 2. Gemini analysiert: Wo sind Figures/Tables?
 * 3. Nur diese Seiten → OCR.space API
 * 4. Kombiniere: Volltext + OCR-Tabellen
 */

/**
 * ✅ Smart PDF Processing mit Selective OCR
 */
function processBoosterPdfWithSelectiveOcr(rowIndex, base64PdfData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
    const uuid = sheet.getRange(rowIndex, 1).getValue();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const titleCol = headers.indexOf("Titel") + 1;
    const title = sheet.getRange(rowIndex, titleCol).getValue() || "Unnamed";
    
    Logger.log("Smart PDF processing for: " + title);
    
    // 1. PDF speichern
    const pdfBlob = Utilities.newBlob(
      Utilities.base64Decode(base64PdfData), 
      'application/pdf', 
      sanitizeFilename(title) + '.pdf'
    );
    
    const pdfFolder = DriveApp.getRootFolder();
    const savedPdf = pdfFolder.createFile(pdfBlob);
    const pdfUrl = savedPdf.getUrl();
    const pdfId = savedPdf.getId();
    
    Logger.log("PDF saved: " + pdfUrl);
    
    // 2. Text extrahieren mit Drive API v3
    const conversionResult = convertPdfToTextV3(pdfBlob, 'TEMP_SMART_' + uuid);
    
    if (!conversionResult.success) {
      return {
        error: "❌ Konvertierung fehlgeschlagen!\n\n" + conversionResult.error + "\n\nPDF: " + pdfUrl
      };
    }
    
    let fullText = conversionResult.text;
    Logger.log("Base text: " + fullText.length + " chars");
    
    // 3. Gemini: Finde Tabellen/Figures
    const figurePages = findFigureAndTablePages(fullText, pdfId);
    Logger.log("Gemini found: " + figurePages.length + " figures/tables");
    
    // 4. Selective OCR für identifizierte Seiten
    let ocrResults = [];
    if (figurePages && figurePages.length > 0) {
      ocrResults = performSelectiveOcr(pdfId, figurePages);
      Logger.log("OCR extracted: " + ocrResults.length + " items");
    }
    
    // 5. Kombiniere Text + OCR-Tabellen
    if (ocrResults.length > 0) {
      fullText += "\n\n" + "=".repeat(50);
      fullText += "\n📊 EXTRAHIERTE TABELLEN (OCR):\n";
      fullText += "=".repeat(50) + "\n\n";
      
      ocrResults.forEach(result => {
        fullText += `--- ${result.type} (Seiten ${result.pages.join(", ")}) ---\n`;
        fullText += result.text + "\n\n";
      });
    }
    
    fullText += "\n\n[ℹ️ " + ocrResults.length + " Tabellen extrahiert. PDF: " + pdfUrl + "]";
    
    // 6. Speichern (mit Google Doc bei >45k)
    let textToSave = fullText;
    let volltextDocLink = null;
    
    if (fullText.length > 45000) {
      const volltextDoc = DocumentApp.create('Volltext_Smart_' + sanitizeFilename(title) + '_' + uuid);
      volltextDoc.getBody().setText(fullText);
      
      const volltextDocFile = DriveApp.getFileById(volltextDoc.getId());
      volltextDocLink = volltextDocFile.getUrl();
      
      textToSave = "[📄 Text zu lang - " + fullText.length + " Zeichen]\n\n" +
                   "📄 Volltext-Doc: " + volltextDocLink + "\n" +
                   "📎 Original-PDF: " + pdfUrl + "\n\n" +
                   "✅ " + ocrResults.length + " Tabellen via OCR";
    }
    
    const volltextCol = headers.indexOf("Volltext/Extrakt") + 1;
    const statusCol = headers.indexOf("Volltext-Status") + 1;
    const linkCol = headers.indexOf("Volltext-Datei-Link") + 1;
    
    if (volltextCol > 0) sheet.getRange(rowIndex, volltextCol).setValue(textToSave);
    if (statusCol > 0) sheet.getRange(rowIndex, statusCol).setValue("VOLLTEXT_MIT_OCR");
    if (linkCol > 0) sheet.getRange(rowIndex, linkCol).setValue(volltextDocLink || pdfUrl);
    
    // 7. OCR Usage Tracking
    if (ocrResults.length > 0 && typeof trackOcrUsage === 'function') {
      for (let i = 0; i < ocrResults.length; i++) trackOcrUsage();
    }
    
    // ✅ Auto-Refresh Trigger
    onVolltextAdded(uuid);
    
    const usage = typeof getOcrUsageStats === 'function' ? getOcrUsageStats() : {currentMonth: '?', limit: 25000, percentUsed: '?'};
    
    // 8. Rückmeldung
    let resultMessage = "✅ Smart PDF verarbeitet!\n\n";
    resultMessage += "📝 " + fullText.length + " Zeichen\n";
    resultMessage += "🔍 Gemini fand: " + figurePages.length + " Seiten\n";
    resultMessage += "📊 OCR extrahiert: " + ocrResults.length + " Tabellen\n\n";
    
    if (volltextDocLink) {
      resultMessage += "📄 Als Google Doc gespeichert\n\n";
    }
    
    resultMessage += "🔗 PDF: " + pdfUrl + "\n\n";
    resultMessage += "OCR: " + usage.currentMonth + "/" + usage.limit + " (" + usage.percentUsed + "%)\n\n";
    resultMessage += "➡️ Jetzt Analyse starten!";
    
    return { message: resultMessage };
    
  } catch (e) {
    Logger.log("ERROR: " + e.message);
    
    if (typeof logError === 'function') {
      logError("processBoosterPdfWithSelectiveOcr", e);
    }
    
    return { error: "Fehler: " + e.message };
  }
}

/**
 * ✅ Gemini: Finde Figures & Tables
 */
function findFigureAndTablePages(fullText, pdfId) {
  try {
    const geminiKey = PropertiesService.getScriptProperties().getProperty("GEMINI_API_KEY");
    
    if (!geminiKey) {
      Logger.log("⚠️ Gemini API Key fehlt - skip Figure Detection");
      return [];
    }
    
    const prompt = `Analysiere folgenden wissenschaftlichen Text und identifiziere ALLE Erwähnungen von Abbildungen und Tabellen.

TEXT:
${fullText.substring(0, 50000)}

AUFGABE:
Finde alle Stellen wo folgendes erwähnt wird:
- "Figure 1", "Figure 2", "Fig. 3", etc.
- "Table 1", "Table 2", "Tabelle 1", etc.
- "Abbildung 1", "Abb. 2", etc.

WICHTIG: 
- Schätze auf welchen Seiten diese Abbildungen/Tabellen vermutlich stehen
- Wenn mehrere Figures/Tables nahe beieinander sind, gruppiere sie
- Wenn eine Tabelle über mehrere Seiten geht, gib den Range an

BEISPIEL:
"Figure 1 und Figure 2 sind vermutlich auf Seite 4-5, Table 1 geht über Seite 6-8"

ANTWORTE NUR MIT JSON (keine Erklärung):
{
  "figures_and_tables": [
    {"type": "Figure 1", "page_start": 4, "page_end": 4},
    {"type": "Figure 2", "page_start": 5, "page_end": 5},
    {"type": "Table 1 (mehrseitig)", "page_start": 6, "page_end": 8}
  ]
}`;

    const payload = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }],
      generationConfig: {
        temperature: 0.1,
        maxOutputTokens: 2000
      }
    };
    
    const options = {
      method: "post",
      contentType: "application/json",
      headers: {
        "x-goog-api-key": geminiKey
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(
      "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent",
      options
    );
    
    if (response.getResponseCode() !== 200) {
      Logger.log("Gemini API Error: " + response.getContentText());
      return [];
    }
    
    const result = JSON.parse(response.getContentText());
    
    if (!result.candidates || result.candidates.length === 0) {
      return [];
    }
    
    const textContent = result.candidates[0].content.parts[0].text;
    
    let cleanJson = textContent.replace(/```json/g, "").replace(/```/g, "").trim();
    const parsed = JSON.parse(cleanJson);
    
    return parsed.figures_and_tables || [];
    
  } catch (e) {
    logError("findFigureAndTablePages", e);
    return [];
  }
}

/**
 * ✅ Selective OCR mit 3-Seiten-Batching
 */
function performSelectiveOcr(pdfId, figurePages) {
  const results = [];
  
  try {
    const ocrApiKey = PropertiesService.getScriptProperties().getProperty("OCR_SPACE_API_KEY");
    
    if (!ocrApiKey) {
      Logger.log("⚠️ OCR.space API Key fehlt - skip OCR");
      return results;
    }
    
    // 1. Sammle alle Seiten
    const allPages = new Set();
    figurePages.forEach(item => {
      const start = item.page_start || item.estimated_page || 1;
      const end = item.page_end || item.estimated_page || start;
      
      for (let p = start; p <= end; p++) {
        allPages.add(p);
      }
    });
    
    const sortedPages = Array.from(allPages).sort((a, b) => a - b);
    
    Logger.log(`📊 OCR für ${sortedPages.length} Seiten: ${sortedPages.join(", ")}`);
    
    // 2. Gruppiere in 3er-Batches
    const batches = createPageBatches(sortedPages, 3);
    
    Logger.log(`📦 ${batches.length} Batches (max 3 Seiten pro Batch)`);
    
    // 3. Verarbeite jeden Batch
    batches.forEach((batch, batchIndex) => {
      try {
        Logger.log(`Batch ${batchIndex + 1}/${batches.length}: Seiten ${batch.join(", ")}`);
        
        const multiPageBase64 = convertPdfPagesToMultiImage(pdfId, batch);
        
        if (!multiPageBase64) {
          Logger.log(`⚠️ Batch ${batchIndex + 1} Konvertierung fehlgeschlagen`);
          return;
        }
        
        // OCR.space API Call
        const payload = {
          apikey: ocrApiKey,
          base64Image: `data:image/png;base64,${multiPageBase64}`,
          language: "ger",
          isTable: true,
          OCREngine: 2,
          scale: true,
          detectOrientation: true
        };
        
        const options = {
          method: "post",
          payload: payload,
          muteHttpExceptions: true
        };
        
        const response = UrlFetchApp.fetch("https://api.ocr.space/parse/image", options);
        const ocrResult = JSON.parse(response.getContentText());
        
        if (ocrResult.ParsedResults && ocrResult.ParsedResults.length > 0) {
          const extractedText = ocrResult.ParsedResults[0].ParsedText;
          
          const itemsInBatch = figurePages.filter(item => {
            const start = item.page_start || item.estimated_page || 1;
            const end = item.page_end || item.estimated_page || start;
            return batch.some(p => p >= start && p <= end);
          });
          
          const types = itemsInBatch.map(f => f.type).join(", ");
          
          results.push({
            type: types,
            pages: batch,
            text: extractedText
          });
          
          Logger.log(`✅ Batch ${batchIndex + 1} erfolgreich: ${extractedText.length} Zeichen`);
        }
        
        if (ocrResult.IsErroredOnProcessing) {
          logError("OCR Batch Error", new Error(ocrResult.ErrorMessage));
        }
        
        Utilities.sleep(2000);
        
      } catch (e) {
        logError("performSelectiveOcr - Batch " + (batchIndex + 1), e);
      }
    });
    
  } catch (e) {
    logError("performSelectiveOcr", e);
  }
  
  return results;
}

/**
 * ✅ Gruppiert Seiten in Batches
 */
function createPageBatches(pages, maxPerBatch) {
  const batches = [];
  let currentBatch = [];
  
  pages.forEach((page, index) => {
    currentBatch.push(page);
    
    const nextPage = pages[index + 1];
    const isLastPage = index === pages.length - 1;
    const nextNotConsecutive = nextPage && (nextPage !== page + 1);
    
    if (currentBatch.length === maxPerBatch || nextNotConsecutive || isLastPage) {
      batches.push([...currentBatch]);
      currentBatch = [];
    }
  });
  
  return batches;
}

/**
 * ✅ Konvertiert PDF-Seiten zu Image
 */
function convertPdfPagesToMultiImage(pdfId, pageNumbers) {
  try {
    const pdf = DriveApp.getFileById(pdfId);
    const pdfBlob = pdf.getBlob();
    
    if (pageNumbers.length === 1) {
      const url = `https://drive.google.com/thumbnail?id=${pdfId}&sz=w2000`;
      const response = UrlFetchApp.fetch(url, {
        headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      });
      
      if (response.getResponseCode() === 200) {
        return Utilities.base64Encode(response.getBlob().getBytes());
      }
    }
    
    // Für Multiple Pages: PDF direkt senden
    const pdfBase64 = Utilities.base64Encode(pdfBlob.getBytes());
    return pdfBase64;
    
  } catch (e) {
    logError("convertPdfPagesToMultiImage", e);
    return null;
  }
}

/**
 * ✅ OCR Usage Tracking
 */
function trackOcrUsage() {
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  
  const storedMonth = props.getProperty("OCR_USAGE_MONTH");
  let count = parseInt(props.getProperty("OCR_USAGE_COUNT")) || 0;
  
  if (storedMonth !== currentMonth) {
    count = 0;
    props.setProperty("OCR_USAGE_MONTH", currentMonth);
  }
  
  count++;
  props.setProperty("OCR_USAGE_COUNT", String(count));
  
  return count;
}

function getOcrUsageStats() {
  const props = PropertiesService.getScriptProperties();
  const now = new Date();
  const currentMonth = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
  
  const storedMonth = props.getProperty("OCR_USAGE_MONTH");
  let count = parseInt(props.getProperty("OCR_USAGE_COUNT")) || 0;
  
  if (storedMonth !== currentMonth) {
    count = 0;
  }
  
  return {
    currentMonth: count,
    limit: 25000,
    monthName: currentMonth,
    percentUsed: ((count / 25000) * 100).toFixed(1)
  };
}
