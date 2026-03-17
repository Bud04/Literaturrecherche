// FILE: gemini_api.gs
/**
 * ==========================================
 * GEMINI API INTEGRATION
 * Vollständige Integration - funktioniert sobald API Key gesetzt ist
 * ==========================================
 */

/**
 * ✅ Analysiert Papers mit Gemini API (automatisch)
 */
function analyzeWithGeminiAPI() {
  const ui = SpreadsheetApp.getUi();
  
  // Check API Key
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    ui.alert(
      'Gemini API Key fehlt!',
      'Um die automatische Analyse zu nutzen, musst du einen Gemini API Key einrichten.\n\n' +
      'Gehe zu:\n' +
      '1. Erweiterungen → Apps Script\n' +
      '2. Projekt-Einstellungen ⚙️\n' +
      '3. Script-Properties\n' +
      '4. Property hinzufügen:\n' +
      '   Name: GEMINI_API_KEY\n' +
      '   Wert: [dein API Key]\n\n' +
      'API Key bekommst du auf: https://aistudio.google.com/apikey',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Frage welche Papers
  const response = ui.alert(
    'Papers analysieren',
    'Welche Papers sollen analysiert werden?',
    ui.ButtonSet.YES_NO_CANCEL
  );
  
  let papers = [];
  
  if (response === ui.Button.YES) {
    // Alle Papers mit Status "NEU"
    papers = getPapersForAnalysis("NEU");
  } else if (response === ui.Button.NO) {
    // Nur ausgewählte Zeilen
    papers = getSelectedPapersForAnalysis();
  } else {
    return;
  }
  
  if (papers.length === 0) {
    ui.alert('Keine Papers zum Analysieren gefunden!');
    return;
  }
  
  ui.alert(`Starte Analyse von ${papers.length} Papers...\n\nDies kann einige Minuten dauern.`);
  
  // Analysiere!
  analyzeMultiplePapers(papers, apiKey);
}

/**
 * ✅ Analysiert mehrere Papers
 */
function analyzeMultiplePapers(papers, apiKey) {
  let successCount = 0;
  let failCount = 0;
  
  for (let i = 0; i < papers.length; i++) {
    const paper = papers[i];
    
    Logger.log(`\n[${i+1}/${papers.length}] ${paper.title.substring(0, 50)}...`);
    
    try {
      const analysis = callGeminiAPI(paper, apiKey);
      
      if (analysis) {
        writeAnalysisToDashboard(paper.uuid, analysis);
        successCount++;
        Logger.log(`  ✅ Erfolgreich analysiert`);
      } else {
        failCount++;
        Logger.log(`  ❌ Analyse fehlgeschlagen`);
      }
      
    } catch (e) {
      failCount++;
      Logger.log(`  ❌ Error: ${e.message}`);
    }
    
    // Rate limiting
    Utilities.sleep(2000);
  }
  
  // Ergebnis
  let message = `=== GEMINI ANALYSE ABGESCHLOSSEN ===\n\n`;
  message += `Papers analysiert: ${papers.length}\n`;
  message += `✅ Erfolgreich: ${successCount}\n`;
  message += `❌ Fehlgeschlagen: ${failCount}`;
  
  SpreadsheetApp.getUi().alert(message);
  
  if (typeof logAction === 'function') {
    logAction("Gemini API Analyse", `${successCount}/${papers.length} Papers analysiert`);
  }
}

/**
 * ✅ Ruft Gemini API auf
 */
function callGeminiAPI(paper, apiKey) {
  try {
    const model = "gemini-2.0-flash-exp";
    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    
    // Baue Prompt
    const prompt = buildGeminiPrompt(paper);
    
    // API Request
    const payload = {
      contents: [{
        parts: [{
          text: prompt
        }]
      }],
      generationConfig: {
        temperature: 0.7,
        topK: 40,
        topP: 0.95,
        maxOutputTokens: 4096
      }
    };
    
    const options = {
      method: 'POST',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    Logger.log(`  → Rufe Gemini API auf...`);
    
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    
    if (responseCode !== 200) {
      Logger.log(`  ❌ API Error: ${responseCode}`);
      Logger.log(response.getContentText());
      return null;
    }
    
    const data = JSON.parse(response.getContentText());
    
    if (data.candidates && data.candidates[0] && data.candidates[0].content) {
      const text = data.candidates[0].content.parts[0].text;
      Logger.log(`  ✅ Antwort erhalten (${text.length} Zeichen)`);
      return text;
    }
    
    return null;
    
  } catch (e) {
    Logger.log(`  ❌ callGeminiAPI Error: ${e.message}`);
    return null;
  }
}

/**
 * ✅ Baut Prompt für Gemini
 */
function buildGeminiPrompt(paper) {
  // Nutze Volltext wenn vorhanden, sonst Abstract
  const content = paper.volltext || paper.abstract;
  
  const prompt = `Analysiere dieses wissenschaftliche Paper im Kontext von VitalAire (Medizintechnik für Diabetes & Atemwegserkrankungen):

**PAPER DETAILS:**
Titel: ${paper.title}
Autoren: ${paper.authors}
Journal: ${paper.journal} (${paper.year})
DOI: ${paper.doi}
PMID: ${paper.pmid}

**INHALT:**
${content}

**AUFGABEN:**

1. **RELEVANZ** (Hoch/Mittel/Niedrig):
   Bewerte die Relevanz für VitalAire's Produktportfolio (CGM, Insulinpumpen, Atemtherapie).

2. **RELEVANZ-BEGRÜNDUNG** (2-3 Sätze):
   Erkläre warum das Paper diese Relevanz hat.

3. **PRODUKT-FOKUS**:
   - VitalAire Kernprodukt (mit positivem Ergebnis)
   - VitalAire Kernprodukt (mit negativem/neutralem Ergebnis)
   - Konkurrenz (Erwähnung)
   - Konkurrenz (Head-to-Head Vergleich)
   - Technologie-Trend
   - Regulierung/Policy
   - Keine direkte Produktrelevanz

4. **HAUPTERKENNTNIS** (1 Satz):
   Die wichtigste Erkenntnis des Papers.

5. **KERNAUSSAGEN** (3-5 Bullet Points):
   Die zentralen Findings.

6. **ZUSAMMENFASSUNG** (3-4 Sätze):
   Kompakte Zusammenfassung des Papers.

7. **PRAKTISCHE IMPLIKATIONEN** (2-3 Sätze):
   Was bedeutet das für die Praxis?

8. **KRITISCHE BEWERTUNG** (1-2 Sätze):
   Limitationen oder kritische Punkte.

9. **EVIDENZGRAD** (Ia/Ib/IIa/IIb/III/IV):
   Nach Oxford Centre for Evidence-Based Medicine.

Formatiere die Antwort als strukturierten Text mit klaren Überschriften.`;

  return prompt;
}

/**
 * ✅ Schreibt Analyse ins Dashboard
 */
function writeAnalysisToDashboard(uuid, analysis) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName("Dashboard");
    
    if (!dashboard) return;
    
    const headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
    const data = dashboard.getRange(2, 1, dashboard.getLastRow() - 1, headers.length).getValues();
    
    const uuidCol = headers.indexOf("UUID");
    const geminiCol = headers.indexOf("Gemini Master Analyse");
    
    // Finde Zeile
    for (let i = 0; i < data.length; i++) {
      if (data[i][uuidCol] === uuid) {
        const row = i + 2;
        
        // Schreibe Analyse
        dashboard.getRange(row, geminiCol + 1).setValue(analysis);
        
        // Parse und schreibe strukturierte Felder (optional)
        parseAndWriteStructuredAnalysis(dashboard, row, headers, analysis);
        
        break;
      }
    }
    
  } catch (e) {
    Logger.log(`writeAnalysisToDashboard Error: ${e.message}`);
  }
}

/**
 * ✅ Parst Analyse und schreibt in strukturierte Spalten
 */
function parseAndWriteStructuredAnalysis(dashboard, row, headers, analysis) {
  try {
    // Extrahiere Relevanz
    const relevanzMatch = analysis.match(/RELEVANZ[:\s]*\*?\*?(Hoch|Mittel|Niedrig)/i);
    if (relevanzMatch) {
      const relevanzCol = headers.indexOf("Relevanz");
      if (relevanzCol >= 0) {
        dashboard.getRange(row, relevanzCol + 1).setValue(relevanzMatch[1]);
      }
    }
    
    // Extrahiere Haupterkenntnis
    const erkenntnisMatch = analysis.match(/HAUPTERKENNTNIS[:\s]*\*?\*?([^\n]+)/i);
    if (erkenntnisMatch) {
      const erkenntnisCol = headers.indexOf("Haupterkenntnis");
      if (erkenntnisCol >= 0) {
        dashboard.getRange(row, erkenntnisCol + 1).setValue(erkenntnisMatch[1].trim());
      }
    }
    
    // Extrahiere Evidenzgrad
    const evidenzMatch = analysis.match(/EVIDENZGRAD[:\s]*\*?\*?(Ia|Ib|IIa|IIb|III|IV)/i);
    if (evidenzMatch) {
      const evidenzCol = headers.indexOf("Evidenzgrad");
      if (evidenzCol >= 0) {
        dashboard.getRange(row, evidenzCol + 1).setValue(evidenzMatch[1]);
      }
    }
    
    // Setze Status
    const statusCol = headers.indexOf("Status");
    if (statusCol >= 0) {
      dashboard.getRange(row, statusCol + 1).setValue("ANALYSIERT");
    }
    
  } catch (e) {
    Logger.log(`parseAndWriteStructuredAnalysis Error: ${e.message}`);
  }
}

/**
 * ✅ Bereitet Gemini Prompt für manuelles Gemini vor
 */
function prepareGeminiPromptForPaper(uuid) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName("Dashboard");
    
    if (!dashboard) return;
    
    const headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
    const data = dashboard.getRange(2, 1, dashboard.getLastRow() - 1, headers.length).getValues();
    
    const uuidCol = headers.indexOf("UUID");
    
    // Finde Paper
    for (let i = 0; i < data.length; i++) {
      if (data[i][uuidCol] === uuid) {
        const paper = {
          uuid: uuid,
          title: data[i][headers.indexOf("Titel")],
          authors: data[i][headers.indexOf("Autoren")],
          journal: data[i][headers.indexOf("Journal/Quelle")],
          year: data[i][headers.indexOf("Jahr")],
          doi: data[i][headers.indexOf("DOI")],
          pmid: data[i][headers.indexOf("PMID")],
          abstract: data[i][headers.indexOf("Inhalt/Abstract")],
          volltext: data[i][headers.indexOf("Volltext/Extrakt")]
        };
        
        const prompt = buildGeminiPrompt(paper);
        
        // Schreibe Prompt in Spalte (für manuelles Gemini)
        const promptCol = headers.indexOf("Gemini Master Analyse");
        if (promptCol >= 0) {
          dashboard.getRange(i + 2, promptCol + 1).setValue(prompt);
        }
        
        break;
      }
    }
    
  } catch (e) {
    Logger.log(`prepareGeminiPromptForPaper Error: ${e.message}`);
  }
}

/**
 * ✅ Bereitet Prompts für alle neuen Papers vor (manuelles Gemini)
 */
function prepareAllGeminiPrompts() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Prompts vorbereiten',
    'Sollen Gemini-Prompts für alle neuen Papers vorbereitet werden?\n\n' +
    'Du kannst dann "Gemini in Sheets" manuell aufrufen.',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    return;
  }
  
  const papers = getPapersForAnalysis("NEU");
  
  if (papers.length === 0) {
    ui.alert('Keine neuen Papers gefunden!');
    return;
  }
  
  let count = 0;
  
  for (const paper of papers) {
    prepareGeminiPromptForPaper(paper.uuid);
    count++;
  }
  
  ui.alert(`✅ ${count} Prompts vorbereitet!\n\nDu kannst jetzt "Gemini in Sheets" nutzen.`);
}

/**
 * ✅ Holt Papers für Analyse
 */
function getPapersForAnalysis(status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName("Dashboard");
    
    if (!dashboard || dashboard.getLastRow() < 2) {
      return [];
    }
    
    const headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
    const data = dashboard.getRange(2, 1, dashboard.getLastRow() - 1, headers.length).getValues();
    
    const papers = [];
    
    for (let i = 0; i < data.length; i++) {
      const rowStatus = String(data[i][headers.indexOf("Status")] || "").trim();
      
      if (rowStatus === status) {
        papers.push({
          uuid: data[i][headers.indexOf("UUID")],
          title: data[i][headers.indexOf("Titel")],
          authors: data[i][headers.indexOf("Autoren")],
          journal: data[i][headers.indexOf("Journal/Quelle")],
          year: data[i][headers.indexOf("Jahr")],
          doi: data[i][headers.indexOf("DOI")],
          pmid: data[i][headers.indexOf("PMID")],
          abstract: data[i][headers.indexOf("Inhalt/Abstract")],
          volltext: data[i][headers.indexOf("Volltext/Extrakt")]
        });
      }
    }
    
    return papers;
    
  } catch (e) {
    Logger.log(`getPapersForAnalysis Error: ${e.message}`);
    return [];
  }
}

/**
 * ✅ Holt ausgewählte Papers für Analyse
 */
function getSelectedPapersForAnalysis() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const dashboard = ss.getSheetByName("Dashboard");
    const selection = dashboard.getActiveRange();
    
    if (!selection) {
      return [];
    }
    
    const startRow = selection.getRow();
    const numRows = selection.getNumRows();
    
    if (startRow < 2) {
      return [];
    }
    
    const headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
    const data = dashboard.getRange(startRow, 1, numRows, headers.length).getValues();
    
    const papers = [];
    
    for (const row of data) {
      papers.push({
        uuid: row[headers.indexOf("UUID")],
        title: row[headers.indexOf("Titel")],
        authors: row[headers.indexOf("Autoren")],
        journal: row[headers.indexOf("Journal/Quelle")],
        year: row[headers.indexOf("Jahr")],
        doi: row[headers.indexOf("DOI")],
        pmid: row[headers.indexOf("PMID")],
        abstract: row[headers.indexOf("Inhalt/Abstract")],
        volltext: row[headers.indexOf("Volltext/Extrakt")]
      });
    }
    
    return papers;
    
  } catch (e) {
    Logger.log(`getSelectedPapersForAnalysis Error: ${e.message}`);
    return [];
  }
}
