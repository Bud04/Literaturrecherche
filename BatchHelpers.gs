// FILE: batchHelpers.gs

/**
 * ==========================================
 * BATCH HELPER-FUNKTIONEN
 * ==========================================
 * Zusätzliche Utilities für bessere UX
 */

/**
 * ✅ Färbt Batch-Phase Spalte basierend auf Status
 * Aufruf: Menü → System → Batch-Phase Farben aktualisieren
 */
function updateBatchPhaseColors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  if (sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("Keine Daten im Dashboard");
    return;
  }
  
  const batchPhaseCol = DASHBOARD_HEADERS.indexOf("Batch-Phase") + 1;
  if (batchPhaseCol === 0) {
    SpreadsheetApp.getUi().alert("Spalte 'Batch-Phase' nicht gefunden");
    return;
  }
  
  const data = sheet.getRange(2, batchPhaseCol, sheet.getLastRow() - 1, 1).getValues();
  
  const colorMap = {
    "Prompt generiert": "#fff2cc",     // Gelb - wartet auf Antwort
    "Abgeschlossen ✓": "#d9ead3",      // Grün - fertig
    "Fehler": "#f4cccc",               // Rot - Fehler
    "": "#ffffff"                       // Weiß - leer
  };
  
  let updated = 0;
  
  data.forEach((row, index) => {
    const phaseText = row[0];
    let color = "#ffffff"; // Default
    
    // Prüfe ob Phrase im Text vorkommt
    Object.keys(colorMap).forEach(key => {
      if (phaseText && phaseText.toString().includes(key)) {
        color = colorMap[key];
      }
    });
    
    sheet.getRange(index + 2, batchPhaseCol).setBackground(color);
    updated++;
  });
  
  SpreadsheetApp.getUi().alert(`${updated} Zeilen eingefärbt`);
  logAction("Batch", "Phase-Farben aktualisiert");
}

/**
 * ✅ Dashboard Filter: Zeige nur Zeilen einer bestimmten Batch-Phase
 */
function filterByBatchPhase() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt(
    'Batch-Phase Filter',
    'Welche Phase anzeigen?\n\n' +
    'Optionen:\n' +
    '- Phase 0\n' +
    '- Phase 1\n' +
    '- Phase 2\n' +
    '- Phase 3\n' +
    '- Phase 4\n' +
    '- Phase 5\n' +
    '- Phase 6\n' +
    '- Prompt generiert\n' +
    '- Abgeschlossen\n' +
    '- (leer lassen = alle zeigen)',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  const filterText = response.getResponseText().trim();
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  // Alle Zeilen einblenden
  sheet.showRows(1, sheet.getMaxRows());
  
  if (!filterText) {
    ui.alert("Alle Zeilen werden angezeigt");
    return;
  }
  
  const batchPhaseCol = DASHBOARD_HEADERS.indexOf("Batch-Phase");
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, DASHBOARD_HEADERS.length).getValues();
  
  let hiddenCount = 0;
  
  data.forEach((row, index) => {
    const phaseText = row[batchPhaseCol] || "";
    
    if (!phaseText.toString().toLowerCase().includes(filterText.toLowerCase())) {
      sheet.hideRows(index + 2);
      hiddenCount++;
    }
  });
  
  ui.alert(`Filter aktiv: "${filterText}"\n${data.length - hiddenCount} Zeilen sichtbar, ${hiddenCount} ausgeblendet`);
}

/**
 * ✅ Zeige Batch-Progress als Sankey-Diagramm (Text-basiert)
 */
function showBatchProgress() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  
  if (sheet.getLastRow() <= 1) {
    SpreadsheetApp.getUi().alert("Keine Daten im Dashboard");
    return;
  }
  
  const batchPhaseCol = DASHBOARD_HEADERS.indexOf("Batch-Phase");
  const data = sheet.getRange(2, batchPhaseCol, sheet.getLastRow() - 1, 1).getValues();
  
  // Zähle Phase-Status
  const stats = {
    "Leer / Nicht gestartet": 0,
    "Phase 0 - Prompt generiert": 0,
    "Phase 0 - Abgeschlossen ✓": 0,
    "Phase 1 - Prompt generiert": 0,
    "Phase 1 - Abgeschlossen ✓": 0,
    "Phase 2 - Prompt generiert": 0,
    "Phase 2 - Abgeschlossen ✓": 0,
    "Phase 3 - Prompt generiert": 0,
    "Phase 3 - Abgeschlossen ✓": 0,
    "Phase 4 - Prompt generiert": 0,
    "Phase 4 - Abgeschlossen ✓": 0,
    "Phase 5 - Prompt generiert": 0,
    "Phase 5 - Abgeschlossen ✓": 0,
    "Phase 6 - Prompt generiert": 0,
    "Phase 6 - Abgeschlossen ✓": 0
  };
  
  data.forEach(row => {
    const phase = row[0] || "";
    
    if (!phase || phase === "") {
      stats["Leer / Nicht gestartet"]++;
      return;
    }
    
    // Exakte Matches
    Object.keys(stats).forEach(key => {
      if (phase.toString().includes(key.replace("Leer / Nicht gestartet", ""))) {
        stats[key]++;
      }
    });
  });
  
  // Formatiere Output
  let message = "=== BATCH PIPELINE PROGRESS ===\n\n";
  
  message += `📊 Gesamt: ${data.length} Publikationen\n\n`;
  
  Object.keys(stats).forEach(key => {
    if (stats[key] > 0) {
      const bar = "█".repeat(Math.ceil(stats[key] / 2));
      message += `${key.padEnd(35)} ${stats[key].toString().padStart(3)} ${bar}\n`;
    }
  });
  
  SpreadsheetApp.getUi().alert(message);
}

/**
 * ✅ Reset Batch-Phase für ausgewählte Zeilen
 */
function resetBatchPhaseForSelected() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(
    'Batch-Phase zurücksetzen',
    'Batch-Phase für alle markierten Zeilen löschen?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(DASHBOARD_SHEET_NAME);
  const selection = sheet.getActiveRange();
  
  if (!selection) {
    ui.alert("Bitte Zeilen auswählen");
    return;
  }
  
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  const batchPhaseCol = DASHBOARD_HEADERS.indexOf("Batch-Phase") + 1;
  
  if (batchPhaseCol === 0) {
    ui.alert("Spalte 'Batch-Phase' nicht gefunden");
    return;
  }
  
  for (let i = 0; i < numRows; i++) {
    sheet.getRange(startRow + i, batchPhaseCol).setValue("");
  }
  
  ui.alert(`${numRows} Zeilen zurückgesetzt`);
  logAction("Batch", `${numRows} Batch-Phasen zurückgesetzt`);
}
