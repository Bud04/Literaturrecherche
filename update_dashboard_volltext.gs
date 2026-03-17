// FILE: update_dashboard_volltext.gs
// ==========================================
// DASHBOARD HEADER UPDATES
// Fügt neue Spalten für Volltext-Splitting hinzu
// ==========================================

/**
 * ✅ Fügt neue Spalten für Volltext-Management hinzu
 */
function addVolltextManagementColumns() {
  const ui = SpreadsheetApp.getUi();
  
  const confirm = ui.alert(
    'Dashboard erweitern',
    'Folgende neue Spalten werden hinzugefügt:\n\n' +
    '• Volltext_Teil2 (für lange Volltexte)\n' +
    '• Volltext_Teil3 (für sehr lange Volltexte)\n' +
    '• Volltext Original Länge\n' +
    '• Volltext Bereinigt Länge\n\n' +
    'Fortfahren?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName('Dashboard');
  
  if (!dashboard) {
    ui.alert('Dashboard Sheet nicht gefunden!');
    return;
  }
  
  const headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  
  const newColumns = [
    {name: 'Volltext_Teil2', after: 'Volltext/Extrakt'},
    {name: 'Volltext_Teil3', after: 'Volltext_Teil2'},
    {name: 'Volltext Original Länge', after: 'Volltext_Teil3'},
    {name: 'Volltext Bereinigt Länge', after: 'Volltext Original Länge'}
  ];
  
  let added = 0;
  
  for (const col of newColumns) {
    // Prüfe ob Spalte schon existiert
    if (headers.indexOf(col.name) >= 0) {
      Logger.log(`Spalte "${col.name}" existiert bereits`);
      continue;
    }
    
    // Finde Position der "after"-Spalte
    const afterIndex = headers.indexOf(col.after);
    
    if (afterIndex < 0) {
      Logger.log(`Spalte "${col.after}" nicht gefunden, füge "${col.name}" am Ende hinzu`);
      dashboard.insertColumnAfter(dashboard.getLastColumn());
      dashboard.getRange(1, dashboard.getLastColumn()).setValue(col.name);
      added++;
    } else {
      // Füge nach der "after"-Spalte ein
      dashboard.insertColumnAfter(afterIndex + 1);
      dashboard.getRange(1, afterIndex + 2).setValue(col.name);
      headers.splice(afterIndex + 1, 0, col.name); // Update lokales Header-Array
      added++;
    }
  }
  
  if (added > 0) {
    ui.alert(`✅ ${added} Spalten hinzugefügt!`);
  } else {
    ui.alert('Alle Spalten existieren bereits!');
  }
}

/**
 * ✅ Zeigt Info über Volltext-Spalten
 */
function showVolltextColumnInfo() {
  const ui = SpreadsheetApp.getUi();
  const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  if (!dashboard) {
    ui.alert('Dashboard nicht gefunden!');
    return;
  }
  
  const headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  
  const volltextCols = {
    'Volltext/Extrakt': headers.indexOf('Volltext/Extrakt') + 1,
    'Volltext_Teil2': headers.indexOf('Volltext_Teil2') + 1,
    'Volltext_Teil3': headers.indexOf('Volltext_Teil3') + 1,
    'Volltext Original Länge': headers.indexOf('Volltext Original Länge') + 1,
    'Volltext Bereinigt Länge': headers.indexOf('Volltext Bereinigt Länge') + 1
  };
  
  let info = '=== VOLLTEXT SPALTEN ===\n\n';
  
  for (const [name, col] of Object.entries(volltextCols)) {
    if (col > 0) {
      const colLetter = columnToLetter(col);
      info += `✅ ${name}\n   → Spalte ${colLetter} (${col})\n\n`;
    } else {
      info += `❌ ${name}\n   → NICHT VORHANDEN\n\n`;
    }
  }
  
  info += '─────────────────────\n\n';
  info += 'VERWENDUNG:\n';
  info += '• Teil 1: Haupt-Volltext (max 49k Zeichen)\n';
  info += '• Teil 2: Fortsetzung bei >49k\n';
  info += '• Teil 3: Fortsetzung bei >98k\n';
  info += '• Original Länge: Vor Bereinigung\n';
  info += '• Bereinigt Länge: Nach Bereinigung\n\n';
  info += 'In Workspace Flows kombinieren:\n';
  info += '{Volltext/Extrakt}{Volltext_Teil2}{Volltext_Teil3}';
  
  ui.alert('Volltext Spalten Info', info, ui.ButtonSet.OK);
}

/**
 * ✅ Zeigt Volltext-Statistik
 */
function showVolltextStats() {
  const ui = SpreadsheetApp.getUi();
  const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  
  if (!dashboard) {
    ui.alert('Dashboard nicht gefunden!');
    return;
  }
  
  const totalPapers = dashboard.getLastRow() - 1;
  
  const headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  const volltext1Col = headers.indexOf('Volltext/Extrakt') + 1;
  const volltext2Col = headers.indexOf('Volltext_Teil2') + 1;
  const volltext3Col = headers.indexOf('Volltext_Teil3') + 1;
  
  if (volltext1Col === 0) {
    ui.alert('Volltext-Spalte nicht gefunden!');
    return;
  }
  
  let hasVolltext = 0;
  let hasPart2 = 0;
  let hasPart3 = 0;
  let totalChars = 0;
  
  const data1 = dashboard.getRange(2, volltext1Col, totalPapers, 1).getValues();
  const data2 = volltext2Col > 0 ? dashboard.getRange(2, volltext2Col, totalPapers, 1).getValues() : [];
  const data3 = volltext3Col > 0 ? dashboard.getRange(2, volltext3Col, totalPapers, 1).getValues() : [];
  
  for (let i = 0; i < totalPapers; i++) {
    const text1 = data1[i][0] || '';
    const text2 = volltext2Col > 0 ? (data2[i][0] || '') : '';
    const text3 = volltext3Col > 0 ? (data3[i][0] || '') : '';
    
    if (text1.length > 100) {
      hasVolltext++;
      totalChars += text1.length;
      
      if (text2.length > 0) {
        hasPart2++;
        totalChars += text2.length;
      }
      
      if (text3.length > 0) {
        hasPart3++;
        totalChars += text3.length;
      }
    }
  }
  
  const avgChars = hasVolltext > 0 ? Math.round(totalChars / hasVolltext) : 0;
  
  let stats = '=== VOLLTEXT STATISTIK ===\n\n';
  stats += `📊 Gesamt Papers: ${totalPapers}\n\n`;
  stats += `📄 Mit Volltext: ${hasVolltext} (${((hasVolltext / totalPapers) * 100).toFixed(1)}%)\n`;
  stats += `📋 Aufgeteilt (Teil 2): ${hasPart2}\n`;
  stats += `📋 Aufgeteilt (Teil 3): ${hasPart3}\n\n`;
  stats += `📏 Durchschnitt: ${avgChars.toLocaleString()} Zeichen\n`;
  stats += `📏 Total: ${totalChars.toLocaleString()} Zeichen\n\n`;
  stats += `─────────────────────\n\n`;
  
  if (hasPart2 > 0 || hasPart3 > 0) {
    stats += `⚠️ ${hasPart2 + hasPart3} Papers sind sehr lang!\n`;
    stats += `Diese nutzen mehrere Spalten.\n\n`;
    stats += `In Workspace Flows müssen alle Teile\n`;
    stats += `kombiniert werden!`;
  } else {
    stats += `✅ Alle Volltexte passen in eine Spalte!\n`;
    stats += `Teil2 & Teil3 sind leer.`;
  }
  
  ui.alert('Volltext Statistik', stats, ui.ButtonSet.OK);
}

/**
 * ✅ Hilfsfunktion: Spalten-Nummer → Buchstabe
 */
function columnToLetter(column) {
  let temp;
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}
