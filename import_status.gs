var IMPORT_STATUS_SHEET = 'Import_Status';

function ensureImportStatusSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(IMPORT_STATUS_SHEET);
  if (sheet) return sheet;

  sheet = ss.insertSheet(IMPORT_STATUS_SHEET);
  sheet.getRange(1, 1, 1, 5).setValues([[
    'Kategorie', 'Von', 'Bis', 'Import-Datum', 'Anzahl'
  ]]);
  sheet.getRange(1, 1, 1, 5)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('white');
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 120);
  sheet.setColumnWidth(3, 120);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 80);
  return sheet;
}

function saveImportRange(kategorie, vonDate, bisDate, anzahl) {
  var sheet = ensureImportStatusSheet();
  sheet.appendRow([
    kategorie,
    vonDate  || '–',
    bisDate  || '–',
    new Date(),
    anzahl   || 0
  ]);
  SpreadsheetApp.flush();
  Logger.log('[IMPORT_STATUS] ' + kategorie + ' | ' + vonDate + ' → ' + bisDate + ' | ' + anzahl + ' Papers');
}

function getLastImportDateForCategory(kategorie) {
  var sheet   = ensureImportStatusSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  var data      = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  var latestBis = null;

  for (var r = 0; r < data.length; r++) {
    if (String(data[r][0]).trim() === kategorie) {
      var bis = String(data[r][2]).trim();
      if (bis && bis !== '–' && (!latestBis || bis > latestBis)) latestBis = bis;
    }
  }
  return latestBis;
}

// Rückwärtskompatibilität mit altem Code
function getLastImportDate(key) {
  return getLastImportDateForCategory(key) ||
         PropertiesService.getDocumentProperties().getProperty('lastPubDate_' + key) ||
         null;
}

function getImportedRanges(kategorie) {
  var sheet   = ensureImportStatusSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data   = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
  var ranges = [];

  for (var r = 0; r < data.length; r++) {
    if (String(data[r][0]).trim() === kategorie) {
      ranges.push({
        von:   String(data[r][1]).trim(),
        bis:   String(data[r][2]).trim(),
        datum: data[r][3],
        count: data[r][4]
      });
    }
  }
  ranges.sort(function(a, b) { return a.von < b.von ? -1 : 1; });
  return ranges;
}

function findGapsForCategory(kategorie) {
  var ranges = getImportedRanges(kategorie);
  if (ranges.length < 2) return [];

  var gaps = [];
  for (var i = 0; i < ranges.length - 1; i++) {
    var bisObj = parseDateStr(ranges[i].bis);
    var vonObj = parseDateStr(ranges[i+1].von);
    if (!bisObj || !vonObj) continue;

    var nextDay = new Date(bisObj.getTime());
    nextDay.setDate(nextDay.getDate() + 1);
    if (nextDay < vonObj) {
      gaps.push({
        von: formatDateStr(nextDay),
        bis: formatDateStr(addDays(vonObj, -1))
      });
    }
  }
  return gaps;
}

function getAllImportKategorien() {
  var sheet   = ensureImportStatusSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var data   = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  var seen   = {};
  var result = [];
  for (var r = 0; r < data.length; r++) {
    var k = String(data[r][0]).trim();
    if (k && k !== '–' && !seen[k]) { seen[k] = true; result.push(k); }
  }
  return result;
}

function getQueryForKategorie(kategorieName) {
  if (typeof IMPORT_KATEGORIEN !== 'undefined' && IMPORT_KATEGORIEN[kategorieName]) {
    return IMPORT_KATEGORIEN[kategorieName].query;
  }
  return null;
}

function showGapOverview() {
  var ui         = SpreadsheetApp.getUi();
  var kategorien = getAllImportKategorien();

  if (kategorien.length === 0) {
    ui.alert('Keine Import-Historien gefunden.\nFühre erst einen Import durch.');
    return;
  }

  var msg     = '=== IMPORT LÜCKEN ÜBERSICHT ===\n\n';
  var hasGaps = false;

  for (var k = 0; k < kategorien.length; k++) {
    var kat    = kategorien[k];
    var gaps   = findGapsForCategory(kat);
    var ranges = getImportedRanges(kat);

    msg += '📂 ' + kat + '\n';
    if (ranges.length > 0) {
      msg += '   Importiert: ' + ranges[0].von + ' → ' + ranges[ranges.length-1].bis + '\n';
    }
    if (gaps.length > 0) {
      hasGaps = true;
      for (var g = 0; g < gaps.length; g++) {
        msg += '   ⚠️ Lücke: ' + gaps[g].von + ' → ' + gaps[g].bis + '\n';
      }
    } else {
      msg += '   ✅ Keine Lücken\n';
    }
    msg += '\n';
  }

  if (!hasGaps) {
    ui.alert('Lücken-Übersicht', msg + '✅ Alle Kategorien lückenlos!', ui.ButtonSet.OK);
    return;
  }

  var response = ui.alert(
    'Lücken-Übersicht',
    msg + 'Lücken jetzt schließen?',
    ui.ButtonSet.YES_NO
  );
  if (response === ui.Button.YES) showGapFillMenu(kategorien);
}

function showGapFillMenu(kategorien) {
  var ui      = SpreadsheetApp.getUi();
  var gapKats = [];

  for (var k = 0; k < kategorien.length; k++) {
    var gaps = findGapsForCategory(kategorien[k]);
    if (gaps.length > 0) gapKats.push({ kat: kategorien[k], gaps: gaps });
  }

  if (gapKats.length === 0) { ui.alert('Keine Lücken gefunden!'); return; }

  var msg = 'Welche Lücken schließen?\n\n';
  msg    += '0. Alle Lücken schließen\n\n';
  for (var i = 0; i < gapKats.length; i++) {
    msg += (i+1) + '. ' + gapKats[i].kat + ' (' + gapKats[i].gaps.length + ' Lücke(n))\n';
    for (var g = 0; g < gapKats[i].gaps.length; g++) {
      msg += '   → ' + gapKats[i].gaps[g].von + ' bis ' + gapKats[i].gaps[g].bis + '\n';
    }
  }

  var resp = ui.prompt('Lücken schließen', msg, ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  var choice = parseInt(resp.getResponseText().trim());
  var toFill = [];
  if (choice === 0) {
    toFill = gapKats;
  } else if (choice >= 1 && choice <= gapKats.length) {
    toFill = [gapKats[choice - 1]];
  } else {
    ui.alert('Ungültige Auswahl!'); return;
  }

  var countResp = ui.prompt(
    'Anzahl pro Lücke',
    'Wie viele Papers pro Lücke importieren? (1-200)\n\n' +
    '⚠️ Falls die Lücke größer ist als die Anzahl,\n' +
    'wird der Fortschritt gespeichert und du kannst\n' +
    'beim nächsten Mal weitermachen.',
    ui.ButtonSet.OK_CANCEL
  );
  if (countResp.getSelectedButton() !== ui.Button.OK) return;
  var maxPerGap = parseInt(countResp.getResponseText()) || 50;

  var totalFilled = 0;
  for (var t = 0; t < toFill.length; t++) {
    var item  = toFill[t];
    var query = getQueryForKategorie(item.kat);
    if (!query) { Logger.log('Kein Query für: ' + item.kat); continue; }

    for (var g = 0; g < item.gaps.length; g++) {
      var gap = item.gaps[g];
      Logger.log('[GAP-FILL] ' + item.kat + ' | ' + gap.von + ' → ' + gap.bis);

      // Import starten — übergibt gap.bis als gewünschtes Enddatum
      // saveImportRange wird in importPapersFromPubMed mit dem
      // tatsächlich erreichten Datum aufgerufen
      importPapersFromPubMed(query, gap.von, maxPerGap, item.kat, gap.bis);
      totalFilled++;
    }
  }
  ui.alert('✅ ' + totalFilled + ' Lücke(n) bearbeitet!\n\nFalls Papers limit erreicht wurde,\nsiehe Lücken-Übersicht für verbleibende Lücken.');
}

// ── Datums-Helpers ────────────────────────────────────────────────
function parseDateStr(str) {
  if (!str || str === '–') return null;
  var parts = str.split('/');
  if (parts.length !== 3) return null;
  return new Date(parseInt(parts[0]), parseInt(parts[1])-1, parseInt(parts[2]), 12, 0, 0);
}

function formatDateStr(date) {
  return Utilities.formatDate(date, 'GMT+1', 'yyyy/MM/dd');
}

function addDays(date, days) {
  var d = new Date(date.getTime());
  d.setDate(d.getDate() + days);
  return d;
}
