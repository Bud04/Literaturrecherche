// ================================================================
// DUPLICATE CHECK
// ================================================================

// ── Einstieg: bestehende Papers prüfen (Menü) ────────────────────
function checkExistingDuplicates() {
  var result = findAndProcessDuplicates(null);
  if (result.autoCleaned > 0) {
    SpreadsheetApp.getUi().alert(
      '🧹 ' + result.autoCleaned + ' Duplikate automatisch bereinigt.\n\n' +
      (result.manualPairs.length > 0
        ? result.manualPairs.length + ' weitere Paare brauchen deine Entscheidung.'
        : 'Keine weiteren Duplikate gefunden.')
    );
  }
  if (result.manualPairs.length === 0) {
    if (result.autoCleaned === 0) {
      SpreadsheetApp.getUi().alert('✅ Keine Duplikate gefunden!');
    }
    return;
  }
  showDuplicateDialog(result.manualPairs);
}

// ── Einstieg: nach Import ─────────────────────────────────────────
function checkDuplicatesAfterImport(newUuids) {
  var result = findAndProcessDuplicates(newUuids);
  if (result.manualPairs.length === 0) return;
  showDuplicateDialog(result.manualPairs);
}

// ── Dialog öffnen ─────────────────────────────────────────────────
function showDuplicateDialog(pairs) {
  var html = HtmlService.createHtmlOutputFromFile('duplicateCheckDialog')
    .setWidth(900)
    .setHeight(700)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(
    html, '🔍 Duplikat-Check – ' + pairs.length + ' mögliche Duplikate'
  );
}

// ── Haupt-Logik: finden + auto-bereinigen ─────────────────────────
function findAndProcessDuplicates(filterUuids) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) return { autoCleaned: 0, manualPairs: [] };

  var headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  var lastRow = dashboard.getLastRow();
  if (lastRow < 3) return { autoCleaned: 0, manualPairs: [] };

  var data = dashboard.getRange(2, 1, lastRow - 1, headers.length).getValues();

  var getIdx = function(name) {
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c]).trim() === name) return c;
    }
    return -1;
  };

  var isValid = function(val) {
    if (!val) return false;
    var v = val.toString().trim().toLowerCase();
    return v !== '' && v !== 'n/a' && v !== 'na' && v !== '-'
        && v !== 'null' && v !== 'undefined';
  };

  // Zählt wie viele Felder valide befüllt sind
  var countValidFields = function(p) {
    var score = 0;
    if (isValid(p.doi)) score++;
    if (isValid(p.pmid)) score++;
    if (isValid(p.journal)) score++;
    if (isValid(p.autoren)) score++;
    if (isValid(p.pubdate)) score++;
    return score;
  };

  // Relevanz aus Status extrahieren
  var getRelevanz = function(status) {
    if (!status) return '';
    var s = status.toLowerCase();
    if (s.indexOf('irrelevant') >= 0) return 'Irrelevant';
    if (s.indexOf('niedrig') >= 0) return 'Niedrig';
    if (s.indexOf('mittel') >= 0) return 'Mittel';
    if (s.indexOf('hoch') >= 0) return 'Hoch';
    if (s.indexOf('sehr hoch') >= 0) return 'Sehr hoch';
    if (s.indexOf('komplett') >= 0) return 'Komplett';
    return '';
  };

  // Papers laden
  var papers = [];
  for (var r = 0; r < data.length; r++) {
    var uuid = String(data[r][getIdx('UUID')] || '').trim();
    var titel = String(data[r][getIdx('Titel')] || '').trim();
    if (!uuid || !titel) continue;

    papers.push({
      uuid: uuid,
      rowNum: r + 2,
      doi: String(data[r][getIdx('DOI')] || '').trim().toLowerCase(),
      titel: titel.toLowerCase(),
      titelOrig: titel,
      pmid: String(data[r][getIdx('PMID')] || '').trim(),
      link: String(data[r][getIdx('Link')] || '').trim(),
      autoren: String(data[r][getIdx('Autoren')] || '').trim(),
      journal: String(data[r][getIdx('Journal/Quelle')] || '').trim(),
      pubdate: String(data[r][getIdx('Publikationsdatum')]|| '').trim(),
      status: String(data[r][getIdx('Status')] || '').trim(),
      isNew: filterUuids ? filterUuids.indexOf(uuid) >= 0 : false
    });
  }

  var autoCleaned = 0;
  var manualPairs = [];
  var seen = {};
  var deleted = {}; // bereits gelöschte UUIDs nicht nochmal prüfen

  for (var i = 0; i < papers.length; i++) {
    for (var j = i + 1; j < papers.length; j++) {
      var a = papers[i];
      var b = papers[j];

      if (deleted[a.uuid] || deleted[b.uuid]) continue;
      if (filterUuids && filterUuids.length > 0 && !a.isNew && !b.isNew) continue;

      var pairKey = [a.uuid, b.uuid].sort().join('|');
      if (seen[pairKey]) continue;

      var matchType = '';
      var confidence = 0;

      if (isValid(a.doi) && isValid(b.doi) && a.doi === b.doi) {
        matchType = 'DOI identisch'; confidence = 100;
      } else if (isValid(a.pmid) && isValid(b.pmid) && a.pmid === b.pmid) {
        matchType = 'PMID identisch'; confidence = 100;
      } else if (isValid(a.titel) && isValid(b.titel)) {
        var sim = titleSimilarity(a.titel, b.titel);
        if (sim >= 0.85) { matchType = 'Titel sehr ähnlich'; confidence = Math.round(sim * 100); }
        else if (sim >= 0.70) { matchType = 'Titel ähnlich'; confidence = Math.round(sim * 100); }
      }

      if (!matchType) continue;
      seen[pairKey] = true;

      // ── AUTO-ENTSCHEIDUNG ─────────────────────────────────────
      var autoDeleteUuid = null;

      var scoreA = countValidFields(a);
      var scoreB = countValidFields(b);
      var relevanzA = getRelevanz(a.status);
      var relevanzB = getRelevanz(b.status);

      // Fall 1: Eines hat fehlende Metadaten, Relevanz gleich oder eines N/A
      // → behalte das mit mehr validen Feldern
      if (confidence === 100) {
        if (scoreA !== scoreB) {
          // Unterschiedlich viele valide Felder → mehr Infos gewinnt
          autoDeleteUuid = scoreA > scoreB ? b.uuid : a.uuid;
          Logger.log('[AUTO-DUPLIKAT] Metadaten-Score: A=' + scoreA + ' B=' + scoreB
            + ' → lösche ' + (scoreA > scoreB ? 'B' : 'A'));
        } else if (relevanzA === relevanzB) {
          // Gleiche Relevanz, gleiche Metadaten-Anzahl → A behalten, B löschen
          autoDeleteUuid = b.uuid;
          Logger.log('[AUTO-DUPLIKAT] Identisch → lösche B (' + b.uuid + ')');
        }
      }

      if (autoDeleteUuid) {
        // Zeile löschen
        var deleteUuid = autoDeleteUuid;
        var rowToDelete = -1;
        // Neu einlesen weil vorherige Löschungen Zeilennummern verschieben
        var currentData = dashboard.getRange(2, 1, dashboard.getLastRow() - 1, headers.length).getValues();
        for (var rd = 0; rd < currentData.length; rd++) {
          if (String(currentData[rd][getIdx('UUID')] || '').trim() === deleteUuid) {
            rowToDelete = rd + 2;
            break;
          }
        }
        if (rowToDelete > 0) {
          dashboard.deleteRow(rowToDelete);
          SpreadsheetApp.flush();
          deleted[deleteUuid] = true;
          autoCleaned++;
          Logger.log('[AUTO-DUPLIKAT] Gelöscht: Zeile ' + rowToDelete + ' UUID: ' + deleteUuid);
        }
      } else {
        // Manuell entscheiden
        manualPairs.push({
          a: a, b: b,
          matchType: matchType,
          confidence: confidence
        });
      }
    }
  }

  manualPairs.sort(function(x, y) { return y.confidence - x.confidence; });
  return { autoCleaned: autoCleaned, manualPairs: manualPairs };
}

// ── Titel-Ähnlichkeit (Jaccard auf Wort-Ebene) ───────────────────
function titleSimilarity(a, b) {
  var wordsA = a.replace(/[^a-z0-9\s]/g, '').split(/\s+/).filter(function(w) { return w.length > 2; });
  var wordsB = b.replace(/[^a-z0-9\s]/g, '').split(/\s+/).filter(function(w) { return w.length > 2; });

  if (wordsA.length === 0 || wordsB.length === 0) return 0;

  var setA = {};
  wordsA.forEach(function(w) { setA[w] = true; });

  var intersection = 0;
  wordsB.forEach(function(w) { if (setA[w]) intersection++; });

  var union = Object.keys(setA).length + wordsB.length - intersection;
  return union === 0 ? 0 : intersection / union;
}

// ── Duplikat-Paare ans Frontend liefern ──────────────────────────
function getDuplicatePairs() {
  var result = findAndProcessDuplicates(null);
  return result.manualPairs;
}

// ── Entscheidung verarbeiten (manuell vom Dialog) ─────────────────
function processDuplicateDecision(decision) {
  if (decision.action === 'KEEP_BOTH') return { success: true };

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) return { success: false, error: 'Dashboard nicht gefunden' };

  var headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  var lastRow = dashboard.getLastRow();
  var data = dashboard.getRange(2, 1, lastRow - 1, headers.length).getValues();

  var getIdx = function(name) {
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c]).trim() === name) return c;
    }
    return -1;
  };

  var rowNum = -1;
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][getIdx('UUID')] || '').trim() === decision.deleteUuid) {
      rowNum = r + 2;
      break;
    }
  }

  if (rowNum < 0) return { success: false, error: 'UUID nicht gefunden: ' + decision.deleteUuid };

  dashboard.deleteRow(rowNum);
  SpreadsheetApp.flush();
  Logger.log('[DUPLIKAT-MANUELL] Zeile ' + rowNum + ' gelöscht (UUID: ' + decision.deleteUuid + ')');

  return { success: true };
}
