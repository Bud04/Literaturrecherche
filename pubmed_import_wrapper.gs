// FILE: pubmed_import_wrapper.gs

function importPapersFromPubMed(query, startDate, maxResults, category, endDate) {
  var ui = SpreadsheetApp.getUi();
  if (!query || query.trim() === '') {
    ui.alert('Fehler', 'Kein Suchbegriff angegeben', ui.ButtonSet.OK);
    return;
  }
  maxResults = parseInt(maxResults) || 50;

  try {
    importFromPubMedWithDate(query, startDate, maxResults);

    var actualEndDate = PropertiesService.getDocumentProperties()
                          .getProperty('lastPubDate') ||
                        endDate ||
                        Utilities.formatDate(new Date(), 'GMT+1', 'yyyy/MM/dd');
    var effectiveEnd = (endDate && actualEndDate > endDate) ? endDate : actualEndDate;

    saveImportRange(category || 'Unbekannt', startDate || '–', effectiveEnd, maxResults);
    if (typeof logAction === 'function') {
      logAction('PubMed Import', 'Kategorie: ' + (category || 'Unbekannt') +
        ' | ' + (startDate || '–') + ' → ' + effectiveEnd);
    }
  } catch (e) {
    Logger.log('importPapersFromPubMed FEHLER: ' + e.message);
    ui.alert('Import-Fehler', e.message, ui.ButtonSet.OK);
  }
}

// ── Ersetzt importFromPubMedWithDate aus pubmed_import.gs ──────────────────
function importFromPubMedWithDate(query, startDate, maxResults) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');

  var fullQuery = query;
  if (startDate) {
    fullQuery += ' AND (""' + formatDateForPubMed(startDate) + '""[PDAT] : ""3000""[PDAT])';
  }

  var searchUrl = SMART_PUBMED.SEARCH +
    '?db=pubmed&term=' + encodeURIComponent(fullQuery) +
    '&retmax=' + maxResults +
    '&retmode=json&sort=pub+date&usehistory=y';

  var pmids = JSON.parse(UrlFetchApp.fetch(searchUrl).getContentText())
                .esearchresult.idlist || [];

  if (pmids.length === 0) {
    SpreadsheetApp.getUi().alert('Keine Papers gefunden.');
    return;
  }

  var fetchUrl = SMART_PUBMED.FETCH +
    '?db=pubmed&id=' + pmids.join(',') + '&retmode=xml';
  var papers = smartParsePubMedXML(UrlFetchApp.fetch(fetchUrl).getContentText());

  var imported = 0;
  var skipped  = 0;
  var lastDate = startDate || '';

  for (var i = 0; i < papers.length; i++) {
    var paper = papers[i];
    if (smartPaperExists(dashboard, paper.pmid)) { skipped++; continue; }

    var volltext = smartFetchFulltext(paper);
    smartSavePaper(dashboard, paper, volltext);
    imported++;
    if (paper.pubDate) lastDate = paper.pubDate;
    Utilities.sleep(SMART_PUBMED.SLEEP);
  }

  // Datum für Wrapper speichern
  if (lastDate) {
    PropertiesService.getDocumentProperties().setProperty('lastPubDate', lastDate);
  }

  Logger.log('importFromPubMedWithDate: ' + imported + ' importiert, ' + skipped + ' übersprungen');
}

// ── Datums-Formatierung (wird oben benötigt) ───────────────────────────────
function formatDateForPubMed(dateString) {
  if (!dateString) return null;
  var cleaned = dateString.replace(/[.-]/g, '/');
  var parts   = cleaned.split('/');
  if (parts.length === 3) {
    if (parts[0].length === 4) return parts[0] + '/' + parts[1].padStart(2,'0') + '/' + parts[2].padStart(2,'0');
    else                       return parts[2] + '/' + parts[1].padStart(2,'0') + '/' + parts[0].padStart(2,'0');
  }
  return dateString;
}
