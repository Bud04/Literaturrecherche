// FILE: smart_pubmed_import.gs
// VERSION 7.2 - Angepasst an echtes 69-Spalten-Dashboard (A bis BQ)
//
// FIXES vs. v7.1:
// - onOpen() entfernt (Duplikat - ist in ui.gs)
// - EMAIL: Semikolon weg, aus Script Properties
// - savePaperToDashboard: header-basiert statt hardcoded
// - ""Jahr"" -> ""Publikationsdatum"" (echte Spalte F)
// - Alle Funktionsnamen mit ""smart"" Prefix (kein Konflikt mit pubmed_import.gs)

const SMART_PUBMED = {
  SEARCH: 'https://eutils.ncbi.nlm.nih.gov/entrez/eutils/esearch.fcgi',
  FETCH:  'https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi',
  SLEEP:  334
};

function getImportEmail() {
  return PropertiesService.getScriptProperties().getProperty('IMPORT_EMAIL')
    || Session.getActiveUser().getEmail()
    || 'research@example.com';
}

// ===================================================
// MENÜ-EINSTIEG (onOpen ist in ui.gs)
// ===================================================

function showStats() {
  var ui = SpreadsheetApp.getUi();
  var dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  if (!dashboard) { ui.alert('Dashboard nicht gefunden'); return; }
  ui.alert('Statistik', 'Gesamt Papers: ' + Math.max(0, dashboard.getLastRow() - 1), ui.ButtonSet.OK);
}

function importAllTopicsSmart() {
  importNewPubMedPapers();
}

// ===================================================
// HAUPTFUNKTION
// ===================================================

function importNewPubMedPapers() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');

  if (!dashboard) {
    ui.alert('Fehler', 'Dashboard nicht gefunden!', ui.ButtonSet.OK);
    return;
  }

  var query =
    '((continuous glucose monitoring OR CGM OR insulin pump OR CSII OR ' +
    'continuous subcutaneous insulin infusion OR diabetes technology OR diabetes device OR ' +
    'automated insulin delivery OR AID OR closed loop OR hybrid closed loop OR ' +
    'flash glucose monitoring OR FGM OR real-time CGM OR rtCGM OR sensor-augmented pump OR SAP) ' +
    'AND (diabetes OR diabetic OR type 1 diabetes OR type 2 diabetes OR T1D OR T2D OR T1DM OR T2DM OR ' +
    'diabetes mellitus OR insulin dependent diabetes OR NIDDM OR insulin-requiring diabetes OR insulin therapy)) ' +
    'OR ' +
    '(primary immunodeficiency OR primary immunodeficiencies OR PID OR PIDD OR ' +
    'secondary immunodeficiency OR secondary immunodeficiencies OR acquired immunodeficiency OR ' +
    'immunoglobulin OR immunoglobulins OR IVIG OR SCIG OR subcutaneous immunoglobulin OR ' +
    'IgG replacement OR immunoglobulin therapy OR immunoglobulin treatment)';

  var defaultDate = smartGetLastDate(dashboard);
  var dateResponse = ui.prompt(
    'Smart Import – Schritt 1/3: Startdatum',
    'Ab welchem Datum importieren?\n\n' +
    'Format: YYYY/MM/DD (z.B. 2025/01/01)\n' +
    'Leer lassen = letzter Import (' + defaultDate + ')',
    ui.ButtonSet.OK_CANCEL
  );
  if (dateResponse.getSelectedButton() !== ui.Button.OK) return;

  var startDateInput = dateResponse.getResponseText().trim();
  var startDate = startDateInput === '' ? defaultDate : startDateInput;

  var countResponse = ui.prompt(
    'Smart Import – Schritt 2/3: Anzahl',
    'Wie viele Papers importieren?\n\n' +
    'Leer lassen = 100\n' +
    'Oder Zahl eingeben (1 – 500)',
    ui.ButtonSet.OK_CANCEL
  );
  if (countResponse.getSelectedButton() !== ui.Button.OK) return;

  var countInput = countResponse.getResponseText().trim();
  var targetCount = 100;
  if (countInput !== '') {
    var parsed = parseInt(countInput);
    if (isNaN(parsed) || parsed < 1 || parsed > 500) {
      ui.alert('Fehler', 'Bitte Zahl zwischen 1 und 500 eingeben!', ui.ButtonSet.OK);
      return;
    }
    targetCount = parsed;
  }

  var confirm = ui.alert(
    'Smart Import – Schritt 3/3: Bestätigung',
    '=== SMART IMPORT – BESTÄTIGUNG ===\n\n' +
    'Startdatum: ' + startDate + '\n' +
    'Ziel: ' + targetCount + ' Papers\n' +
    'Methode: Tag für Tag ab Startdatum\n\n' +
    'Jetzt starten?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  try {
    var parts = startDate.split('/');
    var currentDate = new Date(
      parseInt(parts[0]),
      parseInt(parts[1]) - 1,
      parseInt(parts[2]),
      12, 0, 0
    );

    var today = new Date();
    var imported = 0;
    var skipped = 0;
    var daysChecked = 0;
    var lastImportedDate = startDate;
    var dayFullyDone = false;
    var MAX_DAYS = 400;
    var newUuids = []; // ← NEU: Duplikat-Tracking

    while (imported < targetCount && currentDate <= today && daysChecked < MAX_DAYS) {
      var dateStr = Utilities.formatDate(currentDate, 'GMT+1', 'yyyy/MM/dd');
      daysChecked++;
      dayFullyDone = false;

      Logger.log('Prüfe Tag: ' + dateStr);

      var searchUrl = SMART_PUBMED.SEARCH +
        '?db=pubmed&term=' + encodeURIComponent(query) +
        '&retmax=100' +
        '&retmode=json' +
        '&datetype=pdat' +
        '&mindate=' + dateStr +
        '&maxdate=' + dateStr;

      var searchResp = UrlFetchApp.fetch(searchUrl, { muteHttpExceptions: true });
      if (searchResp.getResponseCode() !== 200) {
        Logger.log('API Fehler am ' + dateStr);
        currentDate.setDate(currentDate.getDate() + 1);
        Utilities.sleep(SMART_PUBMED.SLEEP);
        continue;
      }

      var pmids = JSON.parse(searchResp.getContentText()).esearchresult.idlist || [];

      if (pmids.length === 0) {
        Logger.log('Keine Papers am ' + dateStr);
        dayFullyDone = true;
        currentDate.setDate(currentDate.getDate() + 1);
        Utilities.sleep(SMART_PUBMED.SLEEP);
        continue;
      }

      Logger.log(pmids.length + ' PMIDs am ' + dateStr);
      Utilities.sleep(SMART_PUBMED.SLEEP);

      var fetchUrl = SMART_PUBMED.FETCH +
        '?db=pubmed&id=' + pmids.join(',') + '&retmode=xml';
      var papers = smartParsePubMedXML(
        UrlFetchApp.fetch(fetchUrl).getContentText()
      );

      var reachedLimit = false;
      for (var p = 0; p < papers.length; p++) {
        if (imported >= targetCount) {
          reachedLimit = true;
          break;
        }

        var paper = papers[p];
        if (smartPaperExists(dashboard, paper.pmid)) {
          skipped++;
          continue;
        }

        var volltext = smartFetchFulltext(paper);
        var savedUuid = smartSavePaper(dashboard, paper, volltext); // ← gibt jetzt UUID zurück
        if (savedUuid) newUuids.push(savedUuid); // ← NEU: UUID merken
        imported++;
        Logger.log('Importiert (' + imported + '/' + targetCount + '): ' + paper.pmid);
        Utilities.sleep(SMART_PUBMED.SLEEP);
      }

      lastImportedDate = dateStr;

      if (reachedLimit) {
        dayFullyDone = false;
        break;
      } else {
        dayFullyDone = true;
        currentDate.setDate(currentDate.getDate() + 1);
        Utilities.sleep(SMART_PUBMED.SLEEP);
      }
    }

    var saveDate;
    if (dayFullyDone) {
      var saveParts = lastImportedDate.split('/');
      var saveDateObj = new Date(
        parseInt(saveParts[0]),
        parseInt(saveParts[1]) - 1,
        parseInt(saveParts[2]),
        12, 0, 0
      );
      saveDateObj.setDate(saveDateObj.getDate() + 1);
      saveDate = Utilities.formatDate(saveDateObj, 'GMT+1', 'yyyy/MM/dd');
    } else {
      saveDate = lastImportedDate;
    }

    PropertiesService.getDocumentProperties().setProperty('lastPubDate', saveDate);

    var endMsg = '=== IMPORT ABGESCHLOSSEN ===\n\n';
    endMsg += '✅ Importiert: ' + imported + '\n';
    endMsg += '⏭️ Übersprungen: ' + skipped + '\n';
    endMsg += '📅 Tage geprüft: ' + daysChecked + '\n';
    endMsg += '📅 Nächster Import ab: ' + saveDate + '\n\n';
    endMsg += dayFullyDone
      ? '✅ Tag vollständig – nächster Import startet am Folgetag.'
      : '⚠️ Limit mitten im Tag erreicht – nächster Import setzt am gleichen Tag fort.';

    ui.alert('Import fertig', endMsg, ui.ButtonSet.OK);

    // ── DUPLIKAT-CHECK ── (nach dem Alert, damit Import-Ergebnis zuerst sichtbar)
    if (newUuids.length > 0) {
      checkDuplicatesAfterImport(newUuids);
    }

  } catch (e) {
    Logger.log('FEHLER: ' + e.message);
    ui.alert('Fehler', e.message, ui.ButtonSet.OK);
  }
}

// ===================================================
// DATUM-HELPER
// ===================================================

function smartGetLastDate(dashboard) {
  var props = PropertiesService.getDocumentProperties();
  var saved = props.getProperty('lastPubDate');
  if (saved) return saved;

  var lastRow = dashboard.getLastRow();
  if (lastRow < 2) return '2020/01/01';

  // Lese ""Publikationsdatum"" (Spalte F = Index 6)
  var headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  var col = headers.indexOf('Publikationsdatum') + 1;
  if (col === 0) col = headers.indexOf('Jahr') + 1; // Fallback
  if (col === 0) return '2020/01/01';

  var dates = dashboard.getRange(2, col, lastRow - 1, 1).getValues().flat()
    .filter(function(d) { return d; })
    .map(function(d) { return String(d).substring(0, 4); }) // Nur Jahr extrahieren
    .filter(function(y) { return /^\d{4}$/.test(y); })
    .sort();

  if (dates.length === 0) return '2020/01/01';

  var formatted = dates[dates.length - 1] + '/12/31';
  props.setProperty('lastPubDate', formatted);
  return formatted;
}

function smartPaperExists(dashboard, pmid) {
  var lastRow = dashboard.getLastRow();
  if (lastRow < 2 || !pmid) return false;

  var headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  var col = headers.indexOf('PMID') + 1;
  if (col === 0) return false;

  var pmids = dashboard.getRange(2, col, lastRow - 1, 1).getValues().flat();
  return pmids.indexOf(pmid) >= 0;
}

// ===================================================
// XML PARSING
// ===================================================

function smartParsePubMedXML(xmlText) {
  var papers = [];
  var articles = xmlText.match(/<PubmedArticle>[\s\S]*?<\/PubmedArticle>/g) || [];

  for (var i = 0; i < articles.length; i++) {
    try {
      papers.push(smartParseArticle(articles[i]));
    } catch (e) {
      Logger.log('Parse Error: ' + e.message);
    }
  }
  return papers;
}

function smartParseArticle(xml) {
  return {
    pmid:     smartRx(xml, /<PMID[^>]*>(\d+)<\/PMID>/),
    doi:      smartRx(xml, /<ArticleId IdType=""doi"">([^<]+)<\/ArticleId>/),
    pmcid:    smartRx(xml, /<ArticleId IdType=""pmc"">([^<]+)<\/ArticleId>/),
    title:    smartRx(xml, /<ArticleTitle>([^<]+)<\/ArticleTitle>/),
    abstract: smartParseAbstract(xml),
    authors:  smartParseAuthors(xml),
    journal:  smartRx(xml, /<Title>([^<]+)<\/Title>/),
    pubDate:  smartParseDate(xml),
    year:     smartRx(xml, /<Year>(\d{4})<\/Year>/)
  };
}

function smartRx(text, pattern) {
  var m = text.match(pattern);
  return m ? m[1].trim() : '';
}

function smartParseAbstract(xml) {
  var m = xml.match(/<Abstract>([\s\S]*?)<\/Abstract>/);
  if (!m) return '';
  var parts = m[1].match(/<AbstractText[^>]*>([^<]*)<\/AbstractText>/g) || [];
  return parts.map(function(t) {
    var r = t.match(/<AbstractText[^>]*>([^<]*)<\/AbstractText>/);
    return r ? r[1].trim() : '';
  }).join(' ');
}

function smartParseAuthors(xml) {
  var matches = xml.match(/<Author[^>]*>([\s\S]*?)<\/Author>/g) || [];
  var names = matches.map(function(a) {
    var last = smartRx(a, /<LastName>([^<]+)<\/LastName>/);
    var fore = smartRx(a, /<ForeName>([^<]+)<\/ForeName>/);
    return last && fore ? last + ', ' + fore : last || null;
  }).filter(function(n) { return n; });
  return names.slice(0, 3).join('; ') + (names.length > 3 ? ' et al.' : '');
}

function smartParseDate(xml) {
  // Strategie 1: ArticleDate → echtes Publikationsdatum (wie auf Website)
  var ad = xml.match(/<ArticleDate[^>]*>([\s\S]*?)<\/ArticleDate>/);
  if (ad) {
    var y = smartRx(ad[1], /<Year>(\d{4})<\/Year>/);
    var mo = smartRx(ad[1], /<Month>(\d{1,2})<\/Month>/);
    var d = smartRx(ad[1], /<Day>(\d{1,2})<\/Day>/);
    if (y && mo && d) return y + '-' + mo.padStart(2,'0') + '-' + d.padStart(2,'0');
  }

  // Strategie 2: PubDate im Journal
  var pd = xml.match(/<PubDate>([\s\S]*?)<\/PubDate>/);
  if (pd) {
    var y2 = smartRx(pd[1], /<Year>(\d{4})<\/Year>/);
    var rawM = smartRx(pd[1], /<Month>(\d{1,2}|[A-Za-z]+)<\/Month>/);
    var d2 = smartRx(pd[1], /<Day>(\d{1,2})<\/Day>/);
    if (y2 && rawM) {
      var nums = {'Jan':'01','Feb':'02','Mar':'03','Apr':'04','May':'05','Jun':'06',
                  'Jul':'07','Aug':'08','Sep':'09','Oct':'10','Nov':'11','Dec':'12'};
      var mNum = /^\d+$/.test(rawM) ? rawM.padStart(2,'0') : (nums[rawM] || '01');
      var dayStr = d2 ? d2.padStart(2,'0') : '01';
      return y2 + '-' + mNum + '-' + dayStr;
    }
  }

  // Strategie 3: PubMedPubDate (Indexierungsdatum, nur Fallback)
  var blocks = xml.match(/<PubMedPubDate[^>]*>[\s\S]*?<\/PubMedPubDate>/g) || [];
  for (var i = 0; i < blocks.length; i++) {
    var b = blocks[i];
    var st = b.match(/PubStatus=""([^""]+)""/);
    if (st && (st[1] === 'pubmed' || st[1] === 'medline')) {
      var y3 = smartRx(b, /<Year>(\d{4})<\/Year>/);
      var mo3 = smartRx(b, /<Month>(\d{1,2})<\/Month>/);
      var d3 = smartRx(b, /<Day>(\d{1,2})<\/Day>/);
      if (y3 && mo3 && d3) return y3 + '-' + mo3.padStart(2,'0') + '-' + d3.padStart(2,'0');
    }
  }
  return '';
}

// ===================================================
// VOLLTEXT
// ===================================================


function smartCleanText(text) {
  if (!text) return '';
  return text
    .replace(/[^\x20-\x7E\n\r\t]/g, '')
    .replace(/\n\s*References\s*\n[\s\S]*$/i, '')
    .replace(/\n\s*REFERENCES\s*\n[\s\S]*$/i, '')
    .replace(/\n\s*Figure\s+\d+[:\.]?[^\n]*\n/gi, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .replace(/[ \t]+/g, ' ')
    .trim();
}

function smartParseAbstract(xml) {
  var m = xml.match(/<Abstract>([\s\S]*?)<\/Abstract>/);
  if (!m) return '';
  // ✅ [\s\S]*? statt [^<]* → matcht auch Tags innerhalb AbstractText
  var parts = m[1].match(/<AbstractText[^>]*>([\s\S]*?)<\/AbstractText>/g) || [];
  return parts.map(function(t) {
    var r = t.match(/<AbstractText[^>]*>([\s\S]*?)<\/AbstractText>/);
    if (!r) return '';
    // Tags innerhalb entfernen
    return r[1].replace(/<[^>]+>/g, '').trim();
  }).filter(function(s) { return s; }).join(' ');
}

function extractTextFromPdfUrl(pdfUrl) {
  var pdfFile = null;
  var docFile = null;

  try {
    Logger.log('PDF Extraktion: ' + pdfUrl);
    var response = UrlFetchApp.fetch(pdfUrl, { muteHttpExceptions: true, followRedirects: true });
    if (response.getResponseCode() !== 200) {
      Logger.log('PDF Download Fehler: HTTP ' + response.getResponseCode());
      return '';
    }

    var pdfBlob = response.getBlob()
      .setContentType('application/pdf')
      .setName('temp_pdf_' + new Date().getTime() + '.pdf');

    // ── Schritt 1: PDF in Drive hochladen (DriveApp, kein Advanced Service nötig)
    pdfFile = DriveApp.createFile(pdfBlob);
    var pdfFileId = pdfFile.getId();
    Logger.log('PDF hochgeladen: ' + pdfFileId);

    // ── Schritt 2: PDF als Google Doc kopieren (konvertiert automatisch)
    var copyResponse = UrlFetchApp.fetch(
      'https://www.googleapis.com/drive/v3/files/' + pdfFileId + '/copy',
      {
        method: 'POST',
        headers: {
          'Authorization': 'Bearer ' + ScriptApp.getOAuthToken(),
          'Content-Type': 'application/json'
        },
        payload: JSON.stringify({
          mimeType: 'application/vnd.google-apps.document',
          name: 'temp_doc_' + new Date().getTime()
        }),
        muteHttpExceptions: true
      }
    );

    // PDF sofort löschen
    pdfFile.setTrashed(true);
    pdfFile = null;

    if (copyResponse.getResponseCode() !== 200) {
      Logger.log('Doc-Konvertierung Fehler: ' + copyResponse.getContentText());
      return '';
    }

    var docId = JSON.parse(copyResponse.getContentText()).id;
    Logger.log('Google Doc erstellt: ' + docId);

    // Kurz warten bis Konvertierung fertig
    Utilities.sleep(3000);

    // ── Schritt 3: Text aus Google Doc lesen
    var doc = DocumentApp.openById(docId);
    var text = doc.getBody().getText();

    // Google Doc löschen
    DriveApp.getFileById(docId).setTrashed(true);

    if (text && text.length > 100) {
      Logger.log('PDF Text extrahiert: ' + text.length + ' Zeichen');
      return text;
    }

    Logger.log('PDF Text zu kurz oder leer');
    return '';

  } catch(e) {
    Logger.log('PDF Extraktion Fehler: ' + e.message);
    // Aufräumen falls Fehler
    try { if (pdfFile) pdfFile.setTrashed(true); } catch(e2) {}
    return '';
  }
}

function smartFetchFulltext(paper) {
  var volltext = '';

  // ── 1. PMC Fulltext via API ────────────────────────────────────
  if (paper.pmcid) {
    try {
      var pmcUrl = 'https://www.ncbi.nlm.nih.gov/research/bionlp/RESTful/pmcoa.cgi/BioC_json/' + paper.pmcid + '/unicode';
      var pmcResponse = UrlFetchApp.fetch(pmcUrl, { muteHttpExceptions: true });
      if (pmcResponse.getResponseCode() === 200) {
        var pmcText = pmcResponse.getContentText();
        if (pmcText && !pmcText.trim().startsWith('%PDF-') && pmcText.length > 500) {
          try {
            var pmcJson = JSON.parse(pmcText);
            var texts = [];
            if (pmcJson.documents) {
              pmcJson.documents.forEach(function(doc) {
                if (doc.passages) {
                  doc.passages.forEach(function(passage) {
                    if (passage.text) texts.push(passage.text);
                  });
                }
              });
            }
            if (texts.length > 0) {
              volltext = texts.join('\n\n');
              Logger.log('PMC API Volltext: ' + volltext.length + ' Zeichen');
              return volltext;
            }
          } catch(e) {
            Logger.log('PMC JSON Parse Fehler: ' + e.message);
          }
        }
      }
    } catch(e) {
      Logger.log('PMC API Fehler: ' + e.message);
    }
  }

  // ── 2. Unpaywall ──────────────────────────────────────────────
  if (paper.doi) {
    try {
      var unpayUrl = 'https://api.unpaywall.org/v2/' + encodeURIComponent(paper.doi) + '?email=vitalaire@research.de';
      var unpayResponse = UrlFetchApp.fetch(unpayUrl, { muteHttpExceptions: true });
      if (unpayResponse.getResponseCode() === 200) {
        var unpayJson = JSON.parse(unpayResponse.getContentText());
        var pdfUrl = null;

        if (unpayJson.best_oa_location && unpayJson.best_oa_location.url_for_pdf) {
          pdfUrl = unpayJson.best_oa_location.url_for_pdf;
        } else if (unpayJson.best_oa_location && unpayJson.best_oa_location.url_for_landing_page) {
          pdfUrl = unpayJson.best_oa_location.url_for_landing_page;
        } else if (unpayJson.oa_locations && unpayJson.oa_locations.length > 0) {
          for (var i = 0; i < unpayJson.oa_locations.length; i++) {
            if (unpayJson.oa_locations[i].url_for_pdf) {
              pdfUrl = unpayJson.oa_locations[i].url_for_pdf;
              break;
            }
          }
        }

        if (pdfUrl) {
          Logger.log('Unpaywall URL: ' + pdfUrl);
          var pageResponse = UrlFetchApp.fetch(pdfUrl, { muteHttpExceptions: true, followRedirects: true });
          var pageContent = pageResponse.getContentText();

          if (pageContent && pageContent.trim().startsWith('%PDF-')) {
            // ✅ Raw PDF → via Drive extrahieren
            Logger.log('Raw PDF von Unpaywall → Drive Extraktion');
            var extracted = extractTextFromPdfUrl(pdfUrl);
            if (extracted.length > 100) return extracted;
          } else if (pageContent && pageContent.length > 500) {
            volltext = pageContent
              .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
              .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
              .replace(/<[^>]+>/g, ' ')
              .replace(/\s+/g, ' ')
              .trim();
            if (volltext.length > 500) {
              Logger.log('Unpaywall HTML Volltext: ' + volltext.length + ' Zeichen');
              return volltext;
            }
          }
        }
      }
    } catch(e) {
      Logger.log('Unpaywall Fehler: ' + e.message);
    }
  }

  // ── 3. PubMed Central HTML direkt ─────────────────────────────
  if (paper.pmcid) {
    try {
      var htmlUrl = 'https://www.ncbi.nlm.nih.gov/pmc/articles/' + paper.pmcid + '/';
      var htmlResponse = UrlFetchApp.fetch(htmlUrl, { muteHttpExceptions: true });
      if (htmlResponse.getResponseCode() === 200) {
        var htmlContent = htmlResponse.getContentText();

        if (htmlContent && htmlContent.trim().startsWith('%PDF-')) {
          // ✅ Raw PDF → via Drive extrahieren
          Logger.log('Raw PDF von PMC HTML → Drive Extraktion');
          var extracted = extractTextFromPdfUrl(htmlUrl);
          if (extracted.length > 100) return extracted;
        } else if (htmlContent) {
          var articleMatch = htmlContent.match(/<div[^>]*class=""""[^""""]*article[^""""]*""""[^>]*>([\s\S]*?)<\/div>/i);
          var bodyText = articleMatch ? articleMatch[1] : htmlContent;

          volltext = bodyText
            .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, '')
            .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, '')
            .replace(/<[^>]+>/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();

          if (volltext.length > 500) {
            Logger.log('PMC HTML Volltext: ' + volltext.length + ' Zeichen');
            return volltext;
          }
        }
      }
    } catch(e) {
      Logger.log('PMC HTML Fehler: ' + e.message);
    }
  }

  // ── 4. Europe PMC als Fallback ─────────────────────────────────
  if (paper.pmid) {
    try {
      var euroUrl = 'https://www.ebi.ac.uk/europepmc/webservices/rest/' + paper.pmid + '/fullTextXML';
      var euroResponse = UrlFetchApp.fetch(euroUrl, { muteHttpExceptions: true });
      if (euroResponse.getResponseCode() === 200) {
        var euroText = euroResponse.getContentText();

        if (euroText && euroText.trim().startsWith('%PDF-')) {
          // ✅ Raw PDF → via Drive extrahieren
          Logger.log('Raw PDF von Europe PMC → Drive Extraktion');
          var extracted = extractTextFromPdfUrl(euroUrl);
          if (extracted.length > 100) return extracted;
        } else if (euroText && euroText.length > 500) {
          volltext = euroText
            .replace(/<[^>]+>/g, ' ')
            .replace(/\s+/g, ' ')
            .trim();
          if (volltext.length > 500) {
            Logger.log('Europe PMC Volltext: ' + volltext.length + ' Zeichen');
            return volltext;
          }
        }
      }
    } catch(e) {
      Logger.log('Europe PMC Fehler: ' + e.message);
    }
  }

  Logger.log('Kein Volltext gefunden für PMID: ' + paper.pmid);
  return '';
}

// ─── PMC HTML (Fallback) ──────────────────────────────────────────
function smartFetchPMCHtml(pmcid) {
  var id = pmcid.replace('PMC', '');
  var resp = UrlFetchApp.fetch('https://pmc.ncbi.nlm.nih.gov/articles/PMC' + id + '/', {
    muteHttpExceptions: true,
    headers: { 'User-Agent': 'Mozilla/5.0' }
  });
  if (resp.getResponseCode() !== 200) return '';

  var html = resp.getContentText();
  var content = '';

  // Mehrere Selektoren versuchen
  var m = html.match(/<div[^>]*class=""""[^""""]*article[^""""]*body[^""""]*""""[^>]*>([\s\S]*?)<\/div>\s*(?=<div[^>]*class=""""[^""""]*(?:back|ref|foot)[^""""]*"""")/i);
  if (m) content = m[1];

  if (!content) {
    m = html.match(/<div[^>]*id=""""[^""""]*body[^""""]*""""[^>]*>([\s\S]*?)<\/div>\s*(?=<div[^>]*id=""""[^""""]*(?:back|ref)[^""""]*"""")/i);
    if (m) content = m[1];
  }

  if (!content) {
    m = html.match(/<div[^>]*class=""""[^""""]*main-article-body[^""""]*""""[^>]*>([\s\S]*?)<\/div>\s*<\/div>/i);
    if (m) content = m[1];
  }

  if (!content) return '';

  return content
    .replace(/<section[^>]*id=""""[^""""]*(?:ref|bib|ack)[^""""]*""""[^>]*>[\s\S]*?<\/section>/gi, '')
    .replace(/<figure[^>]*>[\s\S]*?<\/figure>/gi, '')
    .replace(/<math[^>]*>[\s\S]*?<\/math>/gi, '[FORMEL]')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&nbsp;/g, ' ').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/[ \t]+/g, ' ')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

// ─── UNPAYWALL PDF ────────────────────────────────────────────────
function smartFetchUnpaywall(doi) {
  var url = 'https://api.unpaywall.org/v2/' + doi + '?email=' + getImportEmail();
  Utilities.sleep(SMART_PUBMED.SLEEP);
  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() !== 200) return '';

  var data = JSON.parse(resp.getContentText());
  if (!data.is_oa) return '';

  var loc = data.best_oa_location || (data.oa_locations && data.oa_locations[0]);
  if (!loc || !loc.url_for_pdf) return '';

  Utilities.sleep(SMART_PUBMED.SLEEP);
  var pdfResp = UrlFetchApp.fetch(loc.url_for_pdf, { muteHttpExceptions: true });
  if (pdfResp.getResponseCode() !== 200) return '';

  return smartCleanText(pdfResp.getBlob().getDataAsString());
}

// ===================================================
// SAVE TO DASHBOARD
// Header-basiert - passt automatisch zum echten Sheet
// ===================================================

function repairBrokenVolltexte() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) return;

  var headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  var lastRow = dashboard.getLastRow();
  if (lastRow < 2) return;

  var data = dashboard.getRange(2, 1, lastRow - 1, headers.length).getValues();

  var getIdx = function(name) {
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c]).trim() === name) return c;
    }
    return -1;
  };

  var geminiCols = [
    'Gemini Researcher', 'Gemini Triage', 'Gemini Metadaten',
    'Gemini Master Analyse', 'Gemini Redaktion', 'Gemini Faktencheck', 'Gemini Review'
  ];

  var triggerCols = [
    'Flow_Trigger_Researcher', 'Flow_Trigger_Triage', 'Flow_Trigger_Metadaten',
    'Flow_Trigger_Analyse', 'Flow_Trigger_Redaktion', 'Flow_Trigger_Faktencheck',
    'Flow_Trigger_Review'
  ];

  var resetCols = [
    'Relevanz', 'Hauptkategorie', 'Unterkategorien', 'Haupterkenntnis',
    'Zusammenfassung', 'Kritische Bewertung', 'Produkt-Fokus',
    'Review-Status', 'Relevanz-Begründung', 'Praktische Implikationen',
    'Fehler-Details'
  ];

  var MAX_CHARS = 50000;

  var splitVolltext = function(text) {
    return {
      teil1: text.substring(0, MAX_CHARS),
      teil2: text.length > MAX_CHARS ? text.substring(MAX_CHARS, MAX_CHARS * 2) : '',
      teil3: text.length > MAX_CHARS * 2 ? text.substring(MAX_CHARS * 2, MAX_CHARS * 3) : ''
    };
  };

  var volltextStatusIdx = getIdx('Volltext-Status');
  if (volltextStatusIdx < 0) {
    Logger.log('⚠️ Spalte """"Volltext-Status"""" nicht gefunden!');
    SpreadsheetApp.getUi().alert('Spalte """"Volltext-Status"""" nicht gefunden! Bitte Spaltenname prüfen.');
    return;
  }

  var repaired = 0;
  var failed   = 0;
  var skipped  = 0;
  var cleaned  = 0;

  for (var r = 0; r < data.length; r++) {
    var row    = data[r];
    var rowNum = r + 2;

    var uuid           = String(row[getIdx('UUID')]             || '').trim();
    var volltext       = String(row[getIdx('Volltext/Extrakt')] || '').trim();
    var volltext2      = String(row[getIdx('Volltext_Teil2')]   || '').trim();
    var volltext3      = String(row[getIdx('Volltext_Teil3')]   || '').trim();
    var volltextStatus = String(row[volltextStatusIdx]          || '').trim();
    var pmid           = String(row[getIdx('PMID')]             || '').trim();
    var pmcid          = String(row[getIdx('PMCID')]            || '').trim();
    var doi            = String(row[getIdx('DOI')]              || '').trim();
    var titel          = String(row[getIdx('Titel')]            || '').trim();

    if (!uuid) continue;

    // ── """"Kein Volltext"""" → Volltextzellen bereinigen falls Inhalt vorhanden ──
    var isKeinVolltext = volltextStatus.toLowerCase().indexOf('kein volltext') >= 0;
    if (isKeinVolltext) {
      var hadContent = volltext !== '' || volltext2 !== '' || volltext3 !== '';
      if (hadContent) {
        var v1Idx = getIdx('Volltext/Extrakt');
        var v2Idx = getIdx('Volltext_Teil2');
        var v3Idx = getIdx('Volltext_Teil3');
        if (v1Idx >= 0 && volltext  !== '') dashboard.getRange(rowNum, v1Idx + 1).clearContent();
        if (v2Idx >= 0 && volltext2 !== '') dashboard.getRange(rowNum, v2Idx + 1).clearContent();
        if (v3Idx >= 0 && volltext3 !== '') dashboard.getRange(rowNum, v3Idx + 1).clearContent();
        Logger.log('[CLEANUP] ' + uuid + ' | Kein Volltext → Zellen geleert');
        cleaned++;
        SpreadsheetApp.flush();
      } else {
        skipped++;
      }
      continue;
    }

    // ── Nur """"Volltext vorhanden"""" weiterverarbeiten ──────────────────────────
    var hasVolltext = volltextStatus === 'Volltext vorhanden';
    if (!hasVolltext) {
      skipped++;
      continue;
    }

    // ── Kaputte Zustände erkennen ───────────────────────────────────────────
    var isBroken = volltext.startsWith('%PDF') ||
                   volltext.indexOf('Volltext nicht abrufbar') >= 0 ||
                   volltext === '';

    if (!isBroken) {
      skipped++;
      continue;
    }

    Logger.log('[REPAIR-VOLLTEXT] ' + uuid + ' | ' + titel.substring(0, 60));
    Logger.log('[REPAIR-VOLLTEXT] Status: """"' + volltextStatus + '"""" | Grund: ' + (
      volltext.startsWith('%PDF') ? 'Raw PDF' :
      volltext === '' ? 'Leer' : 'Platzhalter'
    ));

    var paper = { pmid: pmid, pmcid: pmcid, doi: doi };
    var neuerVolltext = '';

    try {
      neuerVolltext = smartFetchFulltext(paper);
    } catch(e) {
      Logger.log('[REPAIR-VOLLTEXT] Fetch Fehler: ' + e.message);
      neuerVolltext = '';
    }

    if (neuerVolltext && neuerVolltext.length > 100 && !neuerVolltext.startsWith('%PDF')) {
      // ✅ Volltext auf 3 Spalten aufteilen
      var teile = splitVolltext(neuerVolltext);

      dashboard.getRange(rowNum, getIdx('Volltext/Extrakt') + 1).setValue(teile.teil1);

      var t2Idx = getIdx('Volltext_Teil2');
      if (t2Idx >= 0) dashboard.getRange(rowNum, t2Idx + 1).setValue(teile.teil2);

      var t3Idx = getIdx('Volltext_Teil3');
      if (t3Idx >= 0) dashboard.getRange(rowNum, t3Idx + 1).setValue(teile.teil3);

      Logger.log('[REPAIR-VOLLTEXT] Aufgeteilt: ' +
        teile.teil1.length + ' / ' + teile.teil2.length + ' / ' + teile.teil3.length + ' Zeichen');

      // Alle Gemini-Antworten löschen
      for (var g = 0; g < geminiCols.length; g++) {
        var gIdx = getIdx(geminiCols[g]);
        if (gIdx >= 0) dashboard.getRange(rowNum, gIdx + 1).clearContent();
      }

      // Alle Trigger löschen
      for (var t = 0; t < triggerCols.length; t++) {
        var tIdx = getIdx(triggerCols[t]);
        if (tIdx >= 0) dashboard.getRange(rowNum, tIdx + 1).clearContent();
      }

      // Abhängige Felder zurücksetzen
      for (var rc = 0; rc < resetCols.length; rc++) {
        var rcIdx = getIdx(resetCols[rc]);
        if (rcIdx >= 0) dashboard.getRange(rowNum, rcIdx + 1).clearContent();
      }

      // Status + Farbe zurücksetzen
      dashboard.getRange(rowNum, getIdx('Status') + 1)
        .setValue('🔄 Volltext repariert – Workflow neu');
      dashboard.getRange(rowNum, 1, 1, dashboard.getLastColumn())
        .setBackground('#d9ead3');

      // Researcher PENDING → Kaskade startet neu
      dashboard.getRange(rowNum, getIdx('Flow_Trigger_Researcher') + 1).setValue('PENDING');

      Logger.log('[REPAIR-VOLLTEXT] ✅ Repariert: ' + uuid + ' (' + neuerVolltext.length + ' Zeichen gesamt)');
      repaired++;

    } else {
      // ❌ Kein brauchbarer Volltext — alle drei Teile leeren
      dashboard.getRange(rowNum, getIdx('Volltext/Extrakt') + 1)
        .setValue('[Volltext nicht abrufbar – nur Abstract verfügbar]');

      var t2Idx = getIdx('Volltext_Teil2');
      if (t2Idx >= 0) dashboard.getRange(rowNum, t2Idx + 1).clearContent();

      var t3Idx = getIdx('Volltext_Teil3');
      if (t3Idx >= 0) dashboard.getRange(rowNum, t3Idx + 1).clearContent();

      dashboard.getRange(rowNum, getIdx('Status') + 1)
        .setValue('⚠️ Volltext nicht abrufbar – kein PENDING');
      dashboard.getRange(rowNum, 1, 1, dashboard.getLastColumn())
        .setBackground('#f4cccc');

      Logger.log('[REPAIR-VOLLTEXT] ❌ Fehlgeschlagen: ' + uuid);
      failed++;
    }

    SpreadsheetApp.flush();
    Utilities.sleep(1000);
  }

  var msg = 'Volltext-Reparatur abgeschlossen:\n\n' +
            '✅ Repariert: '                   + repaired + '\n' +
            '🧹 Bereinigt (Kein Volltext): '   + cleaned  + '\n' +
            '❌ Fehlgeschlagen: '               + failed   + '\n' +
            '⏭️ Übersprungen: '                + skipped;

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}

// ── Bereinigt Volltext und entfernt bei Bedarf Methodenteil ────────────────
// Gibt Objekt zurück: { text: bereinigter Text, length: Länge, methodsRemoved: bool }
var VOLLTEXT_MAX_CHARS = 60000; // ~15.000 Tokens – sicher für Gemini Flow

function cleanVolltextForFlow(text) {
  if (!text || text.length === 0) return { text: '', length: 0, methodsRemoved: false };

  var lines = text.split('\n');
  var result = [];
  var skip = false;

  var skipHeaders = [
    'references', 'bibliography', 'literatur', 'literaturverzeichnis', 'referenzen',
    'acknowledgements', 'acknowledgments', 'danksagung',
    'funding', 'financial support', 'financial disclosure',
    'conflict of interest', 'conflicts of interest', 'competing interests',
    'competing interest', 'disclosure', 'disclosures',
    'interessenkonflikt', 'interessenskonflikte',
    'author contributions', 'authors contributions',
    'ethics statement', 'ethics approval', 'ethical approval',
    'institutional review board', 'data availability',
    'availability of data', 'supplementary', 'supporting information',
    'abbreviations', 'abkürzungen', 'appendix', 'anhang'
  ];

  var resumeHeaders = [
    'abstract', 'background', 'introduction', 'methods', 'results',
    'discussion', 'conclusion', 'zusammenfassung', 'hintergrund',
    'einleitung', 'methoden', 'ergebnisse', 'diskussion', 'schlussfolgerung'
  ];

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i];
    var trimmed = line.trim();
    var lineLower = trimmed.toLowerCase();

    if (trimmed.length > 0 && trimmed.length < 60) {
      var isSkip = false;
      for (var s = 0; s < skipHeaders.length; s++) {
        if (lineLower === skipHeaders[s] ||
            lineLower === skipHeaders[s] + ':' ||
            lineLower.indexOf(skipHeaders[s] + ':') === 0) {
          isSkip = true;
          break;
        }
      }
      if (isSkip) { skip = true; continue; }

      if (skip) {
        for (var r = 0; r < resumeHeaders.length; r++) {
          if (lineLower === resumeHeaders[r] ||
              lineLower === resumeHeaders[r] + ':' ||
              lineLower.indexOf(resumeHeaders[r] + ':') === 0) {
            skip = false;
            break;
          }
        }
      }
    }

    if (!skip) result.push(line);
  }

  var cleaned = result.join('\n').trim();
  var methodsRemoved = false;

  // ── Schritt 2: Methodenteil entfernen wenn immer noch zu lang ───────────
  if (cleaned.length > VOLLTEXT_MAX_CHARS) {
    var methodsSkipHeaders = [
      'methods', 'materials and methods', 'patients and methods',
      'study design', 'methoden', 'material und methoden',
      'patienten und methoden', 'studiendesign', 'statistical analysis',
      'statistische analyse', 'statistical methods'
    ];
    var methodsResumeHeaders = [
      'results', 'ergebnisse', 'findings', 'outcomes', 'discussion',
      'diskussion', 'conclusion', 'schlussfolgerung'
    ];

    var lines2 = cleaned.split('\n');
    var result2 = [];
    var skip2 = false;

    for (var i = 0; i < lines2.length; i++) {
      var line2 = lines2[i];
      var trimmed2 = line2.trim();
      var lower2 = trimmed2.toLowerCase();

      if (trimmed2.length > 0 && trimmed2.length < 80) {
        var isMethodSkip = false;
        for (var s = 0; s < methodsSkipHeaders.length; s++) {
          if (lower2 === methodsSkipHeaders[s] ||
              lower2 === methodsSkipHeaders[s] + ':' ||
              lower2.indexOf(methodsSkipHeaders[s] + ':') === 0) {
            isMethodSkip = true;
            break;
          }
        }
        if (isMethodSkip) { skip2 = true; continue; }

        if (skip2) {
          for (var r = 0; r < methodsResumeHeaders.length; r++) {
            if (lower2 === methodsResumeHeaders[r] ||
                lower2 === methodsResumeHeaders[r] + ':' ||
                lower2.indexOf(methodsResumeHeaders[r] + ':') === 0) {
              skip2 = false;
              break;
            }
          }
        }
      }

      if (!skip2) result2.push(line2);
    }

    var cleaned2 = result2.join('\n').trim();
    if (cleaned2.length < cleaned.length) {
      cleaned = cleaned2;
      methodsRemoved = true;
    }
  }

  return { text: cleaned, length: cleaned.length, methodsRemoved: methodsRemoved };
}

// Wrapper der nur die Länge zurückgibt (für Abwärtskompatibilität)
function cleanVolltextForLength(text) {
  return cleanVolltextForFlow(text).length;
}


function smartSavePaper(dashboard, paper, volltext) {
  var uuid = Utilities.getUuid();
  var now  = new Date();
  var volltextStatus = (volltext && volltext.length > 100) ? 'Volltext vorhanden' : 'Kein Volltext';

  var CHUNK_SIZE = 45000;
  var vt1 = '', vt2 = '', vt3 = '';
  var methodsRemovedNote = '';
  if (volltext && volltext.length > 0) {
    // Bereinigung + ggf. Methodenteil entfernen wenn zu lang
    var cleanResult = cleanVolltextForFlow(volltext);
    var volltextForChunks = cleanResult.text || volltext;
    if (cleanResult.methodsRemoved) {
      methodsRemovedNote = '⚠️ Methodenteil automatisch entfernt (Volltext > ' +
        VOLLTEXT_MAX_CHARS + ' Zeichen). Original: ' + volltext.length +
        ' → Bereinigt: ' + cleanResult.length + ' Zeichen';
      Logger.log('[CLEAN] ' + uuid + ' | Methodenteil entfernt | ' +
        volltext.length + ' → ' + cleanResult.length + ' Zeichen');
    }
    vt1 = volltextForChunks.substring(0, CHUNK_SIZE);
    if (volltextForChunks.length > CHUNK_SIZE)     vt2 = volltextForChunks.substring(CHUNK_SIZE, CHUNK_SIZE * 2);
    if (volltextForChunks.length > CHUNK_SIZE * 2) vt3 = volltextForChunks.substring(CHUNK_SIZE * 2, CHUNK_SIZE * 3);
  }

  var dataMap = {
    'UUID':                    uuid,
    'PMID':                    paper.pmid,
    'DOI':                     paper.doi,
    'Titel':                   paper.title,
    'Autoren':                 paper.authors,
    'Publikationsdatum':       paper.pubDate,
    'Journal/Quelle':          paper.journal,
    'Quelle':                  'Smart Import v7.2',
    'Link':                    'https://pubmed.ncbi.nlm.nih.gov/' + paper.pmid + '/',
    'Volltext-Status':         volltextStatus,
    'Inhalt/Abstract':         paper.abstract,
    'Volltext/Extrakt':        vt1,
    'Volltext_Teil2':          vt2,
    'Volltext_Teil3':          vt3,
    'Volltext Original Länge':  volltext ? volltext.length : 0,
    'Volltext Bereinigt Länge': volltext ? cleanVolltextForFlow(volltext).length : 0,
    'Fehler-Details':           methodsRemovedNote,
    'Status':                  volltextStatus === 'Kein Volltext'
                                 ? 'Neu importiert | Nur Abstract'
                                 : 'Neu importiert | Volltext',
    'Import-Timestamp':        now,
    'Letzte Änderung':         now,
    'Mail-Datum':              paper.pubDate
  };

  var headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];

  var lastDataRow = dashboard.getLastRow();
  var allRows = dashboard.getRange(2, 1, Math.max(lastDataRow, 2), 1).getValues();
  var nextRow = 2;
  for (var r = 0; r < allRows.length; r++) {
    if (allRows[r][0] && allRows[r][0] !== '') {
      nextRow = r + 3;
    }
  }

  var row = headers.map(function(header) {
    var key = String(header).trim();
    return dataMap.hasOwnProperty(key) ? dataMap[key] : '';
  });

  dashboard.getRange(nextRow, 1, 1, row.length).setValues([row]);
  Logger.log('Zeile ' + nextRow + ' | UUID: ' + uuid + ' | PMID: ' + paper.pmid);

  var triggerCol = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === 'Flow_Trigger_Researcher') {
      triggerCol = c + 1;
      break;
    }
  }
  if (triggerCol > 0) {
    dashboard.getRange(nextRow, triggerCol).setValue('PENDING');
    Logger.log('Flow_Trigger_Researcher PENDING gesetzt für Zeile ' + nextRow);
  } else {
    Logger.log('⚠️ Flow_Trigger_Researcher Spalte nicht gefunden! Headers: ' + headers.join('|'));
  }
  return uuid; // ← NEU: für Duplikat-Tracking
}


