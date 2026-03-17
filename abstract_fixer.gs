// FILE: abstract_fixer.gs
/**
 * ==========================================
 * ABSTRACT VERVOLLSTÄNDIGUNG
 * Alle Quellen: Semantic Scholar, bioRxiv, medRxiv, Crossref, PubMed, Europe PMC
 * ==========================================
 */

/**
 * ✅ Hauptfunktion: Alle gekürzte Abstracts vervollständigen
 */
function fixTruncatedAbstracts() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Dashboard");
  
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert("Keine Daten im Dashboard");
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();
  
  const abstractCol = headers.indexOf("Inhalt/Abstract");
  const pmidCol = headers.indexOf("PMID");
  const doiCol = headers.indexOf("DOI");
  
  let foundCount = 0;
  let fixedCount = 0;
  let failedCount = 0;
  const failed = [];
  
  SpreadsheetApp.getUi().alert("Starte Abstract-Vervollständigung mit allen Quellen...\n\nDies kann einige Minuten dauern.");
  
  data.forEach((row, index) => {
    const rowIndex = index + 2;
    const abstract = String(row[abstractCol] || "").trim();
    const pmid = String(row[pmidCol] || "");
    const doi = String(row[doiCol] || "");
    const titel = String(row[headers.indexOf("Titel")] || "");
    
    if (isTruncated(abstract)) {
      foundCount++;
      Logger.log(`\n=== Zeile ${rowIndex}: "${titel.substring(0, 50)}..." ===`);
      
      let fullAbstract = fetchFullAbstract(doi, pmid);
      
      if (fullAbstract && fullAbstract.length > abstract.length) {
        sheet.getRange(rowIndex, abstractCol + 1).setValue(fullAbstract);
        sheet.getRange(rowIndex, abstractCol + 1).setBackground("#d9ead3");
        fixedCount++;
        Logger.log(`✅ ERFOLG: ${fullAbstract.length} Zeichen`);
        Utilities.sleep(1000);
      } else {
        failedCount++;
        failed.push({ row: rowIndex, titel: titel.substring(0, 50), pmid: pmid, doi: doi });
        Logger.log(`❌ FEHLGESCHLAGEN`);
      }
    }
  });
  
  let message = `=== ABSTRACT VERVOLLSTÄNDIGUNG ===\n\n`;
  message += `🔍 Gefunden: ${foundCount} gekürzte Abstracts\n`;
  message += `✅ Vervollständigt: ${fixedCount}\n`;
  message += `❌ Fehlgeschlagen: ${failedCount}\n\n`;
  
  if (failed.length > 0) {
    message += `--- NICHT VERVOLLSTÄNDIGT (erste 5) ---\n`;
    failed.slice(0, 5).forEach(item => {
      message += `\nZeile ${item.row}: ${item.titel}\n`;
    });
  }
  
  SpreadsheetApp.getUi().alert(message);
  
  if (typeof logAction === 'function') {
    logAction("Abstract Fix", `${fixedCount}/${foundCount} vervollständigt`);
  }
}

/**
 * ✅ Nur für ausgewählte Zeilen
 */
function fixTruncatedAbstractsForSelected() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  const selection = sheet.getActiveRange();
  
  if (!selection) {
    SpreadsheetApp.getUi().alert("Bitte Zeilen auswählen");
    return;
  }
  
  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  
  if (startRow < 2) {
    SpreadsheetApp.getUi().alert("Bitte Datenzeilen auswählen");
    return;
  }
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const abstractCol = headers.indexOf("Inhalt/Abstract");
  const pmidCol = headers.indexOf("PMID");
  const doiCol = headers.indexOf("DOI");
  
  let fixedCount = 0;
  
  for (let i = 0; i < numRows; i++) {
    const row = startRow + i;
    const abstract = sheet.getRange(row, abstractCol + 1).getValue();
    const pmid = sheet.getRange(row, pmidCol + 1).getValue();
    const doi = sheet.getRange(row, doiCol + 1).getValue();
    
    if (isTruncated(String(abstract))) {
      let fullAbstract = fetchFullAbstract(doi, pmid);
      
      if (fullAbstract && fullAbstract.length > String(abstract).length) {
        sheet.getRange(row, abstractCol + 1).setValue(fullAbstract);
        sheet.getRange(row, abstractCol + 1).setBackground("#d9ead3");
        fixedCount++;
        Utilities.sleep(1000);
      }
    }
  }
  
  SpreadsheetApp.getUi().alert(`✅ ${fixedCount} von ${numRows} Abstracts vervollständigt`);
}

/**
 * ✅ ZENTRALE Funktion: Versucht ALLE Quellen in optimaler Reihenfolge
 */
function fetchFullAbstract(doi, pmid) {
  let abstract = null;
  
  // 1. Semantic Scholar (beste Quelle!)
  if (doi) {
    Logger.log(`  → Semantic Scholar (DOI)`);
    abstract = fetchSemanticScholar(doi, pmid);
    if (abstract) return abstract;
  }
  
  // 1.5. eLife DOI bereinigen (.sa* = Peer Review Artifact)
  if (doi && doi.includes('.sa')) {
    const cleanDoi = doi.replace(/\.sa\d+$/, '');
    Logger.log(`  🔧 eLife DOI bereinigt: ${doi} → ${cleanDoi}`);
    doi = cleanDoi; // Nutze bereinigte DOI für alle weiteren Versuche
  }
  
  // 2. DIREKT von DOI-URL scrapen (für neue Papers!)
  if (doi) {
    Logger.log(`  → Direkt-Scraping via DOI`);
    abstract = scrapeDirectDOI(doi);
    if (abstract) return abstract;
  }
  
  // 3. Crossref (inkl. besseres Parsing)
  if (doi) {
    Logger.log(`  → Crossref API`);
    abstract = fetchCrossref(doi);
    if (abstract) return abstract;
  }
  
  // 3. bioRxiv/medRxiv (Preprints)
  if (doi) {
    Logger.log(`  → bioRxiv/medRxiv`);
    abstract = fetchBiorxiv(doi);
    if (abstract) return abstract;
  }
  
  // 3.5. IMR Press direkt (DOI-Redirect funktioniert nicht!)
  if (doi && doi.includes('10.31083/')) {
    Logger.log(`  → IMR Press (direkt)`);
    abstract = fetchIMRPressDirect(doi);
    if (abstract) return abstract;
  }
  
  // 4. PubMed
  if (pmid) {
    Logger.log(`  → PubMed`);
    abstract = fetchPubmed(pmid);
    if (abstract) return abstract;
  }
  
  // 5. Europe PMC
  if (pmid || doi) {
    Logger.log(`  → Europe PMC`);
    abstract = fetchEuropePMC(pmid, doi);
    if (abstract) return abstract;
  }
  
  // 6. Unpaywall
  if (doi) {
    Logger.log(`  → Unpaywall`);
    abstract = fetchUnpaywall(doi);
    if (abstract) return abstract;
  }
  
  // 7. Web-Scraping (letzter Versuch)
  if (doi) {
    Logger.log(`  → Web-Scraping`);
    abstract = fetchViaScraping(doi);
    if (abstract) return abstract;
  }
  
  return null;
}

/**
 * ✅ Semantic Scholar API
 */
function fetchSemanticScholar(doi, pmid) {
  try {
    let identifier = doi ? `DOI:${doi}` : (pmid ? `PMID:${pmid}` : null);
    if (!identifier) return null;
    
    const url = `https://api.semanticscholar.org/graph/v1/paper/${encodeURIComponent(identifier)}?fields=abstract`;
    
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return null;
    
    const data = JSON.parse(response.getContentText());
    
    if (data.abstract && data.abstract.length > 100) {
      Logger.log(`    ✅ ${data.abstract.length} Zeichen`);
      return data.abstract;
    }
  } catch (e) {
    Logger.log(`    ❌ Error: ${e.message}`);
  }
  return null;
}

/**
 * ✅ DIREKT von DOI-URL scrapen (für sehr neue Papers!)
 */
function scrapeDirectDOI(doi) {
  try {
    // Baue DOI-URL
    const url = `https://doi.org/${doi}`;
    
    Logger.log(`    → Lade direkt: ${url}`);
    
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; ResearchBot/1.0)' },
      followRedirects: true
    });
    
    if (response.getResponseCode() !== 200) {
      Logger.log(`    ❌ HTTP ${response.getResponseCode()}`);
      return null;
    }
    
    const html = response.getContentText();
    const finalUrl = response.getResponseCode() === 200 ? url : null;
    
    Logger.log(`    → Gelandet auf: ${finalUrl || 'unknown'}`);
    
    // Publisher-Detection via HTML-Inhalt
    let abstract = null;
    
    // BMC/BioMed Central
    if (html.includes('biomedcentral.com') || html.includes('springeropen.com')) {
      Logger.log(`    → Erkannt: BMC`);
      abstract = parseBMC(html);
    }
    // Elsevier/ScienceDirect
    else if (html.includes('sciencedirect.com') || html.includes('elsevier.com')) {
      Logger.log(`    → Erkannt: Elsevier`);
      abstract = parseElsevier(html);
    }
    // Springer/Nature
    else if (html.includes('springer.com') || html.includes('nature.com')) {
      Logger.log(`    → Erkannt: Springer`);
      abstract = parseSpringer(html);
    }
    // Wiley
    else if (html.includes('wiley.com')) {
      Logger.log(`    → Erkannt: Wiley`);
      abstract = parseWiley(html);
    }
    // Oxford
    else if (html.includes('oup.com') || html.includes('academic.oup')) {
      Logger.log(`    → Erkannt: Oxford`);
      abstract = parseOxford(html);
    }
    // MDPI
    else if (html.includes('mdpi.com')) {
      Logger.log(`    → Erkannt: MDPI`);
      abstract = parseMDPI(html);
    }
    // Cureus
    else if (html.includes('cureus.com')) {
      Logger.log(`    → Erkannt: Cureus`);
      abstract = parseCureus(html);
    }
    // Frontiers
    else if (html.includes('frontiersin.org')) {
      Logger.log(`    → Erkannt: Frontiers`);
      abstract = parseFrontiers(html);
    }
    // PLOS
    else if (html.includes('plos.org')) {
      Logger.log(`    → Erkannt: PLOS`);
      abstract = parsePLOS(html);
    }
    // IMR Press (NEU!)
    else if (html.includes('imrpress.com')) {
      Logger.log(`    → Erkannt: IMR Press`);
      abstract = parseIMRPress(html);
    }
    // eLife (NEU!)
    else if (html.includes('elifesciences.org')) {
      Logger.log(`    → Erkannt: eLife`);
      abstract = parseELife(html);
    }
    // Generic Fallback
    else {
      Logger.log(`    → Generic Parser`);
      abstract = parseGeneric(html);
    }
    
    if (abstract && abstract.length > 100) {
      Logger.log(`    ✅ ${abstract.length} Zeichen`);
      return abstract;
    }
    
    Logger.log(`    ❌ Kein Abstract gefunden`);
    return null;
    
  } catch (e) {
    Logger.log(`    ❌ Error: ${e.message}`);
    return null;
  }
}

/**
 * ✅ Crossref API (verbessertes Parsing)
 */
function fetchCrossref(doi) {
  try {
    const url = `https://api.crossref.org/works/${encodeURIComponent(doi)}`;
    
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return null;
    
    const data = JSON.parse(response.getContentText());
    if (!data.message) return null;
    
    if (data.message.abstract) {
      let text = data.message.abstract;
      
      // Entferne alle XML/JATS Tags
      text = text.replace(/<jats:title[^>]*>.*?<\/jats:title>/gi, "");
      text = text.replace(/<jats:sec[^>]*>/gi, "");
      text = text.replace(/<\/jats:sec>/gi, "");
      text = text.replace(/<jats:p>/gi, "");
      text = text.replace(/<\/jats:p>/gi, " ");
      text = text.replace(/<jats:italic>/gi, "");
      text = text.replace(/<\/jats:italic>/gi, "");
      text = text.replace(/<jats:bold>/gi, "");
      text = text.replace(/<\/jats:bold>/gi, "");
      text = text.replace(/<jats:sub>/gi, "");
      text = text.replace(/<\/jats:sub>/gi, "");
      text = text.replace(/<jats:sup>/gi, "");
      text = text.replace(/<\/jats:sup>/gi, "");
      text = text.replace(/<[^>]+>/g, "");
      text = text.replace(/\s+/g, " ");
      text = text.trim();
      
      if (text.length > 100) {
        Logger.log(`    ✅ ${text.length} Zeichen`);
        return text;
      }
    }
    
    // Wenn kein Abstract in API, versuche Web-Scraping der Crossref-URL
    if (data.message.URL) {
      return scrapeCrossrefUrl(data.message.URL);
    }
    
  } catch (e) {
    Logger.log(`    ❌ Error: ${e.message}`);
  }
  return null;
}

/**
 * ✅ bioRxiv/medRxiv API
 */
function fetchBiorxiv(doi) {
  try {
    // bioRxiv API
    let url = `https://api.biorxiv.org/details/biorxiv/${doi}`;
    let response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.collection && data.collection.length > 0) {
        const abstract = data.collection[0].abstract;
        if (abstract && abstract.length > 100) {
          Logger.log(`    ✅ bioRxiv: ${abstract.length} Zeichen`);
          return abstract;
        }
      }
    }
    
    // medRxiv API (gleiche Struktur)
    url = `https://api.biorxiv.org/details/medrxiv/${doi}`;
    response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.collection && data.collection.length > 0) {
        const abstract = data.collection[0].abstract;
        if (abstract && abstract.length > 100) {
          Logger.log(`    ✅ medRxiv: ${abstract.length} Zeichen`);
          return abstract;
        }
      }
    }
    
  } catch (e) {
    Logger.log(`    ❌ Error: ${e.message}`);
  }
  return null;
}

/**
 * ✅ IMR Press direkt laden (DOI-Redirect kaputt!)
 */
function fetchIMRPressDirect(doi) {
  try {
    // DOI: 10.31083/HSF50002
    // URL: https://www.imrpress.com/journal/HSF/29/1/10.31083/HSF50002
    
    // Extrahiere Journal-Code und Paper-ID
    const match = doi.match(/10\.31083\/([A-Z]+)(\d+)/);
    if (!match) return null;
    
    const journal = match[1]; // z.B. "HSF"
    const paperId = match[2];  // z.B. "50002"
    
    // Baue URL-Varianten (probiere häufige Volume/Issue Kombinationen)
    const possibleUrls = [
      // Mit Volume/Issue (häufigste Struktur)
      `https://www.imrpress.com/journal/${journal}/1/1/${doi}`,
      `https://www.imrpress.com/journal/${journal}/29/1/${doi}`,
      `https://www.imrpress.com/journal/${journal}/30/1/${doi}`,
      `https://www.imrpress.com/journal/${journal}/28/1/${doi}`,
      `https://www.imrpress.com/journal/${journal}/27/1/${doi}`,
      
      // Ohne Volume/Issue
      `https://www.imrpress.com/journal/${journal}/article/${doi}`,
      `https://www.imrpress.com/journal/${journal}/${doi}`,
      
      // Alternative Strukturen
      `https://www.imrpress.com/journal/${journal}/articles/${doi}`,
      `https://imrpress.com/journal/${journal}/${doi}`
    ];
    
    for (const url of possibleUrls) {
      try {
        Logger.log(`    → Versuche: ${url}`);
        
        const response = UrlFetchApp.fetch(url, {
          muteHttpExceptions: true,
          headers: { 'User-Agent': 'Mozilla/5.0 (compatible; ResearchBot/1.0)' },
          followRedirects: true
        });
        
        if (response.getResponseCode() === 200) {
          const html = response.getContentText();
          
          if (html.includes('ipub-html-title') || html.includes('Abstract')) {
            const abstract = parseIMRPress(html);
            if (abstract) {
              Logger.log(`    ✅ IMR Press direkt: ${abstract.length} chars`);
              return abstract;
            }
          }
        }
        
        Utilities.sleep(200); // Rate limiting
        
      } catch (e) {
        // Continue to next URL
      }
    }
    
  } catch (e) {
    Logger.log(`    ❌ IMR Press Error: ${e.message}`);
  }
  return null;
}

/**
 * ✅ PubMed API
 */
function fetchPubmed(pmid) {
  try {
    const url = `https://eutils.ncbi.nlm.nih.gov/entrez/eutils/efetch.fcgi?db=pubmed&id=${pmid}&retmode=xml&rettype=abstract`;
    
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return null;
    
    const xml = response.getContentText();
    const match = xml.match(/<Abstract>([\s\S]*?)<\/Abstract>/);
    if (!match) return null;
    
    let text = match[1];
    text = text.replace(/<AbstractText[^>]*>/g, "");
    text = text.replace(/<\/AbstractText>/g, " ");
    text = text.replace(/<[^>]+>/g, "");
    text = text.replace(/&lt;/g, "<");
    text = text.replace(/&gt;/g, ">");
    text = text.replace(/&amp;/g, "&");
    text = text.replace(/\s+/g, " ");
    text = text.trim();
    
    if (text.length > 100) {
      Logger.log(`    ✅ ${text.length} Zeichen`);
      return text;
    }
  } catch (e) {
    Logger.log(`    ❌ Error: ${e.message}`);
  }
  return null;
}

/**
 * ✅ Europe PMC API
 */
function fetchEuropePMC(pmid, doi) {
  try {
    let query = "";
    if (pmid) query = `EXT_ID:${pmid}`;
    else if (doi) query = `DOI:"${doi}"`;
    else return null;
    
    const url = `https://www.ebi.ac.uk/europepmc/webservices/rest/search?query=${encodeURIComponent(query)}&format=json`;
    
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return null;
    
    const data = JSON.parse(response.getContentText());
    
    if (data.resultList && data.resultList.result && data.resultList.result.length > 0) {
      const abstract = data.resultList.result[0].abstractText;
      if (abstract && abstract.length > 100) {
        Logger.log(`    ✅ ${abstract.length} Zeichen`);
        return abstract;
      }
    }
  } catch (e) {
    Logger.log(`    ❌ Error: ${e.message}`);
  }
  return null;
}

/**
 * ✅ Unpaywall API
 */
function fetchUnpaywall(doi) {
  try {
    const email = Session.getActiveUser().getEmail() || "research@example.com";
    const url = `https://api.unpaywall.org/v2/${encodeURIComponent(doi)}?email=${email}`;
    
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return null;
    
    const data = JSON.parse(response.getContentText());
    
    if (data.abstract && data.abstract.length > 100) {
      Logger.log(`    ✅ ${data.abstract.length} Zeichen`);
      return data.abstract;
    }
  } catch (e) {
    Logger.log(`    ❌ Error: ${e.message}`);
  }
  return null;
}

/**
 * ✅ Web-Scraping via Crossref URL
 */
function scrapeCrossrefUrl(url) {
  try {
    Logger.log(`    → Lade: ${url}`);
    
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0 (compatible; ResearchBot/1.0)' },
      followRedirects: true
    });
    
    if (response.getResponseCode() !== 200) return null;
    
    const html = response.getContentText();
    
    // Publisher-spezifisches Parsing
    let abstract = null;
    
    if (url.includes('sciencedirect') || url.includes('elsevier')) {
      abstract = parseElsevier(html);
    } else if (url.includes('springer') || url.includes('nature')) {
      abstract = parseSpringer(html);
    } else if (url.includes('wiley')) {
      abstract = parseWiley(html);
    } else if (url.includes('oup.com') || url.includes('academic.oup')) {
      abstract = parseOxford(html);
    } else if (url.includes('biomedcentral') || url.includes('bmj')) {
      abstract = parseBMC(html);
    } else if (url.includes('mdpi')) {
      abstract = parseMDPI(html);
    } else if (url.includes('cureus')) {
      abstract = parseCureus(html);
    } else if (url.includes('frontiersin')) {
      abstract = parseFrontiers(html);
    } else if (url.includes('plos')) {
      abstract = parsePLOS(html);
    } else if (url.includes('imrpress')) {
      abstract = parseIMRPress(html);
    } else if (url.includes('elifesciences')) {
      abstract = parseELife(html);
    } else {
      abstract = parseGeneric(html);
    }
    
    if (abstract && abstract.length > 100) {
      Logger.log(`    ✅ Web-Scraping: ${abstract.length} Zeichen`);
      return abstract;
    }
    
  } catch (e) {
    Logger.log(`    ❌ Scraping Error: ${e.message}`);
  }
  return null;
}

/**
 * ✅ Fallback Web-Scraping
 */
function fetchViaScraping(doi) {
  try {
    // Hole URL von Crossref
    const crossrefUrl = `https://api.crossref.org/works/${encodeURIComponent(doi)}`;
    const response = UrlFetchApp.fetch(crossrefUrl, { muteHttpExceptions: true });
    
    if (response.getResponseCode() === 200) {
      const data = JSON.parse(response.getContentText());
      if (data.message && data.message.URL) {
        return scrapeCrossrefUrl(data.message.URL);
      }
    }
  } catch (e) {
    Logger.log(`    ❌ Error: ${e.message}`);
  }
  return null;
}

// ==========================================
// PUBLISHER-SPEZIFISCHE PARSER
// ==========================================

function parseElsevier(html) {
  // Methode 1: Finde <h2>Abstract</h2> und dann <div id="as0005">
  const abstractDivMatch = html.match(/<h2[^>]*>Abstract<\/h2>[\s\S]*?<div[^>]*id="as\d+"[^>]*>([\s\S]*?)<\/div>/i);
  if (abstractDivMatch) {
    let text = abstractDivMatch[1];
    text = text.replace(/<[^>]+>/g, " ");
    text = text.replace(/\s+/g, " ");
    text = text.trim();
    
    if (text.length > 100) {
      Logger.log(`    → Elsevier (as-div): ${text.length} chars`);
      return text;
    }
  }
  
  // Methode 2: Standard patterns
  const patterns = [
    /<section[^>]*class="[^"]*abstract[^"]*"[^>]*>([\s\S]*?)<\/section>/i,
    /<div[^>]*class="[^"]*abstract[^"]*"[^>]*>([\s\S]*?)<\/div>/i,
    /<div[^>]*id="abstracts?"[^>]*>([\s\S]*?)<\/div>/i,
    /<div[^>]*id="abspara[^"]*"[^>]*>([\s\S]*?)<\/div>/i,
    /<ce:abstract[^>]*>([\s\S]*?)<\/ce:abstract>/i,
    /<ce:para[^>]*view="all"[^>]*>([\s\S]*?)<\/ce:para>/i
  ];
  
  for (const pattern of patterns) {
    const match = html.match(pattern);
    if (match) {
      let text = match[1];
      text = text.replace(/<h[0-9][^>]*>.*?<\/h[0-9]>/gi, "");
      text = text.replace(/<ce:title[^>]*>.*?<\/ce:title>/gi, "");
      text = text.replace(/<[^>]+>/g, " ");
      text = text.replace(/\s+/g, " ");
      text = text.trim();
      
      if (text.length > 100) {
        Logger.log(`    → Elsevier (pattern): ${text.length} chars`);
        return text;
      }
    }
  }
  
  return null;
}

function parseSpringer(html) {
  const patterns = [
    /<div[^>]*id="Abs1-content"[^>]*>([\s\S]*?)<\/div>/i,
    /<section[^>]*data-title="Abstract"[^>]*>([\s\S]*?)<\/section>/i,
    /<div[^>]*class="[^"]*AbstractSection[^"]*"[^>]*>([\s\S]*?)<\/div>/i
  ];
  return extractWithPatterns(html, patterns);
}

function parseWiley(html) {
  const patterns = [
    /<section[^>]*class="[^"]*article-section__abstract[^"]*"[^>]*>([\s\S]*?)<\/section>/i,
    /<div[^>]*class="[^"]*article-section__content[^"]*"[^>]*>([\s\S]*?)<\/div>/i
  ];
  return extractWithPatterns(html, patterns);
}

function parseOxford(html) {
  // Methode 1: scrollto-destination Abstract (NEU - aus Screenshot!)
  const scrollMatch = html.match(/<h2[^>]*id="[^"]*"[^>]*class="abstract-title[^"]*"[^>]*>Abstract<\/h2>[\s\S]*?<section[^>]*class="abstract"[^>]*>([\s\S]*?)<\/section>/i);
  if (scrollMatch) {
    let text = scrollMatch[1];
    text = text.replace(/<div[^>]*class="abstract[^"]*Border[^"]*"[^>]*>.*?<\/div>/gi, "");
    text = text.replace(/<div[^>]*class="article-metadata[^"]*"[^>]*>.*?<\/div>/gi, "");
    text = text.replace(/<h[0-9][^>]*>.*?<\/h[0-9]>/gi, "");
    text = text.replace(/<[^>]+>/g, " ");
    text = text.replace(/\s+/g, " ");
    text = text.trim();
    
    if (text.length > 100) {
      Logger.log(`    → Oxford (scrollto): ${text.length} chars`);
      return text;
    }
  }
  
  // Methode 2: Standard abstract section
  const patterns = [
    /<section[^>]*class="[^"]*abstract[^"]*"[^>]*>([\s\S]*?)<\/section>/i,
    /<div[^>]*class="[^"]*abstract[^"]*"[^>]*>([\s\S]*?)<\/div>/i
  ];
  return extractWithPatterns(html, patterns);
}

function parseBMC(html) {
  const patterns = [
    /<section[^>]*data-title="Abstract"[^>]*>([\s\S]*?)<\/section>/i,
    /<div[^>]*class="[^"]*Abstract[^"]*"[^>]*>([\s\S]*?)<\/div>/i,
    /<section[^>]*id="Abs1"[^>]*>([\s\S]*?)<\/section>/i
  ];
  return extractWithPatterns(html, patterns);
}

function parseMDPI(html) {
  const patterns = [
    /<div[^>]*class="[^"]*art-abstract[^"]*"[^>]*>([\s\S]*?)<\/div>/i,
    /<section[^>]*class="[^"]*abstract[^"]*"[^>]*>([\s\S]*?)<\/section>/i
  ];
  return extractWithPatterns(html, patterns);
}

function parseCureus(html) {
  // Methode 1: <h3 class="reg" id="abstract">
  const regMatch = html.match(/<h3[^>]*class="reg"[^>]*id="abstract"[^>]*>Abstract<\/h3>[\s\S]{0,100}?<p[^>]*>([\s\S]*?)<\/p>/i);
  if (regMatch) {
    let text = regMatch[1];
    text = text.replace(/<[^>]+>/g, " ");
    text = text.replace(/\s+/g, " ");
    text = text.trim();
    
    if (text.length > 100) {
      Logger.log(`    → Cureus (reg-p): ${text.length} chars`);
      return text;
    }
  }
  
  // Methode 2: Alle <p> zwischen abstract und nächster section
  const sectionMatch = html.match(/<h3[^>]*id="abstract"[^>]*>Abstract<\/h3>([\s\S]*?)<h3[^>]*>/i);
  if (sectionMatch) {
    let text = sectionMatch[1];
    // Entferne alle außer <p> tags
    const paragraphs = text.match(/<p[^>]*>([\s\S]*?)<\/p>/gi);
    if (paragraphs && paragraphs.length > 0) {
      text = paragraphs.join(" ");
      text = text.replace(/<[^>]+>/g, " ");
      text = text.replace(/\s+/g, " ");
      text = text.trim();
      
      if (text.length > 100) {
        Logger.log(`    → Cureus (section): ${text.length} chars`);
        return text;
      }
    }
  }
  
  // Methode 3: abstract-content div
  const patterns = [
    /<div[^>]*class="[^"]*abstract-content[^"]*"[^>]*>([\s\S]*?)<\/div>/i,
    /<section[^>]*id="abstract"[^>]*>([\s\S]*?)<\/section>/i
  ];
  return extractWithPatterns(html, patterns);
}

function parseFrontiers(html) {
  const patterns = [
    /<div[^>]*class="[^"]*abstract[^"]*"[^>]*>([\s\S]*?)<\/div>/i,
    /<section[^>]*class="[^"]*abstract[^"]*"[^>]*>([\s\S]*?)<\/section>/i
  ];
  return extractWithPatterns(html, patterns);
}

function parsePLOS(html) {
  const patterns = [
    /<div[^>]*class="[^"]*abstract[^"]*"[^>]*>([\s\S]*?)<\/div>/i,
    /<section[^>]*id="abstract"[^>]*>([\s\S]*?)<\/section>/i
  ];
  return extractWithPatterns(html, patterns);
}

function parseIMRPress(html) {
  // IMR Press: <h3 class="ipub-html-title" id="Abstract">
  const imrMatch = html.match(/<h3[^>]*class="ipub-html-title"[^>]*id="Abstract"[^>]*>[\s\S]*?<div[^>]*class="ipub-html-content[^"]*"[^>]*>([\s\S]*?)<\/div>/i);
  if (imrMatch) {
    let text = imrMatch[1];
    text = text.replace(/<[^>]+>/g, " ");
    text = text.replace(/\s+/g, " ");
    text = text.trim();
    
    if (text.length > 100) {
      Logger.log(`    → IMR Press: ${text.length} chars`);
      return text;
    }
  }
  
  return null;
}

function parseELife(html) {
  // eLife: <h2 class="article-section__header_text">Abstract</h2>
  const elifeMatch = html.match(/<h2[^>]*class="article-section__header_text"[^>]*>Abstract<\/h2>[\s\S]*?<div[^>]*class="article-section__body"[^>]*>([\s\S]*?)<\/div>/i);
  if (elifeMatch) {
    let text = elifeMatch[1];
    text = text.replace(/<h[0-9][^>]*>.*?<\/h[0-9]>/gi, "");
    text = text.replace(/<[^>]+>/g, " ");
    text = text.replace(/\s+/g, " ");
    text = text.trim();
    
    if (text.length > 100) {
      Logger.log(`    → eLife: ${text.length} chars`);
      return text;
    }
  }
  
  return null;
}

function parseGeneric(html) {
  const patterns = [
    /<section[^>]*class="[^"]*abstract[^"]*"[^>]*>([\s\S]*?)<\/section>/i,
    /<div[^>]*class="[^"]*abstract[^"]*"[^>]*>([\s\S]*?)<\/div>/i,
    /<div[^>]*id="abstract"[^>]*>([\s\S]*?)<\/div>/i,
    /<section[^>]*id="abstract"[^>]*>([\s\S]*?)<\/section>/i
  ];
  return extractWithPatterns(html, patterns);
}

function extractWithPatterns(html, patterns) {
  for (const pattern of patterns) {
    const match = html.match(pattern);
    if (match) {
      let text = match[1];
      text = text.replace(/<h[0-9][^>]*>.*?<\/h[0-9]>/gi, "");
      text = text.replace(/<[^>]+>/g, " ");
      text = text.replace(/\s+/g, " ");
      text = text.trim();
      
      if (text.length > 100) return text;
    }
  }
  return null;
}

// ==========================================
// HILFSFUNKTIONEN
// ==========================================

function isTruncated(abstract) {
  if (!abstract || abstract.length < 50) return false;
  
  // REGEL 1: Wenn Abstract sehr lang ist (>1000 Zeichen), ist er wahrscheinlich vollständig
  if (abstract.length > 1000) return false;
  
  // REGEL 2: Wenn Abstract mit normalem Satzende endet (nicht "..."), ist er vermutlich vollständig
  const endsNormally = /[.!?][\s]*$/.test(abstract.trim());
  const hasEllipsis = /\.{3}$|…$/.test(abstract.trim());
  
  if (abstract.length > 500 && endsNormally && !hasEllipsis) {
    return false; // Lang genug + normales Ende = vollständig
  }
  
  // REGEL 3: Erkenne explizite Kürzungsmarker
  const truncationPatterns = [
    /\.{3}$/,           // "..."
    /…$/,               // "…"
    /\[truncated\]$/i,
    /\[see more\]$/i,
    /\(continued\)$/i
  ];
  
  for (const pattern of truncationPatterns) {
    if (pattern.test(abstract.trim())) return true;
  }
  
  // REGEL 4: Sehr kurze Abstracts (<300 Zeichen) mit "..." sind verdächtig
  if (abstract.length < 300 && abstract.includes("...")) return true;
  
  // REGEL 5: Mittelgroße Abstracts (300-500 Zeichen) die MIT "..." enden
  if (abstract.length >= 300 && abstract.length < 500 && hasEllipsis) return true;
  
  // Default: Nicht gekürzt
  return false;
}

function autoFixAbstractOnImport(uuid) {
  try {
    const data = getDashboardDataByUUID(uuid);
    if (!data) return;
    
    const abstract = String(data["Inhalt/Abstract"] || "").trim();
    
    if (isTruncated(abstract)) {
      const fullAbstract = fetchFullAbstract(data.DOI, data.PMID);
      
      if (fullAbstract && fullAbstract.length > abstract.length) {
        updateDashboardField(uuid, "Inhalt/Abstract", fullAbstract);
        Logger.log("✅ Abstract automatisch vervollständigt");
      }
    }
  } catch (e) {
    Logger.log("autoFixAbstractOnImport Error: " + e.message);
  }
}
