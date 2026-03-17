// FILE: ui_pubmed_menu_extended.gs

// ── Globale Ausschlüsse (an alle Gesamt-Queries angehängt) ────────────────
// Entfernt Tier-/Veterinär-Papers und klar fachfremde Infektionskrankheiten
var ANIMAL_EXCLUSIONS =
  ' NOT (veterinary[Title/Abstract] OR ""dairy cow""[Title/Abstract] OR ' +
  'bovine[Title/Abstract] OR canine[Title/Abstract] OR feline[Title/Abstract] OR ' +
  'murine[Title/Abstract] OR ""animal model""[Title/Abstract] OR ' +
  '""mouse model""[Title/Abstract] OR rats[MeSH Terms] OR mice[MeSH Terms])';

var IMMUNO_EXCLUSIONS =
  ' NOT (malaria[Title] OR tuberculosis[Title] OR ""hepatitis C""[Title] OR ' +
  '""hepatitis B""[Title] OR ""HIV infection""[Title] OR dengue[Title])';

var IMPORT_KATEGORIEN = {

  'Diabetes Typ 1': {
    icon: '🩸',
    // humans[MeSH] added → filtert Tier-Studien heraus
    // Produkt-spezifische Terme (tandem, Dexcom etc.) bleiben erhalten
    query: '((type 1 diabetes OR T1D OR T1DM OR insulin-dependent diabetes OR juvenile diabetes) AND ' +
           '(continuous glucose monitoring OR CGM OR insulin pump OR CSII OR automated insulin delivery OR AID OR ' +
           'closed-loop OR hybrid closed-loop OR flash glucose monitoring OR FGM OR real-time CGM OR rtCGM OR ' +
           'sensor-augmented pump OR SAP OR time in range OR TIR OR HbA1c OR hypoglycemia OR hypoglycaemia OR ' +
           'tandem OR t:slim OR Control-IQ OR Dexcom OR FreeStyle Libre) AND humans[MeSH])',
    einzelsuchen: [
      { name: 'AID / Closed-Loop (T1D)',  query: '(automated insulin delivery OR AID OR closed-loop OR hybrid closed-loop) AND (type 1 diabetes OR T1D) AND humans[MeSH]' },
      { name: 'CGM (T1D)',                query: '(continuous glucose monitoring OR CGM OR Dexcom OR FreeStyle Libre) AND (type 1 diabetes OR T1D) AND humans[MeSH]' },
      // Produkt-spezifisch → kein humans[MeSH], sonst entfallen IEEE-Vergleichsstudien
      { name: 't:slim X2 / Control-IQ',   query: '(tandem OR t:slim OR Control-IQ OR Basal-IQ)' },
      { name: 'Dexcom G6/G7 (T1D)',       query: '(Dexcom G6 OR Dexcom G7 OR Dexcom) AND (type 1 diabetes OR T1D)' },
      { name: 'Hypoglykämie (T1D)',        query: '(hypoglycemia OR hypoglycaemia) AND (type 1 diabetes OR T1D) AND (CGM OR pump OR automated) AND humans[MeSH]' },
      { name: 'HbA1c & TIR (T1D)',        query: '(HbA1c OR time in range OR TIR) AND (type 1 diabetes OR T1D) AND (pump OR CGM OR sensor) AND humans[MeSH]' }
    ]
  },

  'Diabetes Typ 2': {
    icon: '🩸',
    query: '((type 2 diabetes OR T2D OR T2DM OR NIDDM OR non-insulin-dependent diabetes) AND ' +
           '(continuous glucose monitoring OR CGM OR insulin pump OR automated insulin delivery OR AID OR ' +
           'flash glucose monitoring OR time in range OR TIR OR HbA1c OR hypoglycemia OR ' +
           'GLP-1 OR SGLT2 OR FreeStyle Libre OR Dexcom) AND humans[MeSH])',
    einzelsuchen: [
      { name: 'CGM (T2D)',              query: '(continuous glucose monitoring OR CGM OR FreeStyle Libre OR Dexcom) AND (type 2 diabetes OR T2D) AND humans[MeSH]' },
      { name: 'Insulin-Therapie (T2D)', query: '(insulin therapy OR insulin pump OR CSII) AND (type 2 diabetes OR T2D) AND humans[MeSH]' },
      { name: 'AID (T2D)',              query: '(automated insulin delivery OR AID OR closed-loop) AND (type 2 diabetes OR T2D) AND humans[MeSH]' },
      { name: 'HbA1c & TIR (T2D)',     query: '(HbA1c OR time in range OR TIR) AND (type 2 diabetes OR T2D) AND (CGM OR sensor) AND humans[MeSH]' },
      { name: 'Hypoglykämie (T2D)',     query: '(hypoglycemia OR hypoglycaemia) AND (type 2 diabetes OR T2D) AND humans[MeSH]' }
    ]
  },

  'Diabetes Gesamt': {
    icon: '🩸',
    // ""diabetes technology OR diabetes device"" war zu breit → ersetzt durch
    // konkrete Technologie-Terme; humans[MeSH] für Gesamt-Query
    query: '((diabetes mellitus OR diabetes OR diabetic OR T1D OR T2D OR T1DM OR T2DM) AND ' +
           '(continuous glucose monitoring OR CGM OR insulin pump OR CSII OR automated insulin delivery OR ' +
           'AID OR closed-loop OR flash glucose monitoring OR FGM OR time in range OR TIR OR ' +
           'HbA1c OR tandem OR t:slim OR Dexcom OR FreeStyle Libre OR sensor-augmented pump OR ' +
           'hybrid closed-loop OR real-time CGM) AND humans[MeSH])',
    einzelsuchen: [
      { name: 'Diabetes Allgemein',         query: '(diabetes mellitus) AND (insulin pump OR CGM OR sensor OR closed-loop) AND humans[MeSH]' },
      { name: 'CGM Allgemein',              query: '(continuous glucose monitoring OR CGM) AND (diabetes) AND humans[MeSH]' },
      { name: 'AID / Closed-Loop',          query: '(automated insulin delivery OR AID OR closed-loop OR hybrid closed-loop) AND diabetes AND humans[MeSH]' },
      // Produkt-Suchen ohne humans[MeSH] → IEEE-Vergleichsstudien bleiben
      { name: 't:slim X2 / Control-IQ',     query: '(tandem OR t:slim OR Control-IQ OR Basal-IQ)' },
      { name: 'Dexcom G6/G7',               query: '(Dexcom G6 OR Dexcom G7 OR Dexcom)' },
      { name: 'Medtronic 780G',             query: '(Medtronic 780G OR MiniMed 780G OR SmartGuard OR Simplera)' },
      { name: 'Omnipod 5',                  query: '(Omnipod 5 OR Omnipod) AND diabetes' },
      { name: 'Abbott Libre',               query: '(FreeStyle Libre OR Abbott Libre OR Libre 2 OR Libre 3)' },
      { name: 'Time in Range',              query: '(time in range OR TIR) AND (diabetes OR CGM) AND humans[MeSH]' },
      { name: 'HbA1c & Technology',         query: '(HbA1c OR hemoglobin A1c) AND (pump OR CGM OR sensor) AND humans[MeSH]' },
      { name: 'Hypoglykämie & Technology',  query: '(hypoglycemia OR hypoglycaemia) AND (CGM OR pump OR automated) AND humans[MeSH]' }
    ]
  },

  'Immunologie': {
    icon: '🧬',
    // ANIMAL_EXCLUSIONS + IMMUNO_EXCLUSIONS angehängt
    // humans[MeSH] bewusst NICHT global, da z.B. in-vitro-Studien zu
    // Immunoglobulin-Mechanismen relevant sein können
    query: '(immunoglobulin OR immunoglobulins OR IVIG OR SCIG OR subcutaneous immunoglobulin OR ' +
           'IgG replacement OR immunoglobulin therapy OR MOBILIZE OR VITALIZE OR Riliprubart OR ' +
           'primary immunodeficiency OR PID OR secondary immunodeficiency OR antibody deficiency OR ' +
           'hypogammaglobulinemia OR myasthenia gravis OR myositides OR myositis OR polymyositis OR ' +
           'dermatomyositis OR inclusion body myositis OR Efgartigimod OR Rozanolixizumab OR ' +
           'Rituximab OR anti-FcRn OR Deferoxamine)' +
           ANIMAL_EXCLUSIONS + IMMUNO_EXCLUSIONS,
    einzelsuchen: [
      { name: 'MOBILIZE (Kernprodukt)',      query: 'MOBILIZE AND (immunoglobulin OR SCIG OR subcutaneous)' },
      { name: 'VITALIZE (Kernprodukt)',      query: 'VITALIZE AND (immunoglobulin OR SCIG OR subcutaneous)' },
      { name: 'Riliprubart (Kernprodukt)',   query: 'Riliprubart OR (anti-FcRn AND immunoglobulin)' },
      { name: 'Immunoglobulin Allgemein',    query: '(immunoglobulin therapy OR IgG therapy OR SCIG OR IVIG)' + ANIMAL_EXCLUSIONS },
      { name: 'Primary Immunodeficiency',    query: '(primary immunodeficiency OR primary immune deficiency OR PID) AND (immunoglobulin OR IgG)' + ANIMAL_EXCLUSIONS },
      { name: 'Secondary Immunodeficiency',  query: '(secondary immunodeficiency OR secondary immune deficiency) AND (immunoglobulin OR treatment)' + ANIMAL_EXCLUSIONS },
      { name: 'Antibody Deficiency',         query: '(antibody deficiency OR hypogammaglobulinemia) AND (immunoglobulin OR IgG OR SCIG)' + ANIMAL_EXCLUSIONS },
      { name: 'Myasthenia Gravis',           query: '(myasthenia gravis) AND (treatment OR therapy OR immunoglobulin OR IgG)' },
      { name: 'Myositides',                  query: '(myositides OR myositis OR polymyositis OR dermatomyositis) AND (immunoglobulin OR treatment)' + ANIMAL_EXCLUSIONS },
      { name: 'Inclusion Body Myositis',     query: '(inclusion body myositis OR IBM) AND (immunoglobulin OR treatment)' },
      { name: 'Efgartigimod / Vyvgart',      query: '(Efgartigimod OR Vyvgart) AND (myasthenia OR immunoglobulin OR FcRn)' },
      { name: 'Rozanolixizumab',             query: 'Rozanolixizumab AND (myasthenia OR FcRn OR immunoglobulin)' },
      { name: 'Rituximab',                   query: 'Rituximab AND (myasthenia OR neuropathy OR myositis OR immunodeficiency)' },
      { name: 'Deferoxamine',                query: 'Deferoxamine AND (therapy OR treatment)' }
    ]
  },

  'Neurologie': {
    icon: '🧠',
    // humans[MeSH] für Gesamt; Produkt-Terme (Foslevodopa etc.) ohne Filter
    query: '(Foslevodopa OR Produodopa OR (Parkinson AND subcutaneous infusion) OR ' +
           'CIDP OR chronic inflammatory demyelinating polyneuropathy OR ' +
           '(neuropathy AND immunoglobulin) OR (neuropathy AND IgG) OR (neuropathy AND IVIG))' +
           ANIMAL_EXCLUSIONS,
    einzelsuchen: [
      // Produkt-spezifisch → kein Filter
      { name: 'Foslevodopa (Kernprodukt)',  query: '(Foslevodopa OR Produodopa) AND (Parkinson OR subcutaneous)' },
      { name: 'Produodopa (Kernprodukt)',   query: 'Produodopa AND (Parkinson OR subcutaneous infusion)' },
      { name: 'Parkinson & Infusion',      query: '(Parkinson OR Parkinsons disease) AND (subcutaneous infusion OR continuous infusion OR pump)' + ANIMAL_EXCLUSIONS },
      { name: 'CIDP',                      query: '(CIDP OR chronic inflammatory demyelinating polyneuropathy) AND (treatment OR immunoglobulin)' + ANIMAL_EXCLUSIONS },
      { name: 'Neuropathy & IgG',          query: '(neuropathy OR peripheral neuropathy) AND (immunoglobulin OR IgG OR IVIG)' + ANIMAL_EXCLUSIONS }
    ]
  }
};

// ── Hauptfunktion ─────────────────────────────────────────────────────────
function showPredefinedPubMedSearches() {
  var ui      = SpreadsheetApp.getUi();
  var katKeys = Object.keys(IMPORT_KATEGORIEN);

  // Schritt 1: Oberkategorie
  var step1 = '=== PUBMED IMPORT – KATEGORIE WÄHLEN ===\n\n';
  for (var i = 0; i < katKeys.length; i++) {
    step1 += (i+1) + '. ' + IMPORT_KATEGORIEN[katKeys[i]].icon + ' ' + katKeys[i] + '\n';
  }
  step1 += '\n0. 🔄 ALLE Kategorien (Smart Import gesamt)\n\nGib die Nummer ein:';

  var resp1 = ui.prompt('Import – Schritt 1/4: Kategorie', step1, ui.ButtonSet.OK_CANCEL);
  if (resp1.getSelectedButton() !== ui.Button.OK) return;

  var katChoice = parseInt(resp1.getResponseText().trim());
  if (katChoice === 0) { importAllTopicsSmart(); return; }
  if (isNaN(katChoice) || katChoice < 1 || katChoice > katKeys.length) { ui.alert('Ungültige Auswahl!'); return; }

  var selectedKatName = katKeys[katChoice - 1];
  var selectedKat     = IMPORT_KATEGORIEN[selectedKatName];

  // Schritt 2: Gesamt oder Einzelsuche
  var step2 = selectedKat.icon + ' ' + selectedKatName + '\n\n';
  step2    += '0. 📦 GESAMT (alle Suchen kombiniert)\n\n';
  for (var j = 0; j < selectedKat.einzelsuchen.length; j++) {
    step2 += (j+1) + '. ' + selectedKat.einzelsuchen[j].name + '\n';
  }
  step2 += '\nGib die Nummer ein:';

  var resp2 = ui.prompt('Import – Schritt 2/4: Suche', step2, ui.ButtonSet.OK_CANCEL);
  if (resp2.getSelectedButton() !== ui.Button.OK) return;

  var suchChoice = parseInt(resp2.getResponseText().trim());
  var finalQuery, finalName;

  if (suchChoice === 0) {
    finalQuery = selectedKat.query;
    finalName  = selectedKatName + ' (Gesamt)';
  } else if (suchChoice >= 1 && suchChoice <= selectedKat.einzelsuchen.length) {
    finalQuery = selectedKat.einzelsuchen[suchChoice - 1].query;
    finalName  = selectedKat.einzelsuchen[suchChoice - 1].name;
  } else {
    ui.alert('Ungültige Auswahl!'); return;
  }

  // Schritt 3: Datum
  var lastDate = getLastImportDateForCategory(selectedKatName);
  var step3    = 'Startdatum für:\n' + finalName + '\n\n';
  if (lastDate) {
    step3 += '📅 Letzter Import dieser Kategorie: ' + lastDate + '\n\n';
    step3 += '• Leer lassen = ab letztem Import (' + lastDate + ')\n';
  } else {
    step3 += '• Leer lassen = neueste Papers\n';
  }
  step3 += '• Datum eingeben (z.B. 2024/01/01)\n• ""alle"" = ohne Datumsfilter';

  var resp3 = ui.prompt('Import – Schritt 3/4: Datum', step3, ui.ButtonSet.OK_CANCEL);
  if (resp3.getSelectedButton() !== ui.Button.OK) return;

  var startDate = resp3.getResponseText().trim();
  if (!startDate && lastDate)                                startDate = lastDate;
  else if (!startDate || startDate.toLowerCase() === 'alle') startDate = null;

  // Schritt 4: Anzahl
  var resp4 = ui.prompt(
    'Import – Schritt 4/4: Anzahl',
    'Wie viele Papers importieren? (1-500)\nEmpfehlung: 50-100',
    ui.ButtonSet.OK_CANCEL
  );
  if (resp4.getSelectedButton() !== ui.Button.OK) return;

  var maxResults = parseInt(resp4.getResponseText().trim()) || 50;
  if (maxResults < 1 || maxResults > 500) { ui.alert('Bitte Zahl zwischen 1 und 500 eingeben!'); return; }

  var confirm = ui.alert(
    'Import starten?',
    '📂 ' + finalName + '\n' +
    '📅 Ab: ' + (startDate || 'neueste Papers') + '\n' +
    '🔢 Anzahl: ' + maxResults + '\n\nJetzt starten?',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  importPapersFromPubMed(finalQuery, startDate, maxResults, selectedKatName);
}

// ── Quick-Imports ─────────────────────────────────────────────────────────
function quickImportDiabetes() {
  importPapersFromPubMed(
    IMPORT_KATEGORIEN['Diabetes Gesamt'].query,
    getLastImportDateForCategory('Diabetes Gesamt'),
    50, 'Diabetes Gesamt'
  );
}

function quickImportImmunoNeuro() {
  importPapersFromPubMed(
    IMPORT_KATEGORIEN['Immunologie'].query + ' OR ' + IMPORT_KATEGORIEN['Neurologie'].query,
    null, 50, 'Immunologie'
  );
}

function quickImportCIDP() {
  importPapersFromPubMed(
    '(CIDP OR chronic inflammatory demyelinating polyneuropathy) AND (treatment OR immunoglobulin)' + ANIMAL_EXCLUSIONS,
    getLastImportDateForCategory('Neurologie'), 30, 'Neurologie'
  );
}

function quickImportMyastheniaGravis() {
  importPapersFromPubMed(
    '(myasthenia gravis) AND (treatment OR immunoglobulin OR Efgartigimod OR Vyvgart)',
    getLastImportDateForCategory('Immunologie'), 30, 'Immunologie'
  );
}
