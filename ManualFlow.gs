// FILE: manualFlow.gs

function openManualFlowDialog() {
  var html = HtmlService.createHtmlOutputFromFile('manualFlowDialog')
    .setWidth(860)
    .setHeight(720)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, '📋 Manual Flow – Gemini Analyse');
}

function openFlowStatusDialog() {
  var html = HtmlService.createHtmlOutputFromFile('flowStatusDialog')
    .setWidth(520)
    .setHeight(680)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModelessDialog(html, '📊 Pipeline Übersicht');
}

// ── getNextFlowItem ───────────────────────────────────────────────
function getNextFlowItem(filterFlow, skipKeys, filterKategorie) {
  skipKeys        = skipKeys        || [];
  filterKategorie = filterKategorie || 'ALLE';

  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) return null;

  var headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  var lastRow = dashboard.getLastRow();
  if (lastRow < 2) return null;

  var data = dashboard.getRange(2, 1, lastRow - 1, headers.length).getValues();

  var getIdx = function(name) {
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c]).trim() === name) return c;
    }
    return -1;
  };

  var isValidJson = function(str) {
    if (!str || str.length < 10) return false;
    var s = str.replace(/```json/g, '').replace(/```/g, '').trim();
    var firstBrace = s.indexOf('{');
    if (firstBrace === -1) return false;
    s = s.substring(firstBrace);
    try { JSON.parse(s); return true; }
    catch(e) { return false; }
  };

  var flows = [
    { name: 'RESEARCHER',     trigger: 'Flow_Trigger_Researcher',  output: 'Gemini Researcher' },
    { name: 'TRIAGE',         trigger: 'Flow_Trigger_Triage',      output: 'Gemini Triage' },
    { name: 'METADATEN',      trigger: 'Flow_Trigger_Metadaten',   output: 'Gemini Metadaten' },
    { name: 'MASTER_ANALYSE', trigger: 'Flow_Trigger_Analyse',     output: 'Gemini Master Analyse' },
    { name: 'REDAKTION',      trigger: 'Flow_Trigger_Redaktion',   output: 'Gemini Redaktion' },
    { name: 'FAKTENCHECK',    trigger: 'Flow_Trigger_Faktencheck', output: 'Gemini Faktencheck' },
    { name: 'REVIEW',         trigger: 'Flow_Trigger_Review',      output: 'Gemini Review' }
  ];

  var stopStatuses = [
    '✅ Workflow komplett',
    '⛔ Irrelevant – Workflow gestoppt',
    '🔻 Niedrig – Workflow gestoppt',
    '⚠️ Mittel – Manuell fortsetzen?'
  ];

  var cleanTriggers = [
    'Flow_Trigger_Metadaten', 'Flow_Trigger_Analyse',
    'Flow_Trigger_Redaktion', 'Flow_Trigger_Faktencheck', 'Flow_Trigger_Review'
  ];

  for (var r = 0; r < data.length; r++) {
    var row    = data[r];
    var uuid   = String(row[getIdx('UUID')]   || '').trim();
    var status = String(row[getIdx('Status')] || '').trim();
    var rowNum = r + 2;

    if (!uuid) continue;

    var skip = false;
    for (var s = 0; s < stopStatuses.length; s++) {
      if (status.indexOf(stopStatuses[s]) >= 0) { skip = true; break; }
    }
    if (status.indexOf('dauerhafter Fehler') >= 0) skip = true;
    if (skip) continue;

    // ── Hauptkategorie-Filter ─────────────────────────────────────
    if (filterKategorie !== 'ALLE') {
      var hauptkat = String(row[getIdx('Hauptkategorie')] || '').trim();
      // Wenn Hauptkategorie noch leer (vor Triage): nur zeigen wenn
      // Researcher oder Triage Flow gefiltert wird
      if (hauptkat && hauptkat !== filterKategorie) continue;
    }

    // ── Triage-Konsistenz-Check ───────────────────────────────────
    var triageOutput = String(row[getIdx('Gemini Triage')] || '').trim();
    if (triageOutput && isValidJson(triageOutput)) {
      try {
        var cleanT = triageOutput.replace(/```json/g, '').replace(/```/g, '').trim();
        var fbT    = cleanT.indexOf('{');
        if (fbT > 0) cleanT = cleanT.substring(fbT);
        var tjT    = JSON.parse(cleanT);
        var relT   = (tjT.relevanz || '').trim();
        if (relT === 'Irrelevant' || relT === 'Niedrig' || relT === 'Mittel') {
          for (var ct = 0; ct < cleanTriggers.length; ct++) {
            if (String(row[getIdx(cleanTriggers[ct])] || '').trim() !== '') {
              dashboard.getRange(rowNum, getIdx(cleanTriggers[ct]) + 1).clearContent();
              SpreadsheetApp.flush();
            }
          }
          continue;
        }
      } catch(e) {}
    }

    // ── Flow-Suche: linear durchlaufen ───────────────────────────
    for (var f = 0; f < flows.length; f++) {
      var flow       = flows[f];
      var triggerVal = String(row[getIdx(flow.trigger)] || '').trim();
      var outputVal  = String(row[getIdx(flow.output)]  || '').trim();

      // Output vorhanden und valide → weiter zum nächsten Flow
      if (outputVal !== '' && isValidJson(outputVal)) continue;

      // Trigger gesetzt, Output leer → dieser Flow ist dran
      if (triggerVal === 'PENDING' || triggerVal === 'START') {
        var itemKey = uuid + '|' + flow.name;
        if (skipKeys.indexOf(itemKey) >= 0) break;
        if (filterFlow && filterFlow !== 'ALLE' && flow.name !== filterFlow) break;

        var prompt = buildPrompt(uuid, flow.name, data, headers, getIdx);
        if (!prompt) continue;

        return {
          uuid:     uuid,
          rowNum:   rowNum,
          flowName: flow.name,
          titel:    String(row[getIdx('Titel')] || '').trim(),
          prompt:   prompt,
          trigger:  flow.trigger,
          output:   flow.output
        };
      }

      // Crash-Recovery: Trigger leer aber Vorgänger hat validen Output
      if (triggerVal === '' && f > 0) {
        var prevOutputVal = String(row[getIdx(flows[f-1].output)] || '').trim();
        if (prevOutputVal !== '' && isValidJson(prevOutputVal)) {
          if (flows[f-1].name === 'TRIAGE') {
            try {
              var cleanR = prevOutputVal.replace(/```json/g, '').replace(/```/g, '').trim();
              var fbR    = cleanR.indexOf('{');
              if (fbR > 0) cleanR = cleanR.substring(fbR);
              var relR   = (JSON.parse(cleanR).relevanz || '').trim();
              if (relR === 'Irrelevant' || relR === 'Niedrig' || relR === 'Mittel') break;
            } catch(e) { break; }
          }

          var itemKeyR = uuid + '|' + flow.name;
          if (skipKeys.indexOf(itemKeyR) >= 0) break;

          dashboard.getRange(rowNum, getIdx(flow.trigger) + 1).setValue('PENDING');
          SpreadsheetApp.flush();

          if (filterFlow && filterFlow !== 'ALLE' && flow.name !== filterFlow) break;

          var promptR = buildPrompt(uuid, flow.name, data, headers, getIdx);
          if (!promptR) break;

          return {
            uuid:     uuid,
            rowNum:   rowNum,
            flowName: flow.name,
            titel:    String(row[getIdx('Titel')] || '').trim(),
            prompt:   promptR,
            trigger:  flow.trigger,
            output:   flow.output
          };
        }
      }

      break;
    }
  }

  return null;
}

// ── buildPrompt ───────────────────────────────────────────────────
function buildPrompt(uuid, flowName, data, headers, getIdx) {
  var row = null;
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][getIdx('UUID')] || '').trim() === uuid) { row = data[r]; break; }
  }
  if (!row) return null;

  var get = function(col) { return String(row[getIdx(col)] || '').trim(); };

  var analyse     = get('Gemini Master Analyse');
  var prompts     = getPromptsFromSheet();
  var template    = prompts[flowName] || getDefaultPrompt(flowName);

  var fillVariables = function(t) {
    return t
      .replace(/\u200B/g, '').replace(/\u200C/g, '').replace(/\u200D/g, '')
      .replace(/\[Variable: Titel\]/g,                get('Titel'))
      .replace(/\[Variable: Inhalt\/Abstract\]/g,      get('Inhalt/Abstract'))
      .replace(/\[Variable: Volltext\/Extrakt\]/g,     get('Volltext/Extrakt'))
      .replace(/\[Variable: Volltext_Teil2\]/g,        get('Volltext_Teil2'))
      .replace(/\[Variable: Volltext_Teil3\]/g,        get('Volltext_Teil3'))
      .replace(/\[Variable: Autoren\]/g,               get('Autoren'))
      .replace(/\[Variable: DOI\]/g,                   get('DOI'))
      .replace(/\[Variable: Link\]/g,                  get('Link'))
      .replace(/\[Variable: Journal\/Quelle\]/g,       get('Journal/Quelle'))
      .replace(/\[Variable: Gemini Researcher\]/g,     get('Gemini Researcher'))
      .replace(/\[Variable: Gemini Triage\]/g,         get('Gemini Triage'))
      .replace(/\[Variable: Gemini Metadaten\]/g,      get('Gemini Metadaten'))
      .replace(/\[Variable: Gemini Master Analyse\]/g, analyse)
      .replace(/\[Variable: Gemini Redaktion\]/g,      get('Gemini Redaktion'))
      .replace(/\[Variable: Gemini Faktencheck\]/g,    get('Gemini Faktencheck'))
      .replace(/\[Variable: Haupterkenntnis\]/g,       getHaupterkenntnis(analyse))
      .replace(/\[Variable: Kernaussagen\]/g,          getKernaussagen(analyse));
  };

  return fillVariables(template);
}

function getKernaussagen(analyseJson) {
  try {
    var clean = analyseJson.replace(/```json/g, '').replace(/```/g, '').trim();
    var fb    = clean.indexOf('{');
    if (fb > 0) clean = clean.substring(fb);
    var arr = JSON.parse(clean).kernaussagen;
    return Array.isArray(arr) ? arr.join('\n- ') : String(arr || '');
  } catch(e) { return ''; }
}

function getHaupterkenntnis(analyseJson) {
  try {
    var clean = analyseJson.replace(/```json/g, '').replace(/```/g, '').trim();
    var fb    = clean.indexOf('{');
    if (fb > 0) clean = clean.substring(fb);
    return JSON.parse(clean).haupterkenntnis || '';
  } catch(e) { return ''; }
}

function getPromptsFromSheet() {
  var ss          = SpreadsheetApp.getActiveSpreadsheet();
  var promptSheet = ss.getSheetByName('Prompts');
  if (!promptSheet) return {};

  var prompts = {};
  var data    = promptSheet.getDataRange().getValues();
  for (var r = 0; r < data.length; r++) {
    var name    = String(data[r][0] || '').trim();
    var content = String(data[r][1] || '').trim();
    if (name && content) prompts[name] = content;
  }
  return prompts;
}

function getDefaultPrompt(flowName) {
  return '[Prompt für ' + flowName + ' nicht gefunden – bitte in "Prompts"-Sheet hinterlegen]\n\n' +
         'TITEL: [Variable: Titel]\nABSTRACT: [Variable: Inhalt/Abstract]';
}

// ── submitFlowResponse ────────────────────────────────────────────
function submitFlowResponse(uuid, flowName, triggerCol, outputCol, response) {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) return { success: false, error: 'Dashboard nicht gefunden' };

  var headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  var lastRow = dashboard.getLastRow();

  var getIdx = function(name) {
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c]).trim() === name) return c;
    }
    return -1;
  };

  var data   = dashboard.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var rowNum = -1;
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][getIdx('UUID')] || '').trim() === uuid) { rowNum = r + 2; break; }
  }
  if (rowNum < 0) return { success: false, error: 'UUID nicht gefunden' };

  var outIdx = getIdx(outputCol);
  if (outIdx < 0) return { success: false, error: 'Output-Spalte nicht gefunden: ' + outputCol };
  dashboard.getRange(rowNum, outIdx + 1).setValue(response);

  // Manuell-Tracking
  var manuelIdx = getIdx('Manuell_verarbeitet');
  if (manuelIdx >= 0) {
    var bisherige  = String(dashboard.getRange(rowNum, manuelIdx + 1).getValue() || '').trim();
    dashboard.getRange(rowNum, manuelIdx + 1).setValue(
      bisherige ? bisherige + ', ' + flowName : flowName
    );
  }

  // Trigger löschen
  var trigIdx = getIdx(triggerCol);
  if (trigIdx >= 0) dashboard.getRange(rowNum, trigIdx + 1).clearContent();

  try { applyResultsToDashboard(uuid, flowName, response); } catch(e) {
    Logger.log('applyResultsToDashboard Fehler: ' + e.message);
  }

  var statusCol  = getIdx('Status') + 1;
  var nextAction = determineNextAction(uuid, flowName, response, dashboard, headers, getIdx, rowNum, statusCol);

  SpreadsheetApp.flush();
  return { success: true, nextAction: nextAction };
}

// ── determineNextAction ───────────────────────────────────────────
function determineNextAction(uuid, flowName, response, dashboard, headers, getIdx, rowNum, statusCol) {
  var isValidJson = function(str) {
    if (!str || str.length < 10) return false;
    var s = str.replace(/```json/g, '').replace(/```/g, '').trim();
    var fb = s.indexOf('{');
    if (fb === -1) return false;
    s = s.substring(fb);
    try { JSON.parse(s); return true; } catch(e) { return false; }
  };

  if (flowName === 'TRIAGE' && isValidJson(response)) {
    try {
      var clean      = response.replace(/```json/g, '').replace(/```/g, '').trim();
      var fb         = clean.indexOf('{');
      if (fb > 0) clean = clean.substring(fb);
      var triageJson = JSON.parse(clean);
      var relevanz   = (triageJson.relevanz || '').trim();

      if (relevanz === 'Irrelevant') {
        dashboard.getRange(rowNum, statusCol).setValue('⛔ Irrelevant – Workflow gestoppt');
        return { action: 'SKIP_TO_NEXT', reason: 'Irrelevant' };
      }
      if (relevanz === 'Niedrig') {
        dashboard.getRange(rowNum, statusCol).setValue('🔻 Niedrig – Workflow gestoppt');
        return { action: 'SKIP_TO_NEXT', reason: 'Niedrig' };
      }
      if (relevanz === 'Mittel') {
        dashboard.getRange(rowNum, statusCol).setValue('⚠️ Mittel – Manuell fortsetzen?');
        return { action: 'SKIP_TO_NEXT', reason: 'Mittel' };
      }
      dashboard.getRange(rowNum, getIdx('Flow_Trigger_Metadaten') + 1).setValue('PENDING');
      return { action: 'NEXT_FLOW', flow: 'METADATEN' };
    } catch(e) {}
  }

  if (flowName === 'REVIEW') {
    dashboard.getRange(rowNum, statusCol).setValue('✅ Workflow komplett');
    return { action: 'COMPLETE' };
  }

  var nextFlowMap = {
    'RESEARCHER':     { trigger: 'Flow_Trigger_Triage',      flow: 'TRIAGE' },
    'METADATEN':      { trigger: 'Flow_Trigger_Analyse',     flow: 'MASTER_ANALYSE' },
    'MASTER_ANALYSE': { trigger: 'Flow_Trigger_Redaktion',   flow: 'REDAKTION' },
    'REDAKTION':      { trigger: 'Flow_Trigger_Faktencheck', flow: 'FAKTENCHECK' },
    'FAKTENCHECK':    { trigger: 'Flow_Trigger_Review',      flow: 'REVIEW' }
  };

  var next = nextFlowMap[flowName];
  if (next) {
    dashboard.getRange(rowNum, getIdx(next.trigger) + 1).setValue('PENDING');
    return { action: 'NEXT_FLOW', flow: next.flow };
  }
  return { action: 'UNKNOWN' };
}

// ── repairTriageInconsistencies ───────────────────────────────────
function repairTriageInconsistencies() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
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

  var isValidJson = function(str) {
    if (!str || str.length < 10) return false;
    var s = str.replace(/```json/g, '').replace(/```/g, '').trim();
    var fb = s.indexOf('{');
    if (fb === -1) return false;
    try { JSON.parse(s.substring(fb)); return true; } catch(e) { return false; }
  };

  var cleanTriggers = [
    'Flow_Trigger_Metadaten', 'Flow_Trigger_Analyse',
    'Flow_Trigger_Redaktion', 'Flow_Trigger_Faktencheck', 'Flow_Trigger_Review'
  ];

  var statusMap = {
    'Irrelevant': '⛔ Irrelevant – Workflow gestoppt',
    'Niedrig':    '🔻 Niedrig – Workflow gestoppt',
    'Mittel':     '⚠️ Mittel – Manuell fortsetzen?'
  };

  var repaired = 0, skipped = 0;

  for (var r = 0; r < data.length; r++) {
    var row          = data[r];
    var uuid         = String(row[getIdx('UUID')]          || '').trim();
    var rowNum       = r + 2;
    if (!uuid) continue;

    var triageOutput = String(row[getIdx('Gemini Triage')] || '').trim();
    if (!triageOutput || !isValidJson(triageOutput)) { skipped++; continue; }

    var relevanz = '';
    try {
      var clean = triageOutput.replace(/```json/g, '').replace(/```/g, '').trim();
      var fb    = clean.indexOf('{');
      if (fb > 0) clean = clean.substring(fb);
      relevanz  = (JSON.parse(clean).relevanz || '').trim();
    } catch(e) { skipped++; continue; }

    if (relevanz !== 'Irrelevant' && relevanz !== 'Niedrig' && relevanz !== 'Mittel') { skipped++; continue; }

    var needsClean = false;
    for (var ct = 0; ct < cleanTriggers.length; ct++) {
      if (String(row[getIdx(cleanTriggers[ct])] || '').trim() !== '') { needsClean = true; break; }
    }
    if (!needsClean) { skipped++; continue; }

    for (var ct2 = 0; ct2 < cleanTriggers.length; ct2++) {
      var ctIdx = getIdx(cleanTriggers[ct2]);
      if (ctIdx >= 0) dashboard.getRange(rowNum, ctIdx + 1).clearContent();
    }

    var currentStatus = String(row[getIdx('Status')] || '').trim();
    var correctStatus = statusMap[relevanz];
    if (currentStatus !== correctStatus) {
      dashboard.getRange(rowNum, getIdx('Status') + 1).setValue(correctStatus);
    }

    SpreadsheetApp.flush();
    repaired++;
    Logger.log('[REPAIR-TRIAGE] ' + uuid + ' | ' + relevanz + ' | Trigger bereinigt');
  }

  SpreadsheetApp.getUi().alert(
    '✅ Triage-Reparatur abgeschlossen\n\n' +
    '🔧 Repariert: ' + repaired + '\n' +
    '⏭️ Übersprungen: ' + skipped
  );
}

// ── getFlowStatusOverview ─────────────────────────────────────────
function getFlowStatusOverview() {
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) return null;

  var headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  var lastRow = dashboard.getLastRow();
  if (lastRow < 2) return null;

  var data = dashboard.getRange(2, 1, lastRow - 1, headers.length).getValues();

  var getIdx = function(name) {
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c]).trim() === name) return c;
    }
    return -1;
  };

  var isValidJson = function(str) {
    if (!str || str.length < 10) return false;
    var s = str.replace(/```json/g, '').replace(/```/g, '').trim();
    var fb = s.indexOf('{');
    if (fb === -1) return false;
    try { JSON.parse(s.substring(fb)); return true; } catch(e) { return false; }
  };

  var weekAgo = new Date(new Date().getTime() - 7 * 24 * 60 * 60 * 1000);

  var stats = {
    total: 0, brauchtResearcher: 0, brauchtTriage: 0,
    triageIrrelevant: 0, triageNiedrig: 0, triageMittel: 0,
    triageHoch: 0, triageSehrHoch: 0,
    brauchtMetadaten: 0, brauchtAnalyse: 0, brauchtRedaktion: 0,
    brauchtFaktencheck: 0, brauchtReview: 0,
    komplett: 0, fehler: 0,
    fehlerPerFlow: { Researcher: 0, Triage: 0, Metadaten: 0, 'Master Analyse': 0, Redaktion: 0, Faktencheck: 0, Review: 0 },
    hatVolltext: 0, nurAbstract: 0,
    fehlendeDoi: 0, fehlendePmid: 0,
    neuDieseWoche: 0
  };

  for (var r = 0; r < data.length; r++) {
    var row  = data[r];
    var uuid = String(row[getIdx('UUID')] || '').trim();
    if (!uuid) continue;
    stats.total++;

    var status      = String(row[getIdx('Status')]               || '').trim();
    var researcher  = String(row[getIdx('Gemini Researcher')]    || '').trim();
    var triage      = String(row[getIdx('Gemini Triage')]        || '').trim();
    var metadaten   = String(row[getIdx('Gemini Metadaten')]     || '').trim();
    var analyse     = String(row[getIdx('Gemini Master Analyse')]|| '').trim();
    var redaktion   = String(row[getIdx('Gemini Redaktion')]     || '').trim();
    var faktencheck = String(row[getIdx('Gemini Faktencheck')]   || '').trim();
    var review      = String(row[getIdx('Gemini Review')]        || '').trim();
    var volltext    = String(row[getIdx('Volltext/Extrakt')]      || '').trim();
    var abstract_   = String(row[getIdx('Inhalt/Abstract')]      || '').trim();
    var doi         = String(row[getIdx('DOI')]                  || '').trim();
    var pmid        = String(row[getIdx('PMID')]                 || '').trim();

    if (volltext.length > 100) stats.hatVolltext++;
    else if (abstract_.length > 50) stats.nurAbstract++;

    var inv = ['', 'n/a', '-', 'null', 'undefined'];
    if (inv.indexOf(doi.toLowerCase())  >= 0 || doi  === '') stats.fehlendeDoi++;
    if (inv.indexOf(pmid.toLowerCase()) >= 0 || pmid === '') stats.fehlendePmid++;

    var importTs = row[getIdx('Import-Timestamp')];
    if (importTs) {
      var d = new Date(importTs);
      if (!isNaN(d.getTime()) && d >= weekAgo) stats.neuDieseWoche++;
    }

    if (status.indexOf('dauerhafter Fehler') >= 0 || status.indexOf('❌') >= 0) {
      stats.fehler++;
      for (var fn in stats.fehlerPerFlow) {
        if (status.indexOf(fn) >= 0) { stats.fehlerPerFlow[fn]++; break; }
      }
      continue;
    }
    if (status === '✅ Workflow komplett')              { stats.komplett++;        continue; }
    if (status === '⛔ Irrelevant – Workflow gestoppt') { stats.triageIrrelevant++;continue; }
    if (status === '🔻 Niedrig – Workflow gestoppt')   { stats.triageNiedrig++;   continue; }
    if (status === '⚠️ Mittel – Manuell fortsetzen?')  { stats.triageMittel++;    continue; }

    if (!researcher || !isValidJson(researcher)) { stats.brauchtResearcher++; continue; }
    if (!triage     || !isValidJson(triage))     { stats.brauchtTriage++;     continue; }

    var relevanz = '';
    try {
      var cleanT = triage.replace(/```json/g,'').replace(/```/g,'').trim();
      var fbT    = cleanT.indexOf('{');
      if (fbT > 0) cleanT = cleanT.substring(fbT);
      relevanz   = (JSON.parse(cleanT).relevanz || '').trim();
    } catch(e) {}

    if (relevanz === 'Irrelevant') { stats.triageIrrelevant++; continue; }
    if (relevanz === 'Niedrig')    { stats.triageNiedrig++;    continue; }
    if (relevanz === 'Mittel')     { stats.triageMittel++;     continue; }
    if (relevanz === 'Hoch')       stats.triageHoch++;
    else if (relevanz === 'Sehr hoch') stats.triageSehrHoch++;

    if (!metadaten   || !isValidJson(metadaten))   { stats.brauchtMetadaten++;   continue; }
    if (!analyse     || !isValidJson(analyse))     { stats.brauchtAnalyse++;     continue; }
    if (!redaktion   || !isValidJson(redaktion))   { stats.brauchtRedaktion++;   continue; }
    if (!faktencheck || !isValidJson(faktencheck)) { stats.brauchtFaktencheck++; continue; }
    if (!review      || !isValidJson(review))      { stats.brauchtReview++;      continue; }

    stats.komplett++;
  }
  return stats;
}

// Einmalig ausführen: befüllt ""Volltext Bereinigt Länge"" für alle vorhandenen Papers
function repairBereinigtLaenge() {
  var ui        = SpreadsheetApp.getUi();
  var ss        = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) { ui.alert('Dashboard nicht gefunden!'); return; }

  var headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  var lastRow = dashboard.getLastRow();
  if (lastRow < 2) { ui.alert('Keine Papers vorhanden.'); return; }

  var getIdx = function(name) {
    for (var c = 0; c < headers.length; c++) {
      if (String(headers[c]).trim() === name) return c;
    }
    return -1;
  };

  var vt1Idx      = getIdx('Volltext/Extrakt');
  var vt2Idx      = getIdx('Volltext_Teil2');
  var vt3Idx      = getIdx('Volltext_Teil3');
  var berIdx      = getIdx('Volltext Bereinigt Länge');
  var origIdx     = getIdx('Volltext Original Länge');

  if (berIdx < 0) { ui.alert('""Volltext Bereinigt Länge"" Spalte nicht gefunden!'); return; }

  var data = dashboard.getRange(2, 1, lastRow - 1, headers.length).getValues();

  var updated  = 0;
  var skipped  = 0;

  for (var r = 0; r < data.length; r++) {
    var uuid = String(data[r][0] || '').trim();
    if (!uuid) continue;

    // Bereits befüllt → überspringen
    var existing = data[r][berIdx];
    if (existing && existing !== '' && existing !== 0) { skipped++; continue; }

    // Volltext zusammensetzen
    var vt1 = vt1Idx >= 0 ? String(data[r][vt1Idx] || '') : '';
    var vt2 = vt2Idx >= 0 ? String(data[r][vt2Idx] || '') : '';
    var vt3 = vt3Idx >= 0 ? String(data[r][vt3Idx] || '') : '';
    var volltext = (vt1 + vt2 + vt3).trim();

    if (!volltext) { skipped++; continue; }

    // Bereinigte Länge berechnen
    var bereinigt = cleanVolltextForLength(volltext);
    dashboard.getRange(r + 2, berIdx + 1).setValue(bereinigt);

    // Original-Länge nachfüllen falls leer
    if (origIdx >= 0 && (!data[r][origIdx] || data[r][origIdx] === 0)) {
      dashboard.getRange(r + 2, origIdx + 1).setValue(volltext.length);
    }

    updated++;

    // Alle 50 Zeilen flushen
    if (updated % 50 === 0) {
      SpreadsheetApp.flush();
      Logger.log('[REPAIR] ' + updated + ' Papers aktualisiert...');
    }
  }

  SpreadsheetApp.flush();
  ui.alert(
    'Volltext Bereinigt Länge – Reparatur abgeschlossen\n\n' +
    '✅ Aktualisiert: ' + updated + '\n' +
    '⏭️ Übersprungen:  ' + skipped
  );
}
