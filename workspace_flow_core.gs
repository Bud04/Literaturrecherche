// FILE: workspace_flow_core.gs
// ==========================================
// WORKSPACE FLOW CORE
// Integration mit Google Workspace Studio Flows
// ==========================================

const WORKSPACE_FLOW_MAPPING = {
  RESEARCHER: {
    name: "Researcher",
    triggerColumn: "Flow_Trigger_Researcher",
    outputColumn:  "Gemini Researcher",
    timeout: 120
  },
  TRIAGE: {
    name: "Triage",
    triggerColumn: "Flow_Trigger_Triage",
    outputColumn:  "Gemini Triage",
    timeout: 60
  },
  METADATEN: {
    name: "Metadaten",
    triggerColumn: "Flow_Trigger_Metadaten",
    outputColumn:  "Gemini Metadaten",
    timeout: 60
  },
  MASTER_ANALYSE: {
    name: "Master Analyse",
    triggerColumn: "Flow_Trigger_Analyse",
    outputColumn:  "Gemini Master Analyse",
    timeout: 180
  },
  REDAKTION: {
    name: "Redaktion",
    triggerColumn: "Flow_Trigger_Redaktion",
    outputColumn:  "Gemini Redaktion",
    timeout: 120
  },
  FAKTENCHECK: {
    name: "Faktencheck",
    triggerColumn: "Flow_Trigger_Faktencheck",
    outputColumn:  "Gemini Faktencheck",
    timeout: 180
  },
  REVIEW: {
    name: "Review",
    triggerColumn: "Flow_Trigger_Review",
    outputColumn:  "Gemini Review",
    timeout: 90
  }
};

// ==========================================
// POLL WORKFLOW PROGRESS
// Läuft jede Minute per Zeitbasiertem Trigger
// ==========================================

function pollWorkflowProgress() {
  // ── Lock: verhindert parallele Ausführungen ──────────────────────────────
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    Logger.log('[LOCK] Anderer pollWorkflowProgress läuft noch – abgebrochen');
    return;
  }
  try {
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

  var isValidJson = function(str) {
    if (!str || str.length < 10) return false;
    var s = str
      .replace(/```json/g, '')
      .replace(/```/g, '')
      .replace(/^\s+/, '')
      .replace(/\u00EF\u00BB\u00BF/, '')
      .trim();
    var firstBrace = s.indexOf('{');
    if (firstBrace === -1) return false;
    s = s.substring(firstBrace);
    try { JSON.parse(s); return true; }
    catch(e) { return false; }
  };

  var getRetryCount = function(uuid, flowName) {
    try {
      var val = getDashboardField(uuid, 'Fehler-Details') || '';
      var match = val.match(new RegExp(flowName + '_retry:(\\d+)'));
      return match ? parseInt(match[1]) : 0;
    } catch(e) { return 0; }
  };

  var incrementRetry = function(uuid, flowName) {
    try {
      var count = getRetryCount(uuid, flowName) + 1;
      var current = getDashboardField(uuid, 'Fehler-Details') || '';
      var newVal = current.replace(new RegExp(flowName + '_retry:\\d+'), '').trim();
      updateDashboardField(uuid, 'Fehler-Details', newVal + ' ' + flowName + '_retry:' + count);
      return count;
    } catch(e) { return 1; }
  };

  var clearRetry = function(uuid, flowName) {
    try {
      var current = getDashboardField(uuid, 'Fehler-Details') || '';
      var newVal = current.replace(new RegExp(flowName + '_retry:\\d+'), '').trim();
      updateDashboardField(uuid, 'Fehler-Details', newVal);
    } catch(e) {}
  };

  // ── FIX Bug 3: Alle Retry-Marker eines Flows löschen ─────────────
  var clearAllRetries = function(uuid, flowName) {
    try {
      var markers = [flowName + '_STUCK', flowName, 'AUTO_RETRY'];
      var current = getDashboardField(uuid, 'Fehler-Details') || '';
      var cleaned = current;
      for (var m = 0; m < markers.length; m++) {
        cleaned = cleaned.replace(new RegExp(markers[m] + '_retry:\\d+', 'g'), '').trim();
      }
      // Mehrfache Leerzeichen bereinigen
      cleaned = cleaned.replace(/\s{2,}/g, ' ').trim();
      if (cleaned !== current) updateDashboardField(uuid, 'Fehler-Details', cleaned);
    } catch(e) {}
  };

  var markError = function(rowNum, statusCol, message) {
    dashboard.getRange(rowNum, statusCol).setValue(message);
    dashboard.getRange(rowNum, 1, 1, dashboard.getLastColumn()).setBackground('#ffcccc');
  };

  var safeApply = function(uuid, type, json) {
    try {
      applyResultsToDashboard(uuid, type, json);
      SpreadsheetApp.flush();
    } catch(e) {
      Logger.log('[' + uuid + '] applyResultsToDashboard ' + type + ' Fehler: ' + e.message);
    }
  };

  // ================================================================
  // ✅ SCHRITT 0: AUTOMATISCHE NACHKONTROLLE + TIMEOUT-RETRY
  // ================================================================
  var repairChecks = [
    {
      label: 'Researcher→Triage',
      prevOutput: 'Gemini Researcher',
      nextTrigger: 'Flow_Trigger_Triage',
      nextOutput: 'Gemini Triage',
      errorStatus: '❌ Triage Fehler – Safety/Timeout',
      triggerCol: 'Flow_Trigger_Triage',
      stopStatuses: []
    },
    {
      label: 'Triage→Metadaten',
      prevOutput: 'Gemini Triage',
      nextTrigger: 'Flow_Trigger_Metadaten',
      nextOutput: 'Gemini Metadaten',
      errorStatus: '❌ Metadaten Fehler – Safety/Timeout',
      triggerCol: 'Flow_Trigger_Metadaten',
      stopStatuses: ['⛔ Irrelevant', '🔻 Niedrig', '⚠️ Mittel']
    },
    {
      label: 'Metadaten→Analyse',
      prevOutput: 'Gemini Metadaten',
      nextTrigger: 'Flow_Trigger_Analyse',
      nextOutput: 'Gemini Master Analyse',
      errorStatus: '❌ Master Analyse Fehler – Safety/Timeout',
      triggerCol: 'Flow_Trigger_Analyse',
      stopStatuses: []
    },
    {
      label: 'Analyse→Redaktion',
      prevOutput: 'Gemini Master Analyse',
      nextTrigger: 'Flow_Trigger_Redaktion',
      nextOutput: 'Gemini Redaktion',
      errorStatus: '❌ Redaktion Fehler – Safety/Timeout',
      triggerCol: 'Flow_Trigger_Redaktion',
      stopStatuses: []
    },
    {
      label: 'Redaktion→Faktencheck',
      prevOutput: 'Gemini Redaktion',
      nextTrigger: 'Flow_Trigger_Faktencheck',
      nextOutput: 'Gemini Faktencheck',
      errorStatus: '❌ Faktencheck Fehler – Safety/Timeout',
      triggerCol: 'Flow_Trigger_Faktencheck',
      stopStatuses: []
    },
    {
      label: 'Faktencheck→Review',
      prevOutput: 'Gemini Faktencheck',
      nextTrigger: 'Flow_Trigger_Review',
      nextOutput: 'Gemini Review',
      errorStatus: '❌ Review Fehler – Safety/Timeout',
      triggerCol: 'Flow_Trigger_Review',
      stopStatuses: []
    }
  ];

  var postTriageTriggers = [
    'Flow_Trigger_Metadaten', 'Flow_Trigger_Analyse',
    'Flow_Trigger_Redaktion', 'Flow_Trigger_Faktencheck', 'Flow_Trigger_Review'
  ];

  var statusMapRelevanz = {
    'Irrelevant': '⛔ Irrelevant – Workflow gestoppt',
    'Niedrig': '🔻 Niedrig – Workflow gestoppt',
    'Mittel': '⚠️ Mittel – Manuell fortsetzen?'
  };

  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var rowNum = r + 2;
    var uuid = String(row[getIdx('UUID')] || '').trim();
    var status = String(row[getIdx('Status')] || '').trim();

    if (!uuid) continue;
    if (status === '✅ Workflow komplett') continue;
    if (status.indexOf('❌') >= 0 && status.indexOf('Safety/Timeout') < 0 &&
        status.indexOf('max retries') < 0) continue;

    var statusCol = getIdx('Status') + 1;

    // ── Auto-Retry für Timeout/Spreadsheet-Fehler ──────────────────
    if (status.indexOf('❌') >= 0 && status.indexOf('Safety/Timeout') >= 0) {
      var autoRetries = incrementRetry(uuid, 'AUTO_RETRY');
      if (autoRetries <= 2) {
        Logger.log('[AUTO-RETRY] ' + uuid + ' | Versuch ' + autoRetries + ' | ' + status);
        for (var c = 0; c < repairChecks.length; c++) {
          var chk = repairChecks[c];
          if (status === chk.errorStatus) {
            dashboard.getRange(rowNum, statusCol).setValue('🔄 Auto-Retry läuft...');
            dashboard.getRange(rowNum, 1, 1, dashboard.getLastColumn()).setBackground('#fff3cd');
            dashboard.getRange(rowNum, getIdx(chk.triggerCol) + 1).setValue('PENDING');
            Logger.log('[AUTO-RETRY] ' + uuid + ' | ' + chk.label + ' → PENDING');
            data[r][getIdx('Status')] = '🔄 Auto-Retry läuft...';
            break;
          }
        }
      } else {
        Logger.log('[AUTO-RETRY] ' + uuid + ' | Max Auto-Retries erreicht → dauerhafter Fehler');
        dashboard.getRange(rowNum, statusCol).setValue(status.replace('Safety/Timeout', 'dauerhafter Fehler'));
      }
      continue;
    }

    // ── Auto-Retry für max retries Fehler ─────────────────────────
    if (status.indexOf('❌') >= 0 && status.indexOf('max retries') >= 0) {
      var autoRetries = incrementRetry(uuid, 'AUTO_RETRY');
      if (autoRetries <= 1) {
        Logger.log('[AUTO-RETRY] ' + uuid + ' | max retries reset → 1 weiterer Versuch');
        var flowMap = {
          'Researcher': { output: 'Gemini Researcher', trigger: 'Flow_Trigger_Researcher' },
          'Triage': { output: 'Gemini Triage', trigger: 'Flow_Trigger_Triage' },
          'Metadaten': { output: 'Gemini Metadaten', trigger: 'Flow_Trigger_Metadaten' },
          'Master Analyse': { output: 'Gemini Master Analyse', trigger: 'Flow_Trigger_Analyse' },
          'Redaktion': { output: 'Gemini Redaktion', trigger: 'Flow_Trigger_Redaktion' },
          'Faktencheck': { output: 'Gemini Faktencheck', trigger: 'Flow_Trigger_Faktencheck' },
          'Review': { output: 'Gemini Review', trigger: 'Flow_Trigger_Review' }
        };
        for (var flowName in flowMap) {
          if (status.indexOf(flowName) >= 0) {
            var fm = flowMap[flowName];
            clearRetry(uuid, flowName);
            clearRetry(uuid, flowName + '_STUCK');
            dashboard.getRange(rowNum, getIdx(fm.output) + 1).clearContent();
            dashboard.getRange(rowNum, getIdx(fm.trigger) + 1).setValue('PENDING');
            dashboard.getRange(rowNum, statusCol).setValue('🔄 Auto-Retry läuft...');
            dashboard.getRange(rowNum, 1, 1, dashboard.getLastColumn()).setBackground('#fff3cd');
            Logger.log('[AUTO-RETRY] ' + uuid + ' | ' + flowName + ' reset → PENDING');
            data[r][getIdx('Status')] = '🔄 Auto-Retry läuft...';
            break;
          }
        }
      } else {
        Logger.log('[AUTO-RETRY] ' + uuid + ' | Aufgegeben nach Auto-Retry');
        dashboard.getRange(rowNum, statusCol).setValue(status.replace('max retries', 'dauerhafter Fehler'));
      }
      continue;
    }

    // ── Nachkontrolle: fehlende PENDING auffüllen ──────────────────
    for (var c = 0; c < repairChecks.length; c++) {
      var chk = repairChecks[c];

      var isStopStatus = false;
      for (var s = 0; s < chk.stopStatuses.length; s++) {
        if (status.indexOf(chk.stopStatuses[s]) >= 0) { isStopStatus = true; break; }
      }
      if (isStopStatus) continue;

      // ── FIX Bug 1 (Schritt 0): Triage→Metadaten direkt per JSON prüfen ──
      if (chk.label === 'Triage→Metadaten') {
        var triageVal = String(row[getIdx('Gemini Triage')] || '').trim();
        if (triageVal && isValidJson(triageVal)) {
          try {
            var cleanTriage = triageVal.replace(/```json/g, '').replace(/```/g, '').trim();
            var fbTriage = cleanTriage.indexOf('{');
            if (fbTriage > 0) cleanTriage = cleanTriage.substring(fbTriage);
            var tjTriage = JSON.parse(cleanTriage);
            var relTriage = (tjTriage.relevanz || '').trim();
            if (relTriage === 'Irrelevant' || relTriage === 'Niedrig' || relTriage === 'Mittel') {
              var correctStatusT = statusMapRelevanz[relTriage];
              if (status !== correctStatusT) {
                dashboard.getRange(rowNum, statusCol).setValue(correctStatusT);
                data[r][getIdx('Status')] = correctStatusT;
              }
              continue; // Nicht zu Metadaten weitergehen
            }
          } catch(e) {}
        }
      }

      var prevVal = String(row[getIdx(chk.prevOutput)] || '').trim();
      var triggerVal = String(row[getIdx(chk.nextTrigger)] || '').trim();
      var nextVal = String(row[getIdx(chk.nextOutput)] || '').trim();

      if (prevVal !== '' && isValidJson(prevVal) &&
          triggerVal === '' && nextVal === '') {
        Logger.log('[REPAIR] ' + uuid + ' | ' + chk.label + ' → PENDING');
        dashboard.getRange(rowNum, getIdx(chk.nextTrigger) + 1).setValue('PENDING');
        data[r][getIdx(chk.nextTrigger)] = 'PENDING';
        break;
      }
    }
  }

  SpreadsheetApp.flush();

  // Daten neu laden nach Nachkontrolle
  data = dashboard.getRange(2, 1, lastRow - 1, headers.length).getValues();

  // ================================================================
  // ✅ SCHRITT 1: Trigger löschen wo Output vorhanden
  // ================================================================
  var flowChecks = [
    { trigger: 'Flow_Trigger_Researcher', output: 'Gemini Researcher' },
    { trigger: 'Flow_Trigger_Triage', output: 'Gemini Triage' },
    { trigger: 'Flow_Trigger_Metadaten', output: 'Gemini Metadaten' },
    { trigger: 'Flow_Trigger_Analyse', output: 'Gemini Master Analyse' },
    { trigger: 'Flow_Trigger_Redaktion', output: 'Gemini Redaktion' },
    { trigger: 'Flow_Trigger_Faktencheck', output: 'Gemini Faktencheck' },
    { trigger: 'Flow_Trigger_Review', output: 'Gemini Review' }
  ];

  for (var f = 0; f < flowChecks.length; f++) {
    var fc = flowChecks[f];
    var tIdx = getIdx(fc.trigger);
    var oIdx = getIdx(fc.output);
    if (tIdx < 0 || oIdx < 0) continue;

    var startCount = 0;
    for (var r = 0; r < data.length; r++) {
      var tVal = String(data[r][tIdx] || '').trim();
      var oVal = String(data[r][oIdx] || '').trim();

      if (oVal !== '' && (tVal === 'START' || tVal === 'PENDING')) {
        dashboard.getRange(r + 2, tIdx + 1).clearContent();
        data[r][tIdx] = '';
      }

      if (String(data[r][tIdx] || '').trim() === 'START' && oVal === '') {
        startCount++;
        if (startCount > 1) {
          dashboard.getRange(r + 2, tIdx + 1).clearContent();
          data[r][tIdx] = '';
        }
      }
    }
  }

  // ================================================================
  // ✅ SCHRITT 2: Aktive Flows erfassen
  // ================================================================
  var activeFlows = {
    'Flow_Trigger_Researcher': false,
    'Flow_Trigger_Triage': false,
    'Flow_Trigger_Metadaten': false,
    'Flow_Trigger_Analyse': false,
    'Flow_Trigger_Redaktion': false,
    'Flow_Trigger_Faktencheck': false,
    'Flow_Trigger_Review': false
  };

  for (var r = 0; r < data.length; r++) {
    for (var flowTrigger in activeFlows) {
      var idx = getIdx(flowTrigger);
      if (idx >= 0 && String(data[r][idx] || '').trim() === 'START') {
        activeFlows[flowTrigger] = true;
      }
    }
  }

  Logger.log('Aktive Flows: ' + JSON.stringify(activeFlows));

  // ================================================================
  // ✅ SCHRITT 3: PENDING → START befördern (Prioritätsmodus + Datumsreihenfolge)
  // ================================================================

  // ── Konfiguration lesen ──────────────────────────────────────────
  var prioritaetsModus = false;
  var flowReihenfolge  = 'ZEILE'; // ZEILE | DATUM_NEU | DATUM_ALT
  var konfSheet = ss.getSheetByName('Konfiguration');
  if (konfSheet) {
    var konfData = konfSheet.getDataRange().getValues();
    for (var k = 0; k < konfData.length; k++) {
      var konfKey = String(konfData[k][0]).trim();
      if (konfKey === 'PRIORITAETS_MODUS') {
        prioritaetsModus = String(konfData[k][1]).trim().toUpperCase() === 'EIN';
      }
      if (konfKey === 'FLOW_REIHENFOLGE') {
        var rfVal = String(konfData[k][1]).trim().toUpperCase();
        if (rfVal) flowReihenfolge = rfVal;
      }
    }
  }
  Logger.log('[PRIORITÄT] Modus: ' + (prioritaetsModus ? 'EIN' : 'AUS'));
  Logger.log('[REIHENFOLGE] ' + flowReihenfolge);

  // ── Sortierte Indexliste für PENDING→START ───────────────────────
  // DATUM_NEU = neueste zuerst (heute rückwärts)
  // DATUM_ALT = älteste zuerst
  // ZEILE     = natürliche Zeilenreihenfolge (Default)
  var sortedIndices = [];
  for (var si = 0; si < data.length; si++) sortedIndices.push(si);

  if (flowReihenfolge === 'DATUM_NEU' || flowReihenfolge === 'DATUM_ALT') {
    var datumIdx = getIdx('Publikationsdatum');
    if (datumIdx >= 0) {
      sortedIndices.sort(function(a, b) {
        var rawA = data[a][datumIdx];
        var rawB = data[b][datumIdx];
        var dA = rawA ? new Date(String(rawA).replace(/\//g, '-').substring(0, 10)) : new Date(0);
        var dB = rawB ? new Date(String(rawB).replace(/\//g, '-').substring(0, 10)) : new Date(0);
        if (isNaN(dA.getTime())) dA = new Date(0);
        if (isNaN(dB.getTime())) dB = new Date(0);
        return flowReihenfolge === 'DATUM_NEU' ? dB - dA : dA - dB;
      });
      Logger.log('[REIHENFOLGE] Sortiert (' + flowReihenfolge + '), erstes Datum: ' +
        String(data[sortedIndices[0]][datumIdx] || ''));
    }
  }

  var postTriageFlowTriggers = [
    'Flow_Trigger_Metadaten', 'Flow_Trigger_Analyse',
    'Flow_Trigger_Redaktion', 'Flow_Trigger_Faktencheck', 'Flow_Trigger_Review'
  ];

  for (var pf = 0; pf < flowChecks.length; pf++) {
    var pfc = flowChecks[pf];
    if (activeFlows[pfc.trigger]) continue;

    var pTrigIdx = getIdx(pfc.trigger);
    var pOutIdx = getIdx(pfc.output);
    if (pTrigIdx < 0 || pOutIdx < 0) continue;

    // ── Prioritätsmodus: Researcher/Triage zurückstellen ────────────
    // solange Hoch/Sehr-hoch Papers noch Folge-Flows offen haben
    if (prioritaetsModus &&
        (pfc.trigger === 'Flow_Trigger_Researcher' ||
         pfc.trigger === 'Flow_Trigger_Triage')) {

      var higherFlowsPending = false;

      for (var pr = 0; pr < data.length && !higherFlowsPending; pr++) {
        var prRelevanz = String(data[pr][getIdx('Relevanz')] || '').trim();
        if (prRelevanz !== 'Hoch' && prRelevanz !== 'Sehr hoch') continue;

        for (var pt = 0; pt < postTriageFlowTriggers.length; pt++) {
          var ptIdx = getIdx(postTriageFlowTriggers[pt]);
          if (ptIdx >= 0) {
            var ptVal = String(data[pr][ptIdx] || '').trim();
            if (ptVal === 'PENDING' || ptVal === 'START') {
              higherFlowsPending = true;
              Logger.log('[PRIORITÄT] ' + pfc.trigger + ' zurückgestellt – ' +
                String(data[pr][getIdx('UUID')] || '') + ' hat noch ' +
                postTriageFlowTriggers[pt] + ' offen');
              break;
            }
          }
        }

        if (!higherFlowsPending) {
          var metaOut = String(data[pr][getIdx('Gemini Metadaten')] || '').trim();
          if (metaOut && isValidJson(metaOut)) {
            var analyseOut = String(data[pr][getIdx('Gemini Master Analyse')] || '').trim();
            if (!analyseOut || !isValidJson(analyseOut)) {
              higherFlowsPending = true;
              Logger.log('[PRIORITÄT] ' + pfc.trigger + ' zurückgestellt – ' +
                String(data[pr][getIdx('UUID')] || '') + ' braucht noch Master Analyse');
            }
          }
        }
      }

      if (higherFlowsPending) continue;
    }

    // ── Nächstes PENDING Paper in sortierter Reihenfolge promoten ───
    for (var si = 0; si < sortedIndices.length; si++) {
      var r = sortedIndices[si];
      var pVal = String(data[r][pTrigIdx] || '').trim();
      var oVal = String(data[r][pOutIdx] || '').trim();
      if (pVal === 'PENDING' && oVal === '') {
        // ── Gestoppte/irrelevante Papers überspringen + PENDING wegräumen ──
        var pStatus = String(data[r][getIdx('Status')] || '').trim();
        if (pStatus.indexOf('Irrelevant') >= 0 ||
            pStatus.indexOf('Niedrig')    >= 0 ||
            pStatus.indexOf('Mittel')     >= 0 ||
            pStatus.indexOf('gestoppt')   >= 0 ||
            pStatus.indexOf('dauerhafter Fehler') >= 0) {
          dashboard.getRange(r + 2, pTrigIdx + 1).clearContent();
          data[r][pTrigIdx] = '';
          Logger.log('[SKIP] ' + String(data[r][getIdx('UUID')] || '') +
            ' | ' + pfc.trigger + ' gelöscht (Status: ' + pStatus + ')');
          continue;
        }
        // ───────────────────────────────────────────────────────────────────
        Logger.log('PENDING → START: ' + pfc.trigger + ' Zeile ' + (r + 2));
        dashboard.getRange(r + 2, pTrigIdx + 1).setValue('START');
        data[r][pTrigIdx] = 'START';
        activeFlows[pfc.trigger] = true;
        break;
      }
    }
  }

  // ================================================================
  // ✅ SCHRITT 4: Papers verarbeiten
  // ================================================================
  for (var r = 0; r < data.length; r++) {
    var row = data[r];
    var rowNum = r + 2;

    var uuid = String(row[getIdx('UUID')] || '').trim();
    if (!uuid) continue;

    var geminiResearcher = String(row[getIdx('Gemini Researcher')] || '').trim();
    var geminiTriage = String(row[getIdx('Gemini Triage')] || '').trim();
    var geminiMetadaten = String(row[getIdx('Gemini Metadaten')] || '').trim();
    var geminiAnalyse = String(row[getIdx('Gemini Master Analyse')] || '').trim();
    var geminiRedaktion = String(row[getIdx('Gemini Redaktion')] || '').trim();
    var geminiFaktencheck = String(row[getIdx('Gemini Faktencheck')] || '').trim();
    var geminiReview = String(row[getIdx('Gemini Review')] || '').trim();

    var triggerResearcher = String(row[getIdx('Flow_Trigger_Researcher')] || '').trim();
    var triggerTriage = String(row[getIdx('Flow_Trigger_Triage')] || '').trim();
    var triggerMetadaten = String(row[getIdx('Flow_Trigger_Metadaten')] || '').trim();
    var triggerAnalyse = String(row[getIdx('Flow_Trigger_Analyse')] || '').trim();
    var triggerRedaktion = String(row[getIdx('Flow_Trigger_Redaktion')] || '').trim();
    var triggerFaktencheck = String(row[getIdx('Flow_Trigger_Faktencheck')] || '').trim();
    var triggerReview = String(row[getIdx('Flow_Trigger_Review')] || '').trim();
    var status = String(row[getIdx('Status')] || '').trim();

    var statusCol = getIdx('Status') + 1;
    var triggerResCol = getIdx('Flow_Trigger_Researcher') + 1;
    var triggerTriaCol = getIdx('Flow_Trigger_Triage') + 1;
    var triggerMetaCol = getIdx('Flow_Trigger_Metadaten') + 1;
    var triggerAnalCol = getIdx('Flow_Trigger_Analyse') + 1;
    var triggerRedaCol = getIdx('Flow_Trigger_Redaktion') + 1;
    var triggerFaktCol = getIdx('Flow_Trigger_Faktencheck') + 1;
    var triggerRevCol = getIdx('Flow_Trigger_Review') + 1;

    if (status === '✅ Workflow komplett' ||
        status === '⛔ Irrelevant – Workflow gestoppt' ||
        status === '🔻 Niedrig – Workflow gestoppt' ||
        status === '⚠️ Mittel – Manuell fortsetzen?' ||
        status.indexOf('dauerhafter Fehler') >= 0 ||
        status.indexOf('❌') >= 0) continue;

    // ── FIX Bug 1 (Schritt 4): Harte Triage-Relevanz-Sperre ─────────
    // Läuft VOR allen anderen Checks — unabhängig vom Trigger-Zustand
    if (geminiTriage !== '' && isValidJson(geminiTriage)) {
      try {
        var cleanHard = geminiTriage.replace(/```json/g, '').replace(/```/g, '').trim();
        var fbHard = cleanHard.indexOf('{');
        if (fbHard > 0) cleanHard = cleanHard.substring(fbHard);
        var tjHard = JSON.parse(cleanHard);
        var relHard = (tjHard.relevanz || '').trim();
        if (relHard === 'Irrelevant' || relHard === 'Niedrig' || relHard === 'Mittel') {
          var hadWrongTrigger = false;
          for (var pt = 0; pt < postTriageTriggers.length; pt++) {
            if (String(row[getIdx(postTriageTriggers[pt])] || '').trim() !== '') {
              dashboard.getRange(rowNum, getIdx(postTriageTriggers[pt]) + 1).clearContent();
              hadWrongTrigger = true;
            }
          }
          dashboard.getRange(rowNum, statusCol).setValue(statusMapRelevanz[relHard]);
          if (hadWrongTrigger) SpreadsheetApp.flush();
          continue;
        }
      } catch(e) {}
    }

    // ── Researcher stuck ───────────────────────────────────────────
    if (triggerResearcher === 'START' && geminiResearcher === '') {
      var retries = incrementRetry(uuid, 'RESEARCHER_STUCK');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Researcher Fehler – Safety/Timeout');
        dashboard.getRange(rowNum, triggerResCol).clearContent();
      } else {
        Logger.log('[' + uuid + '] Researcher stuck re-trigger (' + retries + ')');
        dashboard.getRange(rowNum, triggerResCol).clearContent();
        SpreadsheetApp.flush();
        Utilities.sleep(500);
        dashboard.getRange(rowNum, triggerResCol).setValue('START');
      }
      continue;
    }

    // ── Triage stuck ───────────────────────────────────────────────
    if (triggerTriage === 'START' && geminiTriage === '') {
      var retries = incrementRetry(uuid, 'TRIAGE_STUCK');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Triage Fehler – Safety/Timeout');
        dashboard.getRange(rowNum, triggerTriaCol).clearContent();
      } else {
        Logger.log('[' + uuid + '] Triage stuck re-trigger (' + retries + ')');
        dashboard.getRange(rowNum, triggerTriaCol).clearContent();
        SpreadsheetApp.flush();
        Utilities.sleep(500);
        dashboard.getRange(rowNum, triggerTriaCol).setValue('START');
      }
      continue;
    }

    // ── Metadaten stuck ────────────────────────────────────────────
    if (triggerMetadaten === 'START' && geminiMetadaten === '') {
      var retries = incrementRetry(uuid, 'METADATEN_STUCK');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Metadaten Fehler – Safety/Timeout');
        dashboard.getRange(rowNum, triggerMetaCol).clearContent();
      } else {
        Logger.log('[' + uuid + '] Metadaten stuck re-trigger (' + retries + ')');
        dashboard.getRange(rowNum, triggerMetaCol).clearContent();
        SpreadsheetApp.flush();
        Utilities.sleep(500);
        dashboard.getRange(rowNum, triggerMetaCol).setValue('START');
      }
      continue;
    }

    // ── Master Analyse stuck ───────────────────────────────────────
    if (triggerAnalyse === 'START' && geminiAnalyse === '') {
      var retries = incrementRetry(uuid, 'ANALYSE_STUCK');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Master Analyse Fehler – Safety/Timeout');
        dashboard.getRange(rowNum, triggerAnalCol).clearContent();
      } else {
        Logger.log('[' + uuid + '] Analyse stuck re-trigger (' + retries + ')');
        dashboard.getRange(rowNum, triggerAnalCol).clearContent();
        SpreadsheetApp.flush();
        Utilities.sleep(500);
        dashboard.getRange(rowNum, triggerAnalCol).setValue('START');
      }
      continue;
    }

    // ── Redaktion stuck ────────────────────────────────────────────
    if (triggerRedaktion === 'START' && geminiRedaktion === '') {
      var retries = incrementRetry(uuid, 'REDAKTION_STUCK');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Redaktion Fehler – Safety/Timeout');
        dashboard.getRange(rowNum, triggerRedaCol).clearContent();
      } else {
        Logger.log('[' + uuid + '] Redaktion stuck re-trigger (' + retries + ')');
        dashboard.getRange(rowNum, triggerRedaCol).clearContent();
        SpreadsheetApp.flush();
        Utilities.sleep(500);
        dashboard.getRange(rowNum, triggerRedaCol).setValue('START');
      }
      continue;
    }

    // ── Faktencheck stuck ──────────────────────────────────────────
    if (triggerFaktencheck === 'START' && geminiFaktencheck === '') {
      var retries = incrementRetry(uuid, 'FAKTENCHECK_STUCK');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Faktencheck Fehler – Safety/Timeout');
        dashboard.getRange(rowNum, triggerFaktCol).clearContent();
      } else {
        Logger.log('[' + uuid + '] Faktencheck stuck re-trigger (' + retries + ')');
        dashboard.getRange(rowNum, triggerFaktCol).clearContent();
        SpreadsheetApp.flush();
        Utilities.sleep(500);
        dashboard.getRange(rowNum, triggerFaktCol).setValue('START');
      }
      continue;
    }

    // ── Review stuck ───────────────────────────────────────────────
    if (triggerReview === 'START' && geminiReview === '') {
      var retries = incrementRetry(uuid, 'REVIEW_STUCK');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Review Fehler – Safety/Timeout');
        dashboard.getRange(rowNum, triggerRevCol).clearContent();
      } else {
        Logger.log('[' + uuid + '] Review stuck re-trigger (' + retries + ')');
        dashboard.getRange(rowNum, triggerRevCol).clearContent();
        SpreadsheetApp.flush();
        Utilities.sleep(500);
        dashboard.getRange(rowNum, triggerRevCol).setValue('START');
      }
      continue;
    }

    // ── Researcher kein JSON → re-trigger ─────────────────────────
    if (geminiResearcher !== '' && !isValidJson(geminiResearcher) &&
        triggerTriage === '' && geminiTriage === '') {
      var retries = incrementRetry(uuid, 'RESEARCHER');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Researcher Fehler – max retries');
      } else {
        dashboard.getRange(rowNum, getIdx('Gemini Researcher') + 1).clearContent();
        dashboard.getRange(rowNum, triggerResCol).setValue('PENDING');
      }
      continue;
    }

    // ── Researcher fertig → Triage PENDING ────────────────────────
    if (geminiResearcher !== '' && isValidJson(geminiResearcher) &&
        triggerTriage === '' && geminiTriage === '') {
      Logger.log('[' + uuid + '] Researcher fertig → Triage PENDING');
      clearAllRetries(uuid, 'RESEARCHER'); // ← FIX Bug 3
      dashboard.getRange(rowNum, triggerTriaCol).setValue('PENDING');
      continue;
    }

    // ── Triage kein JSON → re-trigger ─────────────────────────────
    if (geminiTriage !== '' && !isValidJson(geminiTriage) &&
        triggerMetadaten === '' && geminiMetadaten === '') {
      var retries = incrementRetry(uuid, 'TRIAGE');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Triage Fehler – max retries');
      } else {
        dashboard.getRange(rowNum, getIdx('Gemini Triage') + 1).clearContent();
        dashboard.getRange(rowNum, triggerTriaCol).setValue('PENDING');
      }
      continue;
    }

    // ── Triage fertig → Relevanz prüfen ───────────────────────────
    if (geminiTriage !== '' && isValidJson(geminiTriage) &&
        triggerMetadaten === '' && geminiMetadaten === '') {
      var relevanz = '';
      try {
        var clean = geminiTriage.replace(/```json/g, '').replace(/```/g, '').trim();
        var firstBrace = clean.indexOf('{');
        if (firstBrace > 0) clean = clean.substring(firstBrace);
        var triageJson = JSON.parse(clean);
        relevanz = (triageJson.relevanz || '').trim();
      } catch(e) {
        Logger.log('[' + uuid + '] Triage JSON Fehler: ' + e.message);
        relevanz = 'Unbekannt';
      }

      safeApply(uuid, 'TRIAGE', geminiTriage);
      clearAllRetries(uuid, 'TRIAGE'); // ← FIX Bug 3

      if (relevanz === 'Irrelevant') {
        dashboard.getRange(rowNum, statusCol).setValue('⛔ Irrelevant – Workflow gestoppt');
      } else if (relevanz === 'Niedrig') {
        dashboard.getRange(rowNum, statusCol).setValue('🔻 Niedrig – Workflow gestoppt');
      } else if (relevanz === 'Mittel') {
        dashboard.getRange(rowNum, statusCol).setValue('⚠️ Mittel – Manuell fortsetzen?');
      } else {
        Logger.log('[' + uuid + '] Relevanz ""' + relevanz + '"" → Metadaten PENDING');
        dashboard.getRange(rowNum, triggerMetaCol).setValue('PENDING');
      }
      continue;
    }

    // ── Metadaten kein JSON → re-trigger ──────────────────────────
    if (geminiMetadaten !== '' && !isValidJson(geminiMetadaten) &&
        triggerAnalyse === '' && geminiAnalyse === '') {
      var retries = incrementRetry(uuid, 'METADATEN');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Metadaten Fehler – max retries');
      } else {
        dashboard.getRange(rowNum, getIdx('Gemini Metadaten') + 1).clearContent();
        dashboard.getRange(rowNum, triggerMetaCol).setValue('PENDING');
      }
      continue;
    }

    // ── Metadaten fertig → Master Analyse PENDING ─────────────────
    if (geminiMetadaten !== '' && isValidJson(geminiMetadaten) &&
        triggerAnalyse === '' && geminiAnalyse === '') {
      safeApply(uuid, 'METADATEN', geminiMetadaten);
      clearAllRetries(uuid, 'METADATEN'); // ← FIX Bug 3
      Logger.log('[' + uuid + '] Metadaten fertig → Master Analyse PENDING');
      dashboard.getRange(rowNum, triggerAnalCol).setValue('PENDING');
      continue;
    }

    // ── Master Analyse kein JSON → re-trigger ─────────────────────
    if (geminiAnalyse !== '' && !isValidJson(geminiAnalyse) &&
        triggerRedaktion === '' && geminiRedaktion === '') {
      var retries = incrementRetry(uuid, 'MASTER_ANALYSE');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Master Analyse Fehler – max retries');
      } else {
        dashboard.getRange(rowNum, getIdx('Gemini Master Analyse') + 1).clearContent();
        dashboard.getRange(rowNum, triggerAnalCol).setValue('PENDING');
      }
      continue;
    }

    // ── Master Analyse fertig → Redaktion PENDING ─────────────────
    if (geminiAnalyse !== '' && isValidJson(geminiAnalyse) &&
        triggerRedaktion === '' && geminiRedaktion === '') {
      safeApply(uuid, 'MASTER_ANALYSE', geminiAnalyse);
      clearAllRetries(uuid, 'MASTER_ANALYSE'); // ← FIX Bug 3
      Logger.log('[' + uuid + '] Master Analyse fertig → Redaktion PENDING');
      dashboard.getRange(rowNum, triggerRedaCol).setValue('PENDING');
      continue;
    }

    // ── Redaktion kein JSON → re-trigger ──────────────────────────
    if (geminiRedaktion !== '' && !isValidJson(geminiRedaktion) &&
        triggerFaktencheck === '' && geminiFaktencheck === '') {
      var retries = incrementRetry(uuid, 'REDAKTION');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Redaktion Fehler – max retries');
      } else {
        dashboard.getRange(rowNum, getIdx('Gemini Redaktion') + 1).clearContent();
        dashboard.getRange(rowNum, triggerRedaCol).setValue('PENDING');
      }
      continue;
    }

    // ── Redaktion fertig → Faktencheck PENDING ────────────────────
    if (geminiRedaktion !== '' && isValidJson(geminiRedaktion) &&
        triggerFaktencheck === '' && geminiFaktencheck === '') {
      clearAllRetries(uuid, 'REDAKTION'); // ← FIX Bug 3
      Logger.log('[' + uuid + '] Redaktion fertig → Faktencheck PENDING');
      dashboard.getRange(rowNum, triggerFaktCol).setValue('PENDING');
      continue;
    }

    // ── Faktencheck kein JSON → re-trigger ────────────────────────
    if (geminiFaktencheck !== '' && !isValidJson(geminiFaktencheck) &&
        triggerReview === '' && geminiReview === '') {
      var retries = incrementRetry(uuid, 'FAKTENCHECK');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Faktencheck Fehler – max retries');
      } else {
        dashboard.getRange(rowNum, getIdx('Gemini Faktencheck') + 1).clearContent();
        dashboard.getRange(rowNum, triggerFaktCol).setValue('PENDING');
      }
      continue;
    }

    // ── Faktencheck fertig → Review PENDING ───────────────────────
    if (geminiFaktencheck !== '' && isValidJson(geminiFaktencheck) &&
        triggerReview === '' && geminiReview === '') {
      safeApply(uuid, 'FAKTENCHECK', geminiFaktencheck);
      clearAllRetries(uuid, 'FAKTENCHECK'); // ← FIX Bug 3
      Logger.log('[' + uuid + '] Faktencheck fertig → Review PENDING');
      dashboard.getRange(rowNum, triggerRevCol).setValue('PENDING');
      continue;
    }

    // ── Review kein JSON → re-trigger ─────────────────────────────
    if (geminiReview !== '' && !isValidJson(geminiReview)) {
      var retries = incrementRetry(uuid, 'REVIEW');
      if (retries > 3) {
        markError(rowNum, statusCol, '❌ Review Fehler – max retries');
      } else {
        dashboard.getRange(rowNum, getIdx('Gemini Review') + 1).clearContent();
        dashboard.getRange(rowNum, triggerRevCol).setValue('PENDING');
      }
      continue;
    }

    // ── Review fertig → Workflow komplett ─────────────────────────
    if (geminiReview !== '' && isValidJson(geminiReview)) {
      safeApply(uuid, 'REVIEW', geminiReview);
      clearAllRetries(uuid, 'REVIEW'); // ← FIX Bug 3
      dashboard.getRange(rowNum, statusCol).setValue('✅ Workflow komplett');
      Logger.log('[' + uuid + '] ✅ Workflow komplett');
      continue;
    }
  }
  } finally {
    lock.releaseLock();
  }
}


// ==========================================
// MANUELLE FORTSETZUNG (Mittel-Relevanz)
// ==========================================

function continueWorkflowChainFromMetadata() {
  var ui = SpreadsheetApp.getUi();
  var dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  var selection = dashboard.getActiveRange();

  if (!selection || selection.getRow() < 2) {
    ui.alert('Bitte ein Paper im Dashboard auswählen.');
    return;
  }

  var startRow = selection.getRow();
  var numRows  = selection.getNumRows();
  var headers  = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];

  var triggerMetaCol = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === 'Flow_Trigger_Metadaten') {
      triggerMetaCol = c + 1;
      break;
    }
  }

  if (triggerMetaCol === -1) {
    ui.alert('Flow_Trigger_Metadaten Spalte nicht gefunden!');
    return;
  }

  var uuids = [];
  for (var i = 0; i < numRows; i++) {
    var uuid = dashboard.getRange(startRow + i, 1).getValue();
    if (uuid) uuids.push({ uuid: uuid, row: startRow + i });
  }

  if (uuids.length === 0) {
    ui.alert('Keine gültigen Papers ausgewählt.');
    return;
  }

  var confirm = ui.alert(
    'Workflow fortsetzen',
    uuids.length + ' Paper(s) ausgewählt.\n\nAb Metadaten-Flow fortsetzen?\n(Researcher + Triage werden NICHT nochmal ausgeführt)',
    ui.ButtonSet.YES_NO
  );
  if (confirm !== ui.Button.YES) return;

  for (var j = 0; j < uuids.length; j++) {
    dashboard.getRange(uuids[j].row, triggerMetaCol).setValue('START');
    Logger.log('Manuell fortgesetzt: ' + uuids[j].uuid);
  }

  ui.alert('✅ ' + uuids.length + ' Paper(s) werden ab Metadaten fortgesetzt.\npollWorkflowProgress übernimmt automatisch.');
}

// ==========================================
// ERGEBNISSE INS DASHBOARD SCHREIBEN
// ==========================================

function applyResultsToDashboard(uuid, flowName, jsonString) {
  var data = {};
  try {
    var clean = jsonString.replace(/```json/g, '').replace(/```/g, '').trim();
    data = JSON.parse(clean);
  } catch (e) {
    Logger.log('[' + uuid + '] JSON Parse Fehler für ' + flowName + ': ' + e.message);
    updateDashboardField(uuid, 'Fehler-Details', flowName + ' JSON Fehler: ' + e.message);
    return;
  }

  var mappings = {

    'TRIAGE': {
      'relevanz':                  'Relevanz',
      'relevanz_begruendung':      'Relevanz-Begründung',
      'evidenzgrad':               'Evidenzgrad',
      'kategorie.hauptkategorie':  'Hauptkategorie',
      'kategorie.unterkategorien': 'Unterkategorien',
      'kategorie.schlagwoerter':   'Schlagwörter'
    },

    'METADATEN': {
      'studientyp':   'Artikeltyp/Studientyp',
      'volume':       'Volume',
      'issue':        'Issue',
      'pages':        'Pages',
      'doi_string':   'DOI',
      'pmid':         'PMID',
      'journal_name': 'Journal/Quelle'
    },

    'MASTER_ANALYSE': {
      'relevanz':                  'Relevanz',
      'relevanz_begruendung':      'Relevanz-Begründung',
      'haupterkenntnis':           'Haupterkenntnis',
      'kernaussagen':              'Kernaussagen',
      'zusammenfassung':           'Zusammenfassung',
      'praktische_implikationen':  'Praktische Implikationen',
      'evidenzgrad':               'Evidenzgrad',
      'kritische_bewertung':       'Kritische Bewertung',
      'action_required':           'Review-Status',
      'produkt_fokus':             'Produkt-Fokus',
      'kategorie.hauptkategorie':  'Hauptkategorie',
      'kategorie.unterkategorien': 'Unterkategorien',
      'schlagwoerter':             'Schlagwörter',
      'pico.population':           'PICO Population',
      'pico.intervention':         'PICO Intervention',
      'pico.comparator':           'PICO Comparator',
      'pico.outcomes':             'PICO Outcomes'
    },

    'FAKTENCHECK': {
     'is_supported': 'Batch-Phase',
     'warnung': 'Kritische Bewertung'
    },

    'REVIEW': {
      'freigabe_kommentar':                    'Relevanz-Begründung',
      'korrigiertes_json.relevanz':            'Relevanz',
      'korrigiertes_json.haupterkenntnis':     'Haupterkenntnis',
      'korrigiertes_json.zusammenfassung':     'Zusammenfassung',
      'korrigiertes_json.action_required':     'Review-Status',
      'korrigiertes_json.kritische_bewertung': 'Kritische Bewertung',
      'korrigiertes_json.produkt_fokus':       'Produkt-Fokus',
      'korrigiertes_json.hauptkategorie':      'Hauptkategorie'
    }
  };

  var fieldMap = mappings[flowName];
  if (!fieldMap) {
    Logger.log('[' + uuid + '] Kein Mapping für Flow: ' + flowName);
    return;
  }

  for (var jsonPath in fieldMap) {
    var dashboardCol = fieldMap[jsonPath];
    var value = getNestedValue(data, jsonPath);
    if (value === null || value === undefined || value === '') continue;

    if (Array.isArray(value)) {
      value = value.join(', ');
    } else if (typeof value === 'boolean') {
      value = value ? 'Ja' : 'Nein';
    } else if (typeof value === 'object') {
      value = JSON.stringify(value);
    }

    updateDashboardField(uuid, dashboardCol, String(value));
  }

  Logger.log('[' + uuid + '] ' + flowName + ' → Felder geschrieben');
}

function getNestedValue(obj, path) {
  var parts = path.split('.');
  var current = obj;
  for (var i = 0; i < parts.length; i++) {
    if (current === null || current === undefined) return null;
    current = current[parts[i]];
  }
  return current;
}

// ==========================================
// DASHBOARD HELPER FUNKTIONEN
// ==========================================

function updateDashboardField(uuid, columnName, value) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) return;

  var headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  var colIndex = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === columnName.trim()) {
      colIndex = c + 1;
      break;
    }
  }

  if (colIndex === -1) {
    Logger.log('updateDashboardField: Spalte nicht gefunden: ' + columnName);
    return;
  }

  var data = dashboard.getRange(2, 1, dashboard.getLastRow() - 1, 1).getValues();
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][0]).trim() === String(uuid).trim()) {
      dashboard.getRange(r + 2, colIndex).setValue(value);
      return;
    }
  }

  Logger.log('updateDashboardField: UUID nicht gefunden: ' + uuid);
}

function getDashboardField(uuid, columnName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) return null;

  var headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  var colIndex = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === columnName.trim()) {
      colIndex = c + 1;
      break;
    }
  }

  if (colIndex === -1) {
    Logger.log('getDashboardField: Spalte nicht gefunden: ' + columnName);
    return null;
  }

  var data = dashboard.getRange(2, 1, dashboard.getLastRow() - 1, 1).getValues();
  for (var r = 0; r < data.length; r++) {
    if (String(data[r][0]).trim() === String(uuid).trim()) {
      return dashboard.getRange(r + 2, colIndex).getValue();
    }
  }

  Logger.log('getDashboardField: UUID nicht gefunden: ' + uuid);
  return null;
}

// ==========================================
// FLOW STATUS ANZEIGE
// ==========================================

function showWorkspaceFlowStatus() {
  var ui = SpreadsheetApp.getUi();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dashboard = ss.getSheetByName('Dashboard');
  if (!dashboard) { ui.alert('Dashboard nicht gefunden'); return; }

  var lastRow = dashboard.getLastRow();
  if (lastRow < 2) { ui.alert('Keine Papers vorhanden'); return; }

  var headers = dashboard.getRange(1, 1, 1, dashboard.getLastColumn()).getValues()[0];
  var data    = dashboard.getRange(2, 1, lastRow - 1, headers.length).getValues();

  var statusCol = -1;
  for (var c = 0; c < headers.length; c++) {
    if (String(headers[c]).trim() === 'Status') { statusCol = c; break; }
  }

  var total = 0, komplett = 0, irrelevant = 0, niedrig = 0,
      mittel = 0, laufend = 0, wartend = 0;

  for (var r = 0; r < data.length; r++) {
    var uuid = String(data[r][0] || '').trim();
    if (!uuid) continue;
    total++;
    var st = statusCol >= 0 ? String(data[r][statusCol] || '').trim() : '';
    if (st === '✅ Workflow komplett')              komplett++;
    else if (st === '⛔ Irrelevant – Workflow gestoppt') irrelevant++;
    else if (st === '🔻 Niedrig – Workflow gestoppt')    niedrig++;
    else if (st === '⚠️ Mittel – Manuell fortsetzen?')   mittel++;
    else if (st.indexOf('❌') >= 0)                      laufend++;
    else                                                  wartend++;
  }

  var message =
    '=== WORKSPACE FLOW STATUS ===\n\n' +
    '📊 Gesamt Papers: '      + total     + '\n' +
    '✅ Workflow komplett: '   + komplett  + '\n' +
    '⏳ Wartend/Laufend: '    + wartend   + '\n' +
    '⚠️ Mittel (manuell): '   + mittel    + '\n' +
    '🔻 Niedrig (gestoppt): ' + niedrig   + '\n' +
    '⛔ Irrelevant: '          + irrelevant + '\n' +
    '❌ Fehler: '              + laufend   + '\n\n' +
    'Tipp: Mittel-Papers manuell fortsetzen über Menü.';

  ui.alert('Flow Status', message, ui.ButtonSet.OK);
}

// ==========================================
// TOKEN CHECK
// ==========================================

function checkTokensBeforeFlow(uuid, flowName) {
  var volltextFlows = ['MASTER_ANALYSE', 'REDAKTION', 'FAKTENCHECK'];
  if (volltextFlows.indexOf(flowName) < 0) return { ok: true };

  var teil1 = getDashboardField(uuid, 'Volltext/Extrakt') || '';
  var teil2 = getDashboardField(uuid, 'Volltext_Teil2')   || '';
  var teil3 = getDashboardField(uuid, 'Volltext_Teil3')   || '';
  var volltext = (teil1 + teil2 + teil3).trim();
  var text = volltext || getDashboardField(uuid, 'Inhalt/Abstract') || '';

  if (!text || text.length < 100) {
    return { ok: false, error: 'Kein Volltext/Abstract verfügbar' };
  }

  var estimatedTokens = text.length / 4;
  var TOKEN_LIMIT = 29000;

  if (estimatedTokens > TOKEN_LIMIT) {
    return {
      ok: false,
      error: 'Volltext zu groß (' + Math.round(estimatedTokens) + ' Tokens)',
      needsManual: true
    };
  }

  return { ok: true, tokens: estimatedTokens };
}
