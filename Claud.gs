// FILE: Claude.gs
// ==========================================
// CLAUDE API INTEGRATION
// NUR Claude-spezifische Funktionen!
// Alle PubMed-Funktionen → pubmed_import.gs
// Alle Shared Helpers → helpers_shared.gs
// ==========================================

// ==========================================
// SETUP & HILFE
// ==========================================

/**
 * ✅ Claude API Setup
 */
function showClaudeAPISetup() {
  const ui = SpreadsheetApp.getUi();
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');

  let message = '=== CLAUDE API EINRICHTUNG ===\n\n';

  if (apiKey) {
    message += '✅ API Key ist eingerichtet!\n\n';
    message += `Key: ${apiKey.substring(0, 10)}...${apiKey.substring(apiKey.length - 4)}\n\n`;
    message += 'Du kannst jetzt Claude nutzen!';
  } else {
    message += '❌ API Key noch nicht eingerichtet\n\n';
    message += 'SETUP:\n\n';
    message += '1. Gehe zu: https://console.anthropic.com/\n';
    message += '2. Account erstellen/anmelden\n';
    message += '3. API Keys → "Create Key"\n';
    message += '4. Key kopieren\n\n';
    message += '5. In diesem Sheet:\n';
    message += '   • Erweiterungen → Apps Script\n';
    message += '   • Projekt-Einstellungen ⚙️\n';
    message += '   • Script Properties\n';
    message += '   • Property hinzufügen:\n';
    message += '     Name: CLAUDE_API_KEY\n';
    message += '     Wert: [dein API Key]\n\n';
    message += '──────────────────────\n\n';
    message += '💰 KOSTEN:\n';
    message += '• Haiku: ~$0.40/1M Tokens\n';
    message += '• Sonnet: ~$4/1M Tokens\n';
    message += '• Opus: ~$20/1M Tokens\n\n';
    message += 'Ein Paper (komplett): ~$0.05-0.15\n';
    message += '100 Papers: ~$5-15\n\n';
    message += '💡 TIPP: Erst mit kleiner Menge testen!';
  }

  ui.alert('Claude API Setup', message, ui.ButtonSet.OK);
}

/**
 * ✅ Claude Hilfe
 */
function showClaudeHelp() {
  const ui = SpreadsheetApp.getUi();

  const message =
    '=== CLAUDE HILFE ===\n\n' +
    '🎯 WORKFLOW:\n\n' +
    '1️⃣ ERSTE ANALYSE:\n' +
    '  → "Komplette Analyse (alle unbearbeitet)"\n' +
    '  → STOPP bei niedriger/mittlerer Relevanz!\n\n' +
    '2️⃣ GEZIELT ERGÄNZEN:\n' +
    '  → Markiere Papers im Dashboard\n' +
    '  → "Intelligente Ergänzung (markiert)"\n' +
    '  → Füllt nur fehlende Schritte\n\n' +
    '🛡️ RELEVANZ-FILTER:\n\n' +
    '• Nach Schritt 1 (Triage):\n' +
    '  → Relevanz = Niedrig? → STOPP!\n' +
    '  → Relevanz = Hoch? → Weiter!\n\n' +
    '💰 KOSTEN SPAREN:\n\n' +
    '1. Erst nur Schritt 1 (Triage)\n' +
    '2. Dann nur hohe Relevanz analysieren\n' +
    '3. Kosten-Rechner nutzen!';

  ui.alert('Claude Hilfe', message, ui.ButtonSet.OK);
}

// ==========================================
// WORKFLOW-LOGIK (Claude-spezifisch)
// ==========================================

/**
 * ✅ Prüft ob Schritt ausgeführt werden soll
 */
function shouldExecuteStep(uuid, stepNumber) {
  const data = getDashboardDataByUUID(uuid);
  if (!data) return false;

  const stepFields = {
    0: 'Researcher',
    1: 'Relevanz',
    2: 'Tags',
    3: 'Haupterkenntnis',
    4: 'Zusammenfassung',
    5: 'Faktencheck',
    6: 'Review-Status'
  };

  const field = stepFields[stepNumber];
  if (!field) return false;

  return !data[field] || data[field] === '' || data[field] === 'UNGEPRÜFT';
}

/**
 * ✅ Findet fehlende Schritte für Paper
 */
function getMissingSteps(uuid) {
  const data = getDashboardDataByUUID(uuid);
  if (!data) return {};

  return {
    0: !data.Researcher || data.Researcher === '',
    1: !data.Relevanz || data.Relevanz === '',
    2: !data.Tags || data.Tags === '',
    3: !data.Haupterkenntnis || data.Haupterkenntnis === '',
    4: !data.Zusammenfassung || data.Zusammenfassung === '',
    5: !data.Faktencheck || data.Faktencheck === '',
    6: !data['Review-Status'] || data['Review-Status'] === 'UNGEPRÜFT'
  };
}

/**
 * ✅ Findet Papers wo bestimmter Schritt fehlt
 */
function getPapersWithMissingStep(stepNumber) {
  return getAllPaperUUIDs().filter(uuid => {
    const missing = getMissingSteps(uuid);
    return missing[stepNumber] === true;
  });
}

/**
 * ✅ Holt Papers mit bestimmtem Status
 * HINWEIS: getPapersForAnalysis() → gemini_api.gs
 */
function getPapersWithStatus(status) {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  if (!dashboard || dashboard.getLastRow() <= 1) return [];

  const data = dashboard.getDataRange().getValues();
  const headers = data[0];
  const statusCol = headers.indexOf('Status');
  if (statusCol === -1) return [];

  const uuids = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][statusCol] === status) {
      uuids.push(data[i][0]);
    }
  }
  return uuids;
}

/**
 * ✅ Holt ausgewählte Papers für Analyse
 * HINWEIS: getSelectedPapersForAnalysis() (mit vollem Objekt) → gemini_api.gs
 */
function getSelectedPaperUUIDs() {
  const dashboard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Dashboard');
  if (!dashboard) return [];

  const selection = dashboard.getActiveRange();
  if (!selection || selection.getRow() < 2) return [];

  const startRow = selection.getRow();
  const numRows = selection.getNumRows();
  const uuids = [];

  for (let i = 0; i < numRows; i++) {
    const uuid = dashboard.getRange(startRow + i, 1).getValue();
    if (uuid) uuids.push(uuid);
  }
  return uuids;
}

/**
 * ✅ Führt einzelnen Schritt aus (Wrapper für Menü)
 */
function claudeExecuteSingleStep(uuids, stepNumber) {
  const ui = SpreadsheetApp.getUi();
  let processed = 0;
  let errors = 0;

  for (const uuid of uuids) {
    try {
      claudeExecuteStepForPaper(uuid, stepNumber);
      processed++;
    } catch (e) {
      Logger.log(`Fehler bei Paper ${uuid}, Schritt ${stepNumber}: ${e.message}`);
      errors++;
    }
  }

  ui.alert('Schritt abgeschlossen', `${processed} Papers verarbeitet\n${errors} Fehler`);
}

/**
 * ✅ Führt einen Schritt für ein Paper aus via Claude API
 */
function claudeExecuteStepForPaper(uuid, stepNumber) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('CLAUDE_API_KEY');
  if (!apiKey) throw new Error('CLAUDE_API_KEY nicht gesetzt');

  const data = getDashboardDataByUUID(uuid);
  if (!data) throw new Error(`Paper ${uuid} nicht gefunden`);

  // Hier Prompt aufbauen und Claude API aufrufen
  // Nutzt dieselbe Prompt-Infrastruktur wie geminiSheets.gs
  const phaseMap = ['RESEARCHER', 'TRIAGE', 'METADATEN', 'MASTER_ANALYSE', 'REDAKTION', 'FAKTENCHECK', 'REVIEW'];
  const phase = phaseMap[stepNumber];

  if (!phase) throw new Error(`Unbekannter Schritt: ${stepNumber}`);

  // TODO: Claude API Call implementieren wenn API Key vorhanden
  // Vorlage: callGeminiAPI() in gemini_api.gs
  Logger.log(`[Claude] Schritt ${stepNumber} (${phase}) für ${uuid}`);
}

/**
 * ✅ Parst Claude-Ergebnis und speichert
 * Nutzt dieselbe Logik wie Gemini (applyJSONToDashboardPhaseAware)
 */
function parseAndSaveClaudeResult(uuid, stepNumber, result) {
  const phaseMap = ['RESEARCHER', 'TRIAGE', 'METADATEN', 'MASTER_ANALYSE', 'REDAKTION', 'FAKTENCHECK', 'REVIEW'];
  const phase = phaseMap[stepNumber];

  try {
    const jsonData = parseJsonFromLlm(result); // aus geminiSheets.gs
    applyJSONToDashboardPhaseAware(uuid, phase, jsonData); // aus geminiSheets.gs
  } catch (e) {
    Logger.log(`parseAndSaveClaudeResult Error: ${e.message}`);
    throw e;
  }
}

// ==========================================
// KOSTEN-KALKULATION
// ==========================================

/**
 * ✅ Schätzt Claude-Kosten
 */
function estimateClaudeCost(numPapers, steps) {
  const avgTokensPerStep = { 0: 3000, 1: 4000, 2: 5000, 3: 8000, 4: 6000, 5: 5000, 6: 4000 };
  const costPerMillion = { haiku: 0.40, sonnet: 4.00, opus: 20.00 };

  let totalCost = 0;
  let totalTokens = 0;

  for (const step of steps) {
    const tokens = avgTokensPerStep[step] || 5000;
    totalTokens += tokens * numPapers;

    let costPerToken;
    if (step <= 2) costPerToken = costPerMillion.haiku / 1000000;
    else if (step <= 4) costPerToken = costPerMillion.sonnet / 1000000;
    else costPerToken = costPerMillion.opus / 1000000;

    totalCost += tokens * numPapers * costPerToken;
  }

  return {
    tokens: totalTokens,
    cost: totalCost,
    estimate: `$${totalCost.toFixed(2)}`,
    perPaper: `$${(totalCost / numPapers).toFixed(3)}`
  };
}

/**
 * ✅ Kosten-Rechner (Menü-Funktion)
 */
function showClaudeCostCalculator() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt('Claude Kosten-Rechner', 'Wie viele Papers?', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;

  const numPapers = parseInt(response.getResponseText()) || 10;
  const costAll = estimateClaudeCost(numPapers, [0, 1, 2, 3, 4, 5, 6]);
  const costTriage = estimateClaudeCost(numPapers, [1]);
  const costTriageAnalysis = estimateClaudeCost(numPapers, [1, 3]);

  const message =
    `=== CLAUDE KOSTEN-RECHNER ===\n\n` +
    `Für ${numPapers} Papers:\n\n` +
    `📊 Komplett (Schritt 0-6): ${costAll.estimate} (${costAll.perPaper}/Paper)\n` +
    `🎯 Nur Triage: ${costTriage.estimate} (${costTriage.perPaper}/Paper)\n` +
    `📝 Triage + Analyse: ${costTriageAnalysis.estimate} (${costTriageAnalysis.perPaper}/Paper)\n\n` +
    `💡 TIPP: Erst Triage → nur hohe Relevanz analysieren!`;

  ui.alert('Kosten-Rechner', message, ui.ButtonSet.OK);
}
