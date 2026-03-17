// FILE: config.gs

/**
 * Zentrale Konfiguration
 * ✅ ENHANCED: Aktive Prompt-Versionierung
 */

const CONFIG_SHEET_NAME = "Konfiguration";
const DASHBOARD_SHEET_NAME = "Dashboard";
const KATEGORIEN_SHEET_NAME = "Kategorien";
const MANUAL_PROMPTS_SHEET_NAME = "Manuelle Prompts";
const FULLTEXT_SEARCH_SHEET_NAME = "Volltext-Suche";
const MANUAL_FULLTEXT_SHEET_NAME = "Manuelle Volltextsuche";
const MANUAL_IMPORT_SHEET_NAME = "Manueller Import";
const EXPORT_HISTORY_SHEET_NAME = "Export-Historie";
const ONEPAGER_HISTORY_SHEET_NAME = "OnePager-Historie";
const ERROR_LOG_SHEET_NAME = "Fehlerprotokoll";
const ERROR_LIST_SHEET_NAME = "Fehlerliste";
const LOGBOOK_SHEET_NAME = "Logbuch";
const IMPORT_REPORT_SHEET_NAME = "Import-Report";

const DASHBOARD_HEADERS = [
 "UUID",                    // A  1
  "PMID",                    // B  2
  "DOI",                     // C  3
  "Titel",                   // D  4
  "Autoren",                 // E  5
  "Publikationsdatum",       // F  6  ← WAR "Jahr" - GEÄNDERT
  "Journal/Quelle",          // G  7
  "Volume",                  // H  8
  "Issue",                   // I  9
  "Pages",                   // J  10
  "Artikeltyp/Studientyp",   // K  11
  "Quelle",                  // L  12
  "Link",                    // M  13
  "Link zum Volltext",       // N  14
  "Volltext-Status",         // O  15
  "Volltext-Datei-Link",     // P  16
  "Inhalt/Abstract",         // Q  17
  "Volltext/Extrakt",        // R  18
  "Volltext_Teil2",          // S  19
  "Volltext_Teil3",          // T  20
  "Volltext Original Länge", // U  21
  "Volltext Bereinigt Länge",// V  22
  "Hauptkategorie",          // W  23
  "Unterkategorien",         // X  24
  "Schlagwörter",            // Y  25
  "Relevanz",                // Z  26
  "Relevanz-Begründung",     // AA 27
  "Produkt-Fokus",           // AB 28
  "Haupterkenntnis",         // AC 29
  "Kernaussagen",            // AD 30
  "Zusammenfassung",         // AE 31
  "Praktische Implikationen",// AF 32
  "Kritische Bewertung",     // AG 33
  "Evidenzgrad",             // AH 34
  "PICO Population",         // AI 35
  "PICO Intervention",       // AJ 36
  "PICO Comparator",         // AK 37
  "PICO Outcomes",           // AL 38
  "Review-Status",           // AM 39
  "Status",                  // AN 40
  "Batch-Phase",             // AO 41
  "Für Citavi-Export",       // AP 42
  "Export-Status",           // AQ 43
  "Exportiert am",           // AR 44
  "OnePager Link",           // AS 45
  "OnePager Status",         // AT 46
  "OnePager erstellt am",    // AU 47
  "Fehler-Details",          // AV 48
  "Import-Timestamp",        // AW 49
  "Letzte Änderung",         // AX 50
  "Fingerprint",             // AY 51
  "Prompt-Version",          // AZ 52
  "Mail-Betreff",            // BA 53
  "Mail-Absender",           // BB 54
  "Mail-Datum",              // BC 55
  "Gemini Researcher",       // BD 56  ← WAR AY(51)
  "Gemini Triage",           // BE 57  ← WAR AZ(52)
  "Gemini Metadaten",        // BF 58  ← WAR BA(53)
  "Gemini Master Analyse",   // BG 59  ← WAR BB(54)
  "Gemini Redaktion",        // BH 60  ← WAR BC(55)
  "Gemini Faktencheck",      // BI 61  ← WAR BD(56)
  "Gemini Review",           // BJ 62  ← WAR BE(57)
  "Flow_Trigger_Researcher", // BK 63  ← NEU
  "Flow_Trigger_Triage",     // BL 64  ← NEU
  "Flow_Trigger_Metadaten",  // BM 65  ← NEU
  "Flow_Trigger_Analyse",    // BN 66  ← NEU
  "Flow_Trigger_Redaktion",  // BO 67  ← NEU
  "Flow_Trigger_Faktencheck",// BP 68  ← NEU
  "Flow_Trigger_Review"      // BQ 69  ← NEU

];

const CONFIG_DEFAULTS = {
  "Gmail Label Name": "citavi-literatur-updates",
  "Gmail Search Query": "",
  "Onepager Hauptordner ID": "1IQ1sM8F4SMFlhR_mIg27ACshAA6YZg4z",
  "PDF Hauptordner ID": "1ix2rvtnYCOK5ebmnmzGI2BkJLYBHo-dG",
  "Abstracts Hauptordner ID": "1Z6qC0eU1BGZHgMfAIs3eFqj6A3v5X_7j",
  "Onepager Vorlagen ID": "1Mi40EcHZn6xxL8FuGNjkCfp75A812xMjxvL3lXKGcoM",
  "Archiv Spreadsheet ID": "1r8iNHJjVo6GIf_jMSBx3vLb6XGm1e2bvcqbmnFu0W0E",
  "Schnell-Generator Sheet ID": "1YStV542IiiNvlQ0RLFsctRPd4tBQVvVVYYIEMQimXRI",
  "Citavi Export Ordner ID": "1ND6JbIaDK8znUo5wXZPLqomGsyCb5qd4",
  "ONEPAGER_TEMPLATE_FILE_ID": "",
  "ONEPAGER_OUTPUT_FOLDER_ID": "",
  "Use API": "FALSE",
  "OpenAI API Key": "",
  "OpenAI Model": "gpt-4",
  "PROMPT_TEMPLATE_TRIAGE": "",
  "PROMPT_TEMPLATE_ANALYSIS": "",
  "PROMPT_TEMPLATE_METADATA_EXTRACT": "",
  "PROMPT_TEMPLATE_FACTCHECK_BATCH": "",
  "PROMPT_TEMPLATE_REVIEW": ""
};

function setConfig(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!sheet) return;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  sheet.appendRow([key, value]);
}

/**
 * ✅ ENHANCED: Aktive Versionierung basierend auf ALLEN 7 PROMPTS (B15-B21)
 * Nutzt MD5-Hash der ersten 1000 Zeichen kombiniert für Performance
 */
function getCurrentPromptVersion() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  
  if (!sheet) return "N/A";
  
  try {
    // Hole ALLE 7 Prompts und kombiniere sie
    const allPrompts = [
      sheet.getRange("B15").getValue() || "", // RESEARCHER
      sheet.getRange("B16").getValue() || "", // MASTER_ANALYSE
      sheet.getRange("B17").getValue() || "", // METADATEN
      sheet.getRange("B18").getValue() || "", // REDAKTION
      sheet.getRange("B19").getValue() || "", // TRIAGE
      sheet.getRange("B20").getValue() || "", // FAKTENCHECK
      sheet.getRange("B21").getValue() || ""  // REVIEW
    ].join("|||"); // Trennzeichen zwischen Prompts
    
    if (allPrompts === "||||||") {
      // Alle Prompts leer
      return "N/A";
    }
    
    // Verwende erste 1000 Zeichen für Hash (Performance + Aussagekraft)
    const textForHash = allPrompts.substring(0, 1000);
    return computeHash(textForHash);
    
  } catch (e) {
    Logger.log("getCurrentPromptVersion Error: " + e.message);
    return "ERROR";
  }
}

function computeHash(text) {
  if (!text) return "N/A";
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, text);
  return digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('').substring(0, 12);
}

/**
 * ✅ NEU: Prüft ob Prompt-Versionierung aktiv ist
 * Wird beim Import/Analyse-Start aufgerufen
 */
function ensurePromptVersionTracking() {
  const currentVersion = getCurrentPromptVersion();
  
  if (currentVersion === "N/A" || currentVersion === "ERROR") {
    Logger.log("⚠️ Warnung: Prompt-Versionierung nicht aktiv (Kein Prompt in B16)");
    return false;
  }
  
  Logger.log("✅ Prompt-Versionierung aktiv: " + currentVersion);
  return true;
}
