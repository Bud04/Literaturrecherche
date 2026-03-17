VitalAire Literaturrecherche – Projekt-Kontext für Claude
Kurzanleitung für neue Chats
Schick diese URL: `https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/CONTEXT.md`
Schick die Raw-URLs der relevanten Dateien (siehe "Welche Dateien für welche Aufgabe")
Beschreibe das Problem / Feature – kein weiterer Kontext nötig
---
Was ist das System?
Google Apps Script Projekt für VitalAire GmbH. Automatisiert Import, KI-Analyse und Verwaltung wissenschaftlicher Publikationen aus PubMed. Läuft in Google Sheets + Google Workspace Studio Flows (Gemini AI).
Technischer Stack: Google Apps Script (JavaScript), Google Sheets, Google Workspace Studio Flows, PubMed API, Gemini API, Unpaywall API, OCR.space API
Arbeitsumgebung: Apps Script läuft auf einem Arbeits-PC (kein Claude-Zugriff dort). GitHub dient als Brücke – Isabells Privat-PC nutzt GitHub Raw-URLs um Claude den Code zu zeigen. Änderungen werden manuell zwischen Apps Script und GitHub synchronisiert.
---
Architektur
7-Phasen KI-Workflow (Herzstück des Systems)
Papers durchlaufen automatisch 7 Phasen via Google Workspace Studio Flows:
Phase	Name	Prompt-Zelle	Stoppt bei
0	Researcher	B15	–
1	Triage	B19	Irrelevant / Niedrig / Mittel
2	Metadaten	B17	–
3	Master Analyse	B16	–
4	Redaktion	B18	–
5	Faktencheck	B20	–
6	Review	B21	–
`pollWorkflowProgress()` in `workspace_flow_core.gs` läuft jede Minute per Zeitbasiertem Trigger und orchestriert die gesamte Pipeline.
Trigger-System (KRITISCH – nicht ändern ohne guten Grund)
Trigger-Spalten: `Flow_Trigger_Researcher` bis `Flow_Trigger_Review` (BK–BQ)
Output-Spalten: `Gemini Researcher` bis `Gemini Review` (BD–BJ)
Werte-Zyklus: leer → `PENDING` → `START` → leer (nach Fertigstellung)
Nur ein Flow pro Typ darf gleichzeitig `START` haben
`pollWorkflowProgress()` befördert `PENDING` → `START`, erkennt Stuck-Flows und retriggt sie
Dashboard-Struktur
69 Spalten (A–BQ), autoritativ definiert in `DASHBOARD_HEADERS` in `config.gs`.
Wichtigste Spalten:
A = UUID (Primärschlüssel – IMMER, nie PMID/DOI)
D = Titel
F = Publikationsdatum (NICHT "Jahr" – das ist veraltet)
Q = Inhalt/Abstract
R = Volltext/Extrakt (Teil 1, max 45.000 Zeichen)
S = Volltext_Teil2
T = Volltext_Teil3
AN = Status
BD–BJ = Gemini-Outputs (Researcher bis Review)
BK–BQ = Flow-Trigger
Prompts
Im Konfiguration-Sheet, Zellen B15–B21 (siehe Tabelle oben).
Variablen im Prompt: `[Variable: Titel]`, `[Variable: Inhalt/Abstract]`, `[Variable: Volltext/Extrakt]` etc.
---
Dateien – Übersicht und Abhängigkeiten
Kern-Dateien (am häufigsten bearbeitet)
Datei	Zweck	Ruft auf
`workspace_flow_core.gs`	Pipeline-Motor, pollWorkflowProgress(), applyResultsToDashboard()	utils.gs, helpers_shared.gs
`smart_pubmed_import.gs`	PubMed Import Tag-für-Tag, smartSavePaper(), smartFetchFulltext()	utils.gs, DuplicateCheck.gs
`pubmed_import_wrapper.gs`	importPapersFromPubMed(), importFromPubMedWithDate()	smart_pubmed_import.gs, import_status.gs
`helpers_shared.gs`	Geteilte Hilfsfunktionen – zentrale Sammlung	utils.gs, geminiSheets.gs
`ManualFlow.gs`	Manueller Flow-Dialog, buildPrompt(), submitFlowResponse(), getNextFlowItem()	workspace_flow_core.gs, utils.gs
`DuplicateCheck.gs`	Duplikat-Erkennung nach Import (DOI, PMID, Titel-Ähnlichkeit)	utils.gs
`import_status.gs`	Import-Historie pro Kategorie, Lücken-Erkennung, saveImportRange()	–
`ui.gs`	onOpen(), Menü, alle Menü-Callbacks	alle
`ui_pubmed_menu_extended.gs`	Vordefinierte PubMed-Suchen (Diabetes T1/T2, Immunologie, Neurologie)	pubmed_import_wrapper.gs, import_status.gs
Weitere aktive Dateien
Datei	Zweck
`config.gs`	DASHBOARD_HEADERS, Sheet-Namen, CONFIG_DEFAULTS, getCurrentPromptVersion()
`utils.gs`	updateDashboardField(), getDashboardDataByUUID(), getDashboardColumnIndex(), loadCategoriesAndKeywords()
`geminiSheets.gs`	Gemini Sheets Batch-Workflow, applyJSONToDashboardPhaseAware(), parseJsonFromLlm()
`gemini_api.gs`	Direkte Gemini API Calls (GEMINI_API_KEY)
`sidebar_backend.gs`	Sidebar-Logik, getSidebarData(), saveSidebarData()
`manualimport.gs`	Manueller Import aus "Manueller Import" Sheet, importPublicationToDashboard()
`log.gs`	logAction(), logError(), logToImportReport()
`abstract_fixer.gs`	Abstracts vervollständigen (12 Publisher-Parser)
`fulltext.gs`	Volltext-Suche (Unpaywall, PMC, Kandidaten-Links)
`onepager.gs`	OnePager-Erstellung aus Google Doc Vorlage
`exportCitavi.gs`	RIS-Export für Citavi
`review.gs`	JSON-Validierung, applyJSONToDashboard(), validatePhaseData()
`quality_checks.gs`	Readiness-Checks für OnePager + Citavi
`pdf_processing.gs`	PDF-Upload, Drive-Konvertierung, OCR
`selective Ocr.gs`	Selective OCR via OCR.space API für Tabellen/Figures
`batchPipeline.gs`	Älterer manueller Batch-Workflow (selten genutzt)
`BatchHelpers.gs`	Batch-Phase Farben, Filter, Progress
`prompt_automation.gs`	Auto-Regenerierung bei Volltext-Update
`triggerManager.gs`	Trigger installieren/entfernen
`update_dashboard_volltext.gs`	Volltext-Spalten verwalten, Statistik
`init.gs`	initialSetup(), ensureSheetsAndHeaders()
`events.g`	onEdit() – setzt "Letzte Änderung" Timestamp
HTML-Dateien
`Sidebar.html`, `Dialog.html`, `BatchPromptDialog.html`, `manualFlowDialog.html`, `flowStatusDialog.html`, `duplicateCheckDialog.html`
---
Welche Dateien für welche Aufgabe laden?
Flow/Pipeline-Probleme
`workspace_flow_core.gs` + `helpers_shared.gs` + `utils.gs`
Import-Probleme (PubMed)
`smart_pubmed_import.gs` + `pubmed_import_wrapper.gs` + `import_status.gs`
Manueller Flow / Dialog
`ManualFlow.gs` + `workspace_flow_core.gs` + `helpers_shared.gs`
Duplikat-Probleme
`DuplicateCheck.gs` + `smart_pubmed_import.gs`
Menü / UI-Probleme
`ui.gs` + `ui_pubmed_menu_extended.gs`
JSON-Parsing / KI-Output Probleme
`geminiSheets.gs` + `workspace_flow_core.gs`
Dashboard-Spalten / Konfiguration
`config.gs` + `utils.gs`
Volltext / PDF
`fulltext.gs` + `abstract_fixer.gs` + `pdf_processing.gs` + `selective Ocr.gs`
OnePager / Export
`onepager.gs` + `exportCitavi.gs` + `review.gs`
---
Konventionen – was Claude beachten muss
Immer einhalten
Spaltenname `Publikationsdatum` verwenden, NICHT `Jahr` (veraltet)
UUID als Primärschlüssel – alle Schreiboperationen über UUID
Dashboard-Schreiboperationen NUR über `updateDashboardField(uuid, spaltenname, wert)` aus `utils.gs`
`getConfig(key)` aus `helpers_shared.gs` ist die autoritative Version
`getDashboardColumnIndex()` aus `utils.gs` ist die autoritative Version
Einziges `onOpen()` in `ui.gs` – in keiner anderen Datei!
Alle neuen Menüpunkte nur in `ui.gs` ergänzen
Nie ändern ohne explizite Aufforderung
Die Reihenfolge von `DASHBOARD_HEADERS` in `config.gs`
Die Prompt-Zellen B15–B21 (nur Inhalt, nicht Zell-Referenz)
Die Trigger-Logik PENDING→START in `pollWorkflowProgress()`
Die Spalten-Nummern in `GEMINI_COLUMNS` in `geminiSheets.gs`
Stil
`workspace_flow_core.gs` und `smart_pubmed_import.gs` verwenden `var` (altes JS) – so beibehalten
Neuere Dateien verwenden `const`/`let` – so beibehalten
Fehler immer über `logError()` aus `log.gs` loggen
Benutzer-Feedback immer über `SpreadsheetApp.getUi().alert()`
---
Bekannte Altlasten (noch vorhanden, nicht primär genutzt)
Datei	Problem
`workflow.gs`	Alte OpenAI/GPT-4o Pipeline. Enthält Duplikate für `setCell`, `getRowData`, `getCategoriesAndTags`, `writeResultsToSheetExtended`. Nicht aktiv genutzt.
`Claud.gs`	Claude API nie fertiggestellt (TODO-Stubs). Funktionen nicht aufgerufen.
`geminiWorkspace.gs`	Nur eine Testfunktion, Syntaxfehler am Dateiende (`}geminiWorkspace.gs`).
`batchPipeline.gs`	Referenziert `rowData.Jahr` statt `rowData["Publikationsdatum"]`. Selten genutzt.
Hinweis für Claude: Falls diese Dateien geladen werden, ihre Duplikate nicht in andere Dateien übertragen. Der aktive Code in `helpers_shared.gs` und `utils.gs` hat Vorrang.
---
Raw-URL Liste (für README)
```
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/workspace_flow_core.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/smart_pubmed_import.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/pubmed_import_wrapper.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/helpers_shared.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/ManualFlow.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/DuplicateCheck.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/import_status.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/ui.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/ui_pubmed_menu_extended.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/config.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/utils.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/geminiSheets.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/gemini_api.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/sidebar_backend.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/manualimport.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/log.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/abstract_fixer.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/fulltext.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/onepager.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/exportCitavi.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/review.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/quality_checks.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/pdf_processing.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/prompt_automation.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/triggerManager.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/update_dashboard_volltext.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/init.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/events.g
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/batchPipeline.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/BatchHelpers.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/selective%20Ocr.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/workflow.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/Claud.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/geminiWorkspace.gs
https://raw.githubusercontent.com/Bud04/Literaturrecherche/main/CONTEXT.md
```
