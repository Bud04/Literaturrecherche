// FILE: geminiWorkspace.gs
function testGeminiViaFormulaSimple() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Dashboard");
  const row = 2;
  
  // 1. Abstract holen
  const abstract = sheet.getRange("Q" + row).getValue();
  
  if (!abstract || String(abstract).length < 50) {
    SpreadsheetApp.getUi().alert("Kein Abstract");
    return;
  }
  
  // 2. Stark kürzen und alle Sonderzeichen entfernen
  const cleanAbstract = String(abstract)
    .substring(0, 300)
    .replace(/"/g, '')
    .replace(/'/g, '')
    .replace(/\n/g, ' ')
    .replace(/\r/g, ' ')
    .replace(/[^\w\s.,;:\-()]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  
  Logger.log("Clean Abstract: " + cleanAbstract);
  
  // 3. Einfache Formel
  const tempCell = sheet.getRange("BA" + row);
  const formula = `=GEMINI("Ist dieser Text relevant für Diabetes oder Herz? Antworte nur Ja oder Nein. Text: ${cleanAbstract}")`;
  
  Logger.log("Formel: " + formula);
  
  tempCell.setFormula(formula);
  SpreadsheetApp.flush();
  Utilities.sleep(5000);
  
  const result = tempCell.getValue();
  tempCell.clear();
  
  SpreadsheetApp.getUi().alert("Ergebnis: " + result);
  return result;
}geminiWorkspace.gs
