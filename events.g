// FILE: events.gs

/**
 * Event Handler (nur onEdit)
 */

function onEdit(e) {
  if (!e) return;
  
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  if (sheet.getName() === DASHBOARD_SHEET_NAME) {
    const row = range.getRow();
    if (row > 1) {
      const lastChangeCol = getDashboardColumnIndex("Letzte Änderung");
      if (lastChangeCol > 0) {
        sheet.getRange(row, lastChangeCol).setValue(new Date());
      }
    }
  }
}
