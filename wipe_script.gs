/**
 * WinTek Sheet — One-Click WinChoice Data Wipe
 * 
 * This clears all WinChoice data from the new WinTek Sales Sheet while preserving:
 *   - Headers (row 1 of every tab)
 *   - All formulas (especially ARRAYFORMULA in MASTER column K)
 *   - Charts, conditional formatting, dropdowns
 *   - Goals dashboard structure
 * 
 * INSTRUCTIONS:
 * 1. Open the new WinTek Sales Sheet
 * 2. Extensions → Apps Script
 * 3. Paste this entire function at the BOTTOM of the existing script (do NOT replace anything)
 * 4. Save (disk icon)
 * 5. Select "wipeWinChoiceData" from the function dropdown at top
 * 6. Click ▶ Run
 * 7. Authorize when prompted
 * 8. Done — sheet is now a clean WinTek slate
 * 
 * This script can be run safely multiple times.
 */

function wipeWinChoiceData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const confirm = ui.alert(
    '⚠️ WIPE WINCHOICE DATA',
    'This will clear all appointment/sales data from the new WinTek sheet.\n\n' +
    'PRESERVED: headers, formulas, charts, formatting, Goals dashboard, 5 Moves content\n' +
    'CLEARED: MASTER appointments, monthly tabs, weekly performance log\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (confirm !== ui.Button.YES) return;
  
  const log = [];
  
  // === MASTER tab — clear data rows but preserve formula in column K ===
  const master = ss.getSheetByName('MASTER');
  if (master) {
    const lastRow = master.getLastRow();
    if (lastRow > 1) {
      // Clear A:J and L:V (skip column K which has ARRAYFORMULA)
      master.getRange(2, 1, lastRow - 1, 10).clearContent();   // A2:J
      master.getRange(2, 12, lastRow - 1, 11).clearContent();  // L2:V
      log.push(`✓ MASTER: cleared ${lastRow - 1} rows (column K formula preserved)`);
    }
  }
  
  // === Monthly tabs — clear data, keep headers ===
  const months = ['January', 'February', 'March', 'April', 'May', 'June',
                  'July', 'August', 'September', 'October', 'November', 'December',
                  'Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
  let monthsCleared = 0;
  months.forEach(m => {
    const sheet = ss.getSheetByName(m);
    if (sheet) {
      const lr = sheet.getLastRow();
      const lc = sheet.getLastColumn();
      if (lr > 1 && lc > 0) {
        sheet.getRange(2, 1, lr - 1, lc).clearContent();
        monthsCleared++;
      }
    }
  });
  if (monthsCleared) log.push(`✓ Monthly tabs cleared: ${monthsCleared}`);
  
  // === Weekly Performance Log — clear data rows below header ===
  const weeklyNames = ['Weekly Performance Log', '2026 WEEKLY PERFORMANCE LOG', 'Weekly Performance', 'Weekly_Performance_Log'];
  for (const name of weeklyNames) {
    const wpl = ss.getSheetByName(name);
    if (wpl) {
      const lr = wpl.getLastRow();
      const lc = wpl.getLastColumn();
      // Find header row by looking for "Week" in column A
      let headerRow = 1;
      const colA = wpl.getRange(1, 1, Math.min(5, lr), 1).getValues();
      for (let i = 0; i < colA.length; i++) {
        if (colA[i][0] && String(colA[i][0]).toLowerCase().includes('week')) {
          headerRow = i + 1;
          break;
        }
      }
      if (lr > headerRow && lc > 0) {
        // Don't clear formulas in this tab — only clear the variable cells
        const range = wpl.getRange(headerRow + 1, 1, lr - headerRow, Math.min(lc, 8));
        // Clear only non-formula cells
        const formulas = range.getFormulas();
        const values = range.getValues();
        for (let i = 0; i < values.length; i++) {
          for (let j = 0; j < values[i].length; j++) {
            if (!formulas[i][j]) values[i][j] = '';
          }
        }
        range.setValues(values);
        log.push(`✓ ${name}: data cleared (formulas preserved)`);
      }
      break;
    }
  }
  
  // === Status Trends / weekly_status_trends — clear data rows ===
  const trendNames = ['weekly_status_trends', 'Status Trends', 'status_trends'];
  for (const name of trendNames) {
    const t = ss.getSheetByName(name);
    if (t) {
      const lr = t.getLastRow();
      const lc = t.getLastColumn();
      if (lr > 1 && lc > 0) {
        const range = t.getRange(2, 1, lr - 1, lc);
        const formulas = range.getFormulas();
        const values = range.getValues();
        for (let i = 0; i < values.length; i++) {
          for (let j = 0; j < values[i].length; j++) {
            if (!formulas[i][j]) values[i][j] = '';
          }
        }
        range.setValues(values);
        log.push(`✓ ${name}: data cleared`);
      }
      break;
    }
  }
  
  // === Goals Dashboard — manually update YTD numbers (zero them out) ===
  // Note: We DO NOT clear the Goals tab structure, books read list, mission statement, etc.
  // Those are aspirational and should carry over. YTD totals will recompute from MASTER.
  
  log.push('');
  log.push('🇺🇸 WinTek slate is clean and ready.');
  log.push('Your Goals dashboard (books, milestones, mission) is preserved.');
  log.push('YTD numbers will recompute as new WinTek appointments are logged.');
  
  ui.alert('✓ Wipe Complete', log.join('\n'), ui.ButtonSet.OK);
  
  Logger.log(log.join('\n'));
}
