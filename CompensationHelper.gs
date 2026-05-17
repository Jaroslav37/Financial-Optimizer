/**
 * КОМПЕНСАЦИИ - автоматический расчёт строк 36 и 38
 */

function updateAllCompensations(yearSheetName) {
  var sheet = getSheetSafe(yearSheetName);
  if (!sheet) { showMessage('Ошибка', 'Лист "'+yearSheetName+'" не найден', 'error'); return; }
  var lastCol = sheet.getLastColumn();
  var liquidated = false;
  for (var col = 2; col <= lastCol; col += 2) {
    var decisionCell = sheet.getRange(ROWS.PAYMENT, col);
    var decision = decisionCell.getValue();
    var decisionStr = (decision || '').toString().trim();
    var monthlyIncome = sheet.getRange(ROWS.TOTAL, col).getValue();
    if (isNaN(monthlyIncome)) monthlyIncome = 0;
    var plannedGifts = liquidated ? 0 : monthlyIncome * 0.1;
    var plannedInvest = liquidated ? 0 : monthlyIncome * 0.2;
    var actualGifts = sheet.getRange(ROWS.GIFTS, col).getValue() || 0;
    var actualInvest = sheet.getRange(ROWS.PROJECTS, col).getValue() || 0;
    var prevCompGifts = 0, prevCompInvest = 0;
    if (col > 2 && !liquidated) {
      prevCompGifts = sheet.getRange(ROWS.COMPENSATING, col - 2).getValue() || 0;
      prevCompInvest = sheet.getRange(ROWS.COMPENSATING2, col - 2).getValue() || 0;
    }
    var compGifts, compInvest;
    if (decisionStr === 'Скорректировано') {
      compGifts = sheet.getRange(ROWS.COMPENSATING, col).getValue() || 0;
      compInvest = sheet.getRange(ROWS.COMPENSATING2, col).getValue() || 0;
    } else if (decisionStr === 'Ликвидировано') {
      compGifts = 0; compInvest = 0; liquidated = true;
      decisionCell.setValue('Ликвидировано');
    } else {
      var rawGifts = prevCompGifts + plannedGifts - actualGifts;
      var rawInvest = prevCompInvest + plannedInvest - actualInvest;
      if (decisionStr === 'Погашено') { compGifts = plannedGifts - actualGifts; compInvest = plannedInvest - actualInvest; }
      else { compGifts = rawGifts; compInvest = rawInvest; }
    }
    sheet.getRange(ROWS.COMPENSATING, col).setValue(compGifts);
    sheet.getRange(ROWS.COMPENSATING2, col).setValue(compInvest);
    if (!decisionStr && !liquidated) decisionCell.setValue('Перенесено');
  }
  showMessage('Готово', 'Компенсации для листа "'+yearSheetName+'" пересчитаны', 'info');
  addLogEntry('Пересчитаны компенсации для '+yearSheetName, 'INFO', 'CompensationHelper');
}
function fixCurrentYearCompensations() { updateAllCompensations(SHEETS.CURRENT_YEAR); }
function fixAllYearsCompensations() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) { var name = sheets[i].getName(); if (/^\d{4}$/.test(name)) updateAllCompensations(name); }
}