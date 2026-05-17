/**
 * УПРАВЛЕНИЕ ГОДОВЫМИ ЛИСТАМИ (2026, 2025, ...)
 * - Формулы для строк 33–42
 * - Расчёт компенсаций
 * - Перенос остатков при создании нового года
 */

function fixYearSheetFormulas(yearSheetName) {
  var sheet = getSheetSafe(yearSheetName);
  if (!sheet) return;
  var lastCol = sheet.getLastColumn();
  for (var col = 2; col <= lastCol; col += 2) {
    var colLetter = String.fromCharCode(64 + col);
    var nextColLetter = String.fromCharCode(64 + col + 1);
    sheet.getRange(ROWS.TOTAL, col).setFormula(`=SUM(${colLetter}${ROWS.DATA_START}:${colLetter}${ROWS.DATA_END})`);
    sheet.getRange(ROWS.TOTAL, col + 1).setFormula(`=SUM(${nextColLetter}${ROWS.DATA_START}:${nextColLetter}${ROWS.DATA_END})`);
    [ROWS.GIFTS, ROWS.PROJECTS, ROWS.INSURANCE, ROWS.PENSION, ROWS.COMPENSATING, ROWS.COMPENSATING2].forEach(function(row) {
      sheet.getRange(row, col).setFormula(`=SUM(${colLetter}${ROWS.DATA_START}:${colLetter}${ROWS.DATA_END})`);
    });
    var balanceFormula = `${colLetter}${ROWS.TOTAL}-${nextColLetter}${ROWS.TOTAL}`;
    balanceFormula += `-(${colLetter}${ROWS.GIFTS}+${colLetter}${ROWS.COMPENSATING})`;
    balanceFormula += `-(${colLetter}${ROWS.PROJECTS}+${colLetter}${ROWS.COMPENSATING2})`;
    balanceFormula += `-${colLetter}${ROWS.INSURANCE}-${colLetter}${ROWS.PENSION}`;
    sheet.getRange(ROWS.BALANCE, col + 1).setFormula(balanceFormula);
  }
  showMessage('Готово', `Формулы для листа "${yearSheetName}" обновлены`, 'info');
  addLogEntry(`Обновлены формулы для листа ${yearSheetName}`, 'INFO', 'YearSheetManager');
}

function recalcYearSheetCompensations(yearSheetName) { updateAllCompensations(yearSheetName); }
function fixCurrentYearCompensations() { updateAllCompensations(SHEETS.CURRENT_YEAR); }
function fixAllYearsCompensations() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    if (/^\d{4}$/.test(name)) updateAllCompensations(name);
  }
}
// =============================================
// НОВЫЙ ГОД – создание листа на следующий год
// =============================================

function createNewYearSheet() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var currentYearSheet = ss.getSheetByName(SHEETS.CURRENT_YEAR);
    if (!currentYearSheet) {
      showMessage('Ошибка', 'Лист текущего года не найден', 'error');
      return;
    }
    var nextYear = CONFIG.CURRENT_YEAR + 1;
    var nextYearSheetName = nextYear.toString();
    var nextYearSheet = ss.getSheetByName(nextYearSheetName);
    
    if (nextYearSheet) {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert(
        'Создание листа на следующий год',
        'Лист "' + nextYearSheetName + '" уже существует. Пересоздать? Данные будут потеряны.',
        ui.ButtonSet.YES_NO
      );
      if (response !== ui.Button.YES) return;
      ss.deleteSheet(nextYearSheet);
    }
    
    // Копируем структуру текущего года
    nextYearSheet = currentYearSheet.copyTo(ss);
    nextYearSheet.setName(nextYearSheetName);
    nextYearSheet.setTabColor('#f9cb9c'); // меняем цвет, чтобы отличать
    
    // Очищаем данные за предыдущий год, но оставляем формулы
    var lastCol = nextYearSheet.getLastColumn();
    for (var col = 2; col <= lastCol; col++) {
      for (var row = ROWS.DATA_START; row <= ROWS.DATA_END; row++) {
        var cell = nextYearSheet.getRange(row, col);
        if (!cell.getFormula()) cell.clearContent();
      }
      // Очищаем строки компенсаций (кроме формул)
      var specialRows = [ROWS.GIFTS, ROWS.PROJECTS, ROWS.INSURANCE, ROWS.PENSION, ROWS.COMPENSATING, ROWS.COMPENSATING2];
      for (var r = 0; r < specialRows.length; r++) {
        var specialCell = nextYearSheet.getRange(specialRows[r], col);
        if (!specialCell.getFormula()) specialCell.clearContent();
      }
      // Сбрасываем решения (строка PAYMENT)
      nextYearSheet.getRange(ROWS.PAYMENT, col).clearContent();
    }
    
    // Переносим остаток с декабря на январь нового года
    var decemberBalance = currentYearSheet.getRange(ROWS.BALANCE, 25).getValue(); // 25 - колонка для декабря (2 + 23)
    if (decemberBalance && decemberBalance !== 0) {
      // В новом листе январь — это колонки 2 (доходы) и 3 (расходы). Остаток января пишется в колонку 3? Логика: строка BALANCE, нечётная колонка
      var januaryBalanceCell = nextYearSheet.getRange(ROWS.BALANCE, 3);
      var currentBalance = safeParseNumber(januaryBalanceCell.getValue());
      januaryBalanceCell.setValue(currentBalance + decemberBalance);
    }
    
    // Обновляем глобальную переменную CONFIG.CURRENT_YEAR (только в скрипте, не в Properties)
    CONFIG.CURRENT_YEAR = nextYear;
    SHEETS.CURRENT_YEAR = nextYearSheetName;
    
    showMessage('Готово', '✅ Создан лист "' + nextYearSheetName + '"', 'info');
    addLogEntry('Создан лист на следующий год: ' + nextYearSheetName, 'INFO', 'YearSheetManager');
    
  } catch(e) {
    logError(e, 'createNewYearSheet');
    showMessage('Ошибка', e.toString(), 'error');
  }
}

function showBalanceToTransfer() {
  try {
    var sheet = getSheetSafe(SHEETS.CURRENT_YEAR);
    if (!sheet) return;
    var decemberBalance = sheet.getRange(ROWS.BALANCE, 25).getValue(); // декабрь — колонка 25 (2 + 23)
    var message = '💰 Остаток на конец года: ' + formatMoney(decemberBalance) + '\n\n';
    message += 'При создании листа на следующий год эта сумма будет перенесена как начальный остаток января.';
    showMessage('Перенос остатка', message, 'info');
  } catch(e) {
    logError(e, 'showBalanceToTransfer');
    showMessage('Ошибка', e.toString(), 'error');
  }
}