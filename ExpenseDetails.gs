/**
 * ДЕТАЛИЗАЦИЯ РАСХОДОВ - лист "Детали расходов"
 */

class ExpenseDetailsModule {
  constructor() {
    this.detailsSheet = getSheetSafe(SHEETS.EXPENSE_DETAILS);
    this.yearSheet = getSheetSafe(SHEETS.CURRENT_YEAR);
    this.currentYear = CONFIG.CURRENT_YEAR;
  }

  initDetailsSheet() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEETS.EXPENSE_DETAILS);
      if (!sheet) sheet = ss.insertSheet(SHEETS.EXPENSE_DETAILS);
      else {
        var ui = SpreadsheetApp.getUi();
        if (ui.alert('Пересоздать лист?', 'Лист детализации уже существует. Пересоздать? Все данные будут потеряны.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
        sheet.clear();
      }
      var headers = [['Дата', 'День', 'Месяц', 'Год', 'Категория', 'Сумма', 'Описание', 'ID записи']];
      sheet.getRange(1, 1, 1, 8).setValues(headers);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f0f0f0');
      sheet.setColumnWidth(1, 100); sheet.setColumnWidth(2, 60); sheet.setColumnWidth(3, 80); sheet.setColumnWidth(4, 60);
      sheet.setColumnWidth(5, 120); sheet.setColumnWidth(6, 100); sheet.setColumnWidth(7, 300); sheet.setColumnWidth(8, 150);
      sheet.getRange(2, 1, 1000, 1).setNumberFormat('dd.mm.yyyy');
      sheet.getRange(2, 6, 1000, 1).setNumberFormat('#,##0.00');
      var examples = [
        [new Date(2026, 0, 15), 15, 'Январь', 2026, 'Продукты', 3500, 'Продукты в магазине', 'EXP-20260115-001'],
        [new Date(2026, 0, 15), 15, 'Январь', 2026, 'Транспорт', 500, 'Такси', 'EXP-20260115-002']
      ];
      sheet.getRange(2, 1, examples.length, 8).setValues(examples);
      sheet.getRange(1, 1).setNote('Детализация расходов:\nA - Дата\nB - День\nC - Месяц\nD - Год\nE - Категория\nF - Сумма\nG - Описание\nH - ID записи');
      showMessage('Готово', 'Лист детализации расходов создан', 'info');
      addLogEntry('Лист детализации расходов создан', 'INFO', 'ExpenseDetailsModule');
    } catch(e) { logError(e, 'ExpenseDetailsModule.initDetailsSheet'); showMessage('Ошибка', e.toString(), 'error'); }
  }

  addExpenseRecord(date, category, amount, description, source, subcategory) {
    try {
      if (!this.detailsSheet) { this.detailsSheet = getSheetSafe(SHEETS.EXPENSE_DETAILS); if (!this.detailsSheet) { this.initDetailsSheet(); this.detailsSheet = getSheetSafe(SHEETS.EXPENSE_DETAILS); } }
      var day = date.getDate(), month = date.getMonth(), year = date.getFullYear(), monthName = MONTHS[month];
      var dateStr = formatDateForSearch(date).replace(/\./g, '');
      var lastRow = this.detailsSheet.getLastRow();
      var seq = (lastRow + 1).toString().padStart(3, '0');
      var recordId = 'EXP-' + dateStr + '-' + seq;
      var newRow = lastRow + 1;
      this.detailsSheet.getRange(newRow, 1).setValue(date);
      this.detailsSheet.getRange(newRow, 2).setValue(day);
      this.detailsSheet.getRange(newRow, 3).setValue(monthName);
      this.detailsSheet.getRange(newRow, 4).setValue(year);
      this.detailsSheet.getRange(newRow, 5).setValue(category);
      this.detailsSheet.getRange(newRow, 6).setValue(amount).setNumberFormat('#,##0.00');
      this.detailsSheet.getRange(newRow, 7).setValue(description);
      this.detailsSheet.getRange(newRow, 8).setValue(recordId);
      this.updateYearSheetTotal(date, amount);
      return { success: true, recordId: recordId };
    } catch(e) { logError(e, 'ExpenseDetailsModule.addExpenseRecord'); return { success: false, error: e.toString() }; }
  }

  updateYearSheetTotal(date, amount) {
    if (!this.yearSheet) return;
    var day = date.getDate(), month = date.getMonth(), row = day + 1;
    var expenseCol = -1;
    var headers = this.yearSheet.getRange(1, 1, 1, this.yearSheet.getLastColumn()).getValues()[0];
    for (var col = 1; col < headers.length; col++) {
      var header = headers[col];
      if (header && header.indexOf(MONTHS[month]) !== -1 && header.indexOf('-') !== -1) { expenseCol = col + 1; break; }
    }
    if (expenseCol === -1) return;
    var current = safeParseNumber(this.yearSheet.getRange(row, expenseCol).getValue());
    this.yearSheet.getRange(row, expenseCol).setValue(current + amount);
    this.updateYearTotals();
  }

  getExpensesByPeriod(startDate, endDate) {
    if (!this.detailsSheet) return [];
    var data = this.detailsSheet.getRange(2, 1, this.detailsSheet.getLastRow() - 1, 8).getValues();
    var expenses = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var recordDate = row[0];
      if (recordDate && recordDate >= startDate && recordDate <= endDate) {
        expenses.push({ date: recordDate, day: row[1], month: row[2], year: row[3], category: row[4], amount: row[5], description: row[6], id: row[7] });
      }
    }
    return expenses;
  }

  getExpensesByCategory(category, startDate, endDate) {
    var all = this.getExpensesByPeriod(startDate, endDate);
    return all.filter(function(e) { return e.category === category; });
  }

  getCategoryStats(startDate, endDate) {
    var expenses = this.getExpensesByPeriod(startDate, endDate);
    var stats = {}, total = 0;
    for (var i = 0; i < expenses.length; i++) {
      var e = expenses[i];
      if (!stats[e.category]) stats[e.category] = { total: 0, count: 0, items: [] };
      stats[e.category].total += e.amount;
      stats[e.category].count++;
      stats[e.category].items.push(e);
      total += e.amount;
    }
    return { stats: stats, total: total };
  }

  updateYearTotals() {
    var lastCol = this.yearSheet.getLastColumn();
    for (var col = 2; col <= lastCol; col++) {
      var range = this.yearSheet.getRange(ROWS.DATA_START, col, ROWS.DATA_END - ROWS.DATA_START + 1, 1);
      var values = range.getValues();
      var total = 0;
      for (var i = 0; i < values.length; i++) total += safeParseNumber(values[i][0]);
      this.yearSheet.getRange(ROWS.TOTAL, col).setValue(total);
    }
    for (var col = 2; col <= lastCol; col += 2) {
      var income = safeParseNumber(this.yearSheet.getRange(ROWS.TOTAL, col).getValue());
      var expense = safeParseNumber(this.yearSheet.getRange(ROWS.TOTAL, col + 1).getValue());
      this.yearSheet.getRange(ROWS.BALANCE, col + 1).setValue(income - expense);
    }
  }

  deleteExpenseRecord(recordId) {
    if (!this.detailsSheet) return false;
    var data = this.detailsSheet.getRange(2, 1, this.detailsSheet.getLastRow() - 1, 8).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][7] === recordId) {
        var rowToDelete = i + 2;
        var amount = data[i][5];
        var date = data[i][0];
        this.detailsSheet.deleteRow(rowToDelete);
        this.recalculateYearSheetTotal(date);
        addLogEntry('Удалена запись о расходе: ' + formatMoney(amount) + ' от ' + formatDate(date), 'INFO', 'ExpenseDetailsModule');
        return true;
      }
    }
    return false;
  }

  recalculateYearSheetTotal(date) {
    var day = date.getDate(), month = date.getMonth(), row = day + 1;
    var expenseCol = -1;
    var headers = this.yearSheet.getRange(1, 1, 1, this.yearSheet.getLastColumn()).getValues()[0];
    for (var col = 1; col < headers.length; col++) {
      var header = headers[col];
      if (header && header.indexOf(MONTHS[month]) !== -1 && header.indexOf('-') !== -1) { expenseCol = col + 1; break; }
    }
    if (expenseCol === -1) return;
    var start = new Date(date); start.setHours(0,0,0,0);
    var end = new Date(date); end.setHours(23,59,59,999);
    var dayExpenses = this.getExpensesByPeriod(start, end);
    var total = 0;
    for (var i = 0; i < dayExpenses.length; i++) total += dayExpenses[i].amount;
    if (total > 0) this.yearSheet.getRange(row, expenseCol).setValue(total);
    else this.yearSheet.getRange(row, expenseCol).clearContent();
    this.updateYearTotals();
  }
}

// =============================================
// ФУНКЦИИ ДЛЯ ВЫЗОВА ИЗ МЕНЮ
// =============================================

function initExpenseDetails() { new ExpenseDetailsModule().initDetailsSheet(); }

function showExpenseStats() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('📊 Статистика расходов', 'Введите период в формате: ДД.ММ.ГГГГ - ДД.ММ.ГГГГ\n\nИли оставьте пустым для текущего месяца', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var input = response.getResponseText().trim();
  var startDate, endDate, today = new Date();
  if (!input) {
    startDate = new Date(today.getFullYear(), today.getMonth(), 1);
    endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  } else {
    var dates = input.split('-').map(function(s) { return s.trim(); });
    if (dates.length !== 2) { ui.alert('Ошибка', 'Неверный формат. Используйте: ДД.ММ.ГГГГ - ДД.ММ.ГГГГ', ui.ButtonSet.OK); return; }
    startDate = safeParseDate(dates[0]); endDate = safeParseDate(dates[1]);
    if (!startDate || !endDate) { ui.alert('Ошибка', 'Неверный формат даты', ui.ButtonSet.OK); return; }
  }
  var module = new ExpenseDetailsModule();
  var stats = module.getCategoryStats(startDate, endDate);
  var report = '📊 СТАТИСТИКА РАСХОДОВ\nПериод: ' + formatDate(startDate) + ' - ' + formatDate(endDate) + '\n══════════════════════\n\n💰 ВСЕГО: ' + formatMoney(stats.total) + '\n\n';
  var categories = Object.keys(stats.stats);
  for (var i = 0; i < categories.length; i++) {
    var cat = categories[i];
    var data = stats.stats[cat];
    var percent = (data.total / stats.total * 100).toFixed(1);
    report += cat + ':\n   Сумма: ' + formatMoney(data.total) + '\n   Доля: ' + percent + '%\n   Кол-во: ' + data.count + '\n\n';
  }
  showMessage('Статистика', report);
}

function showInvestmentExpensesStats() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('📊 Инвестиционные расходы', 'Введите период в формате: ДД.ММ.ГГГГ - ДД.ММ.ГГГГ\n\nИли оставьте пустым для текущего месяца', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var input = response.getResponseText().trim();
  var startDate, endDate, today = new Date();
  if (!input) {
    startDate = new Date(today.getFullYear(), today.getMonth(), 1);
    endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
  } else {
    var dates = input.split('-').map(function(s) { return s.trim(); });
    if (dates.length !== 2) { ui.alert('Ошибка', 'Неверный формат', ui.ButtonSet.OK); return; }
    startDate = safeParseDate(dates[0]); endDate = safeParseDate(dates[1]);
    if (!startDate || !endDate) { ui.alert('Ошибка', 'Неверный формат даты', ui.ButtonSet.OK); return; }
  }
  var expenseModule = new ExpenseDetailsModule();
  var expenses = expenseModule.getExpensesByCategory('Инвестиции', startDate, endDate);
  var total = expenses.reduce(function(s, e) { return s + e.amount; }, 0);
  var report = '📊 ИНВЕСТИЦИОННЫЕ РАСХОДЫ\nПериод: ' + formatDate(startDate) + ' - ' + formatDate(endDate) + '\n═══════════════════════════\n\n💰 ВСЕГО: ' + formatMoney(total) + '\n📊 Количество операций: ' + expenses.length + '\n\n';
  if (expenses.length > 0) {
    report += '📋 ДЕТАЛИЗАЦИЯ:\n';
    expenses.forEach(function(exp) {
      report += formatDate(exp.date) + ': ' + formatMoney(exp.amount) + ' - ' + exp.description + '\n';
    });
  } else report += 'Нет инвестиционных расходов за указанный период\n';
  showMessage('Инвестиционные расходы', report);
}