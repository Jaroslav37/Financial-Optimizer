/**
 * ДЕТАЛИЗАЦИЯ ДОХОДОВ - лист "Детали доходов"
 */

class IncomeDetailsModule {
  constructor() {
    this.detailsSheet = getSheetSafe(SHEETS.INCOME_DETAILS);
    this.yearSheet = getSheetSafe(SHEETS.CURRENT_YEAR);
    this.currentYear = CONFIG.CURRENT_YEAR;
  }

  initDetailsSheet() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEETS.INCOME_DETAILS);
      if (!sheet) sheet = ss.insertSheet(SHEETS.INCOME_DETAILS);
      else {
        var ui = SpreadsheetApp.getUi();
        if (ui.alert('Пересоздать лист?', 'Лист детализации доходов уже существует. Пересоздать? Все данные будут потеряны.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
        sheet.clear();
      }
      var headers = [['Дата', 'День', 'Месяц', 'Год', 'Категория', 'Сумма', 'Описание', 'ID записи', 'Источник']];
      sheet.getRange(1, 1, 1, 9).setValues(headers);
      sheet.getRange(1, 1, 1, 9).setFontWeight('bold').setBackground('#f0f0f0');
      sheet.setColumnWidth(1,100); sheet.setColumnWidth(2,60); sheet.setColumnWidth(3,80); sheet.setColumnWidth(4,60);
      sheet.setColumnWidth(5,120); sheet.setColumnWidth(6,100); sheet.setColumnWidth(7,300); sheet.setColumnWidth(8,150); sheet.setColumnWidth(9,120);
      sheet.getRange(2,1,1000,1).setNumberFormat('dd.mm.yyyy');
      sheet.getRange(2,6,1000,1).setNumberFormat('#,##0.00');
      var examples = [
        [new Date(2026,0,15),15,'Январь',2026,'Зарплата',50000,'Зарплата за январь','INC-20260115-001','Основная работа']
      ];
      sheet.getRange(2,1,examples.length,9).setValues(examples);
      sheet.getRange(1,1).setNote('Детализация доходов:\nA - Дата\nB - День\nC - Месяц\nD - Год\nE - Категория\nF - Сумма\nG - Описание\nH - ID записи\nI - Источник');
      showMessage('Готово', 'Лист детализации доходов создан', 'info');
      addLogEntry('Лист детализации доходов создан', 'INFO', 'IncomeDetailsModule');
    } catch(e) { logError(e, 'IncomeDetailsModule.initDetailsSheet'); showMessage('Ошибка', e.toString(), 'error'); }
  }

  addIncomeRecord(date, category, amount, description, source, subcategory, ticker, tax) {
    // Валидация
    if (!date || !(date instanceof Date) || isNaN(date.getTime())) {
      return { success: false, error: 'Неверная дата' };
    }
    if (!amount || amount <= 0 || !isFinite(amount)) {
      return { success: false, error: 'Сумма должна быть положительным числом' };
    }
    if (!category || typeof category !== 'string' || category.trim() === '') {
      return { success: false, error: 'Категория не указана' };
    }
    if (!description || typeof description !== 'string' || description.trim() === '') {
      return { success: false, error: 'Описание не может быть пустым' };
    }
    try {
      if (!this.detailsSheet) { this.initDetailsSheet(); this.detailsSheet = getSheetSafe(SHEETS.INCOME_DETAILS); }
      var day = date.getDate(), month = date.getMonth(), year = date.getFullYear(), monthName = MONTHS[month];
      var dateStr = formatDateForSearch(date).replace(/\./g, '');
      var lastRow = this.detailsSheet.getLastRow();
      var seq = (lastRow + 1).toString().padStart(3, '0');
      var recordId = 'INC-' + dateStr + '-' + seq;
      var newRow = lastRow + 1;
      this.detailsSheet.getRange(newRow,1).setValue(date);
      this.detailsSheet.getRange(newRow,2).setValue(day);
      this.detailsSheet.getRange(newRow,3).setValue(monthName);
      this.detailsSheet.getRange(newRow,4).setValue(year);
      this.detailsSheet.getRange(newRow,5).setValue(category);
      this.detailsSheet.getRange(newRow,6).setValue(amount).setNumberFormat('#,##0.00');
      this.detailsSheet.getRange(newRow,7).setValue(description);
      this.detailsSheet.getRange(newRow,8).setValue(recordId);
      this.detailsSheet.getRange(newRow,9).setValue(source || 'manual');
      if (subcategory || ticker || tax) {
        var note = '';
        if (subcategory) note += 'Подкатегория: ' + subcategory + '\n';
        if (ticker) note += 'Тикер: ' + ticker + '\n';
        if (tax > 0) note += 'Налог: ' + formatMoney(tax) + '\n';
        this.detailsSheet.getRange(newRow,7).setNote(note);
      }
      this.updateYearSheetTotal(date, amount);
      return { success: true, recordId: recordId };
    } catch(e) { logError(e, 'IncomeDetailsModule.addIncomeRecord'); return { success: false, error: e.toString() }; }
  }

  updateYearSheetTotal(date, amount) {
    if (!this.yearSheet) return;
    var day = date.getDate(), month = date.getMonth(), row = day + 1;
    var incomeCol = -1;
    var headers = this.yearSheet.getRange(1,1,1,this.yearSheet.getLastColumn()).getValues()[0];
    for (var col = 1; col < headers.length; col++) {
      var header = headers[col];
      if (header && header.indexOf(MONTHS[month]) !== -1 && header.indexOf('+') !== -1) { incomeCol = col + 1; break; }
    }
    if (incomeCol === -1) return;
    var current = safeParseNumber(this.yearSheet.getRange(row, incomeCol).getValue());
    this.yearSheet.getRange(row, incomeCol).setValue(current + amount);
    this.updateYearTotals();
  }

  getIncomesByPeriod(startDate, endDate) {
    if (!this.detailsSheet) return [];
    var data = this.detailsSheet.getRange(2,1,this.detailsSheet.getLastRow()-1,9).getValues();
    var incomes = [];
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var recordDate = row[0];
      if (recordDate && recordDate >= startDate && recordDate <= endDate) {
        incomes.push({ date: recordDate, day: row[1], month: row[2], year: row[3], category: row[4], amount: row[5], description: row[6], id: row[7], source: row[8] });
      }
    }
    return incomes;
  }

  getIncomesByCategory(category, startDate, endDate) {
    var all = this.getIncomesByPeriod(startDate, endDate);
    return all.filter(function(i) { return i.category === category; });
  }

  getCategoryStats(startDate, endDate) {
    var incomes = this.getIncomesByPeriod(startDate, endDate);
    var stats = {}, total = 0;
    for (var i = 0; i < incomes.length; i++) {
      var inc = incomes[i];
      if (!stats[inc.category]) stats[inc.category] = { total: 0, count: 0, items: [] };
      stats[inc.category].total += inc.amount;
      stats[inc.category].count++;
      stats[inc.category].items.push(inc);
      total += inc.amount;
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

  deleteIncomeRecord(recordId) {
    if (!this.detailsSheet) return false;
    var data = this.detailsSheet.getRange(2,1,this.detailsSheet.getLastRow()-1,9).getValues();
    for (var i = 0; i < data.length; i++) {
      if (data[i][7] === recordId) {
        var rowToDelete = i + 2;
        var amount = data[i][5];
        var date = data[i][0];
        this.detailsSheet.deleteRow(rowToDelete);
        this.recalculateYearSheetTotal(date);
        addLogEntry('Удалена запись о доходе: ' + formatMoney(amount) + ' от ' + formatDate(date), 'INFO', 'IncomeDetailsModule');
        return true;
      }
    }
    return false;
  }

  recalculateYearSheetTotal(date) {
    var day = date.getDate(), month = date.getMonth(), row = day + 1;
    var incomeCol = -1;
    var headers = this.yearSheet.getRange(1,1,1,this.yearSheet.getLastColumn()).getValues()[0];
    for (var col = 1; col < headers.length; col++) {
      var header = headers[col];
      if (header && header.indexOf(MONTHS[month]) !== -1 && header.indexOf('+') !== -1) { incomeCol = col + 1; break; }
    }
    if (incomeCol === -1) return;
    var start = new Date(date); start.setHours(0,0,0,0);
    var end = new Date(date); end.setHours(23,59,59,999);
    var dayIncomes = this.getIncomesByPeriod(start, end);
    var total = 0;
    for (var i = 0; i < dayIncomes.length; i++) total += dayIncomes[i].amount;
    if (total > 0) this.yearSheet.getRange(row, incomeCol).setValue(total);
    else this.yearSheet.getRange(row, incomeCol).clearContent();
    this.updateYearTotals();
  }

  getDayIncomes(date) {
    var start = new Date(date); start.setHours(0,0,0,0);
    var end = new Date(date); end.setHours(23,59,59,999);
    return this.getIncomesByPeriod(start, end);
  }

  showIncomeStats() {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('📊 Статистика доходов', 'Введите период в формате: ДД.ММ.ГГГГ - ДД.ММ.ГГГГ\n\nИли оставьте пустым для текущего месяца', ui.ButtonSet.OK_CANCEL);
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
    var stats = this.getCategoryStats(startDate, endDate);
    var report = '📊 СТАТИСТИКА ДОХОДОВ\nПериод: ' + formatDate(startDate) + ' - ' + formatDate(endDate) + '\n══════════════════════\n\n💰 ВСЕГО: ' + formatMoney(stats.total) + '\n\n';
    var categories = Object.keys(stats.stats);
    for (var i = 0; i < categories.length; i++) {
      var cat = categories[i];
      var data = stats.stats[cat];
      var percent = (data.total / stats.total * 100).toFixed(1);
      report += cat + ':\n   Сумма: ' + formatMoney(data.total) + '\n   Доля: ' + percent + '%\n   Кол-во: ' + data.count + '\n\n';
    }
    showMessage('Статистика доходов', report);
    addLogEntry('Показана статистика доходов', 'INFO', 'IncomeDetailsModule');
  }

  showInvestmentIncomeStats() {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('📊 Инвестиционные доходы', 'Введите период в формате: ДД.ММ.ГГГГ - ДД.ММ.ГГГГ\n\nИли оставьте пустым для текущего месяца', ui.ButtonSet.OK_CANCEL);
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
    var incomes = this.getIncomesByPeriod(startDate, endDate);
    var investmentCategories = ['Дивиденды', 'Купоны по облигациям', 'Продажа активов'];
    var investmentIncomes = incomes.filter(function(i) { return investmentCategories.indexOf(i.category) !== -1; });
    var total = investmentIncomes.reduce(function(s, i) { return s + i.amount; }, 0);
    var report = '📊 ИНВЕСТИЦИОННЫЕ ДОХОДЫ\nПериод: ' + formatDate(startDate) + ' - ' + formatDate(endDate) + '\n═══════════════════════════\n\n💰 ВСЕГО: ' + formatMoney(total) + '\n📊 Количество операций: ' + investmentIncomes.length + '\n\n';
    if (investmentIncomes.length > 0) {
      report += '📋 ДЕТАЛИЗАЦИЯ:\n';
      investmentIncomes.forEach(function(inc) {
        report += formatDate(inc.date) + ': ' + inc.category + ' - ' + formatMoney(inc.amount);
        if (inc.source && inc.source !== 'manual') report += ' [' + inc.source + ']';
        report += '\n';
      });
    } else report += 'Нет инвестиционных доходов за указанный период\n';
    showMessage('Инвестиционные доходы', report);
  }
}

function initIncomeDetails() { new IncomeDetailsModule().initDetailsSheet(); }
function showIncomeStats() { new IncomeDetailsModule().showIncomeStats(); }
function showInvestmentIncomeStats() { new IncomeDetailsModule().showInvestmentIncomeStats(); }