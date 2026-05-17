/**
 * БЮДЖЕТ - лист "Бюджет"
 */

class BudgetManager {
  constructor() {
    this.sheet = getSheetSafe(SHEETS.BUDGET);
    this.currentYear = CONFIG.CURRENT_YEAR;
    this.detailsSheet = getSheetSafe(SHEETS.EXPENSE_DETAILS);
  }

  getDaysInMonth(month) { return new Date(this.currentYear, month, 0).getDate(); }
  getMonthFactor(month) { var daysInMonth = this.getDaysInMonth(month); var daysInYear = (this.currentYear % 4 === 0) ? 366 : 365; return daysInMonth / daysInYear; }
  calculateMonthlyLimit(annualLimit, month) { return annualLimit * this.getMonthFactor(month); }

  initBudgetSheet() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEETS.BUDGET);
      if (sheet) {
        var ui = SpreadsheetApp.getUi();
        if (ui.alert('Инициализация бюджета', 'Лист бюджета уже существует. Пересоздать? Все данные будут потеряны.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
        ss.deleteSheet(sheet);
      }
      this.sheet = ss.insertSheet(SHEETS.BUDGET);
      this.sheet.setTabColor('#34a853');
      var headers = ['Категория', 'Годовой лимит', 'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь', 'Итого', 'Остаток', 'Месячный лимит', 'Отклонение'];
      this.sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      this.sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0').setHorizontalAlignment('center');
      var defaultCategories = [['Жильё', 360000], ['Продукты', 240000], ['Транспорт', 60000], ['Здоровье', 48000], ['Развлечения', 84000], ['Одежда', 60000], ['Связь', 18000], ['Долги', 120000], ['Инвестиции', 60000], ['Прочее', 36000]];
      var row = 2;
      for (var c = 0; c < defaultCategories.length; c++) {
        var category = defaultCategories[c][0];
        var limit = defaultCategories[c][1];
        this.sheet.getRange(row, 1).setValue(category);
        this.sheet.getRange(row, 2).setValue(limit).setNumberFormat('#,##0.00');
        for (var m = 0; m < 12; m++) {
          var col = m + 3;
          var monthNum = m + 1;
          var formula = '=SUMIFS(\'' + SHEETS.EXPENSE_DETAILS + '\'!B:B, \'' + SHEETS.EXPENSE_DETAILS + '\'!C:C, "' + category + '", MONTH(\'' + SHEETS.EXPENSE_DETAILS + '\'!A:A), ' + monthNum + ', YEAR(\'' + SHEETS.EXPENSE_DETAILS + '\'!A:A), ' + this.currentYear + ')';
          this.sheet.getRange(row, col).setFormula(formula);
        }
        this.sheet.getRange(row, 15).setFormula('=SUM(C' + row + ':N' + row + ')');
        this.sheet.getRange(row, 16).setFormula('=B' + row + '-O' + row);
        this.sheet.getRange(row, 17).setValue(0).setNumberFormat('#,##0.00');
        this.sheet.getRange(row, 18).setValue(0).setNumberFormat('#,##0.00');
        row++;
      }
      var totalRow = row;
      this.sheet.getRange(totalRow, 1).setValue('ИТОГО:').setFontWeight('bold');
      this.sheet.getRange(totalRow, 2).setFormula('=SUM(B2:B' + (row - 1) + ')');
      for (var colMonth = 3; colMonth <= 14; colMonth++) {
        var colLetter = String.fromCharCode(64 + colMonth);
        this.sheet.getRange(totalRow, colMonth).setFormula('=SUM(' + colLetter + '2:' + colLetter + (row - 1) + ')');
      }
      this.sheet.getRange(totalRow, 15).setFormula('=SUM(O2:O' + (row - 1) + ')');
      this.sheet.setColumnWidth(1, 200); for (var i = 2; i <= 18; i++) this.sheet.setColumnWidth(i, 100);
      this.sheet.getRange(2, 2, row, 17).setNumberFormat('#,##0.00').setHorizontalAlignment('right');
      this.sheet.setFrozenRows(1);
      var rule1 = SpreadsheetApp.newConditionalFormatRule().whenNumberGreaterThan(0).setBackground('#ffcdd2').setRanges([this.sheet.getRange(2, 18, row - 2, 1)]).build();
      var rule2 = SpreadsheetApp.newConditionalFormatRule().whenNumberLessThan(0).setBackground('#b6d7a8').setRanges([this.sheet.getRange(2, 18, row - 2, 1)]).build();
      var rules = this.sheet.getConditionalFormatRules(); rules.push(rule1, rule2); this.sheet.setConditionalFormatRules(rules);
      this.checkBudgetStatus();
      showMessage('Готово', '✅ Лист бюджета инициализирован', 'info');
      addLogEntry('Лист бюджета инициализирован', 'INFO', 'BudgetManager');
      return true;
    } catch(e) { logError(e, 'BudgetManager.initBudgetSheet'); showMessage('Ошибка', e.toString(), 'error'); return false; }
  }

  checkBudgetStatus() {
    if (!this.sheet) { this.initBudgetSheet(); return; }
    var lastRow = this.sheet.getLastRow();
    if (lastRow < 2) return;
    var data = this.sheet.getRange(2, 1, lastRow - 1, 18).getValues();
    var currentMonth = new Date().getMonth() + 1;
    var alerts = [];
    for (var i = 0; i < data.length - 1; i++) {
      var row = i + 2, rowData = data[i], category = rowData[0], annualLimit = rowData[1];
      if (category === 'ИТОГО:' || !category) continue;
      var monthSpent = rowData[currentMonth + 1] || 0, yearSpent = rowData[14] || 0;
      if (annualLimit <= 0) continue;
      var monthlyLimit = this.calculateMonthlyLimit(annualLimit, currentMonth);
      var monthPercent = monthlyLimit > 0 ? (monthSpent / monthlyLimit * 100) : 0;
      var deviation = monthSpent - monthlyLimit;
      this.sheet.getRange(row, 17).setValue(monthlyLimit).setNumberFormat('#,##0.00');
      this.sheet.getRange(row, 18).setValue(deviation).setNumberFormat('#,##0.00');
      var monthCell = this.sheet.getRange(row, currentMonth + 2);
      if (monthPercent >= BUDGET_CONFIG.CRITICAL_THRESHOLD) {
        monthCell.setBackground('#ffcdd2');
        if (BUDGET_CONFIG.ALERT_ON_EXCEED && monthSpent > monthlyLimit) alerts.push('🔴 КРИТИЧНО: "' + category + '" - превышение на ' + formatMoney(deviation) + ' (' + monthPercent.toFixed(1) + '%)');
      } else if (monthPercent >= BUDGET_CONFIG.WARNING_THRESHOLD) {
        monthCell.setBackground('#fff9c4');
        if (BUDGET_CONFIG.ALERT_ON_EXCEED) alerts.push('🟡 Внимание: "' + category + '" - использовано ' + monthPercent.toFixed(1) + '% месячного лимита');
      } else monthCell.setBackground('#ffffff');
      var yearPercent = annualLimit > 0 ? (yearSpent / annualLimit * 100) : 0;
      if (yearPercent >= 100) alerts.push('⚠️ Годовой лимит "' + category + '" исчерпан: ' + formatMoney(yearSpent) + ' из ' + formatMoney(annualLimit));
    }
    var totalRow = lastRow, totalSpent = this.sheet.getRange(totalRow, 15).getValue(), totalLimit = this.sheet.getRange(totalRow, 2).getValue();
    if (totalLimit > 0) {
      var totalPercent = (totalSpent / totalLimit * 100);
      if (totalPercent > 100) this.sheet.getRange(totalRow, 15).setBackground('#ffcdd2');
      else if (totalPercent > 90) this.sheet.getRange(totalRow, 15).setBackground('#fff9c4');
      else this.sheet.getRange(totalRow, 15).setBackground('#ffffff');
    }
    if (alerts.length > 0) this.sendBudgetAlerts(alerts);
    return alerts;
  }

  sendBudgetAlerts(alerts) {
    try {
      var email = Session.getActiveUser().getEmail();
      MailApp.sendEmail(email, '⚠️ Превышение бюджета', 'Обнаружены проблемы с бюджетом:\n\n' + alerts.join('\n\n') + '\n\nПроверьте лист "Бюджет" для деталей.');
      addLogEntry('Отправлены уведомления о превышении бюджета: ' + alerts.length, 'INFO', 'BudgetManager');
    } catch(e) { logError(e, 'BudgetManager.sendBudgetAlerts'); }
  }

  showBudgetReport() {
    security.checkAccess('showBudgetReport');
    if (!this.sheet) this.initBudgetSheet();
    var lastRow = this.sheet.getLastRow();
    if (lastRow < 2) { showMessage('Информация', 'Нет данных в бюджете', 'warning'); return; }
    var data = this.sheet.getRange(2, 1, lastRow - 1, 18).getValues();
    var currentMonth = new Date().getMonth() + 1;
    var report = '📊 ОТЧЁТ ПО БЮДЖЕТУ\n═══════════════════\n\n📅 Текущий месяц: ' + MONTHS[currentMonth - 1] + '\n\n📊 ПО КАТЕГОРИЯМ:\n────────────────\n';
    var totalAnnualLimit = 0, totalAnnualSpent = 0, totalMonthlyLimit = 0, totalMonthlySpent = 0, totalDeviation = 0;
    for (var i = 0; i < data.length - 1; i++) {
      var rowData = data[i], category = rowData[0];
      if (category === 'ИТОГО:' || !category) continue;
      var annualLimit = rowData[1] || 0, monthSpent = rowData[currentMonth + 1] || 0, yearSpent = rowData[14] || 0;
      if (annualLimit <= 0) continue;
      var monthlyLimit = this.calculateMonthlyLimit(annualLimit, currentMonth);
      var monthPercent = monthlyLimit > 0 ? (monthSpent / monthlyLimit * 100) : 0, yearPercent = annualLimit > 0 ? (yearSpent / annualLimit * 100) : 0, deviation = monthSpent - monthlyLimit;
      totalAnnualLimit += annualLimit; totalAnnualSpent += yearSpent; totalMonthlyLimit += monthlyLimit; totalMonthlySpent += monthSpent; totalDeviation += deviation;
      var emoji = '✅'; if (monthPercent > BUDGET_CONFIG.CRITICAL_THRESHOLD) emoji = '🔴'; else if (monthPercent > BUDGET_CONFIG.WARNING_THRESHOLD) emoji = '🟡';
      report += emoji + ' ' + category + ':\n   Месяц: ' + formatMoney(monthSpent) + ' / ' + formatMoney(monthlyLimit) + ' (' + monthPercent.toFixed(1) + '%)\n';
      if (deviation > 0) report += '   🔴 Перерасход: ' + formatMoney(deviation) + '\n'; else if (deviation < 0) report += '   ✅ Экономия: ' + formatMoney(Math.abs(deviation)) + '\n';
      report += '   Год: ' + formatMoney(yearSpent) + ' / ' + formatMoney(annualLimit) + ' (' + yearPercent.toFixed(1) + '%)\n\n';
    }
    report += '═══════════════════\n💰 ЗА МЕСЯЦ:\n   План: ' + formatMoney(totalMonthlyLimit) + '\n   Факт: ' + formatMoney(totalMonthlySpent) + '\n   ' + (totalMonthlySpent / totalMonthlyLimit * 100).toFixed(1) + '% от бюджета\n';
    if (totalDeviation > 0) report += '   🔴 Общий перерасход: ' + formatMoney(totalDeviation) + '\n\n'; else if (totalDeviation < 0) report += '   ✅ Общая экономия: ' + formatMoney(Math.abs(totalDeviation)) + '\n\n'; else report += '   ✅ Точное выполнение бюджета\n\n';
    report += '📅 ЗА ГОД:\n   План: ' + formatMoney(totalAnnualLimit) + '\n   Факт: ' + formatMoney(totalAnnualSpent) + '\n   Остаток: ' + formatMoney(totalAnnualLimit - totalAnnualSpent);
    showMessage('Бюджетный отчёт', report);
  }

  refreshBudgetDirect() {
    if (!this.sheet) this.initBudgetSheet();
    if (!this.detailsSheet) { showMessage('Ошибка', 'Лист "Детали расходов" не найден', 'error'); return; }
    var lastRow = this.sheet.getLastRow() - 1;
    if (lastRow < 2) { showMessage('Информация', 'Нет данных в бюджете', 'warning'); return; }
    var detailsLastRow = this.detailsSheet.getLastRow();
    if (detailsLastRow < 2) { showMessage('Информация', 'Нет данных в детализации', 'warning'); return; }
    var detailsData = this.detailsSheet.getRange(2, 1, detailsLastRow - 1, 5).getValues();
    var categories = this.sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var monthlySums = {};
    for (var i = 0; i < categories.length; i++) { var cat = categories[i][0]; if (!cat || cat === 'ИТОГО:') continue; monthlySums[cat] = {}; for (var m = 1; m <= 12; m++) monthlySums[cat][m] = 0; }
    for (var d = 0; d < detailsData.length; d++) { var row = detailsData[d], date = row[0], amount = row[1], category = row[2]; if (!date || !amount || !category) continue; var year = date.getFullYear(), month = date.getMonth() + 1; if (year === this.currentYear && monthlySums[category]) monthlySums[category][month] += amount; }
    for (var catIdx = 0; catIdx < categories.length; catIdx++) { var rowNum = catIdx + 2, catName = categories[catIdx][0]; if (!catName || catName === 'ИТОГО:') continue; for (var monthNum = 1; monthNum <= 12; monthNum++) { var colNum = monthNum + 2; this.sheet.getRange(rowNum, colNum).setValue(monthlySums[catName][monthNum]).setNumberFormat('#,##0.00'); } this.sheet.getRange(rowNum, 15).setFormula('=SUM(C' + rowNum + ':N' + rowNum + ')'); this.sheet.getRange(rowNum, 16).setFormula('=B' + rowNum + '-O' + rowNum); }
    this.checkBudgetStatus();
    showMessage('Готово', 'Бюджет обновлён прямым расчётом', 'info');
    addLogEntry('Бюджет обновлён прямым расчётом', 'INFO', 'BudgetManager');
  }

  setCategoryLimit(category, limit) {
    security.checkAccess('setCategoryLimit');
    if (!this.sheet) this.initBudgetSheet();
    var lastRow = this.sheet.getLastRow() - 1;
    var categories = this.sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var found = false;
    for (var i = 0; i < categories.length; i++) { if (categories[i][0] === category) { var row = i + 2; this.sheet.getRange(row, 2).setValue(limit).setNumberFormat('#,##0.00'); found = true; break; } }
    if (!found) { var newRow = this.sheet.getLastRow(); this.sheet.getRange(newRow, 1).setValue(category); this.sheet.getRange(newRow, 2).setValue(limit).setNumberFormat('#,##0.00'); for (var m = 0; m < 12; m++) { var col = m + 3, monthNum = m + 1; var formula = '=SUMIFS(\'' + SHEETS.EXPENSE_DETAILS + '\'!B:B, \'' + SHEETS.EXPENSE_DETAILS + '\'!C:C, "' + category + '", MONTH(\'' + SHEETS.EXPENSE_DETAILS + '\'!A:A), ' + monthNum + ', YEAR(\'' + SHEETS.EXPENSE_DETAILS + '\'!A:A), ' + this.currentYear + ')'; this.sheet.getRange(newRow, col).setFormula(formula); } this.sheet.getRange(newRow, 15).setFormula('=SUM(C' + newRow + ':N' + newRow + ')'); this.sheet.getRange(newRow, 16).setFormula('=B' + newRow + '-O' + newRow); }
    this.checkBudgetStatus();
    showMessage('Успешно', 'Лимит для категории "' + category + '" установлен: ' + formatMoney(limit), 'info');
    addLogEntry('Установлен лимит для категории "' + category + '": ' + formatMoney(limit), 'INFO', 'BudgetManager');
    return true;
  }

  setLimitInteractive() {
    var ui = SpreadsheetApp.getUi();
    var categories = this.getCategories();
    var message = 'Выберите категорию (введите номер):\n\n';
    for (var i = 0; i < categories.length; i++) message += (i + 1) + '. ' + categories[i] + '\n';
    var response = ui.prompt('Выбор категории', message, ui.ButtonSet.OK_CANCEL);
    if (response.getSelectedButton() !== ui.Button.OK) return;
    var catIndex = parseInt(response.getResponseText()) - 1;
    if (isNaN(catIndex) || catIndex < 0 || catIndex >= categories.length) { ui.alert('Ошибка', 'Неверный номер категории', ui.ButtonSet.OK); return; }
    var category = categories[catIndex];
    var limitResponse = ui.prompt('Установка лимита', 'Введите ГОДОВОЙ лимит для категории "' + category + '" в рублях:\n\n(Месячные лимиты будут рассчитаны автоматически с учётом дней в месяце)', ui.ButtonSet.OK_CANCEL);
    if (limitResponse.getSelectedButton() !== ui.Button.OK) return;
    var limit = safeParseNumber(limitResponse.getResponseText());
    if (limit <= 0) { ui.alert('Ошибка', 'Лимит должен быть положительным числом', ui.ButtonSet.OK); return; }
    this.setCategoryLimit(category, limit);
  }

  getCategories() {
    var categories = [];
    if (this.sheet) { var lastRow = this.sheet.getLastRow(); if (lastRow >= 2) { var data = this.sheet.getRange(2, 1, lastRow - 1, 1).getValues(); for (var i = 0; i < data.length; i++) { if (data[i][0] && data[i][0] !== 'ИТОГО:') categories.push(data[i][0]); } } }
    var defaultCats = ['Жильё', 'Продукты', 'Транспорт', 'Здоровье', 'Развлечения', 'Одежда', 'Связь', 'Долги', 'Инвестиции', 'Прочее'];
    for (var d = 0; d < defaultCats.length; d++) if (categories.indexOf(defaultCats[d]) === -1) categories.push(defaultCats[d]);
    categories.sort();
    return categories;
  }

  showBudgetRecommendations() {
    if (!this.sheet) this.initBudgetSheet();
    var lastRow = this.sheet.getLastRow();
    if (lastRow < 2) { showMessage('Информация', 'Нет данных в бюджете', 'warning'); return; }
    var data = this.sheet.getRange(2, 1, lastRow - 1, 18).getValues();
    var currentMonth = new Date().getMonth() + 1;
    var recommendations = [], totalSavings = 0;
    for (var i = 0; i < data.length - 1; i++) {
      var row = data[i], category = row[0];
      if (category === 'ИТОГО:' || !category) continue;
      var annualLimit = row[1] || 0, monthSpent = row[currentMonth + 1] || 0, yearSpent = row[14] || 0;
      if (annualLimit <= 0) continue;
      var monthlyLimit = this.calculateMonthlyLimit(annualLimit, currentMonth);
      var monthPercent = (monthSpent / monthlyLimit * 100), yearPercent = (yearSpent / annualLimit * 100), deviation = monthSpent - monthlyLimit;
      if (monthPercent > 100) { recommendations.push({ priority: 'high', category: category, message: 'Превышение месячного лимита на ' + formatMoney(deviation) + ' (' + monthPercent.toFixed(1) + '%)', saving: deviation * 0.3 }); totalSavings += deviation * 0.3; }
      else if (monthPercent > 80) { var potentialSaving = (monthlyLimit - monthSpent) * 0.2; recommendations.push({ priority: 'medium', category: category, message: 'Близко к лимиту: ' + monthPercent.toFixed(1) + '%', saving: potentialSaving }); totalSavings += potentialSaving; }
      if (yearPercent > 90 && yearPercent < 100) { var remaining = annualLimit - yearSpent; recommendations.push({ priority: 'medium', category: category, message: 'Годовой лимит почти исчерпан (' + yearPercent.toFixed(1) + '%). Осталось ' + formatMoney(remaining), saving: remaining * 0.5 }); }
      if (category === 'Продукты' && monthSpent > 25000) { recommendations.push({ priority: 'medium', category: category, message: 'Высокие расходы на продукты: ' + formatMoney(monthSpent) + '. Рекомендуется оптимизировать закупки.', saving: monthSpent * 0.1 }); totalSavings += monthSpent * 0.1; }
      if (category === 'Развлечения' && monthSpent > 8000) { recommendations.push({ priority: 'medium', category: category, message: 'Расходы на развлечения выше среднего: ' + formatMoney(monthSpent), saving: monthSpent * 0.2 }); totalSavings += monthSpent * 0.2; }
    }
    var report = '💡 РЕКОМЕНДАЦИИ ПО БЮДЖЕТУ\n══════════════════════════\n\n';
    if (recommendations.length > 0) {
      recommendations.sort(function(a,b) { return (b.priority === 'high' ? 2 : 1) - (a.priority === 'high' ? 2 : 1); });
      for (var r = 0; r < recommendations.length; r++) { var rec = recommendations[r]; report += (rec.priority === 'high' ? '🔴' : '🟡') + ' ' + rec.category + ':\n   ' + rec.message + '\n   💰 Потенциальная экономия: ' + formatMoney(rec.saving) + '\n\n'; }
      report += '══════════════════════════\n💰 Общая потенциальная экономия: ' + formatMoney(totalSavings);
    } else report += '✅ Отличная работа! Все лимиты соблюдаются.\n';
    showMessage('Рекомендации', report);
  }

  exportBudget() {
    if (!this.sheet) this.initBudgetSheet();
    var lastRow = this.sheet.getLastRow(), lastCol = this.sheet.getLastColumn();
    var data = this.sheet.getRange(1, 1, lastRow, lastCol).getValues();
    var csv = '';
    for (var i = 0; i < data.length; i++) csv += data[i].join(',') + '\n';
    var blob = Utilities.newBlob(csv, 'text/csv', 'budget_' + this.currentYear + '.csv');
    var file = DriveApp.createFile(blob);
    showMessage('Экспорт готов', 'Бюджет экспортирован: ' + file.getUrl());
    addLogEntry('Бюджет экспортирован в CSV', 'INFO', 'BudgetManager');
  }

  fixBudgetFormulas() {
    if (!this.sheet) this.initBudgetSheet();
    var lastRow = this.sheet.getLastRow() - 1;
    if (lastRow < 2) return;
    var categories = this.sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < categories.length; i++) { var row = i + 2, category = categories[i][0]; if (!category || category === 'ИТОГО:') continue; for (var m = 0; m < 12; m++) { var col = m + 3, monthNum = m + 1; var formula = '=SUMIFS(\'' + SHEETS.EXPENSE_DETAILS + '\'!B:B, \'' + SHEETS.EXPENSE_DETAILS + '\'!C:C, "' + category + '", MONTH(\'' + SHEETS.EXPENSE_DETAILS + '\'!A:A), ' + monthNum + ', YEAR(\'' + SHEETS.EXPENSE_DETAILS + '\'!A:A), ' + this.currentYear + ')'; this.sheet.getRange(row, col).setFormula(formula); } this.sheet.getRange(row, 15).setFormula('=SUM(C' + row + ':N' + row + ')'); this.sheet.getRange(row, 16).setFormula('=B' + row + '-O' + row); }
    var totalRow = lastRow + 1;
    this.sheet.getRange(totalRow, 2).setFormula('=SUM(B2:B' + lastRow + ')');
    for (var colNum = 3; colNum <= 14; colNum++) { var colLetter = String.fromCharCode(64 + colNum); this.sheet.getRange(totalRow, colNum).setFormula('=SUM(' + colLetter + '2:' + colLetter + lastRow + ')'); }
    this.sheet.getRange(totalRow, 15).setFormula('=SUM(O2:O' + lastRow + ')');
    this.checkBudgetStatus();
    showMessage('Готово', 'Формулы бюджета восстановлены', 'info');
    addLogEntry('Формулы бюджета восстановлены', 'INFO', 'BudgetManager');
  }

  validateDetailsData() {
    if (!this.detailsSheet) { showMessage('Ошибка', 'Лист "Детали расходов" не найден', 'error'); return; }
    var lastRow = this.detailsSheet.getLastRow();
    if (lastRow < 2) { showMessage('Информация', 'Нет данных в детализации', 'warning'); return; }
    var data = this.detailsSheet.getRange(2, 1, lastRow - 1, 5).getValues();
    var issues = [];
    for (var i = 0; i < data.length; i++) { var row = data[i], rowNum = i + 2, date = row[0], amount = row[1], category = row[2], id = row[4]; if (!date || isNaN(date.getTime())) issues.push('Строка ' + rowNum + ': Неверная дата'); if (amount <= 0) issues.push('Строка ' + rowNum + ': Сумма должна быть положительной (' + amount + ')'); if (!category) issues.push('Строка ' + rowNum + ': Отсутствует категория'); if (!id) issues.push('Строка ' + rowNum + ': Отсутствует ID'); }
    if (issues.length > 0) showMessage('Найдены проблемы', issues.slice(0, 10).join('\n') + (issues.length > 10 ? '\n... и ещё ' + (issues.length - 10) : ''), 'warning');
    else showMessage('Проверка пройдена', 'Все ' + data.length + ' записей корректны', 'info');
  }
}

function initBudgetSheet() { new BudgetManager().initBudgetSheet(); }
function refreshBudget() { new BudgetManager().refreshBudgetDirect(); }
function showBudgetReport() { new BudgetManager().showBudgetReport(); }
function setCategoryLimit() { new BudgetManager().setLimitInteractive(); }
function checkBudgetStatus() { new BudgetManager().checkBudgetStatus(); }
function showBudgetRecommendations() { new BudgetManager().showBudgetRecommendations(); }
function exportBudget() { new BudgetManager().exportBudget(); }
function fixBudgetFormulas() { new BudgetManager().fixBudgetFormulas(); }
function validateDetailsData() { new BudgetManager().validateDetailsData(); }