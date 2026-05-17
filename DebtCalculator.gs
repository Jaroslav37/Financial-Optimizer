/**
 * ДОЛГИ - лист "Долговые обязательства"
 */

class DebtCalculator {
  constructor() { this.sheet = getSheetSafe(SHEETS.DEBTS); this.currentYear = CONFIG.CURRENT_YEAR; }

  calculateWithInterest(amount, rate, startDate, today, type) {
    if (amount <= 0) return 0;
    if (rate <= 0) return amount;
    var daysDiff = Math.ceil((today - startDate) / (1000 * 60 * 60 * 24));
    if (daysDiff <= 0) return amount;
    var years = daysDiff / 365;
    if (type === 'simple') return Number((amount * (1 + (rate / 100) * years)).toFixed(2));
    else return Number((amount * Math.pow(1 + (rate / 100), years)).toFixed(2));
  }

  calculateAllDebts() {
    security.checkAccess('calculateAllDebts');
    var sheet = getSheetSafe(SHEETS.DEBTS);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) { showMessage('Информация', 'Нет данных для расчёта', 'warning'); return; }
    var data = sheet.getRange(2, 1, lastRow - 1, 8).getValues();
    var today = new Date();
    var totalOwedToMe = 0, totalIOwe = 0, debtsOwedToMe = [], debtsIOwe = [];
    for (var i = 0; i < data.length; i++) {
      var row = i + 2, name = data[i][0], obligation = data[i][1], externalDebt = data[i][2], rate = data[i][3], startDate = data[i][5], type = data[i][6];
      if (!name || name.toString().trim() === '' || name === 'СВОДКА:' || name === 'В том числе:' || name === 'ИТОГО:') continue;
      var nameStr = name.toString().trim(), rateNum = safeParseNumber(rate), startDateObj = safeParseDate(startDate), interestType = (type && type.toString().toLowerCase().indexOf('simple') !== -1) ? 'simple' : 'compound';
      var rowTotal = 0, rowOwedToMe = 0, rowIOwe = 0;
      var obligationAmount = safeParseNumber(obligation);
      if (obligationAmount > 0) {
        var totalWithInterest = obligationAmount;
        if (startDateObj && rateNum > 0) totalWithInterest = this.calculateWithInterest(obligationAmount, rateNum, startDateObj, today, interestType);
        rowTotal += totalWithInterest; rowOwedToMe = totalWithInterest; totalOwedToMe += totalWithInterest;
        debtsOwedToMe.push({ name: nameStr, amount: totalWithInterest, original: obligationAmount, rate: rateNum });
      }
      var externalAmount = safeParseNumber(externalDebt);
      if (externalAmount > 0) {
        var totalWithInterestExt = externalAmount;
        if (startDateObj && rateNum > 0) totalWithInterestExt = this.calculateWithInterest(externalAmount, rateNum, startDateObj, today, interestType);
        rowTotal += totalWithInterestExt; rowIOwe = totalWithInterestExt; totalIOwe += totalWithInterestExt;
        debtsIOwe.push({ name: nameStr, amount: totalWithInterestExt, original: externalAmount, rate: rateNum });
      }
      if (rowTotal > 0) { sheet.getRange(row, 5).setValue(rowTotal).setNumberFormat('#,##0.00'); var cell = sheet.getRange(row, 5); if (rowOwedToMe > 0 && rowIOwe === 0) cell.setBackground('#b6d7a8'); else if (rowOwedToMe === 0 && rowIOwe > 0) cell.setBackground('#ffcdd2'); else if (rowOwedToMe > 0 && rowIOwe > 0) cell.setBackground('#fff9c4'); }
    }
    var netBalance = totalOwedToMe - totalIOwe;
    var report = '💰 ОТЧЁТ ПО ДОЛГАМ\n══════════════════\n\n📌 Мне должны (всего): ' + formatMoney(totalOwedToMe) + '\n🌍 Я должен (всего): ' + formatMoney(totalIOwe) + '\n💎 Чистый баланс: ' + formatMoney(netBalance) + '\n\n';
    if (debtsOwedToMe.length) { report += '📌 КТО ДОЛЖЕН МНЕ:\n'; debtsOwedToMe.forEach(function(d) { report += '   • ' + d.name + ': ' + formatMoney(d.amount) + (d.rate > 0 ? ' (' + d.rate + '% годовых, исходная: ' + formatMoney(d.original) + ')' : '') + '\n'; }); report += '\n'; }
    if (debtsIOwe.length) { report += '🌍 КОМУ ДОЛЖЕН Я:\n'; debtsIOwe.forEach(function(d) { report += '   • ' + d.name + ': ' + formatMoney(d.amount) + (d.rate > 0 ? ' (' + d.rate + '% годовых, исходная: ' + formatMoney(d.original) + ')' : '') + '\n'; }); report += '\n'; }
    showMessage('Расчёт долгов завершён', report);
    addLogEntry('Расчёт долгов выполнен', 'INFO', 'DebtCalculator');
  }

  initDebtsSheet() {
    security.checkAccess('initDebtsSheet');
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEETS.DEBTS);
      if (!sheet) sheet = ss.insertSheet(SHEETS.DEBTS);
      else { var ui = SpreadsheetApp.getUi(); if (ui.alert('Пересоздать лист?', 'Лист долгов уже существует. Пересоздать с правильной структурой? Все данные будут потеряны.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return; sheet.clear(); }
      var headers = [['Имя', 'Мне должны', 'Я должен', 'Ставка %', 'С учётом %', 'Дата начала', 'Тип %', 'Примечания']];
      sheet.getRange(1, 1, 1, 8).setValues(headers);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f0f0f0');
      sheet.setColumnWidth(1,200); sheet.setColumnWidth(2,120); sheet.setColumnWidth(3,120); sheet.setColumnWidth(4,80); sheet.setColumnWidth(5,120); sheet.setColumnWidth(6,100); sheet.setColumnWidth(7,100); sheet.setColumnWidth(8,300);
      var examples = [['Иван Петров', '50000', '', '12', '', '01.01.2025', 'compound', 'Друг, должен за ремонт'], ['Сбербанк', '', '100000', '15', '', '15.03.2025', 'simple', 'Кредитная карта']];
      sheet.getRange(2, 1, examples.length, 8).setValues(examples);
      if (sheet.getLastRow() >= 2) { sheet.getRange(2, 2, sheet.getLastRow()-1, 1).setNumberFormat('#,##0.00'); sheet.getRange(2, 3, sheet.getLastRow()-1, 1).setNumberFormat('#,##0.00'); sheet.getRange(2, 4, sheet.getLastRow()-1, 1).setNumberFormat('0.00'); sheet.getRange(2, 5, sheet.getLastRow()-1, 1).setNumberFormat('#,##0.00'); sheet.getRange(2, 6, sheet.getLastRow()-1, 1).setNumberFormat('dd.mm.yyyy'); }
      sheet.getRange(1, 1).setNote('Структура листа долгов:\nA - Имя\nB - Мне должны (сумма)\nC - Я должен (сумма)\nD - Процентная ставка (% годовых)\nE - Сумма с процентами (заполняется скриптом)\nF - Дата начала\nG - Тип процентов (simple/compound)\nH - Примечания');
      showMessage('Готово', 'Лист долгов инициализирован', 'info');
      addLogEntry('Лист долгов инициализирован', 'INFO', 'DebtCalculator');
      this.calculateAllDebts();
    } catch(e) { logError(e, 'DebtCalculator.initDebtsSheet'); showMessage('Ошибка', e.toString(), 'error'); }
  }

  showDebtReport() {
    security.checkAccess('showDebtReport');
    var sheet = getSheetSafe(SHEETS.DEBTS);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) { showMessage('Информация', 'Нет данных о долгах', 'warning'); return; }
    var data = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    var report = '💰 ОТЧЁТ ПО ДОЛГАМ\n══════════════════\n\n';
    var totalOwedToMe = 0, totalIOwe = 0;
    for (var i = 0; i < data.length; i++) {
      var row = data[i], name = row[0]; if (name === 'СВОДКА:' || name === 'В том числе:') continue;
      var obligation = safeParseNumber(row[1]), external = safeParseNumber(row[2]), rate = row[3], startDate = row[5], type = row[6];
      if (obligation > 0) { totalOwedToMe += obligation; report += '📌 ' + name + ' должен мне: ' + formatMoney(obligation); if (rate > 0) report += ' (' + rate + '% годовых, ' + (type || 'compound') + ', с ' + formatDate(startDate) + ')'; report += '\n'; }
      if (external > 0) { totalIOwe += external; report += '🌍 Я должен ' + name + ': ' + formatMoney(external); if (rate > 0) report += ' (' + rate + '% годовых, ' + (type || 'compound') + ', с ' + formatDate(startDate) + ')'; report += '\n'; }
    }
    report += '\n══════════════════\n📌 Всего должны мне: ' + formatMoney(totalOwedToMe) + '\n🌍 Всего я должен: ' + formatMoney(totalIOwe) + '\n💎 Чистый баланс: ' + formatMoney(totalOwedToMe - totalIOwe);
    showMessage('Долги', report);
  }

  addNewDebtInteractive() {
    security.checkAccess('addNewDebtInteractive');
    var ui = SpreadsheetApp.getUi();
    var typeResponse = ui.alert('💰 Добавление долга - Шаг 1/5', 'Это долг ВАМ или ВЫ должны?', ui.ButtonSet.YES_NO_CANCEL);
    if (typeResponse === ui.Button.CANCEL) return;
    var type = (typeResponse === ui.Button.YES) ? 'obligation' : 'debt';
    var nameResponse = ui.prompt('💰 Добавление долга - Шаг 2/5', type === 'obligation' ? 'Введите имя человека/организации, КОТОРЫЙ ДОЛЖЕН ВАМ:' : 'Введите имя человека/организации, КОТОРОЙ ВЫ ДОЛЖНЫ:', ui.ButtonSet.OK_CANCEL);
    if (nameResponse.getSelectedButton() !== ui.Button.OK) return;
    var name = nameResponse.getResponseText().trim();
    if (!name) { ui.alert('Ошибка', 'Имя не может быть пустым', ui.ButtonSet.OK); return; }
    var amountResponse = ui.prompt('💰 Добавление долга - Шаг 3/5', 'Введите сумму долга (в рублях):', ui.ButtonSet.OK_CANCEL);
    if (amountResponse.getSelectedButton() !== ui.Button.OK) return;
    var amount = safeParseNumber(amountResponse.getResponseText());
    if (amount <= 0) { ui.alert('Ошибка', 'Сумма должна быть положительным числом', ui.ButtonSet.OK); return; }
    var rateChoice = ui.alert('💰 Добавление долга - Шаг 4/5', 'Хотите указать процентную ставку?', ui.ButtonSet.YES_NO);
    var rate = 0, startDate = null, interestType = 'compound';
    if (rateChoice === ui.Button.YES) {
      var rateResponse = ui.prompt('💰 Добавление долга - Шаг 4/5 (продолжение)', 'Введите ГОДОВУЮ процентную ставку (например: 12 для 12% годовых):', ui.ButtonSet.OK_CANCEL);
      if (rateResponse.getSelectedButton() !== ui.Button.OK) return;
      rate = safeParseNumber(rateResponse.getResponseText());
      if (rate < 0) { ui.alert('Ошибка', 'Ставка не может быть отрицательной', ui.ButtonSet.OK); return; }
      if (rate > 0) {
        var dateResponse = ui.prompt('💰 Добавление долга - Шаг 4/5 (продолжение)', 'Введите дату начала начисления процентов (ДД.ММ.ГГГГ):', ui.ButtonSet.OK_CANCEL);
        if (dateResponse.getSelectedButton() === ui.Button.OK) { var dateStr = dateResponse.getResponseText().trim(); if (dateStr) { startDate = safeParseDate(dateStr); if (!startDate) { ui.alert('Ошибка', 'Неверный формат даты', ui.ButtonSet.OK); return; } } }
        var typeChoice = ui.alert('💰 Добавление долга - Шаг 5/5', 'Выберите тип процентов:', ui.ButtonSet.YES_NO_CANCEL);
        if (typeChoice === ui.Button.CANCEL) return;
        interestType = (typeChoice === ui.Button.YES) ? 'simple' : 'compound';
      }
    }
    this.saveDebt(name, type, amount, rate, startDate, interestType);
  }

  saveDebt(name, type, amount, rate, startDate, interestType) {
    var sheet = getSheetSafe(SHEETS.DEBTS);
    if (!sheet) return;
    var lastRow = sheet.getLastRow();
    var targetRow = -1;
    if (lastRow >= 2) { var names = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); for (var i = 0; i < names.length; i++) if (names[i][0] === name) { targetRow = i + 2; break; } }
    if (targetRow === -1) { targetRow = sheet.getLastRow() + 1; sheet.getRange(targetRow, 1).setValue(name); if (type === 'obligation') sheet.getRange(targetRow, 2).setValue(amount).setNumberFormat('#,##0.00'); else sheet.getRange(targetRow, 3).setValue(amount).setNumberFormat('#,##0.00'); if (rate > 0) sheet.getRange(targetRow, 4).setValue(rate).setNumberFormat('0.00'); if (startDate) sheet.getRange(targetRow, 6).setValue(startDate).setNumberFormat('dd.mm.yyyy'); if (rate > 0) sheet.getRange(targetRow, 7).setValue(interestType); var note = []; note.push('Тип: ' + (type === 'obligation' ? 'Долг ВАМ' : 'ВЫ должны')); note.push('Сумма: ' + formatMoney(amount)); if (rate > 0) note.push('Ставка: ' + rate + '% годовых (' + (interestType === 'simple' ? 'простые' : 'сложные') + ')'); if (startDate) note.push('С даты: ' + formatDate(startDate)); else note.push('Без процентов'); sheet.getRange(targetRow, 8).setValue(note.join('\n')); var range = sheet.getRange(targetRow, 1, 1, 8); if (type === 'obligation') range.setBackground('#e6f4ea'); else range.setBackground('#fce8e8'); }
    this.calculateAllDebts();
    addLogEntry('Долг добавлен: ' + name + ', сумма: ' + formatMoney(amount), 'INFO', 'DebtCalculator');
  }
}

function calculateAllDebts() { new DebtCalculator().calculateAllDebts(); }
function initDebtsSheet() { new DebtCalculator().initDebtsSheet(); }
function showDebtReport() { new DebtCalculator().showDebtReport(); }
function addNewDebtInteractive() { new DebtCalculator().addNewDebtInteractive(); }
function addNewDebtQuick() { addNewDebtInteractive(); }
function markDebtAsPaid() { showMessage('Информация', 'Функция в разработке', 'info'); }