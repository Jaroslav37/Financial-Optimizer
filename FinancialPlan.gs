/**
 * ФИНАНСОВЫЙ ПЛАН - лист "Финансовый план"
 */

class FinancialPlanManager {
  constructor() {
    this.planSheet = getSheetSafe(SHEETS.FINANCIAL_PLAN);
    this.currentYear = CONFIG.CURRENT_YEAR;
  }

  getYearRow(year) { return 7 + (year - 2024); }

  updateFinancialPlan() {
    security.checkAccess('updateFinancialPlan');
    if (!this.planSheet) { addLogEntry('Лист "Финансовый план" не найден', 'WARNING', 'FinancialPlan'); return; }
    var year = this.currentYear;
    var yearSheetName = year.toString();
    var yearSheet = getSheetSafe(yearSheetName);
    if (!yearSheet) { showMessage('Ошибка', `Лист "${yearSheetName}" не найден`, 'error'); return; }
    var targetRow = this.findOrCreateYearRow(year);
    var yearRow = this.getYearRow(year);
    this.planSheet.getRange(targetRow, 1).setValue(year);
    // B: доходы (чётные колонки TOTAL)
    var incomeFormula = `=SUM('${yearSheetName}'!C${yearRow},'${yearSheetName}'!E${yearRow},'${yearSheetName}'!G${yearRow},'${yearSheetName}'!I${yearRow},'${yearSheetName}'!K${yearRow},'${yearSheetName}'!M${yearRow},'${yearSheetName}'!O${yearRow},'${yearSheetName}'!Q${yearRow},'${yearSheetName}'!S${yearRow},'${yearSheetName}'!U${yearRow},'${yearSheetName}'!W${yearRow},'${yearSheetName}'!Y${yearRow})`;
    this.planSheet.getRange(targetRow, 2).setFormula(incomeFormula);
    // C: расходы (нечётные колонки TOTAL)
    var expenseFormula = `=SUM('${yearSheetName}'!B${yearRow},'${yearSheetName}'!D${yearRow},'${yearSheetName}'!F${yearRow},'${yearSheetName}'!H${yearRow},'${yearSheetName}'!J${yearRow},'${yearSheetName}'!L${yearRow},'${yearSheetName}'!N${yearRow},'${yearSheetName}'!P${yearRow},'${yearSheetName}'!R${yearRow},'${yearSheetName}'!T${yearRow},'${yearSheetName}'!V${yearRow},'${yearSheetName}'!X${yearRow})`;
    this.planSheet.getRange(targetRow, 3).setFormula(expenseFormula);
    // D: прибыль
    var profitFormula = `=${this.getColumnLetter(2)}${targetRow}-${this.getColumnLetter(3)}${targetRow}-${this.getColumnLetter(6)}${targetRow}-${this.getColumnLetter(7)}${targetRow}-${this.getColumnLetter(9)}${targetRow}-${this.getColumnLetter(5)}${targetRow}`;
    this.planSheet.getRange(targetRow, 4).setFormula(profitFormula);
    // E: инвестиции (PROJECTS + COMPENSATING2)
    var investmentsFormula = `=SUM('${yearSheetName}'!B${ROWS.PROJECTS},'${yearSheetName}'!D${ROWS.PROJECTS},'${yearSheetName}'!F${ROWS.PROJECTS},'${yearSheetName}'!H${ROWS.PROJECTS},'${yearSheetName}'!J${ROWS.PROJECTS},'${yearSheetName}'!L${ROWS.PROJECTS},'${yearSheetName}'!N${ROWS.PROJECTS},'${yearSheetName}'!P${ROWS.PROJECTS},'${yearSheetName}'!R${ROWS.PROJECTS},'${yearSheetName}'!T${ROWS.PROJECTS},'${yearSheetName}'!V${ROWS.PROJECTS},'${yearSheetName}'!X${ROWS.PROJECTS}) + SUM('${yearSheetName}'!B${ROWS.COMPENSATING2},'${yearSheetName}'!D${ROWS.COMPENSATING2},'${yearSheetName}'!F${ROWS.COMPENSATING2},'${yearSheetName}'!H${ROWS.COMPENSATING2},'${yearSheetName}'!J${ROWS.COMPENSATING2},'${yearSheetName}'!L${ROWS.COMPENSATING2},'${yearSheetName}'!N${ROWS.COMPENSATING2},'${yearSheetName}'!P${ROWS.COMPENSATING2},'${yearSheetName}'!R${ROWS.COMPENSATING2},'${yearSheetName}'!T${ROWS.COMPENSATING2},'${yearSheetName}'!V${ROWS.COMPENSATING2},'${yearSheetName}'!X${ROWS.COMPENSATING2})`;
    this.planSheet.getRange(targetRow, 5).setFormula(investmentsFormula);
    // F: страховка
    var insuranceFormula = `=SUM('${yearSheetName}'!B${ROWS.INSURANCE},'${yearSheetName}'!D${ROWS.INSURANCE},'${yearSheetName}'!F${ROWS.INSURANCE},'${yearSheetName}'!H${ROWS.INSURANCE},'${yearSheetName}'!J${ROWS.INSURANCE},'${yearSheetName}'!L${ROWS.INSURANCE},'${yearSheetName}'!N${ROWS.INSURANCE},'${yearSheetName}'!P${ROWS.INSURANCE},'${yearSheetName}'!R${ROWS.INSURANCE},'${yearSheetName}'!T${ROWS.INSURANCE},'${yearSheetName}'!V${ROWS.INSURANCE},'${yearSheetName}'!X${ROWS.INSURANCE})`;
    this.planSheet.getRange(targetRow, 6).setFormula(insuranceFormula);
    // G: пенсия
    var pensionFormula = `=SUM('${yearSheetName}'!B${ROWS.PENSION},'${yearSheetName}'!D${ROWS.PENSION},'${yearSheetName}'!F${ROWS.PENSION},'${yearSheetName}'!H${ROWS.PENSION},'${yearSheetName}'!J${ROWS.PENSION},'${yearSheetName}'!L${ROWS.PENSION},'${yearSheetName}'!N${ROWS.PENSION},'${yearSheetName}'!P${ROWS.PENSION},'${yearSheetName}'!R${ROWS.PENSION},'${yearSheetName}'!T${ROWS.PENSION},'${yearSheetName}'!V${ROWS.PENSION},'${yearSheetName}'!X${ROWS.PENSION})`;
    this.planSheet.getRange(targetRow, 7).setFormula(pensionFormula);
    // H: резерв
    this.planSheet.getRange(targetRow, 8).setValue(20000);
    // I: подарки (GIFTS + COMPENSATING)
    var giftsFormula = `=SUM('${yearSheetName}'!B${ROWS.GIFTS},'${yearSheetName}'!D${ROWS.GIFTS},'${yearSheetName}'!F${ROWS.GIFTS},'${yearSheetName}'!H${ROWS.GIFTS},'${yearSheetName}'!J${ROWS.GIFTS},'${yearSheetName}'!L${ROWS.GIFTS},'${yearSheetName}'!N${ROWS.GIFTS},'${yearSheetName}'!P${ROWS.GIFTS},'${yearSheetName}'!R${ROWS.GIFTS},'${yearSheetName}'!T${ROWS.GIFTS},'${yearSheetName}'!V${ROWS.GIFTS},'${yearSheetName}'!X${ROWS.GIFTS}) + SUM('${yearSheetName}'!B${ROWS.COMPENSATING},'${yearSheetName}'!D${ROWS.COMPENSATING},'${yearSheetName}'!F${ROWS.COMPENSATING},'${yearSheetName}'!H${ROWS.COMPENSATING},'${yearSheetName}'!J${ROWS.COMPENSATING},'${yearSheetName}'!L${ROWS.COMPENSATING},'${yearSheetName}'!N${ROWS.COMPENSATING},'${yearSheetName}'!P${ROWS.COMPENSATING},'${yearSheetName}'!R${ROWS.COMPENSATING},'${yearSheetName}'!T${ROWS.COMPENSATING},'${yearSheetName}'!V${ROWS.COMPENSATING},'${yearSheetName}'!X${ROWS.COMPENSATING})`;
    this.planSheet.getRange(targetRow, 9).setFormula(giftsFormula);
    for (var col = 2; col <= 9; col++) this.planSheet.getRange(targetRow, col).setNumberFormat('#,##0.00');
    var note = this.createPlanNote(year);
    this.planSheet.getRange(targetRow, 1).setNote(note);
    showMessage('Финансовый план обновлён', `📊 Данные за ${year} год успешно добавлены в план`, 'info');
    addLogEntry(`Финансовый план обновлён за ${year} год`, 'INFO', 'FinancialPlan');
  }

  getColumnLetter(colNum) { var result = ''; while (colNum > 0) { colNum--; result = String.fromCharCode(65 + (colNum % 26)) + result; colNum = Math.floor(colNum / 26); } return result; }

  findOrCreateYearRow(year) {
    var lastRow = this.planSheet.getLastRow();
    var years = this.planSheet.getRange(2, 1, Math.max(1, lastRow - 1), 1).getValues();
    for (var i = 0; i < years.length; i++) if (years[i][0] == year) return i + 2;
    var newRow = this.planSheet.getLastRow() + 1;
    this.planSheet.getRange(newRow, 1).setValue(year);
    return newRow;
  }

  createPlanNote(year) {
    return '📊 Финансовый план за ' + year + ' год\n═══════════════════════════\n\nB - Сумма доходов (строка TOTAL, чётные колонки)\nC - Сумма расходов (строка TOTAL, нечётные колонки)\nD - Прибыль = B - C - F - G - I - E\nE - Инвестиции (PROJECTS + COMPENSATING2)\nF - Страховка\nG - Пенсия\nH - Резерв\nI - Подарки (GIFTS + COMPENSATING)\n\nОбновлено: ' + new Date().toLocaleString('ru-RU');
  }

  createYearForecast() {
    security.checkAccess('createYearForecast');
    if (!this.planSheet) { addLogEntry('Лист "Финансовый план" не найден', 'WARNING', 'FinancialPlan'); return; }
    var lastRow = this.planSheet.getLastRow();
    if (lastRow < 3) { showMessage('Ошибка', 'Недостаточно данных для прогноза', 'error'); return; }
    var data = this.planSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    var lastYearsCount = Math.min(3, data.length);
    var lastYears = [];
    for (var i = data.length - lastYearsCount; i < data.length; i++) lastYears.push(data[i]);
    var incomes = [], expenses = [], profits = [], investments = [];
    for (var j = 0; j < lastYears.length; j++) { if (lastYears[j][1] > 0) incomes.push(lastYears[j][1]); if (lastYears[j][2] > 0) expenses.push(lastYears[j][2]); if (lastYears[j][3] > 0) profits.push(lastYears[j][3]); if (lastYears[j][4] > 0) investments.push(lastYears[j][4]); }
    var avgIncome = this.calculateAverage(incomes), avgExpenses = this.calculateAverage(expenses), avgProfit = this.calculateAverage(profits), avgInvestments = this.calculateAverage(investments);
    var nextYear = this.currentYear + 1;
    var forecastRow = this.planSheet.getLastRow() + 1;
    this.planSheet.getRange(forecastRow, 1).setValue(nextYear);
    this.planSheet.getRange(forecastRow, 2).setValue(avgIncome * 1.08);
    this.planSheet.getRange(forecastRow, 3).setValue(avgExpenses * 1.05);
    this.planSheet.getRange(forecastRow, 4).setValue(avgProfit * 1.1);
    this.planSheet.getRange(forecastRow, 5).setValue(avgInvestments * 1.1);
    this.planSheet.getRange(forecastRow, 6).setValue(avgProfit * 0.1);
    this.planSheet.getRange(forecastRow, 7).setValue(avgProfit * 0.1);
    this.planSheet.getRange(forecastRow, 8).setValue(20000);
    this.planSheet.getRange(forecastRow, 9).setValue(avgProfit * 0.05);
    for (var col = 2; col <= 9; col++) this.planSheet.getRange(forecastRow, col).setNumberFormat('#,##0.00');
    var note = '📈 Прогноз на основе данных ' + lastYears.length + ' лет\nСредний доход: ' + formatMoney(avgIncome) + '\nСредний расход: ' + formatMoney(avgExpenses) + '\nСредняя прибыль: ' + formatMoney(avgProfit);
    this.planSheet.getRange(forecastRow, 1).setNote(note);
    showMessage('Прогноз создан', `📊 Прогноз на ${nextYear} год добавлен в финансовый план`, 'info');
    addLogEntry(`Создан прогноз на ${nextYear} год`, 'INFO', 'FinancialPlan');
  }

  calculateAverage(numbers) { var valid = numbers.filter(function(n) { return !isNaN(n) && n > 0; }); if (valid.length === 0) return 0; return valid.reduce(function(a,b) { return a+b; }, 0) / valid.length; }

  showFinancialHistory() {
    security.checkAccess('showFinancialHistory');
    if (!this.planSheet) { addLogEntry('Лист "Финансовый план" не найден', 'WARNING', 'FinancialPlan'); return; }
    var lastRow = this.planSheet.getLastRow();
    if (lastRow < 2) { showMessage('Информация', 'Нет данных в финансовом плане', 'warning'); return; }
    var data = this.planSheet.getRange(2, 1, lastRow - 1, 10).getValues();
    data.sort(function(a,b) { return b[0] - a[0]; });
    var report = '📊 ИСТОРИЯ ФИНАНСОВОГО ПЛАНА\n═══════════════════════════\n\n';
    for (var i = 0; i < data.length; i++) { var row = data[i]; if (row[0]) { report += '📅 ' + row[0] + ' год:\n  Доход: ' + formatMoney(row[1]) + '\n  Расход: ' + formatMoney(row[2]) + '\n  Прибыль: ' + formatMoney(row[3]) + '\n  Инвестиции: ' + formatMoney(row[4]) + '\n  Страховка: ' + formatMoney(row[5]) + '\n  Пенсия: ' + formatMoney(row[6]) + '\n  Резерв: ' + formatMoney(row[7]) + '\n  Подарки: ' + formatMoney(row[8]) + '\n\n'; } }
    showMessage('История плана', report);
    addLogEntry('Показана история финансового плана', 'INFO', 'FinancialPlan');
  }

  initFinancialPlanSheet() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEETS.FINANCIAL_PLAN);
      if (sheet) { var ui = SpreadsheetApp.getUi(); if (ui.alert('Инициализация', 'Лист "Финансовый план" уже существует. Пересоздать? Все данные будут потеряны.', ui.ButtonSet.YES_NO) !== ui.Button.YES) return; ss.deleteSheet(sheet); }
      sheet = ss.insertSheet(SHEETS.FINANCIAL_PLAN);
      sheet.setTabColor('#34a853');
      var headers = [['Год', 'Доход (₽)', 'Расход (₽)', 'Прибыль (₽)', 'Инвестиции (₽)', 'Страховка (₽)', 'Пенсия (₽)', 'Резерв (₽)', 'Подарки (₽)', 'Примечания']];
      sheet.getRange(1, 1, 1, 10).setValues(headers);
      sheet.getRange(1, 1, 1, 10).setFontWeight('bold').setBackground('#f0f0f0').setHorizontalAlignment('center');
      sheet.setColumnWidth(1, 80); for (var i=2;i<=9;i++) sheet.setColumnWidth(i,120); sheet.setColumnWidth(10,300);
      sheet.setFrozenRows(1);
      showMessage('Готово', '✅ Лист "Финансовый план" создан', 'info');
      addLogEntry('Лист "Финансовый план" создан', 'INFO', 'FinancialPlan');
    } catch(e) { logError(e, 'FinancialPlanManager.initFinancialPlanSheet'); showMessage('Ошибка', e.toString(), 'error'); }
  }
}

function updateFinancialPlan() { new FinancialPlanManager().updateFinancialPlan(); }
function createYearForecast() { new FinancialPlanManager().createYearForecast(); }
function showFinancialHistory() { new FinancialPlanManager().showFinancialHistory(); }
function initFinancialPlanSheet() { new FinancialPlanManager().initFinancialPlanSheet(); }