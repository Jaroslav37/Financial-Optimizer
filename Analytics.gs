/**
 * АНАЛИТИКА - лист "Аналитика"
 */

class AnalyticsModule {
  constructor() { 
    this.security = security; 
    this.bondCalculator = new BondCalculator(); 
  }

  analyzeExpenses() {
    this.security.checkAccess('analyzeExpenses');
    try {
      var yearData = collectYearData();
      if (!yearData) return null;
      var startDate = new Date(CONFIG.CURRENT_YEAR, 0, 1), endDate = new Date(CONFIG.CURRENT_YEAR, 11, 31);
      var expenseDetails = null, incomeDetails = null;
      try { expenseDetails = collectExpenseDetails(startDate, endDate); } catch(e) { logError(e, 'AnalyticsModule.analyzeExpenses - expenseDetails'); }
      try { incomeDetails = collectIncomeDetails(startDate, endDate); } catch(e) { logError(e, 'AnalyticsModule.analyzeExpenses - incomeDetails'); }
      var analysis = { totalIncome: yearData.income, totalExpenses: yearData.expenses, profit: yearData.profit, categories: yearData.categories, monthly: {}, specialRows: yearData.specialRows, recommendations: [], savingsRate: yearData.income > 0 ? (yearData.profit / yearData.income * 100) : 0, expenseDetails: expenseDetails, incomeDetails: incomeDetails, investmentStats: { totalInvestmentIncome: 0, totalInvestmentExpenses: 0, dividendIncome: 0, couponIncome: 0, capitalGains: 0, brokerCommissions: 0, taxPaid: 0 } };
      if (expenseDetails && expenseDetails.investmentExpenses) { for (var i = 0; i < expenseDetails.investmentExpenses.length; i++) { var expense = expenseDetails.investmentExpenses[i]; if (expense.description && expense.description.toLowerCase().indexOf('комиссия') !== -1) analysis.investmentStats.brokerCommissions += expense.amount; analysis.investmentStats.totalInvestmentExpenses += expense.amount; } }
      if (incomeDetails && incomeDetails.investmentIncome) { for (var j = 0; j < incomeDetails.investmentIncome.length; j++) { var income = incomeDetails.investmentIncome[j]; if (income.category === 'Дивиденды') analysis.investmentStats.dividendIncome += income.amount; else if (income.category === 'Купоны по облигациям') analysis.investmentStats.couponIncome += income.amount; else if (income.category === 'Продажа активов') analysis.investmentStats.capitalGains += income.amount; analysis.investmentStats.totalInvestmentIncome += income.amount; analysis.investmentStats.taxPaid += income.tax || 0; } }
      for (var month = 0; month < 12; month++) { if (yearData.byMonth[month]) { var monthName = MONTHS[month], monthData = yearData.byMonth[month]; analysis.monthly[monthName] = { income: monthData.income, expenses: monthData.expenses, balance: Number((monthData.income - monthData.expenses).toFixed(2)), investmentIncome: monthData.investmentIncome || 0, investmentExpenses: monthData.investmentExpenses || 0, items: monthData.items }; } }
      var categoryPercent = {};
      var categoryKeys = Object.keys(analysis.categories);
      for (var k = 0; k < categoryKeys.length; k++) { var cat = categoryKeys[k], amount = analysis.categories[cat]; categoryPercent[cat] = { amount: amount, percent: analysis.totalExpenses > 0 ? Number((amount / analysis.totalExpenses * 100).toFixed(1)) : 0 }; }
      analysis.categoryPercent = categoryPercent;
      if (analysis.totalExpenses > analysis.totalIncome * 0.9) analysis.recommendations.push({ priority: 'high', type: 'expense', message: 'Расходы (' + formatMoney(analysis.totalExpenses) + ') составляют более 90% от доходов (' + formatMoney(analysis.totalIncome) + '). Рекомендуется сократить траты.' });
      var catKeys = Object.keys(categoryPercent);
      for (var m = 0; m < catKeys.length; m++) { var catItem = catKeys[m], dataItem = categoryPercent[catItem]; if (dataItem.percent > 30 && catItem !== 'Жильё' && catItem !== 'Прочее' && catItem !== 'Инвестиции') analysis.recommendations.push({ priority: 'medium', type: 'category', category: catItem, message: 'Категория "' + catItem + '" составляет ' + dataItem.percent + '% всех расходов (' + formatMoney(dataItem.amount) + '). Рассмотрите возможность оптимизации.' }); }
      var balance = 0;
      if (yearData.specialRows && yearData.specialRows['balance']) { var balanceValues = Object.values(yearData.specialRows['balance']); for (var n = 0; n < balanceValues.length; n++) balance += balanceValues[n]; }
      if (balance < 0) analysis.recommendations.push({ priority: 'high', type: 'balance', message: 'Отрицательный остаток: ' + formatMoney(balance) + '. Требуется срочная корректировка бюджета.' });
      if (analysis.savingsRate < 10) analysis.recommendations.push({ priority: 'medium', type: 'savings', message: 'Норма сбережений составляет ' + analysis.savingsRate.toFixed(1) + '%. Рекомендуемый минимум - 10-15%.' });
      else if (analysis.savingsRate > 20) analysis.recommendations.push({ priority: 'low', type: 'savings', message: 'Отличная норма сбережений - ' + analysis.savingsRate.toFixed(1) + '%. Рекомендуется направить сбережения в инвестиции.' });
      if (analysis.investmentStats.brokerCommissions > analysis.investmentStats.totalInvestmentIncome * 0.05 && analysis.investmentStats.totalInvestmentIncome > 0) analysis.recommendations.push({ priority: 'medium', type: 'investment', message: 'Комиссии брокера (' + formatMoney(analysis.investmentStats.brokerCommissions) + ') составляют более 5% от инвестиционного дохода. Рассмотрите смену тарифа.' });
      return analysis;
    } catch(e) { logError(e, 'AnalyticsModule.analyzeExpenses'); this.security.secureLog('Error in analyzeExpenses', { error: e.toString() }); return null; }
  }

  runFullFinancialAnalysis() {
    this.security.checkAccess('runFullFinancialAnalysis');
    try {
      var expenseAnalysis = this.analyzeExpenses();
      var fullAnalysis = { timestamp: new Date(), expenses: expenseAnalysis, combinedRecommendations: [], summary: {} };
      if (expenseAnalysis) fullAnalysis.summary = { totalIncome: expenseAnalysis.totalIncome, totalExpenses: expenseAnalysis.totalExpenses, profit: expenseAnalysis.profit, savingsRate: expenseAnalysis.savingsRate, investmentIncome: expenseAnalysis.investmentStats.totalInvestmentIncome, investmentExpenses: expenseAnalysis.investmentStats.totalInvestmentExpenses };
      if (expenseAnalysis && expenseAnalysis.recommendations) { for (var i = 0; i < expenseAnalysis.recommendations.length; i++) fullAnalysis.combinedRecommendations.push(expenseAnalysis.recommendations[i]); }
      this.saveFullAnalytics(fullAnalysis);
      return fullAnalysis;
    } catch(e) { logError(e, 'AnalyticsModule.runFullFinancialAnalysis'); this.security.secureLog('Error in runFullFinancialAnalysis', { error: e.toString() }); return null; }
  }

  saveFullAnalytics(analysis) {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sheet = ss.getSheetByName(SHEETS.ANALYTICS);
      if (!sheet) sheet = ss.insertSheet(SHEETS.ANALYTICS); else sheet.clear();
      var row = 1;
      sheet.getRange(row, 1).setValue('📊 ФИНАНСОВЫЙ АНАЛИЗ').setFontWeight('bold').setFontSize(14); sheet.getRange(row, 1, 1, 7).merge(); row += 2;
      sheet.getRange(row, 1).setValue('Дата анализа:'); sheet.getRange(row, 2).setValue(analysis.timestamp); row += 2;
      if (analysis.summary) {
        sheet.getRange(row, 1).setValue('📋 СВОДКА').setFontWeight('bold').setFontSize(12); row++;
        sheet.getRange(row, 1).setValue('Показатель'); sheet.getRange(row, 2).setValue('Значение'); sheet.getRange(row, 3).setValue('Показатель'); sheet.getRange(row, 4).setValue('Значение'); row++;
        sheet.getRange(row, 1).setValue('Доходы:'); sheet.getRange(row, 2).setValue(formatMoney(analysis.summary.totalIncome || 0)); sheet.getRange(row, 3).setValue('Инвестиционный доход:'); sheet.getRange(row, 4).setValue(formatMoney(analysis.summary.investmentIncome || 0)); row++;
        sheet.getRange(row, 1).setValue('Расходы:'); sheet.getRange(row, 2).setValue(formatMoney(analysis.summary.totalExpenses || 0)); sheet.getRange(row, 3).setValue('Инвестиционные расходы:'); sheet.getRange(row, 4).setValue(formatMoney(analysis.summary.investmentExpenses || 0)); row++;
        sheet.getRange(row, 1).setValue('Прибыль:'); sheet.getRange(row, 2).setValue(formatMoney(analysis.summary.profit || 0)); row++;
        sheet.getRange(row, 1).setValue('Норма сбережений:'); sheet.getRange(row, 2).setValue((analysis.summary.savingsRate || 0).toFixed(1) + '%'); row += 2;
      }
      if (analysis.expenses) {
        sheet.getRange(row, 1).setValue('💰 ОСНОВНЫЕ ПОКАЗАТЕЛИ (БЮДЖЕТ):').setFontWeight('bold'); row++;
        sheet.getRange(row, 1).setValue('Доходы:'); sheet.getRange(row, 2).setValue(analysis.expenses.totalIncome).setNumberFormat('#,##0.00 ₽'); row++;
        sheet.getRange(row, 1).setValue('Расходы:'); sheet.getRange(row, 2).setValue(analysis.expenses.totalExpenses).setNumberFormat('#,##0.00 ₽'); row++;
        sheet.getRange(row, 1).setValue('Прибыль:'); sheet.getRange(row, 2).setValue(analysis.expenses.profit).setNumberFormat('#,##0.00 ₽'); row++;
        sheet.getRange(row, 1).setValue('Норма сбережений:'); sheet.getRange(row, 2).setValue(analysis.expenses.savingsRate.toFixed(1) + '%'); row += 2;
        if (analysis.expenses.categoryPercent) {
          sheet.getRange(row, 1).setValue('📊 РАСХОДЫ ПО КАТЕГОРИЯМ:').setFontWeight('bold'); row++;
          sheet.getRange(row, 1).setValue('Категория').setFontWeight('bold'); sheet.getRange(row, 2).setValue('Сумма').setFontWeight('bold'); sheet.getRange(row, 3).setValue('Доля').setFontWeight('bold'); row++;
          var sorted = Object.entries(analysis.expenses.categoryPercent).sort(function(a, b) { return b[1].amount - a[1].amount; });
          for (var c = 0; c < sorted.length; c++) { var cat = sorted[c][0], dataCat = sorted[c][1]; sheet.getRange(row, 1).setValue(cat); sheet.getRange(row, 2).setValue(dataCat.amount).setNumberFormat('#,##0.00 ₽'); sheet.getRange(row, 3).setValue(dataCat.percent + '%'); row++; }
          row += 2;
        }
        if (analysis.expenses.investmentStats) {
          sheet.getRange(row, 1).setValue('📈 ИНВЕСТИЦИОННАЯ СТАТИСТИКА:').setFontWeight('bold'); row++;
          sheet.getRange(row, 1).setValue('Дивиденды:'); sheet.getRange(row, 2).setValue(formatMoney(analysis.expenses.investmentStats.dividendIncome)); row++;
          sheet.getRange(row, 1).setValue('Купоны:'); sheet.getRange(row, 2).setValue(formatMoney(analysis.expenses.investmentStats.couponIncome)); row++;
          sheet.getRange(row, 1).setValue('Комиссии брокера:'); sheet.getRange(row, 2).setValue(formatMoney(analysis.expenses.investmentStats.brokerCommissions)); row++;
          sheet.getRange(row, 1).setValue('Уплачено налогов:'); sheet.getRange(row, 2).setValue(formatMoney(analysis.expenses.investmentStats.taxPaid)); row += 2;
        }
      }
      sheet.getRange(row, 1).setValue('💡 РЕКОМЕНДАЦИИ:').setFontWeight('bold').setFontSize(12); row++;
      if (analysis.combinedRecommendations.length > 0) {
        var priorityOrder = { high: 0, medium: 1, low: 2 };
        var sortedRecs = analysis.combinedRecommendations.slice(); sortedRecs.sort(function(a,b){return priorityOrder[a.priority] - priorityOrder[b.priority];});
        for (var r = 0; r < sortedRecs.length; r++) { var rec = sortedRecs[r]; var priorityIcon = rec.priority === 'high' ? '🔴' : (rec.priority === 'medium' ? '🟡' : '🟢'); sheet.getRange(row, 1).setValue(priorityIcon + ' [' + rec.priority + '] ' + rec.message); row++; }
      } else { sheet.getRange(row, 1).setValue('✅ Рекомендаций нет. Отличная работа!'); row++; }
      sheet.getRange(1, 1, row, 7).setWrap(true); sheet.setColumnWidth(1,250); sheet.setColumnWidth(2,150); sheet.setColumnWidth(3,200); sheet.setColumnWidth(4,150); sheet.setColumnWidth(5,100); sheet.setColumnWidth(6,100); sheet.setColumnWidth(7,100);
      sheet.setFrozenRows(1);
      this.security.secureLog('Analytics saved', { rows: row });
      addLogEntry('Финансовый анализ сохранён', 'INFO', 'AnalyticsModule');
    } catch(e) { logError(e, 'AnalyticsModule.saveFullAnalytics'); this.security.secureLog('Error saving analytics', { error: e.toString() }); throw e; }
  }

  saveAnalytics(analysis) {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID), sheet = ss.getSheetByName(SHEETS.ANALYTICS);
      if (!sheet) sheet = ss.insertSheet(SHEETS.ANALYTICS); else sheet.clear();
      var row = 1;
      sheet.getRange(row, 1).setValue('📊 ФИНАНСОВЫЙ АНАЛИЗ').setFontWeight('bold').setFontSize(14); row += 2;
      sheet.getRange(row, 1).setValue('ОСНОВНЫЕ ПОКАЗАТЕЛИ:').setFontWeight('bold'); row++;
      sheet.getRange(row, 1).setValue('Доходы:'); sheet.getRange(row, 2).setValue(analysis.totalIncome).setNumberFormat('#,##0.00'); row++;
      sheet.getRange(row, 1).setValue('Расходы:'); sheet.getRange(row, 2).setValue(analysis.totalExpenses).setNumberFormat('#,##0.00'); row++;
      sheet.getRange(row, 1).setValue('Прибыль:'); sheet.getRange(row, 2).setValue(analysis.profit).setNumberFormat('#,##0.00'); row += 2;
      sheet.getRange(row, 1).setValue('РАСХОДЫ ПО КАТЕГОРИЯМ:').setFontWeight('bold'); row++;
      sheet.getRange(row, 1).setValue('Категория').setFontWeight('bold'); sheet.getRange(row, 2).setValue('Сумма').setFontWeight('bold'); sheet.getRange(row, 3).setValue('Доля').setFontWeight('bold'); row++;
      var catKeysOld = Object.keys(analysis.categoryPercent);
      for (var o = 0; o < catKeysOld.length; o++) { var catOld = catKeysOld[o], dataOld = analysis.categoryPercent[catOld]; sheet.getRange(row, 1).setValue(catOld); sheet.getRange(row, 2).setValue(dataOld.amount).setNumberFormat('#,##0.00'); sheet.getRange(row, 3).setValue(dataOld.percent + '%'); row++; }
      row += 2;
      sheet.getRange(row, 1).setValue('РЕКОМЕНДАЦИИ:').setFontWeight('bold'); row++;
      if (analysis.recommendations.length > 0) { for (var rOld = 0; rOld < analysis.recommendations.length; rOld++) { var recOld = analysis.recommendations[rOld]; sheet.getRange(row, 1).setValue('[' + recOld.priority + '] ' + recOld.message); row++; } } else { sheet.getRange(row, 1).setValue('Рекомендаций нет'); row++; }
      sheet.getRange(1, 1, row, 3).setWrap(true); sheet.setColumnWidth(2,200);
    } catch(e) { logError(e, 'AnalyticsModule.saveAnalytics'); this.security.secureLog('Error saving analytics', { error: e.toString() }); }
  }

  generateFinancialGoalsReport() {
    this.security.checkAccess('generateFinancialGoalsReport');
    try {
      var expenseAnalysis = this.analyzeExpenses();
      var report = '🎯 ОТЧЁТ ПО ФИНАНСОВЫМ ЦЕЛЯМ\n═══════════════════════════\n\n';
      if (expenseAnalysis) {
        report += '💰 Ежемесячные сбережения: ' + formatMoney(expenseAnalysis.profit / 12) + '\n📈 Годовые сбережения: ' + formatMoney(expenseAnalysis.profit) + '\n💎 Норма сбережений: ' + expenseAnalysis.savingsRate.toFixed(1) + '%\n\n';
        if (expenseAnalysis.investmentStats.totalInvestmentIncome > 0) report += '📊 ИНВЕСТИЦИОННЫЙ ДОХОД ЗА ГОД:\n   Дивиденды: ' + formatMoney(expenseAnalysis.investmentStats.dividendIncome) + '\n   Купоны: ' + formatMoney(expenseAnalysis.investmentStats.couponIncome) + '\n   Налог: ' + formatMoney(expenseAnalysis.investmentStats.taxPaid) + '\n   Чистый доход: ' + formatMoney(expenseAnalysis.investmentStats.dividendIncome + expenseAnalysis.investmentStats.couponIncome - expenseAnalysis.investmentStats.taxPaid) + '\n\n';
      }
      // Финансовые цели без инвестиционного портфеля – только сбережения
      var targetAmounts = [1000000, 5000000, 10000000, 20000000, 50000000];
      var monthlySavings = expenseAnalysis ? expenseAnalysis.profit / 12 : 0;
      var currentPortfolio = 0; // Нет автоматического портфеля
      var annualReturn = 0.08; // Допущение
      report += '📋 ДОСТИЖЕНИЕ ФИНАНСОВЫХ ЦЕЛЕЙ (при норме сбережений):\n';
      for (var t = 0; t < targetAmounts.length; t++) { var target = targetAmounts[t]; var needed = Math.max(0, target - currentPortfolio); if (needed === 0) report += '✅ ' + formatMoney(target) + ': Достигнуто!\n'; else if (monthlySavings > 0) { var monthsNeeded = 0, current = currentPortfolio, monthlyReturn = annualReturn / 12; while (current < target && monthsNeeded < 600) { current = current * (1 + monthlyReturn) + monthlySavings; monthsNeeded++; } if (monthsNeeded < 600) { var yearsNeeded = Math.floor(monthsNeeded / 12); var monthsRemain = monthsNeeded % 12; report += '🎯 ' + formatMoney(target) + ': ' + yearsNeeded + ' лет ' + monthsRemain + ' мес.\n'; } else report += '🎯 ' + formatMoney(target) + ': Более 50 лет\n'; } }
      return report;
    } catch(e) { logError(e, 'AnalyticsModule.generateFinancialGoalsReport'); this.security.secureLog('Error generating goals report', { error: e.toString() }); return 'Ошибка генерации отчёта'; }
  }
}

function runFullAnalysis() { try { var analytics = new AnalyticsModule(); var fullAnalysis = analytics.runFullFinancialAnalysis(); if (!fullAnalysis) { showMessage('Ошибка', 'Не удалось выполнить анализ', 'error'); return; } var report = '📊 РЕЗУЛЬТАТЫ КОМПЛЕКСНОГО АНАЛИЗА\n═══════════════════════════════════\n\n'; if (fullAnalysis.expenses) { report += '💰 БЮДЖЕТ:\n   Доходы: ' + formatMoney(fullAnalysis.expenses.totalIncome) + '\n   Расходы: ' + formatMoney(fullAnalysis.expenses.totalExpenses) + '\n   Прибыль: ' + formatMoney(fullAnalysis.expenses.profit) + '\n   Норма сбережений: ' + fullAnalysis.expenses.savingsRate.toFixed(1) + '%\n'; if (fullAnalysis.expenses.investmentStats.totalInvestmentIncome > 0) report += '   Инвестиционный доход: ' + formatMoney(fullAnalysis.expenses.investmentStats.totalInvestmentIncome) + '\n'; report += '\n'; } if (fullAnalysis.combinedRecommendations.length > 0) { report += '💡 РЕКОМЕНДАЦИИ:\n'; for (var i = 0; i < fullAnalysis.combinedRecommendations.length; i++) { var rec = fullAnalysis.combinedRecommendations[i]; var icon = rec.priority === 'high' ? '🔴' : (rec.priority === 'medium' ? '🟡' : '🟢'); report += icon + ' ' + rec.message + '\n'; } } else report += '✅ Рекомендаций нет. Отличная работа!\n'; showMessage('Комплексный анализ', report); } catch(e) { logError(e, 'runFullAnalysis'); showMessage('Ошибка', e.toString(), 'error'); } }
function runExpenseAnalysis() { try { var analytics = new AnalyticsModule(); var analysis = analytics.analyzeExpenses(); if (!analysis) { showMessage('Ошибка', 'Не удалось выполнить анализ', 'error'); return; } analytics.saveAnalytics(analysis); var report = '📊 РЕЗУЛЬТАТЫ АНАЛИЗА РАСХОДОВ\n═══════════════════════════\n\n💰 Доходы: ' + formatMoney(analysis.totalIncome) + '\n💸 Расходы: ' + formatMoney(analysis.totalExpenses) + '\n💎 Прибыль: ' + formatMoney(analysis.profit) + '\n💎 Норма сбережений: ' + analysis.savingsRate.toFixed(1) + '%\n\n'; if (analysis.investmentStats.totalInvestmentIncome > 0) report += '📈 ИНВЕСТИЦИОННЫЙ ДОХОД:\n   Дивиденды: ' + formatMoney(analysis.investmentStats.dividendIncome) + '\n   Купоны: ' + formatMoney(analysis.investmentStats.couponIncome) + '\n   Налог: ' + formatMoney(analysis.investmentStats.taxPaid) + '\n   Чистый: ' + formatMoney(analysis.investmentStats.totalInvestmentIncome - analysis.investmentStats.taxPaid) + '\n\n'; report += '📊 РАСХОДЫ ПО КАТЕГОРИЯМ:\n'; var sorted = Object.entries(analysis.categoryPercent).sort(function(a,b){return b[1].amount - a[1].amount;}); for (var s=0; s<sorted.length; s++) { var cat=sorted[s][0], dataCat=sorted[s][1]; var icon = cat === 'Инвестиции' ? '📈' : '💰'; report += icon + ' ' + cat + ': ' + formatMoney(dataCat.amount) + ' (' + dataCat.percent + '%)\n'; } report += '\n💡 РЕКОМЕНДАЦИИ:\n'; if (analysis.recommendations.length > 0) { for (var r=0; r<analysis.recommendations.length; r++) { var rec=analysis.recommendations[r]; var recIcon=rec.priority==='high'?'🔴':(rec.priority==='medium'?'🟡':'🟢'); report+=recIcon+' '+rec.message+'\n'; } } else report+='✅ Рекомендаций нет. Отличная работа!\n'; showMessage('Анализ расходов', report); } catch(e) { logError(e, 'runExpenseAnalysis'); showMessage('Ошибка', e.toString(), 'error'); } }
function generateFinancialGoalsReport() { try { var analytics = new AnalyticsModule(); var report = analytics.generateFinancialGoalsReport(); showMessage('Финансовые цели', report); } catch(e) { logError(e, 'generateFinancialGoalsReport'); showMessage('Ошибка', e.toString(), 'error'); } }