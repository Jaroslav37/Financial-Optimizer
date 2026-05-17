// =============================================
// СБОР ДАННЫХ ДЛЯ АНАЛИТИКИ
// =============================================

function collectYearData() {
  try {
    var sheet = getSheetSafe(SHEETS.CURRENT_YEAR);
    if (!sheet) return null;
    var lastCol = sheet.getLastColumn();
    var totalIncome = 0, totalExpenses = 0;
    var categories = {};
    var byMonth = {};
    var specialRows = { balance: {}, projects: {}, gifts: {}, compensating: {}, compensating2: {} };
    
    for (var col = 2; col <= lastCol; col += 2) {
      var monthIndex = Math.floor((col - 2) / 2);
      var monthName = MONTHS[monthIndex];
      var income = safeParseNumber(sheet.getRange(ROWS.TOTAL, col).getValue());
      var expense = safeParseNumber(sheet.getRange(ROWS.TOTAL, col + 1).getValue());
      totalIncome += income;
      totalExpenses += expense;
      
      byMonth[monthIndex] = { income: income, expenses: expense, items: [] };
      
      // Собираем категории один раз (из декабрьских итогов)
      if (monthIndex === 11) {
        var expDetails = collectExpenseDetails(new Date(CONFIG.CURRENT_YEAR,0,1), new Date(CONFIG.CURRENT_YEAR,11,31));
        if (expDetails && expDetails.categories) categories = expDetails.categories;
      }
      
      specialRows.projects[monthName] = safeParseNumber(sheet.getRange(ROWS.PROJECTS, col).getValue());
      specialRows.gifts[monthName] = safeParseNumber(sheet.getRange(ROWS.GIFTS, col).getValue());
      specialRows.compensating[monthName] = safeParseNumber(sheet.getRange(ROWS.COMPENSATING, col).getValue());
      specialRows.compensating2[monthName] = safeParseNumber(sheet.getRange(ROWS.COMPENSATING2, col).getValue());
      specialRows.balance[monthName] = safeParseNumber(sheet.getRange(ROWS.BALANCE, col + 1).getValue());
    }
    
    return {
      income: totalIncome,
      expenses: totalExpenses,
      profit: totalIncome - totalExpenses,
      categories: categories,
      byMonth: byMonth,
      specialRows: specialRows
    };
  } catch(e) {
    logError(e, 'collectYearData');
    return null;
  }
}

function collectExpenseDetails(startDate, endDate) {
  var expenseModule = new ExpenseDetailsModule();
  var stats = expenseModule.getCategoryStats(startDate, endDate);
  var investmentExpenses = expenseModule.getExpensesByCategory('Инвестиции', startDate, endDate);
  return { categories: stats.stats, total: stats.total, investmentExpenses: investmentExpenses };
}

function collectIncomeDetails(startDate, endDate) {
  var incomeModule = new IncomeDetailsModule();
  var stats = incomeModule.getCategoryStats(startDate, endDate);
  var investmentCategories = ['Дивиденды', 'Купоны по облигациям', 'Продажа активов'];
  var investmentIncome = [];
  if (stats.stats) {
    for (var cat in stats.stats) {
      if (investmentCategories.indexOf(cat) !== -1) {
        investmentIncome = investmentIncome.concat(stats.stats[cat].items);
      }
    }
  }
  return { categories: stats.stats, total: stats.total, investmentIncome: investmentIncome };
}

function collectAllFinancialData() {
  return {
    yearData: collectYearData(),
    expenseDetails: collectExpenseDetails(new Date(CONFIG.CURRENT_YEAR,0,1), new Date(CONFIG.CURRENT_YEAR,11,31)),
    incomeDetails: collectIncomeDetails(new Date(CONFIG.CURRENT_YEAR,0,1), new Date(CONFIG.CURRENT_YEAR,11,31))
  };
}