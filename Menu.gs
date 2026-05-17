/**
 * МЕНЮ – настройка интерфейса пользователя
 * Версия 7.2 – очищено от зависимостей удалённых модулей
 */

function onOpen() {
  try {
    var ui = SpreadsheetApp.getUi();
    var menu = ui.createMenu('📊 Финансы');

    // =============================================
    // 0. УНИФИЦИРОВАННЫЙ ВВОД
    // =============================================
    menu.addItem('📝 День за минуту (доходы + расходы)', 'showUnifiedDayForm');
    menu.addSeparator();

    // =============================================
    // 1. ДОХОДЫ
    // =============================================
    var incomeMenu = ui.createMenu('💰 Только доходы');
    incomeMenu.addItem('📝 Ввести доходы', 'showIncomeForm');
    incomeMenu.addSeparator();
    incomeMenu.addItem('📊 Статистика доходов', 'showIncomeStats');
    incomeMenu.addItem('📈 Статистика инвестиционных доходов', 'showInvestmentIncomeStats');
    incomeMenu.addItem('🆕 Инициализировать детализацию', 'initIncomeDetails');
    menu.addSubMenu(incomeMenu);

    // =============================================
    // 2. РАСХОДЫ
    // =============================================
    var expensesMenu = ui.createMenu('💸 Только расходы');
    expensesMenu.addItem('📝 Ввести расходы', 'showExpenseForm');
    expensesMenu.addSeparator();
    expensesMenu.addItem('📊 Статистика расходов', 'showExpenseStats');
    expensesMenu.addItem('📈 Статистика инвестиционных расходов', 'showInvestmentExpensesStats');
    menu.addSubMenu(expensesMenu);

    // =============================================
    // 3. ДОЛГИ
    // =============================================
    var debtsMenu = ui.createMenu('💰 Долги');
    debtsMenu.addItem('📊 Рассчитать все долги', 'calculateAllDebts');
    debtsMenu.addItem('➕ Добавить долг (пошагово)', 'addNewDebtInteractive');
    debtsMenu.addItem('⚡ Добавить долг (быстро)', 'addNewDebtQuick');
    debtsMenu.addItem('✅ Отметить как погашенный', 'markDebtAsPaid');
    debtsMenu.addItem('📈 Показать отчёт', 'showDebtReport');
    menu.addSubMenu(debtsMenu);

    // =============================================
    // 4. ФИНАНСОВЫЙ ПЛАН
    // =============================================
    var planMenu = ui.createMenu('📈 Финансовый план');
    planMenu.addItem('📊 Обновить из текущего года', 'updateFinancialPlan');
    planMenu.addItem('🔮 Создать прогноз на следующий год', 'createYearForecast');
    planMenu.addItem('📜 Показать историю', 'showFinancialHistory');
    menu.addSubMenu(planMenu);

    // =============================================
    // 5. БЮДЖЕТ
    // =============================================
    var budgetMenu = ui.createMenu('💰 Бюджет');
    budgetMenu.addItem('📊 Показать отчёт по бюджету', 'showBudgetReport');
    budgetMenu.addItem('💡 Показать рекомендации', 'showBudgetRecommendations');
    budgetMenu.addItem('⚙️ Установить лимит категории', 'setCategoryLimit');
    budgetMenu.addItem('🔍 Проверить статус бюджета', 'checkBudgetStatus');
    budgetMenu.addItem('🔄 Обновить бюджет (прямой расчёт)', 'refreshBudget');
    budgetMenu.addItem('📤 Экспорт бюджета', 'exportBudget');
    budgetMenu.addItem('🔧 Исправить формулы', 'fixBudgetFormulas');
    budgetMenu.addItem('🔍 Проверить данные детализации', 'validateDetailsData');
    menu.addSubMenu(budgetMenu);

    // =============================================
    // 6. АНАЛИТИКА
    // =============================================
    var analyticsMenu = ui.createMenu('📊 Аналитика');
    analyticsMenu.addItem('🔍 Запустить полный анализ', 'runFullAnalysis');
    analyticsMenu.addItem('📈 Показать дашборд', 'createDashboard');
    analyticsMenu.addItem('📉 Дашборд в таблице', 'createInlineDashboard');
    menu.addSubMenu(analyticsMenu);

    // =============================================
    // 7. ТЕХНИЧЕСКИЕ
    // =============================================
    var techMenu = ui.createMenu('🛠️ Технические');
    techMenu.addItem('🆕 Создать лист бюджета', 'initBudgetSheet');
    techMenu.addItem('🆕 Инициализировать лист долгов', 'initDebtsSheet');
    techMenu.addItem('🆕 Инициализировать детализацию', 'initExpenseDetails');
    techMenu.addItem('📋 Создать лист для следующего года', 'createNewYearSheet');
    techMenu.addItem('💰 Показать остаток для переноса', 'showBalanceToTransfer');
    techMenu.addItem('📋 Инициализация листов (выборочная)', 'showInitSheetsDialog');
    menu.addSubMenu(techMenu);

    // =============================================
    // 8. ТЕСТИРОВАНИЕ
    // =============================================
    var testingMenu = ui.createMenu('🧪 Тестирование');
    testingMenu.addItem('✅ Запустить все тесты', 'runAllTests');
    testingMenu.addItem('📝 Создать тестовые данные', 'createTestData');
    testingMenu.addItem('🔍 Тест сбора данных', 'testDataCollection');
    testingMenu.addItem('🔍 Проверить целостность данных', 'validateDataIntegrity');
    menu.addSubMenu(testingMenu);

    // =============================================
    // 9. ТРИГГЕРЫ
    // =============================================
    var triggersMenu = ui.createMenu('⚙️ Триггеры');
    triggersMenu.addItem('🎛️ Управление триггерами', 'manageTriggers');
    triggersMenu.addItem('▶️ Запустить все триггеры', 'setupAllTriggers');
    triggersMenu.addItem('⏸️ Остановить все триггеры', 'stopAllTriggers');
    triggersMenu.addItem('📋 Проверить триггеры', 'checkAllTriggers');
    triggersMenu.addItem('📝 Логи триггеров', 'logAllTriggers');
    menu.addSubMenu(triggersMenu);

    // =============================================
    // 10. ОБЛИГАЦИИ (исправлено)
    // =============================================
    var bondMenu = ui.createMenu('📈 Облигации');
    bondMenu.addItem('🆕 Создать лист "Облигации"', 'initBondSheet');
    bondMenu.addItem('➕ Ручной ввод облигации', 'manualBondInput');
    bondMenu.addItem('🔄 Сгенерировать IMPORTXML формулы', 'generateImportFormulas');
    bondMenu.addItem('📊 Рассчитать доходность', 'calculateBondYield');
    bondMenu.addItem('🔍 Проверить ISIN', 'testIsin');
    bondMenu.addItem('ℹ️ Справка по доходностям', 'showBondHelp');
    menu.addSubMenu(bondMenu);

    // =============================================
    // 11. ФЬЮЧЕРСЫ
    // =============================================
    var futuresMenu = ui.createMenu('📈 Фьючерсы');
    futuresMenu.addItem('➕ Добавить сделку', 'showAddFuturesTrade');
    futuresMenu.addItem('🎯 Точки безубыточности', 'showFuturesBreakeven');
    menu.addSubMenu(futuresMenu);

    // =============================================
    // 12. АКЦИИ
    // =============================================
    var stocksMenu = ui.createMenu('📈 Акции');
    stocksMenu.addItem('➕ Добавить сделку', 'showAddStockTrade');
    stocksMenu.addItem('🔄 Добавить сплит/консолидацию', 'showAddStockSplit');
    stocksMenu.addItem('🎯 Точки безубыточности', 'showStocksBreakeven');
    menu.addSubMenu(stocksMenu);

    // =============================================
    // 13. СПЕЦИАЛЬНЫЕ ОПЕРАЦИИ
    // =============================================
    var specialMenu = ui.createMenu('🎁 Спец. операции');
    specialMenu.addItem('🎁 Добавить подарок', 'showGiftForm');
    specialMenu.addItem('📊 Отчёт по подаркам', 'showGiftsReport');
    specialMenu.addItem('📈 Инвестиционное пополнение', 'showInvestmentContributionForm');
    specialMenu.addItem('📊 Отчёт по пополнениям', 'showInvestmentsReport');
    specialMenu.addItem('🆕 Инициализировать листы', 'initSpecialSheets');
    menu.addSubMenu(specialMenu);

    // =============================================
    // 14. НОВЫЙ ГОД
    // =============================================
    var newYearMenu = ui.createMenu('🎄 Новый год');
    newYearMenu.addItem('📋 Создать лист на следующий год', 'createNewYearSheet');
    newYearMenu.addItem('💰 Показать остаток для переноса', 'showBalanceToTransfer');
    newYearMenu.addItem('📊 Анализ за год', 'showYearAnalysis');
    menu.addSubMenu(newYearMenu);

    // =============================================
    // 15. ЛОГИ
    // =============================================
    var logsMenu = ui.createMenu('📋 Логи');
    logsMenu.addItem('📊 Статистика логов', 'showLogsStats');
    logsMenu.addItem('📤 Экспорт логов', 'exportLogs');
    logsMenu.addItem('🗑️ Очистить старые логи', 'cleanupLogs');
    logsMenu.addItem('🆕 Создать лист логов', 'createLogsSheet');
    menu.addSubMenu(logsMenu);

    // =============================================
    // 16. БЕЗОПАСНОСТЬ
    // =============================================
    var securityMenu = ui.createMenu('🔒 Безопасность');
    securityMenu.addItem('🔐 Информация о безопасности', 'showSecurityInfo');
    securityMenu.addItem('➕ Добавить разрешённый email', 'addAllowedEmail');
    securityMenu.addItem('🗑️ Удалить разрешённый email', 'removeAllowedEmail');
    menu.addSubMenu(securityMenu);

    // =============================================
    // 17. ПОДКАТЕГОРИИ
    // =============================================
    var subcatMenu = ui.createMenu('📋 Подкатегории');
    subcatMenu.addItem('➕ Добавить подкатегорию дохода', 'showAddIncomeSubcategoryDialog');
    subcatMenu.addItem('➕ Добавить подкатегорию расхода', 'showAddExpenseSubcategoryDialog');
    menu.addSubMenu(subcatMenu);

    menu.addToUi();
    
    addLogEntry('Меню успешно создано (v7.2)', 'INFO', 'Menu');
  } catch (e) {
    logError(e, 'onOpen');
    try {
      SpreadsheetApp.getUi().alert('❌ Ошибка меню', 'Не удалось создать меню. Ошибка: ' + e.toString(), SpreadsheetApp.getUi().ButtonSet.OK);
    } catch (f) {}
  }
}

// =============================================
// ОБЁРТКИ ДЛЯ ВЫЗОВА ФУНКЦИЙ
// =============================================

function runFullAnalysis() {
  try {
    var analytics = new AnalyticsModule();
    var fullAnalysis = analytics.runFullFinancialAnalysis();
    if (!fullAnalysis) { showMessage('Ошибка', 'Не удалось выполнить анализ', 'error'); return; }
    var report = '📊 РЕЗУЛЬТАТЫ КОМПЛЕКСНОГО АНАЛИЗА\n═══════════════════════════════════\n\n';
    if (fullAnalysis.expenses) {
      report += '💰 БЮДЖЕТ:\n   Доходы: ' + formatMoney(fullAnalysis.expenses.totalIncome) + '\n   Расходы: ' + formatMoney(fullAnalysis.expenses.totalExpenses) + '\n   Прибыль: ' + formatMoney(fullAnalysis.expenses.profit) + '\n   Норма сбережений: ' + fullAnalysis.expenses.savingsRate.toFixed(1) + '%\n\n';
    }
    if (fullAnalysis.combinedRecommendations && fullAnalysis.combinedRecommendations.length > 0) {
      report += '💡 РЕКОМЕНДАЦИИ:\n';
      for (var i = 0; i < fullAnalysis.combinedRecommendations.length; i++) {
        var rec = fullAnalysis.combinedRecommendations[i];
        var icon = rec.priority === 'high' ? '🔴' : (rec.priority === 'medium' ? '🟡' : '🟢');
        report += icon + ' ' + rec.message + '\n';
      }
    } else {
      report += '✅ Рекомендаций нет. Отличная работа!\n';
    }
    showMessage('Комплексный анализ', report);
  } catch (e) { logError(e, 'runFullAnalysis'); showMessage('Ошибка', e.toString(), 'error'); }
}

function createDashboard() {
  try { new VisualizationModule().createDashboard(); } catch (e) { logError(e, 'createDashboard'); showMessage('Ошибка', e.toString(), 'error'); }
}

function createInlineDashboard() {
  try { new VisualizationModule().createInlineDashboard(); } catch (e) { logError(e, 'createInlineDashboard'); showMessage('Ошибка', e.toString(), 'error'); }
}

function showUnifiedDayForm() {
  try { new UnifiedDayForm().show(); } catch (e) { logError(e, 'showUnifiedDayForm'); showMessage('Ошибка', e.toString(), 'error'); }
}

function showUnifiedDayFormForDate(date) {
  new UnifiedDayForm().show(date);
}