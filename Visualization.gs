/**
 * ВИЗУАЛИЗАЦИЯ - Дашборды и отчёты
 * Версия 3.1 - Без инвестиционного портфеля (чистка зависимостей)
 */

class VisualizationModule {
  constructor() {
    this.security = security;
  }

  collectDashboardData() {
    try {
      const yearData = collectYearData();
      if (!yearData) return null;

      const analytics = new AnalyticsModule();
      const expenseAnalysis = analytics.analyzeExpenses();
      
      const totalIncome = yearData.income || 0;
      const totalExpenses = yearData.expenses || 0;
      const profit = yearData.profit || 0;
      const savingsRate = totalIncome > 0 ? (profit / totalIncome * 100) : 0;

      const monthlyData = [];
      for (let month = 0; month < 12; month++) {
        const monthData = yearData.byMonth[month] || { income: 0, expenses: 0 };
        monthlyData.push({
          month: MONTHS[month],
          income: monthData.income || 0,
          expenses: monthData.expenses || 0,
          balance: (monthData.income || 0) - (monthData.expenses || 0)
        });
      }

      const categoryData = [];
      if (yearData.categories) {
        const catKeys = Object.keys(yearData.categories);
        for (let i = 0; i < catKeys.length; i++) {
          const cat = catKeys[i];
          const amount = yearData.categories[cat];
          if (amount > 0) {
            categoryData.push({
              category: cat,
              amount: Math.round(amount * 100) / 100,
              percent: totalExpenses > 0 ? Math.round((amount / totalExpenses * 100) * 10) / 10 : 0
            });
          }
        }
      }
      categoryData.sort((a, b) => b.amount - a.amount);
      const topCategories = categoryData.slice(0, 5);

      const budgetData = [];
      if (typeof BUDGET_CONFIG !== 'undefined' && BUDGET_CONFIG.DEFAULT_LIMITS) {
        const limits = BUDGET_CONFIG.DEFAULT_LIMITS;
        const catKeys = Object.keys(limits);
        const currentMonth = new Date().getMonth() + 1;
        
        for (let i = 0; i < catKeys.length; i++) {
          const cat = catKeys[i];
          const limit = limits[cat];
          let spent = 0;
          if (yearData.byMonth && yearData.byMonth[currentMonth - 1]) {
            spent = yearData.byMonth[currentMonth - 1].expensesByCategory?.[cat] || 0;
          }
          if (spent === 0 && yearData.categories && yearData.categories[cat]) {
            spent = yearData.categories[cat] / 12;
          }
          if (limit > 0) {
            budgetData.push({
              category: cat,
              limit: limit,
              spent: Math.round(spent * 100) / 100,
              percentUsed: Math.min(100, Math.round((spent / limit * 100) * 10) / 10)
            });
          }
        }
      }
      budgetData.sort((a, b) => b.percentUsed - a.percentUsed);

      const recommendations = [];
      if (expenseAnalysis && expenseAnalysis.recommendations) {
        for (let i = 0; i < expenseAnalysis.recommendations.length; i++) {
          recommendations.push(expenseAnalysis.recommendations[i]);
        }
      }

      return {
        totalIncome: Math.round(totalIncome * 100) / 100,
        totalExpenses: Math.round(totalExpenses * 100) / 100,
        profit: Math.round(profit * 100) / 100,
        savingsRate: Math.round(savingsRate * 10) / 10,
        monthlyData: monthlyData,
        categoryData: categoryData,
        topCategories: topCategories,
        budgetData: budgetData,
        recommendations: recommendations,
        year: CONFIG.CURRENT_YEAR,
        currentMonth: MONTHS[new Date().getMonth()]
      };
    } catch (e) {
      logError(e, 'VisualizationModule.collectDashboardData');
      return null;
    }
  }

  createDashboard() {
    this.security.checkAccess('createDashboard');
    try {
      const dashboardData = this.collectDashboardData();
      if (!dashboardData) {
        showMessage('Ошибка', 'Не удалось собрать данные для дашборда', 'error');
        return;
      }
      const html = HtmlService.createTemplateFromFile('DashboardHTML');
      html.dashboardData = JSON.stringify(dashboardData);
      const htmlOutput = html.evaluate().setWidth(1100).setHeight(800).setTitle('📊 Финансовый дашборд');
      SpreadsheetApp.getUi().showModalDialog(htmlOutput, '📊 Финансовый дашборд');
    } catch (e) {
      logError(e, 'VisualizationModule.createDashboard');
      showMessage('Ошибка', 'Не удалось создать дашборд: ' + e.toString(), 'error');
    }
  }

  createInlineDashboard() {
    this.security.checkAccess('createInlineDashboard');
    try {
      const dashboardData = this.collectDashboardData();
      if (!dashboardData) {
        showMessage('Ошибка', 'Не удалось собрать данные для дашборда', 'error');
        return;
      }
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      let sheet = ss.getSheetByName('Дашборд');
      if (sheet) sheet.clear();
      else sheet = ss.insertSheet('Дашборд');
      sheet.setTabColor('#4285f4');
      let row = 1;
      sheet.getRange(row, 1, 1, 5).merge();
      sheet.getRange(row, 1).setValue(`📊 ФИНАНСОВЫЙ ДАШБОРД ${dashboardData.year}`)
        .setFontWeight('bold').setFontSize(16).setBackground('#e8f0fe');
      row += 2;
      sheet.getRange(row, 1).setValue('КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ').setFontWeight('bold').setFontSize(12);
      row++;
      const metrics = [
        ['💵 Доходы', formatMoney(dashboardData.totalIncome)],
        ['💸 Расходы', formatMoney(dashboardData.totalExpenses)],
        ['💰 Прибыль', formatMoney(dashboardData.profit)],
        ['📈 Норма сбережений', dashboardData.savingsRate.toFixed(1) + '%']
      ];
      for (let i = 0; i < metrics.length; i++) {
        sheet.getRange(row, 1).setValue(metrics[i][0]).setFontWeight('bold');
        sheet.getRange(row, 2).setValue(metrics[i][1]);
        if (i === 3) sheet.getRange(row, 2).setFontColor(dashboardData.savingsRate >= 10 ? '#34a853' : '#ea4335');
        row++;
      }
      row += 1;
      sheet.getRange(row, 1).setValue('ТОП-5 КАТЕГОРИЙ РАСХОДОВ').setFontWeight('bold').setFontSize(12);
      row++;
      sheet.getRange(row, 1).setValue('Категория').setFontWeight('bold');
      sheet.getRange(row, 2).setValue('Сумма').setFontWeight('bold');
      sheet.getRange(row, 3).setValue('Доля').setFontWeight('bold');
      row++;
      for (let i = 0; i < dashboardData.topCategories.length; i++) {
        const cat = dashboardData.topCategories[i];
        sheet.getRange(row, 1).setValue(cat.category);
        sheet.getRange(row, 2).setValue(cat.amount).setNumberFormat('#,##0.00 ₽');
        sheet.getRange(row, 3).setValue(cat.percent + '%');
        row++;
      }
      row += 1;
      if (dashboardData.budgetData.length > 0) {
        sheet.getRange(row, 1).setValue('ИСПОЛНЕНИЕ БЮДЖЕТА').setFontWeight('bold').setFontSize(12);
        row++;
        sheet.getRange(row, 1).setValue('Категория').setFontWeight('bold');
        sheet.getRange(row, 2).setValue('Лимит').setFontWeight('bold');
        sheet.getRange(row, 3).setValue('Потрачено').setFontWeight('bold');
        sheet.getRange(row, 4).setValue('%').setFontWeight('bold');
        row++;
        for (let i = 0; i < Math.min(10, dashboardData.budgetData.length); i++) {
          const b = dashboardData.budgetData[i];
          sheet.getRange(row, 1).setValue(b.category);
          sheet.getRange(row, 2).setValue(b.limit).setNumberFormat('#,##0.00 ₽');
          sheet.getRange(row, 3).setValue(b.spent).setNumberFormat('#,##0.00 ₽');
          sheet.getRange(row, 4).setValue(b.percentUsed + '%');
          if (b.percentUsed >= 90) sheet.getRange(row, 4).setFontColor('#ea4335');
          else if (b.percentUsed >= 70) sheet.getRange(row, 4).setFontColor('#fbbc04');
          row++;
        }
        row += 1;
      }
      if (dashboardData.recommendations.length > 0) {
        sheet.getRange(row, 1).setValue('💡 РЕКОМЕНДАЦИИ').setFontWeight('bold').setFontSize(12);
        row++;
        for (let i = 0; i < dashboardData.recommendations.length; i++) {
          const rec = dashboardData.recommendations[i];
          const priorityIcon = rec.priority === 'high' ? '🔴' : (rec.priority === 'medium' ? '🟡' : '🟢');
          sheet.getRange(row, 1).setValue(priorityIcon + ' ' + rec.message);
          row++;
        }
        row += 1;
      }
      sheet.setColumnWidth(1, 250);
      sheet.setColumnWidth(2, 150);
      sheet.setColumnWidth(3, 100);
      sheet.setColumnWidth(4, 100);
      sheet.setColumnWidth(5, 100);
      sheet.setFrozenRows(0);
      showMessage('Дашборд создан', 'Дашборд добавлен на лист "Дашборд"', 'info');
      addLogEntry('Создан дашборд в листе', 'INFO', 'VisualizationModule');
    } catch (e) {
      logError(e, 'VisualizationModule.createInlineDashboard');
      showMessage('Ошибка', 'Не удалось создать дашборд', 'error');
    }
  }
}

function createDashboard() {
  new VisualizationModule().createDashboard();
}

function createInlineDashboard() {
  new VisualizationModule().createInlineDashboard();
}