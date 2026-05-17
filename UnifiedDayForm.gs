/**
 * УНИВЕРСАЛЬНАЯ ФОРМА ДНЯ – расширена: автоопределение категории, месячные итоги, повтор вчера
 * Версия 3.0
 */

class UnifiedDayForm {
  constructor() {
    this.sheet = getSheetSafe(SHEETS.CURRENT_YEAR);
    this.today = new Date();
    this.currentYear = CONFIG.CURRENT_YEAR;
  }

  show(date) {
    const ui = SpreadsheetApp.getUi();
    const html = HtmlService.createTemplateFromFile('UnifiedDayFormHTML');
    html.todayDate = formatDate(date || new Date());
    html.categories = JSON.stringify(this._getAllCategories());
    html.quickTemplates = JSON.stringify(this._getQuickTemplates());
    
    const htmlOutput = html.evaluate()
      .setWidth(750)
      .setHeight(950)
      .setTitle('📝 День за минуту');
    
    ui.showModalDialog(htmlOutput, '📝 День за минуту');
  }

  _getAllCategories() {
    const subMgr = new SubcategoriesManager();

    const baseIncomeCats = {
      'Зарплата': ['основная', 'премия', 'аванс'],
      'Фриланс': ['разработка', 'дизайн', 'консалтинг'],
      'Дивиденды': ['акции', 'дивиденды'],
      'Купоны по облигациям': ['офз', 'корпоративные'],
      'Продажа активов': ['акции', 'облигации'],
      'Проценты': ['депозит', 'накопительный счёт'],
      'Подарки': ['деньги'],
      'Возврат долга': ['возврат'],
      'Кэшбэк': ['банк', 'бонусы'],
      'Другое': []
    };

    const baseExpenseCats = {
      'Продукты': ['супермаркет', 'рынок', 'доставка'],
      'Транспорт': ['бензин', 'такси', 'метро'],
      'Здоровье': ['лекарства', 'врачи', 'аптека'],
      'Жильё': ['квартплата', 'ремонт', 'аренда'],
      'Развлечения': ['кино', 'кафе', 'подписки'],
      'Одежда': ['одежда', 'обувь'],
      'Связь': ['интернет', 'телефон'],
      'Образование': ['курсы', 'книги'],
      'Инвестиции': ['комиссия', 'налоги', 'депозитарий'],
      'Долги': ['кредит', 'проценты'],
      'Прочее': []
    };

    return {
      income: subMgr.getIncomeSubcategories(baseIncomeCats),
      expense: subMgr.getExpenseSubcategories(baseExpenseCats),
      gift: {
        'Подарок': ['деньги', 'вещь'],
        'Поздравление': ['цветы', 'открытка'],
        'Благотворительность': ['помощь', 'фонд'],
        'Другое': []
      },
      investment: {
        'Накопление': ['Сбер', 'Тинькофф', 'ИИС', 'Брокерский счёт'],
        'Пенсия': ['ИИС', 'Брокерский счёт'],
        'Крупная покупка': ['Сбер', 'Тинькофф'],
        'Дивидендный доход': ['Сбер', 'Тинькофф'],
        'Другое': []
      }
    };
  }

  _getQuickTemplates() {
    return [
      { type: 'expense', icon: '🛒', label: 'Продукты', category: 'Продукты', subcategory: 'супермаркет', source: 'manual' },
      { type: 'expense', icon: '⛽', label: 'Бензин', category: 'Транспорт', subcategory: 'бензин', source: 'manual' },
      { type: 'expense', icon: '🚕', label: 'Такси', category: 'Транспорт', subcategory: 'такси', source: 'manual' },
      { type: 'expense', icon: '☕', label: 'Кафе', category: 'Развлечения', subcategory: 'кафе', source: 'manual' },
      { type: 'income', icon: '💼', label: 'Зарплата', category: 'Зарплата', subcategory: 'основная', source: 'manual' },
      { type: 'income', icon: '💵', label: 'Дивиденды', category: 'Дивиденды', subcategory: 'акции', source: 'sber' },
      { type: 'income', icon: '🎫', label: 'Купоны', category: 'Купоны по облигациям', subcategory: 'офз', source: 'sber' },
      { type: 'gift', icon: '🎁', label: 'Подарок', category: 'Подарок', subcategory: 'деньги', source: 'manual' },
      { type: 'investment', icon: '📈', label: 'Инв. пополнение', category: 'Накопление', subcategory: 'Сбер', source: 'manual' }
    ];
  }

  getDayData(dateStr) {
    const date = safeParseDate(dateStr);
    if (!date) return { success: false, error: 'Неверная дата' };

    const incomeModule = new IncomeDetailsModule();
    const expenseModule = new ExpenseDetailsModule();
    const giftForm = new GiftForm();
    const investForm = new InvestmentContributionForm();

    const incomes = incomeModule.getDayIncomes(date);
    const expenses = expenseModule.getExpensesByPeriod(
      new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0),
      new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59)
    );
    const gifts = giftForm.getGiftsForDate(dateStr);
    const contributions = investForm.getContributionsForDate(dateStr);

    const items = [];

    for (const inc of incomes) {
      items.push({
        id: inc.id,
        type: 'income',
        description: inc.description,
        amount: inc.amount,
        category: inc.category,
        subcategory: inc.subcategory || '',
        source: inc.source || 'manual',
        ticker: inc.ticker || '',
        tax: inc.tax || 0
      });
    }

    for (const exp of expenses) {
      items.push({
        id: exp.id,
        type: 'expense',
        description: exp.description,
        amount: exp.amount,
        category: exp.category,
        subcategory: exp.subcategory || '',
        source: exp.source || 'manual'
      });
    }

    for (const g of gifts) {
      items.push({
        id: g.id,
        type: 'gift',
        description: g.description,
        amount: g.amount,
        category: g.category,
        event: g.event || '',
        person: g.person || '',
        source: 'manual'
      });
    }

    for (const c of contributions) {
      items.push({
        id: c.id,
        type: 'investment',
        description: c.purpose,
        amount: c.amount,
        category: c.goal,
        account: c.account || '',
        note: c.note || '',
        source: 'manual'
      });
    }

    items.sort((a, b) => {
      const order = { income: 0, expense: 1, gift: 2, investment: 3 };
      if (a.type !== b.type) return order[a.type] - order[b.type];
      return 0;
    });

    return {
      success: true,
      date: dateStr,
      items: items,
      totalIncome: incomes.reduce((s, i) => s + i.amount, 0),
      totalExpense: expenses.reduce((s, e) => s + e.amount, 0),
      totalGifts: gifts.reduce((s, g) => s + g.amount, 0),
      totalInvestments: contributions.reduce((s, c) => s + c.amount, 0)
    };
  }

  saveDayData(dateStr, items) {
    const date = safeParseDate(dateStr);
    if (!date) return { success: false, error: 'Неверная дата' };

    const incomeItems = items.filter(i => i.type === 'income');
    const expenseItems = items.filter(i => i.type === 'expense');
    const giftItems = items.filter(i => i.type === 'gift');
    const investmentItems = items.filter(i => i.type === 'investment');

    const results = {
      income: { added: 0, total: 0 },
      expense: { added: 0, total: 0 },
      gift: { added: 0, total: 0 },
      investment: { added: 0, total: 0 },
      errors: []
    };

    if (incomeItems.length > 0) {
      const incomeModule = new IncomeDetailsModule();
      const existingIncomes = incomeModule.getDayIncomes(date);
      for (const inc of existingIncomes) incomeModule.deleteIncomeRecord(inc.id);
      for (const item of incomeItems) {
        const result = incomeModule.addIncomeRecord(date, item.category, item.amount, item.description, item.source || 'manual', item.subcategory || '', item.ticker || '', item.tax || 0);
        if (result.success) { results.income.added++; results.income.total += item.amount; }
        else results.errors.push('Доход: ' + result.error);
      }
    }

    if (expenseItems.length > 0) {
      const expenseModule = new ExpenseDetailsModule();
      const existingExpenses = expenseModule.getExpensesByPeriod(
        new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0),
        new Date(date.getFullYear(), date.getMonth(), date.getDate(), 23, 59, 59)
      );
      for (const exp of existingExpenses) expenseModule.deleteExpenseRecord(exp.id);
      for (const item of expenseItems) {
        const result = expenseModule.addExpenseRecord(date, item.category, item.amount, item.description, item.source || 'manual', item.subcategory || '');
        if (result.success) { results.expense.added++; results.expense.total += item.amount; }
        else results.errors.push('Расход: ' + result.error);
      }
    }

    if (giftItems.length > 0) {
      const giftForm = new GiftForm();
      const gifts = giftItems.map(i => ({
        person: i.person || '',
        description: i.description,
        amount: i.amount,
        category: i.category,
        event: i.event || ''
      }));
      const giftResult = giftForm.saveGifts(dateStr, gifts);
      if (giftResult.success) {
        results.gift.added = giftResult.added;
        results.gift.total = giftResult.totalAmount;
      } else {
        results.errors.push('Подарки: ' + giftResult.error);
      }
    }

    if (investmentItems.length > 0) {
      const investForm = new InvestmentContributionForm();
      const contribs = investmentItems.map(i => ({
        account: i.account || '',
        amount: i.amount,
        purpose: i.description,
        goal: i.category,
        note: i.note || ''
      }));
      const investResult = investForm.saveContributions(dateStr, contribs);
      if (investResult.success) {
        results.investment.added = investResult.added;
        results.investment.total = investResult.totalAmount;
      } else {
        results.errors.push('Инвестиции: ' + investResult.error);
      }
    }

    try { new BudgetManager().checkBudgetStatus(); } catch (e) {}

    const totalAdded = results.income.added + results.expense.added + results.gift.added + results.investment.added;
    const msgParts = [];
    if (results.income.added > 0) msgParts.push(`${results.income.added} доходов (${formatMoney(results.income.total)})`);
    if (results.expense.added > 0) msgParts.push(`${results.expense.added} расходов (${formatMoney(results.expense.total)})`);
    if (results.gift.added > 0) msgParts.push(`${results.gift.added} подарков (${formatMoney(results.gift.total)})`);
    if (results.investment.added > 0) msgParts.push(`${results.investment.added} пополнений (${formatMoney(results.investment.total)})`);

    return {
      success: results.errors.length === 0,
      added: totalAdded,
      totalIncome: results.income.total,
      totalExpense: results.expense.total,
      totalGifts: results.gift.total,
      totalInvestments: results.investment.total,
      message: `✅ Добавлено: ${msgParts.join(', ') || 'нет новых записей'}`
    };
  }

  getFormData() {
    return {
      categories: this._getAllCategories(),
      quickTemplates: this._getQuickTemplates()
    };
  }

  /** Возвращает доходы, расходы и баланс за месяц, к которому принадлежит дата */
  getMonthTotals(dateStr) {
    const date = safeParseDate(dateStr);
    if (!date) return { income: 0, expenses: 0, balance: 0 };

    const incomeModule = new IncomeDetailsModule();
    const expenseModule = new ExpenseDetailsModule();

    const startOfMonth = new Date(date.getFullYear(), date.getMonth(), 1);
    const endOfMonth = new Date(date.getFullYear(), date.getMonth() + 1, 0, 23, 59, 59);

    const monthIncomes = incomeModule.getIncomesByPeriod(startOfMonth, endOfMonth);
    const monthExpenses = expenseModule.getExpensesByPeriod(startOfMonth, endOfMonth);

    const totalIncome = monthIncomes.reduce((s, i) => s + i.amount, 0);
    const totalExpenses = monthExpenses.reduce((s, e) => s + e.amount, 0);

    return {
      income: Math.round(totalIncome * 100) / 100,
      expenses: Math.round(totalExpenses * 100) / 100,
      balance: Math.round((totalIncome - totalExpenses) * 100) / 100
    };
  }
}

// Глобальные функции
function showUnifiedDayForm() { new UnifiedDayForm().show(); }
function getUnifiedDayData(dateStr) { return new UnifiedDayForm().getDayData(dateStr); }
function saveUnifiedDayData(dateStr, items) { return new UnifiedDayForm().saveDayData(dateStr, items); }
function showUnifiedDayFormForDate(date) { new UnifiedDayForm().show(date); }
function getFormData() { return new UnifiedDayForm().getFormData(); }
function getUnifiedMonthTotals(dateStr) { return new UnifiedDayForm().getMonthTotals(dateStr); }