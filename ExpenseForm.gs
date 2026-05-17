/**
 * ВЕБ-ФОРМА ДЛЯ ВВОДА РАСХОДОВ
 */

class ExpenseForm {
  constructor() {
    this.sheet = getSheetSafe(SHEETS.CURRENT_YEAR);
    this.detailsModule = new ExpenseDetailsModule();
  }

  show() {
    var ui = SpreadsheetApp.getUi();
    var html = this.createHtmlForm();
    ui.showModalDialog(html, '💸 Ввод расходов');
  }

  createHtmlForm() {
    var today = formatDate(new Date());
    var categories = Object.keys(EXPENSE_CATEGORIES).join(',');
    var html = HtmlService.createTemplate(`
      <!DOCTYPE html>
      <html><head><base target="_top"><meta charset="utf-8"><title>Ввод расходов</title>
      <style>
        *{box-sizing:border-box} body{font-family:'Segoe UI',Arial,sans-serif;padding:20px;background:#f5f5f5;margin:0}
        .container{max-width:650px;margin:0 auto;background:white;padding:25px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,0.1)}
        h2{color:#1a73e8;margin-top:0;margin-bottom:20px;font-size:22px}
        .date-group{margin-bottom:20px} label{display:block;margin-bottom:5px;font-weight:500;color:#333;font-size:14px}
        input,select,textarea{width:100%;padding:10px;border:1px solid #ddd;border-radius:6px;font-size:14px}
        button{padding:10px 16px;border:none;border-radius:6px;font-size:14px;font-weight:500;cursor:pointer;transition:background 0.2s}
        .btn-primary{background:#1a73e8;color:white} .btn-success{background:#34a853;color:white}
        .btn-danger{background:#ea4335;color:white} .btn-secondary{background:#f1f3f4;color:#333}
        .row{display:flex;gap:10px;margin-top:15px}
        .expense-item{background:#f8f9fa;padding:12px;margin-bottom:8px;border-radius:6px;border-left:4px solid #ea4335;display:flex;justify-content:space-between;align-items:center}
        .expense-item.investment{border-left-color:#34a853;background:#e6f4ea}
        .expense-item .info{flex-grow:1} .expense-item .category{font-size:12px;color:#666;margin-top:3px}
        .expense-item .amount{font-weight:bold;color:#ea4335;font-size:16px;margin:0 10px}
        .total{text-align:right;font-size:18px;font-weight:bold;margin-top:15px;padding-top:15px;border-top:2px solid #e0e0e0}
        .list-container{max-height:300px;overflow-y:auto;margin-bottom:20px}
      </style></head>
      <body><div class="container">
        <h2>💸 Ввод расходов</h2>
        <div class="date-group"><label>Дата:</label><input type="text" id="expenseDate" value="<?= today ?>" placeholder="ДД.ММ.ГГГГ" onchange="loadExpenses()"></div>
        <div class="list-container" id="expensesList"></div>
        <div class="total">Итого: <span id="totalAmount">0.00 ₽</span></div>
        <div class="form-group"><label>Описание:</label><input type="text" id="desc" placeholder="Например: Продукты, Комиссия брокера"></div>
        <div class="form-group"><label>Сумма (₽):</label><input type="number" id="amount" step="0.01" placeholder="0.00"></div>
        <div class="form-group"><label>Категория:</label><select id="category">
          <? for (var cat of categories.split(',')) { ?>
            <option value="<?= cat ?>"><?= cat ?></option>
          <? } ?>
        </select></div>
        <div class="form-group"><label>Подкатегория:</label><input type="text" id="subcategory" placeholder="например: продукты, комиссия"></div>
        <div class="form-group"><label>Источник:</label><select id="source">
          <option value="manual">Ручной ввод</option><option value="sber">Сбер Инвестиции</option>
          <option value="tinkoff">Тинькофф</option><option value="other">Другое</option>
        </select></div>
        <div class="row"><button class="btn-success" onclick="addExpense()">➕ Добавить</button>
        <button class="btn-primary" onclick="saveAll()">💾 Сохранить все</button>
        <button class="btn-secondary" onclick="google.script.host.close()">❌ Отмена</button></div>
      </div>
      <script>
        let expenses = [];
        function loadExpenses() {
          var dateStr = document.getElementById('expenseDate').value.trim();
          if (!dateStr) return;
          google.script.run.withSuccessHandler(function(data) {
            expenses = data.map(function(e) { return { id: e.id, description: e.description, amount: e.amount, category: e.category, subcategory: e.subcategory || '', source: e.source || 'manual' }; });
            renderExpenses();
          }).getExpensesForDate(dateStr);
        }
        function addExpense() {
          var desc = document.getElementById('desc').value.trim();
          var amount = parseFloat(document.getElementById('amount').value);
          var category = document.getElementById('category').value;
          var subcategory = document.getElementById('subcategory').value.trim();
          var source = document.getElementById('source').value;
          if (!desc) { alert('Введите описание'); return; }
          if (isNaN(amount) || amount <= 0) { alert('Введите корректную сумму'); return; }
          expenses.push({ id: null, description: desc, amount: amount, category: category, subcategory: subcategory, source: source });
          document.getElementById('desc').value = '';
          document.getElementById('amount').value = '';
          document.getElementById('subcategory').value = '';
          renderExpenses();
        }
        function deleteExpense(idx) {
          if (confirm('Удалить этот расход?')) { expenses.splice(idx, 1); renderExpenses(); }
        }
        function renderExpenses() {
          var container = document.getElementById('expensesList');
          var total = 0;
          if (expenses.length === 0) container.innerHTML = '<p style="color:#999; text-align:center; padding:20px;">Нет добавленных расходов</p>';
          else {
            container.innerHTML = '';
            expenses.forEach(function(exp, idx) {
              total += exp.amount;
              var isInvestment = exp.category === 'Инвестиции';
              var div = document.createElement('div');
              div.className = 'expense-item' + (isInvestment ? ' investment' : '');
              div.innerHTML = '<div class="info"><strong>' + escapeHtml(exp.description) + '</strong><div class="category">' + escapeHtml(exp.category) + (exp.subcategory ? ' → ' + escapeHtml(exp.subcategory) : '') + '</div></div>' +
                '<div class="amount">' + exp.amount.toFixed(2) + ' ₽</div>' +
                '<div class="actions"><button class="btn-danger" onclick="deleteExpense(' + idx + ')">✕</button></div>';
              container.appendChild(div);
            });
          }
          document.getElementById('totalAmount').textContent = total.toFixed(2) + ' ₽';
        }
        function saveAll() {
          if (expenses.length === 0 && !confirm('Нет добавленных расходов. Продолжить?')) return;
          var dateStr = document.getElementById('expenseDate').value.trim();
          if (!dateStr) { alert('Введите дату'); return; }
          var saveBtn = document.querySelector('.btn-primary');
          var origText = saveBtn.textContent;
          saveBtn.textContent = '⏳ Сохранение...';
          saveBtn.disabled = true;
          google.script.run.withSuccessHandler(function(result) {
            if (result.success) { alert('✅ ' + result.message); google.script.host.close(); }
            else { alert('❌ Ошибка: ' + result.error); saveBtn.textContent = origText; saveBtn.disabled = false; }
          }).withFailureHandler(function(error) {
            alert('❌ Ошибка вызова: ' + error); saveBtn.textContent = origText; saveBtn.disabled = false;
          }).saveExpenses(dateStr, expenses);
        }
        function escapeHtml(str) { if (!str) return ''; return str.replace(/[&<>]/g, function(m) { if (m === '&') return '&amp;'; if (m === '<') return '&lt;'; if (m === '>') return '&gt;'; return m; }); }
        loadExpenses();
      </script></body></html>
    `);
    html.today = today;
    html.categories = categories;
    return html.evaluate().setWidth(650).setHeight(700);
  }

  getExpensesForDate(dateStr) {
    var date = safeParseDate(dateStr);
    if (!date) return [];
    return new ExpenseDetailsModule().getExpensesByPeriod(date, date);
  }

  saveExpenses(dateStr, expenses) {
    try {
      var date = safeParseDate(dateStr);
      if (!date) return { success: false, error: 'Неверная дата' };
      var module = new ExpenseDetailsModule();
      var existing = module.getExpensesByPeriod(date, date);
      for (var i = 0; i < existing.length; i++) { if (existing[i].id) module.deleteExpenseRecord(existing[i].id); }
      var added = 0, totalAmount = 0;
      for (var j = 0; j < expenses.length; j++) {
        var exp = expenses[j];
        var result = module.addExpenseRecord(date, exp.category, exp.amount, exp.description, exp.source || 'manual', exp.subcategory || '');
        if (result.success) { added++; totalAmount += exp.amount; }
      }
      return { success: true, message: 'Добавлено ' + added + ' расходов на сумму ' + formatMoney(totalAmount), added: added, totalAmount: totalAmount };
    } catch (e) { return { success: false, error: e.toString() }; }
  }
}

function showExpenseForm() { new ExpenseForm().show(); }
function getExpensesForDate(dateStr) { return new ExpenseForm().getExpensesForDate(dateStr); }
function saveExpenses(dateStr, expenses) { return new ExpenseForm().saveExpenses(dateStr, expenses); }