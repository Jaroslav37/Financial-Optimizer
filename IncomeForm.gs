/**
 * ВЕБ-ФОРМА ДЛЯ ВВОДА ДОХОДОВ
 * Исправлено: XSS-защита через шаблон
 */

class IncomeForm {
  constructor() {
    this.sheet = getSheetSafe(SHEETS.CURRENT_YEAR);
    this.detailsModule = new IncomeDetailsModule();
  }

  show() {
    var ui = SpreadsheetApp.getUi();
    var html = this.createHtmlForm();
    ui.showModalDialog(html, '💰 Ввод доходов');
  }

  createHtmlForm() {
    var today = formatDate(new Date());
    var categories = Object.keys(INCOME_CATEGORIES).join(',');
    var template = HtmlService.createTemplate(`
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <meta charset="utf-8">
          <title>Ввод доходов</title>
          <style>
            *{box-sizing:border-box} body{font-family:'Segoe UI',Arial,sans-serif;padding:20px;background:#f5f5f5;margin:0}
            .container{max-width:650px;margin:0 auto;background:white;padding:25px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,0.1)}
            h2{color:#1a73e8;margin-top:0;margin-bottom:20px;font-size:22px}
            .date-group{margin-bottom:20px} label{display:block;margin-bottom:5px;font-weight:500;color:#333;font-size:14px}
            input,select,textarea{width:100%;padding:10px;border:1px solid #ddd;border-radius:6px;font-size:14px}
            button{padding:10px 16px;border:none;border-radius:6px;font-size:14px;font-weight:500;cursor:pointer}
            .btn-primary{background:#1a73e8;color:white} .btn-success{background:#34a853;color:white}
            .btn-danger{background:#ea4335;color:white} .btn-secondary{background:#f1f3f4;color:#333}
            .row{display:flex;gap:10px;margin-top:15px}
            .income-item{background:#f8f9fa;padding:12px;margin-bottom:8px;border-radius:6px;border-left:4px solid #34a853;display:flex;justify-content:space-between;align-items:center}
            .income-item.investment{border-left-color:#1a73e8;background:#e8f0fe}
            .income-item .info{flex-grow:1} .income-item .category{font-size:12px;color:#666;margin-top:3px}
            .income-item .amount{font-weight:bold;color:#34a853;font-size:16px;margin:0 10px}
            .total{text-align:right;font-size:18px;font-weight:bold;margin-top:15px;padding-top:15px;border-top:2px solid #e0e0e0}
            .list-container{max-height:300px;overflow-y:auto;margin-bottom:20px}
          </style>
        </head>
        <body>
          <div class="container">
            <h2>💰 Ввод доходов</h2>
            <div class="date-group"><label>Дата:</label><input type="text" id="incomeDate" value="<?= today ?>" placeholder="ДД.ММ.ГГГГ" onchange="loadIncomes()"></div>
            <div class="list-container" id="incomesList"></div>
            <div class="total">Итого: <span id="totalAmount">0.00 ₽</span></div>
            <div class="form-group"><label>Описание:</label><input type="text" id="desc" placeholder="Например: Зарплата, Дивиденды Сбербанк"></div>
            <div class="form-group"><label>Сумма (₽):</label><input type="number" id="amount" step="0.01" placeholder="0.00"></div>
            <div class="form-group"><label>Категория:</label><select id="category">
              <? for (var cat of categories.split(',')) { ?>
                <option value="<?= cat ?>"><?= cat ?></option>
              <? } ?>
            </select></div>
            <div class="form-group"><label>Подкатегория:</label><input type="text" id="subcategory" placeholder="например: премия, акции"></div>
            <div class="form-group"><label>Тикер/ISIN:</label><input type="text" id="ticker" placeholder="SBER, SU26239RMFS5"></div>
            <div class="form-group"><label>Налог (₽):</label><input type="number" id="tax" step="0.01" value="0"></div>
            <div class="form-group"><label>Источник:</label><select id="source">
              <option value="manual">Ручной ввод</option><option value="sber">Сбер Инвестиции</option>
              <option value="tinkoff">Тинькофф</option><option value="broker">Другой брокер</option>
            </select></div>
            <div class="row"><button class="btn-success" onclick="addIncome()">➕ Добавить</button>
            <button class="btn-primary" onclick="saveAll()">💾 Сохранить все</button>
            <button class="btn-secondary" onclick="google.script.host.close()">❌ Отмена</button></div>
          </div>
          <script>
            let incomes = [];
            function loadIncomes() {
              var dateStr = document.getElementById('incomeDate').value.trim();
              if (!dateStr) return;
              google.script.run.withSuccessHandler(function(data) {
                incomes = data.map(function(i) { return { id: i.id, description: i.description, amount: i.amount, category: i.category, subcategory: i.subcategory || '', ticker: i.ticker || '', tax: i.tax || 0, source: i.source || 'manual' }; });
                renderIncomes();
              }).getIncomesForDate(dateStr);
            }
            function addIncome() {
              var desc = document.getElementById('desc').value.trim();
              var amount = parseFloat(document.getElementById('amount').value);
              var category = document.getElementById('category').value;
              var subcategory = document.getElementById('subcategory').value.trim();
              var ticker = document.getElementById('ticker').value.trim();
              var tax = parseFloat(document.getElementById('tax').value) || 0;
              var source = document.getElementById('source').value;
              if (!desc) { alert('Введите описание'); return; }
              if (isNaN(amount) || amount <= 0) { alert('Введите корректную сумму'); return; }
              incomes.push({ id: null, description: desc, amount: amount, category: category, subcategory: subcategory, ticker: ticker, tax: tax, source: source });
              document.getElementById('desc').value = ''; document.getElementById('amount').value = '';
              document.getElementById('subcategory').value = ''; document.getElementById('ticker').value = ''; document.getElementById('tax').value = '0';
              renderIncomes();
            }
            function deleteIncome(idx) {
              if (confirm('Удалить этот доход?')) { incomes.splice(idx, 1); renderIncomes(); }
            }
            function renderIncomes() {
              var container = document.getElementById('incomesList');
              var total = 0;
              if (incomes.length === 0) container.innerHTML = '<p style="color:#999; text-align:center; padding:20px;">Нет добавленных доходов</p>';
              else {
                container.innerHTML = '';
                incomes.forEach(function(inc, idx) {
                  total += inc.amount;
                  var isInvestment = inc.category === 'Дивиденды' || inc.category === 'Купоны по облигациям' || inc.category === 'Продажа активов';
                  var div = document.createElement('div');
                  div.className = 'income-item' + (isInvestment ? ' investment' : '');
                  div.innerHTML = '<div class="info"><strong>' + escapeHtml(inc.description) + '</strong><div class="category">' + escapeHtml(inc.category) + (inc.subcategory ? ' → ' + escapeHtml(inc.subcategory) : '') + '</div>' +
                    (inc.ticker ? '<div class="category">Тикер: ' + escapeHtml(inc.ticker) + '</div>' : '') +
                    (inc.tax > 0 ? '<div class="category">Налог: ' + inc.tax.toFixed(2) + ' ₽</div>' : '') + '</div>' +
                    '<div class="amount">' + inc.amount.toFixed(2) + ' ₽</div>' +
                    '<div class="actions"><button class="btn-danger" onclick="deleteIncome(' + idx + ')">✕</button></div>';
                  container.appendChild(div);
                });
              }
              document.getElementById('totalAmount').textContent = total.toFixed(2) + ' ₽';
            }
            function saveAll() {
              if (incomes.length === 0 && !confirm('Нет добавленных доходов. Продолжить?')) return;
              var dateStr = document.getElementById('incomeDate').value.trim();
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
              }).saveIncomes(dateStr, incomes);
            }
            function escapeHtml(str) { if (!str) return ''; return str.replace(/[&<>]/g, function(m) { if (m === '&') return '&amp;'; if (m === '<') return '&lt;'; if (m === '>') return '&gt;'; return m; }); }
            loadIncomes();
          </script>
        </body>
      </html>
    `);
    template.today = today;
    template.categories = categories;
    return template.evaluate().setWidth(650).setHeight(700);
  }

  getIncomesForDate(dateStr) {
    var date = safeParseDate(dateStr);
    if (!date) return [];
    return new IncomeDetailsModule().getDayIncomes(date);
  }

  saveIncomes(dateStr, incomes) {
    try {
      var date = safeParseDate(dateStr);
      if (!date) return { success: false, error: 'Неверная дата' };
      var module = new IncomeDetailsModule();
      var existing = module.getDayIncomes(date);
      for (var i = 0; i < existing.length; i++) { if (existing[i].id) module.deleteIncomeRecord(existing[i].id); }
      var added = 0, totalAmount = 0, totalTax = 0;
      for (var j = 0; j < incomes.length; j++) {
        var inc = incomes[j];
        var result = module.addIncomeRecord(date, inc.category, inc.amount, inc.description, inc.source || 'manual', inc.subcategory || '', inc.ticker || '', inc.tax || 0);
        if (result.success) { added++; totalAmount += inc.amount; totalTax += inc.tax || 0; }
      }
      return { success: true, message: 'Добавлено ' + added + ' доходов на сумму ' + formatMoney(totalAmount) + (totalTax > 0 ? ' (налог: ' + formatMoney(totalTax) + ')' : ''), added: added, totalAmount: totalAmount };
    } catch(e) { return { success: false, error: e.toString() }; }
  }
}

function showIncomeForm() { new IncomeForm().show(); }
function getIncomesForDate(dateStr) { return new IncomeForm().getIncomesForDate(dateStr); }
function saveIncomes(dateStr, incomes) { return new IncomeForm().saveIncomes(dateStr, incomes); }