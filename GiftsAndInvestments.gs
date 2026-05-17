/**
 * ПОДАРКИ И ИНВЕСТИЦИОННЫЕ ПОПОЛНЕНИЯ - листы "Подарки", "Инвестиционные пополнения"
 * Исправлено: XSS-защита через шаблоны, безопасное логирование
 */

class GiftForm {
  constructor() {
    this.sheet = getSheetSafe(SHEETS.CURRENT_YEAR);
    this.giftsSheet = this.getOrCreateGiftsSheet();
  }

  getOrCreateGiftsSheet() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Подарки');
    if (!sheet) {
      sheet = ss.insertSheet('Подарки');
      sheet.setTabColor('#f9cb9c');
      var headers = [['Дата', 'Кому/От кого', 'Описание', 'Сумма (₽)', 'Категория', 'Событие', 'ID записи', 'Месяц']];
      sheet.getRange(1, 1, 1, 8).setValues(headers);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f0f0f0');
      sheet.setColumnWidth(1, 100);
      sheet.setColumnWidth(2, 200);
      sheet.setColumnWidth(3, 250);
      sheet.setColumnWidth(4, 120);
      sheet.setColumnWidth(5, 120);
      sheet.setColumnWidth(6, 150);
      sheet.setColumnWidth(7, 150);
      sheet.setColumnWidth(8, 80);
      sheet.getRange(2, 4, 1000, 1).setNumberFormat('#,##0.00 ₽');
    }
    return sheet;
  }

  show() {
    var ui = SpreadsheetApp.getUi();
    var html = this.createHtmlForm();
    ui.showModalDialog(html, '🎁 Ввод подарков');
  }

  createHtmlForm() {
    var today = formatDate(new Date());
    var categories = ['Подарок', 'Поздравление', 'Благотворительность', 'Другое'];
    var template = HtmlService.createTemplate(`
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <meta charset="utf-8">
          <title>Ввод подарков</title>
          <style>
            * { box-sizing: border-box; }
            body {
              font-family: 'Segoe UI', Arial, sans-serif;
              padding: 20px;
              background: #f5f5f5;
              margin: 0;
            }
            .container {
              max-width: 650px;
              margin: 0 auto;
              background: white;
              padding: 25px;
              border-radius: 12px;
              box-shadow: 0 2px 10px rgba(0,0,0,0.1);
            }
            h2 {
              color: #1a73e8;
              margin-top: 0;
              margin-bottom: 20px;
              font-size: 22px;
            }
            .date-group {
              margin-bottom: 20px;
            }
            label {
              display: block;
              margin-bottom: 5px;
              font-weight: 500;
              color: #333;
              font-size: 14px;
            }
            input, select, textarea {
              width: 100%;
              padding: 10px;
              border: 1px solid #ddd;
              border-radius: 6px;
              font-size: 14px;
              box-sizing: border-box;
            }
            button {
              padding: 10px 16px;
              border: none;
              border-radius: 6px;
              font-size: 14px;
              font-weight: 500;
              cursor: pointer;
              transition: background 0.2s;
            }
            .btn-primary {
              background: #1a73e8;
              color: white;
            }
            .btn-primary:hover {
              background: #1557b0;
            }
            .btn-success {
              background: #34a853;
              color: white;
            }
            .btn-success:hover {
              background: #2d8745;
            }
            .btn-danger {
              background: #ea4335;
              color: white;
            }
            .btn-danger:hover {
              background: #d33426;
            }
            .btn-secondary {
              background: #f1f3f4;
              color: #333;
            }
            .row {
              display: flex;
              gap: 10px;
              margin-top: 15px;
            }
            .gift-item {
              background: #f8f9fa;
              padding: 12px;
              margin-bottom: 8px;
              border-radius: 6px;
              border-left: 4px solid #f9cb9c;
              display: flex;
              justify-content: space-between;
              align-items: center;
            }
            .gift-item .info {
              flex-grow: 1;
            }
            .gift-item .amount {
              font-weight: bold;
              color: #ea4335;
              font-size: 16px;
              margin: 0 10px;
            }
            .total {
              text-align: right;
              font-size: 18px;
              font-weight: bold;
              margin-top: 15px;
              padding-top: 15px;
              border-top: 2px solid #e0e0e0;
            }
            .list-container {
              max-height: 300px;
              overflow-y: auto;
              margin-bottom: 20px;
            }
          </style>
        </head>
        <body>
          <div class="container">
            <h2>🎁 Ввод подарков</h2>
            <div class="date-group">
              <label>Дата:</label>
              <input type="text" id="giftDate" value="<?= today ?>" placeholder="ДД.ММ.ГГГГ" onchange="loadGifts()">
            </div>
            <div class="list-container" id="giftsList"></div>
            <div class="total">
              Итого: <span id="totalAmount">0.00 ₽</span>
            </div>
            <div class="form-group">
              <label>Кому / От кого:</label>
              <input type="text" id="person" placeholder="Например: Мама, Иван, Коллеги">
            </div>
            <div class="form-group">
              <label>Описание:</label>
              <input type="text" id="description" placeholder="Например: День рождения, Новый год">
            </div>
            <div class="form-group">
              <label>Сумма (₽):</label>
              <input type="number" id="amount" step="0.01" placeholder="0.00">
            </div>
            <div class="form-group">
              <label>Категория:</label>
              <select id="category">
                <? for (var cat of categories) { ?>
                  <option value="<?= cat ?>"><?= cat ?></option>
                <? } ?>
              </select>
            </div>
            <div class="form-group">
              <label>Событие:</label>
              <input type="text" id="event" placeholder="Например: День рождения, 8 Марта">
            </div>
            <div class="row">
              <button class="btn-success" onclick="addGift()">➕ Добавить</button>
              <button class="btn-primary" onclick="saveAll()">💾 Сохранить все</button>
              <button class="btn-secondary" onclick="google.script.host.close()">❌ Отмена</button>
            </div>
          </div>
          <script>
            let gifts = [];
            
            function loadGifts() {
              var dateStr = document.getElementById('giftDate').value.trim();
              if (!dateStr) return;
              google.script.run
                .withSuccessHandler(function(data) {
                  gifts = data.map(function(g) {
                    return {
                      id: g.id,
                      person: g.person,
                      description: g.description,
                      amount: g.amount,
                      category: g.category,
                      event: g.event
                    };
                  });
                  renderGifts();
                })
                .getGiftsForDate(dateStr);
            }
            
            function addGift() {
              var person = document.getElementById('person').value.trim();
              var description = document.getElementById('description').value.trim();
              var amount = parseFloat(document.getElementById('amount').value);
              var category = document.getElementById('category').value;
              var event = document.getElementById('event').value.trim();
              
              if (!person) {
                alert('Введите получателя/дарителя');
                return;
              }
              if (!description) {
                alert('Введите описание');
                return;
              }
              if (isNaN(amount) || amount <= 0) {
                alert('Введите корректную сумму');
                return;
              }
              
              gifts.push({
                id: null,
                person: person,
                description: description,
                amount: amount,
                category: category,
                event: event
              });
              
              // Очистить форму
              document.getElementById('person').value = '';
              document.getElementById('description').value = '';
              document.getElementById('amount').value = '';
              document.getElementById('event').value = '';
              
              renderGifts();
            }
            
            function deleteGift(index) {
              if (confirm('Удалить этот подарок?')) {
                gifts.splice(index, 1);
                renderGifts();
              }
            }
            
            function renderGifts() {
              var container = document.getElementById('giftsList');
              var total = 0;
              
              if (gifts.length === 0) {
                container.innerHTML = '<p style="color:#999; text-align:center; padding:20px;">Нет добавленных подарков</p>';
              } else {
                container.innerHTML = '';
                gifts.forEach(function(g, idx) {
                  total += g.amount;
                  var div = document.createElement('div');
                  div.className = 'gift-item';
                  div.innerHTML = 
                    '<div class="info">' +
                      '<strong>' + escapeHtml(g.person) + '</strong>' +
                      '<div>' + escapeHtml(g.description) + '</div>' +
                      '<div class="category">' + escapeHtml(g.category) + (g.event ? ' (' + escapeHtml(g.event) + ')' : '') + '</div>' +
                    '</div>' +
                    '<div class="amount">' + g.amount.toFixed(2) + ' ₽</div>' +
                    '<div class="actions">' +
                      '<button class="btn-danger" onclick="deleteGift(' + idx + ')">✕</button>' +
                    '</div>';
                  container.appendChild(div);
                });
              }
              document.getElementById('totalAmount').textContent = total.toFixed(2) + ' ₽';
            }
            
            function saveAll() {
              if (gifts.length === 0 && !confirm('Нет добавленных подарков. Продолжить?')) return;
              
              var dateStr = document.getElementById('giftDate').value.trim();
              if (!dateStr) {
                alert('Введите дату');
                return;
              }
              
              var saveBtn = document.querySelector('.btn-primary');
              var origText = saveBtn.textContent;
              saveBtn.textContent = '⏳ Сохранение...';
              saveBtn.disabled = true;
              
              google.script.run
                .withSuccessHandler(function(result) {
                  if (result.success) {
                    alert('✅ ' + result.message);
                    google.script.host.close();
                  } else {
                    alert('❌ Ошибка: ' + result.error);
                    saveBtn.textContent = origText;
                    saveBtn.disabled = false;
                  }
                })
                .withFailureHandler(function(error) {
                  alert('❌ Ошибка вызова: ' + error);
                  saveBtn.textContent = origText;
                  saveBtn.disabled = false;
                })
                .saveGifts(dateStr, gifts);
            }
            
            function escapeHtml(str) {
              if (!str) return '';
              return str.replace(/[&<>]/g, function(m) {
                if (m === '&') return '&amp;';
                if (m === '<') return '&lt;';
                if (m === '>') return '&gt;';
                return m;
              });
            }
            
            // Загрузить при открытии
            loadGifts();
          </script>
        </body>
      </html>
    `);
    template.today = today;
    template.categories = categories;
    return template.evaluate().setWidth(650).setHeight(700);
  }

  getGiftsForDate(dateStr) {
    var date = safeParseDate(dateStr);
    if (!date) return [];
    var data = this.giftsSheet.getDataRange().getValues();
    var result = [];
    var start = new Date(date);
    start.setHours(0, 0, 0, 0);
    var end = new Date(date);
    end.setHours(23, 59, 59, 999);
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var recordDate = row[0];
      if (recordDate && recordDate >= start && recordDate <= end) {
        result.push({
          id: row[6],
          person: row[1],
          description: row[2],
          amount: row[3],
          category: row[4],
          event: row[5]
        });
      }
    }
    return result;
  }

  saveGifts(dateStr, gifts) {
    try {
      var date = safeParseDate(dateStr);
      if (!date) return { success: false, error: 'Неверная дата' };
      
      // 1. Удалить все существующие записи за этот день
      var existing = this.getGiftsForDate(dateStr);
      for (var g of existing) {
        if (g.id) this.deleteGiftRecord(g.id);
      }
      
      // 2. Добавить новые записи
      var added = 0;
      var totalAmount = 0;
      var monthName = MONTHS[date.getMonth()];
      var dateStrFormatted = formatDateForSearch(date).replace(/\./g, '');
      var lastRow = this.giftsSheet.getLastRow();
      
      for (var gift of gifts) {
        var seq = (lastRow + 1).toString().padStart(3, '0');
        var recordId = 'GIFT-' + dateStrFormatted + '-' + seq;
        var newRow = lastRow + 1;
        this.giftsSheet.getRange(newRow, 1).setValue(date);
        this.giftsSheet.getRange(newRow, 2).setValue(gift.person);
        this.giftsSheet.getRange(newRow, 3).setValue(gift.description);
        this.giftsSheet.getRange(newRow, 4).setValue(gift.amount);
        this.giftsSheet.getRange(newRow, 5).setValue(gift.category);
        this.giftsSheet.getRange(newRow, 6).setValue(gift.event);
        this.giftsSheet.getRange(newRow, 7).setValue(recordId);
        this.giftsSheet.getRange(newRow, 8).setValue(monthName);
        lastRow++;
        added++;
        totalAmount += gift.amount;
      }
      
      // 3. Обновить строку компенсации (подарки) в основном листе
      this.updateCompensationRow(date, totalAmount, ROWS.COMPENSATING);
      
      var message = added > 0 
        ? 'Добавлено ' + added + ' подарков на сумму ' + formatMoney(totalAmount)
        : 'Новых подарков не добавлено (все уже существуют)';
      
      return { success: true, message: message, added: added, totalAmount: totalAmount };
      
    } catch (e) {
      logError(e, 'GiftForm.saveGifts');
      return { success: false, error: e.toString() };
    }
  }

  deleteGiftRecord(recordId) {
    var data = this.giftsSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][6] === recordId) {
        this.giftsSheet.deleteRow(i + 1);
        break;
      }
    }
  }

  updateCompensationRow(date, amount, row) {
    var month = date.getMonth();
    var col = month * 2 + 2; // B+C для января и т.д.
    var plusCol = col;
    var minusCol = col + 1;
    
    var currentPlus = safeParseNumber(this.sheet.getRange(row, plusCol).getValue());
    var currentMinus = safeParseNumber(this.sheet.getRange(row, minusCol).getValue());
    
    this.sheet.getRange(row, plusCol).setValue(currentPlus + amount);
    this.sheet.getRange(row, minusCol).setValue(currentMinus + amount);
    
    // Обновляем итоги
    var lastCol = this.sheet.getLastColumn();
    for (var c = 2; c <= lastCol; c++) {
      var total = 0;
      for (var r = ROWS.DATA_START; r <= ROWS.DATA_END; r++) {
        total += safeParseNumber(this.sheet.getRange(r, c).getValue());
      }
      this.sheet.getRange(ROWS.TOTAL, c).setValue(total);
    }
    for (var c = 2; c <= lastCol; c += 2) {
      var income = safeParseNumber(this.sheet.getRange(ROWS.TOTAL, c).getValue());
      var expense = safeParseNumber(this.sheet.getRange(ROWS.TOTAL, c + 1).getValue());
      this.sheet.getRange(ROWS.BALANCE, c + 1).setValue(income - expense);
    }
  }
}

class InvestmentContributionForm {
  constructor() {
    this.sheet = getSheetSafe(SHEETS.CURRENT_YEAR);
    this.investSheet = this.getOrCreateInvestSheet();
  }

  getOrCreateInvestSheet() {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName('Инвестиционные пополнения');
    if (!sheet) {
      sheet = ss.insertSheet('Инвестиционные пополнения');
      sheet.setTabColor('#b6d7a8');
      var headers = [['Дата', 'Счёт', 'Сумма (₽)', 'Назначение', 'Цель', 'Примечание', 'ID записи', 'Месяц']];
      sheet.getRange(1, 1, 1, 8).setValues(headers);
      sheet.getRange(1, 1, 1, 8).setFontWeight('bold').setBackground('#f0f0f0');
      sheet.setColumnWidth(1, 100);
      sheet.setColumnWidth(2, 150);
      sheet.setColumnWidth(3, 120);
      sheet.setColumnWidth(4, 200);
      sheet.setColumnWidth(5, 150);
      sheet.setColumnWidth(6, 200);
      sheet.setColumnWidth(7, 150);
      sheet.setColumnWidth(8, 80);
      sheet.getRange(2, 3, 1000, 1).setNumberFormat('#,##0.00 ₽');
    }
    return sheet;
  }

  show() {
    var ui = SpreadsheetApp.getUi();
    var template = HtmlService.createTemplate(this.getContributionFormHtml());
    var htmlOutput = template.evaluate().setWidth(650).setHeight(700);
    ui.showModalDialog(htmlOutput, '📈 Инвестиционные пополнения');
  }

  getContributionFormHtml() {
    return `
      <!DOCTYPE html>
      <html>
        <head>
          <base target="_top">
          <style>
            *{box-sizing:border-box} body{font-family:Arial,sans-serif;padding:20px;background:#f5f5f5}
            .container{max-width:650px;margin:0 auto;background:white;padding:25px;border-radius:12px}
            h2{color:#1a73e8;margin-top:0}.form-group{margin-bottom:15px}label{display:block;margin-bottom:5px;font-weight:500}
            input,select,textarea{width:100%;padding:10px;border:1px solid #ddd;border-radius:6px}
            button{padding:10px;border:none;border-radius:6px;cursor:pointer}
            .btn-success{background:#34a853;color:white}.btn-primary{background:#1a73e8;color:white}
            .btn-secondary{background:#f1f3f4;color:#333}.row{display:flex;gap:10px;margin-top:15px}
            .contribution-item{background:#f8f9fa;padding:12px;margin-bottom:8px;border-radius:6px;border-left:4px solid #34a853;display:flex;justify-content:space-between}
            .list-container{max-height:300px;overflow-y:auto;margin-bottom:20px}
          </style>
        </head>
        <body>
          <div class="container">
            <h2>📈 Инвестиционные пополнения</h2>
            <div class="date-group"><label>Дата:</label><input type="text" id="contributionDate" value="<?= formatDate(new Date()) ?>" onchange="loadContributions()"></div>
            <div class="list-container" id="contributionsList"></div>
            <div class="total">Итого: <span id="totalAmount">0.00 ₽</span></div>
            <div class="form-group"><label>Счёт:</label><select id="account"><option>Сбер Инвестиции</option><option>Тинькофф Инвестиции</option><option>ИИС</option><option>Брокерский счёт</option><option>Другое</option></select></div>
            <div class="form-group"><label>Сумма (₽):</label><input type="number" id="amount" step="0.01"></div>
            <div class="form-group"><label>Назначение:</label><input type="text" id="purpose" placeholder="Например: Покупка акций"></div>
            <div class="form-group"><label>Цель:</label><select id="goal"><option>Накопление</option><option>Пенсия</option><option>Крупная покупка</option><option>Дивидендный доход</option><option>Другое</option></select></div>
            <div class="form-group"><label>Примечание:</label><textarea id="note" rows="2"></textarea></div>
            <div class="row"><button class="btn-success" onclick="addContribution()">➕ Добавить</button><button class="btn-primary" onclick="saveAll()">💾 Сохранить все</button><button class="btn-secondary" onclick="google.script.host.close()">❌ Отмена</button></div>
          </div>
          <script>
            let contributions = [];
            function loadContributions() {
              var dateStr = document.getElementById('contributionDate').value.trim();
              if (dateStr) google.script.run.withSuccessHandler(function(data) {
                contributions = data.map(c => ({ id: c.id, account: c.account, amount: c.amount, purpose: c.purpose, goal: c.goal, note: c.note }));
                renderContributions();
              }).getContributionsForDate(dateStr);
            }
            function addContribution() {
              var account = document.getElementById('account').value;
              var amount = parseFloat(document.getElementById('amount').value);
              var purpose = document.getElementById('purpose').value.trim();
              var goal = document.getElementById('goal').value;
              var note = document.getElementById('note').value.trim();
              if (isNaN(amount) || amount <= 0) { alert('Введите корректную сумму'); return; }
              if (!purpose) { alert('Введите назначение'); return; }
              contributions.push({ id: null, account: account, amount: amount, purpose: purpose, goal: goal, note: note });
              document.getElementById('amount').value = ''; document.getElementById('purpose').value = ''; document.getElementById('note').value = '';
              renderContributions();
            }
            function deleteContribution(idx) { if (confirm('Удалить это пополнение?')) { contributions.splice(idx, 1); renderContributions(); } }
            function renderContributions() {
              var container = document.getElementById('contributionsList');
              var total = 0;
              if (contributions.length === 0) container.innerHTML = '<p>Нет добавленных пополнений</p>';
              else {
                container.innerHTML = '';
                contributions.forEach((c, idx) => {
                  total += c.amount;
                  var div = document.createElement('div');
                  div.className = 'contribution-item';
                  div.innerHTML = '<div><strong>' + escapeHtml(c.account) + '</strong><br>' + escapeHtml(c.purpose) + '<br><small>' + escapeHtml(c.goal) + (c.note ? ' — ' + escapeHtml(c.note) : '') + '</small></div>' +
                    '<div>' + c.amount.toFixed(2) + ' ₽ <button onclick="deleteContribution(' + idx + ')">✕</button></div>';
                  container.appendChild(div);
                });
              }
              document.getElementById('totalAmount').textContent = total.toFixed(2) + ' ₽';
            }
            function saveAll() {
              if (contributions.length === 0 && !confirm('Нет пополнений. Продолжить?')) return;
              var dateStr = document.getElementById('contributionDate').value.trim();
              if (!dateStr) { alert('Введите дату'); return; }
              google.script.run.saveContributions(dateStr, contributions);
            }
            function escapeHtml(str) { if (!str) return ''; return str.replace(/[&<>]/g, function(m) { if (m === '&') return '&amp;'; if (m === '<') return '&lt;'; if (m === '>') return '&gt;'; return m; }); }
            loadContributions();
          </script>
        </body>
      </html>
    `;
  }

  getContributionsForDate(dateStr) {
    var date = safeParseDate(dateStr);
    if (!date) return [];
    var data = this.investSheet.getDataRange().getValues();
    var result = [];
    var start = new Date(date); start.setHours(0,0,0,0);
    var end = new Date(date); end.setHours(23,59,59,999);
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var recordDate = row[0];
      if (recordDate && recordDate >= start && recordDate <= end) {
        result.push({ id: row[6], account: row[1], amount: row[2], purpose: row[3], goal: row[4], note: row[5] });
      }
    }
    return result;
  }

  saveContributions(dateStr, contributions) {
    try {
      var date = safeParseDate(dateStr);
      if (!date) return { success: false, error: 'Неверная дата' };
      var existing = this.getContributionsForDate(dateStr);
      for (var c of existing) if (c.id) this.deleteContributionRecord(c.id);
      var added = 0, totalAmount = 0;
      var monthName = MONTHS[date.getMonth()];
      var dateStrFormatted = formatDateForSearch(date).replace(/\./g, '');
      var lastRow = this.investSheet.getLastRow();
      for (var contrib of contributions) {
        var seq = (lastRow + 1).toString().padStart(3, '0');
        var recordId = 'INV-' + dateStrFormatted + '-' + seq;
        var newRow = lastRow + 1;
        this.investSheet.getRange(newRow, 1).setValue(date);
        this.investSheet.getRange(newRow, 2).setValue(contrib.account);
        this.investSheet.getRange(newRow, 3).setValue(contrib.amount);
        this.investSheet.getRange(newRow, 4).setValue(contrib.purpose);
        this.investSheet.getRange(newRow, 5).setValue(contrib.goal);
        this.investSheet.getRange(newRow, 6).setValue(contrib.note);
        this.investSheet.getRange(newRow, 7).setValue(recordId);
        this.investSheet.getRange(newRow, 8).setValue(monthName);
        lastRow++;
        added++;
        totalAmount += contrib.amount;
      }
      this.updateCompensationRow(date, totalAmount, ROWS.COMPENSATING2);
      showMessage('Успешно', 'Добавлено ' + added + ' пополнений на сумму ' + formatMoney(totalAmount), 'info');
      addLogEntry('Добавлено ' + added + ' пополнений на сумму ' + formatMoney(totalAmount), 'INFO', 'InvestmentContributionForm');
      return { success: true, added: added, totalAmount: totalAmount };
    } catch(e) { logError(e, 'InvestmentContributionForm.saveContributions'); return { success: false, error: e.toString() }; }
  }

  deleteContributionRecord(recordId) {
    var data = this.investSheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][6] === recordId) {
        this.investSheet.deleteRow(i + 1);
        break;
      }
    }
  }

  updateCompensationRow(date, amount, row) {
    var month = date.getMonth();
    var col = month * 2 + 2;
    var currentPlus = safeParseNumber(this.sheet.getRange(row, col).getValue());
    var currentMinus = safeParseNumber(this.sheet.getRange(row, col + 1).getValue());
    this.sheet.getRange(row, col).setValue(currentPlus + amount);
    this.sheet.getRange(row, col + 1).setValue(currentMinus + amount);
    var lastCol = this.sheet.getLastColumn();
    for (var c = 2; c <= lastCol; c++) {
      var total = 0;
      for (var r = ROWS.DATA_START; r <= ROWS.DATA_END; r++) total += safeParseNumber(this.sheet.getRange(r, c).getValue());
      this.sheet.getRange(ROWS.TOTAL, c).setValue(total);
    }
    for (var c = 2; c <= lastCol; c += 2) {
      var income = safeParseNumber(this.sheet.getRange(ROWS.TOTAL, c).getValue());
      var expense = safeParseNumber(this.sheet.getRange(ROWS.TOTAL, c + 1).getValue());
      this.sheet.getRange(ROWS.BALANCE, c + 1).setValue(income - expense);
    }
  }
}

function showGiftForm() { new GiftForm().show(); }
function getGiftsForDate(dateStr) { return new GiftForm().getGiftsForDate(dateStr); }
function saveGifts(dateStr, gifts) { return new GiftForm().saveGifts(dateStr, gifts); }
function showInvestmentContributionForm() { new InvestmentContributionForm().show(); }
function getContributionsForDate(dateStr) { return new InvestmentContributionForm().getContributionsForDate(dateStr); }
function saveContributions(dateStr, contributions) { return new InvestmentContributionForm().saveContributions(dateStr, contributions); }
function showAddGiftDialog() { showGiftForm(); }
function showAddInvestmentDialog() { showInvestmentContributionForm(); }
function initSpecialSheets() { new GiftForm().getOrCreateGiftsSheet(); new InvestmentContributionForm().getOrCreateInvestSheet(); showMessage('Готово', 'Листы для специальных операций созданы', 'info'); }
function showGiftsReport() { showMessage('Информация', 'Отчёт по подаркам можно посмотреть в листе "Подарки"', 'info'); }
function showInvestmentsReport() { showMessage('Информация', 'Отчёт по инвестиционным пополнениям можно посмотреть в листе "Инвестиционные пополнения"', 'info'); }