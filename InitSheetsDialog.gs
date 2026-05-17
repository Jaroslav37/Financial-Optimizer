/**
 * ДИАЛОГ ВЫБОРОЧНОЙ ИНИЦИАЛИЗАЦИИ ЛИСТОВ
 */

function showInitSheetsDialog() {
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html>
    <html>
    <head>
      <base target="_top">
      <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; padding: 20px; background: #f5f5f5; }
        .container { max-width: 500px; margin: 0 auto; background: white; padding: 25px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); }
        h2 { color: #1a73e8; margin-top: 0; }
        .sheet-list { margin: 20px 0; max-height: 400px; overflow-y: auto; border: 1px solid #ddd; border-radius: 8px; padding: 10px; }
        .sheet-item { margin-bottom: 12px; display: flex; align-items: center; }
        .sheet-item input { margin-right: 10px; transform: scale(1.2); }
        .sheet-item label { font-weight: normal; cursor: pointer; }
        button { width: 100%; padding: 12px; background: #1a73e8; color: white; border: none; border-radius: 6px; cursor: pointer; font-size: 16px; margin-top: 10px; }
        button:hover { background: #1557b0; }
        .select-all { margin-bottom: 15px; display: flex; align-items: center; }
        .select-all input { margin-right: 10px; transform: scale(1.2); }
        .status { margin-top: 15px; font-size: 14px; color: #333; white-space: pre-wrap; }
      </style>
    </head>
    <body>
      <div class="container">
        <h2>📋 Инициализация листов</h2>
        <p>Выберите листы, которые нужно создать или пересоздать.<br>
        <strong>Внимание:</strong> при пересоздании все существующие данные в этих листах будут потеряны.</p>
        
        <div class="select-all">
          <input type="checkbox" id="selectAll" onclick="toggleAll(this.checked)">
          <label for="selectAll"><strong>✅ Выбрать все</strong></label>
        </div>
        
        <div class="sheet-list" id="sheetList">
          <!-- чекбоксы будут заполнены скриптом -->
        </div>
        
        <button onclick="initSelected()">🚀 Инициализировать выбранные листы</button>
        <div id="status" class="status"></div>
      </div>
      
      <script>
        // Список листов с их идентификаторами и отображаемыми названиями
        const sheets = [
          { id: 'budget', name: '💰 Бюджет', func: 'initBudgetSheet' },
          { id: 'debts', name: '💸 Долговые обязательства', func: 'initDebtsSheet' },
          { id: 'expenseDetails', name: '📝 Детали расходов', func: 'initExpenseDetails' },
          { id: 'incomeDetails', name: '📈 Детали доходов', func: 'initIncomeDetails' },
          { id: 'giftsInvest', name: '🎁 Подарки и инвестиционные пополнения', func: 'initSpecialSheets' },
          { id: 'bonds', name: '📊 Облигации (калькулятор)', func: 'initBondSheet' },
          { id: 'financialPlan', name: '📅 Финансовый план', func: 'initFinancialPlanSheet' },
          { id: 'investAnalytics', name: '📉 Инвест аналитика', func: 'createAnalysisSheet' },
          { id: 'logs', name: '📋 Логи', func: 'createLogsSheet' }
        ];
        
        function renderSheetList() {
          const container = document.getElementById('sheetList');
          container.innerHTML = '';
          sheets.forEach(sheet => {
            const div = document.createElement('div');
            div.className = 'sheet-item';
            div.innerHTML = \`
              <input type="checkbox" id="\${sheet.id}" value="\${sheet.func}">
              <label for="\${sheet.id}">\${sheet.name}</label>
            \`;
            container.appendChild(div);
          });
        }
        
        function toggleAll(checked) {
          const checkboxes = document.querySelectorAll('#sheetList input[type="checkbox"]');
          checkboxes.forEach(cb => cb.checked = checked);
        }
        
        function initSelected() {
          const selected = [];
          const checkboxes = document.querySelectorAll('#sheetList input[type="checkbox"]:checked');
          checkboxes.forEach(cb => selected.push(cb.value));
          
          if (selected.length === 0) {
            document.getElementById('status').innerHTML = '⚠️ Не выбран ни один лист.';
            return;
          }
          
          document.getElementById('status').innerHTML = '⏳ Инициализация... Пожалуйста, подождите.';
          const button = document.querySelector('button');
          button.disabled = true;
          
          google.script.run
            .withSuccessHandler(result => {
              document.getElementById('status').innerHTML = result.message;
              button.disabled = false;
            })
            .withFailureHandler(error => {
              document.getElementById('status').innerHTML = '❌ Ошибка: ' + error.message;
              button.disabled = false;
            })
            .runInitSheets(selected);
        }
        
        renderSheetList();
      </script>
    </body>
    </html>
  `)
  .setWidth(550)
  .setHeight(600)
  .setTitle('Инициализация листов');
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Инициализация листов');
}

function runInitSheets(selectedFunctions) {
  let results = [];
  for (let i = 0; i < selectedFunctions.length; i++) {
    const funcName = selectedFunctions[i];
    try {
      // Вызов глобальной функции по имени
      const func = this[funcName];
      if (typeof func === 'function') {
        func();
        results.push(`✅ ${funcName} – выполнено`);
      } else {
        results.push(`⚠️ ${funcName} – функция не найдена`);
      }
    } catch(e) {
      results.push(`❌ ${funcName} – ошибка: ${e.toString()}`);
    }
  }
  return { message: results.join('\n') };
}