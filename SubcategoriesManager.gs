/**
 * МЕНЕДЖЕР ПОДКАТЕГОРИЙ — справочники для доходов и расходов
 * Версия 1.1 — загрузка в UnifiedDayForm + диалоги добавления
 */

class SubcategoriesManager {
  constructor() {
    this.expenseSheet = this._getOrCreateSheet('Подкатегории расходов');
    this.incomeSheet  = this._getOrCreateSheet('Подкатегории доходов');
  }

  _getOrCreateSheet(sheetName) {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.setTabColor(sheetName.includes('доходов') ? '#d9ead3' : '#fce5cd');
      const headers = ['Категория', 'Подкатегория'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
      sheet.setColumnWidth(1, 200);
      sheet.setColumnWidth(2, 200);
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  /** Возвращает объект { Категория: [подкат1, подкат2, ...] } */
  getCategoryMap(sheet) {
    const data = sheet.getDataRange().getValues();
    const map = {};
    for (let i = 1; i < data.length; i++) {
      const cat = (data[i][0] || '').toString().trim();
      const sub = (data[i][1] || '').toString().trim();
      if (!cat || !sub) continue;
      if (!map[cat]) map[cat] = [];
      if (map[cat].indexOf(sub) === -1) map[cat].push(sub);
    }
    return map;
  }

  /** Возвращает объединённые подкатегории расходов (базовые + из справочника) */
  getExpenseSubcategories(baseCategories) {
    const extra = this.getCategoryMap(this.expenseSheet);
    const merged = {};
    for (const cat of Object.keys(baseCategories)) {
      merged[cat] = (baseCategories[cat] || []).slice();
    }
    for (const cat of Object.keys(extra)) {
      if (!merged[cat]) merged[cat] = [];
      const existing = merged[cat];
      for (const sub of extra[cat]) {
        if (existing.indexOf(sub) === -1) existing.push(sub);
      }
    }
    return merged;
  }

  /** Возвращает объединённые подкатегории доходов */
  getIncomeSubcategories(baseCategories) {
    const extra = this.getCategoryMap(this.incomeSheet);
    const merged = {};
    for (const cat of Object.keys(baseCategories)) {
      merged[cat] = (baseCategories[cat] || []).slice();
    }
    for (const cat of Object.keys(extra)) {
      if (!merged[cat]) merged[cat] = [];
      const existing = merged[cat];
      for (const sub of extra[cat]) {
        if (existing.indexOf(sub) === -1) existing.push(sub);
      }
    }
    return merged;
  }

  /** Добавляет одну запись в справочник (диалог) */
  showAddDialog(type) {
    const isIncome = (type === 'income');
    const sheetName = isIncome ? 'Подкатегории доходов' : 'Подкатегории расходов';
    const title = isIncome ? 'Добавить подкатегорию дохода' : 'Добавить подкатегорию расхода';
    
    const html = HtmlService.createHtmlOutput(`
      <!DOCTYPE html><html><head><base target="_top">
      <style>
        body{font-family:Arial;padding:20px} input,select{width:100%;padding:8px;margin:5px 0;border:1px solid #ddd;border-radius:6px}
        button{width:100%;padding:10px;background:#1a73e8;color:white;border:none;border-radius:6px;cursor:pointer;margin-top:10px}
        .cancel{background:#f1f3f4;color:#333}
      </style></head><body>
        <h3>${title}</h3>
        <label>Категория:</label><input type="text" id="cat" placeholder="Существующая или новая">
        <label>Подкатегория:</label><input type="text" id="sub" placeholder="Новая подкатегория">
        <button onclick="save()">💾 Сохранить</button>
        <button class="cancel" onclick="google.script.host.close()">Отмена</button>
        <script>
          function save() {
            var cat = document.getElementById('cat').value.trim();
            var sub = document.getElementById('sub').value.trim();
            if (!cat || !sub) { alert('Заполните оба поля'); return; }
            google.script.run.withSuccessHandler(function() {
              alert('✅ Добавлено в справочник «${sheetName}»');
              google.script.host.close();
            }).addSubcategory('${type}', cat, sub);
          }
        </script>
      </body></html>
    `).setWidth(400).setHeight(300);
    
    SpreadsheetApp.getUi().showModalDialog(html, title);
  }
}

// ─── Глобальные функции ───────────────────────────

function getExpenseSubcategories(baseCategoriesJson) {
  const base = JSON.parse(baseCategoriesJson);
  return new SubcategoriesManager().getExpenseSubcategories(base);
}

function getIncomeSubcategories(baseCategoriesJson) {
  const base = JSON.parse(baseCategoriesJson);
  return new SubcategoriesManager().getIncomeSubcategories(base);
}

function addSubcategory(type, category, subcategory) {
  const mgr = new SubcategoriesManager();
  const sheet = type === 'income' ? mgr.incomeSheet : mgr.expenseSheet;
  const lr = sheet.getLastRow() + 1;
  sheet.getRange(lr, 1).setValue(category);
  sheet.getRange(lr, 2).setValue(subcategory);
  return true;
}

function showAddExpenseSubcategoryDialog() {
  new SubcategoriesManager().showAddDialog('expense');
}

function showAddIncomeSubcategoryDialog() {
  new SubcategoriesManager().showAddDialog('income');
}