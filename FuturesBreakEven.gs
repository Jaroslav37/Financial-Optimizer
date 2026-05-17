/**
 * FUTURES BREAK‑EVEN – учёт сделок в пунктах и расчёт безубыточности (Лонг + Шорт)
 * Версия 2.4 – выпадающий список контрактов из таблицы
 */

class FuturesBreakEven {
  constructor() {
    this.sheet = this._getOrCreateSheet();
    this.refSheet = this._getOrCreateRefSheet();
  }

  _getOrCreateSheet() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Фьючерсы');
    if (!sheet) {
      sheet = ss.insertSheet('Фьючерсы');
      sheet.setTabColor('#e69138');
      const headers = ['Дата', 'Контракт', 'Тип (Покупка/Продажа)', 'Количество', 'Пункты', 'Комиссия биржи (₽)', 'Комиссия брокера (₽)'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
      sheet.setColumnWidth(1, 100); sheet.setColumnWidth(2, 100); sheet.setColumnWidth(3, 120);
      sheet.setColumnWidth(4, 80); sheet.setColumnWidth(5, 80); sheet.setColumnWidth(6, 120); sheet.setColumnWidth(7, 120);
      sheet.setFrozenRows(1);
      sheet.getRange(2, 1, 1000, 1).setNumberFormat('dd.MM.yyyy');
      sheet.getRange(2, 4, 1000, 1).setNumberFormat('#,##0');
      sheet.getRange(2, 5, 1000, 1).setNumberFormat('#,##0');
      sheet.getRange(2, 6, 1000, 2).setNumberFormat('#,##0.00 ₽');
    }
    return sheet;
  }

  _getOrCreateRefSheet() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Справочник фьючерсов');
    if (!sheet) {
      sheet = ss.insertSheet('Справочник фьючерсов');
      sheet.setTabColor('#a4c2f4');
      const headers = ['Контракт', 'Цена пункта (₽)'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
      sheet.setColumnWidth(1, 150); sheet.setColumnWidth(2, 120);
      sheet.setFrozenRows(1);
      sheet.getRange(2, 2, 1000, 1).setNumberFormat('#,##0.00');
    }
    return sheet;
  }

  /** Собирает список уникальных контрактов из листа сделок */
  getContracts() {
    const data = this.sheet.getDataRange().getValues();
    const contracts = new Set();
    for (let i = 1; i < data.length; i++) {
      const contract = (data[i][1] || '').toString().trim().toUpperCase();
      if (contract) contracts.add(contract);
    }
    return [...contracts].sort();
  }

  _getPointPrice(contract) {
    const data = this.refSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === contract.toUpperCase()) {
        return parseFloat(data[i][1]) || 0;
      }
    }
    return null;
  }

  _setPointPrice(contract, price) {
    const data = this.refSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toUpperCase() === contract.toUpperCase()) {
        this.refSheet.getRange(i + 1, 2).setValue(price); return;
      }
    }
    const lr = this.refSheet.getLastRow() + 1;
    this.refSheet.getRange(lr, 1).setValue(contract);
    this.refSheet.getRange(lr, 2).setValue(price);
  }

  addTrade(date, contract, side, qty, points, exchangeComm, brokerComm) {
    const lr = this.sheet.getLastRow() + 1;
    this.sheet.getRange(lr, 1).setValue(date);
    this.sheet.getRange(lr, 2).setValue(contract);
    this.sheet.getRange(lr, 3).setValue(side);
    this.sheet.getRange(lr, 4).setValue(qty);
    this.sheet.getRange(lr, 5).setValue(points);
    this.sheet.getRange(lr, 6).setValue(exchangeComm);
    this.sheet.getRange(lr, 7).setValue(brokerComm);
    return true;
  }

  calculateBreakeven(contract) {
    const pointPrice = this._getPointPrice(contract);
    if (!pointPrice || pointPrice <= 0) {
      return { contract, error: `Не задана цена пункта для контракта "${contract}". Добавьте её в справочник.` };
    }

    const data = this.sheet.getDataRange().getValues();
    const trades = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[1] === contract && row[0]) {
        trades.push({
          date: row[0], side: row[2],
          qty: parseFloat(row[3]) || 0,
          points: parseFloat(row[4]) || 0,
          exchangeComm: parseFloat(row[5]) || 0,
          brokerComm: parseFloat(row[6]) || 0,
          totalCommission: (parseFloat(row[5]) || 0) + (parseFloat(row[6]) || 0)
        });
      }
    }

    if (trades.length === 0) return { contract, error: 'Нет сделок по этому контракту' };

    trades.sort((a, b) => new Date(a.date) - new Date(b.date));

    let netPosition = 0, totalBuyCostRub = 0, totalBuyQty = 0;
    let totalSellRevenueRub = 0, totalSellQty = 0, totalCommission = 0;

    for (const t of trades) {
      const priceRub = t.points * pointPrice;
      if (t.side === 'Покупка') {
        netPosition += t.qty; totalBuyCostRub += t.qty * priceRub; totalBuyQty += t.qty; totalCommission += t.totalCommission;
      } else {
        netPosition -= t.qty; totalSellRevenueRub += t.qty * priceRub; totalSellQty += t.qty; totalCommission += t.totalCommission;
      }
    }

    if (netPosition === 0) return { contract, position: 0, message: 'Позиция закрыта' };

    const lastPriceRub = trades[trades.length - 1].points * pointPrice;
    let result = { contract, pointPrice };

    if (netPosition > 0) {
      const avgPriceRub = totalBuyQty > 0 ? totalBuyCostRub / totalBuyQty : 0;
      const breakevenRub = avgPriceRub + (totalCommission / netPosition);
      result = {
        ...result, direction: 'Лонг', position: netPosition,
        avgPriceRub: Math.round(avgPriceRub * 100) / 100,
        avgPricePoints: Math.round(avgPriceRub / pointPrice * 100) / 100,
        totalCommission: Math.round(totalCommission * 100) / 100,
        breakevenRub: Math.round(breakevenRub * 100) / 100,
        breakevenPoints: Math.round(breakevenRub / pointPrice * 100) / 100,
        lastPriceRub: Math.round(lastPriceRub * 100) / 100,
      };
      result.recommendation = lastPriceRub >= breakevenRub
        ? '✅ Цена выше безубыточности. Можно продавать с прибылью.'
        : '⏳ Цена ниже безубыточности. Дождитесь роста или усредните позицию.';
    } else {
      const absPos = Math.abs(netPosition);
      const avgSellPriceRub = totalSellQty > 0 ? totalSellRevenueRub / totalSellQty : 0;
      const breakevenRub = avgSellPriceRub - (totalCommission / absPos);
      result = {
        ...result, direction: 'Шорт', position: absPos,
        avgPriceRub: Math.round(avgSellPriceRub * 100) / 100,
        avgPricePoints: Math.round(avgSellPriceRub / pointPrice * 100) / 100,
        totalCommission: Math.round(totalCommission * 100) / 100,
        breakevenRub: Math.round(breakevenRub * 100) / 100,
        breakevenPoints: Math.round(breakevenRub / pointPrice * 100) / 100,
        lastPriceRub: Math.round(lastPriceRub * 100) / 100,
      };
      result.recommendation = lastPriceRub <= breakevenRub
        ? '✅ Цена ниже безубыточности. Можно покупать с прибылью.'
        : '⏳ Цена выше безубыточности. Дождитесь снижения или продавайте ещё.';
    }
    return result;
  }

  showAllBreakeven() {
    const data = this.sheet.getDataRange().getValues();
    const contracts = new Set();
    for (let i = 1; i < data.length; i++) { if (data[i][1]) contracts.add(data[i][1]); }
    if (contracts.size === 0) { showMessage('Фьючерсы', 'Нет данных о сделках', 'warning'); return; }

    let report = '📈 ТОЧКИ БЕЗУБЫТОЧНОСТИ ФЬЮЧЕРСОВ\n═══════════════════════════════\n\n';
    for (const contract of contracts) {
      const result = this.calculateBreakeven(contract);
      if (result.error) { report += `⚠️ ${contract}: ${result.error}\n\n`; continue; }
      if (result.position === 0) { report += `🔹 ${contract}: позиция закрыта\n\n`; continue; }
      const dirIcon = result.direction === 'Лонг' ? '📈' : '📉';
      report += `${dirIcon} ${contract} (${result.direction}, цена пункта ${formatMoney(result.pointPrice)})\n`;
      report += `   Позиция: ${result.position} контрактов\n`;
      report += `   Средняя цена: ${formatMoney(result.avgPriceRub)} (${result.avgPricePoints} п.)\n`;
      report += `   Комиссии: ${formatMoney(result.totalCommission)}\n`;
      report += `   📉 Последняя цена: ${formatMoney(result.lastPriceRub)}\n`;
      report += `   🎯 Безубыточность: ${formatMoney(result.breakevenRub)} (${result.breakevenPoints} п.)\n`;
      report += `   💡 ${result.recommendation}\n\n`;
    }
    showMessage('Безубыточность фьючерсов', report);
  }

  showAddTradeDialog() {
    const contracts = this.getContracts();
    const contractOptions = contracts.map(c => `<option value="${c}">${c}</option>`).join('');
    
    const html = HtmlService.createHtmlOutput(`
      <!DOCTYPE html>
      <html><head><base target="_top">
      <style>
        body{font-family:Arial;padding:20px} input,select{width:100%;padding:8px;margin:5px 0;border:1px solid #ddd;border-radius:6px}
        button{width:100%;padding:10px;background:#1a73e8;color:white;border:none;border-radius:6px;cursor:pointer;margin-top:10px}
        .cancel{background:#f1f3f4;color:#333} .hint{font-size:12px;color:#666} .new-contract-row{display:none;margin-top:8px}
      </style></head><body>
        <h3>📈 Добавить сделку (пункты)</h3>
        <label>Дата:</label><input type="text" id="date" value="${formatDate(new Date())}">
        <label>Контракт:</label>
        <select id="contractSelect" onchange="handleContractChange()">
          <option value="">-- Выберите контракт --</option>
          ${contractOptions}
          <option value="__new__">➕ Ввести новый...</option>
        </select>
        <div id="newContractRow" class="new-contract-row">
          <label>Новый контракт:</label>
          <input type="text" id="newContract" placeholder="Введите новый контракт">
        </div>
        <label>Тип:</label><select id="side"><option value="Покупка">Покупка</option><option value="Продажа">Продажа</option></select>
        <label>Количество:</label><input type="number" id="qty" step="1" value="1">
        <label>Цена (пункты):</label><input type="number" id="points" step="0.01">
        <label>Цена пункта (₽):</label>
        <input type="number" id="pointPrice" step="0.01"><span class="hint" id="ppHint"></span>
        <label>Комиссия биржи (₽):</label><input type="number" id="exchangeComm" step="0.01" value="0">
        <label>Комиссия брокера (₽):</label><input type="number" id="brokerComm" step="0.01" value="0">
        <button onclick="save()">💾 Сохранить сделку</button>
        <button class="cancel" onclick="google.script.host.close()">Отмена</button>
        <script>
          document.getElementById('contractSelect').addEventListener('change', function() {
            handleContractChange();
          });
          function handleContractChange() {
            var sel = document.getElementById('contractSelect');
            var newRow = document.getElementById('newContractRow');
            if (sel.value === '__new__') {
              newRow.style.display = 'block';
              document.getElementById('newContract').focus();
              document.getElementById('ppHint').innerText = '(введите цену пункта – будет сохранена)';
              document.getElementById('pointPrice').value = '';
            } else {
              newRow.style.display = 'none';
              var c = sel.value.trim();
              if (c) {
                google.script.run.withSuccessHandler(function(price) {
                  if (price) {
                    document.getElementById('pointPrice').value = price;
                    document.getElementById('ppHint').innerText = '(найдено в справочнике)';
                  } else {
                    document.getElementById('pointPrice').value = '';
                    document.getElementById('ppHint').innerText = '(введите цену пункта – будет сохранена)';
                  }
                }).getPointPrice(c);
              }
            }
          }
          function save() {
            var date = document.getElementById('date').value.trim();
            var contractSelect = document.getElementById('contractSelect');
            var contract = contractSelect.value === '__new__' ? document.getElementById('newContract').value.trim().toUpperCase() : contractSelect.value;
            if (!contract) { alert('Выберите или введите контракт'); return; }
            var side = document.getElementById('side').value;
            var qty = parseInt(document.getElementById('qty').value);
            var points = parseFloat(document.getElementById('points').value);
            var pointPrice = parseFloat(document.getElementById('pointPrice').value);
            var exchangeComm = parseFloat(document.getElementById('exchangeComm').value) || 0;
            var brokerComm = parseFloat(document.getElementById('brokerComm').value) || 0;
            if (!date || isNaN(qty) || isNaN(points) || isNaN(pointPrice)) { alert('Заполните все поля'); return; }
            google.script.run.withSuccessHandler(function() { alert('✅ Сделка добавлена'); google.script.host.close(); })
              .addFuturesTradeFull(date, contract, side, qty, points, pointPrice, exchangeComm, brokerComm);
          }
        </script>
      </body></html>
    `).setWidth(400).setHeight(580);
    SpreadsheetApp.getUi().showModalDialog(html, 'Добавить сделку (пункты)');
  }
}

// Глобальные функции
function showFuturesBreakeven() { new FuturesBreakEven().showAllBreakeven(); }
function showAddFuturesTrade() { new FuturesBreakEven().showAddTradeDialog(); }
function addFuturesTradeFull(dateStr, contract, side, qty, points, pointPrice, exchangeComm, brokerComm) {
  const date = safeParseDate(dateStr); if (!date) return false;
  const fb = new FuturesBreakEven();
  if (pointPrice > 0) fb._setPointPrice(contract, pointPrice);
  return fb.addTrade(date, contract, side, qty, points, exchangeComm, brokerComm);
}
function getPointPrice(contract) { return new FuturesBreakEven()._getPointPrice(contract); }