/**
 * STOCKS BREAK‑EVEN – расчёт безубыточности акций с учётом сплитов и комиссий
 * Версия 1.3 – рекомендации по рыночной цене
 */

class StocksBreakEven {
  constructor() {
    this.tradesSheet = this._getOrCreateTradesSheet();
    this.splitsSheet = this._getOrCreateSplitsSheet();
  }

  _getOrCreateTradesSheet() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Сделки по акциям');
    if (!sheet) {
      sheet = ss.insertSheet('Сделки по акциям');
      sheet.setTabColor('#a4c2f4');
      const headers = ['Дата', 'Тикер', 'Тип (Покупка/Продажа)', 'Количество', 'Цена (₽)', 'Комиссия биржи (₽)', 'Комиссия брокера (₽)'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
      sheet.setColumnWidth(1, 100); sheet.setColumnWidth(2, 100);
      sheet.setColumnWidth(3, 120); sheet.setColumnWidth(4, 80);
      sheet.setColumnWidth(5, 100); sheet.setColumnWidth(6, 120); sheet.setColumnWidth(7, 120);
      sheet.setFrozenRows(1);
      sheet.getRange(2, 1, 1000, 1).setNumberFormat('dd.MM.yyyy');
      sheet.getRange(2, 4, 1000, 1).setNumberFormat('#,##0');
      sheet.getRange(2, 5, 1000, 1).setNumberFormat('#,##0.00');
      sheet.getRange(2, 6, 1000, 2).setNumberFormat('#,##0.00 ₽');
    }
    return sheet;
  }

  _getOrCreateSplitsSheet() {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Сплиты акций');
    if (!sheet) {
      sheet = ss.insertSheet('Сплиты акций');
      sheet.setTabColor('#f9cb9c');
      const headers = ['Дата', 'Тикер', 'Тип события', 'Коэффициент'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#f0f0f0');
      sheet.setColumnWidth(1, 100); sheet.setColumnWidth(2, 100);
      sheet.setColumnWidth(3, 130); sheet.setColumnWidth(4, 100);
      sheet.setFrozenRows(1);
      sheet.getRange(2, 1, 1000, 1).setNumberFormat('dd.MM.yyyy');
      sheet.getRange(2, 4, 1000, 1).setNumberFormat('0.000');
    }
    return sheet;
  }

  getStockTickers() {
    const data = this.tradesSheet.getDataRange().getValues();
    const tickers = new Set();
    for (let i = 1; i < data.length; i++) {
      const ticker = (data[i][1] || '').toString().trim().toUpperCase();
      if (ticker) tickers.add(ticker);
    }
    return [...tickers].sort();
  }

  addTrade(date, ticker, side, qty, price, exchangeComm, brokerComm) {
    const lr = this.tradesSheet.getLastRow() + 1;
    this.tradesSheet.getRange(lr, 1).setValue(date);
    this.tradesSheet.getRange(lr, 2).setValue(ticker.toUpperCase());
    this.tradesSheet.getRange(lr, 3).setValue(side);
    this.tradesSheet.getRange(lr, 4).setValue(qty);
    this.tradesSheet.getRange(lr, 5).setValue(price);
    this.tradesSheet.getRange(lr, 6).setValue(exchangeComm);
    this.tradesSheet.getRange(lr, 7).setValue(brokerComm);
    return true;
  }

  addSplit(date, ticker, type, ratio) {
    const lr = this.splitsSheet.getLastRow() + 1;
    this.splitsSheet.getRange(lr, 1).setValue(date);
    this.splitsSheet.getRange(lr, 2).setValue(ticker.toUpperCase());
    this.splitsSheet.getRange(lr, 3).setValue(type === 'split' ? 'Сплит' : 'Консолидация');
    this.splitsSheet.getRange(lr, 4).setValue(ratio);
    return true;
  }

  _getSplits(ticker) {
    const data = this.splitsSheet.getDataRange().getValues();
    const splits = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim().toUpperCase() === ticker.toUpperCase()) {
        splits.push({
          date: data[i][0],
          type: data[i][2] === 'Сплит' ? 'split' : 'reverse',
          ratio: parseFloat(data[i][3]) || 1
        });
      }
    }
    return splits.sort((a, b) => a.date - b.date);
  }

  calculateBreakeven(ticker) {
    const rawTrades = [];
    const tradesData = this.tradesSheet.getDataRange().getValues();
    for (let i = 1; i < tradesData.length; i++) {
      const row = tradesData[i];
      if (row[1] && row[1].toString().trim().toUpperCase() === ticker.toUpperCase()) {
        rawTrades.push({
          date: row[0],
          side: row[2],
          qty: parseFloat(row[3]) || 0,
          price: parseFloat(row[4]) || 0,
          exchangeComm: parseFloat(row[5]) || 0,
          brokerComm: parseFloat(row[6]) || 0,
          totalCommission: (parseFloat(row[5]) || 0) + (parseFloat(row[6]) || 0)
        });
      }
    }
    if (rawTrades.length === 0) return { ticker, error: 'Нет сделок по этому тикеру' };

    rawTrades.sort((a, b) => a.date - b.date);
    const splits = this._getSplits(ticker);

    const adjustedTrades = rawTrades.map(t => {
      let qty = t.qty;
      let price = t.price;
      for (const s of splits) {
        if (s.date > t.date) { qty *= s.ratio; price /= s.ratio; }
      }
      return { ...t, adjQty: qty, adjPrice: price };
    });

    let position = 0, totalCost = 0, totalQty = 0, totalCommission = 0;
    for (const t of adjustedTrades) {
      if (t.side === 'Покупка') {
        position += t.adjQty;
        totalCost += t.adjQty * t.adjPrice;
        totalQty += t.adjQty;
        totalCommission += t.totalCommission;
      } else if (t.side === 'Продажа') {
        position -= t.adjQty;
        if (position < 0) position = 0;
      }
    }

    if (position === 0) return { ticker, position: 0, message: 'Позиция закрыта' };

    const avgPrice = totalQty > 0 ? totalCost / totalQty : 0;
    const breakeven = avgPrice + (totalCommission / position);

    // Последняя цена (из последней сделки по тикеру)
    const lastPrice = rawTrades[rawTrades.length - 1].price;

    let recommendation = '';
    if (lastPrice > 0) {
      if (breakeven > 0 && lastPrice >= breakeven) {
        recommendation = '✅ Цена выше безубыточности. Можно продавать с прибылью.';
      } else {
        recommendation = '⏳ Цена ниже безубыточности. Дождитесь роста или усредните позицию.';
      }
    }

    return {
      ticker, position,
      avgPrice: Math.round(avgPrice * 100) / 100,
      totalCommission: Math.round(totalCommission * 100) / 100,
      breakeven: Math.round(breakeven * 100) / 100,
      lastPrice: Math.round(lastPrice * 100) / 100,
      recommendation,
      splitsApplied: splits.length
    };
  }

  showAllBreakeven() {
    const data = this.tradesSheet.getDataRange().getValues();
    const tickers = new Set();
    for (let i = 1; i < data.length; i++) { if (data[i][1]) tickers.add(data[i][1].toString().trim().toUpperCase()); }
    if (tickers.size === 0) { showMessage('Акции', 'Нет данных о сделках', 'warning'); return; }

    let report = '📈 ТОЧКИ БЕЗУБЫТОЧНОСТИ АКЦИЙ\n═══════════════════════════════\n\n';
    for (const ticker of tickers) {
      const result = this.calculateBreakeven(ticker);
      if (result.error) { report += `⚠️ ${ticker}: ${result.error}\n\n`; continue; }
      if (result.position === 0) { report += `🔹 ${ticker}: позиция закрыта\n\n`; continue; }
      report += `🔹 ${ticker}\n   Позиция: ${result.position} шт.\n   Средняя цена: ${formatMoney(result.avgPrice)}\n   Комиссии: ${formatMoney(result.totalCommission)}\n`;
      report += `   📉 Последняя цена: ${formatMoney(result.lastPrice)}\n`;
      report += `   🎯 Безубыточность: ${formatMoney(result.breakeven)}\n`;
      report += `   💡 ${result.recommendation}\n`;
      if (result.splitsApplied > 0) report += `   🔄 Применено сплитов: ${result.splitsApplied}\n`;
      report += '\n';
    }
    showMessage('Безубыточность акций', report);
  }

  showAddTradeDialog() {
    const tickers = this.getStockTickers();
    const tickerOptions = tickers.map(t => `<option value="${t}">${t}</option>`).join('');
    const html = HtmlService.createHtmlOutput(`
      <!DOCTYPE html><html><head><base target="_top">
      <style>
        body{font-family:Arial;padding:20px} input,select{width:100%;padding:8px;margin:5px 0;border:1px solid #ddd;border-radius:6px}
        button{width:100%;padding:10px;background:#1a73e8;color:white;border:none;border-radius:6px;cursor:pointer;margin-top:10px}
        .cancel{background:#f1f3f4;color:#333} .new-ticker-row{display:none;margin-top:8px}
      </style></head><body>
        <h3>📈 Добавить сделку с акцией</h3>
        <label>Дата:</label><input type="text" id="date" value="${formatDate(new Date())}">
        <label>Тикер:</label><select id="tickerSelect" onchange="handleTickerChange()">
          <option value="">-- Выберите тикер --</option>${tickerOptions}<option value="__new__">➕ Ввести новый...</option>
        </select>
        <div id="newTickerRow" class="new-ticker-row"><label>Новый тикер:</label><input type="text" id="newTicker" placeholder="Введите новый тикер"></div>
        <label>Тип:</label><select id="side"><option value="Покупка">Покупка</option><option value="Продажа">Продажа</option></select>
        <label>Количество:</label><input type="number" id="qty" step="1" value="1">
        <label>Цена (₽):</label><input type="number" id="price" step="0.01">
        <label>Комиссия биржи (₽):</label><input type="number" id="exchangeComm" step="0.01" value="0">
        <label>Комиссия брокера (₽):</label><input type="number" id="brokerComm" step="0.01" value="0">
        <button onclick="save()">💾 Сохранить</button><button class="cancel" onclick="google.script.host.close()">Отмена</button>
        <script>
          function handleTickerChange(){var s=document.getElementById('tickerSelect');document.getElementById('newTickerRow').style.display=s.value==='__new__'?'block':'none';if(s.value==='__new__')document.getElementById('newTicker').focus();}
          function save(){
            var date=document.getElementById('date').value.trim();
            var tickerSelect=document.getElementById('tickerSelect');
            var ticker=tickerSelect.value==='__new__'?document.getElementById('newTicker').value.trim().toUpperCase():tickerSelect.value;
            if(!ticker){alert('Выберите или введите тикер');return}
            var side=document.getElementById('side').value;
            var qty=parseInt(document.getElementById('qty').value);
            var price=parseFloat(document.getElementById('price').value);
            var exchangeComm=parseFloat(document.getElementById('exchangeComm').value)||0;
            var brokerComm=parseFloat(document.getElementById('brokerComm').value)||0;
            if(!date||isNaN(qty)||isNaN(price)){alert('Заполните все поля');return}
            google.script.run.withSuccessHandler(function(){alert('✅ Сделка добавлена');google.script.host.close()}).addStockTrade(date,ticker,side,qty,price,exchangeComm,brokerComm);
          }
        </script>
      </body></html>
    `).setWidth(400).setHeight(580);
    SpreadsheetApp.getUi().showModalDialog(html, 'Сделка с акцией');
  }

  showAddSplitDialog() {
    const tickers = this.getStockTickers();
    const tickerOptions = tickers.map(t => `<option value="${t}">${t}</option>`).join('');
    const html = HtmlService.createHtmlOutput(`
      <!DOCTYPE html><html><head><base target="_top">
      <style>
        body{font-family:Arial;padding:20px} input,select{width:100%;padding:8px;margin:5px 0;border:1px solid #ddd;border-radius:6px}
        button{width:100%;padding:10px;background:#1a73e8;color:white;border:none;border-radius:6px;cursor:pointer;margin-top:10px}
        .cancel{background:#f1f3f4;color:#333} .hint{font-size:12px;color:#666} .new-ticker-row{display:none;margin-top:8px}
      </style></head><body>
        <h3>🔄 Добавить сплит / консолидацию</h3>
        <div class="hint">Сплит 2:1 → коэфф. 2. Консолидация 1:5 → коэфф. 0.2</div>
        <label>Дата:</label><input type="text" id="date" value="${formatDate(new Date())}">
        <label>Тикер:</label><select id="tickerSelect" onchange="handleTickerChange()">
          <option value="">-- Выберите тикер --</option>${tickerOptions}<option value="__new__">➕ Ввести новый...</option>
        </select>
        <div id="newTickerRow" class="new-ticker-row"><label>Новый тикер:</label><input type="text" id="newTicker" placeholder="Введите новый тикер"></div>
        <label>Тип:</label><select id="type"><option value="split">Сплит</option><option value="reverse">Консолидация</option></select>
        <label>Коэффициент:</label><input type="number" id="ratio" step="0.001" placeholder="2.0">
        <button onclick="save()">💾 Сохранить</button><button class="cancel" onclick="google.script.host.close()">Отмена</button>
        <script>
          function handleTickerChange(){var s=document.getElementById('tickerSelect');document.getElementById('newTickerRow').style.display=s.value==='__new__'?'block':'none';if(s.value==='__new__')document.getElementById('newTicker').focus();}
          function save(){
            var date=document.getElementById('date').value.trim();
            var tickerSelect=document.getElementById('tickerSelect');
            var ticker=tickerSelect.value==='__new__'?document.getElementById('newTicker').value.trim().toUpperCase():tickerSelect.value;
            if(!ticker){alert('Выберите или введите тикер');return}
            var type=document.getElementById('type').value;
            var ratio=parseFloat(document.getElementById('ratio').value);
            if(!date||isNaN(ratio)||ratio<=0){alert('Заполните все поля');return}
            google.script.run.withSuccessHandler(function(){alert('✅ Событие сохранено');google.script.host.close()}).addStockSplit(date,ticker,type,ratio);
          }
        </script>
      </body></html>
    `).setWidth(400).setHeight(450);
    SpreadsheetApp.getUi().showModalDialog(html, 'Сплит / Консолидация');
  }
}

// Глобальные функции
function showStocksBreakeven() { new StocksBreakEven().showAllBreakeven(); }
function showAddStockTrade() { new StocksBreakEven().showAddTradeDialog(); }
function showAddStockSplit() { new StocksBreakEven().showAddSplitDialog(); }
function addStockTrade(dateStr, ticker, side, qty, price, exchangeComm, brokerComm) {
  const date = safeParseDate(dateStr); if (!date) return false;
  return new StocksBreakEven().addTrade(date, ticker, side, qty, price, exchangeComm, brokerComm);
}
function addStockSplit(dateStr, ticker, type, ratio) {
  const date = safeParseDate(dateStr); if (!date) return false;
  return new StocksBreakEven().addSplit(date, ticker, type, ratio);
}
function getStockTickers() { return new StocksBreakEven().getStockTickers(); }