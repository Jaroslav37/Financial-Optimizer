/**
 * КАЛЬКУЛЯТОР ОБЛИГАЦИЙ - расчёт доходности облигаций
 * 
 * Поддерживает:
 * - Автоматическую загрузку данных через MOEX ISS API
 * - Ручной ввод параметров
 * - Расчёт текущей, простой и эффективной доходности (YTM)
 * - Расчёт дюрации и НКД
 * - Учёт комиссии брокера
 * - Создание IMPORTXML формул для автоматического обновления
 * 
 * Исправлено: безопасное логирование, удалены зависимости от инвестиционного парсера
 */

class BondCalculator {
  constructor() {
    this.sheet = getSheetSafe(SHEETS.BONDS);
    this.currentYear = CONFIG.CURRENT_YEAR;
    this.dataLoader = new BondDataLoader();
  }

  // =============================================
  // ИНИЦИАЛИЗАЦИЯ ЛИСТА ОБЛИГАЦИЙ
  // =============================================

  initBondSheet() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEETS.BONDS);
      
      if (sheet) {
        var ui = SpreadsheetApp.getUi();
        var response = ui.alert(
          'Инициализация',
          'Пересоздать лист "Облигации"? Все данные будут потеряны.',
          ui.ButtonSet.YES_NO
        );
        if (response !== ui.Button.YES) return;
        ss.deleteSheet(sheet);
      }
      
      sheet = ss.insertSheet(SHEETS.BONDS);
      sheet.setTabColor('#f9cb9c');
      
      // Заголовок
      sheet.getRange(1, 1).setValue('📊 КАЛЬКУЛЯТОР ДОХОДНОСТИ ОБЛИГАЦИЙ').setFontWeight('bold').setFontSize(16);
      sheet.getRange(1, 1, 1, 8).merge();
      sheet.getRange(1, 1).setBackground('#e6f4ea');
      
      // Инструкция
      sheet.getRange(3, 1).setValue('🚀 КАК ПОЛЬЗОВАТЬСЯ:').setFontWeight('bold');
      sheet.getRange(4, 1).setValue('1️⃣ Введите код облигации (ISIN)');
      sheet.getRange(5, 1).setValue('2️⃣ Выберите тип расчёта (Текущая / Простая / Эффективная)');
      sheet.getRange(6, 1).setValue('3️⃣ Укажите комиссию брокера (по умолчанию 0.05%)');
      sheet.getRange(7, 1).setValue('4️⃣ Укажите цену покупки (или используйте рыночную)');
      sheet.getRange(8, 1).setValue('5️⃣ Нажмите "Рассчитать доходность"');
      
      // Параметры
      sheet.getRange(10, 1).setValue('📝 ПАРАМЕТРЫ ОБЛИГАЦИИ').setFontWeight('bold').setFontSize(12);
      sheet.getRange(10, 1, 1, 6).merge();
      sheet.getRange(10, 1).setBackground('#f0f0f0');
      
      sheet.getRange(11, 1).setValue('Код облигации (ISIN):').setFontWeight('bold');
      sheet.getRange(11, 2).setValue('').setBackground('#fff9c4');
      sheet.getRange(11, 2).setBorder(true, true, true, true, false, false);
      
      sheet.getRange(12, 1).setValue('Тип расчёта:').setFontWeight('bold');
      var types = ['Текущая доходность', 'Простая доходность', 'Эффективная доходность (YTM)'];
      sheet.getRange(12, 2).setDataValidation(
        SpreadsheetApp.newDataValidation()
          .requireValueInList(types)
          .build()
      );
      sheet.getRange(12, 2).setValue('Эффективная доходность (YTM)');
      
      sheet.getRange(13, 1).setValue('Комиссия брокера (%):').setFontWeight('bold');
      sheet.getRange(13, 2).setValue(0.05).setNumberFormat('0.00%');
      
      sheet.getRange(14, 1).setValue('Цена покупки (% от номинала):').setFontWeight('bold');
      sheet.getRange(14, 2).setValue(100).setNumberFormat('#,##0.00');
      
      sheet.getRange(16, 1).setValue('🔍 РАССЧИТАТЬ ДОХОДНОСТЬ').setFontWeight('bold').setBackground('#1a73e8').setFontColor('#ffffff');
      sheet.getRange(16, 1, 1, 2).merge();
      sheet.getRange(16, 1).setHorizontalAlignment('center');
      
      // Данные по облигации
      sheet.getRange(18, 1).setValue('📊 ДАННЫЕ ПО ОБЛИГАЦИИ').setFontWeight('bold').setFontSize(12);
      sheet.getRange(18, 1, 1, 6).merge();
      sheet.getRange(18, 1).setBackground('#f0f0f0');
      
      var dataHeaders = [
        ['Название', '', 'Номинал', '', 'Тип', ''],
        ['ISIN', '', 'Рыночная цена', '', 'Купон %', ''],
        ['Выпуск', '', 'НКД', '', 'Частота', ''],
        ['Погашение', '', 'Следующий купон', '', 'Дней до погаш.', '']
      ];
      sheet.getRange(20, 1, 4, 6).setValues(dataHeaders);
      sheet.getRange(20, 1, 4, 1).setFontWeight('bold');
      sheet.getRange(20, 3, 4, 1).setFontWeight('bold');
      sheet.getRange(20, 5, 4, 1).setFontWeight('bold');
      
      // Результаты
      sheet.getRange(26, 1).setValue('📈 РЕЗУЛЬТАТЫ РАСЧЁТА').setFontWeight('bold').setFontSize(12);
      sheet.getRange(26, 1, 1, 6).merge();
      sheet.getRange(26, 1).setBackground('#f0f0f0');
      
      var resultsHeaders = [
        ['Текущая доходность', '0.00%', 'Купонный доход (год)', '0.00 ₽', '', ''],
        ['Простая доходность', '0.00%', 'Накопленный доход', '0.00 ₽', '', ''],
        ['Эффективная доходность (YTM)', '0.00%', 'Доход к погашению', '0.00 ₽', '', ''],
        ['Дюрация (лет)', '0.00', 'С учетом комиссии', '0.00%', '', ''],
        ['НКД к уплате', '0.00 ₽', 'Итого к оплате', '0.00 ₽', '', '']
      ];
      sheet.getRange(28, 1, 5, 6).setValues(resultsHeaders);
      sheet.getRange(28, 1, 5, 1).setFontWeight('bold');
      sheet.getRange(28, 3, 5, 1).setFontWeight('bold');
      
      // Форматирование
      sheet.setColumnWidth(1, 200);
      sheet.setColumnWidth(2, 150);
      sheet.setColumnWidth(3, 180);
      sheet.setColumnWidth(4, 120);
      sheet.setColumnWidth(5, 150);
      sheet.setColumnWidth(6, 120);
      
      sheet.getRange(13, 2).setNumberFormat('0.00%');
      sheet.getRange(14, 2).setNumberFormat('#,##0.00');
      sheet.getRange(28, 2, 3, 1).setNumberFormat('0.00%');
      sheet.getRange(31, 2, 1, 1).setNumberFormat('0.00%');
      sheet.getRange(28, 4, 5, 1).setNumberFormat('#,##0.00 ₽');
      
      // Реестр облигаций
      sheet.getRange(36, 1).setValue('📋 РЕЕСТР ОБЛИГАЦИЙ').setFontWeight('bold').setFontSize(14);
      sheet.getRange(36, 1, 1, 8).merge();
      sheet.getRange(36, 1).setBackground('#e6f4ea');
      
      var registryHeaders = [
        ['Тикер/ISIN', 'Название', 'Номинал', 'Купон %', 'Частота', 
         'Дата погашения', 'Рыночная цена', 'Тип']
      ];
      sheet.getRange(38, 1, 1, 8).setValues(registryHeaders);
      sheet.getRange(38, 1, 1, 8).setFontWeight('bold').setBackground('#f0f0f0');
      
      sheet.setColumnWidth(1, 150);
      sheet.setColumnWidth(2, 250);
      sheet.setColumnWidth(3, 100);
      sheet.setColumnWidth(4, 100);
      sheet.setColumnWidth(5, 100);
      sheet.setColumnWidth(6, 120);
      sheet.setColumnWidth(7, 120);
      sheet.setColumnWidth(8, 100);
      
      showMessage('Готово', '✅ Лист "Облигации" создан', 'info');
      addLogEntry('Лист "Облигации" создан', 'INFO', 'BondCalculator');
      
    } catch (e) {
      logError(e, 'BondCalculator.initBondSheet');
      showMessage('Ошибка', e.toString(), 'error');
    }
  }

  // =============================================
  // ПОЛУЧЕНИЕ ДАННЫХ ИЗ MOEX ISS API
  // =============================================

  fetchBondData(isin) {
    if (!isin || typeof isin !== 'string') {
      addLogEntry('ISIN is empty or invalid', 'WARNING', 'BondCalculator');
      return null;
    }
    
    var isinPattern = /^[A-Z]{2}[A-Z0-9]{10}$/;
    if (!isinPattern.test(isin.toUpperCase())) {
      addLogEntry('Invalid ISIN format: ' + isin, 'WARNING', 'BondCalculator');
      return null;
    }
    
    var upperIsin = isin.toUpperCase();
    
    try {
      var board = this.isOfz(upperIsin) ? 'TQOB' : 'TQCB';
      var bondData = this.dataLoader.getFullBondData(upperIsin, board);
      
      if (!bondData) {
        bondData = this.fetchBondDataDirect(upperIsin, board);
      }
      
      return bondData;
      
    } catch (e) {
      logError(e, 'BondCalculator.fetchBondData');
      return null;
    }
  }

  fetchBondDataDirect(isin, board) {
    try {
      var url = 'https://iss.moex.com/iss/engines/stock/markets/bonds/boards/' + board + '/securities.xml?iss.meta=off&iss.only=securities&securities.columns=SECID,SECNAME,FACEVALUE,PREVADMITTEDQUOTE,COUPONVALUE,COUPONPERCENT,MATDATE,COUPONFREQUENCY';
      
      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' }
      });
      
      var xml = response.getContentText();
      var regex = new RegExp('<row[^>]*SECID=["\']' + isin + '["\'][^>]*>', 'i');
      var match = xml.match(regex);
      
      if (!match) return null;
      
      var row = match[0];
      
      return {
        secid: isin,
        name: this.extractAttribute(row, 'SECNAME'),
        faceValue: parseFloat(this.extractAttribute(row, 'FACEVALUE') || 1000),
        marketPrice: parseFloat(this.extractAttribute(row, 'PREVADMITTEDQUOTE') || 0),
        couponValue: parseFloat(this.extractAttribute(row, 'COUPONVALUE') || 0),
        couponRate: parseFloat(this.extractAttribute(row, 'COUPONPERCENT') || 0) / 100,
        maturityDate: this.extractAttribute(row, 'MATDATE') || '',
        couponFrequency: parseInt(this.extractAttribute(row, 'COUPONFREQUENCY') || 2),
        accruedInterest: 0
      };
      
    } catch (e) {
      logError(e, 'BondCalculator.fetchBondDataDirect');
      return null;
    }
  }

  fetchAccruedInterest(isin) {
    try {
      var board = this.isOfz(isin) ? 'TQOB' : 'TQCB';
      return this.dataLoader.getAccruedInterest(isin, board);
    } catch (e) {
      logError(e, 'BondCalculator.fetchAccruedInterest');
      return 0;
    }
  }

  isOfz(secid) {
    var upper = secid.toUpperCase();
    return upper.indexOf('SU') === 0 || 
           upper.indexOf('OFZ') !== -1 ||
           upper.indexOf('ОФЗ') !== -1;
  }

  extractAttribute(xmlString, attrName) {
    var regex = new RegExp(attrName + '=["\']([^"\']*)["\']', 'i');
    var match = xmlString.match(regex);
    return match ? match[1] : '';
  }

  // =============================================
  // РУЧНОЙ ВВОД ПАРАМЕТРОВ
  // =============================================

  manualBondInput() {
    var ui = SpreadsheetApp.getUi();
    
    var html = HtmlService.createHtmlOutput('<!DOCTYPE html><html><head><base target="_top"><style>body{font-family:"Segoe UI",Arial,sans-serif;padding:20px;background:#f5f5f5}.container{max-width:500px;margin:0 auto;background:white;padding:25px;border-radius:12px;box-shadow:0 2px 10px rgba(0,0,0,0.1)}h2{color:#1a73e8;margin-top:0}.form-group{margin-bottom:15px}label{display:block;margin-bottom:5px;font-weight:500;color:#333}input,select{width:100%;padding:10px;border:1px solid #ddd;border-radius:6px;font-size:14px;box-sizing:border-box}button{background:#1a73e8;color:white;padding:12px;border:none;border-radius:6px;cursor:pointer;width:100%;font-size:16px;margin-top:10px}button:hover{background:#1557b0}.note{font-size:12px;color:#666;margin-top:5px}</style></head><body><div class="container"><h2>📝 Ручной ввод параметров</h2><p>Введите параметры облигации вручную:</p><div class="form-group"><label>Название:</label><input type="text" id="name" placeholder="ОФЗ 26239"></div><div class="form-group"><label>ISIN:</label><input type="text" id="isin" placeholder="SU26239RMFS5"></div><div class="form-group"><label>Дата погашения (ДД.ММ.ГГГГ):</label><input type="text" id="maturityDate" placeholder="15.05.2036"></div><div class="form-group"><label>Купонная ставка (%):</label><input type="number" id="couponRate" step="0.01" placeholder="7.5"></div><div class="form-group"><label>Номинал (₽):</label><input type="number" id="faceValue" step="100" value="1000"></div><div class="form-group"><label>Частота выплат (раз в год):</label><select id="couponFrequency"><option value="1">1 раз в год</option><option value="2" selected>2 раза в год (раз в полгода)</option><option value="4">4 раза в год (раз в квартал)</option><option value="12">12 раз в год (ежемесячно)</option></select></div><div class="form-group"><label>Тип облигации:</label><select id="bondType"><option value="corp">Корпоративная</option><option value="ofz">ОФЗ</option><option value="mun">Муниципальная</option></select></div><div class="form-group"><label>Текущая рыночная цена (% от номинала):</label><input type="number" id="marketPrice" step="0.01" placeholder="98.5"></div><div class="form-group"><label>НКД (₽):</label><input type="number" id="accruedInterest" step="0.01" placeholder="15.23"></div><button onclick="submitManualData()">💾 Сохранить параметры</button></div><script>function submitManualData(){var data={name:document.getElementById("name").value,isin:document.getElementById("isin").value,maturityDate:document.getElementById("maturityDate").value,couponRate:parseFloat(document.getElementById("couponRate").value),faceValue:parseFloat(document.getElementById("faceValue").value),couponFrequency:parseInt(document.getElementById("couponFrequency").value),bondType:document.getElementById("bondType").value,marketPrice:parseFloat(document.getElementById("marketPrice").value),accruedInterest:parseFloat(document.getElementById("accruedInterest").value)};if(!data.name){alert("Введите название");return}if(!data.maturityDate){alert("Введите дату погашения");return}if(isNaN(data.couponRate)||data.couponRate<=0){alert("Введите корректную купонную ставку");return}google.script.run.withSuccessHandler(function(){alert("✅ Параметры сохранены!");google.script.host.close()}).withFailureHandler(function(error){alert("❌ Ошибка: "+error)}).saveManualBondData(data)}</script></body></html>')
      .setWidth(550)
      .setHeight(700)
      .setTitle('Ручной ввод облигации');
    
    ui.showModalDialog(html, 'Ручной ввод параметров');
  }

  saveManualBondData(data) {
    try {
      var sheet = getSheetSafe(SHEETS.BONDS);
      if (!sheet) {
        this.initBondSheet();
        return;
      }
      
      var lastRow = sheet.getLastRow();
      var registryStartRow = 39;
      var targetRow = registryStartRow;
      
      for (var i = registryStartRow; i <= lastRow + 1; i++) {
        if (sheet.getRange(i, 1).isBlank()) {
          targetRow = i;
          break;
        }
      }
      
      sheet.getRange(targetRow, 1).setValue(data.isin);
      sheet.getRange(targetRow, 2).setValue(data.name);
      sheet.getRange(targetRow, 3).setValue(data.faceValue);
      sheet.getRange(targetRow, 4).setValue(data.couponRate);
      sheet.getRange(targetRow, 5).setValue(data.couponFrequency);
      sheet.getRange(targetRow, 6).setValue(data.maturityDate);
      sheet.getRange(targetRow, 7).setValue(data.marketPrice);
      sheet.getRange(targetRow, 8).setValue(data.bondType === 'ofz' ? 'ОФЗ' : 'Корпоративная');
      
      sheet.getRange(20, 2).setValue(data.name);
      sheet.getRange(21, 2).setValue(data.isin);
      sheet.getRange(22, 2).setValue('');
      sheet.getRange(23, 2).setValue(data.maturityDate);
      sheet.getRange(24, 2).setValue(data.couponRate + '%');
      sheet.getRange(25, 2).setValue(data.couponFrequency + ' раз(а)');
      sheet.getRange(20, 4).setValue(data.faceValue);
      sheet.getRange(21, 4).setValue(data.marketPrice);
      sheet.getRange(22, 4).setValue(data.accruedInterest);
      sheet.getRange(23, 4).setValue('');
      sheet.getRange(24, 4).setValue((data.faceValue * data.couponRate / 100).toFixed(2));
      sheet.getRange(25, 4).setValue('');
      
      var props = PropertiesService.getScriptProperties();
      props.setProperty('MANUAL_BOND_DATA', JSON.stringify(data));
      
      showMessage('Успешно', 'Параметры сохранены', 'info');
      addLogEntry('Ручные параметры сохранены для ISIN: ' + data.isin, 'INFO', 'BondCalculator');
      
    } catch (e) {
      logError(e, 'BondCalculator.saveManualBondData');
      throw new Error(e.toString());
    }
  }

  // =============================================
  // РАСЧЁТЫ ДОХОДНОСТИ
  // =============================================

  calculateCurrentYield(couponRate, faceValue, pricePercent) {
    var annualCoupon = couponRate * faceValue;
    var priceRub = (pricePercent / 100) * faceValue;
    return (annualCoupon / priceRub) * 100;
  }

  calculateSimpleYield(couponRate, faceValue, pricePercent, yearsToMaturity) {
    if (yearsToMaturity <= 0) return 0;
    var annualCoupon = couponRate * faceValue;
    var priceRub = (pricePercent / 100) * faceValue;
    var discount = (faceValue - priceRub) / yearsToMaturity;
    var annualIncome = annualCoupon + discount;
    return (annualIncome / priceRub) * 100;
  }

  calculateXIRR(cashFlows, guess) {
    var xirr = guess || 0.1;
    var iteration = 0;
    var maxIterations = 100;
    var tolerance = 0.00001;
    
    if (cashFlows.length < 2) return 0;
    
    cashFlows.sort(function(a, b) { return a.date - b.date; });
    var firstDate = cashFlows[0].date;
    
    while (iteration < maxIterations) {
      var npv = 0;
      var derivative = 0;
      
      for (var i = 0; i < cashFlows.length; i++) {
        var flow = cashFlows[i];
        var daysDiff = (flow.date - firstDate) / (1000 * 60 * 60 * 24);
        var years = daysDiff / 365;
        var factor = Math.pow(1 + xirr, years);
        
        if (factor === 0) continue;
        npv += flow.amount / factor;
        derivative += -years * flow.amount / (factor * (1 + xirr));
      }
      
      if (derivative === 0) break;
      
      var newXirr = xirr - npv / derivative;
      if (Math.abs(newXirr - xirr) < tolerance) {
        return newXirr * 100;
      }
      xirr = newXirr;
      iteration++;
    }
    
    return xirr * 100;
  }

  calculateYTMCashFlows(purchaseDate, maturityDate, couponRate, faceValue, couponFrequency, pricePercent, commission, accruedInterest) {
    var cashFlows = [];
    var priceRub = (pricePercent / 100) * faceValue;
    var totalCost = priceRub + accruedInterest + (priceRub * commission / 100);
    cashFlows.push({ date: purchaseDate, amount: -totalCost });
    
    var currentDate = new Date(purchaseDate);
    var monthsStep = 12 / couponFrequency;
    var couponAmount = (couponRate * faceValue) / couponFrequency;
    
    var nextDate = new Date(currentDate);
    var maxIterations = 100;
    var iter = 0;
    
    while (iter < maxIterations) {
      nextDate = new Date(currentDate);
      nextDate.setMonth(nextDate.getMonth() + monthsStep);
      
      if (nextDate >= maturityDate) break;
      
      cashFlows.push({ date: new Date(nextDate), amount: couponAmount });
      currentDate = nextDate;
      iter++;
    }
    
    cashFlows.push({ date: maturityDate, amount: faceValue });
    
    return this.calculateXIRR(cashFlows);
  }

  calculateDuration(cashFlows, ytm) {
    var weightedSum = 0;
    var totalPV = 0;
    var firstDate = cashFlows[0].date;
    
    for (var i = 0; i < cashFlows.length; i++) {
      var flow = cashFlows[i];
      var daysDiff = (flow.date - firstDate) / (1000 * 60 * 60 * 24);
      var years = daysDiff / 365;
      var pv = flow.amount / Math.pow(1 + ytm / 100, years);
      weightedSum += years * pv;
      totalPV += pv;
    }
    
    return totalPV > 0 ? weightedSum / totalPV : 0;
  }

  // =============================================
  // ОСНОВНОЙ РАСЧЁТ
  // =============================================

  calculateBondYield() {
    try {
      var ui = SpreadsheetApp.getUi();
      
      var isin = this.sheet.getRange(11, 2).getValue();
      var calculationType = this.sheet.getRange(12, 2).getValue();
      var commission = this.sheet.getRange(13, 2).getValue() / 100;
      var pricePercent = this.sheet.getRange(14, 2).getValue();
      
      if (!isin) {
        ui.alert('Ошибка', 'Введите код облигации (ISIN)', ui.ButtonSet.OK);
        return;
      }
      
      showMessage('Информация', 'Загрузка данных об облигации...', 'info');
      
      var bondData = this.fetchBondData(isin);
      
      if (!bondData) {
        var props = PropertiesService.getScriptProperties();
        var manualData = props.getProperty('MANUAL_BOND_DATA');
        
        if (manualData) {
          var parsed = JSON.parse(manualData);
          if (parsed.isin === isin) {
            bondData = {
              name: parsed.name,
              isin: parsed.isin,
              maturityDate: parsed.maturityDate,
              couponRate: parsed.couponRate / 100,
              couponFrequency: parsed.couponFrequency,
              faceValue: parsed.faceValue,
              marketPrice: parsed.marketPrice,
              accruedInterest: parsed.accruedInterest
            };
          }
        }
      }
      
      if (!bondData) {
        var response = ui.alert(
          'Облигация не найдена',
          'Не удалось загрузить данные для ISIN: ' + isin + '\n\nХотите ввести параметры вручную?',
          ui.ButtonSet.YES_NO
        );
        
        if (response === ui.Button.YES) {
          this.manualBondInput();
        }
        return;
      }
      
      var accruedInterest = bondData.accruedInterest || this.fetchAccruedInterest(isin);
      
      this.sheet.getRange(20, 2).setValue(bondData.name);
      this.sheet.getRange(21, 2).setValue(bondData.isin);
      this.sheet.getRange(22, 2).setValue(bondData.issueDate || '');
      this.sheet.getRange(23, 2).setValue(bondData.maturityDate);
      this.sheet.getRange(24, 2).setValue((bondData.couponRate * 100).toFixed(2) + '%');
      this.sheet.getRange(25, 2).setValue((bondData.couponFrequency || 2) + ' раз(а) в год');
      this.sheet.getRange(20, 4).setValue(bondData.faceValue);
      this.sheet.getRange(21, 4).setValue(bondData.marketPrice);
      this.sheet.getRange(22, 4).setValue(accruedInterest);
      this.sheet.getRange(23, 4).setValue('');
      this.sheet.getRange(24, 4).setValue((bondData.faceValue * bondData.couponRate).toFixed(2));
      this.sheet.getRange(25, 4).setValue('');
      
      var today = new Date();
      var maturityDate = new Date(bondData.maturityDate);
      var yearsToMaturity = (maturityDate - today) / (1000 * 60 * 60 * 24 * 365);
      var couponRate = bondData.couponRate;
      var faceValue = bondData.faceValue;
      var couponFrequency = bondData.couponFrequency || 2;
      
      var currentYield = this.calculateCurrentYield(couponRate, faceValue, pricePercent);
      var simpleYield = this.calculateSimpleYield(couponRate, faceValue, pricePercent, yearsToMaturity);
      var ytm = this.calculateYTMCashFlows(
        today, maturityDate, couponRate, faceValue,
        couponFrequency, pricePercent, commission, accruedInterest
      );
      
      var ytmWithCommission = ytm - (commission * 100);
      var nkdToPay = accruedInterest;
      var totalToPay = (pricePercent / 100) * faceValue + nkdToPay + (pricePercent / 100) * faceValue * commission / 100;
      
      this.sheet.getRange(28, 2).setValue((currentYield / 100).toFixed(2) + '%');
      this.sheet.getRange(28, 4).setValue((couponRate * faceValue).toFixed(2));
      this.sheet.getRange(29, 2).setValue((simpleYield / 100).toFixed(2) + '%');
      this.sheet.getRange(30, 2).setValue((ytm / 100).toFixed(2) + '%');
      this.sheet.getRange(30, 4).setValue(((faceValue - (pricePercent / 100) * faceValue) + (couponRate * faceValue * yearsToMaturity)).toFixed(2));
      this.sheet.getRange(31, 2).setValue((ytmWithCommission / 100).toFixed(2) + '%');
      this.sheet.getRange(31, 4).setValue(ytmWithCommission.toFixed(2) + '%');
      this.sheet.getRange(32, 2).setValue(nkdToPay.toFixed(2));
      this.sheet.getRange(32, 4).setValue(totalToPay.toFixed(2));
      
      var cashFlows = [];
      cashFlows.push({ date: today, amount: -((pricePercent / 100) * faceValue) });
      
      var currentDate = new Date(today);
      var monthsStep = 12 / couponFrequency;
      var couponAmount = (couponRate * faceValue) / couponFrequency;
      
      var nextDate = new Date(currentDate);
      var maxIterations = 100;
      var iter = 0;
      
      while (iter < maxIterations) {
        nextDate = new Date(currentDate);
        nextDate.setMonth(nextDate.getMonth() + monthsStep);
        if (nextDate >= maturityDate) break;
        cashFlows.push({ date: new Date(nextDate), amount: couponAmount });
        currentDate = nextDate;
        iter++;
      }
      cashFlows.push({ date: maturityDate, amount: faceValue });
      
      var duration = this.calculateDuration(cashFlows, ytm);
      this.sheet.getRange(31, 4).setValue(duration.toFixed(2) + ' лет');
      
      var message = '📊 РЕЗУЛЬТАТ РАСЧЁТА\n\n';
      message += 'Облигация: ' + bondData.name + '\n';
      message += 'ISIN: ' + isin + '\n\n';
      message += '💰 Текущая доходность: ' + (currentYield / 100).toFixed(2) + '%\n';
      message += '📈 Простая доходность: ' + (simpleYield / 100).toFixed(2) + '%\n';
      message += '⭐ Эффективная доходность (YTM): ' + (ytm / 100).toFixed(2) + '%\n';
      message += '💳 С учётом комиссии: ' + ytmWithCommission.toFixed(2) + '%\n';
      message += '📅 Дюрация: ' + duration.toFixed(2) + ' лет\n\n';
      message += '💰 НКД к уплате: ' + formatMoney(nkdToPay) + '\n';
      message += '💵 Итого к оплате: ' + formatMoney(totalToPay);
      
      showMessage('Расчёт завершён', message);
      addLogEntry('Расчёт доходности выполнен для ISIN: ' + isin, 'INFO', 'BondCalculator');
      
    } catch (e) {
      logError(e, 'BondCalculator.calculateBondYield');
      showMessage('Ошибка', e.toString(), 'error');
    }
  }

  // =============================================
  // ГЕНЕРАЦИЯ IMPORTXML ФОРМУЛ
  // =============================================

  generateImportFormulas() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var sheet = ss.getSheetByName(SHEETS.BONDS);
      
      if (!sheet) {
        this.initBondSheet();
        sheet = ss.getSheetByName(SHEETS.BONDS);
      }
      
      var data = sheet.getDataRange().getValues();
      var lastRowWithIsin = 38;
      
      for (var i = 38; i < data.length; i++) {
        if (data[i][0]) {
          lastRowWithIsin = i + 1;
        }
      }
      
      if (lastRowWithIsin <= 38) {
        showMessage('Информация', 'Нет облигаций для создания формул. Добавьте ISIN в колонку A реестра.', 'warning');
        return;
      }
      
      var formulasCreated = 0;
      
      for (var row = 39; row <= lastRowWithIsin; row++) {
        var secid = sheet.getRange(row, 1).getValue();
        if (!secid) continue;
        
        var board = this.isOfz(secid) ? 'TQOB' : 'TQCB';
        
        try {
          var nameFormula = '=IF(ISBLANK(A' + row + '), " ", IMPORTXML("https://iss.moex.com/iss/engines/stock/markets/bonds/boards/' + board + '/securities.xml?iss.meta=off&iss.only=securities&securities.columns=SECID,SECNAME", "//row[@SECID=\'" & A' + row + ' & "\']/@SECNAME"))';
          var faceFormula = '=IF(ISBLANK(A' + row + '), " ", SUBSTITUTE(IMPORTXML("https://iss.moex.com/iss/engines/stock/markets/bonds/boards/' + board + '/securities.xml?iss.meta=off&iss.only=securities&securities.columns=SECID,FACEVALUE", "//row[@SECID=\'" & A' + row + ' & "\']/@FACEVALUE"), ".", ","))';
          var couponFormula = '=IF(ISBLANK(A' + row + '), " ", SUBSTITUTE(IMPORTXML("https://iss.moex.com/iss/engines/stock/markets/bonds/boards/' + board + '/securities.xml?iss.meta=off&iss.only=securities&securities.columns=SECID,COUPONPERCENT", "//row[@SECID=\'" & A' + row + ' & "\']/@COUPONPERCENT"), ".", ","))';
          var priceFormula = '=IF(ISBLANK(A' + row + '), " ", SUBSTITUTE(IMPORTXML("https://iss.moex.com/iss/engines/stock/markets/bonds/boards/' + board + '/securities.xml?iss.meta=off&iss.only=securities&securities.columns=SECID,PREVADMITTEDQUOTE", "//row[@SECID=\'" & A' + row + ' & "\']/@PREVADMITTEDQUOTE"), ".", ","))';
          var maturityFormula = '=IF(ISBLANK(A' + row + '), " ", IMPORTXML("https://iss.moex.com/iss/engines/stock/markets/bonds/boards/' + board + '/securities.xml?iss.meta=off&iss.only=securities&securities.columns=SECID,MATDATE", "//row[@SECID=\'" & A' + row + ' & "\']/@MATDATE"))';
          
          sheet.getRange(row, 2).setFormula(nameFormula);
          sheet.getRange(row, 3).setFormula(faceFormula);
          sheet.getRange(row, 4).setFormula(couponFormula);
          sheet.getRange(row, 5).setFormula('');
          sheet.getRange(row, 6).setFormula(maturityFormula);
          sheet.getRange(row, 7).setFormula(priceFormula);
          
          formulasCreated++;
          
        } catch (e) {
          logError(e, 'BondCalculator.generateImportFormulas - ' + secid);
        }
      }
      
      showMessage('Готово', '✅ Созданы IMPORTXML формулы для ' + formulasCreated + ' облигаций\n\n⚠️ Обновление данных может занять некоторое время', 'info');
      addLogEntry('Созданы IMPORTXML формулы для ' + formulasCreated + ' облигаций', 'INFO', 'BondCalculator');
      
    } catch (e) {
      logError(e, 'BondCalculator.generateImportFormulas');
      showMessage('Ошибка', e.toString(), 'error');
    }
  }

  // =============================================
  // ТЕСТИРОВАНИЕ
  // =============================================

  testIsin(isin) {
    var result = this.fetchBondData(isin);
    
    if (result) {
      var message = 
        '✅ Облигация найдена!\n\n' +
        '📌 Название: ' + result.name + '\n' +
        '🔢 ISIN: ' + result.isin + '\n' +
        '💰 Номинал: ' + formatMoney(result.faceValue) + '\n' +
        '📅 Погашение: ' + result.maturityDate + '\n' +
        '💵 Купон: ' + (result.couponRate * 100).toFixed(2) + '%\n' +
        '📊 Частота: ' + result.couponFrequency + ' раз(а) в год\n' +
        '📈 Рыночная цена: ' + formatMoney(result.marketPrice) + '\n' +
        '📉 НКД: ' + formatMoney(result.accruedInterest);
      
      showMessage('ISIN найден', message, 'info');
    } else {
      showMessage('ISIN не найден', 
        '❌ Облигация с кодом ' + isin + ' не найдена.\n\n' +
        'Проверьте правильность ISIN.\n' +
        'Формат ISIN: 2 буквы + 10 символов\n' +
        'Пример: SU26239RMFS5 (ОФЗ)\n' +
        '       RU000A0JXQJ3 (Сбербанк)',
        'warning'
      );
    }
  }

  generateCouponReport(monthsAhead) {
    return '💵 ПЛАН КУПОННЫХ ВЫПЛАТ\n═══════════════════════\n\nФункция в разработке. Будет доступна в следующей версии.';
  }
}

// =============================================
// ГЛОБАЛЬНЫЕ ФУНКЦИИ ДЛЯ ВЫЗОВА ИЗ МЕНЮ
// =============================================

function initBondSheet() {
  new BondCalculator().initBondSheet();
}

function calculateBondYield() {
  new BondCalculator().calculateBondYield();
}

function manualBondInput() {
  new BondCalculator().manualBondInput();
}

function testIsin() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    '🔍 Проверка ISIN',
    'Введите ISIN облигации (например: SU26239RMFS5):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    var isin = response.getResponseText().trim().toUpperCase();
    new BondCalculator().testIsin(isin);
  }
}

// Функция analyzeBondPortfolio() УДАЛЕНА – зависела от collectInvestmentData()

function showCouponSchedule() {
  var calculator = new BondCalculator();
  var report = calculator.generateCouponReport(12);
  showMessage('Купонный график', report);
}

function generateImportFormulas() {
  new BondCalculator().generateImportFormulas();
}

function showBondHelp() {
  var ui = SpreadsheetApp.getUi();
  
  var helpText = 
'📈 СПРАВКА ПО ДОХОДНОСТЯМ ОБЛИГАЦИЙ\n' +
'═══════════════════════════════\n\n' +
'📌 ТЕКУЩАЯ ДОХОДНОСТЬ (Current Yield)\n' +
'Показывает только купонный доход относительно цены.\n' +
'Формула: (Годовой купон / Цена покупки) × 100%\n\n' +
'📌 ПРОСТАЯ ДОХОДНОСТЬ К ПОГАШЕНИЮ\n' +
'Учитывает купонный доход и разницу между ценой и номиналом.\n' +
'Формула: (Купон + (Номинал - Цена)/Лет) / Цена × 100%\n\n' +
'📌 ЭФФЕКТИВНАЯ ДОХОДНОСТЬ (YTM)\n' +
'Наиболее точный показатель. Учитывает реинвестирование купонов.\n' +
'Рассчитывается по формуле ЧИСТВНДОХ (XIRR)\n\n' +
'📌 ДЮРАЦИЯ\n' +
'Средневзвешенный срок возврата инвестиций.\n' +
'Показывает чувствительность цены к изменению ставок.\n\n' +
'📌 НКД (Накопленный купонный доход)\n' +
'Часть купона, накопленная с момента последней выплаты.\n' +
'При покупке вы платите продавцу НКД.\n\n' +
'💡 Совет: Для сравнения с депозитами используйте YTM.\n' +
'   Для портфеля из нескольких облигаций рассчитывайте средневзвешенную YTM.\n\n' +
'📋 Тестовые ISIN:\n' +
'• SU26239RMFS5 (ОФЗ 26239)\n' +
'• RU000A0JXQJ3 (Сбербанк БО)\n' +
'• RU000A0JWSK7 (Газпром БО)\n\n' +
'✨ Новое: При ручном вводе вы можете задать параметры облигации самостоятельно.';
  
  ui.alert('Справка по доходностям', helpText, ui.ButtonSet.OK);
}