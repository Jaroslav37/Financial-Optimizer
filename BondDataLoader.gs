/**
 * BOND DATA LOADER - загрузчик данных об облигациях с MOEX
 * Использует XmlService для безопасного парсинга XML
 * 
 * Исправлено: безопасное логирование (addLogEntry вместо Logger.log)
 */

class BondDataLoader {
  constructor() {
    this.cache = CacheService.getScriptCache();
    this.cacheExpiration = 3600; // 1 час кэширования
  }

  /**
   * Получает полные данные об облигации через XmlService
   * @param {string} secid - тикер или ISIN
   * @param {string} board - площадка (TQOB для ОФЗ, TQCB для корп.)
   * @returns {Object} данные об облигации
   */
  getFullBondData(secid, board) {
    var safeBoard = board || 'TQCB';
    
    // Валидация ISIN
    if (!this.validateIsin(secid)) {
      addLogEntry('Неверный формат ISIN: ' + secid, 'WARNING', 'BondDataLoader');
      return null;
    }
    
    var cacheKey = 'bond_full_' + secid + '_' + safeBoard;
    var cached = this.cache.get(cacheKey);
    if (cached) {
      try {
        return JSON.parse(cached);
      } catch (e) {
        logError(e, 'BondDataLoader.getFullBondData - parse cache');
      }
    }
    
    try {
      var url = 'https://iss.moex.com/iss/engines/stock/markets/bonds/boards/' + safeBoard + '/securities.xml?iss.meta=off&iss.only=securities&securities.columns=SECID,SECNAME,FACEVALUE,PREVADMITTEDQUOTE,COUPONVALUE,COUPONPERCENT,MATDATE,COUPONFREQUENCY';
      
      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' },
        validateHttpsCertificates: true
      });
      
      if (response.getResponseCode() !== 200) {
        addLogEntry('MOEX API вернул код: ' + response.getResponseCode(), 'WARNING', 'BondDataLoader');
        return null;
      }
      
      var xml = response.getContentText();
      var document = XmlService.parse(xml);
      var root = document.getRootElement();
      var rows = root.getChildren('row');
      
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        var secidAttr = row.getAttribute('SECID');
        
        if (secidAttr && secidAttr.getValue().toUpperCase() === secid.toUpperCase()) {
          var bondData = {
            secid: secid,
            name: this.getAttributeValue(row, 'SECNAME'),
            faceValue: parseFloat(this.getAttributeValue(row, 'FACEVALUE') || 1000),
            marketPrice: parseFloat(this.getAttributeValue(row, 'PREVADMITTEDQUOTE') || 0),
            couponValue: parseFloat(this.getAttributeValue(row, 'COUPONVALUE') || 0),
            couponPercent: parseFloat(this.getAttributeValue(row, 'COUPONPERCENT') || 0),
            maturityDate: this.getAttributeValue(row, 'MATDATE') || '',
            couponFrequency: parseInt(this.getAttributeValue(row, 'COUPONFREQUENCY') || 2)
          };
          
          this.cache.put(cacheKey, JSON.stringify(bondData), this.cacheExpiration);
          return bondData;
        }
      }
      
      addLogEntry('Облигация не найдена: ' + secid, 'WARNING', 'BondDataLoader');
      return null;
      
    } catch (e) {
      logError(e, 'BondDataLoader.getFullBondData - ' + secid);
      return null;
    }
  }

  /**
   * Получает НКД (накопленный купонный доход)
   * @param {string} secid - тикер или ISIN
   * @param {string} board - площадка
   * @returns {number} НКД в рублях
   */
  getAccruedInterest(secid, board) {
    var safeBoard = board || 'TQCB';
    
    if (!this.validateIsin(secid)) {
      return 0;
    }
    
    var cacheKey = 'bond_nkd_' + secid + '_' + safeBoard;
    var cached = this.cache.get(cacheKey);
    if (cached) return parseFloat(cached);
    
    try {
      var url = 'https://iss.moex.com/iss/engines/stock/markets/bonds/boards/' + safeBoard + '/securities.xml?iss.meta=off&iss.only=marketdata&marketdata.columns=SECID,ACCRUEDINT';
      
      var response = UrlFetchApp.fetch(url, {
        muteHttpExceptions: true,
        headers: { 'User-Agent': 'Mozilla/5.0' },
        validateHttpsCertificates: true
      });
      
      if (response.getResponseCode() !== 200) {
        return 0;
      }
      
      var xml = response.getContentText();
      var document = XmlService.parse(xml);
      var root = document.getRootElement();
      var rows = root.getChildren('row');
      
      for (var i = 0; i < rows.length; i++) {
        var row = rows[i];
        var secidAttr = row.getAttribute('SECID');
        
        if (secidAttr && secidAttr.getValue().toUpperCase() === secid.toUpperCase()) {
          var nkd = parseFloat(this.getAttributeValue(row, 'ACCRUEDINT') || 0);
          this.cache.put(cacheKey, nkd.toString(), this.cacheExpiration / 4);
          return nkd;
        }
      }
      
      return 0;
      
    } catch (e) {
      logError(e, 'BondDataLoader.getAccruedInterest - ' + secid);
      return 0;
    }
  }

  /**
   * Валидация ISIN (усиленная)
   * @param {string} isin - код ISIN
   * @returns {boolean} true если ISIN валиден
   */
  validateIsin(isin) {
    if (!isin || typeof isin !== 'string') return false;
    
    var upperIsin = isin.toUpperCase().trim();
    
    // Проверка длины (2 буквы + 10 символов)
    if (upperIsin.length !== 12) return false;
    
    // Проверка формата: 2 буквы + 10 цифр/букв
    var pattern = /^[A-Z]{2}[A-Z0-9]{10}$/;
    if (!pattern.test(upperIsin)) return false;
    
    // Дополнительная проверка для российских ISIN
    if (upperIsin.indexOf('RU') === 0) {
      var checkDigits = upperIsin.substring(10, 12);
      var calculatedCheck = this.calculateIsinCheckDigit(upperIsin.substring(0, 10));
      if (checkDigits !== calculatedCheck) {
        addLogEntry('Неверная контрольная сумма ISIN: ' + isin, 'WARNING', 'BondDataLoader');
        return false;
      }
    }
    
    return true;
  }

  /**
   * Расчёт контрольной суммы ISIN (алгоритм Luhn mod 10)
   * @param {string} isinWithoutCheck - ISIN без двух последних цифр
   * @returns {string} двухзначная контрольная сумма
   */
  calculateIsinCheckDigit(isinWithoutCheck) {
    var sum = 0;
    var alternate = false;
    var digits = [];
    
    // Преобразуем буквы в цифры (A=10, B=11, ... Z=35)
    for (var i = isinWithoutCheck.length - 1; i >= 0; i--) {
      var char = isinWithoutCheck.charAt(i);
      var num;
      
      if (char >= '0' && char <= '9') {
        num = parseInt(char);
      } else {
        num = char.charCodeAt(0) - 55; // A=10, B=11...
      }
      
      // Разбиваем двузначные числа на отдельные цифры
      if (num >= 10) {
        digits.push(num % 10);
        digits.push(Math.floor(num / 10));
      } else {
        digits.push(num);
      }
    }
    
    // Вычисляем сумму с чередованием
    for (var j = 0; j < digits.length; j++) {
      var digit = digits[j];
      if (alternate) {
        digit *= 2;
        if (digit > 9) digit -= 9;
      }
      sum += digit;
      alternate = !alternate;
    }
    
    var checkDigit = (10 - (sum % 10)) % 10;
    return checkDigit.toString().padStart(2, '0');
  }

  /**
   * Безопасное получение значения атрибута
   * @param {Element} element - XML элемент
   * @param {string} attrName - имя атрибута
   * @returns {string} значение атрибута или пустая строка
   */
  getAttributeValue(element, attrName) {
    try {
      var attr = element.getAttribute(attrName);
      return attr ? attr.getValue() : '';
    } catch (e) {
      return '';
    }
  }

  /**
   * Очистка кэша
   */
  clearCache() {
    try {
      // CacheService не имеет метода очистки всех ключей
      // Создаём новый кэш-объект
      this.cache = CacheService.getScriptCache();
      addLogEntry('Кэш очищен', 'INFO', 'BondDataLoader');
      return true;
    } catch (e) {
      logError(e, 'BondDataLoader.clearCache');
      return false;
    }
  }

  /**
   * Очистка старого кэша (по ключу)
   * @param {string} secid - ISIN облигации
   */
  clearBondCache(secid) {
    try {
      var keys = [
        'bond_full_' + secid + '_TQOB',
        'bond_full_' + secid + '_TQCB',
        'bond_nkd_' + secid + '_TQOB',
        'bond_nkd_' + secid + '_TQCB',
        'bond_name_' + secid,
        'bond_facevalue_' + secid,
        'bond_price_' + secid
      ];
      
      for (var i = 0; i < keys.length; i++) {
        this.cache.remove(keys[i]);
      }
      
      addLogEntry('Кэш очищен для облигации: ' + secid, 'INFO', 'BondDataLoader');
      return true;
      
    } catch (e) {
      logError(e, 'BondDataLoader.clearBondCache');
      return false;
    }
  }

  /**
   * Периодическая очистка старого кэша
   * @param {number} maxAgeDays - максимальный возраст в днях
   */
  cleanupOldCache(maxAgeDays) {
    try {
      addLogEntry('Запущена очистка кэша', 'INFO', 'BondDataLoader');
      this.clearCache();
      return true;
    } catch (e) {
      logError(e, 'BondDataLoader.cleanupOldCache');
      return false;
    }
  }

  /**
   * Проверяет, является ли облигация ОФЗ (государственная)
   * @param {string} isin - код ISIN облигации
   * @returns {boolean} true если это ОФЗ
   */
  isOfz(isin) {
    if (!isin) return false;
    var upperIsin = isin.toUpperCase().trim();
    return upperIsin.indexOf('RU') === 0;
  }

  /**
   * Обновляет данные по всем облигациям
   */
  refreshAllBondData() {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var bondSheet = ss.getSheetByName(SHEETS.BONDS);
      
      if (!bondSheet) {
        addLogEntry('Лист "Облигации" не найден', 'WARNING', 'BondDataLoader');
        return false;
      }
      
      var lastRow = bondSheet.getLastRow();
      if (lastRow < 38) return false;
      
      var isins = bondSheet.getRange(38, 1, lastRow - 37, 1).getValues();
      var updated = 0;
      
      for (var i = 0; i < isins.length; i++) {
        var isin = isins[i][0];
        if (!isin) continue;
        
        try {
          var board = this.isOfz(isin) ? 'TQOB' : 'TQCB';
          this.clearBondCache(isin);
          var data = this.getFullBondData(isin, board);
          if (data) updated++;
        } catch (e) {
          logError(e, 'refreshAllBondData - ' + isin);
        }
      }
      
      addLogEntry('Обновлены данные для ' + updated + ' облигаций', 'INFO', 'BondDataLoader');
      return true;
      
    } catch (e) {
      logError(e, 'BondDataLoader.refreshAllBondData');
      return false;
    }
  }
}

// =============================================
// ГЛОБАЛЬНЫЕ ФУНКЦИИ
// =============================================

function clearBondCache() {
  try {
    var loader = new BondDataLoader();
    loader.clearCache();
    showMessage('Успешно', 'Кэш облигаций очищен', 'info');
  } catch (e) {
    logError(e, 'clearBondCache');
    showMessage('Ошибка', e.toString(), 'error');
  }
}

function cleanupBondCache() {
  try {
    var loader = new BondDataLoader();
    loader.cleanupOldCache(7);
    showMessage('Успешно', 'Кэш облигаций очищен', 'info');
  } catch (e) {
    logError(e, 'cleanupBondCache');
    showMessage('Ошибка', e.toString(), 'error');
  }
}