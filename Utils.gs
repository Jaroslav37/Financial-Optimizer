/**
 * ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ - утилиты для всех скриптов
 * Исправлено: безопасное логирование, экспорт в защищённую папку
 */

// =============================================
// БАЗОВЫЕ УТИЛИТЫ
// =============================================

function showMessage(title, message, type) {
  try {
    var safeTitle = title || 'Информация';
    var safeMessage = message || '';
    var safeType = type || 'info';
    var ui = SpreadsheetApp.getUi();
    if (ui) {
      var prefix = safeType === 'error' ? '❌ ' : (safeType === 'warning' ? '⚠️ ' : '✅ ');
      ui.alert(prefix + safeTitle, safeMessage, ui.ButtonSet.OK);
    } else Logger.log(safeTitle + ': ' + safeMessage);
  } catch(e) { /* тихо */ }
}

function safeParseNumber(value, defaultValue) {
  var defValue = (defaultValue !== undefined) ? defaultValue : 0;
  if (value === undefined || value === null || value === '') return defValue;
  try {
    if (typeof value === 'number' && !isNaN(value) && isFinite(value)) return Number(value.toFixed(2));
    var str = String(value).replace(/[^\d.,-]/g, '').replace(',', '.');
    var num = parseFloat(str);
    if (isNaN(num) || !isFinite(num)) return defValue;
    return Number(num.toFixed(2));
  } catch(e) { return defValue; }
}

function safeParseDate(dateStr, defaultValue) {
  var defValue = (defaultValue !== undefined) ? defaultValue : null;
  if (!dateStr) return defValue;
  try {
    // Если уже объект Date, возвращаем
    if (dateStr instanceof Date && !isNaN(dateStr.getTime())) return dateStr;
    
    // Если строка, пытаемся разобрать как ДД.ММ.ГГГГ
    if (typeof dateStr === 'string') {
      var parts = dateStr.split('.');
      if (parts.length === 3) {
        var day = parseInt(parts[0], 10);
        var month = parseInt(parts[1], 10) - 1;
        var year = parseInt(parts[2], 10);
        if (!isNaN(day) && !isNaN(month) && !isNaN(year) && 
            day >= 1 && day <= 31 && month >= 0 && month <= 11 && year >= 1900 && year <= 2100) {
          var date = new Date(year, month, day);
          if (!isNaN(date.getTime())) return date;
        }
      }
      // Если не удалось, пробуем стандартный парсинг
      var isoDate = new Date(dateStr);
      if (!isNaN(isoDate.getTime())) return isoDate;
    }
    return defValue;
  } catch(e) {
    return defValue;
  }
}

function formatMoney(value) {
  if (value === undefined || value === null || isNaN(value)) return '0,00 ₽';
  try {
    var num = Number(value).toFixed(2);
    var parts = num.split('.');
    parts[0] = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ' ');
    return parts.join(',') + ' ₽';
  } catch(e) { return '0,00 ₽'; }
}

function formatPercent(value, decimals) {
  var dec = (decimals !== undefined) ? decimals : 2;
  if (value === undefined || value === null || isNaN(value)) return '0%';
  try { return (Number(value) * 100).toFixed(dec) + '%'; } catch(e) { return '0%'; }
}

function formatDate(date) {
  if (!date) return '';
  var d = new Date(date);
  return d.toLocaleDateString('ru-RU');
}

function formatDateForSearch(date) {
  if (!date) return '';
  var d = new Date(date);
  var day = String(d.getDate()).padStart(2, '0');
  var month = String(d.getMonth() + 1).padStart(2, '0');
  var year = d.getFullYear();
  return day + '.' + month + '.' + year;
}

// =============================================
// БЕЗОПАСНОЕ ЛОГИРОВАНИЕ
// =============================================

function logDebug(message, data) {
  // Используем addLogEntry для записи в лист Логи (без чувствительных данных)
  if (data && typeof data === 'object') {
    addLogEntry(message + ' (data logged separately)', 'DEBUG', 'System');
    // Если нужно залогировать данные без sensitive полей, используем security.secureLog
    if (security && security.secureLog) security.secureLog(message, data);
  } else {
    addLogEntry(message, 'DEBUG', 'System');
  }
}

// =============================================
// РАБОТА С ЛИСТАМИ
// =============================================

function getSheetSafe(sheetName) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) { showMessage('Ошибка', 'Лист "' + sheetName + '" не найден', 'error'); return null; }
    return sheet;
  } catch(e) { showMessage('Ошибка', 'Не удалось открыть таблицу: ' + e.toString(), 'error'); return null; }
}

function getCellValue(sheet, row, col, defaultValue) {
  var def = defaultValue !== undefined ? defaultValue : '';
  try { if (!sheet) return def; var v = sheet.getRange(row, col).getValue(); return v !== '' ? v : def; } catch(e) { return def; }
}

function setCellValue(sheet, row, col, value) {
  try { if (!sheet) return false; sheet.getRange(row, col).setValue(value); return true; } catch(e) { return false; }
}

function getLastRowInColumn(sheet, column, startRow) {
  var start = startRow !== undefined ? startRow : 1;
  try {
    if (!sheet) return start;
    var lastRow = sheet.getLastRow();
    if (lastRow < start) return start;
    var values = sheet.getRange(start, column, lastRow - start + 1, 1).getValues();
    for (var i = values.length - 1; i >= 0; i--) if (values[i][0] !== '' && values[i][0] !== null) return start + i;
    return start;
  } catch(e) { return start; }
}

// =============================================
// КАТЕГОРИЗАЦИЯ
// =============================================

function getMonthIndex(header) {
  if (!header || typeof header !== 'string') return -1;
  for (var i = 0; i < MONTHS.length; i++) if (header.indexOf(MONTHS[i]) !== -1) return i;
  return -1;
}

function detectCategory(description) {
  if (!description || typeof description !== 'string') return 'Прочее';
  var desc = description.toLowerCase().trim();
  if (desc === '') return 'Прочее';
  var catKeys = Object.keys(EXPENSE_CATEGORIES);
  for (var i = 0; i < catKeys.length; i++) {
    var category = catKeys[i];
    var keywords = EXPENSE_CATEGORIES[category];
    for (var j = 0; j < keywords.length; j++) if (desc.indexOf(keywords[j]) !== -1) return category;
  }
  return 'Прочее';
}

function detectIncomeCategory(description) {
  if (!description || typeof description !== 'string') return 'Прочее';
  var desc = description.toLowerCase().trim();
  var incomeCategories = {
    'Зарплата': ['зарплата', 'зп', 'аванс', 'оклад', 'salary'],
    'Фриланс': ['фриланс', 'проект', 'заказ', 'подработка', 'freelance'],
    'Инвестиции': ['дивиденд', 'купон', 'продажа акций', 'кэшбэк', 'бонус'],
    'Подарки': ['подарок', 'премия', 'бонус', 'gift'],
    'Возвраты': ['возврат', 'компенсация', 'refund']
  };
  var incKeys = Object.keys(incomeCategories);
  for (var i = 0; i < incKeys.length; i++) {
    var category = incKeys[i];
    var keywords = incomeCategories[category];
    for (var j = 0; j < keywords.length; j++) if (desc.indexOf(keywords[j]) !== -1) return category;
  }
  return 'Прочее';
}

function detectAssetType(name) {
  if (!name) return 'Другое';
  var lowerName = name.toLowerCase();
  if (lowerName.indexOf('облигац') !== -1 || lowerName.indexOf('офз') !== -1 || lowerName.indexOf('бонд') !== -1) return 'Облигация';
  if (lowerName.indexOf('акция') !== -1 || lowerName.indexOf('ао') !== -1 || lowerName.indexOf('ап') !== -1) return 'Акция';
  if (lowerName.indexOf('etf') !== -1 || lowerName.indexOf('фонд') !== -1) return 'ETF';
  return 'Другое';
}

// =============================================
// ЛОГИРОВАНИЕ (основное)
// =============================================

function addLogEntry(message, level, module) {
  try {
    var lvl = level || 'INFO';
    var mod = module || 'System';
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logsSheet = ss.getSheetByName('Логи');
    if (logsSheet) {
      var timestamp = new Date();
      var currentUser = '';
      try { currentUser = Session.getActiveUser().getEmail() || 'anonymous'; } catch(e) { currentUser = 'unknown'; }
      var lastRow = logsSheet.getLastRow();
      logsSheet.getRange(lastRow + 1, 1).setValue(timestamp);
      logsSheet.getRange(lastRow + 1, 2).setValue(lvl);
      logsSheet.getRange(lastRow + 1, 3).setValue(mod);
      logsSheet.getRange(lastRow + 1, 4).setValue(message);
      logsSheet.getRange(lastRow + 1, 5).setValue(currentUser);
      logsSheet.getRange(lastRow + 1, 1).setNumberFormat('dd.MM.yyyy HH:mm:ss');
    } else Logger.log('[' + lvl + '] ' + message);
  } catch(e) { Logger.log('Error in addLogEntry: ' + e.toString()); }
}

function logInfo(message, module) { addLogEntry(message, 'INFO', module); }
function logWarning(message, module) { addLogEntry(message, 'WARNING', module); }
function logError(message, module) { addLogEntry(message, 'ERROR', module); }
function logMessage(message, level) { addLogEntry(message, level || 'INFO', 'System'); }

// =============================================
// ВАЛИДАЦИЯ
// =============================================

function isValidIsin(isin) {
  if (!isin || typeof isin !== 'string') return false;
  var upperIsin = isin.toUpperCase().trim();
  if (upperIsin.length !== 12) return false;
  var pattern = /^[A-Z]{2}[A-Z0-9]{10}$/;
  if (!pattern.test(upperIsin)) return false;
  if (upperIsin.indexOf('RU') === 0) {
    var checkDigits = upperIsin.substring(10, 12);
    var calculatedCheck = calculateIsinCheckDigit(upperIsin.substring(0, 10));
    if (checkDigits !== calculatedCheck) return false;
  }
  return true;
}

function calculateIsinCheckDigit(isinWithoutCheck) {
  var sum = 0, alternate = false, digits = [];
  for (var i = isinWithoutCheck.length - 1; i >= 0; i--) {
    var ch = isinWithoutCheck.charAt(i);
    var num = (ch >= '0' && ch <= '9') ? parseInt(ch) : (ch.charCodeAt(0) - 55);
    if (num >= 10) { digits.push(num % 10); digits.push(Math.floor(num / 10)); }
    else digits.push(num);
  }
  for (var j = 0; j < digits.length; j++) {
    var d = digits[j];
    if (alternate) { d *= 2; if (d > 9) d -= 9; }
    sum += d;
    alternate = !alternate;
  }
  var checkDigit = (10 - (sum % 10)) % 10;
  return checkDigit.toString().padStart(2, '0');
}

function isValidDate(date) { return date instanceof Date && !isNaN(date.getTime()); }
function isPositiveNumber(value) { var num = parseFloat(value); return !isNaN(num) && num > 0 && isFinite(num); }

// =============================================
// РАБОТА С КЭШЕМ
// =============================================

function setCache(key, value, expirationSeconds) {
  var exp = expirationSeconds || 3600;
  try { var cache = CacheService.getScriptCache(); cache.put(key, JSON.stringify(value), exp); return true; } catch(e) { logError(e, 'setCache'); return false; }
}

function getCache(key) {
  try { var cache = CacheService.getScriptCache(); var cached = cache.get(key); if (cached) return JSON.parse(cached); return null; } catch(e) { logError(e, 'getCache'); return null; }
}

function clearCache(key) {
  try { var cache = CacheService.getScriptCache(); cache.remove(key); return true; } catch(e) { logError(e, 'clearCache'); return false; }
}

// =============================================
// РАСЧЁТНЫЕ УТИЛИТЫ
// =============================================

function calculateAnnualReturn(initialValue, finalValue, days) {
  if (!initialValue || initialValue <= 0 || days <= 0) return 0;
  var totalReturn = (finalValue - initialValue) / initialValue;
  var years = days / 365;
  return years > 0 ? (Math.pow(1 + totalReturn, 1 / years) - 1) : 0;
}

function roundTo(value, decimals) {
  var dec = decimals !== undefined ? decimals : 2;
  var factor = Math.pow(10, dec);
  return Math.round(value * factor) / factor;
}

function calculatePercent(value, total) {
  if (!total || total === 0) return 0;
  return (value / total) * 100;
}

// =============================================
// ЭКСПОРТ ДАННЫХ (БЕЗОПАСНЫЙ)
// =============================================

function exportToCsv(data, filename) {
  try {
    var csv = '';
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var processedRow = [];
      for (var j = 0; j < row.length; j++) {
        var cell = row[j];
        if (typeof cell === 'string' && (cell.indexOf(',') !== -1 || cell.indexOf('"') !== -1)) {
          cell = '"' + cell.replace(/"/g, '""') + '"';
        }
        processedRow.push(cell);
      }
      csv += processedRow.join(',') + '\n';
    }
    var blob = Utilities.newBlob(csv, 'text/csv', filename);
    // Создаём папку для резервных копий
    var backupFolderName = 'Резервные копии ' + CONFIG.CURRENT_YEAR;
    var folders = DriveApp.getFoldersByName(backupFolderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(backupFolderName);
    var file = folder.createFile(blob);
    // Ограничиваем доступ: только владелец
    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch(e) {
    logError(e, 'exportToCsv');
    return null;
  }
}

// =============================================
// ДОПОЛНИТЕЛЬНЫЕ ФУНКЦИИ ДЛЯ ИНВЕСТИЦИЙ
// =============================================

function calculateWeightedYield(positions, values, yields) {
  var totalValue = 0, weightedSum = 0;
  for (var i = 0; i < positions.length; i++) {
    var val = values[i], y = yields[i];
    if (val > 0 && !isNaN(y)) { totalValue += val; weightedSum += y * val; }
  }
  return totalValue > 0 ? weightedSum / totalValue : 0;
}

function getCurrentDateForApi() {
  var now = new Date();
  return now.getFullYear() + '-' + String(now.getMonth() + 1).padStart(2, '0') + '-' + String(now.getDate()).padStart(2, '0');
}

function sleep(ms) { Utilities.sleep(ms); }

function retry(func, maxAttempts, delayMs) {
  var attempts = maxAttempts || 3;
  var delay = delayMs || 1000;
  for (var a = 1; a <= attempts; a++) {
    try { return func(); } catch(e) { if (a === attempts) throw e; sleep(delay * a); }
  }
}