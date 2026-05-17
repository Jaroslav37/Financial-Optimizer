/**
 * МЕНЕДЖЕР ТАБЛИЦ - оптимизированная работа с данными
 * Исправлено: безопасное логирование
 */

class SheetManager {
  constructor(spreadsheetId) {
    this.spreadsheetId = spreadsheetId;
    this.cache = {};
    this.pendingUpdates = {};
  }
  
  getSheetData(sheetName, range) {
    try {
      var cacheKey = sheetName + ':' + range;
      if (this.cache[cacheKey] !== undefined) return this.cache[cacheKey];
      var ss = SpreadsheetApp.openById(this.spreadsheetId);
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        logError(new Error('Лист "' + sheetName + '" не найден'), 'SheetManager.getSheetData');
        return null;
      }
      var data = sheet.getRange(range).getValues();
      this.cache[cacheKey] = data;
      return data;
    } catch(e) {
      logError(e, 'SheetManager.getSheetData');
      if (security && security.secureLog) security.secureLog('Error getting sheet data', { error: e.toString() });
      return null;
    }
  }
  
  scheduleUpdate(sheetName, row, col, value, note) {
    try {
      var key = sheetName + ':' + row + ':' + col;
      if (!this.pendingUpdates[sheetName]) this.pendingUpdates[sheetName] = {};
      this.pendingUpdates[sheetName][key] = { row: row, col: col, value: value, note: note || null };
    } catch(e) { logError(e, 'SheetManager.scheduleUpdate'); }
  }
  
  applyUpdates() {
    try {
      var ss = SpreadsheetApp.openById(this.spreadsheetId);
      var sheetNames = Object.keys(this.pendingUpdates);
      for (var s = 0; s < sheetNames.length; s++) {
        var sheetName = sheetNames[s];
        var sheet = ss.getSheetByName(sheetName);
        if (!sheet) continue;
        var updates = this.pendingUpdates[sheetName];
        var updateKeys = Object.keys(updates);
        var rowsMap = {};
        for (var u = 0; u < updateKeys.length; u++) {
          var update = updates[updateKeys[u]];
          if (!rowsMap[update.row]) rowsMap[update.row] = [];
          rowsMap[update.row].push(update);
        }
        var rowNumbers = Object.keys(rowsMap);
        for (var r = 0; r < rowNumbers.length; r++) {
          var row = parseInt(rowNumbers[r]);
          var rowUpdates = rowsMap[row];
          for (var ru = 0; ru < rowUpdates.length; ru++) {
            var updateItem = rowUpdates[ru];
            sheet.getRange(row, updateItem.col).setValue(updateItem.value);
            if (updateItem.note) sheet.getRange(row, updateItem.col).setNote(updateItem.note);
          }
        }
      }
      this.pendingUpdates = {};
      this.cache = {};
    } catch(e) {
      logError(e, 'SheetManager.applyUpdates');
      if (security && security.secureLog) security.secureLog('Error applying updates', { error: e.toString() });
    }
  }
  
  loadFullSheet(sheetName) {
    try {
      var ss = SpreadsheetApp.openById(this.spreadsheetId);
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        logError(new Error('Лист "' + sheetName + '" не найден'), 'SheetManager.loadFullSheet');
        return null;
      }
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow === 0 || lastCol === 0) return { headers: [], data: [], lastRow: 0, lastCol: 0 };
      var range = sheet.getRange(1, 1, lastRow, lastCol);
      var values = range.getValues();
      return { headers: values[0] || [], data: values, lastRow: lastRow, lastCol: lastCol };
    } catch(e) {
      logError(e, 'SheetManager.loadFullSheet');
      if (security && security.secureLog) security.secureLog('Error loading full sheet', { error: e.toString() });
      return null;
    }
  }
  
  clearCache() { try { this.cache = {}; } catch(e) { logError(e, 'SheetManager.clearCache'); } }
  
  getCellValue(sheetName, row, col) {
    try {
      var cacheKey = sheetName + ':cell:' + row + ':' + col;
      if (this.cache[cacheKey] !== undefined) return this.cache[cacheKey];
      var ss = SpreadsheetApp.openById(this.spreadsheetId);
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return null;
      var value = sheet.getRange(row, col).getValue();
      this.cache[cacheKey] = value;
      return value;
    } catch(e) { logError(e, 'SheetManager.getCellValue'); return null; }
  }
  
  setCellValue(sheetName, row, col, value, immediate) {
    try {
      if (immediate) {
        var ss = SpreadsheetApp.openById(this.spreadsheetId);
        var sheet = ss.getSheetByName(sheetName);
        if (sheet) sheet.getRange(row, col).setValue(value);
      } else {
        this.scheduleUpdate(sheetName, row, col, value);
      }
    } catch(e) { logError(e, 'SheetManager.setCellValue'); }
  }
  
  getRangeData(sheetName, startRow, startCol, numRows, numCols) {
    try {
      var cacheKey = sheetName + ':range:' + startRow + ':' + startCol + ':' + numRows + ':' + numCols;
      if (this.cache[cacheKey] !== undefined) return this.cache[cacheKey];
      var ss = SpreadsheetApp.openById(this.spreadsheetId);
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return null;
      var data = sheet.getRange(startRow, startCol, numRows, numCols).getValues();
      this.cache[cacheKey] = data;
      return data;
    } catch(e) { logError(e, 'SheetManager.getRangeData'); return null; }
  }
  
  batchUpdate(sheetName, updates) {
    try {
      if (!updates || updates.length === 0) return;
      var ss = SpreadsheetApp.openById(this.spreadsheetId);
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return;
      for (var i = 0; i < updates.length; i++) {
        var update = updates[i];
        sheet.getRange(update.row, update.col).setValue(update.value);
        if (update.note) sheet.getRange(update.row, update.col).setNote(update.note);
      }
    } catch(e) { logError(e, 'SheetManager.batchUpdate'); }
  }
  
  getLastRow(sheetName) {
    try {
      var ss = SpreadsheetApp.openById(this.spreadsheetId);
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return 0;
      return sheet.getLastRow();
    } catch(e) { logError(e, 'SheetManager.getLastRow'); return 0; }
  }
  
  getLastColumn(sheetName) {
    try {
      var ss = SpreadsheetApp.openById(this.spreadsheetId);
      var sheet = ss.getSheetByName(sheetName);
      if (!sheet) return 0;
      return sheet.getLastColumn();
    } catch(e) { logError(e, 'SheetManager.getLastColumn'); return 0; }
  }
}

var sheetManager = null;
function getSheetManager() {
  try {
    if (!sheetManager) sheetManager = new SheetManager(SPREADSHEET_ID);
    return sheetManager;
  } catch(e) { logError(e, 'getSheetManager'); return null; }
}
var sheetManager = getSheetManager();