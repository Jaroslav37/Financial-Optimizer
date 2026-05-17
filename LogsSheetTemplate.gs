/**
 * ЛИСТ ЛОГОВ - создание и управление
 */

function createLogsSheet() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var logsSheet = ss.getSheetByName('Логи');
    if (logsSheet) { var ui = SpreadsheetApp.getUi(); if (ui.alert('Создание листа логов', 'Лист "Логи" уже существует. Пересоздать?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return; ss.deleteSheet(logsSheet); }
    logsSheet = ss.insertSheet('Логи');
    logsSheet.setTabColor('#f1c40f');
    var headers = ['Дата и время', 'Уровень', 'Модуль', 'Сообщение', 'ID пользователя'];
    logsSheet.getRange(1,1,1,headers.length).setValues([headers]);
    logsSheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#f0f0f0').setHorizontalAlignment('center');
    logsSheet.setColumnWidth(1,150); logsSheet.setColumnWidth(2,80); logsSheet.setColumnWidth(3,120); logsSheet.setColumnWidth(4,500); logsSheet.setColumnWidth(5,150);
    logsSheet.setFrozenRows(1);
    logsSheet.getDataRange().createFilter();
    logsSheet.getRange(2,1,1000,1).setNumberFormat('dd.MM.yyyy HH:mm:ss');
    var exampleLogs = [[new Date(),'INFO','System','Лист "Логи" создан',Session.getActiveUser().getEmail()]];
    if (exampleLogs.length) logsSheet.getRange(2,1,exampleLogs.length,exampleLogs[0].length).setValues(exampleLogs);
    logsSheet.getDataRange().sort({column:1,ascending:false});
    SpreadsheetApp.getUi().alert('✅ Лист "Логи" успешно создан!');
    return { success: true };
  } catch(error) { return { success: false, error: error.toString() }; }
}
function clearLogsSheet() { var sheet = getSheetSafe('Логи'); if(!sheet) return; var lastRow=sheet.getLastRow(); if(lastRow>1) sheet.getRange(2,1,lastRow-1,5).clearContent(); SpreadsheetApp.getUi().alert('✅ Лист "Логи" очищен'); }
function exportLogsToCsv() { var sheet=getSheetSafe('Логи'); if(!sheet) return; var data=sheet.getDataRange().getValues(); if(data.length<2) return; var csv=''; for(var i=0;i<data.length;i++){ var row=data[i]; var processedRow=[]; for(var j=0;j<row.length;j++){ var cell=row[j]; if(typeof cell==='string' && (cell.indexOf(',')!==-1||cell.indexOf('"')!==-1)) cell='"'+cell.replace(/"/g,'""')+'"'; processedRow.push(cell); } csv+=processedRow.join(',')+'\n'; } var blob=Utilities.newBlob(csv,'text/csv','logs_'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd_HH-mm')+'.csv'); var file=DriveApp.createFile(blob); SpreadsheetApp.getUi().alert('✅ Логи экспортированы: '+file.getUrl()); }
function cleanupOldLogs(daysToKeep) { var keepDays=daysToKeep||30; var sheet=getSheetSafe('Логи'); if(!sheet) return; var data=sheet.getDataRange().getValues(); if(data.length<2) return; var cutoff=new Date(); cutoff.setDate(cutoff.getDate()-keepDays); var rowsToDelete=[]; for(var i=1;i<data.length;i++) if(data[i][0] && data[i][0]<cutoff) rowsToDelete.push(i+1); rowsToDelete.reverse(); for(var j=0;j<rowsToDelete.length;j++) sheet.deleteRow(rowsToDelete[j]); SpreadsheetApp.getUi().alert('✅ Удалено старых записей: '+rowsToDelete.length); }
function showLogsStats() { var sheet=getSheetSafe('Логи'); if(!sheet) return; var data=sheet.getDataRange().getValues(); var stats={total:data.length-1,byLevel:{INFO:0,WARNING:0,ERROR:0}}; for(var i=1;i<data.length;i++){ var level=data[i][1]; if(stats.byLevel[level]!==undefined) stats.byLevel[level]++; } var report='📊 СТАТИСТИКА ЛОГОВ\n\n📋 Всего записей: '+stats.total+'\n\n📈 ПО УРОВНЯМ:\n   INFO: '+stats.byLevel.INFO+'\n   WARNING: '+stats.byLevel.WARNING+'\n   ERROR: '+stats.byLevel.ERROR; SpreadsheetApp.getUi().alert(report); }
function initLogs() { createLogsSheet(); addLogEntry('Система инициализирована','INFO','System'); }
function showLogsStatsMenu() { showLogsStats(); }
function exportLogs() { exportLogsToCsv(); }
function cleanupLogsMenu() { var ui=SpreadsheetApp.getUi(); var response=ui.prompt('Очистка старых логов','Введите количество дней для хранения (по умолчанию 30):',ui.ButtonSet.OK_CANCEL); if(response.getSelectedButton()!==ui.Button.OK) return; var days=parseInt(response.getResponseText())||30; cleanupOldLogs(days); }