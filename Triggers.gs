/**
 * ТРИГГЕРЫ – автоматические задачи (без инвестиций, ипотеки, целей)
 */

function setupAllTriggers() {
  if (!AUTOMATION || !AUTOMATION.AUTO_SETUP_TRIGGERS) {
    var ui = SpreadsheetApp.getUi();
    if (ui.alert('⚠️ Триггеры отключены', 'Автоматическое создание триггеров отключено. Включить?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  }
  var triggers = ScriptApp.getProjectTriggers();
  for (var t = 0; t < triggers.length; t++) try { ScriptApp.deleteTrigger(triggers[t]); } catch(e) {}

  ScriptApp.newTrigger('calculateAllDebts').timeBased().everyDays(1).atHour(3).create();
  ScriptApp.newTrigger('updateFinancialPlan').timeBased().onMonthDay(1).atHour(4).create();
  ScriptApp.newTrigger('checkBudgetStatus').timeBased().everyDays(1).atHour(9).create();
  ScriptApp.newTrigger('createNewYearSheet').timeBased().onMonthDay(31).atHour(23).nearMinute(55).create();
  ScriptApp.newTrigger('sendWeeklyReport').timeBased().onWeekDay(ScriptApp.WeekDay.SUNDAY).atHour(19).create();
  ScriptApp.newTrigger('backupData').timeBased().onWeekDay(ScriptApp.WeekDay.SATURDAY).atHour(2).create();

  showMessage('Информация', '✅ Триггеры успешно созданы', 'info');
  addLogEntry('Все триггеры настроены', 'INFO', 'Triggers');
}

function backupData() {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var backupFolderName = 'Резервные копии ' + CONFIG.CURRENT_YEAR;
    var folders = DriveApp.getFoldersByName(backupFolderName);
    var backupFolder = folders.hasNext() ? folders.next() : DriveApp.createFolder(backupFolderName);
    var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    var backupName = 'Резервная копия ' + dateStr;
    var backupFile = DriveApp.getFileById(SPREADSHEET_ID).makeCopy(backupName, backupFolder);
    backupFile.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
    var files = backupFolder.getFiles();
    var fileList = [];
    while (files.hasNext()) fileList.push(files.next());
    if (fileList.length > 10) {
      fileList.sort(function(a, b) { return b.getDateCreated() - a.getDateCreated(); });
      for (var i = 10; i < fileList.length; i++) fileList[i].setTrashed(true);
    }
    addLogEntry('Создана резервная копия: ' + backupName, 'INFO', 'Triggers');
  } catch(e) { logError(e, 'backupData'); }
}

function sendWeeklyReport() {
  try {
    var data = collectYearData();
    if (!data) return;
    var email = Session.getActiveUser().getEmail();
    var body = '📊 Еженедельный отчёт\n\n💰 Доходы: ' + formatMoney(data.income) + '\n💸 Расходы: ' + formatMoney(data.expenses);
    MailApp.sendEmail(email, '📊 Еженедельный финансовый отчёт', body);
    addLogEntry('Еженедельный отчёт отправлен', 'INFO', 'Triggers');
  } catch(e) { logError(e, 'sendWeeklyReport'); }
}

// Вспомогательные функции для меню
function manageTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  var report = '🎛️ УПРАВЛЕНИЕ ТРИГГЕРАМИ\n\n';
  if (triggers.length === 0) report += 'Нет активных триггеров\n';
  else {
    report += 'Активных триггеров: ' + triggers.length + '\n\n';
    for (var i = 0; i < triggers.length; i++) {
      var trigger = triggers[i];
      report += (i + 1) + '. ' + trigger.getHandlerFunction() + '\n';
    }
  }
  SpreadsheetApp.getUi().alert('📋 Триггеры', report, SpreadsheetApp.getUi().ButtonSet.OK);
}

function stopAllTriggers() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Остановка триггеров', 'Вы уверены?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) ScriptApp.deleteTrigger(triggers[i]);
  showMessage('Информация', '✅ Все триггеры остановлены', 'info');
}

function checkAllTriggers() { manageTriggers(); }
function logAllTriggers() {
  var triggers = ScriptApp.getProjectTriggers();
  addLogEntry('Total triggers: ' + triggers.length, 'INFO', 'Triggers');
  for (var i = 0; i < triggers.length; i++) addLogEntry('Trigger: ' + triggers[i].getHandlerFunction(), 'INFO', 'Triggers');
  showMessage('Логи', 'Триггеры записаны в логи', 'info');
}