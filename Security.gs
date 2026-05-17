/**
 * МОДУЛЬ БЕЗОПАСНОСТИ - контроль доступа и логирование
 * Исправлено: все try-catch конструкции корректны
 */

class SecurityManager {
  constructor() {
    try {
      this.currentUser = Session.getActiveUser().getEmail();
      this.allowedEmails = SECURITY_CONFIG.ALLOWED_EMAILS || [];
      this.strictMode = SECURITY_CONFIG.STRICT_MODE;
      Logger.log(`SecurityManager init: user=${this.currentUser}, strict=${this.strictMode}, allowedCount=${this.allowedEmails.length}`);
      this.isAuthorized = this.checkAuthorization();
      if (this.strictMode && this.allowedEmails.length === 0) {
        this.isAuthorized = false;
        this.logSecurityEvent('STRICT_MODE_NO_ALLOWED_EMAILS', { user: this.currentUser });
        console.error('SECURITY: Strict mode enabled but ALLOWED_EMAILS is empty. Access denied for all users.');
      }
    } catch(e) {
      Logger.log('SecurityManager constructor error: ' + e.toString());
      this.currentUser = 'unknown';
      this.isAuthorized = false;
      this.strictMode = true;
    }
  }

  checkAuthorization() {
    if (!this.strictMode) return true;
    if (!this.allowedEmails || this.allowedEmails.length === 0) return false;
    for (var i = 0; i < this.allowedEmails.length; i++) {
      if (this.allowedEmails[i] === this.currentUser) return true;
    }
    this.logSecurityEvent('UNAUTHORIZED_ACCESS_ATTEMPT', { user: this.currentUser });
    return false;
  }

  secureLog(message, data) {
    if (!this.isAuthorized) return;
    var safeData = data;
    if (data && typeof data === 'object') safeData = this.sanitizeData(data);
    var logEntry = {
      timestamp: new Date().toISOString(),
      user: this.currentUser,
      message: message,
      data: safeData ? JSON.stringify(safeData).substring(0, 1000) : null
    };
    Logger.log(JSON.stringify(logEntry));
  }

  sanitizeData(data) {
    if (!data) return data;
    if (Array.isArray(data)) {
      var arr = [];
      for (var i = 0; i < data.length; i++) {
        arr.push(typeof data[i] === 'object' ? this.sanitizeData(data[i]) : this.cleanValue(data[i]));
      }
      return arr;
    }
    var sanitized = {};
    var keys = Object.keys(data);
    for (var j = 0; j < keys.length; j++) {
      var key = keys[j];
      var isSensitive = false;
      for (var k = 0; k < SECURITY_CONFIG.SENSITIVE_FIELDS.length; k++) {
        if (key.toLowerCase().indexOf(SECURITY_CONFIG.SENSITIVE_FIELDS[k]) !== -1) {
          isSensitive = true;
          break;
        }
      }
      if (isSensitive) {
        sanitized[key] = '[REDACTED]';
      } else if (typeof data[key] === 'object') {
        sanitized[key] = this.sanitizeData(data[key]);
      } else {
        sanitized[key] = this.cleanValue(data[key]);
      }
    }
    return sanitized;
  }

  cleanValue(value) {
    if (typeof value !== 'string') return value;
    for (var i = 0; i < SECURITY_CONFIG.SENSITIVE_FIELDS.length; i++) {
      if (value.toLowerCase().indexOf(SECURITY_CONFIG.SENSITIVE_FIELDS[i]) !== -1) return '[REDACTED]';
    }
    return value;
  }

  logSecurityEvent(eventType, details) {
    var event = { type: eventType, timestamp: new Date().toISOString(), user: this.currentUser };
    if (details) {
      var detailKeys = Object.keys(details);
      for (var i = 0; i < detailKeys.length; i++) event[detailKeys[i]] = details[detailKeys[i]];
    }
    this.saveSecurityEvent(event);
  }

  saveSecurityEvent(event) {
    try {
      var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      var securitySheet = ss.getSheetByName('SecurityLog');
      if (!securitySheet) {
        securitySheet = ss.insertSheet('SecurityLog');
        securitySheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Event Type', 'User', 'Details']]);
        securitySheet.getRange(1, 1, 1, 4).setFontWeight('bold');
        securitySheet.setColumnWidth(1, 180);
        securitySheet.setColumnWidth(2, 120);
        securitySheet.setColumnWidth(3, 200);
        securitySheet.setColumnWidth(4, 400);
      }
      var safeDetails = {};
      if (event.details) {
        var detailKeys = Object.keys(event.details);
        for (var i = 0; i < detailKeys.length; i++) safeDetails[detailKeys[i]] = '[REDACTED]';
      }
      securitySheet.appendRow([event.timestamp, event.type, event.user, JSON.stringify(safeDetails)]);
    } catch(e) {
      Logger.log('Security log error: ' + e.toString());
    }
  }

  checkAccess(functionName) {
    if (!this.isAuthorized) {
      this.logSecurityEvent('UNAUTHORIZED_ACCESS_ATTEMPT', { function: functionName, user: this.currentUser });
      throw new Error('Unauthorized access attempt to function: ' + functionName);
    }
    if (this.strictMode && /delete|remove|clear|export|send|addAllowedEmail|removeAllowedEmail/i.test(functionName)) {
      Logger.log(`Strict mode: sensitive action ${functionName} by ${this.currentUser}`);
    }
    return true;
  }

  addAllowedEmail(email) {
    if (!this.isAuthorized && this.strictMode) throw new Error('Only authorized users can modify allowed emails');
    if (email && typeof email === 'string') {
      this.allowedEmails.push(email);
      var props = PropertiesService.getScriptProperties();
      props.setProperty('ALLOWED_EMAILS', this.allowedEmails.join(','));
      SECURITY_CONFIG.ALLOWED_EMAILS = this.allowedEmails;
      this.logSecurityEvent('EMAIL_ADDED', { email: email });
      return true;
    }
    return false;
  }

  removeAllowedEmail(email) {
    if (!this.isAuthorized && this.strictMode) throw new Error('Only authorized users can modify allowed emails');
    if (!email) return false;
    var newList = [];
    for (var i = 0; i < this.allowedEmails.length; i++) {
      if (this.allowedEmails[i] !== email) newList.push(this.allowedEmails[i]);
    }
    if (newList.length !== this.allowedEmails.length) {
      this.allowedEmails = newList;
      var props = PropertiesService.getScriptProperties();
      props.setProperty('ALLOWED_EMAILS', this.allowedEmails.join(','));
      SECURITY_CONFIG.ALLOWED_EMAILS = this.allowedEmails;
      this.logSecurityEvent('EMAIL_REMOVED', { email: email });
      return true;
    }
    return false;
  }

  getAllowedEmails() { return this.allowedEmails.slice(); }
  getCurrentUser() { return this.currentUser; }
  isAuthorizedUser() { return this.isAuthorized; }
}

var security = null;
function getSecurityManager() {
  if (!security) security = new SecurityManager();
  return security;
}
var security = getSecurityManager();

function showSecurityInfo() {
  var sec = getSecurityManager();
  if (!sec) { showMessage('Ошибка', 'Не удалось инициализировать менеджер безопасности', 'error'); return; }
  var report = '🔒 ИНФОРМАЦИЯ БЕЗОПАСНОСТИ\n═══════════════════════════\n\n';
  report += '👤 Текущий пользователь: ' + sec.getCurrentUser() + '\n';
  report += '🔐 Авторизован: ' + (sec.isAuthorizedUser() ? '✅ Да' : '❌ Нет') + '\n';
  report += '🛡️ Строгий режим: ' + (SECURITY_CONFIG.STRICT_MODE ? '✅ Включён' : '❌ Выключен') + '\n\n';
  var emails = sec.getAllowedEmails();
  report += '📧 Разрешённые email (' + emails.length + '):\n';
  if (emails.length === 0) report += '   (список пуст - доступ только при выключенном строгом режиме)\n';
  else for (var i = 0; i < emails.length; i++) report += '   • ' + emails[i] + '\n';
  showMessage('Безопасность', report, 'info');
}

function addAllowedEmail() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('➕ Добавить email', 'Введите email для добавления в список разрешённых:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var email = response.getResponseText().trim();
  try {
    getSecurityManager().addAllowedEmail(email);
    showMessage('Успешно', 'Email ' + email + ' добавлен в список разрешённых', 'info');
  } catch(e) { showMessage('Ошибка', e.toString(), 'error'); }
}

function removeAllowedEmail() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt('🗑️ Удалить email', 'Введите email для удаления из списка разрешённых:', ui.ButtonSet.OK_CANCEL);
  if (response.getSelectedButton() !== ui.Button.OK) return;
  var email = response.getResponseText().trim();
  try {
    getSecurityManager().removeAllowedEmail(email);
    showMessage('Успешно', 'Email ' + email + ' удалён из списка разрешённых', 'info');
  } catch(e) { showMessage('Ошибка', e.toString(), 'error'); }
}