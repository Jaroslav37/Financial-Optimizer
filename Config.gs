/**
 * КОНФИГУРАЦИЯ - общие настройки для всех скриптов
 * Исправлено: безопасное логирование (addLogEntry вместо Logger.log)
 */

function getScriptProperty(key, defaultValue) {
  try {
    var props = PropertiesService.getScriptProperties();
    var value = props.getProperty(key);
    return (value !== null && value !== '') ? value : defaultValue;
  } catch(e) { logError(e, 'getScriptProperty - ' + key); return defaultValue; }
}

function getScriptPropertyNumber(key, defaultValue) {
  return parseFloat(getScriptProperty(key, defaultValue));
}

function getScriptPropertyBoolean(key, defaultValue) {
  var value = getScriptProperty(key, defaultValue);
  if (typeof value === 'string') return value.toLowerCase() === 'true' || value === '1';
  return value === true;
}

var SPREADSHEET_ID = getScriptProperty('SPREADSHEET_ID', '');
if (!SPREADSHEET_ID) {
  addLogEntry('SPREADSHEET_ID не настроен в PropertiesService', 'WARNING', 'Config');
}

var MOEX_API_KEY = getScriptProperty('MOEX_API_KEY', '');
var YAHOO_API_KEY = getScriptProperty('YAHOO_API_KEY', '');
var ALPHA_VANTAGE_KEY = getScriptProperty('ALPHA_VANTAGE_KEY', '');

var SECURITY_CONFIG = {
  ALLOWED_EMAILS: (function() {
    var emails = getScriptProperty('ALLOWED_EMAILS', '');
    if (emails === '') return [];
    return emails.split(',').map(function(e) { return e.trim(); }).filter(function(e) { return e !== ''; });
  })(),
  STRICT_MODE: getScriptPropertyBoolean('STRICT_MODE', true),
  SENSITIVE_FIELDS: ['password', 'token', 'key', 'secret', 'пароль', 'ключ', 'api', 'id', 'email']
};

var NOTIFICATION_CONFIG = {
  EMAIL_RECIPIENT: getScriptProperty('EMAIL_RECIPIENT', Session.getActiveUser().getEmail()),
  SEND_DAILY_REPORT: getScriptPropertyBoolean('SEND_DAILY_REPORT', false),
  SEND_WEEKLY_REPORT: getScriptPropertyBoolean('SEND_WEEKLY_REPORT', true),
  SEND_ALERTS: getScriptPropertyBoolean('SEND_ALERTS', true),
  SEND_GOAL_NOTIFICATIONS: getScriptPropertyBoolean('SEND_GOAL_NOTIFICATIONS', true)
};

var SHEETS = {
  CURRENT_YEAR: '2026',
  FINANCIAL_PLAN: 'Финансовый план',
  BUDGET: 'Бюджет',
  DEBTS: 'Долговые обязательства',
  MORTGAGE: 'Ипотека',
  GOALS: 'Смета целей',
  ANALYTICS: 'Аналитика',
  EXPENSE_DETAILS: 'Детали расходов',
  INCOME_DETAILS: 'Детали доходов',
  REPORTS: 'Отчёты',
  SETTINGS: 'Настройки',
  INVESTMENTS: 'Инвестиции',
  BONDS: 'Облигации',
  STOCKS: 'Акции',
  SECURITY_LOG: 'SecurityLog'
};

var INVESTMENT_COLUMNS = {
  DATE: 1, ACCOUNT: 2, NAME: 3, TYPE: 4, ISIN: 5, QUANTITY: 6, VALUE: 7, CURRENCY: 8,
  MATURITY_DATE: 9, COUPON_VALUE: 10, COUPON_PERIOD: 11, REMAINING_COUPONS: 12, NKD: 13,
  FACE_VALUE: 14, PRICE_PERCENT: 15, CURRENT_YIELD: 16, YTM: 17, YTM_RUB: 18,
  DIVIDEND_PER_SHARE: 19, DIVIDEND_YIELD: 20, GROWTH_POTENTIAL: 21, TICKER: 22,
  NOTE: 23, RECORD_ID: 24, SOURCE: 25, PERIOD: 26
};

var ROWS = {
  DATA_START: 2, DATA_END: 32, TOTAL: 33, BALANCE: 34, GIFTS: 35, COMPENSATING: 36,
  PROJECTS: 37, COMPENSATING2: 38, PAYMENT: 39, INSURANCE: 40, PENSION: 41, HOURS: 42
};

var COLUMNS = { A:1, B:2, C:3, D:4, E:5, F:6, G:7, H:8, I:9, J:10, K:11, L:12, M:13, N:14, O:15, P:16, Q:17, R:18, S:19, T:20 };

var MONTHS = ['Январь','Февраль','Март','Апрель','Май','Июнь','Июль','Август','Сентябрь','Октябрь','Ноябрь','Декабрь'];
var MONTHS_INDEX = {'Январь':0,'Февраль':1,'Март':2,'Апрель':3,'Май':4,'Июнь':5,'Июль':6,'Август':7,'Сентябрь':8,'Октябрь':9,'Ноябрь':10,'Декабрь':11};

var MORTGAGE_CONFIG = {
  ANNUAL_RATE: getScriptPropertyNumber('MORTGAGE_ANNUAL_RATE', 15),
  LATE_PENALTY_RATE: getScriptPropertyNumber('MORTGAGE_LATE_PENALTY_RATE', 0.1),
  MONTHLY_PENALTY_RATE: getScriptPropertyNumber('MORTGAGE_MONTHLY_PENALTY_RATE', 0.01),
  PAYMENT_DAY: getScriptPropertyNumber('MORTGAGE_PAYMENT_DAY', 15)
};

var CONFIG = {
  CURRENT_YEAR: 2026,
  CURRENT_MONTH: new Date().getMonth(),
  MORTGAGE: MORTGAGE_CONFIG
};

var EXPENSE_CATEGORIES = {
  'Жильё': ['квартплата','коммунальные','ремонт','аренда','жильё','ипотека','жкх'],
  'Продукты': ['еда','продукты','магазин','супермаркет','продуктовый','пятёрочка','магнит'],
  'Транспорт': ['бензин','такси','метро','автобус','транспорт','проезд','тройка'],
  'Здоровье': ['аптека','врач','лекарства','больница','здоровье','стоматолог'],
  'Развлечения': ['кино','кафе','ресторан','досуг','развлечения','игры','подписки'],
  'Одежда': ['одежда','обувь','аксессуары','wildberries','ozon'],
  'Связь': ['интернет','телефон','мобильный','связь'],
  'Долги': ['кредит','долг','проценты','займ'],
  'Образование': ['курсы','книги','обучение'],
  'Инвестиции': ['комиссия','брокер','депозитарий','налог'],
  'Прочее': []
};

var INCOME_CATEGORIES = {
  'Зарплата': ['зарплата','зп','аванс','оклад'],
  'Фриланс': ['фриланс','проект','заказ','подработка'],
  'Дивиденды': ['дивиденды','дивиденд'],
  'Купоны': ['купон','купоны','облигации'],
  'Продажа активов': ['продажа','акции','облигации'],
  'Кэшбэк': ['кэшбэк','бонусы','мили'],
  'Подарки': ['подарок','премия'],
  'Возвраты': ['возврат','компенсация']
};

var BUDGET_CONFIG = {
  ENABLED: true,
  DEFAULT_LIMITS: {'Жильё':40000,'Продукты':30000,'Транспорт':8000,'Здоровье':5000,'Развлечения':10000,'Одежда':8000,'Связь':2500,'Долги':20000,'Образование':5000,'Инвестиции':5000,'Прочее':5000},
  WARNING_THRESHOLD: 80,
  CRITICAL_THRESHOLD: 95,
  ALERT_ON_EXCEED: true,
  SAVE_EXCESS: true,
  ROLLOVER_UNUSED: true
};

var AUTOMATION = {
  AUTO_CATEGORIZE: true,
  AUTO_BACKUP: true,
  BACKUP_FREQUENCY_DAYS: 7,
  AUTO_EMAIL_REPORT: NOTIFICATION_CONFIG.SEND_WEEKLY_REPORT,
  EMAIL_RECIPIENT: NOTIFICATION_CONFIG.EMAIL_RECIPIENT,
  NOTIFY_ON_EXCEED: NOTIFICATION_CONFIG.SEND_ALERTS,
  NOTIFY_ON_GOAL_ACHIEVED: NOTIFICATION_CONFIG.SEND_GOAL_NOTIFICATIONS,
  AUTO_SETUP_TRIGGERS: false
};

var CURRENCY = { CODE: 'RUB', SYMBOL: '₽', DECIMAL_PLACES: 2, THOUSAND_SEPARATOR: ' ' };
var INVESTMENT_CONFIG = { TAX_RATE: 0.13, DIVIDEND_TAX: 0.13, COUPON_TAX: 0.13, IIS_BENEFIT: 52000, BROKER_COMMISSION: 0.0005 };
var MENU_CONFIG = getScriptProperty('MENU_CONFIG', 'default');

function getCurrentYearSheet() {
  try {
    if (!SPREADSHEET_ID) { logError(new Error('SPREADSHEET_ID не настроен'), 'getCurrentYearSheet'); return null; }
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return ss.getSheetByName(SHEETS.CURRENT_YEAR);
  } catch(e) { logError(e, 'getCurrentYearSheet'); return null; }
}

function sheetExists(sheetName) {
  try {
    if (!SPREADSHEET_ID) return false;
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    return ss.getSheetByName(sheetName) !== null;
  } catch(e) { logError(e, 'sheetExists - ' + sheetName); return false; }
}

function getApiKey(keyName) {
  try {
    var key = getScriptProperty(keyName, '');
    if (!key) addLogEntry('API ключ не найден: ' + keyName, 'WARNING', 'Config');
    return key;
  } catch(e) { logError(e, 'getApiKey - ' + keyName); return ''; }
}

function validateConfig() {
  try {
    var errors = [], warnings = [];
    if (!SPREADSHEET_ID) errors.push('❌ SPREADSHEET_ID не настроен в PropertiesService');
    if (SECURITY_CONFIG.ALLOWED_EMAILS.length === 0 && SECURITY_CONFIG.STRICT_MODE) warnings.push('⚠️ ALLOWED_EMAILS не настроен, но строгий режим включён');
    if (errors.length > 0) { var message = 'Ошибки конфигурации:\n' + errors.join('\n'); if (typeof showMessage === 'function') showMessage('Ошибка конфигурации', message, 'error'); addLogEntry(message, 'ERROR', 'Config'); return false; }
    if (warnings.length > 0) { var warnMsg = 'Предупреждения:\n' + warnings.join('\n'); if (typeof showMessage === 'function') showMessage('Конфигурация', warnMsg, 'warning'); addLogEntry(warnMsg, 'WARNING', 'Config'); }
    addLogEntry('Конфигурация валидна', 'INFO', 'Config');
    return true;
  } catch(e) { logError(e, 'validateConfig'); return false; }
}

function showConfig() {
  try {
    var report = '⚙️ КОНФИГУРАЦИЯ\n═══════════════════\n\n📁 SPREADSHEET_ID: ' + (SPREADSHEET_ID ? '✅ настроен' : '❌ не настроен') + '\n🔑 MOEX_API_KEY: ' + (MOEX_API_KEY ? '✅ настроен' : '❌ не настроен') + '\n📧 EMAIL_RECIPIENT: ' + (NOTIFICATION_CONFIG.EMAIL_RECIPIENT || 'не настроен') + '\n🛡️ Строгий режим: ' + (SECURITY_CONFIG.STRICT_MODE ? '✅ вкл' : '❌ выкл') + '\n\n📧 Разрешённые email (' + SECURITY_CONFIG.ALLOWED_EMAILS.length + '):\n';
    if (SECURITY_CONFIG.ALLOWED_EMAILS.length === 0) report += '   (список пуст)\n';
    else for (var i = 0; i < SECURITY_CONFIG.ALLOWED_EMAILS.length; i++) report += '   • ' + SECURITY_CONFIG.ALLOWED_EMAILS[i] + '\n';
    report += '\n🏠 Ипотека:\n   Ставка: ' + MORTGAGE_CONFIG.ANNUAL_RATE + '%\n   Пеня за просрочку: ' + (MORTGAGE_CONFIG.LATE_PENALTY_RATE * 100) + '%\n   Ежемесячная пеня: ' + (MORTGAGE_CONFIG.MONTHLY_PENALTY_RATE * 100) + '%\n';
    if (typeof showMessage === 'function') showMessage('Конфигурация', report, 'info');
    else addLogEntry(report, 'INFO', 'Config');
  } catch(e) { logError(e, 'showConfig'); if (typeof showMessage === 'function') showMessage('Ошибка', e.toString(), 'error'); }
}

function logConfig() {
  try {
    addLogEntry('=== CONFIG LOADED ===', 'DEBUG', 'Config');
    addLogEntry('SPREADSHEET_ID: [REDACTED]', 'DEBUG', 'Config');
    addLogEntry('CURRENT_YEAR: ' + CONFIG.CURRENT_YEAR, 'DEBUG', 'Config');
    addLogEntry('Budget enabled: ' + BUDGET_CONFIG.ENABLED, 'DEBUG', 'Config');
    addLogEntry('Auto-categorize: ' + AUTOMATION.AUTO_CATEGORIZE, 'DEBUG', 'Config');
    addLogEntry('Strict mode: ' + SECURITY_CONFIG.STRICT_MODE, 'DEBUG', 'Config');
    addLogEntry('Allowed emails: ' + SECURITY_CONFIG.ALLOWED_EMAILS.length, 'DEBUG', 'Config');
    addLogEntry('Конфигурация загружена', 'INFO', 'Config');
  } catch(e) { logError(e, 'logConfig'); }
}

logConfig();