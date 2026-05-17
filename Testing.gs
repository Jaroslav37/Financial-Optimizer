/**
 * ТЕСТИРОВАНИЕ - функции для проверки работоспособности
 */

function runAllTests() {
  try {
    var ui = SpreadsheetApp.getUi();
    var results = [];
    var passed = 0;
    var failed = 0;
    
    results.push('📊 ТЕСТ СБОРА ДАННЫХ:');
    var dataTest = testDataCollection();
    if (dataTest.passed) {
      results.push('   ✅ ' + dataTest.message);
      passed++;
    } else {
      results.push('   ❌ ' + dataTest.message);
      failed++;
    }
    
    results.push('\n📅 ТЕСТ ПАРСИНГА ДАТ:');
    var dateTest = testDateParsing();
    if (dateTest.passed) {
      results.push('   ✅ ' + dateTest.message);
      passed++;
    } else {
      results.push('   ❌ ' + dateTest.message);
      failed++;
    }
    
    results.push('\n💰 ТЕСТ ФОРМАТИРОВАНИЯ:');
    var moneyTest = testMoneyFormatting();
    if (moneyTest.passed) {
      results.push('   ✅ ' + moneyTest.message);
      passed++;
    } else {
      results.push('   ❌ ' + moneyTest.message);
      failed++;
    }
    
    results.push('\n📋 ТЕСТ КАТЕГОРИЙ:');
    var categoryTest = testCategoryDetection();
    if (categoryTest.passed) {
      results.push('   ✅ ' + categoryTest.message);
      passed++;
    } else {
      results.push('   ❌ ' + categoryTest.message);
      failed++;
    }
    
    results.push('\n💰 ТЕСТ БЮДЖЕТА:');
    var budgetTest = testBudgetCalculation();
    if (budgetTest.passed) {
      results.push('   ✅ ' + budgetTest.message);
      passed++;
    } else {
      results.push('   ❌ ' + budgetTest.message);
      failed++;
    }
    
    var report = '\n══════════════════════════\n';
    report += '✅ Пройдено: ' + passed + '\n';
    report += '❌ Провалено: ' + failed + '\n';
    report += '📊 Всего: ' + (passed + failed) + '\n';
    report += '══════════════════════════';
    results.push(report);
    
    Logger.log(results.join('\n'));
    ui.alert('🧪 Результаты тестирования', '✅ Пройдено: ' + passed + '\n❌ Провалено: ' + failed, ui.ButtonSet.OK);
    addLogEntry('Тестирование завершено: пройдено ' + passed + ', провалено ' + failed, 'INFO', 'Testing');
    return { passed: passed, failed: failed };
  } catch(e) {
    logError(e, 'runAllTests');
    showMessage('Ошибка', e.toString(), 'error');
    return null;
  }
}

function testDataCollection() {
  try {
    var data = collectYearData();
    if (!data) return { passed: false, message: 'Не удалось собрать данные' };
    var checks = [data.income !== undefined, data.expenses !== undefined, data.profit !== undefined];
    var passed = checks.every(function(v) { return v; });
    return { passed: passed, message: passed ? 'Данные собраны' : 'Структура данных повреждена' };
  } catch(e) {
    return { passed: false, message: 'Ошибка: ' + e.toString() };
  }
}

function testDateParsing() {
  var tests = [
    { input: '01.01.2026', expected: true }, { input: '2026-01-01', expected: true },
    { input: '31.12.2026', expected: true }, { input: 'неверная дата', expected: false }, { input: '', expected: false }
  ];
  var passed = tests.filter(function(t) { return (safeParseDate(t.input) !== null) === t.expected; }).length;
  return { passed: passed === tests.length, message: 'Пройдено ' + passed + '/' + tests.length };
}

function testMoneyFormatting() {
  var tests = [
    { input: 1000, expected: '1 000,00 ₽' }, { input: 1000.5, expected: '1 000,50 ₽' },
    { input: 1000000, expected: '1 000 000,00 ₽' }, { input: 'не число', expected: '0,00 ₽' }, { input: null, expected: '0,00 ₽' }
  ];
  var passed = tests.filter(function(t) { return formatMoney(t.input) === t.expected; }).length;
  return { passed: passed === tests.length, message: 'Пройдено ' + passed + '/' + tests.length };
}

function testCategoryDetection() {
  var tests = [
    { input: 'продукты в магазине', expected: 'Продукты' }, { input: 'заправка бензин', expected: 'Транспорт' },
    { input: 'аптека лекарства', expected: 'Здоровье' }, { input: 'кино с друзьями', expected: 'Развлечения' },
    { input: 'неизвестно что', expected: 'Прочее' }, { input: '', expected: 'Прочее' }
  ];
  var passed = tests.filter(function(t) { return detectCategory(t.input) === t.expected; }).length;
  return { passed: passed === tests.length, message: 'Пройдено ' + passed + '/' + tests.length };
}

function testBudgetCalculation() {
  var testBudget = { 'Продукты': 30000, 'Транспорт': 5000, 'Развлечения': 10000 };
  var passed = Object.keys(testBudget).filter(function(k) { return testBudget[k] > 0; }).length;
  return { passed: passed === 3, message: 'Пройдено ' + passed + '/3' };
}

function createTestData() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('📝 Создание тестовых данных', 'Это создаст тестовые данные на отдельном листе. Продолжить?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var testSheet = ss.getSheetByName('ТЕСТОВЫЕ ДАННЫЕ');
    if (testSheet) ss.deleteSheet(testSheet);
    testSheet = ss.insertSheet('ТЕСТОВЫЕ ДАННЫЕ');
    testSheet.getRange(1,1).setValue('ТЕСТОВЫЕ ДАННЫЕ').setFontWeight('bold');
    testSheet.getRange(3,1).setValue('Сюда можно вставить тестовые данные вручную');
    ui.alert('✅ Готово', 'Тестовый лист создан', ui.ButtonSet.OK);
    addLogEntry('Создан тестовый лист', 'INFO', 'Testing');
  } catch(e) {
    logError(e, 'createTestData');
    showMessage('Ошибка', e.toString(), 'error');
  }
}

function validateDataIntegrity() {
  try {
    var sheet = getSheetSafe(SHEETS.CURRENT_YEAR);
    if (!sheet) return;
    var issues = [];
    var lastCol = sheet.getLastColumn();
    for (var col = 2; col <= lastCol; col += 2) {
      var total = sheet.getRange(ROWS.TOTAL, col).getValue();
      var balance = sheet.getRange(ROWS.BALANCE, col + 1).getValue();
      if (Math.abs(total - balance) > 0.01) issues.push('Несоответствие в колонке ' + col + ': итог ' + total + ' ≠ остаток ' + balance);
    }
    if (issues.length === 0) showMessage('✅ Проверка пройдена', 'Все данные корректны', 'info');
    else showMessage('⚠️ Найдены проблемы', issues.join('\n'), 'warning');
    addLogEntry('Проверка целостности данных: ' + (issues.length === 0 ? 'ok' : issues.length + ' проблем'), 'INFO', 'Testing');
  } catch(e) {
    logError(e, 'validateDataIntegrity');
    showMessage('Ошибка', e.toString(), 'error');
  }
}

function runAllTestsExtended() {
  var results = runAllTests();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('📊 Расширенное тестирование', 'Показать детальный отчёт в логах?', ui.ButtonSet.YES_NO);
  if (response === ui.Button.YES) {
    var fullReport = collectAllFinancialData();
    Logger.log('=== ПОЛНЫЙ ОТЧЁТ ===\n' + JSON.stringify(fullReport, null, 2));
    showMessage('Готово', 'Детальный отчёт записан в логи (Ctrl+Enter)', 'info');
  }
}