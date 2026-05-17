/**
 * МОДУЛЬ ВВОДА ДАННЫХ - проверка пропусков и навигация по дням
 * Версия 3.0 – использует унифицированную форму «День за минуту»
 */

class DataEntryModule {
  constructor() {
    this.sheet = getSheetSafe(SHEETS.CURRENT_YEAR);
    this.today = new Date();
    this.currentYear = CONFIG.CURRENT_YEAR;

    this.cache = {
      headers: null,
      data: null,
      lastCol: 0,
      lastRow: 0
    };
  }

  /**
   * Загружает данные одним запросом (оптимизация)
   */
  loadSheetData() {
    if (!this.sheet) return false;

    try {
      this.cache.lastCol = this.sheet.getLastColumn();
      this.cache.lastRow = this.sheet.getLastRow();

      var range = this.sheet.getRange(1, 1, Math.max(this.cache.lastRow, ROWS.DATA_END), this.cache.lastCol);
      var values = range.getValues();

      this.cache.headers = values[0];
      this.cache.data = values;

      return true;
    } catch (e) {
      logError(e, 'DataEntryModule.loadSheetData');
      return false;
    }
  }

  /**
   * Находит последнюю дату, на которую есть данные
   */
  findLastDate() {
    try {
      if (!this.sheet) return null;

      if (!this.cache.data) {
        this.loadSheetData();
      }

      var lastDate = null;
      var data = this.cache.data;

      for (var col = 1; col <= 24; col++) {
        var header = this.cache.headers[col];
        if (!header) continue;

        var month = -1;
        for (var m = 0; m < MONTHS.length; m++) {
          if (header.indexOf(MONTHS[m]) !== -1) {
            month = m;
            break;
          }
        }
        if (month === -1) continue;

        var daysInMonth = new Date(this.currentYear, month + 1, 0).getDate();

        for (var row = 1; row <= 31; row++) {
          var value = data[row] ? data[row][col] : null;

          if (value && value !== 0 && value !== '') {
            var day = row;
            if (day <= daysInMonth) {
              var date = new Date(this.currentYear, month, day);

              if (!lastDate || date > lastDate) {
                lastDate = date;
              }
            }
          }
        }
      }

      return lastDate;
    } catch (e) {
      logError(e, 'DataEntryModule.findLastDate');
      return null;
    }
  }

  /**
   * Проверка пропущенных дней и предложение ввести данные
   */
  checkMissingData() {
    security.checkAccess('checkMissingData');

    try {
      if (!this.sheet) return;

      var lastDate = this.findLastDate();

      if (!lastDate) {
        var ui = SpreadsheetApp.getUi();
        var response = ui.alert(
          'Начало ввода данных',
          'В таблице нет данных. Начать ввод с 1 января ' + this.currentYear + '?',
          ui.ButtonSet.YES_NO
        );

        if (response === ui.Button.YES) {
          var startDate = new Date(this.currentYear, 0, 1);
          this.openUnifiedForm(startDate);
        }
        return;
      }

      var todayThisYear = new Date(this.currentYear, this.today.getMonth(), this.today.getDate());
      var diffDays = Math.ceil((todayThisYear - lastDate) / (1000 * 60 * 60 * 24));

      if (diffDays <= 1) {
        showMessage('Информация', 'Данные актуальны. Последние данные от ' + formatDate(lastDate), 'info');
        return;
      }

      var ui = SpreadsheetApp.getUi();
      var response = ui.alert(
        'Проверка данных',
        'Последние данные от ' + formatDate(lastDate) + '.\n\nПропущено ' + (diffDays - 1) + ' дней.\n\nХотите ввести данные за пропущенные дни?',
        ui.ButtonSet.YES_NO
      );

      if (response === ui.Button.YES) {
        var nextDate = new Date(lastDate);
        nextDate.setDate(nextDate.getDate() + 1);
        this.openUnifiedForm(nextDate);
      }
    } catch (e) {
      logError(e, 'DataEntryModule.checkMissingData');
      security.secureLog('Error in checkMissingData', { error: e.toString() });
      showMessage('Ошибка', e.toString(), 'error');
    }
  }

  /**
   * Открывает унифицированную форму «День за минуту» на указанную дату
   * @param {Date} date - дата, для которой вводятся данные
   */
  openUnifiedForm(date) {
    // Вызываем глобальную функцию из UnifiedDayForm.gs
    showUnifiedDayFormForDate(date);
  }

  /**
   * Ввод данных за сегодня
   */
  enterTodayData() {
    var todayThisYear = new Date(this.currentYear, this.today.getMonth(), this.today.getDate());
    this.openUnifiedForm(todayThisYear);
  }

  /**
   * Ввод данных за выбранную дату
   */
  enterCustomDate() {
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt(
      'Ввод данных',
      'Введите дату в формате ДД.ММ.ГГГГ:',
      ui.ButtonSet.OK_CANCEL
    );

    if (response.getSelectedButton() === ui.Button.OK) {
      var dateStr = response.getResponseText();
      var date = safeParseDate(dateStr);

      if (date) {
        if (date.getFullYear() !== this.currentYear) {
          ui.alert('Ошибка', 'Год должен быть ' + this.currentYear, ui.ButtonSet.OK);
          return;
        }
        this.openUnifiedForm(date);
      } else {
        ui.alert('Ошибка', 'Неверный формат даты', ui.ButtonSet.OK);
      }
    }
  }
}

// =============================================
// ГЛОБАЛЬНЫЕ ФУНКЦИИ ДЛЯ ВЫЗОВА ИЗ МЕНЮ
// =============================================

function checkMissingData() {
  try {
    new DataEntryModule().checkMissingData();
  } catch (e) {
    logError(e, 'checkMissingData');
    showMessage('Ошибка', e.toString(), 'error');
  }
}

function enterTodayData() {
  try {
    new DataEntryModule().enterTodayData();
  } catch (e) {
    logError(e, 'enterTodayData');
    showMessage('Ошибка', e.toString(), 'error');
  }
}

function enterCustomDate() {
  try {
    new DataEntryModule().enterCustomDate();
  } catch (e) {
    logError(e, 'enterCustomDate');
    showMessage('Ошибка', e.toString(), 'error');
  }
}