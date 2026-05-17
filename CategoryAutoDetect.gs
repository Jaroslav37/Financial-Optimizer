/**
 * АВТООПРЕДЕЛЕНИЕ КАТЕГОРИЙ – по истории последних 30 дней
 */

class CategoryAutoDetect {
  constructor() {
    this.expenseSheet = getSheetSafe(SHEETS.EXPENSE_DETAILS);
    this.incomeSheet  = getSheetSafe(SHEETS.INCOME_DETAILS);
    this.cacheDays = 30;
  }

  /** Возвращает {category, subcategory} или null */
  detect(type, description) {
    if (!description || description.length < 2) return null;

    const desc = description.toLowerCase().trim();
    const sheet = type === 'income' ? this.incomeSheet : this.expenseSheet;
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    const cutoff = new Date();
    cutoff.setDate(cutoff.getDate() - this.cacheDays);

    // Собираем статистику: описание → {категория → счёт, подкатегория → счёт}
    const stats = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const date = row[0];
      if (!date || date < cutoff) continue;

      const rowDesc = (row[6] || row[4] || '').toString().toLowerCase().trim();
      if (!rowDesc) continue;

      const category = (row[4] || row[2] || '').toString().trim();
      const subcategory = (row[3] || '').toString().trim();

      // Простое сравнение: если описание содержит введённый текст или наоборот
      if (rowDesc.indexOf(desc) !== -1 || desc.indexOf(rowDesc) !== -1) {
        if (!stats[category]) stats[category] = { count: 0, subs: {} };
        stats[category].count++;
        if (subcategory) {
          if (!stats[category].subs[subcategory]) stats[category].subs[subcategory] = 0;
          stats[category].subs[subcategory]++;
        }
      }
    }

    // Выбираем самую частую категорию
    let bestCat = null, bestCount = 0;
    for (const cat in stats) {
      if (stats[cat].count > bestCount) {
        bestCount = stats[cat].count;
        bestCat = cat;
      }
    }
    if (!bestCat) return null;

    // Выбираем самую частую подкатегорию внутри категории
    const subs = stats[bestCat].subs;
    let bestSub = null, bestSubCount = 0;
    for (const sub in subs) {
      if (subs[sub] > bestSubCount) {
        bestSubCount = subs[sub];
        bestSub = sub;
      }
    }

    return { category: bestCat, subcategory: bestSub || '' };
  }
}

function autoDetectCategory(type, description) {
  return new CategoryAutoDetect().detect(type, description);
}