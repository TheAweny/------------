// ==================== КОНСТАНТЫ И КОНФИГУРАЦИЯ ====================

const CONFIG = {
  // Названия листов
  SHEETS: {
    INPUT: "Ввод",
    DICTIONARY: "Словарь", 
    SKEWS: "Перекосы",
    OUTPUT: "Вывод"
  },
  
  // Индексы столбцов (0-based)
  COLUMNS: {
    CLIENT_REQUEST: 0,  // A
    OIL: 1,            // B  
    ZONE: 2,           // C
    TROIKA: 3,         // D
    PE: 5,             // F
    S: 6,              // G
    COMBINATIONS: 8,    // I
    DIAGNOSTICS: 17    // R
  },
  
  // Допустимые значения для "Тройки"
  TOP_VALUES: ["1", "2", "3"],
  
  // Группы эфирных масел
  OIL_GROUPS: {
    CITRUS: "Цитрусовая",
    CONIFEROUS: "Хвойная", 
    SPICE: "Пряная",
    FLORAL: "Цветочная",
    WOODY_HERBAL: "Древесно-травяная"
  },
  
  // Зоны воздействия
  ZONES: ["+++", "+", "N", "-", "---", "0", "R"],
  
  // Сообщения
  MESSAGES: {
    INVALID_TROIKA: "Можно выбрать только 1, 2 или 3!",
    DUPLICATE_TROIKA: "Значение уже выбрано!",
    NO_DATA: "Нет данных. Заполните информацию в столбцах",
    KEY_NOT_FOUND: "Ключ не найден:"
  }
};

// ==================== КЛАССЫ ДЛЯ ВАЛИДАЦИИ И ОБРАБОТКИ ====================

/**
 * Класс для валидации и обработки редактирования ячеек
 */
class CellValidator {
  constructor(editEvent) {
    this.event = editEvent;
    this.sheet = editEvent.range.getSheet();
    this.editedCell = editEvent.range;
    this.row = this.editedCell.getRow();
    this.col = this.editedCell.getColumn();
  }
  
  /**
   * Проверяет, нужно ли обрабатывать это редактирование
   */
  shouldProcess() {
    // Игнорируем редактирование не в листе "Ввод"
    if (this.sheet.getName() !== CONFIG.SHEETS.INPUT) return false;
    
    // Игнорируем редактирование заголовка
    if (this.row === 1) return false;
    
    // Обрабатываем только столбцы Зона и Тройка
    return this.col === CONFIG.COLUMNS.ZONE + 1 || this.col === CONFIG.COLUMNS.TROIKA + 1;
  }
  
  /**
   * Выполняет валидацию и обработку
   */
  validateAndProcess() {
    if (this.col === CONFIG.COLUMNS.TROIKA + 1) {
      this.validateTroika();
    }
    
    this.updateCellFormatting();
  }
  
  /**
   * Валидация значений в столбце "Тройка"
   */
  validateTroika() {
    const value = this.editedCell.getValue();
    if (value === "" || value === null) return;
    
    const stringValue = value.toString();
    
    // Проверка допустимых значений
    if (!CONFIG.TOP_VALUES.includes(stringValue)) {
      this.showError(CONFIG.MESSAGES.INVALID_TROIKA);
      this.editedCell.clearContent();
      return;
    }
    
    // Проверка уникальности в зоне "+++"
    const zona = this.sheet.getRange(this.row, CONFIG.COLUMNS.ZONE + 1).getValue();
    if (zona === "+++") {
      this.validateUniqueTroika(stringValue);
    }
  }
  
  /**
   * Проверка уникальности значения "Тройки" в зоне "+++"
   */
  validateUniqueTroika(editedValue) {
    const lastRow = this.sheet.getLastRow();
    
    for (let r = 2; r <= lastRow; r++) {
      if (r === this.row) continue; // Пропускаем текущую строку
      
      const zona = this.sheet.getRange(r, CONFIG.COLUMNS.ZONE + 1).getValue();
      const troika = this.sheet.getRange(r, CONFIG.COLUMNS.TROIKA + 1).getValue();
      
      if (zona === "+++" && 
          troika !== "" && 
          troika !== null && 
          troika.toString() === editedValue) {
        
        this.showError(`${CONFIG.MESSAGES.DUPLICATE_TROIKA} ${editedValue}`);
        this.editedCell.clearContent();
        return;
      }
    }
  }
  
  /**
   * Обновление форматирования и валидации ячеек
   */
  updateCellFormatting() {
    const lastRow = this.sheet.getLastRow();
    
    for (let r = 2; r <= lastRow; r++) {
      const zona = this.sheet.getRange(r, CONFIG.COLUMNS.ZONE + 1).getValue();
      const troikaCell = this.sheet.getRange(r, CONFIG.COLUMNS.TROIKA + 1);
      
      if (zona === "+++") {
        this.setupTroikaValidation(troikaCell);
      } else {
        this.clearTroikaCell(troikaCell);
      }
    }
  }
  
  /**
   * Настройка валидации для ячейки "Тройка"
   */
  setupTroikaValidation(cell) {
    cell.setBackground("white");
    
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.TOP_VALUES, true)
      .setAllowInvalid(false)
      .setHelpText("Выберите 1, 2 или 3")
      .build();
      
    cell.setDataValidation(rule);
  }
  
  /**
   * Очистка ячейки "Тройка"
   */
  clearTroikaCell(cell) {
    cell.clearDataValidations();
    cell.clearContent();
    cell.setBackground("#d3d3d3");
  }
  
  /**
   * Показ ошибки пользователю
   */
  showError(message) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message);
  }
}

// ==================== КЛАССЫ ДЛЯ АВТОМАТИЧЕСКИХ СОЧЕТАНИЙ ====================

/**
 * Улучшенный класс для автоматической проверки сочетаний масел
 * Читает данные напрямую из CSV словаря и парсит номерованные сочетания
 */
class CombinationChecker {
  constructor(inputSheet) {
    this.inputSheet = inputSheet;
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    this.dictionarySheet = this.spreadsheet.getSheetByName(CONFIG.SHEETS.DICTIONARY);
    
    this.inputData = this.processInputData();
    this.combinationRules = this.loadCombinationRules();
    this.foundCombinations = [];
  }
  
  /**
   * Обработка входных данных с оптимизацией
   */
  processInputData() {
    const data = this.inputSheet.getDataRange().getValues();
    const processedData = new Map();
    
    for (let i = 1; i < data.length; i++) {
      const [clientRequest, oil, zone, troika] = data[i];
      
      if (oil && zone) {
        const key = `${oil}|${zone}`;
        if (!processedData.has(oil)) {
          processedData.set(oil, []);
        }
        processedData.get(oil).push({
          zone: zone,
          troika: troika,
          row: i + 1,
          key: key
        });
      }
    }
    
    return processedData;
  }
  
  /**
   * Загрузка правил сочетаний из CSV файла
   */
  loadCombinationRules() {
    const dictData = this.dictionarySheet.getDataRange().getValues();
    const combinationRules = new Map();
    
    for (let i = 1; i < dictData.length; i++) {
      const [oil, zone, pe, s, group, combinations, single] = dictData[i];
      
      if (combinations && combinations.trim()) {
        const key = `${oil}|${zone}`;
        const parsedCombinations = this.parseCombinations(combinations, oil, zone);
        
        if (parsedCombinations.length > 0) {
          if (!combinationRules.has(oil)) {
            combinationRules.set(oil, []);
          }
          combinationRules.get(oil).push({
            zone: zone,
            combinations: parsedCombinations,
            pe: pe,
            s: s,
            group: group
          });
        }
      }
    }
    
    return combinationRules;
  }
  
  /**
   * Продвинутый парсер сочетаний из CSV
   */
  parseCombinations(combinationsText, mainOil, mainZone) {
    const combinations = [];
    
    // Регулярное выражение для поиска номерованных сочетаний
    const regex = /\[(\d+)\]\s*([^[]+?)(?=\s*\[|\s*$)/g;
    let match;
    
    while ((match = regex.exec(combinationsText)) !== null) {
      const number = match[1];
      const description = match[2].trim();
      
      const parsedCombination = this.parseIndividualCombination(description, mainOil, mainZone, number);
      if (parsedCombination) {
        combinations.push(parsedCombination);
      }
    }
    
    return combinations;
  }
  
  /**
   * Парсинг отдельного сочетания
   */
  parseIndividualCombination(description, mainOil, mainZone, number) {
    // Извлекаем названия масел из описания
    const oilNames = this.extractOilNames(description);
    if (oilNames.length === 0) return null;
    
    // Извлекаем зоны из описания
    const zones = this.extractZones(description);
    const targetZones = zones.length > 0 ? zones : [mainZone];
    
    // Определяем логику требований (все или любое)
    const requireAll = description.includes('и ') || description.includes('сочетании с');
    const requireAny = description.includes('или ') || description.includes('/');
    
    return {
      number: number,
      oils: oilNames,
      zones: targetZones,
      description: description,
      requireAll: requireAll && !requireAny,
      requireAny: requireAny && !requireAll,
      mainOil: mainOil,
      mainZone: mainZone
    };
  }
  
  /**
   * Извлечение названий масел из описания
   */
  extractOilNames(description) {
    const oilNames = [];
    const commonOils = [
      'Апельсин', 'Бергамот', 'Лимон', 'Мандарин', 'Грейпфрут', 'Литцея Кубеба', 'Литцея',
      'Кедр', 'Кипарис', 'Ель', 'Пихта', 'Можжевеловые ягоды', 'Можжевелов',
      'Анис', 'Гвоздика', 'Мята', 'Чабрец', 'Фенхель', 'Каяпут', 'Розмарин',
      'Лаванда', 'Герань', 'Пальмароза', 'Ваниль', 'Иланг-иланг', 'Бензоин', 'Ладан',
      'Берёза', 'Аир', 'Полынь', 'Эвкалипт', 'Ветивер'
    ];
    
    for (const oil of commonOils) {
      if (description.includes(oil) && oil !== this.mainOil) {
        oilNames.push(oil);
      }
    }
    
    // Обработка специальных случаев
    if (description.includes('Литцей') || description.includes('Литцеей')) {
      oilNames.push('Литцея Кубеба');
    }
    
    return [...new Set(oilNames)]; // Удаляем дубликаты
  }
  
  /**
   * Извлечение зон из описания
   */
  extractZones(description) {
    const zones = [];
    const zonePatterns = {
      '+++': /\+\+\+/g,
      '+': /(?<!\+)\+(?!\+)/g,
      '---': /---/g,
      '-': /(?<!-)-(?!-)/g,
      'N': /\bN\b/g,
      '0': /\b0\b/g,
      'R': /\bR\b/g
    };
    
    for (const [zone, pattern] of Object.entries(zonePatterns)) {
      if (pattern.test(description)) {
        zones.push(zone);
      }
    }
    
    return zones;
  }
  
  /**
   * Основная функция проверки всех сочетаний
   */
  checkAllCombinations() {
    this.foundCombinations = [];
    
    for (const [mainOil, rules] of this.combinationRules.entries()) {
      if (this.inputData.has(mainOil)) {
        for (const rule of rules) {
          for (const combination of rule.combinations) {
            const result = this.checkCombination(mainOil, combination);
            if (result) {
              this.foundCombinations.push({
                mainOil: mainOil,
                combination: combination,
                foundOils: result.foundOils,
                priority: this.calculatePriority(combination, result.foundOils),
                rule: rule
              });
            }
          }
        }
      }
    }
    
    // Сортируем по приоритету
    this.foundCombinations.sort((a, b) => b.priority - a.priority);
    
    return this.foundCombinations;
  }
  
  /**
   * Проверка конкретного сочетания
   */
  checkCombination(mainOil, combination) {
    const foundOils = [];
    
    for (const targetOil of combination.oils) {
      if (this.inputData.has(targetOil)) {
        const oilEntries = this.inputData.get(targetOil);
        
        for (const entry of oilEntries) {
          if (combination.zones.includes(entry.zone)) {
            foundOils.push({
              oil: targetOil,
              zone: entry.zone,
              troika: entry.troika,
              row: entry.row
            });
          }
        }
      }
    }
    
    // Проверяем логику требований
    const uniqueOils = new Set(foundOils.map(f => f.oil));
    
    if (combination.requireAll) {
      return uniqueOils.size === combination.oils.length ? { foundOils } : null;
    } else if (combination.requireAny) {
      return foundOils.length > 0 ? { foundOils } : null;
    } else {
      // По умолчанию требуем все масла
      return uniqueOils.size === combination.oils.length ? { foundOils } : null;
    }
  }
  
  /**
   * Расчет приоритета сочетания для сортировки
   */
  calculatePriority(combination, foundOils) {
    let priority = 0;
    
    // Базовый приоритет по номеру сочетания
    priority += parseInt(combination.number) || 0;
    
    // Увеличиваем приоритет для критических зон
    for (const oil of foundOils) {
      if (oil.zone === '+++') priority += 100;
      else if (oil.zone === '---') priority += 80;
      else if (oil.zone === '+') priority += 60;
      else if (oil.zone === '-') priority += 40;
      
      // Дополнительный приоритет для троек
      if (oil.troika) priority += 50;
    }
    
    // Приоритет по количеству задействованных масел
    priority += foundOils.length * 10;
    
    return priority;
  }
  
  /**
   * Генерация красивого отчета по сочетаниям
   */
  getCombinationsReport() {
    if (this.foundCombinations.length === 0) {
      return this.generateEmptyReport();
    }
    
    return this.generateDetailedReport();
  }
  
  /**
   * Генерация отчета при отсутствии сочетаний
   */
  generateEmptyReport() {
    return `
╔════════════════════════════════════════════════════════════════════════════════╗
║                          🌿 АНАЛИЗ СОЧЕТАНИЙ МАСЕЛ 🌿                          ║
╚════════════════════════════════════════════════════════════════════════════════╝

📋 РЕЗУЛЬТАТ: Специальных диагностических сочетаний не обнаружено.

ℹ️  Это означает, что текущий набор масел не формирует особых комбинаций,
   требующих специального внимания или интерпретации.

🔍 Для получения более детального анализа рекомендуется проверить:
   • Правильность заполнения зон для всех масел
   • Соответствие названий масел словарю
   • Наличие достаточного количества масел для анализа
`;
  }
  
  /**
   * Генерация детального отчета
   */
  generateDetailedReport() {
    let report = `
╔════════════════════════════════════════════════════════════════════════════════╗
║                          🌿 АНАЛИЗ СОЧЕТАНИЙ МАСЕЛ 🌿                          ║
╚════════════════════════════════════════════════════════════════════════════════╝

📊 ОБНАРУЖЕНО СОЧЕТАНИЙ: ${this.foundCombinations.length}

`;
    
    // Группируем по приоритету
    const criticalCombinations = this.foundCombinations.filter(c => c.priority >= 150);
    const importantCombinations = this.foundCombinations.filter(c => c.priority >= 100 && c.priority < 150);
    const normalCombinations = this.foundCombinations.filter(c => c.priority < 100);
    
    if (criticalCombinations.length > 0) {
      report += this.generatePrioritySection("🚨 КРИТИЧЕСКИЕ СОЧЕТАНИЯ", criticalCombinations, "🔴");
    }
    
    if (importantCombinations.length > 0) {
      report += this.generatePrioritySection("⚠️  ВАЖНЫЕ СОЧЕТАНИЯ", importantCombinations, "🟡");
    }
    
    if (normalCombinations.length > 0) {
      report += this.generatePrioritySection("ℹ️  ДОПОЛНИТЕЛЬНЫЕ СОЧЕТАНИЯ", normalCombinations, "🔵");
    }
    
    report += this.generateSummarySection();
    
    return report;
  }
  
  /**
   * Генерация секции по приоритету
   */
  generatePrioritySection(title, combinations, icon) {
    let section = `
╔════════════════════════════════════════════════════════════════════════════════╗
║ ${title.padEnd(77)} ║
╚════════════════════════════════════════════════════════════════════════════════╝

`;
    
    combinations.forEach((combo, index) => {
      section += `${icon} ${index + 1}. СОЧЕТАНИЕ [${combo.combination.number}]:\n`;
      section += `   🧬 Основное масло: ${combo.mainOil}\n`;
      
      const oilsText = combo.foundOils.map(f => 
        `${f.oil} (${f.zone}${f.troika ? `, топ ${f.troika}` : ''})`
      ).join(" + ");
      section += `   🔗 Найденные масла: ${oilsText}\n`;
      
      section += `   📋 Описание: ${combo.combination.description}\n`;
      section += `   📈 Приоритет: ${combo.priority}\n`;
      section += `\n`;
    });
    
    return section;
  }
  
  /**
   * Генерация итоговой секции
   */
  generateSummarySection() {
    const totalCombinations = this.foundCombinations.length;
    const criticalCount = this.foundCombinations.filter(c => c.priority >= 150).length;
    const importantCount = this.foundCombinations.filter(c => c.priority >= 100 && c.priority < 150).length;
    
    return `
╔════════════════════════════════════════════════════════════════════════════════╗
║                                   📊 ИТОГО                                    ║
╚════════════════════════════════════════════════════════════════════════════════╝

🔴 Критические сочетания: ${criticalCount}
🟡 Важные сочетания: ${importantCount}
🔵 Дополнительные сочетания: ${totalCombinations - criticalCount - importantCount}

📋 ОБЩЕЕ КОЛИЧЕСТВО: ${totalCombinations}

💡 РЕКОМЕНДАЦИИ:
   • Уделите особое внимание критическим сочетаниям
   • Критические сочетания требуют первоочередной проработки  
   • Важные сочетания указывают на значимые процессы
   • Дополнительные сочетания дают общую картину состояния

`;
  }
  
  /**
   * Получение краткого списка для интеграции с основным отчетом
   */
  getCombinationsSummary() {
    return this.foundCombinations.map(combo => {
      const oilsText = combo.foundOils.map(f => `${f.oil} (${f.zone})`).join(" + ");
      return `[${combo.combination.number}] ${combo.mainOil} + ${oilsText}: ${combo.combination.description}`;
    });
  }
}

// ==================== КЛАССЫ ДЛЯ АНАЛИЗА ДАННЫХ ====================

/**
 * Основной класс для анализа ароматерапии
 */
class AromatherapyAnalyzer {
  constructor() {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    this.initializeSheets();
    this.dictionary = new Map();
    this.groups = this.initializeGroups();
    this.analysisData = this.initializeAnalysisData();
    
    // Кэш для улучшения производительности
    this.dictionaryCache = new Map();
    this.processingStartTime = Date.now();
  }
  
  /**
   * Инициализация листов
   */
  initializeSheets() {
    this.inputSheet = this.spreadsheet.getSheetByName(CONFIG.SHEETS.INPUT);
    this.dictionarySheet = this.spreadsheet.getSheetByName(CONFIG.SHEETS.DICTIONARY);
    this.skewsSheet = this.spreadsheet.getSheetByName(CONFIG.SHEETS.SKEWS) || 
                     this.spreadsheet.insertSheet(CONFIG.SHEETS.SKEWS);
    this.outputSheet = this.spreadsheet.getSheetByName(CONFIG.SHEETS.OUTPUT) || 
                      this.spreadsheet.insertSheet(CONFIG.SHEETS.OUTPUT);
  }
  
  /**
   * Инициализация групп масел
   */
  initializeGroups() {
    const groups = {};
    Object.values(CONFIG.OIL_GROUPS).forEach(groupName => {
      groups[groupName] = {
        oils: [],
        ...CONFIG.ZONES.reduce((acc, zone) => ({ ...acc, [zone]: 0 }), {})
      };
    });
    return groups;
  }
  
  /**
   * Инициализация данных для анализа
   */
  initializeAnalysisData() {
    return {
      neutralZoneSize: 0,
      singleOils: [],
      zeroZone: [],
      reversZone: [],
      plusPlusPlusPE: [],
      plusPlusPlusS: [],
      minusMinusMinusPE: [],
      minusMinusMinusS: [],
      plusPE: [],
      plusS: [],
      minusPE: [],
      minusS: [],
      combosAll: [],
      patterns: []
    };
  }
  
  /**
   * Оптимизированная загрузка словаря с кэшированием
   */
  loadDictionary() {
    if (this.dictionaryCache.size > 0) {
      this.dictionary = this.dictionaryCache;
      return;
    }
    
    this.showProgress("Загрузка словаря...");
    
    const dictData = this.dictionarySheet.getDataRange().getValues();
    
    // Пакетная обработка данных для улучшения производительности
    const batchSize = 50;
    for (let i = 1; i < dictData.length; i += batchSize) {
      const batch = dictData.slice(i, Math.min(i + batchSize, dictData.length));
      
      batch.forEach((row, index) => {
        const actualIndex = i + index;
        const [oil, zone, pe, s, group, combos, single] = row;
        const key = `${oil}|${zone}`;
        
        const entry = {
          pe: pe || "",
          s: s || "",
          group: group || "",
          combos: combos || "",
          single: single || ""
        };
        
        this.dictionary.set(key, entry);
      });
      
      // Небольшая пауза для предотвращения блокировки UI
      if (i % 200 === 0) {
        Utilities.sleep(10);
      }
    }
    
    // Кэшируем результат
    this.dictionaryCache = new Map(this.dictionary);
  }
  
  /**
   * Основной метод анализа с прогресс-баром
   */
  performAnalysis() {
    try {
      this.showProgress("Начинаем анализ...");
      
      this.loadDictionary();
      this.showProgress("Обработка входных данных...", 25);
      
      this.processInputData();
      this.showProgress("Поиск единичных масел...", 50);
      
      this.findSingleOils();
      this.showProgress("Генерация отчета по перекосам...", 75);
      
      this.generateSkewsReport();
      this.showProgress("Генерация итогового отчета...", 90);
      
      this.generateOutputReport();
      this.showProgress("Анализ завершен!", 100);
      
      const processingTime = (Date.now() - this.processingStartTime) / 1000;
      this.showCompletion(processingTime);
      
    } catch (error) {
      this.showError("Ошибка при выполнении анализа: " + error.message);
      throw error;
    }
  }
  
  /**
   * Обработка входных данных с батчингом
   */
  processInputData() {
    const data = this.inputSheet.getDataRange().getValues();
    const batchSize = 25;
    
    for (let row = 1; row < data.length; row += batchSize) {
      const batch = Math.min(batchSize, data.length - row);
      
      for (let i = 0; i < batch; i++) {
        const currentRow = row + i;
        if (currentRow < data.length) {
          this.processDataRow(data[currentRow], currentRow + 1);
        }
      }
      
      // Пауза для больших наборов данных
      if (data.length > 100 && row % 100 === 0) {
        Utilities.sleep(10);
      }
    }
  }
  
  /**
   * Обработка одной строки данных (оптимизировано)
   */
  processDataRow(rowData, rowIndex) {
    const [clientRequest, oil, zone, troika] = rowData;
    
    // Очистка и установка значений по умолчанию
    this.clearAndSetDefaults(rowIndex);
    
    if (!oil || !zone) return;
    
    const key = `${oil}|${zone}`;
    const entry = this.dictionary.get(key);
    
    if (entry) {
      this.processValidEntry(oil, zone, troika, entry, rowIndex);
    } else {
      this.inputSheet.getRange(rowIndex, CONFIG.COLUMNS.DIAGNOSTICS + 1)
        .setValue(`${CONFIG.MESSAGES.KEY_NOT_FOUND} ${key}`);
    }
  }
  
  /**
   * Очистка ячеек и установка значений по умолчанию
   */
  clearAndSetDefaults(rowIndex) {
    const sheet = this.inputSheet;
    const updates = [
      [rowIndex, CONFIG.COLUMNS.PE + 1, CONFIG.MESSAGES.NO_DATA],
      [rowIndex, CONFIG.COLUMNS.S + 1, CONFIG.MESSAGES.NO_DATA],
      [rowIndex, CONFIG.COLUMNS.COMBINATIONS + 1, ""],
      [rowIndex, CONFIG.COLUMNS.DIAGNOSTICS + 1, ""]
    ];
    
    // Пакетное обновление для лучшей производительности
    updates.forEach(([row, col, value]) => {
      sheet.getRange(row, col).setValue(value);
    });
  }
  
  /**
   * Обработка корректной записи из словаря
   */
  processValidEntry(oil, zone, troika, entry, rowIndex) {
    // Заполнение основных данных
    this.inputSheet.getRange(rowIndex, CONFIG.COLUMNS.PE + 1).setValue(entry.pe);
    this.inputSheet.getRange(rowIndex, CONFIG.COLUMNS.S + 1).setValue(entry.s);
    
    // Проверка комбинаций
    const combos = this.checkCombinations(oil, zone, entry.combos);
    this.inputSheet.getRange(rowIndex, CONFIG.COLUMNS.COMBINATIONS + 1).setValue(combos);
    
    // Обновление групп
    if (entry.group && this.groups[entry.group]) {
      this.groups[entry.group][zone]++;
      this.groups[entry.group].oils.push(`${oil} (${zone})`);
    }
    
    // Сбор данных для анализа
    this.collectAnalysisData(oil, zone, troika, entry);
  }
  
  /**
   * Сбор данных для алгоритма анализа 3.1
   */
  collectAnalysisData(oil, zone, troika, entry) {
    const data = this.analysisData;
    const troikaText = troika ? ` (топ ${troika})` : "";
    
    switch (zone) {
      case "N":
        data.neutralZoneSize++;
        break;
      case "0":
        data.zeroZone.push(oil);
        break;
      case "R":
        data.reversZone.push(oil);
        break;
      case "+++":
        data.plusPlusPlusPE.push(`${oil}${troikaText}: ${entry.pe}`);
        data.plusPlusPlusS.push(`${oil}${troikaText}: ${entry.s}`);
        break;
      case "---":
        data.minusMinusMinusPE.push(`${oil}: ${entry.pe}`);
        data.minusMinusMinusS.push(`${oil}: ${entry.s}`);
        break;
      case "+":
        data.plusPE.push(`${oil}: ${entry.pe}`);
        data.plusS.push(`${oil}: ${entry.s}`);
        break;
      case "-":
        data.minusPE.push(`${oil}: ${entry.pe}`);
        data.minusS.push(`${oil}: ${entry.s}`);
        break;
    }
    
    // Сбор повторяющихся закономерностей
    this.findPatterns(oil, zone, entry);
  }
  
  /**
   * Поиск повторяющихся закономерностей
   */
  findPatterns(oil, zone, entry) {
    // Базовые паттерны из словаря
    if (entry.pe.includes("*ЕД") || entry.s.includes("*ЕД")) {
      this.analysisData.patterns.push(`${oil} (${zone}): ${entry.pe} / ${entry.s}`);
    }
    
    // Паттерны по зонам из методички
    if (zone === "---") {
      const patterns = {
        "Апельсин": "Запрет на радость, глубокая депрессия.",
        "Бергамот": "Глубокая депрессия, застревание в подростковом периоде.",
        "Лимон": "Высокая раздражительность, агрессия."
      };
      if (patterns[oil]) {
        this.analysisData.patterns.push(patterns[oil]);
      }
    }
  }
  
  /**
   * Поиск единичных масел в группах
   */
  findSingleOils() {
    Object.entries(this.groups).forEach(([groupName, groupData]) => {
      CONFIG.ZONES.forEach(zone => {
        if (zone !== "oils" && groupData[zone] === 1) {
          const oilInfo = groupData.oils.find(oil => oil.includes(`(${zone})`));
          if (oilInfo) {
            const [oilName] = oilInfo.split(' (');
            const key = `${oilName}|${zone}`;
            const entry = this.dictionary.get(key);
            const singleText = entry ? entry.single : '';
            const message = singleText ? `${oilInfo}: ${singleText}` : `${oilInfo} единственное в группе ${groupName}.`;
            this.analysisData.singleOils.push(message);
          }
        }
      });
    });
  }
  
  /**
   * Проверка комбинаций масел
   */
  checkCombinations(oil, zone, baseCombos) {
    // Создаем экземпляр проверки сочетаний
    const combinationChecker = new CombinationChecker(this.inputSheet);
    
    // Проверяем все сочетания
    const foundCombinations = combinationChecker.checkAllCombinations();
    
    // Добавляем найденные сочетания в анализ
    this.analysisData.combosAll = combinationChecker.getCombinationsSummary();
    
    // Возвращаем базовые комбинации если есть
    return baseCombos || "";
  }
  
  /**
   * Генерация отчета по перекосам
   */
  generateSkewsReport() {
    const reporter = new SkewsReporter(this.skewsSheet, this.groups);
    reporter.generate();
  }
  
  /**
   * Генерация итогового отчета
   */
  generateOutputReport() {
    const reporter = new OutputReporter(this.outputSheet, this.analysisData, this.groups);
    const clientRequest = this.inputSheet.getRange("A1").getValue() || "Общий запрос";
    reporter.generate(clientRequest);
  }
  
  // ==================== UX УЛУЧШЕНИЯ ====================
  
  /**
   * Показ прогресса выполнения
   */
  showProgress(message, percent = null) {
    const progressMessage = percent !== null 
      ? `${message} (${percent}%)`
      : message;
    
    SpreadsheetApp.getActiveSpreadsheet().toast(progressMessage, "Анализ ароматерапии", 2);
  }
  
  /**
   * Показ сообщения об ошибке
   */
  showError(message) {
    SpreadsheetApp.getUi().alert('Ошибка', message, SpreadsheetApp.getUi().ButtonSet.OK);
    SpreadsheetApp.getActiveSpreadsheet().toast("❌ " + message, "Ошибка", 5);
  }
  
  /**
   * Показ сообщения о завершении
   */
  showCompletion(processingTime) {
    const message = `✅ Анализ успешно завершен за ${processingTime.toFixed(1)} сек.`;
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Готово", 3);
    
    // Дополнительная статистика
    const stats = this.generateProcessingStats();
    console.log('Статистика обработки:', stats);
  }
  
  /**
   * Генерация статистики обработки
   */
  generateProcessingStats() {
    const totalOils = Object.values(this.groups).reduce((sum, group) => 
      sum + CONFIG.ZONES.reduce((zoneSum, zone) => zoneSum + (group[zone] || 0), 0), 0);
    
    return {
      totalOils: totalOils,
      dictionaryEntries: this.dictionary.size,
      foundCombinations: this.analysisData.combosAll.length,
      singleOils: this.analysisData.singleOils.length,
      processingTime: (Date.now() - this.processingStartTime) / 1000
    };
  }
}

// ==================== КЛАССЫ ДЛЯ ГЕНЕРАЦИИ ОТЧЕТОВ ====================

/**
 * Улучшенный генератор отчета по перекосам с красивым форматированием
 */
class SkewsReporter {
  constructor(sheet, groups) {
    this.sheet = sheet;
    this.groups = groups;
  }
  
  generate() {
    this.sheet.clear();
    this.createBeautifulHeader();
    this.fillGroupAnalysis();
    this.addSummarySection();
    this.formatReport();
  }
  
  /**
   * Создание красивого заголовка
   */
  createBeautifulHeader() {
    const headerText = `
╔════════════════════════════════════════════════════════════════════════════════╗
║                         📊 АНАЛИЗ ПЕРЕКОСОВ ПО ГРУППАМ МАСЕЛ                  ║
║                          Детальная оценка дисбалансов                          ║
╚════════════════════════════════════════════════════════════════════════════════╝

💡 Этот отчет показывает распределение масел по группам и выявляет значимые перекосы,
   которые указывают на особые состояния и потребности организма.

${"═".repeat(80)}
`;
    
    const headerCell = this.sheet.getRange(1, 1, 1, 11);
    headerCell.merge().setValue(headerText);
    headerCell.setFontWeight('bold').setBackground('#e6f3ff').setWrap(true);
    
    // Создаем заголовки таблицы
    const headers = [
      "🏷️ Группа", "🚨 +++", "➕ +", "⚪ N", "➖ -", "🔴 ---", "⚫ 0", "🔄 R", 
      "🍃 Масла в группе", "🧠 ПЭ Анализ", "💊 С Анализ"
    ];
    
    this.sheet.getRange(4, 1, 1, 11).setValues([headers]);
    this.sheet.getRange(4, 1, 1, 11)
      .setFontWeight('bold')
      .setBackground('#fff2cc')
      .setBorder(true, true, true, true, true, true)
      .setHorizontalAlignment('center');
  }
  
  /**
   * Заполнение анализа групп
   */
  fillGroupAnalysis() {
    let rowIndex = 5;
    
    // Сначала рассчитываем общие показатели цитрусовой группы для анализа
    const citrusData = this.groups[CONFIG.OIL_GROUPS.CITRUS] || {};
    const citrusCount = (citrusData["+++"] || 0) + (citrusData["+"] || 0);
    const citrusNegativeCount = (citrusData["---"] || 0) + (citrusData["-"] || 0);
    
    Object.entries(this.groups).forEach(([groupName, groupData]) => {
      const [peAnalysis, sAnalysis] = this.analyzeGroupSkews(groupName, groupData, {
        citrusCount,
        citrusNegativeCount
      });
      
      const rowData = [
        `${this.getGroupIcon(groupName)} ${groupName}`,
        groupData["+++"] || 0,
        groupData["+"] || 0,
        groupData["N"] || 0,
        groupData["-"] || 0,
        groupData["---"] || 0,
        groupData["0"] || 0,
        groupData["R"] || 0,
        this.formatOilsList(groupData.oils || []),
        peAnalysis || "Сбалансированное состояние",
        sAnalysis || "В пределах нормы"
      ];
      
      this.sheet.getRange(rowIndex, 1, 1, 11).setValues([rowData]);
      
      // Применяем форматирование к строке
      this.formatDataRow(rowIndex, groupName);
      
      rowIndex++;
    });
  }
  
  /**
   * Добавление итоговой секции
   */
  addSummarySection() {
    const startRow = this.sheet.getLastRow() + 2;
    
    const summaryText = `
╔════════════════════════════════════════════════════════════════════════════════╗
║                                 📋 ИТОГОВЫЙ АНАЛИЗ                            ║
╚════════════════════════════════════════════════════════════════════════════════╝

${this.generateOverallAssessment()}

💡 КЛЮЧЕВЫЕ ВЫВОДЫ:
${this.generateKeyConclusions()}

🎯 РЕКОМЕНДАЦИИ ПО РАБОТЕ С ПЕРЕКОСАМИ:
${this.generateSkewRecommendations()}

${"─".repeat(80)}
ℹ️  Для получения детального плана коррекции обратитесь к основному отчету
`;
    
    const summaryCell = this.sheet.getRange(startRow, 1, 1, 11);
    summaryCell.merge().setValue(summaryText);
    summaryCell.setBackground('#f0f8ff').setWrap(true).setVerticalAlignment('top');
  }
  
  /**
   * Анализ перекосов группы (улучшенная версия)
   */
  analyzeGroupSkews(groupName, groupData, citrusInfo) {
    let peSkew = "";
    let sSkew = "";
    
    const totalPositive = (groupData["+++"] || 0) + (groupData["+"] || 0);
    const totalNegative = (groupData["---"] || 0) + (groupData["-"] || 0);
    const neutralCount = groupData["N"] || 0;
    const specialZones = (groupData["0"] || 0) + (groupData["R"] || 0);
    
    switch (groupName) {
      case CONFIG.OIL_GROUPS.CITRUS:
        // Цитрусовая группа - эмоциональное регулирование
        if (citrusInfo.citrusCount >= 5) {
          peSkew += "🚨 Зависимость от чужого мнения, потребность во внешнем одобрении. ";
          sSkew += "⚠️ Окислительный стресс, высокая закисленность организма. ";
        }
        if (citrusInfo.citrusNegativeCount >= 5) {
          peSkew += "🔒 Только собственное мнение, не считается с окружением. ";
          sSkew += "🩸 Хронические застойные процессы, нарушения гормонального фона. ";
        }
        if (totalPositive >= 3 && totalNegative >= 3) {
          peSkew += "⚖️ Внутренний конфликт между зависимостью и автономией. ";
        }
        break;
        
      case CONFIG.OIL_GROUPS.CONIFEROUS:
        // Хвойная группа - управление страхами и стрессом
        if (totalPositive + (groupData["0"] || 0) >= 5) {
          peSkew += "😰 Состояние паники, гиперстресс, избегание проблем. ";
          sSkew += "🔥 Острая стрессовая реакция организма. ";
        }
        if (totalNegative >= 5) {
          peSkew += "😴 Потеря чувства опасности, безразличие к угрозам. ";
          sSkew += "🧊 Подавление защитных реакций. ";
        }
        if ((groupData["-"] || 0) > 0) {
          sSkew += "🦠 Острый воспалительный процесс (требует проверки). ";
        }
        break;
        
      case CONFIG.OIL_GROUPS.SPICE:
        // Пряная группа - самооценка и пищеварение
        if (totalPositive >= 5) {
          peSkew += "🤗 Потребность в признании, тепле и заботе. ";
          sSkew += "⚡ Активация пищеварительной системы. ";
        }
        if (totalNegative >= 4) {
          peSkew += "😞 Низкая самооценка, самоуничижение. ";
          sSkew += "🍽️ Хронические нарушения ЖКТ и эндокринной системы. ";
        }
        break;
        
      case CONFIG.OIL_GROUPS.FLORAL:
        // Цветочная группа - женская энергия и гормоны
        if (neutralCount > 3) {
          peSkew += "🌸 Принятие женственности как данность без напряжения. ";
          sSkew += "⚖️ Гормональная стабильность. ";
        }
        if (totalPositive > 4) {
          peSkew += "💖 Активная работа с женской энергией. ";
          sSkew += "🌺 Активация репродуктивной системы. ";
        }
        if (totalNegative > 3) {
          peSkew += "❌ Отвержение женской роли или сексуальности. ";
          sSkew += "⚠️ Гормональные дисбалансы. ";
        }
        break;
        
      case CONFIG.OIL_GROUPS.WOODY_HERBAL:
        // Древесно-травяная группа - интеллект и детокс
        if (totalPositive > 3) {
          peSkew += "🧠 Активная интеллектуальная деятельность, творчество. ";
          sSkew += "🧹 Активные детоксикационные процессы. ";
        }
        if (totalNegative > 3) {
          peSkew += "😴 Ментальная усталость, творческий застой. ";
          sSkew += "🗑️ Накопление токсинов, нарушение очищения. ";
        }
        break;
    }
    
    // Анализ специальных зон
    if (specialZones > 2) {
      peSkew += "⚠️ Множественные конфликты и поиск решений. ";
    }
    
    return [peSkew.trim(), sSkew.trim()];
  }
  
  /**
   * Получение иконки для группы
   */
  getGroupIcon(groupName) {
    const icons = {
      [CONFIG.OIL_GROUPS.CITRUS]: "🍊",
      [CONFIG.OIL_GROUPS.CONIFEROUS]: "🌲", 
      [CONFIG.OIL_GROUPS.SPICE]: "🌶️",
      [CONFIG.OIL_GROUPS.FLORAL]: "🌸",
      [CONFIG.OIL_GROUPS.WOODY_HERBAL]: "🌿"
    };
    return icons[groupName] || "🔹";
  }
  
  /**
   * Форматирование списка масел
   */
  formatOilsList(oils) {
    if (!oils || oils.length === 0) {
      return "❌ Нет масел";
    }
    
    if (oils.length <= 3) {
      return oils.join(", ");
    }
    
    return `${oils.slice(0, 3).join(", ")} и ещё ${oils.length - 3}...`;
  }
  
  /**
   * Форматирование строки данных
   */
  formatDataRow(rowIndex, groupName) {
    const range = this.sheet.getRange(rowIndex, 1, 1, 11);
    
    // Цветовое кодирование по группам
    const groupColors = {
      [CONFIG.OIL_GROUPS.CITRUS]: "#fff2cc",      // Жёлтый
      [CONFIG.OIL_GROUPS.CONIFEROUS]: "#d5e8d4",  // Зелёный
      [CONFIG.OIL_GROUPS.SPICE]: "#f8cecc",       // Красный
      [CONFIG.OIL_GROUPS.FLORAL]: "#e1d5e7",      // Фиолетовый
      [CONFIG.OIL_GROUPS.WOODY_HERBAL]: "#dae8fc" // Голубой
    };
    
    range.setBackground(groupColors[groupName] || "#f4f4f4");
    range.setBorder(true, true, true, true, true, true);
    range.setWrap(true).setVerticalAlignment('top');
    
    // Специальное форматирование для числовых столбцов (зоны)
    for (let col = 2; col <= 8; col++) {
      const cell = this.sheet.getRange(rowIndex, col);
      const value = cell.getValue();
      
      if (value > 0) {
        if (col === 2 || col === 8) { // +++ или R
          cell.setFontWeight('bold').setFontColor('#cc0000');
        } else if (col === 5 || col === 6) { // --- или -
          cell.setFontWeight('bold').setFontColor('#0066cc');
        }
      }
    }
  }
  
  /**
   * Генерация общей оценки
   */
  generateOverallAssessment() {
    const totalOils = Object.values(this.groups).reduce((sum, group) => 
      sum + CONFIG.ZONES.reduce((zoneSum, zone) => zoneSum + (group[zone] || 0), 0), 0);
    
    const criticalZones = Object.values(this.groups).reduce((sum, group) => 
      sum + (group["+++"] || 0), 0);
    
    const resourceZones = Object.values(this.groups).reduce((sum, group) => 
      sum + (group["---"] || 0), 0);
    
    let assessment = `📊 ОБЩАЯ СТАТИСТИКА:\n`;
    assessment += `   • Всего масел в анализе: ${totalOils}\n`;
    assessment += `   • Критических зон (+++): ${criticalZones}\n`;
    assessment += `   • Ресурсных зон (---): ${resourceZones}\n`;
    assessment += `   • Соотношение проблема/ресурс: ${(criticalZones/Math.max(resourceZones,1)).toFixed(1)}:1\n\n`;
    
    if (criticalZones > resourceZones * 2) {
      assessment += "⚠️ ВЫСОКИЙ РИСК: Критических зон значительно больше ресурсных.";
    } else if (resourceZones > criticalZones) {
      assessment += "✅ ХОРОШИЙ БАЛАНС: Достаточно ресурсов для работы с проблемами.";
    } else {
      assessment += "⚖️ УМЕРЕННОЕ СОСТОЯНИЕ: Баланс между проблемами и ресурсами.";
    }
    
    return assessment;
  }
  
  /**
   * Генерация ключевых выводов
   */
  generateKeyConclusions() {
    const conclusions = [];
    
    // Анализ каждой группы
    Object.entries(this.groups).forEach(([groupName, groupData]) => {
      const totalPositive = (groupData["+++"] || 0) + (groupData["+"] || 0);
      const totalNegative = (groupData["---"] || 0) + (groupData["-"] || 0);
      
      if (totalPositive >= 4) {
        conclusions.push(`🔴 ${groupName}: Высокая активность - требует внимания`);
      }
      if (totalNegative >= 4) {
        conclusions.push(`🔵 ${groupName}: Мощные ресурсы для восстановления`);
      }
      if ((groupData["0"] || 0) + (groupData["R"] || 0) >= 2) {
        conclusions.push(`⚪ ${groupName}: Конфликтные процессы`);
      }
    });
    
    return conclusions.length > 0 
      ? conclusions.map(c => `   • ${c}`).join('\n')
      : '   ✅ Все группы в относительном балансе';
  }
  
  /**
   * Генерация рекомендаций по работе с перекосами
   */
  generateSkewRecommendations() {
    const recommendations = [
      "1️⃣ Работайте в первую очередь с группами, имеющими высокие показатели +++",
      "2️⃣ Используйте ресурсы групп с высокими показателями --- для поддержки",
      "3️⃣ Особое внимание уделите группам с множественными 0 и R зонами",
      "4️⃣ Поддерживайте баланс - не перегружайте одну группу за счёт другой",
      "5️⃣ Регулярно отслеживайте изменения в распределении масел по группам"
    ];
    
    return recommendations.map(r => `   ${r}`).join('\n');
  }
  
  /**
   * Форматирование всего отчета
   */
  formatReport() {
    const lastRow = this.sheet.getLastRow();
    
    if (lastRow > 0) {
      // Автоподбор ширины столбцов
      this.sheet.setColumnWidth(1, 180);  // Группа
      this.sheet.setColumnWidths(2, 7, 50); // Зоны
      this.sheet.setColumnWidth(9, 200);   // Масла
      this.sheet.setColumnWidth(10, 250);  // ПЭ анализ
      this.sheet.setColumnWidth(11, 250);  // С анализ
      
      // Высота строк
      this.sheet.setRowHeight(1, 80);      // Заголовок
      this.sheet.setRowHeight(4, 40);      // Заголовки таблицы
      this.sheet.setRowHeights(5, Math.max(1, lastRow - 4), 60); // Данные
      
      // Заморозка заголовков
      this.sheet.setFrozenRows(4);
      this.sheet.setFrozenColumns(1);
    }
  }
}

/**
 * Улучшенный генератор итогового отчета с красивым форматированием
 */
class OutputReporter {
  constructor(sheet, analysisData, groups) {
    this.sheet = sheet;
    this.data = analysisData;
    this.groups = groups;
    this.currentRow = 1;
  }
  
  generate(clientRequest) {
    this.sheet.clear();
    this.currentRow = 1;
    
    // Создаем красивый отчет
    this.writeHeader(clientRequest);
    this.writeExecutiveSummary();
    this.writeCombinationsSection();
    this.writeGroupAnalysis();
    this.writeKeyFindings();
    this.writeResourceAnalysis();
    this.writeDetailedBreakdown();
    this.writeRecommendations();
    this.writeFooter();
    
    this.formatReport();
  }
  
  /**
   * Запись заголовка отчета
   */
  writeHeader(clientRequest) {
    const headerText = `
╔════════════════════════════════════════════════════════════════════════════════╗
║                     🌿 ПОЛНЫЙ АНАЛИЗ АРОМАТЕРАПЕВТИЧЕСКОГО ТЕСТА 🌿            ║
║                           Алгоритм расшифровки версия 3.1                      ║
╚════════════════════════════════════════════════════════════════════════════════╝

📋 ЗАПРОС КЛИЕНТА: ${clientRequest || 'Общий анализ состояния'}
📅 ДАТА АНАЛИЗА: ${new Date().toLocaleDateString('ru-RU')}
⏰ ВРЕМЯ ГЕНЕРАЦИИ: ${new Date().toLocaleTimeString('ru-RU')}

${"═".repeat(80)}
`;
    
    this.writeFormattedText(headerText, 'header');
    this.addSpacing(2);
  }
  
  /**
   * Краткое резюме
   */
  writeExecutiveSummary() {
    const totalOils = Object.values(this.groups).reduce((sum, group) => 
      sum + CONFIG.ZONES.reduce((zoneSum, zone) => zoneSum + (group[zone] || 0), 0), 0);
    
    const criticalZoneCount = Object.values(this.groups).reduce((sum, group) => 
      sum + (group["+++"] || 0), 0);
    
    const resourceZoneCount = Object.values(this.groups).reduce((sum, group) => 
      sum + (group["---"] || 0), 0);
    
    const summaryText = `
📊 КРАТКОЕ РЕЗЮМЕ
${"─".repeat(50)}

🔢 Общее количество масел: ${totalOils}
🚨 Масла в критической зоне (+++): ${criticalZoneCount}
💪 Масла в ресурсной зоне (---): ${resourceZoneCount}
📈 Нейтральная зона (N): ${this.data.neutralZoneSize} ${this.interpretNeutralZone()}
⚠️  Особые зоны: ${"0-зона: " + (this.data.zeroZone.length || 0) + ", R-зона: " + (this.data.reversZone.length || 0)}

${this.generateQuickInsights()}
`;
    
    this.writeFormattedText(summaryText, 'summary');
    this.addSpacing(2);
  }
  
  /**
   * Секция сочетаний масел
   */
  writeCombinationsSection() {
    // Создаем новый экземпляр для получения сочетаний
    const combinationChecker = new CombinationChecker(
      SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEETS.INPUT)
    );
    
    const foundCombinations = combinationChecker.checkAllCombinations();
    const combinationsReport = combinationChecker.getCombinationsReport();
    
    this.writeFormattedText(combinationsReport, 'combinations');
    this.addSpacing(2);
  }
  
  /**
   * Анализ групп масел
   */
  writeGroupAnalysis() {
    const groupText = `
╔════════════════════════════════════════════════════════════════════════════════╗
║                            📊 АНАЛИЗ ГРУПП МАСЕЛ                              ║
╚════════════════════════════════════════════════════════════════════════════════╝
`;
    
    this.writeFormattedText(groupText, 'section-header');
    
    Object.entries(this.groups).forEach(([groupName, groupData]) => {
      this.writeGroupDetail(groupName, groupData);
    });
    
    this.addSpacing(2);
  }
  
  /**
   * Детальный анализ группы
   */
  writeGroupDetail(groupName, groupData) {
    const totalInGroup = CONFIG.ZONES.reduce((sum, zone) => sum + (groupData[zone] || 0), 0);
    const criticalCount = (groupData["+++"] || 0) + (groupData["+"] || 0);
    const resourceCount = (groupData["---"] || 0) + (groupData["-"] || 0);
    
    const groupAnalysis = this.analyzeGroupPatterns(groupName, groupData);
    
    const groupText = `
🔸 ${groupName.toUpperCase()}
${"─".repeat(60)}

📈 Распределение: +++:${groupData["+++"] || 0} | +:${groupData["+"] || 0} | N:${groupData["N"] || 0} | -:${groupData["-"] || 0} | ---:${groupData["---"] || 0} | 0:${groupData["0"] || 0} | R:${groupData["R"] || 0}
🍃 Масла в группе: ${groupData.oils.join(", ") || "Нет масел"}
📊 Всего масел: ${totalInGroup} | Активных (+/+++): ${criticalCount} | Ресурсных (-/---): ${resourceCount}

💡 ИНТЕРПРЕТАЦИЯ:
${groupAnalysis.pe ? `   🧠 Психоэмоциональный аспект: ${groupAnalysis.pe}` : ''}
${groupAnalysis.s ? `   💊 Соматический аспект: ${groupAnalysis.s}` : ''}
${groupAnalysis.patterns ? `   🔍 Особенности: ${groupAnalysis.patterns}` : ''}

`;
    
    this.writeFormattedText(groupText, 'group-detail');
  }
  
  /**
   * Ключевые находки
   */
  writeKeyFindings() {
    const findingsText = `
╔════════════════════════════════════════════════════════════════════════════════╗
║                              🎯 КЛЮЧЕВЫЕ НАХОДКИ                              ║
╚════════════════════════════════════════════════════════════════════════════════╝

🚨 ПРИОРИТЕТНЫЕ ЗАДАЧИ (зона +++):
${this.formatFindings(this.data.plusPlusPlusPE, '🧠')}
${this.formatFindings(this.data.plusPlusPlusS, '💊')}

💪 РЕСУРСЫ ДЛЯ ВОССТАНОВЛЕНИЯ (зона ---):
${this.formatFindings(this.data.minusMinusMinusPE, '🧠')}
${this.formatFindings(this.data.minusMinusMinusS, '💊')}

⚠️  ДОПОЛНИТЕЛЬНЫЕ ОБЛАСТИ ВНИМАНИЯ:
${this.formatAdditionalFindings()}

🔍 ЕДИНИЧНЫЕ МАСЛА (требуют особого внимания):
${this.formatSingleOils()}
`;
    
    this.writeFormattedText(findingsText, 'key-findings');
    this.addSpacing(2);
  }
  
  /**
   * Анализ ресурсов
   */
  writeResourceAnalysis() {
    const resourceText = `
╔════════════════════════════════════════════════════════════════════════════════╗
║                            💎 РЕСУРСНЫЙ АНАЛИЗ                                ║
╚════════════════════════════════════════════════════════════════════════════════╝

${this.analyzeResourceBalance()}

🔄 СПЕЦИАЛЬНЫЕ ЗОНЫ:
   🔴 0-зона (конфликт): ${this.data.zeroZone.join(", ") || "Нет"}
   🔄 R-зона (поиск решения): ${this.data.reversZone.join(", ") || "Нет"}
   🟡 N-зона (принятие): ${this.data.neutralZoneSize} масел - ${this.interpretNeutralZone()}

${this.generateResourceRecommendations()}
`;
    
    this.writeFormattedText(resourceText, 'resource-analysis');
    this.addSpacing(2);
  }
  
  /**
   * Детальная разбивка
   */
  writeDetailedBreakdown() {
    const detailText = `
╔════════════════════════════════════════════════════════════════════════════════╗
║                           📋 ДЕТАЛЬНАЯ РАЗБИВКА                               ║
╚════════════════════════════════════════════════════════════════════════════════╝

🧠 ПСИХОЭМОЦИОНАЛЬНОЕ СОСТОЯНИЕ:
${"─".repeat(50)}
${this.generatePsychoemotionalAnalysis()}

💊 СОМАТИЧЕСКОЕ СОСТОЯНИЕ:
${"─".repeat(50)}
${this.generateSomaticAnalysis()}

🔄 ВЫЯВЛЕННЫЕ ЗАКОНОМЕРНОСТИ:
${"─".repeat(50)}
${this.data.patterns.length > 0 ? 
  this.data.patterns.map(p => `   • ${p}`).join('\n') : 
  '   ℹ️  Специальных закономерностей не выявлено'}
`;
    
    this.writeFormattedText(detailText, 'detailed-breakdown');
    this.addSpacing(2);
  }
  
  /**
   * Рекомендации
   */
  writeRecommendations() {
    const recommendationsText = `
╔════════════════════════════════════════════════════════════════════════════════╗
║                              💡 РЕКОМЕНДАЦИИ                                  ║
╚════════════════════════════════════════════════════════════════════════════════╝

🎯 ПЕРВООЧЕРЕДНЫЕ ДЕЙСТВИЯ:
${this.generatePriorityRecommendations()}

🔧 РАБОТА С РЕСУРСАМИ:
${this.generateResourceWorkRecommendations()}

⚖️  БАЛАНС И ГАРМОНИЗАЦИЯ:
${this.generateBalanceRecommendations()}

📅 ПЛАН ДАЛЬНЕЙШИХ ДЕЙСТВИЙ:
${this.generateActionPlan()}
`;
    
    this.writeFormattedText(recommendationsText, 'recommendations');
    this.addSpacing(2);
  }
  
  /**
   * Подвал отчета
   */
  writeFooter() {
    const footerText = `
${"═".repeat(80)}
📝 ЗАКЛЮЧЕНИЕ:

${this.generateConclusion()}

${"─".repeat(80)}
ℹ️  Этот анализ основан на алгоритме расшифровки ароматерапевтического теста версии 3.1
🔬 Для более глубокого понимания рекомендуется консультация специалиста
📞 При возникновении вопросов обратитесь к вашему ароматерапевту

╔════════════════════════════════════════════════════════════════════════════════╗
║                      Конец отчета • Будьте здоровы! 🌿                        ║
╚════════════════════════════════════════════════════════════════════════════════╝
`;
    
    this.writeFormattedText(footerText, 'footer');
  }
  
  // ==================== ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ ====================
  
  /**
   * Запись форматированного текста
   */
  writeFormattedText(text, style = 'normal') {
    const cell = this.sheet.getRange(this.currentRow, 1);
    cell.setValue(text);
    
    // Применяем стили
    switch (style) {
      case 'header':
        cell.setFontWeight('bold').setFontSize(12).setBackground('#e6f3ff');
        break;
      case 'summary':
        cell.setBackground('#f0f8ff').setFontSize(11);
        break;
      case 'section-header':
        cell.setFontWeight('bold').setBackground('#fff2cc');
        break;
      case 'combinations':
        cell.setBackground('#f4f4f4').setFontFamily('Courier New');
        break;
      case 'key-findings':
        cell.setBackground('#fef7e0');
        break;
      case 'recommendations':
        cell.setBackground('#e8f5e8');
        break;
      default:
        cell.setFontSize(10);
    }
    
    cell.setWrap(true).setVerticalAlignment('top');
    this.currentRow++;
  }
  
  /**
   * Добавление отступов
   */
  addSpacing(rows = 1) {
    this.currentRow += rows;
  }
  
  /**
   * Анализ паттернов группы
   */
  analyzeGroupPatterns(groupName, groupData) {
    const analysis = { pe: '', s: '', patterns: '' };
    
    // Специфичный анализ для каждой группы
    switch (groupName) {
      case CONFIG.OIL_GROUPS.CITRUS:
        analysis.pe = this.analyzeCitrusGroupPE(groupData);
        analysis.s = this.analyzeCitrusGroupS(groupData);
        break;
      case CONFIG.OIL_GROUPS.CONIFEROUS:
        analysis.pe = this.analyzeConiferousGroupPE(groupData);
        analysis.s = this.analyzeConiferousGroupS(groupData);
        break;
      case CONFIG.OIL_GROUPS.SPICE:
        analysis.pe = this.analyzeSpiceGroupPE(groupData);
        analysis.s = this.analyzeSpiceGroupS(groupData);
        break;
      case CONFIG.OIL_GROUPS.FLORAL:
        analysis.pe = this.analyzeFloralGroupPE(groupData);
        analysis.s = this.analyzeFloralGroupS(groupData);
        break;
      case CONFIG.OIL_GROUPS.WOODY_HERBAL:
        analysis.pe = this.analyzeWoodyHerbalGroupPE(groupData);
        analysis.s = this.analyzeWoodyHerbalGroupS(groupData);
        break;
    }
    
    return analysis;
  }
  
  /**
   * Анализ цитрусовой группы (ПЭ)
   */
  analyzeCitrusGroupPE(groupData) {
    const positiveCount = (groupData["+++"] || 0) + (groupData["+"] || 0);
    const negativeCount = (groupData["---"] || 0) + (groupData["-"] || 0);
    
    let analysis = '';
    
    if (positiveCount >= 5) {
      analysis += 'Зависимость от чужого мнения, потребность во внешнем одобрении. ';
    }
    
    if (negativeCount >= 5) {
      analysis += 'Только собственное мнение, сложности с учетом мнения окружения. ';
    }
    
    if (positiveCount >= 3 && negativeCount >= 3) {
      analysis += 'Внутренний конфликт между зависимостью от мнения и стремлением к автономии. ';
    }
    
    return analysis || 'Сбалансированное состояние по цитрусовой группе.';
  }
  
  /**
   * Анализ цитрусовой группы (С)
   */
  analyzeCitrusGroupS(groupData) {
    const positiveCount = (groupData["+++"] || 0) + (groupData["+"] || 0);
    const negativeCount = (groupData["---"] || 0) + (groupData["-"] || 0);
    
    let analysis = '';
    
    if (positiveCount >= 5) {
      analysis += 'Окислительный стресс, высокая закисленность организма. ';
    }
    
    if (negativeCount >= 5) {
      analysis += 'Хронические застойные процессы, возможны гормональные нарушения. ';
    }
    
    return analysis || 'Нормальное состояние кислотно-щелочного баланса.';
  }
  
  /**
   * Анализ хвойной группы (ПЭ)
   */
  analyzeConiferousGroupPE(groupData) {
    const totalPositive = (groupData["+++"] || 0) + (groupData["+"] || 0) + (groupData["0"] || 0);
    const negativeCount = (groupData["---"] || 0) + (groupData["-"] || 0);
    
    let analysis = '';
    
    if (totalPositive === 5) {
      analysis += 'Состояние паники, гиперстресс, стремление избежать проблем. ';
    }
    
    if (negativeCount >= 5) {
      analysis += 'Потеря чувства опасности, безразличие к угрозам. ';
    }
    
    return analysis || 'Адекватное восприятие внешних угроз и стрессов.';
  }
  
  /**
   * Анализ хвойной группы (С)
   */
  analyzeConiferousGroupS(groupData) {
    if ((groupData["-"] || 0) > 0) {
      return 'Острый воспалительный процесс, требует проверки основной соматики.';
    }
    return 'Нормальное состояние противовоспалительной системы.';
  }
  
  /**
   * Интерпретация нейтральной зоны
   */
  interpretNeutralZone() {
    return this.data.neutralZoneSize > 3 
      ? 'принятие как данность, стабильное состояние'
      : 'напряжение, неопределенность';
  }
  
  /**
   * Форматирование находок
   */
  formatFindings(findings, icon) {
    if (!findings || findings.length === 0) {
      return `${icon} Нет специфических находок в данной области`;
    }
    
    return findings.map(finding => `${icon} ${finding}`).join('\n');
  }
  
  /**
   * Дополнительные находки
   */
  formatAdditionalFindings() {
    let additional = '';
    
    if (this.data.plusPE.length > 0) {
      additional += `   🧠 Зона "+": ${this.data.plusPE.slice(0, 3).join(', ')}${this.data.plusPE.length > 3 ? '...' : ''}\n`;
    }
    
    if (this.data.minusPE.length > 0) {
      additional += `   🧠 Зона "-": ${this.data.minusPE.slice(0, 3).join(', ')}${this.data.minusPE.length > 3 ? '...' : ''}\n`;
    }
    
    return additional || '   ℹ️  Дополнительных областей не выявлено';
  }
  
  /**
   * Форматирование единичных масел
   */
  formatSingleOils() {
    return this.data.singleOils.length > 0 
      ? this.data.singleOils.map(oil => `   • ${oil}`).join('\n')
      : '   ℹ️  Единичных масел не обнаружено';
  }
  
  /**
   * Генерация быстрых инсайтов
   */
  generateQuickInsights() {
    const insights = [];
    
    // Рассчитываем показатели
    const criticalZoneCount = Object.values(this.groups).reduce((sum, group) => 
      sum + (group["+++"] || 0), 0);
    
    const resourceZoneCount = Object.values(this.groups).reduce((sum, group) => 
      sum + (group["---"] || 0), 0);
    
    if (criticalZoneCount > 7) {
      insights.push('🚨 Высокий уровень стресса - требуется комплексная работа');
    }
    
    if (resourceZoneCount > 5) {
      insights.push('💪 Хороший ресурсный потенциал для восстановления');
    }
    
    if (this.data.neutralZoneSize > 5) {
      insights.push('✅ Стабильное принятие многих аспектов жизни');
    }
    
    return insights.length > 0 
      ? `\n💡 КЛЮЧЕВЫЕ ИНСАЙТЫ:\n${insights.map(i => `   ${i}`).join('\n')}`
      : '';
  }
  
  /**
   * Анализ баланса ресурсов
   */
  analyzeResourceBalance() {
    const totalCritical = Object.values(this.groups).reduce((sum, group) => 
      sum + (group["+++"] || 0) + (group["+"] || 0), 0);
    
    const totalResource = Object.values(this.groups).reduce((sum, group) => 
      sum + (group["---"] || 0) + (group["-"] || 0), 0);
    
    const ratio = totalResource / (totalCritical || 1);
    
    let analysis = `📊 Соотношение проблемных и ресурсных зон: ${totalCritical}:${totalResource}\n`;
    
    if (ratio > 1.2) {
      analysis += '✅ Отличный ресурсный потенциал - организм готов к восстановлению';
    } else if (ratio > 0.8) {
      analysis += '⚖️  Сбалансированное состояние - есть ресурсы для работы с проблемами';
    } else if (ratio > 0.5) {
      analysis += '⚠️  Ограниченные ресурсы - требуется осторожный подход';
    } else {
      analysis += '🚨 Дефицит ресурсов - приоритет на восстановление сил';
    }
    
    return analysis;
  }
  
  /**
   * Генерация психоэмоционального анализа
   */
  generatePsychoemotionalAnalysis() {
    let analysis = '';
    
    if (this.data.plusPlusPlusPE.length > 0) {
      analysis += '🚨 ОСТРЫЕ ПРОБЛЕМЫ:\n';
      analysis += this.data.plusPlusPlusPE.map(item => `   • ${item}`).join('\n') + '\n\n';
    }
    
    if (this.data.minusMinusMinusPE.length > 0) {
      analysis += '💪 РЕСУРСЫ:\n';
      analysis += this.data.minusMinusMinusPE.map(item => `   • ${item}`).join('\n') + '\n\n';
    }
    
    return analysis || 'ℹ️  Стабильное психоэмоциональное состояние';
  }
  
  /**
   * Генерация соматического анализа
   */
  generateSomaticAnalysis() {
    let analysis = '';
    
    if (this.data.plusPlusPlusS.length > 0) {
      analysis += '🚨 ТРЕБУЕТ ВНИМАНИЯ:\n';
      analysis += this.data.plusPlusPlusS.map(item => `   • ${item}`).join('\n') + '\n\n';
    }
    
    if (this.data.minusMinusMinusS.length > 0) {
      analysis += '💊 ПОДДЕРЖИВАЮЩИЕ ФАКТОРЫ:\n';
      analysis += this.data.minusMinusMinusS.map(item => `   • ${item}`).join('\n') + '\n\n';
    }
    
    return analysis || 'ℹ️  Соматические показатели в пределах нормы';
  }
  
  /**
   * Генерация рекомендаций по приоритетам
   */
  generatePriorityRecommendations() {
    const recommendations = [];
    
    if (this.data.plusPlusPlusPE.length > 3) {
      recommendations.push('1. Работа с психоэмоциональным состоянием - высший приоритет');
    }
    
    if (this.data.plusPlusPlusS.length > 3) {
      recommendations.push('2. Медицинское обследование по выявленным соматическим проблемам');
    }
    
    if (this.data.singleOils.length > 0) {
      recommendations.push('3. Особое внимание единичным маслам - они показывают ключевые проблемы');
    }
    
    return recommendations.length > 0 
      ? recommendations.map(r => `   ${r}`).join('\n')
      : '   ✅ Состояние стабильное, специальных приоритетов нет';
  }
  
  /**
   * Генерация рекомендаций по работе с ресурсами
   */
  generateResourceWorkRecommendations() {
    const recommendations = [];
    
    if (this.data.minusMinusMinusPE.length > 0) {
      recommendations.push('• Активизируйте выявленные психоэмоциональные ресурсы');
    }
    
    if (this.data.neutralZoneSize > 3) {
      recommendations.push('• Используйте стабильные зоны как опору для изменений');
    }
    
    return recommendations.length > 0 
      ? recommendations.map(r => `   ${r}`).join('\n')
      : '   ⚠️  Ресурсы ограничены - фокус на восстановлении';
  }
  
  /**
   * Генерация рекомендаций по балансу
   */
  generateBalanceRecommendations() {
    return `   • Работайте постепенно, не перегружайте организм
   • Соблюдайте баланс между активностью и отдыхом  
   • Регулярно отслеживайте изменения в состоянии
   • При необходимости повторите тестирование через 2-4 недели`;
  }
  
  /**
   * Генерация плана действий
   */
  generateActionPlan() {
    return `   1️⃣ Ближайшие 7 дней: работа с критическими сочетаниями
   2️⃣ Следующие 2-3 недели: активизация ресурсных зон
   3️⃣ Через месяц: контрольное тестирование
   4️⃣ При улучшении: переход к профилактической работе`;
  }
  
  /**
   * Генерация заключения
   */
  generateConclusion() {
    const totalIssues = this.data.plusPlusPlusPE.length + this.data.plusPlusPlusS.length;
    const totalResources = this.data.minusMinusMinusPE.length + this.data.minusMinusMinusS.length;
    
    if (totalIssues > 7 && totalResources < 3) {
      return 'Состояние требует комплексного подхода и профессиональной поддержки. Рекомендуется начать с восстановления ресурсов.';
    } else if (totalIssues > 3 && totalResources > 3) {
      return 'Сбалансированная ситуация с хорошим потенциалом для улучшения. Есть ресурсы для работы с выявленными проблемами.';
    } else if (totalIssues < 3) {
      return 'Стабильное состояние. Рекомендуется профилактическая работа и поддержание текущего баланса.';
    } else {
      return 'Умеренные отклонения, которые можно скорректировать при правильном подходе.';
    }
  }
  
  /**
   * Форматирование отчета
   */
  formatReport() {
    const lastRow = this.sheet.getLastRow();
    if (lastRow > 0) {
      this.sheet.getRange(1, 1, lastRow, 1).setWrap(true).setVerticalAlignment('top');
      this.sheet.setColumnWidth(1, 900);
      this.sheet.setRowHeights(1, lastRow, 25);
    }
  }
  
  // Добавляем недостающие методы анализа групп
  analyzeSpiceGroupPE(groupData) {
    const positiveCount = (groupData["+++"] || 0) + (groupData["+"] || 0);
    const negativeCount = (groupData["---"] || 0) + (groupData["-"] || 0);
    
    if (positiveCount === 5) {
      return 'Потребность в признании, тепле и заботе.';
    }
    
    return 'Сбалансированное состояние по пряной группе.';
  }
  
  analyzeSpiceGroupS(groupData) {
    const negativeCount = (groupData["---"] || 0) + (groupData["-"] || 0);
    
    if (negativeCount >= 4) {
      return 'Хронические нарушения ЖКТ и эндокринной системы.';
    }
    
    return 'Нормальное состояние пищеварительной системы.';
  }
  
  analyzeFloralGroupPE(groupData) {
    if ((groupData["N"] || 0) > 3) {
      return 'Принятие женственности как данность без напряжения.';
    }
    
    return 'Работа с женской энергией и принятием.';
  }
  
  analyzeFloralGroupS(groupData) {
    return 'Состояние гормональной и репродуктивной системы.';
  }
  
  analyzeWoodyHerbalGroupPE(groupData) {
    return 'Творческие и интеллектуальные процессы.';
  }
  
  analyzeWoodyHerbalGroupS(groupData) {
    return 'Детоксикационные и очистительные процессы.';
  }
}

// ==================== ОСНОВНЫЕ ОБРАБОТЧИКИ СОБЫТИЙ ====================

/**
 * Главный обработчик редактирования ячеек
 */
function onEdit(e) {
  try {
    const validator = new CellValidator(e);
    if (!validator.shouldProcess()) return;
    
    validator.validateAndProcess();
    updateAnalysis();
    
  } catch (error) {
    console.error('Ошибка в onEdit:', error);
    SpreadsheetApp.getUi().alert('Произошла ошибка: ' + error.message);
  }
}

/**
 * Принудительное обновление анализа (можно вызвать из меню)
 */
function forceUpdateAnalysis() {
  updateAnalysis();
}

// ==================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ====================

/**
 * Основная функция обновления анализа
 */
function updateAnalysis() {
  try {
    const analyzer = new AromatherapyAnalyzer();
    analyzer.performAnalysis();
    
    SpreadsheetApp.getActiveSpreadsheet().toast("Анализ успешно обновлен!", "Готово", 3);
    
  } catch (error) {
    console.error('Ошибка в updateAnalysis:', error);
    SpreadsheetApp.getUi().alert('Ошибка при обновлении анализа: ' + error.message);
  }
}

/**
 * Создание меню для ручного запуска функций
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🌿 Ароматерапия')
    .addItem('🔄 Обновить анализ', 'forceUpdateAnalysis')
    .addItem('🧹 Очистить форматирование', 'clearAllFormatting')
    .addItem('ℹ️ Справка', 'showHelp')
    .addToUi();
}

/**
 * Очистка форматирования (полезно для отладки)
 */
function clearAllFormatting() {
  const sheet = SpreadsheetApp.getActiveSheet();
  if (sheet.getName() === CONFIG.SHEETS.INPUT) {
    const range = sheet.getRange(2, CONFIG.COLUMNS.TROIKA + 1, sheet.getLastRow() - 1, 1);
    range.clearDataValidations();
    range.setBackground("white");
    SpreadsheetApp.getActiveSpreadsheet().toast("Форматирование очищено", "Готово", 2);
  }
}

/**
 * Показ справки
 */
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  const helpText = `
СПРАВКА ПО СИСТЕМЕ АНАЛИЗА АРОМАТЕРАПИИ

ОСНОВНОЕ НАЗНАЧЕНИЕ:
Система помогает автоматически анализировать данные по эфирным маслам,
выявлять ключевые зоны, перекосы и взаимосвязи, а также формировать
подробный отчет для интерпретации.

СТРУКТУРА ЛИСТОВ:
• "Ввод" — введите данные для анализа
• "Словарь" — справочник масел и их свойств  
• "Перекосы" — автоматический анализ дисбалансов по группам
• "Вывод" — итоговый отчет с интерпретацией

КАК ПОЛЬЗОВАТЬСЯ:
1. Перейдите на лист "Ввод" и заполните данные (масло, зона, тройка).
2. Система автоматически проверит корректность ввода.
3. Анализ и отчеты будут обновлены автоматически.
4. При необходимости используйте меню 🌿 Ароматерапия → «Обновить анализ».

ВОЗМОЖНОСТИ СИСТЕМЫ:
• Автоматическая проверка и валидация ввода
• Подсветка ошибок и предупреждений
• Автоматический анализ сочетаний масел
• Генерация отчетов по алгоритму интерпретации
• Автоматическое форматирование для удобства чтения

НОВЫЕ ВОЗМОЖНОСТИ:
• Автоматическая индексация всех масел по зонам
• Автоматическая проверка всех возможных сочетаний
• Структурированный вывод результатов сочетаний
• Улучшенная читаемость отчетов
  `;
  
  ui.alert('Справка по системе', helpText, ui.ButtonSet.OK);
}

// ==================== ФУНКЦИИ ДЛЯ ОТЛАДКИ И ТЕСТИРОВАНИЯ ====================

/**
 * Функция для тестирования системы (только для разработки)
 */
function runTests() {
  console.log('Запуск тестов системы анализа ароматерапии...');
  
  try {
    // Тест инициализации
    const analyzer = new AromatherapyAnalyzer();
    console.log('✓ Инициализация прошла успешно');
    
    // Тест загрузки словаря
    analyzer.loadDictionary();
    console.log(`✓ Словарь загружен, записей: ${analyzer.dictionary.size}`);
    
    // Тест валидатора
    console.log('✓ Все тесты пройдены успешно');
    
  } catch (error) {
    console.error('✗ Ошибка в тестах:', error);
  }
}

/**
 * Функция для экспорта конфигурации
 */
function exportConfig() {
  console.log('Текущая конфигурация системы:', JSON.stringify(CONFIG, null, 2));
}

// END
