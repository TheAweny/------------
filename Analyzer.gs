// ==================== ОСНОВНОЙ АНАЛИЗАТОР ====================

/**
 * Главный класс для анализа ароматерапии
 * Координирует все компоненты системы и управляет процессом анализа
 */
class AromatherapyAnalyzer {
  constructor() {
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    this.sheets = this.initializeSheets();
    this.dictionary = new Map();
    this.oilGroups = this.initializeOilGroups();
    this.analysisResults = this.initializeAnalysisResults();
  }
  
  /**
   * Инициализирует все необходимые листы
   * @returns {Object} Объект с ссылками на листы
   */
  initializeSheets() {
    return {
      input: this.spreadsheet.getSheetByName(CONFIG.SHEETS.INPUT),
      dictionary: this.spreadsheet.getSheetByName(CONFIG.SHEETS.DICTIONARY),
      skews: this.spreadsheet.getSheetByName(CONFIG.SHEETS.SKEWS) || 
             this.spreadsheet.insertSheet(CONFIG.SHEETS.SKEWS),
      output: this.spreadsheet.getSheetByName(CONFIG.SHEETS.OUTPUT) || 
              this.spreadsheet.insertSheet(CONFIG.SHEETS.OUTPUT)
    };
  }
  
  /**
   * Инициализирует группы масел
   * @returns {Object} Объект групп масел
   */
  initializeOilGroups() {
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
   * Инициализирует структуру результатов анализа
   * @returns {Object} Объект результатов анализа
   */
  initializeAnalysisResults() {
    return {
      neutralZoneSize: 0,
      specialZones: {
        zero: [],
        reverse: []
      },
      mainTasks: {
        plusPlusPlusPE: [],
        plusPlusPlusS: [],
        minusMinusMinusPE: [],
        minusMinusMinusS: []
      },
      additionalTasks: {
        plusPE: [],
        plusS: [],
        minusPE: [],
        minusS: []
      },
      singleOils: [],
      combinations: [],
      patterns: []
    };
  }
  
  /**
   * Загружает словарь масел
   */
  loadDictionary() {
    try {
      const dictData = this.sheets.dictionary.getDataRange().getValues();
      
      for (let i = 1; i < dictData.length; i++) {
        const [oil, zone, pe, s, group, combos, single] = dictData[i];
        if (Utils.isNotEmpty(oil) && Utils.isNotEmpty(zone)) {
          const key = `${oil}|${zone}`;
          const parsedCombos = Utils.parseCombos(combos || "");
          this.dictionary.set(key, {
            pe: pe || "",
            s: s || "",
            group: group || "",
            combos: combos || "",
            combosIndex: parsedCombos.index,
            single: single || ""
          });
        }
      }
    } catch (error) {
      console.error(`Ошибка загрузки словаря: ${error.message}`);
    }
  }
  
  /**
   * Выполняет полный анализ
   */
  performAnalysis() {
    try {
      this.loadDictionary();
      this.processInputData();
      this.findSingleOils();
      this.analyzeCombinations();
      this.generateReports();
      
      SpreadsheetApp.getActiveSpreadsheet().toast(CONFIG.MESSAGES.ANALYSIS_COMPLETE, "Готово", 3);
    } catch (error) {
      console.error(`Ошибка анализа: ${error.message}`);
      SpreadsheetApp.getUi().alert(`${CONFIG.MESSAGES.ANALYSIS_ERROR} ${error.message}`);
    }
  }
  
  /**
   * Обрабатывает данные с листа ввода
   */
  processInputData() {
    const data = this.sheets.input.getDataRange().getValues();
    
    for (let row = 1; row < data.length; row++) {
      this.processDataRow(data[row], row + 1);
    }
  }
  
  /**
   * Обрабатывает одну строку данных
   * @param {Array} rowData - Данные строки
   * @param {number} rowIndex - Индекс строки (1-based)
   */
  processDataRow(rowData, rowIndex) {
    const [clientRequest, oil, zone, troika] = rowData.map(cell => cell ? cell.toString() : "");
    
    // Очистка ячеек
    this.clearRowCells(rowIndex);
    
    if (!Utils.isNotEmpty(oil) || !Utils.isNotEmpty(zone)) return;
    
    const key = `${oil}|${zone}`;
    const entry = this.dictionary.get(key);
    
    if (entry) {
      this.processValidEntry(oil, zone, troika, entry, rowIndex);
    } else {
      Utils.setCellValue(
        this.sheets.input.getRange(rowIndex, CONFIG.COLUMNS.DIAGNOSTICS + 1),
        `${CONFIG.MESSAGES.KEY_NOT_FOUND} ${key}`
      );
    }
  }
  
  /**
   * Очищает ячейки строки и устанавливает значения по умолчанию
   * @param {number} rowIndex - Индекс строки
   */
  clearRowCells(rowIndex) {
    const sheet = this.sheets.input;
    Utils.setCellValue(sheet.getRange(rowIndex, CONFIG.COLUMNS.PE + 1), CONFIG.MESSAGES.NO_DATA);
    Utils.setCellValue(sheet.getRange(rowIndex, CONFIG.COLUMNS.S + 1), CONFIG.MESSAGES.NO_DATA);
    Utils.setCellValue(sheet.getRange(rowIndex, CONFIG.COLUMNS.COMBINATIONS + 1), "");
    Utils.setCellValue(sheet.getRange(rowIndex, CONFIG.COLUMNS.DIAGNOSTICS + 1), "");
  }
  
  /**
   * Обрабатывает корректную запись из словаря
   * @param {string} oil - Название масла
   * @param {string} zone - Зона воздействия
   * @param {string} troika - Значение тройки
   * @param {Object} entry - Запись из словаря
   * @param {number} rowIndex - Индекс строки
   */
  processValidEntry(oil, zone, troika, entry, rowIndex) {
    // Заполнение основных данных
    const expandedPe = Utils.expandCombinationReferences(entry.pe, entry.combosIndex);
    const expandedS = Utils.expandCombinationReferences(entry.s, entry.combosIndex);
    Utils.setCellValue(this.sheets.input.getRange(rowIndex, CONFIG.COLUMNS.PE + 1), expandedPe);
    Utils.setCellValue(this.sheets.input.getRange(rowIndex, CONFIG.COLUMNS.S + 1), expandedS);
    
    // Обновление групп масел
    if (entry.group && this.oilGroups[entry.group]) {
      this.oilGroups[entry.group][zone]++;
      this.oilGroups[entry.group].oils.push(`${oil} (${zone}${troika ? `, топ ${troika}` : ''})`);
    }
    
    // Сбор данных для анализа
    // Передаем в анализ уже развернутые тексты
    const expandedEntry = { ...entry, pe: expandedPe, s: expandedS };
    this.collectAnalysisData(oil, zone, troika, expandedEntry);
  }
  
  /**
   * Собирает данные для анализа
   * @param {string} oil - Название масла
   * @param {string} zone - Зона воздействия  
   * @param {string} troika - Значение тройки
   * @param {Object} entry - Запись из словаря
   */
  collectAnalysisData(oil, zone, troika, entry) {
    const troikaText = troika ? ` (топ ${troika})` : "";
    const oilText = `${oil}${troikaText}`;
    
    switch (zone) {
      case "N":
        this.analysisResults.neutralZoneSize++;
        break;
      case "0":
        this.analysisResults.specialZones.zero.push(oil);
        break;
      case "R":
        this.analysisResults.specialZones.reverse.push(oil);
        break;
      case "+++":
        this.analysisResults.mainTasks.plusPlusPlusPE.push(`${oilText}: ${entry.pe}`);
        this.analysisResults.mainTasks.plusPlusPlusS.push(`${oilText}: ${entry.s}`);
        break;
      case "---":
        this.analysisResults.mainTasks.minusMinusMinusPE.push(`${oil}: ${entry.pe}`);
        this.analysisResults.mainTasks.minusMinusMinusS.push(`${oil}: ${entry.s}`);
        break;
      case "+":
        this.analysisResults.additionalTasks.plusPE.push(`${oil}: ${entry.pe}`);
        this.analysisResults.additionalTasks.plusS.push(`${oil}: ${entry.s}`);
        break;
      case "-":
        this.analysisResults.additionalTasks.minusPE.push(`${oil}: ${entry.pe}`);
        this.analysisResults.additionalTasks.minusS.push(`${oil}: ${entry.s}`);
        break;
    }
    
    // Сбор паттернов
    this.findPatterns(oil, zone, entry);
  }
  
  /**
   * Находит паттерны в данных
   * @param {string} oil - Название масла
   * @param {string} zone - Зона воздействия
   * @param {Object} entry - Запись из словаря
   */
  findPatterns(oil, zone, entry) {
    // Базовые паттерны из словаря
    if (entry.pe.includes("*ЕД") || entry.s.includes("*ЕД")) {
      this.analysisResults.patterns.push(`${oil} (${zone}): ${entry.pe} / ${entry.s}`);
    }
    
    // Специальные паттерны для зоны "---"
    if (zone === "---") {
      const specialPatterns = {
        "Апельсин": "Запрет на радость, глубокая депрессия.",
        "Бергамот": "Глубокая депрессия, застревание в подростковом периоде.",
        "Лимон": "Высокая раздражительность, агрессия."
      };
      
      if (specialPatterns[oil]) {
        this.analysisResults.patterns.push(specialPatterns[oil]);
      }
    }
  }
  
  /**
   * Находит единичные масла в группах
   */
  findSingleOils() {
    Object.entries(this.oilGroups).forEach(([groupName, groupData]) => {
      CONFIG.ZONES.forEach(zone => {
        if (zone !== "oils" && groupData[zone] === 1) {
          const oilInfo = groupData.oils.find(oil => oil.includes(`(${zone})`));
          if (oilInfo) {
            const [oilName] = oilInfo.split(' (');
            const key = `${oilName}|${zone}`;
            const entry = this.dictionary.get(key);
            const singleText = entry && entry.single ? entry.single : '';
            const message = singleText ? 
              `${oilInfo}: ${singleText}` : 
              `${oilInfo} единственное в группе ${groupName}.`;
            this.analysisResults.singleOils.push(message);
          }
        }
      });
    });
  }
  
  /**
   * Анализирует сочетания масел
   */
  analyzeCombinations() {
    const combinationChecker = new CombinationChecker(this.sheets.input, this.dictionary);
    const foundCombinations = combinationChecker.checkAllCombinations();
    this.analysisResults.combinations = combinationChecker.getStructuredCombinations();
    
    // Добавляем информацию о сочетаниях в паттерны
    if (foundCombinations.length > 0) {
      const comboPatterns = combinationChecker.getCombinationsByGroups();
      Object.entries(comboPatterns).forEach(([groupName, combos]) => {
        if (combos.length > 0) {
          this.analysisResults.patterns.push(
            `${groupName} группа: выявлено ${combos.length} сочетаний масел`
          );
        }
      });
    }
  }
  
  /**
   * Генерирует все отчеты
   */
  generateReports() {
    const skewsReporter = new ImprovedSkewsReporter(this.sheets.skews, this.oilGroups);
    const outputReporter = new ImprovedOutputReporter(this.sheets.output, this.analysisResults, this.oilGroups, this.dictionary);
    
    skewsReporter.generate();
    outputReporter.generate();
  }
}
