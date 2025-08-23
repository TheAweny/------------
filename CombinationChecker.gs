// ==================== АНАЛИЗ СОЧЕТАНИЙ МАСЕЛ ====================

/**
 * Класс для автоматической проверки сочетаний эфирных масел
 * Динамически загружает правила сочетаний из словаря и проверяет их
 */
class CombinationChecker {
  constructor(inputSheet, dictionary) {
    this.inputSheet = inputSheet;
    this.dictionary = dictionary;
    this.data = this.loadInputData();
    this.oilZones = this.indexOilsByZones();
    this.foundCombinations = [];
  }
  
  /**
   * Загружает данные с листа ввода
   * @returns {Array} Массив данных
   */
  loadInputData() {
    try {
      return this.inputSheet.getDataRange().getValues();
    } catch (error) {
      console.error(`Ошибка загрузки данных: ${error.message}`);
      return [];
    }
  }
  
  /**
   * Индексирует все масла по зонам для быстрого поиска
   * @returns {Object} Объект с маслами, сгруппированными по зонам
   */
  indexOilsByZones() {
    const oilZones = {};
    
    for (let i = 1; i < this.data.length; i++) {
      const oil = this.data[i][CONFIG.COLUMNS.OIL];
      const zone = this.data[i][CONFIG.COLUMNS.ZONE];
      const troika = this.data[i][CONFIG.COLUMNS.TROIKA];
      
      if (Utils.isNotEmpty(oil) && Utils.isNotEmpty(zone)) {
        if (!oilZones[oil]) {
          oilZones[oil] = [];
        }
        oilZones[oil].push({
          zone: zone.toString(),
          troika: troika ? troika.toString() : "",
          row: i + 1
        });
      }
    }
    
    return oilZones;
  }
  
  /**
   * Проверяет все возможные сочетания масел на основе данных из словаря
   * @returns {Array} Массив найденных сочетаний
   */
  checkAllCombinations() {
    this.foundCombinations = [];
    
    // Проходим по всем маслам в вводе
    Object.keys(this.oilZones).forEach(oil => {
      // Проверяем каждую зону для данного масла
      this.oilZones[oil].forEach(zoneInfo => {
        const key = `${oil}|${zoneInfo.zone}`;
        const dictEntry = this.dictionary.get(key);
        
        if (dictEntry && dictEntry.combos) {
          // Парсим сочетания из словаря
          const parsedCombos = Utils.parseCombos(dictEntry.combos);
          
          // Проверяем каждое сочетание
          parsedCombos.order.forEach(comboIndex => {
            const comboText = parsedCombos.index.get(comboIndex);
            if (comboText) {
              this.checkCombinationFromText(oil, zoneInfo, comboText, comboIndex);
            }
          });
        }
      });
    });
    
    return this.foundCombinations;
  }
  
  /**
   * Проверяет конкретное сочетание на основе текста из словаря
   * @param {string} mainOil - Основное масло
   * @param {Object} zoneInfo - Информация о зоне
   * @param {string} comboText - Текст сочетания
   * @param {number} comboIndex - Индекс сочетания
   */
  checkCombinationFromText(mainOil, zoneInfo, comboText, comboIndex) {
    // Ищем упоминания других масел в тексте сочетания
    const mentionedOils = this.extractMentionedOils(comboText);
    
    if (mentionedOils.length > 0) {
      // Проверяем, есть ли эти масла в вводе в соответствующих зонах
      const foundOils = this.findMentionedOilsInInput(mentionedOils, zoneInfo.zone);
      
      if (foundOils.length > 0) {
        // Создаем запись о найденном сочетании
        this.foundCombinations.push({
          mainOil: mainOil,
          mainOilZone: zoneInfo.zone,
          mainOilTroika: zoneInfo.troika,
          comboIndex: comboIndex,
          comboText: comboText,
          foundOils: foundOils,
          interpretation: this.extractInterpretation(comboText)
        });
      }
    }
  }
  
  /**
   * Извлекает упомянутые масла из текста сочетания
   * @param {string} comboText - Текст сочетания
   * @returns {Array} Массив названий масел
   */
  extractMentionedOils(comboText) {
    const oils = [];
    
    // Ищем упоминания масел в тексте
    Object.keys(this.oilZones).forEach(oilName => {
      if (comboText.includes(oilName)) {
        oils.push(oilName);
      }
    });
    
    return oils;
  }
  
  /**
   * Находит упомянутые масла во вводе в соответствующих зонах
   * @param {Array} mentionedOils - Упомянутые масла
   * @param {string} mainZone - Зона основного масла
   * @returns {Array} Массив найденных масел
   */
  findMentionedOilsInInput(mentionedOils, mainZone) {
    const found = [];
    
    mentionedOils.forEach(oilName => {
      if (this.oilZones[oilName]) {
        this.oilZones[oilName].forEach(zoneInfo => {
          // Проверяем, подходит ли зона для сочетания
          if (this.isZoneCompatibleForCombination(mainZone, zoneInfo.zone)) {
            found.push({
              oil: oilName,
              zone: zoneInfo.zone,
              troika: zoneInfo.troika,
              row: zoneInfo.row
            });
          }
        });
      }
    });
    
    return found;
  }
  
  /**
   * Проверяет совместимость зон для сочетания
   * @param {string} mainZone - Зона основного масла
   * @param {string} otherZone - Зона другого масла
   * @returns {boolean} true если зоны совместимы
   */
  isZoneCompatibleForCombination(mainZone, otherZone) {
    // Базовые правила совместимости зон
    const zoneCompatibility = {
      "+++": ["+++", "+", "---", "-"],
      "+": ["+++", "+", "---", "-"],
      "---": ["+++", "+", "---", "-"],
      "-": ["+++", "+", "---", "-"],
      "0": ["+++", "+", "---", "-"],
      "R": ["+++", "+", "---", "-"],
      "N": ["+++", "+", "---", "-"]
    };
    
    return zoneCompatibility[mainZone] && zoneCompatibility[mainZone].includes(otherZone);
  }
  
  /**
   * Извлекает интерпретацию из текста сочетания
   * @param {string} comboText - Текст сочетания
   * @returns {string} Интерпретация
   */
  extractInterpretation(comboText) {
    // Убираем упоминания масел и оставляем только интерпретацию
    let interpretation = comboText;
    
    // Убираем фразы типа "в сочетании с", "В сочетании с"
    interpretation = interpretation.replace(/[Вв]\s+сочетании\s+с\s+[^\.]+\.?\s*/g, '');
    
    // Убираем лишние пробелы
    interpretation = interpretation.replace(/\s+/g, ' ').trim();
    
    return interpretation;
  }
  
  /**
   * Получает структурированный отчет по сочетаниям для отчета
   * @returns {Array} Массив объектов сочетаний
   */
  getStructuredCombinations() {
    return this.foundCombinations.map(combo => ({
      mainOil: combo.mainOil,
      mainOilZone: combo.mainOilZone,
      mainOilTroika: combo.mainOilTroika,
      comboIndex: combo.comboIndex,
      comboText: combo.comboText,
      foundOils: combo.foundOils.map(oil => ({
        name: oil.oil,
        zone: oil.zone,
        troika: oil.troika,
        displayText: `${oil.oil} (${oil.zone}${oil.troika ? `, топ ${oil.troika}` : ''})`
      })),
      interpretation: combo.interpretation,
      zones: [combo.mainOilZone, ...combo.foundOils.map(oil => oil.zone)]
    }));
  }
  
  /**
   * Получает количество найденных сочетаний
   * @returns {number} Количество сочетаний
   */
  getCombinationsCount() {
    return this.foundCombinations.length;
  }
  
  /**
   * Получает сочетания по группам масел
   * @returns {Object} Объект с сочетаниями по группам
   */
  getCombinationsByGroups() {
    const groups = {};
    
    this.foundCombinations.forEach(combo => {
      const mainOilEntry = this.dictionary.get(`${combo.mainOil}|${combo.mainOilZone}`);
      if (mainOilEntry && mainOilEntry.group) {
        const groupName = mainOilEntry.group;
        if (!groups[groupName]) {
          groups[groupName] = [];
        }
        groups[groupName].push(combo);
      }
    });
    
    return groups;
  }
}
