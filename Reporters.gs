// ==================== УЛУЧШЕННЫЕ ГЕНЕРАТОРЫ ОТЧЕТОВ ====================

/**
 * Улучшенный генератор отчета по перекосам
 * Создает структурированную таблицу с анализом групп масел
 */
class ImprovedSkewsReporter {
  constructor(sheet, groups) {
    this.sheet = sheet;
    this.groups = groups;
  }
  
  /**
   * Генерирует отчет по перекосам
   */
  generate() {
    this.sheet.clear();
    this.createHeader();
    this.fillGroupData();
    this.addSummarySection();
    this.formatReport();
  }
  
  /**
   * Создает заголовок таблицы
   */
  createHeader() {
    const headers = [
      "🌿 Группа масел", "🔥 +++", "⚡ +", "⚪ N", "⚠️ -", "💀 ---", "🚫 0", "🔄 R", 
      "📊 Всего", "🛢️ Масла в группе", "🧠 ПЭ Анализ", "💊 Соматический анализ"
    ];
    
    this.sheet.getRange("A1:L1").setValues([headers])
      .setFontWeight("bold")
      .setFontSize(11)
      .setBackground("#1f4e79")
      .setFontColor("white")
      .setHorizontalAlignment("center");
  }
  
  /**
   * Заполняет данные по группам
   */
  fillGroupData() {
    let rowIndex = 2;
    
    Object.entries(this.groups).forEach(([groupName, groupData]) => {
      const [peAnalysis, sAnalysis] = this.analyzeGroup(groupName, groupData);
      const totalOils = CONFIG.ZONES.reduce((sum, zone) => sum + (groupData[zone] || 0), 0);
      
      const rowData = [
        groupName,
        groupData["+++"] || 0,
        groupData["+"] || 0,
        groupData["N"] || 0,
        groupData["-"] || 0,
        groupData["---"] || 0,
        groupData["0"] || 0,
        groupData["R"] || 0,
        totalOils,
        groupData.oils.join(", ") || "Нет масел",
        peAnalysis,
        sAnalysis
      ];
      
      this.sheet.getRange(rowIndex, 1, 1, 12).setValues([rowData]);
      
      // Цветовое кодирование строк
      if (totalOils > 5) {
        this.sheet.getRange(rowIndex, 1, 1, 12).setBackground("#fff2cc");
      } else if (totalOils === 0) {
        this.sheet.getRange(rowIndex, 1, 1, 12).setBackground("#f4cccc");
      } else if (totalOils <= 2) {
        this.sheet.getRange(rowIndex, 1, 1, 12).setBackground("#d5e8d4");
      }
      
      rowIndex++;
    });
  }
  
  /**
   * Анализирует группу масел
   * @param {string} groupName - Название группы
   * @param {Object} groupData - Данные группы
   * @returns {Array} Массив с ПЭ и соматическим анализом
   */
  analyzeGroup(groupName, groupData) {
    let peAnalysis = "";
    let sAnalysis = "";
    
    const positiveCount = (groupData["+++"] || 0) + (groupData["+"] || 0);
    const negativeCount = (groupData["---"] || 0) + (groupData["-"] || 0);
    const totalCount = CONFIG.ZONES.reduce((sum, zone) => sum + (groupData[zone] || 0), 0);
    
    switch (groupName) {
      case CONFIG.OIL_GROUPS.CITRUS:
        if (positiveCount >= 5) {
          peAnalysis = "🚨 ПЕРЕКОС: Зависимость от чужого мнения";
          sAnalysis = "⚠️ РИСК: Окислительный стресс";
        }
        if (negativeCount >= 5) {
          peAnalysis += (peAnalysis ? " | " : "") + "🚨 ПЕРЕКОС: Игнорирование мнения окружения";
          sAnalysis += (sAnalysis ? " | " : "") + "⚠️ РИСК: Хронические застойные процессы";
        }
        break;
        
      case CONFIG.OIL_GROUPS.CONIFEROUS:
        if (positiveCount >= 4) {
          peAnalysis = "🚨 ПЕРЕКОС: Острые воспалительные процессы";
          sAnalysis = "⚠️ РИСК: Нарушение работы ЦНС";
        }
        if (negativeCount >= 4) {
          peAnalysis += (peAnalysis ? " | " : "") + "🚨 ПЕРЕКОС: Хронические воспаления";
          sAnalysis += (sAnalysis ? " | " : "") + "⚠️ РИСК: Застойные процессы";
        }
        break;
        
      case CONFIG.OIL_GROUPS.SPICE:
        if (positiveCount >= 5) {
          peAnalysis = "🚨 ПЕРЕКОС: Работа с потребностью в признании";
          sAnalysis = "⚠️ РИСК: Нарушения ЖКТ и эндокринной системы";
        }
        if (negativeCount >= 5) {
          peAnalysis += (peAnalysis ? " | " : "") + "🚨 ПЕРЕКОС: Подавление эмоций";
          sAnalysis += (sAnalysis ? " | " : "") + "⚠️ РИСК: Хронические нарушения";
        }
        break;
        
      default:
        if (totalCount > 0) {
          peAnalysis = "📊 Анализ требует дополнительного изучения";
          sAnalysis = "📊 Анализ требует дополнительного изучения";
        }
    }
    
    if (!peAnalysis) {
      peAnalysis = "✅ Перекосов не выявлено";
      sAnalysis = "✅ Перекосов не выявлено";
    }
    
    return [peAnalysis, sAnalysis];
  }
  
  /**
   * Добавляет сводную секцию
   */
  addSummarySection() {
    const lastRow = this.sheet.getLastRow();
    const summaryRow = lastRow + 2;
    
    // Заголовок сводки
    this.sheet.getRange(summaryRow, 1).setValue("📋 СВОДКА ПО ГРУППАМ")
      .setFontSize(14).setFontWeight("bold")
      .setBackground("#e8f4f8");
    
    this.sheet.getRange(summaryRow, 1, 1, 12).merge();
    
    // Данные сводки
    const summaryData = this.generateSummaryData();
    const summaryHeaders = ["Параметр", "Значение", "Статус"];
    
    this.addDataTable(summaryData, summaryHeaders, summaryRow + 1);
  }
  
  /**
   * Генерирует данные для сводки
   * @returns {Array} Массив данных сводки
   */
  generateSummaryData() {
    const summary = [];
    
    Object.entries(this.groups).forEach(([groupName, groupData]) => {
      const totalOils = CONFIG.ZONES.reduce((sum, zone) => sum + (groupData[zone] || 0), 0);
      const positiveCount = (groupData["+++"] || 0) + (groupData["+"] || 0);
      const negativeCount = (groupData["---"] || 0) + (groupData["-"] || 0);
      
      let status = "✅ Норма";
      if (totalOils === 0) status = "⚠️ Нет масел";
      else if (positiveCount >= 5 || negativeCount >= 5) status = "🚨 Перекос";
      else if (totalOils <= 2) status = "💡 Мало масел";
      
      summary.push([groupName, `${totalOils} масел`, status]);
    });
    
    return summary;
  }
  
  /**
   * Добавляет таблицу данных
   * @param {Array} data - Данные таблицы
   * @param {Array} headers - Заголовки таблицы
   * @param {number} startRow - Начальная строка
   */
  addDataTable(data, headers, startRow = 1) {
    // Заголовки
    this.sheet.getRange(startRow, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight("bold")
      .setBackground("#e8f4f8")
      .setHorizontalAlignment("center");
    
    // Данные
    if (data.length > 0) {
      this.sheet.getRange(startRow + 1, 1, data.length, headers.length)
        .setValues(data)
        .setWrap(true)
        .setVerticalAlignment("top");
    }
    
    // Границы таблицы
    const tableRange = this.sheet.getRange(startRow, 1, data.length + 1, headers.length);
    tableRange.setBorder(true, true, true, true, true, true);
  }
  
  /**
   * Применяет форматирование к отчету
   */
  formatReport() {
    const lastRow = this.sheet.getLastRow();
    const lastColumn = 12;
    
    // Основная таблица
    this.sheet.getRange(1, 1, lastRow, lastColumn)
      .setBorder(true, true, true, true, true, true)
      .setWrap(true)
      .setVerticalAlignment("top");
    
    // Автоширина колонок
    for (let col = 1; col <= lastColumn; col++) {
      this.sheet.autoResizeColumn(col);
    }
    
    // Ограничение ширины текстовых колонок
    this.sheet.setColumnWidth(10, 300); // Масла в группе
    this.sheet.setColumnWidth(11, 250); // ПЭ Анализ  
    this.sheet.setColumnWidth(12, 250); // Соматический анализ
  }
}

/**
 * Улучшенный генератор итогового отчета
 * Создает структурированный отчет с таблицами и четкими разделами
 */
class ImprovedOutputReporter {
  constructor(sheet, analysisResults, groups, dictionary) {
    this.sheet = sheet;
    this.data = analysisResults;
    this.groups = groups;
    this.dictionary = dictionary;
    this.currentRow = 1;
  }
  
  /**
   * Генерирует полный отчет
   */
  generate() {
    this.sheet.clear();
    this.createMainHeader();
    this.createExecutiveSummary();
    this.createZoneAnalysisTable();
    this.createCombinationsTable();
    this.createSingleOilsTable();
    this.createPatternsSection();
    this.createRecommendationsSection();
    this.formatReport();
  }
  
  /**
   * Создает главный заголовок
   */
  createMainHeader() {
    const header = "🌿 ПОЛНЫЙ АНАЛИЗ АРОМАТЕРАПИИ ПО АЛГОРИТМУ 3.1 🌿";
    
    this.sheet.getRange(this.currentRow, 1).setValue(header)
      .setFontSize(18).setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#1f4e79")
      .setFontColor("white");
    
    this.sheet.getRange(this.currentRow, 1, 1, 8).merge();
    this.currentRow += 2;
  }
  
  /**
   * Создает краткое резюме
   */
  createExecutiveSummary() {
    this.addSectionHeader("📊 КРАТКОЕ РЕЗЮМЕ", "executive");
    
    const totalOils = Object.values(this.groups)
      .reduce((sum, group) => sum + CONFIG.ZONES.reduce((gSum, zone) => gSum + (group[zone] || 0), 0), 0);
    
    const summaryData = [
      ["🔢 Общее количество масел:", totalOils, this.getStatusForTotal(totalOils)],
      ["⚪ Размер нейтральной зоны (N):", `${this.data.neutralZoneSize}`, this.getStatusForNeutral(this.data.neutralZoneSize)],
      ["🚨 Основные проблемы (++++):", this.data.mainTasks.plusPlusPlusPE.length, this.getStatusForProblems(this.data.mainTasks.plusPlusPlusPE.length)],
      ["💪 Ресурсные состояния (---):", this.data.mainTasks.minusMinusMinusPE.length, this.getStatusForResources(this.data.mainTasks.minusMinusMinusPE.length)],
      ["🔗 Найдено сочетаний масел:", this.data.combinations.length, this.getStatusForCombinations(this.data.combinations.length)],
      ["🔍 Единичные масла:", this.data.singleOils.length, this.getStatusForSingleOils(this.data.singleOils.length)]
    ];
    
    this.addDataTable(summaryData, ["Параметр", "Значение", "Статус"]);
    this.currentRow += 2;
  }
  
  /**
   * Создает таблицу анализа по зонам
   */
  createZoneAnalysisTable() {
    this.addSectionHeader("🎯 АНАЛИЗ ПО ЗОНАМ ВОЗДЕЙСТВИЯ", "zones");
    
    // Основные проблемы (зона +++)
    if (this.data.mainTasks.plusPlusPlusPE.length > 0) {
      this.addSubsectionHeader("🚨 ОСНОВНЫЕ ПРОБЛЕМЫ (ЗОНА +++)", "problems");
      
      const mainProblemsHeaders = ["🛢️ Масло", "🧠 Психоэмоциональное", "💊 Соматическое", "📊 Приоритет"];
      const mainProblemsData = [];
      
      for (let i = 0; i < Math.max(this.data.mainTasks.plusPlusPlusPE.length, this.data.mainTasks.plusPlusPlusS.length); i++) {
        const peEntry = this.data.mainTasks.plusPlusPlusPE[i] || "";
        const sEntry = this.data.mainTasks.plusPlusPlusS[i] || "";
        
        const peParts = peEntry.split(": ");
        const sParts = sEntry.split(": ");
        
        const oilName = peParts[0] || sParts[0] || "";
        const peDesc = peParts[1] || "";
        const sDesc = sParts[1] || "";
        const priority = this.getPriorityForOil(oilName, "+++");
        
        mainProblemsData.push([oilName, peDesc, sDesc, priority]);
      }
      
      this.addDataTable(mainProblemsData, mainProblemsHeaders);
      this.currentRow += 1;
    }
    
    // Ресурсные состояния (зона ---)
    if (this.data.mainTasks.minusMinusMinusPE.length > 0) {
      this.addSubsectionHeader("💪 РЕСУРСНЫЕ СОСТОЯНИЯ (ЗОНА ---)", "resources");
      
      const resourceHeaders = ["🛢️ Масло", "🧠 Психоэмоциональное", "💊 Соматическое", "📊 Потенциал"];
      const resourceData = [];
      
      for (let i = 0; i < Math.max(this.data.mainTasks.minusMinusMinusPE.length, this.data.mainTasks.minusMinusMinusS.length); i++) {
        const peEntry = this.data.mainTasks.minusMinusMinusPE[i] || "";
        const sEntry = this.data.mainTasks.minusMinusMinusS[i] || "";
        
        const peParts = peEntry.split(": ");
        const sParts = sEntry.split(": ");
        
        const oilName = peParts[0] || sParts[0] || "";
        const peDesc = peParts[1] || "";
        const sDesc = sParts[1] || "";
        const potential = this.getPotentialForOil(oilName, "---");
        
        resourceData.push([oilName, peDesc, sDesc, potential]);
      }
      
      this.addDataTable(resourceData, resourceHeaders);
      this.currentRow += 1;
    }
    
    // Специальные зоны
    if (this.data.specialZones.zero.length > 0 || this.data.specialZones.reverse.length > 0) {
      this.addSubsectionHeader("⚡ СПЕЦИАЛЬНЫЕ ЗОНЫ", "special");
      
      const specialZonesData = [
        ["🚫 0-зона (блокировка):", this.data.specialZones.zero.join(", ") || "Нет", this.getStatusForSpecialZone(this.data.specialZones.zero.length)],
        ["🔄 R-зона (реверс):", this.data.specialZones.reverse.length === 1 ? 
          this.data.specialZones.reverse[0] : this.data.specialZones.reverse.join(", ") || "Нет", this.getStatusForSpecialZone(this.data.specialZones.reverse.length)]
      ];
      
      this.addDataTable(specialZonesData, ["Зона", "Масла", "Статус"]);
      this.currentRow += 1;
    }
  }
  
  /**
   * Создает таблицу сочетаний масел
   */
  createCombinationsTable() {
    if (this.data.combinations.length === 0) {
      this.addSectionHeader("🔗 СОЧЕТАНИЯ МАСЕЛ", "combinations");
      this.sheet.getRange(this.currentRow, 1).setValue("✅ Специальных сочетаний не обнаружено.");
      this.currentRow += 2;
      return;
    }
    
    this.addSectionHeader("🔗 НАЙДЕННЫЕ СОЧЕТАНИЯ МАСЕЛ", "combinations");
    
    const combinationsHeaders = ["🛢️ Основное масло", "🔗 Сочетающиеся масла", "🎯 Зоны", "📝 Интерпретация", "📊 Группа"];
    const combinationsData = [];
    
    this.data.combinations.forEach(combo => {
      const foundOilsText = combo.foundOils.map(oil => oil.displayText).join(", ");
      const zonesText = combo.zones.join(", ");
      const groupName = this.getGroupForOil(combo.mainOil);
      
      combinationsData.push([
        `${combo.mainOil} (${combo.mainOilZone}${combo.mainOilTroika ? `, топ ${combo.mainOilTroika}` : ''})`,
        foundOilsText,
        zonesText,
        combo.interpretation,
        groupName
      ]);
    });
    
    this.addDataTable(combinationsData, combinationsHeaders);
    this.currentRow += 2;
  }
  
  /**
   * Создает таблицу единичных масел
   */
  createSingleOilsTable() {
    if (this.data.singleOils.length === 0) {
      this.addSectionHeader("🔍 ЕДИНИЧНЫЕ МАСЛА В ГРУППАХ", "single");
      this.sheet.getRange(this.currentRow, 1).setValue("✅ Единичных масел не обнаружено.");
      this.currentRow += 2;
      return;
    }
    
    this.addSectionHeader("🔍 ЕДИНИЧНЫЕ МАСЛА В ГРУППАХ", "single");
    
    const singleOilsHeaders = ["🔢 №", "🛢️ Масло и интерпретация", "📊 Группа"];
    const singleOilsData = this.data.singleOils.map((oil, index) => {
      const oilName = oil.split(':')[0];
      const groupName = this.getGroupForOil(oilName);
      return [index + 1, oil, groupName];
    });
    
    this.addDataTable(singleOilsData, singleOilsHeaders);
    this.currentRow += 2;
  }
  
  /**
   * Создает секцию паттернов
   */
  createPatternsSection() {
    this.addSectionHeader("🔄 ВЫЯВЛЕННЫЕ ЗАКОНОМЕРНОСТИ", "patterns");
    
    if (this.data.patterns.length === 0) {
      this.sheet.getRange(this.currentRow, 1).setValue("✅ Особых закономерностей не выявлено.");
      this.currentRow += 2;
      return;
    }
    
    const patternsHeaders = ["🔢 №", "📊 Закономерность", "🎯 Тип"];
    const patternsData = this.data.patterns.map((pattern, index) => {
      const type = this.getPatternType(pattern);
      return [index + 1, pattern, type];
    });
    
    this.addDataTable(patternsData, patternsHeaders);
    this.currentRow += 2;
  }
  
  /**
   * Создает секцию рекомендаций
   */
  createRecommendationsSection() {
    this.addSectionHeader("📋 ИТОГОВЫЕ ВЫВОДЫ И РЕКОМЕНДАЦИИ", "recommendations");
    
    // Психоэмоциональный вывод
    this.addSubsectionHeader("🧠 ПСИХОЭМОЦИОНАЛЬНОЕ СОСТОЯНИЕ", "pe");
    const peConclusion = this.generatePsychoemotionalConclusion();
    this.sheet.getRange(this.currentRow, 1, 1, 8).setValue(peConclusion).setWrap(true);
    this.currentRow += 2;
    
    // Соматический вывод
    this.addSubsectionHeader("💊 СОМАТИЧЕСКОЕ СОСТОЯНИЕ", "somatic");
    const sConclusion = this.generateSomaticConclusion();
    this.sheet.getRange(this.currentRow, 1, 1, 8).setValue(sConclusion).setWrap(true);
    this.currentRow += 2;
    
    // Общие рекомендации
    this.addSubsectionHeader("✅ ОБЩИЕ РЕКОМЕНДАЦИИ", "general");
    const recommendations = this.generateGeneralRecommendations();
    
    const recommendationsHeaders = ["🔢 №", "📝 Рекомендация", "📊 Приоритет", "⏰ Сроки"];
    const recommendationsData = recommendations.map((rec, index) => [
      index + 1, 
      rec.text, 
      rec.priority,
      rec.timeline
    ]);
    
    this.addDataTable(recommendationsData, recommendationsHeaders);
  }
  
  /**
   * Добавляет заголовок секции
   * @param {string} title - Заголовок секции
   * @param {string} type - Тип секции для стилизации
   */
  addSectionHeader(title, type = "default") {
    const colors = {
      "executive": "#1f4e79",
      "zones": "#2e7d32",
      "combinations": "#f57c00",
      "single": "#7b1fa2",
      "patterns": "#1976d2",
      "recommendations": "#d32f2f",
      "default": "#1f4e79"
    };
    
    this.sheet.getRange(this.currentRow, 1).setValue(title)
      .setFontSize(16).setFontWeight("bold")
      .setBackground(colors[type] || colors.default)
      .setFontColor("white");
    
    this.sheet.getRange(this.currentRow, 1, 1, 8).merge();
    this.currentRow += 1;
  }
  
  /**
   * Добавляет заголовок подсекции
   * @param {string} title - Заголовок подсекции
   * @param {string} type - Тип подсекции
   */
  addSubsectionHeader(title, type = "default") {
    const colors = {
      "problems": "#d32f2f",
      "resources": "#388e3c",
      "special": "#f57c00",
      "pe": "#1976d2",
      "somatic": "#7b1fa2",
      "general": "#2e7d32",
      "default": "#f0f8ff"
    };
    
    this.sheet.getRange(this.currentRow, 1).setValue(title)
      .setFontSize(14).setFontWeight("bold")
      .setBackground(colors[type] || colors.default)
      .setFontColor(colors[type] ? "white" : "black");
    
    this.sheet.getRange(this.currentRow, 1, 1, 8).merge();
    this.currentRow += 1;
  }
  
  /**
   * Добавляет таблицу данных
   * @param {Array} data - Данные таблицы
   * @param {Array} headers - Заголовки таблицы
   */
  addDataTable(data, headers) {
    // Заголовки
    this.sheet.getRange(this.currentRow, 1, 1, headers.length)
      .setValues([headers])
      .setFontWeight("bold")
      .setBackground("#e8f4f8")
      .setHorizontalAlignment("center");
    
    this.currentRow++;
    
    // Данные
    if (data.length > 0) {
      this.sheet.getRange(this.currentRow, 1, data.length, headers.length)
        .setValues(data)
        .setWrap(true)
        .setVerticalAlignment("top");
      
      this.currentRow += data.length;
    }
    
    // Границы таблицы
    const tableRange = this.sheet.getRange(this.currentRow - data.length - 1, 1, data.length + 1, headers.length);
    tableRange.setBorder(true, true, true, true, true, true);
  }
  
  // Вспомогательные методы для статусов и приоритетов
  getStatusForTotal(total) {
    if (total === 0) return "⚠️ Нет данных";
    if (total <= 10) return "✅ Оптимально";
    if (total <= 20) return "⚠️ Много масел";
    return "🚨 Очень много масел";
  }
  
  getStatusForNeutral(size) {
    if (size === 0) return "🚨 Нет принятия";
    if (size <= 3) return "⚠️ Мало принятия";
    if (size <= 6) return "✅ Умеренное принятие";
    return "✅ Хорошее принятие";
  }
  
  getStatusForProblems(count) {
    if (count === 0) return "✅ Проблем нет";
    if (count <= 3) return "⚠️ Несколько проблем";
    if (count <= 6) return "🚨 Много проблем";
    return "🚨 Критично много проблем";
  }
  
  getStatusForResources(count) {
    if (count === 0) return "⚠️ Нет ресурсов";
    if (count <= 2) return "⚠️ Мало ресурсов";
    if (count <= 4) return "✅ Умеренные ресурсы";
    return "✅ Много ресурсов";
  }
  
  getStatusForCombinations(count) {
    if (count === 0) return "✅ Простая схема";
    if (count <= 5) return "⚠️ Средняя сложность";
    if (count <= 10) return "🚨 Сложная схема";
    return "🚨 Очень сложная схема";
  }
  
  getStatusForSingleOils(count) {
    if (count === 0) return "✅ Группы сбалансированы";
    if (count <= 3) return "⚠️ Несколько единичных";
    if (count <= 6) return "🚨 Много единичных";
    return "🚨 Критично много единичных";
  }
  
  getPriorityForOil(oilName, zone) {
    if (zone === "+++") return "🔴 Высокий";
    if (zone === "+") return "🟡 Средний";
    return "🟢 Низкий";
  }
  
  getPotentialForOil(oilName, zone) {
    if (zone === "---") return "💪 Высокий";
    if (zone === "-") return "⚠️ Средний";
    return "🟢 Низкий";
  }
  
  getStatusForSpecialZone(count) {
    if (count === 0) return "✅ Нет";
    if (count === 1) return "⚠️ Одно масло";
    return "🚨 Несколько масел";
  }
  
  getGroupForOil(oilName) {
    // Ищем группу масла в словаре
    for (const [key, entry] of this.dictionary.entries()) {
      if (key.startsWith(oilName + "|")) {
        return entry.group || "Неизвестно";
      }
    }
    return "Неизвестно";
  }
  
  getPatternType(pattern) {
    if (pattern.includes("сочетаний")) return "🔗 Сочетания";
    if (pattern.includes("группа")) return "🌿 Группа";
    if (pattern.includes("зона")) return "🎯 Зона";
    return "📊 Общий";
  }
  
  /**
   * Генерирует вывод по психоэмоциональному состоянию
   * @returns {string} Вывод
   */
  generatePsychoemotionalConclusion() {
    let conclusion = "";
    
    if (this.data.mainTasks.plusPlusPlusPE.length > 0) {
      conclusion += `Выявлены ${this.data.mainTasks.plusPlusPlusPE.length} основных психоэмоциональных проблем в зоне +++. `;
      
      if (this.data.patterns.length > 0) {
        conclusion += `Обнаружены ${this.data.patterns.length} специфических закономерностей. `;
      }
      
      if (this.data.mainTasks.minusMinusMinusPE.length > 0) {
        conclusion += `Присутствуют ${this.data.mainTasks.minusMinusMinusPE.length} ресурсных состояний (зона ---), что указывает на потенциал восстановления. `;
      } else {
        conclusion += "Ресурсные состояния ограничены - требуется активация внутренних резервов. ";
      }
    } else {
      conclusion += "Острых психоэмоциональных проблем не выявлено, состояние относительно стабильное. ";
    }
    
    const neutralAnalysis = this.data.neutralZoneSize > 3 ? 
      "Большая нейтральная зона указывает на принятие ситуации." : 
      "Маленькая нейтральная зона указывает на внутреннее напряжение.";
    
    conclusion += neutralAnalysis;
    
    return conclusion;
  }
  
  /**
   * Генерирует вывод по соматическому состоянию
   * @returns {string} Вывод
   */
  generateSomaticConclusion() {
    let conclusion = "";
    
    if (this.data.mainTasks.plusPlusPlusS.length > 0 || this.data.mainTasks.minusMinusMinusS.length > 0) {
      conclusion += `Выявлены активные соматические процессы: ${this.data.mainTasks.plusPlusPlusS.length} основных проблем и ${this.data.mainTasks.minusMinusMinusS.length} ресурсных состояний. `;
      
      // Анализ по группам масел
      const citrusGroup = this.groups[CONFIG.OIL_GROUPS.CITRUS];
      const coniferousGroup = this.groups[CONFIG.OIL_GROUPS.CONIFEROUS];
      const spiceGroup = this.groups[CONFIG.OIL_GROUPS.SPICE];
      
      const citrusIssues = ((citrusGroup["---"] || 0) + (citrusGroup["-"] || 0)) >= 5;
      const coniferousIssues = (coniferousGroup["-"] || 0) > 0 || (coniferousGroup["---"] || 0) > 0;
      const spiceIssues = ((spiceGroup["---"] || 0) + (spiceGroup["-"] || 0)) >= 5;
      
      if (citrusIssues || coniferousIssues || spiceIssues) {
        conclusion += "Выявленные риски: ";
        const risks = [];
        if (citrusIssues) risks.push("проблемы с кислотно-щелочным балансом (цитрусовые)");
        if (coniferousIssues) risks.push("острые воспалительные процессы (хвойные)");
        if (spiceIssues) risks.push("хронические нарушения ЖКТ и эндокринной системы (пряные)");
        conclusion += risks.join(", ") + ". ";
      }
    } else {
      conclusion += "Серьезных соматических нарушений не выявлено. ";
    }
    
    if (this.data.additionalTasks.plusS.length > 0) {
      conclusion += `Дополнительно выявлены ${this.data.additionalTasks.plusS.length} поддерживающих факторов в зоне +.`;
    }
    
    return conclusion;
  }
  
  /**
   * Генерирует общие рекомендации
   * @returns {Array} Массив рекомендаций с приоритетами
   */
  generateGeneralRecommendations() {
    const recommendations = [];
    
    // Высокий приоритет
    if (this.data.mainTasks.plusPlusPlusPE.length > 0) {
      recommendations.push({
        text: "Работа с основными психоэмоциональными проблемами (зона +++)",
        priority: "🔴 Высокий",
        timeline: "1-2 недели"
      });
    }
    
    if (this.data.combinations.length > 0) {
      recommendations.push({
        text: `Учесть ${this.data.combinations.length} выявленных сочетаний масел в терапевтической схеме`,
        priority: "🔴 Высокий",
        timeline: "Немедленно"
      });
    }
    
    // Средний приоритет
    if (this.data.singleOils.length > 0) {
      recommendations.push({
        text: "Особое внимание к единичным маслам в группах - они могут указывать на специфические потребности",
        priority: "🟡 Средний",
        timeline: "2-3 недели"
      });
    }
    
    if (this.data.neutralZoneSize <= 3) {
      recommendations.push({
        text: "Работа по снижению внутреннего напряжения (маленькая нейтральная зона)",
        priority: "🟡 Средний",
        timeline: "3-4 недели"
      });
    }
    
    // Низкий приоритет
    if (this.data.additionalTasks.plusPE.length > 0 || this.data.additionalTasks.minusPE.length > 0) {
      recommendations.push({
        text: "Работа с дополнительными психоэмоциональными задачами (зоны + и -)",
        priority: "🟢 Низкий",
        timeline: "4-6 недель"
      });
    }
    
    if (recommendations.length === 0) {
      recommendations.push({
        text: "Состояние стабильное. Рекомендуется профилактическое наблюдение.",
        priority: "ℹ️ Информация",
        timeline: "По необходимости"
      });
    }
    
    return recommendations;
  }
  
  /**
   * Применяет форматирование ко всему отчету
   */
  formatReport() {
    // Автоширина основных колонок
    for (let col = 1; col <= 8; col++) {
      this.sheet.autoResizeColumn(col);
    }
    
    // Установка максимальной ширины для читаемости
    this.sheet.setColumnWidth(1, Math.min(this.sheet.getColumnWidth(1), 200));
    this.sheet.setColumnWidth(2, Math.min(this.sheet.getColumnWidth(2), 300));
    this.sheet.setColumnWidth(3, Math.min(this.sheet.getColumnWidth(3), 250));
    this.sheet.setColumnWidth(4, Math.min(this.sheet.getColumnWidth(4), 400));
    
    // Общее форматирование
    const fullRange = this.sheet.getRange(1, 1, this.currentRow, 8);
    fullRange.setVerticalAlignment("top");
  }
}
