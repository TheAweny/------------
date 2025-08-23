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
      "Группа масел", "+++", "+", "N", "-", "---", "0", "R", 
      "Всего масел", "Масла в группе", "ПЭ Анализ", "Соматический анализ"
    ];
    
    this.sheet.getRange("A1:L1").setValues([headers])
      .setFontWeight("bold")
      .setBackground("#e8f4f8")
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
          peAnalysis = "ПЕРЕКОС: Зависимость от чужого мнения";
          sAnalysis = "РИСК: Окислительный стресс";
        }
        if (negativeCount >= 5) {
          peAnalysis += (peAnalysis ? " | " : "") + "ПЕРЕКОС: Игнорирование мнения окружения";
          sAnalysis += (sAnalysis ? " | " : "") + "РИСК: Хронические застойные процессы";
        }
        break;
        
      case CONFIG.OIL_GROUPS.CONIFEROUS:
        if (positiveCount + (groupData["0"] || 0) === 5) {
          peAnalysis = "КРИТИЧНО: Состояние паники, гиперстресс";
        }
        if (negativeCount >= 5) {
          peAnalysis = "РИСК: Пофигизм, не чувствует опасности";
        }
        if ((groupData["-"] || 0) > 0) {
          sAnalysis = "ВНИМАНИЕ: Острый воспалительный процесс";
        }
        break;
        
      case CONFIG.OIL_GROUPS.SPICE:
        if (positiveCount === 5) {
          peAnalysis = "ПОТРЕБНОСТЬ: Признание, тепло и забота";
        }
        if (negativeCount >= 4) {
          sAnalysis = "ХРОНИЧНО: Нарушения ЖКТ и эндокринной системы";
        }
        break;
        
      case CONFIG.OIL_GROUPS.FLORAL:
        if ((groupData["N"] || 0) > 3) {
          peAnalysis = "НОРМА: Принятие женственности без напряжения";
        }
        break;
    }
    
    // Добавляем общий анализ если специфического нет
    if (!peAnalysis && totalCount > 0) {
      if (positiveCount > negativeCount) {
        peAnalysis = "Преобладание активирующих масел";
      } else if (negativeCount > positiveCount) {
        peAnalysis = "Преобладание ресурсных масел";
      } else {
        peAnalysis = "Сбалансированное распределение";
      }
    }
    
    if (!sAnalysis && totalCount > 0) {
      sAnalysis = totalCount > 3 ? "Активная группа" : "Умеренная активность";
    }
    
    return [peAnalysis || "Нет данных", sAnalysis || "Нет данных"];
  }
  
  /**
   * Добавляет секцию общего анализа
   */
  addSummarySection() {
    const startRow = Object.keys(this.groups).length + 3;
    
    this.sheet.getRange(startRow, 1).setValue("ОБЩИЙ АНАЛИЗ ПЕРЕКОСОВ:")
      .setFontWeight("bold").setFontSize(12);
    
    let summaryRow = startRow + 1;
    const totalOilsCount = Object.values(this.groups)
      .reduce((sum, group) => sum + CONFIG.ZONES.reduce((gSum, zone) => gSum + (group[zone] || 0), 0), 0);
    
    this.sheet.getRange(summaryRow++, 1, 1, 2)
      .setValues([["Общее количество масел:", totalOilsCount]]);
    
    // Анализ доминирующих групп
    const dominantGroups = Object.entries(this.groups)
      .filter(([name, data]) => CONFIG.ZONES.reduce((sum, zone) => sum + (data[zone] || 0), 0) >= 3)
      .map(([name]) => name);
    
    this.sheet.getRange(summaryRow++, 1, 1, 2)
      .setValues([["Доминирующие группы:", dominantGroups.join(", ") || "Нет доминирующих групп"]]);
    
    // Рекомендации
    const recommendations = this.generateRecommendations();
    this.sheet.getRange(summaryRow++, 1).setValue("РЕКОМЕНДАЦИИ:")
      .setFontWeight("bold");
    
    recommendations.forEach((rec, index) => {
      this.sheet.getRange(summaryRow + index, 1, 1, 3)
        .setValues([[`${index + 1}. ${rec}`, "", ""]]);
    });
  }
  
  /**
   * Генерирует рекомендации на основе анализа
   * @returns {Array} Массив рекомендаций
   */
  generateRecommendations() {
    const recommendations = [];
    
    Object.entries(this.groups).forEach(([groupName, groupData]) => {
      const positiveCount = (groupData["+++"] || 0) + (groupData["+"] || 0);
      const negativeCount = (groupData["---"] || 0) + (groupData["-"] || 0);
      
      if (positiveCount >= 5 && groupName === CONFIG.OIL_GROUPS.CITRUS) {
        recommendations.push("Работа с самооценкой и независимостью мнения (цитрусовый перекос)");
      }
      
      if (negativeCount >= 5 && groupName === CONFIG.OIL_GROUPS.CONIFEROUS) {
        recommendations.push("Обратить внимание на воспалительные процессы (хвойная группа)");
      }
      
      if (positiveCount === 5 && groupName === CONFIG.OIL_GROUPS.SPICE) {
        recommendations.push("Работа с потребностью в признании и заботе (пряная группа)");
      }
    });
    
    if (recommendations.length === 0) {
      recommendations.push("Перекосов не выявлено. Продолжить наблюдение.");
    }
    
    return recommendations;
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
  constructor(sheet, analysisResults, groups) {
    this.sheet = sheet;
    this.data = analysisResults;
    this.groups = groups;
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
      .setFontSize(16).setFontWeight("bold")
      .setHorizontalAlignment("center")
      .setBackground("#e8f4f8");
    
    this.sheet.getRange(this.currentRow, 1, 1, 6).merge();
    this.currentRow += 2;
  }
  
  /**
   * Создает краткое резюме
   */
  createExecutiveSummary() {
    this.addSectionHeader("📊 КРАТКОЕ РЕЗЮМЕ");
    
    const totalOils = Object.values(this.groups)
      .reduce((sum, group) => sum + CONFIG.ZONES.reduce((gSum, zone) => gSum + (group[zone] || 0), 0), 0);
    
    const summaryData = [
      ["Общее количество масел:", totalOils],
      ["Размер нейтральной зоны (N):", `${this.data.neutralZoneSize} ${this.data.neutralZoneSize > 3 ? '(большая - принятие)' : '(маленькая - напряжение)'}`],
      ["Основные проблемы (++++):", this.data.mainTasks.plusPlusPlusPE.length],
      ["Ресурсные состояния (---):", this.data.mainTasks.minusMinusMinusPE.length],
      ["Найдено сочетаний масел:", this.data.combinations.length],
      ["Единичные масла:", this.data.singleOils.length]
    ];
    
    this.addDataTable(summaryData, ["Параметр", "Значение"]);
    this.currentRow += 2;
  }
  
  /**
   * Создает таблицу анализа по зонам
   */
  createZoneAnalysisTable() {
    this.addSectionHeader("🎯 АНАЛИЗ ПО ЗОНАМ ВОЗДЕЙСТВИЯ");
    
    // Основные проблемы (зона +++)
    if (this.data.mainTasks.plusPlusPlusPE.length > 0) {
      this.addSubsectionHeader("🚨 ОСНОВНЫЕ ПРОБЛЕМЫ (ЗОНА +++)");
      
      const mainProblemsHeaders = ["Масло", "Психоэмоциональное", "Соматическое"];
      const mainProblemsData = [];
      
      for (let i = 0; i < Math.max(this.data.mainTasks.plusPlusPlusPE.length, this.data.mainTasks.plusPlusPlusS.length); i++) {
        const peEntry = this.data.mainTasks.plusPlusPlusPE[i] || "";
        const sEntry = this.data.mainTasks.plusPlusPlusS[i] || "";
        
        const peParts = peEntry.split(": ");
        const sParts = sEntry.split(": ");
        
        const oilName = peParts[0] || sParts[0] || "";
        const peDesc = peParts[1] || "";
        const sDesc = sParts[1] || "";
        
        mainProblemsData.push([oilName, peDesc, sDesc]);
      }
      
      this.addDataTable(mainProblemsData, mainProblemsHeaders);
      this.currentRow += 1;
    }
    
    // Ресурсные состояния (зона ---)
    if (this.data.mainTasks.minusMinusMinusPE.length > 0) {
      this.addSubsectionHeader("💪 РЕСУРСНЫЕ СОСТОЯНИЯ (ЗОНА ---)");
      
      const resourceHeaders = ["Масло", "Психоэмоциональное", "Соматическое"];
      const resourceData = [];
      
      for (let i = 0; i < Math.max(this.data.mainTasks.minusMinusMinusPE.length, this.data.mainTasks.minusMinusMinusS.length); i++) {
        const peEntry = this.data.mainTasks.minusMinusMinusPE[i] || "";
        const sEntry = this.data.mainTasks.minusMinusMinusS[i] || "";
        
        const peParts = peEntry.split(": ");
        const sParts = sEntry.split(": ");
        
        const oilName = peParts[0] || sParts[0] || "";
        const peDesc = peParts[1] || "";
        const sDesc = sParts[1] || "";
        
        resourceData.push([oilName, peDesc, sDesc]);
      }
      
      this.addDataTable(resourceData, resourceHeaders);
      this.currentRow += 1;
    }
    
    // Специальные зоны
    if (this.data.specialZones.zero.length > 0 || this.data.specialZones.reverse.length > 0) {
      this.addSubsectionHeader("⚡ СПЕЦИАЛЬНЫЕ ЗОНЫ");
      
      const specialZonesData = [
        ["0-зона (блокировка):", this.data.specialZones.zero.join(", ") || "Нет"],
        ["R-зона (реверс):", this.data.specialZones.reverse.length === 1 ? 
          this.data.specialZones.reverse[0] : this.data.specialZones.reverse.join(", ") || "Нет"]
      ];
      
      this.addDataTable(specialZonesData, ["Зона", "Масла"]);
      this.currentRow += 1;
    }
  }
  
  /**
   * Создает таблицу сочетаний масел
   */
  createCombinationsTable() {
    if (this.data.combinations.length === 0) {
      this.addSectionHeader("🔗 СОЧЕТАНИЯ МАСЕЛ");
      this.sheet.getRange(this.currentRow, 1).setValue("Специальных сочетаний не обнаружено.");
      this.currentRow += 2;
      return;
    }
    
    this.addSectionHeader("🔗 НАЙДЕННЫЕ СОЧЕТАНИЯ МАСЕЛ");
    
    const combinationsHeaders = ["Основное масло", "Сочетающиеся масла", "Зоны", "Интерпретация"];
    const combinationsData = [];
    
    this.data.combinations.forEach(combo => {
      const foundOilsText = combo.foundOils.map(oil => oil.displayText).join(", ");
      const zonesText = combo.zones.join(", ");
      const resultsText = combo.results.join(" | ");
      
      combinationsData.push([
        combo.mainOil,
        foundOilsText,
        zonesText,
        resultsText
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
      this.addSectionHeader("🔍 ЕДИНИЧНЫЕ МАСЛА В ГРУППАХ");
      this.sheet.getRange(this.currentRow, 1).setValue("Единичных масел не обнаружено.");
      this.currentRow += 2;
      return;
    }
    
    this.addSectionHeader("🔍 ЕДИНИЧНЫЕ МАСЛА В ГРУППАХ");
    
    const singleOilsHeaders = ["№", "Масло и интерпретация"];
    const singleOilsData = this.data.singleOils.map((oil, index) => [index + 1, oil]);
    
    this.addDataTable(singleOilsData, singleOilsHeaders);
    this.currentRow += 2;
  }
  
  /**
   * Создает секцию паттернов
   */
  createPatternsSection() {
    this.addSectionHeader("🔄 ВЫЯВЛЕННЫЕ ЗАКОНОМЕРНОСТИ");
    
    if (this.data.patterns.length === 0) {
      this.sheet.getRange(this.currentRow, 1).setValue("Особых закономерностей не выявлено.");
      this.currentRow += 2;
      return;
    }
    
    const patternsHeaders = ["№", "Закономерность"];
    const patternsData = this.data.patterns.map((pattern, index) => [index + 1, pattern]);
    
    this.addDataTable(patternsData, patternsHeaders);
    this.currentRow += 2;
  }
  
  /**
   * Создает секцию рекомендаций
   */
  createRecommendationsSection() {
    this.addSectionHeader("📋 ИТОГОВЫЕ ВЫВОДЫ И РЕКОМЕНДАЦИИ");
    
    // Психоэмоциональный вывод
    this.addSubsectionHeader("🧠 ПСИХОЭМОЦИОНАЛЬНОЕ СОСТОЯНИЕ");
    const peConclusion = this.generatePsychoemotionalConclusion();
    this.sheet.getRange(this.currentRow, 1, 1, 6).setValue(peConclusion).setWrap(true);
    this.currentRow += 2;
    
    // Соматический вывод
    this.addSubsectionHeader("💊 СОМАТИЧЕСКОЕ СОСТОЯНИЕ");
    const sConclusion = this.generateSomaticConclusion();
    this.sheet.getRange(this.currentRow, 1, 1, 6).setValue(sConclusion).setWrap(true);
    this.currentRow += 2;
    
    // Общие рекомендации
    this.addSubsectionHeader("✅ ОБЩИЕ РЕКОМЕНДАЦИИ");
    const recommendations = this.generateGeneralRecommendations();
    
    const recommendationsHeaders = ["№", "Рекомендация", "Приоритет"];
    const recommendationsData = recommendations.map((rec, index) => [
      index + 1, 
      rec.text, 
      rec.priority
    ]);
    
    this.addDataTable(recommendationsData, recommendationsHeaders);
  }
  
  /**
   * Добавляет заголовок секции
   * @param {string} title - Заголовок секции
   */
  addSectionHeader(title) {
    this.sheet.getRange(this.currentRow, 1).setValue(title)
      .setFontSize(14).setFontWeight("bold")
      .setBackground("#f0f8ff");
    
    this.sheet.getRange(this.currentRow, 1, 1, 6).merge();
    this.currentRow += 1;
  }
  
  /**
   * Добавляет заголовок подсекции
   * @param {string} title - Заголовок подсекции
   */
  addSubsectionHeader(title) {
    this.sheet.getRange(this.currentRow, 1).setValue(title)
      .setFontSize(12).setFontWeight("bold")
      .setBackground("#f8f8ff");
    
    this.sheet.getRange(this.currentRow, 1, 1, 6).merge();
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
        priority: "Высокий"
      });
    }
    
    if (this.data.combinations.length > 0) {
      recommendations.push({
        text: `Учесть ${this.data.combinations.length} выявленных сочетаний масел в терапевтической схеме`,
        priority: "Высокий"
      });
    }
    
    // Средний приоритет
    if (this.data.singleOils.length > 0) {
      recommendations.push({
        text: "Особое внимание к единичным маслам в группах - они могут указывать на специфические потребности",
        priority: "Средний"
      });
    }
    
    if (this.data.neutralZoneSize <= 3) {
      recommendations.push({
        text: "Работа по снижению внутреннего напряжения (маленькая нейтральная зона)",
        priority: "Средний"
      });
    }
    
    // Низкий приоритет
    if (this.data.additionalTasks.plusPE.length > 0 || this.data.additionalTasks.minusPE.length > 0) {
      recommendations.push({
        text: "Работа с дополнительными психоэмоциональными задачами (зоны + и -)",
        priority: "Низкий"
      });
    }
    
    if (recommendations.length === 0) {
      recommendations.push({
        text: "Состояние стабильное. Рекомендуется профилактическое наблюдение.",
        priority: "Информация"
      });
    }
    
    return recommendations;
  }
  
  /**
   * Применяет форматирование ко всему отчету
   */
  formatReport() {
    // Автоширина основных колонок
    for (let col = 1; col <= 6; col++) {
      this.sheet.autoResizeColumn(col);
    }
    
    // Установка максимальной ширины для читаемости
    this.sheet.setColumnWidth(1, Math.min(this.sheet.getColumnWidth(1), 200));
    this.sheet.setColumnWidth(2, Math.min(this.sheet.getColumnWidth(2), 300));
    this.sheet.setColumnWidth(3, Math.min(this.sheet.getColumnWidth(3), 250));
    this.sheet.setColumnWidth(4, Math.min(this.sheet.getColumnWidth(4), 400));
    
    // Общее форматирование
    const fullRange = this.sheet.getRange(1, 1, this.currentRow, 6);
    fullRange.setVerticalAlignment("top");
  }
}
