// ==================== УТИЛИТЫ И ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ====================

/**
 * Утилитарный класс для общих операций
 */
class Utils {
  /**
   * Безопасное получение значения ячейки
   * @param {Range} range - Диапазон ячейки
   * @returns {string} Значение ячейки или пустая строка
   */
  static getCellValue(range) {
    try {
      const value = range.getValue();
      return value === null || value === undefined ? "" : value.toString();
    } catch (error) {
      console.error(`Ошибка получения значения ячейки: ${error.message}`);
      return "";
    }
  }

  /**
   * Безопасная установка значения ячейки
   * @param {Range} range - Диапазон ячейки
   * @param {any} value - Значение для установки
   */
  static setCellValue(range, value) {
    try {
      range.setValue(value || "");
    } catch (error) {
      console.error(`Ошибка установки значения ячейки: ${error.message}`);
    }
  }

  /**
   * Проверка, что строка не пуста
   * @param {string} str - Строка для проверки
   * @returns {boolean} true если строка не пуста
   */
  static isNotEmpty(str) {
    return str && str.toString().trim() !== "";
  }

  /**
   * Парсит поле "Сочетания" вида "[1] ... [2] ..." в словарь номер->описание
   * @param {string} combosText - Исходный текст сочетаний
   * @returns {Object} Объект { index: Map<number,string>, order: number[] }
   */
  static parseCombos(combosText) {
    const result = new Map();
    const order = [];
    
    if (!combosText || typeof combosText !== 'string') {
      return { index: result, order };
    }
    
    // Улучшенный regex для поиска сочетаний
    const regex = /\[(\d+)\]\s*([^\[]+?)(?=\s*\[\d+\]|$)/g;
    let match;
    
    while ((match = regex.exec(combosText)) !== null) {
      const num = parseInt(match[1], 10);
      let desc = match[2].trim();
      
      // Очищаем описание от лишних символов
      desc = desc.replace(/\s+/g, ' ').trim();
      
      if (!Number.isNaN(num) && desc.length > 0) {
        // Убираем лишние пробелы и точки в конце
        desc = desc.replace(/\s*\.\s*$/, '').trim();
        if (!desc.endsWith('.')) {
          desc += '.';
        }
        
        result.set(num, desc);
        order.push(num);
      }
    }
    
    return { index: result, order: order.sort((a, b) => a - b) };
  }

  /**
   * Разворачивает ссылки вида "См. сочетание [1],[2]" в явные тексты из словаря сочетаний
   * @param {string} text - Исходный текст (ПЭ/С)
   * @param {Map<number,string>} combosIndex - Словарь номеров к описаниям
   * @returns {string} Текст без ссылок с подставленными расшифровками
   */
  static expandCombinationReferences(text, combosIndex) {
    if (!text) return '';
    const numbers = new Set();
    const numRegex = /\[(\d+)\]/g;
    let m;
    while ((m = numRegex.exec(text)) !== null) {
      const n = parseInt(m[1], 10);
      if (!Number.isNaN(n)) numbers.add(n);
    }

    // Удаляем фразы "См. сочетание(я) ..."
    let cleaned = text
      .replace(/См\.?\s*сочетани[ея][^\.]*\.?/gi, '')
      .replace(/\s{2,}/g, ' ')
      .trim();

    if (numbers.size === 0 || !combosIndex || combosIndex.size === 0) {
      return cleaned;
    }

    const additions = [];
    numbers.forEach((n) => {
      const desc = combosIndex.get(n);
      if (desc) additions.push(desc);
    });

    if (additions.length > 0) {
      const extra = additions.join(' ');
      cleaned = cleaned ? `${cleaned} ${extra}` : extra;
    }
    return cleaned.trim();
  }
  
  /**
   * Форматирует текст для красивого отображения в отчётах
   * @param {string} text - Исходный текст
   * @param {string} type - Тип форматирования
   * @returns {string} Отформатированный текст
   */
  static formatReportText(text, type = 'data') {
    if (!text) return '';
    
    switch (type) {
      case 'header':
        return `\n${'='.repeat(80)}\n${text.toUpperCase()}\n${'='.repeat(80)}\n`;
      case 'subheader':
        return `\n${'-'.repeat(60)}\n${text}\n${'-'.repeat(60)}\n`;
      case 'emphasis':
        return `**${text}**`;
      case 'warning':
        return `⚠️ ${text}`;
      case 'success':
        return `✅ ${text}`;
      case 'error':
        return `🚨 ${text}`;
      default:
        return text;
    }
  }
  
  /**
   * Создает красивую таблицу для отчётов
   * @param {Array} data - Данные таблицы
   * @param {Array} headers - Заголовки
   * @returns {string} Отформатированная таблица
   */
  static createFormattedTable(data, headers) {
    if (!data || data.length === 0) return "Нет данных для отображения";
    
    let table = "\n";
    
    // Заголовки
    table += "| " + headers.join(" | ") + " |\n";
    table += "| " + headers.map(() => "---").join(" | ") + " |\n";
    
    // Данные
    data.forEach(row => {
      table += "| " + row.join(" | ") + " |\n";
    });
    
    return table;
  }
  
  /**
   * Проверяет корректность данных масла
   * @param {string} oil - Название масла
   * @param {string} zone - Зона
   * @param {string} group - Группа
   * @returns {Object} Результат проверки
   */
  static validateOilData(oil, zone, group) {
    const result = {
      isValid: true,
      errors: [],
      warnings: []
    };
    
    if (!oil || oil.trim() === '') {
      result.isValid = false;
      result.errors.push("Название масла не может быть пустым");
    }
    
    if (!zone || !CONFIG.ZONES.includes(zone)) {
      result.isValid = false;
      result.errors.push(`Некорректная зона: ${zone}`);
    }
    
    if (group && !Object.values(CONFIG.OIL_GROUPS).includes(group)) {
      result.warnings.push(`Неизвестная группа: ${group}`);
    }
    
    return result;
  }
  
  /**
   * Генерирует уникальный ключ для масла
   * @param {string} oil - Название масла
   * @param {string} zone - Зона
   * @returns {string} Уникальный ключ
   */
  static generateOilKey(oil, zone) {
    return `${oil.trim()}|${zone}`;
  }
  
  /**
   * Очищает и нормализует текст
   * @param {string} text - Исходный текст
   * @returns {string} Очищенный текст
   */
  static cleanText(text) {
    if (!text) return '';
    
    return text
      .replace(/\s+/g, ' ')           // Убираем множественные пробелы
      .replace(/\n+/g, ' ')           // Заменяем переносы строк на пробелы
      .replace(/[^\w\s\-\.\,\:\;\(\)\[\]]/g, '') // Убираем специальные символы
      .trim();
  }
}
