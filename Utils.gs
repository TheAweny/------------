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
   * Форматирование текста для отчетов
   * @param {string} text - Исходный текст
   * @param {string} type - Тип форматирования (header, subheader, data)
   * @returns {string} Отформатированный текст
   */
  static formatReportText(text, type = 'data') {
    switch (type) {
      case 'header':
        return `\n${'='.repeat(80)}\n${text.toUpperCase()}\n${'='.repeat(80)}\n`;
      case 'subheader':
        return `\n${'-'.repeat(60)}\n${text}\n${'-'.repeat(60)}\n`;
      default:
        return text;
    }
  }

  /**
   * Парсит поле "Сочетания" вида "[1] ... [2] ..." в словарь номер->описание
   * @param {string} combosText - Исходный текст сочетаний
   * @returns {Object} Объект { index: Map<number,string>, order: number[] }
   */
  static parseCombos(combosText) {
    const result = new Map();
    const order = [];
    if (!combosText) return { index: result, order };
    const regex = /\[(\d+)\]\s*([^\[]+)/g; // [n] followed by description until next [ or end
    let match;
    while ((match = regex.exec(combosText)) !== null) {
      const num = parseInt(match[1], 10);
      const desc = match[2].trim().replace(/\s+/g, ' ');
      if (!Number.isNaN(num)) {
        result.set(num, desc.endsWith('.') ? desc : desc + '.');
        order.push(num);
      }
    }
    return { index: result, order };
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
}
