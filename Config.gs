// ==================== КОНСТАНТЫ И КОНФИГУРАЦИЯ ====================

/**
 * Глобальная конфигурация системы анализа ароматерапии
 * @const {Object} CONFIG - Объект конфигурации с именами листов, индексами столбцов и константами
 */
const CONFIG = {
  // Названия листов в таблице
  SHEETS: {
    INPUT: "Ввод",
    DICTIONARY: "Словарь", 
    SKEWS: "Перекосы",
    OUTPUT: "Вывод"
  },
  
  // Индексы столбцов (0-based для внутренних операций)
  COLUMNS: {
    CLIENT_REQUEST: 0,  // A - Запрос клиента
    OIL: 1,            // B - Название масла
    ZONE: 2,           // C - Зона воздействия
    TROIKA: 3,         // D - Тройка (1, 2, 3)
    PE: 5,             // F - Психоэмоциональное воздействие
    S: 6,              // G - Соматическое воздействие
    COMBINATIONS: 8,    // I - Сочетания масел
    DIAGNOSTICS: 17    // R - Диагностические сообщения
  },
  
  // Допустимые значения для "Тройки"
  VALID_TROIKA_VALUES: ["1", "2", "3"],
  
  // Группы эфирных масел для анализа
  OIL_GROUPS: {
    CITRUS: "Цитрусовая",
    CONIFEROUS: "Хвойная", 
    SPICE: "Пряная",
    FLORAL: "Цветочная",
    WOODY_HERBAL: "Древесно-травяная"
  },
  
  // Возможные зоны воздействия
  ZONES: ["+++", "+", "N", "-", "---", "0", "R"],
  
  // Пользовательские сообщения
  MESSAGES: {
    INVALID_TROIKA: "Можно выбрать только 1, 2 или 3!",
    DUPLICATE_TROIKA: "Значение уже выбрано!",
    NO_DATA: "Нет данных",
    KEY_NOT_FOUND: "Ключ не найден:",
    ANALYSIS_COMPLETE: "Анализ успешно обновлен!",
    ANALYSIS_ERROR: "Ошибка при обновлении анализа:"
  },
  
  // Настройки форматирования отчетов
  REPORT_FORMATTING: {
    HEADER_STYLE: { fontWeight: "bold", fontSize: 18, background: "#1f4e79", fontColor: "white" },
    SUBHEADER_STYLE: { fontWeight: "bold", fontSize: 14, background: "#2e7d32", fontColor: "white" },
    DATA_STYLE: { fontSize: 10, wrap: true, verticalAlignment: "top" },
    TABLE_BORDER: { top: true, bottom: true, left: true, right: true },
    COLORS: {
      EXECUTIVE: "#1f4e79",
      ZONES: "#2e7d32", 
      COMBINATIONS: "#f57c00",
      SINGLE: "#7b1fa2",
      PATTERNS: "#1976d2",
      RECOMMENDATIONS: "#d32f2f",
      PROBLEMS: "#d32f2f",
      RESOURCES: "#388e3c",
      SPECIAL: "#f57c00",
      PE: "#1976d2",
      SOMATIC: "#7b1fa2",
      GENERAL: "#2e7d32"
    }
  }
};
