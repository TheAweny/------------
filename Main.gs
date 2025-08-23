// ==================== ОБРАБОТЧИКИ СОБЫТИЙ ====================

/**
 * Главный обработчик редактирования ячеек
 * @param {Event} e - Событие редактирования
 */
function onEdit(e) {
  try {
    const validator = new CellValidator(e);
    if (!validator.shouldProcess()) return;
    
    validator.validateAndProcess();
    
    // Запускаем анализ с небольшой задержкой для производительности
    Utilities.sleep(100);
    updateAnalysis();
    
  } catch (error) {
    console.error(`Ошибка в onEdit: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Произошла ошибка: ${error.message}`);
  }
}

/**
 * Принудительное обновление анализа
 */
function forceUpdateAnalysis() {
  updateAnalysis();
}

/**
 * Основная функция обновления анализа
 */
function updateAnalysis() {
  try {
    const analyzer = new AromatherapyAnalyzer();
    analyzer.performAnalysis();
  } catch (error) {
    console.error(`Ошибка в updateAnalysis: ${error.message}`);
    SpreadsheetApp.getUi().alert(`${CONFIG.MESSAGES.ANALYSIS_ERROR} ${error.message}`);
  }
}

/**
 * Создание меню при открытии файла
 */
function onOpen() {
  try {
    const ui = SpreadsheetApp.getUi();
    ui.createMenu('🌿 Ароматерапия')
      .addItem('🔄 Обновить анализ', 'forceUpdateAnalysis')
      .addItem('🧹 Очистить форматирование', 'clearAllFormatting')
      .addItem('🧪 Тестировать систему', 'runSystemTest')
      .addItem('ℹ️ Справка', 'showHelp')
      .addToUi();
  } catch (error) {
    console.error(`Ошибка создания меню: ${error.message}`);
  }
}

/**
 * Очистка форматирования (для отладки)
 */
function clearAllFormatting() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    if (sheet.getName() === CONFIG.SHEETS.INPUT) {
      const range = sheet.getRange(2, CONFIG.COLUMNS.TROIKA + 1, sheet.getLastRow() - 1, 1);
      range.clearDataValidations();
      range.setBackground("white");
      SpreadsheetApp.getActiveSpreadsheet().toast("Форматирование очищено", "Готово", 2);
    } else {
      SpreadsheetApp.getUi().alert("Функция доступна только на листе 'Ввод'");
    }
  } catch (error) {
    console.error(`Ошибка очистки форматирования: ${error.message}`);
  }
}

/**
 * Тестирование системы
 */
function runSystemTest() {
  try {
    console.log('🧪 Запуск системного теста...');
    
    // Тест инициализации
    const analyzer = new AromatherapyAnalyzer();
    console.log('✅ Инициализация анализатора');
    
    // Тест загрузки словаря
    analyzer.loadDictionary();
    console.log(`✅ Словарь загружен: ${analyzer.dictionary.size} записей`);
    
    // Тест проверки сочетаний
    const combinationChecker = new CombinationChecker(analyzer.sheets.input);
    const combinations = combinationChecker.checkAllCombinations();
    console.log(`✅ Проверка сочетаний: найдено ${combinations.length} сочетаний`);
    
    SpreadsheetApp.getActiveSpreadsheet().toast("Системный тест пройден успешно!", "Тест", 3);
    
  } catch (error) {
    console.error(`❌ Ошибка теста: ${error.message}`);
    SpreadsheetApp.getUi().alert(`Ошибка теста: ${error.message}`);
  }
}

/**
 * Показ справки пользователю
 */
function showHelp() {
  const ui = SpreadsheetApp.getUi();
  const helpText = `
🌿 СИСТЕМА АНАЛИЗА АРОМАТЕРАПИИ (Версия 2.0)

📋 ОСНОВНЫЕ ФУНКЦИИ:
• Автоматическая валидация введенных данных
• Проверка уникальности значений "Тройка" в зоне +++
• Анализ сочетаний эфирных масел по 830+ правилам
• Структурированные отчеты с таблицами и выводами
• Выявление единичных масел и закономерностей

📊 СТРУКТУРА ЛИСТОВ:
• "Ввод" — ввод данных для анализа
• "Словарь" — база масел и их свойств
• "Перекосы" — анализ дисбалансов по группам с рекомендациями
• "Вывод" — итоговый структурированный отчет

🚀 КАК ПОЛЬЗОВАТЬСЯ:
1. Заполните лист "Ввод": масло, зона, тройка
2. Система автоматически проверит данные и обновит анализ
3. Просмотрите отчеты на листах "Перекосы" и "Вывод"
4. При необходимости используйте "Обновить анализ" в меню

✨ НОВЫЕ ВОЗМОЖНОСТИ:
• Структурированные таблицы вместо текстовых блоков
• Явный вывод всех сочетаний без ссылок
• Цветовое кодирование важных данных  
• Приоритизация рекомендаций
• Улучшенная читаемость отчетов

⚙️ ТЕХНИЧЕСКАЯ ИНФОРМАЦИЯ:
• Оптимизированная архитектура кода
• Улучшенная обработка ошибок
• Автоматическое форматирование отчетов
• Системные тесты для проверки работоспособности
  `;
  
  ui.alert('Справка по системе', helpText, ui.ButtonSet.OK);
}

// ==================== ЭКСПОРТ КОНФИГУРАЦИИ ====================

/**
 * Функция для экспорта текущей конфигурации (для разработчиков)
 */
function exportConfig() {
  console.log('📤 Экспорт конфигурации системы:');
  console.log(JSON.stringify(CONFIG, null, 2));
}

// END OF SCRIPT
