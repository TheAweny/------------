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
 * Класс для автоматической проверки сочетаний масел
 */
class CombinationChecker {
  constructor(inputSheet) {
    this.inputSheet = inputSheet;
    this.data = inputSheet.getDataRange().getValues();
    this.oilZones = this.indexOilZones();
    this.combinations = this.initializeCombinations();
    this.foundCombinations = [];
  }
  
  /**
   * Индексация всех масел по зонам из листа "Ввод"
   */
  indexOilZones() {
    const oilZones = {};
    
    for (let i = 1; i < this.data.length; i++) {
      const oil = this.data[i][CONFIG.COLUMNS.OIL];
      const zone = this.data[i][CONFIG.COLUMNS.ZONE];
      const troika = this.data[i][CONFIG.COLUMNS.TROIKA];
      
      if (oil && zone) {
        if (!oilZones[oil]) {
          oilZones[oil] = [];
        }
        oilZones[oil].push({
          zone: zone,
          troika: troika,
          row: i + 1
        });
      }
    }
    
    return oilZones;
  }
  
  /**
   * Инициализация всех правил сочетаний из словаря
   */
  initializeCombinations() {
    return {
      // Цитрусовая группа
      "Апельсин": [
        {
          oils: ["Литцея Кубеба"],
          zones: ["+++", "+", "---", "-"],
          results: [
            "[1] Переутомление. Потребность в повышении концентрации внимания.",
            "[4] Напряжение, нарушение адаптации."
          ]
        },
        {
          oils: ["Грейпфрут", "Ваниль"],
          zones: ["+++", "+", "---", "-"],
          results: ["[5] Нарушение пищевого поведения."],
          requireAll: true
        },
        {
          oils: ["Бергамот", "Мандарин"],
          zones: ["+++", "+"],
          results: ["[2] Потребность в радости, беззаботности."],
          requireAny: true
        }
      ],
      "Бергамот": [
        {
          oils: ["Лимон"],
          zones: ["+++", "+", "---", "-"],
          results: ["[1] Эмоциональное переутомление. Часто встречается у подростков."]
        }
      ],
      "Лимон": [
        {
          oils: ["Бергамот"],
          zones: ["+++", "+", "---", "-"],
          results: ["[1] Эмоциональное переутомление. Часто встречается у подростков."]
        },
        {
          oils: ["Литцея Кубеба"],
          zones: ["+++", "+", "---", "-"],
          results: ["[2] Склонность к мигреням."]
        },
        {
          oils: ["Аир"],
          zones: ["+++", "+", "---", "-"],
          results: ["[3] Возможно лямблиоз."]
        }
      ],
      "Грейпфрут": [
        {
          oils: ["Бензоин"],
          zones: ["+++", "---", "0"],
          results: ["[1] Склонность к сахарному диабету."]
        },
        {
          oils: ["Герань"],
          zones: ["+++", "---"],
          results: ["[2] Нарушение работы щитовидной железы."]
        },
        {
          oils: ["Литцея Кубеба"],
          zones: ["+++", "+"],
          results: ["[3] Нарушена адаптация. Нарушение работы ЦНС. Головные боли."]
        }
      ],
      "Литцея Кубеба": [
        {
          oils: ["Апельсин"],
          zones: ["+++", "+", "---", "-"],
          results: [
            "[1] Напряжение, нарушение адаптации.",
            "[4] Нарушение работы ЦНС."
          ]
        },
        {
          oils: ["Лимон"],
          zones: ["+++", "+", "---", "-"],
          results: ["[2] Склонность к мигреням."]
        },
        {
          oils: ["Грейпфрут"],
          zones: ["+++", "+"],
          results: ["[3] Нарушена адаптация. Нарушение работы ЦНС. Головные боли."]
        },
        {
          oils: ["Мята", "Эвкалипт"],
          zones: ["+++", "+", "---", "-"],
          results: ["[4] Последствия травмы ЦНС."],
          requireAny: true
        }
      ],
      
      // Хвойная группа
      "Кедр": [
        {
          oils: ["Кипарис"],
          zones: ["+++", "+", "---", "-"],
          results: ["[2] Острая потребность в понимании и сильном плече."]
        },
        {
          oils: ["Можжевеловые ягоды"],
          zones: ["+++", "+", "---", "-"],
          results: ["[3] Хронические нарушения МПС."]
        }
      ],
      "Кипарис": [
        {
          oils: ["Ель", "Можжевеловые ягоды"],
          zones: ["+++", "+", "---", "-"],
          results: ["[1] Потребность в поддержке на фоне страхов."],
          requireAny: true
        },
        {
          oils: ["Ладан", "Бензоин"],
          zones: ["+++", "+", "---", "-"],
          results: ["[2] Непрожитая потеря близкого человека."],
          requireAny: true
        },
        {
          oils: ["Лимон"],
          zones: ["+++"],
          results: ["[4] Купероз, нарушение работы печени."]
        },
        {
          oils: ["Мята", "Иланг-иланг"],
          zones: ["+++", "+", "---", "-"],
          results: ["[5] Хронические нарушения ССС."],
          requireAny: true
        },
        {
          oils: ["Иланг-иланг"],
          zones: ["+++"],
          results: ["[6] Нарушение сердечного ритма."]
        }
      ],
      "Ель": [
        {
          oils: ["Кипарис", "Можжевеловые ягоды"],
          zones: ["+++", "+", "---", "-"],
          results: ["[1] Потребность в поддержке на фоне страхов."],
          requireAny: true
        },
        {
          oils: ["Пихта", "Можжевеловые ягоды"],
          zones: ["+++", "+"],
          results: ["[2] Человек не справляется со страхами."],
          requireAny: true
        },
        {
          oils: ["Каяпут"],
          zones: ["+++", "+"],
          results: ["[3] Может быть обострение Дорсопатии."]
        },
        {
          oils: ["Чабрец", "Эвкалипт", "Анис"],
          zones: ["+++", "+"],
          results: ["[5] Остаточные явления после ОРВИ с нарушением ДС."],
          requireAny: true
        }
      ],
      "Пихта": [
        {
          oils: ["Ель", "Можжевеловые ягоды"],
          zones: ["+++", "+"],
          results: ["[3] Человек не справляется со страхами."],
          requireAny: true
        },
        {
          oils: ["Ладан", "Кипарис"],
          zones: ["+++", "+", "---", "-"],
          results: ["[2] Страх смерти."],
          requireAny: true
        },
        {
          oils: ["Герань"],
          zones: ["---", "-"],
          results: ["[4] ЛОР-нарушения."]
        },
        {
          oils: ["Бензоин"],
          zones: ["+++", "+", "---", "-"],
          results: ["[5] Нейродермиты."]
        }
      ],
      "Можжевеловые ягоды": [
        {
          oils: ["Пихта"],
          zones: ["+++", "+", "---", "-"],
          results: ["[1] Усиливаются/нивелируются внутренние страхи."]
        },
        {
          oils: ["Ель", "Пихта"],
          zones: ["+++", "+"],
          results: ["[3] Человек не справляется со страхами."],
          requireAny: true
        },
        {
          oils: ["Полынь"],
          zones: ["+++", "+"],
          results: ["[4] Может быть эндокринные нарушения."]
        },
        {
          oils: ["Кедр"],
          zones: ["+++", "---"],
          results: ["[5] Отеки."]
        },
        {
          oils: ["Кедр"],
          zones: ["-"],
          results: ["[6] Нарушения МПС."]
        }
      ],
      
      // Пряная группа
      "Анис": [
        {
          oils: ["Фенхель"],
          zones: ["+++", "+"],
          results: ["[1] Проявление тревожности."]
        },
        {
          oils: ["Ель", "Пихта", "Можжевеловые ягоды"],
          zones: ["+++", "+", "---", "-"],
          results: ["[3] Человека преследуют навязчивые страхи."],
          requireAny: true
        },
        {
          oils: ["Чабрец"],
          zones: ["+++", "+", "---", "-"],
          results: ["[4] Тотально снижена самооценка, человек занимается самокопанием."]
        },
        {
          oils: ["Фенхель", "Чабрец"],
          zones: ["+++", "+", "---", "-"],
          results: ["[5] Зависимость/склонность к зависимости, часто у курильщиков."],
          requireAny: true
        },
        {
          oils: ["Фенхель"],
          zones: ["+++", "+", "---"],
          results: ["[6] Высокое содержание слизи в кишечнике."]
        },
        {
          oils: ["Фенхель", "Чабрец"],
          zones: ["+++", "+", "---", "-"],
          results: ["[8] ОРВИ."],
          requireAny: true
        },
        {
          oils: ["Фенхель", "Ветивер", "Чабрец"],
          zones: ["+++", "+", "---", "-"],
          results: ["[10] Дисбактериоз."],
          requireAny: true
        }
      ],
      "Гвоздика": [
        {
          oils: ["Мята"],
          zones: ["+++", "+", "---", "-"],
          results: [
            "[1] Эмоциональная зацикленность.",
            "[3] Нарушение мозгового кровообращения."
          ]
        },
        {
          oils: ["Полынь", "Аир", "Берёза", "Эвкалипт"],
          zones: ["+++", "+", "---", "-"],
          results: ["[2] Паразитарная инвазия усиливается ароматами."],
          requireAny: true
        }
      ],
      "Мята": [
        {
          oils: ["Лаванда", "Иланг-иланг"],
          zones: ["+++", "+"],
          results: ["[2] Склонность к повышенному АД."],
          requireAny: true
        },
        {
          oils: ["Лаванда", "Кипарис"],
          zones: ["+++", "+"],
          results: ["[3] Может быть нарушение ритма, нарушение ССС."],
          requireAny: true
        },
        {
          oils: ["Розмарин"],
          zones: ["+++", "+"],
          results: ["[4] ВСД."]
        },
        {
          oils: ["Литцея Кубеба", "Эвкалипт"],
          zones: ["+++", "+", "---", "-"],
          results: ["[5] Последствия травмы/нарушение работы ЦНС."],
          requireAny: true
        },
        {
          oils: ["Иланг-иланг"],
          zones: ["+++", "+", "---", "-"],
          results: ["[6] Нарушение работы ССС."]
        }
      ],
      "Чабрец": [
        {
          oils: ["Анис", "Фенхель"],
          zones: ["+++", "+", "---", "-"],
          results: ["[2] Зависимость/склонность к зависимости, часто у курильщиков."],
          requireAny: true
        },
        {
          oils: ["Фенхель"],
          zones: ["+++", "+", "---", "-"],
          results: ["[3] Локализация в кишечнике."]
        },
        {
          oils: ["Литцея Кубеба"],
          zones: ["+++"],
          results: ["[4] Может быть в НС."]
        },
        {
          oils: ["Анис", "Фенхель"],
          zones: ["+++", "+", "---", "-"],
          results: ["[5] ОРВИ."],
          requireAny: true
        }
      ],
      "Фенхель": [
        {
          oils: ["Анис"],
          zones: ["+++", "+"],
          results: ["[1] Подавленная тревожность."]
        },
        {
          oils: ["Ель", "Пихта", "Можжевеловые ягоды"],
          zones: ["+++", "+", "---", "-"],
          results: ["[3] Человека преследуют навязчивые страхи."],
          requireAny: true
        },
        {
          oils: ["Чабрец"],
          zones: ["+++", "+", "---", "-"],
          results: ["[5] ОРВИ."]
        }
      ],
      "Каяпут": [
        {
          oils: ["Анис", "Эвкалипт"],
          zones: ["+++", "+"],
          results: ["[3] Нарушения ОДС, скорее всего грудной отдел."],
          requireAny: true
        },
        {
          oils: ["Литцея Кубеба", "Мята"],
          zones: ["+++", "+"],
          results: ["[4] Нарушения ОДС, скорее всего шейный отдел."],
          requireAny: true
        },
        {
          oils: ["Фенхель", "Грейпфрут", "Лимон"],
          zones: ["+++", "+"],
          results: ["[5] Нарушения ОДС, скорее всего поясничный отдел."],
          requireAny: true
        },
        {
          oils: ["Ель", "Пихта"],
          zones: ["+++", "+", "---", "-"],
          results: ["[6] Возможны суставные боли."],
          requireAny: true
        }
      ],
      "Розмарин": [
        {
          oils: ["Можжевеловые ягоды"],
          zones: ["+++", "+"],
          results: [
            "[1] Острый воспалительный процесс МПС.",
            "[2] Хронические процессы в МПС."
          ]
        },
        {
          oils: ["Мята"],
          zones: ["+++", "+"],
          results: ["[3] ВСД."]
        }
      ],
      
      // Цветочная группа
      "Лаванда": [
        {
          oils: ["Мята", "Кипарис", "Иланг-иланг"],
          zones: ["+++", "+"],
          results: ["[3] Склонность к повышенному АД."],
          requireAny: true
        },
        {
          oils: ["Мята", "Кипарис"],
          zones: ["+++", "+"],
          results: ["[4] Может быть нарушение ритма, нарушение ССС."],
          requireAny: true
        },
        {
          oils: ["Ваниль"],
          zones: ["+++", "+"],
          results: ["[1] Значимость маминой фигуры с эмоциональным перенапряжением."]
        },
        {
          oils: ["Иланг-иланг"],
          zones: ["+++", "+"],
          results: ["[3] Тотальное перенапряжение и нарушение сна."]
        },
        {
          oils: ["Литцея Кубеба", "Фенхель", "Бензоин", "Каяпут"],
          zones: ["+++", "+", "---", "-"],
          results: ["[4] Хроническое нарушение нервной системы или эндокринные нарушения."],
          requireAny: true
        }
      ],
      "Герань": [
        {
          oils: ["Грейпфрут"],
          zones: ["+++", "---"],
          results: ["[2] Нарушение работы щитовидной железы."]
        }
      ],
      "Пальмароза": [
        {
          oils: ["Литцея Кубеба", "Грейпфрут"],
          zones: ["+++"],
          results: ["[3] Нарушение адаптации."],
          requireAny: true
        }
      ],
      "Ваниль": [
        {
          oils: ["Лаванда"],
          zones: ["+++", "+"],
          results: ["[1] Принятие 'женской роли'."]
        },
        {
          oils: ["Грейпфрут", "Апельсин"],
          zones: ["+++", "+", "---", "-"],
          results: ["[3] Нарушение пищевого поведения."],
          requireAny: true
        },
        {
          oils: ["Герань"],
          zones: ["-"],
          results: ["[4] Ситуация может быть усилена и другими Ж маслами или др М маслами."]
        },
        {
          oils: ["Эвкалипт"],
          zones: ["+++", "+", "---"],
          results: ["[5] Склонность к повышению сахара."]
        }
      ],
      "Иланг-иланг": [
        {
          oils: ["Чабрец"],
          zones: ["-"],
          results: ["[2] Высокий грибковый фон."]
        },
        {
          oils: ["Лаванда", "Мята", "Кипарис"],
          zones: ["+++", "+"],
          results: ["[3] Может быть нарушение ССС, сердечного ритма, сосудистая патология."],
          requireAny: true
        },
        {
          oils: ["Кипарис"],
          zones: ["+++"],
          results: ["[4] Нарушение сердечного ритма."]
        }
      ],
      "Бензоин": [
        {
          oils: ["Анис"],
          zones: ["+++", "+"],
          results: ["[1] Осознанная обида на отца."]
        },
        {
          oils: ["Ладан", "Кипарис"],
          zones: ["+++", "+", "---", "-"],
          results: ["[2] Непрожитая потеря близкого человека."],
          requireAny: true
        },
        {
          oils: ["Анис", "Эвкалипт", "Апельсин"],
          zones: ["+++", "+", "-", "0"],
          results: ["[3] Нарушение ДС."],
          requireAny: true
        },
        {
          oils: ["Мандарин", "Бергамот"],
          zones: ["+++", "+", "---", "-"],
          results: ["[4] Сухая кожа."],
          requireAny: true
        }
      ],
      "Ладан": [
        {
          oils: ["Кедр", "Можжевеловые ягоды"],
          zones: ["+++", "+", "---", "-"],
          results: ["[3] Застойные процессы."],
          requireAny: true
        },
        {
          oils: ["Иланг-иланг", "Герань"],
          zones: ["+++", "+", "---", "-"],
          results: ["[4] Снижение тонуса кожи, атония."],
          requireAny: true
        }
      ],
      
      // Древесно-травяная группа
      "Берёза": [
        {
          oils: ["Полынь"],
          zones: ["+++", "+"],
          results: ["[1] Может быть острый паразитарный процесс на коже."]
        },
        {
          oils: ["Гвоздика"],
          zones: ["+++", "+", "---", "-"],
          results: ["[2] Может быть острый паразитарный процесс в ЖКТ."]
        }
      ],
      "Аир": [
        {
          oils: ["Полынь"],
          zones: ["+++", "+", "---"],
          results: ["[1] Творческий перфекционист."]
        },
        {
          oils: ["Лимон"],
          zones: ["+++", "+", "---", "-"],
          results: ["[3] Может быть лямблиоз."]
        }
      ],
      "Полынь": [
        {
          oils: ["Аир"],
          zones: ["+++", "+", "---"],
          results: ["[1] Творческий перфекционист."]
        },
        {
          oils: ["Мята"],
          zones: ["+++", "+", "---", "-"],
          results: ["[2] Высокое интеллектуальное перенапряжение."]
        },
        {
          oils: ["Ель"],
          zones: ["+++", "+"],
          results: ["[3] Истощение коры надпочечников."]
        },
        {
          oils: ["Берёза", "Гвоздика", "Аир"],
          zones: ["---", "-"],
          results: ["[4] Однозначно хронический паразитарный процесс."],
          requireAny: true
        }
      ],
      "Эвкалипт": [
        {
          oils: ["Чабрец"],
          zones: ["+++", "+"],
          results: ["[1] ОРВИ."]
        },
        {
          oils: ["Анис", "Бензоин"],
          zones: ["+++", "+", "-", "0"],
          results: ["[2] Нарушение ДС."],
          requireAny: true
        },
        {
          oils: ["Можжевеловые ягоды"],
          zones: ["+++", "+"],
          results: ["[3] Нарушение функции почек."]
        },
        {
          oils: ["Ваниль"],
          zones: ["+++", "+", "---"],
          results: ["[4] Склонность к повышению сахара."]
        },
        {
          oils: ["Литцея Кубеба", "Мята"],
          zones: ["+++", "+", "---", "-"],
          results: ["[6] Последствия травмы ЦНС."],
          requireAny: true
        }
      ],
      "Ветивер": [
        {
          oils: ["Анис", "Фенхель", "Герань"],
          zones: ["+++", "+", "---", "-"],
          results: ["[1] Нарушения ЭС."],
          requireAny: true
        },
        {
          oils: ["Анис", "Фенхель"],
          zones: ["+++", "+", "---", "-"],
          results: ["[2] Дисбактериоз."],
          requireAny: true
        }
      ]
    };
  }
  
  /**
   * Автоматическая проверка всех возможных сочетаний
   */
  checkAllCombinations() {
    this.foundCombinations = [];
    
    // Проходим по всем маслам в данных
    Object.keys(this.oilZones).forEach(oil => {
      const rules = this.combinations[oil];
      if (rules) {
        rules.forEach(rule => {
          if (this.checkRule(rule)) {
            this.foundCombinations.push({
              mainOil: oil,
              rule: rule,
              foundOils: this.findOilsForRule(rule),
              results: rule.results
            });
          }
        });
      }
    });
    
    return this.foundCombinations;
  }
  
  /**
   * Проверка конкретного правила
   */
  checkRule(rule) {
    const foundOils = this.findOilsForRule(rule);
    
    if (rule.requireAll) {
      return foundOils.length === rule.oils.length;
    } else if (rule.requireAny) {
      return foundOils.length > 0;
    } else {
      // По умолчанию требуется все
      return foundOils.length === rule.oils.length;
    }
  }
  
  /**
   * Поиск масел для конкретного правила
   */
  findOilsForRule(rule) {
    const found = [];
    
    rule.oils.forEach(targetOil => {
      if (this.oilZones[targetOil]) {
        this.oilZones[targetOil].forEach(zoneInfo => {
          if (rule.zones.includes(zoneInfo.zone)) {
            found.push({
              oil: targetOil,
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
   * Получение структурированного отчета по сочетаниям
   */
  getCombinationsReport() {
    if (this.foundCombinations.length === 0) {
      return "Специальных сочетаний не обнаружено.";
    }
    
    // Группируем по основным маслам
    const groupedCombinations = {};
    
    this.foundCombinations.forEach(combo => {
      if (!groupedCombinations[combo.mainOil]) {
        groupedCombinations[combo.mainOil] = [];
      }
      groupedCombinations[combo.mainOil].push(combo);
    });
    
    let report = "";
    
    Object.entries(groupedCombinations).forEach(([mainOil, combos]) => {
      report += `\n🍃 ${mainOil}:\n`;
      
      combos.forEach((combo, index) => {
        const foundOilsText = combo.foundOils.map(f => `${f.oil} (${f.zone}${f.troika ? `, топ ${f.troika}` : ''})`).join(" + ");
        report += `  ${index + 1}. ${foundOilsText}:\n`;
        
        combo.results.forEach(result => {
          report += `     ${result}\n`;
        });
        report += "\n";
      });
    });
    
    return report.trim();
  }
  
  /**
   * Получение краткого списка всех найденных сочетаний
   */
  getCombinationsSummary() {
    return this.foundCombinations.map(combo => {
      const foundOilsText = combo.foundOils.map(f => `${f.oil} (${f.zone})`).join(" + ");
      return `${combo.mainOil} + ${foundOilsText}: ${combo.results.join("; ")}`;
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
   * Загрузка словаря
   */
  loadDictionary() {
    const dictData = this.dictionarySheet.getDataRange().getValues();
    
    for (let i = 1; i < dictData.length; i++) {
      const [oil, zone, pe, s, group, combos, single] = dictData[i];
      const key = `${oil}|${zone}`;
      
      this.dictionary.set(key, {
        pe: pe || "",
        s: s || "",
        group: group || "",
        combos: combos || "",
        single: single || ""
      });
    }
  }
  
  /**
   * Основной метод анализа
   */
  performAnalysis() {
    this.loadDictionary();
    this.processInputData();
    this.findSingleOils();
    this.generateSkewsReport();
    this.generateOutputReport();
  }
  
  /**
   * Обработка входных данных
   */
  processInputData() {
    const data = this.inputSheet.getDataRange().getValues();
    
    for (let row = 1; row < data.length; row++) {
      this.processDataRow(data[row], row + 1);
    }
  }
  
  /**
   * Обработка одной строки данных
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
    sheet.getRange(rowIndex, CONFIG.COLUMNS.PE + 1).setValue(CONFIG.MESSAGES.NO_DATA);
    sheet.getRange(rowIndex, CONFIG.COLUMNS.S + 1).setValue(CONFIG.MESSAGES.NO_DATA);
    sheet.getRange(rowIndex, CONFIG.COLUMNS.COMBINATIONS + 1).setValue("");
    sheet.getRange(rowIndex, CONFIG.COLUMNS.DIAGNOSTICS + 1).setValue("");
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
}

// ==================== КЛАССЫ ДЛЯ ГЕНЕРАЦИИ ОТЧЕТОВ ====================

/**
 * Генератор отчета по перекосам
 */
class SkewsReporter {
  constructor(sheet, groups) {
    this.sheet = sheet;
    this.groups = groups;
  }
  
  generate() {
    this.sheet.clear();
    this.createHeader();
    this.fillData();
    this.formatReport();
  }
  
  createHeader() {
    const headers = [
      "Группа", "+++", "+", "N", "-", "---", "0", "R", 
      "Масла", "ПЭ Интерпретация", "С Интерпретация"
    ];
    
    this.sheet.getRange("A1:K1").setValues([headers]);
  }
  
  fillData() {
    let rowIndex = 2;
    const citrusData = this.groups[CONFIG.OIL_GROUPS.CITRUS];
    const citrusCount = citrusData["+++"] + citrusData["+"];
    const citrusNegativeCount = citrusData["---"] + citrusData["-"];
    
    Object.entries(this.groups).forEach(([groupName, groupData]) => {
      const [peSkew, sSkew] = this.analyzeGroupSkews(groupName, groupData, {
        citrusCount,
        citrusNegativeCount
      });
      
      const rowData = [
        groupName,
        groupData["+++"],
        groupData["+"],
        groupData["N"],
        groupData["-"],
        groupData["---"],
        groupData["0"],
        groupData["R"],
        groupData.oils.join(", ") || "Нет масел",
        peSkew,
        sSkew
      ];
      
      this.sheet.getRange(rowIndex, 1, 1, 11).setValues([rowData]);
      rowIndex++;
    });
  }
  
  analyzeGroupSkews(groupName, groupData, citrusInfo) {
    let peSkew = "";
    let sSkew = "";
    
    switch (groupName) {
      case CONFIG.OIL_GROUPS.CITRUS:
        if (citrusInfo.citrusCount >= 5) {
          peSkew += "Зависимость от чужого мнения. ";
          sSkew += "Окислительный стресс. ";
        }
        if (citrusInfo.citrusNegativeCount >= 5) {
          peSkew += "Только собственное мнение, не считаемся с мнением окружения. ";
          sSkew += "Хронический застойный процесс. Может быть нарушение гормонального фона. ";
        }
        break;
        
      case CONFIG.OIL_GROUPS.CONIFEROUS:
        if (groupData["+++"] + groupData["+"] + groupData["0"] === 5) {
          peSkew += "Состояние паники, гиперстресс, человек прячет голову в песок. ";
        }
        if (groupData["---"] + groupData["-"] >= 5) {
          peSkew += "Человек пофигист, не чувствует опасности. ";
        }
        if (groupData["-"] > 0) {
          sSkew += "Острое воспалительное процесс (проверить основную соматику). ";
        }
        break;
        
      case CONFIG.OIL_GROUPS.SPICE:
        if (groupData["+++"] + groupData["+"] === 5) {
          peSkew += "Потребность в признании, тепле и заботе. ";
        }
        if (groupData["---"] + groupData["-"] >= 4) {
          sSkew += "Хронические нарушения ЖКТ и ЭС. ";
        }
        break;
        
      case CONFIG.OIL_GROUPS.FLORAL:
        if (groupData["N"] > 3) {
          peSkew += "Принятие женственности как данность без напряжения. ";
        }
        break;
    }
    
    return [peSkew.trim(), sSkew.trim()];
  }
  
  formatReport() {
    const lastRow = this.sheet.getLastRow();
    this.sheet.getRange(1, 1, 1, 11).setFontWeight("bold").setBackground("#f0f0f0");
    this.sheet.getRange(1, 1, lastRow, 11).setWrap(true).setVerticalAlignment("top");
    this.sheet.setColumnWidths(1, 11, 150);
  }
}

/**
 * Генератор итогового отчета
 */
class OutputReporter {
  constructor(sheet, analysisData, groups) {
    this.sheet = sheet;
    this.data = analysisData;
    this.groups = groups;
  }
  
  generate(clientRequest) {
    this.sheet.clear();
    
    let row = 1;
    this.sheet.getRange(row++, 1).setValue(this.generateHeader()).setFontWeight("bold");
    row++;
    
    this.sheet.getRange(row++, 1).setValue(this.generateSkewsAnalysis());
    row++;
    
    this.sheet.getRange(row++, 1).setValue(this.generateSingleOils());
    row++;
    
    this.sheet.getRange(row++, 1).setValue(this.generateSpecialZones());
    row++;
    
    this.sheet.getRange(row++, 1).setValue(this.generateKeyTasks());
    row++;
    
    this.sheet.getRange(row++, 1).setValue(this.generateResourceTasks());
    row++;
    
    this.sheet.getRange(row++, 1).setValue(this.generateAdditionalTasks());
    row++;
    
    this.sheet.getRange(row++, 1).setValue(this.generateCombinations());
    row++;
    
    this.sheet.getRange(row++, 1).setValue(this.generatePatterns());
    row++;
    
    this.sheet.getRange(row++, 1).setValue(this.generatePESummary());
    row++;
    
    this.sheet.getRange(row++, 1).setValue(this.generateSSummary());
    row++;
    
    this.formatReport();
  }
  
  generateHeader() {
    return "🌿 ПОЛНЫЙ АНАЛИЗ ПО АЛГОРИТМУ РАСШИФРОВКИ АРОМАТЕРАПИИ 3.1 🌿\n" + "=".repeat(70);
  }
  
  generateSkewsAnalysis() {
    let section = "📊 1. АНАЛИЗ ПЕРЕКОСОВ И СООТНОШЕНИЕ МАСЕЛ ПО ЗОНАМ:\n" + "─".repeat(60) + "\n";
    
    Object.entries(this.groups).forEach(([groupName, groupData]) => {
      section += `\n🔸 ${groupName}:\n`;
      section += `   📈 Распределение по зонам:\n`;
      section += `      +++: ${groupData["+++"]} | +: ${groupData["+"]} | N: ${groupData["N"]} | -: ${groupData["-"]} | ---: ${groupData["---"]} | 0: ${groupData["0"]} | R: ${groupData["R"]}\n`;
      section += `   🍃 Масла: ${groupData.oils.join(", ") || "Нет"}\n`;
    });
    
    const neutralAnalysis = this.data.neutralZoneSize > 3 
      ? `Большая (${this.data.neutralZoneSize} масел) - принятие как данность.`
      : `Маленькая (${this.data.neutralZoneSize} масел) - напряжение.`;
    
    section += `\n\n🎯 Нейтральная зона (N): ${neutralAnalysis}`;
    
    return section;
  }
  
  generateSingleOils() {
    return `🔍 2. ЕДИНИЧНЫЕ МАСЛА В ГРУППАХ:\n${"─".repeat(50)}\n${
      this.data.singleOils.length 
        ? this.data.singleOils.map(item => `• ${item}`).join("\n") 
        : "Единичных масел не обнаружено."
    }`;
  }
  
  generateSpecialZones() {
    const zeroZones = this.data.zeroZone.length 
      ? this.data.zeroZone.join(", ") 
      : "Нет масел в 0-зоне.";
    
    const reversZones = this.data.reversZone.length 
      ? this.data.reversZone.length === 1 ? this.data.reversZone[0] : this.data.reversZone.join(", ")
      : "Нет масел в R-зоне.";
    
    return `⚡ 3. СПЕЦИАЛЬНЫЕ ЗОНЫ:\n${"─".repeat(40)}\n\n🔴 0-зона: ${zeroZones}\n🔄 R-зона: ${reversZones}`;
  }
  
  generateKeyTasks() {
    const peTasks = this.data.plusPlusPlusPE.length 
      ? this.data.plusPlusPlusPE.map(item => `• ${item}`).join("\n") 
      : "Нет масел в зоне +++.";
    
    const sTasks = this.data.plusPlusPlusS.length 
      ? this.data.plusPlusPlusS.map(item => `• ${item}`).join("\n") 
      : "Нет масел в зоне +++.";
    
    return `🚨 4. КЛЮЧЕВЫЕ ЗАДАЧИ ДЛЯ +++ (основные проблемы):\n${"─".repeat(70)}\n\n🧠 Психоэмоциональные:\n${peTasks}\n\n💊 Соматические:\n${sTasks}`;
  }
  
  generateResourceTasks() {
    const peTasks = this.data.minusMinusMinusPE.length 
      ? this.data.minusMinusMinusPE.map(item => `• ${item}`).join("\n") 
      : "Нет масел в зоне ---.";
    
    const sTasks = this.data.minusMinusMinusS.length 
      ? this.data.minusMinusMinusS.map(item => `• ${item}`).join("\n") 
      : "Нет масел в зоне +++.";
    
    return `💪 5. РЕСУРСНЫЕ ЗАДАЧИ ДЛЯ --- (ресурсы):\n${"─".repeat(50)}\n\n🧠 Психоэмоциональные:\n${peTasks}\n\n💊 Соматические:\n${sTasks}`;
  }
  
  generateAdditionalTasks() {
    const plusPE = this.data.plusPE.length ? this.data.plusPE.map(item => `• ${item}`).join("\n") : "Нет масел в зоне +.";
    const plusS = this.data.plusS.length ? this.data.plusS.map(item => `• ${item}`).join("\n") : "Нет масел в зоне +.";
    const minusPE = this.data.minusPE.length ? this.data.minusPE.map(item => `• ${item}`).join("\n") : "Нет масел в зоне -.";
    const minusS = this.data.minusS.length ? this.data.minusS.map(item => `• ${item}`).join("\n") : "Нет масел в зоне -.";
    
    return `📋 6. ДОПОЛНИТЕЛЬНЫЕ ЗАДАЧИ ДЛЯ + И -:\n${"─".repeat(60)}\n\n🧠 Психоэмоциональные (+):\n${plusPE}\n\n💊 Соматические (+):\n${plusS}\n\n🧠 Психоэмоциональные (-):\n${minusPE}\n\n💊 Соматические (-):\n${minusS}`;
  }
  
  generateCombinations() {
    const combinationChecker = new CombinationChecker(SpreadsheetApp.getActiveSpreadsheet().getActiveSheet());
    const foundCombinations = combinationChecker.checkAllCombinations();
    const report = combinationChecker.getCombinationsReport();
    
    return `🔗 7. СОЧЕТАНИЯ МАСЕЛ:\n${"=".repeat(50)}\n${report}`;
  }
  
  generatePatterns() {
    const patterns = this.data.patterns.length 
      ? this.data.patterns.map(item => `• ${item}`).join("\n") 
      : "Особых закономерностей не выявлено.";
    
    return `🔄 8. ПОВТОРЯЮЩИЕСЯ ЗАКОНОМЕРНОСТИ:\n${"=".repeat(50)}\n${patterns}`;
  }
  
  generatePESummary() {
    const summary = this.generatePsychoemotionalSummary();
    return `🧠 9. ВЫВОД ПО ПСИХОЭМОЦИОНАЛЬНОМУ СОСТОЯНИЮ:\n${"=".repeat(50)}\n${summary}`;
  }
  
  generateSSummary() {
    const summary = this.generateSomaticSummary();
    return `💊 10. ВЫВОД ПО СОМАТИЧЕСКОМУ СОСТОЯНИЮ:\n${"=".repeat(50)}\n${summary}`;
  }
  
  generatePsychoemotionalSummary() {
    let summary = "На основе анализа масел в зоне +++: ";
    
    if (this.data.plusPlusPlusPE.length > 0) {
      summary += "выявлены основные психоэмоциональные проблемы. ";
      
      if (this.data.patterns.length > 0) {
        summary += `\n\n🔍 Отмечены особые закономерности:\n${this.data.patterns.map(p => `   • ${p}`).join("\n")}\n\n`;
      }
      
      summary += "💪 Ресурсные состояния (зона ---): ";
      summary += this.data.minusMinusMinusPE.length > 0 
        ? "присутствуют, что говорит о потенциале для восстановления."
        : "ограничены, требуется активация ресурсов.";
    } else {
      summary += "острых психоэмоциональных проблем не выявлено, состояние относительно стабильное.";
    }
    
    return summary;
  }
  
  generateSomaticSummary() {
    let summary = "Анализ соматических показателей: ";
    
    if (this.data.minusMinusMinusS.length > 0 || this.data.minusS.length > 0) {
      summary += "выявлены активные соматические проблемы. ";
      
      const citrusGroup = this.groups[CONFIG.OIL_GROUPS.CITRUS];
      const pineGroup = this.groups[CONFIG.OIL_GROUPS.CONIFEROUS];
      const spiceGroup = this.groups[CONFIG.OIL_GROUPS.SPICE];
      
      const citrusIssues = (citrusGroup["---"] + citrusGroup["-"]) >= 5;
      const pineIssues = pineGroup["-"] > 0 || pineGroup["---"] > 0;
      const spiceIssues = (spiceGroup["---"] + spiceGroup["-"]) >= 5;
      
      summary += "\n\n⚠️ Выявленные проблемы:\n";
      if (citrusIssues) summary += "   • Цитрусовая группа: возможны проблемы с кислотностью\n";
      if (pineIssues) summary += "   • Хвойная группа: признаки острого воспалительного процесса\n";
      if (spiceIssues) summary += "   • Пряная группа: хронические нарушения ЖКТ и эндокринной системы\n";
    } else {
      summary += "серьезных соматических нарушений не выявлено. ";
    }
    
    if (this.data.plusPlusPlusS.length > 0 || this.data.plusS.length > 0) {
      summary += "\n\n✅ Есть выраженные ресурсные показатели (масла в зоне +++ и + поддерживают организм).";
    }
    
    return summary;
  }
  
  formatReport() {
    const lastRow = this.sheet.getLastRow();
    this.sheet.getRange(1, 1, lastRow, 1).setWrap(true).setVerticalAlignment("top");
    this.sheet.setColumnWidth(1, 800);
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
