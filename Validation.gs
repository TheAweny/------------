// ==================== ВАЛИДАЦИЯ И ОБРАБОТКА ВВОДА ====================

/**
 * Класс для валидации редактирования ячеек
 * Обрабатывает события редактирования и проверяет корректность данных
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
   * Проверяет, требует ли это редактирование обработки
   * @returns {boolean} true если нужно обрабатывать
   */
  shouldProcess() {
    return this.sheet.getName() === CONFIG.SHEETS.INPUT && 
           this.row > 1 && 
           (this.col === CONFIG.COLUMNS.ZONE + 1 || this.col === CONFIG.COLUMNS.TROIKA + 1);
  }
  
  /**
   * Выполняет валидацию и обработку редактирования
   */
  validateAndProcess() {
    try {
      if (this.col === CONFIG.COLUMNS.TROIKA + 1) {
        this.validateTroika();
      }
      this.updateCellFormatting();
    } catch (error) {
      console.error(`Ошибка валидации: ${error.message}`);
      this.showError(`Ошибка валидации: ${error.message}`);
    }
  }
  
  /**
   * Валидация значений в столбце "Тройка"
   */
  validateTroika() {
    const value = Utils.getCellValue(this.editedCell);
    
    if (!Utils.isNotEmpty(value)) return;
    
    // Проверка допустимых значений
    if (!CONFIG.VALID_TROIKA_VALUES.includes(value)) {
      this.showError(CONFIG.MESSAGES.INVALID_TROIKA);
      this.editedCell.clearContent();
      return;
    }
    
    // Проверка уникальности в зоне "+++"
    const zoneValue = Utils.getCellValue(this.sheet.getRange(this.row, CONFIG.COLUMNS.ZONE + 1));
    if (zoneValue === "+++") {
      this.validateUniqueTroika(value);
    }
  }
  
  /**
   * Проверяет уникальность значения "Тройки" в зоне "+++"
   * @param {string} editedValue - Редактируемое значение
   */
  validateUniqueTroika(editedValue) {
    const lastRow = this.sheet.getLastRow();
    
    for (let row = 2; row <= lastRow; row++) {
      if (row === this.row) continue;
      
      const zoneCell = this.sheet.getRange(row, CONFIG.COLUMNS.ZONE + 1);
      const troikaCell = this.sheet.getRange(row, CONFIG.COLUMNS.TROIKA + 1);
      
      const zone = Utils.getCellValue(zoneCell);
      const troika = Utils.getCellValue(troikaCell);
      
      if (zone === "+++" && Utils.isNotEmpty(troika) && troika === editedValue) {
        this.showError(`${CONFIG.MESSAGES.DUPLICATE_TROIKA} ${editedValue}`);
        this.editedCell.clearContent();
        return;
      }
    }
  }
  
  /**
   * Обновляет форматирование ячеек в зависимости от зоны
   */
  updateCellFormatting() {
    const lastRow = this.sheet.getLastRow();
    
    for (let row = 2; row <= lastRow; row++) {
      const zoneValue = Utils.getCellValue(this.sheet.getRange(row, CONFIG.COLUMNS.ZONE + 1));
      const troikaCell = this.sheet.getRange(row, CONFIG.COLUMNS.TROIKA + 1);
      
      if (zoneValue === "+++") {
        this.setupTroikaValidation(troikaCell);
      } else {
        this.clearTroikaCell(troikaCell);
      }
    }
  }
  
  /**
   * Настраивает валидацию для ячейки "Тройка"
   * @param {Range} cell - Ячейка для настройки
   */
  setupTroikaValidation(cell) {
    cell.setBackground("white");
    
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(CONFIG.VALID_TROIKA_VALUES, true)
      .setAllowInvalid(false)
      .setHelpText("Выберите 1, 2 или 3")
      .build();
      
    cell.setDataValidation(rule);
  }
  
  /**
   * Очищает ячейку "Тройка" и устанавливает серый фон
   * @param {Range} cell - Ячейка для очистки
   */
  clearTroikaCell(cell) {
    cell.clearDataValidations();
    cell.clearContent();
    cell.setBackground("#d3d3d3");
  }
  
  /**
   * Показывает сообщение об ошибке пользователю
   * @param {string} message - Сообщение об ошибке
   */
  showError(message) {
    SpreadsheetApp.getActiveSpreadsheet().toast(message, "Ошибка", 3);
  }
}
