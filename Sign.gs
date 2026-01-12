/**
 * Signature Verification Script - Using Installable Trigger
 * Logic: Find signature by active user's email in dt_sign_allData
 * Clean status feedback without popups
 */

// ============================================
// CONFIGURATION - Easy to modify
// ============================================
const CONFIG = {
  // Sheet names and ranges
  SHEET_NAME: 'dt_signs',                     // Main working sheet
  DATA_RANGE_NAME: 'dt_sign_allData',         // Named range for signature data
  
  // Column indexes (1 = A, 2 = B, etc.)
  CHECKBOX_COLUMN: 4,                         // Column D - Checkbox trigger
  STATUS_COLUMN: 3,                           // Column C - Status display
  
  // Row range for processing
  START_ROW: 20,                              // First row to process
  END_ROW: 60,                                // Last row to process
  
  // Default texts
  DEFAULT_STATUS_TEXT: 'Подпись',             // Default text in status column
  DEFAULT_FONT_COLOR: '#000000',              // Default black color
  
  // Status colors
  COLOR_PROCESSING: '#FFA500',                // Orange for processing
  COLOR_SUCCESS: '#28a745',                   // Green for success
  COLOR_ERROR: '#dc3545'                      // Red for error
};

// Property keys for storage
const PROPERTY_KEYS = {
  SHEET_NAME: 'SHEET_NAME',
  DATA_RANGE_NAME: 'DATA_RANGE_NAME',
  CHECKBOX_COLUMN: 'CHECKBOX_COLUMN',
  STATUS_COLUMN: 'STATUS_COLUMN',
  START_ROW: 'START_ROW',
  END_ROW: 'END_ROW',
  DEFAULT_STATUS_TEXT: 'DEFAULT_STATUS_TEXT',
  DEFAULT_FONT_COLOR: 'DEFAULT_FONT_COLOR',
  COLOR_PROCESSING: 'COLOR_PROCESSING',
  COLOR_SUCCESS: 'COLOR_SUCCESS',
  COLOR_ERROR: 'COLOR_ERROR'
};

/**
 * Gets configuration - reads from Properties or uses defaults
 */
function getConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const config = {};
  
  // Read all properties or use defaults
  config.SHEET_NAME = scriptProperties.getProperty(PROPERTY_KEYS.SHEET_NAME) || CONFIG.SHEET_NAME;
  config.DATA_RANGE_NAME = scriptProperties.getProperty(PROPERTY_KEYS.DATA_RANGE_NAME) || CONFIG.DATA_RANGE_NAME;
  config.CHECKBOX_COLUMN = parseInt(scriptProperties.getProperty(PROPERTY_KEYS.CHECKBOX_COLUMN)) || CONFIG.CHECKBOX_COLUMN;
  config.STATUS_COLUMN = parseInt(scriptProperties.getProperty(PROPERTY_KEYS.STATUS_COLUMN)) || CONFIG.STATUS_COLUMN;
  config.START_ROW = parseInt(scriptProperties.getProperty(PROPERTY_KEYS.START_ROW)) || CONFIG.START_ROW;
  config.END_ROW = parseInt(scriptProperties.getProperty(PROPERTY_KEYS.END_ROW)) || CONFIG.END_ROW;
  config.DEFAULT_STATUS_TEXT = scriptProperties.getProperty(PROPERTY_KEYS.DEFAULT_STATUS_TEXT) || CONFIG.DEFAULT_STATUS_TEXT;
  config.DEFAULT_FONT_COLOR = scriptProperties.getProperty(PROPERTY_KEYS.DEFAULT_FONT_COLOR) || CONFIG.DEFAULT_FONT_COLOR;
  config.COLOR_PROCESSING = scriptProperties.getProperty(PROPERTY_KEYS.COLOR_PROCESSING) || CONFIG.COLOR_PROCESSING;
  config.COLOR_SUCCESS = scriptProperties.getProperty(PROPERTY_KEYS.COLOR_SUCCESS) || CONFIG.COLOR_SUCCESS;
  config.COLOR_ERROR = scriptProperties.getProperty(PROPERTY_KEYS.COLOR_ERROR) || CONFIG.COLOR_ERROR;
  
  return config;
}

/**
 * Saves configuration to Properties
 */
function saveConfig(newConfig) {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Save all properties
  scriptProperties.setProperty(PROPERTY_KEYS.SHEET_NAME, newConfig.SHEET_NAME);
  scriptProperties.setProperty(PROPERTY_KEYS.DATA_RANGE_NAME, newConfig.DATA_RANGE_NAME);
  scriptProperties.setProperty(PROPERTY_KEYS.CHECKBOX_COLUMN, newConfig.CHECKBOX_COLUMN.toString());
  scriptProperties.setProperty(PROPERTY_KEYS.STATUS_COLUMN, newConfig.STATUS_COLUMN.toString());
  scriptProperties.setProperty(PROPERTY_KEYS.START_ROW, newConfig.START_ROW.toString());
  scriptProperties.setProperty(PROPERTY_KEYS.END_ROW, newConfig.END_ROW.toString());
  scriptProperties.setProperty(PROPERTY_KEYS.DEFAULT_STATUS_TEXT, newConfig.DEFAULT_STATUS_TEXT);
  scriptProperties.setProperty(PROPERTY_KEYS.DEFAULT_FONT_COLOR, newConfig.DEFAULT_FONT_COLOR);
  scriptProperties.setProperty(PROPERTY_KEYS.COLOR_PROCESSING, newConfig.COLOR_PROCESSING);
  scriptProperties.setProperty(PROPERTY_KEYS.COLOR_SUCCESS, newConfig.COLOR_SUCCESS);
  scriptProperties.setProperty(PROPERTY_KEYS.COLOR_ERROR, newConfig.COLOR_ERROR);
  
  return { success: true, message: 'Configuration saved successfully!' };
}

/**
 * Resets configuration to defaults
 */
function resetConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  
  // Clear all properties
  Object.values(PROPERTY_KEYS).forEach(key => {
    scriptProperties.deleteProperty(key);
  });
  
  return { success: true, message: 'Configuration reset to defaults!' };
}

/**
 * Run this function ONCE to install the trigger
 */
function installTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEditInstallable') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
    
  SpreadsheetApp.getActiveSpreadsheet().toast('Триггер успешно установлен!', '✓ Готово', 3);
}

/**
 * Installable onEdit trigger
 */
function onEditInstallable(e) {
  const config = getConfig();
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Check if edit is in the correct sheet
  if (sheet.getName() !== config.SHEET_NAME) return;
  
  // Check if edit is in the checkbox column
  if (range.getColumn() !== config.CHECKBOX_COLUMN) return;
  
  // Check if checkbox is checked (true)
  if (range.getValue() !== true) return;
  
  const row = range.getRow();
  
  // Check if row is within the processing range
  if (row < config.START_ROW || row > config.END_ROW) return;
  
  verifyAndSign(sheet, row, config);
}

/**
 * Verifies the active user and places their signature
 */
function verifyAndSign(sheet, row, config) {
  const statusCell = sheet.getRange(row, config.STATUS_COLUMN);
  const checkboxCell = sheet.getRange(row, config.CHECKBOX_COLUMN);
  
  try {
    // ═══════════════════════════════════════════════════════
    // STEP 1: Initialize
    // ═══════════════════════════════════════════════════════
    setStatus(statusCell, '🔄 Проверка...', config.COLOR_PROCESSING);
    
    // ═══════════════════════════════════════════════════════
    // STEP 2: Get active user email
    // ═══════════════════════════════════════════════════════
    const activeUserEmail = Session.getActiveUser().getEmail().toLowerCase();
    
    if (!activeUserEmail) {
      handleError(checkboxCell, statusCell, '❌ Ошибка авторизации', config);
      return;
    }
    
    setStatus(statusCell, '🔍 Поиск подписи...', config.COLOR_PROCESSING);
    
    // ═══════════════════════════════════════════════════════
    // STEP 3: Search in reference data using named range
    // ═══════════════════════════════════════════════════════
    const dataRange = sheet.getParent().getRangeByName(config.DATA_RANGE_NAME);
    
    if (!dataRange) {
      handleError(checkboxCell, statusCell, '❌ Набор данных не найден', config);
      return;
    }
    
    const data = dataRange.getValues();
    
    let userSignature = null;
    let userName = null;
    
    // Assuming named range has columns: ID, Name, Email, Signature
    for (let i = 0; i < data.length; i++) {
      const dataEmail = data[i][2]; // Email is in 3rd column (index 2)
      
      if (dataEmail && dataEmail.toString().toLowerCase() === activeUserEmail) {
        userSignature = data[i][3]; // Signature is in 4th column (index 3)
        userName = data[i][1];      // Name is in 2nd column (index 1)
        break;
      }
    }
    
    if (!userSignature) {
      handleError(checkboxCell, statusCell, '❌ Email не найден', config);
      return;
    }
    
    setStatus(statusCell, '🔐 Проверка прав...', config.COLOR_PROCESSING);
    
    // ═══════════════════════════════════════════════════════
    // STEP 4: Verify signer matches expected person
    // ═══════════════════════════════════════════════════════
    const expectedName = sheet.getRange(row, 2).getValue(); // Column B
    
    if (!nameMatches(expectedName, userName)) {
      handleError(checkboxCell, statusCell, '❌ Нет прав на подпись', config);
      return;
    }
    
    // ═══════════════════════════════════════════════════════
    // SUCCESS!
    // ═══════════════════════════════════════════════════════
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
    setStatus(statusCell, '✅ Подписано: ' + timestamp, config.COLOR_SUCCESS);
    
  } catch (error) {
    handleError(checkboxCell, statusCell, '❌ Ошибка', config);
    console.error('Signature error:', error);
  }
}

/**
 * Sets status cell with color
 */
function setStatus(cell, text, color) {
  cell.setValue(text).setFontColor(color);
  SpreadsheetApp.flush();
}

/**
 * Handles errors: resets checkbox, shows error, resets to default text after 5 sec
 */
function handleError(checkboxCell, statusCell, errorText, config) {
  checkboxCell.setValue(false);
  setStatus(statusCell, errorText, config.COLOR_ERROR);
  
  // Wait 5 seconds then reset to default text
  Utilities.sleep(5000);
  statusCell.setValue(config.DEFAULT_STATUS_TEXT)
            .setFontColor(config.DEFAULT_FONT_COLOR);
  SpreadsheetApp.flush();
}

/**
 * Checks if two names match
 */
function nameMatches(shortName, fullName) {
  if (!shortName || !fullName) return false;
  
  const short = shortName.toString().trim().toLowerCase();
  const full = fullName.toString().trim().toLowerCase();
  
  // Exact match
  if (short === full) return true;
  
  // Check surname match (first word)
  const shortSurname = short.split(/[\s\.]+/)[0];
  const fullSurname = full.split(/[\s\.]+/)[0];
  
  return shortSurname === fullSurname;
}

/**
 * Test function to verify configuration
 */
function testConfiguration() {
  const config = getConfig();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(config.SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`Sheet "${config.SHEET_NAME}" not found!`);
  }
  
  const dataRange = ss.getRangeByName(config.DATA_RANGE_NAME);
  if (!dataRange) {
    throw new Error(`Named range "${config.DATA_RANGE_NAME}" not found!`);
  }
  
  console.log('Configuration test passed:');
  console.log(`- Sheet: ${config.SHEET_NAME} ✓`);
  console.log(`- Named range: ${config.DATA_RANGE_NAME} ✓`);
  console.log(`- Data range size: ${dataRange.getNumRows()} rows x ${dataRange.getNumColumns()} columns`);
  console.log(`- Processing rows: ${config.START_ROW} to ${config.END_ROW}`);
  console.log(`- Checkbox column: ${String.fromCharCode(64 + config.CHECKBOX_COLUMN)}`);
  console.log(`- Status column: ${String.fromCharCode(64 + config.STATUS_COLUMN)}`);
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Конфигурация проверена успешно!', '✓ Тест', 5);
  
  return {
    success: true,
    message: 'Configuration test passed!',
    details: {
      sheet: config.SHEET_NAME,
      dataRange: config.DATA_RANGE_NAME,
      rows: `${config.START_ROW}-${config.END_ROW}`,
      checkboxColumn: `${String.fromCharCode(64 + config.CHECKBOX_COLUMN)} (${config.CHECKBOX_COLUMN})`,
      statusColumn: `${String.fromCharCode(64 + config.STATUS_COLUMN)} (${config.STATUS_COLUMN})`
    }
  };
}

// ============================================
// HTML CONFIGURATION INTERFACE
// ============================================

/**
 * Creates a custom menu for the configuration interface
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('⚙️ Подпись')
    .addItem('📋 Конфигурация', 'showConfigPanel')
    .addSeparator()
    .addItem('🔧 Установить триггер', 'installTrigger')
    .addItem('✅ Проверить конфигурацию', 'testConfiguration')
    .addToUi();
}

/**
 * Shows the configuration panel
 */
function showConfigPanel() {
  const html = HtmlService.createHtmlOutputFromFile('ConfigurationPanel')
    .setWidth(500)
    .setHeight(600)
    .setTitle('Конфигурация подписи');
  
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Loads current configuration for the HTML panel
 */
function loadConfigForPanel() {
  return getConfig();
}