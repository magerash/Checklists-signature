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
  SHEET_NAME: 'II. Чек-лист стадии концепции',  // Main working sheet
  DATA_RANGE_NAME: 'dt_sign_allData',         // Named range for signature data
  
  // Column indexes (1 = A, 2 = B, etc.)
  CHECKBOX_COLUMNS: '12,22',                      // Column(s) with checkboxes
  STATUS_COLUMNS: '11,21',                        // Column(s) for status display
  NAME_COLUMNS: '10,20',                          // NEW: Column(s) for name verification
  
  // Row range for processing
  ROW_RANGES: '5-13',                        // Row range(s) for processing
  
  // Default texts
  DEFAULT_STATUS_TEXT: 'Подпись',             // Default text in status column
  DEFAULT_FONT_COLOR: '#999999',              // Default gray color
  
  // Status colors
  COLOR_PROCESSING: '#FFA500',                // Orange for processing
  COLOR_SUCCESS: '#28a745',                   // Green for success
  COLOR_ERROR: '#dc3545'                      // Red for error
};

const PROPERTY_KEYS = {
  SHEET_NAME: 'SHEET_NAME',
  DATA_RANGE_NAME: 'DATA_RANGE_NAME',
  CHECKBOX_COLUMNS: 'CHECKBOX_COLUMNS',
  STATUS_COLUMNS: 'STATUS_COLUMNS',
  NAME_COLUMNS: 'NAME_COLUMNS',  // NEW
  ROW_RANGES: 'ROW_RANGES',
  DEFAULT_STATUS_TEXT: 'DEFAULT_STATUS_TEXT',
  DEFAULT_FONT_COLOR: 'DEFAULT_FONT_COLOR',
  COLOR_PROCESSING: 'COLOR_PROCESSING',
  COLOR_SUCCESS: 'COLOR_SUCCESS',
  COLOR_ERROR: 'COLOR_ERROR'
};

/**
 * Helper function to parse comma-separated or range strings into arrays
 * Supports formats like: "4", "4,5,6", "4-6", "20-30,35-40"
 */
function parseRangeString(rangeString) {
  if (!rangeString) return [];
  
  const result = [];
  const parts = rangeString.split(',').map(part => part.trim());
  
  for (const part of parts) {
    if (part.includes('-')) {
      // Handle range like "20-30"
      const [start, end] = part.split('-').map(num => parseInt(num.trim()));
      if (!isNaN(start) && !isNaN(end) && start <= end) {
        for (let i = start; i <= end; i++) {
          result.push(i);
        }
      }
    } else {
      // Handle single number
      const num = parseInt(part);
      if (!isNaN(num)) {
        result.push(num);
      }
    }
  }
  
  // Remove duplicates and sort
  return [...new Set(result)].sort((a, b) => a - b);
}

/**
 * Helper function to check if a value is in any of the parsed ranges
 */
function isInRanges(value, rangeString) {
  const ranges = parseRangeString(rangeString);
  return ranges.includes(value);
}

/**
 * Helper function to get corresponding status column for a checkbox column
 * Assumes checkbox and status columns are paired in order
 */
function getStatusColumnForCheckbox(checkboxColumn, config) {
  const checkboxCols = parseRangeString(config.CHECKBOX_COLUMNS);
  const statusCols = parseRangeString(config.STATUS_COLUMNS);
  
  const index = checkboxCols.indexOf(checkboxColumn);
  if (index !== -1 && index < statusCols.length) {
    return statusCols[index];
  }
  
  // Fallback: if only one status column defined, use it for all checkboxes
  if (statusCols.length === 1) {
    return statusCols[0];
  }
  
  // Default: assume status column is checkbox column - 1
  return checkboxColumn - 1;
}

/**
 * Helper function to get corresponding name column for a checkbox column
 */
function getNameColumnForCheckbox(checkboxColumn, config) {
  const checkboxCols = parseRangeString(config.CHECKBOX_COLUMNS);
  const nameCols = parseRangeString(config.NAME_COLUMNS);
  
  const index = checkboxCols.indexOf(checkboxColumn);
  if (index !== -1 && index < nameCols.length) {
    return nameCols[index];
  }
  
  // Fallback: if only one name column defined, use it for all checkboxes
  if (nameCols.length === 1) {
    return nameCols[0];
  }
  
  // Default: assume name column is checkbox column - 2 (as per your requirement)
  return checkboxColumn - 2;
}

/**
 * Gets configuration - reads from Properties or uses defaults
 */
function getConfig() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const config = {};
  
  // Read all properties or use defaults
  config.SHEET_NAME = scriptProperties.getProperty(PROPERTY_KEYS.SHEET_NAME) || CONFIG.SHEET_NAME;
  config.DATA_RANGE_NAME = scriptProperties.getProperty(PROPERTY_KEYS.DATA_RANGE_NAME) || CONFIG.DATA_RANGE_NAME;
  config.CHECKBOX_COLUMNS = scriptProperties.getProperty(PROPERTY_KEYS.CHECKBOX_COLUMNS) || CONFIG.CHECKBOX_COLUMNS;
  config.STATUS_COLUMNS = scriptProperties.getProperty(PROPERTY_KEYS.STATUS_COLUMNS) || CONFIG.STATUS_COLUMNS;
  config.NAME_COLUMNS = scriptProperties.getProperty(PROPERTY_KEYS.NAME_COLUMNS) || CONFIG.NAME_COLUMNS;  // NEW
  config.ROW_RANGES = scriptProperties.getProperty(PROPERTY_KEYS.ROW_RANGES) || CONFIG.ROW_RANGES;
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
  scriptProperties.setProperty(PROPERTY_KEYS.CHECKBOX_COLUMNS, newConfig.CHECKBOX_COLUMNS.toString());
  scriptProperties.setProperty(PROPERTY_KEYS.STATUS_COLUMNS, newConfig.STATUS_COLUMNS.toString());
  scriptProperties.setProperty(PROPERTY_KEYS.NAME_COLUMNS, newConfig.NAME_COLUMNS.toString());  // NEW
  scriptProperties.setProperty(PROPERTY_KEYS.ROW_RANGES, newConfig.ROW_RANGES.toString());
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
  
  // Check if edit is in one of the checkbox columns
  const checkboxColumn = range.getColumn();
  if (!isInRanges(checkboxColumn, config.CHECKBOX_COLUMNS)) return;
  
  // Check if checkbox is checked (true)
  if (range.getValue() !== true) return;
  
  const row = range.getRow();
  
  // Check if row is within the processing ranges
  if (!isInRanges(row, config.ROW_RANGES)) return;
  
  verifyAndSign(sheet, row, checkboxColumn, config);
}

/**
 * Verifies the active user and places their signature
 */
function verifyAndSign(sheet, row, checkboxColumn, config) {
  const statusColumn = getStatusColumnForCheckbox(checkboxColumn, config);
  const statusCell = sheet.getRange(row, statusColumn);
  const checkboxCell = sheet.getRange(row, checkboxColumn);
  
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
    const nameColumn = getNameColumnForCheckbox(checkboxColumn, config);
    const expectedName = sheet.getRange(row, nameColumn).getValue();

    if (!expectedName || expectedName.toString().trim() === '') {
      handleError(checkboxCell, statusCell, '❌ Имя не указано', config);
      return;
    }

    if (!nameMatches(expectedName, userName)) {
      handleError(checkboxCell, statusCell, '❌ Нет прав на подпись', config);
      return;
    }   

    // ═══════════════════════════════════════════════════════
    // SUCCESS!
    // ═══════════════════════════════════════════════════════
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
    setStatus(statusCell, ' ' + timestamp, config.COLOR_SUCCESS);
    
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
  
  // Parse the ranges for display
  const checkboxCols = parseRangeString(config.CHECKBOX_COLUMNS);
  const statusCols = parseRangeString(config.STATUS_COLUMNS);
  const rowRanges = parseRangeString(config.ROW_RANGES);
  
  console.log('Configuration test passed:');
  console.log(`- Sheet: ${config.SHEET_NAME} ✓`);
  console.log(`- Named range: ${config.DATA_RANGE_NAME} ✓`);
  console.log(`- Data range size: ${dataRange.getNumRows()} rows x ${dataRange.getNumColumns()} columns`);
  console.log(`- Checkbox columns: ${checkboxCols.map(c => String.fromCharCode(64 + c)).join(', ')} (${checkboxCols.join(', ')})`);
  console.log(`- Status columns: ${statusCols.map(c => String.fromCharCode(64 + c)).join(', ')} (${statusCols.join(', ')})`);
  console.log(`- Processing rows: ${rowRanges.join(', ')}`);
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Конфигурация проверена успешно!', '✓ Тест', 5);
  
  return {
    success: true,
    message: 'Configuration test passed!',
    details: {
      sheet: config.SHEET_NAME,
      dataRange: config.DATA_RANGE_NAME,
      checkboxColumns: `${checkboxCols.map(c => String.fromCharCode(64 + c)).join(', ')} (${config.CHECKBOX_COLUMNS})`,
      statusColumns: `${statusCols.map(c => String.fromCharCode(64 + c)).join(', ')} (${config.STATUS_COLUMNS})`,
      rowRanges: rowRanges.join(', ') + ` (${config.ROW_RANGES})`
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
    .setHeight(650)
    .setTitle('Конфигурация подписи');
  
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Loads current configuration for the HTML panel
 */
function loadConfigForPanel() {
  return getConfig();
}