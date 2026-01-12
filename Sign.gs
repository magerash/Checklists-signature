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
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  // Check if edit is in the correct sheet
  if (sheet.getName() !== CONFIG.SHEET_NAME) return;
  
  // Check if edit is in the checkbox column
  if (range.getColumn() !== CONFIG.CHECKBOX_COLUMN) return;
  
  // Check if checkbox is checked (true)
  if (range.getValue() !== true) return;
  
  const row = range.getRow();
  
  // Check if row is within the processing range
  if (row < CONFIG.START_ROW || row > CONFIG.END_ROW) return;
  
  verifyAndSign(sheet, row);
}

/**
 * Verifies the active user and places their signature
 */
function verifyAndSign(sheet, row) {
  const statusCell = sheet.getRange(row, CONFIG.STATUS_COLUMN);
  const checkboxCell = sheet.getRange(row, CONFIG.CHECKBOX_COLUMN);
  
  try {
    // ═══════════════════════════════════════════════════════
    // STEP 1: Initialize
    // ═══════════════════════════════════════════════════════
    setStatus(statusCell, '🔄 Проверка...', CONFIG.COLOR_PROCESSING);
    
    // ═══════════════════════════════════════════════════════
    // STEP 2: Get active user email
    // ═══════════════════════════════════════════════════════
    const activeUserEmail = Session.getActiveUser().getEmail().toLowerCase();
    
    if (!activeUserEmail) {
      handleError(checkboxCell, statusCell, '❌ Ошибка авторизации');
      return;
    }
    
    setStatus(statusCell, '🔍 Поиск подписи...', CONFIG.COLOR_PROCESSING);
    
    // ═══════════════════════════════════════════════════════
    // STEP 3: Search in reference data using named range
    // ═══════════════════════════════════════════════════════
    const dataRange = sheet.getParent().getRangeByName(CONFIG.DATA_RANGE_NAME);
    
    if (!dataRange) {
      handleError(checkboxCell, statusCell, '❌ Набор данных не найден');
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
      handleError(checkboxCell, statusCell, '❌ Email не найден');
      return;
    }
    
    setStatus(statusCell, '🔐 Проверка прав...', CONFIG.COLOR_PROCESSING);
    
    // ═══════════════════════════════════════════════════════
    // STEP 4: Verify signer matches expected person
    // ═══════════════════════════════════════════════════════
    const expectedName = sheet.getRange(row, 2).getValue(); // Column B
    
    if (!nameMatches(expectedName, userName)) {
      handleError(checkboxCell, statusCell, '❌ Нет прав на подпись');
      return;
    }
    
    // ═══════════════════════════════════════════════════════
    // SUCCESS!
    // ═══════════════════════════════════════════════════════
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
    setStatus(statusCell, '✅ Подписано: ' + timestamp, CONFIG.COLOR_SUCCESS);
    
  } catch (error) {
    handleError(checkboxCell, statusCell, '❌ Ошибка');
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
function handleError(checkboxCell, statusCell, errorText) {
  checkboxCell.setValue(false);
  setStatus(statusCell, errorText, CONFIG.COLOR_ERROR);
  
  // Wait 5 seconds then reset to default text
  Utilities.sleep(5000);
  statusCell.setValue(CONFIG.DEFAULT_STATUS_TEXT)
            .setFontColor(CONFIG.DEFAULT_FONT_COLOR);
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
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  
  if (!sheet) {
    throw new Error(`Sheet "${CONFIG.SHEET_NAME}" not found!`);
  }
  
  const dataRange = ss.getRangeByName(CONFIG.DATA_RANGE_NAME);
  if (!dataRange) {
    throw new Error(`Named range "${CONFIG.DATA_RANGE_NAME}" not found!`);
  }
  
  console.log('Configuration test passed:');
  console.log(`- Sheet: ${CONFIG.SHEET_NAME} ✓`);
  console.log(`- Named range: ${CONFIG.DATA_RANGE_NAME} ✓`);
  console.log(`- Data range size: ${dataRange.getNumRows()} rows x ${dataRange.getNumColumns()} columns`);
  console.log(`- Processing rows: ${CONFIG.START_ROW} to ${CONFIG.END_ROW}`);
  console.log(`- Checkbox column: ${String.fromCharCode(64 + CONFIG.CHECKBOX_COLUMN)}`);
  console.log(`- Status column: ${String.fromCharCode(64 + CONFIG.STATUS_COLUMN)}`);
  
  SpreadsheetApp.getActiveSpreadsheet().toast('Конфигурация проверена успешно!', '✓ Тест', 5);
}