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
  CHECKBOX_COLUMNS: '13,24',                      // Column(s) with checkboxes
  STATUS_COLUMNS: '12,23',                        // Column(s) for status display
  NAME_COLUMNS: '10,21',                          // Column(s) for name verification
  
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

// Property keys for storage
const PROPERTY_KEYS = {
  SHEET_NAME: 'SHEET_NAME',
  DATA_RANGE_NAME: 'DATA_RANGE_NAME',
  CHECKBOX_COLUMNS: 'CHECKBOX_COLUMNS',
  STATUS_COLUMNS: 'STATUS_COLUMNS',
  NAME_COLUMNS: 'NAME_COLUMNS',
  ROW_RANGES: 'ROW_RANGES',
  DEFAULT_STATUS_TEXT: 'DEFAULT_STATUS_TEXT',
  DEFAULT_FONT_COLOR: 'DEFAULT_FONT_COLOR',
  COLOR_PROCESSING: 'COLOR_PROCESSING',
  COLOR_SUCCESS: 'COLOR_SUCCESS',
  COLOR_ERROR: 'COLOR_ERROR'
};

// Cache for storing last valid status values
const CACHE = CacheService.getScriptCache();
const CACHE_PREFIX = 'STATUS_';

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
 * Helper function to get corresponding name column for a status column
 */
function getNameColumnForStatus(statusColumn, config) {
  const statusCols = parseRangeString(config.STATUS_COLUMNS);
  const nameCols = parseRangeString(config.NAME_COLUMNS);
  
  const index = statusCols.indexOf(statusColumn);
  if (index !== -1 && index < nameCols.length) {
    return nameCols[index];
  }
  
  // Fallback: if only one name column defined, use it for all status columns
  if (nameCols.length === 1) {
    return nameCols[0];
  }
  
  // Default: assume name column is status column - 1
  return statusColumn - 1;
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
  config.NAME_COLUMNS = scriptProperties.getProperty(PROPERTY_KEYS.NAME_COLUMNS) || CONFIG.NAME_COLUMNS;
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
  scriptProperties.setProperty(PROPERTY_KEYS.NAME_COLUMNS, newConfig.NAME_COLUMNS.toString());
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
 * Check if current user is authorized to modify this status cell
 */
function isUserAuthorizedForStatus(sheet, row, statusColumn, config) {
  try {
    const activeUserEmail = Session.getActiveUser().getEmail().toLowerCase();
    
    if (!activeUserEmail) {
      console.log('No active user email found');
      return false;
    }
    
    // Get the corresponding name column
    const nameColumn = getNameColumnForStatus(statusColumn, config);
    const expectedName = sheet.getRange(row, nameColumn).getValue();
    
    if (!expectedName || expectedName.toString().trim() === '') {
      console.log('No expected name in name column');
      return false;
    }
    
    // Search for user in reference data
    const dataRange = sheet.getParent().getRangeByName(config.DATA_RANGE_NAME);
    
    if (!dataRange) {
      console.log('Data range not found');
      return false;
    }
    
    const data = dataRange.getValues();
    
    for (let i = 0; i < data.length; i++) {
      const dataEmail = data[i][2]; // Email is in 3rd column (index 2)
      
      if (dataEmail && dataEmail.toString().toLowerCase() === activeUserEmail) {
        const userName = data[i][1]; // Name is in 2nd column (index 1)
        
        // Check if names match
        if (nameMatches(expectedName, userName)) {
          console.log('User authorized:', activeUserEmail);
          return true;
        } else {
          console.log('Name mismatch:', expectedName, 'vs', userName);
          return false;
        }
      }
    }
    
    console.log('User email not found in data range:', activeUserEmail);
    return false;
    
  } catch (error) {
    console.error('Error checking authorization:', error);
    return false;
  }
}

/**
 * Get cache key for a status cell
 */
function getCacheKey(sheet, row, column) {
  return `${CACHE_PREFIX}${sheet.getSheetId()}_${row}_${column}`;
}

/**
 * Cache the current value of a status cell
 */
function cacheStatusValue(sheet, row, column, value) {
  const key = getCacheKey(sheet, row, column);
  if (value && value !== CONFIG.DEFAULT_STATUS_TEXT) {
    CACHE.put(key, value, 21600); // 6 hours
    console.log('Cached value:', value, 'for key:', key);
  } else {
    CACHE.remove(key);
    console.log('Removed cache for key:', key);
  }
}

/**
 * Get cached value for a status cell
 */
function getCachedStatusValue(sheet, row, column) {
  const key = getCacheKey(sheet, row, column);
  return CACHE.get(key);
}

/**
 * Run this function ONCE to install the triggers
 */
function installTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  
  // Remove existing triggers
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEditInstallable' || 
        trigger.getHandlerFunction() === 'onEditStatusProtection') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Install checkbox trigger
  ScriptApp.newTrigger('onEditInstallable')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
    
  // Install status protection trigger
  ScriptApp.newTrigger('onEditStatusProtection')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
    
  SpreadsheetApp.getActiveSpreadsheet().toast('Триггеры успешно установлены!', '✓ Готово', 3);
}

/**
 * Installable onEdit trigger for checkbox handling
 */
function onEditInstallable(e) {
  try {
    console.log('=== onEditInstallable TRIGGERED ===');
    console.log('Event details:', {
      sheet: e.source.getActiveSheet().getName(),
      range: e.range.getA1Notation(),
      row: e.range.getRow(),
      column: e.range.getColumn(),
      value: e.range.getValue()
    });
    
    const config = getConfig();
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const sheetName = sheet.getName();
    
    console.log('Config loaded. Target sheet:', config.SHEET_NAME);
    console.log('Current sheet:', sheetName);
    
    // Check if edit is in the correct sheet
    if (sheetName !== config.SHEET_NAME) {
      console.log('❌ Wrong sheet. Expected:', config.SHEET_NAME, 'Got:', sheetName);
      return;
    }
    console.log('✓ Correct sheet');
    
    // Check if edit is in one of the checkbox columns
    const checkboxColumn = range.getColumn();
    console.log('Checkbox column:', checkboxColumn);
    console.log('Allowed checkbox columns:', config.CHECKBOX_COLUMNS);
    
    if (!isInRanges(checkboxColumn, config.CHECKBOX_COLUMNS)) {
      console.log('❌ Not a checkbox column');
      return;
    }
    console.log('✓ Valid checkbox column');
    
    const row = range.getRow();
    console.log('Row:', row);
    console.log('Allowed row ranges:', config.ROW_RANGES);
    
    // Check if row is within the processing ranges
    if (!isInRanges(row, config.ROW_RANGES)) {
      console.log('❌ Row not in processing range');
      return;
    }
    console.log('✓ Row in processing range');
    
    const checkboxValue = range.getValue();
    console.log('Checkbox value:', checkboxValue);
    
    // Check if checkbox is checked (true) or unchecked (false)
    if (checkboxValue === true) {
      console.log('✓ Checkbox checked - starting verification...');
      verifyAndSign(sheet, row, checkboxColumn, config);
    } else {
      console.log('✓ Checkbox unchecked - handling uncheck...');
      handleCheckboxUnchecked(sheet, row, checkboxColumn, config);
    }
    
    console.log('=== onEditInstallable COMPLETED ===');
    
  } catch (error) {
    console.error('❌ ERROR in onEditInstallable:', error);
    console.error('Stack trace:', error.stack);
  }
}

/**
 * Installable onEdit trigger for status field protection
 * Prevents unauthorized deletion/editing of status fields
 */
function onEditStatusProtection(e) {
  try {
    console.log('=== onEditStatusProtection TRIGGERED ===');
    console.log('Event details:', {
      sheet: e.source.getActiveSheet().getName(),
      range: e.range.getA1Notation(),
      row: e.range.getRow(),
      column: e.range.getColumn(),
      value: e.range.getValue()
    });
    
    const config = getConfig();
    const sheet = e.source.getActiveSheet();
    const range = e.range;
    const sheetName = sheet.getName();
    
    console.log('Config loaded. Target sheet:', config.SHEET_NAME);
    console.log('Current sheet:', sheetName);
    
    // Check if edit is in the correct sheet
    if (sheetName !== config.SHEET_NAME) {
      console.log('❌ Wrong sheet. Expected:', config.SHEET_NAME, 'Got:', sheetName);
      return;
    }
    console.log('✓ Correct sheet');
    
    // Check if edit is in one of the status columns
    const statusColumn = range.getColumn();
    console.log('Status column:', statusColumn);
    console.log('Allowed status columns:', config.STATUS_COLUMNS);
    
    if (!isInRanges(statusColumn, config.STATUS_COLUMNS)) {
      console.log('❌ Not a status column');
      return;
    }
    console.log('✓ Valid status column');
    
    const row = range.getRow();
    console.log('Row:', row);
    console.log('Allowed row ranges:', config.ROW_RANGES);
    
    // Check if row is within the processing ranges
    if (!isInRanges(row, config.ROW_RANGES)) {
      console.log('❌ Row not in processing range');
      return;
    }
    console.log('✓ Row in processing range');
    
    const newValue = range.getValue();
    console.log('New value in status cell:', newValue);
    
    // Get cached value
    const cachedValue = getCachedStatusValue(sheet, row, statusColumn);
    console.log('Cached value:', cachedValue);
    
    // Check if this is a deletion (empty value) or change from a non-default value
    if (cachedValue && (newValue === '' || newValue === config.DEFAULT_STATUS_TEXT || newValue !== cachedValue)) {
      // There's a cached value and user is trying to delete or change it
      console.log('Attempting to change/delete cached value');
      
      // Check if the user is authorized to make this change
      const isAuthorized = isUserAuthorizedForStatus(sheet, row, statusColumn, config);
      console.log('User authorized:', isAuthorized);
      
      if (isAuthorized) {
        // User is authorized - update cache with new value
        if (newValue && newValue !== config.DEFAULT_STATUS_TEXT) {
          cacheStatusValue(sheet, row, statusColumn, newValue);
        } else {
          cacheStatusValue(sheet, row, statusColumn, null);
        }
        console.log('Authorized status change by user');
      } else {
        // User is not authorized - revert to cached value
        console.log('Unauthorized - reverting to cached value');
        Utilities.sleep(100); // Small delay to ensure sheet is ready
        
        range.setValue(cachedValue);
        range.setFontColor(config.COLOR_SUCCESS);
        SpreadsheetApp.flush();
        
        // Show minimalistic toast notification
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Нет прав на изменение этой подписи',
          '❌ Доступ запрещен',
          3
        );
      }
    } else if (!cachedValue && newValue && newValue !== config.DEFAULT_STATUS_TEXT) {
      // No cached value but user is trying to set a non-default value
      // Check if they're authorized to set a new signature
      console.log('Attempting to set new signature');
      const isAuthorized = isUserAuthorizedForStatus(sheet, row, statusColumn, config);
      
      if (isAuthorized) {
        // Authorized - cache the new value
        cacheStatusValue(sheet, row, statusColumn, newValue);
        console.log('Authorized new signature set');
      } else {
        // Not authorized - revert to default
        console.log('Unauthorized new signature attempt');
        Utilities.sleep(100);
        range.setValue(config.DEFAULT_STATUS_TEXT);
        range.setFontColor(config.DEFAULT_FONT_COLOR);
        SpreadsheetApp.flush();
        
        // Show minimalistic toast notification
        SpreadsheetApp.getActiveSpreadsheet().toast(
          'Нет прав на установку подписи',
          '❌ Доступ запрещен',
          3
        );
      }
    } else if (!cachedValue && (newValue === '' || newValue === config.DEFAULT_STATUS_TEXT)) {
      // No cached value and setting to default/empty - allow it
      console.log('Setting to default/empty - allowed');
    }
    
    console.log('=== onEditStatusProtection COMPLETED ===');
    
  } catch (error) {
    console.error('❌ ERROR in onEditStatusProtection:', error);
    console.error('Stack trace:', error.stack);
  }
}

/**
 * Handles checkbox unchecked - clears the signature if authorized
 */
function handleCheckboxUnchecked(sheet, row, checkboxColumn, config) {
  try {
    console.log('=== handleCheckboxUnchecked STARTED ===');
    console.log('Row:', row, 'Checkbox column:', checkboxColumn);
    
    const statusColumn = getStatusColumnForCheckbox(checkboxColumn, config);
    const statusCell = sheet.getRange(row, statusColumn);
    const checkboxCell = sheet.getRange(row, checkboxColumn);
    
    console.log('Status column for checkbox:', statusColumn);
    
    // Check if user is authorized to clear this signature
    const isAuthorized = isUserAuthorizedForStatus(sheet, row, statusColumn, config);
    console.log('User authorized to clear:', isAuthorized);
    
    if (!isAuthorized) {
      // User is not authorized - revert checkbox back to checked
      console.log('User not authorized - reverting checkbox');
      Utilities.sleep(100);
      checkboxCell.setValue(true);
      
      // Show minimalistic toast notification
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Нет прав на удаление подписи',
        '❌ Доступ запрещен',
        3
      );
      
      console.log('Unauthorized checkbox uncheck prevented');
      return;
    }
    
    // User is authorized - clear the signature
    console.log('User authorized - clearing signature');
    setStatus(statusCell, config.DEFAULT_STATUS_TEXT, config.DEFAULT_FONT_COLOR);
    
    // Clear from cache
    cacheStatusValue(sheet, row, statusColumn, null);
    
    console.log('Signature cleared by authorized user');
    console.log('=== handleCheckboxUnchecked COMPLETED ===');
    
  } catch (error) {
    console.error('Error handling checkbox unchecked:', error);
    console.error('Stack trace:', error.stack);
    // Revert checkbox on error
    Utilities.sleep(100);
    checkboxCell.setValue(true);
  }
}

/**
 * Verifies the active user and places their signature
 */
function verifyAndSign(sheet, row, checkboxColumn, config) {
  console.log('=== verifyAndSign STARTED ===');
  console.log('Row:', row, 'Checkbox column:', checkboxColumn);
  
  const statusColumn = getStatusColumnForCheckbox(checkboxColumn, config);
  const statusCell = sheet.getRange(row, statusColumn);
  const checkboxCell = sheet.getRange(row, checkboxColumn);
  
  console.log('Status column for checkbox:', statusColumn);
  
  try {
    // Check if there's already a cached value (already signed)
    const cachedValue = getCachedStatusValue(sheet, row, statusColumn);
    console.log('Existing cached value:', cachedValue);
    
    if (cachedValue) {
      // Already signed - check if user is authorized to re-sign
      console.log('Already signed - checking authorization for re-sign');
      const isAuthorized = isUserAuthorizedForStatus(sheet, row, statusColumn, config);
      console.log('User authorized to re-sign:', isAuthorized);
      
      if (!isAuthorized) {
        handleError(checkboxCell, statusCell, '❌ Нет прав на переподпись', config);
        return;
      }
    }
    
    // ═══════════════════════════════════════════════════════
    // STEP 1: Initialize
    // ═══════════════════════════════════════════════════════
    console.log('Step 1: Initializing...');
    setStatus(statusCell, '🔄 Проверка...', config.COLOR_PROCESSING);
    
    // ═══════════════════════════════════════════════════════
    // STEP 2: Get active user email
    // ═══════════════════════════════════════════════════════
    console.log('Step 2: Getting active user email...');
    const activeUserEmail = Session.getActiveUser().getEmail().toLowerCase();
    console.log('Active user email:', activeUserEmail);
    
    if (!activeUserEmail) {
      console.log('❌ No active user email found');
      handleError(checkboxCell, statusCell, '❌ Ошибка авторизации', config);
      return;
    }
    
    setStatus(statusCell, '🔍 Поиск подписи...', config.COLOR_PROCESSING);
    
    // ═══════════════════════════════════════════════════════
    // STEP 3: Search in reference data using named range
    // ═══════════════════════════════════════════════════════
    console.log('Step 3: Searching in reference data...');
    const dataRange = sheet.getParent().getRangeByName(config.DATA_RANGE_NAME);
    
    if (!dataRange) {
      console.log('❌ Data range not found:', config.DATA_RANGE_NAME);
      handleError(checkboxCell, statusCell, '❌ Набор данных не найден', config);
      return;
    }
    
    console.log('Data range found:', dataRange.getA1Notation());
    const data = dataRange.getValues();
    
    let userSignature = null;
    let userName = null;
    
    // Assuming named range has columns: ID, Name, Email, Signature
    for (let i = 0; i < data.length; i++) {
      const dataEmail = data[i][2]; // Email is in 3rd column (index 2)
      
      if (dataEmail && dataEmail.toString().toLowerCase() === activeUserEmail) {
        userSignature = data[i][3]; // Signature is in 4th column (index 3)
        userName = data[i][1];      // Name is in 2nd column (index 1)
        console.log('Found user in data range:', { userName, userSignature });
        break;
      }
    }
    
    if (!userSignature) {
      console.log('❌ User signature not found for email:', activeUserEmail);
      handleError(checkboxCell, statusCell, '❌ Email не найден', config);
      return;
    }
    
    setStatus(statusCell, '🔐 Проверка прав...', config.COLOR_PROCESSING);
    
    // ═══════════════════════════════════════════════════════
    // STEP 4: Verify signer matches expected person
    // ═══════════════════════════════════════════════════════
    console.log('Step 4: Verifying name match...');
    const nameColumn = getNameColumnForStatus(statusColumn, config);
    console.log('Name column for status:', nameColumn);
    
    const expectedName = sheet.getRange(row, nameColumn).getValue();
    console.log('Expected name from sheet:', expectedName);
    console.log('User name from data:', userName);
    
    if (!expectedName || expectedName.toString().trim() === '') {
      console.log('❌ No expected name in name column');
      handleError(checkboxCell, statusCell, '❌ Имя не указано', config);
      return;
    }
    
    if (!nameMatches(expectedName, userName)) {
      console.log('❌ Name mismatch');
      handleError(checkboxCell, statusCell, '❌ Нет прав на подпись', config);
      return;
    }
    
    console.log('✓ Name match verified successfully');
    
    // ═══════════════════════════════════════════════════════
    // SUCCESS!
    // ═══════════════════════════════════════════════════════
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
    const successMessage = userSignature + ' ' + timestamp;
    console.log('Success message:', successMessage);
    
    setStatus(statusCell, successMessage, config.COLOR_SUCCESS);
    
    // Cache the successful signature
    cacheStatusValue(sheet, row, statusColumn, successMessage);
    
    console.log('✓ Signature successful and cached');
    console.log('=== verifyAndSign COMPLETED ===');
    
  } catch (error) {
    console.error('❌ ERROR in verifyAndSign:', error);
    console.error('Stack trace:', error.stack);
    handleError(checkboxCell, statusCell, '❌ Ошибка', config);
  }
}

/**
 * Sets status cell with color
 */
function setStatus(cell, text, color) {
  console.log('Setting status:', { text, color });
  cell.setValue(text).setFontColor(color);
  SpreadsheetApp.flush();
}

/**
 * Handles errors: resets checkbox, shows error, resets to default text after 5 sec
 */
function handleError(checkboxCell, statusCell, errorText, config) {
  console.log('handleError called:', errorText);
  
  checkboxCell.setValue(false);
  setStatus(statusCell, errorText, config.COLOR_ERROR);
  
  console.log('Error shown, will reset in 5 seconds...');
  
  // Wait 5 seconds then reset to default text
  Utilities.sleep(5000);
  statusCell.setValue(config.DEFAULT_STATUS_TEXT)
            .setFontColor(config.DEFAULT_FONT_COLOR);
  SpreadsheetApp.flush();
  
  console.log('Error reset to default text');
}

/**
 * Checks if two names match
 */
function nameMatches(shortName, fullName) {
  if (!shortName || !fullName) return false;
  
  const short = shortName.toString().trim().toLowerCase();
  const full = fullName.toString().trim().toLowerCase();
  
  // Exact match
  if (short === full) {
    console.log('Exact name match');
    return true;
  }
  
  // Check surname match (first word)
  const shortSurname = short.split(/[\s\.]+/)[0];
  const fullSurname = full.split(/[\s\.]+/)[0];
  
  console.log('Comparing surnames:', shortSurname, 'vs', fullSurname);
  return shortSurname === fullSurname;
}

/**
 * Test function to verify configuration
 */
function testConfiguration() {
  try {
    console.log('=== testConfiguration STARTED ===');
    
    const config = getConfig();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(config.SHEET_NAME);
    
    if (!sheet) {
      throw new Error(`Sheet "${config.SHEET_NAME}" not found!`);
    }
    
    console.log('✓ Sheet found:', config.SHEET_NAME);
    
    const dataRange = ss.getRangeByName(config.DATA_RANGE_NAME);
    if (!dataRange) {
      throw new Error(`Named range "${config.DATA_RANGE_NAME}" not found!`);
    }
    
    console.log('✓ Data range found:', config.DATA_RANGE_NAME);
    console.log('Data range size:', dataRange.getNumRows(), 'rows x', dataRange.getNumColumns(), 'columns');
    
    // Parse the ranges for display
    const checkboxCols = parseRangeString(config.CHECKBOX_COLUMNS);
    const statusCols = parseRangeString(config.STATUS_COLUMNS);
    const nameCols = parseRangeString(config.NAME_COLUMNS);
    const rowRanges = parseRangeString(config.ROW_RANGES);
    
    console.log('Configuration test passed:');
    console.log(`- Sheet: ${config.SHEET_NAME} ✓`);
    console.log(`- Named range: ${config.DATA_RANGE_NAME} ✓`);
    console.log(`- Data range size: ${dataRange.getNumRows()} rows x ${dataRange.getNumColumns()} columns`);
    console.log(`- Checkbox columns: ${checkboxCols.map(c => String.fromCharCode(64 + c)).join(', ')} (${checkboxCols.join(', ')})`);
    console.log(`- Status columns: ${statusCols.map(c => String.fromCharCode(64 + c)).join(', ')} (${statusCols.join(', ')})`);
    console.log(`- Name columns: ${nameCols.map(c => String.fromCharCode(64 + c)).join(', ')} (${nameCols.join(', ')})`);
    console.log(`- Processing rows: ${rowRanges.join(', ')}`);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Конфигурация проверена успешно!', '✓ Тест', 3);
    
    console.log('=== testConfiguration COMPLETED ===');
    
    return {
      success: true,
      message: 'Configuration test passed!',
      details: {
        sheet: config.SHEET_NAME,
        dataRange: config.DATA_RANGE_NAME,
        checkboxColumns: `${checkboxCols.map(c => String.fromCharCode(64 + c)).join(', ')} (${config.CHECKBOX_COLUMNS})`,
        statusColumns: `${statusCols.map(c => String.fromCharCode(64 + c)).join(', ')} (${config.STATUS_COLUMNS})`,
        nameColumns: `${nameCols.map(c => String.fromCharCode(64 + c)).join(', ')} (${config.NAME_COLUMNS})`,
        rowRanges: rowRanges.join(', ') + ` (${config.ROW_RANGES})`
      }
    };
    
  } catch (error) {
    console.error('❌ ERROR in testConfiguration:', error);
    console.error('Stack trace:', error.stack);
    throw error;
  }
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
    .addItem('🔧 Установить триггеры', 'installTrigger')
    .addSeparator()
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