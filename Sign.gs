/**
 * Signature Verification Script - Using Installable Trigger
 * Logic: Find signature by active user's email in dt_sign_allData
 * Clean status feedback without popups
 */

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
  
  if (sheet.getName() !== 'dt_signs') return;
  if (range.getColumn() !== 4) return;
  if (range.getValue() !== true) return;
  
  const row = range.getRow();
  if (row < 20) return;
  
  verifyAndSign(sheet, row);
}

/**
 * Verifies the active user and places their signature
 */
function verifyAndSign(sheet, row) {
  const statusCell = sheet.getRange(row, 3); // Column C for status
  const checkboxCell = sheet.getRange(row, 4);
  
  try {
    // ═══════════════════════════════════════════════════════
    // STEP 1: Initialize
    // ═══════════════════════════════════════════════════════
    setStatus(statusCell, '🔄 Проверка...', '#FFA500');
    
    // ═══════════════════════════════════════════════════════
    // STEP 2: Get active user email
    // ═══════════════════════════════════════════════════════
    const activeUserEmail = Session.getActiveUser().getEmail().toLowerCase();
    
    if (!activeUserEmail) {
      handleError(checkboxCell, statusCell, '❌ Ошибка авторизации');
      return;
    }
    
    setStatus(statusCell, '🔍 Поиск подписи...', '#FFA500');
    
    // ═══════════════════════════════════════════════════════
    // STEP 3: Search in reference data
    // ═══════════════════════════════════════════════════════
    const dataRange = sheet.getRange('A2:D12');
    const data = dataRange.getValues();
    
    let userSignature = null;
    let userName = null;
    
    for (let i = 0; i < data.length; i++) {
      const dataEmail = data[i][2];
      
      if (dataEmail && dataEmail.toString().toLowerCase() === activeUserEmail) {
        userSignature = data[i][3];
        userName = data[i][1];
        break;
      }
    }
    
    if (!userSignature) {
      handleError(checkboxCell, statusCell, '❌ Email не найден');
      return;
    }
    
    setStatus(statusCell, '🔐 Проверка прав...', '#FFA500');
    
    // ═══════════════════════════════════════════════════════
    // STEP 4: Verify signer matches expected person
    // ═══════════════════════════════════════════════════════
    const expectedName = sheet.getRange(row, 2).getValue();
    
    if (!nameMatches(expectedName, userName)) {
      handleError(checkboxCell, statusCell, '❌ Нет прав на подпись');
      return;
    }
    
    // ═══════════════════════════════════════════════════════
    // SUCCESS!
    // ═══════════════════════════════════════════════════════
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm');
    setStatus(statusCell, '✔ Подписано: ' + timestamp, '#28a745');
    
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
 * Handles errors: resets checkbox, shows error, resets to "Подпись" after 5 sec
 */
function handleError(checkboxCell, statusCell, errorText) {
  checkboxCell.setValue(false);
  setStatus(statusCell, errorText, '#dc3545');
  
  // Wait 5 seconds then reset to default text
  Utilities.sleep(5000);
  statusCell.setValue('Подпись').setFontColor('#000000'); // Default black color
  SpreadsheetApp.flush();
}

/**
 * Checks if two names match
 */
function nameMatches(shortName, fullName) {
  if (!shortName || !fullName) return false;
  const short = shortName.toString().trim().toLowerCase();
  const full = fullName.toString().trim().toLowerCase();
  if (short === full) return true;
  const shortSurname = short.split(/[\\s\\.]+/)[0];
  const fullSurname = full.split(/[\\s\\.]+/)[0];
  return shortSurname === fullSurname;
}