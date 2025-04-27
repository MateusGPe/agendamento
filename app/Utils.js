/**
 * Arquivo: Utils.gs
 * Descrição: Contém funções utilitárias genéricas e de formatação/parsing.
 */
function createJsonResponse(success, message, data = null) {
  if (!success) {
    Logger.log(`Operation Failed: ${message}`);
    if (data && data.error) {
      Logger.log(`Error Details: ${data.error.message || data.error}`);
      if (data.error.stack) Logger.log(`Stack: ${data.error.stack}`);
    } else if (data && data.bookingId) {
      Logger.log(`Associated Booking ID (if any): ${data.bookingId}`);
    }
  }
  return JSON.stringify({ success, message, data });
}
function getActiveUserEmail_() {
  try {
    const email = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail();
    if (!email) throw new Error("Session.getEffectiveUser().getEmail() returned empty.");
    return email;
  } catch (e) {
    Logger.log('CRITICAL: Failed to get active/effective user email: ' + e.message);
    throw new Error('Não foi possível identificar o usuário logado.');
  }
}
function acquireScriptLock_(timeoutMilliseconds = SCRIPT_LOCK_TIMEOUT_MS) {
  const lock = LockService.getScriptLock();
  try {
    Logger.log(`Attempting to acquire script lock (timeout: ${timeoutMilliseconds}ms)...`);
    lock.waitLock(timeoutMilliseconds);
    Logger.log("Script lock acquired.");
    return lock;
  } catch (e) {
    Logger.log(`Failed to acquire script lock within ${timeoutMilliseconds}ms: ${e.message}`);
    throw new Error('O sistema está ocupado processando outra solicitação. Tente novamente em alguns instantes.');
  }
}
function releaseScriptLock_(lock) {
  if (lock && typeof lock.releaseLock === 'function') {
    try {
      lock.releaseLock();
      Logger.log("Script lock released.");
    } catch (e) {
      Logger.log(`Warning: Error releasing script lock (may have expired or already released): ${e.message}`);
    }
  } else {
  }
}
function formatValueToDate(rawValue) {
  if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
    if (rawValue.getFullYear() === 1899 && rawValue.getMonth() === 11 && rawValue.getDate() === 30) {
      return null;
    }
    return new Date(Date.UTC(rawValue.getFullYear(), rawValue.getMonth(), rawValue.getDate()));
  }
  return null;
}
function parseDDMMYYYY(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
  }
  if (typeof value === 'string') {
    const dateString = value.trim();
    const parts = dateString.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
    if (!parts) return null;
    const day = parseInt(parts[1], 10);
    const month = parseInt(parts[2], 10) - 1;
    const year = parseInt(parts[3], 10);
    if (month >= 0 && month <= 11 && day >= 1 && day <= 31 && year >= 1000) {
      const date = new Date(Date.UTC(year, month, day));
      if (date.getUTCFullYear() === year && date.getUTCMonth() === month && date.getUTCDate() === day) {
        return date;
      }
    }
    return null;
  }
  return null;
}
function formatValueToHHMM(rawValue, timeZone) {
  try {
    if (!timeZone) {
      Logger.log("Warning: Timezone not provided to formatValueToHHMM. Using default.");
      timeZone = Session.getScriptTimeZone();
    }
    if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
      if (rawValue.getFullYear() === 1899 && rawValue.getMonth() === 11 && rawValue.getDate() === 30) {
        if (rawValue.getHours() !== 0 || rawValue.getMinutes() !== 0 || rawValue.getSeconds() !== 0 || rawValue.getMilliseconds() !== 0) {
          return Utilities.formatDate(rawValue, timeZone, 'HH:mm');
        } else {
          return null;
        }
      }
      return Utilities.formatDate(rawValue, timeZone, 'HH:mm');
    }
    if (typeof rawValue === 'string') {
      const timeMatch = rawValue.trim().match(/^(\d{1,2}):(\d{2})(:\d{2})?(\s*(?:AM|PM))?$/i);
      if (timeMatch) {
        let hour = parseInt(timeMatch[1], 10);
        const minute = parseInt(timeMatch[2], 10);
        const ampm = (timeMatch[4] || '').trim().toUpperCase();
        if (ampm === 'PM' && hour < 12) hour += 12;
        if (ampm === 'AM' && hour === 12) hour = 0;
        if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
          return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
        }
      }
    }
    if (typeof rawValue === 'number' && rawValue >= 0 && rawValue <= 1) {
      const totalMinutes = Math.round(rawValue * 1440);
      if (totalMinutes === 1440) return "00:00";
      const hours = Math.floor(totalMinutes / 60) % 24;
      const minutes = totalMinutes % 60;
      if (hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59) {
        return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
      }
    }
    return null;
  } catch (e) {
    Logger.log(`Error in formatValueToHHMM for value "${rawValue}" (Type: ${typeof rawValue}): ${e.message}`);
    return null;
  }
}
function invalidateSheetCache_(sheetName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = `SHEET_DATA_${SPREADSHEET_ID}_${sheetName}`;
  cache.remove(cacheKey);
  Logger.log(`Cache invalidated for sheet "${sheetName}".`);
}
function appendSheetRows_(sheet, numCols, rowsToAppend) {
  if (!rowsToAppend || rowsToAppend.length === 0) {
    return;
  }
  try {
    invalidateSheetCache_(sheet.getName());
    const finalRows = rowsToAppend.map(row => {
      const finalRow = [...row];
      while (finalRow.length < numCols) finalRow.push('');
      if (finalRow.length > numCols) finalRow.length = numCols;
      return finalRow;
    });
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, finalRows.length, numCols).setValues(finalRows);
    Logger.log(`${finalRows.length} rows appended successfully to sheet "${sheet.getName()}".`);
  } catch (e) {
    Logger.log(`ERROR appending rows to sheet "${sheet.getName()}": ${e.message}`);
    throw new Error(`Erro interno ao adicionar novas linhas na planilha "${sheet.getName()}": ${e.message}`);
  }
}
function updateSheetRow_(sheet, rowIndex, numCols, updatedRowData) {
  try {
    invalidateSheetCache_(sheet.getName());
    const finalRow = [...updatedRowData];
    while (finalRow.length < numCols) finalRow.push('');
    if (finalRow.length > numCols) finalRow.length = numCols;
    sheet.getRange(rowIndex, 1, 1, numCols).setValues([finalRow]);
    Logger.log(`Row ${rowIndex} updated successfully in sheet "${sheet.getName()}".`);
  } catch (e) {
    Logger.log(`ERROR updating row ${rowIndex} in sheet "${sheet.getName()}": ${e.message}`);
    throw new Error(`Erro interno ao atualizar a linha ${rowIndex} na planilha "${sheet.getName()}": ${e.message}`);
  }
}
function updateCalendarGuests_(event, newGuestEmails) {
  if (!event || typeof event.getGuestList !== 'function') return;
  const newGuestsLower = newGuestEmails.map(g => String(g || '').toLowerCase()).filter(g => g && g.includes('@'));
  const existingGuests = event.getGuestList();
  const existingGuestsLower = existingGuests.map(g => g.getEmail().toLowerCase());
  existingGuests.forEach(guest => {
    const emailLower = guest.getEmail().toLowerCase();
    if (!newGuestsLower.includes(emailLower)) {
      try {
        event.removeGuest(guest.getEmail());
        Logger.log(`Removed guest ${guest.getEmail()} from event ${event.getId()}`);
      } catch (removeErr) {
        Logger.log(`Failed to remove guest ${guest.getEmail()}: ${removeErr.message}`);
      }
    }
  });
  newGuestsLower.forEach(guestEmail => {
    if (!existingGuestsLower.includes(guestEmail)) {
      try {
        event.addGuest(guestEmail);
        Logger.log(`Added guest ${guestEmail} to event ${event.getId()}`);
      } catch (addErr) {
        Logger.log(`Failed to add guest ${guestEmail}: ${addErr.message}`);
      }
    }
  });
}