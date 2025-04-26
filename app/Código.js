/*
    MIT License

    Copyright (c) 2025 Mateus G. Pereira

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to
    deal in the Software without restriction, including without limitation the
    rights to use, copy, modify, merge, publish, distribute, sublicense, and/or
    sell copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in
    all copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
    FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
    IN THE SOFTWARE.
*/

// Code.gs
// Obs: Comentários gerados por IA.
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SCRIPT_LOCK_TIMEOUT_MS = 15000; // Timeout for script lock (15 seconds)

// Sheet Names
const SHEETS = Object.freeze({
  CONFIG: 'Configuracoes',
  AUTHORIZED_USERS: 'Usuarios Autorizados',
  BASE_SCHEDULES: 'Horarios Base',
  SCHEDULE_INSTANCES: 'Instancias de Horarios',
  BOOKING_DETAILS: 'Reservas Detalhadas',
  DISCIPLINES: 'Disciplinas'
});

// Column Indices (0-based) - Grouped by Sheet
const HEADERS = Object.freeze({
  CONFIG: Object.freeze({ NOME: 0, VALOR: 1 }),
  AUTHORIZED_USERS: Object.freeze({ EMAIL: 0, NOME: 1, PAPEL: 2 }),
  BASE_SCHEDULES: Object.freeze({
    ID: 0, TIPO: 1, DIA_SEMANA: 2, HORA_INICIO: 3, DURACAO: 4,
    PROFESSOR_PRINCIPAL: 5, TURMA_PADRAO: 6, DISCIPLINA_PADRAO: 7,
    CAPACIDADE: 8, OBSERVATIONS: 9
  }),
  SCHEDULE_INSTANCES: Object.freeze({
    ID_INSTANCIA: 0, ID_BASE_HORARIO: 1, TURMA: 2, PROFESSOR_PRINCIPAL: 3,
    DATA: 4, DIA_SEMANA: 5, HORA_INICIO: 6, TIPO_ORIGINAL: 7,
    STATUS_OCUPACAO: 8, ID_RESERVA: 9, ID_EVENTO_CALENDAR: 10
  }),
  BOOKING_DETAILS: Object.freeze({
    ID_RESERVA: 0, TIPO_RESERVA: 1, ID_INSTANCIA: 2, PROFESSOR_REAL: 3,
    PROFESSOR_ORIGINAL: 4, ALUNOS: 5, TURMAS_AGENDADA: 6, DISCIPLINA_REAL: 7,
    DATA_HORA_INICIO_EFETIVA: 8, STATUS_RESERVA: 9, DATA_CRIACAO: 10,
    CRIADO_POR: 11
  }),
  DISCIPLINES: Object.freeze({ NOME: 0 })
});

// Statuses, Types, Roles
const STATUS_OCUPACAO = Object.freeze({
  DISPONIVEL: 'Disponivel',
  REPOSICAO_AGENDADA: 'Reposicao Agendada',
  SUBSTITUICAO_AGENDADA: 'Substituicao Agendada'
});
const TIPOS_RESERVA = Object.freeze({
  REPOSICAO: 'Reposicao',
  SUBSTITUICAO: 'Substituicao'
});
const TIPOS_HORARIO = Object.freeze({
  FIXO: 'Fixo',
  VAGO: 'Vago'
});
const USER_ROLES = Object.freeze({
  ADMIN: 'Admin',
  PROFESSOR: 'Professor',
  ALUNO: 'Aluno'
});

// Email Configuration
const ADMIN_COPY_EMAILS = ["cae.itq@ifsp.edu.br", "mtm.itq@ifsp.edu.br"]; // Fixed BCC list
const EMAIL_SENDER_NAME = 'Sistema de Reservas IFSP';

// ==========================================================================
//                            Utility & Helper Functions
// ==========================================================================

/**
 * Standardized JSON response creation. Logs failures server-side.
 * @param {boolean} success - Whether the operation succeeded.
 * @param {string} message - A descriptive message.
 * @param {*} [data=null] - Optional data payload.
 * @returns {string} JSON string representation.
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

/**
 * Gets the active/effective user's email safely.
 * @returns {string} The user's email.
 * @throws {Error} If the user email cannot be retrieved.
 */
function getActiveUserEmail_() {
  try {
    // Prioritize getEffectiveUser for broader compatibility (triggers, etc.)
    // Fallback to getActiveUser if effective user isn't available (less common for web apps)
    const email = Session.getEffectiveUser().getEmail() || Session.getActiveUser().getEmail();
    if (!email) throw new Error("Session.getEffectiveUser().getEmail() returned empty.");
    return email;
  } catch (e) {
    Logger.log('CRITICAL: Failed to get active/effective user email: ' + e.message);
    throw new Error('Não foi possível identificar o usuário logado.');
  }
}

/**
 * Gets a sheet by name and handles not found errors.
 * @param {string} sheetName - The name of the sheet.
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The sheet object.
 * @throws {Error} If the sheet is not found.
 */
function getSheetByName_(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`CRITICAL: Sheet "${sheetName}" not found.`);
    throw new Error(`Erro interno: Planilha "${sheetName}" não encontrada.`);
  }
  return sheet;
}

/**
 * Reads data from a sheet, handling empty sheets and providing header/data separately.
 * Optionally checks for a minimum number of columns based on header definitions.
 * @param {string} sheetName - The name of the sheet.
 * @param {object} [headersDefinition=null] - Optional: The HEADERS[SHEET_KEY] object for column count check.
 * @returns {{header: string[], data: any[][], sheet: GoogleAppsScript.Spreadsheet.Sheet}}
 * @throws {Error} If the sheet is not found.
 */
function getSheetData_(sheetName, headersDefinition = null) {
  const sheet = getSheetByName_(sheetName);
  const range = sheet.getDataRange();
  const values = range.getValues(); // Use getValues to preserve Date objects etc.

  if (!values || values.length === 0) {
    Logger.log(`Sheet "${sheetName}" is completely empty.`);
    return { header: [], data: [], sheet: sheet };
  }

  const header = values[0];
  const data = values.length > 1 ? values.slice(1) : [];

  // Optional minimum column check
  if (headersDefinition) {
    const expectedMinCols = Math.max(...Object.values(headersDefinition)) + 1;
    if (header.length < expectedMinCols) {
      Logger.log(`WARNING: Sheet "${sheetName}" header has ${header.length} columns, but expected at least ${expectedMinCols} based on HEADERS definition. Data processing might fail.`);
    }
  }

  if (data.length === 0) {
    Logger.log(`Sheet "${sheetName}" contains only a header row or is empty.`);
  }

  return { header, data, sheet };
}

/**
 * Finds the 1-based row index of a row matching a specific ID in a 2D data array.
 * Assumes the first row of the data array corresponds to row 2 in the sheet.
 * @param {any[][]} data - The 2D array of data (excluding header).
 * @param {number} idColumnIndex - The 0-based index of the column containing the ID.
 * @param {string} targetId - The ID to search for (will be trimmed).
 * @returns {number} The 1-based row index in the *sheet*, or -1 if not found or invalid input.
 */
function findRowIndexById_(data, idColumnIndex, targetId) {
  if (!targetId || typeof targetId !== 'string' || idColumnIndex < 0) return -1;
  const trimmedTargetId = targetId.trim();
  if (trimmedTargetId === '') return -1;

  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    // Check if row exists and has the ID column
    if (row && row.length > idColumnIndex) {
      const currentIdRaw = row[idColumnIndex];
      // Check if ID is string or number before converting/trimming
      if (typeof currentIdRaw === 'string' || typeof currentIdRaw === 'number') {
        const currentId = String(currentIdRaw).trim();
        if (currentId === trimmedTargetId) {
          return i + 2; // Return 1-based sheet row index
        }
      }
    }
  }
  return -1; // Not found
}

/**
 * Acquires a script lock with a timeout, throwing an error if unsuccessful.
 * @param {number} timeoutMilliseconds - Timeout duration.
 * @returns {GoogleAppsScript.Lock.Lock} The acquired lock.
 * @throws {Error} If the lock cannot be acquired within the timeout.
 */
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

/**
 * Releases a script lock safely, logging warnings if release fails.
 * @param {GoogleAppsScript.Lock.Lock | null} lock - The lock object to release.
 */
function releaseScriptLock_(lock) {
  if (lock && typeof lock.releaseLock === 'function') {
    try {
      lock.releaseLock();
      Logger.log("Script lock released.");
    } catch (e) {
      Logger.log(`Warning: Error releasing script lock (may have expired or already released): ${e.message}`);
    }
  } else {
    // Logger.log("No valid lock object provided to releaseScriptLock_."); // Can be noisy
  }
}

/**
 * Safely updates a specific row in a sheet. Pads/trims data to match column count.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object.
 * @param {number} rowIndex - The 1-based row index to update.
 * @param {number} numCols - The total number of columns the sheet range should cover.
 * @param {any[]} updatedRowData - The 1D array of data for the row.
 * @throws {Error} If the update fails.
 */
function updateSheetRow_(sheet, rowIndex, numCols, updatedRowData) {
  try {
    if (!sheet || typeof sheet.getRange !== 'function') throw new Error("Invalid sheet object provided for update.");
    if (rowIndex < 1) throw new Error(`Invalid row index: ${rowIndex}.`);
    if (numCols < 1) throw new Error(`Invalid number of columns: ${numCols}.`);
    if (!Array.isArray(updatedRowData)) throw new Error("updatedRowData must be an array.");

    const finalRowData = [...updatedRowData];
    while (finalRowData.length < numCols) finalRowData.push('');
    if (finalRowData.length > numCols) finalRowData.length = numCols;

    sheet.getRange(rowIndex, 1, 1, numCols).setValues([finalRowData]);
    Logger.log(`Row ${rowIndex} in sheet "${sheet.getName()}" updated successfully.`);
  } catch (e) {
    Logger.log(`ERROR updating row ${rowIndex} in sheet "${sheet.getName()}": ${e.message}`);
    throw new Error(`Erro interno ao atualizar dados na planilha "${sheet.getName()}" (linha ${rowIndex}): ${e.message}`);
  }
}

/**
 * Appends a new row safely to a sheet. Pads/trims data to match column count.
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object.
 * @param {number} numCols - The expected number of columns.
 * @param {any[]} newRowData - The 1D array of data for the new row.
 * @throws {Error} If the append operation fails.
 */
function appendSheetRow_(sheet, numCols, newRowData) {
  try {
    if (!sheet || typeof sheet.appendRow !== 'function') throw new Error("Invalid sheet object provided for append.");
    if (numCols < 1) throw new Error(`Invalid number of columns: ${numCols}.`);
    if (!Array.isArray(newRowData)) throw new Error("newRowData must be an array.");

    const finalRowData = [...newRowData];
    while (finalRowData.length < numCols) finalRowData.push('');
    if (finalRowData.length > numCols) finalRowData.length = numCols;

    sheet.appendRow(finalRowData);
    // Logger.log(`Row appended successfully to sheet "${sheet.getName()}".`); // Can be verbose
  } catch (e) {
    Logger.log(`ERROR appending row to sheet "${sheet.getName()}": ${e.message}`);
    throw new Error(`Erro interno ao adicionar dados na planilha "${sheet.getName()}": ${e.message}`);
  }
}

// ==========================================================================
//                            Formatting Helpers
// ==========================================================================

/**
 * Tries to convert a value to a valid Date object normalized to UTC midnight.
 * Handles Sheets' "zero date" (1899-12-30) by returning null.
 * @param {*} rawValue - The value from the sheet cell.
 * @returns {Date|null} A valid Date object (UTC midnight) or null.
 */
function formatValueToDate(rawValue) {
  if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
    // Explicitly check for and reject the 1899 date artifact
    if (rawValue.getFullYear() === 1899 && rawValue.getMonth() === 11 && rawValue.getDate() === 30) {
      // Logger.log(`formatValueToDate: Detected 1899-12-30 artifact, returning null.`); // Optional log
      return null;
    }
    // Return a *new* Date object representing UTC midnight for consistent date comparisons
    return new Date(Date.UTC(rawValue.getFullYear(), rawValue.getMonth(), rawValue.getDate()));
  }
  return null;
}

/**
 * Parses a string in 'dd/MM/yyyy' format or a Date object to a Date object normalized to UTC midnight.
 * @param {*} value - String 'dd/MM/yyyy', Date object, or other.
 * @returns {Date|null} The Date object (UTC midnight) or null if invalid.
 */
function parseDDMMYYYY(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    // Normalize existing Date objects to UTC midnight
    return new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate()));
  }

  if (typeof value === 'string') {
    const dateString = value.trim();
    // Allow ., / or - as separators
    const parts = dateString.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
    if (!parts) return null;

    const day = parseInt(parts[1], 10);
    const month = parseInt(parts[2], 10) - 1; // 0-indexed for Date constructor
    const year = parseInt(parts[3], 10);

    // Check basic validity and create UTC date
    if (month >= 0 && month <= 11 && day >= 1 && day <= 31 && year >= 1000) {
      const date = new Date(Date.UTC(year, month, day));
      // Final check: Ensure created date components match input (handles Feb 30 etc.)
      if (date.getUTCFullYear() === year && date.getUTCMonth() === month && date.getUTCDate() === day) {
        return date; // Already normalized to UTC midnight
      }
    }
    return null; // Invalid date components or failed validation
  }
  return null; // Not a Date or parseable string
}


/**
 * Formats a value (Date, string, or Sheets time number) to "HH:mm" string using the sheet's timezone.
 * Handles Sheets' "zero date" artifact if it contains time components.
 * @param {*} rawValue - The value from the sheet.
 * @param {string} timeZone - The spreadsheet's time zone ID (e.g., "America/Sao_Paulo").
 * @returns {string|null} Formatted time string "HH:mm" or null.
 */
function formatValueToHHMM(rawValue, timeZone) {
  try {
    if (!timeZone) { // Safety check for timezone
      Logger.log("Warning: Timezone not provided to formatValueToHHMM. Using default.");
      timeZone = Session.getScriptTimeZone(); // Fallback to script timezone
    }

    if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
      // Handle the 1899 "zero date" case specifically if it has time components
      if (rawValue.getFullYear() === 1899 && rawValue.getMonth() === 11 && rawValue.getDate() === 30) {
        // Format time part only if it's not exactly midnight
        if (rawValue.getHours() !== 0 || rawValue.getMinutes() !== 0 || rawValue.getSeconds() !== 0 || rawValue.getMilliseconds() !== 0) {
          return Utilities.formatDate(rawValue, timeZone, 'HH:mm');
        } else {
          return null; // Treat exact 1899-12-30 00:00:00 as invalid time
        }
      }
      // Format regular dates
      return Utilities.formatDate(rawValue, timeZone, 'HH:mm');
    }
    if (typeof rawValue === 'string') {
      const timeMatch = rawValue.trim().match(/^(\d{1,2}):(\d{2})(:\d{2})?(\s*(?:AM|PM))?$/i);
      if (timeMatch) {
        let hour = parseInt(timeMatch[1], 10);
        const minute = parseInt(timeMatch[2], 10);
        const ampm = (timeMatch[4] || '').trim().toUpperCase();

        if (ampm === 'PM' && hour < 12) hour += 12;
        if (ampm === 'AM' && hour === 12) hour = 0; // Midnight case

        if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
          return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
        }
      }
    }
    if (typeof rawValue === 'number' && rawValue >= 0 && rawValue <= 1) { // Includes 0 and 1
      // Using modulo arithmetic on total minutes is generally safer for precision
      const totalMinutes = Math.round(rawValue * 1440); // 24 * 60
      // Handle the edge case where rawValue is exactly 1 (representing 24:00 or 00:00 of next day technically)
      if (totalMinutes === 1440) return "00:00";

      const hours = Math.floor(totalMinutes / 60) % 24; // Use modulo 24 for hours
      const minutes = totalMinutes % 60;

      if (hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59) {
        return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
      }
    }
    return null; // Could not format
  } catch (e) {
    Logger.log(`Error in formatValueToHHMM for value "${rawValue}" (Type: ${typeof rawValue}): ${e.message}`);
    return null;
  }
}


// ==========================================================================
//                       Configuration & Authorization
// ==========================================================================

/**
 * Gets a configuration value from the 'Configuracoes' sheet. Caches values.
 * @param {string} configName - The name of the configuration setting.
 * @returns {string|Date|null} The configuration value or null if not found/error.
 */
const getConfigValue = (() => {
  const cache = {};
  let configData = null;
  let hasReadSheet = false;

  return (configName) => {
    if (configName in cache) {
      return cache[configName];
    }

    if (!configData && !hasReadSheet) {
      try {
        const sheet = getSheetByName_(SHEETS.CONFIG);
        const numRows = sheet.getLastRow();
        if (numRows < 2) {
          configData = [];
        } else {
          configData = sheet.getRange(2, HEADERS.CONFIG.NOME + 1, numRows - 1, 2).getValues();
        }
        hasReadSheet = true;
      } catch (e) {
        Logger.log(`Error reading config sheet in getConfigValue: ${e.message}`);
        configData = [];
        hasReadSheet = true;
      }
    }

    if (!configData || configData.length === 0) {
      if (!cache.hasOwnProperty(configName)) {
        Logger.log(`Configuração "${configName}" não encontrada (sheet empty or read error).`);
        cache[configName] = null;
      }
      return null;
    }

    const configRow = configData.find(row => row && row[0] === configName);

    if (configRow) {
      let value = configRow[1];
      if (value instanceof Date && !isNaN(value.getTime())) {
        if (value.getFullYear() === 1899 && value.getMonth() === 11 && value.getDate() === 30) {
          try {
            const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
            const formattedTime = formatValueToHHMM(value, timeZone);
            if (formattedTime) { cache[configName] = formattedTime; return formattedTime; }
          } catch (tzError) { /* Ignore */ }
        } else {
          cache[configName] = value;
          return value;
        }
      }
      const stringValue = String(value || '').trim();
      cache[configName] = stringValue;
      return stringValue;
    }

    if (!cache.hasOwnProperty(configName)) {
      Logger.log(`Configuração "${configName}" não encontrada na planilha "${SHEETS.CONFIG}".`);
      cache[configName] = null;
    }
    return null;
  };
})();


/**
 * Gets the role of a user based on their email from 'Usuarios Autorizados' sheet. (Internal use)
 * @param {string} userEmail - The email to look up.
 * @returns {string|null} The user's role (Admin, Professor, Aluno) or null.
 */
function getUserRolePlain_(userEmail) {
  if (!userEmail) return null;
  const trimmedEmail = userEmail.trim().toLowerCase();
  if (trimmedEmail === '') return null;

  try {
    // Read only necessary columns for optimization
    const userSheet = getSheetByName_(SHEETS.AUTHORIZED_USERS);
    const lastRow = userSheet.getLastRow();
    if (lastRow < 2) return null; // No data

    const emailCol = HEADERS.AUTHORIZED_USERS.EMAIL + 1; // 1-based
    const roleCol = HEADERS.AUTHORIZED_USERS.PAPEL + 1; // 1-based
    const maxCol = Math.max(emailCol, roleCol);
    if (userSheet.getLastColumn() < maxCol) {
      Logger.log(`WARNING: Sheet "${SHEETS.AUTHORIZED_USERS}" has fewer columns (${userSheet.getLastColumn()}) than expected (${maxCol}). Role lookup might fail.`);
      return null;
    }

    // Read only email and role columns
    const range = userSheet.getRange(2, 1, lastRow - 1, maxCol);
    const data = range.getValues();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      // Adjust column indices back to 0-based for the sliced 'data' array
      const emailInSheet = String(row[emailCol - 1] || '').trim().toLowerCase();
      if (emailInSheet === trimmedEmail) {
        const role = String(row[roleCol - 1] || '').trim();
        if (Object.values(USER_ROLES).includes(role)) {
          return role;
        } else if (role !== '') {
          Logger.log(`Invalid/Unrecognized role "${role}" found for user ${trimmedEmail}.`);
        }
        // Assuming first match is definitive, even if invalid role found later
        return null; // Found email but role is invalid or empty
      }
    }
    // Logger.log(`User "${trimmedEmail}" not found in ${SHEETS.AUTHORIZED_USERS}.`); // Can be verbose
    return null; // User email not found
  } catch (e) {
    Logger.log(`Error in getUserRolePlain_ for ${trimmedEmail}: ${e.message}`);
    return null;
  }
}

// ==========================================================================
//                        Client-Side Callable Functions
// ==========================================================================

/**
 * [CLIENT CALLABLE] Gets the current user's role and email.
 * @returns {string} JSON {success, message, data: {role, email}}
 */
function getUserRole() {
  Logger.log('*** getUserRole called ***');
  let userEmail = '[Unavailable]';
  try {
    userEmail = getActiveUserEmail_();
    const userRole = getUserRolePlain_(userEmail);
    const message = userRole ? 'Papel do usuário obtido.' : 'Usuário não encontrado ou não autorizado.';
    return createJsonResponse(true, message, { role: userRole, email: userEmail });
  } catch (e) {
    return createJsonResponse(false, e.message || 'Erro inesperado ao obter informações do usuário.', { role: null, email: userEmail });
  }
}

/**
 * [CLIENT CALLABLE] Gets a sorted list of professor names.
 * @returns {string} JSON {success, message, data: [professor names]}
 */
function getProfessorsList() {
  Logger.log('*** getProfessorsList called ***');
  try {
    const { data, header } = getSheetData_(SHEETS.AUTHORIZED_USERS, HEADERS.AUTHORIZED_USERS);
    const professors = new Set();
    const nameCol = HEADERS.AUTHORIZED_USERS.NOME;
    const roleCol = HEADERS.AUTHORIZED_USERS.PAPEL;
    const requiredCols = Math.max(nameCol, roleCol) + 1;

    if (header.length < requiredCols) {
      Logger.log(`WARNING: Sheet "${SHEETS.AUTHORIZED_USERS}" columns (${header.length}) insufficient to get Professors. Need ${requiredCols}.`);
    } else if (data.length > 0) {
      data.forEach(row => {
        if (row && row.length >= requiredCols) { // Check individual row length too
          const role = String(row[roleCol] || '').trim();
          const name = String(row[nameCol] || '').trim();
          if (role === USER_ROLES.PROFESSOR && name !== '') {
            professors.add(name);
          }
        }
      });
    }
    const sortedProfessors = Array.from(professors).sort((a, b) => a.localeCompare(b));
    Logger.log(`Found ${sortedProfessors.length} unique professors.`);
    return createJsonResponse(true, 'Lista de professores obtida com sucesso.', sortedProfessors);
  } catch (e) {
    return createJsonResponse(false, `Erro ao obter lista de professores: ${e.message}`, []);
  }
}

/**
 * [CLIENT CALLABLE] Gets the list of available class groups (Turmas) from config.
 * @returns {string} JSON {success, message, data: [turma names]}
 */
function getTurmasList() {
  Logger.log('*** getTurmasList called ***');
  try {
    const turmasConfig = getConfigValue('Turmas Disponiveis');
    if (turmasConfig === null) {
      Logger.log('Configuração "Turmas Disponiveis" não encontrada.');
      return createJsonResponse(true, 'Configuração de turmas não encontrada.', []);
    }
    if (turmasConfig === '') {
      Logger.log('Configuração "Turmas Disponiveis" está vazia.');
      return createJsonResponse(true, 'Configuração de turmas vazia.', []);
    }
    const turmasArray = turmasConfig.split(',')
      .map(t => t.trim())
      .filter(t => t !== '')
      .sort((a, b) => a.localeCompare(b));

    Logger.log(`Found ${turmasArray.length} turmas from config.`);
    return createJsonResponse(true, 'Lista de turmas (config) obtida.', turmasArray);
  } catch (e) {
    return createJsonResponse(false, `Erro ao obter lista de turmas: ${e.message}`, []);
  }
}

/**
 * [CLIENT CALLABLE] Gets the list of available disciplines from the 'Disciplinas' sheet.
 * @returns {string} JSON {success, message, data: [discipline names]}
 */
function getDisciplinesList() {
  Logger.log('*** getDisciplinesList called ***');
  try {
    const { data, header } = getSheetData_(SHEETS.DISCIPLINES, HEADERS.DISCIPLINES);
    const disciplines = new Set();
    const nameCol = HEADERS.DISCIPLINES.NOME;

    if (header.length <= nameCol) {
      Logger.log(`WARNING: Sheet "${SHEETS.DISCIPLINES}" does not have the required Name column (index ${nameCol}).`);
    } else if (data.length > 0) {
      data.forEach(row => {
        if (row && row.length > nameCol) {
          const name = String(row[nameCol] || '').trim();
          if (name !== '') {
            disciplines.add(name);
          }
        }
      });
    }
    const sortedDisciplines = Array.from(disciplines).sort((a, b) => a.localeCompare(b));
    Logger.log(`Found ${sortedDisciplines.length} unique disciplines.`);
    return createJsonResponse(true, 'Lista de disciplinas obtida com sucesso.', sortedDisciplines);
  } catch (e) {
    if (e.message.includes(`Planilha "${SHEETS.DISCIPLINES}" não encontrada`)) {
      return createJsonResponse(false, e.message, []);
    }
    return createJsonResponse(false, `Erro ao obter lista de disciplinas: ${e.message}`, []);
  }
}

/**
 * [CLIENT CALLABLE] Gets filter options (turmas, week start dates UTC Mondays) for the schedule view.
 * @returns {string} JSON {success, message, data: {turmas: [], weekStartDates: []}}
 */
function getScheduleViewFilters() {
  Logger.log('*** getScheduleViewFilters called ***');
  try {
    getActiveUserEmail_(); // Authorization check

    // Get Turmas
    const turmasResponse = JSON.parse(getTurmasList());
    const turmas = turmasResponse.success ? turmasResponse.data : [];
    if (!turmasResponse.success) {
      Logger.log("Warning: Failed to get turmas list for filters: " + turmasResponse.message);
    }

    // Calculate Week Start Dates (UTC Mondays)
    const numWeeks = parseInt(getConfigValue('Semanas Para Gerar Filtros')) || 12;
    const { startGenerationDate: firstMondayUTC } = calculateGenerationRange_(numWeeks);
    const weekStartDates = [];

    Logger.log(`Generating ${numWeeks} week start dates (UTC Mondays) for filters starting from: ${firstMondayUTC.toISOString().slice(0, 10)}`);

    for (let i = 0; i < numWeeks; i++) {
      const weekStartDate = new Date(firstMondayUTC.getTime());
      weekStartDate.setUTCDate(firstMondayUTC.getUTCDate() + (i * 7));
      const valueString = Utilities.formatDate(weekStartDate, 'UTC', 'yyyy-MM-dd');
      weekStartDates.push(valueString);
    }

    Logger.log(`Filters obtained: ${turmas.length} turmas, ${weekStartDates.length} weeks (UTC Mondays).`);
    return createJsonResponse(true, 'Filtros carregados.', { turmas: turmas, weekStartDates: weekStartDates });

  } catch (e) {
    return createJsonResponse(false, `Erro ao obter filtros de horários: ${e.message}`, null);
  }
}

/**
 * [CLIENT CALLABLE] Gets filtered schedule instances for a specific class and week.
 * @param {string} turma - The class group (turma) to filter by.
 * @param {string} weekStartDateString - The starting date of the week (Monday YYYY-MM-DD, representing UTC Monday).
 * @returns {string} JSON {success, message, data: [enriched slot details]}
 */
function getFilteredScheduleInstances(turma, weekStartDateString) {
  Logger.log(`*** getFilteredScheduleInstances called for Turma: ${turma}, Semana (expecting UTC Monday): ${weekStartDateString} ***`);
  try {
    // 1. Authorization & Validation
    getActiveUserEmail_();

    const trimmedTurma = String(turma || '').trim();
    if (!trimmedTurma) {
      return createJsonResponse(false, 'Turma não especificada.', null);
    }
    if (!weekStartDateString || typeof weekStartDateString !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(weekStartDateString)) {
      return createJsonResponse(false, 'Semana de início inválida ou formato incorreto (esperado YYYY-MM-DD).', null);
    }

    // 2. Calculate Date Range (using UTC for consistency)
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const parts = weekStartDateString.split('-');
    const weekStartDate = new Date(Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10)));

    if (isNaN(weekStartDate.getTime())) {
      return createJsonResponse(false, `Data de início da semana inválida: ${weekStartDateString}`, null);
    }
    if (weekStartDate.getUTCDay() !== 1) { // 1 = Monday in UTC
      Logger.log(`Validation Error: Provided week start date ${weekStartDateString} is not a Monday in UTC (UTC day: ${weekStartDate.getUTCDay()}).`);
      return createJsonResponse(false, `A data de início (${weekStartDateString}) não é uma Segunda-feira válida para o sistema.`, null);
    }

    const weekEndDate = new Date(weekStartDate.getTime());
    weekEndDate.setUTCDate(weekEndDate.getUTCDate() + 6);

    Logger.log(`Filtering instances for Turma "${trimmedTurma}" between UTC ${weekStartDate.toISOString().slice(0, 10)} and ${weekEndDate.toISOString().slice(0, 10)}`);

    // 3. Read Data (with column checks)
    const { data: instanceData, header: instanceHeader } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
    const { data: baseData } = getSheetData_(SHEETS.BASE_SCHEDULES);
    const { data: bookingData } = getSheetData_(SHEETS.BOOKING_DETAILS);

    // 4. Pre-process Maps
    const baseScheduleMap = baseData.reduce((map, row) => {
      const idCol = HEADERS.BASE_SCHEDULES.ID;
      const discCol = HEADERS.BASE_SCHEDULES.DISCIPLINA_PADRAO;
      const profCol = HEADERS.BASE_SCHEDULES.PROFESSOR_PRINCIPAL;
      const reqCols = Math.max(idCol, discCol, profCol) + 1;
      if (row && row.length >= reqCols) {
        const baseId = String(row[idCol] || '').trim();
        if (baseId) {
          map[baseId] = {
            disciplina: String(row[discCol] || '').trim(),
            professor: String(row[profCol] || '').trim()
          };
        }
      }
      return map;
    }, {});
    const bookingDetailsMap = bookingData.reduce((map, row) => {
      const idInstCol = HEADERS.BOOKING_DETAILS.ID_INSTANCIA;
      const discCol = HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL;
      const profRealCol = HEADERS.BOOKING_DETAILS.PROFESSOR_REAL;
      const profOrigCol = HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL;
      const statusCol = HEADERS.BOOKING_DETAILS.STATUS_RESERVA;
      const reqCols = Math.max(idInstCol, discCol, profRealCol, profOrigCol, statusCol) + 1;

      if (row && row.length >= reqCols) {
        const instanceId = String(row[idInstCol] || '').trim();
        const statusReserva = String(row[statusCol] || '').trim();
        if (instanceId && statusReserva === 'Agendada') {
          map[instanceId] = {
            disciplinaReal: String(row[discCol] || '').trim(),
            professorReal: String(row[profRealCol] || '').trim(),
            professorOriginalBooking: String(row[profOrigCol] || '').trim()
          };
        }
      }
      return map;
    }, {});

    // 5. Filter and Enrich Instance Data
    const filteredSlots = [];
    const instIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA;
    const baseIdCol = HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO;
    const turmaCol = HEADERS.SCHEDULE_INSTANCES.TURMA;
    const profPrincCol = HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL;
    const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
    const dayCol = HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA;
    const hourCol = HEADERS.SCHEDULE_INSTANCES.HORA_INICIO;
    const typeCol = HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL;
    const statusCol = HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO;
    const maxIndexNeeded = Math.max(instIdCol, baseIdCol, turmaCol, profPrincCol, dateCol, dayCol, hourCol, typeCol, statusCol);
    if (instanceHeader.length <= maxIndexNeeded) {
      throw new Error(`Planilha "${SHEETS.SCHEDULE_INSTANCES}" não contém todas as colunas necessárias (encontradas ${instanceHeader.length}, esperado pelo menos ${maxIndexNeeded + 1}).`);
    }


    instanceData.forEach((row, index) => {
      if (!row || row.length <= maxIndexNeeded) {
        return;
      }

      const instanceId = String(row[instIdCol] || '').trim();
      const baseId = String(row[baseIdCol] || '').trim();
      const instanceTurma = String(row[turmaCol] || '').trim();
      const instanceUTCDate = formatValueToDate(row[dateCol]);
      const formattedHoraInicio = formatValueToHHMM(row[hourCol], timeZone);

      if (!instanceId || !baseId || !instanceTurma || !instanceUTCDate || !formattedHoraInicio) return;
      if (instanceTurma !== trimmedTurma) return;
      if (instanceUTCDate < weekStartDate || instanceUTCDate > weekEndDate) return;

      const professorPrincipalInstance = String(row[profPrincCol] || '').trim();
      const instanceDiaSemana = String(row[dayCol] || '').trim();
      const originalType = String(row[typeCol] || '').trim();
      const instanceStatus = String(row[statusCol] || '').trim();

      let disciplinaParaExibir = '';
      let professorParaExibir = '';
      let professorOriginalNaReserva = '';
      const baseInfo = baseScheduleMap[baseId] || { disciplina: '', professor: '' };
      const bookingDetails = bookingDetailsMap[instanceId];

      if (instanceStatus === STATUS_OCUPACAO.DISPONIVEL) {
        disciplinaParaExibir = baseInfo.disciplina;
        professorParaExibir = (originalType === TIPOS_HORARIO.VAGO) ? '' : baseInfo.professor;
      } else if (bookingDetails) {
        disciplinaParaExibir = bookingDetails.disciplinaReal;
        professorParaExibir = bookingDetails.professorReal;
        professorOriginalNaReserva = bookingDetails.professorOriginalBooking;
      } else {
        Logger.log(`Warning: Instância ${instanceId} (Status: ${instanceStatus}) sem detalhes de reserva 'Agendada'. Usando dados base.`);
        disciplinaParaExibir = baseInfo.disciplina;
        professorParaExibir = professorPrincipalInstance;
      }

      filteredSlots.push({
        idInstancia: instanceId,
        data: Utilities.formatDate(instanceUTCDate, timeZone, 'dd/MM/yyyy'),
        diaSemana: instanceDiaSemana,
        horaInicio: formattedHoraInicio,
        turma: instanceTurma,
        tipoOriginal: originalType,
        statusOcupacao: instanceStatus,
        disciplinaParaExibir: disciplinaParaExibir,
        professorParaExibir: professorParaExibir,
        professorOriginalNaReserva: professorOriginalNaReserva,
        professorPrincipal: professorPrincipalInstance
      });
    });

    Logger.log(`Found ${filteredSlots.length} enriched slots for Turma "${trimmedTurma}" week starting ${weekStartDateString} (UTC).`);
    return createJsonResponse(true, `${filteredSlots.length} horários encontrados.`, filteredSlots);

  } catch (e) {
    return createJsonResponse(false, `Erro ao buscar horários filtrados: ${e.message}`, null);
  }
}


/**
 * [CLIENT CALLABLE] Gets available slots for a specific booking type.
 * @param {string} tipoReserva - 'Reposicao' or 'Substituicao'.
 * @returns {string} JSON {success, message, data: [available slots sorted]}
 */
function getAvailableSlots(tipoReserva) {
  Logger.log(`*** getAvailableSlots called for tipo: ${tipoReserva} ***`);
  try {
    // 1. Authorization & Validation
    const userEmail = getActiveUserEmail_();
    const userRole = getUserRolePlain_(userEmail);
    if (!userRole) {
      return createJsonResponse(false, 'Usuário não autorizado a buscar horários.', null);
    }
    if (tipoReserva !== TIPOS_RESERVA.REPOSICAO && tipoReserva !== TIPOS_RESERVA.SUBSTITUICAO) {
      return createJsonResponse(false, `Tipo de reserva inválido: ${tipoReserva}`, null);
    }

    // 2. Read Instance Data & Setup (Check header)
    const { data: instanceData, header: instanceHeader } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const availableSlots = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // 3. Define required columns and check header
    const instIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA;
    const baseIdCol = HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO;
    const turmaCol = HEADERS.SCHEDULE_INSTANCES.TURMA;
    const profPrincCol = HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL;
    const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
    const dayCol = HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA;
    const hourCol = HEADERS.SCHEDULE_INSTANCES.HORA_INICIO;
    const typeCol = HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL;
    const statusCol = HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO;
    const maxIndexNeeded = Math.max(instIdCol, baseIdCol, turmaCol, profPrincCol, dateCol, dayCol, hourCol, typeCol, statusCol);
    if (instanceHeader.length <= maxIndexNeeded) {
      throw new Error(`Planilha "${SHEETS.SCHEDULE_INSTANCES}" não contém todas as colunas necessárias para getAvailableSlots.`);
    }


    // 4. Filter and Format Data
    instanceData.forEach((row, index) => {
      if (!row || row.length <= maxIndexNeeded) return;

      // Extract and format essential data
      const instanceId = String(row[instIdCol] || '').trim();
      const baseId = String(row[baseIdCol] || '').trim();
      const turma = String(row[turmaCol] || '').trim();
      const professorPrincipal = String(row[profPrincCol] || '').trim();
      const rawInstanceDate = row[dateCol];
      const instanceDiaSemana = String(row[dayCol] || '').trim();
      const formattedHoraInicio = formatValueToHHMM(row[hourCol], timeZone);
      const originalType = String(row[typeCol] || '').trim();
      const instanceStatus = String(row[statusCol] || '').trim();

      if (!instanceId || !baseId || !turma || !rawInstanceDate || !instanceDiaSemana || !formattedHoraInicio || !originalType || !instanceStatus) {
        return;
      }

      // Compare using local date part
      let instanceDateForCompare = null;
      if (rawInstanceDate instanceof Date && !isNaN(rawInstanceDate.getTime())) {
        instanceDateForCompare = new Date(rawInstanceDate.getFullYear(), rawInstanceDate.getMonth(), rawInstanceDate.getDate());
        if (instanceDateForCompare < today) return;
      } else {
        return; // Skip invalid dates
      }

      // Apply filtering based on requested booking type and user role
      let isMatch = false;
      if (tipoReserva === TIPOS_RESERVA.REPOSICAO) {
        if (originalType === TIPOS_HORARIO.VAGO && instanceStatus === STATUS_OCUPACAO.DISPONIVEL) {
          if ([USER_ROLES.ADMIN, USER_ROLES.PROFESSOR, USER_ROLES.ALUNO].includes(userRole)) {
            isMatch = true;
          }
        }
      } else if (tipoReserva === TIPOS_RESERVA.SUBSTITUICAO) {
        if (originalType === TIPOS_HORARIO.FIXO && instanceStatus !== STATUS_OCUPACAO.REPOSICAO_AGENDADA) {
          if ([USER_ROLES.ADMIN, USER_ROLES.PROFESSOR].includes(userRole)) {
            if (instanceStatus === STATUS_OCUPACAO.DISPONIVEL || instanceStatus === STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA) {
              isMatch = true;
            }
          }
        }
      }

      if (isMatch) {
        availableSlots.push({
          idInstancia: instanceId,
          baseId: baseId,
          turma: turma,
          professorPrincipal: professorPrincipal,
          data: Utilities.formatDate(instanceDateForCompare, timeZone, 'dd/MM/yyyy'),
          instanceDateObj: instanceDateForCompare,
          diaSemana: instanceDiaSemana,
          horaInicio: formattedHoraInicio,
          tipoOriginal: originalType,
          statusOcupacao: instanceStatus,
        });
      }
    });

    // 5. Sort Results
    availableSlots.sort((a, b) => {
      const dateComparison = a.instanceDateObj.getTime() - b.instanceDateObj.getTime();
      if (dateComparison !== 0) return dateComparison;
      const timeComparison = a.horaInicio.localeCompare(b.horaInicio);
      if (timeComparison !== 0) return timeComparison;
      return a.turma.localeCompare(b.turma);
    });

    Logger.log(`Found ${availableSlots.length} available slots for type ${tipoReserva}.`);
    return createJsonResponse(true, 'Slots carregados com sucesso.', availableSlots);

  } catch (e) {
    return createJsonResponse(false, `Erro ao buscar horários disponíveis: ${e.message}`, null);
  }
}


/**
 * [CLIENT CALLABLE] Books a slot based on provided details. Orchestrates internal helpers.
 * @param {string} jsonBookingDetailsString - JSON string containing booking details.
 * @returns {string} JSON {success, message, data: {bookingId, eventId}}
 */
function bookSlot(jsonBookingDetailsString) {
  Logger.log(`*** bookSlot called ***`);
  let lock = null;
  let bookingId = null;
  let userEmail = '[Unavailable]';

  try {
    // 1. Initial Checks (before lock)
    userEmail = getActiveUserEmail_();
    Logger.log(`Booking attempt by: ${userEmail}. Details: ${jsonBookingDetailsString}`);
    const userRole = getUserRolePlain_(userEmail);
    if (!userRole) {
      throw new Error('Usuário não autorizado ou perfil não definido.');
    }

    let bookingDetails;
    try {
      if (!jsonBookingDetailsString) throw new Error("Dados da reserva não recebidos (null).");
      bookingDetails = JSON.parse(jsonBookingDetailsString);
    } catch (e) {
      throw new Error(`Erro interno ao processar os dados da reserva (JSON inválido): ${e.message}`);
    }

    // Validate essential structure and required fields
    const { idInstancia, tipoReserva, professorReal, disciplinaReal } = bookingDetails;
    const instanceIdToBook = String(idInstancia || '').trim();
    const bookingType = String(tipoReserva || '').trim();
    const profRealTrimmed = String(professorReal || '').trim();
    const discRealTrimmed = String(disciplinaReal || '').trim();

    if (!instanceIdToBook) throw new Error('ID da instância de horário ausente ou inválido.');
    if (bookingType !== TIPOS_RESERVA.REPOSICAO && bookingType !== TIPOS_RESERVA.SUBSTITUICAO) throw new Error('Tipo de reserva inválido ou ausente.');
    if (!profRealTrimmed) throw new Error('Professor é obrigatório.');
    if (!discRealTrimmed) throw new Error('Disciplina é obrigatória.');
    // Pass trimmed values for further processing
    bookingDetails.professorReal = profRealTrimmed;
    bookingDetails.disciplinaReal = discRealTrimmed;

    // 2. Acquire Lock
    lock = acquireScriptLock_();

    // 3. Process Booking Logic (updates sheets, performs detailed validation)
    const processResult = processBooking_(bookingDetails, userEmail, userRole);
    bookingId = processResult.bookingId;

    // 4. Handle Calendar Integration
    const calendarResult = handleCalendarIntegration_(
      getConfigValue('ID do Calendario'),
      bookingDetails,
      processResult.instanceDetails, // Use the updated instance details from processResult
      processResult.effectiveStartDateTime,
      processResult.guestEmails
    );

    // 5. Update Instance Sheet with Calendar Event ID (if successful)
    if (calendarResult.eventId && processResult.instanceRowIndex > 0) {
      try {
        const instancesSheet = getSheetByName_(SHEETS.SCHEDULE_INSTANCES);
        const eventIdCol = HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR + 1;
        instancesSheet.getRange(processResult.instanceRowIndex, eventIdCol).setValue(calendarResult.eventId);
        Logger.log(`Calendar Event ID ${calendarResult.eventId} saved to instance sheet row ${processResult.instanceRowIndex}.`);
      } catch (e) {
        Logger.log(`WARNING: Failed to save Calendar Event ID ${calendarResult.eventId} to instance sheet row ${processResult.instanceRowIndex}: ${e.message}`);
      }
    }

    // 6. Send Notification Email
    sendBookingNotificationEmail_(
      bookingId,
      bookingType,
      discRealTrimmed, // Use trimmed value
      processResult.instanceDetails[HEADERS.SCHEDULE_INSTANCES.TURMA],
      profRealTrimmed, // Use trimmed value
      processResult.professorOriginal,
      processResult.effectiveStartDateTime,
      processResult.creationTimestamp,
      userEmail,
      calendarResult.eventId,
      calendarResult.error,
      processResult.guestEmails
    );

    cleanupExcessVagoSlots();

    // 7. Prepare Success Response Message
    let successMessage = `Reserva ${bookingType} (${bookingId}) agendada com sucesso!`;
    if (calendarResult.error) {
      successMessage += ` (Aviso: ${calendarResult.error.message || 'Erro ao integrar com Google Calendar.'})`;
    } else if (calendarResult.eventId) {
      successMessage += ` Evento no calendário criado/atualizado.`;
    } else {
      successMessage += ` Não foi possível gerar evento no calendário.`;
    }
    successMessage += ` Notificação enviada.`;

    // 8. Return Success (Lock released in finally)
    return createJsonResponse(true, successMessage, { bookingId: bookingId, eventId: calendarResult.eventId });

  } catch (e) {
    Logger.log(`ERROR during bookSlot for user ${userEmail}: ${e.message}\nStack: ${e.stack}`);
    return createJsonResponse(false, `Falha no agendamento: ${e.message}`, { bookingId: bookingId });
  } finally {
    releaseScriptLock_(lock);
  }
}


// ==========================================================================
//                  Internal Logic Functions (Refactored Parts)
// ==========================================================================

/**
 * Processes the core booking logic: validates instance/permissions, updates sheets. (Internal use)
 * Assumes lock is already acquired.
 * @param {object} bookingDetails - Parsed details from client (already trimmed).
 * @param {string} userEmail - Email of the user performing the booking.
 * @param {string} userRole - Role of the user performing the booking.
 * @returns {{bookingId: string, instanceRowIndex: number, instanceDetails: any[], professorOriginal: string, effectiveStartDateTime: Date, creationTimestamp: Date, guestEmails: string[]}} Details needed for subsequent steps.
 * @throws {Error} If validation fails or sheet updates fail.
 */
function processBooking_(bookingDetails, userEmail, userRole) {
  Logger.log(`processBooking_ started for user ${userEmail} (Role: ${userRole}).`);
  const { idInstancia, tipoReserva, professorReal, disciplinaReal } = bookingDetails; // Use already validated/trimmed values
  const instanceIdToBook = idInstancia;
  const bookingType = tipoReserva;

  const { header: instanceHeader, sheet: instancesSheet } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
  const { sheet: bookingsSheet } = getSheetData_(SHEETS.BOOKING_DETAILS); // Get sheet object only
  const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

  // Find the instance row index first (more efficient than reading all data if sheet is huge)
  const instanceIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA + 1; // 1-based
  const instanceIdFinder = instancesSheet.createTextFinder(instanceIdToBook).matchEntireCell(true);
  const foundCells = instanceIdFinder.findAll();

  if (foundCells.length === 0) {
    throw new Error(`Horário com ID ${instanceIdToBook} não encontrado.`);
  }
  if (foundCells.length > 1) {
    // This shouldn't happen if IDs are unique UUIDs
    Logger.log(`WARNING: Multiple rows found for instance ID ${instanceIdToBook}. Using the first one found.`);
  }
  const instanceRowIndex = foundCells[0].getRow(); // 1-based row index
  Logger.log(`Instance ${instanceIdToBook} found at row ${instanceRowIndex}. Reading row data.`);

  // Read the specific row data now that we know the index
  const instanceDetails = instancesSheet.getRange(instanceRowIndex, 1, 1, instanceHeader.length).getValues()[0];
  const maxIndexNeeded = Math.max(...Object.values(HEADERS.SCHEDULE_INSTANCES));
  if (instanceDetails.length <= maxIndexNeeded) {
    throw new Error(`Dados da linha ${instanceRowIndex} na planilha "${SHEETS.SCHEDULE_INSTANCES}" estão incompletos.`);
  }


  // Validate instance data structure and content
  const currentStatus = String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO] || '').trim();
  const originalType = String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL] || '').trim();
  const rawBookingDate = instanceDetails[HEADERS.SCHEDULE_INSTANCES.DATA];
  const rawBookingTime = instanceDetails[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO];
  const professorPrincipalInstancia = String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL] || '').trim();
  const turmaInstancia = String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.TURMA] || '').trim();

  const bookingDateObj = formatValueToDate(rawBookingDate); // UTC Date
  const bookingHourString = formatValueToHHMM(rawBookingTime, timeZone); // HH:mm String

  if (!currentStatus || !originalType || !turmaInstancia || !bookingDateObj || !bookingHourString) {
    throw new Error(`Erro interno: Dados essenciais do horário ${instanceIdToBook} (linha ${instanceRowIndex}) são inválidos na planilha.`);
  }

  // --- Concurrency, Rule, AND PERMISSION Validation ---
  let professorOriginal = '';
  if (bookingType === TIPOS_RESERVA.REPOSICAO) {
    if (![USER_ROLES.ADMIN, USER_ROLES.PROFESSOR, USER_ROLES.ALUNO].includes(userRole)) throw new Error(`Seu perfil (${userRole}) não permite agendar Reposições.`);
    if (originalType !== TIPOS_HORARIO.VAGO) throw new Error(`Reposição só pode ser feita em horários Vagos (este é ${originalType}).`);
    if (currentStatus !== STATUS_OCUPACAO.DISPONIVEL) throw new Error(`Este horário vago (${instanceIdToBook}) não está mais disponível (Status atual: ${currentStatus}). Atualize a lista.`);
    Logger.log(`Validation OK for Reposicao by ${userRole}`);
  } else if (bookingType === TIPOS_RESERVA.SUBSTITUICAO) {
    if (![USER_ROLES.ADMIN, USER_ROLES.PROFESSOR].includes(userRole)) throw new Error(`Seu perfil (${userRole}) não permite agendar Substituições.`);
    if (originalType !== TIPOS_HORARIO.FIXO) throw new Error(`Substituição só pode ser feita em horários Fixos (este é ${originalType}).`);
    if (!professorPrincipalInstancia) throw new Error(`Erro interno: O horário fixo ${instanceIdToBook} não tem um Professor Principal definido.`);
    professorOriginal = professorPrincipalInstancia;
    if (currentStatus === STATUS_OCUPACAO.REPOSICAO_AGENDADA) throw new Error(`Este horário fixo (${instanceIdToBook}) já está ocupado por uma Reposição.`);
    if (currentStatus !== STATUS_OCUPACAO.DISPONIVEL && currentStatus !== STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA) throw new Error(`Este horário fixo (${instanceIdToBook}) não está disponível para substituição (Status atual: ${currentStatus}). Atualize a lista.`);
    Logger.log(`Validation OK for Substituicao by ${userRole}`);
  }

  // --- Prepare Updates ---
  const bookingId = Utilities.getUuid();
  const creationTimestamp = new Date();
  const newStatus = (bookingType === TIPOS_RESERVA.REPOSICAO) ? STATUS_OCUPACAO.REPOSICAO_AGENDADA : STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA;

  const [hour, minute] = bookingHourString.split(':').map(Number);
  // Construct date/time in sheet's timezone context
  const effectiveStartDateTime = new Date(bookingDateObj.getUTCFullYear(), bookingDateObj.getUTCMonth(), bookingDateObj.getUTCDate(), hour, minute);

  // --- Update Instance Sheet ---
  const updatedInstanceRow = [...instanceDetails];
  updatedInstanceRow[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO] = newStatus;
  updatedInstanceRow[HEADERS.SCHEDULE_INSTANCES.ID_RESERVA] = bookingId;
  updatedInstanceRow[HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR] = ''; // Clear old event ID
  updateSheetRow_(instancesSheet, instanceRowIndex, instanceHeader.length, updatedInstanceRow);

  // --- Add Booking Details Row ---
  const bookingHeader = bookingsSheet.getRange(1, 1, 1, bookingsSheet.getLastColumn()).getValues()[0];
  const numBookingCols = bookingHeader.length;
  const newBookingRow = new Array(numBookingCols).fill('');

  newBookingRow[HEADERS.BOOKING_DETAILS.ID_RESERVA] = bookingId;
  newBookingRow[HEADERS.BOOKING_DETAILS.TIPO_RESERVA] = bookingType;
  newBookingRow[HEADERS.BOOKING_DETAILS.ID_INSTANCIA] = instanceIdToBook;
  newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_REAL] = professorReal;
  newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL] = professorOriginal;
  newBookingRow[HEADERS.BOOKING_DETAILS.ALUNOS] = String(bookingDetails.alunos || '').trim();
  newBookingRow[HEADERS.BOOKING_DETAILS.TURMAS_AGENDADA] = turmaInstancia;
  newBookingRow[HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL] = disciplinaReal;
  newBookingRow[HEADERS.BOOKING_DETAILS.DATA_HORA_INICIO_EFETIVA] = effectiveStartDateTime;
  newBookingRow[HEADERS.BOOKING_DETAILS.STATUS_RESERVA] = 'Agendada';
  newBookingRow[HEADERS.BOOKING_DETAILS.DATA_CRIACAO] = creationTimestamp;
  newBookingRow[HEADERS.BOOKING_DETAILS.CRIADO_POR] = userEmail;

  appendSheetRow_(bookingsSheet, numBookingCols, newBookingRow);

  // --- Determine Guest Emails ---
  const guestEmails = getGuestEmailsForBooking_(professorReal, professorOriginal);

  Logger.log("processBooking_ completed successfully.");
  return {
    bookingId: bookingId,
    instanceRowIndex: instanceRowIndex,
    instanceDetails: updatedInstanceRow,
    professorOriginal: professorOriginal,
    effectiveStartDateTime: effectiveStartDateTime,
    creationTimestamp: creationTimestamp,
    guestEmails: guestEmails
  };
}


/**
 * Gets the email addresses for the relevant professors involved in a booking. (Internal use)
 * @param {string} profReal - The name of the professor performing the class.
 * @param {string} [profOrig] - The name of the original professor (for substitutions). Optional.
 * @returns {string[]} An array of unique email addresses.
 */
function getGuestEmailsForBooking_(profReal, profOrig) {
  const guests = new Set();
  const nameEmailMap = {};
  try {
    const userSheet = getSheetByName_(SHEETS.AUTHORIZED_USERS);
    const nameCol = HEADERS.AUTHORIZED_USERS.NOME + 1;
    const emailCol = HEADERS.AUTHORIZED_USERS.EMAIL + 1;
    const lastRow = userSheet.getLastRow();
    if (lastRow > 1 && userSheet.getLastColumn() >= Math.max(nameCol, emailCol)) {
      const nameRange = userSheet.getRange(2, nameCol, lastRow - 1, 1).getValues();
      const emailRange = userSheet.getRange(2, emailCol, lastRow - 1, 1).getValues();
      for (let i = 0; i < nameRange.length; i++) {
        const name = String(nameRange[i][0] || '').trim();
        const email = String(emailRange[i][0] || '').trim().toLowerCase();
        if (name && email && email.includes('@')) nameEmailMap[name] = email;
      }
      Logger.log(`Built name->email map with ${Object.keys(nameEmailMap).length} entries.`);
    } else {
      Logger.log(`Sheet "${SHEETS.AUTHORIZED_USERS}" empty or insufficient columns for guest email lookup.`);
    }
  } catch (e) {
    Logger.log(`Warning: Could not read ${SHEETS.AUTHORIZED_USERS} to get guest emails: ${e.message}`);
  }
  if (profReal && nameEmailMap[profReal]) {
    guests.add(nameEmailMap[profReal]);
    Logger.log(`Adding guest (Real): ${profReal} -> ${nameEmailMap[profReal]}`);
  } else if (profReal) {
    Logger.log(`Warning: Email for Professor Real "${profReal}" not found.`);
  }
  if (profOrig && profOrig !== profReal && nameEmailMap[profOrig]) {
    guests.add(nameEmailMap[profOrig]);
    Logger.log(`Adding guest (Original): ${profOrig} -> ${nameEmailMap[profOrig]}`);
  } else if (profOrig && profOrig !== profReal) {
    Logger.log(`Warning: Email for Professor Original "${profOrig}" not found.`);
  }
  const guestArray = Array.from(guests);
  Logger.log(`Final guest list for booking: [${guestArray.join(', ')}]`);
  return guestArray;
}

/**
 * Handles Google Calendar integration for a booking. (Internal use)
 * @param {string|null} calendarIdConfig - The Calendar ID from config (or null).
 * @param {object} bookingDetails - Original booking details from client.
 * @param {any[]} instanceDetails - The full row data of the booked instance (updated row).
 * @param {Date} effectiveStartDateTime - The calculated start time.
 * @param {string[]} guests - Array of guest email addresses.
 * @returns {{eventId: string|null, error: Error|null}} Result object.
 */
function handleCalendarIntegration_(calendarIdConfig, bookingDetails, instanceDetails, effectiveStartDateTime, guests) {
  Logger.log("handleCalendarIntegration_ started.");
  let calendarEventId = null;
  let calendarError = null;
  try {
    if (!calendarIdConfig) {
      Logger.log('Calendar ID not configured. Skipping Calendar integration.');
      return { eventId: null, error: null };
    }
    const calendar = CalendarApp.getCalendarById(calendarIdConfig.trim());
    if (!calendar) {
      Logger.log(`Calendar with ID "${calendarIdConfig}" not found or inaccessible. Skipping Calendar integration.`);
      return { eventId: null, error: new Error(`Calendário com ID "${calendarIdConfig}" não encontrado ou inacessível.`) };
    }
    Logger.log(`Accessing calendar "${calendar.getName()}" (ID: ${calendarIdConfig})`);

    let durationMinutes = parseInt(getConfigValue('Duracao Padrao Aula (minutos)')) || 45;
    const endTime = new Date(effectiveStartDateTime.getTime() + durationMinutes * 60 * 1000);
    const bookingType = String(bookingDetails.tipoReserva || '').trim();
    const disciplina = String(bookingDetails.disciplinaReal || '').trim();
    const turma = String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.TURMA] || '').trim();
    const profReal = String(bookingDetails.professorReal || '').trim();
    // Get original professor reliably from the *processed* booking details row if substitution, else from instance row's base prof
    const profOrig = (bookingType === TIPOS_RESERVA.SUBSTITUICAO)
      ? String(instanceDetails[HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL] || instanceDetails[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL] || '').trim()
      : String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL] || '').trim(); // Use base prof for context if fixed
    const bookingId = String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.ID_RESERVA] || '').trim();
    const userEmail = String(instanceDetails[HEADERS.BOOKING_DETAILS.CRIADO_POR] || getActiveUserEmail_());

    const eventTitle = `${bookingType} - ${disciplina} (${turma})`;
    let eventDescription = `Reserva ID: ${bookingId}\nProfessor: ${profReal}`;
    if (bookingType === TIPOS_RESERVA.SUBSTITUICAO && profOrig && profOrig !== profReal) {
      eventDescription += ` (Original: ${profOrig})`;
    }
    eventDescription += `\nTurma: ${turma}\nAgendado por: ${userEmail}`;

    // Read event ID from the *already updated* instance row before trying to find/create
    const existingEventId = String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR] || '').trim();
    let event = null;

    if (existingEventId) {
      try {
        event = calendar.getEventById(existingEventId);
        if (event) {
          Logger.log(`Existing Calendar event ${existingEventId} found. Updating...`);
          event.setTitle(eventTitle);
          event.setTime(effectiveStartDateTime, endTime);
          event.setDescription(eventDescription);
          updateCalendarGuests_(event, guests);
          calendarEventId = event.getId();
        } else {
          Logger.log(`Event with ID ${existingEventId} returned null. Will create a new one.`);
        }
      } catch (e) {
        Logger.log(`Failed to get/update event ${existingEventId}: ${e.message}. Creating new event.`);
        event = null;
      }
    }

    if (!event) {
      Logger.log('Creating new Calendar event.');
      const eventOptions = { description: eventDescription, conferenceDataVersion: 0 };
      if (guests && guests.length > 0) { eventOptions.guests = guests.join(','); eventOptions.sendInvites = true; }
      else { eventOptions.sendInvites = false; }
      event = calendar.createEvent(eventTitle, effectiveStartDateTime, endTime, eventOptions);
      calendarEventId = event.getId();
      Logger.log(`New event created (ID: ${calendarEventId}) without Meet link.`);
    }

  } catch (e) {
    Logger.log(`ERROR during Calendar integration: ${e.message}\nStack: ${e.stack}`);
    calendarError = e;
    calendarEventId = null;
  }
  Logger.log(`handleCalendarIntegration_ finished. Event ID: ${calendarEventId}, Error: ${calendarError ? calendarError.message : 'None'}`);
  return { eventId: calendarEventId, error: calendarError };
}


/**
 * Helper to update guests on a Calendar event. (Internal use)
 * @param {GoogleAppsScript.Calendar.CalendarEvent} event - The event object.
 * @param {string[]} newGuestEmails - Array of email addresses that *should* be guests.
 */
function updateCalendarGuests_(event, newGuestEmails) {
  if (!event || typeof event.getGuestList !== 'function') return;

  const newGuestsLower = newGuestEmails.map(g => String(g || '').toLowerCase()).filter(g => g && g.includes('@')); // Ensure valid emails
  const existingGuests = event.getGuestList();
  const existingGuestsLower = existingGuests.map(g => g.getEmail().toLowerCase());

  // Remove guests no longer in the list
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

  // Add new guests not already present
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

/**
 * Creates the content for the booking notification email. (Internal use)
 * @returns {{subject: string, bodyText: string, bodyHtml: string}} Email content.
 */
function createBookingEmailContent_(bookingId, bookingType, disciplina, turma, profReal, profOrig, startTime, timestampCriacao, userEmail, calendarEventId, calendarError) {
  const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  const dataFormatada = Utilities.formatDate(startTime, timeZone, 'dd/MM/yyyy');
  const horaFormatada = Utilities.formatDate(startTime, timeZone, 'HH:mm');
  const criacaoFormatada = Utilities.formatDate(timestampCriacao, timeZone, 'dd/MM/yyyy HH:mm:ss');
  const isSubstituicao = bookingType === TIPOS_RESERVA.SUBSTITUICAO;

  let subjectStatus = calendarError ? '⚠️ Erro no Google Calendar' : (calendarEventId ? '✅ Confirmada' : '✅ Sem Evento Calendar');
  let subject = `${subjectStatus} - Reserva ${bookingType} - ${disciplina || 'N/D'} - ${dataFormatada}`;

  let bodyText = `Olá,\n\nUma reserva de "${bookingType}" foi registrada no sistema:\n\n`;
  bodyText += `==============================\nDETALHES DA RESERVA\n==============================\n`;
  bodyText += `Tipo: ${bookingType}\nData: ${dataFormatada}\nHorário: ${horaFormatada}\nTurma: ${turma || 'N/A'}\n`;
  bodyText += `Disciplina: ${disciplina || 'N/A'}\nProfessor: ${profReal || 'N/A'}\n`;
  if (isSubstituicao && profOrig && profOrig !== profReal) bodyText += `Professor Original: ${profOrig}\n`;
  bodyText += `------------------------------\nID Reserva: ${bookingId}\nAgendado por: ${userEmail}\nData/Hora Agend.: ${criacaoFormatada}\n`;
  bodyText += `==============================\n\n`;
  if (calendarError) {
    bodyText += `*** ATENÇÃO: Google Calendar ***\nHouve um erro ao criar/atualizar o evento no calendário.\nA reserva está confirmada nas planilhas, mas verifique o calendário manualmente.\nErro: ${calendarError.message}\n\n`;
  } else if (calendarEventId) {
    bodyText += `Evento no Google Calendar criado/atualizado com sucesso (ID: ${calendarEventId}).\n\n`;
  } else {
    bodyText += `*** AVISO: Google Calendar ***\nO evento não foi criado/atualizado no calendário (ID do calendário pode não estar configurado ou calendário inacessível).\n\n`;
  }
  bodyText += `Atenciosamente,\n${EMAIL_SENDER_NAME}`;

  let bodyHtml = `<p>Olá,</p><p>Uma reserva de "<b>${bookingType}</b>" foi registrada no sistema:</p><hr><h3>Detalhes da Reserva</h3>`;
  bodyHtml += `<table border="0" cellpadding="5" style="border-collapse: collapse; font-family: sans-serif; font-size: 11pt;">`;
  bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Tipo:</strong></td><td>${bookingType}</td></tr>`;
  bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Data:</strong></td><td>${dataFormatada}</td></tr>`;
  bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Horário:</strong></td><td>${horaFormatada}</td></tr>`;
  bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Turma:</strong></td><td>${turma || 'N/A'}</td></tr>`;
  bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Disciplina:</strong></td><td>${disciplina || 'N/A'}</td></tr>`;
  bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Professor:</strong></td><td>${profReal || 'N/A'}</td></tr>`;
  if (isSubstituicao && profOrig && profOrig !== profReal) {
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Prof. Original:</strong></td><td>${profOrig}</td></tr>`;
  }
  bodyHtml += `</table><br>`;
  bodyHtml += `<table border="0" cellpadding="5" style="border-collapse: collapse; font-family: sans-serif; font-size: 9pt; color: #555;">`;
  bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><i>ID Reserva:</i></td><td><i>${bookingId}</i></td></tr>`;
  bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><i>Agendado por:</i></td><td><i>${userEmail}</i></td></tr>`;
  bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><i>Data/Hora Agend.:</i></td><td><i>${criacaoFormatada}</i></td></tr>`;
  bodyHtml += `</table><hr>`;
  if (calendarError) {
    bodyHtml += `<div style="border: 1px solid #DC3545; background-color: #F8D7DA; color: #721C24; padding: 15px; margin-top: 15px; border-radius: 4px; font-family: sans-serif; font-size: 11pt;">`;
    bodyHtml += `<strong>*** ATENÇÃO: Google Calendar ***</strong><br>Houve um erro ao criar/atualizar o evento no calendário.<br>A reserva está confirmada nas planilhas, mas verifique o calendário manualmente.<br><span style="font-size: 9pt; color: #721C24;">Erro: ${calendarError.message}</span></div>`;
  } else if (calendarEventId) {
    bodyHtml += `<div style="border: 1px solid #28A745; background-color: #D4EDDA; color: #155724; padding: 15px; margin-top: 15px; border-radius: 4px; font-family: sans-serif; font-size: 11pt;">`;
    bodyHtml += `Evento no Google Calendar criado/atualizado com sucesso.<br>ID do Evento: ${calendarEventId}</div>`;
  } else {
    bodyHtml += `<div style="border: 1px solid #FFC107; background-color: #FFF3CD; color: #856404; padding: 15px; margin-top: 15px; border-radius: 4px; font-family: sans-serif; font-size: 11pt;">`;
    bodyHtml += `<strong>*** AVISO: Google Calendar ***</strong><br>O evento não foi criado/atualizado no calendário (ID do calendário pode não estar configurado ou calendário inacessível).</div>`;
  }
  bodyHtml += `<p style="font-family: sans-serif; font-size: 11pt; margin-top: 20px;">Atenciosamente,<br>${EMAIL_SENDER_NAME}</p>`;

  return { subject, bodyText, bodyHtml };
}


/**
 * Sends the booking notification email using MailApp. (Internal use)
 */
function sendBookingNotificationEmail_(bookingId, bookingType, disciplina, turma, profReal, profOrig, startTime, timestampCriacao, userEmail, calendarEventId, calendarError, guests) {
  Logger.log("sendBookingNotificationEmail_ called.");
  try {
    const emailContent = createBookingEmailContent_(bookingId, bookingType, disciplina, turma, profReal, profOrig, startTime, timestampCriacao, userEmail, calendarEventId, calendarError);
    const validGuests = Array.isArray(guests) ? guests.filter(email => email && typeof email === 'string' && email.includes('@')) : [];
    const validAdminEmails = ADMIN_COPY_EMAILS.filter(email => email && typeof email === 'string' && email.includes('@'));
    const finalRecipientsBcc = [...new Set([...validGuests, ...validAdminEmails])];

    if (finalRecipientsBcc.length === 0) {
      Logger.log("No valid recipients found. Skipping email send.");
      return;
    }

    const toAddress = validAdminEmails[0] || userEmail; // Prefer admin, fallback to user

    Logger.log(`Sending notification for Booking ID ${bookingId}. To: ${toAddress}, BCC: ${finalRecipientsBcc.length} addresses.`);
    MailApp.sendEmail({
      to: toAddress,
      bcc: finalRecipientsBcc.join(','),
      subject: emailContent.subject,
      body: emailContent.bodyText,
      htmlBody: emailContent.bodyHtml,
      name: EMAIL_SENDER_NAME
    });
    Logger.log(`Email notification for Booking ID ${bookingId} sent successfully via MailApp.`);

  } catch (e) {
    Logger.log(`ERROR sending booking notification email for ID ${bookingId}: ${e.message}\nStack: ${e.stack}`);
  }
}


// ==========================================================================
//                         Instance Generation & Cleanup
// ==========================================================================

/**
 * [TRIGGERED/MANUAL] Creates future schedule instances based on 'Horarios Base'.
 */
function createScheduleInstances() {
  Logger.log('*** createScheduleInstances START ***');
  let lock = null;
  try {
    lock = acquireScriptLock_();
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const { header: baseHeader, data: baseData } = getSheetData_(SHEETS.BASE_SCHEDULES, HEADERS.BASE_SCHEDULES);
    const { header: instanceHeader, data: instanceData, sheet: instancesSheet } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);

    if (baseData.length === 0) throw new Error('Planilha "Horarios Base" está vazia.');

    const validBaseSchedules = validateBaseSchedules_(baseData, timeZone);
    if (validBaseSchedules.length === 0) throw new Error('Nenhum horário base válido encontrado após validação.');
    Logger.log(`Found ${validBaseSchedules.length} valid base schedules.`);

    const existingInstanceKeys = createExistingInstanceMap_(instanceData, timeZone);
    Logger.log(`Created map with ${Object.keys(existingInstanceKeys).length} existing instance keys.`);

    const numWeeksToGenerate = parseInt(getConfigValue('Semanas Para Gerar Instancias')) || 4;
    const { startGenerationDate, endGenerationDate } = calculateGenerationRange_(numWeeksToGenerate);
    Logger.log(`Generating instances from UTC ${startGenerationDate.toISOString().slice(0, 10)} to ${endGenerationDate.toISOString().slice(0, 10)}`);

    const newInstanceRows = generateNewInstances_(
      startGenerationDate,
      endGenerationDate,
      validBaseSchedules,
      existingInstanceKeys,
      timeZone
    );

    if (newInstanceRows.length > 0) {
      Logger.log(`Generated ${newInstanceRows.length} new instances. Appending to sheet...`);
      const numInstanceCols = instanceHeader.length || (Math.max(...Object.values(HEADERS.SCHEDULE_INSTANCES)) + 1);
      appendSheetRows_(instancesSheet, numInstanceCols, newInstanceRows);
    } else {
      Logger.log('No new instances needed for the specified period.');
    }

    Logger.log('*** createScheduleInstances FINISHED ***');
  } catch (e) {
    Logger.log(`ERROR in createScheduleInstances: ${e.message}\nStack: ${e.stack}`);
    // Optionally re-throw or handle error notification
  } finally {
    releaseScriptLock_(lock);
  }
}

/**
 * Validates rows from the Base Schedules sheet. (Internal use for createScheduleInstances)
 * @param {any[][]} baseData - Raw data rows from the sheet.
 * @param {string} timeZone - Spreadsheet timezone.
 * @returns {object[]} Array of validated base schedule objects.
 */
function validateBaseSchedules_(baseData, timeZone) {
  const validSchedules = [];
  const daysOfWeek = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
  const idCol = HEADERS.BASE_SCHEDULES.ID;
  const dayCol = HEADERS.BASE_SCHEDULES.DIA_SEMANA;
  const hourCol = HEADERS.BASE_SCHEDULES.HORA_INICIO;
  const typeCol = HEADERS.BASE_SCHEDULES.TIPO;
  const turmaCol = HEADERS.BASE_SCHEDULES.TURMA_PADRAO;
  const profCol = HEADERS.BASE_SCHEDULES.PROFESSOR_PRINCIPAL;
  const reqCols = Math.max(idCol, dayCol, hourCol, typeCol, turmaCol, profCol) + 1;

  baseData.forEach((row, index) => {
    const rowIndex = index + 2; // Sheet row index
    if (!row || row.length < reqCols) {
      // Logger.log(`Skipping base schedule row ${rowIndex} due to insufficient columns.`); // Verbose
      return;
    }

    const baseId = String(row[idCol] || '').trim();
    const baseDayOfWeek = String(row[dayCol] || '').trim();
    const baseHourString = formatValueToHHMM(row[hourCol], timeZone);
    const baseType = String(row[typeCol] || '').trim();
    const baseTurma = String(row[turmaCol] || '').trim();
    const baseProfessorPrincipal = String(row[profCol] || '').trim();

    let isValid = true;
    const errorMessages = [];
    if (!baseId) { errorMessages.push("ID Base inválido"); isValid = false; }
    if (!baseDayOfWeek || !daysOfWeek.includes(baseDayOfWeek)) { errorMessages.push(`Dia da Semana inválido: ${baseDayOfWeek}`); isValid = false; }
    if (!baseHourString) { errorMessages.push(`Hora inválida: ${row[hourCol]}`); isValid = false; }
    if (baseType !== TIPOS_HORARIO.FIXO && baseType !== TIPOS_HORARIO.VAGO) { errorMessages.push(`Tipo inválido: ${baseType}`); isValid = false; }
    if (!baseTurma) { errorMessages.push("Turma Padrão inválida"); isValid = false; }
    if (baseType === TIPOS_HORARIO.FIXO && !baseProfessorPrincipal) { errorMessages.push("Professor Principal ausente para horário Fixo"); isValid = false; }

    if (isValid) {
      validSchedules.push({
        id: baseId,
        dayOfWeek: baseDayOfWeek,
        hour: baseHourString,
        type: baseType,
        turma: baseTurma,
        professorPrincipal: baseProfessorPrincipal
      });
    } else {
      Logger.log(`Skipping Base Schedule row ${rowIndex}: ${errorMessages.join(', ')}.`);
    }
  });
  return validSchedules;
}

/**
 * Creates a map of existing instance keys for duplicate checking. (Internal use for createScheduleInstances)
 * Key format: "baseId_YYYY-MM-DD_HH:mm" (Date is UTC)
 * @param {any[][]} instanceData - Raw instance data rows.
 * @param {string} timeZone - Spreadsheet timezone.
 * @returns {object} Map where keys are unique instance identifiers and values can be true or the instance ID.
 */
function createExistingInstanceMap_(instanceData, timeZone) {
  const existingKeys = {};
  const baseIdCol = HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO;
  const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
  const hourCol = HEADERS.SCHEDULE_INSTANCES.HORA_INICIO;
  const reqCols = Math.max(baseIdCol, dateCol, hourCol) + 1;

  instanceData.forEach((row, index) => {
    if (!row || row.length < reqCols) return; // Skip incomplete

    const baseId = String(row[baseIdCol] || '').trim();
    // Use UTC date for the key to be consistent regardless of DST/TZ
    const instanceUTCDate = formatValueToDate(row[dateCol]); // Gets UTC Date obj
    const hourString = formatValueToHHMM(row[hourCol], timeZone); // HH:mm based on sheet TZ

    if (baseId && instanceUTCDate && hourString) {
      // Format UTC date for the key
      const dateString = Utilities.formatDate(instanceUTCDate, 'UTC', 'yyyy-MM-dd');
      const key = `${baseId}_${dateString}_${hourString}`;
      existingKeys[key] = true; // Value doesn't strictly matter, just existence
    } else {
      // Logger.log(`Could not create key for existing instance row ${index + 2}. Missing data?`); // Too verbose
    }
  });
  return existingKeys;
}

/**
 * Calculates the start (next Mon UTC) and end dates (Sun UTC) for generation. (Internal use)
 * @param {number} numWeeksToGenerate - How many weeks ahead to generate.
 * @returns {{startGenerationDate: Date, endGenerationDate: Date}} UTC Dates.
 */
function calculateGenerationRange_(numWeeksToGenerate) {
  const now = new Date();
  const todayUTC = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
  const dayUTC = todayUTC.getUTCDay();
  const daysToAdd = (dayUTC === 0) ? 1 : (8 - dayUTC) % 7;
  const start = new Date(todayUTC.getTime());
  start.setUTCDate(todayUTC.getUTCDate() + daysToAdd);
  const end = new Date(start.getTime());
  end.setUTCDate(start.getUTCDate() + (numWeeksToGenerate * 7) - 1);
  return { startGenerationDate: start, endGenerationDate: end };
}

/**
 * Generates the data arrays for new instances, avoiding duplicates. (Internal use for createScheduleInstances)
 * @param {Date} startDateUTC - First date (UTC midnight) to generate for.
 * @param {Date} endDateUTC - Last date (UTC midnight) to generate for.
 * @param {object[]} validBaseSchedules - Array of validated base schedule objects.
 * @param {object} existingInstanceKeys - Map of existing instance keys ("baseId_YYYY-MM-DD_HH:mm").
 * @param {string} timeZone - Spreadsheet timezone (for HH:mm key matching).
 * @returns {any[][]} Array of rows to be appended.
 */
function generateNewInstances_(startDateUTC, endDateUTC, validBaseSchedules, existingInstanceKeys, timeZone) {
  const newInstanceRows = [];
  const daysOfWeekMap = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
  const numInstanceCols = Math.max(...Object.values(HEADERS.SCHEDULE_INSTANCES)) + 1;

  let currentDate = new Date(startDateUTC.getTime());
  while (currentDate <= endDateUTC) {
    const targetUTCDate = new Date(currentDate.getTime());
    const targetDayName = daysOfWeekMap[targetUTCDate.getUTCDay()];
    const targetDateStr = Utilities.formatDate(targetUTCDate, 'UTC', 'yyyy-MM-dd');

    const applicableSchedules = validBaseSchedules.filter(s => s.dayOfWeek === targetDayName);

    for (const baseSchedule of applicableSchedules) {
      // Key uses UTC date string + HH:mm string (which implicitly uses sheet TZ from formatValueToHHMM)
      const predictableKey = `${baseSchedule.id}_${targetDateStr}_${baseSchedule.hour}`;

      if (!existingInstanceKeys[predictableKey]) {
        const newRow = new Array(numInstanceCols).fill('');
        newRow[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA] = Utilities.getUuid();
        newRow[HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO] = baseSchedule.id;
        newRow[HEADERS.SCHEDULE_INSTANCES.TURMA] = baseSchedule.turma;
        newRow[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL] = baseSchedule.professorPrincipal;
        // Store the Date object representing the specific day (in sheet's TZ context for display later)
        newRow[HEADERS.SCHEDULE_INSTANCES.DATA] = new Date(targetUTCDate.getUTCFullYear(), targetUTCDate.getUTCMonth(), targetUTCDate.getUTCDate());
        newRow[HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA] = targetDayName;
        newRow[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO] = baseSchedule.hour; // Store HH:mm string
        newRow[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL] = baseSchedule.type;
        newRow[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO] = STATUS_OCUPACAO.DISPONIVEL;
        // ID_RESERVA and ID_EVENTO_CALENDAR start empty

        newInstanceRows.push(newRow);
        existingInstanceKeys[predictableKey] = true; // Mark as generated
      }
    }
    currentDate.setUTCDate(currentDate.getUTCDate() + 1);
  }
  return newInstanceRows;
}

/**
 * Appends multiple rows to a sheet efficiently. (Internal use)
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet - The sheet object.
 * @param {number} numCols - The number of columns expected/to write.
 * @param {any[][]} rowsToAppend - 2D array of row data.
 */
function appendSheetRows_(sheet, numCols, rowsToAppend) {
  if (!rowsToAppend || rowsToAppend.length === 0) {
    // Logger.log("No rows to append."); // Can be verbose
    return;
  }
  try {
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


/**
 * [TRIGGERED/MANUAL] Cleans old schedule instances based on 'Data Limite' config.
 */
function cleanOldScheduleInstances() {
  Logger.log('*** cleanOldScheduleInstances START ***');
  let lock = null;
  try {
    lock = acquireScriptLock_();
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();

    const cleanupDateString = getConfigValue('Data Limite');
    if (!cleanupDateString) throw new Error(`Configuração "Data Limite" não encontrada ou vazia.`);

    const cleanupDateUTC = parseDDMMYYYY(cleanupDateString);
    if (!cleanupDateUTC) throw new Error(`Valor da configuração "Data Limite" inválido: "${cleanupDateString}". Use dd/MM/yyyy.`);

    Logger.log(`Cleaning instances strictly BEFORE UTC date: ${cleanupDateUTC.toISOString().slice(0, 10)}`);

    const { header: instanceHeader, data: instanceData, sheet: instancesSheet } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
    const originalRowCount = instanceData.length;
    if (originalRowCount === 0) {
      Logger.log('No instances found to clean.');
      releaseScriptLock_(lock); return;
    }
    const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
    const numCols = instanceHeader.length;
    if (dateCol >= numCols) throw new Error(`Coluna de Data (índice ${dateCol}) não encontrada na planilha "${SHEETS.SCHEDULE_INSTANCES}".`);

    const rowsToKeep = [];
    instanceData.forEach((row) => {
      if (row && row.length > dateCol) {
        const instanceUTCDate = formatValueToDate(row[dateCol]); // Gets UTC Date obj
        // Keep if date is valid AND on or after the cleanup date (UTC comparison)
        if (instanceUTCDate && instanceUTCDate >= cleanupDateUTC) {
          rowsToKeep.push(row);
        }
      }
      // Rows that are incomplete, have invalid dates, or are before cleanupDateUTC are implicitly skipped
    });

    const deletedCount = originalRowCount - rowsToKeep.length;
    Logger.log(`Filtering complete: ${rowsToKeep.length} rows to keep, ${deletedCount} rows to delete.`);

    if (deletedCount > 0) {
      Logger.log(`Rewriting sheet "${SHEETS.SCHEDULE_INSTANCES}"...`);
      const dataToWrite = [instanceHeader, ...rowsToKeep].map(row => {
        const paddedRow = [...row];
        while (paddedRow.length < numCols) paddedRow.push('');
        if (paddedRow.length > numCols) return paddedRow.slice(0, numCols);
        return paddedRow;
      });

      instancesSheet.clearContents();
      if (dataToWrite.length > 0) {
        instancesSheet.getRange(1, 1, dataToWrite.length, numCols).setValues(dataToWrite);
      }
      Logger.log(`Sheet rewritten with ${rowsToKeep.length} data rows.`);
    } else {
      Logger.log('No instances found before the cleanup date. No changes made to the sheet.');
    }

    Logger.log('*** cleanOldScheduleInstances FINISHED ***');
  } catch (e) {
    Logger.log(`ERROR in cleanOldScheduleInstances: ${e.message}\nStack: ${e.stack}`);
  } finally {
    releaseScriptLock_(lock);
  }
}


// ==========================================================================
//                     Web App Entry Point & Include Helper
// ==========================================================================

/**
 * Main function for serving the web app via GET request. Handles routing and authorization.
 * @param {object} e - Apps Script event object.
 * @returns {HtmlService.HtmlOutput} Rendered HTML output.
 */
function doGet(e) {
  let userEmail = '[Public/Unknown]'; // Default for public/unknown or error
  const page = e && e.parameter ? e.parameter.page : null;
  let userRole = null; // Initialize userRole

  // --- Try to get user info regardless of page initially for logging/potential use ---
  try {
    userEmail = getActiveUserEmail_();
    userRole = getUserRolePlain_(userEmail); // Get role early if possible
  } catch (err) {
    // Ignore error here, might be public access attempt
    Logger.log(`Attempted access, couldn't get user info initially (may be public): ${err.message}`);
  }
  Logger.log(`Web App GET request by: ${userEmail}. Page Parameter: ${page}. Detected Role: ${userRole}`);


  // --- Route: Public View ---
  if (page === 'public') {
    Logger.log(`Serving public view page.`);
    try {
      const template = HtmlService.createTemplateFromFile('PublicView');
      return template.evaluate()
        .setTitle('Horários Públicos')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    } catch (error) {
      Logger.log(`FATAL ERROR loading public view: ${error.message} Stack: ${error.stack}`);
      return HtmlService.createHtmlOutput('<h1>Erro Interno</h1><p>Erro ao carregar visualização pública.</p>').setTitle('Erro');
    }
  }

  // --- Route: Cancel View (Requires Admin Role) ---
  else if (page === 'cancel') {
    Logger.log(`Attempting to serve cancel view page for user: ${userEmail}`);
    // **Strict Admin Check for this page**
    if (userRole === USER_ROLES.ADMIN) {
      Logger.log(`Admin access granted for cancel page.`);
      try {
        const template = HtmlService.createTemplateFromFile('CancelView');
        // Pass user info if needed by the template, though JS usually fetches
        // template.adminEmail = userEmail;
        return template.evaluate()
          .setTitle('Cancelar Reservas')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } catch (error) {
        Logger.log(`FATAL ERROR loading cancel view: ${error.message} Stack: ${error.stack}`);
        return HtmlService.createHtmlOutput('<h1>Erro Interno</h1><p>Erro ao carregar página de cancelamento.</p>').setTitle('Erro');
      }
    } else {
      // Deny access if not Admin
      Logger.log(`Access Denied for user ${userEmail} to cancel page. Role: ${userRole}`);
      return HtmlService
        .createHtmlOutput(`<h1>Acesso Negado</h1><p>Apenas administradores podem acessar esta página (${userEmail}).</p>`)
        .setTitle('Acesso Negado');
    }
  }

  // --- Default Route: Main App (Requires Any Authorized Role) ---
  else {
    Logger.log(`Attempting to serve main page for user: ${userEmail}`);
    // Check if user has *any* authorized role for the main app
    if (userRole) {
      Logger.log(`Serving main page (Index.html) for user ${userEmail} with role ${userRole}.`);
      try {
        const template = HtmlService.createTemplateFromFile('Index');
        return template.evaluate()
          .setTitle('Sistema de Agendamento e Visualização')
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      } catch (error) {
        Logger.log(`FATAL ERROR loading main Index page: ${error.message} Stack: ${error.stack}`);
        return HtmlService.createHtmlOutput('<h1>Erro Interno</h1><p>Erro ao carregar aplicação principal.</p>').setTitle('Erro');
      }
    } else {
      // Deny access if no authorized role for the main app
      Logger.log(`Access Denied for user ${userEmail} to main application. Role: ${userRole}`);
      return HtmlService
        .createHtmlOutput(`<h1>Acesso Negado</h1><p>Seu usuário (${userEmail}) não tem permissão para acessar esta aplicação. Entre em contato com o administrador.</p>`)
        .setTitle('Acesso Negado');
    }
  }
}

/**
 * [CLIENT CALLABLE - For Public View] Gets relevant schedule instances for ALL classes for a specific week.
 * Only returns 'Fixo' slots (any status) and 'Vago' slots that are booked ('Reposicao Agendada').
 * Enriches available 'Fixo' slots with base discipline and professor.
 * No authorization check needed here.
 * @param {string} weekStartDateString - The starting date of the week (Monday, YYYY-MM-DD, representing UTC Monday).
 * @returns {string} JSON {success, message, data: { "TurmaName1": [slots...], "TurmaName2": [slots...] }}
 */
function getPublicScheduleInstances(weekStartDateString) {
  Logger.log(`*** getPublicScheduleInstances (Public View - Fixed/Booked Only) called for Semana: ${weekStartDateString} ***`);
  try {
    // 1. Validation (No Auth check here)
    if (!weekStartDateString || typeof weekStartDateString !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(weekStartDateString)) {
      return createJsonResponse(false, 'Semana de início inválida ou formato incorreto (esperado YYYY-MM-DD).', null);
    }

    // 2. Calculate Date Range (using UTC)
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(); // Still needed for display formatting & HH:mm
    const parts = weekStartDateString.split('-');
    const weekStartDate = new Date(Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10)));

    if (isNaN(weekStartDate.getTime()) || weekStartDate.getUTCDay() !== 1) { // Check UTC Monday
      return createJsonResponse(false, `A data de início (${weekStartDateString}) não é uma Segunda-feira válida para o sistema.`, null);
    }

    const weekEndDate = new Date(weekStartDate.getTime());
    weekEndDate.setUTCDate(weekEndDate.getUTCDate() + 6);
    Logger.log(`Filtering Public (Fixed/Booked) instances between UTC ${weekStartDate.toISOString().slice(0, 10)} and ${weekEndDate.toISOString().slice(0, 10)}`);

    // 3. Read Data
    const { data: instanceData } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
    const { data: baseData } = getSheetData_(SHEETS.BASE_SCHEDULES); // Needed for base info
    const { data: bookingData } = getSheetData_(SHEETS.BOOKING_DETAILS);

    // 4. Pre-process Maps (Make sure baseScheduleMap includes professor)
    const baseScheduleMap = baseData.reduce((map, row) => {
      const idCol = HEADERS.BASE_SCHEDULES.ID;
      const discCol = HEADERS.BASE_SCHEDULES.DISCIPLINA_PADRAO;
      const profCol = HEADERS.BASE_SCHEDULES.PROFESSOR_PRINCIPAL; // Ensure professor is included
      const reqCols = Math.max(idCol, discCol, profCol) + 1;
      if (row && row.length >= reqCols) {
        const baseId = String(row[idCol] || '').trim();
        if (baseId) {
          map[baseId] = {
            disciplina: String(row[discCol] || '').trim(),
            professor: String(row[profCol] || '').trim() // Store professor
          };
        }
      } return map;
    }, {});
    const bookingDetailsMap = bookingData.reduce((map, row) => {
      const idInstCol = HEADERS.BOOKING_DETAILS.ID_INSTANCIA; const discCol = HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL; const profRealCol = HEADERS.BOOKING_DETAILS.PROFESSOR_REAL; const profOrigCol = HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL; const statusCol = HEADERS.BOOKING_DETAILS.STATUS_RESERVA; const reqCols = Math.max(idInstCol, discCol, profRealCol, profOrigCol, statusCol) + 1;
      if (row && row.length >= reqCols) { const instanceId = String(row[idInstCol] || '').trim(); const statusReserva = String(row[statusCol] || '').trim(); if (instanceId && statusReserva === 'Agendada') map[instanceId] = { disciplinaReal: String(row[discCol] || '').trim(), professorReal: String(row[profRealCol] || '').trim(), professorOriginalBooking: String(row[profOrigCol] || '').trim() }; } return map;
    }, {});

    // 5. Filter and Enrich Instance Data, THEN Group by Turma
    const slotsByTurma = {};
    // ... (column index definitions) ...
    const instIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA; const baseIdCol = HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO; const turmaCol = HEADERS.SCHEDULE_INSTANCES.TURMA; const profPrincCol = HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL; const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA; const dayCol = HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA; const hourCol = HEADERS.SCHEDULE_INSTANCES.HORA_INICIO; const typeCol = HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL; const statusCol = HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO; const maxIndexNeeded = Math.max(instIdCol, baseIdCol, turmaCol, profPrincCol, dateCol, dayCol, hourCol, typeCol, statusCol);


    instanceData.forEach((row, index) => {
      if (!row || row.length <= maxIndexNeeded) return;

      const instanceId = String(row[instIdCol] || '').trim();
      const baseId = String(row[baseIdCol] || '').trim();
      const instanceTurma = String(row[turmaCol] || '').trim();
      const instanceUTCDate = formatValueToDate(row[dateCol]);
      const formattedHoraInicio = formatValueToHHMM(row[hourCol], timeZone);

      if (!instanceId || !baseId || !instanceTurma || !instanceUTCDate || !formattedHoraInicio) return;
      if (instanceUTCDate < weekStartDate || instanceUTCDate > weekEndDate) return;

      const originalType = String(row[typeCol] || '').trim();
      const instanceStatus = String(row[statusCol] || '').trim();

      // *** PUBLIC VIEW FILTER LOGIC ***
      let includeSlot = false;
      if (originalType === TIPOS_HORARIO.FIXO) {
        includeSlot = true;
      } else if (originalType === TIPOS_HORARIO.VAGO && instanceStatus === STATUS_OCUPACAO.REPOSICAO_AGENDADA) {
        includeSlot = true;
      }
      if (!includeSlot) return;
      // *** END FILTERING LOGIC ***

      // Extract remaining info needed for enrichment (only if included)
      const professorPrincipalInstancia = String(row[profPrincCol] || '').trim(); // Professor from instance row itself
      const instanceDiaSemana = String(row[dayCol] || '').trim();

      // Enrichment logic
      let disciplinaParaExibir = '';
      let professorParaExibir = '';
      let professorOriginalNaReserva = ''; // From booking details specifically
      const baseInfo = baseScheduleMap[baseId] || { disciplina: '', professor: '' }; // Get base info using baseId
      const bookingDetails = bookingDetailsMap[instanceId]; // Get booking details using instanceId

      if (instanceStatus === STATUS_OCUPACAO.DISPONIVEL) {
        // *** FIX: Use baseInfo for Available Fixed slots ***
        disciplinaParaExibir = baseInfo.disciplina;
        // Use professor from base schedule map, NOT from the instance row directly here for base display
        professorParaExibir = baseInfo.professor;
        // If somehow base professor is empty for a fixed slot, fallback or mark as N/D
        if (!professorParaExibir && originalType === TIPOS_HORARIO.FIXO) {
          Logger.log(`Warning (Public View): Base Professor not found in map for available Fixed slot ${instanceId} (Base ID: ${baseId}).`);
          professorParaExibir = 'Prof. N/D';
        }
      } else if (bookingDetails) { // Booked Fixed (Substituicao) or booked Vago (Reposicao)
        disciplinaParaExibir = bookingDetails.disciplinaReal;
        professorParaExibir = bookingDetails.professorReal;
        professorOriginalNaReserva = bookingDetails.professorOriginalBooking;
      } else { // Booked status but details missing (data inconsistency)
        Logger.log(`Warning (Public View): Instância ${instanceId} (Status: ${instanceStatus}) sem detalhes de reserva 'Agendada'. Usando dados base.`);
        // Fallback to base info, using professor from instance row as last resort if base map failed
        disciplinaParaExibir = baseInfo.disciplina;
        professorParaExibir = professorPrincipalInstancia || baseInfo.professor || 'Prof. N/D';
      }

      // Prepare final slot data object
      const slotData = {
        data: Utilities.formatDate(instanceUTCDate, timeZone, 'dd/MM/yyyy'),
        diaSemana: instanceDiaSemana,
        horaInicio: formattedHoraInicio,
        tipoOriginal: originalType,
        statusOcupacao: instanceStatus,
        disciplinaParaExibir: disciplinaParaExibir,
        professorParaExibir: professorParaExibir,
        professorOriginalNaReserva: professorOriginalNaReserva,
        // Include professorPrincipalInstancia for context if needed (e.g., tooltip)
        professorPrincipal: professorPrincipalInstancia
      };

      // Group by Turma
      if (!slotsByTurma[instanceTurma]) {
        slotsByTurma[instanceTurma] = [];
      }
      slotsByTurma[instanceTurma].push(slotData);
    });

    const turmaCount = Object.keys(slotsByTurma).length;
    Logger.log(`Found relevant public instances for ${turmaCount} turmas for week starting ${weekStartDateString} (UTC).`);
    return createJsonResponse(true, `Horários encontrados para ${turmaCount} turma(s).`, slotsByTurma);

  } catch (e) {
    return createJsonResponse(false, `Erro ao buscar horários públicos: ${e.message}`, null);
  }
}

/**
 * [TRIGGERED/MANUAL] Removes available 'Vago' instances for a specific Turma/Date
 * if the total count of 'Fixo' slots + booked ('Vago' or 'Fixo') slots
 * reaches or exceeds a configured threshold for that day/turma.
 */
function cleanupExcessVagoSlots() {
  Logger.log('*** cleanupExcessVagoSlots START ***');
  let lock = null;
  try {
    // 1. Acquire Lock (essential for safe deletion)
    lock = acquireScriptLock_();

    // 2. Get Configuration
    const threshold = parseInt(getConfigValue('Limite Maximo Aulas Dia Turma')) || 10; // Default to 10
    Logger.log(`Using threshold: ${threshold} (Fixo + Vago Booked)`);

    // 3. Read Instance Data
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const { header: instanceHeader, data: instanceData, sheet: instancesSheet } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
    const originalRowCount = instanceData.length;
    if (originalRowCount === 0) {
      Logger.log('No instances found to process.');
      releaseScriptLock_(lock); return;
    }

    // Define column indices
    const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
    const turmaCol = HEADERS.SCHEDULE_INSTANCES.TURMA;
    const typeCol = HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL;
    const statusCol = HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO;
    const maxIndexNeeded = Math.max(dateCol, turmaCol, typeCol, statusCol);
    if (instanceHeader.length <= maxIndexNeeded) {
      throw new Error(`Planilha "${SHEETS.SCHEDULE_INSTANCES}" não contém todas as colunas necessárias para cleanupExcessVagoSlots.`);
    }
    const today = new Date(); // Use local date for "today" comparison
    today.setHours(0, 0, 0, 0);

    // 4. Group Data by Date (UTC String) and Turma
    const groupedData = {}; // Format: groupedData[dateString][turmaName] = { fixoCount: 0, vagoBookedCount: 0, availableVagoRows: [] }

    Logger.log(`Processing ${originalRowCount} instance rows...`);
    instanceData.forEach((row, index) => {
      if (!row || row.length <= maxIndexNeeded) return; // Skip incomplete

      const instanceUTCDate = formatValueToDate(row[dateCol]); // Get UTC Date obj
      const turma = String(row[turmaCol] || '').trim();
      const originalType = String(row[typeCol] || '').trim();
      const instanceStatus = String(row[statusCol] || '').trim();

      // Validate essential fields for grouping/counting
      if (!instanceUTCDate || !turma) return;

      // --- Optional: Skip processing dates in the past? ---
      // Create a local date representation for comparison with 'today'
      let instanceLocalDate = null;
      const rawDate = row[dateCol];
      if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
        instanceLocalDate = new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate());
        if (instanceLocalDate < today) {
          // Logger.log(`Skipping past date row ${index + 2}`); // Can be verbose
          return; // Skip past dates for this logic
        }
      } else {
        // Logger.log(`Skipping row ${index + 2} due to invalid date object`); // Can be verbose
        return; // Skip if raw date wasn't valid date object
      }
      // --- End Past Date Skip ---

      const dateStringKey = Utilities.formatDate(instanceUTCDate, 'UTC', 'yyyy-MM-dd'); // Use UTC date string as key

      // Initialize group if needed
      if (!groupedData[dateStringKey]) groupedData[dateStringKey] = {};
      if (!groupedData[dateStringKey][turma]) {
        groupedData[dateStringKey][turma] = { fixoCount: 0, vagoBookedCount: 0, availableVagoRows: [] };
      }

      // Count based on type and status
      if (originalType === TIPOS_HORARIO.FIXO) {
        groupedData[dateStringKey][turma].fixoCount++;
      } else if (originalType === TIPOS_HORARIO.VAGO) {
        if (instanceStatus === STATUS_OCUPACAO.DISPONIVEL) {
          // Store the actual sheet row index (1-based)
          groupedData[dateStringKey][turma].availableVagoRows.push(index + 2);
        } else {
          // Count booked Vago slots (Reposicao/Substituicao shouldn't happen but count any non-available)
          groupedData[dateStringKey][turma].vagoBookedCount++;
        }
      }
    });
    Logger.log(`Finished grouping data by Date/Turma.`);

    // 5. Identify Rows for Deletion
    const rowsToDelete = [];
    for (const dateKey in groupedData) {
      for (const turmaName in groupedData[dateKey]) {
        const group = groupedData[dateKey][turmaName];
        const triggerCount = group.fixoCount + group.vagoBookedCount;

        if (triggerCount >= threshold) {
          Logger.log(`Threshold (${threshold}) MET for Turma "${turmaName}" on Date ${dateKey}. Trigger count: ${triggerCount}. Marking ${group.availableVagoRows.length} available Vago slots for deletion.`);
          // Add all row indices from this group to the master list
          rowsToDelete.push(...group.availableVagoRows);
        }
      }
    }

    // 6. Delete Rows (if any identified)
    if (rowsToDelete.length > 0) {
      Logger.log(`Preparing to delete ${rowsToDelete.length} rows...`);

      // Sort row indices in DESCENDING order
      rowsToDelete.sort((a, b) => b - a);

      // Delete rows from bottom to top
      let deletedCount = 0;
      for (const rowIndex of rowsToDelete) {
        try {
          instancesSheet.deleteRow(rowIndex);
          deletedCount++;
          // Logger.log(`Deleted row ${rowIndex}.`); // Can be very verbose
        } catch (e) {
          Logger.log(`ERROR deleting row ${rowIndex}: ${e.message}`);
          // Continue attempting to delete other rows even if one fails
        }
      }
      Logger.log(`Successfully deleted ${deletedCount} out of ${rowsToDelete.length} identified rows.`);
    } else {
      Logger.log('No available Vago slots needed removal based on threshold.');
    }

    Logger.log('*** cleanupExcessVagoSlots FINISHED ***');

  } catch (e) {
    Logger.log(`ERROR in cleanupExcessVagoSlots: ${e.message}\nStack: ${e.stack}`);
  } finally {
    releaseScriptLock_(lock); // Ensure lock release
  }
}

/**
 * [CLIENT CALLABLE - For Admin Cancel View] Gets future bookings with 'Agendada' status.
 * Returns combined data from Bookings and Instances.
 * @returns {string} JSON {success, message, data: [booking details]}
 */
function getCancellableBookings() {
  Logger.log('*** getCancellableBookings called ***');
  try {
    // Authorization: Ensure only Admin can call this
    const userEmail = getActiveUserEmail_();
    const userRole = getUserRolePlain_(userEmail);
    if (userRole !== USER_ROLES.ADMIN) {
      return createJsonResponse(false, 'Acesso negado. Apenas administradores podem visualizar esta lista.', null);
    }

    // Read data
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const { data: bookingData } = getSheetData_(SHEETS.BOOKING_DETAILS, HEADERS.BOOKING_DETAILS);
    const { data: instanceData } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES); // Need instance for context

    const today = new Date();
    today.setHours(0, 0, 0, 0);

    // Create a map of instances for quick lookup
    const instanceMap = instanceData.reduce((map, row) => {
      const idCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA;
      if (row && row.length > idCol) {
        const id = String(row[idCol] || '').trim();
        if (id) {
          map[id] = {
            date: formatValueToDate(row[HEADERS.SCHEDULE_INSTANCES.DATA]), // UTC Date obj
            time: formatValueToHHMM(row[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO], timeZone), // HH:mm string
            turma: String(row[HEADERS.SCHEDULE_INSTANCES.TURMA] || '').trim()
          };
        }
      }
      return map;
    }, {});

    // Filter and combine booking data
    const cancellableBookings = [];
    const bookingIdCol = HEADERS.BOOKING_DETAILS.ID_RESERVA;
    const instanceFkCol = HEADERS.BOOKING_DETAILS.ID_INSTANCIA;
    const statusCol = HEADERS.BOOKING_DETAILS.STATUS_RESERVA;
    const typeCol = HEADERS.BOOKING_DETAILS.TIPO_RESERVA;
    const discCol = HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL;
    const profRealCol = HEADERS.BOOKING_DETAILS.PROFESSOR_REAL;
    const profOrigCol = HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL;
    const creatorCol = HEADERS.BOOKING_DETAILS.CRIADO_POR;
    const reqCols = Math.max(bookingIdCol, instanceFkCol, statusCol, typeCol, discCol, profRealCol, profOrigCol, creatorCol) + 1;


    bookingData.forEach(row => {
      if (!row || row.length < reqCols) return; // Skip incomplete booking rows

      const bookingStatus = String(row[statusCol] || '').trim();
      const instanceId = String(row[instanceFkCol] || '').trim();
      const instanceInfo = instanceMap[instanceId]; // Get instance details from map

      // Filter: Must be 'Agendada', linked to a valid instance, and instance date must be today or in the future
      if (bookingStatus === 'Agendada' && instanceInfo && instanceInfo.date && instanceInfo.date >= today) {
        cancellableBookings.push({
          bookingId: String(row[bookingIdCol] || '').trim(),
          instanceId: instanceId,
          bookingType: String(row[typeCol] || '').trim(),
          date: Utilities.formatDate(instanceInfo.date, timeZone, 'dd/MM/yyyy'), // Format display date
          time: instanceInfo.time || 'N/D',
          turma: instanceInfo.turma || 'N/D',
          disciplina: String(row[discCol] || '').trim(),
          profReal: String(row[profRealCol] || '').trim(),
          profOrig: String(row[profOrigCol] || '').trim(),
          criadoPor: String(row[creatorCol] || '').trim()
        });
      }
    });

    // Sort by date, then time, then turma (optional but good for display)
    cancellableBookings.sort((a, b) => {
      const dateA = a.date.split('/').reverse().join(''); // YYYYMMDD
      const dateB = b.date.split('/').reverse().join('');
      if (dateA !== dateB) return dateA.localeCompare(dateB);
      if (a.time !== b.time) return a.time.localeCompare(b.time);
      return a.turma.localeCompare(b.turma);
    });

    Logger.log(`Found ${cancellableBookings.length} cancellable bookings.`);
    return createJsonResponse(true, `${cancellableBookings.length} reserva(s) encontrada(s).`, cancellableBookings);

  } catch (e) {
    return createJsonResponse(false, `Erro ao buscar reservas canceláveis: ${e.message}`, null);
  }
}

/**
 * [CLIENT CALLABLE - Admin Only] Cancels a booking and reverts the instance status.
 * @param {string} bookingIdToCancel - The ID of the booking to cancel (from Reservas Detalhadas).
 * @returns {string} JSON {success, message, data: {cancelledBookingId}}
 */
function cancelBookingAdmin(bookingIdToCancel) {
  Logger.log(`*** cancelBookingAdmin called for Booking ID: ${bookingIdToCancel} ***`);
  let lock = null;
  try {
    // 1. Authorization (Strict Admin Check)
    const userEmail = getActiveUserEmail_();
    const userRole = getUserRolePlain_(userEmail);
    if (userRole !== USER_ROLES.ADMIN) {
      throw new Error('Apenas administradores podem cancelar reservas.');
    }
    if (!bookingIdToCancel || typeof bookingIdToCancel !== 'string' || bookingIdToCancel.trim() === '') {
      throw new Error('ID da Reserva inválido ou ausente.');
    }
    const trimmedBookingId = bookingIdToCancel.trim();

    // 2. Acquire Lock
    lock = acquireScriptLock_();

    // 3. Find Booking and Corresponding Instance
    const { header: bookingHeader, data: bookingData, sheet: bookingsSheet } = getSheetData_(SHEETS.BOOKING_DETAILS, HEADERS.BOOKING_DETAILS);
    const { header: instanceHeader, sheet: instancesSheet } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES); // Need sheet object

    const bookingIdCol = HEADERS.BOOKING_DETAILS.ID_RESERVA;
    const instanceFkCol = HEADERS.BOOKING_DETAILS.ID_INSTANCIA;
    const bookingStatusCol = HEADERS.BOOKING_DETAILS.STATUS_RESERVA;
    const maxBookingIndex = Math.max(bookingIdCol, instanceFkCol, bookingStatusCol);

    let bookingRowIndex = -1;
    let bookingDetails = null;
    let instanceId = null;

    // Find the booking row by Booking ID
    for (let i = 0; i < bookingData.length; i++) {
      const row = bookingData[i];
      if (row && row.length > maxBookingIndex && String(row[bookingIdCol] || '').trim() === trimmedBookingId) {
        bookingRowIndex = i + 2; // 1-based sheet index
        bookingDetails = row;
        instanceId = String(row[instanceFkCol] || '').trim();
        break;
      }
    }

    if (bookingRowIndex === -1 || !bookingDetails || !instanceId) {
      throw new Error(`Reserva com ID ${trimmedBookingId} não encontrada.`);
    }
    Logger.log(`Booking ${trimmedBookingId} found at row ${bookingRowIndex}, linked to Instance ID ${instanceId}.`);

    // Check if booking is already cancelled
    const currentBookingStatus = String(bookingDetails[bookingStatusCol] || '').trim();
    if (currentBookingStatus !== 'Agendada') {
      throw new Error(`Esta reserva (ID ${trimmedBookingId}) já não está com status "Agendada" (Status atual: ${currentBookingStatus}). Não pode ser cancelada novamente.`);
    }

    // Find the instance row using the linked Instance ID
    const instanceIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA;
    const instanceRowFinder = instancesSheet.createTextFinder(instanceId).matchEntireCell(true);
    const foundInstanceCells = instanceRowFinder.findAll();
    if (foundInstanceCells.length === 0) {
      // Data inconsistency: Booking exists but instance doesn't
      Logger.log(`CRITICAL INCONSISTENCY: Booking ${trimmedBookingId} exists, but linked Instance ${instanceId} not found! Marking booking as cancelled, but cannot revert instance.`);
      // Mark booking as cancelled anyway
      bookingsSheet.getRange(bookingRowIndex, bookingStatusCol + 1).setValue('Cancelada (Instância Não Encontrada)');
      throw new Error(`Erro de dados: Instância ${instanceId} ligada a esta reserva não foi encontrada. Reserva marcada como cancelada, mas o horário pode não ter sido liberado.`);
    }
    const instanceRowIndex = foundInstanceCells[0].getRow();
    const instanceDetails = instancesSheet.getRange(instanceRowIndex, 1, 1, instanceHeader.length).getValues()[0];
    Logger.log(`Linked Instance ${instanceId} found at row ${instanceRowIndex}.`);

    // 4. Update Instance Sheet: Revert status, clear booking/event IDs
    const instanceStatusCol = HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO;
    const instanceBookingIdCol = HEADERS.SCHEDULE_INSTANCES.ID_RESERVA;
    const instanceEventIdCol = HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR;
    const existingEventId = String(instanceDetails[instanceEventIdCol] || '').trim(); // Get event ID *before* clearing

    const updatedInstanceRow = [...instanceDetails];
    updatedInstanceRow[instanceStatusCol] = STATUS_OCUPACAO.DISPONIVEL; // Revert status
    updatedInstanceRow[instanceBookingIdCol] = ''; // Clear booking ID link
    updatedInstanceRow[instanceEventIdCol] = ''; // Clear event ID link
    updateSheetRow_(instancesSheet, instanceRowIndex, instanceHeader.length, updatedInstanceRow);
    Logger.log(`Instance ${instanceId} status reverted to Disponivel, IDs cleared.`);

    // 5. Update Booking Sheet: Change status to 'Cancelada'
    bookingsSheet.getRange(bookingRowIndex, bookingStatusCol + 1).setValue('Cancelada'); // 1-based column
    Logger.log(`Booking ${trimmedBookingId} status updated to Cancelada.`);

    // 6. Delete Calendar Event (if exists)
    if (existingEventId) {
      Logger.log(`Attempting to delete Calendar event ID: ${existingEventId}`);
      try {
        const calendarIdConfig = getConfigValue('ID do Calendario');
        if (calendarIdConfig) {
          const calendar = CalendarApp.getCalendarById(calendarIdConfig.trim());
          if (calendar) {
            const event = calendar.getEventById(existingEventId);
            if (event) {
              event.deleteEvent();
              Logger.log(`Calendar event ${existingEventId} deleted successfully.`);
            } else {
              Logger.log(`Calendar event ${existingEventId} not found (already deleted?).`);
            }
          } else {
            Logger.log(`Calendar ID ${calendarIdConfig} not found/inaccessible, cannot delete event.`);
          }
        } else {
          Logger.log('Calendar ID not configured, cannot delete event.');
        }
      } catch (calError) {
        // Log error but don't fail the overall cancellation if calendar deletion fails
        Logger.log(`WARNING: Failed to delete Calendar event ${existingEventId}: ${calError.message}`);
      }
    }

    // 7. Send Cancellation Notification (Optional)
    // TODO: Implement if needed - gather relevant emails (creator, profs) and send notification.
    // sendCancellationEmail_(bookingDetails, instanceDetails, userEmail);
    createScheduleInstances()

    // 8. Release Lock and Return Success
    releaseScriptLock_(lock);
    return createJsonResponse(true, `Reserva ${trimmedBookingId} cancelada com sucesso.`, { cancelledBookingId: trimmedBookingId });

  } catch (e) {
    Logger.log(`ERROR in cancelBookingAdmin for ID ${bookingIdToCancel}: ${e.message}\nStack: ${e.stack}`);
    releaseScriptLock_(lock);
    return createJsonResponse(false, `Falha ao cancelar reserva: ${e.message}`, { failedBookingId: bookingIdToCancel });
  }
}

/**
 * Utility function to include HTML partials (CSS, JS) in the main HTML file.
 * Used as <?!= include('Stylesheet'); ?> in HTML.
 * @param {string} filename - The name of the HTML file (without .html extension).
 * @returns {string} The content of the file.
 */
function include(filename) {
  try {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  } catch (e) {
    Logger.log(`Error including file "${filename}": ${e.message}`);
    return `<!-- Error including file: ${filename}.html - ${e.message} -->`;
  }
}