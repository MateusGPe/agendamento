/**
 * Arquivo: Sheets.gs
 * Descrição: Funções de baixo nível para interação com as planilhas.
 */
const SHEET_CACHE_EXPIRATION_SECONDS = 300;
function getSheetByName_(sheetName) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
        Logger.log(`CRITICAL: Sheet "${sheetName}" not found.`);
        throw new Error(`Erro interno: Planilha "${sheetName}" não encontrada.`);
    }
    return sheet;
}
function getSheetData_(sheetName, headersDefinition = null) {
    const cache = CacheService.getScriptCache();
    const cacheKey = `SHEET_DATA_${SPREADSHEET_ID}_${sheetName}`;
    const cachedData = cache.get(cacheKey);
    if (cachedData != null) {
        try {
            const parsedData = JSON.parse(cachedData);
            if (parsedData && Array.isArray(parsedData.header) && Array.isArray(parsedData.data)) {
                Logger.log(`Cache HIT for sheet "${sheetName}". Returning cached data.`);
                const sheet = getSheetByName_(sheetName);
                return { header: parsedData.header, data: parsedData.data, sheet: sheet };
            } else {
                Logger.log(`Cache data for sheet "${sheetName}" seems invalid. Fetching fresh data.`);
                cache.remove(cacheKey);
            }
        } catch (e) {
            Logger.log(`Error parsing cached data for sheet "${sheetName}": ${e.message}. Fetching fresh data.`);
            cache.remove(cacheKey);
        }
    }
    Logger.log(`Cache MISS for sheet "${sheetName}". Reading from Sheet API.`);
    const sheet = getSheetByName_(sheetName);
    const range = sheet.getDataRange();
    const values = range.getValues();
    if (!values || values.length === 0) {
        Logger.log(`Sheet "${sheetName}" is completely empty.`);
        const emptyData = { header: [], data: [] };
        cache.put(cacheKey, JSON.stringify(emptyData), SHEET_CACHE_EXPIRATION_SECONDS);
        return { header: [], data: [], sheet: sheet };
    }
    const header = values[0];
    const data = values.length > 1 ? values.slice(1) : [];
    const dataToCache = { header: header, data: data };
    try {
        cache.put(cacheKey, JSON.stringify(dataToCache), SHEET_CACHE_EXPIRATION_SECONDS);
        Logger.log(`Stored fresh data for sheet "${sheetName}" in cache.`);
    } catch (e) {
        Logger.log(`Error putting data for sheet "${sheetName}" into cache: ${e.message}. Cache limit might be exceeded.`);
    }
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
function findRowIndexById_(data, idColumnIndex, targetId) {
    if (!targetId || typeof targetId !== 'string' || idColumnIndex < 0) return -1;
    const trimmedTargetId = targetId.trim();
    if (trimmedTargetId === '') return -1;
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        if (row && row.length > idColumnIndex) {
            const currentIdRaw = row[idColumnIndex];
            if (typeof currentIdRaw === 'string' || typeof currentIdRaw === 'number') {
                const currentId = String(currentIdRaw).trim();
                if (currentId === trimmedTargetId) {
                    return i + 2;
                }
            }
        }
    }
    return -1;
}
function updateSheetRow_(sheet, rowIndex, numCols, updatedRowData) {
    const cache = CacheService.getScriptCache();
    const cacheKey = `SHEET_DATA_${SPREADSHEET_ID}_${sheet.getName()}`;
    cache.remove(cacheKey);
    Logger.log(`Cache invalidated for sheet "${sheet.getName()}" due to update.`);
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
function appendSheetRow_(sheet, numCols, newRowData) {
    const cache = CacheService.getScriptCache();
    const cacheKey = `SHEET_DATA_${SPREADSHEET_ID}_${sheet.getName()}`;
    cache.remove(cacheKey);
    Logger.log(`Cache invalidated for sheet "${sheet.getName()}" due to append.`);
    try {
        if (!sheet || typeof sheet.appendRow !== 'function') throw new Error("Invalid sheet object provided for append.");
        if (numCols < 1) throw new Error(`Invalid number of columns: ${numCols}.`);
        if (!Array.isArray(newRowData)) throw new Error("newRowData must be an array.");
        const finalRowData = [...newRowData];
        while (finalRowData.length < numCols) finalRowData.push('');
        if (finalRowData.length > numCols) finalRowData.length = numCols;
        sheet.appendRow(finalRowData);
    } catch (e) {
        Logger.log(`ERROR appending row to sheet "${sheet.getName()}": ${e.message}`);
        throw new Error(`Erro interno ao adicionar dados na planilha "${sheet.getName()}": ${e.message}`);
    }
}