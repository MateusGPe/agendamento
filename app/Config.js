/**
 * Arquivo: Config.gs
 * Descrição: Funções para obter valores de configuração da planilha 'Configuracoes'.
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