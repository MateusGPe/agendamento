/**
 * Arquivo: Auth.gs
 * Descrição: Funções relacionadas a autenticação, autorização e listagem de usuários/entidades.
 */
function getUserRolePlain_(userEmail) {
    if (!userEmail) return null;
    const trimmedEmail = userEmail.trim().toLowerCase();
    if (trimmedEmail === '') return null;
    try {
        const userSheet = getSheetByName_(SHEETS.AUTHORIZED_USERS);
        const lastRow = userSheet.getLastRow();
        if (lastRow < 2) return null;
        const emailCol = HEADERS.AUTHORIZED_USERS.EMAIL + 1;
        const roleCol = HEADERS.AUTHORIZED_USERS.PAPEL + 1;
        const maxCol = Math.max(emailCol, roleCol);
        if (userSheet.getLastColumn() < maxCol) {
            Logger.log(`WARNING: Sheet "${SHEETS.AUTHORIZED_USERS}" has fewer columns (${userSheet.getLastColumn()}) than expected (${maxCol}). Role lookup might fail.`);
            return null;
        }
        const range = userSheet.getRange(2, 1, lastRow - 1, maxCol);
        const data = range.getValues();
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const emailInSheet = String(row[emailCol - 1] || '').trim().toLowerCase();
            if (emailInSheet === trimmedEmail) {
                const role = String(row[roleCol - 1] || '').trim();
                if (Object.values(USER_ROLES).includes(role)) {
                    return role;
                } else if (role !== '') {
                    Logger.log(`Invalid/Unrecognized role "${role}" found for user ${trimmedEmail}.`);
                }
                return null;
            }
        }
        return null;
    } catch (e) {
        Logger.log(`Error in getUserRolePlain_ for ${trimmedEmail}: ${e.message}`);
        return null;
    }
}
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
                if (row && row.length >= requiredCols) {
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
function getProfessorNameByEmail_(userEmail) {
    if (!userEmail || typeof userEmail !== 'string' || !userEmail.includes('@')) {
        Logger.log(`getProfessorNameByEmail_ called with invalid email: ${userEmail}`);
        return null;
    }
    const trimmedEmail = userEmail.trim().toLowerCase();
    try {
        const userSheet = getSheetByName_(SHEETS.AUTHORIZED_USERS);
        const lastRow = userSheet.getLastRow();
        if (lastRow < 2) return null;
        const emailCol = HEADERS.AUTHORIZED_USERS.EMAIL + 1;
        const nameCol = HEADERS.AUTHORIZED_USERS.NOME + 1;
        const roleCol = HEADERS.AUTHORIZED_USERS.PAPEL + 1;
        const maxCol = Math.max(emailCol, nameCol, roleCol);
        if (userSheet.getLastColumn() < maxCol) {
            Logger.log(`WARNING: Sheet "${SHEETS.AUTHORIZED_USERS}" has fewer columns (${userSheet.getLastColumn()}) than needed (${maxCol}) for name lookup.`);
            return null;
        }
        const range = userSheet.getRange(2, 1, lastRow - 1, maxCol);
        const data = range.getValues();
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const emailInSheet = String(row[emailCol - 1] || '').trim().toLowerCase();
            const roleInSheet = String(row[roleCol - 1] || '').trim();
            if (emailInSheet === trimmedEmail && roleInSheet === USER_ROLES.PROFESSOR) {
                const name = String(row[nameCol - 1] || '').trim();
                return name || null;
            }
        }
        Logger.log(`Professor name not found for email: ${trimmedEmail}`);
        return null;
    } catch (e) {
        Logger.log(`Error in getProfessorNameByEmail_ for ${trimmedEmail}: ${e.message}`);
        return null;
    }
}