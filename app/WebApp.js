/**
 * Arquivo: WebApp.gs
 * Descrição: Ponto de entrada da aplicação web (função doGet) e helpers para servir conteúdo HTML.
 */
function doGet(e) {
    let userEmail = '[Public/Unknown]';
    const page = e && e.parameter ? e.parameter.page : null;
    let userRole = null;
    try {
        userEmail = getActiveUserEmail_();
        userRole = getUserRolePlain_(userEmail);
    } catch (err) {
        Logger.log(`Attempted access, couldn't get user info initially (may be public): ${err.message}`);
    }
    Logger.log(`Web App GET request by: ${userEmail}. Page Parameter: ${page}. Detected Role: ${userRole}`);
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
    else if (page === 'cancel') {
        Logger.log(`Attempting to serve cancel view page for user: ${userEmail}`);
        if (userRole) {
            Logger.log(`Admin access granted for cancel page.`);
            try {
                const template = HtmlService.createTemplateFromFile('CancelView');
                return template.evaluate()
                    .setTitle('Cancelar Reservas')
                    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
            } catch (error) {
                Logger.log(`FATAL ERROR loading cancel view: ${error.message} Stack: ${error.stack}`);
                return HtmlService.createHtmlOutput('<h1>Erro Interno</h1><p>Erro ao carregar página de cancelamento.</p>').setTitle('Erro');
            }
        } else {
            Logger.log(`Access Denied for user ${userEmail} to cancel page. Role: ${userRole}`);
            return HtmlService
                .createHtmlOutput(`<h1>Acesso Negado</h1><p>Apenas administradores podem acessar esta página (${userEmail}).</p>`)
                .setTitle('Acesso Negado');
        }
    }
    else {
        Logger.log(`Attempting to serve main page for user: ${userEmail}`);
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
            Logger.log(`Access Denied for user ${userEmail} to main application. Role: ${userRole}`);
            return HtmlService
                .createHtmlOutput(`<h1>Acesso Negado</h1><p>Seu usuário (${userEmail}) não tem permissão para acessar esta aplicação. Entre em contato com o administrador.</p>`)
                .setTitle('Acesso Negado');
        }
    }
}
function getInitialData() {
    Logger.log('*** getInitialData called ***');
    let userEmail = '[Unavailable]';
    let userRole = null;
    let professors = [];
    let turmas = [];
    let disciplines = [];
    let errorMessages = [];
    try {
        userEmail = getActiveUserEmail_();
        userRole = getUserRolePlain_(userEmail);
        if (!userRole) {
            errorMessages.push('Usuário não encontrado ou não autorizado.');
        }
    } catch (e) {
        userEmail = '[Error]';
        errorMessages.push(`Erro ao obter informações do usuário: ${e.message}`);
    }
    try {
        const { data, header } = getSheetData_(SHEETS.AUTHORIZED_USERS, HEADERS.AUTHORIZED_USERS);
        const profSet = new Set();
        const nameCol = HEADERS.AUTHORIZED_USERS.NOME;
        const roleCol = HEADERS.AUTHORIZED_USERS.PAPEL;
        const requiredCols = Math.max(nameCol, roleCol) + 1;
        if (header.length >= requiredCols && data.length > 0) {
            data.forEach(row => {
                if (row && row.length >= requiredCols) {
                    const role = String(row[roleCol] || '').trim();
                    const name = String(row[nameCol] || '').trim();
                    if (role === USER_ROLES.PROFESSOR && name !== '') {
                        profSet.add(name);
                    }
                }
            });
        }
        professors = Array.from(profSet).sort((a, b) => a.localeCompare(b));
    } catch (e) {
        errorMessages.push(`Erro ao obter lista de professores: ${e.message}`);
    }
    try {
        const turmasConfig = getConfigValue('Turmas Disponiveis');
        if (turmasConfig !== null && turmasConfig !== '') {
            turmas = turmasConfig.split(',')
                .map(t => t.trim())
                .filter(t => t !== '')
                .sort((a, b) => a.localeCompare(b));
        }
    } catch (e) {
        errorMessages.push(`Erro ao obter lista de turmas: ${e.message}`);
    }
    try {
        const { data, header } = getSheetData_(SHEETS.DISCIPLINES, HEADERS.DISCIPLINES);
        const discSet = new Set();
        const nameCol = HEADERS.DISCIPLINES.NOME;
        if (header.length > nameCol && data.length > 0) {
            data.forEach(row => {
                if (row && row.length > nameCol) {
                    const name = String(row[nameCol] || '').trim();
                    if (name !== '') {
                        discSet.add(name);
                    }
                }
            });
        } else if (header.length <= nameCol) {
            Logger.log(`Discipline sheet "${SHEETS.DISCIPLINES}" does not have the required Name column (index ${nameCol}) or is empty.`);
        }
        disciplines = Array.from(discSet).sort((a, b) => a.localeCompare(b));
    } catch (e) {
        if (e.message && e.message.includes(`Planilha "${SHEETS.DISCIPLINES}" não encontrada`)) {
            Logger.log(`Optional sheet "${SHEETS.DISCIPLINES}" not found during initial load. Proceeding without disciplines.`);
            disciplines = [];
        } else {
            errorMessages.push(`Erro ao obter lista de disciplinas: ${e.message}`);
            Logger.log(`Error fetching disciplines: ${e.message}`);
            disciplines = [];
        }
    }
    const criticalFailure = errorMessages.some(msg => msg.includes('Usuário não encontrado') || msg.includes('Erro ao obter informações do usuário'));
    const overallSuccess = !criticalFailure;
    let message = '';
    if (overallSuccess && errorMessages.length > 0) {
        message = `Dados iniciais carregados com avisos: ${errorMessages.join('; ')}`;
    } else if (overallSuccess) {
        message = 'Dados iniciais carregados com sucesso.';
    } else {
        message = `Falha crítica ao carregar dados iniciais: ${errorMessages.join('; ')}`;
    }
    return createJsonResponse(overallSuccess, message, {
        user: { role: userRole, email: userEmail },
        professors: professors,
        turmas: turmas,
        disciplines: disciplines
    });
}
function include(filename) {
    try {
        return HtmlService.createHtmlOutputFromFile(filename).getContent();
    } catch (e) {
        Logger.log(`Error including file "${filename}": ${e.message}`);
        return `<!-- Error including file: ${filename}.html - ${e.message} -->`;
    }
}
// --- Funções Expostas para google.script.run ---
// As funções de nível superior (que não começam com '_') são automaticamente expostas para chamadas do frontend (google.script.run).
// getUserRole()
// getProfessorsList()
// getTurmasList()
// getDisciplinesList()
// getScheduleViewFilters()
// getFilteredScheduleInstances(turma, weekStartDateString)
// getAvailableSlots(tipoReserva)
// bookSlot(jsonBookingDetailsString)
// getCancellableBookings() // Apenas para Admin via WebApp
// cancelBookingAdmin(bookingIdToCancel) // Apenas para Admin via WebApp
// getPublicScheduleInstances(weekStartDateString) // Para a view pública
// createScheduleInstances() // Pode ser chamada via trigger ou manualmente, ou como parte de outra lógica (cancelamento)
// cleanOldScheduleInstances() // Pode ser chamada via trigger ou manualmente
// cleanupExcessVagoSlots() // Pode ser chamada via trigger ou manualmente, ou como parte de outra lógica (agendamento)
