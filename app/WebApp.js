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
