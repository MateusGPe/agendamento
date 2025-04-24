const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

const SHEETS = {
    CONFIG: 'Configuracoes',
    AUTHORIZED_USERS: 'Usuarios Autorizados',
    BASE_SCHEDULES: 'Horarios Base',
    SCHEDULE_INSTANCES: 'Instancias de Horarios',
    BOOKING_DETAILS: 'Reservas Detalhadas',
    DISCIPLINES: 'Disciplinas'
};

const HEADERS = {
    CONFIG: {
        NOME: 0, // Coluna A
        VALOR: 1 // Coluna B
    },
    AUTHORIZED_USERS: {
        EMAIL: 0, // Coluna A
        NOME: 1, // Coluna B
        PAPEL: 2 // Coluna C
    },
    BASE_SCHEDULES: {
        ID: 0, // Coluna A
        TIPO: 1, // Coluna B
        DIA_SEMANA: 2, // Coluna C
        HORA_INICIO: 3, // Coluna D
        DURACAO: 4, // Coluna E
        PROFESSOR_PRINCIPAL: 5, // Coluna F
        TURMA_PADRAO: 6, // Coluna G
        DISCIPLINA_PADRAO: 7, // Coluna H
        CAPACIDADE: 8, // Coluna I
        OBSERVATIONS: 9 // Coluna J
    },
    SCHEDULE_INSTANCES: {
        ID_INSTANCIA: 0, // Coluna A
        ID_BASE_HORARIO: 1, // Coluna B
        TURMA: 2, // Coluna C
        PROFESSOR_PRINCIPAL: 3, // Coluna D
        DATA: 4, // Coluna E
        DIA_SEMANA: 5, // Coluna F
        HORA_INICIO: 6, // Coluna G
        TIPO_ORIGINAL: 7, // Coluna H
        STATUS_OCUPACAO: 8, // Coluna I
        ID_RESERVA: 9, // Coluna J
        ID_EVENTO_CALENDAR: 10 // Coluna K
    },
    BOOKING_DETAILS: {
        ID_RESERVA: 0, // Coluna A
        TIPO_RESERVA: 1, // Coluna B
        ID_INSTANCIA: 2, // Coluna C
        PROFESSOR_REAL: 3, // Coluna D
        PROFESSOR_ORIGINAL: 4, // Coluna E
        ALUNOS: 5, // Coluna F
        TURMAS_AGENDADA: 6, // Coluna G
        DISCIPLINA_REAL: 7, // Coluna H
        DATA_HORA_INICIO_EFETIVA: 8, // Coluna I
        STATUS_RESERVA: 9, // Coluna J
        DATA_CRIACAO: 10, // Coluna K
        CRIADO_POR: 11 // Coluna L
    },
    DISCIPLINES: {
        NOME: 0 // Coluna A
    }
};
const STATUS_OCUPACAO = {
    DISPONIVEL: 'Disponivel',
    REPOSICAO_AGENDADA: 'Reposicao Agendada',
    SUBSTITUICAO_AGENDADA: 'Substituicao Agendada'
};
const TIPOS_RESERVA = {
    REPOSICAO: 'Reposicao',
    SUBSTITUICAO: 'Substituicao'
};
const TIPOS_HORARIO = {
    FIXO: 'Fixo',
    VAGO: 'Vago'
};

/**
 * Converte um valor de data lido de uma célula para um objeto Date válido.
 * Tenta lidar com objetos Date válidos e exclui a data de referência 1899-12-30 se não for uma data real.
 * @param {*} rawValue O valor lido diretamente da célula.
 * @returns {Date|null} Um objeto Date válido ou null se inválido.
 */
function formatValueToDate(rawValue) {
    if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
        if (rawValue.getFullYear() === 1899 && rawValue.getMonth() === 11 && rawValue.getDate() === 30) {
            if (rawValue.getHours() === 0 && rawValue.getMinutes() === 0 && rawValue.getSeconds() === 0) {
                return null;
            }
            return null;
        }
        return rawValue;
    }

    return null;
}


/**
 * Converte um valor de data/hora lido de uma célula para uma string HH:mm.
 * Tenta lidar com objetos Date (extraindo HH:mm no fuso horário da planilha), strings "HH:mm" ou números seriais de hora.
 * @param {*} rawValue O valor lido diretamente da célula.
 * @param {string} timeZone Fuso horário da planilha (SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone()).
 * @returns {string|null} A hora formatada como "HH:mm" ou null se inválido.
 */
function formatValueToHHMM(rawValue, timeZone) {
    if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
        return Utilities.formatDate(rawValue, timeZone, 'HH:mm');
    } else if (typeof rawValue === 'string') {
        const timeMatch = rawValue.trim().match(/^(\d{1,2}):(\d{2})(:\d{2})?$/);
        if (timeMatch) {
            const hour = parseInt(timeMatch[1], 10);
            const minute = parseInt(timeMatch[2], 10);
            if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
            }
        }
    } else if (typeof rawValue === 'number' && rawValue >= 0 && rawValue <= 1) {
        const totalMinutes = Math.round(rawValue * 24 * 60);
        const hours = Math.floor(totalMinutes / 60);
        const minutes = totalMinutes % 60;
        if (hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59) {
            return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
        }
    }
    return null;
}


/**
 * Obtém o papel do usuário logado.
 * Versão interna que retorna apenas a string do papel ou null.
 * @param {string} userEmail O email do usuário logado.
 * @returns {string|null} O papel do usuário (Admin, Professor, Aluno) ou null se não autorizado.
 */
function getUserRolePlain(userEmail) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AUTHORIZED_USERS);
    if (!sheet) {
        Logger.log(`Planilha "${SHEETS.AUTHORIZED_USERS}" não encontrada em getUserRolePlain.`);
        return null;
    }
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
        Logger.log(`Planilha "${SHEETS.AUTHORIZED_USERS}" vazia ou apenas cabeçalho.`);
        return null;
    }

    for (let i = 1; i < data.length; i++) {
        // Adicionado check para garantir que a linha, coluna e email não estão vazios/inválidos
        if (data[i] && data[i].length > HEADERS.AUTHORIZED_USERS.PAPEL && data[i][HEADERS.AUTHORIZED_USERS.EMAIL] && typeof data[i][HEADERS.AUTHORIZED_USERS.EMAIL] === 'string' && data[i][HEADERS.AUTHORIZED_USERS.EMAIL].trim() !== '' && data[i][HEADERS.AUTHORIZED_USERS.EMAIL].trim() === userEmail) {
            const role = data[i][HEADERS.AUTHORIZED_USERS.PAPEL];
            // Valida se o papel lido é um dos esperados
            if (['Admin', 'Professor', 'Aluno'].includes(role)) {
                return role;
            } else {
                Logger.log(`Papel inválido encontrado para o usuário ${userEmail} na linha ${i + 1} da planilha "${SHEETS.AUTHORIZED_USERS}": "${role}".`);
                // Continua procurando outros papéis ou retorna null, dependendo da política
            }
        }
    }
    Logger.log(`Usuário "${userEmail}" não encontrado na lista de autorizados da planilha "${SHEETS.AUTHORIZED_USERS}".`);
    return null;
}

/**
 * Função auxiliar para ler um valor da planilha Configurações.
 * @param {string} configName O nome da configuração a buscar.
 * @returns {string|null} O valor da configuração (como string) ou null se não encontrado.
 */
function getConfigValue(configName) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
    if (!sheet) {
        Logger.log(`Planilha "${SHEETS.CONFIG}" não encontrada.`);
        return null;
    }
    // Use getDataRange() para pegar tudo
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
        Logger.log(`Planilha "${SHEETS.CONFIG}" vazia ou apenas cabeçalho.`);
        return null;
    }

    for (let i = 1; i < data.length; i++) { // Começa da linha 2 (índice 1)
        if (data[i] && data[i].length > HEADERS.CONFIG.VALOR && data[i][HEADERS.CONFIG.NOME] === configName) {
            // Retorna o valor como string, tratando null/undefined e formatando para string
            // Converte para string explicitamente para garantir que funciona com números ou outros tipos
            return String(data[i][HEADERS.CONFIG.VALOR] || '').trim();
        }
    }
    Logger.log(`Configuração "${configName}" não encontrada na planilha "${SHEETS.CONFIG}".`);
    return null;
}


// --- Funções do Web App ---

/**
 * Função principal para servir o Web App.
 * Verifica a autorização do usuário antes de exibir a interface.
 */
function doGet(e) {
    const userEmail = Session.getActiveUser().getEmail();
    // Use a versão interna para verificar o papel
    const userRole = getUserRolePlain(userEmail);

    if (!userRole) {
        // Usuário não encontrado na lista de autorizados
        return HtmlService.createHtmlOutput(
            '<h1>Acesso Negado</h1>' +
            '<p>Seu usuário (' + userEmail + ') não tem permissão para acessar esta aplicação. Entre em contato com o administrador.</p>'
        ).setTitle('Acesso Negado');
    }

    // Usuário autorizado, servir a interface principal
    const htmlOutput = HtmlService.createTemplateFromFile('Index');
    htmlOutput.userRole = userRole; // Passa o papel do usuário para o frontend
    return htmlOutput.evaluate()
        .setTitle('Sistema de Agendamento');
}

/**
 * Inclui um arquivo HTML/CSS/JS no HTML principal.
 * Usado dentro dos templates HTML.
 */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// --- Funções de Backend (Chamadas pelo Frontend via google.script.run) ---

/**
 * Obtém o papel do usuário logado E seu email.
 * Esta versão retorna um objeto { success, message, data: { role: string|null, email: string|null } } SERIALIZADO como JSON.
 * @returns {string} Uma string JSON representando o resultado.
 */
function getUserRole() {
    Logger.log('*** getUserRole chamada ***');
    const userEmail = Session.getActiveUser().getEmail();
    Logger.log('Tentando obter papel para email: ' + userEmail);

    // Chama a versão interna para obter apenas o papel
    const userRole = getUserRolePlain(userEmail);

    // Retorna o papel E o email logado
    return JSON.stringify({
        success: true,
        message: userRole ? 'Papel do usuário obtido.' : 'Usuário não encontrado ou não autorizado.',
        data: {
            role: userRole,
            email: userEmail
        }
    });
}

/**
 * Obtém uma lista de professores da planilha Usuarios Autorizados.
 * Retorna um objeto { success, message, data: Array<string> } SERIALIZADO como JSON.
 * @returns {string} Uma string JSON representando o resultado.
 */
function getProfessorsList() {
    Logger.log('*** getProfessorsList chamada ***');
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AUTHORIZED_USERS);
        if (!sheet) {
            Logger.log(`Erro: Planilha "${SHEETS.AUTHORIZED_USERS}" não encontrada em getProfessorsList.`);
            return JSON.stringify({ success: false, message: `Erro interno: Planilha "${SHEETS.AUTHORIZED_USERS}" não encontrada.`, data: [] });
        }

        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) {
            Logger.log(`Planilha "${SHEETS.AUTHORIZED_USERS}" vazia ou apenas cabeçalho.`);
            return JSON.stringify({ success: true, message: 'Nenhum usuário autorizado encontrado.', data: [] });
        }

        const professors = [];
        // Começa da segunda linha (índice 1) para pular o cabeçalho
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            // Verifica se a linha tem colunas suficientes e se o papel é Professor
            if (row && row.length > HEADERS.AUTHORIZED_USERS.PAPEL) {
                const userRole = String(row[HEADERS.AUTHORIZED_USERS.PAPEL] || '').trim();
                const userName = String(row[HEADERS.AUTHORIZED_USERS.NOME] || '').trim();

                if (userRole === 'Professor' && userName !== '') {
                    professors.push(userName); // Adiciona o nome do professor
                }
            }
        }

        // Ordena a lista de professores alfabeticamente
        professors.sort();

        Logger.log(`Encontrados ${professors.length} professores.`);
        return JSON.stringify({ success: true, message: 'Lista de professores obtida com sucesso.', data: professors });

    } catch (e) {
        Logger.log('Erro em getProfessorsList: ' + e.message + ' Stack: ' + e.stack);
        return JSON.stringify({ success: false, message: 'Ocorreu um erro ao obter a lista de professores: ' + e.message, data: [] });
    }
}


/**
 * Obtém uma lista de turmas da configuração.
 * (Ainda pode ser útil para outras partes do sistema ou futuras funcionalidades).
 * Retorna um objeto { success, message, data: Array<string> } SERIALIZADO como JSON.
 * @returns {string} Uma string JSON representando o resultado.
 */
function getTurmasList() {
    Logger.log('*** getTurmasList chamada ***');
    try {
        const turmasConfig = getConfigValue('Turmas Disponiveis');
        if (!turmasConfig || turmasConfig === '') {
            Logger.log("Configuração 'Turmas Disponiveis' não encontrada ou vazia.");
            return JSON.stringify({ success: true, message: "Configuração de turmas não encontrada ou vazia.", data: [] });
        }
        const turmasArray = turmasConfig.split(',').map(t => t.trim()).filter(t => t !== ''); // Divide por vírgula, remove espaços e vazios
        turmasArray.sort(); // Ordena

        Logger.log(`Encontradas ${turmasArray.length} turmas na configuração.`);
        return JSON.stringify({ success: true, message: 'Lista de turmas (config) obtida.', data: turmasArray });

    } catch (e) {
        Logger.log('Erro em getTurmasList: ' + e.message + ' Stack: ' + e.stack);
        return JSON.stringify({ success: false, message: 'Ocorreu um erro ao obter a lista de turmas: ' + e.message, data: [] });
    }
}


/**
 * Obtém uma lista de disciplinas da planilha dedicada 'Disciplinas'.
 * Retorna um objeto { success, message, data: Array<string> } SERIALIZADO como JSON.
 * @returns {string} Uma string JSON representando o resultado.
 */
function getDisciplinesList() {
    Logger.log(`*** getDisciplinesList chamada (lendo da planilha dedicada: ${SHEETS.DISCIPLINES}) ***`);
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const disciplinesSheet = ss.getSheetByName(SHEETS.DISCIPLINES); // Usa a nova constante

        // Verifica se a planilha existe
        if (!disciplinesSheet) {
            Logger.log(`Erro: Planilha "${SHEETS.DISCIPLINES}" não encontrada.`);
            return JSON.stringify({ success: false, message: `Erro interno: Planilha de Disciplinas "${SHEETS.DISCIPLINES}" não encontrada. Verifique o nome da aba.`, data: [] });
        }

        // Lê os dados da planilha
        const disciplinesData = disciplinesSheet.getDataRange().getValues();

        // Verifica se há dados (mais do que apenas a linha do cabeçalho)
        if (disciplinesData.length <= 1) {
            Logger.log(`Planilha "${SHEETS.DISCIPLINES}" vazia ou apenas cabeçalho.`);
            // Retorna sucesso, mas com lista vazia e mensagem informativa
            return JSON.stringify({ success: true, message: `Nenhuma disciplina cadastrada na planilha "${SHEETS.DISCIPLINES}".`, data: [] });
        }

        const disciplinesArray = [];
        // Itera pelas linhas, começando da segunda (índice 1) para pular o cabeçalho
        for (let i = 1; i < disciplinesData.length; i++) {
            const row = disciplinesData[i];
            // Verifica se a linha existe e tem a coluna do nome
            if (row && row.length > HEADERS.DISCIPLINES.NOME) {
                // Extrai o nome da disciplina da coluna definida em HEADERS
                const disciplineName = String(row[HEADERS.DISCIPLINES.NOME] || '').trim();
                // Adiciona à lista apenas se não for uma string vazia
                if (disciplineName !== '') {
                    disciplinesArray.push(disciplineName);
                }
            }
        }

        // Ordena a lista alfabeticamente
        disciplinesArray.sort();

        Logger.log(`Encontradas ${disciplinesArray.length} disciplinas na planilha "${SHEETS.DISCIPLINES}".`);
        // Retorna a lista no formato JSON padrão
        return JSON.stringify({ success: true, message: 'Lista de disciplinas obtida com sucesso.', data: disciplinesArray });

    } catch (e) {
        // Captura e loga erros inesperados
        Logger.log(`Erro em getDisciplinesList (planilha dedicada "${SHEETS.DISCIPLINES}"): ${e.message} Stack: ${e.stack}`);
        return JSON.stringify({ success: false, message: `Ocorreu um erro ao obter a lista de disciplinas da planilha "${SHEETS.DISCIPLINES}": ${e.message}`, data: [] });
    }
}


/**
 * Busca slots de horários disponíveis para agendamento.
 * Robusto na leitura de Data, Hora, Turma e Professor Principal (para fixos) da planilha Instancias de Horarios.
 * Retorna um objeto { success: boolean, message: string, data: Array<Object>|null } SERIALIZADO como JSON.
 * @param {string} tipoReserva O tipo de reserva ('Reposicao' ou 'Substituicao').
 * @returns {string} Uma string JSON representando o resultado.
 */
function getAvailableSlots(tipoReserva) {
    Logger.log('*** getAvailableSlots chamada para tipo: ' + tipoReserva + ' ***');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const timeZone = ss.getSpreadsheetTimeZone();

    const userEmail = Session.getActiveUser().getEmail();
    const userRole = getUserRolePlain(userEmail);

    if (!userRole) {
        Logger.log('Erro: Usuário no autorizado.');
        return JSON.stringify({ success: false, message: 'Usuário não autorizado a buscar horários.', data: null });
    }

    try {
        const sheet = ss.getSheetByName(SHEETS.SCHEDULE_INSTANCES);
        if (!sheet) {
            Logger.log(`Erro: Planilha "${SHEETS.SCHEDULE_INSTANCES}" não encontrada!`);
            return JSON.stringify({ success: false, message: `Erro interno: Planilha "${SHEETS.SCHEDULE_INSTANCES}" não encontrada.`, data: null });
        }
        // Lê todas as colunas necessárias para ter todos os dados
        const minCols = Math.max(
            HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA,
            HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO,
            HEADERS.SCHEDULE_INSTANCES.TURMA,
            HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL,
            HEADERS.SCHEDULE_INSTANCES.DATA,
            HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA,
            HEADERS.SCHEDULE_INSTANCES.HORA_INICIO,
            HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL,
            HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO,
            HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR
        ) + 1;

        const rawData = sheet.getDataRange().getValues();
        if (rawData.length <= 1) {
            Logger.log('Planilha Instancias de Horarios está vazia ou apenas cabeçalho.');
            return JSON.stringify({ success: true, message: 'Nenhuma instância de horário futuro encontrada. Gere instâncias primeiro.', data: [] });
        }
        // Verifica se a planilha tem o número mínimo de colunas esperado
        if (rawData[0].length < minCols) {
            Logger.log(`Warning: Planilha "${SHEETS.SCHEDULE_INSTANCES}" tem menos colunas (${rawData[0].length}) do que o esperado (${minCols}). A leitura de dados pode falhar.`);
            // Continua, mas pode pular linhas ou lançar erro ao acessar índices
        }


        const data = rawData.slice(1); // Pula o cabeçalho

        const availableSlots = [];
        const today = new Date();
        today.setHours(0, 0, 0, 0); // Para comparar apenas a data

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const rowIndex = i + 2; // Linha no Sheets (baseado em 1)

            // --- Leitura e Tratamento Robusto de Colunas ---
            // Verifica se a linha tem colunas suficientes antes de acessar os índices
            if (!row || row.length < minCols) {
                Logger.log(`Skipping incomplete row ${rowIndex} in Instancias de Horarios. Missing required columns.`);
                continue;
            }

            const instanceIdRaw = row[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA];
            const baseIdRaw = row[HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO];
            const turmaRaw = row[HEADERS.SCHEDULE_INSTANCES.TURMA];
            const professorPrincipalRaw = row[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL];
            const rawDate = row[HEADERS.SCHEDULE_INSTANCES.DATA];
            const instanceDiaSemanaRaw = row[HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA];
            const rawHoraInicio = row[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO];
            const originalTypeRaw = row[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL];
            const instanceStatusRaw = row[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO];


            // Valida e formata os campos essenciais
            const instanceId = (typeof instanceIdRaw === 'string' || typeof instanceIdRaw === 'number') ? String(instanceIdRaw).trim() : null;
            const baseId = (typeof baseIdRaw === 'string' || typeof baseIdRaw === 'number') ? String(baseIdRaw).trim() : null;
            const turma = (typeof turmaRaw === 'string' || typeof turmaRaw === 'number') ? String(turmaRaw).trim() : null;
            const professorPrincipal = (typeof professorPrincipalRaw === 'string' || typeof professorPrincipalRaw === 'number') ? String(professorPrincipalRaw || '').trim() : ''; // Professor Principal (pode ser vazio para Vago)
            const instanceDate = formatValueToDate(rawDate); // Usa função auxiliar para Data
            const instanceDiaSemana = (typeof instanceDiaSemanaRaw === 'string' || typeof instanceDiaSemanaRaw === 'number') ? String(instanceDiaSemanaRaw).trim() : null;
            const formattedHoraInicio = formatValueToHHMM(rawHoraInicio, timeZone); // Usa função auxiliar para Hora
            const originalType = (typeof originalTypeRaw === 'string' || typeof originalTypeRaw === 'number') ? String(originalTypeRaw).trim() : null;
            const instanceStatus = (typeof instanceStatusRaw === 'string' || typeof instanceStatusRaw === 'number') ? String(instanceStatusRaw).trim() : null;


            // Validação de dados essenciais formatados
            if (!instanceId || instanceId === '' || !baseId || baseId === '' || !turma || turma === '' ||
                !instanceDate || // Data válida
                !instanceDiaSemana || instanceDiaSemana === '' ||
                formattedHoraInicio === null || // Hora válida
                !originalType || originalType === '' ||
                !instanceStatus || instanceStatus === '') {
                Logger.log(`Skipping row ${rowIndex} due to invalid or missing essential data after formatting: ID=${instanceIdRaw}, BaseID=${baseIdRaw}, Turma=${turmaRaw}, ProfPrinc=${professorPrincipalRaw}, Data=${rawDate}, Dia=${instanceDiaSemanaRaw}, Hora=${rawHoraInicio}, Tipo=${originalTypeRaw}, Status=${instanceStatusRaw}`);
                continue; // Pula a linha se faltarem dados críticos ou inválidos
            }

            // Valida os valores lidos se estão nos domínios esperados
            const daysOfWeek = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
            if (!daysOfWeek.includes(instanceDiaSemana)) {
                Logger.log(`Skipping row ${rowIndex} due to invalid Dia da Semana: "${instanceDiaSemana}". Raw: ${instanceDiaSemanaRaw}`);
                continue;
            }
            if (originalType !== TIPOS_HORARIO.FIXO && originalType !== TIPOS_HORARIO.VAGO) {
                Logger.log(`Skipping row ${rowIndex} due to invalid Tipo Original: "${originalType}". Raw: ${originalTypeRaw}`);
                continue;
            }
            // Permite qualquer status válido para verificação de regras de agendamento
            if (instanceStatus !== STATUS_OCUPACAO.DISPONIVEL && instanceStatus !== STATUS_OCUPACAO.REPOSICAO_AGENDADA && instanceStatus !== STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA) {
                Logger.log(`Skipping row ${rowIndex} due to invalid Status Ocupação: "${instanceStatus}". Raw: ${instanceStatusRaw}`);
                continue;
            }

            // Ignora horários no passado
            // Cria um objeto Date "apenas com data" para comparação
            const dateOnly = new Date(instanceDate.getFullYear(), instanceDate.getMonth(), instanceDate.getDate());
            if (dateOnly < today) {
                // Logger.log(`Skipping row ${rowIndex} in the past: ${Utilities.formatDate(instanceDate, timeZone, 'yyyy-MM-dd')}`);
                continue;
            }


            // --- Filtra com base no tipo de reserva e regras ---
            if (tipoReserva === TIPOS_RESERVA.REPOSICAO) {
                // Reposição só em horários VAGOS e DISPONÍVEIS
                if (originalType === TIPOS_HORARIO.VAGO && instanceStatus === STATUS_OCUPACAO.DISPONIVEL) {
                    availableSlots.push({
                        idInstancia: instanceId,
                        baseId: baseId,
                        turma: turma,
                        professorPrincipal: professorPrincipal, // Inclui Professor Principal (será vazio para Vago)
                        data: Utilities.formatDate(instanceDate, timeZone, 'dd/MM/yyyy'), // Data formatada para STRING (dd/MM/yyyy)
                        diaSemana: instanceDiaSemana,
                        horaInicio: formattedHoraInicio, // Hora formatada para STRING (HH:mm)
                        tipoOriginal: originalType,
                        statusOcupacao: instanceStatus,
                    });
                }
            } else if (tipoReserva === TIPOS_RESERVA.SUBSTITUICAO) {
                // Substituição só em horários FIXOS
                // E NÃO podem ser em horários que já são Reposições agendadas
                if (originalType === TIPOS_HORARIO.FIXO && instanceStatus !== STATUS_OCUPACAO.REPOSICAO_AGENDADA) {
                    availableSlots.push({
                        idInstancia: instanceId,
                        baseId: baseId,
                        turma: turma,
                        professorPrincipal: professorPrincipal, // Inclui Professor Principal
                        data: Utilities.formatDate(instanceDate, timeZone, 'dd/MM/yyyy'), // Data formatada para STRING (dd/MM/yyyy)
                        diaSemana: instanceDiaSemana,
                        horaInicio: formattedHoraInicio, // Hora formatada para STRING (HH:mm)
                        tipoOriginal: originalType,
                        statusOcupacao: instanceStatus,
                    });
                }
            }
        }

        Logger.log('Número de slots disponíveis encontrados: ' + availableSlots.length);

        // Retorna sucesso com os dados, SERIALIZADO como JSON
        return JSON.stringify({ success: true, message: 'Slots carregados com sucesso.', data: availableSlots });

    } catch (e) {
        Logger.log('Erro no getAvailableSlots: ' + e.message + ' Stack: ' + e.stack);
        // Retorna falha em caso de exceção, SERIALIZADO como JSON
        return JSON.stringify({ success: false, message: 'Ocorreu um erro interno ao buscar horários: ' + e.message, data: null });
    }
}


/**
 * Agenda uma reposição ou substituição.
 * Robusto na leitura de Data e Hora da instância para criar o evento do Calendar.
 * Recebe os detalhes da reserva como uma string JSON (SEM Turma Agendada).
 * Professor Original para Substituição é lido da instância, não do frontend.
 * A Turma Agendada é lida da instância, não do frontend.
 * Retorna um objeto { success: boolean, message: string } SERIALIZADO como JSON.
 * @param {string} jsonBookingDetailsString String JSON com os detalhes da reserva.
 * @returns {string} Uma string JSON representando o resultado.
 */
function bookSlot(jsonBookingDetailsString) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000); // Espera no máximo 10 segundos

    const userEmail = Session.getActiveUser().getEmail();
    const userRole = getUserRolePlain(userEmail);

    if (!userRole) {
        lock.releaseLock();
        return JSON.stringify({ success: false, message: 'Usuário não autorizado a agendar.', data: null });
    }

    let bookingDetails;
    try {
        bookingDetails = JSON.parse(jsonBookingDetailsString); // PARSEA a string JSON recebida
        Logger.log("Booking details received and parsed: " + JSON.stringify(bookingDetails));
    } catch (e) {
        lock.releaseLock();
        Logger.log('Erro ao parsear JSON de detalhes da reserva: ' + e.message);
        return JSON.stringify({ success: false, message: 'Erro ao processar dados da reserva.', data: null });
    }


    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const instancesSheet = ss.getSheetByName(SHEETS.SCHEDULE_INSTANCES);
        const bookingsSheet = ss.getSheetByName(SHEETS.BOOKING_DETAILS);
        //const configSheet = ss.getSheetByName(SHEETS.CONFIG);
        const timeZone = ss.getSpreadsheetTimeZone();


        // Validar entrada básica
        if (!bookingDetails || typeof bookingDetails.idInstancia !== 'string' || bookingDetails.idInstancia.trim() === '') {
            lock.releaseLock();
            return JSON.stringify({ success: false, message: 'Dados de ID da instância de horário incompletos ou inválidos.', data: null });
        }

        const instanceIdToBook = bookingDetails.idInstancia.trim();
        const bookingType = bookingDetails.tipoReserva ? String(bookingDetails.tipoReserva).trim() : null;
        if (!bookingType || (bookingType !== TIPOS_RESERVA.REPOSICAO && bookingType !== TIPOS_RESERVA.SUBSTITUICAO)) {
            lock.releaseLock();
            return JSON.stringify({ success: false, message: 'Tipo de reserva inválido ou ausente.', data: null });
        }


        // --- Encontrar a linha da instância na planilha Instancias de Horarios ---
        const instanceDataRaw = instancesSheet.getDataRange().getValues();
        let instanceRowIndex = -1;
        let instanceDetails = null;

        if (instanceDataRaw.length <= 1) {
            lock.releaseLock();
            return JSON.stringify({ success: false, message: 'Erro interno: Planilha de instâncias vazia ou estrutura incorreta.', data: null });
        }

        const instanceData = instanceDataRaw.slice(1);

        for (let i = 0; i < instanceData.length; i++) {
            const row = instanceData[i];
            const rowIndex = i + 2;
            const minColsForId = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA + 1;
            if (row && row.length >= minColsForId) {
                const currentInstanceIdRaw = row[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA];
                const currentInstanceId = (typeof currentInstanceIdRaw === 'string' || typeof currentInstanceIdRaw === 'number') ? String(currentInstanceIdRaw).trim() : null;
                if (currentInstanceId && currentInstanceId === instanceIdToBook) {
                    instanceRowIndex = rowIndex;
                    instanceDetails = row;
                    break;
                }
            }
        }

        if (instanceRowIndex === -1 || !instanceDetails) {
            lock.releaseLock();
            return JSON.stringify({ success: false, message: 'Este horário não está mais disponível. Por favor, atualize a lista e tente novamente.', data: null });
        }

        // --- Obter detalhes da instância com robustez para validação e Calendar ---
        const expectedInstanceCols = Math.max(
            HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO,
            HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL,
            HEADERS.SCHEDULE_INSTANCES.DATA,
            HEADERS.SCHEDULE_INSTANCES.HORA_INICIO,
            HEADERS.SCHEDULE_INSTANCES.TURMA, // Coluna da turma
            HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL,
            HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR
        ) + 1;

        if (instanceDetails.length < expectedInstanceCols) {
            Logger.log(`Erro: Linha ${instanceRowIndex} na planilha Instancias de Horarios tem menos colunas (${instanceDetails.length}) que o esperado (${expectedInstanceCols}).`);
            lock.releaseLock();
            return JSON.stringify({ success: false, message: "Erro interno: Dados do horário selecionado incompletos na planilha.", data: null });
        }

        const currentStatusRaw = instanceDetails[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO];
        const originalTypeRaw = instanceDetails[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL];
        const rawBookingDate = instanceDetails[HEADERS.SCHEDULE_INSTANCES.DATA];
        const rawBookingTime = instanceDetails[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO];
        const turmaInstanciaRaw = instanceDetails[HEADERS.SCHEDULE_INSTANCES.TURMA]; // Turma da instância
        const professorPrincipalInstanciaRaw = instanceDetails[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL];
        const calendarEventIdExistingRaw = instanceDetails[HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR];


        const currentStatus = (typeof currentStatusRaw === 'string' || typeof currentStatusRaw === 'number') ? String(currentStatusRaw).trim() : null;
        const originalType = (typeof originalTypeRaw === 'string' || typeof originalTypeRaw === 'number') ? String(originalTypeRaw).trim() : null;
        const turmaInstancia = (typeof turmaInstanciaRaw === 'string' || typeof turmaInstanciaRaw === 'number') ? String(turmaInstanciaRaw).trim() : null; // Turma da instância
        const professorPrincipalInstancia = (typeof professorPrincipalInstanciaRaw === 'string' || typeof professorPrincipalInstanciaRaw === 'number') ? String(professorPrincipalInstanciaRaw || '').trim() : '';
        const bookingDateObj = formatValueToDate(rawBookingDate);
        const bookingHourString = formatValueToHHMM(rawBookingTime, timeZone);
        const calendarEventIdExisting = (typeof calendarEventIdExistingRaw === 'string' || typeof calendarEventIdExistingRaw === 'number') ? String(calendarEventIdExistingRaw || '').trim() : null;


        // Re-Validação dos dados críticos lidos da instância (garante consistência)
        if (!currentStatus || currentStatus === '' ||
            !originalType || originalType === '' ||
            !turmaInstancia || turmaInstancia === '' || // Turma da instância é crucial
            !bookingDateObj || bookingHourString === null) {
            Logger.log(`Erro: Dados críticos da instância ${instanceIdToBook} na linha ${instanceRowIndex} são inválidos. Status=${currentStatusRaw}, Tipo=${originalTypeRaw}, Turma=${turmaInstanciaRaw}, Data=${rawBookingDate}, Hora=${rawBookingTime}`);
            lock.releaseLock();
            return JSON.stringify({ success: false, message: "Erro interno: Dados do horário selecionado são inválidos na planilha.", data: null });
        }


        // Re-Validação final baseada nas regras e status atual (para garantir que não houve concorrência)
        if (bookingType === TIPOS_RESERVA.REPOSICAO) {
            if (originalType !== TIPOS_HORARIO.VAGO || currentStatus !== STATUS_OCUPACAO.DISPONIVEL) {
                lock.releaseLock();
                return JSON.stringify({ success: false, message: 'Este horário não está mais disponível para reposição ou não é um horário vago (concorrência).', data: null });
            }
            // Validação de campos específicos para Reposição recebidos do frontend (SEM Turma)
            if (!bookingDetails.professorReal || bookingDetails.professorReal.trim() === '' ||
                !bookingDetails.disciplinaReal || bookingDetails.disciplinaReal.trim() === '') {
                lock.releaseLock();
                return JSON.stringify({ success: false, message: 'Por favor, preencha todos os campos obrigatórios para reposição (Professor, Disciplina).', data: null });
            }

        } else if (bookingType === TIPOS_RESERVA.SUBSTITUICAO) {
            if (originalType !== TIPOS_HORARIO.FIXO) {
                lock.releaseLock();
                return JSON.stringify({ success: false, message: 'Este horário não é um horário fixo e não pode ser substituído.', data: null });
            }
            if (currentStatus === STATUS_OCUPACAO.REPOSICAO_AGENDADA) {
                lock.releaseLock();
                return JSON.stringify({ success: false, message: 'Este horário fixo está sendo usado para uma reposição e não pode ser substituído.', data: null });
            }
            if (currentStatus !== STATUS_OCUPACAO.DISPONIVEL && currentStatus !== STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA) {
                lock.releaseLock();
                return JSON.stringify({ success: false, message: 'Este horário fixo não está disponível para substituição neste momento (concorrência).', data: null });
            }
            // Validação de campos específicos para Substituição recebidos do frontend (SEM Turma)
            if (!bookingDetails.professorReal || bookingDetails.professorReal.trim() === '' ||
                !bookingDetails.disciplinaReal || bookingDetails.disciplinaReal.trim() === '') {
                lock.releaseLock();
                return JSON.stringify({ success: false, message: 'Por favor, preencha todos os campos obrigatórios para substituição (Professor Substituto, Disciplina).', data: null });
            }
            // Adicionalmente, para substituição, o Professor Principal da instância deve ser informado
            if (professorPrincipalInstancia === '') {
                Logger.log(`Erro: Instância de horário fixo ${instanceIdToBook} na linha ${instanceRowIndex} não tem Professor Principal definido na planilha de instâncias.`);
                lock.releaseLock();
                return JSON.stringify({ success: false, message: "Erro interno: Horário fixo não tem Professor Principal definido na planilha de instâncias. Verifique a geração de instâncias.", data: null });
            }
        }

        // Tudo parece válido, proceder com a reserva

        const bookingId = Utilities.getUuid();
        const now = new Date();

        // 1. Atualizar Instancias de Horarios na linha encontrada
        const newStatus = (bookingType === TIPOS_RESERVA.REPOSICAO) ? STATUS_OCUPACAO.REPOSICAO_AGENDADA : STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA;
        const updatedInstanceRow = [...instanceDetails]; // Copia a linha atual
        updatedInstanceRow[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO] = newStatus;
        updatedInstanceRow[HEADERS.SCHEDULE_INSTANCES.ID_RESERVA] = bookingId;
        // O ID do Evento do Calendar será atualizado/adicionado após a criação/atualização do evento

        try {
            // Verifica o número de colunas antes de salvar
            const numColsInstance = HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR + 1;
            // Garante que a linha a ser salva tenha o número correto de colunas
            while (updatedInstanceRow.length < numColsInstance) updatedInstanceRow.push('');
            if (updatedInstanceRow.length > numColsInstance) updatedInstanceRow.length = numColsInstance; // Trunca se houver mais

            instancesSheet.getRange(instanceRowIndex, 1, 1, numColsInstance).setValues([updatedInstanceRow]);
            Logger.log(`Instância de horário ${instanceIdToBook} na linha ${instanceRowIndex} atualizada para ${newStatus}.`);
        } catch (e) {
            Logger.log(`Erro ao atualizar linha ${instanceRowIndex} na planilha "${SHEETS.SCHEDULE_INSTANCES}": ${e.message}`);
            lock.releaseLock();
            return JSON.stringify({ success: false, message: `Erro interno ao atualizar o status do horário na planilha. ${e.message}`, data: null });
        }


        // 2. Criar entrada em Reservas Detalhadas
        const newBookingRow = [];
        // Garante que o array tem o tamanho correto preenchendo com ''
        const numColsBooking = HEADERS.BOOKING_DETAILS.CRIADO_POR + 1;
        for (let colIdx = 0; colIdx < numColsBooking; colIdx++) {
            newBookingRow[colIdx] = '';
        }

        newBookingRow[HEADERS.BOOKING_DETAILS.ID_RESERVA] = bookingId;
        newBookingRow[HEADERS.BOOKING_DETAILS.TIPO_RESERVA] = bookingType;
        newBookingRow[HEADERS.BOOKING_DETAILS.ID_INSTANCIA] = instanceIdToBook;
        newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_REAL] = bookingDetails.professorReal.trim();
        newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL] = (bookingType === TIPOS_RESERVA.SUBSTITUICAO) ? professorPrincipalInstancia.trim() : '';
        newBookingRow[HEADERS.BOOKING_DETAILS.ALUNOS] = bookingDetails.alunos ? bookingDetails.alunos.trim() : '';
        newBookingRow[HEADERS.BOOKING_DETAILS.TURMAS_AGENDADA] = turmaInstancia; // Usa a turma lida da instância
        newBookingRow[HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL] = bookingDetails.disciplinaReal.trim();

        // Define a hora correta no objeto Date para salvar e usar no Calendar
        const [hour, minute] = bookingHourString.split(':').map(Number);
        bookingDateObj.setHours(hour, minute, 0, 0);
        newBookingRow[HEADERS.BOOKING_DETAILS.DATA_HORA_INICIO_EFETIVA] = bookingDateObj;

        newBookingRow[HEADERS.BOOKING_DETAILS.STATUS_RESERVA] = 'Agendada';
        newBookingRow[HEADERS.BOOKING_DETAILS.DATA_CRIACAO] = now;
        newBookingRow[HEADERS.BOOKING_DETAILS.CRIADO_POR] = userEmail;

        try {
            // Adiciona a nova linha na planilha de Reservas Detalhadas
            if (newBookingRow.length !== numColsBooking) {
                Logger.log(`Erro interno: newBookingRow tem ${newBookingRow.length} colunas, esperado ${numColsBooking}.`);
                // Preenche ou trunca por segurança
                while (newBookingRow.length < numColsBooking) newBookingRow.push('');
                if (newBookingRow.length > numColsBooking) newBookingRow.length = numColsBooking;
            }
            bookingsSheet.appendRow(newBookingRow);
            Logger.log(`Reserva ${bookingId} adicionada à planilha de Reservas Detalhadas.`);
        } catch (e) {
            Logger.log(`Erro ao adicionar reserva ${bookingId} à planilha "${SHEETS.BOOKING_DETAILS}": ${e.message}`);
            // Considere reverter a atualização da instância
            lock.releaseLock();
            return JSON.stringify({ success: false, message: `Reserva agendada na instância, mas erro ao salvar os detalhes da reserva. ${e.message}`, data: null });
        }


        // 3. Criar ou atualizar evento no Google Calendar
        let calendarEventId = null;
        try {
            const calendarId = getConfigValue('ID do Calendario');
            if (!calendarId || calendarId === '') {
                Logger.log('ID do Calendário não configurado. Pulando criação de evento.');
                lock.releaseLock();
                return JSON.stringify({ success: true, message: `Reserva agendada com sucesso, mas o ID do calendário não está configurado. Evento não criado/atualizado.`, data: { bookingId: bookingId, eventId: null } });
            }
            const calendar = CalendarApp.getCalendarById(calendarId);
            if (!calendar) {
                Logger.log(`Calendário com ID "${calendarId}" não encontrado ou acessível. Pulando criação/atualização de evento.`);
                lock.releaseLock();
                return JSON.stringify({ success: true, message: `Reserva agendada com sucesso, mas o calendário "${calendarId}" não foi encontrado ou não está acessível. Evento não criado/atualizado.`, data: { bookingId: bookingId, eventId: null } });
            }

            // Obter duração da configuração ou usar um padrão
            let durationMinutes = 45;
            const durationConfig = getConfigValue('Duracao Padrao Aula (minutos)');
            if (durationConfig && !isNaN(parseInt(durationConfig))) {
                durationMinutes = parseInt(durationConfig);
            } else {
                Logger.log(`Configuração "Duracao Padrao Aula (minutos)" não encontrada ou inválida. Usando padrão de ${durationMinutes} minutos.`);
            }

            const startTime = bookingDateObj; // Objeto Date já com a hora correta
            const endTime = new Date(startTime.getTime() + durationMinutes * 60 * 1000);

            let eventTitle = '';
            let eventDescription = `Reserva ID: ${bookingId}\nTipo: ${bookingType}\nCriado por: ${userEmail}`;
            // Inclui a Turma (lida da instância) no título ou descrição
            const disciplina = newBookingRow[HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL] || 'Disciplina Não Informada';
            const turmaTexto = newBookingRow[HEADERS.BOOKING_DETAILS.TURMAS_AGENDADA]; // Já contém turmaInstancia

            eventDescription += `\nTurma(s): ${turmaTexto}`;
            eventTitle = `${bookingType} - ${disciplina} - ${turmaTexto}`;

            // Lógica para adicionar convidados (Emails)
            const guests = [];
            const authUsersSheet = ss.getSheetByName(SHEETS.AUTHORIZED_USERS);
            const nameEmailMap = {};
            if (authUsersSheet) {
                const authUserData = authUsersSheet.getDataRange().getValues();
                if (authUserData.length > 1 && authUserData[0].length > Math.max(HEADERS.AUTHORIZED_USERS.EMAIL, HEADERS.AUTHORIZED_USERS.NOME)) {
                    for (let i = 1; i < authUserData.length; i++) {
                        const row = authUserData[i];
                        const email = (row.length > HEADERS.AUTHORIZED_USERS.EMAIL && typeof row[HEADERS.AUTHORIZED_USERS.EMAIL] === 'string') ? row[HEADERS.AUTHORIZED_USERS.EMAIL].trim() : '';
                        const name = (row.length > HEADERS.AUTHORIZED_USERS.NOME && typeof row[HEADERS.AUTHORIZED_USERS.NOME] === 'string') ? row[HEADERS.AUTHORIZED_USERS.NOME].trim() : '';
                        if (email && name) nameEmailMap[name] = email;
                    }
                } else {
                    Logger.log("Planilha Usuarios Autorizados vazia ou estrutura incorreta para buscar emails.");
                }

                const profRealNome = newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_REAL];
                if (profRealNome && nameEmailMap[profRealNome]) guests.push(nameEmailMap[profRealNome]);

                if (bookingType === TIPOS_RESERVA.SUBSTITUICAO) {
                    const profOriginalNome = newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL];
                    if (profOriginalNome && nameEmailMap[profOriginalNome]) guests.push(nameEmailMap[profOriginalNome]);
                }
                // Adicionar lógica para Alunos se necessário
            } else {
                Logger.log("Planilha Usuarios Autorizados não encontrada para adicionar convidados.");
            }

            // Buscar evento existente pelo ID para ATUALIZAR, ou criar um novo
            let event = null;
            if (calendarEventIdExisting && calendarEventIdExisting !== '') {
                try {
                    event = calendar.getEventById(calendarEventIdExisting);
                    Logger.log(`Encontrado evento existente ${calendarEventIdExisting} para atualização.`);
                    event.setTitle(eventTitle);
                    event.setDescription(eventDescription);
                    event.setTime(startTime, endTime);

                    // Atualizar convidados
                    const existingGuests = event.getGuestList().map(g => g.getEmail());
                    const newGuests = [...new Set(guests)]; // Lista atual de convidados (sem duplicatas)

                    // Remover convidados que não estão mais na lista 'newGuests'
                    existingGuests.forEach(guestEmail => {
                        if (!newGuests.includes(guestEmail)) {
                            try { event.removeGuest(guestEmail); } catch (removeErr) { Logger.log(`Falha ao remover convidado ${guestEmail}: ${removeErr}`); }
                        }
                    });
                    // Adicionar novos convidados que não estavam na lista 'existingGuests'
                    newGuests.forEach(guestEmail => {
                        if (!existingGuests.includes(guestEmail)) {
                            try { event.addGuest(guestEmail); } catch (addErr) { Logger.log(`Falha ao adicionar convidado ${guestEmail}: ${addErr}`); }
                        }
                    });

                } catch (e) {
                    Logger.log(`Evento do Calendar ID ${calendarEventIdExisting} não encontrado para atualização (pode ter sido excluído manualmente ou ID inválido): ${e}. Criando novo evento.`);
                    event = null; // Garante que um novo será criado
                }
            }

            if (!event) {
                // Cria um novo evento se não existia ou não foi encontrado
                const eventOptions = { description: eventDescription };
                if (guests.length > 0) {
                    const uniqueGuests = [...new Set(guests)];
                    eventOptions.guests = uniqueGuests.join(',');
                    eventOptions.sendInvites = true;
                    Logger.log("Convidados adicionados ao novo evento: " + uniqueGuests.join(', '));
                }
                event = calendar.createEvent(eventTitle, startTime, endTime, eventOptions);
                Logger.log(`Evento do Calendar criado com ID: ${event.getId()}`);
            } else {
                Logger.log(`Evento do Calendar ID ${event.getId()} atualizado.`);
            }

            // Salvar o ID do evento (novo ou atualizado) na planilha de Instâncias
            instancesSheet.getRange(instanceRowIndex, HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR + 1).setValue(event.getId());
            calendarEventId = event.getId();

        } catch (calendarError) {
            Logger.log('Erro crítico no Calendar: ' + calendarError.message + ' Stack: ' + calendarError.stack);
            // Não reverter a reserva da planilha, apenas logar o erro do Calendar e retornar sucesso parcial
            lock.releaseLock();
            return JSON.stringify({ success: true, message: `Reserva agendada com sucesso, mas houve um erro ao criar/atualizar o evento no Google Calendar: ${calendarError.message}. Verifique os logs.`, data: { bookingId: bookingId, eventId: null } });
        }

        lock.releaseLock();
        // Retorna sucesso total
        return JSON.stringify({ success: true, message: `${bookingType} agendada com sucesso!`, data: { bookingId: bookingId, eventId: calendarEventId } });

    } catch (e) {
        if (lock.hasLock()) {
            lock.releaseLock();
        }
        Logger.log('Erro no bookSlot: ' + e.message + ' Stack: ' + e.stack);
        return JSON.stringify({ success: false, message: 'Ocorreu um erro ao agendar: ' + e.message, data: null });
    }
} // Fim bookSlot


// --- Funções de Gerenciamento de Instâncias ---

/**
 * GERA INSTÂNCIAS FUTURAS DE HORÁRIOS NA PLANILHA 'Instancias de Horarios'.
 * Esta função deve ser executada periodicamente (usando um gatilho de tempo).
 */
function createScheduleInstances() {
    Logger.log('*** createScheduleInstances chamada ***');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const baseSheet = ss.getSheetByName(SHEETS.BASE_SCHEDULES);
    const instancesSheet = ss.getSheetByName(SHEETS.SCHEDULE_INSTANCES);
    const timeZone = ss.getSpreadsheetTimeZone();

    if (!baseSheet) {
        Logger.log(`Erro: Planilha "${SHEETS.BASE_SCHEDULES}" não encontrada.`);
        return;
    }
    if (!instancesSheet) {
        Logger.log(`Erro: Planilha "${SHEETS.SCHEDULE_INSTANCES}" não encontrada.`);
        return;
    }

    // --- 1. Lê e valida a planilha Horarios Base ---
    const baseDataRaw = baseSheet.getDataRange().getValues();
    if (baseDataRaw.length <= 1) {
        Logger.log(`Planilha "${SHEETS.BASE_SCHEDULES}" está vazia ou apenas cabeçalho.`);
        return;
    }

    const baseSchedules = [];
    const expectedBaseCols = Math.max(
        HEADERS.BASE_SCHEDULES.ID,
        HEADERS.BASE_SCHEDULES.DIA_SEMANA,
        HEADERS.BASE_SCHEDULES.HORA_INICIO,
        HEADERS.BASE_SCHEDULES.TIPO,
        HEADERS.BASE_SCHEDULES.TURMA_PADRAO,
        HEADERS.BASE_SCHEDULES.PROFESSOR_PRINCIPAL
    ) + 1;

    for (let i = 1; i < baseDataRaw.length; i++) {
        const row = baseDataRaw[i];
        const rowIndex = i + 1;

        if (!row || row.length < expectedBaseCols) {
            Logger.log(`Skipping incomplete base schedule row ${rowIndex}. Expected at least ${expectedBaseCols} columns, found ${row ? row.length : 0}.`);
            continue;
        }

        const baseIdRaw = row[HEADERS.BASE_SCHEDULES.ID];
        const baseDayOfWeekRaw = row[HEADERS.BASE_SCHEDULES.DIA_SEMANA];
        const baseHourRaw = row[HEADERS.BASE_SCHEDULES.HORA_INICIO];
        const baseTypeRaw = row[HEADERS.BASE_SCHEDULES.TIPO];
        const baseTurmaRaw = row[HEADERS.BASE_SCHEDULES.TURMA_PADRAO];
        const baseProfessorPrincipalRaw = row[HEADERS.BASE_SCHEDULES.PROFESSOR_PRINCIPAL];

        const baseId = (typeof baseIdRaw === 'string' || typeof baseIdRaw === 'number') ? String(baseIdRaw).trim() : null;
        const baseDayOfWeek = (typeof baseDayOfWeekRaw === 'string' || typeof baseDayOfWeekRaw === 'number') ? String(baseDayOfWeekRaw).trim() : null;
        const baseHourString = formatValueToHHMM(baseHourRaw, timeZone);
        const baseType = (typeof baseTypeRaw === 'string' || typeof baseTypeRaw === 'number') ? String(baseTypeRaw).trim() : null;
        const baseTurma = (typeof baseTurmaRaw === 'string' || typeof baseTurmaRaw === 'number') ? String(baseTurmaRaw).trim() : null;
        const baseProfessorPrincipal = (typeof baseProfessorPrincipalRaw === 'string' || typeof baseProfessorPrincipalRaw === 'number') ? String(baseProfessorPrincipalRaw || '').trim() : '';

        if (!baseId || baseId === '' || !baseDayOfWeek || baseDayOfWeek === '' || baseHourString === null || !baseType || baseType === '' || !baseTurma || baseTurma === '') {
            Logger.log(`Skipping base schedule row ${rowIndex} due to invalid/missing essential data: ID=${baseIdRaw}, Dia=${baseDayOfWeekRaw}, Hora=${baseHourRaw}, Tipo=${baseTypeRaw}, Turma=${baseTurmaRaw}`);
            continue;
        }

        const daysOfWeek = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
        if (!daysOfWeek.includes(baseDayOfWeek)) {
            Logger.log(`Skipping base schedule row ${rowIndex} with invalid Dia da Semana: "${baseDayOfWeek}"`);
            continue;
        }
        if (baseType !== TIPOS_HORARIO.FIXO && baseType !== TIPOS_HORARIO.VAGO) {
            Logger.log(`Skipping base schedule row ${rowIndex} with invalid Tipo: "${baseType}"`);
            continue;
        }
        if (baseType === TIPOS_HORARIO.FIXO && baseProfessorPrincipal === '') {
            Logger.log(`Skipping base schedule row ${rowIndex}: Horário Fixo (ID ${baseId}) não tem Professor Principal definido.`);
            continue;
        }

        baseSchedules.push({
            id: baseId, dayOfWeek: baseDayOfWeek, hour: baseHourString, type: baseType, turma: baseTurma, professorPrincipal: baseProfessorPrincipal
        });
    }

    if (baseSchedules.length === 0) {
        Logger.log("Nenhum horário base válido encontrado para gerar instâncias.");
        return;
    }
    Logger.log(`Processados ${baseSchedules.length} horários base válidos.`);


    // --- 2. Lê as instâncias existentes para verificação de duplicidade ---
    const existingInstancesRaw = instancesSheet.getDataRange().getValues();
    const existingInstancesMap = {}; // Key: ID_BASE_HORARIO + "_" + YYYY-MM-DD + "_" + HH:MM
    const mapKeyCols = Math.max(HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO, HEADERS.SCHEDULE_INSTANCES.DATA, HEADERS.SCHEDULE_INSTANCES.HORA_INICIO) + 1;
    const instanceIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA; // Para legibilidade

    if (existingInstancesRaw.length > 1) {
        if (existingInstancesRaw[0].length < mapKeyCols) {
            Logger.log(`Warning: Planilha "${SHEETS.SCHEDULE_INSTANCES}" tem menos colunas (${existingInstancesRaw[0].length}) que o esperado (${mapKeyCols}) para verificação de duplicidade.`);
        }

        for (let j = 1; j < existingInstancesRaw.length; j++) {
            const row = existingInstancesRaw[j];
            const rowIndex = j + 1;

            if (!row || row.length < mapKeyCols) {
                // Logger.log(`Skipping malformed existing instance row ${rowIndex} while building map.`);
                continue;
            }

            const existingBaseIdRaw = row[HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO];
            const existingDateRaw = row[HEADERS.SCHEDULE_INSTANCES.DATA];
            const existingHourRaw = row[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO];
            const existingInstanceIdRaw = (row.length > instanceIdCol) ? row[instanceIdCol] : null;

            const existingBaseId = (typeof existingBaseIdRaw === 'string' || typeof existingBaseIdRaw === 'number') ? String(existingBaseIdRaw).trim() : null;
            const existingDate = formatValueToDate(existingDateRaw);
            const existingHourString = formatValueToHHMM(existingHourRaw, timeZone);
            const existingInstanceId = (typeof existingInstanceIdRaw === 'string' || typeof existingInstanceIdRaw === 'number') ? String(existingInstanceIdRaw).trim() : null;

            if (existingBaseId && existingDate && existingHourString && existingInstanceId) {
                const existingDateStr = Utilities.formatDate(existingDate, timeZone, 'yyyy-MM-dd');
                const mapKey = `${existingBaseId}_${existingDateStr}_${existingHourString}`;
                existingInstancesMap[mapKey] = existingInstanceId;
            }
        }
    }
    Logger.log(`Map de instâncias existentes populado com ${Object.keys(existingInstancesMap).length} chaves.`);


    // --- 3. Gera novas instâncias ---
    const numWeeksToGenerate = 4;
    const today = new Date();
    today.setHours(0, 0, 0, 0);

    const newInstances = [];
    const daysOfWeek = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
    const numColsInstance = HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR + 1; // Número de colunas esperado

    let startGenerationDate = new Date(today.getTime());
    const currentDayOfWeek = startGenerationDate.getDay();
    const daysUntilMonday = (currentDayOfWeek === 0) ? 1 : (8 - currentDayOfWeek) % 7;
    if (daysUntilMonday !== 0) {
        startGenerationDate.setDate(startGenerationDate.getDate() + daysUntilMonday);
    }
    startGenerationDate.setHours(0, 0, 0, 0);

    const endGenerationDate = new Date(startGenerationDate.getTime());
    endGenerationDate.setDate(endGenerationDate.getDate() + (numWeeksToGenerate * 7) - 1);
    Logger.log(`Gerando instâncias de ${Utilities.formatDate(startGenerationDate, timeZone, 'yyyy-MM-dd')} até ${Utilities.formatDate(endGenerationDate, timeZone, 'yyyy-MM-dd')}`);


    let currentDate = new Date(startGenerationDate.getTime());
    while (currentDate <= endGenerationDate) {
        const targetDate = new Date(currentDate.getTime());
        const targetDayOfWeekName = daysOfWeek[targetDate.getDay()];

        const schedulesForThisDay = baseSchedules.filter(schedule => schedule.dayOfWeek === targetDayOfWeekName);

        for (const baseSchedule of schedulesForThisDay) {
            const baseId = baseSchedule.id;
            const baseHourString = baseSchedule.hour;
            const baseTurma = baseSchedule.turma;
            const baseProfessorPrincipal = baseSchedule.professorPrincipal;

            const instanceDateStr = Utilities.formatDate(targetDate, timeZone, 'yyyy-MM-dd');
            const predictableInstanceKey = `${baseId}_${instanceDateStr}_${baseHourString}`;

            if (!existingInstancesMap[predictableInstanceKey]) {
                const newRow = [];
                for (let colIdx = 0; colIdx < numColsInstance; colIdx++) { newRow[colIdx] = ''; } // Preenche com vazio

                newRow[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA] = Utilities.getUuid();
                newRow[HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO] = baseId;
                newRow[HEADERS.SCHEDULE_INSTANCES.TURMA] = baseTurma;
                newRow[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL] = baseProfessorPrincipal;
                newRow[HEADERS.SCHEDULE_INSTANCES.DATA] = targetDate;
                newRow[HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA] = targetDayOfWeekName;
                newRow[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO] = baseHourString; // Salva HORA COMO STRING "HH:mm"
                newRow[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL] = baseSchedule.type;
                newRow[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO] = STATUS_OCUPACAO.DISPONIVEL;
                // ID_RESERVA e ID_EVENTO_CALENDAR já estão ''

                newInstances.push(newRow);
                existingInstancesMap[predictableInstanceKey] = newRow[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA]; // Marca como adicionado
            }
        }
        currentDate.setDate(currentDate.getDate() + 1); // Avança para o próximo dia
    }

    Logger.log(`Pronto para inserir ${newInstances.length} novas instâncias.`);

    // --- 4. Adiciona as novas instâncias na planilha ---
    if (newInstances.length > 0) {
        // Validação final do número de colunas (já preenchido corretamente acima)
        if (newInstances[0].length !== numColsInstance) {
            Logger.log(`Erro interno: O array newInstances tem ${newInstances[0].length} colunas, mas esperava ${numColsInstance}.`);
            throw new Error("Erro na estrutura interna dos dados a serem salvos.");
        }

        try {
            instancesSheet.getRange(instancesSheet.getLastRow() + 1, 1, newInstances.length, numColsInstance).setValues(newInstances);
            Logger.log(`Geradas ${newInstances.length} novas instâncias de horários salvas.`);
        } catch (e) {
            Logger.log(`Erro ao salvar novas instâncias na planilha "${SHEETS.SCHEDULE_INSTANCES}": ${e.message} Stack: ${e.stack}`);
            throw new Error(`Erro ao salvar novas instâncias: ${e.message}`);
        }
    } else {
        Logger.log("Nenhuma nova instância de horário gerada para o período.");
    }
    Logger.log('*** createScheduleInstances finalizada ***');
}


// --- Funções Adicionais (Exemplo de Cancelamento - Mantido Comentado) ---
/*
function cancelBooking(bookingId) {
   const lock = LockService.getScriptLock();
   lock.waitLock(5000); // Espera no máximo 5 segundos

   const userEmail = Session.getActiveUser().getEmail();
   const userRole = getUserRolePlain(userEmail);

   if (!userRole) {
     lock.releaseLock();
     return JSON.stringify({ success: false, message: 'Usuário não autorizado a cancelar.', data: null });
   }

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const instancesSheet = ss.getSheetByName(SHEETS.SCHEDULE_INSTANCES);
        const bookingsSheet = ss.getSheetByName(SHEETS.BOOKING_DETAILS);
        const calendarId = getConfigValue('ID do Calendario');
        let calendar = null;
         if (calendarId) {
             try { calendar = CalendarApp.getCalendarById(calendarId.trim()); } catch(e) { Logger.log("Calendar not found for cancellation: " + e); }
         }

        // 1. Encontrar a reserva na planilha Reservas Detalhadas
        const bookingsData = bookingsSheet.getDataRange().getValues();
        let bookingRowIndex = -1; // Índice da linha no Sheets (baseado em 1)
        let bookingDetails = null; // Array da linha da reserva

        if (bookingsData.length <= 1) {
             lock.releaseLock();
             return JSON.stringify({ success: false, message: 'Reserva não encontrada (planilha de reservas vazia).', data: null });
        }

        for (let i = 1; i < bookingsData.length; i++) { // Começa do 1 para pular cabeçalho
            const row = bookingsData[i];
             if (row && row.length > HEADERS.BOOKING_DETAILS.ID_RESERVA) {
                const currentBookingId = (typeof row[HEADERS.BOOKING_DETAILS.ID_RESERVA] === 'string' || typeof row[HEADERS.BOOKING_DETAILS.ID_RESERVA] === 'number') ? String(row[HEADERS.BOOKING_DETAILS.ID_RESERVA]).trim() : null;

                if (currentBookingId && currentBookingId === bookingId) {
                    bookingRowIndex = i + 1; // Índice da linha no Sheets
                    bookingDetails = row; // Salva o array da linha completa
                    break;
                }
             }
        }

        if (bookingRowIndex === -1 || !bookingDetails) {
             lock.releaseLock();
            return JSON.stringify({ success: false, message: 'Reserva non trovata.', data: null }); // Corrigido italiano
        }

         // Opcional: Verificar se o usuário logado tem permissão para cancelar esta reserva
         // Ex: Apenas o criador, o professor envolvido, o admin.
         // const createdBy = (bookingDetails.length > HEADERS.BOOKING_DETAILS.CRIADO_POR) ? String(bookingDetails[HEADERS.BOOKING_DETAILS.CRIADO_POR] || '').trim() : '';
         // const professorReal = (bookingDetails.length > HEADERS.BOOKING_DETAILS.PROFESSOR_REAL) ? String(bookingDetails[HEADERS.BOOKING_DETAILS.PROFESSOR_REAL] || '').trim() : '';
         // if (userRole !== 'Admin' && createdBy !== userEmail && professorReal !== userEmail) {
         //     lock.releaseLock();
         //     return JSON.stringify({ success: false, message: 'Você não tem permissão para cancelar esta reserva.', data: null });
         // }

        // Verifica se a reserva já está cancelada
        const currentBookingStatus = (bookingDetails.length > HEADERS.BOOKING_DETAILS.STATUS_RESERVA) ? String(bookingDetails[HEADERS.BOOKING_DETAILS.STATUS_RESERVA] || '').trim() : '';
         if (currentBookingStatus === 'Cancelada') {
              lock.releaseLock();
              return JSON.stringify({ success: true, message: `Reserva ${bookingId} já estava cancelada.`, data: null }); // Considera sucesso se já cancelada
         }


        // 2. Encontrar a instância de horário correspondente na planilha Instancias de Horarios
        const instanceId = (bookingDetails.length > HEADERS.BOOKING_DETAILS.ID_INSTANCIA) ? String(bookingDetails[HEADERS.BOOKING_DETAILS.ID_INSTANCIA] || '').trim() : null;

        let instanceRowIndex = -1; // Índice da linha no Sheets (baseado em 1)
        let instanceDetails = null; // Array da linha da instância

        if (!instanceId || instanceId === '') {
             Logger.log(`Reserva ${bookingId} não tem ID de instância vinculado.`);
             // Continua cancelamento da reserva e Calendar (se ID Calendar existir direto na reserva?), mas avisa
        } else {
             // Busca a instância
             const instancesData = instancesSheet.getDataRange().getValues();
              // Need to check column count here too based on new headers
             const minInstanceCols = HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR + 1;

              if (instancesData.length > 1 && instancesData[0].length >= minInstanceCols) {
                  for (let i = 1; i < instancesData.length; i++) { // Começa do 1 para pular cabeçalho
                       const row = instancesData[i];
                       const currentInstanceId = (row.length > HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA) ? String(row[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA] || '').trim() : null;
                       if (currentInstanceId && currentInstanceId === instanceId) {
                           instanceRowIndex = i + 1; // Índice da linha no Sheets
                           instanceDetails = row; // Salva o array da linha completa
                           break;
                       }
                  }
              }

             if (instanceRowIndex === -1 || !instanceDetails) {
                 Logger.log(`Instância de horário ${instanceId} não encontrada na planilha Instancias de Horarios para a reserva ${bookingId} durante o cancelamento.`);
                 // Continua o cancelamento da reserva e do Calendar
             } else {
                  // 3. Resetar o status da instância de horário para 'Disponivel'
                  // Verifica se o status atual da instância ainda corresponde a esta reserva antes de resetar
                   const currentInstanceBookingId = (instanceDetails.length > HEADERS.SCHEDULE_INSTANCES.ID_RESERVA) ? String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.ID_RESERVA] || '').trim() : null;

                 if (currentInstanceBookingId && currentInstanceBookingId === bookingId) {
                      // Reseta para Disponivel (ou o status original para horários fixos se apropriado)
                      // Lê o tipo original da instância para decidir para qual status voltar
                      const originalType = (instanceDetails.length > HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL) ? String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL] || '').trim() : null;
                      let newStatus = STATUS_OCUPACAO.DISPONIVEL; // Padrão para Vago
                      if (originalType === TIPOS_HORARIO.FIXO) {
                          // Para Fixo, pode voltar para Disponivel (pronto para substituicao) ou outro status padrão se tiver
                          newStatus = STATUS_OCUPACAO.DISPONIVEL; // Assumindo que Fixo cancelado volta a ser Disponível para substituição
                      }


                      // Verifica se a linha da instância tem colunas suficientes antes de escrever
                       if(instanceDetails.length > HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO && instanceDetails.length > HEADERS.SCHEDULE_INSTANCES.ID_RESERVA) {
                          instancesSheet.getRange(instanceRowIndex, HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO + 1).setValue(newStatus);
                          instancesSheet.getRange(instanceRowIndex, HEADERS.SCHEDULE_INSTANCES.ID_RESERVA + 1).setValue(''); // Limpa o link para a reserva
                           Logger.log(`Status da instância ${instanceId} na linha ${instanceRowIndex} resetado para '${newStatus}'.`);
                       } else {
                            Logger.log(`Warning: Instância ${instanceId} na linha ${instanceRowIndex} não tem colunas suficientes para resetar status/ID Reserva.`);
                       }


                       // Opcional: Limpar o ID do evento do Calendar na instância se necessário, ou se o evento for excluído abaixo
                       // if(instanceDetails.length > HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR) {
                       //     instancesSheet.getRange(instanceRowIndex, HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR + 1).setValue('');
                       // }
                 } else {
                     Logger.log(`Instância ${instanceId} na linha ${instanceRowIndex} já não estava vinculada à reserva ${bookingId}. Status não alterado.`);
                 }
             }
        }


        // 4. Excluir o evento no Google Calendar (se existir)
         let eventId = null;
         if (instanceDetails && instanceDetails.length > HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR) {
              eventId = (typeof instanceDetails[HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR] === 'string' || typeof instanceDetails[HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR] === 'number') ? String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR] || '').trim() : null;
         }
         // Poderia também haver um Calendar ID armazenado diretamente na Reserva Detalhada se a vinculação da Instância falhou no agendamento.

         if (calendar && eventId && eventId !== '') {
              try {
                  const event = calendar.getEventById(eventId);
                   if (event) {
                       event.deleteEvent(); // Exclui o evento
                        // Limpa o ID do evento na planilha Instancias de Horarios (se não fez acima e instanceDetails foi encontrado)
                        if (instanceRowIndex !== -1 && instanceDetails.length > HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR) {
                             instancesSheet.getRange(instanceRowIndex, HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR + 1).setValue('');
                        }
                       Logger.log(`Evento do Calendar ${eventId} excluído.`);
                   } else {
                       Logger.log(`Evento do Calendar ${eventId} não encontrado para exclusão.`);
                   }
              } catch (e) {
                   Logger.log(`Erro ao excluir evento do Calendar ${eventId}: ${e.message}`);
              }
         }


        // 5. Marcar a reserva como cancelada na planilha Reservas Detalhadas
        // Use setValue para atualizar a célula específica de status
         if (bookingDetails.length > HEADERS.BOOKING_DETAILS.STATUS_RESERVA) {
             bookingsSheet.getRange(bookingRowIndex, HEADERS.BOOKING_DETAILS.STATUS_RESERVA + 1).setValue('Cancelada');
              Logger.log(`Reserva ${bookingId} na linha ${bookingRowIndex} marcada como 'Cancelada'.`);
         } else {
             Logger.log(`Warning: Reserva ${bookingId} na linha ${bookingRowIndex} não tem coluna de Status Reserva.`);
         }


        lock.releaseLock();
        let calendarMessage = calendar ? '' : ' (ID do Calendário não configurado ou inválido, verifique as Configurações)';
        let instanceMessage = (instanceRowIndex === -1) ? ' (Instância de horário associada não encontrada)' : '';
        return JSON.stringify({ success: true, message: `Reserva ${bookingId} cancelada com sucesso!${instanceMessage}${calendarMessage}`, data: null });

    } catch (e) {
       if (lock.hasLock()) {
         lock.releaseLock();
       }
       Logger.log('Erro no cancelBooking: ' + e.message + ' Stack: ' + e.stack);
       return JSON.stringify({ success: false, message: 'Ocorreu um erro ao cancelar a reserva: ' + e.message, data: null });
    }
}
*/