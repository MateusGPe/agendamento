// Obtém o ID da planilha Google Sheets atualmente ativa onde o script está sendo executado.
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// Define um objeto para armazenar os nomes exatos das abas (planilhas) usadas no script.
// Isso torna o código mais legível e fácil de manter, evitando erros de digitação nos nomes das planilhas.
const SHEETS = {
    CONFIG: 'Configuracoes', // Aba para configurações gerais do sistema.
    AUTHORIZED_USERS: 'Usuarios Autorizados', // Aba para listar usuários e seus papéis (Admin, Professor, Aluno).
    BASE_SCHEDULES: 'Horarios Base', // Aba com os horários "modelo" ou "template" que se repetem.
    SCHEDULE_INSTANCES: 'Instancias de Horarios', // Aba onde as ocorrências futuras dos horários base são geradas.
    BOOKING_DETAILS: 'Reservas Detalhadas', // Aba para registrar os detalhes de cada reserva feita.
    DISCIPLINES: 'Disciplinas' // Aba para listar as disciplinas disponíveis.
};

// Define objetos para mapear nomes de colunas legíveis para seus índices (base 0) em cada planilha.
// Essencial para acessar os dados corretos nas células, mesmo que a ordem das colunas mude (embora exija atualização aqui).
const HEADERS = {
    // Índices das colunas na aba 'Configuracoes'
    CONFIG: {
        NOME: 0,  // Coluna A: Nome da configuração
        VALOR: 1 // Coluna B: Valor da configuração
    },
    // Índices das colunas na aba 'Usuarios Autorizados'
    AUTHORIZED_USERS: {
        EMAIL: 0, // Coluna A: Email do usuário
        NOME: 1,  // Coluna B: Nome do usuário
        PAPEL: 2 // Coluna C: Papel/Função do usuário (Admin, Professor, Aluno)
    },
    // Índices das colunas na aba 'Horarios Base'
    BASE_SCHEDULES: {
        ID: 0,                    // Coluna A: Identificador único para o horário base
        TIPO: 1,                  // Coluna B: Tipo do horário (Fixo, Vago)
        DIA_SEMANA: 2,            // Coluna C: Dia da semana (Segunda, Terça, etc.)
        HORA_INICIO: 3,           // Coluna D: Hora de início do horário (formato HH:mm)
        DURACAO: 4,               // Coluna E: Duração em minutos (pode não ser usado ativamente no código fornecido, mas está definido)
        PROFESSOR_PRINCIPAL: 5, // Coluna F: Professor associado a este horário (relevante para horários Fixos)
        TURMA_PADRAO: 6,        // Coluna G: Turma padrão associada a este horário
        DISCIPLINA_PADRAO: 7,     // Coluna H: Disciplina padrão (pode não ser usado ativamente no código fornecido, mas está definido)
        CAPACIDADE: 8,            // Coluna I: Capacidade de alunos (pode não ser usado ativamente no código fornecido, mas está definido)
        OBSERVATIONS: 9           // Coluna J: Observações gerais (pode não ser usado ativamente no código fornecido, mas está definido)
    },
    // Índices das colunas na aba 'Instancias de Horarios'
    SCHEDULE_INSTANCES: {
        ID_INSTANCIA: 0,        // Coluna A: Identificador único para esta ocorrência específica do horário
        ID_BASE_HORARIO: 1,     // Coluna B: ID do horário base correspondente
        TURMA: 2,               // Coluna C: Turma associada a esta instância (geralmente copiada do horário base)
        PROFESSOR_PRINCIPAL: 3, // Coluna D: Professor principal associado (relevante se for tipo Fixo)
        DATA: 4,                // Coluna E: Data específica desta instância
        DIA_SEMANA: 5,            // Coluna F: Dia da semana (redundante com Data, mas pode facilitar filtros)
        HORA_INICIO: 6,         // Coluna G: Hora de início (copiada do horário base)
        TIPO_ORIGINAL: 7,       // Coluna H: Tipo original do horário base (Fixo, Vago)
        STATUS_OCUPACAO: 8,     // Coluna I: Status atual da instância (Disponivel, Reposicao Agendada, Substituicao Agendada)
        ID_RESERVA: 9,          // Coluna J: ID da reserva que ocupou esta instância (se aplicável)
        ID_EVENTO_CALENDAR: 10  // Coluna K: ID do evento no Google Calendar associado a esta instância/reserva
    },
    // Índices das colunas na aba 'Reservas Detalhadas'
    BOOKING_DETAILS: {
        ID_RESERVA: 0,              // Coluna A: Identificador único da reserva
        TIPO_RESERVA: 1,            // Coluna B: Tipo da reserva (Reposicao, Substituicao)
        ID_INSTANCIA: 2,            // Coluna C: ID da instância de horário que foi reservada
        PROFESSOR_REAL: 3,          // Coluna D: Professor que efetivamente ministrará a aula (pode ser diferente do original em substituições)
        PROFESSOR_ORIGINAL: 4,      // Coluna E: Professor original do horário (relevante para Substituições)
        ALUNOS: 5,                  // Coluna F: Nomes dos alunos (se aplicável, formato livre)
        TURMAS_AGENDADA: 6,         // Coluna G: Turma(s) para a qual a reserva foi feita (pode ser a mesma da instância ou diferente)
        DISCIPLINA_REAL: 7,         // Coluna H: Disciplina que será ministrada
        DATA_HORA_INICIO_EFETIVA: 8,// Coluna I: Data e hora de início combinadas da aula agendada
        STATUS_RESERVA: 9,          // Coluna J: Status da reserva (Agendada, Realizada, Cancelada - apenas 'Agendada' é usada aqui)
        DATA_CRIACAO: 10,           // Coluna K: Data e hora em que a reserva foi criada
        CRIADO_POR: 11              // Coluna L: Email do usuário que criou a reserva
    },
    // Índices das colunas na aba 'Disciplinas'
    DISCIPLINES: {
        NOME: 0 // Coluna A: Nome da disciplina
    }
};

// Define os valores padrão para o status de ocupação de uma instância de horário.
const STATUS_OCUPACAO = {
    DISPONIVEL: 'Disponivel',                     // O horário está livre para ser agendado.
    REPOSICAO_AGENDADA: 'Reposicao Agendada',     // Uma aula de reposição foi agendada neste horário.
    SUBSTITUICAO_AGENDADA: 'Substituicao Agendada' // Uma aula de substituição foi agendada neste horário.
};

// Define os tipos de reserva possíveis.
const TIPOS_RESERVA = {
    REPOSICAO: 'Reposicao',       // Agendamento em um horário originalmente "Vago".
    SUBSTITUICAO: 'Substituicao'  // Agendamento em um horário originalmente "Fixo", substituindo o professor/aula original.
};

// Define os tipos de horário base.
const TIPOS_HORARIO = {
    FIXO: 'Fixo', // Horário regular com professor e turma definidos.
    VAGO: 'Vago'  // Horário disponível na grade, sem aula fixa associada, usado para reposições.
};


/**
 * Tenta converter um valor bruto (geralmente de uma célula da planilha) para um objeto Date válido.
 * Trata casos específicos como a data "zero" do Google Sheets (30/12/1899) que pode ocorrer
 * se uma célula de hora for formatada incorretamente como data, retornando null nesses casos.
 * @param {*} rawValue O valor lido da célula.
 * @returns {Date|null} Um objeto Date válido ou null se a conversão falhar ou for a data "zero".
 */
function formatValueToDate(rawValue) {
    // Verifica se já é um objeto Date e se é um tempo válido
    if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
        // Verifica se a data é 30/12/1899 (artefato comum do Google Sheets para horas sem data)
        if (rawValue.getFullYear() === 1899 && rawValue.getMonth() === 11 && rawValue.getDate() === 30) {
            // Mesmo sendo 30/12/1899, se tiver hora, minuto ou segundo, poderia ser uma hora válida, mas a lógica atual retorna null.
            // Se for exatamente meia-noite (00:00:00), definitivamente deve ser null.
             if (rawValue.getHours() === 0 && rawValue.getMinutes() === 0 && rawValue.getSeconds() === 0) {
                 return null; // Retorna null para a data "zero" exata.
             }
             // Considerando que 30/12/1899 geralmente indica um problema de formatação, retorna null mesmo se houver horas/minutos.
             return null;
        }
        // Se for uma data válida e não for 30/12/1899, retorna o objeto Date.
        return rawValue;
    }

    // Se não for um objeto Date válido, retorna null.
    return null;
}

/**
 * Formata um valor bruto (Date, string ou número) para uma string de hora no formato "HH:mm".
 * @param {*} rawValue O valor lido da célula (pode ser Date, string "HH:MM" ou número entre 0 e 1).
 * @param {string} timeZone O fuso horário da planilha (ex: "America/Sao_Paulo").
 * @returns {string|null} A hora formatada como "HH:mm" ou null se a formatação falhar.
 */
function formatValueToHHMM(rawValue, timeZone) {
    // Se for um objeto Date válido
    if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
        // Formata a data usando o fuso horário fornecido para o formato HH:mm.
        return Utilities.formatDate(rawValue, timeZone, 'HH:mm');
    }
    // Se for uma string
    else if (typeof rawValue === 'string') {
        // Tenta encontrar um padrão HH:MM ou HH:MM:SS na string (ignorando espaços em branco)
        const timeMatch = rawValue.trim().match(/^(\d{1,2}):(\d{2})(:\d{2})?$/);
        if (timeMatch) {
            const hour = parseInt(timeMatch[1], 10);
            const minute = parseInt(timeMatch[2], 10);
            // Valida se a hora e o minuto estão dentro dos limites permitidos (0-23 para hora, 0-59 para minuto)
            if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                // Retorna a string formatada com zero à esquerda se necessário.
                return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
            }
        }
    }
    // Se for um número (formato de hora do Google Sheets, onde 0 = 00:00, 0.5 = 12:00, 1 = 24:00)
    else if (typeof rawValue === 'number' && rawValue >= 0 && rawValue <= 1) {
        // Converte a fração do dia para minutos totais.
        const totalMinutes = Math.round(rawValue * 24 * 60);
        // Calcula as horas e minutos.
        const hours = Math.floor(totalMinutes / 60);
        const minutes = totalMinutes % 60;
        // Valida se a hora e o minuto calculados são válidos.
        if (hours >= 0 && hours <= 23 && minutes >= 0 && minutes <= 59) {
             // Retorna a string formatada com zero à esquerda.
            return `${String(hours).padStart(2, '0')}:${String(minutes).padStart(2, '0')}`;
        }
    }
    // Se nenhum dos formatos for reconhecido ou for inválido, retorna null.
    return null;
}

/**
 * Obtém o papel (role) de um usuário específico buscando seu email na planilha 'Usuarios Autorizados'.
 * Esta é uma função interna, chamada por outras funções.
 * @param {string} userEmail O email do usuário a ser procurado.
 * @returns {string|null} O papel do usuário ('Admin', 'Professor', 'Aluno') ou null se não encontrado ou inválido.
 */
function getUserRolePlain(userEmail) {
    // Acessa a planilha de usuários autorizados pelo nome definido em SHEETS.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AUTHORIZED_USERS);
    // Verifica se a planilha foi encontrada.
    if (!sheet) {
        Logger.log(`Planilha "${SHEETS.AUTHORIZED_USERS}" não encontrada em getUserRolePlain.`);
        return null; // Retorna null se a planilha não existe.
    }
    // Obtém todos os dados da planilha.
    const data = sheet.getDataRange().getValues();
    // Verifica se há dados além do cabeçalho.
    if (data.length <= 1) {
        Logger.log(`Planilha "${SHEETS.AUTHORIZED_USERS}" vazia ou apenas cabeçalho.`);
        return null; // Retorna null se a planilha estiver vazia.
    }

    // Itera pelas linhas de dados (começando da segunda linha, índice 1, para pular o cabeçalho).
    for (let i = 1; i < data.length; i++) {
        // Verifica se a linha atual existe, tem colunas suficientes, e se a coluna de email existe, é uma string não vazia.
        if (data[i] && data[i].length > HEADERS.AUTHORIZED_USERS.PAPEL && data[i][HEADERS.AUTHORIZED_USERS.EMAIL] && typeof data[i][HEADERS.AUTHORIZED_USERS.EMAIL] === 'string' && data[i][HEADERS.AUTHORIZED_USERS.EMAIL].trim() !== '' && data[i][HEADERS.AUTHORIZED_USERS.EMAIL].trim() === userEmail) {
            // Se o email corresponder (após remover espaços extras), obtém o papel da coluna correspondente.
            const role = data[i][HEADERS.AUTHORIZED_USERS.PAPEL];

            // Verifica se o papel encontrado é um dos papéis válidos definidos.
            if (['Admin', 'Professor', 'Aluno'].includes(role)) {
                return role; // Retorna o papel válido encontrado.
            } else {
                // Se o papel na planilha for inválido, registra um log.
                Logger.log(`Papel inválido encontrado para o usuário ${userEmail} na linha ${i + 1} da planilha "${SHEETS.AUTHORIZED_USERS}": "${role}".`);
                // Continua procurando, caso o email apareça novamente com um papel válido (embora isso não devesse ocorrer).
            }
        }
    }
    // Se o loop terminar sem encontrar o email, registra um log.
    Logger.log(`Usuário "${userEmail}" não encontrado na lista de autorizados da planilha "${SHEETS.AUTHORIZED_USERS}".`);
    return null; // Retorna null se o usuário não foi encontrado.
}

/**
 * Obtém o valor de uma configuração específica da planilha 'Configuracoes'.
 * @param {string} configName O nome da configuração a ser buscada na coluna NOME.
 * @returns {string|null} O valor da configuração como string, ou null se não for encontrada.
 */
function getConfigValue(configName) {
    // Acessa a planilha de configurações.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
    // Verifica se a planilha existe.
    if (!sheet) {
        Logger.log(`Planilha "${SHEETS.CONFIG}" não encontrada.`);
        return null;
    }

    // Obtém todos os dados da planilha.
    const data = sheet.getDataRange().getValues();
     // Verifica se há dados além do cabeçalho.
    if (data.length <= 1) {
        Logger.log(`Planilha "${SHEETS.CONFIG}" vazia ou apenas cabeçalho.`);
        return null;
    }

    // Itera pelas linhas de dados (a partir da segunda linha).
    for (let i = 1; i < data.length; i++) {
         // Verifica se a linha existe, tem colunas suficientes e se o nome na coluna NOME corresponde ao solicitado.
        if (data[i] && data[i].length > HEADERS.CONFIG.VALOR && data[i][HEADERS.CONFIG.NOME] === configName) {
            // Retorna o valor da coluna VALOR, convertido para string e sem espaços extras.
            // Usa '' como fallback caso a célula esteja vazia (null ou undefined) antes de chamar trim().
            return String(data[i][HEADERS.CONFIG.VALOR] || '').trim();
        }
    }
    // Se o loop terminar sem encontrar a configuração, registra um log.
    Logger.log(`Configuração "${configName}" não encontrada na planilha "${SHEETS.CONFIG}".`);
    return null; // Retorna null se a configuração não foi encontrada.
}


/**
 * Função principal que responde a requisições GET (quando o script é acessado como Web App).
 * Verifica a autorização do usuário e serve a interface HTML principal.
 * @param {object} e O objeto de evento do Google Apps Script (não usado diretamente aqui).
 * @returns {HtmlService.HtmlOutput} O conteúdo HTML a ser exibido para o usuário.
 */
function doGet(e) {
    // Obtém o email do usuário que está acessando o Web App.
    const userEmail = Session.getActiveUser().getEmail();

    // Verifica o papel (role) do usuário usando a função auxiliar.
    const userRole = getUserRolePlain(userEmail);

    // Se o usuário não tiver um papel definido (não autorizado), retorna uma página de acesso negado.
    if (!userRole) {
        // Cria uma página HTML simples informando o acesso negado.
        return HtmlService.createHtmlOutput(
            '<h1>Acesso Negado</h1>' +
            '<p>Seu usuário (' + userEmail + ') não tem permissão para acessar esta aplicação. Entre em contato com o administrador.</p>'
        ).setTitle('Acesso Negado'); // Define o título da página.
    }

    // Se o usuário for autorizado, cria a interface a partir do arquivo HTML 'Index.html'.
    const htmlOutput = HtmlService.createTemplateFromFile('Index');
    // Passa o papel do usuário para o template HTML (para que o lado do cliente saiba o que exibir).
    htmlOutput.userRole = userRole;
    // Avalia o template (processa qualquer scriptlet فيه) e retorna o HTML final.
    return htmlOutput.evaluate()
        .setTitle('Sistema de Agendamento'); // Define o título da página principal.
}

/**
 * Função utilitária para ser usada dentro dos templates HTML (arquivos .html).
 * Permite incluir o conteúdo de outros arquivos HTML (ex: CSS, JS em blocos <style> ou <script>).
 * Uso no HTML: <?!= include('NomeDoArquivoSemExtensao'); ?>
 * @param {string} filename O nome do arquivo (sem a extensão .html) a ser incluído.
 * @returns {string} O conteúdo do arquivo solicitado.
 */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * Função exposta para ser chamada pelo lado do cliente (JavaScript no navegador).
 * Retorna o papel e o email do usuário atual em formato JSON.
 * @returns {string} Uma string JSON contendo {success, message, data: {role, email}}.
 */
function getUserRole() {
    Logger.log('*** getUserRole chamada ***'); // Log de início da função.
    // Obtém o email do usuário ativo na sessão.
    const userEmail = Session.getActiveUser().getEmail();
    Logger.log('Tentando obter papel para email: ' + userEmail); // Log do email.

    // Chama a função interna para buscar o papel na planilha.
    const userRole = getUserRolePlain(userEmail);

    // Retorna uma string JSON com o resultado da operação.
    return JSON.stringify({
        success: true, // Indica que a função em si executou (não necessariamente que encontrou o papel).
        message: userRole ? 'Papel do usuário obtido.' : 'Usuário não encontrado ou não autorizado.', // Mensagem informativa.
        data: {
            role: userRole, // O papel encontrado (ou null).
            email: userEmail // O email do usuário.
        }
    });
}

/**
 * Função exposta para ser chamada pelo lado do cliente.
 * Retorna uma lista dos nomes dos professores cadastrados na planilha 'Usuarios Autorizados'.
 * @returns {string} Uma string JSON contendo {success, message, data: [lista de nomes]}.
 */
function getProfessorsList() {
    Logger.log('*** getProfessorsList chamada ***');
    try {
        // Acessa a planilha de usuários autorizados.
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.AUTHORIZED_USERS);
        // Verifica se a planilha existe.
        if (!sheet) {
            Logger.log(`Erro: Planilha "${SHEETS.AUTHORIZED_USERS}" não encontrada em getProfessorsList.`);
            // Retorna JSON indicando falha interna.
            return JSON.stringify({ success: false, message: `Erro interno: Planilha "${SHEETS.AUTHORIZED_USERS}" não encontrada.`, data: [] });
        }

        // Obtém todos os dados.
        const data = sheet.getDataRange().getValues();
        // Verifica se há dados além do cabeçalho.
        if (data.length <= 1) {
            Logger.log(`Planilha "${SHEETS.AUTHORIZED_USERS}" vazia ou apenas cabeçalho.`);
            // Retorna JSON indicando sucesso, mas com lista vazia.
            return JSON.stringify({ success: true, message: 'Nenhum usuário autorizado encontrado.', data: [] });
        }

        const professors = []; // Array para armazenar os nomes dos professores.

        // Itera pelas linhas de dados (a partir da segunda linha).
        for (let i = 1; i < data.length; i++) {
            const row = data[i];

            // Verifica se a linha existe e tem colunas suficientes (até a coluna PAPEL).
            if (row && row.length > HEADERS.AUTHORIZED_USERS.PAPEL) {
                // Obtém o papel e o nome, convertendo para string e removendo espaços extras.
                const userRole = String(row[HEADERS.AUTHORIZED_USERS.PAPEL] || '').trim();
                const userName = String(row[HEADERS.AUTHORIZED_USERS.NOME] || '').trim();

                // Se o papel for 'Professor' e o nome não estiver vazio, adiciona à lista.
                if (userRole === 'Professor' && userName !== '') {
                    professors.push(userName);
                }
            }
        }

        // Ordena a lista de nomes de professores em ordem alfabética.
        professors.sort();

        Logger.log(`Encontrados ${professors.length} professores.`);
        // Retorna JSON indicando sucesso e a lista de nomes.
        return JSON.stringify({ success: true, message: 'Lista de professores obtida com sucesso.', data: professors });

    } catch (e) {
        // Em caso de erro inesperado durante a execução.
        Logger.log('Erro em getProfessorsList: ' + e.message + ' Stack: ' + e.stack);
        // Retorna JSON indicando falha e a mensagem de erro.
        return JSON.stringify({ success: false, message: 'Ocorreu um erro ao obter a lista de professores: ' + e.message, data: [] });
    }
}

/**
 * Função exposta para ser chamada pelo lado do cliente.
 * Retorna a lista de turmas disponíveis, lendo da configuração 'Turmas Disponiveis' na planilha 'Configuracoes'.
 * @returns {string} Uma string JSON contendo {success, message, data: [lista de turmas]}.
 */
function getTurmasList() {
    Logger.log('*** getTurmasList chamada ***');
    try {
        // Obtém o valor da configuração 'Turmas Disponiveis'.
        const turmasConfig = getConfigValue('Turmas Disponiveis');
        // Verifica se a configuração foi encontrada e não está vazia.
        if (!turmasConfig || turmasConfig === '') {
            Logger.log("Configuração 'Turmas Disponiveis' não encontrada ou vazia.");
            // Retorna sucesso, mas com lista vazia e mensagem informativa.
            return JSON.stringify({ success: true, message: "Configuração de turmas não encontrada ou vazia.", data: [] });
        }
        // Divide a string da configuração pela vírgula, remove espaços de cada item e filtra itens vazios.
        const turmasArray = turmasConfig.split(',').map(t => t.trim()).filter(t => t !== '');
        // Ordena as turmas em ordem alfabética.
        turmasArray.sort();

        Logger.log(`Encontradas ${turmasArray.length} turmas na configuração.`);
        // Retorna JSON com sucesso e a lista de turmas.
        return JSON.stringify({ success: true, message: 'Lista de turmas (config) obtida.', data: turmasArray });

    } catch (e) {
         // Em caso de erro inesperado.
        Logger.log('Erro em getTurmasList: ' + e.message + ' Stack: ' + e.stack);
        // Retorna JSON indicando falha.
        return JSON.stringify({ success: false, message: 'Ocorreu um erro ao obter a lista de turmas: ' + e.message, data: [] });
    }
}


/**
 * Função exposta para ser chamada pelo lado do cliente.
 * Retorna a lista de disciplinas disponíveis, lendo da planilha dedicada 'Disciplinas'.
 * @returns {string} Uma string JSON contendo {success, message, data: [lista de disciplinas]}.
 */
function getDisciplinesList() {
    Logger.log(`*** getDisciplinesList chamada (lendo da planilha dedicada: ${SHEETS.DISCIPLINES}) ***`);
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        // Acessa a planilha de disciplinas.
        const disciplinesSheet = ss.getSheetByName(SHEETS.DISCIPLINES);

        // Verifica se a planilha foi encontrada.
        if (!disciplinesSheet) {
            Logger.log(`Erro: Planilha "${SHEETS.DISCIPLINES}" não encontrada.`);
            // Retorna JSON indicando falha interna.
            return JSON.stringify({ success: false, message: `Erro interno: Planilha de Disciplinas "${SHEETS.DISCIPLINES}" não encontrada. Verifique o nome da aba.`, data: [] });
        }

        // Obtém todos os dados da planilha de disciplinas.
        const disciplinesData = disciplinesSheet.getDataRange().getValues();

        // Verifica se há dados além do cabeçalho.
        if (disciplinesData.length <= 1) {
            Logger.log(`Planilha "${SHEETS.DISCIPLINES}" vazia ou apenas cabeçalho.`);
             // Retorna JSON indicando sucesso, mas com lista vazia.
            return JSON.stringify({ success: true, message: `Nenhuma disciplina cadastrada na planilha "${SHEETS.DISCIPLINES}".`, data: [] });
        }

        const disciplinesArray = []; // Array para armazenar os nomes das disciplinas.

        // Itera pelas linhas de dados (a partir da segunda linha).
        for (let i = 1; i < disciplinesData.length; i++) {
            const row = disciplinesData[i];

            // Verifica se a linha existe e tem a coluna de nome.
            if (row && row.length > HEADERS.DISCIPLINES.NOME) {
                // Obtém o nome da disciplina, converte para string e remove espaços.
                const disciplineName = String(row[HEADERS.DISCIPLINES.NOME] || '').trim();

                // Se o nome não for vazio, adiciona à lista.
                if (disciplineName !== '') {
                    disciplinesArray.push(disciplineName);
                }
            }
        }

        // Ordena a lista de disciplinas alfabeticamente.
        disciplinesArray.sort();

        Logger.log(`Encontradas ${disciplinesArray.length} disciplinas na planilha "${SHEETS.DISCIPLINES}".`);
        // Retorna JSON com sucesso e a lista de disciplinas.
        return JSON.stringify({ success: true, message: 'Lista de disciplinas obtida com sucesso.', data: disciplinesArray });

    } catch (e) {
        // Em caso de erro inesperado.
        Logger.log(`Erro em getDisciplinesList (planilha dedicada "${SHEETS.DISCIPLINES}"): ${e.message} Stack: ${e.stack}`);
        // Retorna JSON indicando falha.
        return JSON.stringify({ success: false, message: `Ocorreu um erro ao obter a lista de disciplinas da planilha "${SHEETS.DISCIPLINES}": ${e.message}`, data: [] });
    }
}

/**
 * Função exposta para ser chamada pelo lado do cliente.
 * Busca e retorna as instâncias de horários disponíveis para um determinado tipo de reserva (Reposição ou Substituição).
 * @param {string} tipoReserva O tipo de reserva desejado ('Reposicao' ou 'Substituicao').
 * @returns {string} Uma string JSON contendo {success, message, data: [lista de horários disponíveis]}.
 */
function getAvailableSlots(tipoReserva) {
    Logger.log('*** getAvailableSlots chamada para tipo: ' + tipoReserva + ' ***');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Obtém o fuso horário da planilha para formatação correta de datas/horas.
    const timeZone = ss.getSpreadsheetTimeZone();

    // Verifica a autorização do usuário que está fazendo a chamada.
    const userEmail = Session.getActiveUser().getEmail();
    const userRole = getUserRolePlain(userEmail);

    // Se o usuário não for autorizado, retorna falha.
    if (!userRole) {
        Logger.log('Erro: Usuário no autorizado.');
        return JSON.stringify({ success: false, message: 'Usuário não autorizado a buscar horários.', data: null });
    }

    try {
        // Acessa a planilha de instâncias de horários.
        const sheet = ss.getSheetByName(SHEETS.SCHEDULE_INSTANCES);
        // Verifica se a planilha existe.
        if (!sheet) {
            Logger.log(`Erro: Planilha "${SHEETS.SCHEDULE_INSTANCES}" não encontrada!`);
            return JSON.stringify({ success: false, message: `Erro interno: Planilha "${SHEETS.SCHEDULE_INSTANCES}" não encontrada.`, data: null });
        }

        // Determina o número mínimo de colunas necessárias para ler os dados essenciais.
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
            HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR // Incluído para garantir consistência, mesmo que não usado diretamente no filtro.
        ) + 1; // +1 porque os índices são base 0.

        // Obtém todos os dados da planilha.
        const rawData = sheet.getDataRange().getValues();
        // Verifica se há dados além do cabeçalho.
        if (rawData.length <= 1) {
            Logger.log('Planilha Instancias de Horarios está vazia ou apenas cabeçalho.');
            return JSON.stringify({ success: true, message: 'Nenhuma instância de horário futuro encontrada. Gere instâncias primeiro.', data: [] });
        }

        // Verifica se a planilha tem o número mínimo de colunas esperado.
        if (rawData[0].length < minCols) {
            Logger.log(`Warning: Planilha "${SHEETS.SCHEDULE_INSTANCES}" tem menos colunas (${rawData[0].length}) do que o esperado (${minCols}). A leitura de dados pode falhar.`);
            // O script continua, mas pode falhar se tentar acessar uma coluna inexistente.
        }

        // Remove o cabeçalho dos dados.
        const data = rawData.slice(1);

        const availableSlots = []; // Array para armazenar os horários disponíveis encontrados.
        const today = new Date(); // Obtém a data atual.
        today.setHours(0, 0, 0, 0); // Zera a hora para comparar apenas a data.

        // Itera por todas as linhas de instâncias de horário.
        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const rowIndex = i + 2; // Índice da linha na planilha (considerando cabeçalho e base 1).

            // Pula a linha se for inválida ou não tiver colunas suficientes.
            if (!row || row.length < minCols) {
                Logger.log(`Skipping incomplete row ${rowIndex} in Instancias de Horarios. Missing required columns.`);
                continue; // Pula para a próxima iteração.
            }

            // Extrai os valores brutos das colunas relevantes.
            const instanceIdRaw = row[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA];
            const baseIdRaw = row[HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO];
            const turmaRaw = row[HEADERS.SCHEDULE_INSTANCES.TURMA];
            const professorPrincipalRaw = row[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL];
            const rawDate = row[HEADERS.SCHEDULE_INSTANCES.DATA];
            const instanceDiaSemanaRaw = row[HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA];
            const rawHoraInicio = row[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO];
            const originalTypeRaw = row[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL];
            const instanceStatusRaw = row[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO];

            // Formata e valida os dados essenciais.
            // Converte IDs para string e remove espaços, ou define como null se inválido.
            const instanceId = (typeof instanceIdRaw === 'string' || typeof instanceIdRaw === 'number') ? String(instanceIdRaw).trim() : null;
            const baseId = (typeof baseIdRaw === 'string' || typeof baseIdRaw === 'number') ? String(baseIdRaw).trim() : null;
            const turma = (typeof turmaRaw === 'string' || typeof turmaRaw === 'number') ? String(turmaRaw).trim() : null;
            // Trata professor principal (pode ser vazio).
            const professorPrincipal = (typeof professorPrincipalRaw === 'string' || typeof professorPrincipalRaw === 'number') ? String(professorPrincipalRaw || '').trim() : '';
            // Formata a data usando a função auxiliar.
            const instanceDate = formatValueToDate(rawDate);
            // Formata dia da semana e hora usando funções auxiliares.
            const instanceDiaSemana = (typeof instanceDiaSemanaRaw === 'string' || typeof instanceDiaSemanaRaw === 'number') ? String(instanceDiaSemanaRaw).trim() : null;
            const formattedHoraInicio = formatValueToHHMM(rawHoraInicio, timeZone);
            // Formata tipo original e status.
            const originalType = (typeof originalTypeRaw === 'string' || typeof originalTypeRaw === 'number') ? String(originalTypeRaw).trim() : null;
            const instanceStatus = (typeof instanceStatusRaw === 'string' || typeof instanceStatusRaw === 'number') ? String(instanceStatusRaw).trim() : null;

            // Verifica se algum dado essencial formatado é inválido ou ausente.
            if (!instanceId || instanceId === '' || !baseId || baseId === '' || !turma || turma === '' ||
                !instanceDate || // Verifica se a data é válida.
                !instanceDiaSemana || instanceDiaSemana === '' ||
                formattedHoraInicio === null || // Verifica se a hora é válida.
                !originalType || originalType === '' ||
                !instanceStatus || instanceStatus === '') {
                // Registra um log detalhado sobre a linha pulada e os valores brutos.
                Logger.log(`Skipping row ${rowIndex} due to invalid or missing essential data after formatting: ID=${instanceIdRaw}, BaseID=${baseIdRaw}, Turma=${turmaRaw}, ProfPrinc=${professorPrincipalRaw}, Data=${rawDate}, Dia=${instanceDiaSemanaRaw}, Hora=${rawHoraInicio}, Tipo=${originalTypeRaw}, Status=${instanceStatusRaw}`);
                continue; // Pula para a próxima linha.
            }

            // Validações adicionais de consistência dos dados.
            const daysOfWeek = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
            if (!daysOfWeek.includes(instanceDiaSemana)) {
                Logger.log(`Skipping row ${rowIndex} due to invalid Dia da Semana: "${instanceDiaSemana}". Raw: ${instanceDiaSemanaRaw}`);
                continue;
            }
            if (originalType !== TIPOS_HORARIO.FIXO && originalType !== TIPOS_HORARIO.VAGO) {
                Logger.log(`Skipping row ${rowIndex} due to invalid Tipo Original: "${originalType}". Raw: ${originalTypeRaw}`);
                continue;
            }
            if (instanceStatus !== STATUS_OCUPACAO.DISPONIVEL && instanceStatus !== STATUS_OCUPACAO.REPOSICAO_AGENDADA && instanceStatus !== STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA) {
                 Logger.log(`Skipping row ${rowIndex} due to invalid Status Ocupação: "${instanceStatus}". Raw: ${instanceStatusRaw}`);
                 continue;
            }

            // Compara a data da instância (sem a hora) com a data atual (sem a hora).
            const dateOnly = new Date(instanceDate.getFullYear(), instanceDate.getMonth(), instanceDate.getDate());
            if (dateOnly < today) {
                // Pula instâncias que já ocorreram.
                continue;
            }

            // Lógica de filtragem baseada no tipo de reserva solicitado.
            if (tipoReserva === TIPOS_RESERVA.REPOSICAO) {
                // Para Reposição, o horário deve ser do tipo VAGO e estar com status DISPONIVEL.
                if (originalType === TIPOS_HORARIO.VAGO && instanceStatus === STATUS_OCUPACAO.DISPONIVEL) {
                    // Adiciona o horário à lista de disponíveis.
                    availableSlots.push({
                        idInstancia: instanceId,
                        baseId: baseId,
                        turma: turma,
                        professorPrincipal: professorPrincipal, // Geralmente vazio para tipo VAGO.
                        data: Utilities.formatDate(instanceDate, timeZone, 'dd/MM/yyyy'), // Formata data para exibição.
                        diaSemana: instanceDiaSemana,
                        horaInicio: formattedHoraInicio,
                        tipoOriginal: originalType,
                        statusOcupacao: instanceStatus, // Será 'Disponivel'.
                    });
                }
            } else if (tipoReserva === TIPOS_RESERVA.SUBSTITUICAO) {
                // Para Substituição, o horário deve ser do tipo FIXO.
                // E NÃO pode estar com status REPOSICAO_AGENDADA (pois uma reposição já ocupou esse slot).
                // Pode estar DISPONIVEL ou já ter uma SUBSTITUICAO_AGENDADA (permitindo talvez reagendar, mas a lógica de bookSlot trata isso).
                 if (originalType === TIPOS_HORARIO.FIXO && instanceStatus !== STATUS_OCUPACAO.REPOSICAO_AGENDADA) {
                    // Adiciona o horário à lista de disponíveis para substituição.
                    availableSlots.push({
                        idInstancia: instanceId,
                        baseId: baseId,
                        turma: turma,
                        professorPrincipal: professorPrincipal, // Professor original do horário fixo.
                        data: Utilities.formatDate(instanceDate, timeZone, 'dd/MM/yyyy'),
                        diaSemana: instanceDiaSemana,
                        horaInicio: formattedHoraInicio,
                        tipoOriginal: originalType,
                        statusOcupacao: instanceStatus, // Pode ser 'Disponivel' ou 'Substituicao Agendada'.
                    });
                }
            }
        }

        Logger.log('Número de slots disponíveis encontrados: ' + availableSlots.length);

        // Retorna JSON com sucesso e a lista de horários encontrados.
        return JSON.stringify({ success: true, message: 'Slots carregados com sucesso.', data: availableSlots });

    } catch (e) {
        // Em caso de erro inesperado.
        Logger.log('Erro no getAvailableSlots: ' + e.message + ' Stack: ' + e.stack);
        // Retorna JSON indicando falha.
        return JSON.stringify({ success: false, message: 'Ocorreu um erro interno ao buscar horários: ' + e.message, data: null });
    }
}

/**
 * Função exposta para ser chamada pelo lado do cliente.
 * Tenta reservar um horário específico (instância), atualizando seu status e
 * criando um registro na planilha de detalhes da reserva e um evento no Google Calendar.
 * Usa LockService para evitar condições de corrida (duas pessoas tentando reservar o mesmo horário).
 * @param {string} jsonBookingDetailsString Uma string JSON contendo os detalhes da reserva (idInstancia, tipoReserva, professorReal, disciplinaReal, etc.).
 * @returns {string} Uma string JSON indicando sucesso ou falha da operação {success, message, data: {bookingId, eventId}}.
 */
function bookSlot(jsonBookingDetailsString) {
    // Obtém um bloqueio exclusivo para este script, esperando até 10 segundos se já estiver bloqueado.
    // Isso previne que duas execuções simultâneas desta função tentem modificar o mesmo horário ao mesmo tempo.
    const lock = LockService.getScriptLock();
    lock.waitLock(10000); // Espera até 10 segundos (10000 ms).

    // Verifica a autorização do usuário.
    const userEmail = Session.getActiveUser().getEmail();
    const userRole = getUserRolePlain(userEmail);

    // Se não autorizado, libera o bloqueio e retorna falha.
    if (!userRole) {
        lock.releaseLock();
        return JSON.stringify({ success: false, message: 'Usuário não autorizado a agendar.', data: null });
    }

    let bookingDetails;
    try {
        // Tenta converter a string JSON recebida em um objeto JavaScript.
        bookingDetails = JSON.parse(jsonBookingDetailsString);
        Logger.log("Booking details received and parsed: " + JSON.stringify(bookingDetails));
    } catch (e) {
        // Se o JSON for inválido, libera o bloqueio e retorna falha.
        lock.releaseLock();
        Logger.log('Erro ao parsear JSON de detalhes da reserva: ' + e.message);
        return JSON.stringify({ success: false, message: 'Erro ao processar dados da reserva.', data: null });
    }

    // Inicia o bloco principal de processamento da reserva.
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        // Acessa as planilhas de instâncias e de detalhes das reservas.
        const instancesSheet = ss.getSheetByName(SHEETS.SCHEDULE_INSTANCES);
        const bookingsSheet = ss.getSheetByName(SHEETS.BOOKING_DETAILS);
        // Obtém o fuso horário.
        const timeZone = ss.getSpreadsheetTimeZone();

        // Validação básica dos dados recebidos: ID da instância é obrigatório.
        if (!bookingDetails || typeof bookingDetails.idInstancia !== 'string' || bookingDetails.idInstancia.trim() === '') {
            lock.releaseLock();
            return JSON.stringify({ success: false, message: 'Dados de ID da instância de horário incompletos ou inválidos.', data: null });
        }

        // Obtém e limpa os dados essenciais da reserva.
        const instanceIdToBook = bookingDetails.idInstancia.trim();
        const bookingType = bookingDetails.tipoReserva ? String(bookingDetails.tipoReserva).trim() : null;
        // Valida o tipo de reserva.
        if (!bookingType || (bookingType !== TIPOS_RESERVA.REPOSICAO && bookingType !== TIPOS_RESERVA.SUBSTITUICAO)) {
            lock.releaseLock();
            return JSON.stringify({ success: false, message: 'Tipo de reserva inválido ou ausente.', data: null });
        }

        // --- Busca e Validação da Instância de Horário na Planilha ---
        // Obtém todos os dados da planilha de instâncias.
        const instanceDataRaw = instancesSheet.getDataRange().getValues();
        let instanceRowIndex = -1; // Índice da linha onde a instância foi encontrada.
        let instanceDetails = null; // Array com os dados da linha da instância.

        // Verifica se a planilha de instâncias tem dados.
        if (instanceDataRaw.length <= 1) {
            lock.releaseLock();
            return JSON.stringify({ success: false, message: 'Erro interno: Planilha de instâncias vazia ou estrutura incorreta.', data: null });
        }

        // Remove o cabeçalho.
        const instanceData = instanceDataRaw.slice(1);

        // Itera pelas linhas de instância para encontrar a que corresponde ao ID solicitado.
        for (let i = 0; i < instanceData.length; i++) {
            const row = instanceData[i];
            const rowIndex = i + 2; // Índice real na planilha.
            // Verifica se a linha tem pelo menos a coluna do ID.
            const minColsForId = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA + 1;
            if (row && row.length >= minColsForId) {
                const currentInstanceIdRaw = row[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA];
                // Converte e compara o ID da linha atual com o ID procurado.
                const currentInstanceId = (typeof currentInstanceIdRaw === 'string' || typeof currentInstanceIdRaw === 'number') ? String(currentInstanceIdRaw).trim() : null;
                if (currentInstanceId && currentInstanceId === instanceIdToBook) {
                    instanceRowIndex = rowIndex; // Armazena o índice da linha.
                    instanceDetails = row; // Armazena os dados da linha.
                    break; // Para o loop assim que encontrar.
                }
            }
        }

        // Se a instância não foi encontrada (pode ter sido deletada ou o ID estava errado).
        if (instanceRowIndex === -1 || !instanceDetails) {
            lock.releaseLock();
            // Mensagem importante para o usuário indicando possível concorrência ou dado desatualizado.
            return JSON.stringify({ success: false, message: 'Este horário não está mais disponível. Por favor, atualize a lista e tente novamente.', data: null });
        }

        // Verifica se a linha encontrada tem o número esperado de colunas para evitar erros.
        const expectedInstanceCols = Math.max(
            HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO,
            HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL,
            HEADERS.SCHEDULE_INSTANCES.DATA,
            HEADERS.SCHEDULE_INSTANCES.HORA_INICIO,
            HEADERS.SCHEDULE_INSTANCES.TURMA,
            HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL,
            HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR // Inclui a coluna do ID do evento.
        ) + 1;

        if (instanceDetails.length < expectedInstanceCols) {
            Logger.log(`Erro: Linha ${instanceRowIndex} na planilha Instancias de Horarios tem menos colunas (${instanceDetails.length}) que o esperado (${expectedInstanceCols}).`);
            lock.releaseLock();
            return JSON.stringify({ success: false, message: "Erro interno: Dados do horário selecionado incompletos na planilha.", data: null });
        }

        // Extrai e formata os dados relevantes da linha da instância encontrada.
        const currentStatusRaw = instanceDetails[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO];
        const originalTypeRaw = instanceDetails[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL];
        const rawBookingDate = instanceDetails[HEADERS.SCHEDULE_INSTANCES.DATA];
        const rawBookingTime = instanceDetails[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO];
        const turmaInstanciaRaw = instanceDetails[HEADERS.SCHEDULE_INSTANCES.TURMA];
        const professorPrincipalInstanciaRaw = instanceDetails[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL];
        const calendarEventIdExistingRaw = instanceDetails[HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR];

        // Formata e valida os dados extraídos.
        const currentStatus = (typeof currentStatusRaw === 'string' || typeof currentStatusRaw === 'number') ? String(currentStatusRaw).trim() : null;
        const originalType = (typeof originalTypeRaw === 'string' || typeof originalTypeRaw === 'number') ? String(originalTypeRaw).trim() : null;
        const turmaInstancia = (typeof turmaInstanciaRaw === 'string' || typeof turmaInstanciaRaw === 'number') ? String(turmaInstanciaRaw).trim() : null;
        const professorPrincipalInstancia = (typeof professorPrincipalInstanciaRaw === 'string' || typeof professorPrincipalInstanciaRaw === 'number') ? String(professorPrincipalInstanciaRaw || '').trim() : '';
        const bookingDateObj = formatValueToDate(rawBookingDate);
        const bookingHourString = formatValueToHHMM(rawBookingTime, timeZone);
        const calendarEventIdExisting = (typeof calendarEventIdExistingRaw === 'string' || typeof calendarEventIdExistingRaw === 'number') ? String(calendarEventIdExistingRaw || '').trim() : null;

        // Verifica se dados críticos formatados são válidos.
        if (!currentStatus || currentStatus === '' ||
            !originalType || originalType === '' ||
            !turmaInstancia || turmaInstancia === '' ||
            !bookingDateObj || bookingHourString === null) {
            Logger.log(`Erro: Dados críticos da instância ${instanceIdToBook} na linha ${instanceRowIndex} são inválidos. Status=${currentStatusRaw}, Tipo=${originalTypeRaw}, Turma=${turmaInstanciaRaw}, Data=${rawBookingDate}, Hora=${rawBookingTime}`);
            lock.releaseLock();
            return JSON.stringify({ success: false, message: "Erro interno: Dados do horário selecionado são inválidos na planilha.", data: null });
        }

        // --- Lógica de Validação Específica por Tipo de Reserva (VERIFICAÇÃO DE CONCORRÊNCIA) ---
        if (bookingType === TIPOS_RESERVA.REPOSICAO) {
             // Para REPOSICAO, a instância deve ser do tipo VAGO e estar DISPONIVEL no momento da reserva.
            if (originalType !== TIPOS_HORARIO.VAGO || currentStatus !== STATUS_OCUPACAO.DISPONIVEL) {
                lock.releaseLock();
                // Mensagem indicando que o status mudou desde que o usuário viu a lista.
                return JSON.stringify({ success: false, message: 'Este horário não está mais disponível para reposição ou não é um horário vago (concorrência).', data: null });
            }
            // Verifica se os campos obrigatórios para reposição foram preenchidos no formulário.
            if (!bookingDetails.professorReal || bookingDetails.professorReal.trim() === '' ||
                !bookingDetails.disciplinaReal || bookingDetails.disciplinaReal.trim() === '') {
                lock.releaseLock();
                return JSON.stringify({ success: false, message: 'Por favor, preencha todos os campos obrigatórios para reposição (Professor, Disciplina).', data: null });
            }

        } else if (bookingType === TIPOS_RESERVA.SUBSTITUICAO) {
            // Para SUBSTITUICAO, a instância deve ser do tipo FIXO.
            if (originalType !== TIPOS_HORARIO.FIXO) {
                lock.releaseLock();
                return JSON.stringify({ success: false, message: 'Este horário não é um horário fixo e não pode ser substituído.', data: null });
            }
             // Não pode ser substituído se já houver uma REPOSICAO agendada nele.
            if (currentStatus === STATUS_OCUPACAO.REPOSICAO_AGENDADA) {
                lock.releaseLock();
                return JSON.stringify({ success: false, message: 'Este horário fixo está sendo usado para uma reposição e não pode ser substituído.', data: null });
            }
             // Deve estar DISPONIVEL ou já marcado como SUBSTITUICAO_AGENDADA (permitindo reagendar/atualizar a substituição).
             // Se o status for qualquer outro (ex: algum status futuro inválido), a reserva falha.
            if (currentStatus !== STATUS_OCUPACAO.DISPONIVEL && currentStatus !== STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA) {
                lock.releaseLock();
                return JSON.stringify({ success: false, message: 'Este horário fixo não está disponível para substituição neste momento (concorrência).', data: null });
            }

            // Verifica campos obrigatórios para substituição.
            if (!bookingDetails.professorReal || bookingDetails.professorReal.trim() === '' ||
                !bookingDetails.disciplinaReal || bookingDetails.disciplinaReal.trim() === '') {
                lock.releaseLock();
                return JSON.stringify({ success: false, message: 'Por favor, preencha todos os campos obrigatórios para substituição (Professor Substituto, Disciplina).', data: null });
            }

            // Para substituição, é crucial que o professor original esteja definido na instância.
            if (professorPrincipalInstancia === '') {
                Logger.log(`Erro: Instância de horário fixo ${instanceIdToBook} na linha ${instanceRowIndex} não tem Professor Principal definido na planilha de instâncias.`);
                lock.releaseLock();
                return JSON.stringify({ success: false, message: "Erro interno: Horário fixo não tem Professor Principal definido na planilha de instâncias. Verifique a geração de instâncias.", data: null });
            }
        }

        // --- Se todas as validações passaram, prossegue com a reserva ---

        // Gera um ID único para a nova reserva.
        const bookingId = Utilities.getUuid();
        // Obtém a data/hora atual para registro.
        const now = new Date();

        // Determina o novo status da instância baseado no tipo de reserva.
        const newStatus = (bookingType === TIPOS_RESERVA.REPOSICAO) ? STATUS_OCUPACAO.REPOSICAO_AGENDADA : STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA;
        // Cria uma cópia dos dados da linha da instância para modificação.
        const updatedInstanceRow = [...instanceDetails];
        // Atualiza o status da ocupação.
        updatedInstanceRow[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO] = newStatus;
        // Associa o ID da reserva à instância.
        updatedInstanceRow[HEADERS.SCHEDULE_INSTANCES.ID_RESERVA] = bookingId;

        // --- Atualiza a Planilha de Instâncias ---
        try {
            // Define o número exato de colunas esperado na planilha de instâncias.
            const numColsInstance = HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR + 1;

            // Garante que a linha a ser escrita tenha exatamente o número correto de colunas.
            while (updatedInstanceRow.length < numColsInstance) updatedInstanceRow.push(''); // Adiciona colunas vazias se faltar.
            if (updatedInstanceRow.length > numColsInstance) updatedInstanceRow.length = numColsInstance; // Remove colunas extras se houver.

            // Escreve a linha atualizada de volta na planilha, na posição correta.
            instancesSheet.getRange(instanceRowIndex, 1, 1, numColsInstance).setValues([updatedInstanceRow]);
            Logger.log(`Instância de horário ${instanceIdToBook} na linha ${instanceRowIndex} atualizada para ${newStatus}.`);
        } catch (e) {
            // Se ocorrer um erro ao escrever na planilha de instâncias.
            Logger.log(`Erro ao atualizar linha ${instanceRowIndex} na planilha "${SHEETS.SCHEDULE_INSTANCES}": ${e.message}`);
            lock.releaseLock();
            return JSON.stringify({ success: false, message: `Erro interno ao atualizar o status do horário na planilha. ${e.message}`, data: null });
        }

        // --- Cria a Nova Linha na Planilha de Detalhes da Reserva ---
        const newBookingRow = []; // Array para a nova linha.

        // Define o número de colunas esperado na planilha de reservas.
        const numColsBooking = HEADERS.BOOKING_DETAILS.CRIADO_POR + 1;
        // Inicializa a linha com strings vazias para garantir o tamanho correto.
        for (let colIdx = 0; colIdx < numColsBooking; colIdx++) {
            newBookingRow[colIdx] = '';
        }

        // Preenche os dados da nova reserva.
        newBookingRow[HEADERS.BOOKING_DETAILS.ID_RESERVA] = bookingId;
        newBookingRow[HEADERS.BOOKING_DETAILS.TIPO_RESERVA] = bookingType;
        newBookingRow[HEADERS.BOOKING_DETAILS.ID_INSTANCIA] = instanceIdToBook;
        newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_REAL] = bookingDetails.professorReal.trim();
        // Professor original só é relevante para substituição.
        newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL] = (bookingType === TIPOS_RESERVA.SUBSTITUICAO) ? professorPrincipalInstancia.trim() : '';
        newBookingRow[HEADERS.BOOKING_DETAILS.ALUNOS] = bookingDetails.alunos ? bookingDetails.alunos.trim() : ''; // Campo opcional.
        newBookingRow[HEADERS.BOOKING_DETAILS.TURMAS_AGENDADA] = turmaInstancia; // Usa a turma da instância por padrão.
        newBookingRow[HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL] = bookingDetails.disciplinaReal.trim();

        // Monta o objeto Date/Time completo para o início efetivo da aula.
        const [hour, minute] = bookingHourString.split(':').map(Number);
        bookingDateObj.setHours(hour, minute, 0, 0); // Define a hora e minuto no objeto Date.
        newBookingRow[HEADERS.BOOKING_DETAILS.DATA_HORA_INICIO_EFETIVA] = bookingDateObj;

        newBookingRow[HEADERS.BOOKING_DETAILS.STATUS_RESERVA] = 'Agendada'; // Status inicial.
        newBookingRow[HEADERS.BOOKING_DETAILS.DATA_CRIACAO] = now; // Data/hora da criação.
        newBookingRow[HEADERS.BOOKING_DETAILS.CRIADO_POR] = userEmail; // Email do usuário que fez a reserva.

        // --- Adiciona a Linha à Planilha de Reservas ---
        try {
             // Verificação extra para garantir o número correto de colunas antes de adicionar.
             if (newBookingRow.length !== numColsBooking) {
                 Logger.log(`Erro interno: newBookingRow tem ${newBookingRow.length} colunas, esperado ${numColsBooking}. Ajustando...`);
                 // Ajusta o array se necessário (embora a inicialização acima deva prevenir isso).
                 while (newBookingRow.length < numColsBooking) newBookingRow.push('');
                 if (newBookingRow.length > numColsBooking) newBookingRow.length = numColsBooking;
            }
            // Adiciona a nova linha ao final da planilha de reservas.
            bookingsSheet.appendRow(newBookingRow);
            Logger.log(`Reserva ${bookingId} adicionada à planilha de Reservas Detalhadas.`);
        } catch (e) {
            // Se falhar ao adicionar a linha de reserva (mas a instância já foi atualizada).
            Logger.log(`Erro ao adicionar reserva ${bookingId} à planilha "${SHEETS.BOOKING_DETAILS}": ${e.message}`);
            // É um estado inconsistente, mas informa o usuário. A instância ficou reservada, mas os detalhes não foram salvos.
            // Idealmente, deveria tentar reverter a atualização da instância (rollback), mas isso adiciona complexidade.
            lock.releaseLock();
            return JSON.stringify({ success: false, message: `Reserva agendada na instância, mas erro ao salvar os detalhes da reserva. ${e.message}`, data: null });
        }

        // --- Integração com Google Calendar ---
        let calendarEventId = null; // Variável para armazenar o ID do evento criado/atualizado.
        try {
            // Obtém o ID do calendário da planilha de configurações.
            const calendarId = getConfigValue('ID do Calendario');
             // Se o ID não estiver configurado, pula a criação do evento.
            if (!calendarId || calendarId === '') {
                Logger.log('ID do Calendário não configurado. Pulando criação de evento.');
                lock.releaseLock();
                // Retorna sucesso, mas informa que o evento não foi criado.
                return JSON.stringify({ success: true, message: `Reserva agendada com sucesso, mas o ID do calendário não está configurado. Evento não criado/atualizado.`, data: { bookingId: bookingId, eventId: null } });
            }
            // Tenta obter o objeto Calendar usando o ID.
            const calendar = CalendarApp.getCalendarById(calendarId);
             // Se o calendário não for encontrado ou o script não tiver permissão.
            if (!calendar) {
                Logger.log(`Calendário com ID "${calendarId}" não encontrado ou acessível. Pulando criação/atualização de evento.`);
                lock.releaseLock();
                 // Retorna sucesso, mas informa sobre o problema com o calendário.
                return JSON.stringify({ success: true, message: `Reserva agendada com sucesso, mas o calendário "${calendarId}" não foi encontrado ou não está acessível. Evento não criado/atualizado.`, data: { bookingId: bookingId, eventId: null } });
            }

            // Define a duração padrão da aula (em minutos).
            let durationMinutes = 45; // Valor default.
            // Tenta obter a duração da configuração.
            const durationConfig = getConfigValue('Duracao Padrao Aula (minutos)');
            if (durationConfig && !isNaN(parseInt(durationConfig))) {
                durationMinutes = parseInt(durationConfig);
            } else {
                Logger.log(`Configuração "Duracao Padrao Aula (minutos)" não encontrada ou inválida. Usando padrão de ${durationMinutes} minutos.`);
            }

            // Calcula a hora de início e fim do evento.
            const startTime = bookingDateObj; // Já contém data e hora corretas.
            const endTime = new Date(startTime.getTime() + durationMinutes * 60 * 1000); // Adiciona a duração em milissegundos.

            // Define o título e a descrição do evento.
            let eventTitle = '';
            let eventDescription = `Reserva ID: ${bookingId}\nTipo: ${bookingType}\nCriado por: ${userEmail}`;

            // Usa os dados já formatados da newBookingRow.
            const disciplina = newBookingRow[HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL] || 'Disciplina Não Informada';
            const turmaTexto = newBookingRow[HEADERS.BOOKING_DETAILS.TURMAS_AGENDADA];

            eventDescription += `\nTurma(s): ${turmaTexto}`;
            // Título mais informativo.
            eventTitle = `${bookingType} - ${disciplina} - ${turmaTexto}`;

            // --- Prepara a lista de convidados para o evento ---
            const guests = []; // Array de emails dos convidados.
            const authUsersSheet = ss.getSheetByName(SHEETS.AUTHORIZED_USERS);
            const nameEmailMap = {}; // Mapa para buscar email pelo nome do professor.
            // Tenta ler a planilha de usuários para mapear nomes a emails.
            if (authUsersSheet) {
                const authUserData = authUsersSheet.getDataRange().getValues();
                // Verifica se a planilha tem dados e as colunas necessárias.
                if (authUserData.length > 1 && authUserData[0].length > Math.max(HEADERS.AUTHORIZED_USERS.EMAIL, HEADERS.AUTHORIZED_USERS.NOME)) {
                    // Cria o mapa Nome -> Email.
                    for (let i = 1; i < authUserData.length; i++) {
                        const row = authUserData[i];
                        const email = (row.length > HEADERS.AUTHORIZED_USERS.EMAIL && typeof row[HEADERS.AUTHORIZED_USERS.EMAIL] === 'string') ? row[HEADERS.AUTHORIZED_USERS.EMAIL].trim() : '';
                        const name = (row.length > HEADERS.AUTHORIZED_USERS.NOME && typeof row[HEADERS.AUTHORIZED_USERS.NOME] === 'string') ? row[HEADERS.AUTHORIZED_USERS.NOME].trim() : '';
                        if (email && name) nameEmailMap[name] = email;
                    }
                } else {
                    Logger.log("Planilha Usuarios Autorizados vazia ou estrutura incorreta para buscar emails.");
                }

                // Adiciona o email do professor real (que vai dar a aula).
                const profRealNome = newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_REAL];
                if (profRealNome && nameEmailMap[profRealNome]) guests.push(nameEmailMap[profRealNome]);

                // Se for substituição, adiciona também o email do professor original.
                if (bookingType === TIPOS_RESERVA.SUBSTITUICAO) {
                    const profOriginalNome = newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL];
                    if (profOriginalNome && nameEmailMap[profOriginalNome]) guests.push(nameEmailMap[profOriginalNome]);
                }

            } else {
                Logger.log("Planilha Usuarios Autorizados não encontrada para adicionar convidados.");
            }

            // --- Lógica para Atualizar ou Criar Evento ---
            let event = null;
            // Verifica se já existe um ID de evento associado a esta instância na planilha.
            if (calendarEventIdExisting && calendarEventIdExisting !== '') {
                try {
                    // Tenta obter o evento existente pelo ID.
                    event = calendar.getEventById(calendarEventIdExisting);
                    Logger.log(`Encontrado evento existente ${calendarEventIdExisting} para atualização.`);
                    // Se encontrou, atualiza os detalhes do evento existente.
                    event.setTitle(eventTitle);
                    event.setDescription(eventDescription);
                    event.setTime(startTime, endTime);

                    // Atualiza a lista de convidados (adiciona novos, remove antigos).
                    const existingGuests = event.getGuestList().map(g => g.getEmail());
                    const newGuests = [...new Set(guests)]; // Remove duplicatas da nova lista.

                    // Remove convidados que não estão mais na lista nova.
                    existingGuests.forEach(guestEmail => {
                        if (!newGuests.includes(guestEmail)) {
                            try { event.removeGuest(guestEmail); } catch (removeErr) { Logger.log(`Falha ao remover convidado ${guestEmail}: ${removeErr}`); }
                        }
                    });

                    // Adiciona convidados que estão na lista nova mas não estavam na antiga.
                    newGuests.forEach(guestEmail => {
                        if (!existingGuests.includes(guestEmail)) {
                            try { event.addGuest(guestEmail); } catch (addErr) { Logger.log(`Falha ao adicionar convidado ${guestEmail}: ${addErr}`); }
                        }
                    });

                } catch (e) {
                    // Se getEventById falhar (evento deletado, ID inválido, permissão?).
                    Logger.log(`Evento do Calendar ID ${calendarEventIdExisting} não encontrado para atualização (pode ter sido excluído manualmente ou ID inválido): ${e}. Criando novo evento.`);
                    event = null; // Reseta a variável para forçar a criação de um novo evento.
                }
            }

            // Se não havia evento existente ou a busca/atualização falhou.
            if (!event) {
                 // Cria um novo evento.
                const eventOptions = { description: eventDescription };
                // Adiciona convidados se houver.
                if (guests.length > 0) {
                    const uniqueGuests = [...new Set(guests)]; // Garante emails únicos.
                    eventOptions.guests = uniqueGuests.join(','); // Formato esperado pela API.
                    eventOptions.sendInvites = true; // Envia convites por email.
                    Logger.log("Convidados adicionados ao novo evento: " + uniqueGuests.join(', '));
                }
                event = calendar.createEvent(eventTitle, startTime, endTime, eventOptions);
                Logger.log(`Evento do Calendar criado com ID: ${event.getId()}`);
            } else {
                // Se o evento foi atualizado com sucesso.
                Logger.log(`Evento do Calendar ID ${event.getId()} atualizado.`);
            }

            // --- Salva o ID do Evento na Planilha de Instâncias ---
            // Atualiza a coluna ID_EVENTO_CALENDAR na linha da instância com o ID do evento (novo ou atualizado).
            instancesSheet.getRange(instanceRowIndex, HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR + 1).setValue(event.getId());
            // Armazena o ID para retornar na resposta JSON.
            calendarEventId = event.getId();

        } catch (calendarError) {
             // Se ocorrer um erro durante a interação com o Google Calendar.
            Logger.log('Erro crítico no Calendar: ' + calendarError.message + ' Stack: ' + calendarError.stack);
            // Libera o bloqueio.
            lock.releaseLock();
             // Retorna sucesso na reserva da planilha, mas informa sobre o erro no Calendar.
             // A reserva está feita no sistema, mas o evento pode estar ausente ou incorreto.
            return JSON.stringify({ success: true, message: `Reserva agendada com sucesso, mas houve um erro ao criar/atualizar o evento no Google Calendar: ${calendarError.message}. Verifique os logs.`, data: { bookingId: bookingId, eventId: null } });
        }

        // --- Finalização ---
        // Libera o bloqueio, pois todas as operações foram concluídas.
        lock.releaseLock();

        // Retorna JSON indicando sucesso total.
        return JSON.stringify({ success: true, message: `${bookingType} agendada com sucesso!`, data: { bookingId: bookingId, eventId: calendarEventId } });

    } catch (e) {
        // Captura qualquer erro não tratado no bloco try principal.
        // Verifica se o bloqueio ainda está ativo (pode não estar se o erro ocorreu antes da liberação).
        if (lock.hasLock()) {
            lock.releaseLock();
        }
        Logger.log('Erro no bookSlot: ' + e.message + ' Stack: ' + e.stack);
        // Retorna JSON indicando falha geral.
        return JSON.stringify({ success: false, message: 'Ocorreu um erro ao agendar: ' + e.message, data: null });
    }
}

/**
 * Função para gerar instâncias futuras de horários com base nos 'Horarios Base'.
 * Normalmente executada periodicamente por um gatilho (trigger) de tempo.
 * Cria registros na planilha 'Instancias de Horarios' para um período futuro definido.
 * Evita criar duplicatas se uma instância para o mesmo horário base, data e hora já existir.
 */
function createScheduleInstances() {
    Logger.log('*** createScheduleInstances chamada ***');
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    // Acessa as planilhas de horários base e instâncias.
    const baseSheet = ss.getSheetByName(SHEETS.BASE_SCHEDULES);
    const instancesSheet = ss.getSheetByName(SHEETS.SCHEDULE_INSTANCES);
    // Obtém o fuso horário para consistência de datas/horas.
    const timeZone = ss.getSpreadsheetTimeZone();

    // Verifica se as planilhas necessárias existem.
    if (!baseSheet) {
        Logger.log(`Erro: Planilha "${SHEETS.BASE_SCHEDULES}" não encontrada.`);
        return; // Aborta a execução.
    }
    if (!instancesSheet) {
        Logger.log(`Erro: Planilha "${SHEETS.SCHEDULE_INSTANCES}" não encontrada.`);
        return; // Aborta a execução.
    }

    // --- Leitura e Validação dos Horários Base ---
    const baseDataRaw = baseSheet.getDataRange().getValues();
    // Verifica se há dados na planilha base.
    if (baseDataRaw.length <= 1) {
        Logger.log(`Planilha "${SHEETS.BASE_SCHEDULES}" está vazia ou apenas cabeçalho.`);
        return; // Aborta se não houver horários base.
    }

    const baseSchedules = []; // Array para armazenar os horários base válidos.
    // Define o número mínimo de colunas necessárias nos horários base.
    const expectedBaseCols = Math.max(
        HEADERS.BASE_SCHEDULES.ID,
        HEADERS.BASE_SCHEDULES.DIA_SEMANA,
        HEADERS.BASE_SCHEDULES.HORA_INICIO,
        HEADERS.BASE_SCHEDULES.TIPO,
        HEADERS.BASE_SCHEDULES.TURMA_PADRAO,
        HEADERS.BASE_SCHEDULES.PROFESSOR_PRINCIPAL
    ) + 1;

    // Itera pelas linhas da planilha base (pulando cabeçalho).
    for (let i = 1; i < baseDataRaw.length; i++) {
        const row = baseDataRaw[i];
        const rowIndex = i + 1;

        // Pula linhas incompletas.
        if (!row || row.length < expectedBaseCols) {
            Logger.log(`Skipping incomplete base schedule row ${rowIndex}. Expected at least ${expectedBaseCols} columns, found ${row ? row.length : 0}.`);
            continue;
        }

        // Extrai os dados brutos das colunas relevantes.
        const baseIdRaw = row[HEADERS.BASE_SCHEDULES.ID];
        const baseDayOfWeekRaw = row[HEADERS.BASE_SCHEDULES.DIA_SEMANA];
        const baseHourRaw = row[HEADERS.BASE_SCHEDULES.HORA_INICIO];
        const baseTypeRaw = row[HEADERS.BASE_SCHEDULES.TIPO];
        const baseTurmaRaw = row[HEADERS.BASE_SCHEDULES.TURMA_PADRAO];
        const baseProfessorPrincipalRaw = row[HEADERS.BASE_SCHEDULES.PROFESSOR_PRINCIPAL];

        // Formata e valida os dados essenciais.
        const baseId = (typeof baseIdRaw === 'string' || typeof baseIdRaw === 'number') ? String(baseIdRaw).trim() : null;
        const baseDayOfWeek = (typeof baseDayOfWeekRaw === 'string' || typeof baseDayOfWeekRaw === 'number') ? String(baseDayOfWeekRaw).trim() : null;
        const baseHourString = formatValueToHHMM(baseHourRaw, timeZone); // Formata a hora.
        const baseType = (typeof baseTypeRaw === 'string' || typeof baseTypeRaw === 'number') ? String(baseTypeRaw).trim() : null;
        const baseTurma = (typeof baseTurmaRaw === 'string' || typeof baseTurmaRaw === 'number') ? String(baseTurmaRaw).trim() : null;
        const baseProfessorPrincipal = (typeof baseProfessorPrincipalRaw === 'string' || typeof baseProfessorPrincipalRaw === 'number') ? String(baseProfessorPrincipalRaw || '').trim() : ''; // Permite professor vazio.

        // Pula a linha se dados essenciais forem inválidos após formatação.
        if (!baseId || baseId === '' || !baseDayOfWeek || baseDayOfWeek === '' || baseHourString === null || !baseType || baseType === '' || !baseTurma || baseTurma === '') {
            Logger.log(`Skipping base schedule row ${rowIndex} due to invalid/missing essential data: ID=${baseIdRaw}, Dia=${baseDayOfWeekRaw}, Hora=${baseHourRaw}, Tipo=${baseTypeRaw}, Turma=${baseTurmaRaw}`);
            continue;
        }

        // Validações adicionais de consistência.
        const daysOfWeek = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
        if (!daysOfWeek.includes(baseDayOfWeek)) {
            Logger.log(`Skipping base schedule row ${rowIndex} with invalid Dia da Semana: "${baseDayOfWeek}"`);
            continue;
        }
        if (baseType !== TIPOS_HORARIO.FIXO && baseType !== TIPOS_HORARIO.VAGO) {
            Logger.log(`Skipping base schedule row ${rowIndex} with invalid Tipo: "${baseType}"`);
            continue;
        }
        // Horários fixos DEVEM ter um professor principal definido.
        if (baseType === TIPOS_HORARIO.FIXO && baseProfessorPrincipal === '') {
            Logger.log(`Skipping base schedule row ${rowIndex}: Horário Fixo (ID ${baseId}) não tem Professor Principal definido.`);
            continue;
        }

        // Adiciona o horário base validado à lista.
        baseSchedules.push({
            id: baseId,
            dayOfWeek: baseDayOfWeek,
            hour: baseHourString,
            type: baseType,
            turma: baseTurma,
            professorPrincipal: baseProfessorPrincipal
        });
    }

    // Se nenhum horário base válido foi encontrado, não há o que gerar.
    if (baseSchedules.length === 0) {
        Logger.log("Nenhum horário base válido encontrado para gerar instâncias.");
        return;
    }
    Logger.log(`Processados ${baseSchedules.length} horários base válidos.`);

    // --- Leitura das Instâncias Existentes para Evitar Duplicatas ---
    const existingInstancesRaw = instancesSheet.getDataRange().getValues();
    const existingInstancesMap = {}; // Mapa para armazenar chaves de instâncias existentes.
    // Colunas necessárias para criar a chave única de identificação de uma instância.
    const mapKeyCols = Math.max(HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO, HEADERS.SCHEDULE_INSTANCES.DATA, HEADERS.SCHEDULE_INSTANCES.HORA_INICIO) + 1;
    const instanceIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA; // Coluna do ID da instância.

    // Processa as instâncias existentes apenas se houver dados além do cabeçalho.
    if (existingInstancesRaw.length > 1) {
        // Verifica se a planilha de instâncias tem colunas suficientes para a chave.
        if (existingInstancesRaw[0].length < mapKeyCols) {
            Logger.log(`Warning: Planilha "${SHEETS.SCHEDULE_INSTANCES}" tem menos colunas (${existingInstancesRaw[0].length}) que o esperado (${mapKeyCols}) para verificação de duplicidade.`);
        }

        // Itera pelas linhas de instâncias existentes (pulando cabeçalho).
        for (let j = 1; j < existingInstancesRaw.length; j++) {
            const row = existingInstancesRaw[j];
            const rowIndex = j + 1; // Índice real na planilha.

            // Pula linhas incompletas que não permitem criar a chave.
            if (!row || row.length < mapKeyCols) {
                // Logger.log(`Skipping existing instance row ${rowIndex} for map creation due to insufficient columns.`);
                continue; // Pula silenciosamente para não poluir muito o log.
            }

            // Extrai os dados brutos para a chave e o ID da instância.
            const existingBaseIdRaw = row[HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO];
            const existingDateRaw = row[HEADERS.SCHEDULE_INSTANCES.DATA];
            const existingHourRaw = row[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO];
            // Pega o ID da instância (se a coluna existir).
            const existingInstanceIdRaw = (row.length > instanceIdCol) ? row[instanceIdCol] : null;

            // Formata e valida os dados para a chave.
            const existingBaseId = (typeof existingBaseIdRaw === 'string' || typeof existingBaseIdRaw === 'number') ? String(existingBaseIdRaw).trim() : null;
            const existingDate = formatValueToDate(existingDateRaw);
            const existingHourString = formatValueToHHMM(existingHourRaw, timeZone);
            // Formata o ID da instância.
            const existingInstanceId = (typeof existingInstanceIdRaw === 'string' || typeof existingInstanceIdRaw === 'number') ? String(existingInstanceIdRaw).trim() : null;

            // Se todos os componentes da chave e o ID da instância forem válidos.
            if (existingBaseId && existingDate && existingHourString && existingInstanceId) {
                // Formata a data como string 'yyyy-MM-dd' para a chave do mapa.
                const existingDateStr = Utilities.formatDate(existingDate, timeZone, 'yyyy-MM-dd');
                // Cria a chave única: IDBase_Data_Hora.
                const mapKey = `${existingBaseId}_${existingDateStr}_${existingHourString}`;
                // Adiciona a chave ao mapa, associada ao ID da instância (valor pode ser útil para debug).
                existingInstancesMap[mapKey] = existingInstanceId;
            }
        }
    }
    Logger.log(`Map de instâncias existentes populado com ${Object.keys(existingInstancesMap).length} chaves.`);

    // --- Geração das Novas Instâncias ---
    const numWeeksToGenerate = 4; // Define quantas semanas no futuro gerar.
    const today = new Date(); // Data atual.
    today.setHours(0, 0, 0, 0); // Zera a hora.

    const newInstances = []; // Array para armazenar as novas linhas de instância a serem inseridas.
    const daysOfWeek = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado']; // Mapeamento dia -> nome.
    // Número total de colunas na planilha de instâncias.
    const numColsInstance = HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR + 1;

    // Define a data de início da geração: a próxima segunda-feira a partir de hoje.
    let startGenerationDate = new Date(today.getTime());
    const currentDayOfWeek = startGenerationDate.getDay(); // 0=Domingo, 1=Segunda, ...
    // Calcula quantos dias faltam para a próxima segunda-feira.
    const daysUntilMonday = (currentDayOfWeek === 0) ? 1 : (8 - currentDayOfWeek) % 7;
    // Se hoje não for segunda, avança para a próxima segunda.
    if (daysUntilMonday !== 0) {
        startGenerationDate.setDate(startGenerationDate.getDate() + daysUntilMonday);
    }
    startGenerationDate.setHours(0, 0, 0, 0); // Garante que a hora está zerada.

    // Define a data final da geração (inclusive).
    const endGenerationDate = new Date(startGenerationDate.getTime());
    // Avança (número de semanas * 7 - 1) dias para cobrir o período desejado.
    endGenerationDate.setDate(endGenerationDate.getDate() + (numWeeksToGenerate * 7) - 1);
    Logger.log(`Gerando instâncias de ${Utilities.formatDate(startGenerationDate, timeZone, 'yyyy-MM-dd')} até ${Utilities.formatDate(endGenerationDate, timeZone, 'yyyy-MM-dd')}`);

    // Itera por cada dia dentro do período de geração.
    let currentDate = new Date(startGenerationDate.getTime());
    while (currentDate <= endGenerationDate) {
        const targetDate = new Date(currentDate.getTime()); // Cria cópia da data atual do loop.
        // Obtém o nome do dia da semana para a data atual.
        const targetDayOfWeekName = daysOfWeek[targetDate.getDay()];

        // Filtra os horários base que ocorrem neste dia da semana.
        const schedulesForThisDay = baseSchedules.filter(schedule => schedule.dayOfWeek === targetDayOfWeekName);

        // Para cada horário base que deve ocorrer neste dia:
        for (const baseSchedule of schedulesForThisDay) {
            const baseId = baseSchedule.id;
            const baseHourString = baseSchedule.hour;
            const baseTurma = baseSchedule.turma;
            const baseProfessorPrincipal = baseSchedule.professorPrincipal;

            // Cria a chave única para esta possível instância (IDBase_Data_Hora).
            const instanceDateStr = Utilities.formatDate(targetDate, timeZone, 'yyyy-MM-dd');
            const predictableInstanceKey = `${baseId}_${instanceDateStr}_${baseHourString}`;

            // Verifica se uma instância com essa chave JÁ EXISTE no mapa.
            if (!existingInstancesMap[predictableInstanceKey]) {
                // Se NÃO EXISTE, cria uma nova linha (array) para a instância.
                const newRow = [];
                // Inicializa a linha com strings vazias para todas as colunas.
                for (let colIdx = 0; colIdx < numColsInstance; colIdx++) { newRow[colIdx] = ''; }

                // Preenche os dados da nova instância.
                newRow[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA] = Utilities.getUuid(); // Gera um novo ID único.
                newRow[HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO] = baseId;
                newRow[HEADERS.SCHEDULE_INSTANCES.TURMA] = baseTurma;
                newRow[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL] = baseProfessorPrincipal;
                newRow[HEADERS.SCHEDULE_INSTANCES.DATA] = targetDate; // Objeto Date.
                newRow[HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA] = targetDayOfWeekName;
                newRow[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO] = baseHourString; // String HH:mm.
                newRow[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL] = baseSchedule.type;
                newRow[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO] = STATUS_OCUPACAO.DISPONIVEL; // Status inicial.
                // ID_RESERVA e ID_EVENTO_CALENDAR ficam vazios inicialmente.

                // Adiciona a nova linha ao array de instâncias a serem inseridas.
                newInstances.push(newRow);
                // Adiciona a chave desta nova instância ao mapa para evitar duplicatas dentro do mesmo ciclo de geração.
                existingInstancesMap[predictableInstanceKey] = newRow[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA];
            }
        }
        // Avança para o próximo dia.
        currentDate.setDate(currentDate.getDate() + 1);
    }

    Logger.log(`Pronto para inserir ${newInstances.length} novas instâncias.`);

    // --- Inserção das Novas Instâncias na Planilha ---
    // Verifica se há novas instâncias para inserir.
    if (newInstances.length > 0) {
        // Verificação de segurança: garante que as linhas a serem inseridas têm o número correto de colunas.
        if (newInstances[0].length !== numColsInstance) {
            Logger.log(`Erro interno: O array newInstances tem ${newInstances[0].length} colunas, mas esperava ${numColsInstance}. Abortando inserção.`);
            // Lança um erro para interromper, pois isso indica um problema na lógica de criação da linha.
            throw new Error("Erro na estrutura interna dos dados a serem salvos.");
        }

        try {
            // Insere TODAS as novas linhas de uma vez na planilha para melhor performance.
            // Começa a inserir na primeira linha vazia (getLastRow() + 1).
            instancesSheet.getRange(instancesSheet.getLastRow() + 1, 1, newInstances.length, numColsInstance).setValues(newInstances);
            Logger.log(`Geradas ${newInstances.length} novas instâncias de horários salvas.`);
        } catch (e) {
             // Se ocorrer um erro durante a inserção em lote.
            Logger.log(`Erro ao salvar novas instâncias na planilha "${SHEETS.SCHEDULE_INSTANCES}": ${e.message} Stack: ${e.stack}`);
             // Lança um erro para que o gatilho (se houver) registre a falha.
            throw new Error(`Erro ao salvar novas instâncias: ${e.message}`);
        }
    } else {
        // Se nenhuma nova instância foi gerada (talvez já existissem todas para o período).
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