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
// Obtém o ID da planilha Google Sheets atualmente ativa onde o script está
// sendo executado.
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

// Define um objeto para armazenar os nomes exatos das abas (planilhas) usadas
// no script. Isso torna o código mais legível e fácil de manter, evitando erros
// de digitação nos nomes das planilhas.
const SHEETS = {
    CONFIG: 'Configuracoes',  // Aba para configurações gerais do sistema.
    AUTHORIZED_USERS: 'Usuarios Autorizados',  // Aba para listar usuários e seus
    // papéis (Admin, Professor).
    BASE_SCHEDULES: 'Horarios Base',  // Aba com os horários "modelo" ou
    // "template" que se repetem.
    SCHEDULE_INSTANCES:
        'Instancias de Horarios',  // Aba onde as ocorrências futuras dos horários
    // base são geradas.
    BOOKING_DETAILS: 'Reservas Detalhadas',  // Aba para registrar os detalhes de
    // cada reserva feita.
    DISCIPLINES: 'Disciplinas'  // Aba para listar as disciplinas disponíveis.
};

// Define objetos para mapear nomes de colunas legíveis para seus índices (base
// 0) em cada planilha. Essencial para acessar os dados corretos nas células,
// mesmo que a ordem das colunas mude (embora exija atualização aqui).
const HEADERS = {
    // Índices das colunas na aba 'Configuracoes'
    CONFIG: {
        NOME: 0,  // Coluna A: Nome da configuração
        VALOR: 1  // Coluna B: Valor da configuração
    },
    // Índices das colunas na aba 'Usuarios Autorizados'
    AUTHORIZED_USERS: {
        EMAIL: 0,  // Coluna A: Email do usuário
        NOME: 1,   // Coluna B: Nome do usuário
        PAPEL: 2   // Coluna C: Papel/Função do usuário (Admin, Professor)
    },
    // Índices das colunas na aba 'Horarios Base'
    BASE_SCHEDULES: {
        ID: 0,           // Coluna A: Identificador único para o horário base
        TIPO: 1,         // Coluna B: Tipo do horário (Fixo, Vago)
        DIA_SEMANA: 2,   // Coluna C: Dia da semana (Segunda, Terça, etc.)
        HORA_INICIO: 3,  // Coluna D: Hora de início do horário (formato HH:mm)
        DURACAO: 4,  // Coluna E: Duração em minutos (pode não ser usado ativamente
        // no código fornecido, mas está definido)
        PROFESSOR_PRINCIPAL: 5,  // Coluna F: Professor associado a este horário
        // (relevante para horários Fixos)
        TURMA_PADRAO: 6,       // Coluna G: Turma padrão associada a este horário
        DISCIPLINA_PADRAO: 7,  // Coluna H: Disciplina padrão (pode não ser usado
        // ativamente no código fornecido, mas está definido)
        CAPACIDADE: 8,  // Coluna I: Capacidade de alunos (pode não ser usado
        // ativamente no código fornecido, mas está definido)
        OBSERVATIONS: 9  // Coluna J: Observações gerais (pode não ser usado
        // ativamente no código fornecido, mas está definido)
    },
    // Índices das colunas na aba 'Instancias de Horarios'
    SCHEDULE_INSTANCES: {
        ID_INSTANCIA: 0,  // Coluna A: Identificador único para esta ocorrência
        // específica do horário
        ID_BASE_HORARIO: 1,  // Coluna B: ID do horário base correspondente
        TURMA: 2,  // Coluna C: Turma associada a esta instância (geralmente copiada
        // do horário base)
        PROFESSOR_PRINCIPAL: 3,  // Coluna D: Professor principal associado
        // (relevante se for tipo Fixo)
        DATA: 4,        // Coluna E: Data específica desta instância
        DIA_SEMANA: 5,  // Coluna F: Dia da semana (redundante com Data, mas pode
        // facilitar filtros)
        HORA_INICIO: 6,    // Coluna G: Hora de início (copiada do horário base)
        TIPO_ORIGINAL: 7,  // Coluna H: Tipo original do horário base (Fixo, Vago)
        STATUS_OCUPACAO: 8,  // Coluna I: Status atual da instância (Disponivel,
        // Reposicao Agendada, Substituicao Agendada)
        ID_RESERVA:
            9,  // Coluna J: ID da reserva que ocupou esta instância (se aplicável)
        ID_EVENTO_CALENDAR: 10  // Coluna K: ID do evento no Google Calendar
        // associado a esta instância/reserva
    },
    // Índices das colunas na aba 'Reservas Detalhadas'
    BOOKING_DETAILS: {
        ID_RESERVA: 0,    // Coluna A: Identificador único da reserva
        TIPO_RESERVA: 1,  // Coluna B: Tipo da reserva (Reposicao, Substituicao)
        ID_INSTANCIA: 2,  // Coluna C: ID da instância de horário que foi reservada
        PROFESSOR_REAL:
            3,  // Coluna D: Professor que efetivamente ministrará a aula (pode ser
        // diferente do original em substituições)
        PROFESSOR_ORIGINAL: 4,  // Coluna E: Professor original do horário
        // (relevante para Substituições)
        ALUNOS: 5,  // Coluna F: Nomes dos alunos (se aplicável, formato livre)
        TURMAS_AGENDADA: 6,  // Coluna G: Turma(s) para a qual a reserva foi feita
        // (pode ser a mesma da instância ou diferente)
        DISCIPLINA_REAL: 7,  // Coluna H: Disciplina que será ministrada
        DATA_HORA_INICIO_EFETIVA:
            8,  // Coluna I: Data e hora de início combinadas da aula agendada
        STATUS_RESERVA: 9,  // Coluna J: Status da reserva (Agendada, Realizada,
        // Cancelada - apenas 'Agendada' é usada aqui)
        DATA_CRIACAO: 10,  // Coluna K: Data e hora em que a reserva foi criada
        CRIADO_POR: 11     // Coluna L: Email do usuário que criou a reserva
    },
    // Índices das colunas na aba 'Disciplinas'
    DISCIPLINES: {
        NOME: 0  // Coluna A: Nome da disciplina
    }
};

// Define os valores padrão para o status de ocupação de uma instância de
// horário.
const STATUS_OCUPACAO = {
    DISPONIVEL: 'Disponivel',  // O horário está livre para ser agendado.
    REPOSICAO_AGENDADA: 'Reposicao Agendada',  // Uma aula de reposição foi
    // agendada neste horário.
    SUBSTITUICAO_AGENDADA: 'Substituicao Agendada'  // Uma aula de substituição
    // foi agendada neste horário.
};

// Define os tipos de reserva possíveis.
const TIPOS_RESERVA = {
    REPOSICAO: 'Reposicao',  // Agendamento em um horário originalmente "Vago".
    SUBSTITUICAO:
        'Substituicao'  // Agendamento em um horário originalmente "Fixo",
    // substituindo o professor/aula original.
};

// Define os tipos de horário base.
const TIPOS_HORARIO = {
    FIXO: 'Fixo',  // Horário regular com professor e turma definidos.
    VAGO: 'Vago'  // Horário disponível na grade, sem aula fixa associada, usado
    // para reposições.
};


/**
 * Tenta converter um valor bruto (geralmente de uma célula da planilha) para um
 * objeto Date válido. Trata casos específicos como a data "zero" do Google
 * Sheets (30/12/1899) que pode ocorrer se uma célula de hora for formatada
 * incorretamente como data, retornando null nesses casos.
 * @param {*} rawValue O valor lido da célula.
 * @returns {Date|null} Um objeto Date válido ou null se a conversão falhar ou
 *     for a data "zero".
 */
function formatValueToDate(rawValue) {
    // Verifica se já é um objeto Date e se é um tempo válido
    if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
        // Verifica se a data é 30/12/1899 (artefato comum do Google Sheets para
        // horas sem data)
        if (rawValue.getFullYear() === 1899 && rawValue.getMonth() === 11 &&
            rawValue.getDate() === 30) {
            // Mesmo sendo 30/12/1899, se tiver hora, minuto ou segundo, poderia ser
            // uma hora válida, mas a lógica atual retorna null. Se for exatamente
            // meia-noite (00:00:00), definitivamente deve ser null.
            if (rawValue.getHours() === 0 && rawValue.getMinutes() === 0 &&
                rawValue.getSeconds() === 0) {
                return null;  // Retorna null para a data "zero" exata.
            }
            // Considerando que 30/12/1899 geralmente indica um problema de
            // formatação, retorna null mesmo se houver horas/minutos.
            return null;
        }
        // Se for uma data válida e não for 30/12/1899, retorna o objeto Date.
        return rawValue;
    }

    // Se não for um objeto Date válido, retorna null.
    return null;
}

/**
 * Envia um email para uma lista fixa de contatos.
 *
 * @param {string[]} recipientsArray Um array de strings com os emails dos destinatários.
 * @param {string} subject O assunto do email.
 * @param {string} bodyText O corpo do email em texto simples.
 * @param {string} [bodyHtml] (Opcional) O corpo do email em formato HTML.
 * @param {string} [sendAs='to'] (Opcional) Como enviar: 'to', 'cc' ou 'bcc'. Padrão 'to'.
 */
function enviarEmailListaFixa(recipientsArray, subject, bodyText, bodyHtml, sendAs) {
    try {
        if (!recipientsArray || recipientsArray.length === 0) {
            Logger.log("A lista de destinatários está vazia. Nenhum email enviado.");
            return;
        }

        // Filtra emails inválidos (básico)
        const validRecipients = recipientsArray
            .map(email => String(email).trim())
            .filter(email => email && email.includes('@'));

        if (validRecipients.length === 0) {
            Logger.log("Nenhum endereço de email válido encontrado na lista fornecida.");
            return;
        }

        const recipientString = validRecipients.join(',');
        const mailOptions = {
            subject: subject,
            body: bodyText,
        };

        const sendType = (sendAs || 'to').toLowerCase();

        if (sendType === 'bcc') {
            mailOptions.bcc = recipientString;
            mailOptions.to = Session.getActiveUser().getEmail(); // Boa prática
            Logger.log(`Enviando email via Bcc para ${validRecipients.length} destinatários.`);
        } else if (sendType === 'cc') {
            mailOptions.cc = recipientString;
            mailOptions.to = Session.getActiveUser().getEmail(); // Precisa de um 'to'
            Logger.log(`Enviando email com Cc para ${recipientString}`);
        } else {
            mailOptions.to = recipientString;
            Logger.log(`Enviando email para: ${recipientString}`);
        }

        if (bodyHtml) {
            mailOptions.htmlBody = bodyHtml;
        }

        Logger.log(`Assunto: ${subject}`);
        MailApp.sendEmail(mailOptions);
        Logger.log("Email enviado com sucesso para a lista fixa.");

    } catch (error) {
        Logger.log(`Erro ao enviar email: ${error.message}\n${error.stack}`);
    }
}

/**
 * Converte um valor (string 'dd/MM/yyyy' ou objeto Date) para um objeto Date válido.
 * Retorna null se o valor for inválido ou o formato incorreto para string.
 * A hora é definida para 00:00:00.
 * @param {*} value O valor lido da célula (pode ser string 'dd/MM/yyyy', Date, etc.).
 * @returns {Date|null} O objeto Date válido ou null se a conversão falhar ou for inválida.
 */
function parseDDMMYYYY(value) {
    // 1. Verifica se o valor já é um objeto Date válido
    if (value instanceof Date && !isNaN(value.getTime())) {
        // É um objeto Date válido. Zero a hora e retorna.
        const date = new Date(value.getTime()); // Cria uma cópia para não modificar o original (embora getValues retorne cópias)
        date.setHours(0, 0, 0, 0);
        return date;
    }

    // 2. Se não for um Date object, verifica se é uma string para tentar parsear
    if (typeof value === 'string') {
        const dateString = value.trim();
        const parts = dateString.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);
        if (!parts) {
            // Não corresponde ao formato esperado 'dd/MM/yyyy'
            Logger.log(`parseDDMMYYYY: String "${dateString}" não corresponde ao formato dd/MM/yyyy.`);
            return null;
        }

        const day = parseInt(parts[1], 10);
        const month = parseInt(parts[2], 10) - 1; // Mês é baseado em zero (0-11)
        const year = parseInt(parts[3], 10);

        // Validação básica dos componentes
        if (month < 0 || month > 11 || day < 1 || day > 31 || year < 1000) {
            Logger.log(`parseDDMMYYYY: Componentes de data inválidos: Dia=${day}, Mês=${month + 1}, Ano=${year} na string "${dateString}".`);
            return null;
        }

        // Cria o objeto Date. O construtor Date corrigirá dias inválidos (ex: 30 de Fev),
        // então verificamos se a data resultante corresponde às partes de entrada.
        const date = new Date(year, month, day);

        // Verifica se o Date object criado corresponde às partes originais
        if (date.getFullYear() !== year || date.getMonth() !== month || date.getDate() !== day) {
            // A data de entrada era inválida (ex: 30/02/2023)
            Logger.log(`parseDDMMYYYY: Data criada não corresponde aos componentes (provavelmente data inválida como 30/02): "${dateString}".`);
            return null;
        }

        // Zera o tempo para comparação apenas de datas
        date.setHours(0, 0, 0, 0);
        return date;
    }

    // 3. Se não for nem Date nem string, é um tipo inesperado
    Logger.log(`parseDDMMYYYY: Tipo de valor inesperado recebido: ${typeof value}. Valor: "${value}".`);
    return null;
}

/**
 * Formata um valor bruto (Date, string ou número) para uma string de hora no
 * formato "HH:mm".
 * @param {*} rawValue O valor lido da célula (pode ser Date, string "HH:MM" ou
 *     número entre 0 e 1).
 * @param {string} timeZone O fuso horário da planilha (ex:
 *     "America/Sao_Paulo").
 * @returns {string|null} A hora formatada como "HH:mm" ou null se a formatação
 *     falhar.
 */
function formatValueToHHMM(rawValue, timeZone) {
    // Se for um objeto Date válido
    if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
        // Formata a data usando o fuso horário fornecido para o formato HH:mm.
        return Utilities.formatDate(rawValue, timeZone, 'HH:mm');
    }
    // Se for uma string
    else if (typeof rawValue === 'string') {
        // Tenta encontrar um padrão HH:MM ou HH:MM:SS na string (ignorando espaços
        // em branco)
        const timeMatch = rawValue.trim().match(/^(\d{1,2}):(\d{2})(:\d{2})?$/);
        if (timeMatch) {
            const hour = parseInt(timeMatch[1], 10);
            const minute = parseInt(timeMatch[2], 10);
            // Valida se a hora e o minuto estão dentro dos limites permitidos (0-23
            // para hora, 0-59 para minuto)
            if (hour >= 0 && hour <= 23 && minute >= 0 && minute <= 59) {
                // Retorna a string formatada com zero à esquerda se necessário.
                return `${String(hour).padStart(2, '0')}:${String(minute).padStart(2, '0')}`;
            }
        }
    }
    // Se for um número (formato de hora do Google Sheets, onde 0 = 00:00, 0.5 =
    // 12:00, 1 = 24:00)
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
 * Obtém o papel (role) de um usuário específico buscando seu email na planilha
 * 'Usuarios Autorizados'. Esta é uma função interna, chamada por outras
 * funções.
 * @param {string} userEmail O email do usuário a ser procurado.
 * @returns {string|null} O papel do usuário ('Admin', 'Professor') ou null se
 *     não encontrado ou inválido.
 */
function getUserRolePlain(userEmail) {
    // Acessa a planilha de usuários autorizados pelo nome definido em SHEETS.
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
        SHEETS.AUTHORIZED_USERS);
    // Verifica se a planilha foi encontrada.
    if (!sheet) {
        Logger.log(`Planilha "${SHEETS.AUTHORIZED_USERS}" não encontrada em getUserRolePlain.`);
        return null;  // Retorna null se a planilha não existe.
    }
    // Obtém todos os dados da planilha.
    const data = sheet.getDataRange().getValues();
    // Verifica se há dados além do cabeçalho.
    if (data.length <= 1) {
        Logger.log(
            `Planilha "${SHEETS.AUTHORIZED_USERS}" vazia ou apenas cabeçalho.`);
        return null;  // Retorna null se a planilha estiver vazia.
    }

    // Itera pelas linhas de dados (começando da segunda linha, índice 1, para
    // pular o cabeçalho).
    for (let i = 1; i < data.length; i++) {
        // Verifica se a linha atual existe, tem colunas suficientes, e se a coluna
        // de email existe, é uma string não vazia.
        if (data[i] && data[i].length > HEADERS.AUTHORIZED_USERS.PAPEL &&
            data[i][HEADERS.AUTHORIZED_USERS.EMAIL] &&
            typeof data[i][HEADERS.AUTHORIZED_USERS.EMAIL] === 'string' &&
            data[i][HEADERS.AUTHORIZED_USERS.EMAIL].trim() !== '' &&
            data[i][HEADERS.AUTHORIZED_USERS.EMAIL].trim() === userEmail) {
            // Se o email corresponder (após remover espaços extras), obtém o papel da
            // coluna correspondente.
            const role = data[i][HEADERS.AUTHORIZED_USERS.PAPEL];

            // Verifica se o papel encontrado é um dos papéis válidos definidos.
            if (['Admin', 'Professor'].includes(role)) {
                return role;  // Retorna o papel válido encontrado.
            } else {
                // Se o papel na planilha for inválido, registra um log.
                Logger.log(
                    `Papel inválido encontrado para o usuário ${userEmail} na linha ${i + 1} da planilha "${SHEETS.AUTHORIZED_USERS}": "${role}".`);
                // Continua procurando, caso o email apareça novamente com um papel
                // válido (embora isso não devesse ocorrer).
            }
        }
    }
    // Se o loop terminar sem encontrar o email, registra um log.
    Logger.log(`Usuário "${userEmail}" não encontrado na lista de autorizados da planilha "${SHEETS.AUTHORIZED_USERS}".`);
    return null;  // Retorna null se o usuário não foi encontrado.
}

/**
 * Obtém o valor de uma configuração específica da planilha 'Configuracoes'.
 * @param {string} configName O nome da configuração a ser buscada na coluna
 *     NOME.
 * @returns {string|null} O valor da configuração como string, ou null se não
 *     for encontrada.
 */
function getConfigValue(configName) {
    // Acessa a planilha de configurações.
    const sheet =
        SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.CONFIG);
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
        // Verifica se a linha existe, tem colunas suficientes e se o nome na coluna
        // NOME corresponde ao solicitado.
        if (data[i] && data[i].length > HEADERS.CONFIG.VALOR &&
            data[i][HEADERS.CONFIG.NOME] === configName) {
            // Retorna o valor da coluna VALOR, convertido para string e sem espaços
            // extras. Usa '' como fallback caso a célula esteja vazia (null ou
            // undefined) antes de chamar trim().
            let value = data[i][HEADERS.CONFIG.VALOR]
            if (value instanceof Date) {
                return value;
            }
            return String(value || '').trim();
        }
    }
    // Se o loop terminar sem encontrar a configuração, registra um log.
    Logger.log(`Configuração "${configName}" não encontrada na planilha "${SHEETS.CONFIG}".`);
    return null;  // Retorna null se a configuração não foi encontrada.
}

/**
 * Função principal que responde a requisições GET (quando o script é acessado
 * como Web App). Verifica a autorização do usuário e serve a interface HTML
 * principal ou sub-páginas.
 * @param {object} e O objeto de evento do Google Apps Script contendo parâmetros
 *     da requisição.
 * @returns {HtmlService.HtmlOutput} O conteúdo HTML a ser exibido para o
 *     usuário.
 */
function doGet(e) {
    // Obtém o email do usuário que está acessando o Web App.
    const userEmail = Session.getActiveUser().getEmail();
    Logger.log(`Acesso Web App por: ${userEmail}. Parâmetros: ${JSON.stringify(e.parameter)}`);

    // Verifica o papel (role) do usuário usando a função auxiliar.
    const userRole = getUserRolePlain(userEmail);

    // Se o usuário não tiver um papel definido (não autorizado), retorna uma
    // página de acesso negado, independentemente da página solicitada.
    if (!userRole) {
        Logger.log(`Acesso negado para usuário: ${userEmail}. Papel: ${userRole}`);
        // Cria uma página HTML simples informando o acesso negado.
        return HtmlService
            .createHtmlOutput(
                '<h1>Acesso Negado</h1>' +
                '<p>Seu usuário (' + userEmail +
                ') não tem permissão para acessar esta aplicação. Entre em contato com o administrador.</p>')
            .setTitle('Acesso Negado');  // Define o título da página.
    }

    // --- Roteamento para sub-páginas ---
    const page = e.parameter.page; // Obtém o parâmetro 'page' da URL

    if (page === 'scheduleView') {
        Logger.log('Servindo ScheduleView.html');
        // Se a página solicitada for 'scheduleView', serve o HTML correspondente.
        // Não precisamos passar o papel do usuário aqui, pois a ScheduleViewJS
        // fará a própria checagem ao carregar filtros/dados.
        return HtmlService.createTemplateFromFile('ScheduleView')
            .evaluate()
            .setTitle('Visualizar Horários - Sistema de Agendamento');
    } else {
        // Página padrão (se nenhum parâmetro 'page' ou um parâmetro
        // desconhecido for especificado)
        Logger.log('Servindo Index.html (padrão)');
        const htmlOutput = HtmlService.createTemplateFromFile('Index');
        // Passa o papel do usuário para o template HTML principal.
        htmlOutput.userRole = userRole;
        // Avalia o template e retorna o HTML final.
        return htmlOutput.evaluate().setTitle(
            'Sistema de Agendamento');  // Define o título da página principal.
    }
}

/**
 * Função utilitária para ser usada dentro dos templates HTML (arquivos .html).
 * Permite incluir o conteúdo de outros arquivos HTML (ex: CSS, JS em blocos
 * <style> ou <script>). Uso no HTML: <?!= include('NomeDoArquivoSemExtensao');
 * ?>
 * @param {string} filename O nome do arquivo (sem a extensão .html) a ser
 *     incluído.
 * @returns {string} O conteúdo do arquivo solicitado.
 */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


/**
 * Função exposta para ser chamada pelo lado do cliente (JavaScript no
 * navegador). Retorna o papel e o email do usuário atual em formato JSON.
 * @returns {string} Uma string JSON contendo {success, message, data: {role,
 *     email}}.
 */
function getUserRole() {
    Logger.log('*** getUserRole chamada ***');  // Log de início da função.
    // Obtém o email do usuário ativo na sessão.
    const userEmail = Session.getActiveUser().getEmail();
    Logger.log('Tentando obter papel para email: ' + userEmail);  // Log do email.

    // Chama a função interna para buscar o papel na planilha.
    const userRole = getUserRolePlain(userEmail);

    // Retorna uma string JSON com o resultado da operação.
    return JSON.stringify({
        success: true,  // Indica que a função em si executou (não necessariamente
        // que encontrou o papel).
        message: userRole ?
            'Papel do usuário obtido.' :
            'Usuário não encontrado ou não autorizado.',  // Mensagem informativa.
        data: {
            role: userRole,   // O papel encontrado (ou null).
            email: userEmail  // O email do usuário.
        }
    });
}

/**
 * Função exposta para ser chamada pelo lado do cliente.
 * Retorna uma lista dos nomes dos professores cadastrados na planilha 'Usuarios
 * Autorizados'.
 * @returns {string} Uma string JSON contendo {success, message, data: [lista de
 *     nomes]}.
 */
function getProfessorsList() {
    Logger.log('*** getProfessorsList chamada ***');
    try {
        // Acessa a planilha de usuários autorizados.
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(
            SHEETS.AUTHORIZED_USERS);
        // Verifica se a planilha existe.
        if (!sheet) {
            Logger.log(`Erro: Planilha "${SHEETS.AUTHORIZED_USERS}" não encontrada em getProfessorsList.`);
            // Retorna JSON indicando falha interna.
            return JSON.stringify({
                success: false,
                message: `Erro interno: Planilha "${SHEETS.AUTHORIZED_USERS}" não encontrada.`,
                data: []
            });
        }

        // Obtém todos os dados.
        const data = sheet.getDataRange().getValues();
        // Verifica se há dados além do cabeçalho.
        if (data.length <= 1) {
            Logger.log(
                `Planilha "${SHEETS.AUTHORIZED_USERS}" vazia ou apenas cabeçalho.`);
            // Retorna JSON indicando sucesso, mas com lista vazia.
            return JSON.stringify({
                success: true,
                message: 'Nenhum usuário autorizado encontrado.',
                data: []
            });
        }

        const professors = [];  // Array para armazenar os nomes dos professores.

        // Itera pelas linhas de dados (a partir da segunda linha).
        for (let i = 1; i < data.length; i++) {
            const row = data[i];

            // Verifica se a linha existe e tem colunas suficientes (até a coluna
            // PAPEL).
            if (row && row.length > HEADERS.AUTHORIZED_USERS.PAPEL) {
                // Obtém o papel e o nome, convertendo para string e removendo espaços
                // extras.
                const userRole =
                    String(row[HEADERS.AUTHORIZED_USERS.PAPEL] || '').trim();
                const userName =
                    String(row[HEADERS.AUTHORIZED_USERS.NOME] || '').trim();

                // Se o papel for 'Professor' e o nome não estiver vazio, adiciona à
                // lista.
                if (userRole === 'Professor' && userName !== '') {
                    professors.push(userName);
                }
            }
        }

        // Ordena a lista de nomes de professores em ordem alfabética.
        professors.sort();

        Logger.log(`Encontrados ${professors.length} professores.`);
        // Retorna JSON indicando sucesso e a lista de nomes.
        return JSON.stringify({
            success: true,
            message: 'Lista de professores obtida com sucesso.',
            data: professors
        });

    } catch (e) {
        // Em caso de erro inesperado durante a execução.
        Logger.log(
            'Erro em getProfessorsList: ' + e.message + ' Stack: ' + e.stack);
        // Retorna JSON indicando falha e a mensagem de erro.
        return JSON.stringify({
            success: false,
            message: 'Ocorreu um erro ao obter a lista de professores: ' + e.message,
            data: []
        });
    }
}

/**
 * Função exposta para ser chamada pelo lado do cliente.
 * Retorna a lista de turmas disponíveis, lendo da configuração 'Turmas
 * Disponiveis' na planilha 'Configuracoes'.
 * @returns {string} Uma string JSON contendo {success, message, data: [lista de
 *     turmas]}.
 */
function getTurmasList() {
    Logger.log('*** getTurmasList chamada ***');
    try {
        // Obtém o valor da configuração 'Turmas Disponiveis'.
        const turmasConfig = getConfigValue('Turmas Disponiveis');
        // Verifica se a configuração foi encontrada e não está vazia.
        if (!turmasConfig || turmasConfig === '') {
            Logger.log(
                'Configuração \'Turmas Disponiveis\' não encontrada ou vazia.');
            // Retorna sucesso, mas com lista vazia e mensagem informativa.
            return JSON.stringify({
                success: true,
                message: 'Configuração de turmas não encontrada ou vazia.',
                data: []
            });
        }
        // Divide a string da configuração pela vírgula, remove espaços de cada item
        // e filtra itens vazios.
        const turmasArray =
            turmasConfig.split(',').map(t => t.trim()).filter(t => t !== '');
        // Ordena as turmas em ordem alfabética.
        turmasArray.sort();

        Logger.log(`Encontradas ${turmasArray.length} turmas na configuração.`);
        // Retorna JSON com sucesso e a lista de turmas.
        return JSON.stringify({
            success: true,
            message: 'Lista de turmas (config) obtida.',
            data: turmasArray
        });

    } catch (e) {
        // Em caso de erro inesperado.
        Logger.log('Erro em getTurmasList: ' + e.message + ' Stack: ' + e.stack);
        // Retorna JSON indicando falha.
        return JSON.stringify({
            success: false,
            message: 'Ocorreu um erro ao obter a lista de turmas: ' + e.message,
            data: []
        });
    }
}


/**
 * Função exposta para ser chamada pelo lado do cliente.
 * Retorna a lista de disciplinas disponíveis, lendo da planilha dedicada
 * 'Disciplinas'.
 * @returns {string} Uma string JSON contendo {success, message, data: [lista de
 *     disciplinas]}.
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
            return JSON.stringify({
                success: false,
                message: `Erro interno: Planilha de Disciplinas "${SHEETS.DISCIPLINES}" não encontrada. Verifique o nome da aba.`,
                data: []
            });
        }

        // Obtém todos os dados da planilha de disciplinas.
        const disciplinesData = disciplinesSheet.getDataRange().getValues();

        // Verifica se há dados além do cabeçalho.
        if (disciplinesData.length <= 1) {
            Logger.log(`Planilha "${SHEETS.DISCIPLINES}" vazia ou apenas cabeçalho.`);
            // Retorna JSON indicando sucesso, mas com lista vazia.
            return JSON.stringify({
                success: true,
                message: `Nenhuma disciplina cadastrada na planilha "${SHEETS.DISCIPLINES}".`,
                data: []
            });
        }

        const disciplinesArray =
            [];  // Array para armazenar os nomes das disciplinas.

        // Itera pelas linhas de dados (a partir da segunda linha).
        for (let i = 1; i < disciplinesData.length; i++) {
            const row = disciplinesData[i];

            // Verifica se a linha existe e tem a coluna de nome.
            if (row && row.length > HEADERS.DISCIPLINES.NOME) {
                // Obtém o nome da disciplina, converte para string e remove espaços.
                const disciplineName =
                    String(row[HEADERS.DISCIPLINES.NOME] || '').trim();

                // Se o nome não for vazio, adiciona à lista.
                if (disciplineName !== '') {
                    disciplinesArray.push(disciplineName);
                }
            }
        }

        // Ordena a lista de disciplinas alfabeticamente.
        disciplinesArray.sort();

        Logger.log(
            `Encontradas ${disciplinesArray.length} disciplinas na planilha "${SHEETS.DISCIPLINES}".`);
        // Retorna JSON com sucesso e a lista de disciplinas.
        return JSON.stringify({
            success: true,
            message: 'Lista de disciplinas obtida com sucesso.',
            data: disciplinesArray
        });

    } catch (e) {
        // Em caso de erro inesperado.
        Logger.log(`Erro em getDisciplinesList (planilha dedicada "${SHEETS.DISCIPLINES}"): ${e.message} Stack: ${e.stack}`);
        // Retorna JSON indicando falha.
        return JSON.stringify({
            success: false,
            message: `Ocorreu um erro ao obter a lista de disciplinas da planilha "${SHEETS.DISCIPLINES}": ${e.message}`,
            data: []
        });
    }
}

/**
 * Função exposta para o lado do cliente da ScheduleView.
 * Retorna as listas de turmas e datas de início de semana disponíveis para filtros.
 * Calcula as próximas N semanas (começando na Segunda).
 * @returns {string} JSON string {success, message, data: {turmas: [], weekStartDates: []}}
 */
function getScheduleViewFilters() {
    Logger.log('*** getScheduleViewFilters chamada ***');
    try {
        // Verifica autorização mínima (qualquer papel é suficiente para visualizar)
        const userEmail = Session.getActiveUser().getEmail();
        const userRole = getUserRolePlain(userEmail);
        if (!userRole) {
            Logger.log('Erro: Usuário não autorizado a obter filtros de visualização.');
            return JSON.stringify({ success: false, message: 'Usuário não autorizado.', data: null });
        }

        // Obter lista de turmas (reutiliza a lógica existente ou chama a função)
        const turmasResponse = JSON.parse(getTurmasList()); // Chama a função existente e parsea a resposta
        const turmas = turmasResponse.success ? turmasResponse.data : [];
        if (!turmasResponse.success) {
            Logger.log("Falha ao obter lista de turmas para filtros: " + turmasResponse.message);
            // Continua, mas com turmas vazias
        }


        // Calcular próximas semanas (ex: 12 semanas a partir da próxima segunda)
        const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
        const weekStartDates = [];
        const numWeeks = 12; // Quantidade de semanas futuras para mostrar

        const today = new Date();
        today.setHours(0, 0, 0, 0);

        // Encontra a próxima segunda-feira
        const nextMonday = new Date(today.getTime());
        const currentDayOfWeek = nextMonday.getDay(); // 0=Domingo, 1=Segunda
        const daysUntilNextMonday = (currentDayOfWeek === 0) ? 1 : (8 - currentDayOfWeek) % 7; // Se hoje for domingo, próxima segunda é daqui 1 dia. Se for segunda, é hoje. Senão, (8 - dia_da_semana)%7
        nextMonday.setDate(nextMonday.getDate() + daysUntilNextMonday);
        nextMonday.setHours(0, 0, 0, 0); // Garantir hora zerada

        Logger.log("Gerando lista de semanas a partir de: " + Utilities.formatDate(nextMonday, timeZone, 'yyyy-MM-dd'));

        // Adiciona as datas das próximas 'numWeeks' segundas-feiras
        for (let i = 0; i < numWeeks; i++) {
            const weekStartDate = new Date(nextMonday.getTime());
            weekStartDate.setDate(nextMonday.getDate() + (i * 7));
            // Formata para um valor legível E um valor técnico (YYYY-MM-DD)
            const valueString = Utilities.formatDate(weekStartDate, timeZone, 'yyyy-MM-dd');
            const textString = Utilities.formatDate(weekStartDate, timeZone, 'dd/MM/yyyy'); // Formato exibição
            weekStartDates.push({ value: valueString, text: `Semana de ${textString}` });
        }

        // Adapta para o formato esperado pelo populateDropdown simples no JS (apenas value)
        const weekValueStrings = weekStartDates.map(week => week.value);

        Logger.log(`Filtros obtidos: ${turmas.length} turmas, ${weekValueStrings.length} semanas.`);

        return JSON.stringify({
            success: true,
            message: 'Filtros carregados.',
            data: {
                turmas: turmas,
                weekStartDates: weekValueStrings // Retorna apenas os valores YYYY-MM-DD
            }
        });

    } catch (e) {
        Logger.log('Erro em getScheduleViewFilters: ' + e.message + ' Stack: ' + e.stack);
        return JSON.stringify({ success: false, message: 'Ocorreu um erro ao obter os filtros de horários: ' + e.message, data: null });
    }
}

/**
 * Função exposta para o lado do cliente da ScheduleView.
 * Busca instâncias de horários filtradas por turma e data de início da semana.
 * Inclui informações adicionais de disciplina e professor real/original
 * buscando em outras planilhas para horários relevantes.
 * @param {string} turma A turma para filtrar.
 * @param {string} weekStartDateString A data de início da semana (Segunda-feira) no formato 'YYYY-MM-DD'.
 * @returns {string} JSON string {success, message, data: [lista de slots filtrados com detalhes]}
 */
function getFilteredScheduleInstances(turma, weekStartDateString) {
    Logger.log(`*** getFilteredScheduleInstances chamada para Turma: ${turma}, Semana: ${weekStartDateString} ***`);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const instancesSheet = ss.getSheetByName(SHEETS.SCHEDULE_INSTANCES);
    const baseSheet = ss.getSheetByName(SHEETS.BASE_SCHEDULES); // <-- Add this line
    const bookingsSheet = ss.getSheetByName(SHEETS.BOOKING_DETAILS); // <-- Add this line
    const timeZone = ss.getSpreadsheetTimeZone();

    // 1. Verifica autorização mínima
    const userEmail = Session.getActiveUser().getEmail();
    const userRole = getUserRolePlain(userEmail);
    if (!userRole) { // Qualquer usuário autorizado pode visualizar
        Logger.log('Erro: Usuário não autorizado a visualizar horários filtrados.');
        return JSON.stringify({ success: false, message: 'Usuário não autorizado a visualizar.', data: null });
    }

    // 2. Valida parâmetros de entrada
    if (!turma || typeof turma !== 'string' || turma.trim() === '') {
        Logger.log('Erro: Parâmetro turma inválido ou ausente.');
        return JSON.stringify({ success: false, message: 'Turma não especificada.', data: null });
    }
    if (!weekStartDateString || typeof weekStartDateString !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(weekStartDateString)) {
        Logger.log('Erro: Parâmetro weekStartDateString inválido ou ausente: ' + weekStartDateString);
        return JSON.stringify({ success: false, message: 'Semana de início inválida.', data: null });
    }

    // 3. Calcula o intervalo de datas da semana
    const parts = weekStartDateString.split('-');
    const weekStartDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10)); // Month is 0-indexed
    weekStartDate.setHours(0, 0, 0, 0); // Zera a hora para comparação

    // Verifica se a data parseada é válida e se realmente é uma segunda-feira
    if (isNaN(weekStartDate.getTime()) || weekStartDate.getDay() !== 1) { // 1 = Monday
        Logger.log('Erro: A data de início da semana fornecida não é uma Segunda-feira válida: ' + weekStartDateString);
        return JSON.stringify({ success: false, message: 'A data de início da semana deve ser uma Segunda-feira válida.', data: null });
    }

    const weekEndDate = new Date(weekStartDate.getTime());
    weekEndDate.setDate(weekEndDate.getDate() + 6); // Fim da semana (Domingo)
    weekEndDate.setHours(23, 59, 59, 999); // Define para o final do dia

    Logger.log(`Buscando instâncias para Turma "${turma}" entre ${Utilities.formatDate(weekStartDate, timeZone, 'dd/MM/yyyy')} e ${Utilities.formatDate(weekEndDate, timeZone, 'dd/MM/yyyy')}`);


    try {
        // 4. Lê os dados das planilhas necessárias
        if (!instancesSheet) {
            Logger.log(`Erro: Planilha "${SHEETS.SCHEDULE_INSTANCES}" não encontrada!`);
            return JSON.stringify({ success: false, message: `Erro interno: Planilha "${SHEETS.SCHEDULE_INSTANCES}" não encontrada.`, data: null });
        }
        if (!baseSheet) { // Check base sheet
            Logger.log(`Warning: Planilha "${SHEETS.BASE_SCHEDULES}" não encontrada. Disciplinas padrão podem não ser exibidas para horários disponíveis.`);
            // Continue, but log a warning
        }
        if (!bookingsSheet) { // Check bookings sheet
            Logger.log(`Warning: Planilha "${SHEETS.BOOKING_DETAILS}" não encontrada. Detalhes de reservas (disciplina real, profs reais) podem não ser exibidos.`);
            // Continue, but log a warning
        }


        const rawInstanceData = instancesSheet.getDataRange().getValues();

        if (rawInstanceData.length <= 1) {
            Logger.log('Planilha Instancias de Horarios está vazia ou apenas cabeçalho.');
            return JSON.stringify({ success: true, message: 'Nenhuma instância de horário futura encontrada.', data: [] });
        }

        const instanceHeader = rawInstanceData[0];
        const instanceData = rawInstanceData.slice(1);

        // Determine the number of columns based on the header
        const numInstanceCols = instanceHeader.length;

        // Check if essential instance columns exist
        const requiredInstanceCols = Math.max(
            HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA,
            HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO,
            HEADERS.SCHEDULE_INSTANCES.TURMA,
            HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL, // Original Prof
            HEADERS.SCHEDULE_INSTANCES.DATA,
            HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA,
            HEADERS.SCHEDULE_INSTANCES.HORA_INICIO,
            HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL,
            HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO,
            HEADERS.SCHEDULE_INSTANCES.ID_RESERVA // Needed to link to bookings
        ) + 1;

        if (numInstanceCols < requiredInstanceCols) {
            Logger.log(`Erro: Planilha "${SHEETS.SCHEDULE_INSTANCES}" não tem colunas suficientes (${numInstanceCols} vs ${requiredInstanceCols} requeridas) para a visualização. Verifique a estrutura.`);
            return JSON.stringify({ success: false, message: `Erro interno: Estrutura da planilha "${SHEETS.SCHEDULE_INSTANCES}" incorreta.`, data: null });
        }


        // --- Read and Map Base Schedules (for original discipline) ---
        const baseScheduleDisciplineMap = {};
        if (baseSheet) {
            const rawBaseData = baseSheet.getDataRange().getValues();
            const baseHeader = rawBaseData.length > 0 ? rawBaseData[0] : [];
            const baseData = rawBaseData.slice(1);
            const numBaseCols = baseHeader.length;
            const requiredBaseCols = Math.max(HEADERS.BASE_SCHEDULES.ID, HEADERS.BASE_SCHEDULES.DISCIPLINA_PADRAO) + 1;

            if (numBaseCols >= requiredBaseCols) {
                baseData.forEach(row => {
                    if (row && row.length >= requiredBaseCols) {
                        const baseIdRaw = row[HEADERS.BASE_SCHEDULES.ID];
                        const baseId = (typeof baseIdRaw === 'string' || typeof baseIdRaw === 'number') ? String(baseIdRaw).trim() : null;
                        const disciplinaPadraoRaw = row[HEADERS.BASE_SCHEDULES.DISCIPLINA_PADRAO];
                        const disciplinaPadrao = (typeof disciplinaPadraoRaw === 'string' || typeof disciplinaPadraoRaw === 'number') ? String(disciplinaPadraoRaw || '').trim() : '';
                        if (baseId) {
                            baseScheduleDisciplineMap[baseId] = disciplinaPadrao;
                        }
                    }
                });
                Logger.log(`Mapeadas ${Object.keys(baseScheduleDisciplineMap).length} disciplinas base.`);
            } else if (rawBaseData.length > 1) {
                Logger.log(`Warning: Planilha "${SHEETS.BASE_SCHEDULES}" não tem colunas suficientes (${numBaseCols} vs ${requiredBaseCols} requeridas). Disciplinas base podem não ser exibidas.`);
            }
        }


        // --- Read and Map Booking Details (for real discipline and professors) ---
        // Map instanceId -> bookingDetails (assuming one booking per instance for the relevant types)
        const bookingDetailsMap = {};
        if (bookingsSheet) {
            const rawBookingData = bookingsSheet.getDataRange().getValues();
            const bookingHeader = rawBookingData.length > 0 ? rawBookingData[0] : [];
            const bookingData = rawBookingData.slice(1);
            const numBookingCols = bookingHeader.length;
            const requiredBookingCols = Math.max(
                HEADERS.BOOKING_DETAILS.ID_INSTANCIA,
                HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL,
                HEADERS.BOOKING_DETAILS.PROFESSOR_REAL,
                HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL,
                HEADERS.BOOKING_DETAILS.STATUS_RESERVA // To filter for 'Agendada'
            ) + 1;

            if (numBookingCols >= requiredBookingCols) {
                bookingData.forEach(row => {
                    if (row && row.length >= requiredBookingCols) {
                        const instanceIdRaw = row[HEADERS.BOOKING_DETAILS.ID_INSTANCIA];
                        const instanceId = (typeof instanceIdRaw === 'string' || typeof instanceIdRaw === 'number') ? String(instanceIdRaw).trim() : null;
                        const statusReservaRaw = row[HEADERS.BOOKING_DETAILS.STATUS_RESERVA];
                        const statusReserva = (typeof statusReservaRaw === 'string' || typeof statusReservaRaw === 'number') ? String(statusReservaRaw).trim() : '';

                        // Only consider 'Agendada' bookings linked to an instance
                        if (instanceId && statusReserva === 'Agendada') {
                            const disciplinaRealRaw = row[HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL];
                            const disciplinaReal = (typeof disciplinaRealRaw === 'string' || typeof disciplinaRealRaw === 'number') ? String(disciplinaRealRaw || '').trim() : '';
                            const professorRealRaw = row[HEADERS.BOOKING_DETAILS.PROFESSOR_REAL];
                            const professorReal = (typeof professorRealRaw === 'string' || typeof professorRealRaw === 'number') ? String(professorRealRaw || '').trim() : '';
                            const professorOriginalBookingRaw = row[HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL];
                            const professorOriginalBooking = (typeof professorOriginalBookingRaw === 'string' || typeof professorOriginalBookingRaw === 'number') ? String(professorOriginalBookingRaw || '').trim() : '';

                            // Store the details, prioritizing if multiple bookings exist for the same instance (unlikely with current system)
                            bookingDetailsMap[instanceId] = {
                                disciplinaReal: disciplinaReal,
                                professorReal: professorReal,
                                professorOriginalBooking: professorOriginalBooking // Original Prof from the booking detail row
                            };
                        }
                    }
                });
                Logger.log(`Mapeadas ${Object.keys(bookingDetailsMap).length} detalhes de reservas agendadas.`);
            } else if (rawBookingData.length > 1) {
                Logger.log(`Warning: Planilha "${SHEETS.BOOKING_DETAILS}" não tem colunas suficientes (${numBookingCols} vs ${requiredBookingCols} requeridas). Detalhes de reservas podem não ser exibidos.`);
            }
        }


        // 5. Filtra e enriquece os slots
        const filteredSlots = [];

        for (let i = 0; i < instanceData.length; i++) {
            const row = instanceData[i];
            const rowIndex = i + 2; // Índice real na planilha

            // Basic check if row has enough columns based on header
            if (!row || row.length < numInstanceCols) {
                // Logger.log(`Skipping incomplete instance row ${rowIndex}.`); // Too noisy
                continue;
            }

            // Extract & format essential data from instance row
            const instanceIdRaw = row[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA];
            const baseIdRaw = row[HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO]; // Needed to lookup base discipline
            const instanceTurmaRaw = row[HEADERS.SCHEDULE_INSTANCES.TURMA];
            const professorPrincipalInstanceRaw = row[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL]; // Original Prof from instance
            const rawInstanceDate = row[HEADERS.SCHEDULE_INSTANCES.DATA];
            const instanceDiaSemanaRaw = row[HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA];
            const rawHoraInicio = row[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO];
            const originalTypeRaw = row[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL];
            const instanceStatusRaw = row[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO];
            const instanceReservationIdRaw = row[HEADERS.SCHEDULE_INSTANCES.ID_RESERVA]; // Needed to link to bookings

            const instanceId = (typeof instanceIdRaw === 'string' || typeof instanceIdRaw === 'number') ? String(instanceIdRaw).trim() : null;
            const baseId = (typeof baseIdRaw === 'string' || typeof baseIdRaw === 'number') ? String(baseIdRaw).trim() : null;
            const instanceTurma = (typeof instanceTurmaRaw === 'string' || typeof instanceTurmaRaw === 'number') ? String(instanceTurmaRaw).trim() : null;
            const professorPrincipalInstance = (typeof professorPrincipalInstanceRaw === 'string' || typeof professorPrincipalInstanceRaw === 'number') ? String(professorPrincipalInstanceRaw || '').trim() : ''; // Original Prof from instance
            const instanceDate = formatValueToDate(rawInstanceDate);
            const instanceDiaSemana = (typeof instanceDiaSemanaRaw === 'string' || typeof instanceDiaSemanaRaw === 'number') ? String(instanceDiaSemanaRaw).trim() : null;
            const formattedHoraInicio = formatValueToHHMM(rawHoraInicio, timeZone);
            const originalType = (typeof originalTypeRaw === 'string' || typeof originalTypeRaw === 'number') ? String(originalTypeRaw).trim() : null;
            const instanceStatus = (typeof instanceStatusRaw === 'string' || typeof instanceStatusRaw === 'number') ? String(instanceStatusRaw).trim() : null;
            // We don't strictly need instanceReservationIdRaw here, as we map booking details by instance ID


            // Validate essential data for filtering and display
            if (!instanceId || !instanceTurma || !instanceDate || !instanceDiaSemana || formattedHoraInicio === null || !originalType || !instanceStatus || !baseId) { // baseId is now essential too
                // Logger.log(`Skipping instance row ${rowIndex} due to invalid essential data.`); // Too noisy
                continue;
            }


            // Filter by Turma
            if (instanceTurma !== turma.trim()) {
                continue; // Skip if turma does not match
            }

            // Filter by Date range
            // Compare instanceDate (already zered hour by formatValueToDate if valid Date)
            if (instanceDate < weekStartDate || instanceDate > weekEndDate) {
                continue; // Skip if date is outside the week
            }

            // --- Enrich the slot with additional details ---
            let disciplinaParaExibir = '';
            let professorParaExibir = '';
            let professorOriginalNaReserva = ''; // This is professor_original from booking details, relevant for Substituicao

            const bookingDetails = bookingDetailsMap[instanceId]; // Try to find matching booking details

            if (instanceStatus === STATUS_OCUPACAO.DISPONIVEL) {
                // For available slots, use discipline and professor from the base schedule
                disciplinaParaExibir = baseScheduleDisciplineMap[baseId] || ''; // Lookup discipline using baseId
                professorParaExibir = professorPrincipalInstance; // Use original professor from instance
            } else if (bookingDetails) { // If there are booking details linked AND the instance is booked
                // For booked slots, use discipline and real/original professor from booking details
                disciplinaParaExibir = bookingDetails.disciplinaReal;
                professorParaExibir = bookingDetails.professorReal;
                professorOriginalNaReserva = bookingDetails.professorOriginalBooking; // Store original prof from booking details
            } else {
                // Fallback for booked slots without matching booking details (inconsistent data)
                Logger.log(`Warning: Instância ${instanceId} (Status: ${instanceStatus}) não encontrou detalhes de reserva correspondentes no mapa. Row ${rowIndex}.`);
                // We could use base info as a fallback, or leave empty
                disciplinaParaExibir = baseScheduleDisciplineMap[baseId] || ''; // Fallback to base discipline
                professorParaExibir = professorPrincipalInstance; // Fallback to base professor
            }


            // If passed all filters and enriched, add to results
            filteredSlots.push({
                idInstancia: instanceId,
                data: Utilities.formatDate(instanceDate, timeZone, 'dd/MM/yyyy'), // Formatted date for display
                diaSemana: instanceDiaSemana,
                horaInicio: formattedHoraInicio,
                turma: instanceTurma, // Turma from instance (should match filter turma)
                tipoOriginal: originalType,
                statusOcupacao: instanceStatus,

                // Add the enriched details
                disciplinaParaExibir: disciplinaParaExibir,
                professorParaExibir: professorParaExibir, // Real Professor for booked, Original for available (unless overridden by booking)
                professorOriginalNaReserva: professorOriginalNaReserva // Professor ORIGINAL from booking details (only if Substitution booked)
                // We could also include professorPrincipalInstance if needed for fallback logic in frontend
            });
        }

        Logger.log(`Encontrados ${filteredSlots.length} slots filtrados e enriquecidos para Turma "${turma}" na semana de ${weekStartDateString}.`);

        // Return JSON with success and the list of enriched slots.
        return JSON.stringify({ success: true, message: `${filteredSlots.length} horários encontrados.`, data: filteredSlots });

    } catch (e) {
        Logger.log('Erro em getFilteredScheduleInstances: ' + e.message + ' Stack: ' + e.stack);
        return JSON.stringify({ success: false, message: 'Ocorreu um erro interno ao buscar horários: ' + e.message, data: null });
    }
}


/**
 * Função exposta para ser chamada pelo lado do cliente.
 * Busca e retorna as instâncias de horários disponíveis para um determinado tipo de reserva (Reposição ou Substituição).
 * Os resultados são ordenados por Data, Hora e Turma.
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
            HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR
        ) + 1; // +1 porque os índices são base 0.

        // Obtém todos os dados da planilha.
        const rawData = sheet.getDataRange().getValues();
        // Verifica se há dados além do cabeçalho.
        if (rawData.length <= 1) {
            Logger.log('Planilha Instancias de Horarios está vazia ou apenas cabeçalho.');
            return JSON.stringify({ success: true, message: 'Nenhuma instância de horário futura encontrada. Gere instâncias primeiro.', data: [] });
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
                // Logger.log(`Skipping incomplete row ${rowIndex} in Instancias de Horarios.`); // Too noisy
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
            const instanceId = (typeof instanceIdRaw === 'string' || typeof instanceIdRaw === 'number') ? String(instanceIdRaw).trim() : null;
            const baseId = (typeof baseIdRaw === 'string' || typeof baseIdRaw === 'number') ? String(baseIdRaw).trim() : null;
            const turma = (typeof turmaRaw === 'string' || typeof turmaRaw === 'number') ? String(turmaRaw).trim() : null;
            const professorPrincipal = (typeof professorPrincipalRaw === 'string' || typeof professorPrincipalRaw === 'number') ? String(professorPrincipalRaw || '').trim() : '';
            const instanceDate = formatValueToDate(rawDate); // <-- THIS IS THE Date OBJECT
            const instanceDiaSemana = (typeof instanceDiaSemanaRaw === 'string' || typeof instanceDiaSemanaRaw === 'number') ? String(instanceDiaSemanaRaw).trim() : null;
            const formattedHoraInicio = formatValueToHHMM(rawHoraInicio, timeZone); // <-- HH:mm STRING
            const originalType = (typeof originalTypeRaw === 'string' || typeof originalTypeRaw === 'number') ? String(originalTypeRaw).trim() : null;
            const instanceStatus = (typeof instanceStatusRaw === 'string' || typeof instanceStatusRaw === 'number') ? String(instanceStatusRaw).trim() : null;

            // Verifica se algum dado essencial formatado é inválido ou ausente.
            if (!instanceId || instanceId === '' || !baseId || baseId === '' || !turma || turma === '' ||
                !instanceDate || // Checks if instanceDate is a valid Date object
                !instanceDiaSemana || instanceDiaSemana === '' ||
                formattedHoraInicio === null || // Checks if time is valid string
                !originalType || originalType === '' ||
                !instanceStatus || instanceStatus === '') {
                // Log detail for debugging skipped rows
                // Logger.log(`Skipping row ${rowIndex} due to invalid/missing essential data: ID=${instanceIdRaw}, BaseID=${baseIdRaw}, Turma=${turmaRaw}, ProfPrinc=${professorPrincipalRaw}, Data=${rawDate}, Dia=${instanceDiaSemanaRaw}, Hora=${rawHoraInicio}, Tipo=${originalTypeRaw}, Status=${instanceStatusRaw}`);
                continue; // Pula para a próxima linha.
            }

            // Validações adicionais de consistência dos dados (Dias da semana, Tipos, Status)
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
            // instanceDate already has hours/minutes/seconds set to 0 by formatValueToDate if it was parsed correctly
            if (instanceDate < today) {
                // Pula instâncias que já ocorreram.
                continue;
            }

            // Lógica de filtragem baseada no tipo de reserva solicitado e status.
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
                        instanceDateObj: instanceDate, // <-- INCLUDE THE ACTUAL DATE OBJECT FOR SORTING
                        diaSemana: instanceDiaSemana,
                        horaInicio: formattedHoraInicio, // HH:mm string for sorting
                        tipoOriginal: originalType,
                        statusOcupacao: instanceStatus, // Será 'Disponivel'.
                    });
                }
            } else if (tipoReserva === TIPOS_RESERVA.SUBSTITUICAO) {
                // Para Substituição, o horário deve ser do tipo FIXO.
                // E NÃO pode estar com status REPOSICAO_AGENDADA.
                // Pode estar DISPONIVEL ou já ter uma SUBSTITUICAO_AGENDADA.
                if (originalType === TIPOS_HORARIO.FIXO && instanceStatus !== STATUS_OCUPACAO.REPOSICAO_AGENDADA) {
                    // Adiciona o horário à lista de disponíveis para substituição.
                    availableSlots.push({
                        idInstancia: instanceId,
                        baseId: baseId,
                        turma: turma,
                        professorPrincipal: professorPrincipal, // Professor original do horário fixo.
                        data: Utilities.formatDate(instanceDate, timeZone, 'dd/MM/yyyy'), // Formata data para exibição.
                        instanceDateObj: instanceDate, // <-- INCLUDE THE ACTUAL DATE OBJECT FOR SORTING
                        diaSemana: instanceDiaSemana,
                        horaInicio: formattedHoraInicio, // HH:mm string for sorting
                        tipoOriginal: originalType,
                        statusOcupacao: instanceStatus, // Puede ser 'Disponivel' o 'Substituicao Agendada'.
                    });
                }
            }
        }

        // --- ADD SORTING LOGIC HERE ---
        availableSlots.sort((a, b) => {
            // 1. Sort by Date (ascending)
            // Use getTime() for reliable comparison of Date objects
            const dateComparison = a.instanceDateObj.getTime() - b.instanceDateObj.getTime();
            if (dateComparison !== 0) {
                return dateComparison;
            }

            // 2. If Dates are the same, sort by Time (ascending string comparison)
            // String comparison for HH:mm format works correctly (e.g., "09:00" < "10:00")
            const timeComparison = a.horaInicio.localeCompare(b.horaInicio);
            if (timeComparison !== 0) {
                return timeComparison;
            }

            // 3. If Date and Time are the same, sort by Turma (ascending string comparison)
            const turmaComparison = a.turma.localeCompare(b.turma);
            return turmaComparison;
        });
        // --- END SORTING LOGIC ---


        Logger.log('Número de slots disponíveis encontrados e ordenados: ' + availableSlots.length);

        // Return JSON with success and the sorted list of available slots.
        // Note: We are sending the instanceDateObj back in the data payload, which is fine.
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
 * criando um registro na planilha de detalhes da reserva e um evento no Google
 * Calendar. Usa LockService para evitar condições de corrida (duas pessoas
 * tentando reservar o mesmo horário).
 * @param {string} jsonBookingDetailsString Uma string JSON contendo os detalhes
 *     da reserva (idInstancia, tipoReserva, professorReal, disciplinaReal,
 *     etc.).
 * @returns {string} Uma string JSON indicando sucesso ou falha da
 *     operação {success, message, data: {bookingId, eventId}}.
 */
function bookSlot(jsonBookingDetailsString) {
    // Obtém um bloqueio exclusivo para este script, esperando até 10 segundos se
    // já estiver bloqueado. Isso previne que duas execuções simultâneas desta
    // função tentem modificar o mesmo horário ao mesmo tempo.
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);  // Espera até 10 segundos (10000 ms).

    // Verifica a autorização do usuário.
    const userEmail = Session.getActiveUser().getEmail();
    const userRole = getUserRolePlain(userEmail);

    // Se não autorizado, libera o bloqueio e retorna falha.
    if (!userRole) {
        lock.releaseLock();
        return JSON.stringify({
            success: false,
            message: 'Usuário não autorizado a agendar.',
            data: null
        });
    }

    let bookingDetails;
    try {
        // Tenta converter a string JSON recebida em um objeto JavaScript.
        bookingDetails = JSON.parse(jsonBookingDetailsString);
        Logger.log(
            'Booking details received and parsed: ' +
            JSON.stringify(bookingDetails));
    } catch (e) {
        // Se o JSON for inválido, libera o bloqueio e retorna falha.
        lock.releaseLock();
        Logger.log('Erro ao parsear JSON de detalhes da reserva: ' + e.message);
        return JSON.stringify({
            success: false,
            message: 'Erro ao processar dados da reserva.',
            data: null
        });
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
        if (!bookingDetails || typeof bookingDetails.idInstancia !== 'string' ||
            bookingDetails.idInstancia.trim() === '') {
            lock.releaseLock();
            return JSON.stringify({
                success: false,
                message:
                    'Dados de ID da instância de horário incompletos ou inválidos.',
                data: null
            });
        }

        // Obtém e limpa os dados essenciais da reserva.
        const instanceIdToBook = bookingDetails.idInstancia.trim();
        const bookingType = bookingDetails.tipoReserva ?
            String(bookingDetails.tipoReserva).trim() :
            null;
        // Valida o tipo de reserva.
        if (!bookingType ||
            (bookingType !== TIPOS_RESERVA.REPOSICAO &&
                bookingType !== TIPOS_RESERVA.SUBSTITUICAO)) {
            lock.releaseLock();
            return JSON.stringify({
                success: false,
                message: 'Tipo de reserva inválido ou ausente.',
                data: null
            });
        }

        // --- Busca e Validação da Instância de Horário na Planilha ---
        // Obtém todos os dados da planilha de instâncias.
        const instanceDataRaw = instancesSheet.getDataRange().getValues();
        let instanceRowIndex =
            -1;  // Índice da linha onde a instância foi encontrada.
        let instanceDetails = null;  // Array com os dados da linha da instância.

        // Verifica se a planilha de instâncias tem dados.
        if (instanceDataRaw.length <= 1) {
            lock.releaseLock();
            return JSON.stringify({
                success: false,
                message:
                    'Erro interno: Planilha de instâncias vazia ou estrutura incorreta.',
                data: null
            });
        }

        // Remove o cabeçalho.
        const instanceData = instanceDataRaw.slice(1);

        // Itera pelas linhas de instância para encontrar a que corresponde ao ID
        // solicitado.
        for (let i = 0; i < instanceData.length; i++) {
            const row = instanceData[i];
            const rowIndex = i + 2;  // Índice real na planilha.
            // Verifica se a linha tem pelo menos a coluna do ID.
            const minColsForId = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA + 1;
            if (row && row.length >= minColsForId) {
                const currentInstanceIdRaw =
                    row[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA];
                // Converte e compara o ID da linha atual com o ID procurado.
                const currentInstanceId = (typeof currentInstanceIdRaw === 'string' ||
                    typeof currentInstanceIdRaw === 'number') ?
                    String(currentInstanceIdRaw).trim() :
                    null;
                if (currentInstanceId && currentInstanceId === instanceIdToBook) {
                    instanceRowIndex = rowIndex;  // Armazena o índice da linha.
                    instanceDetails = row;        // Armazena os dados da linha.
                    break;                        // Para o loop assim que encontrar.
                }
            }
        }

        // Se a instância não foi encontrada (pode ter sido deletada ou o ID estava
        // errado).
        if (instanceRowIndex === -1 || !instanceDetails) {
            lock.releaseLock();
            // Mensagem importante para o usuário indicando possível concorrência ou
            // dado desatualizado.
            return JSON.stringify({
                success: false,
                message:
                    'Este horário não está mais disponível. Por favor, atualize a lista e tente novamente.',
                data: null
            });
        }

        // Verifica se a linha encontrada tem o número esperado de colunas para
        // evitar erros.
        const expectedInstanceCols =
            Math.max(
                HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO,
                HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL,
                HEADERS.SCHEDULE_INSTANCES.DATA,
                HEADERS.SCHEDULE_INSTANCES.HORA_INICIO,
                HEADERS.SCHEDULE_INSTANCES.TURMA,
                HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL,
                HEADERS.SCHEDULE_INSTANCES
                    .ID_EVENTO_CALENDAR  // Inclui a coluna do ID do evento.
            ) +
            1;

        if (instanceDetails.length < expectedInstanceCols) {
            Logger.log(`Erro: Linha ${instanceRowIndex} na planilha Instancias de Horarios tem menos colunas (${instanceDetails.length}) que o esperado (${expectedInstanceCols}).`);
            lock.releaseLock();
            return JSON.stringify({
                success: false,
                message:
                    'Erro interno: Dados do horário selecionado incompletos na planilha.',
                data: null
            });
        }

        // Extrai e formata os dados relevantes da linha da instância encontrada.
        const currentStatusRaw =
            instanceDetails[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO];
        const originalTypeRaw =
            instanceDetails[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL];
        const rawBookingDate = instanceDetails[HEADERS.SCHEDULE_INSTANCES.DATA];
        const rawBookingTime =
            instanceDetails[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO];
        const turmaInstanciaRaw = instanceDetails[HEADERS.SCHEDULE_INSTANCES.TURMA];
        const professorPrincipalInstanciaRaw =
            instanceDetails[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL];
        const calendarEventIdExistingRaw =
            instanceDetails[HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR];

        // Formata e valida os dados extraídos.
        const currentStatus = (typeof currentStatusRaw === 'string' ||
            typeof currentStatusRaw === 'number') ?
            String(currentStatusRaw).trim() :
            null;
        const originalType = (typeof originalTypeRaw === 'string' ||
            typeof originalTypeRaw === 'number') ?
            String(originalTypeRaw).trim() :
            null;
        const turmaInstancia = (typeof turmaInstanciaRaw === 'string' ||
            typeof turmaInstanciaRaw === 'number') ?
            String(turmaInstanciaRaw).trim() :
            null;
        const professorPrincipalInstancia =
            (typeof professorPrincipalInstanciaRaw === 'string' ||
                typeof professorPrincipalInstanciaRaw === 'number') ?
                String(professorPrincipalInstanciaRaw || '').trim() :
                '';
        const bookingDateObj = formatValueToDate(rawBookingDate);
        const bookingHourString = formatValueToHHMM(rawBookingTime, timeZone);
        const calendarEventIdExisting =
            (typeof calendarEventIdExistingRaw === 'string' ||
                typeof calendarEventIdExistingRaw === 'number') ?
                String(calendarEventIdExistingRaw || '').trim() :
                null;

        // Verifica se dados críticos formatados são válidos.
        if (!currentStatus || currentStatus === '' || !originalType ||
            originalType === '' || !turmaInstancia || turmaInstancia === '' ||
            !bookingDateObj || bookingHourString === null) {
            Logger.log(`Erro: Dados críticos da instância ${instanceIdToBook} na linha ${instanceRowIndex} são inválidos. Status=${currentStatusRaw}, Tipo=${originalTypeRaw}, Turma=${turmaInstanciaRaw}, Data=${rawBookingDate}, Hora=${rawBookingTime}`);
            lock.releaseLock();
            return JSON.stringify({
                success: false,
                message:
                    'Erro interno: Dados do horário selecionado são inválidos na planilha.',
                data: null
            });
        }

        // --- Lógica de Validação Específica por Tipo de Reserva (VERIFICAÇÃO DE
        // CONCORRÊNCIA) ---
        if (bookingType === TIPOS_RESERVA.REPOSICAO) {
            // Para REPOSICAO, a instância deve ser do tipo VAGO e estar DISPONIVEL no
            // momento da reserva.
            if (originalType !== TIPOS_HORARIO.VAGO ||
                currentStatus !== STATUS_OCUPACAO.DISPONIVEL) {
                lock.releaseLock();
                // Mensagem indicando que o status mudou desde que o usuário viu a
                // lista.
                return JSON.stringify({
                    success: false,
                    message:
                        'Este horário não está mais disponível para reposição ou não é um horário vago (concorrência).',
                    data: null
                });
            }
            // Verifica se os campos obrigatórios para reposição foram preenchidos no
            // formulário.
            if (!bookingDetails.professorReal ||
                bookingDetails.professorReal.trim() === '' ||
                !bookingDetails.disciplinaReal ||
                bookingDetails.disciplinaReal.trim() === '') {
                lock.releaseLock();
                return JSON.stringify({
                    success: false,
                    message:
                        'Por favor, preencha todos os campos obrigatórios para reposição (Professor, Disciplina).',
                    data: null
                });
            }

        } else if (bookingType === TIPOS_RESERVA.SUBSTITUICAO) {
            // Para SUBSTITUICAO, a instância deve ser do tipo FIXO.
            if (originalType !== TIPOS_HORARIO.FIXO) {
                lock.releaseLock();
                return JSON.stringify({
                    success: false,
                    message:
                        'Este horário não é um horário fixo e não pode ser substituído.',
                    data: null
                });
            }
            // Não pode ser substituído se já houver uma REPOSICAO agendada nele.
            if (currentStatus === STATUS_OCUPACAO.REPOSICAO_AGENDADA) {
                lock.releaseLock();
                return JSON.stringify({
                    success: false,
                    message:
                        'Este horário fixo está sendo usado para uma reposição e não pode ser substituído.',
                    data: null
                });
            }
            // Deve estar DISPONIVEL ou já marcado como SUBSTITUICAO_AGENDADA
            // (permitindo reagendar/atualizar a substituição). Se o status for
            // qualquer outro (ex: algum status futuro inválido), a reserva falha.
            if (currentStatus !== STATUS_OCUPACAO.DISPONIVEL &&
                currentStatus !== STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA) {
                lock.releaseLock();
                return JSON.stringify({
                    success: false,
                    message:
                        'Este horário fixo não está disponível para substituição neste momento (concorrência).',
                    data: null
                });
            }

            // Verifica campos obrigatórios para substituição.
            if (!bookingDetails.professorReal ||
                bookingDetails.professorReal.trim() === '' ||
                !bookingDetails.disciplinaReal ||
                bookingDetails.disciplinaReal.trim() === '') {
                lock.releaseLock();
                return JSON.stringify({
                    success: false,
                    message:
                        'Por favor, preencha todos os campos obrigatórios para substituição (Professor Substituto, Disciplina).',
                    data: null
                });
            }

            // Para substituição, é crucial que o professor original esteja definido
            // na instância.
            if (professorPrincipalInstancia === '') {
                Logger.log(`Erro: Instância de horário fixo ${instanceIdToBook} na linha ${instanceRowIndex} não tem Professor Principal definido na planilha de instâncias.`);
                lock.releaseLock();
                return JSON.stringify({
                    success: false,
                    message:
                        'Erro interno: Horário fixo não tem Professor Principal definido na planilha de instâncias. Verifique a geração de instâncias.',
                    data: null
                });
            }
        }

        // --- Se todas as validações passaram, prossegue com a reserva ---

        // Gera um ID único para a nova reserva.
        const bookingId = Utilities.getUuid();
        // Obtém a data/hora atual para registro.
        const now = new Date();

        // Determina o novo status da instância baseado no tipo de reserva.
        const newStatus = (bookingType === TIPOS_RESERVA.REPOSICAO) ?
            STATUS_OCUPACAO.REPOSICAO_AGENDADA :
            STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA;
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

            // Garante que a linha a ser escrita tenha exatamente o número correto de
            // colunas.
            while (updatedInstanceRow.length < numColsInstance)
                updatedInstanceRow.push('');  // Adiciona colunas vazias se faltar.
            if (updatedInstanceRow.length > numColsInstance)
                updatedInstanceRow.length =
                    numColsInstance;  // Remove colunas extras se houver.

            // Escreve a linha atualizada de volta na planilha, na posição correta.
            instancesSheet.getRange(instanceRowIndex, 1, 1, numColsInstance)
                .setValues([updatedInstanceRow]);
            Logger.log(`Instância de horário ${instanceIdToBook} na linha ${instanceRowIndex} atualizada para ${newStatus}.`);
        } catch (e) {
            // Se ocorrer um erro ao escrever na planilha de instâncias.
            Logger.log(`Erro ao atualizar linha ${instanceRowIndex} na planilha "${SHEETS.SCHEDULE_INSTANCES}": ${e.message}`);
            lock.releaseLock();
            return JSON.stringify({
                success: false,
                message: `Erro interno ao atualizar o status do horário na planilha. ${e.message}`,
                data: null
            });
        }

        // --- Cria a Nova Linha na Planilha de Detalhes da Reserva ---
        const newBookingRow = [];  // Array para a nova linha.

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
        newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_REAL] =
            bookingDetails.professorReal.trim();
        // Professor original só é relevante para substituição.
        newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL] =
            (bookingType === TIPOS_RESERVA.SUBSTITUICAO) ?
                professorPrincipalInstancia.trim() :
                '';
        newBookingRow[HEADERS.BOOKING_DETAILS.ALUNOS] = bookingDetails.alunos ?
            bookingDetails.alunos.trim() :
            '';  // Campo opcional.
        newBookingRow[HEADERS.BOOKING_DETAILS.TURMAS_AGENDADA] =
            turmaInstancia;  // Usa a turma da instância por padrão.
        newBookingRow[HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL] =
            bookingDetails.disciplinaReal.trim();

        // Monta o objeto Date/Time completo para o início efetivo da aula.
        const [hour, minute] = bookingHourString.split(':').map(Number);
        bookingDateObj.setHours(
            hour, minute, 0, 0);  // Define a hora e minuto no objeto Date.
        newBookingRow[HEADERS.BOOKING_DETAILS.DATA_HORA_INICIO_EFETIVA] =
            bookingDateObj;

        newBookingRow[HEADERS.BOOKING_DETAILS.STATUS_RESERVA] =
            'Agendada';  // Status inicial.
        newBookingRow[HEADERS.BOOKING_DETAILS.DATA_CRIACAO] =
            now;  // Data/hora da criação.
        newBookingRow[HEADERS.BOOKING_DETAILS.CRIADO_POR] =
            userEmail;  // Email do usuário que fez a reserva.

        // --- Adiciona a Linha à Planilha de Reservas ---
        try {
            // Verificação extra para garantir o número correto de colunas antes de
            // adicionar.
            if (newBookingRow.length !== numColsBooking) {
                Logger.log(`Erro interno: newBookingRow tem ${newBookingRow.length} colunas, esperado ${numColsBooking}. Ajustando...`);
                // Ajusta o array se necessário (embora a inicialização acima deva
                // prevenir isso).
                while (newBookingRow.length < numColsBooking) newBookingRow.push('');
                if (newBookingRow.length > numColsBooking)
                    newBookingRow.length = numColsBooking;
            }
            // Adiciona a nova linha ao final da planilha de reservas.
            bookingsSheet.appendRow(newBookingRow);
            Logger.log(
                `Reserva ${bookingId} adicionada à planilha de Reservas Detalhadas.`);
        } catch (e) {
            // Se falhar ao adicionar a linha de reserva (mas a instância já foi
            // atualizada).
            Logger.log(`Erro ao adicionar reserva ${bookingId} à planilha "${SHEETS.BOOKING_DETAILS}": ${e.message}`);
            // É um estado inconsistente, mas informa o usuário. A instância ficou
            // reservada, mas os detalhes não foram salvos. Idealmente, deveria tentar
            // reverter a atualização da instância (rollback), mas isso adiciona
            // complexidade.
            lock.releaseLock();
            return JSON.stringify({
                success: false,
                message:
                    `Reserva agendada na instância, mas erro ao salvar os detalhes da reserva. ${e.message}`,
                data: null
            });
        }

        // --- Integração com Google Calendar ---
        let calendarEventId =
            null;  // Variável para armazenar o ID do evento criado/atualizado.
        try {
            // Obtém o ID do calendário da planilha de configurações.
            const calendarId = getConfigValue('ID do Calendario');
            // Se o ID não estiver configurado, pula a criação do evento.
            if (!calendarId || calendarId === '') {
                Logger.log(
                    'ID do Calendário não configurado. Pulando criação de evento.');
                lock.releaseLock();
                // Retorna sucesso, mas informa que o evento não foi criado.
                return JSON.stringify({
                    success: true,
                    message:
                        `Reserva agendada com sucesso, mas o ID do calendário não está configurado. Evento não criado/atualizado.`,
                    data: { bookingId: bookingId, eventId: null }
                });
            }
            // Tenta obter o objeto Calendar usando o ID.
            const calendar = CalendarApp.getCalendarById(calendarId);
            // Se o calendário não for encontrado ou o script não tiver permissão.
            if (!calendar) {
                Logger.log(`Calendário com ID "${calendarId}" não encontrado ou acessível. Pulando criação/atualização de evento.`);
                lock.releaseLock();
                // Retorna sucesso, mas informa sobre o problema com o calendário.
                return JSON.stringify({
                    success: true,
                    message: `Reserva agendada com sucesso, mas o calendário "${calendarId}" não foi encontrado ou não está acessível. Evento não criado/atualizado.`,
                    data: { bookingId: bookingId, eventId: null }
                });
            }

            // Define a duração padrão da aula (em minutos).
            let durationMinutes = 45;  // Valor default.
            // Tenta obter a duração da configuração.
            const durationConfig = getConfigValue('Duracao Padrao Aula (minutos)');
            if (durationConfig && !isNaN(parseInt(durationConfig))) {
                durationMinutes = parseInt(durationConfig);
            } else {
                Logger.log(
                    `Configuração "Duracao Padrao Aula (minutos)" não encontrada ou inválida. Usando padrão de ${durationMinutes} minutos.`);
            }

            // Calcula a hora de início e fim do evento.
            const startTime = bookingDateObj;  // Já contém data e hora corretas.
            const endTime = new Date(
                startTime.getTime() +
                durationMinutes * 60 * 1000);  // Adiciona a duração em milissegundos.

            // Define o título e a descrição do evento.
            let eventTitle = '';
            let eventDescription = `Reserva ID: ${bookingId}\nTipo: ${bookingType}\nCriado por: ${userEmail}`;

            // Usa os dados já formatados da newBookingRow.
            const disciplina =
                newBookingRow[HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL] ||
                'Disciplina Não Informada';
            const turmaTexto = newBookingRow[HEADERS.BOOKING_DETAILS.TURMAS_AGENDADA];

            eventDescription += `\nTurma(s): ${turmaTexto}`;
            // Título mais informativo.
            eventTitle = `${bookingType} - ${disciplina} - ${turmaTexto}`;

            // --- Prepara a lista de convidados para o evento ---
            const guests = [];  // Array de emails dos convidados.
            const authUsersSheet = ss.getSheetByName(SHEETS.AUTHORIZED_USERS);
            const nameEmailMap =
                {};  // Mapa para buscar email pelo nome do professor.
            // Tenta ler a planilha de usuários para mapear nomes a emails.
            if (authUsersSheet) {
                const authUserData = authUsersSheet.getDataRange().getValues();
                // Verifica se a planilha tem dados e as colunas necessárias.
                if (authUserData.length > 1 &&
                    authUserData[0].length > Math.max(
                        HEADERS.AUTHORIZED_USERS.EMAIL,
                        HEADERS.AUTHORIZED_USERS.NOME)) {
                    // Cria o mapa Nome -> Email.
                    for (let i = 1; i < authUserData.length; i++) {
                        const row = authUserData[i];
                        const email =
                            (row.length > HEADERS.AUTHORIZED_USERS.EMAIL &&
                                typeof row[HEADERS.AUTHORIZED_USERS.EMAIL] === 'string') ?
                                row[HEADERS.AUTHORIZED_USERS.EMAIL].trim() :
                                '';
                        const name =
                            (row.length > HEADERS.AUTHORIZED_USERS.NOME &&
                                typeof row[HEADERS.AUTHORIZED_USERS.NOME] === 'string') ?
                                row[HEADERS.AUTHORIZED_USERS.NOME].trim() :
                                '';
                        if (email && name) nameEmailMap[name] = email;
                    }
                } else {
                    Logger.log(
                        'Planilha Usuarios Autorizados vazia ou estrutura incorreta para buscar emails.');
                }

                // Adiciona o email do professor real (que vai dar a aula).
                const profRealNome =
                    newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_REAL];
                if (profRealNome && nameEmailMap[profRealNome])
                    guests.push(nameEmailMap[profRealNome]);

                // Se for substituição, adiciona também o email do professor original.
                if (bookingType === TIPOS_RESERVA.SUBSTITUICAO) {
                    const profOriginalNome =
                        newBookingRow[HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL];
                    if (profOriginalNome && nameEmailMap[profOriginalNome])
                        guests.push(nameEmailMap[profOriginalNome]);
                }

            } else {
                Logger.log(
                    'Planilha Usuarios Autorizados não encontrada para adicionar convidados.');
            }

            // --- Lógica para Atualizar ou Criar Evento ---
            let event = null;
            // Verifica se já existe um ID de evento associado a esta instância na
            // planilha.
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
                    const newGuests =
                        [...new Set(guests)];  // Remove duplicatas da nova lista.

                    // Remove convidados que não estão mais na lista nova.
                    existingGuests.forEach(guestEmail => {
                        if (!newGuests.includes(guestEmail)) {
                            try {
                                event.removeGuest(guestEmail);
                            } catch (removeErr) {
                                Logger.log(
                                    `Falha ao remover convidado ${guestEmail}: ${removeErr}`);
                            }
                        }
                    });

                    // Adiciona convidados que estão na lista nova mas não estavam na
                    // antiga.
                    newGuests.forEach(guestEmail => {
                        if (!existingGuests.includes(guestEmail)) {
                            try {
                                event.addGuest(guestEmail);
                            } catch (addErr) {
                                Logger.log(
                                    `Falha ao adicionar convidado ${guestEmail}: ${addErr}`);
                            }
                        }
                    });

                } catch (e) {
                    // Se getEventById falhar (evento deletado, ID inválido, permissão?).
                    Logger.log(`Evento do Calendar ID ${calendarEventIdExisting} não encontrado para atualização (pode ter sido excluído manualmente ou ID inválido): ${e}. Criando novo evento.`);
                    event = null;  // Reseta a variável para forçar a criação de um novo
                    // evento.
                }
            }

            // Se não havia evento existente ou a busca/atualização falhou.
            if (!event) {
                // Cria um novo evento.
                const eventOptions = { description: eventDescription };
                // Adiciona convidados se houver.
                if (guests.length > 0) {
                    const uniqueGuests = [...new Set(guests)];  // Garante emails únicos.
                    eventOptions.guests =
                        uniqueGuests.join(',');       // Formato esperado pela API.
                    eventOptions.sendInvites = true;  // Envia convites por email.
                    Logger.log(
                        'Convidados adicionados ao novo evento: ' +
                        uniqueGuests.join(', '));
                }
                event =
                    calendar.createEvent(eventTitle, startTime, endTime, eventOptions);
                Logger.log(`Evento do Calendar criado com ID: ${event.getId()}`);
            } else {
                // Se o evento foi atualizado com sucesso.
                Logger.log(`Evento do Calendar ID ${event.getId()} atualizado.`);
            }

            // --- Salva o ID do Evento na Planilha de Instâncias ---
            // Atualiza a coluna ID_EVENTO_CALENDAR na linha da instância com o ID do
            // evento (novo ou atualizado).
            instancesSheet
                .getRange(
                    instanceRowIndex,
                    HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR + 1)
                .setValue(event.getId());
            // Armazena o ID para retornar na resposta JSON.
            calendarEventId = event.getId();

        } catch (calendarError) {
            // Se ocorrer um erro durante a interação com o Google Calendar.
            Logger.log(
                'Erro crítico no Calendar: ' + calendarError.message +
                ' Stack: ' + calendarError.stack);
            // Libera o bloqueio.
            lock.releaseLock();
            // Retorna sucesso na reserva da planilha, mas informa sobre o erro no
            // Calendar. A reserva está feita no sistema, mas o evento pode estar
            // ausente ou incorreto.
            enviarEmailListaFixa([...new Set(guests)], eventTitle, eventDescription, 'bcc');
            return JSON.stringify({
                success: true,
                message:
                    `Reserva agendada com sucesso, mas houve um erro ao criar/atualizar o evento no Google Calendar: ${calendarError.message}. Verifique os logs.`,
                data: { bookingId: bookingId, eventId: null }
            });
        }

        // --- Finalização ---
        // Libera o bloqueio, pois todas as operações foram concluídas.
        lock.releaseLock();
        enviarEmailListaFixa([...new Set(guests)], eventTitle, eventDescription, 'bcc');

        // Retorna JSON indicando sucesso total.
        return JSON.stringify({
            success: true,
            message: `${bookingType} agendada com sucesso!`,
            data: { bookingId: bookingId, eventId: calendarEventId }
        });

    } catch (e) {
        // Captura qualquer erro não tratado no bloco try principal.
        // Verifica se o bloqueio ainda está ativo (pode não estar se o erro ocorreu
        // antes da liberação).
        if (lock.hasLock()) {
            lock.releaseLock();
        }
        Logger.log('Erro no bookSlot: ' + e.message + ' Stack: ' + e.stack);
        // Retorna JSON indicando falha geral.
        return JSON.stringify({
            success: false,
            message: 'Ocorreu um erro ao agendar: ' + e.message,
            data: null
        });
    }
}

/**
 * Função para gerar instâncias futuras de horários com base nos 'Horarios
 * Base'. Normalmente executada periodicamente por um gatilho (trigger) de
 * tempo. Cria registros na planilha 'Instancias de Horarios' para um período
 * futuro definido. Evita criar duplicatas se uma instância para o mesmo horário
 * base, data e hora já existir.
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
        return;  // Aborta a execução.
    }
    if (!instancesSheet) {
        Logger.log(`Erro: Planilha "${SHEETS.SCHEDULE_INSTANCES}" não encontrada.`);
        return;  // Aborta a execução.
    }

    // --- Leitura e Validação dos Horários Base ---
    const baseDataRaw = baseSheet.getDataRange().getValues();
    // Verifica se há dados na planilha base.
    if (baseDataRaw.length <= 1) {
        Logger.log(
            `Planilha "${SHEETS.BASE_SCHEDULES}" está vazia ou apenas cabeçalho.`);
        return;  // Aborta se não houver horários base.
    }

    const baseSchedules = [];  // Array para armazenar os horários base válidos.
    // Define o número mínimo de colunas necessárias nos horários base.
    const expectedBaseCols =
        Math.max(
            HEADERS.BASE_SCHEDULES.ID, HEADERS.BASE_SCHEDULES.DIA_SEMANA,
            HEADERS.BASE_SCHEDULES.HORA_INICIO, HEADERS.BASE_SCHEDULES.TIPO,
            HEADERS.BASE_SCHEDULES.TURMA_PADRAO,
            HEADERS.BASE_SCHEDULES.PROFESSOR_PRINCIPAL) +
        1;

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
        const baseProfessorPrincipalRaw =
            row[HEADERS.BASE_SCHEDULES.PROFESSOR_PRINCIPAL];

        // Formata e valida os dados essenciais.
        const baseId =
            (typeof baseIdRaw === 'string' || typeof baseIdRaw === 'number') ?
                String(baseIdRaw).trim() :
                null;
        const baseDayOfWeek = (typeof baseDayOfWeekRaw === 'string' ||
            typeof baseDayOfWeekRaw === 'number') ?
            String(baseDayOfWeekRaw).trim() :
            null;
        const baseHourString =
            formatValueToHHMM(baseHourRaw, timeZone);  // Formata a hora.
        const baseType =
            (typeof baseTypeRaw === 'string' || typeof baseTypeRaw === 'number') ?
                String(baseTypeRaw).trim() :
                null;
        const baseTurma =
            (typeof baseTurmaRaw === 'string' || typeof baseTurmaRaw === 'number') ?
                String(baseTurmaRaw).trim() :
                null;
        const baseProfessorPrincipal =
            (typeof baseProfessorPrincipalRaw === 'string' ||
                typeof baseProfessorPrincipalRaw === 'number') ?
                String(baseProfessorPrincipalRaw || '').trim() :
                '';  // Permite professor vazio.

        // Pula a linha se dados essenciais forem inválidos após formatação.
        if (!baseId || baseId === '' || !baseDayOfWeek || baseDayOfWeek === '' ||
            baseHourString === null || !baseType || baseType === '' || !baseTurma ||
            baseTurma === '') {
            Logger.log(`Skipping base schedule row ${rowIndex} due to invalid/missing essential data: ID=${baseIdRaw}, Dia=${baseDayOfWeekRaw}, Hora=${baseHourRaw}, Tipo=${baseTypeRaw}, Turma=${baseTurmaRaw}`);
            continue;
        }

        // Validações adicionais de consistência.
        const daysOfWeek =
            ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
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
        Logger.log('Nenhum horário base válido encontrado para gerar instâncias.');
        return;
    }
    Logger.log(`Processados ${baseSchedules.length} horários base válidos.`);

    // --- Leitura das Instâncias Existentes para Evitar Duplicatas ---
    const existingInstancesRaw = instancesSheet.getDataRange().getValues();
    const existingInstancesMap =
        {};  // Mapa para armazenar chaves de instâncias existentes.
    // Colunas necessárias para criar a chave única de identificação de uma
    // instância.
    const mapKeyCols = Math.max(
        HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO,
        HEADERS.SCHEDULE_INSTANCES.DATA,
        HEADERS.SCHEDULE_INSTANCES.HORA_INICIO) +
        1;
    const instanceIdCol =
        HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA;  // Coluna do ID da instância.

    // Processa as instâncias existentes apenas se houver dados além do cabeçalho.
    if (existingInstancesRaw.length > 1) {
        // Verifica se a planilha de instâncias tem colunas suficientes para a
        // chave.
        if (existingInstancesRaw[0].length < mapKeyCols) {
            Logger.log(`Warning: Planilha "${SHEETS.SCHEDULE_INSTANCES}" tem menos colunas (${existingInstancesRaw[0].length}) que o esperado (${mapKeyCols}) para verificação de duplicidade.`);
        }

        // Itera pelas linhas de instâncias existentes (pulando cabeçalho).
        for (let j = 1; j < existingInstancesRaw.length; j++) {
            const row = existingInstancesRaw[j];
            const rowIndex = j + 1;  // Índice real na planilha.

            // Pula linhas incompletas que não permitem criar a chave.
            if (!row || row.length < mapKeyCols) {
                // Logger.log(`Skipping existing instance row ${rowIndex} for map
                // creation due to insufficient columns.`);
                continue;  // Pula silenciosamente para não poluir muito o log.
            }

            // Extrai os dados brutos para a chave e o ID da instância.
            const existingBaseIdRaw = row[HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO];
            const existingDateRaw = row[HEADERS.SCHEDULE_INSTANCES.DATA];
            const existingHourRaw = row[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO];
            // Pega o ID da instância (se a coluna existir).
            const existingInstanceIdRaw =
                (row.length > instanceIdCol) ? row[instanceIdCol] : null;

            // Formata e valida os dados para a chave.
            const existingBaseId = (typeof existingBaseIdRaw === 'string' ||
                typeof existingBaseIdRaw === 'number') ?
                String(existingBaseIdRaw).trim() :
                null;
            const existingDate = formatValueToDate(existingDateRaw);
            const existingHourString = formatValueToHHMM(existingHourRaw, timeZone);
            // Formata o ID da instância.
            const existingInstanceId = (typeof existingInstanceIdRaw === 'string' ||
                typeof existingInstanceIdRaw === 'number') ?
                String(existingInstanceIdRaw).trim() :
                null;

            // Se todos os componentes da chave e o ID da instância forem válidos.
            if (existingBaseId && existingDate && existingHourString &&
                existingInstanceId) {
                // Formata a data como string 'yyyy-MM-dd' para a chave do mapa.
                const existingDateStr =
                    Utilities.formatDate(existingDate, timeZone, 'yyyy-MM-dd');
                // Cria a chave única: IDBase_Data_Hora.
                const mapKey =
                    `${existingBaseId}_${existingDateStr}_${existingHourString}`;
                // Adiciona a chave ao mapa, associada ao ID da instância (valor pode
                // ser útil para debug).
                existingInstancesMap[mapKey] = existingInstanceId;
            }
        }
    }
    Logger.log(`Map de instâncias existentes populado com ${Object.keys(existingInstancesMap).length} chaves.`);

    // --- Geração das Novas Instâncias ---
    const numWeeksToGenerate = 4;  // Define quantas semanas no futuro gerar.
    const today = new Date();      // Data atual.
    today.setHours(0, 0, 0, 0);    // Zera a hora.

    const newInstances = [
    ];  // Array para armazenar as novas linhas de instância a serem inseridas.
    const daysOfWeek = [
        'Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'
    ];  // Mapeamento dia -> nome.
    // Número total de colunas na planilha de instâncias.
    const numColsInstance = HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR + 1;

    // Define a data de início da geração: a próxima segunda-feira a partir de
    // hoje.
    let startGenerationDate = new Date(today.getTime());
    const currentDayOfWeek =
        startGenerationDate.getDay();  // 0=Domingo, 1=Segunda, ...
    // Calcula quantos dias faltam para a próxima segunda-feira.
    const daysUntilMonday =
        (currentDayOfWeek === 0) ? 1 : (8 - currentDayOfWeek) % 7;
    // Se hoje não for segunda, avança para a próxima segunda.
    if (daysUntilMonday !== 0) {
        startGenerationDate.setDate(
            startGenerationDate.getDate() + daysUntilMonday);
    }
    startGenerationDate.setHours(0, 0, 0, 0);  // Garante que a hora está zerada.

    // Define a data final da geração (inclusive).
    const endGenerationDate = new Date(startGenerationDate.getTime());
    // Avança (número de semanas * 7 - 1) dias para cobrir o período desejado.
    endGenerationDate.setDate(
        endGenerationDate.getDate() + (numWeeksToGenerate * 7) - 1);
    Logger.log(`Gerando instâncias de ${Utilities.formatDate(startGenerationDate, timeZone, 'yyyy-MM-dd')} até ${Utilities.formatDate(endGenerationDate, timeZone, 'yyyy-MM-dd')}`);

    // Itera por cada dia dentro do período de geração.
    let currentDate = new Date(startGenerationDate.getTime());
    while (currentDate <= endGenerationDate) {
        const targetDate =
            new Date(currentDate.getTime());  // Cria cópia da data atual do loop.
        // Obtém o nome do dia da semana para a data atual.
        const targetDayOfWeekName = daysOfWeek[targetDate.getDay()];

        // Filtra os horários base que ocorrem neste dia da semana.
        const schedulesForThisDay = baseSchedules.filter(
            schedule => schedule.dayOfWeek === targetDayOfWeekName);

        // Para cada horário base que deve ocorrer neste dia:
        for (const baseSchedule of schedulesForThisDay) {
            const baseId = baseSchedule.id;
            const baseHourString = baseSchedule.hour;
            const baseTurma = baseSchedule.turma;
            const baseProfessorPrincipal = baseSchedule.professorPrincipal;

            // Cria a chave única para esta possível instância (IDBase_Data_Hora).
            const instanceDateStr =
                Utilities.formatDate(targetDate, timeZone, 'yyyy-MM-dd');
            const predictableInstanceKey =
                `${baseId}_${instanceDateStr}_${baseHourString}`;

            // Verifica se uma instância com essa chave JÁ EXISTE no mapa.
            if (!existingInstancesMap[predictableInstanceKey]) {
                // Se NÃO EXISTE, cria uma nova linha (array) para a instância.
                const newRow = [];
                // Inicializa a linha com strings vazias para todas as colunas.
                for (let colIdx = 0; colIdx < numColsInstance; colIdx++) {
                    newRow[colIdx] = '';
                }

                // Preenche os dados da nova instância.
                newRow[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA] =
                    Utilities.getUuid();  // Gera um novo ID único.
                newRow[HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO] = baseId;
                newRow[HEADERS.SCHEDULE_INSTANCES.TURMA] = baseTurma;
                newRow[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL] =
                    baseProfessorPrincipal;
                newRow[HEADERS.SCHEDULE_INSTANCES.DATA] = targetDate;  // Objeto Date.
                newRow[HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA] = targetDayOfWeekName;
                newRow[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO] =
                    baseHourString;  // String HH:mm.
                newRow[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL] = baseSchedule.type;
                newRow[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO] =
                    STATUS_OCUPACAO.DISPONIVEL;  // Status inicial.
                // ID_RESERVA e ID_EVENTO_CALENDAR ficam vazios inicialmente.

                // Adiciona a nova linha ao array de instâncias a serem inseridas.
                newInstances.push(newRow);
                // Adiciona a chave desta nova instância ao mapa para evitar duplicatas
                // dentro do mesmo ciclo de geração.
                existingInstancesMap[predictableInstanceKey] =
                    newRow[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA];
            }
        }
        // Avança para o próximo dia.
        currentDate.setDate(currentDate.getDate() + 1);
    }

    Logger.log(`Pronto para inserir ${newInstances.length} novas instâncias.`);

    // --- Inserção das Novas Instâncias na Planilha ---
    // Verifica se há novas instâncias para inserir.
    if (newInstances.length > 0) {
        // Verificação de segurança: garante que as linhas a serem inseridas têm o
        // número correto de colunas.
        if (newInstances[0].length !== numColsInstance) {
            Logger.log(`Erro interno: O array newInstances tem ${newInstances[0].length} colunas, mas esperava ${numColsInstance}. Abortando inserção.`);
            // Lança um erro para interromper, pois isso indica um problema na lógica
            // de criação da linha.
            throw new Error('Erro na estrutura interna dos dados a serem salvos.');
        }

        try {
            // Insere TODAS as novas linhas de uma vez na planilha para melhor
            // performance. Começa a inserir na primeira linha vazia (getLastRow() +
            // 1).
            instancesSheet
                .getRange(
                    instancesSheet.getLastRow() + 1, 1, newInstances.length,
                    numColsInstance)
                .setValues(newInstances);
            Logger.log(`Geradas ${newInstances.length} novas instâncias de horários salvas.`);
        } catch (e) {
            // Se ocorrer um erro durante a inserção em lote.
            Logger.log(`Erro ao salvar novas instâncias na planilha "${SHEETS.SCHEDULE_INSTANCES}": ${e.message} Stack: ${e.stack}`);
            // Lança um erro para que o gatilho (se houver) registre a falha.
            throw new Error(`Erro ao salvar novas instâncias: ${e.message}`);
        }
    } else {
        // Se nenhuma nova instância foi gerada (talvez já existissem todas para o
        // período).
        Logger.log('Nenhuma nova instância de horário gerada para o período.');
    }
    Logger.log('*** createScheduleInstances finalizada ***');
}

/**
 * Apaga instâncias de horários (SCHEDULE_INSTANCES) que ocorreram
 * antes da data especificada na configuração 'Data Limite Limpeza Instancias'.
 * Projetada para ser executada por um gatilho (trigger) de tempo.
 */
function cleanOldScheduleInstances() {
    // Adquire um bloqueio para garantir que apenas uma instância deste script
    // esteja acessando e modificando os dados de limpeza por vez.
    // Espera até 30 segundos se o bloqueio já estiver ativo.
    const lock = LockService.getScriptLock();
    const lockAcquired =
        lock.tryLock(30000);  // Espera até 30 segundos (30000 ms)

    if (!lockAcquired) {
        Logger.log(
            'Não foi possível adquirir o bloqueio para cleanOldScheduleInstances. Outro processo pode estar em execução.');
        return;  // Aborta a execução se não puder obter o bloqueio
    }

    Logger.log('*** cleanOldScheduleInstances chamada ***');

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const instancesSheet = ss.getSheetByName(SHEETS.SCHEDULE_INSTANCES);
        const timeZone = ss.getSpreadsheetTimeZone();  // Para logs formatados

        // 1. Verifica se a planilha de instâncias existe
        if (!instancesSheet) {
            Logger.log(
                `Erro: Planilha "${SHEETS.SCHEDULE_INSTANCES}" não encontrada.`);
            return;  // Aborta a execução
        }

        // 2. Obtém e valida a data limite para limpeza da configuração
        const cleanupDateString = getConfigValue('Data Limite');
        if (!cleanupDateString) {
            Logger.log(
                `Erro: Configuração "Data Limite" não encontrada ou vazia na planilha "${SHEETS.CONFIG}".`);
            Logger.log(
                '*** cleanOldScheduleInstances abortada: Configuração de data limite ausente ***');
            return;  // Aborta se a configuração não existe ou está vazia
        }

        const cleanupDate = parseDDMMYYYY(cleanupDateString);

        if (!cleanupDate || isNaN(cleanupDate.getTime())) {
            Logger.log(`Erro: Valor da configuração "Data Limite" inválido: "${cleanupDateString}". Esperado formato dd/MM/yyyy.`);
            Logger.log(
                '*** cleanOldScheduleInstances abortada: Data limite de limpeza inválida na configuração ***');
            return;  // Aborta se a data na configuração for inválida
        }

        // Zera a hora da data limite para garantir que comparações sejam apenas por
        // data
        cleanupDate.setHours(0, 0, 0, 0);
        Logger.log(
            `Data limite para limpeza (instâncias ANTES desta data serão apagadas): ${Utilities.formatDate(cleanupDate, timeZone, 'dd/MM/yyyy')}`);

        // 3. Obtém todos os dados da planilha de instâncias
        const rawData = instancesSheet.getDataRange().getValues();

        // Verifica se há dados além do cabeçalho
        if (rawData.length <= 1) {
            Logger.log(`Planilha "${SHEETS
                .SCHEDULE_INSTANCES}" vazia ou apenas cabeçalho. Nenhuma limpeza necessária.`);
            return;  // Aborta se não houver dados
        }

        const header = rawData[0];      // Guarda o cabeçalho
        const data = rawData.slice(1);  // Dados sem o cabeçalho
        const dataColIndex =
            HEADERS.SCHEDULE_INSTANCES.DATA;  // Índice da coluna de data
        const numCols = header.length;  // Número total de colunas na planilha
        // (baseado no cabeçalho real)

        // Verifica se a coluna de data existe
        if (dataColIndex >= numCols) {
            Logger.log(`Erro: Coluna de Data (índice ${dataColIndex}) não encontrada na planilha de instâncias (tem apenas ${numCols} colunas). Verifique a estrutura da planilha "${SHEETS.SCHEDULE_INSTANCES}".`);
            Logger.log(
                '*** cleanOldScheduleInstances abortada: Coluna de Data não encontrada ***');
            return;  // Aborta se a coluna de data não existe
        }

        // 4. Filtra as linhas que devem ser mantidas
        const rowsToKeep = [];
        let initialRowCount =
            data.length;  // Número de linhas de dados originais (sem cabeçalho)
        let deletedCount = 0;

        for (let i = 0; i < data.length; i++) {
            const row = data[i];
            const rowIndexInSheet = i + 2;  // Índice real da linha na planilha
            // (contando o cabeçalho e base 1)

            // Verifica se a linha tem colunas suficientes para conter a data
            if (!row || row.length <= dataColIndex) {
                Logger.log(`Warning: Ignorando linha ${rowIndexInSheet} na planilha de instâncias devido a colunas insuficientes (${row ? row.length : 0} vs requerido ${dataColIndex + 1}). Será tratada como apagável.`);
                deletedCount++;  // Contabiliza linhas incompletas como apagadas
                continue;        // Pula para a próxima linha
            }

            const rawInstanceDate = row[dataColIndex];
            const instanceDate = formatValueToDate(
                rawInstanceDate);  // Usa a função existente para formatar/validar

            let keepRow = false;

            // Verifica se a data da instância é válida
            if (instanceDate && !isNaN(instanceDate.getTime())) {
                // Zera a hora da data da instância para comparação consistente (já
                // feito em formatValueToDate, mas reforça)
                instanceDate.setHours(0, 0, 0, 0);

                // Mantém a linha se a data da instância for IGUAL ou POSTERIOR à data
                // limite de limpeza
                if (instanceDate >= cleanupDate) {
                    rowsToKeep.push(row);
                    keepRow = true;  // Marca para manter
                } else {
                    // Linha será apagada por ser anterior à data limite
                    // Logger.log(`Linha ${rowIndexInSheet} (Data:
                    // ${Utilities.formatDate(instanceDate, timeZone, 'dd/MM/yyyy')}) é
                    // anterior à data limite ${Utilities.formatDate(cleanupDate,
                    // timeZone, 'dd/MM/yyyy')}. Marcada para apagar.`);
                }
            } else {
                // Linha será apagada por ter uma data inválida/ausente
                Logger.log(`Warning: Linha ${rowIndexInSheet} tem valor de data inválido ou ausente ("${rawInstanceDate}"). Marcada para apagar.`);
            }

            if (!keepRow) {
                // Se a linha não foi marcada para manter, ela será apagada (inclui
                // inválidas e antigas)
                deletedCount++;
            }
        }

        // Recalcula o número real de linhas deletadas com base no filtro
        deletedCount = initialRowCount - rowsToKeep.length;
        Logger.log(
            `Filtragem completa. ${rowsToKeep.length} instâncias serão mantidas, ${deletedCount} serão apagadas.`);


        // 5. Reescreve os dados (cabeçalho + linhas a manter) na planilha
        // Prepara os dados a serem escritos, incluindo o cabeçalho no início
        const dataToWrite = [header, ...rowsToKeep];

        // Verifica se há dados para escrever (deve haver pelo menos o cabeçalho se
        // a planilha não estava vazia e o cabeçalho foi lido)
        if (dataToWrite.length === 0) {
            Logger.log(
                'Erro interno: O array dataToWrite está vazio inesperadamente.');
        } else {
            // Pad/Trim rowsToKeep to match the number of columns in the header.
            // This prevents errors in setValues if some old rows had different column
            // counts.
            const paddedDataToWrite = dataToWrite.map(row => {
                const paddedRow = [...row];  // Create a copy
                while (paddedRow.length < numCols)
                    paddedRow.push('');  // Pad if necessary
                if (paddedRow.length > numCols)
                    return paddedRow.slice(0, numCols);  // Trim if necessary
                return paddedRow;
            });

            // Calcula o número total de linhas a serem escritas (cabeçalho + linhas
            // mantidas)
            const numRowsToWrite = paddedDataToWrite.length;

            // Define o range que cobre a área original que precisa ser limpa
            // (da linha 1 até a última linha onde havia dados, cobrindo todas as
            // colunas originais)
            const originalDataRange =
                instancesSheet.getRange(1, 1, rawData.length, numCols);

            try {
                // Limpa o conteúdo da área original (cabeçalho + dados)
                originalDataRange.clearContent();
                Logger.log(`Conteúdo original (${rawData.length} linhas) da planilha "${SHEETS.SCHEDULE_INSTANCES}" limpo.`);

                // Define o range onde os dados filtrados (e cabeçalho) serão escritos
                // Começa na linha 1, na coluna 1, e cobre o número de linhas e colunas
                // dos dados filtrados
                const targetRange =
                    instancesSheet.getRange(1, 1, numRowsToWrite, numCols);

                // Escreve o cabeçalho e as linhas a manter
                targetRange.setValues(paddedDataToWrite);
                Logger.log(`Dados filtrados reescritos na planilha "${SHEETS.SCHEDULE_INSTANCES}". Total de linhas agora: ${numRowsToWrite}.`);

                // Não é mais necessário limpar abaixo, pois clearContent() já limpou a
                // área maior

            } catch (e) {
                // Se houver um erro durante a reescrita dos dados
                Logger.log(`Erro ao limpar/reescrever dados na planilha "${SHEETS.SCHEDULE_INSTANCES}": ${e.message} Stack: ${e.stack}`);
                // Lança o erro para que o gatilho (se usado) possa registrá-lo
                throw new Error(`Erro ao reescrever dados da instância: ${e.message}`);
            }
        }

        Logger.log(`Limpeza de instâncias antigas concluída. ${deletedCount} instâncias foram apagadas.`);
        Logger.log('*** cleanOldScheduleInstances finalizada com sucesso ***');

    } catch (e) {
        // Captura erros inesperados durante a execução
        Logger.log(
            'Erro inesperado em cleanOldScheduleInstances: ' + e.message +
            ' Stack: ' + e.stack);
        // Dependendo de como você quer que os gatilhos se comportem, pode ser útil
        // relançar o erro: throw e;
    } finally {
        // Garante que o bloqueio seja liberado, mesmo se ocorrer um erro
        if (lock.hasLock()) {
            lock.releaseLock();
            Logger.log('Bloqueio de script liberado.');
        }
    }
}