/**
 * Arquivo: Bookings.gs
 * Descrição: Funções para gerenciar a criação e cancelamento de reservas.
 */
function checkLunchAlertPreBooking(slotHoraInicio, slotTurma, slotData) {
    Logger.log(`*** checkLunchAlertPreBooking called for Time: ${slotHoraInicio}, Turma: ${slotTurma}, Date: ${slotData} ***`);
    try {
        if (!slotHoraInicio || !slotTurma || !slotData) {
            throw new Error("Parâmetros insuficientes para verificação do alerta de almoço.");
        }
        if (!/^\d{2}:\d{2}$/.test(slotHoraInicio)) {
            throw new Error("Formato de hora inválido para verificação do alerta.");
        }
        const bookingDateObjUTC = parseDDMMYYYY(slotData);
        if (!bookingDateObjUTC) {
            throw new Error("Formato de data inválido para verificação do alerta.");
        }
        const needsAlert = checkLunchAlertNeeded_(slotHoraInicio, slotTurma, slotData, bookingDateObjUTC);
        Logger.log(`checkLunchAlertPreBooking result: ${needsAlert}`);
        return createJsonResponse(true, "Verificação do alerta de almoço concluída.", { needsAlert: needsAlert });
    } catch (e) {
        Logger.log(`ERROR in checkLunchAlertPreBooking: ${e.message}\nStack: ${e.stack}`);
        return createJsonResponse(false, `Erro ao verificar alerta: ${e.message}`, { needsAlert: false });
    }
}
function processBooking_(bookingDetails, userEmail, userRole) {
    Logger.log(`processBooking_ started for user ${userEmail} (Role: ${userRole}).`);
    const { idInstancia, tipoReserva, professorReal, disciplinaReal, tipoAulaReposicao } = bookingDetails;
    const instanceIdToBook = idInstancia;
    const bookingType = tipoReserva;
    const { header: instanceHeader, sheet: instancesSheet } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
    const { sheet: bookingsSheet } = getSheetData_(SHEETS.BOOKING_DETAILS);
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const instanceIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA + 1;
    const instanceIdFinder = instancesSheet.createTextFinder(instanceIdToBook).matchEntireCell(true);
    const foundCells = instanceIdFinder.findAll();
    if (foundCells.length === 0) {
        throw new Error(`Horário com ID ${instanceIdToBook} não encontrado.`);
    }
    if (foundCells.length > 1) {
        Logger.log(`WARNING: Multiple rows found for instance ID ${instanceIdToBook}. Using the first one found.`);
    }
    const instanceRowIndex = foundCells[0].getRow();
    Logger.log(`Instance ${instanceIdToBook} found at row ${instanceRowIndex}. Reading row data.`);
    const instanceDetails = instancesSheet.getRange(instanceRowIndex, 1, 1, instanceHeader.length).getValues()[0];
    const maxIndexNeeded = Math.max(...Object.values(HEADERS.SCHEDULE_INSTANCES));
    if (instanceDetails.length <= maxIndexNeeded) {
        throw new Error(`Dados da linha ${instanceRowIndex} na planilha "${SHEETS.SCHEDULE_INSTANCES}" estão incompletos.`);
    }
    const currentStatus = String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO] || '').trim();
    const originalType = String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL] || '').trim();
    const rawBookingDate = instanceDetails[HEADERS.SCHEDULE_INSTANCES.DATA];
    const rawBookingTime = instanceDetails[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO];
    const professorPrincipalInstancia = String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL] || '').trim();
    const turmaInstancia = String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.TURMA] || '').trim();
    const bookingDateObj = formatValueToDate(rawBookingDate);
    const bookingHourString = formatValueToHHMM(rawBookingTime, timeZone);
    if (!currentStatus || !originalType || !turmaInstancia || !bookingDateObj || !bookingHourString) {
        throw new Error(`Erro interno: Dados essenciais do horário ${instanceIdToBook} (linha ${instanceRowIndex}) são inválidos na planilha.`);
    }
    let professorOriginal = '';
    if (bookingType === TIPOS_RESERVA.REPOSICAO) {
        if (![USER_ROLES.ADMIN, USER_ROLES.PROFESSOR].includes(userRole)) throw new Error(`Seu perfil (${userRole}) não permite agendar Reposições/Recuperações.`);
        if (originalType !== TIPOS_HORARIO.VAGO) throw new Error(`Reposição/Recuperação só pode ser feita em horários Vagos (este é ${originalType}).`);
        if (currentStatus !== STATUS_OCUPACAO.DISPONIVEL) throw new Error(`Este horário vago (${instanceIdToBook}) não está mais disponível (Status atual: ${currentStatus}). Atualize a lista.`);
        if (!tipoAulaReposicao || (tipoAulaReposicao !== TIPOS_AULA_REPOSICAO.REPOSICAO && tipoAulaReposicao !== TIPOS_AULA_REPOSICAO.RECUPERACAO_PARALELA)) {
            throw new Error(`Tipo de aula (Reposição/Recuperação) inválido ou ausente: "${tipoAulaReposicao}".`);
        }
        Logger.log(`Validation OK for ${tipoAulaReposicao} (BookingType: ${bookingType}) by ${userRole}`);
    } else if (bookingType === TIPOS_RESERVA.SUBSTITUICAO) {
        if (![USER_ROLES.ADMIN, USER_ROLES.PROFESSOR].includes(userRole)) throw new Error(`Seu perfil (${userRole}) não permite agendar Substituições.`);
        if (originalType !== TIPOS_HORARIO.FIXO) throw new Error(`Substituição só pode ser feita em horários Fixos (este é ${originalType}).`);
        if (!professorPrincipalInstancia) throw new Error(`Erro interno: O horário fixo ${instanceIdToBook} não tem um Professor Principal definido.`);
        professorOriginal = professorPrincipalInstancia;
        if (currentStatus === STATUS_OCUPACAO.REPOSICAO_AGENDADA) throw new Error(`Este horário fixo (${instanceIdToBook}) já está ocupado por uma Reposição/Recuperação.`);
        if (currentStatus !== STATUS_OCUPACAO.DISPONIVEL && currentStatus !== STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA) throw new Error(`Este horário fixo (${instanceIdToBook}) não está disponível para substituição (Status atual: ${currentStatus}). Atualize a lista.`);
        Logger.log(`Validation OK for Substituicao by ${userRole}`);
    }
    const bookingId = Utilities.getUuid();
    const creationTimestamp = new Date();
    // O status na instância será sempre 'Reposicao Agendada' para simplificar,
    // o detalhe (Reposição ou Recuperação) fica na planilha de Reservas Detalhadas.
    const newStatus = (bookingType === TIPOS_RESERVA.REPOSICAO) ? STATUS_OCUPACAO.REPOSICAO_AGENDADA : STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA;
    const [hour, minute] = bookingHourString.split(':').map(Number);
    const effectiveStartDateTime = new Date(bookingDateObj.getUTCFullYear(), bookingDateObj.getUTCMonth(), bookingDateObj.getUTCDate(), hour, minute);
    const updatedInstanceRow = [...instanceDetails];
    updatedInstanceRow[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO] = newStatus;
    updatedInstanceRow[HEADERS.SCHEDULE_INSTANCES.ID_RESERVA] = bookingId;
    updatedInstanceRow[HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR] = '';
    updateSheetRow_(instancesSheet, instanceRowIndex, instanceHeader.length, updatedInstanceRow);
    const bookingHeader = bookingsSheet.getRange(1, 1, 1, bookingsSheet.getLastColumn()).getValues()[0];
    const numBookingCols = Math.max(bookingHeader.length, HEADERS.BOOKING_DETAILS.TIPO_AULA_REPOSICAO + 1); // Garante que a coluna exista

    const newBookingRow = new Array(numBookingCols).fill('');
    newBookingRow[HEADERS.BOOKING_DETAILS.ID_RESERVA] = bookingId;
    newBookingRow[HEADERS.BOOKING_DETAILS.TIPO_RESERVA] = bookingType; // 'Reposicao' ou 'Substituicao'
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
    // Salva o tipo de aula específico para reposições/recuperações
    if (bookingType === TIPOS_RESERVA.REPOSICAO) {
        newBookingRow[HEADERS.BOOKING_DETAILS.TIPO_AULA_REPOSICAO] = tipoAulaReposicao;
    } else {
        newBookingRow[HEADERS.BOOKING_DETAILS.TIPO_AULA_REPOSICAO] = ''; // Limpa para substituições
    }

    appendSheetRow_(bookingsSheet, numBookingCols, newBookingRow);
    const guestEmails = getGuestEmailsForBooking_(professorReal, professorPrincipalInstancia);
    Logger.log("processBooking_ completed successfully.");
    return {
        bookingId: bookingId,
        instanceRowIndex: instanceRowIndex,
        instanceDetails: updatedInstanceRow, // updatedInstanceRow contém os dados da instância após a atualização
        professorOriginal: professorOriginal,
        effectiveStartDateTime: effectiveStartDateTime,
        creationTimestamp: creationTimestamp,
        guestEmails: guestEmails,
        turmaInstancia: turmaInstancia, // Retorna a turma para a verificação de almoço
        bookingHourString: bookingHourString, // Retorna a hora para a verificação de almoço
        bookingDateObj: bookingDateObj // Retorna a data para a verificação de almoço
    };
}
function getGuestEmailsForBooking_(profReal, professorsPrincipalString) {
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
    if (professorsPrincipalString) {
        const principalProfNames = professorsPrincipalString.split(',')
            .map(name => name.trim())
            .filter(name => name !== '');
        Logger.log(`Processing principal professors from string "${professorsPrincipalString}": [${principalProfNames.join(', ')}]`);
        principalProfNames.forEach(name => {
            if (nameEmailMap[name]) {
                guests.add(nameEmailMap[name]);
                Logger.log(`Adding guest (Principal): ${name} -> ${nameEmailMap[name]}`);
            } else {
                Logger.log(`Warning: Email for Principal Professor "${name}" not found.`);
            }
        });
    } else {
        Logger.log(`Principal professors string is empty or null.`);
    }
    const guestArray = Array.from(guests);
    Logger.log(`Final guest list for booking: [${guestArray.join(', ')}]`);
    return guestArray;
}
function getCancellableBookings() {
    Logger.log('*** getCancellableBookings called ***');
    let userEmail = '[Unknown]';
    try {
        userEmail = getActiveUserEmail_();
        const userRole = getUserRolePlain_(userEmail);
        const userProfessorName = (userRole === USER_ROLES.PROFESSOR) ? getProfessorNameByEmail_(userEmail) : null;
        const cancelAdminProfessorEmailRaw = getConfigValue('Professor Admin Cancelamento');
        const cancelAdminProfessorEmail = cancelAdminProfessorEmailRaw ? cancelAdminProfessorEmailRaw.trim().toLowerCase() : null;
        const canViewAll = (userRole === USER_ROLES.ADMIN) || (cancelAdminProfessorEmail && userEmail.toLowerCase() === cancelAdminProfessorEmail);
        if (!canViewAll && userRole !== USER_ROLES.PROFESSOR) {
            Logger.log(`Access denied for ${userEmail} (Role: ${userRole}) to view cancellable bookings.`);
            return createJsonResponse(false, 'Acesso negado. Você não tem permissão para visualizar esta lista.', null);
        }
        const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
        const { data: bookingData } = getSheetData_(SHEETS.BOOKING_DETAILS, HEADERS.BOOKING_DETAILS);
        const { data: instanceData } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
        const today = new Date();
        today.setHours(0, 0, 0, 0);
        const instanceMap = instanceData.reduce((map, row) => {
            const idCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA;
            if (row && row.length > idCol) {
                const id = String(row[idCol] || '').trim();
                if (id) {
                    map[id] = {
                        date: formatValueToDate(row[HEADERS.SCHEDULE_INSTANCES.DATA]),
                        time: formatValueToHHMM(row[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO], timeZone),
                        turma: String(row[HEADERS.SCHEDULE_INSTANCES.TURMA] || '').trim()
                    };
                }
            }
            return map;
        }, {});
        const cancellableBookings = [];
        const bookingIdCol = HEADERS.BOOKING_DETAILS.ID_RESERVA;
        const instanceFkCol = HEADERS.BOOKING_DETAILS.ID_INSTANCIA;
        const statusCol = HEADERS.BOOKING_DETAILS.STATUS_RESERVA;
        const typeCol = HEADERS.BOOKING_DETAILS.TIPO_RESERVA; // 'Reposicao' ou 'Substituicao'
        const tipoAulaCol = HEADERS.BOOKING_DETAILS.TIPO_AULA_REPOSICAO; // 'Reposição' ou 'Recuperação Paralela'
        const discCol = HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL;
        const profRealCol = HEADERS.BOOKING_DETAILS.PROFESSOR_REAL;
        const profOrigCol = HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL;
        const creatorCol = HEADERS.BOOKING_DETAILS.CRIADO_POR;
        const reqCols = Math.max(bookingIdCol, instanceFkCol, statusCol, typeCol, tipoAulaCol, discCol, profRealCol, profOrigCol, creatorCol) + 1;

        bookingData.forEach(row => {
            if (!row || row.length < reqCols) return;
            const bookingStatus = String(row[statusCol] || '').trim();
            const instanceId = String(row[instanceFkCol] || '').trim();
            const instanceInfo = instanceMap[instanceId];
            if (bookingStatus !== 'Agendada' || !instanceInfo || !instanceInfo.date || instanceInfo.date < today) {
                return;
            }
            const bookingId = String(row[bookingIdCol] || '').trim();
            let displayBookingType = String(row[typeCol] || '').trim(); // e.g., 'Reposicao'
            const specificAulaType = String(row[tipoAulaCol] || '').trim(); // e.g., 'Recuperação Paralela'

            // Se for uma "Reposicao" genérica e tiver um tipo específico de aula, use o tipo específico.
            if (displayBookingType === TIPOS_RESERVA.REPOSICAO && specificAulaType) {
                displayBookingType = specificAulaType; // Muda para "Reposição" ou "Recuperação Paralela"
            }

            const disciplina = String(row[discCol] || '').trim();
            const profReal = String(row[profRealCol] || '').trim();
            const profOrig = String(row[profOrigCol] || '').trim();
            const criadoPor = String(row[creatorCol] || '').trim();
            let canViewThisBooking = false;
            if (canViewAll) {
                canViewThisBooking = true;
            } else if (userRole === USER_ROLES.PROFESSOR && userProfessorName) {
                if (criadoPor.toLowerCase() === userEmail.toLowerCase() ||
                    (profReal && profReal === userProfessorName) ||
                    (profOrig && profOrig === userProfessorName)) {
                    canViewThisBooking = true;
                }
            }
            if (canViewThisBooking) {
                cancellableBookings.push({
                    bookingId: bookingId,
                    instanceId: instanceId,
                    bookingType: displayBookingType, // Mostra o tipo de aula mais específico
                    date: Utilities.formatDate(instanceInfo.date, timeZone, 'dd/MM/yyyy'),
                    time: instanceInfo.time || 'N/D',
                    turma: instanceInfo.turma || 'N/D',
                    disciplina: disciplina,
                    profReal: profReal,
                    profOrig: profOrig,
                    criadoPor: criadoPor
                });
            }
        });
        cancellableBookings.sort((a, b) => {
            const dateA = a.date.split('/').reverse().join('');
            const dateB = b.date.split('/').reverse().join('');
            if (dateA !== dateB) return dateA.localeCompare(dateB);
            if (a.time !== b.time) return a.time.localeCompare(b.time);
            return a.turma.localeCompare(b.turma);
        });
        Logger.log(`Found ${cancellableBookings.length} cancellable bookings for user ${userEmail} (Role: ${userRole}, Name: ${userProfessorName || 'N/A'}).`);
        return createJsonResponse(true, `${cancellableBookings.length} reserva(s) encontrada(s) que você pode cancelar.`, cancellableBookings);
    } catch (e) {
        Logger.log(`ERROR in getCancellableBookings for user ${userEmail}: ${e.message}\nStack: ${e.stack}`);
        return createJsonResponse(false, `Erro ao buscar reservas canceláveis: ${e.message}`, null);
    }
}
function cancelBookingAdmin(bookingIdToCancel) {
    Logger.log(`*** cancelBookingAdmin (includes professor logic) called for Booking ID: ${bookingIdToCancel} ***`);
    let lock = null;
    let userEmail = '[Unknown]';
    try {
        userEmail = getActiveUserEmail_();
        const userRole = getUserRolePlain_(userEmail);
        const userProfessorName = (userRole === USER_ROLES.PROFESSOR) ? getProfessorNameByEmail_(userEmail) : null;
        const cancelAdminProfessorEmailRaw = getConfigValue('Professor Admin Cancelamento');
        const cancelAdminProfessorEmail = cancelAdminProfessorEmailRaw ? cancelAdminProfessorEmailRaw.trim().toLowerCase() : null;
        const isActualAdmin = (userRole === USER_ROLES.ADMIN);
        const isCancelAdminProfessor = (cancelAdminProfessorEmail && userEmail.toLowerCase() === cancelAdminProfessorEmail);
        const isPrivilegedUser = isActualAdmin || isCancelAdminProfessor;
        if (!bookingIdToCancel || typeof bookingIdToCancel !== 'string' || bookingIdToCancel.trim() === '') {
            throw new Error('ID da Reserva inválido ou ausente.');
        }
        const trimmedBookingId = bookingIdToCancel.trim();
        lock = acquireScriptLock_();
        const { header: bookingHeader, data: bookingData, sheet: bookingsSheet } = getSheetData_(SHEETS.BOOKING_DETAILS, HEADERS.BOOKING_DETAILS);
        const bookingIdCol = HEADERS.BOOKING_DETAILS.ID_RESERVA;
        const instanceFkCol = HEADERS.BOOKING_DETAILS.ID_INSTANCIA;
        const bookingStatusCol = HEADERS.BOOKING_DETAILS.STATUS_RESERVA;
        const creatorCol = HEADERS.BOOKING_DETAILS.CRIADO_POR;
        const profRealCol = HEADERS.BOOKING_DETAILS.PROFESSOR_REAL;
        const profOrigCol = HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL;
        const maxBookingIndex = Math.max(bookingIdCol, instanceFkCol, bookingStatusCol, creatorCol, profRealCol, profOrigCol);
        let bookingRowIndex = -1;
        let bookingDetails = null;
        let instanceId = null;
        let bookingCreatorEmail = null;
        let bookingProfReal = null;
        let bookingProfOrig = null;
        for (let i = 0; i < bookingData.length; i++) {
            const row = bookingData[i];
            if (row && row.length > maxBookingIndex && String(row[bookingIdCol] || '').trim() === trimmedBookingId) {
                bookingRowIndex = i + 2;
                bookingDetails = row;
                instanceId = String(row[instanceFkCol] || '').trim();
                bookingCreatorEmail = String(row[creatorCol] || '').trim();
                bookingProfReal = String(row[profRealCol] || '').trim();
                bookingProfOrig = String(row[profOrigCol] || '').trim();
                break;
            }
        }
        if (bookingRowIndex === -1 || !bookingDetails || !instanceId) {
            throw new Error(`Reserva com ID ${trimmedBookingId} não encontrada.`);
        }
        Logger.log(`Booking ${trimmedBookingId} found at row ${bookingRowIndex}, linked to Instance ID ${instanceId}. Creator: ${bookingCreatorEmail}, ProfReal: ${bookingProfReal}, ProfOrig: ${bookingProfOrig}`);
        let canCancel = false;
        if (isPrivilegedUser) {
            canCancel = true;
            Logger.log(`User ${userEmail} is privileged (Admin or Cancel Admin Prof), proceeding.`);
        } else if (userRole === USER_ROLES.PROFESSOR && userProfessorName) {
            if (bookingCreatorEmail.toLowerCase() === userEmail.toLowerCase() ||
                (bookingProfReal && bookingProfReal === userProfessorName) ||
                (bookingProfOrig && bookingProfOrig === userProfessorName)) {
                canCancel = true;
                Logger.log(`Professor ${userEmail} (${userProfessorName}) has permission to cancel this booking.`);
            } else {
                Logger.log(`Professor ${userEmail} (${userProfessorName}) DENIED cancellation for booking ${trimmedBookingId}. Not creator or involved professor.`);
                throw new Error('Você só pode cancelar reservas que você criou ou nas quais você é o professor (real ou original).');
            }
        } else {
            Logger.log(`User ${userEmail} (Role: ${userRole}) DENIED cancellation for booking ${trimmedBookingId}. Insufficient permissions.`);
            throw new Error('Você não tem permissão para cancelar esta reserva.');
        }
        const currentBookingStatus = String(bookingDetails[bookingStatusCol] || '').trim();
        if (currentBookingStatus !== 'Agendada') {
            throw new Error(`Esta reserva (ID ${trimmedBookingId}) já não está com status "Agendada" (Status atual: ${currentBookingStatus}). Não pode ser cancelada novamente.`);
        }
        const { header: instanceHeader, sheet: instancesSheet } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
        const instanceIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA;
        const instanceRowFinder = instancesSheet.createTextFinder(instanceId).matchEntireCell(true);
        const foundInstanceCells = instanceRowFinder.findAll();
        if (foundInstanceCells.length === 0) {
            Logger.log(`CRITICAL INCONSISTENCY: Booking ${trimmedBookingId} exists, but linked Instance ${instanceId} not found! Marking booking as cancelled, but cannot revert instance.`);
            bookingsSheet.getRange(bookingRowIndex, bookingStatusCol + 1).setValue('Cancelada (Instância Não Encontrada)');
            throw new Error(`Erro de dados: Instância ${instanceId} ligada a esta reserva não foi encontrada. Reserva marcada como cancelada, mas o horário pode não ter sido liberado.`);
        }
        const instanceRowIndex = foundInstanceCells[0].getRow();
        const instanceDetailsFromSheet = instancesSheet.getRange(instanceRowIndex, 1, 1, instanceHeader.length).getValues()[0];
        Logger.log(`Linked Instance ${instanceId} found at row ${instanceRowIndex}.`);
        const instanceStatusCol = HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO;
        const instanceBookingIdCol = HEADERS.SCHEDULE_INSTANCES.ID_RESERVA;
        const instanceEventIdCol = HEADERS.SCHEDULE_INSTANCES.ID_EVENTO_CALENDAR;
        const existingEventId = String(instanceDetailsFromSheet[instanceEventIdCol] || '').trim();
        const updatedInstanceRow = [...instanceDetailsFromSheet];
        updatedInstanceRow[instanceStatusCol] = STATUS_OCUPACAO.DISPONIVEL;
        updatedInstanceRow[instanceBookingIdCol] = '';
        updatedInstanceRow[instanceEventIdCol] = '';
        updateSheetRow_(instancesSheet, instanceRowIndex, instanceHeader.length, updatedInstanceRow);
        Logger.log(`Instance ${instanceId} status reverted to Disponivel, IDs cleared.`);
        bookingsSheet.getRange(bookingRowIndex, bookingStatusCol + 1).setValue('Cancelada');
        Logger.log(`Booking ${trimmedBookingId} status updated to Cancelada.`);
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
                Logger.log(`WARNING: Failed to delete Calendar event ${existingEventId}: ${calError.message}`);
            }
        }
        createScheduleInstances(); // Isso parece ser um trigger para recriar/verificar instâncias.
        releaseScriptLock_(lock);
        lock = null;
        return createJsonResponse(true, `Reserva ${trimmedBookingId} cancelada com sucesso.`, { cancelledBookingId: trimmedBookingId });
    } catch (e) {
        Logger.log(`ERROR in cancelBookingAdmin for ID ${bookingIdToCancel} by user ${userEmail}: ${e.message}\nStack: ${e.stack}`);
        return createJsonResponse(false, `Falha ao cancelar reserva: ${e.message}`, { failedBookingId: bookingIdToCancel });
    } finally {
        releaseScriptLock_(lock);
    }
}

function bookSlot(jsonBookingDetailsString) {
    Logger.log(`*** bookSlot called ***`);
    let lock = null;
    let bookingId = null;
    let userEmail = '[Unavailable]';
    let lunchAlertMessage = null; // Para armazenar a mensagem de alerta de almoço
    try {
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
        // slotHoraInicio, slotTurma, slotData são passados de Javascript.html para o alerta de almoço
        const { idInstancia, tipoReserva, professorReal, disciplinaReal, tipoAulaReposicao, slotHoraInicio, slotTurma, slotData } = bookingDetails;

        const instanceIdToBook = String(idInstancia || '').trim();
        const bookingType = String(tipoReserva || '').trim(); // 'Reposicao' ou 'Substituicao'
        const profRealTrimmed = String(professorReal || '').trim();
        const discRealTrimmed = String(disciplinaReal || '').trim();
        const tipoAulaReposicaoTrimmed = (bookingType === TIPOS_RESERVA.REPOSICAO) ? String(tipoAulaReposicao || '').trim() : null;

        if (!instanceIdToBook) throw new Error('ID da instância de horário ausente ou inválido.');
        if (bookingType !== TIPOS_RESERVA.REPOSICAO && bookingType !== TIPOS_RESERVA.SUBSTITUICAO) throw new Error('Tipo de reserva inválido ou ausente.');

        if (bookingType === TIPOS_RESERVA.REPOSICAO) {
            if (!tipoAulaReposicaoTrimmed || (tipoAulaReposicaoTrimmed !== TIPOS_AULA_REPOSICAO.REPOSICAO && tipoAulaReposicaoTrimmed !== TIPOS_AULA_REPOSICAO.RECUPERACAO_PARALELA)) {
                throw new Error(`Para ${bookingType}, o tipo de aula (Reposição/Recuperação Paralela) é obrigatório e deve ser válido.`);
            }
            bookingDetails.tipoAulaReposicao = tipoAulaReposicaoTrimmed;
        }

        if (!profRealTrimmed) throw new Error('Professor é obrigatório.');
        if (!discRealTrimmed) throw new Error('Disciplina é obrigatória.');

        bookingDetails.professorReal = profRealTrimmed;
        bookingDetails.disciplinaReal = discRealTrimmed;

        lock = acquireScriptLock_();
        const processResult = processBooking_(bookingDetails, userEmail, userRole);
        bookingId = processResult.bookingId;

        // Lógica de Alerta de Almoço
        if (bookingType === TIPOS_RESERVA.REPOSICAO && slotHoraInicio && slotTurma && slotData) {
            Logger.log(`Verificando alerta de almoço para ${bookingType} (${tipoAulaReposicaoTrimmed}) às ${slotHoraInicio} na turma ${slotTurma} em ${slotData}`);
            if (checkLunchAlertNeeded_(slotHoraInicio, slotTurma, slotData, processResult.bookingDateObj)) {
                lunchAlertMessage = LUNCH_ALERT_MESSAGE;
                Logger.log(`Alerta de almoço ATIVADO para booking ID ${bookingId}.`);
            }
        }


        const calendarResult = handleCalendarIntegration_(
            getConfigValue('ID do Calendario'),
            bookingDetails, // bookingDetails já contém tipoAulaReposicao se for o caso
            processResult.instanceDetails,
            processResult.effectiveStartDateTime,
            processResult.guestEmails
        );
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
        sendBookingNotificationEmail_(
            bookingId,
            bookingType, // 'Reposicao' ou 'Substituicao'
            tipoAulaReposicaoTrimmed, // 'Reposição', 'Recuperação Paralela' ou null
            discRealTrimmed,
            processResult.turmaInstancia, // Usar a turma da instância do processResult
            profRealTrimmed,
            processResult.professorOriginal,
            processResult.effectiveStartDateTime,
            processResult.creationTimestamp,
            userEmail,
            calendarResult.eventId,
            calendarResult.error,
            processResult.guestEmails,
            lunchAlertMessage // Passa a mensagem de alerta de almoço para o e-mail
        );
        cleanupExcessVagoSlots();

        let successMessage = `Reserva ${bookingType === TIPOS_RESERVA.REPOSICAO ? tipoAulaReposicaoTrimmed : bookingType} (${bookingId}) agendada com sucesso!`;
        if (calendarResult.error) {
            successMessage += ` (Aviso: ${calendarResult.error.message || 'Erro ao integrar com Google Calendar.'})`;
        } else if (calendarResult.eventId) {
            successMessage += ` Evento no calendário criado/atualizado.`;
        } else {
            successMessage += ` Não foi possível gerar evento no calendário.`;
        }
        // Não adiciona o alerta de almoço à mensagem de sucesso principal, ele já foi enviado por email se necessário.
        // Se quiser mostrar no UI também, pode adicionar aqui.
        // if (lunchAlertMessage) {
        //     successMessage += ` ${lunchAlertMessage}`;
        // }
        successMessage += ` Notificação enviada.`;

        return createJsonResponse(true, successMessage, {
            bookingId: bookingId,
            eventId: calendarResult.eventId,
            lunchAlert: lunchAlertMessage // Retorna o alerta para o Javascript.html se quiser exibir lá
        });

    } catch (e) {
        Logger.log(`ERROR during bookSlot for user ${userEmail}: ${e.message}\nStack: ${e.stack}`);
        return createJsonResponse(false, `Falha no agendamento: ${e.message}`, { bookingId: bookingId });
    } finally {
        releaseScriptLock_(lock);
    }
}

function checkLunchAlertNeeded_(bookingTimeStr, bookingTurma, bookingDateStrDDMMYYYY, bookingDateObjUTC) {
    Logger.log(`checkLunchAlertNeeded_ - Hora: ${bookingTimeStr}, Turma: ${bookingTurma}, Data: ${bookingDateStrDDMMYYYY}`);
    if (!bookingTimeStr || !bookingTurma || !bookingDateObjUTC) {
        Logger.log("checkLunchAlertNeeded_ - Dados insuficientes para verificação.");
        return false;
    }

    const bookingHour = parseInt(bookingTimeStr.split(':')[0], 10);
    const bookingMinute = parseInt(bookingTimeStr.split(':')[1], 10);

    // Verifica se o horário é 13:15 ou posterior
    if (bookingHour < 13 || (bookingHour === 13 && bookingMinute < 15)) {
        Logger.log("checkLunchAlertNeeded_ - Horário anterior a 13:15. Sem alerta.");
        return false;
    }
    Logger.log("checkLunchAlertNeeded_ - Horário é 13:15 ou posterior.");

    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const { data: instanceData } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
    const { data: baseData } = getSheetData_(SHEETS.BASE_SCHEDULES);
    const { data: bookingDetailsData } = getSheetData_(SHEETS.BOOKING_DETAILS);

    const baseScheduleDisciplines = baseData.reduce((map, row) => {
        const baseId = String(row[HEADERS.BASE_SCHEDULES.ID] || '').trim();
        const disciplina = String(row[HEADERS.BASE_SCHEDULES.DISCIPLINA_PADRAO] || '').trim();
        if (baseId && disciplina) map[baseId] = disciplina;
        return map;
    }, {});

    const bookedDisciplines = bookingDetailsData.reduce((map, row) => {
        const instanceFk = String(row[HEADERS.BOOKING_DETAILS.ID_INSTANCIA] || '').trim();
        const disciplina = String(row[HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL] || '').trim();
        const statusReserva = String(row[HEADERS.BOOKING_DETAILS.STATUS_RESERVA] || '').trim();
        if (instanceFk && disciplina && statusReserva === 'Agendada') map[instanceFk] = disciplina;
        return map;
    }, {});

    const exemptDisciplinesConfig = getConfigValue(CONFIG_KEYS.DISCIPLINAS_ISENTAS_ALERTA_ALMOCO);
    const exemptDisciplines = exemptDisciplinesConfig ? exemptDisciplinesConfig.split(',').map(d => d.trim().toLowerCase()) : [];
    Logger.log(`checkLunchAlertNeeded_ - Disciplinas isentas: [${exemptDisciplines.join(', ')}]`);

    for (const row of instanceData) {
        const instanceDateUTC = formatValueToDate(row[HEADERS.SCHEDULE_INSTANCES.DATA]);
        const instanceTurma = String(row[HEADERS.SCHEDULE_INSTANCES.TURMA] || '').trim();

        if (!instanceDateUTC || instanceTurma !== bookingTurma) continue;

        // Compara apenas dia, mês e ano em UTC
        if (instanceDateUTC.getUTCFullYear() !== bookingDateObjUTC.getUTCFullYear() ||
            instanceDateUTC.getUTCMonth() !== bookingDateObjUTC.getUTCMonth() ||
            instanceDateUTC.getUTCDate() !== bookingDateObjUTC.getUTCDate()) {
            continue;
        }

        const instanceTimeStr = formatValueToHHMM(row[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO], timeZone);
        if (!instanceTimeStr) continue;

        const instanceHour = parseInt(instanceTimeStr.split(':')[0], 10);
        const instanceMinute = parseInt(instanceTimeStr.split(':')[1], 10);

        // Considerar apenas horários após o horário da reserva atual
        if (instanceHour > bookingHour || (instanceHour === bookingHour && instanceMinute > bookingMinute)) {
            const instanceStatus = String(row[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO] || '').trim();
            const instanceType = String(row[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL] || '').trim();
            const instanceId = String(row[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA] || '').trim();
            const baseId = String(row[HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO] || '').trim();

            let disciplineToCheck = '';

            if (instanceType === TIPOS_HORARIO.FIXO && instanceStatus === STATUS_OCUPACAO.DISPONIVEL) {
                disciplineToCheck = baseScheduleDisciplines[baseId] || '';
            } else if (instanceStatus === STATUS_OCUPACAO.REPOSICAO_AGENDADA || instanceStatus === STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA) {
                disciplineToCheck = bookedDisciplines[instanceId] || '';
            }

            if (disciplineToCheck && !exemptDisciplines.includes(disciplineToCheck.toLowerCase())) {
                Logger.log(`checkLunchAlertNeeded_ - Encontrada disciplina NÃO ISENTA "${disciplineToCheck}" às ${instanceTimeStr} para turma ${instanceTurma}. Sem alerta.`);
                return false; // Encontrou uma aula posterior não isenta, não precisa de alerta
            } else if (disciplineToCheck) {
                Logger.log(`checkLunchAlertNeeded_ - Encontrada disciplina ISENTA "${disciplineToCheck}" às ${instanceTimeStr}. Continuando verificação.`);
            }
        }
    }

    // Se chegou aqui, todas as aulas posteriores (se houver) são isentas, ou não há aulas posteriores.
    Logger.log(`checkLunchAlertNeeded_ - Nenhuma disciplina NÃO ISENTA encontrada após ${bookingTimeStr} para turma ${bookingTurma}. ALERTA NECESSÁRIO.`);
    return true;
}