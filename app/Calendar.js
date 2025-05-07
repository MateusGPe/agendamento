/**
 * Arquivo: Calendar.gs
 * Descrição: Funções para integração com o Google Calendar (criação/atualização de eventos).
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
        
        // Usa tipoAulaReposicao se existir (para Reposições/Recuperações), senão usa tipoReserva (para Substituições)
        const eventDisplayType = bookingDetails.tipoAulaReposicao || bookingDetails.tipoReserva;
        
        const disciplina = String(bookingDetails.disciplinaReal || '').trim();
        const turma = String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.TURMA] || '').trim();
        const profReal = String(bookingDetails.professorReal || '').trim();
        
        // Para Substituição, o professor original é pego da reserva ou da instância.
        // Para Reposição/Recuperação, não há "professor original" no mesmo sentido, mas podemos pegar da instância se existir.
        let profOrig = '';
        if (bookingDetails.tipoReserva === TIPOS_RESERVA.SUBSTITUICAO) {
             profOrig = String(bookingDetails.professorOriginalNaReserva || instanceDetails[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL] || '').trim();
        } else if (bookingDetails.tipoReserva === TIPOS_RESERVA.REPOSICAO) {
            // Para reposição, o "original" seria o professor base do horário, se houver (geralmente não há para VAGO).
            // Mas não faz muito sentido exibir no título do evento.
            // Se a lógica precisar, pode-se pegar de instanceDetails[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL]
        }

        const bookingId = String(instanceDetails[HEADERS.SCHEDULE_INSTANCES.ID_RESERVA] || '').trim();
        // Tenta obter o criador da reserva (CRIADO_POR) que agora está na instância,
        // ou cai para o usuário ativo se não encontrar.
        const userEmail = String(instanceDetails[HEADERS.BOOKING_DETAILS.CRIADO_POR] || getActiveUserEmail_());
        
        const eventTitle = `${eventDisplayType} - ${disciplina} (${turma})`;
        
        let eventDescription = `Reserva ID: ${bookingId}\nProfessor: ${profReal}`;
        if (bookingDetails.tipoReserva === TIPOS_RESERVA.SUBSTITUICAO && profOrig && profOrig !== profReal) {
            eventDescription += ` (Original: ${profOrig})`;
        }
        // Adiciona o tipo de aula para Reposição/Recuperação na descrição, se aplicável
        if (bookingDetails.tipoReserva === TIPOS_RESERVA.REPOSICAO && bookingDetails.tipoAulaReposicao) {
            eventDescription += `\nTipo de Aula: ${bookingDetails.tipoAulaReposicao}`;
        }
        eventDescription += `\nTurma: ${turma}\nAgendado por: ${userEmail}`;
        
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