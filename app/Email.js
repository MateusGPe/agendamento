/**
 * Arquivo: Email.gs
 * Descrição: Funções para criar e enviar emails de notificação de reserva.
 */
function createBookingEmailContent_(bookingId, bookingType, tipoAulaReposicao, disciplina, turma, profReal, profOrig, startTime, timestampCriacao, userEmail, calendarEventId, calendarError, lunchAlertMessage) {
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const dataFormatada = Utilities.formatDate(startTime, timeZone, 'dd/MM/yyyy');
    const horaFormatada = Utilities.formatDate(startTime, timeZone, 'HH:mm');
    const criacaoFormatada = Utilities.formatDate(timestampCriacao, timeZone, 'dd/MM/yyyy HH:mm:ss');
    
    // Define o tipo de evento para o assunto e corpo do email
    // Se tipoAulaReposicao estiver presente, usa ele, senão usa bookingType (para Substituições)
    const displayEventType = tipoAulaReposicao || bookingType;

    let subjectStatus = calendarError ? '⚠️ Erro no Google Calendar' : (calendarEventId ? '✅ Confirmada' : '✅ Sem Evento Calendar');
    let subject = `${subjectStatus} - Reserva ${displayEventType} - ${disciplina || 'N/D'} - ${dataFormatada}`;
    
    let bodyText = `Olá,\n\nUma reserva de "${displayEventType}" foi registrada no sistema:\n\n`;
    bodyText += `==============================\nDETALHES DA RESERVA\n==============================\n`;
    bodyText += `Tipo: ${displayEventType}\nData: ${dataFormatada}\nHorário: ${horaFormatada}\nTurma: ${turma || 'N/A'}\n`;
    bodyText += `Disciplina: ${disciplina || 'N/A'}\nProfessor: ${profReal || 'N/A'}\n`;
    if (bookingType === TIPOS_RESERVA.SUBSTITUICAO && profOrig && profOrig !== profReal) {
        bodyText += `Professor Original: ${profOrig}\n`;
    }
    bodyText += `------------------------------\nID Reserva: ${bookingId}\nAgendado por: ${userEmail}\nData/Hora Agend.: ${criacaoFormatada}\n`;
    bodyText += `==============================\n\n`;

    if (lunchAlertMessage) {
        bodyText += `*** ATENÇÃO IMPORTANTE ***\n${lunchAlertMessage}\n\n`;
    }

    if (calendarError) {
        bodyText += `*** ATENÇÃO: Google Calendar ***\nHouve um erro ao criar/atualizar o evento no calendário.\nA reserva está confirmada nas planilhas, mas verifique o calendário manualmente.\nErro: ${calendarError.message}\n\n`;
    } else if (calendarEventId) {
        bodyText += `Evento no Google Calendar criado/atualizado com sucesso (ID: ${calendarEventId}).\n\n`;
    } else {
        bodyText += `*** AVISO: Google Calendar ***\nO evento não foi criado/atualizado no calendário (ID do calendário pode não estar configurado ou calendário inacessível).\n\n`;
    }
    bodyText += `Atenciosamente,\n${EMAIL_SENDER_NAME}`;

    let bodyHtml = `<p>Olá,</p><p>Uma reserva de "<b>${displayEventType}</b>" foi registrada no sistema:</p><hr><h3>Detalhes da Reserva</h3>`;
    bodyHtml += `<table border="0" cellpadding="5" style="border-collapse: collapse; font-family: sans-serif; font-size: 11pt;">`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Tipo:</strong></td><td>${displayEventType}</td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Data:</strong></td><td>${dataFormatada}</td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Horário:</strong></td><td>${horaFormatada}</td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Turma:</strong></td><td>${turma || 'N/A'}</td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Disciplina:</strong></td><td>${disciplina || 'N/A'}</td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Professor:</strong></td><td>${profReal || 'N/A'}</td></tr>`;
    if (bookingType === TIPOS_RESERVA.SUBSTITUICAO && profOrig && profOrig !== profReal) {
        bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Prof. Original:</strong></td><td>${profOrig}</td></tr>`;
    }
    bodyHtml += `</table><br>`;
    bodyHtml += `<table border="0" cellpadding="5" style="border-collapse: collapse; font-family: sans-serif; font-size: 9pt; color: #555;">`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><i>ID Reserva:</i></td><td><i>${bookingId}</i></td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><i>Agendado por:</i></td><td><i>${userEmail}</i></td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><i>Data/Hora Agend.:</i></td><td><i>${criacaoFormatada}</i></td></tr>`;
    bodyHtml += `</table><hr>`;

    if (lunchAlertMessage) {
        bodyHtml += `<div style="border: 1px solid #FF8C00; background-color: #FFF3CD; color: #856404; padding: 15px; margin-top: 15px; border-radius: 4px; font-family: sans-serif; font-size: 11pt;">`;
        bodyHtml += `<strong>*** ATENÇÃO IMPORTANTE ***</strong><br>${lunchAlertMessage}</div>`;
    }

    if (calendarError) {
        bodyHtml += `<div style="border: 1px solid #DC3545; background-color: #F8D7DA; color: #721C24; padding: 15px; margin-top: 15px; border-radius: 4px; font-family: sans-serif; font-size: 11pt;">`;
        bodyHtml += `<strong>*** ATENÇÃO: Google Calendar ***</strong><br>Houve um erro ao criar/atualizar o evento no calendário.<br>A reserva está confirmada nas planilhas, mas verifique o calendário manualmente.<br><span style="font-size: 9pt; color: #721C24;">Erro: ${calendarError.message}</span></div>`;
    } else if (calendarEventId) {
        bodyHtml += `<div style="border: 1px solid #28A745; background-color: #D4EDDA; color: #155724; padding: 15px; margin-top: 15px; border-radius: 4px; font-family: sans-serif; font-size: 11pt;">`;
        bodyHtml += `Evento no Google Calendar criado/atualizado com sucesso.<br>ID do Evento: ${calendarEventId}</div>`;
    } else {
        bodyHtml += `<div style="border: 1px solid #FFC107; background-color: #FFF3CD; color: #856404; padding: 15px; margin-top: 15px; border-radius: 4px; font-family: sans-serif; font-size: 11pt;">`;
        bodyHtml += `<strong>*** AVISO: Google Calendar ***</strong><br>O evento não foi criado/atualizado no calendário (ID do calendário pode não estar configurado ou calendário inacessível).</div>`;
    }
    bodyHtml += `<p style="font-family: sans-serif; font-size: 11pt; margin-top: 20px;">Atenciosamente,<br>${EMAIL_SENDER_NAME}</p>`;
    return { subject, bodyText, bodyHtml };
}

function sendBookingNotificationEmail_(bookingId, bookingType, tipoAulaReposicao, disciplina, turma, profReal, profOrig, startTime, timestampCriacao, userEmail, calendarEventId, calendarError, guests, lunchAlertMessage) {
    Logger.log("sendBookingNotificationEmail_ called.");
    try {
        const emailContent = createBookingEmailContent_(bookingId, bookingType, tipoAulaReposicao, disciplina, turma, profReal, profOrig, startTime, timestampCriacao, userEmail, calendarEventId, calendarError, lunchAlertMessage);
        
        const validGuests = Array.isArray(guests) ? guests.filter(email => email && typeof email === 'string' && email.includes('@')) : [];
        const validAdminEmails = ADMIN_COPY_EMAILS.filter(email => email && typeof email === 'string' && email.includes('@'));
        
        // Garantir que o criador da reserva receba o e-mail, mesmo que não esteja na lista de convidados ou admins.
        const allRecipients = new Set([...validGuests, ...validAdminEmails]);
        if (userEmail && userEmail.includes('@')) {
            allRecipients.add(userEmail.toLowerCase()); // Adiciona o criador
        }

        const finalRecipientsBcc = Array.from(allRecipients);

        if (finalRecipientsBcc.length === 0) {
            Logger.log("No valid recipients found. Skipping email send.");
            return;
        }
        
        // O campo "to" pode ser um dos admins ou o próprio criador se nenhum admin estiver listado
        let toAddress = validAdminEmails.length > 0 ? validAdminEmails[0] : (userEmail && userEmail.includes('@') ? userEmail : null);

        if (!toAddress && finalRecipientsBcc.length > 0) {
            // Se `toAddress` ainda for nulo, mas houver destinatários em BCC, use o primeiro BCC como `to`.
            toAddress = finalRecipientsBcc[0];
            Logger.log(`Using first BCC recipient ${toAddress} as 'To' address as no primary admin/creator was suitable.`);
        } else if (!toAddress) {
            Logger.log("No valid 'To' address could be determined. Skipping email send.");
            return;
        }
        
        // Remover o 'toAddress' da lista de BCC para evitar duplicação se ele já estiver lá
        const bccListForSend = finalRecipientsBcc.filter(email => email.toLowerCase() !== toAddress.toLowerCase());

        Logger.log(`Sending notification for Booking ID ${bookingId}. To: ${toAddress}, BCC: ${bccListForSend.length} addresses ([${bccListForSend.join(', ')}]). Original AllRecipients: ${finalRecipientsBcc.length}`);
        
        MailApp.sendEmail({
            to: toAddress,
            bcc: bccListForSend.join(','), // pode ser string vazia se não houver mais ninguém no BCC
            subject: emailContent.subject,
            body: emailContent.bodyText,
            htmlBody: emailContent.bodyHtml,
            name: EMAIL_SENDER_NAME
        });
        Logger.log(`Email notification for Booking ID ${bookingId} sent successfully via MailApp.`);
    } catch (e) {
        Logger.log(`ERROR sending booking notification email for ID ${bookingId}: ${e.message}\nStack: ${e.stack}`);
    }
}