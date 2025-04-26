/**
 * Arquivo: Email.gs
 * Descrição: Funções para criar e enviar emails de notificação de reserva.
 */
function createBookingEmailContent_(bookingId, bookingType, disciplina, turma, profReal, profOrig, startTime, timestampCriacao, userEmail, calendarEventId, calendarError) {
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const dataFormatada = Utilities.formatDate(startTime, timeZone, 'dd/MM/yyyy');
    const horaFormatada = Utilities.formatDate(startTime, timeZone, 'HH:mm');
    const criacaoFormatada = Utilities.formatDate(timestampCriacao, timeZone, 'dd/MM/yyyy HH:mm:ss');
    const isSubstituicao = bookingType === TIPOS_RESERVA.SUBSTITUICAO;
    let subjectStatus = calendarError ? '⚠️ Erro no Google Calendar' : (calendarEventId ? '✅ Confirmada' : '✅ Sem Evento Calendar');
    let subject = `${subjectStatus} - Reserva ${bookingType} - ${disciplina || 'N/D'} - ${dataFormatada}`;
    let bodyText = `Olá,\n\nUma reserva de "${bookingType}" foi registrada no sistema:\n\n`;
    bodyText += `==============================\nDETALHES DA RESERVA\n==============================\n`;
    bodyText += `Tipo: ${bookingType}\nData: ${dataFormatada}\nHorário: ${horaFormatada}\nTurma: ${turma || 'N/A'}\n`;
    bodyText += `Disciplina: ${disciplina || 'N/A'}\nProfessor: ${profReal || 'N/A'}\n`;
    if (isSubstituicao && profOrig && profOrig !== profReal) bodyText += `Professor Original: ${profOrig}\n`;
    bodyText += `------------------------------\nID Reserva: ${bookingId}\nAgendado por: ${userEmail}\nData/Hora Agend.: ${criacaoFormatada}\n`;
    bodyText += `==============================\n\n`;
    if (calendarError) {
        bodyText += `*** ATENÇÃO: Google Calendar ***\nHouve um erro ao criar/atualizar o evento no calendário.\nA reserva está confirmada nas planilhas, mas verifique o calendário manualmente.\nErro: ${calendarError.message}\n\n`;
    } else if (calendarEventId) {
        bodyText += `Evento no Google Calendar criado/atualizado com sucesso (ID: ${calendarEventId}).\n\n`;
    } else {
        bodyText += `*** AVISO: Google Calendar ***\nO evento não foi criado/atualizado no calendário (ID do calendário pode não estar configurado ou calendário inacessível).\n\n`;
    }
    bodyText += `Atenciosamente,\n${EMAIL_SENDER_NAME}`;
    let bodyHtml = `<p>Olá,</p><p>Uma reserva de "<b>${bookingType}</b>" foi registrada no sistema:</p><hr><h3>Detalhes da Reserva</h3>`;
    bodyHtml += `<table border="0" cellpadding="5" style="border-collapse: collapse; font-family: sans-serif; font-size: 11pt;">`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Tipo:</strong></td><td>${bookingType}</td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Data:</strong></td><td>${dataFormatada}</td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Horário:</strong></td><td>${horaFormatada}</td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Turma:</strong></td><td>${turma || 'N/A'}</td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Disciplina:</strong></td><td>${disciplina || 'N/A'}</td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Professor:</strong></td><td>${profReal || 'N/A'}</td></tr>`;
    if (isSubstituicao && profOrig && profOrig !== profReal) {
        bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><strong>Prof. Original:</strong></td><td>${profOrig}</td></tr>`;
    }
    bodyHtml += `</table><br>`;
    bodyHtml += `<table border="0" cellpadding="5" style="border-collapse: collapse; font-family: sans-serif; font-size: 9pt; color: #555;">`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><i>ID Reserva:</i></td><td><i>${bookingId}</i></td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><i>Agendado por:</i></td><td><i>${userEmail}</i></td></tr>`;
    bodyHtml += `<tr><td style="text-align: right; vertical-align: top; padding-right: 10px;"><i>Data/Hora Agend.:</i></td><td><i>${criacaoFormatada}</i></td></tr>`;
    bodyHtml += `</table><hr>`;
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
function sendBookingNotificationEmail_(bookingId, bookingType, disciplina, turma, profReal, profOrig, startTime, timestampCriacao, userEmail, calendarEventId, calendarError, guests) {
    Logger.log("sendBookingNotificationEmail_ called.");
    try {
        const emailContent = createBookingEmailContent_(bookingId, bookingType, disciplina, turma, profReal, profOrig, startTime, timestampCriacao, userEmail, calendarEventId, calendarError);
        const validGuests = Array.isArray(guests) ? guests.filter(email => email && typeof email === 'string' && email.includes('@')) : [];
        const validAdminEmails = ADMIN_COPY_EMAILS.filter(email => email && typeof email === 'string' && email.includes('@'));
        const finalRecipientsBcc = [...new Set([...validGuests, ...validAdminEmails])];
        if (finalRecipientsBcc.length === 0) {
            Logger.log("No valid recipients found. Skipping email send.");
            return;
        }
        const toAddress = validAdminEmails[0] || userEmail;
        Logger.log(`Sending notification for Booking ID ${bookingId}. To: ${toAddress}, BCC: ${finalRecipientsBcc.length} addresses.`);
        MailApp.sendEmail({
            to: toAddress,
            bcc: finalRecipientsBcc.join(','),
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