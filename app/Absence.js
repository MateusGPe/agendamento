/**
 * Arquivo: Absence.gs
 * Descrição: Funções para registrar ausência de professores em horários fixos.
 */
function reportAbsence(instanceId, professorNameToMarkAbsent) {
    Logger.log(`*** reportAbsence called for Instance ID: ${instanceId}, Professor: ${professorNameToMarkAbsent} ***`);
    let lock = null;
    let userEmail = '[Unknown]';
    try {
        userEmail = getActiveUserEmail_();
        const userRole = getUserRolePlain_(userEmail);
        const userProfessorName = (userRole === USER_ROLES.PROFESSOR) ? getProfessorNameByEmail_(userEmail) : null;
        const cancelAdminProfessorEmailRaw = getConfigValue('Professor Admin Cancelamento');
        const cancelAdminProfessorEmail = cancelAdminProfessorEmailRaw ? cancelAdminProfessorEmailRaw.trim().toLowerCase() : null;
        const isPrivilegedUser = (userRole === USER_ROLES.ADMIN) || (cancelAdminProfessorEmail && userEmail.toLowerCase() === cancelAdminProfessorEmail);
        if (!userRole) throw new Error('Usuário não autorizado.');
        if (!instanceId || typeof instanceId !== 'string' || instanceId.trim() === '') throw new Error('ID da Instância inválido ou ausente.');
        if (!professorNameToMarkAbsent || typeof professorNameToMarkAbsent !== 'string' || professorNameToMarkAbsent.trim() === '') throw new Error('Nome do Professor ausente inválido ou ausente.');
        const trimmedInstanceId = instanceId.trim();
        const trimmedProfessorName = professorNameToMarkAbsent.trim();
        lock = acquireScriptLock_();
        const { header: instanceHeader, data: instanceData, sheet: instancesSheet } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
        const instanceIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA;
        const typeCol = HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL;
        const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
        const profPrincCol = HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL;
        const absentCol = HEADERS.SCHEDULE_INSTANCES.PROFESSORES_AUSENTES;
        const maxIndexNeeded = Math.max(instanceIdCol, typeCol, dateCol, profPrincCol, absentCol);
        if (instanceHeader.length <= maxIndexNeeded) {
            throw new Error(`A planilha "${SHEETS.SCHEDULE_INSTANCES}" não contém a coluna necessária para Ausências (Índice ${absentCol}). Atualize as definições de cabeçalho (HEADERS).`);
        }
        let instanceRowIndex = -1;
        let instanceDetails = null;
        for (let i = 0; i < instanceData.length; i++) {
            const row = instanceData[i];
            if (row && row.length > instanceIdCol && String(row[instanceIdCol] || '').trim() === trimmedInstanceId) {
                instanceRowIndex = i + 2;
                instanceDetails = row;
                break;
            }
        }
        if (instanceRowIndex === -1 || !instanceDetails) {
            throw new Error(`Instância de Horário com ID ${trimmedInstanceId} não encontrada.`);
        }
        const instanceType = String(instanceDetails[typeCol] || '').trim();
        const instanceDateRaw = instanceDetails[dateCol];
        const instanceProfessorsString = String(instanceDetails[profPrincCol] || '').trim();
        if (instanceType !== TIPOS_HORARIO.FIXO) {
            throw new Error(`Só é possível registrar ausência para horários do tipo "${TIPOS_HORARIO.FIXO}". Este é do tipo "${instanceType}".`);
        }
        const instanceDate = formatValueToDate(instanceDateRaw);
        const today = new Date();
        const todayUTC = new Date(Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()));
        if (!instanceDate || instanceDate < todayUTC) {
            throw new Error(`Só é possível registrar ausência para datas futuras. Data do horário: ${instanceDate ? Utilities.formatDate(instanceDate, Session.getScriptTimeZone(), 'dd/MM/yyyy') : 'Inválida'}.`);
        }
        const instanceProfessorsList = instanceProfessorsString.split(',').map(p => p.trim()).filter(p => p !== '');
        if (instanceProfessorsList.length === 0) {
            throw new Error(`Não há professores principais definidos para esta instância (${trimmedInstanceId}).`);
        }
        if (!instanceProfessorsList.includes(trimmedProfessorName)) {
            throw new Error(`O professor "${trimmedProfessorName}" não está listado como professor principal para este horário (${instanceProfessorsList.join(', ')}).`);
        }
        let canReport = false;
        if (isPrivilegedUser) {
            canReport = true;
            Logger.log(`User ${userEmail} (Privileged) reporting absence for ${trimmedProfessorName} in instance ${trimmedInstanceId}.`);
        } else if (userRole === USER_ROLES.PROFESSOR && userProfessorName) {
            if (instanceProfessorsList.includes(userProfessorName) && trimmedProfessorName === userProfessorName) {
                canReport = true;
                Logger.log(`Professor ${userEmail} (${userProfessorName}) reporting own absence for instance ${trimmedInstanceId}.`);
            } else if (!instanceProfessorsList.includes(userProfessorName)) {
                Logger.log(`Professor ${userEmail} (${userProfessorName}) tried to report absence, but is not listed in instance ${trimmedInstanceId} professors: [${instanceProfessorsList.join(', ')}].`);
                throw new Error(`Você (${userProfessorName}) não está listado como professor principal para este horário e só pode informar a própria falta.`);
            } else {
                Logger.log(`Professor ${userEmail} (${userProfessorName}) tried to report absence for ${trimmedProfessorName} (not self) in instance ${trimmedInstanceId}.`);
                throw new Error(`Você (${userProfessorName}) só pode informar a própria falta.`);
            }
        } else {
            Logger.log(`User ${userEmail} (Role: ${userRole}) has no permission to report absence.`);
            throw new Error('Permissão negada para registrar ausência.');
        }
        if (!canReport) {
            throw new Error('Permissão negada.');
        }
        const currentAbsentString = String(instanceDetails[absentCol] || '').trim();
        const currentAbsentList = currentAbsentString.split(',').map(p => p.trim()).filter(p => p !== '');
        if (currentAbsentList.includes(trimmedProfessorName)) {
            Logger.log(`Professor ${trimmedProfessorName} already marked as absent for instance ${trimmedInstanceId}. No change needed.`);
            releaseScriptLock_(lock);
            lock = null;
            return createJsonResponse(true, `Professor ${trimmedProfessorName} já estava marcado como ausente para este horário.`, { instanceId: trimmedInstanceId });
        }
        currentAbsentList.push(trimmedProfessorName);
        const newAbsentString = currentAbsentList.sort().join(',');
        instancesSheet.getRange(instanceRowIndex, absentCol + 1).setValue(newAbsentString);
        SpreadsheetApp.flush();
        invalidateSheetCache_(SHEETS.SCHEDULE_INSTANCES);
        Logger.log(`Absence reported for ${trimmedProfessorName} in instance ${trimmedInstanceId}. New absent list: [${newAbsentString}]`);
        releaseScriptLock_(lock);
        lock = null;
        return createJsonResponse(true, `Ausência de ${trimmedProfessorName} registrada com sucesso para o horário.`, { instanceId: trimmedInstanceId });
    } catch (e) {
        Logger.log(`ERROR in reportAbsence for Instance ${instanceId}, Professor ${professorNameToMarkAbsent} by user ${userEmail}: ${e.message}\nStack: ${e.stack}`);
        return createJsonResponse(false, `Falha ao registrar ausência: ${e.message}`, { instanceId: instanceId });
    } finally {
        releaseScriptLock_(lock);
    }
}