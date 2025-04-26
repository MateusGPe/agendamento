/**
 * Arquivo: BaseSchedules.gs
 * Descrição: Funções para processar e validar dados da planilha 'Horarios Base'.
 */
function validateBaseSchedules_(baseData, timeZone) {
    const validSchedules = [];
    const daysOfWeek = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
    const idCol = HEADERS.BASE_SCHEDULES.ID;
    const dayCol = HEADERS.BASE_SCHEDULES.DIA_SEMANA;
    const hourCol = HEADERS.BASE_SCHEDULES.HORA_INICIO;
    const typeCol = HEADERS.BASE_SCHEDULES.TIPO;
    const turmaCol = HEADERS.BASE_SCHEDULES.TURMA_PADRAO;
    const profCol = HEADERS.BASE_SCHEDULES.PROFESSOR_PRINCIPAL;
    const reqCols = Math.max(idCol, dayCol, hourCol, typeCol, turmaCol, profCol) + 1;
    baseData.forEach((row, index) => {
        const rowIndex = index + 2;
        if (!row || row.length < reqCols) {
            return;
        }
        const baseId = String(row[idCol] || '').trim();
        const baseDayOfWeek = String(row[dayCol] || '').trim();
        const baseHourString = formatValueToHHMM(row[hourCol], timeZone);
        const baseType = String(row[typeCol] || '').trim();
        const baseTurma = String(row[turmaCol] || '').trim();
        const baseProfessorPrincipal = String(row[profCol] || '').trim();
        let isValid = true;
        const errorMessages = [];
        if (!baseId) { errorMessages.push("ID Base inválido"); isValid = false; }
        if (!baseDayOfWeek || !daysOfWeek.includes(baseDayOfWeek)) { errorMessages.push(`Dia da Semana inválido: ${baseDayOfWeek}`); isValid = false; }
        if (!baseHourString) { errorMessages.push(`Hora inválida: ${row[hourCol]}`); isValid = false; }
        if (baseType !== TIPOS_HORARIO.FIXO && baseType !== TIPOS_HORARIO.VAGO) { errorMessages.push(`Tipo inválido: ${baseType}`); isValid = false; }
        if (!baseTurma) { errorMessages.push("Turma Padrão inválida"); isValid = false; }
        if (baseType === TIPOS_HORARIO.FIXO && !baseProfessorPrincipal) { errorMessages.push("Professor Principal ausente para horário Fixo"); isValid = false; }
        if (isValid) {
            validSchedules.push({
                id: baseId,
                dayOfWeek: baseDayOfWeek,
                hour: baseHourString,
                type: baseType,
                turma: baseTurma,
                professorPrincipal: baseProfessorPrincipal
            });
        } else {
            Logger.log(`Skipping Base Schedule row ${rowIndex}: ${errorMessages.join(', ')}.`);
        }
    });
    return validSchedules;
}