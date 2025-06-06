/**
 * Arquivo: ScheduleInstances.gs
 * Descrição: Funções para criar, filtrar e limpar instâncias de horários gerados a partir dos horários base.
 */
function createExistingInstanceMap_(instanceData, timeZone) {
  const existingKeys = {};
  const baseIdCol = HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO;
  const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
  const hourCol = HEADERS.SCHEDULE_INSTANCES.HORA_INICIO;
  const reqCols = Math.max(baseIdCol, dateCol, hourCol) + 1;
  instanceData.forEach((row, index) => {
    if (!row || row.length < reqCols) return;
    const baseId = String(row[baseIdCol] || '').trim();
    const instanceUTCDate = formatValueToDate(row[dateCol]);
    const hourString = formatValueToHHMM(row[hourCol], timeZone);
    if (baseId && instanceUTCDate && hourString) {
      const dateString = Utilities.formatDate(instanceUTCDate, 'UTC', 'yyyy-MM-dd');
      const key = `${baseId}_${dateString}_${hourString}`;
      existingKeys[key] = true;
    } else {
    }
  });
  return existingKeys;
}
function calculateGenerationRange_(numWeeksToGenerate) {
  const now = new Date();
  const todayUTC = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate() - 14));
  const dayUTC = todayUTC.getUTCDay();
  const daysToAdd = (dayUTC === 0) ? 1 : (8 - dayUTC) % 7;
  const start = new Date(todayUTC.getTime());
  start.setUTCDate(todayUTC.getUTCDate() + daysToAdd);
  const end = new Date(start.getTime());
  end.setUTCDate(start.getUTCDate() + (numWeeksToGenerate * 7) - 1);
  return { startGenerationDate: start, endGenerationDate: end };
}
function buildDailyTurmaCounts_(instanceData) {
  const dailyTurmaCounts = {};
  const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
  const turmaCol = HEADERS.SCHEDULE_INSTANCES.TURMA;
  const typeCol = HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL;
  const statusCol = HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO;
  const maxIndexNeeded = Math.max(dateCol, turmaCol, typeCol, statusCol);
  if (!instanceData || instanceData.length === 0) {
    return dailyTurmaCounts;
  }
  Logger.log(`Building daily turma counts from ${instanceData.length} existing instances...`);
  instanceData.forEach((row) => {
    if (!row || row.length <= maxIndexNeeded) return;
    const instanceUTCDate = formatValueToDate(row[dateCol]);
    const turma = String(row[turmaCol] || '').trim();
    const originalType = String(row[typeCol] || '').trim();
    const instanceStatus = String(row[statusCol] || '').trim();
    if (!instanceUTCDate || !turma) return;
    const dateStringKey = Utilities.formatDate(instanceUTCDate, 'UTC', 'yyyy-MM-dd');
    if (!dailyTurmaCounts[dateStringKey]) {
      dailyTurmaCounts[dateStringKey] = {};
    }
    if (!dailyTurmaCounts[dateStringKey][turma]) {
      dailyTurmaCounts[dateStringKey][turma] = { fixedCount: 0, vagoBookedCount: 0 };
    }
    if (originalType === TIPOS_HORARIO.FIXO) {
      dailyTurmaCounts[dateStringKey][turma].fixedCount++;
    } else if (originalType === TIPOS_HORARIO.VAGO) {
      // Considera 'Reposicao Agendada' como vago ocupado
      if (instanceStatus === STATUS_OCUPACAO.REPOSICAO_AGENDADA) {
        dailyTurmaCounts[dateStringKey][turma].vagoBookedCount++;
      }
    }
  });
  Logger.log(`Finished building daily turma counts. Found counts for ${Object.keys(dailyTurmaCounts).length} dates.`);
  return dailyTurmaCounts;
}
function generateNewInstances_(startDateUTC, endDateUTC, validBaseSchedules, existingInstanceKeys, timeZone, initialDailyTurmaCounts, threshold) {
  const newInstanceRows = [];
  const daysOfWeekMap = ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado'];
  const numInstanceCols = Math.max(...Object.values(HEADERS.SCHEDULE_INSTANCES)) + 1;
  Logger.log(`Generating new instances. Threshold (${threshold}) prevents VAGO creation if initial Fixed+BookedVago count is >= threshold.`);
  let skippedVagoDueToThreshold = 0;
  let currentDate = new Date(startDateUTC.getTime());
  while (currentDate <= endDateUTC) {
    const targetUTCDate = new Date(currentDate.getTime());
    const targetDayName = daysOfWeekMap[targetUTCDate.getUTCDay()];
    const targetDateStr = Utilities.formatDate(targetUTCDate, 'UTC', 'yyyy-MM-dd');
    const applicableSchedules = validBaseSchedules.filter(s => s.dayOfWeek === targetDayName);
    applicableSchedules.sort((a, b) => a.hour.localeCompare(b.hour));
    for (const baseSchedule of applicableSchedules) {
      const predictableKey = `${baseSchedule.id}_${targetDateStr}_${baseSchedule.hour}`;
      const turma = baseSchedule.turma;
      const baseType = baseSchedule.type;
      if (existingInstanceKeys[predictableKey]) {
        continue;
      }
      let allowCreation = true;
      if (baseType === TIPOS_HORARIO.VAGO) {
        const initialCounts = (initialDailyTurmaCounts[targetDateStr] && initialDailyTurmaCounts[targetDateStr][turma])
          ? initialDailyTurmaCounts[targetDateStr][turma]
          : { fixedCount: 0, vagoBookedCount: 0 };
        const initialTotalConsidered = initialCounts.fixedCount + initialCounts.vagoBookedCount;
        if (initialTotalConsidered >= threshold) {
          allowCreation = false;
          skippedVagoDueToThreshold++;
        }
      }
      if (allowCreation) {
        const newRow = new Array(numInstanceCols).fill('');
        newRow[HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA] = Utilities.getUuid();
        newRow[HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO] = baseSchedule.id;
        newRow[HEADERS.SCHEDULE_INSTANCES.TURMA] = turma;
        newRow[HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL] = baseSchedule.professorPrincipal;
        newRow[HEADERS.SCHEDULE_INSTANCES.DATA] = new Date(targetUTCDate.getUTCFullYear(), targetUTCDate.getUTCMonth(), targetUTCDate.getUTCDate());
        newRow[HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA] = targetDayName;
        newRow[HEADERS.SCHEDULE_INSTANCES.HORA_INICIO] = baseSchedule.hour;
        newRow[HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL] = baseType;
        newRow[HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO] = STATUS_OCUPACAO.DISPONIVEL;
        newRow[HEADERS.SCHEDULE_INSTANCES.PROFESSORES_AUSENTES] = '';
        newInstanceRows.push(newRow);
        existingInstanceKeys[predictableKey] = true;
      }
    }
    currentDate.setUTCDate(currentDate.getUTCDate() + 1);
  }
  if (skippedVagoDueToThreshold > 0) {
    Logger.log(`Skipped creating ${skippedVagoDueToThreshold} VAGO instances due to the initial Fixed+Booked count meeting/exceeding the threshold (${threshold}).`);
  }
  return newInstanceRows;
}
function createScheduleInstances() {
  Logger.log('*** createScheduleInstances START ***');
  let lock = null;
  try {
    lock = acquireScriptLock_();
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const { header: baseHeader, data: baseData } = getSheetData_(SHEETS.BASE_SCHEDULES, HEADERS.BASE_SCHEDULES);
    const { header: instanceHeader, data: instanceData, sheet: instancesSheet } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
    if (baseData.length === 0) {
      Logger.log('Planilha "Horarios Base" está vazia. Saindo.');
      releaseScriptLock_(lock);
      return;
    };
    const validBaseSchedules = validateBaseSchedules_(baseData, timeZone);
    if (validBaseSchedules.length === 0) {
      Logger.log('Nenhum horário base válido encontrado após validação. Saindo.');
      releaseScriptLock_(lock);
      return;
    }
    Logger.log(`Found ${validBaseSchedules.length} valid base schedules.`);
    const existingInstanceKeys = createExistingInstanceMap_(instanceData, timeZone);
    Logger.log(`Created map with ${Object.keys(existingInstanceKeys).length} existing instance keys.`);
    const initialDailyTurmaCounts = buildDailyTurmaCounts_(instanceData);
    Logger.log(`Built initial daily counts for ${Object.keys(initialDailyTurmaCounts).length} dates.`);
    const numWeeksToGenerate = parseInt(getConfigValue('Semanas Para Gerar Instancias')) || 4;
    let threshold = parseInt(getConfigValue('Limite Maximo Aulas Dia Turma')) || 10;
    if (isNaN(threshold) || threshold <= 0) {
      Logger.log(`WARNING: Invalid threshold value found in config. Using default of 10.`);
      threshold = 10;
    }
    const { startGenerationDate, endGenerationDate } = calculateGenerationRange_(numWeeksToGenerate);
    Logger.log(`Generating instances from UTC ${startGenerationDate.toISOString().slice(0, 10)} to ${endGenerationDate.toISOString().slice(0, 10)} with threshold ${threshold} (for Vago creation)`);
    const newInstanceRows = generateNewInstances_(
      startGenerationDate,
      endGenerationDate,
      validBaseSchedules,
      existingInstanceKeys,
      timeZone,
      initialDailyTurmaCounts,
      threshold
    );
    if (newInstanceRows.length > 0) {
      Logger.log(`Generated ${newInstanceRows.length} new instances (respecting threshold for Vago slots). Appending to sheet...`);
      const numInstanceCols = instanceHeader.length || (Math.max(...Object.values(HEADERS.SCHEDULE_INSTANCES)) + 1);
      appendSheetRows_(instancesSheet, numInstanceCols, newInstanceRows);
    } else {
      Logger.log('No new instances needed or allowed (for Vago type) for the specified period based on existing data and threshold.');
    }
    Logger.log('*** createScheduleInstances FINISHED ***');
  } catch (e) {
    Logger.log(`ERROR in createScheduleInstances: ${e.message}\nStack: ${e.stack}`);
  } finally {
    releaseScriptLock_(lock);
  }
}
function getFilteredScheduleInstances(turma, weekStartDateString) {
  Logger.log(`*** getFilteredScheduleInstances called for Turma: ${turma}, Semana (expecting UTC Monday): ${weekStartDateString} ***`);
  try {
    getActiveUserEmail_();
    const trimmedTurma = String(turma || '').trim();
    if (!trimmedTurma) {
      return createJsonResponse(false, 'Turma não especificada.', null);
    }
    if (!weekStartDateString || typeof weekStartDateString !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(weekStartDateString)) {
      return createJsonResponse(false, 'Semana de início inválida ou formato incorreto (esperado YYYY-MM-DD).', null);
    }
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const parts = weekStartDateString.split('-');
    const weekStartDate = new Date(Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10)));
    if (isNaN(weekStartDate.getTime())) {
      return createJsonResponse(false, `Data de início da semana inválida: ${weekStartDateString}`, null);
    }
    if (weekStartDate.getUTCDay() !== 1) {
      Logger.log(`Validation Error: Provided week start date ${weekStartDateString} is not a Monday in UTC (UTC day: ${weekStartDate.getUTCDay()}).`);
      return createJsonResponse(false, `A data de início (${weekStartDateString}) não é uma Segunda-feira válida para o sistema.`, null);
    }
    const weekEndDate = new Date(weekStartDate.getTime());
    weekEndDate.setUTCDate(weekEndDate.getUTCDate() + 6);
    Logger.log(`Filtering instances for Turma "${trimmedTurma}" between UTC ${weekStartDate.toISOString().slice(0, 10)} and ${weekEndDate.toISOString().slice(0, 10)}`);
    const { data: instanceData, header: instanceHeader } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
    const { data: baseData } = getSheetData_(SHEETS.BASE_SCHEDULES);
    const { data: bookingData } = getSheetData_(SHEETS.BOOKING_DETAILS);
    const baseScheduleMap = baseData.reduce((map, row) => {
      const idCol = HEADERS.BASE_SCHEDULES.ID;
      const discCol = HEADERS.BASE_SCHEDULES.DISCIPLINA_PADRAO;
      const profCol = HEADERS.BASE_SCHEDULES.PROFESSOR_PRINCIPAL;
      const reqCols = Math.max(idCol, discCol, profCol) + 1;
      if (row && row.length >= reqCols) {
        const baseId = String(row[idCol] || '').trim();
        if (baseId) {
          const firstProfessor = (String(row[profCol] || '').split(',')[0] || '').trim();
          map[baseId] = {
            disciplina: String(row[discCol] || '').trim(),
            professor: firstProfessor
          };
        }
      }
      return map;
    }, {});
    const bookingDetailsMap = bookingData.reduce((map, row) => {
      const idInstCol = HEADERS.BOOKING_DETAILS.ID_INSTANCIA;
      const discCol = HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL;
      const profRealCol = HEADERS.BOOKING_DETAILS.PROFESSOR_REAL;
      const profOrigCol = HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL;
      const statusCol = HEADERS.BOOKING_DETAILS.STATUS_RESERVA;
      const tipoAulaCol = HEADERS.BOOKING_DETAILS.TIPO_AULA_REPOSICAO; // Novo: tipo de aula
      const reqCols = Math.max(idInstCol, discCol, profRealCol, profOrigCol, statusCol, tipoAulaCol) + 1;
      if (row && row.length >= reqCols) {
        const instanceId = String(row[idInstCol] || '').trim();
        const statusReserva = String(row[statusCol] || '').trim();
        if (instanceId && statusReserva === 'Agendada') {
          map[instanceId] = {
            disciplinaReal: String(row[discCol] || '').trim(),
            professorReal: String(row[profRealCol] || '').trim(),
            professorOriginalBooking: String(row[profOrigCol] || '').trim(),
            tipoAulaReposicao: String(row[tipoAulaCol] || '').trim() // Novo: tipo de aula
          };
        }
      }
      return map;
    }, {});
    const filteredSlots = [];
    const instIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA;
    const baseIdCol = HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO;
    const turmaCol = HEADERS.SCHEDULE_INSTANCES.TURMA;
    const profPrincCol = HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL;
    const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
    const dayCol = HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA;
    const hourCol = HEADERS.SCHEDULE_INSTANCES.HORA_INICIO;
    const typeCol = HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL;
    const statusCol = HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO;
    const absentCol = HEADERS.SCHEDULE_INSTANCES.PROFESSORES_AUSENTES;
    const maxIndexNeeded = Math.max(instIdCol, baseIdCol, turmaCol, profPrincCol, dateCol, dayCol, hourCol, typeCol, statusCol, absentCol);
    if (instanceHeader.length <= maxIndexNeeded) {
      throw new Error(`Planilha "${SHEETS.SCHEDULE_INSTANCES}" não contém todas as colunas necessárias (encontradas ${instanceHeader.length}, esperado pelo menos ${maxIndexNeeded + 1}).`);
    }
    instanceData.forEach((row, index) => {
      if (!row || row.length <= maxIndexNeeded) {
        return;
      }
      const instanceId = String(row[instIdCol] || '').trim();
      const baseId = String(row[baseIdCol] || '').trim();
      const instanceTurma = String(row[turmaCol] || '').trim();
      const instanceUTCDate = formatValueToDate(row[dateCol]);
      const formattedHoraInicio = formatValueToHHMM(row[hourCol], timeZone);
      if (!instanceId || !baseId || !instanceTurma || !instanceUTCDate || !formattedHoraInicio) {
        return;
      }
      if (instanceTurma !== trimmedTurma) return;
      if (instanceUTCDate < weekStartDate || instanceUTCDate > weekEndDate) return;
      const professorPrincipalInstance = String(row[profPrincCol] || '').trim();
      const instanceDiaSemana = String(row[dayCol] || '').trim();
      const originalType = String(row[typeCol] || '').trim();
      const instanceStatus = String(row[statusCol] || '').trim();
      const professoresAusentes = String(row[absentCol] || '').trim();
      let disciplinaParaExibir = '';
      let professorParaExibir = '';
      let professorOriginalNaReserva = '';
      let tipoAulaParaExibir = ''; // Novo: tipo de aula para exibir
      const baseInfo = baseScheduleMap[baseId] || { disciplina: '', professor: '' };
      const bookingDetails = bookingDetailsMap[instanceId];
      if (instanceStatus === STATUS_OCUPACAO.DISPONIVEL) {
        disciplinaParaExibir = baseInfo.disciplina;
        professorParaExibir = (originalType === TIPOS_HORARIO.VAGO) ? '' : professorPrincipalInstance || baseInfo.professor;
      } else if (bookingDetails) {
        disciplinaParaExibir = bookingDetails.disciplinaReal;
        professorParaExibir = bookingDetails.professorReal;
        professorOriginalNaReserva = bookingDetails.professorOriginalBooking;
        if (instanceStatus === STATUS_OCUPACAO.REPOSICAO_AGENDADA) { // Apenas se for uma reposição agendada
             tipoAulaParaExibir = bookingDetails.tipoAulaReposicao; // Novo: pega o tipo de aula
        }
      } else {
        Logger.log(`Warning: Instância ${instanceId} (Status: ${instanceStatus}) sem detalhes de reserva 'Agendada'. Usando dados base/instância.`);
        disciplinaParaExibir = baseInfo.disciplina;
        professorParaExibir = professorPrincipalInstance || baseInfo.professor;
      }
      filteredSlots.push({
        idInstancia: instanceId,
        data: Utilities.formatDate(instanceUTCDate, timeZone, 'dd/MM/yyyy'),
        diaSemana: instanceDiaSemana,
        horaInicio: formattedHoraInicio,
        turma: instanceTurma,
        tipoOriginal: originalType,
        statusOcupacao: instanceStatus,
        disciplinaParaExibir: disciplinaParaExibir,
        professorParaExibir: professorParaExibir,
        professorOriginalNaReserva: professorOriginalNaReserva,
        professorPrincipal: professorPrincipalInstance,
        professoresAusentes: professoresAusentes,
        tipoAulaReposicao: tipoAulaParaExibir // Novo: adiciona ao objeto do slot
      });
    });
    const dailyCounts = { 'Segunda': 0, 'Terça': 0, 'Quarta': 0, 'Quinta': 0, 'Sexta': 0, 'Sábado': 0 };
    filteredSlots.forEach(slot => {
      if (slot.tipoOriginal === TIPOS_HORARIO.FIXO ||
        (slot.tipoOriginal === TIPOS_HORARIO.VAGO && slot.statusOcupacao === STATUS_OCUPACAO.REPOSICAO_AGENDADA)) {
        if (dailyCounts.hasOwnProperty(slot.diaSemana)) {
          dailyCounts[slot.diaSemana]++;
        }
      }
    });
    Logger.log(`Calculated daily counts for display (Turma ${trimmedTurma}, Week ${weekStartDateString}):`, JSON.stringify(dailyCounts));
    Logger.log(`Found ${filteredSlots.length} enriched slots for Turma "${trimmedTurma}" week starting ${weekStartDateString} (UTC).`);
    return createJsonResponse(true, `${filteredSlots.length} horários encontrados.`, { slots: filteredSlots, dailyCounts: dailyCounts });
  } catch (e) {
    Logger.log(`Error in getFilteredScheduleInstances for Turma ${turma}, Week ${weekStartDateString}: ${e.message}\nStack: ${e.stack}`);
    return createJsonResponse(false, `Erro ao buscar horários filtrados: ${e.message}`, null);
  }
}
function getAvailableSlots(tipoReserva) {
  Logger.log(`*** getAvailableSlots called for tipo: ${tipoReserva} ***`);
  try {
    const userEmail = getActiveUserEmail_();
    const userRole = getUserRolePlain_(userEmail);
    if (!userRole) {
      return createJsonResponse(false, 'Usuário não autorizado a buscar horários.', null);
    }
    if (tipoReserva !== TIPOS_RESERVA.REPOSICAO && tipoReserva !== TIPOS_RESERVA.SUBSTITUICAO) {
      return createJsonResponse(false, `Tipo de reserva inválido: ${tipoReserva}`, null);
    }
    const { data: instanceData, header: instanceHeader } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const availableSlots = [];
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const instIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA;
    const baseIdCol = HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO;
    const turmaCol = HEADERS.SCHEDULE_INSTANCES.TURMA;
    const profPrincCol = HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL;
    const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
    const dayCol = HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA;
    const hourCol = HEADERS.SCHEDULE_INSTANCES.HORA_INICIO;
    const typeCol = HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL;
    const statusCol = HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO;
    const maxIndexNeeded = Math.max(instIdCol, baseIdCol, turmaCol, profPrincCol, dateCol, dayCol, hourCol, typeCol, statusCol);
    if (instanceHeader.length <= maxIndexNeeded) {
      throw new Error(`Planilha "${SHEETS.SCHEDULE_INSTANCES}" não contém todas as colunas necessárias para getAvailableSlots.`);
    }
    instanceData.forEach((row, index) => {
      if (!row || row.length <= maxIndexNeeded) return;
      const instanceId = String(row[instIdCol] || '').trim();
      const baseId = String(row[baseIdCol] || '').trim();
      const turma = String(row[turmaCol] || '').trim();
      const professorPrincipal = String(row[profPrincCol] || '').trim();
      const rawInstanceDate = row[dateCol];
      const instanceDiaSemana = String(row[dayCol] || '').trim();
      const formattedHoraInicio = formatValueToHHMM(row[hourCol], timeZone);
      const originalType = String(row[typeCol] || '').trim();
      const instanceStatus = String(row[statusCol] || '').trim();
      if (!instanceId || !baseId || !turma || !rawInstanceDate || !instanceDiaSemana || !formattedHoraInicio || !originalType || !instanceStatus) {
        return;
      }
      let instanceDateForCompare = null;
      if (rawInstanceDate instanceof Date && !isNaN(rawInstanceDate.getTime())) {
        instanceDateForCompare = new Date(rawInstanceDate.getFullYear(), rawInstanceDate.getMonth(), rawInstanceDate.getDate());
        if (instanceDateForCompare < today) return;
      } else {
        return;
      }
      let isMatch = false;
      if (tipoReserva === TIPOS_RESERVA.REPOSICAO) { // Lógica para Reposição/Recuperação
        if (originalType === TIPOS_HORARIO.VAGO && instanceStatus === STATUS_OCUPACAO.DISPONIVEL) {
          if ([USER_ROLES.ADMIN, USER_ROLES.PROFESSOR].includes(userRole)) {
            isMatch = true;
          }
        }
      } else if (tipoReserva === TIPOS_RESERVA.SUBSTITUICAO) {
        if (originalType === TIPOS_HORARIO.FIXO && instanceStatus !== STATUS_OCUPACAO.REPOSICAO_AGENDADA) { // Não pode substituir se já tem reposição/recuperação
          if ([USER_ROLES.ADMIN, USER_ROLES.PROFESSOR].includes(userRole)) {
            if (instanceStatus === STATUS_OCUPACAO.DISPONIVEL || instanceStatus === STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA) {
              isMatch = true;
            }
          }
        }
      }
      if (isMatch) {
        availableSlots.push({
          idInstancia: instanceId,
          baseId: baseId,
          turma: turma,
          professorPrincipal: professorPrincipal,
          data: Utilities.formatDate(instanceDateForCompare, timeZone, 'dd/MM/yyyy'),
          instanceDateObj: instanceDateForCompare,
          diaSemana: instanceDiaSemana,
          horaInicio: formattedHoraInicio,
          tipoOriginal: originalType,
          statusOcupacao: instanceStatus,
        });
      }
    });
    availableSlots.sort((a, b) => {
      const dateComparison = a.instanceDateObj.getTime() - b.instanceDateObj.getTime();
      if (dateComparison !== 0) return dateComparison;
      const timeComparison = a.horaInicio.localeCompare(b.horaInicio);
      if (timeComparison !== 0) return timeComparison;
      return a.turma.localeCompare(b.turma);
    });
    Logger.log(`Found ${availableSlots.length} available slots for type ${tipoReserva}.`);
    return createJsonResponse(true, 'Slots carregados com sucesso.', availableSlots);
  } catch (e) {
    return createJsonResponse(false, `Erro ao buscar horários disponíveis: ${e.message}`, null);
  }
}
function cleanOldScheduleInstances() {
  Logger.log('*** cleanOldScheduleInstances START ***');
  let lock = null;
  try {
    lock = acquireScriptLock_();
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const cleanupDateString = getConfigValue('Data Limite');
    if (!cleanupDateString) throw new Error(`Configuração "Data Limite" não encontrada ou vazia.`);
    const cleanupDateUTC = parseDDMMYYYY(cleanupDateString);
    if (!cleanupDateUTC) throw new Error(`Valor da configuração "Data Limite" inválido: "${cleanupDateString}". Use dd/MM/yyyy.`);
    Logger.log(`Cleaning instances strictly BEFORE UTC date: ${cleanupDateUTC.toISOString().slice(0, 10)}`);
    const { header: instanceHeader, data: instanceData, sheet: instancesSheet } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
    const originalRowCount = instanceData.length;
    if (originalRowCount === 0) {
      Logger.log('No instances found to clean.');
      releaseScriptLock_(lock);
      return;
    }
    const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
    const numCols = instanceHeader.length > 0 ? instanceHeader.length : (Math.max(...Object.values(HEADERS.SCHEDULE_INSTANCES)) + 1);
    if (dateCol >= numCols) throw new Error(`Coluna de Data (índice ${dateCol}) não encontrada na planilha "${SHEETS.SCHEDULE_INSTANCES}". Check HEADERS definition.`);
    const rowsToKeep = [];
    instanceData.forEach((row) => {
      if (row && row.length > dateCol) {
        const instanceUTCDate = formatValueToDate(row[dateCol]);
        if (instanceUTCDate && instanceUTCDate >= cleanupDateUTC) {
          rowsToKeep.push(row);
        }
      }
    });
    const deletedCount = originalRowCount - rowsToKeep.length;
    Logger.log(`Filtering complete: ${rowsToKeep.length} rows to keep, ${deletedCount} rows to delete.`);
    if (deletedCount > 0) {
      Logger.log(`Rewriting sheet "${SHEETS.SCHEDULE_INSTANCES}"...`);
      const dataToWrite = [instanceHeader, ...rowsToKeep].map(row => {
        const paddedRow = [...row];
        while (paddedRow.length < numCols) paddedRow.push('');
        if (paddedRow.length > numCols) return paddedRow.slice(0, numCols);
        return paddedRow;
      });
      instancesSheet.clearContents();
      if (dataToWrite.length > 0) {
        instancesSheet.getRange(1, 1, dataToWrite.length, numCols).setValues(dataToWrite);
      }
      Logger.log(`Sheet rewritten with ${rowsToKeep.length} data rows.`);
      const cache = CacheService.getScriptCache();
      const cacheKey = `SHEET_DATA_${SPREADSHEET_ID}_${SHEETS.SCHEDULE_INSTANCES}`;
      cache.remove(cacheKey);
      Logger.log(`Cache invalidated for sheet "${SHEETS.SCHEDULE_INSTANCES}" after cleanup rewrite.`);
    } else {
      Logger.log('No instances found before the cleanup date. No changes made to the sheet.');
    }
    Logger.log('*** cleanOldScheduleInstances FINISHED ***');
  } catch (e) {
    Logger.log(`ERROR in cleanOldScheduleInstances: ${e.message}\nStack: ${e.stack}`);
  } finally {
    releaseScriptLock_(lock);
  }
}
function cleanupExcessVagoSlots() {
  Logger.log('*** cleanupExcessVagoSlots START (Rewrite Method) ***');
  let lock = null;
  try {
    lock = acquireScriptLock_();
    const threshold = parseInt(getConfigValue('Limite Maximo Aulas Dia Turma')) || 10;
    Logger.log(`Using threshold: ${threshold} (Fixo + Vago Booked)`);
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const { header: instanceHeader, data: instanceData, sheet: instancesSheet } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
    const originalDataRowCount = instanceData.length;
    const numCols = instanceHeader.length > 0 ? instanceHeader.length : (Math.max(...Object.values(HEADERS.SCHEDULE_INSTANCES)) + 1);
    if (originalDataRowCount === 0) {
      Logger.log('No instances found to process.');
      releaseScriptLock_(lock);
      return;
    }
    const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
    const turmaCol = HEADERS.SCHEDULE_INSTANCES.TURMA;
    const typeCol = HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL;
    const statusCol = HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO;
    const maxIndexNeeded = Math.max(dateCol, turmaCol, typeCol, statusCol /* Add other needed cols if any */);
    if (instanceHeader.length <= maxIndexNeeded) {
      throw new Error(`Planilha "${SHEETS.SCHEDULE_INSTANCES}" não contém todas as colunas necessárias (Index: ${maxIndexNeeded}) para cleanupExcessVagoSlots.`);
    }
    const today = new Date();
    today.setHours(0, 0, 0, 0);
    const groupedData = {};
    const rowsToDeleteIndices = new Set();
    Logger.log(`Processing ${originalDataRowCount} instance rows to identify excess Vago slots...`);
    instanceData.forEach((row, index) => {
      if (!row || row.length <= maxIndexNeeded) return;
      const instanceUTCDate = formatValueToDate(row[dateCol]);
      const turma = String(row[turmaCol] || '').trim();
      const originalType = String(row[typeCol] || '').trim();
      const instanceStatus = String(row[statusCol] || '').trim();
      if (!instanceUTCDate || !turma) return;
      let instanceLocalDate = null;
      const rawDate = row[dateCol];
      if (rawDate instanceof Date && !isNaN(rawDate.getTime())) {
        instanceLocalDate = new Date(rawDate.getFullYear(), rawDate.getMonth(), rawDate.getDate());
        if (instanceLocalDate < today) {
          return; // Não processa slots passados
        }
      } else {
        return; // Data inválida
      }
      const dateStringKey = Utilities.formatDate(instanceUTCDate, 'UTC', 'yyyy-MM-dd');
      if (!groupedData[dateStringKey]) groupedData[dateStringKey] = {};
      if (!groupedData[dateStringKey][turma]) {
        groupedData[dateStringKey][turma] = { fixoCount: 0, vagoBookedCount: 0, availableVagoRowIndices: [] };
      }
      if (originalType === TIPOS_HORARIO.FIXO) {
        groupedData[dateStringKey][turma].fixoCount++;
      } else if (originalType === TIPOS_HORARIO.VAGO) {
        if (instanceStatus === STATUS_OCUPACAO.DISPONIVEL) {
          groupedData[dateStringKey][turma].availableVagoRowIndices.push(index + 2); // Armazena o índice da linha da planilha (1-based)
        } else if (instanceStatus === STATUS_OCUPACAO.REPOSICAO_AGENDADA) { // Considera como ocupado
          groupedData[dateStringKey][turma].vagoBookedCount++;
        }
      }
    });
    Logger.log(`Finished grouping data. Identifying rows exceeding threshold...`);
    for (const dateKey in groupedData) {
      for (const turmaName in groupedData[dateKey]) {
        const group = groupedData[dateKey][turmaName];
        const triggerCount = group.fixoCount + group.vagoBookedCount;
        if (triggerCount >= threshold) {
          group.availableVagoRowIndices.forEach(rowIndex => rowsToDeleteIndices.add(rowIndex));
        }
      }
    }
    if (rowsToDeleteIndices.size > 0) {
      Logger.log(`Identified ${rowsToDeleteIndices.size} available Vago slot rows for deletion.`);
      const rowsToKeep = [];
      instanceData.forEach((row, index) => {
        const currentRowIndex = index + 2; // +1 for header, +1 for 0-based to 1-based
        if (!rowsToDeleteIndices.has(currentRowIndex)) {
          rowsToKeep.push(row);
        }
      });
      Logger.log(`Calculated rows to keep: ${rowsToKeep.length} (Original: ${originalDataRowCount})`);
      if (rowsToKeep.length < originalDataRowCount) {
        Logger.log(`Rewriting sheet "${SHEETS.SCHEDULE_INSTANCES}" with filtered data...`);
        const dataToWrite = [instanceHeader, ...rowsToKeep].map(row => {
          const paddedRow = [...row];
          while (paddedRow.length < numCols) paddedRow.push('');
          return paddedRow.slice(0, numCols);
        });
        instancesSheet.clearContents();
        SpreadsheetApp.flush(); // Força a limpeza antes de escrever
        if (dataToWrite.length > 0) {
          instancesSheet.getRange(1, 1, dataToWrite.length, numCols).setValues(dataToWrite);
          Logger.log(`Sheet rewritten successfully with ${rowsToKeep.length} data rows.`);
        } else {
           // Se dataToWrite estiver vazio, mas o cabeçalho existe, reescreva apenas o cabeçalho.
           // Isso pode acontecer se todas as linhas de dados forem removidas.
           // No entanto, se instanceHeader também estiver vazio (improvável com getSheetData_),
           // a planilha realmente ficaria vazia.
           if (instanceHeader.length > 0) {
               instancesSheet.getRange(1, 1, 1, numCols).setValues([instanceHeader.slice(0,numCols)]);
               Logger.log(`Sheet cleared of data rows, header rewritten.`);
           } else {
               Logger.log(`Sheet cleared, and no data (including header) was identified to write back.`);
           }
        }
        const cache = CacheService.getScriptCache();
        const cacheKey = `SHEET_DATA_${SPREADSHEET_ID}_${SHEETS.SCHEDULE_INSTANCES}`;
        cache.remove(cacheKey);
        Logger.log(`Cache invalidated for sheet "${SHEETS.SCHEDULE_INSTANCES}" after Vago slot cleanup rewrite.`);
      } else {
        Logger.log('No changes needed. Number of rows to keep matches original count.');
      }
    } else {
      Logger.log('No available Vago slots needed removal based on threshold.');
    }
    Logger.log('*** cleanupExcessVagoSlots FINISHED (Rewrite Method) ***');
  } catch (e) {
    Logger.log(`ERROR in cleanupExcessVagoSlots (Rewrite Method): ${e.message}\nStack: ${e.stack}`);
  } finally {
    releaseScriptLock_(lock);
  }
}
function getPublicScheduleInstances(weekStartDateString) {
  Logger.log(`*** getPublicScheduleInstances (Public View - Fixed/Booked Only) called for Semana: ${weekStartDateString} ***`);
  try {
    if (!weekStartDateString || typeof weekStartDateString !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(weekStartDateString)) {
      return createJsonResponse(false, 'Semana de início inválida ou formato incorreto (esperado YYYY-MM-DD).', null);
    }
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const parts = weekStartDateString.split('-');
    const weekStartDate = new Date(Date.UTC(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10)));
    if (isNaN(weekStartDate.getTime()) || weekStartDate.getUTCDay() !== 1) { // Checa se é segunda-feira em UTC
      return createJsonResponse(false, `A data de início (${weekStartDateString}) não é uma Segunda-feira válida para o sistema.`, null);
    }
    const weekEndDate = new Date(weekStartDate.getTime());
    weekEndDate.setUTCDate(weekEndDate.getUTCDate() + 6); // Vai até o final do domingo (UTC)
    Logger.log(`Filtering Public (Fixed/Booked) instances between UTC ${weekStartDate.toISOString().slice(0, 10)} and ${weekEndDate.toISOString().slice(0, 10)}`);
    const { data: instanceData } = getSheetData_(SHEETS.SCHEDULE_INSTANCES, HEADERS.SCHEDULE_INSTANCES);
    const { data: baseData } = getSheetData_(SHEETS.BASE_SCHEDULES);
    const { data: bookingData } = getSheetData_(SHEETS.BOOKING_DETAILS);
    const baseScheduleMap = baseData.reduce((map, row) => {
      const idCol = HEADERS.BASE_SCHEDULES.ID;
      const discCol = HEADERS.BASE_SCHEDULES.DISCIPLINA_PADRAO;
      const profCol = HEADERS.BASE_SCHEDULES.PROFESSOR_PRINCIPAL;
      const reqCols = Math.max(idCol, discCol, profCol) + 1;
      if (row && row.length >= reqCols) {
        const baseId = String(row[idCol] || '').trim();
        if (baseId) {
          map[baseId] = {
            disciplina: String(row[discCol] || '').trim(),
            professor: String(row[profCol] || '').trim() // Pode ter múltiplos, mas para exibição pública, o primeiro é suficiente se não houver substituto.
          };
        }
      } return map;
    }, {});
    const bookingDetailsMap = bookingData.reduce((map, row) => {
      const idInstCol = HEADERS.BOOKING_DETAILS.ID_INSTANCIA;
      const discCol = HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL;
      const profRealCol = HEADERS.BOOKING_DETAILS.PROFESSOR_REAL;
      const profOrigCol = HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL;
      const statusCol = HEADERS.BOOKING_DETAILS.STATUS_RESERVA;
      const tipoAulaCol = HEADERS.BOOKING_DETAILS.TIPO_AULA_REPOSICAO; // Novo: tipo de aula
      const reqCols = Math.max(idInstCol, discCol, profRealCol, profOrigCol, statusCol, tipoAulaCol) + 1;
      if (row && row.length >= reqCols) {
        const instanceId = String(row[idInstCol] || '').trim();
        const statusReserva = String(row[statusCol] || '').trim();
        if (instanceId && statusReserva === 'Agendada') map[instanceId] = { disciplinaReal: String(row[discCol] || '').trim(), professorReal: String(row[profRealCol] || '').trim(), professorOriginalBooking: String(row[profOrigCol] || '').trim(), tipoAulaReposicao: String(row[tipoAulaCol] || '').trim() };
      } return map;
    }, {});
    const slotsByTurma = {};
    const instIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA;
    const baseIdCol = HEADERS.SCHEDULE_INSTANCES.ID_BASE_HORARIO;
    const turmaCol = HEADERS.SCHEDULE_INSTANCES.TURMA;
    const profPrincCol = HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL;
    const dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
    const dayCol = HEADERS.SCHEDULE_INSTANCES.DIA_SEMANA;
    const hourCol = HEADERS.SCHEDULE_INSTANCES.HORA_INICIO;
    const typeCol = HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL;
    const statusCol = HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO;
    const absentCol = HEADERS.SCHEDULE_INSTANCES.PROFESSORES_AUSENTES;
    const maxIndexNeeded = Math.max(instIdCol, baseIdCol, turmaCol, profPrincCol, dateCol, dayCol, hourCol, typeCol, statusCol, absentCol);
    instanceData.forEach((row, index) => {
      if (!row || row.length <= maxIndexNeeded) return;
      const instanceId = String(row[instIdCol] || '').trim();
      const baseId = String(row[baseIdCol] || '').trim();
      const instanceTurma = String(row[turmaCol] || '').trim();
      const instanceUTCDate = formatValueToDate(row[dateCol]);
      const formattedHoraInicio = formatValueToHHMM(row[hourCol], timeZone);
      if (!instanceId || !baseId || !instanceTurma || !instanceUTCDate || !formattedHoraInicio) return;
      if (instanceUTCDate < weekStartDate || instanceUTCDate > weekEndDate) return;
      const originalType = String(row[typeCol] || '').trim();
      const instanceStatus = String(row[statusCol] || '').trim();
      const professoresAusentes = String(row[absentCol] || '').trim();
      let includeSlot = false;
      if (originalType === TIPOS_HORARIO.FIXO) { // Sempre inclui horários FIXOS (Disponível ou com Substituição)
        includeSlot = true;
      } else if (originalType === TIPOS_HORARIO.VAGO && instanceStatus === STATUS_OCUPACAO.REPOSICAO_AGENDADA) { // Inclui VAGO apenas se tiver Reposição/Recuperação
        includeSlot = true;
      }
      if (!includeSlot) return;
      const professorPrincipalInstancia = String(row[profPrincCol] || '').trim();
      const instanceDiaSemana = String(row[dayCol] || '').trim();
      let disciplinaParaExibir = '';
      let professorParaExibir = '';
      let professorOriginalNaReserva = '';
      let tipoAulaParaExibir = ''; // Novo: tipo de aula para exibir

      const baseInfo = baseScheduleMap[baseId] || { disciplina: '', professor: '' };
      const bookingDetails = bookingDetailsMap[instanceId];
      if (instanceStatus === STATUS_OCUPACAO.DISPONIVEL && originalType === TIPOS_HORARIO.FIXO) {
        disciplinaParaExibir = baseInfo.disciplina;
        professorParaExibir = professorPrincipalInstancia || baseInfo.professor;
      } else if (bookingDetails) { // Se há uma reserva (Substituição ou Reposição/Recuperação)
        disciplinaParaExibir = bookingDetails.disciplinaReal;
        professorParaExibir = bookingDetails.professorReal;
        professorOriginalNaReserva = bookingDetails.professorOriginalBooking;
        if (instanceStatus === STATUS_OCUPACAO.REPOSICAO_AGENDADA) {
            tipoAulaParaExibir = bookingDetails.tipoAulaReposicao;
        }
      } else {
        // Caso raro: slot FIXO que não está DISPONIVEL e não tem bookingDetails correspondente (pode ser inconsistência ou estado transitório)
        // Para a visão pública, mostrar dados base.
        Logger.log(`Warning (Public View): Instância ${instanceId} (Tipo: ${originalType}, Status: ${instanceStatus}) sem detalhes de reserva 'Agendada'. Usando dados base/instância.`);
        disciplinaParaExibir = baseInfo.disciplina;
        professorParaExibir = professorPrincipalInstancia || baseInfo.professor || 'Prof. N/D';
      }
      const slotData = {
        data: Utilities.formatDate(instanceUTCDate, timeZone, 'dd/MM/yyyy'),
        diaSemana: instanceDiaSemana,
        horaInicio: formattedHoraInicio,
        tipoOriginal: originalType,
        statusOcupacao: instanceStatus, // Para fins de estilização/lógica no frontend, se necessário
        disciplinaParaExibir: disciplinaParaExibir,
        professorParaExibir: professorParaExibir,
        professorOriginalNaReserva: professorOriginalNaReserva, // Para substituições
        professorPrincipal: professorPrincipalInstancia, // Professor(es) originais do horário fixo
        professoresAusentes: professoresAusentes,
        tipoAulaReposicao: tipoAulaParaExibir // Novo: tipo de aula de reposição/recuperação
      };
      if (!slotsByTurma[instanceTurma]) {
        slotsByTurma[instanceTurma] = [];
      }
      slotsByTurma[instanceTurma].push(slotData);
    });
    const turmaCount = Object.keys(slotsByTurma).length;
    Logger.log(`Found relevant public instances for ${turmaCount} turmas for week starting ${weekStartDateString} (UTC).`);
    return createJsonResponse(true, `Horários encontrados para ${turmaCount} turma(s).`, slotsByTurma);
  } catch (e) {
    return createJsonResponse(false, `Erro ao buscar horários públicos: ${e.message}`, null);
  }
}
function getScheduleViewFilters() {
  Logger.log('*** getScheduleViewFilters called ***');
  try {
    getActiveUserEmail_(); // Ensure user is active, even if role isn't strictly checked here
    const turmasResponse = JSON.parse(getTurmasList());
    const turmas = turmasResponse.success ? turmasResponse.data : [];
    if (!turmasResponse.success) {
      Logger.log("Warning: Failed to get turmas list for filters: " + turmasResponse.message);
      // Não retorna erro aqui, pode prosseguir sem turmas se for o caso.
    }
    // Usa a mesma lógica de `calculateGenerationRange_` para definir o início, mas para filtros
    const numWeeks = parseInt(getConfigValue('Semanas Para Gerar Filtros')) || 12;
    // `calculateGenerationRange_` calcula a partir de 2 semanas atrás para `createScheduleInstances`.
    // Para filtros, queremos a partir da segunda-feira da semana atual ou da próxima,
    // dependendo do dia atual.
    const now = new Date();
    const currentDayUTC = now.getUTCDay(); // Domingo = 0, Segunda = 1, ...
    let firstMondayUTC = new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), now.getUTCDate()));
    // Retrocede para a última segunda-feira ou mantém se hoje for segunda
    firstMondayUTC.setUTCDate(firstMondayUTC.getUTCDate() - (currentDayUTC === 0 ? 6 : currentDayUTC - 1));
    const weekStartDates = [];
    Logger.log(`Generating ${numWeeks} week start dates (UTC Mondays) for filters starting from: ${firstMondayUTC.toISOString().slice(0, 10)}`);
    for (let i = 0; i < numWeeks; i++) {
      const weekStartDate = new Date(firstMondayUTC.getTime());
      weekStartDate.setUTCDate(firstMondayUTC.getUTCDate() + (i * 7));
      const valueString = Utilities.formatDate(weekStartDate, 'UTC', 'yyyy-MM-dd');
      weekStartDates.push(valueString);
    }
    Logger.log(`Filters obtained: ${turmas.length} turmas, ${weekStartDates.length} weeks (UTC Mondays).`);
    return createJsonResponse(true, 'Filtros carregados.', { turmas: turmas, weekStartDates: weekStartDates });
  } catch (e) {
    return createJsonResponse(false, `Erro ao obter filtros de horários: ${e.message}`, null);
  }
}