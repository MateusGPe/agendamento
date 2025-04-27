/**
 * Arquivo: Reports.gs
 * Descrição: Funções para gerar relatórios de horários e eventos.
 */
const STATUS_RESERVA = Object.freeze({
    AGENDADA: 'Agendada',
    REALIZADA: 'Realizada',
    CANCELADA: 'Cancelada',
    FALTA_ALUNO: 'Falta Aluno',
    FALTA_PROFESSOR: 'Falta Professor'
});
function onOpen() {
    SpreadsheetApp.getUi()
        .createMenu('Relatórios')
        .addItem('Gerar Relatório de Horários...', 'showReportDialog')
        .addToUi();
}
function showReportDialog() {
    try {
        const ui = SpreadsheetApp.getUi();
        const sheet = SpreadsheetApp.getActiveSheet();
        const startDate = sheet.getRange('A1').getDisplayValue();
        const endDate = sheet.getRange('B1').getDisplayValue();
        const folderId = sheet.getRange('C1').getDisplayValue();
        const titlePrefix = sheet.getRange('D1').getDisplayValue();
        if (!startDate || !endDate) {
            ui.alert("Por favor, insira as datas de início e fim (formato dd/MM/yyyy) nas células A1 e B1.");
            return;
        }
        ui.alert("Gerando relatório... Isso pode demorar um pouco. Por favor, aguarde.");
        const result = generateScheduleReport(startDate, endDate, folderId || null, titlePrefix || null);
        if (result.success) {
            ui.alert(`Relatório gerado!\n${result.message}\nURL: ${result.docUrl || 'N/A'}`);
        } else {
            ui.alert(`Erro ao gerar o relatório:\n${result.message}`);
            Logger.log(`Falha na geração do relatório: ${result.message}`);
        }
    } catch (e) {
        Logger.log(`Erro em showReportDialog: ${e.message}\nStack: ${e.stack}`);
        SpreadsheetApp.getUi().alert("Ocorreu um erro inesperado ao iniciar a geração: " + e.message);
    }
}
const REPORT_DEFAULTS = {
    TITLE_PREFIX: "Relatório de Horários",
    FONT_MAIN: DocumentApp.FontFamily.ARIAL,
    COLOR_DARK_BLUE: '#2C3E50',
    COLOR_DARK_GREY_TEXT: '#555555',
    COLOR_BLACK_TEXT: '#000000',
    COLOR_TABLE_BORDER: '#DDDDDD',
    COLOR_TABLE_HEADER_BG: '#F5F5F5',
    ERROR_MSG_NO_EVENTS_ACTIVE: "Nenhuma ausência, reposição ou substituição ativa encontrada para o período selecionado.",
    ERROR_MSG_NO_EVENTS_CANCELLED: "Nenhuma reserva foi cancelada no período selecionado.",
    HEADER_ACTIVE_EVENTS: "EVENTOS ATIVOS (AUSÊNCIAS / REPOSIÇÕES / SUBSTITUIÇÕES)",
    HEADER_CANCELLED_EVENTS: "RESERVAS CANCELADAS",
    TABLE_HEADERS_ACTIVE: ['Data', 'Hora', 'Disciplina', 'Tipo', 'Detalhes'],
    TABLE_HEADERS_CANCELLED: ['Data', 'Hora', 'Turma', 'Tipo', 'Disciplina', 'Detalhes']
};
function generateScheduleReport(startDateString, endDateString, targetFolderId = null, reportTitlePrefix = null) {
    const effectiveTitlePrefix = reportTitlePrefix || REPORT_DEFAULTS.TITLE_PREFIX;
    Logger.log(`Iniciando generateScheduleReport de ${startDateString} a ${endDateString}. Folder ID: ${targetFolderId}, Título: ${effectiveTitlePrefix}`);
    try {
        invalidateSheetCache_(SHEETS.SCHEDULE_INSTANCES);
        invalidateSheetCache_(SHEETS.BOOKING_DETAILS);
        Logger.log("Caches explicitly invalidated.");
    } catch (e) {
        Logger.log(`Warning: Error invalidating cache: ${e.message}`);
    }
    try {
        const { startDate, endDate, nextDayAfterEnd, formattedStartDate, formattedEndDate, timeZone } = parseAndValidateDates_(startDateString, endDateString);
        Logger.log(`Datas validadas. Período: ${formattedStartDate} a ${formattedEndDate}. TimeZone: ${timeZone}`);
        Logger.log("Buscando e mapeando dados...");
        const { instanceDetailsMap, bookingMap, bookingData } = fetchAndMapData_();
        Logger.log(`Dados mapeados: ${Object.keys(instanceDetailsMap).length} instâncias, ${Object.keys(bookingMap).length} reservas ativas.`);
        Logger.log("Processando eventos ativos...");
        const reportDataActive = processActiveEvents_(instanceDetailsMap, bookingMap, startDate, nextDayAfterEnd, timeZone);
        Logger.log(`Eventos ativos processados: ${Object.keys(reportDataActive).length} professores/grupos.`);
        Logger.log("Processando eventos cancelados...");
        const cancelledEventsList = processCancelledEvents_(bookingData, instanceDetailsMap, startDate, nextDayAfterEnd, timeZone);
        Logger.log(`Eventos cancelados processados: ${cancelledEventsList.length}.`);
        Logger.log("Aplicando estilos e criando doc...");
        const styles = applyStyles_();
        const { doc, body } = createReportDocument_(effectiveTitlePrefix, formattedStartDate, formattedEndDate, timeZone, styles);
        Logger.log(`Documento criado: ${doc.getId()}`);
        Logger.log("Escrevendo cabeçalho...");
        writeReportHeader_(body, effectiveTitlePrefix, formattedStartDate, formattedEndDate, timeZone, styles);
        Logger.log("Escrevendo seção ativos...");
        writeActiveEventsSection_(body, reportDataActive, timeZone, styles);
        Logger.log("Escrevendo seção cancelados...");
        writeCancelledEventsSection_(body, cancelledEventsList, timeZone, styles);
        Logger.log("Salvando/movendo doc...");
        const saveResult = saveAndMoveDoc_(doc, targetFolderId);
        Logger.log(`Resultado final: ${JSON.stringify(saveResult)}`);
        return saveResult;
    } catch (e) {
        Logger.log(`ERRO FATAL em generateScheduleReport: ${e.message}\n${e.stack}`);
        return { success: false, message: `Erro ao gerar relatório: ${e.message}` };
    }
}
function parseAndValidateDates_(startDateString, endDateString) {
    const startDate = parseDDMMYYYY(startDateString);
    const endDate = parseDDMMYYYY(endDateString);
    if (!startDate || !endDate) throw new Error("Datas inválidas. Use dd/MM/yyyy.");
    if (endDate.getTime() < startDate.getTime()) throw new Error("Data final anterior à inicial.");
    const nextDayAfterEnd = new Date(endDate.getTime());
    nextDayAfterEnd.setUTCDate(endDate.getUTCDate() + 1);
    const timeZone = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
    const formattedStartDate = Utilities.formatDate(startDate, 'UTC', 'dd/MM/yyyy');
    const formattedEndDate = Utilities.formatDate(endDate, 'UTC', 'dd/MM/yyyy');
    Logger.log(`Período UTC: ${startDate.toISOString()} a ${nextDayAfterEnd.toISOString()} (exclusive)`);
    return {
        startDate,
        endDate,
        nextDayAfterEnd,
        formattedStartDate,
        formattedEndDate,
        timeZone
    };
}
function fetchAndMapData_() {
    const { data: instanceData } = getSheetData_(SHEETS.SCHEDULE_INSTANCES);
    const { data: bookingData } = getSheetData_(SHEETS.BOOKING_DETAILS);
    const instanceDetailsMap = {};
    const bookingMap = {};
    const i_instanceIdCol = HEADERS.SCHEDULE_INSTANCES.ID_INSTANCIA;
    const i_dateCol = HEADERS.SCHEDULE_INSTANCES.DATA;
    const i_hourCol = HEADERS.SCHEDULE_INSTANCES.HORA_INICIO;
    const i_turmaCol = HEADERS.SCHEDULE_INSTANCES.TURMA;
    const i_profPrincCol = HEADERS.SCHEDULE_INSTANCES.PROFESSOR_PRINCIPAL;
    const i_typeCol = HEADERS.SCHEDULE_INSTANCES.TIPO_ORIGINAL;
    const i_statusCol = HEADERS.SCHEDULE_INSTANCES.STATUS_OCUPACAO;
    const i_absentCol = HEADERS.SCHEDULE_INSTANCES.PROFESSORES_AUSENTES;
    const i_maxIndexNeeded = Math.max(i_instanceIdCol, i_dateCol, i_hourCol, i_turmaCol, i_profPrincCol, i_typeCol, i_statusCol, i_absentCol);
    const problematicInstanceId = "11a465f9-660a-4ef9-bfb8-b1109c5a86db";
    instanceData.forEach((r, rowIndex) => {
        if (!r || r.length <= i_maxIndexNeeded) return;
        const id = String(r[i_instanceIdCol] || '').trim();
        if (id) {
            const rawDateValue = r[i_dateCol];
            const parsedDate = formatValueToDate(rawDateValue);
            if (id === problematicInstanceId) {
                Logger.log(`DEBUG MAP: Instância ${id} (Linha ${rowIndex + 2}) lida. Valor Bruto Data: ${rawDateValue} (Tipo: ${typeof rawDateValue}). Parsed Date: ${parsedDate ? parsedDate.toISOString() : 'null/inválida'}`);
            }
            instanceDetailsMap[id] = {
                date: parsedDate,
                hora: formatValueToHHMM(r[i_hourCol], Session.getScriptTimeZone()) || 'HH:MM?',
                turma: String(r[i_turmaCol] || '').trim(),
                profPrincipal: String(r[i_profPrincCol] || '').trim(),
                tipoOriginal: String(r[i_typeCol] || '').trim(),
                status: String(r[i_statusCol] || '').trim(),
                ausentes: String(r[i_absentCol] || '').trim()
            };
        }
    });
    Logger.log(`Mapeadas ${Object.keys(instanceDetailsMap).length} instâncias.`);
    const b_instanceFkCol = HEADERS.BOOKING_DETAILS.ID_INSTANCIA;
    const b_statusCol = HEADERS.BOOKING_DETAILS.STATUS_RESERVA;
    const b_profRealCol = HEADERS.BOOKING_DETAILS.PROFESSOR_REAL;
    const b_discRealCol = HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL;
    const b_profOrigCol = HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL;
    const b_maxIndexNeeded = Math.max(b_instanceFkCol, b_statusCol, b_profRealCol, b_discRealCol, b_profOrigCol);
    bookingData.forEach((r) => {
        if (!r || r.length <= b_maxIndexNeeded) return;
        const id = String(r[b_instanceFkCol] || '').trim();
        const st = String(r[b_statusCol] || '').trim();
        if (id && st === STATUS_RESERVA.AGENDADA) {
            bookingMap[id] = {
                professorReal: String(r[b_profRealCol] || '').trim(),
                disciplinaReal: String(r[b_discRealCol] || '').trim(),
                professorOriginalBooking: String(r[b_profOrigCol] || '').trim()
            };
        }
    });
    Logger.log(`Mapeadas ${Object.keys(bookingMap).length} reservas 'Agendada'.`);
    return { instanceDetailsMap, bookingMap, bookingData };
}
function processActiveEvents_(instanceDetailsMap, bookingMap, startDate, nextDayAfterEnd, timeZone) {
    const reportDataActive = {};
    for (const instanceId in instanceDetailsMap) {
        const instance = instanceDetailsMap[instanceId];
        if (!instance.date || instance.date < startDate || instance.date >= nextDayAfterEnd) continue;
        const professoresPrincipais = instance.profPrincipal.split(',').map(p => p.trim()).filter(Boolean);
        let eventType = null;
        let details = "";
        let disciplina = "N/D";
        const bookingInfo = bookingMap[instanceId];
        if (instance.status === STATUS_OCUPACAO.REPOSICAO_AGENDADA) {
            eventType = "Reposição";
            if (bookingInfo) {
                details = `Professor: ${bookingInfo.professorReal || 'N/D'}`;
                disciplina = bookingInfo.disciplinaReal || 'N/D';
            } else {
                details = "ERRO: Reserva ativa ('Agendada') não encontrada.";
                Logger.log(`Aviso: Instância ${instanceId} (Reposição Agendada) sem reserva ativa no mapa.`);
            }
        } else if (instance.status === STATUS_OCUPACAO.SUBSTITUICAO_AGENDADA) {
            eventType = "Substituição";
            if (bookingInfo) {
                let originalBase = instance.profPrincipal || 'N/D';
                let original = bookingInfo.professorOriginalBooking || originalBase;
                details = `Substituto: ${bookingInfo.professorReal || 'N/D'} (Prof. Fixo: ${original})`;
                disciplina = bookingInfo.disciplinaReal || 'N/D';
                if (instance.ausentes) {
                    const ausentesList = instance.ausentes.split(',').map(p => p.trim()).filter(Boolean);
                    if (ausentesList.length > 0) {
                        details += ` / Ausente(s): ${ausentesList.join(', ')}`;
                        if (professoresPrincipais.length > 0 && ausentesList.length === professoresPrincipais.length && ausentesList.every(prof => professoresPrincipais.includes(prof))) {
                            details += ` (Todos originais ausentes)`;
                        }
                    }
                }
            } else {
                details = "ERRO: Reserva ativa ('Agendada') não encontrada.";
                Logger.log(`Aviso: Instância ${instanceId} (Substituição Agendada) sem reserva ativa no mapa.`);
            }
        } else if (instance.tipoOriginal === TIPOS_HORARIO.FIXO && instance.ausentes) {
            eventType = "Ausência";
            const ausentesList = instance.ausentes.split(',').map(p => p.trim()).filter(Boolean);
            details = `Ausente(s): ${ausentesList.join(', ')}`;
            if (professoresPrincipais.length > 0) {
                const presentes = professoresPrincipais.filter(p => !ausentesList.includes(p));
                if (presentes.length > 0 && presentes.length < professoresPrincipais.length) {
                    details += ` / Presente(s): ${presentes.join(', ')}`;
                } else if (presentes.length === 0 && ausentesList.length > 0) {
                    details += ` (Todos ausentes)`;
                }
            } else {
                Logger.log(`Aviso: Instância ${instanceId} tipo Fixo sem Professor Principal mas com Ausentes: ${instance.ausentes}`);
            }
        }
        if (eventType) {
            const eventData = {
                date: instance.date,
                formattedDate: Utilities.formatDate(instance.date, 'UTC', 'dd/MM/yyyy'),
                hora: instance.hora,
                disciplina: disciplina,
                eventType: eventType,
                details: details
            };
            let groupKey = "[Sem Professor Fixo Associado]";
            if (professoresPrincipais.length > 0) {
                groupKey = professoresPrincipais.join(', ');
            } else if (eventType === "Reposição" && bookingInfo && bookingInfo.professorReal) {
                groupKey = `[Reposição por: ${bookingInfo.professorReal}]`;
            }
            if (!reportDataActive[groupKey]) reportDataActive[groupKey] = {};
            if (!reportDataActive[groupKey][instance.turma]) reportDataActive[groupKey][instance.turma] = [];
            reportDataActive[groupKey][instance.turma].push(eventData);
        }
    }
    return reportDataActive;
}
function processCancelledEvents_(bookingData, instanceDetailsMap, startDate, nextDayAfterEnd, timeZone) {
    const cancelledEventsList = [];
    const b_instanceFkCol = HEADERS.BOOKING_DETAILS.ID_INSTANCIA;
    const b_statusCol = HEADERS.BOOKING_DETAILS.STATUS_RESERVA;
    const b_profRealCol = HEADERS.BOOKING_DETAILS.PROFESSOR_REAL;
    const b_discRealCol = HEADERS.BOOKING_DETAILS.DISCIPLINA_REAL;
    const b_profOrigCol = HEADERS.BOOKING_DETAILS.PROFESSOR_ORIGINAL;
    const b_typeCol = HEADERS.BOOKING_DETAILS.TIPO_RESERVA;
    const b_maxIndexNeeded = Math.max(b_instanceFkCol, b_statusCol, b_profRealCol, b_discRealCol, b_profOrigCol, b_typeCol);
    Logger.log(`Processando ${bookingData.length} registros de reserva para cancelamentos...`);
    bookingData.forEach((row, index) => {
        if (!row || row.length <= b_maxIndexNeeded) return;
        const status = String(row[b_statusCol] || '').trim();
        if (status !== STATUS_RESERVA.CANCELADA) return;
        const instanceId = String(row[b_instanceFkCol] || '').trim();
        const bookingId = String(row[HEADERS.BOOKING_DETAILS.ID_RESERVA] || '').trim();
        const instance = instanceDetailsMap[instanceId];
        if (!instance) {
            Logger.log(`AVISO (Cancelados): Reserva Cancelada (ID: ${bookingId}, Linha: ${index + 2}) encontrada, mas instância (${instanceId}) NÃO encontrada no mapa.`);
            return;
        }
        if (!instance.date) {
            Logger.log(`AVISO (Cancelados): Instância ${instanceId} (Reserva ${bookingId}) encontrada, mas data inválida.`);
            return;
        }
        if (instance.date >= startDate && instance.date < nextDayAfterEnd) {
            const formattedDateForReport = Utilities.formatDate(instance.date, 'UTC', 'dd/MM/yyyy');
            Logger.log(`DEBUG CANCELLED: Adicionando evento cancelado ao relatório. Instancia ID: ${instanceId}, Data da Instância (UTC): ${instance.date.toISOString()}, Data Formatada p/ Relatório: ${formattedDateForReport}, Hora: ${instance.hora}, Turma: ${instance.turma}`);
            cancelledEventsList.push({
                date: instance.date,
                formattedDate: formattedDateForReport,
                hora: instance.hora,
                turma: instance.turma,
                eventType: `Cancelada (${String(row[b_typeCol] || 'N/D')})`,
                disciplina: String(row[b_discRealCol] || 'N/D'),
                details: `Prof. Subst.: ${String(row[b_profRealCol] || 'N/D')}${String(row[b_profOrigCol] || '').trim() ? ` (Prof. Fixo: ${String(row[b_profOrigCol]).trim()})` : ''}`
            });
        }
    });
    Logger.log(`Processamento de cancelados concluído. ${cancelledEventsList.length} eventos cancelados adicionados ao relatório.`);
    cancelledEventsList.sort((a, b) => {
        const d = a.date.getTime() - b.date.getTime();
        return d !== 0 ? d : a.hora.localeCompare(b.hora);
    });
    return cancelledEventsList;
}
function applyStyles_() {
    const styles = {};
    styles.default = {};
    styles.default[DocumentApp.Attribute.FONT_FAMILY] = REPORT_DEFAULTS.FONT_MAIN;
    styles.default[DocumentApp.Attribute.FONT_SIZE] = 10;
    styles.default[DocumentApp.Attribute.FOREGROUND_COLOR] = REPORT_DEFAULTS.COLOR_BLACK_TEXT;
    styles.default[DocumentApp.Attribute.LINE_SPACING] = 1.15;
    styles.default[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    styles.default[DocumentApp.Attribute.SPACING_AFTER] = 6;
    styles.default[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
    styles.mainTitle = {};
    styles.mainTitle[DocumentApp.Attribute.FONT_FAMILY] = REPORT_DEFAULTS.FONT_MAIN;
    styles.mainTitle[DocumentApp.Attribute.FONT_SIZE] = 18;
    styles.mainTitle[DocumentApp.Attribute.BOLD] = true;
    styles.mainTitle[DocumentApp.Attribute.FOREGROUND_COLOR] = REPORT_DEFAULTS.COLOR_DARK_BLUE;
    styles.mainTitle[DocumentApp.Attribute.SPACING_AFTER] = 4;
    styles.mainTitle[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
    styles.label = {};
    styles.label[DocumentApp.Attribute.FONT_FAMILY] = REPORT_DEFAULTS.FONT_MAIN;
    styles.label[DocumentApp.Attribute.FONT_SIZE] = 9;
    styles.label[DocumentApp.Attribute.BOLD] = true;
    styles.label[DocumentApp.Attribute.FOREGROUND_COLOR] = REPORT_DEFAULTS.COLOR_DARK_GREY_TEXT;
    styles.label[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
    styles.sectionHeaderBar = {};
    styles.sectionHeaderBar[DocumentApp.Attribute.FONT_FAMILY] = REPORT_DEFAULTS.FONT_MAIN;
    styles.sectionHeaderBar[DocumentApp.Attribute.FONT_SIZE] = 11;
    styles.sectionHeaderBar[DocumentApp.Attribute.BOLD] = true;
    styles.sectionHeaderBar[DocumentApp.Attribute.FOREGROUND_COLOR] = REPORT_DEFAULTS.COLOR_BLACK_TEXT;
    styles.sectionHeaderBar[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
    styles.sectionHeaderBar[DocumentApp.Attribute.SPACING_BEFORE] = 18;
    styles.sectionHeaderBar[DocumentApp.Attribute.SPACING_AFTER] = 0;
    styles.sectionHeaderBar[DocumentApp.Attribute.PADDING_TOP] = 6;
    styles.sectionHeaderBar[DocumentApp.Attribute.PADDING_BOTTOM] = 6;
    styles.sectionHeaderBar[DocumentApp.Attribute.PADDING_LEFT] = 8;
    styles.sectionHeaderBar[DocumentApp.Attribute.PADDING_RIGHT] = 8;
    styles.professorTitle = {};
    styles.professorTitle[DocumentApp.Attribute.FONT_FAMILY] = REPORT_DEFAULTS.FONT_MAIN;
    styles.professorTitle[DocumentApp.Attribute.FONT_SIZE] = 11;
    styles.professorTitle[DocumentApp.Attribute.BOLD] = true;
    styles.professorTitle[DocumentApp.Attribute.SPACING_BEFORE] = 12;
    styles.professorTitle[DocumentApp.Attribute.SPACING_AFTER] = 3;
    styles.professorTitle[DocumentApp.Attribute.FOREGROUND_COLOR] = REPORT_DEFAULTS.COLOR_BLACK_TEXT;
    styles.professorTitle[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
    styles.turmaTitle = {};
    styles.turmaTitle[DocumentApp.Attribute.FONT_FAMILY] = REPORT_DEFAULTS.FONT_MAIN;
    styles.turmaTitle[DocumentApp.Attribute.FONT_SIZE] = 10;
    styles.turmaTitle[DocumentApp.Attribute.ITALIC] = true;
    styles.turmaTitle[DocumentApp.Attribute.SPACING_BEFORE] = 0;
    styles.turmaTitle[DocumentApp.Attribute.SPACING_AFTER] = 4;
    styles.turmaTitle[DocumentApp.Attribute.FOREGROUND_COLOR] = REPORT_DEFAULTS.COLOR_DARK_GREY_TEXT;
    styles.turmaTitle[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
    styles.tableHeaderPara = {};
    styles.tableHeaderPara[DocumentApp.Attribute.FONT_FAMILY] = REPORT_DEFAULTS.FONT_MAIN;
    styles.tableHeaderPara[DocumentApp.Attribute.BOLD] = true;
    styles.tableHeaderPara[DocumentApp.Attribute.FOREGROUND_COLOR] = REPORT_DEFAULTS.COLOR_BLACK_TEXT;
    styles.tableHeaderPara[DocumentApp.Attribute.FONT_SIZE] = 9;
    styles.tableHeaderPara[DocumentApp.Attribute.SPACING_BEFORE] = 3;
    styles.tableHeaderPara[DocumentApp.Attribute.SPACING_AFTER] = 3;
    styles.tableCellPara = {};
    styles.tableCellPara[DocumentApp.Attribute.FONT_FAMILY] = REPORT_DEFAULTS.FONT_MAIN;
    styles.tableCellPara[DocumentApp.Attribute.FONT_SIZE] = 9;
    styles.tableCellPara[DocumentApp.Attribute.FOREGROUND_COLOR] = REPORT_DEFAULTS.COLOR_BLACK_TEXT;
    styles.tableCellPara[DocumentApp.Attribute.SPACING_BEFORE] = 3;
    styles.tableCellPara[DocumentApp.Attribute.SPACING_AFTER] = 3;
    styles.tableCellPara[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
    styles.cellBorderInfo = {
        color: REPORT_DEFAULTS.COLOR_TABLE_BORDER
    };
    styles.noEventsMessage = {};
    styles.noEventsMessage[DocumentApp.Attribute.FONT_FAMILY] = REPORT_DEFAULTS.FONT_MAIN;
    styles.noEventsMessage[DocumentApp.Attribute.FONT_SIZE] = 9;
    styles.noEventsMessage[DocumentApp.Attribute.ITALIC] = true;
    styles.noEventsMessage[DocumentApp.Attribute.FOREGROUND_COLOR] = REPORT_DEFAULTS.COLOR_DARK_GREY_TEXT;
    styles.noEventsMessage[DocumentApp.Attribute.SPACING_BEFORE] = 6;
    styles.noEventsMessage[DocumentApp.Attribute.SPACING_AFTER] = 12;
    styles.noEventsMessage[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
    Logger.log("Estilos do relatório definidos.");
    return styles;
}
function createReportDocument_(reportTitlePrefix, formattedStartDate, formattedEndDate, timeZone, styles) {
    let reportMonthYear = "Período";
    try {
        const startDateObj = parseDDMMYYYY(formattedStartDate);
        if (startDateObj) reportMonthYear = Utilities.formatDate(startDateObj, 'UTC', 'MMMM yyyy');
    } catch (e) { }
    const docName = `${reportTitlePrefix} - ${reportMonthYear} (${formattedStartDate} a ${formattedEndDate})`;
    const doc = DocumentApp.create(docName);
    const body = doc.getBody();
    body.setMarginTop(72).setMarginBottom(72).setMarginLeft(72).setMarginRight(72);
    body.setBackgroundColor(null);
    if (body.getNumChildren() > 0) {
        const firstChild = body.getChild(0);
        if (firstChild.getType() == DocumentApp.ElementType.PARAGRAPH) {
            const firstPara = firstChild.asParagraph();
            if (firstPara.getText() === "") {
                let iStyle = {};
                Object.assign(iStyle, styles.default);
                iStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
                iStyle[DocumentApp.Attribute.SPACING_BEFORE] = 0;
                firstPara.setAttributes(iStyle);
                Logger.log("Parágrafo inicial ajustado.");
            }
        }
    }
    Logger.log(`Documento "${docName}" criado.`);
    return {
        doc,
        body
    };
}
function writeReportHeader_(body, reportTitlePrefix, formattedStartDate, formattedEndDate, timeZone, styles) {
    body.appendParagraph(reportTitlePrefix.toUpperCase()).setAttributes(styles.mainTitle).setAlignment(DocumentApp.HorizontalAlignment.LEFT);
    const hTable = body.appendTable([
        ['PERÍODO:', `${formattedStartDate} a ${formattedEndDate}`],
        ['GERADO EM:', Utilities.formatDate(new Date(), timeZone, 'dd/MM/yyyy HH:mm:ss')]
    ]);
    hTable.setBorderWidth(0);
    for (let i = 0; i < hTable.getNumRows(); i++) {
        const lCell = hTable.getCell(i, 0);
        lCell.getChild(0).asParagraph().setAttributes(styles.label);
        lCell.setWidth(80);
        const vCell = hTable.getCell(i, 1);
        let vStyle = {};
        Object.assign(vStyle, styles.default);
        vStyle[DocumentApp.Attribute.FONT_SIZE] = 9;
        vStyle[DocumentApp.Attribute.FOREGROUND_COLOR] = REPORT_DEFAULTS.COLOR_DARK_GREY_TEXT;
        vStyle[DocumentApp.Attribute.SPACING_AFTER] = 0;
        vCell.getChild(0).asParagraph().setAttributes(vStyle);
    }
    body.appendParagraph("").setAttributes(styles.default).setSpacingAfter(18);
}
function writeActiveEventsSection_(body, reportDataActive, timeZone, styles) {
    const headerPara = body.appendParagraph(REPORT_DEFAULTS.HEADER_ACTIVE_EVENTS);
    headerPara.setAttributes(styles.sectionHeaderBar);
    if (Object.keys(reportDataActive).length === 0) {
        body.appendParagraph(REPORT_DEFAULTS.ERROR_MSG_NO_EVENTS_ACTIVE).setAttributes(styles.noEventsMessage);
    } else {
        body.appendParagraph("").setAttributes(styles.default).setSpacingBefore(6).setSpacingAfter(0);
        const sortedGroups = Object.keys(reportDataActive).sort((a, b) => {
            const isASpecial = a.startsWith("[");
            const isBSpecial = b.startsWith("[");
            if (isASpecial && !isBSpecial) return 1;
            if (!isASpecial && isBSpecial) return -1;
            return a.localeCompare(b);
        });
        sortedGroups.forEach(groupKey => {
            let titleText = groupKey;
            if (!groupKey.startsWith("[")) {
                titleText = `Professor(es): ${groupKey}`;
            }
            body.appendParagraph(titleText).setAttributes(styles.professorTitle);
            const turmas = reportDataActive[groupKey];
            const sortedTurmas = Object.keys(turmas).sort((a, b) => a.localeCompare(b));
            sortedTurmas.forEach(turma => {
                body.appendParagraph(`Turma: ${turma}`).setAttributes(styles.turmaTitle);
                const eventos = turmas[turma];
                eventos.sort((a, b) => {
                    const d = a.date.getTime() - b.date.getTime();
                    return d !== 0 ? d : a.hora.localeCompare(b.hora);
                });
                const tableData = eventos.map(e => [e.formattedDate, e.hora, e.disciplina, e.eventType, e.details]);
                appendStyledTable_(body, REPORT_DEFAULTS.TABLE_HEADERS_ACTIVE, tableData, styles);
            });
        });
        body.appendParagraph("").setAttributes(styles.default).setSpacingAfter(18);
    }
}
function writeCancelledEventsSection_(body, cancelledEventsList, timeZone, styles) {
    const headerPara = body.appendParagraph(REPORT_DEFAULTS.HEADER_CANCELLED_EVENTS);
    headerPara.setAttributes(styles.sectionHeaderBar);
    if (cancelledEventsList.length === 0) {
        body.appendParagraph(REPORT_DEFAULTS.ERROR_MSG_NO_EVENTS_CANCELLED).setAttributes(styles.noEventsMessage);
    } else {
        body.appendParagraph("").setAttributes(styles.default).setSpacingBefore(6).setSpacingAfter(0);
        const tableData = cancelledEventsList.map(e => [e.formattedDate, e.hora, e.turma, e.eventType, e.disciplina, e.details]);
        appendStyledTable_(body, REPORT_DEFAULTS.TABLE_HEADERS_CANCELLED, tableData, styles);
        body.appendParagraph("").setAttributes(styles.default).setSpacingAfter(18);
    }
}
function formatTableCell_(cell, text, paragraphStyle, isHeader = false, styles) {
    cell.clear();
    cell.setBackgroundColor(isHeader ? REPORT_DEFAULTS.COLOR_TABLE_HEADER_BG : null);
    const para = cell.appendParagraph(text || "");
    para.setAttributes(paragraphStyle);
    cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    para.setIndentStart(5);
    para.setIndentEnd(5);
}
function appendStyledTable_(body, headers, dataRows, styles) {
    if (!dataRows || dataRows.length === 0) {
        const p = body.appendParagraph("(Nenhum registro para esta tabela)");
        let s = {};
        Object.assign(s, styles.default);
        s[DocumentApp.Attribute.FONT_SIZE] = 9;
        s[DocumentApp.Attribute.ITALIC] = true;
        s[DocumentApp.Attribute.FOREGROUND_COLOR] = REPORT_DEFAULTS.COLOR_DARK_GREY_TEXT;
        s[DocumentApp.Attribute.SPACING_BEFORE] = 4;
        s[DocumentApp.Attribute.SPACING_AFTER] = 10;
        p.setAttributes(s);
        return;
    }
    const table = body.appendTable([headers, ...dataRows]);
    for (let i = 0; i < table.getNumRows(); i++) {
        const row = table.getRow(i);
        for (let j = 0; j < row.getNumCells(); j++) {
            const cell = row.getCell(j);
            const text = cell.getText();
            const isHeader = (i === 0);
            const pStyle = isHeader ? styles.tableHeaderPara : styles.tableCellPara;
            formatTableCell_(cell, text, pStyle, isHeader, styles);
        }
    }
    body.appendParagraph("").setAttributes(styles.default).setSpacingAfter(12);
}
function saveAndMoveDoc_(doc, targetFolderId) {
    let docId, docUrl;
    try {
        doc.saveAndClose();
        docId = doc.getId();
        docUrl = doc.getUrl();
        let message = `Relatório gerado com sucesso!`;
        Logger.log(`Documento salvo: ID ${docId}, URL: ${docUrl}`);
        if (targetFolderId) {
            try {
                const file = DriveApp.getFileById(docId);
                const folder = DriveApp.getFolderById(targetFolderId);
                if (folder) {
                    file.moveTo(folder);
                    Logger.log(`Movido para pasta ID: ${targetFolderId}`);
                    message += ` Movido para a pasta especificada.`;
                } else {
                    Logger.log(`AVISO: Pasta ${targetFolderId} não encontrada. Salvo na raiz.`);
                    message += ` Pasta não encontrada, salvo na raiz.`;
                }
            } catch (moveError) {
                Logger.log(`ERRO ao mover doc ${docId}: ${moveError}`);
                message = `Relatório gerado, mas erro ao mover: ${moveError.message}. Salvo na raiz.`;
                return {
                    success: true,
                    message: message,
                    docId: docId,
                    docUrl: docUrl
                };
            }
        } else {
            message += ` Salvo na raiz do Google Drive.`;
        }
        return {
            success: true,
            message: message,
            docId: docId,
            docUrl: docUrl
        };
    } catch (saveError) {
        Logger.log(`ERRO FATAL ao salvar/fechar: ${saveError}\nStack: ${saveError.stack}`);
        docId = doc ? doc.getId() : 'N/A';
        docUrl = doc ? doc.getUrl() : 'N/A';
        return {
            success: false,
            message: `Erro ao salvar ou finalizar o documento: ${saveError.message}`,
            docId: docId,
            docUrl: docUrl
        };
    }
}