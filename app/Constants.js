/**
 * Arquivo: Constants.gs
 * Descrição: Contém todas as constantes globais usadas na aplicação.
 */
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SCRIPT_LOCK_TIMEOUT_MS = 15000; // Timeout for script lock (15 seconds)
const SHEETS = Object.freeze({
  CONFIG: 'Configuracoes',
  AUTHORIZED_USERS: 'Usuarios Autorizados',
  BASE_SCHEDULES: 'Horarios Base',
  SCHEDULE_INSTANCES: 'Instancias de Horarios',
  BOOKING_DETAILS: 'Reservas Detalhadas',
  DISCIPLINES: 'Disciplinas'
});
const HEADERS = Object.freeze({
  CONFIG: Object.freeze({ NOME: 0, VALOR: 1 }),
  AUTHORIZED_USERS: Object.freeze({ EMAIL: 0, NOME: 1, PAPEL: 2 }),
  BASE_SCHEDULES: Object.freeze({
    ID: 0, TIPO: 1, DIA_SEMANA: 2, HORA_INICIO: 3, DURACAO: 4,
    PROFESSOR_PRINCIPAL: 5, TURMA_PADRAO: 6, DISCIPLINA_PADRAO: 7,
    CAPACIDADE: 8, OBSERVATIONS: 9
  }),
  SCHEDULE_INSTANCES: Object.freeze({
    ID_INSTANCIA: 0, ID_BASE_HORARIO: 1, TURMA: 2, PROFESSOR_PRINCIPAL: 3,
    DATA: 4, DIA_SEMANA: 5, HORA_INICIO: 6, TIPO_ORIGINAL: 7,
    STATUS_OCUPACAO: 8, ID_RESERVA: 9, ID_EVENTO_CALENDAR: 10,
    PROFESSORES_AUSENTES: 11
  }),
  BOOKING_DETAILS: Object.freeze({
    ID_RESERVA: 0, TIPO_RESERVA: 1, ID_INSTANCIA: 2, PROFESSOR_REAL: 3,
    PROFESSOR_ORIGINAL: 4, ALUNOS: 5, TURMAS_AGENDADA: 6, DISCIPLINA_REAL: 7,
    DATA_HORA_INICIO_EFETIVA: 8, STATUS_RESERVA: 9, DATA_CRIACAO: 10,
    CRIADO_POR: 11
  }),
  DISCIPLINES: Object.freeze({ NOME: 0 })
});
const STATUS_OCUPACAO = Object.freeze({
  DISPONIVEL: 'Disponivel',
  REPOSICAO_AGENDADA: 'Reposicao Agendada',
  SUBSTITUICAO_AGENDADA: 'Substituicao Agendada'
});
const TIPOS_RESERVA = Object.freeze({
  REPOSICAO: 'Reposicao',
  SUBSTITUICAO: 'Substituicao'
});
const TIPOS_HORARIO = Object.freeze({
  FIXO: 'Fixo',
  VAGO: 'Vago'
});
const USER_ROLES = Object.freeze({
  ADMIN: 'Admin',
  PROFESSOR: 'Professor'
});
const ADMIN_COPY_EMAILS = ["cae.itq@ifsp.edu.br", "mtm.itq@ifsp.edu.br"]; // Fixed BCC list
const EMAIL_SENDER_NAME = 'Sistema de Reservas IFSP';
