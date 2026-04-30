/* =========================================================
   CRM Reservas - Calendario tipo Google Calendar (frontend)
   - Salones
   - Estados con colores
   - Reglas de bloqueo
   - Avatares por usuario
   - Persistencia MariaDB (via API)
========================================================= */

const API_SYNC_DEBOUNCE_MS = 700;
let persistServerTimer = null;
let persistInFlight = false;
let persistQueued = false;
let serverStateReady = false;
let pendingPersistAfterSync = false;
let saveErrorNotified = false;
let syncInFlight = false;
const ALL_ROOMS_VALUE = "__all_rooms__";
const CURRENT_ORIGIN_STATE_URL = (() => {
  try {
    if (window.location?.origin && /^https?:\/\//i.test(window.location.origin)) {
      return `${window.location.origin}/api/state`;
    }
  } catch (_) { }
  return null;
})();
const API_STATE_CANDIDATES = Array.from(new Set([
  "http://192.168.10.2:3002/api/state",
  "http://localhost:3002/api/state",
  "http://127.0.0.1:3002/api/state",
  ...(CURRENT_ORIGIN_STATE_URL ? [CURRENT_ORIGIN_STATE_URL] : []),
  "/api/state",
]));
let activeApiStateUrl = API_STATE_CANDIDATES[0];

function buildApiUrlFromStateUrl(stateUrl, endpoint) {
  const target = String(endpoint || "").trim().replace(/^\/+/, "");
  if (!target) return stateUrl;
  if (stateUrl.startsWith("http://") || stateUrl.startsWith("https://")) {
    try {
      const u = new URL(stateUrl);
      return `${u.origin}/api/${target}`;
    } catch (_) {
      return `/api/${target}`;
    }
  }
  return `/api/${target}`;
}

const STATUS = {
  CONFIRMADO: "Confirmado",
  LISTA: "Lista de Espera",
  PRERESERVA: "Pre reserva",
  MANTENIMIENTO: "Mantenimiento",
  CANCELADO: "Cancelado",
  PERDIDO: "Perdido",
  RESERVA_SIN_COTIZACION: "Reserva sin Cotizacion",
  PRIMERA: "1er Cotizacion",
  SEGUIMIENTO: "Seguimiento",
};

const STATUS_META = [
  { key: STATUS.RESERVA_SIN_COTIZACION, colorVar: "--c-reserva-sin-cotizacion" },
  { key: STATUS.PRIMERA, colorVar: "--c-primera" },
  { key: STATUS.PERDIDO, colorVar: "--c-perdido" },
  { key: STATUS.SEGuimiento ?? STATUS.SEGUIMIENTO, colorVar: "--c-seguimiento" }, // safety
  { key: STATUS.LISTA, colorVar: "--c-lista" },
  { key: STATUS.PRERESERVA, colorVar: "--c-prereserva" },
  { key: STATUS.CONFIRMADO, colorVar: "--c-confirmado" },
  { key: STATUS.CANCELADO, colorVar: "--c-cancelado" },
  { key: STATUS.MANTENIMIENTO, colorVar: "--c-mantenimiento" },
].map(x => ({ ...x, key: x.key === undefined ? STATUS.SEGUIMIENTO : x.key }));

const AUTO_STATUSES = new Set([STATUS.RESERVA_SIN_COTIZACION, STATUS.PRIMERA, STATUS.SEGUIMIENTO, STATUS.PERDIDO]);
const DASHBOARD_STATUS_ORDER = [
  STATUS.PERDIDO,
  STATUS.SEGUIMIENTO,
  STATUS.PRIMERA,
  STATUS.RESERVA_SIN_COTIZACION,
  STATUS.CANCELADO,
  STATUS.CONFIRMADO,
  STATUS.PRERESERVA,
  STATUS.LISTA,
];
const LEADS_PIPELINE_STATUSES = new Set([
  STATUS.RESERVA_SIN_COTIZACION,
  STATUS.PRIMERA,
  STATUS.SEGUIMIENTO,
  STATUS.LISTA,
  STATUS.PRERESERVA,
  STATUS.PERDIDO,
]);
const LEADS_PIPELINE_STATUS_ORDER = [
  STATUS.RESERVA_SIN_COTIZACION,
  STATUS.PRIMERA,
  STATUS.SEGUIMIENTO,
  STATUS.LISTA,
  STATUS.PRERESERVA,
  STATUS.PERDIDO,
];
const DASHBOARD_EVENT_TYPES = ["Social", "Corporativo", "Individual"];
const USER_ROLE = {
  SELLER: "vendedor",
  RECEPTIONIST: "recepcionista",
  ADMIN: "admin",
};
const REPORTABLE_USER_ROLES = [USER_ROLE.SELLER, USER_ROLE.RECEPTIONIST];
const USER_ROLE_LABELS = {
  [USER_ROLE.SELLER]: "Vendedor",
  [USER_ROLE.RECEPTIONIST]: "Recepcionista",
  [USER_ROLE.ADMIN]: "Administrador",
};
const USER_ROLE_PLURAL_LABELS = {
  [USER_ROLE.SELLER]: "Vendedores",
  [USER_ROLE.RECEPTIONIST]: "Recepcionistas",
  [USER_ROLE.ADMIN]: "Administradores",
};
function normalizeUserRole(value) {
  const raw = String(value || "").trim().toLowerCase();
  if (raw === USER_ROLE.RECEPTIONIST) return USER_ROLE.RECEPTIONIST;
  if (raw === USER_ROLE.ADMIN) return USER_ROLE.ADMIN;
  return USER_ROLE.SELLER;
}
function userRoleLabel(value) {
  const role = normalizeUserRole(value);
  return USER_ROLE_LABELS[role] || "Vendedor";
}
function userRolePluralLabel(value) {
  const role = normalizeUserRole(value);
  return USER_ROLE_PLURAL_LABELS[role] || "Vendedores";
}
function isAutoStatus(status) {
  return AUTO_STATUSES.has(String(status || "").trim());
}

function reservationStatusFromQuotePresence(ev) {
  return ev?.quote ? STATUS.PRIMERA : STATUS.RESERVA_SIN_COTIZACION;
}

function cssVar(name) {
  return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
}
function statusColor(status) {
  const meta = STATUS_META.find(s => s.key === status);
  return meta ? cssVar(meta.colorVar) : "rgba(255,255,255,0.35)";
}

const SALONES_DEFAULT = [];
const USERS_DEFAULT = [];
const COMPANIES_DEFAULT = [];
const SERVICES_DEFAULT = [];

const HOUR_START = 0;   // 00:00
const HOUR_END = 23;  // Ultima hora editable
const HOUR_HEIGHT = 56; // px
const SNAP_MINUTES = 30;
const AUTO_SCROLL_EDGE_PX = 44;
const AUTO_SCROLL_STEP_PX = 26;
const USE_ENHANCED_SELECTS = false;
const SETTINGS_STORAGE_KEY = "crm_topbar_settings_v1";
const QUICK_TEMPLATES_STORAGE_KEY = "crm_quick_templates_v1";
const AUTH_SESSION_STORAGE_KEY = "crm_auth_session_v1";
const ACTIVE_MODULE_SCREEN_STORAGE_KEY = "crm_active_module_screen_v1";
const CORPORATE_TEMPLATE_ID = "tpl-corporativo";
const CORPORATE_TEMPLATE_NAME = "Corporativo";
const SERVIHOSP_TEMPLATE_ID = "tpl-servi-hosp";
const SERVIHOSP_TEMPLATE_NAME = "Servi Hosp";
const CONTRACT_CORP_TEMPLATE_ID = "tpl-contrato-corp";
const CONTRACT_CORP_TEMPLATE_NAME = "Jardines";
const DEFAULT_TEMPLATE_HEADER_IMAGE = "./Encabezadojdl.png";
const DEFAULT_TEMPLATE_FOOTER_IMAGE = "./piedepaginajdl.png";
const SERVIHOSP_TEMPLATE_HEADER_IMAGE = "./EncabezadoServ.jpg";
const TEMPLATE_SIGNATURE_MIN_W_PCT = 10;
const TEMPLATE_SIGNATURE_MIN_H_PCT = 3;
const TEMPLATE_SIGNATURE_MAX_W_PCT = 35;
const TEMPLATE_SIGNATURE_MAX_H_PCT = 12;
const TEMPLATE_SIGNATURE_FALLBACK_W_PCT = 22;
const TEMPLATE_SIGNATURE_FALLBACK_H_PCT = 5;
const TEMPLATE_COORD_BASE_W_PT = 612;
const TEMPLATE_COORD_BASE_H_PT = 792;
const PAST_EVENT_ADMIN_EDIT_CODE = "JDL-ADMIN-2026";
const FALLBACK_MAX_PAX_PER_SLOT = 2000;
const DEFAULT_TOPBAR_SETTINGS = {
  showLegend: true,
  compactEvents: false,
  showWeekends: true,
};

let state = buildInitialState();
let viewStart = startOfWeek(new Date()); // Monday-based
let selectedSalon = ALL_ROOMS_VALUE;
let navMode = "week";
let monthCursor = startOfMonth(new Date());
let pendingCreateDates = null;
let quoteDraft = null;
let quoteSaveInFlight = false;
let quoteSelectedItemIds = new Set();
let quoteAdvanceEditingId = "";
let companyManagersDraft = [];
let editingCompanyId = "";
let editingCompanyManagerId = "";
let editingServiceId = "";
let editingServiceCategoryId = "";
let editingServiceSubcategoryId = "";
let globalGoalEditingKey = "";
let editingSalonName = "";
let moduleModalReturnScreen = "";
let catalogoCategoriasServicio = [];
let catalogoSubcategoriasServicio = [];
let historyTargetEventId = null;
let appointmentTargetEventId = null;
let userModalEditingId = "";
let userMonthlyGoalsDraft = [];
let editingUserGoalMonth = "";
let occupancySelectedDayIso = "";
let authSession = loadPersistedAuthSession();
let userSignatureNormalizedDataUrl = "";
let checklistTemplateDraft = [];
let checklistTemplateEditingId = "";
let checklistTemplateSectionsDraft = [];
let checklistTemplatesDraft = [];
let checklistTemplateCurrentId = "";
let checklistTemplateSectionEditingId = "";
let currentEventChecklistId = "";
let eventChecklistDraft = null;
let quoteServiceTemplateManagerItems = [];
const interaction = {
  selecting: null,
  selectionBox: null,
  dragging: null,
  stretching: null,
  suppressClickUntil: 0,
};
const uiEnhancers = {
  selectChoices: new Map(),
  selectObserver: null,
  selectQueue: new Map(),
  selectQueueTimer: null,
  customTopbarSelects: new Map(),
  openCustomSelect: null,
};
let topbarSettings = loadTopbarSettings();
let quickTemplates = ensureCorporateTemplateSeed(Array.isArray(state.quickTemplates) ? state.quickTemplates : []);
const signatureImageAnalysisCache = new Map();
const pastEventEditAuthorizedKeys = new Set();
const notifiedReminderKeys = new Set();
let menuMontajeSelectedKey = "";
let menuMontajeSelectedVersion = 0;
let menuMontajeSelectableSelectedKey = "";
let menuMontajeSelectableSelectedVersion = 0;
let menuMontajeSelectableSilentUpdate = false;
let quoteItemsExpanded = false;
let mmsShowAllGuarniciones = false;
let mmsShowAllPostres = false;
let mmsPrimaryMode = "menu";
let mmsCurrentStage = "plato";
let mmsPlatoQty = 1;
let mmsSelectedPlatoItems = [];
let mmsLineItemsDraft = [];
let mmsActiveLineKey = "";
let mmsSelectedSalsaIds = [];
let mmsSelectedBebidaIds = [];
let mmsGuarnicionQtyById = {};
let mmsPostreQtyById = {};
let mmsBebidaQtyById = {};
const mmsNameMapMemo = new WeakMap();
let dashboardHoverTipEl = null;
let menuMontajeSelectableCatalogCache = {
  proteins: [],
  preparationsByProtein: new Map(),
  suggestionsByPair: new Map(),
  suggestionsInFlightByPair: new Map(),
  salsas: [],
  guarniciones: [],
  postres: [],
  bebidas: [],
  comentarios: [],
  montajeTipos: [],
  montajeAdicionales: [],
};
let menuCatalogManagerKind = "plato_fuerte";
let menuCatalogManagerEditingId = "";
let menuCatalogManagerRows = [];
let menuSuggestionDraggingRow = null;

const el = {
  appShell: document.getElementById("appShell"),
  loginScreen: document.getElementById("loginScreen"),
  loginForm: document.getElementById("loginForm"),
  loginUserSelect: document.getElementById("loginUserSelect"),
  loginPassword: document.getElementById("loginPassword"),
  loginAvatar: document.getElementById("loginAvatar"),
  loginError: document.getElementById("loginError"),
  moduleHubScreen: document.getElementById("moduleHubScreen"),
  reportsHubScreen: document.getElementById("reportsHubScreen"),
  settingsScreen: document.getElementById("settingsScreen"),
  btnSideDashboardReports: document.getElementById("btnSideDashboardReports"),
  btnSideCalendar: document.getElementById("btnSideCalendar"),
  btnSideReports: document.getElementById("btnSideReports"),
  btnSideSettings: document.getElementById("btnSideSettings"),
  btnSupportMenu: document.getElementById("btnSupportMenu"),
  supportOptions: document.getElementById("supportOptions"),
  btnModuleCalendar: document.getElementById("btnModuleCalendar"),
  btnModuleReports: document.getElementById("btnModuleReports"),
  btnModuleSettings: document.getElementById("btnModuleSettings"),
  btnBackFromReports: document.getElementById("btnBackFromReports"),
  btnBackFromSettings: document.getElementById("btnBackFromSettings"),
  topbarWelcome: document.getElementById("topbarWelcome"),
  topbarUserAvatar: document.getElementById("topbarUserAvatar"),
  btnLogout: document.getElementById("btnLogout"),
  weekLabel: document.getElementById("weekLabel"),
  topbarReminderWrap: document.getElementById("topbarReminderWrap"),
  btnTopbarReminders: document.getElementById("btnTopbarReminders"),
  topbarReminderCount: document.getElementById("topbarReminderCount"),
  topbarReminderPanel: document.getElementById("topbarReminderPanel"),
  topbarReminderSubtitle: document.getElementById("topbarReminderSubtitle"),
  topbarReminderList: document.getElementById("topbarReminderList"),
  btnPrev: document.getElementById("btnPrev"),
  btnNext: document.getElementById("btnNext"),
  btnToday: document.getElementById("btnToday"),
  navMode: document.getElementById("navMode"),
  btnOpenLeads: document.getElementById("btnOpenLeads"),
  btnFindEvent: document.getElementById("btnFindEvent"),
  btnNew: document.getElementById("btnNew"),
  settingsMenu: document.getElementById("settingsMenu"),
  btnSettings: document.getElementById("btnSettings"),
  settingsPanel: document.getElementById("settingsPanel"),
  settingShowLegend: document.getElementById("settingShowLegend"),
  settingCompactEvents: document.getElementById("settingCompactEvents"),
  settingShowWeekends: document.getElementById("settingShowWeekends"),
  btnToggleQuickAdd: document.getElementById("btnToggleQuickAdd"),
  quickAddGroup: document.getElementById("quickAddGroup"),
  btnToggleReports: document.getElementById("btnToggleReports"),
  reportsGroup: document.getElementById("reportsGroup"),
  btnQuickAddInstitution: document.getElementById("btnQuickAddInstitution"),
  btnQuickAddManager: document.getElementById("btnQuickAddManager"),
  btnQuickAddUser: document.getElementById("btnQuickAddUser"),
  btnQuickAddService: document.getElementById("btnQuickAddService"),
  btnQuickAddSalon: document.getElementById("btnQuickAddSalon"),
  btnQuickAddGlobalGoal: document.getElementById("btnQuickAddGlobalGoal"),
  btnQuickAddChecklist: document.getElementById("btnQuickAddChecklist"),
  btnQuickOpenMenuCatalog: document.getElementById("btnQuickOpenMenuCatalog"),
  btnManageQuoteServiceTemplates: document.getElementById("btnManageQuoteServiceTemplates"),
  quoteServiceTemplateManagerBackdrop: document.getElementById("quoteServiceTemplateManagerBackdrop"),
  btnQuoteServiceTemplateManagerClose: document.getElementById("btnQuoteServiceTemplateManagerClose"),
  settingsQuoteServiceTemplateSelect: document.getElementById("settingsQuoteServiceTemplateSelect"),
  settingsQuoteServiceTemplateName: document.getElementById("settingsQuoteServiceTemplateName"),
  btnSettingsQuoteServiceTemplateNew: document.getElementById("btnSettingsQuoteServiceTemplateNew"),
  btnSettingsQuoteServiceTemplateSave: document.getElementById("btnSettingsQuoteServiceTemplateSave"),
  btnSettingsQuoteServiceTemplateUpdate: document.getElementById("btnSettingsQuoteServiceTemplateUpdate"),
  btnSettingsQuoteServiceTemplateDelete: document.getElementById("btnSettingsQuoteServiceTemplateDelete"),
  settingsQuoteTemplateServiceSearch: document.getElementById("settingsQuoteTemplateServiceSearch"),
  settingsQuoteTemplateServicesList: document.getElementById("settingsQuoteTemplateServicesList"),
  settingsQuoteTemplateServiceQty: document.getElementById("settingsQuoteTemplateServiceQty"),
  btnSettingsQuoteTemplateServiceAdd: document.getElementById("btnSettingsQuoteTemplateServiceAdd"),
  settingsQuoteTemplateItemsBody: document.getElementById("settingsQuoteTemplateItemsBody"),
  btnReportSales: document.getElementById("btnReportSales"),
  btnReportAccounting: document.getElementById("btnReportAccounting"),
  btnReportOccupancy: document.getElementById("btnReportOccupancy"),
  btnReportDashboard: document.getElementById("btnReportDashboard"),
  btnReportInstitution: document.getElementById("btnReportInstitution"),
  salesReportBackdrop: document.getElementById("salesReportBackdrop"),
  btnSalesReportClose: document.getElementById("btnSalesReportClose"),
  salesReportSearch: document.getElementById("salesReportSearch"),
  salesReportFrom: document.getElementById("salesReportFrom"),
  salesReportTo: document.getElementById("salesReportTo"),
  salesReportUser: document.getElementById("salesReportUser"),
  salesReportStatus: document.getElementById("salesReportStatus"),
  salesReportSalon: document.getElementById("salesReportSalon"),
  salesReportCompany: document.getElementById("salesReportCompany"),
  btnSalesReportReset: document.getElementById("btnSalesReportReset"),
  btnSalesReportExportExcel: document.getElementById("btnSalesReportExportExcel"),
  salesReportFiltersMeta: document.getElementById("salesReportFiltersMeta"),
  salesReportSummary: document.getElementById("salesReportSummary"),
  salesReportBody: document.getElementById("salesReportBody"),
  accountingReportBackdrop: document.getElementById("accountingReportBackdrop"),
  btnAccountingReportClose: document.getElementById("btnAccountingReportClose"),
  accountingReportSearch: document.getElementById("accountingReportSearch"),
  accountingReportFrom: document.getElementById("accountingReportFrom"),
  accountingReportTo: document.getElementById("accountingReportTo"),
  accountingReportUser: document.getElementById("accountingReportUser"),
  accountingReportStatus: document.getElementById("accountingReportStatus"),
  accountingReportSalon: document.getElementById("accountingReportSalon"),
  accountingReportCompany: document.getElementById("accountingReportCompany"),
  btnAccountingReportReset: document.getElementById("btnAccountingReportReset"),
  btnAccountingReportExportExcel: document.getElementById("btnAccountingReportExportExcel"),
  accountingReportFiltersMeta: document.getElementById("accountingReportFiltersMeta"),
  accountingReportSummary: document.getElementById("accountingReportSummary"),
  accountingReportBody: document.getElementById("accountingReportBody"),
  accountStatementBackdrop: document.getElementById("accountStatementBackdrop"),
  btnAccountStatementClose: document.getElementById("btnAccountStatementClose"),
  accountStatementCompany: document.getElementById("accountStatementCompany"),
  accountStatementNit: document.getElementById("accountStatementNit"),
  accountStatementManager: document.getElementById("accountStatementManager"),
  accountStatementPhone: document.getElementById("accountStatementPhone"),
  accountStatementAddress: document.getElementById("accountStatementAddress"),
  accountStatementIssueDate: document.getElementById("accountStatementIssueDate"),
  accountStatementIssueTime: document.getElementById("accountStatementIssueTime"),
  accountStatementMeta: document.getElementById("accountStatementMeta"),
  accountStatementSummary: document.getElementById("accountStatementSummary"),
  accountStatementBody: document.getElementById("accountStatementBody"),
  occupancyReportBackdrop: document.getElementById("occupancyReportBackdrop"),
  btnOccupancyReportClose: document.getElementById("btnOccupancyReportClose"),
  occupancyReportSubtitle: document.getElementById("occupancyReportSubtitle"),
  occupancyReportWeek: document.getElementById("occupancyReportWeek"),
  btnOccupancyReportPrevWeek: document.getElementById("btnOccupancyReportPrevWeek"),
  btnOccupancyReportNextWeek: document.getElementById("btnOccupancyReportNextWeek"),
  btnOccupancyReportTodayWeek: document.getElementById("btnOccupancyReportTodayWeek"),
  btnOccupancyReportExportExcel: document.getElementById("btnOccupancyReportExportExcel"),
  occupancyReportSummary: document.getElementById("occupancyReportSummary"),
  occupancyDaysStrip: document.getElementById("occupancyDaysStrip"),
  occupancyDayDetail: document.getElementById("occupancyDayDetail"),
  occupancyReportBody: document.getElementById("occupancyReportBody"),
  dashboardReportBackdrop: document.getElementById("dashboardReportBackdrop"),
  btnDashboardReportClose: document.getElementById("btnDashboardReportClose"),
  dashboardReportTitle: document.getElementById("dashboardReportTitle"),
  dashboardReportSubtitle: document.getElementById("dashboardReportSubtitle"),
  dashboardReportPeriod: document.getElementById("dashboardReportPeriod"),
  dashboardReportMonth: document.getElementById("dashboardReportMonth"),
  dashboardReportWeekField: document.getElementById("dashboardReportWeekField"),
  dashboardReportWeek: document.getElementById("dashboardReportWeek"),
  dashboardReportFrom: document.getElementById("dashboardReportFrom"),
  dashboardReportTo: document.getElementById("dashboardReportTo"),
  dashboardReportRole: document.getElementById("dashboardReportRole"),
  dashboardReportScope: document.getElementById("dashboardReportScope"),
  dashboardReportSeller: document.getElementById("dashboardReportSeller"),
  btnDashboardReportCurrentMonth: document.getElementById("btnDashboardReportCurrentMonth"),
  btnDashboardReportReset: document.getElementById("btnDashboardReportReset"),
  dashboardGoalsGrid: document.getElementById("dashboardGoalsGrid"),
  dashboardCompareTitle: document.getElementById("dashboardCompareTitle"),
  dashboardCompareSubtitle: document.getElementById("dashboardCompareSubtitle"),
  dashboardCompareChart: document.getElementById("dashboardCompareChart"),
  dashboardBestTitle: document.getElementById("dashboardBestTitle"),
  dashboardBestSubtitle: document.getElementById("dashboardBestSubtitle"),
  dashboardBestMonthChart: document.getElementById("dashboardBestMonthChart"),
  dashboardSellerList: document.getElementById("dashboardSellerList"),
  institutionReportBackdrop: document.getElementById("institutionReportBackdrop"),
  btnInstitutionReportClose: document.getElementById("btnInstitutionReportClose"),
  institutionReportCompanySearch: document.getElementById("institutionReportCompanySearch"),
  institutionReportCompanyResults: document.getElementById("institutionReportCompanyResults"),
  institutionReportCompany: document.getElementById("institutionReportCompany"),
  institutionReportFrom: document.getElementById("institutionReportFrom"),
  institutionReportTo: document.getElementById("institutionReportTo"),
  btnInstitutionReportCurrentYear: document.getElementById("btnInstitutionReportCurrentYear"),
  btnInstitutionReportReset: document.getElementById("btnInstitutionReportReset"),
  institutionReportHeadline: document.getElementById("institutionReportHeadline"),
  institutionReportSummary: document.getElementById("institutionReportSummary"),
  institutionReportNav: document.getElementById("institutionReportNav"),
  institutionReportContent: document.getElementById("institutionReportContent"),
  institutionOverviewGrid: document.getElementById("institutionOverviewGrid"),
  institutionReportChartsBody: document.getElementById("institutionReportChartsBody"),
  institutionReportSalonBody: document.getElementById("institutionReportSalonBody"),
  institutionReportDishBody: document.getElementById("institutionReportDishBody"),
  institutionReportManagerBody: document.getElementById("institutionReportManagerBody"),
  institutionReportTimelineBody: document.getElementById("institutionReportTimelineBody"),
  institutionReportEventsBody: document.getElementById("institutionReportEventsBody"),
  checklistTemplateBackdrop: document.getElementById("checklistTemplateBackdrop"),
  btnChecklistTemplateClose: document.getElementById("btnChecklistTemplateClose"),
  checklistTemplateSelect: document.getElementById("checklistTemplateSelect"),
  checklistTemplateName: document.getElementById("checklistTemplateName"),
  checklistTemplateActive: document.getElementById("checklistTemplateActive"),
  checklistTemplateInput: document.getElementById("checklistTemplateInput"),
  checklistTemplateSectionSelect: document.getElementById("checklistTemplateSectionSelect"),
  checklistTemplateSectionEditSelect: document.getElementById("checklistTemplateSectionEditSelect"),
  checklistTemplateSectionInput: document.getElementById("checklistTemplateSectionInput"),
  btnChecklistTemplateAdd: document.getElementById("btnChecklistTemplateAdd"),
  btnChecklistTemplateAddSection: document.getElementById("btnChecklistTemplateAddSection"),
  btnChecklistTemplateResetSection: document.getElementById("btnChecklistTemplateResetSection"),
  btnChecklistTemplateNew: document.getElementById("btnChecklistTemplateNew"),
  btnChecklistTemplateDisable: document.getElementById("btnChecklistTemplateDisable"),
  checklistTemplateSectionsBody: document.getElementById("checklistTemplateSectionsBody"),
  checklistTemplateBody: document.getElementById("checklistTemplateBody"),
  eventChecklistBackdrop: document.getElementById("eventChecklistBackdrop"),
  btnEventChecklistClose: document.getElementById("btnEventChecklistClose"),
  btnEventChecklistDiscard: document.getElementById("btnEventChecklistDiscard"),
  btnEventChecklistSave: document.getElementById("btnEventChecklistSave"),
  eventChecklistTemplateSelect: document.getElementById("eventChecklistTemplateSelect"),
  eventChecklistSubtitle: document.getElementById("eventChecklistSubtitle"),
  eventChecklistDate: document.getElementById("eventChecklistDate"),
  eventChecklistEventName: document.getElementById("eventChecklistEventName"),
  eventChecklistNotes: document.getElementById("eventChecklistNotes"),
  eventChecklistProgressLabel: document.getElementById("eventChecklistProgressLabel"),
  eventChecklistSatisfactionLabel: document.getElementById("eventChecklistSatisfactionLabel"),
  eventChecklistProgressFill: document.getElementById("eventChecklistProgressFill"),
  eventChecklistBody: document.getElementById("eventChecklistBody"),
  legend: document.getElementById("legend"),
  timeCol: document.getElementById("timeCol"),
  daysHeader: document.getElementById("daysHeader"),
  grid: document.getElementById("grid"),
  toast: document.getElementById("toast"),
  roomSelect: document.getElementById("roomSelect"),

  modalBackdrop: document.getElementById("modalBackdrop"),
  btnClose: document.getElementById("btnClose"),
  btnBackToForm: document.getElementById("btnBackToForm"),
  btnDiscard: document.getElementById("btnDiscard"),
  eventForm: document.getElementById("eventForm"),
  eventId: document.getElementById("eventId"),
  eventName: document.getElementById("eventName"),
  eventDate: document.getElementById("eventDate"),
  eventDateEnd: document.getElementById("eventDateEnd"),
  eventStatus: document.getElementById("eventStatus"),
  statusHint: document.getElementById("statusHint"),
  startTime: document.getElementById("startTime"),
  endTime: document.getElementById("endTime"),
  btnAddSlot: document.getElementById("btnAddSlot"),
  slotsBody: document.getElementById("slotsBody"),
  eventUser: document.getElementById("eventUser"),
  eventPax: document.getElementById("eventPax"),
  eventNotes: document.getElementById("eventNotes"),
  modalTitle: document.getElementById("modalTitle"),
  modalSubtitle: document.getElementById("modalSubtitle"),
  btnDelete: document.getElementById("btnDelete"),
  btnCancelEvent: document.getElementById("btnCancelEvent"),
  btnQuoteEvent: document.getElementById("btnQuoteEvent"),
  btnSetMaintenance: document.getElementById("btnSetMaintenance"),
  btnToggleHistory: document.getElementById("btnToggleHistory"),
  btnToggleAppointments: document.getElementById("btnToggleAppointments"),
  btnAddAppointment: document.getElementById("btnAddAppointment"),
  historyPanel: document.getElementById("historyPanel"),
  historyBody: document.getElementById("historyBody"),
  appointmentPanel: document.getElementById("appointmentPanel"),
  appointmentBody: document.getElementById("appointmentBody"),
  conflictsBox: document.getElementById("conflictsBox"),
  conflictsList: document.getElementById("conflictsList"),

  appointmentBackdrop: document.getElementById("appointmentBackdrop"),
  btnAppointmentClose: document.getElementById("btnAppointmentClose"),
  appointmentForm: document.getElementById("appointmentForm"),
  appointmentDate: document.getElementById("appointmentDate"),
  appointmentTime: document.getElementById("appointmentTime"),
  appointmentChannel: document.getElementById("appointmentChannel"),
  appointmentNotes: document.getElementById("appointmentNotes"),

  eventFinderBackdrop: document.getElementById("eventFinderBackdrop"),
  btnEventFinderClose: document.getElementById("btnEventFinderClose"),
  eventFinderSearch: document.getElementById("eventFinderSearch"),
  eventFinderBody: document.getElementById("eventFinderBody"),
  leadsBackdrop: document.getElementById("leadsBackdrop"),
  btnLeadsClose: document.getElementById("btnLeadsClose"),
  leadsSearch: document.getElementById("leadsSearch"),
  leadsStatusFilter: document.getElementById("leadsStatusFilter"),
  leadsSellerFilter: document.getElementById("leadsSellerFilter"),
  leadsDateFrom: document.getElementById("leadsDateFrom"),
  leadsDateTo: document.getElementById("leadsDateTo"),
  btnLeadsReset: document.getElementById("btnLeadsReset"),
  leadsBody: document.getElementById("leadsBody"),

  btnAddUser: document.getElementById("btnAddUser"),
  userBackdrop: document.getElementById("userBackdrop"),
  btnUserClose: document.getElementById("btnUserClose"),
  btnUserDiscard: document.getElementById("btnUserDiscard"),
  userForm: document.getElementById("userForm"),
  userName: document.getElementById("userName"),
  userFullName: document.getElementById("userFullName"),
  userUsername: document.getElementById("userUsername"),
  userEmail: document.getElementById("userEmail"),
  userPhone: document.getElementById("userPhone"),
  userPassword: document.getElementById("userPassword"),
  userSignature: document.getElementById("userSignature"),
  userSignaturePreviewCard: document.getElementById("userSignaturePreviewCard"),
  userSignaturePreview: document.getElementById("userSignaturePreview"),
  userSignatureMeta: document.getElementById("userSignatureMeta"),
  userSignatureWarn: document.getElementById("userSignatureWarn"),
  userAvatar: document.getElementById("userAvatar"),
  userRole: document.getElementById("userRole"),
  userSalesTargetEnabled: document.getElementById("userSalesTargetEnabled"),
  userGoalMonth: document.getElementById("userGoalMonth"),
  userGoalAmount: document.getElementById("userGoalAmount"),
  btnUserGoalAdd: document.getElementById("btnUserGoalAdd"),
  userGoalsBody: document.getElementById("userGoalsBody"),
  userEditSelect: document.getElementById("userEditSelect"),
  userActive: document.getElementById("userActive"),
  btnUserDisable: document.getElementById("btnUserDisable"),
  btnUserSubmit: document.getElementById("btnUserSubmit"),
  userTitle: document.getElementById("userTitle"),

  quoteBackdrop: document.getElementById("quoteBackdrop"),
  quoteDocPanelCard: document.getElementById("quoteDocPanelCard"),
  quoteDocFold: document.getElementById("quoteDocFold"),
  quoteForm: document.getElementById("quoteForm"),
  quoteEventId: document.getElementById("quoteEventId"),
  quoteSubtitle: document.getElementById("quoteSubtitle"),
  btnQuoteClose: document.getElementById("btnQuoteClose"),
  btnQuoteDiscard: document.getElementById("btnQuoteDiscard"),
  quoteVersionSelect: document.getElementById("quoteVersionSelect"),
  quoteTemplateSelect: document.getElementById("quoteTemplateSelect"),
  btnLoadQuoteVersion: document.getElementById("btnLoadQuoteVersion"),
  btnToggleQuoteDoc: document.getElementById("btnToggleQuoteDoc"),
  btnQuoteDocSaveClose: document.getElementById("btnQuoteDocSaveClose"),
  quoteServiceTemplateSelect: document.getElementById("quoteServiceTemplateSelect"),
  quoteServiceTemplateName: document.getElementById("quoteServiceTemplateName"),
  btnQuoteServiceTemplateApply: document.getElementById("btnQuoteServiceTemplateApply"),
  btnQuoteServiceTemplateSave: document.getElementById("btnQuoteServiceTemplateSave"),
  btnQuoteServiceTemplateUpdate: document.getElementById("btnQuoteServiceTemplateUpdate"),
  btnQuoteServiceTemplateDelete: document.getElementById("btnQuoteServiceTemplateDelete"),
  quoteCompanySearch: document.getElementById("quoteCompanySearch"),
  quoteCompanyResults: document.getElementById("quoteCompanyResults"),
  companiesList: document.getElementById("companiesList"),
  quoteCompany: document.getElementById("quoteCompany"),
  quoteManagerSelect: document.getElementById("quoteManagerSelect"),
  quoteContact: document.getElementById("quoteContact"),
  quoteEmail: document.getElementById("quoteEmail"),
  quoteBillTo: document.getElementById("quoteBillTo"),
  quoteAddress: document.getElementById("quoteAddress"),
  quoteEventType: document.getElementById("quoteEventType"),
  quoteVenue: document.getElementById("quoteVenue"),
  quoteSchedule: document.getElementById("quoteSchedule"),
  quoteCode: document.getElementById("quoteCode"),
  quoteDocDate: document.getElementById("quoteDocDate"),
  quotePhone: document.getElementById("quotePhone"),
  quoteNIT: document.getElementById("quoteNIT"),
  quotePeople: document.getElementById("quotePeople"),
  quoteEventDate: document.getElementById("quoteEventDate"),
  quoteFolio: document.getElementById("quoteFolio"),
  quoteEndDate: document.getElementById("quoteEndDate"),
  quoteDueDate: document.getElementById("quoteDueDate"),
  quotePaymentType: document.getElementById("quotePaymentType"),
  quotePaymentTypeSelect: document.getElementById("quotePaymentTypeSelect"),
  btnQuotePaymentAdd: document.getElementById("btnQuotePaymentAdd"),
  btnQuotePaymentClear: document.getElementById("btnQuotePaymentClear"),
  quoteServiceDate: document.getElementById("quoteServiceDate"),
  quoteServiceSearch: document.getElementById("quoteServiceSearch"),
  quoteServiceQty: document.getElementById("quoteServiceQty"),
  quoteServiceQtyHint: document.getElementById("quoteServiceQtyHint"),
  servicesList: document.getElementById("servicesList"),
  serviceDescriptionsList: document.getElementById("serviceDescriptionsList"),
  btnAddServiceToQuote: document.getElementById("btnAddServiceToQuote"),
  quoteItemsPanel: document.getElementById("quoteItemsPanel"),
  quoteSelectAllItems: document.getElementById("quoteSelectAllItems"),
  btnQuoteDuplicateSelected: document.getElementById("btnQuoteDuplicateSelected"),
  btnQuoteMoveSelectedUp: document.getElementById("btnQuoteMoveSelectedUp"),
  btnQuoteMoveSelectedDown: document.getElementById("btnQuoteMoveSelectedDown"),
  btnToggleQuoteItemsExpand: document.getElementById("btnToggleQuoteItemsExpand"),
  quoteItemsBody: document.getElementById("quoteItemsBody"),
  quoteDiscountType: document.getElementById("quoteDiscountType"),
  quoteDiscountValue: document.getElementById("quoteDiscountValue"),
  quoteSubtotal: document.getElementById("quoteSubtotal"),
  quoteDiscountAmount: document.getElementById("quoteDiscountAmount"),
  quoteTotal: document.getElementById("quoteTotal"),
  quoteSidebarDocument: document.getElementById("quoteSidebarDocument"),
  quoteSidebarClient: document.getElementById("quoteSidebarClient"),
  quoteSidebarBreakdown: document.getElementById("quoteSidebarBreakdown"),
  quoteSidebarSubtotal: document.getElementById("quoteSidebarSubtotal"),
  quoteSidebarDiscount: document.getElementById("quoteSidebarDiscount"),
  quoteSidebarTotal: document.getElementById("quoteSidebarTotal"),
  quoteAccountMeta: document.getElementById("quoteAccountMeta"),
  quoteAccountDocument: document.getElementById("quoteAccountDocument"),
  quoteAccountEvent: document.getElementById("quoteAccountEvent"),
  quoteAccountRange: document.getElementById("quoteAccountRange"),
  quoteAccountSummary: document.getElementById("quoteAccountSummary"),
  quoteAccountBody: document.getElementById("quoteAccountBody"),
  btnQuoteCurrencyGTQ: document.getElementById("btnQuoteCurrencyGTQ"),
  btnQuoteCurrencyUSD: document.getElementById("btnQuoteCurrencyUSD"),
  quoteInternalNotes: document.getElementById("quoteInternalNotes"),
  btnMenuMontaje: document.getElementById("btnMenuMontaje"),
  btnQuoteAdvances: document.getElementById("btnQuoteAdvances"),
  btnMenuMontajeSelectable: document.getElementById("btnMenuMontajeSelectable"),
  btnQuotePrintTemplate: document.getElementById("btnQuotePrintTemplate"),
  btnQuoteSave: document.getElementById("btnQuoteSave"),
  quoteSaveProgress: document.getElementById("quoteSaveProgress"),
  quoteSaveProgressText: document.getElementById("quoteSaveProgressText"),
  quoteSaveProgressBar: document.getElementById("quoteSaveProgressBar"),
  quoteSaveProgressFill: document.getElementById("quoteSaveProgressFill"),
  quoteAdvanceBackdrop: document.getElementById("quoteAdvanceBackdrop"),
  btnQuoteAdvanceClose: document.getElementById("btnQuoteAdvanceClose"),
  btnQuoteAdvanceDone: document.getElementById("btnQuoteAdvanceDone"),
  quoteAdvanceAmount: document.getElementById("quoteAdvanceAmount"),
  quoteAdvancePaymentType: document.getElementById("quoteAdvancePaymentType"),
  quoteAdvanceDate: document.getElementById("quoteAdvanceDate"),
  quoteAdvanceVoucherNumber: document.getElementById("quoteAdvanceVoucherNumber"),
  quoteAdvanceDescription: document.getElementById("quoteAdvanceDescription"),
  quoteAdvanceEvidence: document.getElementById("quoteAdvanceEvidence"),
  quoteAdvanceEvidenceHint: document.getElementById("quoteAdvanceEvidenceHint"),
  btnQuoteAdvanceAdd: document.getElementById("btnQuoteAdvanceAdd"),
  quoteAdvanceBody: document.getElementById("quoteAdvanceBody"),
  quoteAdvanceTotal: document.getElementById("quoteAdvanceTotal"),
  quoteAdvancePending: document.getElementById("quoteAdvancePending"),
  quoteAdvanceCredit: document.getElementById("quoteAdvanceCredit"),
  quoteAdvanceHistoryBody: document.getElementById("quoteAdvanceHistoryBody"),

  menuMontajeBackdrop: document.getElementById("menuMontajeBackdrop"),
  btnMenuMontajeClose: document.getElementById("btnMenuMontajeClose"),
  mmDateSalonSelect: document.getElementById("mmDateSalonSelect"),
  mmVersionSelect: document.getElementById("mmVersionSelect"),
  btnMenuMontajeLoadVersion: document.getElementById("btnMenuMontajeLoadVersion"),
  mmDocNo: document.getElementById("mmDocNo"),
  mmMenuTitle: document.getElementById("mmMenuTitle"),
  mmMenuQty: document.getElementById("mmMenuQty"),
  mmMenuDescription: document.getElementById("mmMenuDescription"),
  mmMontajeDescription: document.getElementById("mmMontajeDescription"),
  mmMenuDescCount: document.getElementById("mmMenuDescCount"),
  mmMontajeDescCount: document.getElementById("mmMontajeDescCount"),
  mmEntriesBody: document.getElementById("mmEntriesBody"),
  btnMenuMontajeSave: document.getElementById("btnMenuMontajeSave"),
  btnMenuMontajeSaveCurrent: document.getElementById("btnMenuMontajeSaveCurrent"),
  btnMenuMontajePrintDay: document.getElementById("btnMenuMontajePrintDay"),
  menuMontajeSelectableBackdrop: document.getElementById("menuMontajeSelectableBackdrop"),
  btnMenuMontajeSelectableClose: document.getElementById("btnMenuMontajeSelectableClose"),
  mmsDateSalonSelect: document.getElementById("mmsDateSalonSelect"),
  mmsVersionSelect: document.getElementById("mmsVersionSelect"),
  btnMmsLoadVersion: document.getElementById("btnMmsLoadVersion"),
  mmsDocNo: document.getElementById("mmsDocNo"),
  mmsProtein: document.getElementById("mmsProtein"),
  mmsPreparation: document.getElementById("mmsPreparation"),
  mmsPrimaryTabs: document.getElementById("mmsPrimaryTabs"),
  btnMmsPrimaryMenu: document.getElementById("btnMmsPrimaryMenu"),
  btnMmsPrimaryMontaje: document.getElementById("btnMmsPrimaryMontaje"),
  mmsStageTabs: document.getElementById("mmsStageTabs"),
  btnMmsStagePlato: document.getElementById("btnMmsStagePlato"),
  btnMmsStagePrep: document.getElementById("btnMmsStagePrep"),
  btnMmsStageSalsa: document.getElementById("btnMmsStageSalsa"),
  btnMmsStageGuarnicion: document.getElementById("btnMmsStageGuarnicion"),
  btnMmsStagePostre: document.getElementById("btnMmsStagePostre"),
  btnMmsStageBebida: document.getElementById("btnMmsStageBebida"),
  btnMmsStageMontajeTipo: document.getElementById("btnMmsStageMontajeTipo"),
  btnMmsStageMontajeAdicional: document.getElementById("btnMmsStageMontajeAdicional"),
  mmsStageFilter: document.getElementById("mmsStageFilter"),
  mmsMenuQtyVisible: document.getElementById("mmsMenuQtyVisible"),
  btnMmsStageMoreOptions: document.getElementById("btnMmsStageMoreOptions"),
  btnMmsStageCancelSelection: document.getElementById("btnMmsStageCancelSelection"),
  btnMmsOpenCatalog: document.getElementById("btnMmsOpenCatalog"),
  mmsStageTitle: document.getElementById("mmsStageTitle"),
  mmsStageOptions: document.getElementById("mmsStageOptions"),
  mmsMenuSection: document.getElementById("mmsMenuSection"),
  mmsMenuSectionInput: document.getElementById("mmsMenuSectionInput"),
  btnMmsMenuSectionAdd: document.getElementById("btnMmsMenuSectionAdd"),
  mmsMenuTitle: document.getElementById("mmsMenuTitle"),
  mmsMenuQty: document.getElementById("mmsMenuQty"),
  mmsGuarnicionesSuggested: document.getElementById("mmsGuarnicionesSuggested"),
  mmsGuarnicionesAll: document.getElementById("mmsGuarnicionesAll"),
  mmsGuarnicionFilter: document.getElementById("mmsGuarnicionFilter"),
  btnMmsToggleGuarnicionesGlobal: document.getElementById("btnMmsToggleGuarnicionesGlobal"),
  mmsGuarnicionesQuickSuggested: document.getElementById("mmsGuarnicionesQuickSuggested"),
  mmsGuarnicionesQuickGlobal: document.getElementById("mmsGuarnicionesQuickGlobal"),
  mmsPostresSuggested: document.getElementById("mmsPostresSuggested"),
  mmsPostresAll: document.getElementById("mmsPostresAll"),
  mmsPostreFilter: document.getElementById("mmsPostreFilter"),
  btnMmsTogglePostresGlobal: document.getElementById("btnMmsTogglePostresGlobal"),
  mmsPostresQuickSuggested: document.getElementById("mmsPostresQuickSuggested"),
  mmsPostresQuickGlobal: document.getElementById("mmsPostresQuickGlobal"),
  mmsComandaPreview: document.getElementById("mmsComandaPreview"),
  mmsActivePlateHint: document.getElementById("mmsActivePlateHint"),
  mmsComandaPlato: document.getElementById("mmsComandaPlato"),
  mmsComandaSalsas: document.getElementById("mmsComandaSalsas"),
  mmsComandaGuarniciones: document.getElementById("mmsComandaGuarniciones"),
  mmsComandaPostres: document.getElementById("mmsComandaPostres"),
  mmsComandaBebidas: document.getElementById("mmsComandaBebidas"),
  mmsComandaComentarios: document.getElementById("mmsComandaComentarios"),
  mmsComandaMontaje: document.getElementById("mmsComandaMontaje"),
  mmsPlatoDescripcion: document.getElementById("mmsPlatoDescripcion"),
  mmsComentariosAll: document.getElementById("mmsComentariosAll"),
  mmsBebidaInput: document.getElementById("mmsBebidaInput"),
  btnMmsAddBebida: document.getElementById("btnMmsAddBebida"),
  mmsComentarioLibre: document.getElementById("mmsComentarioLibre"),
  btnMmsUseSuggested: document.getElementById("btnMmsUseSuggested"),
  btnMmsClearMenuSelection: document.getElementById("btnMmsClearMenuSelection"),
  btnMmsMenuAppend: document.getElementById("btnMmsMenuAppend"),
  btnMmsMenuReplace: document.getElementById("btnMmsMenuReplace"),
  mmsSummaryMenu: document.getElementById("mmsSummaryMenu"),
  mmsSummaryGuarniciones: document.getElementById("mmsSummaryGuarniciones"),
  mmsSummaryPostres: document.getElementById("mmsSummaryPostres"),
  mmsSummaryComentarios: document.getElementById("mmsSummaryComentarios"),
  mmsMenuDescription: document.getElementById("mmsMenuDescription"),
  mmsMontajeTipo: document.getElementById("mmsMontajeTipo"),
  mmsMontajeAdicionales: document.getElementById("mmsMontajeAdicionales"),
  mmsMontajeDescription: document.getElementById("mmsMontajeDescription"),
  btnMmsMontajeClear: document.getElementById("btnMmsMontajeClear"),
  btnMmsMontajeAppend: document.getElementById("btnMmsMontajeAppend"),
  btnMmsMontajeReplace: document.getElementById("btnMmsMontajeReplace"),
  mmsSummaryMontajeTipo: document.getElementById("mmsSummaryMontajeTipo"),
  mmsSummaryMontajeAdicionales: document.getElementById("mmsSummaryMontajeAdicionales"),
  mmsEntriesBody: document.getElementById("mmsEntriesBody"),
  btnMmsSave: document.getElementById("btnMmsSave"),
  btnMmsSaveCurrent: document.getElementById("btnMmsSaveCurrent"),
  btnMmsPrintDay: document.getElementById("btnMmsPrintDay"),
  btnAddCompany: document.getElementById("btnAddCompany"),
  btnOpenServiceCreate: document.getElementById("btnOpenServiceCreate"),

  companyBackdrop: document.getElementById("companyBackdrop"),
  companyTitle: document.getElementById("companyTitle"),
  companyForm: document.getElementById("companyForm"),
  companyEditSelect: document.getElementById("companyEditSelect"),
  companyActive: document.getElementById("companyActive"),
  companyName: document.getElementById("companyName"),
  companyOwner: document.getElementById("companyOwner"),
  companyEmail: document.getElementById("companyEmail"),
  companyNIT: document.getElementById("companyNIT"),
  companyBusinessName: document.getElementById("companyBusinessName"),
  companyEventType: document.getElementById("companyEventType"),
  companyAddress: document.getElementById("companyAddress"),
  companyPhone: document.getElementById("companyPhone"),
  companyNotes: document.getElementById("companyNotes"),
  companyRecordSection: document.getElementById("companyRecordSection"),
  companyRecordSummary: document.getElementById("companyRecordSummary"),
  companyRecordBody: document.getElementById("companyRecordBody"),
  managerName: document.getElementById("managerName"),
  managerPhone: document.getElementById("managerPhone"),
  managerEmail: document.getElementById("managerEmail"),
  managerAddress: document.getElementById("managerAddress"),
  btnAddManager: document.getElementById("btnAddManager"),
  managersBody: document.getElementById("managersBody"),
  btnCompanyClose: document.getElementById("btnCompanyClose"),
  btnCompanyDiscard: document.getElementById("btnCompanyDiscard"),
  btnCompanyDisable: document.getElementById("btnCompanyDisable"),

  serviceBackdrop: document.getElementById("serviceBackdrop"),
  serviceTitle: document.getElementById("serviceTitle"),
  serviceForm: document.getElementById("serviceForm"),
  serviceEditSelect: document.getElementById("serviceEditSelect"),
  serviceActive: document.getElementById("serviceActive"),
  serviceName: document.getElementById("serviceName"),
  serviceCategory: document.getElementById("serviceCategory"),
  btnServiceCategoryManage: document.getElementById("btnServiceCategoryManage"),
  serviceSubcategory: document.getElementById("serviceSubcategory"),
  btnServiceSubcategoryManage: document.getElementById("btnServiceSubcategoryManage"),
  servicePrice: document.getElementById("servicePrice"),
  serviceQuantityMode: document.getElementById("serviceQuantityMode"),
  serviceDescription: document.getElementById("serviceDescription"),
  servicesManagerBody: document.getElementById("servicesManagerBody"),
  btnServiceClose: document.getElementById("btnServiceClose"),
  btnServiceDiscard: document.getElementById("btnServiceDiscard"),
  btnServiceDisable: document.getElementById("btnServiceDisable"),
  serviceCategoryBackdrop: document.getElementById("serviceCategoryBackdrop"),
  btnServiceCategoryClose: document.getElementById("btnServiceCategoryClose"),
  serviceCategoryEditSelect: document.getElementById("serviceCategoryEditSelect"),
  serviceCategoryNameInput: document.getElementById("serviceCategoryNameInput"),
  serviceCategoryBody: document.getElementById("serviceCategoryBody"),
  btnServiceCategoryReset: document.getElementById("btnServiceCategoryReset"),
  btnServiceCategorySave: document.getElementById("btnServiceCategorySave"),
  serviceSubcategoryBackdrop: document.getElementById("serviceSubcategoryBackdrop"),
  btnServiceSubcategoryClose: document.getElementById("btnServiceSubcategoryClose"),
  serviceSubcategoryCategorySelect: document.getElementById("serviceSubcategoryCategorySelect"),
  serviceSubcategoryEditSelect: document.getElementById("serviceSubcategoryEditSelect"),
  serviceSubcategoryNameInput: document.getElementById("serviceSubcategoryNameInput"),
  serviceSubcategoryBody: document.getElementById("serviceSubcategoryBody"),
  btnServiceSubcategoryReset: document.getElementById("btnServiceSubcategoryReset"),
  btnServiceSubcategorySave: document.getElementById("btnServiceSubcategorySave"),

  globalGoalsBackdrop: document.getElementById("globalGoalsBackdrop"),
  btnGlobalGoalsClose: document.getElementById("btnGlobalGoalsClose"),
  globalGoalsEditSelect: document.getElementById("globalGoalsEditSelect"),
  globalGoalActive: document.getElementById("globalGoalActive"),
  globalGoalMonth: document.getElementById("globalGoalMonth"),
  globalGoalRole: document.getElementById("globalGoalRole"),
  globalGoalAmount: document.getElementById("globalGoalAmount"),
  globalGoalsBody: document.getElementById("globalGoalsBody"),
  btnGlobalGoalDisable: document.getElementById("btnGlobalGoalDisable"),
  btnGlobalGoalReset: document.getElementById("btnGlobalGoalReset"),
  btnGlobalGoalSave: document.getElementById("btnGlobalGoalSave"),

  salonesBackdrop: document.getElementById("salonesBackdrop"),
  btnSalonesClose: document.getElementById("btnSalonesClose"),
  salonEditSelect: document.getElementById("salonEditSelect"),
  salonActive: document.getElementById("salonActive"),
  salonNameInput: document.getElementById("salonNameInput"),
  salonesBody: document.getElementById("salonesBody"),
  btnSalonDisable: document.getElementById("btnSalonDisable"),
  btnSalonReset: document.getElementById("btnSalonReset"),
  btnSalonSave: document.getElementById("btnSalonSave"),

  menuSuggestionsBackdrop: document.getElementById("menuSuggestionsBackdrop"),
  btnMenuSuggestionsClose: document.getElementById("btnMenuSuggestionsClose"),
  btnMenuSuggestionsDiscard: document.getElementById("btnMenuSuggestionsDiscard"),
  btnMenuSuggestionsSave: document.getElementById("btnMenuSuggestionsSave"),
  btnMenuSuggestionsManageCatalog: document.getElementById("btnMenuSuggestionsManageCatalog"),
  menuSuggestionsProtein: document.getElementById("menuSuggestionsProtein"),
  menuSuggestionsPreparation: document.getElementById("menuSuggestionsPreparation"),
  menuSuggestionsSalsas: document.getElementById("menuSuggestionsSalsas"),
  menuSuggestionsPostres: document.getElementById("menuSuggestionsPostres"),
  menuSuggestionsGuarniciones: document.getElementById("menuSuggestionsGuarniciones"),
  menuSuggestionsBebidas: document.getElementById("menuSuggestionsBebidas"),
  menuSuggestionsMontajeTipos: document.getElementById("menuSuggestionsMontajeTipos"),
  menuSuggestionsMontajeAdicionales: document.getElementById("menuSuggestionsMontajeAdicionales"),

  menuCatalogBackdrop: document.getElementById("menuCatalogBackdrop"),
  btnMenuCatalogClose: document.getElementById("btnMenuCatalogClose"),
  btnMenuCatalogDiscard: document.getElementById("btnMenuCatalogDiscard"),
  btnMenuCatalogOpenSuggestions: document.getElementById("btnMenuCatalogOpenSuggestions"),
  btnMenuCatalogSave: document.getElementById("btnMenuCatalogSave"),
  btnMenuCatalogReset: document.getElementById("btnMenuCatalogReset"),
  menuCatalogKind: document.getElementById("menuCatalogKind"),
  menuCatalogProteinWrap: document.getElementById("menuCatalogProteinWrap"),
  menuCatalogProtein: document.getElementById("menuCatalogProtein"),
  menuCatalogName: document.getElementById("menuCatalogName"),
  menuCatalogDishTypeWrap: document.getElementById("menuCatalogDishTypeWrap"),
  menuCatalogDishType: document.getElementById("menuCatalogDishType"),
  menuCatalogNoProteinWrap: document.getElementById("menuCatalogNoProteinWrap"),
  menuCatalogNoProtein: document.getElementById("menuCatalogNoProtein"),
  menuCatalogBody: document.getElementById("menuCatalogBody"),
};

function goToTodayView() {
  const now = new Date();
  monthCursor = startOfMonth(now);
  viewStart = navMode === "month" ? startOfWeek(monthCursor) : stripTime(now);
}

function loadTopbarSettings() {
  try {
    const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
    if (!raw) return { ...DEFAULT_TOPBAR_SETTINGS };
    const parsed = JSON.parse(raw);
    return {
      showLegend: parsed?.showLegend !== false,
      compactEvents: parsed?.compactEvents === true,
      showWeekends: parsed?.showWeekends !== false,
    };
  } catch (_) {
    return { ...DEFAULT_TOPBAR_SETTINGS };
  }
}

function saveTopbarSettings() {
  try {
    localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(topbarSettings));
  } catch (_) { }
}

function loadPersistedAuthSession() {
  try {
    const raw = localStorage.getItem(AUTH_SESSION_STORAGE_KEY);
    if (!raw) return { userId: "", fullName: "", username: "", avatarDataUrl: "", signatureDataUrl: "" };
    const parsed = JSON.parse(raw);
    return {
      userId: String(parsed?.userId || "").trim(),
      fullName: String(parsed?.fullName || "").trim(),
      username: String(parsed?.username || "").trim(),
      avatarDataUrl: String(parsed?.avatarDataUrl || "").trim(),
      signatureDataUrl: String(parsed?.signatureDataUrl || "").trim(),
    };
  } catch (_) {
    return { userId: "", fullName: "", username: "", avatarDataUrl: "", signatureDataUrl: "" };
  }
}

function savePersistedAuthSession() {
  try {
    if (!String(authSession.userId || "").trim()) {
      localStorage.removeItem(AUTH_SESSION_STORAGE_KEY);
      return;
    }
    localStorage.setItem(AUTH_SESSION_STORAGE_KEY, JSON.stringify({
      userId: String(authSession.userId || "").trim(),
      fullName: String(authSession.fullName || "").trim(),
      username: String(authSession.username || "").trim(),
      avatarDataUrl: String(authSession.avatarDataUrl || "").trim(),
      signatureDataUrl: String(authSession.signatureDataUrl || "").trim(),
    }));
  } catch (_) { }
}

function clearPersistedAuthSession() {
  try {
    localStorage.removeItem(AUTH_SESSION_STORAGE_KEY);
  } catch (_) { }
}

function logoutCurrentSession() {
  authSession = { userId: "", fullName: "", username: "", avatarDataUrl: "", signatureDataUrl: "" };
  clearPersistedAuthSession();
  renderTopbarWelcome();
  refreshTopbarReminders();
  if (el.topbarReminderPanel) el.topbarReminderPanel.hidden = true;
  if (el.btnTopbarReminders) el.btnTopbarReminders.setAttribute("aria-expanded", "false");
  if (el.loginPassword) el.loginPassword.value = "";
  setLoginError("");
  showCalendarModule();
  if (el.loginScreen) el.loginScreen.hidden = false;
  toast("Sesion cerrada.");
}

function loadPersistedModuleScreen() {
  try {
    return String(localStorage.getItem(ACTIVE_MODULE_SCREEN_STORAGE_KEY) || "").trim();
  } catch (_) {
    return "";
  }
}

function savePersistedModuleScreen(screen = getActiveModuleScreenName()) {
  try {
    localStorage.setItem(ACTIVE_MODULE_SCREEN_STORAGE_KEY, String(screen || "").trim());
  } catch (_) { }
}

function restorePersistedAppSession() {
  const persisted = loadPersistedAuthSession();
  const persistedUserId = String(persisted.userId || "").trim();
  if (!persistedUserId) return false;
  const activeUsers = (state.users || []).map(normalizeUserRecord).filter((user) => user.active !== false);
  const matchedUser = activeUsers.find((user) => String(user.id || "") === persistedUserId);
  if (!matchedUser) {
    authSession = { userId: "", fullName: "", username: "", avatarDataUrl: "", signatureDataUrl: "" };
    clearPersistedAuthSession();
    return false;
  }
  authSession = {
    userId: String(matchedUser.id || persistedUserId),
    fullName: String(matchedUser.fullName || matchedUser.name || persisted.fullName || "").trim(),
    username: String(matchedUser.username || persisted.username || "").trim(),
    avatarDataUrl: String(matchedUser.avatarDataUrl || persisted.avatarDataUrl || "").trim(),
    signatureDataUrl: String(matchedUser.signatureDataUrl || persisted.signatureDataUrl || "").trim(),
  };
  savePersistedAuthSession();
  renderTopbarWelcome();
  refreshTopbarReminders();
  if (el.loginScreen) el.loginScreen.hidden = true;
  if (el.eventUser && authSession.userId) {
    syncEnhancedSelectValue(el.eventUser, authSession.userId);
  }
  const targetScreen = loadPersistedModuleScreen();
  if (targetScreen === "settings") {
    showSettingsHub();
  } else if (targetScreen === "reports") {
    showReportsHub();
  } else if (targetScreen === "hub") {
    showCalendarModule();
  } else {
    showCalendarModule();
  }
  return true;
}

function normalizeTemplateRecord(candidate) {
  if (typeof candidate === "string") {
    const name = candidate.trim();
    if (!name) return null;
    return {
      id: uid(),
      name,
      header: "",
      body: "",
      footer: "",
      assets: { pagePdf: "", headerImage: "", footerImage: "" },
      positionedFields: [],
      signatureDefaults: {
        w: TEMPLATE_SIGNATURE_FALLBACK_W_PCT,
        h: TEMPLATE_SIGNATURE_FALLBACK_H_PCT,
      },
      roomRates: [],
      formulas: [],
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
  }
  if (!candidate || typeof candidate !== "object") return null;
  const name = String(candidate.name || "").trim();
  if (!name) return null;
  const formulasRaw = Array.isArray(candidate.formulas) ? candidate.formulas : [];
  const positionedRaw = Array.isArray(candidate.positionedFields) ? candidate.positionedFields : [];
  const roomRatesRaw = Array.isArray(candidate.roomRates) ? candidate.roomRates : [];
  const formulas = formulasRaw
    .map((f) => ({
      key: String(f?.key || "").trim(),
      expression: String(f?.expression || "").trim(),
    }))
    .filter((f) => f.key);
  const positionedFields = positionedRaw
    .map((p) => {
      const token = String(p?.token || "").trim();
      const isSignature = isTemplateSignatureToken(token);
      const minW = isSignature ? TEMPLATE_SIGNATURE_MIN_W_PCT : 4;
      const minH = isSignature ? TEMPLATE_SIGNATURE_MIN_H_PCT : 2;
      const maxW = isSignature ? TEMPLATE_SIGNATURE_MAX_W_PCT : 95;
      const maxH = isSignature ? TEMPLATE_SIGNATURE_MAX_H_PCT : 60;
      return {
        label: String(p?.label || "").trim(),
        token,
        x: clamp(Number(p?.x), 0, 100),
        y: clamp(Number(p?.y), 0, 100),
        w: clamp(Number(p?.w || TEMPLATE_SIGNATURE_FALLBACK_W_PCT), minW, maxW),
        h: clamp(Number(p?.h || TEMPLATE_SIGNATURE_FALLBACK_H_PCT), minH, maxH),
        xPt: Number.isFinite(Number(p?.xPt)) ? Number(p.xPt) : null,
        yPt: Number.isFinite(Number(p?.yPt)) ? Number(p.yPt) : null,
        wPt: Number.isFinite(Number(p?.wPt)) ? Number(p.wPt) : null,
        hPt: Number.isFinite(Number(p?.hPt)) ? Number(p.hPt) : null,
        fontSize: clamp(Number(p?.fontSize || 12), 8, 72),
        fontFamily: String(p?.fontFamily || "Arial").trim() || "Arial",
        bold: p?.bold === true,
        italic: p?.italic === true,
        isSignature,
      };
    })
    .filter((p) => p.label || p.token);
  const roomRates = roomRatesRaw
    .map((r) => ({
      habitacion: String(r?.habitacion || "").trim(),
      precio: Number(r?.precio || 0),
    }))
    .filter((r) => r.habitacion);
  const assets = {
    pagePdf: String(candidate?.assets?.pagePdf || candidate?.assets?.pageImage || "").trim(),
    headerImage: String(candidate?.assets?.headerImage || "").trim(),
    footerImage: String(candidate?.assets?.footerImage || "").trim(),
  };
  const signatureDefaults = normalizeTemplateSignatureDefaults(candidate?.signatureDefaults, positionedFields);
  return {
    id: String(candidate.id || uid()),
    name,
    header: String(candidate.header || "").trim(),
    body: String(candidate.body || "").trim(),
    footer: String(candidate.footer || "").trim(),
    assets,
    positionedFields,
    signatureDefaults,
    roomRates,
    formulas,
    createdAt: String(candidate.createdAt || new Date().toISOString()),
    updatedAt: String(candidate.updatedAt || new Date().toISOString()),
  };
}

function buildCorporateTemplateSeed() {
  const nowIso = new Date().toISOString();
  return normalizeTemplateRecord({
    id: CORPORATE_TEMPLATE_ID,
    name: CORPORATE_TEMPLATE_NAME,
    header: "",
    body: "",
    footer: "",
    assets: {
      pagePdf: "",
      headerImage: DEFAULT_TEMPLATE_HEADER_IMAGE,
      footerImage: DEFAULT_TEMPLATE_FOOTER_IMAGE,
    },
    positionedFields: [],
    signatureDefaults: {
      w: TEMPLATE_SIGNATURE_FALLBACK_W_PCT,
      h: TEMPLATE_SIGNATURE_FALLBACK_H_PCT,
    },
    roomRates: [],
    formulas: [],
    createdAt: nowIso,
    updatedAt: nowIso,
  });
}

function buildServiHospTemplateSeed() {
  const nowIso = new Date().toISOString();
  return normalizeTemplateRecord({
    id: SERVIHOSP_TEMPLATE_ID,
    name: SERVIHOSP_TEMPLATE_NAME,
    header: "",
    body: "",
    footer: "",
    assets: {
      pagePdf: "",
      headerImage: SERVIHOSP_TEMPLATE_HEADER_IMAGE,
      footerImage: "",
    },
    positionedFields: [],
    signatureDefaults: {
      w: TEMPLATE_SIGNATURE_FALLBACK_W_PCT,
      h: TEMPLATE_SIGNATURE_FALLBACK_H_PCT,
    },
    roomRates: [],
    formulas: [],
    createdAt: nowIso,
    updatedAt: nowIso,
  });
}

function buildContractCorpTemplateSeed() {
  const nowIso = new Date().toISOString();
  return normalizeTemplateRecord({
    id: CONTRACT_CORP_TEMPLATE_ID,
    name: CONTRACT_CORP_TEMPLATE_NAME,
    header: "",
    body: "",
    footer: "",
    assets: {
      pagePdf: "",
      headerImage: "",
      footerImage: "",
    },
    positionedFields: [],
    signatureDefaults: {
      w: TEMPLATE_SIGNATURE_FALLBACK_W_PCT,
      h: TEMPLATE_SIGNATURE_FALLBACK_H_PCT,
    },
    roomRates: [],
    formulas: [],
    createdAt: nowIso,
    updatedAt: nowIso,
  });
}

function ensureCorporateTemplateSeed(listLike) {
  const list = Array.isArray(listLike)
    ? listLike.map(normalizeTemplateRecord).filter(Boolean)
    : [];
  const corporateSeed = buildCorporateTemplateSeed();
  const serviHospSeed = buildServiHospTemplateSeed();
  const contractSeed = buildContractCorpTemplateSeed();
  const byId = list.find((t) => String(t?.id || "").trim() === CORPORATE_TEMPLATE_ID) || null;
  const byName = list.find((t) => /corporativ/i.test(String(t?.name || ""))) || null;
  const rich = list.find((t) =>
    String(t?.assets?.pagePdf || "").trim() &&
    Array.isArray(t?.positionedFields) &&
    t.positionedFields.length > 0
  ) || null;
  const serviHospById = list.find((t) => String(t?.id || "").trim() === SERVIHOSP_TEMPLATE_ID) || null;
  const serviHospByName = list.find((t) => /servi\s*hosp/i.test(String(t?.name || ""))) || null;
  const contractById = list.find((t) => String(t?.id || "").trim() === CONTRACT_CORP_TEMPLATE_ID) || null;
  const contractByName = list.find((t) => /contrato\s+corporativ/i.test(String(t?.name || ""))) || null;
  const base = byId || byName || rich || corporateSeed;
  const corporate = normalizeTemplateRecord({
    ...corporateSeed,
    ...base,
    id: CORPORATE_TEMPLATE_ID,
    name: CORPORATE_TEMPLATE_NAME,
    assets: {
      pagePdf: String(base?.assets?.pagePdf || "").trim(),
      headerImage: String(base?.assets?.headerImage || corporateSeed.assets.headerImage || "").trim(),
      footerImage: String(base?.assets?.footerImage || corporateSeed.assets.footerImage || "").trim(),
    },
  });
  const serviHosp = normalizeTemplateRecord({
    ...serviHospSeed,
    ...(serviHospById || serviHospByName || {}),
    id: SERVIHOSP_TEMPLATE_ID,
    name: SERVIHOSP_TEMPLATE_NAME,
    assets: {
      pagePdf: String(serviHospById?.assets?.pagePdf || serviHospByName?.assets?.pagePdf || "").trim(),
      headerImage: String(
        serviHospById?.assets?.headerImage
        || serviHospByName?.assets?.headerImage
        || serviHospSeed.assets.headerImage
        || ""
      ).trim(),
      footerImage: String(
        serviHospById?.assets?.footerImage
        || serviHospByName?.assets?.footerImage
        || serviHospSeed.assets.footerImage
        || ""
      ).trim(),
    },
  });
  const contract = normalizeTemplateRecord({
    ...contractSeed,
    ...(contractById || contractByName || {}),
    id: CONTRACT_CORP_TEMPLATE_ID,
    name: CONTRACT_CORP_TEMPLATE_NAME,
  });
  const out = [corporate, serviHosp, contract];
  for (const t of list) {
    const id = String(t?.id || "").trim();
    if (
      !id ||
      id === String(base?.id || "").trim() ||
      id === String(serviHospById?.id || "").trim() ||
      id === String(serviHospByName?.id || "").trim() ||
      id === String(contractById?.id || "").trim() ||
      id === String(contractByName?.id || "").trim() ||
      id === CORPORATE_TEMPLATE_ID ||
      id === SERVIHOSP_TEMPLATE_ID ||
      id === CONTRACT_CORP_TEMPLATE_ID
    ) continue;
    out.push(t);
  }
  return out.filter(Boolean);
}

function loadQuickTemplates() {
  try {
    const raw = localStorage.getItem(QUICK_TEMPLATES_STORAGE_KEY);
    if (!raw) return ensureCorporateTemplateSeed([]);
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return ensureCorporateTemplateSeed([]);
    return ensureCorporateTemplateSeed(parsed
      .map(normalizeTemplateRecord)
      .filter(Boolean));
  } catch (_) {
    return ensureCorporateTemplateSeed([]);
  }
}

function syncQuickTemplatesIntoState() {
  if (!state || typeof state !== "object") return;
  quickTemplates = ensureCorporateTemplateSeed(quickTemplates);
  state.quickTemplates = quickTemplates;
}

function saveQuickTemplates({ persistRemote = true, backupLocal = true } = {}) {
  syncQuickTemplatesIntoState();
  if (backupLocal) {
    try {
      localStorage.setItem(QUICK_TEMPLATES_STORAGE_KEY, JSON.stringify(quickTemplates));
    } catch (_) { }
  }
  if (persistRemote) persist();
}

function backupQuickTemplatesLocal() {
  try {
    localStorage.setItem(QUICK_TEMPLATES_STORAGE_KEY, JSON.stringify(quickTemplates));
  } catch (_) { }
}

function buildTemplatePrintContextFromQuoteForm() {
  const selectedDate = String(el.quoteDocDate?.value || el.quoteEventDate?.value || toISODate(new Date())).trim();
  const d = selectedDate ? new Date(`${selectedDate}T00:00:00`) : new Date();
  const safeDate = Number.isNaN(d.getTime()) ? new Date() : d;
  const monthName = safeDate.toLocaleDateString("es-GT", { month: "long" });
  const company = (state.companies || []).find((c) => String(c.id || "") === String(el.quoteCompany?.value || ""));
  const quoteEventId = String(el.quoteEventId?.value || "").trim();
  const ev = quoteEventId ? (state.events || []).find((x) => String(x.id || "") === quoteEventId) : null;
  const sellerUser = normalizeUserRecord((state.users || []).find((u) => String(u.id || "") === String(ev?.userId || "")) || {});
  const authUser = normalizeUserRecord(getAuthUserRecord() || {});
  const vendorSignature = String(
    sellerUser?.signatureDataUrl
    || authSession.signatureDataUrl
    || authUser?.signatureDataUrl
    || ""
  ).trim();
  const vendorName = String(
    sellerUser?.fullName
    || sellerUser?.name
    || authSession.fullName
    || authUser?.fullName
    || authUser?.name
    || ""
  ).trim();
  const vendorPhone = String(sellerUser?.phone || authUser?.phone || "").trim();
  const vendorEmail = String(sellerUser?.email || authUser?.email || "").trim();
  const clientName = String(el.quoteCompanySearch?.value || company?.name || quoteDraft?.companyName || "").trim();
  const clientContact = String(el.quoteContact?.value || quoteDraft?.contact || company?.owner || clientName).trim();
  const clientPhone = String(el.quotePhone?.value || quoteDraft?.phone || company?.phone || "").trim();
  const clientEmail = String(el.quoteEmail?.value || quoteDraft?.email || company?.email || "").trim();
  const billTo = String(el.quoteBillTo?.value || quoteDraft?.billTo || company?.billTo || company?.businessName || clientName).trim();
  const quoteAddress = String(el.quoteAddress?.value || quoteDraft?.address || company?.address || "").trim();
  const eventType = String(el.quoteEventType?.value || quoteDraft?.eventType || "").trim();
  const eventDate = String(el.quoteEventDate?.value || quoteDraft?.eventDate || ev?.date || "").trim();
  const schedule = String(el.quoteSchedule?.value || quoteDraft?.schedule || `${ev?.startTime || ""} a ${ev?.endTime || ""}`.trim()).trim();
  const people = String(el.quotePeople?.value || quoteDraft?.people || ev?.pax || "").trim();
  const paymentType = String(el.quotePaymentType?.value || quoteDraft?.paymentType || "").trim();
  const dueDate = String(el.quoteDueDate?.value || quoteDraft?.dueDate || "").trim();
  const quoteNotes = String(el.quoteInternalNotes?.value || quoteDraft?.internalNotes || quoteDraft?.notes || "").trim();
  const venue = String(el.quoteVenue?.value || "").trim();
  const departmentRaw = String(el.quoteAddress?.value || company?.address || "").trim();
  const department = departmentRaw || "Solola";
  return {
    NO_DOC: String(el.quoteCode?.value || quoteDraft?.code || "").trim(),
    CLIENTE: clientName,
    CLIENTE_EMPRESA: clientName,
    LUGAR: venue || "Panajachel",
    DEPARTAMENTO: department,
    DIA: String(safeDate.getDate()),
    MES: String(monthName || "").trim(),
    ANIO: String(safeDate.getFullYear()),
    VENDEDOR_FIRMA_URL: vendorSignature,
    VENDEDOR_NOMBRE: vendorName,
    VENDEDOR_TELEFONO: vendorPhone,
    VENDEDOR_CORREO: vendorEmail,
    CLIENTE_NOMBRE: clientContact,
    CLIENTE_TELEFONO: clientPhone,
    CLIENTE_CORREO: clientEmail,
    CLIENTE_FACTURAR_A: billTo,
    CLIENTE_DIRECCION: quoteAddress,
    EVENTO_TIPO: eventType,
    EVENTO_FECHA: eventDate,
    EVENTO_HORARIO: schedule,
    EVENTO_PAX: people,
    PAGO_TIPO: paymentType,
    PAGO_FECHA_MAXIMA: dueDate,
    OBSERVACIONES: quoteNotes,
  };
}

function buildTemplatePrintContextFromQuoteData(ev, quote, company, manager) {
  const selectedDate = String(quote?.docDate || quote?.eventDate || ev?.date || toISODate(new Date())).trim();
  const d = selectedDate ? new Date(`${selectedDate}T00:00:00`) : new Date();
  const safeDate = Number.isNaN(d.getTime()) ? new Date() : d;
  const monthName = safeDate.toLocaleDateString("es-GT", { month: "long" });
  const sellerUser = normalizeUserRecord((state.users || []).find((u) => String(u.id || "") === String(ev?.userId || "")) || {});
  const authUser = normalizeUserRecord(getAuthUserRecord() || {});
  const vendorSignature = String(
    sellerUser?.signatureDataUrl
    || authSession.signatureDataUrl
    || authUser?.signatureDataUrl
    || ""
  ).trim();
  const vendorName = String(
    sellerUser?.fullName
    || sellerUser?.name
    || authSession.fullName
    || authUser?.fullName
    || authUser?.name
    || ""
  ).trim();
  const vendorPhone = String(sellerUser?.phone || authUser?.phone || "").trim();
  const vendorEmail = String(sellerUser?.email || authUser?.email || "").trim();
  const clientCompany = String(quote?.companyName || company?.name || "").trim();
  const clientContact = String(quote?.contact || manager?.name || company?.owner || clientCompany).trim();
  const clientPhone = String(quote?.phone || manager?.phone || company?.phone || "").trim();
  const clientEmail = String(quote?.email || manager?.email || company?.email || "").trim();
  const billTo = String(quote?.billTo || company?.billTo || company?.businessName || clientCompany).trim();
  const quoteAddress = String(quote?.address || company?.address || "").trim();
  const eventType = String(quote?.eventType || ev?.name || "").trim();
  const eventDate = String(quote?.eventDate || ev?.date || "").trim();
  const schedule = String(quote?.schedule || `${ev?.startTime || ""} a ${ev?.endTime || ""}`.trim()).trim();
  const people = String(quote?.people || ev?.pax || "").trim();
  const paymentType = String(quote?.paymentType || "").trim();
  const dueDate = String(quote?.dueDate || "").trim();
  const quoteNotes = String(quote?.internalNotes || quote?.notes || "").trim();
  const venue = String(quote?.venue || ev?.salon || "").trim();
  const departmentRaw = String(quote?.address || company?.address || "").trim();
  const department = departmentRaw || "Solola";
  return {
    NO_DOC: String(quote?.code || "").trim(),
    CLIENTE: clientCompany,
    CLIENTE_EMPRESA: clientCompany,
    LUGAR: venue || "Panajachel",
    DEPARTAMENTO: department,
    DIA: String(safeDate.getDate()),
    MES: String(monthName || "").trim(),
    ANIO: String(safeDate.getFullYear()),
    VENDEDOR_FIRMA_URL: vendorSignature,
    VENDEDOR_NOMBRE: vendorName,
    VENDEDOR_TELEFONO: vendorPhone,
    VENDEDOR_CORREO: vendorEmail,
    CLIENTE_NOMBRE: clientContact,
    CLIENTE_TELEFONO: clientPhone,
    CLIENTE_CORREO: clientEmail,
    CLIENTE_FACTURAR_A: billTo,
    CLIENTE_DIRECCION: quoteAddress,
    EVENTO_TIPO: eventType,
    EVENTO_FECHA: eventDate,
    EVENTO_HORARIO: schedule,
    EVENTO_PAX: people,
    PAGO_TIPO: paymentType,
    PAGO_FECHA_MAXIMA: dueDate,
    OBSERVACIONES: quoteNotes,
  };
}

function fillTemplateHtmlTokens(htmlText, contextMap) {
  let out = String(htmlText || "");
  const pairs = Object.entries(contextMap || {});
  for (const [key, rawValue] of pairs) {
    const token = "{{" + String(key) + "}}";
    const textValue = String(rawValue || "");
    const value = /_URL$/i.test(String(key)) ? textValue : escapeHtml(textValue);
    out = out.split(token).join(value);
  }
  return out;
}

function getBuiltInQuoteTemplateMeta(templateId = "") {
  const id = String(templateId || "").trim();
  if (!id) return null;
  if (id === CONTRACT_CORP_TEMPLATE_ID) {
    return {
      file: "./Jardines.html",
      label: "Jardines",
      attachToQuote: true,
      headerImage: "",
    };
  }
  if (id === SERVIHOSP_TEMPLATE_ID) {
    return {
      file: "./ServiHosp.html",
      label: "Servi Hosp",
      attachToQuote: true,
      headerImage: SERVIHOSP_TEMPLATE_HEADER_IMAGE,
    };
  }
  if (id === CORPORATE_TEMPLATE_ID) {
    return {
      file: "./Corporativo.html",
      label: "Corporativo",
      attachToQuote: false,
      headerImage: DEFAULT_TEMPLATE_HEADER_IMAGE,
    };
  }
  return null;
}

function applyBuiltInTemplateAssets(htmlText, meta) {
  const out = String(htmlText || "");
  if (!meta || !String(meta.headerImage || "").trim()) return out;
  return out.replace(/Encabezadojdl\.png/g, String(meta.headerImage).trim());
}

async function printSelectedQuoteTemplate() {
  const selectedTemplateId = String(el.quoteTemplateSelect?.value || quoteDraft?.templateId || "").trim();
  if (!selectedTemplateId) return toast("Selecciona una plantilla.");
  const activeTemplateMeta = getBuiltInQuoteTemplateMeta(selectedTemplateId);
  if (!activeTemplateMeta) return toast("La plantilla seleccionada no usa formato HTML imprimible.");
  const win = window.open("", "_blank");
  if (!win) return toast("Tu navegador bloqueo la ventana emergente.");
  try {
    const res = await fetch(activeTemplateMeta.file, { cache: "no-store" });
    if (!res.ok) throw new Error(`HTTP ${res.status}`);
    let html = await res.text();
    html = applyBuiltInTemplateAssets(html, activeTemplateMeta);
    html = html.replace("<head>", `<head><base href="${escapeHtml(String(window.location.href || ""))}" />`);
    const ctx = buildTemplatePrintContextFromQuoteForm();
    html = fillTemplateHtmlTokens(html, ctx);
    win.document.open();
    win.document.write(html);
    win.document.close();
    setTimeout(() => {
      try {
        win.focus();
        win.print();
      } catch (_) { }
    }, 700);
  } catch (err) {
    try { win.close(); } catch (_) { }
    console.error(`No se pudo imprimir plantilla ${activeTemplateMeta.label}:`, err?.message || err);
    toast(`No se pudo abrir la plantilla ${activeTemplateMeta.label}.`);
  }
}

async function promptTextRequired({ title, label = "", placeholder = "" }) {
  if (window.Swal && typeof window.Swal.fire === "function") {
    const result = await window.Swal.fire({
      title,
      input: "text",
      inputLabel: label || undefined,
      inputPlaceholder: placeholder || undefined,
      showCancelButton: true,
      confirmButtonText: "Guardar",
      cancelButtonText: "Cancelar",
      background: "#f8fbff",
      color: "#10243b",
      confirmButtonColor: "#2563eb",
      inputValidator: (value) => {
        if (!String(value || "").trim()) return "Este campo es obligatorio.";
        return null;
      },
    });
    if (!result.isConfirmed) return null;
    return String(result.value || "").trim();
  }

  const raw = window.prompt(title, "");
  const value = String(raw || "").trim();
  return value || null;
}

async function promptSelectRequired({ title, options = [], label = "Selecciona una opcion" }) {
  const cleaned = Array.isArray(options) ? options.filter((o) => o && o.value !== undefined) : [];
  if (!cleaned.length) return null;
  if (window.Swal && typeof window.Swal.fire === "function") {
    const inputOptions = {};
    for (const o of cleaned) inputOptions[String(o.value)] = String(o.label || o.value);
    const result = await window.Swal.fire({
      title,
      input: "select",
      inputLabel: label,
      inputOptions,
      showCancelButton: true,
      confirmButtonText: "Continuar",
      cancelButtonText: "Cancelar",
      background: "#f8fbff",
      color: "#10243b",
      confirmButtonColor: "#2563eb",
      inputValidator: (value) => {
        if (!String(value || "").trim()) return "Debes seleccionar una opcion.";
        return null;
      },
    });
    if (!result.isConfirmed) return null;
    return String(result.value || "").trim();
  }
  const labels = cleaned.map((o, i) => `${i + 1}. ${o.label}`).join("\n");
  const raw = window.prompt(`${title}\n${labels}`, "");
  const idx = Number(raw);
  if (!Number.isFinite(idx) || idx < 1 || idx > cleaned.length) return null;
  return String(cleaned[idx - 1].value || "");
}

async function promptCrudAction(entityLabel) {
  return promptSelectRequired({
    title: `${entityLabel}: accion`,
    label: "Que deseas hacer",
    options: [
      { value: "add", label: "Agregar" },
      { value: "edit", label: "Editar" },
      { value: "disable", label: "Inhabilitar" },
    ],
  });
}

function isCompanyDisabled(companyId) {
  return (state.disabledCompanies || []).includes(String(companyId || ""));
}

function isServiceDisabled(serviceId) {
  return (state.disabledServices || []).includes(String(serviceId || ""));
}

function isManagerDisabled(managerId) {
  return (state.disabledManagers || []).includes(String(managerId || ""));
}

function isSalonDisabled(name) {
  const needle = String(name || "").trim().toLowerCase();
  if (!needle) return false;
  return (state.disabledSalones || []).some((s) => String(s || "").trim().toLowerCase() === needle);
}

function enableCompany(companyId) {
  const id = String(companyId || "").trim();
  state.disabledCompanies = (state.disabledCompanies || []).filter((x) => String(x) !== id);
}

function enableService(serviceId) {
  const id = String(serviceId || "").trim();
  state.disabledServices = (state.disabledServices || []).filter((x) => String(x) !== id);
}

function enableManager(managerId) {
  const id = String(managerId || "").trim();
  state.disabledManagers = (state.disabledManagers || []).filter((x) => String(x) !== id);
}

function enableSalon(name) {
  const needle = String(name || "").trim().toLowerCase();
  state.disabledSalones = (state.disabledSalones || []).filter((s) => String(s || "").trim().toLowerCase() !== needle);
}

async function manageSalonesFromQuickMenu() {
  const action = await promptCrudAction("Salones");
  if (!action) return;
  if (action === "add") {
    const name = await promptTextRequired({
      title: "Agregar salon",
      label: "Nombre del salon",
      placeholder: "Ej: Salon Aurora",
    });
    if (!name) return;
    const exists = (state.salones || []).some((s) => String(s || "").toLowerCase() === name.toLowerCase());
    if (exists) {
      enableSalon(name);
      persist();
      renderRoomSelects();
      return toast("Salon habilitado nuevamente.");
    }
    state.salones.push(name);
    state.salones.sort((a, b) => String(a).localeCompare(String(b), "es", { sensitivity: "base" }));
    enableSalon(name);
    renderRoomSelects();
    persist();
    return toast("Salon agregado.");
  }

  const all = (state.salones || []).slice().sort((a, b) => String(a).localeCompare(String(b), "es", { sensitivity: "base" }));
  if (!all.length) return toast("No hay salones registrados.");
  const selected = await promptSelectRequired({
    title: action === "edit" ? "Editar salon" : "Inhabilitar salon",
    options: all.map((name) => ({
      value: name,
      label: `${name}${isSalonDisabled(name) ? " (Inhabilitado)" : ""}`,
    })),
  });
  if (!selected) return;

  if (action === "edit") {
    const nextName = await promptTextRequired({
      title: "Nuevo nombre de salon",
      label: `Actual: ${selected}`,
      placeholder: selected,
    });
    if (!nextName) return;
    const exists = all.some((s) => String(s).toLowerCase() === nextName.toLowerCase() && String(s) !== String(selected));
    if (exists) return toast("Ya existe un salon con ese nombre.");
    state.salones = (state.salones || []).map((s) => (String(s) === String(selected) ? nextName : s));
    if (isSalonDisabled(selected)) {
      state.disabledSalones.push(nextName);
      enableSalon(selected);
    }
    state.salones.sort((a, b) => String(a).localeCompare(String(b), "es", { sensitivity: "base" }));
    renderRoomSelects();
    persist();
    return toast("Salon actualizado.");
  }

  if (!isSalonDisabled(selected)) {
    state.disabledSalones = Array.from(new Set([...(state.disabledSalones || []), selected]));
    renderRoomSelects();
    persist();
  }
  toast("Salon inhabilitado.");
}

async function manageInstitutionsFromQuickMenu() {
  const action = await promptCrudAction("Empresas");
  if (!action) return;
  if (action === "add") return openCompanyModal("");

  const companies = (state.companies || []).slice().sort((a, b) =>
    String(a.name || "").localeCompare(String(b.name || ""), "es", { sensitivity: "base" })
  );
  if (!companies.length) return toast("No hay instituciones registradas.");
  const selectedId = await promptSelectRequired({
    title: action === "edit" ? "Editar empresa" : "Inhabilitar empresa",
    options: companies.map((c) => ({
      value: c.id,
      label: `${c.name}${isCompanyDisabled(c.id) ? " (Inhabilitada)" : ""}`,
    })),
  });
  if (!selectedId) return;

  if (action === "edit") {
    enableCompany(selectedId);
    return openCompanyModal(selectedId);
  }
  if (!isCompanyDisabled(selectedId)) {
    state.disabledCompanies = Array.from(new Set([...(state.disabledCompanies || []), selectedId]));
    persist();
    renderCompaniesSelect();
  }
  toast("Empresa inhabilitada.");
}

async function manageServicesFromQuickMenu() {
  const action = await promptCrudAction("Servicios");
  if (!action) return;
  if (action === "add") return openServiceModal();

  const services = (state.services || []).slice().sort((a, b) =>
    String(a.name || "").localeCompare(String(b.name || ""), "es", { sensitivity: "base" })
  );
  if (!services.length) return toast("No hay servicios registrados.");
  const selectedId = await promptSelectRequired({
    title: action === "edit" ? "Editar servicio" : "Inhabilitar servicio",
    options: services.map((s) => ({
      value: s.id,
      label: `${s.name}${isServiceDisabled(s.id) ? " (Inhabilitado)" : ""}`,
    })),
  });
  if (!selectedId) return;
  if (action === "edit") {
    enableService(selectedId);
    return openServiceModal(selectedId);
  }
  if (!isServiceDisabled(selectedId)) {
    state.disabledServices = Array.from(new Set([...(state.disabledServices || []), selectedId]));
    persist();
    renderServicesList();
  }
  toast("Servicio inhabilitado.");
}

function buildManagersCatalog() {
  const rows = [];
  for (const c of state.companies || []) {
    for (const m of c.managers || []) {
      rows.push({
        companyId: c.id,
        companyName: c.name,
        manager: m,
      });
    }
  }
  return rows;
}

async function editManagerFlow() {
  const catalog = buildManagersCatalog();
  if (!catalog.length) return toast("No hay encargados registrados.");
  const selectedValue = await promptSelectRequired({
    title: "Editar encargado",
    options: catalog.map((x) => ({
      value: `${x.companyId}::${x.manager.id}`,
      label: `${x.manager.name} (${x.companyName})${isManagerDisabled(x.manager.id) ? " (Inhabilitado)" : ""}`,
    })),
  });
  if (!selectedValue) return;
  const [companyId, managerId] = String(selectedValue).split("::");
  const company = (state.companies || []).find((c) => String(c.id) === String(companyId));
  const manager = company?.managers?.find((m) => String(m.id) === String(managerId));
  if (!company || !manager) return toast("Encargado no encontrado.");

  const name = await promptTextRequired({ title: "Nombre encargado", label: "Nombre completo", placeholder: manager.name || "" });
  if (!name) return;
  const phone = await promptTextRequired({ title: "Telefono encargado", label: "Telefono", placeholder: manager.phone || "" });
  if (!phone) return;
  const email = await promptTextRequired({ title: "Correo encargado", label: "Correo", placeholder: manager.email || "" });
  if (!email || !isValidEmail(email)) return toast("Correo de encargado invalido.");
  const addressRaw = window.prompt("Direccion (opcional)", manager.address || "");
  const address = addressRaw === null ? String(manager.address || "") : String(addressRaw || "");

  manager.name = name;
  manager.phone = phone;
  manager.email = email;
  manager.address = String(address || "").trim();
  enableManager(manager.id);
  persist();
  renderCompaniesSelect(company.id);
  toast("Encargado actualizado.");
}

async function addManagerQuickFlow() {
  const companies = (state.companies || []).slice().sort((a, b) =>
    String(a.name || "").localeCompare(String(b.name || ""), "es", { sensitivity: "base" })
  );
  if (!companies.length) return toast("Primero crea una institucion.");
  const companyId = await promptSelectRequired({
    title: "Agregar encargado",
    label: "Institucion",
    options: companies.map((c) => ({ value: c.id, label: c.name })),
  });
  if (!companyId) return;
  const company = (state.companies || []).find((c) => String(c.id) === String(companyId));
  if (!company) return;
  const name = await promptTextRequired({ title: "Nombre encargado", label: "Nombre completo", placeholder: "Ej: Pedro Juan" });
  if (!name) return;
  const phone = await promptTextRequired({ title: "Telefono encargado", label: "Telefono", placeholder: "Ej: 55551234" });
  if (!phone) return;
  const email = await promptTextRequired({ title: "Correo encargado", label: "Correo", placeholder: "correo@dominio.com" });
  if (!email || !isValidEmail(email)) return toast("Correo de encargado invalido.");
  const address = window.prompt("Direccion (opcional)", "") || "";
  const manager = {
    id: uid(),
    name,
    phone,
    email,
    address: String(address || "").trim(),
  };
  company.managers = Array.isArray(company.managers) ? company.managers : [];
  company.managers.push(manager);
  enableManager(manager.id);
  persist();
  renderCompaniesSelect(company.id);
  toast("Encargado agregado.");
}

async function disableManagerFlow() {
  const catalog = buildManagersCatalog();
  if (!catalog.length) return toast("No hay encargados registrados.");
  const selectedValue = await promptSelectRequired({
    title: "Inhabilitar encargado",
    options: catalog.map((x) => ({
      value: `${x.companyId}::${x.manager.id}`,
      label: `${x.manager.name} (${x.companyName})${isManagerDisabled(x.manager.id) ? " (Inhabilitado)" : ""}`,
    })),
  });
  if (!selectedValue) return;
  const [, managerId] = String(selectedValue).split("::");
  if (!managerId) return;
  if (!isManagerDisabled(managerId)) {
    state.disabledManagers = Array.from(new Set([...(state.disabledManagers || []), managerId]));
    persist();
    renderCompaniesSelect();
  }
  toast("Encargado inhabilitado.");
}

async function manageManagersFromQuickMenu() {
  const action = await promptCrudAction("Encargados");
  if (!action) return;
  if (action === "add") return addManagerQuickFlow();
  if (action === "edit") return editManagerFlow();
  return disableManagerFlow();
}

function normalizeGlobalMonthlyGoalRecord(goal) {
  return {
    month: String(goal?.month || "").trim(),
    role: normalizeUserRole(goal?.role || USER_ROLE.SELLER),
    amount: Math.max(0, Number(goal?.amount || 0)),
    active: goal?.active === false ? false : true,
  };
}

function globalGoalCompositeKey(goalLike) {
  const row = normalizeGlobalMonthlyGoalRecord(goalLike || {});
  return `${row.month}|${row.role}`;
}

function getGlobalMonthlyGoals(options = {}) {
  const includeInactive = options && options.includeInactive === true;
  const role = options && options.role ? normalizeUserRole(options.role) : "";
  const rows = Array.isArray(state.globalMonthlyGoals) ? state.globalMonthlyGoals : [];
  return rows
    .map(normalizeGlobalMonthlyGoalRecord)
    .filter((g) => /^\d{4}-\d{2}$/.test(g.month))
    .filter((g) => !role || g.role === role)
    .filter((g) => includeInactive || g.active !== false)
    .sort((a, b) => a.month.localeCompare(b.month) || userRoleLabel(a.role).localeCompare(userRoleLabel(b.role), "es", { sensitivity: "base" }));
}

function formatMonthKeyLabel(monthKey) {
  const raw = String(monthKey || "").trim();
  if (!/^\d{4}-\d{2}$/.test(raw)) return raw;
  const parsed = new Date(`${raw}-01T00:00:00`);
  if (Number.isNaN(parsed.getTime())) return raw;
  const label = fmtMonthYear(parsed);
  return label ? `${label.charAt(0).toUpperCase()}${label.slice(1)}` : raw;
}

function buildGlobalGoalMonthOptions() {
  const months = new Set();
  const now = new Date();
  const base = new Date(now.getFullYear() - 2, 0, 1);
  for (let i = 0; i < 96; i++) {
    const d = new Date(base.getFullYear(), base.getMonth() + i, 1);
    const key = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
    months.add(key);
  }
  for (const row of getGlobalMonthlyGoals({ includeInactive: true })) {
    if (/^\d{4}-\d{2}$/.test(row.month)) months.add(row.month);
  }
  return Array.from(months).sort((a, b) => a.localeCompare(b));
}

function renderGlobalGoalMonthSelect(selectedMonth = "") {
  if (!el.globalGoalMonth) return;
  const keep = String(selectedMonth || "").trim();
  const options = buildGlobalGoalMonthOptions();
  el.globalGoalMonth.innerHTML = `<option value="">Selecciona un mes</option>`;
  for (const monthKey of options) {
    const opt = document.createElement("option");
    opt.value = monthKey;
    opt.textContent = formatMonthKeyLabel(monthKey);
    el.globalGoalMonth.appendChild(opt);
  }
  el.globalGoalMonth.value = keep;
}

function renderGlobalGoalsEditSelect(selectedMonth = "") {
  if (!el.globalGoalsEditSelect) return;
  const keep = String(selectedMonth || "").trim();
  const rows = getGlobalMonthlyGoals({ includeInactive: true });
  el.globalGoalsEditSelect.innerHTML = `<option value="">Crear nueva meta</option>`;
  for (const row of rows) {
    const opt = document.createElement("option");
    opt.value = globalGoalCompositeKey(row);
    opt.textContent = `${formatMonthKeyLabel(row.month)} | ${userRoleLabel(row.role)}${row.active === false ? " (Inhabilitada)" : ""}`;
    el.globalGoalsEditSelect.appendChild(opt);
  }
  el.globalGoalsEditSelect.value = keep;
}

function updateGlobalGoalModalControls(month = "") {
  const key = String(month || "").trim();
  const target = key ? getGlobalMonthlyGoals({ includeInactive: true }).find((g) => globalGoalCompositeKey(g) === key) : null;
  if (el.globalGoalActive) el.globalGoalActive.checked = target ? target.active !== false : true;
  if (el.btnGlobalGoalDisable) {
    el.btnGlobalGoalDisable.disabled = !target;
    el.btnGlobalGoalDisable.textContent = target && target.active === false ? "Reactivar" : "Inhabilitar";
  }
}

function renderGlobalGoalsTable() {
  if (!el.globalGoalsBody) return;
  const rows = getGlobalMonthlyGoals({ includeInactive: true });
  el.globalGoalsBody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="5">Sin metas globales registradas.</td>`;
    el.globalGoalsBody.appendChild(tr);
    return;
  }
  for (const row of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(formatMonthKeyLabel(row.month))}</td>
      <td>${escapeHtml(userRoleLabel(row.role))}</td>
      <td>${moneyGT(row.amount)}</td>
      <td>${row.active === false ? "Inhabilitada" : "Activa"}</td>
      <td class="appointmentActions">
        <button class="apptIconBtn apptEdit" type="button" data-global-goal-edit="${escapeHtml(globalGoalCompositeKey(row))}" title="Editar" aria-label="Editar">&#9998;</button>
        <button class="apptIconBtn ${row.active === false ? "" : "apptDelete"}" type="button" data-global-goal-toggle="${escapeHtml(globalGoalCompositeKey(row))}" title="${row.active === false ? "Reactivar" : "Inhabilitar"}" aria-label="${row.active === false ? "Reactivar" : "Inhabilitar"}">${row.active === false ? "&#8635;" : "&#9940;"}</button>
      </td>
    `;
    el.globalGoalsBody.appendChild(tr);
  }
}

function loadGlobalGoalInModal(month = "") {
  const key = String(month || "").trim();
  globalGoalEditingKey = key;
  const target = key ? getGlobalMonthlyGoals({ includeInactive: true }).find((g) => globalGoalCompositeKey(g) === key) : null;
  renderGlobalGoalsEditSelect(key);
  renderGlobalGoalMonthSelect(target?.month || "");
  if (el.globalGoalRole) el.globalGoalRole.value = target ? normalizeUserRole(target.role) : USER_ROLE.SELLER;
  updateGlobalGoalModalControls(key);
  if (el.globalGoalAmount) el.globalGoalAmount.value = target ? String(Number(target.amount || 0)) : "";
}

function openGlobalGoalsModal() {
  if (!el.globalGoalsBackdrop) return;
  renderGlobalGoalsTable();
  loadGlobalGoalInModal("");
  el.globalGoalsBackdrop.hidden = false;
  setTimeout(() => {
    try { el.globalGoalMonth?.focus(); } catch (_) { }
  }, 0);
}

function closeGlobalGoalsModal() {
  if (!el.globalGoalsBackdrop) return;
  el.globalGoalsBackdrop.hidden = true;
  globalGoalEditingKey = "";
  if (el.globalGoalsEditSelect) el.globalGoalsEditSelect.value = "";
  renderGlobalGoalMonthSelect("");
  if (el.globalGoalRole) el.globalGoalRole.value = USER_ROLE.SELLER;
  if (el.globalGoalAmount) el.globalGoalAmount.value = "";
  if (el.globalGoalActive) el.globalGoalActive.checked = true;
  if (el.btnGlobalGoalDisable) {
    el.btnGlobalGoalDisable.disabled = true;
    el.btnGlobalGoalDisable.textContent = "Inhabilitar";
  }
  restoreModuleScreenAfterModal();
}

function saveGlobalGoalFromModal() {
  const month = String(el.globalGoalMonth?.value || "").trim();
  const role = normalizeUserRole(el.globalGoalRole?.value || USER_ROLE.SELLER);
  const amountRaw = String(el.globalGoalAmount?.value || "").trim();
  const amount = Math.max(0, Number(amountRaw || 0));
  if (!/^\d{4}-\d{2}$/.test(month)) return toast("Mes invalido. Usa formato AAAA-MM.");
  if (!Number.isFinite(amount) || amount <= 0) return toast("Monto de meta invalido.");

  const rows = getGlobalMonthlyGoals({ includeInactive: true });
  const editingKey = String(globalGoalEditingKey || "").trim();
  const nextKey = globalGoalCompositeKey({ month, role });
  const duplicate = rows.find((g) => globalGoalCompositeKey(g) === nextKey && globalGoalCompositeKey(g) !== editingKey);
  if (duplicate) return toast("Ya existe una meta para ese mes. Seleccionala para editar.");

  const payload = normalizeGlobalMonthlyGoalRecord({
    month,
    role,
    amount,
    active: el.globalGoalActive?.checked !== false,
  });

  if (editingKey) {
    let changed = false;
    state.globalMonthlyGoals = rows.map((g) => {
      if (globalGoalCompositeKey(g) !== editingKey) return g;
      changed = true;
      return payload;
    });
    if (!changed) state.globalMonthlyGoals.push(payload);
  } else {
    state.globalMonthlyGoals = rows.concat(payload);
  }

  state.globalMonthlyGoals = getGlobalMonthlyGoals({ includeInactive: true });
  persist();
  renderGlobalGoalsTable();
  loadGlobalGoalInModal(nextKey);
  toast(editingKey ? "Meta global actualizada." : "Meta global guardada.");
}

function toggleGlobalGoalActive(targetMonth = "") {
  const month = String(targetMonth || globalGoalEditingKey || "").trim();
  if (!month) return toast("Selecciona una meta global.");
  let changed = false;
  state.globalMonthlyGoals = getGlobalMonthlyGoals({ includeInactive: true }).map((g) => {
    if (globalGoalCompositeKey(g) !== month) return g;
    changed = true;
    return { ...g, active: g.active === false ? true : false };
  });
  if (!changed) return toast("Meta global no encontrada.");
  persist();
  renderGlobalGoalsTable();
  loadGlobalGoalInModal(month);
  const current = getGlobalMonthlyGoals({ includeInactive: true }).find((g) => globalGoalCompositeKey(g) === month);
  toast(current?.active === false ? "Meta global inhabilitada." : "Meta global reactivada.");
}

function renderSalonesEditSelect(selectedName = "") {
  if (!el.salonEditSelect) return;
  const keep = String(selectedName || "").trim();
  const rows = (state.salones || [])
    .slice()
    .sort((a, b) => String(a || "").localeCompare(String(b || ""), "es", { sensitivity: "base" }));
  el.salonEditSelect.innerHTML = `<option value="">Crear nuevo salon</option>`;
  for (const name of rows) {
    const value = String(name || "").trim();
    if (!value) continue;
    const opt = document.createElement("option");
    opt.value = value;
    opt.textContent = `${value}${isSalonDisabled(value) ? " (Inhabilitado)" : ""}`;
    el.salonEditSelect.appendChild(opt);
  }
  el.salonEditSelect.value = keep;
}

function updateSalonModalControls(salonName = "") {
  const name = String(salonName || "").trim();
  if (el.salonActive) el.salonActive.checked = name ? !isSalonDisabled(name) : true;
  if (el.btnSalonDisable) {
    el.btnSalonDisable.disabled = !name;
    el.btnSalonDisable.textContent = name && isSalonDisabled(name) ? "Reactivar" : "Inhabilitar";
  }
}

function renderSalonesTable() {
  if (!el.salonesBody) return;
  const rows = (state.salones || [])
    .slice()
    .sort((a, b) => String(a || "").localeCompare(String(b || ""), "es", { sensitivity: "base" }));
  el.salonesBody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="3">Sin salones registrados.</td>`;
    el.salonesBody.appendChild(tr);
    return;
  }
  for (const name of rows) {
    const label = String(name || "").trim();
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(label)}</td>
      <td>${isSalonDisabled(label) ? "Inhabilitado" : "Activo"}</td>
      <td class="appointmentActions salonActionIcons">
        <button class="apptIconBtn apptEdit" type="button" data-salon-edit="${escapeHtml(label)}" title="Editar" aria-label="Editar">&#9998;</button>
        <button class="apptIconBtn ${isSalonDisabled(label) ? "" : "apptDelete"}" type="button" data-salon-toggle="${escapeHtml(label)}" title="${isSalonDisabled(label) ? "Reactivar" : "Inhabilitar"}" aria-label="${isSalonDisabled(label) ? "Reactivar" : "Inhabilitar"}">${isSalonDisabled(label) ? "&#8635;" : "&#9940;"}</button>
      </td>
    `;
    el.salonesBody.appendChild(tr);
  }
}

function loadSalonInModal(salonName = "") {
  const name = String(salonName || "").trim();
  editingSalonName = name;
  renderSalonesEditSelect(name);
  updateSalonModalControls(name);
  if (el.salonNameInput) el.salonNameInput.value = name;
}

function openSalonesModal() {
  if (!el.salonesBackdrop) return;
  renderSalonesTable();
  loadSalonInModal("");
  el.salonesBackdrop.hidden = false;
  setTimeout(() => {
    try { el.salonNameInput?.focus(); } catch (_) { }
  }, 0);
}

function closeSalonesModal() {
  if (!el.salonesBackdrop) return;
  el.salonesBackdrop.hidden = true;
  editingSalonName = "";
  if (el.salonEditSelect) el.salonEditSelect.value = "";
  if (el.salonNameInput) el.salonNameInput.value = "";
  if (el.salonActive) el.salonActive.checked = true;
  if (el.btnSalonDisable) {
    el.btnSalonDisable.disabled = true;
    el.btnSalonDisable.textContent = "Inhabilitar";
  }
  restoreModuleScreenAfterModal();
}

function saveSalonFromModal() {
  const nextName = String(el.salonNameInput?.value || "").trim();
  if (!nextName) return toast("Nombre del salon es obligatorio.");
  const currentName = String(editingSalonName || "").trim();
  const exists = (state.salones || []).some((s) =>
    String(s || "").trim().toLowerCase() === nextName.toLowerCase() && String(s || "").trim() !== currentName
  );
  if (exists) return toast("Ya existe un salon con ese nombre.");

  if (currentName) {
    state.salones = (state.salones || []).map((s) => (String(s || "").trim() === currentName ? nextName : s));
    if (isSalonDisabled(currentName)) {
      state.disabledSalones = Array.from(new Set([...(state.disabledSalones || []), nextName]));
      enableSalon(currentName);
    }
  } else {
    state.salones = Array.from(new Set([...(state.salones || []), nextName]));
  }

  if (el.salonActive?.checked === false) {
    state.disabledSalones = Array.from(new Set([...(state.disabledSalones || []), nextName]));
  } else {
    enableSalon(nextName);
  }

  state.salones.sort((a, b) => String(a || "").localeCompare(String(b || ""), "es", { sensitivity: "base" }));
  renderRoomSelects();
  persist();
  renderSalonesTable();
  loadSalonInModal(currentName ? nextName : "");
  if (!currentName) {
    setTimeout(() => {
      try { el.salonNameInput?.focus(); } catch (_) { }
    }, 0);
  }
  toast(currentName ? "Salon actualizado." : "Salon agregado.");
}

function toggleSalonActive(targetSalon = "") {
  const name = String(targetSalon || editingSalonName || "").trim();
  if (!name) return toast("Selecciona un salon.");
  if (isSalonDisabled(name)) enableSalon(name);
  else state.disabledSalones = Array.from(new Set([...(state.disabledSalones || []), name]));
  renderRoomSelects();
  persist();
  renderSalonesTable();
  loadSalonInModal(name);
  toast(isSalonDisabled(name) ? "Salon inhabilitado." : "Salon reactivado.");
}

async function manageGlobalGoalsFromQuickMenu() {
  const action = await promptSelectRequired({
    title: "Metas globales: accion",
    label: "Que deseas hacer",
    options: [
      { value: "add", label: "Agregar meta mensual" },
      { value: "edit", label: "Editar meta mensual" },
      { value: "delete", label: "Eliminar meta mensual" },
    ],
  });
  if (!action) return;

  const current = getGlobalMonthlyGoals();
  if (action === "add") {
    const month = await promptTextRequired({
      title: "Meta global mensual",
      label: "Mes (AAAA-MM)",
      placeholder: "2026-03",
    });
    if (!month || !/^\d{4}-\d{2}$/.test(month)) return toast("Mes invalido. Usa formato AAAA-MM.");
    const role = await promptSelectRequired({
      title: "Rol para meta global",
      label: "Selecciona rol",
      options: REPORTABLE_USER_ROLES.map((item) => ({ value: item, label: userRoleLabel(item) })),
    });
    if (!role) return;
    const amountRaw = await promptTextRequired({
      title: "Monto meta global",
      label: `Mes ${month} | ${userRoleLabel(role)}`,
      placeholder: "Ej: 250000",
    });
    const amount = Math.max(0, Number(amountRaw || 0));
    if (!Number.isFinite(amount) || amount <= 0) return toast("Monto de meta invalido.");
    const nextKey = globalGoalCompositeKey({ month, role });
    const next = current.filter((g) => globalGoalCompositeKey(g) !== nextKey);
    next.push({ month, role, amount });
    next.sort((a, b) => a.month.localeCompare(b.month) || userRoleLabel(a.role).localeCompare(userRoleLabel(b.role), "es", { sensitivity: "base" }));
    state.globalMonthlyGoals = next;
    persist();
    return toast("Meta global mensual guardada.");
  }

  if (!current.length) return toast("No hay metas globales registradas.");
  const selectedMonth = await promptSelectRequired({
    title: action === "edit" ? "Editar meta global" : "Eliminar meta global",
    options: current.map((g) => ({
      value: globalGoalCompositeKey(g),
      label: `${g.month} | ${userRoleLabel(g.role)} - Q ${g.amount.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
    })),
  });
  if (!selectedMonth) return;

  if (action === "edit") {
    const target = current.find((g) => globalGoalCompositeKey(g) === selectedMonth);
    if (!target) return;
    const amountRaw = await promptTextRequired({
      title: "Nuevo monto meta global",
      label: `Mes ${target.month} | ${userRoleLabel(target.role)}`,
      placeholder: String(target.amount),
    });
    const amount = Math.max(0, Number(amountRaw || 0));
    if (!Number.isFinite(amount) || amount <= 0) return toast("Monto de meta invalido.");
    state.globalMonthlyGoals = current.map((g) => (globalGoalCompositeKey(g) === selectedMonth ? { month: target.month, role: target.role, amount } : g));
    persist();
    return toast("Meta global actualizada.");
  }

  const ok = await modernConfirm({
    title: "Eliminar meta global",
    message: `Esta seguro de eliminar la meta global seleccionada?`,
    confirmText: "Si, eliminar",
    cancelText: "No",
  });
  if (!ok) return;
  state.globalMonthlyGoals = current.filter((g) => globalGoalCompositeKey(g) !== selectedMonth);
  persist();
  toast("Meta global eliminada.");
}

async function readMenuCatalog(kind, extraQuery = "") {
  const q = String(extraQuery || "").trim();
  const endpoint = buildApiUrlFromStateUrl(activeApiStateUrl, `menu-catalog/${encodeURIComponent(kind)}${q ? `?${q}` : ""}`);
  const res = await fetch(endpoint, { cache: "no-store" });
  if (!res.ok) throw new Error(`menu_catalog_read_${kind}`);
  const payload = await res.json();
  return Array.isArray(payload?.items) ? payload.items : [];
}

async function createMenuCatalog(kind, body) {
  const endpoint = buildApiUrlFromStateUrl(activeApiStateUrl, `menu-catalog/${encodeURIComponent(kind)}`);
  const res = await fetch(endpoint, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body || {}),
  });
  if (!res.ok) {
    let detail = "";
    try {
      const payload = await res.json();
      detail = String(payload?.detail || payload?.message || "").trim();
    } catch (_) { }
    throw new Error(detail || `menu_catalog_create_${kind}`);
  }
}

async function updateMenuCatalog(kind, id, body) {
  const endpoint = buildApiUrlFromStateUrl(activeApiStateUrl, `menu-catalog/${encodeURIComponent(kind)}/${encodeURIComponent(String(id || ""))}`);
  const res = await fetch(endpoint, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(body || {}),
  });
  if (!res.ok) {
    let detail = "";
    try {
      const payload = await res.json();
      detail = String(payload?.detail || payload?.message || "").trim();
    } catch (_) { }
    throw new Error(detail || `menu_catalog_update_${kind}`);
  }
}

async function readMenuSuggestions({ platoId, preparacionId }) {
  const q = `plato_id=${encodeURIComponent(String(platoId || ""))}&preparacion_id=${encodeURIComponent(String(preparacionId || ""))}`;
  const endpoint = buildApiUrlFromStateUrl(activeApiStateUrl, `menu-suggestions?${q}`);
  const res = await fetch(endpoint);
  if (!res.ok) throw new Error("menu_suggestions_read_failed");
  return res.json();
}

function normalizeMmsSuggestionLinks(raw) {
  const links = raw && typeof raw === "object" ? raw : {};
  const normArray = (value) => (Array.isArray(value) ? value : [])
    .map((x) => Number(x))
    .filter((n) => Number.isFinite(n) && n > 0);
  return {
    salsaIds: normArray(links.salsaIds),
    postreIds: normArray(links.postreIds),
    guarnicionIds: normArray(links.guarnicionIds),
    bebidaIds: normArray(links.bebidaIds),
    montajeTipoIds: normArray(links.montajeTipoIds),
    montajeAdicionalIds: normArray(links.montajeAdicionalIds),
  };
}

async function readMenuSuggestionsCached(platoId, preparacionId) {
  ensureMmsCatalogDefaults();
  const pid = Number(platoId || 0);
  const prepId = Number(preparacionId || 0);
  if (!Number.isFinite(pid) || pid <= 0 || !Number.isFinite(prepId) || prepId <= 0) {
    return normalizeMmsSuggestionLinks(null);
  }
  const cache = menuMontajeSelectableCatalogCache;
  const key = `${pid}:${prepId}`;
  if (cache.suggestionsByPair.has(key)) {
    return cache.suggestionsByPair.get(key);
  }
  if (cache.suggestionsInFlightByPair.has(key)) {
    return cache.suggestionsInFlightByPair.get(key);
  }
  const pending = readMenuSuggestions({ platoId: pid, preparacionId: prepId })
    .then((payload) => {
      const normalized = normalizeMmsSuggestionLinks(payload);
      cache.suggestionsByPair.set(key, normalized);
      cache.suggestionsInFlightByPair.delete(key);
      return normalized;
    })
    .catch((err) => {
      cache.suggestionsInFlightByPair.delete(key);
      throw err;
    });
  cache.suggestionsInFlightByPair.set(key, pending);
  return pending;
}

async function saveMenuSuggestions(payload) {
  const endpoint = buildApiUrlFromStateUrl(activeApiStateUrl, "menu-suggestions");
  const res = await fetch(endpoint, {
    method: "PUT",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload || {}),
  });
  if (!res.ok) {
    let detail = "";
    try {
      const body = await res.json();
      detail = String(body?.detail || body?.message || "").trim();
    } catch (_) { }
    throw new Error(detail || "menu_suggestions_save_failed");
  }
}

function renderMenuSuggestionCheckboxList(container, items, selectedIds) {
  if (!container) return;
  container.innerHTML = "";
  const rows = Array.isArray(items) ? items.filter((x) => x && x.activo !== false) : [];
  if (!rows.length) {
    container.innerHTML = `<div class="menuSuggestEmpty">Sin datos en catalogo.</div>`;
    return;
  }
  const selectedSet = new Set((Array.isArray(selectedIds) ? selectedIds : []).map((x) => String(x)));
  for (const item of rows) {
    const id = String(item.id || "").trim();
    if (!id) continue;
    const isChecked = selectedSet.has(id);
    const row = document.createElement("label");
    row.className = "menuSuggestRow";
    row.dataset.mmSuggestId = id;
    row.draggable = isChecked;
    row.classList.toggle("isChecked", isChecked);
    row.innerHTML = `
      <span class="menuSuggestDragWrap">
        <span class="menuSuggestDrag" title="Arrastra para priorizar">&#9776;</span>
      </span>
      <span class="menuSuggestCheckWrap">
        <input type="checkbox" value="${escapeHtml(id)}" ${isChecked ? "checked" : ""} />
        <span class="menuSuggestCheckVisual" aria-hidden="true"></span>
      </span>
      <span class="menuSuggestMeta">
        <span class="menuSuggestName">${escapeHtml(String(item.nombre || "").trim())}</span>
        <span class="menuSuggestHint">${isChecked ? "Sugerencia activa y ordenable." : "Marca para sugerir esta opcion."}</span>
      </span>
      <span class="menuSuggestState">${isChecked ? "Activo" : "Opcional"}</span>
    `;
    container.appendChild(row);
  }
}

function selectedIdsFromChecklist(container) {
  if (!container) return [];
  const out = [];
  const checks = container.querySelectorAll("input[type='checkbox']:checked");
  for (const node of checks) {
    const n = Number(node.value);
    if (Number.isFinite(n) && n > 0) out.push(n);
  }
  return out;
}

function setMenuSuggestRowDraggableByCheckbox(row) {
  if (!row) return;
  const checkbox = row.querySelector("input[type='checkbox']");
  const isChecked = !!checkbox?.checked;
  row.draggable = isChecked;
  row.classList.toggle("isChecked", isChecked);
}

function bindMenuSuggestDnD(container) {
  if (!container) return;

  container.addEventListener("change", (e) => {
    const checkbox = e.target.closest("input[type='checkbox']");
    if (!checkbox) return;
    const row = checkbox.closest(".menuSuggestRow");
    setMenuSuggestRowDraggableByCheckbox(row);
  });

  container.addEventListener("dragstart", (e) => {
    const row = e.target.closest(".menuSuggestRow");
    if (!row || !row.draggable) return;
    menuSuggestionDraggingRow = row;
    row.classList.add("isDragging");
    if (e.dataTransfer) {
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", row.dataset.mmSuggestId || "");
    }
  });

  container.addEventListener("dragover", (e) => {
    if (!menuSuggestionDraggingRow) return;
    e.preventDefault();
    const over = e.target.closest(".menuSuggestRow");
    if (!over || over === menuSuggestionDraggingRow || over.parentElement !== container) return;
    const rect = over.getBoundingClientRect();
    const placeAfter = e.clientY > (rect.top + rect.height / 2);
    container.insertBefore(menuSuggestionDraggingRow, placeAfter ? over.nextSibling : over);
  });

  container.addEventListener("drop", (e) => {
    if (!menuSuggestionDraggingRow) return;
    e.preventDefault();
  });

  container.addEventListener("dragend", () => {
    if (menuSuggestionDraggingRow) {
      menuSuggestionDraggingRow.classList.remove("isDragging");
    }
    menuSuggestionDraggingRow = null;
  });
}

function formatPlatoCatalogLabel(item) {
  const name = String(item?.nombre || "").trim() || "(sin nombre)";
  const tipo = String(item?.tipo_plato || "NORMAL").trim();
  const sinProteina = item?.es_sin_proteina === true || Number(item?.es_sin_proteina) !== 0;
  const tags = [];
  if (tipo && tipo !== "NORMAL") tags.push(tipo);
  if (sinProteina) tags.push("SIN PROTEINA");
  return tags.length ? `${name} [${tags.join(" | ")}]` : name;
}

function resetMenuCatalogManagerForm() {
  menuCatalogManagerEditingId = "";
  if (el.menuCatalogName) el.menuCatalogName.value = "";
  if (el.menuCatalogDishType) el.menuCatalogDishType.value = "NORMAL";
  if (el.menuCatalogNoProtein) el.menuCatalogNoProtein.checked = false;
}

function syncMenuCatalogManagerFormByKind() {
  const kind = String(el.menuCatalogKind?.value || menuCatalogManagerKind || "plato_fuerte");
  menuCatalogManagerKind = kind;
  const isPlato = kind === "plato_fuerte";
  const isPrep = kind === "preparacion";
  if (el.menuCatalogDishTypeWrap) el.menuCatalogDishTypeWrap.hidden = !isPlato;
  if (el.menuCatalogNoProteinWrap) el.menuCatalogNoProteinWrap.hidden = !isPlato;
  if (el.menuCatalogProteinWrap) el.menuCatalogProteinWrap.hidden = !isPrep;
}

async function loadMenuCatalogProteinOptionsForManager() {
  if (!el.menuCatalogProtein) return [];
  const platos = await readMenuCatalog("plato_fuerte");
  el.menuCatalogProtein.innerHTML = "";
  for (const p of platos.filter((x) => x && x.activo !== false)) {
    const opt = document.createElement("option");
    opt.value = String(p.id);
    opt.textContent = formatPlatoCatalogLabel(p);
    el.menuCatalogProtein.appendChild(opt);
  }
  if (!el.menuCatalogProtein.options.length) {
    el.menuCatalogProtein.innerHTML = `<option value="">Sin proteinas activas</option>`;
  }
  return platos;
}

function renderMenuCatalogManagerRows(kind, rows, proteins = []) {
  if (!el.menuCatalogBody) return;
  el.menuCatalogBody.innerHTML = "";
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="4">Sin registros.</td>`;
    el.menuCatalogBody.appendChild(tr);
    return;
  }
  const proteinById = new Map((Array.isArray(proteins) ? proteins : []).map((p) => [String(p.id), p]));
  for (const item of list) {
    const tr = document.createElement("tr");
    const id = String(item.id || "");
    const isActive = item.activo !== false;
    let detail = "-";
    if (kind === "plato_fuerte") {
      const tipo = String(item.tipo_plato || "NORMAL");
      const sp = item.es_sin_proteina ? " | SIN PROTEINA" : "";
      detail = `${tipo}${sp}`;
    } else if (kind === "preparacion") {
      const protein = proteinById.get(String(item.id_plato_fuerte || ""));
      detail = `Proteina: ${protein ? formatPlatoCatalogLabel(protein) : String(item.id_plato_fuerte || "-")}`;
    } else if (kind === "montaje_adicional") {
      detail = String(item.tipo || "-");
    }
    tr.innerHTML = `
      <td>${escapeHtml(String(item.nombre || "-"))}</td>
      <td>${escapeHtml(detail)}</td>
      <td>${isActive ? "Activo" : "Inhabilitado"}</td>
      <td>
        <div class="appointmentActions">
          <button type="button" class="apptIconBtn apptEdit" data-mmcat-action="edit" data-mmcat-id="${escapeHtml(id)}" title="Editar registro" aria-label="Editar registro">&#9998;</button>
          <button type="button" class="apptIconBtn ${isActive ? "apptDelete" : ""}" data-mmcat-action="toggle" data-mmcat-id="${escapeHtml(id)}" title="${isActive ? "Inhabilitar registro" : "Reactivar registro"}" aria-label="${isActive ? "Inhabilitar registro" : "Reactivar registro"}">${isActive ? "&#9940;" : "&#8635;"}</button>
        </div>
      </td>
    `;
    el.menuCatalogBody.appendChild(tr);
  }
}

async function refreshMenuCatalogManagerRows() {
  const kind = String(el.menuCatalogKind?.value || menuCatalogManagerKind || "plato_fuerte");
  menuCatalogManagerKind = kind;
  const proteins = await loadMenuCatalogProteinOptionsForManager();
  let rows;
  if (kind === "preparacion") {
    const proteinId = Number(el.menuCatalogProtein?.value || 0);
    rows = await readMenuCatalog("preparacion", `plato_id=${encodeURIComponent(String(proteinId || ""))}`);
  } else {
    rows = await readMenuCatalog(kind);
  }
  menuCatalogManagerRows = Array.isArray(rows) ? rows : [];
  renderMenuCatalogManagerRows(kind, menuCatalogManagerRows, proteins);
}

async function saveMenuCatalogManagerRecord() {
  const kind = String(el.menuCatalogKind?.value || menuCatalogManagerKind || "plato_fuerte");
  const name = String(el.menuCatalogName?.value || "").trim();
  if (!name) return toast("Nombre requerido.");
  const editingId = String(menuCatalogManagerEditingId || "").trim();

  if (kind === "preparacion") {
    const proteinId = Number(el.menuCatalogProtein?.value || 0);
    if (!Number.isFinite(proteinId) || proteinId <= 0) return toast("Selecciona proteina base.");
    if (editingId) {
      await updateMenuCatalog("preparacion", editingId, {
        nombre: name,
        id_plato_fuerte: proteinId,
        activo: true,
      });
    } else {
      await createMenuCatalog("preparacion", {
        nombre: name,
        id_plato_fuerte: proteinId,
      });
    }
  } else if (kind === "plato_fuerte") {
    const tipoPlato = String(el.menuCatalogDishType?.value || "NORMAL");
    const sinProteina = !!el.menuCatalogNoProtein?.checked;
    if (editingId) {
      await updateMenuCatalog(kind, editingId, {
        nombre: name,
        tipo_plato: tipoPlato,
        es_sin_proteina: sinProteina ? 1 : 0,
        activo: true,
      });
    } else {
      await createMenuCatalog(kind, {
        nombre: name,
        tipo_plato: tipoPlato,
        es_sin_proteina: sinProteina ? 1 : 0,
      });
    }
  } else {
    if (editingId) {
      await updateMenuCatalog(kind, editingId, { nombre: name, activo: true });
    } else {
      await createMenuCatalog(kind, { nombre: name });
    }
  }

  resetMenuCatalogManagerForm();
  await refreshMenuCatalogManagerRows();
}

async function openMenuCatalogManagerModal(initialKind = "plato_fuerte") {
  if (!el.menuCatalogBackdrop || !el.menuCatalogKind) return;
  menuCatalogManagerEditingId = "";
  menuCatalogManagerKind = String(initialKind || "plato_fuerte");
  el.menuCatalogKind.value = menuCatalogManagerKind;
  resetMenuCatalogManagerForm();
  syncMenuCatalogManagerFormByKind();
  await refreshMenuCatalogManagerRows();
  el.menuCatalogBackdrop.hidden = false;
}

function closeMenuCatalogManagerModal() {
  if (!el.menuCatalogBackdrop) return;
  el.menuCatalogBackdrop.hidden = true;
  resetMenuCatalogManagerForm();
}

async function refreshMenuSuggestionsModalData() {
  if (!el.menuSuggestionsProtein || !el.menuSuggestionsPreparation) return;
  const platoId = Number(el.menuSuggestionsProtein.value || 0);
  const preparacionId = Number(el.menuSuggestionsPreparation.value || 0);
  if (!Number.isFinite(platoId) || platoId <= 0 || !Number.isFinite(preparacionId) || preparacionId <= 0) {
    renderMenuSuggestionCheckboxList(el.menuSuggestionsSalsas, [], []);
    renderMenuSuggestionCheckboxList(el.menuSuggestionsPostres, [], []);
    renderMenuSuggestionCheckboxList(el.menuSuggestionsGuarniciones, [], []);
    renderMenuSuggestionCheckboxList(el.menuSuggestionsBebidas, [], []);
    renderMenuSuggestionCheckboxList(el.menuSuggestionsMontajeTipos, [], []);
    renderMenuSuggestionCheckboxList(el.menuSuggestionsMontajeAdicionales, [], []);
    return;
  }

  const [salsas, postres, guarniciones, bebidas, montajeTipos, montajeAdicionales, links] = await Promise.all([
    readMenuCatalog("salsa"),
    readMenuCatalog("postre"),
    readMenuCatalog("guarnicion"),
    readMenuCatalog("bebida"),
    readMenuCatalog("montaje_tipo"),
    readMenuCatalog("montaje_adicional"),
    readMenuSuggestions({ platoId, preparacionId }),
  ]);
  renderMenuSuggestionCheckboxList(el.menuSuggestionsSalsas, salsas, links?.salsaIds || []);
  renderMenuSuggestionCheckboxList(el.menuSuggestionsPostres, postres, links?.postreIds || []);
  renderMenuSuggestionCheckboxList(el.menuSuggestionsGuarniciones, guarniciones, links?.guarnicionIds || []);
  renderMenuSuggestionCheckboxList(el.menuSuggestionsBebidas, bebidas, links?.bebidaIds || []);
  renderMenuSuggestionCheckboxList(el.menuSuggestionsMontajeTipos, montajeTipos, links?.montajeTipoIds || []);
  renderMenuSuggestionCheckboxList(el.menuSuggestionsMontajeAdicionales, montajeAdicionales, links?.montajeAdicionalIds || []);
}

async function openMenuSuggestionsModal() {
  if (!el.menuSuggestionsBackdrop || !el.menuSuggestionsProtein || !el.menuSuggestionsPreparation) return;
  const platos = await readMenuCatalog("plato_fuerte");
  el.menuSuggestionsProtein.innerHTML = "";
  for (const p of platos.filter((x) => x && x.activo !== false)) {
    const opt = document.createElement("option");
    opt.value = String(p.id);
    opt.textContent = formatPlatoCatalogLabel(p);
    el.menuSuggestionsProtein.appendChild(opt);
  }
  if (!el.menuSuggestionsProtein.options.length) {
    el.menuSuggestionsProtein.innerHTML = `<option value="">Sin proteinas registradas</option>`;
    el.menuSuggestionsPreparation.innerHTML = `<option value="">Sin preparaciones</option>`;
    renderMenuSuggestionCheckboxList(el.menuSuggestionsSalsas, [], []);
    renderMenuSuggestionCheckboxList(el.menuSuggestionsPostres, [], []);
    renderMenuSuggestionCheckboxList(el.menuSuggestionsGuarniciones, [], []);
    el.menuSuggestionsBackdrop.hidden = false;
    document.body.classList.add("menuSuggestionsModalOpen");
    return;
  }

  const proteinId = Number(el.menuSuggestionsProtein.value || el.menuSuggestionsProtein.options[0].value || 0);
  const preps = await readMenuCatalog("preparacion", `plato_id=${encodeURIComponent(String(proteinId || ""))}`);
  el.menuSuggestionsPreparation.innerHTML = "";
  for (const p of preps.filter((x) => x && x.activo !== false)) {
    const opt = document.createElement("option");
    opt.value = String(p.id);
    opt.textContent = String(p.nombre || "");
    el.menuSuggestionsPreparation.appendChild(opt);
  }
  if (!el.menuSuggestionsPreparation.options.length) {
    el.menuSuggestionsPreparation.innerHTML = `<option value="">Sin preparaciones para esta proteina</option>`;
  }

  await refreshMenuSuggestionsModalData();
  el.menuSuggestionsBackdrop.hidden = false;
  document.body.classList.add("menuSuggestionsModalOpen");
}

function closeMenuSuggestionsModal() {
  if (!el.menuSuggestionsBackdrop) return;
  el.menuSuggestionsBackdrop.hidden = true;
  document.body.classList.remove("menuSuggestionsModalOpen");
}

async function manageMenuMontajeCatalogFromQuickMenu() {
  const kind = await promptSelectRequired({
    title: "Catalogo Menu & Montaje",
    options: [
      { value: "plato_fuerte", label: "Proteina / Plato fuerte" },
      { value: "preparacion", label: "Preparacion (ej. empanizado)" },
      { value: "salsa", label: "Salsa o aderezo" },
      { value: "guarnicion", label: "Guarnicion" },
      { value: "postre", label: "Postre" },
      { value: "comentario", label: "Comentario adicional" },
      { value: "montaje_tipo", label: "Tipo de montaje" },
      { value: "montaje_adicional", label: "Adicional de montaje" },
    ],
  });
  if (!kind) return;

  const action = await promptCrudAction("Catalogo");
  if (!action) return;

  const titleByKind = {
    plato_fuerte: "Nueva proteina / plato fuerte",
    salsa: "Nueva salsa o aderezo",
    guarnicion: "Nueva guarnicion",
    postre: "Nuevo postre",
    comentario: "Nuevo comentario adicional",
    montaje_tipo: "Nuevo tipo de montaje",
    montaje_adicional: "Nuevo adicional de montaje",
  };

  if (kind === "preparacion") {
    const platos = await readMenuCatalog("plato_fuerte");
    if (!platos.length) return toast("Primero agrega una proteina/plato fuerte.");
    const platoId = await promptSelectRequired({
      title: "Proteina base",
      options: platos.map((p) => ({ value: String(p.id), label: formatPlatoCatalogLabel(p) })),
    });
    if (!platoId) return;

    const preparaciones = await readMenuCatalog("preparacion", `plato_id=${encodeURIComponent(String(platoId))}`);
    if (action === "add") {
      const nombrePrep = await promptTextRequired({
        title: "Nueva preparacion",
        label: "Nombre de la preparacion",
        placeholder: "Ej: A la parrilla",
      });
      if (!nombrePrep) return;
      await createMenuCatalog("preparacion", {
        nombre: nombrePrep,
        id_plato_fuerte: Number(platoId),
      });
      return toast("Preparacion de menu agregada.");
    }

    if (!preparaciones.length) return toast("No hay preparaciones registradas para esa proteina.");
    const selectedPrepId = await promptSelectRequired({
      title: action === "edit" ? "Editar preparacion" : "Inhabilitar preparacion",
      options: preparaciones.map((p) => ({
        value: String(p.id),
        label: `${String(p.nombre || "")}${p.activo === false ? " (Inhabilitada)" : ""}`,
      })),
    });
    if (!selectedPrepId) return;

    if (action === "edit") {
      const target = preparaciones.find((p) => String(p.id) === String(selectedPrepId));
      const nextName = await promptTextRequired({
        title: "Nuevo nombre de preparacion",
        label: "Nombre",
        placeholder: String(target?.nombre || ""),
      });
      if (!nextName) return;
      await updateMenuCatalog("preparacion", selectedPrepId, {
        nombre: nextName,
        activo: true,
        id_plato_fuerte: Number(platoId),
      });
      return toast("Preparacion actualizada.");
    }

    await updateMenuCatalog("preparacion", selectedPrepId, { activo: false });
    return toast("Preparacion inhabilitada.");
  }

  const items = await readMenuCatalog(kind);
  if (action === "add") {
    const nombre = await promptTextRequired({
      title: titleByKind[kind] || "Nuevo registro",
      label: "Nombre",
      placeholder: "Escribe el nombre",
    });
    if (!nombre) return;
    if (kind === "plato_fuerte") {
      const tipoPlato = await promptSelectRequired({
        title: "Tipo de plato",
        options: [
          { value: "NORMAL", label: "Normal" },
          { value: "VEGETARIANO", label: "Vegetariano" },
          { value: "VEGANO", label: "Vegano" },
        ],
      });
      if (!tipoPlato) return;
      const sinProteina = await promptSelectRequired({
        title: "Este plato puede ser sin proteina?",
        options: [
          { value: "0", label: "No" },
          { value: "1", label: "Si" },
        ],
      });
      if (sinProteina === null || sinProteina === undefined) return;
      await createMenuCatalog(kind, {
        nombre,
        tipo_plato: tipoPlato,
        es_sin_proteina: Number(sinProteina) ? 1 : 0,
      });
    } else {
      await createMenuCatalog(kind, { nombre });
    }
    return toast("Catalogo de Menu & Montaje actualizado.");
  }

  if (!items.length) return toast("No hay registros en ese catalogo.");
  const selectedId = await promptSelectRequired({
    title: action === "edit" ? "Editar registro" : "Inhabilitar registro",
    options: items.map((it) => ({
      value: String(it.id),
      label: `${kind === "plato_fuerte" ? formatPlatoCatalogLabel(it) : String(it.nombre || "")}${it.activo === false ? " (Inhabilitado)" : ""}`,
    })),
  });
  if (!selectedId) return;

  if (action === "edit") {
    const target = items.find((it) => String(it.id) === String(selectedId));
    const nextName = await promptTextRequired({
      title: "Nuevo nombre",
      label: "Nombre",
      placeholder: String(target?.nombre || ""),
    });
    if (!nextName) return;
    if (kind === "plato_fuerte") {
      const tipoPlato = await promptSelectRequired({
        title: "Tipo de plato",
        options: [
          { value: "NORMAL", label: "Normal" },
          { value: "VEGETARIANO", label: "Vegetariano" },
          { value: "VEGANO", label: "Vegano" },
        ],
      });
      if (!tipoPlato) return;
      const sinProteina = await promptSelectRequired({
        title: "Este plato puede ser sin proteina?",
        options: [
          { value: "0", label: "No" },
          { value: "1", label: "Si" },
        ],
      });
      if (sinProteina === null || sinProteina === undefined) return;
      await updateMenuCatalog(kind, selectedId, {
        nombre: nextName,
        activo: true,
        tipo_plato: tipoPlato,
        es_sin_proteina: Number(sinProteina) ? 1 : 0,
      });
    } else {
      await updateMenuCatalog(kind, selectedId, { nombre: nextName, activo: true });
    }
    return toast("Registro actualizado.");
  }

  await updateMenuCatalog(kind, selectedId, { activo: false });
  toast("Registro inhabilitado.");
}

function readImageFileAsDataUrl(file) {
  return new Promise((resolve) => {
    if (!file) return resolve("");
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => resolve("");
    reader.readAsDataURL(file);
  });
}

function readBlobAsDataUrl(blob) {
  return new Promise((resolve) => {
    if (!blob) return resolve("");
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => resolve("");
    reader.readAsDataURL(blob);
  });
}

async function resolveTemplateImageAssetToDataUrl(value) {
  const raw = String(value || "").trim();
  if (!raw) return "";
  if (isImageDataUrl(raw)) return raw;
  try {
    const res = await fetch(raw, { cache: "no-store" });
    if (!res.ok) return "";
    const blob = await res.blob();
    return await readBlobAsDataUrl(blob);
  } catch (_) {
    return "";
  }
}
function isTemplateSignatureToken(token) {
  const t = String(token || "").toLowerCase().trim();
  return t === "{{cliente.firma}}" || t === "{{vendedor.firma}}" || t.includes(".firma");
}

function normalizeTemplateSignatureDefaults(rawDefaults, positionedFields = []) {
  const firstSignature = (positionedFields || []).find((p) => p?.isSignature === true || isTemplateSignatureToken(p?.token));
  const fallbackW = clamp(Number(firstSignature?.w || TEMPLATE_SIGNATURE_FALLBACK_W_PCT), TEMPLATE_SIGNATURE_MIN_W_PCT, TEMPLATE_SIGNATURE_MAX_W_PCT);
  const fallbackH = clamp(Number(firstSignature?.h || TEMPLATE_SIGNATURE_FALLBACK_H_PCT), TEMPLATE_SIGNATURE_MIN_H_PCT, TEMPLATE_SIGNATURE_MAX_H_PCT);
  return {
    w: clamp(Number(rawDefaults?.w || fallbackW), TEMPLATE_SIGNATURE_MIN_W_PCT, TEMPLATE_SIGNATURE_MAX_W_PCT),
    h: clamp(Number(rawDefaults?.h || fallbackH), TEMPLATE_SIGNATURE_MIN_H_PCT, TEMPLATE_SIGNATURE_MAX_H_PCT),
  };
}

function loadImageFromDataUrl(dataUrl) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => resolve(img);
    img.onerror = reject;
    img.src = dataUrl;
  });
}

async function analyzeSignatureDataUrl(dataUrl) {
  const safeData = String(dataUrl || "").trim();
  if (!isImageDataUrl(safeData)) return null;
  if (signatureImageAnalysisCache.has(safeData)) return signatureImageAnalysisCache.get(safeData);
  try {
    const img = await loadImageFromDataUrl(safeData);
    const srcW = Math.max(1, Number(img.naturalWidth || img.width || 1));
    const srcH = Math.max(1, Number(img.naturalHeight || img.height || 1));
    const scale = Math.min(1, 920 / Math.max(srcW, srcH));
    const w = Math.max(1, Math.round(srcW * scale));
    const h = Math.max(1, Math.round(srcH * scale));
    const canvas = document.createElement("canvas");
    canvas.width = w;
    canvas.height = h;
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    if (!ctx) return null;
    ctx.drawImage(img, 0, 0, w, h);
    const pixels = ctx.getImageData(0, 0, w, h).data;
    let minX = w;
    let minY = h;
    let maxX = -1;
    let maxY = -1;
    for (let i = 0; i < pixels.length; i += 4) {
      const r = pixels[i];
      const g = pixels[i + 1];
      const b = pixels[i + 2];
      const a = pixels[i + 3];
      const luma = 0.2126 * r + 0.7152 * g + 0.0722 * b;
      const isStroke = a > 22 && luma < 248;
      if (!isStroke) continue;
      const p = i / 4;
      const x = p % w;
      const y = Math.floor(p / w);
      if (x < minX) minX = x;
      if (y < minY) minY = y;
      if (x > maxX) maxX = x;
      if (y > maxY) maxY = y;
    }
    const hasStroke = maxX >= minX && maxY >= minY;
    const contentW = hasStroke ? (maxX - minX + 1) : 0;
    const contentH = hasStroke ? (maxY - minY + 1) : 0;
    const contentArea = contentW * contentH;
    const areaPct = hasStroke ? (contentArea / Math.max(1, w * h)) * 100 : 0;
    const contentWPct = hasStroke ? (contentW / Math.max(1, w)) * 100 : 0;
    const contentHPct = hasStroke ? (contentH / Math.max(1, h)) * 100 : 0;
    const whitespaceHeavy = !hasStroke || areaPct < 22 || contentWPct < 58 || contentHPct < 18;
    const result = {
      width: srcW,
      height: srcH,
      hasStroke,
      contentAreaPct: areaPct,
      contentWPct,
      contentHPct,
      whitespaceHeavy,
      recommendedAspectRatio: hasStroke ? clamp(contentW / Math.max(1, contentH), 1.8, 10) : 4,
    };
    signatureImageAnalysisCache.set(safeData, result);
    return result;
  } catch (_) {
    return null;
  }
}

async function normalizeSignatureDataUrlForContract(dataUrl) {
  const safeData = String(dataUrl || "").trim();
  if (!isImageDataUrl(safeData)) return "";
  try {
    const img = await loadImageFromDataUrl(safeData);
    const srcW = Math.max(1, Number(img.naturalWidth || img.width || 1));
    const srcH = Math.max(1, Number(img.naturalHeight || img.height || 1));
    const scanMax = 1000;
    const scanScale = Math.min(1, scanMax / Math.max(srcW, srcH));
    const scanW = Math.max(1, Math.round(srcW * scanScale));
    const scanH = Math.max(1, Math.round(srcH * scanScale));
    const scanCanvas = document.createElement("canvas");
    scanCanvas.width = scanW;
    scanCanvas.height = scanH;
    const scanCtx = scanCanvas.getContext("2d", { willReadFrequently: true });
    if (!scanCtx) return safeData;
    scanCtx.drawImage(img, 0, 0, scanW, scanH);
    const pixels = scanCtx.getImageData(0, 0, scanW, scanH).data;

    let minX = scanW;
    let minY = scanH;
    let maxX = -1;
    let maxY = -1;
    const binary = new Uint8Array(scanW * scanH);
    for (let i = 0; i < pixels.length; i += 4) {
      const r = pixels[i];
      const g = pixels[i + 1];
      const b = pixels[i + 2];
      const a = pixels[i + 3];
      const luma = 0.2126 * r + 0.7152 * g + 0.0722 * b;
      const colorDiff = Math.max(r, g, b) - Math.min(r, g, b);
      const isInkLike = (a > 18 && luma < 235 && colorDiff > 8) || (a > 30 && luma < 210);
      const isStroke = isInkLike;
      if (!isStroke) continue;
      const p = i / 4;
      const x = p % scanW;
      const y = Math.floor(p / scanW);
      binary[p] = 1;
      if (x < minX) minX = x;
      if (y < minY) minY = y;
      if (x > maxX) maxX = x;
      if (y > maxY) maxY = y;
    }

    const hasStroke = maxX >= minX && maxY >= minY;
    const srcToScan = 1 / scanScale;
    let cropX = 0;
    let cropY = 0;
    let cropW = srcW;
    let cropH = srcH;

    if (hasStroke) {
      const visited = new Uint8Array(scanW * scanH);
      const qx = new Int32Array(scanW * scanH);
      const qy = new Int32Array(scanW * scanH);
      const comps = [];
      for (let y = 0; y < scanH; y++) {
        for (let x = 0; x < scanW; x++) {
          const idx = y * scanW + x;
          if (!binary[idx] || visited[idx]) continue;
          let head = 0;
          let tail = 0;
          qx[tail] = x;
          qy[tail] = y;
          tail++;
          visited[idx] = 1;
          let cMinX = x, cMaxX = x, cMinY = y, cMaxY = y, count = 0;
          while (head < tail) {
            const cx = qx[head];
            const cy = qy[head];
            head++;
            count++;
            if (cx < cMinX) cMinX = cx;
            if (cx > cMaxX) cMaxX = cx;
            if (cy < cMinY) cMinY = cy;
            if (cy > cMaxY) cMaxY = cy;
            for (let oy = -1; oy <= 1; oy++) {
              for (let ox = -1; ox <= 1; ox++) {
                if (ox === 0 && oy === 0) continue;
                const nx = cx + ox;
                const ny = cy + oy;
                if (nx < 0 || ny < 0 || nx >= scanW || ny >= scanH) continue;
                const nIdx = ny * scanW + nx;
                if (!binary[nIdx] || visited[nIdx]) continue;
                visited[nIdx] = 1;
                qx[tail] = nx;
                qy[tail] = ny;
                tail++;
              }
            }
          }
          const cw = cMaxX - cMinX + 1;
          const ch = cMaxY - cMinY + 1;
          const ratio = cw / Math.max(1, ch);
          const cxMid = (cMinX + cMaxX) / 2;
          const leftWeight = 1 - Math.max(0, (cxMid / Math.max(1, scanW)) - 0.62);
          const shapeWeight = ratio >= 1.6 ? 1.2 : 0.8;
          const score = count * leftWeight * shapeWeight;
          comps.push({ cMinX, cMinY, cMaxX, cMaxY, count, score, ratio });
        }
      }
      let target = null;
      const minPixels = Math.max(40, Math.floor((scanW * scanH) * 0.0004));
      const candidates = comps.filter((c) => c.count >= minPixels);
      if (candidates.length) {
        candidates.sort((a, b) => b.score - a.score);
        target = candidates[0];
      }
      const bx0Base = target ? target.cMinX : minX;
      const by0Base = target ? target.cMinY : minY;
      const bx1Base = target ? target.cMaxX : maxX;
      const by1Base = target ? target.cMaxY : maxY;
      const padX = Math.max(6, Math.round((bx1Base - bx0Base + 1) * 0.12));
      const padY = Math.max(6, Math.round((by1Base - by0Base + 1) * 0.20));
      const bx0 = Math.max(0, bx0Base - padX);
      const by0 = Math.max(0, by0Base - padY);
      const bx1 = Math.min(scanW - 1, bx1Base + padX);
      const by1 = Math.min(scanH - 1, by1Base + padY);
      cropX = Math.max(0, Math.floor(bx0 * srcToScan));
      cropY = Math.max(0, Math.floor(by0 * srcToScan));
      cropW = Math.min(srcW - cropX, Math.max(1, Math.ceil((bx1 - bx0 + 1) * srcToScan)));
      cropH = Math.min(srcH - cropY, Math.max(1, Math.ceil((by1 - by0 + 1) * srcToScan)));
    }

    const targetW = 1100;
    const targetH = 320;
    const out = document.createElement("canvas");
    out.width = targetW;
    out.height = targetH;
    const outCtx = out.getContext("2d");
    if (!outCtx) return safeData;
    outCtx.clearRect(0, 0, targetW, targetH);

    const padOutX = 28;
    const padOutY = 26;
    const availW = targetW - (padOutX * 2);
    const availH = targetH - (padOutY * 2);
    const scale = Math.min(availW / Math.max(1, cropW), availH / Math.max(1, cropH));
    const drawW = Math.max(1, Math.round(cropW * scale));
    const drawH = Math.max(1, Math.round(cropH * scale));
    const dx = Math.round((targetW - drawW) / 2);
    const dy = Math.round((targetH - drawH) / 2);
    outCtx.imageSmoothingEnabled = true;
    outCtx.imageSmoothingQuality = "high";
    outCtx.drawImage(img, cropX, cropY, cropW, cropH, dx, dy, drawW, drawH);
    return out.toDataURL("image/png");
  } catch (_) {
    return safeData;
  }
}

function getSignatureWhitespaceWarning(analysis) {
  if (!analysis) return "";
  if (!analysis.hasStroke) return "No se detecta trazo claro de firma. Revisa el archivo.";
  if (analysis.whitespaceHeavy) return "La firma tiene mucho espacio en blanco; recorta la imagen para que no se vea pequena.";
  return "";
}

function renderUserSignaturePreview(dataUrl = "") {
  const src = String(dataUrl || "").trim();
  if (!el.userSignaturePreviewCard || !el.userSignaturePreview || !el.userSignatureMeta || !el.userSignatureWarn) return;
  if (!src || !isImageDataUrl(src)) {
    el.userSignaturePreviewCard.hidden = true;
    el.userSignaturePreview.removeAttribute("src");
    el.userSignatureMeta.textContent = "Sin firma cargada.";
    el.userSignatureWarn.hidden = true;
    el.userSignatureWarn.textContent = "";
    return;
  }
  el.userSignaturePreviewCard.hidden = false;
  el.userSignaturePreview.src = src;
  analyzeSignatureDataUrl(src).then((analysis) => {
    if (!analysis || String(el.userSignaturePreview.src || "") !== src) return;
    const area = Math.round(Number(analysis.contentAreaPct || 0));
    el.userSignatureMeta.textContent = `Resolucion ${analysis.width}x${analysis.height}px | Area util aprox. ${area}%`;
    const warn = getSignatureWhitespaceWarning(analysis);
    el.userSignatureWarn.hidden = !warn;
    el.userSignatureWarn.textContent = warn;
  }).catch(() => {
    if (String(el.userSignaturePreview.src || "") !== src) return;
    el.userSignatureMeta.textContent = "No se pudo analizar la firma.";
    el.userSignatureWarn.hidden = true;
    el.userSignatureWarn.textContent = "";
  });
}

function getBestAvailableSignatureDataUrl() {
  const sessionSig = String(authSession.signatureDataUrl || "").trim();
  if (isImageDataUrl(sessionSig)) return sessionSig;
  const authUser = (state.users || []).map(normalizeUserRecord).find((u) => String(u.id) === String(authSession.userId || ""));
  const authSig = String(authUser?.signatureDataUrl || "").trim();
  if (isImageDataUrl(authSig)) return authSig;
  const anySig = (state.users || [])
    .map(normalizeUserRecord)
    .map((u) => String(u.signatureDataUrl || "").trim())
    .find((sig) => isImageDataUrl(sig));
  return anySig || "";
}

const CHECKLIST_STATUS_CYCLE = ["ok", "x", "na"];

function normalizeChecklistStatus(status) {
  const s = String(status || "").trim().toLowerCase();
  if (s === "ok" || s === "x" || s === "na") return s;
  return "";
}

function checklistStatusLabel(status) {
  const s = normalizeChecklistStatus(status);
  if (s === "ok") return "Correcto";
  if (s === "x") return "X";
  if (s === "na") return "No aplica";
  return "Pendiente";
}

function checklistStatusBadgeText(status) {
  const s = normalizeChecklistStatus(status);
  if (s === "ok") return "OK";
  if (s === "x") return "X";
  if (s === "na") return "N/A";
  return "...";
}

function cycleChecklistStatus(status) {
  const s = normalizeChecklistStatus(status);
  if (!s) return CHECKLIST_STATUS_CYCLE[0];
  const idx = CHECKLIST_STATUS_CYCLE.indexOf(s);
  const next = CHECKLIST_STATUS_CYCLE[(idx + 1) % CHECKLIST_STATUS_CYCLE.length];
  return normalizeChecklistStatus(next);
}

function normalizeChecklistSectionRecord(raw) {
  const name = String(raw?.name || raw?.label || "").trim();
  if (!name) return null;
  return {
    id: String(raw?.id || uid()).trim(),
    name,
    active: raw?.active !== false,
  };
}

function ensureChecklistStores() {
  if (!Array.isArray(state.checklistTemplateItems)) state.checklistTemplateItems = [];
  if (!Array.isArray(state.checklistTemplateSections)) state.checklistTemplateSections = ["General"];
  if (!Array.isArray(state.checklistTemplates)) {
    const legacySections = state.checklistTemplateSections
      .map((s) => String(s || "").trim())
      .filter(Boolean);
    const sectionMap = new Map();
    const sections = Array.from(new Set(["General", ...legacySections])).map((name) => {
      const rec = { id: uid(), name, active: true };
      sectionMap.set(name.toLowerCase(), rec.id);
      return rec;
    });
    const items = (state.checklistTemplateItems || [])
      .map((raw) => {
        const item = normalizeChecklistTemplateItem(raw);
        if (!item) return null;
        const sectionName = String(raw?.section || "General").trim() || "General";
        item.sectionId = sectionMap.get(sectionName.toLowerCase()) || sections[0]?.id || "";
        return item;
      })
      .filter(Boolean);
    state.checklistTemplates = [{
      id: "tpl_default",
      name: "Checklist general",
      active: true,
      sections,
      items,
    }];
  }
  state.checklistTemplates = (state.checklistTemplates || [])
    .map(normalizeChecklistTemplateRecord)
    .filter(Boolean);
  if (!state.checklistTemplates.length) {
    state.checklistTemplates = [normalizeChecklistTemplateRecord({
      id: "tpl_default",
      name: "Checklist general",
      active: true,
      sections: [{ id: uid(), name: "General", active: true }],
      items: [],
    })];
  }
  syncLegacyChecklistStateFromTemplates();
  if (!state.eventChecklists || typeof state.eventChecklists !== "object") state.eventChecklists = {};
}

function normalizeChecklistTemplateItem(raw) {
  const label = String(raw?.label || raw?.name || "").trim();
  if (!label) return null;
  return {
    id: String(raw?.id || uid()).trim(),
    label,
    sectionId: String(raw?.sectionId || "").trim(),
    active: raw?.active !== false,
  };
}

function normalizeChecklistTemplateRecord(raw) {
  const name = String(raw?.name || raw?.title || "").trim() || "Checklist";
  let sections = Array.isArray(raw?.sections) ? raw.sections.map(normalizeChecklistSectionRecord).filter(Boolean) : [];
  const sectionMap = new Map(sections.map((s) => [String(s.id || "").trim(), s]));
  if (!sections.some((s) => String(s.name || "").trim().toLowerCase() === "general")) {
    const general = normalizeChecklistSectionRecord({ name: "General" });
    sections = [general, ...sections];
    sectionMap.set(general.id, general);
  }
  const generalId = sections[0]?.id || uid();
  const items = (Array.isArray(raw?.items) ? raw.items : [])
    .map((item) => {
      const normalized = normalizeChecklistTemplateItem(item);
      if (!normalized) return null;
      let sectionId = String(normalized.sectionId || "").trim();
      if (!sectionId && item?.section) {
        const sectionName = String(item.section || "").trim().toLowerCase();
        const found = sections.find((s) => String(s.name || "").trim().toLowerCase() === sectionName);
        sectionId = String(found?.id || "");
      }
      normalized.sectionId = sectionId && sectionMap.has(sectionId) ? sectionId : generalId;
      return normalized;
    })
    .filter(Boolean);
  return {
    id: String(raw?.id || uid()).trim(),
    name,
    active: raw?.active !== false,
    sections,
    items,
  };
}

function syncLegacyChecklistStateFromTemplates() {
  const templates = Array.isArray(state.checklistTemplates) ? state.checklistTemplates : [];
  const fallback = templates.find((tpl) => tpl?.active !== false) || templates[0] || null;
  const sections = Array.isArray(fallback?.sections) ? fallback.sections : [];
  const items = Array.isArray(fallback?.items) ? fallback.items : [];
  state.checklistTemplateSections = sections
    .filter((s) => s?.active !== false)
    .map((s) => String(s?.name || "").trim())
    .filter(Boolean);
  if (!state.checklistTemplateSections.length) state.checklistTemplateSections = ["General"];
  state.checklistTemplateItems = items
    .filter((x) => x?.active !== false)
    .map((x) => {
      const section = sections.find((s) => String(s?.id || "") === String(x?.sectionId || ""));
      return {
        id: String(x?.id || uid()).trim(),
        label: String(x?.label || "").trim(),
        section: String(section?.name || "General").trim() || "General",
        active: x?.active !== false,
      };
    })
    .filter((x) => x.label);
}

function getChecklistTemplates(options = {}) {
  ensureChecklistStores();
  const includeInactive = options && options.includeInactive === true;
  return (state.checklistTemplates || [])
    .map(normalizeChecklistTemplateRecord)
    .filter((tpl) => includeInactive || tpl.active !== false);
}

function getChecklistTemplateById(templateId = "", includeInactive = true) {
  const key = String(templateId || "").trim();
  if (!key) return null;
  return getChecklistTemplates({ includeInactive }).find((tpl) => String(tpl.id || "") === key) || null;
}

function getDefaultChecklistTemplate(includeInactive = false) {
  const rows = getChecklistTemplates({ includeInactive: includeInactive === true });
  if (!rows.length) return null;
  return rows.find((tpl) => tpl.active !== false) || rows[0] || null;
}

function getChecklistSections(templateId = "") {
  const template = getChecklistTemplateById(templateId, true) || getDefaultChecklistTemplate(true);
  const rows = Array.isArray(template?.sections) ? template.sections : [];
  return rows.filter((s) => s?.active !== false);
}

function getChecklistTemplateItems(templateId = "") {
  const template = getChecklistTemplateById(templateId, true) || getDefaultChecklistTemplate(true);
  const sections = Array.isArray(template?.sections) ? template.sections : [];
  const activeSections = new Map(
    sections
      .filter((s) => s?.active !== false)
      .map((s) => [String(s.id || ""), String(s.name || "General").trim() || "General"])
  );
  const items = Array.isArray(template?.items) ? template.items : [];
  return items
    .filter((x) => x?.active !== false)
    .filter((x) => {
      const sectionId = String(x?.sectionId || "").trim();
      return !sectionId || activeSections.has(sectionId);
    })
    .map((x) => {
      const sectionId = String(x?.sectionId || "").trim();
      const section = activeSections.get(sectionId) || "General";
      return {
        id: String(x?.id || uid()).trim(),
        label: String(x?.label || "").trim(),
        sectionId: sectionId || "",
        section,
        active: x?.active !== false,
      };
    })
    .filter((x) => x.label);
}

function normalizeEventChecklistRecord(raw, fallbackEventId = "") {
  const eventId = String(raw?.eventId || fallbackEventId || "").trim();
  const items = Array.isArray(raw?.items) ? raw.items : [];
  return {
    eventId,
    templateKey: String(raw?.templateKey || raw?.templateId || "").trim(),
    templateName: String(raw?.templateName || "").trim(),
    notes: String(raw?.notes || "").trim(),
    items: items.map((it) => ({
      id: String(it?.id || uid()).trim(),
      templateId: String(it?.templateId || "").trim(),
      label: String(it?.label || "").trim(),
      section: String(it?.section || "General").trim() || "General",
      status: normalizeChecklistStatus(it?.status),
      comment: String(it?.comment || "").trim(),
    })).filter((it) => it.label),
    updatedAt: String(raw?.updatedAt || "").trim(),
    completedAt: String(raw?.completedAt || "").trim(),
  };
}

function isChecklistCompleted(record) {
  const items = Array.isArray(record?.items) ? record.items : [];
  if (!items.length) return false;
  return items.every((it) => ["ok", "x", "na"].includes(normalizeChecklistStatus(it?.status)));
}

function getEventChecklistMeta(eventId) {
  ensureChecklistStores();
  const key = String(eventId || "").trim();
  const raw = key ? state.eventChecklists?.[key] : null;
  const rec = raw ? normalizeEventChecklistRecord(raw, key) : null;
  return {
    hasChecklist: !!rec && Array.isArray(rec.items) && rec.items.length > 0,
    completed: !!rec && isChecklistCompleted(rec),
    updatedAt: String(rec?.updatedAt || "").trim(),
  };
}

function buildEventChecklistDraft(eventId, preferredTemplateId = "") {
  ensureChecklistStores();
  const key = String(eventId || "").trim();
  const ev = (state.events || []).find((x) => String(x.id || "") === key);
  if (!ev) return null;
  const savedRaw = state.eventChecklists?.[key] || null;
  const saved = savedRaw ? normalizeEventChecklistRecord(savedRaw, key) : null;
  const chosenTemplate = getChecklistTemplateById(preferredTemplateId, true)
    || getChecklistTemplateById(saved?.templateKey || "", true)
    || getDefaultChecklistTemplate(true);
  const templateItems = getChecklistTemplateItems(chosenTemplate?.id || "");
  const savedByTemplate = new Map();
  const savedByLabel = new Map();
  for (const it of saved?.items || []) {
    const tpl = String(it.templateId || "").trim();
    const lbl = String(it.label || "").trim().toLowerCase();
    if (tpl) savedByTemplate.set(tpl, it);
    if (lbl) savedByLabel.set(lbl, it);
  }
  const items = templateItems.map((tpl) => {
    const tplId = String(tpl.id || "").trim();
    const lbl = String(tpl.label || "").trim();
    const savedHit = (tplId && savedByTemplate.get(tplId)) || savedByLabel.get(lbl.toLowerCase()) || null;
    return {
      id: String(savedHit?.id || uid()).trim(),
      templateId: tplId,
      label: lbl,
      section: String(tpl.section || "General").trim() || "General",
      status: normalizeChecklistStatus(savedHit?.status),
      comment: String(savedHit?.comment || "").trim(),
    };
  });
  return {
    eventId: key,
    templateKey: String(chosenTemplate?.id || "").trim(),
    templateName: String(chosenTemplate?.name || "").trim(),
    eventName: String(ev.name || "").trim(),
    eventDate: String(ev.date || "").trim(),
    salon: String(ev.salon || "").trim(),
    notes: String(saved?.notes || "").trim(),
    items,
    updatedAt: String(saved?.updatedAt || "").trim(),
    completedAt: String(saved?.completedAt || "").trim(),
  };
}

function setActiveModuleScreen(screen) {
  const target = String(screen || "").trim();
  if (el.moduleHubScreen) el.moduleHubScreen.hidden = target !== "hub";
  if (el.reportsHubScreen) el.reportsHubScreen.hidden = target !== "reports";
  if (el.settingsScreen) el.settingsScreen.hidden = target !== "settings";
  savePersistedModuleScreen(target);
}

function setSideNavActive(key = "calendar") {
  const items = [
    { key: "dashboard", node: el.btnSideDashboardReports },
    { key: "calendar", node: el.btnSideCalendar },
    { key: "reports", node: el.btnSideReports },
    { key: "settings", node: el.btnSideSettings },
  ];
  const target = String(key || "").trim() || "calendar";
  for (const item of items) {
    if (!item?.node) continue;
    const active = item.key === target;
    item.node.classList.toggle("isActive", active);
    item.node.setAttribute("aria-current", active ? "page" : "false");
  }
}

function showModuleHub() {
  setActiveModuleScreen("hub");
}

function showCalendarModule() {
  setActiveModuleScreen("");
  setSideNavActive("calendar");
}

function showReportsHub() {
  setActiveModuleScreen("reports");
  setSideNavActive("reports");
}

function showSettingsHub() {
  setQuickAddGroupOpen(true);
  setSettingsPanelOpen(true);
  setSideNavActive("settings");
}

function getActiveModuleScreenName() {
  if (el.settingsScreen && !el.settingsScreen.hidden) return "settings";
  if (el.reportsHubScreen && !el.reportsHubScreen.hidden) return "reports";
  if (el.moduleHubScreen && !el.moduleHubScreen.hidden) return "hub";
  return "";
}

function prepareModuleModalOpen(preferredScreen = "") {
  const target = String(preferredScreen || "").trim() || getActiveModuleScreenName();
  moduleModalReturnScreen = target;
  if (target) {
    setActiveModuleScreen("");
  }
}

function restoreModuleScreenAfterModal() {
  const target = String(moduleModalReturnScreen || "").trim();
  if (!target) return;
  setActiveModuleScreen(target);
  moduleModalReturnScreen = "";
}

function closeBlockingOverlaysForNavigation() {
  const backdrops = [
    el.salesReportBackdrop,
    el.accountingReportBackdrop,
    el.accountStatementBackdrop,
    el.occupancyReportBackdrop,
    el.dashboardReportBackdrop,
    el.institutionReportBackdrop,
    el.checklistTemplateBackdrop,
    el.eventChecklistBackdrop,
    el.modalBackdrop,
    el.appointmentBackdrop,
    el.eventFinderBackdrop,
    el.leadsBackdrop,
    el.userBackdrop,
    el.quoteBackdrop,
    el.quoteAdvanceBackdrop,
    el.menuMontajeBackdrop,
    el.menuMontajeSelectableBackdrop,
    el.quoteServiceTemplateManagerBackdrop,
    el.companyBackdrop,
    el.serviceBackdrop,
    el.serviceCategoryBackdrop,
    el.serviceSubcategoryBackdrop,
    el.globalGoalsBackdrop,
    el.salonesBackdrop,
    el.menuSuggestionsBackdrop,
    el.menuCatalogBackdrop,
  ];

  for (const node of backdrops) {
    if (node) node.hidden = true;
  }

  if (el.quoteBackdrop) el.quoteBackdrop.classList.remove("docFloatOpen");
  document.body.classList.remove("eventModalOpen", "quoteModeOpen", "menuSuggestionsModalOpen");
  moduleModalReturnScreen = "";
}

function setSettingsPanelOpen(open) {
  if (!el.settingsPanel) return;
  if (el.settingsScreen) el.settingsScreen.hidden = !open;
  if (el.btnSettings) el.btnSettings.setAttribute("aria-expanded", open ? "true" : "false");
  if (open) {
    setActiveModuleScreen("settings");
  }
  if (open) {
    if (!el.settingsPanel.hasAttribute("tabindex")) {
      el.settingsPanel.setAttribute("tabindex", "-1");
    }
    setTimeout(() => {
      try { el.settingsPanel.focus(); } catch (_) { }
    }, 0);
  } else {
    if (getActiveModuleScreenName() === "settings") {
      showCalendarModule();
    }
    setQuickAddGroupOpen(false);
    setReportsGroupOpen(false);
  }
}

function closeSettingsPanel() {
  setSettingsPanelOpen(false);
}

function setQuickAddGroupOpen(open) {
  if (!el.quickAddGroup || !el.btnToggleQuickAdd) return;
  el.quickAddGroup.hidden = !open;
  el.btnToggleQuickAdd.setAttribute("aria-expanded", open ? "true" : "false");
}

function setReportsGroupOpen(open) {
  if (!el.reportsGroup || !el.btnToggleReports) return;
  el.reportsGroup.hidden = !open;
  el.btnToggleReports.setAttribute("aria-expanded", open ? "true" : "false");
}

function normalizeBucketKey(value) {
  return String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, " ")
    .trim();
}

function tokenizeLikeSearch(value) {
  const normalized = normalizeBucketKey(value);
  return normalized ? normalized.split(/\s+/).filter(Boolean) : [];
}

function matchesLikeSearch(haystack, query) {
  const tokens = Array.isArray(query) ? query.filter(Boolean) : tokenizeLikeSearch(query);
  if (!tokens.length) return true;
  const normalizedHaystack = normalizeBucketKey(haystack);
  if (!normalizedHaystack) return false;
  return tokens.every((token) => normalizedHaystack.includes(token));
}

function matchesAliases(value, aliases = []) {
  const base = normalizeBucketKey(value);
  if (!base) return false;
  return aliases.some((a) => base.includes(normalizeBucketKey(a)));
}

function formatMoneyGTValue(v) {
  return Number(v || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function moneyGT(v) {
  return `Q ${formatMoneyGTValue(v)}`;
}

function normalizeQuoteCurrency(value) {
  return String(value || "").trim().toUpperCase() === "USD" ? "USD" : "GTQ";
}

function getQuoteCurrencySymbol(currency) {
  return normalizeQuoteCurrency(currency) === "USD" ? "$" : "Q";
}

function getQuoteCurrencyWord(currency, whole) {
  if (normalizeQuoteCurrency(currency) === "USD") {
    return Number(whole || 0) === 1 ? "DOLAR" : "DOLARES";
  }
  return Number(whole || 0) === 1 ? "QUETZAL" : "QUETZALES";
}

function moneyQuote(v, currency = "GTQ") {
  return `${getQuoteCurrencySymbol(currency)} ${formatMoneyGTValue(v)}`;
}

function numberToWordsQuoteEs(value, currency = "GTQ") {
  const units = ["", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE"];
  const teens = ["DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE"];
  const tens = ["", "", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"];
  const hundreds = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"];
  const belowHundred = (n) => {
    if (n < 10) return units[n];
    if (n < 20) return teens[n - 10];
    if (n < 30) return n === 20 ? "VEINTE" : `VEINTI${units[n - 20].toLowerCase()}`.toUpperCase();
    const t = Math.floor(n / 10);
    const u = n % 10;
    return u ? `${tens[t]} Y ${units[u]}` : tens[t];
  };
  const belowThousand = (n) => {
    if (n === 0) return "";
    if (n === 100) return "CIEN";
    const h = Math.floor(n / 100);
    const r = n % 100;
    if (!h) return belowHundred(r);
    return r ? `${hundreds[h]} ${belowHundred(r)}` : hundreds[h];
  };
  const toWords = (n) => {
    if (n === 0) return "CERO";
    const millions = Math.floor(n / 1000000);
    const thousands = Math.floor((n % 1000000) / 1000);
    const rest = n % 1000;
    const parts = [];
    if (millions) parts.push(millions === 1 ? "UN MILLON" : `${toWords(millions)} MILLONES`);
    if (thousands) parts.push(thousands === 1 ? "MIL" : `${belowThousand(thousands)} MIL`);
    if (rest) parts.push(belowThousand(rest));
    return parts.join(" ").replace(/\s+/g, " ").trim();
  };
  const amount = Math.max(0, Number(value || 0));
  const whole = Math.floor(amount);
  const cents = Math.round((amount - whole) * 100);
  const centsText = String(cents).padStart(2, "0");
  return `${toWords(whole)} ${getQuoteCurrencyWord(currency, whole)} CON ${centsText}/100`;
}

function getEventLastUpdatedLabel(ev) {
  const key = reservationKeyFromEvent(ev);
  const rows = Array.isArray(state.changeHistory?.[key]) ? state.changeHistory[key] : [];
  const latest = rows[0]?.at || ev?.quote?.quotedAt || "";
  if (!latest) return "";
  return formatDateTime(latest);
}

function aggregateQuoteBuckets(quote) {
  const items = Array.isArray(quote?.items) ? quote.items : [];
  const subcatBuckets = {
    desayunos: { qty: 0, amount: 0, aliases: ["desayuno"] },
    refa: { qty: 0, amount: 0, aliases: ["refa"] },
    almuerzos: { qty: 0, amount: 0, aliases: ["almuerzo"] },
    amRefa: { qty: 0, amount: 0, aliases: ["am refa", "refa am"] },
    pmRefa: { qty: 0, amount: 0, aliases: ["pm refa", "refa pm"] },
    cenasBuffet: { qty: 0, amount: 0, aliases: ["cena buffet", "buffet cena"] },
    miscelaneos: { qty: 0, amount: 0, aliases: ["miscelaneo", "miscelaneos"] },
  };
  const catBuckets = {
    alimentosBebidas: { qty: 0, amount: 0, aliases: ["alimentos y bebidas", "alimentos", "bebidas"] },
    hospedajeJdl: { qty: 0, amount: 0, aliases: ["hospedaje jdl"] },
    hospedajeTerceros: { qty: 0, amount: 0, aliases: ["hospedaje terceros", "hospedaje tercero"] },
    miscelaneos: { qty: 0, amount: 0, aliases: ["miscelaneo", "miscelaneos"] },
  };

  for (const it of items) {
    const qty = Math.max(0, Number(it?.qty || 0));
    const unit = Math.max(0, Number(it?.price || 0));
    const lineAmount = qty * unit;
    const subcategory = String(it?.subcategory || it?.subcategoria || "");
    const category = String(it?.category || it?.categoria || "");
    const subNorm = normalizeBucketKey(subcategory);
    let subKey = "";
    if (matchesAliases(subNorm, subcatBuckets.amRefa.aliases)) subKey = "amRefa";
    else if (matchesAliases(subNorm, subcatBuckets.pmRefa.aliases)) subKey = "pmRefa";
    else if (matchesAliases(subNorm, subcatBuckets.cenasBuffet.aliases)) subKey = "cenasBuffet";
    else if (matchesAliases(subNorm, subcatBuckets.desayunos.aliases)) subKey = "desayunos";
    else if (matchesAliases(subNorm, subcatBuckets.almuerzos.aliases)) subKey = "almuerzos";
    else if (matchesAliases(subNorm, subcatBuckets.refa.aliases)) subKey = "refa";
    else if (matchesAliases(subNorm, subcatBuckets.miscelaneos.aliases)) subKey = "miscelaneos";
    if (subKey) {
      subcatBuckets[subKey].qty += qty;
      subcatBuckets[subKey].amount += lineAmount;
    }
    for (const bucket of Object.values(catBuckets)) {
      if (matchesAliases(category, bucket.aliases)) {
        bucket.qty += qty;
        bucket.amount += lineAmount;
      }
    }
  }
  return { subcatBuckets, catBuckets };
}

function buildSalesReportRows() {
  const rows = [];
  const seenReservations = new Set();
  for (const ev of state.events || []) {
    const reservationKey = String(reservationKeyFromEvent(ev) || ev?.id || "").trim();
    if (reservationKey) {
      if (seenReservations.has(reservationKey)) continue;
      seenReservations.add(reservationKey);
    }
    const financialMeta = getEventSeriesFinancialMeta(ev);
    const primaryEvent = financialMeta.primaryEvent || ev;
    const quote = primaryEvent?.quote || ev?.quote || null;
    const user = (state.users || []).find((u) => String(u.id) === String(primaryEvent?.userId || ev?.userId));
    const company = quote?.companyId ? (state.companies || []).find((c) => String(c.id) === String(quote.companyId)) : null;
    const manager = company?.managers?.find((m) => String(m.id) === String(quote?.managerId));
    const totals = getQuoteTotals(quote || {});
    const { subcatBuckets, catBuckets } = aggregateQuoteBuckets(quote || {});
    rows.push({
      event: primaryEvent,
      reservationKey,
      actionEventId: String(primaryEvent?.id || ev?.id || ""),
      companyId: String(quote?.companyId || company?.id || ""),
      status: String(primaryEvent?.status || ev?.status || ""),
      statusColor: statusColor(primaryEvent?.status || ev?.status),
      refId: String(quote?.code || reservationKey || primaryEvent?.id || ev?.id || ""),
      seller: String(user?.fullName || user?.name || ""),
      eventDate: String(financialMeta.startDate || primaryEvent?.date || ev?.date || ""),
      eventType: String(quote?.eventType || primaryEvent?.name || ev?.name || ""),
      startTime: String(financialMeta.startTime || primaryEvent?.startTime || ev?.startTime || ""),
      endTime: String(financialMeta.endTime || primaryEvent?.endTime || ev?.endTime || ""),
      salon: String(financialMeta.mainSalon || primaryEvent?.salon || ev?.salon || ""),
      mainSalon: String(financialMeta.mainSalon || primaryEvent?.salon || ev?.salon || ""),
      salones: Array.isArray(financialMeta.salones) ? financialMeta.salones.slice() : [],
      salonesLabel: Array.isArray(financialMeta.salones) ? financialMeta.salones.join(", ") : "",
      company: String(company?.name || quote?.companyName || ""),
      manager: String(manager?.phone || quote?.managerPhone || ""),
      pax: Number(primaryEvent?.pax || ev?.pax || quote?.people || 0),
      paymentType: String(quote?.paymentType || "").trim(),
      dueDate: String(quote?.dueDate || "").trim(),
      docDate: String(quote?.docDate || "").trim(),
      advances: normalizeQuoteAdvancesForSnapshot(quote?.advances),
      subtotal: Number(totals.subtotal || 0),
      total: Number(totals.total || 0),
      subcatBuckets,
      catBuckets,
      discount: Number(totals.discountAmount || 0),
      updatedAt: getEventLastUpdatedLabel(primaryEvent || ev),
    });
  }
  return rows;
}

function getSalesReportFilteredRows() {
  const search = String(el.salesReportSearch?.value || "").trim().toLowerCase();
  const from = String(el.salesReportFrom?.value || "").trim();
  const to = String(el.salesReportTo?.value || "").trim();
  const userId = String(el.salesReportUser?.value || "").trim();
  const status = String(el.salesReportStatus?.value || "").trim();
  const salon = String(el.salesReportSalon?.value || "").trim();
  const company = String(el.salesReportCompany?.value || "").trim();

  return enrichAccountingReportRows(buildSalesReportRows()).filter((r) => {
    if (from && r.eventDate && r.eventDate < from) return false;
    if (to && r.eventDate && r.eventDate > to) return false;
    if (userId && String(r.event?.userId || "") !== userId) return false;
    if (status && r.status !== status) return false;
    if (salon) {
      const eventSalones = Array.isArray(r.salones) ? r.salones : [];
      if (r.salon !== salon && !eventSalones.includes(salon)) return false;
    }
    if (company && String(r.event?.quote?.companyId || "") !== company) return false;
    if (search) {
      const blob = [
        r.refId, r.seller, r.eventType, r.salon, r.salonesLabel, r.company, r.manager, r.status,
      ].join(" ");
      if (!matchesLikeSearch(blob, search)) return false;
    }
    return true;
  }).sort((a, b) => {
    const d = String(a.eventDate || "").localeCompare(String(b.eventDate || ""));
    if (d !== 0) return d;
    return String(a.startTime || "").localeCompare(String(b.startTime || ""));
  });
}

function renderSalesReportFilters() {
  const users = (state.users || []).filter((u) => u.active !== false);
  const statuses = Array.from(new Set((state.events || []).map((e) => String(e.status || "")).filter(Boolean))).sort();
  const salones = getUsedSalonesForReports();
  const companies = (state.companies || []).filter((c) => !isCompanyDisabled(c.id));

  const fillSelect = (node, rows, allLabel) => {
    if (!node) return;
    const previous = String(node.value || "");
    node.innerHTML = "";
    const all = document.createElement("option");
    all.value = "";
    all.textContent = allLabel;
    node.appendChild(all);
    for (const row of rows) {
      const opt = document.createElement("option");
      opt.value = String(row.value);
      opt.textContent = String(row.label);
      node.appendChild(opt);
    }
    if (previous && rows.some((r) => String(r.value) === previous)) node.value = previous;
    enhanceSelectControl(node, true);
  };

  fillSelect(el.salesReportUser, users.map((u) => ({ value: u.id, label: u.fullName || u.name })), "Todos vendedores");
  fillSelect(el.salesReportStatus, statuses.map((s) => ({ value: s, label: s })), "Todos estados");
  fillSelect(el.salesReportSalon, salones.map((s) => ({ value: s, label: s })), "Todos salones");
  fillSelect(el.salesReportCompany, companies.map((c) => ({ value: c.id, label: c.name })), "Todas instituciones");
}

function renderSalesReportSummary(rows = []) {
  const safeRows = Array.isArray(rows) ? rows : [];
  if (el.salesReportFiltersMeta) {
    el.salesReportFiltersMeta.textContent = salesReportFiltersSummaryText();
  }
  if (!el.salesReportSummary) return;
  const confirmedCount = safeRows.filter((row) => String(row?.status || "") === STATUS.CONFIRMADO).length;
  const preReservedCount = safeRows.filter((row) => String(row?.status || "") === STATUS.PRE_RESERVA).length;
  const totalPax = safeRows.reduce((acc, row) => acc + Math.max(0, Number(row?.pax || 0)), 0);
  const totalAmount = safeRows.reduce((acc, row) => acc + Math.max(0, Number(row?.total || 0)), 0);
  const avgTicket = safeRows.length ? (totalAmount / safeRows.length) : 0;
  const baseRowsCount = enrichAccountingReportRows(buildSalesReportRows()).length || safeRows.length;
  const pipelinePct = baseRowsCount ? Math.max(0, Math.round((safeRows.length / baseRowsCount) * 100)) : 0;
  const conversionPct = safeRows.length ? Math.max(0, Math.round((confirmedCount / safeRows.length) * 100)) : 0;
  const topSellerMap = new Map();
  for (const row of safeRows) {
    const seller = String(row?.seller || "Sin vendedor").trim() || "Sin vendedor";
    topSellerMap.set(seller, Number(topSellerMap.get(seller) || 0) + Math.max(0, Number(row?.total || 0)));
  }
  const topSeller = Array.from(topSellerMap.entries()).sort((a, b) => Number(b[1] || 0) - Number(a[1] || 0))[0];
  const cards = [
    { tone: "primary", icon: "inventory_2", eyebrow: "Pipeline", label: "Registros visibles", value: String(safeRows.length), meta: "Registros visibles", trend: `+${pipelinePct}%` },
    { tone: "success", icon: "check_circle", eyebrow: "Conversion", label: "Confirmado", value: String(confirmedCount), meta: "Confirmado", trend: `${conversionPct}%` },
    { tone: "info", icon: "groups", eyebrow: "Capacidad", label: "PAX total", value: String(totalPax), meta: "PAX total", trend: "" },
    { tone: "ticket", icon: "payments", eyebrow: "Ticket Promedio", label: "Por evento", value: moneyGT(avgTicket), meta: "Por evento", trend: "" },
    { tone: "accent", icon: "trending_up", eyebrow: "Facturacion", label: "Total cotizado", value: moneyGT(totalAmount), meta: "Total cotizado", trend: topSeller ? `Top vendedor ${topSeller[0]}` : "" },
  ];
  el.salesReportSummary.innerHTML = cards.map((card) => `
    <article class="salesSummaryCard salesSummaryCard--${escapeHtml(card.tone)}" title="${escapeHtml(({ Pipeline: "Pipeline = cuantos registros u oportunidades tienes activas en el flujo", Conversion: "Conversion = cuantas ya pasaron a confirmados", Capacidad: "Capacidad = volumen o PAX asociado", Facturacion: "Facturacion = cuanto dinero representa ese pipeline o lo ya cerrado" }[card.eyebrow] || ""))}">
      <div class="salesSummaryHead">
        <span class="salesSummaryIcon material-symbols-outlined">${escapeHtml(card.icon || "insights")}</span>
        ${card.trend ? `<span class="salesSummaryTrend">${escapeHtml(card.trend)}</span>` : ""}
      </div>
      <span class="salesSummaryEyebrow">${escapeHtml(card.eyebrow)}</span>
      <small>${escapeHtml(card.label)}</small>
      <strong>${escapeHtml(card.value)}</strong>
      <div class="salesSummaryMeta">${escapeHtml(card.meta)}</div>
    </article>
  `).join("");
}

function renderSalesReportTable() {
  if (!el.salesReportBody) return;
  const rows = getSalesReportFilteredRows();
  renderSalesReportSummary(rows);
  el.salesReportBody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="9">Sin resultados para los filtros seleccionados.</td>`;
    el.salesReportBody.appendChild(tr);
    return;
  }
  for (const r of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><span class="salesStatusBadge" style="background:${hexToRgba(r.statusColor, 0.25)};border-color:${hexToRgba(r.statusColor, 0.6)}">${escapeHtml(r.status || "-")}</span></td>
      <td>${escapeHtml(r.refId || "-")}</td>
      <td class="salesSellerCell">
        <span class="salesSellerAvatar">${escapeHtml(getInitials(r.seller || "-"))}</span>
        <span>${escapeHtml(r.seller || "-")}</span>
      </td>
      <td>${escapeHtml(r.eventDate || "-")}</td>
      <td>${escapeHtml(r.eventType || "-")}</td>
      <td>${escapeHtml(r.startTime || "-")} - ${escapeHtml(r.endTime || "-")}</td>
      <td>${escapeHtml(r.salon || "-")}</td>
      <td>${escapeHtml(String(r.pax || 0))}</td>
      <td>${escapeHtml(moneyGT(r.total || 0))}</td>
    `;
    el.salesReportBody.appendChild(tr);
  }
}

function salesReportFiltersSummaryText() {
  const parts = [];
  const from = String(el.salesReportFrom?.value || "").trim();
  const to = String(el.salesReportTo?.value || "").trim();
  const sellerOpt = el.salesReportUser?.selectedOptions?.[0]?.textContent || "";
  const statusOpt = el.salesReportStatus?.selectedOptions?.[0]?.textContent || "";
  const salonOpt = el.salesReportSalon?.selectedOptions?.[0]?.textContent || "";
  const companyOpt = el.salesReportCompany?.selectedOptions?.[0]?.textContent || "";
  const search = String(el.salesReportSearch?.value || "").trim();
  if (from || to) parts.push(`Rango: ${from || "..."} a ${to || "..."}`);
  if (sellerOpt && !/^todos/i.test(sellerOpt)) parts.push(`Vendedor: ${sellerOpt}`);
  if (statusOpt && !/^todos/i.test(statusOpt)) parts.push(`Estado: ${statusOpt}`);
  if (salonOpt && !/^todos/i.test(salonOpt)) parts.push(`Salon: ${salonOpt}`);
  if (companyOpt && !/^todas/i.test(companyOpt)) parts.push(`Institucion: ${companyOpt}`);
  if (search) parts.push(`Buscar: ${search}`);
  return parts.length ? parts.join(" | ") : "Sin filtros";
}

function exportSalesReportToExcel() {
  const rows = getSalesReportFilteredRows();
  if (!rows.length) return toast("No hay datos para exportar.");
  const pick = (obj, key) => obj?.[key] || { qty: 0, amount: 0 };
  const generatedAt = new Date().toLocaleString("es-GT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
  const emittedBy = String(authSession.fullName || authSession.username || "Sistema").trim();

  const htmlRows = rows.map((r) => `
    <tr>
      <td style="background:${escapeHtml(hexToRgba(r.statusColor, 0.25))}; border:1px solid #c7d5ea; font-weight:700;">${escapeHtml(r.status || "-")}</td>
      <td>${escapeHtml(r.refId || "-")}</td>
      <td>${escapeHtml(r.seller || "-")}</td>
      <td>${escapeHtml(r.eventDate || "-")}</td>
      <td>${escapeHtml(r.eventType || "-")}</td>
      <td>${escapeHtml(r.startTime || "-")}</td>
      <td>${escapeHtml(r.endTime || "-")}</td>
      <td>${escapeHtml(r.salon || "-")}</td>
      <td>${escapeHtml(r.company || "-")}</td>
      <td>${escapeHtml(r.manager || "-")}</td>
      <td>${escapeHtml(String(r.pax || 0))}</td>
      <td>${escapeHtml(String(pick(r.subcatBuckets, "desayunos").qty || 0))}</td>
      <td>${escapeHtml(moneyGT(pick(r.subcatBuckets, "desayunos").amount))}</td>
      <td>${escapeHtml(String(pick(r.subcatBuckets, "refa").qty || 0))}</td>
      <td>${escapeHtml(moneyGT(pick(r.subcatBuckets, "refa").amount))}</td>
      <td>${escapeHtml(String(pick(r.subcatBuckets, "almuerzos").qty || 0))}</td>
      <td>${escapeHtml(moneyGT(pick(r.subcatBuckets, "almuerzos").amount))}</td>
      <td>${escapeHtml(String(pick(r.subcatBuckets, "amRefa").qty || 0))}</td>
      <td>${escapeHtml(moneyGT(pick(r.subcatBuckets, "amRefa").amount))}</td>
      <td>${escapeHtml(String(pick(r.subcatBuckets, "pmRefa").qty || 0))}</td>
      <td>${escapeHtml(moneyGT(pick(r.subcatBuckets, "pmRefa").amount))}</td>
      <td>${escapeHtml(String(pick(r.subcatBuckets, "cenasBuffet").qty || 0))}</td>
      <td>${escapeHtml(moneyGT(pick(r.subcatBuckets, "cenasBuffet").amount))}</td>
      <td>${escapeHtml(String(pick(r.subcatBuckets, "miscelaneos").qty || 0))}</td>
      <td>${escapeHtml(moneyGT(pick(r.subcatBuckets, "miscelaneos").amount))}</td>
      <td>${escapeHtml(String(pick(r.catBuckets, "alimentosBebidas").qty || 0))}</td>
      <td>${escapeHtml(moneyGT(pick(r.catBuckets, "alimentosBebidas").amount))}</td>
      <td>${escapeHtml(String(pick(r.catBuckets, "hospedajeJdl").qty || 0))}</td>
      <td>${escapeHtml(moneyGT(pick(r.catBuckets, "hospedajeJdl").amount))}</td>
      <td>${escapeHtml(String(pick(r.catBuckets, "hospedajeTerceros").qty || 0))}</td>
      <td>${escapeHtml(moneyGT(pick(r.catBuckets, "hospedajeTerceros").amount))}</td>
      <td>${escapeHtml(String(pick(r.catBuckets, "miscelaneos").qty || 0))}</td>
      <td>${escapeHtml(moneyGT(pick(r.catBuckets, "miscelaneos").amount))}</td>
      <td>${escapeHtml(moneyGT(r.discount || 0))}</td>
      <td>${escapeHtml(r.updatedAt || "-")}</td>
    </tr>
  `).join("");
  const html = `<!doctype html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="utf-8" />
  <meta name="ProgId" content="Excel.Sheet" />
  <meta name="Generator" content="CRM Jardines" />
      <style>
    body{ font-family: Calibri, Arial, sans-serif; background:#eef3fb; margin:0; padding:16px; color:#0f172a; }
    .card{ background:#ffffff; border:1px solid #c5d4ea; border-radius:10px; overflow:hidden; }
    .meta{ padding:10px 14px; border-top:1px solid #bfd3ee; border-bottom:1px solid #bfd3ee; background:#eaf3ff; font-size:12px; }
    .meta div{ margin:2px 0; }
    table{ width:100%; border-collapse:collapse; }
    th,td{ border:1px solid #c7d5ea; padding:6px 7px; font-size:10.5px; white-space:nowrap; }
    thead th{ background:#0f3c67; color:#fff; font-weight:700; text-transform:uppercase; }
    .titleTable{ width:100%; border-collapse:collapse; }
    .titleCell{
      border:1px solid #c7d5ea;
      background:#d8e3f3;
      color:#000;
      font-weight:800;
      font-size:20px;
      letter-spacing:.3px;
      padding:12px 14px;
      text-transform:uppercase;
    }
  </style>
</head>
<body>
  <div class="card">
    <table class="titleTable">
      <tr><td class="titleCell">CRM JARDINES - REPORTE DE VENTAS</td></tr>
    </table>
    <div class="meta">
      <div><b>Fecha:</b> ${escapeHtml(generatedAt)}</div>
      <div><b>Quien emitio el reporte:</b> ${escapeHtml(emittedBy)}</div>
      <div><b>Filtros aplicados:</b> ${escapeHtml(salesReportFiltersSummaryText())}</div>
      <div><b>Total registros:</b> ${rows.length}</div>
    </div>
    <table class="quoteItemsTablePrint">
      <colgroup>
        <col class="qtyCol" />
        <col class="descCol" />
        <col class="priceCol" />
        <col class="totalCol" />
      </colgroup>
      <thead>
        <tr>
          <th>Estado</th><th>ID Cotizacion/Reserva</th><th>Vendedor</th><th>Fecha evento</th><th>Tipo evento</th><th>Hora inicio</th><th>Hora final</th><th>Salon</th><th>Institucion</th><th>Encargado evento</th><th>PAX</th><th>Cant Desayunos</th><th>Monto Desayunos</th><th>Cant Refa</th><th>Monto Refa</th><th>Cant Almuerzos</th><th>Monto Almuerzos</th><th>Cant AM Refa</th><th>Monto AM Refa</th><th>Cant PM Refa</th><th>Monto PM Refa</th><th>Cant Cenas Buffet</th><th>Monto Cenas Buffet</th><th>Cant Miscelaneos</th><th>Monto Miscelaneos</th><th>Cat A&B Cant</th><th>Cat A&B Monto</th><th>Cat Hospedaje JDL Cant</th><th>Cat Hospedaje JDL Monto</th><th>Cat Hospedaje Terceros Cant</th><th>Cat Hospedaje Terceros Monto</th><th>Cat Miscelaneos Cant</th><th>Cat Miscelaneos Monto</th><th>Descuento</th><th>Ultima modificacion</th>
        </tr>
      </thead>
      <tbody>${htmlRows}</tbody>
    </table>
  </div>
</body>
</html>`;

  const blob = new Blob([`\uFEFF${html}`], { type: "application/vnd.ms-excel;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const stamp = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = `reporte_ventas_${stamp}.xls`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function resetSalesReportFilters() {
  if (el.salesReportSearch) el.salesReportSearch.value = "";
  if (el.salesReportFrom) el.salesReportFrom.value = "";
  if (el.salesReportTo) el.salesReportTo.value = "";
  if (el.salesReportUser) el.salesReportUser.value = "";
  if (el.salesReportStatus) el.salesReportStatus.value = "";
  if (el.salesReportSalon) el.salesReportSalon.value = "";
  if (el.salesReportCompany) el.salesReportCompany.value = "";
}

function openSalesReportModal() {
  if (!el.salesReportBackdrop) return;
  if (el.occupancyReportBackdrop) el.occupancyReportBackdrop.hidden = true;
  if (el.dashboardReportBackdrop) el.dashboardReportBackdrop.hidden = true;
  if (el.institutionReportBackdrop) el.institutionReportBackdrop.hidden = true;
  renderSalesReportFilters();
  resetSalesReportFilters();
  renderSalesReportTable();
  el.salesReportBackdrop.hidden = false;
}

function closeSalesReportModal() {
  if (!el.salesReportBackdrop) return;
  el.salesReportBackdrop.hidden = true;
  restoreModuleScreenAfterModal();
}

function accountingReportFiltersSummaryText() {
  const parts = [];
  const from = String(el.accountingReportFrom?.value || "").trim();
  const to = String(el.accountingReportTo?.value || "").trim();
  const sellerOpt = el.accountingReportUser?.selectedOptions?.[0]?.textContent || "";
  const statusOpt = el.accountingReportStatus?.selectedOptions?.[0]?.textContent || "";
  const salonOpt = el.accountingReportSalon?.selectedOptions?.[0]?.textContent || "";
  const companyOpt = el.accountingReportCompany?.selectedOptions?.[0]?.textContent || "";
  const search = String(el.accountingReportSearch?.value || "").trim();
  if (from || to) parts.push(`Rango: ${from || "..."} a ${to || "..."}`);
  if (sellerOpt && !/^todos/i.test(sellerOpt)) parts.push(`Vendedor: ${sellerOpt}`);
  if (statusOpt && !/^todos/i.test(statusOpt)) parts.push(`Estado: ${statusOpt}`);
  if (salonOpt && !/^todos/i.test(salonOpt)) parts.push(`Salon: ${salonOpt}`);
  if (companyOpt && !/^todas/i.test(companyOpt)) parts.push(`Institucion: ${companyOpt}`);
  if (search) parts.push(`Buscar: ${search}`);
  return parts.length ? parts.join(" | ") : "Sin filtros";
}

const accountingReportExpandedAccounts = new Set();
let accountingReportReturnAfterPayment = false;
let accountingReturnAfterPaymentTarget = "";
let accountingStatementActiveCompanyKey = "";

function summarizeAccountingCollectionState(account) {
  const pending = Math.max(0, Number(account?.pendingAmount || 0));
  const credit = Math.max(0, Number(account?.creditAmount || 0));
  const due = parseQuoteDate(String(account?.nextDueDate || "").trim());
  if (credit > 0 && pending <= 0) {
    return { label: "Saldo a favor", tone: "credit", eta: "Disponible para compensar en otro evento", dueLabel: "Sin cobro urgente", sortWeight: 4 };
  }
  if (pending <= 0) {
    return { label: "Saldado", tone: "ok", eta: "Sin gestion pendiente", dueLabel: "Cuenta al dia", sortWeight: 5 };
  }
  if (!due) {
    return { label: "Sin fecha limite", tone: "neutral", eta: "Definir fecha maxima de pago", dueLabel: "Cobro sin programar", sortWeight: 3 };
  }
  const today = stripTime(new Date());
  const diffDays = Math.round((due.getTime() - today.getTime()) / 86400000);
  if (diffDays < 0) {
    const days = Math.abs(diffDays);
    return { label: `Vencido ${days} ${days === 1 ? "dia" : "dias"}`, tone: "overdue", eta: "Solicitar pago hoy", dueLabel: fmtDateShort(due), sortWeight: 0 };
  }
  if (diffDays === 0) {
    return { label: "Cobro hoy", tone: "due", eta: "Solicitar pago hoy", dueLabel: fmtDateShort(due), sortWeight: 1 };
  }
  if (diffDays <= 3) {
    return { label: `Por vencer ${diffDays} ${diffDays === 1 ? "dia" : "dias"}`, tone: "due", eta: `Solicitar en ${diffDays} ${diffDays === 1 ? "dia" : "dias"}`, dueLabel: fmtDateShort(due), sortWeight: 2 };
  }
  return { label: "En tiempo", tone: "ok", eta: `Solicitar en ${diffDays} dias`, dueLabel: fmtDateShort(due), sortWeight: 3 };
}

function getAccountingCompanyAccounts(rows = []) {
  const sourceRows = Array.isArray(rows) ? rows : [];
  const groups = new Map();
  const pickActionRow = (list) => list.slice().sort((a, b) => {
    const aPending = Number(a?.balancePending || 0) > 0 ? 0 : 1;
    const bPending = Number(b?.balancePending || 0) > 0 ? 0 : 1;
    if (aPending !== bPending) return aPending - bPending;
    const dueCmp = String(a?.dueDate || "9999-12-31").localeCompare(String(b?.dueDate || "9999-12-31"));
    if (dueCmp !== 0) return dueCmp;
    return String(b?.eventDate || "").localeCompare(String(a?.eventDate || ""));
  })[0] || null;

  for (const row of sourceRows) {
    const key = String(row?.companyId || row?.company || row?.refId || row?.event?.id || `sin_empresa_${groups.size + 1}`);
    const account = groups.get(key) || {
      key,
      companyId: String(row?.companyId || ""),
      companyName: String(row?.company || "Sin institucion").trim() || "Sin institucion",
      contactPhone: String(row?.manager || "").trim(),
      rows: [],
      netAmount: 0,
      collectedAmount: 0,
      pendingAmount: 0,
      creditAmount: 0,
      eventsCount: 0,
      pendingEventsCount: 0,
      paidEventsCount: 0,
      advancesCount: 0,
      sellerMap: new Map(),
      lastAdvanceDate: "",
      nextDueDate: "",
    };
    account.rows.push(row);
    account.netAmount += Math.max(0, Number(row?.total || 0));
    account.collectedAmount += Math.max(0, Number(row?.advancesTotal || 0));
    account.pendingAmount += Math.max(0, Number(row?.balancePending || 0));
    account.creditAmount += Math.max(0, Number(row?.creditBalance || 0));
    account.eventsCount += 1;
    account.advancesCount += Math.max(0, Number(row?.advancesCount || 0));
    if (Number(row?.balancePending || 0) > 0) account.pendingEventsCount += 1;
    else account.paidEventsCount += 1;
    if (row?.seller) {
      const seller = String(row.seller).trim();
      account.sellerMap.set(seller, Number(account.sellerMap.get(seller) || 0) + Math.max(0, Number(row?.total || 0)));
    }
    if (row?.lastAdvanceDate && (!account.lastAdvanceDate || String(row.lastAdvanceDate).localeCompare(account.lastAdvanceDate) > 0)) {
      account.lastAdvanceDate = String(row.lastAdvanceDate);
    }
    if (Number(row?.balancePending || 0) > 0 && row?.dueDate && (!account.nextDueDate || String(row.dueDate).localeCompare(account.nextDueDate) < 0)) {
      account.nextDueDate = String(row.dueDate);
    }
    groups.set(key, account);
  }

  return Array.from(groups.values()).map((account) => {
    const sellerEntry = Array.from(account.sellerMap.entries()).sort((a, b) => Number(b[1] || 0) - Number(a[1] || 0))[0];
    const collection = summarizeAccountingCollectionState(account);
    const actionRow = pickActionRow(account.rows);
    return {
      ...account,
      primarySeller: sellerEntry?.[0] || "",
      collectionLabel: collection.label,
      collectionTone: collection.tone,
      collectionEta: collection.eta,
      collectionDueLabel: collection.dueLabel,
      collectionSortWeight: collection.sortWeight,
      actionEventId: String(actionRow?.actionEventId || actionRow?.event?.id || ""),
      actionHasCredit: Math.max(0, Number(account.creditAmount || 0)) > 0,
    };
  }).sort((a, b) => {
    if (a.collectionSortWeight !== b.collectionSortWeight) return a.collectionSortWeight - b.collectionSortWeight;
    if (Number(b.pendingAmount || 0) !== Number(a.pendingAmount || 0)) return Number(b.pendingAmount || 0) - Number(a.pendingAmount || 0);
    return String(a.companyName || "").localeCompare(String(b.companyName || ""));
  });
}

function enrichAccountingReportRows(rows = []) {
  return (Array.isArray(rows) ? rows : []).map((row) => {
    const advances = normalizeQuoteAdvancesForSnapshot(row?.advances);
    const advancesSorted = advances.slice().sort((a, b) => {
      const d = String(a.date || "").localeCompare(String(b.date || ""));
      if (d !== 0) return d;
      return String(a.createdAt || "").localeCompare(String(b.createdAt || ""));
    });
    const lastAdvance = advancesSorted.length ? advancesSorted[advancesSorted.length - 1] : null;
    const advancesTotal = advancesSorted.reduce((acc, item) => acc + Math.max(0, Number(item.amount || 0)), 0);
    const total = Math.max(0, Number(row?.total || 0));
    const delta = total - advancesTotal;
    return {
      ...row,
      advances: advancesSorted,
      advancesCount: advancesSorted.length,
      advancesTotal,
      balancePending: Math.max(0, delta),
      creditBalance: Math.max(0, -delta),
      lastAdvanceDate: String(lastAdvance?.date || "").trim(),
      lastAdvancePaymentType: String(lastAdvance?.paymentType || "").trim(),
      lastAdvanceDescription: String(lastAdvance?.description || "").trim(),
    };
  });
}

function buildAccountingLedgerEntries(account) {
  const rows = Array.isArray(account?.rows) ? account.rows : [];
  const entries = [];
  for (const row of rows) {
    const eventId = String(row?.actionEventId || row?.event?.id || "").trim();
    const ref = String(row?.refId || eventId || "-").trim() || "-";
    const company = String(account?.companyName || row?.company || "").trim();
    const eventDate = String(row?.eventDate || "").trim();
    const docDate = String(row?.docDate || "").trim();
    const baseDate = docDate || eventDate;
    const total = Math.max(0, Number(row?.total || 0));
    const mainSalon = String(row?.mainSalon || row?.salon || "").trim();
    const salones = Array.isArray(row?.salones) ? row.salones.map((item) => String(item || "").trim()).filter(Boolean) : [];
    const otherSalones = salones.filter((item) => item !== mainSalon);
    const salonContext = [
      mainSalon ? `Salon principal: ${mainSalon}` : "",
      otherSalones.length ? `Salones del evento: ${otherSalones.join(", ")}` : "",
    ].filter(Boolean).join(" | ");
    if (total > 0) {
      entries.push({
        date: baseDate,
        type: "Cargo",
        eventId,
        refId: ref,
        concept: `Cotizacion ${ref}${company ? ` | ${company}` : ""}${salonContext ? ` | ${salonContext}` : ""}`,
        debit: total,
        credit: 0,
        sortBucket: 0,
      });
    }
    const advances = Array.isArray(row?.advances) ? row.advances : [];
    for (const adv of advances) {
      const amount = Math.max(0, Number(adv?.amount || 0));
      if (amount <= 0) continue;
      const advDate = String(adv?.date || baseDate || "").trim();
      const paymentType = String(adv?.paymentType || "").trim();
      const description = String(adv?.description || "").trim();
      const notes = [paymentType, description].filter(Boolean).join(" | ");
      entries.push({
        date: advDate,
        type: "Abono",
        eventId,
        refId: ref,
        concept: `${notes || `Pago aplicado a ${ref}`}${salonContext ? ` | ${salonContext}` : ""}`,
        debit: 0,
        credit: amount,
        sortBucket: 1,
      });
    }
  }

  const normalized = entries.slice().sort((a, b) => {
    const d = String(a.date || "9999-12-31").localeCompare(String(b.date || "9999-12-31"));
    if (d !== 0) return d;
    if (a.sortBucket !== b.sortBucket) return a.sortBucket - b.sortBucket;
    return String(a.refId || "").localeCompare(String(b.refId || ""));
  });
  let runningBalance = 0;
  return normalized.map((entry) => {
    runningBalance += Math.max(0, Number(entry.debit || 0));
    runningBalance -= Math.max(0, Number(entry.credit || 0));
    return { ...entry, runningBalance };
  });
}
function getAccountingReportFilteredRows() {
  const search = String(el.accountingReportSearch?.value || "").trim().toLowerCase();
  const from = String(el.accountingReportFrom?.value || "").trim();
  const to = String(el.accountingReportTo?.value || "").trim();
  const userId = String(el.accountingReportUser?.value || "").trim();
  const status = String(el.accountingReportStatus?.value || "").trim();
  const salon = String(el.accountingReportSalon?.value || "").trim();
  const company = String(el.accountingReportCompany?.value || "").trim();

  return enrichAccountingReportRows(buildSalesReportRows()).filter((r) => {
    if (from && r.eventDate && r.eventDate < from) return false;
    if (to && r.eventDate && r.eventDate > to) return false;
    if (userId && String(r.event?.userId || "") !== userId) return false;
    if (status && r.status !== status) return false;
    if (salon) {
      const eventSalones = Array.isArray(r.salones) ? r.salones : [];
      if (r.salon !== salon && !eventSalones.includes(salon)) return false;
    }
    if (company && String(r.event?.quote?.companyId || "") !== company) return false;
    if (search) {
      const blob = [r.refId, r.seller, r.eventType, r.salon, r.salonesLabel, r.company, r.manager, r.status, r.paymentType, r.lastAdvancePaymentType, r.lastAdvanceDescription].join(" ");
      if (!matchesLikeSearch(blob, search)) return false;
    }
    return true;
  }).sort((a, b) => {
    const d = String(a.eventDate || "").localeCompare(String(b.eventDate || ""));
    if (d !== 0) return d;
    return String(a.startTime || "").localeCompare(String(b.startTime || ""));
  });
}

function renderAccountingReportFilters() {
  const users = (state.users || []).filter((u) => u.active !== false);
  const statuses = Array.from(new Set((state.events || []).map((e) => String(e.status || "")).filter(Boolean))).sort();
  const salones = getUsedSalonesForReports();
  const companies = (state.companies || []).filter((c) => !isCompanyDisabled(c.id));

  const fillSelect = (node, rows, allLabel) => {
    if (!node) return;
    const previous = String(node.value || "");
    node.innerHTML = "";
    const all = document.createElement("option");
    all.value = "";
    all.textContent = allLabel;
    node.appendChild(all);
    for (const row of rows) {
      const opt = document.createElement("option");
      opt.value = String(row.value);
      opt.textContent = String(row.label);
      node.appendChild(opt);
    }
    if (previous && rows.some((r) => String(r.value) === previous)) node.value = previous;
    enhanceSelectControl(node, true);
  };

  fillSelect(el.accountingReportUser, users.map((u) => ({ value: u.id, label: u.fullName || u.name })), "Todos vendedores");
  fillSelect(el.accountingReportStatus, statuses.map((s) => ({ value: s, label: s })), "Todos estados");
  fillSelect(el.accountingReportSalon, salones.map((s) => ({ value: s, label: s })), "Todos salones");
  fillSelect(el.accountingReportCompany, companies.map((c) => ({ value: c.id, label: c.name })), "Todas instituciones");
}

function renderAccountingReportSummary(rows = []) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const accounts = getAccountingCompanyAccounts(safeRows);
  if (el.accountingReportFiltersMeta) {
    el.accountingReportFiltersMeta.textContent = accountingReportFiltersSummaryText();
  }
  if (!el.accountingReportSummary) return;
  const discountAmount = safeRows.reduce((acc, row) => acc + Math.max(0, Number(row?.discount || 0)), 0);
  const netAmount = safeRows.reduce((acc, row) => acc + Math.max(0, Number(row?.total || 0)), 0);
  const collectedAmount = safeRows.reduce((acc, row) => acc + Math.max(0, Number(row?.advancesTotal || 0)), 0);
  const pendingAmount = safeRows.reduce((acc, row) => acc + Math.max(0, Number(row?.balancePending || 0)), 0);
  const creditAmount = safeRows.reduce((acc, row) => acc + Math.max(0, Number(row?.creditBalance || 0)), 0);
  const overdueCount = accounts.filter((account) => account.collectionTone === "overdue").length;
  const nextCollection = accounts.find((account) => account.nextDueDate);
  const cards = [
    { tone: "primary", eyebrow: "Cartera", label: "Empresas visibles", value: String(accounts.length), meta: accounts.length ? `${safeRows.length} eventos | Venta neta ${moneyGT(netAmount)}` : "No hay resultados para el filtro actual" },
    { tone: "info", eyebrow: "Cobrado", label: "Anticipos cobrados", value: moneyGT(collectedAmount), meta: `${safeRows.reduce((acc, row) => acc + Math.max(0, Number(row?.advancesCount || 0)), 0)} pago(s) registrados` },
    { tone: "accent", eyebrow: "Pendiente", label: "Saldo por cobrar", value: moneyGT(pendingAmount), meta: `Saldo a favor ${moneyGT(creditAmount)} | Descuentos ${moneyGT(discountAmount)}` },
    { tone: overdueCount ? "accent" : "success", eyebrow: "Gestion", label: "Cobro sugerido", value: overdueCount ? `${overdueCount} vencida(s)` : "Al dia", meta: nextCollection ? `${nextCollection.companyName}: ${nextCollection.collectionEta}` : "Sin cobros pendientes programados" },
  ];
  el.accountingReportSummary.innerHTML = cards.map((card) => `
    <article class="salesSummaryCard salesSummaryCard--${escapeHtml(card.tone)}">
      <span class="salesSummaryEyebrow">${escapeHtml(card.eyebrow)}</span>
      <small>${escapeHtml(card.label)}</small>
      <strong>${escapeHtml(card.value)}</strong>
      <div class="salesSummaryMeta">${escapeHtml(card.meta)}</div>
    </article>
  `).join("");
}
function renderAccountingReportTable() {
  if (!el.accountingReportBody) return;
  const rows = getAccountingReportFilteredRows();
  renderAccountingReportSummary(rows);
  el.accountingReportBody.innerHTML = "";
  const accounts = getAccountingCompanyAccounts(rows);
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="11">Sin resultados para los filtros seleccionados.</td>`;
    el.accountingReportBody.appendChild(tr);
    return;
  }
  const pick = (obj, key) => obj?.[key] || { qty: 0, amount: 0 };
  for (const account of accounts) {
    const expanded = accountingReportExpandedAccounts.has(account.key);
    const canApplyPayment = !!String(account.actionEventId || "").trim();
    const tr = document.createElement("tr");
    tr.className = "accountingCompanyRow";
    tr.innerHTML = `
      <td><span class="accountingIndicator accountingIndicator--${escapeHtml(account.collectionTone)}">${escapeHtml(account.collectionLabel)}</span></td>
      <td>
        <div class="accountingCompanyCell">
          <strong>${escapeHtml(account.companyName)}</strong>
          <small>${escapeHtml(account.primarySeller ? `Principal: ${account.primarySeller}` : "Sin vendedor asignado")}</small>
        </div>
      </td>
      <td>
        <div class="accountingMetaCell">
          <strong>${escapeHtml(account.contactPhone || "-")}</strong>
          <small>${escapeHtml(`${account.pendingEventsCount} pendiente(s) | ${account.paidEventsCount} al dia`)}</small>
        </div>
      </td>
      <td>${escapeHtml(String(account.eventsCount || 0))}</td>
      <td>${escapeHtml(moneyGT(account.netAmount || 0))}</td>
      <td>${escapeHtml(moneyGT(account.collectedAmount || 0))}</td>
      <td>${escapeHtml(moneyGT(account.pendingAmount || 0))}</td>
      <td>${escapeHtml(moneyGT(account.creditAmount || 0))}</td>
      <td>
        <div class="accountingMetaCell">
          <strong>${escapeHtml(account.collectionDueLabel || "-")}</strong>
          <small>${escapeHtml(account.collectionEta || "Sin gestion pendiente")}</small>
        </div>
      </td>
      <td>${escapeHtml(account.lastAdvanceDate || "-")}</td>
      <td>
        <div class="accountingActionGroup">
          <button class="accountingActionBtn accountingActionBtn--secondary" type="button" data-accounting-toggle-company="${escapeHtml(account.key)}">${escapeHtml(expanded ? "Ocultar detalle" : "Ver detalle")}</button>
          <button class="accountingActionBtn accountingActionBtn--secondary" type="button" data-accounting-open-statement="${escapeHtml(account.key)}">Estado de cuenta</button>
          <button class="accountingActionBtn accountingActionBtn--secondary" type="button" data-accounting-open-quote="${escapeHtml(account.actionEventId || "")}" ${canApplyPayment ? "" : "disabled"}>Ver cotizacion</button>
          <button class="accountingActionBtn" type="button" data-accounting-apply-payment="${escapeHtml(account.actionEventId || "")}" ${canApplyPayment ? "" : "disabled"}>${escapeHtml(account.actionHasCredit ? "Ajustar saldo" : "Aplicar pago")}</button>
        </div>
      </td>
    `;
    el.accountingReportBody.appendChild(tr);

    if (!expanded) continue;
    const ledgerEntries = buildAccountingLedgerEntries(account);
    const detailTr = document.createElement("tr");
    detailTr.className = "accountingCompanyDetailRow";
    detailTr.innerHTML = `
      <td colspan="11">
        <div class="accountingDetailCard">
          <div class="accountingDetailHead">
            <strong>Estado de cuenta por evento</strong>
            <small>${escapeHtml(`${account.companyName} | ${account.eventsCount} evento(s) | ${account.advancesCount} pago(s)`)}</small>
          </div>
          <div class="accountingDetailTableWrap">
            <table class="quoteTable accountingDetailTable">
              <thead>
                <tr>
                  <th>Estado</th>
                  <th>No cotizacion</th>
                  <th>Fecha evento</th>
                  <th>Fecha maxima pago</th>
                  <th>Salon principal</th>
                  <th>Salones del evento</th>
                  <th>Vendedor</th>
                  <th>Forma pago</th>
                  <th>Total neto</th>
                  <th>Cobrado</th>
                  <th>Pendiente</th>
                  <th>Saldo a favor</th>
                  <th>Ultimo deposito</th>
                  <th>Detalle</th>
                  <th>Accion</th>
                </tr>
              </thead>
              <tbody>
                ${account.rows.map((r) => `
                  <tr>
                    <td><span class="salesStatusBadge" style="background:${escapeHtml(hexToRgba(r.statusColor, 0.25))};border-color:${escapeHtml(hexToRgba(r.statusColor, 0.6))}">${escapeHtml(r.status || "-")}</span></td>
                    <td>${escapeHtml(r.refId || "-")}</td>
                    <td>${escapeHtml(r.eventDate || "-")}</td>
                    <td>${escapeHtml(r.dueDate || "-")}</td>
                    <td>${escapeHtml(r.mainSalon || r.salon || "-")}</td>
                    <td>${escapeHtml(r.salonesLabel || r.salon || "-")}</td>
                    <td>${escapeHtml(r.seller || "-")}</td>
                    <td>${escapeHtml(r.paymentType || "-")}</td>
                    <td>${escapeHtml(moneyGT(r.total || 0))}</td>
                    <td>${escapeHtml(moneyGT(r.advancesTotal || 0))}</td>
                    <td>${escapeHtml(moneyGT(r.balancePending || 0))}</td>
                    <td>${escapeHtml(moneyGT(r.creditBalance || 0))}</td>
                    <td>${escapeHtml(r.lastAdvanceDate || "-")}</td>
                    <td>
                      <div class="accountingMetaCell">
                        <strong>${escapeHtml(r.lastAdvancePaymentType || "-")}</strong>
                        <small>${escapeHtml(r.lastAdvanceDescription || "Sin anticipo registrado")}</small>
                      </div>
                    </td>
                    <td><div class="accountingActionGroup"><button class="accountingActionBtn accountingActionBtn--secondary" type="button" data-accounting-open-quote="${escapeHtml(String(r?.actionEventId || r?.event?.id || ""))}" ${String(r?.actionEventId || r?.event?.id || "").trim() ? "" : "disabled"}>Ver cotizacion</button><button class="accountingActionBtn" type="button" data-accounting-apply-payment="${escapeHtml(String(r?.actionEventId || r?.event?.id || ""))}" ${String(r?.actionEventId || r?.event?.id || "").trim() ? "" : "disabled"}>${escapeHtml(r.creditBalance > 0 ? "Ajustar saldo" : "Aplicar pago")}</button></div></td>
                  </tr>
                `).join("")}
              </tbody>
            </table>
          </div>
          <div class="accountingDetailMetrics">
            <span>Total A&B ${escapeHtml(moneyGT(account.rows.reduce((acc, row) => acc + Math.max(0, Number(pick(row.catBuckets, "alimentosBebidas").amount || 0)), 0)))}</span>
            <span>Hospedaje JDL ${escapeHtml(moneyGT(account.rows.reduce((acc, row) => acc + Math.max(0, Number(pick(row.catBuckets, "hospedajeJdl").amount || 0)), 0)))}</span>
            <span>Hospedaje terceros ${escapeHtml(moneyGT(account.rows.reduce((acc, row) => acc + Math.max(0, Number(pick(row.catBuckets, "hospedajeTerceros").amount || 0)), 0)))}</span>
            <span>Miscelaneos ${escapeHtml(moneyGT(account.rows.reduce((acc, row) => acc + Math.max(0, Number(pick(row.catBuckets, "miscelaneos").amount || 0)), 0)))}</span>
          </div>
          <div class="accountingDetailTableWrap">
            <table class="quoteTable accountingDetailTable">
              <thead>
                <tr>
                  <th>Fecha</th>
                  <th>Tipo</th>
                  <th>No cotizacion</th>
                  <th>Concepto</th>
                  <th>Cargo</th>
                  <th>Abono</th>
                  <th>Saldo acumulado</th>
                  <th>Accion</th>
                </tr>
              </thead>
              <tbody>
                ${ledgerEntries.length ? ledgerEntries.map((entry) => `
                  <tr>
                    <td>${escapeHtml(entry.date || "-")}</td>
                    <td>${escapeHtml(entry.type || "-")}</td>
                    <td>${escapeHtml(entry.refId || "-")}</td>
                    <td>${escapeHtml(entry.concept || "-")}</td>
                    <td>${escapeHtml(entry.debit > 0 ? moneyGT(entry.debit) : "-")}</td>
                    <td>${escapeHtml(entry.credit > 0 ? moneyGT(entry.credit) : "-")}</td>
                    <td>${escapeHtml(moneyGT(entry.runningBalance || 0))}</td>
                    <td><div class="accountingActionGroup"><button class="accountingActionBtn accountingActionBtn--secondary" type="button" data-accounting-open-quote="${escapeHtml(String(entry.eventId || ""))}" ${String(entry.eventId || "").trim() ? "" : "disabled"}>Ver cotizacion</button><button class="accountingActionBtn" type="button" data-accounting-apply-payment="${escapeHtml(String(entry.eventId || ""))}" ${String(entry.eventId || "").trim() ? "" : "disabled"}>${escapeHtml(entry.runningBalance > 0 ? "Aplicar pago" : "Ajustar saldo")}</button></div></td>
                  </tr>
                `).join("") : `<tr><td colspan="8">Sin movimientos para esta institucion.</td></tr>`}
              </tbody>
            </table>
          </div>
        </div>
      </td>
    `;
    el.accountingReportBody.appendChild(detailTr);
  }
}

function renderAccountingStatementModal(account) {
  if (!el.accountStatementCompany || !el.accountStatementMeta || !el.accountStatementSummary || !el.accountStatementBody) return;
  const safe = account || null;
  const rows = Array.isArray(safe?.rows) ? safe.rows : [];
  const ledgerEntries = buildAccountingLedgerEntries(safe);
  const now = new Date();
  const issueDate = now.toLocaleDateString("es-GT", { day: "2-digit", month: "2-digit", year: "numeric" });
  const issueTime = now.toLocaleTimeString("es-GT", { hour: "2-digit", minute: "2-digit" });
  const company = String(safe?.companyName || "Institucion").trim() || "Institucion";
  const pending = Math.max(0, Number(safe?.pendingAmount || 0));
  const credit = Math.max(0, Number(safe?.creditAmount || 0));
  const total = Math.max(0, Number(safe?.netAmount || 0));
  const collected = Math.max(0, Number(safe?.collectedAmount || 0));

  const firstRow = rows[0] || {};
  const firstEvent = firstRow?.event || {};
  const quote = firstEvent?.quote || {};
  const managerName = String(quote?.contact || "-").trim() || "-";
  const phone = String(quote?.phone || firstRow?.manager || "-").trim() || "-";
  const address = String(quote?.address || "-").trim() || "-";
  const nit = String(quote?.nit || "-").trim() || "-";

  el.accountStatementCompany.textContent = company;
  if (el.accountStatementNit) el.accountStatementNit.textContent = nit;
  if (el.accountStatementManager) el.accountStatementManager.textContent = managerName;
  if (el.accountStatementPhone) el.accountStatementPhone.textContent = phone;
  if (el.accountStatementAddress) el.accountStatementAddress.textContent = address;
  if (el.accountStatementIssueDate) el.accountStatementIssueDate.textContent = issueDate;
  if (el.accountStatementIssueTime) el.accountStatementIssueTime.textContent = issueTime;

  el.accountStatementMeta.textContent = `${rows.length} evento(s) financiero(s) | ${ledgerEntries.length} movimiento(s)`;
  el.accountStatementSummary.innerHTML = `
    <article class="accountStatementKpi"><small>Total neto</small><strong>${escapeHtml(moneyGT(total))}</strong></article>
    <article class="accountStatementKpi"><small>Total pagado</small><strong>${escapeHtml(moneyGT(collected))}</strong></article>
    <article class="accountStatementKpi"><small>Saldo pendiente</small><strong>${escapeHtml(moneyGT(pending))}</strong></article>
    <article class="accountStatementKpi"><small>Saldo a favor</small><strong>${escapeHtml(moneyGT(credit))}</strong></article>
  `;

  if (!ledgerEntries.length) {
    el.accountStatementBody.innerHTML = '<tr><td colspan="7">Sin movimientos para esta empresa.</td></tr>';
    return;
  }

  el.accountStatementBody.innerHTML = ledgerEntries.map((entry) => `
    <tr>
      <td>${escapeHtml(entry.date || "-")}</td>
      <td>${escapeHtml(entry.refId || "-")}</td>
      <td>${escapeHtml(entry.concept || "-")}</td>
      <td>${escapeHtml(entry.debit > 0 ? moneyGT(entry.debit) : "-")}</td>
      <td>${escapeHtml(entry.credit > 0 ? moneyGT(entry.credit) : "-")}</td>
      <td>${escapeHtml(moneyGT(entry.runningBalance || 0))}</td>
      <td><div class="accountingActionGroup"><button class="accountingActionBtn accountingActionBtn--secondary" type="button" data-accounting-open-quote="${escapeHtml(String(entry.eventId || ""))}" ${String(entry.eventId || "").trim() ? "" : "disabled"}>Ver cotizacion</button><button class="accountingActionBtn accountingActionBtn--secondary" type="button" data-accounting-apply-payment="${escapeHtml(String(entry.eventId || ""))}" ${String(entry.eventId || "").trim() ? "" : "disabled"}>${escapeHtml(entry.runningBalance > 0 ? "Aplicar pago" : "Ajustar")}</button></div></td>
    </tr>
  `).join("");
}

function openAccountingStatementModal(companyKey) {
  if (!el.accountStatementBackdrop) return;
  const key = String(companyKey || "").trim();
  if (!key) return toast("No se pudo identificar la empresa.");
  const rows = getAccountingReportFilteredRows();
  const accounts = getAccountingCompanyAccounts(rows);
  const account = accounts.find((x) => String(x.key || "") === key);
  if (!account) return toast("No se encontro estado de cuenta para esa empresa.");
  accountingStatementActiveCompanyKey = key;
  renderAccountingStatementModal(account);
  el.accountStatementBackdrop.hidden = false;
}

function closeAccountingStatementModal() {
  if (!el.accountStatementBackdrop) return;
  el.accountStatementBackdrop.hidden = true;
}
function exportAccountingReportToExcel() {
  const rows = getAccountingReportFilteredRows();
  if (!rows.length) return toast("No hay datos para exportar.");
  const accounts = getAccountingCompanyAccounts(rows);
  const generatedAt = new Date().toLocaleString("es-GT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
  const emittedBy = String(authSession.fullName || authSession.username || "Sistema").trim();
  const netAmount = rows.reduce((acc, row) => acc + Math.max(0, Number(row?.total || 0)), 0);
  const collectedAmount = rows.reduce((acc, row) => acc + Math.max(0, Number(row?.advancesTotal || 0)), 0);
  const pendingAmount = rows.reduce((acc, row) => acc + Math.max(0, Number(row?.balancePending || 0)), 0);
  const creditAmount = rows.reduce((acc, row) => acc + Math.max(0, Number(row?.creditBalance || 0)), 0);
  const overdueCount = accounts.filter((account) => account.collectionTone === "overdue").length;

  const htmlRows = accounts.map((account) => `
    <tr>
      <td>${escapeHtml(account.collectionLabel || "-")}</td>
      <td>${escapeHtml(account.companyName || "-")}</td>
      <td>${escapeHtml(account.contactPhone || "-")}</td>
      <td>${escapeHtml(account.primarySeller || "-")}</td>
      <td>${escapeHtml(String(account.eventsCount || 0))}</td>
      <td>${escapeHtml(String(account.pendingEventsCount || 0))}</td>
      <td>${escapeHtml(moneyGT(account.netAmount || 0))}</td>
      <td>${escapeHtml(moneyGT(account.collectedAmount || 0))}</td>
      <td>${escapeHtml(moneyGT(account.pendingAmount || 0))}</td>
      <td>${escapeHtml(moneyGT(account.creditAmount || 0))}</td>
      <td>${escapeHtml(account.collectionDueLabel || "-")}</td>
      <td>${escapeHtml(account.collectionEta || "-")}</td>
      <td>${escapeHtml(account.lastAdvanceDate || "-")}</td>
    </tr>
  `).join("");

  const htmlDetailRows = rows.slice().sort((a, b) => {
    const c = String(a.company || "").localeCompare(String(b.company || ""));
    if (c !== 0) return c;
    const d = String(a.refId || "").localeCompare(String(b.refId || ""));
    if (d !== 0) return d;
    return String(a.eventDate || "").localeCompare(String(b.eventDate || ""));
  }).map((row) => `
    <tr>
      <td>${escapeHtml(row.company || "-")}</td>
      <td>${escapeHtml(row.refId || "-")}</td>
      <td>${escapeHtml(row.eventDate || "-")}</td>
      <td>${escapeHtml(row.dueDate || "-")}</td>
      <td>${escapeHtml(row.seller || "-")}</td>
      <td>${escapeHtml(moneyGT(row.total || 0))}</td>
      <td>${escapeHtml(moneyGT(row.advancesTotal || 0))}</td>
      <td>${escapeHtml(moneyGT(row.balancePending || 0))}</td>
      <td>${escapeHtml(moneyGT(row.creditBalance || 0))}</td>
      <td>${escapeHtml(row.lastAdvanceDate || "-")}</td>
    </tr>
  `).join("");

  const html = `<!doctype html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="utf-8" />
  <meta name="ProgId" content="Excel.Sheet" />
  <meta name="Generator" content="CRM Jardines" />
  <style>
    body{ font-family: Calibri, Arial, sans-serif; background:#eef3fb; margin:0; padding:16px; color:#0f172a; }
    .card{ background:#ffffff; border:1px solid #c5d4ea; border-radius:10px; overflow:hidden; }
    .meta{ padding:10px 14px; border-top:1px solid #bfd3ee; border-bottom:1px solid #bfd3ee; background:#eaf3ff; font-size:12px; }
    .meta div{ margin:2px 0; }
    table{ width:100%; border-collapse:collapse; }
    th,td{ border:1px solid #c7d5ea; padding:6px 7px; font-size:10.5px; white-space:nowrap; }
    thead th{ background:#0f3c67; color:#fff; font-weight:700; text-transform:uppercase; }
    .titleTable{ width:100%; border-collapse:collapse; }
    .titleCell{ border:1px solid #c7d5ea; background:#d8e3f3; color:#000; font-weight:800; font-size:20px; letter-spacing:.3px; padding:12px 14px; text-transform:uppercase; }
    .sectionTitle{ background:#f1f6ff; color:#0f3c67; font-weight:700; border:1px solid #c7d5ea; padding:8px 10px; margin-top:14px; }
  </style>
</head>
<body>
  <div class="card">
    <table class="titleTable">
      <tr><td class="titleCell">CRM JARDINES - ESTADO DE CUENTA POR INSTITUCION</td></tr>
    </table>
    <div class="meta">
      <div><b>Fecha:</b> ${escapeHtml(generatedAt)}</div>
      <div><b>Quien emitio el reporte:</b> ${escapeHtml(emittedBy)}</div>
      <div><b>Filtros aplicados:</b> ${escapeHtml(accountingReportFiltersSummaryText())}</div>
      <div><b>Total empresas:</b> ${accounts.length}</div>
      <div><b>Total eventos:</b> ${rows.length}</div>
      <div><b>Anticipos cobrados:</b> ${escapeHtml(moneyGT(collectedAmount))}</div>
      <div><b>Saldo pendiente:</b> ${escapeHtml(moneyGT(pendingAmount))}</div>
      <div><b>Saldo a favor:</b> ${escapeHtml(moneyGT(creditAmount))}</div>
      <div><b>Venta neta:</b> ${escapeHtml(moneyGT(netAmount))}</div>
      <div><b>Empresas vencidas:</b> ${overdueCount}</div>
    </div>
    <div class="sectionTitle">Resumen por institucion</div>
    <table class="quoteItemsTablePrint">
      <colgroup>
        <col class="qtyCol" />
        <col class="descCol" />
        <col class="priceCol" />
        <col class="totalCol" />
      </colgroup>
      <thead>
        <tr>
          <th>Indicador</th><th>Institucion</th><th>Contacto cobro</th><th>Vendedor principal</th><th>Eventos</th><th>Eventos pendientes</th><th>Venta neta</th><th>Cobrado</th><th>Saldo pendiente</th><th>Saldo a favor</th><th>Proximo cobro</th><th>Tiempo sugerido</th><th>Ultimo deposito</th>
        </tr>
      </thead>
      <tbody>${htmlRows}</tbody>
    </table>
    <div class="sectionTitle">Detalle por numero de cotizacion</div>
    <table>
      <thead>
        <tr>
          <th>Institucion</th><th>No cotizacion</th><th>Fecha evento</th><th>Fecha maxima pago</th><th>Vendedor</th><th>Total neto</th><th>Cobrado</th><th>Saldo pendiente</th><th>Saldo a favor</th><th>Ultimo deposito</th>
        </tr>
      </thead>
      <tbody>${htmlDetailRows}</tbody>
    </table>
  </div>
</body>
</html>`;

  const blob = new Blob([`?${html}`], { type: "application/vnd.ms-excel;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  const stamp = new Date().toISOString().slice(0, 10);
  a.href = url;
  a.download = `estado_cuenta_instituciones_${stamp}.xls`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}
function resetAccountingReportFilters() {
  if (el.accountingReportSearch) el.accountingReportSearch.value = "";
  if (el.accountingReportFrom) el.accountingReportFrom.value = "";
  if (el.accountingReportTo) el.accountingReportTo.value = "";
  if (el.accountingReportUser) el.accountingReportUser.value = "";
  if (el.accountingReportStatus) el.accountingReportStatus.value = "";
  if (el.accountingReportSalon) el.accountingReportSalon.value = "";
  if (el.accountingReportCompany) el.accountingReportCompany.value = "";
}

function openAccountingReportModal() {
  if (!el.accountingReportBackdrop) return;
  if (el.salesReportBackdrop) el.salesReportBackdrop.hidden = true;
  if (el.occupancyReportBackdrop) el.occupancyReportBackdrop.hidden = true;
  if (el.dashboardReportBackdrop) el.dashboardReportBackdrop.hidden = true;
  if (el.institutionReportBackdrop) el.institutionReportBackdrop.hidden = true;
  renderAccountingReportFilters();
  resetAccountingReportFilters();
  renderAccountingReportTable();
  el.accountingReportBackdrop.hidden = false;
}

function closeAccountingReportModal() {
  if (!el.accountingReportBackdrop) return;
  accountingReportReturnAfterPayment = false;
  accountingReturnAfterPaymentTarget = "";
  document.body.classList.remove("accountingPaymentMode");
  el.accountingReportBackdrop.hidden = true;
  closeAccountingStatementModal();
  restoreModuleScreenAfterModal();
}
function weekInputFromDate(date) {
  return toISODate(startOfWeek(date));
}

function mondayFromWeekInput(value) {
  const raw = String(value || "").trim();
  if (!raw) return startOfWeek(new Date());
  const dateMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (dateMatch) {
    const parsed = new Date(`${raw}T00:00:00`);
    return Number.isNaN(parsed.getTime()) ? startOfWeek(new Date()) : startOfWeek(parsed);
  }
  const m = String(value || "").match(/^(\d{4})-W(\d{2})$/);
  if (!m) return startOfWeek(new Date());
  const year = Number(m[1]);
  const week = Number(m[2]);
  const jan4 = new Date(year, 0, 4);
  const jan4Day = (jan4.getDay() + 6) % 7;
  const monday = new Date(jan4);
  monday.setDate(jan4.getDate() - jan4Day + (week - 1) * 7);
  monday.setHours(0, 0, 0, 0);
  return monday;
}

function getOccupancyWeekRange() {
  const monday = mondayFromWeekInput(el.occupancyReportWeek?.value || "");
  const sunday = addDays(monday, 6);
  return { monday, sunday };
}

function updateOccupancyReportWeekUi() {
  const { monday, sunday } = getOccupancyWeekRange();
  occupancySelectedDayIso = toISODate(monday);
  if (el.occupancyReportWeek) {
    el.occupancyReportWeek.value = weekInputFromDate(monday);
  }
  if (el.occupancyReportSubtitle) {
    el.occupancyReportSubtitle.textContent = `Semana ${toISODate(monday)} a ${toISODate(sunday)} (Lunes a Domingo)`;
  }
  renderOccupancyReportTable();
}

function moveOccupancyReportWeek(deltaWeeks = 0) {
  const { monday } = getOccupancyWeekRange();
  const nextMonday = addDays(monday, Number(deltaWeeks || 0) * 7);
  if (el.occupancyReportWeek) {
    el.occupancyReportWeek.value = weekInputFromDate(nextMonday);
  }
  updateOccupancyReportWeekUi();
}

function getLatestQuoteSnapshotFromSeries(series) {
  const snapshots = [];
  for (const ev of Array.isArray(series) ? series : []) {
    const latest = getLatestQuoteSnapshotForEvent(ev);
    if (latest) snapshots.push(latest);
  }
  if (!snapshots.length) return null;
  snapshots.sort((a, b) => {
    const verDiff = Number(b.version || 0) - Number(a.version || 0);
    if (verDiff !== 0) return verDiff;
    const ta = new Date(a.quotedAt || 0).getTime() || 0;
    const tb = new Date(b.quotedAt || 0).getTime() || 0;
    return tb - ta;
  });
  return snapshots[0];
}

function getLatestMenuMontajeSnapshotFromSeries(series) {
  const candidates = [];
  for (const ev of Array.isArray(series) ? series : []) {
    const snap = getLatestQuoteSnapshotForEvent(ev);
    const mmVersions = normalizeMenuMontajeVersionHistory(snap?.menuMontajeVersions);
    const currentMmVersion = Math.max(1, Number(snap?.menuMontajeVersion || mmVersions[mmVersions.length - 1]?.version || 1));
    const versionSnap = mmVersions.find((v) => Number(v.version) === currentMmVersion)
      || mmVersions[mmVersions.length - 1]
      || null;
    const entries = Array.isArray(versionSnap?.entries) ? versionSnap.entries : (Array.isArray(snap?.menuMontajeEntries) ? snap.menuMontajeEntries : []);
    if (!entries.length) continue;
    const latestEntryAt = String(versionSnap?.savedAt || "").trim() || entries.reduce((maxIso, row) => {
      const iso = String(row?.updatedAt || "").trim();
      return iso > maxIso ? iso : maxIso;
    }, "");
    const mmVersion = Number(versionSnap?.version || currentMmVersion || 1);
    candidates.push({ snap, latestEntryAt, mmVersion, entries });
  }
  if (!candidates.length) return null;
  candidates.sort((a, b) => {
    const aKey = String(a.latestEntryAt || a.snap?.quotedAt || "");
    const bKey = String(b.latestEntryAt || b.snap?.quotedAt || "");
    if (aKey !== bKey) return bKey.localeCompare(aKey);
    const verDiff = Number(b.snap?.version || 0) - Number(a.snap?.version || 0);
    if (verDiff !== 0) return verDiff;
    const ta = new Date(a.snap?.quotedAt || 0).getTime() || 0;
    const tb = new Date(b.snap?.quotedAt || 0).getTime() || 0;
    return tb - ta;
  });
  return candidates[0];
}

function formatQuoteSentAtLabel(isoText) {
  const raw = String(isoText || "").trim();
  if (!raw) return "";
  const d = new Date(raw);
  if (Number.isNaN(d.getTime())) return raw;
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  const hh = String(d.getHours()).padStart(2, "0");
  const mi = String(d.getMinutes()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd} ${hh}:${mi}`;
}

function buildOccupancyQuoteActionHtml(row) {
  if (!row?.hasQuote) return `<span class="occupancyQuoteEmpty">-</span>`;
  const version = Number(row.latestQuoteVersion || 0);
  const versionLabel = version > 0 ? `V${version}` : "Ver";
  const sentLabel = formatQuoteSentAtLabel(row.latestQuoteSentAt);
  const text = sentLabel ? `${versionLabel} - ${sentLabel}` : versionLabel;
  return `<button type="button" class="occupancyQuoteLinkBtn" data-event-id="${escapeHtml(String(row.eventId || ""))}" data-quote-version="${escapeHtml(String(version || ""))}">${escapeHtml(text)}</button>`;
}

function buildOccupancyMenuMontajeActionHtml(row) {
  if (!row?.hasMenuMontajeReport) return `<span class="occupancyQuoteEmpty">-</span>`;
  const version = Number(row.latestMenuMontajeVersion || 0);
  const versionLabel = version > 0 ? `V${version}` : "Ver";
  const sentLabel = formatQuoteSentAtLabel(row.latestMenuMontajeAt);
  const text = sentLabel ? `${versionLabel} - ${sentLabel}` : versionLabel;
  return `<button type="button" class="occupancyQuoteLinkBtn occupancyMenuMontajeLinkBtn" data-event-id="${escapeHtml(String(row.eventId || ""))}" data-quote-version="${escapeHtml(String(version || ""))}">${escapeHtml(text)}</button>`;
}

function buildOccupancyChecklistActionHtml(row) {
  const eventId = String(row?.eventId || "").trim();
  if (!eventId) return `<span class="occupancyQuoteEmpty">-</span>`;
  const completed = row?.checklistCompleted === true;
  const hasChecklist = row?.hasChecklist === true;
  const label = completed ? "Completo" : (hasChecklist ? "En proceso" : "Iniciar");
  const cls = hasChecklist
    ? "occupancyChecklistLinkBtn occupancyChecklistLinkBtn--has"
    : "occupancyChecklistLinkBtn occupancyChecklistLinkBtn--missing";
  return `<button type="button" class="${cls}" data-event-id="${escapeHtml(eventId)}"><span class="occupancyChecklistActionLabel">${escapeHtml(label)}</span><span class="occupancyChecklistActionIcon" aria-hidden="true">&#8250;</span></button>`;
}

async function openOccupancyQuoteByRow(eventId, versionRaw = "") {
  const id = String(eventId || "").trim();
  if (!id) return;
  const ev = (state.events || []).find((x) => String(x.id) === id);
  if (!ev?.quote) return toast("Este evento no tiene cotizacion.");
  const versions = normalizeQuoteVersionHistory(ev.quote.versions);
  const requestedVersion = Number(versionRaw || 0);
  let snapshot = null;
  if (Number.isFinite(requestedVersion) && requestedVersion > 0) {
    snapshot = versions.find((v) => Number(v.version || 0) === requestedVersion) || null;
  }
  if (!snapshot) snapshot = getLatestQuoteSnapshotForEvent(ev);
  if (!snapshot) return toast("No se encontro una version de cotizacion para abrir.");
  try {
    await openQuoteDocument(ev, snapshot);
  } catch (_) {
    console.error("No se pudo abrir la cotizacion desde reporte de ocupacion:", _);
    toast("No se pudo abrir la cotizacion.");
  }
}

function openOccupancyMenuMontajeByRow(eventId, versionRaw = "") {
  const id = String(eventId || "").trim();
  if (!id) return;
  const ev = (state.events || []).find((x) => String(x.id) === id);
  if (!ev?.quote) return toast("Este evento no tiene informe de Menu & Montaje.");
  const requestedVersion = Number(versionRaw || 0);
  const series = getEventSeries(ev);
  const latestMeta = getLatestMenuMontajeSnapshotFromSeries(series);
  let snapshot = latestMeta?.snap || getLatestQuoteSnapshotForEvent(ev) || deepClone(ev.quote);
  if (!snapshot) return toast("No se encontro un informe de Menu & Montaje.");
  const mmVersions = normalizeMenuMontajeVersionHistory(snapshot.menuMontajeVersions);
  const targetVersion = Number.isFinite(requestedVersion) && requestedVersion > 0
    ? requestedVersion
    : Math.max(1, Number(snapshot.menuMontajeVersion || latestMeta?.mmVersion || mmVersions[mmVersions.length - 1]?.version || 1));
  const mmSnap = mmVersions.find((v) => Number(v.version) === targetVersion)
    || mmVersions[mmVersions.length - 1]
    || null;
  snapshot = { ...deepClone(snapshot), menuMontajeVersion: targetVersion, menuMontajeEntries: Array.isArray(mmSnap?.entries) ? deepClone(mmSnap.entries) : [] };
  const entries = Array.isArray(snapshot.menuMontajeEntries) ? snapshot.menuMontajeEntries : [];
  if (!entries.length) return toast("La version seleccionada no contiene informe de Menu & Montaje.");
  openMenuMontajeReportDocument(ev, snapshot);
}

function openEventChecklistByRow(eventId) {
  const id = String(eventId || "").trim();
  if (!id) return;
  openEventChecklistModal(id);
}

function buildOccupancyReportRows() {
  const allowed = new Set([STATUS.PRERESERVA, STATUS.CONFIRMADO]);
  const { monday, sunday } = getOccupancyWeekRange();
  const fromIso = toISODate(monday);
  const toIso = toISODate(sunday);
  const rows = [];
  const metricsByReservation = new Map();
  for (const ev of state.events || []) {
    const eventDate = String(ev.date || "");
    if (!eventDate || eventDate < fromIso || eventDate > toIso) continue;
    if (!allowed.has(String(ev.status || ""))) continue;
    const reservationKey = reservationKeyFromEvent(ev);
    if (!metricsByReservation.has(reservationKey)) {
      const series = getEventSeries(ev);
      const dateSet = new Set();
      const salonDaySet = new Set();
      for (const s of series) {
        const d = String(s?.date || "").trim();
        const salon = String(s?.salon || "").trim();
        if (d) dateSet.add(d);
        if (d && salon) salonDaySet.add(`${d}|${salon}`);
      }
      const latestQuoteFromSeries = getLatestQuoteSnapshotFromSeries(series) || getLatestQuoteSnapshotForEvent(ev);
      const latestMenuMontaje = getLatestMenuMontajeSnapshotFromSeries(series);
      const totalsFromSeries = getQuoteTotals(latestQuoteFromSeries || ev.quote || {});
      const totalEvent = Math.max(0, Number(totalsFromSeries.total || 0));
      const days = Math.max(1, dateSet.size || 1);
      const salonDayUnits = Math.max(1, salonDaySet.size || 1);
      metricsByReservation.set(reservationKey, {
        reservationKey,
        days,
        salonDayUnits,
        totalEvent,
        incomePerDay: totalEvent / days,
        incomePerSalonDay: totalEvent / salonDayUnits,
        latestQuote: latestQuoteFromSeries || null,
        latestMenuMontajeSnap: latestMenuMontaje?.snap || null,
        latestMenuMontajeVersion: Number(latestMenuMontaje?.mmVersion || 0),
        latestMenuMontajeEntries: Array.isArray(latestMenuMontaje?.entries) ? latestMenuMontaje.entries : [],
        latestMenuMontajeAt: String(latestMenuMontaje?.latestEntryAt || ""),
      });
    }
    const metrics = metricsByReservation.get(reservationKey);
    const quote = ev.quote || null;
    const latestQuote = metrics?.latestQuote || getLatestQuoteSnapshotForEvent(ev);
    const latestQuoteVersion = Number(latestQuote?.version || 0);
    const latestMenuMontajeSnap = metrics?.latestMenuMontajeSnap || null;
    const latestMenuMontajeVersion = Number(metrics?.latestMenuMontajeVersion || 0);
    const latestMenuMontajeEntries = Array.isArray(metrics?.latestMenuMontajeEntries) ? metrics.latestMenuMontajeEntries : [];
    const user = (state.users || []).find((u) => String(u.id) === String(ev.userId));
    const company = quote?.companyId ? (state.companies || []).find((c) => String(c.id) === String(quote.companyId)) : null;
    const manager = company?.managers?.find((m) => String(m.id) === String(quote?.managerId));
    const totals = getQuoteTotals(quote || {});
    const checklistMeta = getEventChecklistMeta(ev.id);
    rows.push({
      eventId: String(ev.id || ""),
      status: String(ev.status || ""),
      statusColor: statusColor(ev.status),
      refId: String(quote?.code || reservationKeyFromEvent(ev) || ev.id || ""),
      eventDate,
      startTime: String(ev.startTime || ""),
      endTime: String(ev.endTime || ""),
      eventName: String(ev.name || ""),
      salon: String(ev.salon || ""),
      company: String(company?.name || quote?.companyName || ""),
      manager: String(manager?.phone || quote?.managerPhone || ""),
      seller: String(user?.fullName || user?.name || ""),
      pax: Number(ev.pax || quote?.people || 0),
      total: Number(totals.total || 0),
      reservationKey,
      totalEvent: Number(metrics?.totalEvent || 0),
      incomePerDay: Number(metrics?.incomePerDay || 0),
      incomePerSalonDay: Number(metrics?.incomePerSalonDay || 0),
      hasQuote: !!latestQuote,
      latestQuoteVersion: latestQuoteVersion > 0 ? latestQuoteVersion : "",
      latestQuoteSentAt: String(latestQuote?.quotedAt || ""),
      hasMenuMontajeReport: !!latestMenuMontajeEntries.length,
      latestMenuMontajeVersion: latestMenuMontajeVersion > 0 ? latestMenuMontajeVersion : "",
      latestMenuMontajeAt: String(metrics?.latestMenuMontajeAt || latestMenuMontajeSnap?.quotedAt || ""),
      hasChecklist: checklistMeta.hasChecklist,
      checklistCompleted: checklistMeta.completed,
      checklistUpdatedAt: checklistMeta.updatedAt,
      updatedAt: getEventLastUpdatedLabel(ev),
    });
  }
  return rows.sort((a, b) => {
    const d = a.eventDate.localeCompare(b.eventDate);
    if (d !== 0) return d;
    const t = a.startTime.localeCompare(b.startTime);
    if (t !== 0) return t;
    return a.salon.localeCompare(b.salon);
  });
}

function getOccupancyUniqueReservationTotal(rows, fieldName = "totalEvent") {
  const totalsByReservation = new Map();
  for (const r of rows) {
    const key = String(r?.reservationKey || r?.eventId || "");
    if (!key) continue;
    const amount = Math.max(0, Number(r?.[fieldName] || 0));
    if (!totalsByReservation.has(key)) totalsByReservation.set(key, amount);
  }
  return Array.from(totalsByReservation.values()).reduce((acc, n) => acc + n, 0);
}

function renderOccupancySummary(rows) {
  if (!el.occupancyReportSummary) return;
  const confirmed = rows.filter((r) => r.status === STATUS.CONFIRMADO).length;
  const pre = rows.filter((r) => r.status === STATUS.PRERESERVA).length;
  const pax = rows.reduce((acc, r) => acc + Math.max(0, Number(r.pax || 0)), 0);
  const total = getOccupancyUniqueReservationTotal(rows, "totalEvent");
  const activeDays = new Set(rows.map((r) => String(r.eventDate || "").trim()).filter(Boolean)).size;
  const avgPax = rows.length ? Math.round(pax / rows.length) : 0;
  const confirmedPct = rows.length ? Math.round((confirmed / rows.length) * 100) : 0;
  const prePct = rows.length ? Math.round((pre / rows.length) * 100) : 0;
  const avgTicket = rows.length ? total / rows.length : 0;
  const cards = [
    { tone: "events", code: "EV", label: "Eventos semana", value: String(rows.length), meta: `${activeDays} dia(s) con actividad`, hint: rows.length ? "Panorama general de ocupacion" : "Sin actividad en el periodo" },
    { tone: "confirmed", code: "OK", label: "Confirmados", value: String(confirmed), meta: `${confirmedPct}% del total semanal`, hint: "Eventos listos para ejecutar" },
    { tone: "pre", code: "PR", label: "Pre reserva", value: String(pre), meta: `${prePct}% del total semanal`, hint: "Pendientes de consolidacion" },
    { tone: "pax", code: "PX", label: "PAX total", value: String(pax), meta: `Promedio ${avgPax} por evento`, hint: "Capacidad movilizada en la semana" },
    { tone: "revenue", code: "GT", label: "Total cotizado", value: moneyGT(total), meta: `Ticket promedio ${moneyGT(avgTicket)}`, hint: "Monto unico por reserva consolidada" },
  ];
  el.occupancyReportSummary.innerHTML = cards.map((card) => `
    <article class="occupancyCard occupancyCard--${escapeHtml(card.tone)}">
      <div class="occupancyCardTop">
        <span class="occupancyCardKicker">${escapeHtml(card.label)}</span>
        <span class="occupancyCardIcon" aria-hidden="true">${escapeHtml(card.code)}</span>
      </div>
      <strong>${escapeHtml(card.value)}</strong>
      <div class="occupancyCardMeta">${escapeHtml(card.meta)}</div>
      <div class="occupancyCardHint">${escapeHtml(card.hint)}</div>
    </article>
  `).join("");
}

function formatDayCardLabel(isoDate) {
  const d = new Date(`${isoDate}T00:00:00`);
  if (Number.isNaN(d.getTime())) return isoDate;
  return d.toLocaleDateString("es-GT", { weekday: "short", day: "2-digit", month: "2-digit" }).toUpperCase();
}

function renderOccupancyDayCards(rows) {
  if (!el.occupancyDaysStrip) return;
  const { monday } = getOccupancyWeekRange();
  const dates = Array.from({ length: 7 }, (_, i) => toISODate(addDays(monday, i)));
  if (!occupancySelectedDayIso || !dates.includes(occupancySelectedDayIso)) {
    occupancySelectedDayIso = dates[0];
  }
  el.occupancyDaysStrip.innerHTML = "";
  for (const d of dates) {
    const dayRows = rows.filter((r) => r.eventDate === d);
    const count = dayRows.length;
    const confirmedCount = dayRows.filter((r) => r.status === STATUS.CONFIRMADO).length;
    const preCount = dayRows.filter((r) => r.status === STATUS.PRERESERVA).length;
    const revenue = getOccupancyUniqueReservationTotal(dayRows, "totalEvent");
    const label = formatDayCardLabel(d);
    const parts = label.split(",");
    const dayName = String(parts[0] || label).trim();
    const dayDate = String(parts[1] || d).trim();
    const card = document.createElement("button");
    card.type = "button";
    card.className = `occupancyDayCard${d === occupancySelectedDayIso ? " active" : ""}${count ? " hasEvents" : ""}`;
    card.innerHTML = `
      <div class="occupancyDayCardTop">
        <small>${escapeHtml(dayName)}</small>
        <span class="occupancyDayDate">${escapeHtml(dayDate)}</span>
      </div>
      <div class="occupancyDayPrimary">
        <strong>${count}</strong>
        <span>evento${count === 1 ? "" : "s"}</span>
      </div>
      <div class="occupancyDayMeta" style="display:flex; align-items:center; gap:8px; margin-top:6px;"><span class="occupancyDayStatChip occupancyDayStatChip--confirmed" style="display:inline-flex; align-items:center; justify-content:space-between; gap:6px; min-width:84px; padding:5px 9px; border-radius:999px; border:1px solid rgba(34,197,94,.34); background:rgba(34,197,94,.12); box-shadow:inset 0 1px 0 rgba(255,255,255,.04);"><b style="color:#14532d; font-size:12px; font-weight:900; line-height:1;">${escapeHtml(String(confirmedCount))}</b><small style="display:inline-block; width:34px; text-align:center; color:#166534; font-size:10px; font-weight:800; letter-spacing:.04em; text-transform:uppercase; line-height:1;">Conf.</small></span><span class="occupancyDayStatChip occupancyDayStatChip--pre" style="display:inline-flex; align-items:center; justify-content:space-between; gap:6px; min-width:84px; padding:5px 9px; border-radius:999px; border:1px solid rgba(56,189,248,.34); background:rgba(56,189,248,.12); box-shadow:inset 0 1px 0 rgba(255,255,255,.04);"><b style="color:#0c4a6e; font-size:12px; font-weight:900; line-height:1;">${escapeHtml(String(preCount))}</b><small style="display:inline-block; width:34px; text-align:center; color:#075985; font-size:10px; font-weight:800; letter-spacing:.04em; text-transform:uppercase; line-height:1;">Pre</small></span></div>
      <div class="occupancyDayRevenue">${escapeHtml(count ? moneyGT(revenue) : "Sin monto")}</div>
    `;
    card.addEventListener("click", () => {
      occupancySelectedDayIso = d;
      renderOccupancyDayCards(rows);
      renderOccupancyDayDetail(rows);
    });
    el.occupancyDaysStrip.appendChild(card);
  }
}

function renderOccupancyDayDetail(rows) {
  if (!el.occupancyDayDetail) return;
  const target = String(occupancySelectedDayIso || "").trim();
  const dayRows = rows.filter((r) => r.eventDate === target);
  const title = target ? `Detalle ${target}` : "Detalle del dia";
  const confirmedCount = dayRows.filter((r) => r.status === STATUS.CONFIRMADO).length;
  const preCount = dayRows.filter((r) => r.status === STATUS.PRERESERVA).length;
  const pax = dayRows.reduce((acc, r) => acc + Math.max(0, Number(r.pax || 0)), 0);
  const revenue = getOccupancyUniqueReservationTotal(dayRows, "totalEvent");
  const metrics = `
    <div class="occupancyDetailStat"><small>Eventos</small><strong>${escapeHtml(String(dayRows.length))}</strong></div>
    <div class="occupancyDetailStat"><small>Confirmados</small><strong>${escapeHtml(String(confirmedCount))}</strong></div>
    <div class="occupancyDetailStat"><small>Pre reserva</small><strong>${escapeHtml(String(preCount))}</strong></div>
    <div class="occupancyDetailStat"><small>PAX</small><strong>${escapeHtml(String(pax))}</strong></div>
    <div class="occupancyDetailStat"><small>Total</small><strong>${escapeHtml(moneyGT(revenue))}</strong></div>
  `;
  if (!dayRows.length) {
    el.occupancyDayDetail.innerHTML = `
      <div class="occupancyDetailHeader">
        <div>
          <div class="occupancyDayDetailTitle">${escapeHtml(title)}</div>
          <div class="occupancyDayDetailText">No se registran eventos Confirmados o Pre reserva en este dia.</div>
        </div>
        <div class="occupancyDetailStats">${metrics}</div>
      </div>
      <div class="occupancyDayDetailEmpty">Elige otro dia de la semana para revisar eventos, montos y acciones operativas.</div>
    `;
    return;
  }
  const cards = dayRows.map((r) => {
    const rowTone = r.status === STATUS.CONFIRMADO ? "confirmed" : "pre";
    return `
      <article class="occupancyEventCard occupancyEventCard--${escapeHtml(rowTone)}">
        <div class="occupancyEventHead">
          <div class="occupancyEventHeading">
            <span class="salesStatusBadge" style="background:${escapeHtml(hexToRgba(r.statusColor, 0.18))};border-color:${escapeHtml(hexToRgba(r.statusColor, 0.48))};color:#0f172a">${escapeHtml(r.status)}</span>
            <strong title="${escapeHtml(r.eventName || "-")}">${escapeHtml(r.eventName || "-")}</strong>
            <small>${escapeHtml(r.refId || "Sin referencia")} | ${escapeHtml(r.startTime || "-")} - ${escapeHtml(r.endTime || "-")}</small>
          </div>
          <div class="occupancyEventMoney">
            <span>Total evento</span>
            <strong>${escapeHtml(moneyGT(r.totalEvent || 0))}</strong>
          </div>
        </div>
        <div class="occupancyEventBody">
          <div class="occupancyInfoGrid">
            <div class="occupancyInfoItem occupancyInfoWide"><small>Institucion</small><strong title="${escapeHtml(r.company || "-")}">${escapeHtml(r.company || "-")}</strong></div>
            <div class="occupancyInfoItem"><small>Salon</small><strong>${escapeHtml(r.salon || "-")}</strong></div>
            <div class="occupancyInfoItem"><small>Encargado</small><strong title="${escapeHtml(r.manager || "-")}">${escapeHtml(r.manager || "-")}</strong></div>
            <div class="occupancyInfoItem"><small>Vendedor</small><strong title="${escapeHtml(r.seller || "-")}">${escapeHtml(r.seller || "-")}</strong></div>
            <div class="occupancyInfoItem"><small>PAX</small><strong>${escapeHtml(String(r.pax || 0))}</strong></div>
          </div>
          <div class="occupancyMetricsGrid">
            <div class="occupancyMetricItem"><small>Ingreso dia</small><strong>${escapeHtml(moneyGT(r.incomePerDay || 0))}</strong></div>
            <div class="occupancyMetricItem"><small>Ingreso salon-dia</small><strong>${escapeHtml(moneyGT(r.incomePerSalonDay || 0))}</strong></div>
            <div class="occupancyMetricItem"><small>Ultima modificacion</small><strong>${escapeHtml(r.updatedAt || "-")}</strong></div>
          </div>
          <div class="occupancyActionGrid">
            <div class="occupancyActionItem"><small>Ultima cotizacion</small><div>${buildOccupancyQuoteActionHtml(r)}</div></div>
            <div class="occupancyActionItem"><small>Ultimo informe</small><div>${buildOccupancyMenuMontajeActionHtml(r)}</div></div>
            <div class="occupancyActionItem"><small>Check List</small><div>${buildOccupancyChecklistActionHtml(r)}</div></div>
          </div>
        </div>
      </article>
    `;
  }).join("");
  el.occupancyDayDetail.innerHTML = `
    <div class="occupancyDetailHeader">
      <div>
        <div class="occupancyDayDetailTitle">${escapeHtml(title)}</div>
        <div class="occupancyDayDetailText">Lectura rapida del dia seleccionado con eventos, responsables y acciones relacionadas.</div>
      </div>
      <div class="occupancyDetailStats">${metrics}</div>
    </div>
    <div class="occupancyEventCards">${cards}</div>
  `;
}

function renderOccupancyReportTable() {
  if (!el.occupancyReportBody) return;
  const rows = buildOccupancyReportRows();
  el.occupancyReportBody.innerHTML = "";
  renderOccupancySummary(rows);
  renderOccupancyDayCards(rows);
  renderOccupancyDayDetail(rows);
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="17" class="occupancyTableEmpty">Sin eventos Confirmados/Pre reserva para esta semana.</td>`;
    el.occupancyReportBody.appendChild(tr);
    return;
  }
  for (const r of rows) {
    const tr = document.createElement("tr");
    const rowTone = r.status === STATUS.CONFIRMADO ? "confirmed" : "pre";
    tr.className = `occupancyTableRow occupancyTableRow--${rowTone}`;
    tr.innerHTML = `
      <td class="occupancyTableStatus"><span class="salesStatusBadge" style="background:${escapeHtml(hexToRgba(r.statusColor, 0.18))};border-color:${escapeHtml(hexToRgba(r.statusColor, 0.48))};color:#0f172a">${escapeHtml(r.status || "-")}</span></td>
      <td class="occupancyTablePax">${escapeHtml(String(r.pax || 0))}</td>
      <td class="occupancyTableDate">${escapeHtml(r.eventDate || "-")}</td>
      <td class="occupancyTableTime">${escapeHtml(r.startTime || "-")}</td>
      <td class="occupancyTableTime">${escapeHtml(r.endTime || "-")}</td>
      <td class="occupancyTableEvent"><strong title="${escapeHtml(r.eventName || "-")}">${escapeHtml(r.eventName || "-")}</strong></td>
      <td class="occupancyTableSalon">${escapeHtml(r.salon || "-")}</td>
      <td class="occupancyTableInstitution">${escapeHtml(r.company || "-")}</td>
      <td class="occupancyTableManager">${escapeHtml(r.manager || "-")}</td>
      <td class="occupancyTableSeller">${escapeHtml(r.seller || "-")}</td>
      <td>${buildOccupancyQuoteActionHtml(r)}</td>
      <td>${buildOccupancyMenuMontajeActionHtml(r)}</td>
      <td>${buildOccupancyChecklistActionHtml(r)}</td>
      <td class="occupancyTableMoney">${escapeHtml(moneyGT(r.totalEvent || 0))}</td>
      <td class="occupancyTableMoney">${escapeHtml(moneyGT(r.incomePerDay || 0))}</td>
      <td class="occupancyTableMoney">${escapeHtml(moneyGT(r.incomePerSalonDay || 0))}</td>
      <td class="occupancyTableUpdated">${escapeHtml(r.updatedAt || "-")}</td>
    `;
    el.occupancyReportBody.appendChild(tr);
  }
}

function exportOccupancyReportToExcel() {
  const rows = buildOccupancyReportRows();
  if (!rows.length) return toast("No hay datos para exportar.");
  const { monday, sunday } = getOccupancyWeekRange();
  const weekLabel = `${toISODate(monday)} a ${toISODate(sunday)}`;
  const generatedAt = new Date().toLocaleString("es-GT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
  const emittedBy = String(authSession.fullName || authSession.username || "Sistema").trim();
  const htmlRows = rows.map((r) => `
    <tr>
      <td style="background:${escapeHtml(hexToRgba(r.statusColor, 0.25))}; border:1px solid #c7d5ea; font-weight:700;">${escapeHtml(r.status)}</td>
      <td>${escapeHtml(String(r.pax || 0))}</td>
      <td>${escapeHtml(r.eventDate)}</td>
      <td>${escapeHtml(r.startTime)}</td>
      <td>${escapeHtml(r.endTime)}</td>
      <td>${escapeHtml(r.eventName)}</td>
      <td>${escapeHtml(r.salon)}</td>
      <td>${escapeHtml(r.company)}</td>
      <td>${escapeHtml(r.manager)}</td>
      <td>${escapeHtml(r.seller)}</td>
      <td>${escapeHtml((() => {
    const v = Number(r.latestQuoteVersion || 0);
    const versionLabel = v > 0 ? `V${v}` : "-";
    const sent = formatQuoteSentAtLabel(r.latestQuoteSentAt);
    return sent ? `${versionLabel} - ${sent}` : versionLabel;
  })())}</td>
      <td>${escapeHtml((() => {
    const v = Number(r.latestMenuMontajeVersion || 0);
    const versionLabel = v > 0 ? `V${v}` : "-";
    const sent = formatQuoteSentAtLabel(r.latestMenuMontajeAt);
    return sent ? `${versionLabel} - ${sent}` : versionLabel;
  })())}</td>
      <td>${escapeHtml(r.checklistCompleted ? "Completo" : (r.hasChecklist ? "En proceso" : "Sin iniciar"))}</td>
      <td>${escapeHtml(moneyGT(r.totalEvent || 0))}</td>
      <td>${escapeHtml(moneyGT(r.incomePerDay || 0))}</td>
      <td>${escapeHtml(moneyGT(r.incomePerSalonDay || 0))}</td>
      <td>${escapeHtml(r.updatedAt || "-")}</td>
    </tr>
  `).join("");
  const html = `<!doctype html>
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
<head>
  <meta charset="utf-8" />
  <meta name="ProgId" content="Excel.Sheet" />
  <style>
    body{ font-family: Calibri, Arial, sans-serif; background:#eef3fb; margin:0; padding:16px; color:#0f172a; }
    .card{ background:#ffffff; border:1px solid #c5d4ea; border-radius:10px; overflow:hidden; }
    .titleCell{ border:1px solid #c7d5ea; background:#d8e3f3; color:#000; font-weight:800; font-size:20px; padding:12px 14px; text-transform:uppercase; }
    .meta{ padding:10px 14px; border-top:1px solid #bfd3ee; border-bottom:1px solid #bfd3ee; background:#eaf3ff; font-size:12px; }
    .meta div{ margin:2px 0; }
    table{ width:100%; border-collapse:collapse; }
    th,td{ border:1px solid #c7d5ea; padding:6px 7px; font-size:10.5px; white-space:nowrap; }
    thead th{ background:#0f3c67; color:#fff; font-weight:700; text-transform:uppercase; }
  </style>
</head>
<body>
  <div class="card">
    <table><tr><td class="titleCell">CRM JARDINES - REPORTE DE OCUPACION</td></tr></table>
    <div class="meta">
      <div><b>Fecha:</b> ${escapeHtml(generatedAt)}</div>
      <div><b>Quien emitio el reporte:</b> ${escapeHtml(emittedBy)}</div>
      <div><b>Semana:</b> ${escapeHtml(weekLabel)} (Lunes a Domingo)</div>
      <div><b>Estados incluidos:</b> Confirmado y Pre reserva</div>
      <div><b>Total registros:</b> ${rows.length}</div>
    </div>
    <table>
      <thead>
        <tr>
          <th>Estado</th><th>PAX</th><th>Fecha evento</th><th>Hora inicio</th><th>Hora final</th><th>Evento</th><th>Salon</th><th>Institucion</th><th>Encargado evento</th><th>Vendedor</th><th>Ultima cotizacion enviada</th><th>Ultimo informe menu/montaje</th><th>Check List</th><th>Total evento</th><th>Ingreso dia</th><th>Ingreso salon-dia</th><th>Ultima modificacion</th>
        </tr>
      </thead>
      <tbody>${htmlRows}</tbody>
    </table>
  </div>
</body>
</html>`;
  const blob = new Blob([`\uFEFF${html}`], { type: "application/vnd.ms-excel;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `reporte_ocupacion_${toISODate(monday)}_${toISODate(sunday)}.xls`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function setOccupancyCurrentWeek() {
  if (!el.occupancyReportWeek) return;
  el.occupancyReportWeek.value = weekInputFromDate(new Date());
}

function openOccupancyReportModal() {
  if (!el.occupancyReportBackdrop) return;
  if (!String(el.occupancyReportWeek?.value || "").trim()) {
    setOccupancyCurrentWeek();
  }
  updateOccupancyReportWeekUi();
  el.occupancyReportBackdrop.hidden = false;
}

function closeOccupancyReportModal() {
  if (!el.occupancyReportBackdrop) return;
  el.occupancyReportBackdrop.hidden = true;
  restoreModuleScreenAfterModal();
}

function dashboardMonthKeyFromIso(iso) {
  const raw = String(iso || "").trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(raw)) return "";
  return raw.slice(0, 7);
}

function dashboardResolveMonthKey() {
  const month = String(el.dashboardReportMonth?.value || "").trim();
  if (/^\d{4}-\d{2}$/.test(month)) return month;
  const from = String(el.dashboardReportFrom?.value || "").trim();
  const fallback = dashboardMonthKeyFromIso(from);
  return fallback || toISODate(new Date()).slice(0, 7);
}

function dashboardMonthBounds(monthKey) {
  const m = String(monthKey || "").trim();
  if (!/^\d{4}-\d{2}$/.test(m)) {
    const now = new Date();
    const start = startOfMonth(now);
    const end = addDays(startOfMonth(addMonths(start, 1)), -1);
    return { fromIso: toISODate(start), toIso: toISODate(end) };
  }
  const [y, mo] = m.split("-").map(Number);
  const start = new Date(y, mo - 1, 1);
  start.setHours(0, 0, 0, 0);
  const end = addDays(startOfMonth(addMonths(start, 1)), -1);
  return { fromIso: toISODate(start), toIso: toISODate(end) };
}

function setDashboardFiltersByMonth(monthKey) {
  const { fromIso, toIso } = dashboardMonthBounds(monthKey);
  if (el.dashboardReportMonth) el.dashboardReportMonth.value = String(monthKey || "").trim();
  if (el.dashboardReportFrom) el.dashboardReportFrom.value = fromIso;
  if (el.dashboardReportTo) el.dashboardReportTo.value = toIso;
}

function dashboardResolvePeriod() {
  return String(el.dashboardReportPeriod?.value || "month").trim() === "week" ? "week" : "month";
}

function dashboardWeekBounds(weekIso) {
  const monday = mondayFromWeekInput(weekIso || "");
  const sunday = addDays(monday, 6);
  return { fromIso: toISODate(monday), toIso: toISODate(sunday) };
}

function setDashboardFiltersByWeek(weekIso) {
  const { fromIso, toIso } = dashboardWeekBounds(weekIso);
  if (el.dashboardReportWeek) el.dashboardReportWeek.value = fromIso;
  if (el.dashboardReportFrom) el.dashboardReportFrom.value = fromIso;
  if (el.dashboardReportTo) el.dashboardReportTo.value = toIso;
  if (el.dashboardReportMonth && /^\d{4}-\d{2}$/.test(fromIso.slice(0, 7))) {
    el.dashboardReportMonth.value = fromIso.slice(0, 7);
  }
}

function syncDashboardPeriodControls() {
  const period = dashboardResolvePeriod();
  const isWeek = period === "week";
  if (el.dashboardReportWeekField) el.dashboardReportWeekField.hidden = !isWeek;
  if (el.dashboardReportMonth) el.dashboardReportMonth.disabled = isWeek;
  if (el.dashboardReportWeek) el.dashboardReportWeek.disabled = !isWeek;
  if (el.btnDashboardReportCurrentMonth) {
    el.btnDashboardReportCurrentMonth.textContent = isWeek ? "Semana actual" : "Mes actual";
  }
}

function dashboardFormatRangeLabel(period, fromIso, toIso, monthKey) {
  if (period === "week") return `Semana ${fromIso} a ${toIso}`;
  return formatMonthKeyLabel(monthKey || dashboardMonthKeyFromIso(fromIso || ""));
}

function dashboardResolveRole() {
  return normalizeUserRole(el.dashboardReportRole?.value || USER_ROLE.SELLER);
}

function dashboardGetActiveUsersByRole(role = dashboardResolveRole()) {
  const targetRole = normalizeUserRole(role);
  return (state.users || [])
    .map(normalizeUserRecord)
    .filter((u) => u.active !== false)
    .filter((u) => REPORTABLE_USER_ROLES.includes(targetRole) ? normalizeUserRole(u.role) === targetRole : true)
    .sort((a, b) => String(a.fullName || a.name || "").localeCompare(String(b.fullName || b.name || ""), "es", { sensitivity: "base" }));
}

function renderDashboardSellerFilterOptions() {
  if (!el.dashboardReportSeller) return;
  const previous = String(el.dashboardReportSeller.value || "").trim();
  const role = dashboardResolveRole();
  const roleLabel = userRoleLabel(role);
  const sellers = dashboardGetActiveUsersByRole(role);
  el.dashboardReportSeller.innerHTML = "";
  const empty = document.createElement("option");
  empty.value = "";
  empty.textContent = `Selecciona ${roleLabel.toLowerCase()}`;
  el.dashboardReportSeller.appendChild(empty);
  for (const s of sellers) {
    const opt = document.createElement("option");
    opt.value = s.id;
    opt.textContent = s.fullName || s.name;
    el.dashboardReportSeller.appendChild(opt);
  }
  if (previous && sellers.some((s) => String(s.id) === previous)) {
    el.dashboardReportSeller.value = previous;
  }
}

function dashboardResolveScopeUserId() {
  const scope = String(el.dashboardReportScope?.value || "mine").trim();
  const sessionUserId = String(authSession.userId || "").trim();
  const authUser = normalizeUserRecord(getAuthUserRecord() || {});
  const role = dashboardResolveRole();
  const selectedSellerId = String(el.dashboardReportSeller?.value || "").trim();
  if (scope === "mine" && sessionUserId && normalizeUserRole(authUser.role) === role) return sessionUserId;
  if (scope === "seller" && selectedSellerId) return selectedSellerId;
  return "";
}

function dashboardNormalizeEventType(value) {
  const norm = normalizeBucketKey(value);
  if (matchesAliases(norm, ["corporativo", "corporate", "empresa", "empresarial"])) return "Corporativo";
  if (matchesAliases(norm, ["social", "boda", "xv", "xv anos", "cumple", "aniversario", "fiesta"])) return "Social";
  return "Individual";
}

function dashboardNormalizeStatus(status) {
  const raw = String(status || "").trim();
  const norm = normalizeBucketKey(raw);
  if (matchesAliases(norm, ["reserva sin cotizacion", "reservacion sin cotizacion", "sin cotizacion", "reserva"])) {
    return STATUS.RESERVA_SIN_COTIZACION;
  }
  if (matchesAliases(norm, ["1er cotizacion", "1ra cotizacion", "primera cotizacion", "1er reserva", "primera reserva"])) {
    return STATUS.PRIMERA;
  }
  return raw;
}

function buildDashboardReportRows(fromIso, toIso) {
  const groups = new Map();
  for (const ev of state.events || []) {
    const key = reservationKeyFromEvent(ev);
    if (!key) continue;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push(ev);
  }
  const rows = [];
  for (const [reservationKey, seriesRaw] of groups.entries()) {
    const series = (seriesRaw || []).slice().sort((a, b) => {
      const d = String(a.date || "").localeCompare(String(b.date || ""));
      if (d !== 0) return d;
      return String(a.startTime || "").localeCompare(String(b.startTime || ""));
    });
    if (!series.length) continue;
    const intersects = series.some((ev) => {
      const date = String(ev.date || "");
      if (!date) return false;
      if (fromIso && date < fromIso) return false;
      if (toIso && date > toIso) return false;
      return true;
    });
    if (!intersects) continue;
    const head = series[0];
    const quoteSnapshot = getLatestQuoteSnapshotFromSeries(series) || getLatestQuoteSnapshotForEvent(head) || head.quote || {};
    const totals = getQuoteTotals(quoteSnapshot || {});
    const eventTypeRaw = String(quoteSnapshot?.eventType || head.name || "").trim();
    rows.push({
      reservationKey: String(reservationKey),
      userId: String(head.userId || "").trim(),
      status: dashboardNormalizeStatus(head.status),
      eventDate: String(head.date || "").trim(),
      monthKey: dashboardMonthKeyFromIso(head.date),
      salon: String(head.salon || "").trim(),
      total: Math.max(0, Number(totals.total || 0)),
      type: dashboardNormalizeEventType(eventTypeRaw),
    });
  }
  return rows;
}

function isGoalStatus(status) {
  const s = String(status || "").trim();
  return s === STATUS.CONFIRMADO || s === STATUS.PRERESERVA;
}

function dashboardGoalCardState(progressPct) {
  if (progressPct >= 100) return "over";
  if (progressPct >= 80) return "near";
  return "far";
}

function getDashboardHeroStatusMeta() {
  return [
    { key: STATUS.CONFIRMADO, label: "Confirmado", color: "#10c972" },
    { key: STATUS.RESERVA_SIN_COTIZACION, label: "Reserva sin cotizacion", color: "#0ea5e9" },
    { key: STATUS.PRIMERA, label: "1era. Cotizacion", color: "#aa97df" },
    { key: STATUS.SEGUIMIENTO, label: "Negociacion", color: "#ff6b3a" },
    { key: STATUS.PERDIDO, label: "Perdido", color: "#7c5cff" },
    { key: STATUS.LISTA, label: "Lista de Espera", color: "#f5c400" },
    { key: STATUS.PRERESERVA, label: "Pre Reserva", color: "#d07db8" },
    { key: STATUS.CANCELADO, label: "Cancelado", color: "#e42a48" },
  ];
}

function buildDashboardHeroStatusSummary(rows) {
  const meta = getDashboardHeroStatusMeta();
  const counters = new Map(meta.map((item) => [item.key, 0]));
  for (const row of Array.isArray(rows) ? rows : []) {
    const key = String(row?.status || "").trim();
    if (!counters.has(key)) continue;
    counters.set(key, Number(counters.get(key) || 0) + 1);
  }
  const total = Array.from(counters.values()).reduce((acc, n) => acc + Number(n || 0), 0);
  const segments = meta.map((item) => {
    const count = Number(counters.get(item.key) || 0);
    const pct = total > 0 ? (count / total) * 100 : 0;
    return {
      ...item,
      count,
      pct,
    };
  });
  const confirmed = Number(counters.get(STATUS.CONFIRMADO) || 0);
  const confirmedPct = total > 0 ? (confirmed / total) * 100 : 0;
  return { total, confirmed, confirmedPct, segments };
}

function ensureDashboardHoverTip() {
  if (dashboardHoverTipEl && document.body.contains(dashboardHoverTipEl)) return dashboardHoverTipEl;
  dashboardHoverTipEl = document.createElement("div");
  dashboardHoverTipEl.className = "dashboardHoverTip";
  dashboardHoverTipEl.hidden = true;
  document.body.appendChild(dashboardHoverTipEl);
  return dashboardHoverTipEl;
}

function showDashboardHoverTip(text, clientX, clientY) {
  const message = String(text || "").trim();
  if (!message) {
    hideDashboardHoverTip();
    return;
  }
  const tip = ensureDashboardHoverTip();
  tip.textContent = message;
  tip.hidden = false;
  const pad = 14;
  const width = tip.offsetWidth || 180;
  const height = tip.offsetHeight || 44;
  let left = Number(clientX || 0) + 16;
  let top = Number(clientY || 0) + 16;
  if (left + width + pad > window.innerWidth) left = Math.max(pad, Number(clientX || 0) - width - 16);
  if (top + height + pad > window.innerHeight) top = Math.max(pad, Number(clientY || 0) - height - 16);
  tip.style.left = `${Math.round(left)}px`;
  tip.style.top = `${Math.round(top)}px`;
}

function hideDashboardHoverTip() {
  if (!dashboardHoverTipEl) return;
  dashboardHoverTipEl.hidden = true;
}

function handleDashboardHoverTooltip(event) {
  const target = event.target instanceof Element ? event.target.closest("[data-dashboard-tooltip]") : null;
  if (!target) {
    hideDashboardHoverTip();
    return;
  }
  showDashboardHoverTip(target.getAttribute("data-dashboard-tooltip") || "", event.clientX, event.clientY);
}

function renderDashboardGoals(globalMeta, personalMeta) {
  if (!el.dashboardGoalsGrid) return;
  const gp = Math.max(0, Number(globalMeta.progress || 0));
  const pp = Math.max(0, Number(personalMeta.progress || 0));
  const gState = dashboardGoalCardState(gp);
  const pState = dashboardGoalCardState(pp);
  const globalRemaining = Math.max(0, Number(globalMeta.remaining || 0));
  const personalRemaining = Math.max(0, Number(personalMeta.remaining || 0));
  const summary = globalMeta.statusSummary || { total: 0, segments: [] };
  const heroPct = Math.max(0, Number(summary.confirmedPct || 0));
  const visibleSegments = (summary.segments || []).filter((item) => Number(item.count || 0) > 0);
  const segmentHtml = visibleSegments.length
    ? visibleSegments.map((item) => `
        <span class="dashboardHeroSegment" style="width:${Math.max(2, item.pct)}%;background:${escapeHtml(item.color)}" data-dashboard-tooltip="${escapeHtml(item.label)}: ${escapeHtml(String(item.count))} evento(s) | ${escapeHtml(item.pct.toFixed(1))}% del periodo"></span>
      `).join("")
    : `<span class="dashboardHeroSegment dashboardHeroSegmentEmpty"></span>`;
  const legendHtml = (summary.segments || [])
    .filter((item) => Number(item.count || 0) > 0)
    .map((item) => `
      <span class="dashboardHeroLegendItem" data-dashboard-tooltip="${escapeHtml(item.label)}: ${escapeHtml(String(item.count))} evento(s) | ${escapeHtml(item.pct.toFixed(1))}% del periodo">
        <i style="background:${escapeHtml(item.color)}"></i>
        <span>${escapeHtml(item.label)} ${escapeHtml(item.pct.toFixed(1))}%</span>
      </span>
    `).join("");
  el.dashboardGoalsGrid.innerHTML = `
    <article class="dashboardHeroCard dashboardGoalCard--${escapeHtml(gState)}">
      <div class="dashboardHeroHead">
        <div>
          <small>EFICIENCIA EN ${escapeHtml(userRolePluralLabel(globalMeta.roleKey || USER_ROLE.SELLER).toUpperCase())} "CRM"</small>
          <strong>${escapeHtml(globalMeta.periodLabel || formatMonthKeyLabel(globalMeta.monthKey || ""))}</strong>
        </div>
        <div class="dashboardHeroPct" data-dashboard-tooltip="${escapeHtml(String(summary.confirmed || 0))} confirmados de ${escapeHtml(String(summary.total || 0))} eventos del periodo">
          <b>${escapeHtml(heroPct.toFixed(1))}%</b>
          <span>Confirmado</span>
        </div>
      </div>
      <div class="dashboardHeroBarWrap">
        <div class="dashboardHeroBar">${segmentHtml}</div>
      </div>
      <div class="dashboardHeroLegend">
        ${legendHtml || `<span class="dashboardEmpty">Sin estados para el periodo.</span>`}
      </div>
    </article>
    <article class="dashboardGoalCard dashboardGoalCard--${escapeHtml(gState)}">
      <small>Meta ${escapeHtml(globalMeta.roleLabel || "global")}</small>
      <strong>${escapeHtml(moneyGT(globalMeta.goal || 0))}</strong>
      <div class="dashboardGoalMeta">Avance ${escapeHtml(moneyGT(globalMeta.achieved || 0))} | ${escapeHtml(gp.toFixed(1))}%</div>
      <div class="dashboardGoalProgress"><span style="width:${Math.min(gp, 100)}%"></span></div>
    </article>
    <article class="dashboardGoalCard dashboardGoalCard--${escapeHtml(gState)}">
      <small>Pendiente del rol</small>
      <strong>${escapeHtml(moneyGT(globalRemaining))}</strong>
      <div class="dashboardGoalMeta">${globalRemaining <= 0 ? `Meta ${escapeHtml(globalMeta.roleLabel || "global")} superada` : "Ingreso pendiente para cumplir meta"}</div>
      <div class="dashboardGoalProgress"><span style="width:${Math.min(gp, 100)}%"></span></div>
    </article>
    <article class="dashboardGoalCard dashboardGoalCard--${escapeHtml(pState)}">
      <small>Meta personal (${escapeHtml(personalMeta.sellerLabel)})</small>
      <strong>${escapeHtml(moneyGT(personalMeta.goal || 0))}</strong>
      <div class="dashboardGoalMeta">Avance ${escapeHtml(moneyGT(personalMeta.achieved || 0))} | ${escapeHtml(pp.toFixed(1))}%</div>
      <div class="dashboardGoalProgress"><span style="width:${Math.min(pp, 100)}%"></span></div>
    </article>
    <article class="dashboardGoalCard dashboardGoalCard--${escapeHtml(pState)}">
      <small>Falta para meta personal</small>
      <strong>${escapeHtml(moneyGT(personalRemaining))}</strong>
      <div class="dashboardGoalMeta">${personalRemaining <= 0 ? "Meta personal superada" : "Ingreso pendiente para cumplir meta"}</div>
      <div class="dashboardGoalProgress"><span style="width:${Math.min(pp, 100)}%"></span></div>
    </article>
  `;
}

function renderDashboardBars(container, series, highlightMax = false) {
  if (!container) return;
  if (!Array.isArray(series) || !series.length) {
    container.innerHTML = `<div class="dashboardEmpty">Sin datos para esta grafica.</div>`;
    return;
  }
  const max = Math.max(1, ...series.map((x) => Number(x.value || 0)));
  const best = highlightMax ? Math.max(...series.map((x) => Number(x.value || 0))) : -1;
  container.innerHTML = series.map((row) => {
    const value = Math.max(0, Number(row.value || 0));
    const pct = Math.min(100, (value / max) * 100);
    const bestCls = highlightMax && value === best && best > 0 ? " best" : "";
    return `
      <div class="dashboardBarRow">
        <div class="dashboardBarLabel">${escapeHtml(row.label || "-")}</div>
        <div class="dashboardBarTrack"><div class="dashboardBarFill${bestCls}" style="width:${pct}%"></div></div>
        <div class="dashboardBarValue">${escapeHtml(moneyGT(value))}</div>
      </div>
    `;
  }).join("");
}

function renderDashboardSalonUsageChart(rows) {
  if (!el.dashboardCompareChart) return;
  const counts = new Map();
  for (const row of Array.isArray(rows) ? rows : []) {
    const label = String(row?.salon || "").trim() || "Sin salon";
    counts.set(label, Number(counts.get(label) || 0) + 1);
  }
  const palette = ["#5b95f0", "#facc15", "#9b5de5", "#e92f55", "#10c972", "#67b7e1", "#ffb23d", "#11945a", "#8b5cf6", "#ec4899", "#ef4444", "#84cc16"];
  const ordered = Array.from(counts.entries())
    .map(([label, count], idx) => ({ label, count: Number(count || 0), color: palette[idx % palette.length] }))
    .sort((a, b) => Number(b.count || 0) - Number(a.count || 0) || String(a.label || "").localeCompare(String(b.label || ""), "es", { sensitivity: "base" }));
  const total = ordered.reduce((acc, item) => acc + Number(item.count || 0), 0);
  if (!ordered.length || total <= 0) {
    el.dashboardCompareChart.innerHTML = `<div class="dashboardEmpty">Sin salones con actividad para el periodo.</div>`;
    return;
  }
  const slices = [];
  let cursor = 0;
  for (const item of ordered) {
    const pct = (Number(item.count || 0) / total) * 100;
    const next = cursor + pct;
    slices.push(`${item.color} ${cursor.toFixed(2)}% ${next.toFixed(2)}%`);
    item.pct = pct;
    cursor = next;
  }
  const visibleLegend = ordered.slice(0, 6);
  const hiddenCount = Math.max(0, ordered.length - visibleLegend.length);
  const legend = visibleLegend.map((item) => `
    <div class="dashboardPieLegendItem" data-dashboard-tooltip="${escapeHtml(item.label)}: ${escapeHtml(String(item.count))} evento(s) | ${escapeHtml(item.pct.toFixed(1))}% del periodo">
      <i style="background:${escapeHtml(item.color)}"></i>
      <span>${escapeHtml(item.label)}: ${escapeHtml(item.pct.toFixed(1))}%</span>
    </div>
  `).join("");
  el.dashboardCompareChart.innerHTML = `
    <div class="dashboardPieLayout">
      <div class="dashboardPieChart" style="background:conic-gradient(${slices.join(", ")})"></div>
      <div class="dashboardPieLegend">
        ${legend}
        ${hiddenCount ? `<div class="dashboardPieLegendMore">+${hiddenCount} area(s) mas</div>` : ""}
      </div>
    </div>
  `;
}

function renderDashboardEventTypeChart(rows, selectedMonthKey) {
  if (!el.dashboardBestMonthChart) return;
  const year = Number(String(selectedMonthKey || "").slice(0, 4)) || new Date().getFullYear();
  const months = Array.from({ length: 12 }, (_, idx) => {
    const key = `${year}-${pad2(idx + 1)}`;
    const monthRows = rows.filter((r) => r.monthKey === key);
    const corp = rows
      .filter((r) => {
        if (r.monthKey !== key) return false;
        const type = String(r.type || "").trim().toLowerCase();
        return type === "corporativo" || type.includes("corpor");
      })
      .reduce((acc, r) => acc + Math.max(0, Number(r.total || 0)), 0);
    const social = rows
      .filter((r) => {
        if (r.monthKey !== key) return false;
        const type = String(r.type || "").trim().toLowerCase();
        return type === "social" || type.includes("social");
      })
      .reduce((acc, r) => acc + Math.max(0, Number(r.total || 0)), 0);
    return { label: String(idx + 1), corporativo: corp, social, totalEvents: monthRows.length };
  });
  const max = Math.max(1, ...months.flatMap((m) => [Number(m.corporativo || 0), Number(m.social || 0)]));
  const bars = months.map((row) => {
    const corpH = Math.max(0, (Number(row.corporativo || 0) / max) * 100);
    const socialH = Math.max(0, (Number(row.social || 0) / max) * 100);
    const hasCorp = Number(row.corporativo || 0) > 0;
    const hasSocial = Number(row.social || 0) > 0;
    const corpLabel = hasCorp ? `<span class="dashboardMonthBarValue">${escapeHtml(moneyGT(row.corporativo || 0))}</span>` : `<span class="dashboardMonthBarValue is-empty"></span>`;
    const socialLabel = hasSocial ? `<span class="dashboardMonthBarValue">${escapeHtml(moneyGT(row.social || 0))}</span>` : `<span class="dashboardMonthBarValue is-empty"></span>`;
    const corpHeight = hasCorp ? Math.max(14, corpH) : 6;
    const socialHeight = hasSocial ? Math.max(14, socialH) : 6;
    return `
      <div class="dashboardMonthGroup">
        <div class="dashboardMonthBars">
          <span class="dashboardMonthBarCol" data-dashboard-tooltip="Corporativo ${escapeHtml(moneyGT(row.corporativo || 0))}">
            ${corpLabel}
            <span class="dashboardMonthBar corp${hasCorp ? "" : " is-empty"}" style="height:${corpHeight}%"></span>
          </span>
          <span class="dashboardMonthBarCol" data-dashboard-tooltip="Social ${escapeHtml(moneyGT(row.social || 0))}">
            ${socialLabel}
            <span class="dashboardMonthBar social${hasSocial ? "" : " is-empty"}" style="height:${socialHeight}%"></span>
          </span>
        </div>
        <div class="dashboardMonthLabel" data-dashboard-tooltip="${escapeHtml(String(row.totalEvents || 0))} evento(s) en el mes ${escapeHtml(row.label)}">${escapeHtml(row.label)}</div>
      </div>
    `;
  }).join("");
  el.dashboardBestMonthChart.innerHTML = `
    <div class="dashboardMonthChart">
      <div class="dashboardMonthBarsGrid">${bars}</div>
      <div class="dashboardMonthLegend">
        <span class="dashboardPieLegendItem"><i style="background:#d4c23c"></i><span>Corporativo</span></span>
        <span class="dashboardPieLegendItem"><i style="background:#4b4b52"></i><span>Social</span></span>
      </div>
    </div>
  `;
}

function renderDashboardCharts(rows, selectedMonthKey) {
  const scopeUserId = dashboardResolveScopeUserId();
  const scopedRows = scopeUserId ? rows.filter((r) => String(r.userId) === scopeUserId) : rows.slice();
  if (el.dashboardCompareTitle) el.dashboardCompareTitle.textContent = "Areas mas utilizadas";
  if (el.dashboardCompareSubtitle) el.dashboardCompareSubtitle.textContent = "Distribucion de salones en el periodo";
  if (el.dashboardBestTitle) el.dashboardBestTitle.textContent = "Ventas por tipo de evento";
  if (el.dashboardBestSubtitle) el.dashboardBestSubtitle.textContent = "Corporativo vs Social por mes";
  renderDashboardSalonUsageChart(scopedRows);
  renderDashboardEventTypeChart(scopedRows, selectedMonthKey);
}

function renderDashboardSellerList(rows, periodLabel = "") {
  if (!el.dashboardSellerList) return;
  const scopeUserId = dashboardResolveScopeUserId();
  const role = dashboardResolveRole();
  const roleLabel = userRoleLabel(role);
  const sellers = dashboardGetActiveUsersByRole(role).filter((u) => !scopeUserId || String(u.id) === scopeUserId);
  if (!sellers.length) {
    el.dashboardSellerList.innerHTML = `<div class="dashboardEmpty">No hay usuarios activos del rol ${escapeHtml(roleLabel.toLowerCase())} para mostrar.</div>`;
    return;
  }

  const periodRows = Array.isArray(rows) ? rows : [];
  const metrics = sellers
    .map((seller) => {
      const sellerRows = periodRows.filter((r) => String(r.userId || "") === String(seller.id || ""));
      const confirmedRows = sellerRows.filter((r) => String(r.status || "") === STATUS.CONFIRMADO);
      const totalEvents = sellerRows.length;
      const confirmedEvents = confirmedRows.length;
      const totalAmount = sellerRows.reduce((acc, r) => acc + Math.max(0, Number(r.total || 0)), 0);
      const confirmedAmount = confirmedRows.reduce((acc, r) => acc + Math.max(0, Number(r.total || 0)), 0);
      const averageTicket = totalEvents > 0 ? (totalAmount / totalEvents) : 0;
      return {
        id: String(seller.id || ""),
        name: String(seller.fullName || seller.name || roleLabel),
        avatar: String(seller.avatarDataUrl || "").trim() || avatarDataUri(seller.fullName || seller.name || roleLabel),
        totalEvents,
        confirmedEvents,
        totalAmount,
        confirmedAmount,
        averageTicket,
      };
    })
    .sort((a, b) => {
      const amountDiff = Number(b.confirmedAmount || 0) - Number(a.confirmedAmount || 0);
      if (amountDiff !== 0) return amountDiff;
      const eventDiff = Number(b.confirmedEvents || 0) - Number(a.confirmedEvents || 0);
      if (eventDiff !== 0) return eventDiff;
      return String(a.name || "").localeCompare(String(b.name || ""), "es", { sensitivity: "base" });
    });

  const maxAmount = Math.max(1, ...metrics.map((item) => Number(item.confirmedAmount || 0)));
  const cardsHtml = metrics.map((item) => {
    const amount = Math.max(0, Number(item.confirmedAmount || 0));
    const heightPct = amount > 0 ? Math.max(8, (amount / maxAmount) * 100) : 4;
    const tooltip = `${item.name} | Eventos del periodo: ${item.totalEvents} | Confirmados: ${item.confirmedEvents} | Total cotizado: ${moneyGT(item.totalAmount)} | Total confirmado: ${moneyGT(item.confirmedAmount)} | Ticket promedio: ${moneyGT(item.averageTicket)}`;
    return `
      <article class="dashboardSellerPerfCard" data-dashboard-tooltip="${escapeHtml(tooltip)}">
        <div class="dashboardSellerPerfValue">${escapeHtml(moneyGT(item.confirmedAmount))}</div>
        <div class="dashboardSellerPerfBarArea">
          <span class="dashboardSellerPerfBar${amount > 0 ? "" : " is-empty"}" style="height:${heightPct}%"></span>
        </div>
        <div class="dashboardSellerPerfAvatar"><img alt="Avatar ${escapeHtml(roleLabel.toLowerCase())}" src="${escapeHtml(item.avatar)}"></div>
        <div class="dashboardSellerPerfName">${escapeHtml(item.name)}</div>
        <div class="dashboardSellerPerfMeta">${escapeHtml(String(item.confirmedEvents))} confirmados de ${escapeHtml(String(item.totalEvents))} evento(s)</div>
      </article>
    `;
  }).join("");

  el.dashboardSellerList.innerHTML = `
    <div class="dashboardSellerPerfWrap">
      <div class="dashboardSellerPerfHead">
        <strong>Eficiencia en confirmacion por ${escapeHtml(roleLabel.toLowerCase())}</strong>
        <small>${escapeHtml(periodLabel)} | Barra = monto confirmado</small>
      </div>
      <div class="dashboardSellerPerfGrid">
        ${cardsHtml || `<div class="dashboardEmpty">Sin datos para los filtros seleccionados.</div>`}
      </div>
      <div class="dashboardSellerPerfLegend">
        <span class="dashboardPieLegendItem" data-dashboard-tooltip="Suma de eventos en estado Confirmado del periodo seleccionado.">
          <i style="background:#10c972"></i>
          <span>Monto confirmado del periodo</span>
        </span>
      </div>
    </div>
  `;
}
function renderDashboardReport() {
  syncDashboardPeriodControls();
  const period = dashboardResolvePeriod();
  const monthKey = dashboardResolveMonthKey();
  const fromIso = String(el.dashboardReportFrom?.value || "").trim();
  const toIso = String(el.dashboardReportTo?.value || "").trim();
  if (period === "week") {
    const weekRaw = String(el.dashboardReportWeek?.value || "").trim();
    if (!/^\d{4}-\d{2}-\d{2}$/.test(weekRaw)) {
      setDashboardFiltersByWeek(weekInputFromDate(new Date()));
    } else {
      setDashboardFiltersByWeek(weekRaw);
    }
  } else if (!/^\d{4}-\d{2}-\d{2}$/.test(fromIso) || !/^\d{4}-\d{2}-\d{2}$/.test(toIso) || fromIso > toIso) {
    setDashboardFiltersByMonth(monthKey);
  }
  const safeFrom = String(el.dashboardReportFrom?.value || "").trim();
  const safeTo = String(el.dashboardReportTo?.value || "").trim();
  const safeMonthKey = dashboardMonthKeyFromIso(safeFrom) || monthKey;
  const periodLabel = dashboardFormatRangeLabel(period, safeFrom, safeTo, safeMonthKey);
  const role = dashboardResolveRole();
  const roleLabel = userRoleLabel(role);
  const authUser = normalizeUserRecord(getAuthUserRecord() || {});
  if (String(el.dashboardReportScope?.value || "mine") === "mine" && normalizeUserRole(authUser.role) !== role) {
    if (el.dashboardReportScope) el.dashboardReportScope.value = "all";
  }
  const rows = buildDashboardReportRows(safeFrom, safeTo).filter((r) => {
    const user = normalizeUserRecord((state.users || []).find((u) => String(u.id || "") === String(r.userId || "")) || {});
    return normalizeUserRole(user.role) === role;
  });
  const periodRows = rows.filter((r) => {
    const date = String(r?.eventDate || "").trim();
    if (!date) return false;
    if (safeFrom && date < safeFrom) return false;
    if (safeTo && date > safeTo) return false;
    return true;
  });
  renderDashboardSellerFilterOptions();
  const scope = String(el.dashboardReportScope?.value || "mine").trim();
  if (el.dashboardReportSeller) el.dashboardReportSeller.disabled = scope !== "seller";
  const globalGoal = (getGlobalMonthlyGoals({ role }).find((g) => String(g.month) === safeMonthKey)?.amount) || 0;
  const globalAchieved = periodRows.filter((r) => isGoalStatus(r.status)).reduce((acc, r) => acc + Number(r.total || 0), 0);
  const scopeUserId = dashboardResolveScopeUserId();
  const focusedUser = (state.users || []).map(normalizeUserRecord).find((u) => String(u.id) === scopeUserId) || null;
  const personalGoal = focusedUser
    ? ((focusedUser.monthlyGoals || []).find((g) => String(g.month) === safeMonthKey)?.amount || 0)
    : 0;
  const personalAchieved = focusedUser
    ? periodRows
      .filter((r) => String(r.userId) === String(focusedUser.id) && isGoalStatus(r.status))
      .reduce((acc, r) => acc + Number(r.total || 0), 0)
    : 0;
  renderDashboardGoals(
    {
      monthKey: safeMonthKey,
      period,
      periodLabel,
      roleKey: role,
      roleLabel,
      statusSummary: buildDashboardHeroStatusSummary(periodRows),
      goal: globalGoal,
      achieved: globalAchieved,
      remaining: Math.max(0, globalGoal - globalAchieved),
      progress: globalGoal > 0 ? (globalAchieved / globalGoal) * 100 : 0,
    },
    {
      sellerLabel: focusedUser ? (focusedUser.fullName || focusedUser.name || roleLabel) : `Selecciona ${roleLabel.toLowerCase()}`,
      goal: personalGoal,
      achieved: personalAchieved,
      remaining: Math.max(0, personalGoal - personalAchieved),
      progress: personalGoal > 0 ? (personalAchieved / personalGoal) * 100 : 0,
    },
  );
  renderDashboardCharts(periodRows, safeMonthKey);
  renderDashboardSellerList(periodRows, periodLabel);
  if (el.dashboardReportTitle) el.dashboardReportTitle.textContent = `Reporte ${roleLabel}`;
  if (el.dashboardReportSubtitle) {
    const viewLabel = period === "week" ? "Vista semanal" : "Vista mensual";
    el.dashboardReportSubtitle.textContent = `${viewLabel} de ${userRolePluralLabel(role).toLowerCase()}, metas y comparativos`;
  }
}

function resetDashboardReportFilters() {
  const currentMonth = toISODate(new Date()).slice(0, 7);
  if (el.dashboardReportPeriod) el.dashboardReportPeriod.value = "month";
  syncDashboardPeriodControls();
  setDashboardFiltersByMonth(currentMonth);
  if (el.dashboardReportWeek) {
    el.dashboardReportWeek.value = weekInputFromDate(new Date());
  }
  const authUser = normalizeUserRecord(getAuthUserRecord() || {});
  const sessionRole = REPORTABLE_USER_ROLES.includes(normalizeUserRole(authUser.role)) ? normalizeUserRole(authUser.role) : USER_ROLE.SELLER;
  if (el.dashboardReportRole) {
    el.dashboardReportRole.value = sessionRole;
  }
  if (el.dashboardReportScope) {
    el.dashboardReportScope.value = String(authSession.userId || "").trim() ? "mine" : "all";
  }
  renderDashboardSellerFilterOptions();
  if (el.dashboardReportSeller) {
    const sessionUserId = String(authSession.userId || "").trim();
    el.dashboardReportSeller.value = sessionUserId;
  }
}

function openDashboardReportModal() {
  if (!el.dashboardReportBackdrop) return;
  setSideNavActive("dashboard");
  resetDashboardReportFilters();
  renderDashboardReport();
  el.dashboardReportBackdrop.hidden = false;
}

function closeDashboardReportModal() {
  if (!el.dashboardReportBackdrop) return;
  el.dashboardReportBackdrop.hidden = true;
  restoreModuleScreenAfterModal();
  if (el.settingsScreen && !el.settingsScreen.hidden) setSideNavActive("settings");
  else if (el.reportsHubScreen && !el.reportsHubScreen.hidden) setSideNavActive("reports");
  else setSideNavActive("calendar");
}

function getInstitutionReportDefaultRange() {
  const now = new Date();
  const start = new Date(now.getFullYear(), 0, 1);
  start.setHours(0, 0, 0, 0);
  return {
    fromIso: toISODate(start),
    toIso: toISODate(now),
  };
}

function setInstitutionReportDefaultRange() {
  const { fromIso, toIso } = getInstitutionReportDefaultRange();
  if (el.institutionReportFrom) el.institutionReportFrom.value = fromIso;
  if (el.institutionReportTo) el.institutionReportTo.value = toIso;
}

function getInstitutionReportFilteredCompanies() {
  const search = String(el.institutionReportCompanySearch?.value || "").trim();
  return (state.companies || [])
    .filter((c) => {
      if (!String(c?.id || "").trim() || isCompanyDisabled(c.id)) return false;
      if (!search) return true;
      const haystack = [
        c?.name,
        c?.owner,
        c?.email,
        c?.phone,
        c?.billTo,
        c?.businessName,
      ].filter(Boolean).join(" ");
      return matchesLikeSearch(haystack, search);
    })
    .slice()
    .sort((a, b) => String(a?.name || "").localeCompare(String(b?.name || ""), "es", { sensitivity: "base" }));
}

function renderInstitutionReportCompanyResults(companies = []) {
  if (!el.institutionReportCompanyResults) return;
  const selectedId = String(el.institutionReportCompany?.value || "").trim();
  const search = String(el.institutionReportCompanySearch?.value || "").trim();
  if (!companies.length) {
    el.institutionReportCompanyResults.innerHTML = search
      ? '<div class="institutionSearchEmpty">No se encontraron instituciones con ese criterio.</div>'
      : "";
    return;
  }
  el.institutionReportCompanyResults.innerHTML = companies.slice(0, 6).map((company) => {
    const companyId = String(company.id || "").trim();
    const metaPrimary = String(company.owner || company.email || company.businessName || "Sin contacto principal").trim();
    const metaSecondary = [company.phone, company.email].filter(Boolean).join(" | ");
    const activeClass = companyId === selectedId ? " isActive" : "";
    return `
      <button class="institutionSearchCard${activeClass}" type="button" data-institution-company-pick="${escapeHtml(companyId)}">
        <span class="institutionSearchCardTitle">${escapeHtml(String(company.name || "Institucion").trim())}</span>
        <span class="institutionSearchCardMeta">${escapeHtml(metaPrimary)}</span>
        <span class="institutionSearchCardSub">${escapeHtml(metaSecondary || "Sin telefono o correo")}</span>
      </button>
    `;
  }).join("");
}

function renderInstitutionReportCompanyOptions() {
  if (!el.institutionReportCompany) return;
  const previous = String(el.institutionReportCompany.value || "").trim();
  const companies = getInstitutionReportFilteredCompanies();
  el.institutionReportCompany.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = companies.length ? "Selecciona institucion" : "Sin coincidencias";
  el.institutionReportCompany.appendChild(placeholder);
  for (const company of companies) {
    const opt = document.createElement("option");
    const detail = [company.owner || company.email || "Sin contacto", company.phone || "Sin telefono"]
      .filter(Boolean)
      .join(" | ");
    opt.value = String(company.id || "").trim();
    opt.textContent = `${String(company.name || "Institucion").trim()} - ${detail}`;
    el.institutionReportCompany.appendChild(opt);
  }
  if (previous && companies.some((c) => String(c.id) === previous)) {
    el.institutionReportCompany.value = previous;
  } else if (companies.length) {
    el.institutionReportCompany.value = String(companies[0].id || "");
  } else {
    el.institutionReportCompany.value = "";
  }
  syncEnhancedSelectValue(el.institutionReportCompany, el.institutionReportCompany.value || "");
  queueSelectEnhancement(el.institutionReportCompany, true);
  renderInstitutionReportCompanyResults(companies);
}

function normalizeInstitutionItemLabel(item) {
  const primary = String(item?.name || "").trim();
  const secondary = String(item?.description || "").trim();
  return primary || secondary || "Servicio sin nombre";
}

function buildInstitutionMetricListHtml(rows, emptyText, formatter) {
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) {
    return `<div class="dashboardEmpty">${escapeHtml(emptyText)}</div>`;
  }
  return list.map((row, idx) => formatter(row, idx)).join("");
}

function buildInstitutionMonthlyBarChartHtml(monthRows) {
  const rows = (Array.isArray(monthRows) ? monthRows : [])
    .slice()
    .sort((a, b) => String(a?.monthKey || "").localeCompare(String(b?.monthKey || "")))
    .slice(-8);
  if (!rows.length) {
    return `<div class="dashboardEmpty">Sin meses con actividad para graficar.</div>`;
  }
  const maxAmount = Math.max(1, ...rows.map((r) => Math.max(0, Number(r?.amount || 0))));
  const bars = rows.map((row) => {
    const amount = Math.max(0, Number(row?.amount || 0));
    const count = Math.max(0, Number(row?.count || 0));
    const pct = Math.max(8, Math.round((amount / maxAmount) * 100));
    const label = String(row?.label || "-");
    const compact = label.replace(/\s+de\s+/i, " ").trim();
    const tip = `${label} | ${count} reserva(s) | ${moneyGT(amount)}`;
    return `
      <div class="institutionBarCol">
        <div class="institutionBarValue">${escapeHtml(moneyGT(amount))}</div>
        <div class="institutionBarTrack">
          <div class="institutionBarFill" style="height:${pct}%;" tabindex="0">
            <span class="institutionChartTip">${escapeHtml(tip)}</span>
          </div>
        </div>
        <div class="institutionBarLabel">${escapeHtml(compact)}</div>
      </div>
    `;
  }).join("");
  return `<div class="institutionBarChart">${bars}</div>`;
}

function buildInstitutionStatusDonutHtml(data) {
  const segments = [
    { label: "Confirmados", value: Math.max(0, Number(data?.confirmed || 0)), color: "#22c55e" },
    { label: "Pre reserva", value: Math.max(0, Number(data?.preReserved || 0)), color: "#38bdf8" },
    { label: "Cancelados", value: Math.max(0, Number(data?.canceled || 0)), color: "#f97316" },
    { label: "Perdidos", value: Math.max(0, Number(data?.lost || 0)), color: "#f43f5e" },
  ].filter((x) => x.value > 0);
  const total = segments.reduce((acc, x) => acc + x.value, 0);
  if (!total) {
    return `<div class="dashboardEmpty">Sin estados suficientes para graficar.</div>`;
  }
  let cursor = 0;
  const stops = segments.map((seg) => {
    const start = cursor;
    const pct = (seg.value / total) * 100;
    cursor += pct;
    return `${seg.color} ${start.toFixed(2)}% ${cursor.toFixed(2)}%`;
  }).join(", ");
  const legend = segments.map((seg) => {
    const pct = (seg.value / total) * 100;
    const tip = `${seg.label}: ${seg.value} reserva(s) (${pct.toFixed(1)}%)`;
    return `
      <div class="institutionDonutLegendItem" tabindex="0" style="min-width:0; overflow-wrap:anywhere;">
        <span class="institutionDonutDot" style="background:${seg.color}"></span>
        <b>${escapeHtml(seg.label)}</b>
        <span>${escapeHtml(String(seg.value))} (${escapeHtml(pct.toFixed(1))}%)</span>
        <span class="institutionChartTip">${escapeHtml(tip)}</span>
      </div>
    `;
  }).join("");
  return `
    <div class="institutionDonutWrap" style="display:flex; align-items:center; justify-content:center; gap:18px; flex-wrap:wrap; min-width:0;">
      <div class="institutionDonut" style="background:conic-gradient(${stops}); width:min(230px, 70vw); height:min(230px, 70vw); max-width:100%; flex:0 0 auto; margin:0 auto;">
        <div class="institutionDonutHole" style="width:58%; height:58%;">
          <small>Total</small>
          <strong>${escapeHtml(String(total))}</strong>
        </div>
      </div>
      <div class="institutionDonutLegend" style="flex:1 1 240px; min-width:0; width:100%;">${legend}</div>
    </div>
  `;
}
function buildInstitutionReportData(companyId, fromIso, toIso) {
  const id = String(companyId || "").trim();
  const company = (state.companies || []).find((c) => String(c?.id || "") === id) || null;
  const empty = {
    company,
    reservations: 0,
    eventRows: 0,
    totalRevenue: 0,
    totalPax: 0,
    avgTicket: 0,
    avgPax: 0,
    confirmed: 0,
    preReserved: 0,
    canceled: 0,
    lost: 0,
    conversionPct: 0,
    firstVisitIso: "",
    latestVisitIso: "",
    daysSinceLastVisit: null,
    topSalonLabel: "",
    topDishLabel: "",
    topManagerLabel: "",
    topMonthLabel: "",
    topSellerLabel: "",
    activeMonths: 0,
    salonRows: [],
    dishRows: [],
    managerRows: [],
    sellerRows: [],
    monthRows: [],
    eventRowsDetailed: [],
  };
  if (!company) return empty;

  const grouped = new Map();
  for (const ev of state.events || []) {
    if (!quoteBelongsToCompany(ev?.quote, company)) continue;
    const key = String(reservationKeyFromEvent(ev) || ev?.id || "").trim();
    if (!key) continue;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(ev);
  }

  const salonCounter = new Map();
  const dishCounter = new Map();
  const managerCounter = new Map();
  const sellerCounter = new Map();
  const monthCounter = new Map();
  const monthRevenueCounter = new Map();
  const allDates = [];

  for (const seriesRaw of grouped.values()) {
    const series = (seriesRaw || []).slice().sort((a, b) => {
      const dateCmp = String(a?.date || "").localeCompare(String(b?.date || ""));
      if (dateCmp !== 0) return dateCmp;
      return compareTime(String(a?.startTime || "00:00"), String(b?.startTime || "00:00"));
    });
    if (!series.length) continue;

    for (const ev of series) {
      const d = String(ev?.date || "").trim();
      if (d) allDates.push(d);
    }

    const seriesInRange = series.filter((ev) => {
      const date = String(ev?.date || "").trim();
      if (!date) return false;
      if (fromIso && date < fromIso) return false;
      if (toIso && date > toIso) return false;
      return true;
    });
    if (!seriesInRange.length) continue;

    const firstInRange = seriesInRange[0];
    const quoteHost = series.find((ev) => quoteBelongsToCompany(ev?.quote, company)) || firstInRange;
    const latestSnapshot = getLatestQuoteSnapshotFromSeries(series) || getLatestQuoteSnapshotForEvent(quoteHost) || quoteHost?.quote || {};
    const totals = getQuoteTotals(latestSnapshot || {});
    const firstMonthKey = dashboardMonthKeyFromIso(firstInRange?.date || "");
    const currentStatus = String(firstInRange?.status || "").trim();
    const userName = getUserNameById(firstInRange?.userId || "");
    const managerName = String(
      latestSnapshot?.managerName
      || latestSnapshot?.contact
      || quoteHost?.quote?.managerName
      || "Sin encargado"
    ).trim();

    empty.reservations += 1;
    empty.totalRevenue += Math.max(0, Number(totals.total || 0));
    if (currentStatus === STATUS.CONFIRMADO) empty.confirmed += 1;
    if (currentStatus === STATUS.PRERESERVA) empty.preReserved += 1;
    if (currentStatus === STATUS.CANCELADO) empty.canceled += 1;
    if (currentStatus === STATUS.PERDIDO) empty.lost += 1;

    managerCounter.set(managerName, Number(managerCounter.get(managerName) || 0) + 1);
    sellerCounter.set(userName || "Sin vendedor", Number(sellerCounter.get(userName || "Sin vendedor") || 0) + 1);
    if (firstMonthKey) {
      monthCounter.set(firstMonthKey, Number(monthCounter.get(firstMonthKey) || 0) + 1);
      monthRevenueCounter.set(firstMonthKey, Number(monthRevenueCounter.get(firstMonthKey) || 0) + Math.max(0, Number(totals.total || 0)));
    }

    const items = Array.isArray(latestSnapshot?.items) ? latestSnapshot.items : [];
    for (const item of items) {
      const label = normalizeInstitutionItemLabel(item);
      const qty = Math.max(0, Number(item?.qty || 0));
      const amount = qty * Math.max(0, Number(item?.price || 0));
      const current = dishCounter.get(label) || { label, qty: 0, amount: 0 };
      current.qty += qty;
      current.amount += amount;
      dishCounter.set(label, current);
    }

    for (const ev of seriesInRange) {
      const date = String(ev?.date || "").trim();
      const salon = String(ev?.salon || "").trim() || "Sin salon";
      const pax = Math.max(0, Number(ev?.pax || latestSnapshot?.people || 0));
      empty.eventRows += 1;
      empty.totalPax += pax;
      salonCounter.set(salon, Number(salonCounter.get(salon) || 0) + 1);
      empty.eventRowsDetailed.push({
        status: String(ev?.status || ""),
        statusColor: statusColor(ev?.status),
        reservationKey: String(latestSnapshot?.code || reservationKeyFromEvent(ev) || ev?.id || "-"),
        eventDate: date,
        eventName: String(ev?.name || "Reserva"),
        salon,
        managerName,
        pax,
        total: Math.max(0, Number(totals.total || 0)),
      });
    }
  }

  allDates.sort();
  empty.firstVisitIso = allDates[0] || "";
  empty.latestVisitIso = allDates[allDates.length - 1] || "";
  if (empty.latestVisitIso) {
    const latest = new Date(`${empty.latestVisitIso}T00:00:00`);
    const today = stripTime(new Date());
    if (!Number.isNaN(latest.getTime())) {
      empty.daysSinceLastVisit = Math.max(0, Math.floor((today.getTime() - latest.getTime()) / (1000 * 60 * 60 * 24)));
    }
  }

  empty.avgTicket = empty.reservations > 0 ? empty.totalRevenue / empty.reservations : 0;
  empty.avgPax = empty.eventRows > 0 ? empty.totalPax / empty.eventRows : 0;
  empty.conversionPct = empty.reservations > 0 ? ((empty.confirmed + empty.preReserved) / empty.reservations) * 100 : 0;
  empty.activeMonths = monthCounter.size;

  empty.salonRows = Array.from(salonCounter.entries())
    .map(([label, count]) => ({ label, count: Number(count || 0) }))
    .sort((a, b) => Number(b.count || 0) - Number(a.count || 0) || String(a.label || "").localeCompare(String(b.label || ""), "es", { sensitivity: "base" }));
  empty.dishRows = Array.from(dishCounter.values())
    .sort((a, b) => Number(b.qty || 0) - Number(a.qty || 0) || Number(b.amount || 0) - Number(a.amount || 0));
  empty.managerRows = Array.from(managerCounter.entries())
    .map(([label, count]) => ({ label, count: Number(count || 0) }))
    .sort((a, b) => Number(b.count || 0) - Number(a.count || 0) || String(a.label || "").localeCompare(String(b.label || ""), "es", { sensitivity: "base" }));
  empty.sellerRows = Array.from(sellerCounter.entries())
    .map(([label, count]) => ({ label, count: Number(count || 0) }))
    .sort((a, b) => Number(b.count || 0) - Number(a.count || 0) || String(a.label || "").localeCompare(String(b.label || ""), "es", { sensitivity: "base" }));
  empty.monthRows = Array.from(monthCounter.entries())
    .map(([monthKey, count]) => {
      const [yy, mm] = String(monthKey || "").split("-").map(Number);
      const d = Number.isFinite(yy) && Number.isFinite(mm) ? new Date(yy, mm - 1, 1) : null;
      return {
        monthKey,
        label: d ? fmtMonthYear(d) : monthKey,
        count: Number(count || 0),
        amount: Math.max(0, Number(monthRevenueCounter.get(monthKey) || 0)),
      };
    })
    .sort((a, b) => Number(b.amount || 0) - Number(a.amount || 0) || Number(b.count || 0) - Number(a.count || 0));
  empty.eventRowsDetailed.sort((a, b) => {
    const d = String(b.eventDate || "").localeCompare(String(a.eventDate || ""));
    if (d !== 0) return d;
    return String(a.eventName || "").localeCompare(String(b.eventName || ""), "es", { sensitivity: "base" });
  });

  empty.topSalonLabel = empty.salonRows[0]?.label || "";
  empty.topDishLabel = empty.dishRows[0]?.label || "";
  empty.topManagerLabel = empty.managerRows[0]?.label || "";
  empty.topMonthLabel = empty.monthRows[0]?.label || "";
  empty.topSellerLabel = empty.sellerRows[0]?.label || "";
  return empty;
}

function renderInstitutionReport() {
  const companyId = String(el.institutionReportCompany?.value || "").trim();
  const fromIso = String(el.institutionReportFrom?.value || "").trim();
  const toIso = String(el.institutionReportTo?.value || "").trim();
  if (fromIso && toIso && fromIso > toIso) {
    if (el.institutionReportTo) el.institutionReportTo.value = fromIso;
  }
  const safeFrom = String(el.institutionReportFrom?.value || "").trim();
  const safeTo = String(el.institutionReportTo?.value || "").trim();
  const data = buildInstitutionReportData(companyId, safeFrom, safeTo);

  if (el.institutionReportHeadline) {
    const companyName = String(data.company?.name || "Institucion").trim() || "Institucion";
    const companyType = String(
      data.company?.businessName
      || data.company?.billTo
      || data.company?.owner
      || "Selecciona una institucion para ver su tablero"
    ).trim();
    const contactLine = [
      data.company?.owner || data.company?.email || "Sin contacto principal",
      data.company?.phone || "Sin telefono",
    ].filter(Boolean).join(" | ");
    const lastVisitText = data.latestVisitIso
      ? fmtDateShort(new Date(`${data.latestVisitIso}T00:00:00`))
      : "Sin visitas registradas";
    const firstVisitText = data.firstVisitIso
      ? fmtDateShort(new Date(`${data.firstVisitIso}T00:00:00`))
      : "Sin historial previo";
    const rangeText = safeFrom && safeTo
      ? `${fmtDateShort(new Date(`${safeFrom}T00:00:00`))} al ${fmtDateShort(new Date(`${safeTo}T00:00:00`))}`
      : "Define un rango para analizar a tus clientes";
    el.institutionReportHeadline.classList.toggle("isEmpty", !data.company);
    if (!data.company) {
      el.institutionReportHeadline.innerHTML = `
        <div class="institutionHeadlineEmpty">
          <div class="institutionHeadlineEmptyMain">
            <span class="institutionHeroEyebrow">Vista premium</span>
            <strong>Selecciona una institucion para desbloquear el dashboard</strong>
            <small>Usa el buscador, el selector o ajusta el rango para mostrar consumo, encargados, salones y tendencia historica en una sola vista.</small>
          </div>
          <div class="institutionHeadlineHintGrid">
            <article class="institutionHeadlineHintCard">
              <span>Periodo activo</span>
              <b>${escapeHtml(rangeText)}</b>
            </article>
            <article class="institutionHeadlineHintCard">
              <span>Instituciones disponibles</span>
              <b>${escapeHtml(String((state.companies || []).filter((c) => !isCompanyDisabled(c?.id)).length))}</b>
            </article>
            <article class="institutionHeadlineHintCard">
              <span>Que veras aqui</span>
              <b>Ingresos, historial, estados y consumo</b>
            </article>
          </div>
        </div>
      `;
    } else {
      el.institutionReportHeadline.innerHTML = `
        <div class="institutionHeadlineTop">
          <div class="institutionHeadlineIdentity">
            <span class="institutionHeroEyebrow">Cliente destacado</span>
            <strong>${escapeHtml(companyName)}</strong>
            <small>${escapeHtml(companyType)} | ${escapeHtml(contactLine)}</small>
          </div>
          <div class="institutionHeadlineValueBlock">
            <span class="institutionHeadlineValueLabel">Ingreso acumulado</span>
            <div class="institutionHeadlineMoney">${escapeHtml(moneyGT(data.totalRevenue || 0))}</div>
          </div>
        </div>
        <div class="institutionHeadlineStats">
          <article class="institutionStatCard institutionStatPrimary">
            <span class="institutionStatLabel">Ultima visita</span>
            <strong>${escapeHtml(lastVisitText)}</strong>
          </article>
          <article class="institutionStatCard institutionStatTeal">
            <span class="institutionStatLabel">Primera visita</span>
            <strong>${escapeHtml(firstVisitText)}</strong>
          </article>
          <article class="institutionStatCard institutionStatAmber">
            <span class="institutionStatLabel">Encargado top</span>
            <strong>${escapeHtml(data.topManagerLabel || "-")}</strong>
          </article>
          <article class="institutionStatCard institutionStatMagenta">
            <span class="institutionStatLabel">Vendedor top</span>
            <strong>${escapeHtml(data.topSellerLabel || "-")}</strong>
          </article>
        </div>
      `;
    }
  }

  if (el.institutionReportSummary) {
    if (!data.company) {
      el.institutionReportSummary.innerHTML = "";
    } else {
      const cards = [
        { label: "Eventos", value: String(data.eventRows || 0), meta: "Registros en el rango", target: "institutionSectionEvents" },
        { label: "Reservas", value: String(data.reservations || 0), meta: "Reservas consolidadas", target: "institutionSectionOverview" },
        { label: "PAX total", value: String(data.totalPax || 0), meta: `Promedio ${Math.round(data.avgPax || 0)}`, target: "institutionSectionOverview" },
        { label: "Salon top", value: data.topSalonLabel || "-", meta: "Mas usado por frecuencia", target: "institutionSectionSalons" },
        { label: "Platillo top", value: data.topDishLabel || "-", meta: "Mas pedido", target: "institutionSectionDishes" },
        { label: "Mes mas fuerte", value: data.topMonthLabel || "-", meta: "Mayor monto del periodo", target: "institutionSectionTimeline" },
        { label: "Encargado top", value: data.topManagerLabel || "-", meta: "Mas eventos generados", target: "institutionSectionManagers" },
        { label: "Conversion", value: `${Math.round(data.conversionPct || 0)}%`, meta: "Confirmado + pre reserva", target: "institutionSectionOverview" },
      ];
      el.institutionReportSummary.innerHTML = cards.map((card) => `
        <button class="dashboardGoalCard institutionSummaryCard" type="button" data-target-section="${escapeHtml(card.target)}">
          <small>${escapeHtml(card.label)}</small>
          <strong>${escapeHtml(card.value)}</strong>
          <div class="dashboardGoalMeta">${escapeHtml(card.meta)}</div>
        </button>
      `).join("");
    }
  }

  if (el.institutionOverviewGrid) {
    if (!data.company) {
      el.institutionOverviewGrid.innerHTML = buildInstitutionReportEmptyStateHtml(
        "Selecciona una institucion",
        "Al elegir una institucion veras ingresos, ticket promedio, conversion y actividad historica."
      );
    } else {
      const overviewRows = [
        { label: "Ingreso total", value: moneyGT(data.totalRevenue || 0) },
        { label: "Ticket promedio", value: moneyGT(data.avgTicket || 0) },
        { label: "Confirmados", value: String(data.confirmed || 0) },
        { label: "Pre reserva", value: String(data.preReserved || 0) },
        { label: "Cancelados", value: String(data.canceled || 0) },
        { label: "Perdidos", value: String(data.lost || 0) },
        { label: "Meses activos", value: String(data.activeMonths || 0) },
        { label: "Dias desde ultima visita", value: Number.isFinite(data.daysSinceLastVisit) ? String(data.daysSinceLastVisit) : "-" },
      ];
      el.institutionOverviewGrid.innerHTML = overviewRows.map((row) => `
        <div class="dashboardStatusChip institutionOverviewChip">
          <b>${escapeHtml(row.label)}</b>
          <span>${escapeHtml(row.value)}</span>
        </div>
      `).join("");
    }
  }

  if (el.institutionReportChartsBody) {
    if (!data.company) {
      el.institutionReportChartsBody.innerHTML = buildInstitutionReportEmptyStateHtml(
        "Graficas listas para analizar",
        "Selecciona una institucion y el sistema armara las graficas de ingresos por mes y distribucion por estado."
      );
    } else {
      el.institutionReportChartsBody.innerHTML = `
        <article class="institutionChartCard" style="min-width:0; overflow:hidden; display:flex; flex-direction:column; gap:14px;">
          <header>
            <strong>Ingresos por mes</strong>
            <small>Top 8 meses del rango (hover para detalle)</small>
          </header>
          ${buildInstitutionMonthlyBarChartHtml(data.monthRows)}
        </article>
        <article class="institutionChartCard" style="min-width:0; overflow:hidden; display:flex; flex-direction:column; gap:14px;">
          <header>
            <strong>Distribucion por estado</strong>
            <small>Confirmado, pre reserva, cancelado y perdido</small>
          </header>
          ${buildInstitutionStatusDonutHtml(data)}
        </article>
      `;
    }
  }
  if (el.institutionReportSalonBody) {
    if (!data.company) {
      el.institutionReportSalonBody.innerHTML = buildInstitutionReportEmptyStateHtml(
        "Sin salon seleccionado aun",
        "Aqui apareceran los salones mas usados por la institucion dentro del rango que definas."
      );
    } else {
      el.institutionReportSalonBody.innerHTML = buildInstitutionMetricListHtml(
        data.salonRows.slice(0, 8),
        "Sin uso de salones para el rango seleccionado.",
        (row, idx) => `
          <article class="institutionMetricCard">
            <strong>#${idx + 1} ${escapeHtml(row.label || "-")}</strong>
            <span>${escapeHtml(String(row.count || 0))} uso(s)</span>
          </article>
        `
      );
    }
  }

  if (el.institutionReportDishBody) {
    if (!data.company) {
      el.institutionReportDishBody.innerHTML = buildInstitutionReportEmptyStateHtml(
        "Consumo pendiente de cargar",
        "Cuando selecciones una institucion veras aqui los platillos y servicios mas cotizados."
      );
    } else {
      el.institutionReportDishBody.innerHTML = buildInstitutionMetricListHtml(
        data.dishRows.slice(0, 10),
        "Sin platillos o servicios cotizados para este rango.",
        (row, idx) => `
          <article class="institutionMetricCard">
            <strong>#${idx + 1} ${escapeHtml(row.label || "-")}</strong>
            <span>${escapeHtml(String(row.qty || 0))} unidad(es) | ${escapeHtml(moneyGT(row.amount || 0))}</span>
          </article>
        `
      );
    }
  }

  if (el.institutionReportManagerBody) {
    if (!data.company) {
      el.institutionReportManagerBody.innerHTML = buildInstitutionReportEmptyStateHtml(
        "Encargados por mostrar",
        "El panel listara a las personas que mas eventos han coordinado con nosotros para esa institucion."
      );
    } else {
      el.institutionReportManagerBody.innerHTML = buildInstitutionMetricListHtml(
        data.managerRows.slice(0, 8),
        "Sin encargados asociados en el rango.",
        (row, idx) => `
          <article class="institutionMetricCard">
            <strong>#${idx + 1} ${escapeHtml(row.label || "-")}</strong>
            <span>${escapeHtml(String(row.count || 0))} reserva(s)</span>
          </article>
        `
      );
    }
  }

  if (el.institutionReportTimelineBody) {
    if (!data.company) {
      el.institutionReportTimelineBody.innerHTML = buildInstitutionReportEmptyStateHtml(
        "Historial aun no disponible",
        "Selecciona una institucion para revisar meses fuertes, ultima visita y seguimiento comercial."
      );
    } else {
      const monthHtml = buildInstitutionMetricListHtml(
        data.monthRows.slice(0, 6),
        "No hay meses con actividad en el rango.",
        (row, idx) => `
          <article class="institutionMetricCard">
            <strong>#${idx + 1} ${escapeHtml(row.label || "-")}</strong>
            <span>${escapeHtml(String(row.count || 0))} reserva(s) | ${escapeHtml(moneyGT(row.amount || 0))}</span>
          </article>
        `
      );
      const sellerHtml = buildInstitutionMetricListHtml(
        data.sellerRows.slice(0, 4),
        "Sin vendedores con actividad.",
        (row, idx) => `
          <article class="institutionMetricCard">
            <strong>#${idx + 1} ${escapeHtml(row.label || "-")}</strong>
            <span>${escapeHtml(String(row.count || 0))} reserva(s)</span>
          </article>
        `
      );
      el.institutionReportTimelineBody.innerHTML = `
        <div class="institutionTimelineCol">
          <div class="institutionTimelineLabel">Meses mas fuertes</div>
          ${monthHtml}
        </div>
        <div class="institutionTimelineCol">
          <div class="institutionTimelineLabel">Vendedores con mas seguimiento</div>
          ${sellerHtml}
        </div>
      `;
    }
  }

  if (el.institutionReportEventsBody) {
    el.institutionReportEventsBody.innerHTML = "";
    if (!data.company || !data.eventRowsDetailed.length) {
      const tr = document.createElement("tr");
      tr.innerHTML = `<td colspan="9">Sin eventos para los filtros seleccionados.</td>`;
      el.institutionReportEventsBody.appendChild(tr);
    } else {
      for (const row of data.eventRowsDetailed.slice(0, 120)) {
        const tr = document.createElement("tr");
        tr.innerHTML = `
          <td><span class="salesStatusBadge" style="background:${hexToRgba(row.statusColor, 0.25)};border-color:${hexToRgba(row.statusColor, 0.6)}">${escapeHtml(row.status || "-")}</span></td>
          <td>${escapeHtml(row.reservationKey || "-")}</td>
          <td>${escapeHtml(row.eventDate || "-")}</td>
          <td>${escapeHtml(row.eventName || "-")}</td>
          <td>${escapeHtml(row.salon || "-")}</td>
          <td>${escapeHtml(row.managerName || "-")}</td>
          <td>${escapeHtml(String(row.pax || 0))}</td>
          <td>${escapeHtml(moneyGT(row.total || 0))}</td>
          <td>${escapeHtml(data.latestVisitIso || "-")}</td>
        `;
        el.institutionReportEventsBody.appendChild(tr);
      }
    }
  }
}

function buildInstitutionReportEmptyStateHtml(title, text) {
  return `
    <div class="institutionPanelEmpty">
      <strong>${escapeHtml(title)}</strong>
      <span>${escapeHtml(text)}</span>
    </div>
  `;
}

function scrollInstitutionReportToSection(sectionId) {
  const id = String(sectionId || "").trim();
  if (!id) return;
  const node = document.getElementById(id);
  if (!node) return;
  try {
    node.scrollIntoView({ behavior: "smooth", block: "start" });
  } catch (_) {
    node.scrollIntoView();
  }
}

function resetInstitutionReportFilters() {
  if (el.institutionReportCompanySearch) el.institutionReportCompanySearch.value = "";
  renderInstitutionReportCompanyOptions();
  setInstitutionReportDefaultRange();
}

function openInstitutionReportModal() {
  if (!el.institutionReportBackdrop) return;
  resetInstitutionReportFilters();
  renderInstitutionReport();
  el.institutionReportBackdrop.hidden = false;
}

function closeInstitutionReportModal() {
  if (!el.institutionReportBackdrop) return;
  el.institutionReportBackdrop.hidden = true;
  restoreModuleScreenAfterModal();
}

function renderChecklistTemplateTable() {
  if (!el.checklistTemplateBody) return;
  checklistTemplateDraft = (checklistTemplateDraft || []).map(normalizeChecklistTemplateItem).filter(Boolean);
  const sectionMap = new Map((checklistTemplateSectionsDraft || []).map((s) => [String(s.id || ""), s]));
  el.checklistTemplateBody.innerHTML = "";
  if (!checklistTemplateDraft.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="5">Sin puntos configurados.</td>`;
    el.checklistTemplateBody.appendChild(tr);
    return;
  }
  checklistTemplateDraft.forEach((item, idx) => {
    const section = sectionMap.get(String(item.sectionId || ""));
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>${escapeHtml(String(section?.name || "General"))}</td>
      <td>${escapeHtml(String(item.label || ""))}</td>
      <td>
        <button class="btn" type="button" data-checklist-template-up="${escapeHtml(String(item.id || ""))}">Subir</button>
        <button class="btn" type="button" data-checklist-template-down="${escapeHtml(String(item.id || ""))}">Bajar</button>
      </td>
      <td>
        <button class="apptIconBtn apptEdit" type="button" data-checklist-template-edit="${escapeHtml(String(item.id || ""))}" title="Editar plantilla" aria-label="Editar plantilla">&#9998;</button>
        <button class="apptIconBtn apptDelete" type="button" data-checklist-template-remove="${escapeHtml(String(item.id || ""))}" title="Eliminar plantilla" aria-label="Eliminar plantilla">&#128465;</button>
      </td>
    `;
    el.checklistTemplateBody.appendChild(tr);
  });
}

function renderChecklistSectionSelect(selected = "") {
  if (!el.checklistTemplateSectionSelect) return;
  const sections = checklistTemplateSectionsDraft.length
    ? checklistTemplateSectionsDraft
    : getChecklistSections();
  el.checklistTemplateSectionSelect.innerHTML = "";
  for (const s of sections) {
    const opt = document.createElement("option");
    opt.value = s;
    opt.textContent = s;
    el.checklistTemplateSectionSelect.appendChild(opt);
  }
  const preferred = String(selected || "").trim();
  if (preferred) el.checklistTemplateSectionSelect.value = preferred;
  if (!el.checklistTemplateSectionSelect.value && el.checklistTemplateSectionSelect.options.length) {
    el.checklistTemplateSectionSelect.value = el.checklistTemplateSectionSelect.options[0].value;
  }
}

function resetChecklistTemplateEditor() {
  checklistTemplateEditingId = "";
  if (el.btnChecklistTemplateAdd) el.btnChecklistTemplateAdd.textContent = "Agregar punto";
  if (el.checklistTemplateInput) el.checklistTemplateInput.value = "";
  renderChecklistSectionSelect("General");
}

function openChecklistTemplateModal() {
  ensureChecklistStores();
  checklistTemplateDraft = getChecklistTemplateItems().map((x) => ({ ...x }));
  checklistTemplateSectionsDraft = getChecklistSections().slice();
  renderChecklistTemplateTable();
  resetChecklistTemplateEditor();
  if (el.checklistTemplateSectionInput) el.checklistTemplateSectionInput.value = "";
  if (el.checklistTemplateBackdrop) el.checklistTemplateBackdrop.hidden = false;
}

function closeChecklistTemplateModal() {
  if (el.checklistTemplateBackdrop) el.checklistTemplateBackdrop.hidden = true;
  resetChecklistTemplateEditor();
}

function saveChecklistTemplateDraft() {
  ensureChecklistStores();
  state.checklistTemplateItems = (checklistTemplateDraft || []).map(normalizeChecklistTemplateItem).filter(Boolean);
  state.checklistTemplateSections = Array.from(new Set((checklistTemplateSectionsDraft || []).map((s) => String(s || "").trim()).filter(Boolean)));
  persist();
}

function addChecklistSectionFromInput() {
  const name = String(el.checklistTemplateSectionInput?.value || "").trim();
  if (!name) return toast("Escribe el nombre de la seccion.");
  const exists = (checklistTemplateSectionsDraft || []).some((s) => String(s || "").trim().toLowerCase() === name.toLowerCase());
  if (exists) return toast("Esa seccion ya existe.");
  checklistTemplateSectionsDraft.push(name);
  saveChecklistTemplateDraft();
  renderChecklistSectionSelect(name);
  if (el.checklistTemplateSectionInput) {
    el.checklistTemplateSectionInput.value = "";
    el.checklistTemplateSectionInput.focus();
  }
  toast("Seccion agregada.");
}

function addChecklistTemplateItemFromInput() {
  const label = String(el.checklistTemplateInput?.value || "").trim();
  const section = String(el.checklistTemplateSectionSelect?.value || "General").trim() || "General";
  if (!label) return toast("Escribe un punto para el check list.");
  const wasEditing = !!checklistTemplateEditingId;
  const exists = (checklistTemplateDraft || []).some((x) => {
    const sameLabel = String(x?.label || "").trim().toLowerCase() === label.toLowerCase();
    const sameId = String(x?.id || "") === String(checklistTemplateEditingId || "");
    return sameLabel && !sameId;
  });
  if (exists) return toast("Ese punto ya existe en el check list.");
  if (checklistTemplateEditingId) {
    const idx = (checklistTemplateDraft || []).findIndex((x) => String(x?.id || "") === String(checklistTemplateEditingId));
    if (idx >= 0) checklistTemplateDraft[idx] = { ...checklistTemplateDraft[idx], label, section };
  } else {
    checklistTemplateDraft.push({ id: uid(), label, section, active: true });
  }
  saveChecklistTemplateDraft();
  renderChecklistTemplateTable();
  resetChecklistTemplateEditor();
  if (el.checklistTemplateInput) el.checklistTemplateInput.focus();
  toast(wasEditing ? "Punto actualizado." : "Punto agregado al check list.");
}

function getCurrentChecklistTemplateDraft() {
  const currentId = String(checklistTemplateCurrentId || "").trim();
  return (checklistTemplatesDraft || []).find((tpl) => String(tpl?.id || "") === currentId) || null;
}

function renderChecklistTemplateSelect(selected = "") {
  if (!el.checklistTemplateSelect) return;
  const keep = String(selected || "").trim();
  el.checklistTemplateSelect.innerHTML = "";
  for (const tpl of checklistTemplatesDraft || []) {
    const opt = document.createElement("option");
    opt.value = String(tpl?.id || "").trim();
    if (!opt.value) continue;
    opt.textContent = `${String(tpl?.name || "Checklist").trim()}${tpl?.active === false ? " (Inhabilitada)" : ""}`;
    el.checklistTemplateSelect.appendChild(opt);
  }
  el.checklistTemplateSelect.value = keep;
}

function renderChecklistTemplateTable() {
  if (!el.checklistTemplateBody) return;
  checklistTemplateDraft = (checklistTemplateDraft || []).map(normalizeChecklistTemplateItem).filter(Boolean);
  const sectionMap = new Map((checklistTemplateSectionsDraft || []).map((s) => [String(s.id || ""), s]));
  el.checklistTemplateBody.innerHTML = "";
  if (!checklistTemplateDraft.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="5">Sin puntos configurados.</td>`;
    el.checklistTemplateBody.appendChild(tr);
    return;
  }
  checklistTemplateDraft.forEach((item, idx) => {
    const section = sectionMap.get(String(item.sectionId || ""));
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td>${escapeHtml(String(section?.name || "General"))}</td>
      <td>${escapeHtml(String(item.label || ""))}</td>
      <td>
        <button class="btn" type="button" data-checklist-template-up="${escapeHtml(String(item.id || ""))}">Subir</button>
        <button class="btn" type="button" data-checklist-template-down="${escapeHtml(String(item.id || ""))}">Bajar</button>
      </td>
      <td>
        <button class="apptIconBtn apptEdit" type="button" data-checklist-template-edit="${escapeHtml(String(item.id || ""))}" title="Editar plantilla" aria-label="Editar plantilla">&#9998;</button>
        <button class="apptIconBtn apptDelete" type="button" data-checklist-template-remove="${escapeHtml(String(item.id || ""))}" title="Eliminar plantilla" aria-label="Eliminar plantilla">&#128465;</button>
      </td>
    `;
    el.checklistTemplateBody.appendChild(tr);
  });
}

function renderChecklistSectionSelect(selected = "") {
  if (!el.checklistTemplateSectionSelect) return;
  const sections = (checklistTemplateSectionsDraft || []).filter((s) => s?.active !== false);
  el.checklistTemplateSectionSelect.innerHTML = "";
  for (const s of sections) {
    const opt = document.createElement("option");
    opt.value = String(s?.id || "").trim();
    opt.textContent = String(s?.name || "General").trim();
    el.checklistTemplateSectionSelect.appendChild(opt);
  }
  const preferred = String(selected || "").trim();
  if (preferred) el.checklistTemplateSectionSelect.value = preferred;
  if (!el.checklistTemplateSectionSelect.value && el.checklistTemplateSectionSelect.options.length) {
    el.checklistTemplateSectionSelect.value = el.checklistTemplateSectionSelect.options[0].value;
  }
}

function renderChecklistSectionEditSelect(selected = "") {
  if (!el.checklistTemplateSectionEditSelect) return;
  const keep = String(selected || "").trim();
  el.checklistTemplateSectionEditSelect.innerHTML = `<option value="">Nueva seccion</option>`;
  for (const s of checklistTemplateSectionsDraft || []) {
    const opt = document.createElement("option");
    opt.value = String(s?.id || "").trim();
    if (!opt.value) continue;
    opt.textContent = `${String(s?.name || "Seccion").trim()}${s?.active === false ? " (Inhabilitada)" : ""}`;
    el.checklistTemplateSectionEditSelect.appendChild(opt);
  }
  el.checklistTemplateSectionEditSelect.value = keep;
}

function renderChecklistSectionsTable() {
  if (!el.checklistTemplateSectionsBody) return;
  el.checklistTemplateSectionsBody.innerHTML = "";
  if (!(checklistTemplateSectionsDraft || []).length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="3">Sin secciones configuradas.</td>`;
    el.checklistTemplateSectionsBody.appendChild(tr);
    return;
  }
  for (const row of checklistTemplateSectionsDraft || []) {
    const rowId = String(row?.id || "").trim();
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(String(row?.name || "Seccion"))}</td>
      <td>${row?.active === false ? "Inhabilitada" : "Activa"}</td>
      <td class="appointmentActions">
        <button class="apptIconBtn apptEdit" type="button" data-checklist-section-edit="${escapeHtml(rowId)}" title="Editar" aria-label="Editar">&#9998;</button>
        <button class="apptIconBtn ${row?.active === false ? "" : "apptDelete"}" type="button" data-checklist-section-toggle="${escapeHtml(rowId)}" title="${row?.active === false ? "Reactivar" : "Inhabilitar"}" aria-label="${row?.active === false ? "Reactivar" : "Inhabilitar"}">${row?.active === false ? "&#8635;" : "&#9940;"}</button>
      </td>
    `;
    el.checklistTemplateSectionsBody.appendChild(tr);
  }
}

function resetChecklistSectionEditor() {
  checklistTemplateSectionEditingId = "";
  if (el.checklistTemplateSectionInput) el.checklistTemplateSectionInput.value = "";
  renderChecklistSectionEditSelect("");
}

function resetChecklistTemplateEditor() {
  checklistTemplateEditingId = "";
  if (el.btnChecklistTemplateAdd) el.btnChecklistTemplateAdd.textContent = "Agregar punto";
  if (el.checklistTemplateInput) el.checklistTemplateInput.value = "";
  renderChecklistSectionSelect("");
}

function loadChecklistTemplateEditor(templateId = "") {
  const fallback = (checklistTemplatesDraft || [])[0] || null;
  const target = (checklistTemplatesDraft || []).find((tpl) => String(tpl?.id || "") === String(templateId || "").trim()) || fallback;
  if (!target) return;
  checklistTemplateCurrentId = String(target.id || "").trim();
  checklistTemplateSectionsDraft = Array.isArray(target.sections) ? deepClone(target.sections) : [];
  checklistTemplateDraft = Array.isArray(target.items) ? deepClone(target.items) : [];
  if (el.checklistTemplateName) el.checklistTemplateName.value = String(target.name || "");
  if (el.checklistTemplateActive) el.checklistTemplateActive.checked = target.active !== false;
  if (el.btnChecklistTemplateDisable) {
    el.btnChecklistTemplateDisable.disabled = false;
    el.btnChecklistTemplateDisable.textContent = target.active === false ? "Reactivar" : "Inhabilitar";
  }
  renderChecklistTemplateSelect(checklistTemplateCurrentId);
  renderChecklistSectionSelect("");
  renderChecklistSectionEditSelect("");
  renderChecklistSectionsTable();
  renderChecklistTemplateTable();
  resetChecklistTemplateEditor();
  resetChecklistSectionEditor();
}

function openChecklistTemplateModal() {
  ensureChecklistStores();
  checklistTemplatesDraft = getChecklistTemplates({ includeInactive: true }).map((tpl) => deepClone(tpl));
  if (!checklistTemplatesDraft.length) {
    checklistTemplatesDraft = [normalizeChecklistTemplateRecord({
      name: "Checklist general",
      active: true,
      sections: [{ name: "General", active: true }],
      items: [],
    })];
  }
  loadChecklistTemplateEditor(checklistTemplatesDraft[0]?.id || "");
  if (el.checklistTemplateBackdrop) el.checklistTemplateBackdrop.hidden = false;
}

function closeChecklistTemplateModal() {
  if (checklistTemplatesDraft.length) {
    saveChecklistTemplateDraft();
  }
  if (el.checklistTemplateBackdrop) el.checklistTemplateBackdrop.hidden = true;
  checklistTemplatesDraft = [];
  checklistTemplateCurrentId = "";
  resetChecklistTemplateEditor();
  resetChecklistSectionEditor();
  restoreModuleScreenAfterModal();
}

function saveChecklistTemplateDraft() {
  ensureChecklistStores();
  const current = getCurrentChecklistTemplateDraft();
  if (current) {
    current.name = String(el.checklistTemplateName?.value || "").trim() || "Checklist";
    current.active = el.checklistTemplateActive?.checked !== false;
    current.sections = (checklistTemplateSectionsDraft || []).map(normalizeChecklistSectionRecord).filter(Boolean);
    current.items = (checklistTemplateDraft || []).map(normalizeChecklistTemplateItem).filter(Boolean);
  }
  state.checklistTemplates = (checklistTemplatesDraft || []).map(normalizeChecklistTemplateRecord).filter(Boolean);
  syncLegacyChecklistStateFromTemplates();
  persist();
}

function addChecklistSectionFromInput() {
  const name = String(el.checklistTemplateSectionInput?.value || "").trim();
  if (!name) return toast("Escribe el nombre de la seccion.");
  const exists = (checklistTemplateSectionsDraft || []).some((s) =>
    String(s?.name || "").trim().toLowerCase() === name.toLowerCase()
    && String(s?.id || "") !== String(checklistTemplateSectionEditingId || "")
  );
  if (exists) return toast("Esa seccion ya existe.");
  const wasEditing = !!checklistTemplateSectionEditingId;
  if (wasEditing) {
    const idx = (checklistTemplateSectionsDraft || []).findIndex((s) => String(s?.id || "") === String(checklistTemplateSectionEditingId || ""));
    if (idx >= 0) checklistTemplateSectionsDraft[idx] = { ...checklistTemplateSectionsDraft[idx], name };
  } else {
    checklistTemplateSectionsDraft.push({ id: uid(), name, active: true });
  }
  saveChecklistTemplateDraft();
  renderChecklistSectionSelect("");
  renderChecklistSectionEditSelect("");
  renderChecklistSectionsTable();
  renderChecklistTemplateTable();
  if (el.checklistTemplateSectionInput) el.checklistTemplateSectionInput.focus();
  resetChecklistSectionEditor();
  toast(wasEditing ? "Seccion actualizada." : "Seccion agregada.");
}

function addChecklistTemplateItemFromInput() {
  const label = String(el.checklistTemplateInput?.value || "").trim();
  const sectionId = String(el.checklistTemplateSectionSelect?.value || "").trim();
  if (!label) return toast("Escribe un punto para el check list.");
  if (!sectionId) return toast("Selecciona una seccion.");
  const wasEditing = !!checklistTemplateEditingId;
  const exists = (checklistTemplateDraft || []).some((x) => {
    const sameLabel = String(x?.label || "").trim().toLowerCase() === label.toLowerCase();
    const sameId = String(x?.id || "") === String(checklistTemplateEditingId || "");
    return sameLabel && !sameId;
  });
  if (exists) return toast("Ese punto ya existe en el check list.");
  if (checklistTemplateEditingId) {
    const idx = (checklistTemplateDraft || []).findIndex((x) => String(x?.id || "") === String(checklistTemplateEditingId));
    if (idx >= 0) checklistTemplateDraft[idx] = { ...checklistTemplateDraft[idx], label, sectionId };
  } else {
    checklistTemplateDraft.push({ id: uid(), label, sectionId, active: true });
  }
  saveChecklistTemplateDraft();
  renderChecklistTemplateTable();
  resetChecklistTemplateEditor();
  if (el.checklistTemplateInput) el.checklistTemplateInput.focus();
  toast(wasEditing ? "Punto actualizado." : "Punto agregado al check list.");
}

function updateEventChecklistProgress() {
  const items = Array.isArray(eventChecklistDraft?.items) ? eventChecklistDraft.items : [];
  const total = items.length;
  const answered = items.filter((it) => normalizeChecklistStatus(it?.status) !== "").length;
  const okCount = items.filter((it) => normalizeChecklistStatus(it?.status) === "ok").length;
  const progressPct = total > 0 ? Math.round((answered / total) * 100) : 0;
  const satisfactionPct = total > 0 ? Math.round((okCount / total) * 100) : 0;
  if (el.eventChecklistProgressLabel) {
    el.eventChecklistProgressLabel.textContent = `Resultado general ${satisfactionPct}%`;
  }
  if (el.eventChecklistSatisfactionLabel) {
    el.eventChecklistSatisfactionLabel.textContent = `Avance respondido ${progressPct}%`;
  }
  if (el.eventChecklistProgressFill) {
    const safePct = Math.max(0, Math.min(100, satisfactionPct));
    el.eventChecklistProgressFill.style.width = `${safePct}%`;
    el.eventChecklistProgressFill.style.background =
      safePct >= 85
        ? "linear-gradient(90deg, #10b981, #34d399)"
        : (safePct >= 60
          ? "linear-gradient(90deg, #f59e0b, #fbbf24)"
          : "linear-gradient(90deg, #ef4444, #f87171)");
  }
}

function renderEventChecklistRows() {
  if (!el.eventChecklistBody) return;
  const items = Array.isArray(eventChecklistDraft?.items) ? eventChecklistDraft.items : [];
  el.eventChecklistBody.innerHTML = "";
  if (!items.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="4">No hay puntos configurados. Usa "Agregar Check List" en Configuracion.</td>`;
    el.eventChecklistBody.appendChild(tr);
    return;
  }
  let lastSection = "";
  items.forEach((item, idx) => {
    const section = String(item?.section || "General").trim() || "General";
    if (section !== lastSection) {
      const sectionTr = document.createElement("tr");
      sectionTr.className = "checklistSectionRow";
      sectionTr.innerHTML = `<td colspan="4">${escapeHtml(section)}</td>`;
      el.eventChecklistBody.appendChild(sectionTr);
      lastSection = section;
    }
    const status = normalizeChecklistStatus(item?.status);
    const statusCls = status === "ok"
      ? "checklistStateBtn checklistStateBtn--ok"
      : (status === "x"
        ? "checklistStateBtn checklistStateBtn--x"
        : (status === "na"
          ? "checklistStateBtn checklistStateBtn--na"
          : "checklistStateBtn checklistStateBtn--pending"));
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${idx + 1}</td>
      <td><b>${escapeHtml(String(item?.section || "General"))}</b> - ${escapeHtml(String(item?.label || ""))}</td>
      <td>
        <button type="button" class="${statusCls}" data-checklist-cycle-index="${idx}" data-checklist-state="${escapeHtml(status)}">
          ${escapeHtml(checklistStatusBadgeText(status))} - ${escapeHtml(checklistStatusLabel(status))}
        </button>
      </td>
      <td>
        <input type="text" class="quoteInput" data-checklist-comment-index="${idx}" value="${escapeHtml(String(item?.comment || ""))}" placeholder="Comentario adicional" />
      </td>
    `;
    el.eventChecklistBody.appendChild(tr);
  });
  updateEventChecklistProgress();
}

function renderEventChecklistTemplateSelect(selected = "") {
  if (!el.eventChecklistTemplateSelect) return;
  const keep = String(selected || "").trim();
  const rows = getChecklistTemplates({ includeInactive: true });
  el.eventChecklistTemplateSelect.innerHTML = "";
  for (const tpl of rows) {
    const opt = document.createElement("option");
    opt.value = String(tpl?.id || "").trim();
    if (!opt.value) continue;
    opt.textContent = `${String(tpl?.name || "Checklist").trim()}${tpl?.active === false ? " (Inhabilitada)" : ""}`;
    el.eventChecklistTemplateSelect.appendChild(opt);
  }
  el.eventChecklistTemplateSelect.value = keep;
}

function openEventChecklistModal(eventId) {
  const draft = buildEventChecklistDraft(eventId);
  if (!draft) return toast("No se pudo abrir el check list del evento.");
  currentEventChecklistId = String(eventId || "").trim();
  eventChecklistDraft = draft;
  renderEventChecklistTemplateSelect(draft.templateKey || "");
  if (el.eventChecklistDate) el.eventChecklistDate.value = draft.eventDate || "";
  if (el.eventChecklistEventName) el.eventChecklistEventName.value = draft.eventName || "";
  if (el.eventChecklistSubtitle) {
    el.eventChecklistSubtitle.textContent = `${draft.eventName || "-"} | ${draft.eventDate || "-"} | ${draft.salon || "-"}`;
  }
  if (el.eventChecklistNotes) el.eventChecklistNotes.value = String(draft.notes || "");
  renderEventChecklistRows();
  if (el.eventChecklistBackdrop) el.eventChecklistBackdrop.hidden = false;
}

function closeEventChecklistModal() {
  if (el.eventChecklistBackdrop) el.eventChecklistBackdrop.hidden = true;
  currentEventChecklistId = "";
  eventChecklistDraft = null;
}

function saveEventChecklistFromModal() {
  if (!eventChecklistDraft || !currentEventChecklistId) return;
  ensureChecklistStores();
  const notes = String(el.eventChecklistNotes?.value || "").trim();
  const items = Array.isArray(eventChecklistDraft.items) ? eventChecklistDraft.items : [];
  const normalizedItems = items.map((it) => ({
    id: String(it?.id || uid()).trim(),
    templateId: String(it?.templateId || "").trim(),
    label: String(it?.label || "").trim(),
    section: String(it?.section || "General").trim() || "General",
    status: normalizeChecklistStatus(it?.status),
    comment: String(it?.comment || "").trim(),
  })).filter((it) => it.label);
  const nowIso = new Date().toISOString();
  const completed = normalizedItems.length > 0 && normalizedItems.every((it) => ["ok", "x", "na"].includes(it.status));
  state.eventChecklists[currentEventChecklistId] = {
    eventId: currentEventChecklistId,
    templateKey: String(eventChecklistDraft.templateKey || "").trim(),
    templateName: String(eventChecklistDraft.templateName || "").trim(),
    notes,
    items: normalizedItems,
    updatedAt: nowIso,
    completedAt: completed ? nowIso : "",
  };
  persist();
  renderOccupancyReportTable();
  toast(completed ? "Check list completado y guardado." : "Check list guardado.");
}

function applyTopbarSettings({ rerender = false } = {}) {
  if (el.settingShowLegend) el.settingShowLegend.checked = !!topbarSettings.showLegend;
  if (el.settingCompactEvents) el.settingCompactEvents.checked = !!topbarSettings.compactEvents;
  if (el.settingShowWeekends) el.settingShowWeekends.checked = !!topbarSettings.showWeekends;
  if (el.legend) el.legend.hidden = !topbarSettings.showLegend;

  if (rerender) render();
}

function closeOpenCustomTopbarSelect() {
  const active = uiEnhancers.openCustomSelect;
  if (!active) return;
  const comp = uiEnhancers.customTopbarSelects.get(active);
  if (!comp) return;
  comp.root.classList.remove("open");
  comp.menu.hidden = true;
  uiEnhancers.openCustomSelect = null;
}

function updateCustomTopbarSelectFromNative(select) {
  const comp = uiEnhancers.customTopbarSelects.get(select);
  if (!comp) return;
  const selectedOpt = select.options[select.selectedIndex] || null;
  comp.button.textContent = selectedOpt?.textContent?.trim() || "Selecciona";
  const items = comp.menu.querySelectorAll(".cselectItem");
  for (const item of items) {
    const isSelected = item.dataset.value === String(select.value || "");
    item.classList.toggle("selected", isSelected);
  }
}

function rebuildCustomTopbarSelectOptions(select) {
  const comp = uiEnhancers.customTopbarSelects.get(select);
  if (!comp) return;
  comp.menu.innerHTML = "";
  for (const opt of Array.from(select.options)) {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "cselectItem";
    item.textContent = String(opt.textContent || "").trim();
    item.dataset.value = String(opt.value || "");
    if (opt.disabled) item.disabled = true;
    item.addEventListener("click", () => {
      select.value = item.dataset.value || "";
      select.dispatchEvent(new Event("change", { bubbles: true }));
      updateCustomTopbarSelectFromNative(select);
      closeOpenCustomTopbarSelect();
    });
    comp.menu.appendChild(item);
  }
  updateCustomTopbarSelectFromNative(select);
}

function ensureCustomTopbarSelect(select) {
  if (!select) return;
  const already = uiEnhancers.customTopbarSelects.get(select);
  if (already) {
    rebuildCustomTopbarSelectOptions(select);
    return;
  }
  const root = document.createElement("div");
  root.className = "cselect";
  const button = document.createElement("button");
  button.type = "button";
  button.className = "cselectBtn";
  const menu = document.createElement("div");
  menu.className = "cselectMenu";
  menu.hidden = true;
  root.appendChild(button);
  root.appendChild(menu);
  select.style.display = "none";
  select.insertAdjacentElement("afterend", root);

  button.addEventListener("click", (e) => {
    e.stopPropagation();
    const isOpen = !menu.hidden;
    closeOpenCustomTopbarSelect();
    if (isOpen) return;
    root.classList.add("open");
    menu.hidden = false;
    uiEnhancers.openCustomSelect = select;
  });
  select.addEventListener("change", () => updateCustomTopbarSelectFromNative(select));

  uiEnhancers.customTopbarSelects.set(select, { root, button, menu });
  rebuildCustomTopbarSelectOptions(select);
}

function initCustomTopbarSelects() {
  ensureCustomTopbarSelect(el.navMode);
  ensureCustomTopbarSelect(el.roomSelect);
}

function enhanceSelectControl(select, force = false) {
  if (!USE_ENHANCED_SELECTS) return;
  if (!select || typeof window.SlimSelect !== "function") return;
  if (select === el.navMode || select === el.roomSelect || select.closest?.(".topbar")) return;
  // Keep quote page selects native to avoid floating/overlapping slim-select wrappers.
  if (select.closest?.("#quoteBackdrop")) return;
  const existing = uiEnhancers.selectChoices.get(select);
  if (existing && !force) return;
  if (existing) {
    try { existing.destroy(); } catch (_) { }
    uiEnhancers.selectChoices.delete(select);
  }

  try {
    const choices = new window.SlimSelect({
      select,
      settings: {
        showSearch: true,
        openPosition: "auto",
        contentPosition: "fixed",
        contentLocation: document.body,
        closeOnSelect: true,
        placeholderText: "Selecciona opcion",
        searchPlaceholder: "Escribe para buscar...",
        searchText: "Sin resultados",
      },
    });
    const ssMain = select.nextElementSibling;
    if (ssMain && ssMain.classList?.contains("ss-main")) {
      ssMain.classList.add("ss-glass");
      if (select.classList.contains("quoteInput") || select.closest(".quoteTable")) {
        ssMain.classList.add("ss-compact");
      }
    }
    uiEnhancers.selectChoices.set(select, choices);
  } catch (err) {
    console.warn("SlimSelect error on", select?.id || select?.name || "select", err?.message || err);
  }
}

function syncEnhancedSelectValue(select, value = "") {
  if (!select) return;
  const next = String(value || "");
  select.value = next;
  const instance = uiEnhancers.selectChoices.get(select);
  if (instance && typeof instance.setSelected === "function") {
    try {
      instance.setSelected(next || "");
    } catch (_) { }
  }
}

function queueSelectEnhancement(select, force = false) {
  if (!select) return;
  const prev = uiEnhancers.selectQueue.get(select) === true;
  uiEnhancers.selectQueue.set(select, prev || force);
  if (uiEnhancers.selectQueueTimer) return;
  uiEnhancers.selectQueueTimer = setTimeout(() => {
    const batch = Array.from(uiEnhancers.selectQueue.entries());
    uiEnhancers.selectQueue.clear();
    uiEnhancers.selectQueueTimer = null;
    for (const [node, mustForce] of batch) {
      if (!node || !document.documentElement.contains(node)) continue;
      enhanceSelectControl(node, mustForce);
    }
  }, 0);
}

function queueSelectsInNode(node, force = false) {
  if (!node || node.nodeType !== 1) return;
  if (node.tagName === "SELECT") queueSelectEnhancement(node, force);
  const nested = node.querySelectorAll ? node.querySelectorAll("select") : [];
  for (const s of nested) queueSelectEnhancement(s, force);
}

function ensureSelectEnhancerObserver() {
  if (uiEnhancers.selectObserver || typeof MutationObserver !== "function") return;
  uiEnhancers.selectObserver = new MutationObserver((mutations) => {
    for (const m of mutations) {
      if (m.type !== "childList") continue;
      if (m.target && m.target.nodeType === 1 && m.target.tagName === "SELECT") {
        queueSelectEnhancement(m.target, true);
      }
      for (const n of m.addedNodes || []) {
        queueSelectsInNode(n, false);
      }
    }
  });
  uiEnhancers.selectObserver.observe(document.body, { childList: true, subtree: true });
}

function initEnhancedSelects() {
  if (!USE_ENHANCED_SELECTS) {
    if (uiEnhancers.selectObserver) {
      try { uiEnhancers.selectObserver.disconnect(); } catch (_) { }
      uiEnhancers.selectObserver = null;
    }
    for (const [node, instance] of Array.from(uiEnhancers.selectChoices.entries())) {
      try { instance.destroy(); } catch (_) { }
      uiEnhancers.selectChoices.delete(node);
    }
    return;
  }
  if (typeof window.SlimSelect !== "function") return;
  if (uiEnhancers.selectObserver) {
    try { uiEnhancers.selectObserver.disconnect(); } catch (_) { }
    uiEnhancers.selectObserver = null;
  }
  const selects = Array.from(document.querySelectorAll("select"));
  for (const s of selects) {
    try {
      enhanceSelectControl(s, true);
    } catch (_) { }
  }
  ensureSelectEnhancerObserver();
}

function getVisibleDayCount() {
  if (navMode === "day") return 1;
  if (navMode === "week" && !topbarSettings.showWeekends) return 5;
  return 7;
}

// -------- init ----------
try {
  goToTodayView();
  autoMarkLostEvents();
  renderTopbarWelcome();
  renderLegend();
  applyTopbarSettings();
  renderTimeColumn();
  renderRoomSelects();
  renderStatusSelect();
  renderUsersSelect();
  renderCompaniesSelect();
  renderServicesList();
  render();
  initModernDatePickers();
  runUpcomingReminderChecks();
  refreshTopbarReminders();
  setInterval(runUpcomingReminderChecks, 60 * 1000);
  setInterval(refreshTopbarReminders, 60 * 1000);

  bindEvents();
  initEnhancedSelects();
  initCustomTopbarSelects();
  syncWithServerState()
    .catch(() => { })
    .finally(() => {
      loadLoginUsers()
        .then(() => {
          const restored = restorePersistedAppSession();
          if (!restored && el.loginScreen) el.loginScreen.hidden = false;
        })
        .catch(() => {
          if (el.loginScreen) el.loginScreen.hidden = false;
          setLoginError("No se pudo cargar usuarios desde MariaDB.");
        });
    });
} catch (bootErr) {
  console.error("Fallo al iniciar app:", bootErr);
  if (el.loginScreen) el.loginScreen.hidden = false;
  setLoginError("Fallo al iniciar la app. Revisa consola.");
}

// ================== Rendering ==================

function render() {
  interaction.selecting = null;
  clearSelectionBox();
  const visibleDays = getVisibleDayCount();

  // Header label
  const end = addDays(viewStart, visibleDays - 1);
  if (navMode === "month") {
    el.weekLabel.textContent = fmtMonthYear(monthCursor);
  } else if (navMode === "day") {
    el.weekLabel.textContent = fmtDateShort(viewStart);
  } else {
    el.weekLabel.textContent = `${fmtDateShort(viewStart)} - ${fmtDateShort(end)}`;
  }

  const columnsTemplate = `repeat(${visibleDays}, minmax(240px, 1fr))`;
  el.daysHeader.style.gridTemplateColumns = columnsTemplate;
  el.grid.style.gridTemplateColumns = columnsTemplate;

  // Header days
  el.daysHeader.innerHTML = "";
  for (let i = 0; i < visibleDays; i++) {
    const d = addDays(viewStart, i);
    const head = document.createElement("div");
    head.className = "dayHead" + (isSameDay(d, new Date()) ? " today" : "");
    head.innerHTML = `
      <div class="dayName">${fmtWeekday(d)}</div>
      <div class="dayDate">${fmtDayMonth(d)}</div>
    `;
    el.daysHeader.appendChild(head);
  }
  syncCalendarVerticalOffset();

  // Grid columns with hour lines
  el.grid.innerHTML = "";
  for (let i = 0; i < visibleDays; i++) {
    const col = document.createElement("div");
    col.className = "dayCol";
    col.dataset.dayIndex = String(i);

    for (let h = HOUR_START; h <= HOUR_END + 1; h++) {
      const line = document.createElement("div");
      line.className = "hourLine";
      col.appendChild(line);
    }

    col.addEventListener("mousedown", (ev) => {
      if (ev.button !== 0) return;
      if (ev.target.closest(".event")) return;
      startSelection(ev, col);
    });

    el.grid.appendChild(col);
  }

  // Renderiza eventos en rango y salon seleccionado
  const visibleEvents = getEventsInWeek(viewStart, selectedSalon, visibleDays);
  const eventsByDay = new Map();
  for (const ev of visibleEvents) {
    const dayKey = String(ev?.date || "");
    if (!eventsByDay.has(dayKey)) eventsByDay.set(dayKey, []);
    eventsByDay.get(dayKey).push(ev);
  }
  for (let i = 0; i < visibleDays; i++) {
    const dayDate = toISODate(addDays(viewStart, i));
    const dayEvents = eventsByDay.get(dayDate) || [];
    const layoutMap = computeDayEventLayout(dayEvents);
    for (const e of dayEvents) {
      placeEvent(e, layoutMap.get(e.id));
    }
  }
  el.timeCol.scrollTop = el.grid.scrollTop;
  refreshTopbarReminders();
}

function syncCalendarVerticalOffset() {
  const headerHeight = Number(el.daysHeader?.offsetHeight || 0);
  if (!Number.isFinite(headerHeight) || headerHeight <= 0) return;
  document.documentElement.style.setProperty("--calendar-header-offset", `${headerHeight}px`);
}

function renderLegend() {
  el.legend.innerHTML = "";
  for (const s of STATUS_META) {
    const badge = document.createElement("div");
    badge.className = "badge";
    badge.innerHTML = `<span class="dot" style="background:${cssVar(s.colorVar)}"></span>${s.key}`;
    el.legend.appendChild(badge);
  }
}

function renderTimeColumn() {
  el.timeCol.innerHTML = "";
  for (let h = HOUR_START; h <= HOUR_END + 1; h++) {
    const slot = document.createElement("div");
    slot.className = "timeSlot";
    slot.textContent = formatHourAmPm(h);
    el.timeCol.appendChild(slot);
  }
}

function renderRoomSelects() {
  const disabledSalones = new Set((state.disabledSalones || []).map((x) => String(x).toLowerCase()));
  const activeSalones = (state.salones || []).filter((r) => !disabledSalones.has(String(r || "").toLowerCase()));
  // selector de salon en barra superior
  el.roomSelect.innerHTML = "";
  const optAll = document.createElement("option");
  optAll.value = ALL_ROOMS_VALUE;
  optAll.textContent = "Todos los salones";
  el.roomSelect.appendChild(optAll);

  for (const r of activeSalones) {
    const opt = document.createElement("option");
    opt.value = r; opt.textContent = r;
    el.roomSelect.appendChild(opt);
  }
  if (!activeSalones.includes(selectedSalon) && selectedSalon !== ALL_ROOMS_VALUE) {
    selectedSalon = ALL_ROOMS_VALUE;
  }
  el.roomSelect.value = selectedSalon;
  enhanceSelectControl(el.roomSelect, true);
  ensureCustomTopbarSelect(el.roomSelect);
  rerenderSlotRoomOptions();
}

function renderStatusSelect() {
  el.eventStatus.innerHTML = "";
  // Orden visual importante
  const order = [
    STATUS.RESERVA_SIN_COTIZACION,
    STATUS.PRIMERA,
    STATUS.PERDIDO,
    STATUS.SEGUIMIENTO,
    STATUS.LISTA,
    STATUS.PRERESERVA,
    STATUS.CONFIRMADO,
    STATUS.CANCELADO,
    STATUS.MANTENIMIENTO,
  ];

  for (const st of order) {
    const opt = document.createElement("option");
    opt.value = st;
    opt.textContent = st;
    if (isAutoStatus(st)) opt.disabled = true;
    el.eventStatus.appendChild(opt);
  }
  applyStatusSelectTheme();
}

function applyStatusSelectTheme() {
  if (!el.eventStatus) return;
  const color = statusColor(el.eventStatus.value);
  el.eventStatus.style.borderColor = hexToRgba(color, 0.6);
  el.eventStatus.style.background = `linear-gradient(135deg, ${hexToRgba(color, 0.32)}, rgba(255,255,255,0.06))`;
  el.eventStatus.style.boxShadow = `inset 0 0 0 1px ${hexToRgba(color, 0.28)}`;
}

function renderUsersSelect() {
  const previousValue = String(el.eventUser?.value || "").trim();
  el.eventUser.innerHTML = "";
  const activeUsers = (state.users || []).map(normalizeUserRecord).filter((u) => u.active !== false);
  for (const u of activeUsers) {
    const opt = document.createElement("option");
    opt.value = u.id;
    opt.textContent = u.fullName || u.name;
    el.eventUser.appendChild(opt);
  }
  const hasPrev = activeUsers.some((u) => String(u.id) === previousValue);
  const sessionUserExists = activeUsers.some((u) => String(u.id) === String(authSession.userId || ""));
  const fallback = sessionUserExists ? String(authSession.userId) : (activeUsers[0]?.id || "");
  const selected = hasPrev ? previousValue : fallback;
  enhanceSelectControl(el.eventUser, true);
  syncEnhancedSelectValue(el.eventUser, selected);
}

function setLoginError(message = "") {
  if (!el.loginError) return;
  el.loginError.textContent = String(message || "").trim();
}

function renderTopbarWelcome() {
  if (!el.topbarWelcome) return;
  const displayName = String(authSession.fullName || authSession.username || "").trim();
  el.topbarWelcome.textContent = displayName ? `Bienvenido - ${displayName}` : "Bienvenido -";
  if (el.topbarUserAvatar) {
    const avatar = String(authSession.avatarDataUrl || "").trim() || avatarDataUri(displayName || "Usuario");
    el.topbarUserAvatar.src = avatar;
  }
}

function updateLoginAvatarFromSelect() {
  if (!el.loginUserSelect || !el.loginAvatar) return;
  const selectedOpt = el.loginUserSelect.selectedOptions?.[0] || null;
  const avatar = String(selectedOpt?.dataset?.avatar || "").trim();
  const fallbackName = String(selectedOpt?.dataset?.name || "Usuario").trim();
  el.loginAvatar.src = avatar || avatarDataUri(fallbackName);
}

async function loadLoginUsers() {
  if (!el.loginUserSelect) return;
  const loginUsersUrl = buildApiUrlFromStateUrl(activeApiStateUrl, "login-users");
  const response = await fetch(loginUsersUrl, { cache: "no-store" });
  if (!response.ok) throw new Error("No se pudo cargar lista de usuarios.");
  const payload = await response.json();
  const users = Array.isArray(payload?.users) ? payload.users : [];
  el.loginUserSelect.innerHTML = "";
  for (const user of users) {
    const id = String(user?.id || "").trim();
    if (!id) continue;
    const fullName = String(user?.fullName || user?.name || "").trim() || "Usuario";
    const username = String(user?.username || "").trim();
    const opt = document.createElement("option");
    opt.value = id;
    opt.textContent = username ? `${username} - ${fullName}` : fullName;
    opt.dataset.avatar = String(user?.avatarDataUrl || "").trim();
    opt.dataset.name = fullName;
    el.loginUserSelect.appendChild(opt);
  }
  if (!el.loginUserSelect.options.length) {
    const emptyOpt = document.createElement("option");
    emptyOpt.value = "";
    emptyOpt.textContent = "No hay usuarios activos";
    el.loginUserSelect.appendChild(emptyOpt);
  }
  updateLoginAvatarFromSelect();
}

async function doLogin() {
  if (!el.loginForm || !el.loginScreen) return;
  const userId = String(el.loginUserSelect?.value || "").trim();
  const password = String(el.loginPassword?.value || "");
  if (!userId) {
    setLoginError("No hay usuarios activos.");
    return;
  }
  if (!password) {
    setLoginError("Ingresa la contrasena.");
    return;
  }
  setLoginError("");
  const loginUrl = buildApiUrlFromStateUrl(activeApiStateUrl, "login");
  const response = await fetch(loginUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ userId, password }),
  });
  const payload = await response.json().catch(() => ({}));
  if (!response.ok || payload?.ok !== true) {
    setLoginError(payload?.message || "Contrasena incorrecta.");
    return;
  }
  authSession = {
    userId: String(payload?.user?.id || userId),
    fullName: String(payload?.user?.fullName || payload?.user?.name || "").trim(),
    username: String(payload?.user?.username || "").trim(),
    avatarDataUrl: String(payload?.user?.avatarDataUrl || "").trim(),
    signatureDataUrl: String(payload?.user?.signatureDataUrl || "").trim(),
  };
  savePersistedAuthSession();
  renderTopbarWelcome();
  refreshTopbarReminders();
  el.loginScreen.hidden = true;
  showCalendarModule();
  if (el.loginPassword) el.loginPassword.value = "";
  if (el.eventUser && authSession.userId) {
    syncEnhancedSelectValue(el.eventUser, authSession.userId);
  }
  toast(`Sesion iniciada: ${authSession.fullName || authSession.username}`);
}

function salonOptionsHtml(selected = "", includePlaceholder = false) {
  const disabledSalones = new Set((state.disabledSalones || []).map((x) => String(x).toLowerCase()));
  const placeholder = includePlaceholder
    ? `<option value=""${selected ? "" : " selected"} disabled>Selecciona salon</option>`
    : "";
  const rows = (state.salones || [])
    .filter((r) => !disabledSalones.has(String(r || "").toLowerCase()) || String(r) === String(selected))
    .map(r => `<option value="${escapeHtml(r)}"${r === selected ? " selected" : ""}>${escapeHtml(r)}</option>`)
    .join("");
  return `${placeholder}${rows}`;
}

function initModernTimePicker(input) {
  if (!input) return;
  if (typeof window.flatpickr === "function") {
    if (input._flatpickr) input._flatpickr.destroy();
    input.classList.add("modernPickerInput", "modernTimeInput");
    window.flatpickr(input, {
      enableTime: true,
      noCalendar: true,
      dateFormat: "H:i",
      time_24hr: true,
      minuteIncrement: SNAP_MINUTES,
      allowInput: true,
      disableMobile: true,
      nextArrow: "<span aria-hidden=\"true\">&rsaquo;</span>",
      prevArrow: "<span aria-hidden=\"true\">&lsaquo;</span>",
      onReady: (_, __, instance) => {
        instance.calendarContainer?.classList.add("crmTimePicker");
        if (input.id === "appointmentTime") {
          instance.calendarContainer?.classList.add("appointmentPickerFlat", "appointmentTimePickerFlat");
        }
        input.addEventListener("focus", () => instance.open());
        input.addEventListener("click", () => instance.open());
      },
      onOpen: () => input.classList.add("modernPickerInputOpen"),
      onClose: () => input.classList.remove("modernPickerInputOpen"),
    });
    return;
  }
  input.type = "time";
  input.step = String(SNAP_MINUTES * 60);
}

function syncEventDateRangePickerConstraints() {
  const startPicker = el.eventDate?._flatpickr || null;
  const endPicker = el.eventDateEnd?._flatpickr || null;
  const startValue = String(el.eventDate?.value || "").trim();
  const endValue = String(el.eventDateEnd?.value || "").trim();

  if (endPicker) endPicker.set("minDate", startValue || null);
  if (startPicker) startPicker.set("maxDate", endValue || null);
}

function normalizeEventDateRange(changedInput) {
  const startValue = String(el.eventDate?.value || "").trim();
  const endValue = String(el.eventDateEnd?.value || "").trim();
  if (!startValue || !endValue) {
    syncEventDateRangePickerConstraints();
    return;
  }
  if (startValue <= endValue) {
    syncEventDateRangePickerConstraints();
    return;
  }
  if (changedInput === el.eventDate && el.eventDateEnd?._flatpickr) {
    el.eventDateEnd._flatpickr.setDate(startValue, true);
  } else if (changedInput === el.eventDateEnd && el.eventDate?._flatpickr) {
    el.eventDate._flatpickr.setDate(endValue, true);
  } else if (changedInput === el.eventDate) {
    el.eventDateEnd.value = startValue;
  } else if (changedInput === el.eventDateEnd) {
    el.eventDate.value = endValue;
  }
  syncEventDateRangePickerConstraints();
}

function initModernDatePicker(input) {
  if (!input) return;
  if (typeof window.flatpickr === "function") {
    if (input._flatpickr) input._flatpickr.destroy();
    input.type = "text";
    input.autocomplete = "off";
    input.classList.add("modernPickerInput", "modernDateInput");
    if (input.id === "eventDate") input.classList.add("modernDateInputStart");
    if (input.id === "eventDateEnd") input.classList.add("modernDateInputEnd");
    window.flatpickr(input, {
      dateFormat: "Y-m-d",
      allowInput: false,
      clickOpens: true,
      disableMobile: true,
      locale: window.flatpickr?.l10ns?.es || "default",
      monthSelectorType: "static",
      nextArrow: "<span aria-hidden=\"true\">&rsaquo;</span>",
      prevArrow: "<span aria-hidden=\"true\">&lsaquo;</span>",
      onChange: () => {
        normalizeEventDateRange(input);
        input.dispatchEvent(new Event("input", { bubbles: true }));
        input.dispatchEvent(new Event("change", { bubbles: true }));
      },
      onReady: (_, __, instance) => {
        instance.calendarContainer?.classList.add("crmDatePicker");
        if (input.id === "appointmentDate") {
          instance.calendarContainer?.classList.add("appointmentPickerFlat", "appointmentDatePickerFlat");
        }
        normalizeEventDateRange(input);
        input.addEventListener("focus", () => instance.open());
        input.addEventListener("click", () => instance.open());
      },
      onOpen: () => input.classList.add("modernPickerInputOpen"),
      onClose: () => input.classList.remove("modernPickerInputOpen"),
    });
    return;
  }
  input.type = "date";
}

function setModernDateValue(input, value, triggerChange = false) {
  if (!input) return;
  const clean = String(value || "").trim();
  if (input._flatpickr) {
    input._flatpickr.setDate(clean || null, !!triggerChange, "Y-m-d");
    return;
  }
  input.value = clean;
  if (triggerChange) {
    input.dispatchEvent(new Event("input", { bubbles: true }));
    input.dispatchEvent(new Event("change", { bubbles: true }));
  }
}

function initModernDatePickers() {
  for (const input of Array.from(document.querySelectorAll('input[type="date"]'))) {
    initModernDatePicker(input);
  }
}

function addSlotRow(slot = null) {
  const row = document.createElement("tr");
  row.className = "slotRow";
  const salon = slot?.salon || "";
  const paxValue = Number(slot?.slotPax ?? slot?.pax ?? 0);
  const pax = paxValue > 0 ? String(Math.round(paxValue)) : "";
  const slotDateStart = slot?.dateStart || slot?.startDate || el.eventDate?.value || "";
  const slotDateEnd = slot?.dateEnd || slot?.endDate || el.eventDateEnd?.value || slotDateStart;
  const start = slot?.startTime || "";
  const end = slot?.endTime || "";
  row.innerHTML = `
    <td><select class="quoteInput slotRoom">${salonOptionsHtml(salon, true)}</select></td>
    <td><input class="quoteInput slotPax" type="number" min="1" step="1" placeholder="PAX" value="${escapeHtml(pax)}" /></td>
    <td><input class="quoteInput slotDateStart" type="date" value="${escapeHtml(slotDateStart)}" /></td>
    <td><input class="quoteInput slotDateEnd" type="date" value="${escapeHtml(slotDateEnd)}" /></td>
    <td><input class="quoteInput slotStart" type="text" inputmode="numeric" placeholder="HH:mm" value="${escapeHtml(start)}" /></td>
    <td><input class="quoteInput slotEnd" type="text" inputmode="numeric" placeholder="HH:mm" value="${escapeHtml(end)}" /></td>
    <td><button type="button" class="btnDanger slotRemoveBtn">X</button></td>
  `;
  el.slotsBody.appendChild(row);
  const slotRoomSelect = row.querySelector(".slotRoom");
  initModernDatePicker(row.querySelector(".slotDateStart"));
  initModernDatePicker(row.querySelector(".slotDateEnd"));
  initModernTimePicker(row.querySelector(".slotStart"));
  initModernTimePicker(row.querySelector(".slotEnd"));
  queueSelectEnhancement(slotRoomSelect, true);
}

function rerenderSlotRoomOptions() {
  if (!el.slotsBody) return;
  for (const row of Array.from(el.slotsBody.querySelectorAll(".slotRow"))) {
    const select = row.querySelector(".slotRoom");
    if (!select) continue;
    const current = select.value;
    select.innerHTML = salonOptionsHtml(current, true);
    if (!select.value && select.options.length) select.value = select.options[0].value;
    queueSelectEnhancement(select, true);
    syncEnhancedSelectValue(select, select.value || "");
  }
}

function getSlotsFromForm() {
  const rows = Array.from(el.slotsBody.querySelectorAll(".slotRow"));
  return rows.map(row => ({
    salon: row.querySelector(".slotRoom")?.value || "",
    slotPax: Math.max(0, Number(row.querySelector(".slotPax")?.value || 0)) || 0,
    dateStart: row.querySelector(".slotDateStart")?.value || "",
    dateEnd: row.querySelector(".slotDateEnd")?.value || "",
    startTime: row.querySelector(".slotStart")?.value || "",
    endTime: row.querySelector(".slotEnd")?.value || "",
  }));
}

function normalizeSlotDateRange(slot, fallbackStart = "", fallbackEnd = "") {
  const rawStart = String(slot?.dateStart || fallbackStart || "").trim();
  const rawEnd = String(slot?.dateEnd || fallbackEnd || rawStart).trim();
  if (!rawStart && !rawEnd) return { start: "", end: "" };
  if (!rawStart) return { start: rawEnd, end: rawEnd };
  if (!rawEnd) return { start: rawStart, end: rawStart };
  return rawStart <= rawEnd
    ? { start: rawStart, end: rawEnd }
    : { start: rawEnd, end: rawStart };
}

function getEventGeneralDateStart(ev) {
  return String(ev?.eventDateStart || ev?.reservationDateStart || ev?.seriesDateStart || ev?.date || "").trim();
}

function getEventGeneralDateEnd(ev) {
  return String(ev?.eventDateEnd || ev?.reservationDateEnd || ev?.seriesDateEnd || ev?.date || "").trim();
}

function syncEventPaxFromSlots() {
  if (!el.eventPax) return 0;
  const total = getSlotsFromForm().reduce((acc, slot) => acc + Math.max(0, Number(slot?.slotPax || 0)), 0);
  el.eventPax.value = total > 0 ? String(total) : "";
  return total;
}

function syncHiddenTimesFromFirstSlot() {
  const first = getSlotsFromForm()[0];
  el.startTime.value = first?.startTime || "";
  el.endTime.value = first?.endTime || "";
}

function getQuoteFilteredCompanies(rawTerm = "") {
  const term = String(rawTerm || "").trim();
  const disabledCompanies = new Set((state.disabledCompanies || []).map((x) => String(x)));
  return (state.companies || [])
    .filter((c) => {
      if (!String(c?.id || "").trim() || disabledCompanies.has(String(c.id))) return false;
      if (!term) return true;
      const haystack = [
        c?.name,
        c?.businessName,
        c?.billTo,
        c?.owner,
        c?.email,
        c?.phone,
        c?.nit,
      ].filter(Boolean).join(" ");
      return matchesLikeSearch(haystack, term);
    })
    .slice()
    .sort((a, b) => String(a?.name || "").localeCompare(String(b?.name || ""), "es", { sensitivity: "base" }));
}

function hideQuoteCompanyResults() {
  if (!el.quoteCompanyResults) return;
  el.quoteCompanyResults.innerHTML = "";
  el.quoteCompanyResults.classList.remove("isVisible");
}

function renderQuoteCompanyResults(companies = []) {
  if (!el.quoteCompanyResults) return;
  const selectedId = String(el.quoteCompany?.value || "").trim();
  const search = String(el.quoteCompanySearch?.value || "").trim();
  if (!companies.length) {
    if (search) {
      el.quoteCompanyResults.innerHTML = '<div class="quoteCompanyEmpty">No se encontraron instituciones con ese criterio.</div>';
      el.quoteCompanyResults.classList.add("isVisible");
    } else {
      hideQuoteCompanyResults();
    }
    return;
  }
  el.quoteCompanyResults.innerHTML = companies.slice(0, 7).map((company) => {
    const companyId = String(company.id || "").trim();
    const metaPrimary = String(company.businessName || company.owner || company.email || "Sin razon social").trim();
    const metaSecondary = [company.phone, company.email, company.nit ? `NIT ${company.nit}` : ""]
      .filter(Boolean)
      .join(" | ");
    const activeClass = companyId === selectedId ? " isActive" : "";
    return `
      <button class="quoteCompanyCard${activeClass}" type="button" data-quote-company-pick="${escapeHtml(companyId)}">
        <span class="quoteCompanyCardTitle">${escapeHtml(String(company.name || "Institucion").trim())}</span>
        <span class="quoteCompanyCardMeta">${escapeHtml(metaPrimary)}</span>
        <span class="quoteCompanyCardSub">${escapeHtml(metaSecondary || "Sin telefono o correo")}</span>
      </button>
    `;
  }).join("");
  el.quoteCompanyResults.classList.add("isVisible");
}

function renderCompaniesSelect(selectedId = null) {
  if (!el.quoteCompany) return;
  const previousValue = String(el.quoteCompany.value || "").trim();
  el.quoteCompany.innerHTML = "";
  if (el.companiesList) el.companiesList.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Selecciona empresa";
  el.quoteCompany.appendChild(placeholder);
  const disabledCompanies = new Set((state.disabledCompanies || []).map((x) => String(x)));
  const keepId = String(selectedId || "").trim();
  for (const c of state.companies || []) {
    if (disabledCompanies.has(String(c.id)) && String(c.id) !== keepId) continue;
    const opt = document.createElement("option");
    opt.value = c.id;
    opt.textContent = disabledCompanies.has(String(c.id)) ? `${c.name} (Inhabilitada)` : c.name;
    el.quoteCompany.appendChild(opt);
  }
  if (selectedId) el.quoteCompany.value = selectedId;
  const finalCompanyId = String(el.quoteCompany.value || (keepId ? "" : previousValue) || "").trim();
  enhanceSelectControl(el.quoteCompany, true);
  syncEnhancedSelectValue(el.quoteCompany, finalCompanyId);
  const selectedCompany = (state.companies || []).find((c) => c.id === el.quoteCompany.value);
  if (el.quoteCompanySearch) {
    el.quoteCompanySearch.value = selectedCompany?.name || "";
  }
  hideQuoteCompanyResults();
  renderQuoteManagerSelect(el.quoteCompany.value, quoteDraft?.managerId || null);
}

function refreshCompanySuggestions(rawTerm = "") {
  if (el.companiesList) el.companiesList.innerHTML = "";
  const companies = getQuoteFilteredCompanies(rawTerm);
  renderQuoteCompanyResults(companies);
}

function renderServicesList() {
  if (!el.servicesList) return;
  el.servicesList.innerHTML = "";
  const descriptionValues = new Set();
  if (el.serviceDescriptionsList) el.serviceDescriptionsList.innerHTML = "";
  const disabledServices = new Set((state.disabledServices || []).map((x) => String(x)));
  for (const s of state.services || []) {
    if (disabledServices.has(String(s.id))) continue;
    const opt = document.createElement("option");
    opt.value = s.name;
    opt.textContent = s.description ? `${s.name} - ${s.description}` : s.name;
    el.servicesList.appendChild(opt);
    const desc = String(s.description || "").trim();
    if (el.serviceDescriptionsList && desc && !descriptionValues.has(desc.toLowerCase())) {
      descriptionValues.add(desc.toLowerCase());
      const optDesc = document.createElement("option");
      optDesc.value = desc;
      optDesc.textContent = `${s.name} - ${desc}`;
      el.serviceDescriptionsList.appendChild(optDesc);
    }
  }
}

function findCompanyMatches(rawTerm) {
  const term = String(rawTerm || "").trim();
  if (!term) return [];
  const disabledCompanies = new Set((state.disabledCompanies || []).map((x) => String(x)));
  const companies = (state.companies || []).filter((c) => !disabledCompanies.has(String(c.id)));
  return companies.filter((c) => {
    const haystack = [
      c?.name,
      c?.businessName,
      c?.billTo,
      c?.owner,
      c?.email,
      c?.phone,
      c?.nit,
    ].filter(Boolean).join(" ");
    return matchesLikeSearch(haystack, term);
  });
}

function resolveCompanyFromSearch(rawTerm) {
  const term = String(rawTerm || "").trim();
  if (!term) return null;
  const normalizedTerm = normalizeBucketKey(term);
  const disabledCompanies = new Set((state.disabledCompanies || []).map((x) => String(x)));
  const companies = (state.companies || []).filter((c) => !disabledCompanies.has(String(c.id)));
  const exactName = companies.find((c) => normalizeBucketKey(c?.name || "") === normalizedTerm);
  if (exactName) return exactName;
  const exactBusiness = companies.find((c) => normalizeBucketKey(c?.businessName || "") === normalizedTerm);
  if (exactBusiness) return exactBusiness;
  const matches = findCompanyMatches(term);
  return matches[0] || null;
}

function applyQuoteCompanyDefaults() {
  if (!quoteDraft) return;
  const company = (state.companies || []).find((c) => c.id === el.quoteCompany.value);
  const manager = company?.managers?.find((m) => m.id === el.quoteManagerSelect.value);
  quoteDraft.contact = manager?.name || company?.owner || "";
  quoteDraft.email = manager?.email || company?.email || "";
  quoteDraft.billTo = company?.billTo || company?.businessName || company?.name || "";
  quoteDraft.address = company?.address || "";
  quoteDraft.eventType = company?.eventType || "";
  quoteDraft.phone = manager?.phone || company?.phone || "";
  quoteDraft.nit = company?.nit || "";
}

function selectCompanyInQuote(companyId) {
  const company = (state.companies || []).find((c) => c.id === companyId);
  if (!company) return;
  if (el.quoteCompany) syncEnhancedSelectValue(el.quoteCompany, company.id);
  if (el.quoteCompanySearch) el.quoteCompanySearch.value = company.name || "";
  if (quoteDraft) quoteDraft.companyId = company.id;
  renderQuoteManagerSelect(company.id, null);
  if (quoteDraft) quoteDraft.managerId = String(el.quoteManagerSelect.value || "").trim();
  applyQuoteCompanyDefaults();
  fillQuoteHeaderFields(true, false);
  hideQuoteCompanyResults();
}

function findServiceMatches(rawTerm) {
  const term = String(rawTerm || "").trim();
  if (!term) return [];
  const disabledServices = new Set((state.disabledServices || []).map((x) => String(x)));
  const services = (state.services || []).filter((s) => !disabledServices.has(String(s.id)));
  return services.filter((s) => {
    const haystack = [
      s?.name,
      s?.description,
      s?.category,
      s?.subcategory,
    ].filter(Boolean).join(" ");
    return matchesLikeSearch(haystack, term);
  });
}

function resolveServiceFromSearch(rawTerm) {
  const term = String(rawTerm || "").trim();
  if (!term) return null;
  const normalizedTerm = normalizeBucketKey(term);
  const disabledServices = new Set((state.disabledServices || []).map((x) => String(x)));
  const services = (state.services || []).filter((s) => !disabledServices.has(String(s.id)));
  const exactName = services.find((s) => normalizeBucketKey(s?.name || "") === normalizedTerm);
  if (exactName) return exactName;
  const exactDesc = services.find((s) => normalizeBucketKey(s?.description || "") === normalizedTerm);
  if (exactDesc) return exactDesc;
  const matches = findServiceMatches(term);
  return matches[0] || null;
}

function applyServiceToQuoteItem(item, service) {
  if (!item || !service) return;
  const unit = Math.max(0, Number(service.price || 0));
  const mode = String(service.quantityMode || "").toUpperCase() === "PAX" ? "PAX" : "MANUAL";
  item.serviceId = service.id || null;
  item.name = service.name || item.name || "";
  item.description = service.description || item.description || item.name || "";
  item.category = service.category || "";
  item.subcategory = service.subcategory || "";
  item.quantityMode = mode;
  item.unitPrice = unit;
  if (mode === "PAX") {
    item.qty = 1;
    item.price = unit * Math.max(0, normalizeQuotePeopleValue());
    return;
  }
  if (!Number.isFinite(Number(item.qty)) || Number(item.qty) <= 0) {
    item.qty = 1;
  }
  item.price = unit;
}

function selectedOptionText(selectEl) {
  if (!selectEl) return "";
  const idx = selectEl.selectedIndex;
  if (idx < 0 || !selectEl.options[idx]) return "";
  return String(selectEl.options[idx].textContent || "").trim();
}

function sanitizeCatalogOptionLabel(rawText) {
  return String(rawText || "")
    .replace(/\s*\(Inhabilitad[ao]\)\s*$/i, "")
    .trim();
}

function renderCategoriasServicioSelect() {
  if (!el.serviceCategory) return;
  el.serviceCategory.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Seleccione categoria";
  el.serviceCategory.appendChild(placeholder);
  for (const c of catalogoCategoriasServicio) {
    const opt = document.createElement("option");
    opt.value = String(c.id);
    const isActive = c?.activo !== false;
    opt.textContent = `${c.nombre}${isActive ? "" : " (Inhabilitada)"}`;
    opt.dataset.active = isActive ? "1" : "0";
    el.serviceCategory.appendChild(opt);
  }
}

function renderSubcategoriasServicioSelect(categoriaId) {
  if (!el.serviceSubcategory) return;
  el.serviceSubcategory.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Seleccione subcategoria";
  el.serviceSubcategory.appendChild(placeholder);
  const catId = Number(categoriaId);
  const categoryActiveMap = new Map(
    catalogoCategoriasServicio.map((c) => [Number(c.id), c?.activo !== false])
  );
  const list = Number.isFinite(catId)
    ? catalogoSubcategoriasServicio.filter((s) => Number(s.id_categoria) === catId)
    : [];
  for (const s of list) {
    const opt = document.createElement("option");
    opt.value = String(s.id);
    const categoryActive = categoryActiveMap.get(Number(s.id_categoria)) !== false;
    const subActive = s?.activo !== false;
    const isActive = categoryActive && subActive;
    opt.textContent = `${s.nombre}${isActive ? "" : " (Inhabilitada)"}`;
    opt.dataset.active = isActive ? "1" : "0";
    el.serviceSubcategory.appendChild(opt);
  }
}

async function syncServiceCatalogFromDb(allowRecovery = true) {
  async function loadCatalogLists() {
    try {
      const categoriasUrl = buildApiUrlFromStateUrl(activeApiStateUrl, "categorias-servicio");
      const categoriasRes = await fetch(categoriasUrl, { cache: "no-store" });
      if (categoriasRes.ok) {
        const payloadCategorias = await categoriasRes.json();
        const categorias = Array.isArray(payloadCategorias?.categorias) ? payloadCategorias.categorias : [];
        catalogoCategoriasServicio = categorias
          .map((c) => ({ id: Number(c.id), nombre: String(c.nombre || "").trim(), activo: c?.activo !== false && Number(c?.activo) !== 0 }))
          .filter((c) => Number.isFinite(c.id) && c.nombre);
      }
    } catch (_) { }

    try {
      const subcategoriasUrl = buildApiUrlFromStateUrl(activeApiStateUrl, "subcategorias-servicio");
      const subcategoriasRes = await fetch(subcategoriasUrl, { cache: "no-store" });
      if (subcategoriasRes.ok) {
        const payloadSubcategorias = await subcategoriasRes.json();
        const subcategorias = Array.isArray(payloadSubcategorias?.subcategorias) ? payloadSubcategorias.subcategorias : [];
        catalogoSubcategoriasServicio = subcategorias
          .map((s) => ({
            id: Number(s.id),
            id_categoria: Number(s.id_categoria),
            nombre: String(s.nombre || "").trim(),
            activo: s?.activo !== false && Number(s?.activo) !== 0,
          }))
          .filter((s) => Number.isFinite(s.id) && Number.isFinite(s.id_categoria) && s.nombre);
      }
    } catch (_) { }
  }

  await loadCatalogLists();
  if (
    allowRecovery
    && catalogoCategoriasServicio.length === 0
    && Array.isArray(state.services)
    && state.services.length > 0
  ) {
    try {
      const recoverUrl = buildApiUrlFromStateUrl(activeApiStateUrl, "service-catalog/recover");
      const recoverRes = await fetch(recoverUrl, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ forceRelink: false }),
      });
      if (recoverRes.ok) {
        await loadCatalogLists();
      }
    } catch (_) { }
  }

  syncServiceLabelsFromCatalogInState();
  renderCategoriasServicioSelect();
  renderSubcategoriasServicioSelect(Number(el.serviceCategory?.value || NaN));
}

function syncServiceLabelsFromCatalogInState() {
  const services = Array.isArray(state.services) ? state.services : [];
  if (!services.length) return;
  const hasCatalogCategories = Array.isArray(catalogoCategoriasServicio) && catalogoCategoriasServicio.length > 0;
  const hasCatalogSubcategories = Array.isArray(catalogoSubcategoriasServicio) && catalogoSubcategoriasServicio.length > 0;
  if (!hasCatalogCategories && !hasCatalogSubcategories) return;

  const categoryById = new Map(
    (catalogoCategoriasServicio || []).map((c) => [Number(c.id), String(c.nombre || "").trim()])
  );
  const subcategoryById = new Map(
    (catalogoSubcategoriasServicio || []).map((s) => [
      Number(s.id),
      {
        nombre: String(s.nombre || "").trim(),
        id_categoria: Number(s.id_categoria),
      },
    ])
  );

  let changed = false;
  for (const rawService of services) {
    if (!rawService || typeof rawService !== "object") continue;
    const service = rawService;
    const categoryId = Number(service.categoryId);
    const subcategoryId = Number(service.subcategoryId);
    const subInfo = Number.isFinite(subcategoryId) ? subcategoryById.get(subcategoryId) : null;

    let nextCategoryId = Number.isFinite(categoryId) ? categoryId : null;
    if ((!Number.isFinite(categoryId) || categoryId <= 0) && subInfo && Number.isFinite(subInfo.id_categoria)) {
      nextCategoryId = Number(subInfo.id_categoria);
    }

    const hasNextCategory = Number.isFinite(nextCategoryId) && nextCategoryId > 0;
    const mappedCategoryName = hasNextCategory ? String(categoryById.get(nextCategoryId) || "") : "";
    const mappedSubcategoryName = subInfo ? String(subInfo.nombre || "") : "";
    const nextCategoryName = mappedCategoryName || String(service.category || "").trim();
    const nextSubcategoryName = mappedSubcategoryName || String(service.subcategory || "").trim();

    if (nextCategoryId !== null && Number(service.categoryId) !== Number(nextCategoryId)) {
      service.categoryId = nextCategoryId;
      changed = true;
    }
    if (String(service.category || "") !== nextCategoryName) {
      service.category = nextCategoryName;
      changed = true;
    }
    if (String(service.subcategory || "") !== nextSubcategoryName) {
      service.subcategory = nextSubcategoryName;
      changed = true;
    }
  }

  if (!changed) return;
  renderServicesList();
  try { renderServiceManagerTable(); } catch (_) { }
  persist();
}

function renderServiceCategoryManagerSelect(keep = "") {
  if (!el.serviceCategoryEditSelect) return;
  const current = String(keep || "").trim();
  el.serviceCategoryEditSelect.innerHTML = `<option value="">Crear nueva categoria</option>`;
  for (const c of catalogoCategoriasServicio.slice().sort((a, b) => String(a?.nombre || "").localeCompare(String(b?.nombre || ""), "es", { sensitivity: "base" }))) {
    const opt = document.createElement("option");
    opt.value = String(c.id);
    opt.textContent = `${c.nombre}${c?.activo === false ? " (Inhabilitada)" : ""}`;
    el.serviceCategoryEditSelect.appendChild(opt);
  }
  el.serviceCategoryEditSelect.value = current;
}

function renderServiceCategoryManagerTable() {
  if (!el.serviceCategoryBody) return;
  el.serviceCategoryBody.innerHTML = "";
  const rows = catalogoCategoriasServicio.slice().sort((a, b) => String(a?.nombre || "").localeCompare(String(b?.nombre || ""), "es", { sensitivity: "base" }));
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="3">Sin categorias registradas.</td>`;
    el.serviceCategoryBody.appendChild(tr);
    return;
  }
  for (const row of rows) {
    const isActive = row?.activo !== false;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(String(row?.nombre || ""))}</td>
      <td>${isActive ? "Activa" : "Inhabilitada"}</td>
      <td class="tableActionsCell">
        <button class="apptIconBtn apptEdit" type="button" data-service-category-edit="${escapeHtml(String(row?.id || ""))}" title="Editar" aria-label="Editar">&#9998;</button>
        <button class="apptIconBtn ${isActive ? "apptDelete" : ""}" type="button" data-service-category-toggle="${escapeHtml(String(row?.id || ""))}" title="${isActive ? "Inhabilitar" : "Reactivar"}" aria-label="${isActive ? "Inhabilitar" : "Reactivar"}">${isActive ? "&#9940;" : "&#8635;"}</button>
      </td>`;
    el.serviceCategoryBody.appendChild(tr);
  }
}

function loadServiceCategoryInManager(id = "") {
  const textId = String(id || "").trim();
  editingServiceCategoryId = textId;
  renderServiceCategoryManagerSelect(textId);
  const target = catalogoCategoriasServicio.find((c) => String(c?.id || "") === textId);
  if (el.serviceCategoryNameInput) el.serviceCategoryNameInput.value = String(target?.nombre || "");
}

async function openServiceCategoryManagerModal() {
  if (!el.serviceCategoryBackdrop) return;
  await syncServiceCatalogFromDb();
  renderServiceCategoryManagerTable();
  loadServiceCategoryInManager("");
  el.serviceCategoryBackdrop.hidden = false;
  setTimeout(() => {
    try { el.serviceCategoryNameInput?.focus(); } catch (_) { }
  }, 0);
}

function closeServiceCategoryManagerModal() {
  if (!el.serviceCategoryBackdrop) return;
  el.serviceCategoryBackdrop.hidden = true;
  editingServiceCategoryId = "";
  if (el.serviceCategoryEditSelect) el.serviceCategoryEditSelect.value = "";
  if (el.serviceCategoryNameInput) el.serviceCategoryNameInput.value = "";
}

async function saveServiceCategoryFromManagerModal() {
  const name = String(el.serviceCategoryNameInput?.value || "").trim();
  if (!name) return toast("Nombre de categoria es obligatorio.");
  const editingId = String(
    editingServiceCategoryId ||
    el.serviceCategoryEditSelect?.value ||
    ""
  ).trim();
  try {
    const endpoint = editingId
      ? buildApiUrlFromStateUrl(activeApiStateUrl, `categorias-servicio/${encodeURIComponent(editingId)}`)
      : buildApiUrlFromStateUrl(activeApiStateUrl, "categorias-servicio");
    const res = await fetch(endpoint, {
      method: editingId ? "PUT" : "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ nombre: name }),
    });
    const payload = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(payload?.detail || payload?.message || "No se pudo guardar la categoria.");
    const selectedId = String(payload?.categoria?.id || editingId || "").trim();
    await syncServiceCatalogFromDb();
    if (el.serviceCategory) el.serviceCategory.value = selectedId || el.serviceCategory.value || "";
    renderSubcategoriasServicioSelect(Number(el.serviceCategory?.value || NaN));
    renderServiceCategoryManagerTable();
    loadServiceCategoryInManager("");
    toast(editingId ? "Categoria actualizada." : "Categoria creada.");
  } catch (error) {
    toast(error?.message || "No se pudo guardar la categoria.");
  }
}

function renderServiceSubcategoryCategorySelect(keep = "") {
  if (!el.serviceSubcategoryCategorySelect) return;
  const selected = String(keep || "").trim();
  el.serviceSubcategoryCategorySelect.innerHTML = `<option value="">Selecciona categoria</option>`;
  for (const c of catalogoCategoriasServicio.slice().sort((a, b) => String(a?.nombre || "").localeCompare(String(b?.nombre || ""), "es", { sensitivity: "base" }))) {
    const opt = document.createElement("option");
    opt.value = String(c.id);
    opt.textContent = `${c.nombre}${c?.activo === false ? " (Inhabilitada)" : ""}`;
    el.serviceSubcategoryCategorySelect.appendChild(opt);
  }
  el.serviceSubcategoryCategorySelect.value = selected;
}

function renderServiceSubcategoryManagerSelect(keep = "", categoryId = "") {
  if (!el.serviceSubcategoryEditSelect) return;
  const current = String(keep || "").trim();
  const category = Number(categoryId);
  el.serviceSubcategoryEditSelect.innerHTML = `<option value="">Crear nueva subcategoria</option>`;
  const rows = catalogoSubcategoriasServicio
    .filter((s) => !Number.isFinite(category) || category <= 0 || Number(s?.id_categoria) === category)
    .sort((a, b) => String(a?.nombre || "").localeCompare(String(b?.nombre || ""), "es", { sensitivity: "base" }));
  for (const s of rows) {
    const opt = document.createElement("option");
    opt.value = String(s.id);
    opt.textContent = `${s.nombre}${s?.activo === false ? " (Inhabilitada)" : ""}`;
    el.serviceSubcategoryEditSelect.appendChild(opt);
  }
  el.serviceSubcategoryEditSelect.value = current;
}

function renderServiceSubcategoryManagerTable(categoryFilter = "") {
  if (!el.serviceSubcategoryBody) return;
  const filterId = Number(categoryFilter);
  const categoryById = new Map(catalogoCategoriasServicio.map((c) => [Number(c.id), String(c.nombre || "")]));
  const rows = catalogoSubcategoriasServicio
    .filter((s) => !Number.isFinite(filterId) || filterId <= 0 || Number(s?.id_categoria) === filterId)
    .slice()
    .sort((a, b) => String(a?.nombre || "").localeCompare(String(b?.nombre || ""), "es", { sensitivity: "base" }));
  el.serviceSubcategoryBody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="4">Sin subcategorias registradas.</td>`;
    el.serviceSubcategoryBody.appendChild(tr);
    return;
  }
  for (const row of rows) {
    const isActive = row?.activo !== false;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(String(row?.nombre || ""))}</td>
      <td>${escapeHtml(String(categoryById.get(Number(row?.id_categoria)) || "-"))}</td>
      <td>${isActive ? "Activa" : "Inhabilitada"}</td>
      <td class="tableActionsCell">
        <button class="apptIconBtn apptEdit" type="button" data-service-subcategory-edit="${escapeHtml(String(row?.id || ""))}" title="Editar" aria-label="Editar">&#9998;</button>
        <button class="apptIconBtn ${isActive ? "apptDelete" : ""}" type="button" data-service-subcategory-toggle="${escapeHtml(String(row?.id || ""))}" title="${isActive ? "Inhabilitar" : "Reactivar"}" aria-label="${isActive ? "Inhabilitar" : "Reactivar"}">${isActive ? "&#9940;" : "&#8635;"}</button>
      </td>`;
    el.serviceSubcategoryBody.appendChild(tr);
  }
}

function loadServiceSubcategoryInManager(id = "", defaultCategoryId = "") {
  const textId = String(id || "").trim();
  editingServiceSubcategoryId = textId;
  const target = catalogoSubcategoriasServicio.find((s) => String(s?.id || "") === textId);
  const categoryId = String(target?.id_categoria || defaultCategoryId || "").trim();
  renderServiceSubcategoryCategorySelect(categoryId);
  renderServiceSubcategoryManagerSelect(textId, categoryId);
  renderServiceSubcategoryManagerTable(categoryId);
  if (el.serviceSubcategoryNameInput) el.serviceSubcategoryNameInput.value = String(target?.nombre || "");
}

async function openServiceSubcategoryManagerModal() {
  if (!el.serviceSubcategoryBackdrop) return;
  await syncServiceCatalogFromDb();
  const categoryFromService = String(el.serviceCategory?.value || "").trim();
  loadServiceSubcategoryInManager("", categoryFromService);
  el.serviceSubcategoryBackdrop.hidden = false;
  setTimeout(() => {
    try { el.serviceSubcategoryNameInput?.focus(); } catch (_) { }
  }, 0);
}

function closeServiceSubcategoryManagerModal() {
  if (!el.serviceSubcategoryBackdrop) return;
  el.serviceSubcategoryBackdrop.hidden = true;
  editingServiceSubcategoryId = "";
  if (el.serviceSubcategoryCategorySelect) el.serviceSubcategoryCategorySelect.value = "";
  if (el.serviceSubcategoryEditSelect) el.serviceSubcategoryEditSelect.value = "";
  if (el.serviceSubcategoryNameInput) el.serviceSubcategoryNameInput.value = "";
}

async function saveServiceSubcategoryFromManagerModal() {
  const categoryId = Number(el.serviceSubcategoryCategorySelect?.value || NaN);
  const name = String(el.serviceSubcategoryNameInput?.value || "").trim();
  if (!Number.isFinite(categoryId) || categoryId <= 0) return toast("Categoria es obligatoria.");
  if (!name) return toast("Nombre de subcategoria es obligatorio.");
  const editingId = String(
    editingServiceSubcategoryId ||
    el.serviceSubcategoryEditSelect?.value ||
    ""
  ).trim();
  try {
    const endpoint = editingId
      ? buildApiUrlFromStateUrl(activeApiStateUrl, `subcategorias-servicio/${encodeURIComponent(editingId)}`)
      : buildApiUrlFromStateUrl(activeApiStateUrl, "subcategorias-servicio");
    const res = await fetch(endpoint, {
      method: editingId ? "PUT" : "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ id_categoria: categoryId, nombre: name }),
    });
    const payload = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(payload?.detail || payload?.message || "No se pudo guardar la subcategoria.");
    const subcategoryId = String(payload?.subcategoria?.id || editingId || "").trim();
    await syncServiceCatalogFromDb();
    if (el.serviceCategory) el.serviceCategory.value = String(categoryId);
    renderSubcategoriasServicioSelect(categoryId);
    if (el.serviceSubcategory) el.serviceSubcategory.value = subcategoryId;
    loadServiceSubcategoryInManager("", String(categoryId));
    toast(editingId ? "Subcategoria actualizada." : "Subcategoria creada.");
  } catch (error) {
    toast(error?.message || "No se pudo guardar la subcategoria.");
  }
}

async function toggleServiceCategoryActiveFromManager(categoryId) {
  const id = String(categoryId || "").trim();
  if (!id) return;
  const current = catalogoCategoriasServicio.find((c) => String(c?.id || "") === id);
  if (!current) return;
  const nextActive = current?.activo === false ? true : false;
  try {
    const endpoint = buildApiUrlFromStateUrl(activeApiStateUrl, `categorias-servicio/${encodeURIComponent(id)}/activo`);
    const res = await fetch(endpoint, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ activo: nextActive ? 1 : 0 }),
    });
    const payload = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(payload?.detail || payload?.message || "No se pudo cambiar el estado de categoria.");
    await syncServiceCatalogFromDb();
    renderServiceCategoryManagerTable();
    loadServiceCategoryInManager(editingServiceCategoryId || "");
    toast(nextActive ? "Categoria reactivada." : "Categoria inhabilitada.");
  } catch (error) {
    toast(error?.message || "No se pudo cambiar el estado de categoria.");
  }
}

async function toggleServiceSubcategoryActiveFromManager(subcategoryId) {
  const id = String(subcategoryId || "").trim();
  if (!id) return;
  const current = catalogoSubcategoriasServicio.find((s) => String(s?.id || "") === id);
  if (!current) return;
  const nextActive = current?.activo === false ? true : false;
  try {
    const endpoint = buildApiUrlFromStateUrl(activeApiStateUrl, `subcategorias-servicio/${encodeURIComponent(id)}/activo`);
    const res = await fetch(endpoint, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ activo: nextActive ? 1 : 0 }),
    });
    const payload = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(payload?.detail || payload?.message || "No se pudo cambiar el estado de subcategoria.");
    const selectedCategoryId = String(el.serviceSubcategoryCategorySelect?.value || "").trim();
    await syncServiceCatalogFromDb();
    loadServiceSubcategoryInManager("", selectedCategoryId);
    toast(nextActive ? "Subcategoria reactivada." : "Subcategoria inhabilitada.");
  } catch (error) {
    toast(error?.message || "No se pudo cambiar el estado de subcategoria.");
  }
}

function renderQuoteManagerSelect(companyId, selectedManagerId = null) {
  if (!el.quoteManagerSelect) return;
  const previousValue = String(el.quoteManagerSelect.value || "").trim();
  el.quoteManagerSelect.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Selecciona encargado";
  el.quoteManagerSelect.appendChild(placeholder);
  const company = (state.companies || []).find(c => c.id === companyId);
  const disabledManagers = new Set((state.disabledManagers || []).map((x) => String(x)));
  const keepId = String(selectedManagerId || "").trim();
  const managers = (company?.managers || []).filter((m) => !disabledManagers.has(String(m.id)) || String(m.id) === keepId);
  for (const m of managers) {
    const opt = document.createElement("option");
    opt.value = m.id;
    opt.textContent = `${m.name}${disabledManagers.has(String(m.id)) ? " (Inhabilitado)" : ""}`;
    el.quoteManagerSelect.appendChild(opt);
  }
  if (selectedManagerId && managers.some(m => m.id === selectedManagerId)) {
    el.quoteManagerSelect.value = selectedManagerId;
  }
  const managerId = String(el.quoteManagerSelect.value || selectedManagerId || previousValue || "").trim();
  enhanceSelectControl(el.quoteManagerSelect, true);
  syncEnhancedSelectValue(el.quoteManagerSelect, managerId);
}

function normalizeQuoteItemForSnapshot(rawItem) {
  const qty = Math.max(0, Number(rawItem?.qty || 0));
  const price = Math.max(0, Number(rawItem?.price || 0));
  const unitPrice = Math.max(0, Number(rawItem?.unitPrice || price || 0));
  const rowId = String(rawItem?.rowId || uid());
  return {
    rowId,
    serviceId: rawItem?.serviceId ? String(rawItem.serviceId) : null,
    serviceDate: String(rawItem?.serviceDate || "").trim(),
    name: String(rawItem?.name || "").trim(),
    description: String(rawItem?.description || "").trim(),
    category: String(rawItem?.category || "").trim(),
    subcategory: String(rawItem?.subcategory || "").trim(),
    quantityMode: String(rawItem?.quantityMode || "MANUAL").toUpperCase() === "PAX" ? "PAX" : "MANUAL",
    qty: Number.isFinite(qty) ? qty : 0,
    unitPrice: Number.isFinite(unitPrice) ? unitPrice : 0,
    price: Number.isFinite(price) ? price : 0,
  };
}

function normalizeQuoteItemsForSnapshot(rawItems) {
  const list = Array.isArray(rawItems) ? rawItems : [];
  return list.map((item) => normalizeQuoteItemForSnapshot(item));
}

function buildQuoteVersionComparable(quoteLike) {
  const q = quoteLike || {};
  const items = normalizeQuoteItemsForSnapshot(q.items).map((item) => ({
    serviceDate: item.serviceDate,
    name: item.name,
    description: item.description,
    qty: Number(item.qty || 0),
    price: Number(item.price || 0),
    category: item.category,
    subcategory: item.subcategory,
    quantityMode: item.quantityMode,
  })).sort((a, b) => {
    const d = String(a.serviceDate || "").localeCompare(String(b.serviceDate || ""));
    if (d !== 0) return d;
    const n = String(a.name || "").localeCompare(String(b.name || ""));
    if (n !== 0) return n;
    const x = String(a.description || "").localeCompare(String(b.description || ""));
    if (x !== 0) return x;
    if (Number(a.qty || 0) !== Number(b.qty || 0)) return Number(a.qty || 0) - Number(b.qty || 0);
    return Number(a.price || 0) - Number(b.price || 0);
  });
  return {
    companyId: String(q.companyId || "").trim(),
    managerId: String(q.managerId || "").trim(),
    contact: String(q.contact || "").trim(),
    email: String(q.email || "").trim(),
    billTo: String(q.billTo || "").trim(),
    address: String(q.address || "").trim(),
    eventType: String(q.eventType || "").trim(),
    venue: String(q.venue || "").trim(),
    schedule: String(q.schedule || "").trim(),
    code: String(q.code || "").trim(),
    docDate: String(q.docDate || "").trim(),
    phone: String(q.phone || "").trim(),
    nit: String(q.nit || "").trim(),
    people: String(q.people || "").trim(),
    eventDate: String(q.eventDate || "").trim(),
    folio: String(q.folio || "").trim(),
    endDate: String(q.endDate || "").trim(),
    dueDate: String(q.dueDate || "").trim(),
    paymentType: quotePaymentTypesToStorage(q.paymentType),
    discountType: normalizeDiscountType(q.discountType),
    discountValue: Math.max(0, Number(q.discountValue || 0)),
    internalNotes: String(q.internalNotes || q.notes || "").trim(),
    templateId: String(q.templateId || "").trim(),
    currency: normalizeQuoteCurrency(q.currency),
    items,
  };
}

function areQuotesEquivalentForVersioning(a, b) {
  const aa = buildQuoteVersionComparable(a);
  const bb = buildQuoteVersionComparable(b);
  return JSON.stringify(aa) === JSON.stringify(bb);
}

function buildCompanyComparable(companyLike) {
  const c = normalizeCompanyRecord(companyLike || {});
  return {
    name: String(c.name || "").trim(),
    owner: String(c.owner || "").trim(),
    email: String(c.email || "").trim(),
    nit: String(c.nit || "").trim(),
    businessName: String(c.businessName || "").trim(),
    billTo: String(c.billTo || "").trim(),
    eventType: String(c.eventType || "").trim(),
    address: String(c.address || "").trim(),
    phone: String(c.phone || "").trim(),
    notes: String(c.notes || "").trim(),
    managers: (Array.isArray(c.managers) ? c.managers : []).map((m) => ({
      name: String(m.name || "").trim(),
      phone: String(m.phone || "").trim(),
      email: String(m.email || "").trim(),
      address: String(m.address || "").trim(),
    })),
  };
}

function areCompaniesEquivalent(a, b) {
  return JSON.stringify(buildCompanyComparable(a)) === JSON.stringify(buildCompanyComparable(b));
}

function buildServiceComparable(serviceLike) {
  const s = normalizeServiceRecord(serviceLike || {});
  return {
    name: String(s.name || "").trim(),
    price: Math.max(0, Number(s.price || 0)),
    description: String(s.description || "").trim(),
    categoryId: Number(s.categoryId || 0) || 0,
    subcategoryId: Number(s.subcategoryId || 0) || 0,
    category: String(s.category || "").trim(),
    subcategory: String(s.subcategory || "").trim(),
    quantityMode: String(s.quantityMode || "").trim().toUpperCase() === "PAX" ? "PAX" : "MANUAL",
  };
}

function areServicesEquivalent(a, b) {
  return JSON.stringify(buildServiceComparable(a)) === JSON.stringify(buildServiceComparable(b));
}

function buildUserComparable(userLike) {
  const u = normalizeUserRecord(userLike || {});
  return {
    name: String(u.fullName || u.name || "").trim(),
    role: normalizeUserRole(u.role),
    username: String(u.username || "").trim(),
    email: String(u.email || "").trim(),
    phone: String(u.phone || "").trim(),
    password: String(u.password || "").trim(),
    signatureDataUrl: String(u.signatureDataUrl || "").trim(),
    avatarDataUrl: String(u.avatarDataUrl || "").trim(),
    active: u.active !== false,
    salesTargetEnabled: u.salesTargetEnabled === true,
    monthlyGoals: Array.isArray(u.monthlyGoals) ? u.monthlyGoals.map((g) => ({
      month: String(g.month || "").trim(),
      amount: Math.max(0, Number(g.amount || 0)),
    })) : [],
  };
}

function areUsersEquivalent(a, b) {
  return JSON.stringify(buildUserComparable(a)) === JSON.stringify(buildUserComparable(b));
}

function cloneQuoteSnapshot(rawQuote, forcedVersion = null) {
  const base = deepClone(rawQuote || {});
  delete base.versions;
  const versionRaw = forcedVersion ?? base.version;
  const parsedVersion = Number(versionRaw);
  base.version = Number.isFinite(parsedVersion) && parsedVersion > 0 ? Math.floor(parsedVersion) : 1;
  base.items = normalizeQuoteItemsForSnapshot(base.items);
  base.discountType = normalizeDiscountType(base.discountType);
  base.discountValue = Math.max(0, Number(base.discountValue || 0));
  base.currency = normalizeQuoteCurrency(base.currency);
  return base;
}

function cloneQuoteVersionSnapshot(rawQuote, forcedVersion = null) {
  const snap = cloneQuoteSnapshot(rawQuote, forcedVersion);
  delete snap.menuMontajeEntries;
  delete snap.menuMontajeVersions;
  delete snap.menuMontajeVersion;
  return snap;
}

function normalizeQuoteVersionHistory(rawVersions) {
  const versions = Array.isArray(rawVersions) ? rawVersions : [];
  return versions
    .map((v, idx) => cloneQuoteVersionSnapshot(v, Number(v?.version) || (idx + 1)))
    .sort((a, b) => Number(a.version || 0) - Number(b.version || 0));
}

function getLatestQuoteSnapshotForEvent(ev) {
  const q = ev?.quote;
  if (!q || typeof q !== "object") return null;
  const versions = normalizeQuoteVersionHistory(q.versions);
  const current = cloneQuoteSnapshot(q, Number(q.version) || (versions.length + 1));
  const candidates = [...versions, current];
  if (!candidates.length) return null;
  candidates.sort((a, b) => {
    const verDiff = Number(b.version || 0) - Number(a.version || 0);
    if (verDiff !== 0) return verDiff;
    const ta = new Date(a.quotedAt || 0).getTime() || 0;
    const tb = new Date(b.quotedAt || 0).getTime() || 0;
    return tb - ta;
  });
  return candidates[0];
}

function getEventInstitutionName(ev) {
  const latestQuote = getLatestQuoteSnapshotForEvent(ev);
  const companyName = String(latestQuote?.companyName || "").trim();
  if (companyName) return companyName;
  const companyId = String(latestQuote?.companyId || ev?.quote?.companyId || "").trim();
  if (!companyId) return "";
  return String((state.companies || []).find((c) => c.id === companyId)?.name || "").trim();
}

function getEventLatestQuoteTotalLabel(ev) {
  const latestQuote = getLatestQuoteSnapshotForEvent(ev);
  if (!latestQuote) return "";
  const totals = getQuoteTotals(latestQuote);
  return `Cot Q ${totals.total.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

function renderQuoteVersionControls() {
  if (!el.quoteVersionSelect || !quoteDraft) return;
  const versions = normalizeQuoteVersionHistory(quoteDraft.versions);
  const currentVersion = Number(quoteDraft.version || (versions.length + 1));
  const formatVersionDateTime = (isoText) => {
    const raw = String(isoText || "").trim();
    if (!raw) return "sin fecha";
    const d = new Date(raw);
    if (Number.isNaN(d.getTime())) return raw;
    const yyyy = d.getFullYear();
    const mm = String(d.getMonth() + 1).padStart(2, "0");
    const dd = String(d.getDate()).padStart(2, "0");
    const hh = String(d.getHours()).padStart(2, "0");
    const mi = String(d.getMinutes()).padStart(2, "0");
    return `${yyyy}-${mm}-${dd} ${hh}:${mi}`;
  };
  const formatVersionTotal = (quoteLike) => {
    const totals = getQuoteTotals(quoteLike);
    return `Q ${totals.total.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
  };
  el.quoteVersionSelect.innerHTML = "";

  const currentOpt = document.createElement("option");
  currentOpt.value = String(currentVersion);
  currentOpt.textContent = `V${currentVersion} (actual) - ${formatVersionDateTime(quoteDraft.quotedAt)} - ${formatVersionTotal(quoteDraft)}`;
  el.quoteVersionSelect.appendChild(currentOpt);

  for (let i = versions.length - 1; i >= 0; i--) {
    const v = versions[i];
    const opt = document.createElement("option");
    opt.value = String(v.version);
    opt.textContent = `V${v.version} - ${formatVersionDateTime(v.quotedAt)} - ${formatVersionTotal(v)}`;
    el.quoteVersionSelect.appendChild(opt);
  }
  el.quoteVersionSelect.value = String(currentVersion);
}

function renderQuoteTemplateSelect(selectedId = "") {
  if (!el.quoteTemplateSelect) return;
  el.quoteTemplateSelect.innerHTML = "";
  const none = document.createElement("option");
  none.value = "";
  none.textContent = "Sin plantilla";
  el.quoteTemplateSelect.appendChild(none);

  const ordered = (quickTemplates || [])
    .filter((t) => String(t?.id || "").trim() !== CORPORATE_TEMPLATE_ID)
    .slice()
    .sort((a, b) => String(a?.name || "").localeCompare(String(b?.name || ""), "es", { sensitivity: "base" }));
  for (const tpl of ordered) {
    const opt = document.createElement("option");
    opt.value = String(tpl.id || "");
    opt.textContent = String(tpl.name || "Plantilla");
    el.quoteTemplateSelect.appendChild(opt);
  }
  const fallbackId = ordered.some((x) => String(x.id) === CONTRACT_CORP_TEMPLATE_ID)
    ? CONTRACT_CORP_TEMPLATE_ID
    : (ordered[0]?.id || "");
  const selected = selectedId && ordered.some((x) => x.id === selectedId)
    ? selectedId
    : fallbackId;
  el.quoteTemplateSelect.value = selected || "";
}

function getQuoteServiceTemplatesFromState() {
  return Array.isArray(state?.quoteServiceTemplates) ? state.quoteServiceTemplates : [];
}

function normalizeQuoteServiceTemplateItems(rawItems) {
  const rangeDates = getQuoteRangeDates();
  const fallbackDate = String(el.quoteServiceDate?.value || rangeDates[0] || "").trim();
  return normalizeQuoteItemsForSnapshot(rawItems)
    .map((item) => {
      const next = { ...item, rowId: uid() };
      const candidateDate = String(item?.serviceDate || "").trim();
      next.serviceDate = rangeDates.length
        ? (rangeDates.includes(candidateDate) ? candidateDate : fallbackDate)
        : (candidateDate || fallbackDate);
      if (String(next.quantityMode || "").toUpperCase() !== "PAX") {
        next.qty = Math.max(1, Number(next.qty || 1));
      }
      return next;
    })
    .filter((item) => String(item.name || "").trim());
}

function normalizeQuoteServiceTemplateRecord(candidate, index = 0) {
  const row = candidate && typeof candidate === "object" ? candidate : {};
  const id = String(row.id || `qst_${index + 1}`).trim() || `qst_${index + 1}`;
  return {
    id,
    name: String(row.name || "").trim() || `Plantilla ${index + 1}`,
    items: normalizeQuoteItemsForSnapshot(row.items),
    createdAt: String(row.createdAt || "").trim() || new Date().toISOString(),
    updatedAt: String(row.updatedAt || "").trim() || new Date().toISOString(),
  };
}

function normalizeQuoteServiceTemplatesCollection(rawTemplates) {
  const list = Array.isArray(rawTemplates) ? rawTemplates : [];
  const normalized = [];
  const seenIds = new Set();
  for (let i = 0; i < list.length; i++) {
    const rec = normalizeQuoteServiceTemplateRecord(list[i], i);
    if (!rec.id || seenIds.has(rec.id)) continue;
    seenIds.add(rec.id);
    normalized.push(rec);
  }
  normalized.sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "es", { sensitivity: "base" }));
  return normalized;
}

function renderQuoteServiceTemplateSelect(selectedId = "") {
  if (!el.quoteServiceTemplateSelect) return;
  const templates = normalizeQuoteServiceTemplatesCollection(getQuoteServiceTemplatesFromState());
  state.quoteServiceTemplates = templates;
  el.quoteServiceTemplateSelect.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = templates.length ? "Selecciona plantilla" : "No hay plantillas guardadas";
  el.quoteServiceTemplateSelect.appendChild(placeholder);
  for (const tpl of templates) {
    const opt = document.createElement("option");
    opt.value = String(tpl.id || "");
    opt.textContent = String(tpl.name || "Plantilla");
    el.quoteServiceTemplateSelect.appendChild(opt);
  }
  const selected = templates.some((t) => String(t.id) === String(selectedId || "")) ? String(selectedId || "") : "";
  el.quoteServiceTemplateSelect.value = selected;
  if (el.quoteServiceTemplateName) {
    const found = templates.find((t) => String(t.id) === selected);
    el.quoteServiceTemplateName.value = found ? String(found.name || "") : "";
  }
}

function renderSettingsQuoteTemplateServicesList() {
  if (!el.settingsQuoteTemplateServicesList) return;
  el.settingsQuoteTemplateServicesList.innerHTML = "";
  const disabledServices = new Set((state.disabledServices || []).map((x) => String(x)));
  const services = (state.services || [])
    .filter((s) => !disabledServices.has(String(s?.id || "")))
    .slice()
    .sort((a, b) => String(a?.name || "").localeCompare(String(b?.name || ""), "es", { sensitivity: "base" }));
  for (const service of services) {
    const opt = document.createElement("option");
    opt.value = String(service.name || "");
    const meta = [service.category, service.subcategory].filter(Boolean).join(" / ");
    if (meta) opt.label = meta;
    el.settingsQuoteTemplateServicesList.appendChild(opt);
  }
}

function renderSettingsQuoteServiceTemplateSelect(selectedId = "") {
  if (!el.settingsQuoteServiceTemplateSelect) return;
  const templates = normalizeQuoteServiceTemplatesCollection(getQuoteServiceTemplatesFromState());
  state.quoteServiceTemplates = templates;
  el.settingsQuoteServiceTemplateSelect.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = templates.length ? "Selecciona plantilla" : "No hay plantillas guardadas";
  el.settingsQuoteServiceTemplateSelect.appendChild(placeholder);
  for (const tpl of templates) {
    const opt = document.createElement("option");
    opt.value = String(tpl.id || "");
    opt.textContent = String(tpl.name || "Plantilla");
    el.settingsQuoteServiceTemplateSelect.appendChild(opt);
  }
  const selected = templates.some((t) => String(t.id) === String(selectedId || "")) ? String(selectedId || "") : "";
  el.settingsQuoteServiceTemplateSelect.value = selected;
}

function renderQuoteServiceTemplateManagerItems() {
  if (!el.settingsQuoteTemplateItemsBody) return;
  const items = normalizeQuoteItemsForSnapshot(quoteServiceTemplateManagerItems);
  quoteServiceTemplateManagerItems = items;
  if (!items.length) {
    el.settingsQuoteTemplateItemsBody.innerHTML = `<tr><td colspan="5">Sin servicios en esta plantilla.</td></tr>`;
    return;
  }
  el.settingsQuoteTemplateItemsBody.innerHTML = items.map((item) => {
    const qty = Number(item.qty || 0);
    const unit = Math.max(0, Number(item.unitPrice || item.price || 0));
    const total = String(item.quantityMode || "").toUpperCase() === "PAX" ? unit : unit * Math.max(1, qty || 1);
    return `
      <tr>
        <td><strong>${escapeHtml(item.name || "Servicio")}</strong><small>${escapeHtml([item.category, item.subcategory].filter(Boolean).join(" / "))}</small></td>
        <td>${escapeHtml(String(item.quantityMode || "").toUpperCase() === "PAX" ? "PAX" : String(qty || 1))}</td>
        <td>Q ${unit.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        <td>Q ${total.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
        <td><button class="btn" type="button" data-qstm-remove="${escapeHtml(String(item.rowId || ""))}">Quitar</button></td>
      </tr>`;
  }).join("");
}

function loadQuoteServiceTemplateManager(templateId = "") {
  const id = String(templateId || "").trim();
  const templates = normalizeQuoteServiceTemplatesCollection(getQuoteServiceTemplatesFromState());
  const found = templates.find((t) => String(t.id) === id);
  renderSettingsQuoteServiceTemplateSelect(id);
  if (el.settingsQuoteServiceTemplateName) el.settingsQuoteServiceTemplateName.value = found ? String(found.name || "") : "";
  quoteServiceTemplateManagerItems = found ? normalizeQuoteItemsForSnapshot(found.items).map((item) => ({ ...item, rowId: uid() })) : [];
  renderQuoteServiceTemplateManagerItems();
}

function openQuoteServiceTemplateManager() {
  if (!el.quoteServiceTemplateManagerBackdrop) return;
  renderSettingsQuoteTemplateServicesList();
  loadQuoteServiceTemplateManager(el.settingsQuoteServiceTemplateSelect?.value || "");
  el.quoteServiceTemplateManagerBackdrop.hidden = false;
  setTimeout(() => el.settingsQuoteServiceTemplateSelect?.focus(), 0);
}

function closeQuoteServiceTemplateManager() {
  if (el.quoteServiceTemplateManagerBackdrop) el.quoteServiceTemplateManagerBackdrop.hidden = true;
}

function addServiceToQuoteServiceTemplateManager() {
  const rawName = String(el.settingsQuoteTemplateServiceSearch?.value || "").trim();
  if (!rawName) return toast("Escribe un servicio.");
  const service = resolveServiceFromSearch(rawName);
  const requestedQty = Math.max(0, Math.floor(Number(el.settingsQuoteTemplateServiceQty?.value || 0)));
  const isPaxService = String(service?.quantityMode || "").trim().toUpperCase() === "PAX";
  if (!isPaxService && requestedQty <= 0) return toast("Este producto requiere una cantidad manual.");
  const unitPrice = Math.max(0, Number(service?.price || 0));
  quoteServiceTemplateManagerItems.push({
    rowId: uid(),
    serviceId: service?.id || null,
    name: service?.name || rawName,
    qty: isPaxService ? 1 : requestedQty,
    unitPrice,
    price: unitPrice,
    description: service?.description || "",
    category: service?.category || "",
    subcategory: service?.subcategory || "",
    quantityMode: isPaxService ? "PAX" : "MANUAL",
    serviceDate: "",
  });
  if (el.settingsQuoteTemplateServiceSearch) el.settingsQuoteTemplateServiceSearch.value = "";
  if (el.settingsQuoteTemplateServiceQty) el.settingsQuoteTemplateServiceQty.value = "";
  renderQuoteServiceTemplateManagerItems();
}

function saveQuoteServiceTemplateFromManager(updateExisting = false) {
  const name = String(el.settingsQuoteServiceTemplateName?.value || "").trim();
  if (!name) return toast("Escribe el nombre de la plantilla.");
  const sanitizedItems = normalizeQuoteItemsForSnapshot(quoteServiceTemplateManagerItems).filter((item) => String(item.name || "").trim());
  if (!sanitizedItems.length) return toast("Agrega al menos un servicio a la plantilla.");
  const templates = normalizeQuoteServiceTemplatesCollection(getQuoteServiceTemplatesFromState());
  const nowIso = new Date().toISOString();
  if (updateExisting) {
    const selectedId = String(el.settingsQuoteServiceTemplateSelect?.value || "").trim();
    if (!selectedId) return toast("Selecciona una plantilla para actualizar.");
    const idx = templates.findIndex((t) => String(t.id) === selectedId);
    if (idx < 0) return toast("Plantilla no encontrada para actualizar.");
    templates[idx] = { ...templates[idx], name, items: sanitizedItems, updatedAt: nowIso };
    state.quoteServiceTemplates = normalizeQuoteServiceTemplatesCollection(templates);
    renderQuoteServiceTemplateSelect(selectedId);
    loadQuoteServiceTemplateManager(selectedId);
    persist();
    toast("Plantilla de servicios actualizada.");
    return;
  }
  const duplicate = templates.find((t) => String(t.name || "").trim().toLowerCase() === name.toLowerCase());
  if (duplicate) return toast("Ya existe una plantilla con ese nombre. Seleccionala para actualizar.");
  const nextId = `qst_${uid()}`;
  templates.push({ id: nextId, name, items: sanitizedItems, createdAt: nowIso, updatedAt: nowIso });
  state.quoteServiceTemplates = normalizeQuoteServiceTemplatesCollection(templates);
  renderQuoteServiceTemplateSelect(nextId);
  loadQuoteServiceTemplateManager(nextId);
  persist();
  toast("Plantilla de servicios guardada.");
}

async function deleteQuoteServiceTemplateFromManager() {
  const selectedId = String(el.settingsQuoteServiceTemplateSelect?.value || "").trim();
  if (!selectedId) return toast("Selecciona una plantilla para eliminar.");
  const templates = normalizeQuoteServiceTemplatesCollection(getQuoteServiceTemplatesFromState());
  const target = templates.find((t) => String(t.id) === selectedId);
  if (!target) return toast("Plantilla no encontrada.");
  const ok = await modernConfirm({
    title: "Eliminar plantilla",
    message: `Esta seguro de eliminar la plantilla "${target.name}"?`,
    confirmText: "Si, eliminar",
    cancelText: "No",
  });
  if (!ok) return;
  state.quoteServiceTemplates = templates.filter((t) => String(t.id) !== selectedId);
  renderQuoteServiceTemplateSelect("");
  loadQuoteServiceTemplateManager("");
  persist();
  toast("Plantilla eliminada.");
}

async function applyQuoteServiceTemplateToDraft(templateId) {
  if (!quoteDraft) return;
  const id = String(templateId || el.quoteServiceTemplateSelect?.value || "").trim();
  if (!id) return toast("Selecciona una plantilla de servicios.");
  const templates = normalizeQuoteServiceTemplatesCollection(getQuoteServiceTemplatesFromState());
  state.quoteServiceTemplates = templates;
  const target = templates.find((t) => String(t.id) === id);
  if (!target) return toast("Plantilla no encontrada.");
  if (Array.isArray(quoteDraft.items) && quoteDraft.items.length) {
    const ok = await modernConfirm({
      title: "Aplicar plantilla",
      message: "Esto reemplazara los servicios actuales de la cotizacion. Deseas continuar?",
      confirmText: "Si, aplicar",
      cancelText: "No",
    });
    if (!ok) return;
  }
  quoteDraft.items = normalizeQuoteServiceTemplateItems(target.items);
  if (el.quoteServiceTemplateName) el.quoteServiceTemplateName.value = String(target.name || "");
  syncPaxQuantityItems();
  renderQuoteItems();
  toast(`Plantilla aplicada: ${target.name}`);
}

function saveQuoteServiceTemplateFromDraft(updateExisting = false) {
  if (!quoteDraft) return;
  const name = String(el.quoteServiceTemplateName?.value || "").trim();
  if (!name) return toast("Escribe el nombre de la plantilla.");
  if (!Array.isArray(quoteDraft.items) || !quoteDraft.items.length) {
    return toast("Agrega al menos un servicio para guardar la plantilla.");
  }
  const templates = normalizeQuoteServiceTemplatesCollection(getQuoteServiceTemplatesFromState());
  const nowIso = new Date().toISOString();
  const sanitizedItems = normalizeQuoteServiceTemplateItems(quoteDraft.items);
  if (!sanitizedItems.length) return toast("La plantilla no tiene servicios validos para guardar.");

  if (updateExisting) {
    const selectedId = String(el.quoteServiceTemplateSelect?.value || "").trim();
    if (!selectedId) return toast("Selecciona una plantilla para actualizar.");
    const idx = templates.findIndex((t) => String(t.id) === selectedId);
    if (idx < 0) return toast("Plantilla no encontrada para actualizar.");
    templates[idx] = { ...templates[idx], name, items: sanitizedItems, updatedAt: nowIso };
    state.quoteServiceTemplates = normalizeQuoteServiceTemplatesCollection(templates);
    renderQuoteServiceTemplateSelect(selectedId);
    persist();
    toast("Plantilla de servicios actualizada.");
    return;
  }

  const duplicate = templates.find((t) => String(t.name || "").trim().toLowerCase() === name.toLowerCase());
  if (duplicate) return toast("Ya existe una plantilla con ese nombre. Usa Actualizar o cambia el nombre.");

  const nextId = `qst_${uid()}`;
  templates.push({ id: nextId, name, items: sanitizedItems, createdAt: nowIso, updatedAt: nowIso });
  state.quoteServiceTemplates = normalizeQuoteServiceTemplatesCollection(templates);
  renderQuoteServiceTemplateSelect(nextId);
  persist();
  toast("Plantilla de servicios guardada.");
}

async function deleteQuoteServiceTemplateSelected() {
  const selectedId = String(el.quoteServiceTemplateSelect?.value || "").trim();
  if (!selectedId) return toast("Selecciona una plantilla para eliminar.");
  const templates = normalizeQuoteServiceTemplatesCollection(getQuoteServiceTemplatesFromState());
  const target = templates.find((t) => String(t.id) === selectedId);
  if (!target) return toast("Plantilla no encontrada.");
  const ok = await modernConfirm({
    title: "Eliminar plantilla",
    message: `Esta seguro de eliminar la plantilla "${target.name}"?`,
    confirmText: "Si, eliminar",
    cancelText: "No",
  });
  if (!ok) return;
  state.quoteServiceTemplates = templates.filter((t) => String(t.id) !== selectedId);
  renderQuoteServiceTemplateSelect("");
  if (el.quoteServiceTemplateName) el.quoteServiceTemplateName.value = "";
  persist();
  toast("Plantilla eliminada.");
}
function applyQuoteSnapshotToDraft(snapshot) {
  if (!quoteDraft || !snapshot) return;
  const keepVersions = normalizeQuoteVersionHistory(quoteDraft.versions);
  const keepCurrentVersion = Number(quoteDraft.version || (keepVersions.length + 1));
  const next = cloneQuoteSnapshot(snapshot, snapshot.version);
  quoteDraft = {
    ...quoteDraft,
    ...next,
    version: keepCurrentVersion,
    versions: keepVersions,
  };
  quoteDraft.paymentType = quotePaymentTypesToStorage(quoteDraft.paymentType || "");
  quoteDraft.advances = normalizeQuoteAdvancesForSnapshot(quoteDraft.advances);
  quoteDraft.currency = normalizeQuoteCurrency(quoteDraft.currency);
  renderCompaniesSelect(quoteDraft.companyId);
  renderQuoteManagerSelect(quoteDraft.companyId, quoteDraft.managerId || null);
  renderQuoteTemplateSelect(quoteDraft.templateId || "");
  renderQuoteServiceTemplateSelect("");
  fillQuoteHeaderFields(true, true);
  renderQuoteCurrencyPicker();
  setQuotePaymentTypesOnForm(quoteDraft.paymentType || "");
  if (el.quoteDueDate) el.quoteDueDate.dataset.autoDefault = quoteDraft.dueDate ? "0" : "1";
  syncQuoteDateDefaults({ forceDueDate: !quoteDraft.dueDate });
  renderQuoteServiceDateSelect();
  renderQuoteItems();
  syncPaxQuantityItems();
  renderQuoteVersionControls();
}

function getQuoteRangeDates() {
  const eventId = el.quoteEventId?.value;
  const ev = state.events.find(x => x.id === eventId);
  if (!ev) return [];
  const series = getEventSeries(ev).slice().sort((a, b) => a.date.localeCompare(b.date));
  const firstDate = series[0]?.date || ev.date;
  const lastDate = series[series.length - 1]?.date || ev.date;
  return listDatesBetween(firstDate, lastDate);
}

function compactMenuMontajeEntries(entries) {
  const list = Array.isArray(entries) ? entries : [];
  const byKey = new Map();
  for (const raw of list) {
    const date = String(raw?.date || "").trim();
    const salon = String(raw?.salon || "").trim();
    if (!date || !salon) continue;
    const qtyRaw = String(raw?.menuQty ?? "").trim();
    const qtyNum = qtyRaw === "" ? "" : Math.max(0, Math.floor(Number(qtyRaw) || 0));
    const key = `${date}|${salon}`;
    byKey.set(key, {
      date,
      salon,
      menuTitle: String(raw?.menuTitle || "").trim(),
      menuQty: qtyNum === 0 && qtyRaw === "" ? "" : qtyNum,
      menuDescription: String(raw?.menuDescription || "").trim(),
      montajeDescription: String(raw?.montajeDescription || "").trim(),
      menuSelection: raw?.menuSelection ? deepClone(raw.menuSelection) : null,
      montajeSelection: raw?.montajeSelection ? deepClone(raw.montajeSelection) : null,
    });
  }
  return Array.from(byKey.values()).sort((a, b) => {
    const d = String(a.date || "").localeCompare(String(b.date || ""));
    if (d !== 0) return d;
    return String(a.salon || "").localeCompare(String(b.salon || ""));
  });
}

function inflateMenuMontajeEntries(entries) {
  const compact = compactMenuMontajeEntries(entries);
  const nowIso = new Date().toISOString();
  return compact.map((row) => ({
    id: uid(),
    date: row.date,
    salon: row.salon,
    menuTitle: row.menuTitle,
    menuQty: row.menuQty,
    menuDescription: row.menuDescription,
    montajeDescription: row.montajeDescription,
    menuSelection: row.menuSelection ? deepClone(row.menuSelection) : null,
    montajeSelection: row.montajeSelection ? deepClone(row.montajeSelection) : null,
    updatedAt: nowIso,
  }));
}

function areMenuMontajeEntriesEqual(a, b) {
  const aa = compactMenuMontajeEntries(a);
  const bb = compactMenuMontajeEntries(b);
  if (aa.length !== bb.length) return false;
  for (let i = 0; i < aa.length; i += 1) {
    const x = aa[i];
    const y = bb[i];
    if (!y) return false;
    if (x.date !== y.date) return false;
    if (x.salon !== y.salon) return false;
    if (String(x.menuTitle || "") !== String(y.menuTitle || "")) return false;
    if (String(x.menuQty || "") !== String(y.menuQty || "")) return false;
    if (String(x.menuDescription || "") !== String(y.menuDescription || "")) return false;
    if (String(x.montajeDescription || "") !== String(y.montajeDescription || "")) return false;
    if (JSON.stringify(x.menuSelection || null) !== JSON.stringify(y.menuSelection || null)) return false;
    if (JSON.stringify(x.montajeSelection || null) !== JSON.stringify(y.montajeSelection || null)) return false;
  }
  return true;
}

function normalizeMenuMontajeVersionHistory(rawVersions) {
  const list = Array.isArray(rawVersions) ? rawVersions : [];
  return list
    .map((v, idx) => ({
      version: Math.max(1, Number(v?.version || (idx + 1))),
      entries: compactMenuMontajeEntries(v?.entries),
      savedAt: String(v?.savedAt || "").trim(),
    }))
    .sort((a, b) => Number(a.version || 0) - Number(b.version || 0));
}

function ensureMenuMontajeModel() {
  if (!quoteDraft) return { currentVersion: 1, versions: [] };
  let versions = normalizeMenuMontajeVersionHistory(quoteDraft.menuMontajeVersions);
  if (!versions.length) {
    const seedEntries = compactMenuMontajeEntries(quoteDraft.menuMontajeEntries);
    versions = [{ version: 1, entries: seedEntries, savedAt: String(quoteDraft.quotedAt || new Date().toISOString()) }];
  }
  let currentVersion = Math.max(1, Number(quoteDraft.menuMontajeVersion || versions[versions.length - 1]?.version || 1));
  if (!versions.some((v) => Number(v.version) === currentVersion)) currentVersion = Number(versions[versions.length - 1].version || 1);
  quoteDraft.menuMontajeVersions = versions;
  quoteDraft.menuMontajeVersion = currentVersion;
  const active = versions.find((v) => Number(v.version) === currentVersion) || versions[versions.length - 1];
  quoteDraft.menuMontajeEntries = inflateMenuMontajeEntries(active?.entries);
  return { currentVersion, versions };
}

function getMenuMontajeVersionSnapshot(versionNumber) {
  const model = ensureMenuMontajeModel();
  const target = Number(versionNumber || model.currentVersion);
  const found = model.versions.find((v) => Number(v.version) === target);
  if (found) return found;
  return model.versions[model.versions.length - 1] || { version: 1, entries: [], savedAt: "" };
}

function ensureMenuMontajeDraft() {
  ensureMenuMontajeModel();
  return quoteDraft.menuMontajeEntries;
}

function getQuoteDateSalonCombos() {
  const eventId = String(el.quoteEventId?.value || "").trim();
  const ev = state.events.find((x) => String(x.id || "") === eventId);
  if (!ev) return [];
  const series = getEventSeries(ev).slice().sort((a, b) => {
    const d = String(a.date || "").localeCompare(String(b.date || ""));
    if (d !== 0) return d;
    const s = String(a.salon || "").localeCompare(String(b.salon || ""));
    if (s !== 0) return s;
    return String(a.startTime || "").localeCompare(String(b.startTime || ""));
  });
  const seen = new Set();
  const combos = [];
  for (const row of series) {
    const date = String(row?.date || "").trim();
    const salon = String(row?.salon || "").trim();
    if (!date || !salon) continue;
    const key = `${date}|${salon}`;
    if (seen.has(key)) continue;
    seen.add(key);
    combos.push({ key, date, salon });
  }
  return combos;
}

function getCurrentQuoteHistoryContext() {
  const eventId = String(el.quoteEventId?.value || "").trim();
  if (!eventId) return null;
  const ev = (state.events || []).find((x) => String(x.id || "") === eventId);
  if (!ev) return null;
  const key = reservationKeyFromEvent(ev);
  if (!key) return null;
  return { ev, key };
}

// Sincroniza el borrador de cotizacion con el formulario sin validar todavia.
// Sirve para mantener el estado interno actualizado mientras el usuario escribe,
// pero respeta valores sensibles como el codigo generado automaticamente.
function syncQuoteDraftFromQuoteFormLoose() {
  if (!quoteDraft) return;
  quoteDraft.companyId = String(el.quoteCompany?.value || quoteDraft.companyId || "").trim();
  quoteDraft.managerId = String(el.quoteManagerSelect?.value || quoteDraft.managerId || "").trim();
  quoteDraft.contact = String(el.quoteContact?.value || quoteDraft.contact || "").trim();
  quoteDraft.email = String(el.quoteEmail?.value || quoteDraft.email || "").trim();
  quoteDraft.billTo = String(el.quoteBillTo?.value || quoteDraft.billTo || "").trim();
  quoteDraft.address = String(el.quoteAddress?.value || quoteDraft.address || "").trim();
  quoteDraft.eventType = String(el.quoteEventType?.value || quoteDraft.eventType || "").trim();
  quoteDraft.venue = String(el.quoteVenue?.value || quoteDraft.venue || "").trim();
  quoteDraft.schedule = String(el.quoteSchedule?.value || quoteDraft.schedule || "").trim();
  quoteDraft.code = String(quoteDraft.code || el.quoteCode?.value || "").trim();
  quoteDraft.docDate = String(el.quoteDocDate?.value || quoteDraft.docDate || "").trim();
  quoteDraft.phone = String(el.quotePhone?.value || quoteDraft.phone || "").trim();
  quoteDraft.nit = String(el.quoteNIT?.value || quoteDraft.nit || "").trim();
  quoteDraft.people = String(el.quotePeople?.value || quoteDraft.people || "").trim();
  quoteDraft.eventDate = String(el.quoteEventDate?.value || quoteDraft.eventDate || "").trim();
  quoteDraft.folio = String(el.quoteFolio?.value || quoteDraft.folio || "").trim();
  quoteDraft.endDate = String(el.quoteEndDate?.value || quoteDraft.endDate || "").trim();
  quoteDraft.dueDate = String(el.quoteDueDate?.value || quoteDraft.dueDate || "").trim();
  quoteDraft.paymentType = quotePaymentTypesToStorage(getSelectedQuotePaymentTypesFromForm().length ? getSelectedQuotePaymentTypesFromForm() : quoteDraft.paymentType);
  quoteDraft.currency = normalizeQuoteCurrency(quoteDraft.currency);
  quoteDraft.internalNotes = String(el.quoteInternalNotes?.value || quoteDraft.internalNotes || "").trim();
  quoteDraft.notes = quoteDraft.internalNotes;
}

function renderMenuMontajeEntriesTable() {
  if (!el.mmEntriesBody || !quoteDraft) return;
  const entries = ensureMenuMontajeDraft()
    .slice()
    .sort((a, b) => {
      const d = String(a.date || "").localeCompare(String(b.date || ""));
      if (d !== 0) return d;
      return String(a.salon || "").localeCompare(String(b.salon || ""));
    });
  el.mmEntriesBody.innerHTML = "";
  if (!entries.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="5">Sin informacion de menu/montaje.</td>`;
    el.mmEntriesBody.appendChild(tr);
    return;
  }
  for (const item of entries) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(String(item.date || "-"))}</td>
      <td>${escapeHtml(String(item.salon || "-"))}</td>
      <td>${escapeHtml(String(item.menuTitle || "-"))}</td>
      <td>${escapeHtml(String(item.menuQty || "-"))}</td>
      <td>${escapeHtml(String(item.montajeDescription || "-").slice(0, 120))}</td>
    `;
    el.mmEntriesBody.appendChild(tr);
  }
}

function renderMenuMontajeVersionControls() {
  if (!el.mmVersionSelect || !quoteDraft) return;
  const model = ensureMenuMontajeModel();
  const previous = Number(menuMontajeSelectedVersion || model.currentVersion);
  el.mmVersionSelect.innerHTML = "";
  for (const v of model.versions.slice().sort((a, b) => Number(b.version || 0) - Number(a.version || 0))) {
    const opt = document.createElement("option");
    const stamp = formatQuoteSentAtLabel(v.savedAt);
    opt.value = String(v.version);
    opt.textContent = stamp ? `V${v.version} - ${stamp}` : `V${v.version}`;
    el.mmVersionSelect.appendChild(opt);
  }
  const selected = model.versions.some((v) => Number(v.version) === previous)
    ? previous
    : model.currentVersion;
  menuMontajeSelectedVersion = selected;
  el.mmVersionSelect.value = String(selected);
}

function loadMenuMontajeFormByKey(key) {
  if (!quoteDraft) return;
  const entries = ensureMenuMontajeDraft();
  const [date, salon] = String(key || "").split("|");
  const found = entries.find((x) => String(x.date || "") === String(date || "") && String(x.salon || "") === String(salon || ""));
  el.mmMenuTitle.value = String(found?.menuTitle || "");
  el.mmMenuQty.value = found?.menuQty === null || found?.menuQty === undefined || found?.menuQty === "" ? "" : String(found.menuQty);
  el.mmMenuDescription.value = String(found?.menuDescription || "");
  el.mmMontajeDescription.value = String(found?.montajeDescription || "");
  updateMenuMontajeCounters();
}

function getMenuMontajeSnippet(snippetKey) {
  switch (String(snippetKey || "").trim()) {
    case "menu_section":
      return "\n[SECCION MENU]\n- ";
    case "montaje_section":
      return "\n[SECCION MONTAJE]\n- ";
    case "separator":
      return "\n------------------------------\n\n";
    case "bullet":
      return "\n- ";
    case "line_break":
      return "\n";
    default:
      return "";
  }
}

function renderMenuMontajeRichText(rawText) {
  const value = String(rawText || "").trim();
  if (!value) return `<p class="mmReportText">-</p>`;
  const lines = String(rawText || "").split(/\r?\n/);
  const htmlParts = [];
  let paragraphBuffer = [];
  let sectionTitle = "";
  let sectionItems = [];
  const flushParagraph = () => {
    if (!paragraphBuffer.length) return;
    const chunk = paragraphBuffer.join("\n");
    htmlParts.push(`<p class="mmReportText">${escapeHtml(chunk)}</p>`);
    paragraphBuffer = [];
  };
  const flushSection = () => {
    if (!sectionTitle) return;
    if (/^PLATO\s+\d+$/i.test(sectionTitle) && sectionItems.length) {
      const compactSegments = sectionItems
        .flatMap((item) => String(item || "").split(/\s+\|\s+/))
        .map((item) => String(item || "").trim())
        .filter(Boolean);
      const rows = compactSegments.map((segment) => {
        const match = segment.match(/^([^()]+?)\s*\((.*)\)$/);
        if (match) {
          return {
            label: String(match[1] || "").trim().toUpperCase(),
            value: String(match[2] || "").trim() || "-",
          };
        }
        return { label: "DETALLE", value: segment };
      });
      htmlParts.push(`
        <section class="mmReportSection">
          <div class="mmReportSubtitle">${escapeHtml(sectionTitle)}</div>
          <div class="mmReportInlineFlow">
            ${rows.map((row) => `
              <span class="mmReportInlinePair"><span class="mmReportInlineLabel">[${escapeHtml(row.label)}]</span> <span class="mmReportMiniValue">${escapeHtml(row.value)}</span></span>
            `).join("")}
          </div>
        </section>
      `);
      sectionTitle = "";
      sectionItems = [];
      return;
    }
    if (/^MONTAJE$/i.test(sectionTitle) && sectionItems.length) {
      const compactSegments = sectionItems
        .flatMap((item) => String(item || "").split(/\s+\|\s+/))
        .map((item) => String(item || "").trim())
        .filter(Boolean);
      const rows = compactSegments.map((segment) => {
        const match = segment.match(/^([^()]+?)\s*\((.*)\)$/);
        if (match) {
          return {
            label: String(match[1] || "").trim().toUpperCase(),
            value: String(match[2] || "").trim() || "-",
          };
        }
        const pair = segment.match(/^([^:]+):\s*(.*)$/);
        if (pair) {
          return {
            label: String(pair[1] || "").trim().toUpperCase(),
            value: String(pair[2] || "").trim() || "-",
          };
        }
        return { label: "DETALLE", value: segment };
      });
      htmlParts.push(`
        <section class="mmReportSection">
          <div class="mmReportSubtitle">${escapeHtml(sectionTitle)}</div>
          <div class="mmReportInlineFlow">
            ${rows.map((row) => `
              <span class="mmReportInlinePair"><span class="mmReportInlineLabel">[${escapeHtml(row.label)}]</span> <span class="mmReportMiniValue">${escapeHtml(row.value)}</span></span>
            `).join("")}
          </div>
        </section>
      `);
      sectionTitle = "";
      sectionItems = [];
      return;
    }
    const textValue = sectionItems.length ? sectionItems.join(" | ") : "-";
    const itemsHtml = `<p class="mmReportSectionLine">${escapeHtml(textValue)}</p>`;
    htmlParts.push(`
      <section class="mmReportSection">
        <div class="mmReportSubtitle">${escapeHtml(sectionTitle)}</div>
        ${itemsHtml}
      </section>
    `);
    sectionTitle = "";
    sectionItems = [];
  };
  for (const line of lines) {
    const normalized = String(line || "").trim();
    if (/^<hr\s*\/?>$/i.test(normalized) || /^[-_]{6,}$/.test(normalized) || /^\[\[HR\]\]$/i.test(normalized)) {
      flushParagraph();
      flushSection();
      htmlParts.push(`<hr class="mmReportHr" />`);
      continue;
    }
    const sectionMatch = normalized.match(/^\[(.+)\]$/);
    if (sectionMatch) {
      flushParagraph();
      flushSection();
      sectionTitle = String(sectionMatch[1] || "").trim();
      continue;
    }
    if (!normalized) {
      flushParagraph();
      flushSection();
      continue;
    }
    if (sectionTitle) {
      const content = normalized.replace(/^-+\s*/, "");
      if (content) sectionItems.push(content);
      continue;
    }
    paragraphBuffer.push(line);
  }
  flushParagraph();
  flushSection();
  return htmlParts.join("") || `<p class="mmReportText">-</p>`;
}

function updateMenuMontajeCounters() {
  if (el.mmMenuDescCount) {
    const size = String(el.mmMenuDescription?.value || "").length;
    el.mmMenuDescCount.textContent = `${size} caracteres`;
  }
  if (el.mmMontajeDescCount) {
    const size = String(el.mmMontajeDescription?.value || "").length;
    el.mmMontajeDescCount.textContent = `${size} caracteres`;
  }
}

function insertTextAtCursor(textareaEl, snippet) {
  if (!textareaEl || !snippet) return;
  const start = Number.isFinite(textareaEl.selectionStart) ? textareaEl.selectionStart : textareaEl.value.length;
  const end = Number.isFinite(textareaEl.selectionEnd) ? textareaEl.selectionEnd : start;
  let textToInsert = snippet;
  if (!textareaEl.value) {
    textToInsert = textToInsert.replace(/^\n+/, "");
  } else if (start > 0 && textareaEl.value[start - 1] !== "\n" && !textToInsert.startsWith("\n")) {
    textToInsert = `\n${textToInsert}`;
  }
  if (typeof textareaEl.setRangeText === "function") {
    textareaEl.setRangeText(textToInsert, start, end, "end");
  } else {
    const left = textareaEl.value.slice(0, start);
    const right = textareaEl.value.slice(end);
    textareaEl.value = `${left}${textToInsert}${right}`;
    const pos = left.length + textToInsert.length;
    textareaEl.selectionStart = pos;
    textareaEl.selectionEnd = pos;
  }
  textareaEl.dispatchEvent(new Event("input", { bubbles: true }));
}

function loadMenuMontajeVersion(versionNumber) {
  if (!quoteDraft) return;
  const snap = getMenuMontajeVersionSnapshot(versionNumber);
  quoteDraft.menuMontajeVersion = Number(snap.version || 1);
  quoteDraft.menuMontajeEntries = inflateMenuMontajeEntries(snap.entries);
  menuMontajeSelectedVersion = Number(snap.version || 1);
  renderMenuMontajeVersionControls();
  renderMenuMontajeSelect();
  renderMenuMontajeEntriesTable();
}

function renderMenuMontajeSelect() {
  if (!el.mmDateSalonSelect || !quoteDraft) return;
  ensureMenuMontajeModel();
  const combos = getQuoteDateSalonCombos();
  const previous = String(menuMontajeSelectedKey || "").trim();
  el.mmDateSalonSelect.innerHTML = "";
  for (const c of combos) {
    const opt = document.createElement("option");
    opt.value = c.key;
    opt.textContent = `${c.date} - ${c.salon}`;
    el.mmDateSalonSelect.appendChild(opt);
  }
  if (previous && combos.some((x) => x.key === previous)) {
    menuMontajeSelectedKey = previous;
  } else {
    menuMontajeSelectedKey = combos[0]?.key || "";
  }
  el.mmDateSalonSelect.value = menuMontajeSelectedKey;
  loadMenuMontajeFormByKey(menuMontajeSelectedKey);
  if (el.mmDocNo) el.mmDocNo.value = String(quoteDraft.code || "").trim() || "(sin codigo)";
}

async function saveMenuMontajeFromModal({ updateCurrentVersion = false } = {}) {
  if (!quoteDraft) return;
  const key = String(el.mmDateSalonSelect?.value || "").trim();
  if (!key || !key.includes("|")) return toast("Selecciona fecha y salon.");
  const [date, salon] = key.split("|");
  const menuTitle = String(el.mmMenuTitle?.value || "").trim();
  const menuQtyRaw = String(el.mmMenuQty?.value || "").trim();
  const menuQty = menuQtyRaw ? Math.max(0, Number(menuQtyRaw)) : "";
  const menuDescription = String(el.mmMenuDescription?.value || "").trim();
  const montajeDescription = String(el.mmMontajeDescription?.value || "").trim();
  if (!menuTitle && !menuDescription && !montajeDescription) {
    return toast("Agrega al menos menu o montaje para guardar.");
  }

  const entries = ensureMenuMontajeDraft();
  const idx = entries.findIndex((x) => String(x.date || "") === date && String(x.salon || "") === salon);
  const row = {
    id: idx >= 0 ? String(entries[idx].id || uid()) : uid(),
    date,
    salon,
    menuTitle,
    menuQty: menuQty === "" ? "" : Number.isFinite(menuQty) ? Math.floor(menuQty) : "",
    menuDescription,
    montajeDescription,
    updatedAt: new Date().toISOString(),
  };
  if (idx >= 0) entries[idx] = row;
  else entries.push(row);

  syncQuoteDraftFromQuoteFormLoose();
  const model = ensureMenuMontajeModel();
  let targetVersion = Number(menuMontajeSelectedVersion || model.currentVersion || 1);
  const nowIso = new Date().toISOString();
  const compactEntries = compactMenuMontajeEntries(entries);
  let createdNewVersion = false;
  let unchanged = false;
  if (updateCurrentVersion) {
    const targetIdx = model.versions.findIndex((v) => Number(v.version) === targetVersion);
    if (targetIdx >= 0) {
      model.versions[targetIdx] = {
        ...model.versions[targetIdx],
        entries: compactEntries,
        savedAt: nowIso,
      };
    } else {
      model.versions.push({ version: targetVersion, entries: compactEntries, savedAt: nowIso });
    }
  } else {
    const currentSnapshot = model.versions.find((v) => Number(v.version || 0) === Number(targetVersion || 0))
      || model.versions[model.versions.length - 1]
      || null;
    if (currentSnapshot && areMenuMontajeEntriesEqual(currentSnapshot.entries, compactEntries)) {
      unchanged = true;
    } else {
      const nextVersion = Math.max(0, ...model.versions.map((v) => Number(v.version || 0))) + 1;
      targetVersion = nextVersion;
      model.versions.push({ version: nextVersion, entries: compactEntries, savedAt: nowIso });
      createdNewVersion = true;
    }
  }
  quoteDraft.menuMontajeVersions = normalizeMenuMontajeVersionHistory(model.versions);
  quoteDraft.menuMontajeVersion = targetVersion;
  quoteDraft.menuMontajeEntries = deepClone(entries);

  if (!String(quoteDraft.code || "").trim()) {
    const code = await requestServerQuoteCode();
    quoteDraft.code = code || buildQuoteCode();
    if (el.quoteCode) el.quoteCode.value = quoteDraft.code;
  }
  if (el.mmDocNo) el.mmDocNo.value = quoteDraft.code;
  quoteDraft.quotedAt = nowIso;
  menuMontajeSelectedVersion = targetVersion;
  renderMenuMontajeVersionControls();

  const eventId = String(el.quoteEventId?.value || "").trim();
  const ev = (state.events || []).find((x) => String(x.id || "") === eventId);
  if (ev) {
    const reservationKey = reservationKeyFromEvent(ev);
    const series = getEventSeries(ev);
    for (const item of series) {
      item.quote = deepClone(quoteDraft);
    }
    quoteDraft = deepClone(quoteDraft);
    appendHistoryByKey(reservationKey, ev.userId || "", unchanged
      ? `Menu & Montaje verificado sin cambios (V${targetVersion}).`
      : `Menu & Montaje ${updateCurrentVersion ? "actualizado" : "guardado"} en V${targetVersion}.`);
    persist();
    render();
    renderQuoteVersionControls();
    renderMenuMontajeSelect();
    renderMenuMontajeEntriesTable();
    toast(unchanged
      ? `Sin cambios detectados. Se mantiene V${targetVersion}.`
      : (updateCurrentVersion
        ? `Menu & Montaje actualizado en V${targetVersion}.`
        : (createdNewVersion
          ? `Menu & Montaje guardado. Version V${targetVersion} creada.`
          : `Menu & Montaje guardado en V${targetVersion}.`)));
    return;
  }

  renderMenuMontajeSelect();
  renderMenuMontajeEntriesTable();
  toast(unchanged
    ? `Sin cambios detectados. Se mantiene V${targetVersion}.`
    : (updateCurrentVersion
      ? `Menu & Montaje actualizado en V${targetVersion}.`
      : (createdNewVersion
        ? `Menu & Montaje guardado en borrador (V${targetVersion}).`
        : `Menu & Montaje guardado en borrador (V${targetVersion}).`)));
}

function buildMenuMontajeReportHtml(ev, quoteLike, reportMode = "full") {
  const quote = quoteLike || {};
  const entries = (Array.isArray(quote?.menuMontajeEntries) ? quote.menuMontajeEntries : [])
    .filter((x) => String(x?.date || "").trim() && String(x?.salon || "").trim());
  if (!entries.length) return "";
  const company = (state.companies || []).find((c) => String(c.id || "") === String(quote.companyId || ""));
  const institutionName = String(company?.name || quote.companyName || "-").trim() || "-";
  const manager = company?.managers?.find((m) => String(m.id || "") === String(quote.managerId || ""));
  const seller = (state.users || []).find((u) => String(u.id || "") === String(ev?.userId || ""));

  const byDate = new Map();
  for (const row of entries) {
    if (!byDate.has(row.date)) byDate.set(row.date, []);
    byDate.get(row.date).push(row);
  }
  const orderedDates = Array.from(byDate.keys()).sort((a, b) => a.localeCompare(b));

  const sectionHtml = orderedDates.map((date) => {
    const rows = byDate.get(date) || [];
    rows.sort((a, b) => String(a.salon || "").localeCompare(String(b.salon || "")));
    const blocks = rows.map((r) => `
      <div class="mmReportBlock">
        <h2 class="mmReportTitle">MENU - ${escapeHtml(String(r.salon || "").toUpperCase())} - ${escapeHtml(date)}</h2>
        ${renderMenuMontajeRichText(String(r.menuDescription || ""))}
      </div>
      ${reportMode === "menu_only" ? "" : `
      <div class="mmReportBlock">
        <h2 class="mmReportTitle">MONTAJE - ${escapeHtml(String(r.salon || "").toUpperCase())} - ${escapeHtml(date)}</h2>
        ${renderMenuMontajeRichText(String(r.montajeDescription || ""))}
      </div>`}
    `).join("");
    return `
      <article class="mmReportCard" style="page-break-after: always;">
        <div class="mmReportHead">${escapeHtml(institutionName)} - MENU & MONTAJE - ${escapeHtml(date)}</div>
        <div class="mmReportMeta">
          <div><b>Encargado evento:</b> ${escapeHtml(String(manager?.name || quote.contact || "-"))}</div>
          <div><b>No cotizacion:</b> ${escapeHtml(String(quote.code || "-"))}</div>
          <div><b>Tipo evento:</b> ${escapeHtml(String(quote.eventType || ev?.name || "-"))}</div>
          <div><b>Fecha cotizacion:</b> ${escapeHtml(String(quote.docDate || "-"))}</div>
          <div><b>Horario:</b> ${escapeHtml(String(quote.schedule || `${ev?.startTime || ""} a ${ev?.endTime || ""}`.trim() || "-"))}</div>
          <div><b>Telefono:</b> ${escapeHtml(String(quote.phone || manager?.phone || "-"))}</div>
          <div><b>Fecha evento:</b> ${escapeHtml(date)}</div>
          <div><b>No Pax:</b> ${escapeHtml(String(quote.people || ev?.pax || "-"))}</div>
          <div><b>No Folio:</b> ${escapeHtml(String(quote.folio || "0"))}</div>
          <div><b>Vendedor:</b> ${escapeHtml(String(seller?.fullName || seller?.name || "-"))}</div>
        </div>
        ${blocks}
      </article>
    `;
  }).join("");

  const liveStyles = collectCurrentDocumentStyles();
  const docTitle = reportMode === "menu_only"
    ? `${institutionName} - INFORME MENU`
    : `${institutionName} - MENU & MONTAJE`;
  return `<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>${escapeHtml(docTitle)}</title>
  ${liveStyles ? `<style data-mm-live-styles>${liveStyles}</style>` : ""}
  <style>
    :root{
      --ink:#0f172a;
      --line:#bdd0e9;
      --line2:#d9e2f2;
      --soft:#edf5ff;
      --brand:#0f3c67;
      --brand2:#165d90;
    }
    *{ box-sizing:border-box; }
    html,body{ margin:0; padding:0; }
    body{
      font-family:"Segoe UI", Arial, sans-serif;
      background:#eef3fb;
      color:var(--ink);
      -webkit-print-color-adjust: exact;
      print-color-adjust: exact;
    }
    .mmReportWrap{
      padding:18px;
    }
    .mmReportCard{
      max-width:1100px;
      margin:0 auto 16px;
      border:1px solid var(--line);
      border-radius:12px;
      overflow:hidden;
      background:#ffffff;
      box-shadow:0 10px 28px rgba(15,23,42,0.12);
    }
    .mmReportHead{
      background:linear-gradient(135deg, var(--brand), var(--brand2));
      color:#fff;
      padding:12px 14px;
      font-size:20px;
      font-weight:900;
      letter-spacing:.3px;
      text-transform:uppercase;
    }
    .mmReportMeta{
      display:grid;
      grid-template-columns:repeat(2, minmax(240px, 1fr));
      gap:8px 12px;
      padding:12px 14px;
      background:var(--soft);
      border-top:1px solid #c9d8ee;
      border-bottom:1px solid #c9d8ee;
      font-size:14px;
    }
    .mmReportBlock{
      padding:14px;
      border-bottom:1px solid var(--line2);
    }
    .mmReportBlock:last-child{ border-bottom:none; }
    .mmReportTitle{
      margin:0 0 8px;
      font-size:18px;
      line-height:1.1;
      font-weight:900;
      color:#0a3f67;
      text-transform:uppercase;
    }
    .mmReportText{
      margin:0;
      white-space:pre-wrap;
      font-size:12.5px;
      line-height:1.35;
      color:var(--ink);
    }
    .mmReportSection{
      margin-top:10px;
    }
    .mmReportSection:first-child{
      margin-top:0;
    }
    .mmReportSubtitle{
      margin:0 0 6px;
      font-size:15px;
      font-weight:900;
      letter-spacing:.45px;
      color:#0a3f67;
      text-transform:uppercase;
      border-left:4px solid #0a5a92;
      padding-left:8px;
    }
    .mmReportStack{
      display:grid;
      gap:6px;
      margin-left:10px;
    }
    .mmReportInlineFlow{
      display:flex;
      flex-wrap:wrap;
      gap:8px 18px;
      margin-left:10px;
      align-items:baseline;
    }
    .mmReportInlinePair{
      display:inline-flex;
      align-items:baseline;
      gap:4px;
      white-space:nowrap;
    }
    .mmReportInlineLabel{
      font-size:12.5px;
      line-height:1.45;
      font-weight:900;
      letter-spacing:.2px;
      text-transform:uppercase;
      color:#0a4b80;
      white-space:nowrap;
    }
    .mmReportMiniValue{
      font-size:12.5px;
      line-height:1.45;
      color:#0f172a;
      white-space:nowrap;
    }
    @media print {
      .mmReportGrid{
        grid-template-columns: repeat(2, minmax(260px, 1fr));
      }
    }
    .mmReportSectionLine{
      margin:0 0 8px 12px;
      font-size:12.5px;
      line-height:1.45;
      color:#0f172a;
    }
    .mmReportHr{
      margin:10px 0;
      border:none;
      border-top:1px solid #9ab4d6;
    }
    @page { size: auto; margin: 10mm; }
    @media print {
      body{ background:#ffffff; }
      .mmReportWrap{ padding:0; }
      article.mmReportCard{
        box-shadow:none;
        margin:0 0 10mm 0;
        page-break-after:always;
      }
    }
  </style>
</head>
<body>
  <div class="mmReportWrap">
    ${sectionHtml}
  </div>
</body>
</html>`;
}

function collectCurrentDocumentStyles() {
  try {
    const sheets = Array.from(document.styleSheets || []);
    let cssText = "";
    for (const sheet of sheets) {
      try {
        const rules = Array.from(sheet.cssRules || []);
        cssText += rules.map((r) => r.cssText).join("\n");
      } catch (_) {
        // Ignore stylesheets that are not readable due to browser security rules.
      }
    }
    return cssText.trim();
  } catch (_) {
    return "";
  }
}

function openMenuMontajeReportDocument(ev, quoteLike, reportMode = "full") {
  const html = buildMenuMontajeReportHtml(ev, quoteLike, reportMode);
  if (!html) return toast("No hay datos de menu/montaje para imprimir.");
  const w = window.open("about:blank", "_blank");
  if (!w) return toast("Habilita ventanas emergentes para generar el informe.");
  w.document.open();
  w.document.write(html);
  w.document.close();
  w.focus();
}

function printMenuMontajeByDay() {
  if (!quoteDraft) return;
  const eventId = String(el.quoteEventId?.value || "").trim();
  const ev = (state.events || []).find((x) => String(x.id || "") === eventId);
  if (!ev) return toast("No se encontro el evento para imprimir.");
  const entries = ensureMenuMontajeDraft().filter((x) => String(x.date || "").trim() && String(x.salon || "").trim());
  if (!entries.length) return toast("No hay datos de menu/montaje para imprimir.");
  openMenuMontajeReportDocument(ev, quoteDraft, "full");
}

function openMenuMontajeModal() {
  if (!quoteDraft) return toast("Primero abre una cotizacion.");
  const model = ensureMenuMontajeModel();
  menuMontajeSelectedVersion = Number(model.currentVersion || 1);
  loadMenuMontajeVersion(menuMontajeSelectedVersion);
  renderMenuMontajeVersionControls();
  renderMenuMontajeSelect();
  renderMenuMontajeEntriesTable();
  if (el.menuMontajeBackdrop) el.menuMontajeBackdrop.hidden = false;
}

function closeMenuMontajeModal() {
  if (el.menuMontajeBackdrop) el.menuMontajeBackdrop.hidden = true;
}

function closeMenuMontajeSelectableModal() {
  if (el.menuMontajeSelectableBackdrop) el.menuMontajeSelectableBackdrop.hidden = true;
}

function selectedIdsUnionFromTwoLists(a, b) {
  const set = new Set();
  for (const n of selectedIdsFromChecklist(a)) set.add(n);
  for (const n of selectedIdsFromChecklist(b)) set.add(n);
  return Array.from(set.values());
}

function setChecklistCheckedByIds(container, idSet) {
  if (!container) return;
  const checks = container.querySelectorAll("input[type='checkbox']");
  for (const node of checks) {
    const id = Number(node.value);
    node.checked = Number.isFinite(id) && idSet.has(id);
    const row = node.closest(".menuSuggestRow");
    setMenuSuggestRowDraggableByCheckbox(row);
  }
}

function syncMirrorChecklistValue(source, target, changed) {
  if (!source || !target || !changed) return;
  const id = Number(changed.value || 0);
  if (!Number.isFinite(id) || id <= 0) return;
  const mirror = target.querySelector(`input[type='checkbox'][value="${id}"]`);
  if (!mirror) return;
  mirror.checked = !!changed.checked;
  const row = mirror.closest(".menuSuggestRow");
  setMenuSuggestRowDraggableByCheckbox(row);
}

function ensureMmsCatalogDefaults() {
  if (!menuMontajeSelectableCatalogCache || typeof menuMontajeSelectableCatalogCache !== "object") {
    menuMontajeSelectableCatalogCache = {
      proteins: [],
      preparationsByProtein: new Map(),
      suggestionsByPair: new Map(),
      suggestionsInFlightByPair: new Map(),
      salsas: [],
      guarniciones: [],
      postres: [],
      bebidas: [],
      comentarios: [],
      montajeTipos: [],
      montajeAdicionales: [],
    };
  }
  if (!(menuMontajeSelectableCatalogCache.preparationsByProtein instanceof Map)) {
    menuMontajeSelectableCatalogCache.preparationsByProtein = new Map();
  }
  if (!(menuMontajeSelectableCatalogCache.suggestionsByPair instanceof Map)) {
    menuMontajeSelectableCatalogCache.suggestionsByPair = new Map();
  }
  if (!(menuMontajeSelectableCatalogCache.suggestionsInFlightByPair instanceof Map)) {
    menuMontajeSelectableCatalogCache.suggestionsInFlightByPair = new Map();
  }
}

function ensureMenuMontajeSectionsStore() {
  if (!Array.isArray(state.menuMontajeSections)) state.menuMontajeSections = ["General"];
  state.menuMontajeSections = state.menuMontajeSections
    .map((s) => String(s || "").trim())
    .filter(Boolean);
  if (!state.menuMontajeSections.length) state.menuMontajeSections = ["General"];
  return state.menuMontajeSections;
}

function renderMmsMenuSectionSelect(selected = "") {
  if (!el.mmsMenuSection) return;
  const sections = ensureMenuMontajeSectionsStore();
  el.mmsMenuSection.innerHTML = "";
  for (const s of sections) {
    const opt = document.createElement("option");
    opt.value = s;
    opt.textContent = s;
    el.mmsMenuSection.appendChild(opt);
  }
  const preferred = String(selected || "").trim();
  if (preferred) el.mmsMenuSection.value = preferred;
  if (!el.mmsMenuSection.value && el.mmsMenuSection.options.length) {
    el.mmsMenuSection.value = el.mmsMenuSection.options[0].value;
  }
}

function addMmsMenuSectionFromInput() {
  const value = String(el.mmsMenuSectionInput?.value || "").trim();
  if (!value) return toast("Escribe un nombre de seccion.");
  const sections = ensureMenuMontajeSectionsStore();
  const exists = sections.some((s) => String(s || "").trim().toLowerCase() === value.toLowerCase());
  if (exists) return toast("Esa seccion ya existe.");
  sections.push(value);
  state.menuMontajeSections = sections;
  persist();
  renderMmsMenuSectionSelect(value);
  if (el.mmsMenuSectionInput) {
    el.mmsMenuSectionInput.value = "";
    el.mmsMenuSectionInput.focus();
  }
  refreshMmsDescriptionAuto();
  toast("Seccion agregada.");
}

function ensureMenuMontajeBebidasStore() {
  if (!Array.isArray(state.menuMontajeBebidas)) state.menuMontajeBebidas = [];
  state.menuMontajeBebidas = state.menuMontajeBebidas
    .map((x) => {
      if (typeof x === "string") return { id: uid(), nombre: String(x).trim(), activo: true };
      return {
        id: String(x?.id || uid()).trim() || uid(),
        nombre: String(x?.nombre || x?.name || "").trim(),
        activo: x?.activo === false ? false : true,
      };
    })
    .filter((x) => x.nombre);
  return state.menuMontajeBebidas;
}

function addMmsBebidaFromInput() {
  const nombre = String(el.mmsBebidaInput?.value || "").trim();
  if (!nombre) return toast("Escribe una bebida.");
  const rows = ensureMenuMontajeBebidasStore();
  if (rows.some((x) => String(x.nombre || "").toLowerCase() === nombre.toLowerCase())) {
    return toast("Esa bebida ya existe.");
  }
  rows.push({ id: uid(), nombre, activo: true });
  state.menuMontajeBebidas = rows;
  if (el.mmsBebidaInput) el.mmsBebidaInput.value = "";
  persist();
  renderMmsStageOptions();
  renderMmsComandaPreview();
  toast("Bebida agregada.");
}

function getMmsSuggestedSalsaIds() {
  return Array.isArray(menuMontajeSelectableCatalogCache?.suggestedSalsaIds)
    ? menuMontajeSelectableCatalogCache.suggestedSalsaIds.map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0)
    : [];
}

function setMmsStage(stage) {
  const valid = new Set(["plato", "preparacion", "salsa", "guarnicion", "postre", "bebida", "montaje_tipo", "montaje_adicional"]);
  const next = valid.has(String(stage || "")) ? String(stage) : "plato";
  mmsCurrentStage = next;
  mmsPrimaryMode = (next === "montaje_tipo" || next === "montaje_adicional") ? "montaje" : "menu";
  // Evita que un filtro previo deje vacia la siguiente etapa (ej. montaje).
  if (el.mmsStageFilter) el.mmsStageFilter.value = "";
  renderMmsPrimaryTabs();
  renderMmsStageTabs();
  renderMmsStageOptions();
}

function getMmsStagesForPrimaryMode(mode) {
  if (String(mode || "") === "montaje") return ["montaje_tipo", "montaje_adicional"];
  return ["plato", "preparacion", "salsa", "guarnicion", "postre", "bebida"];
}

function renderMmsPrimaryTabs() {
  const buttons = Array.from(el.mmsPrimaryTabs?.querySelectorAll("[data-mms-primary]") || []);
  for (const btn of buttons) {
    btn.classList.toggle("isActive", String(btn.dataset.mmsPrimary || "") === mmsPrimaryMode);
  }
}

function setMmsPrimaryMode(mode) {
  const nextMode = String(mode || "") === "montaje" ? "montaje" : "menu";
  mmsPrimaryMode = nextMode;
  const allowedStages = getMmsStagesForPrimaryMode(nextMode);
  if (!allowedStages.includes(mmsCurrentStage)) {
    mmsCurrentStage = allowedStages[0] || "plato";
  }
  if (el.mmsStageFilter) el.mmsStageFilter.value = "";
  renderMmsPrimaryTabs();
  renderMmsStageTabs();
  renderMmsStageOptions();
}

function renderMmsStageTabs() {
  const buttons = Array.from(el.mmsStageTabs?.querySelectorAll("[data-mms-stage]") || []);
  const allowedStages = new Set(getMmsStagesForPrimaryMode(mmsPrimaryMode));
  for (const btn of buttons) {
    const stage = String(btn.dataset.mmsStage || "");
    const isVisible = allowedStages.has(stage);
    btn.hidden = !isVisible;
    btn.classList.toggle("isActive", isVisible && stage === mmsCurrentStage);
  }
}

function getMmsStageLabel(stage) {
  const map = {
    plato: "Elija plato fuerte",
    preparacion: "Elija preparacion",
    salsa: "Elija salsa/aderezo",
    guarnicion: "Elija guarnicion",
    postre: "Elija postre",
    bebida: "Elija bebida",
    montaje_tipo: "Elija tipo de montaje",
    montaje_adicional: "Elija adicionales de montaje",
  };
  return map[stage] || "Elija una opcion";
}

function renderMmsStageOptions() {
  ensureMmsCatalogDefaults();
  if (el.mmsStageTitle) el.mmsStageTitle.textContent = getMmsStageLabel(mmsCurrentStage);
  if (el.btnMmsStageMoreOptions) {
    const canToggle = mmsCurrentStage === "salsa" || mmsCurrentStage === "guarnicion" || mmsCurrentStage === "postre";
    el.btnMmsStageMoreOptions.disabled = !canToggle;
    if (mmsCurrentStage === "postre") {
      el.btnMmsStageMoreOptions.textContent = mmsShowAllPostres ? "Solo sugeridas" : "Mas opciones";
    } else if (mmsCurrentStage === "salsa" || mmsCurrentStage === "guarnicion") {
      el.btnMmsStageMoreOptions.textContent = mmsShowAllGuarniciones ? "Solo sugeridas" : "Mas opciones";
    } else {
      el.btnMmsStageMoreOptions.textContent = "Mas opciones";
    }
  }
  if (!el.mmsStageOptions) return;

  const filter = String(el.mmsStageFilter?.value || "").trim().toLowerCase();
  let rows = [];
  let selected = new Set();
  let kind = mmsCurrentStage;

  if (mmsCurrentStage === "plato") {
    rows = menuMontajeSelectableCatalogCache.proteins || [];
    selected = new Set((Array.isArray(mmsSelectedPlatoItems) ? mmsSelectedPlatoItems : []).map((item) => Number(item?.platoId || 0)).filter((n) => Number.isFinite(n) && n > 0));
  } else if (mmsCurrentStage === "preparacion") {
    const pid = Number(el.mmsProtein?.value || 0);
    rows = pid > 0 ? (menuMontajeSelectableCatalogCache.preparationsByProtein.get(pid) || []) : [];
    selected = new Set((Array.isArray(mmsSelectedPlatoItems) ? mmsSelectedPlatoItems : [])
      .filter((item) => Number(item?.platoId || 0) === pid)
      .map((item) => Number(item?.preparacionId || 0))
      .filter((n) => Number.isFinite(n) && n > 0));
  } else if (mmsCurrentStage === "salsa") {
    const all = menuMontajeSelectableCatalogCache.salsas || [];
    const selectedSalsaSet = new Set((Array.isArray(mmsSelectedSalsaIds) ? mmsSelectedSalsaIds : []).map((x) => Number(x)));
    const suggestedSet = new Set(getMmsSuggestedSalsaIds().map((x) => Number(x)));
    const suggestedRows = all.filter((x) => suggestedSet.has(Number(x.id || 0)));
    rows = mmsShowAllGuarniciones ? all : (suggestedRows.length ? suggestedRows : all);
    selected = selectedSalsaSet;
  } else if (mmsCurrentStage === "guarnicion") {
    const all = menuMontajeSelectableCatalogCache.guarniciones || [];
    const selectedGuarnicionSet = new Set(getMmsSelectedGuarnicionIds().map((x) => Number(x)));
    const suggestedSet = new Set(listAllCheckboxIds(el.mmsGuarnicionesSuggested));
    const suggestedRows = all.filter((x) => suggestedSet.has(Number(x.id || 0)));
    rows = mmsShowAllGuarniciones ? all : (suggestedRows.length ? suggestedRows : all);
    selected = selectedGuarnicionSet;
  } else if (mmsCurrentStage === "postre") {
    const all = menuMontajeSelectableCatalogCache.postres || [];
    const selectedPostreSet = new Set(getMmsSelectedPostreIds().map((x) => Number(x)));
    const suggestedSet = new Set(listAllCheckboxIds(el.mmsPostresSuggested));
    const suggestedRows = all.filter((x) => suggestedSet.has(Number(x.id || 0)));
    rows = mmsShowAllPostres ? all : (suggestedRows.length ? suggestedRows : all);
    selected = selectedPostreSet;
  } else if (mmsCurrentStage === "bebida") {
    rows = menuMontajeSelectableCatalogCache.bebidas || [];
    const selectedBebidaSet = new Set((Array.isArray(mmsSelectedBebidaIds) ? mmsSelectedBebidaIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0));
    selected = selectedBebidaSet;
    kind = "bebida";
  } else if (mmsCurrentStage === "montaje_tipo") {
    rows = menuMontajeSelectableCatalogCache.montajeTipos || [];
    selected = new Set([Number(el.mmsMontajeTipo?.value || 0)]);
  } else if (mmsCurrentStage === "montaje_adicional") {
    const all = menuMontajeSelectableCatalogCache.montajeAdicionales || [];
    rows = mmsShowAllPostres ? all : all;
    selected = new Set(selectedIdsFromChecklist(el.mmsMontajeAdicionales).map((x) => Number(x)));
  }

  rows = rows.filter((x) => {
    const name = String(x?.nombre || "").trim().toLowerCase();
    return !filter || name.includes(filter);
  });

  renderMmsQuickButtonsGroup(el.mmsStageOptions, rows, selected, kind);
}

async function handleMmsStageOptionClick(kind, id) {
  if (kind === "plato") {
    const nextId = Number(id || 0);
    if (!Number.isFinite(nextId) || nextId <= 0) return;
    ensureMmsCatalogDefaults();
    const plateName = namesFromIds(menuMontajeSelectableCatalogCache.proteins, [nextId])[0] || "Plato fuerte";
    const qty = getFastMmsPlatoQty();
    mmsPlatoQty = qty;
    if (el.mmsMenuQty) el.mmsMenuQty.value = String(qty);
    if (el.mmsProtein) el.mmsProtein.value = String(nextId);
    await populateMmsPreparationsForProtein(nextId, { preserveSelection: false }).catch(() => []);
    const prepOptions = Array.from(el.mmsPreparation?.options || []).filter((opt) => String(opt.value || "").trim());
    if (!prepOptions.length) {
      commitCurrentMmsLineItem();
      const item = upsertMmsPlatoItem({ platoId: nextId, preparacionId: null, qty });
      mmsActiveLineKey = String(item?.key || "");
      clearMmsActiveComplements();
      renderMmsSelectionSummary();
      renderMmsComandaPreview();
      if (item) notifyMmsSelectionAdded("plato", item.key, plateName, qty);
      refreshMmsDescriptionAuto();
      setMmsStage("guarnicion");
      modernGuideToast("Plato agregado. Ahora elige guarniciones.");
      return;
    }
    setMmsStage("preparacion");
    if (prepOptions.length === 1) {
      await handleMmsStageOptionClick("preparacion", prepOptions[0].value);
      return;
    }
    modernGuideToast("Selecciona la preparacion para agregar este plato fuerte.");
    return;
  }
  if (kind === "preparacion") {
    const prepId = Number(id || 0);
    if (!Number.isFinite(prepId) || prepId <= 0) return;
    if (el.mmsPreparation) el.mmsPreparation.value = String(prepId);
    const platoId = Number(el.mmsProtein?.value || 0);
    if (!Number.isFinite(platoId) || platoId <= 0) return;
    const label = [namesFromIds(menuMontajeSelectableCatalogCache.proteins, [platoId])[0] || "", String(el.mmsPreparation?.selectedOptions?.[0]?.textContent || "").trim()].filter(Boolean).join(" - ") || "Plato fuerte";
    const qty = Math.max(1, Math.floor(Number(mmsPlatoQty || 1)));
    commitCurrentMmsLineItem();
    const item = upsertMmsPlatoItem({ platoId, preparacionId: prepId, qty });
    mmsActiveLineKey = String(item?.key || "");
    clearMmsActiveComplements();
    refreshMmsDescriptionAuto();
    renderMmsSelectionSummary();
    renderMmsComandaPreview();
    if (item) notifyMmsSelectionAdded("plato", item.key, label, qty);
    setMmsStage("guarnicion");
    modernGuideToast("Plato fuerte agregado. Ahora puedes elegir guarniciones.");
    refreshMmsByProteinPreparation({ preserveSelection: true }).catch(() => { });
    return;
  }
  if (kind === "salsa") {
    const n = Number(id || 0);
    const set = new Set((Array.isArray(mmsSelectedSalsaIds) ? mmsSelectedSalsaIds : []).map((x) => Number(x)));
    if (set.has(n)) set.delete(n);
    else {
      set.add(n);
      const label = namesFromIds(menuMontajeSelectableCatalogCache.salsas, [n])[0] || "Salsa";
      notifyMmsSelectionAdded("salsa", n, label, 1);
    }
    mmsSelectedSalsaIds = Array.from(set.values()).filter((x) => Number.isFinite(x) && x > 0);
    refreshMmsDescriptionAuto();
    renderMmsStageOptions();
    renderMmsComandaPreview();
    return;
  }
  if (kind === "guarnicion") {
    await toggleMmsQuickItem("guarnicion", id);
    return;
  }
  if (kind === "postre") {
    await toggleMmsQuickItem("postre", id);
    return;
  }
  if (kind === "bebida") {
    const n = Number(id || 0);
    if (!Number.isFinite(n) || n <= 0) return;
    const set = new Set((Array.isArray(mmsSelectedBebidaIds) ? mmsSelectedBebidaIds : []).map((x) => Number(x)).filter((v) => Number.isFinite(v) && v > 0));
    if (set.has(n)) {
      set.delete(n);
      deleteMmsItemQty("bebida", n);
    } else {
      ensureMmsCatalogDefaults();
      const label = namesFromIds(menuMontajeSelectableCatalogCache.bebidas, [n])[0] || "Bebida";
      const qty = await requestMmsItemQty(label, getMmsItemQty("bebida", n), "bebidas");
      if (qty === null) return;
      set.add(n);
      setMmsItemQty("bebida", n, qty);
      notifyMmsSelectionAdded("bebida", n, label, qty);
    }
    mmsSelectedBebidaIds = Array.from(set.values());
    syncMmsBebidaQtyWithSelection();
    refreshMmsDescriptionAuto();
    renderMmsSelectionSummary();
    renderMmsStageOptions();
    renderMmsComandaPreview();
    return;
  }
  if (kind === "montaje_tipo") {
    if (el.mmsMontajeTipo) el.mmsMontajeTipo.value = String(id || "");
    refreshMmsDescriptionAuto();
    renderMmsStageOptions();
    renderMmsComandaPreview();
    return;
  }
  if (kind === "montaje_adicional") {
    const n = Number(id || 0);
    const current = new Set(selectedIdsFromChecklist(el.mmsMontajeAdicionales).map((x) => Number(x)));
    if (current.has(n)) current.delete(n);
    else current.add(n);
    setMmsSelectionSets({
      guarnicionIds: getMmsSelectedGuarnicionIds(),
      postreIds: getMmsSelectedPostreIds(),
      comentarioIds: selectedIdsFromChecklist(el.mmsComentariosAll),
      adicionalIds: Array.from(current.values()),
    });
    refreshMmsDescriptionAuto();
    renderMmsStageOptions();
    renderMmsComandaPreview();
  }
}

function cancelMmsCurrentStageSelection() {
  if (mmsCurrentStage === "plato") {
    mmsSelectedPlatoItems = [];
    mmsPlatoQty = 1;
    syncMmsMenuQtyField();
    if (el.mmsProtein) el.mmsProtein.value = "";
    if (el.mmsPreparation) el.mmsPreparation.value = "";
  } else if (mmsCurrentStage === "salsa") {
    mmsSelectedSalsaIds = [];
  } else if (mmsCurrentStage === "guarnicion") {
    mmsGuarnicionQtyById = {};
    setMmsSelectionSets({
      guarnicionIds: [],
      postreIds: getMmsSelectedPostreIds(),
      comentarioIds: selectedIdsFromChecklist(el.mmsComentariosAll),
      adicionalIds: selectedIdsFromChecklist(el.mmsMontajeAdicionales),
    });
  } else if (mmsCurrentStage === "postre") {
    setMmsSelectionSets({
      guarnicionIds: getMmsSelectedGuarnicionIds(),
      postreIds: [],
      comentarioIds: selectedIdsFromChecklist(el.mmsComentariosAll),
      adicionalIds: selectedIdsFromChecklist(el.mmsMontajeAdicionales),
    });
  } else if (mmsCurrentStage === "bebida") {
    mmsSelectedBebidaIds = [];
    mmsBebidaQtyById = {};
  } else if (mmsCurrentStage === "montaje_tipo") {
    if (el.mmsMontajeTipo && el.mmsMontajeTipo.options.length) el.mmsMontajeTipo.value = el.mmsMontajeTipo.options[0].value;
  } else if (mmsCurrentStage === "montaje_adicional") {
    setMmsSelectionSets({
      guarnicionIds: getMmsSelectedGuarnicionIds(),
      postreIds: getMmsSelectedPostreIds(),
      comentarioIds: selectedIdsFromChecklist(el.mmsComentariosAll),
      adicionalIds: [],
    });
  } else if (mmsCurrentStage === "preparacion") {
    if (el.mmsPreparation && el.mmsPreparation.options.length) el.mmsPreparation.value = el.mmsPreparation.options[0].value;
  }
  refreshMmsDescriptionAuto();
  renderMmsSelectionSummary();
  renderMmsStageOptions();
  renderMmsComandaPreview();
}

async function ensureMenuMontajeSelectableCatalogLoaded(force = false) {
  ensureMmsCatalogDefaults();
  const cache = menuMontajeSelectableCatalogCache;
  if (!force && cache.proteins.length) return cache;
  const [proteins, salsas, guarniciones, postres, bebidas, comentarios, montajeTipos, montajeAdicionales, preparaciones] = await Promise.all([
    readMenuCatalog("plato_fuerte"),
    readMenuCatalog("salsa"),
    readMenuCatalog("guarnicion"),
    readMenuCatalog("postre"),
    readMenuCatalog("bebida"),
    readMenuCatalog("comentario"),
    readMenuCatalog("montaje_tipo"),
    readMenuCatalog("montaje_adicional"),
    readMenuCatalog("preparacion"),
  ]);
  cache.proteins = Array.isArray(proteins) ? proteins.filter((x) => x && x.activo !== false) : [];
  cache.salsas = Array.isArray(salsas) ? salsas.filter((x) => x && x.activo !== false) : [];
  cache.guarniciones = Array.isArray(guarniciones) ? guarniciones.filter((x) => x && x.activo !== false) : [];
  cache.postres = Array.isArray(postres) ? postres.filter((x) => x && x.activo !== false) : [];
  cache.bebidas = Array.isArray(bebidas) ? bebidas.filter((x) => x && x.activo !== false) : [];
  cache.comentarios = Array.isArray(comentarios) ? comentarios.filter((x) => x && x.activo !== false) : [];
  cache.montajeTipos = Array.isArray(montajeTipos) ? montajeTipos.filter((x) => x && x.activo !== false) : [];
  cache.montajeAdicionales = Array.isArray(montajeAdicionales) ? montajeAdicionales.filter((x) => x && x.activo !== false) : [];
  const prepsByProtein = new Map();
  for (const prep of Array.isArray(preparaciones) ? preparaciones : []) {
    if (!prep || prep.activo === false) continue;
    const platoId = Number(prep?.plato_id || prep?.platoId || 0);
    if (!Number.isFinite(platoId) || platoId <= 0) continue;
    if (!prepsByProtein.has(platoId)) prepsByProtein.set(platoId, []);
    prepsByProtein.get(platoId).push(prep);
  }
  cache.preparationsByProtein = prepsByProtein;
  return cache;
}

async function loadMmsPreparationsByProtein(proteinId) {
  ensureMmsCatalogDefaults();
  const pid = Number(proteinId || 0);
  if (!Number.isFinite(pid) || pid <= 0) return [];
  const cache = menuMontajeSelectableCatalogCache;
  if (cache.preparationsByProtein.has(pid)) return cache.preparationsByProtein.get(pid) || [];
  const rows = await readMenuCatalog("preparacion", `plato_id=${encodeURIComponent(String(pid))}`);
  const active = Array.isArray(rows) ? rows.filter((x) => x && x.activo !== false) : [];
  cache.preparationsByProtein.set(pid, active);
  return active;
}

function buildMmsNameMap(rows) {
  if (Array.isArray(rows) && mmsNameMapMemo.has(rows)) {
    return mmsNameMapMemo.get(rows);
  }
  const m = new Map();
  for (const row of Array.isArray(rows) ? rows : []) {
    const id = Number(row?.id || 0);
    if (!Number.isFinite(id) || id <= 0) continue;
    m.set(id, String(row?.nombre || "").trim());
  }
  if (Array.isArray(rows)) mmsNameMapMemo.set(rows, m);
  return m;
}

function namesFromIds(rows, ids) {
  const map = buildMmsNameMap(rows);
  const out = [];
  for (const idRaw of Array.isArray(ids) ? ids : []) {
    const id = Number(idRaw || 0);
    if (!Number.isFinite(id) || id <= 0) continue;
    const name = String(map.get(id) || "").trim();
    if (name) out.push(name);
  }
  return out;
}

function listAllCheckboxIds(container) {
  if (!container) return [];
  return Array.from(container.querySelectorAll("input[type='checkbox']"))
    .map((x) => Number(x.value || 0))
    .filter((n) => Number.isFinite(n) && n > 0);
}

function setMmsSelectionSets({ guarnicionIds = [], postreIds = [], comentarioIds = [], adicionalIds = [] } = {}) {
  const guarnicionSet = new Set((Array.isArray(guarnicionIds) ? guarnicionIds : []).map((x) => Number(x)));
  const postreSet = new Set((Array.isArray(postreIds) ? postreIds : []).map((x) => Number(x)));
  const comentarioSet = new Set((Array.isArray(comentarioIds) ? comentarioIds : []).map((x) => Number(x)));
  const adicionalSet = new Set((Array.isArray(adicionalIds) ? adicionalIds : []).map((x) => Number(x)));
  setChecklistCheckedByIds(el.mmsGuarnicionesSuggested, guarnicionSet);
  setChecklistCheckedByIds(el.mmsGuarnicionesAll, guarnicionSet);
  setChecklistCheckedByIds(el.mmsPostresSuggested, postreSet);
  setChecklistCheckedByIds(el.mmsPostresAll, postreSet);
  setChecklistCheckedByIds(el.mmsComentariosAll, comentarioSet);
  setChecklistCheckedByIds(el.mmsMontajeAdicionales, adicionalSet);
  syncMmsGuarnicionQtyWithSelection();
  syncMmsPostreQtyWithSelection();
}

function getMmsSelectedGuarnicionIds() {
  return selectedIdsUnionFromTwoLists(el.mmsGuarnicionesSuggested, el.mmsGuarnicionesAll);
}

function getMmsSelectedPostreIds() {
  return selectedIdsUnionFromTwoLists(el.mmsPostresSuggested, el.mmsPostresAll);
}

function normalizeMmsQtyMap(raw) {
  const out = {};
  if (!raw || typeof raw !== "object") return out;
  for (const [k, v] of Object.entries(raw)) {
    const id = Number(k);
    const qty = Math.max(0, Math.floor(Number(v || 0)));
    if (!Number.isFinite(id) || id <= 0 || qty <= 0) continue;
    out[id] = qty;
  }
  return out;
}

function normalizeMmsPostreQtyMap(raw) {
  return normalizeMmsQtyMap(raw);
}

function syncMmsQtyMapWithSelection(selectedIds, qtyMap) {
  const selected = new Set((Array.isArray(selectedIds) ? selectedIds : []).map((x) => Number(x)));
  const next = {};
  for (const id of selected) {
    const qty = Math.max(1, Math.floor(Number(qtyMap?.[id] || 1)));
    next[id] = qty;
  }
  return next;
}

function syncMmsPostreQtyWithSelection() {
  mmsPostreQtyById = syncMmsQtyMapWithSelection(getMmsSelectedPostreIds(), mmsPostreQtyById);
}

function syncMmsGuarnicionQtyWithSelection() {
  mmsGuarnicionQtyById = syncMmsQtyMapWithSelection(getMmsSelectedGuarnicionIds(), mmsGuarnicionQtyById);
}

function syncMmsBebidaQtyWithSelection() {
  mmsBebidaQtyById = syncMmsQtyMapWithSelection(mmsSelectedBebidaIds, mmsBebidaQtyById);
}

function getMmsItemQty(kind, id) {
  const itemId = Number(id || 0);
  if (!Number.isFinite(itemId) || itemId <= 0) return 1;
  if (kind === "guarnicion") return Math.max(1, Math.floor(Number(mmsGuarnicionQtyById[itemId] || 1)));
  if (kind === "bebida") return Math.max(1, Math.floor(Number(mmsBebidaQtyById[itemId] || 1)));
  return Math.max(1, Math.floor(Number(mmsPostreQtyById[itemId] || 1)));
}

function getFastMmsPlatoQty() {
  const direct = Math.floor(Number(el.mmsMenuQtyVisible?.value || el.mmsMenuQty?.value || mmsPlatoQty || 1));
  return Number.isFinite(direct) && direct > 0 ? direct : 1;
}

function normalizeMmsPlatoItems(raw, fallback = null) {
  const input = Array.isArray(raw) ? raw : [];
  const out = input
    .map((item) => ({
      key: String(item?.key || `pl_${uid()}`),
      platoId: Number(item?.platoId || 0),
      preparacionId: Number(item?.preparacionId || 0) || null,
      qty: Math.max(1, Math.floor(Number(item?.qty || 1))),
    }))
    .filter((item) => Number.isFinite(item.platoId) && item.platoId > 0);
  if (out.length) return out;
  const legacyPlatoId = Number(fallback?.platoId || 0);
  if (!Number.isFinite(legacyPlatoId) || legacyPlatoId <= 0) return [];
  return [{
    key: `pl_${uid()}`,
    platoId: legacyPlatoId,
    preparacionId: Number(fallback?.preparacionId || 0) || null,
    qty: Math.max(1, Math.floor(Number(fallback?.platoQty || 1))),
  }];
}

function normalizeMmsLineItems(raw, fallback = null) {
  const list = Array.isArray(raw) ? raw : [];
  const normalized = list
    .map((item) => {
      const platoId = Number(item?.platoId || 0);
      if (!Number.isFinite(platoId) || platoId <= 0) return null;
      return {
        key: String(item?.key || `ln_${uid()}`),
        platoId,
        preparacionId: Number(item?.preparacionId || 0) || null,
        qty: Math.max(1, Math.floor(Number(item?.qty || 1))),
        salsaIds: (Array.isArray(item?.salsaIds) ? item.salsaIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0),
        guarnicionIds: (Array.isArray(item?.guarnicionIds) ? item.guarnicionIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0),
        guarnicionQtys: normalizeMmsQtyMap(item?.guarnicionQtys),
        postreIds: (Array.isArray(item?.postreIds) ? item.postreIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0),
        postreQtys: normalizeMmsQtyMap(item?.postreQtys),
        bebidaIds: (Array.isArray(item?.bebidaIds) ? item.bebidaIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0),
        bebidaQtys: normalizeMmsQtyMap(item?.bebidaQtys),
        comentarioIds: (Array.isArray(item?.comentarioIds) ? item.comentarioIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0),
        comentarioLibre: String(item?.comentarioLibre || "").trim(),
      };
    })
    .filter(Boolean);
  if (normalized.length) return normalized;
  const legacyPlatos = normalizeMmsPlatoItems(fallback?.platoItems, fallback);
  if (!legacyPlatos.length) return [];
  const salsaIds = (Array.isArray(fallback?.salsaIds) ? fallback.salsaIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0);
  const guarnicionIds = (Array.isArray(fallback?.guarnicionIds) ? fallback.guarnicionIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0);
  const guarnicionQtys = normalizeMmsQtyMap(fallback?.guarnicionQtys);
  const postreIds = (Array.isArray(fallback?.postreIds) ? fallback.postreIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0);
  const postreQtys = normalizeMmsQtyMap(fallback?.postreQtys);
  const bebidaIds = (Array.isArray(fallback?.bebidaIds) ? fallback.bebidaIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0);
  const bebidaQtys = normalizeMmsQtyMap(fallback?.bebidaQtys);
  const comentarioIds = (Array.isArray(fallback?.comentarioIds) ? fallback.comentarioIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0);
  const comentarioLibre = String(fallback?.comentarioLibre || "").trim();
  return legacyPlatos.map((plato) => ({
    key: String(plato.key || `ln_${uid()}`),
    platoId: Number(plato.platoId || 0),
    preparacionId: Number(plato.preparacionId || 0) || null,
    qty: Math.max(1, Math.floor(Number(plato.qty || 1))),
    salsaIds: [...salsaIds],
    guarnicionIds: [...guarnicionIds],
    guarnicionQtys: { ...guarnicionQtys },
    postreIds: [...postreIds],
    postreQtys: { ...postreQtys },
    bebidaIds: [...bebidaIds],
    bebidaQtys: { ...bebidaQtys },
    comentarioIds: [...comentarioIds],
    comentarioLibre,
  }));
}

function getMmsPlatoItemLabel(item) {
  ensureMmsCatalogDefaults();
  const platoId = Number(item?.platoId || 0);
  const prepId = Number(item?.preparacionId || 0);
  const platoName = getMmsPlatoBaseName(item);
  const prepName = getMmsPreparationName(item);
  return [platoName, prepName].filter(Boolean).join(" - ") || "Plato fuerte";
}

function getMmsPlatoBaseName(item) {
  ensureMmsCatalogDefaults();
  const platoId = Number(item?.platoId || 0);
  return namesFromIds(menuMontajeSelectableCatalogCache.proteins, [platoId])[0] || "";
}

function getMmsPreparationName(item) {
  ensureMmsCatalogDefaults();
  const platoId = Number(item?.platoId || 0);
  const prepId = Number(item?.preparacionId || 0);
  const prepList = platoId > 0 ? (menuMontajeSelectableCatalogCache.preparationsByProtein.get(platoId) || []) : [];
  return namesFromIds(prepList, prepId > 0 ? [prepId] : [])[0] || "";
}

function getMmsTotalPlatoQty() {
  return (Array.isArray(mmsSelectedPlatoItems) ? mmsSelectedPlatoItems : [])
    .reduce((sum, item) => sum + Math.max(1, Math.floor(Number(item?.qty || 1))), 0);
}

function syncMmsMenuQtyField() {
  const total = getMmsTotalPlatoQty();
  const fallback = Math.max(1, Math.floor(Number(mmsPlatoQty || 1)));
  const value = total > 0 ? String(total) : String(fallback);
  if (el.mmsMenuQty) el.mmsMenuQty.value = value;
  if (el.mmsMenuQtyVisible) el.mmsMenuQtyVisible.value = value;
}

function upsertMmsPlatoItem({ platoId, preparacionId = null, qty = 1 } = {}) {
  const pid = Number(platoId || 0);
  const prep = Number(preparacionId || 0) || null;
  const safeQty = Math.max(1, Math.floor(Number(qty || 1)));
  if (!Number.isFinite(pid) || pid <= 0) return null;
  const list = Array.isArray(mmsSelectedPlatoItems) ? [...mmsSelectedPlatoItems] : [];
  const idx = list.findIndex((item) => Number(item?.platoId || 0) === pid && Number(item?.preparacionId || 0) === Number(prep || 0));
  if (idx >= 0) {
    list[idx] = { ...list[idx], qty: safeQty };
  } else {
    list.push({ key: `pl_${uid()}`, platoId: pid, preparacionId: prep, qty: safeQty });
  }
  mmsSelectedPlatoItems = list;
  mmsPlatoQty = safeQty;
  syncMmsMenuQtyField();
  return idx >= 0 ? list[idx] : list[list.length - 1];
}

function removeMmsPlatoItemByKey(key) {
  const targetKey = String(key || "").trim();
  if (!targetKey) return;
  mmsLineItemsDraft = (Array.isArray(mmsLineItemsDraft) ? mmsLineItemsDraft : [])
    .filter((item) => String(item?.key || "") !== targetKey);
  mmsSelectedPlatoItems = (Array.isArray(mmsSelectedPlatoItems) ? mmsSelectedPlatoItems : [])
    .filter((item) => String(item?.key || "") !== targetKey);
  const nextActive = mmsSelectedPlatoItems[mmsSelectedPlatoItems.length - 1] || null;
  mmsActiveLineKey = String(nextActive?.key || "");
  mmsPlatoQty = Math.max(1, Math.floor(Number(nextActive?.qty || 1)));
  syncMmsMenuQtyField();
  if (nextActive) {
    loadMmsLineItemIntoEditor(nextActive.key);
  } else {
    clearMmsActiveComplements();
  }
}

function updateMmsPlatoItemQtyByKey(key, nextQty) {
  const targetKey = String(key || "").trim();
  const safeQty = Math.max(1, Math.floor(Number(nextQty || 1)));
  if (!targetKey) return false;
  let changed = false;
  mmsSelectedPlatoItems = (Array.isArray(mmsSelectedPlatoItems) ? mmsSelectedPlatoItems : []).map((item) => {
    if (String(item?.key || "") !== targetKey) return item;
    changed = true;
    return { ...item, qty: safeQty };
  });
  if (changed) {
    mmsPlatoQty = safeQty;
    syncMmsMenuQtyField();
  }
  return changed;
}

function clearMmsActiveComplements({ keepComment = false } = {}) {
  mmsSelectedSalsaIds = [];
  mmsSelectedBebidaIds = [];
  mmsGuarnicionQtyById = {};
  mmsPostreQtyById = {};
  mmsBebidaQtyById = {};
  setMmsSelectionSets({ guarnicionIds: [], postreIds: [], comentarioIds: [] });
  if (!keepComment) {
    if (el.mmsComentarioLibre) el.mmsComentarioLibre.value = "";
    if (el.mmsPlatoDescripcion) el.mmsPlatoDescripcion.value = "";
  }
}

function buildCurrentMmsLineItemSnapshot() {
  const activeKey = String(mmsActiveLineKey || "").trim();
  if (!activeKey) return null;
  const plate = (Array.isArray(mmsSelectedPlatoItems) ? mmsSelectedPlatoItems : []).find((item) => String(item?.key || "") === activeKey);
  if (!plate) return null;
  return {
    key: activeKey,
    platoId: Number(plate.platoId || 0) || null,
    preparacionId: Number(plate.preparacionId || 0) || null,
    qty: Math.max(1, Math.floor(Number(plate.qty || 1))),
    salsaIds: (Array.isArray(mmsSelectedSalsaIds) ? mmsSelectedSalsaIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0),
    guarnicionIds: getMmsSelectedGuarnicionIds(),
    guarnicionQtys: {},
    postreIds: getMmsSelectedPostreIds(),
    postreQtys: normalizeMmsPostreQtyMap(mmsPostreQtyById),
    bebidaIds: (Array.isArray(mmsSelectedBebidaIds) ? mmsSelectedBebidaIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0),
    bebidaQtys: normalizeMmsQtyMap(mmsBebidaQtyById),
    comentarioIds: selectedIdsFromChecklist(el.mmsComentariosAll),
    comentarioLibre: String(el.mmsComentarioLibre?.value || "").trim(),
  };
}

function commitCurrentMmsLineItem() {
  const snapshot = buildCurrentMmsLineItemSnapshot();
  if (!snapshot || !snapshot.platoId) return;
  const list = Array.isArray(mmsLineItemsDraft) ? [...mmsLineItemsDraft] : [];
  const idx = list.findIndex((item) => String(item?.key || "") === String(snapshot.key || ""));
  if (idx >= 0) list[idx] = snapshot;
  else list.push(snapshot);
  mmsLineItemsDraft = list;
}

function getCombinedMmsLineItems() {
  const list = Array.isArray(mmsLineItemsDraft) ? deepClone(mmsLineItemsDraft) : [];
  const snapshot = buildCurrentMmsLineItemSnapshot();
  if (!snapshot || !snapshot.platoId) return list;
  const idx = list.findIndex((item) => String(item?.key || "") === String(snapshot.key || ""));
  if (idx >= 0) list[idx] = snapshot;
  else list.push(snapshot);
  return list;
}

function loadMmsLineItemIntoEditor(key) {
  const targetKey = String(key || "").trim();
  if (!targetKey) {
    mmsActiveLineKey = "";
    clearMmsActiveComplements();
    refreshMmsDescriptionAuto();
    renderMmsStageOptions();
    renderMmsComandaPreview();
    return;
  }
  const line = getCombinedMmsLineItems().find((item) => String(item?.key || "") === targetKey);
  const plate = (Array.isArray(mmsSelectedPlatoItems) ? mmsSelectedPlatoItems : []).find((item) => String(item?.key || "") === targetKey);
  if (!line || !plate) return;
  mmsActiveLineKey = targetKey;
  mmsPlatoQty = Math.max(1, Math.floor(Number(plate.qty || 1)));
  if (el.mmsMenuQty) el.mmsMenuQty.value = String(mmsPlatoQty);
  if (el.mmsProtein) el.mmsProtein.value = String(plate.platoId || "");
  refreshMmsByProteinPreparation({ preserveSelection: false }).catch(() => { });
  mmsSelectedSalsaIds = (Array.isArray(line.salsaIds) ? line.salsaIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0);
  mmsSelectedBebidaIds = (Array.isArray(line.bebidaIds) ? line.bebidaIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0);
  mmsGuarnicionQtyById = {};
  mmsPostreQtyById = normalizeMmsQtyMap(line.postreQtys);
  mmsBebidaQtyById = normalizeMmsQtyMap(line.bebidaQtys);
  setMmsSelectionSets({
    guarnicionIds: Array.isArray(line.guarnicionIds) ? line.guarnicionIds : [],
    postreIds: Array.isArray(line.postreIds) ? line.postreIds : [],
    comentarioIds: Array.isArray(line.comentarioIds) ? line.comentarioIds : [],
  });
  if (el.mmsComentarioLibre) el.mmsComentarioLibre.value = String(line.comentarioLibre || "");
  if (el.mmsPlatoDescripcion) el.mmsPlatoDescripcion.value = String(line.comentarioLibre || "");
  syncMmsBebidaQtyWithSelection();
  refreshMmsDescriptionAuto();
  renderMmsStageOptions();
  renderMmsComandaPreview();
}

function setMmsItemQty(kind, id, qty) {
  const itemId = Number(id || 0);
  const safeQty = Math.max(1, Math.floor(Number(qty || 1)));
  if (!Number.isFinite(itemId) || itemId <= 0) return;
  if (kind === "guarnicion") {
    mmsGuarnicionQtyById[itemId] = safeQty;
  } else if (kind === "bebida") {
    mmsBebidaQtyById[itemId] = safeQty;
  } else if (kind === "postre") {
    mmsPostreQtyById[itemId] = safeQty;
  }
}

function deleteMmsItemQty(kind, id) {
  const itemId = Number(id || 0);
  if (!Number.isFinite(itemId) || itemId <= 0) return;
  if (kind === "guarnicion") delete mmsGuarnicionQtyById[itemId];
  else if (kind === "bebida") delete mmsBebidaQtyById[itemId];
  else if (kind === "postre") delete mmsPostreQtyById[itemId];
}

async function requestMmsItemQty(label, currentQty = 1, kindLabel = "producto") {
  const safeCurrent = Math.max(1, Math.floor(Number(currentQty || 1)));
  const itemLabel = String(label || "producto").trim() || "producto";
  if (window.Swal && typeof window.Swal.fire === "function") {
    const result = await window.Swal.fire({
      title: `Cantidad de ${kindLabel}`,
      html: `<div style="font-size:12.5px;opacity:.92">Selecciona la cantidad para <b>${escapeHtml(itemLabel)}</b>.</div>`,
      input: "number",
      inputValue: String(safeCurrent),
      inputAttributes: {
        min: "1",
        step: "1",
        autocapitalize: "off",
        autocorrect: "off",
      },
      showCancelButton: true,
      confirmButtonText: "Agregar",
      cancelButtonText: "Cancelar",
      background: "#f8fbff",
      color: "#10243b",
      confirmButtonColor: "#0f766e",
      cancelButtonColor: "#94a3b8",
      inputValidator: (value) => {
        const qty = Math.floor(Number(value || 0));
        if (!Number.isFinite(qty) || qty <= 0) return "Ingresa una cantidad valida mayor a 0.";
        return null;
      },
    });
    if (!result?.isConfirmed) return null;
    return Math.max(1, Math.floor(Number(result.value || 1)));
  }
  const raw = window.prompt(`Cantidad para "${itemLabel}":`, String(safeCurrent));
  if (raw === null) return null;
  const qty = Math.floor(Number(raw));
  if (!Number.isFinite(qty) || qty <= 0) {
    toast("Ingresa una cantidad valida mayor a 0.");
    return null;
  }
  return qty;
}

function flashMmsComandaTag(kind, id) {
  const container = el.mmsComandaPreview;
  if (!container) return;
  const tag = container.querySelector(`[data-mms-tag-kind="${String(kind || "").trim()}"][data-mms-tag-id="${String(id || "").trim()}"]`);
  if (!tag) return;
  tag.classList.remove("isFresh");
  void tag.offsetWidth;
  tag.classList.add("isFresh");
}

function notifyMmsSelectionAdded(kind, id, label, qty) {
  const kindMap = {
    plato: "Plato fuerte",
    salsa: "Salsa",
    guarnicion: "Guarnicion",
    postre: "Postre",
    bebida: "Bebida",
  };
  const guideMap = {
    plato: "Siguiente paso: agrega salsas o guarniciones para completar el plato.",
    salsa: "Siguiente paso: continua con guarniciones.",
    guarnicion: "Siguiente paso: agrega otra guarnicion o continua con postres.",
    postre: "Siguiente paso: define bebidas o guarda la comanda.",
    bebida: "Siguiente paso: revisa el resumen y guarda el menu.",
  };
  const kindLabel = kindMap[String(kind || "").trim()] || "Producto";
  const safeQty = Math.max(1, Math.floor(Number(qty || 1)));
  const suffix = kind === "guarnicion" || kind === "salsa" ? "." : ` x${safeQty}.`;
  const message = `${kindLabel} agregado: ${String(label || "").trim() || "Producto"}${suffix} ${guideMap[String(kind || "").trim()] || "Revisa el resumen de comanda."}`;
  toast(message);
  requestAnimationFrame(() => flashMmsComandaTag(kind, id));
}

function renderMmsQuickButtonsGroup(container, rows, selectedSet, kind) {
  if (!container) return;
  container.innerHTML = "";
  const list = Array.isArray(rows) ? rows : [];
  if (!list.length) {
    container.innerHTML = `<span class="muted">Sin opciones.</span>`;
    return;
  }
  for (const row of list) {
    const isBebida = kind === "bebida";
    const idRaw = isBebida ? String(row?.id || "").trim() : Number(row?.id || 0);
    if (isBebida) {
      if (!idRaw) continue;
    } else if (!Number.isFinite(idRaw) || idRaw <= 0) {
      continue;
    }
    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "mmsQuickBtn";
    if (selectedSet.has(idRaw)) btn.classList.add("isSelected");
    btn.dataset.mmsQuickKind = kind;
    btn.dataset.mmsQuickId = String(idRaw);
    btn.textContent = String(row?.nombre || "").trim() || "-";
    container.appendChild(btn);
  }
}

function renderMmsQuickSelectors() {
  ensureMmsCatalogDefaults();
  const cache = menuMontajeSelectableCatalogCache;
  const selectedGuarnicionSet = new Set(getMmsSelectedGuarnicionIds().map((x) => Number(x)));
  const selectedPostreSet = new Set(getMmsSelectedPostreIds().map((x) => Number(x)));

  const suggestedGuarnicionIds = new Set(listAllCheckboxIds(el.mmsGuarnicionesSuggested));
  const suggestedPostreIds = new Set(listAllCheckboxIds(el.mmsPostresSuggested));
  const suggestedGuarniciones = cache.guarniciones.filter((x) => suggestedGuarnicionIds.has(Number(x.id || 0)));
  const suggestedPostres = cache.postres.filter((x) => suggestedPostreIds.has(Number(x.id || 0)));

  const guarnicionFilter = String(el.mmsGuarnicionFilter?.value || "").trim().toLowerCase();
  const postreFilter = String(el.mmsPostreFilter?.value || "").trim().toLowerCase();
  const allGuarnicionesFiltered = cache.guarniciones.filter((x) => {
    const name = String(x?.nombre || "").trim().toLowerCase();
    return !guarnicionFilter || name.includes(guarnicionFilter);
  });
  const allPostresFiltered = cache.postres.filter((x) => {
    const name = String(x?.nombre || "").trim().toLowerCase();
    return !postreFilter || name.includes(postreFilter);
  });

  renderMmsQuickButtonsGroup(el.mmsGuarnicionesQuickSuggested, suggestedGuarniciones, selectedGuarnicionSet, "guarnicion");
  renderMmsQuickButtonsGroup(el.mmsPostresQuickSuggested, suggestedPostres, selectedPostreSet, "postre");
  renderMmsQuickButtonsGroup(el.mmsGuarnicionesQuickGlobal, allGuarnicionesFiltered, selectedGuarnicionSet, "guarnicion");
  renderMmsQuickButtonsGroup(el.mmsPostresQuickGlobal, allPostresFiltered, selectedPostreSet, "postre");

  if (el.mmsGuarnicionesQuickGlobal) el.mmsGuarnicionesQuickGlobal.hidden = !mmsShowAllGuarniciones;
  if (el.mmsPostresQuickGlobal) el.mmsPostresQuickGlobal.hidden = !mmsShowAllPostres;
}

async function toggleMmsQuickItem(kind, id) {
  const itemId = Number(id || 0);
  if (!Number.isFinite(itemId) || itemId <= 0) return;
  ensureMmsCatalogDefaults();
  const cache = menuMontajeSelectableCatalogCache;
  const guarniciones = new Set(getMmsSelectedGuarnicionIds().map((x) => Number(x)));
  const postres = new Set(getMmsSelectedPostreIds().map((x) => Number(x)));
  const comentarios = new Set(selectedIdsFromChecklist(el.mmsComentariosAll).map((x) => Number(x)));
  const adicionales = new Set(selectedIdsFromChecklist(el.mmsMontajeAdicionales).map((x) => Number(x)));
  if (kind === "guarnicion") {
    if (guarniciones.has(itemId)) {
      guarniciones.delete(itemId);
      deleteMmsItemQty("guarnicion", itemId);
    } else {
      guarniciones.add(itemId);
      const label = namesFromIds(cache.guarniciones, [itemId])[0] || "Guarnicion";
      notifyMmsSelectionAdded("guarnicion", itemId, label, 1);
    }
  } else if (kind === "postre") {
    if (postres.has(itemId)) {
      postres.delete(itemId);
      deleteMmsItemQty("postre", itemId);
    } else {
      const label = namesFromIds(cache.postres, [itemId])[0] || "Postre";
      const qty = await requestMmsItemQty(label, getMmsItemQty("postre", itemId), "postres");
      if (qty === null) return;
      postres.add(itemId);
      setMmsItemQty("postre", itemId, qty);
      notifyMmsSelectionAdded("postre", itemId, label, qty);
    }
  }
  setMmsSelectionSets({
    guarnicionIds: Array.from(guarniciones.values()),
    postreIds: Array.from(postres.values()),
    comentarioIds: Array.from(comentarios.values()),
    adicionalIds: Array.from(adicionales.values()),
  });
  refreshMmsDescriptionAuto();
  renderMmsSelectionSummary();
  renderMmsQuickSelectors();
  renderMmsComandaPreview();
}

function renderMmsComandaTag(container, label, removeKind, removeId, { qty = null, allowQty = false } = {}) {
  if (!container) return;
  const tag = document.createElement("span");
  tag.className = "mmsComandaTag";
  tag.dataset.mmsTagKind = String(removeKind || "");
  tag.dataset.mmsTagId = String(removeId || "");
  const safeQty = Math.max(1, Math.floor(Number(qty || 1)));
  const qtyHtml = allowQty && Number.isFinite(Number(qty))
    ? `<span class="mmsTagQty"><span class="mmsTagQtyText">Cant. ${safeQty}</span><button class="mmsTagIconBtn mmsTagQtyBtn isDec" type="button" data-mms-qty-kind="${escapeHtml(removeKind)}" data-mms-qty-action="dec" data-mms-qty-id="${escapeHtml(String(removeId))}" title="Disminuir" aria-label="Disminuir cantidad"></button><button class="mmsTagIconBtn mmsTagQtyBtn isInc" type="button" data-mms-qty-kind="${escapeHtml(removeKind)}" data-mms-qty-action="inc" data-mms-qty-id="${escapeHtml(String(removeId))}" title="Aumentar" aria-label="Aumentar cantidad"></button></span>`
    : "";
  tag.innerHTML = `<span class="mmsComandaTagText">${escapeHtml(label)}</span>${qtyHtml}<button class="mmsTagIconBtn mmsTagRemoveBtn" type="button" data-mms-remove-kind="${escapeHtml(removeKind)}" data-mms-remove-id="${escapeHtml(String(removeId))}" title="Quitar" aria-label="Quitar del resumen"></button>`;
  container.appendChild(tag);
  return tag;
}

function renderMmsComandaPreview() {
  ensureMmsCatalogDefaults();
  const cache = menuMontajeSelectableCatalogCache;
  if (el.mmsActivePlateHint) {
    const activeItem = (Array.isArray(mmsSelectedPlatoItems) ? mmsSelectedPlatoItems : []).find((item) => String(item?.key || "") === String(mmsActiveLineKey || ""));
    if (!activeItem) {
      el.mmsActivePlateHint.textContent = "Toca un plato para editarlo.";
    } else {
      const index = (Array.isArray(mmsSelectedPlatoItems) ? mmsSelectedPlatoItems : []).findIndex((item) => String(item?.key || "") === String(mmsActiveLineKey || ""));
      const plateNo = index >= 0 ? index + 1 : 1;
      el.mmsActivePlateHint.textContent = `Editando Plato ${plateNo}: ${getMmsPlatoItemLabel(activeItem)}. Toca otro plato para cambiar.`;
    }
  }
  if (el.mmsComandaPlato) {
    el.mmsComandaPlato.innerHTML = "";
    const items = Array.isArray(mmsSelectedPlatoItems) ? mmsSelectedPlatoItems : [];
    if (!items.length) {
      const plain = document.createElement("span");
      plain.className = "mmsComandaTag";
      plain.textContent = "(sin plato fuerte)";
      el.mmsComandaPlato.appendChild(plain);
    } else {
      for (let i = 0; i < items.length; i++) {
        const item = items[i];
        const label = `Plato ${i + 1} | Cant ${Math.max(1, Math.floor(Number(item?.qty || 1)))} | ${getMmsPlatoItemLabel(item)}`;
        const tag = renderMmsComandaTag(el.mmsComandaPlato, label, "plato", item.key, {
          qty: Math.max(1, Math.floor(Number(item?.qty || 1))),
          allowQty: false,
        });
        tag?.classList.add("isPlateSelector");
        if (String(item?.key || "") === String(mmsActiveLineKey || "")) tag?.classList.add("isActiveLine");
      }
    }
  }

  if (el.mmsComandaGuarniciones) {
    el.mmsComandaGuarniciones.innerHTML = "";
    const ids = getMmsSelectedGuarnicionIds();
    const names = namesFromIds(cache.guarniciones, ids);
    if (!names.length) {
      const empty = document.createElement("span");
      empty.className = "muted";
      empty.textContent = "Sin guarniciones";
      el.mmsComandaGuarniciones.appendChild(empty);
    } else {
      for (let i = 0; i < names.length; i++) {
        renderMmsComandaTag(el.mmsComandaGuarniciones, names[i], "guarnicion", ids[i]);
      }
    }
  }

  if (el.mmsComandaSalsas) {
    el.mmsComandaSalsas.innerHTML = "";
    const ids = (Array.isArray(mmsSelectedSalsaIds) ? mmsSelectedSalsaIds : []).map((x) => Number(x));
    const names = namesFromIds(cache.salsas, ids);
    if (!names.length) {
      const empty = document.createElement("span");
      empty.className = "muted";
      empty.textContent = "Sin salsas";
      el.mmsComandaSalsas.appendChild(empty);
    } else {
      for (let i = 0; i < names.length; i++) renderMmsComandaTag(el.mmsComandaSalsas, names[i], "salsa", ids[i]);
    }
  }

  if (el.mmsComandaPostres) {
    el.mmsComandaPostres.innerHTML = "";
    const ids = getMmsSelectedPostreIds();
    if (!ids.length) {
      const empty = document.createElement("span");
      empty.className = "muted";
      empty.textContent = "Sin postres";
      el.mmsComandaPostres.appendChild(empty);
    } else {
      const map = new Map((cache.postres || []).map((p) => [Number(p.id), String(p.nombre || "").trim()]));
      for (const id of ids) {
        const name = String(map.get(Number(id)) || "").trim();
        if (!name) continue;
        renderMmsComandaTag(el.mmsComandaPostres, name, "postre", id, {
          qty: getMmsItemQty("postre", id),
          allowQty: true,
        });
      }
    }
  }

  if (el.mmsComandaBebidas) {
    el.mmsComandaBebidas.innerHTML = "";
    const ids = (Array.isArray(mmsSelectedBebidaIds) ? mmsSelectedBebidaIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0);
    const names = namesFromIds(cache.bebidas, ids);
    if (!names.length) {
      const empty = document.createElement("span");
      empty.className = "muted";
      empty.textContent = "Sin bebidas";
      el.mmsComandaBebidas.appendChild(empty);
    } else {
      for (let i = 0; i < names.length; i++) {
        renderMmsComandaTag(el.mmsComandaBebidas, names[i], "bebida", ids[i], {
          qty: getMmsItemQty("bebida", ids[i]),
          allowQty: true,
        });
      }
    }
  }
  if (el.mmsComandaComentarios) {
    el.mmsComandaComentarios.innerHTML = "";
    const commentIds = selectedIdsFromChecklist(el.mmsComentariosAll);
    const commentNames = namesFromIds(cache.comentarios, commentIds);
    const commentFree = String(el.mmsComentarioLibre?.value || "").trim();
    const allComments = [...commentNames, ...(commentFree ? [commentFree] : [])];
    if (!allComments.length) {
      const empty = document.createElement("span");
      empty.className = "muted";
      empty.textContent = "Sin comentarios";
      el.mmsComandaComentarios.appendChild(empty);
    } else {
      for (let i = 0; i < allComments.length; i++) {
        const tag = document.createElement("span");
        tag.className = "mmsComandaTag";
        tag.textContent = allComments[i];
        el.mmsComandaComentarios.appendChild(tag);
      }
    }
  }
  if (el.mmsComandaMontaje) {
    el.mmsComandaMontaje.innerHTML = "";
    const tipo = String(el.mmsMontajeTipo?.selectedOptions?.[0]?.textContent || "").trim();
    const adicionales = namesFromIds(cache.montajeAdicionales, selectedIdsFromChecklist(el.mmsMontajeAdicionales));
    if (tipo) {
      const tag = document.createElement("span");
      tag.className = "mmsComandaTag";
      tag.textContent = `Tipo: ${tipo}`;
      el.mmsComandaMontaje.appendChild(tag);
    }
    if (adicionales.length) {
      for (const a of adicionales) {
        const tag = document.createElement("span");
        tag.className = "mmsComandaTag";
        tag.textContent = a;
        el.mmsComandaMontaje.appendChild(tag);
      }
    }
    if (!tipo && !adicionales.length) {
      const empty = document.createElement("span");
      empty.className = "muted";
      empty.textContent = "Sin montaje";
      el.mmsComandaMontaje.appendChild(empty);
    }
  }
}

function removeMmsComandaItem(kind, id) {
  if (kind === "plato") {
    removeMmsPlatoItemByKey(id);
    refreshMmsDescriptionAuto();
    renderMmsSelectionSummary();
    renderMmsStageOptions();
    renderMmsComandaPreview();
    return;
  }
  const itemId = Number(id || 0);
  if (!Number.isFinite(itemId) || itemId <= 0) return;
  const guarniciones = new Set(getMmsSelectedGuarnicionIds().map((x) => Number(x)));
  const postres = new Set(getMmsSelectedPostreIds().map((x) => Number(x)));
  const salsas = new Set((Array.isArray(mmsSelectedSalsaIds) ? mmsSelectedSalsaIds : []).map((x) => Number(x)));
  const bebidas = new Set((Array.isArray(mmsSelectedBebidaIds) ? mmsSelectedBebidaIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0));
  if (kind === "guarnicion") {
    guarniciones.delete(itemId);
    deleteMmsItemQty("guarnicion", itemId);
  }
  if (kind === "postre") {
    postres.delete(itemId);
    deleteMmsItemQty("postre", itemId);
  }
  if (kind === "salsa") salsas.delete(itemId);
  if (kind === "bebida") {
    bebidas.delete(Number(id || 0));
    deleteMmsItemQty("bebida", itemId);
  }
  setMmsSelectionSets({
    guarnicionIds: Array.from(guarniciones.values()),
    postreIds: Array.from(postres.values()),
    comentarioIds: selectedIdsFromChecklist(el.mmsComentariosAll),
    adicionalIds: selectedIdsFromChecklist(el.mmsMontajeAdicionales),
  });
  mmsSelectedSalsaIds = Array.from(salsas.values()).filter((x) => Number.isFinite(x) && x > 0);
  mmsSelectedBebidaIds = Array.from(bebidas.values()).filter((n) => Number.isFinite(n) && n > 0);
  syncMmsBebidaQtyWithSelection();
  refreshMmsDescriptionAuto();
  renderMmsSelectionSummary();
  renderMmsQuickSelectors();
  renderMmsStageOptions();
  renderMmsComandaPreview();
}

function renderMmsSelectionSummary() {
  ensureMmsCatalogDefaults();
  const cache = menuMontajeSelectableCatalogCache;
  const platoItems = (Array.isArray(mmsSelectedPlatoItems) ? mmsSelectedPlatoItems : [])
    .map((item) => `${getMmsPlatoItemLabel(item)} | Cant: ${Math.max(1, Math.floor(Number(item?.qty || 1)))}`);
  const guarnicionIds = selectedIdsUnionFromTwoLists(el.mmsGuarnicionesSuggested, el.mmsGuarnicionesAll);
  const postreIds = selectedIdsUnionFromTwoLists(el.mmsPostresSuggested, el.mmsPostresAll);
  const guarniciones = guarnicionIds.map((id) => {
    const name = namesFromIds(cache.guarniciones, [id])[0] || "";
    return name ? `${name} x${getMmsItemQty("guarnicion", id)}` : "";
  }).filter(Boolean);
  const postres = postreIds.map((id) => {
    const name = namesFromIds(cache.postres, [id])[0] || "";
    return name ? `${name} x${getMmsItemQty("postre", id)}` : "";
  }).filter(Boolean);
  const comentarios = namesFromIds(cache.comentarios, selectedIdsFromChecklist(el.mmsComentariosAll));
  const comentarioLibre = String(el.mmsComentarioLibre?.value || "").trim();
  const montajeTipo = String(el.mmsMontajeTipo?.selectedOptions?.[0]?.textContent || "").trim();
  const montajeAdicionales = namesFromIds(cache.montajeAdicionales, selectedIdsFromChecklist(el.mmsMontajeAdicionales));

  if (el.mmsSummaryMenu) {
    el.mmsSummaryMenu.value = platoItems.length ? platoItems.join(", ") : "(sin seleccionar)";
  }
  if (el.mmsSummaryGuarniciones) el.mmsSummaryGuarniciones.value = guarniciones.length ? guarniciones.join(", ") : "(sin seleccionar)";
  if (el.mmsSummaryPostres) el.mmsSummaryPostres.value = postres.length ? postres.join(", ") : "(sin seleccionar)";
  if (el.mmsSummaryComentarios) {
    const all = [...comentarios, ...(comentarioLibre ? [comentarioLibre] : [])];
    el.mmsSummaryComentarios.value = all.length ? all.join(", ") : "(sin comentarios)";
  }
  if (el.mmsSummaryMontajeTipo) el.mmsSummaryMontajeTipo.value = montajeTipo || "(sin seleccionar)";
  if (el.mmsSummaryMontajeAdicionales) el.mmsSummaryMontajeAdicionales.value = montajeAdicionales.length ? montajeAdicionales.join(", ") : "(sin seleccionar)";
}

function appendTextBlock(textarea, blockText) {
  if (!(textarea instanceof HTMLTextAreaElement)) return;
  const current = String(textarea.value || "").trim();
  const incoming = String(blockText || "").trim();
  if (!incoming) return;
  textarea.value = current ? `${current}\n\n${incoming}` : incoming;
}

function applyMmsMenuBuilder(mode = "append") {
  const block = buildMmsMenuDescriptionFromForm();
  if (!String(block || "").trim()) return;
  if (mode === "replace") {
    if (el.mmsMenuDescription) el.mmsMenuDescription.value = block;
  } else {
    appendTextBlock(el.mmsMenuDescription, block);
  }
  maybeAutofillMmsTitle();
  renderMmsSelectionSummary();
}

function applyMmsMontajeBuilder(mode = "append") {
  const block = buildMmsMontajeDescriptionFromForm();
  if (!String(block || "").trim()) return;
  if (mode === "replace") {
    if (el.mmsMontajeDescription) el.mmsMontajeDescription.value = block;
  } else {
    appendTextBlock(el.mmsMontajeDescription, block);
  }
  renderMmsSelectionSummary();
}

function useMmsSuggestedSelections() {
  const guarnicionIds = listAllCheckboxIds(el.mmsGuarnicionesSuggested);
  const postreIds = listAllCheckboxIds(el.mmsPostresSuggested);
  const salsaIds = getMmsSuggestedSalsaIds();
  if (!guarnicionIds.length && !postreIds.length && !salsaIds.length) {
    return toast("No hay sugerencias para esta combinacion.");
  }
  mmsSelectedSalsaIds = salsaIds;
  setMmsSelectionSets({ guarnicionIds, postreIds });
  refreshMmsDescriptionAuto();
  renderMmsSelectionSummary();
  renderMmsStageOptions();
  toast("Sugerencias aplicadas.");
}

function clearMmsMenuSelections() {
  mmsLineItemsDraft = [];
  mmsActiveLineKey = "";
  mmsSelectedPlatoItems = [];
  mmsPlatoQty = 1;
  syncMmsMenuQtyField();
  if (el.mmsProtein) el.mmsProtein.value = "";
  if (el.mmsPreparation) el.mmsPreparation.value = "";
  setMmsSelectionSets({ guarnicionIds: [], postreIds: [], comentarioIds: [] });
  mmsSelectedSalsaIds = [];
  mmsSelectedBebidaIds = [];
  mmsGuarnicionQtyById = {};
  mmsPostreQtyById = {};
  mmsBebidaQtyById = {};
  if (el.mmsComentarioLibre) el.mmsComentarioLibre.value = "";
  if (el.mmsPlatoDescripcion) el.mmsPlatoDescripcion.value = "";
  refreshMmsDescriptionAuto();
  renderMmsQuickSelectors();
  renderMmsStageOptions();
  renderMmsSelectionSummary();
}

function clearMmsMontajeSelections() {
  setMmsSelectionSets({ adicionalIds: [] });
  if (el.mmsMontajeTipo && el.mmsMontajeTipo.options.length) {
    el.mmsMontajeTipo.value = el.mmsMontajeTipo.options[0].value;
  }
  refreshMmsDescriptionAuto();
  renderMmsQuickSelectors();
  renderMmsSelectionSummary();
}

function renderMmsEntriesTable() {
  if (!el.mmsEntriesBody || !quoteDraft) return;
  const entries = ensureMenuMontajeDraft()
    .slice()
    .sort((a, b) => {
      const d = String(a.date || "").localeCompare(String(b.date || ""));
      if (d !== 0) return d;
      return String(a.salon || "").localeCompare(String(b.salon || ""));
    });
  el.mmsEntriesBody.innerHTML = "";
  if (!entries.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="5">Sin informacion de menu/montaje.</td>`;
    el.mmsEntriesBody.appendChild(tr);
    return;
  }
  for (const item of entries) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(String(item.date || "-"))}</td>
      <td>${escapeHtml(String(item.salon || "-"))}</td>
      <td>${escapeHtml(String(item.menuTitle || "-"))}</td>
      <td>${escapeHtml(String(item.menuQty || "-"))}</td>
      <td>${escapeHtml(String(item.montajeDescription || "-").slice(0, 120))}</td>
    `;
    el.mmsEntriesBody.appendChild(tr);
  }
}

function renderMmsVersionControls() {
  if (!el.mmsVersionSelect || !quoteDraft) return;
  const model = ensureMenuMontajeModel();
  const previous = Number(menuMontajeSelectableSelectedVersion || model.currentVersion);
  el.mmsVersionSelect.innerHTML = "";
  for (const v of model.versions.slice().sort((a, b) => Number(b.version || 0) - Number(a.version || 0))) {
    const opt = document.createElement("option");
    const stamp = formatQuoteSentAtLabel(v.savedAt);
    opt.value = String(v.version);
    opt.textContent = stamp ? `V${v.version} - ${stamp}` : `V${v.version}`;
    el.mmsVersionSelect.appendChild(opt);
  }
  const selected = model.versions.some((v) => Number(v.version) === previous)
    ? previous
    : model.currentVersion;
  menuMontajeSelectableSelectedVersion = selected;
  el.mmsVersionSelect.value = String(selected);
}

function renderMmsDateSalonSelect() {
  if (!el.mmsDateSalonSelect || !quoteDraft) return;
  ensureMenuMontajeModel();
  const combos = getQuoteDateSalonCombos();
  const previous = String(menuMontajeSelectableSelectedKey || "").trim();
  el.mmsDateSalonSelect.innerHTML = "";
  for (const c of combos) {
    const opt = document.createElement("option");
    opt.value = c.key;
    opt.textContent = `${c.date} - ${c.salon}`;
    el.mmsDateSalonSelect.appendChild(opt);
  }
  if (previous && combos.some((x) => x.key === previous)) {
    menuMontajeSelectableSelectedKey = previous;
  } else {
    menuMontajeSelectableSelectedKey = combos[0]?.key || "";
  }
  el.mmsDateSalonSelect.value = menuMontajeSelectableSelectedKey;
}

function renderMmsMontajeCatalogs() {
  ensureMmsCatalogDefaults();
  const cache = menuMontajeSelectableCatalogCache;
  if (el.mmsMontajeTipo) {
    el.mmsMontajeTipo.innerHTML = "";
    for (const row of cache.montajeTipos) {
      const opt = document.createElement("option");
      opt.value = String(row.id);
      opt.textContent = String(row.nombre || "");
      el.mmsMontajeTipo.appendChild(opt);
    }
    if (!el.mmsMontajeTipo.options.length) {
      el.mmsMontajeTipo.innerHTML = `<option value="">Sin tipos de montaje</option>`;
    }
  }
  renderMenuSuggestionCheckboxList(el.mmsMontajeAdicionales, cache.montajeAdicionales, []);
  renderMenuSuggestionCheckboxList(el.mmsComentariosAll, cache.comentarios, []);
}

function renderMmsProteinOptions() {
  ensureMmsCatalogDefaults();
  const cache = menuMontajeSelectableCatalogCache;
  if (!el.mmsProtein) return;
  const prev = Number(el.mmsProtein.value || 0);
  el.mmsProtein.innerHTML = "";
  for (const row of cache.proteins) {
    const opt = document.createElement("option");
    opt.value = String(row.id);
    opt.textContent = formatPlatoCatalogLabel(row);
    el.mmsProtein.appendChild(opt);
  }
  const hasPrev = prev > 0 && cache.proteins.some((x) => Number(x.id) === prev);
  if (hasPrev) el.mmsProtein.value = String(prev);
  if (!el.mmsProtein.options.length) {
    el.mmsProtein.innerHTML = `<option value="">Sin proteinas activas</option>`;
  }
}

function buildMmsMenuDescriptionFromForm() {
  ensureMmsCatalogDefaults();
  const cache = menuMontajeSelectableCatalogCache;
  const lineItems = getCombinedMmsLineItems();
  const mapNames = (rows, ids) => namesFromIds(rows, ids);
  const mapItemsWithQtyFromMap = (rows, ids, qtyMap) => ids.map((itemId) => {
    const name = namesFromIds(rows, [itemId])[0] || "";
    const qty = Math.max(1, Math.floor(Number(qtyMap?.[itemId] || 1)));
    return name ? `${name} | Cant: ${qty}` : "";
  }).filter(Boolean);
  const lines = [];
  if (!lineItems.length) {
    lines.push("[PLATOS FUERTES]");
    lines.push("- Por definir");
    return lines.join("\n").trim();
  }
  lineItems.forEach((line, index) => {
    if (index > 0) lines.push("");
    lines.push(`[PLATO ${index + 1}]`);
    const platoFuerte = getMmsPlatoBaseName(line) || "Por definir";
    const preparacion = getMmsPreparationName(line) || "Por definir";
    const salsaNames = mapNames(cache.salsas, line.salsaIds || []);
    const guarniciones = mapNames(cache.guarniciones, line.guarnicionIds || []);
    const postres = mapItemsWithQtyFromMap(cache.postres, line.postreIds || [], line.postreQtys);
    const bebidas = mapItemsWithQtyFromMap(cache.bebidas, line.bebidaIds || [], line.bebidaQtys);
    const compactParts = [
      `PLATO FUERTE (Cant ${Math.max(1, Math.floor(Number(line?.qty || 1)))} - ${platoFuerte})`,
      `PREPARACION (${preparacion})`,
      `SALSAS (${salsaNames.length ? salsaNames.join(", ") : "Por definir"})`,
      `GUARNICIONES (${guarniciones.length ? guarniciones.join(", ") : "Por definir"})`,
      `POSTRES (${postres.length ? postres.join(", ") : "Por definir"})`,
      `BEBIDAS (${bebidas.length ? bebidas.join(", ") : "Por definir"})`,
    ];
    const comentarios = mapNames(cache.comentarios, line.comentarioIds || []);
    const comentarioLibre = String(line.comentarioLibre || "").trim();
    if (comentarios.length || comentarioLibre) {
      const commentParts = [...comentarios];
      if (comentarioLibre) commentParts.push(comentarioLibre);
      compactParts.push(`COMENTARIOS (${commentParts.join(", ")})`);
    }
    lines.push(`- ${compactParts.join(" | ")}`);
  });
  if (!lines.length) {
    lines.push("[PLATOS FUERTES]");
    lines.push("- Por definir");
  }
  return lines.join("\n").trim();
}

function buildMmsMontajeDescriptionFromForm() {
  ensureMmsCatalogDefaults();
  const cache = menuMontajeSelectableCatalogCache;
  const tipoId = Number(el.mmsMontajeTipo?.value || 0);
  const adicionalesIds = selectedIdsFromChecklist(el.mmsMontajeAdicionales);
  const tipoMap = buildMmsNameMap(cache.montajeTipos);
  const tipoName = String(tipoMap.get(tipoId) || "").trim();
  const adicionales = namesFromIds(cache.montajeAdicionales, adicionalesIds);
  const manual = String(el.mmsMontajeDescription?.value || "").trim();
  const lines = [];
  lines.push("[MONTAJE]");
  lines.push(`- TIPO (${tipoName || "Por definir"}) | ADICIONALES (${adicionales.length ? adicionales.join(", ") : "Ninguno"})`);
  if (manual) {
    lines.push("");
    lines.push("[DETALLE]");
    lines.push(manual);
  }
  return lines.join("\n").trim();
}

function maybeAutofillMmsTitle() {
  return;
}

function refreshMmsDescriptionAuto() {
  if (menuMontajeSelectableSilentUpdate) return;
  const autoMenu = buildMmsMenuDescriptionFromForm();
  if (el.mmsMenuDescription) el.mmsMenuDescription.value = autoMenu;
  if (el.mmsMontajeDescription && !String(el.mmsMontajeDescription.value || "").trim()) {
    el.mmsMontajeDescription.value = buildMmsMontajeDescriptionFromForm();
  }
  renderMmsSelectionSummary();
  renderMmsComandaPreview();
}

async function refreshMmsByProteinPreparation({ preserveSelection = true } = {}) {
  ensureMmsCatalogDefaults();
  const cache = menuMontajeSelectableCatalogCache;
  const platoId = Number(el.mmsProtein?.value || 0);
  await populateMmsPreparationsForProtein(platoId, { preserveSelection });
  const prepId = Number(el.mmsPreparation?.value || 0);
  let links = normalizeMmsSuggestionLinks(null);
  if (platoId > 0 && prepId > 0) {
    try {
      links = await readMenuSuggestionsCached(platoId, prepId);
    } catch (_) { }
  }

  const selectedGuarniciones = selectedIdsUnionFromTwoLists(el.mmsGuarnicionesSuggested, el.mmsGuarnicionesAll);
  const selectedPostres = selectedIdsUnionFromTwoLists(el.mmsPostresSuggested, el.mmsPostresAll);
  const selectedBebidas = (Array.isArray(mmsSelectedBebidaIds) ? mmsSelectedBebidaIds : [])
    .map((x) => Number(x))
    .filter((n) => Number.isFinite(n) && n > 0);
  const selectedMontajeAdicionales = selectedIdsFromChecklist(el.mmsMontajeAdicionales)
    .map((x) => Number(x))
    .filter((n) => Number.isFinite(n) && n > 0);

  cache.suggestedSalsaIds = Array.isArray(links.salsaIds) ? links.salsaIds : [];

  const guarnicionSet = new Set([...(Array.isArray(links.guarnicionIds) ? links.guarnicionIds : []), ...selectedGuarniciones].map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0));
  const postreSet = new Set([...(Array.isArray(links.postreIds) ? links.postreIds : []), ...selectedPostres].map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0));
  const bebidaSet = new Set([...(Array.isArray(links.bebidaIds) ? links.bebidaIds : []), ...selectedBebidas].map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0));
  const montajeAdicionalSet = new Set([...(Array.isArray(links.montajeAdicionalIds) ? links.montajeAdicionalIds : []), ...selectedMontajeAdicionales].map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0));

  const guarnicionIds = Array.from(guarnicionSet.values());
  const postreIds = Array.from(postreSet.values());
  const bebidaIds = Array.from(bebidaSet.values());

  const guarnicionesSuggested = cache.guarniciones.filter((x) => guarnicionSet.has(Number(x.id || 0)));
  const postresSuggested = cache.postres.filter((x) => postreSet.has(Number(x.id || 0)));
  renderMenuSuggestionCheckboxList(el.mmsGuarnicionesSuggested, guarnicionesSuggested, guarnicionIds);
  renderMenuSuggestionCheckboxList(el.mmsGuarnicionesAll, cache.guarniciones, guarnicionIds);
  renderMenuSuggestionCheckboxList(el.mmsPostresSuggested, postresSuggested, postreIds);
  renderMenuSuggestionCheckboxList(el.mmsPostresAll, cache.postres, postreIds);

  mmsSelectedBebidaIds = bebidaIds;
  mmsBebidaQtyById = syncMmsQtyMapWithSelection(mmsSelectedBebidaIds, mmsBebidaQtyById);
  setChecklistCheckedByIds(el.mmsMontajeAdicionales, montajeAdicionalSet);

  const suggestedMontajeTipoId = Array.isArray(links.montajeTipoIds)
    ? Number(links.montajeTipoIds[0] || 0)
    : 0;
  if (suggestedMontajeTipoId > 0 && el.mmsMontajeTipo && Array.from(el.mmsMontajeTipo.options).some((o) => Number(o.value || 0) === suggestedMontajeTipoId)) {
    const currentMontajeTipoId = Number(el.mmsMontajeTipo.value || 0);
    if (!currentMontajeTipoId || !preserveSelection) {
      el.mmsMontajeTipo.value = String(suggestedMontajeTipoId);
    }
  }

  syncMmsGuarnicionQtyWithSelection();
  syncMmsPostreQtyWithSelection();
  syncMmsBebidaQtyWithSelection();
  renderMmsQuickSelectors();
  renderMmsStageOptions();
  maybeAutofillMmsTitle();
  refreshMmsDescriptionAuto();
  renderMmsSelectionSummary();
  renderMmsComandaPreview();
}

async function populateMmsPreparationsForProtein(platoIdRaw, { preserveSelection = true } = {}) {
  const platoId = Number(platoIdRaw || 0);
  const previousPrep = Number(el.mmsPreparation?.value || 0);
  const preps = await loadMmsPreparationsByProtein(platoId);
  if (!el.mmsPreparation) return preps;
  el.mmsPreparation.innerHTML = "";
  for (const p of preps) {
    const opt = document.createElement("option");
    opt.value = String(p.id);
    opt.textContent = String(p.nombre || "");
    el.mmsPreparation.appendChild(opt);
  }
  const keepPrep = preserveSelection && previousPrep > 0 && preps.some((x) => Number(x.id) === previousPrep);
  if (keepPrep) el.mmsPreparation.value = String(previousPrep);
  if (!el.mmsPreparation.options.length) {
    el.mmsPreparation.innerHTML = `<option value="">Sin preparaciones para esta proteina</option>`;
  }
  return preps;
}

async function loadMmsFormByKey(key) {
  if (!quoteDraft) return;
  const entries = ensureMenuMontajeDraft();
  const [date, salon] = String(key || "").split("|");
  const found = entries.find((x) => String(x.date || "") === String(date || "") && String(x.salon || "") === String(salon || ""));
  const menuSelection = found?.menuSelection || {};
  const montajeSelection = found?.montajeSelection || {};
  renderMmsMenuSectionSelect(String(menuSelection?.section || "General"));

  menuMontajeSelectableSilentUpdate = true;
  if (el.mmsMenuTitle) el.mmsMenuTitle.value = String(found?.menuTitle || "");
  mmsPlatoQty = Math.max(1, Math.floor(Number(menuSelection?.platoQty || found?.menuQty || 1)));
  if (el.mmsMenuQty) el.mmsMenuQty.value = String(mmsPlatoQty);
  if (el.mmsComentarioLibre) el.mmsComentarioLibre.value = String(menuSelection?.comentarioLibre || "");
  if (el.mmsPlatoDescripcion) el.mmsPlatoDescripcion.value = String(menuSelection?.comentarioLibre || "");
  if (el.mmsMenuDescription) el.mmsMenuDescription.value = String(found?.menuDescription || "");
  if (el.mmsMontajeDescription) el.mmsMontajeDescription.value = String(found?.montajeDescription || "");
  menuMontajeSelectableSilentUpdate = false;

  const lineItems = normalizeMmsLineItems(menuSelection?.lineItems, menuSelection);
  mmsLineItemsDraft = deepClone(lineItems);
  mmsSelectedPlatoItems = lineItems.length
    ? lineItems.map((item) => ({
      key: String(item.key || `pl_${uid()}`),
      platoId: Number(item.platoId || 0),
      preparacionId: Number(item.preparacionId || 0) || null,
      qty: Math.max(1, Math.floor(Number(item.qty || 1))),
    }))
    : normalizeMmsPlatoItems(menuSelection?.platoItems, menuSelection);
  syncMmsMenuQtyField();

  const firstPlato = mmsSelectedPlatoItems[0] || null;
  const pid = Number(firstPlato?.platoId || menuSelection?.platoId || 0);
  if (pid > 0 && el.mmsProtein && Array.from(el.mmsProtein.options).some((o) => Number(o.value || 0) === pid)) {
    el.mmsProtein.value = String(pid);
  }
  await refreshMmsByProteinPreparation({ preserveSelection: false });

  const prepId = Number(firstPlato?.preparacionId || menuSelection?.preparacionId || 0);
  if (prepId > 0 && el.mmsPreparation && Array.from(el.mmsPreparation.options).some((o) => Number(o.value || 0) === prepId)) {
    el.mmsPreparation.value = String(prepId);
    await refreshMmsByProteinPreparation({ preserveSelection: true });
  }
  const activeLine = lineItems[lineItems.length - 1] || null;
  mmsActiveLineKey = String(activeLine?.key || "");
  const mergedSalsaIds = Array.isArray(activeLine?.salsaIds) ? activeLine.salsaIds : (Array.isArray(menuSelection?.salsaIds) ? menuSelection.salsaIds : []);
  const mergedBebidaIds = Array.isArray(activeLine?.bebidaIds) ? activeLine.bebidaIds : (Array.isArray(menuSelection?.bebidaIds) ? menuSelection.bebidaIds : []);
  const mergedGuarnicionIds = Array.isArray(activeLine?.guarnicionIds) ? activeLine.guarnicionIds : (Array.isArray(menuSelection?.guarnicionIds) ? menuSelection.guarnicionIds : []);
  const mergedPostreIds = Array.isArray(activeLine?.postreIds) ? activeLine.postreIds : (Array.isArray(menuSelection?.postreIds) ? menuSelection.postreIds : []);
  const mergedComentarioIds = Array.isArray(activeLine?.comentarioIds) ? activeLine.comentarioIds : (Array.isArray(menuSelection?.comentarioIds) ? menuSelection.comentarioIds : []);
  const mergedPostreQtys = activeLine ? normalizeMmsQtyMap(activeLine.postreQtys) : normalizeMmsPostreQtyMap(menuSelection?.postreQtys);
  const mergedBebidaQtys = activeLine ? normalizeMmsQtyMap(activeLine.bebidaQtys) : normalizeMmsQtyMap(menuSelection?.bebidaQtys);
  const mergedComentarioLibre = activeLine ? String(activeLine.comentarioLibre || "").trim() : String(menuSelection?.comentarioLibre || "").trim();
  mmsSelectedSalsaIds = mergedSalsaIds.map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0);
  mmsSelectedBebidaIds = mergedBebidaIds.map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0);
  mmsGuarnicionQtyById = {};
  mmsPostreQtyById = mergedPostreQtys;
  mmsBebidaQtyById = mergedBebidaQtys;
  if (el.mmsComentarioLibre) el.mmsComentarioLibre.value = mergedComentarioLibre;
  if (el.mmsPlatoDescripcion) el.mmsPlatoDescripcion.value = mergedComentarioLibre;

  const guarnicionSet = new Set(mergedGuarnicionIds.map((x) => Number(x)));
  const postreSet = new Set(mergedPostreIds.map((x) => Number(x)));
  const comentarioSet = new Set(mergedComentarioIds.map((x) => Number(x)));
  const adicionalSet = new Set((Array.isArray(montajeSelection?.adicionalIds) ? montajeSelection.adicionalIds : []).map((x) => Number(x)));
  setChecklistCheckedByIds(el.mmsGuarnicionesSuggested, guarnicionSet);
  setChecklistCheckedByIds(el.mmsGuarnicionesAll, guarnicionSet);
  setChecklistCheckedByIds(el.mmsPostresSuggested, postreSet);
  setChecklistCheckedByIds(el.mmsPostresAll, postreSet);
  syncMmsGuarnicionQtyWithSelection();
  syncMmsPostreQtyWithSelection();
  syncMmsBebidaQtyWithSelection();
  setChecklistCheckedByIds(el.mmsComentariosAll, comentarioSet);
  setChecklistCheckedByIds(el.mmsMontajeAdicionales, adicionalSet);

  const montajeTipoId = Number(montajeSelection?.tipoId || 0);
  if (montajeTipoId > 0 && el.mmsMontajeTipo && Array.from(el.mmsMontajeTipo.options).some((o) => Number(o.value || 0) === montajeTipoId)) {
    el.mmsMontajeTipo.value = String(montajeTipoId);
  }
  if (el.mmsDocNo) el.mmsDocNo.value = String(quoteDraft.code || "").trim() || "(sin codigo)";
  refreshMmsDescriptionAuto();
  renderMmsQuickSelectors();
  renderMmsStageOptions();
  renderMmsSelectionSummary();
  renderMmsComandaPreview();
}

async function loadMmsVersion(versionNumber) {
  loadMenuMontajeVersion(versionNumber);
  menuMontajeSelectableSelectedVersion = Number(versionNumber || quoteDraft?.menuMontajeVersion || 1);
  renderMmsVersionControls();
  renderMmsDateSalonSelect();
  renderMmsEntriesTable();
  await loadMmsFormByKey(menuMontajeSelectableSelectedKey);
}

async function openMenuMontajeSelectableModal() {
  if (!quoteDraft) return toast("Primero abre una cotizacion.");
  await ensureMenuMontajeSelectableCatalogLoaded(false);
  mmsPrimaryMode = "menu";
  mmsShowAllGuarniciones = false;
  mmsShowAllPostres = false;
  mmsLineItemsDraft = [];
  mmsActiveLineKey = "";
  mmsSelectedPlatoItems = [];
  mmsPlatoQty = 1;
  mmsSelectedSalsaIds = [];
  mmsSelectedBebidaIds = [];
  mmsGuarnicionQtyById = {};
  mmsPostreQtyById = {};
  mmsBebidaQtyById = {};
  mmsCurrentStage = "plato";
  if (el.mmsStageFilter) el.mmsStageFilter.value = "";
  if (el.btnMmsToggleGuarnicionesGlobal) el.btnMmsToggleGuarnicionesGlobal.textContent = "Mas guarniciones";
  if (el.btnMmsTogglePostresGlobal) el.btnMmsTogglePostresGlobal.textContent = "Mas postres";
  renderMmsMenuSectionSelect("General");
  renderMmsProteinOptions();
  renderMmsMontajeCatalogs();
  const model = ensureMenuMontajeModel();
  menuMontajeSelectableSelectedVersion = Number(model.currentVersion || 1);
  renderMmsVersionControls();
  renderMmsDateSalonSelect();
  renderMmsEntriesTable();
  await loadMmsFormByKey(menuMontajeSelectableSelectedKey);
  renderMmsPrimaryTabs();
  renderMmsStageTabs();
  renderMmsStageOptions();
  renderMmsSelectionSummary();
  renderMmsComandaPreview();
  if (el.menuMontajeSelectableBackdrop) el.menuMontajeSelectableBackdrop.hidden = false;
}

function buildMmsSelectionPayload() {
  const lineItems = getCombinedMmsLineItems();
  const firstPlato = lineItems[0] || null;
  const salsaIds = Array.from(new Set(lineItems.flatMap((item) => item.salsaIds || [])));
  const guarnicionIds = Array.from(new Set(lineItems.flatMap((item) => item.guarnicionIds || [])));
  const guarnicionQtys = {};
  const postreIds = Array.from(new Set(lineItems.flatMap((item) => item.postreIds || [])));
  const postreQtys = lineItems.reduce((acc, item) => Object.assign(acc, normalizeMmsQtyMap(item.postreQtys)), {});
  const bebidaIds = Array.from(new Set(lineItems.flatMap((item) => item.bebidaIds || [])));
  const bebidaQtys = lineItems.reduce((acc, item) => Object.assign(acc, normalizeMmsQtyMap(item.bebidaQtys)), {});
  const comentarioIds = Array.from(new Set(lineItems.flatMap((item) => item.comentarioIds || [])));
  const comentarioLibre = lineItems.map((item) => String(item.comentarioLibre || "").trim()).find(Boolean) || "";
  return {
    section: String(el.mmsMenuSection?.value || "General").trim() || "General",
    platoId: Number(firstPlato?.platoId || 0) || null,
    platoQty: Math.max(1, Math.floor(Number(firstPlato?.qty || 1))),
    preparacionId: Number(firstPlato?.preparacionId || 0) || null,
    platoItems: lineItems.map((item) => ({
      key: String(item?.key || `pl_${uid()}`),
      platoId: Number(item?.platoId || 0) || null,
      preparacionId: Number(item?.preparacionId || 0) || null,
      qty: Math.max(1, Math.floor(Number(item?.qty || 1))),
    })).filter((item) => item.platoId),
    lineItems: lineItems.map((item) => ({
      key: String(item?.key || `ln_${uid()}`),
      platoId: Number(item?.platoId || 0) || null,
      preparacionId: Number(item?.preparacionId || 0) || null,
      qty: Math.max(1, Math.floor(Number(item?.qty || 1))),
      salsaIds: (Array.isArray(item?.salsaIds) ? item.salsaIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0),
      guarnicionIds: (Array.isArray(item?.guarnicionIds) ? item.guarnicionIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0),
      guarnicionQtys: {},
      postreIds: (Array.isArray(item?.postreIds) ? item.postreIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0),
      postreQtys: normalizeMmsQtyMap(item?.postreQtys),
      bebidaIds: (Array.isArray(item?.bebidaIds) ? item.bebidaIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0),
      bebidaQtys: normalizeMmsQtyMap(item?.bebidaQtys),
      comentarioIds: (Array.isArray(item?.comentarioIds) ? item.comentarioIds : []).map((x) => Number(x)).filter((n) => Number.isFinite(n) && n > 0),
      comentarioLibre: String(item?.comentarioLibre || "").trim(),
    })).filter((item) => item.platoId),
    salsaIds,
    guarnicionIds,
    guarnicionQtys,
    postreIds,
    postreQtys,
    bebidaIds,
    bebidaQtys,
    comentarioIds,
    comentarioLibre,
  };
}

function buildMmsMontajeSelectionPayload() {
  return {
    tipoId: Number(el.mmsMontajeTipo?.value || 0) || null,
    adicionalIds: selectedIdsFromChecklist(el.mmsMontajeAdicionales),
  };
}

async function saveMenuMontajeSelectableFromModal({ updateCurrentVersion = false } = {}) {
  if (!quoteDraft) return;
  const key = String(el.mmsDateSalonSelect?.value || "").trim();
  if (!key || !key.includes("|")) return toast("Selecciona fecha y salon.");
  const [date, salon] = key.split("|");
  const menuTitle = String(el.mmsMenuTitle?.value || "").trim();
  const menuQty = getMmsTotalPlatoQty();
  const menuDescription = buildMmsMenuDescriptionFromForm();
  const montajeDescription = String(el.mmsMontajeDescription?.value || "").trim() || buildMmsMontajeDescriptionFromForm();
  syncMmsMenuQtyField();
  if (el.mmsMenuDescription) el.mmsMenuDescription.value = menuDescription;
  if (!menuTitle && !menuDescription && !montajeDescription) {
    return toast("Agrega al menos menu o montaje para guardar.");
  }
  const entries = ensureMenuMontajeDraft();
  const idx = entries.findIndex((x) => String(x.date || "") === date && String(x.salon || "") === salon);
  const row = {
    id: idx >= 0 ? String(entries[idx].id || uid()) : uid(),
    date,
    salon,
    menuTitle,
    menuQty,
    menuDescription,
    montajeDescription,
    menuSelection: buildMmsSelectionPayload(),
    montajeSelection: buildMmsMontajeSelectionPayload(),
    updatedAt: new Date().toISOString(),
  };
  if (idx >= 0) entries[idx] = row;
  else entries.push(row);

  syncQuoteDraftFromQuoteFormLoose();
  const model = ensureMenuMontajeModel();
  let targetVersion = Number(menuMontajeSelectableSelectedVersion || model.currentVersion || 1);
  const nowIso = new Date().toISOString();
  const compactEntries = compactMenuMontajeEntries(entries);
  let createdNewVersion = false;
  let unchanged = false;
  if (updateCurrentVersion) {
    const targetIdx = model.versions.findIndex((v) => Number(v.version) === targetVersion);
    if (targetIdx >= 0) {
      model.versions[targetIdx] = { ...model.versions[targetIdx], entries: compactEntries, savedAt: nowIso };
    } else {
      model.versions.push({ version: targetVersion, entries: compactEntries, savedAt: nowIso });
    }
  } else {
    const currentSnapshot = model.versions.find((v) => Number(v.version || 0) === Number(targetVersion || 0))
      || model.versions[model.versions.length - 1]
      || null;
    if (currentSnapshot && areMenuMontajeEntriesEqual(currentSnapshot.entries, compactEntries)) {
      unchanged = true;
    } else {
      const nextVersion = Math.max(0, ...model.versions.map((v) => Number(v.version || 0))) + 1;
      targetVersion = nextVersion;
      model.versions.push({ version: nextVersion, entries: compactEntries, savedAt: nowIso });
      createdNewVersion = true;
    }
  }

  quoteDraft.menuMontajeVersions = normalizeMenuMontajeVersionHistory(model.versions);
  quoteDraft.menuMontajeVersion = targetVersion;
  quoteDraft.menuMontajeEntries = deepClone(entries);
  if (!String(quoteDraft.code || "").trim()) {
    const code = await requestServerQuoteCode();
    quoteDraft.code = code || buildQuoteCode();
    if (el.quoteCode) el.quoteCode.value = quoteDraft.code;
  }
  if (el.mmsDocNo) el.mmsDocNo.value = quoteDraft.code;
  quoteDraft.quotedAt = nowIso;
  menuMontajeSelectableSelectedVersion = targetVersion;
  renderMmsVersionControls();

  const eventId = String(el.quoteEventId?.value || "").trim();
  const ev = (state.events || []).find((x) => String(x.id || "") === eventId);
  if (ev) {
    const reservationKey = reservationKeyFromEvent(ev);
    const series = getEventSeries(ev);
    for (const item of series) item.quote = deepClone(quoteDraft);
    quoteDraft = deepClone(quoteDraft);
    appendHistoryByKey(reservationKey, ev.userId || "", unchanged
      ? `Menu & Montaje (listas) verificado sin cambios (V${targetVersion}).`
      : `Menu & Montaje (listas) ${updateCurrentVersion ? "actualizado" : "guardado"} en V${targetVersion}.`);
    persist();
    render();
    renderQuoteVersionControls();
  }

  renderMenuMontajeSelect();
  renderMenuMontajeEntriesTable();
  renderMmsDateSalonSelect();
  renderMmsEntriesTable();
  toast(unchanged
    ? `Sin cambios detectados. Se mantiene V${targetVersion}.`
    : (updateCurrentVersion
      ? `Menu & Montaje actualizado en V${targetVersion}.`
      : (createdNewVersion
        ? `Menu & Montaje guardado. Version V${targetVersion} creada.`
        : `Menu & Montaje guardado en V${targetVersion}.`)));
}

// Solicita al servidor el siguiente codigo de cotizacion disponible.
// Si el backend responde con un codigo valido como COT-123, ese valor se usa
// como fuente principal para evitar duplicados entre varios usuarios.
async function requestServerQuoteCode() {
  try {
    const endpoint = buildApiUrlFromStateUrl(activeApiStateUrl, "doc-code-next");
    const response = await fetch(endpoint, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
    });
    if (!response.ok) return "";
    const payload = await response.json();
    const code = String(payload?.code || "").trim().toUpperCase();
    return /^COT-\d+$/.test(code) ? code : "";
  } catch (_) {
    return "";
  }
}

// Genera un codigo local de respaldo recorriendo las cotizaciones existentes.
// Solo se usa cuando el servidor no pudo entregar el siguiente correlativo,
// para que la cotizacion nunca se quede sin codigo al guardar o versionar.
function buildQuoteCode() {
  let maxNum = 0;
  const parseCodeNum = (raw) => {
    const m = String(raw || "").trim().toUpperCase().match(/^COT-(\d{1,})$/);
    if (!m) return 0;
    const n = Number(m[1]);
    return Number.isFinite(n) && n > 0 ? n : 0;
  };

  for (const ev of state.events || []) {
    const quote = ev?.quote || null;
    const direct = parseCodeNum(quote?.code);
    if (direct > maxNum) maxNum = direct;
    const versions = Array.isArray(quote?.versions) ? quote.versions : [];
    for (const v of versions) {
      const vn = parseCodeNum(v?.code);
      if (vn > maxNum) maxNum = vn;
    }
  }

  const next = maxNum + 1;
  return `COT-${String(next).padStart(3, "0")}`;
}

// Reune el contexto del evento asociado a la cotizacion.
// Devuelve el evento base y su rango de fechas para autocompletar encabezados,
// documentos y datos dependientes del evento sin repetir busquedas en varios lugares.
function getQuoteEventMeta(eventId) {
  const ev = state.events.find(x => x.id === eventId);
  if (!ev) return null;
  const series = getEventSeries(ev).slice().sort((a, b) => a.date.localeCompare(b.date));
  const startDate = series[0]?.date || ev.date;
  const endDate = series[series.length - 1]?.date || ev.date;
  return { ev, startDate, endDate };
}

// Fija en el formulario las fechas reales de la reservacion.
// Con esto la cotizacion siempre usa el mismo rango del evento y evita
// que se guarden fechas fuera del periodo reservado.
function syncQuoteReservationDates(meta = getQuoteEventMeta(el.quoteEventId?.value)) {
  const startDate = String(meta?.startDate || "").trim();
  const endDate = String(meta?.endDate || startDate || "").trim();
  if (el.quoteEventDate) {
    el.quoteEventDate.value = startDate;
    el.quoteEventDate.min = startDate;
    el.quoteEventDate.max = startDate;
    el.quoteEventDate.readOnly = true;
  }
  if (el.quoteEndDate) {
    el.quoteEndDate.value = endDate;
    el.quoteEndDate.min = endDate;
    el.quoteEndDate.max = endDate;
    el.quoteEndDate.readOnly = true;
  }
  if (quoteDraft) {
    quoteDraft.eventDate = startDate;
    quoteDraft.endDate = endDate;
  }
}

function calculateQuoteDueDateDefault(docDateIso, eventDateIso) {
  const eventDate = String(eventDateIso || "").trim();
  if (!eventDate) return "";
  const base = new Date(`${eventDate}T00:00:00`);
  if (Number.isNaN(base.getTime())) return "";
  return toISODate(addDays(base, -30));
}

function syncQuoteDateDefaults({ forceDueDate = false } = {}) {
  const eventDate = String(el.quoteEventDate?.value || quoteDraft?.eventDate || "").trim();
  if (el.quoteDocDate && !String(el.quoteDocDate.value || "").trim() && eventDate) {
    el.quoteDocDate.value = eventDate;
  }
  const docDate = String(el.quoteDocDate?.value || eventDate || "").trim();
  const dueDefault = calculateQuoteDueDateDefault(docDate, eventDate);
  if (!el.quoteDueDate || !dueDefault) return;
  const dueLockedAuto = String(el.quoteDueDate.dataset.autoDefault || "1") === "1";
  if (forceDueDate || !String(el.quoteDueDate.value || "").trim() || dueLockedAuto) {
    el.quoteDueDate.value = dueDefault;
    el.quoteDueDate.dataset.autoDefault = "1";
  }
}

// Rellena el encabezado de la cotizacion con datos del borrador, del evento,
// de la empresa y del encargado. Si force=true, sobrescribe el formulario;
// de lo contrario, solo completa campos vacios para no pisar lo que ya eligio el usuario.
function fillQuoteHeaderFields(force = false, allowContextFallback = true) {
  if (!quoteDraft) return;
  const company = (state.companies || []).find(c => c.id === el.quoteCompany.value);
  const manager = company?.managers?.find(m => m.id === el.quoteManagerSelect.value);
  const meta = getQuoteEventMeta(el.quoteEventId.value);
  const apply = (node, value) => {
    if (!node) return;
    if (force || !String(node.value || "").trim()) {
      node.value = value || "";
    }
  };

  apply(el.quoteContact, quoteDraft.contact || (allowContextFallback ? (manager?.name || company?.owner || "") : ""));
  apply(el.quoteEmail, quoteDraft.email || (allowContextFallback ? (manager?.email || company?.email || "") : ""));
  apply(el.quoteBillTo, quoteDraft.billTo || (allowContextFallback ? (company?.billTo || company?.businessName || company?.name || "") : ""));
  apply(el.quoteAddress, quoteDraft.address || (allowContextFallback ? (company?.address || "") : ""));
  apply(el.quoteEventType, quoteDraft.eventType || (allowContextFallback ? (company?.eventType || "") : ""));
  apply(el.quoteVenue, quoteDraft.venue || (allowContextFallback ? (meta?.ev?.salon || "") : ""));
  apply(el.quoteSchedule, quoteDraft.schedule || (allowContextFallback ? `${meta?.ev?.startTime || ""} a ${meta?.ev?.endTime || ""}`.trim() : ""));
  apply(el.quoteCode, quoteDraft.code || "");
  apply(el.quoteDocDate, quoteDraft.docDate || (allowContextFallback ? (meta?.startDate || "") : ""));
  apply(el.quotePhone, quoteDraft.phone || (allowContextFallback ? (manager?.phone || company?.phone || "") : ""));
  apply(el.quoteNIT, quoteDraft.nit || (allowContextFallback ? (company?.nit || "") : ""));
  apply(el.quotePeople, quoteDraft.people || "");
  apply(el.quoteEventDate, quoteDraft.eventDate || (allowContextFallback ? (meta?.startDate || "") : ""));
  apply(el.quoteEndDate, quoteDraft.endDate || (allowContextFallback ? (meta?.endDate || "") : ""));
  syncQuoteReservationDates(meta);
  syncQuoteDateDefaults({ forceDueDate: !quoteDraft.dueDate });
  apply(el.quoteFolio, quoteDraft.folio || "");
  apply(el.quoteInternalNotes, quoteDraft.internalNotes || quoteDraft.notes || "");
  if (force || !getSelectedQuotePaymentTypesFromForm().length) {
    setQuotePaymentTypesOnForm(quoteDraft.paymentType || "");
  }
}

function renderQuoteServiceDateSelect(selectedDate = null) {
  if (!el.quoteServiceDate) return;
  const dates = getQuoteRangeDates();
  el.quoteServiceDate.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = "Fecha del servicio";
  placeholder.disabled = true;
  placeholder.selected = true;
  el.quoteServiceDate.appendChild(placeholder);

  for (const d of dates) {
    const opt = document.createElement("option");
    opt.value = d;
    opt.textContent = d;
    el.quoteServiceDate.appendChild(opt);
  }

  if (selectedDate && dates.includes(selectedDate)) {
    el.quoteServiceDate.value = selectedDate;
  }
  syncQuoteServiceDateRequired();
}

function syncQuoteServiceDateRequired() {
  if (!el.quoteServiceDate) return;
  const hasItems = !!(quoteDraft && Array.isArray(quoteDraft.items) && quoteDraft.items.length > 0);
  el.quoteServiceDate.required = !hasItems;
  if (hasItems) {
    el.quoteServiceDate.classList.remove("fieldInvalid");
  }
}

function placeEvent(ev, layout = null) {
  // find day col
  const d = new Date(ev.date + "T00:00:00");
  const dayIdx = Math.floor((stripTime(d) - stripTime(viewStart)) / (1000 * 60 * 60 * 24));
  const col = el.grid.querySelector(`.dayCol[data-day-index="${dayIdx}"]`);
  if (!col) return;

  const top = timeToY(ev.startTime);
  const bottom = timeToY(ev.endTime);
  const height = Math.max(42, bottom - top);

  const card = document.createElement("div");
  card.className = "event";
  card.dataset.eventId = ev.id;
  card.style.top = `${top + 6}px`;
  card.style.height = `${height - 10}px`;

  // If events overlap, split the width in lanes so each remains clickable.
  const lane = layout?.lane ?? 0;
  const lanes = Math.max(1, layout?.lanes ?? 1);
  const sidePad = 8;
  const gap = 6;
  const colWidth = col.clientWidth || col.getBoundingClientRect().width || 240;
  const usable = Math.max(40, colWidth - sidePad * 2);
  const laneWidth = Math.max(56, (usable - (lanes - 1) * gap) / lanes);
  const leftPx = sidePad + lane * (laneWidth + gap);
  const isCompact = laneWidth < 165 || lanes >= 3;
  card.style.left = `${leftPx}px`;
  card.style.width = `${laneWidth}px`;
  card.style.right = "auto";
  card.style.zIndex = String(10 + lane);
  card.classList.toggle("compact", isCompact || !!topbarSettings.compactEvents);

  const c = statusColor(ev.status);
  card.style.background = `linear-gradient(180deg, ${hexToRgba(c, 0.20)}, ${hexToRgba(c, 0.12)})`;
  card.style.borderColor = hexToRgba(c, 0.42);
  card.style.setProperty("--status-color", c);

  const user = state.users.find(u => u.id === ev.userId) || { name: "-" };
  const series = getEventSeries(ev).slice().sort((a, b) => a.date.localeCompare(b.date));
  const seriesTotal = series.length;
  const seriesIndex = Math.max(0, series.findIndex(x => x.id === ev.id));
  const isSeriesLastDay = !seriesTotal || series[seriesTotal - 1].id === ev.id;
  const canStretch = uniqueSlotsFromSeries(series).length === 1 && (seriesTotal <= 1 || isSeriesLastDay);
  const statusShort = ev.status
    .split(/\s+/)
    .filter(Boolean)
    .map(w => w[0])
    .join("")
    .slice(0, 3)
    .toUpperCase();
  const followUpLabel = buildFollowUpLabel(ev);
  const reminderLabel = reminderBadgeText(ev);
  const reminderClass = reminderBadgeClass(ev);
  const institutionName = getEventInstitutionName(ev);
  const paxNumber = Number(ev.pax || 0);
  const paxLabel = paxNumber > 0 ? `PAX ${Math.round(paxNumber)}` : "";
  const latestQuoteTotalLabel = getEventLatestQuoteTotalLabel(ev);
  const detailRows = [
    { label: "Horario", value: `${ev.startTime}-${ev.endTime}` },
    { label: "Salon", value: ev.salon || "-" },
    { label: "Vendedor", value: user.name || "-" },
  ];
  if (institutionName) detailRows.push({ label: "Cliente", value: institutionName, optional: true });
  if (paxLabel) detailRows.push({ label: "PAX", value: paxLabel.replace(/^PAX\s+/i, ""), optional: true });
  if (latestQuoteTotalLabel) detailRows.push({ label: "Cot.", value: latestQuoteTotalLabel, optional: true });
  const detailRowsHtml = detailRows
    .map((row) => `
      <div class="eventFact${row.optional ? " optional" : ""}">
        <span class="eventFactLabel">${escapeHtml(row.label)}</span>
        <span class="eventFactValue" title="${escapeHtml(String(row.value || "-"))}">${escapeHtml(String(row.value || "-"))}</span>
      </div>
    `)
    .join("");
  const secondaryBadges = [
    seriesTotal > 1 ? `<span class="eventMiniBadge seriesPill">Reserva ${seriesIndex + 1}/${seriesTotal}</span>` : "",
    followUpLabel ? `<span class="eventMiniBadge followupPill">${escapeHtml(followUpLabel)}</span>` : "",
    reminderLabel ? `<span class="eventMiniBadge ${reminderClass}">${escapeHtml(reminderLabel)}</span>` : "",
  ].filter(Boolean).join("");

  card.innerHTML = `
    <div class="eventStatusChip">
      <span class="dot" style="background:${c}"></span>
      <span>${escapeHtml(ev.status)}</span>
      <small>${escapeHtml(statusShort)}</small>
    </div>
    <div class="eventInner">
      <div class="eventMeta">
        <div class="eventTitle" title="${escapeHtml(ev.name)}">${escapeHtml(ev.name)}</div>
        <div class="eventFacts">${detailRowsHtml}</div>
        ${secondaryBadges ? `<div class="eventSub eventSubBadges">${secondaryBadges}</div>` : ""}
      </div>
    </div>
    ${canStretch ? `<div class="eventStretch" title="Estirar dias"></div>` : ""}
  `;

  card.addEventListener("mousedown", (e) => {
    if (e.button !== 0) return;
    startEventDrag(e, ev, card);
  });
  card.addEventListener("click", () => {
    if (Date.now() < interaction.suppressClickUntil) return;
    openModalForEdit(ev.id);
  });

  const stretchHandle = canStretch ? card.querySelector(".eventStretch") : null;
  if (stretchHandle) {
    stretchHandle.addEventListener("click", (e) => {
      e.preventDefault();
      e.stopPropagation();
    });
    stretchHandle.addEventListener("mousedown", (e) => {
      if (e.button !== 0) return;
      startEventStretch(e, ev);
    });
  }
  col.appendChild(card);
}
function computeDayEventLayout(dayEvents) {
  const layout = new Map();
  if (!dayEvents.length) return layout;

  // 1) Greedy lane assignment by start time.
  const items = dayEvents
    .slice()
    .sort((a, b) => {
      const t = compareTime(a.startTime, b.startTime);
      if (t !== 0) return t;
      return compareTime(a.endTime, b.endTime);
    })
    .map(e => ({
      id: e.id,
      start: timeToMinutes(e.startTime),
      end: timeToMinutes(e.endTime),
      lane: 0,
    }));

  const laneEnd = [];
  for (const it of items) {
    let chosen = -1;
    for (let i = 0; i < laneEnd.length; i++) {
      if (laneEnd[i] <= it.start) {
        chosen = i;
        break;
      }
    }
    if (chosen === -1) {
      chosen = laneEnd.length;
      laneEnd.push(it.end);
    } else {
      laneEnd[chosen] = it.end;
    }
    it.lane = chosen;
  }

  // 2) Connected overlap groups to know how many lanes each event shares.
  const parent = new Map(items.map(x => [x.id, x.id]));
  const find = (x) => {
    let p = parent.get(x);
    while (p !== parent.get(p)) p = parent.get(p);
    return p;
  };
  const union = (a, b) => {
    const ra = find(a), rb = find(b);
    if (ra !== rb) parent.set(rb, ra);
  };

  for (let i = 0; i < dayEvents.length; i++) {
    for (let j = i + 1; j < dayEvents.length; j++) {
      if (timesOverlap(dayEvents[i].startTime, dayEvents[i].endTime, dayEvents[j].startTime, dayEvents[j].endTime)) {
        union(dayEvents[i].id, dayEvents[j].id);
      }
    }
  }

  const componentLaneCount = new Map();
  for (const it of items) {
    const root = find(it.id);
    const prev = componentLaneCount.get(root) ?? 0;
    componentLaneCount.set(root, Math.max(prev, it.lane + 1));
  }

  for (const it of items) {
    const root = find(it.id);
    layout.set(it.id, { lane: it.lane, lanes: componentLaneCount.get(root) ?? 1 });
  }

  return layout;
}

// ================== Events / Modal ==================

function shiftViewBackward() {
  if (navMode === "month") {
    monthCursor = startOfMonth(addMonths(monthCursor, -1));
    viewStart = startOfWeek(monthCursor);
  } else if (navMode === "day") {
    viewStart = addDays(viewStart, -1);
  } else {
    viewStart = addDays(viewStart, -getVisibleDayCount());
  }
  render();
}

function shiftViewForward() {
  if (navMode === "month") {
    monthCursor = startOfMonth(addMonths(monthCursor, 1));
    viewStart = startOfWeek(monthCursor);
  } else if (navMode === "day") {
    viewStart = addDays(viewStart, 1);
  } else {
    viewStart = addDays(viewStart, getVisibleDayCount());
  }
  render();
}

function getValidationFieldLabel(target) {
  if (!target) return "este campo";
  const dataLabel = String(target.getAttribute?.("data-label") || "").trim();
  if (dataLabel) return dataLabel;
  const ariaLabel = String(target.getAttribute?.("aria-label") || "").trim();
  if (ariaLabel) return ariaLabel;
  const placeholder = String(target.getAttribute?.("placeholder") || "").trim();
  if (placeholder) return placeholder;
  const labelWrap = target.closest?.("label");
  if (labelWrap) {
    const titleNode = labelWrap.querySelector("span");
    const labelText = String((titleNode?.textContent || labelWrap.textContent || "")).trim();
    if (labelText) return labelText.replace(/\s+/g, " ").replace(/[:*]\s*$/, "");
  }
  const idText = String(target.id || target.name || "").trim();
  if (idText) return idText;
  return "este campo";
}

function setupStyledInvalidAlerts() {
  if (setupStyledInvalidAlerts._bound) return;
  setupStyledInvalidAlerts._bound = true;

  document.addEventListener("invalid", (e) => {
    const target = e.target;
    if (!(target instanceof HTMLElement)) return;
    e.preventDefault();
    target.classList.add("fieldInvalid");
    const nativeMessage = String(target.validationMessage || "").trim();
    const label = getValidationFieldLabel(target);
    let message = nativeMessage || `Completa el campo ${label}.`;
    if (/completa este campo|fill out this field/i.test(nativeMessage)) {
      message = `Completa el campo ${label}.`;
    }
    toast(message);
  }, true);

  const clearInvalid = (e) => {
    const target = e.target;
    if (!(target instanceof HTMLElement)) return;
    if (target.classList.contains("fieldInvalid")) target.classList.remove("fieldInvalid");
  };
  document.addEventListener("input", clearInvalid, true);
  document.addEventListener("change", clearInvalid, true);
}

function bindSafeBackdropClose(backdropEl, closeFn) {
  if (!backdropEl || typeof closeFn !== "function") return;
  let startedOnBackdrop = false;
  backdropEl.addEventListener("mousedown", (e) => {
    startedOnBackdrop = e.target === backdropEl;
  });
  backdropEl.addEventListener("mouseup", (e) => {
    const endedOnBackdrop = e.target === backdropEl;
    if (startedOnBackdrop && endedOnBackdrop) closeFn();
    startedOnBackdrop = false;
  });
  backdropEl.addEventListener("mouseleave", () => {
    startedOnBackdrop = false;
  });
}

function openEventFinderModal() {
  if (!el.eventFinderBackdrop) return;
  const term = String(el.eventFinderSearch?.value || "").trim();
  if (term) {
    renderEventFinderResults(term);
  } else {
    renderEventFinderIdle();
  }
  el.eventFinderBackdrop.hidden = false;
  setTimeout(() => {
    if (el.eventFinderSearch) {
      el.eventFinderSearch.focus();
      el.eventFinderSearch.select();
    }
  }, 0);
}

function closeEventFinderModal() {
  if (!el.eventFinderBackdrop) return;
  el.eventFinderBackdrop.hidden = true;
}

function renderEventFinderIdle() {
  if (!el.eventFinderBody) return;
  el.eventFinderBody.innerHTML = "";
  const tr = document.createElement("tr");
  tr.innerHTML = `<td colspan="7">Escribe un criterio y presiona Enter para buscar.</td>`;
  el.eventFinderBody.appendChild(tr);
}

function buildEventFinderRows() {
  const usersById = new Map((state.users || []).map((u) => [String(u.id || "").trim(), String(u.fullName || u.name || "").trim()]));
  const companiesById = new Map((state.companies || []).map((c) => [String(c.id || "").trim(), c]));
  const rows = [];
  for (const ev of state.events || []) {
    const quote = ev?.quote || {};
    const companyId = String(quote.companyId || "").trim();
    const company = companiesById.get(companyId) || null;
    const companyName = String(company?.name || quote.companyName || "").trim();
    const code = String(quote.code || reservationKeyFromEvent(ev) || ev.id || "").trim();
    const userName = usersById.get(String(ev.userId || "").trim()) || "";
    const parts = [
      String(code || ""),
      String(ev.name || ""),
      String(ev.date || ""),
      String(ev.startTime || ""),
      String(ev.endTime || ""),
      String(ev.salon || ""),
      String(ev.status || ""),
      String(ev.notes || ""),
      String(code || ""),
      String(companyName || ""),
      String(quote.contact || ""),
      String(quote.managerName || ""),
      String(userName || ""),
    ];
    rows.push({
      ev,
      eventId: String(ev.id || ""),
      docNo: String(code || "").trim(),
      date: String(ev.date || ""),
      eventName: String(ev.name || ""),
      salon: String(ev.salon || ""),
      companyName,
      status: String(ev.status || ""),
      code,
      searchBlob: parts.join(" ").toLowerCase(),
    });
  }
  rows.sort((a, b) => {
    const d = String(b.date || "").localeCompare(String(a.date || ""));
    if (d !== 0) return d;
    return String(a.ev?.startTime || "").localeCompare(String(b.ev?.startTime || ""));
  });
  return rows;
}

function renderEventFinderResults(rawTerm = "") {
  if (!el.eventFinderBody) return;
  const term = String(rawTerm || "").trim();
  const allRows = buildEventFinderRows();
  const rows = term
    ? allRows.filter((r) => matchesLikeSearch(r.searchBlob, term)).slice(0, 150)
    : allRows.slice(0, 120);
  el.eventFinderBody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="7">Sin coincidencias.</td>`;
    el.eventFinderBody.appendChild(tr);
    return;
  }
  for (const r of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(r.docNo || "-")}</td>
      <td>${escapeHtml(r.date || "-")}</td>
      <td>${escapeHtml(r.eventName || "-")}</td>
      <td>${escapeHtml(r.salon || "-")}</td>
      <td>${escapeHtml(r.companyName || "-")}</td>
      <td>${escapeHtml(r.status || "-")}</td>
      <td><button class="btn" type="button" data-find-event-id="${escapeHtml(r.eventId || "")}">Mostrar</button></td>
    `;
    el.eventFinderBody.appendChild(tr);
  }
}

function openLeadsModal() {
  if (!el.leadsBackdrop) return;
  ensureLeadsFilterOptions();
  const term = String(el.leadsSearch?.value || "").trim();
  renderLeadsResults(term);
  el.leadsBackdrop.hidden = false;
  setTimeout(() => {
    if (el.leadsSearch) {
      el.leadsSearch.focus();
      el.leadsSearch.select();
    }
  }, 0);
}

function closeLeadsModal() {
  if (!el.leadsBackdrop) return;
  el.leadsBackdrop.hidden = true;
}

function renderLeadsIdle() {
  if (!el.leadsBody) return;
  el.leadsBody.innerHTML = "";
  const tr = document.createElement("tr");
  tr.innerHTML = `<td colspan="7">No hay leads para los filtros seleccionados.</td>`;
  el.leadsBody.appendChild(tr);
}

function buildLeadsRows() {
  const usersById = new Map((state.users || []).map((u) => [String(u.id || "").trim(), String(u.fullName || u.name || "").trim()]));
  const companiesById = new Map((state.companies || []).map((c) => [String(c.id || "").trim(), c]));
  const rows = [];
  for (const ev of state.events || []) {
    const status = String(ev?.status || "").trim();
    if (!LEADS_PIPELINE_STATUSES.has(status)) continue;
    const quote = ev?.quote || {};
    const companyId = String(quote.companyId || "").trim();
    const company = companiesById.get(companyId) || null;
    const companyName = String(company?.name || quote.companyName || "").trim();
    const managerName = String(quote.managerName || "").trim();
    const contactName = String(quote.contact || "").trim();
    const eventName = String(ev?.name || "").trim();
    const leadLabel = contactName && eventName && contactName !== eventName
      ? `${eventName} / ${contactName}`
      : (eventName || contactName || "-");
    const userName = usersById.get(String(ev?.userId || "").trim()) || "";
    const parts = [
      String(ev?.date || ""),
      String(ev?.startTime || ""),
      String(ev?.endTime || ""),
      String(ev?.salon || ""),
      String(status || ""),
      String(eventName || ""),
      String(contactName || ""),
      String(companyName || ""),
      String(managerName || ""),
      String(userName || ""),
      String(quote?.code || ""),
    ];
    rows.push({
      eventId: String(ev?.id || ""),
      date: String(ev?.date || ""),
      leadLabel,
      companyName,
      managerName: managerName || userName,
      sellerId: String(ev?.userId || "").trim(),
      sellerName: userName,
      status,
      salon: String(ev?.salon || ""),
      statusColor: statusColor(status),
      searchBlob: parts.join(" ").toLowerCase(),
      startTime: String(ev?.startTime || ""),
    });
  }
  rows.sort((a, b) => {
    const d = String(b.date || "").localeCompare(String(a.date || ""));
    if (d !== 0) return d;
    return String(a.startTime || "").localeCompare(String(b.startTime || ""));
  });
  return rows;
}

function ensureLeadsFilterOptions() {
  if (el.leadsStatusFilter) {
    const current = String(el.leadsStatusFilter.value || "");
    const options = ['<option value="">Todos los estados</option>'];
    for (const status of LEADS_PIPELINE_STATUS_ORDER) {
      options.push(`<option value="${escapeHtml(status)}">${escapeHtml(status)}</option>`);
    }
    el.leadsStatusFilter.innerHTML = options.join("");
    if (current && LEADS_PIPELINE_STATUS_ORDER.includes(current)) {
      el.leadsStatusFilter.value = current;
    }
  }
  if (el.leadsSellerFilter) {
    const current = String(el.leadsSellerFilter.value || "");
    const users = (state.users || [])
      .map(normalizeUserRecord)
      .filter((u) => u.active !== false)
      .sort((a, b) => String(a.fullName || a.name || "").localeCompare(String(b.fullName || b.name || ""), "es", { sensitivity: "base" }));
    const options = ['<option value="">Todos los vendedores</option>'];
    for (const u of users) {
      const id = String(u.id || "").trim();
      if (!id) continue;
      const label = String(u.fullName || u.name || u.username || "").trim() || id;
      options.push(`<option value="${escapeHtml(id)}">${escapeHtml(label)}</option>`);
    }
    el.leadsSellerFilter.innerHTML = options.join("");
    if (current && users.some((u) => String(u.id || "") === current)) {
      el.leadsSellerFilter.value = current;
    }
  }
}

function getLeadsFilters() {
  return {
    status: String(el.leadsStatusFilter?.value || "").trim(),
    sellerId: String(el.leadsSellerFilter?.value || "").trim(),
    dateFrom: String(el.leadsDateFrom?.value || "").trim(),
    dateTo: String(el.leadsDateTo?.value || "").trim(),
  };
}

function applyLeadsFilters(rows) {
  const filters = getLeadsFilters();
  return (rows || []).filter((row) => {
    const rowDate = String(row?.date || "").trim();
    if (filters.status && String(row?.status || "").trim() !== filters.status) return false;
    if (filters.sellerId && String(row?.sellerId || "").trim() !== filters.sellerId) return false;
    if (filters.dateFrom && rowDate && rowDate < filters.dateFrom) return false;
    if (filters.dateTo && rowDate && rowDate > filters.dateTo) return false;
    return true;
  });
}

function resetLeadsFilters() {
  if (el.leadsStatusFilter) el.leadsStatusFilter.value = "";
  if (el.leadsSellerFilter) el.leadsSellerFilter.value = "";
  if (el.leadsDateFrom) el.leadsDateFrom.value = "";
  if (el.leadsDateTo) el.leadsDateTo.value = "";
}

function renderLeadsResults(rawTerm = "") {
  if (!el.leadsBody) return;
  const term = String(rawTerm || "").trim();
  const allRows = applyLeadsFilters(buildLeadsRows());
  const rows = term
    ? allRows.filter((r) => matchesLikeSearch(r.searchBlob, term)).slice(0, 180)
    : allRows.slice(0, 140);
  el.leadsBody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="7">No hay leads que coincidan con la busqueda.</td>`;
    el.leadsBody.appendChild(tr);
    return;
  }
  for (const r of rows) {
    const badgeBg = hexToRgba(r.statusColor || "#94a3b8", 0.22);
    const badgeBorder = hexToRgba(r.statusColor || "#94a3b8", 0.56);
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(r.date || "-")}</td>
      <td>${escapeHtml(r.leadLabel || "-")}</td>
      <td>${escapeHtml(r.companyName || "-")}</td>
      <td>${escapeHtml(r.managerName || "-")}</td>
      <td><span class="salesStatusBadge" style="background:${escapeHtml(badgeBg)};border-color:${escapeHtml(badgeBorder)}">${escapeHtml(r.status || "-")}</span></td>
      <td>${escapeHtml(r.salon || "-")}</td>
      <td><button class="btn" type="button" data-lead-event-id="${escapeHtml(r.eventId || "")}">Ver</button></td>
    `;
    el.leadsBody.appendChild(tr);
  }
}

function revealEventInCalendar(eventId) {
  const id = String(eventId || "").trim();
  if (!id) return;
  const ev = (state.events || []).find((x) => String(x.id || "") === id);
  if (!ev) return toast("No se encontro el evento seleccionado.");

  const targetDate = stripTime(new Date(`${String(ev.date || "").trim()}T00:00:00`));
  if (Number.isNaN(targetDate.getTime())) return toast("La fecha del evento es invalida.");

  monthCursor = startOfMonth(targetDate);
  if (navMode === "month") {
    viewStart = startOfWeek(monthCursor);
  } else if (navMode === "week") {
    viewStart = startOfWeek(targetDate);
  } else {
    viewStart = stripTime(targetDate);
  }

  if (selectedSalon !== ALL_ROOMS_VALUE && String(selectedSalon || "") !== String(ev.salon || "")) {
    selectedSalon = ALL_ROOMS_VALUE;
    if (el.roomSelect) {
      el.roomSelect.value = selectedSalon;
      syncEnhancedSelectValue(el.roomSelect, selectedSalon);
      ensureCustomTopbarSelect(el.roomSelect);
    }
  }

  render();
}

function bindEvents() {
  setupStyledInvalidAlerts();
  if (el.navMode) el.navMode.value = navMode;
  if (el.loginForm) {
    el.loginForm.addEventListener("submit", async (e) => {
      e.preventDefault();
      await doLogin();
    });
  }
  if (el.loginUserSelect) {
    el.loginUserSelect.addEventListener("change", () => {
      setLoginError("");
      updateLoginAvatarFromSelect();
    });
  }
  if (el.loginPassword) {
    el.loginPassword.addEventListener("input", () => setLoginError(""));
  }
  if (el.btnSideDashboardReports) {
    el.btnSideDashboardReports.addEventListener("click", () => {
      closeBlockingOverlaysForNavigation();
      closeSettingsPanel();
      prepareModuleModalOpen("reports");
      openSalesReportModal();
      setSideNavActive("reports");
    });
  }
  if (el.btnSideCalendar) {
    el.btnSideCalendar.addEventListener("click", () => {
      closeBlockingOverlaysForNavigation();
      closeSettingsPanel();
      showCalendarModule();
      render();
    });
  }
  if (el.btnSideReports) {
    el.btnSideReports.addEventListener("click", () => {
      closeBlockingOverlaysForNavigation();
      closeSettingsPanel();
      showReportsHub();
    });
  }
  if (el.btnSideSettings) {
    el.btnSideSettings.addEventListener("click", () => {
      closeBlockingOverlaysForNavigation();
      closeSettingsPanel();
      showSettingsHub();
    });
  }
  if (el.btnLogout) {
    el.btnLogout.addEventListener("click", () => {
      logoutCurrentSession();
    });
  }
  if (el.btnSupportMenu) {
    el.btnSupportMenu.setAttribute("aria-expanded", "false");
    if (el.supportOptions) el.supportOptions.hidden = true;
    el.btnSupportMenu.addEventListener("click", () => {
      const expanded = String(el.btnSupportMenu.getAttribute("aria-expanded") || "false") === "true";
      const next = !expanded;
      el.btnSupportMenu.setAttribute("aria-expanded", next ? "true" : "false");
      if (el.supportOptions) el.supportOptions.hidden = !next;
    });
  }
  if (el.btnModuleCalendar) {
    el.btnModuleCalendar.addEventListener("click", () => {
      showCalendarModule();
      render();
    });
  }
  if (el.btnModuleReports) {
    el.btnModuleReports.addEventListener("click", () => {
      showReportsHub();
    });
  }
  if (el.btnModuleSettings) {
    el.btnModuleSettings.addEventListener("click", () => {
      showSettingsHub();
    });
  }
  if (el.btnBackFromReports) {
    el.btnBackFromReports.addEventListener("click", () => {
      showCalendarModule();
      render();
    });
  }
  if (el.btnBackFromSettings) {
    el.btnBackFromSettings.addEventListener("click", () => {
      if (el.settingsScreen) el.settingsScreen.hidden = true;
      showCalendarModule();
      render();
    });
  }

  el.btnPrev.addEventListener("click", shiftViewBackward);
  el.btnNext.addEventListener("click", shiftViewForward);
  el.btnToday.addEventListener("click", () => {
    goToTodayView();
    render();
  });
  window.addEventListener("pageshow", () => {
    goToTodayView();
    render();
  });
  window.addEventListener("resize", () => {
    syncCalendarVerticalOffset();
  });
  if (el.navMode) {
    el.navMode.addEventListener("change", (e) => {
      const nextMode = e.target.value === "month"
        ? "month"
        : (e.target.value === "day" ? "day" : "week");
      const prevMode = navMode;
      navMode = nextMode;
      if (navMode === "month") {
        monthCursor = startOfMonth(addDays(viewStart, 3));
        viewStart = startOfWeek(monthCursor);
      } else {
        if (prevMode === "month") viewStart = stripTime(addDays(viewStart, 3));
        viewStart = navMode === "week" ? startOfWeek(viewStart) : stripTime(viewStart);
      }
      render();
    });
  }

  el.roomSelect.addEventListener("change", (e) => {
    selectedSalon = e.target.value;
    render();
  });

  if (el.settingsPanel) {
    if (el.btnSettings && el.settingsMenu) {
      el.btnSettings.addEventListener("click", (e) => {
        e.stopPropagation();
        setSettingsPanelOpen(el.settingsPanel.hidden);
      });
      el.settingsMenu.addEventListener("focusout", () => {
        setTimeout(() => {
          if (!el.settingsMenu.contains(document.activeElement)) {
            closeSettingsPanel();
          }
        }, 0);
      });
    }
    if (el.btnToggleQuickAdd) {
      el.btnToggleQuickAdd.addEventListener("click", () => {
        const willOpen = el.quickAddGroup ? el.quickAddGroup.hidden : false;
        setQuickAddGroupOpen(willOpen);
      });
    }
    if (el.btnToggleReports) {
      el.btnToggleReports.addEventListener("click", () => {
        const willOpen = el.reportsGroup ? el.reportsGroup.hidden : false;
        setReportsGroupOpen(willOpen);
      });
    }

    if (el.settingShowLegend) {
      el.settingShowLegend.addEventListener("change", (e) => {
        topbarSettings.showLegend = !!e.target.checked;
        saveTopbarSettings();
        applyTopbarSettings();
      });
    }

    if (el.settingCompactEvents) {
      el.settingCompactEvents.addEventListener("change", (e) => {
        topbarSettings.compactEvents = !!e.target.checked;
        saveTopbarSettings();
        applyTopbarSettings({ rerender: true });
      });
    }

    if (el.settingShowWeekends) {
      el.settingShowWeekends.addEventListener("change", (e) => {
        topbarSettings.showWeekends = !!e.target.checked;
        saveTopbarSettings();
        applyTopbarSettings({ rerender: true });
      });
    }

    if (el.btnQuickAddInstitution) {
      el.btnQuickAddInstitution.addEventListener("click", () => {
        prepareModuleModalOpen("settings");
        closeSettingsPanel();
        openCompanyModal("");
      });
    }

    if (el.btnQuickAddManager) {
      el.btnQuickAddManager.addEventListener("click", () => {
        prepareModuleModalOpen("settings");
        closeSettingsPanel();
        openCompanyModal("");
        setTimeout(() => {
          try { el.managerName?.focus(); } catch (_) { }
        }, 0);
      });
    }

    if (el.btnQuickAddUser) {
      el.btnQuickAddUser.addEventListener("click", () => {
        prepareModuleModalOpen("settings");
        closeSettingsPanel();
        openUserModal();
      });
    }

    if (el.btnQuickAddService) {
      el.btnQuickAddService.addEventListener("click", () => {
        prepareModuleModalOpen("settings");
        closeSettingsPanel();
        openServiceModal();
      });
    }

    if (el.btnQuickAddSalon) {
      el.btnQuickAddSalon.addEventListener("click", () => {
        prepareModuleModalOpen("settings");
        closeSettingsPanel();
        openSalonesModal();
      });
    }

    if (el.btnQuickAddGlobalGoal) {
      el.btnQuickAddGlobalGoal.addEventListener("click", () => {
        prepareModuleModalOpen("settings");
        closeSettingsPanel();
        openGlobalGoalsModal();
      });
    }

    if (el.btnQuickAddChecklist) {
      el.btnQuickAddChecklist.addEventListener("click", () => {
        prepareModuleModalOpen("settings");
        closeSettingsPanel();
        openChecklistTemplateModal();
      });
    }

    if (el.btnQuickOpenMenuCatalog) {
      el.btnQuickOpenMenuCatalog.addEventListener("click", async () => {
        prepareModuleModalOpen("settings");
        closeSettingsPanel();
        await openMenuCatalogManagerModal("plato_fuerte");
      });
    }

    if (el.btnManageQuoteServiceTemplates) {
      el.btnManageQuoteServiceTemplates.addEventListener("click", () => {
        prepareModuleModalOpen("settings");
        closeSettingsPanel();
        openQuoteServiceTemplateManager();
      });
    }

    if (el.btnQuoteServiceTemplateManagerClose) {
      el.btnQuoteServiceTemplateManagerClose.addEventListener("click", closeQuoteServiceTemplateManager);
    }

    if (el.quoteServiceTemplateManagerBackdrop) {
      el.quoteServiceTemplateManagerBackdrop.addEventListener("click", (e) => {
        if (e.target === el.quoteServiceTemplateManagerBackdrop) closeQuoteServiceTemplateManager();
        const removeBtn = e.target.closest("[data-qstm-remove]");
        if (removeBtn) {
          const rowId = String(removeBtn.dataset.qstmRemove || "").trim();
          quoteServiceTemplateManagerItems = quoteServiceTemplateManagerItems.filter((item) => String(item.rowId || "") !== rowId);
          renderQuoteServiceTemplateManagerItems();
        }
      });
    }

    if (el.settingsQuoteServiceTemplateSelect) {
      el.settingsQuoteServiceTemplateSelect.addEventListener("change", () => {
        loadQuoteServiceTemplateManager(el.settingsQuoteServiceTemplateSelect.value || "");
      });
    }

    if (el.btnSettingsQuoteServiceTemplateNew) {
      el.btnSettingsQuoteServiceTemplateNew.addEventListener("click", () => {
        loadQuoteServiceTemplateManager("");
        if (el.settingsQuoteServiceTemplateName) el.settingsQuoteServiceTemplateName.value = "";
      });
    }

    if (el.btnSettingsQuoteTemplateServiceAdd) {
      el.btnSettingsQuoteTemplateServiceAdd.addEventListener("click", addServiceToQuoteServiceTemplateManager);
    }

    if (el.settingsQuoteTemplateServiceSearch) {
      el.settingsQuoteTemplateServiceSearch.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          addServiceToQuoteServiceTemplateManager();
        }
      });
    }

    if (el.btnSettingsQuoteServiceTemplateSave) {
      el.btnSettingsQuoteServiceTemplateSave.addEventListener("click", () => saveQuoteServiceTemplateFromManager(false));
    }

    if (el.btnSettingsQuoteServiceTemplateUpdate) {
      el.btnSettingsQuoteServiceTemplateUpdate.addEventListener("click", () => saveQuoteServiceTemplateFromManager(true));
    }

    if (el.btnSettingsQuoteServiceTemplateDelete) {
      el.btnSettingsQuoteServiceTemplateDelete.addEventListener("click", deleteQuoteServiceTemplateFromManager);
    }

    if (el.btnReportSales) {
      el.btnReportSales.addEventListener("click", () => {
        prepareModuleModalOpen("reports");
        openSalesReportModal();
      });
    }
    if (el.btnReportAccounting) {
      el.btnReportAccounting.addEventListener("click", () => {
        prepareModuleModalOpen("reports");
        openAccountingReportModal();
      });
    }
    if (el.btnReportOccupancy) {
      el.btnReportOccupancy.addEventListener("click", () => {
        prepareModuleModalOpen("reports");
        openOccupancyReportModal();
      });
    }
    if (el.btnReportDashboard) {
      el.btnReportDashboard.addEventListener("click", () => {
        prepareModuleModalOpen("reports");
        openDashboardReportModal();
      });
    }
    if (el.btnReportInstitution) {
      el.btnReportInstitution.addEventListener("click", () => {
        prepareModuleModalOpen("reports");
        openInstitutionReportModal();
      });
    }
  }

  el.grid.addEventListener("scroll", () => {
    el.timeCol.scrollTop = el.grid.scrollTop;
    el.daysHeader.scrollLeft = el.grid.scrollLeft;
  });
  el.timeCol.addEventListener("scroll", () => {
    el.grid.scrollTop = el.timeCol.scrollTop;
  });
  el.daysHeader.addEventListener("wheel", (e) => {
    if (!e.deltaY) return;
    e.preventDefault();
    el.grid.scrollTop += e.deltaY;
    el.timeCol.scrollTop = el.grid.scrollTop;
    el.daysHeader.scrollLeft = el.grid.scrollLeft;
  }, { passive: false });

  el.btnNew.addEventListener("click", () => {
    if (!state.users.length) return toast("Primero agrega al menos un usuario.");
    if (!state.salones.length) return toast("Primero agrega al menos un salon.");
    openModalForCreate({
      date: new Date(),
      start: "09:00",
      end: "10:00",
      salon: selectedSalon,
    });
  });

  if (el.btnOpenLeads) {
    el.btnOpenLeads.addEventListener("click", () => {
      closeSettingsPanel();
      openLeadsModal();
    });
  }
  if (el.btnLeadsClose) {
    el.btnLeadsClose.addEventListener("click", closeLeadsModal);
  }
  if (el.leadsBackdrop) {
    bindSafeBackdropClose(el.leadsBackdrop, closeLeadsModal);
  }
  if (el.leadsSearch) {
    el.leadsSearch.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        renderLeadsResults(String(el.leadsSearch.value || "").trim());
        return;
      }
      if (e.key === "Escape") closeLeadsModal();
    });
  }
  const refreshLeadsWithCurrentTerm = () => {
    renderLeadsResults(String(el.leadsSearch?.value || "").trim());
  };
  [el.leadsStatusFilter, el.leadsSellerFilter, el.leadsDateFrom, el.leadsDateTo].forEach((control) => {
    if (!control) return;
    control.addEventListener("change", refreshLeadsWithCurrentTerm);
  });
  if (el.btnLeadsReset) {
    el.btnLeadsReset.addEventListener("click", () => {
      resetLeadsFilters();
      refreshLeadsWithCurrentTerm();
    });
  }
  if (el.leadsBody) {
    el.leadsBody.addEventListener("click", (e) => {
      const btn = e.target.closest?.("[data-lead-event-id]");
      if (!btn) return;
      const eventId = String(btn.dataset.leadEventId || "").trim();
      if (!eventId) return;
      closeLeadsModal();
      revealEventInCalendar(eventId);
    });
  }

  if (el.btnFindEvent) {
    el.btnFindEvent.addEventListener("click", () => {
      closeBlockingOverlaysForNavigation();
      closeSettingsPanel();
      openEventFinderModal();
    });
  }
  if (el.btnEventFinderClose) {
    el.btnEventFinderClose.addEventListener("click", closeEventFinderModal);
  }
  if (el.eventFinderBackdrop) {
    bindSafeBackdropClose(el.eventFinderBackdrop, closeEventFinderModal);
  }
  if (el.eventFinderSearch) {
    el.eventFinderSearch.addEventListener("keydown", (e) => {
      if (e.key === "Enter") {
        e.preventDefault();
        const term = String(el.eventFinderSearch.value || "").trim();
        if (!term) {
          renderEventFinderIdle();
          return;
        }
        renderEventFinderResults(term);
        return;
      }
      if (e.key === "Escape") closeEventFinderModal();
    });
  }
  if (el.eventFinderBody) {
    el.eventFinderBody.addEventListener("click", (e) => {
      const btn = e.target.closest?.("[data-find-event-id]");
      if (!btn) return;
      const eventId = String(btn.dataset.findEventId || "").trim();
      if (!eventId) return;
      closeEventFinderModal();
      revealEventInCalendar(eventId);
    });
  }

  const handleCloseModalClick = (e) => {
    e.preventDefault();
    e.stopPropagation();
    closeModal();
  };
  el.btnClose.addEventListener("click", handleCloseModalClick);
  el.btnClose.addEventListener("pointerdown", handleCloseModalClick);
  if (el.btnDiscard) el.btnDiscard.addEventListener("click", closeModal);
  bindSafeBackdropClose(el.modalBackdrop, closeModal);

  // Validate conflicts & rules on change
  ["eventDate", "eventDateEnd", "eventStatus"].forEach(id => {
    el[id].addEventListener("change", updateRulesAndConflictsUI);
    el[id].addEventListener("input", updateRulesAndConflictsUI);
  });
  ["eventName", "eventDate", "eventDateEnd", "eventUser", "eventPax"].forEach(id => {
    if (!el[id]) return;
    el[id].addEventListener("change", () => validateReservationRequiredFields());
    el[id].addEventListener("input", () => validateReservationRequiredFields());
  });
  el.btnAddSlot.addEventListener("click", () => {
    addSlotRow({
      salon: "",
      slotPax: 0,
      dateStart: el.eventDate?.value || "",
      dateEnd: el.eventDateEnd?.value || el.eventDate?.value || "",
      startTime: "",
      endTime: "",
    });
    syncHiddenTimesFromFirstSlot();
    syncEventPaxFromSlots();
    updateRulesAndConflictsUI();
  });
  const onSlotChange = (e) => {
    if (!e.target.closest(".slotStart, .slotEnd, .slotRoom, .slotPax, .slotDateStart, .slotDateEnd")) return;
    syncHiddenTimesFromFirstSlot();
    syncEventPaxFromSlots();
    updateRulesAndConflictsUI();
    validateReservationRequiredFields();
  };
  el.slotsBody.addEventListener("input", onSlotChange);
  el.slotsBody.addEventListener("change", onSlotChange);
  el.slotsBody.addEventListener("click", async (e) => {
    const btn = e.target.closest(".slotRemoveBtn");
    if (!btn) return;
    if (el.slotsBody.querySelectorAll(".slotRow").length <= 1) {
      return toast("Debe existir al menos un bloque.");
    }
    const ok = await modernConfirm({
      title: "Eliminar bloque",
      message: "Esta seguro de eliminar este bloque de salon/horario?",
      confirmText: "Si, eliminar",
      cancelText: "No",
    });
    if (!ok) return;
    btn.closest(".slotRow")?.remove();
    syncHiddenTimesFromFirstSlot();
    syncEventPaxFromSlots();
    updateRulesAndConflictsUI();
    validateReservationRequiredFields();
  });
  el.eventDate.addEventListener("change", () => {
    if (!el.eventId.value && Array.isArray(pendingCreateDates) && pendingCreateDates.length > 1) {
      pendingCreateDates = [el.eventDate.value];
      setModernDateValue(el.eventDateEnd, el.eventDate.value);
      el.modalSubtitle.textContent = "Nuevo evento";
    }
    for (const input of Array.from(el.slotsBody.querySelectorAll(".slotDateStart"))) {
      if (!input.value) input.value = el.eventDate.value;
    }
    for (const input of Array.from(el.slotsBody.querySelectorAll(".slotDateEnd"))) {
      if (!input.value) input.value = el.eventDateEnd.value || el.eventDate.value;
    }
  });
  el.eventDateEnd.addEventListener("change", () => {
    for (const input of Array.from(el.slotsBody.querySelectorAll(".slotDateEnd"))) {
      if (!input.value) input.value = el.eventDateEnd.value || el.eventDate.value;
    }
  });

  el.eventForm.addEventListener("submit", (e) => {
    e.preventDefault();
    saveEventFromForm();
  });

  el.btnDelete.addEventListener("click", async () => {
    const id = el.eventId.value;
    if (!id) return;
    const target = state.events.find(x => x.id === id);
    const summary = target ? summarizeSeriesWindow(getEventSeries(target)) : "";
    const ok = await modernConfirm({
      title: "Eliminar reserva",
      message: `Esta seguro de eliminar esta reserva${summary ? ` (${summary})` : ""}?`,
      confirmText: "Si, eliminar",
      cancelText: "No",
    });
    if (!ok) return;
    const actorUserId = el.eventUser.value || target?.userId || "";
    removeEvent(id, actorUserId);
    closeModal();
    toast("Evento eliminado.");
  });

  el.btnCancelEvent.addEventListener("click", () => {
    const id = el.eventId.value;
    if (!id) return;
    const ev = state.events.find(x => x.id === id);
    if (!ev) return;
    const previousStatus = ev.status;
    const releasedWindows = isHardBlockingStatus(ev.status) ? buildBlockingWindowsFromEvents([ev]) : [];
    ev.status = STATUS.CANCELADO;
    appendHistoryByKey(reservationKeyFromEvent(ev), ev.userId, "Estado cambiado a Cancelado.");
    notifyReleasedCapacityForWaitingReservations({
      releasedWindows,
      sourceEvent: ev,
      reasonText: "se cancelo",
      actorUserId: ev.userId || "",
    });
    persist();
    render();
    openModalForEdit(id);
    toast(buildStatusChangeToast(previousStatus, ev.status, "Evento cancelado."));
  });

  if (el.btnSetMaintenance) {
    el.btnSetMaintenance.addEventListener("click", async () => {
      const id = String(el.eventId?.value || "").trim();
      const target = id ? state.events.find((x) => String(x.id) === id) : null;
      if (target && target.status === STATUS.MANTENIMIENTO) {
        const release = await modernConfirmReleaseMaintenance();
        if (!release?.isConfirmed) return;
        removeEvent(target.id, el.eventUser.value || target.userId || "");
        closeModal();
        await modernAlert({
          icon: "success",
          title: "Mantenimiento liberado",
          text: "El salon quedo libre nuevamente.",
        });
        return;
      }

      const validation = validateReservationRequiredFields();
      if (validation.issues.length) {
        const first = validation.issues[0];
        const extra = validation.issues.length > 1 ? ` (+${validation.issues.length - 1} pendientes)` : "";
        await modernGuideToast(`Completa: ${first}${extra}`);
        if (validation.firstInvalidEl && typeof validation.firstInvalidEl.focus === "function") {
          validation.firstInvalidEl.focus();
        }
        return;
      }

      const maintenanceBlocks = findMaintenanceBlockingReservationsFromForm();
      if (maintenanceBlocks.length) {
        await modernAlert({
          icon: "warning",
          title: "Mantenimiento bloqueado",
          text: "No puedes poner en Mantenimiento: ya hay una reserva Confirmada o Pre reserva en ese horario.",
        });
        return;
      }

      const res = await modernConfirmMaintenance();
      if (!res?.isConfirmed) return;

      el.eventStatus.value = STATUS.MANTENIMIENTO;
      applyStatusSelectTheme();
      updateRulesAndConflictsUI();
      await modernAlert({
        icon: "success",
        title: "Listo",
        text: "Estado cambiado a Mantenimiento.",
      });
    });
  }
  if (el.btnToggleHistory) {
    el.btnToggleHistory.addEventListener("click", () => {
      if (!historyTargetEventId) return;
      const nextVisible = el.historyPanel?.hidden;
      if (nextVisible) {
        const ev = state.events.find(x => x.id === historyTargetEventId);
        renderHistoryForEvent(ev || null);
      }
      setHistoryPanelVisible(!!nextVisible);
    });
  }
  if (el.btnToggleAppointments) {
    el.btnToggleAppointments.addEventListener("click", () => {
      if (!historyTargetEventId) return;
      const nextVisible = el.appointmentPanel?.hidden;
      if (nextVisible) {
        const ev = state.events.find(x => x.id === historyTargetEventId);
        renderAppointmentsForEvent(ev || null);
      }
      setAppointmentsPanelVisible(!!nextVisible);
    });
  }
  if (el.btnBackToForm) {
    el.btnBackToForm.addEventListener("click", () => {
      backToEventForm();
    });
  }
  if (el.btnTopbarReminders) {
    el.btnTopbarReminders.addEventListener("click", (e) => {
      e.stopPropagation();
      if (!el.topbarReminderPanel) return;
      const willOpen = !!el.topbarReminderPanel.hidden;
      refreshTopbarReminders();
      el.topbarReminderPanel.hidden = !willOpen;
      el.btnTopbarReminders.setAttribute("aria-expanded", willOpen ? "true" : "false");
    });
  }
  if (el.topbarReminderList) {
    el.topbarReminderList.addEventListener("click", (e) => {
      const dismissBtn = e.target.closest("[data-dismiss-global-notification-id]");
      if (dismissBtn) {
        e.preventDefault();
        e.stopPropagation();
        const notificationId = String(dismissBtn.dataset.dismissGlobalNotificationId || "").trim();
        if (!notificationId) return;
        const removed = dismissGlobalNotificationForSession(notificationId, { persistRemote: true });
        if (removed) {
          refreshTopbarReminders();
          toast("Notificacion quitada.");
        }
        return;
      }
      const item = e.target.closest(".topbarReminderItem");
      if (!item) return;
      const eventId = String(item.dataset.eventId || "").trim();
      if (!eventId) return;
      closeTopbarReminderPanel();
      openModalForEdit(eventId);
    });
  }
  if (el.btnAddAppointment) {
    el.btnAddAppointment.addEventListener("click", () => {
      if (!el.eventId.value) return;
      openAppointmentModal(el.eventId.value);
    });
  }

  el.btnQuoteEvent.addEventListener("click", async () => {
    const id = String(el.eventId?.value || "").trim();
    if (!id) return;
    // Evita que el modal de Edicion quede encima y oculte la pantalla de cotizacion.
    closeModal();
    try {
      await openQuoteModal(id);
    } catch (_) {
      toast("No se pudo abrir la cotizacion. Intenta nuevamente.");
    }
  });
  if (el.btnMenuMontaje) {
    el.btnMenuMontaje.addEventListener("click", () => {
      openMenuMontajeModal();
    });
  }
  if (el.btnQuoteAdvances) {
    el.btnQuoteAdvances.addEventListener("click", () => {
      openQuoteAdvanceModal();
    });
  }
  if (el.btnMenuMontajeSelectable) {
    el.btnMenuMontajeSelectable.addEventListener("click", () => {
      openMenuMontajeSelectableModal().catch(() => {
        toast("No se pudo abrir Menu & Montaje (Listas).");
      });
    });
  }
  if (el.btnQuotePrintTemplate) {
    el.btnQuotePrintTemplate.addEventListener("click", async () => {
      await printSelectedQuoteTemplate();
    });
  }
  if (el.btnQuotePaymentAdd) {
    el.btnQuotePaymentAdd.addEventListener("click", () => {
      addQuotePaymentTypeFromPicker();
    });
  }
  if (el.btnQuotePaymentClear) {
    el.btnQuotePaymentClear.addEventListener("click", () => {
      clearQuotePaymentTypesOnForm();
    });
  }
  if (el.btnQuoteAdvanceClose) {
    el.btnQuoteAdvanceClose.addEventListener("click", closeQuoteAdvanceModal);
  }
  if (el.btnQuoteAdvanceDone) {
    el.btnQuoteAdvanceDone.addEventListener("click", () => {
      if (el.accountingReportBackdrop && !el.accountingReportBackdrop.hidden) renderAccountingReportTable();
      closeQuoteAdvanceModal();
    });
  }
  if (el.btnQuoteAdvanceAdd) {
    el.btnQuoteAdvanceAdd.addEventListener("click", () => {
      addQuoteAdvanceFromForm();
    });
  }
  if (el.quoteAdvanceEvidence) {
    el.quoteAdvanceEvidence.addEventListener("change", () => {
      const file = el.quoteAdvanceEvidence.files?.[0] || null;
      if (el.quoteAdvanceEvidenceHint) {
        el.quoteAdvanceEvidenceHint.textContent = file ? `Archivo seleccionado: ${file.name}` : "Sin archivo adjunto";
      }
    });
  }
  if (el.quoteAdvanceBackdrop) {
    el.quoteAdvanceBackdrop.addEventListener("click", (e) => {
      const editBtn = e.target.closest(".quoteAdvanceEditBtn");
      if (editBtn) {
        startEditQuoteAdvance(editBtn.dataset.advanceId);
        return;
      }
      const btn = e.target.closest(".quoteAdvanceRemoveBtn");
      if (!btn) return;
      removeQuoteAdvanceById(btn.dataset.advanceId);
    });
  }
  if (el.accountingReportBody) {
    el.accountingReportBody.addEventListener("click", (e) => {
      const toggleBtn = e.target.closest("[data-accounting-toggle-company]");
      if (toggleBtn) {
        const key = String(toggleBtn.dataset.accountingToggleCompany || "").trim();
        if (!key) return;
        if (accountingReportExpandedAccounts.has(key)) accountingReportExpandedAccounts.delete(key);
        else accountingReportExpandedAccounts.add(key);
        renderAccountingReportTable();
        return;
      }
      const statementBtn = e.target.closest("[data-accounting-open-statement]");
      if (statementBtn) {
        openAccountingStatementModal(statementBtn.dataset.accountingOpenStatement);
        return;
      }
      const quoteBtn = e.target.closest("[data-accounting-open-quote]");
      if (quoteBtn) {
        openLatestQuoteFromAccounting(quoteBtn.dataset.accountingOpenQuote).catch(() => {
          toast("No se pudo abrir la cotizacion.");
        });
        return;
      }
      const btn = e.target.closest("[data-accounting-apply-payment]");
      if (!btn) return;
      applyPaymentFromAccountingReport(btn.dataset.accountingApplyPayment).catch(() => {
        toast("No se pudo abrir la venta para aplicar pagos.");
      });
    });
  }
  if (el.accountStatementBody) {
    el.accountStatementBody.addEventListener("click", (e) => {
      const quoteBtn = e.target.closest("[data-accounting-open-quote]");
      if (quoteBtn) {
        openLatestQuoteFromAccounting(quoteBtn.dataset.accountingOpenQuote).catch(() => {
          toast("No se pudo abrir la cotizacion.");
        });
        return;
      }
      const btn = e.target.closest("[data-accounting-apply-payment]");
      if (!btn) return;
      applyPaymentFromAccountingReport(btn.dataset.accountingApplyPayment).catch(() => {
        toast("No se pudo abrir la venta para aplicar pagos.");
      });
    });
  }  if (el.quoteAdvanceBackdrop) {
    bindSafeBackdropClose(el.quoteAdvanceBackdrop, closeQuoteAdvanceModal);
  }
  if (el.btnMenuMontajeClose) {
    el.btnMenuMontajeClose.addEventListener("click", closeMenuMontajeModal);
  }
  if (el.menuMontajeBackdrop) {
    bindSafeBackdropClose(el.menuMontajeBackdrop, closeMenuMontajeModal);
    el.menuMontajeBackdrop.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-mm-snippet]");
      if (!btn) return;
      e.preventDefault();
      const tools = btn.closest(".mmTextTools");
      const targetId = String(btn.dataset.mmTarget || tools?.dataset.mmTarget || "").trim();
      if (!targetId) return;
      const target = document.getElementById(targetId);
      if (!(target instanceof HTMLTextAreaElement)) return;
      const snippet = getMenuMontajeSnippet(btn.dataset.mmSnippet);
      if (!snippet) return;
      insertTextAtCursor(target, snippet);
      target.focus();
      updateMenuMontajeCounters();
    });
  }
  if (el.mmMenuDescription) {
    el.mmMenuDescription.addEventListener("input", updateMenuMontajeCounters);
  }
  if (el.mmMontajeDescription) {
    el.mmMontajeDescription.addEventListener("input", updateMenuMontajeCounters);
  }
  if (el.mmDateSalonSelect) {
    el.mmDateSalonSelect.addEventListener("change", () => {
      menuMontajeSelectedKey = String(el.mmDateSalonSelect.value || "").trim();
      loadMenuMontajeFormByKey(menuMontajeSelectedKey);
    });
  }
  if (el.btnMenuMontajeLoadVersion) {
    el.btnMenuMontajeLoadVersion.addEventListener("click", () => {
      const v = Number(el.mmVersionSelect?.value || 0);
      if (!Number.isFinite(v) || v <= 0) return toast("Version invalida.");
      loadMenuMontajeVersion(v);
    });
  }
  if (el.btnMenuMontajeSave) {
    el.btnMenuMontajeSave.addEventListener("click", async () => {
      await saveMenuMontajeFromModal({ updateCurrentVersion: false });
    });
  }
  if (el.btnMenuMontajeSaveCurrent) {
    el.btnMenuMontajeSaveCurrent.addEventListener("click", async () => {
      await saveMenuMontajeFromModal({ updateCurrentVersion: true });
    });
  }
  if (el.btnMenuMontajePrintDay) {
    el.btnMenuMontajePrintDay.addEventListener("click", () => {
      printMenuMontajeByDay();
    });
  }
  if (el.btnMenuMontajeSelectableClose) {
    el.btnMenuMontajeSelectableClose.addEventListener("click", closeMenuMontajeSelectableModal);
  }
  if (el.menuMontajeSelectableBackdrop) {
    bindSafeBackdropClose(el.menuMontajeSelectableBackdrop, closeMenuMontajeSelectableModal);
    el.menuMontajeSelectableBackdrop.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-mm-snippet]");
      if (!btn) return;
      e.preventDefault();
      const tools = btn.closest(".mmTextTools");
      const targetId = String(btn.dataset.mmTarget || tools?.dataset.mmTarget || "").trim();
      if (!targetId) return;
      const target = document.getElementById(targetId);
      if (!(target instanceof HTMLTextAreaElement)) return;
      const snippet = getMenuMontajeSnippet(btn.dataset.mmSnippet);
      if (!snippet) return;
      insertTextAtCursor(target, snippet);
      target.focus();
    });
  }
  if (el.mmsProtein) {
    el.mmsProtein.addEventListener("change", () => {
      refreshMmsByProteinPreparation({ preserveSelection: false }).catch(() => {
        toast("No se pudieron cargar preparaciones del menu.");
      });
    });
  }
  if (el.mmsPreparation) {
    el.mmsPreparation.addEventListener("change", () => {
      refreshMmsByProteinPreparation({ preserveSelection: true }).catch(() => {
        toast("No se pudieron cargar sugerencias del menu.");
      });
    });
  }
  if (el.mmsStageTabs) {
    el.mmsStageTabs.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-mms-stage]");
      if (!btn) return;
      setMmsStage(String(btn.dataset.mmsStage || "plato"));
    });
  }
  if (el.mmsPrimaryTabs) {
    el.mmsPrimaryTabs.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-mms-primary]");
      if (!btn) return;
      setMmsPrimaryMode(String(btn.dataset.mmsPrimary || "menu"));
    });
  }
  if (el.btnMmsStageMoreOptions) {
    el.btnMmsStageMoreOptions.addEventListener("click", () => {
      if (mmsCurrentStage === "salsa" || mmsCurrentStage === "guarnicion") {
        mmsShowAllGuarniciones = !mmsShowAllGuarniciones;
      } else if (mmsCurrentStage === "postre") {
        mmsShowAllPostres = !mmsShowAllPostres;
      } else {
        return;
      }
      renderMmsStageOptions();
    });
  }
  if (el.btnMmsStageCancelSelection) {
    el.btnMmsStageCancelSelection.addEventListener("click", () => {
      cancelMmsCurrentStageSelection();
    });
  }
  if (el.btnMmsOpenCatalog) {
    el.btnMmsOpenCatalog.addEventListener("click", async () => {
      await openMenuCatalogManagerModal("plato_fuerte");
    });
  }
  if (el.mmsStageFilter) {
    el.mmsStageFilter.addEventListener("input", () => {
      renderMmsStageOptions();
    });
  }
  if (el.mmsStageOptions) {
    el.mmsStageOptions.addEventListener("click", async (e) => {
      const btn = e.target.closest("[data-mms-quick-kind]");
      if (!btn) return;
      await handleMmsStageOptionClick(btn.dataset.mmsQuickKind, btn.dataset.mmsQuickId);
    });
  }
  if (el.btnMmsAddBebida) {
    el.btnMmsAddBebida.addEventListener("click", () => {
      addMmsBebidaFromInput();
    });
  }
  if (el.mmsBebidaInput) {
    el.mmsBebidaInput.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      addMmsBebidaFromInput();
    });
  }
  if (el.btnMmsToggleGuarnicionesGlobal) {
    el.btnMmsToggleGuarnicionesGlobal.addEventListener("click", () => {
      mmsShowAllGuarniciones = !mmsShowAllGuarniciones;
      if (el.btnMmsToggleGuarnicionesGlobal) {
        el.btnMmsToggleGuarnicionesGlobal.textContent = mmsShowAllGuarniciones ? "Ocultar globales" : "Mas guarniciones";
      }
      renderMmsQuickSelectors();
    });
  }
  if (el.btnMmsTogglePostresGlobal) {
    el.btnMmsTogglePostresGlobal.addEventListener("click", () => {
      mmsShowAllPostres = !mmsShowAllPostres;
      if (el.btnMmsTogglePostresGlobal) {
        el.btnMmsTogglePostresGlobal.textContent = mmsShowAllPostres ? "Ocultar globales" : "Mas postres";
      }
      renderMmsQuickSelectors();
    });
  }
  if (el.mmsGuarnicionFilter) {
    el.mmsGuarnicionFilter.addEventListener("input", () => {
      renderMmsQuickSelectors();
    });
  }
  if (el.mmsPostreFilter) {
    el.mmsPostreFilter.addEventListener("input", () => {
      renderMmsQuickSelectors();
    });
  }
  if (el.mmsGuarnicionesQuickSuggested) {
    el.mmsGuarnicionesQuickSuggested.addEventListener("click", async (e) => {
      const btn = e.target.closest("[data-mms-quick-kind='guarnicion']");
      if (!btn) return;
      await toggleMmsQuickItem("guarnicion", btn.dataset.mmsQuickId);
    });
  }
  if (el.mmsGuarnicionesQuickGlobal) {
    el.mmsGuarnicionesQuickGlobal.addEventListener("click", async (e) => {
      const btn = e.target.closest("[data-mms-quick-kind='guarnicion']");
      if (!btn) return;
      await toggleMmsQuickItem("guarnicion", btn.dataset.mmsQuickId);
    });
  }
  if (el.mmsPostresQuickSuggested) {
    el.mmsPostresQuickSuggested.addEventListener("click", async (e) => {
      const btn = e.target.closest("[data-mms-quick-kind='postre']");
      if (!btn) return;
      await toggleMmsQuickItem("postre", btn.dataset.mmsQuickId);
    });
  }
  if (el.mmsPostresQuickGlobal) {
    el.mmsPostresQuickGlobal.addEventListener("click", async (e) => {
      const btn = e.target.closest("[data-mms-quick-kind='postre']");
      if (!btn) return;
      await toggleMmsQuickItem("postre", btn.dataset.mmsQuickId);
    });
  }
  if (el.mmsComandaPreview) {
    el.mmsComandaPreview.addEventListener("click", (e) => {
      const plateTag = e.target.closest(".mmsComandaTag[data-mms-tag-kind='plato']");
      if (plateTag && !e.target.closest("button")) {
        loadMmsLineItemIntoEditor(String(plateTag.dataset.mmsTagId || ""));
        return;
      }
      const qtyBtn = e.target.closest("[data-mms-qty-kind]");
      if (qtyBtn) {
        const kind = String(qtyBtn.dataset.mmsQtyKind || "");
        const action = String(qtyBtn.dataset.mmsQtyAction || "");
        const idRaw = String(qtyBtn.dataset.mmsQtyId || "").trim();
        const id = Number(idRaw || 0);
        if (kind === "plato" && idRaw) {
          const target = (Array.isArray(mmsSelectedPlatoItems) ? mmsSelectedPlatoItems : []).find((item) => String(item?.key || "") === idRaw);
          if (target) {
            const current = Math.max(1, Math.floor(Number(target.qty || 1)));
            if (action === "dec") {
              if (current <= 1) {
                removeMmsComandaItem("plato", idRaw);
                return;
              }
              updateMmsPlatoItemQtyByKey(idRaw, current - 1);
            } else if (action === "inc") {
              updateMmsPlatoItemQtyByKey(idRaw, current + 1);
            }
            refreshMmsDescriptionAuto();
            renderMmsSelectionSummary();
            renderMmsComandaPreview();
            return;
          }
        }
        if ((kind === "guarnicion" || kind === "postre" || kind === "bebida") && Number.isFinite(id) && id > 0) {
          const current = getMmsItemQty(kind, id);
          if (action === "dec") {
            if (current <= 1) {
              removeMmsComandaItem(kind, id);
              return;
            }
            setMmsItemQty(kind, id, current - 1);
          } else if (action === "inc") {
            setMmsItemQty(kind, id, current + 1);
          }
          refreshMmsDescriptionAuto();
          renderMmsSelectionSummary();
          renderMmsComandaPreview();
        }
        return;
      }
      const btn = e.target.closest("[data-mms-remove-kind]");
      if (!btn) return;
      removeMmsComandaItem(btn.dataset.mmsRemoveKind, btn.dataset.mmsRemoveId);
    });
  }
  if (el.btnMmsUseSuggested) {
    el.btnMmsUseSuggested.addEventListener("click", () => {
      useMmsSuggestedSelections();
    });
  }
  if (el.btnMmsClearMenuSelection) {
    el.btnMmsClearMenuSelection.addEventListener("click", () => {
      clearMmsMenuSelections();
    });
  }
  if (el.btnMmsMenuAppend) {
    el.btnMmsMenuAppend.addEventListener("click", () => {
      applyMmsMenuBuilder("append");
      toast("Bloque de menu agregado.");
    });
  }
  if (el.btnMmsMenuReplace) {
    el.btnMmsMenuReplace.addEventListener("click", () => {
      applyMmsMenuBuilder("replace");
      toast("Descripcion de menu reemplazada.");
    });
  }
  if (el.mmsMenuSection) {
    el.mmsMenuSection.addEventListener("change", () => {
      refreshMmsDescriptionAuto();
    });
  }
  if (el.btnMmsMenuSectionAdd) {
    el.btnMmsMenuSectionAdd.addEventListener("click", () => {
      addMmsMenuSectionFromInput();
    });
  }
  if (el.mmsMenuSectionInput) {
    el.mmsMenuSectionInput.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      addMmsMenuSectionFromInput();
    });
  }
  if (el.mmsDateSalonSelect) {
    el.mmsDateSalonSelect.addEventListener("change", () => {
      menuMontajeSelectableSelectedKey = String(el.mmsDateSalonSelect.value || "").trim();
      loadMmsFormByKey(menuMontajeSelectableSelectedKey).catch(() => {
        toast("No se pudo cargar el detalle de menu.");
      });
    });
  }
  if (el.btnMmsLoadVersion) {
    el.btnMmsLoadVersion.addEventListener("click", () => {
      const v = Number(el.mmsVersionSelect?.value || 0);
      if (!Number.isFinite(v) || v <= 0) return toast("Version invalida.");
      loadMmsVersion(v).catch(() => {
        toast("No se pudo cargar la version.");
      });
    });
  }
  if (el.btnMmsSave) {
    el.btnMmsSave.addEventListener("click", async () => {
      await saveMenuMontajeSelectableFromModal({ updateCurrentVersion: false });
    });
  }
  if (el.btnMmsSaveCurrent) {
    el.btnMmsSaveCurrent.addEventListener("click", async () => {
      await saveMenuMontajeSelectableFromModal({ updateCurrentVersion: true });
    });
  }
  if (el.btnMmsMontajeClear) {
    el.btnMmsMontajeClear.addEventListener("click", () => {
      clearMmsMontajeSelections();
    });
  }
  if (el.btnMmsMontajeAppend) {
    el.btnMmsMontajeAppend.addEventListener("click", () => {
      applyMmsMontajeBuilder("append");
      toast("Bloque de montaje agregado.");
    });
  }
  if (el.btnMmsMontajeReplace) {
    el.btnMmsMontajeReplace.addEventListener("click", () => {
      applyMmsMontajeBuilder("replace");
      toast("Descripcion de montaje reemplazada.");
    });
  }
  if (el.btnMmsPrintDay) {
    el.btnMmsPrintDay.addEventListener("click", () => {
      printMenuMontajeByDay();
    });
  }
  if (el.mmsMontajeTipo) {
    el.mmsMontajeTipo.addEventListener("change", () => {
      refreshMmsDescriptionAuto();
    });
  }
  if (el.mmsMenuTitle) {
    el.mmsMenuTitle.addEventListener("input", () => {
      renderMmsSelectionSummary();
    });
  }
  if (el.mmsMenuQty) {
    el.mmsMenuQty.addEventListener("input", () => {
      const qty = Math.max(1, Math.floor(Number(el.mmsMenuQty?.value || 1)));
      mmsPlatoQty = qty;
      if (el.mmsMenuQtyVisible) el.mmsMenuQtyVisible.value = String(qty);
    });
  }
  if (el.mmsMenuQtyVisible) {
    el.mmsMenuQtyVisible.addEventListener("input", () => {
      const qty = Math.max(1, Math.floor(Number(el.mmsMenuQtyVisible?.value || 1)));
      mmsPlatoQty = qty;
      if (el.mmsMenuQty) el.mmsMenuQty.value = String(qty);
    });
  }
  if (el.mmsMenuDescription) {
    el.mmsMenuDescription.addEventListener("input", () => {
      renderMmsSelectionSummary();
    });
  }
  if (el.mmsMontajeDescription) {
    el.mmsMontajeDescription.addEventListener("input", () => {
      renderMmsSelectionSummary();
    });
  }
  const bindMirror = (left, right) => {
    if (!left || !right) return;
    const resolveKind = () => {
      if (
        (left === el.mmsGuarnicionesSuggested && right === el.mmsGuarnicionesAll)
        || (left === el.mmsGuarnicionesAll && right === el.mmsGuarnicionesSuggested)
      ) return "guarnicion";
      if (
        (left === el.mmsPostresSuggested && right === el.mmsPostresAll)
        || (left === el.mmsPostresAll && right === el.mmsPostresSuggested)
      ) return "postre";
      return "";
    };
    const syncQtyFromChecklist = async (origin, mirror, check) => {
      const kind = resolveKind();
    if (!kind) return true;
    const id = Number(check?.value || 0);
    if (!Number.isFinite(id) || id <= 0) return true;
    if (check.checked) {
      const label = String(check.closest("label")?.textContent || check.dataset.label || "Producto").trim();
      if (kind === "guarnicion") {
        notifyMmsSelectionAdded(kind, id, label, 1);
      } else {
        const qty = await requestMmsItemQty(label, getMmsItemQty(kind, id), "postres");
        if (qty === null) {
          check.checked = false;
          syncMirrorChecklistValue(origin, mirror, check);
          return false;
        }
        setMmsItemQty(kind, id, qty);
        notifyMmsSelectionAdded(kind, id, label, qty);
      }
    } else {
      deleteMmsItemQty(kind, id);
    }
      if (kind === "guarnicion") syncMmsGuarnicionQtyWithSelection();
      else syncMmsPostreQtyWithSelection();
      return true;
    };
    left.addEventListener("change", async (e) => {
      const check = e.target.closest("input[type='checkbox']");
      if (!check) return;
      syncMirrorChecklistValue(left, right, check);
      if (!(await syncQtyFromChecklist(left, right, check))) return;
      refreshMmsDescriptionAuto();
      renderMmsSelectionSummary();
      renderMmsQuickSelectors();
      renderMmsComandaPreview();
    });
    right.addEventListener("change", async (e) => {
      const check = e.target.closest("input[type='checkbox']");
      if (!check) return;
      syncMirrorChecklistValue(right, left, check);
      if (!(await syncQtyFromChecklist(right, left, check))) return;
      refreshMmsDescriptionAuto();
      renderMmsSelectionSummary();
      renderMmsQuickSelectors();
      renderMmsComandaPreview();
    });
  };
  bindMirror(el.mmsGuarnicionesSuggested, el.mmsGuarnicionesAll);
  bindMirror(el.mmsPostresSuggested, el.mmsPostresAll);
  if (el.mmsComentariosAll) {
    el.mmsComentariosAll.addEventListener("change", () => {
      refreshMmsDescriptionAuto();
    });
  }
  if (el.mmsMontajeAdicionales) {
    el.mmsMontajeAdicionales.addEventListener("change", () => {
      refreshMmsDescriptionAuto();
    });
  }
  if (el.mmsComentarioLibre) {
    el.mmsComentarioLibre.addEventListener("input", () => {
      refreshMmsDescriptionAuto();
    });
  }
  if (el.mmsPlatoDescripcion) {
    el.mmsPlatoDescripcion.addEventListener("input", () => {
      if (el.mmsComentarioLibre) el.mmsComentarioLibre.value = String(el.mmsPlatoDescripcion.value || "");
      refreshMmsDescriptionAuto();
    });
  }

  // User modal
  el.btnAddUser.addEventListener("click", () => openUserModal());
  el.btnUserClose.addEventListener("click", (e) => {
    e.preventDefault();
    e.stopPropagation();
    closeUserModal();
  });
  if (el.btnUserDiscard) {
    el.btnUserDiscard.addEventListener("click", (e) => {
      e.preventDefault();
      closeUserModal();
    });
  }
  bindSafeBackdropClose(el.userBackdrop, closeUserModal);
  if (el.btnUserGoalAdd) {
    el.btnUserGoalAdd.addEventListener("click", () => {
      upsertUserMonthlyGoalDraft();
    });
  }
  if (el.userGoalsBody) {
    el.userGoalsBody.addEventListener("click", async (e) => {
      const btn = e.target.closest("[data-goal-action]");
      if (!btn) return;
      const action = String(btn.dataset.goalAction || "");
      const month = String(btn.dataset.goalMonth || "");
      if (!month) return;
      if (action === "remove") {
        const ok = await modernConfirm({
          title: "Eliminar meta mensual",
          message: `Esta seguro de eliminar la meta del mes ${month}?`,
          confirmText: "Si, eliminar",
          cancelText: "No",
        });
        if (!ok) return;
        userMonthlyGoalsDraft = userMonthlyGoalsDraft.filter((g) => String(g.month) !== month);
        if (editingUserGoalMonth === month) clearUserGoalEditorFields();
        renderUserMonthlyGoalsDraft();
        toast("Meta mensual eliminada del borrador.");
        return;
      }
      const target = userMonthlyGoalsDraft.find((g) => String(g.month) === month);
      if (!target) return;
      editingUserGoalMonth = month;
      if (el.userGoalMonth) el.userGoalMonth.value = target.month;
      if (el.userGoalAmount) el.userGoalAmount.value = String(target.amount);
      if (el.btnUserGoalAdd) el.btnUserGoalAdd.textContent = "Guardar meta";
    });
  }
  if (el.userEditSelect) {
    el.userEditSelect.addEventListener("change", () => {
      const selectedId = String(el.userEditSelect.value || "").trim();
      if (!selectedId) {
        resetUserModalForm();
        return;
      }
      loadUserInModal(selectedId);
    });
  }
  if (el.userSignature) {
    el.userSignature.addEventListener("change", async () => {
      const file = el.userSignature.files?.[0] || null;
      if (!file) {
        userSignatureNormalizedDataUrl = "";
        const existing = userModalEditingId
          ? (state.users || []).map(normalizeUserRecord).find((u) => String(u.id) === userModalEditingId)?.signatureDataUrl || ""
          : "";
        renderUserSignaturePreview(existing);
        return;
      }
      const isPng = /image\/png/i.test(file.type) || /\.png$/i.test(file.name || "");
      const isJpg = /image\/jpeg/i.test(file.type) || /\.(jpe?g)$/i.test(file.name || "");
      if (!isPng && !isJpg) {
        toast("La firma debe ser JPG o PNG.");
        el.userSignature.value = "";
        userSignatureNormalizedDataUrl = "";
        renderUserSignaturePreview("");
        return;
      }
      const dataUrl = await readImageFileAsDataUrl(file);
      const normalized = await normalizeSignatureDataUrlForContract(dataUrl);
      userSignatureNormalizedDataUrl = String(normalized || dataUrl || "").trim();
      renderUserSignaturePreview(userSignatureNormalizedDataUrl);
    });
  }
  if (el.btnUserDisable) {
    el.btnUserDisable.addEventListener("click", async () => {
      if (!userModalEditingId) return;
      const idx = (state.users || []).findIndex((u) => String(u.id) === userModalEditingId);
      if (idx < 0) return;
      const target = normalizeUserRecord(state.users[idx]);
      if (authSession.userId && authSession.userId === target.id) {
        toast("No puedes inhabilitar el usuario con sesion activa.");
        return;
      }
      target.active = false;
      state.users[idx] = target;
      persist();
      renderUsersSelect();
      await loadLoginUsers().catch(() => { });
      closeUserModal();
      toast("Usuario inhabilitado.");
    });
  }

  el.userForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    const fullName = String(el.userFullName?.value || "").trim();
    const username = String(el.userUsername?.value || "").trim();
    const email = String(el.userEmail?.value || "").trim();
    const phone = String(el.userPhone?.value || "").trim();
    const role = normalizeUserRole(el.userRole?.value || USER_ROLE.SELLER);
    const password = String(el.userPassword?.value || "").trim();
    const signatureFile = el.userSignature?.files?.[0] || null;
    const avatarFile = el.userAvatar?.files?.[0] || null;
    const isEdit = !!userModalEditingId;
    const editIndex = isEdit ? (state.users || []).findIndex((u) => String(u.id) === userModalEditingId) : -1;

    if (!fullName || !username || !email || !phone || (!isEdit && !password)) {
      return toast("Completa nombre, usuario, correo, telefono y contrasena.");
    }
    if (!isValidEmail(email)) return toast("Correo de usuario invalido.");
    const usernameExists = (state.users || []).some(
      (u, idx) => String(u.username || "").toLowerCase() === username.toLowerCase() && idx !== editIndex
    );
    if (usernameExists) return toast("Ese nombre de usuario ya existe.");
    const emailExists = (state.users || []).some(
      (u, idx) => String(u.email || "").toLowerCase() === email.toLowerCase() && idx !== editIndex
    );
    if (emailExists) return toast("Ese correo ya existe.");

    let signatureDataUrl = "";
    if (signatureFile) {
      const isPng = /image\/png/i.test(signatureFile.type) || /\.png$/i.test(signatureFile.name || "");
      const isJpg = /image\/jpeg/i.test(signatureFile.type) || /\.(jpe?g)$/i.test(signatureFile.name || "");
      if (!isPng && !isJpg) return toast("La firma debe ser JPG o PNG.");
      signatureDataUrl = String(userSignatureNormalizedDataUrl || "").trim();
      if (!signatureDataUrl) {
        const rawSignature = await readImageFileAsDataUrl(signatureFile);
        signatureDataUrl = await normalizeSignatureDataUrlForContract(rawSignature);
      }
      const signatureAnalysis = await analyzeSignatureDataUrl(signatureDataUrl);
      const warn = getSignatureWhitespaceWarning(signatureAnalysis);
      if (warn) toast(`Aviso firma: ${warn}`);
    }

    let avatarDataUrl = "";
    if (avatarFile) {
      const isPng = /image\/png/i.test(avatarFile.type) || /\.png$/i.test(avatarFile.name || "");
      const isJpg = /image\/jpeg/i.test(avatarFile.type) || /\.(jpe?g)$/i.test(avatarFile.name || "");
      if (!isPng && !isJpg) return toast("El avatar debe ser JPG o PNG.");
      avatarDataUrl = await readImageFileAsDataUrl(avatarFile);
    }
    const active = !!el.userActive?.checked;
    const salesTargetEnabled = !!el.userSalesTargetEnabled?.checked;
    const monthlyGoals = (userMonthlyGoalsDraft || [])
      .map((g) => ({ month: String(g.month || "").trim(), amount: Math.max(0, Number(g.amount || 0)) }))
      .filter((g) => /^\d{4}-\d{2}$/.test(g.month) && Number.isFinite(g.amount) && g.amount > 0)
      .sort((a, b) => a.month.localeCompare(b.month));
    if (salesTargetEnabled && !monthlyGoals.length) {
      return toast("Si influye en meta, agrega al menos una meta mensual.");
    }
    if (isEdit && !active && authSession.userId && authSession.userId === userModalEditingId) {
      return toast("No puedes inhabilitar el usuario con sesion activa.");
    }
  if (isEdit && editIndex >= 0) {
    const previous = normalizeUserRecord(state.users[editIndex]);
    const nextUser = {
      ...previous,
      name: fullName,
      fullName,
      role,
      username,
      email,
        phone,
        password: password || previous.password,
        signatureDataUrl: signatureDataUrl || previous.signatureDataUrl,
        avatarDataUrl: avatarDataUrl || previous.avatarDataUrl,
      active,
      salesTargetEnabled,
      monthlyGoals,
    };
    if (areUsersEquivalent(previous, nextUser)) {
      return toast("Sin cambios detectados en usuario.");
    }
    state.users[editIndex] = nextUser;
    if (authSession.userId && authSession.userId === state.users[editIndex].id) {
      authSession.fullName = state.users[editIndex].fullName || state.users[editIndex].name || authSession.fullName;
      authSession.username = state.users[editIndex].username || authSession.username;
      authSession.avatarDataUrl = state.users[editIndex].avatarDataUrl || "";
        authSession.signatureDataUrl = state.users[editIndex].signatureDataUrl || "";
        renderTopbarWelcome();
      }
    } else {
      state.users.push({
        id: uid(),
        name: fullName,
        fullName,
        role,
        username,
        email,
        phone,
        password,
        signatureDataUrl,
        avatarDataUrl,
        active: true,
        salesTargetEnabled,
        monthlyGoals,
      });
    }
    persist();
    renderUsersSelect();
    const targetId = isEdit && editIndex >= 0 ? state.users[editIndex].id : state.users[state.users.length - 1].id;
    syncEnhancedSelectValue(el.eventUser, targetId);
    await loadLoginUsers();
    closeUserModal();
    toast(isEdit ? "Usuario actualizado." : "Usuario agregado.");
  });

  el.btnQuoteClose.addEventListener("click", closeQuoteModal);
  if (el.btnQuoteDiscard) el.btnQuoteDiscard.addEventListener("click", closeQuoteModal);
  bindSafeBackdropClose(el.quoteBackdrop, closeQuoteModal);

  el.btnAddCompany.addEventListener("click", () => openCompanyModal(""));
  if (el.btnOpenServiceCreate) el.btnOpenServiceCreate.addEventListener("click", openServiceModal);
  if (el.companyEditSelect) {
    el.companyEditSelect.addEventListener("change", () => {
      loadCompanyInModal(el.companyEditSelect.value || "");
    });
  }
  if (el.btnCompanyDisable) {
    el.btnCompanyDisable.addEventListener("click", () => {
      const id = String(editingCompanyId || "").trim();
      if (!id) return toast("Selecciona una empresa para cambiar su estado.");
      if (isCompanyDisabled(id)) enableCompany(id);
      else state.disabledCompanies = Array.from(new Set([...(state.disabledCompanies || []), id]));
      persist();
      renderCompaniesSelect(id);
      loadCompanyInModal(id);
      toast(isCompanyDisabled(id) ? "Empresa inhabilitada." : "Empresa reactivada.");
    });
  }
  el.btnCompanyClose.addEventListener("click", closeCompanyModal);
  if (el.btnCompanyDiscard) el.btnCompanyDiscard.addEventListener("click", closeCompanyModal);
  bindSafeBackdropClose(el.companyBackdrop, closeCompanyModal);
  if (el.serviceEditSelect) {
    el.serviceEditSelect.addEventListener("change", () => {
      openServiceModal(el.serviceEditSelect.value || "");
    });
  }
  if (el.btnServiceDisable) {
    el.btnServiceDisable.addEventListener("click", () => {
      const id = String(editingServiceId || "").trim();
      if (!id) return toast("Selecciona un servicio para cambiar su estado.");
      if (isServiceDisabled(id)) enableService(id);
      else state.disabledServices = Array.from(new Set([...(state.disabledServices || []), id]));
      persist();
      renderServicesList();
      openServiceModal(id);
      toast(isServiceDisabled(id) ? "Servicio inhabilitado." : "Servicio reactivado.");
    });
  }
  if (el.btnServiceClose) el.btnServiceClose.addEventListener("click", closeServiceModal);
  if (el.btnServiceDiscard) el.btnServiceDiscard.addEventListener("click", closeServiceModal);
  if (el.serviceBackdrop) {
    bindSafeBackdropClose(el.serviceBackdrop, closeServiceModal);
  }
  if (el.btnGlobalGoalsClose) el.btnGlobalGoalsClose.addEventListener("click", closeGlobalGoalsModal);
  if (el.btnGlobalGoalReset) {
    el.btnGlobalGoalReset.addEventListener("click", () => {
      loadGlobalGoalInModal("");
    });
  }
  if (el.btnGlobalGoalSave) {
    el.btnGlobalGoalSave.addEventListener("click", () => {
      saveGlobalGoalFromModal();
    });
  }
  if (el.btnGlobalGoalDisable) {
    el.btnGlobalGoalDisable.addEventListener("click", () => {
      toggleGlobalGoalActive();
    });
  }
  if (el.globalGoalsEditSelect) {
    el.globalGoalsEditSelect.addEventListener("change", () => {
      loadGlobalGoalInModal(el.globalGoalsEditSelect.value || "");
    });
  }
  if (el.globalGoalsBody) {
    el.globalGoalsBody.addEventListener("click", (e) => {
      const editBtn = e.target.closest("[data-global-goal-edit]");
      if (editBtn) {
        loadGlobalGoalInModal(editBtn.dataset.globalGoalEdit || "");
        return;
      }
      const toggleBtn = e.target.closest("[data-global-goal-toggle]");
      if (toggleBtn) {
        toggleGlobalGoalActive(toggleBtn.dataset.globalGoalToggle || "");
      }
    });
  }
  if (el.globalGoalsBackdrop) {
    bindSafeBackdropClose(el.globalGoalsBackdrop, closeGlobalGoalsModal);
  }
  if (el.btnSalonesClose) el.btnSalonesClose.addEventListener("click", closeSalonesModal);
  if (el.btnSalonReset) {
    el.btnSalonReset.addEventListener("click", () => {
      loadSalonInModal("");
    });
  }
  if (el.btnSalonSave) {
    el.btnSalonSave.addEventListener("click", () => {
      saveSalonFromModal();
    });
  }
  if (el.btnSalonDisable) {
    el.btnSalonDisable.addEventListener("click", () => {
      toggleSalonActive();
    });
  }
  if (el.salonEditSelect) {
    el.salonEditSelect.addEventListener("change", () => {
      loadSalonInModal(el.salonEditSelect.value || "");
    });
  }
  if (el.salonesBody) {
    el.salonesBody.addEventListener("click", (e) => {
      const editBtn = e.target.closest("[data-salon-edit]");
      if (editBtn) {
        loadSalonInModal(editBtn.dataset.salonEdit || "");
        return;
      }
      const toggleBtn = e.target.closest("[data-salon-toggle]");
      if (toggleBtn) {
        toggleSalonActive(toggleBtn.dataset.salonToggle || "");
      }
    });
  }
  if (el.salonesBackdrop) {
    bindSafeBackdropClose(el.salonesBackdrop, closeSalonesModal);
  }
  if (el.menuSuggestionsBackdrop) {
    bindSafeBackdropClose(el.menuSuggestionsBackdrop, closeMenuSuggestionsModal);
  }
  bindMenuSuggestDnD(el.menuSuggestionsSalsas);
  bindMenuSuggestDnD(el.menuSuggestionsPostres);
  bindMenuSuggestDnD(el.menuSuggestionsGuarniciones);
  bindMenuSuggestDnD(el.mmsGuarnicionesSuggested);
  bindMenuSuggestDnD(el.mmsGuarnicionesAll);
  bindMenuSuggestDnD(el.mmsPostresSuggested);
  bindMenuSuggestDnD(el.mmsPostresAll);
  bindMenuSuggestDnD(el.mmsComentariosAll);
  bindMenuSuggestDnD(el.mmsMontajeAdicionales);
  if (el.btnMenuSuggestionsClose) {
    el.btnMenuSuggestionsClose.addEventListener("click", closeMenuSuggestionsModal);
  }
  if (el.btnMenuSuggestionsDiscard) {
    el.btnMenuSuggestionsDiscard.addEventListener("click", closeMenuSuggestionsModal);
  }
  if (el.menuSuggestionsProtein) {
    el.menuSuggestionsProtein.addEventListener("change", async () => {
      const proteinId = Number(el.menuSuggestionsProtein.value || 0);
      const preps = await readMenuCatalog("preparacion", `plato_id=${encodeURIComponent(String(proteinId || ""))}`);
      if (el.menuSuggestionsPreparation) {
        el.menuSuggestionsPreparation.innerHTML = "";
        for (const p of preps.filter((x) => x && x.activo !== false)) {
          const opt = document.createElement("option");
          opt.value = String(p.id);
          opt.textContent = String(p.nombre || "");
          el.menuSuggestionsPreparation.appendChild(opt);
        }
        if (!el.menuSuggestionsPreparation.options.length) {
          el.menuSuggestionsPreparation.innerHTML = `<option value="">Sin preparaciones para esta proteina</option>`;
        }
      }
      await refreshMenuSuggestionsModalData();
    });
  }
  if (el.menuSuggestionsPreparation) {
    el.menuSuggestionsPreparation.addEventListener("change", () => {
      refreshMenuSuggestionsModalData().catch(() => {
        toast("No se pudieron cargar sugerencias.");
      });
    });
  }
  if (el.btnMenuSuggestionsSave) {
    el.btnMenuSuggestionsSave.addEventListener("click", async () => {
      const platoId = Number(el.menuSuggestionsProtein?.value || 0);
      const preparacionId = Number(el.menuSuggestionsPreparation?.value || 0);
      if (!Number.isFinite(platoId) || platoId <= 0) return toast("Selecciona un plato base.");
      if (!Number.isFinite(preparacionId) || preparacionId <= 0) return toast("Selecciona una preparacion.");
      await saveMenuSuggestions({
        id_plato_fuerte: platoId,
        id_preparacion: preparacionId,
        salsaIds: selectedIdsFromChecklist(el.menuSuggestionsSalsas),
        postreIds: selectedIdsFromChecklist(el.menuSuggestionsPostres),
        guarnicionIds: selectedIdsFromChecklist(el.menuSuggestionsGuarniciones),
        bebidaIds: selectedIdsFromChecklist(el.menuSuggestionsBebidas),
        montajeTipoIds: selectedIdsFromChecklist(el.menuSuggestionsMontajeTipos),
        montajeAdicionalIds: selectedIdsFromChecklist(el.menuSuggestionsMontajeAdicionales),
      });
      const pairKey = `${platoId}:${preparacionId}`;
      ensureMmsCatalogDefaults();
      if (menuMontajeSelectableCatalogCache.suggestionsByPair instanceof Map) {
        menuMontajeSelectableCatalogCache.suggestionsByPair.delete(pairKey);
      }
      if (menuMontajeSelectableCatalogCache.suggestionsInFlightByPair instanceof Map) {
        menuMontajeSelectableCatalogCache.suggestionsInFlightByPair.delete(pairKey);
      }
      toast("Combinacion de platillo actualizada.");
    });
  }
  if (el.btnMenuSuggestionsManageCatalog) {
    el.btnMenuSuggestionsManageCatalog.addEventListener("click", async () => {
      await openMenuCatalogManagerModal("plato_fuerte");
    });
  }
  if (el.menuCatalogBackdrop) {
    bindSafeBackdropClose(el.menuCatalogBackdrop, closeMenuCatalogManagerModal);
  }
  if (el.btnMenuCatalogClose) {
    el.btnMenuCatalogClose.addEventListener("click", closeMenuCatalogManagerModal);
  }
  if (el.btnMenuCatalogDiscard) {
    el.btnMenuCatalogDiscard.addEventListener("click", closeMenuCatalogManagerModal);
  }
  if (el.btnMenuCatalogOpenSuggestions) {
    el.btnMenuCatalogOpenSuggestions.addEventListener("click", async () => {
      closeMenuCatalogManagerModal();
      await openMenuSuggestionsModal();
    });
  }
  if (el.btnMenuCatalogReset) {
    el.btnMenuCatalogReset.addEventListener("click", () => {
      resetMenuCatalogManagerForm();
    });
  }
  if (el.menuCatalogKind) {
    el.menuCatalogKind.addEventListener("change", async () => {
      resetMenuCatalogManagerForm();
      syncMenuCatalogManagerFormByKind();
      await refreshMenuCatalogManagerRows();
    });
  }
  if (el.menuCatalogProtein) {
    el.menuCatalogProtein.addEventListener("change", async () => {
      if (String(el.menuCatalogKind?.value || "") !== "preparacion") return;
      await refreshMenuCatalogManagerRows();
    });
  }
  if (el.btnMenuCatalogSave) {
    el.btnMenuCatalogSave.addEventListener("click", async () => {
      try {
        await saveMenuCatalogManagerRecord();
        toast("Catalogo actualizado.");
      } catch (_) {
        toast("No se pudo guardar en catalogo.");
      }
    });
  }
  if (el.menuCatalogBody) {
    el.menuCatalogBody.addEventListener("click", async (e) => {
      const btn = e.target.closest("[data-mmcat-action]");
      if (!btn) return;
      const action = String(btn.dataset.mmcatAction || "").trim();
      const id = String(btn.dataset.mmcatId || "").trim();
      if (!action || !id) return;
      const kind = String(el.menuCatalogKind?.value || menuCatalogManagerKind || "plato_fuerte");
      const target = menuCatalogManagerRows.find((x) => String(x.id) === id);
      if (!target) return;

      if (action === "edit") {
        menuCatalogManagerEditingId = id;
        if (el.menuCatalogName) el.menuCatalogName.value = String(target.nombre || "");
        if (kind === "plato_fuerte") {
          if (el.menuCatalogDishType) el.menuCatalogDishType.value = String(target.tipo_plato || "NORMAL");
          if (el.menuCatalogNoProtein) el.menuCatalogNoProtein.checked = target.es_sin_proteina === true || Number(target.es_sin_proteina) !== 0;
        }
        if (kind === "preparacion" && el.menuCatalogProtein) {
          const pid = String(target.id_plato_fuerte || "").trim();
          if (pid) el.menuCatalogProtein.value = pid;
        }
        return;
      }

      if (action === "toggle") {
        const nextActive = target.activo === false ? true : false;
        const ok = await modernConfirm({
          title: nextActive ? "Reactivar registro" : "Inhabilitar registro",
          message: `Esta seguro de ${nextActive ? "reactivar" : "inhabilitar"} "${target.nombre}"?`,
          confirmText: nextActive ? "Si, reactivar" : "Si, inhabilitar",
          cancelText: "No",
        });
        if (!ok) return;
        await updateMenuCatalog(kind, id, { activo: nextActive ? 1 : 0 });
        await refreshMenuCatalogManagerRows();
        toast(nextActive ? "Registro reactivado." : "Registro inhabilitado.");
      }
    });
  }
  if (el.serviceForm) {
    el.serviceForm.addEventListener("submit", (e) => {
      e.preventDefault();
      saveServiceFromForm();
    });
  }
  if (el.serviceCategory) {
    el.serviceCategory.addEventListener("change", () => {
      renderSubcategoriasServicioSelect(Number(el.serviceCategory.value));
    });
  }
  if (el.btnServiceCategoryManage) {
    el.btnServiceCategoryManage.addEventListener("click", () => {
      openServiceCategoryManagerModal();
    });
  }
  if (el.btnServiceSubcategoryManage) {
    el.btnServiceSubcategoryManage.addEventListener("click", () => {
      openServiceSubcategoryManagerModal();
    });
  }
  if (el.btnServiceCategoryClose) el.btnServiceCategoryClose.addEventListener("click", closeServiceCategoryManagerModal);
  if (el.btnServiceCategoryReset) {
    el.btnServiceCategoryReset.addEventListener("click", () => {
      loadServiceCategoryInManager("");
    });
  }
  if (el.btnServiceCategorySave) {
    el.btnServiceCategorySave.addEventListener("click", () => {
      saveServiceCategoryFromManagerModal();
    });
  }
  if (el.serviceCategoryEditSelect) {
    el.serviceCategoryEditSelect.addEventListener("change", () => {
      loadServiceCategoryInManager(el.serviceCategoryEditSelect.value || "");
    });
  }
  if (el.serviceCategoryBody) {
    el.serviceCategoryBody.addEventListener("click", (e) => {
      const editBtn = e.target.closest("[data-service-category-edit]");
      if (editBtn) {
        loadServiceCategoryInManager(editBtn.dataset.serviceCategoryEdit || "");
        return;
      }
      const toggleBtn = e.target.closest("[data-service-category-toggle]");
      if (toggleBtn) {
        toggleServiceCategoryActiveFromManager(toggleBtn.dataset.serviceCategoryToggle || "");
      }
    });
  }
  if (el.btnServiceSubcategoryClose) el.btnServiceSubcategoryClose.addEventListener("click", closeServiceSubcategoryManagerModal);
  if (el.btnServiceSubcategoryReset) {
    el.btnServiceSubcategoryReset.addEventListener("click", () => {
      const categoryId = String(el.serviceSubcategoryCategorySelect?.value || "").trim();
      loadServiceSubcategoryInManager("", categoryId);
    });
  }
  if (el.btnServiceSubcategorySave) {
    el.btnServiceSubcategorySave.addEventListener("click", () => {
      saveServiceSubcategoryFromManagerModal();
    });
  }
  if (el.serviceSubcategoryCategorySelect) {
    el.serviceSubcategoryCategorySelect.addEventListener("change", () => {
      const categoryId = String(el.serviceSubcategoryCategorySelect.value || "").trim();
      renderServiceSubcategoryManagerSelect("", categoryId);
      renderServiceSubcategoryManagerTable(categoryId);
      editingServiceSubcategoryId = "";
      if (el.serviceSubcategoryNameInput) el.serviceSubcategoryNameInput.value = "";
    });
  }
  if (el.serviceSubcategoryEditSelect) {
    el.serviceSubcategoryEditSelect.addEventListener("change", () => {
      const id = String(el.serviceSubcategoryEditSelect.value || "").trim();
      const categoryId = String(el.serviceSubcategoryCategorySelect?.value || "").trim();
      loadServiceSubcategoryInManager(id, categoryId);
    });
  }
  if (el.serviceSubcategoryBody) {
    el.serviceSubcategoryBody.addEventListener("click", (e) => {
      const editBtn = e.target.closest("[data-service-subcategory-edit]");
      if (editBtn) {
        loadServiceSubcategoryInManager(editBtn.dataset.serviceSubcategoryEdit || "");
        return;
      }
      const toggleBtn = e.target.closest("[data-service-subcategory-toggle]");
      if (toggleBtn) {
        toggleServiceSubcategoryActiveFromManager(toggleBtn.dataset.serviceSubcategoryToggle || "");
      }
    });
  }
  if (el.serviceCategoryBackdrop) {
    bindSafeBackdropClose(el.serviceCategoryBackdrop, closeServiceCategoryManagerModal);
  }
  if (el.serviceSubcategoryBackdrop) {
    bindSafeBackdropClose(el.serviceSubcategoryBackdrop, closeServiceSubcategoryManagerModal);
  }
  if (el.servicesManagerBody) {
    el.servicesManagerBody.addEventListener("click", (e) => {
      const editBtn = e.target.closest("[data-service-row-edit]");
      if (!editBtn) return;
      openServiceModal(editBtn.dataset.serviceRowEdit || "");
    });
  }

  if (el.btnAppointmentClose) el.btnAppointmentClose.addEventListener("click", closeAppointmentModal);
  if (el.btnSalesReportClose) el.btnSalesReportClose.addEventListener("click", closeSalesReportModal);
  if (el.btnAccountingReportClose) el.btnAccountingReportClose.addEventListener("click", closeAccountingReportModal);
  if (el.btnAccountStatementClose) el.btnAccountStatementClose.addEventListener("click", closeAccountingStatementModal);
  if (el.salesReportBackdrop) {
    bindSafeBackdropClose(el.salesReportBackdrop, closeSalesReportModal);
  }
  if (el.accountingReportBackdrop) {
    bindSafeBackdropClose(el.accountingReportBackdrop, closeAccountingReportModal);
  }
  if (el.accountStatementBackdrop) {
    bindSafeBackdropClose(el.accountStatementBackdrop, closeAccountingStatementModal);
  }
  [
    el.salesReportFrom,
    el.salesReportTo,
    el.salesReportUser,
    el.salesReportStatus,
    el.salesReportSalon,
    el.salesReportCompany,
  ].forEach((node) => {
    if (!node) return;
    const evt = node.tagName === "INPUT" ? "input" : "change";
    node.addEventListener(evt, () => renderSalesReportTable());
  });
  if (el.salesReportSearch) {
    el.salesReportSearch.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      renderSalesReportTable();
    });
    el.salesReportSearch.addEventListener("change", () => {
      renderSalesReportTable();
    });
  }
  if (el.btnSalesReportReset) {
    el.btnSalesReportReset.addEventListener("click", () => {
      resetSalesReportFilters();
      renderSalesReportTable();
    });
  }
  if (el.btnSalesReportExportExcel) {
    el.btnSalesReportExportExcel.addEventListener("click", () => {
      exportSalesReportToExcel();
    });
  }
  [
    el.accountingReportFrom,
    el.accountingReportTo,
    el.accountingReportUser,
    el.accountingReportStatus,
    el.accountingReportSalon,
    el.accountingReportCompany,
  ].forEach((node) => {
    if (!node) return;
    const evt = node.tagName === "INPUT" ? "input" : "change";
    node.addEventListener(evt, () => renderAccountingReportTable());
  });
  if (el.accountingReportSearch) {
    el.accountingReportSearch.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      renderAccountingReportTable();
    });
    el.accountingReportSearch.addEventListener("change", () => {
      renderAccountingReportTable();
    });
  }
  if (el.btnAccountingReportReset) {
    el.btnAccountingReportReset.addEventListener("click", () => {
      resetAccountingReportFilters();
      renderAccountingReportTable();
    });
  }
  if (el.btnAccountingReportExportExcel) {
    el.btnAccountingReportExportExcel.addEventListener("click", () => {
      exportAccountingReportToExcel();
    });
  }
  if (el.btnOccupancyReportClose) el.btnOccupancyReportClose.addEventListener("click", closeOccupancyReportModal);
  if (el.occupancyReportBackdrop) {
    bindSafeBackdropClose(el.occupancyReportBackdrop, closeOccupancyReportModal);
  }
  if (el.occupancyReportWeek) {
    el.occupancyReportWeek.addEventListener("change", () => {
      updateOccupancyReportWeekUi();
    });
  }
  if (el.btnOccupancyReportPrevWeek) {
    el.btnOccupancyReportPrevWeek.addEventListener("click", () => {
      moveOccupancyReportWeek(-1);
    });
  }
  if (el.btnOccupancyReportNextWeek) {
    el.btnOccupancyReportNextWeek.addEventListener("click", () => {
      moveOccupancyReportWeek(1);
    });
  }
  if (el.btnOccupancyReportTodayWeek) {
    el.btnOccupancyReportTodayWeek.addEventListener("click", () => {
      setOccupancyCurrentWeek();
      updateOccupancyReportWeekUi();
    });
  }
  if (el.btnOccupancyReportExportExcel) {
    el.btnOccupancyReportExportExcel.addEventListener("click", () => {
      exportOccupancyReportToExcel();
    });
  }
  if (el.btnDashboardReportClose) el.btnDashboardReportClose.addEventListener("click", closeDashboardReportModal);
  if (el.dashboardReportBackdrop) {
    bindSafeBackdropClose(el.dashboardReportBackdrop, closeDashboardReportModal);
    el.dashboardReportBackdrop.addEventListener("mousemove", handleDashboardHoverTooltip);
    el.dashboardReportBackdrop.addEventListener("mouseleave", hideDashboardHoverTip);
    el.dashboardReportBackdrop.addEventListener("scroll", hideDashboardHoverTip, true);
  }
  if (el.dashboardReportPeriod) {
    el.dashboardReportPeriod.addEventListener("change", () => {
      syncDashboardPeriodControls();
      if (dashboardResolvePeriod() === "week") {
        const weekIso = String(el.dashboardReportWeek?.value || "").trim() || weekInputFromDate(new Date());
        setDashboardFiltersByWeek(weekIso);
      } else {
        const month = String(el.dashboardReportMonth?.value || "").trim() || toISODate(new Date()).slice(0, 7);
        if (/^\d{4}-\d{2}$/.test(month)) setDashboardFiltersByMonth(month);
      }
      renderDashboardReport();
    });
  }
  if (el.dashboardReportMonth) {
    el.dashboardReportMonth.addEventListener("change", () => {
      const month = String(el.dashboardReportMonth.value || "").trim();
      if (/^\d{4}-\d{2}$/.test(month)) setDashboardFiltersByMonth(month);
      renderDashboardReport();
    });
  }
  if (el.dashboardReportWeek) {
    el.dashboardReportWeek.addEventListener("change", () => {
      const weekIso = String(el.dashboardReportWeek.value || "").trim();
      setDashboardFiltersByWeek(weekIso);
      renderDashboardReport();
    });
  }
  [
    el.dashboardReportFrom,
    el.dashboardReportTo,
    el.dashboardReportRole,
    el.dashboardReportScope,
    el.dashboardReportSeller,
  ].forEach((node) => {
    if (!node) return;
    node.addEventListener("change", () => {
      renderDashboardReport();
    });
  });
  if (el.btnDashboardReportCurrentMonth) {
    el.btnDashboardReportCurrentMonth.addEventListener("click", () => {
      if (dashboardResolvePeriod() === "week") {
        const weekIso = weekInputFromDate(new Date());
        setDashboardFiltersByWeek(weekIso);
      } else {
        const month = toISODate(new Date()).slice(0, 7);
        setDashboardFiltersByMonth(month);
      }
      renderDashboardReport();
    });
  }  if (el.btnDashboardReportReset) {
    el.btnDashboardReportReset.addEventListener("click", () => {
      resetDashboardReportFilters();
      renderDashboardReport();
    });
  }
  if (el.btnInstitutionReportClose) el.btnInstitutionReportClose.addEventListener("click", closeInstitutionReportModal);
  if (el.institutionReportBackdrop) {
    bindSafeBackdropClose(el.institutionReportBackdrop, closeInstitutionReportModal);
  }
  [
    el.institutionReportCompany,
    el.institutionReportFrom,
    el.institutionReportTo,
  ].forEach((node) => {
    if (!node) return;
    node.addEventListener("change", () => {
      renderInstitutionReport();
    });
  });
  if (el.institutionReportCompanySearch) {
    el.institutionReportCompanySearch.addEventListener("input", () => {
      renderInstitutionReportCompanyOptions();
      renderInstitutionReport();
    });
    el.institutionReportCompanySearch.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      renderInstitutionReportCompanyOptions();
      renderInstitutionReport();
    });
  }
  if (el.institutionReportCompanyResults) {
    el.institutionReportCompanyResults.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-institution-company-pick]");
      if (!btn) return;
      const companyId = String(btn.dataset.institutionCompanyPick || "").trim();
      if (!companyId) return;
      syncEnhancedSelectValue(el.institutionReportCompany, companyId);
      renderInstitutionReportCompanyOptions();
      renderInstitutionReport();
    });
  }
  if (el.btnInstitutionReportCurrentYear) {
    el.btnInstitutionReportCurrentYear.addEventListener("click", () => {
      setInstitutionReportDefaultRange();
      renderInstitutionReport();
    });
  }
  if (el.btnInstitutionReportReset) {
    el.btnInstitutionReportReset.addEventListener("click", () => {
      resetInstitutionReportFilters();
      renderInstitutionReport();
    });
  }
  if (el.institutionReportSummary) {
    el.institutionReportSummary.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-target-section]");
      if (!btn) return;
      scrollInstitutionReportToSection(btn.dataset.targetSection);
    });
  }
  if (el.institutionReportNav) {
    el.institutionReportNav.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-target-section]");
      if (!btn) return;
      scrollInstitutionReportToSection(btn.dataset.targetSection);
    });
  }
  if (el.occupancyReportBody) {
    el.occupancyReportBody.addEventListener("click", (e) => {
      const btn = e.target.closest(".occupancyQuoteLinkBtn");
      if (!btn) return;
      openOccupancyQuoteByRow(btn.dataset.eventId, btn.dataset.quoteVersion);
    });
    el.occupancyReportBody.addEventListener("click", (e) => {
      const btn = e.target.closest(".occupancyMenuMontajeLinkBtn");
      if (!btn) return;
      openOccupancyMenuMontajeByRow(btn.dataset.eventId, btn.dataset.quoteVersion);
    });
    el.occupancyReportBody.addEventListener("click", (e) => {
      const btn = e.target.closest(".occupancyChecklistLinkBtn");
      if (!btn) return;
      openEventChecklistByRow(btn.dataset.eventId);
    });
  }
  if (el.occupancyDayDetail) {
    el.occupancyDayDetail.addEventListener("click", (e) => {
      const btn = e.target.closest(".occupancyQuoteLinkBtn");
      if (!btn) return;
      openOccupancyQuoteByRow(btn.dataset.eventId, btn.dataset.quoteVersion);
    });
    el.occupancyDayDetail.addEventListener("click", (e) => {
      const btn = e.target.closest(".occupancyMenuMontajeLinkBtn");
      if (!btn) return;
      openOccupancyMenuMontajeByRow(btn.dataset.eventId, btn.dataset.quoteVersion);
    });
    el.occupancyDayDetail.addEventListener("click", (e) => {
      const btn = e.target.closest(".occupancyChecklistLinkBtn");
      if (!btn) return;
      openEventChecklistByRow(btn.dataset.eventId);
    });
  }
  if (el.btnChecklistTemplateClose) el.btnChecklistTemplateClose.addEventListener("click", closeChecklistTemplateModal);
  if (el.checklistTemplateBackdrop) bindSafeBackdropClose(el.checklistTemplateBackdrop, closeChecklistTemplateModal);
  if (el.checklistTemplateSelect) {
    el.checklistTemplateSelect.addEventListener("change", () => {
      saveChecklistTemplateDraft();
      loadChecklistTemplateEditor(el.checklistTemplateSelect.value || "");
    });
  }
  if (el.checklistTemplateName) {
    el.checklistTemplateName.addEventListener("change", () => {
      saveChecklistTemplateDraft();
      renderChecklistTemplateSelect(checklistTemplateCurrentId);
    });
  }
  if (el.checklistTemplateActive) {
    el.checklistTemplateActive.addEventListener("change", () => {
      saveChecklistTemplateDraft();
      renderChecklistTemplateSelect(checklistTemplateCurrentId);
      if (el.btnChecklistTemplateDisable) {
        el.btnChecklistTemplateDisable.textContent = el.checklistTemplateActive.checked ? "Inhabilitar" : "Reactivar";
      }
    });
  }
  if (el.btnChecklistTemplateNew) {
    el.btnChecklistTemplateNew.addEventListener("click", () => {
      saveChecklistTemplateDraft();
      const next = normalizeChecklistTemplateRecord({
        name: "Nueva plantilla",
        active: true,
        sections: [{ name: "General", active: true }],
        items: [],
      });
      checklistTemplatesDraft.push(next);
      loadChecklistTemplateEditor(next.id);
      if (el.checklistTemplateName) el.checklistTemplateName.focus();
    });
  }
  if (el.btnChecklistTemplateDisable) {
    el.btnChecklistTemplateDisable.addEventListener("click", () => {
      const current = getCurrentChecklistTemplateDraft();
      if (!current) return toast("Selecciona una plantilla.");
      current.active = current.active === false ? true : false;
      saveChecklistTemplateDraft();
      loadChecklistTemplateEditor(current.id);
      toast(current.active === false ? "Plantilla inhabilitada." : "Plantilla reactivada.");
    });
  }
  if (el.btnChecklistTemplateAdd) {
    el.btnChecklistTemplateAdd.addEventListener("click", () => {
      addChecklistTemplateItemFromInput();
    });
  }
  if (el.btnChecklistTemplateAddSection) {
    el.btnChecklistTemplateAddSection.addEventListener("click", () => {
      addChecklistSectionFromInput();
    });
  }
  if (el.btnChecklistTemplateResetSection) {
    el.btnChecklistTemplateResetSection.addEventListener("click", () => {
      resetChecklistSectionEditor();
    });
  }
  if (el.checklistTemplateSectionEditSelect) {
    el.checklistTemplateSectionEditSelect.addEventListener("change", () => {
      const id = String(el.checklistTemplateSectionEditSelect.value || "").trim();
      checklistTemplateSectionEditingId = id;
      const target = (checklistTemplateSectionsDraft || []).find((s) => String(s?.id || "") === id);
      if (el.checklistTemplateSectionInput) el.checklistTemplateSectionInput.value = String(target?.name || "");
    });
  }
  if (el.checklistTemplateSectionInput) {
    el.checklistTemplateSectionInput.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      addChecklistSectionFromInput();
    });
  }
  if (el.checklistTemplateInput) {
    el.checklistTemplateInput.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      e.preventDefault();
      addChecklistTemplateItemFromInput();
    });
  }
  if (el.checklistTemplateBody) {
    el.checklistTemplateBody.addEventListener("click", (e) => {
      const editBtn = e.target.closest("[data-checklist-template-edit]");
      if (editBtn) {
        const id = String(editBtn.dataset.checklistTemplateEdit || "").trim();
        const row = (checklistTemplateDraft || []).find((x) => String(x?.id || "") === id);
        if (!row) return;
        checklistTemplateEditingId = id;
        if (el.checklistTemplateInput) {
          el.checklistTemplateInput.value = String(row.label || "");
          el.checklistTemplateInput.focus();
        }
        renderChecklistSectionSelect(String(row.sectionId || ""));
        if (el.btnChecklistTemplateAdd) el.btnChecklistTemplateAdd.textContent = "Guardar cambio";
        return;
      }
      const upBtn = e.target.closest("[data-checklist-template-up]");
      if (upBtn) {
        const id = String(upBtn.dataset.checklistTemplateUp || "").trim();
        const idx = (checklistTemplateDraft || []).findIndex((x) => String(x?.id || "") === id);
        if (idx > 0) {
          const tmp = checklistTemplateDraft[idx - 1];
          checklistTemplateDraft[idx - 1] = checklistTemplateDraft[idx];
          checklistTemplateDraft[idx] = tmp;
          saveChecklistTemplateDraft();
          renderChecklistTemplateTable();
        }
        return;
      }
      const downBtn = e.target.closest("[data-checklist-template-down]");
      if (downBtn) {
        const id = String(downBtn.dataset.checklistTemplateDown || "").trim();
        const idx = (checklistTemplateDraft || []).findIndex((x) => String(x?.id || "") === id);
        if (idx >= 0 && idx < checklistTemplateDraft.length - 1) {
          const tmp = checklistTemplateDraft[idx + 1];
          checklistTemplateDraft[idx + 1] = checklistTemplateDraft[idx];
          checklistTemplateDraft[idx] = tmp;
          saveChecklistTemplateDraft();
          renderChecklistTemplateTable();
        }
        return;
      }
      const btn = e.target.closest("[data-checklist-template-remove]");
      if (!btn) return;
      const id = String(btn.dataset.checklistTemplateRemove || "").trim();
      checklistTemplateDraft = (checklistTemplateDraft || []).filter((x) => String(x?.id || "") !== id);
      saveChecklistTemplateDraft();
      renderChecklistTemplateTable();
      if (checklistTemplateEditingId === id) resetChecklistTemplateEditor();
      toast("Punto eliminado.");
    });
  }
  if (el.checklistTemplateSectionsBody) {
    el.checklistTemplateSectionsBody.addEventListener("click", (e) => {
      const editBtn = e.target.closest("[data-checklist-section-edit]");
      if (editBtn) {
        const id = String(editBtn.dataset.checklistSectionEdit || "").trim();
        checklistTemplateSectionEditingId = id;
        const row = (checklistTemplateSectionsDraft || []).find((s) => String(s?.id || "") === id);
        if (el.checklistTemplateSectionInput) {
          el.checklistTemplateSectionInput.value = String(row?.name || "");
          el.checklistTemplateSectionInput.focus();
        }
        renderChecklistSectionEditSelect(id);
        return;
      }
      const toggleBtn = e.target.closest("[data-checklist-section-toggle]");
      if (!toggleBtn) return;
      const id = String(toggleBtn.dataset.checklistSectionToggle || "").trim();
      const idx = (checklistTemplateSectionsDraft || []).findIndex((s) => String(s?.id || "") === id);
      if (idx < 0) return;
      checklistTemplateSectionsDraft[idx] = {
        ...checklistTemplateSectionsDraft[idx],
        active: checklistTemplateSectionsDraft[idx].active === false ? true : false,
      };
      saveChecklistTemplateDraft();
      renderChecklistSectionSelect("");
      renderChecklistSectionEditSelect("");
      renderChecklistSectionsTable();
      renderChecklistTemplateTable();
      toast(checklistTemplateSectionsDraft[idx].active === false ? "Seccion inhabilitada." : "Seccion reactivada.");
    });
  }
  if (el.btnEventChecklistClose) el.btnEventChecklistClose.addEventListener("click", closeEventChecklistModal);
  if (el.btnEventChecklistDiscard) el.btnEventChecklistDiscard.addEventListener("click", closeEventChecklistModal);
  if (el.eventChecklistBackdrop) bindSafeBackdropClose(el.eventChecklistBackdrop, closeEventChecklistModal);
  if (el.eventChecklistTemplateSelect) {
    el.eventChecklistTemplateSelect.addEventListener("change", () => {
      if (!currentEventChecklistId) return;
      const notes = String(el.eventChecklistNotes?.value || "").trim();
      const nextDraft = buildEventChecklistDraft(currentEventChecklistId, el.eventChecklistTemplateSelect.value || "");
      if (!nextDraft) return;
      nextDraft.notes = notes;
      eventChecklistDraft = nextDraft;
      renderEventChecklistRows();
    });
  }
  if (el.btnEventChecklistSave) {
    el.btnEventChecklistSave.addEventListener("click", () => {
      saveEventChecklistFromModal();
    });
  }
  if (el.eventChecklistBody) {
    el.eventChecklistBody.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-checklist-cycle-index]");
      if (!btn || !eventChecklistDraft) return;
      const idx = Number(btn.dataset.checklistCycleIndex || -1);
      if (!Number.isFinite(idx) || idx < 0) return;
      const item = eventChecklistDraft.items?.[idx];
      if (!item) return;
      item.status = cycleChecklistStatus(item.status);
      renderEventChecklistRows();
    });
    el.eventChecklistBody.addEventListener("input", (e) => {
      const input = e.target.closest("[data-checklist-comment-index]");
      if (!input || !eventChecklistDraft) return;
      const idx = Number(input.dataset.checklistCommentIndex || -1);
      if (!Number.isFinite(idx) || idx < 0) return;
      const item = eventChecklistDraft.items?.[idx];
      if (!item) return;
      item.comment = String(input.value || "").trim();
    });
  }
  if (el.appointmentBackdrop) {
    bindSafeBackdropClose(el.appointmentBackdrop, closeAppointmentModal);
  }
  if (el.appointmentForm) {
    el.appointmentForm.addEventListener("submit", (e) => {
      e.preventDefault();
      saveAppointmentFromForm();
    });
  }
  if (el.appointmentBody) {
    el.appointmentBody.addEventListener("click", async (e) => {
      const btn = e.target.closest("button[data-reminder-action]");
      if (!btn) return;
      const action = String(btn.dataset.reminderAction || "");
      const eventId = String(btn.dataset.eventId || "");
      const reminderId = String(btn.dataset.reminderId || "");
      if (!eventId || !reminderId) return;
      if (action === "edit") {
        await openReminderEditor(eventId, reminderId);
        return;
      }
      if (action === "done") {
        markReminderDone(eventId, reminderId);
        return;
      }
      if (action === "delete") {
        const ok = await modernConfirm({
          title: "Eliminar cita",
          message: "Esta seguro de eliminar esta cita?",
          confirmText: "Si, eliminar",
        });
        if (!ok) return;
        removeReminder(eventId, reminderId);
      }
    });
  }

  el.companyForm.addEventListener("submit", (e) => {
    e.preventDefault();
    const name = el.companyName.value.trim();
    const owner = el.companyOwner.value.trim();
    const email = el.companyEmail.value.trim();
    const nit = el.companyNIT.value.trim();
    const businessName = el.companyBusinessName.value.trim();
    const eventType = el.companyEventType.value;
    const address = el.companyAddress.value.trim();
    const phone = el.companyPhone.value.trim();
    const notes = el.companyNotes.value.trim();
    if (!name || !owner || !email || !nit || !businessName || !eventType || !address || !phone) {
      return toast("Completa todos los campos obligatorios de empresa.");
    }
    if (!isValidEmail(email)) {
      return toast("Correo de empresa invalido.");
    }
    if (!companyManagersDraft.length) {
      return toast("Agrega al menos un encargado para la empresa.");
    }
    const invalidManager = companyManagersDraft.find((m) => !isValidEmail(m?.email));
    if (invalidManager) {
      return toast(`Correo invalido en encargado: ${String(invalidManager.name || "sin nombre")}.`);
    }
    const editingId = String(editingCompanyId || "").trim();
    const company = normalizeCompanyRecord({
      id: editingId || uid(),
      name,
      owner,
      email,
      nit,
      businessName,
      billTo: businessName,
      eventType,
      address,
      phone,
      notes,
      managers: deepClone(companyManagersDraft),
    });
    if (editingId) {
      const idx = (state.companies || []).findIndex((c) => String(c.id || "") === editingId);
      if (idx >= 0) {
        const statusChanged = !!el.companyActive && el.companyActive.checked !== !isCompanyDisabled(editingId);
        if (areCompaniesEquivalent(state.companies[idx], company) && !statusChanged) {
          return toast("Sin cambios detectados en empresa.");
        }
        state.companies[idx] = company;
      } else {
        state.companies.push(company);
      }
      enableCompany(company.id);
      companyManagersDraft.forEach((m) => enableManager(m.id));
    } else {
      state.companies.push(company);
      companyManagersDraft.forEach((m) => enableManager(m.id));
    }
    if (el.companyActive?.checked === false) {
      state.disabledCompanies = Array.from(new Set([...(state.disabledCompanies || []), company.id]));
    } else {
      enableCompany(company.id);
    }
    persist();
    const quoteCtx = getCurrentQuoteHistoryContext();
    if (quoteCtx) {
      appendHistoryByKey(
        quoteCtx.key,
        authSession.userId || quoteCtx.ev.userId || "",
        editingId ? `Empresa actualizada en cotizacion: ${name}` : `Empresa creada y usada en cotizacion: ${name}`
      );
    }
    renderCompaniesSelect(company.id);
    closeCompanyModal();
    toast(editingId ? "Empresa actualizada." : "Empresa agregada.");
  });

  el.btnAddServiceToQuote.addEventListener("click", () => {
    addServiceToQuoteDraft(el.quoteServiceSearch.value);
  });
  el.quoteServiceSearch.addEventListener("keydown", (e) => {
    if (e.key === "Enter") {
      e.preventDefault();
      addServiceToQuoteDraft(el.quoteServiceSearch.value);
    }
  });
  if (el.quoteServiceSearch) {
    el.quoteServiceSearch.addEventListener("input", () => {
      syncQuoteServiceQuantityField();
    });
    el.quoteServiceSearch.addEventListener("change", () => {
      syncQuoteServiceQuantityField();
    });
  }
  if (el.quoteServiceQty) {
    el.quoteServiceQty.addEventListener("input", () => {
      el.quoteServiceQty.classList.remove("fieldInvalid");
    });
  }
  if (el.btnToggleQuoteItemsExpand) {
    el.btnToggleQuoteItemsExpand.addEventListener("click", () => {
      setQuoteItemsExpanded(!quoteItemsExpanded);
    });
  }
  if (el.quoteSelectAllItems) {
    el.quoteSelectAllItems.addEventListener("change", () => {
      if (!quoteDraft) return;
      const checked = !!el.quoteSelectAllItems.checked;
      if (checked) {
        quoteSelectedItemIds = new Set((quoteDraft.items || []).map((item) => String(item.rowId || "")));
      } else {
        quoteSelectedItemIds = new Set();
      }
      renderQuoteItems();
    });
  }
  if (el.btnQuoteDuplicateSelected) {
    el.btnQuoteDuplicateSelected.addEventListener("click", () => {
      duplicateSelectedQuoteItems();
    });
  }
  if (el.btnQuoteMoveSelectedUp) {
    el.btnQuoteMoveSelectedUp.addEventListener("click", () => {
      moveSelectedQuoteItems("up");
    });
  }
  if (el.btnQuoteMoveSelectedDown) {
    el.btnQuoteMoveSelectedDown.addEventListener("click", () => {
      moveSelectedQuoteItems("down");
    });
  }
    el.quoteCompany.addEventListener("change", () => {
    selectCompanyInQuote(el.quoteCompany.value);
  });
  if (el.quoteCompanySearch) {
    el.quoteCompanySearch.addEventListener("focus", () => {
      refreshCompanySuggestions(el.quoteCompanySearch.value);
    });
    el.quoteCompanySearch.addEventListener("input", () => {
      refreshCompanySuggestions(el.quoteCompanySearch.value);
    });
    el.quoteCompanySearch.addEventListener("change", () => {
      const matched = resolveCompanyFromSearch(el.quoteCompanySearch.value);
      if (matched) {
        selectCompanyInQuote(matched.id);
        return;
      }
      if (!String(el.quoteCompanySearch.value || "").trim()) {
        if (el.quoteCompany) syncEnhancedSelectValue(el.quoteCompany, "");
        hideQuoteCompanyResults();
      }
    });
    el.quoteCompanySearch.addEventListener("blur", () => {
      setTimeout(() => hideQuoteCompanyResults(), 140);
    });
    el.quoteCompanySearch.addEventListener("keydown", (e) => {
      if (e.key === "Escape") {
        hideQuoteCompanyResults();
        return;
      }
      if (e.key !== "Enter") return;
      const matched = resolveCompanyFromSearch(el.quoteCompanySearch.value);
      if (!matched) return;
      e.preventDefault();
      selectCompanyInQuote(matched.id);
    });
  }
  if (el.quoteCompanyResults) {
    el.quoteCompanyResults.addEventListener("click", (e) => {
      const btn = e.target.closest("[data-quote-company-pick]");
      if (!btn) return;
      const companyId = String(btn.dataset.quoteCompanyPick || "").trim();
      if (!companyId) return;
      selectCompanyInQuote(companyId);
    });
  }
  el.quoteManagerSelect.addEventListener("change", () => {
    if (quoteDraft) quoteDraft.managerId = el.quoteManagerSelect.value;
    applyQuoteCompanyDefaults();
    fillQuoteHeaderFields(true, false);
  });
  if (el.quoteTemplateSelect) {
    el.quoteTemplateSelect.addEventListener("change", () => {
      if (!quoteDraft) return;
      quoteDraft.templateId = String(el.quoteTemplateSelect.value || "").trim();
    });
  }
  if (el.btnQuoteCurrencyGTQ) {
    el.btnQuoteCurrencyGTQ.addEventListener("click", () => {
      if (!quoteDraft) return;
      quoteDraft.currency = "GTQ";
      renderQuoteItems();
    });
  }
  if (el.btnQuoteCurrencyUSD) {
    el.btnQuoteCurrencyUSD.addEventListener("click", () => {
      if (!quoteDraft) return;
      quoteDraft.currency = "USD";
      renderQuoteItems();
    });
  }
  if (el.quoteServiceTemplateSelect) {
    el.quoteServiceTemplateSelect.addEventListener("change", () => {
      const selectedId = String(el.quoteServiceTemplateSelect.value || "").trim();
      const templates = normalizeQuoteServiceTemplatesCollection(getQuoteServiceTemplatesFromState());
      const found = templates.find((t) => String(t.id) === selectedId);
      if (el.quoteServiceTemplateName) {
        el.quoteServiceTemplateName.value = found ? String(found.name || "") : "";
      }
    });
  }
  if (el.btnQuoteServiceTemplateApply) {
    el.btnQuoteServiceTemplateApply.addEventListener("click", async () => {
      await applyQuoteServiceTemplateToDraft(el.quoteServiceTemplateSelect?.value || "");
    });
  }
  if (el.btnQuoteServiceTemplateSave) {
    el.btnQuoteServiceTemplateSave.addEventListener("click", () => {
      saveQuoteServiceTemplateFromDraft(false);
    });
  }
  if (el.btnQuoteServiceTemplateUpdate) {
    el.btnQuoteServiceTemplateUpdate.addEventListener("click", () => {
      saveQuoteServiceTemplateFromDraft(true);
    });
  }
  if (el.btnQuoteServiceTemplateDelete) {
    el.btnQuoteServiceTemplateDelete.addEventListener("click", async () => {
      await deleteQuoteServiceTemplateSelected();
    });
  }
  if (el.btnLoadQuoteVersion) {
    el.btnLoadQuoteVersion.addEventListener("click", () => {
      if (!quoteDraft || !el.quoteVersionSelect) return;
      const selected = Number(el.quoteVersionSelect.value || 0);
      if (!Number.isFinite(selected) || selected <= 0) return;
      const currentVersion = Number(quoteDraft.version || 0);
      if (selected === currentVersion) return;
      const versions = normalizeQuoteVersionHistory(quoteDraft.versions);
      const snapshot = versions.find((v) => Number(v.version) === selected);
      if (!snapshot) return toast("Version no disponible.");
      applyQuoteSnapshotToDraft(snapshot);
      toast(`Version V${selected} cargada. Guarda para crear una nueva version.`);
    });
  }
  if (el.quoteDocFold && el.quoteBackdrop) {
    el.quoteDocFold.addEventListener("toggle", () => {
      if (el.quoteDocPanelCard && !el.quoteDocFold.open && !el.quoteDocPanelCard.hidden) {
        setQuoteDocPanelVisible(false);
        return;
      }
      syncQuoteDocFoldUi();
    });
  }
  if (el.btnToggleQuoteDoc) {
    el.btnToggleQuoteDoc.addEventListener("click", () => {
      const nextVisible = el.quoteDocPanelCard ? el.quoteDocPanelCard.hidden : !el.quoteDocFold?.open;
      setQuoteDocPanelVisible(nextVisible);
      if (nextVisible && el.quoteDocFold) {
        el.quoteDocFold.scrollIntoView({ behavior: "smooth", block: "start" });
      }
    });
  }
  if (el.btnQuoteDocSaveClose) {
    el.btnQuoteDocSaveClose.addEventListener("click", handleQuoteDocSaveClose);
  }
  if (el.quoteDocPanelCard) {
    el.quoteDocPanelCard.addEventListener("click", (e) => {
      const btn = e.target.closest("#btnQuoteDocSaveClose");
      if (!btn) return;
      e.preventDefault();
      handleQuoteDocSaveClose();
    });
  }
  if (el.quoteDocDate) {
    el.quoteDocDate.addEventListener("change", () => {
      syncQuoteDateDefaults();
      if (quoteDraft) quoteDraft.docDate = String(el.quoteDocDate.value || "").trim();
    });
  }
  if (el.quoteDueDate) {
    el.quoteDueDate.addEventListener("input", () => {
      el.quoteDueDate.dataset.autoDefault = "0";
      if (quoteDraft) quoteDraft.dueDate = String(el.quoteDueDate.value || "").trim();
    });
  }
  el.quotePeople.addEventListener("input", () => {
    if (!quoteDraft) return;
    quoteDraft.people = el.quotePeople.value;
    syncPaxQuantityItems();
  });
  if (el.quoteDiscountType) {
    el.quoteDiscountType.addEventListener("change", () => {
      if (!quoteDraft) return;
      quoteDraft.discountType = normalizeDiscountType(el.quoteDiscountType.value);
      renderQuoteItems();
    });
  }
  if (el.quoteDiscountValue) {
    el.quoteDiscountValue.addEventListener("input", () => {
      if (!quoteDraft) return;
      quoteDraft.discountValue = Math.max(0, Number(el.quoteDiscountValue.value || 0));
      renderQuoteItems();
    });
  }

  el.btnAddManager.addEventListener("click", addManagerToCompanyDraft);
  el.managersBody.addEventListener("click", async (e) => {
    const editBtn = e.target.closest(".editManagerBtn");
    if (editBtn) {
      const id = String(editBtn.dataset.managerId || "").trim();
      const target = companyManagersDraft.find((x) => String(x.id || "") === id);
      if (!target) return;
      editingCompanyManagerId = id;
      el.managerName.value = String(target.name || "");
      el.managerPhone.value = String(target.phone || "");
      el.managerEmail.value = String(target.email || "");
      el.managerAddress.value = String(target.address || "");
      if (el.btnAddManager) el.btnAddManager.textContent = "Actualizar encargado";
      try { el.managerName.focus(); } catch (_) { }
      return;
    }
    const btn = e.target.closest(".removeManagerBtn");
    if (!btn) return;
    const id = btn.dataset.managerId;
    const target = companyManagersDraft.find((x) => String(x.id || "") === String(id || ""));
    const ok = await modernConfirm({
      title: "Eliminar encargado",
      message: `Esta seguro de eliminar al encargado${target?.name ? ` "${target.name}"` : ""}?`,
      confirmText: "Si, eliminar",
      cancelText: "No",
    });
    if (!ok) return;
    companyManagersDraft = companyManagersDraft.filter(x => x.id !== id);
    if (String(editingCompanyManagerId || "") === String(id || "")) {
      editingCompanyManagerId = "";
      el.managerName.value = "";
      el.managerPhone.value = "";
      el.managerEmail.value = "";
      el.managerAddress.value = "";
      if (el.btnAddManager) el.btnAddManager.textContent = "+ Encargado";
    }
    renderCompanyManagersDraft();
    toast("Encargado eliminado del borrador.");
  });

  el.quoteForm.noValidate = true;
  el.quoteForm.addEventListener("submit", async (e) => {
    e.preventDefault();
    await saveQuoteFromForm();
  });

  el.quoteItemsBody.addEventListener("change", handleQuoteItemsInput);
  el.quoteItemsBody.addEventListener("click", handleQuoteItemsClick);

  el.eventUser.addEventListener("change", () => {
    if (!el.userBackdrop.hidden) closeUserModal();
  });

  window.addEventListener("mousemove", onGlobalPointerMove);
  window.addEventListener("mouseup", onGlobalPointerUp);

  // ESC closes
  window.addEventListener("keydown", (e) => {
    if (e.key === "Escape") closeOpenCustomTopbarSelect();
    if (e.key === "Escape") closeSettingsPanel();
    if (e.key === "Escape") closeTopbarReminderPanel();
    if (e.key === "Escape") {
      if (!el.userBackdrop.hidden) closeUserModal();
      else if (el.accountStatementBackdrop && !el.accountStatementBackdrop.hidden) closeAccountingStatementModal();
      else if (el.salesReportBackdrop && !el.salesReportBackdrop.hidden) closeSalesReportModal();
      else if (el.accountingReportBackdrop && !el.accountingReportBackdrop.hidden) closeAccountingReportModal();
      else if (el.occupancyReportBackdrop && !el.occupancyReportBackdrop.hidden) closeOccupancyReportModal();
      else if (el.dashboardReportBackdrop && !el.dashboardReportBackdrop.hidden) closeDashboardReportModal();
      else if (el.institutionReportBackdrop && !el.institutionReportBackdrop.hidden) closeInstitutionReportModal();
      else if (!el.appointmentBackdrop.hidden) closeAppointmentModal();
      else if (el.menuMontajeBackdrop && !el.menuMontajeBackdrop.hidden) closeMenuMontajeModal();
      else if (el.menuMontajeSelectableBackdrop && !el.menuMontajeSelectableBackdrop.hidden) closeMenuMontajeSelectableModal();
      else if (el.menuSuggestionsBackdrop && !el.menuSuggestionsBackdrop.hidden) closeMenuSuggestionsModal();
      else if (el.menuCatalogBackdrop && !el.menuCatalogBackdrop.hidden) closeMenuCatalogManagerModal();
      else if (!el.serviceBackdrop.hidden) closeServiceModal();
      else if (!el.companyBackdrop.hidden) closeCompanyModal();
      else if (!el.quoteBackdrop.hidden) closeQuoteModal();
      else if (!el.modalBackdrop.hidden) closeModal();
    }
  });
  document.addEventListener("click", (e) => {
    closeOpenCustomTopbarSelect();
    if (!el.settingsMenu) return;
    if (el.settingsMenu.contains(e.target)) return;
    closeSettingsPanel();
  });
  document.addEventListener("click", (e) => {
    if (!el.topbarReminderWrap) return;
    if (el.topbarReminderWrap.contains(e.target)) return;
    closeTopbarReminderPanel();
  });
}

function openModalForCreate({ date, start, end, salon, rangeDates = null }) {
  const d = stripTime(date);
  const startTime = isValidClockTime(start) ? start : "09:00";
  const startMinuteFromGrid = Math.max(0, timeToMinutes(startTime) - HOUR_START * 60);
  const fallbackEnd = minutesToTime(startMinuteFromGrid + 60);
  const endTime = (isValidClockTime(end) && compareTime(end, startTime) > 0) ? end : fallbackEnd;
  const selectedSalonValue = String(salon || "").trim();
  const initialSalon = (selectedSalonValue && selectedSalonValue !== ALL_ROOMS_VALUE) ? selectedSalonValue : "";
  el.modalTitle.textContent = "Reservar salon";
  const totalDates = Array.isArray(rangeDates) && rangeDates.length ? rangeDates.length : 1;
  const rangeLabel = totalDates > 1
    ? `${toISODate(d)} a ${rangeDates[rangeDates.length - 1]}`
    : toISODate(d);
  const salonLabel = initialSalon || "Sin salon";
  el.modalSubtitle.textContent = `${salonLabel} - ${rangeLabel} - ${startTime}-${endTime}`;
  pendingCreateDates = totalDates > 1 ? rangeDates.slice() : null;

  el.eventId.value = "";
  el.eventName.value = "";
  setModernDateValue(el.eventDate, toISODate(d));
  setModernDateValue(el.eventDateEnd, totalDates > 1 ? rangeDates[rangeDates.length - 1] : toISODate(d));
  el.slotsBody.innerHTML = "";
  addSlotRow({ salon: initialSalon, slotPax: 0, startTime, endTime });
  syncHiddenTimesFromFirstSlot();
  el.eventStatus.value = STATUS.RESERVA_SIN_COTIZACION; // default: apartado sin cotizacion
  const sessionUserId = String(authSession.userId || "").trim();
  const sessionAvailable = (state.users || []).some((u) => String(u.id) === sessionUserId && u.active !== false);
  el.eventUser.value = sessionAvailable ? sessionUserId : (state.users[0]?.id || "");
  syncEventPaxFromSlots();
  el.eventNotes.value = "";

  el.btnDelete.hidden = true;
  el.btnCancelEvent.hidden = true;
  el.btnQuoteEvent.hidden = true;
  el.btnQuoteEvent.textContent = "Cotizar evento";
  if (el.btnSetMaintenance) el.btnSetMaintenance.textContent = "Poner en mantenimiento";
  historyTargetEventId = null;
  if (el.btnToggleHistory) el.btnToggleHistory.hidden = true;
  if (el.btnToggleAppointments) el.btnToggleAppointments.hidden = true;
  if (el.btnAddAppointment) el.btnAddAppointment.hidden = true;
  renderHistoryForEvent(null);
  setHistoryPanelVisible(false);
  renderAppointmentsForEvent(null);
  setAppointmentsPanelVisible(false);

  showModal();
  updateRulesAndConflictsUI();
  validateReservationRequiredFields();
}

async function openModalForEdit(id) {
  const ev = state.events.find(x => x.id === id);
  if (!ev) return;
  if (isEventSeriesInPast(ev)) {
    const authorized = await requestPastEventEditAuthorization(ev);
    if (!authorized) {
      await modernGuideToast("Evento de fecha pasada bloqueado. Requiere codigo de administrador.");
      return;
    }
  }
  pendingCreateDates = null;

  el.modalTitle.textContent = "Editar reserva";
  const series = getEventSeries(ev).sort((a, b) => a.date.localeCompare(b.date));
  const firstDate = getEventGeneralDateStart(series[0] || ev) || series[0]?.date || ev.date;
  const lastDate = getEventGeneralDateEnd(series[0] || ev) || series[series.length - 1]?.date || ev.date;
  el.modalSubtitle.textContent = series.length > 1
    ? `${ev.salon} - ${firstDate} a ${lastDate} - ${ev.startTime}-${ev.endTime}`
    : `${ev.salon} - ${ev.date} - ${ev.startTime}-${ev.endTime}`;

  el.eventId.value = ev.id;
  el.eventName.value = ev.name;
  setModernDateValue(el.eventDate, firstDate);
  setModernDateValue(el.eventDateEnd, lastDate);
  const slots = uniqueSlotsFromSeries(series);
  el.slotsBody.innerHTML = "";
  if (slots.length) {
    for (const slot of slots) addSlotRow(slot);
  } else {
    addSlotRow({ salon: ev.salon, slotPax: ev.slotPax ?? ev.pax ?? 0, startTime: ev.startTime, endTime: ev.endTime });
  }
  syncHiddenTimesFromFirstSlot();
  el.eventStatus.value = ev.status;
  el.eventUser.value = ev.userId;
  syncEventPaxFromSlots();
  el.eventNotes.value = ev.notes || "";

  el.btnDelete.hidden = true;
  el.btnCancelEvent.hidden = (ev.status === STATUS.CANCELADO);
  el.btnQuoteEvent.hidden = false;
  el.btnQuoteEvent.textContent = ev.quote ? "Editar cotizacion" : "Cotizar evento";
  if (el.btnSetMaintenance) {
    el.btnSetMaintenance.textContent = ev.status === STATUS.MANTENIMIENTO
      ? "Liberar mantenimiento"
      : "Poner en mantenimiento";
  }
  historyTargetEventId = id;
  if (el.btnToggleHistory) el.btnToggleHistory.hidden = false;
  if (el.btnToggleAppointments) el.btnToggleAppointments.hidden = false;
  if (el.btnAddAppointment) el.btnAddAppointment.hidden = false;
  renderHistoryForEvent(ev);
  setHistoryPanelVisible(false);
  renderAppointmentsForEvent(ev);
  setAppointmentsPanelVisible(false);

  showModal();
  updateRulesAndConflictsUI();
  validateReservationRequiredFields();
}

function showModal() {
  el.modalBackdrop.hidden = false;
  document.body.classList.add("eventModalOpen");
  // focus
  setTimeout(() => el.eventName.focus(), 0);
}

function setQuoteItemsExpanded(expanded) {
  quoteItemsExpanded = !!expanded;
  if (el.quoteItemsPanel) {
    el.quoteItemsPanel.classList.toggle("isExpanded", quoteItemsExpanded);
  }
  if (el.quoteForm) {
    el.quoteForm.classList.toggle("quoteItemsExpanded", quoteItemsExpanded);
  }
  if (el.btnToggleQuoteItemsExpand) {
    el.btnToggleQuoteItemsExpand.textContent = quoteItemsExpanded ? "Reducir" : "Expandir";
    el.btnToggleQuoteItemsExpand.setAttribute("aria-expanded", quoteItemsExpanded ? "true" : "false");
  }
}

function closeModal() {
  el.modalBackdrop.hidden = true;
  document.body.classList.remove("eventModalOpen");
  el.conflictsBox.hidden = true;
  el.statusHint.textContent = "";
  historyTargetEventId = null;
  if (el.btnToggleHistory) el.btnToggleHistory.hidden = true;
  if (el.btnToggleAppointments) el.btnToggleAppointments.hidden = true;
  if (el.btnAddAppointment) el.btnAddAppointment.hidden = true;
  if (el.historyPanel) el.historyPanel.hidden = true;
  if (el.historyBody) el.historyBody.innerHTML = "";
  if (el.appointmentPanel) el.appointmentPanel.hidden = true;
  if (el.appointmentBody) el.appointmentBody.innerHTML = "";
  setHistoryPanelVisible(false);
  setAppointmentsPanelVisible(false);
  pendingCreateDates = null;
}

function renderUserEditSelect(selectedId = "") {
  if (!el.userEditSelect) return;
  el.userEditSelect.innerHTML = "";
  const baseOpt = document.createElement("option");
  baseOpt.value = "";
  baseOpt.textContent = "Crear nuevo usuario";
  el.userEditSelect.appendChild(baseOpt);
  const ordered = (state.users || [])
    .map(normalizeUserRecord)
    .slice()
    .sort((a, b) => String(a.fullName || "").localeCompare(String(b.fullName || ""), "es", { sensitivity: "base" }));
  for (const u of ordered) {
    const opt = document.createElement("option");
    opt.value = u.id;
    opt.textContent = `${u.fullName || u.name}${u.active === false ? " (Inhabilitado)" : ""}`;
    el.userEditSelect.appendChild(opt);
  }
  el.userEditSelect.value = selectedId || "";
}

function clearUserGoalEditorFields() {
  editingUserGoalMonth = "";
  if (el.userGoalMonth) el.userGoalMonth.value = "";
  if (el.userGoalAmount) el.userGoalAmount.value = "";
  if (el.btnUserGoalAdd) el.btnUserGoalAdd.textContent = "Agregar meta";
}

function renderUserMonthlyGoalsDraft() {
  if (!el.userGoalsBody) return;
  el.userGoalsBody.innerHTML = "";
  const list = (userMonthlyGoalsDraft || [])
    .slice()
    .sort((a, b) => String(a.month || "").localeCompare(String(b.month || "")));
  for (const g of list) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(g.month)}</td>
      <td>Q ${Number(g.amount || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}</td>
      <td>
        <button type="button" class="apptIconBtn apptEdit" data-goal-action="edit" data-goal-month="${escapeHtml(g.month)}" title="Editar meta mensual" aria-label="Editar meta mensual">&#9998;</button>
        <button type="button" class="apptIconBtn apptDelete" data-goal-action="remove" data-goal-month="${escapeHtml(g.month)}" title="Eliminar meta mensual" aria-label="Eliminar meta mensual">&#128465;</button>
      </td>
    `;
    el.userGoalsBody.appendChild(tr);
  }
}

function upsertUserMonthlyGoalDraft() {
  const month = String(el.userGoalMonth?.value || "").trim();
  const amount = Math.max(0, Number(el.userGoalAmount?.value || 0));
  if (!/^\d{4}-\d{2}$/.test(month)) return toast("Selecciona mes valido para meta.");
  if (!Number.isFinite(amount) || amount <= 0) return toast("Monto de meta invalido.");
  const idx = userMonthlyGoalsDraft.findIndex((g) => String(g.month) === month);
  if (idx >= 0) {
    userMonthlyGoalsDraft[idx] = { month, amount };
  } else {
    userMonthlyGoalsDraft.push({ month, amount });
  }
  userMonthlyGoalsDraft.sort((a, b) => a.month.localeCompare(b.month));
  renderUserMonthlyGoalsDraft();
  clearUserGoalEditorFields();
}

function setUserModalMode(mode = "create", targetUser = null) {
  const isEdit = mode === "edit" && !!targetUser;
  if (el.userTitle) el.userTitle.textContent = isEdit ? "Editar usuario" : "Nuevo usuario";
  if (el.btnUserSubmit) el.btnUserSubmit.textContent = isEdit ? "Guardar cambios" : "Crear usuario";
  if (el.btnUserDisable) el.btnUserDisable.hidden = !isEdit;
  if (el.userPassword) {
    el.userPassword.required = !isEdit;
    el.userPassword.placeholder = isEdit ? "Dejar vacio para no cambiar" : "Contrasena";
  }
  if (el.userActive) {
    el.userActive.disabled = !isEdit;
    el.userActive.checked = isEdit ? targetUser.active !== false : true;
  }
  if (el.userSalesTargetEnabled) {
    el.userSalesTargetEnabled.checked = isEdit ? targetUser.salesTargetEnabled === true : false;
  }
  if (el.userRole) {
    el.userRole.value = isEdit ? normalizeUserRole(targetUser.role) : USER_ROLE.SELLER;
  }
}

function loadUserInModal(userId) {
  const target = (state.users || []).map(normalizeUserRecord).find((u) => u.id === userId);
  if (!target) return;
  userModalEditingId = target.id;
  userSignatureNormalizedDataUrl = "";
  if (el.userFullName) el.userFullName.value = target.fullName || "";
  if (el.userUsername) el.userUsername.value = target.username || "";
  if (el.userEmail) el.userEmail.value = target.email || "";
  if (el.userPhone) el.userPhone.value = target.phone || "";
  if (el.userRole) el.userRole.value = normalizeUserRole(target.role);
  if (el.userPassword) el.userPassword.value = "";
  if (el.userSignature) el.userSignature.value = "";
  renderUserSignaturePreview(target.signatureDataUrl || "");
  if (el.userAvatar) el.userAvatar.value = "";
  userMonthlyGoalsDraft = Array.isArray(target.monthlyGoals) ? deepClone(target.monthlyGoals) : [];
  renderUserMonthlyGoalsDraft();
  clearUserGoalEditorFields();
  setUserModalMode("edit", target);
}

function resetUserModalForm() {
  userModalEditingId = "";
  userSignatureNormalizedDataUrl = "";
  if (el.userName) el.userName.value = "";
  if (el.userFullName) el.userFullName.value = "";
  if (el.userUsername) el.userUsername.value = "";
  if (el.userEmail) el.userEmail.value = "";
  if (el.userPhone) el.userPhone.value = "";
  if (el.userRole) el.userRole.value = USER_ROLE.SELLER;
  if (el.userPassword) el.userPassword.value = "";
  if (el.userSignature) el.userSignature.value = "";
  renderUserSignaturePreview("");
  if (el.userAvatar) el.userAvatar.value = "";
  if (el.userSalesTargetEnabled) el.userSalesTargetEnabled.checked = false;
  userMonthlyGoalsDraft = [];
  renderUserMonthlyGoalsDraft();
  clearUserGoalEditorFields();
  renderUserEditSelect("");
  setUserModalMode("create");
}

function openUserModal(userId = "") {
  resetUserModalForm();
  if (userId) {
    renderUserEditSelect(userId);
    loadUserInModal(userId);
  }
  el.userBackdrop.hidden = false;
  setTimeout(() => (el.userFullName || el.userName)?.focus(), 0);
}

function closeUserModal() {
  el.userBackdrop.hidden = true;
  resetUserModalForm();
  restoreModuleScreenAfterModal();
}

// Abre el modal principal de cotizacion para un evento.
// Aqui se prepara el borrador, se asegura un codigo si hace falta,
// y se cargan en pantalla los datos que luego se podran guardar o versionar.
async function openQuoteModal(eventId) {
  const ev = state.events.find(x => x.id === eventId);
  if (!ev) return;
  const series = getEventSeries(ev).sort((a, b) => a.date.localeCompare(b.date));
  const firstDate = series[0]?.date || ev.date;
  const lastDate = series[series.length - 1]?.date || ev.date;
  const user = state.users.find(u => u.id === ev.userId);
  const existingQuote = ev.quote ? deepClone(ev.quote) : null;
  const existingVersions = normalizeQuoteVersionHistory(existingQuote?.versions);
  const existingVersionNumber = Number(existingQuote?.version);
  const currentVersion = Number.isFinite(existingVersionNumber) && existingVersionNumber > 0
    ? Math.floor(existingVersionNumber)
    : (existingQuote ? (existingVersions.length + 1) : 1);

  quoteDraft = existingQuote || {
    companyId: "",
    managerId: "",
    dueDate: "",
    docDate: "",
    paymentType: "",
    code: "",
    contact: "",
    email: "",
    billTo: "",
    address: "",
    eventType: "",
    venue: "",
    schedule: "",
    phone: "",
    nit: "",
    people: ev.pax ? String(ev.pax) : "",
    eventDate: "",
    folio: "",
    endDate: "",
    internalNotes: "",
    discountType: "AMOUNT",
    discountValue: 0,
    items: [],
    notes: "",
    templateId: CONTRACT_CORP_TEMPLATE_ID,
    currency: "GTQ",
  };
  quoteDraft.version = currentVersion;
  quoteDraft.versions = existingVersions;
  quoteSelectedItemIds = new Set();
  if (!String(quoteDraft.code || "").trim()) {
    const code = await requestServerQuoteCode();
    quoteDraft.code = code || buildQuoteCode();
  }
  quoteDraft.paymentType = quotePaymentTypesToStorage(quoteDraft.paymentType || "");
  quoteDraft.advances = normalizeQuoteAdvancesForSnapshot(quoteDraft.advances);
  quoteDraft.currency = normalizeQuoteCurrency(quoteDraft.currency);
  if (quoteDraft.companyId && !quoteDraft.managerId) {
    const cmp = (state.companies || []).find(c => c.id === quoteDraft.companyId);
    const byUserName = cmp?.managers?.find(m => m.name.toLowerCase() === String(user?.name || "").toLowerCase());
    quoteDraft.managerId = byUserName?.id || cmp?.managers?.[0]?.id || "";
  }
  if (existingQuote && !quoteDraft.managerId && existingQuote.manager) {
    const cmp = (state.companies || []).find(c => c.id === quoteDraft.companyId);
    const byName = cmp?.managers?.find(m => m.name.toLowerCase() === String(existingQuote.manager).toLowerCase());
    if (byName) quoteDraft.managerId = byName.id;
  }

  el.quoteEventId.value = ev.id;
  el.quoteSubtitle.textContent = `${ev.name} - ${ev.salon} - ${firstDate} a ${lastDate} - ${ev.startTime}-${ev.endTime}`;
  renderCompaniesSelect(quoteDraft.companyId);
  renderQuoteManagerSelect(quoteDraft.companyId, quoteDraft.managerId || null);
  renderQuoteTemplateSelect(quoteDraft.templateId || "");
  renderQuoteServiceTemplateSelect("");
  if (el.quoteServiceTemplateName) el.quoteServiceTemplateName.value = "";
  fillQuoteHeaderFields(true, !!existingQuote);
  renderQuoteCurrencyPicker();
  setQuotePaymentTypesOnForm(quoteDraft.paymentType || "");
  if (el.quoteDueDate) el.quoteDueDate.dataset.autoDefault = quoteDraft.dueDate ? "0" : "1";
  syncQuoteDateDefaults({ forceDueDate: !quoteDraft.dueDate });
  if (el.quoteDiscountType) el.quoteDiscountType.value = normalizeDiscountType(quoteDraft.discountType);
  if (el.quoteDiscountValue) el.quoteDiscountValue.value = String(Math.max(0, Number(quoteDraft.discountValue || 0)));
  renderQuoteServiceDateSelect();
  el.quoteServiceSearch.value = "";
  if (el.quoteServiceQty) {
    el.quoteServiceQty.value = "";
    el.quoteServiceQty.classList.remove("fieldInvalid");
  }
  syncQuoteServiceQuantityField();
  renderQuoteItems();
  syncPaxQuantityItems();
  renderQuoteVersionControls();
  setQuoteItemsExpanded(false);

  el.quoteBackdrop.hidden = false;
  document.body.classList.add("quoteModeOpen");
  setQuoteDocPanelVisible(false);
  if (!String(quoteDraft.templateId || "").trim()) {
    quoteDraft.templateId = CORPORATE_TEMPLATE_ID;
  }
  setTimeout(() => el.quoteServiceSearch.focus(), 0);
}

function closeQuoteModal() {
  setQuoteItemsExpanded(false);
  el.quoteBackdrop.hidden = true;
  document.body.classList.remove("quoteModeOpen");
  el.quoteBackdrop.classList.remove("docFloatOpen");
  setQuoteDocPanelVisible(false);
  closeQuoteAdvanceModal();
  closeMenuMontajeModal();
  closeMenuMontajeSelectableModal();
  closeServiceModal();
  quoteDraft = null;
  quoteSelectedItemIds = new Set();
  if (accountingReportReturnAfterPayment && el.accountingReportBackdrop) {
    renderAccountingReportTable();
    document.body.classList.remove("accountingPaymentMode");
    el.accountingReportBackdrop.hidden = false;
    if (accountingReturnAfterPaymentTarget === "statement" && accountingStatementActiveCompanyKey) {
      openAccountingStatementModal(accountingStatementActiveCompanyKey);
    }
    accountingReportReturnAfterPayment = false;
    accountingReturnAfterPaymentTarget = "";
  }
}

function persistQuoteDraftFinancials({ showToast = false } = {}) {
  if (!quoteDraft) return false;
  const eventId = String(el.quoteEventId?.value || "").trim();
  const ev = (state.events || []).find((item) => String(item.id || "") === eventId);
  if (!ev || !ev.quote) return false;
  const savedQuote = cloneQuoteSnapshot(quoteDraft, Math.max(1, Number(quoteDraft.version || ev.quote?.version || 1)));
  savedQuote.paymentType = quotePaymentTypesToStorage(savedQuote.paymentType || "");
  savedQuote.advances = normalizeQuoteAdvancesForSnapshot(savedQuote.advances);
  savedQuote.currency = normalizeQuoteCurrency(savedQuote.currency);
  savedQuote.companyName = savedQuote.companyName || ev.quote?.companyName || "";
  savedQuote.managerName = savedQuote.managerName || ev.quote?.managerName || "";
  savedQuote.managerPhone = savedQuote.managerPhone || ev.quote?.managerPhone || "";
  savedQuote.quotedAt = savedQuote.quotedAt || ev.quote?.quotedAt || new Date().toISOString();
  savedQuote.versions = normalizeQuoteVersionHistory(ev.quote?.versions || quoteDraft.versions);
  const totals = getQuoteTotals(savedQuote);
  savedQuote.subtotal = totals.subtotal;
  savedQuote.discountAmount = totals.discountAmount;
  savedQuote.total = totals.total;
  for (const item of getEventSeries(ev)) {
    item.quote = deepClone(savedQuote);
  }
  quoteDraft = deepClone(savedQuote);
  persist();
  render();
  if (el.accountingReportBackdrop && !el.accountingReportBackdrop.hidden) {
    renderAccountingReportTable();
  }
  if (showToast) toast("Pagos actualizados.");
  return true;
}

async function applyPaymentFromAccountingReport(eventId) {
  const id = String(eventId || "").trim();
  if (!id) return;
  const openedFromStatement = !!(el.accountStatementBackdrop && !el.accountStatementBackdrop.hidden);
  const openedFromReport = !!(el.accountingReportBackdrop && !el.accountingReportBackdrop.hidden);
  accountingReportReturnAfterPayment = openedFromStatement || openedFromReport;
  accountingReturnAfterPaymentTarget = openedFromStatement ? "statement" : (openedFromReport ? "report" : "");
  if (accountingReportReturnAfterPayment && el.accountingReportBackdrop) {
    document.body.classList.add("accountingPaymentMode");
    el.accountingReportBackdrop.hidden = true;
  }
  if (openedFromStatement) closeAccountingStatementModal();
  await openQuoteModal(id);
  if (el.quoteBackdrop) el.quoteBackdrop.hidden = false;
  openQuoteAdvanceModal();
  if (el.quoteAdvanceBackdrop) el.quoteAdvanceBackdrop.hidden = false;
  setTimeout(() => el.quoteAdvanceAmount?.focus(), 30);
}

async function openLatestQuoteFromAccounting(eventId) {
  const id = String(eventId || "").trim();
  if (!id) return;
  const ev = (state.events || []).find((item) => String(item.id || "") === id);
  if (!ev?.quote) return toast("Este evento no tiene cotizacion registrada.");
  accountingReportReturnAfterPayment = !!(el.accountingReportBackdrop && !el.accountingReportBackdrop.hidden);
  if (el.accountStatementBackdrop && !el.accountStatementBackdrop.hidden) {
    closeAccountingStatementModal();
  }
  if (accountingReportReturnAfterPayment && el.accountingReportBackdrop) {
    document.body.classList.add("accountingPaymentMode");
    el.accountingReportBackdrop.hidden = true;
  }
  await openQuoteModal(id);
  if (el.quoteBackdrop) el.quoteBackdrop.hidden = false;
  setQuoteDocPanelVisible(false);
  setQuoteItemsExpanded(false);
  setTimeout(() => el.quoteItemsPanel?.scrollIntoView({ behavior: "smooth", block: "start" }), 120);
}

function renderQuoteAdvanceHistoryLog() {
  if (!el.quoteAdvanceHistoryBody) return;
  const ctx = getCurrentQuoteHistoryContext();
  const inferAdvanceLogTone = (text = "") => {
    const lower = String(text || "").toLowerCase();
    if (lower.includes("eliminado")) return { key: "deleted", label: "Eliminado" };
    if (lower.includes("editado") || lower.includes("ajustado")) return { key: "edited", label: "Editado" };
    if (lower.includes("favor")) return { key: "credit", label: "Saldo a favor" };
    return { key: "added", label: "Agregado" };
  };

  let html = "";
  if (ctx) {
    ensureHistoryStore();
    const historyRows = (Array.isArray(state.changeHistory?.[ctx.key]) ? state.changeHistory[ctx.key] : [])
      .filter((row) => /(anticipo|abono|pago|saldo a favor)/i.test(String(row?.change || "")));
    html = historyRows.map((row) => {
      const tone = inferAdvanceLogTone(row?.change);
      return `
        <tr class="quoteAdvanceLogRow quoteAdvanceLogRow--${escapeHtml(tone.key)}">
          <td>${escapeHtml(formatDateTime(row.at))}</td>
          <td>${escapeHtml(row.actorName || "Sistema")}</td>
          <td>
            <div class="quoteAdvanceLogEntry">
              <span class="quoteAdvanceLogTag quoteAdvanceLogTag--${escapeHtml(tone.key)}">${escapeHtml(tone.label)}</span>
              <span>${escapeHtml(row.change || "")}</span>
            </div>
          </td>
        </tr>
      `;
    }).join("");
  }

  if (!html) {
    const advances = normalizeQuoteAdvancesForSnapshot(quoteDraft?.advances)
      .slice()
      .sort((a, b) => {
        const createdCmp = String(b.createdAt || "").localeCompare(String(a.createdAt || ""));
        if (createdCmp !== 0) return createdCmp;
        return String(b.date || "").localeCompare(String(a.date || ""));
      });
    html = advances.map((advance) => {
      const stamp = String(advance.createdAt || "").trim();
      const fallbackDate = String(advance.date || "").trim();
      const detail = [
        `Abono registrado: ${fallbackDate || "sin fecha"}`,
        formatQuoteAdvanceDetail(advance),
        moneyGT(advance.amount || 0),
      ].filter(Boolean).join(" | ");
      return `
        <tr class="quoteAdvanceLogRow quoteAdvanceLogRow--added">
          <td>${escapeHtml(stamp ? formatDateTime(stamp) : (fallbackDate || "-"))}</td>
          <td>${escapeHtml(advance.createdByName || "Sistema")}</td>
          <td>
            <div class="quoteAdvanceLogEntry">
              <span class="quoteAdvanceLogTag quoteAdvanceLogTag--added">Agregado</span>
              <span>${escapeHtml(detail)}</span>
            </div>
          </td>
        </tr>
      `;
    }).join("");
  }

  el.quoteAdvanceHistoryBody.innerHTML = html || `
    <tr class="quoteAdvanceLogEmptyRow">
      <td colspan="3">Sin movimientos de pagos registrados.</td>
    </tr>
  `;
}

function renderQuoteAdvancesModal() {
  const advanceWrap = document.querySelector("#quoteAdvanceBackdrop .quoteAdvanceTableWrap");
  if (!advanceWrap || !el.quoteAdvanceTotal) return;
  const advances = normalizeQuoteAdvancesForSnapshot(quoteDraft?.advances)
    .slice()
    .sort((a, b) => {
      const d = String(a.date || "").localeCompare(String(b.date || ""));
      if (d !== 0) return d;
      return String(a.createdAt || "").localeCompare(String(b.createdAt || ""));
    });
  const balance = getQuoteAdvanceBalanceSummary(quoteDraft || {});
  const rowsHtml = !advances.length
    ? `<div class="quoteAdvanceLedgerEmpty">Sin anticipos registrados.</div>`
    : advances.map((row) => {
      const evidenceAnchor = row.evidenceDataUrl
        ? `<a class="btn" href="${escapeHtml(row.evidenceDataUrl)}" download="${escapeHtml(row.evidenceName || `evidencia_${row.id}.pdf`)}" target="_blank" rel="noopener">Ver</a>`
        : `<span class="quoteAdvanceLedgerMuted">-</span>`;
      return `
        <div class="quoteAdvanceLedgerRow">
          <div class="quoteAdvanceLedgerCell"><span class="quoteAdvanceLedgerLabel">Fecha</span><span>${escapeHtml(row.date || "-")}</span></div>
          <div class="quoteAdvanceLedgerCell"><span class="quoteAdvanceLedgerLabel">Tipo</span><span>${escapeHtml(row.paymentType || "-")}</span></div>
          <div class="quoteAdvanceLedgerCell"><span class="quoteAdvanceLedgerLabel">No. boleta</span><span>${escapeHtml(row.voucherNumber || "-")}</span></div>
          <div class="quoteAdvanceLedgerCell"><span class="quoteAdvanceLedgerLabel">Descripcion</span><span>${escapeHtml(row.description || "-")}</span></div>
          <div class="quoteAdvanceLedgerCell"><span class="quoteAdvanceLedgerLabel">Monto</span><strong>${moneyGT(row.amount || 0)}</strong></div>
          <div class="quoteAdvanceLedgerCell quoteAdvanceLedgerCell--evidence"><span class="quoteAdvanceLedgerLabel">Evidencia</span>${evidenceAnchor}</div>
          <div class="quoteAdvanceLedgerCell quoteAdvanceLedgerCell--actions">
            <span class="quoteAdvanceLedgerLabel">Acciones</span>
            <div class="appointmentActions">
              <button class="apptIconBtn apptEdit quoteAdvanceEditBtn" type="button" data-advance-id="${escapeHtml(row.id)}" title="Editar" aria-label="Editar">&#9998;</button>
              <button class="apptIconBtn apptDelete quoteAdvanceRemoveBtn" type="button" data-advance-id="${escapeHtml(row.id)}" title="Eliminar" aria-label="Eliminar">&#128465;</button>
            </div>
          </div>
        </div>
      `;
    }).join("");

  advanceWrap.innerHTML = `
    <div class="quoteAdvanceLedgerHead">
      <span>Fecha</span>
      <span>Tipo</span>
      <span>No. boleta</span>
      <span>Descripcion</span>
      <span>Monto</span>
      <span>Evidencia</span>
      <span>Acciones</span>
    </div>
    <div class="quoteAdvanceLedgerBody">${rowsHtml}</div>
  `;
  el.quoteAdvanceBody = advanceWrap;
  el.quoteAdvanceTotal.textContent = moneyGT(balance.advancesTotal || 0);
  if (el.quoteAdvancePending) el.quoteAdvancePending.textContent = moneyGT(balance.balancePending || 0);
  if (el.quoteAdvanceCredit) el.quoteAdvanceCredit.textContent = moneyGT(balance.creditBalance || 0);
  renderQuoteAdvanceHistoryLog();
}
function resetQuoteAdvanceForm() {
  quoteAdvanceEditingId = "";
  if (el.quoteAdvanceAmount) el.quoteAdvanceAmount.value = "";
  if (el.quoteAdvancePaymentType) el.quoteAdvancePaymentType.value = "Efectivo";
  if (el.quoteAdvanceDate) el.quoteAdvanceDate.value = quoteDraft?.docDate || toISODate(new Date());
  if (el.quoteAdvanceVoucherNumber) el.quoteAdvanceVoucherNumber.value = "";
  if (el.quoteAdvanceDescription) el.quoteAdvanceDescription.value = "";
  if (el.quoteAdvanceEvidence) el.quoteAdvanceEvidence.value = "";
  if (el.quoteAdvanceEvidenceHint) el.quoteAdvanceEvidenceHint.textContent = "Sin archivo adjunto";
  if (el.btnQuoteAdvanceAdd) el.btnQuoteAdvanceAdd.textContent = "Agregar anticipo";
}

function startEditQuoteAdvance(advanceId) {
  if (!quoteDraft || !Array.isArray(quoteDraft.advances)) return;
  const id = String(advanceId || "").trim();
  if (!id) return;
  const item = quoteDraft.advances.find((x) => String(x?.id || "") === id);
  if (!item) return;
  quoteAdvanceEditingId = id;
  if (el.quoteAdvanceAmount) el.quoteAdvanceAmount.value = String(Number(item.amount || 0).toFixed(2));
  if (el.quoteAdvancePaymentType) el.quoteAdvancePaymentType.value = normalizeAdvancePaymentType(item.paymentType);
  if (el.quoteAdvanceDate) el.quoteAdvanceDate.value = String(item.date || "");
  if (el.quoteAdvanceVoucherNumber) el.quoteAdvanceVoucherNumber.value = String(item.voucherNumber || "");
  if (el.quoteAdvanceDescription) el.quoteAdvanceDescription.value = String(item.description || "");
  if (el.quoteAdvanceEvidence) el.quoteAdvanceEvidence.value = "";
  if (el.quoteAdvanceEvidenceHint) {
    el.quoteAdvanceEvidenceHint.textContent = item.evidenceName ? `Archivo actual: ${item.evidenceName}` : "Sin archivo adjunto";
  }
  if (el.btnQuoteAdvanceAdd) el.btnQuoteAdvanceAdd.textContent = "Guardar cambios";
}

async function addQuoteAdvanceFromForm() {
  if (!quoteDraft) return toast("Primero abre una cotizacion.");
  const amountRaw = String(el.quoteAdvanceAmount?.value || "").trim();
  const amount = Math.max(0, Number(amountRaw || 0));
  const paymentType = normalizeAdvancePaymentType(el.quoteAdvancePaymentType?.value || "");
  const date = String(el.quoteAdvanceDate?.value || "").trim();
  const voucherNumber = String(el.quoteAdvanceVoucherNumber?.value || "").trim();
  const description = String(el.quoteAdvanceDescription?.value || "").trim();
  const evidenceFile = el.quoteAdvanceEvidence?.files?.[0] || null;
  if (!amountRaw || Number.isNaN(Number(amountRaw)) || amount <= 0) return toast("Anticipo: el monto es obligatorio y debe ser mayor a 0.");
  if (!String(el.quoteAdvancePaymentType?.value || "").trim()) return toast("Anticipo: la forma de pago es obligatoria.");
  if (!date) return toast("Anticipo: la fecha es obligatoria.");
  if (!description) return toast("Anticipo: la descripcion es obligatoria.");
  if (paymentType !== "Efectivo" && !voucherNumber) return toast("Anticipo: el No. de boleta es obligatorio para este tipo de pago.");
  if (!Array.isArray(quoteDraft.advances)) quoteDraft.advances = [];

  let evidenceDataUrl = "";
  let evidenceName = "";
  let evidenceType = "";
  if (evidenceFile) {
    const maxBytes = 6 * 1024 * 1024;
    if (Number(evidenceFile.size || 0) > maxBytes) {
      return toast("Evidencia: el archivo supera 6 MB.");
    }
    evidenceDataUrl = await readImageFileAsDataUrl(evidenceFile);
    if (!evidenceDataUrl) return toast("No se pudo leer el archivo de evidencia.");
    evidenceName = String(evidenceFile.name || "evidencia").trim();
    evidenceType = String(evidenceFile.type || "").trim();
  }

  const ctx = getCurrentQuoteHistoryContext();
  if (quoteAdvanceEditingId) {
    const idx = quoteDraft.advances.findIndex((x) => String(x?.id || "") === quoteAdvanceEditingId);
    if (idx >= 0) {
      const prev = normalizeQuoteAdvanceForSnapshot(quoteDraft.advances[idx]);
      quoteDraft.advances[idx] = {
        ...quoteDraft.advances[idx],
        amount,
        paymentType,
        date,
        voucherNumber,
        description,
        evidenceDataUrl: evidenceDataUrl || String(prev.evidenceDataUrl || ""),
        evidenceName: evidenceName || String(prev.evidenceName || ""),
        evidenceType: evidenceType || String(prev.evidenceType || ""),
      };
      if (ctx) {
        appendHistoryByKey(
          ctx.key,
          authSession.userId || ctx.ev.userId || "",
          `Anticipo editado: ${String(prev?.date || "")} ${formatQuoteAdvanceDetail(prev)} ${moneyGT(prev?.amount || 0)} -> ${date} ${formatQuoteAdvanceDetail({ paymentType, voucherNumber, description })} ${moneyGT(amount || 0)}`
        );
      }
      toast("Anticipo actualizado.");
    }
  } else {
    quoteDraft.advances.push({
      id: uid(),
      amount,
      paymentType,
      date,
      voucherNumber,
      description,
      evidenceDataUrl,
      evidenceName,
      evidenceType,
      createdAt: new Date().toISOString(),
      createdByUserId: String(authSession.userId || ctx?.ev?.userId || "").trim(),
      createdByName: String(authSession.fullName || authSession.username || getUserNameById(authSession.userId || ctx?.ev?.userId || "") || "Sistema").trim(),
    });
    if (ctx) {
      appendHistoryByKey(
        ctx.key,
        authSession.userId || ctx.ev.userId || "",
        `Anticipo agregado: ${date} ${formatQuoteAdvanceDetail({ paymentType, voucherNumber, description })} ${moneyGT(amount || 0)}`
      );
    }
    toast("Anticipo agregado.");
  }
  resetQuoteAdvanceForm();
  renderQuoteAdvancesModal();
  persistQuoteDraftFinancials();
}

function removeQuoteAdvanceById(advanceId) {
  if (!quoteDraft || !Array.isArray(quoteDraft.advances)) return;
  const id = String(advanceId || "").trim();
  if (!id) return;
  const removed = quoteDraft.advances.find((x) => String(x?.id || "") === id);
  quoteDraft.advances = quoteDraft.advances.filter((x) => String(x?.id || "") !== id);
  if (quoteAdvanceEditingId && quoteAdvanceEditingId === id) {
    resetQuoteAdvanceForm();
  }
  renderQuoteAdvancesModal();
  persistQuoteDraftFinancials();
  const ctx = getCurrentQuoteHistoryContext();
  if (ctx && removed) {
    appendHistoryByKey(
      ctx.key,
      authSession.userId || ctx.ev.userId || "",
      `Anticipo eliminado: ${String(removed.date || "")} ${formatQuoteAdvanceDetail(removed)} ${moneyGT(removed.amount || 0)}`
    );
  }
  toast("Anticipo eliminado.");
}

function openQuoteAdvanceModal() {
  if (!quoteDraft) return toast("Primero abre una cotizacion.");
  if (!Array.isArray(quoteDraft.advances)) quoteDraft.advances = [];
  resetQuoteAdvanceForm();
  renderQuoteAdvancesModal();
  if (el.quoteAdvanceBackdrop) el.quoteAdvanceBackdrop.hidden = false;
}

function closeQuoteAdvanceModal() {
  resetQuoteAdvanceForm();
  if (el.quoteAdvanceBackdrop) el.quoteAdvanceBackdrop.hidden = true;
  if (accountingReportReturnAfterPayment && el.quoteBackdrop && !el.quoteBackdrop.hidden) {
    closeQuoteModal();
  }
}

function openAppointmentModal(eventId) {
  const ev = state.events.find(x => x.id === eventId);
  if (!ev || !el.appointmentBackdrop) return;
  appointmentTargetEventId = eventId;
  el.appointmentDate.value = ev.date || toISODate(new Date());
  el.appointmentTime.value = ev.startTime || "09:00";
  initModernTimePicker(el.appointmentTime);
  el.appointmentChannel.value = "Telefono";
  el.appointmentNotes.value = "";
  el.appointmentBackdrop.hidden = false;
  setTimeout(() => el.appointmentDate.focus(), 0);
}

function closeAppointmentModal() {
  if (!el.appointmentBackdrop) return;
  el.appointmentBackdrop.hidden = true;
  appointmentTargetEventId = null;
  if (el.appointmentForm) el.appointmentForm.reset();
}

function saveAppointmentFromForm() {
  const eventId = appointmentTargetEventId;
  if (!eventId) return;
  const ev = state.events.find(x => x.id === eventId);
  if (!ev) return;

  const date = String(el.appointmentDate.value || "").trim();
  const time = String(el.appointmentTime.value || "").trim();
  const channel = String(el.appointmentChannel.value || "").trim();
  const notes = String(el.appointmentNotes.value || "").trim();
  if (!date || !time || !channel) {
    return toast("Completa fecha, hora y medio de la cita.");
  }
  if (!isValidClockTime(time)) {
    return toast("Hora de cita invalida. Usa HH:mm.");
  }

  const key = reservationKeyFromEvent(ev);
  ensureReminderStore();
  const currentReminders = Array.isArray(state.reminders[key]) ? state.reminders[key] : [];
  const duplicate = currentReminders.some((r) =>
    r?.done !== true
    && String(r?.date || "").trim() === date
    && String(r?.time || "").trim() === time
    && String(r?.channel || "").trim().toLowerCase() === channel.toLowerCase()
    && String(r?.notes || "").trim() === notes
  );
  if (duplicate) {
    return toast("Sin cambios detectados en la cita.");
  }

  addReminderForEvent(ev, {
    date,
    time,
    channel,
    notes,
    createdByUserId: el.eventUser.value || ev.userId || "",
  });
  appendHistoryByKey(
    reservationKeyFromEvent(ev),
    el.eventUser.value || ev.userId || "",
    `Cita agregada: ${date} ${time} via ${channel}${notes ? ` (${notes})` : ""}.`
  );

  persist();
  render();
  runUpcomingReminderChecks();
  refreshTopbarReminders();
  closeAppointmentModal();
  openModalForEdit(eventId);
  toast("Cita agregada y recordatorio activo.");
}

function normalizeCompanyMatchText(value) {
  return String(value || "").trim().toLowerCase();
}

function quoteBelongsToCompany(quote, company) {
  if (!quote || !company) return false;
  const companyId = String(company.id || "").trim();
  const quoteCompanyId = String(quote.companyId || "").trim();
  if (companyId && quoteCompanyId && companyId === quoteCompanyId) return true;
  const companyName = normalizeCompanyMatchText(company.name);
  const quoteCompanyName = normalizeCompanyMatchText(quote.companyName);
  return !!companyName && !!quoteCompanyName && companyName === quoteCompanyName;
}

function quoteVersionCount(quote) {
  if (!quote || typeof quote !== "object") return 0;
  const versions = normalizeQuoteVersionHistory(quote.versions);
  const currentVersion = Math.max(1, Number(quote.version || (versions.length + 1)) || 1);
  const versionSet = new Set(versions.map((v) => String(Math.max(1, Number(v.version || 1)))));
  versionSet.add(String(currentVersion));
  return versionSet.size;
}

function collectCompanyRecord(company) {
  const stats = {
    quotedReservations: 0,
    quoteVersionsSent: 0,
    confirmed: 0,
    preReserved: 0,
    canceled: 0,
    lost: 0,
    topManagerName: "",
    topManagerEvents: 0,
    rows: [],
  };
  if (!company) return stats;

  const grouped = new Map();
  const managerCounter = new Map();
  for (const ev of state.events || []) {
    if (!ev) continue;
    const key = String(reservationKeyFromEvent(ev) || ev.id || "").trim();
    if (!key) continue;
    if (!grouped.has(key)) grouped.set(key, []);
    grouped.get(key).push(ev);
  }

  for (const [reservationKey, series] of grouped.entries()) {
    const ordered = series.slice().sort((a, b) => {
      const dateCmp = String(a.date || "").localeCompare(String(b.date || ""));
      if (dateCmp !== 0) return dateCmp;
      return compareTime(String(a.startTime || "00:00"), String(b.startTime || "00:00"));
    });
    const representative = ordered[0];
    const quoteHost = ordered.find((e) => quoteBelongsToCompany(e.quote, company));
    if (!quoteHost) continue;

    const currentStatus = String(representative?.status || "").trim();
    if (currentStatus === STATUS.CONFIRMADO) stats.confirmed += 1;
    if (currentStatus === STATUS.PRERESERVA) stats.preReserved += 1;
    if (currentStatus === STATUS.CANCELADO) stats.canceled += 1;
    if (currentStatus === STATUS.PERDIDO) stats.lost += 1;

    const quote = quoteHost.quote || {};
    const versionCount = quoteVersionCount(quote);
    stats.quotedReservations += 1;
    stats.quoteVersionsSent += versionCount;
    const managerId = String(quote.managerId || "").trim();
    const managerById = managerId
      ? (company.managers || []).find((m) => String(m.id || "") === managerId)
      : null;
    const managerName = String(
      managerById?.name
      || quote.manager
      || quote.contact
      || "Sin encargado"
    ).trim();
    const currentManagerCount = Number(managerCounter.get(managerName) || 0);
    managerCounter.set(managerName, currentManagerCount + 1);

    const latestSnapshot = getLatestQuoteSnapshotForEvent(quoteHost) || quote || {};
    const totals = getQuoteTotals(latestSnapshot);
    const sellerName = getUserNameById(representative?.userId || "");
    stats.rows.push({
      reservationKey,
      code: String(quote.code || latestSnapshot.code || reservationKey || "-"),
      versionCount,
      docDate: String(latestSnapshot.docDate || quote.docDate || representative?.date || ""),
      eventName: String(representative?.name || "Reserva"),
      status: currentStatus || "-",
      total: Number(totals.total || 0),
      sellerName: String(sellerName || "").trim(),
      managerName,
    });
  }

  stats.rows.sort((a, b) => String(b.docDate || "").localeCompare(String(a.docDate || "")));
  for (const [managerName, total] of managerCounter.entries()) {
    const n = Number(total || 0);
    if (n > stats.topManagerEvents) {
      stats.topManagerEvents = n;
      stats.topManagerName = managerName;
    }
  }
  return stats;
}

function renderCompanyRecord(company) {
  if (!el.companyRecordSection || !el.companyRecordSummary || !el.companyRecordBody) return;
  if (!company) {
    el.companyRecordSection.hidden = true;
    el.companyRecordSummary.innerHTML = "";
    el.companyRecordBody.innerHTML = "";
    return;
  }

  const record = collectCompanyRecord(company);
  el.companyRecordSection.hidden = false;
  el.companyRecordSummary.innerHTML = `
    <span class="pill">Cotizaciones (reservas): ${escapeHtml(String(record.quotedReservations))}</span>
    <span class="pill">Versiones cotizadas: ${escapeHtml(String(record.quoteVersionsSent))}</span>
    <span class="pill">Confirmados: ${escapeHtml(String(record.confirmed))}</span>
    <span class="pill">Pre reserva: ${escapeHtml(String(record.preReserved))}</span>
    <span class="pill">Cancelados: ${escapeHtml(String(record.canceled))}</span>
    <span class="pill">Perdidos: ${escapeHtml(String(record.lost))}</span>
    <span class="pill">Encargado top: ${escapeHtml(record.topManagerName || "-")} (${escapeHtml(String(record.topManagerEvents || 0))})</span>
  `;

  el.companyRecordBody.innerHTML = "";
  if (!record.rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="6">Sin cotizaciones registradas para esta empresa.</td>`;
    el.companyRecordBody.appendChild(tr);
    return;
  }

  for (const row of record.rows.slice(0, 80)) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(row.code || "-")}</td>
      <td>${escapeHtml(String(row.versionCount || 0))}</td>
      <td>${escapeHtml(row.docDate || "-")}</td>
      <td>${escapeHtml(row.eventName || "-")}</td>
      <td>${escapeHtml(row.status || "-")}</td>
      <td>Q ${escapeHtml(Number(row.total || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }))}</td>
    `;
    el.companyRecordBody.appendChild(tr);
  }
}

function resetCompanyModalState() {
  editingCompanyId = "";
  if (el.companyEditSelect) el.companyEditSelect.value = "";
  if (el.companyActive) el.companyActive.checked = true;
  if (el.btnCompanyDisable) {
    el.btnCompanyDisable.disabled = true;
    el.btnCompanyDisable.textContent = "Inhabilitar";
  }
  if (el.companyTitle) el.companyTitle.textContent = "Nueva empresa";
  editingCompanyManagerId = "";
  companyManagersDraft = [];
  el.companyName.value = "";
  el.companyOwner.value = "";
  el.companyEmail.value = "";
  el.companyNIT.value = "";
  el.companyBusinessName.value = "";
  el.companyEventType.value = "";
  el.companyAddress.value = "";
  el.companyPhone.value = "";
  el.companyNotes.value = "";
  renderCompanyRecord(null);
  el.managerName.value = "";
  el.managerPhone.value = "";
  el.managerEmail.value = "";
  if (el.btnAddManager) el.btnAddManager.textContent = "+ Encargado";
  el.managerAddress.value = "";
  renderCompanyManagersDraft();
}

function renderCompanyEditSelect(selectedId = "") {
  if (!el.companyEditSelect) return;
  const keep = String(selectedId || "").trim();
  const rows = (state.companies || [])
    .slice()
    .sort((a, b) => String(a?.name || "").localeCompare(String(b?.name || ""), "es", { sensitivity: "base" }));
  el.companyEditSelect.innerHTML = `<option value="">Crear nueva empresa</option>`;
  for (const company of rows) {
    const opt = document.createElement("option");
    opt.value = String(company?.id || "").trim();
    if (!opt.value) continue;
    opt.textContent = `${String(company?.name || "Empresa").trim()}${isCompanyDisabled(opt.value) ? " (Inhabilitada)" : ""}`;
    el.companyEditSelect.appendChild(opt);
  }
  el.companyEditSelect.value = keep;
}

function updateCompanyModalStateControls(companyId = "") {
  const id = String(companyId || "").trim();
  const isEdit = !!id;
  if (el.companyActive) {
    el.companyActive.checked = isEdit ? !isCompanyDisabled(id) : true;
  }
  if (el.btnCompanyDisable) {
    el.btnCompanyDisable.disabled = !isEdit;
    el.btnCompanyDisable.textContent = isEdit && isCompanyDisabled(id) ? "Reactivar" : "Inhabilitar";
  }
}

function loadCompanyInModal(companyId = "") {
  const id = String(companyId || "").trim();
  editingCompanyId = id;
  editingCompanyManagerId = "";
  const target = id
    ? (state.companies || []).find((c) => String(c.id) === id)
    : null;
  companyManagersDraft = target ? deepClone(target.managers || []) : [];
  renderCompanyManagersDraft();
  renderCompanyEditSelect(id);
  updateCompanyModalStateControls(id);
  if (target) {
    if (el.companyTitle) el.companyTitle.textContent = "Editar empresa";
    el.companyName.value = String(target.name || "");
    el.companyOwner.value = String(target.owner || "");
    el.companyEmail.value = String(target.email || "");
    el.companyNIT.value = String(target.nit || "");
    el.companyBusinessName.value = String(target.billTo || target.businessName || target.name || "");
    el.companyEventType.value = String(target.eventType || "");
    el.companyAddress.value = String(target.address || "");
    el.companyPhone.value = String(target.phone || "");
    el.companyNotes.value = String(target.notes || "");
    renderCompanyRecord(target);
  } else {
    if (el.companyTitle) el.companyTitle.textContent = "Nueva empresa";
    el.companyName.value = "";
    el.companyOwner.value = "";
    el.companyEmail.value = "";
    el.companyNIT.value = "";
    el.companyBusinessName.value = "";
    el.companyEventType.value = "";
    el.companyAddress.value = "";
    el.companyPhone.value = "";
    el.companyNotes.value = "";
    renderCompanyRecord(null);
  }
  el.managerName.value = "";
  el.managerPhone.value = "";
  if (el.btnAddManager) el.btnAddManager.textContent = "+ Encargado";
  el.managerEmail.value = "";
  el.managerAddress.value = "";
}

function openCompanyModal(companyId = "") {
  const rawId = (companyId && typeof companyId === "object" && "type" in companyId)
    ? ""
    : companyId;
  resetCompanyModalState();
  loadCompanyInModal(String(rawId || "").trim());
  el.companyBackdrop.hidden = false;
  setTimeout(() => el.companyName.focus(), 0);
}

function closeCompanyModal() {
  el.companyBackdrop.hidden = true;
  resetCompanyModalState();
  restoreModuleScreenAfterModal();
}

function renderServiceManagerTable() {
  if (!el.servicesManagerBody) return;
  const rows = (state.services || [])
    .slice()
    .sort((a, b) => String(a?.name || "").localeCompare(String(b?.name || ""), "es", { sensitivity: "base" }));
  el.servicesManagerBody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="6">Sin servicios registrados.</td>`;
    el.servicesManagerBody.appendChild(tr);
    return;
  }
  for (const s of rows) {
    const serviceId = String(s?.id || "").trim();
    const disabled = serviceId ? isServiceDisabled(serviceId) : false;
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(String(s?.name || ""))}</td>
      <td>${escapeHtml(String(s?.category || "-"))}</td>
      <td>${escapeHtml(String(s?.subcategory || "-"))}</td>
      <td>${escapeHtml(Number(s?.price || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }))}</td>
      <td>${disabled ? "Inhabilitado" : "Activo"}</td>
      <td class="tableActionsCell">
        <button class="apptIconBtn apptEdit" type="button" data-service-row-edit="${escapeHtml(serviceId)}" title="Editar" aria-label="Editar">&#9998;</button>
      </td>`;
    el.servicesManagerBody.appendChild(tr);
  }
}

function openServiceModal(serviceId = "") {
  if (!el.serviceBackdrop) return;
  if (!catalogoCategoriasServicio.length) {
    syncServiceCatalogFromDb().catch(() => { });
  }
  const serviceIdText = String(serviceId || "").trim();
  editingServiceId = "";
  if (el.serviceEditSelect) {
    const rows = (state.services || [])
      .slice()
      .sort((a, b) => String(a?.name || "").localeCompare(String(b?.name || ""), "es", { sensitivity: "base" }));
    el.serviceEditSelect.innerHTML = `<option value="">Crear nuevo servicio</option>`;
    for (const service of rows) {
      const opt = document.createElement("option");
      opt.value = String(service?.id || "").trim();
      if (!opt.value) continue;
      opt.textContent = `${String(service?.name || "Servicio").trim()}${isServiceDisabled(opt.value) ? " (Inhabilitado)" : ""}`;
      el.serviceEditSelect.appendChild(opt);
    }
    el.serviceEditSelect.value = serviceIdText;
  }
  renderServiceManagerTable();
  renderCategoriasServicioSelect();
  el.serviceBackdrop.hidden = false;
  const target = serviceIdText
    ? (state.services || []).find((s) => String(s.id) === serviceIdText)
    : null;
  editingServiceId = serviceIdText;
  if (el.serviceActive) {
    el.serviceActive.checked = serviceIdText ? !isServiceDisabled(serviceIdText) : true;
  }
  if (el.btnServiceDisable) {
    el.btnServiceDisable.disabled = !serviceIdText;
    el.btnServiceDisable.textContent = serviceIdText && isServiceDisabled(serviceIdText) ? "Reactivar" : "Inhabilitar";
  }
  if (target) {
    if (el.serviceTitle) el.serviceTitle.textContent = "Editar servicio";
    el.serviceName.value = String(target.name || "").trim();
    el.serviceCategory.value = String(target.categoryId || "");
    renderSubcategoriasServicioSelect(Number(target.categoryId || NaN));
    el.serviceSubcategory.value = String(target.subcategoryId || "");
    el.servicePrice.value = String(Number(target.price || 0));
    el.serviceQuantityMode.value = String(target.quantityMode || "MANUAL").toUpperCase() === "PAX" ? "PAX" : "MANUAL";
    el.serviceDescription.value = String(target.description || "");
  } else {
    if (el.serviceTitle) el.serviceTitle.textContent = "Nuevo servicio";
    el.serviceName.value = String(el.quoteServiceSearch?.value || "").trim();
    el.serviceCategory.value = "";
    el.serviceSubcategory.value = "";
    renderSubcategoriasServicioSelect(Number.NaN);
    el.servicePrice.value = "";
    el.serviceQuantityMode.value = "MANUAL";
    el.serviceDescription.value = "";
  }
  setTimeout(() => el.serviceName.focus(), 0);
}

function closeServiceModal() {
  if (!el.serviceBackdrop) return;
  el.serviceBackdrop.hidden = true;
  editingServiceId = "";
  if (el.serviceEditSelect) el.serviceEditSelect.value = "";
  if (el.serviceActive) el.serviceActive.checked = true;
  if (el.btnServiceDisable) {
    el.btnServiceDisable.disabled = true;
    el.btnServiceDisable.textContent = "Inhabilitar";
  }
  if (el.serviceTitle) el.serviceTitle.textContent = "Nuevo servicio";
  if (el.serviceForm) el.serviceForm.reset();
  if (el.servicesManagerBody) el.servicesManagerBody.innerHTML = "";
  restoreModuleScreenAfterModal();
}

function saveServiceFromForm() {
  const name = String(el.serviceName.value || "").trim();
  const categoryIdRaw = String(el.serviceCategory.value || "").trim();
  const subcategoryIdRaw = String(el.serviceSubcategory.value || "").trim();
  const categoryId = Number(categoryIdRaw);
  const subcategoryId = Number(subcategoryIdRaw);
  const category = sanitizeCatalogOptionLabel(selectedOptionText(el.serviceCategory));
  const subcategory = sanitizeCatalogOptionLabel(selectedOptionText(el.serviceSubcategory));
  const price = Math.max(0, Number(el.servicePrice.value || 0));
  const quantityMode = String(el.serviceQuantityMode.value || "MANUAL").trim().toUpperCase() === "PAX" ? "PAX" : "MANUAL";
  const description = String(el.serviceDescription.value || "").trim();

  if (!name) return toast("Nombre de servicio es obligatorio.");
  if (!Number.isFinite(categoryId) || categoryId <= 0) return toast("Categoria es obligatoria.");
  if (!Number.isFinite(subcategoryId) || subcategoryId <= 0) return toast("Subcategoria es obligatoria.");
  if (!Number.isFinite(price) || price < 0) return toast("Precio base invalido.");

  const selectedCategory = catalogoCategoriasServicio.find((c) => Number(c?.id) === categoryId);
  if (selectedCategory?.activo === false) return toast("La categoria esta inhabilitada. Reactivala para usarla.");
  const selectedSubcategory = catalogoSubcategoriasServicio.find((s) => Number(s?.id) === subcategoryId);
  if (!selectedSubcategory || selectedSubcategory?.activo === false || Number(selectedSubcategory?.id_categoria) !== categoryId) {
    return toast("La subcategoria esta inhabilitada o no corresponde a la categoria.");
  }

  const editingId = String(editingServiceId || "").trim();
  const exists = (state.services || []).some(
    (s) => String(s.name || "").toLowerCase() === name.toLowerCase() && String(s.id || "") !== editingId
  );
  if (exists) return toast("Ya existe un servicio con ese nombre.");
  const payload = normalizeServiceRecord({
    id: editingId || uid(),
    name,
    price,
    description,
    categoryId,
    subcategoryId,
    category,
    subcategory,
    quantityMode,
  });
  if (editingId) {
    const idx = (state.services || []).findIndex((s) => String(s.id || "") === editingId);
    if (idx >= 0) {
      const statusChanged = !!el.serviceActive && el.serviceActive.checked !== !isServiceDisabled(editingId);
      if (areServicesEquivalent(state.services[idx], payload) && !statusChanged) {
        return toast("Sin cambios detectados en servicio.");
      }
      state.services[idx] = payload;
    }
  } else {
    state.services.push(payload);
  }

  if (el.serviceActive?.checked === false) {
    state.disabledServices = Array.from(new Set([...(state.disabledServices || []), payload.id]));
  } else {
    enableService(payload.id);
  }
  persist();
  renderServicesList();
  renderServiceManagerTable();
  editingServiceId = "";
  if (el.serviceTitle) el.serviceTitle.textContent = "Nuevo servicio";
  if (el.serviceEditSelect) {
    const rows = (state.services || [])
      .slice()
      .sort((a, b) => String(a?.name || "").localeCompare(String(b?.name || ""), "es", { sensitivity: "base" }));
    el.serviceEditSelect.innerHTML = `<option value="">Crear nuevo servicio</option>`;
    for (const service of rows) {
      const opt = document.createElement("option");
      opt.value = String(service?.id || "").trim();
      if (!opt.value) continue;
      opt.textContent = `${String(service?.name || "Servicio").trim()}${isServiceDisabled(opt.value) ? " (Inhabilitado)" : ""}`;
      el.serviceEditSelect.appendChild(opt);
    }
  }
  if (el.serviceActive) el.serviceActive.checked = true;
  if (el.btnServiceDisable) {
    el.btnServiceDisable.disabled = true;
    el.btnServiceDisable.textContent = "Inhabilitar";
  }
  if (el.serviceName) el.serviceName.value = "";
  if (el.serviceCategory) el.serviceCategory.value = "";
  renderSubcategoriasServicioSelect(Number.NaN);
  if (el.serviceSubcategory) el.serviceSubcategory.value = "";
  if (el.servicePrice) el.servicePrice.value = "";
  if (el.serviceQuantityMode) el.serviceQuantityMode.value = "MANUAL";
  if (el.serviceDescription) el.serviceDescription.value = "";
  if (el.quoteServiceSearch) el.quoteServiceSearch.value = "";
  setTimeout(() => {
    try { el.serviceName?.focus(); } catch (_) { }
  }, 0);
  toast(editingId ? "Servicio actualizado." : "Servicio creado.");
}

function addManagerToCompanyDraft() {
  const name = el.managerName.value.trim();
  const phone = el.managerPhone.value.trim();
  const email = el.managerEmail.value.trim();
  const address = el.managerAddress.value.trim();
  if (!name || !phone || !email) {
    return toast("Encargado requiere nombre, telefono y correo.");
  }
  if (!isValidEmail(email)) {
    return toast("Correo de encargado invalido.");
  }
  const editingId = String(editingCompanyManagerId || "").trim();
  if (editingId) {
    const target = companyManagersDraft.find((m) => String(m.id || "") === editingId);
    if (!target) {
      editingCompanyManagerId = "";
      if (el.btnAddManager) el.btnAddManager.textContent = "+ Encargado";
      return toast("Encargado a editar no encontrado.");
    }
    target.name = name;
    target.phone = phone;
    target.email = email;
    target.address = address;
    editingCompanyManagerId = "";
    toast("Encargado actualizado.");
  } else {
    companyManagersDraft.push({
      id: uid(),
      name,
      phone,
      email,
      address,
    });
    toast("Encargado agregado.");
  }
  el.managerName.value = "";
  el.managerPhone.value = "";
  el.managerEmail.value = "";
  el.managerAddress.value = "";
  if (el.btnAddManager) el.btnAddManager.textContent = "+ Encargado";
  renderCompanyManagersDraft();
}

function renderCompanyManagersDraft() {
  if (!el.managersBody) return;
  el.managersBody.innerHTML = "";
  for (const m of companyManagersDraft) {
    const isEditing = String(editingCompanyManagerId || "") === String(m.id || "");
    const tr = document.createElement("tr");
    if (isEditing) tr.classList.add("isEditingRow");
    tr.innerHTML = `
      <td>${escapeHtml(m.name)}</td>
      <td>${escapeHtml(m.phone)}</td>
      <td>${escapeHtml(m.email)}</td>
      <td>${escapeHtml(m.address || "")}</td>
      <td class="tableActionsCell">
        <button type="button" class="apptIconBtn apptEdit editManagerBtn" data-manager-id="${m.id}" title="Editar encargado" aria-label="Editar encargado">&#9998;</button>
        <button type="button" class="apptIconBtn apptDelete removeManagerBtn" data-manager-id="${m.id}" title="Eliminar encargado" aria-label="Eliminar encargado">&#8854;</button>
      </td>
    `;
    el.managersBody.appendChild(tr);
  }
}

function normalizeQuotePeopleValue() {
  const raw = String(quoteDraft?.people || el.quotePeople?.value || "").trim();
  const n = Number(raw);
  return Number.isFinite(n) && n > 0 ? Math.floor(n) : 0;
}

function normalizeDiscountType(rawType) {
  return String(rawType || "").toUpperCase() === "PERCENT" ? "PERCENT" : "AMOUNT";
}

const QUOTE_PAYMENT_TYPES = ["Credito", "Deposito", "Efectivo", "Tarjeta"];

function normalizeQuotePaymentTypes(rawValue) {
  const values = Array.isArray(rawValue)
    ? rawValue
    : String(rawValue || "").split("|").join(",").split(",");
  const out = [];
  for (const item of values) {
    const raw = String(item || "").trim();
    if (!raw) continue;
    if (raw.toLowerCase() === "mixto") {
      if (!out.includes("Credito")) out.push("Credito");
      if (!out.includes("Efectivo")) out.push("Efectivo");
      continue;
    }
    const found = QUOTE_PAYMENT_TYPES.find((x) => x.toLowerCase() === raw.toLowerCase());
    if (!found) continue;
    if (!out.includes(found)) out.push(found);
  }
  return out;
}

function quotePaymentTypesToStorage(rawValue) {
  return normalizeQuotePaymentTypes(rawValue).join(", ");
}

function getSelectedQuotePaymentTypesFromForm() {
  if (!el.quotePaymentType) return [];
  return normalizeQuotePaymentTypes(el.quotePaymentType.value);
}

function setQuotePaymentTypesOnForm(rawValue) {
  if (!el.quotePaymentType) return;
  const values = normalizeQuotePaymentTypes(rawValue);
  el.quotePaymentType.value = values.join(", ");
}

function addQuotePaymentTypeFromPicker() {
  if (!el.quotePaymentTypeSelect || !el.quotePaymentType) return;
  const picked = String(el.quotePaymentTypeSelect.value || "").trim();
  if (!picked) return;
  const current = normalizeQuotePaymentTypes(el.quotePaymentType.value);
  if (!current.includes(picked)) current.push(picked);
  el.quotePaymentType.value = current.join(", ");
}

function clearQuotePaymentTypesOnForm() {
  if (!el.quotePaymentType) return;
  el.quotePaymentType.value = "";
}

function normalizeAdvancePaymentType(rawType) {
  const value = String(rawType || "").trim().toLowerCase();
  if (value === "credito") return "Credito";
  if (value === "deposito") return "Deposito";
  if (value === "efectivo") return "Efectivo";
  if (value === "tarjeta") return "Tarjeta";
  return "Efectivo";
}

function normalizeQuoteAdvanceForSnapshot(rawAdvance, index = 0) {
  const amountRaw = Number(rawAdvance?.amount || 0);
  const evidenceDataUrl = String(rawAdvance?.evidenceDataUrl || rawAdvance?.evidence?.dataUrl || "").trim();
  return {
    id: String(rawAdvance?.id || `adv_${index + 1}`).trim() || `adv_${index + 1}`,
    amount: Number.isFinite(amountRaw) ? Math.max(0, amountRaw) : 0,
    paymentType: normalizeAdvancePaymentType(rawAdvance?.paymentType),
    date: String(rawAdvance?.date || "").trim(),
    voucherNumber: String(rawAdvance?.voucherNumber || rawAdvance?.boleta || "").trim(),
    description: String(rawAdvance?.description || "").trim(),
    evidenceName: String(rawAdvance?.evidenceName || rawAdvance?.evidence?.name || "").trim(),
    evidenceType: String(rawAdvance?.evidenceType || rawAdvance?.evidence?.type || "").trim(),
    evidenceDataUrl,
    createdAt: String(rawAdvance?.createdAt || "").trim(),
  };
}

function formatQuoteAdvanceDetail(advanceLike) {
  const paymentType = String(advanceLike?.paymentType || "Anticipo").trim() || "Anticipo";
  const voucher = String(advanceLike?.voucherNumber || "").trim();
  const description = String(advanceLike?.description || "").trim();
  return [paymentType, voucher ? `Boleta ${voucher}` : "", description].filter(Boolean).join(" | ");
}

function normalizeQuoteAdvancesForSnapshot(rawAdvances) {
  const list = Array.isArray(rawAdvances) ? rawAdvances : [];
  return list
    .map((item, idx) => normalizeQuoteAdvanceForSnapshot(item, idx))
    .filter((item) => item.amount >= 0);
}

function getQuoteAdvanceTotal(quoteLike) {
  const advances = normalizeQuoteAdvancesForSnapshot(quoteLike?.advances);
  return advances.reduce((acc, item) => acc + Number(item.amount || 0), 0);
}

function getQuoteAdvanceBalanceSummary(quoteLike) {
  const advancesTotal = getQuoteAdvanceTotal(quoteLike);
  const totals = getQuoteTotals(quoteLike || {});
  const total = Math.max(0, Number(totals?.total ?? quoteLike?.total ?? 0));
  const delta = total - advancesTotal;
  return {
    total,
    advancesTotal,
    balancePending: Math.max(0, delta),
    creditBalance: Math.max(0, -delta),
  };
}

function renderQuoteAccountStatement() {
  if (!el.quoteAccountSummary || !el.quoteAccountBody) return;
  const quoteLike = quoteDraft || {};
  const totals = getQuoteTotals(quoteLike);
  const balance = getQuoteAdvanceBalanceSummary(quoteLike);
  const currency = quoteDraft?.currency || "GTQ";
  const money = (n) => moneyQuote(n, currency);
  const eventId = String(el.quoteEventId?.value || "").trim();
  const ev = eventId ? (state.events || []).find((x) => String(x.id || "") === eventId) : null;
  const advances = normalizeQuoteAdvancesForSnapshot(quoteLike.advances)
    .slice()
    .sort((a, b) => {
      const d = String(a.date || "").localeCompare(String(b.date || ""));
      if (d !== 0) return d;
      return String(a.createdAt || "").localeCompare(String(b.createdAt || ""));
    });
  const dueDate = String(quoteLike?.dueDate || "").trim();
  const docDate = String(quoteLike?.docDate || "").trim();
  const startDate = String(quoteLike?.eventDate || ev?.date || "").trim();
  const endDate = String(quoteLike?.endDate || startDate || ev?.date || "").trim();
  const eventName = String(ev?.name || quoteLike?.eventName || "Evento sin nombre").trim() || "Evento sin nombre";
  const eventSalon = String(ev?.salon || quoteLike?.salon || "").trim();
  const due = parseQuoteDate(dueDate);
  const today = stripTime(new Date());
  let dueMeta = "Sin fecha maxima de pago";
  if (due) {
    const diffDays = Math.round((due.getTime() - today.getTime()) / 86400000);
    if (diffDays < 0) dueMeta = `Vencido hace ${Math.abs(diffDays)} ${Math.abs(diffDays) === 1 ? "dia" : "dias"}`;
    else if (diffDays === 0) dueMeta = "Vence hoy";
    else dueMeta = `Vence en ${diffDays} ${diffDays === 1 ? "dia" : "dias"}`;
  }

  if (el.quoteAccountDocument) {
    const code = String(quoteLike?.code || el.quoteCode?.value || "").trim();
    el.quoteAccountDocument.textContent = code ? `Estado de cuenta ${code}` : "Estado de cuenta sin codigo";
  }
  if (el.quoteAccountEvent) {
    el.quoteAccountEvent.textContent = eventSalon ? `${eventName} | ${eventSalon}` : eventName;
  }
  if (el.quoteAccountRange) {
    if (startDate && endDate) {
      el.quoteAccountRange.textContent = startDate === endDate
        ? `Fecha del evento: ${startDate}`
        : `Periodo del evento: ${startDate} al ${endDate}`;
    } else if (startDate) {
      el.quoteAccountRange.textContent = `Fecha del evento: ${startDate}`;
    } else {
      el.quoteAccountRange.textContent = "Sin rango de fechas definido";
    }
  }

  const cards = [
    { tone: "primary", label: "Cargo total", value: money(totals.total), meta: `Subtotal ${money(totals.subtotal)}` },
    { tone: "info", label: "Abonos recibidos", value: money(balance.advancesTotal), meta: `${advances.length} movimiento(s)` },
    { tone: "accent", label: "Pendiente por cobrar", value: money(balance.balancePending), meta: dueMeta },
    { tone: "success", label: "Saldo a favor", value: money(balance.creditBalance), meta: dueDate ? `Fecha maxima ${dueDate}` : "Sin fecha maxima" },
  ];
  el.quoteAccountSummary.innerHTML = cards.map((card) => `
    <article class="salesSummaryCard salesSummaryCard--${escapeHtml(card.tone)} quoteAccountCard">
      <small>${escapeHtml(card.label)}</small>
      <strong>${escapeHtml(card.value)}</strong>
      <div class="salesSummaryMeta">${escapeHtml(card.meta)}</div>
    </article>
  `).join("");

  let runningBalance = totals.total;
  const rows = [
    {
      type: "Cargo",
      date: docDate || startDate || "-",
      detail: `Cotizacion emitida${dueDate ? ` | Fecha maxima pago ${dueDate}` : ""}`,
      charge: totals.total,
      payment: 0,
      balance: runningBalance,
      tone: "charge",
    },
    ...advances.map((advance) => {
      runningBalance -= Number(advance.amount || 0);
      return {
        type: "Abono",
        date: advance.date || "-",
        detail: formatQuoteAdvanceDetail(advance),
        charge: 0,
        payment: Number(advance.amount || 0),
        balance: runningBalance,
        tone: "payment",
      };
    }),
    {
      type: balance.creditBalance > 0 ? "Saldo a favor" : "Pendiente",
      date: dueDate || endDate || startDate || "-",
      detail: balance.creditBalance > 0 ? "Disponible para aplicar en futuras reservas" : "Monto pendiente de cobro",
      charge: 0,
      payment: 0,
      balance: balance.creditBalance > 0 ? -balance.creditBalance : balance.balancePending,
      tone: balance.creditBalance > 0 ? "credit" : "pending",
    },
  ];

  el.quoteAccountBody.innerHTML = rows.map((row) => {
    const saldoTexto = row.balance < 0 ? `${money(Math.abs(row.balance))} a favor` : money(row.balance);
    return `
      <tr class="quoteAccountRow quoteAccountRow--${escapeHtml(row.tone)}">
        <td>${escapeHtml(row.type)}</td>
        <td>${escapeHtml(row.date)}</td>
        <td>${escapeHtml(row.detail)}</td>
        <td>${row.charge > 0 ? escapeHtml(money(row.charge)) : "-"}</td>
        <td>${row.payment > 0 ? escapeHtml(money(row.payment)) : "-"}</td>
        <td>${escapeHtml(saldoTexto)}</td>
      </tr>
    `;
  }).join("");

  if (el.quoteAccountMeta) {
    el.quoteAccountMeta.textContent = advances.length
      ? `Ultimo abono ${advances[advances.length - 1]?.date || '-'} | Saldo actual ${balance.creditBalance > 0 ? `${money(balance.creditBalance)} a favor` : money(balance.balancePending)}`
      : `Sin abonos registrados | Saldo actual ${balance.creditBalance > 0 ? `${money(balance.creditBalance)} a favor` : money(balance.balancePending)}`;
  }
}

// Calcula subtotal, descuento, impuestos y total final de una cotizacion.
// Centralizar este calculo evita diferencias entre pantalla, PDF y versiones guardadas.
function getQuoteTotals(quoteLike) {
  const q = quoteLike || {};
  const items = Array.isArray(q.items) ? q.items : [];
  const subtotal = items.reduce((acc, x) => acc + Number(x.qty || 0) * Number(x.price || 0), 0);
  const discountType = normalizeDiscountType(q.discountType);
  const rawValue = Math.max(0, Number(q.discountValue || 0));
  const discountValue = Number.isFinite(rawValue) ? rawValue : 0;
  const discountAmount = discountType === "PERCENT"
    ? Math.max(0, Math.min(subtotal, subtotal * Math.min(100, discountValue) / 100))
    : Math.max(0, Math.min(subtotal, discountValue));
  const total = Math.max(0, subtotal - discountAmount);
  return { subtotal, discountType, discountValue, discountAmount, total };
}

function syncPaxQuantityItems() {
  if (!quoteDraft || !Array.isArray(quoteDraft.items)) return;
  const pax = normalizeQuotePeopleValue();
  let changed = false;
  for (const item of quoteDraft.items) {
    if (String(item.quantityMode || "").toUpperCase() !== "PAX") continue;
    if (!Number.isFinite(Number(item.unitPrice))) {
      item.unitPrice = Number(item.price || 0);
    }
    const nextQty = 1;
    const nextPrice = Math.max(0, Number(item.unitPrice || 0) * Math.max(0, pax));
    if (Number(item.qty || 0) !== nextQty || Number(item.price || 0) !== nextPrice) {
      item.qty = nextQty;
      item.price = nextPrice;
      changed = true;
    }
  }
  if (changed) renderQuoteItems();
}

function addServiceToQuoteDraft(rawName) {
  if (!quoteDraft) return;
  const name = String(rawName || "").trim();
  if (!name) return toast("Escribe un servicio.");
  const selectedServiceDate = String(el.quoteServiceDate?.value || "").trim();
  if (!selectedServiceDate) return toast("Selecciona la fecha del servicio.");
  const service = resolveServiceFromSearch(name);
  const rangeDates = getQuoteRangeDates();
  if (rangeDates.length && !rangeDates.includes(selectedServiceDate)) {
    return toast("La fecha del servicio no pertenece al rango del evento.");
  }
  const defaultServiceDate = selectedServiceDate;
  const paxQty = normalizeQuotePeopleValue();
  const requestedQty = Math.max(0, Math.floor(Number(el.quoteServiceQty?.value || 0)));
  const isPaxService = String(service?.quantityMode || "").trim().toUpperCase() === "PAX";
  if (!isPaxService && requestedQty <= 0) {
    if (el.quoteServiceQty) {
      el.quoteServiceQty.classList.add("fieldInvalid");
      el.quoteServiceQty.focus();
      el.quoteServiceQty.select?.();
    }
    return toast("Este producto requiere una cantidad manual antes de agregarlo.");
  }
  const item = service
    ? {
      rowId: uid(),
      serviceId: service.id,
      name: service.name,
      qty: isPaxService ? 1 : requestedQty,
      unitPrice: Number(service.price || 0),
      price: isPaxService
        ? Math.max(0, Number(service.price || 0) * Math.max(0, paxQty))
        : Number(service.price || 0),
      description: service.description || "",
      category: service.category || "",
      subcategory: service.subcategory || "",
      quantityMode: service.quantityMode || "MANUAL",
      serviceDate: defaultServiceDate,
    }
    : {
      rowId: uid(),
      serviceId: null,
      name,
      qty: requestedQty,
      unitPrice: 0,
      price: 0,
      description: "",
      category: "",
      subcategory: "",
      quantityMode: "MANUAL",
      serviceDate: defaultServiceDate,
    };
  quoteDraft.items.push(item);
  el.quoteServiceSearch.value = "";
  if (el.quoteServiceQty) {
    el.quoteServiceQty.value = "";
    el.quoteServiceQty.classList.remove("fieldInvalid");
  }
  syncQuoteServiceQuantityField();
  renderQuoteItems();
  const ctx = getCurrentQuoteHistoryContext();
  if (ctx) {
    appendHistoryByKey(
      ctx.key,
      authSession.userId || ctx.ev.userId || "",
      `Servicio agregado a cotizacion: ${String(item.name || "").trim() || "(sin nombre)"}`
    );
  }
}

function renderQuoteCurrencyPicker() {
  if (!quoteDraft) return;
  quoteDraft.currency = normalizeQuoteCurrency(quoteDraft.currency);
  const currency = quoteDraft.currency;
  if (el.btnQuoteCurrencyGTQ) {
    el.btnQuoteCurrencyGTQ.classList.toggle("isActive", currency === "GTQ");
    el.btnQuoteCurrencyGTQ.setAttribute("aria-pressed", currency === "GTQ" ? "true" : "false");
  }
  if (el.btnQuoteCurrencyUSD) {
    el.btnQuoteCurrencyUSD.classList.toggle("isActive", currency === "USD");
    el.btnQuoteCurrencyUSD.setAttribute("aria-pressed", currency === "USD" ? "true" : "false");
  }
  if (el.quoteDiscountType) {
    const amountOption = Array.from(el.quoteDiscountType.options || []).find((opt) => String(opt.value || "") === "AMOUNT");
    if (amountOption) amountOption.textContent = `Monto (${getQuoteCurrencySymbol(currency)})`;
  }
}

function setQuoteDocPanelVisible(visible) {
  const shouldShow = !!visible;
  const docCard = el.quoteDocPanelCard;
  const docFold = el.quoteDocFold;
  if (docCard && !docCard.__homeParent) {
    docCard.__homeParent = docCard.parentNode;
    docCard.__homeNextSibling = docCard.nextSibling;
  }

  const attachDocPanelHome = () => {
    if (!docCard || !docCard.__homeParent) return;
    if (docCard.parentNode === docCard.__homeParent) return;
    if (docCard.__homeNextSibling && docCard.__homeNextSibling.parentNode === docCard.__homeParent) {
      docCard.__homeParent.insertBefore(docCard, docCard.__homeNextSibling);
    } else {
      docCard.__homeParent.appendChild(docCard);
    }
  };

  const attachDocPanelOverlay = () => {
    if (!docCard) return;
    if (docCard.parentNode !== document.body) {
      document.body.appendChild(docCard);
    }
  };

  if (el.quoteDocPanelCard) {
    if (shouldShow) {
      attachDocPanelOverlay();
      el.quoteDocPanelCard.hidden = false;
      el.quoteDocPanelCard.classList.add("quoteDocDetached", "isForceVisible");
      el.quoteDocPanelCard.classList.remove("isForceHidden");
      el.quoteDocPanelCard.style.setProperty("display", "block", "important");
      el.quoteDocPanelCard.style.setProperty("visibility", "visible", "important");
      el.quoteDocPanelCard.style.setProperty("opacity", "1", "important");
      el.quoteDocPanelCard.style.setProperty("position", "fixed", "important");
      el.quoteDocPanelCard.style.setProperty("top", "96px", "important");
      el.quoteDocPanelCard.style.setProperty("left", "280px", "important");
      el.quoteDocPanelCard.style.setProperty("right", "24px", "important");
      el.quoteDocPanelCard.style.setProperty("bottom", "88px", "important");
      el.quoteDocPanelCard.style.setProperty("max-height", "none", "important");
      el.quoteDocPanelCard.style.setProperty("overflow", "auto", "important");
      el.quoteDocPanelCard.style.setProperty("z-index", "10050", "important");
      el.quoteDocPanelCard.style.setProperty("background", "var(--ui-surface, #ffffff)", "important");
      el.quoteDocPanelCard.style.setProperty("border", "1px solid var(--ui-border, #cdd8e7)", "important");
      el.quoteDocPanelCard.style.setProperty("border-radius", "12px", "important");
      el.quoteDocPanelCard.style.setProperty("box-shadow", "0 20px 45px rgba(2, 6, 23, 0.25)", "important");
      el.quoteDocPanelCard.style.setProperty("padding", "0", "important");
    } else {
      el.quoteDocPanelCard.hidden = true;
      el.quoteDocPanelCard.classList.remove("quoteDocDetached", "isForceVisible");
      el.quoteDocPanelCard.classList.add("isForceHidden");
      el.quoteDocPanelCard.style.setProperty("display", "none", "important");
      el.quoteDocPanelCard.style.setProperty("visibility", "hidden", "important");
      el.quoteDocPanelCard.style.setProperty("opacity", "0", "important");
      attachDocPanelHome();
      el.quoteDocPanelCard.style.removeProperty("position");
      el.quoteDocPanelCard.style.removeProperty("top");
      el.quoteDocPanelCard.style.removeProperty("left");
      el.quoteDocPanelCard.style.removeProperty("right");
      el.quoteDocPanelCard.style.removeProperty("bottom");
      el.quoteDocPanelCard.style.removeProperty("max-height");
      el.quoteDocPanelCard.style.removeProperty("overflow");
      el.quoteDocPanelCard.style.removeProperty("z-index");
      el.quoteDocPanelCard.style.removeProperty("background");
      el.quoteDocPanelCard.style.removeProperty("border");
      el.quoteDocPanelCard.style.removeProperty("border-radius");
      el.quoteDocPanelCard.style.removeProperty("box-shadow");
      el.quoteDocPanelCard.style.removeProperty("padding");
    }
  }
  if (el.quoteBackdrop) {
    el.quoteBackdrop.classList.toggle("quoteDocVisible", shouldShow);
  }
  if (el.quoteDocFold) {
    el.quoteDocFold.open = shouldShow;
    el.quoteDocFold.style.setProperty("position", "static", "important");
    el.quoteDocFold.style.setProperty("transform", "none", "important");
    el.quoteDocFold.style.setProperty("inset", "auto", "important");
    el.quoteDocFold.style.setProperty("max-height", "none", "important");
    el.quoteDocFold.style.setProperty("overflow", "visible", "important");
  }
  syncQuoteDocFoldUi();
}

function syncQuoteDocFoldUi() {
  if (!el.btnToggleQuoteDoc) return;
  const isVisible = el.quoteDocPanelCard ? !el.quoteDocPanelCard.hidden : !!el.quoteDocFold?.open;
  el.btnToggleQuoteDoc.textContent = isVisible ? "Ocultar datos" : "Datos empresa";
  el.btnToggleQuoteDoc.setAttribute("aria-expanded", isVisible ? "true" : "false");
}

function syncQuoteServiceQuantityField() {
  if (!el.quoteServiceQty) return;
  const matched = resolveServiceFromSearch(el.quoteServiceSearch?.value || "");
  const isPaxMode = String(matched?.quantityMode || "").trim().toUpperCase() === "PAX";
  if (isPaxMode) {
    el.quoteServiceQty.value = "1";
    el.quoteServiceQty.disabled = true;
    el.quoteServiceQty.readOnly = true;
    el.quoteServiceQty.placeholder = "Automatico por pax";
    el.quoteServiceQty.classList.remove("fieldInvalid");
    if (el.quoteServiceQtyHint) {
      el.quoteServiceQtyHint.textContent = "Este servicio usa identificador PAX: la cantidad se calcula automaticamente con el numero de personas del evento.";
    }
    return;
  }
  el.quoteServiceQty.disabled = false;
  el.quoteServiceQty.readOnly = false;
  el.quoteServiceQty.placeholder = "Ingresa cantidad";
  if (el.quoteServiceQtyHint) {
    el.quoteServiceQtyHint.textContent = matched
      ? "Este servicio es manual: debes indicar una cantidad valida antes de agregarlo al carrito."
      : "Si no existe identificador PAX, debes ingresar una cantidad valida antes de agregar el servicio.";
  }
}

function renderQuoteSidebarSummary() {
  if (!quoteDraft) return;
  const currency = normalizeQuoteCurrency(quoteDraft.currency);
  const money = (n) => moneyQuote(n, currency);
  const totals = getQuoteTotals(quoteDraft);
  const { catBuckets } = aggregateQuoteBuckets(quoteDraft || {});
  const breakdownRows = [
    { label: "Alimentos y bebidas", amount: Number(catBuckets?.alimentosBebidas?.amount || 0) },
    { label: "Miscelaneos", amount: Number(catBuckets?.miscelaneos?.amount || 0) },
    { label: "Hospedaje JDL", amount: Number(catBuckets?.hospedajeJdl?.amount || 0) },
    { label: "Hospedaje terceros", amount: Number(catBuckets?.hospedajeTerceros?.amount || 0) },
  ].filter((row) => Number(row.amount || 0) > 0);
  const company = (state.companies || []).find((c) => String(c?.id || "") === String(quoteDraft.companyId || ""));
  const documentLabel = String(quoteDraft.code || "COTIZACION SIN CODIGO").trim() || "COTIZACION SIN CODIGO";
  const clientLabel = String(
    quoteDraft.companyName
    || company?.name
    || quoteDraft.contact
    || "Selecciona una institucion para comenzar"
  ).trim() || "Selecciona una institucion para comenzar";

  if (el.quoteSidebarDocument) el.quoteSidebarDocument.textContent = documentLabel;
  if (el.quoteSidebarClient) el.quoteSidebarClient.textContent = clientLabel;
  if (el.quoteSidebarSubtotal) el.quoteSidebarSubtotal.textContent = money(totals.subtotal);
  if (el.quoteSidebarDiscount) el.quoteSidebarDiscount.textContent = money(totals.discountAmount);
  if (el.quoteSidebarTotal) el.quoteSidebarTotal.textContent = money(totals.total);
  if (el.quoteSidebarBreakdown) {
    if (!breakdownRows.length) {
      el.quoteSidebarBreakdown.innerHTML = `
        <div class="quoteSidebarBreakdownEmpty">
          Agrega servicios para ver aqui la distribucion por categoria.
        </div>
      `;
    } else {
      el.quoteSidebarBreakdown.innerHTML = breakdownRows.map((row) => `
        <div class="quoteSidebarBreakdownRow">
          <span>${escapeHtml(row.label)}</span>
          <strong>${escapeHtml(money(row.amount))}</strong>
        </div>
      `).join("");
    }
  }
}

function syncQuoteItemsSelectionUi() {
  const ids = quoteDraft?.items?.map((item) => String(item.rowId || "")) || [];
  quoteSelectedItemIds = new Set(Array.from(quoteSelectedItemIds).filter((id) => ids.includes(id)));
  const selectedCount = quoteSelectedItemIds.size;
  if (el.btnQuoteDuplicateSelected) el.btnQuoteDuplicateSelected.disabled = selectedCount === 0;
  if (el.btnQuoteMoveSelectedUp) el.btnQuoteMoveSelectedUp.disabled = selectedCount === 0;
  if (el.btnQuoteMoveSelectedDown) el.btnQuoteMoveSelectedDown.disabled = selectedCount === 0;
  if (el.quoteSelectAllItems) {
    const total = ids.length;
    const all = total > 0 && selectedCount === total;
    const some = selectedCount > 0 && selectedCount < total;
    el.quoteSelectAllItems.checked = all;
    el.quoteSelectAllItems.indeterminate = some;
  }
}

function duplicateSelectedQuoteItems() {
  if (!quoteDraft || !quoteDraft.items?.length || !quoteSelectedItemIds.size) return;
  const next = [];
  for (const item of quoteDraft.items) {
    next.push(item);
    if (!quoteSelectedItemIds.has(String(item.rowId || ""))) continue;
    const cloned = {
      ...deepClone(item),
      rowId: uid(),
    };
    next.push(cloned);
    quoteSelectedItemIds.add(String(cloned.rowId || ""));
  }
  quoteDraft.items = next;
  renderQuoteItems();
  toast("Servicios duplicados.");
}

function moveSelectedQuoteItems(direction = "up") {
  if (!quoteDraft || !quoteDraft.items?.length || !quoteSelectedItemIds.size) return;
  const arr = quoteDraft.items;
  if (direction === "up") {
    for (let i = 1; i < arr.length; i++) {
      const currentSelected = quoteSelectedItemIds.has(String(arr[i].rowId || ""));
      const prevSelected = quoteSelectedItemIds.has(String(arr[i - 1].rowId || ""));
      if (currentSelected && !prevSelected) {
        [arr[i - 1], arr[i]] = [arr[i], arr[i - 1]];
      }
    }
  } else {
    for (let i = arr.length - 2; i >= 0; i--) {
      const currentSelected = quoteSelectedItemIds.has(String(arr[i].rowId || ""));
      const nextSelected = quoteSelectedItemIds.has(String(arr[i + 1].rowId || ""));
      if (currentSelected && !nextSelected) {
        [arr[i], arr[i + 1]] = [arr[i + 1], arr[i]];
      }
    }
  }
  renderQuoteItems();
}

// Redibuja la tabla de servicios de la cotizacion y recalcula la vista.
// Cada vez que cambia cantidad, precio, descuento o moneda,
// esta funcion refleja el cambio visual y prepara el resumen final.
function renderQuoteItems() {
  if (!quoteDraft || !el.quoteItemsBody) return;
  quoteDraft.currency = normalizeQuoteCurrency(quoteDraft.currency);
  renderQuoteCurrencyPicker();
  const money = (n) => moneyQuote(n, quoteDraft.currency);
  quoteDraft.discountType = normalizeDiscountType(quoteDraft.discountType);
  quoteDraft.discountValue = Math.max(0, Number(quoteDraft.discountValue || 0));
  if (el.quoteDiscountType && el.quoteDiscountType.value !== quoteDraft.discountType) {
    el.quoteDiscountType.value = quoteDraft.discountType;
  }
  if (el.quoteDiscountValue) {
    const nextDiscountValue = String(quoteDraft.discountValue);
    if (el.quoteDiscountValue.value !== nextDiscountValue) el.quoteDiscountValue.value = nextDiscountValue;
  }
  el.quoteItemsBody.innerHTML = "";
  const rangeDates = getQuoteRangeDates();
  const defaultDate = rangeDates[0] || "";
  const availableDates = rangeDates.length ? rangeDates : [defaultDate].filter(Boolean);
  const byDate = new Map();

  for (const item of quoteDraft.items) {
    if (!item.serviceDate || (availableDates.length && !availableDates.includes(item.serviceDate))) {
      item.serviceDate = defaultDate;
    }
    const key = item.serviceDate || defaultDate || "sin_fecha";
    if (!byDate.has(key)) byDate.set(key, []);
    byDate.get(key).push(item);
  }

  const orderedDates = availableDates.length
    ? availableDates.filter(d => byDate.has(d))
    : Array.from(byDate.keys()).sort((a, b) => a.localeCompare(b));

  if (!quoteDraft.items.length) {
    const tr = document.createElement("tr");
    tr.className = "quoteItemsEmptyRow";
    tr.innerHTML = `
      <td colspan="7">
        <div class="quoteItemsEmptyState">
          <strong>Tu carrito aun esta vacio</strong>
          <span>Busca un servicio en el panel derecho o crea uno nuevo para comenzar a cotizar.</span>
        </div>
      </td>
    `;
    el.quoteItemsBody.appendChild(tr);
  }

  for (const dateKey of orderedDates) {
    const dayItems = byDate.get(dateKey) || [];
    for (const item of dayItems) {
      const tr = document.createElement("tr");
      tr.dataset.rowId = item.rowId;
      const isSelected = quoteSelectedItemIds.has(String(item.rowId || ""));
      if (isSelected) tr.classList.add("quoteItemSelected");
      const subtotal = Number(item.qty || 0) * Number(item.price || 0);
      const isPaxMode = String(item.quantityMode || "").toUpperCase() === "PAX";
      const qtyValue = Number(item.qty || 0);
      const priceValue = Number(item.price || 0);
      const qtyInputHtml = `<input class="quoteInput quoteSimpleInput${isPaxMode ? " quotePaxReadonly" : ""}" data-field="qty" type="number" min="0" step="1" value="${qtyValue}" ${isPaxMode ? "readonly disabled" : ""} title="${isPaxMode ? "Cantidad automatica por pax" : ""}" />`;
      const priceInputHtml = `<input class="quoteInput quoteSimpleInput${isPaxMode ? " quotePaxReadonly" : ""}" data-field="price" type="number" min="0" step="1" value="${priceValue}" ${isPaxMode ? "readonly disabled" : ""} title="${isPaxMode ? "Precio calculado por pax" : ""}" />`;
      const currentName = String(item.name || "").trim();
      const currentDetail = String(item.description || "").trim();
      const customDetail = currentDetail && currentDetail !== currentName ? currentDetail : "";
      tr.innerHTML = `
        <td class="quoteItemSelectCell">
          <input type="checkbox" class="quoteItemSelect" data-action="toggle-select" ${isSelected ? "checked" : ""} />
        </td>
        <td class="quoteItemDateCell">
          <select class="quoteInput" data-field="serviceDate">
            ${availableDates.map(d => `<option value="${escapeHtml(d)}"${d === item.serviceDate ? " selected" : ""}>${escapeHtml(d)}</option>`).join("")}
          </select>
        </td>
        <td class="quoteItemQtyCell">
          ${qtyInputHtml}
        </td>
        <td class="quoteItemDescriptionCell">
          <div class="quoteItemServiceStack">
            <div class="quoteItemServiceRow">
              <input class="quoteInput quoteItemNameInput" data-field="name" list="servicesList" value="${escapeHtml(item.name || "")}" placeholder="Buscar servicio..." />
              <button type="button" class="quoteDetailBtn" data-action="edit-detail" title="Editar detalle del servicio">Detalle</button>
            </div>
            ${customDetail ? `<div class="quoteItemDetailPreview">${escapeHtml(customDetail)}</div>` : ""}
          </div>
        </td>
        <td class="quoteItemPriceCell">${priceInputHtml}</td>
        <td class="quoteMoney quoteItemTotalCell">
          <strong>${money(subtotal)}</strong>
          <span>${qtyValue} x ${money(priceValue)}</span>
        </td>
        <td class="quoteItemActionCell">
          <button type="button" class="apptIconBtn quoteMoveUpBtn" data-action="move-up" title="Subir servicio" aria-label="Subir servicio"><span class="material-symbols-outlined" aria-hidden="true">keyboard_arrow_up</span></button>
          <button type="button" class="apptIconBtn quoteMoveDownBtn" data-action="move-down" title="Bajar servicio" aria-label="Bajar servicio"><span class="material-symbols-outlined" aria-hidden="true">keyboard_arrow_down</span></button>
          <button type="button" class="apptIconBtn apptDelete quoteRemoveBtn" title="Eliminar servicio" aria-label="Eliminar servicio"><span class="material-symbols-outlined" aria-hidden="true">delete</span></button>
        </td>
      `;
      el.quoteItemsBody.appendChild(tr);
    }

    const subtotalDay = dayItems.reduce((acc, x) => acc + Number(x.qty || 0) * Number(x.price || 0), 0);
    const subtotalRow = document.createElement("tr");
    subtotalRow.className = "quoteSubtotalRow";
    subtotalRow.innerHTML = `
      <td colspan="5">Subtotal ${escapeHtml(dateKey)}</td>
      <td class="quoteMoney">${money(subtotalDay)}</td>
      <td></td>
    `;
    el.quoteItemsBody.appendChild(subtotalRow);
  }

  const totals = getQuoteTotals(quoteDraft);
  if (el.quoteSubtotal) el.quoteSubtotal.textContent = money(totals.subtotal);
  if (el.quoteDiscountAmount) el.quoteDiscountAmount.textContent = money(totals.discountAmount);
  el.quoteTotal.textContent = money(totals.total);
  renderQuoteSidebarSummary();
  renderQuoteAccountStatement();
  syncQuoteServiceDateRequired();
  syncQuoteItemsSelectionUi();
}

// Procesa cambios directos en los renglones de servicios.
// Actualiza el borrador en memoria y vuelve a recalcular totales
// para que el usuario vea de inmediato el efecto de cada ajuste.
function handleQuoteItemsInput(e) {
  const input = e.target.closest("[data-field]");
  if (!input || !quoteDraft) return;
  const row = input.closest("tr");
  const rowId = row?.dataset.rowId;
  if (!rowId) return;
  const item = quoteDraft.items.find(x => x.rowId === rowId);
  if (!item) return;

  const field = input.dataset.field;
  if (field === "qty") {
    if (String(item.quantityMode || "").toUpperCase() === "PAX") {
      item.qty = 1;
      item.price = Math.max(0, Number(item.unitPrice || item.price || 0) * normalizeQuotePeopleValue());
    } else {
      item.qty = Math.max(0, Number(input.value || 0));
    }
  } else if (field === "price") {
    if (String(item.quantityMode || "").toUpperCase() === "PAX") {
      const pax = normalizeQuotePeopleValue();
      const effective = Math.max(0, Number(input.value || 0));
      item.unitPrice = pax > 0 ? (effective / pax) : effective;
      item.price = Math.max(0, Number(item.unitPrice || 0) * pax);
      item.qty = 1;
    } else {
      item.price = Math.max(0, Number(input.value || 0));
    }
  } else if (field === "name") {
    const previousName = String(item.name || "").trim();
    const previousDescription = String(item.description || "").trim();
    const term = String(input.value || "").trim();
    item.name = term;
    if (!previousDescription || previousDescription === previousName) item.description = term;
    const matched = resolveServiceFromSearch(term);
    if (matched) {
      applyServiceToQuoteItem(item, matched);
    } else {
      item.serviceId = null;
      item.quantityMode = "MANUAL";
    }
  } else if (field === "description") {
    const term = String(input.value || "").trim();
    item.description = term || String(item.name || "").trim();
  } else if (field === "serviceDate") {
    item.serviceDate = input.value;
  }
  renderQuoteItems();
}

async function handleQuoteItemsClick(e) {
  const selectToggle = e.target.closest(".quoteItemSelect");
  if (selectToggle && quoteDraft) {
    const row = selectToggle.closest("tr");
    const rowId = String(row?.dataset.rowId || "").trim();
    if (!rowId) return;
    if (selectToggle.checked) quoteSelectedItemIds.add(rowId);
    else quoteSelectedItemIds.delete(rowId);
    syncQuoteItemsSelectionUi();
    row?.classList.toggle("quoteItemSelected", selectToggle.checked);
    return;
  }
  const moveBtn = e.target.closest("[data-action='move-up'], [data-action='move-down']");
  if (moveBtn && quoteDraft) {
    const row = moveBtn.closest("tr");
    const rowId = String(row?.dataset.rowId || "").trim();
    if (!rowId) return;
    const index = quoteDraft.items.findIndex((x) => String(x.rowId || "") === rowId);
    if (index < 0) return;
    if (moveBtn.dataset.action === "move-up" && index > 0) {
      [quoteDraft.items[index - 1], quoteDraft.items[index]] = [quoteDraft.items[index], quoteDraft.items[index - 1]];
      renderQuoteItems();
    } else if (moveBtn.dataset.action === "move-down" && index < quoteDraft.items.length - 1) {
      [quoteDraft.items[index], quoteDraft.items[index + 1]] = [quoteDraft.items[index + 1], quoteDraft.items[index]];
      renderQuoteItems();
    }
    return;
  }
  const detailBtn = e.target.closest(".quoteDetailBtn");
  if (detailBtn && quoteDraft) {
    const row = detailBtn.closest("tr");
    const rowId = row?.dataset.rowId;
    if (!rowId) return;
    const item = quoteDraft.items.find(x => x.rowId === rowId);
    if (!item) return;
    const serviceName = String(item.name || "").trim();
    const currentDetail = String(item.description || "").trim();
    const editableDetail = currentDetail && currentDetail !== serviceName ? currentDetail : "";
    const detail = window.prompt("Detalle adicional del servicio (deja vacio para quitarlo):", editableDetail);
    if (detail === null) return;
    const cleaned = String(detail || "").trim();
    item.description = cleaned || serviceName;
    renderQuoteItems();
    toast(cleaned ? "Detalle actualizado." : "Detalle eliminado.");
    return;
  }
  const btn = e.target.closest(".quoteRemoveBtn");
  if (!btn || !quoteDraft) return;
  const row = btn.closest("tr");
  const rowId = row?.dataset.rowId;
  if (!rowId) return;
  const ok = await modernConfirm({
    title: "Eliminar servicio",
    message: "Esta seguro de eliminar este servicio de la cotizacion?",
    confirmText: "Si, eliminar",
    cancelText: "No",
  });
  if (!ok) return;
  const removed = quoteDraft.items.find(x => x.rowId === rowId);
  quoteDraft.items = quoteDraft.items.filter(x => x.rowId !== rowId);
  quoteSelectedItemIds.delete(String(rowId));
  renderQuoteItems();
  const ctx = getCurrentQuoteHistoryContext();
  if (ctx) {
    appendHistoryByKey(
      ctx.key,
      authSession.userId || ctx.ev.userId || "",
      `Servicio eliminado de cotizacion: ${String(removed?.name || "").trim() || "(sin nombre)"}`
    );
  }
  toast("Servicio eliminado de la cotizacion.");
}

function focusQuoteValidationTarget(node, message, opts = {}) {
  const openDoc = opts?.openDoc !== false;
  const expandItems = opts?.expandItems === true;
  if (openDoc) setQuoteDocPanelVisible(true);
  if (expandItems) setQuoteItemsExpanded(true);
  toast(String(message || "Completa los datos de la cotizacion."));
  if (!node) return false;
  setTimeout(() => {
    try {
      node.scrollIntoView({ behavior: "smooth", block: "center" });
      node.focus();
      if (typeof node.select === "function" && /^(INPUT|TEXTAREA)$/i.test(String(node.tagName || ""))) {
        node.select();
      }
    } catch (_) { }
  }, 40);
  return false;
}

function setQuoteSavingProgress(percent, label = "") {
  const value = Math.max(0, Math.min(100, Number(percent || 0)));
  if (el.quoteSaveProgressFill) {
    el.quoteSaveProgressFill.style.width = `${value}%`;
  }
  if (el.quoteSaveProgressBar) {
    el.quoteSaveProgressBar.setAttribute("aria-valuenow", String(Math.round(value)));
  }
  if (el.quoteSaveProgressText && label) {
    el.quoteSaveProgressText.textContent = String(label);
  }
}

function setQuoteSavingState(isSaving, label = "") {
  const busy = !!isSaving;
  quoteSaveInFlight = busy;
  if (el.quoteSaveProgress) {
    el.quoteSaveProgress.hidden = !busy;
  }
  if (el.btnQuoteSave) {
    el.btnQuoteSave.disabled = busy;
    el.btnQuoteSave.textContent = busy ? "Guardando..." : "Guardar cotizacion";
  }
  if (busy) {
    setQuoteSavingProgress(5, label || "Iniciando guardado...");
  } else {
    setQuoteSavingProgress(0, "Guardando cotizacion...");
  }
}

function handleQuoteDocSaveClose() {
  if (quoteDraft) {
    quoteDraft.docDate = String(el.quoteDocDate?.value || quoteDraft.docDate || "").trim();
    quoteDraft.dueDate = String(el.quoteDueDate?.value || quoteDraft.dueDate || "").trim();
  }
  setQuoteDocPanelVisible(false);
  toast("Datos de cotizacion guardados.");
}
window.handleQuoteDocSaveClose = handleQuoteDocSaveClose;
// Guarda la cotizacion tomando el estado actual del formulario.
// Esta rutina valida los campos obligatorios, normaliza montos y pagos,
// y decide si debe crear o actualizar versiones dentro del evento actual.
async function saveQuoteFromForm() {
  if (quoteSaveInFlight) return;
  if (!quoteDraft) return;
  setQuoteSavingState(true, "Validando datos...");
  try {
  const eventId = el.quoteEventId.value;
  const ev = state.events.find(x => x.id === eventId);
  if (!ev) return;
  const reservationKey = reservationKeyFromEvent(ev);
  const quoteMeta = getQuoteEventMeta(eventId);
  syncQuoteReservationDates(quoteMeta);

  quoteDraft.companyId = el.quoteCompany.value;
  quoteDraft.managerId = el.quoteManagerSelect.value;
  quoteDraft.contact = el.quoteContact.value.trim();
  quoteDraft.email = el.quoteEmail.value.trim();
  quoteDraft.billTo = el.quoteBillTo.value.trim();
  quoteDraft.address = el.quoteAddress.value.trim();
  quoteDraft.eventType = el.quoteEventType.value.trim();
  quoteDraft.venue = el.quoteVenue.value.trim();
  quoteDraft.schedule = el.quoteSchedule.value.trim();
  quoteDraft.code = String(quoteDraft.code || el.quoteCode.value || "").trim();
  quoteDraft.docDate = el.quoteDocDate.value;
  quoteDraft.phone = el.quotePhone.value.trim();
  quoteDraft.nit = el.quoteNIT.value.trim();
  quoteDraft.people = el.quotePeople.value;
  syncPaxQuantityItems();
  quoteDraft.eventDate = String(quoteMeta?.startDate || el.quoteEventDate.value || "").trim();
  quoteDraft.folio = el.quoteFolio.value.trim();
  quoteDraft.endDate = String(quoteMeta?.endDate || el.quoteEndDate.value || quoteDraft.eventDate || "").trim();
  quoteDraft.dueDate = el.quoteDueDate.value;
  quoteDraft.paymentType = quotePaymentTypesToStorage(getSelectedQuotePaymentTypesFromForm());
  quoteDraft.discountType = normalizeDiscountType(el.quoteDiscountType?.value || quoteDraft.discountType);
  quoteDraft.discountValue = Math.max(0, Number(el.quoteDiscountValue?.value || quoteDraft.discountValue || 0));
  quoteDraft.templateId = String(el.quoteTemplateSelect?.value || quoteDraft.templateId || "").trim();
  quoteDraft.currency = normalizeQuoteCurrency(quoteDraft.currency);
  quoteDraft.internalNotes = el.quoteInternalNotes.value.trim();
  quoteDraft.notes = quoteDraft.internalNotes;
  setQuoteSavingProgress(22, "Revisando validaciones...");

  if (!quoteDraft.companyId) return focusQuoteValidationTarget(el.quoteCompanySearch, "Selecciona la institucion en Datos de cotizacion.");
  if (!quoteDraft.managerId) return focusQuoteValidationTarget(el.quoteManagerSelect, "Selecciona Encargado Evento en Datos de cotizacion.");
  if (!quoteDraft.contact) return focusQuoteValidationTarget(el.quoteContact, "Completa Contacto en Datos de cotizacion.");
  if (!quoteDraft.email) return focusQuoteValidationTarget(el.quoteEmail, "Completa Email en Datos de cotizacion.");
  if (!isValidEmail(quoteDraft.email)) return focusQuoteValidationTarget(el.quoteEmail, "El Email de cotizacion no es valido.");
  if (!quoteDraft.billTo) return focusQuoteValidationTarget(el.quoteBillTo, "Completa Facturar A en Datos de cotizacion.");
  if (!quoteDraft.address) return focusQuoteValidationTarget(el.quoteAddress, "Completa Direccion en Datos de cotizacion.");
  if (!quoteDraft.eventType) return focusQuoteValidationTarget(el.quoteEventType, "Completa Tipo Evento en Datos de cotizacion.");
  if (!quoteDraft.venue) return focusQuoteValidationTarget(el.quoteVenue, "Completa Salon o Jardin en Datos de cotizacion.");
  if (!quoteDraft.schedule) return focusQuoteValidationTarget(el.quoteSchedule, "Completa Horario y Evento en Datos de cotizacion.");
  if (!quoteDraft.code) return focusQuoteValidationTarget(el.quoteCode, "Completa Codigo en Datos de cotizacion.");
  if (!quoteDraft.docDate) return focusQuoteValidationTarget(el.quoteDocDate, "Completa Fecha Documento en Datos de cotizacion.");
  if (!quoteDraft.phone) return focusQuoteValidationTarget(el.quotePhone, "Completa Telefono en Datos de cotizacion.");
  if (!quoteDraft.nit) return focusQuoteValidationTarget(el.quoteNIT, "Completa NIT en Datos de cotizacion.");
  if (!quoteDraft.paymentType) return focusQuoteValidationTarget(el.quotePaymentTypeSelect || el.quotePaymentType, "Agrega al menos una Forma de Pago en Datos de cotizacion.");
  if (!quoteDraft.people || Number(quoteDraft.people) <= 0) return focusQuoteValidationTarget(el.quotePeople, "Ingresa un numero valido en No Personas.");
  if (!quoteDraft.eventDate) return focusQuoteValidationTarget(el.quoteEventDate, "Completa Fecha evento en Datos de cotizacion.");
  if (!quoteDraft.endDate) return focusQuoteValidationTarget(el.quoteEndDate, "Completa Fecha Finalizacion en Datos de cotizacion.");
  if (quoteMeta?.startDate && quoteDraft.eventDate !== quoteMeta.startDate) return focusQuoteValidationTarget(el.quoteEventDate, "La Fecha evento debe coincidir con el inicio de la reservacion.");
  if (quoteMeta?.endDate && quoteDraft.endDate !== quoteMeta.endDate) return focusQuoteValidationTarget(el.quoteEndDate, "La Fecha Finalizacion debe coincidir con el fin de la reservacion.");
  if (!quoteDraft.dueDate) return focusQuoteValidationTarget(el.quoteDueDate, "Completa Fecha Maxima Pago en Datos de cotizacion.");
  if (!quoteDraft.items.length) return focusQuoteValidationTarget(el.quoteServiceSearch, "Agrega al menos un servicio en el bloque Agregar servicio.", { openDoc: false, expandItems: true });

  const company = (state.companies || []).find(c => c.id === quoteDraft.companyId);
  const manager = company?.managers?.find(m => m.id === quoteDraft.managerId);
  if (!company || !manager) return toast("Empresa/encargado invalido.");
  quoteDraft.companyName = company.name;
  quoteDraft.managerName = manager.name;
  quoteDraft.managerPhone = manager.phone || "";
  quoteDraft.quotedAt = new Date().toISOString();
  const previousQuote = ev.quote ? deepClone(ev.quote) : null;
  const historyVersions = normalizeQuoteVersionHistory(previousQuote?.versions);
  setQuoteSavingProgress(42, "Preparando version...");
  const unchangedQuote = previousQuote ? areQuotesEquivalentForVersioning(quoteDraft, previousQuote) : false;
  if (previousQuote && !unchangedQuote) {
    const prevVersion = Number(previousQuote.version || (historyVersions.length + 1));
    const prevSnapshot = cloneQuoteVersionSnapshot(previousQuote, prevVersion);
    if (!historyVersions.some((v) => Number(v.version) === Number(prevSnapshot.version))) {
      historyVersions.push(prevSnapshot);
    }
  }
  const nextVersion = previousQuote
    ? (unchangedQuote
      ? Math.max(1, Number(previousQuote.version || 1))
      : (Math.max(0, Number(previousQuote.version || (historyVersions.length))) + 1))
    : 1;
  const savedQuote = cloneQuoteSnapshot(quoteDraft, nextVersion);
  savedQuote.companyName = quoteDraft.companyName;
  savedQuote.managerName = quoteDraft.managerName;
  savedQuote.managerPhone = quoteDraft.managerPhone || "";
  savedQuote.quotedAt = quoteDraft.quotedAt;
  savedQuote.version = nextVersion;
  savedQuote.versions = historyVersions.sort((a, b) => Number(a.version) - Number(b.version));
  const totals = getQuoteTotals(savedQuote);
  savedQuote.subtotal = totals.subtotal;
  savedQuote.discountAmount = totals.discountAmount;
  savedQuote.total = totals.total;

  const rangeDates = getQuoteRangeDates();
  const validItems = savedQuote.items.every(x => {
    const nameOk = String(x.name || "").trim().length > 0;
    const qtyOk = Number(x.qty) > 0;
    const priceOk = Number(x.price) >= 0;
    const dateOk = !rangeDates.length || rangeDates.includes(String(x.serviceDate || ""));
    return nameOk && qtyOk && priceOk && dateOk;
  });
  if (!validItems) return toast("Completa descripcion, cantidad, precio y fecha valida en cada servicio.");
  setQuoteSavingProgress(58, "Aplicando cambios...");

  const series = getEventSeries(ev);
  const isFirstQuote = !previousQuote;
  let movedToSeguimiento = false;
  let movedToFirstQuote = false;
  for (const item of series) {
    item.quote = deepClone(savedQuote);
    if (isFirstQuote && item.status === STATUS.RESERVA_SIN_COTIZACION) {
      item.status = STATUS.PRIMERA;
      movedToFirstQuote = true;
    } else if (
      !isFirstQuote
      && !unchangedQuote
      && (item.status === STATUS.PRIMERA || item.status === STATUS.RESERVA_SIN_COTIZACION)
    ) {
      item.status = STATUS.SEGUIMIENTO;
      movedToSeguimiento = true;
    }
  }
  const totalQuote = totals.total;
  appendHistoryByKey(
    reservationKey,
    ev.userId,
    unchangedQuote
      ? `Cotizacion verificada sin cambios (V${savedQuote.version}). Total ${moneyGT(totalQuote)}.`
      : (movedToFirstQuote
        ? `Primera cotizacion guardada. Total ${moneyGT(totalQuote)}. Estado a 1er Cotizacion.`
        : (movedToSeguimiento
          ? `Cotizacion actualizada. Total ${moneyGT(totalQuote)}. Estado a Seguimiento.`
          : `Cotizacion guardada. Total ${moneyGT(totalQuote)}. Estado conservado.`))
  );

  setQuoteSavingProgress(72, "Guardando en sistema...");
  persist();
  setQuoteSavingProgress(82, "Generando documento...");
  await openQuoteDocument(ev, savedQuote);
  setQuoteSavingProgress(95, "Finalizando...");
  render();
  closeQuoteModal();
  openModalForEdit(eventId);
  toast(unchangedQuote
    ? `Sin cambios detectados. Se mantiene la version V${savedQuote.version}.`
    : `Cotizacion guardada. Version V${savedQuote.version}.`);
  setQuoteSavingProgress(100, "Guardado completado.");
  } finally {
    setTimeout(() => setQuoteSavingState(false), 450);
  }
}

function dataUrlToUint8Array(dataUrl) {
  const raw = String(dataUrl || "");
  const idx = raw.indexOf(",");
  const b64 = idx >= 0 ? raw.slice(idx + 1) : raw;
  if (!b64) return new Uint8Array(0);
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

function mapTemplateFontToPdfLib(StandardFonts, field) {
  const family = String(field?.fontFamily || "Arial").toLowerCase();
  const bold = field?.bold === true;
  const italic = field?.italic === true;
  const useTimes = family.includes("times") || family.includes("georgia");
  const useCourier = family.includes("courier");
  if (useCourier) {
    if (bold && italic) return StandardFonts.CourierBoldOblique;
    if (bold) return StandardFonts.CourierBold;
    if (italic) return StandardFonts.CourierOblique;
    return StandardFonts.Courier;
  }
  if (useTimes) {
    if (bold && italic) return StandardFonts.TimesBoldItalic;
    if (bold) return StandardFonts.TimesBold;
    if (italic) return StandardFonts.TimesItalic;
    return StandardFonts.TimesRoman;
  }
  if (bold && italic) return StandardFonts.HelveticaBoldOblique;
  if (bold) return StandardFonts.HelveticaBold;
  if (italic) return StandardFonts.HelveticaOblique;
  return StandardFonts.Helvetica;
}

function isImageDataUrl(value) {
  return /^data:image\/(png|jpe?g|webp);base64,/i.test(String(value || "").trim());
}

function detectImageKindFromBytes(bytes) {
  if (!(bytes instanceof Uint8Array) || bytes.length < 4) return "";
  const b0 = bytes[0];
  const b1 = bytes[1];
  const b2 = bytes[2];
  const b3 = bytes[3];
  // PNG: 89 50 4E 47
  if (b0 === 0x89 && b1 === 0x50 && b2 === 0x4E && b3 === 0x47) return "png";
  // JPEG: FF D8 FF
  if (b0 === 0xFF && b1 === 0xD8 && b2 === 0xFF) return "jpg";
  return "";
}

async function embedTemplateImageFromDataUrl(doc, dataUrl) {
  const raw = String(dataUrl || "").trim();
  if (!raw) return null;
  const bytes = dataUrlToUint8Array(raw);
  if (!bytes.length) return null;
  const hintedPng = /^data:image\/png/i.test(raw);
  const hintedJpg = /^data:image\/jpe?g/i.test(raw);
  const kind = detectImageKindFromBytes(bytes);

  // Prefer actual bytes signature. Fallback to MIME hint, then try both.
  if (kind === "png") return doc.embedPng(bytes);
  if (kind === "jpg") return doc.embedJpg(bytes);
  if (hintedPng) {
    try { return await doc.embedPng(bytes); } catch (_) { }
    try { return await doc.embedJpg(bytes); } catch (_) { }
    return null;
  }
  if (hintedJpg) {
    try { return await doc.embedJpg(bytes); } catch (_) { }
    try { return await doc.embedPng(bytes); } catch (_) { }
    return null;
  }
  try { return await doc.embedPng(bytes); } catch (_) { }
  try { return await doc.embedJpg(bytes); } catch (_) { }
  return null;
}

function formatTemplateRoomRates(roomRates = []) {
  const rows = Array.isArray(roomRates) ? roomRates : [];
  if (!rows.length) return "";
  return rows
    .map((r) => `${String(r?.habitacion || "").trim()}: ${moneyGT(r?.precio || 0)}`)
    .filter(Boolean)
    .join("\n");
}

function evaluateTemplateFormula(expression, ctx) {
  const code = String(expression || "").trim();
  if (!code) return "";
  try {
    const fn = new Function("ctx", `with(ctx){ return (${code}); }`); // user-authored formulas
    const out = fn(ctx);
    if (out == null) return "";
    if (typeof out === "number" && Number.isFinite(out)) return String(out);
    return String(out);
  } catch (_) {
    return "";
  }
}

function buildQuoteSummaryText(quote) {
  const items = Array.isArray(quote?.items) ? quote.items : [];
  if (!items.length) return "Sin detalle comercial.";
  const lines = [];
  for (const item of items) {
    const name = String(item?.name || "").trim();
    if (!name) continue;
    const qty = Math.max(0, Number(item?.qty || 0));
    const price = Math.max(0, Number(item?.price || 0));
    const total = qty * price;
    lines.push(`- ${name}${qty > 0 ? ` (${qty})` : ""} | ${moneyGT(total)}`);
  }
  return lines.length ? lines.join("\n") : "Sin detalle comercial.";
}

function buildQuoteTermsText(quote, template) {
  const notes = String(quote?.internalNotes || quote?.notes || "").trim();
  if (notes) return notes;
  const tplBody = String(template?.body || "").trim();
  if (tplBody) return tplBody;
  return "Condiciones sujetas a politicas internas y disponibilidad.";
}

function buildTemplateTokenContext(ev, quote, company, manager, template) {
  const totals = getQuoteTotals(quote);
  const sellerUser = (state.users || []).map(normalizeUserRecord).find((u) => String(u.id) === String(ev?.userId || ""));
  const authUser = normalizeUserRecord(getAuthUserRecord() || {});
  const authSessionSignature = String(authSession.signatureDataUrl || "").trim();
  const authSignature = String(authUser?.signatureDataUrl || "").trim();
  const sellerSignature = String(sellerUser?.signatureDataUrl || "").trim();
  const vendorSignature = authSessionSignature || authSignature || sellerSignature || "________________";
  const vendorName = String(authSession.fullName || authUser?.fullName || authUser?.name || sellerUser?.fullName || sellerUser?.name || quote?.managerName || "").trim();
  const summaryText = buildQuoteSummaryText(quote);
  const termsText = buildQuoteTermsText(quote, template);
  const ctx = {
    cliente: {
      nombre: String(quote?.contact || "").trim(),
      firma: "________________",
      telefono: String(quote?.phone || "").trim(),
      direccion: String(quote?.address || "").trim(),
    },
    vendedor: {
      nombre: vendorName,
      firma: vendorSignature,
      telefono: String(manager?.phone || quote?.phone || "").trim(),
      direccion: String(company?.address || "").trim(),
    },
    empresa: {
      nombre: String(company?.name || quote?.companyName || "").trim(),
      nit: String(company?.nit || quote?.nit || "").trim(),
      direccion: String(company?.address || quote?.address || "").trim(),
    },
    institucion: {
      nombre: String(company?.name || quote?.companyName || "").trim(),
    },
    fecha: {
      hoy: toISODate(new Date()),
      evento: String(quote?.eventDate || ev?.date || "").trim(),
    },
    evento: {
      nombre: String(ev?.name || "").trim(),
      pax: String(quote?.people || "").trim(),
    },
    quote: {
      summary: summaryText,
      terms: termsText,
      notes: String(quote?.internalNotes || quote?.notes || "").trim(),
    },
    monto: {
      subtotal: Number(totals.subtotal || 0),
      descuento: Number(totals.discountAmount || 0),
      total: Number(totals.total || 0),
    },
    tabla: {
      habitaciones: formatTemplateRoomRates(template?.roomRates || []),
    },
  };
  return ctx;
}

function resolveTemplateTokenValue(token, ctx, template) {
  const t = String(token || "").trim();
  if (!t) return "";
  const mFormula = t.match(/^\{\{\s*formula:([a-zA-Z0-9_.-]+)\s*\}\}$/);
  if (mFormula) {
    const key = mFormula[1];
    const formulas = Array.isArray(template?.formulas) ? template.formulas : [];
    const f = formulas.find((x) => String(x?.key || "").trim() === key);
    if (!f) return "";
    return evaluateTemplateFormula(f.expression, ctx);
  }
  const mToken = t.match(/^\{\{\s*([a-zA-Z0-9_.-]+)\s*\}\}$/);
  if (!mToken) return t;
  const path = mToken[1].split(".");
  let cur = ctx;
  for (const k of path) {
    if (!cur || typeof cur !== "object" || !(k in cur)) return "";
    cur = cur[k];
  }
  if (cur == null) return "";
  if (typeof cur === "number" && Number.isFinite(cur)) return String(cur);
  return String(cur);
}

function resolveTemplateFieldByHint(hint, ctx) {
  const h = String(hint || "").toLowerCase();
  if (!h) return "";
  if (h.includes("resumen") && h.includes("comercial")) return String(ctx?.quote?.summary || "");
  if (h.includes("condicion")) return String(ctx?.quote?.terms || "");
  if (h.includes("nombre") && h.includes("cliente")) return String(ctx?.cliente?.nombre || "");
  if (h.includes("telefono") && h.includes("cliente")) return String(ctx?.cliente?.telefono || "");
  if (h.includes("nombre") && h.includes("vendedor")) return String(ctx?.vendedor?.nombre || "");
  if (h.includes("telefono") && h.includes("vendedor")) return String(ctx?.vendedor?.telefono || "");
  if (h.includes("firma") && h.includes("vendedor")) return String(ctx?.vendedor?.firma || "");
  if (h.includes("firma") && h.includes("cliente")) return String(ctx?.cliente?.firma || "");
  return "";
}

function tryGetPdfFormFieldText(field) {
  try {
    if (typeof field?.getText === "function") return String(field.getText() || "");
  } catch (_) { }
  try {
    if (typeof field?.getValue === "function") return String(field.getValue() || "");
  } catch (_) { }
  return "";
}

function fillTemplateAcroFormIfPresent(srcDoc, tokenCtx) {
  try {
    const form = srcDoc?.getForm?.();
    if (!form) return;
    const fields = form.getFields?.() || [];
    for (const field of fields) {
      if (!field || typeof field.getName !== "function") continue;
      if (typeof field.setText !== "function") continue;
      const name = String(field.getName() || "");
      const current = tryGetPdfFormFieldText(field);
      const hint = `${name} ${current}`.trim();
      const value = resolveTemplateFieldByHint(hint, tokenCtx);
      if (!String(value || "").trim()) continue;
      try {
        field.setText(String(value));
      } catch (_) { }
    }
    try {
      form.flatten();
    } catch (_) { }
  } catch (_) { }
}

async function buildQuotePdfDocument(ev, quote, company, manager) {
  if (!window.PDFLib) return null;
  const { PDFDocument, StandardFonts } = window.PDFLib;
  const doc = await PDFDocument.create();
  const font = await doc.embedFont(StandardFonts.Helvetica);
  const fontBold = await doc.embedFont(StandardFonts.HelveticaBold);
  const fontOblique = await doc.embedFont(StandardFonts.HelveticaOblique);
  const pageSize = { w: 595.28, h: 841.89 }; // A4 points
  const margin = 26;
  const contentW = pageSize.w - margin * 2;
  const selectedTemplateId = String(quote?.templateId || "").trim();
  const isServiHospTemplate = selectedTemplateId === SERVIHOSP_TEMPLATE_ID;
  const activeTemplate = (quickTemplates || []).find((t) => String(t?.id || "") === selectedTemplateId) || null;
  const useCorporateBands = !selectedTemplateId || selectedTemplateId === CORPORATE_TEMPLATE_ID;
  const headerImageSource = useCorporateBands
    ? String(activeTemplate?.assets?.headerImage || DEFAULT_TEMPLATE_HEADER_IMAGE).trim()
    : String(activeTemplate?.assets?.headerImage || "").trim();
  const footerImageSource = isServiHospTemplate
    ? ""
    : (useCorporateBands
      ? String(activeTemplate?.assets?.footerImage || DEFAULT_TEMPLATE_FOOTER_IMAGE).trim()
      : String(activeTemplate?.assets?.footerImage || "").trim());
  const [headerImageDataUrl, footerImageDataUrl] = await Promise.all([
    resolveTemplateImageAssetToDataUrl(headerImageSource),
    resolveTemplateImageAssetToDataUrl(footerImageSource),
  ]);
  const headerImage = headerImageDataUrl ? await embedTemplateImageFromDataUrl(doc, headerImageDataUrl) : null;
  const footerImage = footerImageDataUrl ? await embedTemplateImageFromDataUrl(doc, footerImageDataUrl) : null;
  const headerSlotH = headerImage ? 34 : 0;
  const footerSlotH = footerImage ? 24 : 0;
  let page = doc.addPage([pageSize.w, pageSize.h]);
  let y = pageSize.h - margin;
  const headerBandH = 28;
  const quoteLabelColor = window.PDFLib.rgb(114 / 255, 169 / 255, 178 / 255);
  const quoteSectionColor = window.PDFLib.rgb(33 / 255, 89 / 255, 103 / 255);
  const quoteSectionBorderColor = window.PDFLib.rgb(24 / 255, 67 / 255, 78 / 255);
  const quoteGridBorderColor = isServiHospTemplate ? window.PDFLib.rgb(1, 1, 1) : window.PDFLib.rgb(0.86, 0.9, 0.96);
  const quoteSoftBorderColor = isServiHospTemplate ? window.PDFLib.rgb(1, 1, 1) : window.PDFLib.rgb(0.82, 0.88, 0.96);
  const quotePanelBorderColor = isServiHospTemplate ? window.PDFLib.rgb(1, 1, 1) : window.PDFLib.rgb(0.75, 0.84, 0.94);
  const quoteNotesBorderColor = isServiHospTemplate ? window.PDFLib.rgb(1, 1, 1) : window.PDFLib.rgb(0.72, 0.8, 0.9);
  const quoteSummaryBorderColor = isServiHospTemplate ? window.PDFLib.rgb(1, 1, 1) : window.PDFLib.rgb(0.68, 0.76, 0.86);
  const quoteSectionStrokeColor = isServiHospTemplate ? quoteSectionColor : quoteSectionBorderColor;
  const quoteCurrency = normalizeQuoteCurrency(quote?.currency);
  const quoteMoney = (n) => moneyQuote(n, quoteCurrency);

  const drawRect = (x, yTop, w, h, opts = {}) => {
    page.drawRectangle({
      x,
      y: yTop - h,
      width: w,
      height: h,
      color: opts.fill || undefined,
      borderColor: opts.border || undefined,
      borderWidth: opts.borderWidth ?? (opts.border ? 1 : 0),
    });
  };

  const drawPageBands = () => {
    if (headerImage) {
      const maxW = contentW;
      const maxH = 28;
      const scale = Math.min(maxW / Math.max(1, headerImage.width), maxH / Math.max(1, headerImage.height));
      const w = Math.max(10, headerImage.width * scale);
      const h = Math.max(10, headerImage.height * scale);
      const x = margin + (contentW - w) / 2;
      const yTop = pageSize.h - margin;
      page.drawImage(headerImage, { x, y: yTop - h, width: w, height: h });
    }
    if (footerImage) {
      const maxW = contentW;
      const maxH = 20;
      const scale = Math.min(maxW / Math.max(1, footerImage.width), maxH / Math.max(1, footerImage.height));
      const w = Math.max(10, footerImage.width * scale);
      const h = Math.max(10, footerImage.height * scale);
      const x = margin + (contentW - w) / 2;
      const yBase = margin;
      page.drawImage(footerImage, { x, y: yBase, width: w, height: h });
    }
  };

  const resetCursorForPage = () => {
    y = pageSize.h - margin - headerSlotH;
  };

  const drawQuotePageHeader = () => {
    const title = `Cotizacion ${quote?.code ? quote.code : ""}${quote?.version ? ` - V${quote.version}` : ""}`.trim();
    drawRect(margin, y, contentW, headerBandH, {
      fill: quoteSectionColor,
      border: quoteSectionStrokeColor,
      borderWidth: 1,
    });
    page.drawText(title || "Cotizacion", {
      x: margin + 10,
      y: y - 18,
      size: 13,
      font: fontBold,
      color: window.PDFLib.rgb(1, 1, 1),
    });
    y -= headerBandH + 6;
  };

  const ensure = (need = 16) => {
    if (y - need > margin + footerSlotH) return;
    page = doc.addPage([pageSize.w, pageSize.h]);
    drawPageBands();
    resetCursorForPage();
    y -= 2;
    drawQuotePageHeader();
  };

  drawPageBands();
  resetCursorForPage();

  const drawLine = (text, opts = {}) => {
    const lh = opts.lineHeight || 14;
    ensure(lh + 2);
    const size = opts.size || 10;
    const useFont = opts.italic ? fontOblique : (opts.bold ? fontBold : font);
    page.drawText(String(text || ""), {
      x: opts.x ?? margin,
      y,
      size,
      font: useFont,
    });
    y -= lh;
  };
  const wrapTextByWidth = (rawText, maxWidth, useFont, size) => {
    const text = String(rawText || "").trim();
    if (!text) return [""];
    const words = text.split(/\s+/).filter(Boolean);
    const lines = [];
    let current = "";
    for (const word of words) {
      const candidate = current ? `${current} ${word}` : word;
      const w = useFont.widthOfTextAtSize(candidate, size);
      if (w <= maxWidth || !current) {
        current = candidate;
      } else {
        lines.push(current);
        current = word;
      }
    }
    if (current) lines.push(current);
    return lines.length ? lines : [text];
  };

  const numberToWordsEs = (value) => {
    const units = ["", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE"];
    const teens = ["DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE"];
    const tens = ["", "", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"];
    const hundreds = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"];
    const belowHundred = (n) => {
      if (n < 10) return units[n];
      if (n < 20) return teens[n - 10];
      if (n < 30) return n === 20 ? "VEINTE" : `VEINTI${units[n - 20].toLowerCase()}`.toUpperCase();
      const t = Math.floor(n / 10);
      const u = n % 10;
      return u ? `${tens[t]} Y ${units[u]}` : tens[t];
    };
    const belowThousand = (n) => {
      if (n === 0) return "";
      if (n === 100) return "CIEN";
      const h = Math.floor(n / 100);
      const r = n % 100;
      if (!h) return belowHundred(r);
      return r ? `${hundreds[h]} ${belowHundred(r)}` : hundreds[h];
    };
    const toWords = (n) => {
      if (n === 0) return "CERO";
      const millions = Math.floor(n / 1000000);
      const thousands = Math.floor((n % 1000000) / 1000);
      const rest = n % 1000;
      const parts = [];
      if (millions) parts.push(millions === 1 ? "UN MILLON" : `${toWords(millions)} MILLONES`);
      if (thousands) parts.push(thousands === 1 ? "MIL" : `${belowThousand(thousands)} MIL`);
      if (rest) parts.push(belowThousand(rest));
      return parts.join(" ").replace(/\s+/g, " ").trim();
    };
    const amount = Math.max(0, Number(value || 0));
    const whole = Math.floor(amount);
    const cents = Math.round((amount - whole) * 100);
    const centsText = String(cents).padStart(2, "0");
    const moneda = whole === 1 ? "QUETZAL" : "QUETZALES";
    return `${toWords(whole)} ${moneda} CON ${centsText}/100`;
  };

  const formatDocDate = (iso) => {
    const raw = String(iso || "").trim();
    if (!raw) return "";
    const d = new Date(`${raw}T00:00:00`);
    if (Number.isNaN(d.getTime())) return raw;
    return d.toLocaleDateString("es-GT", { day: "2-digit", month: "2-digit", year: "numeric" });
  };

  const totals = getQuoteTotals(quote);
  const { catBuckets } = aggregateQuoteBuckets(quote || {});
  const grandTotalWords = numberToWordsQuoteEs(totals.total, quoteCurrency);

  // Header band (first page)
  drawQuotePageHeader();

  // Info grid (2 columns)
  const pairs = [
    ["Contacto", quote?.contact || manager?.name || ""],
    ["Codigo", quote?.code || ""],
    ["Encargado Evento", manager?.name || ""],
    ["Fecha Documento", quote?.docDate || ""],
    ["Email", quote?.email || manager?.email || ""],
    ["Telefono", quote?.phone || manager?.phone || ""],
    ["Institucion", quote?.companyName || company?.name || ""],
    ["Forma de Pago", quote?.paymentType || ""],
    ["Facturar A", quote?.billTo || company?.billTo || company?.businessName || ""],
    ["NIT", quote?.nit || company?.nit || ""],
    ["Direccion", quote?.address || company?.address || ""],
    ["No Personas", String(quote?.people || "")],
    ["Tipo Evento", quote?.eventType || ""],
    ["Fecha evento", quote?.eventDate || ev?.date || ""],
    ["Salon o Jardin", quote?.venue || ev?.salon || ""],
    ["Folio No", quote?.folio || ""],
    ["Horario y Evento", quote?.schedule || `${ev?.startTime || ""} a ${ev?.endTime || ""}`.trim()],
    ["Fecha Finalizacion", quote?.endDate || ev?.date || ""],
    ["", ""],
    ["Fecha Maxima Pago", quote?.dueDate || ""],
  ];
  const cellH = 16;
  const colW = contentW / 2;
  const keyW = 80;
  for (let i = 0; i < pairs.length; i += 2) {
    ensure(cellH * 2 + 8);
    const left = pairs[i];
    const right = pairs[i + 1] || ["", ""];
    const rowTop = y;

    // left cell
    drawRect(margin, rowTop, colW, cellH, { border: quoteGridBorderColor, borderWidth: 0.8 });
    drawRect(margin, rowTop, keyW, cellH, { fill: quoteLabelColor, border: quoteGridBorderColor, borderWidth: 0.8 });
    page.drawText(String(left[0] || ""), { x: margin + 4, y: rowTop - 11, size: 7.7, font: fontBold, color: window.PDFLib.rgb(1, 1, 1) });
    page.drawText(String(left[1] || ""), { x: margin + keyW + 4, y: rowTop - 11, size: 7.7, font });

    // right cell
    drawRect(margin + colW, rowTop, colW, cellH, { border: quoteGridBorderColor, borderWidth: 0.8 });
    drawRect(margin + colW, rowTop, keyW, cellH, { fill: quoteLabelColor, border: quoteGridBorderColor, borderWidth: 0.8 });
    page.drawText(String(right[0] || ""), { x: margin + colW + 4, y: rowTop - 11, size: 7.7, font: fontBold, color: window.PDFLib.rgb(1, 1, 1) });
    page.drawText(String(right[1] || ""), { x: margin + colW + keyW + 4, y: rowTop - 11, size: 7.7, font });

    y -= cellH;
  }
  y -= 8;

  // Table header
  const col1 = 52;
  const col3 = 84;
  const col4 = 84;
  const col2 = contentW - col1 - col3 - col4;
  const tableHeaderH = 18;
  ensure(tableHeaderH + 10);
  drawRect(margin, y, contentW, tableHeaderH, { fill: window.PDFLib.rgb(0.95, 0.97, 1), border: quoteSoftBorderColor, borderWidth: 1 });
  page.drawText("Cant", { x: margin + 6, y: y - 12, size: 8, font: fontBold });
  page.drawText("Descripcion", { x: margin + col1 + 6, y: y - 12, size: 8, font: fontBold });
  page.drawText("Precio U", { x: margin + col1 + col2 + 6, y: y - 12, size: 8, font: fontBold });
  page.drawText("Total", { x: margin + col1 + col2 + col3 + 6, y: y - 12, size: 8, font: fontBold });
  y -= tableHeaderH;

  const items = Array.isArray(quote?.items) ? quote.items : [];
  const grouped = new Map();
  for (const item of items) {
    const d = String(item?.serviceDate || "");
    if (!grouped.has(d)) grouped.set(d, []);
    grouped.get(d).push(item);
  }
  const orderedDates = Array.from(grouped.keys()).sort((a, b) => a.localeCompare(b));

  const drawServiceRow = (qtyText, descText, priceText, totalText, opts = {}) => {
    const h = opts.height || 16;
    const size = opts.size || 8;
    const textY = y - (opts.textOffsetY || 11);
    ensure(h + 2);
    drawRect(margin, y, contentW, h, { border: quoteGridBorderColor, borderWidth: 0.8, fill: opts.fill });
    // vertical lines
    page.drawLine({ start: { x: margin + col1, y: y }, end: { x: margin + col1, y: y - h }, thickness: 0.8, color: quoteGridBorderColor });
    page.drawLine({ start: { x: margin + col1 + col2, y: y }, end: { x: margin + col1 + col2, y: y - h }, thickness: 0.8, color: quoteGridBorderColor });
    page.drawLine({ start: { x: margin + col1 + col2 + col3, y: y }, end: { x: margin + col1 + col2 + col3, y: y - h }, thickness: 0.8, color: quoteGridBorderColor });
    const textColor = opts.color || window.PDFLib.rgb(0.08, 0.18, 0.29);
    page.drawText(String(qtyText || ""), { x: margin + 6, y: textY, size, font: opts.bold ? fontBold : font, color: textColor });
    page.drawText(String(descText || ""), { x: margin + col1 + 6, y: textY, size, font: opts.bold ? fontBold : font, color: textColor });
    page.drawText(String(priceText || ""), { x: margin + col1 + col2 + 6, y: textY, size, font: opts.bold ? fontBold : font, color: textColor });
    page.drawText(String(totalText || ""), { x: margin + col1 + col2 + col3 + 6, y: textY, size, font: opts.bold ? fontBold : font, color: textColor });
    y -= h;
  };

  for (const d of orderedDates) {
    const dayItems = grouped.get(d) || [];
    drawServiceRow("", `SERVICIOS DEL DIA ${formatDocDate(d)}`, "", "", {
      height: 15,
      fill: quoteSectionColor,
      bold: true,
      color: window.PDFLib.rgb(1, 1, 1),
    });
    let daySubtotal = 0;
    for (const item of dayItems) {
      const qty = Number(item?.qty || 0);
      const price = Number(item?.price || 0);
      const lineTotal = qty * price;
      daySubtotal += lineTotal;
      drawServiceRow(
        String(qty),
        String(item?.name || item?.description || ""),
        quoteMoney(price),
        quoteMoney(lineTotal)
      );
    }
    drawServiceRow("", "", `SUBTOTAL ${formatDocDate(d)}`, quoteMoney(daySubtotal), {
      height: 15,
      fill: quoteLabelColor,
      bold: true,
      size: 7.2,
      textOffsetY: 10.4,
      color: window.PDFLib.rgb(1, 1, 1),
    });
  }

  y -= 8;
  const hasDiscount = Number(totals.discountAmount || 0) > 0;
  const totalsPanelH = hasDiscount ? 86 : 72;
  ensure(totalsPanelH + 10);
  const panelX = margin;
  const panelY = y;
  const panelW = contentW;
  drawRect(panelX, panelY, panelW, totalsPanelH, {
    fill: window.PDFLib.rgb(0.985, 0.992, 1),
    border: quotePanelBorderColor,
    borderWidth: 1,
  });
  page.drawRectangle({
    x: panelX,
    y: panelY - totalsPanelH,
    width: 4,
    height: totalsPanelH,
    color: window.PDFLib.rgb(0.06, 0.41, 0.72),
  });

  const totalsLabelX = panelX + 14;
  const totalsValueRight = panelX + panelW - 12;
  const row1Y = panelY - 20;
  const row2Y = panelY - 34;
  const separatorY = hasDiscount ? panelY - 42 : panelY - 28;
  const totalBoxTopY = hasDiscount ? panelY - 64 : panelY - 50;
  const row3Y = hasDiscount ? panelY - 56 : panelY - 42;
  const money = (n) => quoteMoney(n);
  const drawSummaryLine = (lineY, label, value, opts = {}) => {
    const labelFont = opts.bold ? fontBold : font;
    const valueFont = opts.bold ? fontBold : font;
    const size = opts.size || 9.2;
    const valueText = String(value || "");
    const valueW = valueFont.widthOfTextAtSize(valueText, size);
    page.drawText(String(label || ""), {
      x: totalsLabelX,
      y: lineY,
      size,
      font: labelFont,
      color: opts.color || window.PDFLib.rgb(0.12, 0.24, 0.38),
    });
    page.drawText(valueText, {
      x: Math.max(totalsLabelX + 120, totalsValueRight - valueW),
      y: lineY,
      size,
      font: valueFont,
      color: opts.color || window.PDFLib.rgb(0.12, 0.24, 0.38),
    });
  };

  const discountType = String(totals.discountType || "");
  const discountValue = Number(totals.discountValue || 0);
  const discountLabel = discountType === "PERCENT"
    ? `Descuento (${discountValue.toFixed(2)}%)`
     : `Descuento (${quoteMoney(discountValue)})`;

  drawRect(panelX + 10, panelY - 10, panelW - 20, 14, {
    fill: quoteLabelColor,
    border: quoteSectionStrokeColor,
    borderWidth: 0.8,
  });
  drawSummaryLine(row1Y, "Subtotal evento", money(totals.subtotal), {
    size: 8.2,
    color: window.PDFLib.rgb(1, 1, 1),
  });
  if (hasDiscount) {
    drawRect(panelX + 10, panelY - 24, panelW - 20, 14, {
      fill: quoteLabelColor,
      border: quoteSectionStrokeColor,
      borderWidth: 0.8,
    });
    drawSummaryLine(row2Y, discountLabel, money(totals.discountAmount), {
      size: 8.2,
      color: window.PDFLib.rgb(1, 1, 1),
    });
  }

  page.drawLine({
    start: { x: panelX + 10, y: separatorY },
    end: { x: panelX + panelW - 10, y: separatorY },
    thickness: 0.8,
    color: window.PDFLib.rgb(0.78, 0.86, 0.95),
  });

  drawRect(panelX + 10, totalBoxTopY, panelW - 20, 18, {
    fill: quoteLabelColor,
    border: quoteSectionStrokeColor,
    borderWidth: 0.9,
  });
  drawSummaryLine(row3Y, "TOTAL EVENTO", money(totals.total), {
    bold: true,
    size: 10.8,
    color: window.PDFLib.rgb(1, 1, 1),
  });

  const wordsY = hasDiscount ? panelY - 78 : panelY - 64;
  page.drawText(`SON: ${grandTotalWords}`, {
    x: panelX + 12,
    y: wordsY,
    size: 7.8,
    font: fontBold,
    color: window.PDFLib.rgb(0.1, 0.25, 0.4),
  });
  y -= totalsPanelH;

  y -= 8;
  const notesTitleH = 14;
  const notesLines = [
    "Incrementos en menos de 24 horas de servicios y/o productos tendran cargo adicional del 10% sobre el excedente solicitado.",
    "El monto de anticipo para asegurar el evento depende de las clausulas de formalizacion.",
    "Estamos sujetos a pagos trimestrales. Aplicamos clausula de No Show.",
    "Esta propuesta se formaliza al tener firma del asesor de ventas y del cliente, junto a comprobante y/o orden de compra.",
    "No se reembolsa el anticipo si no realiza su evento por cualquier causa.",
  ];
  const notesFontSize = 7.2;
  const notesLineH = 9;
  const notesBodyPad = 5;
  const notesTextW = contentW - 12;
  const wrappedNotes = notesLines.map((line) => wrapTextByWidth(line, notesTextW, font, notesFontSize)).flat();
  const notesBodyH = notesBodyPad * 2 + wrappedNotes.length * notesLineH;
  ensure(notesTitleH + notesBodyH + 8);
  drawRect(margin, y, contentW, notesTitleH, {
    fill: window.PDFLib.rgb(0.15, 0.4, 0.47),
    border: window.PDFLib.rgb(0.11, 0.31, 0.37),
    borderWidth: 1,
  });
  page.drawText("NOTAS", {
    x: margin + (contentW / 2) - (fontBold.widthOfTextAtSize("NOTAS", 9) / 2),
    y: y - 10,
    size: 9,
    font: fontBold,
    color: window.PDFLib.rgb(1, 1, 1),
  });
  y -= notesTitleH;
  drawRect(margin, y, contentW, notesBodyH, {
    fill: window.PDFLib.rgb(1, 1, 1),
    border: quoteNotesBorderColor,
    borderWidth: 0.9,
  });
  let noteY = y - notesBodyPad - notesLineH + 2;
  for (const line of wrappedNotes) {
    page.drawText(String(line || ""), {
      x: margin + 6,
      y: noteY,
      size: notesFontSize,
      font,
      color: window.PDFLib.rgb(0.1, 0.16, 0.23),
    });
    noteY -= notesLineH;
  }
  y -= notesBodyH;

  y -= 6;
  const cargoRows = [
    { label: "Alimentos y Bebidas", amount: Number(catBuckets?.alimentosBebidas?.amount || 0) },
    { label: "Miscelaneos", amount: Number(catBuckets?.miscelaneos?.amount || 0) },
    { label: "Hospedaje JDL", amount: Number(catBuckets?.hospedajeJdl?.amount || 0) },
    { label: "Hospedaje Terceros", amount: Number(catBuckets?.hospedajeTerceros?.amount || 0) },
  ];
  const advanceRows = normalizeQuoteAdvancesForSnapshot(quote?.advances)
    .slice()
    .sort((a, b) => {
      const d = String(a.date || "").localeCompare(String(b.date || ""));
      if (d !== 0) return d;
      return String(a.createdAt || "").localeCompare(String(b.createdAt || ""));
    });
  const totalContratado = cargoRows.reduce((acc, r) => acc + Number(r.amount || 0), 0);
  const totalAnticipos = getQuoteAdvanceTotal(quote);
  const saldoPendiente = Math.max(0, totalContratado - totalAnticipos);
  const showMoney = (n) => (Number(n || 0) > 0 ? quoteMoney(n) : "-");
  const formatAdvanceDatePdf = (iso) => {
    const raw = String(iso || "").trim();
    if (!raw) return "-";
    const d = new Date(`${raw}T00:00:00`);
    if (Number.isNaN(d.getTime())) return raw;
    return d.toLocaleDateString("es-GT", { day: "2-digit", month: "2-digit", year: "numeric" });
  };

  const summaryTitleH = 14;
  const tableHeaderH2 = 14;
  const rowH2 = 13;
  const summaryRows = [
    ...cargoRows.map((r) => ({ label: r.label, amount: r.amount, tone: "base", bold: false })),
    { label: "TOTAL CONTRATADO", amount: totalContratado, tone: "em", bold: true },
    ...(
      advanceRows.length
        ? advanceRows.map((r) => ({
          label: `${formatAdvanceDatePdf(r.date)} - ${formatQuoteAdvanceDetail(r)}`,
          amount: Number(r.amount || 0),
          tone: "advance",
          bold: false,
          size: 7.4,
        }))
        : [{ label: "ANTICIPOS: SIN REGISTROS", amount: 0, tone: "base", bold: false }]
    ),
    { label: "TOTAL ANTICIPOS", amount: totalAnticipos, tone: "em", bold: true },
    { label: "SALDO", amount: saldoPendiente, tone: "final", bold: true },
  ];
  const summaryBodyH = tableHeaderH2 + summaryRows.length * rowH2;
  ensure(summaryTitleH + summaryBodyH + 8);

  drawRect(margin, y, contentW, summaryTitleH, {
    fill: window.PDFLib.rgb(0.15, 0.4, 0.47),
    border: window.PDFLib.rgb(0.11, 0.31, 0.37),
    borderWidth: 1,
  });
  page.drawText("RESUMEN DE CARGOS", {
    x: margin + (contentW / 2) - (fontBold.widthOfTextAtSize("RESUMEN DE CARGOS", 8.8) / 2),
    y: y - 10,
    size: 8.8,
    font: fontBold,
    color: window.PDFLib.rgb(1, 1, 1),
  });
  y -= summaryTitleH;

  const tX = margin + 22;
  const tW = contentW - 44;
  const descW = tW * 0.68;
  const borderCol = quoteSummaryBorderColor;
  const fitSummaryText = (text, useFont, size, maxWidth) => {
    const raw = String(text || "").trim();
    if (!raw) return "";
    if (useFont.widthOfTextAtSize(raw, size) <= maxWidth) return raw;
    const ellipsis = "...";
    let out = raw;
    while (out.length > 1 && useFont.widthOfTextAtSize(`${out}${ellipsis}`, size) > maxWidth) {
      out = out.slice(0, -1);
    }
    return `${out}${ellipsis}`;
  };
  const drawSummaryGridRow = (yy, h, descText, valNumber, opts = {}) => {
    const fill = opts.fill || window.PDFLib.rgb(1, 1, 1);
    const textColor = opts.color || window.PDFLib.rgb(0.08, 0.18, 0.29);
    drawRect(tX, yy, tW, h, {
      fill,
      border: borderCol,
      borderWidth: 0.7,
    });
    page.drawLine({ start: { x: tX + descW, y: yy }, end: { x: tX + descW, y: yy - h }, thickness: 0.7, color: borderCol });
    const rowFont = opts.bold ? fontBold : font;
    const rowSize = opts.size || 8;
    const safeDesc = fitSummaryText(String(descText || ""), rowFont, rowSize, descW - 14);
    page.drawText(safeDesc, {
      x: tX + 6,
      y: yy - h + 3.8,
      size: rowSize,
      font: rowFont,
      color: textColor,
    });
    const valText = showMoney(valNumber);
    const valFont = rowFont;
    const valSize = rowSize;
    const valWidth = valFont.widthOfTextAtSize(valText, valSize);
    page.drawText(valText, {
      x: tX + tW - 8 - valWidth,
      y: yy - h + 3.8,
      size: valSize,
      font: valFont,
      color: textColor,
    });
  };

  drawRect(tX, y, tW, tableHeaderH2, { fill: quoteLabelColor, border: quoteSectionStrokeColor, borderWidth: 0.7 });
  page.drawLine({ start: { x: tX + descW, y }, end: { x: tX + descW, y: y - tableHeaderH2 }, thickness: 0.7, color: borderCol });
  page.drawText("DESCRIPCION", {
    x: tX + 6,
    y: y - 9,
    size: 8,
    font: fontBold,
    color: window.PDFLib.rgb(1, 1, 1),
  });
  page.drawText("TOTAL", {
    x: tX + tW - 8 - fontBold.widthOfTextAtSize("TOTAL", 8),
    y: y - 9,
    size: 8,
    font: fontBold,
    color: window.PDFLib.rgb(1, 1, 1),
  });
  y -= tableHeaderH2;

  for (let i = 0; i < summaryRows.length; i++) {
    const row = summaryRows[i];
    const baseFill = i % 2 === 0 ? window.PDFLib.rgb(1, 1, 1) : window.PDFLib.rgb(0.985, 0.992, 1);
    const fill = row.tone === "final"
      ? quoteLabelColor
      : row.tone === "em"
        ? quoteLabelColor
        : row.tone === "advance"
          ? window.PDFLib.rgb(0.975, 0.985, 1)
          : baseFill;
    const rowColor = row.tone === "final" || row.tone === "em"
      ? window.PDFLib.rgb(1, 1, 1)
      : undefined;
    drawSummaryGridRow(y, rowH2, row.label, row.amount, { bold: row.bold, fill, size: row.size || 8, color: rowColor });
    y -= rowH2;
  }
}

async function appendTemplateToFinalPdf(finalDoc, template, ev, quote, company, manager) {
  if (!window.PDFLib) return;
  const { PDFDocument, StandardFonts } = window.PDFLib;
  const pdfData = String(template?.assets?.pagePdf || "").trim();
  if (!pdfData) return;
  const src = await PDFDocument.load(dataUrlToUint8Array(pdfData));
  const tokenCtx = buildTemplateTokenContext(ev, quote, company, manager, template);
  fillTemplateAcroFormIfPresent(src, tokenCtx);
  const pageIndices = src.getPageIndices();
  const copied = await finalDoc.copyPages(src, pageIndices);
  for (const p of copied) finalDoc.addPage(p);

  const templatePages = finalDoc.getPages().slice(-copied.length);
  const totalHeight = templatePages.reduce((acc, p) => acc + p.getHeight(), 0);
  if (!totalHeight) return;
  const fontCache = new Map();
  const getFont = async (field) => {
    const fontName = mapTemplateFontToPdfLib(StandardFonts, field);
    if (fontCache.has(fontName)) return fontCache.get(fontName);
    const f = await finalDoc.embedFont(fontName);
    fontCache.set(fontName, f);
    return f;
  };

  let yTopOffset = 0;
  const pageRanges = templatePages.map((p) => {
    const h = p.getHeight();
    const from = yTopOffset;
    const to = yTopOffset + h;
    yTopOffset += h;
    return { page: p, from, to, h };
  });

  const positioned = Array.isArray(template?.positionedFields) ? template.positionedFields : [];
  for (const field of positioned) {
    try {
      const text = resolveTemplateTokenValue(field?.token, tokenCtx, template) || String(field?.label || "").trim();
      if (!text) continue;
      const isSignature = field?.isSignature === true || isTemplateSignatureToken(field?.token);
      const hasYPt = Number.isFinite(Number(field?.yPt));
      const yPctTop = hasYPt
        ? clamp((Number(field.yPt) / TEMPLATE_COORD_BASE_H_PT) * 100, 0, 100)
        : clamp(Number(field?.y || 0), 0, 100);
      const yAbsTop = (yPctTop / 100) * totalHeight;
      const target = pageRanges.find((r) => yAbsTop >= r.from && yAbsTop <= r.to) || pageRanges[pageRanges.length - 1];
      if (!target) continue;
      const hasXPt = Number.isFinite(Number(field?.xPt));
      const xPct = (hasXPt
        ? clamp((Number(field.xPt) / TEMPLATE_COORD_BASE_W_PT) * 100, 0, 100)
        : clamp(Number(field?.x || 0), 0, 100)) / 100;
      const fontSize = clamp(Number(field?.fontSize || 12), 8, 72);
      const xLeft = xPct * target.page.getWidth();
      const topOnPage = yAbsTop - target.from;
      let yCursor = target.h - topOnPage;

      if (isSignature && isImageDataUrl(text)) {
        const hasWPt = Number.isFinite(Number(field?.wPt));
        const hasHPt = Number.isFinite(Number(field?.hPt));
        const boxWPct = hasWPt
          ? clamp((Number(field.wPt) / TEMPLATE_COORD_BASE_W_PT) * 100, TEMPLATE_SIGNATURE_MIN_W_PCT, TEMPLATE_SIGNATURE_MAX_W_PCT)
          : clamp(Number(field?.w || TEMPLATE_SIGNATURE_FALLBACK_W_PCT), TEMPLATE_SIGNATURE_MIN_W_PCT, TEMPLATE_SIGNATURE_MAX_W_PCT);
        const boxHPct = hasHPt
          ? clamp((Number(field.hPt) / TEMPLATE_COORD_BASE_H_PT) * 100, TEMPLATE_SIGNATURE_MIN_H_PCT, TEMPLATE_SIGNATURE_MAX_H_PCT)
          : clamp(Number(field?.h || TEMPLATE_SIGNATURE_FALLBACK_H_PCT), TEMPLATE_SIGNATURE_MIN_H_PCT, TEMPLATE_SIGNATURE_MAX_H_PCT);
        const boxW = (boxWPct / 100) * target.page.getWidth();
        const boxH = (boxHPct / 100) * target.h;
        const drawW = Math.max(10, boxW);
        const drawH = Math.max(10, boxH);
        const image = await embedTemplateImageFromDataUrl(finalDoc, text);
        if (!image) {
          console.warn("Firma no valida para plantilla:", field?.token || field?.label || "-");
          continue;
        }
        const scale = Math.min(drawW / Math.max(1, image.width), drawH / Math.max(1, image.height));
        const iw = Math.max(10, image.width * scale);
        const ih = Math.max(10, image.height * scale);
        const x = clamp(xLeft + (drawW - iw) / 2, 0, Math.max(0, target.page.getWidth() - iw));
        const y = clamp((target.h - topOnPage) - drawH + (drawH - ih) / 2, 0, Math.max(0, target.h - ih));
        target.page.drawImage(image, { x, y, width: iw, height: ih });
        continue;
      }

      const font = await getFont(field);
      const xCenter = xLeft;
      const lines = String(text).split(/\r?\n/);
      for (const ln of lines) {
        const line = String(ln || "");
        const w = font.widthOfTextAtSize(line, fontSize);
        const x = clamp(xCenter - (w / 2), 10, Math.max(10, target.page.getWidth() - w - 10));
        target.page.drawText(line, { x, y: Math.max(8, yCursor), size: fontSize, font });
        yCursor -= fontSize * 1.2;
      }
    } catch (err) {
      console.warn("Campo de plantilla omitido por error:", field?.token || field?.label || "-", err?.message || err);
      continue;
    }
  }
}

async function tryOpenCombinedQuotePdf(ev, quote) {
  if (!window.PDFLib || !quote) return false;
  const company = (state.companies || []).find((c) => c.id === quote.companyId);
  const manager = company?.managers?.find((m) => m.id === quote.managerId);
  const quotePdf = await buildQuotePdfDocument(ev, quote, company, manager);
  if (!quotePdf) return false;

  const { PDFDocument } = window.PDFLib;
  const finalDoc = await PDFDocument.create();
  const quotePages = await finalDoc.copyPages(quotePdf, quotePdf.getPageIndices());
  for (const p of quotePages) finalDoc.addPage(p);

  const templateId = String(quote.templateId || "").trim();
  if (templateId) {
    const template = (quickTemplates || []).find((t) => String(t?.id || "") === templateId);
    if (template && String(template?.assets?.pagePdf || "").trim()) {
      await appendTemplateToFinalPdf(finalDoc, template, ev, quote, company, manager);
    }
  }

  const bytes = await finalDoc.save();
  const blob = new Blob([bytes], { type: "application/pdf" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = `cotizacion_${String(quote.code || "final").replace(/[^a-zA-Z0-9_-]+/g, "_")}.pdf`;
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 2000);
  return true;
}

function normalizeQuotePrintVariant(rawVariant) {
  const value = String(rawVariant || "").trim().toLowerCase();
  if (value === "with_menu" || value === "completa" || value === "full") return "with_menu";
  if (value === "without_prices" || value === "sin_precios" || value === "commercial") return "without_prices";
  return "standard";
}

function getQuotePrintVariantMeta(rawVariant) {
  const variant = normalizeQuotePrintVariant(rawVariant);
  if (variant === "with_menu") {
    return {
      variant,
      title: "Cotizacion completa",
      shortTitle: "Completa",
      description: "Incluye precios, menu y montaje, y contrato al final.",
      includePrices: true,
      includeMenuMontaje: true,
      includeContract: true,
      modeBadge: "VERSION COMPLETA CON MENU Y MONTAJE",
    };
  }
  if (variant === "without_prices") {
    return {
      variant,
      title: "Cotizacion comercial sin precios",
      shortTitle: "Sin precios",
      description: "Oculta precios y totales. Incluye menu y montaje. No incluye contrato.",
      includePrices: false,
      includeMenuMontaje: true,
      includeContract: false,
      modeBadge: "VERSION COMERCIAL SIN PRECIOS",
    };
  }
  return {
    variant: "standard",
    title: "Cotizacion estandar",
    shortTitle: "Estandar",
    description: "Mantiene el formato actual con precios y contrato.",
    includePrices: true,
    includeMenuMontaje: false,
    includeContract: true,
    modeBadge: "COTIZACION ESTANDAR",
  };
}

async function promptQuotePrintVariant(ev, quote) {
  const hasMenuMontaje = Array.isArray(quote?.menuMontajeEntries)
    && quote.menuMontajeEntries.some((row) => String(row?.date || "").trim() && String(row?.salon || "").trim());

  const cardOptions = [
    {
      value: "standard",
      title: "Estandar",
      description: "Cotizacion con precios y contrato, igual que el formato actual.",
      tone: "slate",
      disabled: false,
    },
    {
      value: "with_menu",
      title: "Completa con Menu & Montaje",
      description: "Cotizacion con precios, luego Menu & Montaje, y al final el contrato.",
      tone: "blue",
      disabled: !hasMenuMontaje,
    },
    {
      value: "without_prices",
      title: "Sin precios",
      description: "Muestra la cotizacion sin precios ni totales, y abajo agrega Menu & Montaje sin contrato.",
      tone: "green",
      disabled: !hasMenuMontaje,
    },
  ];

  if (window.Swal && typeof window.Swal.fire === "function") {
    const cardsHtml = cardOptions.map((opt) => `
      <button
        type="button"
        class="quotePrintOptionCard quotePrintOptionCard--${opt.tone}${opt.value === "standard" ? " isActive" : ""}${opt.disabled ? " isDisabled" : ""}"
        data-quote-print-value="${opt.value}"
        ${opt.disabled ? 'disabled aria-disabled="true"' : ''}
      >
        <span class="quotePrintOptionTitle">${opt.title}</span>
        <span class="quotePrintOptionText">${opt.description}</span>
        ${opt.disabled ? '<span class="quotePrintOptionNote">Requiere Menu & Montaje guardado</span>' : ''}
      </button>`).join('');

    const result = await window.Swal.fire({
      title: "Formato de impresion",
      html: `
        <style>
          .quotePrintPicker{ display:grid; gap:12px; margin-top:10px; text-align:left; }
          .quotePrintOptionCard{
            width:100%;
            border:none;
            text-align:left;
            padding:14px 16px;
            border-radius:16px;
            background:#ffffff;
            box-shadow:inset 0 0 0 1px rgba(203,213,225,.9), 0 8px 20px rgba(15,23,42,.08);
            transition:transform .15s ease, box-shadow .15s ease, background .15s ease;
            cursor:pointer;
            display:grid;
            gap:6px;
          }
          .quotePrintOptionCard:hover:not(.isDisabled){
            transform:translateY(-1px);
            box-shadow:inset 0 0 0 1px rgba(147,197,253,.65), 0 10px 24px rgba(37,99,235,.12);
          }
          .quotePrintOptionCard.isActive{
            box-shadow:inset 0 0 0 2px rgba(148,163,184,.4), 0 12px 28px rgba(15,23,42,.1);
          }
          .quotePrintOptionCard--slate.isActive{
            background:#f1f5f9;
            box-shadow:inset 0 0 0 2px rgba(148,163,184,.46), 0 12px 28px rgba(2,6,23,.24);
          }
          .quotePrintOptionCard--blue.isActive{
            background:#eff6ff;
            box-shadow:inset 0 0 0 2px rgba(96,165,250,.56), 0 12px 28px rgba(2,6,23,.24);
          }
          .quotePrintOptionCard--green.isActive{
            background:#ecfdf5;
            box-shadow:inset 0 0 0 2px rgba(52,211,153,.48), 0 12px 28px rgba(2,6,23,.24);
          }
          .quotePrintOptionCard.isDisabled{
            opacity:.58;
            cursor:not-allowed;
            filter:saturate(.72);
          }
          .quotePrintOptionTitle{
            display:block;
            font-size:16px;
            font-weight:900;
            color:#0f172a;
            letter-spacing:.1px;
          }
          .quotePrintOptionText{
            display:block;
            font-size:12px;
            line-height:1.45;
            color:#475569;
          }
          .quotePrintOptionNote{
            display:inline-flex;
            width:max-content;
            margin-top:2px;
            padding:4px 8px;
            border-radius:999px;
            background:rgba(148,163,184,.14);
            color:#334155;
            font-size:11px;
            font-weight:700;
          }
        </style>
        <input type="hidden" id="quotePrintVariantInput" value="standard" />
        <div class="quotePrintPicker">${cardsHtml}</div>
      `,
      showCancelButton: true,
      confirmButtonText: "Abrir documento",
      cancelButtonText: "Cancelar",
      background: "#f8fbff",
      color: "#10243b",
      confirmButtonColor: "#2563eb",
      cancelButtonColor: "#94a3b8",
      didOpen: (popup) => {
        const hidden = popup.querySelector('#quotePrintVariantInput');
        const cards = Array.from(popup.querySelectorAll('[data-quote-print-value]'));
        const activate = (value) => {
          if (!hidden) return;
          hidden.value = value;
          cards.forEach((card) => {
            card.classList.toggle('isActive', card.dataset.quotePrintValue === value);
          });
        };
        cards.forEach((card) => {
          card.addEventListener('click', () => {
            if (card.disabled) return;
            activate(card.dataset.quotePrintValue || 'standard');
          });
        });
      },
      preConfirm: () => {
        const hidden = document.getElementById('quotePrintVariantInput');
        const value = normalizeQuotePrintVariant(hidden?.value || '');
        if (!value) {
          window.Swal.showValidationMessage('Selecciona un formato.');
          return false;
        }
        if ((value === 'with_menu' || value === 'without_prices') && !hasMenuMontaje) {
          window.Swal.showValidationMessage('Esta cotizacion aun no tiene Menu & Montaje guardado.');
          return false;
        }
        return value;
      },
    });
    if (!result?.isConfirmed) return null;
    return normalizeQuotePrintVariant(result.value);
  }

  let answer = window.prompt(
    hasMenuMontaje
      ? "Elige formato: standard, with_menu o without_prices"
      : "Elige formato: standard (Menu & Montaje no esta disponible aun)",
    "standard"
  );
  answer = normalizeQuotePrintVariant(answer);
  if ((answer === "with_menu" || answer === "without_prices") && !hasMenuMontaje) {
    toast("Esta cotizacion aun no tiene Menu & Montaje guardado.");
    return null;
  }
  return answer;
}

function buildQuoteMenuMontajeAttachmentHtml(ev, quoteLike) {
  const quote = quoteLike || {};
  const entries = (Array.isArray(quote?.menuMontajeEntries) ? quote.menuMontajeEntries : [])
    .filter((x) => String(x?.date || "").trim() && String(x?.salon || "").trim());
  if (!entries.length) return "";
  const company = (state.companies || []).find((c) => String(c.id || "") === String(quote.companyId || ""));
  const institutionName = String(company?.name || quote.companyName || "-").trim() || "-";
  const manager = company?.managers?.find((m) => String(m.id || "") === String(quote.managerId || ""));
  const seller = (state.users || []).find((u) => String(u.id || "") === String(ev?.userId || ""));
  const byDate = new Map();
  for (const row of entries) {
    if (!byDate.has(row.date)) byDate.set(row.date, []);
    byDate.get(row.date).push(row);
  }
  const orderedDates = Array.from(byDate.keys()).sort((a, b) => a.localeCompare(b));
  const cardsHtml = orderedDates.map((date) => {
    const rows = byDate.get(date) || [];
    rows.sort((a, b) => String(a.salon || "").localeCompare(String(b.salon || "")));
    const blocks = rows.map((r) => `
      <article class="quoteMmDayCard">
        <div class="quoteMmDayHead">
          <div class="quoteMmDayTitle">${escapeHtml(String(r.salon || "").toUpperCase())}</div>
          <div class="quoteMmDayMeta">${escapeHtml(date)}</div>
        </div>
        <div class="quoteMmBlock">
          <div class="quoteMmBlockTitle">Menu</div>
          <div class="quoteMmBlockBody">${renderMenuMontajeRichText(String(r.menuDescription || "")) || '<p class="quoteMmEmpty">Sin detalle de menu.</p>'}</div>
        </div>
        <div class="quoteMmBlock">
          <div class="quoteMmBlockTitle">Montaje</div>
          <div class="quoteMmBlockBody">${renderMenuMontajeRichText(String(r.montajeDescription || "")) || '<p class="quoteMmEmpty">Sin detalle de montaje.</p>'}</div>
        </div>
      </article>`).join('');
    return `
      <section class="quoteMmDateSection">
        <div class="quoteMmSectionHead">
          <div class="quoteMmSectionTitle">Menu & Montaje - ${escapeHtml(date)}</div>
          <div class="quoteMmSectionMeta">${escapeHtml(institutionName)}</div>
        </div>
        <div class="quoteMmMetaGrid">
          <div><b>Encargado evento:</b> ${escapeHtml(String(manager?.name || quote.contact || '-'))}</div>
          <div><b>No cotizacion:</b> ${escapeHtml(String(quote.code || '-'))}</div>
          <div><b>Tipo evento:</b> ${escapeHtml(String(quote.eventType || ev?.name || '-'))}</div>
          <div><b>Horario:</b> ${escapeHtml(String(quote.schedule || `${ev?.startTime || ''} a ${ev?.endTime || ''}`.trim() || '-'))}</div>
          <div><b>No pax:</b> ${escapeHtml(String(quote.people || ev?.pax || '-'))}</div>
          <div><b>Vendedor:</b> ${escapeHtml(String(seller?.fullName || seller?.name || '-'))}</div>
        </div>
        <div class="quoteMmCards">${blocks}</div>
      </section>`;
  }).join('');

  return `
    <section class="templateAttach quoteMenuMontajeAttach">
      <div class="templateAttachHead">Menu y Montaje</div>
      <div class="templateAttachBody">
        <div class="quoteMmWrap">${cardsHtml}</div>
      </div>
    </section>`;
}

async function openQuoteDocument(ev, quote, printOptions = null) {
  if (!ev || !quote) return;
  const selectedVariant = printOptions?.variant
    ? normalizeQuotePrintVariant(printOptions.variant)
    : await promptQuotePrintVariant(ev, quote);
  if (!selectedVariant) return;
  const printMeta = getQuotePrintVariantMeta(selectedVariant);
  const showPrices = printMeta.includePrices !== false;
  const includeMenuMontaje = printMeta.includeMenuMontaje === true;
  const includeContract = printMeta.includeContract !== false;
  const isMacClient = /Mac|iPhone|iPad|iPod/i.test(String(window?.navigator?.platform || "")) || /Mac OS X/i.test(String(window?.navigator?.userAgent || ""));
  const shouldUseStablePdf = selectedVariant === "standard" && !includeMenuMontaje && includeContract;
  if (shouldUseStablePdf) {
    try {
      const opened = await tryOpenCombinedQuotePdf(ev, quote);
      if (opened) {
        if (isMacClient) toast("Cotizacion PDF generada correctamente.");
        return;
      }
    } catch (err) {
      console.warn("Fallback a vista de impresion HTML:", err?.message || err);
    }
  }
  const company = (state.companies || []).find(c => c.id === quote.companyId);
  const manager = company?.managers?.find(m => m.id === quote.managerId);
  const items = Array.isArray(quote.items) ? quote.items : [];
  const quoteCurrency = normalizeQuoteCurrency(quote?.currency);
  const quoteMoney = (n) => moneyQuote(n, quoteCurrency);
  const numberToWordsEs = (value) => {
    const units = ["", "UNO", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", "OCHO", "NUEVE"];
    const teens = ["DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE"];
    const tens = ["", "", "VEINTE", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", "SETENTA", "OCHENTA", "NOVENTA"];
    const hundreds = ["", "CIENTO", "DOSCIENTOS", "TRESCIENTOS", "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", "OCHOCIENTOS", "NOVECIENTOS"];

    const belowHundred = (n) => {
      if (n < 10) return units[n];
      if (n < 20) return teens[n - 10];
      if (n < 30) {
        if (n === 20) return "VEINTE";
        const u = n - 20;
        return `VEINTI${units[u].toLowerCase()}`.toUpperCase();
      }
      const t = Math.floor(n / 10);
      const u = n % 10;
      return u ? `${tens[t]} Y ${units[u]}` : tens[t];
    };

    const belowThousand = (n) => {
      if (n === 0) return "";
      if (n === 100) return "CIEN";
      const h = Math.floor(n / 100);
      const r = n % 100;
      if (!h) return belowHundred(r);
      return r ? `${hundreds[h]} ${belowHundred(r)}` : hundreds[h];
    };

    const toWords = (n) => {
      if (n === 0) return "CERO";
      const millions = Math.floor(n / 1000000);
      const thousands = Math.floor((n % 1000000) / 1000);
      const rest = n % 1000;
      const parts = [];

      if (millions) {
        if (millions === 1) parts.push("UN MILLON");
        else parts.push(`${toWords(millions)} MILLONES`);
      }
      if (thousands) {
        if (thousands === 1) parts.push("MIL");
        else parts.push(`${belowThousand(thousands)} MIL`);
      }
      if (rest) parts.push(belowThousand(rest));
      return parts.join(" ").replace(/\s+/g, " ").trim();
    };

    const amount = Math.max(0, Number(value || 0));
    const whole = Math.floor(amount);
    const cents = Math.round((amount - whole) * 100);
    const words = toWords(whole);
    const centsText = String(cents).padStart(2, "0");
    const moneda = whole === 1 ? "QUETZAL" : "QUETZALES";
    return `${words} ${moneda} CON ${centsText}/100`;
  };
  const formatDocDate = (iso) => {
    const raw = String(iso || "").trim();
    if (!raw) return "";
    const d = new Date(`${raw}T00:00:00`);
    if (Number.isNaN(d.getTime())) return raw;
    return d.toLocaleDateString("es-GT", { day: "2-digit", month: "2-digit", year: "numeric" });
  };

  const grouped = new Map();
  for (const item of items) {
    const dateKey = String(item.serviceDate || "");
    if (!grouped.has(dateKey)) grouped.set(dateKey, []);
    grouped.get(dateKey).push(item);
  }
  const orderedDates = Array.from(grouped.keys()).sort((a, b) => a.localeCompare(b));

  const rowHtml = [];
  for (const d of orderedDates) {
    const dayItems = grouped.get(d) || [];
    const daySubtotal = dayItems.reduce((acc, x) => acc + Number(x.qty || 0) * Number(x.price || 0), 0);
    rowHtml.push(`
      <tr class="dayHeaderRow">
        <td colspan="${showPrices ? 4 : 2}">SERVICIOS DEL DIA ${escapeHtml(formatDocDate(d))}</td>
      </tr>
    `);
    for (const item of dayItems) {
      const qty = Number(item.qty || 0);
      const price = Number(item.price || 0);
      const lineTotal = qty * price;
      rowHtml.push(`
        <tr>
          <td style="text-align:right">${qty}</td>
          <td>${escapeHtml(item.name || item.description || "")}</td>
          <td style="text-align:right">${showPrices ? quoteMoney(price) : ""}</td>
          <td style="text-align:right">${showPrices ? quoteMoney(lineTotal) : ""}</td>
        </tr>
      `);
    }
    if (showPrices) {
      rowHtml.push(`
      <tr class="daySubtotalRow">
        <td colspan="2"></td>
        <td class="sumLabel">SUBTOTAL ${escapeHtml(formatDocDate(d))}</td>
        <td class="sumValue">${quoteMoney(daySubtotal)}</td>
      </tr>
    `);
    }
  }
  const totalsDoc = getQuoteTotals(quote);
  const { catBuckets: catBucketsDoc } = aggregateQuoteBuckets(quote || {});
  const subtotalDoc = totalsDoc.subtotal;
  const discountDoc = totalsDoc.discountAmount;
  const totalDoc = totalsDoc.total;
  const discountTypeDoc = totalsDoc.discountType;
  const discountValueDoc = totalsDoc.discountValue;
  const showDiscountRowDoc = Number(discountDoc || 0) > 0;
  const discountRowHtmlDoc = showDiscountRowDoc ? `
        <tr>
          <td colspan="2"></td>
          <td class="sumLabel">DESCUENTO ${discountTypeDoc === "PERCENT" ? `(${discountValueDoc.toFixed(2)}%)` : `(${quoteMoney(discountValueDoc)})`}</td>
          <td class="sumValue">${quoteMoney(discountDoc)}</td>
        </tr>
  ` : "";
  const cargoRowsDoc = [
    { label: "Alimentos y Bebidas", amount: Number(catBucketsDoc?.alimentosBebidas?.amount || 0) },
    { label: "Miscelaneos", amount: Number(catBucketsDoc?.miscelaneos?.amount || 0) },
    { label: "Hospedaje JDL", amount: Number(catBucketsDoc?.hospedajeJdl?.amount || 0) },
    { label: "Hospedaje Terceros", amount: Number(catBucketsDoc?.hospedajeTerceros?.amount || 0) },
  ];
  const advanceRowsDoc = normalizeQuoteAdvancesForSnapshot(quote?.advances)
    .slice()
    .sort((a, b) => {
      const d = String(a.date || "").localeCompare(String(b.date || ""));
      if (d !== 0) return d;
      return String(a.createdAt || "").localeCompare(String(b.createdAt || ""));
    });
  const totalContratadoDoc = cargoRowsDoc.reduce((acc, r) => acc + Number(r.amount || 0), 0);
  const totalAnticiposDoc = getQuoteAdvanceTotal(quote);
  const saldoDoc = Math.max(0, totalContratadoDoc - totalAnticiposDoc);
  const moneyQDoc = (n) => Number(n || 0) > 0 ? quoteMoney(n) : "-";
  const grandTotalWords = numberToWordsQuoteEs(totalDoc, quoteCurrency);
  const isServiHospTemplate = String(quote.templateId || "").trim() === SERVIHOSP_TEMPLATE_ID;
  let appendedTemplateHtml = "";
  const builtInTemplateMeta = getBuiltInQuoteTemplateMeta(String(quote.templateId || "").trim());
  if (includeContract && builtInTemplateMeta?.attachToQuote) {
    try {
      const res = await fetch(builtInTemplateMeta.file, { cache: "no-store" });
      if (res.ok) {
        let tplHtml = await res.text();
        tplHtml = applyBuiltInTemplateAssets(tplHtml, builtInTemplateMeta);
        tplHtml = fillTemplateHtmlTokens(tplHtml, buildTemplatePrintContextFromQuoteData(ev, quote, company, manager));
        const parser = new DOMParser();
        const tplDoc = parser.parseFromString(tplHtml, "text/html");
        const rawTplStyleText = tplDoc.head
          ? Array.from(tplDoc.head.querySelectorAll("style"))
              .map((node) => node.textContent || "")
              .join("\n")
          : "";
        const tplStyleText = rawTplStyleText
          .replace(/@media print\s*\{[\s\S]*?\}(?=\.[a-zA-Z0-9_-]+\{|[a-z]+\{|$)/g, "")
          .replace(/@page\s*\{[\s\S]*?\}/g, "")
          .replace(/html,body\{[\s\S]*?\}/g, "")
          .replace(/body\.c26\.doc-content\{[\s\S]*?\}/g, "")
          .replace(/\.contract-header-print\{[\s\S]*?\}/g, "")
          .replace(/\.contract-footer-print\{[\s\S]*?\}/g, "")
          .replace(/\.contract-header\{[\s\S]*?\}/g, "")
          .replace(/\.contract-footer\{[\s\S]*?\}/g, "")
          .replace(/\.contract-band\{[\s\S]*?\}/g, "");
        const tplStyles = tplStyleText ? `<style>${tplStyleText}</style>` : "";
        let tplBody = tplDoc.body ? tplDoc.body.innerHTML : tplHtml;
        if (tplDoc.body) {
          const headerImg = tplDoc.body.querySelector(".contract-header img, .contract-header-print img");
          const footerImg = isServiHospTemplate ? null : tplDoc.body.querySelector(".contract-footer img, .contract-footer-print img");
          tplDoc.body
            .querySelectorAll(".contract-header, .contract-header-print, .contract-footer, .contract-footer-print, script")
            .forEach((node) => node.remove());
          if (headerImg || footerImg) {
            const repeatFrameStyles = `<style>
  .contractPrintFrame{ width:100%; border-collapse:collapse; table-layout:fixed; }
  .contractPrintHead,.contractPrintFoot{ display:none; }
  .contractPrintCell{ padding:0; }
  .contractPrintRow{ break-inside:auto; page-break-inside:auto; }
  .contractPrintContent{ padding:0; }
  .contractPrintContent > :first-child{ margin-top:0 !important; }
  .contractPrintContent .c1:first-child{ display:none; }
  @media print{
    .contractTemplateAttach{
      padding:0;
      border-top:none;
      background:#ffffff;
      break-before:page;
    }
    .contractTemplateAttach .templateAttachHead{
      display:none;
    }
    .contractTemplateBody{
      border:none !important;
      border-radius:0 !important;
      overflow:visible !important;
      background:#fff !important;
    }
    .contractPrintFrame thead{ display:table-header-group; }
    .contractPrintFrame tfoot{ display:table-footer-group; }
    .contractPrintHead,.contractPrintFoot{ display:table-row; }
    .contractPrintHead .contractPrintCell{ padding:0 0 4mm; }
    .contractPrintFoot .contractPrintCell{ padding:4mm 0 0; }
    .contractPrintHead img{
      display:block;
      width:100%;
      max-height:17mm;
      object-fit:contain;
      object-position:center top;
    }
    .contractPrintFoot img{
      display:block;
      width:100%;
      max-height:11mm;
      object-fit:contain;
      object-position:center bottom;
    }
  }
</style>`;
            const repeatHead = headerImg
              ? `<tr class="contractPrintHead"><td class="contractPrintCell">${headerImg.outerHTML}</td></tr>`
              : "";
            const repeatFoot = footerImg
              ? `<tr class="contractPrintFoot"><td class="contractPrintCell">${footerImg.outerHTML}</td></tr>`
              : "";
            const bodyRows = Array.from(tplDoc.body.children)
              .map((node) => `<tr class="contractPrintRow"><td class="contractPrintContent">${node.outerHTML}</td></tr>`)
              .join("");
            tplBody = `${repeatFrameStyles}<table class="contractPrintFrame"><thead>${repeatHead}</thead><tfoot>${repeatFoot}</tfoot><tbody>${bodyRows}</tbody></table>`;
          } else {
            tplBody = tplDoc.body.innerHTML;
          }
        }
        appendedTemplateHtml = `
    <section class="templateAttach contractTemplateAttach">
      <div class="templateAttachHead">Contrato Anexo</div>
      <div class="templateAttachBody contractTemplateBody">${tplStyles}${tplBody}</div>
      __CONTRACT_SIGNATURES_BLOCK__
    </section>`;
      }
    } catch (err) {
      console.warn("No se pudo anexar Jardines:", err?.message || err);
    }
  }

  const menuMontajeAttachHtml = includeMenuMontaje ? buildQuoteMenuMontajeAttachmentHtml(ev, quote) : "";
  if (includeMenuMontaje && !menuMontajeAttachHtml) {
    await modernAlert({
      icon: "warning",
      title: "Menu y montaje no disponible",
      text: "Esta version necesita un Menu & Montaje guardado para poder imprimirse.",
    });
    return;
  }

  const quoteMetaRowsCompact = [
    ["Codigo", quote.code || ""],
    ["Fecha Doc", quote.docDate || ""],
    ["Encargado", manager?.name || ""],
    ["Contacto", quote.contact || manager?.name || ""],
    ["Email", quote.email || manager?.email || ""],
    ["Telefono", quote.phone || manager?.phone || ""],
    ["Institucion", quote.companyName || company?.name || ""],
    ["Facturar A", quote.billTo || company?.billTo || company?.businessName || ""],
    ["Nit", quote.nit || company?.nit || ""],
    ["Direccion", quote.address || company?.address || ""],
    ["Forma de Pago", quote.paymentType || ""],
    ["Fecha maxima pago", quote.dueDate || ""],
    ["Salon", quote.venue || ev.salon || ""],
    ["Horario", quote.schedule || `${ev.startTime} a ${ev.endTime}`],
    ["No Personas", String(quote.people || "")],
    ["Tipo de Evento", quote.eventType || ""],
    ["Folio No", quote.folio || ""],
    ["Fecha Inicio", quote.eventDate || ev.date || ""],
    ["Fecha Fin", quote.endDate || ev.date || ""],
  ];
  const quoteMetaTableHtml = showPrices ? `
    <table class="quoteMetaTable">
      <colgroup>
        <col class="metaLabel" />
        <col class="metaValue" />
        <col class="metaValue" />
        <col class="metaLabel" />
        <col class="metaValue" />
        <col class="metaValue" />
      </colgroup>
      <tbody>
        <tr>
          <td class="metaLabel">Codigo</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.code || "")}</td>
          <td class="metaLabel">Fecha Doc</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.docDate || "")}</td>
        </tr>
        <tr>
          <td class="metaLabel">Encargado</td>
          <td class="metaValue" colspan="2">${escapeHtml(manager?.name || "")}</td>
          <td class="metaLabel">Contacto</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.contact || manager?.name || "")}</td>
        </tr>
        <tr>
          <td class="metaLabel">Email</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.email || manager?.email || "")}</td>
          <td class="metaLabel">Telefono</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.phone || manager?.phone || "")}</td>
        </tr>
        <tr>
          <td class="metaLabel">Institucion</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.companyName || company?.name || "")}</td>
          <td class="metaLabel">Facturar A</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.billTo || company?.billTo || company?.businessName || "")}</td>
        </tr>
        <tr>
          <td class="metaLabel">Nit</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.nit || company?.nit || "")}</td>
          <td class="metaLabel">Direccion</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.address || company?.address || "")}</td>
        </tr>
        <tr class="metaDivider">
          <td class="metaLabel">Forma de Pago</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.paymentType || "")}</td>
          <td class="metaLabel">Fecha&nbsp;maximo<br>Pago</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.dueDate || "")}</td>
        </tr>
        <tr class="metaGap">
          <td colspan="6"></td>
        </tr>
        <tr>
          <td class="metaLabel">Salon</td>
          <td class="metaValue">${escapeHtml(quote.venue || ev.salon || "")}</td>
          <td class="metaLabel">Horario</td>
          <td class="metaValue metaWide">${escapeHtml(quote.schedule || `${ev.startTime} a ${ev.endTime}`)}</td>
          <td class="metaLabel">No Personas</td>
          <td class="metaValue">${escapeHtml(String(quote.people || ""))}</td>
        </tr>
        <tr>
          <td class="metaLabel">Tipo de Evento</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.eventType || "")}</td>
          <td class="metaLabel">Folio No</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.folio || "")}</td>
        </tr>
        <tr>
          <td class="metaLabel">Fecha Inicio</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.eventDate || ev.date || "")}</td>
          <td class="metaLabel">Fecha Fin</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.endDate || ev.date || "")}</td>
        </tr>
      </tbody>
    </table>` : `
    <table class="quoteMetaTableCompact">
      <tbody>
        ${quoteMetaRowsCompact.map(([label, value]) => `
          <tr>
            <td class="metaLabelCompact">${escapeHtml(String(label || ""))}</td>
            <td class="metaValueCompact">${escapeHtml(String(value || ""))}</td>
          </tr>`).join("")}
      </tbody>
    </table>`;
  const quoteItemsTableHtml = showPrices ? `
    <table>
      <thead>
        <tr>
          <th style="text-align:right">Cantidad</th>
          <th>Descripcion</th>
          <th style="text-align:right">Precio</th>
          <th style="text-align:right">Total</th>
        </tr>
      </thead>
      <tbody>
        ${rowHtml.join("")}
      </tbody>
      <tfoot>
        <tr>
          <td colspan="2"></td>
          <td class="sumLabel">SUBTOTAL EVENTO</td>
          <td class="sumValue">${quoteMoney(subtotalDoc)}</td>
        </tr>
        ${discountRowHtmlDoc}
        <tr class="sumTotal">
          <td colspan="2" class="sumWords">SON: ${escapeHtml(grandTotalWords)}</td>
          <td class="sumLabel">TOTAL EVENTO</td>
          <td class="sumValue">${quoteMoney(totalDoc)}</td>
        </tr>
      </tfoot>
    </table>` : `
    <table class="quoteItemsCompactTable">
      <thead>
        <tr>
          <th style="width:18%; text-align:right">Cantidad</th>
          <th>Descripcion</th>
        </tr>
      </thead>
      <tbody>
        ${orderedDates.map((d) => {
          const dayItems = grouped.get(d) || [];
          const dayRows = dayItems.map((item) => `
            <tr>
              <td style="text-align:right">${escapeHtml(String(Number(item.qty || 0)))}</td>
              <td>${escapeHtml(item.name || item.description || "")}</td>
            </tr>`).join("");
          return `
            <tr class="dayHeaderRow">
              <td colspan="2">SERVICIOS DEL DIA ${escapeHtml(formatDocDate(d))}</td>
            </tr>
            ${dayRows}`;
        }).join("")}
      </tbody>
    </table>`;
  const quoteInstitutionName = escapeHtml(String(quote.companyName || company?.name || "Cotizacion").trim() || "Cotizacion");
  const quoteDocLabel = `COTIZACION ${quoteInstitutionName}`;
  const quoteModeBadgeHtml = "";
  const quoteHeaderSrc = String(
    builtInTemplateMeta?.headerImage
    || (isServiHospTemplate ? SERVIHOSP_TEMPLATE_HEADER_IMAGE : DEFAULT_TEMPLATE_HEADER_IMAGE)
    || DEFAULT_TEMPLATE_HEADER_IMAGE
  ).trim();
  const sellerUser = normalizeUserRecord((state.users || []).find((u) => String(u.id || "") === String(ev?.userId || "")) || {});
  const sellerNameDoc = String(sellerUser?.fullName || sellerUser?.name || authSession.fullName || authSession.username || "").trim();
  const sellerEmailDoc = String(sellerUser?.email || "").trim();
  const sellerPhoneDoc = String(sellerUser?.phone || "").trim();
  const sellerSignatureDoc = String(sellerUser?.signatureDataUrl || authSession.signatureDataUrl || "").trim();
  const sellerSignatureHtmlDoc = sellerSignatureDoc
    ? `<img class="signatureImage" src="${escapeHtml(sellerSignatureDoc)}" alt="Firma vendedor" />`
    : "";
  const eventManagerNameDoc = String(quote?.contact || manager?.name || "").trim();
  const eventManagerEmailDoc = String(quote?.email || manager?.email || "").trim();
  const eventManagerPhoneDoc = String(quote?.phone || manager?.phone || "").trim();
  const contractSignaturesBlockHtml = `
    <section class="signatureSection contractTailSignatures">
      <div class="signatureGrid">
        <div class="signatureCard">
          <div class="signatureSignArea">${sellerSignatureHtmlDoc}</div>
          <div class="signatureLine"></div>
          <div class="signatureRole">Firma Vendedor</div>
          <div class="signatureData">${escapeHtml(sellerNameDoc || "-")}</div>
          <div class="signatureData">${escapeHtml(sellerEmailDoc || "-")}</div>
          <div class="signatureData">${escapeHtml(sellerPhoneDoc || "-")}</div>
        </div>
        <div class="signatureCard">
          <div class="signatureSignArea"></div>
          <div class="signatureLine"></div>
          <div class="signatureRole">Firma Encargado Evento</div>
          <div class="signatureData">${escapeHtml(eventManagerNameDoc || "-")}</div>
          <div class="signatureData">${escapeHtml(eventManagerEmailDoc || "-")}</div>
          <div class="signatureData">${escapeHtml(eventManagerPhoneDoc || "-")}</div>
        </div>
      </div>
    </section>`;
  if (appendedTemplateHtml && appendedTemplateHtml.includes("__CONTRACT_SIGNATURES_BLOCK__")) {
    appendedTemplateHtml = appendedTemplateHtml.replace("__CONTRACT_SIGNATURES_BLOCK__", contractSignaturesBlockHtml);
  }
  const html = `<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <base href="${escapeHtml(document.baseURI || window.location.href)}" />
  <title>Cotizacion ${escapeHtml(quote.code || "")}</title>
    <style>
    :root{
      --ink:#183447;
      --line:#d7e2e8;
      --line-strong:#b9cad4;
      --paper:#fbfdfd;
      --surface:#ffffff;
      --brand:#1ea7ad;
      --brand-deep:#117d8f;
      --brand-soft:#e8f8f8;
      --brand-muted:#d7efef;
      --navy:#1e3f57;
      --title:#16465e;
      --head-soft:#eff9fa;
      --zebra:#f6fbfc;
    }
    *{ box-sizing:border-box; }
    body{
      margin:0;
      padding:18px;
      font-family:"Segoe UI",Arial,sans-serif;
      font-size:12px;
      color:var(--ink);
      background:radial-gradient(circle at top left,#eaf7f8 0%,#eff4f8 40%,#edf2f6 100%);
    }
    .doc{
      max-width:1080px;
      margin:0 auto;
      border:1px solid var(--line);
      border-radius:18px;
      overflow:hidden;
      background:var(--surface);
      box-shadow:0 18px 42px rgba(17,47,66,0.14);
    }
    .quoteContractHeader{
      display:block;
      border-bottom:1px solid var(--line);
      background:#ffffff;
    }
    .quoteContractHeader img{
      display:block;
      width:100%;
      height:auto;
      max-height:92px;
      object-fit:contain;
      object-position:center top;
    }
    .head{
      background:var(--brand);
      color:#ffffff;
      padding:11px 16px;
      border-top:1px solid rgba(255,255,255,0.28);
      border-bottom:1px solid rgba(14,78,90,0.25);
      font-weight:800;
      font-size:13px;
      letter-spacing:0.42px;
      text-transform:uppercase;
      text-shadow:0 1px 0 rgba(9,48,56,0.2);
    }
    .quoteModeBadge{
      margin:0 16px 12px;
      padding:8px 12px;
      border:1px solid rgba(17,125,143,0.22);
      border-radius:12px;
      background:#e6f3f6;
      color:#11506a;
      font-size:11px;
      font-weight:900;
      letter-spacing:.28px;
      text-transform:uppercase;
    }
    .quoteMenuMontajeAttach{
      break-before:page;
      page-break-before:always;
    }
    .quoteMmWrap{
      display:grid;
      gap:14px;
      padding:2px;
      background:#f7fbfd;
    }
    .quoteMmDateSection{
      border:1px solid #cfe1ea;
      border-radius:14px;
      overflow:hidden;
      background:#ffffff;
      box-shadow:0 8px 18px rgba(15,23,42,.06);
    }
    .quoteMmSectionHead{
      display:flex;
      align-items:center;
      justify-content:space-between;
      gap:12px;
      padding:12px 14px;
      background:#1d6f88;
      color:#fff;
    }
    .quoteMmSectionTitle{
      font-size:14px;
      font-weight:900;
      letter-spacing:.24px;
      text-transform:uppercase;
    }
    .quoteMmSectionMeta{
      font-size:11px;
      font-weight:700;
      opacity:.92;
    }
    .quoteMmMetaGrid{
      display:grid;
      grid-template-columns:repeat(2, minmax(220px, 1fr));
      gap:8px 14px;
      padding:12px 14px;
      background:#eef7fb;
      border-bottom:1px solid #d6e6ee;
      font-size:12px;
      color:#1e3a4c;
    }
    .quoteMmCards{
      display:grid;
      gap:12px;
      padding:14px;
    }
    .quoteMmDayCard{
      border:1px solid #d6e4ea;
      border-radius:12px;
      overflow:hidden;
      background:#fff;
    }
    .quoteMmDayHead{
      display:flex;
      justify-content:space-between;
      gap:10px;
      align-items:center;
      padding:10px 12px;
      background:#eaf3f7;
      border-bottom:1px solid #d4e3ea;
    }
    .quoteMmDayTitle{
      font-size:13px;
      font-weight:900;
      color:#0f4864;
      letter-spacing:.2px;
      text-transform:uppercase;
    }
    .quoteMmDayMeta{
      font-size:11px;
      font-weight:700;
      color:#37607a;
    }
    .quoteMmBlock{
      padding:12px;
      border-top:1px solid #e1ebef;
    }
    .quoteMmBlock:first-of-type{
      border-top:none;
    }
    .quoteMmBlockTitle{
      margin:0 0 8px;
      font-size:12px;
      font-weight:900;
      letter-spacing:.24px;
      color:#0f4f6f;
      text-transform:uppercase;
    }
    .quoteMmBlockBody{
      font-size:12px;
      line-height:1.45;
      color:#1f3b4d;
    }
    .quoteMmEmpty{
      margin:0;
      color:#5d7487;
      font-style:italic;
    }
    .doc.quoteNoPrices tfoot,
    .doc.quoteNoPrices .financialPolicyTitle,
    .doc.quoteNoPrices .financialPolicyBox,
    .doc.quoteNoPrices .financialSummaryWrap{
      display:none;
    }
    .doc.quoteNoPrices .quoteContractHeader img{
      max-height:76px;
    }
    .doc.quoteNoPrices .signatureSection{
      margin-top:14px;
    }
    @media print{
      .quoteModeBadge{ margin-bottom:8px; }
      .quoteMenuMontajeAttach{ break-before:page; page-break-before:always; }
      .quoteMmDateSection{ break-inside:avoid; page-break-inside:avoid; }
    }
    .quoteMetaTable{
      width:100%;
      border-collapse:collapse;
      table-layout:fixed;
      margin:0 0 10px;
      border-top:3px solid rgba(30,167,173,0.32);
      border-left:1px solid var(--line);
      border-right:1px solid var(--line);
      background:var(--surface);
    }
    .quoteMetaTable col.metaLabel{ width:13%; }
    .quoteMetaTable col.metaValue{ width:20.333%; }
    .quoteMetaTable td{
      border-bottom:1px solid var(--line);
      border-right:1px solid var(--line);
      padding:6px 8px;
      font-size:12px;
      color:#274152;
      vertical-align:middle;
      background:var(--surface);
    }
    .quoteMetaTable tr td:last-child{ border-right:none; }
    .quoteMetaTable td.metaLabel{ background:#e7f5f6; color:var(--title); font-weight:700; letter-spacing:0.14px; }
    .quoteMetaTable td.metaValue{ font-weight:600; background:var(--surface); }
    .quoteMetaTable tr.metaDivider td{ border-top:2px solid rgba(30,167,173,0.35); }
    .quoteMetaTable tr.metaGap td{ padding:0; height:9px; border:none; background:#f4fafc; }
    .quoteMetaTable td.metaWide{ text-align:left; }
    .quoteMetaTableCompact{
      width:100%;
      border-collapse:collapse;
      table-layout:fixed;
      margin:0 0 10px;
      border-top:3px solid rgba(30,167,173,0.32);
      border-left:1px solid var(--line);
      border-right:1px solid var(--line);
      background:var(--surface);
    }
    .quoteMetaTableCompact td{
      border-bottom:1px solid var(--line);
      border-right:1px solid var(--line);
      padding:8px 10px;
      font-size:12px;
      color:#274152;
      vertical-align:middle;
      background:var(--surface);
    }
    .quoteMetaTableCompact tr td:last-child{ border-right:none; }
    .quoteMetaTableCompact .metaLabelCompact{
      width:24%;
      background:#e7f5f6;
      color:var(--title);
      font-weight:700;
      letter-spacing:0.14px;
    }
    .quoteMetaTableCompact .metaValueCompact{
      font-weight:600;
      background:var(--surface);
    }
    .quoteItemsCompactTable thead th:first-child,
    .quoteItemsCompactTable tbody td:first-child{
      width:18%;
    }
    table{ width:100%; border-collapse:collapse; }
    .quoteItemsTablePrint{
      table-layout:fixed;
    }
    .quoteItemsTablePrint col.qtyCol{ width:11%; }
    .quoteItemsTablePrint col.descCol{ width:57%; }
    .quoteItemsTablePrint col.priceCol{ width:16%; }
    .quoteItemsTablePrint col.totalCol{ width:16%; }
    thead th{
      background:#e5f5f6;
      border:1px solid #cddde5;
      padding:6px 8px;
      text-align:center;
      font-size:11.5px;
      font-weight:800;
      color:var(--title);
      text-transform:uppercase;
      letter-spacing:0.2px;
    }
    .quoteItemsTablePrint thead th:nth-child(1){ text-align:center; }
    .quoteItemsTablePrint thead th:nth-child(2){ text-align:center; }
    .quoteItemsTablePrint thead th:nth-child(3),
    .quoteItemsTablePrint thead th:nth-child(4){ text-align:center; }
    tbody td{ border:1px solid var(--line); padding:7px 9px; font-size:12px; background:var(--surface); color:#1f3b4d; }
    .quoteItemsTablePrint tbody td:nth-child(1){
      text-align:center;
      padding-right:16px;
      border-right:2px solid #c7d8e2;
      font-variant-numeric:tabular-nums;
    }
    .quoteItemsTablePrint tbody td:nth-child(2){
      text-align:left;
      padding-left:14px;
    }
    .quoteItemsTablePrint tbody td:nth-child(3),
    .quoteItemsTablePrint tbody td:nth-child(4){
      text-align:right;
      padding-right:12px;
      font-variant-numeric:tabular-nums;
    }
    tbody tr:nth-child(even):not(.dayHeaderRow):not(.daySubtotalRow) td{ background:var(--zebra); }
    .dayHeaderRow td{
      background:#215d78;
      color:#ffffff !important;
      font-weight:800;
      letter-spacing:0.32px;
      text-transform:uppercase;
      text-align:center;
      border:1px solid rgba(16,64,79,0.24);
      padding:6px 9px;
      box-shadow:inset 0 1px 0 rgba(255,255,255,0.15);
    }
    .daySubtotalRow td{
      border-top:2px solid rgba(17,125,143,0.42);
      border-bottom:1px solid #cfe1e7;
      background:var(--brand-muted) !important;
      color:#15516b;
      font-weight:700;
    }
    tfoot td{ border:1px solid #c8d7e0; padding:6px 8px; font-size:11px; background:#e9f5f7; color:#18445a; }
    .sumLabel{ text-align:right; font-weight:700; color:#15495f; background:#e2f1f3; }
    .sumValue{ text-align:right; font-weight:800; color:#103b52; background:#e2f1f3; font-variant-numeric:tabular-nums; }
    .sumTotal .sumLabel,
    .sumTotal .sumValue{ background:var(--brand-deep); color:#ffffff; font-size:12.2px; border-color:rgba(16,77,89,0.4); letter-spacing:0.14px; }
    .sumWords{ border-top:1px solid #c7dde4; background:#e0f2f3; color:#16536d; font-weight:700; font-size:12px; text-transform:uppercase; line-height:1.35; }
    .notes{ border-top:1px solid var(--line); padding:11px 13px; min-height:56px; background:var(--paper); }
    .notes b{ display:block; margin-bottom:6px; }
    .policyTitle{
      margin-top:9px;
      background:#1f5f7a;
      color:#fff;
      font-weight:800;
      text-align:center;
      padding:4px 8px;
      border:1px solid rgba(18,70,86,0.24);
      letter-spacing:.24px;
      text-transform:uppercase;
    }
    .policyBox{ border:1px solid #b5cad8; border-top:none; padding:7px 9px; font-size:10.4px; line-height:1.4; background:#f9fcfd; }
    .policyBox p{ margin:0 0 6px; }
    .policyBox p:last-child{ margin-bottom:0; }
    .cargoSummaryWrap{ margin:0 0 11px; border:1px solid #c4d7e3; border-top:none; border-radius:0 0 12px 12px; overflow:hidden; background:var(--surface); box-shadow:inset 0 1px 0 rgba(255,255,255,0.72); }
    .cargoTable{ width:100%; border-collapse:collapse; table-layout:fixed; font-size:12px; }
    .cargoTable col:first-child{ width:72%; }
    .cargoTable col:last-child{ width:28%; }
    .cargoTable th,.cargoTable td{ padding:9px 10px; border-bottom:1px solid var(--line); }
    .cargoTable thead th{ background:#e6f4f6; color:var(--title); font-size:11px; font-weight:800; letter-spacing:.24px; text-transform:uppercase; }
    .cargoTable thead th:first-child{ text-align:left; border-right:1px solid var(--line); }
    .cargoTable thead th:last-child{ text-align:right; }
    .cargoTable tbody td:first-child{ border-right:1px solid var(--line); }
    .cargoTable tbody tr:last-child td{ border-bottom:none; }
    .cargoTable tbody tr:nth-child(odd):not(.cargoEm):not(.cargoEmFinal){ background:#fcfefe; }
    .cargoTable tbody tr:nth-child(even):not(.cargoEm):not(.cargoEmFinal){ background:#f3f9fb; }
    .cargoLabel{ font-weight:700; color:#1c3f52; }
    .cargoAmount{ text-align:right; white-space:nowrap; font-weight:800; color:#163f55; font-variant-numeric:tabular-nums; }
    .cargoEm{ font-weight:800; background:#e1f2f4 !important; }
    .cargoEm td{ border-bottom-color:#c7dbe4; }
    .cargoEmFinal{ background:var(--brand-deep) !important; }
    .cargoEm .cargoLabel,
    .cargoEm .cargoAmount{ color:#14546e; }
    .cargoEmFinal .cargoLabel,
    .cargoEmFinal .cargoAmount{ color:#ffffff; }
    .actions{ padding:12px; display:flex; justify-content:flex-end; gap:8px; border-top:1px solid var(--line); background:#f0f7fa; }
    .actions button{ border:1px solid rgba(17,125,143,0.32); background:#ffffff; color:#0f4760; border-radius:10px; padding:7px 12px; font-weight:700; cursor:pointer; }
    .actions button:hover{ background:#e9f7f8; }
    .templateAttach{ border-top:1px solid var(--line); background:#f3f9fc; padding:12px 12px 0; }
    .templateAttachHead{ font-weight:800; color:#124c67; margin:0 0 8px; font-size:12.5px; letter-spacing:0.12px; text-transform:uppercase; }
    .templateAttachBody{ background:var(--surface); border:1px solid var(--line); border-radius:12px; overflow:visible; font-size:10px; line-height:1.15; page-break-inside:auto; }
    .signatureSection{
      margin:10px 12px 0;
      padding:14px 10px 8px;
      border:1px solid #c7dbe2;
      border-radius:12px;
      background:#ffffff;
    }
    .contractTailSignatures{
      margin:12px;
      border-top:2px solid #d7e1ea;
      padding-top:14px;
    }
    .signatureGrid{
      display:grid;
      grid-template-columns:1fr 1fr;
      gap:24px;
    }
    .signatureSignArea{
      width:100%;
      min-height:72px;
      display:flex;
      align-items:flex-end;
      justify-content:center;
      margin-bottom:8px;
      background:#ffffff;
      border:1px solid #d7e1ea;
      border-radius:6px;
    }
    .signatureImage{
      max-height:68px;
      width:auto;
      max-width:320px;
      object-fit:contain;
      object-position:center bottom;
      display:block;
      background:#ffffff;
      padding:2px 6px;
    }
    .signatureCard{
      padding:6px 10px;
      min-height:120px;
      display:flex;
      flex-direction:column;
      align-items:center;
      text-align:center;
    }
    .signatureLine{
      width:78%;
      max-width:320px;
      border-top:2px solid #6b8798;
      margin:0 auto;
    }
    .signatureRole{
      font-weight:800;
      color:#164d66;
      text-transform:uppercase;
      letter-spacing:.2px;
      font-size:11px;
      margin-bottom:6px;
      text-align:center;
    }
    .signatureData{
      margin:2px 0;
      font-size:11px;
      color:#2a4b5e;
      line-height:1.35;
      text-align:center;
      width:100%;
    }
    @media print{
      .signatureSection{ break-inside:avoid; }
    }
    .templateAttachBody p,
    .templateAttachBody td,
    .templateAttachBody th,
    .templateAttachBody span,
    .templateAttachBody div{ font-size:inherit !important; }
    .doc.serviHospClean,
    .doc.serviHospClean *{ box-shadow:none !important; }
    .doc.serviHospClean,
     .doc.serviHospClean .quoteContractHeader,
    
    .doc.serviHospClean .head,
    .doc.serviHospClean .cell,
    .doc.serviHospClean .k,
    .doc.serviHospClean thead th,
    .doc.serviHospClean tbody td,
    .doc.serviHospClean tfoot td,
    .doc.serviHospClean .dayHeaderRow td,
    .doc.serviHospClean .daySubtotalRow td,
    .doc.serviHospClean .notes,
    .doc.serviHospClean .policyTitle,
    .doc.serviHospClean .policyBox,
    .doc.serviHospClean .cargoSummaryWrap,
    .doc.serviHospClean .cargoTable th,
    .doc.serviHospClean .cargoTable td,
    .doc.serviHospClean .actions,
    .doc.serviHospClean .templateAttach,
    .doc.serviHospClean .templateAttachBody,
    .doc.serviHospClean .templateAttachBody table,
    .doc.serviHospClean .templateAttachBody td,
    .doc.serviHospClean .templateAttachBody th{ border:none !important; border-radius:0 !important; }
    .doc.serviHospClean,
     .doc.serviHospClean .quoteContractHeader,
    
    .doc.serviHospClean .templateAttach,
    .doc.serviHospClean .templateAttachBody{ background:#fff !important; }
    .doc.serviHospClean .templateAttach{ padding-top:4px; }
    .doc.serviHospClean .contractPrintFoot,
    .doc.serviHospClean .contract-footer,
    .doc.serviHospClean .contract-footer-print{ display:none !important; }
    @media print{
      html, body{ -webkit-print-color-adjust:exact; print-color-adjust:exact; }
      .actions{ display:none; }
      body{ padding:0; background:#ffffff; }
      .doc{ border:none; box-shadow:none; }
      .templateAttach{ padding:0; background:transparent; border-top:none; }
      .templateAttachHead{ display:none; }
      .templateAttachBody{ border:none; border-radius:0; overflow:visible; }
    }
  </style>
</head>
<body>
  <div class="doc${isServiHospTemplate ? " serviHospClean" : ""}${showPrices ? "" : " quoteNoPrices"}">
    <div class="quoteContractHeader">
      <img src="${escapeHtml(quoteHeaderSrc)}" alt="Encabezado contrato" />
    </div>
    <div class="head">${quoteDocLabel}</div>
    ${quoteModeBadgeHtml}
    <table class="quoteMetaTable">
      <colgroup>
        <col class="metaLabel" />
        <col class="metaValue" />
        <col class="metaValue" />
        <col class="metaLabel" />
        <col class="metaValue" />
        <col class="metaValue" />
      </colgroup>
      <tbody>
        <tr>
          <td class="metaLabel">Codigo</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.code || "")}</td>
          <td class="metaLabel">Fecha Doc</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.docDate || "")}</td>
        </tr>
        <tr>
          <td class="metaLabel">Encargado</td>
          <td class="metaValue" colspan="2">${escapeHtml(manager?.name || "")}</td>
          <td class="metaLabel">Contacto</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.contact || manager?.name || "")}</td>
        </tr>
        <tr>
          <td class="metaLabel">Email</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.email || manager?.email || "")}</td>
          <td class="metaLabel">Telefono</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.phone || manager?.phone || "")}</td>
        </tr>
        <tr>
          <td class="metaLabel">Institucion</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.companyName || company?.name || "")}</td>
          <td class="metaLabel">Facturar A</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.billTo || company?.billTo || company?.businessName || "")}</td>
        </tr>
        <tr>
          <td class="metaLabel">Nit</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.nit || company?.nit || "")}</td>
          <td class="metaLabel">Direccion</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.address || company?.address || "")}</td>
        </tr>
        <tr class="metaDivider">
          <td class="metaLabel">Forma de Pago</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.paymentType || "")}</td>
          <td class="metaLabel">Fecha&nbsp;maximo<br>Pago</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.dueDate || "")}</td>
        </tr>
        <tr class="metaGap">
          <td colspan="6"></td>
        </tr>
        <tr>
          <td class="metaLabel">Salon</td>
          <td class="metaValue">${escapeHtml(quote.venue || ev.salon || "")}</td>
          <td class="metaLabel">Horario</td>
          <td class="metaValue metaWide">${escapeHtml(quote.schedule || `${ev.startTime} a ${ev.endTime}`)}</td>
          <td class="metaLabel">No Personas</td>
          <td class="metaValue">${escapeHtml(String(quote.people || ""))}</td>
        </tr>
        <tr>
          <td class="metaLabel">Tipo de Evento</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.eventType || "")}</td>
          <td class="metaLabel">Folio No</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.folio || "")}</td>
        </tr>
        <tr>
          <td class="metaLabel">Fecha Inicio</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.eventDate || ev.date || "")}</td>
          <td class="metaLabel">Fecha Fin</td>
          <td class="metaValue" colspan="2">${escapeHtml(quote.endDate || ev.date || "")}</td>
        </tr>
      </tbody>
    </table>

    <table>
      <thead>
        <tr>
          <th style="text-align:right">Cantidad</th>
          <th>Descripcion</th>
          <th style="text-align:right">Precio</th>
          <th style="text-align:right">Total</th>
        </tr>
      </thead>
      <tbody>
        ${rowHtml.join("")}
      </tbody>
      <tfoot>
        <tr>
          <td colspan="2"></td>
          <td class="sumLabel">SUBTOTAL EVENTO</td>
          <td class="sumValue">${quoteMoney(subtotalDoc)}</td>
        </tr>
        ${discountRowHtmlDoc}
        <tr class="sumTotal">
          <td colspan="2" class="sumWords">SON: ${escapeHtml(grandTotalWords)}</td>
          <td class="sumLabel">TOTAL EVENTO</td>
          <td class="sumValue">${quoteMoney(totalDoc)}</td>
        </tr>
      </tfoot>
    </table>

    <div class="policyTitle financialPolicyTitle">NOTAS</div>
    <div class="policyBox financialPolicyBox">
      <p>Incrementos en menos de 24 horas de servicios y/o productos tendran cargo adicional del 10% en relacion con el excedente.</p>
      <p>El monto de anticipo para asegurar el evento depende de las clausulas de formalizacion.</p>
      <p>Estamos sujetos a pagos trimestrales. Aplicamos Clausula de No Show.</p>
      <p>Esta propuesta se formaliza al tener la firma del asesor de ventas y del cliente, junto a comprobante de anticipo y/o orden de compra.</p>
      <p>No se reembolsa el anticipo si no realiza su evento por cualquier causa.</p>
    </div>

    <div class="policyTitle financialPolicyTitle">RESUMEN DE CARGOS</div>
    <div class="cargoSummaryWrap financialSummaryWrap">
      <table class="cargoTable"><colgroup><col><col></colgroup>
        <thead>
          <tr>
            <th style="width:68%">Descripcion</th>
            <th>Total</th>
          </tr>
        </thead>
        <tbody>
          ${cargoRowsDoc.map((r) => `
            <tr>
              <td class="cargoLabel">${escapeHtml(r.label)}</td>
              <td class="cargoAmount">${Number(r.amount || 0) > 0 ? escapeHtml(moneyQDoc(r.amount)) : "-"}</td>
            </tr>
          `).join("")}
          <tr class="cargoEm">
            <td class="cargoLabel">Total contratado</td>
            <td class="cargoAmount">${totalContratadoDoc > 0 ? escapeHtml(moneyQDoc(totalContratadoDoc)) : "-"}</td>
          </tr>
          ${advanceRowsDoc.length ? advanceRowsDoc.map((r) => `
            <tr>
              <td class="cargoLabel">${escapeHtml(`${formatDocDate(r.date)} - ${formatQuoteAdvanceDetail(r)}`)}</td>
              <td class="cargoAmount">${Number(r.amount || 0) > 0 ? escapeHtml(moneyQDoc(r.amount)) : "-"}</td>
            </tr>
          `).join("") : `
            <tr>
              <td class="cargoLabel">Anticipos: sin registros</td>
              <td class="cargoAmount">-</td>
            </tr>
          `}
          <tr class="cargoEm">
            <td class="cargoLabel">Total anticipos</td>
            <td class="cargoAmount">${totalAnticiposDoc > 0 ? escapeHtml(moneyQDoc(totalAnticiposDoc)) : "-"}</td>
          </tr>
          <tr class="cargoEm cargoEmFinal">
            <td class="cargoLabel">Saldo</td>
            <td class="cargoAmount">${saldoDoc > 0 ? escapeHtml(moneyQDoc(saldoDoc)) : "-"}</td>
          </tr>
        </tbody>
      </table>
    </div>

    <section class="signatureSection">
      <div class="signatureGrid">
        <div class="signatureCard">
          <div class="signatureSignArea">${sellerSignatureHtmlDoc}</div>
          <div class="signatureLine"></div>
          <div class="signatureRole">Firma Vendedor</div>
          <div class="signatureData">${escapeHtml(sellerNameDoc || "-")}</div>
          <div class="signatureData">${escapeHtml(sellerEmailDoc || "-")}</div>
          <div class="signatureData">${escapeHtml(sellerPhoneDoc || "-")}</div>
        </div>
        <div class="signatureCard">
          <div class="signatureSignArea"></div>
          <div class="signatureLine"></div>
          <div class="signatureRole">Firma Encargado Evento</div>
          <div class="signatureData">${escapeHtml(eventManagerNameDoc || "-")}</div>
          <div class="signatureData">${escapeHtml(eventManagerEmailDoc || "-")}</div>
          <div class="signatureData">${escapeHtml(eventManagerPhoneDoc || "-")}</div>
        </div>
      </div>
    </section>
    ${menuMontajeAttachHtml}
    ${appendedTemplateHtml}
    <div class="actions"> 
      <button onclick="window.print()">Imprimir</button>
    </div>
  </div>
</body>
</html>`;

  const win = window.open("", "_blank");
  if (!win) {
    toast("Cotizacion guardada. Habilita ventanas emergentes para ver el documento.");
    return;
  }
  win.document.open();
  win.document.write(html);
  win.document.close();
}

function saveEventFromForm() {
  if (!serverStateReady) {
    pendingPersistAfterSync = true;
    syncWithServerState();
    return toast("Espera: cargando datos desde MariaDB.");
  }

  const editingId = el.eventId.value;
  const dateStart = el.eventDate.value;
  const dateEnd = el.eventDateEnd.value || dateStart;
  const rangeStart = dateStart <= dateEnd ? dateStart : dateEnd;
  const rangeEnd = dateStart <= dateEnd ? dateEnd : dateStart;
  const normalizedName = String(el.eventName?.value || "").replace(/\s+/g, " ").trim();
  const normalizedNotes = String(el.eventNotes?.value || "").replace(/\s+/g, " ").trim();
  const hasVisibleChars = (value) => String(value || "").replace(/[\s\u200B-\u200D\uFEFF]/g, "").length > 0;
  const name = normalizedName;
  const rawStatus = String(el.eventStatus.value || "").trim();
  const editingEvent = editingId ? state.events.find(x => x.id === editingId) : null;
  if (editingEvent && isEventSeriesInPast(editingEvent) && !hasPastEventEditAuthorization(editingEvent)) {
    return toast("Evento de fecha pasada bloqueado. Solicita codigo de administrador.");
  }
  const editingCurrentStatus = String(editingEvent?.status || "").trim();
  let status = isAutoStatus(rawStatus)
    ? (editingId ? (editingCurrentStatus || STATUS.RESERVA_SIN_COTIZACION) : STATUS.RESERVA_SIN_COTIZACION)
    : (rawStatus || STATUS.RESERVA_SIN_COTIZACION);
  const userId = el.eventUser.value;
  const notes = normalizedNotes;
  const slots = getSlotsFromForm();
  const pax = syncEventPaxFromSlots() || null;
  const allowedStatuses = new Set(Object.values(STATUS));
  const selectedUser = (state.users || []).find((u) => String(u?.id || "") === String(userId || ""));
  const salonCapacityMap = state?.salonCapacities && typeof state.salonCapacities === "object"
    ? state.salonCapacities
    : null;

  if (!name || !hasVisibleChars(name)) return toast("Falta nombre del evento.");
  if (el.eventName) el.eventName.value = name;
  if (String(el.eventNotes?.value || "") !== notes && el.eventNotes) el.eventNotes.value = notes;
  if (!rangeStart || !rangeEnd) return toast("Falta la fecha.");
  if (!userId) return toast("Selecciona un vendedor.");
  if (!selectedUser || selectedUser.active === false) return toast("Selecciona un usuario valido y activo.");
  if (!status || !allowedStatuses.has(status)) return toast("Estado invalido. Selecciona un estado permitido.");
  if (!pax) return toast("Ingresa una cantidad valida de personas por salon.");
  if (!slots.length) return toast("Agrega al menos un bloque de salon/horario.");
  for (const s of slots) {
    const slotRange = normalizeSlotDateRange(s, rangeStart, rangeEnd);
    s.dateStart = slotRange.start;
    s.dateEnd = slotRange.end;
    if (!s.salon || !s.dateStart || !s.dateEnd || !s.startTime || !s.endTime) {
      return toast("Completa salon, fechas, inicio y fin en cada bloque.");
    }
    if (!s.slotPax || Number(s.slotPax) <= 0) return toast("Completa un PAX valido en cada salon.");
    const slotPax = Number(s.slotPax || 0);
    const configuredMax = salonCapacityMap
      ? Number(salonCapacityMap[String(s.salon || "").trim()] || 0)
      : 0;
    const maxAllowedPax = configuredMax > 0 ? configuredMax : FALLBACK_MAX_PAX_PER_SLOT;
    if (slotPax > maxAllowedPax) {
      return toast(`El PAX del salon ${s.salon} supera el maximo permitido (${maxAllowedPax}).`);
    }
    if (s.dateStart < rangeStart || s.dateEnd > rangeEnd) {
      return toast("Las fechas de cada salon deben estar dentro de la fecha inicial y final del evento.");
    }
    if (!isValidClockTime(s.startTime) || !isValidClockTime(s.endTime)) {
      return toast("Formato de hora invalido. Usa HH:mm.");
    }
    if (compareTime(s.endTime, s.startTime) <= 0) return toast("En cada bloque, la hora final debe ser mayor que inicio.");
  }

  const replaceEvents = editingId
    ? getEventSeries(state.events.find(x => x.id === editingId))
    : [];
  const previousStatus = String(replaceEvents[0]?.status || "").trim();
  if (editingId && previousStatus === STATUS.PERDIDO && rangeEnd >= toISODate(new Date())) {
    status = reservationStatusFromQuotePresence(replaceEvents[0] || editingEvent);
  }
  const oldBlockingWindows = buildBlockingWindowsFromEvents(replaceEvents);
  const oldKey = replaceEvents[0] ? reservationKeyFromEvent(replaceEvents[0]) : "";
  const oldSnapshot = buildSeriesSnapshot(replaceEvents);
  const replaceIds = new Set(replaceEvents.map(e => e.id));
  const currentGroupId = replaceEvents[0]?.groupId || null;
  const slotPlans = slots.map((slot) => ({
    ...slot,
    targetDates: listDatesBetween(slot.dateStart, slot.dateEnd),
  }));
  const targetDates = Array.from(new Set(slotPlans.flatMap((slot) => slot.targetDates))).sort();
  const needsGroup = targetDates.length > 1 || slots.length > 1 || replaceEvents.length > 1;
  const resultingGroupId = needsGroup ? (currentGroupId || `grp_${uid()}`) : null;
  const existingByKey = new Map(replaceEvents.map(e => [`${e.date}|${e.salon}|${e.startTime}|${e.endTime}`, e]));

  const principalSalon = String(slots[0]?.salon || "").trim();
  const reservationSalones = Array.from(new Set(slots.map((slot) => String(slot?.salon || "").trim()).filter(Boolean)));
  const drafts = [];
  for (const s of slotPlans) {
    for (const d of s.targetDates) {
      const key = `${d}|${s.salon}|${s.startTime}|${s.endTime}`;
      const existing = existingByKey.get(key);
      drafts.push({
        id: existing?.id || uid(),
        name,
        salon: s.salon,
        mainSalon: String(existing?.mainSalon || principalSalon || s.salon || "").trim(),
        salones: reservationSalones.slice(),
        date: d,
        eventDateStart: rangeStart,
        eventDateEnd: rangeEnd,
        groupId: resultingGroupId,
        status,
        startTime: s.startTime,
        endTime: s.endTime,
        userId,
        pax,
        slotPax: s.slotPax,
        notes,
        quote: existing?.quote,
      });
    }
  }

  for (let i = 0; i < drafts.length; i++) {
    for (let j = i + 1; j < drafts.length; j++) {
      const a = drafts[i];
      const b = drafts[j];
      if (a.date !== b.date || a.salon !== b.salon) continue;
      if (timesOverlap(a.startTime, a.endTime, b.startTime, b.endTime)) {
        return toast("Hay bloques que se sobreponen en el mismo salon.");
      }
    }
  }

  for (const draft of drafts) {
    const rules = evaluateRules(draft, replaceIds);
    if (!rules.ok) {
      toast(rules.message);
      return;
    }
  }

  if (status === STATUS.MANTENIMIENTO) {
    const maintenanceHardBlocks = drafts.flatMap((draft) => findHardBlocks(draft, replaceIds));
    if (maintenanceHardBlocks.length) {
      return toast("No puedes poner en Mantenimiento: ya hay una reserva Confirmada o Pre reserva en ese horario.");
    }
  }

  if (editingId) {
    const newSnapshot = buildDraftSnapshot(drafts);
    const oldComparable = oldSnapshot ? JSON.stringify(oldSnapshot) : "";
    const newComparable = newSnapshot ? JSON.stringify(newSnapshot) : "";
    if (oldComparable && newComparable && oldComparable === newComparable) {
      return toast("Sin cambios detectados en la reserva.");
    }
  }

  state.events = state.events.filter(x => !replaceIds.has(x.id));
  state.events.push(...drafts);
  const newKey = resultingGroupId || drafts[0]?.id || oldKey;
  const newBlockingWindows = buildBlockingWindowsFromEvents(drafts);
  const releasedWindows = computeReleasedWindows(oldBlockingWindows, newBlockingWindows);
  moveHistoryKey(oldKey, newKey);
  moveReminderKey(oldKey, newKey);
  if (editingId) {
    const newSnapshot = buildDraftSnapshot(drafts);
    appendDetailedEditHistory(newKey, userId, oldSnapshot, newSnapshot);
  } else {
    appendHistoryByKey(newKey, userId, `Reserva creada: ${summarizeDraftWindow(drafts)}.`);
  }

  if (editingId && releasedWindows.length) {
    notifyReleasedCapacityForWaitingReservations({
      releasedWindows,
      sourceEvent: drafts[0] || replaceEvents[0] || null,
      reasonText: "dejo de bloquear ese horario",
      actorUserId: userId || "",
    });
  }

  persist();
  autoMarkLostEvents(); // in case
  interaction.suppressClickUntil = Date.now() + 450;
  closeModal();
  render();
  requestAnimationFrame(() => {
    render();
  });
  setTimeout(() => {
    render();
  }, 120);
  const nextStatus = String(drafts[0]?.status || "").trim();
  const statusToast = buildStatusChangeToast(previousStatus, nextStatus, "Guardado con exito.");
  toast(`${statusToast} Guardado: ${drafts.length} bloque(s).`);
}

function removeEvent(id, actorUserId = "") {
  const ev = state.events.find(x => x.id === id);
  if (!ev) return;
  const series = getEventSeries(ev);
  const releasedWindows = buildBlockingWindowsFromEvents(series);
  const key = reservationKeyFromEvent(ev);
  const summary = summarizeSeriesWindow(series);
  appendHistoryByKey(key, actorUserId || ev.userId, `Reserva eliminada (${summary}).`);
  const removeIds = new Set(series.map(x => x.id));
  state.events = state.events.filter(x => !removeIds.has(x.id));
  ensureReminderStore();
  delete state.reminders[key];
  if (releasedWindows.length) {
    notifyReleasedCapacityForWaitingReservations({
      releasedWindows,
      sourceEvent: ev,
      reasonText: "fue eliminada",
      actorUserId: actorUserId || ev.userId || "",
    });
  }
  persist();
  render();
}

function getEventSeries(ev) {
  if (!ev) return [];
  if (!ev.groupId) return [ev];
  return state.events.filter(x => x.groupId === ev.groupId);
}

function getEventSeriesFinancialMeta(ev) {
  const series = getEventSeries(ev)
    .slice()
    .sort((a, b) => {
      const byDate = String(a?.date || "").localeCompare(String(b?.date || ""));
      if (byDate !== 0) return byDate;
      const byStart = String(a?.startTime || "").localeCompare(String(b?.startTime || ""));
      if (byStart !== 0) return byStart;
      return String(a?.salon || "").localeCompare(String(b?.salon || ""));
    });
  const salones = [];
  for (const item of series) {
    const eventSalones = Array.isArray(item?.salones) ? item.salones : [];
    for (const salon of [...eventSalones, item?.salon]) {
      const label = String(salon || "").trim();
      if (label) salones.push(label);
    }
  }
  const uniqueSalones = Array.from(new Set(salones));
  const explicitMainSalon = series.map((item) => String(item?.mainSalon || "").trim()).find(Boolean) || "";
  const mainSalon = explicitMainSalon || uniqueSalones[0] || String(ev?.salon || "").trim();
  const primaryEvent = series.find((item) => String(item?.salon || "").trim() === mainSalon)
    || series.find((item) => String(item?.id || "").trim() === String(ev?.id || "").trim())
    || series[0]
    || ev
    || null;
  const firstEvent = series[0] || ev || null;
  const lastEvent = series[series.length - 1] || ev || null;
  const eventDateStart = getEventGeneralDateStart(firstEvent) || String(firstEvent?.date || "").trim();
  const eventDateEnd = getEventGeneralDateEnd(firstEvent) || String(lastEvent?.date || "").trim();
  return {
    series,
    salones: uniqueSalones,
    mainSalon,
    primaryEvent,
    startDate: eventDateStart,
    endDate: eventDateEnd,
    startTime: String(primaryEvent?.startTime || firstEvent?.startTime || "").trim(),
    endTime: String(primaryEvent?.endTime || firstEvent?.endTime || "").trim(),
  };
}

function getUsedSalonesForReports() {
  const seenReservations = new Set();
  const salones = new Set();
  for (const ev of state.events || []) {
    const reservationKey = String(reservationKeyFromEvent(ev) || ev?.id || "").trim();
    if (reservationKey) {
      if (seenReservations.has(reservationKey)) continue;
      seenReservations.add(reservationKey);
    }
    for (const salon of getEventSeriesFinancialMeta(ev).salones) {
      salones.add(salon);
    }
  }
  return Array.from(salones).sort((a, b) => String(a).localeCompare(String(b), "es", { sensitivity: "base" }));
}

function uniqueSlotsFromSeries(series) {
  const map = new Map();
  for (const e of series) {
    const key = `${e.salon}|${e.startTime}|${e.endTime}`;
    if (!map.has(key)) {
      map.set(key, {
        salon: e.salon,
        slotPax: Math.max(0, Number(e?.slotPax || 0)) || 0,
        dateStart: e.date || "",
        dateEnd: e.date || "",
        startTime: e.startTime,
        endTime: e.endTime,
      });
    } else {
      const slot = map.get(key);
      const date = String(e?.date || "").trim();
      if (date) {
        if (!slot.dateStart || date < slot.dateStart) slot.dateStart = date;
        if (!slot.dateEnd || date > slot.dateEnd) slot.dateEnd = date;
      }
    }
  }
  const slots = Array.from(map.values()).sort((a, b) => compareTime(a.startTime, b.startTime));
  const storedTotalPax = Math.max(0, Number(series?.[0]?.pax || 0)) || 0;
  const hasSlotPax = slots.some((slot) => Math.max(0, Number(slot?.slotPax || 0)) > 0);
  if (!hasSlotPax && storedTotalPax > 0 && slots.length) {
    const basePax = Math.floor(storedTotalPax / slots.length);
    let remainder = storedTotalPax % slots.length;
    for (const slot of slots) {
      slot.slotPax = basePax + (remainder > 0 ? 1 : 0);
      if (remainder > 0) remainder -= 1;
    }
  }
  return slots;
}

function summarizeSeriesWindow(series) {
  if (!series.length) return "";
  const sorted = series.slice().sort((a, b) => a.date.localeCompare(b.date));
  const first = getEventGeneralDateStart(sorted[0]) || sorted[0].date;
  const last = getEventGeneralDateEnd(sorted[0]) || sorted[sorted.length - 1].date;
  const slots = uniqueSlotsFromSeries(sorted);
  const slotText = slots.map(s => {
    const dates = s.dateStart && s.dateEnd && (s.dateStart !== first || s.dateEnd !== last)
      ? ` ${s.dateStart}${s.dateStart !== s.dateEnd ? ` a ${s.dateEnd}` : ""}`
      : "";
    return `${s.salon}${dates} ${s.startTime}-${s.endTime}`;
  }).join(", ");
  return `${first}${first !== last ? ` a ${last}` : ""}${slotText ? ` (${slotText})` : ""}`;
}

function summarizeDraftWindow(drafts) {
  if (!drafts.length) return "";
  const sorted = drafts.slice().sort((a, b) => a.date.localeCompare(b.date));
  const first = getEventGeneralDateStart(sorted[0]) || sorted[0].date;
  const last = getEventGeneralDateEnd(sorted[0]) || sorted[sorted.length - 1].date;
  const slots = uniqueSlotsFromSeries(sorted);
  const slotText = slots.map(s => {
    const dates = s.dateStart && s.dateEnd && (s.dateStart !== first || s.dateEnd !== last)
      ? ` ${s.dateStart}${s.dateStart !== s.dateEnd ? ` a ${s.dateEnd}` : ""}`
      : "";
    return `${s.salon}${dates} ${s.startTime}-${s.endTime}`;
  }).join(", ");
  return `${first}${first !== last ? ` a ${last}` : ""}${slotText ? ` (${slotText})` : ""}`;
}

function formatSlotsForHistory(slots) {
  if (!Array.isArray(slots) || !slots.length) return "-";
  return slots.map(s => {
    const dates = s.dateStart && s.dateEnd ? ` ${s.dateStart}${s.dateStart !== s.dateEnd ? ` a ${s.dateEnd}` : ""}` : "";
    return `${s.salon}${dates} ${s.startTime}-${s.endTime}`;
  }).join(" | ");
}

function buildSeriesSnapshot(series) {
  if (!Array.isArray(series) || !series.length) return null;
  const sorted = series.slice().sort((a, b) => a.date.localeCompare(b.date));
  const first = sorted[0];
  const last = sorted[sorted.length - 1];
  return {
    name: first.name || "",
    status: first.status || "",
    userId: first.userId || "",
    notes: String(first.notes || ""),
    dateStart: getEventGeneralDateStart(first) || first.date || "",
    dateEnd: getEventGeneralDateEnd(first) || last.date || "",
    slots: uniqueSlotsFromSeries(sorted),
  };
}

function buildDraftSnapshot(drafts) {
  if (!Array.isArray(drafts) || !drafts.length) return null;
  const sorted = drafts.slice().sort((a, b) => a.date.localeCompare(b.date));
  const first = sorted[0];
  const last = sorted[sorted.length - 1];
  return {
    name: first.name || "",
    status: first.status || "",
    userId: first.userId || "",
    notes: String(first.notes || ""),
    dateStart: getEventGeneralDateStart(first) || first.date || "",
    dateEnd: getEventGeneralDateEnd(first) || last.date || "",
    slots: uniqueSlotsFromSeries(sorted),
  };
}

function appendDetailedEditHistory(key, actorUserId, oldSnap, newSnap) {
  if (!oldSnap || !newSnap) return;
  const changes = [];

  if (oldSnap.name !== newSnap.name) {
    changes.push(`Nombre: "${oldSnap.name}" -> "${newSnap.name}".`);
  }
  if (oldSnap.status !== newSnap.status) {
    changes.push(`Estado: ${oldSnap.status} -> ${newSnap.status}.`);
  }
  if (oldSnap.userId !== newSnap.userId) {
    changes.push(`Vendedor: ${getUserNameById(oldSnap.userId)} -> ${getUserNameById(newSnap.userId)}.`);
  }
  if (oldSnap.dateStart !== newSnap.dateStart || oldSnap.dateEnd !== newSnap.dateEnd) {
    changes.push(`Fechas: ${oldSnap.dateStart}${oldSnap.dateStart !== oldSnap.dateEnd ? ` a ${oldSnap.dateEnd}` : ""} -> ${newSnap.dateStart}${newSnap.dateStart !== newSnap.dateEnd ? ` a ${newSnap.dateEnd}` : ""}.`);
  }

  const oldSlotsText = formatSlotsForHistory(oldSnap.slots);
  const newSlotsText = formatSlotsForHistory(newSnap.slots);
  if (oldSlotsText !== newSlotsText) {
    changes.push(`Salon/Horario: ${oldSlotsText} -> ${newSlotsText}.`);
  }

  if ((oldSnap.notes || "").trim() !== (newSnap.notes || "").trim()) {
    const oldNotes = (oldSnap.notes || "").trim() || "(vacio)";
    const newNotes = (newSnap.notes || "").trim() || "(vacio)";
    changes.push(`Notas: ${oldNotes} -> ${newNotes}.`);
  }

  if (!changes.length) {
    changes.push("Edicion guardada sin cambios detectables.");
  }
  for (const msg of changes) {
    appendHistoryByKey(key, actorUserId, msg);
  }
}

function listDatesBetween(startIso, endIso) {
  const dates = [];
  let d = new Date(`${startIso}T00:00:00`);
  const end = new Date(`${endIso}T00:00:00`);
  while (d <= end) {
    dates.push(toISODate(d));
    d = addDays(d, 1);
  }
  return dates;
}

function reservationKeyFromEvent(ev) {
  if (!ev) return "";
  return ev.groupId || ev.id;
}

function getUserNameById(userId) {
  return state.users.find(u => u.id === userId)?.name || "Sistema";
}

function getAuthUserRecord() {
  const sid = String(authSession.userId || "").trim();
  if (!sid) return null;
  return (state.users || []).find((u) => String(u.id || "") === sid) || null;
}

function isAdminSession() {
  const user = getAuthUserRecord();
  const role = String(user?.role || "").trim().toLowerCase();
  const byFlag = user?.isAdmin === true || role === "admin";
  const sessionUsername = String(authSession.username || user?.username || "").trim().toLowerCase();
  const byUsername = sessionUsername === "admin";
  return byFlag || byUsername;
}

function isCurrentUserEventOwner(ev) {
  if (!ev) return false;
  const sid = String(authSession.userId || "").trim();
  if (!sid) return false;
  return String(ev.userId || "").trim() === sid;
}

function ensureHistoryStore() {
  if (!state.changeHistory || typeof state.changeHistory !== "object") {
    state.changeHistory = {};
  }
}

function ensureReminderStore() {
  if (!state.reminders || typeof state.reminders !== "object") {
    state.reminders = {};
  }
}

function ensureGlobalNotificationsStore() {
  if (!Array.isArray(state.globalNotifications)) {
    state.globalNotifications = [];
  }
}

function normalizeReleasedWindowRecord(candidate) {
  const row = candidate && typeof candidate === "object" ? candidate : {};
  return {
    date: String(row.date || "").trim(),
    salon: String(row.salon || "").trim(),
    startTime: String(row.startTime || "").trim(),
    endTime: String(row.endTime || "").trim(),
  };
}

function normalizeGlobalNotificationRecord(candidate) {
  const row = candidate && typeof candidate === "object" ? candidate : {};
  const createdAt = String(row.createdAt || new Date().toISOString()).trim();
  const expiresAt = String(row.expiresAt || "").trim();
  return {
    id: String(row.id || uid()).trim(),
    title: String(row.title || "Aviso del sistema").trim(),
    notes: String(row.notes || "").trim(),
    createdAt,
    expiresAt,
    eventId: String(row.eventId || "").trim(),
    salon: String(row.salon || "").trim(),
    date: String(row.date || "").trim(),
    time: String(row.time || "").trim(),
    releasedWindows: Array.isArray(row.releasedWindows)
      ? row.releasedWindows.map(normalizeReleasedWindowRecord).filter((w) => w.date && w.salon && w.startTime && w.endTime)
      : [],
    dismissedByUserIds: Array.isArray(row.dismissedByUserIds)
      ? row.dismissedByUserIds.map((x) => String(x || "").trim()).filter(Boolean)
      : [],
  };
}

function pruneExpiredGlobalNotifications() {
  ensureGlobalNotificationsStore();
  const nowMs = Date.now();
  const before = state.globalNotifications.length;
  state.globalNotifications = state.globalNotifications.filter((n) => {
    const expiry = String(n?.expiresAt || "").trim();
    if (!expiry) return true;
    const dt = new Date(expiry);
    if (Number.isNaN(dt.getTime())) return true;
    return dt.getTime() > nowMs;
  });
  return state.globalNotifications.length !== before;
}

function isReleasedWindowOccupiedNow(windowRow) {
  if (!windowRow?.date || !windowRow?.salon || !windowRow?.startTime || !windowRow?.endTime) return false;
  for (const ev of state.events || []) {
    if (!ev) continue;
    if (!isHardBlockingStatus(String(ev.status || "").trim())) continue;
    if (String(ev.date || "") !== String(windowRow.date || "")) continue;
    if (String(ev.salon || "") !== String(windowRow.salon || "")) continue;
    if (timesOverlap(String(ev.startTime || ""), String(ev.endTime || ""), String(windowRow.startTime || ""), String(windowRow.endTime || ""))) {
      return true;
    }
  }
  return false;
}

function pruneResolvedGlobalNotifications() {
  ensureGlobalNotificationsStore();
  const before = state.globalNotifications.length;
  state.globalNotifications = state.globalNotifications.filter((n) => {
    const windows = Array.isArray(n?.releasedWindows) ? n.releasedWindows : [];
    if (!windows.length) return true;
    const allTakenAgain = windows.every((w) => isReleasedWindowOccupiedNow(w));
    return !allTakenAgain;
  });
  return state.globalNotifications.length !== before;
}

function isGlobalNotificationDismissedForSession(notification) {
  const sid = String(authSession.userId || "").trim();
  if (!sid) return false;
  const dismissed = Array.isArray(notification?.dismissedByUserIds) ? notification.dismissedByUserIds : [];
  return dismissed.includes(sid);
}

function dismissGlobalNotificationForSession(notificationId, { persistRemote = true } = {}) {
  const id = String(notificationId || "").trim();
  if (!id) return false;
  ensureGlobalNotificationsStore();
  const target = state.globalNotifications.find((n) => String(n?.id || "") === id);
  if (!target) return false;

  const sid = String(authSession.userId || "").trim();
  if (!sid) {
    state.globalNotifications = state.globalNotifications.filter((n) => String(n?.id || "") !== id);
  } else {
    if (!Array.isArray(target.dismissedByUserIds)) target.dismissedByUserIds = [];
    if (!target.dismissedByUserIds.includes(sid)) target.dismissedByUserIds.push(sid);
  }
  if (persistRemote) persist();
  return true;
}

function addGlobalNotification(payload) {
  ensureGlobalNotificationsStore();
  const now = new Date();
  const createdAt = now.toISOString();
  const expiresAt = new Date(now.getTime() + (72 * 60 * 60 * 1000)).toISOString();
  const row = normalizeGlobalNotificationRecord({
    ...payload,
    createdAt,
    expiresAt,
    releasedWindows: Array.isArray(payload?.releasedWindows) ? payload.releasedWindows : [],
    dismissedByUserIds: [],
  });
  state.globalNotifications.unshift(row);
  if (state.globalNotifications.length > 200) {
    state.globalNotifications = state.globalNotifications.slice(0, 200);
  }
}

function appendHistoryByKey(key, actorUserId, changeText) {
  if (!key || !changeText) return;
  ensureHistoryStore();
  if (!Array.isArray(state.changeHistory[key])) state.changeHistory[key] = [];
  state.changeHistory[key].unshift({
    at: new Date().toISOString(),
    actorUserId: actorUserId || "",
    actorName: getUserNameById(actorUserId),
    change: changeText,
  });
  if (state.changeHistory[key].length > 200) {
    state.changeHistory[key] = state.changeHistory[key].slice(0, 200);
  }
}

function moveHistoryKey(oldKey, newKey) {
  if (!oldKey || !newKey || oldKey === newKey) return;
  ensureHistoryStore();
  const oldRows = Array.isArray(state.changeHistory[oldKey]) ? state.changeHistory[oldKey] : [];
  const newRows = Array.isArray(state.changeHistory[newKey]) ? state.changeHistory[newKey] : [];
  if (!oldRows.length && !newRows.length) return;
  state.changeHistory[newKey] = [...oldRows, ...newRows].slice(0, 200);
  delete state.changeHistory[oldKey];
}

function moveReminderKey(oldKey, newKey) {
  if (!oldKey || !newKey || oldKey === newKey) return;
  ensureReminderStore();
  const oldRows = Array.isArray(state.reminders[oldKey]) ? state.reminders[oldKey] : [];
  const newRows = Array.isArray(state.reminders[newKey]) ? state.reminders[newKey] : [];
  if (!oldRows.length && !newRows.length) return;
  state.reminders[newKey] = [...oldRows, ...newRows].slice(0, 200);
  delete state.reminders[oldKey];
}

function reminderDateTime(reminder) {
  const d = String(reminder?.date || "").trim();
  const t = String(reminder?.time || "").trim() || "00:00";
  const dt = new Date(`${d}T${t}:00`);
  return Number.isNaN(dt.getTime()) ? null : dt;
}

function pruneExpiredReminders({ persistRemote = true } = {}) {
  ensureReminderStore();
  const now = new Date();
  let changed = false;
  for (const [key, list] of Object.entries(state.reminders || {})) {
    const rows = Array.isArray(list) ? list : [];
    const next = rows.filter((r) => {
      if (r?.done === true) return false;
      const dt = reminderDateTime(r);
      return !!dt && dt.getTime() >= now.getTime();
    });
    if (next.length !== rows.length) changed = true;
    if (next.length) state.reminders[key] = next;
    else delete state.reminders[key];
  }
  if (changed && persistRemote) persist();
  return changed;
}

function getReservationReminders(ev) {
  ensureReminderStore();
  pruneExpiredReminders({ persistRemote: false });
  const key = reservationKeyFromEvent(ev);
  const rows = Array.isArray(state.reminders[key]) ? state.reminders[key] : [];
  return rows.slice().sort((a, b) => {
    const da = reminderDateTime(a)?.getTime() || 0;
    const db = reminderDateTime(b)?.getTime() || 0;
    return da - db;
  });
}

function addReminderForEvent(ev, payload) {
  if (!ev) return;
  ensureReminderStore();
  const key = reservationKeyFromEvent(ev);
  if (!Array.isArray(state.reminders[key])) state.reminders[key] = [];
  state.reminders[key].push({
    id: uid(),
    date: payload.date,
    time: payload.time,
    channel: payload.channel,
    notes: payload.notes || "",
    done: false,
    createdAt: new Date().toISOString(),
    createdByUserId: payload.createdByUserId || ev.userId || "",
  });
}

function reminderBadgeText(ev) {
  const rows = getReservationReminders(ev);
  if (!rows.length) return "";
  const primary = getPrimaryReminderForEvent(ev);
  const chosen = primary?.reminder;
  if (!chosen) return "";
  const dateText = String(chosen.date || "");
  const timeText = String(chosen.time || "");
  const channelText = String(chosen.channel || "");
  const status = String(primary?.status || "upcoming");
  const mins = Number(primary?.minutes || 0);
  const extra = rows.length > 1 ? ` +${rows.length - 1}` : "";
  if (status === "overdue") {
    return `Cita pendiente atrasada ${dateText} ${timeText}${channelText ? ` ${channelText}` : ""}${extra}`.trim();
  }
  if (status === "soon") {
    return `Cita pendiente en ${Math.max(1, mins)} min${channelText ? ` ${channelText}` : ""}${extra}`.trim();
  }
  if (status === "today") {
    return `Cita hoy ${timeText}${channelText ? ` ${channelText}` : ""}${extra}`.trim();
  }
  return `Cita ${dateText} ${timeText}${channelText ? ` ${channelText}` : ""}${extra}`.trim();
}

function reminderBadgeClass(ev) {
  const primary = getPrimaryReminderForEvent(ev);
  const status = String(primary?.status || "");
  if (status === "overdue") return "reminderPill reminderOverdue";
  if (status === "soon") return "reminderPill reminderSoon";
  if (status === "today") return "reminderPill reminderToday";
  return "reminderPill";
}

function getReminderStatus(reminder, now = new Date()) {
  const dt = reminderDateTime(reminder);
  if (!dt) return { status: "upcoming", minutes: Number.POSITIVE_INFINITY, dt: null };
  const minutes = Math.floor((dt.getTime() - now.getTime()) / 60000);
  if (minutes <= 120) return { status: "soon", minutes, dt };
  const sameDay = toISODate(dt) === toISODate(now);
  if (sameDay) return { status: "today", minutes, dt };
  return { status: "upcoming", minutes, dt };
}

function getPrimaryReminderForEvent(ev) {
  const rows = getReservationReminders(ev);
  if (!rows.length) return null;
  const now = new Date();
  const withMeta = rows
    .map((r) => ({ reminder: r, ...getReminderStatus(r, now) }))
    .filter((x) => x.dt);
  if (!withMeta.length) return null;
  const upcoming = withMeta.filter((x) => x.minutes >= 0).sort((a, b) => a.minutes - b.minutes);
  return upcoming.length ? upcoming[0] : null;
}

function findReminderContext(eventId, reminderId) {
  const ev = (state.events || []).find((x) => String(x.id || "") === String(eventId || ""));
  if (!ev) return null;
  const key = reservationKeyFromEvent(ev);
  const rows = Array.isArray(state.reminders?.[key]) ? state.reminders[key] : [];
  const idx = rows.findIndex((r) => String(r?.id || "") === String(reminderId || ""));
  if (idx < 0) return null;
  return { ev, key, rows, idx, reminder: rows[idx] };
}

async function openReminderEditor(eventId, reminderId) {
  const ctx = findReminderContext(eventId, reminderId);
  if (!ctx) return;
  const r = ctx.reminder;
  let nextDate = String(r.date || "");
  let nextTime = String(r.time || "");
  let nextChannel = String(r.channel || "Telefono");
  let nextNotes = String(r.notes || "");

  if (window.Swal && typeof window.Swal.fire === "function") {
    const result = await window.Swal.fire({
      title: "Editar cita",
      html: `
        <div style="display:grid;gap:8px;text-align:left">
          <label style="font-size:12px;color:#35506b;">Fecha</label>
          <input id="swalReminderDate" class="swal2-input" type="date" value="${escapeHtml(nextDate)}" style="margin:0" />
          <label style="font-size:12px;color:#35506b;">Hora (HH:mm)</label>
          <input id="swalReminderTime" class="swal2-input" type="text" value="${escapeHtml(nextTime)}" style="margin:0" />
          <label style="font-size:12px;color:#35506b;">Medio</label>
          <select id="swalReminderChannel" class="swal2-input" style="margin:0">
            ${["Telefono", "Correo", "Teams", "Google Meet", "WhatsApp"].map((x) => `<option value="${x}"${x === nextChannel ? " selected" : ""}>${x}</option>`).join("")}
          </select>
          <label style="font-size:12px;color:#35506b;">Detalle</label>
          <input id="swalReminderNotes" class="swal2-input" type="text" value="${escapeHtml(nextNotes)}" style="margin:0" />
        </div>
      `,
      showCancelButton: true,
      confirmButtonText: "Guardar cambios",
      cancelButtonText: "Cancelar",
      background: "#f8fbff",
      color: "#10243b",
      confirmButtonColor: "#2563eb",
      preConfirm: () => {
        const date = String(document.getElementById("swalReminderDate")?.value || "").trim();
        const time = String(document.getElementById("swalReminderTime")?.value || "").trim();
        const channel = String(document.getElementById("swalReminderChannel")?.value || "").trim();
        const notes = String(document.getElementById("swalReminderNotes")?.value || "").trim();
        if (!date || !time || !channel) {
          window.Swal.showValidationMessage("Fecha, hora y medio son obligatorios.");
          return null;
        }
        if (!isValidClockTime(time)) {
          window.Swal.showValidationMessage("Hora invalida. Usa HH:mm.");
          return null;
        }
        const dt = new Date(`${date}T${time}:00`);
        if (Number.isNaN(dt.getTime()) || dt.getTime() < Date.now()) {
          window.Swal.showValidationMessage("La cita debe ser futura.");
          return null;
        }
        return { date, time, channel, notes };
      },
    });
    if (!result.isConfirmed || !result.value) return;
    nextDate = result.value.date;
    nextTime = result.value.time;
    nextChannel = result.value.channel;
    nextNotes = result.value.notes;
  } else {
    const date = String(window.prompt("Fecha (AAAA-MM-DD):", nextDate) || "").trim();
    const time = String(window.prompt("Hora (HH:mm):", nextTime) || "").trim();
    const channel = String(window.prompt("Medio:", nextChannel) || "").trim();
    const notes = String(window.prompt("Detalle:", nextNotes) || "").trim();
    if (!date || !time || !channel || !isValidClockTime(time)) return;
    const dt = new Date(`${date}T${time}:00`);
    if (Number.isNaN(dt.getTime()) || dt.getTime() < Date.now()) return;
    nextDate = date;
    nextTime = time;
    nextChannel = channel;
    nextNotes = notes;
  }

  const fresh = findReminderContext(eventId, reminderId);
  if (!fresh) return;
  fresh.rows[fresh.idx] = {
    ...fresh.reminder,
    date: nextDate,
    time: nextTime,
    channel: nextChannel,
    notes: nextNotes,
  };
  appendHistoryByKey(
    fresh.key,
    authSession.userId || fresh.ev.userId || "",
    `Cita reprogramada: ${nextDate} ${nextTime} via ${nextChannel}${nextNotes ? ` (${nextNotes})` : ""}.`
  );
  persist();
  render();
  const currentEv = historyTargetEventId ? state.events.find((x) => x.id === historyTargetEventId) : fresh.ev;
  renderAppointmentsForEvent(currentEv || null);
  runUpcomingReminderChecks();
  refreshTopbarReminders();
  toast("Cita actualizada.");
}

function removeReminder(eventId, reminderId) {
  const ctx = findReminderContext(eventId, reminderId);
  if (!ctx) return;
  const removed = ctx.rows.splice(ctx.idx, 1)[0];
  if (!ctx.rows.length) delete state.reminders[ctx.key];
  appendHistoryByKey(
    ctx.key,
    authSession.userId || ctx.ev.userId || "",
    `Cita eliminada: ${removed?.date || ""} ${removed?.time || ""} ${removed?.channel || ""}.`.trim()
  );
  persist();
  render();
  const currentEv = historyTargetEventId ? state.events.find((x) => x.id === historyTargetEventId) : ctx.ev;
  renderAppointmentsForEvent(currentEv || null);
  runUpcomingReminderChecks();
  refreshTopbarReminders();
  toast("Cita eliminada.");
}

function markReminderDone(eventId, reminderId) {
  const ctx = findReminderContext(eventId, reminderId);
  if (!ctx) return;
  const target = ctx.rows[ctx.idx];
  ctx.rows[ctx.idx] = { ...target, done: true };
  appendHistoryByKey(
    ctx.key,
    authSession.userId || ctx.ev.userId || "",
    `Cita atendida: ${String(target?.date || "")} ${String(target?.time || "")}.`
  );
  pruneExpiredReminders({ persistRemote: false });
  persist();
  render();
  const currentEv = historyTargetEventId ? state.events.find((x) => x.id === historyTargetEventId) : ctx.ev;
  renderAppointmentsForEvent(currentEv || null);
  runUpcomingReminderChecks();
  refreshTopbarReminders();
  toast("Cita marcada como atendida.");
}

function collectTopbarReminderFeed() {
  const byReservation = new Map();
  for (const ev of state.events || []) {
    if (!ev) continue;
    if (!isCurrentUserEventOwner(ev)) continue;
    if (ev.status === STATUS.CANCELADO || ev.status === STATUS.PERDIDO) continue;
    const key = String(reservationKeyFromEvent(ev) || "").trim();
    if (!key) continue;
    if (!byReservation.has(key)) byReservation.set(key, ev);
  }
  const feed = [];
  const now = new Date();
  for (const ev of byReservation.values()) {
    const reminders = getReservationReminders(ev);
    for (const r of reminders) {
      const meta = getReminderStatus(r, now);
      if (!meta.dt) continue;
      feed.push({
        reservationKey: reservationKeyFromEvent(ev),
        eventId: ev.id,
        eventName: String(ev.name || "Reserva"),
        salon: String(ev.salon || ""),
        date: String(r.date || ""),
        time: String(r.time || ""),
        channel: String(r.channel || ""),
        notes: String(r.notes || ""),
        status: meta.status,
        minutes: meta.minutes,
        dt: meta.dt,
      });
    }
  }
  pruneExpiredGlobalNotifications();
  pruneResolvedGlobalNotifications();
  for (const item of state.globalNotifications || []) {
    if (isGlobalNotificationDismissedForSession(item)) continue;
    const createdDt = new Date(String(item.createdAt || ""));
    if (Number.isNaN(createdDt.getTime())) continue;
    feed.push({
      reservationKey: "",
      eventId: String(item.eventId || ""),
      globalNotificationId: String(item.id || ""),
      eventName: String(item.title || "Aviso del sistema"),
      salon: String(item.salon || ""),
      date: String(item.date || toISODate(createdDt)),
      time: String(item.time || toHHMM(createdDt)),
      channel: "Sistema",
      notes: String(item.notes || ""),
      status: "today",
      minutes: 0,
      dt: createdDt,
    });
  }
  feed.sort((a, b) => a.dt.getTime() - b.dt.getTime());
  return feed;
}

function closeTopbarReminderPanel() {
  if (!el.topbarReminderPanel || !el.btnTopbarReminders) return;
  el.topbarReminderPanel.hidden = true;
  el.btnTopbarReminders.setAttribute("aria-expanded", "false");
}

function renderTopbarReminderPanel(feed) {
  if (!el.topbarReminderList || !el.topbarReminderSubtitle) return;
  const urgent = feed.filter((x) => x.minutes <= 24 * 60);
  el.topbarReminderSubtitle.textContent = urgent.length
    ? `${urgent.length} dentro de 24h`
    : (feed.length ? `${feed.length} pendientes` : "Sin pendientes");
  el.topbarReminderList.innerHTML = "";
  if (!feed.length) {
    el.topbarReminderList.innerHTML = `<div class="topbarReminderEmpty">No hay citas pendientes.</div>`;
    return;
  }
  const rows = feed.slice(0, 20);
  for (const item of rows) {
    const cls = item.status === "overdue"
      ? "overdue"
      : (item.status === "soon" ? "soon" : (item.status === "today" ? "today" : ""));
    const relative = item.status === "overdue"
      ? `Vencida hace ${Math.abs(item.minutes)} min`
      : (item.status === "soon" ? `En ${Math.max(1, item.minutes)} min` : item.date);
    const node = document.createElement("button");
    node.type = "button";
    node.className = `topbarReminderItem ${cls}`.trim();
    node.dataset.eventId = String(item.eventId || "");
    if (item.globalNotificationId) {
      node.dataset.globalNotificationId = String(item.globalNotificationId || "");
    }
    node.innerHTML = `
      <div class="topbarReminderItemHead">
        <strong>${escapeHtml(item.eventName)}</strong>
        <span>${escapeHtml(relative)}</span>
      </div>
      <div class="topbarReminderItemMeta">
        <span>${escapeHtml(item.date)} ${escapeHtml(item.time)}</span>
        <span>${escapeHtml(item.salon || "-")}</span>
        <span>${escapeHtml(item.channel || "-")}</span>
      </div>
      ${item.notes ? `<div class="topbarReminderItemNotes">${escapeHtml(item.notes)}</div>` : ""}
      ${item.globalNotificationId ? `<div class="topbarReminderItemMeta"><span></span><span data-dismiss-global-notification-id="${escapeHtml(String(item.globalNotificationId || ""))}" style="cursor:pointer;text-decoration:underline">Quitar</span></div>` : ""}
    `;
    el.topbarReminderList.appendChild(node);
  }
}

function refreshTopbarReminders() {
  if (!el.btnTopbarReminders || !el.topbarReminderCount) return;
  const feed = collectTopbarReminderFeed();
  const urgent = feed.filter((x) => x.minutes <= 24 * 60);
  const urgentCount = urgent.length;
  el.btnTopbarReminders.classList.toggle("hasUrgent", urgentCount > 0);
  if (urgentCount > 0) {
    el.topbarReminderCount.hidden = false;
    el.topbarReminderCount.style.display = "";
    el.topbarReminderCount.textContent = String(Math.min(99, urgentCount));
  } else {
    el.topbarReminderCount.hidden = true;
    el.topbarReminderCount.style.display = "none";
    el.topbarReminderCount.textContent = "0";
  }
  renderTopbarReminderPanel(feed);
}

function renderAppointmentsForEvent(ev) {
  if (!el.appointmentBody) return;
  el.appointmentBody.innerHTML = "";
  if (!ev) return;
  const rows = getReservationReminders(ev);
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="5">Sin citas pendientes para esta reserva.</td>`;
    el.appointmentBody.appendChild(tr);
    return;
  }
  const now = new Date();
  for (const r of rows) {
    const meta = getReminderStatus(r, now);
    const status = meta.status === "soon" ? "Proxima" : (meta.status === "today" ? "Hoy" : "Pendiente");
    const cls = meta.status === "soon" ? "reminderState soon" : (meta.status === "today" ? "reminderState today" : "reminderState");
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(`${String(r.date || "")} ${String(r.time || "")}`.trim())}</td>
      <td>${escapeHtml(String(r.channel || "-"))}</td>
      <td>${escapeHtml(String(r.notes || "-"))}</td>
      <td><span class="${cls}">${escapeHtml(status)}</span></td>
      <td>
        <div class="appointmentActions">
          <button type="button" class="apptIconBtn apptDone" title="Marcar atendida" aria-label="Marcar atendida" data-reminder-action="done" data-reminder-id="${escapeHtml(String(r.id || ""))}" data-event-id="${escapeHtml(String(ev.id || ""))}">&#10003;</button>
          <button type="button" class="apptIconBtn apptEdit" title="Editar cita" aria-label="Editar cita" data-reminder-action="edit" data-reminder-id="${escapeHtml(String(r.id || ""))}" data-event-id="${escapeHtml(String(ev.id || ""))}">&#9998;</button>
          <button type="button" class="apptIconBtn apptDelete" title="Eliminar cita" aria-label="Eliminar cita" data-reminder-action="delete" data-reminder-id="${escapeHtml(String(r.id || ""))}" data-event-id="${escapeHtml(String(ev.id || ""))}">&#128465;</button>
        </div>
      </td>
    `;
    el.appointmentBody.appendChild(tr);
  }
}

function setAppointmentsPanelVisible(visible) {
  if (!el.appointmentPanel) return;
  el.appointmentPanel.hidden = !visible;
  if (el.btnToggleAppointments) {
    el.btnToggleAppointments.textContent = visible ? "Ocultar citas" : "Ver citas";
  }
  updateEventPanelsBackButtonVisibility();
}

function runUpcomingReminderChecks() {
  if (el.loginScreen && !el.loginScreen.hidden) return;
  pruneExpiredReminders();
  const now = new Date();
  for (const ev of state.events || []) {
    if (!isCurrentUserEventOwner(ev)) continue;
    if (ev.status === STATUS.CANCELADO || ev.status === STATUS.PERDIDO) continue;
    const primary = getPrimaryReminderForEvent(ev);
    if (!primary?.reminder) continue;
    const minutes = Number(primary.minutes);
    const key = `${reservationKeyFromEvent(ev)}|${primary.reminder.id || `${primary.reminder.date}|${primary.reminder.time}`}`;
    if (notifiedReminderKeys.has(key)) continue;
    if (minutes > 30 || minutes < 0) continue;
    notifiedReminderKeys.add(key);
    const eventName = String(ev.name || "Reserva").trim();
    toast(`Reserva pendiente: cita en ${Math.max(1, minutes)} min (${eventName})`);
  }
}

function formatDateTime(iso) {
  const d = new Date(iso);
  if (Number.isNaN(d.getTime())) return String(iso || "");
  return d.toLocaleString("es-GT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function renderHistoryForEvent(ev) {
  if (!el.historyPanel || !el.historyBody) {
    return;
  }
  if (!ev) {
    el.historyPanel.hidden = true;
    el.historyBody.innerHTML = "";
    return;
  }
  ensureHistoryStore();
  const key = reservationKeyFromEvent(ev);
  const rows = Array.isArray(state.changeHistory[key]) ? state.changeHistory[key] : [];
  el.historyBody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="3">Sin cambios registrados.</td>`;
    el.historyBody.appendChild(tr);
  } else {
    for (const row of rows) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(formatDateTime(row.at))}</td>
        <td>${escapeHtml(row.actorName || "Sistema")}</td>
        <td>${escapeHtml(row.change || "")}</td>
      `;
      el.historyBody.appendChild(tr);
    }
  }
}

function setHistoryPanelVisible(visible) {
  if (!el.historyPanel) return;
  el.historyPanel.hidden = !visible;
  if (el.btnToggleHistory) {
    el.btnToggleHistory.textContent = visible ? "Ocultar historial" : "Historial";
  }
  updateEventPanelsBackButtonVisibility();
}

function updateEventPanelsBackButtonVisibility() {
  if (!el.btnBackToForm) return;
  const historyVisible = !!el.historyPanel && !el.historyPanel.hidden;
  const appointmentsVisible = !!el.appointmentPanel && !el.appointmentPanel.hidden;
  el.btnBackToForm.hidden = !(historyVisible || appointmentsVisible);
}

function backToEventForm() {
  setHistoryPanelVisible(false);
  setAppointmentsPanelVisible(false);
  if (el.eventForm) {
    el.eventForm.scrollIntoView({ block: "start", behavior: "smooth" });
  }
}

function totalGridMinutes() {
  return (HOUR_END - HOUR_START + 1) * 60; // includes 24:00 end marker
}

function maxSelectableMinutes() {
  return (HOUR_END - HOUR_START) * 60; // 23:00 as last selectable end in form
}

function startSelection(ev, col) {
  const point = getGridPoint(ev.clientX, ev.clientY);
  if (!point) return;
  const startDay = Number(col.dataset.dayIndex || point.dayIndex);
  interaction.selecting = {
    startDay,
    currentDay: startDay,
    startMinute: point.minute,
    currentMinute: point.minute,
  };
  ensureSelectionBox();
  updateSelectionBox();
  ev.preventDefault();
}

function ensureSelectionBox() {
  if (interaction.selectionBox) return;
  const box = document.createElement("div");
  box.className = "selectionBox";
  interaction.selectionBox = box;
  el.grid.appendChild(box);
}

function updateSelectionBox() {
  if (!interaction.selecting || !interaction.selectionBox) return;
  const sel = interaction.selecting;
  const dayStart = Math.min(sel.startDay, sel.currentDay);
  const dayEnd = Math.max(sel.startDay, sel.currentDay);
  const minA = Math.min(sel.startMinute, sel.currentMinute);
  const minB = Math.max(sel.startMinute, sel.currentMinute);
  const minDuration = 60;
  const endMinute = minB === minA ? Math.min(minA + minDuration, maxSelectableMinutes()) : Math.min(minB, maxSelectableMinutes());

  const colWidth = el.grid.querySelector(".dayCol")?.offsetWidth || 0;
  const topPx = (minA / 60) * HOUR_HEIGHT;
  const heightPx = Math.max((endMinute - minA) / 60 * HOUR_HEIGHT, HOUR_HEIGHT / 2);
  const leftPx = dayStart * colWidth + 10;
  const widthPx = Math.max(((dayEnd - dayStart + 1) * colWidth) - 20, 40);

  interaction.selectionBox.style.top = `${topPx}px`;
  interaction.selectionBox.style.left = `${leftPx}px`;
  interaction.selectionBox.style.width = `${widthPx}px`;
  interaction.selectionBox.style.height = `${heightPx}px`;
  interaction.selectionBox.hidden = false;
}

function finishSelection() {
  if (!interaction.selecting) return;
  if (!state.users.length) {
    interaction.selecting = null;
    clearSelectionBox();
    return toast("Primero agrega al menos un usuario.");
  }
  if (!state.salones.length) {
    interaction.selecting = null;
    clearSelectionBox();
    return toast("Primero agrega al menos un salon.");
  }
  const sel = interaction.selecting;
  const dayStart = Math.min(sel.startDay, sel.currentDay);
  const dayEnd = Math.max(sel.startDay, sel.currentDay);
  const minA = Math.min(sel.startMinute, sel.currentMinute);
  const minB = Math.max(sel.startMinute, sel.currentMinute);
  const endMinute = minB === minA ? Math.min(minA + 60, maxSelectableMinutes()) : Math.min(minB, maxSelectableMinutes());

  const dates = [];
  for (let d = dayStart; d <= dayEnd; d++) {
    dates.push(toISODate(addDays(viewStart, d)));
  }

  openModalForCreate({
    date: addDays(viewStart, dayStart),
    start: minutesToTime(minA),
    end: minutesToTime(endMinute),
    salon: selectedSalon,
    rangeDates: dates,
  });

  interaction.selecting = null;
  clearSelectionBox();
}

function clearSelectionBox() {
  if (!interaction.selectionBox) return;
  interaction.selectionBox.remove();
  interaction.selectionBox = null;
}

function startEventDrag(ev, eventData, card) {
  if (ev.target.closest("button, input, select, textarea")) return;
  if (isEventSeriesInPast(eventData) && !hasPastEventEditAuthorization(eventData)) {
    toast("Evento pasado bloqueado. Abre el evento y solicita codigo de administrador.");
    return;
  }
  const point = getGridPoint(ev.clientX, ev.clientY);
  if (!point) return;
  const dayIndex = Math.floor((stripTime(new Date(eventData.date + "T00:00:00")) - stripTime(viewStart)) / (1000 * 60 * 60 * 24));
  const startMinute = timeToMinutes(eventData.startTime) - HOUR_START * 60;
  const endMinute = timeToMinutes(eventData.endTime) - HOUR_START * 60;
  const duration = Math.max(30, endMinute - startMinute);

  interaction.dragging = {
    eventId: eventData.id,
    originDay: dayIndex,
    originMinute: startMinute,
    dayIndex,
    minute: startMinute,
    duration,
    offset: point.minute - startMinute,
    startClientX: ev.clientX,
    startClientY: ev.clientY,
    hasMoved: false,
  };

  ev.preventDefault();
}

function startEventStretch(ev, eventData) {
  if (isEventSeriesInPast(eventData) && !hasPastEventEditAuthorization(eventData)) {
    toast("Evento pasado bloqueado. Abre el evento y solicita codigo de administrador.");
    return;
  }
  const baseSeries = getEventSeries(eventData);
  const sorted = baseSeries.slice().sort((a, b) => a.date.localeCompare(b.date));
  const anchorDate = sorted[0]?.date || eventData.date;
  const lastDate = sorted[sorted.length - 1]?.date || eventData.date;
  interaction.stretching = {
    eventId: eventData.id,
    anchorDate,
    endDate: lastDate,
  };
  ev.preventDefault();
  ev.stopPropagation();
}

function onGlobalPointerMove(ev) {
  if (interaction.stretching) {
    const point = getGridPoint(ev.clientX, ev.clientY);
    if (!point) return;
    const candidate = toISODate(addDays(viewStart, point.dayIndex));
    interaction.stretching.endDate = candidate < interaction.stretching.anchorDate
      ? interaction.stretching.anchorDate
      : candidate;
    return;
  }

  if (interaction.selecting) {
    autoScrollGridForSelection(ev.clientY);
    const point = getGridPoint(ev.clientX, ev.clientY, true);
    if (!point) return;
    interaction.selecting.currentDay = point.dayIndex;
    interaction.selecting.currentMinute = point.minute;
    updateSelectionBox();
    return;
  }

  if (interaction.dragging) {
    const point = getGridPoint(ev.clientX, ev.clientY);
    if (!point) return;
    if (!interaction.dragging.hasMoved) {
      const dx = ev.clientX - interaction.dragging.startClientX;
      const dy = ev.clientY - interaction.dragging.startClientY;
      if ((dx * dx + dy * dy) < 16) return;
      interaction.dragging.hasMoved = true;
      const card = el.grid.querySelector(`.event[data-event-id="${interaction.dragging.eventId}"]`);
      if (card) card.classList.add("dragging");
      document.body.classList.add("draggingEvent");
    }
    interaction.dragging.dayIndex = point.dayIndex;
    const nextStart = clampToGridMinute(point.minute - interaction.dragging.offset, interaction.dragging.duration);
    interaction.dragging.minute = nextStart;
  }
}

function onGlobalPointerUp() {
  if (interaction.stretching) {
    const stretch = interaction.stretching;
    interaction.stretching = null;
    applyStretchToEvent(stretch);
    interaction.suppressClickUntil = Date.now() + 220;
    return;
  }

  if (interaction.selecting) {
    finishSelection();
    interaction.suppressClickUntil = Date.now() + 120;
    return;
  }

  if (interaction.dragging) {
    const drag = interaction.dragging;
    interaction.dragging = null;
    document.body.classList.remove("draggingEvent");
    const oldCard = el.grid.querySelector(`.event[data-event-id="${drag.eventId}"]`);
    if (oldCard) oldCard.classList.remove("dragging");
    if (!drag.hasMoved) return;

    const ev = state.events.find(x => x.id === drag.eventId);
    if (!ev) return;
    const key = reservationKeyFromEvent(ev);
    const oldDate = ev.date;
    const oldStart = ev.startTime;
    const oldEnd = ev.endTime;

    const nextDate = toISODate(addDays(viewStart, drag.dayIndex));
    const nextStart = minutesToTime(drag.minute);
    const nextEnd = minutesToTime(drag.minute + drag.duration);

    if (ev.date === nextDate && ev.startTime === nextStart && ev.endTime === nextEnd) {
      interaction.suppressClickUntil = Date.now() + 150;
      return;
    }

    const draft = { ...ev, date: nextDate, startTime: nextStart, endTime: nextEnd };
    const rules = evaluateRules(draft, ignoreIds);
    if (!rules.ok) {
      toast(rules.message || "No se puede mover por reglas/choques.");
      render();
      interaction.suppressClickUntil = Date.now() + 220;
      return;
    }

    ev.date = nextDate;
    ev.startTime = nextStart;
    ev.endTime = nextEnd;
    appendHistoryByKey(
      key,
      ev.userId,
      `Movido: ${oldDate} ${oldStart}-${oldEnd} -> ${nextDate} ${nextStart}-${nextEnd}.`
    );
    persist();
    render();
    toast("Evento movido.");
    interaction.suppressClickUntil = Date.now() + 220;
  }
}

function applyStretchToEvent(stretch) {
  const ev = state.events.find(x => x.id === stretch.eventId);
  if (!ev) return;
  if (isEventSeriesInPast(ev) && !hasPastEventEditAuthorization(ev)) {
    toast("Evento pasado bloqueado. Solicita codigo de administrador.");
    return;
  }
  const series = getEventSeries(ev);
  const oldKey = reservationKeyFromEvent(ev);
  const oldSummary = summarizeSeriesWindow(series);
  const replaceIds = new Set(series.map(x => x.id));
  const targetDates = listDatesBetween(stretch.anchorDate, stretch.endDate);
  const groupId = targetDates.length > 1 ? (ev.groupId || `grp_${uid()}`) : null;
  const byDate = new Map(series.map(x => [x.date, x]));

  const drafts = targetDates.map((d) => ({
    id: byDate.get(d)?.id || uid(),
    name: ev.name,
    salon: ev.salon,
    date: d,
    groupId,
    status: ev.status,
    startTime: ev.startTime,
    endTime: ev.endTime,
    userId: ev.userId,
    pax: ev.pax ?? null,
    notes: ev.notes || "",
  }));

  for (const draft of drafts) {
    const rules = evaluateRules(draft, replaceIds);
    if (!rules.ok) {
      toast(rules.message || "No se puede estirar por validaciones.");
      return;
    }
  }

  state.events = state.events.filter(x => !replaceIds.has(x.id));
  state.events.push(...drafts);
  const finalKey = groupId || drafts[0]?.id || oldKey;
  moveHistoryKey(oldKey, finalKey);
  moveReminderKey(oldKey, finalKey);
  appendHistoryByKey(finalKey, ev.userId, `Rango ajustado: ${oldSummary} -> ${summarizeDraftWindow(drafts)}.`);
  persist();
  render();
  toast(drafts.length > 1 ? `Reserva extendida a ${drafts.length} dias.` : "Reserva ajustada.");
}

function getGridPoint(clientX, clientY, allowOutside = false) {
  const rect = el.grid.getBoundingClientRect();
  if (!allowOutside && (clientX < rect.left || clientX > rect.right || clientY < rect.top || clientY > rect.bottom)) {
    return null;
  }
  const col = el.grid.querySelector(".dayCol");
  if (!col) return null;

  const clampedX = clamp(clientX, rect.left, rect.right - 1);
  const clampedY = clamp(clientY, rect.top, rect.bottom - 1);
  const colWidth = col.offsetWidth || 1;
  const x = clampedX - rect.left + el.grid.scrollLeft;
  const y = clampedY - rect.top + el.grid.scrollTop;
  const dayIndex = clamp(Math.floor(x / colWidth), 0, getVisibleDayCount() - 1);

  const totalMinutes = maxSelectableMinutes();
  const minuteRaw = (y / HOUR_HEIGHT) * 60;
  const minute = Math.round(minuteRaw / SNAP_MINUTES) * SNAP_MINUTES;
  return { dayIndex, minute: clamp(minute, 0, totalMinutes) };
}

function autoScrollGridForSelection(clientY) {
  if (!interaction.selecting) return;
  const rect = el.grid.getBoundingClientRect();
  if (clientY > rect.bottom - AUTO_SCROLL_EDGE_PX) {
    el.grid.scrollTop += AUTO_SCROLL_STEP_PX;
    el.timeCol.scrollTop = el.grid.scrollTop;
  } else if (clientY < rect.top + AUTO_SCROLL_EDGE_PX) {
    el.grid.scrollTop -= AUTO_SCROLL_STEP_PX;
    el.timeCol.scrollTop = el.grid.scrollTop;
  }
}

function clampToGridMinute(minuteValue, duration) {
  const totalMinutes = maxSelectableMinutes();
  const effectiveDuration = duration > 0 ? duration : SNAP_MINUTES;
  const maxStart = Math.max(0, totalMinutes - effectiveDuration);
  const snapped = Math.round(minuteValue / SNAP_MINUTES) * SNAP_MINUTES;
  return clamp(snapped, 0, maxStart);
}

// ================== Rules & Conflicts ==================

function buildConflictIgnoreIdsForEditing() {
  const editingId = String(el.eventId?.value || "").trim();
  if (!editingId) return null;
  const editingEvent = (state.events || []).find((x) => String(x?.id || "") === editingId);
  if (!editingEvent) return null;
  return new Set(getEventSeries(editingEvent).map((x) => String(x?.id || "")).filter(Boolean));
}

function updateRulesAndConflictsUI() {
  applyStatusSelectTheme();
  const draft = currentDraftFromForm();
  if (!draft) return;

  const ignoreIds = buildConflictIgnoreIdsForEditing();
  // Lista de conflictos (mismo salon + traslape + misma fecha), excluyendo la reserva en edicion
  const conflicts = findConflicts(draft, ignoreIds);
  const hardBlocks = findHardBlocks(draft, ignoreIds);

  // Render conflict UI
  if (conflicts.length) {
    el.conflictsBox.hidden = false;
    el.conflictsList.innerHTML = "";
    const maxVisible = 2;
    const visible = conflicts.slice(0, maxVisible);
    for (const c of visible) {
      const item = document.createElement("div");
      item.className = "conflictItem";
      const col = statusColor(c.status);
      item.innerHTML = `
        <div class="conflictLeft">
          <span class="dot" style="background:${col}"></span>
          <span class="conflictName">${escapeHtml(c.name)}</span>
        </div>
        <div class="conflictRight">${escapeHtml(c.status)} - ${escapeHtml(c.startTime)}-${escapeHtml(c.endTime)}</div>
      `;
      el.conflictsList.appendChild(item);
    }
    if (conflicts.length > visible.length) {
      const more = document.createElement("div");
      more.className = "conflictItem conflictMore";
      more.innerHTML = `<div class="conflictLeft"><span class="conflictName">Y ${conflicts.length - visible.length} choque(s) mas...</span></div>`;
      el.conflictsList.appendChild(more);
    }
  } else {
    el.conflictsBox.hidden = true;
    el.conflictsList.innerHTML = "";
  }

  const rules = evaluateRules(draft, ignoreIds);
  el.statusHint.textContent = rules.hint || "";
  // Soft-disable some status options when blocked
  applyStatusOptionDisabling(draft, hardBlocks);
}
function applyStatusOptionDisabling(draft, hardBlocks) {
  // Reset all enabled
  Array.from(el.eventStatus.options).forEach(o => {
    o.disabled = isAutoStatus(o.value);
  });

  const hasHardBlock = hardBlocks.length > 0;
  const hasMaintenanceDayBlock = findMaintenanceDayBlocks(draft).length > 0;

  // Si hay Confirmado/Pre reserva en choque:
  // - No permitir elegir Confirmado ni Pre reserva para un evento que se cruce
  if (hasHardBlock) {
    for (const opt of Array.from(el.eventStatus.options)) {
      if (
        opt.value === STATUS.CONFIRMADO
        || opt.value === STATUS.PRERESERVA
        || opt.value === STATUS.MANTENIMIENTO
      ) {
        opt.disabled = true;
      }
    }
    // Si el usuario ya lo tenia seleccionado, moverlo a Lista
    if (
      draft.status === STATUS.CONFIRMADO
      || draft.status === STATUS.PRERESERVA
      || draft.status === STATUS.MANTENIMIENTO
    ) {
      el.eventStatus.value = STATUS.LISTA;
    }
  }

  if (hasMaintenanceDayBlock) {
    for (const opt of Array.from(el.eventStatus.options)) {
      if (opt.value === STATUS.CONFIRMADO || opt.value === STATUS.PRERESERVA) {
        opt.disabled = true;
      }
    }
    if (draft.status === STATUS.CONFIRMADO || draft.status === STATUS.PRERESERVA) {
      el.eventStatus.value = STATUS.LISTA;
    }
  }

  // Estados automaticos quedan bloqueados para seleccion manual.
}

function evaluateRules(draft, ignoreIds = null) {
  // Draft must be parseable
  if (!draft.date || !draft.startTime || !draft.endTime) return { ok: true, hint: "" };

  // Rule: if overlaps with existing Confirmado/Pre reserva,
  // block Confirmado/Pre reserva.
  const hardBlocks = findHardBlocks(draft, ignoreIds);
  const hasHardBlock = hardBlocks.length > 0;
  const maintenanceBlocks = findMaintenanceDayBlocks(draft, ignoreIds);
  const hasMaintenanceDayBlock = maintenanceBlocks.length > 0;

  if (hasMaintenanceDayBlock && draft.status !== STATUS.LISTA && draft.status !== STATUS.MANTENIMIENTO) {
    return {
      ok: false,
      message: "Salon en Mantenimiento este dia. Solo se permite Lista de Espera.",
      hint: "Mantenimiento activo: no se permite Confirmado ni Pre reserva.",
    };
  }

  if (hasHardBlock) {
    if (draft.status === STATUS.MANTENIMIENTO) {
      return {
        ok: false,
        message: "No puedes poner en Mantenimiento: ya hay una reserva Confirmada o Pre reserva en ese horario.",
        hint: "Mantenimiento solo se permite si no hay Confirmado ni Pre reserva cruzados.",
      };
    }
    if (draft.status === STATUS.CONFIRMADO || draft.status === STATUS.PRERESERVA) {
      return {
        ok: false,
        message: "Ya hay un evento Confirmado/Pre reserva en ese horario. Deseas ponerlo en Lista de Espera o cambiar de hora?",
        hint: "Solo puede existir un Confirmado o Pre reserva por horario. Puedes usar Lista de Espera.",
      };
    }
    return {
      ok: true,
      hint: "Hay cruce con un Confirmado/Pre reserva. Recomendado: Lista de Espera.",
    };
  }

  // Rule: Pre reserva se comporta como Confirmado (bloquea igual)
  // (La regla real se aplica por conflictos al guardar)

  // Cancelado: manual (permitido)
  if (draft.status === STATUS.CANCELADO) {
    return { ok: true, hint: "Cancelado: se usa cuando el usuario cancela manualmente." };
  }

  // Default ok
  return { ok: true, hint: "" };
}

function findConflicts(draft, ignoreIds = null) {
  const id = draft.id;
  return state.events.filter(e => {
    if (ignoreIds && ignoreIds.has(String(e.id || ""))) return false;
    if (e.id === id) return false;
    if (e.salon !== draft.salon) return false;
    if (e.date !== draft.date) return false;
    // ignore canceled? (decision: cancelado no bloquea)
    if (e.status === STATUS.CANCELADO) return false;

    return timesOverlap(e.startTime, e.endTime, draft.startTime, draft.endTime);
  });
}

function findHardBlocks(draft, ignoreIds = null) {
  const id = draft.id;
  return state.events.filter(e => {
    if (ignoreIds && ignoreIds.has(e.id)) return false;
    if (e.id === id) return false;
    if (e.salon !== draft.salon) return false;
    if (e.date !== draft.date) return false;
    if (e.status === STATUS.CANCELADO) return false;
    if (e.status !== STATUS.CONFIRMADO && e.status !== STATUS.PRERESERVA) return false;

    return timesOverlap(e.startTime, e.endTime, draft.startTime, draft.endTime);
  });
}

function findMaintenanceDayBlocks(draft, ignoreIds = null) {
  const id = draft.id;
  return state.events.filter(e => {
    if (ignoreIds && ignoreIds.has(e.id)) return false;
    if (e.id === id) return false;
    if (e.salon !== draft.salon) return false;
    if (e.date !== draft.date) return false;
    if (e.status !== STATUS.MANTENIMIENTO) return false;
    return true;
  });
}

function isHardBlockingStatus(status) {
  return status === STATUS.CONFIRMADO || status === STATUS.PRERESERVA;
}

function canBeAutoNotifiedForReleasedCapacity(status) {
  return status === STATUS.LISTA
    || status === STATUS.RESERVA_SIN_COTIZACION
    || status === STATUS.PRIMERA
    || status === STATUS.SEGUIMIENTO;
}

function buildStatusChangeToast(prevStatus, nextStatus, fallback = "Estado actualizado.") {
  const from = String(prevStatus || "").trim();
  const to = String(nextStatus || "").trim();
  if (from && to && from !== to) return `Estado actualizado: ${from} -> ${to}.`;
  if (to) return `Estado guardado: ${to}.`;
  return fallback;
}

function buildBlockingWindowsFromEvents(events) {
  if (!Array.isArray(events) || !events.length) return [];
  return events
    .filter((e) => isHardBlockingStatus(String(e?.status || "").trim()))
    .map((e) => ({
      id: String(e.id || ""),
      date: String(e.date || ""),
      salon: String(e.salon || ""),
      startTime: String(e.startTime || ""),
      endTime: String(e.endTime || ""),
    }))
    .filter((w) => w.date && w.salon && w.startTime && w.endTime);
}

function computeReleasedWindows(previousBlockingWindows, nextBlockingWindows) {
  const prev = Array.isArray(previousBlockingWindows) ? previousBlockingWindows : [];
  const next = Array.isArray(nextBlockingWindows) ? nextBlockingWindows : [];
  if (!prev.length) return [];
  return prev.filter((oldW) => {
    return !next.some((newW) => {
      if (newW.date !== oldW.date) return false;
      if (newW.salon !== oldW.salon) return false;
      return timesOverlap(oldW.startTime, oldW.endTime, newW.startTime, newW.endTime);
    });
  });
}

function summarizeReleasedWindows(windows, maxItems = 3) {
  const rows = (Array.isArray(windows) ? windows : [])
    .filter((w) => w?.date && w?.salon && w?.startTime && w?.endTime)
    .slice(0, maxItems)
    .map((w) => `${w.date} ${w.salon} ${w.startTime}-${w.endTime}`);
  if (!rows.length) return "";
  const extra = (Array.isArray(windows) ? windows.length : 0) - rows.length;
  return extra > 0 ? `${rows.join(", ")} (+${extra} bloque(s))` : rows.join(", ");
}

function createSystemNotificationReminder(ev, notes, createdByUserId = "") {
  if (!ev || !notes) return;
  const now = new Date(Date.now() + 60 * 60 * 1000); // +60 min para mantenerla visible en campana
  addReminderForEvent(ev, {
    date: toISODate(now),
    time: toHHMM(now),
    channel: "Sistema",
    notes,
    createdByUserId,
  });
}

function notifyReleasedCapacityForWaitingReservations({
  releasedWindows,
  sourceEvent,
  reasonText,
  actorUserId = "",
}) {
  try {
    const windows = (Array.isArray(releasedWindows) ? releasedWindows : [])
      .filter((w) => w && typeof w === "object")
      .map((w) => ({
        date: String(w.date || "").trim(),
        salon: String(w.salon || "").trim(),
        startTime: String(w.startTime || "").trim(),
        endTime: String(w.endTime || "").trim(),
      }))
      .filter((w) => w.date && w.salon && isValidClockTime(w.startTime) && isValidClockTime(w.endTime));
    if (!windows.length) return 0;

    const sourceId = String(sourceEvent?.id || "");
    const sourceGroupId = String(sourceEvent?.groupId || "");
    const releaseSummary = summarizeReleasedWindows(windows);
    const sourceLabel = String(sourceEvent?.name || "una reserva");
    const notes = `Se libero horario (${releaseSummary}) porque "${sourceLabel}" ${reasonText}. Ya puedes intentar mover tu reserva a Confirmado/Pre reserva.`;

    addGlobalNotification({
      title: "Horario disponible",
      notes,
      eventId: sourceId,
      salon: String(windows[0]?.salon || sourceEvent?.salon || ""),
      date: String(windows[0]?.date || sourceEvent?.date || ""),
      time: String(windows[0]?.startTime || sourceEvent?.startTime || ""),
      releasedWindows: windows,
    });
    toast("Aviso global enviado: se libero un horario.");

    const notifiedByReservation = new Map();
    for (const ev of state.events || []) {
      if (!ev) continue;
      if (String(ev.id || "") === sourceId) continue;
      if (sourceGroupId && String(ev.groupId || "") === sourceGroupId) continue;
      if (!canBeAutoNotifiedForReleasedCapacity(String(ev.status || "").trim())) continue;
      if (!isValidClockTime(String(ev.startTime || "")) || !isValidClockTime(String(ev.endTime || ""))) continue;

      const hasOverlap = windows.some((w) => {
        if (String(ev.date || "") !== w.date) return false;
        if (String(ev.salon || "") !== w.salon) return false;
        return timesOverlap(String(ev.startTime || ""), String(ev.endTime || ""), w.startTime, w.endTime);
      });
      if (!hasOverlap) continue;

      const key = String(reservationKeyFromEvent(ev) || "").trim();
      if (!key || notifiedByReservation.has(key)) continue;
      notifiedByReservation.set(key, ev);
    }

    for (const ev of notifiedByReservation.values()) {
      createSystemNotificationReminder(ev, notes, actorUserId || sourceEvent?.userId || "");
      appendHistoryByKey(
        reservationKeyFromEvent(ev),
        actorUserId || sourceEvent?.userId || "",
        `Aviso automatico: horario liberado (${releaseSummary}) porque "${sourceLabel}" ${reasonText}.`
      );
      if (String(authSession.userId || "").trim() === String(ev.userId || "").trim()) {
        toast(`Aviso: se libero horario para tu reserva (${ev.date} ${ev.salon}).`);
      }
    }
    return notifiedByReservation.size;
  } catch (err) {
    console.error("Fallo al generar aviso de horario liberado:", err);
    return 0;
  }
}

function autoMarkLostEvents() {
  const now = new Date();
  let changed = false;

  for (const e of state.events) {
    if (e.status === STATUS.CANCELADO) continue;
    if (e.status === STATUS.CONFIRMADO) continue;
    if (e.status === STATUS.MANTENIMIENTO) continue;

    const end = new Date(`${e.date}T${e.endTime}:00`);
    if (end < now) {
      if (e.status !== STATUS.PERDIDO) {
        e.status = STATUS.PERDIDO;
        changed = true;
      }
    }
  }

  if (changed) persist();
}

// ================== State ==================

function persist() {
  if (!serverStateReady) {
    pendingPersistAfterSync = true;
    syncWithServerState();
    return;
  }
  schedulePersistToServer();
}

function normalizeUserRecord(candidate) {
  const u = candidate && typeof candidate === "object" ? candidate : {};
  const fullName = String(u.fullName || u.name || "").trim();
  const monthlyGoals = Array.isArray(u.monthlyGoals) ? u.monthlyGoals : [];
  const normalizedMonthlyGoals = monthlyGoals
    .map((g) => ({
      month: String(g?.month || "").trim(),
      amount: Math.max(0, Number(g?.amount || 0)),
    }))
    .filter((g) => /^\d{4}-\d{2}$/.test(g.month))
    .sort((a, b) => a.month.localeCompare(b.month));
  return {
    ...u,
    id: String(u.id || "").trim(),
    name: fullName,
    fullName,
    role: normalizeUserRole(u.role),
    username: String(u.username || "").trim(),
    email: String(u.email || "").trim(),
    phone: String(u.phone || "").trim(),
    password: String(u.password || "").trim(),
    signatureDataUrl: String(u.signatureDataUrl || "").trim(),
    avatarDataUrl: String(u.avatarDataUrl || "").trim(),
    active: u.active !== false,
    salesTargetEnabled: u.salesTargetEnabled === true,
    monthlyGoals: normalizedMonthlyGoals,
  };
}

function normalizeState(candidate) {
  if (!candidate || typeof candidate !== "object") return null;
  const salones = Array.isArray(candidate.salones) ? candidate.salones : [];
  const hasRemoteTemplates = Object.prototype.hasOwnProperty.call(candidate, "quickTemplates");
  const fallbackLocalTemplates = loadQuickTemplates();
  const templateSource = hasRemoteTemplates ? candidate.quickTemplates : fallbackLocalTemplates;
  const normalizeEventSalonNames = (value, fallbackSalon = "") => {
    const merged = [...(Array.isArray(value) ? value : []), fallbackSalon]
      .map((name) => String(name || "").trim())
      .filter(Boolean);
    return Array.from(new Set(merged));
  };

  const normalized = {
    ...candidate,
    salones,
    events: Array.isArray(candidate.events)
      ? candidate.events.map((e) => {
        const salon = String(e?.salon ?? "").trim();
        const eventSalones = normalizeEventSalonNames(e?.salones, salon);
        const mainSalon = String(e?.mainSalon || eventSalones[0] || salon).trim();
        return {
          ...e,
          salon,
          mainSalon: mainSalon || salon,
          salones: eventSalones,
        };
      })
      : [],
    users: Array.isArray(candidate.users) ? candidate.users.map(normalizeUserRecord).filter((u) => u.id && u.name) : [],
    companies: Array.isArray(candidate.companies) ? candidate.companies : [],
    services: Array.isArray(candidate.services) ? candidate.services.map(normalizeServiceRecord) : [],
    quickTemplates: ensureCorporateTemplateSeed(Array.isArray(templateSource) ? templateSource : []),
    quoteServiceTemplates: normalizeQuoteServiceTemplatesCollection(candidate.quoteServiceTemplates),
    disabledCompanies: Array.isArray(candidate.disabledCompanies) ? candidate.disabledCompanies.map((x) => String(x || "").trim()).filter(Boolean) : [],
    disabledServices: Array.isArray(candidate.disabledServices) ? candidate.disabledServices.map((x) => String(x || "").trim()).filter(Boolean) : [],
    disabledManagers: Array.isArray(candidate.disabledManagers) ? candidate.disabledManagers.map((x) => String(x || "").trim()).filter(Boolean) : [],
    disabledSalones: Array.isArray(candidate.disabledSalones) ? candidate.disabledSalones.map((x) => String(x || "").trim()).filter(Boolean) : [],
    globalMonthlyGoals: Array.isArray(candidate.globalMonthlyGoals) ? candidate.globalMonthlyGoals : [],
    changeHistory: (candidate.changeHistory && typeof candidate.changeHistory === "object") ? candidate.changeHistory : {},
    reminders: (candidate.reminders && typeof candidate.reminders === "object") ? candidate.reminders : {},
    checklistTemplates: Array.isArray(candidate.checklistTemplates) ? candidate.checklistTemplates : [],
    checklistTemplateItems: Array.isArray(candidate.checklistTemplateItems) ? candidate.checklistTemplateItems : [],
    checklistTemplateSections: Array.isArray(candidate.checklistTemplateSections) ? candidate.checklistTemplateSections : [],
    menuMontajeSections: Array.isArray(candidate.menuMontajeSections) ? candidate.menuMontajeSections : [],
    menuMontajeBebidas: Array.isArray(candidate.menuMontajeBebidas) ? candidate.menuMontajeBebidas : [],
    eventChecklists: (candidate.eventChecklists && typeof candidate.eventChecklists === "object") ? candidate.eventChecklists : {},
    globalNotifications: Array.isArray(candidate.globalNotifications)
      ? candidate.globalNotifications.map(normalizeGlobalNotificationRecord)
      : [],
  };
  normalized.companies = normalized.companies.map(normalizeCompanyRecord);
  return normalized;
}

function buildInitialState() {
  const templateSeed = ensureCorporateTemplateSeed(loadQuickTemplates());
  return {
    salones: SALONES_DEFAULT.slice(),
    users: USERS_DEFAULT.slice(),
    companies: COMPANIES_DEFAULT.slice(),
    services: SERVICES_DEFAULT.slice(),
    quickTemplates: Array.isArray(templateSeed) ? templateSeed : ensureCorporateTemplateSeed([]),
    quoteServiceTemplates: [],
    disabledCompanies: [],
    disabledServices: [],
    disabledManagers: [],
    disabledSalones: [],
    globalMonthlyGoals: [],
    changeHistory: {},
    reminders: {},
    checklistTemplates: [],
    checklistTemplateItems: [],
    checklistTemplateSections: ["General"],
    menuMontajeSections: ["General"],
    menuMontajeBebidas: [],
    eventChecklists: {},
    globalNotifications: [],
    events: [],
  };
}

function schedulePersistToServer() {
  clearTimeout(persistServerTimer);
  persistServerTimer = setTimeout(() => {
    persistToServer().catch(() => { });
  }, API_SYNC_DEBOUNCE_MS);
}

async function persistToServer() {
  if (persistInFlight) {
    persistQueued = true;
    return;
  }
  persistInFlight = true;
  try {
    const response = await fetch(activeApiStateUrl, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ state }),
    });
    if (!response.ok) {
      let detail = "";
      try {
        const payload = await response.json();
        detail = payload?.detail || payload?.message || "";
      } catch (_) { }
      throw new Error(detail || `save_failed_${response.status}`);
    }
    saveErrorNotified = false;
  } catch (_) {
    if (!saveErrorNotified) {
      saveErrorNotified = true;
      toast("No se pudo guardar en MariaDB (revisa consola).");
    }
    console.error("Persistencia MariaDB fallida:", _.message || _);
  } finally {
    persistInFlight = false;
    if (persistQueued) {
      persistQueued = false;
      schedulePersistToServer();
    }
  }
}

async function syncWithServerState() {
  if (syncInFlight) return;
  syncInFlight = true;
  try {
    let response = null;
    for (const candidate of API_STATE_CANDIDATES) {
      try {
        const r = await fetch(candidate, { cache: "no-store" });
        const contentType = String(r.headers.get("content-type") || "").toLowerCase();
        const isJson404 = r.status === 404 && contentType.includes("application/json");
        if (r.ok || isJson404) {
          response = r;
          activeApiStateUrl = candidate;
          break;
        }
      } catch (_) { }
    }

    if (!response) {
      serverStateReady = false;
      toast("Sin conexion con MariaDB API.");
      return;
    }

    if (response.status === 404) {
      serverStateReady = true;
      await persistToServer();
      pendingPersistAfterSync = false;
      toast("Base inicial guardada en MariaDB");
      return;
    }

    const payload = await response.json();
    const normalized = normalizeState(payload?.state);
    if (!normalized) {
      serverStateReady = false;
      toast("Estado remoto invalido.");
      return;
    }

    state = normalized;
    quickTemplates = Array.isArray(state.quickTemplates) ? state.quickTemplates : [];
    backupQuickTemplatesLocal();
    serverStateReady = true;
    await syncRoomsFromDb();
    await syncServiceCatalogFromDb();
    autoMarkLostEvents();
    renderRoomSelects();
    renderUsersSelect();
    renderCompaniesSelect();
    renderServicesList();
    render();
    runUpcomingReminderChecks();
    refreshTopbarReminders();
    if (pendingPersistAfterSync) {
      pendingPersistAfterSync = false;
      schedulePersistToServer();
    }
    toast("Datos cargados");
  } catch (_) {
    serverStateReady = false;
    toast("Sin conexion con servidor.");
  } finally {
    syncInFlight = false;
  }
}

async function syncRoomsFromDb() {
  try {
    const salonesUrl = buildApiUrlFromStateUrl(activeApiStateUrl, "salones");
    const response = await fetch(salonesUrl, { cache: "no-store" });
    if (!response.ok) return;
    const payload = await response.json();
    const list = Array.isArray(payload?.salones) ? payload.salones.map(x => String(x || "").trim()).filter(Boolean) : [];
    if (!list.length) return;
    state.salones = list;
  } catch (_) { }
}

// ================== Helpers ==================

function uid() {
  return "ev_" + Math.random().toString(16).slice(2) + "_" + Date.now().toString(16);
}

function startOfWeek(date) {
  // Monday as start
  const d = stripTime(date);
  const day = d.getDay(); // 0 Sun..6 Sat
  const diff = (day === 0 ? -6 : 1) - day; // move to Monday
  return addDays(d, diff);
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function addMonths(date, months) {
  const d = new Date(date);
  d.setMonth(d.getMonth() + months);
  return d;
}

function startOfMonth(date) {
  const d = new Date(date);
  d.setDate(1);
  d.setHours(0, 0, 0, 0);
  return d;
}

function stripTime(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  return d;
}

function isSameDay(a, b) {
  return stripTime(a).getTime() === stripTime(b).getTime();
}

function fmtWeekday(d) {
  return d.toLocaleDateString("es-GT", { weekday: "short" }).replace(".", "").toUpperCase();
}
function fmtDayMonth(d) {
  return d.toLocaleDateString("es-GT", { day: "2-digit", month: "short" }).toUpperCase();
}
function fmtDateShort(d) {
  return d.toLocaleDateString("es-GT", { day: "2-digit", month: "short", year: "numeric" });
}
function fmtMonthYear(d) {
  return d.toLocaleDateString("es-GT", { month: "long", year: "numeric" });
}
function parseQuoteDate(raw) {
  if (!raw) return null;
  let d = new Date(raw);
  if (!Number.isNaN(d.getTime())) return stripTime(d);
  d = new Date(`${raw}T00:00:00`);
  return Number.isNaN(d.getTime()) ? null : stripTime(d);
}
function buildFollowUpLabel(ev) {
  if (!ev || ev.status !== STATUS.SEGUIMIENTO) return "";
  const quotedDate = parseQuoteDate(ev.quote?.quotedAt || ev.quote?.quotedDate);
  if (!quotedDate) return "Seguimiento";
  const today = stripTime(new Date());
  const diffMs = today.getTime() - quotedDate.getTime();
  const days = Math.max(0, Math.floor(diffMs / (1000 * 60 * 60 * 24)));
  const dayWord = days === 1 ? "dia" : "dias";
  return `Seguimiento hace ${days} ${dayWord} - Cotizado ${fmtDateShort(quotedDate)}`;
}
function toISODate(d) {
  const x = new Date(d);
  const y = x.getFullYear();
  const m = pad2(x.getMonth() + 1);
  const da = pad2(x.getDate());
  return `${y}-${m}-${da}`;
}

function toHHMM(d) {
  const x = new Date(d);
  return `${pad2(x.getHours())}:${pad2(x.getMinutes())}`;
}

function pad2(n) { return String(n).padStart(2, "0"); }

function formatHourAmPm(h24) {
  const h = Number(h24) % 24;
  const suffix = h >= 12 ? "PM" : "AM";
  const h12 = (h % 12) || 12;
  return `${h12}:00 ${suffix}`;
}

function timeToY(t) {
  const [hh, mm] = t.split(":").map(Number);
  const minutes = (hh - HOUR_START) * 60 + mm;
  return (minutes / 60) * HOUR_HEIGHT;
}

function timeToMinutes(t) {
  const [hh, mm] = t.split(":").map(Number);
  return hh * 60 + mm;
}

function minutesToTime(minuteFromGridStart) {
  const maxGridMinute = maxSelectableMinutes();
  const safeMinute = clamp(minuteFromGridStart, 0, maxGridMinute);
  const base = HOUR_START * 60 + safeMinute;
  const hh = Math.floor(base / 60);
  const mm = base % 60;
  return `${pad2(hh)}:${pad2(mm)}`;
}

function compareTime(a, b) {
  // returns >0 if a > b
  const [ah, am] = a.split(":").map(Number);
  const [bh, bm] = b.split(":").map(Number);
  return (ah * 60 + am) - (bh * 60 + bm);
}

function isValidClockTime(raw) {
  const m = String(raw || "").trim().match(/^(\d{2}):(\d{2})$/);
  if (!m) return false;
  const hh = Number(m[1]);
  const mm = Number(m[2]);
  return Number.isInteger(hh) && Number.isInteger(mm) && hh >= 0 && hh <= 23 && mm >= 0 && mm <= 59;
}

function modernAlert({ icon = "info", title = "", text = "", html = "", confirmText = "Aceptar" }) {
  if (window.Swal && typeof window.Swal.fire === "function") {
    return window.Swal.fire({
      icon,
      title,
      text: text || undefined,
      html: html || undefined,
      confirmButtonText: confirmText,
      background: "#f8fbff",
      color: "#10243b",
      confirmButtonColor: "#2563eb",
      allowOutsideClick: false,
    });
  }
  toast(text || title || "Aviso");
  return Promise.resolve({ isConfirmed: true });
}

function modernGuideToast(message) {
  if (window.Swal && typeof window.Swal.fire === "function") {
    return window.Swal.fire({
      toast: true,
      position: "top",
      icon: "warning",
      title: message,
      showConfirmButton: false,
      timer: 2600,
      timerProgressBar: true,
      customClass: { popup: "swal-ios-toast" },
    });
  }
  toast(message);
  return Promise.resolve();
}

async function modernConfirm({ title = "Confirmar", message = "", confirmText = "Confirmar", cancelText = "Cancelar" } = {}) {
  if (window.Swal && typeof window.Swal.fire === "function") {
    const result = await window.Swal.fire({
      icon: "warning",
      title,
      text: message || undefined,
      showCancelButton: true,
      confirmButtonText: confirmText,
      cancelButtonText: cancelText,
      background: "#f8fbff",
      color: "#10243b",
      confirmButtonColor: "#2563eb",
      cancelButtonColor: "#94a3b8",
    });
    return !!result?.isConfirmed;
  }
  return !!window.confirm(message || title);
}

function modernConfirmMaintenance() {
  if (window.Swal && typeof window.Swal.fire === "function") {
    return window.Swal.fire({
      icon: "warning",
      title: "Confirmar mantenimiento",
      text: "Esta seguro de ponerla en Mantenimiento?",
      showCancelButton: true,
      confirmButtonText: "Si, poner en mantenimiento",
      cancelButtonText: "No, cancelar",
      background: "#f8fbff",
      color: "#10243b",
      confirmButtonColor: "#8b5cf6",
      cancelButtonColor: "#94a3b8",
    });
  }
  return Promise.resolve({ isConfirmed: window.confirm("Esta seguro de ponerla en Mantenimiento?") });
}

function modernConfirmReleaseMaintenance() {
  if (window.Swal && typeof window.Swal.fire === "function") {
    return window.Swal.fire({
      icon: "warning",
      title: "Liberar mantenimiento",
      text: "Estas seguro de quitar mantenimiento?",
      showCancelButton: true,
      confirmButtonText: "Si, liberar",
      cancelButtonText: "No, cancelar",
      background: "#f8fbff",
      color: "#10243b",
      confirmButtonColor: "#dc2626",
      cancelButtonColor: "#94a3b8",
    });
  }
  return Promise.resolve({ isConfirmed: window.confirm("Estas seguro de quitar mantenimiento?") });
}

function getEventSeriesLastDate(ev) {
  const series = getEventSeries(ev);
  if (!series.length) return String(ev?.date || "");
  return series.reduce((maxDate, item) => {
    const d = String(item?.date || "");
    return d > maxDate ? d : maxDate;
  }, String(series[0]?.date || ev?.date || ""));
}

function isEventSeriesInPast(ev) {
  const lastDate = getEventSeriesLastDate(ev);
  if (!lastDate) return false;
  const todayIso = toISODate(new Date());
  return lastDate < todayIso;
}

function hasPastEventEditAuthorization(ev) {
  const key = String(reservationKeyFromEvent(ev) || "").trim();
  return !!key && pastEventEditAuthorizedKeys.has(key);
}

async function requestPastEventEditAuthorization(ev) {
  const key = String(reservationKeyFromEvent(ev) || "").trim();
  if (!key) return false;
  if (pastEventEditAuthorizedKeys.has(key)) return true;

  let code = "";
  if (window.Swal && typeof window.Swal.fire === "function") {
    const result = await window.Swal.fire({
      icon: "warning",
      title: "Evento de fecha pasada",
      text: "Para editar este evento ingresa codigo de administrador.",
      input: "password",
      inputPlaceholder: "Codigo admin",
      showCancelButton: true,
      confirmButtonText: "Autorizar",
      cancelButtonText: "Cancelar",
      background: "#f8fbff",
      color: "#10243b",
      confirmButtonColor: "#2563eb",
      inputValidator: (value) => {
        if (!String(value || "").trim()) return "Ingresa el codigo.";
        return null;
      },
    });
    if (!result.isConfirmed) return false;
    code = String(result.value || "").trim();
  } else {
    code = String(window.prompt("Codigo de administrador para editar evento pasado:", "") || "").trim();
    if (!code) return false;
  }

  if (code !== PAST_EVENT_ADMIN_EDIT_CODE) {
    await modernAlert({
      icon: "error",
      title: "Codigo invalido",
      text: "No tienes autorizacion para editar eventos de fechas pasadas.",
    });
    return false;
  }
  pastEventEditAuthorizedKeys.add(key);
  return true;
}

function validateReservationRequiredFields() {
  const issues = [];
  let firstInvalidEl = null;
  const mark = (el, ok) => {
    if (!el) return;
    el.classList.toggle("req-missing", !ok);
    el.classList.toggle("req-ok", !!ok);
    if (el.tagName === "SELECT") {
      const visual = el.nextElementSibling;
      if (visual && visual.classList?.contains("ss-main")) {
        visual.classList.toggle("req-missing", !ok);
        visual.classList.toggle("req-ok", !!ok);
      }
    }
    if (!ok && !firstInvalidEl) firstInvalidEl = el;
  };
  const name = String(el.eventName?.value || "").trim();
  const dateStart = String(el.eventDate?.value || "").trim();
  const dateEnd = String(el.eventDateEnd?.value || "").trim();
  const userId = String(el.eventUser?.value || "").trim();
  const slots = getSlotsFromForm();

  const okName = !!name;
  mark(el.eventName, okName);
  if (!okName) issues.push("Nombre del evento");
  const okDateStart = !!dateStart;
  mark(el.eventDate, okDateStart);
  if (!okDateStart) issues.push("Fecha inicial");
  const okDateEnd = !!dateEnd;
  mark(el.eventDateEnd, okDateEnd);
  if (!okDateEnd) issues.push("Fecha final");
  const okUser = !!userId;
  mark(el.eventUser, okUser);
  if (!okUser) issues.push("Usuario");
  const okPax = syncEventPaxFromSlots() > 0;
  mark(el.eventPax, okPax);
  if (!okPax) issues.push("PAX total");
  if (!slots.length) issues.push("Al menos un bloque de salon y horario");

  const rows = Array.from(el.slotsBody.querySelectorAll(".slotRow"));
  for (let i = 0; i < slots.length; i++) {
    const s = slots[i];
    const idx = i + 1;
    const row = rows[i];
    const roomEl = row?.querySelector(".slotRoom");
    const paxEl = row?.querySelector(".slotPax");
    const slotDateStartEl = row?.querySelector(".slotDateStart");
    const slotDateEndEl = row?.querySelector(".slotDateEnd");
    const startEl = row?.querySelector(".slotStart");
    const endEl = row?.querySelector(".slotEnd");
    const okRoom = !!s.salon;
    mark(roomEl, okRoom);
    if (!okRoom) issues.push(`Bloque ${idx}: Salon`);
    const okSlotPax = Number(s.slotPax) > 0;
    mark(paxEl, okSlotPax);
    if (!okSlotPax) issues.push(`Bloque ${idx}: PAX`);
    const slotRange = normalizeSlotDateRange(s, dateStart, dateEnd);
    const okSlotDateStart = !!slotRange.start;
    const okSlotDateEnd = !!slotRange.end;
    mark(slotDateStartEl, okSlotDateStart);
    mark(slotDateEndEl, okSlotDateEnd);
    if (!okSlotDateStart) issues.push(`Bloque ${idx}: Fecha desde`);
    if (!okSlotDateEnd) issues.push(`Bloque ${idx}: Fecha hasta`);
    if (slotRange.start && slotRange.end && dateStart && dateEnd) {
      const eventRange = normalizeSlotDateRange({ dateStart, dateEnd });
      const inEventRange = slotRange.start >= eventRange.start && slotRange.end <= eventRange.end;
      mark(slotDateStartEl, inEventRange);
      mark(slotDateEndEl, inEventRange);
      if (!inEventRange) issues.push(`Bloque ${idx}: Fechas fuera del rango del evento`);
    }
    const okStart = !!s.startTime && isValidClockTime(s.startTime);
    mark(startEl, okStart);
    if (!s.startTime) issues.push(`Bloque ${idx}: Hora inicio`);
    const okEnd = !!s.endTime && isValidClockTime(s.endTime);
    mark(endEl, okEnd);
    if (!s.endTime) issues.push(`Bloque ${idx}: Hora fin`);
    if (s.startTime && s.endTime) {
      if (!isValidClockTime(s.startTime) || !isValidClockTime(s.endTime)) {
        issues.push(`Bloque ${idx}: Formato de hora invalido`);
        if (!isValidClockTime(s.startTime)) mark(startEl, false);
        if (!isValidClockTime(s.endTime)) mark(endEl, false);
      } else if (compareTime(s.endTime, s.startTime) <= 0) {
        issues.push(`Bloque ${idx}: Hora fin debe ser mayor a inicio`);
        mark(endEl, false);
      }
    }
  }

  return { issues, firstInvalidEl };
}

function timesOverlap(aStart, aEnd, bStart, bEnd) {
  // overlap if start < otherEnd AND end > otherStart
  if (!isValidClockTime(aStart) || !isValidClockTime(aEnd) || !isValidClockTime(bStart) || !isValidClockTime(bEnd)) {
    return false;
  }
  return compareTime(aStart, bEnd) < 0 && compareTime(aEnd, bStart) > 0;
}

function clamp(n, min, max) { return Math.max(min, Math.min(max, n)); }

function toast(msg) {
  const text = String(msg || "").trim();
  if (window.Swal && typeof window.Swal.fire === "function") {
    const successPattern = /\b(agregado|agregada|guardado|guardada|creado|creada|cargado|cargada|movido|movida|eliminado|eliminada|ajustada|ajustado|extendida|extendido|cambiado|cambiada|actualizado|actualizada|listo)\b/i;
    const errorPattern = /\b(no se pudo|error|invalido|invalida|falta|faltan|completa|sin conexion|debe|obligatoria|obligatorio)\b/i;
    const icon = errorPattern.test(text) ? "error" : (successPattern.test(text) ? "success" : "info");
    const iconHtml = icon === "success" ? "&#10003;" : (icon === "error" ? "&#215;" : "i");
    window.Swal.fire({
      toast: true,
      position: "top",
      icon,
      iconHtml,
      title: text,
      showConfirmButton: false,
      timer: 2400,
      timerProgressBar: true,
      customClass: { popup: "swal-ios-toast" },
    });
    return;
  }
  el.toast.textContent = text;
  el.toast.classList.add("show");
  clearTimeout(toast._t);
  toast._t = setTimeout(() => el.toast.classList.remove("show"), 2400);
}

function escapeHtml(s) {
  return String(s)
    .split("&").join("&amp;")
    .split("<").join("&lt;")
    .split(">").join("&gt;")
    .split('"').join("&quot;")
    .split("'").join("&#039;");
}

function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function isValidEmail(email) {
  const value = String(email || "").trim();
  if (!value || value.length > 254 || /\.\./.test(value)) return false;
  return /^[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}$/i.test(value);
}

function normalizeCompanyRecord(company) {
  const c = company || {};
  const managers = Array.isArray(c.managers) ? c.managers : [];
  const normalizedManagers = managers
    .map(m => ({
      id: m.id || uid(),
      name: String(m.name || "").trim(),
      phone: String(m.phone || "").trim(),
      email: String(m.email || "").trim(),
      address: String(m.address || "").trim(),
    }))
    .filter(m => m.name);

  if (!normalizedManagers.length) {
    const fallbackName = String(c.contact || c.owner || "Encargado");
    normalizedManagers.push({
      id: uid(),
      name: fallbackName,
      phone: String(c.phone || "").trim(),
      email: String(c.email || "").trim(),
      address: String(c.address || "").trim(),
    });
  }

  return {
    id: c.id || uid(),
    name: String(c.name || "Empresa").trim(),
    owner: String(c.owner || c.contact || "").trim(),
    email: String(c.email || "").trim(),
    nit: String(c.nit || "CF").trim(),
    businessName: String(c.businessName || c.billTo || c.name || "").trim(),
    billTo: String(c.billTo || c.businessName || c.name || "").trim(),
    eventType: String(c.eventType || "Social").trim(),
    address: String(c.address || "").trim(),
    phone: String(c.phone || "").trim(),
    notes: String(c.notes || "").trim(),
    managers: normalizedManagers,
  };
}

function currentDraftFromForm() {
  const date = el.eventDate.value;
  const firstSlot = getSlotsFromForm()[0] || {};
  const startTime = firstSlot.startTime || "";
  const endTime = firstSlot.endTime || "";
  const salon = firstSlot.salon || "";
  const status = el.eventStatus.value;
  const id = el.eventId.value || "__draft__";
  if (!date || !startTime || !endTime || !salon || !status) return null;
  return { id, date, startTime, endTime, salon, status };
}

function findMaintenanceBlockingReservationsFromForm() {
  const editingId = String(el.eventId?.value || "").trim();
  const replaceEvents = editingId
    ? getEventSeries(state.events.find(x => x.id === editingId))
    : [];
  const replaceIds = new Set(replaceEvents.map((e) => e.id));
  const dateStart = String(el.eventDate?.value || "").trim();
  const dateEndRaw = String(el.eventDateEnd?.value || "").trim();
  const dateEnd = dateEndRaw || dateStart;
  if (!dateStart || !dateEnd) return [];
  const rangeStart = dateStart <= dateEnd ? dateStart : dateEnd;
  const rangeEnd = dateStart <= dateEnd ? dateEnd : dateStart;
  const slots = getSlotsFromForm().filter((slot) =>
    slot?.salon
    && isValidClockTime(slot.startTime)
    && isValidClockTime(slot.endTime)
    && compareTime(slot.endTime, slot.startTime) > 0
  );
  if (!slots.length) return [];

  const seen = new Set();
  const blocks = [];
  for (const date of listDatesBetween(rangeStart, rangeEnd)) {
    for (const slot of slots) {
      const draft = {
        id: editingId || "__maintenance__",
        date,
        salon: slot.salon,
        startTime: slot.startTime,
        endTime: slot.endTime,
        status: STATUS.MANTENIMIENTO,
      };
      for (const conflict of findHardBlocks(draft, replaceIds)) {
        const key = String(conflict.id || "");
        if (seen.has(key)) continue;
        seen.add(key);
        blocks.push(conflict);
      }
    }
  }
  return blocks;
}

function normalizeServiceRecord(service) {
  const s = service || {};
  const modeRaw = String(s.quantityMode || s.modoCantidad || s.indicadorCantidad || "MANUAL").trim().toUpperCase();
  const quantityMode = modeRaw === "PAX" ? "PAX" : "MANUAL";
  return {
    id: s.id || uid(),
    name: String(s.name || "").trim(),
    price: Number(s.price || 0),
    description: String(s.description || "").trim(),
    categoryId: s.categoryId ?? s.idCategoria ?? null,
    subcategoryId: s.subcategoryId ?? s.idSubcategoria ?? null,
    category: String(s.category || s.categoria || "").trim(),
    subcategory: String(s.subcategory || s.subcategoria || "").trim(),
    quantityMode,
  };
}

// ================== Avatars (SVG) ==================

function avatarDataUri(name) {
  const svg = avatarSvg(name);
  return "data:image/svg+xml;utf8," + encodeURIComponent(svg);
}
function avatarSvg(name) {
  const initials = getInitials(name);
  const { a, b } = gradientFromString(name);
  return `
  <svg xmlns="http://www.w3.org/2000/svg" width="80" height="80" viewBox="0 0 80 80">
    <defs>
      <linearGradient id="g" x1="0" y1="0" x2="1" y2="1">
        <stop offset="0" stop-color="${a}" />
        <stop offset="1" stop-color="${b}" />
      </linearGradient>
      <filter id="s" x="-20%" y="-20%" width="140%" height="140%">
        <feDropShadow dx="0" dy="6" stdDeviation="6" flood-color="rgba(0,0,0,0.35)"/>
      </filter>
    </defs>
    <rect x="6" y="6" width="68" height="68" rx="18" fill="url(#g)" filter="url(#s)" />
    <circle cx="58" cy="22" r="10" fill="rgba(255,255,255,0.22)"/>
    <circle cx="24" cy="56" r="14" fill="rgba(0,0,0,0.12)"/>
    <text x="40" y="47" text-anchor="middle" font-family="ui-sans-serif,system-ui" font-size="26" font-weight="900" fill="rgba(255,255,255,0.95)">
      ${escapeHtml(initials)}
    </text>
  </svg>`;
}

function getInitials(name) {
  const parts = String(name).trim().split(/\s+/).filter(Boolean);
  if (parts.length === 0) return "U";
  if (parts.length === 1) return parts[0].slice(0, 2).toUpperCase();
  return (parts[0][0] + parts[1][0]).toUpperCase();
}

function gradientFromString(s) {
  // deterministic colors
  let h = 0;
  for (let i = 0; i < s.length; i++) {
    h = (h * 31 + s.charCodeAt(i)) >>> 0;
  }
  const c1 = hslToHex((h % 360), 80, 55);
  const c2 = hslToHex(((h * 3) % 360), 80, 45);
  return { a: c1, b: c2 };
}

function hslToHex(h, s, l) {
  s /= 100; l /= 100;
  const k = n => (n + h / 30) % 12;
  const a = s * Math.min(l, 1 - l);
  const f = n => l - a * Math.max(-1, Math.min(k(n) - 3, Math.min(9 - k(n), 1)));
  const r = Math.round(255 * f(0));
  const g = Math.round(255 * f(8));
  const b = Math.round(255 * f(4));
  return "#" + [r, g, b].map(x => x.toString(16).padStart(2, "0")).join("");
}

function hexToRgba(hex, alpha) {
  // accepts #rrggbb
  const h = String(hex || "").trim();
  if (!h.startsWith("#") || h.length !== 7) return `rgba(255,255,255,${alpha})`;
  const r = parseInt(h.slice(1, 3), 16);
  const g = parseInt(h.slice(3, 5), 16);
  const b = parseInt(h.slice(5, 7), 16);
  return `rgba(${r},${g},${b},${alpha})`;
}

// ================== Utilities ==================

function getEventsInWeek(weekStart, salon, dayCount = 7) {
  const start = stripTime(weekStart).getTime();
  const end = stripTime(addDays(weekStart, dayCount)).getTime(); // exclusive
  return state.events
    .filter(e => salon === ALL_ROOMS_VALUE || e.salon === salon)
    .filter(e => {
      const t = stripTime(new Date(e.date + "T00:00:00")).getTime();
      return t >= start && t < end;
    })
    // show cancelled too (but they won't block)
    .sort((a, b) => {
      if (a.date !== b.date) return a.date.localeCompare(b.date);
      return compareTime(a.startTime, b.startTime);
    });
}














































































































































