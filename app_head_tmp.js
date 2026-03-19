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
  PRIMERA: "1er Cotizacion",
  SEGUIMIENTO: "Seguimiento",
};

const STATUS_META = [
  { key: STATUS.PRIMERA, colorVar: "--c-primera" },
  { key: STATUS.PERDIDO, colorVar: "--c-perdido" },
  { key: STATUS.SEGuimiento ?? STATUS.SEGUIMIENTO, colorVar: "--c-seguimiento" }, // safety
  { key: STATUS.LISTA, colorVar: "--c-lista" },
  { key: STATUS.PRERESERVA, colorVar: "--c-prereserva" },
  { key: STATUS.CONFIRMADO, colorVar: "--c-confirmado" },
  { key: STATUS.CANCELADO, colorVar: "--c-cancelado" },
  { key: STATUS.MANTENIMIENTO, colorVar: "--c-mantenimiento" },
].map(x => ({ ...x, key: x.key === undefined ? STATUS.SEGUIMIENTO : x.key }));

const AUTO_STATUSES = new Set([STATUS.PRIMERA, STATUS.SEGUIMIENTO, STATUS.PERDIDO]);
const DASHBOARD_STATUS_ORDER = [
  STATUS.PERDIDO,
  STATUS.SEGUIMIENTO,
  STATUS.PRIMERA,
  STATUS.CANCELADO,
  STATUS.CONFIRMADO,
  STATUS.PRERESERVA,
  STATUS.LISTA,
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
const CORPORATE_TEMPLATE_ID = "tpl-corporativo";
const CORPORATE_TEMPLATE_NAME = "Corporativo";
const SERVIHOSP_TEMPLATE_ID = "tpl-servi-hosp";
const SERVIHOSP_TEMPLATE_NAME = "Servi Hosp";
const CONTRACT_CORP_TEMPLATE_ID = "tpl-contrato-corp";
const CONTRACT_CORP_TEMPLATE_NAME = "Contrato Corporativo";
const DEFAULT_TEMPLATE_HEADER_IMAGE = "./Encabezadojdl.png";
const DEFAULT_TEMPLATE_FOOTER_IMAGE = "./piedepaginajdl.png";
const SERVIHOSP_TEMPLATE_HEADER_IMAGE = "./ServiHosp_header.png";
const TEMPLATE_SIGNATURE_MIN_W_PCT = 10;
const TEMPLATE_SIGNATURE_MIN_H_PCT = 3;
const TEMPLATE_SIGNATURE_MAX_W_PCT = 35;
const TEMPLATE_SIGNATURE_MAX_H_PCT = 12;
const TEMPLATE_SIGNATURE_FALLBACK_W_PCT = 22;
const TEMPLATE_SIGNATURE_FALLBACK_H_PCT = 5;
const TEMPLATE_COORD_BASE_W_PT = 612;
const TEMPLATE_COORD_BASE_H_PT = 792;
const PAST_EVENT_ADMIN_EDIT_CODE = "JDL-ADMIN-2026";
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
let quoteAdvanceEditingId = "";
let companyManagersDraft = [];
let editingCompanyId = "";
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
let authSession = { userId: "", fullName: "", username: "", avatarDataUrl: "", signatureDataUrl: "" };
let userSignatureNormalizedDataUrl = "";
let checklistTemplateDraft = [];
let checklistTemplateEditingId = "";
let checklistTemplateSectionsDraft = [];
let checklistTemplatesDraft = [];
let checklistTemplateCurrentId = "";
let checklistTemplateSectionEditingId = "";
let currentEventChecklistId = "";
let eventChecklistDraft = null;
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
let dashboardHoverTipEl = null;
let menuMontajeSelectableCatalogCache = {
  proteins: [],
  preparationsByProtein: new Map(),
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
  btnOpenModules: document.getElementById("btnOpenModules"),
  btnModuleCalendar: document.getElementById("btnModuleCalendar"),
  btnModuleReports: document.getElementById("btnModuleReports"),
  btnModuleSettings: document.getElementById("btnModuleSettings"),
  btnBackFromReports: document.getElementById("btnBackFromReports"),
  btnBackFromSettings: document.getElementById("btnBackFromSettings"),
  topbarWelcome: document.getElementById("topbarWelcome"),
  topbarUserAvatar: document.getElementById("topbarUserAvatar"),
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
  btnReportSales: document.getElementById("btnReportSales"),
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
  salesReportBody: document.getElementById("salesReportBody"),
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
  quoteDocFold: document.getElementById("quoteDocFold"),
  quoteForm: document.getElementById("quoteForm"),
  quoteEventId: document.getElementById("quoteEventId"),
  quoteSubtitle: document.getElementById("quoteSubtitle"),
  btnQuoteClose: document.getElementById("btnQuoteClose"),
  btnQuoteDiscard: document.getElementById("btnQuoteDiscard"),
  quoteVersionSelect: document.getElementById("quoteVersionSelect"),
  quoteTemplateSelect: document.getElementById("quoteTemplateSelect"),
  btnLoadQuoteVersion: document.getElementById("btnLoadQuoteVersion"),
  quoteServiceTemplateSelect: document.getElementById("quoteServiceTemplateSelect"),
  quoteServiceTemplateName: document.getElementById("quoteServiceTemplateName"),
  btnQuoteServiceTemplateApply: document.getElementById("btnQuoteServiceTemplateApply"),
  btnQuoteServiceTemplateSave: document.getElementById("btnQuoteServiceTemplateSave"),
  btnQuoteServiceTemplateUpdate: document.getElementById("btnQuoteServiceTemplateUpdate"),
  btnQuoteServiceTemplateDelete: document.getElementById("btnQuoteServiceTemplateDelete"),
  quoteCompanySearch: document.getElementById("quoteCompanySearch"),
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
  servicesList: document.getElementById("servicesList"),
  serviceDescriptionsList: document.getElementById("serviceDescriptionsList"),
  btnAddServiceToQuote: document.getElementById("btnAddServiceToQuote"),
  quoteItemsPanel: document.getElementById("quoteItemsPanel"),
  btnToggleQuoteItemsExpand: document.getElementById("btnToggleQuoteItemsExpand"),
  quoteItemsBody: document.getElementById("quoteItemsBody"),
  quoteDiscountType: document.getElementById("quoteDiscountType"),
  quoteDiscountValue: document.getElementById("quoteDiscountValue"),
  quoteSubtotal: document.getElementById("quoteSubtotal"),
  quoteDiscountAmount: document.getElementById("quoteDiscountAmount"),
  quoteTotal: document.getElementById("quoteTotal"),
  quoteInternalNotes: document.getElementById("quoteInternalNotes"),
  btnMenuMontaje: document.getElementById("btnMenuMontaje"),
  btnQuoteAdvances: document.getElementById("btnQuoteAdvances"),
  btnMenuMontajeSelectable: document.getElementById("btnMenuMontajeSelectable"),
  btnQuotePrintTemplate: document.getElementById("btnQuotePrintTemplate"),
  quoteAdvanceBackdrop: document.getElementById("quoteAdvanceBackdrop"),
  btnQuoteAdvanceClose: document.getElementById("btnQuoteAdvanceClose"),
  btnQuoteAdvanceDone: document.getElementById("btnQuoteAdvanceDone"),
  quoteAdvanceAmount: document.getElementById("quoteAdvanceAmount"),
  quoteAdvancePaymentType: document.getElementById("quoteAdvancePaymentType"),
  quoteAdvanceDate: document.getElementById("quoteAdvanceDate"),
  quoteAdvanceDescription: document.getElementById("quoteAdvanceDescription"),
  btnQuoteAdvanceAdd: document.getElementById("btnQuoteAdvanceAdd"),
  quoteAdvanceBody: document.getElementById("quoteAdvanceBody"),
  quoteAdvanceTotal: document.getElementById("quoteAdvanceTotal"),

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
      file: "./ContradoCorp.html",
      label: "Contrato Corporativo",
      attachToQuote: true,
      headerImage: "",
    };
  }
  if (id === SERVIHOSP_TEMPLATE_ID) {
    return {
      file: "./Contrato.html",
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
      background: "#0b1a32",
      color: "#f8fafc",
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
      background: "#0b1a32",
      color: "#f8fafc",
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
        <button class="apptIconBtn apptEdit" type="button" data-global-goal-edit="${escapeHtml(globalGoalCompositeKey(row))}" title="Editar">ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã¢â‚¬Â¦ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â½</button>
        <button class="apptIconBtn ${row.active === false ? "" : "apptDelete"}" type="button" data-global-goal-toggle="${escapeHtml(globalGoalCompositeKey(row))}" title="${row.active === false ? "Reactivar" : "Inhabilitar"}">${row.active === false ? "ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Âº" : "ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â"}</button>
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
      <td class="appointmentActions">
        <button class="apptIconBtn apptEdit" type="button" data-salon-edit="${escapeHtml(label)}" title="Editar">ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã¢â‚¬Â¦ÃƒÂ¢Ã¢â€šÂ¬Ã…â€œÃƒÆ’Ã¢â‚¬Â¦Ãƒâ€šÃ‚Â½</button>
        <button class="apptIconBtn ${isSalonDisabled(label) ? "" : "apptDelete"}" type="button" data-salon-toggle="${escapeHtml(label)}" title="${isSalonDisabled(label) ? "Reactivar" : "Inhabilitar"}">${isSalonDisabled(label) ? "ÃƒÆ’Ã†â€™Ãƒâ€šÃ‚Â¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬Ãƒâ€šÃ‚Â ÃƒÆ’Ã¢â‚¬Å¡Ãƒâ€šÃ‚Âº" : "ÃƒÆ’Ã†â€™Ãƒâ€ Ã¢â‚¬â„¢ÃƒÆ’Ã‚Â¢ÃƒÂ¢Ã¢â‚¬Å¡Ã‚Â¬ÃƒÂ¢Ã¢â€šÂ¬Ã‚Â"}</button>
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
  const res = await fetch(endpoint, { cache: "no-store" });
  if (!res.ok) throw new Error("menu_suggestions_read_failed");
  return res.json();
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
      <span class="menuSuggestDrag" title="Arrastra para priorizar">&#9776;</span>
      <input type="checkbox" value="${escapeHtml(id)}" ${isChecked ? "checked" : ""} />
      <span>${escapeHtml(String(item.nombre || "").trim())}</span>
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
          <button type="button" class="btn" data-mmcat-action="edit" data-mmcat-id="${escapeHtml(id)}">Editar</button>
          <button type="button" class="btnDanger" data-mmcat-action="toggle" data-mmcat-id="${escapeHtml(id)}">${isActive ? "Inhabilitar" : "Reactivar"}</button>
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
    return;
  }

  const [salsas, postres, guarniciones, links] = await Promise.all([
    readMenuCatalog("salsa"),
    readMenuCatalog("postre"),
    readMenuCatalog("guarnicion"),
    readMenuSuggestions({ platoId, preparacionId }),
  ]);
  renderMenuSuggestionCheckboxList(el.menuSuggestionsSalsas, salsas, links?.salsaIds || []);
  renderMenuSuggestionCheckboxList(el.menuSuggestionsPostres, postres, links?.postreIds || []);
  renderMenuSuggestionCheckboxList(el.menuSuggestionsGuarniciones, guarniciones, links?.guarnicionIds || []);
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
}

function closeMenuSuggestionsModal() {
  if (!el.menuSuggestionsBackdrop) return;
  el.menuSuggestionsBackdrop.hidden = true;
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
}

function showModuleHub() {
  setActiveModuleScreen("hub");
}

function showCalendarModule() {
  setActiveModuleScreen("");
}

function showReportsHub() {
  setActiveModuleScreen("reports");
}

function showSettingsHub() {
  setQuickAddGroupOpen(true);
  setSettingsPanelOpen(true);
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
  for (const ev of state.events || []) {
    const quote = ev?.quote || null;
    const user = (state.users || []).find((u) => String(u.id) === String(ev.userId));
    const company = quote?.companyId ? (state.companies || []).find((c) => String(c.id) === String(quote.companyId)) : null;
    const manager = company?.managers?.find((m) => String(m.id) === String(quote?.managerId));
    const totals = getQuoteTotals(quote || {});
    const { subcatBuckets, catBuckets } = aggregateQuoteBuckets(quote || {});
    rows.push({
      event: ev,
      status: String(ev.status || ""),
      statusColor: statusColor(ev.status),
      refId: String(quote?.code || reservationKeyFromEvent(ev) || ev.id || ""),
      seller: String(user?.fullName || user?.name || ""),
      eventDate: String(ev.date || ""),
      eventType: String(quote?.eventType || ev.name || ""),
      startTime: String(ev.startTime || ""),
      endTime: String(ev.endTime || ""),
      salon: String(ev.salon || ""),
      company: String(company?.name || quote?.companyName || ""),
      manager: String(manager?.phone || quote?.managerPhone || ""),
      pax: Number(ev.pax || quote?.people || 0),
      subcatBuckets,
      catBuckets,
      discount: Number(totals.discountAmount || 0),
      updatedAt: getEventLastUpdatedLabel(ev),
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

  return buildSalesReportRows().filter((r) => {
    if (from && r.eventDate && r.eventDate < from) return false;
    if (to && r.eventDate && r.eventDate > to) return false;
    if (userId && String(r.event?.userId || "") !== userId) return false;
    if (status && r.status !== status) return false;
    if (salon && r.salon !== salon) return false;
    if (company && String(r.event?.quote?.companyId || "") !== company) return false;
    if (search) {
      const blob = [
        r.refId, r.seller, r.eventType, r.salon, r.company, r.manager, r.status,
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
  const salones = Array.from(new Set((state.events || []).map((e) => String(e.salon || "")).filter(Boolean))).sort();
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
  };

  fillSelect(el.salesReportUser, users.map((u) => ({ value: u.id, label: u.fullName || u.name })), "Todos vendedores");
  fillSelect(el.salesReportStatus, statuses.map((s) => ({ value: s, label: s })), "Todos estados");
  fillSelect(el.salesReportSalon, salones.map((s) => ({ value: s, label: s })), "Todos salones");
  fillSelect(el.salesReportCompany, companies.map((c) => ({ value: c.id, label: c.name })), "Todas instituciones");
}

function renderSalesReportTable() {
  if (!el.salesReportBody) return;
  const rows = getSalesReportFilteredRows();
  el.salesReportBody.innerHTML = "";
  if (!rows.length) {
    const tr = document.createElement("tr");
    tr.innerHTML = `<td colspan="35">Sin resultados para los filtros seleccionados.</td>`;
    el.salesReportBody.appendChild(tr);
    return;
  }
  const pick = (obj, key) => obj?.[key] || { qty: 0, amount: 0 };
  for (const r of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><span class="salesStatusBadge" style="background:${hexToRgba(r.statusColor, 0.25)};border-color:${hexToRgba(r.statusColor, 0.6)}">${escapeHtml(r.status || "-")}</span></td>
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
    <table>
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
  return `<button type="button" class="${cls}" data-event-id="${escapeHtml(eventId)}">${escapeHtml(label)}</button>`;
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

function renderOccupancySummary(rows) {
  if (!el.occupancyReportSummary) return;
  const confirmed = rows.filter((r) => r.status === STATUS.CONFIRMADO).length;
  const pre = rows.filter((r) => r.status === STATUS.PRERESERVA).length;
  const pax = rows.reduce((acc, r) => acc + Math.max(0, Number(r.pax || 0)), 0);
  const totalsByReservation = new Map();
  for (const r of rows) {
    const key = String(r.reservationKey || r.eventId || "");
    if (!key) continue;
    if (!totalsByReservation.has(key)) totalsByReservation.set(key, Math.max(0, Number(r.totalEvent || 0)));
  }
  const total = Array.from(totalsByReservation.values()).reduce((acc, n) => acc + n, 0);
  el.occupancyReportSummary.innerHTML = `
    <div class="occupancyCard">
      <small>Eventos semana</small>
      <strong>${rows.length}</strong>
    </div>
    <div class="occupancyCard occupancyConfirmed">
      <small>Confirmados</small>
      <strong>${confirmed}</strong>
    </div>
    <div class="occupancyCard occupancyPre">
      <small>Pre reserva</small>
      <strong>${pre}</strong>
    </div>
    <div class="occupancyCard">
      <small>PAX total</small>
      <strong>${pax}</strong>
    </div>
    <div class="occupancyCard">
      <small>Total cotizado</small>
      <strong>${moneyGT(total)}</strong>
    </div>
  `;
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
  const countByDay = new Map(dates.map((d) => [d, 0]));
  for (const r of rows) {
    const prev = countByDay.get(r.eventDate) || 0;
    countByDay.set(r.eventDate, prev + 1);
  }
  el.occupancyDaysStrip.innerHTML = "";
  for (const d of dates) {
    const count = Number(countByDay.get(d) || 0);
    const card = document.createElement("button");
    card.type = "button";
    card.className = `occupancyDayCard${d === occupancySelectedDayIso ? " active" : ""}`;
    card.innerHTML = `
      <small>${escapeHtml(formatDayCardLabel(d))}</small>
      <strong>${count}</strong>
      <span>evento${count === 1 ? "" : "s"}</span>
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
  if (!dayRows.length) {
    el.occupancyDayDetail.innerHTML = `
      <div class="occupancyDayDetailTitle">${escapeHtml(title)}</div>
      <div class="occupancyDayDetailEmpty">Sin eventos Confirmados/Pre reserva para este dia.</div>
    `;
    return;
  }
  const cards = dayRows.map((r) => `
    <article class="occupancyEventCard">
      <div class="occupancyEventHead">
        <span class="salesStatusBadge" style="background:${escapeHtml(hexToRgba(r.statusColor, 0.25))};border-color:${escapeHtml(hexToRgba(r.statusColor, 0.6))}">${escapeHtml(r.status)}</span>
        <strong>${escapeHtml(r.refId)}</strong>
      </div>
      <div class="occupancyEventBody">
        <div class="occupancyInfoGrid">
          <div class="occupancyInfoItem occupancyInfoWide">
            <small>Evento</small>
            <strong title="${escapeHtml(r.eventName || "-")}">${escapeHtml(r.eventName || "-")}</strong>
          </div>
          <div class="occupancyInfoItem">
            <small>Horario</small>
            <strong>${escapeHtml(r.startTime)} - ${escapeHtml(r.endTime)}</strong>
          </div>
          <div class="occupancyInfoItem">
            <small>Salon</small>
            <strong>${escapeHtml(r.salon || "-")}</strong>
          </div>
          <div class="occupancyInfoItem">
            <small>Institucion</small>
            <strong title="${escapeHtml(r.company || "-")}">${escapeHtml(r.company || "-")}</strong>
          </div>
          <div class="occupancyInfoItem">
            <small>Encargado</small>
            <strong title="${escapeHtml(r.manager || "-")}">${escapeHtml(r.manager || "-")}</strong>
          </div>
          <div class="occupancyInfoItem">
            <small>Vendedor</small>
            <strong title="${escapeHtml(r.seller || "-")}">${escapeHtml(r.seller || "-")}</strong>
          </div>
          <div class="occupancyInfoItem">
            <small>PAX</small>
            <strong>${escapeHtml(String(r.pax || 0))}</strong>
          </div>
        </div>
        <div class="occupancyMetricsGrid">
          <div class="occupancyMetricItem">
            <small>Total evento</small>
            <strong>${escapeHtml(moneyGT(r.totalEvent || 0))}</strong>
          </div>
          <div class="occupancyMetricItem">
            <small>Ingreso dia</small>
            <strong>${escapeHtml(moneyGT(r.incomePerDay || 0))}</strong>
          </div>
          <div class="occupancyMetricItem">
            <small>Ingreso salon-dia</small>
            <strong>${escapeHtml(moneyGT(r.incomePerSalonDay || 0))}</strong>
          </div>
        </div>
        <div class="occupancyActionGrid">
          <div class="occupancyActionItem">
            <small>Ult. cotizacion</small>
            <div>${buildOccupancyQuoteActionHtml(r)}</div>
          </div>
          <div class="occupancyActionItem">
            <small>Ult. informe</small>
            <div>${buildOccupancyMenuMontajeActionHtml(r)}</div>
          </div>
          <div class="occupancyActionItem">
            <small>Check List</small>
            <div>${buildOccupancyChecklistActionHtml(r)}</div>
          </div>
        </div>
      </div>
    </article>
  `).join("");
  el.occupancyDayDetail.innerHTML = `
    <div class="occupancyDayDetailTitle">${escapeHtml(title)}</div>
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
    tr.innerHTML = `<td colspan="17">Sin eventos Confirmados/Pre reserva para esta semana.</td>`;
    el.occupancyReportBody.appendChild(tr);
    return;
  }
  for (const r of rows) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td><span class="salesStatusBadge" style="background:${hexToRgba(r.statusColor, 0.25)};border-color:${hexToRgba(r.statusColor, 0.6)}">${escapeHtml(r.status || "-")}</span></td>
      <td>${escapeHtml(String(r.pax || 0))}</td>
      <td>${escapeHtml(r.eventDate || "-")}</td>
      <td>${escapeHtml(r.startTime || "-")}</td>
      <td>${escapeHtml(r.endTime || "-")}</td>
      <td>${escapeHtml(r.eventName || "-")}</td>
      <td>${escapeHtml(r.salon || "-")}</td>
      <td>${escapeHtml(r.company || "-")}</td>
      <td>${escapeHtml(r.manager || "-")}</td>
      <td>${escapeHtml(r.seller || "-")}</td>
      <td>${buildOccupancyQuoteActionHtml(r)}</td>
      <td>${buildOccupancyMenuMontajeActionHtml(r)}</td>
      <td>${buildOccupancyChecklistActionHtml(r)}</td>
      <td>${escapeHtml(moneyGT(r.totalEvent || 0))}</td>
      <td>${escapeHtml(moneyGT(r.incomePerDay || 0))}</td>
      <td>${escapeHtml(moneyGT(r.incomePerSalonDay || 0))}</td>
      <td>${escapeHtml(r.updatedAt || "-")}</td>
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
  resetDashboardReportFilters();
  renderDashboardReport();
  el.dashboardReportBackdrop.hidden = false;
}

function closeDashboardReportModal() {
  if (!el.dashboardReportBackdrop) return;
  el.dashboardReportBackdrop.hidden = true;
  restoreModuleScreenAfterModal();
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

function renderInstitutionReportCompanyOptions() {
  if (!el.institutionReportCompany) return;
  const previous = String(el.institutionReportCompany.value || "").trim();
  const search = String(el.institutionReportCompanySearch?.value || "").trim();
  const companies = (state.companies || [])
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
  el.institutionReportCompany.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = search ? "Selecciona una coincidencia" : "Selecciona institucion";
  el.institutionReportCompany.appendChild(placeholder);
  for (const company of companies) {
    const opt = document.createElement("option");
    opt.value = String(company.id || "").trim();
    opt.textContent = String(company.name || "Institucion").trim();
    el.institutionReportCompany.appendChild(opt);
  }
  if (previous && companies.some((c) => String(c.id) === previous)) {
    el.institutionReportCompany.value = previous;
  } else if (!previous && companies.length) {
    el.institutionReportCompany.value = String(companies[0].id || "");
  } else if (companies.length) {
    el.institutionReportCompany.value = String(companies[0].id || "");
  } else {
    el.institutionReportCompany.value = "";
  }
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
      <div class="institutionDonutLegendItem" tabindex="0">
        <span class="institutionDonutDot" style="background:${seg.color}"></span>
        <b>${escapeHtml(seg.label)}</b>
        <span>${escapeHtml(String(seg.value))} (${escapeHtml(pct.toFixed(1))}%)</span>
        <span class="institutionChartTip">${escapeHtml(tip)}</span>
      </div>
    `;
  }).join("");
  return `
    <div class="institutionDonutWrap">
      <div class="institutionDonut" style="background:conic-gradient(${stops});">
        <div class="institutionDonutHole">
          <small>Total</small>
          <strong>${escapeHtml(String(total))}</strong>
        </div>
      </div>
      <div class="institutionDonutLegend">${legend}</div>
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
    if (!data.company) {
      el.institutionReportHeadline.innerHTML = `<div class="dashboardEmpty">Selecciona una institucion para ver su dashboard.</div>`;
    } else {
      const lastVisitText = data.latestVisitIso
        ? `${data.latestVisitIso}${Number.isFinite(data.daysSinceLastVisit) ? ` | hace ${data.daysSinceLastVisit} dias` : ""}`
        : "Sin visitas registradas";
      const companyType = String(data.company.eventType || "").trim() || "Sin tipo definido";
      el.institutionReportHeadline.innerHTML = `
        <div class="institutionHeadlineTop">
          <div>
            <strong>${escapeHtml(String(data.company.name || "Institucion"))}</strong>
            <small>${escapeHtml(companyType)} | ${escapeHtml(String(data.company.owner || data.company.email || "Sin contacto principal"))}</small>
          </div>
          <div class="institutionHeadlineMoney">${escapeHtml(moneyGT(data.totalRevenue || 0))}</div>
        </div>
        <div class="institutionHeadlineMeta">
          <span class="pill">Ultima visita: ${escapeHtml(lastVisitText)}</span>
          <span class="pill">Primera visita: ${escapeHtml(data.firstVisitIso || "-")}</span>
          <span class="pill">Encargado top: ${escapeHtml(data.topManagerLabel || "-")}</span>
          <span class="pill">Vendedor top: ${escapeHtml(data.topSellerLabel || "-")}</span>
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
      el.institutionOverviewGrid.innerHTML = `<div class="dashboardEmpty">Sin datos para mostrar.</div>`;
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
      el.institutionReportChartsBody.innerHTML = `<div class="dashboardEmpty">Sin datos para graficas.</div>`;
    } else {
      el.institutionReportChartsBody.innerHTML = `
        <article class="institutionChartCard">
          <header>
            <strong>Ingresos por mes</strong>
            <small>Top 8 meses del rango (hover para detalle)</small>
          </header>
          ${buildInstitutionMonthlyBarChartHtml(data.monthRows)}
        </article>
        <article class="institutionChartCard">
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

  if (el.institutionReportDishBody) {
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

  if (el.institutionReportManagerBody) {
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

  if (el.institutionReportTimelineBody) {
    if (!data.company) {
      el.institutionReportTimelineBody.innerHTML = `<div class="dashboardEmpty">Sin datos de historial.</div>`;
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
        <button class="btn" type="button" data-checklist-template-edit="${escapeHtml(String(item.id || ""))}">Editar</button>
        <button class="btnDanger" type="button" data-checklist-template-remove="${escapeHtml(String(item.id || ""))}">X</button>
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
        <button class="btn" type="button" data-checklist-template-edit="${escapeHtml(String(item.id || ""))}">Editar</button>
        <button class="btnDanger" type="button" data-checklist-template-remove="${escapeHtml(String(item.id || ""))}">X</button>
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
        showSearch: false,
        openPosition: "auto",
        contentPosition: "fixed",
        contentLocation: document.body,
        closeOnSelect: true,
        placeholderText: "Selecciona opcion",
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
          if (el.loginScreen) el.loginScreen.hidden = false;
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
