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
  "/api/state",
  ...(CURRENT_ORIGIN_STATE_URL ? [CURRENT_ORIGIN_STATE_URL] : []),
  "http://localhost:3000/api/state",
  "http://127.0.0.1:3000/api/state",
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
let companyManagersDraft = [];
let editingCompanyId = "";
let editingServiceId = "";
let catalogoCategoriasServicio = [];
let catalogoSubcategoriasServicio = [];
let historyTargetEventId = null;
let appointmentTargetEventId = null;
let userModalEditingId = "";
let userMonthlyGoalsDraft = [];
let editingUserGoalMonth = "";
let occupancySelectedDayIso = "";
let authSession = { userId: "", fullName: "", username: "", avatarDataUrl: "", signatureDataUrl: "" };
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
let quickTemplates = Array.isArray(state.quickTemplates) ? state.quickTemplates : [];
let activeTemplateId = "";
let activeTemplateEditor = null;
let templateAssetsDraft = {
  pagePdf: "",
  headerImage: "",
  footerImage: "",
};
const templateUiState = {
  activePositionIndex: -1,
  dragging: null,
  resizing: null,
  pdfRenderKey: "",
  signatureDefaultW: 22,
  signatureDefaultH: 5,
};
const TEMPLATE_SIGNATURE_MIN_W_PCT = 10;
const TEMPLATE_SIGNATURE_MIN_H_PCT = 3;
const TEMPLATE_SIGNATURE_MAX_W_PCT = 35;
const TEMPLATE_SIGNATURE_MAX_H_PCT = 12;
const TEMPLATE_SIGNATURE_FALLBACK_W_PCT = 22;
const TEMPLATE_SIGNATURE_FALLBACK_H_PCT = 5;
const TEMPLATE_COORD_BASE_W_PT = 612;
const TEMPLATE_COORD_BASE_H_PT = 792;
const signatureImageAnalysisCache = new Map();
const pastEventEditAuthorizedKeys = new Set();
const notifiedReminderKeys = new Set();
let menuMontajeSelectedKey = "";
let menuMontajeSelectedVersion = 0;

const el = {
  loginScreen: document.getElementById("loginScreen"),
  loginForm: document.getElementById("loginForm"),
  loginUserSelect: document.getElementById("loginUserSelect"),
  loginPassword: document.getElementById("loginPassword"),
  loginAvatar: document.getElementById("loginAvatar"),
  loginError: document.getElementById("loginError"),
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
  btnQuickAddTemplate: document.getElementById("btnQuickAddTemplate"),
  btnReportSales: document.getElementById("btnReportSales"),
  btnReportOccupancy: document.getElementById("btnReportOccupancy"),
  btnReportDashboard: document.getElementById("btnReportDashboard"),
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
  btnOccupancyReportTodayWeek: document.getElementById("btnOccupancyReportTodayWeek"),
  btnOccupancyReportExportExcel: document.getElementById("btnOccupancyReportExportExcel"),
  occupancyReportSummary: document.getElementById("occupancyReportSummary"),
  occupancyDaysStrip: document.getElementById("occupancyDaysStrip"),
  occupancyDayDetail: document.getElementById("occupancyDayDetail"),
  occupancyReportBody: document.getElementById("occupancyReportBody"),
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
  btnMarkQuoted: document.getElementById("btnMarkQuoted"),
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
  quoteServiceDate: document.getElementById("quoteServiceDate"),
  quoteServiceSearch: document.getElementById("quoteServiceSearch"),
  servicesList: document.getElementById("servicesList"),
  serviceDescriptionsList: document.getElementById("serviceDescriptionsList"),
  btnAddServiceToQuote: document.getElementById("btnAddServiceToQuote"),
  quoteItemsBody: document.getElementById("quoteItemsBody"),
  quoteDiscountType: document.getElementById("quoteDiscountType"),
  quoteDiscountValue: document.getElementById("quoteDiscountValue"),
  quoteSubtotal: document.getElementById("quoteSubtotal"),
  quoteDiscountAmount: document.getElementById("quoteDiscountAmount"),
  quoteTotal: document.getElementById("quoteTotal"),
  quoteInternalNotes: document.getElementById("quoteInternalNotes"),
  btnMenuMontaje: document.getElementById("btnMenuMontaje"),

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
  btnAddCompany: document.getElementById("btnAddCompany"),
  btnOpenServiceCreate: document.getElementById("btnOpenServiceCreate"),

  companyBackdrop: document.getElementById("companyBackdrop"),
  companyTitle: document.getElementById("companyTitle"),
  companyForm: document.getElementById("companyForm"),
  companyName: document.getElementById("companyName"),
  companyOwner: document.getElementById("companyOwner"),
  companyEmail: document.getElementById("companyEmail"),
  companyNIT: document.getElementById("companyNIT"),
  companyBusinessName: document.getElementById("companyBusinessName"),
  companyEventType: document.getElementById("companyEventType"),
  companyAddress: document.getElementById("companyAddress"),
  companyPhone: document.getElementById("companyPhone"),
  companyNotes: document.getElementById("companyNotes"),
  managerName: document.getElementById("managerName"),
  managerPhone: document.getElementById("managerPhone"),
  managerEmail: document.getElementById("managerEmail"),
  managerAddress: document.getElementById("managerAddress"),
  btnAddManager: document.getElementById("btnAddManager"),
  managersBody: document.getElementById("managersBody"),
  btnCompanyClose: document.getElementById("btnCompanyClose"),
  btnCompanyDiscard: document.getElementById("btnCompanyDiscard"),

  serviceBackdrop: document.getElementById("serviceBackdrop"),
  serviceTitle: document.getElementById("serviceTitle"),
  serviceForm: document.getElementById("serviceForm"),
  serviceName: document.getElementById("serviceName"),
  serviceCategory: document.getElementById("serviceCategory"),
  serviceSubcategory: document.getElementById("serviceSubcategory"),
  servicePrice: document.getElementById("servicePrice"),
  serviceQuantityMode: document.getElementById("serviceQuantityMode"),
  serviceDescription: document.getElementById("serviceDescription"),
  btnServiceClose: document.getElementById("btnServiceClose"),
  btnServiceDiscard: document.getElementById("btnServiceDiscard"),

  templateBackdrop: document.getElementById("templateBackdrop"),
  templateForm: document.getElementById("templateForm"),
  templateName: document.getElementById("templateName"),
  templateSelect: document.getElementById("templateSelect"),
  btnTemplateLoad: document.getElementById("btnTemplateLoad"),
  btnTemplateDelete: document.getElementById("btnTemplateDelete"),
  templatePageImage: document.getElementById("templatePageImage"),
  templateHeaderImage: document.getElementById("templateHeaderImage"),
  templateFooterImage: document.getElementById("templateFooterImage"),
  templatePagePreview: document.getElementById("templatePagePreview"),
  templateFieldPanel: document.getElementById("templateFieldPanel"),
  templateHeader: document.getElementById("templateHeader"),
  templateBody: document.getElementById("templateBody"),
  templateFooter: document.getElementById("templateFooter"),
  templatePositionBody: document.getElementById("templatePositionBody"),
  btnTemplateAddPosition: document.getElementById("btnTemplateAddPosition"),
  btnTemplateFitSignature: document.getElementById("btnTemplateFitSignature"),
  templateFormulaBody: document.getElementById("templateFormulaBody"),
  btnTemplateAddFormula: document.getElementById("btnTemplateAddFormula"),
  templateTokenGrid: document.getElementById("templateTokenGrid"),
  templateRoomRates: document.getElementById("templateRoomRates"),
  templatePreview: document.getElementById("templatePreview"),
  btnTemplateClose: document.getElementById("btnTemplateClose"),
  btnTemplateDiscard: document.getElementById("btnTemplateDiscard"),
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

function loadQuickTemplates() {
  try {
    const raw = localStorage.getItem(QUICK_TEMPLATES_STORAGE_KEY);
    if (!raw) return [];
    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];
    return parsed
      .map(normalizeTemplateRecord)
      .filter(Boolean);
  } catch (_) {
    return [];
  }
}

function syncQuickTemplatesIntoState() {
  if (!state || typeof state !== "object") return;
  state.quickTemplates = quickTemplates
    .map(normalizeTemplateRecord)
    .filter(Boolean);
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
  const action = await promptCrudAction("Instituciones");
  if (!action) return;
  if (action === "add") return openCompanyModal();

  const companies = (state.companies || []).slice().sort((a, b) =>
    String(a.name || "").localeCompare(String(b.name || ""), "es", { sensitivity: "base" })
  );
  if (!companies.length) return toast("No hay instituciones registradas.");
  const selectedId = await promptSelectRequired({
    title: action === "edit" ? "Editar institucion" : "Inhabilitar institucion",
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
  toast("Institucion inhabilitada.");
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

function getGlobalMonthlyGoals() {
  const rows = Array.isArray(state.globalMonthlyGoals) ? state.globalMonthlyGoals : [];
  return rows
    .map((g) => ({
      month: String(g?.month || "").trim(),
      amount: Math.max(0, Number(g?.amount || 0)),
    }))
    .filter((g) => /^\d{4}-\d{2}$/.test(g.month))
    .sort((a, b) => a.month.localeCompare(b.month));
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
    const amountRaw = await promptTextRequired({
      title: "Monto meta global",
      label: `Mes ${month}`,
      placeholder: "Ej: 250000",
    });
    const amount = Math.max(0, Number(amountRaw || 0));
    if (!Number.isFinite(amount) || amount <= 0) return toast("Monto de meta invalido.");
    const next = current.filter((g) => g.month !== month);
    next.push({ month, amount });
    next.sort((a, b) => a.month.localeCompare(b.month));
    state.globalMonthlyGoals = next;
    persist();
    return toast("Meta global mensual guardada.");
  }

  if (!current.length) return toast("No hay metas globales registradas.");
  const selectedMonth = await promptSelectRequired({
    title: action === "edit" ? "Editar meta global" : "Eliminar meta global",
    options: current.map((g) => ({
      value: g.month,
      label: `${g.month} - Q ${g.amount.toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`,
    })),
  });
  if (!selectedMonth) return;

  if (action === "edit") {
    const target = current.find((g) => g.month === selectedMonth);
    if (!target) return;
    const amountRaw = await promptTextRequired({
      title: "Nuevo monto meta global",
      label: `Mes ${selectedMonth}`,
      placeholder: String(target.amount),
    });
    const amount = Math.max(0, Number(amountRaw || 0));
    if (!Number.isFinite(amount) || amount <= 0) return toast("Monto de meta invalido.");
    state.globalMonthlyGoals = current.map((g) => (g.month === selectedMonth ? { month: selectedMonth, amount } : g));
    persist();
    return toast("Meta global actualizada.");
  }

  const ok = await modernConfirm({
    title: "Eliminar meta global",
    message: `Esta seguro de eliminar la meta global del mes ${selectedMonth}?`,
    confirmText: "Si, eliminar",
    cancelText: "No",
  });
  if (!ok) return;
  state.globalMonthlyGoals = current.filter((g) => g.month !== selectedMonth);
  persist();
  toast("Meta global eliminada.");
}

function renderTemplateSelect(selectedId = "") {
  if (!el.templateSelect) return;
  el.templateSelect.innerHTML = "";
  const placeholder = document.createElement("option");
  placeholder.value = "";
  placeholder.textContent = quickTemplates.length
    ? "Selecciona plantilla guardada"
    : "Sin plantillas guardadas";
  el.templateSelect.appendChild(placeholder);

  const ordered = quickTemplates
    .slice()
    .sort((a, b) => a.name.localeCompare(b.name, "es", { sensitivity: "base" }));
  for (const t of ordered) {
    const opt = document.createElement("option");
    opt.value = t.id;
    opt.textContent = t.name;
    el.templateSelect.appendChild(opt);
  }
  el.templateSelect.value = selectedId || "";
}

function addTemplateFormulaRow({ key = "", expression = "" } = {}) {
  if (!el.templateFormulaBody) return;
  const tr = document.createElement("tr");
  tr.className = "templateFormulaRow";
  tr.innerHTML = `
    <td><input class="quoteInput formulaKey" type="text" placeholder="ej: impuesto" value="${escapeHtml(key)}" /></td>
    <td><input class="quoteInput formulaExpr" type="text" placeholder="ej: monto.total * 0.12" value="${escapeHtml(expression)}" /></td>
    <td class="formulaRowActions"><button class="btnDanger formulaRemoveBtn" type="button">X</button></td>
  `;
  el.templateFormulaBody.appendChild(tr);
}

function addTemplatePositionRow({
  label = "",
  token = "",
  x = 50,
  y = 50,
  w = null,
  h = null,
  fontSize = 12,
  fontFamily = "Arial",
  bold = false,
  italic = false,
} = {}) {
  const isSignature = isTemplateSignatureToken(token);
  const minW = isSignature ? TEMPLATE_SIGNATURE_MIN_W_PCT : 4;
  const minH = isSignature ? TEMPLATE_SIGNATURE_MIN_H_PCT : 2;
  const baseW = Number.isFinite(Number(w))
    ? Number(w)
    : (isSignature ? Number(templateUiState.signatureDefaultW || TEMPLATE_SIGNATURE_FALLBACK_W_PCT) : TEMPLATE_SIGNATURE_FALLBACK_W_PCT);
  const baseH = Number.isFinite(Number(h))
    ? Number(h)
    : (isSignature ? Number(templateUiState.signatureDefaultH || TEMPLATE_SIGNATURE_FALLBACK_H_PCT) : TEMPLATE_SIGNATURE_FALLBACK_H_PCT);
  const widthValue = clamp(baseW, minW, isSignature ? TEMPLATE_SIGNATURE_MAX_W_PCT : 95);
  const heightValue = clamp(baseH, minH, isSignature ? TEMPLATE_SIGNATURE_MAX_H_PCT : 60);
  if (!el.templatePositionBody) return;
  const tr = document.createElement("tr");
  tr.className = "templatePositionRow";
  tr.dataset.signature = isSignature ? "1" : "0";
  tr.innerHTML = `
    <td><input class="quoteInput posLabel" type="text" placeholder="Ej: Firma cliente" value="${escapeHtml(label)}" /></td>
    <td><input class="quoteInput posToken" type="text" placeholder="{{cliente.firma}}" value="${escapeHtml(token)}" /></td>
    <td><input class="quoteInput posX" type="number" min="0" max="100" step="0.1" value="${escapeHtml(String(clamp(Number(x), 0, 100)))}" /></td>
    <td><input class="quoteInput posY" type="number" min="0" max="100" step="0.1" value="${escapeHtml(String(clamp(Number(y), 0, 100)))}" /></td>
    <td><input class="quoteInput posW" type="number" min="${isSignature ? TEMPLATE_SIGNATURE_MIN_W_PCT : 4}" max="${isSignature ? TEMPLATE_SIGNATURE_MAX_W_PCT : 95}" step="0.1" value="${escapeHtml(String(widthValue))}" ${isSignature ? "" : "disabled"} /></td>
    <td><input class="quoteInput posH" type="number" min="${isSignature ? TEMPLATE_SIGNATURE_MIN_H_PCT : 2}" max="${isSignature ? TEMPLATE_SIGNATURE_MAX_H_PCT : 60}" step="0.1" value="${escapeHtml(String(heightValue))}" ${isSignature ? "" : "disabled"} /></td>
    <td><input class="quoteInput posFontSize" type="number" min="8" max="72" step="1" value="${escapeHtml(String(clamp(Number(fontSize), 8, 72)))}" /></td>
    <td>
      <select class="quoteInput posFontFamily">
        <option value="Arial"${fontFamily === "Arial" ? " selected" : ""}>Arial</option>
        <option value="Calibri"${fontFamily === "Calibri" ? " selected" : ""}>Calibri</option>
        <option value="Verdana"${fontFamily === "Verdana" ? " selected" : ""}>Verdana</option>
        <option value="Tahoma"${fontFamily === "Tahoma" ? " selected" : ""}>Tahoma</option>
        <option value="Georgia"${fontFamily === "Georgia" ? " selected" : ""}>Georgia</option>
        <option value="Avant Garde"${fontFamily === "Avant Garde" ? " selected" : ""}>Avant Garde</option>
        <option value="Times New Roman"${fontFamily === "Times New Roman" ? " selected" : ""}>Times New Roman</option>
        <option value="Courier New"${fontFamily === "Courier New" ? " selected" : ""}>Courier New</option>
      </select>
    </td>
    <td><input class="posBold" type="checkbox"${bold ? " checked" : ""} /></td>
    <td><input class="posItalic" type="checkbox"${italic ? " checked" : ""} /></td>
    <td class="formulaRowActions"><button class="btnDanger posRemoveBtn" type="button">X</button></td>
  `;
  el.templatePositionBody.appendChild(tr);
  applyTemplatePositionRowSignatureState(tr);
}

function addPresetFieldToTemplatePosition(label, token) {
  if (!el.templatePositionBody) return;
  const markerLayer = el.templatePagePreview?.querySelector(".templateMarkerLayer");
  const canvas = el.templatePagePreview;
  const totalHeight = Math.max(1, markerLayer?.clientHeight || 1);
  const defaultX = 50;
  const defaultY = canvas
    ? clamp(((canvas.scrollTop + canvas.clientHeight * 0.5) / totalHeight) * 100, 4, 96)
    : 50;
  addTemplatePositionRow({ label, token, x: defaultX, y: defaultY });
  const rows = getTemplatePositionRows();
  const idx = rows.length - 1;
  setActiveTemplatePosition(idx);
  refreshTemplatePreview();
  rows[idx]?.scrollIntoView({ block: "nearest" });
}

function collectTemplatePositionedFields() {
  if (!el.templatePositionBody) return [];
  const rows = Array.from(el.templatePositionBody.querySelectorAll(".templatePositionRow"));
  const result = [];
  for (const row of rows) {
    const label = String(row.querySelector(".posLabel")?.value || "").trim();
    const token = String(row.querySelector(".posToken")?.value || "").trim();
    const x = clamp(Number(row.querySelector(".posX")?.value || 0), 0, 100);
    const y = clamp(Number(row.querySelector(".posY")?.value || 0), 0, 100);
    const isSignature = row.dataset.signature === "1";
    const minW = isSignature ? TEMPLATE_SIGNATURE_MIN_W_PCT : 4;
    const minH = isSignature ? TEMPLATE_SIGNATURE_MIN_H_PCT : 2;
    const w = clamp(Number(row.querySelector(".posW")?.value || TEMPLATE_SIGNATURE_FALLBACK_W_PCT), minW, isSignature ? TEMPLATE_SIGNATURE_MAX_W_PCT : 95);
    const h = clamp(Number(row.querySelector(".posH")?.value || TEMPLATE_SIGNATURE_FALLBACK_H_PCT), minH, isSignature ? TEMPLATE_SIGNATURE_MAX_H_PCT : 60);
    const fontSize = clamp(Number(row.querySelector(".posFontSize")?.value || 12), 8, 72);
    const fontFamily = String(row.querySelector(".posFontFamily")?.value || "Arial").trim() || "Arial";
    const bold = row.querySelector(".posBold")?.checked === true;
    const italic = row.querySelector(".posItalic")?.checked === true;
    if (!label && !token) continue;
    result.push({
      label,
      token,
      x,
      y,
      w,
      h,
      xPt: Number(((x / 100) * TEMPLATE_COORD_BASE_W_PT).toFixed(2)),
      yPt: Number(((y / 100) * TEMPLATE_COORD_BASE_H_PT).toFixed(2)),
      wPt: Number(((w / 100) * TEMPLATE_COORD_BASE_W_PT).toFixed(2)),
      hPt: Number(((h / 100) * TEMPLATE_COORD_BASE_H_PT).toFixed(2)),
      isSignature,
      fontSize,
      fontFamily,
      bold,
      italic,
    });
  }
  return result;
}

async function readImageFileAsDataUrl(file) {
  return new Promise((resolve) => {
    if (!file) return resolve("");
    const reader = new FileReader();
    reader.onload = () => resolve(String(reader.result || ""));
    reader.onerror = () => resolve("");
    reader.readAsDataURL(file);
  });
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

function getTemplateSignatureDefaultsFromEditor(positionedFields = null) {
  const fields = Array.isArray(positionedFields) ? positionedFields : collectTemplatePositionedFields();
  const firstSignature = fields.find((p) => p?.isSignature === true || isTemplateSignatureToken(p?.token));
  const defaults = normalizeTemplateSignatureDefaults({
    w: firstSignature?.w || templateUiState.signatureDefaultW || TEMPLATE_SIGNATURE_FALLBACK_W_PCT,
    h: firstSignature?.h || templateUiState.signatureDefaultH || TEMPLATE_SIGNATURE_FALLBACK_H_PCT,
  }, fields);
  templateUiState.signatureDefaultW = defaults.w;
  templateUiState.signatureDefaultH = defaults.h;
  return defaults;
}

function getBestTemplateSignatureRowIndex() {
  const rows = getTemplatePositionRows();
  const activeIdx = templateUiState.activePositionIndex;
  if (Number.isInteger(activeIdx) && activeIdx >= 0 && rows[activeIdx]?.dataset.signature === "1") return activeIdx;
  return rows.findIndex((row) => row.dataset.signature === "1");
}

function fitSignatureRowToAspect(index, ratio = 4) {
  const rows = getTemplatePositionRows();
  const row = rows[index];
  if (!row || row.dataset.signature !== "1") return false;
  const x = clamp(Number(row.querySelector(".posX")?.value || 0), 0, 100);
  const y = clamp(Number(row.querySelector(".posY")?.value || 0), 0, 100);
  const minW = TEMPLATE_SIGNATURE_MIN_W_PCT;
  const minH = TEMPLATE_SIGNATURE_MIN_H_PCT;
  const maxW = Math.min(TEMPLATE_SIGNATURE_MAX_W_PCT, Math.max(minW, 100 - x));
  const maxH = Math.min(TEMPLATE_SIGNATURE_MAX_H_PCT, Math.max(minH, 100 - y));
  const safeRatio = clamp(Number(ratio) || 4, 1.8, 10);
  let nextW = clamp(Number(row.querySelector(".posW")?.value || templateUiState.signatureDefaultW || TEMPLATE_SIGNATURE_FALLBACK_W_PCT), minW, maxW);
  let nextH = clamp(nextW / safeRatio, minH, maxH);
  if (Math.abs((nextW / Math.max(0.1, nextH)) - safeRatio) > 0.08) {
    nextW = clamp(nextH * safeRatio, minW, maxW);
  }
  updateTemplatePositionSizeRow(index, nextW, nextH);
  updateTemplateMarkerSize(index, nextW, nextH);
  syncTemplateSignatureBounds(index);
  templateUiState.signatureDefaultW = clamp(nextW, minW, TEMPLATE_SIGNATURE_MAX_W_PCT);
  templateUiState.signatureDefaultH = clamp(nextH, minH, TEMPLATE_SIGNATURE_MAX_H_PCT);
  return true;
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

function setTemplateEditorContent(template) {
  const t = normalizeTemplateRecord(template) || normalizeTemplateRecord({ name: "" });
  if (!t) return;
  activeTemplateId = t.id || "";
  templateAssetsDraft = { ...t.assets };
  templateUiState.activePositionIndex = -1;
  templateUiState.pdfRenderKey = "";
  templateUiState.signatureDefaultW = clamp(Number(t.signatureDefaults?.w || TEMPLATE_SIGNATURE_FALLBACK_W_PCT), TEMPLATE_SIGNATURE_MIN_W_PCT, TEMPLATE_SIGNATURE_MAX_W_PCT);
  templateUiState.signatureDefaultH = clamp(Number(t.signatureDefaults?.h || TEMPLATE_SIGNATURE_FALLBACK_H_PCT), TEMPLATE_SIGNATURE_MIN_H_PCT, TEMPLATE_SIGNATURE_MAX_H_PCT);
  if (el.templateName) el.templateName.value = t.name || "";
  if (el.templateHeader) el.templateHeader.value = t.header || "";
  if (el.templateBody) el.templateBody.value = t.body || "";
  if (el.templateFooter) el.templateFooter.value = t.footer || "";
  if (el.templateRoomRates) el.templateRoomRates.value = JSON.stringify(t.roomRates || [], null, 2);
  if (el.templatePositionBody) {
    el.templatePositionBody.innerHTML = "";
    for (const p of t.positionedFields || []) addTemplatePositionRow(p);
    if (!(t.positionedFields || []).length) {
      addTemplatePositionRow({ label: "Firma cliente", token: "{{cliente.firma}}", x: 25, y: 84 });
      addTemplatePositionRow({ label: "Firma vendedor", token: "{{vendedor.firma}}", x: 75, y: 84 });
    }
  }
  if (el.templateFormulaBody) {
    el.templateFormulaBody.innerHTML = "";
    for (const f of t.formulas || []) addTemplateFormulaRow(f);
    if (!(t.formulas || []).length) addTemplateFormulaRow();
  }
  renderTemplateSelect(activeTemplateId);
  activeTemplateEditor = el.templateBody || null;
  refreshTemplatePreview();
}

function clearTemplateEditor() {
  activeTemplateId = "";
  templateAssetsDraft = { pagePdf: "", headerImage: "", footerImage: "" };
  templateUiState.activePositionIndex = -1;
  templateUiState.pdfRenderKey = "";
  templateUiState.dragging = null;
  templateUiState.signatureDefaultW = TEMPLATE_SIGNATURE_FALLBACK_W_PCT;
  templateUiState.signatureDefaultH = TEMPLATE_SIGNATURE_FALLBACK_H_PCT;
  if (el.templateName) el.templateName.value = "";
  if (el.templateHeader) el.templateHeader.value = "";
  if (el.templateBody) el.templateBody.value = "";
  if (el.templateFooter) el.templateFooter.value = "";
  if (el.templateRoomRates) el.templateRoomRates.value = "";
  if (el.templateFormulaBody) {
    el.templateFormulaBody.innerHTML = "";
    addTemplateFormulaRow();
  }
  if (el.templatePositionBody) {
    el.templatePositionBody.innerHTML = "";
    addTemplatePositionRow({ label: "Firma cliente", token: "{{cliente.firma}}", x: 25, y: 84 });
    addTemplatePositionRow({ label: "Firma vendedor", token: "{{vendedor.firma}}", x: 75, y: 84 });
    addTemplatePositionRow({ label: "Nombre cliente", token: "{{cliente.nombre}}", x: 25, y: 90 });
    addTemplatePositionRow({ label: "Nombre vendedor", token: "{{vendedor.nombre}}", x: 75, y: 90 });
  }
  if (el.templatePageImage) el.templatePageImage.value = "";
  if (el.templateHeaderImage) el.templateHeaderImage.value = "";
  if (el.templateFooterImage) el.templateFooterImage.value = "";
  renderTemplateSelect("");
  refreshTemplatePreview();
}

function collectTemplateFormulas() {
  if (!el.templateFormulaBody) return [];
  const rows = Array.from(el.templateFormulaBody.querySelectorAll(".templateFormulaRow"));
  const formulas = [];
  for (const row of rows) {
    const key = String(row.querySelector(".formulaKey")?.value || "").trim();
    const expression = String(row.querySelector(".formulaExpr")?.value || "").trim();
    if (!key && !expression) continue;
    formulas.push({ key, expression });
  }
  return formulas;
}

function collectTemplateDraft() {
  const name = String(el.templateName?.value || "").trim();
  const header = String(el.templateHeader?.value || "").trim();
  const body = String(el.templateBody?.value || "").trim();
  const footer = String(el.templateFooter?.value || "").trim();
  const positionedFields = collectTemplatePositionedFields();
  const formulas = collectTemplateFormulas();
  const roomRates = parseRoomRatesFromEditor().items;
  const signatureDefaults = getTemplateSignatureDefaultsFromEditor(positionedFields);
  return { name, header, body, footer, formulas, positionedFields, signatureDefaults, roomRates, assets: { ...templateAssetsDraft } };
}

function parseRoomRatesFromEditor() {
  const roomRatesRaw = String(el.templateRoomRates?.value || "").trim();
  if (!roomRatesRaw) return { ok: true, items: [] };
  try {
    const parsed = JSON.parse(roomRatesRaw);
    if (!Array.isArray(parsed)) return { ok: false, items: [] };
    const items = parsed
      .map((r) => ({
        habitacion: String(r?.habitacion || "").trim(),
        precio: Number(r?.precio || 0),
      }))
      .filter((r) => r.habitacion);
    return { ok: true, items };
  } catch (_) {
    return { ok: false, items: [] };
  }
}

function applyTemplateTokenInsertion(token) {
  const editor = activeTemplateEditor || el.templateBody || el.templateHeader || el.templateFooter;
  if (!editor) return;
  const value = String(editor.value || "");
  const start = Number(editor.selectionStart || 0);
  const end = Number(editor.selectionEnd || start);
  const next = `${value.slice(0, start)}${token}${value.slice(end)}`;
  editor.value = next;
  const cursor = start + token.length;
  editor.focus();
  editor.setSelectionRange(cursor, cursor);
  refreshTemplatePreview();
}

function refreshTemplatePreview() {
  if (!el.templatePreview) return;
  const draft = collectTemplateDraft();
  const samples = {
    "{{cliente.nombre}}": "Juan Perez",
    "{{cliente.firma}}": "Firma Cliente",
    "{{cliente.telefono}}": "5555-0101",
    "{{cliente.direccion}}": "Zona 10, Ciudad",
    "{{vendedor.nombre}}": "Ana Lopez",
    "{{vendedor.firma}}": "Firma Vendedor",
    "{{vendedor.telefono}}": "5555-0202",
    "{{vendedor.direccion}}": "Sucursal Central",
    "{{fecha.hoy}}": toISODate(new Date()),
    "{{fecha.evento}}": toISODate(addDays(new Date(), 15)),
    "{{evento.nombre}}": "Cena corporativa",
    "{{evento.pax}}": "120",
    "{{institucion.nombre}}": "Instituto Ejemplo",
    "{{monto.total}}": "Q 12,500.00",
    "{{tabla.habitaciones}}": "[Tabla habitaciones]",
  };
  const material = [draft.header, draft.body, draft.footer].filter(Boolean).join("\n\n");
  let preview = material || "Sin contenido.";
  for (const [token, sample] of Object.entries(samples)) {
    preview = preview.split(token).join(sample);
  }
  preview = preview.replace(/\{\{\s*formula:([a-zA-Z0-9_.-]+)\s*\}\}/g, (_, key) => {
    const row = draft.formulas.find((f) => f.key === key);
    if (!row) return `[formula:${key}]`;
    return `[${row.key} = ${row.expression || "sin expresion"}]`;
  });
  if (preview.includes("[Tabla habitaciones]") && draft.roomRates.length) {
    const tableRows = draft.roomRates.map((r) => `${r.habitacion}: Q ${Number(r.precio || 0).toFixed(2)}`).join(" | ");
    preview = preview.replace("[Tabla habitaciones]", tableRows);
  }
  el.templatePreview.textContent = preview;
  renderTemplatePagePreview(draft);
}

function getTemplatePositionRows() {
  return Array.from(el.templatePositionBody?.querySelectorAll(".templatePositionRow") || []);
}

function setActiveTemplatePosition(index) {
  templateUiState.activePositionIndex = Number.isInteger(index) ? index : -1;
  const rows = getTemplatePositionRows();
  for (let i = 0; i < rows.length; i++) {
    rows[i].classList.toggle("active", i === templateUiState.activePositionIndex);
  }
  const markers = Array.from(el.templatePagePreview?.querySelectorAll(".templatePosBadge") || []);
  for (let i = 0; i < markers.length; i++) {
    markers[i].classList.toggle("active", i === templateUiState.activePositionIndex);
  }
}

function updateTemplatePositionRow(index, x, y) {
  const rows = getTemplatePositionRows();
  const row = rows[index];
  if (!row) return;
  const xInput = row.querySelector(".posX");
  const yInput = row.querySelector(".posY");
  const isSignature = row.dataset.signature === "1";
  let maxX = 100;
  let maxY = 100;
  if (isSignature) {
    const width = clamp(Number(row.querySelector(".posW")?.value || TEMPLATE_SIGNATURE_FALLBACK_W_PCT), TEMPLATE_SIGNATURE_MIN_W_PCT, TEMPLATE_SIGNATURE_MAX_W_PCT);
    const height = clamp(Number(row.querySelector(".posH")?.value || TEMPLATE_SIGNATURE_FALLBACK_H_PCT), TEMPLATE_SIGNATURE_MIN_H_PCT, TEMPLATE_SIGNATURE_MAX_H_PCT);
    maxX = Math.max(0, 100 - width);
    maxY = Math.max(0, 100 - height);
  }
  if (xInput) xInput.value = clamp(Number(x), 0, maxX).toFixed(1);
  if (yInput) yInput.value = clamp(Number(y), 0, maxY).toFixed(1);
}

function updateTemplatePositionSizeRow(index, w, h) {
  const rows = getTemplatePositionRows();
  const row = rows[index];
  if (!row) return;
  const wInput = row.querySelector(".posW");
  const hInput = row.querySelector(".posH");
  const isSignature = row.dataset.signature === "1";
  const minW = isSignature ? TEMPLATE_SIGNATURE_MIN_W_PCT : 4;
  const minH = isSignature ? TEMPLATE_SIGNATURE_MIN_H_PCT : 2;
  if (wInput) wInput.value = clamp(Number(w), minW, isSignature ? TEMPLATE_SIGNATURE_MAX_W_PCT : 95).toFixed(1);
  if (hInput) hInput.value = clamp(Number(h), minH, isSignature ? TEMPLATE_SIGNATURE_MAX_H_PCT : 60).toFixed(1);
}

function updateTemplateMarkerPosition(index, x, y) {
  const marker = el.templatePagePreview?.querySelector(`.templatePosBadge[data-index="${index}"]`);
  if (!marker) return;
  marker.style.left = `${clamp(Number(x), 0, 100)}%`;
  marker.style.top = `${clamp(Number(y), 0, 100)}%`;
}

function updateTemplateMarkerSize(index, w, h) {
  const marker = el.templatePagePreview?.querySelector(`.templatePosBadge[data-index="${index}"]`);
  if (!marker || !marker.classList.contains("signatureBox")) return;
  marker.style.width = `${clamp(Number(w), TEMPLATE_SIGNATURE_MIN_W_PCT, TEMPLATE_SIGNATURE_MAX_W_PCT)}%`;
  marker.style.height = `${clamp(Number(h), TEMPLATE_SIGNATURE_MIN_H_PCT, TEMPLATE_SIGNATURE_MAX_H_PCT)}%`;
}

function isTemplateSignatureToken(token) {
  const t = String(token || "").toLowerCase().trim();
  return t === "{{cliente.firma}}" || t === "{{vendedor.firma}}" || t.includes(".firma");
}

function applyTemplatePositionRowSignatureState(row) {
  if (!row) return;
  const token = String(row.querySelector(".posToken")?.value || "").trim();
  const isSignature = isTemplateSignatureToken(token);
  row.dataset.signature = isSignature ? "1" : "0";
  row.classList.toggle("signatureRow", isSignature);
  const wInput = row.querySelector(".posW");
  const hInput = row.querySelector(".posH");
  if (wInput) {
    wInput.disabled = !isSignature;
    wInput.min = String(isSignature ? TEMPLATE_SIGNATURE_MIN_W_PCT : 4);
    wInput.max = String(isSignature ? TEMPLATE_SIGNATURE_MAX_W_PCT : 95);
    if (isSignature) wInput.value = String(clamp(Number(wInput.value || TEMPLATE_SIGNATURE_FALLBACK_W_PCT), TEMPLATE_SIGNATURE_MIN_W_PCT, TEMPLATE_SIGNATURE_MAX_W_PCT));
  }
  if (hInput) {
    hInput.disabled = !isSignature;
    hInput.min = String(isSignature ? TEMPLATE_SIGNATURE_MIN_H_PCT : 2);
    hInput.max = String(isSignature ? TEMPLATE_SIGNATURE_MAX_H_PCT : 60);
    if (isSignature) hInput.value = String(clamp(Number(hInput.value || TEMPLATE_SIGNATURE_FALLBACK_H_PCT), TEMPLATE_SIGNATURE_MIN_H_PCT, TEMPLATE_SIGNATURE_MAX_H_PCT));
  }
}

function renderTemplatePositionMarkers(fields = []) {
  const layer = el.templatePagePreview?.querySelector(".templateMarkerLayer");
  if (!layer) return;
  layer.innerHTML = fields
    .map((p, idx) => {
      const left = clamp(Number(p.x), 0, 100);
      const top = clamp(Number(p.y), 0, 100);
      const isSignature = p.isSignature === true || isTemplateSignatureToken(p.token);
      const width = clamp(Number(p.w || TEMPLATE_SIGNATURE_FALLBACK_W_PCT), isSignature ? TEMPLATE_SIGNATURE_MIN_W_PCT : 4, isSignature ? TEMPLATE_SIGNATURE_MAX_W_PCT : 95);
      const height = clamp(Number(p.h || TEMPLATE_SIGNATURE_FALLBACK_H_PCT), isSignature ? TEMPLATE_SIGNATURE_MIN_H_PCT : 2, isSignature ? TEMPLATE_SIGNATURE_MAX_H_PCT : 60);
      const label = escapeHtml(p.label || p.token || "Campo");
      const title = escapeHtml(p.token || "");
      const fontSize = clamp(Number(p.fontSize || 12), 8, 72);
      const fontFamily = escapeHtml(String(p.fontFamily || "Arial"));
      const fontWeight = p.bold ? 800 : 600;
      const fontStyle = p.italic ? "italic" : "normal";
      if (isSignature) {
        return `<div class="templatePosBadge signatureBox${idx === templateUiState.activePositionIndex ? " active" : ""}" data-signature="1" data-index="${idx}" style="left:${left}%;top:${top}%;width:${width}%;height:${height}%" title="${title}"><span class="signatureBoxLabel">${label}</span><span class="templateResizeHandle" data-action="resize"></span></div>`;
      }
      return `<div class="templatePosBadge${idx === templateUiState.activePositionIndex ? " active" : ""}" data-index="${idx}" style="left:${left}%;top:${top}%;font-size:${fontSize}px;font-family:${fontFamily}, sans-serif;font-weight:${fontWeight};font-style:${fontStyle}" title="${title}">${label}</div>`;
    })
    .join("");
}

function syncTemplateSignatureBounds(index) {
  const rows = getTemplatePositionRows();
  const row = rows[index];
  if (!row || row.dataset.signature !== "1") return;
  const xInput = row.querySelector(".posX");
  const yInput = row.querySelector(".posY");
  const wInput = row.querySelector(".posW");
  const hInput = row.querySelector(".posH");
  const w = clamp(Number(wInput?.value || TEMPLATE_SIGNATURE_FALLBACK_W_PCT), TEMPLATE_SIGNATURE_MIN_W_PCT, TEMPLATE_SIGNATURE_MAX_W_PCT);
  const h = clamp(Number(hInput?.value || TEMPLATE_SIGNATURE_FALLBACK_H_PCT), TEMPLATE_SIGNATURE_MIN_H_PCT, TEMPLATE_SIGNATURE_MAX_H_PCT);
  const maxX = Math.max(0, 100 - w);
  const maxY = Math.max(0, 100 - h);
  if (xInput) xInput.value = clamp(Number(xInput.value || 0), 0, maxX).toFixed(1);
  if (yInput) yInput.value = clamp(Number(yInput.value || 0), 0, maxY).toFixed(1);
}

function syncTemplateDocumentHeight() {
  if (!el.templatePagePreview) return;
  const doc = el.templatePagePreview.querySelector(".templatePageDocument");
  const pageLayer = el.templatePagePreview.querySelector(".templatePageLayer.page");
  const markerLayer = el.templatePagePreview.querySelector(".templateMarkerLayer");
  if (!doc || !pageLayer || !markerLayer) return;
  syncTemplatePreviewLetterFrame();
  const h = Math.max(420, pageLayer.scrollHeight || pageLayer.offsetHeight || 420);
  doc.style.height = `${h}px`;
  markerLayer.style.height = `${h}px`;
}

function syncTemplatePreviewLetterFrame() {
  if (!el.templatePagePreview) return;
  const doc = el.templatePagePreview.querySelector(".templatePageDocument");
  if (!doc) return;
  const pad = 18;
  const availW = Math.max(320, (el.templatePagePreview.clientWidth || 760) - pad);
  const availH = Math.max(420, (el.templatePagePreview.clientHeight || 680) - pad);
  const ratio = TEMPLATE_COORD_BASE_W_PT / TEMPLATE_COORD_BASE_H_PT;
  const fitByH = availH * ratio;
  const frameW = Math.max(320, Math.min(availW, fitByH));
  doc.style.width = `${Math.round(frameW)}px`;
  doc.style.margin = "0 auto";
}

async function renderTemplatePdfToLayer(pdfDataUrl) {
  const layer = el.templatePagePreview?.querySelector(".templatePageLayer.page");
  if (!layer) return;
  const key = String(pdfDataUrl || "");
  if (!key) {
    templateUiState.pdfRenderKey = "";
    layer.innerHTML = "";
    syncTemplateDocumentHeight();
    return;
  }
  if (templateUiState.pdfRenderKey === key && layer.querySelector("canvas")) return;
  templateUiState.pdfRenderKey = key;

  if (!window.pdfjsLib || typeof window.pdfjsLib.getDocument !== "function") {
    layer.innerHTML = `<div class="templatePosBadge" style="left:50%;top:50%;">No se pudo cargar visor PDF</div>`;
    return;
  }
  try {
    if (!window.pdfjsLib.GlobalWorkerOptions.workerSrc) {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js";
    }
    const b64 = String(pdfDataUrl).split(",")[1] || "";
    const binary = atob(b64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
    const loadingTask = window.pdfjsLib.getDocument({ data: bytes });
    const pdf = await loadingTask.promise;
    layer.innerHTML = "";
    syncTemplatePreviewLetterFrame();
    const doc = el.templatePagePreview?.querySelector(".templatePageDocument");
    const desiredWidth = Math.max(320, doc?.clientWidth || layer.clientWidth || el.templatePagePreview?.clientWidth || 760);
    for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
      const page = await pdf.getPage(pageNum);
      const baseViewport = page.getViewport({ scale: 1 });
      const scale = desiredWidth / baseViewport.width;
      const viewport = page.getViewport({ scale });
      const canvas = document.createElement("canvas");
      canvas.width = Math.round(viewport.width);
      canvas.height = Math.round(viewport.height);
      canvas.dataset.page = String(pageNum);
      const ctx = canvas.getContext("2d");
      await page.render({ canvasContext: ctx, viewport }).promise;
      layer.appendChild(canvas);
    }
    if (templateUiState.pdfRenderKey !== key) return;
    syncTemplateDocumentHeight();
    renderTemplatePositionMarkers(collectTemplatePositionedFields());
  } catch (_) {
    layer.innerHTML = `<div class="templatePosBadge" style="left:50%;top:50%;">Error al renderizar PDF</div>`;
    syncTemplateDocumentHeight();
  }
}

function renderTemplatePagePreview(draft = null) {
  if (!el.templatePagePreview) return;
  const data = draft || collectTemplateDraft();
  let pageLayer = el.templatePagePreview.querySelector(".templatePageLayer.page");
  let headerLayer = el.templatePagePreview.querySelector(".templatePageLayer.header");
  let footerLayer = el.templatePagePreview.querySelector(".templatePageLayer.footer");
  let markerLayer = el.templatePagePreview.querySelector(".templateMarkerLayer");
  if (!pageLayer || !headerLayer || !footerLayer || !markerLayer) {
    el.templatePagePreview.innerHTML = `
      <div class="templatePageDocument">
        <div class="templatePageLayer page"></div>
        <div class="templatePageLayer header"></div>
        <div class="templatePageLayer footer"></div>
        <div class="templateMarkerLayer"></div>
      </div>
    `;
    pageLayer = el.templatePagePreview.querySelector(".templatePageLayer.page");
    headerLayer = el.templatePagePreview.querySelector(".templatePageLayer.header");
    footerLayer = el.templatePagePreview.querySelector(".templatePageLayer.footer");
    markerLayer = el.templatePagePreview.querySelector(".templateMarkerLayer");
  }
  if (headerLayer) {
    headerLayer.style.backgroundImage = data.assets?.headerImage ? `url('${data.assets.headerImage}')` : "none";
  }
  if (footerLayer) {
    footerLayer.style.backgroundImage = data.assets?.footerImage ? `url('${data.assets.footerImage}')` : "none";
  }
  renderTemplatePositionMarkers(data.positionedFields || []);
  syncTemplateDocumentHeight();
  renderTemplatePdfToLayer(data.assets?.pagePdf || "").catch(() => { });
}

function beginTemplateMarkerDrag(e, index) {
  if (!el.templatePagePreview) return;
  const markerLayer = el.templatePagePreview.querySelector(".templateMarkerLayer");
  const dragRect = markerLayer?.getBoundingClientRect() || el.templatePagePreview.getBoundingClientRect();
  const marker = el.templatePagePreview.querySelector(`.templatePosBadge[data-index="${index}"]`);
  if (!marker) return;
  if (e.target?.closest(".templateResizeHandle")) return;
  const isSignature = marker.classList.contains("signatureBox");
  const markerRect = marker.getBoundingClientRect();
  templateUiState.dragging = {
    index,
    rect: dragRect,
    markerW: isSignature ? clamp((markerRect.width / Math.max(1, dragRect.width)) * 100, TEMPLATE_SIGNATURE_MIN_W_PCT, TEMPLATE_SIGNATURE_MAX_W_PCT) : 0,
    markerH: isSignature ? clamp((markerRect.height / Math.max(1, dragRect.height)) * 100, TEMPLATE_SIGNATURE_MIN_H_PCT, TEMPLATE_SIGNATURE_MAX_H_PCT) : 0,
  };
  setActiveTemplatePosition(index);
  if (marker) marker.classList.add("dragging");
  e.preventDefault();
}

function onTemplateMarkerDragMove(e) {
  const drag = templateUiState.dragging;
  if (!drag) return;
  const markerLayer = el.templatePagePreview?.querySelector(".templateMarkerLayer");
  const rect = markerLayer?.getBoundingClientRect() || drag.rect;
  const maxX = 100 - Math.max(0, Number(drag.markerW || 0));
  const maxY = 100 - Math.max(0, Number(drag.markerH || 0));
  const x = clamp(((e.clientX - rect.left) / Math.max(1, rect.width)) * 100, 0, Math.max(0, maxX));
  const y = clamp(((e.clientY - rect.top) / Math.max(1, rect.height)) * 100, 0, Math.max(0, maxY));
  updateTemplatePositionRow(drag.index, x, y);
  updateTemplateMarkerPosition(drag.index, x, y);
}

function endTemplateMarkerDrag() {
  if (!templateUiState.dragging) return;
  const drag = templateUiState.dragging;
  const marker = el.templatePagePreview?.querySelector(`.templatePosBadge[data-index="${drag.index}"]`);
  if (marker) marker.classList.remove("dragging");
  templateUiState.dragging = null;
}

function beginTemplateMarkerResize(e, index) {
  if (!el.templatePagePreview) return;
  const markerLayer = el.templatePagePreview.querySelector(".templateMarkerLayer");
  const rect = markerLayer?.getBoundingClientRect() || el.templatePagePreview.getBoundingClientRect();
  const marker = el.templatePagePreview.querySelector(`.templatePosBadge[data-index="${index}"]`);
  if (!marker || !marker.classList.contains("signatureBox")) return;
  const markerRect = marker.getBoundingClientRect();
  templateUiState.resizing = {
    index,
    rect,
    startClientX: e.clientX,
    startClientY: e.clientY,
    startW: clamp((markerRect.width / Math.max(1, rect.width)) * 100, TEMPLATE_SIGNATURE_MIN_W_PCT, TEMPLATE_SIGNATURE_MAX_W_PCT),
    startH: clamp((markerRect.height / Math.max(1, rect.height)) * 100, TEMPLATE_SIGNATURE_MIN_H_PCT, TEMPLATE_SIGNATURE_MAX_H_PCT),
  };
  setActiveTemplatePosition(index);
  marker.classList.add("resizing");
  e.preventDefault();
  e.stopPropagation();
}

function onTemplateMarkerResizeMove(e) {
  const resize = templateUiState.resizing;
  if (!resize) return;
  const rows = getTemplatePositionRows();
  const row = rows[resize.index];
  if (!row) return;
  const rect = resize.rect;
  const dxPct = ((e.clientX - resize.startClientX) / Math.max(1, rect.width)) * 100;
  const dyPct = ((e.clientY - resize.startClientY) / Math.max(1, rect.height)) * 100;
  const posX = clamp(Number(row.querySelector(".posX")?.value || 0), 0, 100);
  const posY = clamp(Number(row.querySelector(".posY")?.value || 0), 0, 100);
  const maxW = Math.min(TEMPLATE_SIGNATURE_MAX_W_PCT, Math.max(TEMPLATE_SIGNATURE_MIN_W_PCT, 100 - posX));
  const maxH = Math.min(TEMPLATE_SIGNATURE_MAX_H_PCT, Math.max(TEMPLATE_SIGNATURE_MIN_H_PCT, 100 - posY));
  const nextW = clamp(resize.startW + dxPct, TEMPLATE_SIGNATURE_MIN_W_PCT, maxW);
  const nextH = clamp(resize.startH + dyPct, TEMPLATE_SIGNATURE_MIN_H_PCT, maxH);
  updateTemplatePositionSizeRow(resize.index, nextW, nextH);
  updateTemplateMarkerSize(resize.index, nextW, nextH);
}

function endTemplateMarkerResize() {
  if (!templateUiState.resizing) return;
  const resize = templateUiState.resizing;
  const marker = el.templatePagePreview?.querySelector(`.templatePosBadge[data-index="${resize.index}"]`);
  if (marker) marker.classList.remove("resizing");
  templateUiState.resizing = null;
}

function openTemplateBuilderModal() {
  if (!el.templateBackdrop) return;
  closeSettingsPanel();
  clearTemplateEditor();
  activeTemplateEditor = el.templateBody || null;
  el.templateBackdrop.hidden = false;
  setTimeout(() => el.templateName?.focus(), 0);
}

function closeTemplateBuilderModal() {
  if (!el.templateBackdrop) return;
  el.templateBackdrop.hidden = true;
  activeTemplateEditor = null;
  endTemplateMarkerDrag();
  endTemplateMarkerResize();
}

function saveTemplateFromBuilder() {
  const draft = collectTemplateDraft();
  if (!draft.name) return toast("Nombre de plantilla es obligatorio.");
  const roomRatesCheck = parseRoomRatesFromEditor();
  if (!roomRatesCheck.ok) return toast("JSON de tabla habitaciones invalido.");

  const seen = new Set();
  for (const f of draft.formulas) {
    if (!f.key) return toast("Cada formula debe tener nombre de campo.");
    if (seen.has(f.key)) return toast("No repitas nombres de campos de formula.");
    seen.add(f.key);
  }

  const existingByName = quickTemplates.find(
    (t) => t.name.toLowerCase() === draft.name.toLowerCase() && t.id !== activeTemplateId
  );
  if (existingByName) return toast("Ya existe una plantilla con ese nombre.");

  const nowIso = new Date().toISOString();
  const payload = {
    id: activeTemplateId || uid(),
    name: draft.name,
    header: draft.header,
    body: draft.body,
    footer: draft.footer,
    assets: { ...draft.assets },
    positionedFields: draft.positionedFields,
    signatureDefaults: draft.signatureDefaults,
    roomRates: roomRatesCheck.items,
    formulas: draft.formulas,
    createdAt: nowIso,
    updatedAt: nowIso,
  };

  const idx = quickTemplates.findIndex((t) => t.id === payload.id);
  if (idx >= 0) {
    payload.createdAt = quickTemplates[idx].createdAt || nowIso;
    quickTemplates[idx] = payload;
  } else {
    quickTemplates.push(payload);
  }
  saveQuickTemplates();
  renderTemplateSelect(payload.id);
  activeTemplateId = payload.id;
  toast("Plantilla guardada.");
}

function loadSelectedTemplateIntoBuilder() {
  const id = String(el.templateSelect?.value || "").trim();
  if (!id) return toast("Selecciona una plantilla para cargar.");
  const found = quickTemplates.find((t) => t.id === id);
  if (!found) return toast("Plantilla no encontrada.");
  setTemplateEditorContent(found);
  activeTemplateEditor = el.templateBody || null;
  toast("Plantilla cargada.");
}

async function deleteSelectedTemplateFromBuilder() {
  const id = String(el.templateSelect?.value || activeTemplateId || "").trim();
  if (!id) return toast("Selecciona una plantilla para eliminar.");
  const found = quickTemplates.find((t) => t.id === id);
  if (!found) return toast("Plantilla no encontrada.");
  const ok = await modernConfirm({
    title: "Eliminar plantilla",
    message: `Esta seguro de eliminar la plantilla "${found.name}"?`,
    confirmText: "Si, eliminar",
    cancelText: "No",
  });
  if (!ok) return;
  quickTemplates = quickTemplates.filter((t) => t.id !== id);
  saveQuickTemplates();
  clearTemplateEditor();
  toast(`Plantilla eliminada: ${found.name}`);
}

function setSettingsPanelOpen(open) {
  if (!el.settingsPanel || !el.btnSettings) return;
  el.settingsPanel.hidden = !open;
  el.btnSettings.setAttribute("aria-expanded", open ? "true" : "false");
  if (open) {
    if (!el.settingsPanel.hasAttribute("tabindex")) {
      el.settingsPanel.setAttribute("tabindex", "-1");
    }
    setTimeout(() => {
      try { el.settingsPanel.focus(); } catch (_) { }
    }, 0);
  } else {
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

function matchesAliases(value, aliases = []) {
  const base = normalizeBucketKey(value);
  if (!base) return false;
  return aliases.some((a) => base.includes(normalizeBucketKey(a)));
}

function moneyGT(v) {
  return `Q ${Number(v || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
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
      ].join(" ").toLowerCase();
      if (!blob.includes(search)) return false;
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
    .card{ background:#fff; border:1px solid #c5d4ea; border-radius:10px; overflow:hidden; }
    .meta{ padding:10px 14px; border-top:1px solid #bfd3ee; border-bottom:1px solid #bfd3ee; background:#eaf3ff; font-size:12px; }
    .meta div{ margin:2px 0; }
    table{ width:100%; border-collapse:collapse; }
    th,td{ border:1px solid #c7d5ea; padding:6px 7px; font-size:11px; white-space:nowrap; }
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
}

function weekInputFromDate(date) {
  const d = stripTime(date);
  const day = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - day + 3);
  const year = d.getFullYear();
  const firstThursday = new Date(year, 0, 4);
  const firstDay = (firstThursday.getDay() + 6) % 7;
  firstThursday.setDate(firstThursday.getDate() - firstDay + 3);
  const week = 1 + Math.round((d.getTime() - firstThursday.getTime()) / (7 * 24 * 60 * 60 * 1000));
  return `${year}-W${pad2(week)}`;
}

function mondayFromWeekInput(value) {
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
  return `<button type="button" class="occupancyMenuMontajeLinkBtn" data-event-id="${escapeHtml(String(row.eventId || ""))}" data-quote-version="${escapeHtml(String(version || ""))}">${escapeHtml(text)}</button>`;
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
      <div class="occupancyEventGrid">
        <span><b>Evento:</b> ${escapeHtml(r.eventName || "-")}</span>
        <span><b>Horario:</b> ${escapeHtml(r.startTime)} - ${escapeHtml(r.endTime)}</span>
        <span><b>Salon:</b> ${escapeHtml(r.salon || "-")}</span>
        <span><b>Institucion:</b> ${escapeHtml(r.company || "-")}</span>
        <span><b>Encargado:</b> ${escapeHtml(r.manager || "-")}</span>
        <span><b>Vendedor:</b> ${escapeHtml(r.seller || "-")}</span>
        <span><b>PAX:</b> ${escapeHtml(String(r.pax || 0))}</span>
        <span><b>Ult. cotizacion:</b> ${buildOccupancyQuoteActionHtml(r)}</span>
        <span><b>Ult. informe:</b> ${buildOccupancyMenuMontajeActionHtml(r)}</span>
        <span><b>Total evento:</b> ${escapeHtml(moneyGT(r.totalEvent || 0))}</span>
        <span><b>Ingreso dia:</b> ${escapeHtml(moneyGT(r.incomePerDay || 0))}</span>
        <span><b>Ingreso salon-dia:</b> ${escapeHtml(moneyGT(r.incomePerSalonDay || 0))}</span>
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
    tr.innerHTML = `<td colspan="16">Sin eventos Confirmados/Pre reserva para esta semana.</td>`;
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
    .card{ background:#fff; border:1px solid #c5d4ea; border-radius:10px; overflow:hidden; }
    .titleCell{ border:1px solid #c7d5ea; background:#d8e3f3; color:#000; font-weight:800; font-size:20px; padding:12px 14px; text-transform:uppercase; }
    .meta{ padding:10px 14px; border-top:1px solid #bfd3ee; border-bottom:1px solid #bfd3ee; background:#eaf3ff; font-size:12px; }
    .meta div{ margin:2px 0; }
    table{ width:100%; border-collapse:collapse; }
    th,td{ border:1px solid #c7d5ea; padding:6px 7px; font-size:11px; white-space:nowrap; }
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
          <th>Estado</th><th>PAX</th><th>Fecha evento</th><th>Hora inicio</th><th>Hora final</th><th>Evento</th><th>Salon</th><th>Institucion</th><th>Encargado evento</th><th>Vendedor</th><th>Ultima cotizacion enviada</th><th>Ultimo informe menu/montaje</th><th>Total evento</th><th>Ingreso dia</th><th>Ingreso salon-dia</th><th>Ultima modificacion</th>
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
  const { monday, sunday } = getOccupancyWeekRange();
  if (el.occupancyReportSubtitle) {
    el.occupancyReportSubtitle.textContent = `Semana ${toISODate(monday)} a ${toISODate(sunday)} (Lunes a Domingo)`;
  }
  renderOccupancyReportTable();
  el.occupancyReportBackdrop.hidden = false;
}

function closeOccupancyReportModal() {
  if (!el.occupancyReportBackdrop) return;
  el.occupancyReportBackdrop.hidden = true;
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
      .catch(() => setLoginError("No se pudo cargar usuarios desde MariaDB."));
  });

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
  renderTopbarWelcome();
  refreshTopbarReminders();
  el.loginScreen.hidden = true;
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
    window.flatpickr(input, {
      enableTime: true,
      noCalendar: true,
      dateFormat: "H:i",
      time_24hr: true,
      minuteIncrement: SNAP_MINUTES,
      allowInput: true,
      disableMobile: true,
    });
    return;
  }
  input.type = "time";
  input.step = String(SNAP_MINUTES * 60);
}

function addSlotRow(slot = null) {
  const row = document.createElement("tr");
  row.className = "slotRow";
  const salon = slot?.salon || "";
  const start = slot?.startTime || "";
  const end = slot?.endTime || "";
  row.innerHTML = `
    <td><select class="quoteInput slotRoom">${salonOptionsHtml(salon, true)}</select></td>
    <td><input class="quoteInput slotStart" type="text" inputmode="numeric" placeholder="HH:mm" value="${escapeHtml(start)}" /></td>
    <td><input class="quoteInput slotEnd" type="text" inputmode="numeric" placeholder="HH:mm" value="${escapeHtml(end)}" /></td>
    <td><button type="button" class="btnDanger slotRemoveBtn">X</button></td>
  `;
  el.slotsBody.appendChild(row);
  initModernTimePicker(row.querySelector(".slotStart"));
  initModernTimePicker(row.querySelector(".slotEnd"));
}

function rerenderSlotRoomOptions() {
  if (!el.slotsBody) return;
  for (const row of Array.from(el.slotsBody.querySelectorAll(".slotRow"))) {
    const select = row.querySelector(".slotRoom");
    if (!select) continue;
    const current = select.value;
    select.innerHTML = salonOptionsHtml(current, true);
    if (!select.value && select.options.length) select.value = select.options[0].value;
  }
}

function getSlotsFromForm() {
  const rows = Array.from(el.slotsBody.querySelectorAll(".slotRow"));
  return rows.map(row => ({
    salon: row.querySelector(".slotRoom")?.value || "",
    startTime: row.querySelector(".slotStart")?.value || "",
    endTime: row.querySelector(".slotEnd")?.value || "",
  }));
}

function syncHiddenTimesFromFirstSlot() {
  const first = getSlotsFromForm()[0];
  el.startTime.value = first?.startTime || "";
  el.endTime.value = first?.endTime || "";
}

function renderCompaniesSelect(selectedId = null) {
  if (!el.quoteCompany) return;
  const previousValue = String(el.quoteCompany.value || "").trim();
  el.quoteCompany.innerHTML = "";
  if (el.companiesList) el.companiesList.innerHTML = "";
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
  if (!el.quoteCompany.value && el.quoteCompany.options.length) {
    el.quoteCompany.value = el.quoteCompany.options[0].value;
  }
  const finalCompanyId = String(el.quoteCompany.value || previousValue || "").trim();
  enhanceSelectControl(el.quoteCompany, true);
  syncEnhancedSelectValue(el.quoteCompany, finalCompanyId);
  const selectedCompany = (state.companies || []).find(c => c.id === el.quoteCompany.value);
  if (el.quoteCompanySearch) {
    el.quoteCompanySearch.value = selectedCompany?.name || "";
    refreshCompanySuggestions(el.quoteCompanySearch.value);
  }
  renderQuoteManagerSelect(el.quoteCompany.value, quoteDraft?.managerId || null);
}

function refreshCompanySuggestions(rawTerm = "") {
  if (!el.companiesList) return;
  const term = String(rawTerm || "").trim().toLowerCase();
  const disabledCompanies = new Set((state.disabledCompanies || []).map((x) => String(x)));
  const companies = (state.companies || []).filter((c) => !disabledCompanies.has(String(c.id)));
  let source = companies;
  if (term) {
    source = findCompanyMatches(term);
  }
  const seen = new Set();
  el.companiesList.innerHTML = "";
  for (const c of source.slice(0, 20)) {
    const name = String(c.name || "").trim();
    if (!name) continue;
    const key = name.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    const opt = document.createElement("option");
    opt.value = name;
    opt.textContent = c.businessName ? `${name} - ${c.businessName}` : name;
    el.companiesList.appendChild(opt);
  }
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
  const term = String(rawTerm || "").trim().toLowerCase();
  if (!term) return [];
  const disabledCompanies = new Set((state.disabledCompanies || []).map((x) => String(x)));
  const companies = (state.companies || []).filter((c) => !disabledCompanies.has(String(c.id)));
  return companies.filter((c) => {
    const name = String(c.name || "").toLowerCase();
    const businessName = String(c.businessName || "").toLowerCase();
    const nit = String(c.nit || "").toLowerCase();
    return name.includes(term) || businessName.includes(term) || nit.includes(term);
  });
}

function resolveCompanyFromSearch(rawTerm) {
  const term = String(rawTerm || "").trim().toLowerCase();
  if (!term) return null;
  const disabledCompanies = new Set((state.disabledCompanies || []).map((x) => String(x)));
  const companies = (state.companies || []).filter((c) => !disabledCompanies.has(String(c.id)));
  const exactName = companies.find((c) => String(c.name || "").toLowerCase() === term);
  if (exactName) return exactName;
  const exactBusiness = companies.find((c) => String(c.businessName || "").toLowerCase() === term);
  if (exactBusiness) return exactBusiness;
  const matches = findCompanyMatches(term);
  return matches.length === 1 ? matches[0] : null;
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
  fillQuoteHeaderFields(true);
}

function findServiceMatches(rawTerm) {
  const term = String(rawTerm || "").trim().toLowerCase();
  if (!term) return [];
  const disabledServices = new Set((state.disabledServices || []).map((x) => String(x)));
  const services = (state.services || []).filter((s) => !disabledServices.has(String(s.id)));
  return services.filter((s) => {
    const name = String(s.name || "").toLowerCase();
    const desc = String(s.description || "").toLowerCase();
    return name.includes(term) || desc.includes(term);
  });
}

function resolveServiceFromSearch(rawTerm) {
  const term = String(rawTerm || "").trim().toLowerCase();
  if (!term) return null;
  const disabledServices = new Set((state.disabledServices || []).map((x) => String(x)));
  const services = (state.services || []).filter((s) => !disabledServices.has(String(s.id)));
  const exactName = services.find((s) => String(s.name || "").toLowerCase() === term);
  if (exactName) return exactName;
  const exactDesc = services.find((s) => String(s.description || "").toLowerCase() === term);
  if (exactDesc) return exactDesc;
  const matches = findServiceMatches(term);
  return matches.length === 1 ? matches[0] : null;
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
    opt.textContent = c.nombre;
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
  const list = Number.isFinite(catId)
    ? catalogoSubcategoriasServicio.filter((s) => Number(s.id_categoria) === catId)
    : [];
  for (const s of list) {
    const opt = document.createElement("option");
    opt.value = String(s.id);
    opt.textContent = s.nombre;
    el.serviceSubcategory.appendChild(opt);
  }
}

async function syncServiceCatalogFromDb() {
  try {
    const categoriasUrl = buildApiUrlFromStateUrl(activeApiStateUrl, "categorias-servicio");
    const categoriasRes = await fetch(categoriasUrl, { cache: "no-store" });
    if (categoriasRes.ok) {
      const payloadCategorias = await categoriasRes.json();
      const categorias = Array.isArray(payloadCategorias?.categorias) ? payloadCategorias.categorias : [];
      catalogoCategoriasServicio = categorias
        .map((c) => ({ id: Number(c.id), nombre: String(c.nombre || "").trim() }))
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
        }))
        .filter((s) => Number.isFinite(s.id) && Number.isFinite(s.id_categoria) && s.nombre);
    }
  } catch (_) { }

  renderCategoriasServicioSelect();
  renderSubcategoriasServicioSelect(Number(el.serviceCategory?.value || NaN));
}

function renderQuoteManagerSelect(companyId, selectedManagerId = null) {
  if (!el.quoteManagerSelect) return;
  const previousValue = String(el.quoteManagerSelect.value || "").trim();
  el.quoteManagerSelect.innerHTML = "";
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
  if (!el.quoteManagerSelect.value && el.quoteManagerSelect.options.length) {
    el.quoteManagerSelect.value = el.quoteManagerSelect.options[0].value;
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
    paymentType: String(q.paymentType || "").trim(),
    discountType: normalizeDiscountType(q.discountType),
    discountValue: Math.max(0, Number(q.discountValue || 0)),
    internalNotes: String(q.internalNotes || q.notes || "").trim(),
    templateId: String(q.templateId || "").trim(),
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
    .slice()
    .sort((a, b) => String(a?.name || "").localeCompare(String(b?.name || ""), "es", { sensitivity: "base" }));
  for (const tpl of ordered) {
    const opt = document.createElement("option");
    opt.value = String(tpl.id || "");
    opt.textContent = String(tpl.name || "Plantilla");
    el.quoteTemplateSelect.appendChild(opt);
  }
  el.quoteTemplateSelect.value = selectedId && ordered.some((x) => x.id === selectedId) ? selectedId : "";
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
  renderCompaniesSelect(quoteDraft.companyId);
  renderQuoteManagerSelect(quoteDraft.companyId, quoteDraft.managerId || null);
  renderQuoteTemplateSelect(quoteDraft.templateId || "");
  fillQuoteHeaderFields(true);
  el.quoteDueDate.value = quoteDraft.dueDate || "";
  el.quotePaymentType.value = quoteDraft.paymentType || "Credito";
  el.quoteDocDate.value = quoteDraft.docDate || toISODate(new Date());
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
  quoteDraft.code = String(el.quoteCode?.value || quoteDraft.code || "").trim();
  quoteDraft.docDate = String(el.quoteDocDate?.value || quoteDraft.docDate || "").trim();
  quoteDraft.phone = String(el.quotePhone?.value || quoteDraft.phone || "").trim();
  quoteDraft.nit = String(el.quoteNIT?.value || quoteDraft.nit || "").trim();
  quoteDraft.people = String(el.quotePeople?.value || quoteDraft.people || "").trim();
  quoteDraft.eventDate = String(el.quoteEventDate?.value || quoteDraft.eventDate || "").trim();
  quoteDraft.folio = String(el.quoteFolio?.value || quoteDraft.folio || "").trim();
  quoteDraft.endDate = String(el.quoteEndDate?.value || quoteDraft.endDate || "").trim();
  quoteDraft.dueDate = String(el.quoteDueDate?.value || quoteDraft.dueDate || "").trim();
  quoteDraft.paymentType = String(el.quotePaymentType?.value || quoteDraft.paymentType || "").trim();
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
  let buffer = [];
  const flush = () => {
    if (!buffer.length) return;
    const chunk = buffer.join("\n");
    htmlParts.push(`<p class="mmReportText">${escapeHtml(chunk)}</p>`);
    buffer = [];
  };
  for (const line of lines) {
    const normalized = String(line || "").trim();
    if (/^<hr\s*\/?>$/i.test(normalized) || /^[-_]{6,}$/.test(normalized) || /^\[\[HR\]\]$/i.test(normalized)) {
      flush();
      htmlParts.push(`<hr class="mmReportHr" />`);
      continue;
    }
    buffer.push(line);
  }
  flush();
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

function buildMenuMontajeReportHtml(ev, quoteLike) {
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
        <p class="mmReportText">${escapeHtml(`${r.menuQty ? `${r.menuQty} ` : ""}${r.menuTitle || ""}`.trim() || "-")}</p>
        ${renderMenuMontajeRichText(String(r.menuDescription || ""))}
      </div>
      <div class="mmReportBlock">
        <h2 class="mmReportTitle">MONTAJE - ${escapeHtml(String(r.salon || "").toUpperCase())} - ${escapeHtml(date)}</h2>
        ${renderMenuMontajeRichText(String(r.montajeDescription || ""))}
      </div>
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
  const docTitle = `${institutionName} - MENU & MONTAJE`;
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
      background:#fff;
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
      font-size:13px;
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
      font-size:13px;
      line-height:1.35;
      color:var(--ink);
    }
    .mmReportHr{
      margin:10px 0;
      border:none;
      border-top:1px solid #9ab4d6;
    }
    @page { size: auto; margin: 10mm; }
    @media print {
      body{ background:#fff; }
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

function openMenuMontajeReportDocument(ev, quoteLike) {
  const html = buildMenuMontajeReportHtml(ev, quoteLike);
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
  openMenuMontajeReportDocument(ev, quoteDraft);
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

function getQuoteEventMeta(eventId) {
  const ev = state.events.find(x => x.id === eventId);
  if (!ev) return null;
  const series = getEventSeries(ev).slice().sort((a, b) => a.date.localeCompare(b.date));
  const startDate = series[0]?.date || ev.date;
  const endDate = series[series.length - 1]?.date || ev.date;
  return { ev, startDate, endDate };
}

function fillQuoteHeaderFields(force = false) {
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

  apply(el.quoteContact, quoteDraft.contact || manager?.name || company?.owner || "");
  apply(el.quoteEmail, quoteDraft.email || manager?.email || company?.email || "");
  apply(el.quoteBillTo, quoteDraft.billTo || company?.billTo || company?.businessName || company?.name || "");
  apply(el.quoteAddress, quoteDraft.address || company?.address || "");
  apply(el.quoteEventType, quoteDraft.eventType || company?.eventType || "");
  apply(el.quoteVenue, quoteDraft.venue || meta?.ev?.salon || "");
  apply(el.quoteSchedule, quoteDraft.schedule || `${meta?.ev?.startTime || ""} a ${meta?.ev?.endTime || ""}`.trim());
  apply(el.quoteCode, quoteDraft.code || "");
  apply(el.quoteDocDate, quoteDraft.docDate || toISODate(new Date()));
  apply(el.quotePhone, quoteDraft.phone || manager?.phone || company?.phone || "");
  apply(el.quoteNIT, quoteDraft.nit || company?.nit || "");
  apply(el.quotePeople, quoteDraft.people || "");
  apply(el.quoteEventDate, quoteDraft.eventDate || meta?.startDate || "");
  apply(el.quoteEndDate, quoteDraft.endDate || meta?.endDate || "");
  apply(el.quoteFolio, quoteDraft.folio || "");
  apply(el.quoteInternalNotes, quoteDraft.internalNotes || quoteDraft.notes || "");
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
  card.style.background = `linear-gradient(135deg, ${hexToRgba(c, 0.35)}, ${hexToRgba("#000000", 0.18)})`;
  card.style.borderColor = hexToRgba(c, 0.35);
  card.style.setProperty("--status-color", c);

  const user = state.users.find(u => u.id === ev.userId) || { name: "-" };
  const avatar = String(user.avatarDataUrl || "").trim() || avatarDataUri(user.name);
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

  card.innerHTML = `
    <div class="eventStatusChip">
      <span class="dot" style="background:${c}"></span>
      <span>${escapeHtml(ev.status)}</span>
      <small>${escapeHtml(statusShort)}</small>
    </div>
    <div class="eventInner">
      <div class="avatar"><img alt="" src="${avatar}" style="width:100%;height:100%;display:block"/></div>
      <div class="eventMeta">
        <div class="eventTitle" title="${escapeHtml(ev.name)}">${escapeHtml(ev.name)}</div>
        <div class="eventSub">
          <span class="pill">${escapeHtml(ev.startTime)}-${escapeHtml(ev.endTime)}</span>
          <span class="pill">${escapeHtml(ev.salon || "")}</span>
          <span class="pill">${escapeHtml(user.name)}</span>
          ${institutionName ? `<span class="pill" title="${escapeHtml(institutionName)}">${escapeHtml(institutionName)}</span>` : ""}
          ${paxLabel ? `<span class="pill">${escapeHtml(paxLabel)}</span>` : ""}
          ${latestQuoteTotalLabel ? `<span class="pill">${escapeHtml(latestQuoteTotalLabel)}</span>` : ""}
          ${seriesTotal > 1 ? `<span class="pill seriesPill">Reserva ${seriesIndex + 1}/${seriesTotal}</span>` : ""}
          ${followUpLabel ? `<span class="pill followupPill">${escapeHtml(followUpLabel)}</span>` : ""}
          ${reminderLabel ? `<span class="pill ${reminderClass}">${escapeHtml(reminderLabel)}</span>` : ""}
        </div>
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
  const term = String(rawTerm || "").trim().toLowerCase();
  const allRows = buildEventFinderRows();
  const rows = term
    ? allRows.filter((r) => r.searchBlob.includes(term)).slice(0, 150)
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
    if (el.templateBackdrop && !el.templateBackdrop.hidden) {
      templateUiState.pdfRenderKey = "";
      refreshTemplatePreview();
    }
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

  if (el.btnSettings && el.settingsPanel && el.settingsMenu) {
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
      el.btnQuickAddInstitution.addEventListener("click", async () => {
        closeSettingsPanel();
        await manageInstitutionsFromQuickMenu();
      });
    }

    if (el.btnQuickAddManager) {
      el.btnQuickAddManager.addEventListener("click", async () => {
        closeSettingsPanel();
        await manageManagersFromQuickMenu();
      });
    }

    if (el.btnQuickAddUser) {
      el.btnQuickAddUser.addEventListener("click", () => {
        closeSettingsPanel();
        openUserModal();
      });
    }

    if (el.btnQuickAddService) {
      el.btnQuickAddService.addEventListener("click", async () => {
        closeSettingsPanel();
        await manageServicesFromQuickMenu();
      });
    }

    if (el.btnQuickAddSalon) {
      el.btnQuickAddSalon.addEventListener("click", async () => {
        closeSettingsPanel();
        await manageSalonesFromQuickMenu();
      });
    }

    if (el.btnQuickAddGlobalGoal) {
      el.btnQuickAddGlobalGoal.addEventListener("click", async () => {
        closeSettingsPanel();
        await manageGlobalGoalsFromQuickMenu();
      });
    }

    if (el.btnQuickAddTemplate) {
      el.btnQuickAddTemplate.addEventListener("click", () => {
        openTemplateBuilderModal();
      });
    }
    if (el.btnReportSales) {
      el.btnReportSales.addEventListener("click", () => {
        closeSettingsPanel();
        openSalesReportModal();
      });
    }
    if (el.btnReportOccupancy) {
      el.btnReportOccupancy.addEventListener("click", () => {
        closeSettingsPanel();
        openOccupancyReportModal();
      });
    }
    if (el.btnReportDashboard) {
      el.btnReportDashboard.addEventListener("click", () => {
        closeSettingsPanel();
        toast("Reporte Dashboard: pendiente de construir.");
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

  if (el.btnFindEvent) {
    el.btnFindEvent.addEventListener("click", () => {
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

  el.btnClose.addEventListener("click", closeModal);
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
    addSlotRow({ salon: "", startTime: "", endTime: "" });
    syncHiddenTimesFromFirstSlot();
    updateRulesAndConflictsUI();
  });
  const onSlotChange = (e) => {
    if (!e.target.closest(".slotStart, .slotEnd, .slotRoom")) return;
    syncHiddenTimesFromFirstSlot();
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
    updateRulesAndConflictsUI();
  });
  el.eventDate.addEventListener("change", () => {
    if (!el.eventId.value && Array.isArray(pendingCreateDates) && pendingCreateDates.length > 1) {
      pendingCreateDates = [el.eventDate.value];
      el.eventDateEnd.value = el.eventDate.value;
      el.modalSubtitle.textContent = "Nuevo evento";
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
    ev.status = STATUS.CANCELADO;
    appendHistoryByKey(reservationKeyFromEvent(ev), ev.userId, "Estado cambiado a Cancelado.");
    persist();
    render();
    openModalForEdit(id);
    toast("Evento cancelado.");
  });

  el.btnMarkQuoted.addEventListener("click", () => {
    const id = el.eventId.value;
    if (!id) return;
    const ev = state.events.find(x => x.id === id);
    if (!ev) return;
    ev.status = STATUS.SEGUIMIENTO;
    appendHistoryByKey(reservationKeyFromEvent(ev), ev.userId, "Estado cambiado a Seguimiento.");
    persist();
    render();
    openModalForEdit(id);
    toast("Movido a Seguimiento.");
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

  el.btnQuoteEvent.addEventListener("click", () => {
    const id = el.eventId.value;
    if (!id) return;
    openQuoteModal(id);
  });
  if (el.btnMenuMontaje) {
    el.btnMenuMontaje.addEventListener("click", () => {
      openMenuMontajeModal();
    });
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
        renderUserSignaturePreview("");
        return;
      }
      const dataUrl = await readImageFileAsDataUrl(file);
      renderUserSignaturePreview(dataUrl);
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
      signatureDataUrl = await readImageFileAsDataUrl(signatureFile);
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

  el.btnAddCompany.addEventListener("click", openCompanyModal);
  if (el.btnOpenServiceCreate) el.btnOpenServiceCreate.addEventListener("click", openServiceModal);
  el.btnCompanyClose.addEventListener("click", closeCompanyModal);
  if (el.btnCompanyDiscard) el.btnCompanyDiscard.addEventListener("click", closeCompanyModal);
  bindSafeBackdropClose(el.companyBackdrop, closeCompanyModal);
  if (el.btnServiceClose) el.btnServiceClose.addEventListener("click", closeServiceModal);
  if (el.btnServiceDiscard) el.btnServiceDiscard.addEventListener("click", closeServiceModal);
  if (el.serviceBackdrop) {
    bindSafeBackdropClose(el.serviceBackdrop, closeServiceModal);
  }
  if (el.serviceForm) {
    el.serviceForm.addEventListener("submit", (e) => {
      e.preventDefault();
      saveServiceFromForm();
    });
  }

  if (el.btnTemplateClose) el.btnTemplateClose.addEventListener("click", closeTemplateBuilderModal);
  if (el.btnTemplateDiscard) el.btnTemplateDiscard.addEventListener("click", closeTemplateBuilderModal);
  if (el.templateBackdrop) {
    bindSafeBackdropClose(el.templateBackdrop, closeTemplateBuilderModal);
  }
  if (el.btnTemplateAddFormula) {
    el.btnTemplateAddFormula.addEventListener("click", () => {
      addTemplateFormulaRow();
      refreshTemplatePreview();
    });
  }
  if (el.templateFormulaBody) {
    el.templateFormulaBody.addEventListener("click", async (e) => {
      const removeBtn = e.target.closest(".formulaRemoveBtn");
      if (!removeBtn) return;
      const ok = await modernConfirm({
        title: "Eliminar formula",
        message: "Esta seguro de eliminar esta formula?",
        confirmText: "Si, eliminar",
        cancelText: "No",
      });
      if (!ok) return;
      removeBtn.closest(".templateFormulaRow")?.remove();
      if (!el.templateFormulaBody.querySelector(".templateFormulaRow")) {
        addTemplateFormulaRow();
      }
      refreshTemplatePreview();
    });
    el.templateFormulaBody.addEventListener("input", refreshTemplatePreview);
    el.templateFormulaBody.addEventListener("change", refreshTemplatePreview);
  }
  if (el.btnTemplateAddPosition) {
    el.btnTemplateAddPosition.addEventListener("click", () => {
      addTemplatePositionRow({ label: "Nuevo campo", token: "{{cliente.nombre}}", x: 50, y: 50 });
      refreshTemplatePreview();
    });
  }
  if (el.btnTemplateFitSignature) {
    el.btnTemplateFitSignature.addEventListener("click", async () => {
      const idx = getBestTemplateSignatureRowIndex();
      if (idx < 0) return toast("No hay un campo de firma para ajustar.");
      let ratio = 4;
      const signatureDataUrl = getBestAvailableSignatureDataUrl();
      if (signatureDataUrl) {
        const analysis = await analyzeSignatureDataUrl(signatureDataUrl);
        if (analysis?.recommendedAspectRatio) ratio = analysis.recommendedAspectRatio;
      }
      const ok = fitSignatureRowToAspect(idx, ratio);
      if (!ok) return toast("No se pudo ajustar la firma.");
      setActiveTemplatePosition(idx);
      refreshTemplatePreview();
      toast("Firma ajustada al cuadro.");
    });
  }
  if (el.templatePositionBody) {
    el.templatePositionBody.addEventListener("click", async (e) => {
      const removeBtn = e.target.closest(".posRemoveBtn");
      if (!removeBtn) return;
      const ok = await modernConfirm({
        title: "Eliminar posicion",
        message: "Esta seguro de eliminar esta posicion de plantilla?",
        confirmText: "Si, eliminar",
        cancelText: "No",
      });
      if (!ok) return;
      removeBtn.closest(".templatePositionRow")?.remove();
      if (!el.templatePositionBody.querySelector(".templatePositionRow")) {
        addTemplatePositionRow({ label: "Firma cliente", token: "{{cliente.firma}}", x: 25, y: 84 });
      }
      setActiveTemplatePosition(-1);
      refreshTemplatePreview();
    });
    el.templatePositionBody.addEventListener("click", (e) => {
      const row = e.target.closest(".templatePositionRow");
      if (!row) return;
      const rows = getTemplatePositionRows();
      const idx = rows.indexOf(row);
      setActiveTemplatePosition(idx);
    });
    el.templatePositionBody.addEventListener("input", (e) => {
      const row = e.target.closest(".templatePositionRow");
      if (row && e.target.closest(".posToken")) {
        applyTemplatePositionRowSignatureState(row);
        const rows = getTemplatePositionRows();
        const idx = rows.indexOf(row);
        if (idx >= 0) syncTemplateSignatureBounds(idx);
      }
      if (row && (e.target.closest(".posW") || e.target.closest(".posH"))) {
        const rows = getTemplatePositionRows();
        const idx = rows.indexOf(row);
        if (idx >= 0) syncTemplateSignatureBounds(idx);
      }
      refreshTemplatePreview();
    });
    el.templatePositionBody.addEventListener("change", (e) => {
      const row = e.target.closest(".templatePositionRow");
      if (row) {
        applyTemplatePositionRowSignatureState(row);
        const rows = getTemplatePositionRows();
        const idx = rows.indexOf(row);
        if (idx >= 0) syncTemplateSignatureBounds(idx);
      }
      refreshTemplatePreview();
    });
  }
  if (el.templatePageImage) {
    el.templatePageImage.addEventListener("change", async () => {
      const file = el.templatePageImage.files?.[0];
      if (!file) {
        templateAssetsDraft.pagePdf = "";
        refreshTemplatePreview();
        return;
      }
      if (!/pdf$/i.test(file.type) && !/\.pdf$/i.test(file.name || "")) {
        toast("Solo se permite PDF para la hoja base.");
        el.templatePageImage.value = "";
        return;
      }
      templateAssetsDraft.pagePdf = await readImageFileAsDataUrl(file);
      refreshTemplatePreview();
    });
  }
  if (el.templateHeaderImage) {
    el.templateHeaderImage.addEventListener("change", async () => {
      const file = el.templateHeaderImage.files?.[0];
      templateAssetsDraft.headerImage = await readImageFileAsDataUrl(file);
      refreshTemplatePreview();
    });
  }
  if (el.templateFooterImage) {
    el.templateFooterImage.addEventListener("change", async () => {
      const file = el.templateFooterImage.files?.[0];
      templateAssetsDraft.footerImage = await readImageFileAsDataUrl(file);
      refreshTemplatePreview();
    });
  }
  if (el.templateFieldPanel) {
    el.templateFieldPanel.addEventListener("click", (e) => {
      const btn = e.target.closest(".templatePresetBtn");
      if (!btn) return;
      let token = String(btn.dataset.token || "").trim();
      let label = String(btn.dataset.label || btn.textContent || "").trim();
      if (!token) return;
      if (token === "{{formula:campo}}") {
        const key = prompt("Nombre del campo formula (ej: impuesto):", "impuesto");
        const safe = String(key || "").trim().replace(/\s+/g, "_");
        if (!safe) return;
        token = `{{formula:${safe}}}`;
        label = `Formula ${safe}`;
      }
      addPresetFieldToTemplatePosition(label, token);
    });
  }
  [el.templateHeader, el.templateBody, el.templateFooter].forEach((node) => {
    if (!node) return;
    node.addEventListener("focus", () => { activeTemplateEditor = node; });
    node.addEventListener("click", () => { activeTemplateEditor = node; });
    node.addEventListener("keyup", () => { activeTemplateEditor = node; });
    node.addEventListener("input", refreshTemplatePreview);
    node.addEventListener("change", refreshTemplatePreview);
  });
  if (el.templateName) {
    el.templateName.addEventListener("input", refreshTemplatePreview);
    el.templateName.addEventListener("change", refreshTemplatePreview);
  }
  if (el.templateRoomRates) {
    el.templateRoomRates.addEventListener("input", refreshTemplatePreview);
    el.templateRoomRates.addEventListener("change", refreshTemplatePreview);
  }
  if (el.templatePagePreview) {
    el.templatePagePreview.addEventListener("mousedown", (e) => {
      const marker = e.target.closest(".templatePosBadge");
      if (!marker) return;
      const idx = Number(marker.dataset.index);
      if (!Number.isInteger(idx)) return;
      if (e.target.closest(".templateResizeHandle")) {
        beginTemplateMarkerResize(e, idx);
        return;
      }
      beginTemplateMarkerDrag(e, idx);
    });
    el.templatePagePreview.addEventListener("click", (e) => {
      const marker = e.target.closest(".templatePosBadge");
      if (!marker) return;
      const idx = Number(marker.dataset.index);
      if (!Number.isInteger(idx)) return;
      setActiveTemplatePosition(idx);
    });
  }
  window.addEventListener("mousemove", onTemplateMarkerDragMove);
  window.addEventListener("mousemove", onTemplateMarkerResizeMove);
  window.addEventListener("mouseup", endTemplateMarkerDrag);
  window.addEventListener("mouseup", endTemplateMarkerResize);
  if (el.btnTemplateLoad) el.btnTemplateLoad.addEventListener("click", loadSelectedTemplateIntoBuilder);
  if (el.btnTemplateDelete) {
    el.btnTemplateDelete.addEventListener("click", async () => {
      await deleteSelectedTemplateFromBuilder();
    });
  }
  if (el.templateSelect) {
    el.templateSelect.addEventListener("change", () => {
      activeTemplateId = String(el.templateSelect.value || "").trim();
    });
  }
  if (el.templateForm) {
    el.templateForm.addEventListener("submit", (e) => {
      e.preventDefault();
      saveTemplateFromBuilder();
    });
  }
  if (el.serviceCategory) {
    el.serviceCategory.addEventListener("change", () => {
      renderSubcategoriasServicioSelect(Number(el.serviceCategory.value));
    });
  }

  if (el.btnAppointmentClose) el.btnAppointmentClose.addEventListener("click", closeAppointmentModal);
  if (el.btnSalesReportClose) el.btnSalesReportClose.addEventListener("click", closeSalesReportModal);
  if (el.salesReportBackdrop) {
    bindSafeBackdropClose(el.salesReportBackdrop, closeSalesReportModal);
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
  if (el.btnOccupancyReportClose) el.btnOccupancyReportClose.addEventListener("click", closeOccupancyReportModal);
  if (el.occupancyReportBackdrop) {
    bindSafeBackdropClose(el.occupancyReportBackdrop, closeOccupancyReportModal);
  }
  if (el.occupancyReportWeek) {
    el.occupancyReportWeek.addEventListener("change", () => {
      const { monday, sunday } = getOccupancyWeekRange();
      occupancySelectedDayIso = toISODate(monday);
      if (el.occupancyReportSubtitle) {
        el.occupancyReportSubtitle.textContent = `Semana ${toISODate(monday)} a ${toISODate(sunday)} (Lunes a Domingo)`;
      }
      renderOccupancyReportTable();
    });
  }
  if (el.btnOccupancyReportTodayWeek) {
    el.btnOccupancyReportTodayWeek.addEventListener("click", () => {
      setOccupancyCurrentWeek();
      const { monday, sunday } = getOccupancyWeekRange();
      occupancySelectedDayIso = toISODate(monday);
      if (el.occupancyReportSubtitle) {
        el.occupancyReportSubtitle.textContent = `Semana ${toISODate(monday)} a ${toISODate(sunday)} (Lunes a Domingo)`;
      }
      renderOccupancyReportTable();
    });
  }
  if (el.btnOccupancyReportExportExcel) {
    el.btnOccupancyReportExportExcel.addEventListener("click", () => {
      exportOccupancyReportToExcel();
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
        if (areCompaniesEquivalent(state.companies[idx], company)) {
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
      enableCompany(company.id);
      companyManagersDraft.forEach((m) => enableManager(m.id));
    }
    persist();
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
  el.quoteCompany.addEventListener("change", () => {
    selectCompanyInQuote(el.quoteCompany.value);
  });
  if (el.quoteCompanySearch) {
    el.quoteCompanySearch.addEventListener("input", () => {
      refreshCompanySuggestions(el.quoteCompanySearch.value);
      // No autoseleccionar al teclear: permite borrar/escribir libremente.
    });
    el.quoteCompanySearch.addEventListener("change", () => {
      const matched = resolveCompanyFromSearch(el.quoteCompanySearch.value);
      if (matched) {
        selectCompanyInQuote(matched.id);
        return;
      }
      if (String(el.quoteCompanySearch.value || "").trim()) {
        toast("Institucion no encontrada en el catalogo.");
      }
    });
    el.quoteCompanySearch.addEventListener("keydown", (e) => {
      if (e.key !== "Enter") return;
      const matched = resolveCompanyFromSearch(el.quoteCompanySearch.value);
      if (!matched) return;
      e.preventDefault();
      selectCompanyInQuote(matched.id);
    });
  }
  el.quoteManagerSelect.addEventListener("change", () => {
    if (quoteDraft) quoteDraft.managerId = el.quoteManagerSelect.value;
    applyQuoteCompanyDefaults();
    fillQuoteHeaderFields(true);
  });
  if (el.quoteTemplateSelect) {
    el.quoteTemplateSelect.addEventListener("change", () => {
      if (!quoteDraft) return;
      quoteDraft.templateId = String(el.quoteTemplateSelect.value || "").trim();
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
      if (el.quoteDocFold.open) {
        el.quoteBackdrop.classList.add("docFloatOpen");
      } else {
        el.quoteBackdrop.classList.remove("docFloatOpen");
      }
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
    renderCompanyManagersDraft();
    toast("Encargado eliminado del borrador.");
  });

  el.quoteForm.addEventListener("submit", (e) => {
    e.preventDefault();
    saveQuoteFromForm();
  });

  el.quoteItemsBody.addEventListener("input", handleQuoteItemsInput);
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
      else if (el.salesReportBackdrop && !el.salesReportBackdrop.hidden) closeSalesReportModal();
      else if (el.occupancyReportBackdrop && !el.occupancyReportBackdrop.hidden) closeOccupancyReportModal();
      else if (!el.appointmentBackdrop.hidden) closeAppointmentModal();
      else if (el.menuMontajeBackdrop && !el.menuMontajeBackdrop.hidden) closeMenuMontajeModal();
      else if (el.templateBackdrop && !el.templateBackdrop.hidden) closeTemplateBuilderModal();
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
  el.modalTitle.textContent = "Reservar salon";
  const totalDates = Array.isArray(rangeDates) && rangeDates.length ? rangeDates.length : 1;
  el.modalSubtitle.textContent = totalDates > 1
    ? `Nuevo evento (${totalDates} dias seleccionados)`
    : "Nuevo evento";
  pendingCreateDates = totalDates > 1 ? rangeDates.slice() : null;

  el.eventId.value = "";
  el.eventName.value = "";
  el.eventDate.value = toISODate(d);
  el.eventDateEnd.value = totalDates > 1 ? rangeDates[rangeDates.length - 1] : toISODate(d);
  el.slotsBody.innerHTML = "";
  addSlotRow({ salon: "", startTime: "", endTime: "" });
  syncHiddenTimesFromFirstSlot();
  el.eventStatus.value = STATUS.PRIMERA; // default razonable
  const sessionUserId = String(authSession.userId || "").trim();
  const sessionAvailable = (state.users || []).some((u) => String(u.id) === sessionUserId && u.active !== false);
  el.eventUser.value = sessionAvailable ? sessionUserId : (state.users[0]?.id || "");
  el.eventPax.value = "";
  el.eventNotes.value = "";

  el.btnDelete.hidden = true;
  el.btnCancelEvent.hidden = true;
  el.btnQuoteEvent.hidden = true;
  el.btnQuoteEvent.textContent = "Cotizar evento";
  el.btnMarkQuoted.hidden = true;
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
  const firstDate = series[0]?.date || ev.date;
  const lastDate = series[series.length - 1]?.date || ev.date;
  el.modalSubtitle.textContent = series.length > 1
    ? `${ev.salon} - ${firstDate} a ${lastDate} - ${ev.startTime}-${ev.endTime}`
    : `${ev.salon} - ${ev.date} - ${ev.startTime}-${ev.endTime}`;

  el.eventId.value = ev.id;
  el.eventName.value = ev.name;
  el.eventDate.value = firstDate;
  el.eventDateEnd.value = lastDate;
  const slots = uniqueSlotsFromSeries(series);
  el.slotsBody.innerHTML = "";
  if (slots.length) {
    for (const slot of slots) addSlotRow(slot);
  } else {
    addSlotRow({ salon: ev.salon, startTime: ev.startTime, endTime: ev.endTime });
  }
  syncHiddenTimesFromFirstSlot();
  el.eventStatus.value = ev.status;
  el.eventUser.value = ev.userId;
  el.eventPax.value = Number(ev.pax || 0) > 0 ? String(ev.pax) : "";
  el.eventNotes.value = ev.notes || "";

  el.btnDelete.hidden = false;
  el.btnCancelEvent.hidden = (ev.status === STATUS.CANCELADO);
  el.btnQuoteEvent.hidden = false;
  el.btnQuoteEvent.textContent = ev.quote ? "Editar cotizacion" : "Cotizar evento";
  el.btnMarkQuoted.hidden = !(ev.status === STATUS.PRIMERA);
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
  // focus
  setTimeout(() => el.eventName.focus(), 0);
}
function closeModal() {
  el.modalBackdrop.hidden = true;
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
        <button type="button" class="btn" data-goal-action="edit" data-goal-month="${escapeHtml(g.month)}">Editar</button>
        <button type="button" class="btnDanger" data-goal-action="remove" data-goal-month="${escapeHtml(g.month)}">X</button>
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
}

function loadUserInModal(userId) {
  const target = (state.users || []).map(normalizeUserRecord).find((u) => u.id === userId);
  if (!target) return;
  userModalEditingId = target.id;
  if (el.userFullName) el.userFullName.value = target.fullName || "";
  if (el.userUsername) el.userUsername.value = target.username || "";
  if (el.userEmail) el.userEmail.value = target.email || "";
  if (el.userPhone) el.userPhone.value = target.phone || "";
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
  if (el.userName) el.userName.value = "";
  if (el.userFullName) el.userFullName.value = "";
  if (el.userUsername) el.userUsername.value = "";
  if (el.userEmail) el.userEmail.value = "";
  if (el.userPhone) el.userPhone.value = "";
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
}

function openQuoteModal(eventId) {
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
    companyId: state.companies?.[0]?.id || "",
    managerId: state.companies?.[0]?.managers?.[0]?.id || "",
    dueDate: ev.date,
    docDate: toISODate(new Date()),
    paymentType: "Credito",
    code: "",
    contact: "",
    email: "",
    billTo: "",
    address: "",
    eventType: "",
    venue: ev.salon,
    schedule: `${ev.startTime} a ${ev.endTime}`,
    phone: "",
    nit: "",
    people: ev.pax ? String(ev.pax) : "",
    eventDate: firstDate,
    folio: "",
    endDate: lastDate,
    internalNotes: "",
    discountType: "AMOUNT",
    discountValue: 0,
    items: [],
    notes: "",
    templateId: "",
  };
  quoteDraft.version = currentVersion;
  quoteDraft.versions = existingVersions;
  if (!quoteDraft.managerId) {
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
  fillQuoteHeaderFields(true);
  el.quoteDueDate.value = quoteDraft.dueDate || ev.date;
  el.quotePaymentType.value = quoteDraft.paymentType || "Credito";
  el.quoteDocDate.value = quoteDraft.docDate || toISODate(new Date());
  if (el.quoteDiscountType) el.quoteDiscountType.value = normalizeDiscountType(quoteDraft.discountType);
  if (el.quoteDiscountValue) el.quoteDiscountValue.value = String(Math.max(0, Number(quoteDraft.discountValue || 0)));
  renderQuoteServiceDateSelect();
  el.quoteServiceSearch.value = "";
  renderQuoteItems();
  syncPaxQuantityItems();
  renderQuoteVersionControls();

  el.quoteBackdrop.hidden = false;
  if (!existingQuote && !String(quoteDraft.code || "").trim()) {
    requestServerQuoteCode().then((serverCode) => {
      if (!quoteDraft) return;
      if (String(quoteDraft.code || "").trim()) return;
      const nextCode = String(serverCode || "").trim() || buildQuoteCode();
      quoteDraft.code = nextCode;
      if (el.quoteCode && !String(el.quoteCode.value || "").trim()) {
        el.quoteCode.value = nextCode;
      }
    });
  }
  setTimeout(() => el.quoteServiceSearch.focus(), 0);
}

function closeQuoteModal() {
  el.quoteBackdrop.hidden = true;
  el.quoteBackdrop.classList.remove("docFloatOpen");
  if (el.quoteDocFold) el.quoteDocFold.open = false;
  closeMenuMontajeModal();
  closeServiceModal();
  quoteDraft = null;
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

function openCompanyModal(companyId = "") {
  editingCompanyId = String(companyId || "").trim();
  const target = editingCompanyId
    ? (state.companies || []).find((c) => String(c.id) === editingCompanyId)
    : null;
  companyManagersDraft = target ? deepClone(target.managers || []) : [];
  renderCompanyManagersDraft();
  if (target) {
    if (el.companyTitle) el.companyTitle.textContent = "Editar empresa";
    el.companyName.value = String(target.name || "");
    el.companyOwner.value = String(target.owner || "");
    el.companyEmail.value = String(target.email || "");
    el.companyNIT.value = String(target.nit || "");
    el.companyBusinessName.value = String(target.billTo || target.businessName || target.name || "");
    el.companyEventType.value = String(target.eventType || "Social");
    el.companyAddress.value = String(target.address || "");
    el.companyPhone.value = String(target.phone || "");
    el.companyNotes.value = String(target.notes || "");
  } else {
    if (el.companyTitle) el.companyTitle.textContent = "Nueva empresa";
    el.companyName.value = "";
    el.companyOwner.value = "";
    el.companyEmail.value = "";
    el.companyNIT.value = "";
    el.companyBusinessName.value = "";
    el.companyEventType.value = "Social";
    el.companyAddress.value = "";
    el.companyPhone.value = "";
    el.companyNotes.value = "";
  }
  el.managerName.value = "";
  el.managerPhone.value = "";
  el.managerEmail.value = "";
  el.managerAddress.value = "";
  el.companyBackdrop.hidden = false;
  setTimeout(() => el.companyName.focus(), 0);
}

function closeCompanyModal() {
  el.companyBackdrop.hidden = true;
  editingCompanyId = "";
  if (el.companyTitle) el.companyTitle.textContent = "Nueva empresa";
  companyManagersDraft = [];
  el.companyName.value = "";
  el.companyOwner.value = "";
  el.companyEmail.value = "";
  el.companyNIT.value = "";
  el.companyBusinessName.value = "";
  el.companyEventType.value = "Social";
  el.companyAddress.value = "";
  el.companyPhone.value = "";
  el.companyNotes.value = "";
  el.managerName.value = "";
  el.managerPhone.value = "";
  el.managerEmail.value = "";
  el.managerAddress.value = "";
  renderCompanyManagersDraft();
}

function openServiceModal(serviceId = "") {
  if (!el.serviceBackdrop) return;
  editingServiceId = String(serviceId || "").trim();
  const target = editingServiceId
    ? (state.services || []).find((s) => String(s.id) === editingServiceId)
    : null;
  if (!catalogoCategoriasServicio.length) {
    syncServiceCatalogFromDb().catch(() => { });
  }
  renderCategoriasServicioSelect();
  el.serviceBackdrop.hidden = false;
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
  if (el.serviceTitle) el.serviceTitle.textContent = "Nuevo servicio";
  if (el.serviceForm) el.serviceForm.reset();
}

function saveServiceFromForm() {
  const name = String(el.serviceName.value || "").trim();
  const categoryIdRaw = String(el.serviceCategory.value || "").trim();
  const subcategoryIdRaw = String(el.serviceSubcategory.value || "").trim();
  const categoryId = Number(categoryIdRaw);
  const subcategoryId = Number(subcategoryIdRaw);
  const category = selectedOptionText(el.serviceCategory);
  const subcategory = selectedOptionText(el.serviceSubcategory);
  const price = Math.max(0, Number(el.servicePrice.value || 0));
  const quantityMode = String(el.serviceQuantityMode.value || "MANUAL").trim().toUpperCase() === "PAX" ? "PAX" : "MANUAL";
  const description = String(el.serviceDescription.value || "").trim();

  if (!name) return toast("Nombre de servicio es obligatorio.");
  if (!Number.isFinite(categoryId) || categoryId <= 0) return toast("Categoria es obligatoria.");
  if (!Number.isFinite(subcategoryId) || subcategoryId <= 0) return toast("Subcategoria es obligatoria.");
  if (!Number.isFinite(price) || price < 0) return toast("Precio base invalido.");

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
      if (areServicesEquivalent(state.services[idx], payload)) {
        return toast("Sin cambios detectados en servicio.");
      }
      state.services[idx] = payload;
    }
    enableService(editingId);
  } else {
    state.services.push(payload);
    enableService(payload.id);
  }

  persist();
  renderServicesList();
  closeServiceModal();
  if (el.quoteServiceSearch) el.quoteServiceSearch.value = name;
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
  companyManagersDraft.push({
    id: uid(),
    name,
    phone,
    email,
    address,
  });
  el.managerName.value = "";
  el.managerPhone.value = "";
  el.managerEmail.value = "";
  el.managerAddress.value = "";
  renderCompanyManagersDraft();
}

function renderCompanyManagersDraft() {
  if (!el.managersBody) return;
  el.managersBody.innerHTML = "";
  for (const m of companyManagersDraft) {
    const tr = document.createElement("tr");
    tr.innerHTML = `
      <td>${escapeHtml(m.name)}</td>
      <td>${escapeHtml(m.phone)}</td>
      <td>${escapeHtml(m.email)}</td>
      <td>${escapeHtml(m.address || "")}</td>
      <td><button type="button" class="btnDanger removeManagerBtn" data-manager-id="${m.id}">X</button></td>
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
  const item = service
    ? {
      rowId: uid(),
      serviceId: service.id,
      name: service.name,
      qty: service.quantityMode === "PAX" ? 1 : 1,
      unitPrice: Number(service.price || 0),
      price: service.quantityMode === "PAX"
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
      qty: 1,
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
  renderQuoteItems();
}

function renderQuoteItems() {
  if (!quoteDraft || !el.quoteItemsBody) return;
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

  for (const dateKey of orderedDates) {
    const dayItems = byDate.get(dateKey) || [];
    for (const item of dayItems) {
      const tr = document.createElement("tr");
      tr.dataset.rowId = item.rowId;
      const subtotal = Number(item.qty || 0) * Number(item.price || 0);
      const isPaxMode = String(item.quantityMode || "").toUpperCase() === "PAX";
      tr.innerHTML = `
        <td>
          <select class="quoteInput" data-field="serviceDate">
            ${availableDates.map(d => `<option value="${escapeHtml(d)}"${d === item.serviceDate ? " selected" : ""}>${escapeHtml(d)}</option>`).join("")}
          </select>
        </td>
        <td><input class="quoteInput" data-field="qty" type="number" min="0" step="1" value="${Number(item.qty || 0)}" ${isPaxMode ? "readonly disabled" : ""} title="${isPaxMode ? "Cantidad automatica por pax" : ""}" /></td>
        <td>
          <input class="quoteInput" data-field="name" list="servicesList" value="${escapeHtml(item.name || "")}" placeholder="Buscar servicio..." />
          <input class="quoteInput" data-field="description" list="serviceDescriptionsList" value="${escapeHtml(item.description || "")}" placeholder="Descripcion del servicio" style="margin-top:6px;" />
        </td>
        <td><input class="quoteInput" data-field="price" type="number" min="0" step="0.01" value="${Number(item.price || 0)}" ${isPaxMode ? "readonly disabled" : ""} title="${isPaxMode ? "Precio calculado por pax" : ""}" /></td>
        <td class="quoteMoney">Q ${subtotal.toFixed(2)}</td>
        <td><button type="button" class="btnDanger quoteRemoveBtn">X</button></td>
      `;
      el.quoteItemsBody.appendChild(tr);
    }

    const subtotalDay = dayItems.reduce((acc, x) => acc + Number(x.qty || 0) * Number(x.price || 0), 0);
    const subtotalRow = document.createElement("tr");
    subtotalRow.className = "quoteSubtotalRow";
    subtotalRow.innerHTML = `
      <td colspan="4">Subtotal ${escapeHtml(dateKey)}</td>
      <td class="quoteMoney">Q ${subtotalDay.toFixed(2)}</td>
      <td></td>
    `;
    el.quoteItemsBody.appendChild(subtotalRow);
  }

  const totals = getQuoteTotals(quoteDraft);
  if (el.quoteSubtotal) el.quoteSubtotal.textContent = `Q ${totals.subtotal.toFixed(2)}`;
  if (el.quoteDiscountAmount) el.quoteDiscountAmount.textContent = `Q ${totals.discountAmount.toFixed(2)}`;
  el.quoteTotal.textContent = `Q ${totals.total.toFixed(2)}`;
  syncQuoteServiceDateRequired();
}

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
    const term = String(input.value || "").trim();
    item.name = term;
    if (!String(item.description || "").trim()) item.description = term;
    const matched = resolveServiceFromSearch(term);
    if (matched) {
      applyServiceToQuoteItem(item, matched);
    } else {
      item.serviceId = null;
      item.quantityMode = "MANUAL";
    }
  } else if (field === "description") {
    const term = String(input.value || "").trim();
    item.description = term;
    const matched = resolveServiceFromSearch(term);
    if (matched) {
      applyServiceToQuoteItem(item, matched);
    }
  } else if (field === "serviceDate") {
    item.serviceDate = input.value;
  }
  renderQuoteItems();
}

async function handleQuoteItemsClick(e) {
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
  quoteDraft.items = quoteDraft.items.filter(x => x.rowId !== rowId);
  renderQuoteItems();
  toast("Servicio eliminado de la cotizacion.");
}

async function saveQuoteFromForm() {
  if (!quoteDraft) return;
  const eventId = el.quoteEventId.value;
  const ev = state.events.find(x => x.id === eventId);
  if (!ev) return;
  const reservationKey = reservationKeyFromEvent(ev);

  quoteDraft.companyId = el.quoteCompany.value;
  quoteDraft.managerId = el.quoteManagerSelect.value;
  quoteDraft.contact = el.quoteContact.value.trim();
  quoteDraft.email = el.quoteEmail.value.trim();
  quoteDraft.billTo = el.quoteBillTo.value.trim();
  quoteDraft.address = el.quoteAddress.value.trim();
  quoteDraft.eventType = el.quoteEventType.value.trim();
  quoteDraft.venue = el.quoteVenue.value.trim();
  quoteDraft.schedule = el.quoteSchedule.value.trim();
  quoteDraft.code = el.quoteCode.value.trim();
  quoteDraft.docDate = el.quoteDocDate.value;
  quoteDraft.phone = el.quotePhone.value.trim();
  quoteDraft.nit = el.quoteNIT.value.trim();
  quoteDraft.people = el.quotePeople.value;
  syncPaxQuantityItems();
  quoteDraft.eventDate = el.quoteEventDate.value;
  quoteDraft.folio = el.quoteFolio.value.trim();
  quoteDraft.endDate = el.quoteEndDate.value;
  quoteDraft.dueDate = el.quoteDueDate.value;
  quoteDraft.paymentType = el.quotePaymentType.value;
  quoteDraft.discountType = normalizeDiscountType(el.quoteDiscountType?.value || quoteDraft.discountType);
  quoteDraft.discountValue = Math.max(0, Number(el.quoteDiscountValue?.value || quoteDraft.discountValue || 0));
  quoteDraft.templateId = String(el.quoteTemplateSelect?.value || quoteDraft.templateId || "").trim();
  quoteDraft.internalNotes = el.quoteInternalNotes.value.trim();
  quoteDraft.notes = quoteDraft.internalNotes;
  if (!quoteDraft.code) {
    const serverCode = await requestServerQuoteCode();
    if (serverCode) {
      quoteDraft.code = serverCode;
      if (el.quoteCode) el.quoteCode.value = serverCode;
    }
  }

  if (!quoteDraft.companyId) return toast("Selecciona una empresa.");
  if (!quoteDraft.managerId) return toast("Selecciona encargado del evento.");
  if (!quoteDraft.contact || !quoteDraft.email || !quoteDraft.billTo || !quoteDraft.address) return toast("Completa contacto, email, facturar a y direccion.");
  if (!quoteDraft.code || !quoteDraft.docDate || !quoteDraft.phone || !quoteDraft.nit) return toast("Completa codigo, fecha documento, telefono y NIT.");
  if (!quoteDraft.people || Number(quoteDraft.people) <= 0) return toast("Ingresa un numero valido de personas.");
  if (!quoteDraft.eventDate || !quoteDraft.endDate) return toast("Completa fecha evento y finalizacion.");
  if (!isValidEmail(quoteDraft.email)) return toast("Correo de cotizacion invalido.");
  if (!quoteDraft.dueDate) return toast("Falta fecha maxima de pago.");
  if (!quoteDraft.items.length) return toast("Agrega al menos un servicio.");
  if (!quoteDraft.templateId) return toast("Selecciona una plantilla de contrato para generar el PDF final.");
  const selectedTemplate = (quickTemplates || []).find((t) => String(t?.id || "") === quoteDraft.templateId);
  if (!selectedTemplate) return toast("La plantilla seleccionada ya no existe.");
  if (!String(selectedTemplate?.assets?.pagePdf || "").trim()) {
    return toast("La plantilla seleccionada no tiene PDF base.");
  }

  const company = (state.companies || []).find(c => c.id === quoteDraft.companyId);
  const manager = company?.managers?.find(m => m.id === quoteDraft.managerId);
  if (!company || !manager) return toast("Empresa/encargado invalido.");
  quoteDraft.companyName = company.name;
  quoteDraft.managerName = manager.name;
  quoteDraft.managerPhone = manager.phone || "";
  quoteDraft.quotedAt = new Date().toISOString();
  const previousQuote = ev.quote ? deepClone(ev.quote) : null;
  const historyVersions = normalizeQuoteVersionHistory(previousQuote?.versions);
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

  const series = getEventSeries(ev);
  let movedToSeguimiento = false;
  for (const item of series) {
    item.quote = deepClone(savedQuote);
    if (item.status === STATUS.PRIMERA) {
      item.status = STATUS.SEGUIMIENTO;
      movedToSeguimiento = true;
    }
  }
  const totalQuote = totals.total;
  appendHistoryByKey(
    reservationKey,
    ev.userId,
    unchangedQuote
      ? `Cotizacion verificada sin cambios (V${savedQuote.version}). Total Q ${totalQuote.toFixed(2)}.`
      : (movedToSeguimiento
        ? `Cotizacion guardada. Total Q ${totalQuote.toFixed(2)}. Estado a Seguimiento.`
        : `Cotizacion guardada. Total Q ${totalQuote.toFixed(2)}. Estado conservado.`)
  );

  persist();
  let mergedOk = false;
  try {
    mergedOk = await tryOpenCombinedQuotePdf(ev, savedQuote);
  } catch (err) {
    console.error("PDF final combinado fallo:", err?.message || err);
  }
  if (!mergedOk) {
    toast("No se pudo generar el PDF final con plantilla.");
  }
  await openQuoteDocument(ev, savedQuote);
  render();
  closeQuoteModal();
  openModalForEdit(eventId);
  toast(unchangedQuote
    ? `Sin cambios detectados. Se mantiene la version V${savedQuote.version}.`
    : `Cotizacion guardada. Version V${savedQuote.version}.`);
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
    .map((r) => `${String(r?.habitacion || "").trim()}: Q ${Number(r?.precio || 0).toFixed(2)}`)
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

function buildTemplateTokenContext(ev, quote, company, manager, template) {
  const totals = getQuoteTotals(quote);
  const sellerUser = (state.users || []).map(normalizeUserRecord).find((u) => String(u.id) === String(ev?.userId || ""));
  const authUser = normalizeUserRecord(getAuthUserRecord() || {});
  const authSessionSignature = String(authSession.signatureDataUrl || "").trim();
  const authSignature = String(authUser?.signatureDataUrl || "").trim();
  const sellerSignature = String(sellerUser?.signatureDataUrl || "").trim();
  const vendorSignature = authSessionSignature || authSignature || sellerSignature || "________________";
  const vendorName = String(authSession.fullName || authUser?.fullName || authUser?.name || sellerUser?.fullName || sellerUser?.name || quote?.managerName || "").trim();
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
  let page = doc.addPage([pageSize.w, pageSize.h]);
  let y = pageSize.h - margin;

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

  const ensure = (need = 16) => {
    if (y - need > margin) return;
    page = doc.addPage([pageSize.w, pageSize.h]);
    y = pageSize.h - margin - 2;
  };

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
  const grandTotalWords = numberToWordsEs(totals.total);

  // Header band
  const headH = 28;
  drawRect(margin, y, contentW, headH, {
    fill: window.PDFLib.rgb(0.06, 0.43, 0.72),
    border: window.PDFLib.rgb(0.05, 0.33, 0.56),
    borderWidth: 1,
  });
  page.drawText(`Cotizacion ${quote?.code ? quote.code : ""}${quote?.version ? ` - V${quote.version}` : ""}`, {
    x: margin + 10,
    y: y - 18,
    size: 13,
    font: fontBold,
    color: window.PDFLib.rgb(1, 1, 1),
  });
  y -= headH + 6;

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
    drawRect(margin, rowTop, colW, cellH, { border: window.PDFLib.rgb(0.86, 0.9, 0.96), borderWidth: 0.8 });
    drawRect(margin, rowTop, keyW, cellH, { fill: window.PDFLib.rgb(0.97, 0.99, 1), border: window.PDFLib.rgb(0.86, 0.9, 0.96), borderWidth: 0.8 });
    page.drawText(String(left[0] || ""), { x: margin + 4, y: rowTop - 11, size: 7.7, font: fontBold, color: window.PDFLib.rgb(0.2, 0.29, 0.4) });
    page.drawText(String(left[1] || ""), { x: margin + keyW + 4, y: rowTop - 11, size: 7.7, font });

    // right cell
    drawRect(margin + colW, rowTop, colW, cellH, { border: window.PDFLib.rgb(0.86, 0.9, 0.96), borderWidth: 0.8 });
    drawRect(margin + colW, rowTop, keyW, cellH, { fill: window.PDFLib.rgb(0.97, 0.99, 1), border: window.PDFLib.rgb(0.86, 0.9, 0.96), borderWidth: 0.8 });
    page.drawText(String(right[0] || ""), { x: margin + colW + 4, y: rowTop - 11, size: 7.7, font: fontBold, color: window.PDFLib.rgb(0.2, 0.29, 0.4) });
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
  drawRect(margin, y, contentW, tableHeaderH, { fill: window.PDFLib.rgb(0.95, 0.97, 1), border: window.PDFLib.rgb(0.82, 0.88, 0.96), borderWidth: 1 });
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
    drawRect(margin, y, contentW, h, { border: window.PDFLib.rgb(0.86, 0.9, 0.96), borderWidth: 0.8, fill: opts.fill });
    // vertical lines
    page.drawLine({ start: { x: margin + col1, y: y }, end: { x: margin + col1, y: y - h }, thickness: 0.8, color: window.PDFLib.rgb(0.86, 0.9, 0.96) });
    page.drawLine({ start: { x: margin + col1 + col2, y: y }, end: { x: margin + col1 + col2, y: y - h }, thickness: 0.8, color: window.PDFLib.rgb(0.86, 0.9, 0.96) });
    page.drawLine({ start: { x: margin + col1 + col2 + col3, y: y }, end: { x: margin + col1 + col2 + col3, y: y - h }, thickness: 0.8, color: window.PDFLib.rgb(0.86, 0.9, 0.96) });
    page.drawText(String(qtyText || ""), { x: margin + 6, y: textY, size, font: opts.bold ? fontBold : font });
    page.drawText(String(descText || ""), { x: margin + col1 + 6, y: textY, size, font: opts.bold ? fontBold : font });
    page.drawText(String(priceText || ""), { x: margin + col1 + col2 + 6, y: textY, size, font: opts.bold ? fontBold : font });
    page.drawText(String(totalText || ""), { x: margin + col1 + col2 + col3 + 6, y: textY, size, font: opts.bold ? fontBold : font });
    y -= h;
  };

  for (const d of orderedDates) {
    const dayItems = grouped.get(d) || [];
    drawServiceRow("", `SERVICIOS DEL DIA ${formatDocDate(d)}`, "", "", {
      height: 15,
      fill: window.PDFLib.rgb(0.91, 0.96, 1),
      bold: true,
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
        `Q ${price.toFixed(2)}`,
        `Q ${lineTotal.toFixed(2)}`
      );
    }
    drawServiceRow("", "", `SUBTOTAL ${formatDocDate(d)}`, `Q ${daySubtotal.toFixed(2)}`, {
      height: 15,
      fill: window.PDFLib.rgb(0.97, 0.99, 1),
      bold: true,
      size: 7.2,
      textOffsetY: 10.4,
    });
  }

  y -= 8;
  const totalsPanelH = 86;
  ensure(totalsPanelH + 10);
  const panelX = margin;
  const panelY = y;
  const panelW = contentW;
  drawRect(panelX, panelY, panelW, totalsPanelH, {
    fill: window.PDFLib.rgb(0.985, 0.992, 1),
    border: window.PDFLib.rgb(0.75, 0.84, 0.94),
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
  const row3Y = panelY - 56;
  const money = (n) => `Q ${Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
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
    : `Descuento (Q ${discountValue.toFixed(2)})`;

  drawSummaryLine(row1Y, "Subtotal evento", money(totals.subtotal), { size: 8.2 });
  drawSummaryLine(row2Y, discountLabel, money(totals.discountAmount), {
    size: 8.2,
    color: Number(totals.discountAmount || 0) > 0 ? window.PDFLib.rgb(0.62, 0.22, 0.08) : window.PDFLib.rgb(0.12, 0.24, 0.38),
  });

  page.drawLine({
    start: { x: panelX + 10, y: panelY - 42 },
    end: { x: panelX + panelW - 10, y: panelY - 42 },
    thickness: 0.8,
    color: window.PDFLib.rgb(0.78, 0.86, 0.95),
  });

  drawRect(panelX + 10, panelY - 64, panelW - 20, 18, {
    fill: window.PDFLib.rgb(0.16, 0.47, 0.79),
    border: window.PDFLib.rgb(0.08, 0.35, 0.62),
    borderWidth: 0.9,
  });
  drawSummaryLine(row3Y, "TOTAL EVENTO", money(totals.total), {
    bold: true,
    size: 10.8,
    color: window.PDFLib.rgb(1, 1, 1),
  });

  const wordsY = panelY - 78;
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
    border: window.PDFLib.rgb(0.72, 0.8, 0.9),
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
  const totalContratado = cargoRows.reduce((acc, r) => acc + Number(r.amount || 0), 0);
  const totalAnticipos = 0;
  const saldoPendiente = Math.max(0, totalContratado - totalAnticipos);
  const showQ = (n) => (Number(n || 0) > 0 ? Number(n).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : "-");

  const summaryTitleH = 14;
  const tableHeaderH2 = 13;
  const rowH2 = 12;
  const totalRowH2 = 12;
  const spacerH2 = 5;
  const anticipoHeaderH = 12;
  const anticipoRowH = 12;
  const summaryBodyH = tableHeaderH2 + cargoRows.length * rowH2 + totalRowH2 + spacerH2 + anticipoHeaderH + anticipoRowH + anticipoRowH;
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

  const tX = margin + 70;
  const tW = contentW - 140;
  const descW = tW * 0.7;
  const qW = 20;
  const valW = tW - descW - qW;
  const drawSummaryGridRow = (yy, h, descText, valNumber, opts = {}) => {
    drawRect(tX, yy, tW, h, {
      fill: opts.fill || window.PDFLib.rgb(1, 1, 1),
      border: window.PDFLib.rgb(0.2, 0.2, 0.2),
      borderWidth: 0.8,
    });
    page.drawLine({ start: { x: tX + descW, y: yy }, end: { x: tX + descW, y: yy - h }, thickness: 0.8, color: window.PDFLib.rgb(0.2, 0.2, 0.2) });
    page.drawLine({ start: { x: tX + descW + qW, y: yy }, end: { x: tX + descW + qW, y: yy - h }, thickness: 0.8, color: window.PDFLib.rgb(0.2, 0.2, 0.2) });
    page.drawText(String(descText || ""), {
      x: tX + 3,
      y: yy - h + 3.5,
      size: 7.8,
      font: opts.bold ? fontBold : font,
      color: window.PDFLib.rgb(0.05, 0.05, 0.05),
    });
    page.drawText("Q", {
      x: tX + descW + 6,
      y: yy - h + 3.5,
      size: 7.8,
      font: opts.bold ? fontBold : font,
      color: window.PDFLib.rgb(0.05, 0.05, 0.05),
    });
    const valText = showQ(valNumber);
    const valFont = opts.bold ? fontBold : font;
    const valSize = 7.8;
    const valWidth = valFont.widthOfTextAtSize(valText, valSize);
    page.drawText(valText, {
      x: tX + tW - 4 - valWidth,
      y: yy - h + 3.5,
      size: valSize,
      font: valFont,
      color: window.PDFLib.rgb(0.05, 0.05, 0.05),
    });
  };

  drawRect(tX, y, tW, tableHeaderH2, { fill: window.PDFLib.rgb(1, 1, 1), border: window.PDFLib.rgb(0.2, 0.2, 0.2), borderWidth: 0.8 });
  page.drawLine({ start: { x: tX + descW, y }, end: { x: tX + descW, y: y - tableHeaderH2 }, thickness: 0.8, color: window.PDFLib.rgb(0.2, 0.2, 0.2) });
  page.drawLine({ start: { x: tX + descW + qW, y }, end: { x: tX + descW + qW, y: y - tableHeaderH2 }, thickness: 0.8, color: window.PDFLib.rgb(0.2, 0.2, 0.2) });
  page.drawText("DESCRIPCION", {
    x: tX + (descW / 2) - (fontBold.widthOfTextAtSize("DESCRIPCION", 7.8) / 2),
    y: y - 9,
    size: 7.8,
    font: fontBold,
    color: window.PDFLib.rgb(0.05, 0.05, 0.05),
  });
  page.drawText("TOTAL", {
    x: tX + descW + qW + (valW / 2) - (fontBold.widthOfTextAtSize("TOTAL", 7.8) / 2),
    y: y - 9,
    size: 7.8,
    font: fontBold,
    color: window.PDFLib.rgb(0.05, 0.05, 0.05),
  });
  y -= tableHeaderH2;

  for (const row of cargoRows) {
    drawSummaryGridRow(y, rowH2, row.label, row.amount);
    y -= rowH2;
  }
  drawSummaryGridRow(y, totalRowH2, "TOTAL CONTRATADO", totalContratado, { bold: true, fill: window.PDFLib.rgb(0.98, 0.98, 0.98) });
  y -= totalRowH2;

  y -= spacerH2;
  drawRect(tX, y, tW, anticipoHeaderH, { fill: window.PDFLib.rgb(1, 1, 1), border: window.PDFLib.rgb(0.2, 0.2, 0.2), borderWidth: 0.8 });
  page.drawLine({ start: { x: tX + descW, y }, end: { x: tX + descW, y: y - anticipoHeaderH }, thickness: 0.8, color: window.PDFLib.rgb(0.2, 0.2, 0.2) });
  page.drawLine({ start: { x: tX + descW + qW, y }, end: { x: tX + descW + qW, y: y - anticipoHeaderH }, thickness: 0.8, color: window.PDFLib.rgb(0.2, 0.2, 0.2) });
  page.drawText("ANTICIPOS", {
    x: tX + (descW / 2) - (fontBold.widthOfTextAtSize("ANTICIPOS", 7.8) / 2),
    y: y - 9,
    size: 7.8,
    font: fontBold,
    color: window.PDFLib.rgb(0.05, 0.05, 0.05),
  });
  page.drawText("TOTAL", {
    x: tX + descW + qW + (valW / 2) - (fontBold.widthOfTextAtSize("TOTAL", 7.8) / 2),
    y: y - 9,
    size: 7.8,
    font: fontBold,
    color: window.PDFLib.rgb(0.05, 0.05, 0.05),
  });
  y -= anticipoHeaderH;
  drawSummaryGridRow(y, anticipoRowH, "TOTAL ANTICIPOS", totalAnticipos, { bold: true, fill: window.PDFLib.rgb(0.98, 0.98, 0.98) });
  y -= anticipoRowH;
  drawSummaryGridRow(y, anticipoRowH, "SALDO", saldoPendiente, { bold: true });
  y -= anticipoRowH;

  if (quote?.internalNotes) {
    y -= 8;
    drawLine("Notas internas:", { bold: true, size: 10, lineHeight: 14 });
    for (const line of String(quote.internalNotes).split(/\r?\n/)) {
      drawLine(line);
    }
  }
  return doc;
}

async function appendTemplateToFinalPdf(finalDoc, template, ev, quote, company, manager) {
  if (!window.PDFLib) return;
  const { PDFDocument, StandardFonts } = window.PDFLib;
  const pdfData = String(template?.assets?.pagePdf || "").trim();
  if (!pdfData) return;
  const src = await PDFDocument.load(dataUrlToUint8Array(pdfData));
  const pageIndices = src.getPageIndices();
  const copied = await finalDoc.copyPages(src, pageIndices);
  for (const p of copied) finalDoc.addPage(p);

  const templatePages = finalDoc.getPages().slice(-copied.length);
  const totalHeight = templatePages.reduce((acc, p) => acc + p.getHeight(), 0);
  if (!totalHeight) return;
  const tokenCtx = buildTemplateTokenContext(ev, quote, company, manager, template);
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

async function openQuoteDocument(ev, quote) {
  if (!ev || !quote) return;
  const company = (state.companies || []).find(c => c.id === quote.companyId);
  const manager = company?.managers?.find(m => m.id === quote.managerId);
  const items = Array.isArray(quote.items) ? quote.items : [];
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
        <td colspan="4">SERVICIOS DEL DIA ${escapeHtml(formatDocDate(d))}</td>
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
          <td style="text-align:right">Q ${price.toFixed(2)}</td>
          <td style="text-align:right">Q ${lineTotal.toFixed(2)}</td>
        </tr>
      `);
    }
    rowHtml.push(`
      <tr class="daySubtotalRow">
        <td colspan="2"></td>
        <td class="sumLabel">SUBTOTAL ${escapeHtml(formatDocDate(d))}</td>
        <td class="sumValue">Q ${daySubtotal.toFixed(2)}</td>
      </tr>
    `);
  }
  const totalsDoc = getQuoteTotals(quote);
  const { catBuckets: catBucketsDoc } = aggregateQuoteBuckets(quote || {});
  const subtotalDoc = totalsDoc.subtotal;
  const discountDoc = totalsDoc.discountAmount;
  const totalDoc = totalsDoc.total;
  const discountTypeDoc = totalsDoc.discountType;
  const discountValueDoc = totalsDoc.discountValue;
  const cargoRowsDoc = [
    { label: "Alimentos y Bebidas", amount: Number(catBucketsDoc?.alimentosBebidas?.amount || 0) },
    { label: "Miscelaneos", amount: Number(catBucketsDoc?.miscelaneos?.amount || 0) },
    { label: "Hospedaje JDL", amount: Number(catBucketsDoc?.hospedajeJdl?.amount || 0) },
    { label: "Hospedaje Terceros", amount: Number(catBucketsDoc?.hospedajeTerceros?.amount || 0) },
  ];
  const totalContratadoDoc = cargoRowsDoc.reduce((acc, r) => acc + Number(r.amount || 0), 0);
  const totalAnticiposDoc = 0;
  const saldoDoc = Math.max(0, totalContratadoDoc - totalAnticiposDoc);
  const moneyQDoc = (n) => Number(n || 0) > 0
    ? Number(n || 0).toLocaleString("en-US", { minimumFractionDigits: 2, maximumFractionDigits: 2 })
    : "-";
  const grandTotalWords = numberToWordsEs(totalDoc);

  const html = `<!doctype html>
<html lang="es">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Cotizacion ${escapeHtml(quote.code || "")}</title>
  <style>
    :root{
      --ink:#0b1220;
      --muted:#334155;
      --line:#dbe3ef;
      --line-strong:#c8d4e8;
      --head:#f3f7ff;
      --brand:#0f6db8;
      --brand-soft:#e8f4ff;
    }
    *{ box-sizing:border-box; }
    body{ margin:0; padding:24px; font-family:"Segoe UI",Arial,sans-serif; color:var(--ink); background:#f5f8fc; }
    .doc{
      max-width:1080px;
      margin:0 auto;
      border:1px solid var(--line);
      border-radius:16px;
      overflow:hidden;
      background:#fff;
      box-shadow:0 14px 38px rgba(10,24,40,0.1);
    }
    .head{
      background:linear-gradient(135deg,#0f6db8,#0b548e);
      color:#fff;
      padding:14px 16px;
      border-bottom:1px solid #0a4c7f;
      font-weight:700;
      font-size:18px;
      letter-spacing:0.2px;
    }
    .grid{ display:grid; grid-template-columns:1fr 1fr; gap:0; }
    .cell{ display:grid; grid-template-columns:180px 1fr; border-right:1px solid var(--line); border-bottom:1px solid var(--line); min-height:36px; }
    .cell:nth-child(2n){ border-right:none; }
    .k{ background:#f8fbff; padding:8px 10px; color:var(--muted); font-weight:600; border-right:1px solid var(--line); }
    .v{ padding:8px 10px; font-weight:600; }
    table{ width:100%; border-collapse:collapse; }
    thead th{
      background:var(--head);
      border:1px solid var(--line);
      padding:10px 8px;
      text-align:left;
      font-size:13px;
      font-weight:700;
      color:#173a5c;
    }
    tbody td{ border:1px solid var(--line); padding:9px 8px; font-size:13px; }
    .dayHeaderRow td{
      background:linear-gradient(135deg,#d9eefc,#edf7ff);
      color:#104064;
      font-weight:800;
      letter-spacing:0.2px;
      text-transform:uppercase;
      text-align:center;
      border:1px solid var(--line-strong);
      padding:8px 10px;
    }
    .daySubtotalRow td{
      border-top:2px solid #111827;
      border-bottom:1px solid var(--line-strong);
      background:#f8fbff;
    }
    tfoot td{
      border:1px solid var(--line-strong);
      padding:10px 8px;
      font-size:13px;
      background:#f8fbff;
    }
    .sumLabel{
      text-align:right;
      font-weight:700;
      color:#153a5a;
      background:var(--brand-soft);
    }
    .sumValue{
      text-align:right;
      font-weight:800;
      color:#0d2f4d;
      background:var(--brand-soft);
    }
    .sumTotal .sumLabel,
    .sumTotal .sumValue{
      background:#dff0ff;
      color:#0b2c47;
      font-size:14px;
    }
    .sumWords{
      background:#f7fcff;
      color:#0f3555;
      font-weight:700;
      font-size:12.5px;
      text-transform:uppercase;
      line-height:1.35;
    }
    .notes{ border-top:1px solid var(--line); padding:12px; min-height:70px; }
    .notes b{ display:block; margin-bottom:6px; }
    .policyTitle{
      margin-top:12px;
      background:#1f5f75;
      color:#fff;
      font-weight:800;
      text-align:center;
      padding:6px 8px;
      border:1px solid #16485a;
      letter-spacing:.2px;
    }
    .policyBox{
      border:1px solid #9fb3cc;
      border-top:none;
      padding:8px 10px;
      font-size:12px;
      line-height:1.35;
      background:#fff;
    }
    .policyBox p{ margin:0 0 6px; }
    .policyBox p:last-child{ margin-bottom:0; }
    .cargoTable{
      width:100%;
      border-collapse:collapse;
      margin-top:0;
      font-size:12px;
    }
    .cargoTable th,.cargoTable td{
      border:1px solid #2b2b2b;
      padding:4px 6px;
    }
    .cargoTable thead th{
      background:#fff;
      text-align:center;
      color:#0f172a;
      font-size:12px;
    }
    .cargoNum{ text-align:right; white-space:nowrap; font-weight:700; }
    .cargoLabel{ font-weight:700; }
    .cargoEm{ font-weight:800; background:#f6f7f9; }
    .actions{ padding:12px; display:flex; justify-content:flex-end; gap:8px; border-top:1px solid var(--line); background:#fbfdff; }
    .actions button{
      border:1px solid #b8cde8;
      background:#fff;
      color:#103654;
      border-radius:10px;
      padding:7px 12px;
      font-weight:700;
      cursor:pointer;
    }
    .actions button:hover{ background:#eef6ff; }
    @media print{ .actions{ display:none; } body{ padding:0; } .doc{ border:none; } }
  </style>
</head>
<body>
  <div class="doc">
    <div class="head">Cotizacion ${escapeHtml(quote.code || "")}${quote.version ? ` - V${escapeHtml(String(quote.version))}` : ""}</div>
    <div class="grid">
      <div class="cell"><div class="k">Contacto</div><div class="v">${escapeHtml(quote.contact || manager?.name || "")}</div></div>
      <div class="cell"><div class="k">Codigo</div><div class="v">${escapeHtml(quote.code || "")}</div></div>
      <div class="cell"><div class="k">Encargado Evento</div><div class="v">${escapeHtml(manager?.name || "")}</div></div>
      <div class="cell"><div class="k">Fecha Documento</div><div class="v">${escapeHtml(quote.docDate || "")}</div></div>
      <div class="cell"><div class="k">Email</div><div class="v">${escapeHtml(quote.email || manager?.email || "")}</div></div>
      <div class="cell"><div class="k">Telefono</div><div class="v">${escapeHtml(quote.phone || manager?.phone || "")}</div></div>
      <div class="cell"><div class="k">Institucion</div><div class="v">${escapeHtml(quote.companyName || company?.name || "")}</div></div>
      <div class="cell"><div class="k">Forma de Pago</div><div class="v">${escapeHtml(quote.paymentType || "")}</div></div>
      <div class="cell"><div class="k">Facturar A</div><div class="v">${escapeHtml(quote.billTo || company?.billTo || company?.businessName || "")}</div></div>
      <div class="cell"><div class="k">NIT</div><div class="v">${escapeHtml(quote.nit || company?.nit || "")}</div></div>
      <div class="cell"><div class="k">Direccion</div><div class="v">${escapeHtml(quote.address || company?.address || "")}</div></div>
      <div class="cell"><div class="k">No Personas</div><div class="v">${escapeHtml(String(quote.people || ""))}</div></div>
      <div class="cell"><div class="k">Tipo Evento</div><div class="v">${escapeHtml(quote.eventType || "")}</div></div>
      <div class="cell"><div class="k">Fecha evento</div><div class="v">${escapeHtml(quote.eventDate || ev.date || "")}</div></div>
      <div class="cell"><div class="k">Salon o Jardin</div><div class="v">${escapeHtml(quote.venue || ev.salon || "")}</div></div>
      <div class="cell"><div class="k">Folio No</div><div class="v">${escapeHtml(quote.folio || "")}</div></div>
      <div class="cell"><div class="k">Horario y Evento</div><div class="v">${escapeHtml(quote.schedule || `${ev.startTime} a ${ev.endTime}`)}</div></div>
      <div class="cell"><div class="k">Fecha Finalizacion</div><div class="v">${escapeHtml(quote.endDate || ev.date || "")}</div></div>
      <div class="cell"><div class="k"></div><div class="v"></div></div>
      <div class="cell"><div class="k">Fecha Maxima Pago</div><div class="v">${escapeHtml(quote.dueDate || "")}</div></div>
    </div>

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
          <td class="sumValue">Q ${subtotalDoc.toFixed(2)}</td>
        </tr>
        <tr>
          <td colspan="2"></td>
          <td class="sumLabel">DESCUENTO ${discountTypeDoc === "PERCENT" ? `(${discountValueDoc.toFixed(2)}%)` : `(Q ${discountValueDoc.toFixed(2)})`}</td>
          <td class="sumValue">Q ${discountDoc.toFixed(2)}</td>
        </tr>
        <tr class="sumTotal">
          <td colspan="2" class="sumWords">SON: ${escapeHtml(grandTotalWords)}</td>
          <td class="sumLabel">TOTAL EVENTO</td>
          <td class="sumValue">Q ${totalDoc.toFixed(2)}</td>
        </tr>
      </tfoot>
    </table>

    <div class="policyTitle">NOTAS</div>
    <div class="policyBox">
      <p>Incrementos en menos de 24 horas de servicios y/o productos tendran cargo adicional del 10% en relacion con el excedente.</p>
      <p>El monto de anticipo para asegurar el evento depende de las clausulas de formalizacion.</p>
      <p>Estamos sujetos a pagos trimestrales. Aplicamos Clausula de No Show.</p>
      <p>Esta propuesta se formaliza al tener la firma del asesor de ventas y del cliente, junto a comprobante de anticipo y/o orden de compra.</p>
      <p>No se reembolsa el anticipo si no realiza su evento por cualquier causa.</p>
    </div>

    <div class="policyTitle">RESUMEN DE CARGOS</div>
    <table class="cargoTable">
      <thead>
        <tr>
          <th style="width:68%">DESCRIPCION</th>
          <th colspan="2">TOTAL</th>
        </tr>
      </thead>
      <tbody>
        ${cargoRowsDoc.map((r) => `
          <tr>
            <td>${escapeHtml(r.label)}</td>
            <td style="width:28px;text-align:center">Q</td>
            <td class="cargoNum">${escapeHtml(moneyQDoc(r.amount))}</td>
          </tr>
        `).join("")}
        <tr class="cargoEm">
          <td class="cargoLabel">TOTAL CONTRATADO</td>
          <td style="text-align:center">Q</td>
          <td class="cargoNum">${escapeHtml(moneyQDoc(totalContratadoDoc))}</td>
        </tr>
        <tr>
          <td class="cargoLabel">ANTICIPOS</td>
          <td style="text-align:center">Q</td>
          <td class="cargoNum">-</td>
        </tr>
        <tr class="cargoEm">
          <td class="cargoLabel">TOTAL ANTICIPOS</td>
          <td style="text-align:center">Q</td>
          <td class="cargoNum">${escapeHtml(moneyQDoc(totalAnticiposDoc))}</td>
        </tr>
        <tr class="cargoEm">
          <td class="cargoLabel">SALDO</td>
          <td style="text-align:center">Q</td>
          <td class="cargoNum">${escapeHtml(moneyQDoc(saldoDoc))}</td>
        </tr>
      </tbody>
    </table>

    <div class="notes"><b>Notas internas:</b>${escapeHtml(quote.internalNotes || quote.notes || "")}</div>
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
  const name = el.eventName.value.trim();
  const rawStatus = String(el.eventStatus.value || "").trim();
  const editingEvent = editingId ? state.events.find(x => x.id === editingId) : null;
  if (editingEvent && isEventSeriesInPast(editingEvent) && !hasPastEventEditAuthorization(editingEvent)) {
    return toast("Evento de fecha pasada bloqueado. Solicita codigo de administrador.");
  }
  const editingCurrentStatus = String(editingEvent?.status || "").trim();
  const status = isAutoStatus(rawStatus)
    ? (editingId ? (editingCurrentStatus || STATUS.PRIMERA) : STATUS.PRIMERA)
    : (rawStatus || STATUS.PRIMERA);
  const userId = el.eventUser.value;
  const notes = el.eventNotes.value.trim();
  const paxRaw = String(el.eventPax?.value || "").trim();
  const pax = paxRaw ? Math.max(1, Number(paxRaw)) : null;
  const slots = getSlotsFromForm();

  if (!name) return toast("Falta nombre del evento.");
  if (!rangeStart || !rangeEnd) return toast("Falta la fecha.");
  if (!userId) return toast("Selecciona un vendedor.");
  if (!paxRaw || Number(paxRaw) <= 0) return toast("Ingresa una cantidad valida de personas.");
  if (!slots.length) return toast("Agrega al menos un bloque de salon/horario.");
  for (const s of slots) {
    if (!s.salon || !s.startTime || !s.endTime) return toast("Completa salon, inicio y fin en cada bloque.");
    if (!isValidClockTime(s.startTime) || !isValidClockTime(s.endTime)) {
      return toast("Formato de hora invalido. Usa HH:mm.");
    }
    if (compareTime(s.endTime, s.startTime) <= 0) return toast("En cada bloque, la hora final debe ser mayor que inicio.");
  }

  const replaceEvents = editingId
    ? getEventSeries(state.events.find(x => x.id === editingId))
    : [];
  const oldKey = replaceEvents[0] ? reservationKeyFromEvent(replaceEvents[0]) : "";
  const oldSnapshot = buildSeriesSnapshot(replaceEvents);
  const replaceIds = new Set(replaceEvents.map(e => e.id));
  const currentGroupId = replaceEvents[0]?.groupId || null;
  const targetDates = listDatesBetween(rangeStart, rangeEnd);
  const needsGroup = targetDates.length > 1 || slots.length > 1 || replaceEvents.length > 1;
  const resultingGroupId = needsGroup ? (currentGroupId || `grp_${uid()}`) : null;
  const existingByKey = new Map(replaceEvents.map(e => [`${e.date}|${e.salon}|${e.startTime}|${e.endTime}`, e]));

  const drafts = [];
  for (const d of targetDates) {
    for (const s of slots) {
      const key = `${d}|${s.salon}|${s.startTime}|${s.endTime}`;
      const existing = existingByKey.get(key);
      drafts.push({
        id: existing?.id || uid(),
        name,
        salon: s.salon,
        date: d,
        groupId: resultingGroupId,
        status,
        startTime: s.startTime,
        endTime: s.endTime,
        userId,
        pax,
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
  moveHistoryKey(oldKey, newKey);
  moveReminderKey(oldKey, newKey);
  if (editingId) {
    const newSnapshot = buildDraftSnapshot(drafts);
    appendDetailedEditHistory(newKey, userId, oldSnapshot, newSnapshot);
  } else {
    appendHistoryByKey(newKey, userId, `Reserva creada: ${summarizeDraftWindow(drafts)}.`);
  }

  persist();
  autoMarkLostEvents(); // in case
  render();
  closeModal();
  toast(`Guardado: ${drafts.length} bloque(s).`);
}

function removeEvent(id, actorUserId = "") {
  const ev = state.events.find(x => x.id === id);
  if (!ev) return;
  const key = reservationKeyFromEvent(ev);
  const summary = summarizeSeriesWindow(getEventSeries(ev));
  appendHistoryByKey(key, actorUserId || ev.userId, `Reserva eliminada (${summary}).`);
  const removeIds = new Set(getEventSeries(ev).map(x => x.id));
  state.events = state.events.filter(x => !removeIds.has(x.id));
  ensureReminderStore();
  delete state.reminders[key];
  persist();
  render();
}

function getEventSeries(ev) {
  if (!ev) return [];
  if (!ev.groupId) return [ev];
  return state.events.filter(x => x.groupId === ev.groupId);
}

function uniqueSlotsFromSeries(series) {
  const map = new Map();
  for (const e of series) {
    const key = `${e.salon}|${e.startTime}|${e.endTime}`;
    if (!map.has(key)) {
      map.set(key, { salon: e.salon, startTime: e.startTime, endTime: e.endTime });
    }
  }
  return Array.from(map.values()).sort((a, b) => compareTime(a.startTime, b.startTime));
}

function summarizeSeriesWindow(series) {
  if (!series.length) return "";
  const sorted = series.slice().sort((a, b) => a.date.localeCompare(b.date));
  const first = sorted[0].date;
  const last = sorted[sorted.length - 1].date;
  const slots = uniqueSlotsFromSeries(sorted);
  const slotText = slots.map(s => `${s.salon} ${s.startTime}-${s.endTime}`).join(", ");
  return `${first}${first !== last ? ` a ${last}` : ""}${slotText ? ` (${slotText})` : ""}`;
}

function summarizeDraftWindow(drafts) {
  if (!drafts.length) return "";
  const sorted = drafts.slice().sort((a, b) => a.date.localeCompare(b.date));
  const first = sorted[0].date;
  const last = sorted[sorted.length - 1].date;
  const slots = uniqueSlotsFromSeries(sorted);
  const slotText = slots.map(s => `${s.salon} ${s.startTime}-${s.endTime}`).join(", ");
  return `${first}${first !== last ? ` a ${last}` : ""}${slotText ? ` (${slotText})` : ""}`;
}

function formatSlotsForHistory(slots) {
  if (!Array.isArray(slots) || !slots.length) return "-";
  return slots.map(s => `${s.salon} ${s.startTime}-${s.endTime}`).join(" | ");
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
    dateStart: first.date || "",
    dateEnd: last.date || "",
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
    dateStart: first.date || "",
    dateEnd: last.date || "",
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
          <label style="font-size:12px;color:#bfdbfe;">Fecha</label>
          <input id="swalReminderDate" class="swal2-input" type="date" value="${escapeHtml(nextDate)}" style="margin:0" />
          <label style="font-size:12px;color:#bfdbfe;">Hora (HH:mm)</label>
          <input id="swalReminderTime" class="swal2-input" type="text" value="${escapeHtml(nextTime)}" style="margin:0" />
          <label style="font-size:12px;color:#bfdbfe;">Medio</label>
          <select id="swalReminderChannel" class="swal2-input" style="margin:0">
            ${["Telefono", "Correo", "Teams", "Google Meet", "WhatsApp"].map((x) => `<option value="${x}"${x === nextChannel ? " selected" : ""}>${x}</option>`).join("")}
          </select>
          <label style="font-size:12px;color:#bfdbfe;">Detalle</label>
          <input id="swalReminderNotes" class="swal2-input" type="text" value="${escapeHtml(nextNotes)}" style="margin:0" />
        </div>
      `,
      showCancelButton: true,
      confirmButtonText: "Guardar cambios",
      cancelButtonText: "Cancelar",
      background: "#0b1a32",
      color: "#f8fafc",
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
    el.topbarReminderCount.textContent = String(Math.min(99, urgentCount));
  } else {
    el.topbarReminderCount.hidden = true;
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
    const cls = meta.status === "soon" ? "reminderState soon" : "reminderState";
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
    const rules = evaluateRules(draft);
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

function updateRulesAndConflictsUI() {
  applyStatusSelectTheme();
  const draft = currentDraftFromForm();
  if (!draft) return;

  // Lista de conflictos (mismo salon + traslape + misma fecha), excluyendo el propio
  const conflicts = findConflicts(draft);
  const hardBlocks = findHardBlocks(draft);

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

  const rules = evaluateRules(draft);
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
      if (opt.value === STATUS.CONFIRMADO || opt.value === STATUS.PRERESERVA) {
        opt.disabled = true;
      }
    }
    // Si el usuario ya lo tenia seleccionado, moverlo a Lista
    if (draft.status === STATUS.CONFIRMADO || draft.status === STATUS.PRERESERVA) {
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

function findConflicts(draft) {
  const id = draft.id;
  return state.events.filter(e => {
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

  const normalized = {
    ...candidate,
    salones,
    events: Array.isArray(candidate.events)
      ? candidate.events.map(e => ({ ...e, salon: e?.salon ?? "" }))
      : [],
    users: Array.isArray(candidate.users) ? candidate.users.map(normalizeUserRecord).filter((u) => u.id && u.name) : [],
    companies: Array.isArray(candidate.companies) ? candidate.companies : [],
    services: Array.isArray(candidate.services) ? candidate.services.map(normalizeServiceRecord) : [],
    quickTemplates: Array.isArray(templateSource) ? templateSource.map(normalizeTemplateRecord).filter(Boolean) : [],
    disabledCompanies: Array.isArray(candidate.disabledCompanies) ? candidate.disabledCompanies.map((x) => String(x || "").trim()).filter(Boolean) : [],
    disabledServices: Array.isArray(candidate.disabledServices) ? candidate.disabledServices.map((x) => String(x || "").trim()).filter(Boolean) : [],
    disabledManagers: Array.isArray(candidate.disabledManagers) ? candidate.disabledManagers.map((x) => String(x || "").trim()).filter(Boolean) : [],
    disabledSalones: Array.isArray(candidate.disabledSalones) ? candidate.disabledSalones.map((x) => String(x || "").trim()).filter(Boolean) : [],
    globalMonthlyGoals: Array.isArray(candidate.globalMonthlyGoals) ? candidate.globalMonthlyGoals : [],
    changeHistory: (candidate.changeHistory && typeof candidate.changeHistory === "object") ? candidate.changeHistory : {},
    reminders: (candidate.reminders && typeof candidate.reminders === "object") ? candidate.reminders : {},
  };
  normalized.companies = normalized.companies.map(normalizeCompanyRecord);
  return normalized;
}

function buildInitialState() {
  const templateSeed = loadQuickTemplates();
  return {
    salones: SALONES_DEFAULT.slice(),
    users: USERS_DEFAULT.slice(),
    companies: COMPANIES_DEFAULT.slice(),
    services: SERVICES_DEFAULT.slice(),
    quickTemplates: Array.isArray(templateSeed) ? templateSeed : [],
    disabledCompanies: [],
    disabledServices: [],
    disabledManagers: [],
    disabledSalones: [],
    globalMonthlyGoals: [],
    changeHistory: {},
    reminders: {},
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
        if (r.ok || r.status === 404) {
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
    toast("Datos cargados desde MariaDB");
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
      background: "#0b1a32",
      color: "#f8fafc",
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
      background: "#0b1a32",
      color: "#f8fafc",
      confirmButtonColor: "#2563eb",
      cancelButtonColor: "#334155",
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
      background: "#0b1a32",
      color: "#f8fafc",
      confirmButtonColor: "#8b5cf6",
      cancelButtonColor: "#334155",
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
      background: "#0b1a32",
      color: "#f8fafc",
      confirmButtonColor: "#dc2626",
      cancelButtonColor: "#334155",
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
      background: "#0b1a32",
      color: "#f8fafc",
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
  const paxRaw = String(el.eventPax?.value || "").trim();
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
  const okPax = !!paxRaw && Number(paxRaw) > 0;
  mark(el.eventPax, okPax);
  if (!okPax) issues.push("Cantidad Pax");
  if (!slots.length) issues.push("Al menos un bloque de salon y horario");

  const rows = Array.from(el.slotsBody.querySelectorAll(".slotRow"));
  for (let i = 0; i < slots.length; i++) {
    const s = slots[i];
    const idx = i + 1;
    const row = rows[i];
    const roomEl = row?.querySelector(".slotRoom");
    const startEl = row?.querySelector(".slotStart");
    const endEl = row?.querySelector(".slotEnd");
    const okRoom = !!s.salon;
    mark(roomEl, okRoom);
    if (!okRoom) issues.push(`Bloque ${idx}: Salon`);
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
  return compareTime(aStart, bEnd) < 0 && compareTime(aEnd, bStart) > 0;
}

function clamp(n, min, max) { return Math.max(min, Math.min(max, n)); }

function toast(msg) {
  const text = String(msg || "").trim();
  if (window.Swal && typeof window.Swal.fire === "function") {
    const successPattern = /\b(agregado|agregada|guardado|guardada|creado|creada|cargado|cargada|movido|movida|eliminado|eliminada|ajustada|ajustado|extendida|extendido|cambiado|cambiada|listo)\b/i;
    const errorPattern = /\b(no se pudo|error|invalido|invalida|falta|faltan|completa|sin conexion|debe|obligatoria|obligatorio|espera)\b/i;
    const icon = errorPattern.test(text) ? "error" : (successPattern.test(text) ? "success" : "info");
    window.Swal.fire({
      toast: true,
      position: "top",
      icon,
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
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function deepClone(obj) {
  return JSON.parse(JSON.stringify(obj));
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(email || "").trim());
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




