import { createClient } from "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/+esm";

const STORAGE_KEY = "luxeshade-quote-builder-v2";
const SAMPLE_PRICELIST_PATH = "./data/pricelist.csv";
const MINIMUM_SQUARE_FEET = 15;
const SUPABASE_URL = "https://pdadkkpizdsdiavziimh.supabase.co";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6InBkYWRra3BpemRzZGlhdnppaW1oIiwicm9sZSI6ImFub24iLCJpYXQiOjE3NzYyNTgwMDksImV4cCI6MjA5MTgzNDAwOX0.ycNCa_FPKV8z_XWlSxttWyK46cviIwRok6oyJCFePFo";

const supabase = createClient(SUPABASE_URL, SUPABASE_ANON_KEY);

const currencyFormatter = new Intl.NumberFormat("en-PH", {
  style: "currency",
  currency: "PHP",
});

const dateTimeFormatter = new Intl.DateTimeFormat("en-PH", {
  dateStyle: "medium",
  timeStyle: "short",
});

const contractDateFormatter = new Intl.DateTimeFormat("en-PH", {
  dateStyle: "long",
});

const PDFMAKE_TENOR_SANS_URL =
  "https://cdn.jsdelivr.net/fontsource/fonts/tenor-sans@latest/latin-400-normal.ttf";
const PDFMAKE_FONTS = {
  TenorSans: {
    normal: PDFMAKE_TENOR_SANS_URL,
    bold: PDFMAKE_TENOR_SANS_URL,
    italics: PDFMAKE_TENOR_SANS_URL,
    bolditalics: PDFMAKE_TENOR_SANS_URL,
  },
};

const MEASUREMENT_TYPE_OPTIONS = [
  "BLACKOUT CURTAIN",
  "COMBI BLINDS",
  "PREMIUM BLACKOUT ROLLER BLINDS",
  "ROLLER BLINDS",
  "ROMAN BLINDS",
  "SEMI-BLACKOUT CURTAIN",
  "SHEER CURTAIN",
  "SILHOUETTE BLINDS",
  "SOFT B/O CURTAIN",
  "SUNSCREEN",
  "VERTICAL BLINDS",
  "WOOD BLINDS",
];

/** Must match the configured material category in Material Setup (e.g. pricelist row). */
const MOTORIZED_MATERIAL_CATEGORY = "Curtains Motorized";

const QUOTE_SELECT_COLUMNS = [
  "id",
  "client_name",
  "project_name",
  "quote_date",
  "project_architect",
  "contact_number",
  "email_address",
  "quote_reference",
  "notes",
  "delivery_amount",
  "delivery_is_free",
  "install_steam_amount",
  "install_steam_is_free",
  "discount",
  "discount_type",
  "discount_value",
  "subtotal_amount",
  "applied_discount_amount",
  "final_total_amount",
  "created_at",
  "updated_at",
].join(", ");

const refs = {
  savedQuotesMenuBtn: document.querySelector("#saved-quotes-menu-btn"),
  savedQuotesDrawer: document.querySelector("#saved-quotes-drawer"),
  savedQuotesDrawerBackdrop: document.querySelector("#saved-quotes-drawer-backdrop"),
  savedQuotesDrawerCloseBtn: document.querySelector("#saved-quotes-drawer-close-btn"),
  adminMenuBtn: document.querySelector("#admin-menu-btn"),
  adminDrawer: document.querySelector("#admin-tools-drawer"),
  adminDrawerBackdrop: document.querySelector("#admin-drawer-backdrop"),
  adminDrawerCloseBtn: document.querySelector("#admin-drawer-close-btn"),
  activeQuoteBar: document.querySelector("#active-quote-bar"),
  activeQuoteTitle: document.querySelector("#active-quote-title"),
  activeQuoteDpValue: document.querySelector("#active-quote-dp-value"),
  activeQuoteDiscountLabel: document.querySelector("#active-quote-discount-label"),
  activeQuoteDiscountValue: document.querySelector("#active-quote-discount-value"),
  activeQuoteFinalTotal: document.querySelector("#active-quote-final-total"),
  activeQuoteSaveBtn: document.querySelector("#active-save-quote-btn"),
  activeExportPdfBtn: document.querySelector("#active-export-pdf-btn"),
  unloadQuoteBtn: document.querySelector("#unload-quote-btn"),
  authForm: document.querySelector("#auth-form"),
  authSession: document.querySelector("#auth-session"),
  authStatus: document.querySelector("#auth-status"),
  authUserEmail: document.querySelector("#auth-user-email"),
  signInBtn: document.querySelector("#sign-in-btn"),
  signOutBtn: document.querySelector("#sign-out-btn"),
  syncSupabaseBtn: document.querySelector("#sync-supabase-btn"),
  csvInput: document.querySelector("#csv-input"),
  csvStatus: document.querySelector("#csv-status"),
  loadSampleBtn: document.querySelector("#load-sample-btn"),
  pdfOrgSelect: document.querySelector("#pdf-org-select"),
  deliveryAmount: document.querySelector("#delivery-amount"),
  deliveryFree: document.querySelector("#delivery-free"),
  installSteamAmount: document.querySelector("#install-steam-amount"),
  installSteamFree: document.querySelector("#install-steam-free"),
  quoteClientName: document.querySelector("#quote-client-name"),
  quoteProjectName: document.querySelector("#quote-project-name"),
  quoteDate: document.querySelector("#quote-date"),
  quoteProjectArchitect: document.querySelector("#quote-project-architect"),
  quoteContactNumber: document.querySelector("#quote-contact-number"),
  quoteEmailAddress: document.querySelector("#quote-email-address"),
  quoteNotes: document.querySelector("#quote-notes"),
  newQuoteBtn: document.querySelector("#new-quote-btn"),
  saveQuoteBtn: document.querySelector("#save-quote-btn"),
  saveAsNewQuoteBtn: document.querySelector("#save-as-new-quote-btn"),
  exportPdfBtn: document.querySelector("#export-pdf-btn"),
  refreshQuotesBtn: document.querySelector("#refresh-quotes-btn"),
  currentQuoteLabel: document.querySelector("#current-quote-label"),
  quoteSaveIndicator: document.querySelector("#quote-save-indicator"),
  quoteStatus: document.querySelector("#quote-status"),
  saveReminderToast: document.querySelector("#save-reminder-toast"),
  savedQuotesCount: document.querySelector("#saved-quotes-count"),
  savedQuotesList: document.querySelector("#saved-quotes-list"),
  materialPanel: document.querySelector("#material-setup-panel"),
  measurementsPanel: document.querySelector("#measurements-panel"),
  chargesPanel: document.querySelector("#charges-panel"),
  materialBody: document.querySelector("#material-setup-body"),
  measurementBody: document.querySelector("#measurement-body"),
  addMaterialBtn: document.querySelector("#add-material-btn"),
  addMeasurementBtn: document.querySelector("#add-measurement-btn"),
  discountType: document.querySelector("#discount-type"),
  discountValue: document.querySelector("#discount-value"),
  summaryPanel: document.querySelector(".summary-panel"),
  motorQuantityDialog: document.querySelector("#motor-quantity-dialog"),
  motorQuantityForm: document.querySelector("#motor-quantity-form"),
  motorQuantityInput: document.querySelector("#motor-quantity-input"),
  motorQuantityCancel: document.querySelector("#motor-quantity-cancel"),
  deleteQuoteDialog: document.querySelector("#delete-quote-dialog"),
  deleteQuoteForm: document.querySelector("#delete-quote-form"),
  deleteQuoteCopy: document.querySelector("#delete-quote-copy"),
  deleteQuoteCancel: document.querySelector("#delete-quote-cancel"),
  deleteQuoteConfirm: document.querySelector("#delete-quote-confirm"),
};

const state = loadState();
const ACTIVE_QUOTE_BAR_REVEAL_SCROLL_Y = 300;
/** Wait this long after Add Row before running autosave (debounce reset on each Add Row). */
const AUTOSAVE_DELAY_MS = 6000;
/** Debounced Supabase save after Measurements → Add Row (see AUTOSAVE_DELAY_MS). */
const AUTOSAVE_ENABLED = false;
/** Show save reminder after this long without a manual save. */
const SAVE_REMINDER_DELAY_MS = 180000;
/** Repeat save reminder toast while quote remains unsaved. */
const SAVE_REMINDER_REPEAT_MS = 60000;
/** Pointer movement before a measurement row drag activates (mouse + touch/tablet). */
const MEASUREMENT_DRAG_ACTIVATE_PX = 8;

const runtime = {
  session: null,
  authBusy: false,
  sourceBusy: false,
  quoteBusy: false,
  quoteListBusy: false,
  quoteList: [],
  expandedQuoteId: "",
  savedQuotesDrawerOpen: false,
  adminDrawerOpen: false,
  loadedQuoteFingerprint: "",
  quoteWorkspaceActive: false,
  autosaveTimer: 0,
  exportPdfBusy: false,
  lastDraftEditAtMs: 0,
  autosaveInterval: 0,
  saveReminderTimer: 0,
  saveReminderRepeatTimer: 0,
  saveReminderActive: false,
  recentMeasurementMaterialIds: [],
};

let contractPdfAssetsPromise = null;
let pdfFontsRegistered = false;
/** Active pointer-drag session for measurement row reorder (HTML5 DnD is unreliable on touch). */
let measurementPointerDragSession = null;
let deleteQuoteDialogSession = null;

bootstrap();

async function bootstrap() {
  refs.savedQuotesMenuBtn?.addEventListener("click", () => {
    setSavedQuotesDrawerOpen(!runtime.savedQuotesDrawerOpen);
  });
  refs.savedQuotesDrawerCloseBtn?.addEventListener("click", () => {
    setSavedQuotesDrawerOpen(false);
  });
  refs.savedQuotesDrawerBackdrop?.addEventListener("click", () => {
    setSavedQuotesDrawerOpen(false);
  });
  refs.adminMenuBtn?.addEventListener("click", () => {
    setAdminDrawerOpen(!runtime.adminDrawerOpen);
  });
  refs.adminDrawerCloseBtn?.addEventListener("click", () => {
    setAdminDrawerOpen(false);
  });
  refs.adminDrawerBackdrop?.addEventListener("click", () => {
    setAdminDrawerOpen(false);
  });
  document.addEventListener("keydown", handleGlobalKeydown);
  window.addEventListener("scroll", handleWindowScroll, { passive: true });
  refs.csvInput.addEventListener("change", handleCsvUpload);
  refs.loadSampleBtn.addEventListener("click", handleLoadBundledSample);
  refs.addMaterialBtn.addEventListener("click", handleAddMaterial);
  refs.addMeasurementBtn.addEventListener("click", handleAddMeasurement);
  if (refs.pdfOrgSelect) {
    refs.pdfOrgSelect.value = state.pdfOrganization || "luxe";
    refs.pdfOrgSelect.addEventListener("change", (event) => {
      const next = String(event.target.value || "").toLowerCase();
      state.pdfOrganization = next === "nds" || next === "kk" ? next : "luxe";
      persistDraftChange();
    });
  }
  refs.discountType.addEventListener("change", (event) => {
    state.discountType = event.target.value === "percent" ? "percent" : "amount";
    persistDraftChange();
    renderSummary();
  });
  refs.discountValue.addEventListener("input", (event) => {
    state.discountValue = normalizeInputNumber(event.target.value);
    persistDraftChange();
    renderSummary();
  });

  refs.deliveryAmount?.addEventListener("input", (event) => {
    state.quoteMeta.deliveryAmount = normalizeInputNumber(event.target.value);
    persistDraftChange();
    renderSummary();
  });
  refs.deliveryFree?.addEventListener("change", (event) => {
    state.quoteMeta.deliveryIsFree = Boolean(event.target.checked);
    persistDraftChange();
    renderSummary();
  });
  refs.installSteamAmount?.addEventListener("input", (event) => {
    state.quoteMeta.installSteamAmount = normalizeInputNumber(event.target.value);
    persistDraftChange();
    renderSummary();
  });
  refs.installSteamFree?.addEventListener("change", (event) => {
    state.quoteMeta.installSteamIsFree = Boolean(event.target.checked);
    persistDraftChange();
    renderSummary();
  });
  refs.signInBtn.addEventListener("click", handleSignIn);
  refs.signOutBtn.addEventListener("click", handleSignOut);
  refs.syncSupabaseBtn.addEventListener("click", () =>
    loadMaterialsFromSupabase({ showAlertOnFailure: true }),
  );
  refs.newQuoteBtn.addEventListener("click", handleNewQuote);
  refs.saveAsNewQuoteBtn?.addEventListener("click", () => {
    void handleSaveAsNewQuote();
  });
  refs.activeQuoteSaveBtn?.addEventListener("click", () => {
    void handleSaveQuote();
  });
  refs.activeExportPdfBtn?.addEventListener("click", () => {
    void handleExportPdf();
  });
  refs.unloadQuoteBtn?.addEventListener("click", handleUnloadQuote);
  refs.saveQuoteBtn.addEventListener("click", () => {
    void handleSaveQuote();
  });
  refs.exportPdfBtn?.addEventListener("click", () => {
    void handleExportPdf();
  });
  refs.refreshQuotesBtn.addEventListener("click", () =>
    refreshSavedQuotes({ showAlertOnFailure: true }),
  );

  if (AUTOSAVE_ENABLED && !runtime.autosaveInterval) {
    runtime.autosaveInterval = window.setInterval(() => {
      if (!canAutosaveCurrentQuote()) {
        return;
      }
      // Only autosave when edits have been idle for a bit (debounce alone can be cancelled by constant typing).
      if (Date.now() - (runtime.lastDraftEditAtMs || 0) < Math.max(1500, AUTOSAVE_DELAY_MS)) {
        return;
      }
      void handleSaveQuote({ autosave: true });
    }, 15000);
  }
  refs.quoteClientName.addEventListener("input", (event) => {
    state.quoteMeta.clientName = event.target.value;
    persistDraftChange();
    renderQuoteStatus();
  });
  refs.quoteProjectName.addEventListener("input", (event) => {
    state.quoteMeta.projectName = event.target.value;
    persistDraftChange();
    renderQuoteStatus();
  });
  refs.quoteDate?.addEventListener("input", (event) => {
    state.quoteMeta.quoteDate = event.target.value;
    persistDraftChange();
  });
  refs.quoteProjectArchitect?.addEventListener("input", (event) => {
    state.quoteMeta.projectArchitect = event.target.value;
    persistDraftChange();
  });
  refs.quoteContactNumber?.addEventListener("input", (event) => {
    state.quoteMeta.contactNumber = event.target.value;
    persistDraftChange();
  });
  refs.quoteEmailAddress?.addEventListener("input", (event) => {
    state.quoteMeta.emailAddress = event.target.value;
    persistDraftChange();
  });
  refs.quoteNotes.addEventListener("input", (event) => {
    state.quoteMeta.notes = event.target.value;
    persistDraftChange();
  });

  ensureStarterRows();
  render();
  await initializeAuth();
}

function handleGlobalKeydown(event) {
  if (event.key !== "Escape") {
    return;
  }

  if (runtime.savedQuotesDrawerOpen) {
    setSavedQuotesDrawerOpen(false);
  }

  if (runtime.adminDrawerOpen) {
    setAdminDrawerOpen(false);
  }
}

function handleWindowScroll() {
  renderActiveQuoteBar();
}

function setSavedQuotesDrawerOpen(isOpen) {
  runtime.savedQuotesDrawerOpen = Boolean(isOpen);
  if (runtime.savedQuotesDrawerOpen) {
    runtime.adminDrawerOpen = false;
  }
  renderSavedQuotesDrawer();
  renderAdminDrawer();
}

function renderSavedQuotesDrawer() {
  if (
    !refs.savedQuotesDrawer ||
    !refs.savedQuotesDrawerBackdrop ||
    !refs.savedQuotesMenuBtn
  ) {
    return;
  }

  refs.savedQuotesMenuBtn.setAttribute(
    "aria-expanded",
    runtime.savedQuotesDrawerOpen ? "true" : "false",
  );
  refs.savedQuotesDrawer.setAttribute(
    "aria-hidden",
    runtime.savedQuotesDrawerOpen ? "false" : "true",
  );
  refs.savedQuotesDrawer.classList.toggle("is-open", runtime.savedQuotesDrawerOpen);
  refs.savedQuotesDrawerBackdrop.classList.toggle(
    "hidden",
    !runtime.savedQuotesDrawerOpen,
  );
  document.body.classList.toggle(
    "admin-drawer-open",
    runtime.adminDrawerOpen || runtime.savedQuotesDrawerOpen,
  );
}

function setAdminDrawerOpen(isOpen) {
  runtime.adminDrawerOpen = Boolean(isOpen);
  if (runtime.adminDrawerOpen) {
    runtime.savedQuotesDrawerOpen = false;
  }
  renderSavedQuotesDrawer();
  renderAdminDrawer();
}

function renderAdminDrawer() {
  if (!refs.adminDrawer || !refs.adminDrawerBackdrop || !refs.adminMenuBtn) {
    return;
  }

  refs.adminMenuBtn.setAttribute(
    "aria-expanded",
    runtime.adminDrawerOpen ? "true" : "false",
  );
  refs.adminDrawer.setAttribute(
    "aria-hidden",
    runtime.adminDrawerOpen ? "false" : "true",
  );
  refs.adminDrawer.classList.toggle("is-open", runtime.adminDrawerOpen);
  refs.adminDrawerBackdrop.classList.toggle("hidden", !runtime.adminDrawerOpen);
  document.body.classList.toggle(
    "admin-drawer-open",
    runtime.adminDrawerOpen || runtime.savedQuotesDrawerOpen,
  );
}

function handleUnloadQuote() {
  if (!state.quoteMeta.id) {
    return;
  }

  if (
    hasLoadedQuoteUnsavedChanges() &&
    !window.confirm("Unload this quote and discard the unsaved changes?")
  ) {
    return;
  }

  clearQueuedAutosave();
  clearSaveReminder();
  resetQuoteDraft();
  runtime.quoteWorkspaceActive = false;
  saveState();
  // Bring header menus back into view immediately.
  setSavedQuotesDrawerOpen(false);
  setAdminDrawerOpen(false);
  render();
  setQuoteStatus("Quote unloaded. Select a saved quote or start a new one.");
  window.scrollTo({ top: 0, behavior: "smooth" });
}

async function initializeAuth() {
  setAuthStatus("Checking existing session...");

  const {
    data: { session },
    error,
  } = await supabase.auth.getSession();

  if (error) {
    console.error(error);
    setAuthStatus("Could not check Supabase session.", true);
    render();
    return;
  }

  applySession(session);

  supabase.auth.onAuthStateChange((event, sessionState) => {
    applySession(sessionState);

    if (!sessionState) {
      return;
    }

    window.setTimeout(() => {
      void handleAuthSessionChange(event, sessionState);
    }, 0);
  });

  if (session) {
    // Always refetch: localStorage may hold a stale catalog count after DB rows change.
    await loadMaterialsFromSupabase({ showAlertOnFailure: false });
    await refreshSavedQuotes({ showAlertOnFailure: false, silent: true });
  }
}

async function handleAuthSessionChange(event, sessionState) {
  if (!sessionState) {
    return;
  }

  if (event === "SIGNED_OUT") {
    return;
  }

  if (!isSupabaseSourceLoaded()) {
    await loadMaterialsFromSupabase({ showAlertOnFailure: false });
  }

  await refreshSavedQuotes({ showAlertOnFailure: false, silent: true });
}

function applySession(session) {
  runtime.session = session;

  if (!session) {
    clearQueuedAutosave();
    runtime.authBusy = false;
    runtime.sourceBusy = false;
    runtime.quoteBusy = false;
    runtime.quoteListBusy = false;
    runtime.quoteList = [];
    runtime.expandedQuoteId = "";
    setAuthStatus("Not signed in yet.");
    setQuoteStatus("Sign in to save and reopen quotes from Supabase.");
  } else if (isSupabaseSourceLoaded() && state.sourceMaterials.length > 0) {
    setAuthStatus(
      `Signed in and synced ${state.sourceMaterials.length} materials from Supabase.`,
    );
  } else {
    setAuthStatus("Signed in. Admin tools are ready.");
  }

  render();
}

function ensureStarterRows() {
  if (state.selectedMaterials.length === 0) {
    state.selectedMaterials.push(createMaterialSetupRow());
  }

  if (state.measurementRows.length === 0) {
    state.measurementRows.push(createMeasurementRow());
  }
}

function createId(prefix) {
  if (window.crypto?.randomUUID) {
    return `${prefix}-${window.crypto.randomUUID()}`;
  }

  return `${prefix}-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function createMaterialSetupRow() {
  return {
    id: createId("material"),
    sourceKey: "",
    category: "",
    division: "",
    retailPrice: "",
    askingPrice: "",
  };
}

function createMeasurementRow() {
  return {
    id: createId("measurement"),
    room: "",
    type: "",
    materialCode: "",
    label: "",
    width: "",
    height: "",
    materialId: "",
    unitQuantity: "",
    isFree: false,
  };
}

function getDefaultQuoteMeta() {
  return {
    id: "",
    clientName: "",
    projectName: "",
    quoteDate: "",
    projectArchitect: "",
    contactNumber: "",
    emailAddress: "",
    quoteReference: "",
    notes: "",
    deliveryAmount: "",
    deliveryIsFree: false,
    installSteamAmount: "",
    installSteamIsFree: false,
    createdAt: "",
    updatedAt: "",
  };
}

function loadState() {
  try {
    const saved = window.localStorage.getItem(STORAGE_KEY);
    if (!saved) {
      return getDefaultState();
    }

    const parsed = JSON.parse(saved);
    return {
      sourceLabel: parsed.sourceLabel || "",
      sourceMaterials: Array.isArray(parsed.sourceMaterials)
        ? parsed.sourceMaterials
        : [],
      // Always default PDF branding to Luxe on refresh (do not persist org selection).
      pdfOrganization: "luxe",
      quoteMeta: getDefaultQuoteMeta(),
      selectedMaterials: [],
      measurementRows: [],
      discountType: "amount",
      discountValue: "",
    };
  } catch (error) {
    console.error("Failed to restore app state", error);
    return getDefaultState();
  }
}

function getDefaultState() {
  return {
    sourceLabel: "",
    sourceMaterials: [],
    pdfOrganization: "luxe",
    quoteMeta: getDefaultQuoteMeta(),
    selectedMaterials: [],
    measurementRows: [],
    discountType: "amount",
    discountValue: "",
  };
}

function saveState() {
  // Do not persist PDF branding choice; it should reset to Luxe after refresh.
  const { pdfOrganization: _pdfOrganization, ...persistedState } = state;
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(persistedState));
}

function persistDraftChange() {
  saveState();
  runtime.lastDraftEditAtMs = Date.now();
  queueAutosave();
  scheduleSaveReminder();
  renderQuoteSaveIndicator();
  renderSaveReminderButtons();
}

async function handleSignIn() {
  runtime.authBusy = true;
  renderAuth();
  setAuthStatus("Redirecting to GitHub sign-in...");

  const redirectTo = window.location.href.split("#")[0];
  const { error } = await supabase.auth.signInWithOAuth({
    provider: "github",
    options: {
      redirectTo,
    },
  });

  runtime.authBusy = false;
  renderAuth();

  if (error) {
    console.error(error);
    setAuthStatus(error.message || "GitHub sign in failed.", true);
  }
}

async function handleSignOut() {
  runtime.authBusy = true;
  renderAuth();

  const { error } = await supabase.auth.signOut();

  runtime.authBusy = false;
  renderAuth();

  if (error) {
    console.error(error);
    setAuthStatus(error.message || "Sign out failed.", true);
    return;
  }

  setAuthStatus("Signed out. Supabase sync is disabled until you sign in.");
}

async function loadMaterialsFromSupabase({ showAlertOnFailure }) {
  if (!runtime.session) {
    setAuthStatus("Sign in first before loading from Supabase.", true);
    return;
  }

  runtime.sourceBusy = true;
  renderSourceStatus();
  setAuthStatus("Loading material catalog from Supabase...");

  const { data, error } = await supabase
    .from("material_catalog")
    .select("id, division, category, retail_price")
    .order("division", { ascending: true })
    .order("category", { ascending: true });

  runtime.sourceBusy = false;

  if (error) {
    console.error(error);
    setAuthStatus(error.message || "Could not load material catalog.", true);
    renderSourceStatus();
    if (showAlertOnFailure) {
      window.alert("Supabase materials could not be loaded. Check the auth user.");
    }
    return;
  }

  state.sourceMaterials = data.map((row) => ({
    catalogId: row.id,
    key: `${row.division || ""}::${row.category}`,
    category: row.category,
    division: row.division || "",
    retailPrice: Number(row.retail_price),
  }));
  state.sourceLabel = "Supabase material_catalog";
  sanitizeSelectionsAfterSourceLoad();
  saveState();
  render();
  setAuthStatus(`Signed in and synced ${data.length} materials from Supabase.`);
}

async function refreshSavedQuotes({
  showAlertOnFailure = false,
  silent = false,
} = {}) {
  if (!runtime.session) {
    runtime.quoteListBusy = false;
    runtime.quoteList = [];
    renderQuoteWorkspace();
    return;
  }

  runtime.quoteListBusy = true;
  renderQuoteWorkspace();
  if (!silent) {
    setQuoteStatus("Loading saved quotes...");
  }

  const { data, error } = await supabase
    .from("quotes")
    .select("id, client_name, project_name, quote_reference, final_total_amount, created_at, updated_at")
    .order("updated_at", { ascending: false });

  runtime.quoteListBusy = false;

  if (error) {
    console.error(error);
    setQuoteStatus(
      error.message || "Could not load saved quotes. Run the quote security migration first.",
      true,
    );
    renderQuoteWorkspace();
    if (showAlertOnFailure) {
      window.alert("Saved quotes could not be loaded from Supabase.");
    }
    return;
  }

  runtime.quoteList = data || [];
  if (
    runtime.expandedQuoteId &&
    !runtime.quoteList.some((quote) => quote.id === runtime.expandedQuoteId)
  ) {
    runtime.expandedQuoteId = state.quoteMeta.id || "";
  }
  renderQuoteWorkspace();

  if (!silent) {
    setQuoteStatus(`Loaded ${runtime.quoteList.length} saved quote(s).`);
  }
}

async function handleCsvUpload(event) {
  const [file] = event.target.files || [];
  if (!file) {
    return;
  }

  try {
    const csvText = await file.text();
    hydrateSourceMaterials(csvText, file.name);
    refs.csvInput.value = "";
  } catch (error) {
    console.error(error);
    window.alert("The CSV file could not be read.");
  }
}

async function handleLoadBundledSample() {
  try {
    const response = await fetch(SAMPLE_PRICELIST_PATH);
    if (!response.ok) {
      throw new Error(`Unexpected response ${response.status}`);
    }

    const csvText = await response.text();
    hydrateSourceMaterials(csvText, "Bundled sample pricelist");
  } catch (error) {
    console.error(error);
    window.alert(
      "The bundled sample could not be loaded. Upload the CSV manually instead.",
    );
  }
}

function hydrateSourceMaterials(csvText, sourceLabel) {
  const rows = parseCsv(csvText);
  if (rows.length <= 1) {
    throw new Error("CSV does not contain any data rows.");
  }

  const [headerRow, ...dataRows] = rows;
  const headerIndexes = headerRow.reduce((accumulator, heading, index) => {
    accumulator[heading.trim()] = index;
    return accumulator;
  }, {});

  const categoryIndex = headerIndexes.CATEGORY;
  const divisionIndex = headerIndexes.DIVISION;
  const retailIndex = headerIndexes["FABRIC WITH ACETATE SQFT"];

  if (
    categoryIndex === undefined ||
    divisionIndex === undefined ||
    retailIndex === undefined
  ) {
    throw new Error("CSV is missing one or more required columns.");
  }

  let activeDivision = "";

  state.sourceMaterials = dataRows
    .map((row) => {
      const category = (row[categoryIndex] || "").trim();
      const divisionCell = (row[divisionIndex] || "").trim();
      const retailPrice = parseCurrencyLikeNumber(row[retailIndex]);

      if (divisionCell) {
        activeDivision = divisionCell;
      }

      if (!category || retailPrice === null) {
        return null;
      }

      return {
        catalogId: null,
        key: `${activeDivision}::${category}`,
        category,
        division: activeDivision,
        retailPrice,
      };
    })
    .filter(Boolean);

  state.sourceLabel = sourceLabel;
  sanitizeSelectionsAfterSourceLoad();
  saveState();
  render();
}

function sanitizeSelectionsAfterSourceLoad() {
  const validKeys = new Set(state.sourceMaterials.map((material) => material.key));

  state.selectedMaterials = state.selectedMaterials.map((row) => {
    if (!row.sourceKey || validKeys.has(row.sourceKey)) {
      return row;
    }

    return {
      ...row,
      sourceKey: "",
    };
  });

  sanitizeMeasurementMaterialSelections();
}

function sanitizeMeasurementMaterialSelections() {
  const validMaterialIds = new Set(
    state.selectedMaterials
      .filter((material) => material.category)
      .map((material) => material.id),
  );

  state.measurementRows = state.measurementRows.map((row) => ({
    ...row,
    materialId: validMaterialIds.has(row.materialId) ? row.materialId : "",
  }));
}

function handleAddMaterial() {
  if (state.sourceMaterials.length === 0) {
    window.alert("Load materials from Supabase or CSV first before adding rows.");
    return;
  }

  state.selectedMaterials.push(createMaterialSetupRow());
  persistDraftChange();
  renderMaterials();
  renderMeasurements();
}

function handleAddMeasurement() {
  if (getConfiguredMaterials().length === 0) {
    window.alert("Add at least one configured material before adding measurements.");
    return;
  }

  state.measurementRows.push(createMeasurementRow());
  persistDraftChange();
  queueAutosave();
  renderMeasurements();
  renderSummary();
}

function handleNewQuote() {
  if (
    hasMeaningfulDraftChanges() &&
    !window.confirm("Start a new quote and clear the current draft?")
  ) {
    return;
  }

  clearQueuedAutosave();
  setSavedQuotesDrawerOpen(false);
  runtime.quoteWorkspaceActive = true;
  resetQuoteDraft();
  saveState();
  render();
  setQuoteStatus("New quote draft started.");
}

async function handleSaveQuote(options = {}) {
  const { autosave = false } = options;

  clearQueuedAutosave();
  if (!autosave) {
    clearSaveReminder();
  }
  syncQuoteMetaFromInputs();

  if (!runtime.session) {
    setQuoteStatus(
      autosave
        ? "Autosave paused. Sign in again to keep saving changes."
        : "Sign in first before saving quotes.",
      !autosave,
    );
    return false;
  }

  const validation = validateQuoteForSave();
  if (!validation.ok) {
    setQuoteStatus(
      autosave
        ? "Unsaved changes detected. Autosave will resume after the required fields are complete."
        : validation.message,
      !autosave,
    );
    if (!autosave) {
      window.alert(validation.message);
    }
    return false;
  }

  runtime.quoteBusy = true;
  render();
  setQuoteStatus(
    autosave ? "Autosaving quote..." : state.quoteMeta.id ? "Updating quote..." : "Saving quote...",
  );
  const summaryTotals = getSummaryTotals();

  const quotePayload = {
    owner_user_id: runtime.session.user.id,
    client_name: state.quoteMeta.clientName.trim(),
    project_name: sanitizeOptionalText(state.quoteMeta.projectName),
    quote_date: sanitizeOptionalDate(state.quoteMeta.quoteDate),
    project_architect: sanitizeOptionalText(state.quoteMeta.projectArchitect),
    contact_number: sanitizeOptionalText(state.quoteMeta.contactNumber),
    email_address: sanitizeOptionalText(state.quoteMeta.emailAddress),
    quote_reference: sanitizeOptionalText(state.quoteMeta.quoteReference),
    notes: sanitizeOptionalText(state.quoteMeta.notes),
    delivery_amount: parseCurrencyLikeNumber(state.quoteMeta.deliveryAmount) || 0,
    delivery_is_free: Boolean(state.quoteMeta.deliveryIsFree),
    install_steam_amount: parseCurrencyLikeNumber(state.quoteMeta.installSteamAmount) || 0,
    install_steam_is_free: Boolean(state.quoteMeta.installSteamIsFree),
    discount: summaryTotals.discountAmount,
    discount_type: state.discountType,
    discount_value: parseCurrencyLikeNumber(state.discountValue) || 0,
    subtotal_amount: summaryTotals.subtotal,
    applied_discount_amount: summaryTotals.discountAmount,
    final_total_amount: summaryTotals.finalTotal,
  };

  const quoteQuery = state.quoteMeta.id
    ? supabase
        .from("quotes")
        .update(quotePayload)
        .eq("id", state.quoteMeta.id)
        .select(QUOTE_SELECT_COLUMNS)
        .single()
    : supabase
        .from("quotes")
        .insert(quotePayload)
        .select(QUOTE_SELECT_COLUMNS)
        .single();

  const { data: savedQuote, error: quoteError } = await quoteQuery;

  if (quoteError) {
    runtime.quoteBusy = false;
    render();
    console.error(quoteError);
    setQuoteStatus(
      quoteError.message || "Could not save the quote header. Run the quote security migration first.",
      true,
    );
    return false;
  }

  const { error: measurementDeleteError } = await supabase
    .from("quote_measurements")
    .delete()
    .eq("quote_id", savedQuote.id);

  if (measurementDeleteError) {
    runtime.quoteBusy = false;
    render();
    console.error(measurementDeleteError);
    setQuoteStatus(measurementDeleteError.message || "Could not replace saved measurements.", true);
    return false;
  }

  const { error: materialDeleteError } = await supabase
    .from("quote_materials")
    .delete()
    .eq("quote_id", savedQuote.id);

  if (materialDeleteError) {
    runtime.quoteBusy = false;
    render();
    console.error(materialDeleteError);
    setQuoteStatus(materialDeleteError.message || "Could not replace saved materials.", true);
    return false;
  }

  const materialDrafts = getMaterialDraftsForSave();
  const measurementDrafts = getMeasurementDraftsForSave();
  const materialIdMap = new Map();

  if (materialDrafts.length > 0) {
    const { data: savedMaterials, error: materialInsertError } = await supabase
      .from("quote_materials")
      .insert(
        materialDrafts.map((item, index) => ({
          quote_id: savedQuote.id,
          material_catalog_id: item.catalogId,
          division: item.division || null,
          category: item.category,
          retail_price: item.retailPrice,
          asking_price: item.askingPrice,
          sort_order: index,
        })),
      )
      .select("id, category");

    if (materialInsertError) {
      runtime.quoteBusy = false;
      render();
      console.error(materialInsertError);
      setQuoteStatus(materialInsertError.message || "Could not save quote materials.", true);
      return false;
    }

    materialDrafts.forEach((draft) => {
      const savedMaterial = savedMaterials.find((item) => item.category === draft.category);
      if (savedMaterial) {
        materialIdMap.set(draft.localMaterialId, savedMaterial.id);
      }
    });
  }

  if (measurementDrafts.length > 0) {
    const { error: measurementInsertError } = await supabase
      .from("quote_measurements")
      .insert(
        measurementDrafts.map((item, index) => ({
          quote_id: savedQuote.id,
          quote_material_id: materialIdMap.get(item.localMaterialId) || null,
          room_section: sanitizeOptionalText(item.room),
          measurement_type: sanitizeOptionalText(item.type),
          material_code: sanitizeOptionalText(item.materialCode),
          label: item.label,
          width_mm: item.width,
          height_mm: item.height,
          material_label: item.materialLabel,
          asking_price: item.askingPrice,
          sort_order: index,
          unit_quantity: item.unitQuantity,
          line_cost: item.lineCost,
          line_is_free: Boolean(item.isFree),
        })),
      );

    if (measurementInsertError) {
      runtime.quoteBusy = false;
      render();
      console.error(measurementInsertError);
      setQuoteStatus(measurementInsertError.message || "Could not save quote measurements.", true);
      return false;
    }
  }

  state.quoteMeta = {
    id: savedQuote.id,
    clientName: savedQuote.client_name || "",
    projectName: savedQuote.project_name || "",
    quoteDate: savedQuote.quote_date || "",
    projectArchitect: savedQuote.project_architect || "",
    contactNumber: savedQuote.contact_number || "",
    emailAddress: savedQuote.email_address || "",
    quoteReference: savedQuote.quote_reference || "",
    notes: savedQuote.notes || "",
    deliveryAmount: getDiscountInputDisplayValue(savedQuote.delivery_amount) || "",
    deliveryIsFree: Boolean(savedQuote.delivery_is_free),
    installSteamAmount: getDiscountInputDisplayValue(savedQuote.install_steam_amount) || "",
    installSteamIsFree: Boolean(savedQuote.install_steam_is_free),
    createdAt: savedQuote.created_at || "",
    updatedAt: savedQuote.updated_at || "",
  };
  state.discountType = savedQuote.discount_type === "percent" ? "percent" : "amount";
  state.discountValue =
    getDiscountInputDisplayValue(savedQuote.discount_value) ||
    getDiscountInputDisplayValue(savedQuote.discount);
  runtime.loadedQuoteFingerprint = buildCurrentQuoteFingerprint();
  clearSaveReminder();
  saveState();

  runtime.quoteBusy = false;
  await refreshSavedQuotes({ showAlertOnFailure: false, silent: true });
  render();
  setQuoteStatus(autosave ? "Quote autosaved to Supabase." : "Quote saved to Supabase.");
  renderQuoteSaveIndicator();
  return true;
}

async function handleSaveAsNewQuote() {
  syncQuoteMetaFromInputs();

  if (!runtime.session) {
    setQuoteStatus("Sign in first before saving quotes.", true);
    window.alert("Sign in first before saving quotes.");
    return;
  }

  if (!isQuoteWorkspaceActive() || !hasMeaningfulDraftChanges()) {
    window.alert("Nothing to save yet.");
    return;
  }

  const validation = validateQuoteForSave();
  if (!validation.ok) {
    setQuoteStatus(validation.message, true);
    window.alert(validation.message);
    return;
  }

  if (state.quoteMeta.id) {
    const ok = window.confirm(
      "Save this quote as a new copy? The original quote will remain unchanged.",
    );
    if (!ok) {
      return;
    }
  }

  state.quoteMeta = {
    ...state.quoteMeta,
    id: "",
    createdAt: "",
    updatedAt: "",
  };
  runtime.loadedQuoteFingerprint = "";
  saveState();
  render();
  void handleSaveQuote();
}

async function loadQuoteById(quoteId) {
  if (!runtime.session) {
    setQuoteStatus("Sign in first before loading saved quotes.", true);
    return;
  }

  clearQueuedAutosave();
  runtime.recentMeasurementMaterialIds = [];
  runtime.quoteBusy = true;
  render();
  setQuoteStatus("Loading quote...");

  const [quoteResult, materialsResult, measurementsResult] = await Promise.all([
    supabase
      .from("quotes")
      .select(QUOTE_SELECT_COLUMNS)
      .eq("id", quoteId)
      .single(),
    supabase
      .from("quote_materials")
      .select("id, material_catalog_id, division, category, retail_price, asking_price, sort_order")
      .eq("quote_id", quoteId)
      .order("sort_order", { ascending: true }),
    supabase
      .from("quote_measurements")
      .select(
        "id, quote_material_id, room_section, measurement_type, material_code, label, width_mm, height_mm, material_label, asking_price, unit_quantity, line_cost, line_is_free, sort_order",
      )
      .eq("quote_id", quoteId)
      .order("sort_order", { ascending: true }),
  ]);

  runtime.quoteBusy = false;

  if (quoteResult.error || materialsResult.error || measurementsResult.error) {
    console.error(quoteResult.error || materialsResult.error || measurementsResult.error);
    setQuoteStatus("Could not load the selected quote.", true);
    render();
    return;
  }

  const localIdMap = new Map();

  state.quoteMeta = {
    id: quoteResult.data.id,
    clientName: quoteResult.data.client_name || "",
    projectName: quoteResult.data.project_name || "",
    quoteDate: quoteResult.data.quote_date || "",
    projectArchitect: quoteResult.data.project_architect || "",
    contactNumber: quoteResult.data.contact_number || "",
    emailAddress: quoteResult.data.email_address || "",
    quoteReference: quoteResult.data.quote_reference || "",
    notes: quoteResult.data.notes || "",
    deliveryAmount: getDiscountInputDisplayValue(quoteResult.data.delivery_amount) || "",
    deliveryIsFree: Boolean(quoteResult.data.delivery_is_free),
    installSteamAmount: getDiscountInputDisplayValue(quoteResult.data.install_steam_amount) || "",
    installSteamIsFree: Boolean(quoteResult.data.install_steam_is_free),
    createdAt: quoteResult.data.created_at || "",
    updatedAt: quoteResult.data.updated_at || "",
  };

  state.selectedMaterials =
    materialsResult.data.map((row) => {
      const localId = createId("material");
      localIdMap.set(row.id, localId);
      return {
        id: localId,
        sourceKey: findSourceKey(row.material_catalog_id, row.division, row.category),
        category: row.category,
        division: row.division || "",
        retailPrice: row.retail_price === null ? "" : String(row.retail_price),
        askingPrice: row.asking_price === null ? "" : String(row.asking_price),
      };
    }) || [];

  state.measurementRows =
    measurementsResult.data.map((row) => {
      const unitQty =
        row.unit_quantity !== null && row.unit_quantity !== undefined
          ? String(row.unit_quantity)
          : "";
      const isQtyLine =
        row.unit_quantity !== null &&
        row.unit_quantity !== undefined &&
        Number(row.unit_quantity) > 0;

      return {
        id: createId("measurement"),
        room: row.room_section || "",
        type: row.measurement_type || "",
        materialCode: row.material_code || "",
        label: row.label || "",
        width: isQtyLine ? "" : row.width_mm === null ? "" : String(row.width_mm),
        height: isQtyLine ? "" : row.height_mm === null ? "" : String(row.height_mm),
        materialId: localIdMap.get(row.quote_material_id) || "",
        unitQuantity: unitQty,
        isFree: Boolean(row.line_is_free),
      };
    }) || [];

  state.discountType =
    quoteResult.data.discount_type === "percent" ? "percent" : "amount";
  state.discountValue =
    getDiscountInputDisplayValue(quoteResult.data.discount_value) ||
    getDiscountInputDisplayValue(quoteResult.data.discount);
  ensureStarterRows();
  runtime.expandedQuoteId = quoteId;
  runtime.loadedQuoteFingerprint = buildCurrentQuoteFingerprint();
  runtime.quoteWorkspaceActive = true;
  clearSaveReminder();
  setSavedQuotesDrawerOpen(false);
  saveState();
  render();
  setQuoteStatus(`Loaded quote for ${state.quoteMeta.clientName || "client"}.`);
}

async function deleteQuoteById(quoteId) {
  if (!runtime.session) {
    setQuoteStatus("Sign in first before deleting saved quotes.", true);
    return;
  }

  const quote = runtime.quoteList.find((item) => item.id === quoteId);
  const quoteName = quote?.client_name || "this quote";

  const ok = await confirmQuoteDeleteWithCountdown(quoteName);
  if (!ok) {
    return;
  }

  clearQueuedAutosave();
  runtime.quoteBusy = true;
  render();
  setQuoteStatus("Deleting quote...");

  const { error } = await supabase.from("quotes").delete().eq("id", quoteId);

  runtime.quoteBusy = false;

  if (error) {
    console.error(error);
    setQuoteStatus(error.message || "Could not delete the quote.", true);
    render();
    return;
  }

  if (state.quoteMeta.id === quoteId) {
    resetQuoteDraft();
    runtime.quoteWorkspaceActive = false;
  }

  if (runtime.expandedQuoteId === quoteId) {
    runtime.expandedQuoteId = state.quoteMeta.id || "";
  }

  await refreshSavedQuotes({ showAlertOnFailure: false, silent: true });
  saveState();
  render();
  setQuoteStatus("Quote deleted.");
}

function confirmQuoteDeleteWithCountdown(quoteName, seconds = 5) {
  if (!refs.deleteQuoteDialog || !refs.deleteQuoteForm || !refs.deleteQuoteConfirm) {
    return Promise.resolve(
      window.confirm(`Delete ${quoteName}? This cannot be undone.`),
    );
  }

  if (deleteQuoteDialogSession?.cleanup) {
    deleteQuoteDialogSession.cleanup();
  }

  const dialog = refs.deleteQuoteDialog;
  const confirmBtn = refs.deleteQuoteConfirm;
  const cancelBtn = refs.deleteQuoteCancel;
  const copyEl = refs.deleteQuoteCopy;

  if (copyEl) {
    copyEl.textContent = `Delete ${quoteName}? This cannot be undone.`;
  }

  let remaining = Math.max(1, Number(seconds) || 5);
  confirmBtn.disabled = true;
  confirmBtn.textContent = `DELETE (${remaining})`;

  let intervalId = window.setInterval(() => {
    remaining -= 1;
    if (remaining <= 0) {
      window.clearInterval(intervalId);
      intervalId = 0;
      confirmBtn.disabled = false;
      confirmBtn.textContent = "DELETE";
      return;
    }
    confirmBtn.textContent = `DELETE (${remaining})`;
  }, 1000);

  return new Promise((resolve) => {
    let settled = false;
    const settle = (value) => {
      if (settled) return;
      settled = true;
      resolve(value);
    };

    const cleanup = () => {
      if (intervalId) {
        window.clearInterval(intervalId);
        intervalId = 0;
      }
      dialog.removeEventListener("close", onClose);
      cancelBtn?.removeEventListener("click", onCancel);
      refs.deleteQuoteForm.removeEventListener("submit", onSubmit);
      deleteQuoteDialogSession = null;
    };

    const onCancel = () => {
      cleanup();
      dialog.close("cancel");
      settle(false);
    };

    const onSubmit = (event) => {
      event.preventDefault();
      if (confirmBtn.disabled) {
        return;
      }
      cleanup();
      dialog.close("delete");
      settle(true);
    };

    const onClose = () => {
      cleanup();
      settle(false);
    };

    deleteQuoteDialogSession = { cleanup };
    cancelBtn?.addEventListener("click", onCancel);
    refs.deleteQuoteForm.addEventListener("submit", onSubmit);
    dialog.addEventListener("close", onClose, { once: true });
    dialog.showModal();
  });
}

function render() {
  renderSavedQuotesDrawer();
  renderAdminDrawer();
  renderAuth();
  renderSourceStatus();
  renderQuoteWorkspace();
  renderQuoteDetailPanels();
  renderMaterials();
  renderMeasurements();
  renderSummary();
}

function renderActiveQuoteBar() {
  if (!refs.activeQuoteBar || !refs.activeQuoteTitle || !refs.unloadQuoteBtn) {
    return;
  }

  const hasLoadedQuote = Boolean(state.quoteMeta.id);
  const shouldShow =
    hasLoadedQuote && window.scrollY > ACTIVE_QUOTE_BAR_REVEAL_SCROLL_Y;
  const { discountAmount, finalTotal, half } = getSummaryTotals();
  const discountInputValue = parseCurrencyLikeNumber(state.discountValue);
  const discountLabel = state.discountType === "percent"
    ? `${discountInputValue ?? 0}% Discount`
    : `${formatCurrency(discountAmount)} Discount`;
  refs.activeQuoteBar.classList.toggle("hidden", !shouldShow);
  refs.activeQuoteTitle.textContent = hasLoadedQuote
    ? buildActiveQuoteTitle()
    : "-";
  if (refs.activeQuoteDpValue) {
    refs.activeQuoteDpValue.textContent = formatCurrency(half);
  }
  if (refs.activeQuoteDiscountValue) {
    refs.activeQuoteDiscountValue.textContent = formatCurrency(discountAmount);
  }
  if (refs.activeQuoteDiscountLabel) {
    refs.activeQuoteDiscountLabel.textContent = discountLabel;
  }
  if (refs.activeQuoteFinalTotal) {
    refs.activeQuoteFinalTotal.textContent = formatCurrency(finalTotal);
  }
  if (refs.activeQuoteSaveBtn) {
    refs.activeQuoteSaveBtn.disabled = !runtime.session || runtime.quoteBusy;
    refs.activeQuoteSaveBtn.textContent = runtime.quoteBusy ? "Saving..." : "Save Quote";
  }
  renderSaveReminderButtons();
  if (refs.activeExportPdfBtn) {
    refs.activeExportPdfBtn.disabled =
      runtime.quoteBusy || runtime.exportPdfBusy || !isQuoteWorkspaceActive();
    refs.activeExportPdfBtn.textContent =
      runtime.exportPdfBusy ? "Preparing PDF..." : "Export PDF";
  }
  refs.unloadQuoteBtn.disabled = runtime.quoteBusy;
}

function renderQuoteDetailPanels() {
  const shouldShow = isQuoteWorkspaceActive();

  refs.materialPanel?.classList.toggle("hidden", !shouldShow);
  refs.measurementsPanel?.classList.toggle("hidden", !shouldShow);
  refs.chargesPanel?.classList.toggle("hidden", !shouldShow);
  refs.summaryPanel?.classList.toggle("hidden", !shouldShow);
}

function renderAuth() {
  const signedIn = Boolean(runtime.session);

  refs.authForm.classList.toggle("hidden", signedIn);
  refs.authSession.classList.toggle("hidden", !signedIn);
  refs.authUserEmail.textContent = runtime.session?.user?.email || "-";

  refs.signInBtn.disabled = runtime.authBusy;
  refs.signOutBtn.disabled = runtime.authBusy;
  refs.signInBtn.textContent = runtime.authBusy
    ? "Redirecting..."
    : "Continue with GitHub";
}

function renderSourceStatus() {
  const count = state.sourceMaterials.length;
  refs.csvStatus.textContent =
    count > 0
      ? `${state.sourceLabel} loaded with ${count} material options.`
      : "No material source loaded yet.";

  refs.syncSupabaseBtn.disabled = !runtime.session || runtime.sourceBusy;
  refs.addMaterialBtn.disabled = count === 0 || runtime.quoteBusy;
  refs.addMeasurementBtn.disabled =
    getConfiguredMaterials().length === 0 || runtime.quoteBusy;
}

function renderQuoteWorkspace() {
  refs.quoteClientName.value = state.quoteMeta.clientName;
  refs.quoteProjectName.value = state.quoteMeta.projectName;
  refs.quoteDate.value = state.quoteMeta.quoteDate;
  refs.quoteProjectArchitect.value = state.quoteMeta.projectArchitect;
  refs.quoteContactNumber.value = state.quoteMeta.contactNumber;
  refs.quoteEmailAddress.value = state.quoteMeta.emailAddress;
  refs.quoteNotes.value = state.quoteMeta.notes;

  if (refs.deliveryAmount) {
    refs.deliveryAmount.value = state.quoteMeta.deliveryAmount || "";
    refs.deliveryAmount.disabled = runtime.quoteBusy;
  }
  if (refs.deliveryFree) {
    refs.deliveryFree.checked = Boolean(state.quoteMeta.deliveryIsFree);
    refs.deliveryFree.disabled = runtime.quoteBusy;
  }
  if (refs.installSteamAmount) {
    refs.installSteamAmount.value = state.quoteMeta.installSteamAmount || "";
    refs.installSteamAmount.disabled = runtime.quoteBusy;
  }
  if (refs.installSteamFree) {
    refs.installSteamFree.checked = Boolean(state.quoteMeta.installSteamIsFree);
    refs.installSteamFree.disabled = runtime.quoteBusy;
  }

  const signedIn = Boolean(runtime.session);
  refs.quoteClientName.disabled = runtime.quoteBusy;
  refs.quoteProjectName.disabled = runtime.quoteBusy;
  refs.quoteDate.disabled = runtime.quoteBusy;
  refs.quoteProjectArchitect.disabled = runtime.quoteBusy;
  refs.quoteContactNumber.disabled = runtime.quoteBusy;
  refs.quoteEmailAddress.disabled = runtime.quoteBusy;
  refs.quoteNotes.disabled = runtime.quoteBusy;
  refs.newQuoteBtn.disabled = runtime.quoteBusy;
  refs.saveQuoteBtn.disabled = !signedIn || runtime.quoteBusy;
  renderSaveReminderButtons();
  if (refs.saveAsNewQuoteBtn) {
    refs.saveAsNewQuoteBtn.disabled =
      !signedIn || runtime.quoteBusy || !isQuoteWorkspaceActive();
  }
  if (refs.exportPdfBtn) {
    refs.exportPdfBtn.disabled = runtime.quoteBusy || runtime.exportPdfBusy || !isQuoteWorkspaceActive();
    refs.exportPdfBtn.textContent = runtime.exportPdfBusy ? "Preparing PDF..." : "Export PDF";
  }
  refs.refreshQuotesBtn.disabled =
    !signedIn || runtime.quoteBusy || runtime.quoteListBusy;

  renderQuoteStatus();
  renderQuoteSaveIndicator();
  renderSavedQuotesList();
  renderActiveQuoteBar();
}

function renderQuoteStatus() {
  refs.currentQuoteLabel.textContent = getCurrentQuoteLabel();

  if (!refs.quoteStatus.classList.contains("is-error")) {
    if (state.quoteMeta.id) {
      refs.quoteStatus.textContent = state.quoteMeta.updatedAt
        ? `Editing a saved quote. Last updated ${formatDateTime(state.quoteMeta.updatedAt)}.`
        : "Editing a saved quote.";
    } else if (hasMeaningfulDraftChanges()) {
      refs.quoteStatus.textContent =
        "Working on a new draft. Save it when the client quote is ready.";
    } else {
      refs.quoteStatus.textContent =
        "No quote loaded yet. Select a saved quote or start a new one.";
    }
  }
}

function renderSavedQuotesList() {
  refs.savedQuotesCount.textContent = `${runtime.quoteList.length} quote${runtime.quoteList.length === 1 ? "" : "s"}`;
  refs.savedQuotesList.innerHTML = "";

  if (!runtime.session) {
    refs.savedQuotesList.append(
      buildEmptyBlock("Sign in to load and manage saved quotes."),
    );
    return;
  }

  if (runtime.quoteList.length === 0) {
    refs.savedQuotesList.append(buildEmptyBlock("No saved quotes yet."));
    return;
  }

  if (!runtime.expandedQuoteId && state.quoteMeta.id) {
    runtime.expandedQuoteId = state.quoteMeta.id;
  }

  runtime.quoteList.forEach((quote) => {
    const isExpanded = quote.id === runtime.expandedQuoteId;
    const card = document.createElement("article");
    card.className = "saved-quote-card";
    card.dataset.quoteId = quote.id;
    if (quote.id === state.quoteMeta.id) {
      card.classList.add("is-active");
    }
    if (isExpanded) {
      card.classList.add("is-expanded");
    }

    const header = document.createElement("button");
    header.type = "button";
    header.className = "saved-quote-header";
    header.disabled = runtime.quoteBusy || runtime.quoteListBusy;
    header.setAttribute(
      "aria-expanded",
      isExpanded ? "true" : "false",
    );
    header.addEventListener("click", () => {
      const willExpand = runtime.expandedQuoteId !== quote.id;
      runtime.expandedQuoteId = willExpand ? quote.id : "";
      renderSavedQuotesList();
      if (willExpand) {
        queueSavedQuoteScrollIntoView(quote.id);
      }
    });

    const title = document.createElement("strong");
    title.className = "saved-quote-title";
    title.textContent =
      sanitizeOptionalText(quote.project_name) ||
      sanitizeOptionalText(quote.client_name) ||
      "Untitled quote";

    const chevron = document.createElement("span");
    chevron.className = "saved-quote-chevron";
    chevron.textContent = isExpanded ? "Hide" : "Show";

    const actions = document.createElement("div");
    actions.className = "saved-quote-actions";

    const loadButton = document.createElement("button");
    loadButton.type = "button";
    loadButton.className = "secondary-button";
    loadButton.textContent = quote.id === state.quoteMeta.id ? "Loaded" : "Open";
    loadButton.disabled =
      runtime.quoteBusy || runtime.quoteListBusy || quote.id === state.quoteMeta.id;
    loadButton.addEventListener("click", () => {
      loadQuoteById(quote.id);
    });

    const deleteButton = document.createElement("button");
    deleteButton.type = "button";
    deleteButton.className = "ghost-danger-button";
    deleteButton.textContent = "Delete";
    deleteButton.disabled = runtime.quoteBusy || runtime.quoteListBusy;
    deleteButton.addEventListener("click", () => {
      deleteQuoteById(quote.id);
    });

    actions.append(loadButton, deleteButton);
    header.append(title, chevron);

    const details = document.createElement("div");
    details.className = "saved-quote-details";
    details.hidden = !isExpanded;

    const meta = document.createElement("div");
    meta.className = "saved-quote-meta";

    const client = document.createElement("p");
    client.className = "saved-quote-detail-line";
    client.textContent = `Client: ${sanitizeOptionalText(quote.client_name) || "Unnamed client"}`;

    const project = document.createElement("p");
    project.className = "saved-quote-detail-line";
    project.textContent = `Project Address: ${sanitizeOptionalText(quote.project_name) || "No project yet."}`;

    const total = document.createElement("p");
    total.className = "saved-quote-detail-line saved-quote-detail-total";
    total.textContent = `Final Total: ${formatCurrency(quote.final_total_amount || 0)}`;

    const updated = document.createElement("p");
    updated.className = "saved-quote-detail-line";
    updated.textContent = `Updated ${formatDateTime(quote.updated_at || quote.created_at)}`;

    meta.append(client, project, total, updated);
    details.append(meta, actions);
    card.append(header, details);
    refs.savedQuotesList.append(card);
  });
}

function queueSavedQuoteScrollIntoView(quoteId) {
  if (!quoteId || !refs.savedQuotesList) {
    return;
  }

  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      scrollSavedQuoteIntoView(quoteId);
      window.setTimeout(() => {
        scrollSavedQuoteIntoView(quoteId);
      }, 180);
    });
  });
}

function scrollSavedQuoteIntoView(quoteId) {
  if (!quoteId || !refs.savedQuotesList) {
    return;
  }

  const container = refs.savedQuotesList;
  const card = container.querySelector(`[data-quote-id="${quoteId}"]`);

  if (!card) {
    return;
  }

  const scrollPadding = 12;
  const containerRect = container.getBoundingClientRect();
  const cardRect = card.getBoundingClientRect();
  const cardTop = cardRect.top - containerRect.top + container.scrollTop;
  const cardBottom = cardRect.bottom - containerRect.top + container.scrollTop;
  const cardHeight = cardBottom - cardTop;
  const viewTop = container.scrollTop;
  const viewBottom = viewTop + container.clientHeight;
  const availableHeight = Math.max(container.clientHeight - scrollPadding * 2, 0);

  let nextScrollTop = viewTop;

  if (cardHeight > availableHeight) {
    if (cardTop - scrollPadding < viewTop || cardBottom + scrollPadding > viewBottom) {
      nextScrollTop = Math.max(cardTop - scrollPadding, 0);
    }
  } else if (cardTop - scrollPadding < viewTop) {
    nextScrollTop = Math.max(cardTop - scrollPadding, 0);
  } else if (cardBottom + scrollPadding > viewBottom) {
    nextScrollTop = Math.max(
      cardBottom - container.clientHeight + scrollPadding,
      0,
    );
  }

  if (Math.abs(nextScrollTop - viewTop) > 1) {
    container.scrollTo({
      top: nextScrollTop,
      behavior: "smooth",
    });
  }
}

function getSourceMaterialOptionLabel(material) {
  if (!material) {
    return "";
  }
  return material.division
    ? `${material.category} (${material.division})`
    : String(material.category || "");
}

function sourceMaterialMatchesSearch(material, query) {
  const q = query.trim().toLowerCase();
  if (!q) {
    return true;
  }
  const cat = String(material.category || "").toLowerCase();
  const div = String(material.division || "").toLowerCase();
  const full = getSourceMaterialOptionLabel(material).toLowerCase();
  return cat.includes(q) || div.includes(q) || full.includes(q);
}

function pushRecentMeasurementMaterial(materialId) {
  const id = String(materialId || "");
  if (!id) {
    return;
  }
  runtime.recentMeasurementMaterialIds = [
    id,
    ...(runtime.recentMeasurementMaterialIds || []).filter((existing) => existing !== id),
  ].slice(0, 10);
}

function renderMaterials() {
  refs.materialBody.innerHTML = "";

  if (state.selectedMaterials.length === 0) {
    refs.materialBody.append(createEmptyStateRow(5, "Add a material to begin."));
    return;
  }

  const selectedKeys = state.selectedMaterials
    .map((r) => r.sourceKey)
    .filter(Boolean);

  state.selectedMaterials.forEach((row) => {
    const tr = document.createElement("tr");

    const materialCell = document.createElement("td");
    materialCell.className = "material-combobox-cell";

    const wrap = document.createElement("div");
    wrap.className = "material-combobox";

    const input = document.createElement("input");
    input.type = "text";
    input.className = "material-combobox-input";
    input.autocomplete = "off";
    input.spellcheck = false;
    input.placeholder = "Search material…";
    input.setAttribute("role", "combobox");
    input.setAttribute("aria-autocomplete", "list");
    input.setAttribute("aria-expanded", "false");

    const currentMaterial = row.sourceKey
      ? state.sourceMaterials.find((m) => m.key === row.sourceKey)
      : null;
    input.value = currentMaterial ? getSourceMaterialOptionLabel(currentMaterial) : "";

    const list = document.createElement("ul");
    list.className = "material-combobox-list";
    list.setAttribute("role", "listbox");
    list.hidden = true;

    let tableShellScrollEl = null;
    let outsidePointerActive = false;

    const catalogEmpty = state.sourceMaterials.length === 0;
    input.disabled = runtime.quoteBusy || catalogEmpty;
    if (catalogEmpty) {
      input.placeholder = "Load catalog first";
    }

    const isKeyTakenElsewhere = (key) =>
      selectedKeys.includes(key) && key !== row.sourceKey;

    const positionDropdown = () => {
      if (list.hidden) {
        return;
      }
      const rect = input.getBoundingClientRect();
      const margin = 6;
      const spaceBelow = window.innerHeight - rect.bottom - margin - 8;
      const maxH = Math.min(280, Math.max(120, spaceBelow));
      list.style.position = "fixed";
      list.style.left = `${rect.left}px`;
      list.style.top = `${rect.bottom + margin}px`;
      list.style.width = `${Math.max(rect.width, 200)}px`;
      list.style.maxHeight = `${maxH}px`;
      list.style.zIndex = "3000";
    };

    const syncInputLabelFromRow = () => {
      const m = row.sourceKey
        ? state.sourceMaterials.find((item) => item.key === row.sourceKey)
        : null;
      if (m) {
        input.value = getSourceMaterialOptionLabel(m);
      }
    };

    const isPointerInsideCombobox = (event) => {
      const x = event.clientX;
      const y = event.clientY;
      if (Number.isFinite(x) && Number.isFinite(y)) {
        const over = document.elementFromPoint(x, y);
        if (
          over &&
          (over === input ||
            over === list ||
            over === wrap ||
            wrap.contains(over))
        ) {
          return true;
        }
        // Overlay scrollbars: elementFromPoint often misses the <ul>; use geometry.
        const listRect = list.getBoundingClientRect();
        if (
          x >= listRect.left &&
          x <= listRect.right &&
          y >= listRect.top &&
          y <= listRect.bottom
        ) {
          return true;
        }
        const inputRect = input.getBoundingClientRect();
        if (
          x >= inputRect.left &&
          x <= inputRect.right &&
          y >= inputRect.top &&
          y <= inputRect.bottom
        ) {
          return true;
        }
      }
      if (event.target instanceof Node && wrap.contains(event.target)) {
        return true;
      }
      if (typeof event.composedPath === "function") {
        for (const node of event.composedPath()) {
          if (node === wrap || node === list || node === input) {
            return true;
          }
          if (node instanceof HTMLElement && wrap.contains(node)) {
            return true;
          }
        }
      }
      return false;
    };

    const onWindowScrollCapture = (event) => {
      if (list.hidden) {
        return;
      }
      const t = event.target;
      if (t === list) {
        return;
      }
      if (t instanceof Node && list.contains(t)) {
        return;
      }
      closeList();
    };

    const onOutsidePointerDown = (event) => {
      if (list.hidden) {
        return;
      }
      if (isPointerInsideCombobox(event)) {
        return;
      }
      syncInputLabelFromRow();
      closeList();
    };

    const registerOutsideCloser = () => {
      if (outsidePointerActive) {
        return;
      }
      document.addEventListener("pointerdown", onOutsidePointerDown, true);
      outsidePointerActive = true;
    };

    const unregisterOutsideCloser = () => {
      if (!outsidePointerActive) {
        return;
      }
      document.removeEventListener("pointerdown", onOutsidePointerDown, true);
      outsidePointerActive = false;
    };

    const onTableShellScroll = () => {
      if (!list.hidden) {
        // Reposition instead of closing: wheel/trackpad often scroll-chains from
        // the dropdown list into this ancestor; closing there felt like "scroll
        // closes the menu".
        positionDropdown();
      }
    };

    const closeList = () => {
      list.hidden = true;
      input.setAttribute("aria-expanded", "false");
      unregisterOutsideCloser();
      window.removeEventListener("scroll", onWindowScrollCapture, true);
      window.removeEventListener("resize", positionDropdown);
      tableShellScrollEl?.removeEventListener("scroll", onTableShellScroll);
      tableShellScrollEl = null;
    };

    const openList = () => {
      if (catalogEmpty || runtime.quoteBusy) {
        return;
      }
      list.hidden = false;
      input.setAttribute("aria-expanded", "true");
      renderMaterialListOptions();
      positionDropdown();
      registerOutsideCloser();
      window.addEventListener("scroll", onWindowScrollCapture, true);
      window.addEventListener("resize", positionDropdown);
      tableShellScrollEl = document.querySelector(
        "#material-setup-panel .table-shell",
      );
      tableShellScrollEl?.addEventListener("scroll", onTableShellScroll, {
        passive: true,
      });
    };

    function renderMaterialListOptions() {
      list.innerHTML = "";
      const filtered = state.sourceMaterials.filter((m) =>
        sourceMaterialMatchesSearch(m, input.value),
      );

      if (filtered.length === 0) {
        const empty = document.createElement("li");
        empty.className = "material-combobox-empty";
        empty.textContent = "No materials match your search.";
        list.append(empty);
        return;
      }

      filtered.forEach((material) => {
        const li = document.createElement("li");
        li.className = "material-combobox-option";
        li.setAttribute("role", "option");
        const taken = isKeyTakenElsewhere(material.key);
        if (taken) {
          li.classList.add("is-disabled");
        }

        const main = document.createElement("div");
        main.className = "material-combobox-option-main";
        main.textContent = material.category || "";

        const sub = document.createElement("div");
        sub.className = "material-combobox-option-sub";
        if (material.division) {
          sub.textContent = material.division;
        } else {
          sub.hidden = true;
        }

        li.append(main, sub);

        if (!taken) {
          li.addEventListener("mousedown", (event) => {
            event.preventDefault();
            row.sourceKey = material.key;
            row.category = material.category || "";
            row.division = material.division || "";
            row.retailPrice = material ? material.retailPrice : "";
            input.value = getSourceMaterialOptionLabel(material);
            closeList();
            sanitizeMeasurementMaterialSelections();
            persistDraftChange();
            render();
          });
        }

        list.append(li);
      });
    }

    input.addEventListener("focus", () => {
      if (catalogEmpty || runtime.quoteBusy) {
        return;
      }
      if (row.sourceKey) {
        input.select();
      }
      openList();
    });

    input.addEventListener("input", () => {
      if (catalogEmpty || runtime.quoteBusy) {
        return;
      }
      renderMaterialListOptions();
      if (list.hidden) {
        openList();
      } else {
        positionDropdown();
      }
    });

    input.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        syncInputLabelFromRow();
        closeList();
        input.blur();
      }
    });

    input.addEventListener("focusout", (event) => {
      if (list.hidden) {
        return;
      }
      const next = event.relatedTarget;
      if (next === null) {
        return;
      }
      if (wrap.contains(next)) {
        return;
      }
      syncInputLabelFromRow();
      closeList();
    });

    wrap.append(input, list);
    materialCell.append(wrap);

    const retailCell = document.createElement("td");
    retailCell.className = "money-cell";
    retailCell.textContent =
      row.retailPrice === "" ? "PHP 0.00" : formatCurrency(row.retailPrice);

    const askingCell = document.createElement("td");
    const askingInput = document.createElement("input");
    askingInput.type = "number";
    askingInput.min = "0";
    askingInput.step = "0.01";
    askingInput.placeholder = "0.00";
    askingInput.disabled = !row.category || runtime.quoteBusy;

    const pocketCell = document.createElement("td");
    pocketCell.className = "money-cell";
    pocketCell.textContent = formatCurrency(getPocketValue(row) ?? 0);

    askingInput.value = row.askingPrice;
    askingInput.addEventListener("input", (event) => {
      row.askingPrice = normalizeInputNumber(event.target.value);
      pocketCell.textContent = formatCurrency(getPocketValue(row) ?? 0);
      persistDraftChange();
      renderSummary();
    });
    askingInput.addEventListener("change", () => {
      renderMeasurements();
    });
    askingCell.append(askingInput);

    const actionCell = document.createElement("td");
    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "ghost-danger-button";
    removeButton.textContent = "Remove";
    removeButton.disabled = runtime.quoteBusy;
    removeButton.addEventListener("click", () => {
      const isUsed = state.measurementRows.some(
        (measurement) => measurement.materialId === row.id,
      );

      if (
        isUsed &&
        !window.confirm(
          "This material is used in the measurement table. Remove it and clear those selections?",
        )
      ) {
        return;
      }

      state.selectedMaterials = state.selectedMaterials.filter(
        (material) => material.id !== row.id,
      );

      if (state.selectedMaterials.length === 0) {
        state.selectedMaterials.push(createMaterialSetupRow());
      }

      sanitizeMeasurementMaterialSelections();
      persistDraftChange();
      render();
    });
    actionCell.append(removeButton);

    tr.append(materialCell, retailCell, askingCell, pocketCell, actionCell);
    refs.materialBody.append(tr);
  });
}

/** Custom drag preview: card-shaped ghost (native DnD only snapshots setDragImage). */
function buildMeasurementDragPreviewCard(row) {
  const wrap = document.createElement("div");
  wrap.className = "measurement-drag-card measurement-drag-card--ghost";
  wrap.setAttribute("aria-hidden", "true");

  const title = document.createElement("div");
  title.className = "measurement-drag-card-title";
  const room = (row.room || "").trim() || "Room / section";
  const label = (row.label || "").trim() || "Label";
  title.textContent = `${room} · ${label}`;

  const meta = document.createElement("div");
  meta.className = "measurement-drag-card-meta";
  const type = (row.type || "").trim() || "—";
  const code = (row.materialCode || "").trim() || "—";
  meta.textContent = `${type} · ${code}`;

  const dims = document.createElement("div");
  dims.className = "measurement-drag-card-dims";
  if (isMotorizedMaterialRow(row)) {
    const q = parseMotorQuantity(row.unitQuantity);
    dims.textContent =
      q !== null ? `Motorized · qty ${q}` : "Motorized — quantity not set";
  } else {
    const w = (row.width || "").trim();
    const h = (row.height || "").trim();
    dims.textContent = w && h ? `${w} × ${h} mm` : "Width × height (mm)";
  }

  const costEl = document.createElement("div");
  costEl.className = "measurement-drag-card-cost";
  const cost = getMeasurementCost(row);
  if (cost === null) {
    costEl.textContent = "Cost —";
  } else {
    costEl.textContent = row.isFree
      ? `${formatCurrency(cost)} (FREE)`
      : formatCurrency(cost);
  }

  wrap.append(title, meta, dims, costEl);
  return wrap;
}

function updateMeasurementRowDropHighlight(clientX, clientY, draggedRowId) {
  const target = document.elementFromPoint(clientX, clientY);
  const tr = target?.closest?.("tr.measurement-row");
  refs.measurementBody.querySelectorAll(".measurement-row--drag-over").forEach((el) => {
    if (el !== tr) {
      el.classList.remove("measurement-row--drag-over");
    }
  });
  if (tr?.dataset.measurementRowId && tr.dataset.measurementRowId !== draggedRowId) {
    tr.classList.add("measurement-row--drag-over");
  }
}

function cleanupMeasurementPointerDragSession(session) {
  if (!session) {
    return;
  }
  if (session.ghostEl?.parentNode) {
    session.ghostEl.remove();
  }
  session.sourceTr?.classList.remove("measurement-row--dragging");
  refs.measurementBody?.querySelectorAll(".measurement-row--drag-over").forEach((el) => {
    el.classList.remove("measurement-row--drag-over");
  });
  if (measurementPointerDragSession === session) {
    measurementPointerDragSession = null;
  }
}

function applyMeasurementRowDropAtPoint(draggedRowId, clientX, clientY) {
  const target = document.elementFromPoint(clientX, clientY);
  const tr = target?.closest?.("tr.measurement-row");
  if (!tr?.dataset?.measurementRowId) {
    return false;
  }
  const targetId = tr.dataset.measurementRowId;
  if (targetId === draggedRowId) {
    return false;
  }
  const rect = tr.getBoundingClientRect();
  const insertAfter = clientY > rect.top + rect.height / 2;
  if (!reorderMeasurementRowsByIds(draggedRowId, targetId, insertAfter)) {
    return false;
  }
  persistDraftChange();
  renderMeasurements();
  renderSummary();
  return true;
}

/**
 * Pointer-based row drag (mouse, touch, pen). Native HTML5 DnD does not work on most tablets.
 */
function bindMeasurementRowDragControls(dragHandle, tr, row) {
  const onPointerDown = (event) => {
    if (runtime.quoteBusy || event.button !== 0) {
      return;
    }
    if (measurementPointerDragSession) {
      return;
    }
    measurementPointerDragSession = {
      rowId: row.id,
      row,
      pointerId: event.pointerId,
      startX: event.clientX,
      startY: event.clientY,
      active: false,
      ghostEl: null,
      offsetX: 0,
      offsetY: 0,
      handle: dragHandle,
      sourceTr: tr,
    };
    try {
      dragHandle.setPointerCapture(event.pointerId);
    } catch (_) {
      /* ignore */
    }
    event.preventDefault();
  };

  const onPointerMove = (event) => {
    const session = measurementPointerDragSession;
    if (!session || event.pointerId !== session.pointerId) {
      return;
    }
    if (runtime.quoteBusy) {
      return;
    }
    const dx = event.clientX - session.startX;
    const dy = event.clientY - session.startY;
    if (!session.active) {
      if (Math.hypot(dx, dy) < MEASUREMENT_DRAG_ACTIVATE_PX) {
        return;
      }
      session.active = true;
      const ghost = buildMeasurementDragPreviewCard(session.row);
      ghost.style.cssText =
        "position:fixed;left:-9999px;top:0;pointer-events:none;z-index:2147483647;";
      document.body.append(ghost);
      const width = ghost.offsetWidth;
      const height = ghost.offsetHeight;
      session.offsetX = Math.min(72, Math.max(28, Math.floor(width * 0.14)));
      session.offsetY = Math.floor(height / 2);
      ghost.style.left = `${event.clientX - session.offsetX}px`;
      ghost.style.top = `${event.clientY - session.offsetY}px`;
      session.ghostEl = ghost;
      session.sourceTr.classList.add("measurement-row--dragging");
    } else {
      session.ghostEl.style.left = `${event.clientX - session.offsetX}px`;
      session.ghostEl.style.top = `${event.clientY - session.offsetY}px`;
      updateMeasurementRowDropHighlight(event.clientX, event.clientY, session.rowId);
    }
    if (session.active && event.pointerType === "touch") {
      event.preventDefault();
    }
  };

  const onPointerUpOrCancel = (event) => {
    const session = measurementPointerDragSession;
    if (!session || event.pointerId !== session.pointerId) {
      return;
    }
    try {
      dragHandle.releasePointerCapture(session.pointerId);
    } catch (_) {
      /* ignore */
    }
    if (session.active) {
      applyMeasurementRowDropAtPoint(session.rowId, event.clientX, event.clientY);
    }
    cleanupMeasurementPointerDragSession(session);
  };

  dragHandle.addEventListener("pointerdown", onPointerDown);
  dragHandle.addEventListener("pointermove", onPointerMove, { passive: false });
  dragHandle.addEventListener("pointerup", onPointerUpOrCancel);
  dragHandle.addEventListener("pointercancel", onPointerUpOrCancel);
}

/** Move a measurement row relative to another; persists sort order via save pipeline. */
function reorderMeasurementRowsByIds(dragId, targetId, insertAfter) {
  if (dragId === targetId) {
    return false;
  }
  const dragged = state.measurementRows.find((r) => r.id === dragId);
  if (!dragged) {
    return false;
  }
  const without = state.measurementRows.filter((r) => r.id !== dragId);
  const insertBeforeIndex = without.findIndex((r) => r.id === targetId);
  if (insertBeforeIndex === -1) {
    return false;
  }
  const insertIdx = insertAfter ? insertBeforeIndex + 1 : insertBeforeIndex;
  without.splice(insertIdx, 0, dragged);
  state.measurementRows = without;
  return true;
}

function renderMeasurements() {
  refs.measurementBody.innerHTML = "";

  if (state.measurementRows.length === 0) {
    refs.measurementBody.append(
      createEmptyStateRow(11, "Add a measurement row to begin."),
    );
    return;
  }

  const configuredMaterials = getConfiguredMaterials();
  refs.addMeasurementBtn.disabled =
    configuredMaterials.length === 0 || runtime.quoteBusy;

  const preferChevronReorder = Boolean(
    window.matchMedia?.("(pointer: coarse)")?.matches,
  );

  state.measurementRows.forEach((row) => {
    const tr = document.createElement("tr");
    tr.classList.add("measurement-row");
    tr.dataset.measurementRowId = row.id;
    const motorized = isMotorizedMaterialRow(row);

    const dragCell = document.createElement("td");
    dragCell.className = "measurement-drag-cell";

    if (preferChevronReorder) {
      dragCell.classList.add("measurement-reorder-cell");
      const wrap = document.createElement("div");
      wrap.className = "measurement-reorder-controls";

      const buildChevronButton = (direction) => {
        const btn = document.createElement("button");
        btn.type = "button";
        btn.className = "measurement-reorder-button";
        btn.disabled = runtime.quoteBusy;
        btn.setAttribute(
          "aria-label",
          direction === "up" ? "Move row up" : "Move row down",
        );
        btn.innerHTML =
          direction === "up"
            ? '<svg viewBox="0 0 24 24" width="16" height="16" focusable="false" aria-hidden="true"><path fill="currentColor" d="M7.41 14.59 12 10l4.59 4.59L18 13.17 12 7.17l-6 6z"/></svg>'
            : '<svg viewBox="0 0 24 24" width="16" height="16" focusable="false" aria-hidden="true"><path fill="currentColor" d="m7.41 8.41 4.59 4.59 4.59-4.59L18 9.83l-6 6-6-6z"/></svg>';

        btn.addEventListener("click", () => {
          const idx = state.measurementRows.findIndex((r) => r.id === row.id);
          if (idx < 0) {
            return;
          }
          const nextIdx = direction === "up" ? idx - 1 : idx + 1;
          if (nextIdx < 0 || nextIdx >= state.measurementRows.length) {
            return;
          }
          const next = state.measurementRows.slice();
          const tmp = next[idx];
          next[idx] = next[nextIdx];
          next[nextIdx] = tmp;
          state.measurementRows = next;
          persistDraftChange();
          renderMeasurements();
          renderSummary();
        });

        return btn;
      };

      wrap.append(buildChevronButton("up"), buildChevronButton("down"));
      dragCell.append(wrap);
    } else {
      const dragHandle = document.createElement("div");
      dragHandle.className = "measurement-drag-handle";
      dragHandle.setAttribute("role", "button");
      dragHandle.tabIndex = runtime.quoteBusy ? -1 : 0;
      dragHandle.setAttribute("aria-label", "Drag to reorder row");
      dragHandle.setAttribute("data-measurement-drag-handle", "");
      if (runtime.quoteBusy) {
        dragHandle.setAttribute("aria-disabled", "true");
      }
      const dragGlyph = document.createElement("span");
      dragGlyph.className = "measurement-drag-glyph";
      dragGlyph.setAttribute("aria-hidden", "true");
      dragGlyph.textContent = "\u22EE\u22EE";
      dragHandle.append(dragGlyph);
      dragCell.append(dragHandle);
      bindMeasurementRowDragControls(dragHandle, tr, row);
    }

    const roomCell = document.createElement("td");
    const roomInput = buildTextInput({
      value: row.room,
      placeholder: "",
      className: "measurement-fit-field",
      disabled: runtime.quoteBusy,
      onInput: (value) => {
        row.room = value;
        persistDraftChange();
      },
    });
    attachMeasurementFieldFit(roomInput);
    roomCell.append(roomInput);

    const typeCell = document.createElement("td");
    typeCell.className = "material-combobox-cell";

    const typeWrap = document.createElement("div");
    typeWrap.className = "material-combobox";

    const typeInput = document.createElement("input");
    typeInput.type = "text";
    typeInput.className = "material-combobox-input measurement-fit-field";
    typeInput.autocomplete = "off";
    typeInput.spellcheck = false;
    typeInput.placeholder = "Search type…";
    typeInput.setAttribute("role", "combobox");
    typeInput.setAttribute("aria-autocomplete", "list");
    typeInput.setAttribute("aria-expanded", "false");
    typeInput.disabled = runtime.quoteBusy;
    typeInput.value = row.type || "";
    attachMeasurementFieldFit(typeInput);

    const typeList = document.createElement("ul");
    typeList.className = "material-combobox-list";
    typeList.setAttribute("role", "listbox");
    typeList.hidden = true;

    let typeOutsidePointerActive = false;

    const closeTypeList = () => {
      typeList.hidden = true;
      typeInput.setAttribute("aria-expanded", "false");
      if (typeOutsidePointerActive) {
        document.removeEventListener("pointerdown", onOutsideTypePointerDown, true);
        typeOutsidePointerActive = false;
      }
      window.removeEventListener("scroll", onTypeWindowScrollCapture, true);
      window.removeEventListener("resize", positionTypeDropdown);
    };

    const positionTypeDropdown = () => {
      if (typeList.hidden) {
        return;
      }
      const rect = typeInput.getBoundingClientRect();
      const margin = 6;
      const spaceBelow = window.innerHeight - rect.bottom - margin - 8;
      const maxH = Math.min(280, Math.max(120, spaceBelow));
      typeList.style.position = "fixed";
      typeList.style.left = `${rect.left}px`;
      typeList.style.top = `${rect.bottom + margin}px`;
      typeList.style.width = `${Math.max(rect.width, 200)}px`;
      typeList.style.maxHeight = `${maxH}px`;
      typeList.style.zIndex = "3000";
    };

    const onTypeWindowScrollCapture = (event) => {
      if (typeList.hidden) {
        return;
      }
      const t = event.target;
      if (t === typeList) {
        return;
      }
      if (t instanceof Node && typeList.contains(t)) {
        return;
      }
      closeTypeList();
    };

    const onOutsideTypePointerDown = (event) => {
      if (typeList.hidden) {
        return;
      }
      if (event.target instanceof Node && typeWrap.contains(event.target)) {
        return;
      }
      // revert label if it's not an exact option
      typeInput.value = row.type || "";
      closeTypeList();
    };

    function renderTypeListOptions() {
      typeList.innerHTML = "";
      const q = typeInput.value.trim().toLowerCase();
      const filtered = MEASUREMENT_TYPE_OPTIONS.filter((label) =>
        !q ? true : String(label).toLowerCase().includes(q),
      );

      if (filtered.length === 0) {
        const empty = document.createElement("li");
        empty.className = "material-combobox-empty";
        empty.textContent = "No types match your search.";
        typeList.append(empty);
        return;
      }

      filtered.forEach((label) => {
        const li = document.createElement("li");
        li.className = "material-combobox-option";
        li.setAttribute("role", "option");
        const main = document.createElement("div");
        main.className = "material-combobox-option-main";
        main.textContent = label;
        const sub = document.createElement("div");
        sub.className = "material-combobox-option-sub";
        sub.hidden = true;
        li.append(main, sub);
        li.addEventListener("mousedown", (event) => {
          event.preventDefault();
          row.type = label;
          typeInput.value = label;
          closeTypeList();
          persistDraftChange();
          renderMeasurements();
        });
        typeList.append(li);
      });
    }

    const openTypeList = () => {
      if (runtime.quoteBusy) {
        return;
      }
      typeList.hidden = false;
      typeInput.setAttribute("aria-expanded", "true");
      renderTypeListOptions();
      positionTypeDropdown();
      window.addEventListener("scroll", onTypeWindowScrollCapture, true);
      window.addEventListener("resize", positionTypeDropdown);
      if (!typeOutsidePointerActive) {
        document.addEventListener("pointerdown", onOutsideTypePointerDown, true);
        typeOutsidePointerActive = true;
      }
    };

    typeInput.addEventListener("focus", () => {
      if (runtime.quoteBusy) {
        return;
      }
      if (row.type) {
        typeInput.select();
      }
      openTypeList();
    });
    typeInput.addEventListener("input", () => {
      if (runtime.quoteBusy) {
        return;
      }
      renderTypeListOptions();
      if (typeList.hidden) {
        openTypeList();
      } else {
        positionTypeDropdown();
      }
    });
    typeInput.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        typeInput.value = row.type || "";
        closeTypeList();
        typeInput.blur();
      }
    });
    typeInput.addEventListener("blur", () => {
      window.setTimeout(() => {
        if (!typeList.hidden) {
          typeInput.value = row.type || "";
          closeTypeList();
        }
      }, 0);
    });

    typeWrap.append(typeInput, typeList);
    typeCell.append(typeWrap);

    const materialCodeCell = document.createElement("td");
    const materialCodeInput = buildTextInput({
      value: row.materialCode,
      placeholder: "Material code",
      className: "measurement-fit-field",
      disabled: runtime.quoteBusy,
      onInput: (value) => {
        row.materialCode = value;
        persistDraftChange();
      },
    });
    attachMeasurementFieldFit(materialCodeInput);
    materialCodeCell.append(materialCodeInput);

    const labelCell = document.createElement("td");
    const labelInput = buildTextInput({
      value: row.label,
      placeholder: "W1",
      className: "measurement-fit-field",
      disabled: runtime.quoteBusy,
      onInput: (value) => {
        row.label = value;
        persistDraftChange();
      },
    });
    attachMeasurementFieldFit(labelInput);
    labelCell.append(labelInput);

    const widthCell = document.createElement("td");
    const widthInput = buildNumberInput({
      value: motorized ? "" : row.width,
      placeholder: "0",
      className: "measurement-fit-field",
      disabled: runtime.quoteBusy || motorized,
      onInput: (value) => {
        row.width = value;
        persistDraftChange();
        updateMeasurementOutputs();
        renderSummary();
      },
    });
    attachMeasurementFieldFit(widthInput);
    widthCell.append(widthInput);

    const heightCell = document.createElement("td");
    const heightInput = buildNumberInput({
      value: motorized ? "" : row.height,
      placeholder: "0",
      className: "measurement-fit-field",
      disabled: runtime.quoteBusy || motorized,
      onInput: (value) => {
        row.height = value;
        persistDraftChange();
        updateMeasurementOutputs();
        renderSummary();
      },
    });
    attachMeasurementFieldFit(heightInput);
    heightCell.append(heightInput);

    const materialCell = document.createElement("td");
    materialCell.className = "material-combobox-cell";

    const materialWrap = document.createElement("div");
    materialWrap.className = "material-combobox";

    const materialInput = document.createElement("input");
    materialInput.type = "text";
    materialInput.className = "material-combobox-input measurement-fit-field";
    materialInput.autocomplete = "off";
    materialInput.spellcheck = false;
    materialInput.placeholder = configuredMaterials.length === 0
      ? "Add configured material first"
      : "Search material…";
    materialInput.setAttribute("role", "combobox");
    materialInput.setAttribute("aria-autocomplete", "list");
    materialInput.setAttribute("aria-expanded", "false");
    materialInput.disabled = configuredMaterials.length === 0 || runtime.quoteBusy;
    attachMeasurementFieldFit(materialInput);

    const currentConfigured = configuredMaterials.find((m) => m.id === row.materialId);
    materialInput.value = currentConfigured
      ? `${currentConfigured.category} (${formatCurrency(currentConfigured.askingPrice)})`
      : "";

    const materialList = document.createElement("ul");
    materialList.className = "material-combobox-list";
    materialList.setAttribute("role", "listbox");
    materialList.hidden = true;

    let materialOutsidePointerActive = false;
    let prevMaterialId = row.materialId || "";

    const closeMaterialList = () => {
      materialList.hidden = true;
      materialInput.setAttribute("aria-expanded", "false");
      if (materialOutsidePointerActive) {
        document.removeEventListener("pointerdown", onOutsideMaterialPointerDown, true);
        materialOutsidePointerActive = false;
      }
      window.removeEventListener("scroll", onMaterialWindowScrollCapture, true);
      window.removeEventListener("resize", positionMaterialDropdown);
    };

    const positionMaterialDropdown = () => {
      if (materialList.hidden) {
        return;
      }
      const rect = materialInput.getBoundingClientRect();
      const margin = 6;
      const spaceBelow = window.innerHeight - rect.bottom - margin - 8;
      const maxH = Math.min(320, Math.max(140, spaceBelow));
      materialList.style.position = "fixed";
      materialList.style.left = `${rect.left}px`;
      materialList.style.top = `${rect.bottom + margin}px`;
      materialList.style.width = `${Math.max(rect.width, 220)}px`;
      materialList.style.maxHeight = `${maxH}px`;
      materialList.style.zIndex = "3000";
    };

    const syncMaterialInputLabelFromRow = () => {
      const selected = configuredMaterials.find((m) => m.id === row.materialId);
      materialInput.value = selected
        ? `${selected.category} (${formatCurrency(selected.askingPrice)})`
        : "";
    };

    const onMaterialWindowScrollCapture = (event) => {
      if (materialList.hidden) {
        return;
      }
      const t = event.target;
      if (t === materialList) {
        return;
      }
      if (t instanceof Node && materialList.contains(t)) {
        return;
      }
      closeMaterialList();
    };

    const onOutsideMaterialPointerDown = (event) => {
      if (materialList.hidden) {
        return;
      }
      if (event.target instanceof Node && materialWrap.contains(event.target)) {
        return;
      }
      syncMaterialInputLabelFromRow();
      closeMaterialList();
    };

    const buildMaterialOptionMain = (material) =>
      `${material.category} (${formatCurrency(material.askingPrice)})`;

    function renderMaterialListOptions() {
      materialList.innerHTML = "";
      const q = materialInput.value.trim().toLowerCase();
      const matches = (m) => buildMaterialOptionMain(m).toLowerCase().includes(q);

      const seen = new Set();
      const recentIds = runtime.recentMeasurementMaterialIds || [];
      const recent = !q
        ? recentIds
            .map((id) => configuredMaterials.find((m) => m.id === id))
            .filter(Boolean)
        : [];

      const addItem = (material, { subLabel } = {}) => {
        if (!material || seen.has(material.id)) {
          return;
        }
        seen.add(material.id);
        const li = document.createElement("li");
        li.className = "material-combobox-option";
        li.setAttribute("role", "option");

        const main = document.createElement("div");
        main.className = "material-combobox-option-main";
        main.textContent = buildMaterialOptionMain(material);

        const sub = document.createElement("div");
        sub.className = "material-combobox-option-sub";
        if (subLabel) {
          sub.textContent = subLabel;
        } else {
          sub.hidden = true;
        }

        li.append(main, sub);
        li.addEventListener("mousedown", (event) => {
          event.preventDefault();
          const picked = material;
          prevMaterialId = row.materialId || "";
          row.materialId = picked.id;
          syncMaterialInputLabelFromRow();
          closeMaterialList();

          if (picked && isMotorizedCategoryFromMaterial(picked)) {
            void openMotorQuantityDialog(row).then((result) => {
              if (!result.ok) {
                row.materialId = prevMaterialId;
                syncMaterialInputLabelFromRow();
              } else {
                row.unitQuantity = String(result.quantity);
                row.width = "";
                row.height = "";
                pushRecentMeasurementMaterial(picked.id);
              }
              persistDraftChange();
              renderMeasurements();
              renderSummary();
            });
            return;
          }

          row.unitQuantity = "";
          pushRecentMeasurementMaterial(picked.id);
          persistDraftChange();
          renderMeasurements();
          renderSummary();
        });
        materialList.append(li);
      };

      if (recent.length > 0) {
        recent.forEach((m) => addItem(m, { subLabel: "Recent" }));
      }

      configuredMaterials
        .filter((m) => (q ? matches(m) : true))
        .forEach((m) => addItem(m));

      if (materialList.children.length === 0) {
        const empty = document.createElement("li");
        empty.className = "material-combobox-empty";
        empty.textContent = "No materials match your search.";
        materialList.append(empty);
      }
    }

    const openMaterialList = () => {
      if (configuredMaterials.length === 0 || runtime.quoteBusy) {
        return;
      }
      materialList.hidden = false;
      materialInput.setAttribute("aria-expanded", "true");
      renderMaterialListOptions();
      positionMaterialDropdown();
      window.addEventListener("scroll", onMaterialWindowScrollCapture, true);
      window.addEventListener("resize", positionMaterialDropdown);
      if (!materialOutsidePointerActive) {
        document.addEventListener("pointerdown", onOutsideMaterialPointerDown, true);
        materialOutsidePointerActive = true;
      }
    };

    materialInput.addEventListener("focus", () => {
      if (configuredMaterials.length === 0 || runtime.quoteBusy) {
        return;
      }
      prevMaterialId = row.materialId || "";
      if (row.materialId) {
        materialInput.select();
      }
      openMaterialList();
    });
    materialInput.addEventListener("input", () => {
      if (configuredMaterials.length === 0 || runtime.quoteBusy) {
        return;
      }
      renderMaterialListOptions();
      if (materialList.hidden) {
        openMaterialList();
      } else {
        positionMaterialDropdown();
      }
    });
    materialInput.addEventListener("keydown", (event) => {
      if (event.key === "Escape") {
        event.preventDefault();
        syncMaterialInputLabelFromRow();
        closeMaterialList();
        materialInput.blur();
      }
    });
    materialInput.addEventListener("blur", () => {
      window.setTimeout(() => {
        if (!materialList.hidden) {
          syncMaterialInputLabelFromRow();
          closeMaterialList();
        }
      }, 0);
    });

    materialWrap.append(materialInput, materialList);
    materialCell.append(materialWrap);

    const costCell = document.createElement("td");
    costCell.className = "money-cell";
    const updateMeasurementOutputs = () => {
      const cost = getMeasurementCost(row);
      const squareFootage = getMeasurementSquareFootage(row);
      const motorizedRow = isMotorizedMaterialRow(row);
      const qty = parseMotorQuantity(row.unitQuantity);

      const baseCostText = cost === null ? "PHP 0.00" : formatCurrency(cost);
      costCell.textContent = row.isFree ? `${baseCostText} (FREE)` : baseCostText;

      if (motorizedRow) {
        const helper = document.createElement("span");
        helper.className = "muted-helper";
        helper.textContent = qty
          ? `${qty} quantity`
          : "Quantity required — pick Curtains Motorized again to enter";
        costCell.append(helper);
        return;
      }

      if (squareFootage !== null) {
        const helper = document.createElement("span");
        helper.className = squareFootage.minimumApplied
          ? "muted-helper sqft-warning"
          : "muted-helper";
        helper.textContent = squareFootage.minimumApplied
          ? `${squareFootage.rawRounded} sqft -> rounded to ${squareFootage.billed} sqft minimum`
          : `${squareFootage.billed} sqft rounded`;
        costCell.append(helper);
      }
    };
    updateMeasurementOutputs();

    const freeCell = document.createElement("td");
    freeCell.className = "measurement-free-cell";
    const freeInput = document.createElement("input");
    freeInput.type = "checkbox";
    freeInput.className = "measurement-free-checkbox";
    freeInput.checked = Boolean(row.isFree);
    freeInput.disabled = runtime.quoteBusy;
    freeInput.setAttribute("aria-label", "Complimentary line (reference cost only)");
    freeInput.addEventListener("change", (event) => {
      row.isFree = Boolean(event.target.checked);
      persistDraftChange();
      updateMeasurementOutputs();
      renderSummary();
    });
    freeCell.append(freeInput);

    const actionCell = document.createElement("td");
    const actionStack = document.createElement("div");
    actionStack.className = "measurement-row-actions";

    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "ghost-danger-button measurement-remove-button";
    removeButton.setAttribute("aria-label", "Remove row");
    removeButton.innerHTML = [
      '<span class="measurement-remove-label">Remove</span>',
      '<span class="measurement-remove-icon" aria-hidden="true">',
      '<svg viewBox="0 0 24 24" width="18" height="18" focusable="false" aria-hidden="true">',
      '<path fill="currentColor" d="M9 3h6l1 2h4v2H4V5h4l1-2zm1 6h2v10h-2V9zm4 0h2v10h-2V9zM6 9h2v10H6V9zm2 12h8a2 2 0 0 0 2-2V9H4v10a2 2 0 0 0 2 2z"/>',
      "</svg>",
      "</span>",
    ].join("");
    removeButton.disabled = runtime.quoteBusy;
    removeButton.addEventListener("click", () => {
      state.measurementRows = state.measurementRows.filter(
        (measurement) => measurement.id !== row.id,
      );

      if (state.measurementRows.length === 0) {
        state.measurementRows.push(createMeasurementRow());
      }

      persistDraftChange();
      renderMeasurements();
      renderSummary();
    });

    const addBelowButton = document.createElement("button");
    addBelowButton.type = "button";
    addBelowButton.className = "secondary-button measurement-add-below-button";
    addBelowButton.setAttribute("aria-label", "Add row below");
    addBelowButton.disabled = runtime.quoteBusy;
    addBelowButton.innerHTML = [
      '<span class="measurement-add-below-label">Add below</span>',
      '<span class="measurement-add-below-icon" aria-hidden="true">+</span>',
    ].join("");
    addBelowButton.addEventListener("click", () => {
      const rowIndex = state.measurementRows.findIndex((item) => item.id === row.id);
      if (rowIndex < 0) {
        return;
      }
      state.measurementRows.splice(rowIndex + 1, 0, createMeasurementRow());
      persistDraftChange();
      queueAutosave();
      renderMeasurements();
      renderSummary();
    });

    actionStack.append(removeButton, addBelowButton);
    actionCell.append(actionStack);

    tr.append(
      dragCell,
      roomCell,
      labelCell,
      typeCell,
      materialCodeCell,
      widthCell,
      heightCell,
      materialCell,
      freeCell,
      costCell,
      actionCell,
    );
    refs.measurementBody.append(tr);
  });
}

function renderSummary() {
  refs.discountType.value = state.discountType;
  refs.discountValue.value = getDiscountInputDisplayValue(state.discountValue);
  refs.discountType.disabled = runtime.quoteBusy;
  refs.discountValue.disabled = runtime.quoteBusy;
  refs.discountValue.placeholder =
    state.discountType === "percent" ? "Enter percent" : "Enter amount";
  refs.discountValue.step = state.discountType === "percent" ? "0.01" : "0.01";
  refs.discountValue.max = state.discountType === "percent" ? "100" : "";

  const { subtotal, discountAmount, finalTotal, half } = getSummaryTotals();

  document.querySelector("#subtotal-value").textContent = formatCurrency(subtotal);
  document.querySelector("#applied-discount-value").textContent =
    formatCurrency(discountAmount);
  document.querySelector("#final-total-value").textContent =
    formatCurrency(finalTotal);
  document.querySelector("#downpayment-value").textContent = formatCurrency(half);
  document.querySelector("#remaining-value").textContent = formatCurrency(half);
  renderActiveQuoteBar();
}

function getConfiguredMaterials() {
  return state.selectedMaterials.filter((row) => {
    const askingPrice = parseCurrencyLikeNumber(row.askingPrice);
    return Boolean(row.category) && askingPrice !== null;
  });
}

function getMeasurementMaterialForRow(row) {
  return state.selectedMaterials.find((material) => material.id === row.materialId);
}

function isMotorizedCategoryFromMaterial(material) {
  return material?.category?.trim() === MOTORIZED_MATERIAL_CATEGORY;
}

function isMotorizedMaterialRow(row) {
  return isMotorizedCategoryFromMaterial(getMeasurementMaterialForRow(row));
}

function parseMotorQuantity(value) {
  const parsed = Number.parseInt(String(value ?? "").trim(), 10);
  if (!Number.isFinite(parsed) || parsed < 1) {
    return null;
  }
  return parsed;
}

function openMotorQuantityDialog(row) {
  const dialog = refs.motorQuantityDialog;
  const input = refs.motorQuantityInput;
  const form = refs.motorQuantityForm;
  const cancelBtn = refs.motorQuantityCancel;

  return new Promise((resolve) => {
    if (!dialog || !input || !form) {
      resolve({ ok: false });
      return;
    }

    input.value = row.unitQuantity ? String(row.unitQuantity) : "";

    let settled = false;

    const cleanup = () => {
      form.removeEventListener("submit", onSubmit);
      cancelBtn?.removeEventListener("click", onCancel);
      dialog.removeEventListener("cancel", onDismiss);
    };

    const finish = (result) => {
      if (settled) {
        return;
      }
      settled = true;
      cleanup();
      resolve(result);
    };

    const onSubmit = (event) => {
      event.preventDefault();
      const qty = parseMotorQuantity(input.value);
      if (qty === null) {
        input.focus();
        return;
      }
      dialog.close();
      finish({ ok: true, quantity: qty });
    };

    const onCancel = () => {
      dialog.close();
      finish({ ok: false });
    };

    const onDismiss = () => {
      finish({ ok: false });
    };

    form.addEventListener("submit", onSubmit);
    cancelBtn?.addEventListener("click", onCancel);
    dialog.addEventListener("cancel", onDismiss);
    dialog.showModal();
    window.requestAnimationFrame(() => input.focus());
  });
}

function getPocketValue(materialRow) {
  const askingPrice = parseCurrencyLikeNumber(materialRow.askingPrice);
  const retailPrice = parseCurrencyLikeNumber(materialRow.retailPrice);

  if (askingPrice === null || retailPrice === null) {
    return null;
  }

  return askingPrice - retailPrice;
}

function getMeasurementSquareFootage(row) {
  if (isMotorizedMaterialRow(row)) {
    return null;
  }

  const width = parseCurrencyLikeNumber(row.width);
  const height = parseCurrencyLikeNumber(row.height);

  if (width === null || height === null) {
    return null;
  }

  if (width <= 0 || height <= 0) {
    return null;
  }

  const rawRounded = Math.round(((width / 1000) * (height / 1000)) * 10.76);

  return {
    rawRounded,
    billed: Math.max(rawRounded, MINIMUM_SQUARE_FEET),
    minimumApplied: rawRounded < MINIMUM_SQUARE_FEET,
  };
}

function getMeasurementCost(row) {
  const material = getMeasurementMaterialForRow(row);
  const askingPrice = parseCurrencyLikeNumber(material?.askingPrice);

  if (askingPrice === null) {
    return null;
  }

  if (isMotorizedMaterialRow(row)) {
    const qty = parseMotorQuantity(row.unitQuantity);
    if (qty === null) {
      return null;
    }
    return qty * askingPrice;
  }

  const squareFootage = getMeasurementSquareFootage(row);
  if (squareFootage === null) {
    return null;
  }

  return squareFootage.billed * askingPrice;
}

function getSubtotal() {
  const delivery = state.quoteMeta.deliveryIsFree
    ? 0
    : parseCurrencyLikeNumber(state.quoteMeta.deliveryAmount) || 0;
  const installSteam = state.quoteMeta.installSteamIsFree
    ? 0
    : parseCurrencyLikeNumber(state.quoteMeta.installSteamAmount) || 0;

  return state.measurementRows.reduce((total, row) => {
    if (row.isFree) {
      return total;
    }
    const cost = getMeasurementCost(row);
    return total + (cost ?? 0);
  }, 0) + delivery + installSteam;
}

function getSummaryTotals() {
  const subtotal = getSubtotal();
  const discountAmount = getAppliedDiscountAmount(subtotal);
  const finalTotal = Math.max(0, subtotal - discountAmount);
  const half = finalTotal / 2;

  return {
    subtotal,
    discountAmount,
    finalTotal,
    half,
  };
}

function getAppliedDiscountAmount(subtotal = getSubtotal()) {
  const rawDiscountValue = parseCurrencyLikeNumber(state.discountValue) || 0;

  if (rawDiscountValue <= 0 || subtotal <= 0) {
    return 0;
  }

  if (state.discountType === "percent") {
    const boundedPercent = Math.min(Math.max(rawDiscountValue, 0), 100);
    return subtotal * (boundedPercent / 100);
  }

  return Math.min(Math.max(rawDiscountValue, 0), subtotal);
}

function validateQuoteForSave() {
  if (!state.quoteMeta.clientName.trim()) {
    return { ok: false, message: "Client Name is required before saving a quote." };
  }

  const deliveryAmount = parseCurrencyLikeNumber(state.quoteMeta.deliveryAmount) || 0;
  if (deliveryAmount < 0) {
    return { ok: false, message: "Delivery amount cannot be negative." };
  }
  const installSteamAmount = parseCurrencyLikeNumber(state.quoteMeta.installSteamAmount) || 0;
  if (installSteamAmount < 0) {
    return { ok: false, message: "Installation and steam amount cannot be negative." };
  }

  const discountValue = parseCurrencyLikeNumber(state.discountValue) || 0;
  if (discountValue < 0) {
    return { ok: false, message: "Discount cannot be negative." };
  }

  if (state.discountType === "percent" && discountValue > 100) {
    return {
      ok: false,
      message: "Percent discount must be between 0 and 100.",
    };
  }

  if (state.discountType === "amount" && discountValue > getSubtotal()) {
    return {
      ok: false,
      message: "Amount discount cannot exceed the subtotal.",
    };
  }

  const materialValidation = getMaterialDraftsForSave(true);
  if (!materialValidation.ok) {
    return materialValidation;
  }

  const measurementValidation = getMeasurementDraftsForSave(true);
  if (!measurementValidation.ok) {
    return measurementValidation;
  }

  return { ok: true };
}

function getMaterialDraftsForSave(validateOnly = false) {
  const drafts = [];

  for (const row of state.selectedMaterials) {
    const hasAnyValue =
      Boolean(row.sourceKey) ||
      Boolean(row.category) ||
      Boolean(row.division) ||
      row.retailPrice !== "" ||
      row.askingPrice !== "";

    if (!hasAnyValue) {
      continue;
    }

    const retailPrice = parseCurrencyLikeNumber(row.retailPrice);
    const askingPrice = parseCurrencyLikeNumber(row.askingPrice);

    if (!row.category || retailPrice === null || askingPrice === null) {
      return {
        ok: false,
        message:
          "Every material row must have a material, retail price, and asking price before saving.",
      };
    }

    if (!validateOnly) {
      const sourceMaterial = state.sourceMaterials.find(
        (material) => material.key === row.sourceKey,
      );

      drafts.push({
        localMaterialId: row.id,
        catalogId: sourceMaterial?.catalogId || null,
        division: row.division || sourceMaterial?.division || "",
        category: row.category,
        retailPrice,
        askingPrice,
      });
    }
  }

  return validateOnly ? { ok: true } : drafts;
}

function getMeasurementDraftsForSave(validateOnly = false) {
  const drafts = [];

  for (const row of state.measurementRows) {
    const hasAnyValue =
      Boolean(row.room) ||
      Boolean(row.type) ||
      Boolean(row.materialCode) ||
      Boolean(row.label) ||
      row.width !== "" ||
      row.height !== "" ||
      Boolean(row.materialId) ||
      row.unitQuantity !== "";

    if (!hasAnyValue) {
      continue;
    }

    const selectedMaterial = state.selectedMaterials.find(
      (material) => material.id === row.materialId,
    );
    const askingPrice = parseCurrencyLikeNumber(selectedMaterial?.askingPrice);

    if (!selectedMaterial || askingPrice === null) {
      return {
        ok: false,
        message:
          "Every measurement row must use a configured material with a valid asking price before saving.",
      };
    }

    const motorized = isMotorizedCategoryFromMaterial(selectedMaterial);

    if (motorized) {
      const qty = parseMotorQuantity(row.unitQuantity);
      if (qty === null) {
        return {
          ok: false,
          message:
            "Curtains Motorized rows need a quantity. Select the material again and enter a whole number (1 or more).",
        };
      }

      const lineCost = qty * askingPrice;

      if (!validateOnly) {
        drafts.push({
          localMaterialId: selectedMaterial.id,
          room: row.room,
          type: row.type,
          materialCode: row.materialCode,
          label: row.label.trim(),
          width: 0,
          height: 0,
          unitQuantity: qty,
          lineCost,
          materialLabel: selectedMaterial.category,
          askingPrice,
          isFree: Boolean(row.isFree),
        });
      }
      continue;
    }

    const width = parseCurrencyLikeNumber(row.width);
    const height = parseCurrencyLikeNumber(row.height);

    if (width === null || height === null) {
      return {
        ok: false,
        message:
          "Every measurement row must have width, height, and a configured material before saving.",
      };
    }

    const squareFootage = getMeasurementSquareFootage(row);
    if (squareFootage === null) {
      return {
        ok: false,
        message:
          "Every measurement row must have width, height, and a configured material before saving.",
      };
    }

    const lineCost = squareFootage.billed * askingPrice;

    if (!validateOnly) {
      drafts.push({
        localMaterialId: selectedMaterial.id,
        room: row.room,
        type: row.type,
        materialCode: row.materialCode,
        label: row.label.trim(),
        width,
        height,
        unitQuantity: null,
        lineCost,
        materialLabel: selectedMaterial.category,
        askingPrice,
        isFree: Boolean(row.isFree),
      });
    }
  }

  return validateOnly ? { ok: true } : drafts;
}

function findSourceKey(catalogId, division, category) {
  const byCatalogId = state.sourceMaterials.find(
    (item) => item.catalogId && item.catalogId === catalogId,
  );
  if (byCatalogId) {
    return byCatalogId.key;
  }

  const byLabels = state.sourceMaterials.find(
    (item) => item.category === category && (item.division || "") === (division || ""),
  );

  return byLabels?.key || "";
}

function resetQuoteDraft() {
  state.quoteMeta = getDefaultQuoteMeta();
  state.discountType = "amount";
  state.discountValue = "";
  state.selectedMaterials = [createMaterialSetupRow()];
  state.measurementRows = [createMeasurementRow()];
  runtime.loadedQuoteFingerprint = "";
  clearSaveReminder();
  runtime.recentMeasurementMaterialIds = [];
}

function clearQueuedAutosave() {
  if (!runtime.autosaveTimer) {
    return;
  }

  window.clearTimeout(runtime.autosaveTimer);
  runtime.autosaveTimer = 0;
}

function clearSaveReminderTimer() {
  if (!runtime.saveReminderTimer) {
    return;
  }
  window.clearTimeout(runtime.saveReminderTimer);
  runtime.saveReminderTimer = 0;
}

function clearSaveReminderRepeatTimer() {
  if (!runtime.saveReminderRepeatTimer) {
    return;
  }
  window.clearInterval(runtime.saveReminderRepeatTimer);
  runtime.saveReminderRepeatTimer = 0;
}

function clearSaveReminder() {
  clearSaveReminderTimer();
  clearSaveReminderRepeatTimer();
  runtime.saveReminderActive = false;
  hideSaveReminderToast();
  renderSaveReminderButtons();
}

function showSaveReminderToast() {
  if (!refs.saveReminderToast) {
    return;
  }
  refs.saveReminderToast.textContent = "Unsaved changes reminder: tap Save Quote.";
  refs.saveReminderToast.classList.add("is-visible");
  window.setTimeout(() => {
    refs.saveReminderToast?.classList.remove("is-visible");
  }, 3200);
}

function hideSaveReminderToast() {
  refs.saveReminderToast?.classList.remove("is-visible");
}

function scheduleSaveReminder() {
  clearSaveReminderTimer();
  clearSaveReminderRepeatTimer();
  runtime.saveReminderActive = false;

  if (!runtime.session || runtime.quoteBusy || !hasMeaningfulDraftChanges()) {
    hideSaveReminderToast();
    renderSaveReminderButtons();
    return;
  }

  runtime.saveReminderTimer = window.setTimeout(() => {
    runtime.saveReminderTimer = 0;
    if (!runtime.session || runtime.quoteBusy || !hasMeaningfulDraftChanges()) {
      hideSaveReminderToast();
      renderSaveReminderButtons();
      return;
    }
    runtime.saveReminderActive = true;
    showSaveReminderToast();
    runtime.saveReminderRepeatTimer = window.setInterval(() => {
      if (!shouldShowSaveReminder()) {
        clearSaveReminderRepeatTimer();
        hideSaveReminderToast();
        return;
      }
      showSaveReminderToast();
    }, SAVE_REMINDER_REPEAT_MS);
    renderSaveReminderButtons();
  }, SAVE_REMINDER_DELAY_MS);
}

function shouldShowSaveReminder() {
  return Boolean(runtime.saveReminderActive) &&
    !runtime.quoteBusy &&
    Boolean(runtime.session) &&
    hasMeaningfulDraftChanges();
}

function renderSaveReminderButtons() {
  const shouldPulse = shouldShowSaveReminder();
  refs.saveQuoteBtn?.classList.toggle(
    "save-reminder-pulse",
    shouldPulse && !refs.saveQuoteBtn.disabled,
  );
  refs.activeQuoteSaveBtn?.classList.toggle(
    "save-reminder-pulse",
    shouldPulse && !refs.activeQuoteSaveBtn.disabled,
  );
}

function canAutosaveCurrentQuote() {
  return Boolean(runtime.session && state.quoteMeta.id) &&
    !runtime.quoteBusy &&
    hasLoadedQuoteUnsavedChanges();
}

/** Schedules a debounced Supabase save. Call sites: Measurements → Add Row. */
function queueAutosave() {
  clearQueuedAutosave();

  if (!AUTOSAVE_ENABLED) {
    return;
  }

  if (!canAutosaveCurrentQuote()) {
    return;
  }

  runtime.autosaveTimer = window.setTimeout(() => {
    runtime.autosaveTimer = 0;

    if (!canAutosaveCurrentQuote()) {
      return;
    }

    void handleSaveQuote({ autosave: true });
  }, AUTOSAVE_DELAY_MS);
}

function isQuoteWorkspaceActive() {
  return runtime.quoteWorkspaceActive || Boolean(state.quoteMeta.id);
}

function hasMeaningfulDraftChanges() {
  return Boolean(
    state.quoteMeta.clientName ||
      state.quoteMeta.projectName ||
      state.quoteMeta.quoteDate ||
      state.quoteMeta.projectArchitect ||
      state.quoteMeta.contactNumber ||
      state.quoteMeta.emailAddress ||
      state.quoteMeta.notes ||
      state.discountValue ||
      state.selectedMaterials.some((row) => row.category || row.askingPrice !== "") ||
      state.measurementRows.some(
        (row) =>
          row.room ||
          row.type ||
          row.materialCode ||
          row.label ||
          row.width !== "" ||
          row.height !== "" ||
          row.materialId ||
          row.unitQuantity !== "" ||
          row.isFree,
      ),
  );
}

function hasLoadedQuoteUnsavedChanges() {
  if (!state.quoteMeta.id || !runtime.loadedQuoteFingerprint) {
    return false;
  }

  return buildCurrentQuoteFingerprint() !== runtime.loadedQuoteFingerprint;
}

function buildCurrentQuoteFingerprint() {
  return JSON.stringify({
    quoteId: state.quoteMeta.id || "",
    clientName: state.quoteMeta.clientName.trim(),
    projectName: state.quoteMeta.projectName.trim(),
    quoteDate: state.quoteMeta.quoteDate,
    projectArchitect: state.quoteMeta.projectArchitect.trim(),
    contactNumber: state.quoteMeta.contactNumber.trim(),
    emailAddress: state.quoteMeta.emailAddress.trim(),
    notes: state.quoteMeta.notes.trim(),
    discountType: state.discountType,
    discountValue: state.discountValue,
    selectedMaterials: state.selectedMaterials.map((row) => ({
      sourceKey: row.sourceKey || "",
      category: row.category || "",
      division: row.division || "",
      retailPrice: row.retailPrice === "" ? "" : String(row.retailPrice),
      askingPrice: row.askingPrice === "" ? "" : String(row.askingPrice),
    })),
    measurementRows: state.measurementRows.map((row) => ({
      room: row.room || "",
      type: row.type || "",
      materialCode: row.materialCode || "",
      label: row.label || "",
      width: row.width === "" ? "" : String(row.width),
      height: row.height === "" ? "" : String(row.height),
      materialId: row.materialId || "",
      unitQuantity: row.unitQuantity === "" ? "" : String(row.unitQuantity),
      isFree: Boolean(row.isFree),
    })),
  });
}

function buildActiveQuoteTitle() {
  const parts = [];
  if (state.quoteMeta.clientName.trim()) {
    parts.push(state.quoteMeta.clientName.trim());
  }
  if (state.quoteMeta.projectName.trim()) {
    parts.push(state.quoteMeta.projectName.trim());
  }

  return parts.join(" • ") || "Saved quote";
}

function getCurrentQuoteLabel() {
  if (state.quoteMeta.clientName.trim()) {
    return state.quoteMeta.id
      ? `${state.quoteMeta.clientName} (saved)`
      : `${state.quoteMeta.clientName} (draft)`;
  }

  if (state.quoteMeta.id) {
    return "Saved quote";
  }

  return hasMeaningfulDraftChanges() ? "Unsaved draft" : "No quote selected";
}

function buildQuoteCardSubtitle(quote) {
  const parts = [];
  if (quote.project_name) {
    parts.push(quote.project_name);
  }

  return parts.length > 0 ? parts.join(" • ") : "No project yet.";
}

function setAuthStatus(message, isError = false) {
  refs.authStatus.textContent = message;
  refs.authStatus.classList.toggle("is-error", isError);
}

function setQuoteStatus(message, isError = false) {
  refs.quoteStatus.textContent = message;
  refs.quoteStatus.classList.toggle("is-error", isError);
}

function formatTimeAgoShort(isoValue) {
  if (!isoValue) {
    return "";
  }
  const ms = new Date(isoValue).getTime();
  if (!Number.isFinite(ms)) {
    return "";
  }
  const delta = Date.now() - ms;
  if (!Number.isFinite(delta) || delta < 0) {
    return "";
  }
  if (delta < 10_000) return "just now";
  if (delta < 60_000) return `${Math.round(delta / 1000)}s ago`;
  if (delta < 60 * 60_000) return `${Math.round(delta / 60_000)}m ago`;
  return `${Math.round(delta / (60 * 60_000))}h ago`;
}

function getQuoteSaveIndicatorState() {
  const signedIn = Boolean(runtime.session);
  const hasId = Boolean(state.quoteMeta.id);
  const unsaved = hasLoadedQuoteUnsavedChanges();

  if (!signedIn) {
    if (hasMeaningfulDraftChanges()) {
      return { text: "Unsaved (sign in to save)", tone: "warning" };
    }
    return { text: "", tone: "" };
  }
  if (runtime.quoteBusy) {
    return { text: "Saving...", tone: "" };
  }
  if (hasId && unsaved) {
    return {
      text: shouldShowSaveReminder() ? "Unsaved changes — please save" : "Unsaved changes",
      tone: "warning",
    };
  }
  if (hasId) {
    const timeAgo = formatTimeAgoShort(state.quoteMeta.updatedAt);
    return { text: timeAgo ? `Saved ${timeAgo}` : "Saved", tone: "" };
  }
  if (hasMeaningfulDraftChanges()) {
    return { text: "Unsaved draft", tone: "warning" };
  }
  return { text: "", tone: "" };
}

function renderQuoteSaveIndicator() {
  if (!refs.quoteSaveIndicator) {
    return;
  }
  const { text, tone } = getQuoteSaveIndicatorState();
  refs.quoteSaveIndicator.textContent = text;
  refs.quoteSaveIndicator.classList.toggle("is-warning", tone === "warning");
  refs.quoteSaveIndicator.classList.toggle("is-error", tone === "error");
}

async function handleExportPdf() {
  syncQuoteMetaFromInputs();
  const validation = validateQuoteForSave();
  if (!validation.ok) {
    setQuoteStatus(validation.message, true);
    window.alert(validation.message);
    return;
  }

  const pdfMake = window.pdfMake;
  if (!pdfMake?.createPdf) {
    const message = "PDF export library did not load. Refresh the page and try again.";
    setQuoteStatus(message, true);
    window.alert(message);
    return;
  }

  registerPdfFonts(pdfMake);

  runtime.exportPdfBusy = true;
  renderQuoteWorkspace();
  setQuoteStatus("Preparing contract PDF (1/3)...");

  const popupWindow = window.open("", "_blank");

  try {
    setQuoteStatus("Loading PDF assets (2/3)...");
    const contract = buildContractPreviewData();
    const assets = await loadContractPdfAssets();
    setQuoteStatus("Rendering PDF (3/3)...");
    const documentDefinition = buildContractPdfDefinition(contract, assets, {
      organization: state.pdfOrganization,
    });
    const pdfBlob = await getPdfBlob(pdfMake.createPdf(documentDefinition));
    const pdfUrl = URL.createObjectURL(pdfBlob);
    const fileName = buildContractPdfFileName(contract, state.pdfOrganization);

    if (openPdfPreviewWindow(popupWindow, pdfUrl, fileName)) {
      popupWindow.focus();
      setQuoteStatus("Contract PDF opened in a new tab. Use Download PDF to save it with the correct filename.");
    } else {
      triggerPdfDownload(pdfUrl, fileName);
      setQuoteStatus(
        popupWindow
          ? "Contract PDF downloaded."
          : "Popup blocked by browser. Contract PDF downloaded instead.",
      );
    }

    schedulePdfUrlCleanup(pdfUrl, popupWindow);
  } catch (error) {
    console.error(error);
    if (popupWindow && !popupWindow.closed) {
      popupWindow.close();
    }
    const details = error?.message ? `\n\nDetails: ${error.message}` : "";
    const message = `Could not generate the contract PDF.${details}\n\nTry refreshing the page, then export again. If the issue persists, confirm you're signed in and that your internet connection is stable.`;
    setQuoteStatus(message, true);
    window.alert(message);
  } finally {
    runtime.exportPdfBusy = false;
    renderQuoteWorkspace();
  }
}

function registerPdfFonts(pdfMake) {
  if (pdfFontsRegistered || !pdfMake) {
    return;
  }

  if (typeof pdfMake.addFonts === "function") {
    pdfMake.addFonts(PDFMAKE_FONTS);
  } else {
    pdfMake.fonts = {
      ...(pdfMake.fonts || {}),
      ...PDFMAKE_FONTS,
    };
  }

  pdfFontsRegistered = true;
}

function isSupabaseSourceLoaded() {
  return state.sourceLabel === "Supabase material_catalog";
}

function formatDateTime(value) {
  if (!value) {
    return "recently";
  }

  try {
    return dateTimeFormatter.format(new Date(value));
  } catch (_error) {
    return value;
  }
}

function sanitizeOptionalText(value) {
  const trimmed = value?.trim();
  return trimmed ? trimmed : null;
}

function sanitizeOptionalDate(value) {
  const trimmed = value?.trim();
  return trimmed ? trimmed : null;
}

function syncQuoteMetaFromInputs() {
  if (!refs.quoteClientName) {
    return;
  }

  state.quoteMeta.clientName = refs.quoteClientName.value;
  state.quoteMeta.projectName = refs.quoteProjectName.value;
  state.quoteMeta.quoteDate = refs.quoteDate.value;
  state.quoteMeta.projectArchitect = refs.quoteProjectArchitect.value;
  state.quoteMeta.contactNumber = refs.quoteContactNumber.value;
  state.quoteMeta.emailAddress = refs.quoteEmailAddress.value;
  state.quoteMeta.notes = refs.quoteNotes.value;
  state.quoteMeta.deliveryAmount = refs.deliveryAmount?.value || state.quoteMeta.deliveryAmount || "";
  state.quoteMeta.deliveryIsFree = Boolean(refs.deliveryFree?.checked);
  state.quoteMeta.installSteamAmount =
    refs.installSteamAmount?.value || state.quoteMeta.installSteamAmount || "";
  state.quoteMeta.installSteamIsFree = Boolean(refs.installSteamFree?.checked);
  saveState();
}

function buildContractPreviewData() {
  const lineItems = state.measurementRows
    .map((row) => {
      const cost = getMeasurementCost(row);
      if (cost === null) {
        return null;
      }

      const motorized = isMotorizedMaterialRow(row);
      return {
        room: row.room?.trim() || "",
        label: row.label?.trim() || "",
        type: row.type?.trim() || "-",
        materialCode: row.materialCode?.trim() || "-",
        width: motorized ? "—" : formatMeasurementDimension(row.width),
        height: motorized ? "—" : formatMeasurementDimension(row.height),
        srp: cost,
        isFree: Boolean(row.isFree),
      };
    })
    .filter(Boolean);

  const { subtotal, discountAmount, finalTotal, half } = getSummaryTotals();
  const deliveryAmount = parseCurrencyLikeNumber(state.quoteMeta.deliveryAmount) || 0;
  const installSteamAmount = parseCurrencyLikeNumber(state.quoteMeta.installSteamAmount) || 0;
  const discountValueRaw = String(state.discountValue ?? "").trim();
  const discountValueNumeric = parseCurrencyLikeNumber(state.discountValue) || 0;
  const hasDiscount = discountValueRaw !== "" && discountValueNumeric > 0;
  const discountType = state.discountType === "percent" ? "percent" : "amount";
  const discountPercent = discountType === "percent" ? discountValueNumeric : 0;

  return {
    clientName: state.quoteMeta.clientName.trim() || "-",
    clientAddress: state.quoteMeta.projectName.trim() || "-",
    quoteDate: formatContractDate(state.quoteMeta.quoteDate),
    projectArchitect: state.quoteMeta.projectArchitect.trim() || "-",
    contactNumber: state.quoteMeta.contactNumber.trim() || "-",
    emailAddress: state.quoteMeta.emailAddress.trim() || "-",
    notes: state.quoteMeta.notes.trim(),
    deliveryAmount,
    deliveryIsFree: Boolean(state.quoteMeta.deliveryIsFree),
    installSteamAmount,
    installSteamIsFree: Boolean(state.quoteMeta.installSteamIsFree),
    lineItems,
    subtotal,
    hasDiscount,
    discountAmount,
    discountType,
    discountPercent,
    finalTotal,
    downpayment: half,
    remainingBalance: half,
  };
}

function loadContractPdfAssets() {
  if (contractPdfAssetsPromise) {
    return contractPdfAssetsPromise;
  }

  const luxeLogoUrl = new URL("./assets/luxeshade-logo.png", window.location.href).href;
  const ndsLogoUrl = new URL("./assets/nds-trading-logo.png", window.location.href).href;
  const kkLogoUrl = new URL("./assets/kurtina-kultura-logo.png", window.location.href).href;
  contractPdfAssetsPromise = Promise.all([
    fileUrlToDataUrl(luxeLogoUrl),
    fileUrlToDataUrl(ndsLogoUrl),
    fileUrlToDataUrl(kkLogoUrl),
  ]).then(([luxeLogoDataUrl, ndsLogoDataUrl, kkLogoDataUrl]) => ({
    luxeLogoDataUrl,
    ndsLogoDataUrl,
    kkLogoDataUrl,
  }));
  return contractPdfAssetsPromise;
}

async function fileUrlToDataUrl(url) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`Failed to load asset: ${url}`);
  }

  const blob = await response.blob();
  return await new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result);
    reader.onerror = () => reject(reader.error || new Error("Could not read asset."));
    reader.readAsDataURL(blob);
  });
}

function getPdfBlob(pdfDocument) {
  try {
    const result = pdfDocument.getBlob();
    if (result && typeof result.then === "function") {
      return result;
    }
  } catch (_error) {
    // Fall through to callback-style support for older builds.
  }

  return new Promise((resolve, reject) => {
    try {
      pdfDocument.getBlob(resolve);
    } catch (error) {
      reject(error);
    }
  });
}

function openPdfPreviewWindow(previewWindow, pdfUrl, fileName) {
  if (!previewWindow || previewWindow.closed) {
    return false;
  }

  const previewTitle = escapeHtml(fileName);
  const viewerUrl = `${pdfUrl}#toolbar=0&navpanes=0&scrollbar=1&view=FitH`;
  previewWindow.document.open();
  previewWindow.document.write(`<!doctype html>
<html lang="en">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>${previewTitle}</title>
    <style>
      :root {
        color-scheme: light;
      }

      * {
        box-sizing: border-box;
      }

      html, body {
        margin: 0;
        height: 100%;
        background: #efe8df;
      }

      body {
        display: flex;
        flex-direction: column;
        font-family: "Segoe UI", Arial, sans-serif;
        color: #2f2a26;
      }

      .pdf-preview-bar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        gap: 16px;
        padding: 12px 16px;
        border-bottom: 1px solid #dccab8;
        background: #fbf7f2;
      }

      .pdf-preview-name {
        font-size: 14px;
        font-weight: 600;
        overflow: hidden;
        text-overflow: ellipsis;
        white-space: nowrap;
      }

      .pdf-preview-actions {
        display: flex;
        align-items: center;
        gap: 12px;
        flex-shrink: 0;
      }

      .pdf-preview-note {
        font-size: 12px;
        color: #6f6257;
      }

      .pdf-preview-download {
        display: inline-flex;
        align-items: center;
        justify-content: center;
        min-height: 38px;
        padding: 0 14px;
        border-radius: 999px;
        border: 1px solid #9e7149;
        background: #9e7149;
        color: #ffffff;
        text-decoration: none;
        font-size: 13px;
        font-weight: 600;
      }

      .pdf-preview-download:hover {
        background: #875f3b;
        border-color: #875f3b;
      }

      .pdf-preview-frame {
        flex: 1;
        width: 100%;
        border: 0;
        background: #5f5a55;
      }

      @media (max-width: 720px) {
        .pdf-preview-bar {
          flex-direction: column;
          align-items: stretch;
        }

        .pdf-preview-actions {
          width: 100%;
          justify-content: space-between;
        }

        .pdf-preview-note {
          display: none;
        }
      }
    </style>
  </head>
  <body>
    <div class="pdf-preview-bar">
      <div class="pdf-preview-name">${previewTitle}</div>
      <div class="pdf-preview-actions">
        <span class="pdf-preview-note">Use Download PDF to save with the correct filename.</span>
        <a class="pdf-preview-download" data-pdf-download>Download PDF</a>
      </div>
    </div>
    <iframe class="pdf-preview-frame" data-pdf-viewer title="${previewTitle}"></iframe>
  </body>
</html>`);
  previewWindow.document.close();

  const downloadLink = previewWindow.document.querySelector("[data-pdf-download]");
  if (downloadLink) {
    downloadLink.setAttribute("href", pdfUrl);
    downloadLink.setAttribute("download", fileName);
  }

  const viewerFrame = previewWindow.document.querySelector("[data-pdf-viewer]");
  if (viewerFrame) {
    viewerFrame.setAttribute("src", viewerUrl);
  }

  return true;
}

function schedulePdfUrlCleanup(pdfUrl, previewWindow) {
  let revoked = false;
  const revokeUrl = () => {
    if (revoked) {
      return;
    }

    revoked = true;
    URL.revokeObjectURL(pdfUrl);
  };

  if (previewWindow && !previewWindow.closed) {
    previewWindow.addEventListener("beforeunload", revokeUrl, { once: true });
    window.setTimeout(revokeUrl, 30 * 60 * 1000);
    return;
  }

  window.setTimeout(revokeUrl, 60000);
}

function triggerPdfDownload(pdfUrl, fileName) {
  const link = document.createElement("a");
  link.href = pdfUrl;
  link.download = fileName;
  document.body.append(link);
  link.click();
  link.remove();
}

function buildContractPdfFileName(contract, organization = "luxe") {
  const clientName = contract.clientName.replace(/\s+/g, " ").trim();
  const prefix = organization === "nds"
    ? "NDS Contract"
    : organization === "kk"
      ? "KK Contract"
      : "LuxeShade Contract";
  const rawName = clientName ? `${prefix} - ${clientName}` : prefix;

  const safeName = rawName
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  return `${safeName || prefix}.pdf`;
}

function getContractPdfBranding(organization, assets) {
  if (organization === "kk") {
    const kkAspectRatio = 585 / 1024;
    const watermarkWidth = 250;
    return {
      logoDataUrl: assets.kkLogoDataUrl,
      watermarkWidth,
      watermarkHeight: Math.round(watermarkWidth * kkAspectRatio),
      watermarkOpacity: 0.12,
      borderColor: "#d9d2cd",
      accentColor: "#111111",
      textColor: "#1a1a1a",
      lightFill: "#f3f1ef",
      accentFill: "#eae6e2",
    };
  }
  if (organization === "nds") {
    const ndsAspectRatio = 125 / 425;
    const watermarkWidth = 260;
    return {
      logoDataUrl: assets.ndsLogoDataUrl,
      watermarkWidth,
      watermarkHeight: Math.round(watermarkWidth * ndsAspectRatio),
      watermarkOpacity: 0.14,
      borderColor: "#e4c7c6",
      accentColor: "#8d2a24",
      textColor: "#2b2422",
      lightFill: "#f7eaea",
      accentFill: "#f2dada",
    };
  }

  const luxeAspectRatio = 872 / 1024;
  const watermarkWidth = 220;
  return {
    logoDataUrl: assets.luxeLogoDataUrl,
    watermarkWidth,
    watermarkHeight: Math.round(watermarkWidth * luxeAspectRatio),
    watermarkOpacity: 0.16,
    borderColor: "#e6d8ca",
    accentColor: "#9e7149",
    textColor: "#2f2a26",
    lightFill: "#f5ece2",
    accentFill: "#f3e5d7",
  };
}

function buildContractPdfDefinition(contract, assets, { organization = "luxe" } = {}) {
  const pageMargin = 43.2;
  const orderTableWidths = [92, 76, 112, 63, 63, 76];
  const totalsTableWidths = [340, 146];
  const branding = getContractPdfBranding(organization, assets);
  const { borderColor, accentColor, textColor, lightFill, accentFill } = branding;
  const orderTableLayout = {
    hLineColor: () => borderColor,
    vLineColor: () => borderColor,
    hLineWidth: () => 0.75,
    vLineWidth: () => 0.75,
    paddingLeft: () => 4,
    paddingRight: () => 4,
    paddingTop: () => 6,
    paddingBottom: () => 6,
  };
  const totalsTableLayout = {
    hLineColor: () => borderColor,
    vLineColor: () => borderColor,
    hLineWidth: () => 0.75,
    vLineWidth: () => 0.75,
    paddingLeft: () => 10,
    paddingRight: () => 10,
    paddingTop: () => 10,
    paddingBottom: () => 10,
  };
  const formFieldLayout = {
    hLineWidth: (rowIndex, node) => (rowIndex === 0 ? 0 : 0.75),
    vLineWidth: () => 0,
    hLineColor: () => textColor,
    paddingLeft: () => 0,
    paddingRight: () => 0,
    paddingTop: () => 0,
    paddingBottom: () => 6,
  };
  const termsHeadingLayout = {
    hLineWidth: (rowIndex, node) => (rowIndex === node.table.body.length ? 1.1 : 0),
    vLineWidth: () => 0,
    hLineColor: () => accentColor,
    paddingLeft: () => 0,
    paddingRight: () => 0,
    paddingTop: () => 0,
    paddingBottom: () => 6,
  };
  const termsSectionLayout = {
    hLineWidth: (rowIndex, node) => (rowIndex === node.table.body.length ? 0.9 : 0),
    vLineWidth: () => 0,
    hLineColor: () => accentColor,
    paddingLeft: () => 10,
    paddingRight: () => 10,
    paddingTop: () => 5,
    paddingBottom: () => 5,
  };

  const buildField = (label, value) => ({
    table: {
      widths: [100, "*"],
      body: [[
        { text: label, color: textColor, margin: [0, 4, 8, 0] },
        { text: value, color: textColor, margin: [0, 0, 0, 0] },
      ]],
    },
    layout: formFieldLayout,
    margin: [0, 0, 0, 10],
  });

  const buildTermsEntries = (entries) =>
    entries.flatMap((entry, index) => [
      {
        text: `${index + 1}. ${entry.title}`,
        bold: true,
        margin: [0, index === 0 ? 0 : 6, 0, 4],
      },
      ...entry.items.map((item) => ({
        text: `○ ${item}`,
        margin: [14, 0, 0, 3],
      })),
    ]);

  const buildTermsSection = (title, entries, { topMargin = 0 } = {}) => ({
    stack: [
      {
        text: title.toUpperCase(),
        bold: true,
        fontSize: 10,
        characterSpacing: 0.5,
        margin: [0, topMargin, 0, 8],
      },
      ...buildTermsEntries(entries),
    ],
    margin: [0, 0, 0, 8],
  });

  const buildStyledTermsDocumentHeading = (title) => ({
    table: {
      widths: ["*"],
      body: [[{
        text: title.toUpperCase(),
        bold: true,
        fontSize: 13,
        characterSpacing: 0.35,
        color: textColor,
      }]],
    },
    layout: termsHeadingLayout,
    margin: [0, 0, 0, 10],
  });

  const buildStyledTermsEntries = (entries) =>
    entries.flatMap((entry, index) => [
      {
        columns: [
          {
            width: 16,
            text: `${index + 1}.`,
            bold: true,
            fontSize: 10.6,
            color: accentColor,
          },
          {
            width: "*",
            text: entry.title,
            bold: true,
            ...(entry.titleFont ? { font: entry.titleFont } : {}),
            fontSize: entry.titleFontSize || 10.7,
            color: textColor,
            lineHeight: 1.12,
          },
        ],
        columnGap: 4,
        margin: [0, index === 0 ? 0 : 8, 0, 4],
      },
      ...entry.items.map((item) => ({
        margin: [20, 0, 0, 4],
        ...(typeof item === "string" ? { text: item } : item),
      })),
    ]);

  const buildStyledTermsSection = (title, entries, { topMargin = 0 } = {}) => ({
    stack: [
      {
        table: {
          widths: ["*"],
          body: [[{
            text: title.toUpperCase(),
            bold: true,
            fontSize: 11,
            characterSpacing: 0.2,
            color: textColor,
            fillColor: lightFill,
          }]],
        },
        layout: termsSectionLayout,
        margin: [0, topMargin, 0, 10],
      },
      ...buildStyledTermsEntries(entries),
    ],
    margin: [0, 0, 0, 10],
  });

  return {
    pageSize: "LETTER",
    pageMargins: [pageMargin, pageMargin, pageMargin, pageMargin],
    background(currentPage, pageSize) {
      const safePageWidth = pageSize?.width || 612;
      const safePageHeight = pageSize?.height || 792;
      return {
        image: branding.logoDataUrl,
        width: branding.watermarkWidth,
        height: branding.watermarkHeight,
        opacity: branding.watermarkOpacity,
        absolutePosition: {
          x: (safePageWidth - branding.watermarkWidth) / 2,
          y: (safePageHeight - branding.watermarkHeight) / 2,
        },
      };
    },
    footer(currentPage, pageCount) {
      return {
        text: `Page ${currentPage} of ${pageCount}`,
        alignment: "right",
        color: accentColor,
        fontSize: 9,
        margin: [pageMargin, 12, pageMargin, 22],
      };
    },
    defaultStyle: {
      font: "TenorSans",
      fontSize: 10,
      color: textColor,
      lineHeight: 1.25,
    },
    content: [
      {
        stack: [
          {
            image: branding.logoDataUrl,
            width: 110,
            alignment: "center",
            margin: [0, 0, 0, 10],
          },
          {
            text: "CLIENT INFORMATION",
            alignment: "center",
            fontSize: 12,
            characterSpacing: 1.2,
            margin: [0, 0, 0, 22],
          },
          {
            columns: [
              {
                width: "*",
                stack: [
                  buildField("Date", contract.quoteDate),
                  buildField("Client's Name", contract.clientName),
                  buildField("Contact No.", contract.contactNumber),
                ],
              },
              {
                width: 18,
                text: "",
              },
              {
                width: "*",
                stack: [
                  buildField("Client's Address", contract.clientAddress),
                  buildField("Project Architect", contract.projectArchitect),
                  buildField("Email Address", contract.emailAddress),
                ],
              },
            ],
            columnGap: 18,
          },
          ...(contract.notes
            ? [
                {
                  stack: [
                    { text: "Notes", margin: [0, 0, 0, 4] },
                    {
                      text: contract.notes,
                      margin: [0, 0, 0, 0],
                    },
                  ],
                  margin: [0, 4, 0, 0],
                },
              ]
            : []),
          {
            text: "ORDER DETAILS",
            fontSize: 10,
            bold: true,
            characterSpacing: 1.2,
            color: accentColor,
            margin: [0, 26, 0, 10],
          },
          {
            table: {
              headerRows: 1,
              dontBreakRows: true,
              keepWithHeaderRows: 1,
              widths: orderTableWidths,
              body: buildContractPdfOrderTableBody(contract.lineItems, {
                headerFill: lightFill,
              }),
            },
            layout: orderTableLayout,
          },
          {
            stack: [
              {
                table: {
                  widths: totalsTableWidths,
                  body: (() => {
                    const rows = [
                      [
                        { text: "Sub Total", bold: true },
                        {
                          text: formatPdfPesoAmount(contract.subtotal),
                          font: "Roboto",
                          alignment: "right",
                        },
                      ],
                    ];

                    const shouldShowDelivery =
                      Boolean(contract.deliveryIsFree) || Number(contract.deliveryAmount) > 0;
                    if (shouldShowDelivery) {
                      rows.push([
                        { text: "Delivery", bold: true },
                        {
                          text: formatPdfAdditionalChargeDisplay(
                            contract.deliveryAmount,
                            contract.deliveryIsFree,
                          ),
                          font: "Roboto",
                          alignment: "right",
                        },
                      ]);
                    }

                    const shouldShowInstallSteam =
                      Boolean(contract.installSteamIsFree) || Number(contract.installSteamAmount) > 0;
                    if (shouldShowInstallSteam) {
                      rows.push([
                        { text: "Installation and Steam", bold: true },
                        {
                          text: formatPdfAdditionalChargeDisplay(
                            contract.installSteamAmount,
                            contract.installSteamIsFree,
                          ),
                          font: "Roboto",
                          alignment: "right",
                        },
                      ]);
                    }

                    if (contract.hasDiscount) {
                      const discountSuffix =
                        contract.discountType === "percent" && Number(contract.discountPercent) > 0
                          ? ` (${Number(contract.discountPercent)}%)`
                          : "";
                      rows.push([
                        { text: "Discount", bold: true },
                        {
                          text: `${formatPdfPesoAmount(contract.discountAmount)}${discountSuffix}`,
                          font: "Roboto",
                          alignment: "right",
                        },
                      ]);
                    }

                    rows.push([
                      { text: "Total", bold: true, fontSize: 14, fillColor: accentFill },
                      {
                        text: formatPdfPesoAmount(contract.finalTotal),
                        font: "Roboto",
                        bold: true,
                        fontSize: 14,
                        alignment: "right",
                        fillColor: accentFill,
                      },
                    ]);

                    return rows;
                  })(),
                },
                layout: totalsTableLayout,
              },
            ],
            unbreakable: true,
            margin: [0, 14, 0, 0],
            pageBreak: "after",
          },
        ],
      },
      {
        stack: [
          buildStyledTermsDocumentHeading("Terms of Contract"),
          buildStyledTermsSection("Section 1. Payment Schedule and Conditions", [
            {
              title: "Down Payment (50%)",
              items: [
                {
                  text: [
                    { text: "Amount: " },
                    { text: formatPdfPesoAmount(contract.downpayment), bold: true },
                  ],
                  font: "Roboto",
                  fontSize: 11.5,
                },
                "Payable upon signing or approval of the contract.",
              ],
            },
            {
              title: "Final Payment (50%)",
              items: [
                {
                  text: [
                    { text: "Amount: " },
                    { text: formatPdfPesoAmount(contract.remainingBalance), bold: true },
                  ],
                  font: "Roboto",
                  fontSize: 11.5,
                },
                "Payable on the same day of installation or within 24 hours upon completion of work.",
              ],
            },
            {
              title: "Payment Options",
              items: [
                {
                  text: [
                    { text: "Online payment via " },
                    { text: "BDO", bold: true },
                    { text: " is accepted." },
                  ],
                  font: "Roboto",
                },
                {
                  text: [
                    { text: "Account name: " },
                    { text: "Ma. Elena Bernardo", font: "Roboto", bold: true, fontSize: 11.5 },
                  ],
                },
                {
                  text: [
                    { text: "Account number: " },
                    { text: "0110 1002 1573", font: "Roboto", bold: true, fontSize: 11.5 },
                  ],
                },
                {
                  text: [
                    { text: "Online payment via " },
                    { text: "Metrobank", bold: true },
                    { text: " is accepted." },
                  ],
                  font: "Roboto",
                },
                {
                  text: [
                    { text: "Account name: " },
                    {
                      text: "Jan Robert A Gregorio",
                      font: "Roboto",
                      bold: true,
                      fontSize: 11.5,
                    },
                  ],
                },
                {
                  text: [
                    { text: "Account number: " },
                    { text: "494 349 462 2220", font: "Roboto", bold: true, fontSize: 11.5 },
                  ],
                },
                {
                  text: [
                    { text: "Other payment methods such as " },
                    { text: "checks or cash", bold: true },
                    { text: " should be coordinated with the supplier as needed." },
                  ],
                  font: "Roboto",
                },
                {
                  text: [
                    { text: "For check payments, " },
                    { text: "clearance must be confirmed", bold: true },
                    { text: " before work or deliveries commence." },
                  ],
                  font: "Roboto",
                },
              ],
            },
          ]),
          buildStyledTermsSection("Section 2. Work and Delivery Lead Time", [
            {
              title: "Fabrication and Delivery",
              items: [
                "Typically 12 to 14 working days, if fabric is available, from the date of down payment or from the clearing date for check payments.",
                "This lead time covers the production and preparation of materials.",
              ],
            },
            {
              title: "Approximate Completion Time On-Site",
              items: [
                "Generally 1 to 3 working days for small projects and 3 to 7 working days for big projects, depending on scope.",
                "Actual time may vary due to building administration requirements, power scheduling, access limitations, or unforeseen events.",
              ],
            },
            {
              title: "Cooperation with Building or Project Management",
              items: [
                "The customer acknowledges that certain installation activities may require coordination with building management, such as permits, elevator scheduling, and access arrangements.",
                "Delays caused by building administration rules, weather, power interruptions, or similar external factors are beyond the supplier's control.",
              ],
            },
          ], { topMargin: 8 }),
        ],
      },
      {
        stack: [
          buildStyledTermsSection("Section 3. Installation Work", [
            {
              title: "Coverage of Installation",
              items: [
                "The installation includes fabrication and assembly of the contracted items and proper placement in the customer's designated area.",
                "The supplier will only be responsible for connecting and installing components directly related to the agreed work.",
              ],
            },
            {
              title: "Exclusions",
              items: [
                "Electrical wiring, water lines, sanitary connections, or structural modifications not stated in the contract are excluded unless otherwise agreed in writing.",
                "Plumbing or electrical work requiring specialized trades remains the responsibility of the customer or a separately engaged licensed professional.",
              ],
            },
            {
              title: "Building Permits and Approvals",
              items: [
                "The supplier shall secure the necessary building permits and condo work permits required for installation.",
                "Any associated permit fees may be charged to the customer if not included in the initial quotation.",
              ],
            },
            {
              title: "Liability",
              items: [
                "The supplier will not be held liable for damage or incidents caused by existing building conditions such as defective pipes, concealed wiring, or structural issues.",
                "Damage resulting from negligence or mishandling by supplier personnel will be addressed in accordance with applicable local laws or as otherwise agreed by both parties.",
                "Fragile or delicate items should be removed or secured by the customer before installation to avoid damage.",
              ],
            },
          ]),
        ],
      },
      {
        pageBreak: "before",
        stack: [
          buildStyledTermsSection("Section 4. Customer Cancellation of Order", [
            {
              title: "No Refund Policy",
              items: [
                "Once a down payment has been made, it is strictly non-refundable under any circumstances after contract signing.",
              ],
            },
          ]),
          {
            stack: [
              {
                text: [{ text: "Dated: ", bold: true }, contract.quoteDate],
                margin: [0, 88, 0, 24],
              },
              {
                columns: [
                  {
                    width: "*",
                    stack: [
                      { text: "", margin: [0, 0, 0, 26] },
                      { canvas: [{ type: "line", x1: 0, y1: 0, x2: 220, y2: 0, lineWidth: 1, lineColor: textColor }] },
                      { text: "Company Representative", margin: [0, 6, 0, 0] },
                    ],
                  },
                  {
                    width: 24,
                    text: "",
                  },
                  {
                    width: "*",
                    stack: [
                      { text: "", margin: [0, 0, 0, 26] },
                      { canvas: [{ type: "line", x1: 0, y1: 0, x2: 220, y2: 0, lineWidth: 1, lineColor: textColor }] },
                      { text: "Client Signature", margin: [0, 6, 0, 0] },
                    ],
                  },
                ],
              },
            ],
            unbreakable: true,
          },
        ],
      },
    ],
  };
}

function buildContractPdfOrderTableBody(lineItems, { headerFill }) {
  const buildHeaderCell = (text) => ({
    text,
    font: "Roboto",
    bold: true,
    fillColor: headerFill,
    alignment: "center",
    verticalAlignment: "middle",
    fontSize: 9.4,
    color: "#000000",
    characterSpacing: 0.2,
  });

  const body = [[
    buildHeaderCell("AREA"),
    buildHeaderCell("TYPE"),
    buildHeaderCell("MATERIAL CODE"),
    buildHeaderCell("WIDTH"),
    buildHeaderCell("HEIGHT"),
    buildHeaderCell("SRP"),
  ]];

  let previousRoom = "";
  lineItems.forEach((item) => {
    if (item.room && item.room !== previousRoom) {
      body.push([
        {
          text: item.room,
          colSpan: 6,
          font: "Roboto",
          bold: true,
          color: "#000000",
          alignment: "left",
          verticalAlignment: "middle",
          fontSize: 10,
          characterSpacing: 0.15,
          margin: [6, 1, 0, 1],
        },
        {},
        {},
        {},
        {},
        {},
      ]);
    }

    previousRoom = item.room || previousRoom;
    body.push([
      { text: item.label || " ", alignment: "center", verticalAlignment: "middle", fontSize: 9 },
      { text: item.type, alignment: "center", verticalAlignment: "middle", fontSize: 9 },
      { text: item.materialCode, alignment: "center", verticalAlignment: "middle", fontSize: 9 },
      { text: item.width, alignment: "center", verticalAlignment: "middle", fontSize: 9 },
      { text: item.height, alignment: "center", verticalAlignment: "middle", fontSize: 9 },
      {
        text: formatPdfAdditionalChargeDisplay(item.srp, item.isFree),
        font: "Roboto",
        alignment: "center",
        verticalAlignment: "middle",
        fontSize: 9,
      },
    ]);
  });

  return body;
}

function formatContractDate(value) {
  const fallbackDate = new Date();

  if (!value) {
    return contractDateFormatter.format(fallbackDate);
  }

  const parsedDate = new Date(`${value}T00:00:00`);
  if (Number.isNaN(parsedDate.getTime())) {
    return contractDateFormatter.format(fallbackDate);
  }

  return contractDateFormatter.format(parsedDate);
}

function formatMeasurementDimension(value) {
  const parsedValue = parseCurrencyLikeNumber(value);
  if (parsedValue === null) {
    return "-";
  }

  return Number.isInteger(parsedValue)
    ? `${parsedValue.toLocaleString("en-PH")} mm`
    : `${parsedValue.toLocaleString("en-PH", { maximumFractionDigits: 2 })} mm`;
}

function escapeHtml(value) {
  return String(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

const MEASUREMENT_FIT_MIN_FONT_PX = 11;
/** Reserve space for the native dropdown arrow on `<select>` (see fit logic). */
const MEASUREMENT_FIT_SELECT_ARROW_RESERVE_PX = 28;

let measurementFitMeasureCanvas;

function getMeasurementFitMeasureContext() {
  if (!measurementFitMeasureCanvas) {
    measurementFitMeasureCanvas = document.createElement("canvas");
  }
  return measurementFitMeasureCanvas.getContext("2d");
}

function measureMeasurementFieldTextWidth(cs, text, fontSizePx) {
  const ctx = getMeasurementFitMeasureContext();
  if (!ctx) {
    return 0;
  }
  const family = cs.fontFamily || "sans-serif";
  const weight = cs.fontWeight || "400";
  const style = cs.fontStyle || "normal";
  ctx.font = `${style} ${weight} ${fontSizePx}px ${family}`;
  return ctx.measureText(text).width;
}

function getSelectMeasurementTextMaxPx(el) {
  const cs = window.getComputedStyle(el);
  const pl = parseFloat(cs.paddingLeft) || 0;
  const pr = parseFloat(cs.paddingRight) || 0;
  const inner = el.clientWidth - pl - pr;
  return Math.max(0, inner - MEASUREMENT_FIT_SELECT_ARROW_RESERVE_PX);
}

function fitMeasurementFieldToCell(el) {
  if (!el?.classList?.contains("measurement-fit-field")) {
    return;
  }

  if (!el.dataset.measurementFitFontBasePx) {
    el.style.removeProperty("font-size");
    const px = parseFloat(window.getComputedStyle(el).fontSize);
    el.dataset.measurementFitFontBasePx = String(
      Number.isFinite(px) && px > 0 ? px : 16,
    );
  }

  const basePx = Number(el.dataset.measurementFitFontBasePx) || 16;
  el.style.fontSize = `${basePx}px`;
  let size = basePx;

  if (el.tagName === "SELECT") {
    const text = el.options[el.selectedIndex]?.text ?? "";
    const cs = window.getComputedStyle(el);
    const maxW = getSelectMeasurementTextMaxPx(el);
    if (!text || maxW <= 0) {
      return;
    }
    while (
      size > MEASUREMENT_FIT_MIN_FONT_PX &&
      measureMeasurementFieldTextWidth(cs, text, size) > maxW
    ) {
      size -= 0.5;
      el.style.fontSize = `${size}px`;
    }
    return;
  }

  while (el.scrollWidth > el.clientWidth + 1 && size > MEASUREMENT_FIT_MIN_FONT_PX) {
    size -= 0.5;
    el.style.fontSize = `${size}px`;
  }
}

function attachMeasurementFieldFit(el) {
  const scheduleFit = () => {
    window.requestAnimationFrame(() => fitMeasurementFieldToCell(el));
  };

  const onCellResize = () => {
    delete el.dataset.measurementFitFontBasePx;
    scheduleFit();
  };

  el.addEventListener("input", scheduleFit);
  if (el.tagName === "SELECT") {
    el.addEventListener("change", scheduleFit);
  }

  const td = el.closest("td");
  if (td && typeof ResizeObserver !== "undefined") {
    const resizeObserver = new ResizeObserver(onCellResize);
    resizeObserver.observe(td);
  }

  scheduleFit();
}

function buildTextInput({ value, placeholder, disabled = false, onInput, className }) {
  const input = document.createElement("input");
  input.type = "text";
  input.value = value;
  input.placeholder = placeholder;
  input.disabled = disabled;
  if (className) {
    input.className = className;
  }
  input.addEventListener("input", (event) => onInput(event.target.value));
  return input;
}

function buildNumberInput({
  value,
  placeholder,
  disabled = false,
  onInput,
  className,
}) {
  const input = document.createElement("input");
  input.type = "number";
  input.min = "0";
  input.step = "0.01";
  input.inputMode = "decimal";
  input.value = value;
  input.placeholder = placeholder;
  input.disabled = disabled;
  if (className) {
    input.className = className;
  }
  input.addEventListener("input", (event) =>
    onInput(normalizeInputNumber(event.target.value)),
  );
  return input;
}

function buildEmptyBlock(copy) {
  const block = document.createElement("div");
  block.className = "empty-state-cell";
  block.textContent = copy;
  return block;
}

function normalizeInputNumber(value) {
  if (value === "" || value === null || value === undefined) {
    return "";
  }

  return value;
}

function getDiscountInputDisplayValue(value) {
  const parsedValue = parseCurrencyLikeNumber(value);
  return parsedValue === null || parsedValue === 0 ? "" : String(value);
}

function parseCurrencyLikeNumber(value) {
  if (value === "" || value === null || value === undefined) {
    return null;
  }

  const sanitized = String(value).replace(/,/g, "").trim();
  if (!sanitized) {
    return null;
  }

  const parsed = Number(sanitized);
  return Number.isFinite(parsed) ? parsed : null;
}

function formatCurrency(value) {
  return currencyFormatter.format(value || 0);
}

/** Peso sign (U+20B1) + formatted amount. Use with font: "Roboto" in pdfmake when default font is TenorSans — Tenor Sans may not embed ₱. */
function formatPdfPesoAmount(value) {
  const numericValue = Number(value);
  const safeValue = Number.isFinite(numericValue) ? numericValue : 0;
  return `\u20b1${safeValue.toLocaleString("en-PH", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
}

/** Shows list price; when marked free, appends (FREE) for transparency while subtotal still treats it as zero. */
function formatPdfAdditionalChargeDisplay(amount, isFree) {
  const line = formatPdfPesoAmount(amount);
  return isFree ? `${line} (FREE)` : line;
}

function createEmptyStateRow(colspan, copy) {
  const tr = document.createElement("tr");
  const td = document.createElement("td");
  td.colSpan = colspan;
  td.className = "empty-state-cell";
  td.textContent = copy;
  tr.append(td);
  return tr;
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    const nextCharacter = text[index + 1];

    if (character === '"') {
      if (inQuotes && nextCharacter === '"') {
        cell += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (character === "," && !inQuotes) {
      row.push(cell);
      cell = "";
      continue;
    }

    if ((character === "\n" || character === "\r") && !inQuotes) {
      if (character === "\r" && nextCharacter === "\n") {
        index += 1;
      }

      row.push(cell);
      if (row.some((entry) => entry !== "")) {
        rows.push(row);
      }
      row = [];
      cell = "";
      continue;
    }

    cell += character;
  }

  if (cell !== "" || row.length > 0) {
    row.push(cell);
    if (row.some((entry) => entry !== "")) {
      rows.push(row);
    }
  }

  return rows;
}
