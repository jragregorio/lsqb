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
  "Sheer Curtain",
  "Blackout Curtain",
  "Semi-Blackout Curtain",
  "Soft B/O Curtain",
  "Combi Blinds",
  "Roller Blinds",
  "Wood Blinds",
  "Silhouette Blinds",
  "Vertical Blinds",
  "Sunscreen",
  "Roman Blinds",
];

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
  quoteClientName: document.querySelector("#quote-client-name"),
  quoteProjectName: document.querySelector("#quote-project-name"),
  quoteDate: document.querySelector("#quote-date"),
  quoteProjectArchitect: document.querySelector("#quote-project-architect"),
  quoteContactNumber: document.querySelector("#quote-contact-number"),
  quoteEmailAddress: document.querySelector("#quote-email-address"),
  quoteNotes: document.querySelector("#quote-notes"),
  newQuoteBtn: document.querySelector("#new-quote-btn"),
  saveQuoteBtn: document.querySelector("#save-quote-btn"),
  exportPdfBtn: document.querySelector("#export-pdf-btn"),
  refreshQuotesBtn: document.querySelector("#refresh-quotes-btn"),
  currentQuoteLabel: document.querySelector("#current-quote-label"),
  quoteStatus: document.querySelector("#quote-status"),
  savedQuotesCount: document.querySelector("#saved-quotes-count"),
  savedQuotesList: document.querySelector("#saved-quotes-list"),
  materialPanel: document.querySelector("#material-setup-panel"),
  measurementsPanel: document.querySelector("#measurements-panel"),
  materialBody: document.querySelector("#material-setup-body"),
  measurementBody: document.querySelector("#measurement-body"),
  addMaterialBtn: document.querySelector("#add-material-btn"),
  addMeasurementBtn: document.querySelector("#add-measurement-btn"),
  discountType: document.querySelector("#discount-type"),
  discountValue: document.querySelector("#discount-value"),
  summaryPanel: document.querySelector(".summary-panel"),
};

const state = loadState();
const ACTIVE_QUOTE_BAR_REVEAL_SCROLL_Y = 300;
const AUTOSAVE_DELAY_MS = 1800;

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
};

let contractPdfAssetsPromise = null;
let pdfFontsRegistered = false;

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
  refs.signInBtn.addEventListener("click", handleSignIn);
  refs.signOutBtn.addEventListener("click", handleSignOut);
  refs.syncSupabaseBtn.addEventListener("click", () =>
    loadMaterialsFromSupabase({ showAlertOnFailure: true }),
  );
  refs.newQuoteBtn.addEventListener("click", handleNewQuote);
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
  resetQuoteDraft();
  runtime.quoteWorkspaceActive = false;
  saveState();
  render();
  setQuoteStatus("Quote unloaded. Select a saved quote or start a new one.");
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
    if (!isSupabaseSourceLoaded()) {
      await loadMaterialsFromSupabase({ showAlertOnFailure: false });
    }
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
    quoteMeta: getDefaultQuoteMeta(),
    selectedMaterials: [],
    measurementRows: [],
    discountType: "amount",
    discountValue: "",
  };
}

function saveState() {
  window.localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function persistDraftChange() {
  saveState();
  queueAutosave();
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
    createdAt: savedQuote.created_at || "",
    updatedAt: savedQuote.updated_at || "",
  };
  state.discountType = savedQuote.discount_type === "percent" ? "percent" : "amount";
  state.discountValue =
    getDiscountInputDisplayValue(savedQuote.discount_value) ||
    getDiscountInputDisplayValue(savedQuote.discount);
  runtime.loadedQuoteFingerprint = buildCurrentQuoteFingerprint();
  saveState();

  runtime.quoteBusy = false;
  await refreshSavedQuotes({ showAlertOnFailure: false, silent: true });
  render();
  setQuoteStatus(autosave ? "Quote autosaved to Supabase." : "Quote saved to Supabase.");
  return true;
}

async function loadQuoteById(quoteId) {
  if (!runtime.session) {
    setQuoteStatus("Sign in first before loading saved quotes.", true);
    return;
  }

  clearQueuedAutosave();
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
      .select("id, quote_material_id, room_section, measurement_type, material_code, label, width_mm, height_mm, material_label, asking_price, sort_order")
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
    measurementsResult.data.map((row) => ({
      id: createId("measurement"),
      room: row.room_section || "",
      type: row.measurement_type || "",
      materialCode: row.material_code || "",
      label: row.label || "",
      width: row.width_mm === null ? "" : String(row.width_mm),
      height: row.height_mm === null ? "" : String(row.height_mm),
      materialId: localIdMap.get(row.quote_material_id) || "",
    })) || [];

  state.discountType =
    quoteResult.data.discount_type === "percent" ? "percent" : "amount";
  state.discountValue =
    getDiscountInputDisplayValue(quoteResult.data.discount_value) ||
    getDiscountInputDisplayValue(quoteResult.data.discount);
  ensureStarterRows();
  runtime.expandedQuoteId = quoteId;
  runtime.loadedQuoteFingerprint = buildCurrentQuoteFingerprint();
  runtime.quoteWorkspaceActive = true;
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

  if (!window.confirm(`Delete ${quoteName}? This cannot be undone.`)) {
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
  if (refs.exportPdfBtn) {
    refs.exportPdfBtn.disabled = runtime.quoteBusy || runtime.exportPdfBusy || !isQuoteWorkspaceActive();
    refs.exportPdfBtn.textContent = runtime.exportPdfBusy ? "Preparing PDF..." : "Export PDF";
  }
  refs.refreshQuotesBtn.disabled =
    !signedIn || runtime.quoteBusy || runtime.quoteListBusy;

  renderQuoteStatus();
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

function renderMaterials() {
  refs.materialBody.innerHTML = "";

  if (state.selectedMaterials.length === 0) {
    refs.materialBody.append(createEmptyStateRow(5, "Add a material to begin."));
    return;
  }

  const selectedKeys = state.selectedMaterials
    .map((row) => row.sourceKey)
    .filter(Boolean);

  state.selectedMaterials.forEach((row) => {
    const tr = document.createElement("tr");

    const materialCell = document.createElement("td");
    const select = document.createElement("select");
    const placeholder = new Option("Select material", "");
    select.append(placeholder);

    state.sourceMaterials.forEach((material) => {
      const option = new Option(
        material.division
          ? `${material.category} (${material.division})`
          : material.category,
        material.key,
      );
      option.disabled =
        selectedKeys.includes(material.key) && material.key !== row.sourceKey;
      option.selected = material.key === row.sourceKey;
      select.append(option);
    });

    select.disabled = runtime.quoteBusy;
    select.addEventListener("change", (event) => {
      const nextKey = event.target.value;
      const material = state.sourceMaterials.find((item) => item.key === nextKey);

      row.sourceKey = nextKey;
      row.category = material?.category || "";
      row.division = material?.division || "";
      row.retailPrice = material ? material.retailPrice : "";

      if (!nextKey) {
        row.askingPrice = "";
      }

      sanitizeMeasurementMaterialSelections();
      persistDraftChange();
      render();
    });

    materialCell.append(select);

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

function renderMeasurements() {
  refs.measurementBody.innerHTML = "";

  if (state.measurementRows.length === 0) {
    refs.measurementBody.append(
      createEmptyStateRow(9, "Add a measurement row to begin."),
    );
    return;
  }

  const configuredMaterials = getConfiguredMaterials();
  refs.addMeasurementBtn.disabled =
    configuredMaterials.length === 0 || runtime.quoteBusy;

  state.measurementRows.forEach((row) => {
    const tr = document.createElement("tr");

    const roomCell = document.createElement("td");
    roomCell.append(
      buildTextInput({
        value: row.room,
        placeholder: "Living Room",
        disabled: runtime.quoteBusy,
        onInput: (value) => {
          row.room = value;
          persistDraftChange();
        },
      }),
    );

    const typeCell = document.createElement("td");
    const typeSelect = document.createElement("select");
    typeSelect.append(new Option("Select type", ""));
    MEASUREMENT_TYPE_OPTIONS.forEach((optionLabel) => {
      const option = new Option(optionLabel, optionLabel);
      option.selected = optionLabel === row.type;
      typeSelect.append(option);
    });
    typeSelect.disabled = runtime.quoteBusy;
    typeSelect.addEventListener("change", (event) => {
      row.type = event.target.value;
      persistDraftChange();
    });
    typeCell.append(typeSelect);

    const materialCodeCell = document.createElement("td");
    materialCodeCell.append(
      buildTextInput({
        value: row.materialCode,
        placeholder: "Material code",
        disabled: runtime.quoteBusy,
        onInput: (value) => {
          row.materialCode = value;
          persistDraftChange();
        },
      }),
    );

    const labelCell = document.createElement("td");
    labelCell.append(
      buildTextInput({
        value: row.label,
        placeholder: "W1",
        disabled: runtime.quoteBusy,
        onInput: (value) => {
          row.label = value;
          persistDraftChange();
        },
      }),
    );

    const widthCell = document.createElement("td");
    widthCell.append(
      buildNumberInput({
        value: row.width,
        placeholder: "0",
        disabled: runtime.quoteBusy,
        onInput: (value) => {
          row.width = value;
          persistDraftChange();
          updateMeasurementOutputs();
          renderSummary();
        },
      }),
    );

    const heightCell = document.createElement("td");
    heightCell.append(
      buildNumberInput({
        value: row.height,
        placeholder: "0",
        disabled: runtime.quoteBusy,
        onInput: (value) => {
          row.height = value;
          persistDraftChange();
          updateMeasurementOutputs();
          renderSummary();
        },
      }),
    );

    const materialCell = document.createElement("td");
    const select = document.createElement("select");
    select.append(new Option("Select configured material", ""));
    configuredMaterials.forEach((material) => {
      const option = new Option(
        `${material.category} (${formatCurrency(material.askingPrice)})`,
        material.id,
      );
      option.selected = material.id === row.materialId;
      select.append(option);
    });
    select.disabled = configuredMaterials.length === 0 || runtime.quoteBusy;
    select.addEventListener("change", (event) => {
      row.materialId = event.target.value;
      persistDraftChange();
      updateMeasurementOutputs();
      renderSummary();
    });
    materialCell.append(select);

    const costCell = document.createElement("td");
    costCell.className = "money-cell";
    const updateMeasurementOutputs = () => {
      const cost = getMeasurementCost(row);
      const squareFootage = getMeasurementSquareFootage(row);
      costCell.textContent = cost === null ? "PHP 0.00" : formatCurrency(cost);
      if (squareFootage !== null) {
        const helper = document.createElement("span");
        helper.className = squareFootage.minimumApplied
          ? "muted-helper sqft-warning"
          : "muted-helper";
        helper.textContent = squareFootage.minimumApplied
          ? `${squareFootage.rawRounded} sqft rounded -> billed as ${squareFootage.billed} sqft minimum`
          : `${squareFootage.billed} sqft rounded`;
        costCell.append(helper);
      }
    };
    updateMeasurementOutputs();

    const actionCell = document.createElement("td");
    const removeButton = document.createElement("button");
    removeButton.type = "button";
    removeButton.className = "ghost-danger-button";
    removeButton.textContent = "Remove";
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
    actionCell.append(removeButton);

    tr.append(
      roomCell,
      labelCell,
      typeCell,
      materialCodeCell,
      widthCell,
      heightCell,
      materialCell,
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

function getPocketValue(materialRow) {
  const askingPrice = parseCurrencyLikeNumber(materialRow.askingPrice);
  const retailPrice = parseCurrencyLikeNumber(materialRow.retailPrice);

  if (askingPrice === null || retailPrice === null) {
    return null;
  }

  return askingPrice - retailPrice;
}

function getMeasurementSquareFootage(row) {
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
  const squareFootage = getMeasurementSquareFootage(row);
  const material = state.selectedMaterials.find((item) => item.id === row.materialId);
  const askingPrice = parseCurrencyLikeNumber(material?.askingPrice);

  if (squareFootage === null || askingPrice === null) {
    return null;
  }

  return squareFootage.billed * askingPrice;
}

function getSubtotal() {
  return state.measurementRows.reduce((total, row) => {
    const cost = getMeasurementCost(row);
    return total + (cost ?? 0);
  }, 0);
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
      Boolean(row.materialId);

    if (!hasAnyValue) {
      continue;
    }

    const width = parseCurrencyLikeNumber(row.width);
    const height = parseCurrencyLikeNumber(row.height);
    const selectedMaterial = state.selectedMaterials.find(
      (material) => material.id === row.materialId,
    );
    const askingPrice = parseCurrencyLikeNumber(selectedMaterial?.askingPrice);

    if (width === null || height === null || !selectedMaterial || askingPrice === null) {
      return {
        ok: false,
        message:
          "Every measurement row must have width, height, and a configured material before saving.",
      };
    }

    if (!validateOnly) {
      drafts.push({
        localMaterialId: selectedMaterial.id,
        room: row.room,
        type: row.type,
        materialCode: row.materialCode,
        label: row.label.trim(),
        width,
        height,
        materialLabel: selectedMaterial.category,
        askingPrice,
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
}

function clearQueuedAutosave() {
  if (!runtime.autosaveTimer) {
    return;
  }

  window.clearTimeout(runtime.autosaveTimer);
  runtime.autosaveTimer = 0;
}

function canAutosaveCurrentQuote() {
  return Boolean(runtime.session && state.quoteMeta.id) &&
    !runtime.quoteBusy &&
    hasLoadedQuoteUnsavedChanges();
}

function queueAutosave() {
  clearQueuedAutosave();

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
          row.materialId,
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
  setQuoteStatus("Preparing contract PDF...");

  const popupWindow = window.open("", "_blank");

  try {
    const contract = buildContractPreviewData();
    const assets = await loadContractPdfAssets();
    const documentDefinition = buildContractPdfDefinition(contract, assets);
    const pdfBlob = await getPdfBlob(pdfMake.createPdf(documentDefinition));
    const pdfUrl = URL.createObjectURL(pdfBlob);
    const fileName = buildContractPdfFileName(contract);

    if (openPdfPreviewWindow(popupWindow, pdfUrl, fileName)) {
      popupWindow.focus();
      setQuoteStatus("Contract PDF opened in a new tab. Use Download PDF to save it with the correct filename.");
    } else {
      triggerPdfDownload(pdfUrl, fileName);
      setQuoteStatus("Contract PDF downloaded.");
    }

    schedulePdfUrlCleanup(pdfUrl, popupWindow);
  } catch (error) {
    console.error(error);
    if (popupWindow && !popupWindow.closed) {
      popupWindow.close();
    }
    const message = "Could not generate the contract PDF.";
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
  saveState();
}

function buildContractPreviewData() {
  const lineItems = state.measurementRows
    .map((row) => {
      const cost = getMeasurementCost(row);
      if (cost === null) {
        return null;
      }

      return {
        room: row.room?.trim() || "",
        label: row.label?.trim() || "",
        type: row.type?.trim() || "-",
        materialCode: row.materialCode?.trim() || "-",
        width: formatMeasurementDimension(row.width),
        height: formatMeasurementDimension(row.height),
        srp: cost,
      };
    })
    .filter(Boolean);

  const { subtotal, discountAmount, finalTotal, half } = getSummaryTotals();

  return {
    clientName: state.quoteMeta.clientName.trim() || "-",
    clientAddress: state.quoteMeta.projectName.trim() || "-",
    quoteDate: formatContractDate(state.quoteMeta.quoteDate),
    projectArchitect: state.quoteMeta.projectArchitect.trim() || "-",
    contactNumber: state.quoteMeta.contactNumber.trim() || "-",
    emailAddress: state.quoteMeta.emailAddress.trim() || "-",
    notes: state.quoteMeta.notes.trim(),
    lineItems,
    subtotal,
    discountAmount,
    finalTotal,
    downpayment: half,
    remainingBalance: half,
  };
}

function loadContractPdfAssets() {
  if (contractPdfAssetsPromise) {
    return contractPdfAssetsPromise;
  }

  const logoUrl = new URL("./assets/luxeshade-logo.png", window.location.href).href;
  contractPdfAssetsPromise = fileUrlToDataUrl(logoUrl).then((logoDataUrl) => ({
    logoDataUrl,
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

function buildContractPdfFileName(contract) {
  const clientName = contract.clientName.replace(/\s+/g, " ").trim();
  const rawName = clientName
    ? `LuxeShade Contract - ${clientName}`
    : "LuxeShade Contract";

  const safeName = rawName
    .replace(/[<>:"/\\|?*\u0000-\u001f]/g, "")
    .replace(/\s+/g, " ")
    .trim();

  return `${safeName || "LuxeShade Contract"}.pdf`;
}

function buildContractPdfDefinition(contract, assets) {
  const pageMargin = 43.2;
  const pageWidth = 612;
  const pageHeight = 792;
  const orderTableWidths = [92, 76, 112, 60, 60, 76];
  const totalsTableWidths = [340, 146];
  const borderColor = "#e6d8ca";
  const accentColor = "#9e7149";
  const textColor = "#2f2a26";
  const lightFill = "#f5ece2";
  const accentFill = "#f3e5d7";
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
    background(currentPage) {
      return {
        image: assets.logoDataUrl,
        width: 220,
        opacity: 0.16,
        absolutePosition: {
          x: (pageWidth - 220) / 2,
          y: (pageHeight - 220) / 2,
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
            image: assets.logoDataUrl,
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
                  body: [
                    [
                      { text: "Sub Total", bold: true },
                      { text: formatCurrency(contract.subtotal), alignment: "right" },
                    ],
                    [
                      { text: "Discount", bold: true },
                      { text: formatCurrency(contract.discountAmount), alignment: "right" },
                    ],
                    [
                      { text: "Total", bold: true, fontSize: 14, fillColor: accentFill },
                      { text: formatCurrency(contract.finalTotal), bold: true, fontSize: 14, alignment: "right", fillColor: accentFill },
                    ],
                  ],
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
        pageBreak: "after",
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
        pageBreak: "after",
      },
      {
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
            text: "",
            margin: [0, 250, 0, 0],
          },
          {
            text: [{ text: "Dated: ", bold: true }, contract.quoteDate],
            margin: [0, 0, 0, 24],
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
    fontSize: 9.4,
    color: "#000000",
    characterSpacing: 0.2,
  });

  const body = [[
    buildHeaderCell("Area"),
    buildHeaderCell("Type"),
    buildHeaderCell("Material Code"),
    buildHeaderCell("Width"),
    buildHeaderCell("Height"),
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
      { text: item.label || " ", alignment: "center", fontSize: 9 },
      { text: item.type, alignment: "center", fontSize: 9 },
      { text: item.materialCode, alignment: "center", fontSize: 9 },
      { text: item.width, alignment: "center", fontSize: 9 },
      { text: item.height, alignment: "center", fontSize: 9 },
      { text: formatCurrency(item.srp), alignment: "center", fontSize: 9 },
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

function buildTextInput({ value, placeholder, disabled = false, onInput }) {
  const input = document.createElement("input");
  input.type = "text";
  input.value = value;
  input.placeholder = placeholder;
  input.disabled = disabled;
  input.addEventListener("input", (event) => onInput(event.target.value));
  return input;
}

function buildNumberInput({
  value,
  placeholder,
  disabled = false,
  onInput,
}) {
  const input = document.createElement("input");
  input.type = "number";
  input.min = "0";
  input.step = "0.01";
  input.inputMode = "decimal";
  input.value = value;
  input.placeholder = placeholder;
  input.disabled = disabled;
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

function formatPdfPesoAmount(value) {
  const numericValue = Number(value);
  const safeValue = Number.isFinite(numericValue) ? numericValue : 0;
  return `\u20b1${safeValue.toLocaleString("en-PH", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  })}`;
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
