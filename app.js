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
  activeQuoteSaveBtn: document.querySelector("#active-save-quote-btn"),
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
  previewContractBtn: document.querySelector("#preview-contract-btn"),
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
  floatingTotalCard: document.querySelector(".floating-total-card"),
  floatingFinalTotal: document.querySelector("#floating-final-total"),
  floatingTotalMeta: document.querySelector("#floating-total-meta"),
};

const state = loadState();
const ACTIVE_QUOTE_BAR_REVEAL_SCROLL_Y = 100;
const AUTOSAVE_DELAY_MS = 1800;

const runtime = {
  session: null,
  authBusy: false,
  sourceBusy: false,
  quoteBusy: false,
  quoteListBusy: false,
  quoteList: [],
  floatingObserverInitialized: false,
  expandedQuoteId: "",
  savedQuotesDrawerOpen: false,
  adminDrawerOpen: false,
  loadedQuoteFingerprint: "",
  quoteWorkspaceActive: false,
  autosaveTimer: 0,
};

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
  refs.unloadQuoteBtn?.addEventListener("click", handleUnloadQuote);
  refs.saveQuoteBtn.addEventListener("click", () => {
    void handleSaveQuote();
  });
  refs.previewContractBtn?.addEventListener("click", handlePreviewContract);
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
  initializeFloatingTotalVisibility();
  await initializeAuth();
}

function initializeFloatingTotalVisibility() {
  if (
    runtime.floatingObserverInitialized ||
    !refs.summaryPanel ||
    !refs.floatingTotalCard ||
    typeof IntersectionObserver === "undefined"
  ) {
    return;
  }

  const observer = new IntersectionObserver(
    ([entry]) => {
      updateFloatingTotalVisibility(entry.isIntersecting);
    },
    {
      threshold: 0.15,
    },
  );

  observer.observe(refs.summaryPanel);
  runtime.floatingObserverInitialized = true;
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

function updateFloatingTotalVisibility(summaryInView = false) {
  if (!refs.floatingTotalCard) {
    return;
  }

  const shouldHide = summaryInView || !hasMeaningfulDraftChanges();
  refs.floatingTotalCard.classList.toggle("is-hidden", shouldHide);
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
  refs.activeQuoteBar.classList.toggle("hidden", !shouldShow);
  refs.activeQuoteTitle.textContent = hasLoadedQuote
    ? buildActiveQuoteTitle()
    : "-";
  if (refs.activeQuoteSaveBtn) {
    refs.activeQuoteSaveBtn.disabled = !runtime.session || runtime.quoteBusy;
    refs.activeQuoteSaveBtn.textContent = runtime.quoteBusy ? "Saving..." : "Save Quote";
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
  refs.previewContractBtn.disabled = runtime.quoteBusy || !isQuoteWorkspaceActive();
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
  refs.floatingFinalTotal.textContent = formatCurrency(finalTotal);
  refs.floatingTotalMeta.textContent = `Discount ${formatCurrency(discountAmount)} • 50% DP ${formatCurrency(half)}`;
  updateFloatingTotalVisibility();
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

    if (!row.label || width === null || height === null || !selectedMaterial || askingPrice === null) {
      return {
        ok: false,
        message:
          "Every measurement row must have a label, width, height, and configured material before saving.",
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

function handlePreviewContract() {
  const validation = validateQuoteForSave();
  if (!validation.ok) {
    setQuoteStatus(validation.message, true);
    window.alert(validation.message);
    return;
  }

  const previewWindow = window.open("", "_blank");
  if (!previewWindow) {
    window.alert("Allow pop-ups for this site so the contract preview can open.");
    return;
  }

  previewWindow.document.open();
  previewWindow.document.write(buildContractDocumentHtml());
  previewWindow.document.close();
  previewWindow.focus();
  setQuoteStatus("Client contract preview opened in a new tab.");
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

function buildContractPreviewData() {
  const lineItems = state.measurementRows
    .map((row) => {
      const cost = getMeasurementCost(row);
      if (cost === null) {
        return null;
      }

      const areaLabel = [row.room?.trim(), row.label?.trim()]
        .filter(Boolean)
        .join(" - ");

      return {
        area: areaLabel || "Measurement",
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

function buildLuxeShadeLogoMarkup(logoSrc, { watermark = false } = {}) {
  const logoClass = watermark ? "logo-image logo-image-watermark" : "logo-image";
  const altText = watermark ? "" : "LuxeShade logo";
  return `<img class="${logoClass}" src="${escapeHtml(logoSrc)}" alt="${altText}" data-chroma-logo="true" />`;
}

function buildContractDocumentHtml() {
  const contract = buildContractPreviewData();
  const logoSrc = new URL("./assets/luxeshade-logo.png", window.location.href).href;
  const pageWatermark = buildLuxeShadeLogoMarkup(logoSrc, { watermark: true });
  const heroLogo = buildLuxeShadeLogoMarkup(logoSrc);
  const notesPanelHtml = contract.notes
    ? `
          <div class="info-card info-card-wide">
            <span>Notes</span>
            <div class="info-card-copy">${escapeHtml(contract.notes)}</div>
          </div>`
    : "";
  const lineItemsHtml = contract.lineItems
    .map(
      (item) => `
        <tr>
          <td>${escapeHtml(item.area)}</td>
          <td>${escapeHtml(item.type)}</td>
          <td>${escapeHtml(item.materialCode)}</td>
          <td class="align-right">${escapeHtml(item.width)}</td>
          <td class="align-right">${escapeHtml(item.height)}</td>
          <td class="align-right">${escapeHtml(formatCurrency(item.srp))}</td>
        </tr>`,
    )
    .join("");

  return `<!doctype html>
<html lang="en">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>LuxeShade Contract Preview</title>
    <style>
      :root {
        color-scheme: light;
        font-family: "Tenor Sans", "Segoe UI", Roboto, sans-serif;
        line-height: 1.45;
        color: #2f2a26;
        background: #efe8df;
      }

      * {
        box-sizing: border-box;
      }

      body {
        margin: 0;
        padding: 1.25rem;
        background: #efe8df;
      }

      @page {
        size: Letter;
        margin: 0.6in;
      }

      .contract-toolbar {
        position: sticky;
        top: 0;
        z-index: 10;
        display: flex;
        gap: 0.75rem;
        align-items: center;
        justify-content: space-between;
        margin: 0 auto 1rem;
        width: min(960px, 100%);
        padding: 0.9rem 1rem;
        border: 1px solid #ddcec1;
        border-radius: 1rem;
        background: rgba(255, 250, 246, 0.96);
        box-shadow: 0 10px 24px rgba(84, 58, 37, 0.1);
      }

      .contract-toolbar-copy {
        margin: 0;
        color: #76685d;
      }

      .toolbar-actions {
        display: flex;
        gap: 0.75rem;
        flex-wrap: wrap;
      }

      .toolbar-actions button {
        border: 0;
        border-radius: 999px;
        padding: 0.72rem 1.2rem;
        font: inherit;
        cursor: pointer;
      }

      .print-button {
        background: #b8895d;
        color: #ffffff;
      }

      .close-button {
        background: #f0e4d8;
        color: #4a392d;
      }

      .contract-page {
        position: relative;
        width: min(8.5in, 100%);
        min-height: 11in;
        margin: 0 auto 1rem;
        padding: 2.2rem;
        border: 1px solid #ddcec1;
        border-radius: 1.2rem;
        background: #ffffff;
        box-shadow: 0 18px 38px rgba(84, 58, 37, 0.08);
        overflow: hidden;
        font-size: 10pt;
        line-height: 1.35;
      }

      .contract-page.page-break {
        page-break-after: always;
      }

      .contract-page-content {
        position: relative;
        z-index: 1;
        min-height: calc(11in - 4.4rem);
      }

      .page-watermark {
        position: absolute;
        inset: 0;
        display: flex;
        align-items: center;
        justify-content: center;
        pointer-events: none;
        z-index: 0;
      }

      .logo-image {
        display: block;
        width: 10rem;
        height: auto;
        margin: 0 auto;
        opacity: 1;
      }

      .logo-image-watermark {
        width: 20rem;
        opacity: 0.22;
      }

      h1,
      h2,
      h3,
      p {
        margin-top: 0;
      }

      .contract-title {
        margin-bottom: 1.8rem;
        text-align: center;
      }

      .contract-title h1 {
        margin-bottom: 0.35rem;
        font-size: 12pt;
        font-weight: 700;
      }

      .contract-title p {
        color: #76685d;
        font-size: 10pt;
      }

      .hero-logo {
        margin-bottom: 1.1rem;
      }

      .info-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 0.8rem 1rem;
        margin-bottom: 1.4rem;
      }

      .info-card {
        padding: 0.85rem 0.95rem;
        border: 1px solid #e6d8ca;
        border-radius: 0.85rem;
        background: #fffaf6;
      }

      .info-card-wide {
        grid-column: 1 / -1;
      }

      .info-card span {
        display: block;
        margin-bottom: 0.35rem;
        color: #76685d;
        font-size: 10pt;
      }

      .info-card strong {
        display: block;
        font-size: 10pt;
        font-weight: 700;
      }

      .info-card-copy {
        font-size: 10pt;
        white-space: pre-wrap;
      }

      .section-heading {
        margin: 1.5rem 0 0.8rem;
        font-size: 10pt;
        font-weight: 700;
        letter-spacing: 0.08em;
        text-transform: uppercase;
        color: #9e7149;
      }

      table {
        width: 100%;
        border-collapse: collapse;
      }

      th,
      td {
        padding: 0.72rem 0.75rem;
        border: 1px solid #e6d8ca;
        vertical-align: top;
        font-size: 10pt;
      }

      th {
        background: #f5ece2;
        text-align: left;
        font-weight: 700;
      }

      .align-right {
        text-align: right;
      }

      .order-details-table th,
      .order-details-table td,
      .order-details-table .align-right {
        text-align: center;
        vertical-align: middle;
      }

      .totals-table {
        margin-top: 1rem;
      }

      .totals-table td:first-child {
        width: 68%;
        font-weight: 700;
      }

      .totals-table tr.is-accent td {
        background: #f3e5d7;
        font-weight: 700;
      }

      .terms-section {
        margin-bottom: 1.3rem;
        font-size: 10pt;
      }

      .terms-section h2 {
        margin-bottom: 0.45rem;
        font-size: 11pt;
        font-weight: 700;
      }

      .terms-section h3 {
        margin-bottom: 0.45rem;
        font-size: 10pt;
        font-weight: 700;
      }

      .terms-section ol,
      .terms-section ul {
        margin: 0.45rem 0 0 1.2rem;
        padding: 0;
        font-size: 10pt;
      }

      .terms-section li {
        margin-bottom: 0.35rem;
        font-size: 10pt;
      }

      .signature-grid {
        display: grid;
        grid-template-columns: repeat(2, minmax(0, 1fr));
        gap: 2rem;
      }

      .signature-line {
        padding-top: 2.6rem;
        border-top: 1px solid #2f2a26;
        font-size: 10pt;
      }

      .section-four-page {
        display: flex;
        flex-direction: column;
      }

      .section-four-spacer {
        flex: 1;
        min-height: 3.8in;
      }

      .page-footer {
        margin-top: 2rem;
        color: #76685d;
        font-size: 9pt;
        text-align: right;
      }

      @media print {
        body {
          padding: 0;
          background: #ffffff;
        }

        .contract-toolbar {
          display: none;
        }

        .contract-page {
          width: 100%;
          min-height: auto;
          margin: 0;
          padding: 0.6in;
          border: 0;
          border-radius: 0;
          box-shadow: none;
        }

        .contract-page-content {
          min-height: auto;
        }

        .section-four-spacer {
          min-height: 4.6in;
        }
      }
    </style>
  </head>
  <body>
    <div class="contract-toolbar">
      <p class="contract-toolbar-copy">Preview uses the current quote values. Choose Print and save as PDF when ready.</p>
      <div class="toolbar-actions">
        <button class="print-button" onclick="window.print()">Print / Save PDF</button>
        <button class="close-button" onclick="window.close()">Close Preview</button>
      </div>
    </div>

    <section class="contract-page page-break">
      <div class="page-watermark">${pageWatermark}</div>
      <div class="contract-page-content">
        <div class="contract-title">
          <div class="hero-logo">${heroLogo}</div>
          <h1>Client Information</h1>
        </div>

        <div class="info-grid">
          <div class="info-card">
            <span>Date</span>
            <strong>${escapeHtml(contract.quoteDate)}</strong>
          </div>
          <div class="info-card">
            <span>Client's Address</span>
            <strong>${escapeHtml(contract.clientAddress)}</strong>
          </div>
          <div class="info-card">
            <span>Client's Name</span>
            <strong>${escapeHtml(contract.clientName)}</strong>
          </div>
          <div class="info-card">
            <span>Project Architect</span>
            <strong>${escapeHtml(contract.projectArchitect)}</strong>
          </div>
          <div class="info-card">
            <span>Contact No.</span>
            <strong>${escapeHtml(contract.contactNumber)}</strong>
          </div>
          <div class="info-card">
            <span>Email Address</span>
            <strong>${escapeHtml(contract.emailAddress)}</strong>
          </div>
          ${notesPanelHtml}
        </div>

        <h2 class="section-heading">Order Details</h2>
        <table class="order-details-table">
          <thead>
            <tr>
              <th>Area</th>
              <th>Type</th>
              <th>Material Code</th>
              <th class="align-right">Width</th>
              <th class="align-right">Height</th>
              <th class="align-right">SRP</th>
            </tr>
          </thead>
          <tbody>
            ${lineItemsHtml}
          </tbody>
        </table>

        <table class="totals-table">
          <tbody>
            <tr>
              <td>Sub Total</td>
              <td class="align-right">${escapeHtml(formatCurrency(contract.subtotal))}</td>
            </tr>
            <tr>
              <td>Discount</td>
              <td class="align-right">${escapeHtml(formatCurrency(contract.discountAmount))}</td>
            </tr>
            <tr class="is-accent">
              <td>Total</td>
              <td class="align-right">${escapeHtml(formatCurrency(contract.finalTotal))}</td>
            </tr>
          </tbody>
        </table>

        <p class="page-footer">Page 1 of 4</p>
      </div>
    </section>

    <section class="contract-page page-break">
      <div class="page-watermark">${pageWatermark}</div>
      <div class="contract-page-content">
        <div class="terms-section">
          <h2>Terms of Contract</h2>
          <h3>Section 1. Payment Schedule and Conditions</h3>
          <ol>
            <li>
              <strong>Down Payment (50%)</strong>
              <ul>
                <li>Amount: ${escapeHtml(formatCurrency(contract.downpayment))}</li>
                <li>Payable upon signing or approval of the contract.</li>
              </ul>
            </li>
            <li>
              <strong>Final Payment (50%)</strong>
              <ul>
                <li>Amount: ${escapeHtml(formatCurrency(contract.remainingBalance))}</li>
                <li>Payable on the same day of installation or within 24 hours upon completion of work.</li>
              </ul>
            </li>
            <li>
              <strong>Payment Options</strong>
              <ul>
                <li>Online payment via BDO is accepted.</li>
                <li>Account name: Ma. Elena Bernardo</li>
                <li>Account number: 0110 1002 1573</li>
                <li>Other payment methods such as checks or cash should be coordinated with the supplier as needed.</li>
                <li>For check payments, clearance must be confirmed before work or deliveries commence.</li>
              </ul>
            </li>
          </ol>
        </div>

        <div class="terms-section">
          <h3>Section 2. Work and Delivery Lead Time</h3>
          <ol>
            <li>
              <strong>Fabrication and Delivery</strong>
              <ul>
                <li>Typically 12 to 14 working days, if fabric is available, from the date of down payment or from the clearing date for check payments.</li>
                <li>This lead time covers the production and preparation of materials.</li>
              </ul>
            </li>
            <li>
              <strong>Approximate Completion Time On-Site</strong>
              <ul>
                <li>Generally 1 to 3 working days for small projects and 3 to 7 working days for big projects, depending on scope.</li>
                <li>Actual time may vary due to building administration requirements, power scheduling, access limitations, or unforeseen events.</li>
              </ul>
            </li>
            <li>
              <strong>Cooperation with Building or Project Management</strong>
              <ul>
                <li>The customer acknowledges that certain installation activities may require coordination with building management, such as permits, elevator scheduling, and access arrangements.</li>
                <li>Delays caused by building administration rules, weather, power interruptions, or similar external factors are beyond the supplier's control.</li>
              </ul>
            </li>
          </ol>
        </div>

        <p class="page-footer">Page 2 of 4</p>
      </div>
    </section>

    <section class="contract-page page-break">
      <div class="page-watermark">${pageWatermark}</div>
      <div class="contract-page-content">
        <div class="terms-section">
          <h3>Section 3. Installation Work</h3>
          <ol>
            <li>
              <strong>Coverage of Installation</strong>
              <ul>
                <li>The installation includes fabrication and assembly of the contracted items and proper placement in the customer's designated area.</li>
                <li>The supplier will only be responsible for connecting and installing components directly related to the agreed work.</li>
              </ul>
            </li>
            <li>
              <strong>Exclusions</strong>
              <ul>
                <li>Electrical wiring, water lines, sanitary connections, or structural modifications not stated in the contract are excluded unless otherwise agreed in writing.</li>
                <li>Plumbing or electrical work requiring specialized trades remains the responsibility of the customer or a separately engaged licensed professional.</li>
              </ul>
            </li>
            <li>
              <strong>Building Permits and Approvals</strong>
              <ul>
                <li>The supplier shall secure the necessary building permits and condo work permits required for installation.</li>
                <li>Any associated permit fees may be charged to the customer if not included in the initial quotation.</li>
              </ul>
            </li>
            <li>
              <strong>Liability</strong>
              <ul>
                <li>The supplier will not be held liable for damage or incidents caused by existing building conditions such as defective pipes, concealed wiring, or structural issues.</li>
                <li>Damage resulting from negligence or mishandling by supplier personnel will be addressed in accordance with applicable local laws or as otherwise agreed by both parties.</li>
                <li>Fragile or delicate items should be removed or secured by the customer before installation to avoid damage.</li>
              </ul>
            </li>
          </ol>
        </div>

        <p class="page-footer">Page 3 of 4</p>
      </div>
    </section>

    <section class="contract-page section-four-page">
      <div class="page-watermark">${pageWatermark}</div>
      <div class="contract-page-content section-four-page">
        <div class="terms-section">
          <h3>Section 4. Customer Cancellation of Order</h3>
          <ol>
            <li>
              <strong>No Refund Policy</strong>
              <ul>
                <li>Once a down payment has been made, it is strictly non-refundable under any circumstances after contract signing.</li>
              </ul>
            </li>
          </ol>
        </div>

        <div class="section-four-spacer"></div>

        <div>
          <p><strong>Dated:</strong> ${escapeHtml(contract.quoteDate)}</p>

          <div class="signature-grid">
            <div class="signature-line">Company Representative</div>
            <div class="signature-line">Client Signature</div>
          </div>

          <p class="page-footer">Page 4 of 4</p>
        </div>
      </div>
    </section>
    <script>
      (() => {
        const cleanupLogo = (img) => {
          const canvas = document.createElement("canvas");
          const context = canvas.getContext("2d");
          if (!context) {
            return;
          }

          canvas.width = img.naturalWidth;
          canvas.height = img.naturalHeight;
          context.drawImage(img, 0, 0);

          const imageData = context.getImageData(0, 0, canvas.width, canvas.height);
          const { data } = imageData;

          for (let index = 0; index < data.length; index += 4) {
            const red = data[index];
            const green = data[index + 1];
            const blue = data[index + 2];
            const alpha = data[index + 3];
            const brightness = (red + green + blue) / 3;

            if (alpha === 0) {
              continue;
            }

            if (brightness > 245) {
              data[index + 3] = 0;
              continue;
            }

            data[index + 3] = 255;
          }

          context.putImageData(imageData, 0, 0);
          img.src = canvas.toDataURL("image/png");
        };

        document.querySelectorAll("img[data-chroma-logo='true']").forEach((img) => {
          if (img.complete) {
            cleanupLogo(img);
          } else {
            img.addEventListener("load", () => cleanupLogo(img), { once: true });
          }
        });
      })();
    </script>
  </body>
</html>`;
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
