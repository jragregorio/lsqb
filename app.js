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

const refs = {
  adminMenuBtn: document.querySelector("#admin-menu-btn"),
  adminDrawer: document.querySelector("#admin-tools-drawer"),
  adminDrawerBackdrop: document.querySelector("#admin-drawer-backdrop"),
  adminDrawerCloseBtn: document.querySelector("#admin-drawer-close-btn"),
  activeQuoteBar: document.querySelector("#active-quote-bar"),
  activeQuoteTitle: document.querySelector("#active-quote-title"),
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
  quoteNotes: document.querySelector("#quote-notes"),
  newQuoteBtn: document.querySelector("#new-quote-btn"),
  saveQuoteBtn: document.querySelector("#save-quote-btn"),
  refreshQuotesBtn: document.querySelector("#refresh-quotes-btn"),
  currentQuoteLabel: document.querySelector("#current-quote-label"),
  quoteStatus: document.querySelector("#quote-status"),
  savedQuotesCount: document.querySelector("#saved-quotes-count"),
  savedQuotesList: document.querySelector("#saved-quotes-list"),
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

const runtime = {
  session: null,
  authBusy: false,
  sourceBusy: false,
  quoteBusy: false,
  quoteListBusy: false,
  quoteList: [],
  floatingObserverInitialized: false,
  expandedQuoteId: "",
  adminDrawerOpen: false,
  loadedQuoteFingerprint: "",
};

bootstrap();

async function bootstrap() {
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
    saveState();
    renderSummary();
  });
  refs.discountValue.addEventListener("input", (event) => {
    state.discountValue = normalizeInputNumber(event.target.value);
    saveState();
    renderSummary();
  });
  refs.signInBtn.addEventListener("click", handleSignIn);
  refs.signOutBtn.addEventListener("click", handleSignOut);
  refs.syncSupabaseBtn.addEventListener("click", () =>
    loadMaterialsFromSupabase({ showAlertOnFailure: true }),
  );
  refs.newQuoteBtn.addEventListener("click", handleNewQuote);
  refs.unloadQuoteBtn?.addEventListener("click", handleUnloadQuote);
  refs.saveQuoteBtn.addEventListener("click", handleSaveQuote);
  refs.refreshQuotesBtn.addEventListener("click", () =>
    refreshSavedQuotes({ showAlertOnFailure: true }),
  );
  refs.quoteClientName.addEventListener("input", (event) => {
    state.quoteMeta.clientName = event.target.value;
    saveState();
    renderQuoteStatus();
  });
  refs.quoteProjectName.addEventListener("input", (event) => {
    state.quoteMeta.projectName = event.target.value;
    saveState();
    renderQuoteStatus();
  });
  refs.quoteNotes.addEventListener("input", (event) => {
    state.quoteMeta.notes = event.target.value;
    saveState();
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
  if (event.key === "Escape" && runtime.adminDrawerOpen) {
    setAdminDrawerOpen(false);
  }
}

function handleWindowScroll() {
  renderActiveQuoteBar();
}

function setAdminDrawerOpen(isOpen) {
  runtime.adminDrawerOpen = Boolean(isOpen);
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
  document.body.classList.toggle("admin-drawer-open", runtime.adminDrawerOpen);
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

  resetQuoteDraft();
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
    runtime.authBusy = false;
    runtime.sourceBusy = false;
    runtime.quoteBusy = false;
    runtime.quoteListBusy = false;
    runtime.quoteList = [];
    runtime.expandedQuoteId = "";
    setAuthStatus("Not signed in yet.");
    setQuoteStatus("Sign in to save and reopen quotes from Supabase.");
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
  saveState();
  renderMaterials();
  renderMeasurements();
}

function handleAddMeasurement() {
  if (getConfiguredMaterials().length === 0) {
    window.alert("Add at least one configured material before adding measurements.");
    return;
  }

  state.measurementRows.push(createMeasurementRow());
  saveState();
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

  resetQuoteDraft();
  saveState();
  render();
  setQuoteStatus("New quote draft started.");
}

async function handleSaveQuote() {
  if (!runtime.session) {
    setQuoteStatus("Sign in first before saving quotes.", true);
    return;
  }

  const validation = validateQuoteForSave();
  if (!validation.ok) {
    setQuoteStatus(validation.message, true);
    window.alert(validation.message);
    return;
  }

  runtime.quoteBusy = true;
  renderQuoteWorkspace();
  setQuoteStatus(state.quoteMeta.id ? "Updating quote..." : "Saving quote...");
  const summaryTotals = getSummaryTotals();

  const quotePayload = {
    owner_user_id: runtime.session.user.id,
    client_name: state.quoteMeta.clientName.trim(),
    project_name: sanitizeOptionalText(state.quoteMeta.projectName),
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
        .select("id, client_name, project_name, quote_reference, notes, discount, discount_type, discount_value, subtotal_amount, applied_discount_amount, final_total_amount, created_at, updated_at")
        .single()
    : supabase
        .from("quotes")
        .insert(quotePayload)
        .select("id, client_name, project_name, quote_reference, notes, discount, discount_type, discount_value, subtotal_amount, applied_discount_amount, final_total_amount, created_at, updated_at")
        .single();

  const { data: savedQuote, error: quoteError } = await quoteQuery;

  if (quoteError) {
    runtime.quoteBusy = false;
    renderQuoteWorkspace();
    console.error(quoteError);
    setQuoteStatus(
      quoteError.message || "Could not save the quote header. Run the quote security migration first.",
      true,
    );
    return;
  }

  const { error: measurementDeleteError } = await supabase
    .from("quote_measurements")
    .delete()
    .eq("quote_id", savedQuote.id);

  if (measurementDeleteError) {
    runtime.quoteBusy = false;
    renderQuoteWorkspace();
    console.error(measurementDeleteError);
    setQuoteStatus(measurementDeleteError.message || "Could not replace saved measurements.", true);
    return;
  }

  const { error: materialDeleteError } = await supabase
    .from("quote_materials")
    .delete()
    .eq("quote_id", savedQuote.id);

  if (materialDeleteError) {
    runtime.quoteBusy = false;
    renderQuoteWorkspace();
    console.error(materialDeleteError);
    setQuoteStatus(materialDeleteError.message || "Could not replace saved materials.", true);
    return;
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
      renderQuoteWorkspace();
      console.error(materialInsertError);
      setQuoteStatus(materialInsertError.message || "Could not save quote materials.", true);
      return;
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
      renderQuoteWorkspace();
      console.error(measurementInsertError);
      setQuoteStatus(measurementInsertError.message || "Could not save quote measurements.", true);
      return;
    }
  }

  state.quoteMeta = {
    id: savedQuote.id,
    clientName: savedQuote.client_name || "",
    projectName: savedQuote.project_name || "",
    quoteReference: savedQuote.quote_reference || "",
    notes: savedQuote.notes || "",
    createdAt: savedQuote.created_at || "",
    updatedAt: savedQuote.updated_at || "",
  };
  state.discountType = savedQuote.discount_type === "percent" ? "percent" : "amount";
  state.discountValue =
    savedQuote.discount_value !== null && savedQuote.discount_value !== undefined
      ? String(savedQuote.discount_value)
      : savedQuote.discount
        ? String(savedQuote.discount)
        : "";
  runtime.loadedQuoteFingerprint = buildCurrentQuoteFingerprint();
  saveState();

  runtime.quoteBusy = false;
  await refreshSavedQuotes({ showAlertOnFailure: false, silent: true });
  render();
  setQuoteStatus("Quote saved to Supabase.");
}

async function loadQuoteById(quoteId) {
  if (!runtime.session) {
    setQuoteStatus("Sign in first before loading saved quotes.", true);
    return;
  }

  runtime.quoteBusy = true;
  renderQuoteWorkspace();
  setQuoteStatus("Loading quote...");

  const [quoteResult, materialsResult, measurementsResult] = await Promise.all([
    supabase
      .from("quotes")
      .select("id, client_name, project_name, quote_reference, notes, discount, discount_type, discount_value, subtotal_amount, applied_discount_amount, final_total_amount, created_at, updated_at")
      .eq("id", quoteId)
      .single(),
    supabase
      .from("quote_materials")
      .select("id, material_catalog_id, division, category, retail_price, asking_price, sort_order")
      .eq("quote_id", quoteId)
      .order("sort_order", { ascending: true }),
    supabase
      .from("quote_measurements")
      .select("id, quote_material_id, room_section, label, width_mm, height_mm, material_label, asking_price, sort_order")
      .eq("quote_id", quoteId)
      .order("sort_order", { ascending: true }),
  ]);

  runtime.quoteBusy = false;

  if (quoteResult.error || materialsResult.error || measurementsResult.error) {
    console.error(quoteResult.error || materialsResult.error || measurementsResult.error);
    setQuoteStatus("Could not load the selected quote.", true);
    renderQuoteWorkspace();
    return;
  }

  const localIdMap = new Map();

  state.quoteMeta = {
    id: quoteResult.data.id,
    clientName: quoteResult.data.client_name || "",
    projectName: quoteResult.data.project_name || "",
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
      label: row.label || "",
      width: row.width_mm === null ? "" : String(row.width_mm),
      height: row.height_mm === null ? "" : String(row.height_mm),
      materialId: localIdMap.get(row.quote_material_id) || "",
    })) || [];

  state.discountType =
    quoteResult.data.discount_type === "percent" ? "percent" : "amount";
  state.discountValue =
    quoteResult.data.discount_value !== null &&
    quoteResult.data.discount_value !== undefined
      ? String(quoteResult.data.discount_value)
      : quoteResult.data.discount
        ? String(quoteResult.data.discount)
        : "";
  ensureStarterRows();
  runtime.expandedQuoteId = quoteId;
  runtime.loadedQuoteFingerprint = buildCurrentQuoteFingerprint();
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

  runtime.quoteBusy = true;
  renderQuoteWorkspace();
  setQuoteStatus("Deleting quote...");

  const { error } = await supabase.from("quotes").delete().eq("id", quoteId);

  runtime.quoteBusy = false;

  if (error) {
    console.error(error);
    setQuoteStatus(error.message || "Could not delete the quote.", true);
    renderQuoteWorkspace();
    return;
  }

  if (state.quoteMeta.id === quoteId) {
    resetQuoteDraft();
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
  renderActiveQuoteBar();
  renderAdminDrawer();
  renderAuth();
  renderSourceStatus();
  renderQuoteWorkspace();
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
  refs.unloadQuoteBtn.disabled = runtime.quoteBusy;
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
  refs.quoteNotes.value = state.quoteMeta.notes;

  const signedIn = Boolean(runtime.session);
  refs.quoteClientName.disabled = runtime.quoteBusy;
  refs.quoteProjectName.disabled = runtime.quoteBusy;
  refs.quoteNotes.disabled = runtime.quoteBusy;
  refs.newQuoteBtn.disabled = runtime.quoteBusy;
  refs.saveQuoteBtn.disabled = !signedIn || runtime.quoteBusy;
  refs.refreshQuotesBtn.disabled =
    !signedIn || runtime.quoteBusy || runtime.quoteListBusy;

  renderQuoteStatus();
  renderSavedQuotesList();
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

      saveState();
      sanitizeMeasurementMaterialSelections();
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
      saveState();
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
      saveState();
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
      createEmptyStateRow(7, "Add a measurement row to begin."),
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
          saveState();
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
          saveState();
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
          saveState();
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
          saveState();
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
      saveState();
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

      saveState();
      renderMeasurements();
      renderSummary();
    });
    actionCell.append(removeButton);

    tr.append(
      roomCell,
      labelCell,
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
  refs.discountValue.value = state.discountValue;
  refs.discountType.disabled = runtime.quoteBusy;
  refs.discountValue.disabled = runtime.quoteBusy;
  refs.discountValue.placeholder =
    state.discountType === "percent" ? "0-100" : "0.00";
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

function hasMeaningfulDraftChanges() {
  return Boolean(
    state.quoteMeta.clientName ||
      state.quoteMeta.projectName ||
      state.quoteMeta.notes ||
      state.discountValue ||
      state.selectedMaterials.some((row) => row.category || row.askingPrice !== "") ||
      state.measurementRows.some(
        (row) => row.room || row.label || row.width !== "" || row.height !== "" || row.materialId,
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
