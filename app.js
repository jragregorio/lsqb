import { createClient } from "https://cdn.jsdelivr.net/npm/@supabase/supabase-js@2/+esm";

const STORAGE_KEY = "luxeshade-quote-builder-v1";
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

const refs = {
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
  materialBody: document.querySelector("#material-setup-body"),
  measurementBody: document.querySelector("#measurement-body"),
  addMaterialBtn: document.querySelector("#add-material-btn"),
  addMeasurementBtn: document.querySelector("#add-measurement-btn"),
  discountInput: document.querySelector("#discount-input"),
};

const state = loadState();
const runtime = {
  session: null,
  authBusy: false,
  sourceBusy: false,
};

bootstrap();

async function bootstrap() {
  refs.csvInput.addEventListener("change", handleCsvUpload);
  refs.loadSampleBtn.addEventListener("click", handleLoadBundledSample);
  refs.addMaterialBtn.addEventListener("click", handleAddMaterial);
  refs.addMeasurementBtn.addEventListener("click", handleAddMeasurement);
  refs.discountInput.addEventListener("input", (event) => {
    state.discount = normalizeInputNumber(event.target.value);
    saveState();
    renderSummary();
  });
  refs.signInBtn.addEventListener("click", handleSignIn);
  refs.signOutBtn.addEventListener("click", handleSignOut);
  refs.syncSupabaseBtn.addEventListener("click", () =>
    loadMaterialsFromSupabase({ showAlertOnFailure: true }),
  );

  ensureStarterRows();
  render();
  await initializeAuth();
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
    renderSourceStatus();
    return;
  }

  applySession(session);

  supabase.auth.onAuthStateChange(async (_event, sessionState) => {
    applySession(sessionState);
    if (sessionState && !isSupabaseSourceLoaded()) {
      await loadMaterialsFromSupabase({ showAlertOnFailure: false });
    }
  });

  if (session) {
    await loadMaterialsFromSupabase({ showAlertOnFailure: false });
  }
}

function applySession(session) {
  runtime.session = session;
  if (!session) {
    setAuthStatus("Not signed in yet.");
  }
  renderAuth();
  renderSourceStatus();
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
      selectedMaterials: Array.isArray(parsed.selectedMaterials)
        ? parsed.selectedMaterials
        : [],
      measurementRows: Array.isArray(parsed.measurementRows)
        ? parsed.measurementRows
        : [],
      discount: parsed.discount ?? "",
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
    selectedMaterials: [],
    measurementRows: [],
    discount: "",
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
    .select("division, category, retail_price")
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

    return createMaterialSetupRow();
  });

  sanitizeMeasurementMaterialSelections();
}

function sanitizeMeasurementMaterialSelections() {
  const validMaterialIds = new Set(
    state.selectedMaterials
      .filter((material) => material.sourceKey)
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

function render() {
  renderAuth();
  renderSourceStatus();
  renderMaterials();
  renderMeasurements();
  renderSummary();
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
  refs.addMaterialBtn.disabled = count === 0;
  refs.addMeasurementBtn.disabled = getConfiguredMaterials().length === 0;
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

    const pocketCell = document.createElement("td");
    pocketCell.className = "money-cell";
    const pocket = getPocketValue(row);
    pocketCell.textContent = pocket === null ? "PHP 0.00" : formatCurrency(pocket);

    askingInput.value = row.askingPrice;
    askingInput.disabled = !row.sourceKey;
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
  refs.addMeasurementBtn.disabled = configuredMaterials.length === 0;

  state.measurementRows.forEach((row) => {
    const tr = document.createElement("tr");

    const roomCell = document.createElement("td");
    roomCell.append(
      buildTextInput({
        value: row.room,
        placeholder: "Living Room",
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
    select.disabled = configuredMaterials.length === 0;
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
  refs.discountInput.value = state.discount;

  const subtotal = getSubtotal();
  const discount = parseCurrencyLikeNumber(state.discount) || 0;
  const finalTotal = Math.max(0, subtotal - discount);
  const half = finalTotal / 2;

  document.querySelector("#subtotal-value").textContent = formatCurrency(subtotal);
  document.querySelector("#final-total-value").textContent =
    formatCurrency(finalTotal);
  document.querySelector("#downpayment-value").textContent = formatCurrency(half);
  document.querySelector("#remaining-value").textContent = formatCurrency(half);
}

function getConfiguredMaterials() {
  return state.selectedMaterials.filter(
    (row) => row.sourceKey && parseCurrencyLikeNumber(row.askingPrice) !== null,
  );
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

function setAuthStatus(message, isError = false) {
  refs.authStatus.textContent = message;
  refs.authStatus.classList.toggle("is-error", isError);
}

function isSupabaseSourceLoaded() {
  return state.sourceLabel === "Supabase material_catalog";
}

function buildTextInput({ value, placeholder, onInput }) {
  const input = document.createElement("input");
  input.type = "text";
  input.value = value;
  input.placeholder = placeholder;
  input.addEventListener("input", (event) => onInput(event.target.value));
  return input;
}

function buildNumberInput({ value, placeholder, onInput }) {
  const input = document.createElement("input");
  input.type = "number";
  input.min = "0";
  input.step = "0.01";
  input.inputMode = "decimal";
  input.value = value;
  input.placeholder = placeholder;
  input.addEventListener("input", (event) =>
    onInput(normalizeInputNumber(event.target.value)),
  );
  return input;
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
